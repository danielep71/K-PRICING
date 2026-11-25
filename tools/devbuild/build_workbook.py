#!/usr/bin/env python3
"""Build a macro-enabled workbook from the source of one commit.

Usage: python tools/devbuild/build_workbook.py [--configuration NAME] [--override DIR]
       [OUTPUT]
Default: the development configuration, written to
test-results/build/Workbook-development.xlsm, which Git ignores.

The workbook has the sheets, code names and defined names of
tools/devbuild/workbook.json, which tests/test_devbuild.py keeps equal to
docs/methodology/WORKBOOK_CONTRACT.md. Sheets with the Test role, and their
names, are built only in the development configuration. It carries the build stamp on the
reference sheet.

Without src/workbook/Template.xlsx, every sheet is written blank, with only
the stamp, a title in A1 of the reference sheet and, to the left of each name
that has one, its label. With it, the build starts
from a copy of the template, which keeps its layout, styles and protection:
its sheets must be the configuration's sheets other than Test, in order, and
it must define their names on the same cells. The build sets the code names,
writes the stamp values into the stamp cells, keeping their style, makes the
main part macro-enabled and adds the VBA project and its relationship. A
development build appends each Test sheet blank, so Test sheets come last in
the configuration. Every other part of the template is copied unchanged.

Its VBA project is written from source by vba_project.py:
the standard modules (.bas) and class modules (.cls) of the configuration's
directories, plus a document module for ThisWorkbook and for each sheet. A
document module comes from src/workbook/<code name>.cls when that file
exists. The project holds no compiled code, so Excel reads it from source
and compiles each module when it is first used
(docs/methodology/VALIDATION_PLAN.md, HB-05): the build needs no Trust access
to the VBA project object model.

Forms are not built: a .frm in a source directory stops the build, and forms
are imported by hand. For a given commit and working tree the output is
byte-identical, because the stamp carries the commit's committer time, not
the clock. A tracked change, or an untracked file in a source directory,
stamps the build as not evidence-grade.

With --override DIR, each .bas or .cls file found in DIR or its subfolders
replaces the module of the configuration whose Attribute VB_Name it carries,
and other files there are ignored. A file whose name matches no module, two
files with the same name, a file without exactly one VB_Name, a .frm, or a
file that breaks the encoding, header or line ending rules of exported source
stops the build: an override never adds or drops a module. The stamp's commit
then records the folder and the SHA-256 of each replacement and says
NOT evidence-grade, so tools/check_excel_run.py never accepts the workbook as
evidence for the commit. Standard library only.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import posixpath
import re
import subprocess
import sys
import uuid
import zipfile
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree
from xml.sax.saxutils import escape, quoteattr

sys.path.insert(0, str(Path(__file__).resolve().parent))
from vba_project import (  # noqa: E402
    CLASS,
    DOCUMENT,
    STANDARD,
    Module,
    build_vba_project,
    document_module,
)

ROOT = Path(__file__).resolve().parents[2]
CONFIG = "tools/devbuild/workbook.json"
DOCUMENT_HOME = "src/workbook"
TEMPLATE = "src/workbook/Template.xlsx"
CLASS_HOME = "src/classes"
DEVELOPMENT = "development"
# Sheet roles of the workbook contract; Test sheets exist only in a development build.
ROLES = ("Reference", "Input", "Parameter", "Output", "Diagnostics", "Test")
TEST_ROLE = "Test"
FORM_HOME = "src/forms"
VBA_ENCODING = "cp1252"
NOT_EVIDENCE = " (working tree with changes: NOT evidence-grade)"
# The start of the stamp's note on a build with --override, which
# tools/check_excel_run.py looks for.
OVERRIDDEN = " (built with overrides from "
OVERRIDE_SUFFIXES = (".bas", ".cls")
VB_NAME = re.compile(r'^Attribute VB_Name = "([^"]*)"[ \t]*\r?$', re.M)
ZIP_TIME = (1980, 1, 1, 0, 0, 0)
CELL = re.compile(r"^([A-Z]{1,3})([1-9][0-9]{0,6})$")
NAME = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
CODE_NAME = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,30}$")
HEADER = re.compile(r"^(VERSION \d+\.\d+( CLASS)?|(?i:begin)(\s.*)?|(?i:end)|\s+\S.*)$")
TITLE = "Built from source by tools/devbuild/build_workbook.py. Never commit this workbook."

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships"
WORKSHEET_TYPE = f"{RELATIONSHIPS}/worksheet"
STYLES_TYPE = f"{RELATIONSHIPS}/styles"
VBA_TYPE = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
DOCUMENT_TYPE = f"{RELATIONSHIPS}/officeDocument"
XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
CONTENT_TYPES = "[Content_Types].xml"
WORKBOOK_PART = "xl/workbook.xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
VBA_PART = "xl/vbaProject.bin"
SHEET_MAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
MACRO_MAIN = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
VBA_CONTENT = "application/vnd.ms-office.vbaProject"
WORKSHEET_CONTENT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
MACRO_PART = re.compile(r"vbaProject|macroEnabled", re.I)
# A defined name's reference to one cell: a quoted or plain tab, then $column$row.
REFERENCE = re.compile(r"^(?:'((?:[^']|'')+)'|([^'!]+))!\$([A-Z]{1,3})\$([0-9]+)$")
ROW = re.compile(r"<row\b([^>]*?)(/?)>")
CELL_ELEMENT = re.compile(r"<c\b([^>]*?)(?:/>|>.*?</c>)", re.S)


class BuildError(Exception):
    """The configuration or the source cannot be built."""


@dataclass(frozen=True)
class Sheet:
    tab: str
    code_name: str
    role: str


@dataclass(frozen=True)
class DefinedName:
    name: str
    sheet: str
    cell: str
    label: str


@dataclass(frozen=True)
class Layout:
    workbook_code_name: str
    sheets: tuple[Sheet, ...]
    names: tuple[DefinedName, ...]
    stamp: dict[str, str]
    configurations: dict[str, tuple[str, ...]]

    def sheet(self, code_name: str) -> Sheet:
        return next(sheet for sheet in self.sheets if sheet.code_name == code_name)


@dataclass(frozen=True)
class Override:
    """A module that a file of the override folder replaces."""

    name: str
    path: Path
    sha256: str
    source: str = field(repr=False, default="")


@dataclass(frozen=True)
class Build:
    commit: str
    time: str
    clean: bool
    configuration: str
    components: tuple[Path, ...]
    template: bool = False
    overrides: tuple[Override, ...] = ()
    override_folder: str = ""


def load_layout(root: Path) -> Layout:
    """The build configuration, validated; BuildError names the first problem."""
    path = root / CONFIG
    try:
        data: dict[str, Any] = json.loads(path.read_text(encoding="utf-8"))
        layout = Layout(
            data["workbook_code_name"],
            tuple(Sheet(item["tab"], item["code_name"], item["role"]) for item in data["sheets"]),
            tuple(DefinedName(item["name"], item["sheet"], item["cell"], item.get("label", ""))
                  for item in data["names"]),
            dict(data["stamp"]),
            {key: tuple(value) for key, value in data["configurations"].items()})
    except (OSError, ValueError, KeyError, TypeError) as error:
        raise BuildError(f"{CONFIG}: {error!r}") from error
    problems = layout_problems(layout)
    if problems:
        raise BuildError(f"{CONFIG}: {problems[0]}")
    return layout


def layout_problems(layout: Layout) -> list[str]:
    problems = []
    codes = [layout.workbook_code_name, *(sheet.code_name for sheet in layout.sheets)]
    tabs = [sheet.tab.casefold() for sheet in layout.sheets]
    if not layout.sheets:
        problems.append("no sheet")
    if len({code.casefold() for code in codes}) != len(codes):
        problems.append("a code name is used twice")
    if len(set(tabs)) != len(tabs):
        problems.append("a tab name is used twice")
    problems += [f"code name {code!r} is not a VBA identifier of at most 31 characters"
                 for code in codes if not CODE_NAME.match(code)]
    problems += [f"tab name {sheet.tab!r} is empty, longer than 31 characters or uses []:*?/\\"
                 for sheet in layout.sheets
                 if not sheet.tab or len(sheet.tab) > 31 or re.search(r"[\[\]:*?/\\]", sheet.tab)]
    names = [item.name.casefold() for item in layout.names]
    if len(set(names)) != len(names):
        problems.append("a defined name is used twice")
    for item in layout.names:
        if not NAME.match(item.name) or CELL.match(item.name.upper()):
            problems.append(f"defined name {item.name!r} is not a valid name")
        if item.sheet not in codes[1:]:
            problems.append(f"defined name {item.name} refers to unknown sheet {item.sheet}")
        if not CELL.match(item.cell):
            problems.append(f"defined name {item.name} refers to {item.cell!r}, not one cell")
    problems += [f"sheet {sheet.code_name} has role {sheet.role!r}, not one of "
                 f"{', '.join(ROLES)}" for sheet in layout.sheets if sheet.role not in ROLES]
    tests = {sheet.code_name for sheet in layout.sheets if sheet.role == TEST_ROLE}
    for role in ("commit", "time", "configuration"):
        stamp = next((item for item in layout.names if item.name == layout.stamp.get(role)), None)
        if stamp is None:
            problems.append(f"stamp {role} is not a defined name")
        elif stamp.sheet in tests:
            problems.append(f"stamp {role} is on a {TEST_ROLE} sheet, which a production "
                            "build leaves out")
    if DEVELOPMENT not in layout.configurations:
        problems.append("no development configuration")
    return problems


def git(root: Path, *args: str) -> str:
    completed = subprocess.run(["git", "-C", str(root), *args], capture_output=True, text=True,
                               check=False)
    if completed.returncode:
        raise BuildError(f"git {' '.join(args)}: {completed.stderr.strip()}")
    return completed.stdout.strip()


def source_directories(layout: Layout, configuration: str) -> tuple[str, ...]:
    if configuration not in layout.configurations:
        raise BuildError(f"unknown configuration {configuration!r}; {CONFIG} defines "
                         f"{', '.join(sorted(layout.configurations))}")
    return layout.configurations[configuration]


def components(root: Path, layout: Layout, configuration: str) -> list[Path]:
    """The standard and class modules the configuration embeds, in file name order."""
    found: list[Path] = []
    for directory in (*source_directories(layout, configuration), FORM_HOME):
        forms = sorted(path.name for path in (root / directory).glob("*.frm")) \
            if (root / directory).is_dir() else []
        if forms:
            raise BuildError(f"{directory}: forms are not built ({', '.join(forms)}); import "
                             "them by hand as INSTALLATION.md describes")
    for directory in source_directories(layout, configuration):
        folder = root / directory
        if not folder.is_dir():
            continue
        found += sorted(folder.glob("*.bas"))
        if directory == CLASS_HOME:
            found += sorted(folder.glob("*.cls"))
    documents = {code.casefold(): code for code in
                 [layout.workbook_code_name, *(sheet.code_name for sheet in layout.sheets)]}
    exported = sorted((root / DOCUMENT_HOME).glob("*.cls")) \
        if (root / DOCUMENT_HOME).is_dir() else []
    strays = [path.name for path in exported if path.stem.casefold() not in documents]
    if strays:
        raise BuildError(f"{DOCUMENT_HOME}: {', '.join(strays)} matches no sheet of {CONFIG}")
    # The file is found by its exact name, so a name that differs only in case
    # would be missed on a case-sensitive file system and replaced by an empty
    # document module.
    miscased = [f"{path.name} (code name {documents[path.stem.casefold()]})"
                for path in exported if path.stem != documents[path.stem.casefold()]]
    if miscased:
        raise BuildError(f"{DOCUMENT_HOME}: {', '.join(miscased)} differs from its code name "
                         "in case; rename it to the code name")
    return found


def module_text(path: Path) -> str:
    """A component's source as the VBE stores it: CRLF, without the VERSION and
    BEGIN ... END lines an exported class writes before its attributes."""
    lines = path.read_bytes().decode(VBA_ENCODING).replace("\r\n", "\n").split("\n")
    start = next((index for index, line in enumerate(lines) if line.startswith("Attribute ")), 0)
    if any(not HEADER.match(line) for line in lines[:start]):
        start = 0
    text = "\r\n".join(lines[start:])
    return text if text.endswith("\r\n") else text + "\r\n"


def vba_modules(root: Path, layout: Layout, paths: Sequence[Path]) -> list[Module]:
    modules = []
    for code in [layout.workbook_code_name, *(sheet.code_name for sheet in layout.sheets)]:
        path = root / DOCUMENT_HOME / f"{code}.cls"
        modules.append(Module(code, module_text(path), kind=DOCUMENT) if path.is_file()
                       else document_module(code, code == layout.workbook_code_name))
    for path in paths:
        modules.append(Module(path.stem, module_text(path),
                              kind=CLASS if path.suffix.lower() == ".cls" else STANDARD))
    return modules


def override_problem(path: Path, data: bytes) -> str:
    """How an override file, whose bytes are ``data``, breaks the encoding,
    header or line ending rules of exported source; empty when it does not."""
    if data.startswith(b"\xef\xbb\xbf"):
        return "byte-order mark; VBA source is ASCII only"
    wide = next((byte for byte in data if byte > 0x7F), None)
    if wide is not None:
        return f"non-ASCII byte 0x{wide:02X}; VBA source is ASCII only"
    if re.search(rb"\r(?!\n)", data):
        return "carriage return without a line feed; end each line with CRLF or LF"
    text = module_text(path)
    if len(VB_NAME.findall(text)) != 1:
        return 'needs exactly one Attribute VB_Name = "<module>" line'
    if not text.startswith('Attribute VB_Name = "'):
        return ("does not start with its Attribute VB_Name line (after the VERSION and "
                "BEGIN ... END header of a class)")
    return ""


def read_overrides(folder: Path) -> tuple[Override, ...]:
    """The .bas and .cls files of ``folder`` and its subfolders, each with the
    module it names; BuildError on a form, an invalid file or a name used twice."""
    if not folder.is_dir():
        raise BuildError(f"override folder {folder}: not a folder")
    files = sorted((path for path in folder.rglob("*") if path.is_file() and
                    path.suffix.lower() in (*OVERRIDE_SUFFIXES, ".frm")),
                   key=lambda path: path.relative_to(folder).as_posix())
    forms = [path.relative_to(folder).as_posix() for path in files
             if path.suffix.lower() == ".frm"]
    if forms:
        raise BuildError(f"override folder {folder}: forms are not built ({', '.join(forms)}); "
                         "an override replaces a standard, class or document module")
    if not files:
        raise BuildError(f"override folder {folder}: holds no .bas or .cls file")
    found: dict[str, Override] = {}
    for path in files:
        data = path.read_bytes()
        problem = override_problem(path, data)
        if problem:
            raise BuildError(f"override {path}: {problem}")
        text = module_text(path)
        name = VB_NAME.findall(text)[0]
        other = found.get(name.casefold())
        if other is not None:
            raise BuildError(f"overrides {other.path} and {path} both name the module {name} "
                             "(VBA names are case-insensitive)")
        found[name.casefold()] = Override(name, path, hashlib.sha256(data).hexdigest(), text)
    return tuple(found.values())


def apply_overrides(modules: Sequence[Module], overrides: Sequence[Override],
                    configuration: str) -> list[Module]:
    """The modules with each override in place of the module it names: the same
    name, in the same case, and the same kind; BuildError otherwise."""
    replaced = list(modules)
    index = {module.name.casefold(): number for number, module in enumerate(modules)}
    for item in overrides:
        number = index.get(item.name.casefold())
        if number is None:
            raise BuildError(f"override {item.path}: VB_Name {item.name} matches no module of "
                             f"the {configuration} configuration; an override replaces a "
                             "module and never adds one")
        module = modules[number]
        if item.name != module.name:
            raise BuildError(f"override {item.path}: VB_Name {item.name} differs from the "
                             f"module {module.name} in case; spell it {module.name}")
        suffix = ".bas" if module.kind == STANDARD else ".cls"
        if item.path.suffix.lower() != suffix:
            raise BuildError(f"override {item.path}: {module.name} is a {module.kind} module, "
                             f"which only a {suffix} file replaces")
        replaced[number] = Module(module.name, item.source, kind=module.kind)
    return replaced


def folder_label(root: Path, folder: Path) -> str:
    """The override folder as the stamp records it: relative to the checkout when
    it lies inside it, absolute otherwise."""
    resolved = folder.resolve()
    try:
        return resolved.relative_to(root.resolve()).as_posix()
    except ValueError:
        return resolved.as_posix()


def is_clean(root: Path, layout: Layout, configuration: str) -> bool:
    """No tracked change, and no untracked file where the build reads source."""
    directories = [d for d in {*source_directories(layout, configuration), DOCUMENT_HOME}
                   if (root / d).is_dir()]
    tracked = git(root, "status", "--porcelain", "--untracked-files=no")
    untracked = git(root, "status", "--porcelain", "--untracked-files=all", "--",
                    *sorted(directories)) if directories else ""
    return not tracked and not untracked


def column_number(letters: str) -> int:
    number = 0
    for letter in letters:
        number = number * 26 + ord(letter) - 64
    return number


def column_letters(number: int) -> str:
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def quoted_tab(tab: str) -> str:
    return "'" + tab.replace("'", "''") + "'"


def commit_stamp(build: Build) -> str:
    """The stamp's commit, marked when the build is not evidence for it."""
    text = build.commit + ("" if build.clean else NOT_EVIDENCE)
    if build.overrides:
        text += (f"{OVERRIDDEN}{build.override_folder}: "
                 + ", ".join(f"{item.name} SHA-256 {item.sha256}" for item in build.overrides)
                 + ": NOT evidence-grade)")
    return text


def stamp_values(layout: Layout, build: Build) -> dict[str, str]:
    """The text of each stamp name."""
    return {layout.stamp["commit"]: commit_stamp(build),
            layout.stamp["time"]: build.time,
            layout.stamp["configuration"]: build.configuration}


def cell_position(cell: str) -> tuple[int, int]:
    """The (row, column) of a cell reference that layout_problems() accepted."""
    match = CELL.match(cell)
    assert match is not None
    return int(match.group(2)), column_number(match.group(1))


def sheet_cells(layout: Layout, build: Build, code_name: str) -> dict[tuple[int, int], str]:
    """The text cells of one sheet, by (row, column): stamp values and name labels."""
    stamp = stamp_values(layout, build)
    cells: dict[tuple[int, int], str] = {}
    stamp_sheet = next(item.sheet for item in layout.names
                       if item.name == layout.stamp["commit"])
    if code_name == stamp_sheet:
        cells[(1, 1)] = TITLE
    for item in layout.names:
        if item.sheet != code_name:
            continue
        row, column = cell_position(item.cell)
        if item.label and column > 1:
            cells[(row, column - 1)] = item.label
        if item.name in stamp:
            cells[(row, column)] = stamp[item.name]
    return cells


def inline_cell(row: int, column: int, text: str, style: str | None = None) -> str:
    """A cell holding ``text`` as an inline string, with the style index ``style``."""
    shown = "" if style is None else f" s={quoteattr(style)}"
    return (f'<c r="{column_letters(column)}{row}"{shown} t="inlineStr"><is>'
            f'<t xml:space="preserve">{escape(text)}</t></is></c>')


def worksheet_xml(code_name: str, cells: dict[tuple[int, int], str]) -> str:
    rows: dict[int, list[tuple[int, str]]] = {}
    for (row, column), text in cells.items():
        rows.setdefault(row, []).append((column, text))
    body = "".join(
        f'<row r="{row}">' + "".join(
            inline_cell(row, column, text) for column, text in sorted(rows[row])) + "</row>"
        for row in sorted(rows))
    return (f'{XML_HEADER}<worksheet xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}">'
            f"<sheetPr codeName={quoteattr(code_name)}/>"
            f"<sheetData>{body}</sheetData></worksheet>")


def absolute_reference(layout: Layout, item: DefinedName) -> str:
    """The reference of a defined name to its cell, such as Inputs!$C$5."""
    match = CELL.match(item.cell)
    assert match is not None
    return f"{quoted_tab(layout.sheet(item.sheet).tab)}!${match.group(1)}${match.group(2)}"


def defined_name(layout: Layout, item: DefinedName) -> str:
    return (f"<definedName name={quoteattr(item.name)}>"
            f"{escape(absolute_reference(layout, item))}</definedName>")


def workbook_xml(layout: Layout) -> str:
    sheets = "".join(f"<sheet name={quoteattr(sheet.tab)} sheetId=\"{index}\" "
                     f'r:id="rIdSheet{index}"/>'
                     for index, sheet in enumerate(layout.sheets, 1))
    names = ""
    for item in sorted(layout.names, key=lambda entry: entry.name.casefold()):
        names += defined_name(layout, item)
    return (f'{XML_HEADER}<workbook xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}">'
            f"<workbookPr codeName={quoteattr(layout.workbook_code_name)}/>"
            f"<bookViews><workbookView/></bookViews><sheets>{sheets}</sheets>"
            f"<definedNames>{names}</definedNames></workbook>")


STYLES = (f'{XML_HEADER}<styleSheet xmlns="{MAIN}">'
          '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
          '<fills count="2"><fill><patternFill patternType="none"/></fill>'
          '<fill><patternFill patternType="gray125"/></fill></fills>'
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
          '</cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" '
          'borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" '
          'xfId="0" builtinId="0"/></cellStyles></styleSheet>')


def package(layout: Layout, build: Build, vba_bin: bytes) -> bytes:
    """The .xlsm package: the same bytes for the same input."""
    count = len(layout.sheets)
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for index in range(1, count + 1))
    parts = {
        "[Content_Types].xml": (
            f'{XML_HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
            'content-types"><Default Extension="rels" ContentType="application/'
            'vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" '
            'ContentType="application/xml"/><Default Extension="bin" '
            'ContentType="application/vnd.ms-office.vbaProject"/><Override '
            'PartName="/xl/workbook.xml" ContentType="application/'
            'vnd.ms-excel.sheet.macroEnabled.main+xml"/><Override PartName="/xl/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.'
            f'styles+xml"/>{overrides}</Types>'),
        "_rels/.rels": (
            f'{XML_HEADER}<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship '
            f'Id="rId1" Type="{RELATIONSHIPS}/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>"),
        "xl/workbook.xml": workbook_xml(layout),
        "xl/_rels/workbook.xml.rels": (
            f'{XML_HEADER}<Relationships xmlns="{PACKAGE_RELATIONSHIPS}">' + "".join(
                f'<Relationship Id="rIdSheet{index}" Type="{WORKSHEET_TYPE}" '
                f'Target="worksheets/sheet{index}.xml"/>' for index in range(1, count + 1))
            + f'<Relationship Id="rIdStyles" Type="{STYLES_TYPE}" Target="styles.xml"/>'
            f'<Relationship Id="rIdVBA" Type="{VBA_TYPE}" Target="vbaProject.bin"/>'
            "</Relationships>"),
        "xl/styles.xml": STYLES,
    }
    for index, sheet in enumerate(layout.sheets, 1):
        parts[f"xl/worksheets/sheet{index}.xml"] = worksheet_xml(
            sheet.code_name, sheet_cells(layout, build, sheet.code_name))
    blobs: dict[str, bytes] = {name: text.encode("utf-8") for name, text in parts.items()}
    blobs[VBA_PART] = vba_bin
    return zip_package(blobs)


def zip_package(blobs: dict[str, bytes]) -> bytes:
    """The ZIP of the parts, in the given order, with fixed times and permissions."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in blobs.items():
            info = zipfile.ZipInfo(name, date_time=ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            # ZipInfo takes the creating system from sys.platform; fix it to Unix, the
            # system the permission bits below belong to, so every platform agrees.
            info.create_system = 3
            info.external_attr = 0o644 << 16
            archive.writestr(info, data)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Building from the layout template. Parts are read with ElementTree but edited
# as text, so that every namespace declaration and prefix the template uses,
# including those that mc:Ignorable names, stays as Excel wrote it.

ATTRIBUTE = re.compile(r"""([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')""")


@dataclass(frozen=True)
class TemplateSheet:
    tab: str
    part: str
    sheet_id: int


def read_template(path: Path) -> dict[str, bytes]:
    """The template's parts, in package order; BuildError if it carries macros."""
    try:
        with zipfile.ZipFile(path) as archive:
            parts = {info.filename: archive.read(info) for info in archive.infolist()
                     if not info.is_dir()}
    except (zipfile.BadZipFile, OSError) as error:
        raise BuildError(f"{TEMPLATE}: not a readable package ({error})") from error
    if any(MACRO_PART.search(name) for name in parts) or any(
            MACRO_PART.search(part_text(parts, name)) for name in (CONTENT_TYPES, WORKBOOK_RELS)):
        raise BuildError(f"{TEMPLATE}: carries macros; keep the template macro-free and the "
                         "code in exported source")
    return parts


def part_text(parts: dict[str, bytes], name: str) -> str:
    try:
        return parts[name].decode("utf-8")
    except KeyError as error:
        raise BuildError(f"{TEMPLATE}: has no {name}") from error
    except UnicodeDecodeError as error:
        raise BuildError(f"{TEMPLATE}: {name} is not UTF-8") from error


def parse_part(parts: dict[str, bytes], name: str) -> ElementTree.Element:
    try:
        return ElementTree.fromstring(parts[name])
    except KeyError as error:
        raise BuildError(f"{TEMPLATE}: has no {name}") from error
    except ElementTree.ParseError as error:
        raise BuildError(f"{TEMPLATE}: {name} is not XML ({error})") from error


def attributes(tag: str) -> dict[str, str]:
    """The attributes of a start tag, by name as written."""
    return {name: double or single for name, double, single in ATTRIBUTE.findall(tag)}


def part_name(target: str) -> str:
    """The package part that a target of the workbook's relationships names."""
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath(f"xl/{target}")


def workbook_relationships(parts: dict[str, bytes]) -> dict[str, str]:
    """The target of each relationship of the workbook part, by Id."""
    return {item.get("Id", ""): item.get("Target", "") for item in
            parse_part(parts, WORKBOOK_RELS).iter(f"{{{PACKAGE_RELATIONSHIPS}}}Relationship")}


def template_sheets(parts: dict[str, bytes]) -> list[TemplateSheet]:
    """The template's sheets in tab order, with their parts."""
    targets = workbook_relationships(parts)
    sheets = []
    for sheet in parse_part(parts, WORKBOOK_PART).iter(f"{{{MAIN}}}sheet"):
        target = targets.get(sheet.get(f"{{{RELATIONSHIPS}}}id", ""))
        if target is None:
            raise BuildError(f"{TEMPLATE}: sheet {sheet.get('name')!r} has no part")
        sheets.append(TemplateSheet(sheet.get("name", ""), part_name(target),
                                    int(sheet.get("sheetId", "0"))))
    return sheets


def same_cell(reference: str, tab: str, cell: str) -> bool:
    """Whether a defined name's reference is the absolute reference to that cell."""
    match = REFERENCE.match(reference)
    if match is None:
        return False
    found = match.group(1).replace("''", "'") if match.group(1) is not None else match.group(2)
    return found == tab and match.group(3) + match.group(4) == cell


def template_problems(layout: Layout, parts: dict[str, bytes]) -> list[str]:
    """How the template differs from the sheets and names the configuration builds."""
    kept = tuple(sheet for sheet in layout.sheets if sheet.role != TEST_ROLE)
    problems = []
    tabs = [sheet.tab for sheet in template_sheets(parts)]
    if tabs != [sheet.tab for sheet in kept]:
        problems.append(f"its sheets {tabs} are not the sheets {[sheet.tab for sheet in kept]} "
                        f"of {CONFIG} other than {TEST_ROLE}, in that order")
    if layout.sheets[:len(kept)] != kept:
        problems.append(f"{TEST_ROLE} sheets must come last in {CONFIG}, after the sheets of the "
                        "template")
    names = {item.get("name", "").casefold(): (item.text or "").strip() for item in
             parse_part(parts, WORKBOOK_PART).iter(f"{{{MAIN}}}definedName")
             if item.get("localSheetId") is None}
    codes = {sheet.code_name for sheet in kept}
    for item in layout.names:
        if item.sheet not in codes:
            continue
        found = names.get(item.name.casefold())
        if found is None:
            problems.append(f"it does not define the workbook name {item.name}")
        elif not same_cell(found, layout.sheet(item.sheet).tab, item.cell):
            problems.append(f"its name {item.name} refers to {found}, not "
                            f"{absolute_reference(layout, item)}")
    return problems


def main_part_problems(parts: dict[str, bytes]) -> list[str]:
    """Whether the package's main part is the workbook part, as Excel writes it."""
    targets = [item.get("Target", "").lstrip("/") for item in
               parse_part(parts, "_rels/.rels").iter(f"{{{PACKAGE_RELATIONSHIPS}}}Relationship")
               if item.get("Type") == DOCUMENT_TYPE]
    return [] if targets == [WORKBOOK_PART] else \
        [f"its main part is {targets}, not {WORKBOOK_PART}"]


def insert_before(text: str, marker: str, insertion: str) -> str:
    """The text with ``insertion`` just before the last ``marker``."""
    index = text.rfind(marker)
    if index < 0:
        raise BuildError(f"{TEMPLATE}: a part has no {marker}")
    return text[:index] + insertion + text[index:]


def with_attribute(tag: str, name: str, value: str) -> str:
    """A start tag with the attribute ``name`` set to ``value``."""
    attribute = f" {name}={quoteattr(value)}"
    pattern = re.compile(rf"""\s{name}\s*=\s*("[^"]*"|'[^']*')""")
    if pattern.search(tag):
        return pattern.sub(lambda _: attribute, tag, count=1)
    end = len(tag) - (2 if tag.endswith("/>") else 1)
    return tag[:end] + attribute + tag[end:]


def set_code_name(xml: str, element: str, after: Sequence[str], code_name: str) -> str:
    """The part with the codeName of ``element`` set. A missing element is added after
    the last of the tags ``after`` that the part holds: the root, then earlier siblings."""
    tag = re.search(rf"<{element}\b[^>]*?/?>", xml)
    if tag is not None:
        return xml[:tag.start()] + with_attribute(tag.group(0), "codeName", code_name) + \
            xml[tag.end():]
    ends = [found.end() for pattern in after if (found := re.search(pattern, xml))]
    if not ends:
        raise BuildError(f"{TEMPLATE}: a part has no {after[0]}")
    return xml[:max(ends)] + f"<{element} codeName={quoteattr(code_name)}/>" + xml[max(ends):]


def column_style(xml: str, column: int) -> str | None:
    """The style a sheet gives a column, from its <col> elements."""
    for tag in re.finditer(r"<col\b[^>]*>", xml):
        found = attributes(tag.group(0))
        if int(found.get("min", "0")) <= column <= int(found.get("max", "0")):
            return found.get("style")
    return None


def put_in_cells(cells: str, row: int, column: int, text: str, style: str | None) -> str:
    """A row's cells with ``text`` in ``column``: an existing cell keeps its own style."""
    for cell in CELL_ELEMENT.finditer(cells):
        found = attributes(cell.group(1))
        letters = re.match(r"[A-Z]+", found.get("r", ""))
        if letters is None:
            raise BuildError(f"{TEMPLATE}: a cell of row {row} has no reference")
        number = column_number(letters.group(0))
        if number < column:
            continue
        if number == column:
            return cells[:cell.start()] + inline_cell(row, column, text, found.get("s")) + \
                cells[cell.end():]
        return cells[:cell.start()] + inline_cell(row, column, text, style) + cells[cell.start():]
    return cells + inline_cell(row, column, text, style)


def set_cell(xml: str, row: int, column: int, text: str) -> str:
    """The sheet with ``text`` in one cell. An existing cell keeps its style; a new one
    takes its row's style, or else its column's, as Excel shows an empty cell."""
    data = re.search(r"<sheetData\b[^>]*?(/?)>", xml)
    if data is None:
        raise BuildError(f"{TEMPLATE}: a sheet has no sheetData")
    if data.group(1):
        xml = xml[:data.start()] + "<sheetData></sheetData>" + xml[data.end():]
    start = data.start() + len("<sheetData>") if data.group(1) else data.end()
    end = xml.index("</sheetData>", start)
    body = xml[start:end]
    for match in ROW.finditer(body):
        found = attributes(match.group(1))
        if not found.get("r", "").isdigit():
            raise BuildError(f"{TEMPLATE}: a row has no row number")
        if int(found["r"]) < row:
            continue
        if int(found["r"]) > row:
            new_row = f'<row r="{row}">{inline_cell(row, column, text, column_style(xml, column))}'
            body = body[:match.start()] + new_row + "</row>" + body[match.start():]
            break
        style = found.get("s") if found.get("customFormat") in ("1", "true") else \
            column_style(xml, column)
        if match.group(2):
            body = (body[:match.start()] + match.group(0)[:-2] + ">" +
                    inline_cell(row, column, text, style) + "</row>" + body[match.end():])
        else:
            close = body.index("</row>", match.end())
            body = body[:match.end()] + put_in_cells(body[match.end():close], row, column,
                                                     text, style) + body[close:]
        break
    else:
        body += f'<row r="{row}">{inline_cell(row, column, text, column_style(xml, column))}</row>'
    return xml[:start] + body + xml[end:]


def unused_id(prefix: str, taken: set[str]) -> str:
    candidate, number = prefix, 0
    while candidate in taken:
        number += 1
        candidate = f"{prefix}{number}"
    taken.add(candidate)
    return candidate


def relationship_attribute(workbook: str) -> str:
    """How a new element of the workbook part writes its relationship id."""
    root = re.search(r"<workbook\b[^>]*>", workbook)
    found = re.search(rf'xmlns:([\w.-]+)="{re.escape(RELATIONSHIPS)}"',
                      root.group(0) if root else "")
    return f"{found.group(1)}:id" if found else f'xmlns:r="{RELATIONSHIPS}" r:id'


def append_test_sheets(edited: dict[str, str], parts: dict[str, bytes], layout: Layout,
                       build: Build, ids: set[str]) -> dict[str, str]:
    """Append each Test sheet blank, with its names; return the new parts."""
    tests = [sheet for sheet in layout.sheets if sheet.role == TEST_ROLE]
    added: dict[str, str] = {}
    sheet_id = max((sheet.sheet_id for sheet in template_sheets(parts)), default=0)
    reference = relationship_attribute(edited[WORKBOOK_PART])
    for sheet in tests:
        number = 1
        while f"xl/worksheets/sheet{number}.xml" in {*parts, *added}:
            number += 1
        part, identifier, sheet_id = f"xl/worksheets/sheet{number}.xml", \
            unused_id("rIdBuildSheet", ids), sheet_id + 1
        added[part] = worksheet_xml(sheet.code_name, sheet_cells(layout, build, sheet.code_name))
        edited[WORKBOOK_RELS] = insert_before(
            edited[WORKBOOK_RELS], "</Relationships>",
            f'<Relationship Id="{identifier}" Type="{WORKSHEET_TYPE}" '
            f'Target="worksheets/sheet{number}.xml"/>')
        edited[CONTENT_TYPES] = insert_before(
            edited[CONTENT_TYPES], "</Types>",
            f'<Override PartName="/{part}" ContentType="{WORKSHEET_CONTENT}"/>')
        edited[WORKBOOK_PART] = insert_before(
            edited[WORKBOOK_PART], "</sheets>",
            f'<sheet name={quoteattr(sheet.tab)} sheetId="{sheet_id}" {reference}="{identifier}"/>')
    codes = {sheet.code_name for sheet in tests}
    names = "".join(defined_name(layout, item) for item in layout.names if item.sheet in codes)
    if names:
        edited[WORKBOOK_PART] = insert_before(edited[WORKBOOK_PART], "</definedNames>", names)
    return added


def macro_enabled(types: str) -> str:
    """The content types with the workbook part declared macro-enabled and the VBA part."""
    override = re.search(r'<Override\b[^>]*\bPartName="/xl/workbook\.xml"[^>]*/>', types, re.I)
    if override is None or SHEET_MAIN not in override.group(0):
        raise BuildError(f"{TEMPLATE}: /{WORKBOOK_PART} is not declared as {SHEET_MAIN}")
    types = types[:override.start()] + override.group(0).replace(SHEET_MAIN, MACRO_MAIN) + \
        types[override.end():]
    return insert_before(types, "</Types>",
                         f'<Override PartName="/{VBA_PART}" ContentType="{VBA_CONTENT}"/>')


def template_package(parts: dict[str, bytes], layout: Layout, build: Build,
                     vba_bin: bytes) -> bytes:
    """The .xlsm built from a copy of the template: the same bytes for the same input."""
    problems = template_problems(layout, parts) + main_part_problems(parts)
    if problems:
        raise BuildError(f"{TEMPLATE}: {problems[0]}")
    edited = {name: part_text(parts, name) for name in (CONTENT_TYPES, WORKBOOK_PART,
                                                         WORKBOOK_RELS)}
    stamp = stamp_values(layout, build)
    kept = [sheet for sheet in layout.sheets if sheet.role != TEST_ROLE]
    for sheet, placed in zip(kept, template_sheets(parts), strict=True):
        xml = set_code_name(part_text(parts, placed.part), "sheetPr", (r"<worksheet\b[^>]*>",),
                            sheet.code_name)
        for item in layout.names:
            if item.sheet == sheet.code_name and item.name in stamp:
                xml = set_cell(xml, *cell_position(item.cell), stamp[item.name])
        edited[placed.part] = xml
    edited[WORKBOOK_PART] = set_code_name(
        edited[WORKBOOK_PART], "workbookPr",
        (r"<workbook\b[^>]*>", r"<fileVersion\b[^>]*/>", r"<fileSharing\b[^>]*/>"),
        layout.workbook_code_name)
    ids = set(workbook_relationships(parts))
    added = append_test_sheets(edited, parts, layout, build, ids)
    edited[WORKBOOK_RELS] = insert_before(
        edited[WORKBOOK_RELS], "</Relationships>",
        f'<Relationship Id="{unused_id("rIdVBA", ids)}" Type="{VBA_TYPE}" '
        'Target="vbaProject.bin"/>')
    edited[CONTENT_TYPES] = macro_enabled(edited[CONTENT_TYPES])
    blobs = {name: edited[name].encode("utf-8") if name in edited else data
             for name, data in parts.items()}
    blobs.update((name, text.encode("utf-8")) for name, text in added.items())
    blobs[VBA_PART] = vba_bin
    return zip_package(blobs)


def for_configuration(layout: Layout, configuration: str) -> Layout:
    """The layout a configuration builds: Test sheets and their names only in
    the development configuration, as the workbook contract requires."""
    if configuration == DEVELOPMENT:
        return layout
    sheets = tuple(sheet for sheet in layout.sheets if sheet.role != TEST_ROLE)
    kept = {sheet.code_name for sheet in sheets}
    return Layout(layout.workbook_code_name, sheets,
                  tuple(item for item in layout.names if item.sheet in kept),
                  layout.stamp, layout.configurations)


def build(root: Path, configuration: str, target: Path, *,
          override: Path | None = None) -> Build:
    """Write the workbook to ``target`` and describe what it holds. With
    ``override``, the folder's modules replace those of the same name."""
    # Document modules are checked against every sheet, Test sheets included,
    # before the configuration drops what it does not build.
    complete = load_layout(root)
    paths = components(root, complete, configuration)
    layout = for_configuration(complete, configuration)
    modules = vba_modules(root, layout, paths)
    overrides: tuple[Override, ...] = ()
    if override is not None:
        overrides = read_overrides(override)
        modules = apply_overrides(modules, overrides, configuration)
    commit = git(root, "rev-parse", "HEAD")
    result = Build(commit, git(root, "show", "-s", "--format=%cI", "HEAD"),
                   is_clean(root, layout, configuration), configuration, tuple(paths),
                   (root / TEMPLATE).is_file(), overrides,
                   folder_label(root, override) if override is not None else "")
    project_id = "{" + str(uuid.uuid5(uuid.NAMESPACE_URL,
                                      f"workbook-build:{commit}:{configuration}")).upper() + "}"
    vba_bin = build_vba_project(modules, project_id)
    data = template_package(read_template(root / TEMPLATE), layout, result, vba_bin)         if result.template else package(layout, result, vba_bin)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(data)
    return result


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("output", nargs="?", type=Path,
                        help="the .xlsm to write (default: test-results/build/"
                             "Workbook-<configuration>.xlsm)")
    parser.add_argument("--configuration", default="development",
                        help=f"a configuration of {CONFIG} (default: development)")
    parser.add_argument("--override", type=Path, metavar="DIR",
                        help="replace each module named by a .bas or .cls file of DIR or its "
                             "subfolders; the stamp then marks the build NOT evidence-grade")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    root = args.root.resolve()
    target = args.output or root / "test-results" / "build" / \
        f"Workbook-{args.configuration}.xlsm"
    try:
        result = build(root, args.configuration, target, override=args.override)
    except (BuildError, ValueError) as error:
        print(f"FAIL: {error}")
        return 1
    names = ", ".join(path.stem for path in result.components) or \
        "no VBA component: the workbook holds only its document modules"
    grade = "" if result.clean else NOT_EVIDENCE
    print(f"built {target}: {result.configuration}, commit {result.commit}{grade}")
    print(f"layout: {TEMPLATE if result.template else 'blank sheets of ' + CONFIG}")
    print(f"components: {names}")
    if result.overrides:
        print(f"overrides from {result.override_folder}: NOT evidence for commit "
              f"{result.commit}")
        for item in result.overrides:
            print(f"  replaced {item.name} with {item.path}, SHA-256 {item.sha256}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
