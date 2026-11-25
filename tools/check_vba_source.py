#!/usr/bin/env python3
"""Check the placement, names, encoding, storage and mechanical style of VBA source.

Applies the mechanical rules of docs/VBA_HOUSE_STYLE.md and the placement
rules of docs/REPOSITORY_STRUCTURE.md to every tracked file: VBA components
only in their homes, each with its role's name prefix and a VB_Name equal to
its file name; ASCII only; Option Explicit, and Option Private Module where
required; no Option Base; no tab or trailing white space, except in the
VBE header of a form; string literals closed; forms beside their .frx; LF
storage with CRLF checkout; no workbook committed except the macro-free
src/workbook/Template.xlsx and the byte-preserved copies of regulatory
sources under docs/methodology/sources/.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Sequence
from pathlib import Path

import gitio
from format_vba import header_length
from vba_lexer import (
    VBA_SUFFIXES,
    Finding,
    facade_pattern,
    logical_units,
    read_component,
    report,
    split_statements,
)

ROOT = Path(__file__).resolve().parents[1]
# Component name pattern by (home, suffix): docs/VBA_HOUSE_STYLE.md, Naming and
# visibility. The facade's prefix comes from that table (vba_lexer.facade_pattern).
# Any other suffix in a home, or a component outside these homes, is misplaced.
NAMING: dict[tuple[str, str], tuple[str, re.Pattern[str]]] = {
    ("src/core/", ".bas"): ("CORE_<Subject>", re.compile(r"CORE_[A-Z][A-Za-z0-9_]*")),
    ("src/modules/", ".bas"): ("<REGIME>_<Subject>", re.compile(r"[A-Z][A-Z0-9]*_[A-Za-z0-9_]*")),
    ("src/classes/", ".cls"): ("C<Noun>", re.compile(r"C[A-Z][A-Za-z0-9]*")),
    ("src/workbook/", ".cls"): ("ThisWorkbook or sh<Noun>",
                                re.compile(r"ThisWorkbook|sh[A-Z][A-Za-z0-9]*")),
    ("src/workbook/", ".bas"): ("M_<Subject>", re.compile(r"M_[A-Z][A-Za-z0-9_]*")),
    ("src/forms/", ".frm"): ("frm<Noun>", re.compile(r"frm[A-Z][A-Za-z0-9]*")),
    ("tests/modules/", ".bas"): ("TEST_<Subject>", re.compile(r"TEST_[A-Z][A-Za-z0-9_]*")),
    ("examples/modules/", ".bas"): ("EX_<Subject>", re.compile(r"EX_[A-Z][A-Za-z0-9_]*")),
}
HOMES = tuple(sorted({home for home, _ in NAMING}))
# Standard modules whose Public procedures must stay inside the project.
PRIVATE_MODULE_HOMES = ("src/core/", "tests/modules/")
MAX_NAME = 31  # the VBE rejects a longer component name
# Workbooks are build outputs. Only the macro-free template and the
# byte-preserved copies of regulatory sources may be committed
# (docs/REPOSITORY_STRUCTURE.md, Locations).
WORKBOOK_SUFFIXES = frozenset({".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm",
                               ".xlam", ".xla", ".xlw"})
APPROVED_WORKBOOKS: frozenset[str] = frozenset({"src/workbook/Template.xlsx"})
SOURCE_COPIES = "docs/methodology/sources/"

VB_NAME = re.compile(r'^Attribute VB_Name = "([^"]*)"[ \t]*$', re.M)
# Option statements, matched against comment-free statements.
OPTION_EXPLICIT = re.compile(r"Option\s+Explicit", re.I)
OPTION_PRIVATE = re.compile(r"Option\s+Private\s+Module", re.I)
# Explicit array bounds, never Option Base (docs/VBA_HOUSE_STYLE.md).
OPTION_BASE = re.compile(r"Option\s+Base\b", re.I)
FRX_REFERENCE = re.compile(r'"([^"\n]+\.frx)":([0-9A-Fa-f]+)')
REM_STATEMENT = re.compile(r"Rem(?=[ \t]|$)", re.I)
LINE_NUMBER = re.compile(r"[ \t]*\d+(?=[ \t])")


def placement(path: str, naming: dict[tuple[str, str], tuple[str, re.Pattern[str]]] = NAMING
              ) -> list[Finding]:
    """A component, or a form's .frx, must sit in its home with its role's name."""
    suffix = Path(path).suffix.lower()
    if suffix == ".frx":
        return [] if path.startswith("src/forms/") and path.count("/") == 2 else \
            [Finding(path, None, "a form resource belongs beside its form in src/forms/")]
    home = next((home for home in HOMES if path.startswith(home)), None)
    if home is None or "/" in path[len(home):]:
        return [Finding(path, None, f"VBA component outside its homes ({', '.join(HOMES)}, "
                                    "no subdirectories)")]
    rule = naming.get((home, suffix))
    if rule is None:
        return [Finding(path, None, f"{home} holds no {suffix} components")]
    name = Path(path).stem
    findings = []
    if not rule[1].fullmatch(name):
        findings.append(Finding(path, None, f"a {suffix} component in {home} is named "
                                            f"{rule[0]}"))
    if len(name) > MAX_NAME:
        findings.append(Finding(path, None, f"component name longer than {MAX_NAME} characters"))
    return findings


def non_ascii(path: str, data: bytes) -> list[Finding]:
    if data.startswith(b"\xef\xbb\xbf"):
        return [Finding(path, 1, "byte-order mark; VBA source is ASCII only")]
    return [Finding(path, number, f"non-ASCII byte 0x{byte:02X}; VBA source is ASCII only, "
                                  "build the character with ChrW")
            for number, line in enumerate(data.split(b"\n"), 1)
            for byte in sorted(set(line)) if byte > 0x7F][:20]


def unclosed_literals(text: str) -> list[int]:
    """Lines whose code leaves a string literal open.

    A ' outside a literal and outside [...] (which quotes a sheet name, as in
    ['Sheet 1'!A1]), or Rem at the start of a statement, also after a line
    number, begins a comment; a comment ending in " _" continues on the next
    line.
    """
    open_lines: list[int] = []
    in_comment = False
    for number, line in enumerate(text.split("\n"), 1):
        if in_comment:
            in_comment = line.rstrip().endswith(" _")
            continue
        quoted, statement_start, brackets = False, True, 0
        label = LINE_NUMBER.match(line)
        for index in range(label.end() if label else 0, len(line)):
            char = line[index]
            if quoted:
                quoted = char != '"'
            elif char == '"':
                quoted = True
            elif (char == "'" and not brackets) or \
                    (statement_start and REM_STATEMENT.match(line, index)):
                in_comment = line.rstrip().endswith(" _")
                break
            elif char == "[":
                brackets += 1
            elif char == "]" and brackets:
                brackets -= 1
            statement_start = (char == ":" and not quoted) or \
                (statement_start and char.isspace())
        if quoted:
            open_lines.append(number)
    return open_lines


def layout(path: str, text: str) -> list[Finding]:
    """Tabs and trailing white space; docs/VBA_HOUSE_STYLE.md, Source representation.

    The VBE header of a form, as tools/format_vba.py finds it, keeps the bytes
    the VBE writes, a trailing blank on its Begin line included, so its lines
    are not checked; the code after it is.
    """
    lines = text.split("\n")
    skip = header_length(lines) if Path(path).suffix.lower() == ".frm" else 0
    findings = []
    for number, line in enumerate(lines[skip:], skip + 1):
        if "\t" in line:
            findings.append(Finding(path, number, "tab character; indent with spaces"))
        if line != line.rstrip(" \t"):
            findings.append(Finding(path, number, "trailing white space"))
    return findings


def form_resources(root: Path, path: str, text: str, tracked: set[str]) -> list[Finding]:
    references = FRX_REFERENCE.findall(text)
    if not references:
        return [Finding(path, None, "no form resource (.frx) reference")]
    findings = []
    for filename, offset in references:
        companion = f"{Path(path).parent.as_posix()}/{filename}"
        if Path(filename).name != filename or companion not in tracked:
            findings.append(Finding(path, None, f"form resource {filename} is not tracked "
                                                "beside the form"))
        elif (root / companion).stat().st_size <= int(offset, 16):
            findings.append(Finding(path, None, f"resource offset is outside {filename}"))
    return findings


def content(root: Path, path: str, tracked: set[str], names: dict[str, str]) -> list[Finding]:
    data = (root / path).read_bytes()
    text = read_component(root, path)
    findings = non_ascii(path, data)
    vb_names = VB_NAME.findall(text)
    stem = Path(path).stem
    if vb_names != [stem]:
        findings.append(Finding(path, None, f'needs exactly one Attribute VB_Name = "{stem}"'))
    elif stem.casefold() in names:
        findings.append(Finding(path, None, f"component name also used by {names[stem.casefold()]}"
                                            " (VBA names are case-insensitive)"))
    else:
        names[stem.casefold()] = path
    suffix = Path(path).suffix.lower()
    if suffix == ".bas" and not text.startswith('Attribute VB_Name = "'):
        findings.append(Finding(path, 1, "a .bas export starts with its Attribute VB_Name line"))
    # Comments, continued ones included, are not statements.
    options = [(line, statement) for line, _, kind, code in logical_units(text.split("\n"))
               if kind == "code" for statement in split_statements(code)]
    if not any(OPTION_EXPLICIT.fullmatch(statement) for _, statement in options):
        findings.append(Finding(path, None, "missing Option Explicit"))
    if suffix == ".bas" and path.startswith(PRIVATE_MODULE_HOMES) and \
            not any(OPTION_PRIVATE.fullmatch(statement) for _, statement in options):
        findings.append(Finding(path, None, "missing Option Private Module"))
    findings.extend(Finding(path, line, "Option Base; declare explicit array bounds")
                    for line, statement in options if OPTION_BASE.match(statement))
    findings.extend(Finding(path, number, "string literal not closed")
                    for number in unclosed_literals(text))
    findings.extend(layout(path, text))
    if suffix == ".frm":
        findings.extend(form_resources(root, path, text, tracked))
    return findings


def storage(root: Path, paths: list[str]) -> list[Finding]:
    """Exported VBA is stored with LF and checked out with CRLF (.gitattributes)."""
    if not paths:
        return []
    findings = []
    for record in gitio.git(root, "ls-files", "--eol", "-z", "--", *paths).split(b"\0"):
        if not record:
            continue
        info, _, path = record.decode("utf-8", errors="surrogateescape").partition("\t")
        fields = info.split()
        index_eol, attributes = fields[0].removeprefix("i/"), " ".join(fields[2:])
        if index_eol not in {"lf", "none", ""}:
            findings.append(Finding(path, None, f"stored with {index_eol} in Git; store with LF"))
        if "eol=crlf" not in attributes:
            findings.append(Finding(path, None, ".gitattributes must check it out with CRLF"))
    return findings


def check(root: Path) -> tuple[list[Finding], int]:
    tracked = gitio.tracked(root)
    tracked_set = set(tracked)
    components = [p for p in tracked if Path(p).suffix.lower() in VBA_SUFFIXES]
    findings: list[Finding] = []
    label, pattern = facade_pattern(root)
    naming = {**NAMING, ("src/modules/", ".bas"): (label, re.compile(pattern))}
    for path in tracked:
        suffix = Path(path).suffix.lower()
        if suffix in VBA_SUFFIXES or suffix == ".frx":
            findings.extend(placement(path, naming))
        if suffix in WORKBOOK_SUFFIXES and path not in APPROVED_WORKBOOKS and \
                not path.startswith(SOURCE_COPIES):
            findings.append(Finding(path, None, "workbook committed; workbooks are built from "
                                                "source, and only src/workbook/Template.xlsx "
                                                f"and copies under {SOURCE_COPIES} may be "
                                                "committed"))
    names: dict[str, str] = {}
    for path in components:
        findings.extend(content(root, path, tracked_set, names))
    findings.extend(storage(root, components))
    return findings, len(components)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components = check(args.root.resolve())
    return report("vba-source", findings, f"{components} VBA component(s)")


if __name__ == "__main__":
    sys.exit(main())
