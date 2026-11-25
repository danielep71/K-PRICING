"""Read back the built workbook and its VBA project with an independent reader.

tools/devbuild/vba_project.py writes vbaProject.bin from the [MS-OVBA] and
[MS-CFB] specifications. These tests decompress and parse its output with a
separate, minimal implementation of the same specifications, so a writer
defect fails here without Excel. They also check that
tools/devbuild/build_workbook.py builds the sheets, code names, defined names
and stamp of tools/devbuild/workbook.json, that this configuration matches
docs/methodology/WORKBOOK_CONTRACT.md, that a rebuild is byte-identical,
from blank sheets and from a synthetic layout template, and that --override
replaces exactly the modules it names and marks the stamp. The builds of a
project's own modules and of the template use the template's synthetic build
configuration (tests/workbook_template.py), not this repository's sheets.
"""

from __future__ import annotations

import contextlib
import hashlib
import io
import json
import random
import shutil
import struct
import sys
import tempfile
import unittest
import unittest.mock
import zipfile
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))
import build_workbook  # noqa: E402
import check_binaries  # noqa: E402
import check_excel_run  # noqa: E402
import export_manifest  # noqa: E402
import gitfixture  # noqa: E402
import lo_smoke  # noqa: E402
import vba_project  # noqa: E402
import workbook_template  # noqa: E402

ENDOFCHAIN = 0xFFFFFFFE


def decompress(data: bytes) -> bytes:
    """[MS-OVBA] 2.4.1.3.1: decompress a CompressedContainer."""
    if not data or data[0] != 0x01:
        raise ValueError("missing signature byte")
    out = bytearray()
    pos = 1
    while pos < len(data):
        start = pos
        header = struct.unpack_from("<H", data, pos)[0]
        pos += 2
        if (header >> 12) & 0x7 != 0b011:
            raise ValueError("bad chunk signature")
        end = min(start + (header & 0x0FFF) + 3, len(data))
        if not header & 0x8000:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        chunk_start = len(out)
        while pos < end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[pos])
                    pos += 1
                    continue
                token = struct.unpack_from("<H", data, pos)[0]
                pos += 2
                difference = len(out) - chunk_start
                bit_count = 4
                while (1 << bit_count) < difference:
                    bit_count += 1
                length_mask = 0xFFFF >> bit_count
                length = (token & length_mask) + 3
                offset = (token >> (16 - bit_count)) + 1
                source = len(out) - offset
                for i in range(length):
                    out.append(out[source + i])
    return bytes(out)


class CompoundFile:
    """[MS-CFB] version-3 reader: FAT, directory tree and mini stream."""

    def __init__(self, blob: bytes) -> None:
        if blob[:8] != bytes.fromhex("D0CF11E0A1B11AE1"):
            raise ValueError("not a compound file")
        self.blob = blob
        self.sector = 1 << struct.unpack_from("<H", blob, 30)[0]
        self.mini = 1 << struct.unpack_from("<H", blob, 32)[0]
        n_fat, dir_start = struct.unpack_from("<II", blob, 44)
        self.cutoff, minifat_start = struct.unpack_from("<II", blob, 56)
        difat = struct.unpack_from("<109I", blob, 76)[:n_fat]
        fat_bytes = b"".join(self._sector(s) for s in difat)
        self.fat = list(struct.unpack(f"<{len(fat_bytes) // 4}I", fat_bytes))
        directory = self._chain_bytes(dir_start)
        self.entries = [self._entry(directory[i:i + 128]) for i in range(0, len(directory), 128)]
        root = self.entries[0]
        self.mini_stream = self._chain_bytes(root["start"])[:root["size"]]
        minifat_bytes = self._chain_bytes(minifat_start) if minifat_start != ENDOFCHAIN else b""
        self.minifat = list(struct.unpack(f"<{len(minifat_bytes) // 4}I", minifat_bytes))

    def _sector(self, index: int) -> bytes:
        offset = self.sector * (index + 1)
        return self.blob[offset:offset + self.sector]

    def _chain(self, start: int, table: list[int]) -> list[int]:
        chain, seen = [], set()
        while start != ENDOFCHAIN:
            if start in seen:
                raise ValueError("cyclic chain")
            seen.add(start)
            chain.append(start)
            start = table[start]
        return chain

    def _chain_bytes(self, start: int) -> bytes:
        return b"".join(self._sector(s) for s in self._chain(start, self.fat))

    @staticmethod
    def _entry(raw: bytes) -> dict[str, Any]:
        name_length = struct.unpack_from("<H", raw, 64)[0]
        return {"name": raw[:max(0, name_length - 2)].decode("utf-16-le"), "kind": raw[66],
                "left": struct.unpack_from("<I", raw, 68)[0], "right": struct.unpack_from("<I", raw, 72)[0],
                "child": struct.unpack_from("<I", raw, 76)[0], "start": struct.unpack_from("<I", raw, 116)[0],
                "size": struct.unpack_from("<Q", raw, 120)[0]}

    def children(self, sid: int) -> dict[str, int]:
        found: dict[str, int] = {}
        stack = [self.entries[sid]["child"]]
        while stack:
            node = stack.pop()
            if node in (0xFFFFFFFF, None):
                continue
            entry = self.entries[node]
            found[entry["name"]] = node
            stack += [entry["left"], entry["right"]]
        return found

    def read(self, path: str) -> bytes:
        sid = 0
        for part in path.split("/"):
            sid = self.children(sid)[part]
        entry = self.entries[sid]
        if entry["size"] >= self.cutoff:
            return self._chain_bytes(entry["start"])[:entry["size"]]
        if entry["size"] == 0:
            return b""
        chunks = [self.mini_stream[s * self.mini:(s + 1) * self.mini] for s in self._chain(entry["start"], self.minifat)]
        return b"".join(chunks)[:entry["size"]]


def sample_modules() -> list[vba_project.Module]:
    body = "\r\n".join(f"' line {i}: some repeated text, repeated text, repeated text" for i in range(200))
    return [vba_project.document_module("ThisWorkbook", True),
            vba_project.document_module("shChecks", False),
            vba_project.Module("CORE_Small", 'Attribute VB_Name = "CORE_Small"\r\nOption Explicit\r\n'),
            vba_project.Module("TEST_Large", 'Attribute VB_Name = "TEST_Large"\r\n' + body + "\r\n")]


class Compression(unittest.TestCase):
    def test_specification_example_decompresses(self) -> None:
        # [MS-OVBA] 3.2.3, "Normal Compression" example.
        compressed = bytes.fromhex(
            "012FB000236161616263646582660070616768696A01380861"
            "6B6C00306D6E6F7006710270041072737475761077"
            "78797A00" "3C")
        self.assertEqual(decompress(compressed), b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa")

    def test_round_trip(self) -> None:
        rng = random.Random(8)
        samples = [b"", b"a", b"abc" * 3000, bytes(rng.randrange(256) for _ in range(9000)),
                   b"Option Explicit\r\n" * 700, bytes(range(256)) * 40]
        for data in samples:
            with self.subTest(length=len(data)):
                self.assertEqual(decompress(vba_project.compress(data)), data)


    def test_incompressible_last_chunk_names_the_module(self) -> None:
        # About 4 KB of random printable text in the last, short chunk does not fit
        # one compressed chunk; the build stops and names the module.
        rng = random.Random(97)
        noise = "".join(chr(rng.randrange(33, 127)) for _ in range(4000))
        module = vba_project.Module("CORE_Noise", f'Attribute VB_Name = "CORE_Noise"\r\n\' {noise}\r\n')
        with self.assertRaisesRegex(ValueError, r"module CORE_Noise: .*INSTALLATION.md"):
            vba_project.build_vba_project([module], ProjectStorage.project_id)


class Encryption(unittest.TestCase):
    def test_project_protection_fields_round_trip(self) -> None:
        rng = random.Random("{ID}")
        for data in (b"\x00", b"\xff", struct.pack("<I", 0), b"abcdefgh"):
            self.assertEqual(vba_project.decrypt(vba_project._encrypt(data, "{ID}", rng)), data)


class ProjectStorage(unittest.TestCase):
    project_id = "{0B9A3C1D-1111-4222-8333-444455556666}"

    def build(self) -> CompoundFile:
        return CompoundFile(vba_project.build_vba_project(sample_modules(), self.project_id))

    def test_every_module_source_reads_back_unchanged(self) -> None:
        cfb = self.build()
        for module in sample_modules():
            with self.subTest(module=module.name):
                self.assertEqual(decompress(cfb.read(f"VBA/{module.name}")).decode("cp1252"), module.source)

    def test_project_and_dir_streams_list_every_module(self) -> None:
        cfb = self.build()
        project = cfb.read("PROJECT").decode("cp1252")
        self.assertIn(f'ID="{self.project_id}"', project)
        self.assertIn("Document=ThisWorkbook/&H00000000", project)
        self.assertIn("Module=CORE_Small", project)
        directory = decompress(cfb.read("VBA/dir"))
        for module in sample_modules():
            self.assertIn(module.name.encode("cp1252"), directory)
        self.assertEqual(cfb.read("VBA/_VBA_PROJECT")[:2], b"\xcc\x61")

    def test_output_is_deterministic_for_a_project_id(self) -> None:
        first = vba_project.build_vba_project(sample_modules(), self.project_id)
        self.assertEqual(first, vba_project.build_vba_project(sample_modules(), self.project_id))
        self.assertNotEqual(first, vba_project.build_vba_project(sample_modules(), "{" + "1" * 8 + "-0000-0000-0000-" + "0" * 12 + "}"))


    def test_class_modules_are_declared_as_classes(self) -> None:
        module = vba_project.Module("CPair", 'Attribute VB_Name = "CPair"\r\nOption Explicit\r\n',
                                    kind=vba_project.CLASS)
        cfb = CompoundFile(vba_project.build_vba_project([module], self.project_id))
        self.assertIn("Class=CPair", cfb.read("PROJECT").decode("cp1252"))
        directory = decompress(cfb.read("VBA/dir"))
        # MODULETYPE 0x0022: a document, class or designer module.
        self.assertIn(struct.pack("<HI", 0x0022, 0), directory)
        with self.assertRaisesRegex(ValueError, "unknown kind"):
            vba_project.Module("X", "", kind="form")


def contract_rows(heading: str) -> list[list[str]]:
    """The body rows of the first table under a heading of the workbook contract,
    each cell without its Illustrative marker and code formatting."""
    text = (ROOT / "docs/methodology/WORKBOOK_CONTRACT.md").read_text(encoding="utf-8")
    section = text.split(f"## {heading}\n", 1)[1].split("\n## ", 1)[0]
    lines = [line for line in section.splitlines() if line.startswith("|")][2:]
    return [[cell.strip().replace("*Illustrative:*", "").strip().strip("`")
             for cell in line.strip("|").split("|")] for line in lines]


class Layout(unittest.TestCase):
    layout = build_workbook.load_layout(ROOT)

    def test_sheets_and_names_match_the_workbook_contract(self) -> None:
        sheets = [(row[1].strip("`"), row[2].strip("`"), row[3].strip("`"))
                  for row in contract_rows("Sheets")]
        self.assertEqual([(sheet.tab, sheet.code_name, sheet.role)
                          for sheet in self.layout.sheets], sheets)
        names = [(row[0], row[1]) for row in contract_rows("Defined names")]
        self.assertEqual([(item.name, item.sheet) for item in self.layout.names], names)

    def test_configurations_follow_the_structure_guide(self) -> None:
        structure = (ROOT / "docs/REPOSITORY_STRUCTURE.md").read_text(encoding="utf-8")
        for directories in self.layout.configurations.values():
            for directory in directories:
                with self.subTest(directory=directory):
                    self.assertIn(f"| `{directory}/` |", structure)
        production = set(self.layout.configurations["production"])
        self.assertFalse({"tests/modules", "examples/modules"} & production)
        self.assertLess(production, set(self.layout.configurations["development"]))

    def test_invalid_layouts_are_refused(self) -> None:
        base = self.layout
        # A sheet of the configuration, whatever the repository names it.
        sheet = base.sheets[0].code_name
        cases = {
            "no sheet": build_workbook.Layout(base.workbook_code_name, (), base.names,
                                              base.stamp, base.configurations),
            "code name is used twice": build_workbook.Layout(
                sheet, base.sheets, base.names, base.stamp, base.configurations),
            "not one cell": build_workbook.Layout(
                base.workbook_code_name, base.sheets,
                (*base.names, build_workbook.DefinedName("cfg_x", sheet, "C5:C6", "")),
                base.stamp, base.configurations),
            "unknown sheet": build_workbook.Layout(
                base.workbook_code_name, base.sheets,
                (*base.names, build_workbook.DefinedName("cfg_x", "shNone", "A1", "")),
                base.stamp, base.configurations),
            "stamp commit": build_workbook.Layout(
                base.workbook_code_name, base.sheets, base.names,
                {**base.stamp, "commit": "missing"}, base.configurations),
            "has role 'Scratch'": build_workbook.Layout(
                base.workbook_code_name,
                (*base.sheets, build_workbook.Sheet("Scratch", "shScratch", "Scratch")),
                base.names, base.stamp, base.configurations),
            "on a Test sheet": build_workbook.Layout(
                base.workbook_code_name,
                (*base.sheets, build_workbook.Sheet("Checks", "shChecks", "Test")),
                (*base.names, build_workbook.DefinedName("test_commit", "shChecks", "A1", "")),
                {**base.stamp, "commit": "test_commit"}, base.configurations),
            "not a valid name": build_workbook.Layout(
                base.workbook_code_name, base.sheets,
                (*base.names, build_workbook.DefinedName("AB12", sheet, "A1", "")),
                base.stamp, base.configurations),
        }
        for expected, layout in cases.items():
            with self.subTest(expected=expected):
                problems = build_workbook.layout_problems(layout)
                self.assertTrue(any(expected in problem for problem in problems), problems)
        self.assertEqual(build_workbook.layout_problems(base), [])

    def test_test_sheets_are_built_only_in_development(self) -> None:
        base = self.layout
        layout = build_workbook.Layout(
            base.workbook_code_name,
            (*base.sheets, build_workbook.Sheet("Checks", "shChecks", "Test")),
            (*base.names, build_workbook.DefinedName("Checks_Results", "shChecks", "A3", "")),
            base.stamp, base.configurations)
        self.assertEqual(build_workbook.layout_problems(layout), [])
        development = build_workbook.for_configuration(layout, "development")
        production = build_workbook.for_configuration(layout, "production")
        self.assertIn("shChecks", [sheet.code_name for sheet in development.sheets])
        self.assertEqual(production.sheets, base.sheets)
        self.assertEqual(production.names, base.names)
        package = build_workbook.package(
            production, build_workbook.Build("0" * 40, "t", True, "production", ()), b"")
        with zipfile.ZipFile(io.BytesIO(package)) as archive:
            book = archive.read("xl/workbook.xml").decode()
        self.assertNotIn("Checks", book)


@contextlib.contextmanager
def source_tree(files: dict[str, bytes]) -> Any:
    """A directory with the build configuration and the given files."""
    with tempfile.TemporaryDirectory(prefix="devbuild-") as temporary:
        root = Path(temporary)
        (root / "tools/devbuild").mkdir(parents=True)
        shutil.copy(ROOT / build_workbook.CONFIG, root / build_workbook.CONFIG)
        for name, data in files.items():
            (root / name).parent.mkdir(parents=True, exist_ok=True)
            (root / name).write_bytes(data)
        yield root


class Components(unittest.TestCase):
    layout = build_workbook.load_layout(ROOT)

    def test_standard_and_class_modules_by_configuration(self) -> None:
        # A document module of one of the configuration's sheets is not a component.
        document = f"src/workbook/{self.layout.sheets[0].code_name}.cls"
        files = {"src/core/CORE_A.bas": b"x", "src/classes/CPair.cls": b"x",
                 "src/workbook/M_Run.bas": b"x", "tests/modules/TEST_A.bas": b"x",
                 document: b"x", "docs/notes.bas": b"x"}
        with source_tree(files) as root:
            names = {configuration: [path.relative_to(root).as_posix() for path in
                                     build_workbook.components(root, self.layout, configuration)]
                     for configuration in ("production", "development")}
        self.assertEqual(names["production"],
                         ["src/core/CORE_A.bas", "src/classes/CPair.cls", "src/workbook/M_Run.bas"])
        self.assertEqual(names["development"], [*names["production"], "tests/modules/TEST_A.bas"])

    def test_forms_and_unknown_document_modules_stop_the_build(self) -> None:
        with source_tree({"src/forms/frmA.frm": b"x"}) as root, \
                self.assertRaisesRegex(build_workbook.BuildError, "forms are not built"):
            build_workbook.components(root, self.layout, "production")
        with source_tree({"src/workbook/shOther.cls": b"x"}) as root, \
                self.assertRaisesRegex(build_workbook.BuildError, "matches no sheet"):
            build_workbook.components(root, self.layout, "production")
        with self.assertRaisesRegex(build_workbook.BuildError, "unknown configuration"):
            build_workbook.components(ROOT, self.layout, "staging")

    def test_module_text_drops_the_class_header_only(self) -> None:
        exported = (b'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\n'
                    b'Attribute VB_Name = "CPair"\r\nOption Explicit\r\n')
        standard = b'Attribute VB_Name = "CORE_A"\nOption Explicit'
        with source_tree({"src/classes/CPair.cls": exported,
                          "src/core/CORE_A.bas": standard}) as root:
            self.assertEqual(build_workbook.module_text(root / "src/classes/CPair.cls"),
                             'Attribute VB_Name = "CPair"\r\nOption Explicit\r\n')
            self.assertEqual(build_workbook.module_text(root / "src/core/CORE_A.bas"),
                             'Attribute VB_Name = "CORE_A"\r\nOption Explicit\r\n')


class Workbook(unittest.TestCase):
    layout = build_workbook.load_layout(ROOT)

    def build(self, configuration: str = "development") -> tuple[bytes, build_workbook.Build]:
        with tempfile.TemporaryDirectory(prefix="devbuild-") as temporary:
            target = Path(temporary) / "Workbook.xlsm"
            result = build_workbook.build(ROOT, configuration, target)
            return target.read_bytes(), result

    def test_a_rebuild_is_byte_identical(self) -> None:
        first, result = self.build()
        second, _ = self.build()
        self.assertEqual(first, second)
        self.assertEqual(result.commit, build_workbook.git(ROOT, "rev-parse", "HEAD"))

    def test_the_package_is_the_same_on_windows(self) -> None:
        # ZipInfo records the creating system from sys.platform: 0 on Windows, 3 elsewhere.
        build = build_workbook.Build("a" * 40, "2026-11-02T10:00:00+01:00", True,
                                     "development", ())
        here = build_workbook.package(self.layout, build, b"project")
        with unittest.mock.patch("zipfile.sys.platform", "win32"):
            windows = build_workbook.package(self.layout, build, b"project")
        self.assertEqual(here, windows)

    def test_the_package_holds_the_contract_layout_and_the_stamp(self) -> None:
        data, result = self.build()
        with zipfile.ZipFile(io.BytesIO(data)) as package:
            self.assertIn("macroEnabled", package.read("[Content_Types].xml").decode())
            book = package.read("xl/workbook.xml").decode()
            self.assertIn(f'codeName="{self.layout.workbook_code_name}"', book)
            for index, sheet in enumerate(self.layout.sheets, 1):
                self.assertIn(f'<sheet name="{sheet.tab}" sheetId="{index}"', book)
                self.assertIn(f'<sheetPr codeName="{sheet.code_name}"/>',
                              package.read(f"xl/worksheets/sheet{index}.xml").decode())
            for item in self.layout.names:
                self.assertRegex(book, rf'<definedName name="{item.name}">[^<]*'
                                       rf'!\$[A-Z]+\$[0-9]+</definedName>')
            for info in package.infolist():
                self.assertEqual(info.date_time, build_workbook.ZIP_TIME)
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "Workbook.xlsm"
            path.write_bytes(data)
            stamp = check_excel_run.read_stamp(path, self.layout.stamp["commit"])
            self.assertEqual(stamp, result.commit + ("" if result.clean
                                                      else build_workbook.NOT_EVIDENCE))
            self.assertEqual(check_excel_run.read_stamp(path, self.layout.stamp["configuration"]),
                             "development")
            self.assertEqual(check_excel_run.read_stamp(path, self.layout.stamp["time"]),
                             result.time)

    def test_the_project_embeds_every_component_and_document_module(self) -> None:
        data, result = self.build()
        with zipfile.ZipFile(io.BytesIO(data)) as package:
            cfb = CompoundFile(package.read("xl/vbaProject.bin"))
        project = cfb.read("PROJECT").decode("cp1252")
        for code in [self.layout.workbook_code_name, *(s.code_name for s in self.layout.sheets)]:
            with self.subTest(document=code):
                self.assertIn(f"Document={code}/&H00000000", project)
                self.assertIn(f'Attribute VB_Name = "{code}"',
                              decompress(cfb.read(f"VBA/{code}")).decode("cp1252"))
        self.assertTrue(result.components)
        for path in result.components:
            with self.subTest(component=path.stem):
                self.assertEqual(decompress(cfb.read(f"VBA/{path.stem}")).decode("cp1252"),
                                 build_workbook.module_text(path))
                record = "Class" if path.suffix.lower() == ".cls" else "Module"
                self.assertIn(f"{record}={path.stem}", project)

    def test_production_holds_no_test_module(self) -> None:
        _, result = self.build("production")
        self.assertFalse([path for path in result.components if "tests" in path.parts])

    def test_main_reports_the_build(self) -> None:
        with tempfile.TemporaryDirectory() as temporary, \
                contextlib.redirect_stdout(io.StringIO()) as output:
            code = build_workbook.main([str(Path(temporary) / "W.xlsm"), "--root", str(ROOT),
                                        "--configuration", "production"])
        self.assertEqual(code, 0)
        selected = build_workbook.components(ROOT, self.layout, "production")
        expected = ", ".join(path.stem for path in selected) or "no VBA component"
        self.assertIn(f"components: {expected}", output.getvalue())
        with git_tree({}) as root, tempfile.TemporaryDirectory() as temporary, \
                contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(build_workbook.main([str(Path(temporary) / "W.xlsm"), "--root",
                                                  str(root), "--configuration", "production"]), 0)
        self.assertIn("no VBA component: the workbook holds only its document modules",
                      output.getvalue())
        self.assertIn(f"layout: blank sheets of {build_workbook.CONFIG}", output.getvalue())
        with contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(build_workbook.main(["--configuration", "staging"]), 1)
        self.assertIn("unknown configuration", output.getvalue())


CLASS_SOURCE = (b'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\n'
                b'Attribute VB_Name = "CPair"\r\nOption Explicit\r\n')
SHEET_SOURCE = (b'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\n'
                b'Attribute VB_Name = "shInputs"\r\nOption Explicit\r\n'
                b'Private Sub Worksheet_Change(ByVal Target As Range)\r\nEnd Sub\r\n')


@contextlib.contextmanager
def git_tree(files: dict[str, bytes]) -> Any:
    """A committed Git repository with the build configuration and the given files."""
    with source_tree(files) as root:
        gitfixture.init(root)
        gitfixture.commit(root)
        yield root


# The synthetic build configuration, whose sheets the project and template builds use.
SYNTHETIC_CONFIG = {build_workbook.CONFIG: workbook_template.build_configuration_bytes()}


class ProjectSource(unittest.TestCase):
    """Builds of a project with its own standard, class and document modules."""

    files = {**SYNTHETIC_CONFIG,
             "src/core/CORE_Demo.bas": b'Attribute VB_Name = "CORE_Demo"\r\nOption Explicit\r\n',
             "src/classes/CPair.cls": CLASS_SOURCE,
             "src/workbook/shInputs.cls": SHEET_SOURCE}

    def test_every_kind_of_module_is_embedded(self) -> None:
        with git_tree(self.files) as root, tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary) / "W.xlsm"
            result = build_workbook.build(root, "production", target)
            with zipfile.ZipFile(target) as package:
                cfb = CompoundFile(package.read("xl/vbaProject.bin"))
        self.assertTrue(result.clean)
        self.assertEqual([path.name for path in result.components], ["CORE_Demo.bas", "CPair.cls"])
        project = cfb.read("PROJECT").decode("cp1252")
        for record in ("Module=CORE_Demo", "Class=CPair", "Document=shInputs/&H00000000"):
            self.assertIn(record, project)
        sheet = decompress(cfb.read("VBA/shInputs")).decode("cp1252")
        self.assertIn("Private Sub Worksheet_Change", sheet)
        self.assertTrue(sheet.startswith('Attribute VB_Name = "shInputs"'))

    def test_a_test_sheet_module_is_left_out_of_production_only(self) -> None:
        config = workbook_template.build_configuration()
        config["sheets"].append({"tab": "Checks", "code_name": "shChecks", "role": "Test"})
        checks = (b'Attribute VB_Name = "shChecks"\r\nOption Explicit\r\n'
                  b"Private Sub Worksheet_Activate()\r\nEnd Sub\r\n")
        files = {build_workbook.CONFIG: json.dumps(config).encode(),
                 "src/workbook/shChecks.cls": checks}
        projects = {}
        with git_tree(files) as root, tempfile.TemporaryDirectory() as temporary:
            for configuration in ("production", "development"):
                target = Path(temporary) / f"{configuration}.xlsm"
                build_workbook.build(root, configuration, target)
                with zipfile.ZipFile(target) as package:
                    projects[configuration] = CompoundFile(package.read("xl/vbaProject.bin")) \
                        .read("PROJECT").decode("cp1252")
        self.assertNotIn("shChecks", projects["production"])
        self.assertIn("Document=shChecks/&H00000000", projects["development"])

    def test_a_document_module_in_the_wrong_case_stops_the_build(self) -> None:
        files = {**SYNTHETIC_CONFIG,
                 "src/workbook/shINPUTS.cls": SHEET_SOURCE.replace(b"shInputs", b"shINPUTS")}
        with git_tree(files) as root, tempfile.TemporaryDirectory() as temporary, \
                self.assertRaisesRegex(build_workbook.BuildError,
                                       r"shINPUTS\.cls \(code name shInputs\) differs"):
            build_workbook.build(root, "production", Path(temporary) / "W.xlsm")


DEMO_OVERRIDE = (b'Attribute VB_Name = "CORE_Demo"\r\nOption Explicit\r\n'
                 b"Private Const MUTANT As Long = 1\r\n")


class OverrideBuild(unittest.TestCase):
    """Builds with --override: a folder's modules replace those of the same VB_Name."""

    files = ProjectSource.files

    def build(self, root: Path, overrides: dict[str, bytes], configuration: str = "production"
              ) -> tuple[build_workbook.Build, Path]:
        """A build of ``root`` with the files ``overrides`` in test-results/mutants/."""
        folder = root / "test-results/mutants"
        shutil.rmtree(folder, ignore_errors=True)
        for name, data in overrides.items():
            (folder / name).parent.mkdir(parents=True, exist_ok=True)
            (folder / name).write_bytes(data)
        target = root / "test-results/W.xlsm"
        target.unlink(missing_ok=True)
        return build_workbook.build(root, configuration, target, override=folder), target

    def test_each_file_replaces_the_module_its_vb_name_names(self) -> None:
        sheet = SHEET_SOURCE.replace(b"End Sub", b"    Debug.Print 1\r\nEnd Sub")
        overrides = {"src/core/CORE_Demo.bas": DEMO_OVERRIDE,
                     "classes/CPair.cls": CLASS_SOURCE + b"Private mValue As Long\r\n",
                     "shInputs.cls": sheet, "report.txt": b"not a module"}
        with git_tree(self.files) as root:
            plain = root / "test-results/plain.xlsm"
            build_workbook.build(root, "production", plain)
            result, target = self.build(root, overrides)
            with zipfile.ZipFile(target) as package:
                cfb = CompoundFile(package.read("xl/vbaProject.bin"))
            with zipfile.ZipFile(plain) as package:
                before = CompoundFile(package.read("xl/vbaProject.bin"))
            stamp = check_excel_run.read_stamp(target, "build_commit")
            folder = root / "test-results/mutants"
        self.assertTrue(result.clean)
        self.assertEqual([(item.name, item.path.relative_to(folder).as_posix(), item.sha256)
                          for item in result.overrides],
                         [("CPair", "classes/CPair.cls",
                           hashlib.sha256(overrides["classes/CPair.cls"]).hexdigest()),
                          ("shInputs", "shInputs.cls", hashlib.sha256(sheet).hexdigest()),
                          ("CORE_Demo", "src/core/CORE_Demo.bas",
                           hashlib.sha256(DEMO_OVERRIDE).hexdigest())])
        # The same modules, of the same kinds, with the replaced ones' new source.
        self.assertEqual(cfb.read("PROJECT"), before.read("PROJECT"))
        self.assertIn("Private Const MUTANT", decompress(cfb.read("VBA/CORE_Demo")).decode())
        self.assertIn("Private mValue", decompress(cfb.read("VBA/CPair")).decode())
        self.assertNotIn("VERSION 1.0 CLASS", decompress(cfb.read("VBA/CPair")).decode())
        self.assertIn("Debug.Print 1", decompress(cfb.read("VBA/shInputs")).decode())
        self.assertEqual(decompress(cfb.read("VBA/ThisWorkbook")),
                         decompress(before.read("VBA/ThisWorkbook")))
        self.assertEqual(stamp, result.commit + " (built with overrides from test-results/"
                         "mutants: " + ", ".join(f"{item.name} SHA-256 {item.sha256}"
                                                 for item in result.overrides)
                         + ": NOT evidence-grade)")
        self.assertIn(build_workbook.OVERRIDDEN, stamp or "")

    def test_a_file_that_replaces_no_module_exactly_stops_the_build(self) -> None:
        cases = {
            "matches no module of the production configuration": {
                "CORE_Other.bas": DEMO_OVERRIDE.replace(b"CORE_Demo", b"CORE_Other")},
            # Test modules are not built in production, so they cannot be replaced there.
            "TEST_Demo matches no module": {
                "TEST_Demo.bas": DEMO_OVERRIDE.replace(b"CORE_Demo", b"TEST_Demo")},
            "both name the module core_demo": {
                "a/CORE_Demo.bas": DEMO_OVERRIDE,
                "b/core_demo.bas": DEMO_OVERRIDE.replace(b"CORE_Demo", b"core_demo")},
            "differs from the module CORE_Demo in case": {
                "CORE_Demo.bas": DEMO_OVERRIDE.replace(b"CORE_Demo", b"core_demo")},
            "needs exactly one Attribute VB_Name": {"CORE_Demo.bas": b"Option Explicit\r\n"},
            "needs exactly one Attribute VB_Name = ": {
                "CORE_Demo.bas": DEMO_OVERRIDE + b'Attribute VB_Name = "CORE_Demo"\r\n'},
            "forms are not built": {"CORE_Demo.bas": DEMO_OVERRIDE,
                                    "frmA.frm": b'Attribute VB_Name = "frmA"\r\n'},
            "CPair is a class module, which only a .cls file replaces": {
                "CPair.bas": b'Attribute VB_Name = "CPair"\r\nOption Explicit\r\n'},
            "CORE_Demo is a standard module, which only a .bas file replaces": {
                "CORE_Demo.cls": DEMO_OVERRIDE},
            "non-ASCII byte 0xE9": {"CORE_Demo.bas": DEMO_OVERRIDE + b"' caf\xe9\r\n"},
            "byte-order mark": {"CORE_Demo.bas": b"\xef\xbb\xbf" + DEMO_OVERRIDE},
            "carriage return without a line feed": {
                "CORE_Demo.bas": DEMO_OVERRIDE.replace(b"\r\nOption", b"\rOption")},
            "does not start with its Attribute VB_Name line": {
                "CORE_Demo.bas": b"' note\r\n" + DEMO_OVERRIDE},
            "holds no .bas or .cls file": {"notes.txt": b"x"},
        }
        with git_tree(self.files) as root:
            for expected, overrides in cases.items():
                with self.subTest(expected=expected):
                    with self.assertRaisesRegex(build_workbook.BuildError, expected):
                        self.build(root, overrides)
                    self.assertFalse((root / "test-results/W.xlsm").exists())
            with self.assertRaisesRegex(build_workbook.BuildError, "not a folder"):
                build_workbook.build(root, "production", root / "W.xlsm",
                                     override=root / "absent")

    def test_a_test_module_is_replaced_in_development(self) -> None:
        files = {**self.files, "tests/modules/TEST_Demo.bas":
                 b'Attribute VB_Name = "TEST_Demo"\r\nOption Explicit\r\n'}
        override = b'Attribute VB_Name = "TEST_Demo"\r\nOption Explicit\r\n\r\n'
        with git_tree(files) as root:
            result, target = self.build(root, {"TEST_Demo.bas": override}, "development")
            with zipfile.ZipFile(target) as package:
                cfb = CompoundFile(package.read("xl/vbaProject.bin"))
        self.assertEqual([item.name for item in result.overrides], ["TEST_Demo"])
        self.assertEqual(decompress(cfb.read("VBA/TEST_Demo")), override)

    def test_without_the_option_the_build_is_unchanged(self) -> None:
        with git_tree(self.files) as root:
            first, second = root / "test-results/first.xlsm", root / "test-results/second.xlsm"
            plain = build_workbook.build(root, "production", first)
            explicit = build_workbook.build(root, "production", second, override=None)
            self.assertEqual(first.read_bytes(), second.read_bytes())
            self.assertEqual(check_excel_run.read_stamp(first, "build_commit"), plain.commit)
        self.assertEqual((plain.overrides, plain.override_folder), ((), ""))
        self.assertEqual(plain, explicit)

    def test_main_prints_each_replacement_with_its_sha256(self) -> None:
        with git_tree(self.files) as root, contextlib.redirect_stdout(io.StringIO()) as output:
            folder = root / "test-results/mutants"
            folder.mkdir(parents=True)
            (folder / "CORE_Demo.bas").write_bytes(DEMO_OVERRIDE)
            code = build_workbook.main([str(root / "test-results/W.xlsm"), "--root", str(root),
                                        "--configuration", "production", "--override",
                                        str(folder)])
            commit = build_workbook.git(root, "rev-parse", "HEAD")
        self.assertEqual(code, 0)
        text = output.getvalue()
        self.assertIn(f"overrides from test-results/mutants: NOT evidence for commit {commit}",
                      text)
        self.assertIn(f"replaced CORE_Demo with {folder / 'CORE_Demo.bas'}, SHA-256 "
                      f"{hashlib.sha256(DEMO_OVERRIDE).hexdigest()}", text)
        with git_tree(self.files) as root, contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(build_workbook.main([str(root / "W.xlsm"), "--root", str(root),
                                                  "--override", str(root / "absent")]), 1)
        self.assertIn("not a folder", output.getvalue())


class TemplateBuild(unittest.TestCase):
    """Builds that start from src/workbook/Template.xlsx."""

    # The parts that a build from the template rewrites; it copies every other one.
    edited = {"[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels",
              "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"}

    def build(self, files: dict[str, bytes], configuration: str = "production"
              ) -> tuple[bytes, bytes, build_workbook.Build]:
        """Two builds of one commit, and what the first holds."""
        with git_tree({**SYNTHETIC_CONFIG, build_workbook.TEMPLATE: workbook_template.template(),
                       **files}) as root, \
                tempfile.TemporaryDirectory() as temporary:
            first, second = Path(temporary) / "first.xlsm", Path(temporary) / "second.xlsm"
            result = build_workbook.build(root, configuration, first)
            build_workbook.build(root, configuration, second)
            return first.read_bytes(), second.read_bytes(), result

    def test_the_fixture_is_a_clean_committable_template(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "Template.xlsx"
            path.write_bytes(workbook_template.template())
            self.assertEqual(check_binaries.package_problems(path), [])

    def test_the_build_keeps_the_layout_and_is_byte_identical(self) -> None:
        first, second, result = self.build({})
        self.assertTrue(result.template)
        self.assertEqual(first, second)
        with zipfile.ZipFile(io.BytesIO(workbook_template.template())) as template, \
                zipfile.ZipFile(io.BytesIO(first)) as built:
            self.assertEqual(built.namelist(), [*template.namelist(), "xl/vbaProject.bin"])
            for name in sorted(set(template.namelist()) - self.edited):
                with self.subTest(part=name):
                    self.assertEqual(built.read(name), template.read(name))
            for info in built.infolist():
                self.assertEqual(info.date_time, build_workbook.ZIP_TIME)
            types = built.read("[Content_Types].xml").decode()
            book = built.read("xl/workbook.xml").decode()
            readme = built.read("xl/worksheets/sheet1.xml").decode()
            inputs = built.read("xl/worksheets/sheet2.xml").decode()
            self.assertIn(build_workbook.MACRO_MAIN, types)
            self.assertNotIn(build_workbook.SHEET_MAIN, types)
            self.assertIn('Target="vbaProject.bin"',
                          built.read("xl/_rels/workbook.xml.rels").decode())
            CompoundFile(built.read("xl/vbaProject.bin"))
        self.assertIn('<workbookPr defaultThemeVersion="166925" codeName="ThisWorkbook"/>', book)
        self.assertIn('localSheetId="1">Inputs!$A$1:$C$5</definedName>', book)
        self.assertIn('mc:Ignorable="x14ac"><sheetPr codeName="shReadme"/>', readme)
        self.assertIn('<sheetPr codeName="shInputs"/>', inputs)
        # The stamp cell keeps its style; the label beside it is the template's own.
        self.assertIn(f'<c r="B3" s="{workbook_template.LOCKED}" t="inlineStr">', readme)
        self.assertIn('<c r="A3" s="0" t="s"><v>1</v></c>', readme)
        self.assertNotIn(build_workbook.TITLE, readme)
        self.assertIn('<sheetProtection sheet="1" objects="1" scenarios="1"/>', inputs)
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "W.xlsm"
            path.write_bytes(first)
            for role, value in (("commit", result.commit), ("time", result.time),
                                ("configuration", "production")):
                with self.subTest(stamp=role):
                    self.assertEqual(check_excel_run.read_stamp(
                        path, workbook_template.build_configuration()["stamp"][role]), value)

    def test_a_stamp_cell_takes_the_style_of_its_row_or_column(self) -> None:
        sheet = workbook_template.Sheet("Readme", {"A1": (0, 0)}, columns=((2, 2, 1),))
        xml = workbook_template.sheet_xml(sheet)
        stamped = build_workbook.set_cell(build_workbook.set_cell(xml, 3, 2, "a"), 1, 2, "b")
        self.assertIn('<c r="A1" s="0" t="s"><v>0</v></c><c r="B1" s="1" t="inlineStr">',
                      stamped)
        self.assertIn('<row r="3"><c r="B3" s="1" t="inlineStr">', stamped)
        row = xml.replace('<row r="1" spans="1:3" x14ac:dyDescent="0.25">',
                          '<row r="1" s="2" customFormat="1"/><row r="2">')
        row = row.replace('<c r="A1" s="0" t="s"><v>0</v></c></row>', "</row>")
        self.assertIn('<row r="1" s="2" customFormat="1"><c r="C1" s="2" t="inlineStr">',
                      build_workbook.set_cell(row, 1, 3, "c"))

    def test_development_appends_the_test_sheets_blank(self) -> None:
        config = workbook_template.build_configuration()
        config["sheets"].append({"tab": "Checks", "code_name": "shChecks", "role": "Test"})
        config["names"].append({"name": "Checks_Results", "sheet": "shChecks", "cell": "B3",
                                "label": "Results"})
        files = {build_workbook.CONFIG: json.dumps(config).encode()}
        first, second, _ = self.build(files, "development")
        self.assertEqual(first, second)
        with zipfile.ZipFile(io.BytesIO(first)) as built:
            book = built.read("xl/workbook.xml").decode()
            self.assertIn('<sheet name="Checks" sheetId="3" r:id="rIdBuildSheet"/>', book)
            self.assertIn("<definedName name=\"Checks_Results\">'Checks'!$B$3</definedName>", book)
            self.assertIn('Id="rIdBuildSheet" Type="http://schemas.openxmlformats.org/'
                          'officeDocument/2006/relationships/worksheet" '
                          'Target="worksheets/sheet3.xml"',
                          built.read("xl/_rels/workbook.xml.rels").decode())
            self.assertIn('PartName="/xl/worksheets/sheet3.xml"',
                          built.read("[Content_Types].xml").decode())
            checks = built.read("xl/worksheets/sheet3.xml").decode()
            self.assertIn('<sheetPr codeName="shChecks"/>', checks)
            self.assertIn("Results", checks)
        production, _, _ = self.build(files, "production")
        with zipfile.ZipFile(io.BytesIO(production)) as built:
            self.assertNotIn("Checks", built.read("xl/workbook.xml").decode())
            self.assertNotIn("xl/worksheets/sheet3.xml", built.namelist())

    def test_a_template_that_differs_from_the_configuration_stops_the_build(self) -> None:
        config = workbook_template.build_configuration()
        config["sheets"].insert(0, {"tab": "Checks", "code_name": "shChecks", "role": "Test"})
        names = dict(workbook_template.DEFAULT_NAMES)
        moved = {**names, "cfg_reporting_date": "Inputs!$C$6"}
        del names["build_time"]
        sheets = workbook_template.default_sheets()
        sheets[1].tab = "Settings"
        cases = {
            "not the sheets": workbook_template.template(sheets),
            "does not define the workbook name build_time": workbook_template.template(
                names=names),
            r"cfg_reporting_date refers to Inputs!\$C\$6": workbook_template.template(
                names=moved),
        }
        for expected, data in cases.items():
            with self.subTest(expected=expected), git_tree(
                    {**SYNTHETIC_CONFIG, build_workbook.TEMPLATE: data}) as root, \
                    tempfile.TemporaryDirectory() as temporary, \
                    self.assertRaisesRegex(build_workbook.BuildError, expected):
                build_workbook.build(root, "production", Path(temporary) / "W.xlsm")
        with git_tree({build_workbook.TEMPLATE: workbook_template.template(),
                       build_workbook.CONFIG: json.dumps(config).encode()}) as root, \
                tempfile.TemporaryDirectory() as temporary, \
                self.assertRaisesRegex(build_workbook.BuildError, "Test sheets must come last"):
            build_workbook.build(root, "development", Path(temporary) / "W.xlsm")

    def test_a_template_with_macros_stops_the_build(self) -> None:
        buffer = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(workbook_template.template())) as template, \
                zipfile.ZipFile(buffer, "w") as copy:
            for name in template.namelist():
                copy.writestr(name, template.read(name))
            copy.writestr("xl/vbaProject.bin", b"project")
        with git_tree({**SYNTHETIC_CONFIG, build_workbook.TEMPLATE: buffer.getvalue()}) as root, \
                tempfile.TemporaryDirectory() as temporary, \
                self.assertRaisesRegex(build_workbook.BuildError, "carries macros"):
            build_workbook.build(root, "production", Path(temporary) / "W.xlsm")

    def test_main_names_the_layout(self) -> None:
        with git_tree({**SYNTHETIC_CONFIG,
                       build_workbook.TEMPLATE: workbook_template.template()}) as root, \
                tempfile.TemporaryDirectory() as temporary, \
                contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(build_workbook.main([str(Path(temporary) / "W.xlsm"), "--root",
                                                  str(root), "--configuration", "production"]), 0)
        self.assertIn(f"layout: {build_workbook.TEMPLATE}", output.getvalue())


class Manifest(unittest.TestCase):
    def test_manifest_records_the_workbook_and_every_export(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            (folder / "exported").mkdir()
            (folder / "exported/A.bas").write_bytes(b"a")
            workbook = folder / "W.xlsm"
            workbook.write_bytes(b"workbook")
            path = export_manifest.write(workbook, folder)
            digest, files = check_excel_run.load_manifest(path)
            self.assertEqual(digest, export_manifest.sha256(workbook))
            self.assertEqual(set(files), {"A.bas"})
            (folder / "exported/A.bas").unlink()
            with contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(export_manifest.main([str(workbook), str(folder)]), 1)
            self.assertIn("holds no exported file", output.getvalue())

    def test_manifest_follows_the_evidence_policy(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            data = json.loads(export_manifest.POLICY.read_text(encoding="utf-8"))
            self.assertEqual(export_manifest.layout(),
                             (data["evidence"]["exported"], data["evidence"]["export_manifest"]))
            data["evidence"].update(exported="export", export_manifest="binding.json")
            policy = folder / "policy.json"
            policy.write_text(json.dumps(data), encoding="utf-8")
            (folder / "export").mkdir()
            (folder / "export/A.bas").write_bytes(b"a")
            workbook = folder / "W.xlsm"
            workbook.write_bytes(b"workbook")
            path = export_manifest.write(workbook, folder, policy)
            self.assertEqual(path, folder / "binding.json")
            self.assertEqual(set(check_excel_run.load_manifest(path)[1]), {"A.bas"})
            data["evidence"]["export_manifest"] = "../binding.json"
            policy.write_text(json.dumps(data), encoding="utf-8")
            with contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(export_manifest.main([str(workbook), str(folder), "--policy",
                                                       str(policy)]), 1)
            self.assertIn("not a plain file name", output.getvalue())


class LibreOfficeSmoke(unittest.TestCase):
    """The exit codes of lo_smoke.py that need no LibreOffice."""

    def test_an_unknown_option_or_a_missing_workbook_exits_2_before_starting_libreoffice(
            self) -> None:
        with unittest.mock.patch("subprocess.Popen") as popen, \
                contextlib.redirect_stderr(io.StringIO()) as errors:
            with self.assertRaises(SystemExit) as raised:
                lo_smoke.main(["--help-me"])
            self.assertEqual(raised.exception.code, 2)
            self.assertEqual(lo_smoke.main(["no-such-workbook.xlsm"]), 2)
        popen.assert_not_called()
        self.assertIn("no workbook at no-such-workbook.xlsm", errors.getvalue())

    def test_a_workbook_libreoffice_cannot_run_exits_1_without_a_traceback(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "Workbook.xlsm"
            workbook.write_bytes(b"not a workbook")
            with unittest.mock.patch("shutil.which", return_value="soffice"), \
                    unittest.mock.patch("importlib.util.find_spec",
                                        return_value=object()), \
                    unittest.mock.patch("subprocess.Popen") as popen, \
                    unittest.mock.patch.object(lo_smoke, "connect"), \
                    unittest.mock.patch.object(lo_smoke, "enable_vba"), \
                    unittest.mock.patch.object(lo_smoke, "certify",
                                               side_effect=RuntimeError("cannot load")), \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(lo_smoke.main([str(workbook)]), 1)
        self.assertIn("lo-smoke: FAIL (LibreOffice could not run the workbook: "
                      "RuntimeError: cannot load)", output.getvalue())
        popen.return_value.wait.assert_called_once_with(timeout=60)

    def test_a_build_failure_exits_1_without_a_traceback_before_starting_libreoffice(
            self) -> None:
        failure = build_workbook.BuildError("src/forms: forms are not built (frmInput)")
        with unittest.mock.patch("shutil.which", return_value="soffice"), \
                unittest.mock.patch("importlib.util.find_spec", return_value=object()), \
                unittest.mock.patch.object(build_workbook, "build",
                                           side_effect=failure) as build, \
                unittest.mock.patch("subprocess.Popen") as popen, \
                contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(lo_smoke.main([]), 1)
        build.assert_called_once()
        popen.assert_not_called()
        self.assertEqual(output.getvalue().strip(),
                         "lo-smoke: FAIL (build: src/forms: forms are not built (frmInput))")

    def test_the_entry_point_logs_and_results_file_come_from_the_evidence_policy(self) -> None:
        data = json.loads((ROOT / check_excel_run.POLICY).read_text(encoding="utf-8"))
        data["suite"]["source"] = "tests/modules/TEST_Renamed.bas"
        data["suite"]["certification_entry_point"] = "CertifyRenamed"
        for item in data["modes"]:
            item["log"] = f"renamed-{item['log']}"
        data["evidence"]["certification_log"] = "renamed-certification.log"
        data["evidence"]["results"] = "renamed-results.json"
        policy = check_excel_run.parse_policy(json.dumps(data))
        injected = policy.logs["INJECTED_FAILURE"]
        procedures: list[str] = []

        def certify(_context: object, workbook: Path, folder: Path,
                    procedure: str) -> dict[str, object]:
            procedures.append(procedure)
            for name in policy.logs.values():
                failure = "FAILURE=exact-double/half-plus-quarter\n" if name == injected else ""
                (folder / name).write_text(f"{failure}RESULT=DONE\n", encoding="utf-8")
            (folder / policy.certification).write_text("CERTIFIED\n", encoding="utf-8")
            (folder / policy.results).write_text('{"kind": "test-results"}', encoding="utf-8")
            return {"workbook": str(workbook), "error": None}

        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "Workbook.xlsm"
            workbook.write_bytes(b"built")
            with unittest.mock.patch.object(check_excel_run, "working_policy",
                                            return_value=policy), \
                    unittest.mock.patch("shutil.which", return_value="soffice"), \
                    unittest.mock.patch("importlib.util.find_spec",
                                        return_value=object()), \
                    unittest.mock.patch("subprocess.Popen"), \
                    unittest.mock.patch.object(lo_smoke, "connect"), \
                    unittest.mock.patch.object(lo_smoke, "enable_vba"), \
                    unittest.mock.patch.object(lo_smoke, "certify", side_effect=certify), \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                code = lo_smoke.main([str(workbook)])
        summary = json.loads(output.getvalue())
        self.assertEqual(procedures, ["TEST_Renamed.CertifyRenamed"])
        self.assertEqual(list(summary["logs"]),
                         [*policy.logs.values(), "renamed-certification.log"])
        self.assertEqual(summary["logs"]["renamed-certification.log"], ["CERTIFIED"])
        self.assertEqual(summary["results"], {"kind": "test-results"})
        # The injected mode's designed failure is found under the policy's log name.
        self.assertEqual(summary["unexpected"], [])
        self.assertEqual(code, 0)


if __name__ == "__main__":
    unittest.main()
