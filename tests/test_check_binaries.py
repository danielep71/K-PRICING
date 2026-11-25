"""Tests for tools/check_binaries.py, in temporary Git repositories."""

from __future__ import annotations

import contextlib
import hashlib
import io
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_binaries  # noqa: E402

PNG = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR"
CONTENT_TYPES = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/'
                 'package/2006/content-types"><Default Extension="xml" '
                 'ContentType="application/xml"/>{extra}</Types>')
CORE = ('<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/'
        'metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">'
        '<dc:creator>{creator}</dc:creator><cp:lastModifiedBy>{editor}</cp:lastModifiedBy>'
        '</cp:coreProperties>')
APP = '<Properties><Company>{company}</Company><Manager>{manager}</Manager></Properties>'


def package(extra_types: str = "", parts: dict[str, str | bytes] | None = None) -> bytes:
    """A minimal Office Open XML package with the given extra content types and parts."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES.format(extra=extra_types))
        archive.writestr("xl/workbook.xml", "<workbook/>")
        for name, text in (parts or {}).items():
            archive.writestr(name, text)
    return buffer.getvalue()


CLEAN = package(parts={"docProps/core.xml": CORE.format(creator="", editor=""),
                       "docProps/app.xml": APP.format(company="", manager="")})


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class BinaryCheckTest(unittest.TestCase):
    def repository(self, files: dict[str, bytes], allowlist: object = None) -> list[str]:
        """Stage files in a new repository with this repository's .gitattributes; check it."""
        with tempfile.TemporaryDirectory(prefix="binaries-") as directory:
            root = Path(directory)
            subprocess.run(["git", "init", "-q"], cwd=root, check=True)
            files = {".gitattributes": (ROOT / ".gitattributes").read_bytes(), **files}
            if allowlist is not None:
                files[".github/binary-allowlist.json"] = json.dumps(allowlist).encode()
            for name, content in files.items():
                (root / name).parent.mkdir(parents=True, exist_ok=True)
                (root / name).write_bytes(content)
            subprocess.run(["git", "add", "-f", "."], cwd=root, check=True)
            return check_binaries.check(root)[0]

    def assertFinding(self, findings: list[str], text: str) -> None:
        self.assertTrue(any(text in finding for finding in findings), findings)

    def test_repository_passes(self) -> None:
        self.assertEqual(check_binaries.check(ROOT)[0], [])

    def test_text_files_pass(self) -> None:
        self.assertEqual(self.repository({"README.md": b"# Title\n", "data.csv": b"a,b\n1,2\n"}),
                         [])

    def test_binary_types_fail(self) -> None:
        for name in ("sample.xlsm", "book.xlsx", "docs/report.pdf", "archive.zip", "tool.exe",
                     "docs/assets/logo.png", "src/forms/frmRun.frx"):
            with self.subTest(name=name):
                self.assertFinding(self.repository({name: b"PK\x03\x04\0binary"}),
                                   f"{name}: binary file outside the approved paths")

    def test_approved_paths_pass(self) -> None:
        self.assertEqual(self.repository({
            "docs/methodology/sources/SRC-01_rule.pdf": b"%PDF-1.7\0",
            "docs/methodology/sources/templates/SRC-02_template.xlsx": b"PK\x03\x04\0",
            "src/workbook/Template.xlsx": b"PK\x03\x04\0",
            "src/forms/frmRun.frm": b"VERSION 5.00\r\n",
            "src/forms/frmRun.frx": b"\0\1\2",
            "tests/fixtures/a-case/crlf/positions.csv": b"id\r\nP1\r\n",
            "tests/fixtures/a-case/utf-16/positions.csv": "id\n".encode("utf-16"),
        }), [])

    def test_file_case_manifest_is_not_a_variant_file(self) -> None:
        self.assertFinding(self.repository({"tests/fixtures/a-case/case.json": b"{\0}"}),
                           "tests/fixtures/a-case/case.json: binary content")
        self.assertFinding(self.repository({"tests/fixtures/data.bin": b"\0"}),
                           "tests/fixtures/data.bin: binary content")

    def test_nul_byte_in_a_text_file_fails(self) -> None:
        self.assertFinding(self.repository({"notes.txt": b"text\0more"}),
                           "notes.txt: binary content in a file treated as text")
        self.assertFinding(self.repository({"payload.bin": b"\0\1\2"}),
                           "payload.bin: binary content")

    def test_nul_byte_late_in_a_text_file_fails(self) -> None:
        # Git reads only the first 8000 bytes, so it treats this file as text.
        late = b"x" * 9000 + b"\0" + b"y" * 100
        self.assertFinding(self.repository({"late.txt": late}), "late.txt")
        tail = b"x" * (1 << 16) + b"\0"  # the NUL in the second chunk read
        self.assertFinding(self.repository({"tail.txt": tail}), "tail.txt")

    def test_a_clean_office_package_passes(self) -> None:
        self.assertEqual(self.repository({"src/workbook/Template.xlsx": CLEAN}), [])

    def test_an_office_package_with_macros_fails(self) -> None:
        cases = {
            "vba part": package(parts={"xl/vbaProject.bin": "\0"}),
            "vba content type": package('<Default Extension="bin" ContentType='
                                        '"application/vnd.ms-office.vbaProject"/>'),
            "vba relationship": package(parts={"xl/_rels/workbook.xml.rels":
                                               '<Relationships><Relationship '
                                               'Target="vbaProject.bin"/></Relationships>'}),
            "macro-enabled main part": package('<Override PartName="/xl/workbook.xml" '
                                               'ContentType="application/vnd.ms-excel.'
                                               'sheet.macroEnabled.main+xml"/>'),
        }
        for label, content in cases.items():
            with self.subTest(label=label):
                self.assertFinding(self.repository({"src/workbook/Template.xlsx": content}),
                                   "carries macros")

    def test_an_office_package_with_personal_metadata_fails(self) -> None:
        cases = {
            "creator": ("docProps/core.xml", CORE.format(creator="A. Person", editor="")),
            "lastModifiedBy": ("docProps/core.xml", CORE.format(creator="", editor="A. Person")),
            "Company": ("docProps/app.xml", APP.format(company="Firm", manager="")),
            "Manager": ("docProps/app.xml", APP.format(company="", manager="A. Boss")),
        }
        for field, (part, text) in cases.items():
            with self.subTest(field=field):
                found = self.repository({"tests/fixtures/a-case/variant/input.xlsx":
                                         package(parts={part: text})})
                self.assertFinding(found, f"input.xlsx: {part}: person or organization "
                                          f"({field} is not blank); clear it")

    def test_an_office_package_with_private_data_fails(self) -> None:
        # One case of each kind; tests/test_check_workbook_privacy.py covers every form.
        absolute = ('<workbook xmlns:x15ac="http://schemas.microsoft.com/office/'
                    'spreadsheetml/2010/11/ac"><x15ac:absPath url="C:\\Users\\avery\\'
                    'OneDrive - Contoso\\"/></workbook>')
        cases: dict[str, dict[str, str | bytes]] = {
            "path or URL": {"xl/workbook.xml": absolute},
            "person or organization": {"xl/persons/person.xml":
                                       '<personList><person displayName="Avery"/></personList>'},
            "sensitivity label": {"docMetadata/LabelInfo.xml": "<labelList/>"},
            "SharePoint metadata": {"customXml/item1.xml":
                                    '<p:properties xmlns:p="http://schemas.microsoft.com/office/'
                                    '2006/metadata/properties"/>'},
        }
        for kind, parts in cases.items():
            with self.subTest(kind=kind):
                found = self.repository({"src/workbook/Template.xlsx": package(parts=parts)})
                self.assertFinding(found, f"src/workbook/Template.xlsx: {next(iter(parts))}: "
                                          f"{kind} (")
                self.assertFalse([finding for finding in found
                                  if "avery" in finding.lower() or "contoso" in finding.lower()],
                                 found)

    def test_encoding_and_character_references_hide_nothing(self) -> None:
        def utf16(text: str) -> bytes:
            return ('<?xml version="1.0" encoding="UTF-16"?>' + text).encode("utf-16")

        cases = {
            "utf-16 creator": (package(parts={"docProps/core.xml": utf16(
                CORE.format(creator="A. Person", editor=""))}),
                "docProps/core.xml: person or organization (creator is not blank)"),
            "utf-16 relationship": (package(parts={"xl/_rels/workbook.xml.rels": utf16(
                '<Relationships><Relationship Target="vbaProject.bin"/></Relationships>')}),
                "carries macros"),
            "character reference": (package('<Default Extension="bin" ContentType='
                                            '"application/vnd.ms-office.&#118;baProject"/>'),
                                    "carries macros"),
        }
        for label, (content, text) in cases.items():
            with self.subTest(label=label):
                self.assertFinding(self.repository({"src/workbook/Template.xlsx": content}),
                                   text)
        clean = package(parts={"docProps/core.xml": utf16(CORE.format(creator="", editor=""))})
        self.assertEqual(self.repository({"src/workbook/Template.xlsx": clean}), [])

    def test_a_part_that_is_not_xml_fails(self) -> None:
        broken = package(parts={"docProps/app.xml": "<Properties><Company>"})
        self.assertFinding(self.repository({"src/workbook/Template.xlsx": broken}),
                           "docProps/app.xml: unreadable (is not well-formed XML, so it "
                           "cannot be checked)")

    def test_regulatory_source_copies_are_not_inspected(self) -> None:
        published = package(parts={"xl/vbaProject.bin": "\0", "docProps/core.xml":
                                   CORE.format(creator="Regulator", editor="Staff"),
                                   "docMetadata/LabelInfo.xml": "<labelList/>"})
        self.assertEqual(self.repository({"docs/methodology/sources/SRC-01_template.xlsm":
                                          published}), [])

    def test_allowlist(self) -> None:
        digest = hashlib.sha256(PNG).hexdigest()
        entry = {"sha256": digest, "reason": "Social preview image"}
        listed = {"schema_version": 1, "files": {"docs/assets/preview.png": entry}}
        self.assertEqual(self.repository({"docs/assets/preview.png": PNG}, listed), [])
        self.assertFinding(self.repository({"docs/assets/preview.png": PNG + b"x"}, listed),
                           "content changed since it was allowlisted")
        self.assertFinding(self.repository({}, listed),
                           "docs/assets/preview.png is not a tracked file")

    def test_allowlisted_file_replaced_by_text_still_fails(self) -> None:
        # Not marked binary in .gitattributes, so only its NUL bytes made it binary.
        original = b"\0\1\2payload"
        entry = {"sha256": hashlib.sha256(original).hexdigest(), "reason": "Reviewed payload"}
        listed = {"schema_version": 1, "files": {"payload.bin": entry}}
        self.assertEqual(self.repository({"payload.bin": original}, listed), [])
        self.assertFinding(self.repository({"payload.bin": b"plain text now\n"}, listed),
                           "payload.bin: content changed since it was allowlisted")

    def test_malformed_allowlist(self) -> None:
        for allowlist, text in (([], "needs schema_version 1"),
                                ({"schema_version": 2, "files": {}}, "needs schema_version 1"),
                                ({"schema_version": 1, "files": {"a.png": {"sha256": "x",
                                                                           "reason": "r"}}},
                                 "a.png needs a 64-character sha256 and a reason")):
            with self.subTest(allowlist=allowlist):
                self.assertFinding(self.repository({}, allowlist), text)

    def test_main_reports_the_verdict(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()) as printed:
            self.assertEqual(check_binaries.main(["--root", str(ROOT)]), 0)
        self.assertIn("binaries: PASS", printed.getvalue())


if __name__ == "__main__":
    unittest.main()
