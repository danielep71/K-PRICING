#!/usr/bin/env python3
"""Tests for tools/workbook_template.py: the template is current, deterministic and macro-free."""

from __future__ import annotations

import contextlib
import io
import tempfile
import unittest
import zipfile
from pathlib import Path

import workbook_template as wt
from inject_ribbon import (
    CUSTOMUI_PART,
    DEFAULT_CUSTOMUI,
    FIXED_DATE_TIME,
    RIBBON_REL_TYPE,
    parse_xml,
)

REL = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
SHEET = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"


def run(*argv: str) -> int:
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        return wt.main(list(argv))


class WorkbookTemplateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.customui = DEFAULT_CUSTOMUI.read_bytes()
        self.archive = zipfile.ZipFile(io.BytesIO(wt.build(self.customui)))
        self.addCleanup(self.archive.close)

    def test_committed_template_is_current(self) -> None:
        self.assertEqual(run("--check"), 0)
        self.assertEqual(wt.TEMPLATE.read_bytes(), wt.build(self.customui))

    def test_build_is_deterministic_and_stored(self) -> None:
        self.assertEqual(wt.build(self.customui), wt.build(self.customui))
        for info in self.archive.infolist():
            self.assertEqual(info.compress_type, zipfile.ZIP_STORED, info.filename)
            self.assertEqual(info.date_time, FIXED_DATE_TIME, info.filename)

    def test_every_part_is_well_formed_xml(self) -> None:
        for name in self.archive.namelist():
            parse_xml(name, self.archive.read(name))

    def test_macro_free_without_document_properties(self) -> None:
        names = self.archive.namelist()
        self.assertFalse([name for name in names if "vba" in name.casefold()])
        self.assertFalse([name for name in names if name.startswith("docProps/")])
        self.assertNotIn(b"macroEnabled", self.archive.read("[Content_Types].xml"))

    def test_ribbon_is_the_source_with_one_relationship(self) -> None:
        self.assertEqual(self.archive.read(CUSTOMUI_PART), self.customui)
        relationships = parse_xml("rels", self.archive.read("_rels/.rels")).iter(REL)
        ribbon = [rel for rel in relationships if rel.get("Type") == RIBBON_REL_TYPE]
        self.assertEqual([rel.get("Target") for rel in ribbon], [CUSTOMUI_PART])
        self.assertIn(b'Extension="xml"', self.archive.read("[Content_Types].xml"))

    def test_home_sheet_and_commit_cell(self) -> None:
        sheets = parse_xml("workbook", self.archive.read("xl/workbook.xml")).iter(SHEET)
        self.assertEqual([sheet.get("name") for sheet in sheets], [wt.HOME_SHEET])
        home = self.archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertIn("K-PRICING", home)
        self.assertIn(wt.REPOSITORY_URL, home)
        row = int(wt.COMMIT_CELL[1:])
        self.assertIn((row, "Commit:", wt.NOTE), wt.HOME_ROWS)
        self.assertNotIn(f'r="{wt.COMMIT_CELL}"', home, "the build fills the commit cell")

    def test_text_is_escaped(self) -> None:
        rows = wt.HOME_ROWS
        try:
            wt.HOME_ROWS = ((1, "a < b & c", wt.NORMAL),)
            parse_xml("home", wt.home_sheet())
            self.assertIn(b"a &lt; b &amp; c", wt.home_sheet())
        finally:
            wt.HOME_ROWS = rows

    def test_stale_or_missing_template_fails_and_write_repairs(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            template = Path(temp) / "KPR_Template.xlsx"
            self.assertEqual(run("--template", str(template)), 1)
            template.write_bytes(b"stale")
            self.assertEqual(run("--check", "--template", str(template)), 1)
            self.assertEqual(run("--write", "--template", str(template)), 0)
            self.assertEqual(run("--check", "--template", str(template)), 0)
            customui = Path(temp) / "customUI14.xml"
            customui.write_bytes(self.customui.replace(b"K-PRICING", b"K-PRICING 2"))
            self.assertEqual(run("--template", str(template), "--customui", str(customui)), 1)

    def test_missing_ribbon_source_is_an_error(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            self.assertEqual(run("--customui", str(Path(temp) / "missing.xml")), 2)


if __name__ == "__main__":
    unittest.main()
