#!/usr/bin/env python3
"""Generate or check the macro-free workbook template src/workbook/KPR_Template.xlsx.

The template is the sheets of the K-PRICING workbook without any code: a Home
sheet and the K-PRICING Ribbon (src/ribbon/customUI14.xml, byte for byte, as
customUI/customUI14.xml with exactly one Ribbon relationship). It has no VBA
project and no document properties. tools/Check-KPR.ps1 saves it as an .xlsm,
imports one commit's modules, compiles them and builds the demo sheet into it.

The file is generated, never edited in Excel: it is a stored (uncompressed)
ZIP with fixed entry order and timestamps, so the same source always gives the
same bytes. --check (the default) fails when the committed template differs
from what this script generates; --write rewrites it.
"""

from __future__ import annotations

import argparse
import io
import sys
import zipfile
from collections.abc import Sequence
from pathlib import Path
from xml.sax.saxutils import escape

from inject_ribbon import (
    CUSTOMUI_PART,
    DEFAULT_CUSTOMUI,
    FIXED_DATE_TIME,
    RIBBON_REL_ID,
    RIBBON_REL_TYPE,
    ROOT,
)

TEMPLATE = ROOT / "src" / "workbook" / "KPR_Template.xlsx"
HOME_SHEET = "Home"
# The build writes the commit's SHA here.
COMMIT_CELL = "B13"
REPOSITORY_URL = "https://github.com/danielep71/K-PRICING"

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_DOCUMENT_TYPE = DOC_REL_NS + "/officeDocument"
DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

# Style indexes into cellXfs below.
NORMAL, TITLE, HEADING, NOTE = 0, 1, 2, 3

# (row, column A text, style); column B of the commit row is filled by the build.
HOME_ROWS: tuple[tuple[int, str, int], ...] = (
    (1, "K-PRICING", TITLE),
    (2, "Date primitives for Excel: the KPR_Dates_* worksheet and VBA functions.", NORMAL),
    (4, "Start here", HEADING),
    (5, "The K-PRICING tab on the Ribbon inserts any function and builds the demo sheets.", NORMAL),
    (6, "K-PRICING > Register function help lists the functions under KPR Dates in Insert Function"
        " (Shift+F3).", NORMAL),
    (7, "The KPR Dates Demo sheet compares each function with Excel's own formulas; change its"
        " orange inputs.", NORMAL),
    (9, "About this file", HEADING),
    (10, "Built from the K-PRICING repository: this template plus the VBA source of one commit.", NOTE),
    (11, "Keep it as a macro-enabled workbook (.xlsm). The functions work only while it is open.", NOTE),
    (12, "Source: " + REPOSITORY_URL, NOTE),
    (13, "Commit:", NOTE),
)


def _xml(body: str) -> bytes:
    return (DECLARATION + body).encode("utf-8")


def content_types() -> bytes:
    return _xml(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>")


def package_relationships() -> bytes:
    return _xml(
        f'<Relationships xmlns="{PACKAGE_REL_NS}">'
        f'<Relationship Id="rId1" Type="{OFFICE_DOCUMENT_TYPE}" Target="xl/workbook.xml"/>'
        f'<Relationship Id="{RIBBON_REL_ID}" Type="{RIBBON_REL_TYPE}" Target="{CUSTOMUI_PART}"/>'
        "</Relationships>")


def workbook() -> bytes:
    return _xml(
        f'<workbook xmlns="{MAIN_NS}" xmlns:r="{DOC_REL_NS}">'
        '<bookViews><workbookView activeTab="0"/></bookViews>'
        f'<sheets><sheet name="{HOME_SHEET}" sheetId="1" r:id="rId1"/></sheets>'
        '<calcPr calcId="191029"/>'
        "</workbook>")


def workbook_relationships() -> bytes:
    return _xml(
        f'<Relationships xmlns="{PACKAGE_REL_NS}">'
        f'<Relationship Id="rId1" Type="{DOC_REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{DOC_REL_NS}/styles" Target="styles.xml"/>'
        "</Relationships>")


def styles() -> bytes:
    return _xml(
        f'<styleSheet xmlns="{MAIN_NS}">'
        '<fonts count="4">'
        '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
        '<font><b/><sz val="20"/><color rgb="FF1F4E79"/><name val="Calibri"/><family val="2"/></font>'
        '<font><b/><sz val="13"/><color rgb="FF1F4E79"/><name val="Calibri"/><family val="2"/></font>'
        '<font><i/><sz val="10"/><color rgb="FF595959"/><name val="Calibri"/><family val="2"/></font>'
        "</fonts>"
        '<fills count="2"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="4">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        "</cellXfs>"
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>")


def home_sheet() -> bytes:
    rows = "".join(
        f'<row r="{row}"><c r="A{row}" s="{style}" t="inlineStr"><is><t>{escape(text)}</t></is></c></row>'
        for row, text, style in HOME_ROWS)
    return _xml(
        f'<worksheet xmlns="{MAIN_NS}" xmlns:r="{DOC_REL_NS}">'
        '<sheetViews><sheetView showGridLines="0" tabSelected="1" workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols><col min="1" max="1" width="12.7109375" customWidth="1"/>'
        '<col min="2" max="2" width="48.7109375" customWidth="1"/></cols>'
        f"<sheetData>{rows}</sheetData>"
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        "</worksheet>")


def entries(customui: bytes) -> list[tuple[str, bytes]]:
    """Every part of the template, in the order the ZIP stores them."""
    return [
        ("[Content_Types].xml", content_types()),
        ("_rels/.rels", package_relationships()),
        ("xl/workbook.xml", workbook()),
        ("xl/_rels/workbook.xml.rels", workbook_relationships()),
        ("xl/styles.xml", styles()),
        ("xl/worksheets/sheet1.xml", home_sheet()),
        (CUSTOMUI_PART, customui),
    ]


def build(customui: bytes) -> bytes:
    """The template's bytes: the same customui always gives the same bytes."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_STORED) as archive:
        for name, data in entries(customui):
            info = zipfile.ZipInfo(name, FIXED_DATE_TIME)
            info.compress_type = zipfile.ZIP_STORED
            info.create_system = 0
            info.external_attr = 0
            archive.writestr(info, data)
    return buffer.getvalue()


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--check", action="store_true", help="fail when the template is stale (default)")
    mode.add_argument("--write", action="store_true", help="rewrite the template from its source")
    parser.add_argument("--template", type=Path, default=TEMPLATE, help="template path")
    parser.add_argument("--customui", type=Path, default=DEFAULT_CUSTOMUI, help="Ribbon source")
    args = parser.parse_args(argv)
    try:
        expected = build(args.customui.read_bytes())
    except OSError as error:
        print(f"ERROR: cannot read the Ribbon source: {error}", file=sys.stderr)
        return 2
    if args.write:
        args.template.parent.mkdir(parents=True, exist_ok=True)
        args.template.write_bytes(expected)
        print(f"Wrote {args.template} ({len(expected)} bytes)")
        return 0
    try:
        actual = args.template.read_bytes()
    except OSError as error:
        print(f"FAIL: {args.template} is missing ({error}); run tools/workbook_template.py --write",
              file=sys.stderr)
        return 1
    if actual != expected:
        print(f"FAIL: {args.template} differs from its source; run tools/workbook_template.py --write",
              file=sys.stderr)
        return 1
    print(f"PASS: {args.template.name} matches its source ({len(expected)} bytes, Ribbon included)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
