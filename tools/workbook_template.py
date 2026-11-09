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
DEMO_SHEET = "KPR Dates Demo"
# The build writes the commit's SHA here.
COMMIT_CELL = "D22"
REPOSITORY_URL = "https://github.com/danielep71/K-PRICING"

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_DOCUMENT_TYPE = DOC_REL_NS + "/officeDocument"
DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

# The Home sheet follows the K-SACCR workbook template: Aptos; column A is the
# black navigation column, B a gutter, C the labels and D the text; row 1 is
# the black band with the gold wordmark, row 2 the blue band with the sheet
# title, row 3 the grey subtitle; rows 1 to 3 and column A are frozen.
# Style indexes into cellXfs below.
(NORMAL, NAV, WORDMARK, BLACK_BAND, SHEET_TITLE, BLUE_BAND, SUBTITLE, SECTION, LABEL, TEXT,
 LINK, NAV_CURRENT, NAV_BUTTON) = range(13)

# (row, height in points, ((column, text, style, internal link or ""), ...))
HomeCell = tuple[str, str, int, str]
HOME_ROWS: tuple[tuple[int, float, tuple[HomeCell, ...]], ...] = (
    (1, 24, (("B", "K-PRICING", WORDMARK, ""), ("C", "", BLACK_BAND, ""), ("D", "", BLACK_BAND, ""))),
    (2, 26.1, (("A", "Home", NAV_CURRENT, ""), ("B", "Home", SHEET_TITLE, ""),
               ("C", "", BLUE_BAND, ""), ("D", "", BLUE_BAND, ""))),
    (3, 33.95, (("A", DEMO_SHEET, NAV_BUTTON, f"'{DEMO_SHEET}'!A1"),
                ("C", "Start here: what K-PRICING provides, how to use it and a link to every sheet."
                      " The commit this workbook was built from is at the bottom.", SUBTITLE, ""))),
    (5, 30, (("C", "What K-PRICING does", SECTION, ""),)),
    (6, 46, (("C", "Provides", LABEL, ""),
             ("D", "The KPR_Dates_* worksheet and VBA functions: Gregorian date primitives for pricing,"
                   " with scalar and dynamic-array inputs, strict parsing, explicit errors and the 1900"
                   " date system.", TEXT, ""))),
    (7, 46, (("C", "Does not cover", LABEL, ""),
             ("D", "Business calendars and holidays (v0.0.5), business-day arithmetic and roll conventions"
                   " (v0.0.6), schedules, day counts and accruals, settlement, curves and pricing.", TEXT, ""))),
    (8, 30, (("C", "Support", LABEL, ""),
             ("D", "Questions and problems: " + REPOSITORY_URL + "/issues", TEXT, ""))),
    (10, 30, (("C", "How to use", SECTION, ""),)),
    (11, 30, (("C", "1. Enable macros", LABEL, ""),
              ("D", "Open this .xlsm with macros enabled: the functions, the K-PRICING tab and the demo work"
                    " only while it is open.", TEXT, ""))),
    (12, 30, (("C", "2. Insert a function", LABEL, ""),
              ("D", "K-PRICING tab, Functions: inserts a function into the active cell and opens Excel's"
                    " Function Arguments dialog for it.", TEXT, ""))),
    (13, 30, (("C", "3. Register function help", LABEL, ""),
              ("D", "K-PRICING tab, Register function help: lists the functions under KPR Dates in Insert"
                    " Function (Shift+F3), with argument help.", TEXT, ""))),
    (14, 30, (("C", "4. Try the demo", LABEL, ""),
              ("D", "KPR Dates Demo compares each function with Excel's own formulas; change its orange"
                    " inputs. K-PRICING tab, Demo sheets builds it again.", TEXT, ""))),
    (16, 30, (("C", "Sheets", SECTION, ""),)),
    (17, 30, (("C", HOME_SHEET, LABEL, ""), ("D", "Start here. This sheet.", TEXT, ""))),
    (18, 30, (("C", DEMO_SHEET, LINK, f"'{DEMO_SHEET}'!A1"),
              ("D", "Demo. Each function next to Excel's own formula, with a status per row.", TEXT, ""))),
    (20, 30, (("C", "About this workbook", SECTION, ""),)),
    (21, 30, (("C", "Source", LABEL, ""), ("D", REPOSITORY_URL, TEXT, ""))),
    (22, 30, (("C", "Commit", LABEL, ""),)),
    (23, 30, (("C", "Built from", LABEL, ""),
              ("D", "src/workbook/KPR_Template.xlsx and the VBA source of that commit, by"
                    " tools/Check-KPR.ps1.", TEXT, ""))),
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
    def font(attributes: str, size: str, rgb: str) -> str:
        return f'<font>{attributes}<sz val="{size}"/><color rgb="FF{rgb}"/><name val="Aptos"/><family val="2"/></font>'

    fonts = (
        font("", "11", "000000"),            # 0 body
        font("", "14", "FFC000"),            # 1 wordmark
        font("<b/>", "18", "FFFFFF"),        # 2 sheet title
        font("<i/>", "10", "595959"),        # 3 subtitle
        font("<b/>", "14", "1F3864"),        # 4 section
        font("<b/>", "10", "000000"),        # 5 label
        font("", "9", "000000"),             # 6 text
        font("<u/>", "9", "0563C1"),         # 7 link
        font("<b/>", "10", "FFFFFF"),        # 8 navigation
    )
    fills = ('<fill><patternFill patternType="none"/></fill>'
             '<fill><patternFill patternType="gray125"/></fill>'
             + "".join(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{rgb}"/>'
                       '<bgColor indexed="64"/></patternFill></fill>'
                       for rgb in ("000000", "265A99", "262626")))
    left = '<alignment horizontal="left" vertical="center" indent="1"/>'
    centre = '<alignment horizontal="center" vertical="center" wrapText="1"/>'
    middle = '<alignment vertical="center"/>'
    wrap = '<alignment vertical="center" wrapText="1"/>'
    # (font, fill, alignment) per style index, in the order of the constants above
    xfs = ((0, 0, ""), (0, 2, ""), (1, 2, left), (0, 2, middle), (2, 3, left), (0, 3, middle),
           (3, 0, middle), (4, 0, wrap), (5, 0, wrap), (6, 0, wrap), (7, 0, wrap),
           (8, 3, centre), (8, 4, centre))
    cell_xfs = ""
    for f, fill, align in xfs:
        flags = (' applyFont="1"' if f else "") + (' applyFill="1"' if fill else "")
        tail = f' applyAlignment="1">{align}</xf>' if align else "/>"
        cell_xfs += f'<xf numFmtId="0" fontId="{f}" fillId="{fill}" borderId="0" xfId="0"{flags}{tail}'
    return _xml(
        f'<styleSheet xmlns="{MAIN_NS}">'
        f'<fonts count="{len(fonts)}">{"".join(fonts)}</fonts>'
        f'<fills count="5">{fills}</fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">{cell_xfs}</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>")


def home_sheet() -> bytes:
    rows = []
    links = []
    for row, height, cells in HOME_ROWS:
        body = ""
        for column, text, style, link in cells:
            ref = f"{column}{row}"
            if text:
                body += (f'<c r="{ref}" s="{style}" t="inlineStr">'
                         f'<is><t xml:space="preserve">{escape(text)}</t></is></c>')
            else:
                body += f'<c r="{ref}" s="{style}"/>'
            if link:
                links.append(f'<hyperlink ref="{ref}" location="{escape(link, {chr(34): "&quot;"})}"'
                             f' display="{escape(text)}"/>')
        rows.append(f'<row r="{row}" ht="{height}" customHeight="1">{body}</row>')
    return _xml(
        f'<worksheet xmlns="{MAIN_NS}" xmlns:r="{DOC_REL_NS}">'
        '<sheetViews><sheetView showGridLines="0" tabSelected="1" zoomScaleNormal="100" workbookViewId="0">'
        '<pane xSplit="1" ySplit="3" topLeftCell="B4" activePane="bottomRight" state="frozen"/>'
        '<selection pane="topRight" activeCell="B1" sqref="B1"/>'
        '<selection pane="bottomLeft" activeCell="A4" sqref="A4"/>'
        '<selection pane="bottomRight" activeCell="C5" sqref="C5"/>'
        "</sheetView></sheetViews>"
        '<sheetFormatPr defaultRowHeight="15"/>'
        f'<cols><col min="1" max="1" width="21.7" style="{NAV}" customWidth="1"/>'
        '<col min="2" max="2" width="4" customWidth="1"/>'
        '<col min="3" max="3" width="30" customWidth="1"/>'
        '<col min="4" max="4" width="130" customWidth="1"/>'
        '<col min="5" max="5" width="4.75" customWidth="1"/>'
        '<col min="6" max="16384" width="9.140625" hidden="1"/></cols>'
        f"<sheetData>{''.join(rows)}</sheetData>"
        f"<hyperlinks>{''.join(links)}</hyperlinks>"
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
