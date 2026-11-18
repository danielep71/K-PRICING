#!/usr/bin/env python3
"""Generate or check the macro-free workbook template src/workbook/KPR_Template.xlsx.

The template is the sheets of the K-PRICING workbook without any code: a Home
sheet with its navigation bar and the K-PRICING Ribbon (src/ribbon/customUI14.xml, byte for byte, as
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
from collections.abc import Iterator, Sequence
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
 LINK) = range(11)

# The navigation bar is the K-SACCR one (tools/navigation_bar.py there): two
# groups of linked shapes in column A. NavBarTop sits in the frozen rows 1 to
# 3 and never scrolls; NavBar is anchored on the first row below them. The
# current sheet is highlighted and not linked. The demo builder
# (examples/modules/KPR_Demo__Builder.bas) draws the same bar, in points, on
# the demo sheet; tools/test_workbook_template.py keeps the numbers equal.
PX = 9525                      # EMU per pixel at 96 dpi
BAR_X = 8                      # left margin of the bar inside column A, px
BAR_W = 141                    # bar width, px
PAD_TOP = 6                    # px above the section title
PAD_BOTTOM = 3                 # px below the last button
HEAD_H = 17                    # section title height, px
BTN_H = 21                     # button height, px
STRIP_W = 3                    # coloured strip at the left of a button, px
FROZEN_ROWS = 3
ROW_HEIGHTS = (24.0, 26.1, 33.95)      # rows 1 to 3 of the Home sheet, points
FONT = "Aptos"
HEAD_COLOUR = "FFC000"         # gold of the wordmark
BUTTON_FILL = "262626"
BUTTON_TEXT = "D9D9D9"
CURRENT_FILL = "265A99"        # blue of the sheet-title band
CURRENT_TEXT = "FFFFFF"
CURRENT_STRIP = "FFC000"
DARK_STRIP = "8FAADC"          # used when a tab colour would vanish on the button
GROUP_TOP = "NavBarTop"
GROUP_REST = "NavBar"
TAB_COLOURS = {HOME_SHEET: "1F3864", DEMO_SHEET: "548235"}   # tab colours; the strip repeats them
# (section title, ((sheet, tooltip), ...)) per group
NavBlock = tuple[tuple[str, tuple[tuple[str, str], ...]], ...]
NAV_TOP: NavBlock = (("Demo", ((DEMO_SHEET, "Each function next to Excel's own formula, with a status per row."),)),)
NAV_REST: NavBlock = (("Start", ((HOME_SHEET, "What K-PRICING provides, how to use it and the commit it was built from."),)),)

# (row, height in points, ((column, text, style, internal link or ""), ...))
HomeCell = tuple[str, str, int, str]
HOME_ROWS: tuple[tuple[int, float, tuple[HomeCell, ...]], ...] = (
    (1, ROW_HEIGHTS[0], (("B", "K-PRICING", WORDMARK, ""), ("C", "", BLACK_BAND, ""), ("D", "", BLACK_BAND, ""))),
    (2, ROW_HEIGHTS[1], (("B", "Home", SHEET_TITLE, ""), ("C", "", BLUE_BAND, ""), ("D", "", BLUE_BAND, ""))),
    (3, ROW_HEIGHTS[2], (("C", "Start here: what K-PRICING provides, how to use it and a link to every sheet."
                           " The commit this workbook was built from is at the bottom.", SUBTITLE, ""),)),
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
        '<Override PartName="/xl/drawings/drawing1.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
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
    )
    fills = ('<fill><patternFill patternType="none"/></fill>'
             '<fill><patternFill patternType="gray125"/></fill>'
             + "".join(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{rgb}"/>'
                       '<bgColor indexed="64"/></patternFill></fill>'
                       for rgb in ("000000", "265A99")))
    left = '<alignment horizontal="left" vertical="center" indent="1"/>'
    middle = '<alignment vertical="center"/>'
    wrap = '<alignment vertical="center" wrapText="1"/>'
    # (font, fill, alignment) per style index, in the order of the constants above
    xfs = ((0, 0, ""), (0, 2, ""), (1, 2, left), (0, 2, middle), (2, 3, left), (0, 3, middle),
           (3, 0, middle), (4, 0, wrap), (5, 0, wrap), (6, 0, wrap), (7, 0, wrap))
    cell_xfs = ""
    for f, fill, align in xfs:
        flags = (' applyFont="1"' if f else "") + (' applyFill="1"' if fill else "")
        tail = f' applyAlignment="1">{align}</xf>' if align else "/>"
        cell_xfs += f'<xf numFmtId="0" fontId="{f}" fillId="{fill}" borderId="0" xfId="0"{flags}{tail}'
    return _xml(
        f'<styleSheet xmlns="{MAIN_NS}">'
        f'<fonts count="{len(fonts)}">{"".join(fonts)}</fonts>'
        f'<fills count="4">{fills}</fills>'
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
        f'<sheetPr><tabColor rgb="FF{TAB_COLOURS[HOME_SHEET]}"/></sheetPr>'
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
        '<drawing r:id="rId1"/>'
        "</worksheet>")


def _strip_colour(rgb: str) -> str:
    red, green, blue = (int(rgb[i : i + 2], 16) for i in (0, 2, 4))
    return DARK_STRIP if 0.2126 * red + 0.7152 * green + 0.0722 * blue < 70 else rgb


def _shape(shape_id: int, name: str, x: int, y: int, width: int, height: int, fill: str | None, text: str = "",
           colour: str = "", bold: bool = False, size: int = 900, link: str = "", tip: str = "",
           bottom: bool = False, inset: int = 0, spacing: int = 0) -> str:
    click = ""
    if link:
        click = (f'<a:hlinkClick xmlns:r="{DOC_REL_NS}" r:id="{link}"'
                 f' tooltip="{escape(tip, {chr(34): "&quot;"})}"/>')
    fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
    body = ""
    if text:
        run = (f'<a:rPr lang="en-GB" sz="{size}" b="{int(bold)}"{f" spc={chr(34)}{spacing}{chr(34)}" if spacing else ""}>'
               f'<a:solidFill><a:srgbClr val="{colour}"/></a:solidFill>'
               f'<a:latin typeface="{FONT}"/><a:cs typeface="{FONT}"/></a:rPr>')
        body = (f'<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" lIns="{inset * PX}"'
                f' tIns="0" rIns="0" bIns="{2 * PX if bottom else 0}" rtlCol="0" anchor="{"b" if bottom else "ctr"}"/>'
                f'<a:lstStyle/><a:p><a:pPr algn="l"/><a:r>{run}<a:t>{escape(text)}</a:t></a:r></a:p></xdr:txBody>')
    return (f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{shape_id}" name="{escape(name)}">{click}'
            f'</xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="{x}" y="{y}"/>'
            f'<a:ext cx="{width}" cy="{height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'{fill_xml}<a:ln><a:noFill/></a:ln></xdr:spPr>{body}</xdr:sp>')


def _group(name: str, block: NavBlock, current: str, row: int, top: int, links: dict[str, str],
           ids: Iterator[int]) -> str:
    """One block of the bar, anchored at the top of `row` (0-based), `top` EMU down the sheet."""
    x0, y, children = BAR_X * PX, PAD_TOP, []
    for index, (title, items) in enumerate(block):
        if index:
            raise ValueError("a block holds one section")
        children.append(_shape(next(ids), "Nav_Section_" + title, x0, top + y * PX, BAR_W * PX, HEAD_H * PX, None,
                               title.upper(), HEAD_COLOUR, True, 700, bottom=True, inset=4, spacing=60))
        y += HEAD_H
        for target, tip in items:
            here = target == current
            children.append(_shape(next(ids), "Nav_" + target, x0, top + y * PX, BAR_W * PX, BTN_H * PX,
                                   CURRENT_FILL if here else BUTTON_FILL, target,
                                   CURRENT_TEXT if here else BUTTON_TEXT, here, 900,
                                   "" if here else links[target], tip, inset=12))
            children.append(_shape(next(ids), "Nav_" + target + "_Mark", x0, top + y * PX, STRIP_W * PX,
                                   BTN_H * PX, CURRENT_STRIP if here else _strip_colour(TAB_COLOURS[target])))
            y += BTN_H
    height = (y + PAD_BOTTOM) * PX
    group = (f'<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="{next(ids)}" name="{name}" descr="Sheet navigation"/>'
             f'<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="{x0}" y="{top}"/>'
             f'<a:ext cx="{BAR_W * PX}" cy="{height}"/><a:chOff x="{x0}" y="{top}"/>'
             f'<a:chExt cx="{BAR_W * PX}" cy="{height}"/></a:xfrm></xdr:grpSpPr>{"".join(children)}</xdr:grpSp>')
    return (f'<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>{x0}</xdr:colOff><xdr:row>{row}</xdr:row>'
            f'<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="{BAR_W * PX}" cy="{height}"/>{group}'
            '<xdr:clientData fPrintsWithSheet="0"/></xdr:oneCellAnchor>')


def _nav_links() -> dict[str, str]:
    """Relationship id of each sheet the Home bar links to."""
    return {target: f"rId{number}" for number, target in enumerate(
        [item[0] for block in (NAV_TOP, NAV_REST) for _, items in block for item in items if item[0] != HOME_SHEET], 1)}


def drawing() -> bytes:
    ids = iter(range(2, 1000))
    links = _nav_links()
    frozen = round(sum(ROW_HEIGHTS[:FROZEN_ROWS]) * 12700)
    anchors = (_group(GROUP_TOP, NAV_TOP, HOME_SHEET, 0, 0, links, ids)
               + _group(GROUP_REST, NAV_REST, HOME_SHEET, FROZEN_ROWS, frozen, links, ids))
    return _xml('<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
                f' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{anchors}</xdr:wsDr>')


def drawing_relationships() -> bytes:
    body = "".join(f'<Relationship Id="{rid}" Type="{DOC_REL_NS}/hyperlink" Target="#\'{escape(sheet)}\'!A1"/>'
                   for sheet, rid in _nav_links().items())
    return _xml(f'<Relationships xmlns="{PACKAGE_REL_NS}">{body}</Relationships>')


def sheet_relationships() -> bytes:
    return _xml(f'<Relationships xmlns="{PACKAGE_REL_NS}">'
                f'<Relationship Id="rId1" Type="{DOC_REL_NS}/drawing" Target="../drawings/drawing1.xml"/>'
                "</Relationships>")


def entries(customui: bytes) -> list[tuple[str, bytes]]:
    """Every part of the template, in the order the ZIP stores them."""
    return [
        ("[Content_Types].xml", content_types()),
        ("_rels/.rels", package_relationships()),
        ("xl/workbook.xml", workbook()),
        ("xl/_rels/workbook.xml.rels", workbook_relationships()),
        ("xl/styles.xml", styles()),
        ("xl/worksheets/sheet1.xml", home_sheet()),
        ("xl/worksheets/_rels/sheet1.xml.rels", sheet_relationships()),
        ("xl/drawings/drawing1.xml", drawing()),
        ("xl/drawings/_rels/drawing1.xml.rels", drawing_relationships()),
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
