#!/usr/bin/env python3
"""Generate or check the macro-free workbook template, src/workbook/Template.xlsx.

The template holds the sheets of tools/devbuild/workbook.json other than the
Test sheets, laid out as docs/methodology/WORKBOOK_CONTRACT.md describes, and
defines their names on the same cells. Its one sheet is Home:

- rows 1 to 3 are the title band: the workbook name, the sheet title and a
  subtitle; they and nothing else are frozen;
- below them, sections of labels in column B and text in column C: the build
  stamp, whose cells the build fills, what the workbook provides, how to use
  it, and its sheets, each other sheet with a link to its cell A1;
- every cell is locked and the sheet is protected without a password, as
  docs/methodology/workbook-protection.json records.

It carries no VBA project, no document properties and no external link, so
tools/check_binaries.py and tools/check_workbook_privacy.py pass it. The file
is generated, never edited in Excel: the same source always gives the same
bytes, with fixed entry times. --check (the default) fails when the committed
template differs from what this script generates, or when the sheets of
tools/devbuild/workbook.json are not the sheets this script lays out; --write
rewrites it.

Usage: python tools/gen_workbook_template.py [--check | --write]
Standard library only.
"""

from __future__ import annotations

import argparse
import math
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr

sys.path.insert(0, str(Path(__file__).resolve().parent / "devbuild"))

import build_workbook  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = build_workbook.TEMPLATE
MAIN = build_workbook.MAIN
RELATIONSHIPS = build_workbook.RELATIONSHIPS
PACKAGE = build_workbook.PACKAGE_RELATIONSHIPS
HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
WORKBOOK_NAME = "K-PRICING"
SOURCE = "https://github.com/danielep71/K-PRICING"
FONT = "Aptos"
# Characters of 10-point text that one line of column C holds, to size the
# height of a wrapped row: Excel does not fit a row to its text on opening.
LINE_CHARACTERS = 92
LINE_HEIGHT = 13.5
FROZEN_ROWS = 3
# Column widths, in characters: A a gutter, B the labels, C the text.
WIDTHS = (("A", 2.7), ("B", 26.0), ("C", 85.0))
# The protection of every sheet: on, without a password. It guards against an
# accidental edit, not against a user who means to change the sheet.
PROTECTION = {"sheet": True, "objects": True, "scenarios": True}

# Cell formats, as indexes into cellXfs: every one is locked.
(NORMAL, BAND, WORDMARK, TITLE_BAND, SHEET_TITLE, SUBTITLE, SECTION, LABEL, TEXT,
 LINK) = range(10)
TITLE_FILL, BAND_FILL = "1F3864", "265A99"


@dataclass(frozen=True)
class Cell:
    column: str
    text: str
    style: int
    link: str = ""


@dataclass(frozen=True)
class Row:
    number: int
    cells: tuple[Cell, ...]
    height: float | None = None


# What each sheet holds, in the words of the Home sheet's Sheets section.
SHEET_TEXT = {
    "shHome": "Start here: what the workbook provides, how to use it, its sheets and the "
              "commit it was built from. This sheet.",
}

ABOUT: tuple[tuple[str, tuple[tuple[str, str], ...]], ...] = (
    ("What it provides", (
        ("Functions", "The 22 KPR_Dates_* worksheet functions: Gregorian date primitives for "
                      "dates from 1 March 1900 to 31 December 9999, with strict parsing, "
                      "explicit error values and dynamic-array arguments."),
        ("Does not cover", "Business calendars and holidays, business-day arithmetic, day-count "
                           "conventions, schedules, valuation and pricing."),
        ("Source", f"{SOURCE}: the exported source, the calculation contracts and the "
                   "evidence of each release."),
    )),
    ("How to use", (
        ("1. Enable macros", "Open this .xlsm with macros enabled, as your organization's "
                             "policy allows: the functions work only while it is open."),
        ("2. Use a function", "Type =KPR_Dates_ in a cell of this workbook and pick a function. "
                              "Each returns a value or an Excel error value, and never changes "
                              "a sheet."),
        ("3. Date system", "Use the 1900 date system: in a workbook on the 1904 date system, "
                           "every function with a value argument returns #N/A, never a value "
                           "shifted by 1,462 days."),
    )),
)


def text_height(text: str) -> float | None:
    """The height of a row whose wrapped text takes more than one line."""
    lines = math.ceil(len(text) / LINE_CHARACTERS)
    return None if lines <= 1 else round(lines * LINE_HEIGHT + 3, 2)


def section(number: int, title: str) -> Row:
    return Row(number, (Cell("B", title, SECTION),), 21.0)


def entry(number: int, label: str, text: str, style: int = TEXT, link: str = "") -> Row:
    return Row(number, (Cell("B", label, LINK if link else LABEL, link),
                        Cell("C", text, style)), text_height(text))


def home_rows(layout: build_workbook.Layout) -> list[Row]:
    """The rows of the Home sheet, from the layout's stamp names and sheets."""
    rows = [Row(1, (Cell("A", "", TITLE_BAND), Cell("B", WORKBOOK_NAME, WORDMARK),
                    Cell("C", "", TITLE_BAND)), 30.0),
            Row(2, (Cell("A", "", BAND), Cell("B", "Home", SHEET_TITLE),
                    Cell("C", "", BAND)), 22.0),
            Row(3, (Cell("B", "Start here: what this workbook provides, how to use it and a "
                              "link to every sheet.", SUBTITLE),), 18.0),
            section(5, "This workbook")]
    stamp = sorted((build_workbook.cell_position(item.cell)[0], item) for item in layout.names
                   if item.sheet == "shHome" and item.name in layout.stamp.values())
    number = 6
    for row, item in stamp:
        if row != number or item.cell != f"C{row}":
            raise build_workbook.BuildError(
                f"{build_workbook.CONFIG}: the stamp names must be on consecutive cells of "
                f"column C of shHome from C6, below the section This workbook; "
                f"{item.name} is on {item.cell}")
        # The build writes the value; the template keeps the cell and its style.
        rows.append(Row(row, (Cell("B", item.label, LABEL), Cell("C", "", TEXT))))
        number += 1
    # From here, number is the last row written; one empty row precedes a section.
    number -= 1
    for title, entries in ABOUT:
        number += 2
        rows.append(section(number, title))
        for label, text in entries:
            number += 1
            rows.append(entry(number, label, text))
    number += 2
    rows.append(section(number, "Sheets"))
    for sheet in template_sheets(layout):
        number += 1
        link = "" if sheet.code_name == "shHome" else f"{quoted(sheet.tab)}!A1"
        rows.append(entry(number, sheet.tab, SHEET_TEXT[sheet.code_name], link=link))
    return rows


def quoted(tab: str) -> str:
    return "'" + tab.replace("'", "''") + "'"


def template_sheets(layout: build_workbook.Layout) -> list[build_workbook.Sheet]:
    """The sheets the template holds: every sheet but the Test sheets, in order."""
    return [sheet for sheet in layout.sheets if sheet.role != build_workbook.TEST_ROLE]


def layout_problems(layout: build_workbook.Layout) -> list[str]:
    """How the layout differs from what this script can lay out."""
    codes = [sheet.code_name for sheet in template_sheets(layout)]
    problems = []
    if codes[:1] != ["shHome"]:
        problems.append(f"{build_workbook.CONFIG}: the first sheet must be shHome, the Home "
                        "sheet")
    problems += [f"{build_workbook.CONFIG}: sheet {code} has no text in SHEET_TEXT of "
                 "tools/gen_workbook_template.py" for code in codes if code not in SHEET_TEXT]
    problems += [f"tools/gen_workbook_template.py: SHEET_TEXT describes {code}, which is not a "
                 f"sheet of {build_workbook.CONFIG}" for code in SHEET_TEXT if code not in codes]
    problems += [f"{build_workbook.CONFIG}: this script lays out only the Home sheet, not "
                 f"{code}" for code in codes[1:]]
    return problems


def xml(body: str) -> bytes:
    return (HEADER + body).encode("utf-8")


def content_types(sheets: int) -> bytes:
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
        f'ContentType="{build_workbook.WORKSHEET_CONTENT}"/>' for index in range(1, sheets + 1))
    return xml(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/xl/workbook.xml" ContentType="{build_workbook.SHEET_MAIN}"/>'
        f"{overrides}"
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>")


def package_relationships() -> bytes:
    return xml(f'<Relationships xmlns="{PACKAGE}">'
               f'<Relationship Id="rId1" Type="{build_workbook.DOCUMENT_TYPE}" '
               'Target="xl/workbook.xml"/></Relationships>')


def workbook(layout: build_workbook.Layout) -> bytes:
    sheets = template_sheets(layout)
    tabs = "".join(f"<sheet name={quoteattr(sheet.tab)} sheetId=\"{index}\" r:id=\"rId{index}\"/>"
                   for index, sheet in enumerate(sheets, 1))
    kept = {sheet.code_name for sheet in sheets}
    names = "".join(
        f"<definedName name={quoteattr(item.name)}>"
        f"{escape(build_workbook.absolute_reference(layout, item))}</definedName>"
        for item in sorted(layout.names, key=lambda item: item.name.casefold())
        if item.sheet in kept)
    return xml(f'<workbook xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}">'
               '<bookViews><workbookView activeTab="0"/></bookViews>'
               f"<sheets>{tabs}</sheets><definedNames>{names}</definedNames>"
               '<calcPr calcId="191029"/></workbook>')


def workbook_relationships(sheets: int) -> bytes:
    body = "".join(f'<Relationship Id="rId{index}" Type="{build_workbook.WORKSHEET_TYPE}" '
                   f'Target="worksheets/sheet{index}.xml"/>' for index in range(1, sheets + 1))
    return xml(f'<Relationships xmlns="{PACKAGE}">{body}'
               f'<Relationship Id="rId{sheets + 1}" Type="{build_workbook.STYLES_TYPE}" '
               'Target="styles.xml"/></Relationships>')


def styles() -> bytes:
    def font(attributes: str, size: int, rgb: str) -> str:
        return (f'<font>{attributes}<sz val="{size}"/><color rgb="FF{rgb}"/>'
                f'<name val="{FONT}"/><family val="2"/></font>')

    fonts = (font("", 11, "000000"),        # 0 body
             font("<b/>", 16, "FFFFFF"),    # 1 wordmark
             font("<b/>", 12, "FFFFFF"),    # 2 sheet title
             font("<i/>", 10, "595959"),    # 3 subtitle
             font("<b/>", 12, TITLE_FILL),  # 4 section
             font("<b/>", 10, "000000"),    # 5 label
             font("", 10, "000000"),        # 6 text
             font("<u/>", 10, "0563C1"))    # 7 link
    fills = ('<fill><patternFill patternType="none"/></fill>'
             '<fill><patternFill patternType="gray125"/></fill>'
             + "".join(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{rgb}"/>'
                       '<bgColor indexed="64"/></patternFill></fill>'
                       for rgb in (TITLE_FILL, BAND_FILL)))
    middle = '<alignment vertical="center"/>'
    top = '<alignment vertical="top" wrapText="1"/>'
    # (font, fill, alignment) of each cell format, in the order of the constants.
    formats = ((0, 0, ""), (0, 3, middle), (1, 2, middle), (0, 2, middle), (2, 3, middle),
               (3, 0, middle), (4, 0, '<alignment vertical="bottom"/>'), (5, 0, top),
               (6, 0, top), (7, 0, top))
    cells = ""
    for font_id, fill, alignment in formats:
        flags = (' applyFont="1"' if font_id else "") + (' applyFill="1"' if fill else "")
        tail = f' applyAlignment="1">{alignment}</xf>' if alignment else "/>"
        cells += (f'<xf numFmtId="0" fontId="{font_id}" fillId="{fill}" borderId="0" xfId="0"'
                  f"{flags}{tail}")
    return xml(f'<styleSheet xmlns="{MAIN}">'
               f'<fonts count="{len(fonts)}">{"".join(fonts)}</fonts>'
               f'<fills count="4">{fills}</fills>'
               '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border>'
               '</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" '
               f'borderId="0"/></cellStyleXfs><cellXfs count="{len(formats)}">{cells}</cellXfs>'
               '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/>'
               "</cellStyles></styleSheet>")


def sheet_xml(rows: Sequence[Row]) -> bytes:
    body, links = [], []
    for row in rows:
        cells = ""
        for cell in row.cells:
            reference = f"{cell.column}{row.number}"
            if cell.text:
                cells += (f'<c r="{reference}" s="{cell.style}" t="inlineStr"><is>'
                          f'<t xml:space="preserve">{escape(cell.text)}</t></is></c>')
            else:
                cells += f'<c r="{reference}" s="{cell.style}"/>'
            if cell.link:
                links.append(f'<hyperlink ref="{reference}" location={quoteattr(cell.link)} '
                             f"display={quoteattr(cell.text)}/>")
        height = f' ht="{row.height:g}" customHeight="1"' if row.height else ""
        body.append(f'<row r="{row.number}"{height}>{cells}</row>')
    columns = "".join(f'<col min="{index}" max="{index}" width="{width:g}" customWidth="1"/>'
                      for index, (_, width) in enumerate(WIDTHS, 1))
    protection = "".join(f' {name}="1"' for name, value in PROTECTION.items() if value)
    hyperlinks = f"<hyperlinks>{''.join(links)}</hyperlinks>" if links else ""
    return xml(
        f'<worksheet xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}">'
        '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView showGridLines="0" tabSelected="1" workbookViewId="0">'
        f'<pane ySplit="{FROZEN_ROWS}" topLeftCell="A{FROZEN_ROWS + 1}" activePane="bottomLeft" '
        'state="frozen"/><selection pane="bottomLeft" activeCell="B5" sqref="B5"/>'
        '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
        f"<cols>{columns}</cols><sheetData>{''.join(body)}</sheetData>"
        f"<sheetProtection{protection}/>{hyperlinks}"
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
        'footer="0.3"/><pageSetup orientation="landscape" fitToHeight="0"/></worksheet>')


def parts(layout: build_workbook.Layout) -> dict[str, bytes]:
    """Every part of the template, in the order the package stores them."""
    problems = layout_problems(layout)
    if problems:
        raise build_workbook.BuildError(problems[0])
    count = len(template_sheets(layout))
    return {"[Content_Types].xml": content_types(count),
            "_rels/.rels": package_relationships(),
            "xl/workbook.xml": workbook(layout),
            "xl/_rels/workbook.xml.rels": workbook_relationships(count),
            "xl/styles.xml": styles(),
            "xl/worksheets/sheet1.xml": sheet_xml(home_rows(layout))}


def last_row(layout: build_workbook.Layout) -> int:
    """The last row the Home sheet writes, which the protection matrix covers."""
    return home_rows(layout)[-1].number


def generate(root: Path) -> bytes:
    """The bytes of Template.xlsx: the same source always gives the same bytes."""
    return build_workbook.zip_package(parts(build_workbook.load_layout(root)))


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--check", action="store_true",
                      help="fail when the template differs from its source (the default)")
    mode.add_argument("--write", action="store_true", help="rewrite the template")
    parser.add_argument("--root", type=Path, default=ROOT, help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    root = args.root.resolve()
    try:
        expected = generate(root)
    except build_workbook.BuildError as error:
        print(f"workbook-template: FAIL ({error})")
        return 1
    path = root / TEMPLATE
    if args.write:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(expected)
        print(f"workbook-template: wrote {TEMPLATE} ({len(expected)} bytes)")
        return 0
    if not path.is_file():
        print(f"workbook-template: FAIL ({TEMPLATE} is missing; run "
              "python tools/gen_workbook_template.py --write)")
        return 1
    if path.read_bytes() != expected:
        print(f"workbook-template: FAIL ({TEMPLATE} differs from its source; run "
              "python tools/gen_workbook_template.py --write)")
        return 1
    print(f"workbook-template: PASS ({TEMPLATE} matches its source, {len(expected)} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
