#!/usr/bin/env python3
"""Generate or check the macro-free workbook template, src/workbook/Template.xlsx.

The template holds the sheets of tools/devbuild/workbook.json other than the
Test sheets, laid out as docs/methodology/WORKBOOK_CONTRACT.md describes, and
defines their names on the same cells. Its first sheet is Home:

- columns A and B are narrow margins, C holds labels, D text, E is a margin
  and the columns after it are hidden;
- rows 1 to 3 are the title band: the workbook name on a black band, the
  sheet title on a blue band, and a subtitle; they are frozen;
- below them, sections: a title, then a bordered table of labels and text,
  with a short empty row before the next section.
  They give the build stamp, whose cells the build fills, what the workbook
  provides, how to use it, and its sheets, each other sheet with a link to
  its cell A1;
- the sheet is not protected, and its cells keep Excel's default lock flag,
  as docs/methodology/workbook-protection.json records.

Its second sheet is Dates Demo, with the content of tools/demo_dates.py: the
same title band and the link back to Home, then the inputs, one row per
function call with the function's result beside a formula of Excel's own
functions and a status, and a side panel of checks, cases, a spill and the
array limits. Its formulas are live: Excel calculates them when it opens the
workbook, as dynamic-array formulas. Only its input cells are unlocked.

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
import datetime
import math
import re
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr

sys.path.insert(0, str(Path(__file__).resolve().parent / "devbuild"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import build_workbook  # noqa: E402
import demo_dates  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = build_workbook.TEMPLATE
MAIN = build_workbook.MAIN
RELATIONSHIPS = build_workbook.RELATIONSHIPS
PACKAGE = build_workbook.PACKAGE_RELATIONSHIPS
HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
WORKBOOK_NAME = "K-PRICING"
SOURCE = "https://github.com/danielep71/K-PRICING"
FONT = "Aptos"
FROZEN_ROWS = 3
# Heights in points: the title band, rows 1 to 3; a section title; a table
# row; the empty row after a table. Rows that hold nothing take the default.
BAND_HEIGHTS = (24.0, 26.1, 30.0)
SECTION_HEIGHT, TABLE_HEIGHT, GAP_HEIGHT = 30.0, 20.1, 9.95
# Column widths, in characters: A and B margins, C labels, D text, E a margin;
# the columns after E are hidden.
WIDTHS = (4.625, 4.625, 26.0, 185.375, 4.625)
# The protection of every sheet, as options of the sheetProtection element: none.
# No sheet is protected unless a decision of the workbook contract says so.
PROTECTION: dict[str, bool] = {}

BLACK, WHITE, GOLD, BLUE, NAVY, GREY, LINK_BLUE = ("000000", "FFFFFF", "FFC000", "265A99",
                                                   "1F3864", "595959", "0563C1")
# The demo's header and input fills, its input text and its border colour.
HEADER_FILL, INPUT_FILL, INPUT_TEXT, RULE_GREY = "DAE9F8", "FFCC99", "3F3F76", "A6A6A6"
# Cell formats, as indexes into cellXfs; every one is locked but DEMO_INPUT and
# DEMO_INPUT_DATE. A table cell's format is TABLE + 4 * kind + position
# (TABLE_KINDS, POSITIONS); the demo's formats follow the table formats.
(NORMAL, BLUE_BAND, SUBTITLE, CENTERED, SECTION, WORDMARK, SHEET_TITLE,
 TABLE) = range(8)
TABLE_KINDS = ("label", "text", "link")
POSITIONS = ("first", "middle", "last", "only")
DEMO_FORMATS = ("band", "header", "label", "input", "input_date", "input_note", "formula_text",
                "value", "value_date", "status", "note", "italic", "spill", "spill_date",
                "case_text", "home_link", "black_band")
(DEMO_BAND, DEMO_HEADER, DEMO_LABEL, DEMO_INPUT, DEMO_INPUT_DATE, DEMO_INPUT_NOTE,
 DEMO_FORMULA_TEXT, DEMO_VALUE, DEMO_VALUE_DATE, DEMO_STATUS, DEMO_NOTE, DEMO_ITALIC,
 DEMO_SPILL, DEMO_SPILL_DATE, DEMO_CASE_TEXT, DEMO_HOME_LINK, DEMO_BLACK_BAND) = \
    range(TABLE + 4 * len(TABLE_KINDS), TABLE + 4 * len(TABLE_KINDS) + len(DEMO_FORMATS))
# The demo's columns, in characters: A and B margins; C to G the main column
# (formula text or input label, the function's result or the input, Excel's
# reference or the input's note, the status, the note); H a gap; I to M the
# side panel (input, result, expected value, status, rule); N a margin. The
# columns after N are hidden.
DEMO_WIDTHS = (4.625, 4.625, 44.0, 15.6, 15.6, 10.0, 56.0, 2.6, 15.6, 15.6, 15.6, 10.0, 48.0,
               4.625)
DEMO_SHEET = "shDatesDemo"
DEMO_CHECKS = "DatesDemo_Checks"
# The demo's dates are serials of the 1900 date system, the system of the workbook.
EPOCH = datetime.date(1899, 12, 30)
DATE_FORMAT_ID, TEXT_FORMAT_ID = 164, 49
# Excel calls the functions of Excel 2010 and later by a prefixed name in the file.
FILE_NAMES = (("SEQUENCE(", "_xlfn.SEQUENCE("),)


@dataclass(frozen=True)
class Cell:
    """A cell: inline text, or a typed value, or a formula as Excel shows it
    (starting with "="), which spills over ``spill`` rows."""

    column: str
    text: str
    style: int
    link: str = ""
    value: demo_dates.Value = None
    formula: str = ""
    spill: int = 1


@dataclass(frozen=True)
class Row:
    number: int
    cells: tuple[Cell, ...]
    height: float | None = None


@dataclass(frozen=True)
class SheetLayout:
    """What a sheet part needs besides its rows: its columns, the last visible
    one, whether it is the selected tab, and where the cursor starts."""

    widths: tuple[float, ...]
    selected: bool
    active_cell: str


HOME_LAYOUT = SheetLayout(WIDTHS, True, "C4")
DEMO_LAYOUT = SheetLayout(DEMO_WIDTHS, False, "D5")

# What each sheet holds, in the words of the Home sheet's Sheets section.
SHEET_TEXT = {
    "shHome": "Start here: what the workbook provides, how to use it, its sheets and the "
              "commit it was built from. This sheet.",
    DEMO_SHEET: "A worked example of each of the 22 functions on inputs you can edit, beside "
                "Excel's own formulas, with a status per row and a count of the statuses.",
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


def table_style(kind: str, index: int, count: int) -> int:
    """The format of a table cell: its kind, and its row's place in the table."""
    position = "only" if count == 1 else "first" if index == 0 else \
        "last" if index == count - 1 else "middle"
    return TABLE + 4 * TABLE_KINDS.index(kind) + POSITIONS.index(position)


def section(number: int, title: str, entries: Sequence[tuple[str, str, str]]) -> list[Row]:
    """A section from row ``number``: its title, then a table of (label, text, link)
    entries, one row each."""
    rows = [Row(number, (Cell("C", title, SECTION),), SECTION_HEIGHT)]
    for index, (label, text, link) in enumerate(entries):
        rows.append(Row(number + 1 + index, (
            Cell("C", label, table_style("link" if link else "label", index, len(entries)),
                 link),
            Cell("D", text, table_style("text", index, len(entries)))), TABLE_HEIGHT))
    return rows


def home_rows(layout: build_workbook.Layout) -> list[Row]:
    """The rows of the Home sheet, from the layout's stamp names and sheets."""
    stamp = sorted((build_workbook.cell_position(item.cell)[0], item) for item in layout.names
                   if item.sheet == "shHome" and item.name in layout.stamp.values())
    for offset, (row, item) in enumerate(stamp):
        if row != 5 + offset or item.cell != f"D{row}":
            raise build_workbook.BuildError(
                f"{build_workbook.CONFIG}: the stamp names must be on consecutive cells of "
                f"column D of shHome from D5, below the section This workbook; "
                f"{item.name} is on {item.cell}")
    sections = [
        # The build writes each stamp value; the template keeps the cell and its style.
        ("This workbook", [(item.label, "", "") for _, item in stamp]),
        *[(title, [(label, text, "") for label, text in entries]) for title, entries in ABOUT],
        ("Sheets", [(sheet.tab, SHEET_TEXT[sheet.code_name],
                     "" if sheet.code_name == "shHome" else f"{quoted(sheet.tab)}!A1")
                    for sheet in template_sheets(layout)]),
    ]
    rows = [Row(1, (Cell("B", WORKBOOK_NAME, WORDMARK), Cell("C", "", WORDMARK),
                    Cell("D", "", WORDMARK)), BAND_HEIGHTS[0]),
            Row(2, (Cell("B", "Home", SHEET_TITLE), Cell("C", "", BLUE_BAND),
                    Cell("D", "", SHEET_TITLE)), BAND_HEIGHTS[1]),
            Row(3, (Cell("C", "Start here: what this workbook provides, how to use it and a "
                              "link to every sheet.", SUBTITLE),), BAND_HEIGHTS[2])]
    number = 4
    for index, (title, entries) in enumerate(sections):
        if index:
            # A short empty row separates a table from the next section.
            rows.append(Row(number, (), GAP_HEIGHT))
            number += 1
        rows += section(number, title, entries)
        number = rows[-1].number + 1
    return rows


# ---------------------------------------------------------------------------
# The Dates Demo sheet.

MAIN_FIRST, MAIN_LAST, SIDE_FIRST, SIDE_LAST = "C", "G", "I", "M"
BAND_HEIGHT, LINE_HEIGHT = 20.1, 13.5
STATUS_OK, STATUS_DIFFERS = "OK", "DIFFERS"


def letter(number: int) -> str:
    return build_workbook.column_letters(number)


def column(text: str) -> int:
    return build_workbook.column_number(text)


def offset(text: str, by: int) -> str:
    return letter(column(text) + by)


def status_formula(actual: str, expected: str, differs_when: str = "") -> str:
    """The status of a result: OK when it equals the expected value, or both are the
    same error value; DIFFERS when they differ while ``differs_when`` holds; NO VBA
    while the functions are not loaded (#NAME?); FAIL otherwise."""
    same = (f"IF(ISERROR({expected}),IFERROR(ERROR.TYPE({actual})=ERROR.TYPE({expected}),"
            f"FALSE),IFERROR({actual}={expected},FALSE))")
    other = (f'IF(IFERROR({differs_when},FALSE),"{STATUS_DIFFERS}","FAIL")' if differs_when
             else '"FAIL"')
    return (f'=IF(IFERROR(ERROR.TYPE({actual})=5,FALSE),"NO VBA",IF({same},"{STATUS_OK}",'
            f"{other}))")


def lines(text: str, width: float) -> int:
    """The lines a wrapped text takes in a column of ``width`` characters."""
    return max(1, math.ceil(len(text) / max(width * 1.1, 1.0)))


@dataclass
class DemoSheet:
    """The demo's rows as they are written: cells by row, heights, and the
    cells every status is in."""

    cells: dict[int, list[Cell]]
    heights: dict[int, float]
    statuses: list[str]

    def put(self, row: int, *cells: Cell, height: float = BAND_HEIGHT) -> None:
        self.cells.setdefault(row, []).extend(cells)
        self.heights[row] = max(self.heights.get(row, 0.0), height)

    def band(self, row: int, first: str, last: str, title: str) -> None:
        """A band across ``first`` to ``last``, its title centred across it."""
        self.put(row, *(Cell(letter(index), title if index == column(first) else "", DEMO_BAND)
                        for index in range(column(first), column(last) + 1)))

    def header(self, row: int, first: str, titles: Sequence[str]) -> None:
        self.put(row, *(Cell(offset(first, index), title, DEMO_HEADER)
                        for index, title in enumerate(titles)))

    def status(self, row: int, at: str, formula: str) -> None:
        self.put(row, Cell(at, "", DEMO_STATUS, formula=formula))
        self.statuses.append(f"{at}{row}")


def value_style(value: demo_dates.Value, date_style: int, text_style: int, other: int) -> int:
    if isinstance(value, datetime.date):
        return date_style
    return text_style if isinstance(value, str) else other


def write_inputs(sheet: DemoSheet, layout: build_workbook.Layout, row: int) -> int:
    """The inputs from ``row``: the band, then one named, editable cell each."""
    sheet.band(row, MAIN_FIRST, "E", demo_dates.INPUTS_TITLE)
    names = {item.name: item for item in layout.names if item.sheet == DEMO_SHEET}
    for index, item in enumerate(demo_dates.INPUTS, row + 1):
        named = names.get(item.name)
        if named is None or named.cell != f"D{index}":
            raise build_workbook.BuildError(
                f"{build_workbook.CONFIG}: the demo input {item.name} must be a name of "
                f"{DEMO_SHEET} on D{index}")
        style = DEMO_INPUT_DATE if isinstance(item.value, datetime.date) else DEMO_INPUT
        sheet.put(index, Cell("C", item.label, DEMO_LABEL), Cell("D", "", style, value=item.value),
                  Cell("E", item.note, DEMO_INPUT_NOTE))
    return row + len(demo_dates.INPUTS)


def write_sections(sheet: DemoSheet, row: int) -> int:
    """Every comparison section after ``row``, each after a blank row."""
    for section in demo_dates.SECTIONS:
        row += 2
        sheet.band(row, MAIN_FIRST, MAIN_LAST, section.title)
        sheet.header(row + 1, MAIN_FIRST,
                     ("Formula", "KPR function", "Expected (Excel)", "Status", "Notes"))
        row += 1
        for item in section.rows:
            row += 1
            style = DEMO_VALUE_DATE if item.date else DEMO_VALUE
            height = LINE_HEIGHT * max(lines(item.note, DEMO_WIDTHS[6]),
                                       lines(item.formula, DEMO_WIDTHS[2] * 0.8)) + 6.6
            sheet.put(row, Cell("C", item.formula, DEMO_FORMULA_TEXT),
                      Cell("D", "", style, formula=item.formula),
                      Cell("E", "", style, formula=item.reference),
                      Cell("G", item.note, DEMO_NOTE), height=max(BAND_HEIGHT, height))
            sheet.status(row, "F", status_formula(f"D{row}", f"E{row}", item.differs_when))
    for note in demo_dates.NOTES:
        row += 2
        sheet.put(row, Cell("C", note, DEMO_ITALIC))
    return row


def write_cases(sheet: DemoSheet, table: demo_dates.CaseTable, row: int) -> int:
    """A table of cases from ``row``: band, header, then one row per case."""
    sheet.band(row, SIDE_FIRST, SIDE_LAST, table.title)
    sheet.header(row + 1, SIDE_FIRST, ("Input", "Result", "Expected", "Status", "Rule"))
    row += 1
    for case in table.cases:
        row += 1
        height = LINE_HEIGHT * lines(case.rule, DEMO_WIDTHS[12]) + 6.6
        cells = [Cell("J", "", DEMO_VALUE,
                      formula=case.formula.replace("{IN}", f"I{row}")),
                 Cell("K", "", DEMO_VALUE, value=case.expected),
                 Cell("M", case.rule, DEMO_NOTE)]
        if case.value is not None:
            cells.insert(0, Cell("I", "", value_style(case.value, DEMO_VALUE_DATE,
                                                       DEMO_CASE_TEXT, DEMO_VALUE),
                                 value=case.value))
        else:
            cells.insert(0, Cell("I", "", DEMO_VALUE))
        sheet.put(row, *cells, height=max(BAND_HEIGHT, height))
        sheet.status(row, "L", status_formula(f"J{row}", f"K{row}"))
    return row


def write_spill(sheet: DemoSheet, row: int) -> int:
    """The spill block from ``row``: band, headers, a status per column, then the
    source spill and one dependent spill per column."""
    spill = demo_dates.SPILL
    last = offset(SIDE_FIRST, len(spill.columns))
    sheet.band(row, SIDE_FIRST, last, spill.title)
    sheet.header(row + 1, SIDE_FIRST, (spill.header, *(item.header for item in spill.columns)))
    check, anchor = row + 2, row + 3
    sheet.put(check, Cell(SIDE_FIRST, "Check", DEMO_STATUS))
    sheet.put(anchor, Cell(SIDE_FIRST, "", DEMO_SPILL_DATE, formula=spill.formula,
                           spill=spill.rows))
    source = f"{SIDE_FIRST}{anchor}#"
    for index, item in enumerate(spill.columns, 1):
        at = offset(SIDE_FIRST, index)
        own = f"{at}{anchor}#"
        sheet.put(anchor, Cell(at, "", DEMO_SPILL_DATE if item.date else DEMO_SPILL,
                               formula=item.formula.replace("{SRC}", source), spill=spill.rows))
        reference = item.reference.replace("{SRC}", source)
        sheet.status(check, at, f'=IF(IFERROR(ERROR.TYPE({at}{anchor})=5,FALSE),"NO VBA",'
                                f"IFERROR(IF(AND(ROWS({own})=ROWS({source}),{own}={reference}),"
                                '"OK","FAIL"),"FAIL"))')
    # Every row of each spill range carries its format, so that the spilled values
    # below the anchor show as dates where the column holds dates.
    formats = [DEMO_SPILL_DATE, *(DEMO_SPILL_DATE if item.date else DEMO_SPILL
                                  for item in spill.columns)]
    for extra in range(anchor + 1, anchor + spill.rows):
        sheet.put(extra, *(Cell(offset(SIDE_FIRST, index), "", style)
                           for index, style in enumerate(formats)), height=0.0)
    return anchor + spill.rows - 1


def write_checks(sheet: DemoSheet, layout: build_workbook.Layout, row: int) -> None:
    """The count of statuses from ``row``, under the anchor DEMO_CHECKS."""
    anchor = next((item for item in layout.names if item.name == DEMO_CHECKS), None)
    if anchor is None or anchor.sheet != DEMO_SHEET or anchor.cell != f"I{row + 1}":
        raise build_workbook.BuildError(
            f"{build_workbook.CONFIG}: the name {DEMO_CHECKS} must be on I{row + 1} of "
            f"{DEMO_SHEET}, the first header cell of the checks")
    sheet.band(row, SIDE_FIRST, SIDE_LAST, demo_dates.CHECKS_TITLE)
    sheet.header(row + 1, SIDE_FIRST, ("Check", "Count", "Meaning"))
    ok = "+".join(f'({cell}="{STATUS_OK}")' for cell in sheet.statuses)
    differs = "+".join(f'({cell}="{STATUS_DIFFERS}")' for cell in sheet.statuses)
    count = len(sheet.statuses)
    lines_ = (("OK", f"={ok}", "The function's result equals Excel's own formula."),
              ("Differs by design", f"={differs}",
               "A documented difference; the row's note explains it."),
              ("Fail", f"={count}-J{row + 2}-J{row + 3}",
               "Must be 0. NO VBA means the functions are not loaded."),
              ("Checks", f"={count}", "Every status on this sheet."))
    for index, (label, formula, note) in enumerate(lines_, row + 2):
        sheet.put(index, Cell("I", label, DEMO_LABEL), Cell("J", "", DEMO_STATUS, formula=formula),
                  Cell("K", note, DEMO_INPUT_NOTE))
    sheet.put(row + 6, Cell("I", "Built from source with the workbook. Edit the orange cells: "
                                 "every status recalculates.", DEMO_ITALIC))


def demo_rows(layout: build_workbook.Layout) -> list[Row]:
    """The rows of the Dates Demo sheet."""
    sheet = DemoSheet({}, {}, [])
    sheet.put(1, Cell("B", WORKBOOK_NAME, WORDMARK),
              *(Cell(letter(index), "Home" if index == 4 else "",
                     DEMO_HOME_LINK if index == 4 else DEMO_BLACK_BAND,
                     link="Home!A1" if index == 4 else "")
                for index in range(3, column(SIDE_LAST) + 1)), height=BAND_HEIGHTS[0])
    sheet.put(2, Cell("B", demo_dates.TITLE, SHEET_TITLE),
              *(Cell(letter(index), "", BLUE_BAND) for index in range(3, column(SIDE_LAST) + 1)),
              height=BAND_HEIGHTS[1])
    sheet.put(3, Cell("C", demo_dates.SUBTITLE, SUBTITLE), height=BAND_HEIGHTS[2])
    main = write_sections(sheet, write_inputs(sheet, layout, 4))
    side = write_cases(sheet, demo_dates.CASES, 12)
    side = write_spill(sheet, side + 2)
    side = write_cases(sheet, demo_dates.LIMITS, side + 2)
    # The checks count every status above, so they are written last, at the top.
    write_checks(sheet, layout, 4)
    rows = []
    for number in range(1, max(main, side) + 1):
        cells = tuple(sorted(sheet.cells.get(number, []), key=lambda cell: column(cell.column)))
        height = sheet.heights.get(number)
        rows.append(Row(number, cells, height or None))
    return [row for row in rows if row.cells or row.height]


def demo_statuses(layout: build_workbook.Layout) -> list[str]:
    """The cells that hold a status on the demo sheet, the checks' own counts aside."""
    return [f"{cell.column}{row.number}" for row in demo_rows(layout) for cell in row.cells
            if cell.style == DEMO_STATUS and cell.formula
            and not (cell.column == "J" and row.number in range(6, 10))]


def demo_protection(layout: build_workbook.Layout) -> dict[str, list[str]]:
    """The demo's input and output ranges, as the protection matrix records them."""
    last = demo_rows(layout)[-1].number
    first_input, last_input = 5, 4 + len(demo_dates.INPUTS)
    edge = letter(len(DEMO_WIDTHS))
    return {"input": [f"D{first_input}:D{last_input}"],
            "output": [f"A1:{edge}{first_input - 1}", f"A{first_input}:C{last_input}",
                       f"E{first_input}:{edge}{last_input}", f"A{last_input + 1}:{edge}{last}"]}


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
    problems += [f"{build_workbook.CONFIG}: this script lays out the Home and Dates Demo "
                 f"sheets, not {code}" for code in codes[1:] if code != DEMO_SHEET]
    return problems


def xml(body: str) -> bytes:
    return (HEADER + body).encode("utf-8")


METADATA_CONTENT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
METADATA_TYPE = f"{RELATIONSHIPS}/sheetMetadata"


def content_types(sheets: int) -> bytes:
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
        f'ContentType="{build_workbook.WORKSHEET_CONTENT}"/>' for index in range(1, sheets + 1))
    overrides += f'<Override PartName="/xl/metadata.xml" ContentType="{METADATA_CONTENT}"/>'
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
               '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')


def workbook_relationships(sheets: int) -> bytes:
    body = "".join(f'<Relationship Id="rId{index}" Type="{build_workbook.WORKSHEET_TYPE}" '
                   f'Target="worksheets/sheet{index}.xml"/>' for index in range(1, sheets + 1))
    return xml(f'<Relationships xmlns="{PACKAGE}">{body}'
               f'<Relationship Id="rId{sheets + 1}" Type="{build_workbook.STYLES_TYPE}" '
               'Target="styles.xml"/>'
               f'<Relationship Id="rId{sheets + 2}" Type="{METADATA_TYPE}" '
               'Target="metadata.xml"/></Relationships>')


def metadata() -> bytes:
    """The cell metadata that marks a formula as a dynamic-array formula, as Excel writes it."""
    return xml(f'<metadata xmlns="{MAIN}" '
               'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
               '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" '
               'copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" '
               'clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>'
               '</metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst>'
               '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
               '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>'
               '</futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk>'
               "</cellMetadata></metadata>")


def styles() -> bytes:
    def font(attributes: str, size: int, rgb: str) -> str:
        return (f'<font>{attributes}<sz val="{size}"/><color rgb="FF{rgb}"/>'
                f'<name val="{FONT}"/><family val="2"/></font>')

    def line(side: str, style: str) -> str:
        return f'<{side} style="{style}"><color auto="1"/></{side}>'

    fonts = (font("", 11, BLACK),         # 0 body
             font("<i/>", 10, GREY),      # 1 subtitle
             font("<b/>", 12, NAVY),      # 2 section
             font("", 14, GOLD),          # 3 wordmark
             font("<b/>", 18, WHITE),     # 4 sheet title
             font("<b/>", 10, BLACK),     # 5 label
             font("", 9, BLACK),          # 6 text
             font("<u/>", 10, LINK_BLUE),  # 7 link
             font("<b/>", 12, WHITE),     # 8 demo band
             font("<b/>", 11, BLACK),     # 9 demo header
             font("", 10, INPUT_TEXT),    # 10 demo input
             font("<i/>", 10, BLACK),     # 11 demo italic
             font("", 10, BLACK),         # 12 demo body
             f'<font><sz val="10"/><color rgb="FF{BLACK}"/><name val="Consolas"/>'
             '<family val="3"/></font>',  # 13 demo formula text
             font("<u/>", 10, WHITE))     # 14 demo link on the black band
    fills = ('<fill><patternFill patternType="none"/></fill>'
             '<fill><patternFill patternType="gray125"/></fill>'
             + "".join(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{rgb}"/>'
                       '<bgColor indexed="64"/></patternFill></fill>'
                       for rgb in (BLUE, BLACK, NAVY, HEADER_FILL, INPUT_FILL)))
    # A table is a box of thin lines, with dashed lines between its rows and
    # between its two columns.
    borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>']
    for position in POSITIONS:
        top = "thin" if position in ("first", "only") else "dashed"
        bottom = "thin" if position in ("last", "only") else "dashed"
        for left, right in (("thin", "dashed"), ("dashed", "thin")):
            borders.append(f"<border>{line('left', left)}{line('right', right)}"
                           f"{line('top', top)}{line('bottom', bottom)}<diagonal/></border>")
    # The demo boxes each cell in thin grey lines, and each band in its own colour.
    for rgb in (RULE_GREY, NAVY):
        side = "".join(f'<{edge} style="thin"><color rgb="FF{rgb}"/></{edge}>'
                       for edge in ("left", "right", "top", "bottom"))
        borders.append(f"<border>{side}<diagonal/></border>")
    grey, band = len(borders) - 2, len(borders) - 1
    center = '<alignment vertical="center"/>'
    banded = '<alignment horizontal="left" vertical="center"/>'
    wrapped = '<alignment vertical="center" wrapText="1"/>'
    middle = '<alignment horizontal="center" vertical="center"/>'
    across = '<alignment horizontal="centerContinuous" vertical="center"/>'
    plain = '<alignment vertical="center"/>'
    # (font, fill, border, alignment, number format, unlocked) of each cell format, in the
    # order of the constants.
    formats: list[tuple[int, int, int, str, int, bool]] = [
        (0, 0, 0, "", 0, False), (0, 2, 0, center, 0, False), (1, 0, 0, center, 0, False),
        (0, 0, 0, center, 0, False), (2, 0, 0, center, 0, False), (3, 3, 0, banded, 0, False),
        (4, 2, 0, banded, 0, False)]
    for kind in TABLE_KINDS:
        font_id, column = {"label": (5, 0), "text": (6, 1), "link": (7, 0)}[kind]
        formats += [(font_id, 0, 1 + 2 * index + column, wrapped, 0, False)
                    for index in range(len(POSITIONS))]
    date, text = DATE_FORMAT_ID, TEXT_FORMAT_ID
    demo = {"band": (8, 4, band, across, 0, False),
            "header": (9, 5, grey, middle, 0, False),
            "label": (5, 5, grey, plain, 0, False),
            "input": (10, 6, grey, middle, 0, True),
            "input_date": (10, 6, grey, middle, date, True),
            "input_note": (11, 0, grey, plain, 0, False),
            "formula_text": (13, 0, grey, wrapped, text, False),
            "value": (12, 0, grey, middle, 0, False),
            "value_date": (12, 0, grey, middle, date, False),
            "status": (5, 0, grey, middle, 0, False),
            "note": (12, 0, grey, wrapped, 0, False),
            "italic": (11, 0, 0, plain, 0, False),
            "spill": (12, 0, 0, middle, 0, False),
            "spill_date": (12, 0, 0, middle, date, False),
            "case_text": (12, 0, grey, middle, text, False),
            "home_link": (14, 3, 0, banded, 0, False),
            "black_band": (0, 3, 0, "", 0, False)}
    formats += [demo[name] for name in DEMO_FORMATS]
    cells = ""
    for font_id, fill, border, alignment, number, unlocked in formats:
        flags = (' applyNumberFormat="1"' if number else "") + \
            (' applyFont="1"' if font_id else "") + (' applyFill="1"' if fill else "") + \
            (' applyBorder="1"' if border else "") + \
            (' applyAlignment="1"' if alignment else "") + \
            (' applyProtection="1"' if unlocked else "")
        inner = alignment + ('<protection locked="0"/>' if unlocked else "")
        tail = f">{inner}</xf>" if inner else "/>"
        cells += (f'<xf numFmtId="{number}" fontId="{font_id}" fillId="{fill}" '
                  f'borderId="{border}" xfId="0"{flags}{tail}')
    return xml(f'<styleSheet xmlns="{MAIN}">'
               f'<numFmts count="1"><numFmt numFmtId="{DATE_FORMAT_ID}" '
               'formatCode="yyyy-mm-dd"/></numFmts>'
               f'<fonts count="{len(fonts)}">{"".join(fonts)}</fonts>'
               f'<fills count="7">{fills}</fills>'
               f'<borders count="{len(borders)}">{"".join(borders)}</borders>'
               '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" '
               f'borderId="0"/></cellStyleXfs><cellXfs count="{len(formats)}">{cells}</cellXfs>'
               '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/>'
               "</cellStyles></styleSheet>")


def file_formula(formula: str) -> str:
    """A formula as the file stores it: without its "=", with the prefixed names of
    Excel's newer functions, and a spill reference such as I30# as ANCHORARRAY."""
    text = formula[1:] if formula.startswith("=") else formula
    for shown, stored in FILE_NAMES:
        text = text.replace(shown, stored)
    return re.sub(r"\b([A-Z]{1,3}[0-9]+)#", r"_xlfn.ANCHORARRAY(\1)", text)


def serial(value: datetime.date) -> int:
    return (value - EPOCH).days


def cell_xml(reference: str, cell: Cell, row: int) -> str:
    """One cell: a formula, a typed value, inline text, or an empty styled cell."""
    head = f'<c r="{reference}" s="{cell.style}"'
    if cell.formula:
        # Every formula is a dynamic-array formula (cell metadata 1), so Excel calculates
        # it as it would one typed today, without implicit intersection.
        last = f"{cell.column}{row + cell.spill - 1}"
        return (f'{head} cm="1"><f t="array" ref="{reference}'
                f'{":" + last if cell.spill > 1 else ""}">'
                f"{escape(file_formula(cell.formula))}</f></c>")
    value = cell.value
    if isinstance(value, demo_dates.ErrorValue):
        return f'{head} t="e"><v>{escape(value.text)}</v></c>'
    if isinstance(value, bool):
        return f'{head} t="b"><v>{int(value)}</v></c>'
    if isinstance(value, datetime.date):
        return f"{head}><v>{serial(value)}</v></c>"
    if isinstance(value, (int, float)):
        return f"{head}><v>{value:g}</v></c>"
    text = value if isinstance(value, str) else cell.text
    if text:
        return (f'{head} t="inlineStr"><is><t xml:space="preserve">{escape(text)}</t></is>'
                "</c>")
    return f"{head}/>"


def sheet_xml(rows: Sequence[Row], layout: SheetLayout = HOME_LAYOUT) -> bytes:
    body, links = [], []
    for row in rows:
        cells = ""
        for cell in row.cells:
            reference = f"{cell.column}{row.number}"
            cells += cell_xml(reference, cell, row.number)
            if cell.link:
                links.append(f'<hyperlink ref="{reference}" location={quoteattr(cell.link)} '
                             f"display={quoteattr(cell.text)}/>")
        height = f' ht="{row.height:g}" customHeight="1"' if row.height else ""
        body.append(f'<row r="{row.number}"{height}>{cells}</row>')
    # Every column takes the centred format; A and B share one element, as Excel writes it.
    widths = layout.widths
    columns = (f'<col min="1" max="2" width="{widths[0]:g}" style="{CENTERED}" customWidth="1"/>'
               + "".join(f'<col min="{index}" max="{index}" width="{width:g}" '
                         f'style="{CENTERED}" customWidth="1"/>'
                         for index, width in enumerate(widths[2:], 3))
               + f'<col min="{len(widths) + 1}" max="16384" width="9" style="{CENTERED}" '
               'hidden="1"/>')
    options = "".join(f' {name}="1"' for name, value in PROTECTION.items() if value)
    protection = f"<sheetProtection{options}/>" if PROTECTION.get("sheet") else ""
    hyperlinks = f"<hyperlinks>{''.join(links)}</hyperlinks>" if links else ""
    selected = ' tabSelected="1"' if layout.selected else ""
    active = layout.active_cell
    return xml(
        f'<worksheet xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}">'
        '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
        f'<sheetViews><sheetView showGridLines="0"{selected} workbookViewId="0">'
        f'<pane ySplit="{FROZEN_ROWS}" topLeftCell="A{FROZEN_ROWS + 1}" activePane="bottomLeft" '
        f'state="frozen"/><selection pane="bottomLeft" activeCell="{active}" sqref="{active}"/>'
        '</sheetView></sheetViews><sheetFormatPr defaultColWidth="0" defaultRowHeight="15"/>'
        f'<cols>{columns}</cols><sheetData>{"".join(body)}</sheetData>'
        f"{protection}{hyperlinks}"
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
        'footer="0.3"/><pageSetup orientation="landscape" fitToHeight="0"/></worksheet>')


def parts(layout: build_workbook.Layout) -> dict[str, bytes]:
    """Every part of the template, in the order the package stores them."""
    problems = layout_problems(layout)
    if problems:
        raise build_workbook.BuildError(problems[0])
    count = len(template_sheets(layout))
    found = {"[Content_Types].xml": content_types(count),
             "_rels/.rels": package_relationships(),
             "xl/workbook.xml": workbook(layout),
             "xl/_rels/workbook.xml.rels": workbook_relationships(count),
             "xl/styles.xml": styles(),
             "xl/metadata.xml": metadata()}
    for index, sheet in enumerate(template_sheets(layout), 1):
        found[f"xl/worksheets/sheet{index}.xml"] = (
            sheet_xml(home_rows(layout)) if sheet.code_name == "shHome"
            else sheet_xml(demo_rows(layout), DEMO_LAYOUT))
    return found


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
