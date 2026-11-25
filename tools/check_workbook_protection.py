#!/usr/bin/env python3
"""Check the protection of the layout template against its protection matrix.

The protection matrix, docs/methodology/workbook-protection.json, states for
each sheet of src/workbook/Template.xlsx, by the code name that
tools/devbuild/workbook.json gives it:

- protection: the options of the sheet's protection, named as the attributes
  of the SpreadsheetML sheetProtection element (ECMA-376). sheet true turns
  protection on; each other option true blocks its action, or protects its
  part, while the sheet is protected. An option left out takes the default of
  the standard, as a missing attribute does;
- input: the ranges users type in, which must be unlocked;
- output: the ranges users must not change, which must be locked.

The check compares the matrix with the template's sheet parts and styles.xml.
A cell takes the style of its own cell element, or else of its row when the
row has a custom format, or else of its column, or else the default style. A
style is locked unless its protection element says locked="0"; without a
protection element of its own, or with applyProtection="0", the cell style it
is based on decides. The password of a sheet is not compared, and the matrix
never holds one. Protection that VBA sets at run time with Worksheet.Protect
is not compared either.

Every sheet of the template needs an entry, and every entry a sheet of the
template. Without a template the check reads the matrix alone, and without
either it passes and says so: the template and its matrix are conventions
until a project adds them. A template without a matrix fails, because its
lock flags would go unchecked. It never changes a file.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

sys.path.insert(0, str(Path(__file__).resolve().parent / "devbuild"))

import build_workbook  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
MATRIX = "docs/methodology/workbook-protection.json"
TEMPLATE = build_workbook.TEMPLATE
MAIN = f"{{{build_workbook.MAIN}}}"
# The sheetProtection attributes of ECMA-376 with their defaults.
OPTIONS: dict[str, bool] = {
    "sheet": False, "objects": False, "scenarios": False, "formatCells": True,
    "formatColumns": True, "formatRows": True, "insertColumns": True, "insertRows": True,
    "insertHyperlinks": True, "deleteColumns": True, "deleteRows": True,
    "selectLockedCells": False, "sort": True, "autoFilter": True, "pivotTables": True,
    "selectUnlockedCells": False,
}
KINDS = {"input": False, "output": True}  # whether the range must be locked
RANGE = re.compile(r"^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$")
MAX_ROW, MAX_COLUMN = 1_048_576, 16_384


@dataclass(frozen=True)
class Area:
    kind: str
    text: str
    top: int
    left: int
    bottom: int
    right: int


@dataclass(frozen=True)
class Entry:
    code_name: str
    protection: dict[str, bool]
    areas: tuple[Area, ...]


@dataclass(frozen=True)
class SheetState:
    """What the template sets for one sheet: protection options and cell styles."""

    protection: dict[str, bool]
    # The style of each cell element, by row, then column.
    cells: dict[int, dict[int, int]]
    # The style of each row with a custom format.
    rows: dict[int, int]
    columns: tuple[tuple[int, int, int], ...]

    def column_runs(self, left: int, right: int) -> list[tuple[int, int, int]]:
        """(first, last, style) runs covering columns left to right of a row that has no
        style of its own: each column takes its first <col>, or else the default style."""
        edges = {left, right + 1}
        for low, high, _ in self.columns:
            edges |= {edge for edge in (low, high + 1) if left < edge <= right}
        bounds = sorted(edges)
        return [(first, after - 1, next((style for low, high, style in self.columns
                                         if low <= first <= high), 0))
                for first, after in zip(bounds, bounds[1:])]


def is_locked(flags: list[bool], style: int) -> bool:
    return flags[style] if style < len(flags) else True


def row_mismatches(state: SheetState, flags: list[bool], area: Area, row: int | None,
                   runs: list[tuple[int, int, int]]) -> tuple[int, int | None]:
    """How many cells of one row of the area have the wrong lock flag, and the first such
    column. ``row`` None is a row with neither cell elements nor a style of its own."""
    locked = KINDS[area.kind]
    explicit = {column: style for column, style in state.cells.get(row or 0, {}).items()
                if area.left <= column <= area.right}
    if row in state.rows:
        runs = [(area.left, area.right, state.rows[row])]
    wrong = [column for column, style in explicit.items() if is_locked(flags, style) != locked]
    count, first = len(wrong), min(wrong, default=None)
    for low, high, style in runs:
        if is_locked(flags, style) == locked:
            continue
        inside = sorted(column for column in explicit if low <= column <= high)
        count += high - low + 1 - len(inside)
        column = low
        for taken in inside:
            if taken != column:
                break
            column += 1
        if column <= high:
            first = column if first is None else min(first, column)
    return count, first


def area_mismatches(state: SheetState, flags: list[bool], area: Area
                    ) -> tuple[int, tuple[int, int] | None]:
    """How many cells of the area have the wrong lock flag, and the first such cell, in
    reading order. Rows without cell elements or a style of their own are all alike, so
    they are counted once; the cost follows the rows the sheet sets, not the area's size."""
    runs = state.column_runs(area.left, area.right)
    special = sorted(row for row in {*state.cells, *state.rows}
                     if area.top <= row <= area.bottom)
    plain_count, plain_first = row_mismatches(state, flags, area, None, runs)
    count = plain_count * (area.bottom - area.top + 1 - len(special))
    first: tuple[int, int] | None = None
    if plain_count and plain_first is not None:
        row = area.top
        for taken in special:
            if taken != row:
                break
            row += 1
        if row <= area.bottom:
            first = (row, plain_first)
    for row in special:
        found, column = row_mismatches(state, flags, area, row, runs)
        count += found
        if column is not None and (first is None or (row, column) < first):
            first = (row, column)
    return count, first


def flag(text: str | None, default: bool) -> bool:
    """An xsd:boolean attribute."""
    return default if text is None else text in ("1", "true")


def parse_area(kind: str, text: object) -> Area | str:
    """The range, or why it is not one."""
    match = RANGE.match(text) if isinstance(text, str) else None
    if match is None:
        return f"{kind} range {text!r} is not a cell or a range such as C5:D10"
    first, top = match.group(1), int(match.group(2))
    last, bottom = match.group(3) or first, int(match.group(4) or top)
    left, right = build_workbook.column_number(first), build_workbook.column_number(last)
    if right < left or bottom < top:
        return f"{kind} range {text} does not start at its top left cell"
    if right > MAX_COLUMN or bottom > MAX_ROW:
        return f"{kind} range {text} is outside a worksheet"
    return Area(kind, match.group(0), top, left, bottom, right)


def entry_problems(code: str, data: Any, codes: dict[str, str]) -> tuple[Entry | None, list[str]]:
    """One sheet's entry of the matrix, and what is wrong with it."""
    where = f"{MATRIX}: sheet {code}"
    if code not in codes:
        return None, [f"{where} is not a sheet of {build_workbook.CONFIG}"]
    if codes[code] == build_workbook.TEST_ROLE:
        return None, [f"{where} is a {build_workbook.TEST_ROLE} sheet, which the template "
                      "does not hold"]
    if not isinstance(data, dict) or set(data) - {"protection", *KINDS}:
        return None, [f"{where} needs an object with protection, input and output only"]
    options = data.get("protection", {})
    problems = []
    if not isinstance(options, dict) or any(
            name not in OPTIONS or not isinstance(value, bool) for name, value in options.items()):
        problems.append(f"{where}: protection needs true or false for options among "
                        f"{', '.join(OPTIONS)}")
        options = {}
    areas: list[Area] = []
    for kind in KINDS:
        texts = data.get(kind, [])
        for text in texts if isinstance(texts, list) else [texts]:
            parsed = parse_area(kind, text)
            if isinstance(parsed, str):
                problems.append(f"{where}: {parsed}")
            else:
                areas.append(parsed)
    for first in (area for area in areas if area.kind == "input"):
        problems += [f"{where}: input range {first.text} overlaps output range {second.text}"
                     for second in areas if second.kind == "output" and overlaps(first, second)]
    return Entry(code, {**OPTIONS, **options}, tuple(areas)), problems


def overlaps(first: Area, second: Area) -> bool:
    return first.left <= second.right and second.left <= first.right and \
        first.top <= second.bottom and second.top <= first.bottom


def load_matrix(root: Path, roles: dict[str, str]) -> tuple[list[Entry], list[str]]:
    """The matrix's entries by sheet, and what is wrong with the file."""
    try:
        data = json.loads((root / MATRIX).read_text(encoding="utf-8"))
    except (UnicodeDecodeError, ValueError) as error:
        return [], [f"{MATRIX}: not valid UTF-8 JSON ({error})"]
    sheets = data.get("sheets") if isinstance(data, dict) else None
    if not isinstance(data, dict) or data.get("schema_version") != 1 or \
            not isinstance(sheets, dict):
        return [], [f"{MATRIX}: needs schema_version 1 and a sheets object"]
    entries, problems = [], []
    for code, item in sheets.items():
        entry, found = entry_problems(code, item, roles)
        problems += found
        if entry is not None:
            entries.append(entry)
    return entries, problems


def lock_flags(styles: ElementTree.Element | None) -> list[bool]:
    """Whether each cell format of styles.xml (cellXfs) is locked."""
    if styles is None:
        return [True]

    def locked(xf: ElementTree.Element, inherited: bool) -> bool:
        protection = xf.find(f"{MAIN}protection")
        if protection is None or xf.get("applyProtection") in ("0", "false"):
            return inherited
        return flag(protection.get("locked"), True)

    bases = [locked(xf, True) for xf in styles.iterfind(f"{MAIN}cellStyleXfs/{MAIN}xf")]
    flags = []
    for xf in styles.iterfind(f"{MAIN}cellXfs/{MAIN}xf"):
        index = int(xf.get("xfId", "0"))
        flags.append(locked(xf, bases[index] if index < len(bases) else True))
    return flags or [True]


def sheet_state(part: ElementTree.Element) -> SheetState:
    element = part.find(f"{MAIN}sheetProtection")
    protection = {name: flag(None if element is None else element.get(name), default)
                  for name, default in OPTIONS.items()}
    cells: dict[int, dict[int, int]] = {}
    rows: dict[int, int] = {}
    for row in part.iterfind(f"{MAIN}sheetData/{MAIN}row"):
        number = int(row.get("r", "0"))
        if flag(row.get("customFormat"), False):
            rows[number] = int(row.get("s", "0"))
        for cell in row.iterfind(f"{MAIN}c"):
            match = build_workbook.CELL.match(cell.get("r", ""))
            if match is None:
                raise ValueError(f"a cell of row {number} has no reference")
            column = build_workbook.column_number(match.group(1))
            cells.setdefault(int(match.group(2)), {})[column] = int(cell.get("s", "0"))
    columns = tuple((int(col.get("min", "0")), int(col.get("max", "0")), int(col.get("style", "0")))
                    for col in part.iterfind(f"{MAIN}cols/{MAIN}col"))
    return SheetState(protection, cells, rows, columns)


def read_states(path: Path) -> tuple[dict[str, SheetState], list[bool]]:
    """Each template sheet's state, by tab, and the lock flag of each cell format."""
    parts = build_workbook.read_template(path)
    states = {sheet.tab: sheet_state(build_workbook.parse_part(parts, sheet.part))
              for sheet in build_workbook.template_sheets(parts)}
    styles = build_workbook.parse_part(parts, "xl/styles.xml") if "xl/styles.xml" in parts \
        else None
    return states, lock_flags(styles)


def compare(entry: Entry, state: SheetState, flags: list[bool]) -> list[str]:
    """How the template's sheet differs from its entry of the matrix."""
    where = f"{TEMPLATE}: sheet {entry.code_name}"
    problems = []
    if entry.protection["sheet"] or state.protection["sheet"]:
        problems += [f"{where}: protection option {name} is "
                     f"{str(state.protection[name]).lower()}, but the matrix says "
                     f"{str(value).lower()}"
                     for name, value in entry.protection.items() if state.protection[name] != value]
    for area in entry.areas:
        count, first = area_mismatches(state, flags, area)
        if first is not None:
            row, column = first
            state_name = "unlocked" if KINDS[area.kind] else "locked"
            problems.append(f"{where}: {count} cell(s) of {area.kind} range {area.text} "
                            f"are {state_name}, first {build_workbook.column_letters(column)}"
                            f"{row}; the matrix wants every {area.kind} cell "
                            f"{'locked' if KINDS[area.kind] else 'unlocked'}")
    return problems


def check(root: Path) -> tuple[list[str], str]:
    """The findings, and what was compared."""
    has_template, has_matrix = (root / TEMPLATE).is_file(), (root / MATRIX).is_file()
    if not has_template and not has_matrix:
        return [], f"inactive: no {TEMPLATE} and no {MATRIX}"
    if not has_matrix:
        return [f"{TEMPLATE}: has no protection matrix; add {MATRIX} as "
                "docs/methodology/WORKBOOK_CONTRACT.md describes"], "template only"
    try:
        layout = build_workbook.load_layout(root)
    except build_workbook.BuildError as error:
        return [str(error)], "matrix not read"
    sheets = {sheet.code_name: sheet for sheet in layout.sheets}
    entries, findings = load_matrix(root, {code: sheet.role for code, sheet in sheets.items()})
    if not has_template:
        return findings, f"{len(entries)} sheet(s) in the matrix; no {TEMPLATE} to compare"
    try:
        states, flags = read_states(root / TEMPLATE)
    except (build_workbook.BuildError, zipfile.BadZipFile, ValueError) as error:
        return [*findings, f"{TEMPLATE}: cannot be read ({error})"], "template not read"
    listed = {sheets[entry.code_name].tab for entry in entries}
    findings += [f"{MATRIX}: has no entry for sheet {tab} of {TEMPLATE}"
                 for tab in states if tab not in listed]
    for entry in entries:
        tab = sheets[entry.code_name].tab
        if tab not in states:
            findings.append(f"{MATRIX}: sheet {entry.code_name} ({tab}) is not in {TEMPLATE}")
        else:
            findings += compare(entry, states[tab], flags)
    return findings, f"{len(entries)} sheet(s) compared with {TEMPLATE}"


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, compared = check(args.root.resolve())
    for finding in findings:
        print(f"FAIL: {finding}")
    status = "FAIL" if findings else "PASS"
    print(f"workbook-protection: {status} ({compared})")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
