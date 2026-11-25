"""Tests that src/core agrees with the date-layer methodology it implements.

CORE_Err and CORE_Parse restate parts of docs/methodology/DATA_CONTRACT.md,
Function arguments, and of the common rules of CALCULATION_CONTRACTS.md: the
registry of conditions with their errors, the one raised error and the date
window. Nothing else compares them, because the VBA runs only in Excel, so
these tests read the VBA source and the guides and fail when they differ.
"""

from __future__ import annotations

import datetime
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA_CONTRACT = ROOT / "docs/methodology/DATA_CONTRACT.md"
ERR = ROOT / "src/core/CORE_Err.bas"
PARSE = ROOT / "src/core/CORE_Parse.bas"

ERROR_NAMES = {"`#VALUE!`": "ErrValue", "`#NUM!`": "ErrNum", "`#N/A`": "ErrNA"}
PROPAGATED = "The error passed in"


def read(path: Path) -> str:
    return path.read_text(encoding="utf-8").replace("\r\n", "\n")


def identifier_of(member: str) -> str:
    """DATE_TEXT_FORMAT for cdDateTextFormat: the member's words, upper-cased."""
    return "_".join(word.upper() for word in re.findall(r"[A-Z][a-z]+\d*", member[2:]))


def registry() -> dict[str, tuple[str, str]]:
    """The data contract's Conditions table: identifier to (error, level)."""
    text = read(DATA_CONTRACT)
    section = text[text.index('<a id="conditions"></a>'):text.index('<a id="internal-failures"></a>')]
    rows = re.findall(r"^\| `([A-Z0-9_]+)` \| .*? \| (`#[A-Z/!]+`|The error passed in) \| "
                      r"(Element|Call) \|$", section, re.M)
    return {identifier: (error, level) for identifier, error, level in rows}


def enum_members() -> dict[str, int]:
    found = re.search(r"Public Enum KPRCondition\n(.*?)\n\s*End Enum", read(ERR), re.S)
    assert found is not None, "CORE_Err declares no KPRCondition enum"
    body = found.group(1)
    return {name: int(number)
            for name, number in re.findall(r"^\s*(cd\w+)\s*=\s*(\d+)", body, re.M)}


class ConditionRegistryTest(unittest.TestCase):
    def test_the_enum_mirrors_the_registry(self) -> None:
        members = enum_members()
        self.assertEqual(len(registry()), 31)
        # cdNone is the success sentinel: success has no identifier.
        self.assertEqual({identifier_of(name) for name in members if name != "cdNone"},
                         set(registry()))
        self.assertEqual(members["cdNone"], 0)

    def test_enum_numbers_are_distinct(self) -> None:
        numbers = list(enum_members().values())
        self.assertEqual(len(numbers), len(set(numbers)))

    def test_condition_name_returns_the_identifier_of_each_member(self) -> None:
        text = read(ERR)
        body = text[text.index("Public Function ConditionName"):]
        found = dict(re.findall(r'Case (cd\w+)\n\s+ConditionName = "([A-Z0-9_]+)"', body))
        self.assertEqual(found, {name: identifier_of(name)
                                 for name in enum_members() if name != "cdNone"})

    def test_each_condition_maps_to_the_error_of_the_registry(self) -> None:
        text = read(ERR)
        body = text[text.index("Public Function ErrForCondition"):text.index("Public Function ConditionName")]
        mapped: dict[str, str] = {}
        for cases, call in re.findall(r"Case ((?:cd\w+,? _\n\s*)+cd\w+|cd\w+)\n\s+ErrForCondition = (Err\w+)\(\)",
                                      body):
            for member in re.findall(r"cd\w+", cases):
                mapped[identifier_of(member)] = call
        expected = {identifier: ERROR_NAMES[error]
                    for identifier, (error, _) in registry().items() if error != PROPAGATED}
        self.assertEqual(mapped, expected)
        # The propagation conditions are the registry's only unmapped identifiers.
        self.assertEqual({i for i, (error, _) in registry().items() if error == PROPAGATED},
                         {"INPUT_ERROR_PROPAGATED", "CONTROL_ERROR_PROPAGATED"})

    def test_native_error_numbers(self) -> None:
        text = read(ERR)
        for name, number in (("ERROR_VALUE_NUMBER", 2015), ("ERROR_NUM_NUMBER", 2036),
                             ("ERROR_NA_NUMBER", 2042)):
            self.assertRegex(text, rf"Private Const {name}\s+As Long = {number}\b")


class InternalFailureTest(unittest.TestCase):
    """The data contract's Internal failures table is the source of the raised error."""

    def table_row(self) -> list[str]:
        text = read(DATA_CONTRACT)
        row = re.search(r"^\| `ERR_UNMAPPED_CONDITION` \| .*$", text, re.M)
        self.assertIsNotNone(row)
        assert row is not None
        return [cell.strip() for cell in row.group(0).strip("|").split("|")]

    def test_number_source_and_message_match_the_code(self) -> None:
        _, number, source, message = self.table_row()
        code = read(ERR)
        offset = int(re.fullmatch(r"`vbObjectError \+ (\d+)`", number).group(1))  # type: ignore[union-attr]
        # vbObjectError is &H80040000; the constant is written as a literal because
        # LibreOffice does not compile vbObjectError in a constant expression.
        self.assertIn(f"ERR_UNMAPPED_CONDITION As Long = &H{0x80040000 + offset:X}", code)
        self.assertEqual(source, "`CORE_Err.ErrForCondition`")
        self.assertIn('Const PROC_NAME     As String = "CORE_Err.ErrForCondition"', code)
        described = re.search(r'UNMAPPED_DESCRIPTION\s+As String = _\n\s+"([^"]+)"', code)
        self.assertIsNotNone(described)
        assert described is not None
        self.assertEqual(f"`{described.group(1)}`", message)

    def test_the_number_is_in_the_projects_range(self) -> None:
        text = read(ROOT / "docs/VBA_HOUSE_STYLE.md").replace("\n", " ")
        bounds = re.search(r"`vbObjectError \+ (\d+)` through `vbObjectError \+ (\d+)`", text)
        self.assertIsNotNone(bounds)
        assert bounds is not None
        low, high = int(bounds.group(1)), int(bounds.group(2))
        offset = int(re.search(r"`vbObjectError \+ (\d+)`", self.table_row()[1]).group(1))  # type: ignore[union-attr]
        self.assertTrue(low <= offset <= high, (low, offset, high))


class WindowTest(unittest.TestCase):
    def test_the_window_constants_are_the_excel_serials_of_its_edges(self) -> None:
        epoch = datetime.date(1899, 12, 30)
        first, last = (datetime.date(1900, 3, 1) - epoch).days, (datetime.date(9999, 12, 31) - epoch).days
        text = read(PARSE)
        self.assertRegex(text, rf"WINDOW_MIN_SERIAL\s+As Double = {first}#")
        self.assertRegex(text, rf"WINDOW_MAX_SERIAL\s+As Double = {last}#")
        self.assertRegex(text, r"WINDOW_MIN_YEAR\s+As Long = 1900\b")
        self.assertRegex(text, r"WINDOW_MAX_YEAR\s+As Long = 9999\b")

    def test_the_guides_state_the_same_window(self) -> None:
        for name in ("DATA_CONTRACT.md", "CALCULATION_CONTRACTS.md"):
            self.assertIn("1900-03-01", read(ROOT / "docs/methodology" / name))
            self.assertIn("9999-12-31", read(ROOT / "docs/methodology" / name))

    def test_the_integer_tolerance_is_one_billionth(self) -> None:
        self.assertRegex(read(PARSE), r"INTEGER_TOLERANCE As Double = 1# / 1000000000#")
        self.assertIn("`1E-9`", read(DATA_CONTRACT))


if __name__ == "__main__":
    unittest.main()
