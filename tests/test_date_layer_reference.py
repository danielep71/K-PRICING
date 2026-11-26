"""Tests of tests/derivations/date_layer_reference.py on single calls.

The reference model derives the expected values of the worksheet-function cases;
tools/check_derivations.py compares it with every case. These tests pin rules
that no case reaches yet.
"""

from __future__ import annotations

import importlib.util
import sys
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any, cast

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "tests/derivations/date_layer_reference.py"


def load_derive() -> Callable[[dict[str, object]], list[dict[str, object]]]:
    spec = importlib.util.spec_from_file_location("date_layer_reference", SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    # Its dataclasses look their module up while it loads.
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return cast(Callable[[dict[str, object]], list[dict[str, object]]], module.derive)


DERIVE = load_derive()


def add_days(days: str, date: str = "date:2024-01-01") -> list[dict[str, Any]]:
    """The outputs of KPR_Dates_AddDays(date, days) called from VBA."""
    return DERIVE({"function": "KPR_Dates_AddDays", "context": "direct",
                   "arguments": [date, days]})


class WholeNumberTest(unittest.TestCase):
    def test_a_double_is_checked_in_its_binary_value(self) -> None:
        # 3.000000001 is about 1.0000000827E-9 above 3 as a Double: a fraction.
        self.assertEqual([item["text"] for item in add_days("num:3.000000001")],
                         ["err:#VALUE!", "INTEGER_FRACTION"])

    def test_residue_within_the_tolerance_reads_as_the_whole_number(self) -> None:
        self.assertEqual([item["text"] for item in add_days("num:3.0000000001")],
                         ["date:2024-01-04"])
        self.assertEqual([item["text"] for item in add_days("num:2.9999999999")],
                         ["date:2024-01-04"])

    def test_a_long_is_exact(self) -> None:
        self.assertEqual([item["text"] for item in add_days("lng:3")], ["date:2024-01-04"])


class DateSerialTest(unittest.TestCase):
    def test_a_double_serial_is_floored_in_its_binary_value(self) -> None:
        # Both literals round to the next whole number as a Double.
        self.assertEqual([item["text"] for item in add_days("lng:0", "num:45292.999999999999")],
                         ["date:2024-01-02"])
        self.assertEqual([item["text"] for item in add_days("lng:0", "num:60.999999999999999")],
                         ["date:1900-03-01"])

    def test_a_fraction_is_removed(self) -> None:
        self.assertEqual([item["text"] for item in add_days("lng:0", "num:45292.75")],
                         ["date:2024-01-01"])
        self.assertEqual([item["text"] for item in add_days("lng:0", "num:60.5")],
                         ["err:#NUM!", "DATE_WINDOW"])


def day_of_week(*arguments: str) -> list[str]:
    """The texts of the outputs of KPR_Dates_DayOfWeek(arguments) called from VBA."""
    return [str(item["text"]) for item in DERIVE(
        {"function": "KPR_Dates_DayOfWeek", "context": "direct", "arguments": list(arguments)})]


class ShapeTest(unittest.TestCase):
    def test_an_empty_or_nested_control_has_no_supported_shape(self) -> None:
        self.assertEqual(day_of_week("date:2024-01-01", "arr1:[]"),
                         ["err:#VALUE!", "SHAPE_UNSUPPORTED"])
        self.assertEqual(day_of_week("date:2024-01-01", "arr1:[arr1:[bool:TRUE]]"),
                         ["err:#VALUE!", "SHAPE_UNSUPPORTED"])
        self.assertEqual(day_of_week("date:2024-01-01", "arr1:[bool:TRUE;bool:FALSE]"),
                         ["err:#VALUE!", "CONTROL_NOT_SCALAR"])

    def test_a_nested_value_refuses_the_call_after_the_capacity_check(self) -> None:
        self.assertEqual(day_of_week("arr1:[arr1:[date:2024-01-01]]"),
                         ["err:#VALUE!", "SHAPE_UNSUPPORTED"])
        self.assertEqual(day_of_week("fill2:2x2[arr1:[date:2024-01-01]]"),
                         ["err:#VALUE!", "SHAPE_UNSUPPORTED"])
        self.assertEqual(day_of_week("fill2:1x100001[arr1:[date:2024-01-01]]"),
                         ["err:#NUM!", "CAPACITY_EXCEEDED"])


if __name__ == "__main__":
    unittest.main()
