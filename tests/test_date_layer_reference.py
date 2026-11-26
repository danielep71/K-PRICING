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


def add_days(days: str) -> list[dict[str, Any]]:
    """The outputs of KPR_Dates_AddDays(2024-01-01, days) called from VBA."""
    return DERIVE({"function": "KPR_Dates_AddDays", "context": "direct",
                   "arguments": ["date:2024-01-01", days]})


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


if __name__ == "__main__":
    unittest.main()
