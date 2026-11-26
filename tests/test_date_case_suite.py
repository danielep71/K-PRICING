"""tests/modules/TEST_DateCases.bas is generated from the worksheet-function cases.

tools/gen_date_case_suite.py writes the suite from tests/fixtures/ and
tests/expected/. These tests fail when the committed module differs from what the
cases give, so a changed case cannot leave the suite stale, and when the counts
the evidence policy states differ from the module's.
"""

from __future__ import annotations

import json
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import gen_date_case_suite  # noqa: E402

POLICY = ROOT / ".github/excel-evidence-policy.json"


class GeneratedSuiteTest(unittest.TestCase):
    def test_the_module_matches_the_cases(self) -> None:
        self.assertEqual(gen_date_case_suite.main(["--check"]), 0)

    def test_every_direct_case_is_asserted_once(self) -> None:
        text = (ROOT / gen_date_case_suite.MODULE).read_text(encoding="ascii")
        asserted = re.findall(r'^\s*TEST_Harness\.AssertEqualText "([^"]+)"', text, re.M)
        direct = []
        for path in (ROOT / "tests/fixtures").glob("*.json"):
            fixture = json.loads(path.read_text(encoding="utf-8"))
            if fixture.get("kind") == "test-call-fixture" and fixture["context"] == "direct":
                direct.append(fixture["id"])
        self.assertEqual(sorted(asserted), sorted(direct))
        recorded = re.findall(r'^\s*TEST_Harness\.RecordTextResult "([^"]+)"', text, re.M)
        self.assertEqual(recorded, asserted)

    def test_the_policy_states_the_module_counts(self) -> None:
        text = (ROOT / gen_date_case_suite.MODULE).read_text(encoding="ascii")
        cases = int(re.search(r"EXPECTED_CASES\s+As Long = (\d+)", text).group(1))  # type: ignore[union-attr]
        assertions = int(re.search(r"EXPECTED_ASSERTIONS\s+As Long = (\d+)",  # type: ignore[union-attr]
                                   text).group(1))
        policy = json.loads(POLICY.read_text(encoding="utf-8"))
        entry = next(item for item in policy["suites"] if item["name"] == "date-cases")
        self.assertEqual((entry["expected_cases"], entry["expected_assertions"]),
                         (cases, assertions))
        self.assertEqual(entry["source"], gen_date_case_suite.MODULE.as_posix())

    def test_a_number_keeps_its_double_value(self) -> None:
        self.assertEqual(gen_date_case_suite.number("45332"), "45332#")
        self.assertEqual(gen_date_case_suite.number("45332.75"), "45332.75#")
        self.assertEqual(gen_date_case_suite.number("1E+20"), "1E+20")
        self.assertEqual(gen_date_case_suite.number("45292.999999999999"), "45293#")


if __name__ == "__main__":
    unittest.main()
