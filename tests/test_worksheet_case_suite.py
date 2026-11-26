"""tests/modules/TEST_WorksheetCases.bas is generated from the worksheet callers' cases.

tools/gen_worksheet_case_suite.py writes the suite from tests/fixtures/ and
tests/expected/. These tests fail when the committed module differs from what the
cases give, when a worksheet caller's case is not asserted and recorded exactly
once, and when the counts the evidence policy states differ from the module's.
"""

from __future__ import annotations

import json
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import gen_worksheet_case_suite  # noqa: E402

POLICY = ROOT / ".github/excel-evidence-policy.json"


def module_text() -> str:
    return (ROOT / gen_worksheet_case_suite.MODULE).read_text(encoding="ascii")


class GeneratedWorksheetSuiteTest(unittest.TestCase):
    def test_the_module_matches_the_cases(self) -> None:
        self.assertEqual(gen_worksheet_case_suite.main(["--check"]), 0)

    def test_every_worksheet_case_is_asserted_and_recorded_once(self) -> None:
        text = module_text()
        asserted = re.findall(r'^\s*TEST_Harness\.AssertEqualText "([^"]+)"', text, re.M)
        callers = []
        for path in (ROOT / "tests/fixtures").glob("*.json"):
            fixture = json.loads(path.read_text(encoding="utf-8"))
            if fixture.get("kind") == "test-call-fixture" and fixture["context"] != "direct":
                callers.append(fixture["id"])
        self.assertTrue(callers)
        self.assertEqual(sorted(asserted), sorted(callers))
        recorded = re.findall(r'^\s*TEST_Harness\.RecordTextResult "([^"]+)"', text, re.M)
        self.assertEqual(recorded, asserted)

    def test_each_case_opens_its_date_system_and_asserts_the_close(self) -> None:
        text = module_text()
        procedures = re.findall(r"^Private Sub (Case\w+)\(\)\r?\n(.*?)^End Sub", text,
                                re.M | re.S)
        self.assertTrue(procedures)
        for name, body in procedures:
            with self.subTest(case=name):
                opened = re.findall(r"OpenScratch\((True|False)\)", body)
                self.assertEqual(opened, ["True" if "1904" in name else "False"])
                self.assertEqual(body.count('TEST_Harness.AssertTrue "scratch-closed"'), 1)

    def test_the_policy_states_the_module_counts(self) -> None:
        text = module_text()
        cases = int(re.search(r"EXPECTED_CASES\s+As Long = (\d+)", text).group(1))  # type: ignore[union-attr]
        assertions = int(re.search(r"EXPECTED_ASSERTIONS\s+As Long = (\d+)",  # type: ignore[union-attr]
                                   text).group(1))
        policy = json.loads(POLICY.read_text(encoding="utf-8"))
        entry = next(item for item in policy["suites"] if item["name"] == "worksheet-cases")
        self.assertEqual((entry["expected_cases"], entry["expected_assertions"]),
                         (cases, assertions))
        self.assertEqual(entry["source"], gen_worksheet_case_suite.MODULE.as_posix())
        self.assertEqual(entry["results"], "worksheet-cases-results.json")


if __name__ == "__main__":
    unittest.main()
