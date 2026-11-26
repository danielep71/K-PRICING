"""tests/modules/TEST_ExcelOracle.bas is the Excel oracle, generated from the cases.

tools/gen_excel_oracle_suite.py writes the suite from tests/fixtures/ and
tests/expected/. These tests fail when the committed module differs from what the
cases give, when a formula calls a worksheet function outside the allowed list
or the module names a procedure of the date layer, when a case is neither
covered nor excluded, and when the evidence policy or
docs/methodology/EXCEL_ORACLE.md states other counts than the generator.
"""

from __future__ import annotations

import json
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import gen_excel_oracle_suite  # noqa: E402

POLICY = ROOT / ".github/excel-evidence-policy.json"
DOC = ROOT / "docs/methodology/EXCEL_ORACLE.md"
FORMULA = re.compile(r'Sheet\.Evaluate\("([^"]*)"\)')
CALLED = re.compile(r"([A-Z][A-Z0-9.]*)\(")


def module_text() -> str:
    return (ROOT / gen_excel_oracle_suite.MODULE).read_text(encoding="ascii")


class ExcelOracleSuiteTest(unittest.TestCase):
    def test_the_module_matches_the_cases(self) -> None:
        self.assertEqual(gen_excel_oracle_suite.main(["--check"]), 0)

    def test_formulas_call_only_the_allowed_functions(self) -> None:
        formulas = FORMULA.findall(module_text())
        self.assertTrue(formulas)
        called = {name for text in formulas for name in CALLED.findall(text)}
        self.assertLessEqual(called, set(gen_excel_oracle_suite.ALLOWED))

    def test_a_formula_outside_the_list_is_seen(self) -> None:
        called = set(CALLED.findall("DATEVALUE(A1)+NETWORKDAYS.INTL(B1,C1)"))
        self.assertEqual(called - set(gen_excel_oracle_suite.ALLOWED),
                         {"DATEVALUE", "NETWORKDAYS.INTL"})

    def test_the_module_names_no_procedure_of_the_date_layer(self) -> None:
        text = module_text()
        self.assertNotIn("KPR_", text)
        self.assertIsNone(re.search(r"\b(CORE_|M_Ranges|M_Host|M_Register)", text))

    def test_every_case_is_covered_or_excluded_once(self) -> None:
        coverage = gen_excel_oracle_suite.coverage(ROOT)
        covered = [case["id"] for case in coverage["covered"]]
        excluded = [case["id"] for case in coverage["excluded"]]
        every = []
        for path in (ROOT / "tests/fixtures").glob("*.json"):
            fixture = json.loads(path.read_text(encoding="utf-8"))
            if fixture.get("kind") == "test-call-fixture":
                every.append(fixture["id"])
        self.assertEqual(sorted(covered + excluded), sorted(every))
        asserted = re.findall(r'^\s*TEST_Harness\.AssertEqualText "([^"]+)"', module_text(),
                              re.M)
        self.assertEqual(sorted(asserted), sorted(covered))
        reasons = set(gen_excel_oracle_suite.REASONS)
        self.assertTrue(all(case["reason"] in reasons for case in coverage["excluded"]))

    def test_the_documentation_states_the_coverage(self) -> None:
        text = DOC.read_text(encoding="utf-8")
        start = "<!-- excel-oracle-coverage:start -->\n"
        stated = text[text.index(start) + len(start):text.index("<!-- excel-oracle-coverage:end")]
        self.assertEqual(stated, gen_excel_oracle_suite.coverage_tables(ROOT))

    def test_the_policy_states_the_module_counts(self) -> None:
        text = module_text()
        cases = int(re.search(r"EXPECTED_CASES\s+As Long = (\d+)", text).group(1))  # type: ignore[union-attr]
        assertions = int(re.search(r"EXPECTED_ASSERTIONS\s+As Long = (\d+)",  # type: ignore[union-attr]
                                   text).group(1))
        policy = json.loads(POLICY.read_text(encoding="utf-8"))
        entry = next(item for item in policy["suites"] if item["name"] == "excel-oracle")
        self.assertEqual((entry["expected_cases"], entry["expected_assertions"]),
                         (cases, assertions))
        self.assertEqual(entry["source"], gen_excel_oracle_suite.MODULE.as_posix())


if __name__ == "__main__":
    unittest.main()
