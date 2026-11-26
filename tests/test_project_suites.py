"""Every project suite of the evidence policy has the shape the certification reads.

tools/check_excel_run.py compares each run log with the suite's source: the cases
its runner calls and the literal assertion names in their bodies
(docs/methodology/TEST_HARNESS.md, Modules). An assertion made in a helper, or
named at run time, is not read, and the certification of the run fails. These
tests read every project suite as the certification does and fail when its cases
or assertions differ from what its constants and the policy state.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_excel_run  # noqa: E402


class ProjectSuiteShapeTest(unittest.TestCase):
    def test_every_suite_reads_as_its_counts_state(self) -> None:
        policy = check_excel_run.working_policy(ROOT)
        self.assertTrue(policy.project_suites)
        for item in policy.project_suites:
            with self.subTest(suite=item.name):
                suite = check_excel_run.read_suite((ROOT / item.source).read_bytes(),
                                                   policy.office, item.runner)
                self.assertEqual((len(suite.cases), len(suite.assertions)),
                                 (suite.expected_cases, suite.expected_assertions))
                self.assertEqual((item.expected_cases, item.expected_assertions),
                                 (suite.expected_cases, suite.expected_assertions))
                self.assertEqual(len(suite.assertions), len(set(suite.assertions)))


if __name__ == "__main__":
    unittest.main()
