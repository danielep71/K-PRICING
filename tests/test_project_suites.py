"""Every project suite of the evidence policy has the shape the certification reads.

tools/check_excel_run.py compares each run log with the suite's source: the cases
its runner calls and the literal assertion names in their bodies
(docs/methodology/TEST_HARNESS.md, Modules). An assertion made in a helper, or
named at run time, is not read, and the certification of the run fails. These
tests read every project suite as the certification does and fail when its cases
or assertions differ from what its constants and the policy state, and when
tests/modules/TEST_Suites.bas, which the Run tests button of the menus calls,
does not run every project suite with a runner that takes no argument.
"""

from __future__ import annotations

import re
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


class RunTestsButtonTest(unittest.TestCase):
    def test_test_suites_runs_every_project_suite_in_policy_order(self) -> None:
        policy = check_excel_run.working_policy(ROOT)
        source = (ROOT / "tests/modules/TEST_Suites.bas").read_text(encoding="ascii")
        listed = re.findall(r'"([\w-]+)=(\w+)\.(\w+);?"', source)
        expected = []
        for item in policy.project_suites:
            module = Path(item.source).stem
            text = (ROOT / item.source).read_text(encoding="ascii")
            runners = re.findall(r"^Public Sub (Run\w+)\(\)", text, re.M)
            with self.subTest(suite=item.name):
                self.assertEqual(len(runners), 1, f"{item.source}: {runners}")
            expected.append((item.name, module, runners[0] if runners else ""))
        self.assertEqual(listed, expected)


if __name__ == "__main__":
    unittest.main()
