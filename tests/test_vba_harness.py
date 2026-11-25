"""Static consistency of the VBA harness and its example suite (tests/modules/).

The harness runs only in Excel. These tests check, without Excel, what its
source can show: that the example suite's expected counts equal the cases and
assertions it contains, that every case runs once with a well-formed unique
identifier, that the documented passing RESULT line matches the counts, and
that every record the harness writes is documented. The relation suite is
checked the same way, and each of its cases must pair the relation with its
corruption check; its tests are skipped in a repository that deleted it.
"""

from __future__ import annotations

import io
import json
import re
import tempfile
import unittest
from pathlib import Path
from typing import ClassVar

ROOT = Path(__file__).resolve().parents[1]
EXAMPLES = ROOT / "tests/modules/TEST_Examples.bas"
RELATIONS = ROOT / "tests/modules/TEST_Relations.bas"
HARNESS = ROOT / "tests/modules/TEST_Harness.bas"
GUIDE = ROOT / "docs/methodology/TEST_HARNESS.md"
REGISTRY = ROOT / "tests/accuracy/contracts.json"
BUILD_CONFIG = ROOT / "tools/devbuild/workbook.json"
CASE_ID = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
# A RecordResult statement, not a comment that names it.
RECORD = re.compile(r"^\s*TEST_Harness\.RecordResult ", re.M)


def source(path: Path) -> str:
    return path.read_text(encoding="ascii").replace("\r\n", "\n")


def constant(text: str, name: str) -> int:
    match = re.search(rf"^\s*Private Const {name}\s+As Long = (\d+)\b", text, re.M)
    assert match, name
    return int(match.group(1))


def procedures(text: str) -> dict[str, str]:
    """Body of every Sub and Function, by name."""
    pattern = re.compile(r"^(?:Public|Private) (?:Sub|Function) (\w+)\(.*?^End (?:Sub|Function)$",
                         re.M | re.S)
    return {match.group(1): match.group(0) for match in pattern.finditer(text)}


class ExampleSuiteTest(unittest.TestCase):
    text = source(EXAMPLES)
    bodies = procedures(text)
    cases: ClassVar[dict[str, str]] = {name: body for name, body in bodies.items() if name.startswith("Case")}

    def test_runner_calls_every_case_once_in_order(self) -> None:
        runner = self.bodies["RunSuite"]
        called = re.findall(r"^\s+(Case\w+)$", runner, re.M)
        self.assertEqual(called, list(self.cases))

    def test_each_case_begins_once_with_a_unique_identifier(self) -> None:
        identifiers = []
        for name, body in self.cases.items():
            begun = re.findall(r'TEST_Harness\.BeginCase "([^"]+)"', body)
            self.assertEqual(len(begun), 1, name)
            self.assertRegex(begun[0], CASE_ID)
            self.assertIn("TEST_Harness.RecordCaseError", body, name)
            identifiers += begun
        self.assertEqual(len(identifiers), len(set(identifiers)))

    def test_expected_counts_match_the_source(self) -> None:
        assertions = sum(len(re.findall(r"TEST_Harness\.Assert\w+ ", body))
                         for body in self.cases.values())
        self.assertEqual(constant(self.text, "EXPECTED_CASES"), len(self.cases))
        self.assertEqual(constant(self.text, "EXPECTED_ASSERTIONS"), assertions)

    def test_documented_assertion_counts_per_case(self) -> None:
        for name, body in self.cases.items():
            stated = re.search(r"(\d+) assertions?\.", body)
            self.assertIsNotNone(stated, name)
            assert stated is not None
            self.assertEqual(int(stated.group(1)),
                             len(re.findall(r"TEST_Harness\.Assert\w+ ", body)), name)

    def test_documented_passing_line(self) -> None:
        cases = constant(self.text, "EXPECTED_CASES")
        assertions = constant(self.text, "EXPECTED_ASSERTIONS")
        line = ("RESULT=PASS; suite=examples; operation=PASS; completeness=COMPLETE; "
                f"cases={cases}; assertions={assertions}; failures=0; cleanup=PASS; "
                "recovery=NONE")
        self.assertIn(line, (ROOT / "INSTALLATION.md").read_text(encoding="utf-8"))
        skipped = self.cases["CaseCleanupOnlyFailure"]
        remaining = assertions - len(re.findall(r"TEST_Harness\.Assert\w+ ", skipped))
        self.assertIn(f"`cases={cases - 1}`, `assertions={remaining}`",
                      (ROOT / "INSTALLATION.md").read_text(encoding="utf-8"))


    def results_cases(self) -> dict[str, str]:
        """The cases that record results, found by their RecordResult calls."""
        return {name: body for name, body in self.cases.items() if RECORD.search(body)}

    def recorded(self) -> list[tuple[str, str | None, str | None, str | None, str | None]]:
        """Case, variant, quantity, dataset and record of every RecordResult call."""
        strings = dict(re.findall(r'^\s*Private Const (\w+)\s+As String = "([^"]*)"', self.text,
                                  re.M))

        def value(token: str) -> str | None:
            if token == "vbNullString":
                return None
            return token.strip('"') if token.startswith('"') else strings[token]
        separator = r",\s*(?:_\s+)?"
        calls = "".join(self.results_cases().values())
        recorded = [(case, value(variant), value(quantity), value(dataset), value(record))
                    for case, variant, quantity, dataset, record in re.findall(
                        r'TEST_Harness\.RecordResult "([^"]+)"' + separator
                        + r'(vbNullString|"[^"]+")' + separator + r'(\w+|"[^"]+")' + separator
                        + r'(vbNullString|"[^"]*")' + separator + r'(vbNullString|"[^"]*")'
                        + separator + r"\w+", calls)]
        # Every call is read: none is skipped for a layout the pattern does not know.
        self.assertEqual(len(recorded), len(RECORD.findall(calls)), recorded)
        return recorded

    def test_the_results_case_records_every_output_of_the_contracts(self) -> None:
        recorded = self.recorded()
        registry = json.loads(REGISTRY.read_text(encoding="utf-8")) if REGISTRY.is_file() \
            else {"contracts": []}
        expected = []
        for contract in registry["contracts"]:
            for case in [*contract["calibration_cases"], *contract["holdout_cases"]]:
                data = json.loads((ROOT / "tests/expected" / f"{case}.json")
                                  .read_text(encoding="utf-8"))
                groups = [(None, data["outputs"])] if "outputs" in data \
                    else sorted(data["variants"].items())
                expected += [(case, variant, output["quantity"],
                              *((output["scope"]["dataset"], output["scope"]["record"])
                                if output["scope"] else (None, None)))
                             for variant, outputs in groups for output in outputs
                             if output["quantity"] in contract["quantities"]]
        self.assertEqual(sorted(recorded, key=str), sorted(expected, key=str))
        self.assertEqual(len(recorded), len(set(recorded)))

    def test_the_results_case_reads_files_that_exist(self) -> None:
        paths = [path for body in self.results_cases().values()
                 for path in re.findall(r'FixturePath\(\s*(?:_\s*)?"([^"]+)"\)', body)]
        for path in paths:
            with self.subTest(path=path):
                self.assertTrue((ROOT / "tests/fixtures" / path).is_file())
        # Each recorded case is computed from its own fixture file or folder.
        for case in {item[0] for item in self.recorded()}:
            with self.subTest(case=case):
                self.assertTrue(any(path == f"{case}.json" or path.startswith(f"{case}/")
                                    for path in paths), paths)

    def test_the_certification_saves_the_results_of_the_normal_run(self) -> None:
        body = self.bodies["CertifyExamples"]
        self.assertRegex(body, r'CertifyMode "NORMAL"[^\n]*\n\s+TEST_Harness\.SaveResults ')


class RelationSuiteTest(unittest.TestCase):
    """The relation suite, read when its tests run: the project may delete the
    module, and its tests are then skipped."""

    path: ClassVar[Path] = RELATIONS
    text: ClassVar[str]
    bodies: ClassVar[dict[str, str]]
    cases: ClassVar[dict[str, str]]

    @classmethod
    def setUpClass(cls) -> None:
        if not cls.path.is_file():
            raise unittest.SkipTest(f"{cls.path.name} is not in this repository")
        cls.text = source(cls.path)
        cls.bodies = procedures(cls.text)
        cls.cases = {name: body for name, body in cls.bodies.items() if name.startswith("Case")}

    @staticmethod
    def assertions(body: str) -> list[str]:
        return re.findall(r'TEST_Harness\.Assert\w+ "([^"]+)"', body)

    def test_runner_calls_every_case_once_in_order(self) -> None:
        called = re.findall(r"^\s+(Case\w+)$", self.bodies["RunRelationTests"], re.M)
        self.assertEqual(called, list(self.cases))
        self.assertTrue(self.cases)

    def test_each_case_begins_once_with_a_unique_identifier(self) -> None:
        identifiers = []
        for name, body in self.cases.items():
            begun = re.findall(r'TEST_Harness\.BeginCase "([^"]+)"', body)
            self.assertEqual(len(begun), 1, name)
            self.assertRegex(begun[0], CASE_ID)
            self.assertIn("TEST_Harness.RecordCaseError", body, name)
            identifiers += begun
        self.assertEqual(len(identifiers), len(set(identifiers)))

    def test_every_relation_has_its_corruption_check(self) -> None:
        for name, body in self.cases.items():
            with self.subTest(case=name):
                self.assertEqual(sorted(self.assertions(body)), ["detects-corruption", "holds"])
                self.assertRegex(body, r'AssertTrue "detects-corruption", _\s+Not ')

    def test_expected_counts_match_the_source(self) -> None:
        assertions = sum(len(self.assertions(body)) for body in self.cases.values())
        cases = constant(self.text, "EXPECTED_CASES")
        self.assertEqual(cases, len(self.cases))
        self.assertEqual(constant(self.text, "EXPECTED_ASSERTIONS"), assertions)
        for name, body in self.cases.items():
            stated = re.search(r"(\d+) assertions?\.", body)
            self.assertIsNotNone(stated, name)
            assert stated is not None
            self.assertEqual(int(stated.group(1)), len(self.assertions(body)), name)
        passing = re.sub(r"\n'\s+", " ", self.bodies["RunRelationTests"])
        self.assertIn(f"cases={cases}; assertions={assertions}; failures=0", passing)


class MissingRelationSuiteTest(unittest.TestCase):
    def test_a_deleted_relation_suite_skips_its_tests(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            missing = type("Missing", (RelationSuiteTest,),
                           {"path": Path(directory) / "TEST_Relations.bas"})
            suite = unittest.TestLoader().loadTestsFromTestCase(missing)
            result = unittest.TextTestRunner(stream=io.StringIO(), verbosity=0).run(suite)
        self.assertEqual((result.testsRun, len(result.errors), len(result.failures)),
                         (0, 0, 0), result.errors)
        self.assertEqual(len(result.skipped), 1)  # the class, skipped once in setUpClass
        self.assertGreater(suite.countTestCases(), 0)


class HarnessTest(unittest.TestCase):
    text = source(HARNESS)

    def test_the_stamp_name_is_the_one_the_build_writes(self) -> None:
        match = re.search(r'Const STAMP_COMMIT_NAME\s+As String = "([^"]+)"', self.text)
        self.assertIsNotNone(match)
        assert match is not None
        layout = json.loads(BUILD_CONFIG.read_text(encoding="utf-8"))
        self.assertEqual(match.group(1), layout["stamp"]["commit"])

    def test_a_result_scope_is_the_object_of_the_expected_format(self) -> None:
        # The accuracy check matches a scope only as {"dataset": ..., "record": ...}.
        bodies = procedures(self.text)
        self.assertIn('""scope"": " & ScopeJson(Dataset, Record)', bodies["RecordResult"])
        scope = re.sub(r" _\n\s+", " ", bodies["ScopeJson"])  # join continued lines
        self.assertIn('"{""dataset"": " & JsonText(Dataset) & ", ""record"": " & '
                      'JsonText(Record) & "}"', scope)
        self.assertIn('ScopeJson = "null"', scope)

    def test_results_are_cleared_with_the_log(self) -> None:
        bodies = procedures(self.text)
        for name in ("BeginRun", "ResetTests"):
            with self.subTest(procedure=name):
                self.assertIn("Set mResults = New Collection", bodies[name])
        self.assertIn('RequireActiveRun "RecordResult"', bodies["RecordResult"])

    def test_every_documented_record_is_written(self) -> None:
        guide = GUIDE.read_text(encoding="utf-8")
        cells = [line.split("|")[1] for line in guide.splitlines() if line.startswith("| `")]
        documented = {key for cell in cells for key in re.findall(r"`([A-Z_]+)=", cell)}
        written = set(re.findall(r'"([A-Z_]+)=', self.text))
        self.assertEqual(documented, written - {"RESET"})
        self.assertIn("RESET=DONE", guide)

    def test_a_refused_start_writes_its_result_last(self) -> None:
        # The guide: RESULT is always the last record of a run, after RECOVERY.
        body = procedures(self.text)["BeginRun"]
        refused = body[body.index("If mRecoveryRequired Then"):body.index("Err.Raise ERR_TEST_DIRTY_START")]
        self.assertLess(refused.index('Say "RECOVERY=REQUIRED'), refused.index("Say mLastResult"))
        self.assertEqual(refused.count("Say "), 2)

    def test_certification_raises_only_after_writing(self) -> None:
        body = procedures(self.text)["EndCertification"]
        self.assertLess(body.index("WriteRecords Path, mCertLog"),
                        body.index("Err.Raise ERR_TEST_FAILED"))
        self.assertLess(body.index("mCertActive = False"), body.index("WriteRecords"))

    def test_error_values_are_not_compared_with_operators(self) -> None:
        # Comparing an error Variant with "=" or "<>" raises Type mismatch.
        self.assertNotRegex(self.text, r"\bActual\s*(?:=|<>)\s*CVErr\(")
        self.assertIn("CStr(Actual) = CStr(CVErr(ExpectedCode))", self.text)

    def test_documented_failure_ids_are_written(self) -> None:
        for failure in ("harness.case-id", "harness.no-case", "harness.completeness",
                        "runner.unexpected-error", ".unexpected-error"):
            with self.subTest(failure=failure):
                self.assertIn(f'"{failure}"', self.text)


if __name__ == "__main__":
    unittest.main()
