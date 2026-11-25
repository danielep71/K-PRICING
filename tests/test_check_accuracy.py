"""Valid and invalid accuracy contracts and results for tools/check_accuracy.py.

Each test builds a throw-away repository from this one's methodology guides
and test cases, adds a non-illustrative calculation CALC-02 with synthetic
cases whose values are marked independent and an illustrative calculation
CALC-03 with synthetic cases, and writes a registry. The values are fixtures
for the check only: no contract here guarantees anything, and none depends on
the cases or contracts this repository holds.
"""

from __future__ import annotations

import contextlib
import copy
import hashlib
import io
import json
import shutil
import sys
import tempfile
import unittest
from collections.abc import Iterator
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import check_accuracy  # noqa: E402
from check_test_cases import Report  # noqa: E402
from gitfixture import git, init  # noqa: E402

QUANTITY = "total_exposure_amount"
CANDIDATE = "a" * 40
ILLUSTRATIVE = {
    "id": "ACC-01", "calculation": "CALC-03", "regime": "illustrative",
    "quantities": [QUANTITY], "metric": "absolute",
    "threshold": 0.000001, "provenance": "Synthetic fixture", "status": "draft",
    "decision": None, "calibration_cases": ["demo-illustrative-calibration"],
    "holdout_cases": ["demo-illustrative-holdout"],
}
FROZEN = {
    "id": "ACC-02", "calculation": "CALC-02", "regime": "CRR3", "quantities": [QUANTITY],
    "metric": "relative", "threshold": 0.0001, "provenance": "Synthetic fixture",
    "status": "frozen", "decision": "#1", "calibration_cases": ["demo-calibration"],
    "holdout_cases": ["demo-holdout-1", "demo-holdout-2"],
}


def registry(*contracts: dict[str, Any], minimum: int = 2) -> dict[str, Any]:
    return {"schema_version": 1, "kind": "accuracy-contracts", "min_holdout_cases": minimum,
            "contracts": [copy.deepcopy(item) for item in contracts]}


def demo_case(root: Path, case: str, value: float, reference: str = "independent",
              calculation: str = "CALC-02") -> None:
    fixture = {"schema_version": 1, "kind": "test-fixture", "id": case, "synthetic": True,
               "description": "Synthetic", "category": "nominal", "calculation": calculation,
               "reporting_date": "2026-01-02", "reporting_currency_code": "EUR",
               "datasets": {"DS-01": [{"id": "P1", "inputs": {"currency_code": "EUR"}}]}}
    expected = {"schema_version": 1, "kind": "test-expected", "fixture": case, "outputs": [
        {"quantity": QUANTITY, "scope": None, "rule": f"{calculation}.1", "value": value,
         "unit": "EUR", "tolerance": {"absolute": 0.01, "relative": 0},
         "reference": {"class": reference}}]}
    (root / "tests/fixtures" / f"{case}.json").write_text(json.dumps(fixture), "utf-8")
    (root / "tests/expected" / f"{case}.json").write_text(json.dumps(expected), "utf-8")


EMPTY_RECORD: dict[str, Any] = {"schema_version": 1, "kind": "holdout-exposures",
                                "exposures": []}


@contextlib.contextmanager
def repository(content: dict[str, Any] | None) -> Iterator[Path]:
    """This repository's guides and cases, CALC-02 and its three cases, CALC-03 and its
    two cases, and ``content`` with an empty exposure record. The two calculations are
    declared last, so they replace any of the same ID."""
    with tempfile.TemporaryDirectory(prefix="accuracy-") as temporary:
        root = Path(temporary)
        for name in ("docs/methodology", "tests/fixtures", "tests/expected"):
            shutil.copytree(ROOT / name, root / name)
        contracts = root / "docs/methodology/CALCULATION_CONTRACTS.md"
        contracts.write_text(contracts.read_text("utf-8") + "\n### `CALC-02` Demo\n"
                             "\n### *Illustrative:* `CALC-03` Illustrative demo\n", "utf-8")
        demo_case(root, "demo-calibration", 100.0)
        demo_case(root, "demo-holdout-1", 200.0)
        demo_case(root, "demo-holdout-2", 0.0)
        for case in ("demo-illustrative-calibration", "demo-illustrative-holdout"):
            demo_case(root, case, 50.0, reference="illustrative", calculation="CALC-03")
        if content is not None:
            (root / "tests/accuracy").mkdir(parents=True)
            (root / "tests/accuracy/contracts.json").write_text(json.dumps(content), "utf-8")
            write_record(root, EMPTY_RECORD)
        yield root


def write_record(root: Path, record: dict[str, Any]) -> None:
    (root / check_accuracy.EXPOSURES).write_text(json.dumps(record, indent=2), "utf-8")


def read_record(root: Path) -> dict[str, Any]:
    data: dict[str, Any] = json.loads((root / check_accuracy.EXPOSURES).read_text("utf-8"))
    return data


def expose(root: Path, *names: str) -> None:
    """Record an exposure of each named contract of the registry in ``root`` under its
    decision, as --record-holdout writes it for results of CANDIDATE."""
    if not names:
        return
    registry = check_accuracy.read_registry(root, None, Report(), "registry")
    assert registry is not None
    record = read_record(root)
    for contract in registry.contracts:
        if contract.id in names:
            earlier = sum(item["contract"] == contract.id for item in record["exposures"])
            record["exposures"].append(check_accuracy.exposure_entry(
                root, contract, earlier, CANDIDATE, [root / check_accuracy.REGISTRY]))
    write_record(root, record)


def findings(content: dict[str, Any] | None, results: dict[str, Any] | None = None,
             commit: str | None = CANDIDATE, exposed: tuple[str, ...] = ()) -> list[str]:
    """The findings on ``content``, measured on ``results`` when given, once the
    holdouts of the ``exposed`` contracts are recorded."""
    with repository(content) as root:
        expose(root, *exposed)
        path = None
        if results is not None:
            path = root / "results.json"
            path.write_text(json.dumps(results), "utf-8")
        return check_accuracy.check(root, path, commit)[0]


def results(**values: float) -> dict[str, Any]:
    """A results file with ``values`` for the CALC-02 cases (underscores for hyphens)."""
    return {"schema_version": 1, "kind": "test-results", "commit": CANDIDATE, "results": [
        {"case": case.replace("_", "-"), "variant": None, "quantity": QUANTITY, "scope": None,
         "value": value} for case, value in values.items()]}


class RegistryTest(unittest.TestCase):
    def assertFinding(self, found: list[str], text: str) -> None:
        self.assertTrue(any(text in item for item in found), found)

    def test_the_repository_registry_passes(self) -> None:
        # Whatever contracts this repository's registry holds, if it has one.
        registry = ROOT / check_accuracy.REGISTRY
        data = json.loads(registry.read_text("utf-8")) if registry.is_file() else {}
        count = len(data.get("contracts", []))
        invariants = len(data.get("invariants", []))
        self.assertEqual(check_accuracy.check(ROOT)[:2], ([], count))
        with contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(check_accuracy.main(["--root", str(ROOT)]), 0)
        self.assertIn(f"accuracy: PASS ({count} contract(s), {invariants} invariant(s); "
                      "registry only", output.getvalue())

    def test_valid_draft_and_frozen_contracts_pass(self) -> None:
        self.assertEqual(findings(registry(ILLUSTRATIVE, FROZEN)), [])

    def test_no_registry_passes_unless_results_are_given(self) -> None:
        self.assertEqual(findings(None), [])
        self.assertFinding(findings(None, results(demo_calibration=100)),
                           "is missing, so there is nothing to measure")

    def test_malformed_contracts(self) -> None:
        cases: dict[str, dict[str, Any]] = {
            "id is not ACC-<number>": {"id": "X-1"},
            "regime is not a text": {"regime": " "},
            "quantities is not a non-empty list": {"quantities": []},
            "metric is not absolute or relative": {"metric": "squared"},
            "threshold is not a finite number of at least 0": {"threshold": -1},
            "status is not draft or frozen": {"status": "final"},
            # An array or an object cannot be hashed: a finding, not a traceback.
            "metric is not absolute or relative ": {"metric": ["absolute"]},
            "status is not draft or frozen ": {"status": {}},
            "a frozen contract names the decision": {"decision": None},
            "holdout_cases is not a list of case IDs": {"holdout_cases": "demo-holdout-1"},
        }
        for text, change in cases.items():
            with self.subTest(text=text):
                self.assertFinding(findings(registry({**FROZEN, **change})), text.strip())
        self.assertFinding(findings(registry({**ILLUSTRATIVE, "decision": "#2"})),
                           "a draft contract has no decision")
        self.assertFinding(findings(registry(FROZEN, FROZEN)), "id appears twice")
        # A list id cannot be compared for repeats: a finding, not a traceback.
        self.assertFinding(findings(registry({**FROZEN, "id": ["ACC-01"]}, FROZEN)),
                           "id is not ACC-<number>")
        self.assertFinding(findings(registry({**FROZEN, "extra": 1})), "keys are not")
        self.assertFinding(findings(registry({**FROZEN, "quantities": [QUANTITY, QUANTITY]})),
                           "quantities lists a name twice")
        self.assertFinding(findings(registry(FROZEN, minimum=0)),
                           "min_holdout_cases is not a whole number of at least 1")
        self.assertFinding(findings({"schema_version": 2}), "envelope is not")

    def test_a_target_no_larger_than_the_threshold_is_optional(self) -> None:
        # Up to the threshold: 0.0001 for ACC-02, 0.000001 for ACC-01.
        for target in (0, 0.00001, 0.0001):
            with self.subTest(target=target):
                self.assertEqual(findings(registry({**FROZEN, "target": target},
                                                   {**ILLUSTRATIVE, "target": target / 100})),
                                 [])
        with repository(registry({**FROZEN, "target": 0.00001}, ILLUSTRATIVE)) as root:
            read = check_accuracy.read_registry(root, None, Report(), "registry")
        assert read is not None
        self.assertEqual([contract.target for contract in read.contracts], [0.00001, None])

    def test_a_malformed_or_larger_target_is_a_finding(self) -> None:
        cases: list[tuple[Any, str]] = [
            (-0.00001, "target is not a finite number of at least 0"),
            ("0.00001", "target is not a finite number of at least 0"),
            (True, "target is not a finite number of at least 0"),
            (None, "target is not a finite number of at least 0"),
            ([0.00001], "target is not a finite number of at least 0"),
            (0.001, "target 0.001 is larger than the threshold 0.0001"),
        ]
        for target, text in cases:
            with self.subTest(target=target):
                found = findings(registry({**FROZEN, "target": target}))
                self.assertEqual(len(found), 1, found)
                self.assertIn(f"ACC-02: {text}", found[0])
        # A malformed threshold is its own finding, not a comparison with the target.
        found = findings(registry({**FROZEN, "threshold": "x", "target": 0.001}))
        self.assertEqual(len(found), 1, found)
        self.assertIn("threshold is not a finite number of at least 0", found[0])

    def test_ids_and_names_are_whole_values_with_ascii_digits(self) -> None:
        cases: list[tuple[dict[str, Any], str]] = [
            ({"id": "ACC-02\n"}, "id is not ACC-<number>"),
            ({"id": "ACC-0\u0662"}, "id is not ACC-<number>"),
            ({"quantities": [f"{QUANTITY}\n"]}, "lower snake case names"),
        ]
        for change, text in cases:
            with self.subTest(text=text):
                self.assertFinding(findings(registry({**FROZEN, **change})), text.strip())

    def test_cases_must_fit_the_contract(self) -> None:
        cases: dict[str, dict[str, Any]] = {
            "calculation CALC-09 is not declared": {"calculation": "CALC-09"},
            "case missing-case has no fixture": {"calibration_cases": ["missing-case"]},
            "case demo-illustrative-calibration tests CALC-03, not CALC-02":
                {"calibration_cases": ["demo-illustrative-calibration"]},
            "case demo-calibration has no expected number for other_amount":
                {"quantities": [QUANTITY, "other_amount"]},
            "case demo-holdout-1 is both calibration and holdout":
                {"calibration_cases": ["demo-holdout-1"]},
            "lists no case": {"calibration_cases": [], "holdout_cases": [], "status": "draft",
                              "decision": None},
        }
        for text, change in cases.items():
            with self.subTest(text=text):
                self.assertFinding(findings(registry({**FROZEN, **change})), text.strip())

    def test_a_repeated_case_does_not_count_twice(self) -> None:
        repeated = {**FROZEN, "holdout_cases": ["demo-holdout-1"] * 3}
        found = findings(registry(repeated, minimum=3))
        self.assertTrue(any("holdout_cases lists a case twice" in item for item in found), found)
        self.assertTrue(any("calibration_cases lists a case twice" in item for item in findings(
            registry({**FROZEN, "calibration_cases": ["demo-calibration"] * 2}))))

    def test_a_frozen_contract_needs_independent_holdout_cases(self) -> None:
        self.assertFinding(findings(registry(FROZEN, minimum=3)),
                           "a frozen contract needs at least 3 holdout case(s), has 2")
        self.assertFinding(findings(registry({**ILLUSTRATIVE, "status": "frozen",
                                              "decision": "#1"}, minimum=1)),
                           "calculation CALC-03 is illustrative, so the contract stays draft")
        with repository(registry(FROZEN)) as root:
            demo_case(root, "demo-holdout-2", 0.0, reference="illustrative")
            self.assertTrue(any("holdout output demo-holdout-2 total_exposure_amount is not "
                                "published or independent" in item
                                for item in check_accuracy.check(root)[0]))
        self.assertEqual(findings(registry({**FROZEN, "status": "draft", "decision": None},
                                           minimum=5)), [])


class ResultsTest(unittest.TestCase):
    def test_a_malformed_results_file_outside_the_root_is_a_finding(self) -> None:
        # An Excel run's results file lives in its evidence folder, not in the checkout.
        with repository(registry(FROZEN)) as root, \
                tempfile.TemporaryDirectory() as outside:
            path = Path(outside) / "results.json"
            path.write_text("{broken", "utf-8")
            found, _, lines = check_accuracy.check(root, path, CANDIDATE)
        self.assertEqual(lines, [])
        self.assertEqual(len(found), 1, found)
        self.assertIn(f"{path}: not valid UTF-8 JSON", found[0])

    def test_results_within_every_bound_pass_and_are_reported(self) -> None:
        good = results(demo_calibration=100.005, demo_holdout_1=200.001, demo_holdout_2=0)
        with repository(registry(FROZEN)) as root:
            expose(root, "ACC-02")
            path = root / "results.json"
            path.write_text(json.dumps(good), "utf-8")
            found, contracts, lines = check_accuracy.check(root, path, CANDIDATE)
        self.assertEqual((found, contracts), ([], 1))
        self.assertEqual([line.split(" error ")[0] for line in lines],
                         ["ACC-02 calibration: worst relative", "ACC-02 holdout: worst relative"])
        worst = [float(line.split(" error ")[1].split(",")[0]) for line in lines]
        self.assertAlmostEqual(worst[0], 5e-05, places=12)
        self.assertAlmostEqual(worst[1], 5e-06, places=12)
        # The holdout line also gives the number of exposures of the holdout.
        self.assertTrue(lines[0].endswith("threshold 0.0001, no target"), lines)
        self.assertTrue(lines[1].endswith("threshold 0.0001, no target, exposures 1"), lines)

    def measured(self, contract: dict[str, Any], **values: float) -> tuple[list[str], list[str]]:
        with repository(registry(contract)) as root:
            expose(root, contract["id"])
            path = root / "results.json"
            path.write_text(json.dumps(results(**values)), "utf-8")
            found, _, lines = check_accuracy.check(root, path, CANDIDATE)
        return found, lines

    def test_a_worst_error_within_the_target_is_reported_as_met(self) -> None:
        found, lines = self.measured({**FROZEN, "target": 0.00001}, demo_calibration=100.0004,
                                     demo_holdout_1=200.001, demo_holdout_2=0)
        self.assertEqual(found, [])
        self.assertEqual(len(lines), 2, lines)
        self.assertEqual([line.endswith(", threshold 0.0001, target 1e-05") for line in lines],
                         [True, False])
        self.assertTrue(lines[1].endswith(", threshold 0.0001, target 1e-05, exposures 1"),
                        lines)

    def test_a_missed_target_within_the_threshold_passes_and_is_reported(self) -> None:
        # Relative 5e-05 on the calibration: above the target 1e-05, within the threshold.
        found, lines = self.measured({**FROZEN, "target": 0.00001}, demo_calibration=100.005,
                                     demo_holdout_1=200.001, demo_holdout_2=0)
        self.assertEqual(found, [])
        self.assertTrue(lines[0].startswith("ACC-02 calibration: worst relative error 4.99"),
                        lines)
        self.assertTrue(lines[0].endswith(", threshold 0.0001, target 1e-05: target missed"),
                        lines)
        self.assertTrue(lines[1].endswith(", threshold 0.0001, target 1e-05, exposures 1"),
                        lines)
        # A target equal to 0 is missed by any error, and still blocks nothing.
        found, lines = self.measured({**FROZEN, "target": 0}, demo_calibration=100.005,
                                     demo_holdout_1=200, demo_holdout_2=0)
        self.assertEqual(found, [])
        self.assertEqual([line.endswith("target 0.0: target missed") for line in lines],
                         [True, False])
        with repository(registry({**FROZEN, "target": 0.00001})) as root, \
                contextlib.redirect_stdout(io.StringIO()) as output:
            expose(root, "ACC-02")
            path = root / "results.json"
            path.write_text(json.dumps(results(demo_calibration=100.005, demo_holdout_1=200,
                                               demo_holdout_2=0)), "utf-8")
            self.assertEqual(check_accuracy.main(["--root", str(root), "--results",
                                                  str(path), "--commit", CANDIDATE]), 0)
        self.assertIn("target 1e-05: target missed\n", output.getvalue())
        self.assertIn("accuracy: PASS (1 contract(s), 0 invariant(s); measured against the "
                      "results)",
                      output.getvalue())

    def test_a_worst_error_above_the_threshold_fails_whatever_the_target(self) -> None:
        # The case tolerance of 0.01 holds; the absolute threshold 0.001 does not.
        loose = {**FROZEN, "metric": "absolute", "threshold": 0.001, "target": 0.0001}
        found, lines = self.measured(loose, demo_calibration=100.0005,
                                     demo_holdout_1=200.005, demo_holdout_2=0)
        self.assertEqual(len(found), 1, found)
        self.assertIn("over holdout cases exceeds the threshold 0.001", found[0])
        self.assertTrue(lines[0].endswith("threshold 0.001, target 0.0001: target missed"),
                        lines)
        self.assertTrue(lines[1].endswith("threshold 0.001, target 0.0001, exposures 1: "
                                          "threshold exceeded"), lines)

    def test_a_result_outside_its_own_tolerance_fails(self) -> None:
        found = findings(registry(FROZEN), results(demo_calibration=100.02, demo_holdout_1=200,
                                                   demo_holdout_2=0), exposed=("ACC-02",))
        self.assertTrue(any("demo-calibration total_exposure_amount is 100.02, expected 100.0 "
                            "within 0.01" in item for item in found), found)

    def test_the_worst_error_is_checked_per_split(self) -> None:
        # Within each case's own tolerance of 0.01, but above the contract's absolute
        # threshold of 0.001 on the holdout only.
        loose = {**FROZEN, "metric": "absolute", "threshold": 0.001}
        found = findings(registry(loose), results(demo_calibration=100.0005,
                                                  demo_holdout_1=200.005, demo_holdout_2=0),
                         exposed=("ACC-02",))
        self.assertEqual(len(found), 1, found)
        self.assertIn("worst absolute error 0.00499", found[0])
        self.assertIn("over holdout cases exceeds the threshold 0.001", found[0])

    def test_a_relative_error_at_zero_and_a_missing_result(self) -> None:
        found = findings(registry({**FROZEN, "holdout_cases": ["demo-holdout-2"]}, minimum=1),
                         results(demo_calibration=100, demo_holdout_2=0.001),
                         exposed=("ACC-02",))
        self.assertTrue(any("worst relative error inf over holdout cases" in item
                            for item in found), found)
        missing = findings(registry(FROZEN), results(demo_calibration=100), exposed=("ACC-02",))
        self.assertTrue(any("no result for demo-holdout-1 total_exposure_amount" in item
                            for item in missing), missing)

    def test_results_must_be_of_the_candidate(self) -> None:
        stale = results(demo_calibration=100, demo_holdout_1=200, demo_holdout_2=0)
        found = findings(registry(FROZEN), stale, commit="b" * 40)
        self.assertTrue(any("results are of commit aaaaaaaaaaaa, not of the candidate "
                            "bbbbbbbbbbbb" in item for item in found), found)
        for commit in (None, "HEAD", "A" * 40, "a" * 40 + "\n"):
            with self.subTest(commit=commit):
                bad = {**stale, "commit": commit}
                self.assertTrue(any("commit is not a full lower-case commit SHA" in item
                                    for item in findings(registry(FROZEN), bad)))
        # Outside a Git checkout, with no --commit, the candidate is unknown.
        self.assertTrue(any("the candidate commit is unknown" in item
                            for item in findings(registry(FROZEN), stale, commit=None)))

    def test_a_malformed_results_file_fails(self) -> None:
        self.assertTrue(any("envelope is not schema_version 1, kind 'test-results'" in item
                            for item in findings(registry(FROZEN), {"kind": "other"})))
        bad = results(demo_calibration=100)
        bad["results"].append({"case": "demo-calibration"})
        self.assertTrue(any("results[1] is not case, variant, quantity, scope" in item
                            for item in findings(registry(FROZEN), bad)))
        for field, value in (("variant", ["lf"]), ("case", {"id": 1}), ("quantity", 7)):
            with self.subTest(field=field):
                bad = results(demo_calibration=100)
                bad["results"][0][field] = value
                self.assertTrue(any("results[0] case and quantity are not text, or its "
                                    "variant is not text or null" in item
                                    for item in findings(registry(FROZEN), bad)))

    def test_main_reports_the_measurement(self) -> None:
        with repository(registry(FROZEN)) as root, \
                contextlib.redirect_stdout(io.StringIO()) as output:
            path = root / "results.json"
            path.write_text(json.dumps(results(demo_calibration=1, demo_holdout_1=200,
                                               demo_holdout_2=0)), "utf-8")
            self.assertEqual(check_accuracy.main(["--root", str(root), "--results",
                                                  str(path), "--commit", CANDIDATE]), 1)
        self.assertIn("accuracy: FAIL (1 contract(s), 0 invariant(s); measured against the "
                      "results)",
                      output.getvalue())

    def test_the_results_files_of_several_suites_are_measured_together(self) -> None:
        with repository(registry(FROZEN)) as root, \
                contextlib.redirect_stdout(io.StringIO()) as output:
            expose(root, "ACC-02")
            first, second = root / "results.json", root / "project-results.json"
            first.write_text(json.dumps(results(demo_calibration=100)), "utf-8")
            second.write_text(json.dumps(results(demo_holdout_1=200, demo_holdout_2=0)),
                              "utf-8")
            self.assertEqual(check_accuracy.check(root, [first, second], CANDIDATE)[:2],
                             ([], 1))
            self.assertEqual(check_accuracy.main(["--root", str(root), "--results", str(first),
                                                  "--results", str(second),
                                                  "--commit", CANDIDATE]), 0)
            # Each output is computed by one suite: a result in two files is rejected.
            second.write_text(json.dumps(results(demo_calibration=100, demo_holdout_1=200,
                                                 demo_holdout_2=0)), "utf-8")
            found = check_accuracy.check(root, [first, second], CANDIDATE)[0]
        self.assertIn("accuracy: PASS (1 contract(s), 0 invariant(s); measured against the "
                      "results)",
                      output.getvalue())
        self.assertEqual(found, [f"{second}: demo-calibration {QUANTITY} is also in {first}"])


@unittest.skipUnless(shutil.which("git"), "needs Git")
class WideningTest(unittest.TestCase):
    def tagged(self, before: dict[str, Any], now: dict[str, Any] | None,
               widening: bool = True) -> list[str]:
        """Commit and tag ``before`` as v1.0.0, then check ``now`` in the working tree."""
        with repository(registry(before)) as root:
            init(root)
            git(root, "add", "-A")
            git(root, "commit", "-qm", "Release")
            git(root, "tag", "-a", "v1.0.0", "-m", "v1.0.0")
            contracts = [] if now is None else [now]
            (root / "tests/accuracy/contracts.json").write_text(
                json.dumps(registry(*contracts)), "utf-8")
            return check_accuracy.check(root, widening=widening)[0]

    def test_the_tag_comparison_can_be_skipped_for_a_copy_of_a_commit(self) -> None:
        # check_excel_run.py measures a copy of a commit's cases, with no history.
        widened = {**FROZEN, "threshold": 0.001}
        self.assertTrue(self.tagged(FROZEN, widened))
        self.assertEqual(self.tagged(FROZEN, widened, widening=False), [])
        self.assertEqual(self.tagged(FROZEN, {**FROZEN, "status": "draft", "decision": None},
                                     widening=False), [])

    def test_a_frozen_threshold_widens_only_with_a_new_decision(self) -> None:
        self.assertEqual(self.tagged(FROZEN, FROZEN), [])
        self.assertEqual(self.tagged(FROZEN, {**FROZEN, "threshold": 0.00001}), [])
        widened = self.tagged(FROZEN, {**FROZEN, "threshold": 0.001})
        self.assertTrue(any("ACC-02 threshold 0.001 is larger than 0.0001 at v1.0.0; a weaker "
                            "or different guarantee needs a new decision" in item
                            for item in widened), widened)
        self.assertEqual(self.tagged(FROZEN, {**FROZEN, "threshold": 0.001,
                                              "decision": "#2"}), [])

    def test_a_change_of_target_needs_no_new_decision(self) -> None:
        aimed = {**FROZEN, "target": 0.00001}
        for before, now in ((FROZEN, aimed), (aimed, FROZEN),
                            (aimed, {**FROZEN, "target": 0.0001}),
                            (aimed, {**FROZEN, "target": 0})):
            with self.subTest(before=before.get("target"), now=now.get("target")):
                self.assertEqual(self.tagged(before, now), [])
        # The threshold keeps its rules, whatever the target does.
        widened = self.tagged(aimed, {**aimed, "threshold": 0.001})
        self.assertTrue(any("ACC-02 threshold 0.001 is larger than 0.0001 at v1.0.0" in item
                            for item in widened), widened)
        self.assertEqual(self.tagged(aimed, {**aimed, "threshold": 0.00001}), [])

    def test_any_other_change_to_a_frozen_guarantee_needs_a_new_decision(self) -> None:
        changes: dict[str, dict[str, Any]] = {
            "metric is 'absolute', not 'relative'": {"metric": "absolute"},
            "regime is 'CRR2', not 'CRR3'": {"regime": "CRR2"},
            "drops holdout cases demo-holdout-2": {"holdout_cases": ["demo-holdout-1"]},
            "drops calibration cases demo-calibration": {"calibration_cases": []},
        }
        for text, change in changes.items():
            with self.subTest(text=text):
                found = self.tagged(FROZEN, {**FROZEN, **change})
                self.assertTrue(any(f"ACC-02 {text} at v1.0.0" in item for item in found),
                                found)
                self.assertFalse(any("ACC-02" in item and "v1.0.0" in item for item in
                                     self.tagged(FROZEN, {**FROZEN, **change,
                                                          "decision": "#2"})))
        # Moving a holdout case to calibration drops it from the holdout.
        moved = {**FROZEN, "calibration_cases": ["demo-calibration", "demo-holdout-2"],
                 "holdout_cases": ["demo-holdout-1"]}
        self.assertTrue(any("drops holdout cases demo-holdout-2" in item
                            for item in self.tagged(FROZEN, moved)))
        # Adding a case only strengthens the guarantee.
        self.assertEqual(self.tagged({**FROZEN, "holdout_cases": ["demo-holdout-1"]}, FROZEN),
                         [])

    def test_a_frozen_contract_is_neither_removed_nor_unfrozen(self) -> None:
        for now in (None, {**FROZEN, "status": "draft", "decision": None}):
            with self.subTest(now=now):
                found = self.tagged(FROZEN, now)
                self.assertTrue(any("ACC-02 was frozen at v1.0.0" in item for item in found),
                                found)
        draft = {**FROZEN, "status": "draft", "decision": None}
        self.assertEqual(self.tagged(draft, None), [])

    def test_deleting_the_registry_does_not_drop_a_released_frozen_contract(self) -> None:
        for before, dropped in ((FROZEN, True), ({**FROZEN, "status": "draft",
                                                  "decision": None}, False)):
            with self.subTest(status=before["status"]), \
                    repository(registry(before)) as root:
                init(root)
                git(root, "add", "-A")
                git(root, "commit", "-qm", "Release")
                git(root, "tag", "-a", "v1.0.0", "-m", "v1.0.0")
                (root / "tests/accuracy/contracts.json").unlink()
                found = check_accuracy.check(root)[0]
                self.assertEqual(any("ACC-02 was frozen at v1.0.0" in item for item in found),
                                 dropped, found)


class MalformedExpectedTest(unittest.TestCase):
    def outputs(self, expected: dict[str, Any]) -> list[Any]:
        with repository(None) as root:
            (root / "tests/expected/demo-calibration.json").write_text(json.dumps(expected), "utf-8")
            return check_accuracy.case_outputs(root, "demo-calibration", Report())[1]

    def test_malformed_outputs_give_no_measurement_and_no_error(self) -> None:
        # check_test_cases.py reports each of these forms; the accuracy check skips them.
        item = {"quantity": QUANTITY, "scope": None, "value": 1.0, "unit": "EUR",
                "tolerance": {"absolute": 0.01, "relative": 0}, "reference": {"class": "independent"}}
        base = {"schema_version": 1, "kind": "test-expected", "fixture": "demo-calibration"}
        cases = {
            "variants as a list": {**base, "variants": [[item]]},
            "no relative tolerance": {**base, "outputs": [{**item, "tolerance": {"absolute": 0.01}}]},
            "a tolerance that is not a number": {**base, "outputs": [
                {**item, "tolerance": {"absolute": "0.01", "relative": 0}}]},
        }
        for name, expected in cases.items():
            with self.subTest(name):
                self.assertEqual(self.outputs(expected), [])
        reference = self.outputs({**base, "outputs": [{**item, "reference": "independent"}]})
        self.assertEqual([output.reference for output in reference], ["None"])
        self.assertEqual([output.reference for output in self.outputs({**base, "outputs": [item]})],
                         ["independent"])


# Invariants on the CALC-02 cases: total_exposure_amount is 0 in demo-holdout-2,
# 100 in demo-calibration and 200 in demo-holdout-1, each within 0.01.
BOUNDS: dict[str, Any] = {
    "id": "INV-01", "kind": "bounds", "calculation": "CALC-02", "basis": "Synthetic fixture",
    "quantity": QUANTITY, "cases": ["demo-calibration", "demo-holdout-1", "demo-holdout-2"],
    "minimum": 0, "maximum": None,
}
MONOTONE: dict[str, Any] = {
    "id": "INV-02", "kind": "monotone", "calculation": "CALC-02", "basis": "Synthetic fixture",
    "quantity": QUANTITY, "scope": None, "direction": "non-decreasing",
    "sequence": [{"case": "demo-holdout-2", "variant": None},
                 {"case": "demo-calibration", "variant": None},
                 {"case": "demo-holdout-1", "variant": None}],
}
# demo-parts: part_a_amount 60 and part_b_amount 40, each within 0.01, add up to
# total_exposure_amount 100, within 0.01, in the case and for its record P1.
SUM: dict[str, Any] = {
    "id": "INV-03", "kind": "sum", "calculation": "CALC-02", "basis": "Synthetic fixture",
    "parts": ["part_a_amount", "part_b_amount"], "total": QUANTITY, "cases": ["demo-parts"],
}
PARTS = {"part_a_amount": 60.0, "part_b_amount": 40.0, QUANTITY: 100.0}
RECORD = {"dataset": "DS-01", "record": "P1"}


def parts_case(root: Path, values: dict[str, float] = PARTS) -> None:
    """demo-parts, a CALC-02 case with ``values`` for the whole case and for record P1."""
    demo_case(root, "demo-parts", 0.0)
    path = root / "tests/expected/demo-parts.json"
    expected = json.loads(path.read_text("utf-8"))
    expected["outputs"] = [
        {"quantity": quantity, "scope": scope, "rule": "CALC-02.1", "value": value,
         "unit": "EUR", "tolerance": {"absolute": 0.01, "relative": 0},
         "reference": {"class": "independent"}}
        for scope in (None, RECORD) for quantity, value in values.items()]
    path.write_text(json.dumps(expected), "utf-8")


def variant_case(root: Path, values: dict[str, float]) -> None:
    """demo-variants, a file-based CALC-02 case with one total per variant."""
    folder = root / "tests/fixtures/demo-variants"
    folder.mkdir()
    (folder / "case.json").write_text(json.dumps({"calculation": "CALC-02"}), "utf-8")
    expected = {"schema_version": 1, "kind": "test-expected", "fixture": "demo-variants",
                "variants": {variant: [
                    {"quantity": QUANTITY, "scope": None, "rule": "CALC-02.1", "value": value,
                     "unit": "EUR", "tolerance": {"absolute": 0.01, "relative": 0},
                     "reference": {"class": "independent"}}]
                    for variant, value in values.items()}}
    (root / "tests/expected/demo-variants.json").write_text(json.dumps(expected), "utf-8")


def with_invariants(*invariants: dict[str, Any], contracts: tuple[dict[str, Any], ...] = ()
                    ) -> dict[str, Any]:
    return {**registry(*contracts), "invariants": [copy.deepcopy(item) for item in invariants]}


def result(case: str, value: float, quantity: str = QUANTITY, variant: str | None = None,
           scope: dict[str, str] | None = None) -> dict[str, Any]:
    return {"case": case, "variant": variant, "quantity": quantity, "scope": scope,
            "value": value}


def results_of(*items: dict[str, Any]) -> dict[str, Any]:
    return {"schema_version": 1, "kind": "test-results", "commit": CANDIDATE,
            "results": list(items)}


def parts_results(values: dict[str, float] = PARTS) -> list[dict[str, Any]]:
    return [result("demo-parts", value, quantity, scope=scope)
            for scope in (None, RECORD) for quantity, value in values.items()]


def checked(content: dict[str, Any], document: dict[str, Any] | None = None,
            parts: dict[str, float] | None = PARTS,
            variants: dict[str, float] | None = None) -> tuple[list[str], list[str]]:
    """The findings and the measurement lines of ``content``, with demo-parts and,
    when given, demo-variants; measured on ``document`` when given, once the holdout
    of each frozen contract is recorded."""
    with repository(content) as root:
        expose(root, *[item["id"] for item in content.get("contracts", [])
                       if isinstance(item, dict) and item.get("status") == "frozen"])
        if parts is not None:
            parts_case(root, parts)
        if variants is not None:
            variant_case(root, variants)
        path = None
        if document is not None:
            path = root / "results.json"
            path.write_text(json.dumps(document), "utf-8")
        found, _, lines = check_accuracy.check(root, path, CANDIDATE)
    return found, lines


# A frozen contract on the CALC-02 cases whose threshold is their tolerance.
ABSOLUTE: dict[str, Any] = {**FROZEN, "metric": "absolute", "threshold": 0.01}
GOOD = results_of(result("demo-calibration", 100), result("demo-holdout-1", 200),
                  result("demo-holdout-2", 0), *parts_results())


class InvariantRegistryTest(unittest.TestCase):
    def assertFinding(self, found: list[str], text: str) -> None:
        self.assertTrue(any(text in item for item in found), found)

    def test_declared_invariants_are_optional_and_pass_when_well_formed(self) -> None:
        self.assertEqual(checked(registry(FROZEN))[0], [])
        self.assertEqual(checked(with_invariants(contracts=(FROZEN,)))[0], [])
        self.assertEqual(checked(with_invariants(BOUNDS, MONOTONE, SUM,
                                                 contracts=(FROZEN,)))[0], [])
        with repository(with_invariants(BOUNDS, MONOTONE, SUM)) as root:
            parts_case(root)
            read = check_accuracy.read_registry(root, None, Report(), "registry")
            with contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(check_accuracy.main(["--root", str(root)]), 0)
        assert read is not None
        self.assertEqual([item.id for item in read.invariants], ["INV-01", "INV-02", "INV-03"])
        self.assertIn("accuracy: PASS (0 contract(s), 3 invariant(s); registry only",
                      output.getvalue())
        # Every limit is optional, as long as one is stated; a constant may be the total.
        for change in ({"minimum": None, "maximum": 300}, {"minimum": 0, "maximum": 200},
                       {"minimum": -1.5, "maximum": None}):
            with self.subTest(change=change):
                self.assertEqual(checked(with_invariants({**BOUNDS, **change}))[0], [])
        self.assertEqual(checked(with_invariants({**SUM, "total": 100}),
                                 parts={"part_a_amount": 60, "part_b_amount": 40})[0], [])

    def test_a_malformed_invariant_is_a_finding(self) -> None:
        cases: list[tuple[dict[str, Any], str]] = [
            ({**BOUNDS, "id": "INV-1\n"}, "id is not INV-<number>"),
            ({**BOUNDS, "calculation": 2}, "calculation is not a text"),
            ({**BOUNDS, "basis": " "}, "basis is not a text"),
            ({**BOUNDS, "quantity": "Total"}, "quantity is not a lower snake case name"),
            ({**BOUNDS, "cases": []}, "cases is not a non-empty list of case IDs"),
            ({**BOUNDS, "cases": ["demo-calibration"] * 2}, "cases lists a case twice"),
            ({**BOUNDS, "minimum": "0"}, "minimum is not a finite number or null"),
            ({**BOUNDS, "maximum": True}, "maximum is not a finite number or null"),
            ({**BOUNDS, "minimum": None}, "minimum and maximum are both null"),
            ({**BOUNDS, "minimum": 2, "maximum": 1}, "minimum 2 is larger than the maximum 1"),
            ({**MONOTONE, "direction": "up"},
             "direction is not non-decreasing or non-increasing"),
            ({**MONOTONE, "scope": {"dataset": "DS-01"}}, "scope is not null or an object"),
            ({**MONOTONE, "sequence": MONOTONE["sequence"][:1]},
             "sequence is not a list of at least two points"),
            ({**MONOTONE, "sequence": [{"case": "demo-calibration"}] * 2},
             "sequence is not a list of at least two points"),
            ({**MONOTONE, "sequence": [MONOTONE["sequence"][0]] * 2},
             "sequence lists a point twice"),
            ({**SUM, "parts": []}, "parts is not a non-empty list of lower snake case names"),
            ({**SUM, "parts": ["part_a_amount"] * 2}, "parts lists a name twice"),
            ({**SUM, "total": "part_a_amount"}, "total part_a_amount is also a part"),
            ({**SUM, "total": None}, "total is not a lower snake case name or a finite number"),
            ({**BOUNDS, "kind": "range"}, "is not an object whose kind is bounds, monotone or "
                                          "sum"),
            # An array or an object cannot be hashed: a finding, not a traceback.
            ({**BOUNDS, "kind": []}, "is not an object whose kind is bounds, monotone or sum"),
            ({**BOUNDS, "kind": {}}, "is not an object whose kind is bounds, monotone or sum"),
            ({**MONOTONE, "direction": ["non-decreasing"]},
             "direction is not non-decreasing or non-increasing"),
            ({**MONOTONE, "direction": {}}, "direction is not non-decreasing or non-increasing"),
            ({**BOUNDS, "direction": "non-decreasing"}, "keys of a bounds invariant are not"),
            ({key: value for key, value in SUM.items() if key != "basis"},
             "keys of a sum invariant are not"),
        ]
        for invariant, text in cases:
            with self.subTest(text=text):
                found = checked(with_invariants(invariant))[0]
                self.assertEqual(len(found), 1, found)
                self.assertIn(text, found[0])
        self.assertFinding(checked(with_invariants(BOUNDS, BOUNDS))[0], "id appears twice")
        self.assertFinding(checked({**registry(), "invariants": {}})[0],
                           "invariants is not a list")
        self.assertFinding(checked({**registry(), "invariant": []})[0],
                           "envelope is not schema_version 1, kind 'accuracy-contracts', "
                           "min_holdout_cases and contracts, and optionally invariants")

    def test_an_unknown_calculation_case_or_quantity_is_a_finding(self) -> None:
        cases: list[tuple[dict[str, Any], str]] = [
            ({**BOUNDS, "calculation": "CALC-09"},
             "calculation CALC-09 is not declared by the calculation contracts"),
            ({**BOUNDS, "cases": ["missing-case"]}, "case missing-case has no fixture"),
            ({**BOUNDS, "cases": ["demo-illustrative-calibration"]},
             "case demo-illustrative-calibration tests CALC-03, not CALC-02"),
            ({**BOUNDS, "quantity": "other_amount"},
             "case demo-calibration has no expected number for other_amount"),
            ({**MONOTONE, "quantity": "other_amount"},
             "point demo-holdout-2 other_amount has no expected number"),
            ({**MONOTONE, "scope": RECORD},
             "point demo-holdout-2 total_exposure_amount (DS-01/P1) has no expected number"),
            ({**MONOTONE, "sequence": [*MONOTONE["sequence"],
                                       {"case": "demo-calibration", "variant": "lf"}]},
             "point demo-calibration[lf] total_exposure_amount has no expected number"),
            ({**MONOTONE, "sequence": [*MONOTONE["sequence"],
                                       {"case": "missing-case", "variant": None}]},
             "case missing-case has no fixture"),
            ({**SUM, "parts": ["part_a_amount", "part_c_amount"]},
             "demo-parts has no expected number for part_c_amount"),
            ({**SUM, "cases": ["demo-calibration"]},
             "demo-calibration has no expected number for part_a_amount"),
            ({**SUM, "cases": ["demo-calibration"], "total": 100},
             "case demo-calibration has no expected number for part_a_amount, "
             "part_b_amount"),
        ]
        for invariant, text in cases:
            with self.subTest(text=text):
                self.assertFinding(checked(with_invariants(invariant))[0], text)

    def test_expected_values_that_break_an_invariant_are_a_finding(self) -> None:
        # The declared invariant does not hold on the cases' own expected values.
        cases: list[tuple[dict[str, Any], str]] = [
            ({**BOUNDS, "minimum": 50},
             "the expected values break it: demo-holdout-2 total_exposure_amount is 0.0, "
             "below the minimum 50.0"),
            ({**BOUNDS, "maximum": 150},
             "the expected values break it: demo-holdout-1 total_exposure_amount is 200.0, "
             "above the maximum 150.0"),
            ({**MONOTONE, "direction": "non-increasing"},
             "the expected values break it: total_exposure_amount increases from 0.0 at "
             "demo-holdout-2 to 100.0 at demo-calibration, by 100.0, beyond the allowance "
             "0.02"),
            ({**SUM, "total": 99},
             "the expected values break it: at demo-parts, part_a_amount + part_b_amount is "
             "100.0 and the total is 99.0: the difference 1.0 is beyond the allowance 0.02"),
        ]
        for invariant, text in cases:
            with self.subTest(text=text):
                found = checked(with_invariants(invariant))[0]
                self.assertTrue(found and all("INV-0" in item for item in found), found)
                self.assertFinding(found, text)


class InvariantResultsTest(unittest.TestCase):
    def assertFinding(self, found: list[str], text: str) -> None:
        self.assertTrue(any(text in item for item in found), found)

    def changed(self, case: str, value: float, quantity: str = QUANTITY,
                scope: dict[str, str] | None = None) -> dict[str, Any]:
        """GOOD with one result changed."""
        document = copy.deepcopy(GOOD)
        for item in document["results"]:
            if (item["case"], item["quantity"], item["scope"]) == (case, quantity, scope):
                item["value"] = value
        return document

    def test_invariants_that_hold_are_reported(self) -> None:
        found, lines = checked(with_invariants(BOUNDS, MONOTONE, SUM, contracts=(FROZEN,)),
                               GOOD)
        self.assertEqual(found, [])
        self.assertEqual(lines[2:], [
            "INV-01 bounds: 0.0 <= total_exposure_amount, 3 result(s) checked",
            "INV-02 monotone: total_exposure_amount non-decreasing over 3 point(s), "
            "2 step(s) checked, 0 reversal(s) within the allowance",
            "INV-03 sum: part_a_amount + part_b_amount = total_exposure_amount, "
            "2 group(s) checked"])
        found, lines = checked(with_invariants({**BOUNDS, "minimum": None, "maximum": 1000},
                                               {**SUM, "total": 100}),
                               GOOD, parts={"part_a_amount": 60, "part_b_amount": 40})
        self.assertEqual(found, [])
        self.assertEqual(lines, ["INV-01 bounds: total_exposure_amount <= 1000.0, "
                                 "3 result(s) checked",
                                 "INV-03 sum: part_a_amount + part_b_amount = 100.0, "
                                 "2 group(s) checked"])

    def test_a_result_outside_the_bounds_fails_however_small(self) -> None:
        # -0.001 is within the output's own tolerance of 0.01, but a domain rule
        # applies before tolerance.
        found, lines = checked(with_invariants(BOUNDS, contracts=(ABSOLUTE,)),
                               self.changed("demo-holdout-2", -0.001))
        self.assertEqual(found, ["tests/accuracy/contracts.json: INV-01: demo-holdout-2 "
                                 "total_exposure_amount is -0.001, below the minimum 0.0"])
        self.assertEqual(lines[2], "INV-01 bounds: 0.0 <= total_exposure_amount, "
                                   "3 result(s) checked, 1 outside the bounds")
        found = checked(with_invariants({**BOUNDS, "maximum": 200}),
                        self.changed("demo-holdout-1", 200.005))[0]
        self.assertEqual(found, ["tests/accuracy/contracts.json: INV-01: demo-holdout-1 "
                                 "total_exposure_amount is 200.005, above the maximum 200.0"])

    def test_a_break_fails_whatever_the_contract_status(self) -> None:
        # A frozen contract, a draft one, and no contract at all over the cases.
        draft = {**ABSOLUTE, "status": "draft", "decision": None}
        for contracts in ((ABSOLUTE,), (draft,), ()):
            with self.subTest(contracts=[item["status"] for item in contracts]):
                found = checked(with_invariants(BOUNDS, contracts=contracts),
                                self.changed("demo-holdout-2", -0.001))[0]
                self.assertEqual(len(found), 1, found)
                self.assertIn("INV-01: demo-holdout-2 total_exposure_amount is -0.001, below "
                              "the minimum 0.0", found[0])

    def variants(self, *values: float, direction: str = "non-decreasing"
                 ) -> tuple[list[str], list[str]]:
        """A monotone invariant over the variants of demo-variants, whose expected
        values are 10, 10 and 20 within 0.01, measured on ``values``. No contract
        covers the case, so only the invariant checks its results."""
        expected = {"a-low": 10.0, "b-mid": 10.0, "c-high": 20.0}
        sequence = [{"case": "demo-variants", "variant": name} for name in expected]
        if direction == "non-increasing":
            sequence.reverse()
        document = results_of(*[result("demo-variants", value, variant=name)
                                for name, value in zip(expected, values)])
        return checked(with_invariants({**MONOTONE, "sequence": sequence,
                                        "direction": direction}),
                       document, parts=None, variants=expected)

    def test_a_monotone_reversal_within_the_allowance_is_only_counted(self) -> None:
        # From 10.008 to 9.993: back by 0.015, within the allowance 0.01 + 0.01.
        found, lines = self.variants(10.008, 9.993, 20)
        self.assertEqual(found, [])
        self.assertEqual(lines, ["INV-02 monotone: total_exposure_amount non-decreasing over "
                                 "3 point(s), 2 step(s) checked, 1 reversal(s) within the "
                                 "allowance"])
        found, lines = self.variants(10.008, 9.993, 20, direction="non-increasing")
        self.assertEqual(found, [])
        self.assertTrue(lines[0].endswith("1 reversal(s) within the allowance"), lines)

    def test_a_monotone_reversal_beyond_the_allowance_fails(self) -> None:
        # From 10.0099 to 9.9897: back by 0.0202, beyond the allowance 0.02.
        found, lines = self.variants(10.0099, 9.9897, 20)
        self.assertEqual(len(found), 1, found)
        self.assertIn("INV-02: total_exposure_amount decreases from 10.0099 at "
                      "demo-variants[a-low] to 9.9897 at demo-variants[b-mid], by 0.0202",
                      found[0])
        self.assertIn("beyond the allowance 0.02", found[0])
        self.assertTrue(lines[0].endswith("2 step(s) checked, 0 reversal(s) within the "
                                          "allowance, 1 reversal(s) beyond the allowance"),
                        lines)
        found = self.variants(10, 10, 9, direction="non-increasing")[0]
        self.assertEqual(len(found), 1, found)
        self.assertIn("INV-02: total_exposure_amount increases from 9.0 at "
                      "demo-variants[c-high] to 10.0 at demo-variants[b-mid], by 1.0, beyond "
                      "the allowance 0.02", found[0])

    def test_a_sum_beyond_the_allowance_fails(self) -> None:
        # demo-parts is in no contract, so only the invariant checks its results.
        def parts(**record: float) -> dict[str, Any]:
            return results_of(*parts_results(), *[
                result("demo-parts", value, quantity, scope=RECORD)
                for quantity, value in record.items()])
        record = {"part_a_amount": 60.009, "part_b_amount": 40.009, QUANTITY: 99.991}
        document = parts(**record)
        document["results"] = document["results"][:3] + document["results"][6:]
        # 100.018 against 99.991: 0.027, within the allowance 0.03 of the three outputs.
        self.assertEqual(checked(with_invariants(SUM), document)[0], [])
        document["results"][3]["value"] = 60.02
        document["results"][4]["value"] = 40.02
        document["results"][5]["value"] = 100
        found, lines = checked(with_invariants(SUM), document)
        self.assertEqual(len(found), 1, found)
        self.assertIn("INV-03: at demo-parts (DS-01/P1), part_a_amount + part_b_amount is "
                      "100.04 and the total is 100.0: the difference 0.04", found[0])
        self.assertIn("is beyond the allowance 0.03", found[0])
        self.assertTrue(lines[0].endswith("2 group(s) checked, 1 beyond the allowance"), lines)
        # A constant total has no tolerance of its own: the allowance is the parts'.
        constant = {**SUM, "total": 100}
        found = checked(with_invariants(constant), results_of(
            result("demo-parts", 60.01, "part_a_amount"),
            result("demo-parts", 40.011, "part_b_amount"),
            result("demo-parts", 60, "part_a_amount", scope=RECORD),
            result("demo-parts", 40, "part_b_amount", scope=RECORD)),
            parts={"part_a_amount": 60, "part_b_amount": 40})[0]
        self.assertEqual(len(found), 1, found)
        self.assertIn("INV-03: at demo-parts, part_a_amount + part_b_amount is 100.021 and "
                      "the total is 100.0: the difference 0.0210", found[0])

    def test_main_prints_the_invariant_lines(self) -> None:
        with repository(with_invariants(BOUNDS)) as root, \
                contextlib.redirect_stdout(io.StringIO()) as output:
            path = root / "results.json"
            path.write_text(json.dumps(results_of(result("demo-calibration", 100),
                                                  result("demo-holdout-1", -1),
                                                  result("demo-holdout-2", 0))), "utf-8")
            self.assertEqual(check_accuracy.main(["--root", str(root), "--results", str(path),
                                                  "--commit", CANDIDATE]), 1)
        self.assertIn("INV-01 bounds: 0.0 <= total_exposure_amount, 3 result(s) checked, "
                      "1 outside the bounds\n", output.getvalue())
        self.assertIn("FAIL: tests/accuracy/contracts.json: INV-01: demo-holdout-1 "
                      "total_exposure_amount is -1.0, below the minimum 0.0\n", output.getvalue())
        self.assertIn("accuracy: FAIL (0 contract(s), 1 invariant(s); measured against the "
                      "results)", output.getvalue())

    def test_a_missing_result_of_an_invariant_is_a_finding(self) -> None:
        document = copy.deepcopy(GOOD)
        document["results"] = [item for item in document["results"]
                               if item["case"] != "demo-parts" or item["scope"] != RECORD
                               or item["quantity"] != "part_b_amount"]
        found, lines = checked(with_invariants(SUM), document)
        self.assertEqual(found, ["tests/accuracy/contracts.json: INV-03: no result for "
                                 "demo-parts part_b_amount (DS-01/P1)"])
        self.assertEqual(lines, ["INV-03 sum: part_a_amount + part_b_amount = "
                                 "total_exposure_amount, 1 group(s) checked"])


# Results of CALC-02 within every bound; demo-holdout-1 is off by 0.5 in OFF_HOLDOUT,
# beyond its own tolerance and any contract threshold here.
EXACT = results(demo_calibration=100, demo_holdout_1=200, demo_holdout_2=0)
OFF_HOLDOUT = results(demo_calibration=100, demo_holdout_1=200.5, demo_holdout_2=0)
DRAFT = {**FROZEN, "status": "draft", "decision": None}


def run(root: Path, *args: str, document: dict[str, Any] | None = EXACT) -> tuple[int, str]:
    """main() on ``root`` with ``args``, measuring ``document`` as results.json of
    CANDIDATE when given; the exit code and the output."""
    options = ["--root", str(root), "--commit", CANDIDATE]
    if document is not None:
        path = root / "results.json"
        path.write_text(json.dumps(document), "utf-8")
        options += ["--results", str(path)]
    with contextlib.redirect_stdout(io.StringIO()) as output:
        code = check_accuracy.main([*options, *args])
    return code, output.getvalue()


def measured_in(root: Path, document: dict[str, Any] = EXACT) -> tuple[list[str], list[str]]:
    path = root / "results.json"
    path.write_text(json.dumps(document), "utf-8")
    found, _, lines = check_accuracy.check(root, path, CANDIDATE)
    return found, lines


class SealedHoldoutTest(unittest.TestCase):
    def test_a_draft_holdout_is_neither_measured_nor_printed(self) -> None:
        with repository(registry(DRAFT)) as root:
            found, lines = measured_in(root, OFF_HOLDOUT)
            code, output = run(root, document=OFF_HOLDOUT)
        self.assertEqual(found, [])
        self.assertEqual(lines, ["ACC-02 calibration: worst relative error 0.0, threshold "
                                 "0.0001, no target",
                                 "ACC-02 holdout: sealed while the contract is draft, "
                                 "exposures 0"])
        self.assertEqual(code, 0, output)
        self.assertNotIn("200.5", output)
        self.assertNotIn("holdout cases", output)
        # A draft holdout needs no result either.
        with repository(registry(DRAFT)) as root:
            self.assertEqual(measured_in(root, results(demo_calibration=100))[0], [])

    def test_a_frozen_holdout_with_no_exposure_is_sealed_and_fails_with_results(self) -> None:
        with repository(registry(FROZEN)) as root:
            self.assertEqual(check_accuracy.check(root)[0], [])
            found, lines = measured_in(root, OFF_HOLDOUT)
        self.assertEqual(found, ["tests/accuracy/contracts.json: ACC-02: its holdout has no "
                                 "exposure recorded under decision #1, so it was not measured: "
                                 "measure it once with tools/check_accuracy.py "
                                 "--record-holdout ACC-02"])
        self.assertEqual(lines[1], "ACC-02 holdout: sealed until its exposure under decision #1 "
                                   "is recorded, exposures 0")


class RecordHoldoutTest(unittest.TestCase):
    def test_the_first_exposure_is_recorded_then_measured(self) -> None:
        design = "docs/methodology/ACC-02-design.md"
        contract = {**FROZEN, "design": design}
        with repository(registry(contract)) as root:
            (root / design).write_text("# Design\n\nThe threshold and the holdout.\n", "utf-8")
            code, output = run(root, "--record-holdout", "ACC-02")
            record = read_record(root)
            files = {name: hashlib.sha256((root / name).read_bytes()).hexdigest() for name in (
                "results.json", design, "tests/fixtures/demo-holdout-1.json",
                "tests/expected/demo-holdout-1.json", "tests/fixtures/demo-holdout-2.json",
                "tests/expected/demo-holdout-2.json")}
            size = (root / design).stat().st_size
            # A later measurement is a regression: it records nothing.
            before = (root / check_accuracy.EXPOSURES).read_bytes()
            again, regression = run(root)
            after = (root / check_accuracy.EXPOSURES).read_bytes()
        self.assertEqual(code, 0, output)
        self.assertIn("ACC-02 holdout: worst relative error 0.0, threshold 0.0001, no target, "
                      "exposures 1\n", output)
        self.assertIn("Recorded: exposure 1 of the holdout of ACC-02 under decision #1, with the "
                      "results of commit aaaaaaaaaaaa, in tests/accuracy/holdout_exposures.json: "
                      "commit it\n", output)
        canonical = json.dumps(contract, sort_keys=True, separators=(",", ":"))
        self.assertEqual(record["exposures"], [{
            "contract": "ACC-02", "decision": "#1", "exposures": 1, "commit": CANDIDATE,
            "results": [{"file": "results.json", "sha256": files["results.json"]}],
            "entry_sha256": hashlib.sha256(canonical.encode()).hexdigest(),
            "threshold": 0.0001,
            "holdout": [{"case": case, "fixture_sha256": files[f"tests/fixtures/{case}.json"],
                         "expected_sha256": files[f"tests/expected/{case}.json"]}
                        for case in ("demo-holdout-1", "demo-holdout-2")],
            "design": {"path": design, "bytes": size, "sha256": files[design]}}])
        self.assertEqual((again, before), (0, after), regression)
        self.assertIn("ACC-02 holdout: worst relative error 0.0, threshold 0.0001, no target, "
                      "exposures 1\n", regression)
        self.assertNotIn("Recorded:", regression)

    def test_a_failing_holdout_is_still_recorded(self) -> None:
        # Its errors were seen: the exposure counts, though the measurement fails.
        with repository(registry(FROZEN)) as root:
            code, output = run(root, "--record-holdout", "ACC-02", document=OFF_HOLDOUT)
            record = read_record(root)
        self.assertEqual(code, 1, output)
        self.assertIn("over holdout cases exceeds the threshold", output)
        self.assertEqual([item["exposures"] for item in record["exposures"]], [1])

    def refused(self, content: dict[str, Any], *args: str,
                prepare: Any = None, document: dict[str, Any] | None = EXACT) -> str:
        """The output of a refused --record-holdout, which leaves the record as it was."""
        with repository(content) as root:
            if prepare is not None:
                prepare(root)
            path = root / check_accuracy.EXPOSURES
            before = path.read_bytes() if path.is_file() else None
            code, output = run(root, *args, document=document)
            after = path.read_bytes() if path.is_file() else None
        self.assertEqual((code, after), (2, before), output)
        self.assertIn("accuracy: REFUSED (nothing measured; tests/accuracy/holdout_exposures.json "
                      "not written)", output)
        self.assertNotIn("worst", output)
        return output

    def test_a_draft_or_unknown_contract_is_refused(self) -> None:
        self.assertIn("REFUSED: ACC-02 is draft: its holdout stays sealed",
                      self.refused(registry(DRAFT), "--record-holdout", "ACC-02"))
        self.assertIn("REFUSED: ACC-09 is not a contract of tests/accuracy/contracts.json",
                      self.refused(registry(FROZEN), "--record-holdout", "ACC-09"))
        self.assertIn("REFUSED: --record-holdout needs --results",
                      self.refused(registry(FROZEN), "--record-holdout", "ACC-02",
                                   document=None))

    def test_a_second_first_exposure_is_refused(self) -> None:
        output = self.refused(registry(FROZEN), "--record-holdout", "ACC-02",
                              prepare=lambda root: expose(root, "ACC-02"))
        self.assertIn("REFUSED: second first exposure of ACC-02 refused: its holdout was exposed "
                      "under decision #1 with the results of commit aaaaaaaaaaaa (exposures 1)",
                      output)

    def test_a_decision_used_again_after_a_later_one_is_refused(self) -> None:
        def exposed_twice(root: Path) -> None:
            expose(root, "ACC-02")
            (root / check_accuracy.REGISTRY).write_text(
                json.dumps(registry({**FROZEN, "decision": "#2"})), "utf-8")
            expose(root, "ACC-02")
            (root / check_accuracy.REGISTRY).write_text(json.dumps(registry(FROZEN)), "utf-8")

        with repository(registry(FROZEN)) as root:
            exposed_twice(root)
            found = check_accuracy.check(root)[0]
        self.assertEqual(found, ["tests/accuracy/contracts.json: ACC-02: decision #1 was already "
                                 "used for exposure 1 (results of commit aaaaaaaaaaaa), before "
                                 "decision #2: a new exposure needs a new decision that states "
                                 "the earlier exposures"])
        output = self.refused(registry(FROZEN), "--record-holdout", "ACC-02",
                              prepare=exposed_twice)
        self.assertIn("decision #1 was already used for exposure 1", output)
        # The guard holds even without the registry check, which refuses first.
        with repository(registry(FROZEN)) as root:
            exposed_twice(root)
            read = check_accuracy.read_registry(root, None, Report(), "registry")
            record = check_accuracy.read_record(root, None, Report(), "record")
        assert read is not None and record is not None
        reused = check_accuracy.reused_decision(read.contracts[0], record.of("ACC-02"))
        self.assertEqual((reused.decision, reused.count) if reused else None, ("#1", 1))

    def test_a_missing_malformed_or_inconsistent_record_is_refused(self) -> None:
        def entry(root: Path) -> dict[str, Any]:
            expose(root, "ACC-02")
            item: dict[str, Any] = read_record(root)["exposures"][0]
            return item

        def remove(root: Path) -> None:
            (root / check_accuracy.EXPOSURES).unlink()

        def broken(root: Path) -> None:
            (root / check_accuracy.EXPOSURES).write_text("{broken", "utf-8")

        def miscounted(root: Path) -> None:
            write_record(root, {**EMPTY_RECORD, "exposures": [{**entry(root), "exposures": 2}]})

        def repeated(root: Path) -> None:
            item = entry(root)
            write_record(root, {**EMPTY_RECORD, "exposures": [item, {**item, "exposures": 2}]})

        def wrong_kind(root: Path) -> None:
            write_record(root, {**EMPTY_RECORD, "kind": "other"})

        cases = [
            (remove, "tests/accuracy/holdout_exposures.json is missing: restore it from Git"),
            (broken, "tests/accuracy/holdout_exposures.json: not valid UTF-8 JSON"),
            (wrong_kind, "envelope is not schema_version 1, kind 'holdout-exposures'"),
            (miscounted, "exposures is 2, but this is exposure 1 of ACC-02 in the record"),
            (repeated, "records a second first exposure of ACC-02 under decision #1"),
        ]
        for prepare, text in cases:
            with self.subTest(text=text):
                self.assertIn(text, self.refused(registry(FROZEN), "--record-holdout",
                                                 "ACC-02", prepare=prepare))

    def test_unusable_or_incomplete_results_are_refused(self) -> None:
        self.assertIn("the results are unusable, so nothing is measured: ",
                      self.refused(registry(FROZEN), "--record-holdout", "ACC-02",
                                   document={**EXACT, "commit": "b" * 40}))
        self.assertIn("ACC-02: the results hold no value for holdout output(s) demo-holdout-2 "
                      "total_exposure_amount, so nothing is measured",
                      self.refused(registry(FROZEN), "--record-holdout", "ACC-02",
                                   document=results(demo_calibration=100, demo_holdout_1=200)))

    def test_a_registry_with_a_finding_is_refused(self) -> None:
        self.assertIn("the check fails, so no holdout is exposed: ",
                      self.refused(registry(FROZEN, minimum=3), "--record-holdout", "ACC-02"))

    def test_a_new_decision_may_be_exposed_and_counts_the_earlier_exposures(self) -> None:
        with repository(registry(FROZEN)) as root:
            expose(root, "ACC-02")
            (root / check_accuracy.REGISTRY).write_text(json.dumps(registry(
                {**FROZEN, "decision": "#2", "threshold": 0.00005})), "utf-8")
            self.assertEqual(check_accuracy.check(root)[0], [])
            found, lines = measured_in(root)
            code, output = run(root, "--record-holdout", "ACC-02")
            record = read_record(root)
        self.assertTrue(found and "no exposure recorded under decision #2" in found[0], found)
        self.assertEqual(lines[1], "ACC-02 holdout: sealed until its exposure under decision #2 "
                                   "is recorded, exposures 1")
        self.assertEqual(code, 0, output)
        self.assertIn("threshold 5e-05, no target, exposures 2\n", output)
        self.assertEqual([(item["decision"], item["exposures"]) for item in record["exposures"]],
                         [("#1", 1), ("#2", 2)])


class ExposureDriftTest(unittest.TestCase):
    def drift(self, change: Any, contract: dict[str, Any] = FROZEN) -> list[str]:
        """The findings once ``change`` is applied to a repository whose ACC-02 holdout
        was exposed under ``contract``."""
        with repository(registry(contract)) as root:
            design = root / "docs/methodology/ACC-02-design.md"
            design.write_text("# Design\n", "utf-8")
            expose(root, "ACC-02")
            change(root)
            return check_accuracy.check(root)[0]

    def contract(self, **change: Any) -> Any:
        def apply(root: Path) -> None:
            (root / check_accuracy.REGISTRY).write_text(
                json.dumps(registry({**FROZEN, **change})), "utf-8")
        return apply

    def test_an_unchanged_contract_passes(self) -> None:
        self.assertEqual(self.drift(lambda root: None), [])

    def test_a_changed_entry_is_a_finding_unless_a_new_decision_states_it(self) -> None:
        found = self.drift(self.contract(provenance="Changed after the exposure"))
        self.assertEqual(len(found), 1, found)
        self.assertIn("ACC-02: its entry changed since its holdout was exposed under decision #1 "
                      "(exposure 1, results of commit aaaaaaaaaaaa): a change after an exposure "
                      "needs a new decision that states the earlier exposures", found[0])
        self.assertEqual(self.drift(self.contract(provenance="Changed", decision="#2")), [])
        # A changed target is a changed entry too, once the holdout was exposed.
        self.assertTrue(self.drift(self.contract(target=0.00001)))

    def test_a_lower_threshold_after_an_exposure_is_a_finding(self) -> None:
        found = self.drift(self.contract(threshold=0.00005))
        self.assertTrue(any("ACC-02: threshold 5e-05 is lower than 0.0001, the threshold when "
                            "its holdout was exposed: a threshold lowered after an exposure "
                            "needs a new decision" in item for item in found), found)
        self.assertEqual(self.drift(self.contract(threshold=0.00005, decision="#2")), [])

    def test_a_changed_holdout_file_is_a_finding(self) -> None:
        def fixture(root: Path) -> None:
            path = root / "tests/fixtures/demo-holdout-1.json"
            path.write_text(path.read_text("utf-8").replace("Synthetic", "Changed"), "utf-8")

        def expected(root: Path) -> None:
            path = root / "tests/expected/demo-holdout-2.json"
            path.write_text(path.read_text("utf-8").replace('"value": 0.0',
                                                            '"value": 0.001'), "utf-8")

        found = self.drift(fixture)
        self.assertEqual(len(found), 1, found)
        self.assertIn("ACC-02: the fixture of holdout case demo-holdout-1 changed since its "
                      "holdout was exposed", found[0])
        found = self.drift(expected)
        self.assertEqual(len(found), 1, found)
        self.assertIn("ACC-02: the expected file of holdout case demo-holdout-2 changed", found[0])
        # Under a new decision, the cases may change.
        def both(root: Path) -> None:
            fixture(root)
            self.contract(decision="#2")(root)
        self.assertEqual(self.drift(both), [])

    def test_a_design_may_only_be_appended_to(self) -> None:
        path = "docs/methodology/ACC-02-design.md"
        designed = {**FROZEN, "design": path}

        def append(root: Path) -> None:
            with (root / path).open("a", encoding="utf-8") as file:
                file.write("\nAmendment: a later note.\n")

        def rewrite(root: Path) -> None:
            (root / path).write_text("# Redesign\n", "utf-8")

        self.assertEqual(self.drift(append, designed), [])
        found = self.drift(rewrite, designed)
        self.assertEqual(len(found), 1, found)
        self.assertIn(f"ACC-02: design {path} no longer begins with the 9 bytes recorded since "
                      "its holdout was exposed under decision #1", found[0])
        self.assertIn("it was changed, not amended by appending", found[0])

    def test_a_design_is_a_file_where_the_excel_run_check_copies_it(self) -> None:
        for design in ("../outside.md", "/docs/methodology/x.md", "docs/other/x.md",
                       "docs/methodology/../x.md", "docs\\methodology\\x.md", 7):
            with self.subTest(design=design):
                found = findings(registry({**FROZEN, "design": design}))
                self.assertEqual(len(found), 1, found)
                self.assertIn("design is not a path to a file under docs/methodology/ or "
                              "tests/accuracy/", found[0])
        self.assertEqual(findings(registry({**FROZEN, "design": "docs/methodology/none.md"})),
                         ["tests/accuracy/contracts.json: ACC-02: design docs/methodology/none.md "
                          "is not a file"])

    def test_a_file_based_fixture_is_bound_file_by_file(self) -> None:
        with repository(None) as root:
            variant_case(root, {"base": 10.0})
            folder = root / "tests/fixtures/demo-variants"
            (folder / "base").mkdir()
            (folder / "base/positions.csv").write_bytes(b"id,amount\r\nP1,10\r\n")
            first = check_accuracy.fixture_sha256(root, "demo-variants")
            (folder / "base/positions.csv").write_bytes(b"id,amount\nP1,10\n")
            second = check_accuracy.fixture_sha256(root, "demo-variants")
            self.assertIsNone(check_accuracy.fixture_sha256(root, "missing-case"))
        self.assertNotEqual(first, second)


class RecordCheckTest(unittest.TestCase):
    def test_the_record_is_kept_beside_a_registry(self) -> None:
        with repository(registry(FROZEN)) as root:
            (root / check_accuracy.EXPOSURES).unlink()
            found = check_accuracy.check(root)[0]
        self.assertEqual(found, ["tests/accuracy/holdout_exposures.json: is missing: a "
                                 "repository with accuracy contracts keeps it, with an empty "
                                 "exposures list until a holdout is first measured. Restore it "
                                 "from Git: a record is never reset"])

    def test_a_malformed_record_is_a_finding(self) -> None:
        with repository(registry(FROZEN)) as root:
            expose(root, "ACC-02")
            good = read_record(root)["exposures"][0]
        cases: list[tuple[dict[str, Any], str]] = [
            ({"commit": "HEAD"}, "commit is not a full lower-case commit SHA"),
            ({"exposures": 0}, "exposures is not a whole number of at least 1"),
            ({"exposures": True}, "exposures is not a whole number of at least 1"),
            ({"contract": ["ACC-02"]}, "contract is not ACC-<number>"),
            ({"decision": {"id": 1}}, "decision is not a text"),
            ({"results": []}, "results is not a non-empty list"),
            ({"entry_sha256": "A" * 64}, "entry_sha256 is not a lower-case SHA-256"),
            ({"threshold": -1}, "threshold is not a finite number of at least 0"),
            ({"holdout": [good["holdout"][0]] * 2}, "holdout lists a case twice"),
            ({"holdout": [{"case": "x"}]}, "holdout is not a non-empty list"),
            ({"design": {"path": "x.md", "bytes": 1, "sha256": "0" * 64}},
             "design is not null or an object"),
            ({"note": "extra"}, "keys are not commit, contract, decision"),
        ]
        for change, text in cases:
            with self.subTest(text=text), repository(registry(FROZEN)) as root:
                write_record(root, {**EMPTY_RECORD, "exposures": [{**good, **change}]})
                found = check_accuracy.check(root)[0]
                self.assertEqual(len(found), 1, found)
                self.assertIn("tests/accuracy/holdout_exposures.json: exposures[0]", found[0])
                self.assertIn(text, found[0])
        envelopes: list[Any] = [{**EMPTY_RECORD, "exposures": {}}, [],
                                {**EMPTY_RECORD, "extra": 1}]
        for envelope in envelopes:
            with self.subTest(envelope=envelope), repository(registry(FROZEN)) as root:
                (root / check_accuracy.EXPOSURES).write_text(json.dumps(envelope), "utf-8")
                found = check_accuracy.check(root)[0]
                self.assertEqual(len(found), 1, found)
                self.assertIn("envelope is not schema_version 1, kind 'holdout-exposures'",
                              found[0])

    def test_an_exposed_contract_is_never_removed(self) -> None:
        with repository(registry(FROZEN)) as root:
            expose(root, "ACC-02")
            (root / check_accuracy.REGISTRY).write_text(json.dumps(registry()), "utf-8")
            removed = check_accuracy.check(root)[0]
            (root / check_accuracy.REGISTRY).unlink()
            deleted = check_accuracy.check(root)[0]
        for found in (removed, deleted):
            self.assertEqual(found, ["tests/accuracy/holdout_exposures.json: ACC-02 has recorded "
                                     "holdout exposures but no contract in "
                                     "tests/accuracy/contracts.json: a contract whose holdout "
                                     "was exposed is never removed"])

    def test_a_draft_after_an_exposure_stays_sealed_and_counts_it(self) -> None:
        with repository(registry(FROZEN)) as root:
            expose(root, "ACC-02")
            (root / check_accuracy.REGISTRY).write_text(json.dumps(registry(DRAFT)), "utf-8")
            found, lines = measured_in(root, OFF_HOLDOUT)
        self.assertEqual(found, [])
        self.assertEqual(lines[1], "ACC-02 holdout: sealed while the contract is draft, "
                                   "exposures 1")

    @unittest.skipUnless(shutil.which("git"), "needs Git")
    def test_the_exposures_recorded_at_a_tag_are_kept(self) -> None:
        with repository(registry(FROZEN)) as root:
            expose(root, "ACC-02")
            init(root)
            git(root, "add", "-A")
            git(root, "commit", "-qm", "Release")
            git(root, "tag", "-a", "v1.0.0", "-m", "v1.0.0")
            self.assertEqual(check_accuracy.check(root)[0], [])
            write_record(root, EMPTY_RECORD)
            (root / check_accuracy.REGISTRY).write_text(
                json.dumps(registry({**FROZEN, "decision": "#2"})), "utf-8")
            found = check_accuracy.check(root)[0]
            skipped = check_accuracy.check(root, widening=False)[0]
        self.assertEqual(found, ["tests/accuracy/holdout_exposures.json: the exposures recorded "
                                 "at v1.0.0 are not all kept, unchanged, at the start of the "
                                 "record: the record only grows"])
        self.assertEqual(skipped, [])


if __name__ == "__main__":
    unittest.main()
