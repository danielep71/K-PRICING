"""Tests for tools/check_verification_depth.py.

Each rule is tested on a fixture inventory, check list and set of collected
tests, never on the live files, so that a change to the repository's checks
or tests cannot change what these tests prove.
"""

from __future__ import annotations

import contextlib
import copy
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_repository  # noqa: E402
import check_verification_depth as depth  # noqa: E402
from check_repository import Check, Context, Finding, Probe  # noqa: E402


def built_in(context: Context) -> list[Finding]:
    """A built-in check of the fixture."""
    return []


CLAIM = "a fixture input with a known defect is rejected with its finding"
LISTED = [
    Check("unit", [sys.executable, "tools/unit.py", "--strict"]),
    Check("built-in", function=built_in),
    Check("lint", ["/opt/bin/lint", "-q"], probe=Probe(["/opt/bin/lint", "bad.yml"], "bad key")),
    Check("advice", [sys.executable, "tools/advice.py"], advisory=True),
]
COLLECTED = depth.Collected({"test_unit.UnitTest.test_rejects_a_bad_input",
                             "test_unit.UnitTest.test_rejects_a_bad_option",
                             "test_lint.LintTest.test_rejects_a_bad_key"})


def inventory() -> dict[str, Any]:
    """A complete inventory of LISTED, proved by COLLECTED."""
    return {"schema_version": 1, "checks": [
        {"check": "unit", "checker": "python tools/unit.py --strict",
         "proofs": ["tests.test_unit.UnitTest.test_rejects_a_bad_input"],
         "negative_claim": CLAIM, "state": "green"},
        {"check": "built-in", "checker": "tools/check_repository.py:built_in",
         "proofs": ["tests.test_unit.UnitTest.test_rejects_a_bad_option"],
         "negative_claim": CLAIM, "state": "green"},
        {"check": "lint", "checker": "lint -q",
         "proofs": ["probe", "tests.test_lint.LintTest.test_rejects_a_bad_key"],
         "negative_claim": CLAIM, "state": "green_in_ci"},
    ]}


def damaged(change: Any) -> dict[str, Any]:
    """The fixture inventory after change(inventory) has damaged it."""
    data = copy.deepcopy(inventory())
    change(data)
    return data


def entry(data: dict[str, Any], name: str) -> dict[str, Any]:
    found: dict[str, Any] = next(item for item in data["checks"] if item["check"] == name)
    return found


class InventoryTest(unittest.TestCase):
    def problems(self, data: Any, listed: list[Check] | None = None,
                 collected: depth.Collected | None = None) -> list[str]:
        return depth.validate(data, LISTED if listed is None else listed,
                              COLLECTED if collected is None else collected)

    def test_a_complete_inventory_passes(self) -> None:
        # The probe proves lint, and the advisory check needs no entry.
        self.assertEqual(self.problems(inventory()), [])

    def test_a_removed_entry_fails_and_names_its_check(self) -> None:
        data = damaged(lambda data: data["checks"].pop(1))
        self.assertEqual(self.problems(data), [
            "built-in: release-blocking check has no entry in tools/verification_depth.json"])

    def test_a_check_added_without_an_entry_fails(self) -> None:
        listed = [*LISTED, Check("added", [sys.executable, "tools/added.py"])]
        self.assertEqual(self.problems(inventory(), listed), [
            "added: release-blocking check has no entry in tools/verification_depth.json"])

    def test_an_entry_for_a_check_that_does_not_exist_fails(self) -> None:
        data = damaged(lambda data: entry(data, "unit").update(check="retired"))
        self.assertEqual(self.problems(data), [
            "retired: no such check in tools/check_repository.py; remove its entry",
            "unit: release-blocking check has no entry in tools/verification_depth.json"])

    def test_an_entry_for_an_advisory_check_fails(self) -> None:
        data = damaged(lambda data: data["checks"].append(
            {**entry(data, "unit"), "check": "advice", "checker": "python tools/advice.py"}))
        self.assertEqual(self.problems(data), [
            "advice: the check is advisory and cannot fail the run; remove its entry"])

    def test_a_duplicate_entry_fails(self) -> None:
        data = damaged(lambda data: data["checks"].append(copy.deepcopy(data["checks"][0])))
        self.assertEqual(self.problems(data), ["unit: listed twice, as checks[0] and checks[3]"])

    def test_a_missing_or_empty_field_fails(self) -> None:
        for key in ("checker", "proofs", "negative_claim", "state"):
            values: tuple[Any, ...] = (None, "", " ", [])
            for value in values:
                with self.subTest(key=key, value=value):
                    def change(data: dict[str, Any], key: str = key, value: Any = value) -> None:
                        if value is None:
                            del entry(data, "unit")[key]
                        else:
                            entry(data, "unit")[key] = value
                    self.assertEqual(self.problems(damaged(change)),
                                     [f"unit: {key} is missing or empty"])
        for value in (None, "", "  "):
            with self.subTest(key="check", value=value):
                data = damaged(lambda data, value=value: data["checks"][0].update(check=value))
                self.assertEqual(self.problems(data)[0], "checks[0]: check is missing or empty")

    def test_a_field_of_the_wrong_type_or_an_unknown_field_fails(self) -> None:
        cases = {"checker": (["python", "tools/unit.py"], "unit: checker is not a text"),
                 "state": (1, "unit: state is not a text"),
                 "proofs": ("tests.test_unit.UnitTest.test_rejects_a_bad_input",
                            "unit: proofs is not a list"),
                 "note": ("extra", "unit: unknown field(s) note")}
        for key, (value, message) in cases.items():
            with self.subTest(key=key):
                data = damaged(lambda data, key=key, value=value:
                               entry(data, "unit").update({key: value}))
                self.assertEqual(self.problems(data), [message])
        data = damaged(lambda data: data["checks"].insert(0, "unit"))
        self.assertEqual(self.problems(data), ["checks[0] is not an object"])

    def test_a_state_outside_the_closed_list_fails(self) -> None:
        for state in ("red", "Green", "green_with_debt"):
            with self.subTest(state=state):
                data = damaged(lambda data, state=state: entry(data, "unit").update(state=state))
                self.assertEqual(self.problems(data), [
                    f"unit: state {state!r} is not one of green, green_in_ci"])

    def test_a_short_negative_claim_fails(self) -> None:
        short = "x" * (depth.MIN_CLAIM - 1)
        data = damaged(lambda data: entry(data, "unit").update(negative_claim=f"  {short} "))
        self.assertEqual(self.problems(data), [
            f"unit: negative_claim is too short to be specific ({depth.MIN_CLAIM - 1} "
            f"characters, at least {depth.MIN_CLAIM})"])
        data = damaged(lambda data: entry(data, "unit").update(negative_claim=short + "x"))
        self.assertEqual(self.problems(data), [])

    def test_a_checker_that_differs_from_the_check_fails(self) -> None:
        cases = {"unit": ("python tools/unit.py", "python tools/unit.py --strict"),
                 "built-in": ("tools/check_repository.py:renamed",
                              "tools/check_repository.py:built_in")}
        for name, (written, runs) in cases.items():
            with self.subTest(check=name):
                data = damaged(lambda data, name=name, written=written:
                               entry(data, name).update(checker=written))
                self.assertEqual(self.problems(data), [
                    f"{name}: checker {written!r} is not what the check runs, {runs!r}"])

    def test_a_renamed_or_deleted_test_fails(self) -> None:
        proof = "tests.test_unit.UnitTest.test_rejects_a_bad_input"
        cases = {
            "method": (COLLECTED.ids - {proof.removeprefix("tests.")}, {},
                       "method is not collected from class UnitTest"),
            "class": ({"test_unit.OtherTest.test_rejects_a_bad_option",
                       "test_lint.LintTest.test_rejects_a_bad_key"}, {},
                      "class UnitTest is not collected from tests/test_unit.py"),
            "module": ({"test_lint.LintTest.test_rejects_a_bad_key"}, {},
                       "module tests/test_unit.py is not collected by python -m unittest "
                       "discover -s tests"),
            "import": ({"test_lint.LintTest.test_rejects_a_bad_key"},
                       {"test_unit": "No module named 'gone'"},
                       "tests/test_unit.py cannot be imported: No module named 'gone'"),
        }
        for name, (ids, failed, message) in cases.items():
            with self.subTest(removed=name):
                problems = self.problems(inventory(), collected=depth.Collected(ids, failed))
                self.assertIn(f"unit: proof {proof}: {message}", problems)

    def test_a_proof_that_is_neither_a_test_id_nor_the_probe_fails(self) -> None:
        for proof in ("test_unit.UnitTest.test_rejects_a_bad_input", "tests/test_unit.py",
                      "tests.test_unit.UnitTest", "tests.test_unit.UnitTest.rejects",
                      "tests.helpers.UnitTest.test_rejects_a_bad_input", "Probe"):
            with self.subTest(proof=proof):
                data = damaged(lambda data, proof=proof: entry(data, "unit").update(
                    proofs=[proof]))
                self.assertEqual(self.problems(data), [
                    f"unit: proof {proof!r} is neither 'probe' nor a unittest id "
                    "tests.<module>.<Class>.<test_method>"])
        data = damaged(lambda data: entry(data, "unit").update(proofs=["probe"]))
        self.assertEqual(self.problems(data), ["unit: proof 'probe', but the check has no probe"])
        data = damaged(lambda data: entry(data, "unit").update(proofs=[""]))
        self.assertEqual(self.problems(data), ["unit: a proof is not a non-empty text"])

    def test_a_proof_listed_twice_fails(self) -> None:
        data = damaged(lambda data: entry(data, "lint")["proofs"].append("probe"))
        self.assertEqual(self.problems(data), ["lint: proof probe is listed twice"])

    def test_a_malformed_envelope_fails(self) -> None:
        expected = ["the inventory must be an object with exactly schema_version 1 and a checks "
                    "list"]
        malformed: tuple[Any, ...] = ([], {"checks": []}, {"schema_version": 2, "checks": []},
                     {"schema_version": True, "checks": []}, {"schema_version": 1, "checks": {}},
                     {"schema_version": 1, "checks": [], "note": "x"})
        for data in malformed:
            with self.subTest(data=data):
                self.assertEqual(self.problems(data), expected)


class CollectTest(unittest.TestCase):
    def test_the_tests_that_discovery_collects_and_the_modules_it_cannot_import(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "tests").mkdir()
            (root / "tests/test_fixture_depth.py").write_text(
                "import unittest\n\n\nclass FixtureTest(unittest.TestCase):\n"
                "    def test_one(self):\n        pass\n\n    def helper(self):\n        pass\n\n"
                "\nclass NotATest:\n    def test_ignored(self):\n        pass\n",
                encoding="utf-8")
            # Discovery reads only test*.py files.
            (root / "tests/fixture_helper.py").write_text(
                "import unittest\n\n\nclass HelperTest(unittest.TestCase):\n"
                "    def test_hidden(self):\n        pass\n", encoding="utf-8")
            (root / "tests/test_fixture_broken.py").write_text("import a_missing_module\n",
                                                               encoding="utf-8")
            collected = depth.collect(root)
        self.assertEqual(collected.ids, {"test_fixture_depth.FixtureTest.test_one"})
        self.assertEqual(list(collected.failed), ["test_fixture_broken"])
        self.assertIn("a_missing_module", collected.failed["test_fixture_broken"])

    def test_without_a_tests_folder_collection_fails(self) -> None:
        with tempfile.TemporaryDirectory() as directory, \
                self.assertRaisesRegex(RuntimeError, "the tests cannot be collected: "):
            depth.collect(Path(directory))


class MainTest(unittest.TestCase):
    def run_main(self, text: str | None) -> tuple[int, str]:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "tests").mkdir()
            (root / "tests/test_unit.py").write_text(
                "import unittest\n\n\nclass UnitTest(unittest.TestCase):\n"
                "    def test_rejects_a_bad_input(self):\n        pass\n\n"
                "    def test_rejects_a_bad_option(self):\n        pass\n", encoding="utf-8")
            (root / "tests/test_lint.py").write_text(
                "import unittest\n\n\nclass LintTest(unittest.TestCase):\n"
                "    def test_rejects_a_bad_key(self):\n        pass\n", encoding="utf-8")
            if text is not None:
                (root / "tools").mkdir()
                (root / depth.INVENTORY).write_text(text, encoding="utf-8")
            with contextlib.redirect_stdout(io.StringIO()) as output:
                code = depth.main(["--root", str(root)], lambda root: LISTED)
        return code, output.getvalue()

    def test_main_reports_the_verdict(self) -> None:
        self.assertEqual(self.run_main(json.dumps(inventory())), (
            0, "verification-depth: PASS (3 release-blocking check(s); the tests check runs "
               "the proofs)\n"))
        code, printed = self.run_main(json.dumps(damaged(lambda data: data["checks"].pop())))
        self.assertEqual(code, 1)
        self.assertEqual(printed.splitlines(), [
            "FAIL: lint: release-blocking check has no entry in tools/verification_depth.json",
            "verification-depth: FAIL (3 release-blocking check(s); the tests check runs the "
            "proofs)"])

    def test_an_unreadable_inventory_fails(self) -> None:
        for text, message in ((None, "tools/verification_depth.json is missing"),
                              ('{"schema_version": 1, "schema_version": 1, "checks": []}',
                               "tools/verification_depth.json is not valid JSON: duplicate key "
                               "'schema_version'")):
            with self.subTest(message=message):
                code, printed = self.run_main(text)
                self.assertEqual(code, 1)
                self.assertEqual(printed.splitlines()[0], f"FAIL: {message}")


class RepositoryTest(unittest.TestCase):
    def test_the_required_checks_are_those_ci_runs_less_the_advisory_ones(self) -> None:
        listed = depth.listed_in_ci(ROOT)
        names = [check.name for check in check_repository.checks(Context(ROOT, ci=True))]
        self.assertEqual([check.name for check in listed], names)
        self.assertIn("clean-working-tree", names)
        self.assertIn("verification-depth", names)
        advisory = {check.name for check in listed if check.advisory}
        self.assertIn("vba-calls", advisory)
        # An inventory of every other check needs no entry for an advisory one.
        probed = [check for check in listed if not check.advisory and check.probe is not None]
        data = {"schema_version": 1, "checks": [
            {"check": check.name, "checker": depth.checker_of(check), "proofs": ["probe"],
             "negative_claim": CLAIM, "state": "green"} for check in probed]}
        problems = depth.validate(data, listed, depth.Collected())
        entered = {check.name for check in probed}
        self.assertEqual(problems, [f"{name}: release-blocking check has no entry in "
                                    "tools/verification_depth.json"
                                    for name in names if name not in advisory | entered])

    def test_a_checker_is_shown_as_the_integrity_checks_table_shows_it(self) -> None:
        listed = {check.name: check for check in depth.listed_in_ci(ROOT)}
        self.assertEqual(depth.checker_of(listed["changelog"]),
                         "python tools/check_changelog.py --repository OWNER/REPOSITORY")
        self.assertEqual(depth.checker_of(listed["actionlint"]), "actionlint -no-color")
        self.assertEqual(depth.checker_of(listed["secret-files"]),
                         "tools/check_repository.py:secret_files")
        contributing = (ROOT / "CONTRIBUTING.md").read_text(encoding="utf-8")
        self.assertIn("| `verification-depth` | `python tools/check_verification_depth.py` |",
                      contributing)


if __name__ == "__main__":
    unittest.main()
