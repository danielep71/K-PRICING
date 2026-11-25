"""Tests for tools/check_derivations.py, with synthetic cases and scripts."""

from __future__ import annotations

import contextlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_derivations  # noqa: E402

SCRIPT = "tests/derivations/total.py"


def output(**values: Any) -> dict[str, Any]:
    base: dict[str, Any] = {
        "quantity": "total_amount", "scope": None, "rule": None, "value": 100.0, "unit": "EUR",
        "tolerance": {"absolute": 0.01, "relative": 0},
        "reference": {"class": "illustrative", "source": None, "locator": None,
                      "derivation": "Synthetic", "script": SCRIPT, "derived_by": None,
                      "reviewed_by": None, "date": "2026-11-02"},
    }
    base.update(values)
    return base


def printing(outputs: object) -> str:
    """A derivation script that prints the given outputs."""
    return f"import json\nprint(json.dumps({{'outputs': {outputs!r}}}))\n"


def run(script: str | None, outputs: list[dict[str, Any]],
        extra: dict[str, str] | None = None) -> tuple[list[str], int]:
    with tempfile.TemporaryDirectory(prefix="derivations-") as directory:
        root = Path(directory)
        for path, text in {"tests/fixtures/case-one.json": "{}",
                           "tests/expected/case-one.json": json.dumps({"outputs": outputs}),
                           **(extra or {})}.items():
            (root / path).parent.mkdir(parents=True, exist_ok=True)
            (root / path).write_text(text, encoding="utf-8")
        if script is not None:
            (root / SCRIPT).parent.mkdir(parents=True, exist_ok=True)
            (root / SCRIPT).write_text(script, encoding="utf-8")
        return check_derivations.check(root)


class DerivationTest(unittest.TestCase):
    def assertFinding(self, result: tuple[list[str], int], text: str) -> None:
        self.assertTrue(any(text in finding for finding in result[0]), result[0])

    def test_repository_derivations_agree(self) -> None:
        self.assertEqual(check_derivations.check(ROOT)[0], [])

    def test_matching_value_within_tolerance_passes(self) -> None:
        script = printing([{"quantity": "total_amount", "scope": None, "value": 100.005}])
        self.assertEqual(run(script, [output()]), ([], 1))

    def test_relative_tolerance_governs_large_values(self) -> None:
        script = printing([{"quantity": "total_amount", "scope": None, "value": 1000100.0}])
        wanted = output(value=1000000.0, tolerance={"absolute": 0, "relative": 0.0001})
        self.assertEqual(run(script, [wanted]), ([], 1))
        self.assertFinding(run(script, [output(value=1000000.0, tolerance={
            "absolute": 0, "relative": 0.00001})]), "expected 1000000.0 within 10.0")

    def test_mismatch_fails(self) -> None:
        script = printing([{"quantity": "total_amount", "scope": None, "value": 101.0}])
        self.assertFinding(run(script, [output()]), "total_amount from tests/derivations/total.py: "
                                                    "value 101.0, expected 100.0")

    def test_record_scope_and_text(self) -> None:
        scope = {"dataset": "DS-01", "record": "P1"}
        script = printing([{"quantity": "total_amount", "scope": scope, "value": 100.0},
                           {"quantity": "run_status", "scope": None, "text": "OK"}])
        status = output(quantity="run_status", text="OK")
        for key in ("value", "unit", "tolerance"):
            status.pop(key)
        self.assertEqual(run(script, [output(scope=scope), status]), ([], 2))
        self.assertFinding(run(script, [output()]), "gives no total_amount")
        self.assertFinding(run(script, [output(scope={"dataset": "DS-01", "record": "P2"})]),
                           "gives no total_amount (DS-01/P2)")
        self.assertFinding(run(script, [{**status, "text": "REJECTED"}]),
                           "text 'OK', expected 'REJECTED'")

    def test_script_failures(self) -> None:
        cases = [
            (None, "is not an existing file under tests/derivations/"),
            ("import sys\nsys.exit('broken input')\n", "exited with 1: broken input"),
            ("print('not json')\n", "did not print JSON"),
            ("print('[]')\n", "must print an object with an outputs list"),
            (printing([{"quantity": "total_amount", "scope": None, "value": "100"}]),
             "is not a finite number"),
        ]
        for script, text in cases:
            with self.subTest(text=text):
                self.assertFinding(run(script, [output()]), text)

    def test_a_script_outside_tests_derivations_is_refused(self) -> None:
        # The script exists and agrees, and tests/derivations/ exists for the ".." path to
        # cross, so only the location rule can refuse it.
        script = printing([{"quantity": "total_amount", "scope": None, "value": 100.0}])
        self.assertEqual(run(script, [output()]), ([], 1))
        for path in ("tools/outside.py", "tests/derivations/../../tools/outside.py"):
            with self.subTest(path=path):
                reference = {**output()["reference"], "script": path}
                self.assertFinding(run(script, [output(reference=reference)],
                                       {"tools/outside.py": script}),
                                   "is not an existing file under tests/derivations/")

    def test_the_script_receives_the_fixture(self) -> None:
        script = ("import json, sys\nfixture = json.load(open(sys.argv[1]))\n"
                  "assert fixture == {}, fixture\n" + printing(
                      [{"quantity": "total_amount", "scope": None, "value": 100.0}]))
        self.assertEqual(run(script, [output()]), ([], 1))

    def test_values_without_a_script_are_not_compared(self) -> None:
        unscripted = output(reference={**output()["reference"], "script": None})
        self.assertEqual(run(None, [unscripted]), ([], 0))

    def test_file_based_case_runs_once_per_variant(self) -> None:
        # The script reads the variant directory it is given and prints its value.
        script = ("import json, sys\nfrom pathlib import Path\n"
                  "value = float((Path(sys.argv[1]) / 'total.txt').read_text())\n"
                  "print(json.dumps({'outputs': [{'quantity': 'total_amount', 'scope': None, "
                  "'value': value}]}))\n")

        def check(expected: dict[str, float]) -> tuple[list[str], int]:
            with tempfile.TemporaryDirectory(prefix="derivations-") as directory:
                root = Path(directory)
                files = {SCRIPT: script, "tests/fixtures/case-one/case.json": "{}",
                         "tests/fixtures/case-one/lf/total.txt": "100\n",
                         "tests/fixtures/case-one/crlf/total.txt": "200\r\n",
                         "tests/expected/case-one.json": json.dumps({"variants": {
                             variant: [output(value=value)]
                             for variant, value in expected.items()}})}
                for path, text in files.items():
                    (root / path).parent.mkdir(parents=True, exist_ok=True)
                    (root / path).write_bytes(text.encode())
                return check_derivations.check(root)

        self.assertEqual(check({"lf": 100.0, "crlf": 200.0}), ([], 2))
        findings, compared = check({"lf": 100.0, "crlf": 100.0})
        self.assertEqual(compared, 2)
        self.assertEqual(findings, ["tests/expected/case-one.json: variants.crlf: total_amount "
                                    "from tests/derivations/total.py: value 200.0, expected "
                                    "100.0 within 0.01"])
        # An expected variant without a directory is left to check_test_cases.py.
        self.assertEqual(check({"lf": 100.0, "gone": 1.0}), ([], 1))

    def test_main_reports_the_verdict(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()) as printed:
            self.assertEqual(check_derivations.main(["--root", str(ROOT)]), 0)
        # Any number of named scripts, so that the project may change its cases.
        self.assertIn("derivations: PASS (", printed.getvalue())
        with tempfile.TemporaryDirectory() as empty, \
                contextlib.redirect_stdout(io.StringIO()) as printed:
            self.assertEqual(check_derivations.main(["--root", empty]), 0)
        self.assertIn("no derivation script named", printed.getvalue())


class MalformedOutputTest(unittest.TestCase):
    def test_an_output_without_a_value_or_tolerance_is_left_to_the_format_check(self) -> None:
        # check_test_cases.py reports these; comparing them must not raise.
        no_relative = output(tolerance={"absolute": 0.01})
        no_value = {key: item for key, item in output().items() if key != "value"}
        no_tolerance = {key: item for key, item in output().items() if key != "tolerance"}
        for malformed in (no_relative, no_value, no_tolerance):
            with self.subTest(output=malformed):
                self.assertEqual(run(printing([output()]), [malformed]), ([], 0))


if __name__ == "__main__":
    unittest.main()
