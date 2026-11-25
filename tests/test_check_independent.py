"""Tests for tools/check_independent.py, with a fake implementation and its adapter.

Each test builds a throw-away repository with a source register that lists
IMPL-01, synthetic cases and an adapter registry. The fake implementation is
a small Python script that totals the numbers it reads; the adapter translates
a fixture into those numbers, runs it and reports its total. The values are
fixtures for the check only and validate nothing.
"""

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

import check_independent  # noqa: E402

ADAPTER = "tests/independent/total_adapter.py"
TOOL = "tool/total.py"
SOURCES = ("# Sources\n\n## Register\n\n| ID | Title |\n| --- | --- |\n"
           "| `IMPL-01` | Fake implementation |\n| `IMPL-02` | Another one |\n")
# The fake implementation: totals the numbers on its standard input, times FACTOR.
FAKE_TOOL = ("import sys\nFACTOR = {factor!r}\n"
             "print(sum(float(line) for line in sys.stdin if line.strip()) * FACTOR)\n")
# The adapter: translates a fixture into the tool's input and its total into an output.
FAKE_ADAPTER = """\
import json, subprocess, sys
from pathlib import Path
tool = Path("tool/total.py")
if not tool.is_file():
    print(json.dumps({"unavailable": "the fake tool is not installed"}))
    sys.exit()
fixture = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
records = fixture["datasets"].get("DS-01", [])
if any("notional_amount" not in record["inputs"] for record in records):
    print(json.dumps({"unsupported": "a record without a notional"}))
    sys.exit()
numbers = "".join(f"{record['inputs']['notional_amount']}\\n" for record in records)
total = subprocess.run([sys.executable, str(tool)], input=numbers, capture_output=True,
                       text=True, check=True).stdout
print(json.dumps({"outputs": [{"quantity": "total_amount", "scope": None,
                               "value": float(total)}]}))
"""


def fixture(*amounts: float | None) -> dict[str, Any]:
    records = [{"id": f"P{index}", "inputs": {"currency_code": "EUR"} | (
        {} if amount is None else {"notional_amount": amount})}
        for index, amount in enumerate(amounts, 1)]
    return {"schema_version": 1, "kind": "test-fixture", "id": "case-one", "synthetic": True,
            "description": "Synthetic", "category": "nominal", "calculation": "CALC-02",
            "reporting_date": "2026-01-02", "reporting_currency_code": "EUR",
            "datasets": {"DS-01": records}}


def output(value: float = 100.0, source: str | None = "IMPL-01",
           **values: Any) -> dict[str, Any]:
    base: dict[str, Any] = {
        "quantity": "total_amount", "scope": None, "rule": "CALC-02.1", "value": value,
        "unit": "EUR", "tolerance": {"absolute": 0.01, "relative": 0},
        "reference": {"class": "independent" if source else "illustrative", "source": source,
                      "locator": "Run", "derivation": "Synthetic", "script": None,
                      "derived_by": "Fake 1.0", "reviewed_by": None, "date": "2026-11-03"},
    }
    base.update(values)
    return base


def adapters(*entries: dict[str, Any]) -> dict[str, Any]:
    return {"schema_version": 1, "kind": "independent-adapters", "adapters": list(entries)}


ENTRY = {"source": "IMPL-01", "command": ["python", ADAPTER]}


def run(registry: object, outputs: list[dict[str, Any]] | None = None,
        case: dict[str, Any] | None = None, factor: float | None = 1.0,
        adapter: str = FAKE_ADAPTER,
        extra: dict[str, str] | None = None) -> check_independent.Outcome:
    """The check on one case, ``case-one``, with the fake tool unless ``factor`` is None."""
    with tempfile.TemporaryDirectory(prefix="independent-") as directory:
        root = Path(directory)
        files = {"docs/methodology/SOURCES.md": SOURCES, ADAPTER: adapter,
                 "tests/fixtures/case-one.json": json.dumps(case or fixture(60.0, 40.0)),
                 "tests/expected/case-one.json": json.dumps(
                     {"outputs": [output()] if outputs is None else outputs}),
                 **({} if factor is None else {TOOL: FAKE_TOOL.format(factor=factor)}),
                 **({} if registry is None else {
                     "tests/independent/adapters.json": json.dumps(registry)}),
                 **(extra or {})}
        for path, text in files.items():
            (root / path).parent.mkdir(parents=True, exist_ok=True)
            (root / path).write_text(text, encoding="utf-8")
        return check_independent.check(root)


class IndependentTest(unittest.TestCase):
    def assertFinding(self, outcome: check_independent.Outcome, text: str) -> None:
        self.assertTrue(any(text in finding for finding in outcome.findings), outcome.findings)

    def test_repository_check_passes(self) -> None:
        outcome = check_independent.check(ROOT)
        self.assertEqual(outcome.findings, [])

    def test_without_an_adapter_the_check_is_skipped_and_says_so(self) -> None:
        for registry in (None, adapters()):
            with self.subTest(registry=registry):
                outcome = run(registry)
                self.assertEqual((outcome.findings, outcome.adapters, outcome.compared),
                                 ([], 0, 0))
                self.assertIn("no adapter registered", check_independent.summary(outcome))
        with tempfile.TemporaryDirectory() as empty, \
                contextlib.redirect_stdout(io.StringIO()) as printed:
            self.assertEqual(check_independent.main(["--root", empty]), 0)
        self.assertIn("independent-implementations: PASS (no adapter registered in "
                      "tests/independent/adapters.json; skipped, nothing cross-checked)",
                      printed.getvalue())

    def test_agreeing_outputs_pass(self) -> None:
        outcome = run(adapters(ENTRY))
        self.assertEqual((outcome.findings, outcome.compared, outcome.skipped,
                          outcome.unsupported), ([], 1, [], []))
        # Within max(absolute, relative * |expected|), the rule of the derivation check.
        self.assertEqual(run(adapters(ENTRY), factor=1.00005).findings, [])
        relative = output(tolerance={"absolute": 0, "relative": 0.001})
        self.assertEqual(run(adapters(ENTRY), [relative], factor=1.0009).findings, [])

    def test_disagreeing_outputs_fail(self) -> None:
        outcome = run(adapters(ENTRY), factor=1.01)
        self.assertEqual(outcome.compared, 1)
        self.assertEqual(outcome.findings, ["tests/expected/case-one.json: total_amount from "
                                            "IMPL-01: value 101.0, expected 100.0 within 0.01"])
        relative = output(tolerance={"absolute": 0, "relative": 0.001})
        self.assertFinding(run(adapters(ENTRY), [relative], factor=1.0011),
                           "expected 100.0 within 0.1")

    def test_values_of_other_classes_are_compared_when_printed(self) -> None:
        self.assertFinding(run(adapters(ENTRY), [output(source=None)], factor=2.0),
                           "value 200.0, expected 100.0")
        # A quantity the adapter does not print is compared only when it cites the source.
        other = output(quantity="other_amount", source=None)
        self.assertEqual(run(adapters(ENTRY), [output(), other]).compared, 1)
        self.assertFinding(run(adapters(ENTRY), [output(quantity="other_amount")]),
                           "IMPL-01 gives no other_amount, which cites it")
        # Every value citing the source must be printed, whatever its class.
        cited = output(quantity="other_amount")
        cited["reference"]["class"] = "illustrative"
        self.assertFinding(run(adapters(ENTRY), [output(), cited]),
                           "IMPL-01 gives no other_amount, which cites it")
        # A value citing another source is not this adapter's to recompute.
        self.assertEqual(run(adapters(ENTRY), [output(quantity="other_amount",
                                                      source="IMPL-02")]).findings, [])

    def test_unrepresentable_fixtures_are_listed_with_the_reason(self) -> None:
        outcome = run(adapters(ENTRY), case=fixture(60.0, None))
        self.assertEqual((outcome.findings, outcome.compared), ([], 0))
        self.assertEqual(outcome.unsupported, [
            "IMPL-01 cannot represent case-one: a record without a notional (1 value(s) "
            "citing it not recomputed)"])
        self.assertIn("1 input(s) not representable", check_independent.summary(outcome))

    def test_an_absent_tool_is_skipped_with_a_note(self) -> None:
        outcome = run(adapters(ENTRY), factor=None)
        self.assertEqual((outcome.findings, outcome.compared), ([], 0))
        self.assertEqual(outcome.skipped, ["IMPL-01 skipped: the fake tool is not installed"])
        missing = {"source": "IMPL-01", "command": ["no-such-program-for-this-test", "x"]}
        outcome = run(adapters(missing))
        self.assertEqual(outcome.findings, [])
        self.assertEqual(outcome.skipped, ["IMPL-01 skipped: no-such-program-for-this-test is "
                                           "not on PATH"])
        self.assertIn("1 adapter(s), 1 skipped", check_independent.summary(outcome))

    def test_unusable_replies_fail(self) -> None:
        cases = [
            ("import sys\nsys.exit('tool crashed')\n", "exited with 1: tool crashed"),
            ("print('not json')\n", "did not print JSON"),
            ("print('[]')\n", "must print an object with one of outputs, unavailable, "
                              "unsupported"),
            ("print('{\"outputs\": [], \"unsupported\": \"x\"}')\n", "must print an object"),
            ("print('{\"outputs\": {}}')\n", "outputs is not a list of objects"),
            ("print('{\"unsupported\": \" \"}')\n", "unsupported is not a reason"),
            ("import json\nprint(json.dumps({'outputs': [{'quantity': 'total_amount', "
             "'scope': None, 'value': 'x'}]}))\n", "is not a finite number"),
        ]
        for adapter, text in cases:
            with self.subTest(text=text):
                self.assertFinding(run(adapters(ENTRY), adapter=adapter), text)

    def test_registry_problems_fail(self) -> None:
        cases = [
            ({"schema_version": 1, "kind": "accuracy-contracts", "adapters": []},
             "envelope is not"),
            (adapters({"source": "IMPL-01"}), "adapters[0]: keys are not command, source"),
            (adapters({**ENTRY, "source": "SRC-01"}), "source is not an IMPL- ID"),
            (adapters({**ENTRY, "source": "IMPL-01\n"}), "source is not an IMPL- ID"),
            (adapters({**ENTRY, "source": "IMPL-0\u0661"}), "source is not an IMPL- ID"),
            (adapters({**ENTRY, "source": "IMPL-09"}), "IMPL-09 is not in the register"),
            (adapters(ENTRY, ENTRY), "adapters[1]: source IMPL-01 has a second adapter"),
            (adapters({**ENTRY, "command": []}), "command is not a non-empty list of words"),
            (adapters({**ENTRY, "command": ["python", ""]}), "command is not"),
        ]
        for registry, text in cases:
            with self.subTest(text=text):
                self.assertFinding(run(registry), text)
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / check_independent.REGISTRY
            path.parent.mkdir(parents=True)
            path.write_text('{"adapters": [], "adapters": []}', encoding="utf-8")
            self.assertTrue(any("duplicate key" in finding for finding in
                                check_independent.check(Path(directory)).findings))

    def test_file_based_case_runs_once_per_variant(self) -> None:
        # The adapter reads the total in the variant directory it is given.
        adapter = ("import json, sys\nfrom pathlib import Path\n"
                   "value = float((Path(sys.argv[1]) / 'total.txt').read_text())\n"
                   "print(json.dumps({'outputs': [{'quantity': 'total_amount', "
                   "'scope': None, 'value': value}]}))\n")
        expected = {"variants": {"lf": [output(100.0)], "crlf": [output(100.0)]}}
        outcome = run(adapters(ENTRY), adapter=adapter, extra={
            "tests/expected/case-one.json": json.dumps(expected),
            "tests/fixtures/case-one/case.json": "{}",
            "tests/fixtures/case-one/lf/total.txt": "100\n",
            "tests/fixtures/case-one/crlf/total.txt": "200\r\n"})
        self.assertEqual(outcome.compared, 2)
        self.assertEqual(outcome.findings, [
            "tests/expected/case-one.json: variants.crlf: total_amount from IMPL-01: value "
            "200.0, expected 100.0 within 0.01"])


if __name__ == "__main__":
    unittest.main()
