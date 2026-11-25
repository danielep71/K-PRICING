"""Valid and invalid test-case fixtures for tools/check_test_cases.py.

Each test writes a temporary repository holding minimal methodology guides,
one fixture and its expected file, then changes one thing. The data is
synthetic and has no regulatory meaning.
"""

from __future__ import annotations

import contextlib
import copy
import io
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_test_cases  # noqa: E402

RULES = "**Rules:**\n\n| Step | Rule | Source |\n| --- | --- | --- |\n"
# Steps 3 and 4 are cells of other tables, and step 5 is outside any contract.
CONTRACTS = ("# Contracts\n\n### `CALC-02` Real calculation\n\n**Inputs:**\n\n"
             "| Input | Step |\n| --- | --- |\n| Amount | 3 |\n\n" + RULES
             + "| 1 | Sum | `SRC-02` |\n| 2 | Average | `SRC-02` |\n\n**Outputs:**\n\n"
             "| Output | Step |\n| --- | --- |\n| Total | 4 |\n\n"
             "### *Illustrative:* `CALC-01` Example\n\n" + RULES + "| 1 | Sum | `SRC-01` |\n"
             "\n### `CALC-03` Not written yet\n\n## Decisions\n\n" + RULES + "| 5 | x | y |\n")
DATA = ("# Data\n\n## Datasets\n\n| ID | Description |\n| --- | --- |\n"
        "| `DS-02` | Real positions |\n| *Illustrative:* `DS-01` | Example |\n\n"
        "## Fields\n\n| `DS-09` | Not in the Datasets section |\n")
SOURCES = ("# Sources\n\n## Register\n\n| ID | Title |\n| --- | --- |\n"
           "| `SRC-02` | A regulation |\n| `IMPL-01` | An external implementation |\n"
           "| *Illustrative:* `SRC-01` | Example |\n")
FIXTURE: dict[str, Any] = {
    "schema_version": 1, "kind": "test-fixture", "id": "case-one", "synthetic": True,
    "description": "A synthetic case", "category": "nominal", "calculation": "CALC-02",
    "reporting_date": "2026-01-02", "reporting_currency_code": "EUR",
    "datasets": {"DS-02": [
        {"id": "P1", "inputs": {"currency_code": "EUR", "notional_amount": 600000,
                                "maturity_years": 2.5, "trade_count": 3, "start_date": "2025-01-02",
                                "netted_flag": False, "position_id": "X1", "risk_weight": 0.2}},
        {"id": "P2", "inputs": {"currency_code": "EUR", "notional_amount": 400000}},
    ]},
}
OUTPUT: dict[str, Any] = {
    "quantity": "total_notional_amount", "scope": None, "rule": "CALC-02.1",
    "value": 1000000, "unit": "EUR", "tolerance": {"absolute": 0.000001, "relative": 0},
    "reference": {"class": "published", "source": "SRC-02", "locator": "Article 1(1)",
                  "derivation": None, "script": None, "derived_by": None,
                  "reviewed_by": None, "date": "2026-11-02"},
}
EXPECTED: dict[str, Any] = {"schema_version": 1, "kind": "test-expected", "fixture": "case-one",
                            "outputs": [OUTPUT]}

Edit = Callable[[dict[str, Any], dict[str, Any]], None]


def write(root: Path, edit: Edit | None = None, name: str = "case-one",
          extra: dict[str, str] | None = None) -> None:
    """Write a repository with one case, after applying edit(fixture, expected)."""
    fixture, expected = copy.deepcopy(FIXTURE), copy.deepcopy(EXPECTED)
    if edit is not None:
        edit(fixture, expected)
    files = {"docs/methodology/CALCULATION_CONTRACTS.md": CONTRACTS,
             "docs/methodology/DATA_CONTRACT.md": DATA,
             "docs/methodology/SOURCES.md": SOURCES,
             f"tests/fixtures/{name}.json": json.dumps(fixture),
             f"tests/expected/{name}.json": json.dumps(expected)}
    files.update(extra or {})
    for path, text in files.items():
        (root / path).parent.mkdir(parents=True, exist_ok=True)
        (root / path).write_text(text, encoding="utf-8")


def run(edit: Edit | None = None, name: str = "case-one",
        extra: dict[str, str] | None = None) -> list[str]:
    """Write the repository, apply edit(fixture, expected), and check it."""
    with tempfile.TemporaryDirectory(prefix="test-cases-") as directory:
        write(Path(directory), edit, name, extra)
        return check_test_cases.check(Path(directory))[0]


def output(expected: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = expected["outputs"][0]
    return result


class ValidCasesTest(unittest.TestCase):
    def test_valid_case_passes(self) -> None:
        self.assertEqual(run(), [])

    def test_repository_cases_pass(self) -> None:
        # Any number of cases, none included: the project removes the
        # illustrative case and adds its own.
        self.assertEqual(check_test_cases.check(ROOT)[0], [])

    def test_no_cases_pass_without_reading_the_guides(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            self.assertEqual(check_test_cases.check(Path(directory)), ([], 0, 0))

    def test_record_scope_and_text_output(self) -> None:
        def edit(fixture: dict[str, Any], expected: dict[str, Any]) -> None:
            output(expected)["scope"] = {"dataset": "DS-02", "record": "P2"}
            status = {key: OUTPUT[key] for key in ("quantity", "scope", "rule", "reference")}
            expected["outputs"].append({**status, "quantity": "run_status", "text": "OK"})
        self.assertEqual(run(edit), [])

    def test_illustrative_case_may_use_illustrative_entries(self) -> None:
        def edit(fixture: dict[str, Any], expected: dict[str, Any]) -> None:
            fixture.update(id="illustrative-one", calculation="CALC-01")
            fixture["datasets"] = {"DS-01": fixture["datasets"].pop("DS-02")}
            expected["fixture"] = "illustrative-one"
            output(expected).update(quantity="illustrative_total_amount", rule=None)
            output(expected)["reference"].update({"class": "illustrative", "source": None,
                                                  "locator": None,
                                                  "derivation": "Sum of the two notionals"})
        self.assertEqual(run(edit, name="illustrative-one"), [])

    def test_independent_reference(self) -> None:
        def edit(fixture: dict[str, Any], expected: dict[str, Any]) -> None:
            output(expected)["reference"].update({
                "class": "independent", "source": "IMPL-01", "locator": "Run of 2026-11-02",
                "derivation": "Computed by the external tool", "derived_by": "Tool 1.0"})
        self.assertEqual(run(edit), [])


class FindingTest(unittest.TestCase):
    def assertFinding(self, findings: list[str], text: str) -> None:
        self.assertTrue(any(text in finding for finding in findings), findings)


class FileTest(FindingTest):
    def test_strict_json(self) -> None:
        for raw, text in (('{"a": 1, "a": 2}', "duplicate key"), ('{"a": NaN}', "non-standard"),
                          ('{"a": 1e400}', "not finite"), ("[", "not valid UTF-8 JSON")):
            with self.subTest(raw=raw):
                self.assertFinding(run(extra={"tests/fixtures/case-one.json": raw}), text)

    def test_fixture_must_be_an_object(self) -> None:
        self.assertFinding(run(extra={"tests/fixtures/case-one.json": "null"}), "is not an object")

    def test_expected_needs_a_fixture_and_the_reverse(self) -> None:
        self.assertFinding(run(extra={"tests/expected/other.json": "{}"}),
                           "for an existing fixture")
        orphan = json.dumps({**FIXTURE, "id": "orphan"})
        self.assertFinding(run(extra={"tests/fixtures/orphan.json": orphan}), "has no expected file")

    def test_missing_guide(self) -> None:
        self.assertFinding(run(extra={"docs/methodology/SOURCES.md": ""}),
                           "source 'SRC-02' is not in the register")
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for folder, data in (("tests/fixtures", FIXTURE), ("tests/expected", EXPECTED)):
                (root / folder).mkdir(parents=True)
                (root / folder / "case-one.json").write_text(json.dumps(data), encoding="utf-8")
            self.assertFinding(check_test_cases.check(root)[0],
                               "docs/methodology/DATA_CONTRACT.md: is missing")


class FixtureTest(FindingTest):
    def test_envelope(self) -> None:
        for key, value, text in (("schema_version", 2, "schema_version is not 1"),
                                 ("kind", "fixture", "kind is not 'test-fixture'"),
                                 ("id", "case-two", "id is not the kebab-case file name"),
                                 ("synthetic", False, "synthetic is not true"),
                                 ("description", " ", "description is not a text"),
                                 ("category", "smoke", "unknown category 'smoke'"),
                                 ("reporting_date", "2026-02-30", "reporting_date is not"),
                                 ("reporting_currency_code", "eur", "is not a currency code")):
            with self.subTest(key=key):
                self.assertFinding(run(lambda f, e: f.update({key: value})), text)

    def test_missing_and_unknown_keys(self) -> None:
        self.assertFinding(run(lambda f, e: f.pop("datasets")), "missing datasets")
        self.assertFinding(run(lambda f, e: f.update(extra=1)), "unknown extra")

    def test_calculation_and_datasets_are_declared(self) -> None:
        self.assertFinding(run(lambda f, e: f.update(calculation="CALC-09")),
                           "calculation CALC-09 is not declared")
        self.assertFinding(run(lambda f, e: f.update(calculation="calc-2")),
                           "is not an ID of the form")
        self.assertFinding(run(lambda f, e: f.update(datasets={"DS-09": FIXTURE["datasets"]
                                                               ["DS-02"]})),
                           "dataset DS-09 is not declared")
        self.assertFinding(run(lambda f, e: f.update(calculation="CALC-01")),
                           "CALC-01 is illustrative")
        self.assertFinding(run(lambda f, e: f.update(datasets={})), "not a non-empty object")

    def test_records(self) -> None:
        def records(value: Any) -> Edit:
            return lambda f, e: f["datasets"].update({"DS-02": value})
        self.assertFinding(run(records([])), "is not a non-empty list")
        self.assertFinding(run(records(["P1"])), "is not an object")
        self.assertFinding(run(records([{"id": "P1"}])), "missing inputs")
        self.assertFinding(run(records([{"id": " ", "inputs": {}}])), "id is not a text")
        self.assertFinding(run(records([{"id": "P1", "inputs": {}}, {"id": "P1", "inputs": {}}])),
                           "is not unique")
        self.assertFinding(run(records([{"id": "P1", "inputs": []}])), "inputs is not an object")

    def test_input_suffixes_and_values(self) -> None:
        for key, value, text in (("notional", 1, "has no unit suffix"),
                                 ("NotionalAmount", 1, "not lower snake case"),
                                 ("notional_amount", "1", "invalid value for its suffix _amount"),
                                 ("trade_count", -1, "suffix _count"),
                                 ("trade_count", True, "suffix _count"),
                                 ("start_date", "2026-13-01", "suffix _date"),
                                 ("netted_flag", 1, "suffix _flag"),
                                 ("currency_code", "EURO", "suffix _code"),
                                 ("position_id", "", "suffix _id")):
            with self.subTest(key=key, value=value):
                def edit(f: dict[str, Any], e: dict[str, Any], k: str = key,
                         v: object = value) -> None:
                    f["datasets"]["DS-02"][0]["inputs"][k] = v
                self.assertFinding(run(edit), text)
        self.assertFinding(run(lambda f, e: f["datasets"]["DS-02"][1]["inputs"]
                               .pop("currency_code")), "amount but no currency_code")

    def test_codes_and_ids_are_whole_values_with_ascii_digits(self) -> None:
        def record(key: str, value: object) -> Edit:
            return lambda f, e: f["datasets"]["DS-02"][0]["inputs"].update({key: value})
        cases: list[tuple[Edit, str]] = [
            (lambda f, e: f.update(reporting_currency_code="EUR\n"),
             "reporting_currency_code is not a currency code"),
            (record("currency_code", "EUR\n"), "invalid value for its suffix _code"),
            (record("notional_amount\n", 1), "is not lower snake case"),
            (lambda f, e: f.update(calculation="CALC-02\n"), "is not an ID of the form"),
            (lambda f, e: f.update(calculation="CALC-0\u0662"), "is not an ID of the form"),
            (lambda f, e: f.update(datasets={"DS-0\u0662": FIXTURE["datasets"]["DS-02"]}),
             "is not an ID of the form"),
            (lambda f, e: output(e).update(unit="EUR\n"), "is not a currency code, 1 or years"),
            (lambda f, e: output(e).update(quantity="total_notional_amount\n"),
             "is not lower snake case"),
            (lambda f, e: output(e).update(rule="CALC-02.1\n"), "is not a calculation step"),
            (lambda f, e: output(e).update(rule="CALC-02.\u0661"), "is not a calculation step"),
        ]
        for edit, text in cases:
            with self.subTest(text=text):
                self.assertFinding(run(edit), text)

    def test_a_step_number_has_ascii_digits(self) -> None:
        contracts = CONTRACTS.replace("| 2 | Average |", "| \u0662 | Average |")
        self.assertFinding(run(lambda f, e: output(e).update(rule="CALC-02.2"),
                               extra={"docs/methodology/CALCULATION_CONTRACTS.md": contracts}),
                           "rule CALC-02.2 is not a step that the Rules table of CALC-02 lists")

    def test_null_input_is_not_applicable(self) -> None:
        def edit(f: dict[str, Any], e: dict[str, Any]) -> None:
            f["datasets"]["DS-02"][0]["inputs"]["maturity_years"] = None
        self.assertEqual(run(edit), [])

    def test_invalid_case_checks_only_the_harness_metadata(self) -> None:
        def edit(f: dict[str, Any], e: dict[str, Any]) -> None:
            f.update(category="invalid", calculation="unknown", datasets={"bad": "x"})
            f.pop("reporting_currency_code")
            e["outputs"] = [{"quantity": "run_status", "text": "REJECTED_UNKNOWN_CALCULATION",
                             "scope": {"dataset": "bad", "record": "R9"}, "rule": "CALC-02.1",
                             "reference": OUTPUT["reference"]}]
        self.assertEqual(run(edit), [])

        def no_status(f: dict[str, Any], e: dict[str, Any]) -> None:
            f["category"] = "invalid"
        self.assertFinding(run(no_status), "needs a _status text output")
        self.assertFinding(run(lambda f, e: f.update(category="invalid", synthetic=False)),
                           "synthetic is not true")

    def test_invalid_case_holds_no_other_key(self) -> None:
        # Only the calculation inputs may be invalid or missing.
        def edit(f: dict[str, Any], e: dict[str, Any]) -> None:
            f.update(category="invalid", extra=1)
            f.pop("datasets")
        findings = run(edit)
        self.assertIn("tests/fixtures/case-one.json: unknown extra", findings)
        self.assertFalse([item for item in findings if "missing" in item], findings)


class ExpectedTest(FindingTest):
    def test_envelope(self) -> None:
        self.assertFinding(run(lambda f, e: e.update(kind="expected")), "envelope does not match")
        self.assertFinding(run(lambda f, e: e.update(fixture="other")), "envelope does not match")
        self.assertFinding(run(lambda f, e: e.pop("outputs")), "missing outputs")
        self.assertFinding(run(lambda f, e: e.update(outputs=[])), "not a non-empty list")
        self.assertFinding(run(lambda f, e: e.update(outputs=["x"])), "output is not an object")

    def test_number_form(self) -> None:
        for key, value, text in (("value", "1", "value is not a finite number"),
                                 ("unit", "eur", "is not a currency code, 1 or years"),
                                 ("tolerance", {"absolute": 0}, "exactly absolute and relative"),
                                 ("tolerance", {"absolute": -1, "relative": 0}, "at least 0")):
            with self.subTest(key=key):
                self.assertFinding(run(lambda f, e: output(e).update({key: value})),
                                   text)
        self.assertFinding(run(lambda f, e: output(e).pop("tolerance")), "missing tolerance")
        self.assertFinding(run(lambda f, e: output(e).update(text=1)), "unknown")

    def test_quantity_scope_and_rule(self) -> None:
        cases: list[tuple[Edit, str]] = [
            (lambda f, e: output(e).update(quantity="Total"), "not lower snake case"),
            (lambda f, e: output(e).update(quantity="illustrative_total_amount"),
             "illustrative_ quantity outside"),
            (lambda f, e: output(e).update(scope={"dataset": "DS-02", "record": "P9"}),
             "DS-02/P9 is not a record"),
            (lambda f, e: output(e).update(scope={"dataset": "DS-02"}), "exactly dataset and"),
            (lambda f, e: output(e).update(rule=None), "only an illustrative value may have"),
            (lambda f, e: output(e).update(rule="CALC-02"), "is not a calculation step"),
            (lambda f, e: output(e).update(rule="CALC-01.1"), "not a step of the case's"),
            (lambda f, e: output(e).update(rule="CALC-02.3"),
             "rule CALC-02.3 is not a step that the Rules table of CALC-02 lists"),
            (lambda f, e: output(e).update(rule="CALC-02.4"), "is not a step that the Rules"),
            (lambda f, e: output(e).update(rule="CALC-02.5"), "is not a step that the Rules"),
        ]
        for edit, text in cases:
            with self.subTest(text=text):
                self.assertFinding(run(edit), text)

    def test_a_rule_is_a_step_its_contract_lists(self) -> None:
        self.assertEqual(run(lambda f, e: output(e).update(rule="CALC-02.2")), [])
        # A contract without a Rules table lists no step.
        def no_rules(f: dict[str, Any], e: dict[str, Any]) -> None:
            f["category"] = "invalid"
            e["outputs"] = [{"quantity": "run_status", "text": "REJECTED", "scope": None,
                             "rule": "CALC-03.1", "reference": OUTPUT["reference"]}]
        self.assertFinding(run(no_rules),
                           "rule CALC-03.1 is not a step that the Rules table of CALC-03 lists")

    def test_reference(self) -> None:
        def ref(**values: object) -> Edit:
            return lambda f, e: output(e)["reference"].update(values)
        cases: list[tuple[Edit, str]] = [
            (lambda f, e: output(e).update(reference=[]), "reference is not an object"),
            (lambda f, e: output(e)["reference"].pop("date"), "reference: missing date"),
            (ref(**{"class": "guess"}), "unknown reference class"),
            (ref(source="SRC-09"), "not in the register"),
            (ref(source="SRC-01"), "illustrative register entry"),
            (ref(locator=None), "published value needs"),
            (ref(source="IMPL-01"), "published value needs"),
            (ref(**{"class": "independent"}), "independent value needs"),
            (ref(**{"class": "illustrative"}), "illustrative value needs a derivation"),
            (ref(script="tests/derivations/missing.py"), "is not an existing file"),
            (ref(date="02/11/2026"), "reference date is not"),
            (ref(derived_by=" "), "derived_by is not a text or null"),
        ]
        for edit, text in cases:
            with self.subTest(text=text):
                self.assertFinding(run(edit), text)

    def test_a_script_outside_tests_derivations_is_refused(self) -> None:
        # The script exists, and so does tests/derivations/ for the ".." path to cross,
        # so only the location rule can refuse it.
        files = {"tools/x.py": "print(1000000)\n", "tests/derivations/total.py": "print(1)\n"}
        for path in ("tools/x.py", "tests/derivations/../../tools/x.py"):
            with self.subTest(path=path):
                self.assertFinding(run(lambda f, e: output(e)["reference"].update(script=path),
                                       extra=files),
                                   "is not an existing file under tests/derivations/")

    def test_existing_derivation_script(self) -> None:
        self.assertEqual(run(lambda f, e: output(e)["reference"].update(
            script="tests/derivations/total.py"),
            extra={"tests/derivations/total.py": "print(1000000)\n"}), [])

    def test_illustrative_prefix_matches_the_classes(self) -> None:
        def all_illustrative(f: dict[str, Any], e: dict[str, Any]) -> None:
            output(e)["reference"].update({"class": "illustrative", "derivation": "Example"})
        self.assertFinding(run(all_illustrative), "illustrative- prefix must be used exactly")


PLAN_HEADER = ("# Plan\n\n## Traceability\n\n"
               "| Source | Paragraph | Calculation step | VBA procedure | Test cases | Status |\n"
               "| --- | --- | --- | --- | --- | --- |\n")
TRACED = "| `SRC-02` | Article 1(1) | `CALC-02.1` | `CORE_Totals.TotalNotional` | `case-one` | Validated |"
CORE = ('Attribute VB_Name = "CORE_Totals"\nOption Explicit\nOption Private Module\n\n'
        "Public Function TotalNotional(ByVal Amount As Double) As Double\n"
        "    TotalNotional = Amount\nEnd Function\n")


def plan(*rows: str) -> dict[str, str]:
    """A validation plan whose Traceability table holds rows, and the core module it names."""
    table = PLAN_HEADER + "".join(f"{row}\n" for row in rows)
    return {"docs/methodology/VALIDATION_PLAN.md": table + "\nText after the table.\n\n"
            "| Status | Meaning |\n| --- | --- |\n| Validated | Not a row |\n",
            "src/core/CORE_Totals.bas": CORE}


def traced(column: str, value: str) -> str:
    """TRACED with one column replaced."""
    columns = ["Source", "Paragraph", "Calculation step", "VBA procedure", "Test cases", "Status"]
    values = [cell.strip() for cell in TRACED.strip("|").split("|")]
    values[columns.index(column)] = value
    return "| " + " | ".join(values) + " |"


class TraceabilityTest(FindingTest):
    def test_a_traced_rule_passes(self) -> None:
        self.assertEqual(run(extra=plan(TRACED)), [])
        # Names ignore case, as in VBA; a rule not implemented yet names no procedure.
        self.assertEqual(run(extra=plan(traced("VBA procedure", "`core_totals.totalnotional`"),
                                        traced("Status", "Exercised"),
                                        "| *Illustrative:* `SRC-01` | Article 2 | `CALC-01.1` "
                                        "| \u2014 | \u2014 | Not implemented |")), [])

    def test_a_named_procedure_or_case_must_exist(self) -> None:
        for column, value, text in (
                ("VBA procedure", "`CORE_Totals.Missing`",
                 "VBA procedure 'CORE_Totals.Missing' is not in a component under src/"),
                ("VBA procedure", "`CORE_Other.TotalNotional`", "is not in a component"),
                ("VBA procedure", "`TotalNotional`", "is not Module.Procedure"),
                ("VBA procedure", "TotalNotional", "must name Module.Procedure"),
                ("Test cases", "`case-one`, `case-two`", "test case 'case-two' does not exist"),
                ("Test cases", "case-one", "must be case IDs")):
            with self.subTest(column=column, value=value):
                findings = run(extra=plan(traced(column, value)))
                self.assertFinding(findings, text)
                self.assertEqual(len(findings), 1, findings)
                self.assertTrue(findings[0].startswith("docs/methodology/VALIDATION_PLAN.md:7: "))

    def test_a_procedure_of_a_form_counts(self) -> None:
        files = plan(traced("VBA procedure", "`F_Main.ApplyRule`"))
        files["src/forms/F_Main.frm"] = (
            "VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Main\n"
            '   Caption         =   "Main"\nEnd\nAttribute VB_Name = "F_Main"\n'
            "Option Explicit\n\nPrivate Sub ApplyRule()\nEnd Sub\n")
        self.assertEqual(run(extra=files), [])

    def test_a_procedure_outside_src_does_not_count(self) -> None:
        files = plan(TRACED)
        files["tests/modules/CORE_Totals.bas"] = files.pop("src/core/CORE_Totals.bas")
        self.assertFinding(run(extra=files), "is not in a component under src/")

    def test_the_row_matches_the_guides(self) -> None:
        for column, value, text in (
                ("Source", "`SRC-09`", "is not one SRC- ID of the register"),
                ("Source", "`IMPL-01`", "is not one SRC- ID of the register"),
                ("Paragraph", "", "names no paragraph"),
                ("Calculation step", "`CALC-09.1`", "is not a step of a declared calculation"),
                ("Calculation step", "`CALC-02`", "is not a step of a declared calculation"),
                ("Calculation step", "`CALC-02.9`",
                 "rule CALC-02.9 is not a step that the Rules table of CALC-02 lists"),
                ("Calculation step", "`CALC-01.1`", "test case 'case-one' tests CALC-02, not "
                                                    "CALC-01")):
            with self.subTest(column=column, value=value):
                self.assertFinding(run(extra=plan(traced(column, value))), text)

    def test_the_status_agrees_with_the_row(self) -> None:
        def illustrative(f: dict[str, Any], e: dict[str, Any]) -> None:
            output(e)["reference"].update({"class": "illustrative", "derivation": "Example"})
            f.update(id="illustrative-one")
            e.update(fixture="illustrative-one")
        no_case = traced("Test cases", "\u2014")
        for row, text in (
                (traced("Status", "Done"), "status 'Done' is not one of"),
                (traced("Status", "Not implemented"), "'Not implemented' needs no VBA procedure"),
                (traced("VBA procedure", "\u2014"), "'Validated' needs a VBA procedure"),
                (traced("Status", "Not validated"), "a rule with test cases is not"),
                (no_case.replace("Validated", "Exercised"), "'Exercised' needs a test case")):
            with self.subTest(row=row):
                self.assertFinding(run(extra=plan(row)), text)
        self.assertEqual(run(extra=plan(no_case.replace("Validated", "Not validated"))), [])
        self.assertFinding(run(illustrative, name="illustrative-one", extra=plan(
            traced("Test cases", "`illustrative-one`"))),
            "'Validated' needs a test case with a published or independent output for "
            "CALC-02.1")

    def test_validated_needs_reference_evidence_for_its_own_step(self) -> None:
        def two_steps(f: dict[str, Any], e: dict[str, Any]) -> None:
            second = copy.deepcopy(OUTPUT)
            second.update(quantity="average_notional_amount", rule="CALC-02.2")
            second["reference"].update({"class": "illustrative", "source": None,
                                        "locator": None, "derivation": "Example"})
            e["outputs"].append(second)
        # case-one holds a published output for CALC-02.1 only.
        self.assertFinding(run(two_steps, extra=plan(traced("Calculation step", "`CALC-02.2`"))),
                           "'Validated' needs a test case with a published or independent "
                           "output for CALC-02.2")
        self.assertEqual(run(two_steps, extra=plan(TRACED)), [])

    def test_an_invalid_case_is_compared_when_its_calculation_is_checked(self) -> None:
        row = ("| *Illustrative:* `SRC-01` | Article 1(1) | `CALC-01.1` "
               "| `CORE_Totals.TotalNotional` | `{}` | Exercised |")

        def invalid(manifest: dict[str, Any], expected: dict[str, Any]) -> None:
            manifest["category"] = "invalid"
        # A file-based case's manifest is checked in every category, so its
        # calculation is trusted and must be the row's.
        self.assertFinding(file_case(invalid, extra=plan(row.format("file-one"))),
                           "test case 'file-one' tests CALC-02, not CALC-01")
        # An invalid JSON fixture may break its calculation, so it is not compared.
        findings = run(lambda f, e: f.update(category="invalid"),
                       extra=plan(row.format("case-one")))
        self.assertFalse([item for item in findings if "VALIDATION_PLAN.md" in item], findings)

    def test_the_table_needs_its_columns(self) -> None:
        files = {"docs/methodology/VALIDATION_PLAN.md":
                 "## Traceability\n\n| Source | Status |\n| --- | --- |\n"}
        self.assertFinding(run(extra=files), "VALIDATION_PLAN.md:3: the Traceability table has no "
                                             "Paragraph, Calculation step, VBA procedure, Test "
                                             "cases column")

    def test_rows_are_checked_without_any_case(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for path, text in {**plan(TRACED),
                               "docs/methodology/CALCULATION_CONTRACTS.md": CONTRACTS,
                               "docs/methodology/DATA_CONTRACT.md": DATA,
                               "docs/methodology/SOURCES.md": SOURCES}.items():
                (root / path).parent.mkdir(parents=True, exist_ok=True)
                (root / path).write_text(text, encoding="utf-8")
            self.assertFinding(check_test_cases.check(root)[0],
                               "test case 'case-one' does not exist")


DATA_FILES = ("# Data\n\n## Datasets\n\n| ID | Description | File or sheet | Encoding |\n"
              "| --- | --- | --- | --- |\n"
              "| `DS-02` | Real positions | `positions.csv`, `limits.csv` | |\n"
              "| `DS-03` | Legacy feed | `legacy.txt` | `cp1252` |\n"
              "| `DS-04` | Manual input | `Inputs` sheet | |\n"
              "| *Illustrative:* `DS-01` | Example | `example.csv` | |\n")
MANIFEST: dict[str, Any] = {
    "schema_version": 1, "kind": "test-file-fixture", "id": "file-one", "synthetic": True,
    "description": "A synthetic import", "category": "nominal", "calculation": "CALC-02",
    "reporting_date": "2026-01-02", "reporting_currency_code": "EUR",
    "datasets": ["DS-02", "DS-03"]}
VARIANT_FILES = {"lf/positions.csv": b"id,notional\nP1,600000\n",
                 "crlf/positions.csv": b"id,notional\r\nP1,600000\r\n",
                 "crlf/legacy.txt": "name\r\ncaf\u00e9\r\n".encode("cp1252")}


def contract_dataset(dataset: str, source: str, fields: list[dict[str, Any]],
                     encoding: str | None = "UTF-8", delimiter: str | None = ",") -> dict[str, Any]:
    return {"id": dataset, "illustrative": dataset == "DS-01", "description": "A dataset",
            "provider": "A system", "frequency": "Daily", "granularity": "One row per record",
            "source": source, "encoding": encoding, "delimiter": delimiter,
            "decimal_separator": "." if delimiter else None,
            "fields": [{"unit": None, "required": True, "allowed": None, "notes": "", **item}
                       for item in fields]}


# The machine-readable contract of the datasets in DATA_FILES.
CONTRACT_FILE: dict[str, Any] = {"schema_version": 1, "kind": "data-contract", "datasets": [
    contract_dataset("DS-02", "positions.csv", [
        {"name": "id", "type": "text"},
        {"name": "notional", "type": "decimal", "allowed": {"min": 0}}]),
    contract_dataset("DS-03", "legacy.txt", [{"name": "name", "type": "text"}], "cp1252"),
    contract_dataset("DS-04", "Inputs", [{"name": "id", "type": "text"}], None, None),
    contract_dataset("DS-01", "example.csv", [{"name": "id", "type": "text"}]),
]}
EXPECTED_VARIANTS: dict[str, Any] = {
    "schema_version": 1, "kind": "test-expected", "fixture": "file-one",
    "variants": {"lf": [OUTPUT], "crlf": [OUTPUT]}}


def file_case(edit: Edit | None = None, files: dict[str, bytes | None] | None = None,
              extra: dict[str, str] | None = None) -> list[str]:
    """Check a repository with one file-based case, after edit(manifest, expected).

    files adds or replaces variant files, relative to the case directory; None
    removes one, and a name ending in / makes an empty directory.
    """
    manifest, expected = copy.deepcopy(MANIFEST), copy.deepcopy(EXPECTED_VARIANTS)
    if edit is not None:
        edit(manifest, expected)
    variant_files: dict[str, bytes | None] = {**VARIANT_FILES, **(files or {})}
    with tempfile.TemporaryDirectory(prefix="file-cases-") as directory:
        root = Path(directory)
        texts = {"docs/methodology/CALCULATION_CONTRACTS.md": CONTRACTS,
                 "docs/methodology/DATA_CONTRACT.md": DATA_FILES,
                 "docs/methodology/data-contract.json": json.dumps(CONTRACT_FILE),
                 "docs/methodology/SOURCES.md": SOURCES,
                 "tests/fixtures/file-one/case.json": json.dumps(manifest),
                 "tests/expected/file-one.json": json.dumps(expected), **(extra or {})}
        for path, text in texts.items():
            (root / path).parent.mkdir(parents=True, exist_ok=True)
            (root / path).write_text(text, encoding="utf-8")
        case = root / "tests/fixtures/file-one"
        for name, content in variant_files.items():
            if content is None:
                continue
            (case / name).parent.mkdir(parents=True, exist_ok=True)
            if name.endswith("/"):
                (case / name).mkdir(exist_ok=True)
            else:
                (case / name).write_bytes(content)
        return check_test_cases.check(root)[0]


class FileBasedCaseTest(FindingTest):
    def test_valid_file_case_passes(self) -> None:
        self.assertEqual(file_case(), [])

    def test_record_scope_names_a_dataset_of_the_case(self) -> None:
        def scoped(dataset: str) -> Edit:
            def edit(m: dict[str, Any], e: dict[str, Any]) -> None:
                e["variants"]["lf"] = [{**OUTPUT, "scope": {"dataset": dataset,
                                                            "record": "line 2"}}]
            return edit
        self.assertEqual(file_case(scoped("DS-02")), [])
        self.assertFinding(file_case(scoped("DS-09")), "is not a dataset of the case")

    def test_an_invalid_case_scopes_a_dataset_of_its_manifest(self) -> None:
        def scoped(dataset: str) -> Edit:
            def edit(m: dict[str, Any], e: dict[str, Any]) -> None:
                m["category"] = "invalid"
                status = {"quantity": "import_status", "text": "REJECTED", "rule": "CALC-02.1",
                          "scope": {"dataset": dataset, "record": "line 2"},
                          "reference": OUTPUT["reference"]}
                e["variants"] = {"lf": [status], "crlf": [status]}
            return edit
        self.assertEqual(file_case(scoped("DS-02")), [])
        self.assertFinding(file_case(scoped("DS-09")), "is not a dataset of the case")

    def test_variants_and_expected_outputs_match(self) -> None:
        self.assertFinding(file_case(lambda m, e: e["variants"].pop("crlf")),
                           "variant crlf has no expected outputs")
        self.assertFinding(file_case(lambda m, e: e["variants"].update(gone=[OUTPUT])),
                           "variant 'gone' has no directory under tests/fixtures/file-one/")
        self.assertFinding(file_case(lambda m, e: e["variants"].update(lf=[])),
                           "variants.lf: outputs is not a non-empty list")
        self.assertFinding(file_case(lambda m, e: e.update(variants={})),
                           "variants is not a non-empty object")
        self.assertFinding(file_case(lambda m, e: e.update(outputs=e.pop("variants"))),
                           "lists outputs by variant")

    def test_variant_holds_only_contract_files(self) -> None:
        self.assertFinding(file_case(files={"lf/notes.txt": b"x\n"}),
                           "lf/notes.txt: is not a file the data contract lists")
        self.assertFinding(file_case(files={"lf/example.csv": b"x\n"}),
                           "lf/example.csv: is not a file the data contract lists")
        self.assertFinding(file_case(files={"lf/deeper/positions.csv": b"x\n"}),
                           "lf/deeper: is not a file")
        self.assertFinding(file_case(files={"empty/": b""}), "empty: is an empty variant")
        self.assertFinding(file_case(files={"Upper/positions.csv": b"x\n"}),
                           "Upper: is not a kebab-case variant name")
        self.assertFinding(file_case(files={"README.md": b"x\n"}),
                           "file-one/README.md: is not case.json or a variant directory")

    def test_files_are_read_in_the_contract_encoding(self) -> None:
        self.assertFinding(file_case(files={"lf/positions.csv": b"id\n\xff\n"}),
                           "lf/positions.csv: is not UTF-8 (invalid start byte at byte 3)")
        self.assertFinding(file_case(files={"crlf/legacy.txt": b"\x81"}),
                           "crlf/legacy.txt: is not cp1252")

        def invalid(m: dict[str, Any], e: dict[str, Any]) -> None:
            m["category"] = "invalid"
            status = {"quantity": "import_status", "text": "REJECTED_ENCODING", "scope": None,
                      "rule": "CALC-02.1", "reference": OUTPUT["reference"]}
            e["variants"] = {"lf": [status], "crlf": [status]}
        self.assertEqual(file_case(invalid, files={"lf/positions.csv": b"\xff"}), [])
        self.assertFinding(file_case(lambda m, e: m.update(category="invalid")),
                           "variants.lf: an invalid case needs a _status text output")

    def test_variant_files_are_validated_against_the_contract(self) -> None:
        self.assertFinding(file_case(files={"lf/positions.csv": b"id,notional\nP1,-5\n"}),
                           "tests/fixtures/file-one/lf/positions.csv: line 2: DC-RANGE: field "
                           "notional value '-5' is not \u2265 0")
        findings = file_case(files={"crlf/positions.csv": b"id,amount\r\nP1,1\r\n",
                                    "crlf/legacy.txt": "nom\r\n".encode("cp1252")})
        self.assertEqual(findings, [
            "tests/fixtures/file-one/crlf/legacy.txt: line 1: DC-UNKNOWN-COLUMN: column 'nom' "
            "is not a field of DS-03",
            "tests/fixtures/file-one/crlf/legacy.txt: line 1: DC-MISSING-COLUMN: required field "
            "name has no column",
            "tests/fixtures/file-one/crlf/positions.csv: line 1: DC-UNKNOWN-COLUMN: column "
            "'amount' is not a field of DS-02",
            "tests/fixtures/file-one/crlf/positions.csv: line 1: DC-MISSING-COLUMN: required "
            "field notional has no column"])

    def test_an_invalid_case_may_break_the_contract(self) -> None:
        def invalid(m: dict[str, Any], e: dict[str, Any]) -> None:
            m["category"] = "invalid"
            status = {"quantity": "import_status", "text": "REJECTED_NEGATIVE", "scope": None,
                      "rule": "CALC-02.1", "reference": OUTPUT["reference"]}
            e["variants"] = {"lf": [status], "crlf": [status]}
        self.assertEqual(file_case(invalid, files={"lf/positions.csv": b"id,notional\nP1,-5\n"}),
                         [])

    def test_the_contract_must_describe_each_file(self) -> None:
        self.assertFinding(file_case(files={"lf/limits.csv": b"id\nL1\n"}),
                           "lf/limits.csv: is not the file of a dataset of "
                           "docs/methodology/data-contract.json")

    def test_the_contract_must_exist_and_be_valid(self) -> None:
        missing = file_case(extra={"docs/methodology/data-contract.json": ""})
        self.assertFinding(missing, "data-contract.json: is not a valid data contract (1 "
                                    "problem(s); python tools/check_data_contract.py lists them)")
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for path, text in {"docs/methodology/DATA_CONTRACT.md": DATA_FILES,
                               "docs/methodology/CALCULATION_CONTRACTS.md": CONTRACTS,
                               "docs/methodology/SOURCES.md": SOURCES,
                               "tests/fixtures/file-one/case.json": json.dumps(MANIFEST),
                               "tests/fixtures/file-one/lf/positions.csv": "id,notional\n",
                               "tests/expected/file-one.json": json.dumps(
                                   {**EXPECTED_VARIANTS, "variants": {"lf": [OUTPUT]}})}.items():
                (root / path).parent.mkdir(parents=True, exist_ok=True)
                (root / path).write_text(text, encoding="utf-8")
            self.assertEqual(check_test_cases.check(root)[0], [
                "docs/methodology/data-contract.json: is missing; the variant files of a "
                "file-based case are validated against it"])

    def test_manifest(self) -> None:
        cases: list[tuple[Edit, str]] = [
            (lambda m, e: m.update(kind="test-fixture"), "kind is not 'test-file-fixture'"),
            (lambda m, e: m.update(id="file-two"), "id is not the kebab-case directory name"),
            (lambda m, e: m.update(datasets=[]), "datasets is not a non-empty list"),
            (lambda m, e: m.update(datasets=["DS-02", "DS-02"]), "of distinct DS- IDs"),
            (lambda m, e: m.update(datasets=["DS-09"]), "dataset DS-09 is not declared"),
            (lambda m, e: m.update(datasets=["DS-02", "DS-04"]),
             "dataset DS-04 names no file in the Datasets table"),
            (lambda m, e: m.update(datasets=["DS-01", "DS-03"]), "DS-01 is illustrative"),
            (lambda m, e: m.update(calculation="CALC-09"), "calculation CALC-09 is not"),
            (lambda m, e: m.pop("reporting_date"), "missing reporting_date"),
        ]
        for edit, text in cases:
            with self.subTest(text=text):
                self.assertFinding(file_case(edit), text)

    def test_case_directory(self) -> None:
        self.assertFinding(file_case(files={"lf/positions.csv": None, "crlf/positions.csv": None,
                                            "crlf/legacy.txt": None}),
                           "file-one: has no variant directory")
        self.assertFinding(file_case(extra={"tests/fixtures/file-one.json": "{}"}),
                           "file-one.json is a second fixture of the same case")
        self.assertFinding(file_case(extra={"tests/fixtures/stray/positions.csv": "x\n"}),
                           "tests/fixtures/stray: is a directory without case.json")

    def test_data_contract_encodings(self) -> None:
        unknown = DATA_FILES.replace("`cp1252`", "`klingon`")
        self.assertFinding(file_case(extra={"docs/methodology/DATA_CONTRACT.md": unknown}),
                           "DS-03 has an unknown encoding 'klingon'")
        without_column = ("# Data\n\n## Datasets\n\n| ID | File or sheet |\n| --- | --- |\n"
                          "| `DS-02` | `positions.csv`, `limits.csv` |\n| `DS-03` | `legacy.txt` |\n")
        # Without an Encoding column every file is UTF-8, which the cp1252 file is not.
        self.assertFinding(file_case(extra={"docs/methodology/DATA_CONTRACT.md": without_column}),
                           "crlf/legacy.txt: is not UTF-8")

    @unittest.skipIf(shutil.which("git") is None, "git is not installed")
    def test_variant_files_are_stored_byte_for_byte(self) -> None:
        paths = ["tests/fixtures/a-case/a-variant/positions.csv", "tests/fixtures/a-case/case.json"]
        printed = subprocess.run(["git", "check-attr", "text", "whitespace", "--", *paths],
                                 cwd=ROOT, capture_output=True, text=True, check=True).stdout
        self.assertEqual(printed.splitlines(), [
            f"{paths[0]}: text: unset", f"{paths[0]}: whitespace: unset",
            f"{paths[1]}: text: set", f"{paths[1]}: whitespace: unspecified"])
        editorconfig = (ROOT / ".editorconfig").read_text(encoding="utf-8")
        section = editorconfig.split("[tests/fixtures/*/*/**]\n", 1)[1].split("\n\n", 1)[0]
        self.assertEqual(section.splitlines(), [
            "charset = unset", "end_of_line = unset", "insert_final_newline = unset",
            "trim_trailing_whitespace = false"])


class MainTest(unittest.TestCase):
    def verdict(self, edit: Edit | None = None, empty: bool = False) -> tuple[int, str]:
        with tempfile.TemporaryDirectory(prefix="test-cases-") as directory, \
                contextlib.redirect_stdout(io.StringIO()) as printed:
            if not empty:
                write(Path(directory), edit)
            code = check_test_cases.main(["--root", directory])
        return code, printed.getvalue()

    def test_pass(self) -> None:
        code, printed = self.verdict()
        self.assertEqual(code, 0)
        self.assertIn("test-cases: PASS (1 fixture(s), 1 expected file(s)", printed)

    def test_fail(self) -> None:
        code, printed = self.verdict(lambda f, e: f.update(synthetic=False))
        self.assertEqual(code, 1)
        self.assertIn("FAIL: tests/fixtures/case-one.json: synthetic is not true", printed)
        self.assertIn("test-cases: FAIL (1 fixture(s), 1 expected file(s)", printed)

    def test_no_cases(self) -> None:
        code, printed = self.verdict(empty=True)
        self.assertEqual(code, 0)
        self.assertIn("test-cases: PASS (no test cases found", printed)


if __name__ == "__main__":
    unittest.main()
