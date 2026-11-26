#!/usr/bin/env python3
"""Generate tests/modules/TEST_ExcelOracle.bas, the Excel oracle of the date layer.

The oracle computes the worksheet-function cases of tests/fixtures/ a second
time, with Excel's own worksheet functions only (ALLOWED below) and never with a
KPR_ procedure: each covered case becomes one formula, evaluated on a scratch
workbook in the 1900 date system, whose encoded value must equal the case's
expected text. A case is covered only where the date layer's contract and
Excel's functions agree; every other case is excluded, with the reason, and
--list prints both lists (docs/methodology/EXCEL_ORACLE.md).

Each Case procedure evaluates up to GROUP cases of one function in one scratch
workbook, and asserts last that the workbook was closed and the calculation mode
restored. An assertion is named after its case.

    python tools/gen_excel_oracle_suite.py --write   # regenerate the module
    python tools/gen_excel_oracle_suite.py --check   # fail if the module is stale
    python tools/gen_excel_oracle_suite.py --list    # print the covered and excluded cases
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from collections import Counter
from collections.abc import Sequence
from pathlib import Path
from types import ModuleType
from typing import Any

import gen_date_case_suite as base

ROOT = Path(__file__).resolve().parents[1]
MODULE = Path("tests/modules/TEST_ExcelOracle.bas")
SUITE_NAME = "excel-oracle"
GROUP = 30
UPDATED = "2026-11-26"
SCRATCH_ASSERTION = "scratch-closed"
# The only worksheet functions a formula of the oracle may call.
ALLOWED = ("DATE", "YEAR", "MONTH", "DAY", "EOMONTH", "EDATE", "WEEKDAY", "INT", "MOD", "IF")
DEFAULT = "DEFAULT"

# The reasons a case is excluded, in the order the documentation lists them.
REASON_CONTEXT = "a 1904 workbook: the refusal is the date layer's own rule"
REASON_FUNCTION = "no Excel worksheet function computes it"
REASON_CONDITION = "an expected error: Excel has no such condition"
REASON_TEXT = "a text input: Excel converts text by the regional settings"
REASON_ARRAY = "an array input: the oracle evaluates one value per case"
REASON_FRACTION = "a fractional number or a date with a time"
REASON_OTHER = "an input of another type, which Excel converts differently"
REASON_1900 = "year 1900, which Excel counts as a leap year"
REASONS = (REASON_CONTEXT, REASON_FUNCTION, REASON_CONDITION, REASON_TEXT, REASON_ARRAY,
           REASON_FRACTION, REASON_OTHER, REASON_1900)

header = base.header
section = base.section
literal = base.literal
comment = base.comment


def argument(text: str, kind: str) -> str | tuple[None, str]:
    """The Excel expression of one argument, or (None, the reason it is excluded)."""
    if text.startswith("date:"):
        year, month, day = (int(part) for part in text[5:].split("-"))
        return f"DATE({year},{month},{day})"
    if text.startswith(("num:", "lng:")):
        value = float(text[4:])
        if not value.is_integer():
            return None, REASON_FRACTION
        return str(int(value))
    if kind.startswith("bool") and text.startswith("bool:"):
        return "TRUE" if text == "bool:TRUE" else "FALSE"
    if kind.startswith("bool") and text == "empty":
        return DEFAULT
    if text.startswith("str:"):
        return None, REASON_TEXT
    if text.startswith(("arr1:", "arr2:", "fill2:")):
        return None, REASON_ARRAY
    if text.startswith("datetime:"):
        return None, REASON_FRACTION
    return None, REASON_OTHER


def control(values: list[str], index: int, default: str) -> str:
    """A control's Excel value: the case's, or the function's default when it is absent."""
    value = values[index] if len(values) > index else DEFAULT
    return default if value == DEFAULT else value


# The identities of one date or one year argument, {0} standing for it.
SIMPLE = {
    "DaysInMonth": "DAY(EOMONTH({0},0))",
    "DaysInYear": "DATE({0},12,31)-DATE({0},1,1)+1",
    "IsLeapYear": "DAY(EOMONTH(DATE({0},2,1),0))=29",
    "BeginOfMonth": "DATE(YEAR({0}),MONTH({0}),1)",
    "EndOfMonth": "EOMONTH({0},0)",
    "BeginOfQuarter": "DATE(YEAR({0}),INT((MONTH({0})-1)/3)*3+1,1)",
    "EndOfQuarter": "EOMONTH(DATE(YEAR({0}),INT((MONTH({0})-1)/3)*3+3,1),0)",
    "BeginOfYear": "DATE(YEAR({0}),1,1)",
    "EndOfYear": "DATE(YEAR({0}),12,31)",
    "IsMonthEnd": "{0}=EOMONTH({0},0)",
    "IsQuarterEnd": "{0}=EOMONTH(DATE(YEAR({0}),INT((MONTH({0})-1)/3)*3+3,1),0)",
    "IsYearEnd": "{0}=DATE(YEAR({0}),12,31)",
    "AddDays": "{0}+{1}",
    "AddWeeks": "{0}+7*{1}",
}


def week_type(values: list[str], index: int) -> int:
    """WEEKDAY's return type for the weekday base: 2 for Monday as 1, 1 for Sunday."""
    return 2 if control(values, index, "TRUE") == "TRUE" else 1


def formula(short: str, values: list[str]) -> str:
    """The formula of a covered case: its function's identity with Excel's functions."""
    if short in SIMPLE:
        return SIMPLE[short].format(*values)
    if short == "DayOfWeek":
        return f"WEEKDAY({values[0]},{week_type(values, 1)})"
    if short in ("AddMonths", "AddYears"):
        date = values[0]
        months = values[1] if short == "AddMonths" else f"12*{values[1]}"
        if control(values, 2, "FALSE") == "TRUE":
            return f"IF({date}=EOMONTH({date},0),EOMONTH({date},{months}),EDATE({date},{months}))"
        return f"EDATE({date},{months})"
    if short == "NthWeekdayOfMonth":
        start = f"DATE({values[0]},{values[1]},1)"
        return (f"{start}+MOD({values[2]}-WEEKDAY({start},{week_type(values, 4)}),7)"
                f"+7*({values[3]}-1)")
    if short == "LastWeekdayOfMonth":
        end = f"EOMONTH(DATE({values[0]},{values[1]},1),0)"
        return f"{end}-MOD(WEEKDAY({end},{week_type(values, 3)})-{values[2]},7)"
    raise ValueError(f"no Excel identity for {short}")


# The calculation-contract step that each covered function's result tests.
RULES = {
    "DayOfWeek": "CALC-02.1", "DaysInMonth": "CALC-03.2", "DaysInYear": "CALC-03.3",
    "IsLeapYear": "CALC-03.4", "BeginOfMonth": "CALC-04.2", "BeginOfQuarter": "CALC-04.2",
    "BeginOfYear": "CALC-04.2", "EndOfMonth": "CALC-04.3", "EndOfQuarter": "CALC-04.3",
    "EndOfYear": "CALC-04.3", "IsMonthEnd": "CALC-05.1", "IsQuarterEnd": "CALC-05.1",
    "IsYearEnd": "CALC-05.1", "AddDays": "CALC-06.1", "AddWeeks": "CALC-06.2",
    "AddMonths": "CALC-06.3", "AddYears": "CALC-06.5", "NthWeekdayOfMonth": "CALC-07.2",
    "LastWeekdayOfMonth": "CALC-07.4",
}


def rule(function: str, text: str) -> str:
    """The step a covered case's result tests, from its function and its formula."""
    short = function.removeprefix("KPR_Dates_")
    if short == "DayOfWeek" and text.endswith(",1)"):
        return "CALC-02.2"
    if short == "AddMonths" and text.startswith("IF("):
        return "CALC-06.4"
    return RULES[short]


COVERED_FUNCTIONS = frozenset({
    "DayOfWeek", "DaysInMonth", "DaysInYear", "IsLeapYear", "BeginOfMonth", "EndOfMonth",
    "BeginOfQuarter", "EndOfQuarter", "BeginOfYear", "EndOfYear", "IsMonthEnd",
    "IsQuarterEnd", "IsYearEnd", "AddDays", "AddWeeks", "AddMonths", "AddYears",
    "NthWeekdayOfMonth", "LastWeekdayOfMonth"})


def classify(fixture: dict[str, Any], expected: dict[str, Any],
             model: ModuleType) -> tuple[str | None, str | None]:
    """(formula, None) for a covered case, or (None, the reason it is excluded)."""
    short = fixture["function"].removeprefix("KPR_Dates_")
    if fixture["context"] == "ws1904":
        return None, REASON_CONTEXT
    if short not in COVERED_FUNCTIONS:
        return None, REASON_FUNCTION
    if any(output["quantity"].endswith("_status") for output in expected["outputs"]):
        return None, REASON_CONDITION
    spec = model.FUNCTIONS[fixture["function"]]
    kinds = [kind for _, kind in (*spec.params, *spec.controls)]
    values = []
    for text, kind in zip(fixture["arguments"], kinds):
        value = argument(text, kind)
        if isinstance(value, tuple):
            return None, value[1]
        values.append(value)
    if short in ("IsLeapYear", "DaysInYear") and values[0] == "1900":
        return None, REASON_1900
    return formula(short, values), None


def cases(root: Path, model: ModuleType) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """(covered, excluded): every worksheet-function case, by function in inventory order."""
    order = re.findall(r"^KPR_Dates\tFunction\t(\w+)$",
                       (root / base.API).read_text(encoding="utf-8"), re.M)
    covered, excluded = [], []
    for path in sorted((root / base.FIXTURES).glob("*.json")):
        fixture = json.loads(path.read_text(encoding="utf-8"))
        if fixture.get("kind") != "test-call-fixture":
            continue
        expected = json.loads((root / base.EXPECTED / path.name).read_text(encoding="utf-8"))
        text, reason = classify(fixture, expected, model)
        item = {"id": fixture["id"], "function": fixture["function"],
                "description": fixture["description"]}
        if text is None:
            excluded.append({**item, "reason": reason})
            continue
        result = expected["outputs"][0]["text"]
        covered.append({**item, "formula": text, "expected": result,
                        "kind": result.split(":", 1)[0]})
    key = lambda case: (order.index(case["function"]), case["id"])  # noqa: E731
    return sorted(covered, key=key), sorted(excluded, key=key)


def groups(found: list[dict[str, Any]]) -> list[tuple[str, str, list[dict[str, Any]]]]:
    """(procedure, case id, cases): up to GROUP cases of one function each."""
    made = []
    for function in dict.fromkeys(case["function"] for case in found):
        mine = [case for case in found if case["function"] == function]
        short = function.removeprefix("KPR_Dates_")
        kebab = re.sub(r"(?<!^)(?=[A-Z])", "-", short).lower()
        for index in range(math.ceil(len(mine) / GROUP)):
            made.append((f"Case{short}{index + 1:02d}", f"{kebab}-{index + 1:02d}",
                         mine[index * GROUP:(index + 1) * GROUP]))
    return made


def case_procedure(name: str, case_id: str, part: list[dict[str, Any]]) -> list[str]:
    short = part[0]["function"].removeprefix("KPR_Dates_")
    lines = [f"Private Sub {name}()", ""]
    lines += header(f"ORACLE OF {short.upper()}", [
        ("PURPOSE", [f"Computes {len(part)} cases of {short} with Excel's worksheet functions, "
                     f"from {part[0]['id']}",
                     f"to {part[-1]['id']}."]),
        ("WHY THIS EXISTS", ["Each assertion compares Excel's result with the case's expected "
                             "file, which the date",
                             "layer's code never computed."]),
        ("INPUTS", ["None"]),
        ("RETURNS", ["Nothing"]),
        ("BEHAVIOR", ["Opens a scratch workbook in the 1900 date system, evaluates each case's "
                      "formula on its",
                      "sheet, and asserts the encoded value under the case's id. Closes the "
                      "workbook on every",
                      "path and asserts that it did."]),
        ("ERROR POLICY", ["An unexpected error is kept, the workbook closed, and the error "
                          "recorded against",
                          "the case by TEST_Harness.RecordCaseError; it does not stop the run"]),
        ("DEPENDENCIES", ["TEST_WorksheetSupport, TEST_Harness; Excel's worksheet functions"]),
        ("NOTES", ["Generated by tools/gen_excel_oracle_suite.py from tests/fixtures/"]),
        ("UPDATED", [UPDATED]),
    ])
    lines += ["", *section("DECLARE"),
              "    Dim Sheet           As Worksheet    'The scratch sheet",
              "    Dim Actual          As String       'The encoded value of a formula",
              "    Dim Detail          As String       'What closing the workbook could not "
              "restore",
              "    Dim ErrNumber       As Long         'Error of the case; 0 for none",
              "    Dim ErrSource       As String       'Its source",
              "    Dim ErrDescription  As String       'Its description", "",
              *section("INITIALIZE"),
              "    'Record an unexpected error against the case",
              "        On Error GoTo CaseFailed",
              "    'Start the case",
              f"        TEST_Harness.BeginCase {literal(case_id)}",
              "    'Close the scratch workbook on every path",
              "        On Error GoTo Failed", "",
              *section("RUN CASES"),
              "    'Open a scratch workbook in the 1900 date system",
              "        Set Sheet = TEST_WorksheetSupport.OpenScratch(False)"]
    for case in part:
        lines += [comment(case["description"]),
                  f"        Actual = TEST_WorksheetSupport.EncodeCell(Sheet.Evaluate("
                  f"{literal(case['formula'])}), {literal(case['kind'])})",
                  f"        TEST_Harness.AssertEqualText {literal(case['id'])}, "
                  f"{literal(case['expected'])}, Actual"]
    lines += ["", *section("EXIT PROCEDURE"), "CleanUp:",
              "    'Restore normal error handling", "        On Error GoTo 0",
              "    'Close the scratch workbook and restore the calculation mode",
              "        Detail = TEST_WorksheetSupport.CloseScratch()",
              f"        TEST_Harness.AssertTrue {literal(SCRATCH_ASSERTION)}, "
              "Len(Detail) = 0, Detail",
              "    'Record the case's error, once Excel is restored",
              "        If ErrNumber <> 0 Then",
              "            Err.Number = ErrNumber",
              "            Err.Source = ErrSource",
              "            Err.Description = ErrDescription",
              "            TEST_Harness.RecordCaseError",
              "        End If",
              "    'Exit before the error handlers", "        Exit Sub", "",
              *section("ERROR HANDLER"), "Failed:",
              "    'Keep the case's error and close the workbook",
              "        ErrNumber = Err.Number", "        ErrSource = Err.Source",
              "        ErrDescription = Err.Description", "        Resume CleanUp", "",
              "CaseFailed:", "    'Record the unexpected error against the case",
              "        TEST_Harness.RecordCaseError", "", "End Sub", ""]
    return lines


def module(root: Path) -> str:
    """The text of the generated module, with CRLF line endings."""
    model = base.load_model(root)
    covered, excluded = cases(root, model)
    made = groups(covered)
    assertions = len(covered) + len(made)
    lines = [f'Attribute VB_Name = "{MODULE.stem}"', "", "Option Explicit",
             "Option Private Module", "", "'", base.RULE, f"' MODULE: {MODULE.stem}", base.RULE]
    body = [
        ("PURPOSE", ["The Excel oracle of the date layer: the worksheet-function cases that "
                     "Excel's own functions",
                     "can compute, computed with them and compared, as encoded text, with "
                     "their expected files"]),
        ("WHY THIS EXISTS", ["A second computation by another author catches an error that "
                             "the derivation of the",
                             "expected values and the date layer's code could share"]),
        ("INPUTS", ["None at module level"]),
        ("RETURNS", ["Nothing at module level"]),
        ("BEHAVIOR", [f"RunExcelOracleTests runs the suite: {len(made)} cases of up to {GROUP} "
                      "test cases of one function,",
                      f"{assertions} assertions: one per covered test case, named after it, "
                      f"and {SCRATCH_ASSERTION} per case.",
                      f"{len(excluded)} test cases are excluded, each for a stated reason "
                      "(docs/methodology/",
                      "EXCEL_ORACLE.md). CertifyExcelOracle certifies its normal run and "
                      "saves its log and its",
                      "certification log"]),
        ("ERROR POLICY", ["As TEST_Harness.EndRun: a run that does not pass raises "
                          "TEST_Harness.ERR_TEST_FAILED,",
                          "or the unexpected error that escaped the runner"]),
        ("DEPENDENCIES", ["TEST_WorksheetSupport, TEST_Harness; Excel's worksheet functions "
                          f"{', '.join(ALLOWED)}"]),
        ("NOTES", ["- GENERATED by tools/gen_excel_oracle_suite.py from tests/fixtures/ and "
                   "tests/expected/:",
                   "  never edit it by hand; regenerate it with --write",
                   "- It calls no procedure of the date layer: each result comes from "
                   "Worksheet.Evaluate",
                   "- Option Private Module keeps the suite out of the workbook's external "
                   "surface"]),
        ("UPDATED", [UPDATED]),
    ]
    for index, (name, text) in enumerate(body):
        lines.append(f"' {name}")
        lines += [f"'   {line}" for line in text]
        if index < len(body) - 1:
            lines.append("'")
    lines += [base.RULE, "", *section("PRIVATE CONSTANTS"),
              "    'A complete run starts every case and evaluates every assertion once",
              f'    Private Const SUITE_NAME            As String = "{SUITE_NAME}"'
              "      'Suite name in the run log",
              f"    Private Const EXPECTED_CASES        As Long = {len(made)}"
              "                  'Cases a complete run starts",
              f"    Private Const EXPECTED_ASSERTIONS   As Long = {assertions}"
              "                 'Assertions a complete run evaluates", "", "",
              "'", base.RULE, "'", "'" + "ENTRY POINTS".center(78).rstrip(), "'", base.RULE,
              "'", ""]
    lines += ["Public Sub RunExcelOracleTests()", ""]
    lines += header("RUN THE SUITE", [
        ("PURPOSE", ["Runs every case of the suite once, in its normal mode."]),
        ("WHY THIS EXISTS", ["The entry point to type in the Immediate window."]),
        ("INPUTS", ["None"]), ("RETURNS", ["Nothing"]),
        ("BEHAVIOR", ["Runs RunSuite."]),
        ("ERROR POLICY", ["As RunSuite"]),
        ("DEPENDENCIES", ["RunSuite"]), ("NOTES", ["None"]), ("UPDATED", [UPDATED])])
    lines += ["", *section("RUN"), "    'Run the suite", "        RunSuite", "", "End Sub", ""]
    lines += ["Public Sub CertifyExcelOracle( _", "    ByVal Folder As String, _",
              "    ByVal Fixtures As String)", ""]
    lines += header("CERTIFY THE SUITE", [
        ("PURPOSE", ["Certifies the suite's normal run and saves its evidence in the folder."]),
        ("WHY THIS EXISTS", ["The certification entry point that .github/"
                             "excel-evidence-policy.json names."]),
        ("INPUTS", ["Folder", "  The evidence folder.", "",
                    "Fixtures", "  The tests/fixtures/ folder, which this generated suite "
                    "does not read."]),
        ("RETURNS", ["Nothing"]),
        ("BEHAVIOR", ["Runs the suite, saves its log, certifies the run, then saves the "
                      "certification log."]),
        ("ERROR POLICY", ["TEST_Harness.EndCertification raises when the certification did "
                          "not pass."]),
        ("DEPENDENCIES", ["RunSuite, TEST_Harness"]), ("NOTES", ["None"]),
        ("UPDATED", [UPDATED])])
    lines += ["", *section("DECLARE"),
              "    Dim RunErrorNumber          As Long     'Error that ended the run; 0 for none",
              "    Dim RunErrorDescription     As String   'Its description", "",
              *section("CERTIFY"),
              "    'Refuse state left by an earlier run and start the certification",
              "        TEST_Harness.BeginCertification SUITE_NAME",
              "    'Run the suite and capture the error that ended it",
              "        On Error Resume Next", "        RunSuite",
              "        RunErrorNumber = Err.Number",
              "        RunErrorDescription = Err.Description",
              "    'Restore normal error handling", "        On Error GoTo 0",
              "    'Save the log and certify the run",
              f'        TEST_Harness.SaveLog Folder & "\\{SUITE_NAME}.log"',
              '        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, '
              "RunErrorDescription",
              "    'Save the certification log, raising when the certification did not pass",
              f'        TEST_Harness.EndCertification Folder & "\\{SUITE_NAME}-'
              'certification.log"', "", "End Sub", "",
              "'", base.RULE, "'", "'" + "RUNNER".center(78).rstrip(), "'", base.RULE, "'", ""]
    lines += ["Private Sub RunSuite()", ""]
    lines += header("RUN EVERY CASE", [
        ("PURPOSE", ["Runs every case once, in a fixed order, between BeginRun and EndRun."]),
        ("WHY THIS EXISTS", ["The certification reads the cases this runner calls."]),
        ("INPUTS", ["None"]), ("RETURNS", ["Nothing"]),
        ("BEHAVIOR", ["Starts the run, calls each case, then reports the run."]),
        ("ERROR POLICY", ["As TEST_Harness.EndRun"]),
        ("DEPENDENCIES", ["TEST_Harness"]), ("NOTES", ["None"]), ("UPDATED", [UPDATED])])
    lines += ["", *section("DECLARE"),
              "    Dim ErrNumber       As Long     'Runner error number; 0 for none",
              "    Dim ErrSource       As String   'Runner error source",
              "    Dim ErrDescription  As String   'Runner error description", "",
              *section("RUN"),
              "    'Start the run; a refused start propagates directly",
              '        TEST_Harness.BeginRun SUITE_NAME, "NORMAL", EXPECTED_CASES, '
              "EXPECTED_ASSERTIONS",
              "    'Send an error that escapes a case to the runner handler",
              "        On Error GoTo RunFailed", "    'Run every case in a fixed order"]
    lines += [f"        {name}" for name, _, _ in made]
    lines += ["", "Finish:", "    'Restore normal error handling", "        On Error GoTo 0",
              "    'Report the run, raising its failure or the escaped error",
              "        TEST_Harness.EndRun ErrNumber, ErrSource, ErrDescription", "",
              *section("EXIT PROCEDURE"), "    'Exit before the error handler",
              "        Exit Sub", "", *section("ERROR HANDLER"), "RunFailed:",
              "    'Keep the escaped error for EndRun", "        ErrNumber = Err.Number",
              "        ErrSource = Err.Source", "        ErrDescription = Err.Description",
              "        Resume Finish", "", "End Sub", "",
              "'", base.RULE, "'", "'" + "CASES".center(78).rstrip(), "'", base.RULE, "'", ""]
    for name, case_id, part in made:
        lines += case_procedure(name, case_id, part)
    while lines and lines[-1] == "":
        lines.pop()
    return "\r\n".join(lines) + "\r\n"


def coverage(root: Path) -> dict[str, Any]:
    """The covered and excluded cases, and the counts by function and reason."""
    covered, excluded = cases(root, base.load_model(root))
    counts: Counter[tuple[str, str]] = Counter()
    for case in covered:
        counts[(case["function"], "covered")] += 1
    for case in excluded:
        counts[(case["function"], case["reason"])] += 1
    return {"covered": [{**{key: case[key] for key in ("id", "formula", "expected")},
                         "rule": rule(case["function"], case["formula"])}
                        for case in covered],
            "excluded": [{key: case[key] for key in ("id", "reason")} for case in excluded],
            "counts": [{"function": function, "outcome": outcome, "cases": number}
                       for (function, outcome), number in sorted(counts.items())]}


def coverage_tables(root: Path) -> str:
    """The Markdown tables of the coverage, as docs/methodology/EXCEL_ORACLE.md holds them."""
    covered, excluded = cases(root, base.load_model(root))
    order = re.findall(r"^KPR_Dates\tFunction\t(\w+)$",
                       (root / base.API).read_text(encoding="utf-8"), re.M)
    lines = ["| Function | Covered | Excluded |", "| --- | --- | --- |"]
    for function in order:
        mine = sum(1 for case in covered if case["function"] == function)
        other = sum(1 for case in excluded if case["function"] == function)
        lines.append(f"| `{function}` | {mine} | {other} |")
    lines.append(f"| **Total** | **{len(covered)}** | **{len(excluded)}** |")
    lines += ["", "| Reason a case is excluded | Cases |", "| --- | --- |"]
    reasons = Counter(case["reason"] for case in excluded)
    lines += [f"| {reason[0].upper()}{reason[1:]} | {reasons[reason]} |"
              for reason in REASONS if reasons[reason]]
    return "\n".join(lines) + "\n"


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    action = parser.add_mutually_exclusive_group(required=True)
    action.add_argument("--write", action="store_true", help="regenerate the module")
    action.add_argument("--check", action="store_true", help="fail if the module is stale")
    action.add_argument("--list", action="store_true",
                        help="print the covered and excluded cases as JSON")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    root = args.root.resolve()
    if args.list:
        print(json.dumps(coverage(root), indent=1))
        return 0
    text = module(root)
    path = root / MODULE
    if args.write:
        path.write_bytes(text.encode("ascii"))
        print(f"excel-oracle-suite: wrote {MODULE.as_posix()}")
        return 0
    current = path.read_bytes().decode("ascii") if path.is_file() else None
    if current != text:
        print(f"FAIL: {MODULE.as_posix()} is stale; run python tools/gen_excel_oracle_suite.py "
              "--write")
        return 1
    print(f"excel-oracle-suite: PASS ({MODULE.as_posix()} matches the cases)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
