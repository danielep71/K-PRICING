#!/usr/bin/env python3
"""Generate tests/modules/TEST_DateCases.bas, the suite of the worksheet-function cases.

The suite replays every worksheet-function case of tests/fixtures/ whose caller is
VBA (context "direct") on the KPR_Dates facade, in a development workbook: it
builds the case's arguments, calls the function, encodes the result as the
expected files do, asserts that the text equals the expected one, and records it
for the exact accuracy contracts (docs/methodology/TEST_CASES.md). A case of a
worksheet caller needs a worksheet, so another suite runs it.

The certification reads each case id and assertion name as a literal of the case
procedure (docs/methodology/TEST_HARNESS.md, Modules), so the suite is
generated rather than reading the fixtures at run time. Each Case procedure runs
up to GROUP cases of one function; an assertion is named after its case.

    python tools/gen_date_case_suite.py --write   # regenerate the module
    python tools/gen_date_case_suite.py --check   # fail if the module is stale
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import math
import re
import sys
from collections.abc import Sequence
from pathlib import Path
from types import ModuleType
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = Path("tests/fixtures")
EXPECTED = Path("tests/expected")
MODEL = Path("tests/derivations/date_layer_reference.py")
API = Path("docs/PUBLIC_API.txt")
MODULE = Path("tests/modules/TEST_DateCases.bas")
SUITE_NAME = "date-cases"
GROUP = 30
UPDATED = "2026-11-26"
RULE = "'" + "-" * 78
COMMENT_WIDTH = 99
ERRORS = {"#NULL!": 2000, "#DIV/0!": 2007, "#VALUE!": 2015, "#REF!": 2023, "#NAME?": 2029,
          "#NUM!": 2036, "#N/A": 2042}
LONG_MIN = -2147483648


def load_model(root: Path) -> ModuleType:
    """The reference model, for its value decoder and its table of contracts."""
    spec = importlib.util.spec_from_file_location("date_layer_reference", root / MODEL)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def number(payload: str) -> str:
    """A VBA Double literal of the number a num value holds."""
    value = float(payload)
    if value.is_integer() and abs(value) < 1e15:
        return f"{int(value)}#"
    text = repr(value)
    if "e" in text:
        mantissa, exponent = text.split("e")
        return f"{mantissa}E{int(exponent):+d}"
    return f"{text}#"


def expression(value: Any) -> str:
    """The VBA expression that builds one decoded value."""
    kind, payload = str(value.kind), str(value.payload)
    if kind == "date":
        year, month, day = (int(part) for part in payload.split("-"))
        return f"DateSerial({year}, {month}, {day})"
    if kind == "datetime":
        date, time = payload.split("T")
        year, month, day = (int(part) for part in date.split("-"))
        hour, minute, second = (int(part) for part in time.split(":"))
        return f"(DateSerial({year}, {month}, {day}) + TimeSerial({hour}, {minute}, {second}))"
    if kind == "num":
        return number(payload)
    if kind == "lng":
        whole = int(payload)
        return "(-2147483647& - 1&)" if whole == LONG_MIN else f"{whole}&"
    if kind == "str":
        return '"' + payload.replace('"', '""') + '"'
    if kind == "bool":
        return "True" if payload == "TRUE" else "False"
    if kind == "empty":
        return "Empty"
    if kind == "null":
        return "Null"
    if kind == "err":
        return f"CVErr({ERRORS[payload]})"
    items = ", ".join(expression(item) for item in value.items)
    if kind == "arr1":
        return f"TEST_DateCaseSupport.A1({items})"
    if kind == "arr2":
        return f"TEST_DateCaseSupport.A2({value.rows}, {value.cols}, {items})"
    if kind == "fill2":
        return f"TEST_DateCaseSupport.F2({value.rows}, {value.cols}, {items})"
    raise ValueError(f"no VBA expression for a {kind} value")


def literal(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def comment(text: str) -> str:
    """One intent comment of at most COMMENT_WIDTH columns, in ASCII."""
    clean = text.encode("ascii", "replace").decode("ascii").strip()
    prefix = "    '"
    if len(prefix) + len(clean) > COMMENT_WIDTH:
        clean = clean[:COMMENT_WIDTH - len(prefix) - 3].rsplit(" ", 1)[0] + "..."
    return prefix + clean


def cases(root: Path, model: ModuleType) -> list[dict[str, Any]]:
    """The direct worksheet-function cases, by function in inventory order, then by id."""
    order = re.findall(r"^KPR_Dates\tFunction\t(\w+)$",
                       (root / API).read_text(encoding="utf-8"), re.M)
    found = []
    for path in sorted((root / FIXTURES).glob("*.json")):
        fixture = json.loads(path.read_text(encoding="utf-8"))
        if fixture.get("kind") != "test-call-fixture" or fixture["context"] != "direct":
            continue
        expected = json.loads((root / EXPECTED / path.name).read_text(encoding="utf-8"))
        quantity = model.CONTRACTS[fixture["function"]][1]
        result = next(output["text"] for output in expected["outputs"]
                      if output["quantity"] == quantity)
        found.append({"id": fixture["id"], "function": fixture["function"],
                      "description": fixture["description"], "arguments": fixture["arguments"],
                      "quantity": quantity, "expected": result})
    return sorted(found, key=lambda case: (order.index(case["function"]), case["id"]))


def groups(found: list[dict[str, Any]]) -> list[tuple[str, str, list[dict[str, Any]]]]:
    """(procedure, case id, cases): up to GROUP cases of one function each."""
    made = []
    functions = list(dict.fromkeys(case["function"] for case in found))
    for function in functions:
        mine = [case for case in found if case["function"] == function]
        short = function.removeprefix("KPR_Dates_")
        count = math.ceil(len(mine) / GROUP)
        for index in range(count):
            part = mine[index * GROUP:(index + 1) * GROUP]
            name = f"Case{short}{index + 1:02d}"
            kebab = re.sub(r"(?<!^)(?=[A-Z])", "-", short).lower()
            made.append((name, f"{kebab}-{index + 1:02d}", part))
    return made


def header(title: str, fields: list[tuple[str, list[str]]]) -> list[str]:
    """A procedure's documentation header in the house style."""
    lines = ["'", RULE, "'" + title.center(78).rstrip(), RULE]
    for index, (name, body) in enumerate(fields):
        lines.append(f"' {name}")
        lines += [f"'   {line}" if line else "'" for line in body]
        if index < len(fields) - 1:
            lines.append("'")
    return lines + [RULE]


def section(name: str) -> list[str]:
    return [RULE, f"' {name}", RULE]


def case_procedure(name: str, case_id: str, part: list[dict[str, Any]],
                   model: ModuleType) -> list[str]:
    function = part[0]["function"]
    lines = [f"Private Sub {name}()", ""]
    lines += header(f"CASES OF {function.removeprefix('KPR_Dates_').upper()}", [
        ("PURPOSE", [f"Replays {len(part)} worksheet-function cases of {function}, from "
                     f"{part[0]['id']}",
                     f"to {part[-1]['id']}, called from VBA."]),
        ("WHY THIS EXISTS", ["Each assertion compares one case's encoded result with its "
                             "expected file."]),
        ("INPUTS", ["None"]),
        ("RETURNS", ["Nothing"]),
        ("BEHAVIOR", ["Calls the function with each case's arguments, encodes the result, "
                      "asserts it",
                      "under the case's id and records it for the exact accuracy contracts."]),
        ("ERROR POLICY", ["An unexpected error is recorded against the case by",
                          "TEST_Harness.RecordCaseError and does not stop the run"]),
        ("DEPENDENCIES", [f"{function}, TEST_DateCaseSupport, TEST_Harness"]),
        ("NOTES", ["Generated by tools/gen_date_case_suite.py from tests/fixtures/"]),
        ("UPDATED", [UPDATED]),
    ])
    lines += ["", *section("DECLARE"),
              "    Dim Actual          As String       'The encoded result of the case", "",
              *section("INITIALIZE"),
              "    'Record an unexpected error against the case",
              "        On Error GoTo CaseFailed",
              "    'Start the case",
              f"        TEST_Harness.BeginCase {literal(case_id)}", "",
              *section("RUN CASES")]
    for case in part:
        values = [model.decode(text) for text in case["arguments"]]
        call = f"{function}({', '.join(expression(value) for value in values)})"
        encode = "EncodeFilled" if case["expected"].startswith("fill2:") else "Encode"
        lines += [comment(case["description"]),
                  f"        Actual = TEST_DateCaseSupport.{encode}({call})",
                  f"        TEST_Harness.AssertEqualText {literal(case['id'])}, "
                  f"{literal(case['expected'])}, Actual",
                  f"        TEST_Harness.RecordTextResult {literal(case['id'])}, vbNullString, "
                  f"{literal(case['quantity'])}, vbNullString, vbNullString, Actual"]
    lines += ["", *section("EXIT PROCEDURE"), "    'Exit before the error handler",
              "        Exit Sub", "", *section("ERROR HANDLER"), "CaseFailed:",
              "    'Record the unexpected error against the case",
              "        TEST_Harness.RecordCaseError", "", "End Sub", ""]
    return lines


# ---------------------------------------------------------------------------
# The house style of the generated modules (docs/VBA_HOUSE_STYLE.md, Layout).

SIGNATURE = re.compile(r"^(?:Public|Private) Sub (\w+)\(")
DECLARATION = re.compile(r"(\s*(?:(?:Public|Private|Dim)\s+(?:Const\s+)?|Const\s+)\w+)(\s+)(As\s+.*)$")
STATEMENT_WIDTH = 100


def code_and_comment(line: str) -> tuple[str, str]:
    """A line's code and its trailing comment, which starts at an apostrophe outside strings."""
    quoted = False
    for index, char in enumerate(line):
        if char == '"':
            quoted = not quoted
        elif char == "'" and not quoted:
            return line[:index], line[index:]
    return line, ""


def commas_at(code: str, level: int) -> list[int]:
    """The positions of the commas outside strings, ``level`` parentheses deep."""
    found, depth, quoted = [], 0, False
    for index, char in enumerate(code):
        if char == '"':
            quoted = not quoted
        elif not quoted and char in "()":
            depth += 1 if char == "(" else -1
        elif not quoted and char == "," and depth == level:
            found.append(index)
    return found


def wrapped(line: str) -> list[str]:
    """A line longer than STATEMENT_WIDTH, broken into lines of at most that width where it
    can be. A comment line wraps at spaces under its own prefix; a statement wraps after commas
    or opening parentheses outside strings, each continuation indented four more spaces."""
    code, comment = code_and_comment(line)
    if len(line) <= STATEMENT_WIDTH:
        return [line]
    if not code.strip():
        prefix = comment[:len(comment) - len(comment[1:].lstrip(" "))]
        out, current = [], prefix
        for word in comment[len(prefix):].split(" "):
            if current != prefix and len(current) + len(word) + 1 > STATEMENT_WIDTH:
                out.append(current)
                current = prefix + word
            else:
                current += ("" if current == prefix else " ") + word
        return [*out, current]
    if comment:
        return [line]
    indent = " " * (len(line) - len(line.lstrip(" ")) + 4)
    # A break may follow any comma or opening parenthesis outside strings.
    quoted, openings = False, []
    for index, char in enumerate(code):
        if char == '"':
            quoted = not quoted
        elif char == "(" and not quoted:
            openings.append(index)
    commas = sorted({*openings, *(comma for level in range(8) for comma in commas_at(code, level))})
    out, start, current = [], 0, ""
    for comma in [*commas, len(code)]:
        piece = code[start:comma + 1] if comma < len(code) else code[start:]
        if current and len(current) + len(piece) + 2 > STATEMENT_WIDTH:
            out.append(current.rstrip() + " _")
            current = indent + piece.lstrip()
        else:
            current += piece
        start = comma + 1
    return [*out, current.rstrip()]


def aligned(lines: list[str]) -> list[str]:
    """Each run of one-line declarations with its As and trailing comments in the columns of
    its first declaration, or further right where a longer one needs it."""
    out = list(lines)
    index = 0
    while index < len(out):
        run = []
        while index < len(out) and DECLARATION.match(code_and_comment(out[index])[0].rstrip()):
            run.append(index)
            index += 1
        if not run:
            index += 1
            continue
        parts = []
        for at in run:
            code, comment = code_and_comment(out[at])
            match = DECLARATION.match(code.rstrip())
            assert match is not None
            parts.append((match.group(1), match.group(3).rstrip(), comment))
        first = DECLARATION.match(code_and_comment(out[run[0]])[0].rstrip())
        assert first is not None
        as_column = max(first.start(3), *(len(head) + 1 for head, _, _ in parts))
        codes = [head.ljust(as_column) + tail for head, tail, _ in parts]
        first_code, first_comment = code_and_comment(out[run[0]])
        comment_column = max(len(first_code) if first_comment else 0,
                             *(len(code) + 1 for code in codes))
        for at, code, (_, _, comment) in zip(run, codes, parts, strict=True):
            out[at] = code.ljust(comment_column) + comment if comment else code
    return out


def house_style(lines: list[str], module: str) -> list[str]:
    """The generated lines in the house style: each DECLARE starts with the procedure's
    PROC_NAME, a signature on several lines is followed by its header directly, the runner
    starts the run and sets its handler in INITIALIZE, long statements wrap, and each run of
    declarations is aligned."""
    out: list[str] = []
    procedure = ""
    index = 0
    while index < len(lines):
        line = lines[index]
        found = SIGNATURE.match(line)
        if found:
            procedure = found.group(1)
        if line == "" and out and out[-1].endswith(")") and out[-1].startswith("    ByVal ") \
                and index + 1 < len(lines) and lines[index + 1] == "'":
            index += 1
            continue
        out.append(line)
        if line == "' DECLARE" and out[-2] == RULE:
            out += [lines[index + 1],
                    f'    Const PROC_NAME     As String = "{module}.{procedure}"', ""]
            index += 2
            continue
        if line == "' RUN" and procedure == "RunSuite" and lines[index + 2].startswith(
                "    'Start the run"):
            out[-1] = "' INITIALIZE"
            out += lines[index + 1:index + 6]
            out += ["", RULE, "' RUN CASES", RULE]
            index += 6
            continue
        index += 1
    return aligned([piece for line in out for piece in wrapped(line)])


def module(root: Path) -> str:
    """The text of the generated module, with CRLF line endings."""
    model = load_model(root)
    found = cases(root, model)
    made = groups(found)
    lines = [f'Attribute VB_Name = "{MODULE.stem}"', "", "Option Explicit",
             "Option Private Module", "", "'", RULE, f"' MODULE: {MODULE.stem}", RULE]
    body = [
        ("PURPOSE", ["The worksheet-function cases of the date layer whose caller is VBA, replayed "
                     "on the",
                     "KPR_Dates facade: each result is compared, as encoded text, with its "
                     "expected file"]),
        ("WHY THIS EXISTS", ["The cases of tests/fixtures/ and the exact accuracy contracts "
                             "ACC-02 to ACC-23 are",
                             "measured on the workbook only through a suite that calls the "
                             "functions in Excel"]),
        ("INPUTS", ["None at module level"]),
        ("RETURNS", ["Nothing at module level"]),
        ("BEHAVIOR", [f"RunDateCaseTests runs the suite: {len(made)} cases of up to {GROUP} "
                      "test cases of one function,",
                      f"{len(found)} assertions, one per test case, named after it. "
                      "CertifyDateCases certifies its",
                      "normal run and saves its log, its certification log and its "
                      "results file"]),
        ("ERROR POLICY", ["As TEST_Harness.EndRun: a run that does not pass raises "
                          "TEST_Harness.ERR_TEST_FAILED,",
                          "or the unexpected error that escaped the runner"]),
        ("DEPENDENCIES", ["KPR_Dates, TEST_DateCaseSupport, TEST_Harness"]),
        ("NOTES", ["- GENERATED by tools/gen_date_case_suite.py from tests/fixtures/ and "
                   "tests/expected/:",
                   "  never edit it by hand; regenerate it with --write",
                   "- A case of a worksheet caller (ws1900, ws1904) needs a worksheet and "
                   "is not here",
                   "- Option Private Module keeps the suite out of the workbook's external "
                   "surface"]),
        ("UPDATED", [UPDATED]),
    ]
    for index, (name, text) in enumerate(body):
        lines.append(f"' {name}")
        lines += [f"'   {line}" for line in text]
        if index < len(body) - 1:
            lines.append("'")
    lines += [RULE, "", *section("PRIVATE CONSTANTS"),
              "    'A complete run starts every case and evaluates every assertion once",
              f'    Private Const SUITE_NAME            As String = "{SUITE_NAME}"'
              "        'Suite name in the run log",
              f"    Private Const EXPECTED_CASES        As Long = {len(made)}"
              "                  'Cases a complete run starts",
              f"    Private Const EXPECTED_ASSERTIONS   As Long = {len(found)}"
              "                'Assertions a complete run evaluates", "", "",
              "'", RULE, "'", "'" + "ENTRY POINTS".center(78).rstrip(), "'", RULE, "'", ""]
    lines += ["Public Sub RunDateCaseTests()", ""]
    lines += header("RUN THE SUITE", [
        ("PURPOSE", ["Runs every case of the suite once, in its normal mode."]),
        ("WHY THIS EXISTS", ["The entry point to type in the Immediate window."]),
        ("INPUTS", ["None"]), ("RETURNS", ["Nothing"]),
        ("BEHAVIOR", ["Runs RunSuite."]),
        ("ERROR POLICY", ["As RunSuite"]),
        ("DEPENDENCIES", ["RunSuite"]), ("NOTES", ["None"]), ("UPDATED", [UPDATED])])
    lines += ["", *section("RUN"), "    'Run the suite", "        RunSuite", "", "End Sub", ""]
    lines += ["Public Sub CertifyDateCases( _", "    ByVal Folder As String, _",
              "    ByVal Fixtures As String)", ""]
    lines += header("CERTIFY THE SUITE", [
        ("PURPOSE", ["Certifies the suite's normal run and saves its evidence in the folder."]),
        ("WHY THIS EXISTS", ["The certification entry point that .github/"
                             "excel-evidence-policy.json names."]),
        ("INPUTS", ["Folder", "  The evidence folder.", "",
                    "Fixtures", "  The tests/fixtures/ folder, which this generated suite "
                    "does not read."]),
        ("RETURNS", ["Nothing"]),
        ("BEHAVIOR", ["Runs the suite, saves its log, certifies the run, saves its results "
                      "and then the",
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
              "    'Save the log, certify the run and save its results",
              f'        TEST_Harness.SaveLog Folder & "\\{SUITE_NAME}.log"',
              '        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, '
              "RunErrorDescription",
              f'        TEST_Harness.SaveResults Folder & "\\{SUITE_NAME}-results.json"',
              "    'Save the certification log, raising when the certification did not pass",
              f'        TEST_Harness.EndCertification Folder & "\\{SUITE_NAME}-'
              'certification.log"', "", "End Sub", "",
              "'", RULE, "'", "'" + "RUNNER".center(78).rstrip(), "'", RULE, "'", ""]
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
              "'", RULE, "'", "'" + "CASES".center(78).rstrip(), "'", RULE, "'", ""]
    for name, case_id, part in made:
        lines += case_procedure(name, case_id, part, model)
    while lines and lines[-1] == "":
        lines.pop()
    return "\r\n".join(house_style(lines, MODULE.stem)) + "\r\n"


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    action = parser.add_mutually_exclusive_group(required=True)
    action.add_argument("--write", action="store_true", help="regenerate the module")
    action.add_argument("--check", action="store_true", help="fail if the module is stale")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    root = args.root.resolve()
    text = module(root)
    path = root / MODULE
    if args.write:
        path.write_bytes(text.encode("ascii"))
        print(f"date-case-suite: wrote {MODULE.as_posix()}")
        return 0
    current = path.read_bytes().decode("ascii") if path.is_file() else None
    if current != text:
        print(f"FAIL: {MODULE.as_posix()} is stale; run python tools/gen_date_case_suite.py "
              "--write")
        return 1
    print(f"date-case-suite: PASS ({MODULE.as_posix()} matches the cases)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
