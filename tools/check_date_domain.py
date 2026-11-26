#!/usr/bin/env python3
"""Check the date-domain rules of the date layer that the generic VBA checks do not cover.

Each rule reads the production source under src/ and enforces a statement of
the methodology guides (docs/methodology/VALIDATION_PLAN.md, Date-domain static
rules, lists them with the section each enforces):

- day-zero: no DateSerial call whose day is a literal of zero or below; the last
  day of a month comes from the month's length, because a day of zero in the
  month after December 9999 overflows the date range;
- window: the supported window is declared once, in CORE_Parse, as the serials
  61 and 2958465 and the years 1900 and 9999, and no other component holds a
  copy of its edges;
- host-guard: every worksheet function of the facade but KPR_Dates_HostDateSystem
  calls EvaluateCall exactly once and the host rule nowhere else, and
  EvaluateCall applies it exactly once, before it reads any argument;
- volatility: Application.Volatile is called exactly once in production, with
  True, as the first statement of KPR_Dates_HostDateSystem;
- host-authority: Application.Caller is read only by M_Host, once, and no
  component of the host path consults the active or the containing workbook;
- locale: no conversion that depends on regional settings (IsDate, DateValue,
  CVDate, IsNumeric, Format, FormatDateTime), and CDate only on a variable named
  ...Serial.

It reads comments out and string literals blank, so a rule is never tripped by
text. It does not compile or run VBA; the self-tests in tests/test_check_date_domain.py
show each rule rejecting a defect.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Mapping, Sequence
from pathlib import Path

from check_vba_identifiers import Procedure, scan
from vba_lexer import Finding, logical_units, read_component, report, tracked_components

ROOT = Path(__file__).resolve().parents[1]

CORE_PARSE = "src/core/CORE_Parse.bas"
FACADE = "src/modules/KPR_Dates.bas"
HOST = "src/workbook/M_Host.bas"
DIAGNOSTIC = "KPR_Dates_HostDateSystem"
EVALUATOR = "EvaluateCall"
CHECK = "TryCheckHost"
READ = "TryReadHostDateSystem"

# The frozen supported window, as the guides state it.
WINDOW = {
    "WINDOW_MIN_SERIAL": r"61#",
    "WINDOW_MAX_SERIAL": r"2958465#",
    "WINDOW_MIN_YEAR": r"1900",
    "WINDOW_MAX_YEAR": r"9999",
}
# The edges of the window written another way, in any component.
WINDOW_COPIES = re.compile(r"\b2958465\b|#\s*12/31/9999\s*#|#\s*3/1/1900\s*#|#\s*9999-12-31\s*#"
                           r"|#\s*1900-03-01\s*#")
# An argument that the facade reads or classifies: none may come before the host rule.
ARGUMENT_READERS = ("M_Ranges.", "CORE_Array.", "CORE_Parse.", "CORE_Dates.")
WORKBOOK_FALLBACKS = re.compile(r"\b(?:ActiveWorkbook|ThisWorkbook|ActiveSheet|ActiveCell|Selection"
                                r"|Workbooks|Application\s*\.\s*(?:ActiveWorkbook|ActiveSheet))\b",
                                re.I)
# Conversions that depend on the Windows or Excel regional settings.
LOCALE_CALLS = re.compile(r"(?<![\w.])(?:VBA\s*\.\s*)?(IsDate|DateValue|CVDate|IsNumeric"
                          r"|FormatDateTime|Format\$?)\s*\(", re.I)
CDATE_CALL = re.compile(r"(?<![\w.])(?:VBA\s*\.\s*)?CDate\s*\(", re.I)
DATESERIAL_CALL = re.compile(r"(?<![\w.])(?:VBA\s*\.\s*)?DateSerial\s*\(", re.I)
DECLARATION = re.compile(r"^\s*(?:Dim|Const|Static|Private|Public)\b", re.I)


def split_arguments(text: str) -> list[str]:
    """The top-level arguments of a call, from the text after its opening parenthesis."""
    depth, start, parts = 0, 0, []
    for index, char in enumerate(text):
        if char == "(":
            depth += 1
        elif char == ")":
            if not depth:
                parts.append(text[start:index])
                return parts
            depth -= 1
        elif char == "," and not depth:
            parts.append(text[start:index])
            start = index + 1
    parts.append(text[start:])
    return parts


def code_lines(text: str) -> list[tuple[int, str]]:
    """Each logical line without its comment, strings blanked, with its first line's number.

    A statement continued with " _" is one logical line, so a call written over several
    lines is read whole."""
    return [(first, code) for first, _, kind, code
            in logical_units(text.split("\n"), blank_strings=True) if kind == "code"]


def day_zero(sources: Mapping[str, str]) -> list[Finding]:
    """No DateSerial with a literal day of zero or below."""
    findings = []
    for path, text in sources.items():
        for number, code in code_lines(text):
            for call in DATESERIAL_CALL.finditer(code):
                arguments = split_arguments(code[call.end():])
                if len(arguments) >= 3 and re.fullmatch(r"\s*(?:-\s*\d+|0+)\s*", arguments[2]):
                    findings.append(Finding(path, number, "DateSerial with a day of zero or below "
                                            "overflows after December 9999: take the last day "
                                            "of a month from the month's length"))
    return findings


def constant_statement(text: str, name: str) -> tuple[int, str] | None:
    for number, code in code_lines(text):
        if re.search(rf"\bConst\s+{re.escape(name)}\b", code, re.I):
            return number, code
    return None


def window(sources: Mapping[str, str]) -> list[Finding]:
    """The window's four constants are pinned in CORE_Parse, and nothing else copies its edges."""
    findings = []
    parse = sources.get(CORE_PARSE, "")
    for name, value in WINDOW.items():
        found = constant_statement(parse, name)
        if found is None:
            findings.append(Finding(CORE_PARSE, None, f"{name} is not declared"))
        elif not re.search(rf"=\s*{value}\s*$", found[1]):
            findings.append(Finding(CORE_PARSE, found[0], f"{name} is no longer {value.rstrip('#')}, "
                                    "the edge of the supported window"))
    for path, text in sources.items():
        for number, code in code_lines(text):
            if WINDOW_COPIES.search(code) and not (path == CORE_PARSE and DECLARATION.match(code)):
                findings.append(Finding(path, number, "an edge of the supported window is written "
                                        "out: use CORE_Parse.WINDOW_MIN_SERIAL, WINDOW_MAX_SERIAL, "
                                        "WINDOW_MIN_YEAR or WINDOW_MAX_YEAR"))
    return findings


def executable(procedure: Procedure) -> list[tuple[int, str]]:
    """The statements of a body that run something: no declaration."""
    return [(line, statement) for line, statement in procedure.body
            if statement.strip() and not DECLARATION.match(statement)]


def calls_of(statements: Sequence[tuple[int, str]], name: str) -> list[int]:
    """The lines whose statements call ``name``, qualified or not, never as an assignment target."""
    pattern = re.compile(rf"(?<![\w.])(?:[A-Za-z_]\w*\s*\.\s*)?{name}\s*\(?", re.I)
    found = []
    for line, statement in statements:
        for match in pattern.finditer(statement):
            if not re.match(rf"\s*{name}\s*=", statement[match.start():], re.I):
                found.append(line)
    return found


def function_findings(procedure: Procedure) -> list[Finding]:
    """What one worksheet function of the facade may call of the host rule."""
    body = executable(procedure)
    name = procedure.name
    if name == DIAGNOSTIC:
        if len(calls_of(body, READ)) == 1 and not calls_of(body, CHECK) \
                and not calls_of(body, EVALUATOR):
            return []
        return [Finding(FACADE, procedure.line, f"{name} must call {READ} exactly once, and leave "
                        f"the host rule to {EVALUATOR}")]
    findings = []
    if len(calls_of(body, EVALUATOR)) != 1:
        findings.append(Finding(FACADE, procedure.line, f"{name} must call {EVALUATOR} exactly once"))
    if calls_of(body, CHECK) or calls_of(body, READ):
        findings.append(Finding(FACADE, procedure.line, f"{name} must leave the host rule to "
                                f"{EVALUATOR}"))
    return findings


def evaluator_findings(procedure: Procedure | None) -> list[Finding]:
    """EvaluateCall applies the host rule once, before it reads any argument."""
    if procedure is None:
        return [Finding(FACADE, None, f"{EVALUATOR} is not declared")]
    body = executable(procedure)
    guards = calls_of(body, CHECK)
    if len(guards) != 1:
        return [Finding(FACADE, procedure.line, f"{EVALUATOR} must call {CHECK} exactly once; it "
                        f"calls it {len(guards)} times")]
    return [Finding(FACADE, line, f"{EVALUATOR} reads an argument before {CHECK}: the host rule "
                    "comes first")
            for line, statement in body
            if line < guards[0] and any(reader in statement for reader in ARGUMENT_READERS)]


def host_guard(sources: Mapping[str, str]) -> tuple[list[Finding], int]:
    """The host rule runs once per call, first, in the evaluator only."""
    module = scan(FACADE, sources.get(FACADE, ""))
    functions = [procedure for procedure in module.procedures
                 if procedure.public and procedure.kind == "function"
                 and procedure.name.startswith("KPR_Dates_")]
    findings = [finding for procedure in functions for finding in function_findings(procedure)]
    findings += evaluator_findings(next((procedure for procedure in module.procedures
                                         if procedure.name == EVALUATOR), None))
    for procedure in module.procedures:
        if procedure not in functions and procedure.name != EVALUATOR:
            findings.extend(Finding(FACADE, line, f"{procedure.name} applies the host rule: only "
                                    f"{EVALUATOR} may")
                            for line in calls_of(executable(procedure), CHECK))
    for path, text in sources.items():
        if path not in (FACADE, HOST):
            findings.extend(Finding(path, number, "the host rule is applied outside the facade")
                            for number, code in code_lines(text)
                            if re.search(rf"\b{CHECK}\b|\b{READ}\b", code))
    return findings, len(functions)


def volatility(sources: Mapping[str, str]) -> list[Finding]:
    """Application.Volatile True, once, first, in the diagnostic."""
    findings: list[Finding] = []
    pattern = re.compile(r"\bApplication\s*\.\s*Volatile\b", re.I)
    total = 0
    for path, text in sources.items():
        module = scan(path, text)
        for procedure in module.procedures:
            hits = [(line, statement) for line, statement in executable(procedure)
                    if pattern.search(statement)]
            total += len(hits)
            if procedure.name == DIAGNOSTIC:
                first = executable(procedure)[:1]
                if len(hits) != 1:
                    findings.append(Finding(path, procedure.line, f"{DIAGNOSTIC} must call "
                                            f"Application.Volatile exactly once; it calls it "
                                            f"{len(hits)} times"))
                elif not first or first[0] != hits[0]:
                    findings.append(Finding(path, hits[0][0], "Application.Volatile must be the "
                                            f"first statement of {DIAGNOSTIC}"))
                elif not re.search(r"Volatile\s*\(?\s*True\b", hits[0][1], re.I):
                    findings.append(Finding(path, hits[0][0], "Application.Volatile must be called "
                                            "with True"))
            else:
                findings.extend(Finding(path, line, f"{procedure.name} must not be volatile: only "
                                        f"{DIAGNOSTIC} calls Application.Volatile")
                                for line, _ in hits)
    if total != 1 and not any(f.message.startswith(DIAGNOSTIC) for f in findings):
        findings.append(Finding(FACADE, None, "Application.Volatile must occur exactly once in "
                                f"production; it occurs {total} times"))
    return findings


def host_authority(sources: Mapping[str, str]) -> list[Finding]:
    """Only M_Host reads Application.Caller, and no host-path component consults a workbook."""
    findings: list[Finding] = []
    caller = re.compile(r"\bApplication\s*\.\s*Caller\b", re.I)
    readers = 0
    for path, text in sources.items():
        for number, code in code_lines(text):
            if caller.search(code):
                if path == HOST:
                    readers += 1
                else:
                    findings.append(Finding(path, number, "Application.Caller is read outside "
                                            f"{HOST}: only {READ} classifies the caller"))
            if path in (HOST, FACADE) and WORKBOOK_FALLBACKS.search(code):
                findings.append(Finding(path, number, "the host path consults the active or the "
                                        "containing workbook: only the caller's own workbook "
                                        "decides the date system"))
    if sources.get(HOST) is not None and readers != 1:
        findings.append(Finding(HOST, None, "Application.Caller must be read exactly once; it is "
                                f"read {readers} times"))
    return findings


def locale(sources: Mapping[str, str]) -> list[Finding]:
    """No regional-settings conversion; CDate only on a serial."""
    findings = []
    for path, text in sources.items():
        for number, code in code_lines(text):
            for hit in LOCALE_CALLS.finditer(code):
                findings.append(Finding(path, number, f"{hit.group(1)} depends on regional "
                                        "settings: read dates and numbers by their own rules"))
            for hit in CDATE_CALL.finditer(code):
                argument = split_arguments(code[hit.end():])[0].strip()
                if not re.fullmatch(r"\w*Serial", argument, re.I):
                    findings.append(Finding(path, number, "CDate may only convert a serial, "
                                            f"a variable named ...Serial, not {argument or 'nothing'}"))
    return findings


def production(root: Path) -> dict[str, str]:
    """The tracked components under src/, with LF line ends."""
    return {path: read_component(root, path) for path in tracked_components(root)
            if path.startswith("src/")}


def analyze(sources: Mapping[str, str]) -> tuple[list[Finding], int]:
    """Every finding, and how many public worksheet functions the host rule covers."""
    guard, functions = host_guard(sources)
    findings = [*day_zero(sources), *window(sources), *guard, *volatility(sources),
                *host_authority(sources), *locale(sources)]
    return findings, functions


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    sources = production(args.root.resolve())
    findings, functions = analyze(sources)
    return report("date-domain", findings, f"{len(sources)} production component(s), "
                  f"{functions} public function(s), 6 rules")


if __name__ == "__main__":
    sys.exit(main())
