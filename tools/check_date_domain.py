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
# A bare 61, 1900 or 9999 cannot be forbidden: they are also an enum value and the 1900 date
# system, so only the typed serial 61# is read as a copy of the lower edge.
WINDOW_COPIES = re.compile(r"\b2958465\b|(?<![\w.])61(?:\.0*)?#|#\s*12/31/9999\s*#|#\s*3/1/1900\s*#"
                           r"|#\s*9999-12-31\s*#|#\s*1900-03-01\s*#")
# The components whose public procedures read or classify an argument: no call of theirs may
# come before the host rule, qualified by the module or, as VBA allows, by its name alone.
READER_MODULES = ("src/workbook/M_Ranges.bas", "src/core/CORE_Array.bas", "src/core/CORE_Parse.bas",
                  "src/core/CORE_Dates.bas")
WORKBOOK_FALLBACKS = re.compile(r"\b(?:ActiveWorkbook|ThisWorkbook|ActiveSheet|ActiveCell|Selection"
                                r"|Workbooks|Application\s*\.\s*(?:ActiveWorkbook|ActiveSheet))\b",
                                re.I)
# Conversions that depend on the Windows or Excel regional settings.
LOCALE_CALLS = re.compile(r"(?<![\w.])(?:VBA\s*\.\s*)?(IsDate|DateValue|CVDate|IsNumeric"
                          r"|FormatDateTime|Format\$?)\s*\(", re.I)
# The one statement that makes the diagnostic volatile, in any of VBA's spellings of it.
VOLATILE_TRUE = re.compile(r"\s*(?:Call\s+)?(?:Excel\s*\.\s*)?Application\s*\.\s*Volatile"
                           r"\s*(?:\(\s*True\s*\)|\s+True)\s*", re.I)
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


NUMBER = re.compile(r"[+-]?\s*(?:(?P<hex>&[hH][0-9A-Fa-f]+)|(?P<octal>&[oO][0-7]+)"
                    r"|(?P<decimal>(?:\d+\.?\d*|\.\d+)(?:[eEdD][+-]?\d+)?))[%&#!@^]*")


def literal_value(text: str) -> float | None:
    """The value of a numeric literal in VBA notation, or None when ``text`` is not one.

    It reads a sign, enclosing parentheses, a type suffix (``0&``, ``0#``), a decimal point, an
    exponent (``0E0``) and hexadecimal or octal (``&H0``)."""
    text = text.strip()
    while text.startswith("(") and text.endswith(")") and split_arguments(text[1:])[0] == text[1:-1]:
        text = text[1:-1].strip()
    match = NUMBER.fullmatch(text)
    if match is None:
        return None
    if match.group("hex"):
        value = float(int(match.group("hex")[2:], 16))
    elif match.group("octal"):
        value = float(int(match.group("octal")[2:], 8))
    else:
        value = float(match.group("decimal").replace("d", "e").replace("D", "e"))
    return -value if text.lstrip().startswith("-") else value


def day_zero(sources: Mapping[str, str]) -> list[Finding]:
    """No DateSerial with a literal day of zero or below."""
    findings = []
    for path, text in sources.items():
        for number, code in code_lines(text):
            for call in DATESERIAL_CALL.finditer(code):
                arguments = split_arguments(code[call.end():])
                day = literal_value(arguments[2]) if len(arguments) >= 3 else None
                if day is not None and day <= 0:
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
    pinned = {found for name in WINDOW if (found := constant_statement(parse, name)) is not None}
    for path, text in sources.items():
        for number, code in code_lines(text):
            for name in WINDOW:
                if re.search(rf"\bConst\s+{name}\b", code, re.I) and (
                        path != CORE_PARSE or constant_statement(parse, name) != (number, code)):
                    findings.append(Finding(path, number, f"{name} is declared again: the window "
                                            f"is declared once, in {CORE_PARSE}"))
            if WINDOW_COPIES.search(code) and not (path == CORE_PARSE and (number, code) in pinned):
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
    if name.casefold() == DIAGNOSTIC.casefold():
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


def argument_readers(sources: Mapping[str, str]) -> re.Pattern[str]:
    """A pattern for a call of a reader module: the module's name, or one of its public procedures."""
    names = [Path(path).stem for path in READER_MODULES]
    procedures = {procedure.name for path in READER_MODULES
                  for procedure in scan(path, sources.get(path, "")).procedures if procedure.public}
    alternatives = [rf"\b{name}\s*\." for name in names] + [rf"\b{name}\b" for name in
                                                              sorted(procedures)]
    return re.compile("|".join(alternatives), re.I)


def evaluator_findings(procedure: Procedure | None, readers: re.Pattern[str]) -> list[Finding]:
    """EvaluateCall applies the host rule once, before it reads any argument."""
    if procedure is None:
        return [Finding(FACADE, None, f"{EVALUATOR} is not declared")]
    body = executable(procedure)
    guards = calls_of(body, CHECK)
    reads = [Finding(FACADE, line, f"{EVALUATOR} reads the host itself: {CHECK} reads it once")
             for line in calls_of(body, READ)]
    if len(guards) != 1:
        return [*reads, Finding(FACADE, procedure.line, f"{EVALUATOR} must call {CHECK} exactly "
                                f"once; it calls it {len(guards)} times")]
    return reads + [Finding(FACADE, line, f"{EVALUATOR} reads an argument before {CHECK}: the host rule "
                    "comes first")
            for line, statement in body
            if line < guards[0] and readers.search(statement)]


def host_module_findings(text: str) -> list[Finding]:
    """In M_Host, TryCheckHost reads the host exactly once, through TryReadHostDateSystem."""
    module = scan(HOST, text)
    check = next((procedure for procedure in module.procedures
                  if procedure.name.casefold() == CHECK.casefold()), None)
    if check is None:
        return [Finding(HOST, None, f"{CHECK} is not declared")] if text else []
    reads = calls_of(executable(check), READ)
    if len(reads) == 1:
        return []
    return [Finding(HOST, check.line, f"{CHECK} must read the host exactly once, through {READ}; "
                    f"it calls it {len(reads)} times")]


def host_guard(sources: Mapping[str, str]) -> tuple[list[Finding], int]:
    """The host rule runs once per call, first, in the evaluator only."""
    module = scan(FACADE, sources.get(FACADE, ""))
    functions = [procedure for procedure in module.procedures
                 if procedure.public and procedure.kind == "function"
                 and procedure.name.casefold().startswith("kpr_dates_")]
    findings = [finding for procedure in functions for finding in function_findings(procedure)]
    findings += evaluator_findings(next((procedure for procedure in module.procedures
                                         if procedure.name.casefold() == EVALUATOR.casefold()), None),
                                   argument_readers(sources))
    for procedure in module.procedures:
        if procedure in functions or procedure.name.casefold() == EVALUATOR.casefold():
            continue
        body = executable(procedure)
        findings.extend(Finding(FACADE, line, f"{procedure.name} applies the host rule: only "
                                f"{EVALUATOR} may") for line in calls_of(body, CHECK))
        findings.extend(Finding(FACADE, line, f"{procedure.name} reads the host: only {DIAGNOSTIC} "
                                f"does, and {CHECK} reads it for the others")
                        for line in calls_of(body, READ))
    findings += host_module_findings(sources.get(HOST, ""))
    for path, text in sources.items():
        if path not in (FACADE, HOST):
            findings.extend(Finding(path, number, "the host rule is applied outside the facade")
                            for number, code in code_lines(text)
                            if re.search(rf"\b{CHECK}\b|\b{READ}\b", code, re.I))
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
                    for _ in pattern.finditer(statement)]
            total += len(hits)
            if procedure.name.casefold() == DIAGNOSTIC.casefold():
                first = executable(procedure)[:1]
                if len(hits) != 1:
                    findings.append(Finding(path, procedure.line, f"{DIAGNOSTIC} must call "
                                            f"Application.Volatile exactly once; it calls it "
                                            f"{len(hits)} times"))
                elif not first or first[0] != hits[0]:
                    findings.append(Finding(path, hits[0][0], "Application.Volatile must be the "
                                            f"first statement of {DIAGNOSTIC}"))
                elif not VOLATILE_TRUE.fullmatch(hits[0][1]):
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


CALLER = re.compile(r"\bApplication\s*\.\s*Caller\b", re.I)


def caller_findings(text: str) -> list[Finding]:
    """In M_Host, Application.Caller is read once, and only by TryReadHostDateSystem."""
    findings: list[Finding] = []
    total = 0
    for procedure in scan(HOST, text).procedures:
        count = sum(len(CALLER.findall(statement)) for _, statement in procedure.body)
        total += count
        if count and procedure.name.casefold() != READ.casefold():
            findings.append(Finding(HOST, procedure.line, f"{procedure.name} reads "
                                    f"Application.Caller: only {READ} does"))
    if text and total != 1:
        findings.append(Finding(HOST, None, f"Application.Caller must be read exactly once; it "
                                f"is read {total} times"))
    return findings


def host_authority(sources: Mapping[str, str]) -> list[Finding]:
    """Only M_Host reads Application.Caller, and no host-path component consults a workbook."""
    findings = caller_findings(sources.get(HOST, ""))
    for path, text in sources.items():
        for number, code in code_lines(text):
            for _ in CALLER.finditer(code):
                if path != HOST:
                    findings.append(Finding(path, number, "Application.Caller is read outside "
                                            f"{HOST}: only {READ} classifies the caller"))
            if path in (HOST, FACADE) and WORKBOOK_FALLBACKS.search(code):
                findings.append(Finding(path, number, "the host path consults the active or the "
                                        "containing workbook: only the caller's own workbook "
                                        "decides the date system"))
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
