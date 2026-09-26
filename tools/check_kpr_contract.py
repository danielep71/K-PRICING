#!/usr/bin/env python3
"""Validate the migrated KPR date-layer contract inside K-PRICING.

This is a project-specific additive gate. The generic K-PRICING repository,
VBA-structure, conditional-compilation and public-API validators remain
authoritative for repository-wide policy.
"""

from __future__ import annotations

import argparse
import copy
import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

SCHEMA_VERSION = 1
TOOL_NAME = "KPR migrated date-layer contract"
CONFIG_PATH = ".github/repository-profile.json"
MANIFEST_PATH = "docs/PUBLIC_API.txt"
CONTRACT_PATH = "docs/DATE_LAYER_CONTRACT.md"

EXPECTED_COMPONENTS = {
    "src/core/KPR_Core_Err.bas": "internal",
    "src/core/KPR_Core_Parse.bas": "internal",
    "src/core/KPR_Core_Dates.bas": "internal",
    "src/core/KPR_Core_Array.bas": "internal",
    "src/modules/KPR_DATES_DAYS.bas": "public",
    "tests/modules/KPR_REGRESSION_TESTS.bas": "test",
}
ALLOWED_DEPENDENCIES = {
    "kpr_core_err": frozenset(),
    "kpr_core_dates": frozenset({"kpr_core_err"}),
    "kpr_core_parse": frozenset({"kpr_core_err"}),
    "kpr_core_array": frozenset({"kpr_core_err"}),
    "kpr_dates_days": frozenset(
        {"kpr_core_err", "kpr_core_parse", "kpr_core_dates", "kpr_core_array"}
    ),
    "kpr_regression_tests": frozenset(
        {
            "kpr_core_err",
            "kpr_core_parse",
            "kpr_core_dates",
            "kpr_core_array",
            "kpr_dates_days",
            "kpr_test_fixtures_generated",
            "kpr_test_oracle",
        }
    ),
    # Generated expectations must stay independent of every production module.
    "kpr_test_fixtures_generated": frozenset(),
    # The Excel cross-oracle compares the public facade only.
    "kpr_test_oracle": frozenset({"kpr_dates_days"}),
}
REQUIRED_MEMBERS = {
    "kpr_core_err": frozenset({"ErrValue", "ErrNum", "ErrNA", "ErrForCondition"}),
    "kpr_core_array": frozenset(
        {
            "Array_Rank",
            "TryUnwrapScalar",
            "TryUnwrapControl",
            "TryClassifyShape",
            "CheckCapacity",
            "TryMaterialize",
            "AccumulateShape",
            "ElementAt",
            "TryAllocateOutput",
        }
    ),
    "kpr_core_parse": frozenset(
        {"TryParseDateScalar", "TryParseLongScalar", "TryParseBoolControl", "TryParseRoundingControl"}
    ),
    "kpr_core_dates": frozenset(
        {
            "KPR_MIN_DATE",
            "KPR_MAX_DATE",
            "IsDateInWindow",
            "IsLeapYear",
            "DaysInMonth",
            "EndOfMonth",
            "BeginOfQuarter",
            "EndOfQuarter",
            "BeginOfYear",
            "EndOfYear",
            "TryAddMonths",
            "TryPillar_Parse",
            "TryPillar_Format",
        }
    ),
    "kpr_dates_days": frozenset({"KPR_Dates_HostDateSystem"}),
    "kpr_test_fixtures_generated": frozenset(
        {"KPR_Fixtures_Count", "KPR_Fixtures_Case", "KPR_Fixtures_SourceHash"}
    ),
    "kpr_test_oracle": frozenset({"KPR_Oracle_RunCases"}),
    "kpr_regression_tests": frozenset({
        "KPR_Tests_Run",
        "KPR_Tests_RunSuite",
        "KPR_Tests_RunAll",
        "KPR_Tests_RunEvidence",
        "KPR_Tests_RunMigrationEvidence",
        "KPR_Tests_RunHost",
        "KPR_Tests_RunShape",
        "KPR_Tests_RunArray",
        "KPR_Tests_RunFixtureHost",
        "KPR_Tests_RunStateCheck",
        "KPR_Tests_RunOracle",
        "KPR_Test_RunAll",
        "KPR_Test_RunSuite",
    }),
}
FORBIDDEN_PARSE_CALLS = ("IsDate", "DateValue", "CVDate", "IsNumeric")
HOST_GUARD = "PassHostGuard"
HOST_CLASSIFIER = "TryResolveHostDateSystem"
HOST_DIAGNOSTIC = "KPR_Dates_HostDateSystem"
VOLATILE_CALL = "Application.Volatile True"
ARGUMENT_RESOLVERS = (
    "TryResolveDate",
    "TryResolveLong",
    "TryResolveBool",
    "TryResolveRounding",
    "TryResolvePillar",
    "TryClassifyShape",
    "TryUnwrapControl",
    "TryMaterialize",
)
WORKBOOK_FALLBACKS = ("ActiveWorkbook", "ThisWorkbook", "ActiveSheet")
HOST_PATH_PROCEDURES = (HOST_CLASSIFIER, HOST_GUARD, HOST_DIAGNOSTIC)
ENGINE_FORBIDDEN_MEMBERS = (
    "UsedRange",
    "Select",
    "Activate",
    "Calculate",
    "EnableEvents",
    "ScreenUpdating",
    "Caller",
    "Run",
)
ENGINE_FORBIDDEN_WORDS = (
    "Selection",
    "AddressOf",
    "CallByName",
    "DateSerial",
    "DateValue",
    "CDate",
    "Weekday",
    "Year",
    "Month",
    "Day",
)


def finding(path: str, message: str, line: int | None = None) -> dict[str, Any]:
    item: dict[str, Any] = {"path": path, "message": message}
    if line is not None:
        item["line"] = line
    return item


def result(rule_id: str, title: str, failures: list[dict[str, Any]], summary: str) -> dict[str, Any]:
    return {
        "id": rule_id,
        "title": title,
        "status": "fail" if failures else "pass",
        "summary": f"{len(failures)} finding{'s' if len(failures) != 1 else ''}" if failures else summary,
        "findings": failures,
    }


def git(root: Path, *args: str) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(
        ["git", "-C", str(root), *args],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )


def read_vba(path: Path) -> str:
    return path.read_bytes().decode("cp1252")


def strip_vba(raw: str) -> str:
    output: list[str] = []
    in_string = False
    index = 0
    while index < len(raw):
        char = raw[index]
        if char == '"':
            if in_string and index + 1 < len(raw) and raw[index + 1] == '"':
                output.extend(('"', '"'))
                index += 2
                continue
            in_string = not in_string
        elif char == "'" and not in_string:
            break
        output.append(char)
        index += 1
    text = "".join(output)
    return "" if re.match(r"^\s*Rem(?:\s|$)", text, re.I) else text.rstrip()


def strip_strings(text: str) -> str:
    return re.sub(r'"(?:[^"]|"")*"', '""', text)


def logical(text: str) -> list[tuple[int, str]]:
    output: list[tuple[int, str]] = []
    buffer: list[str] = []
    start = 0
    for number, raw in enumerate(text.splitlines(), 1):
        code = strip_vba(raw)
        if not buffer:
            start = number
        if re.search(r"\s_\s*$", code):
            buffer.append(re.sub(r"\s_\s*$", " ", code))
            continue
        buffer.append(code)
        statement = " ".join(item.strip() for item in buffer)
        output.append((start, re.sub(r"\s+", " ", statement).strip()))
        buffer.clear()
    if buffer:
        output.append((start, re.sub(r"\s+", " ", " ".join(buffer)).strip()))
    return output


def procedures(text: str) -> list[tuple[str, str, str]]:
    output: list[tuple[str, str, str]] = []
    pattern = re.compile(
        r"^(Public|Private)\s+(?:Function|Sub)\s+(\w+)\b(.*?)^End\s+(?:Function|Sub)\s*$",
        re.I | re.M | re.S,
    )
    for match in pattern.finditer(text):
        raw = match.group(3).splitlines()
        index = 0
        while index < len(raw) and raw[index].rstrip().endswith("_"):
            index += 1
        body = "\n".join(raw[index + 1 :])
        output.append((match.group(1).casefold(), match.group(2), body))
    return output


def load_inputs(root: Path) -> dict[str, Any]:
    config = json.loads((root / CONFIG_PATH).read_text(encoding="utf-8"))
    roles = dict(config["vba"]["components"])
    sources: dict[str, str] = {}
    for path in roles:
        candidate = root / path
        if candidate.is_file() and candidate.suffix.casefold() in {".bas", ".cls", ".frm"}:
            sources[path] = read_vba(candidate)
    return {
        "roles": roles,
        "sources": sources,
        "manifest": (root / MANIFEST_PATH).read_text(encoding="utf-8"),
        "contract": (root / CONTRACT_PATH).read_text(encoding="utf-8"),
    }


def public_names(text: str) -> set[str]:
    names: set[str] = set()
    inside_public_enum = False
    for _, statement in logical(text):
        if re.match(r"^Public\s+Enum\s+\w+", statement, re.I):
            inside_public_enum = True
            match = re.match(r"^Public\s+Enum\s+(\w+)", statement, re.I)
            if match:
                names.add(match.group(1))
            continue
        if re.match(r"^End\s+Enum\b", statement, re.I):
            inside_public_enum = False
            continue
        if inside_public_enum:
            hit = re.match(r"^(KPR_[A-Za-z0-9_]+)\s*=", statement)
            if hit:
                names.add(hit.group(1))
            continue
        hit = re.match(r"^Public\s+(?:Function|Sub|Const)\s+(\w+)", statement, re.I)
        if hit:
            names.add(hit.group(1))
    return names


def rule_components(data: dict[str, Any]) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    roles = data["roles"]
    sources = data["sources"]
    for path, role in EXPECTED_COMPONENTS.items():
        if roles.get(path) != role:
            failures.append(finding(path, f"Expected configured role {role!r}; found {roles.get(path)!r}."))
        if path not in sources:
            failures.append(finding(path, "Required migrated component is not tracked/readable."))
    stale = sorted(path for path in roles if Path(path).stem in {"ProjectCore", "ProjectFacade", "ProjectTests"})
    failures.extend(finding(path, "Neutral starter component remains registered after migration.") for path in stale)
    for path, role in roles.items():
        text = sources.get(path)
        if text is None:
            continue
        private = bool(re.search(r"^\s*Option\s+Private\s+Module\b", text, re.I | re.M))
        if role == "internal" and path.endswith(".bas") and not private:
            failures.append(finding(path, "Internal KPR core module must declare Option Private Module."))
        if role == "public" and path.endswith(".bas") and private:
            failures.append(finding(path, "KPR worksheet facade must not declare Option Private Module."))
    return result(
        "kpr-components",
        "Migrated KPR component roles",
        failures,
        "All six migrated KPR components are registered with the intended roles and visibility",
    )


def manifest_surface(text: str) -> set[str]:
    output: set[str] = set()
    for raw in text.splitlines():
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        fields = raw.split("\t")
        if len(fields) == 3 and fields[0].casefold() == "kpr_dates_days" and fields[1].casefold() == "function":
            output.add(fields[2])
    return output


def contract_surface(text: str) -> set[str]:
    return set(re.findall(r"^Public\s+Function\s+(KPR_Dates_\w+)\b", text, re.I | re.M))


def facade_functions(data: dict[str, Any]) -> dict[str, tuple[int, str]]:
    text = data["sources"].get("src/modules/KPR_DATES_DAYS.bas", "")
    output: dict[str, tuple[int, str]] = {}
    for number, statement in logical(text):
        hit = re.match(r"^Public\s+Function\s+(KPR_Dates_\w+)\b", statement, re.I)
        if hit:
            output[hit.group(1)] = (number, statement)
    return output


def rule_surface(data: dict[str, Any]) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    actual = facade_functions(data)
    actual_names = set(actual)
    manifest = manifest_surface(data["manifest"])
    contract = contract_surface(data["contract"])
    if len(contract) != 22:
        failures.append(finding(CONTRACT_PATH, f"Frozen contract must declare exactly 22 KPR_Dates_* functions; found {len(contract)}."))
    if actual_names != contract:
        missing = sorted(contract - actual_names, key=str.casefold)
        extra = sorted(actual_names - contract, key=str.casefold)
        if missing:
            failures.append(finding("src/modules/KPR_DATES_DAYS.bas", "Missing frozen facade member(s): " + ", ".join(missing) + "."))
        if extra:
            failures.append(finding("src/modules/KPR_DATES_DAYS.bas", "Extra facade member(s) outside the frozen contract: " + ", ".join(extra) + "."))
    if manifest != contract:
        missing = sorted(contract - manifest, key=str.casefold)
        extra = sorted(manifest - contract, key=str.casefold)
        if missing:
            failures.append(finding(MANIFEST_PATH, "Manifest omits contract member(s): " + ", ".join(missing) + "."))
        if extra:
            failures.append(finding(MANIFEST_PATH, "Manifest adds non-contract member(s): " + ", ".join(extra) + "."))
    for name, (line, statement) in actual.items():
        if not re.search(r"\)\s+As\s+Variant\s*$", statement, re.I):
            failures.append(finding("src/modules/KPR_DATES_DAYS.bas", f"{name} must return Variant so native Excel errors remain representable.", line))
        if name.casefold().endswith("_spill"):
            failures.append(finding("src/modules/KPR_DATES_DAYS.bas", f"{name} is a forbidden _Spill twin.", line))
        if name.casefold() == "kpr_dates_datesfrompillar":
            failures.append(finding("src/modules/KPR_DATES_DAYS.bas", "Legacy plural KPR_Dates_DatesFromPillar is forbidden; use KPR_Dates_DateFromPillar.", line))
    return result(
        "kpr-public-surface",
        "Frozen KPR public surface",
        failures,
        "Contract, manifest and facade agree on exactly 22 Variant-returning KPR_Dates_* functions",
    )


def rule_required_members(data: dict[str, Any]) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    by_stem = {Path(path).stem.casefold(): (path, text) for path, text in data["sources"].items()}
    for stem, required in REQUIRED_MEMBERS.items():
        item = by_stem.get(stem)
        if item is None:
            failures.append(finding(stem, "Required architecture component is absent."))
            continue
        path, text = item
        declared = {name.casefold() for name in public_names(text)}
        missing = sorted(name for name in required if name.casefold() not in declared)
        if missing:
            failures.append(finding(path, "Required public in-project member(s) missing: " + ", ".join(missing) + "."))
    return result(
        "kpr-required-members",
        "KPR in-project dependency surface",
        failures,
        "All pinned in-project members required by the migrated architecture are declared",
    )


def executable_text(text: str) -> str:
    return "\n".join(strip_strings(strip_vba(line)) for line in text.splitlines())


def rule_dependencies(data: dict[str, Any]) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    modules = {
        Path(path).stem.casefold(): (path, text)
        for path, text in data["sources"].items()
        if Path(path).stem.casefold() in ALLOWED_DEPENDENCIES
    }
    exported = {stem: public_names(text) for stem, (_, text) in modules.items()}
    for stem, (path, text) in modules.items():
        allowed = ALLOWED_DEPENDENCIES[stem]
        code = executable_text(text)
        for other, names in exported.items():
            if other == stem or other in allowed:
                continue
            hits = sorted(name for name in names if re.search(rf"\b{re.escape(name)}\b", code))
            if hits:
                failures.append(finding(path, f"Module may not depend on {other}; references {', '.join(hits)}."))
    return result(
        "kpr-dependencies",
        "KPR module dependency matrix",
        failures,
        "All migrated KPR architecture modules respect the frozen dependency direction",
    )


def rule_locale(data: dict[str, Any]) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    checked = 0
    for path, text in data["sources"].items():
        if not path.startswith("src/"):
            continue
        checked += 1
        for number, raw in enumerate(text.splitlines(), 1):
            code = strip_strings(strip_vba(raw))
            for name in FORBIDDEN_PARSE_CALLS:
                if re.search(rf"\b{name}\s*\(", code):
                    failures.append(finding(path, f"Locale-sensitive {name} is forbidden in production parsing.", number))
    return result(
        "kpr-locale-parsing",
        "Locale-independent date parsing",
        failures,
        f"None of the {checked} production KPR components use forbidden locale-sensitive parsers",
    )


def _constant(text: str, name: str) -> str | None:
    for _, statement in logical(text):
        if re.search(rf"\bConst\s+{re.escape(name)}\b", statement, re.I):
            return statement
    return None


def rule_window(data: dict[str, Any]) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    dates = data["sources"].get("src/core/KPR_Core_Dates.bas", "")
    parse = data["sources"].get("src/core/KPR_Core_Parse.bas", "")
    expectations = (
        (dates, "src/core/KPR_Core_Dates.bas", "KPR_MIN_DATE", r"=\s*#3/1/1900#"),
        (dates, "src/core/KPR_Core_Dates.bas", "KPR_MAX_DATE", r"=\s*#12/31/9999#"),
        (parse, "src/core/KPR_Core_Parse.bas", "KPR_MIN_SERIAL", r"=\s*61#"),
        (parse, "src/core/KPR_Core_Parse.bas", "KPR_MAX_SERIAL", r"=\s*2958465#"),
        (parse, "src/core/KPR_Core_Parse.bas", "KPR_MIN_YEAR", r"=\s*1900\b"),
        (parse, "src/core/KPR_Core_Parse.bas", "KPR_MAX_YEAR", r"=\s*9999\b"),
    )
    for text, path, name, pattern in expectations:
        statement = _constant(text, name)
        if statement is None:
            failures.append(finding(path, f"{name} is not declared."))
        elif not re.search(pattern, statement, re.I):
            failures.append(finding(path, f"{name} no longer matches the frozen date-window value: {statement}"))
    return result(
        "kpr-date-window",
        "Frozen supported date window",
        failures,
        "Date and parser bounds remain pinned to 1900-03-01 through 9999-12-31",
    )


def facade_procedures(data: dict[str, Any]) -> list[tuple[str, str, str]]:
    return procedures(data["sources"].get("src/modules/KPR_DATES_DAYS.bas", ""))


def rule_host_guard(data: dict[str, Any]) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    checked = 0
    guard_call = re.compile(rf"(?<![\w.])(?<!Function\s){HOST_GUARD}\s*\(", re.I)
    for visibility, name, body in facade_procedures(data):
        code = "\n".join(strip_vba(line) for line in body.splitlines())
        if visibility == "private":
            calls = [
                line for line in code.splitlines()
                if guard_call.search(strip_strings(line))
                and not re.match(rf"^\s*{HOST_GUARD}\s*=", line, re.I)
            ]
            if calls:
                failures.append(finding("src/modules/KPR_DATES_DAYS.bas", f"{name} calls {HOST_GUARD}; only the public boundary may run the host guard."))
            continue
        if not name.casefold().startswith("kpr_dates_"):
            continue
        checked += 1
        if name.casefold() == HOST_DIAGNOSTIC.casefold():
            classifier_calls = len(re.findall(rf"\b{HOST_CLASSIFIER}\s*\(", code, re.I))
            if classifier_calls != 1:
                failures.append(finding("src/modules/KPR_DATES_DAYS.bas", f"{name} must call {HOST_CLASSIFIER} exactly once; found {classifier_calls}."))
            if re.search(rf"\b{HOST_GUARD}\s*\(", code, re.I):
                failures.append(finding("src/modules/KPR_DATES_DAYS.bas", f"{name} must report the date system, not call {HOST_GUARD}."))
            continue
        guards = [m.start() for m in re.finditer(rf"\b{HOST_GUARD}\s*\(", code, re.I)]
        if len(guards) != 1:
            failures.append(finding("src/modules/KPR_DATES_DAYS.bas", f"{name} must call {HOST_GUARD} exactly once; found {len(guards)}."))
            continue
        first_resolver = min(
            (m.start() for resolver in ARGUMENT_RESOLVERS for m in re.finditer(rf"\b{resolver}\s*\(", code, re.I)),
            default=None,
        )
        if first_resolver is not None and first_resolver < guards[0]:
            failures.append(finding("src/modules/KPR_DATES_DAYS.bas", f"{name} resolves an argument before {HOST_GUARD}."))
    return result(
        "kpr-host-guard",
        "Caller date-system guard",
        failures,
        f"All {checked} public date functions apply the frozen host guard policy",
    )


def rule_volatile(data: dict[str, Any]) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    hits_total = 0
    for path, text in data["sources"].items():
        if not path.startswith("src/"):
            continue
        for _, name, body in procedures(text):
            lines = [
                line for line in body.splitlines()
                if line.strip()
                and not line.lstrip().startswith("'")
                and not re.match(r"^\s*(Dim|Const|Static)\b", line, re.I)
            ]
            hits = [line for line in lines if VOLATILE_CALL.casefold() in strip_strings(strip_vba(line)).casefold()]
            hits_total += len(hits)
            if name.casefold() == HOST_DIAGNOSTIC.casefold():
                if len(hits) != 1:
                    failures.append(finding(path, f"{HOST_DIAGNOSTIC} must call {VOLATILE_CALL} exactly once; found {len(hits)}."))
                elif not lines or VOLATILE_CALL.casefold() not in strip_strings(strip_vba(lines[0])).casefold():
                    failures.append(finding(path, f"{VOLATILE_CALL} must be the first executable statement of {HOST_DIAGNOSTIC}."))
            elif hits:
                failures.append(finding(path, f"{name} must not be volatile; only {HOST_DIAGNOSTIC} may call {VOLATILE_CALL}."))
    if hits_total != 1:
        failures.append(finding("src/modules/KPR_DATES_DAYS.bas", f"{VOLATILE_CALL} must occur exactly once in production; found {hits_total}."))
    return result(
        "kpr-volatility",
        "Volatility scope",
        failures,
        f"{VOLATILE_CALL} occurs only at the start of {HOST_DIAGNOSTIC}",
    )


def rule_host_authority(data: dict[str, Any]) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    facade = data["sources"].get("src/modules/KPR_DATES_DAYS.bas", "")
    readers: list[str] = []
    for _, name, body in procedures(facade):
        if name in HOST_PATH_PROCEDURES:
            code = "\n".join(strip_strings(strip_vba(line)) for line in body.splitlines())
            for token in WORKBOOK_FALLBACKS:
                if re.search(rf"\b{token}\b", code, re.I):
                    failures.append(finding("src/modules/KPR_DATES_DAYS.bas", f"{name} references {token}; only the caller workbook may supply worksheet date-system authority."))
        code = "\n".join(strip_strings(strip_vba(line)) for line in body.splitlines())
        if re.search(r"\bApplication\s*\.\s*Caller\b", code, re.I):
            readers.append(name)
    wrong = [name for name in readers if name.casefold() != HOST_CLASSIFIER.casefold()]
    for name in wrong:
        failures.append(finding("src/modules/KPR_DATES_DAYS.bas", f"{name} reads Application.Caller; only {HOST_CLASSIFIER} may classify the host."))
    if sum(name.casefold() == HOST_CLASSIFIER.casefold() for name in readers) != 1:
        failures.append(finding("src/modules/KPR_DATES_DAYS.bas", f"{HOST_CLASSIFIER} must be the sole Application.Caller reader."))
    return result(
        "kpr-host-authority",
        "Caller workbook authority",
        failures,
        f"Application.Caller is owned by {HOST_CLASSIFIER} with no active-workbook fallback",
    )


def rule_array_purity(data: dict[str, Any]) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    path = "src/core/KPR_Core_Array.bas"
    text = data["sources"].get(path, "")
    member = re.compile(r"\.\s*(" + "|".join(ENGINE_FORBIDDEN_MEMBERS) + r")\b", re.I)
    word = re.compile(r"(?<![\w.])(" + "|".join(ENGINE_FORBIDDEN_WORDS) + r")\b", re.I)
    for number, raw in enumerate(text.splitlines(), 1):
        code = strip_strings(strip_vba(raw))
        for pattern in (member, word):
            for hit in pattern.finditer(code):
                failures.append(finding(path, f"Array engine uses {hit.group(1)}; it owns shape only, not host state, dispatch or calendar logic.", number))
    return result(
        "kpr-array-purity",
        "Array-engine purity",
        failures,
        "KPR_Core_Array contains no forbidden Excel-state, dispatch, host or date tokens",
    )


def rule_day_zero(data: dict[str, Any]) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    pattern = re.compile(r"\bDateSerial\s*\([^)]*,[^)]*,\s*0\s*\)", re.I)
    checked = 0
    for path, text in data["sources"].items():
        if not path.startswith("src/"):
            continue
        checked += 1
        for number, statement in logical(text):
            if pattern.search(strip_strings(statement)):
                failures.append(finding(path, "DateSerial(..., ..., 0) is forbidden at the upper supported boundary; use the bounded date core.", number))
    return result(
        "kpr-day-zero",
        "Boundary-safe date construction",
        failures,
        f"No day-zero DateSerial idiom appears in {checked} production components",
    )


ORACLE_MODULE = "kpr_test_oracle"
ORACLE_FUNCTIONS = frozenset({"EOMONTH", "EDATE", "WEEKDAY", "DAY", "YEAR", "MONTH"})
ORACLE_CALL = re.compile(r"\bXl\s*\(", re.I)
ORACLE_PROCEDURE = re.compile(
    r"^(?:(?:Public|Private|Friend|Static)\s+)*(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+(\w+)", re.I
)
# Any way to hand a formula to Excel (evaluation, worksheet functions, cell or
# name formulas, recalculation, macro calls); only the Xl helper may use one.
ORACLE_EVALUATION = re.compile(
    r"\b(?:Evaluate|ExecuteExcel4Macro|WorksheetFunction|CallByName|Range|Cells|Names|SendKeys|DDE\w*)\b"
    r"|\.(?:Formula\w*|Value2?|RefersTo\w*|Calculate\w*|Run)\b|\[",
    re.I,
)
ORACLE_DECLARATION = re.compile(r"(\w+)(?:\s*\([^)]*\))?\s+As\s+(\w+)", re.I)
NUMERIC_TYPES = frozenset({"byte", "integer", "long", "longlong", "single", "double", "currency"})


def _split_top(text: str, separator: str) -> list[str]:
    """Split VBA text on a separator outside string literals and parentheses."""
    parts, current, depth, in_string = [], "", 0, False
    for char in text:
        if char == '"':
            in_string = not in_string
        elif not in_string and char in "()":
            depth += 1 if char == "(" else -1
        if char == separator and not in_string and depth == 0:
            parts.append(current.strip())
            current = ""
        else:
            current += char
    parts.append(current.strip())
    return parts


def oracle_calls(statement: str) -> list[str]:
    """The argument text of each Xl(...) call: the formula Excel evaluates."""
    calls: list[str] = []
    match = ORACLE_CALL.search(statement)
    while match:
        index, depth, in_string = match.end(), 1, False
        while index < len(statement) and depth:
            char = statement[index]
            if char == '"':
                in_string = not in_string
            elif not in_string and char in "()":
                depth += 1 if char == "(" else -1
            index += 1
        calls.append(statement[match.end():index - 1])
        match = ORACLE_CALL.search(statement, index)
    return calls


def oracle_expressions(statement: str) -> list[str]:
    """String literals inside each Xl(...) call: the formulas Excel evaluates."""
    return [
        " ".join(item.replace('""', '"') for item in re.findall(r'"((?:[^"]|"")*)"', call))
        for call in oracle_calls(statement)
    ]


def _numeric_text(expression: str, numeric_names: set[str]) -> bool:
    """True when expression is arithmetic over numeric literals and numeric-typed variables only."""
    if not re.fullmatch(r"[\w\s+\-*/\\().]*", expression) or re.search(r"\w\s*\(", expression):
        return False
    names = {name.casefold() for name in re.findall(r"[A-Za-z_]\w*", expression)}
    return names <= numeric_names | {"mod"}


def _opaque_formula(argument: str, numeric_names: set[str]) -> bool:
    """True when a formula part is neither a literal nor CStr of a numeric expression."""
    for part in _split_top(argument, "&"):
        inner = re.fullmatch(r"CStr\s*\((.*)\)", part, re.I | re.S)
        if re.fullmatch(r'"(?:[^"]|"")*"', part):
            continue
        if not (inner and _numeric_text(inner.group(1), numeric_names)):
            return True
    return False


def _oracle_statement_errors(statement: str, numeric_names: set[str], used: set[str]) -> list[str]:
    errors: list[str] = []
    for literal in re.findall(r'"((?:[^"]|"")*)"', statement):
        if re.search(r"WORKDAY|NETWORKDAYS", literal, re.I):
            errors.append("WORKDAY.INTL and NETWORKDAYS.INTL are out of v0.0.4 scope.")
    for argument in oracle_calls(statement):
        if _opaque_formula(argument, numeric_names):
            errors.append("Every Xl formula must be built from literals and CStr of numeric expressions so it can be inspected.")
    for expression in oracle_expressions(statement):
        for name in re.findall(r"([A-Za-z_][A-Za-z0-9_.]*)\s*\(", expression):
            used.add(name.upper())
            if name.upper() not in ORACLE_FUNCTIONS:
                errors.append(f"Oracle formula calls {name}; only {', '.join(sorted(ORACLE_FUNCTIONS))} are permitted.")
    return errors


def rule_oracle_scope(data: dict[str, Any]) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    used: set[str] = set()
    modules = [(p, t) for p, t in data["sources"].items() if Path(p).stem.casefold() == ORACLE_MODULE]
    if not modules:
        failures.append(finding(CONFIG_PATH, "The Excel cross-oracle module KPR_Test_Oracle is not registered."))
    for path, text in modules:
        statements = logical(text)
        declared = [
            (name.casefold(), kind.casefold())
            for _, statement in statements
            for name, kind in ORACLE_DECLARATION.findall(statement)
        ]
        numeric_names = {n for n, k in declared if k in NUMERIC_TYPES} - {n for n, k in declared if k not in NUMERIC_TYPES}
        procedure = ""
        for number, statement in statements:
            header = ORACLE_PROCEDURE.match(statement)
            if header:
                procedure = header.group(1)
            if procedure.casefold() != "xl" and ORACLE_EVALUATION.search(strip_strings(statement)):
                failures.append(finding(path, "Oracle formulas must reach Excel only through the Xl helper.", number))
            if procedure.casefold() == "xl":
                continue
            failures.extend(finding(path, error, number) for error in _oracle_statement_errors(statement, numeric_names, used))
    return result(
        "kpr-oracle-scope",
        "Excel cross-oracle function scope",
        failures,
        f"The cross-oracle evaluates only {', '.join(sorted(used)) or 'no functions'}",
    )


RULES = (
    rule_components,
    rule_surface,
    rule_required_members,
    rule_dependencies,
    rule_locale,
    rule_window,
    rule_host_guard,
    rule_volatile,
    rule_host_authority,
    rule_array_purity,
    rule_day_zero,
    rule_oracle_scope,
)


def evaluate(data: dict[str, Any]) -> list[dict[str, Any]]:
    return [rule(data) for rule in RULES]


def report(root: Path, data: dict[str, Any] | None = None) -> dict[str, Any]:
    data = load_inputs(root) if data is None else data
    rules = evaluate(data)
    completed = git(root, "rev-parse", "HEAD")
    commit = completed.stdout.decode("utf-8", errors="replace").strip() if completed.returncode == 0 else None
    return {
        "schema_version": SCHEMA_VERSION,
        "tool": TOOL_NAME,
        "commit": commit,
        "scope_note": (
            "Project-specific static contract evidence only; this does not compile VBA, "
            "execute Excel, or replace destination parity validation in issue #17."
        ),
        "status": "fail" if any(item["status"] == "fail" for item in rules) else "pass",
        "counts": {
            "rules": len(rules),
            "passed": sum(item["status"] == "pass" for item in rules),
            "failed": sum(item["status"] == "fail" for item in rules),
            "findings": sum(len(item["findings"]) for item in rules),
        },
        "rules": rules,
    }


def markdown(rep: dict[str, Any]) -> str:
    lines = [
        "# KPR migrated date-layer contract",
        "",
        f"**Status:** {str(rep['status']).upper()}",
        "",
        str(rep["scope_note"]),
        "",
        "| Rule | Status | Summary |",
        "| --- | --- | --- |",
    ]
    for item in rep["rules"]:
        lines.append(f"| {item['id']} | {str(item['status']).upper()} | {item['summary']} |")
    failed = [item for item in rep["rules"] if item["status"] == "fail"]
    if failed:
        lines += ["", "## Findings", ""]
        for item in failed:
            lines += [f"### {item['title']}", ""]
            for hit in item["findings"]:
                location = str(hit["path"])
                if "line" in hit:
                    location += f":{hit['line']}"
                lines.append(f"- `{location}` — {hit['message']}")
            lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def mutate(data: dict[str, Any], path: str, old: str, new: str) -> dict[str, Any]:
    case = copy.deepcopy(data)
    if old not in case["sources"][path]:
        raise RuntimeError(f"Self-test mutation target not found in {path}: {old!r}")
    case["sources"][path] = case["sources"][path].replace(old, new, 1)
    return case


def self_test(root: Path) -> None:
    base = load_inputs(root)
    baseline = report(root, base)
    if baseline["status"] != "pass":
        failed_summary = ", ".join(
            item["id"] for item in baseline["rules"] if item["status"] == "fail"
        )
        raise RuntimeError(
            f"Positive KPR contract fixture is not green: {failed_summary}"
        )

    facade = "src/modules/KPR_DATES_DAYS.bas"
    parse = "src/core/KPR_Core_Parse.bas"
    dates = "src/core/KPR_Core_Dates.bas"
    err = "src/core/KPR_Core_Err.bas"
    array = "src/core/KPR_Core_Array.bas"

    scenarios: list[tuple[str, str, dict[str, Any]]] = []
    scenarios.append(
        (
            "missing API member",
            "kpr-public-surface",
            mutate(
                base,
                facade,
                "Public Function KPR_Dates_AddDays(",
                "Public Function KPR_Dates_AddDays_Renamed(",
            ),
        )
    )
    spill = copy.deepcopy(base)
    spill["sources"][facade] += "\r\nPublic Function KPR_Dates_AddDays_Spill(ByVal DateIn As Variant) As Variant\r\nEnd Function\r\n"
    scenarios.append(("_Spill twin", "kpr-public-surface", spill))
    scenarios.append(
        (
            "legacy plural pillar",
            "kpr-public-surface",
            mutate(
                base,
                facade,
                "Public Function KPR_Dates_DateFromPillar(",
                "Public Function KPR_Dates_DatesFromPillar(",
            ),
        )
    )
    scenarios.append(
        (
            "narrow return type",
            "kpr-public-surface",
            mutate(
                base,
                facade,
                "ByVal nDays As Variant) _\r\n    As Variant",
                "ByVal nDays As Variant) _\r\n    As Long",
            ),
        )
    )
    locale = copy.deepcopy(base)
    locale["sources"][parse] += "\r\nPublic Function ProbeLocale(ByVal S As String) As Boolean\r\n    ProbeLocale = IsDate(S)\r\nEnd Function\r\n"
    scenarios.append(("locale parser", "kpr-locale-parsing", locale))
    scenarios.append(("window drift", "kpr-date-window", mutate(base, parse, "As Double = 61#", "As Double = 60#")))
    scenarios.append(("missing host guard", "kpr-host-guard", mutate(base, facade, "If Not PassHostGuard(FailErr) Then", "If False Then")))
    volatile = copy.deepcopy(base)
    volatile["sources"][facade] += "\r\nPrivate Sub ProbeVolatile()\r\n    Application.Volatile True\r\nEnd Sub\r\n"
    scenarios.append(("stray volatility", "kpr-volatility", volatile))
    scenarios.append(("active workbook fallback", "kpr-host-authority", mutate(base, facade, "Set CallerObject = Application.Caller", "Set CallerObject = ActiveWorkbook")))
    engine = copy.deepcopy(base)
    engine["sources"][array] += "\r\nPublic Function ProbeDate(ByVal V As Variant) As Variant\r\n    ProbeDate = Year(V)\r\nEnd Function\r\n"
    scenarios.append(("array date math", "kpr-array-purity", engine))
    scenarios.append(
        (
            "missing private visibility",
            "kpr-components",
            mutate(
                base,
                parse,
                "    Option Private Module   'Internal module: invisible outside this VBA project",
                "    ' Option Private Module removed by self-test",
            ),
        )
    )
    scenarios.append(("missing required member", "kpr-required-members", mutate(base, dates, "Public Function DaysInMonth", "Public Function DaysInMonth_Missing")))
    dependency = copy.deepcopy(base)
    dependency["sources"][err] += "\r\nPublic Function ProbeDependency() As Boolean\r\n    ProbeDependency = TryParseDateScalar(0, 0, 0)\r\nEnd Function\r\n"
    scenarios.append(("forbidden reverse dependency", "kpr-dependencies", dependency))
    boundary = copy.deepcopy(base)
    boundary["sources"][dates] += "\r\nPublic Function ProbeBoundary(ByVal Y As Long, ByVal M As Long) As Date\r\n    ProbeBoundary = DateSerial(Y, M + 1, 0)\r\nEnd Function\r\n"
    scenarios.append(("day-zero DateSerial", "kpr-day-zero", boundary))

    oracle = next(path for path in base["sources"] if Path(path).stem.casefold() == ORACLE_MODULE)
    scenarios.append((
        "business-day oracle",
        "kpr-oracle-scope",
        mutate(base, oracle, 'Xl("WEEKDAY(" & CStr(Serial) & ",2)")', 'Xl("WORKDAY.INTL(" & CStr(Serial) & ",2)")'),
    ))
    scenarios.append((
        "unlisted oracle function",
        "kpr-oracle-scope",
        mutate(base, oracle, 'Xl("DAY(" & CStr(Serial) & ")")', 'Xl("DATEVALUE(" & CStr(Serial) & ")")'),
    ))
    scenarios.append((
        "lower-case oracle call",
        "kpr-oracle-scope",
        mutate(base, oracle, 'Xl("DAY(" & CStr(Serial) & ")")', 'xl("DATEVALUE(" & CStr(Serial) & ")")'),
    ))
    scenarios.append((
        "evaluation outside Xl",
        "kpr-oracle-scope",
        mutate(base, oracle, 'Xl("DAY(" & CStr(Serial) & ")")', 'mSheet.Evaluate("DATEVALUE(" & CStr(Serial) & ")")'),
    ))
    scenarios.append((
        "formula held in a variable",
        "kpr-oracle-scope",
        mutate(base, oracle, 'Xl("DAY(" & CStr(Serial) & ")")', 'Xl(Tag)'),
    ))
    scenarios.append((
        "text smuggled through CStr",
        "kpr-oracle-scope",
        mutate(base, oracle, 'Xl("DAY(" & CStr(Serial) & ")")', 'Xl("DAY(" & CStr(Tag) & ")")'),
    ))
    scenarios.append((
        "text assigned inside If",
        "kpr-oracle-scope",
        mutate(
            mutate(base, oracle, "Tag = IsoText(Serial)", 'If Serial > 0 Then Tag = "DATEVALUE(1)"'),
            oracle, 'Xl("DAY(" & CStr(Serial) & ")")', "Xl(Tag)",
        ),
    ))
    scenarios.append((
        "text built from character codes",
        "kpr-oracle-scope",
        mutate(base, oracle, 'Xl("DAY(" & CStr(Serial) & ")")', 'Xl(CStr(Chr$(68) & Chr$(65)))'),
    ))
    scenarios.append((
        "bracket evaluation of a name",
        "kpr-oracle-scope",
        mutate(base, oracle, "D = CDate(Serial)", "D = CDate(Serial): Tag = [HiddenFormula]"),
    ))
    for label, probe in (
        ("worksheet function outside Xl", "Nth = Application.WorksheetFunction.NetworkDays_Intl(1, 2)"),
        ("cell formula outside Xl", 'mSheet.Cells(1, 1).Formula = "=DATEVALUE(1)": mSheet.Calculate'),
        ("cell value outside Xl", 'mSheet.Cells(1, 1).Value = "=DATEVALUE(1)"'),
        ("workbook name outside Xl", 'mSheet.Parent.Names.Add "Hidden", "=DATEVALUE(1)"'),
    ):
        scenarios.append((label, "kpr-oracle-scope", mutate(base, oracle, "D = CDate(Serial)", f"D = CDate(Serial): {probe}")))
    for name, expected, case in scenarios:
        rep = report(root, case)
        failed_ids = {
            item["id"] for item in rep["rules"] if item["status"] == "fail"
        }
        if expected not in failed_ids:
            raise RuntimeError(
                f"Degraded self-test {name!r} did not fail {expected}; "
                f"failed={sorted(failed_ids)}"
            )
        print(f"PASS degraded self-test: {name} -> {expected}")
    print(f"PASS positive self-test: {len(RULES)} rules; {len(scenarios)} degraded cases")


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path, default=Path("test-results/kpr-contract.json"))
    parser.add_argument("--summary", type=Path)
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    root = args.root.resolve()
    try:
        if args.self_test:
            self_test(root)
            return 0
        rep = report(root)
        output = args.output if args.output.is_absolute() else root / args.output
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(rep, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        summary = markdown(rep)
        if args.summary is not None:
            destination = args.summary if args.summary.is_absolute() else root / args.summary
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_text(summary, encoding="utf-8")
        print(summary, end="")
        print(f"JSON report: {output}")
        return 0 if rep["status"] == "pass" else 1
    except (OSError, KeyError, ValueError, RuntimeError, subprocess.SubprocessError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
