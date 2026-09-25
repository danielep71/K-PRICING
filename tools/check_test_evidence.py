#!/usr/bin/env python3
"""Validate K-PRICING structured test evidence without executing Excel.

The durable VBA runner (KPR_Test_RunAll / KPR_Test_RunSuite in
tests/modules/KPR_REGRESSION_TESTS.bas) writes kpr-test-evidence.json. This
tool checks a record against the committed schema
(docs/kpr-test-evidence.schema.json), then applies the semantic rules the
schema cannot express: the registry read from VBA source, fixture binding,
count consistency, outcome discipline, determinism and certification
completeness. It uses the standard library only.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import sys
import tempfile
from pathlib import Path
from typing import Any

SCHEMA_PATH = "docs/kpr-test-evidence.schema.json"
RUNNER_PATH = "tests/modules/KPR_REGRESSION_TESTS.bas"
FIXTURE_PATH = "tests/fixtures/date_layer_fixtures.tsv"
EVIDENCE_NAME = "kpr-test-evidence.json"

CERTIFICATION_KEYS = (
    "source_import",
    "vba_compile",
    "regression",
    "cross_oracle",
    "macro_options",
    "ribbonx",
    "commandbars",
    "demo_generation",
    "source_round_trip",
)
OUTCOME_STATUSES = ("PASS", "FAIL", "NOT_RUN", "NOT_APPLICABLE")
# Outcomes a certification record may never mark NOT_APPLICABLE.
ALWAYS_APPLICABLE = ("source_import", "vba_compile", "regression", "cross_oracle")
NONDETERMINISTIC_FIELDS = ("/environment", "/timing", "/suites/*/elapsed_ms")
RUNNER_FAILURE_SUITE = "runner"
REGISTRY_PATTERN = re.compile(
    r"Private Function TestRegistry\(\)[^\n]*\n(?P<body>.*?)\nEnd Function",
    re.DOTALL,
)
REGISTRY_ASSIGNMENT = re.compile(r"TestRegistry\s*=\s*Array\((?P<items>.*?)\)\s*$", re.DOTALL)


class EvidenceError(Exception):
    """A usage or input problem that prevents validation (exit 2)."""


# ---------------------------------------------------------------------------
# Repository inputs
# ---------------------------------------------------------------------------


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise EvidenceError(f"{path} does not exist") from exc
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise EvidenceError(f"{path} is not valid UTF-8 JSON: {exc}") from exc


def read_registry(root: Path) -> list[str]:
    """Return the ordered durable-runner registry declared in VBA source."""
    path = root / RUNNER_PATH
    try:
        text = path.read_bytes().decode("cp1252").replace("\r\n", "\n")
    except FileNotFoundError as exc:
        raise EvidenceError(f"{path} does not exist") from exc
    match = REGISTRY_PATTERN.search(text)
    if not match:
        raise EvidenceError(f"{RUNNER_PATH} declares no TestRegistry function")
    code = "\n".join(
        line for line in match.group("body").splitlines() if not line.lstrip().startswith("'")
    )
    code = code.replace(" _\n", " ")
    assignment = REGISTRY_ASSIGNMENT.search(code.strip())
    if not assignment:
        raise EvidenceError(f"{RUNNER_PATH} TestRegistry is not a single Array(...) assignment")
    names = re.findall(r'"([^"]*)"', assignment.group("items"))
    if not names or len(names) != len(set(names)):
        raise EvidenceError(f"{RUNNER_PATH} TestRegistry must list unique suite names")
    return names


def fixture_identity(root: Path) -> tuple[str, int]:
    """Return the SHA-256 and case count of the canonical fixture TSV."""
    path = root / FIXTURE_PATH
    try:
        data = path.read_bytes()
    except FileNotFoundError as exc:
        raise EvidenceError(f"{path} does not exist") from exc
    lines = [line for line in data.decode("utf-8").split("\n") if line]
    return hashlib.sha256(data).hexdigest(), max(len(lines) - 1, 0)


def suite_kind(name: str) -> str:
    if name == "fixtures":
        return "fixture"
    if name.startswith("worksheet-"):
        return "worksheet"
    return "pure"


# ---------------------------------------------------------------------------
# Schema subset validator
# ---------------------------------------------------------------------------

JSON_TYPES: dict[str, tuple[type, ...]] = {
    "object": (dict,),
    "array": (list,),
    "string": (str,),
    "integer": (int,),
    "number": (int, float),
    "boolean": (bool,),
    "null": (type(None),),
}
SUPPORTED_KEYWORDS = {
    "$schema", "$id", "$defs", "$ref", "title", "description", "type", "required",
    "properties", "additionalProperties", "enum", "const", "pattern", "minLength",
    "minimum", "items", "minItems",
}


def _is_type(value: Any, name: str) -> bool:
    if isinstance(value, bool) and name in {"integer", "number"}:
        return False
    return isinstance(value, JSON_TYPES[name])


def _resolve(schema: dict[str, Any], node: dict[str, Any]) -> dict[str, Any]:
    while "$ref" in node:
        ref = node["$ref"]
        if not isinstance(ref, str) or not ref.startswith("#/$defs/"):
            raise EvidenceError(f"unsupported schema reference {ref!r}")
        node = schema["$defs"][ref.removeprefix("#/$defs/")]
    return node


def check_schema_supported(node: Any, path: str = "#") -> None:
    """Refuse schema keywords this subset validator would silently ignore."""
    if isinstance(node, dict):
        unknown = set(node) - SUPPORTED_KEYWORDS
        # Keys under properties/$defs are names, not keywords.
        if unknown and not path.endswith(("/properties", "/$defs")):
            raise EvidenceError(f"schema uses unsupported keywords at {path}: {sorted(unknown)}")
        for key, value in node.items():
            check_schema_supported(value, f"{path}/{key}")
    elif isinstance(node, list):
        for index, value in enumerate(node):
            check_schema_supported(value, f"{path}/{index}")


def schema_errors(schema: dict[str, Any], value: Any, node: dict[str, Any], at: str) -> list[str]:
    node = _resolve(schema, node)
    if "const" in node and (value != node["const"] or type(value) is not type(node["const"])):
        return [f"{at}: expected {node['const']!r}"]
    if "enum" in node and not any(
        value == item and type(value) is type(item) for item in node["enum"]
    ):
        return [f"{at}: {value!r} is not one of {node['enum']}"]
    if "type" in node:
        names = node["type"] if isinstance(node["type"], list) else [node["type"]]
        if not any(_is_type(value, name) for name in names):
            return [f"{at}: expected type {'/'.join(names)}"]
    if isinstance(value, list):
        return _array_errors(schema, value, node, at)
    if isinstance(value, dict):
        return _object_errors(schema, value, node, at)
    return _scalar_errors(value, node, at)


def _scalar_errors(value: Any, node: dict[str, Any], at: str) -> list[str]:
    errors: list[str] = []
    if isinstance(value, str):
        if len(value) < node.get("minLength", 0):
            errors.append(f"{at}: shorter than {node['minLength']} character(s)")
        if "pattern" in node and not re.search(node["pattern"], value):
            errors.append(f"{at}: does not match {node['pattern']}")
    if _is_type(value, "integer") and "minimum" in node and value < node["minimum"]:
        errors.append(f"{at}: below minimum {node['minimum']}")
    return errors


def _array_errors(schema: dict[str, Any], value: list[Any], node: dict[str, Any], at: str) -> list[str]:
    errors: list[str] = []
    if len(value) < node.get("minItems", 0):
        errors.append(f"{at}: fewer than {node['minItems']} item(s)")
    if "items" in node:
        for index, item in enumerate(value):
            errors.extend(schema_errors(schema, item, node["items"], f"{at}/{index}"))
    return errors


def _object_errors(
    schema: dict[str, Any], value: dict[str, Any], node: dict[str, Any], at: str
) -> list[str]:
    errors: list[str] = []
    properties = node.get("properties", {})
    for key in node.get("required", []):
        if key not in value:
            errors.append(f"{at}: missing required property {key!r}")
    if node.get("additionalProperties") is False:
        for key in value:
            if key not in properties:
                errors.append(f"{at}: unexpected property {key!r}")
    for key, child in properties.items():
        if key in value:
            errors.extend(schema_errors(schema, value[key], child, f"{at}/{key}"))
    return errors


# ---------------------------------------------------------------------------
# Semantic rules
# ---------------------------------------------------------------------------


def _outcome_errors(key: str, outcome: dict[str, Any]) -> list[str]:
    status, detail, reason = outcome["status"], outcome["detail"], outcome["reason"]
    if status in {"PASS", "FAIL"}:
        if not (isinstance(detail, str) and detail.strip()) or reason is not None:
            return [f"certification.{key}: {status} needs a nonempty detail and a null reason"]
    elif not (isinstance(reason, str) and reason.strip()) or detail is not None:
        return [f"certification.{key}: {status} needs a nonempty reason and a null detail"]
    return []


def semantic_errors(
    record: dict[str, Any],
    registry: list[str],
    fixture_sha: str,
    fixture_cases: int,
    candidate_sha: str | None,
) -> list[str]:
    errors = _identity_errors(record, fixture_sha, fixture_cases, candidate_sha)
    errors.extend(_selection_errors(record, registry))
    errors.extend(_suite_errors(record))
    errors.extend(_total_errors(record))
    errors.extend(_outcome_rules(record))
    return errors


def _identity_errors(
    record: dict[str, Any], fixture_sha: str, fixture_cases: int, candidate_sha: str | None
) -> list[str]:
    errors: list[str] = []
    runner = record["runner"]
    if candidate_sha is not None and record["source_sha"] != candidate_sha:
        errors.append(f"source_sha {record['source_sha']} is not the candidate {candidate_sha}")
    if runner["fixture_source_sha256"] != fixture_sha:
        errors.append("runner.fixture_source_sha256 does not match " + FIXTURE_PATH)
    if runner["fixture_cases"] != fixture_cases:
        errors.append(f"runner.fixture_cases is not the {fixture_cases} cases in {FIXTURE_PATH}")
    if tuple(record["nondeterministic_fields"]) != NONDETERMINISTIC_FIELDS:
        errors.append(f"nondeterministic_fields must be {list(NONDETERMINISTIC_FIELDS)}")
    timing = record["timing"]
    if timing["finished_local"] < timing["started_local"]:
        errors.append("timing.finished_local precedes timing.started_local")
    return errors


def _selection_errors(record: dict[str, Any], registry: list[str]) -> list[str]:
    errors: list[str] = []
    runner = record["runner"]
    names = [suite["name"] for suite in record["suites"]]
    selection = runner["selection"]
    if selection == "all":
        if runner["entry_point"] != "KPR_Test_RunAll":
            errors.append("selection 'all' must come from KPR_Test_RunAll")
        if names != registry:
            errors.append("suites must be the complete TestRegistry in order")
        return errors
    if runner["entry_point"] != "KPR_Test_RunSuite":
        errors.append("a single-suite selection must come from KPR_Test_RunSuite")
    if selection not in registry:
        errors.append(f"selection {selection!r} is not in TestRegistry")
    if names != [selection]:
        errors.append("suites must hold exactly the selected suite")
    return errors


def _suite_errors(record: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    listed: dict[str, int] = {}
    for failure in record["failures"]:
        listed[failure["suite"]] = listed.get(failure["suite"], 0) + 1
    for suite in record["suites"]:
        name, status, count = suite["name"], suite["status"], suite["failures"]
        if suite["kind"] != suite_kind(name):
            errors.append(f"suite {name}: kind must be {suite_kind(name)}")
        if listed.get(name, 0) != count:
            errors.append(f"suite {name}: failures {count} but {listed.get(name, 0)} listed")
        if status == "PASS" and count:
            errors.append(f"suite {name}: PASS with {count} failure(s)")
        if status == "FAIL" and not count:
            errors.append(f"suite {name}: FAIL without a listed failure")
        if status == "NOT_RUN" and (count or suite["assertions"]):
            errors.append(f"suite {name}: NOT_RUN with recorded assertions or failures")
    names = {suite["name"] for suite in record["suites"]}
    unknown = set(listed) - names - {RUNNER_FAILURE_SUITE}
    if unknown:
        errors.append(f"failures name suites that did not run: {sorted(unknown)}")
    return errors


def _total_errors(record: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    totals = record["totals"]
    if totals["suites"] != len(record["suites"]):
        errors.append("totals.suites does not match the suites list")
    if totals["assertions"] != sum(suite["assertions"] for suite in record["suites"]):
        errors.append("totals.assertions does not match the suite assertions")
    if totals["failures"] != len(record["failures"]):
        errors.append("totals.failures does not match the failures list")
    return errors


def _outcome_rules(record: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    passed = (
        all(suite["status"] == "PASS" for suite in record["suites"])
        and not record["failures"]
        and record["state_restoration"]["status"] == "PASS"
    )
    expected_result = "PASS" if passed else "FAIL"
    if record["result"] != expected_result:
        errors.append(f"result must be {expected_result} for these suites, failures and state")
    certification = record["certification"]
    for key in CERTIFICATION_KEYS:
        errors.extend(_outcome_errors(key, certification[key]))
    if certification["regression"]["status"] != record["result"]:
        errors.append("certification.regression.status must equal result")
    return errors


def certification_errors(record: dict[str, Any]) -> list[str]:
    """Rules for a record presented as the complete certification record."""
    errors: list[str] = []
    if record["runner"]["selection"] != "all":
        errors.append("certification needs a KPR_Test_RunAll record")
    for key in CERTIFICATION_KEYS:
        status = record["certification"][key]["status"]
        allowed = ("PASS",) if key in ALWAYS_APPLICABLE else ("PASS", "NOT_APPLICABLE")
        if status not in allowed:
            errors.append(f"certification.{key} is {status}; certification needs {' or '.join(allowed)}")
    return errors


def deterministic_view(record: dict[str, Any]) -> dict[str, Any]:
    """The record with every declared nondeterministic field removed."""
    view = copy.deepcopy(record)
    view.pop("environment", None)
    view.pop("timing", None)
    for suite in view.get("suites", []):
        suite.pop("elapsed_ms", None)
    return view


def validate(
    root: Path,
    record: Any,
    candidate_sha: str | None = None,
    certification: bool = False,
    compare: Any = None,
) -> list[str]:
    schema = load_json(root / SCHEMA_PATH)
    check_schema_supported(schema)
    errors = schema_errors(schema, record, schema, "$")
    if errors:
        return errors
    registry = read_registry(root)
    fixture_sha, fixture_cases = fixture_identity(root)
    errors = semantic_errors(record, registry, fixture_sha, fixture_cases, candidate_sha)
    if certification:
        errors.extend(certification_errors(record))
    if compare is not None:
        other = schema_errors(schema, compare, schema, "$")
        if not other:
            other = semantic_errors(compare, registry, fixture_sha, fixture_cases, candidate_sha)
        if other:
            errors.extend(f"comparison record: {item}" for item in other)
        elif deterministic_view(record) != deterministic_view(compare):
            errors.append("records differ outside the declared nondeterministic fields")
    return errors


# ---------------------------------------------------------------------------
# Self-test
# ---------------------------------------------------------------------------


def _outcome(status: str, detail: str | None, reason: str | None) -> dict[str, Any]:
    return {"status": status, "detail": detail, "reason": reason}


def synthetic_record(root: Path) -> dict[str, Any]:
    """A passing KPR_Test_RunAll record shaped exactly like the VBA writer's."""
    registry = read_registry(root)
    fixture_sha, fixture_cases = fixture_identity(root)
    operator = "Outside the regression runner; the certification operator records this outcome."
    suites: list[dict[str, Any]] = [
        {"name": name, "kind": suite_kind(name), "status": "PASS",
         "assertions": 10 + index, "failures": 0, "elapsed_ms": 5}
        for index, name in enumerate(registry)
    ]
    assertions = sum(suite["assertions"] for suite in suites)
    return {
        "schema": "kpr-test-evidence",
        "schema_version": 1,
        "source_sha": "0123456789abcdef0123456789abcdef01234567",
        "runner": {
            "module": "KPR_REGRESSION_TESTS",
            "entry_point": "KPR_Test_RunAll",
            "selection": "all",
            "fixture_source_sha256": fixture_sha,
            "fixture_cases": fixture_cases,
        },
        "environment": {
            "excel_version": "16.0",
            "excel_build": "19029",
            "office_bitness": "64-bit",
            "vba_version": "VBA7",
            "operating_system": "Windows (64-bit) NT 10.00",
            "locale": {"country_code": 1, "country_setting": 39, "decimal_separator": ",",
                       "list_separator": ";", "date_order": "DMY"},
            "workbook_date_system": 1900,
            "caller": {"application_caller_type": "Error", "direct_suites": "direct-vba-1900",
                       "worksheet_suites": "scratch-workbook-range"},
        },
        "timing": {"started_local": "2026-09-25T10:00:00",
                   "finished_local": "2026-09-25T10:00:09", "elapsed_ms": 9000},
        "nondeterministic_fields": list(NONDETERMINISTIC_FIELDS),
        "suites": suites,
        "failures": [],
        "totals": {"suites": len(suites), "assertions": assertions, "failures": 0},
        "state_restoration": {"status": "PASS", "detail": "restored"},
        "certification": {
            "source_import": _outcome("NOT_RUN", None, operator),
            "vba_compile": _outcome("NOT_RUN", None, operator),
            "regression": _outcome("PASS", f"{len(suites)} suite(s)", None),
            "cross_oracle": _outcome("NOT_RUN", None, "No cross-oracle suite is registered."),
            "macro_options": _outcome("NOT_RUN", None, operator),
            "ribbonx": _outcome("NOT_RUN", None, operator),
            "commandbars": _outcome("NOT_RUN", None, operator),
            "demo_generation": _outcome("NOT_RUN", None, operator),
            "source_round_trip": _outcome("NOT_RUN", None, operator),
        },
        "result": "PASS",
    }


def _failed(record: dict[str, Any]) -> dict[str, Any]:
    """Turn a passing record into a consistent single-failure record."""
    record["suites"][0]["status"] = "FAIL"
    record["suites"][0]["failures"] = 1
    record["failures"] = [{"suite": record["suites"][0]["name"], "case": "date-type/x",
                           "detail": "expected Date 2024-02-29, got error 2015 (#VALUE!)"}]
    record["totals"]["failures"] = 1
    record["result"] = "FAIL"
    record["certification"]["regression"]["status"] = "FAIL"
    return record


def _single_suite(record: dict[str, Any]) -> dict[str, Any]:
    suite = next(item for item in record["suites"] if item["name"] == "fixtures")
    record["runner"]["entry_point"] = "KPR_Test_RunSuite"
    record["runner"]["selection"] = "fixtures"
    record["suites"] = [suite]
    record["totals"] = {"suites": 1, "assertions": suite["assertions"], "failures": 0}
    return record


def _certified(record: dict[str, Any]) -> dict[str, Any]:
    for key in CERTIFICATION_KEYS:
        record["certification"][key] = _outcome("PASS", f"observed {key}", None)
    record["certification"]["ribbonx"] = _outcome("NOT_APPLICABLE", None, "No RibbonX part.")
    return record


def _set(path: str, value: Any) -> Any:
    def mutate(record: dict[str, Any]) -> dict[str, Any]:
        node: Any = record
        keys = path.split("/")
        for key in keys[:-1]:
            node = node[int(key)] if isinstance(node, list) else node[key]
        last = keys[-1]
        if isinstance(node, list):
            node[int(last)] = value
        elif value is _DELETE:
            del node[last]
        else:
            node[last] = value
        return record
    return mutate


_DELETE = object()


def writer_consistency_errors(root: Path) -> list[str]:
    """Every schema property and certification key must appear in the VBA writer."""
    schema = load_json(root / SCHEMA_PATH)
    text = (root / RUNNER_PATH).read_bytes().decode("cp1252")
    errors: list[str] = []
    names = set(schema["properties"])
    for node in ("runner", "environment", "timing", "totals", "state_restoration"):
        names |= set(schema["properties"][node]["properties"])
    names |= set(schema["properties"]["certification"]["properties"])
    names |= set(schema["properties"]["suites"]["items"]["properties"])
    names |= set(schema["properties"]["failures"]["items"]["properties"])
    for name in sorted(names):
        if f'""{name}""' not in text:
            errors.append(f"{RUNNER_PATH} never writes property {name!r}")
    if tuple(schema["properties"]["certification"]["required"]) != CERTIFICATION_KEYS:
        errors.append("schema certification keys differ from CERTIFICATION_KEYS")
    statuses = schema["$defs"]["outcome"]["properties"]["status"]["enum"]
    if tuple(statuses) != OUTCOME_STATUSES:
        errors.append("schema outcome statuses differ from OUTCOME_STATUSES")
    if tuple(schema["properties"]["nondeterministic_fields"]["items"]["enum"]) != NONDETERMINISTIC_FIELDS:
        errors.append("schema nondeterministic fields differ from NONDETERMINISTIC_FIELDS")
    if f'"{EVIDENCE_NAME}"' not in text:
        errors.append(f"{RUNNER_PATH} does not write {EVIDENCE_NAME}")
    return errors


def self_test(root: Path) -> int:
    base = synthetic_record(root)
    registry = read_registry(root)
    problems = writer_consistency_errors(root)
    if "fixtures" not in registry or not any(name.startswith("worksheet-") for name in registry):
        problems.append("TestRegistry must include fixtures and the worksheet suites")

    positives: list[tuple[str, Any, dict[str, Any]]] = [
        ("passing all-suite record", base, {}),
        ("consistent failing record", _failed(copy.deepcopy(base)), {}),
        ("single-suite record", _single_suite(copy.deepcopy(base)), {}),
        ("complete certification record", _certified(copy.deepcopy(base)), {"certification": True}),
        ("candidate binding", base, {"candidate_sha": base["source_sha"]}),
        ("determinism across timing", base, {"compare": _set("timing/elapsed_ms", 1)(
            _set("suites/0/elapsed_ms", 99)(copy.deepcopy(base)))}),
    ]
    for label, record, options in positives:
        errors = validate(root, record, **options)
        status = "PASS" if not errors else "FAIL"
        print(f"{status} positive: {label}")
        problems.extend(f"positive {label}: {error}" for error in errors)

    negatives: list[tuple[str, Any, dict[str, Any], str]] = [
        ("missing property", _set("result", _DELETE), {}, "missing required property"),
        ("extra property", _set("extra", 1), {}, "unexpected property"),
        ("short SHA", _set("source_sha", "0123"), {}, "does not match"),
        ("boolean as integer", _set("totals/failures", False), {}, "expected type"),
        ("unknown status", _set("suites/0/status", "SKIPPED"), {}, "is not one of"),
        ("wrong candidate", lambda r: r, {"candidate_sha": "f" * 40}, "is not the candidate"),
        ("stale fixtures", _set("runner/fixture_source_sha256", "0" * 64), {}, "does not match"),
        ("fixture count", _set("runner/fixture_cases", 1), {}, "fixture_cases"),
        ("incomplete registry", lambda r: _drop_last_suite(r), {}, "complete TestRegistry"),
        ("wrong entry point", _set("runner/entry_point", "KPR_Test_RunSuite"), {}, "KPR_Test_RunAll"),
        ("unknown selection", lambda r: _set("runner/selection", "nope")(_single_suite(r)), {},
         "not in TestRegistry"),
        ("wrong kind", _set("suites/0/kind", "worksheet"), {}, "kind must be"),
        ("unlisted failure", _set("suites/0/failures", 1), {}, "listed"),
        ("PASS with failures", lambda r: _set("suites/0/status", "PASS")(_failed(r)), {}, "PASS with"),
        ("FAIL without failures", _set("suites/0/status", "FAIL"), {}, "FAIL without"),
        ("NOT_RUN with assertions", _set("suites/0/status", "NOT_RUN"), {}, "NOT_RUN with"),
        ("orphan failure", lambda r: _orphan_failure(r), {}, "did not run"),
        ("assertion total", _set("totals/assertions", 0), {}, "totals.assertions"),
        ("suite total", _set("totals/suites", 1), {}, "totals.suites"),
        ("failure total", _set("totals/failures", 3), {}, "totals.failures"),
        ("nondeterministic list", _set("nondeterministic_fields", ["/timing"]), {},
         "nondeterministic_fields"),
        ("time order", _set("timing/finished_local", "2026-09-24T10:00:00"), {}, "precedes"),
        ("result with failures", lambda r: _set("result", "PASS")(_failed(r)), {}, "result must be"),
        ("state failure passes", _set("state_restoration/status", "FAIL"), {}, "result must be"),
        ("PASS outcome without detail", _set("certification/regression/detail", None), {},
         "nonempty detail"),
        ("NOT_RUN outcome without reason", _set("certification/ribbonx/reason", None), {},
         "nonempty reason"),
        ("regression disagrees", _set("certification/regression", _outcome("FAIL", "x", None)), {},
         "regression.status"),
        ("certification with NOT_RUN", lambda r: r, {"certification": True}, "certification needs"),
        ("certification of one suite", lambda r: _certified(_single_suite(r)), {"certification": True},
         "KPR_Test_RunAll record"),
        ("compile marked not applicable",
         lambda r: _set("certification/vba_compile", _outcome("NOT_APPLICABLE", None, "n/a"))(
             _certified(r)), {"certification": True}, "vba_compile"),
        ("invalid comparison record", lambda r: r,
         {"compare": _set("timing/finished_local", "2026-09-24T10:00:00")(copy.deepcopy(base))},
         "comparison record: timing.finished_local precedes"),
        ("nondeterminism outside declared fields", lambda r: r,
         {"compare": _set("totals/assertions", base["totals"]["assertions"] + 1)(
             _set("suites/0/assertions", base["suites"][0]["assertions"] + 1)(copy.deepcopy(base)))},
         "differ outside"),
    ]
    for label, mutate, options, expected in negatives:
        errors = validate(root, mutate(copy.deepcopy(base)), **options)
        matched = any(expected in error for error in errors)
        print(f"{'PASS' if matched else 'FAIL'} negative: {label}")
        if not matched:
            problems.append(f"negative {label}: expected {expected!r}, got {errors}")

    with tempfile.TemporaryDirectory() as scratch:
        bad = Path(scratch) / "bad.json"
        bad.write_text("{", encoding="utf-8")
        try:
            load_json(bad)
        except EvidenceError:
            print("PASS negative: malformed JSON is a usage error")
        else:
            problems.append("malformed JSON was accepted")

    for problem in problems:
        print(f"ERROR: {problem}", file=sys.stderr)
    if problems:
        return 1
    print(f"PASS self-test: {len(positives)} positive and {len(negatives) + 1} negative cases; "
          f"registry of {len(registry)} suites")
    return 0


def _drop_last_suite(record: dict[str, Any]) -> dict[str, Any]:
    dropped = record["suites"].pop()
    record["totals"]["suites"] -= 1
    record["totals"]["assertions"] -= dropped["assertions"]
    return record


def _orphan_failure(record: dict[str, Any]) -> dict[str, Any]:
    record = _failed(record)
    record["failures"][0]["suite"] = "not-a-suite"
    return record


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=Path("."), help="repository root")
    parser.add_argument("--evidence", type=Path, help=f"{EVIDENCE_NAME} to validate")
    parser.add_argument("--candidate-sha", help="full commit SHA the record must bind to")
    parser.add_argument("--compare", type=Path, help="second record that must match deterministically")
    parser.add_argument("--certification", action="store_true",
                        help="require a complete #52 certification record")
    parser.add_argument("--self-test", action="store_true", help="exercise synthetic records")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    try:
        if args.self_test:
            return self_test(root)
        if args.evidence is None:
            parser.error("--evidence is required unless --self-test is given")
        if args.candidate_sha is not None and not re.fullmatch(r"[0-9a-f]{40}", args.candidate_sha):
            raise EvidenceError("--candidate-sha must be a full 40-character lowercase SHA")
        record = load_json(args.evidence)
        compare = load_json(args.compare) if args.compare else None
        errors = validate(root, record, args.candidate_sha, args.certification, compare)
    except EvidenceError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    for error in errors:
        print(f"FAIL: {error}")
    if errors:
        return 1
    if record["result"] != "PASS":
        print(f"FAIL: {args.evidence} is valid evidence of a failing run")
        return 1
    print(f"PASS: {args.evidence} is valid evidence of a passing run")
    return 0


if __name__ == "__main__":
    sys.exit(main())
