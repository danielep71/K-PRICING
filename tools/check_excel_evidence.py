#!/usr/bin/env python3
"""Validate optional host evidence, without executing Excel or authenticating logs."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Any

from _gatelib import run_gate
from release_provenance import committed, decode, nonempty, object_keys, relative, require

POLICY = ".github/excel-evidence-policy.json"
STAGES = ("import", "compile", "regression", "cleanup")
OUTCOMES = {"import": "IMPORT_FAILED", "compile": "COMPILE_FAILED",
            "regression": "TEST_FAILED", "cleanup": "CLEANUP_FAILED"}


def positive(value: Any) -> bool:
    return type(value) is int and value > 0


def environment_summary(record: dict[str, Any]) -> str:
    return (f"execution={record['execution']}; environment="
            + json.dumps(record["environment"], sort_keys=True, separators=(",", ":")))


def source_inventory(root: Path, sha: str, config: dict[str, Any]) -> list[dict[str, str]]:
    vba = config["vba"]
    components = vba["components"]
    require(isinstance(components, dict), "candidate must declare VBA components")
    paths = {path for path, role in components.items() if role != "example"}
    # Companion form resources are part of the import even though not VBA modules.
    for path in tuple(paths):
        if path.endswith(".frm"):
            result = subprocess.run(["git", "-C", str(root), "cat-file", "-e",
                                     f"{sha}:{path[:-4]}.frx"], capture_output=True, check=False)
            if result.returncode == 0:
                paths.add(path[:-4] + ".frx")
    require(bool(paths), "candidate has no importable source")
    return [{"path": path, "sha256": hashlib.sha256(committed(root, sha, path)).hexdigest()}
            for path in sorted(paths)]


def load_policy(root: Path, sha: str) -> dict[str, Any]:
    policy = decode(committed(root, sha, POLICY))
    object_keys(policy, "schema_version entry_point cases assertions expected_error_cases", "host policy")
    require(type(policy["schema_version"]) is int and policy["schema_version"] == 1,
            "unsupported host policy schema")
    require(nonempty(policy["entry_point"]) and positive(policy["assertions"]), "invalid harness policy")
    for field in ("cases", "expected_error_cases"):
        values = policy[field]
        require(isinstance(values, list) and all(nonempty(value) for value in values),
                f"invalid {field} policy")
        require(len(values) == len(set(values)), f"duplicate {field} policy")
    require(bool(policy["cases"]) and set(policy["expected_error_cases"]) <= set(policy["cases"]),
            "expected-error cases must belong to the declared suite")
    return policy


def validate_identity(record: Any, config: dict[str, Any], sha: str) -> None:
    object_keys(record, "schema_version repository candidate_sha template_contract execution "
                "availability_reason started_at finished_at runner environment sources stages harness", "host evidence")
    require(type(record["schema_version"]) is int and record["schema_version"] == 1,
            "unsupported host evidence schema")
    for key, value in {"repository": config["repository"], "candidate_sha": sha,
                       "template_contract": config["template_contract"]}.items():
        require(record[key] == value, f"host evidence {key} differs from candidate")
    require(record["execution"] in ("manual", "automated", "unavailable"), "invalid execution mode")
    times = []
    for key in ("started_at", "finished_at"):
        require(nonempty(record[key]), f"{key} is required")
        stamp = datetime.fromisoformat(record[key].replace("Z", "+00:00"))
        require(stamp.utcoffset() is not None, "timestamps require a timezone")
        times.append(stamp)
    require(times[1] >= times[0], "finish precedes start")


def validate_environment(record: dict[str, Any]) -> None:
    runner = record["runner"]
    object_keys(runner, "class identity workflow", "runner")
    require(nonempty(runner["identity"]), "runner/operator identity is required")
    if record["execution"] == "manual":
        require(runner["class"] == "manual-interactive" and runner["workflow"] is None,
                "manual execution must not claim a hosted workflow")
    else:
        require(runner["class"] == "trusted-interactive", "automated host requires eligible runner class")
        workflow = runner["workflow"]
        object_keys(workflow, "repository path sha run_id run_attempt", "runner workflow")
        require(nonempty(workflow["repository"]) and relative(workflow["path"])
                and workflow["path"].startswith(".github/workflows/"), "invalid workflow identity")
        require(isinstance(workflow["sha"], str) and re.fullmatch(r"[0-9a-f]{40}", workflow["sha"])
                and positive(workflow["run_id"]) and positive(workflow["run_attempt"]),
                "workflow requires immutable SHA and positive run/attempt")
    environment = record["environment"]
    object_keys(environment, "excel_version excel_build office_bitness os os_architecture runtime "
                "macro_policy vba_project_access trust_changes", "Excel environment")
    require(environment["office_bitness"] in ("32-bit", "64-bit"), "invalid Office bitness")
    require(environment["trust_changes"] is False, "host job must not change trust configuration")
    require(all(nonempty(value) for key, value in environment.items() if key != "trust_changes"),
            "Excel environment fields must be nonempty")
    require(environment["os"].casefold().startswith("windows"), "this interface requires Windows Excel")


def retained_log(directory: Path, value: Any) -> str:
    object_keys(value, "path sha256", "retained log")
    require(relative(value["path"]), "unsafe log path")
    path = directory / value["path"]
    require(not any(parent.is_symlink() for parent in (path, *path.parents)), "symlinked log path")
    require(path.resolve().is_relative_to(directory.resolve()) and path.is_file(), "missing log")
    raw = path.read_bytes()
    require(hashlib.sha256(raw).hexdigest() == value["sha256"], "retained log digest mismatch")
    return raw.decode("utf-8-sig")


def validate_stages(record: dict[str, Any], directory: Path) -> tuple[list[str], dict[str, str]]:
    stages = record["stages"]
    object_keys(stages, "import compile regression cleanup", "stages")
    outcomes: list[str] = []
    logs = {}
    for name in STAGES:
        stage = stages[name]
        object_keys(stage, "status detail log", name)
        status = stage["status"]
        require(status in ("PASS", "FAIL", "NOT_RUN", "TIMEOUT"), f"invalid {name} status")
        require(nonempty(stage["detail"]), f"{name} requires detail")
        if status == "NOT_RUN":
            require(stage["log"] is None, "NOT_RUN cannot claim execution logs")
            outcomes.append("INCOMPLETE")
        else:
            logs[name] = retained_log(directory, stage["log"])
            if status != "PASS":
                outcomes.append("EXECUTION_TIMEOUT" if status == "TIMEOUT" else OUTCOMES[name])
    if stages["compile"]["status"] == "PASS":
        require(stages["import"]["status"] == "PASS", "compile PASS requires imported source")
    if stages["regression"]["status"] == "PASS":
        require(stages["compile"]["status"] == "PASS", "test PASS requires compile PASS")
    if stages["import"]["status"] != "PASS":
        require(stages["compile"]["status"] == "NOT_RUN", "compile must not run after failed import")
    if stages["compile"]["status"] != "PASS":
        require(stages["regression"]["status"] == "NOT_RUN", "regression must not run after failed compile")
    return list(dict.fromkeys(outcomes)), logs


def validate_harness(record: dict[str, Any], policy: dict[str, Any], log: str) -> None:
    harness = record["harness"]
    object_keys(harness, "entry_point cases assertions failures completeness expected_errors", "harness")
    require(harness["entry_point"] == policy["entry_point"], "wrong harness entry point")
    for field in ("cases", "assertions", "failures"):
        require(type(harness[field]) is int and harness[field] >= 0, f"invalid {field}")
    errors = harness["expected_errors"]
    require(isinstance(errors, list), "expected_errors must be an array")
    observed = []
    for error in errors:
        object_keys(error, "case status detail", "expected error")
        require(nonempty(error["case"]) and error["status"] in ("PASS", "FAIL", "NOT_RUN")
                and nonempty(error["detail"]), "invalid expected-error result")
        observed.append(error["case"])
    require(observed == policy["expected_error_cases"], "expected-error results differ from policy")
    if record["stages"]["regression"]["status"] != "PASS":
        return
    require(harness["cases"] == len(policy["cases"]) and harness["assertions"] == policy["assertions"]
            and harness["failures"] == 0 and harness["completeness"] == "COMPLETE"
            and all(error["status"] == "PASS" for error in errors), "test PASS contradicts harness results")
    cases = re.findall(r"^CASE=(.+)$", log, re.MULTILINE)
    require(cases == policy["cases"], "raw harness cases differ from policy")
    summaries = re.findall(r"^RESULT=(.+)$", log, re.MULTILINE)
    require(len(summaries) == 1, "raw harness must contain one summary")
    for field in ("cases", "assertions", "failures"):
        require(re.findall(rf"^{field.upper()}=(.+)$", log, re.MULTILINE) == [str(harness[field])],
                f"raw {field} differs from harness record")
    cleanup_pass = summaries[0].endswith("; cleanup=PASS")
    verdict, cleanup = ("PASS", "PASS") if cleanup_pass else ("FAIL", "FAIL")
    require(summaries[0] == (
        f"{verdict}; completeness=COMPLETE; cases={harness['cases']}; assertions={harness['assertions']}; "
        f"failures=0; cleanup={cleanup}"), "raw harness summary contradicts complete regression")
    if not cleanup_pass:
        require(record["stages"]["cleanup"]["status"] == "FAIL", "harness cleanup failure must remain failed")


def evaluate(root: Path, sha: str, evidence: Path) -> dict[str, Any]:
    execution = "unresolved"
    outcomes: list[str] = []
    findings: list[str] = []
    release_environment = None
    try:
        head = subprocess.run(["git", "-C", str(root), "rev-parse", "HEAD"],
                              capture_output=True, text=True, check=False)
        require(head.returncode == 0 and head.stdout.strip() == sha, "checkout must be the exact candidate SHA")
        clean = subprocess.run(["git", "-C", str(root), "diff", "--quiet", sha, "--"], check=False)
        require(clean.returncode == 0, "candidate tracked tree must be unchanged")
        config = decode(committed(root, sha, ".github/repository-profile.json"))
        policy = load_policy(root, sha)
        record = decode(evidence.read_bytes())
        validate_identity(record, config, sha)
        execution = record["execution"]
        if execution == "unavailable":
            require(nonempty(record["availability_reason"]), "unavailable requires a reason")
            require(all(record[key] is None for key in ("runner", "environment", "harness"))
                    and record["sources"] == [] and record["stages"] == {},
                    "unavailable cannot carry runtime claims")
            outcomes = ["UNAVAILABLE"]
        else:
            require(record["availability_reason"] is None, "available execution has no unavailability reason")
            validate_environment(record)
            require(record["sources"] == source_inventory(root, sha, config),
                    "import inventory differs from exact candidate source")
            outcomes, logs = validate_stages(record, evidence.parent)
            if record["stages"]["regression"]["status"] == "NOT_RUN":
                require(record["harness"] is None, "NOT_RUN cannot carry harness results")
            else:
                validate_harness(record, policy, logs.get("regression", "").replace("\r\n", "\n"))
            if not outcomes:
                release_environment = environment_summary(record)
    except (ValueError, TypeError, KeyError, OSError, subprocess.SubprocessError) as error:
        outcomes.append("EVIDENCE_INVALID")
        findings.append(str(error))
    return {"status": "fail" if outcomes else "pass", "candidate_sha": sha, "execution": execution,
            "release_environment": release_environment,
            "outcomes": outcomes or ["PASS"], "findings": findings,
            "scope_note": "Validates retained host assertions and log bindings; does not execute Excel or authenticate execution."}


def markdown(report: dict[str, Any]) -> str:
    return (f"# Excel evidence validation\n\nResult: {report['status'].upper()}\n\n"
            f"Execution: {report['execution']}\n\nOutcomes: {', '.join(report['outcomes'])}\n\n"
            + "\n".join(f"- {item}" for item in report["findings"])
            + f"\n\n{report['scope_note']}\n")


def release_findings(root: Path, sha: str, evidence_path: Path,
                     host_path: Path | None) -> list[dict[str, str]]:
    try:
        evidence = decode(evidence_path.read_bytes())
        require(isinstance(evidence, dict) and isinstance(evidence.get("checks"), dict),
                "release checks must be an object")
        binding = evidence["checks"].get("excel-host-evidence")
        if binding is None and host_path is None:
            return []
        require(host_path is not None and isinstance(binding, dict),
                "host evidence requires both --excel-evidence and the excel-host-evidence check")
        assert host_path is not None
        report = evaluate(root, sha, host_path)
        require(report["status"] == "pass", "host outcomes: " + ", ".join(report["outcomes"]))
        raw = host_path.read_bytes()
        record = decode(raw)
        require(binding.get("sha256") == hashlib.sha256(raw).hexdigest()
                and binding.get("execution") == record["execution"], "host evidence binding mismatch")
        regression = evidence["checks"].get("regression", {})
        for key in ("vba-compile", "regression"):
            require(evidence["checks"].get(key, {}).get("environment") == environment_summary(record),
                    f"release/host {key} environment mismatch")
        for field in ("entry_point", "cases", "assertions", "failures", "completeness"):
            require(regression.get(field) == record["harness"][field], f"release/host {field} mismatch")
        return []
    except (ValueError, TypeError, KeyError, OSError) as error:
        return [{"code": "excel-host-evidence", "path": str(host_path or "<not supplied>"),
                 "message": str(error)}]


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--candidate-sha", required=True)
    parser.add_argument("--evidence", type=Path, required=True)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    options = parser.parse_args(arguments)
    for path in (options.output, options.summary):
        if path is not None and (path.resolve() == options.evidence.resolve()
                                 or path.resolve().is_relative_to(options.evidence.resolve().parent)):
            parser.error("reports must be outside the retained evidence directory")
    return run_gate(options, build=lambda: evaluate(options.root, options.candidate_sha, options.evidence),
                    markdown=markdown, errors=(OSError,))


if __name__ == "__main__":
    raise SystemExit(main())
