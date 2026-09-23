from __future__ import annotations

import inspect
import json
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable

from _gatelib import parse_report_args as parse_args
from _gatelib import run_gate
from policy_coverage_cases_config import configuration_cases
from policy_coverage_cases_quality import quality_cases
from policy_coverage_cases_repo import repository_cases
from policy_coverage_core import (
    CORE_TOOL,
    TOOL_NAME,
    CoverageError,
    load_module,
    production_finding_sites,
    rule_by_id,
)


def run_core_coverage(root: Path) -> dict[str, Any]:
    module: Any = load_module(root / CORE_TOOL, "coverage_check_repo")
    source_sites = production_finding_sites(root / CORE_TOOL)
    executed: dict[str, set[str]] = {key: set() for key in source_sites}
    unexpected_sites: dict[str, set[str]] = {}
    current_case = "bootstrap"
    original_finding = module.finding

    def recording_finding(*args, **kwargs):
        frame = inspect.currentframe()
        caller = frame.f_back if frame is not None else None
        if caller is not None and Path(caller.f_code.co_filename).resolve() == (root / CORE_TOOL).resolve():
            key = f"{caller.f_code.co_name}:{caller.f_lineno}"
            if key in executed:
                executed[key].add(current_case)
            else:
                unexpected_sites.setdefault(key, set()).add(current_case)
        return original_finding(*args, **kwargs)

    module.finding = recording_finding
    case_results: list[dict[str, Any]] = []

    def run_case(name: str, expected_rule: str, mutate: Callable[[Path], None], pattern: str | None = None) -> None:
        nonlocal current_case
        current_case = name
        with tempfile.TemporaryDirectory(prefix=f"policy-coverage-{name}-") as temporary:
            fixture_root = Path(temporary)
            module._initialize_fixture(fixture_root)
            try:
                mutate(fixture_root)
                report = module.build_report(fixture_root)
                rule = rule_by_id(report, expected_rule)
                status = "pass"
                detail = ""
                if rule is None:
                    status = "fail"
                    detail = f"expected rule {expected_rule!r} did not run"
                elif rule.get("status") != "fail":
                    status = "fail"
                    detail = f"expected rule {expected_rule!r} did not fail"
                elif pattern is not None and not any(pattern.casefold() in str(item.get("message", "")).casefold() for item in rule.get("findings", [])):
                    status = "fail"
                    detail = f"expected finding pattern was absent: {pattern!r}"
            except Exception as error:
                status = "fail"
                detail = f"fixture raised unexpectedly: {type(error).__name__}: {error}"
            case_results.append({"id": name, "owner": expected_rule, "status": status, "detail": detail})

    for expected_rule, mutate in module.SELF_TEST_CASES:
        run_case(f"rule-{expected_rule}", expected_rule, mutate)
    for case_name, expected_rule, mutate in module.BRANCH_SELF_TEST_CASES:
        run_case(f"branch-{case_name}", expected_rule, mutate)

    selected_quality_cases = [
        item for item in quality_cases(module)
        if item[0] != "vba-export-name-not-leading"
    ]

    def late_class_header(fixture_root: Path) -> None:
        prefix = "\n".join(
            ["VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1", "END"]
            + ["' pad"] * 18
        )
        module._write_fixture(
            fixture_root / "src/modules/Late.cls",
            prefix + '\nAttribute VB_Name = "Late"\nOption Explicit\n',
            crlf=True,
        )
        module._run_git(fixture_root, "add", "src/modules/Late.cls")

    extra_cases = [
        *configuration_cases(module),
        *repository_cases(module),
        *selected_quality_cases,
        (
            "vba-export-name-not-leading",
            "vba-export-header",
            "leading export header",
            late_class_header,
        ),
    ]
    for name, expected_rule, pattern, mutate in extra_cases:
        run_case(name, expected_rule, mutate, pattern)

    current_case = "operational-rule-crash"
    with tempfile.TemporaryDirectory(prefix="policy-coverage-operational-") as temporary:
        fixture_root = Path(temporary)
        module._initialize_fixture(fixture_root)
        original_checks = module.CHECKS

        def crash(repo, config):
            del repo, config
            raise RuntimeError("fixture operational failure")

        crash.__name__ = "check_fixture_operational_failure"
        module.CHECKS = (crash,) + original_checks
        report = module.build_report(fixture_root)
        module.CHECKS = original_checks
        rule = rule_by_id(report, "fixture-operational-failure")
        ok = rule is not None and rule.get("status") == "fail" and any("Rule could not complete" in str(item.get("message", "")) for item in rule.get("findings", []))
        case_results.append({"id": "operational-rule-crash", "owner": "build-report", "status": "pass" if ok else "fail", "detail": "" if ok else "rule exception was not mapped to an explicit policy finding"})

    uncovered = [{**source_sites[key], "cases": []} for key in source_sites if not executed[key]]
    covered = [{**source_sites[key], "cases": sorted(executed[key])} for key in source_sites if executed[key]]
    return {
        "source": CORE_TOOL,
        "sites": len(source_sites),
        "covered": covered,
        "uncovered": uncovered,
        "unexpected_sites": [{"id": key, "cases": sorted(value)} for key, value in sorted(unexpected_sites.items())],
        "cases": case_results,
    }


def run_tool_selftest(root: Path, relative: str) -> dict[str, Any]:
    completed = subprocess.run([sys.executable, str(root / relative), "--root", str(root), "--self-test"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    output = (completed.stdout + completed.stderr).strip()
    return {"tool": relative, "status": "pass" if completed.returncode == 0 else "fail", "exit_code": completed.returncode, "summary": output.splitlines()[-1] if output else "no output"}


def operational_exit_fixture(root: Path) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix="policy-coverage-not-git-") as temporary:
        completed = subprocess.run([sys.executable, str(root / CORE_TOOL), "--root", temporary], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    return {"id": "top-level-operational-error", "status": "pass" if completed.returncode == 2 else "fail", "exit_code": completed.returncode, "expected_exit_code": 2}


def delegated_workflow_fixtures(root: Path) -> dict[str, Any]:
    workflow_source = (root / "tools/test_workflow_validation.py").read_text(encoding="utf-8")
    workflow = (root / ".github/workflows/static-checks.yml").read_text(encoding="utf-8")
    fixture_probes = ("valid-local-action", "invalid-yaml", "duplicate-job", "invalid-job-structure", "malformed-local-action", "missing-local-entrypoint")
    missing = [probe for probe in fixture_probes if probe not in workflow_source]
    wired = "Validate workflows and authoritative fixtures" in workflow and "AUTHORITATIVE_WORKFLOWS_OUTCOME" in workflow
    return {"owner": "tools/test_workflow_validation.py", "execution": "delegated-authoritative-workflow-step", "fixtures": list(fixture_probes), "status": "pass" if not missing and wired else "fail", "missing": missing, "terminally_required": wired}


def build_report(root: Path) -> dict[str, Any]:
    core = run_core_coverage(root)
    focused_tools = (
        "tools/check_committed_whitespace.py",
        "tools/check_vba_jumps.py",
        "tools/check_vba_conditionals.py",
        "tools/check_vba_public_api.py",
        "tools/check_local_actions.py",
        "tools/check_release_semantics.py",
        "tools/check_release.py",
    )
    focused = [run_tool_selftest(root, path) for path in focused_tools]
    operational = operational_exit_fixture(root)
    delegated = delegated_workflow_fixtures(root)
    failures: list[str] = []
    if core["uncovered"]:
        failures.append(f"{len(core['uncovered'])} canonical finding site(s) are not exercised")
    if core["unexpected_sites"]:
        failures.append(f"{len(core['unexpected_sites'])} runtime finding site(s) were not in the static inventory")
    bad_cases = [case for case in core["cases"] if case["status"] != "pass"]
    if bad_cases:
        failures.append(f"{len(bad_cases)} canonical branch fixture(s) failed")
    bad_tools = [tool for tool in focused if tool["status"] != "pass"]
    if bad_tools:
        failures.append(f"{len(bad_tools)} focused gate self-test(s) failed")
    if operational["status"] != "pass":
        failures.append("top-level operational-error exit mapping failed")
    if delegated["status"] != "pass":
        failures.append("authoritative workflow fixture delegation is incomplete")

    matrix = [{"id": item["id"], "owner": item["function"], "source_line": item["line"], "fixtures": item["cases"], "status": "covered"} for item in core["covered"]]
    matrix.extend({"id": f"focused:{Path(tool['tool']).stem}", "owner": tool["tool"], "source_line": None, "fixtures": ["--self-test"], "status": "covered" if tool["status"] == "pass" else "failed"} for tool in focused)
    matrix.append({"id": "delegated:authoritative-workflows", "owner": delegated["owner"], "source_line": None, "fixtures": delegated["fixtures"], "status": "covered" if delegated["status"] == "pass" else "failed"})
    matrix.append({"id": "operational:top-level-exit-2", "owner": CORE_TOOL, "source_line": None, "fixtures": [operational["id"]], "status": "covered" if operational["status"] == "pass" else "failed"})

    exclusions = [
        {"id": "line-coverage-percentage", "reason": "Numeric line coverage is deliberately excluded because certification depends on asserted policy outcomes, not executed lines.", "blocking": False},
        {"id": "hosted-service-failures", "reason": "External network/service behavior is outside the dependency-free fixture model; every required hosted step is nevertheless terminally enforced.", "blocking": False},
    ]
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": "pass" if not failures else "fail",
        "counts": {"canonical_sites": core["sites"], "canonical_covered": len(core["covered"]), "canonical_uncovered": len(core["uncovered"]), "canonical_cases": len(core["cases"]), "focused_selftests": len(focused), "matrix_rows": len(matrix), "exclusions": len(exclusions)},
        "matrix": matrix,
        "uncovered": core["uncovered"],
        "canonical_cases": core["cases"],
        "focused": focused,
        "delegated": delegated,
        "operational": operational,
        "exclusions": exclusions,
        "failures": failures,
    }


def markdown_report(report: dict[str, Any]) -> str:
    counts = report["counts"]
    lines = ["## Policy branch coverage", "", f"- **Status:** {str(report['status']).upper()}", f"- **Canonical finding sites:** {counts['canonical_covered']}/{counts['canonical_sites']} covered", f"- **Canonical synthetic cases:** {counts['canonical_cases']}", f"- **Focused gate self-tests:** {counts['focused_selftests']}", f"- **Matrix rows:** {counts['matrix_rows']}", f"- **Reviewed non-blocking exclusions:** {counts['exclusions']}"]
    if report["uncovered"]:
        lines.extend(["", "### Uncovered canonical finding sites", ""])
        for item in report["uncovered"]:
            lines.append(f"- `{item['id']}` — `{item['expression']}`")
    bad_cases = [case for case in report["canonical_cases"] if case["status"] != "pass"]
    if bad_cases:
        lines.extend(["", "### Failed fixture assertions", ""])
        for case in bad_cases:
            lines.append(f"- `{case['id']}` — {case['detail']}")
    if report["failures"]:
        lines.extend(["", "### Coverage failures", ""])
        lines.extend(f"- {message}" for message in report["failures"])
    lines.extend(["", "### Reviewed exclusions", ""])
    for item in report["exclusions"]:
        lines.append(f"- `{item['id']}` — {item['reason']}")
    return "\n".join(lines) + "\n"




def run_self_test(root: Path) -> int:
    first = build_report(root)
    second = build_report(root)
    first_json = json.dumps(first, indent=2, sort_keys=True) + "\n"
    second_json = json.dumps(second, indent=2, sort_keys=True) + "\n"
    failures: list[str] = []
    if first["status"] != "pass":
        failures.append("policy coverage report is not complete")
    if first_json != second_json:
        failures.append("policy coverage JSON is not deterministic")
    if markdown_report(first) != markdown_report(second):
        failures.append("policy coverage Markdown is not deterministic")
    if failures:
        for message in failures:
            print(f"[FAIL] {message}")
        print(markdown_report(first).rstrip())
        return 1
    print("SELF-TEST PASS: every canonical finding site, focused hardening gate, delegated authoritative workflow fixture, and operational-error mapping is covered deterministically.")
    return 0




def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    return run_gate(
        options,
        build=lambda: build_report(options.root),
        markdown=markdown_report,
        errors=(
            OSError,
            UnicodeError,
            CoverageError,
            RuntimeError,
            subprocess.SubprocessError,
            json.JSONDecodeError,
        ),
        self_test=lambda: run_self_test(options.root),
    )

if __name__ == "__main__":
    raise SystemExit(main())
