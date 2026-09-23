#!/usr/bin/env python3
"""Verify the development contract of the canonical portable repository checker."""

from __future__ import annotations

import argparse
import ast
import contextlib
import hashlib
import importlib.util
import io
import json
import subprocess
import sys
import tempfile
from pathlib import Path
from types import ModuleType
from typing import Any

from _gatelib import parse_report_args as parse_arguments
from _gatelib import run_gate

TOOL_NAME = "Checker development contract"
CHECKER_PATH = Path("tools/check_repo.py")
GATELIB_PATH = Path("tools/_gatelib.py")
SECTION_STARTS = (
    ("runtime-core", "Repository"),
    ("configuration", "_same_keys"),
    ("repository-policy", "check_required_paths"),
    ("vba-policy", "_vba_paths"),
    ("reporting", "build_report"),
    ("fixtures", "_write_fixture"),
    ("cli", "parse_arguments"),
)
EXPECTED_CHECK_FUNCTIONS = (
    "check_required_paths",
    "check_placeholders",
    "check_identity",
    "check_dotfile_policy",
    "check_structured_data",
    "check_markdown_links",
    "check_text_integrity",
    "check_forbidden_artifacts",
    "check_line_endings",
    "check_label_manifest",
    "check_issue_forms",
    "check_workflow_actions",
    "check_version_changelog",
    "check_git_diff",
    "check_vba_option_explicit",
    "check_vba_export_header",
    "check_vba_structure",
    "check_vba_visibility",
    "check_generated_vba_contract",
    "check_vba_public_api",
)
GATE_RUNNER_CONSUMERS = frozenset(
    {
        "check_documentation.py",
        "check_wiki.py",
        "check_external_links.py",
        "check_excel_evidence.py",
        "check_portfolio_drift.py",
        "report_portfolio_quality.py",
        "provision_repository.py",
        "check_committed_whitespace.py",
        "check_local_actions.py",
        "check_release_semantics.py",
        "check_template_contract.py",
        "check_vba_conditionals.py",
        "check_vba_jumps.py",
        "check_vba_public_api.py",
        "checker_development.py",
        "policy_coverage_runner.py",
    }
)
GATE_RUNNER_EXCLUSIONS = {
    "_release_closeout.py": (
        "private post-release provider-snapshot validator with a dedicated self-test; "
        "not a focused run_gate report CLI"
    ),
    "collect_portfolio_snapshot.py": "GET-only evidence capture, not a report gate",
    "create_reusable_workflow_fixture.py": "disposable consumer provisioning, not a report gate",
    "check_release.py": (
        "atomic evidence writes and a console rendering distinct from its Markdown summary"
    ),
    "release_certification.py": (
        "mutually exclusive build/verify bundle modes emitting JSON to stdout; "
        "not a focused run_gate report CLI"
    ),
    "test_workflow_validation.py": "text-only report with no JSON evidence output",
    "initialize_repository.py": "repository provisioning CLI, not a focused report gate",
}
# These CLIs deliberately keep their fixtures separate from operational arguments.
# Reasons identify the alternate test command, or explicitly disclose no offline suite.
SELF_TEST_EXCLUSIONS = {
    "test_documentation.py": "Unittest CLI; normal invocation runs its fixture suite",
    "test_excel_evidence.py": "Unittest CLI; normal invocation runs its fixture suite",
    "test_portfolio_drift.py": "Unittest CLI; normal invocation runs its fixture suite",
    "test_portfolio_quality.py": "Unittest CLI; normal invocation runs its fixture suite",
    "test_provision_repository.py": "Unittest CLI; normal invocation runs its fixture suite",
    "test_release_provenance.py": "Unittest CLI; normal invocation runs its fixture suite",
    "test_verification_depth.py": "Unittest CLI; focused release-critical failure-path coverage",
    "test_wiki.py": "Unittest CLI; normal invocation runs its fixture suite",
    "check_documentation.py": "Offline fixtures: python tools/test_documentation.py -v",
    "check_wiki.py": "Offline fixtures: python tools/test_wiki.py -v",
    "check_external_links.py": "Simulated HTTP fixtures: python tools/test_documentation.py -v",
    "check_excel_evidence.py": "Record fixtures: python tools/test_excel_evidence.py -v; no Office execution",
    "check_portfolio_drift.py": "Offline fixtures: python tools/test_portfolio_drift.py -v",
    "report_portfolio_quality.py": "Offline fixtures: python tools/test_portfolio_quality.py -v",
    "collect_portfolio_snapshot.py": "Simulated capture fixtures: python tools/test_portfolio_quality.py -v",
    "provision_repository.py": "Simulated provisioning: python tools/test_provision_repository.py -v",
    "test_workflow_validation.py": "Normal invocation runs authoritative actionlint fixtures",
    "create_reusable_workflow_fixture.py": (
        "Offline cases: python tools/test_workflow_validation.py; live consumer pilot evidence is separate"
    ),
}


ALLOWED_IMPORT_ROOTS = {
    "argparse",
    "fnmatch",
    "hashlib",
    "json",
    "pathlib",
    "re",
    "subprocess",
    "sys",
    "tempfile",
    "typing",
    "urllib",
    "xml",
    "__future__",
}


class ContractError(RuntimeError):
    """Raised when checker development assumptions cannot be evaluated."""


def load_checker(path: Path) -> ModuleType:
    spec = importlib.util.spec_from_file_location("checker_development_target", path)
    if spec is None or spec.loader is None:
        raise ContractError(f"Cannot load checker module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def definition_nodes(tree: ast.Module) -> list[ast.AST]:
    return [
        node
        for node in tree.body
        if isinstance(node, (ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef))
    ]


def node_name(node: ast.AST) -> str:
    if isinstance(node, (ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
        return node.name
    raise ContractError("Unexpected non-definition node.")


def section_report(source: str, tree: ast.Module) -> tuple[list[dict[str, Any]], list[str]]:
    del source
    definitions = definition_nodes(tree)
    positions = {node_name(node): index for index, node in enumerate(definitions)}
    failures: list[str] = []
    starts: list[tuple[str, int]] = []
    for section, sentinel in SECTION_STARTS:
        index = positions.get(sentinel)
        if index is None:
            failures.append(f"section {section!r} sentinel is missing: {sentinel}")
            continue
        starts.append((section, index))
    if len(starts) == len(SECTION_STARTS):
        indexes = [index for _, index in starts]
        if indexes != sorted(indexes) or len(indexes) != len(set(indexes)):
            failures.append("section sentinels are not strictly ordered")
    rows: list[dict[str, Any]] = []
    if failures:
        return rows, failures
    for offset, (section, start) in enumerate(starts):
        end = starts[offset + 1][1] if offset + 1 < len(starts) else len(definitions)
        owned = definitions[start:end]
        if not owned:
            failures.append(f"section {section!r} owns no definitions")
            continue
        rows.append(
            {
                "id": section,
                "start": node_name(owned[0]),
                "end": node_name(owned[-1]),
                "definitions": len(owned),
                "start_line": getattr(owned[0], "lineno", None),
                "end_line": getattr(
                    owned[-1], "end_lineno", getattr(owned[-1], "lineno", None)
                ),
            }
        )
    covered = sum(int(row["definitions"]) for row in rows)
    if covered != len(definitions):
        failures.append(
            f"section ownership covers {covered} of {len(definitions)} top-level definitions"
        )
    return rows, failures


def import_report(tree: ast.Module) -> tuple[list[str], list[str]]:
    observed: set[str] = set()
    failures: list[str] = []
    for node in tree.body:
        if isinstance(node, ast.Import):
            observed.update(alias.name.split(".", 1)[0] for alias in node.names)
        elif isinstance(node, ast.ImportFrom):
            if node.level:
                failures.append("runtime checker must not use relative imports")
                continue
            if node.module:
                observed.add(node.module.split(".", 1)[0])
    unsupported = sorted(observed - ALLOWED_IMPORT_ROOTS)
    if unsupported:
        failures.append(
            "runtime checker imports non-approved package roots: " + ", ".join(unsupported)
        )
    return sorted(observed), failures


def parser_tests(module: ModuleType) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []

    def record(name: str, ok: bool, detail: str = "") -> None:
        results.append({"id": name, "status": "pass" if ok else "fail", "detail": detail})

    yaml_errors = module.validate_yaml_subset("name: x\non: [push\n")
    record("yaml-invalid-flow", bool(yaml_errors), "malformed flow collection must fail")

    valid_yaml = module.validate_yaml_subset(
        "name: x\njobs:\n  test:\n    runs-on: ubuntu-latest\n"
    )
    record("yaml-basic-valid", not valid_yaml, f"unexpected errors: {valid_yaml!r}")

    slugs = module._github_slugs("# Alpha\n# Alpha\n<a id=\"stable-anchor\"></a>\n")
    record(
        "markdown-slugs",
        {"alpha", "alpha-1", "stable-anchor"}.issubset(slugs),
        f"observed slugs: {sorted(slugs)!r}",
    )

    sections = module._editorconfig_sections(
        "root = true\n\n[*.{bas,cls,frm}]\nend_of_line = crlf\n"
    )
    record(
        "editorconfig-parser",
        sections.get("", {}).get("root") == "true"
        and sections.get("*.{bas,cls,frm}", {}).get("end_of_line") == "crlf",
        f"observed sections: {sections!r}",
    )

    stripped = module._strip_vba_line(
        'Debug.Print "apostrophe stays"; value \' real comment'
    )
    record(
        "vba-line-strip",
        "real comment" not in stripped and "Debug.Print" in stripped,
        f"observed: {stripped!r}",
    )

    return results


def reusable_identity_tests(module: ModuleType) -> list[dict[str, Any]]:
    source = "example/TEMPLATE"
    reference = source + "/.github/workflows/static-checks.yml@" + "a" * 40
    path = ".github/workflows/consumer.yml"
    config: dict[str, Any] = {
        "mode": "generated",
        "identity": {"template_tokens": ["TEMPLATE"], "forbidden_tokens": []},
        "template_contract": {"source": source},
    }
    cases = [
        ("exact-pin", path, reference, False),
        ("quoted-pin", path, 'uses: "' + reference + '"', False),
        ("floating-ref", path, reference[:-40] + "main", True),
        ("short-sha", path, reference[:-1], True),
        ("long-sha", path, reference + "a", True),
        ("other-source", path, reference.replace("example/", "other/"), True),
        ("prefix-spoof", path, "other/" + reference, True),
        ("readme-branding", "README.md", reference, True),
        ("stray-branding", path, reference + "\n# TEMPLATE", True),
        ("path-traversal", path, reference.replace("static-checks", "../static-checks"), True),
    ]
    results = []
    for name, file, text, forbidden in cases:
        actual = module._identity_scan_text(file, text, config, "TEMPLATE")
        ok = ("TEMPLATE" in actual) == forbidden and len(actual) == len(text)
        results.append({"id": "reusable-identity-" + name,
                        "status": "pass" if ok else "fail", "detail": ""})
    config["identity"]["forbidden_tokens"] = ["TEMPLATE"]
    actual = module._identity_scan_text(path, reference, config, "TEMPLATE")
    results.append({"id": "reusable-identity-donor-still-forbidden",
                    "status": "pass" if actual == reference else "fail", "detail": ""})
    return results


def dependency_rollback_tests() -> list[dict[str, Any]]:
    """Prove a grouped pin/config revert restores bytes without rewriting history.

    SHA-shaped values are synthetic; no upstream dependency is executed or
    authenticated by this test. Provenance review remains a human merge gate.
    """
    with tempfile.TemporaryDirectory(prefix="dependency-rollback-") as temporary:
        root = Path(temporary)

        def command(*arguments: str) -> str:
            return subprocess.run(
                ["git", "-C", str(root), "-c", "user.name=Rollback Fixture",
                 "-c", "user.email=fixture@example.invalid", *arguments],
                check=True, capture_output=True, text=True,
            ).stdout.strip()

        command("init", "-b", "main")
        pin = root / "dependency.txt"
        coupled = root / "validator.txt"
        pin.write_text("1" * 40 + "\n", encoding="utf-8")
        coupled.write_text("1.0.0\n", encoding="utf-8")
        command("add", "--all")
        command("commit", "-m", "Record passing synthetic baseline")
        baseline = command("rev-parse", "HEAD")
        tree = command("rev-parse", "HEAD^{tree}")
        pin.write_text("2" * 40 + "\n", encoding="utf-8")
        coupled.write_text("2.0.0\n", encoding="utf-8")
        command("add", "--all")
        command("commit", "-m", "Propose synthetic coupled update")
        update = command("rev-parse", "HEAD")
        changed = command("rev-parse", "HEAD^{tree}") != tree
        command("revert", "--no-edit", update)
        restored = command("rev-parse", "HEAD^{tree}") == tree
        forward = command("rev-parse", "HEAD^") == update
        forward = forward and command("rev-parse", "HEAD") != baseline
    return [
        {"id": name, "status": "pass" if ok else "fail", "detail": ""}
        for name, ok in (
            ("dependency-rollback-update-changes-tree", changed),
            ("dependency-rollback-restores-complete-tree", restored),
            ("dependency-rollback-preserves-forward-history", forward),
        )
    ]


def reporter_tests(module: ModuleType) -> list[dict[str, Any]]:
    finding = module.finding("README.md", "Pipe | must be escaped", 7)
    failed = module.rule_result("sample", "Sample", [finding], "")
    report = {
        "schema_version": 1,
        "tool": "fixture",
        "repository": "example/repo",
        "commit": "a" * 40,
        "mode": "generated",
        "profile": "library",
        "scope_note": "fixture scope",
        "status": "fail",
        "counts": {"rules": 1, "passed": 0, "failed": 1, "findings": 1},
        "rules": [failed],
    }
    first_markdown = module.markdown_report(report)
    second_markdown = module.markdown_report(report)
    first_console = module.console_report(report)
    second_console = module.console_report(report)
    return [
        {
            "id": "markdown-determinism",
            "status": "pass" if first_markdown == second_markdown else "fail",
            "detail": "",
        },
        {
            "id": "console-determinism",
            "status": "pass" if first_console == second_console else "fail",
            "detail": "",
        },
        {
            "id": "markdown-escaping",
            "status": "pass" if r"Pipe \| must be escaped" in first_markdown else "fail",
            "detail": first_markdown if r"Pipe \| must be escaped" not in first_markdown else "",
        },
        {
            "id": "finding-location",
            "status": "pass" if "README.md:7" in first_console else "fail",
            "detail": first_console if "README.md:7" not in first_console else "",
        },
    ]


def cli_tests(module: ModuleType, checker: Path) -> list[dict[str, Any]]:
    parsed = module.parse_arguments(
        ["--root", ".", "--output", "out.json", "--summary", "out.md"]
    )
    flags_ok = (
        parsed.root == Path(".")
        and parsed.output == Path("out.json")
        and parsed.summary == Path("out.md")
        and parsed.self_test is False
    )
    with tempfile.TemporaryDirectory(prefix="checker-contract-not-git-") as temporary:
        completed = subprocess.run(
            [sys.executable, str(checker), "--root", temporary],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
    return [
        {
            "id": "cli-flags",
            "status": "pass" if flags_ok else "fail",
            "detail": repr(parsed) if not flags_ok else "",
        },
        {
            "id": "operational-exit-2",
            "status": "pass" if completed.returncode == 2 else "fail",
            "detail": f"observed exit code {completed.returncode}",
        },
    ]


def _invoke_gate(
    options: argparse.Namespace, **keywords: Any
) -> tuple[int | None, str, str, BaseException | None]:
    """Run ``run_gate`` with captured streams, reporting propagated exceptions."""
    out, err = io.StringIO(), io.StringIO()
    code: int | None = None
    raised: BaseException | None = None
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        try:
            code = run_gate(options, **keywords)
        except BaseException as error:
            raised = error
    return code, out.getvalue(), err.getvalue(), raised


def gate_runner_tests() -> list[dict[str, Any]]:
    """Exercise the shared focused-gate runner's CLI, report and exit contract."""
    results: list[dict[str, Any]] = []

    def record(identifier: str, ok: bool, detail: str = "") -> None:
        results.append(
            {"id": identifier, "status": "pass" if ok else "fail", "detail": "" if ok else detail}
        )

    def options(self_test: bool = False, **paths: Any) -> argparse.Namespace:
        return argparse.Namespace(
            self_test=self_test, output=paths.get("output"), summary=paths.get("summary")
        )

    def markdown(report: dict[str, Any]) -> str:
        return f"## gate\n\n- **Status:** {report['status']}\n"

    def passing() -> dict[str, Any]:
        return {"status": "pass", "findings": []}

    def failing() -> dict[str, Any]:
        return {"status": "fail", "findings": ["one"]}

    def operational() -> dict[str, Any]:
        raise OSError("build failed")

    def programming_error() -> dict[str, Any]:
        raise ZeroDivisionError("division by zero")

    def self_test_ok() -> int:
        return 0

    def self_test_broken() -> int:
        raise OSError("self-test failed")

    parsed = parse_arguments(
        ["--root", ".", "--output", "o.json", "--summary", "o.md", "--self-test"]
    )
    record(
        "gate-cli-flags",
        parsed.root == Path(".")
        and parsed.output == Path("o.json")
        and parsed.summary == Path("o.md")
        and parsed.self_test is True,
        repr(parsed),
    )
    defaults = parse_arguments([])
    record(
        "gate-cli-defaults",
        defaults.output is None and defaults.summary is None and defaults.self_test is False,
        repr(defaults),
    )
    help_code: object = None
    with contextlib.redirect_stdout(io.StringIO()):
        try:
            parse_arguments(["--help"])
        except SystemExit as exit_error:
            help_code = exit_error.code
    record("gate-cli-help", help_code == 0, f"observed {help_code!r}")

    with tempfile.TemporaryDirectory(prefix="gate-runner-evidence-") as temporary:
        output = Path(temporary) / "nested" / "report.json"
        summary = Path(temporary) / "nested" / "report.md"
        code, out, err, raised = _invoke_gate(
            options(output=output, summary=summary),
            build=passing,
            markdown=markdown,
            errors=(OSError,),
        )
        record("gate-pass-exit", code == 0 and raised is None, f"exit {code!r} stderr {err!r}")
        record(
            "gate-pass-stdout", out == markdown(passing()).rstrip() + "\n", repr(out)
        )
        record(
            "gate-pass-json",
            output.read_text(encoding="utf-8")
            == json.dumps(passing(), indent=2, sort_keys=True) + "\n",
            "JSON evidence does not match the canonical serialization",
        )
        record(
            "gate-pass-markdown",
            summary.read_text(encoding="utf-8") == markdown(passing()),
            "Markdown evidence does not match the rendered summary",
        )
        first = (output.read_bytes(), summary.read_bytes())
        _invoke_gate(
            options(output=output, summary=summary),
            build=passing,
            markdown=markdown,
            errors=(OSError,),
        )
        record(
            "gate-deterministic-evidence",
            (output.read_bytes(), summary.read_bytes()) == first,
            "evidence bytes changed between identical runs",
        )

    code, out, err, raised = _invoke_gate(
        options(), build=failing, markdown=markdown, errors=(OSError,)
    )
    record("gate-fail-exit", code == 1 and raised is None, f"exit {code!r}")

    code, out, err, raised = _invoke_gate(
        options(), build=operational, markdown=markdown, errors=(OSError,)
    )
    record(
        "gate-operational-exit",
        code == 2 and err.strip() == "ERROR: build failed",
        f"exit {code!r} stderr {err!r}",
    )

    code, out, err, raised = _invoke_gate(
        options(), build=programming_error, markdown=markdown, errors=(OSError,)
    )
    record(
        "gate-programming-error-propagates",
        code is None and isinstance(raised, ZeroDivisionError),
        f"exit {code!r} raised {raised!r}",
    )

    with tempfile.TemporaryDirectory(prefix="gate-runner-write-") as temporary:
        blocker = Path(temporary) / "blocked"
        blocker.write_text("not a directory\n", encoding="utf-8")
        code, out, err, raised = _invoke_gate(
            options(output=blocker / "report.json"),
            build=passing,
            markdown=markdown,
            errors=(OSError,),
        )
        record(
            "gate-write-failure-exit",
            code == 2 and err.startswith("ERROR: ") and raised is None,
            f"exit {code!r} stderr {err!r}",
        )

    code, out, err, raised = _invoke_gate(
        options(self_test=True),
        build=programming_error,
        markdown=markdown,
        errors=(OSError,),
        self_test=self_test_ok,
    )
    record(
        "gate-self-test-dispatch",
        code == 0 and raised is None,
        f"exit {code!r} raised {raised!r}",
    )

    code, out, err, raised = _invoke_gate(
        options(self_test=True),
        build=passing,
        markdown=markdown,
        errors=(OSError,),
        self_test=self_test_broken,
    )
    record(
        "gate-self-test-default-prefix",
        code == 2 and err.strip() == "ERROR: self-test failed",
        repr(err),
    )

    code, out, err, raised = _invoke_gate(
        options(self_test=True),
        build=passing,
        markdown=markdown,
        errors=(OSError,),
        self_test=self_test_broken,
        self_test_error_prefix="SELF-TEST ERROR",
    )
    record(
        "gate-self-test-isolated-prefix",
        code == 2 and err.strip() == "SELF-TEST ERROR: self-test failed",
        repr(err),
    )

    code, out, err, raised = _invoke_gate(
        options(self_test=True), build=passing, markdown=markdown, errors=(OSError,)
    )
    record("gate-no-self-test-runs-build", code == 0 and raised is None, f"exit {code!r}")

    return results


def check_ids(module: ModuleType) -> tuple[list[str], list[str]]:
    functions = [check.__name__ for check in module.CHECKS]
    failures = []
    if tuple(functions) != EXPECTED_CHECK_FUNCTIONS:
        failures.append("canonical CHECKS order drifted: " + ", ".join(functions))
    return functions, failures


def ownership_scan(root: Path) -> tuple[dict[str, Any], list[str]]:
    """Prove that shared focused-gate helpers have exactly one owner and one consumer set."""
    failures: list[str] = []
    forbidden_helpers = {"git", "run_gate", "tracked_files", "write_text"}
    parser_consumers = {
        "check_local_actions.py",
        "check_release_semantics.py",
        "check_template_contract.py",
        "check_vba_conditionals.py",
        "check_vba_jumps.py",
        "check_vba_public_api.py",
        "checker_development.py",
        "policy_coverage_runner.py",
    }
    local_helper_owners: list[str] = []
    parser_imports: list[str] = []
    gate_runner_imports: list[str] = []
    entry_points: list[str] = []
    for tool in sorted((root / "tools").glob("*.py")):
        if tool.name in {CHECKER_PATH.name, GATELIB_PATH.name}:
            continue
        tool_tree = ast.parse(tool.read_text(encoding="utf-8"), filename=str(tool))
        definitions = {
            node.name
            for node in tool_tree.body
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
        }
        duplicates = sorted(definitions & forbidden_helpers)
        if duplicates:
            local_helper_owners.append(f"{tool.name}: {', '.join(duplicates)}")
        if tool.name in parser_consumers:
            if _imports_helper(tool_tree, "parse_report_args"):
                parser_imports.append(tool.name)
            else:
                failures.append(f"{tool.name} does not consume _gatelib.parse_report_args")
        consumes_runner = _imports_helper(tool_tree, "run_gate")
        if "main" in definitions:
            entry_points.append(tool.name)
        if tool.name in GATE_RUNNER_CONSUMERS:
            gate_runner_imports.append(tool.name)
            if not consumes_runner:
                failures.append(f"{tool.name} does not consume _gatelib.run_gate")
        elif consumes_runner:
            failures.append(
                f"{tool.name} consumes _gatelib.run_gate but is not a declared consumer"
            )
    declared = set(GATE_RUNNER_CONSUMERS) | set(GATE_RUNNER_EXCLUSIONS)
    undeclared = sorted(set(entry_points) - declared)
    if undeclared:
        failures.append(
            "focused-gate entry points are neither run_gate consumers nor documented "
            "exclusions: " + ", ".join(undeclared)
        )
    stale = sorted(declared - set(entry_points))
    if stale:
        failures.append(
            "declared run_gate consumers or exclusions no longer define main: " + ", ".join(stale)
        )
    if local_helper_owners:
        failures.append(
            "shared helpers redefined outside _gatelib.py: " + "; ".join(local_helper_owners)
        )
    evidence = {
        "parser_consumers": parser_imports,
        "gate_runner_consumers": gate_runner_imports,
        "gate_runner_exclusions": [
            {"tool": name, "reason": reason}
            for name, reason in sorted(GATE_RUNNER_EXCLUSIONS.items())
        ],
        "entry_points": sorted(entry_points),
    }
    return evidence, failures


def _imports_helper(tree: ast.Module, name: str) -> bool:
    return any(
        isinstance(node, ast.ImportFrom)
        and node.module == "_gatelib"
        and any(alias.name == name for alias in node.names)
        for node in tree.body
    )


def self_test_declaration_failures(
    advertised: dict[str, bool], exclusions: dict[str, str]
) -> list[str]:
    """Require every discovered CLI to advertise the flag or explain its absence."""
    failures = []
    for name, supported in sorted(advertised.items()):
        if supported and name in exclusions:
            failures.append(f"{name} advertises --self-test but remains excluded")
        elif not supported and not exclusions.get(name, "").strip():
            failures.append(f"{name} lacks --self-test and a documented exclusion")
    for name in sorted(set(exclusions) - set(advertised)):
        failures.append(f"self-test exclusion no longer names a CLI: {name}")
    return failures


def guard_truth_values(node: ast.expr, *, main_mode: bool) -> set[bool]:
    """Model guard polarity without evaluating code; unknown operands stay unknown."""
    if isinstance(node, ast.Constant):
        return {bool(node.value)}
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.Not):
        return {not value for value in guard_truth_values(node.operand, main_mode=main_mode)}
    if isinstance(node, ast.BoolOp):
        conjunction = isinstance(node.op, ast.And)
        result = {conjunction}
        for operand in node.values:
            values = guard_truth_values(operand, main_mode=main_mode)
            result = {left and right if conjunction else left or right
                      for left in result for right in values}
        return result
    if isinstance(node, ast.Compare) and len(node.ops) > 1:
        # Python chains are conjunctions of adjacent comparisons. This only models
        # truth possibilities; it never evaluates operands or repeats side effects.
        operands = [node.left, *node.comparators]
        comparisons: list[ast.expr] = [ast.Compare(left=left, ops=[operator], comparators=[right])
                       for left, operator, right in zip(operands, node.ops, operands[1:])]
        return guard_truth_values(ast.BoolOp(op=ast.And(), values=comparisons),
                                  main_mode=main_mode)
    if isinstance(node, ast.Compare) and len(node.ops) == 1:
        left, right = node.left, node.comparators[0]
        if isinstance(left, ast.Constant) and isinstance(right, ast.Constant):
            if isinstance(node.ops[0], ast.Eq):
                return {left.value == right.value}
            if isinstance(node.ops[0], ast.NotEq):
                return {left.value != right.value}
        for name, value in ((left, right), (right, left)):
            if (isinstance(name, ast.Name) and name.id == "__name__"
                    and isinstance(value, ast.Constant) and value.value == "__main__"):
                if isinstance(node.ops[0], ast.Eq):
                    return {main_mode}
                if isinstance(node.ops[0], ast.NotEq):
                    return {not main_mode}
    return {False, True}


def has_cli_entry_point(tree: ast.Module) -> bool:
    """Include executable main guards regardless of the called function's name."""
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name == "main":
            return True
        if isinstance(node, ast.ImportFrom) and any(
            (alias.asname or alias.name) == "main" for alias in node.names
        ):
            return True
        if not isinstance(node, ast.If):
            continue
        main_values = guard_truth_values(node.test, main_mode=True)
        import_values = guard_truth_values(node.test, main_mode=False)
        # Only branches that can execute as a script and cannot execute on import.
        if True in main_values and True not in import_values:
            return True
        if node.orelse and False in main_values and False not in import_values:
            return True
    return False


def self_test_interface_report(root: Path) -> tuple[list[dict[str, Any]], list[str]]:
    """Inspect live CLI help, including wrappers; never execute operational modes."""
    rows = []
    failures = []
    advertised = {}
    for path in sorted((root / "tools").glob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        if not has_cli_entry_point(tree):
            continue
        completed = subprocess.run(
            [sys.executable, str(path.resolve()), "--help"],
            capture_output=True, text=True, timeout=15,
        )
        if completed.returncode != 0:
            failures.append(f"{path.name} --help failed: exit {completed.returncode}")
        supported = "--self-test" in completed.stdout.split()
        advertised[path.name] = supported
        rows.append({"tool": path.name, "self_test_flag": supported,
                     "exclusion": SELF_TEST_EXCLUSIONS.get(path.name)})
    failures.extend(self_test_declaration_failures(advertised, SELF_TEST_EXCLUSIONS))
    return rows, failures


def cli_discovery_tests() -> list[dict[str, Any]]:
    """Exercise discovery through the actual help audit and its rejection path."""
    bodies = {
        "named-cli": "def cli():\n    parser.parse_args()\nif __name__ == '__main__':\n    cli()\n",
        "qualified-main": (
            "import types\nrunner = types.SimpleNamespace(main=parser.parse_args)\n"
            "if __name__ == '__main__':\n    runner.main()\n"
        ),
        "reversed-guard": "if '__main__' == __name__:\n    parser.parse_args()\n",
        "direct-guard": "if __name__ == '__main__':\n    parser.parse_args()\n",
    }
    results = []
    with tempfile.TemporaryDirectory(prefix="cli-discovery-") as temporary:
        root = Path(temporary)
        (root / "tools").mkdir()
        path = root / "tools/new_gate.py"
        for name, body in bodies.items():
            path.write_text("import argparse\nparser = argparse.ArgumentParser()\n" + body,
                            encoding="utf-8")
            rows, failures = self_test_interface_report(root)
            rejected = any(row["tool"] == path.name for row in rows) and (
                "new_gate.py lacks --self-test and a documented exclusion" in failures
            )
            results.append({"id": "cli-discovery-" + name,
                            "status": "pass" if rejected else "fail", "detail": ""})
    helper = ast.parse("def helper():\n    return 1\n")
    results.append({"id": "cli-discovery-helper-excluded",
                    "status": "pass" if not has_cli_entry_point(helper) else "fail", "detail": ""})
    return results


def guard_polarity_tests() -> list[dict[str, Any]]:
    """Keep import-only helpers out of the actual CLI help audit."""
    cases = (
        ("positive", '__name__ == "__main__"', True),
        ("chain-positive", '"__main__" == __name__ == "__main__"', True),
        ("chain-literal-tail", '__name__ == "__main__" == "__main__"', True),
        ("chain-literal-head", '"__main__" == "__main__" == __name__', True),
        ("chain-negated", 'not ("__main__" == __name__ == "__main__")', False),
        ("chain-contradiction", '"__main__" == __name__ != "__main__"', False),
        ("chain-literal-false", '__name__ == "__main__" == "other"', False),
        ("chain-literal-false-head", '"other" == "__main__" == __name__', False),
        ("chain-unknown-tail", '"__main__" == __name__ == enabled', True),
        ("chain-unknown-head", 'enabled == __name__ == "__main__"', True),
        ("chain-false-and", 'False and "__main__" == __name__ == "__main__"', False),
        ("negated", 'not (__name__ == "__main__")', False),
        ("not-equal", '__name__ != "__main__"', False),
        ("double-negated", 'not (__name__ != "__main__")', True),
        ("false-and", 'False and __name__ == "__main__"', False),
        ("and-false", '__name__ == "__main__" and False', False),
        ("false-or", 'False or __name__ == "__main__"', True),
        ("true-or", 'True or __name__ == "__main__"', False),
        ("unknown-and", 'enabled and __name__ == "__main__"', True),
        ("unknown-or", 'enabled or __name__ == "__main__"', False),
        ("negated-or", 'not (__name__ == "__main__" or enabled)', False),
        ("contradiction", '__name__ == "__main__" and __name__ != "__main__"', False),
        ("literal-main-string", '"__main__"', False),
    )
    results = []
    with tempfile.TemporaryDirectory(prefix="guard-polarity-") as temporary:
        root = Path(temporary)
        (root / "tools").mkdir()
        path = root / "tools/guarded.py"
        for name, condition, expected in cases:
            path.write_text("import argparse\nenabled = True\nif " + condition +
                            ":\n    argparse.ArgumentParser().parse_args()\n", encoding="utf-8")
            rows, failures = self_test_interface_report(root)
            discovered = any(row["tool"] == path.name for row in rows)
            rejected = "guarded.py lacks --self-test and a documented exclusion" in failures
            results.append({"id": "guard-polarity-" + name,
                            "status": "pass" if discovered == rejected == expected else "fail",
                            "detail": ""})
    alternate = ast.parse('if __name__ != "__main__":\n    pass\nelse:\n    cli()\n')
    results.append({"id": "guard-polarity-main-else",
                    "status": "pass" if has_cli_entry_point(alternate) else "fail", "detail": ""})
    return results


def self_test_registry_tests() -> list[dict[str, Any]]:
    cases: tuple[tuple[str, dict[str, bool], dict[str, str], bool], ...] = (
        ("declared-flag", {"gate.py": True}, {}, False),
        ("documented-alternative", {"gate.py": False}, {"gate.py": "Separate suite"}, False),
        ("new-uncovered-cli", {"new.py": False}, {}, True),
        ("removed-flag", {"gate.py": False}, {}, True),
        ("empty-exclusion", {"gate.py": False}, {"gate.py": " "}, True),
        ("stale-exclusion", {}, {"old.py": "Separate suite"}, True),
        ("excluded-flag-added", {"gate.py": True}, {"gate.py": "Separate suite"}, True),
    )
    return [
        {"id": "self-test-registry-" + name,
         "status": "pass" if bool(self_test_declaration_failures(flags, exclusions)) == rejects else "fail",
         "detail": ""}
        for name, flags, exclusions, rejects in cases
    ]


def shared_library_report(root: Path) -> tuple[dict[str, Any], list[str]]:
    failures: list[str] = []
    gatelib = (root / GATELIB_PATH).resolve()
    source = gatelib.read_text(encoding="utf-8")
    tree = ast.parse(source, filename=str(gatelib))
    import_roots: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            import_roots.update(alias.name.split(".", 1)[0] for alias in node.names)
        elif isinstance(node, ast.ImportFrom):
            if node.module is not None:
                import_roots.add(node.module.split(".", 1)[0])
    non_stdlib = sorted(
        root_name
        for root_name in import_roots
        if root_name != "__future__" and root_name not in sys.stdlib_module_names
    )
    if non_stdlib:
        failures.append("_gatelib.py has non-stdlib imports: " + ", ".join(non_stdlib))

    checker_tree = ast.parse((root / CHECKER_PATH).read_text(encoding="utf-8"))
    for node in ast.walk(checker_tree):
        if isinstance(node, ast.ImportFrom) and node.module == "_gatelib":
            failures.append("check_repo.py must remain independent of _gatelib.py")
        elif isinstance(node, ast.Import) and any(alias.name == "_gatelib" for alias in node.names):
            failures.append("check_repo.py must remain independent of _gatelib.py")

    ownership, ownership_failures = ownership_scan(root)
    failures.extend(ownership_failures)
    evidence = {
        "path": GATELIB_PATH.as_posix(),
        "imports": sorted(import_roots),
        "check_repo_independent": not any("check_repo.py must remain" in item for item in failures),
        **ownership,
    }
    return evidence, failures


def build_report(root: Path) -> dict[str, Any]:
    checker = (root / CHECKER_PATH).resolve()
    source = checker.read_text(encoding="utf-8")
    tree = ast.parse(source, filename=str(checker))
    module = load_checker(checker)

    sections, failures = section_report(source, tree)
    imports, import_failures = import_report(tree)
    failures.extend(import_failures)
    ids, id_failures = check_ids(module)
    failures.extend(id_failures)
    shared_library, shared_failures = shared_library_report(root)
    failures.extend(shared_failures)
    interfaces, interface_failures = self_test_interface_report(root)
    failures.extend(interface_failures)

    parser_results = (parser_tests(module) + reusable_identity_tests(module)
                      + dependency_rollback_tests())
    reporter_results = reporter_tests(module)
    cli_results = cli_tests(module, checker)
    gate_results = (gate_runner_tests() + self_test_registry_tests()
                    + cli_discovery_tests() + guard_polarity_tests())
    all_unit_results = [*parser_results, *reporter_results, *cli_results, *gate_results]
    failed_units = [item for item in all_unit_results if item["status"] != "pass"]
    if failed_units:
        failures.append(f"{len(failed_units)} independent parser/reporter/CLI test(s) failed")

    digest = hashlib.sha256(checker.read_bytes()).hexdigest()
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": "pass" if not failures else "fail",
        "artifact": {
            "path": CHECKER_PATH.as_posix(),
            "development_model": "reviewed-source-is-distributable",
            "build_transform": "none",
            "sha256": digest,
            "runtime_dependencies": "python-standard-library-only",
        },
        "sections": sections,
        "imports": imports,
        "canonical_checks": ids,
        "shared_library": shared_library,
        "self_test_interfaces": interfaces,
        "unit_tests": all_unit_results,
        "failures": failures,
    }


def markdown_report(report: dict[str, Any]) -> str:
    artifact = report["artifact"]
    lines = [
        "## Checker development contract",
        "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **Runtime artifact:** `{artifact['path']}`",
        f"- **Development model:** `{artifact['development_model']}`",
        f"- **Build transform:** `{artifact['build_transform']}`",
        f"- **SHA-256:** `{artifact['sha256']}`",
        f"- **Runtime dependencies:** `{artifact['runtime_dependencies']}`",
        f"- **Internal sections:** {len(report['sections'])}",
        f"- **Canonical policy checks:** {len(report['canonical_checks'])}",
        f"- **Shared focused-gate library:** `{report['shared_library']['path']}`",
        f"- **Independent unit tests:** {len(report['unit_tests'])}",
        "",
        "| Section | Start | End | Definitions |",
        "| --- | --- | --- | ---: |",
    ]
    for section in report["sections"]:
        lines.append(
            f"| `{section['id']}` | `{section['start']}` | `{section['end']}` | "
            f"{section['definitions']} |"
        )
    lines.extend(["", "### Independent tests", "", "| Test | Result |", "| --- | --- |"])
    for item in report["unit_tests"]:
        lines.append(f"| `{item['id']}` | {str(item['status']).upper()} |")
    lines.extend(["", "### Self-test interfaces", "",
                  "CLI help is checked here; fixture execution remains a separate CI responsibility.",
                  "", "| CLI | --self-test | Alternative / exclusion |", "| --- | --- | --- |"])
    for item in report["self_test_interfaces"]:
        lines.append(f"| `{item['tool']}` | {'yes' if item['self_test_flag'] else 'no'} | "
                     f"{item['exclusion'] or 'Dedicated flag'} |")
    if report["failures"]:
        lines.extend(["", "### Failures", ""])
        lines.extend(f"- {item}" for item in report["failures"])
    return "\n".join(lines).rstrip() + "\n"




def run_self_test(root: Path) -> int:
    first = build_report(root)
    second = build_report(root)
    failures: list[str] = []
    if first["status"] != "pass":
        failures.append("development contract is not satisfied")
    if json.dumps(first, sort_keys=True) != json.dumps(second, sort_keys=True):
        failures.append("JSON evidence is not deterministic")
    if markdown_report(first) != markdown_report(second):
        failures.append("Markdown evidence is not deterministic")
    if failures:
        for item in failures:
            print(f"[FAIL] {item}")
        print(markdown_report(first).rstrip())
        return 1
    print(
        "SELF-TEST PASS: internal boundaries, parser/reporter units, CLI contract, "
        "canonical check order, artifact identity, shared-helper ownership, and standard-library-only runtime passed."
    )
    return 0




def main(arguments: list[str] | None = None) -> int:
    options = parse_arguments(sys.argv[1:] if arguments is None else arguments)
    return run_gate(
        options,
        build=lambda: build_report(options.root),
        markdown=markdown_report,
        errors=(
            OSError,
            UnicodeError,
            SyntaxError,
            ContractError,
            subprocess.SubprocessError,
        ),
        self_test=lambda: run_self_test(options.root),
    )

if __name__ == "__main__":
    raise SystemExit(main())
