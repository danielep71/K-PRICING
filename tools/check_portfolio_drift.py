#!/usr/bin/env python3
"""Evaluate captured repository evidence without network access or code execution."""
from __future__ import annotations

import argparse
import hashlib
import importlib
import json
import re
from pathlib import Path, PurePosixPath
from typing import Any

from _gatelib import run_gate
from check_template_contract import CONTRACT_RULE_SETS

CONFIG = ".github/repository-profile.json"
PROFILES = {"application", "library", "ui-component"}
DECISIONS = {"REQUIRED", "ADOPT", "KEEP", "DEFER", "NOT APPLICABLE"}
CORE_LABEL_NAMES = {"behavior-change", "blocked", "bug", "ci", "documentation", "duplicate",
                    "enhancement", "good first issue", "help wanted", "invalid", "P1", "P2", "P3",
                    "question", "refactor", "release", "repository", "security", "tests", "wontfix"}
PATHS = {
    "canonical-repository-gate": ["README.md", "CONTRIBUTING.md", "SECURITY.md", "LICENSE", "tools/check_repo.py"],
    "deterministic-initializer": ["tools/initialize_repository.py"],
    "label-policy": [".github/labels.json"],
    "placeholder-schema": [CONFIG],
    "profile-model": [CONFIG],
    "release-integrity": ["RELEASING.md", "VERSION", "CHANGELOG.md", ".github/release-policy.json"],
    "committed-whitespace": ["tools/check_committed_whitespace.py"],
    "complete-public-api": ["docs/PUBLIC_API.txt", "tools/check_vba_public_api.py"],
    "label-drift-detection": [".github/workflows/labels-drift.yml"],
    "nested-conditional-compilation": ["tools/check_vba_conditionals.py"],
    "procedure-scoped-jumps": ["tools/check_vba_jumps.py"],
    "repository-local-actions": ["tools/check_local_actions.py"],
    "strict-release-semantics": ["tools/check_release_semantics.py"],
    "template-contract-version": ["docs/TEMPLATE_CONTRACT.md", "tools/check_template_contract.py"],
    "controlled-dependency-updates": ["docs/DEPENDENCY_UPDATES.md"],
    "documentation-drift": [".github/documentation-policy.json", "docs/DOCUMENTATION_CHECKS.md",
                            "tools/check_documentation.py", "tools/check_external_links.py"],
    "advanced-release-provenance": [".github/release-provenance.json", "docs/RELEASE_PROVENANCE.md",
                                    "tools/release_provenance.py", "tools/test_release_provenance.py"],
}
UNIVERSAL = {"metadata", "workflow-properties", "branch-protection", "release-tag-protection"}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def valid_path(path: str) -> bool:
    item = PurePosixPath(path)
    return bool(path) and not item.is_absolute() and ".." not in item.parts and str(item) == path


def snapshot_shape(repo: dict[str, Any]) -> None:
    require(bool(re.fullmatch(r"[\w.-]+/[\w.-]+", repo.get("repository", ""))), "Invalid repository")
    require(bool(re.fullmatch(r"[0-9a-f]{40}", repo.get("commit", ""))), "Invalid commit")
    paths = repo.get("paths")
    require(paths is None or (isinstance(paths, list) and all(isinstance(p, str) and valid_path(p)
            for p in paths) and len(paths) == len(set(paths))), "paths must be a complete unique tree or null")
    files = repo.get("files")
    if not isinstance(files, dict):
        raise ValueError("files must be an object")
    require(all(valid_path(p) and isinstance(v, str) for p, v in files.items()), "Invalid file content")
    require(paths is None or set(files).issubset(paths), "Content absent from captured tree")
    for key, kind in (("metadata", dict), ("labels", list), ("rulesets", list)):
        require(key in repo and (repo[key] is None or isinstance(repo[key], kind)), f"Invalid {key}")


def read_json(repo: dict[str, Any], path: str) -> Any:
    text = repo["files"].get(path)
    return json.loads(text) if text is not None else None


def observation(ok: bool | None, detail: str) -> tuple[str, str]:
    return ("UNVERIFIED" if ok is None else "PASS" if ok else "DRIFT", detail)


def missing_paths(repo: dict[str, Any], paths: list[str]) -> tuple[str, str]:
    if repo["paths"] is None:
        return observation(None, "Complete tree unavailable")
    missing = sorted(set(paths) - set(repo["paths"]))
    return observation(not missing, "Missing: " + ", ".join(missing) if missing else "Required paths present")


def yaml_document(text: str) -> dict[str, Any]:
    """Walk composed mappings/sequences as data, retaining scalar text without constructors.

    Reject aliases, merge keys, duplicate keys and custom tags; bound size and depth.
    """
    yaml = importlib.import_module("yaml")
    require(yaml.__version__ == "6.0.3", "Install PyYAML==6.0.3")
    require(len(text) <= 1_000_000, "Workflow exceeds parser limit")
    require(not any(isinstance(event, yaml.AliasEvent) for event in yaml.parse(text)),
            "YAML aliases require manual review")

    def value(node: Any, depth: int = 0) -> Any:
        require(depth < 40, "YAML nesting exceeds parser limit")
        require(node is not None and node.tag.startswith("tag:yaml.org,2002:"), "Custom/empty YAML node")
        if isinstance(node, yaml.ScalarNode):
            require(node.tag.startswith("tag:yaml.org,2002:"), "Custom YAML tag")
            return node.value
        if isinstance(node, yaml.SequenceNode):
            return [value(child, depth + 1) for child in node.value]
        require(isinstance(node, yaml.MappingNode), "Invalid YAML node")
        result: dict[str, Any] = {}
        for key, child in node.value:
            name = value(key, depth + 1)
            require(isinstance(name, str) and name not in result and name != "<<", "Duplicate/merge YAML key")
            result[name] = value(child, depth + 1)
        return result

    result = value(yaml.compose(text))
    require(isinstance(result, dict), "Workflow is not an object")
    return result


def workflow_errors(document: dict[str, Any]) -> list[str]:
    errors = []
    permissions = document.get("permissions")
    if permissions != {"contents": "read"}:
        errors.append("Workflow must explicitly use contents: read only")
    jobs = document.get("jobs")
    if not isinstance(jobs, dict) or not jobs:
        raise ValueError("Workflow jobs unavailable")
    for name, job in jobs.items():
        require(isinstance(job, dict), "Invalid workflow job")
        if job.get("permissions", permissions) != {"contents": "read"}:
            errors.append(f"{name}: elevated or unclear permissions")
        if "secrets" in job or "environment" in job:
            errors.append(f"{name}: secrets or environment require separate trusted workflow")
        references = [job] if "uses" in job else job.get("steps", [])
        require(isinstance(references, list) and bool(references), "Missing/invalid workflow steps")
        for step in references:
            require(isinstance(step, dict), "Invalid workflow step")
            require(bool(step.get("uses")) != bool(step.get("run")), "Step needs exactly one run or uses")
            ref = step.get("uses")
            if ref and not re.fullmatch(r"[\w.-]+/[\w./-]+@[0-9a-f]{40}", ref):
                errors.append(f"{name}: non-immutable or local reference needs independent inspection")
    return errors


def workflow_check(repo: dict[str, Any], policy: dict[str, Any]) -> tuple[str, str]:
    path = policy.get("workflow", ".github/workflows/static-checks.yml")
    state, detail = missing_paths(repo, [path])
    if state != "PASS":
        return state, detail
    if path not in repo["files"]:
        return observation(None, "Workflow source was not captured")
    try:
        yaml = importlib.import_module("yaml")
    except ImportError:
        return observation(None, "Install PyYAML==6.0.3 to inspect workflow evidence")
    try:
        document = yaml_document(repo["files"][path])
        errors = workflow_errors(document)
    except (ValueError, TypeError, RecursionError, yaml.YAMLError) as error:
        return observation(None, str(error))
    return observation(not errors, "; ".join(errors) or "Scoped generic workflow has immutable refs and read-only permissions; execution unproven")


def protection(repo: dict[str, Any], target: str) -> tuple[str, str]:
    rulesets = repo["rulesets"]
    if rulesets is None:
        return observation(None, "Rulesets unavailable; absence of evidence is not protection")
    branch = (repo["metadata"] or {}).get("default_branch")
    if target == "branch" and not branch:
        return observation(None, "Default branch unavailable")
    accepted = {"~ALL", "~DEFAULT_BRANCH", f"refs/heads/{branch}"} if target == "branch" else {"~ALL", "refs/tags/v*"}
    applicable = []
    for item in rulesets:
        require(isinstance(item, dict), "Invalid ruleset")
        if item.get("target") != target or item.get("enforcement") != "active":
            continue
        scope = item.get("conditions", {}).get("ref_name", {})
        if not set(scope.get("include", [])) & accepted:
            continue
        if scope.get("exclude"):
            return observation(None, "Ruleset exclusions require manual scope review")
        applicable.append(item)
    if not applicable:
        return observation(False, "No active ruleset demonstrably covers required refs")
    if any(item.get("bypass_actors") for item in applicable):
        return observation(False, "Applicable protection permits bypass")
    rules = [rule for item in applicable for rule in item.get("rules", [])]
    types = {rule.get("type") for rule in rules}
    needed = {"deletion", "non_fast_forward"}
    if target == "branch":
        needed.add("pull_request")
        checks = [rule.get("parameters", {}) for rule in rules if rule.get("type") == "required_status_checks"]
        if not any(item.get("strict_required_status_checks_policy") is True and item.get("required_status_checks") for item in checks):
            return observation(False, "No strict required quality check")
    else:
        needed.add("update")
    return observation(needed <= types, "Required protection types: " + ", ".join(sorted(needed)))


def label_check(repo: dict[str, Any]) -> tuple[str, str]:
    manifest = read_json(repo, ".github/labels.json")
    if manifest is None or repo["labels"] is None:
        return observation(None, "Label manifest or live label snapshot unavailable")
    require(isinstance(manifest, dict) and isinstance(manifest.get("core"), list), "Invalid label manifest")
    live = {item["name"]: item for item in repo["labels"]}
    core = manifest["core"]
    if not CORE_LABEL_NAMES <= {item.get("name") for item in core}:
        return observation(False, "Required baseline core labels removed from declared policy")
    bad = [item["name"] for item in core if any(live.get(item["name"], {}).get(field) != item.get(field)
           for field in ("name", "color", "description"))]
    return observation(not bad, "Core label differences: " + ", ".join(sorted(bad)) if bad else "Declared core labels match; additional local labels retained")


def semantic_check(repo: dict[str, Any], rule: str, config: dict[str, Any], policy: dict[str, Any]) -> tuple[str, str]:
    if rule == "metadata":
        meta = repo["metadata"]
        return observation(None if meta is None else bool(meta.get("description")) and
                           meta.get("has_issues") is True and meta.get("allow_auto_merge") is False,
                           "Description, issue intake and manual merging required")
    if rule == "workflow-properties":
        return workflow_check(repo, policy)
    if rule in {"branch-protection", "release-tag-protection"}:
        return protection(repo, "branch" if rule == "branch-protection" else "tag")
    if rule == "label-policy":
        return label_check(repo)
    if rule == "profile-model":
        return observation(config.get("mode") == "generated" and config.get("profile") in PROFILES
                           and config.get("repository") == repo["repository"], "Recorded generated profile and repository identity")
    if rule == "placeholder-schema":
        return observation(isinstance(config.get("placeholders"), dict) and bool(config["placeholders"].get("catalogue")), "Placeholder catalogue required")
    if rule == "release-integrity":
        release = read_json(repo, ".github/release-policy.json")
        if release is None:
            return observation(None, "Release policy contents unavailable")
        return observation({"repository-integrity", "vba-compile", "regression"} <= set(release.get("core_checks", [])), "Universal release evidence must remain required")
    return observation(True, "Required control paths present; implementation correctness is outside structural drift scope")


def decisions_for(repo: dict[str, Any]) -> dict[str, Any]:
    result = {}
    for entry in repo.get("decisions", []):
        require(isinstance(entry, dict), "Invalid decision")
        rule = entry.get("rule")
        require(isinstance(rule, str) and rule not in result, "Duplicate or missing decision rule")
        require(entry.get("decision") in DECISIONS, "Unknown decision")
        require(isinstance(entry.get("reason"), str) and bool(entry["reason"].strip()), "Decision rationale missing")
        require(entry.get("commit") == repo["commit"], "Decision bound to another candidate")
        evidence = entry.get("evidence_path")
        require(isinstance(evidence, str) and evidence in repo["files"], "Decision evidence must be captured")
        result[rule] = entry
    return result


def assess(repo: dict[str, Any], source: str, contract_commit: str) -> list[dict[str, Any]]:
    snapshot_shape(repo)
    config = read_json(repo, CONFIG)
    decisions = decisions_for(repo)
    require(set(decisions) <= set(PATHS) | UNIVERSAL | {"adoption", "local-specialist", "ui-evidence"},
            "Unknown decision rule")
    rows = []

    def add(rule: str, state: str, detail: str, decision: str = "REQUIRED") -> None:
        entry = decisions.get(rule, {})
        selected = entry.get("decision", decision)
        if selected == "NOT APPLICABLE" and rule in UNIVERSAL | set(PATHS):
            state, detail = "DRIFT", "Universal control cannot be marked NOT APPLICABLE"
        rows.append({"repository": repo["repository"], "commit": repo["commit"],
                     "rule": rule, "decision": selected, "status": state,
                     "detail": detail, "reason": entry.get("reason", ""),
                     "evidence": f"https://github.com/{repo['repository']}/tree/{repo['commit']}",
                     "contract_rule": f"https://github.com/{source}/blob/{contract_commit}/docs/PORTFOLIO_DRIFT.md#{rule}",
                     "decision_evidence": entry.get("evidence_path")})

    if config is None:
        state = "DRIFT" if repo["paths"] is not None and CONFIG not in repo["paths"] else "UNVERIFIED"
        add("adoption", state, "No captured adopted contract; no version or profile assigned", "ADOPT")
        for rule in sorted(UNIVERSAL):
            add(rule, "UNVERIFIED", "No adopted contract; observation cannot be scored as conformance")
        return rows
    require(isinstance(config, dict), "Invalid profile document")
    contract = config.get("template_contract", {})
    version = contract.get("version") if isinstance(contract, dict) else None
    if version not in CONTRACT_RULE_SETS:
        add("adoption", "UNVERIFIED", "Missing/unsupported adopted version; no latest-version fallback", "ADOPT")
        return rows
    require(bool(re.fullmatch(r"[\w.-]+/[\w.-]+", contract.get("source", ""))), "Invalid template source")
    if contract["source"] != source:
        add("adoption", "UNVERIFIED", "Recorded source differs from evaluator contract source", "ADOPT")
        return rows
    add("adoption", "PASS", f"Recorded contract {version}; profile {config.get('profile')}")
    rules = CONTRACT_RULE_SETS[version] | UNIVERSAL
    require(set(decisions) <= rules | {"adoption", "local-specialist", "ui-evidence"}, "Unknown decision rule")
    for rule in sorted(rules):
        if rule not in PATHS and rule not in UNIVERSAL:
            add(rule, "UNVERIFIED", "No evaluator registered for this contract rule")
            continue
        state, detail = missing_paths(repo, PATHS.get(rule, []))
        if state == "PASS":
            state, detail = semantic_check(repo, rule, config, repo.get("policy", {}))
        add(rule, state, detail)
    for rule in ("local-specialist", "ui-evidence"):
        if rule == "ui-evidence" and config.get("profile") != "ui-component":
            add(rule, "NOT_APPLICABLE", "UI-only optional evidence; universal release controls still apply", "NOT APPLICABLE")
        elif rule in decisions:
            add(rule, "UNVERIFIED", "Retain documented specialist control; runtime evidence not certified", "KEEP")
    return rows


def build_report(snapshot: dict[str, Any]) -> dict[str, Any]:
    require(snapshot.get("schema_version") == 1, "Unsupported snapshot schema")
    source = snapshot.get("contract_source")
    if not isinstance(source, str) or not re.fullmatch(r"[\w.-]+/[\w.-]+", source):
        raise ValueError("Missing evaluator contract_source")
    contract_commit = snapshot.get("contract_commit", "")
    require(bool(re.fullmatch(r"[0-9a-f]{40}", contract_commit)), "Invalid contract_commit")
    repos = snapshot.get("repositories")
    if not isinstance(repos, list) or not repos or not all(isinstance(repo, dict) for repo in repos):
        raise ValueError("Empty/invalid portfolio is not a pass")
    names = [repo.get("repository") for repo in repos]
    require(len(names) == len(set(names)), "Duplicate repository snapshot")
    rows = [row for repo in sorted(repos, key=lambda item: item["repository"]) for row in assess(repo, source, contract_commit)]
    digest = hashlib.sha256(json.dumps(snapshot, sort_keys=True, separators=(",", ":")).encode()).hexdigest()
    return {"schema_version": 1, "status": "pass" if all(row["status"] in {"PASS", "NOT_APPLICABLE"} for row in rows) else "fail",
            "snapshot_sha256": digest, "contract_source": source, "findings": rows,
            "scope": "Structural drift from captured evidence; no code execution, remote writes or runtime certification"}


def markdown_report(report: dict[str, Any]) -> str:
    def cell(value: Any) -> str:
        return str(value).replace("|", "\\|").replace("\n", " ").replace("\r", " ")
    lines = ["# Portfolio semantic drift", "", report["scope"], "",
             "| Repository | Rule | Decision | Result | Evidence / finding |",
             "| --- | --- | --- | --- | --- |"]
    for row in report["findings"]:
        annotation = ""
        if row["decision_evidence"]:
            annotation = f"; [decision]({row['evidence']}/{row['decision_evidence']}): {cell(row['reason'])}"
        lines.append(f"| {cell(row['repository'])} | [{cell(row['rule'])}]({row['contract_rule']}) | " +
                     " | ".join(cell(row[key]) for key in ("decision", "status")) +
                     f" | [source]({row['evidence']}) — {cell(row['detail'])}{annotation} |")
    return "\n".join(lines) + "\n"


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--snapshot", type=Path, required=True)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    options = parser.parse_args(arguments)
    if options.output is not None and options.summary is not None and options.output.resolve() == options.summary.resolve():
        parser.error("JSON and Markdown outputs must use different paths")
    for target in (options.output, options.summary):
        if target is not None and target.resolve() == options.snapshot.resolve():
            parser.error("Report must not overwrite the input snapshot")
    return run_gate(options, build=lambda: build_report(json.loads(options.snapshot.read_text(encoding="utf-8"))),
                    markdown=markdown_report, errors=(OSError, ValueError, TypeError, KeyError))


if __name__ == "__main__":
    raise SystemExit(main())
