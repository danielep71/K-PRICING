#!/usr/bin/env python3
"""Render timestamped portfolio evidence; never execute inspected code or use the network."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote

import check_portfolio_drift as drift
from _gatelib import run_gate


def instant(value: Any) -> datetime:
    drift.require(isinstance(value, str), "Timestamp must be an ISO-8601 string")
    result = datetime.fromisoformat(value.replace("Z", "+00:00"))
    drift.require(result.tzinfo is not None, "Timestamp needs a timezone")
    return result.astimezone(timezone.utc)


def freshness(observed: Any, as_of: datetime, days: int) -> str:
    if observed is None:
        return "UNVERIFIED"
    captured = instant(observed)
    if captured > as_of:
        return "UNVERIFIED"
    return "STALE" if as_of - captured > timedelta(days=days) else "CURRENT"


def required_checks(repo: dict[str, Any]) -> list[str] | None:
    """Resolve demonstrably applicable rules only; unknown scopes cannot yield PASS."""
    if repo["rulesets"] is None or repo["metadata"] is None:
        return None
    branch = repo["metadata"].get("default_branch")
    if not branch:
        return None
    accepted = {"~ALL", "~DEFAULT_BRANCH", f"refs/heads/{branch}"}
    contexts = set()
    for item in repo["rulesets"]:
        if item.get("target") != "branch" or item.get("enforcement") != "active":
            continue
        scope = item.get("conditions", {}).get("ref_name", {})
        if scope.get("exclude") or any(any(char in ref for char in "*?[") and ref not in accepted for ref in scope.get("include", [])):
            return None
        if not accepted.intersection(scope.get("include", [])):
            continue
        for rule in item.get("rules", []):
            if rule.get("type") != "required_status_checks":
                continue
            for check in rule.get("parameters", {}).get("required_status_checks", []):
                if check.get("integration_id") is not None:
                    return None  # Actions job evidence does not authenticate another check provider.
                name = check.get("context")
                drift.require(isinstance(name, str) and bool(name), "Invalid required check context")
                contexts.add(name)
    return sorted(contexts)


def workflow_health(repo: dict[str, Any], quality: dict[str, Any]) -> tuple[str, str, list[str]]:
    required = required_checks(repo)
    workflows = quality.get("workflows")
    if required is None or workflows is None:
        return "UNVERIFIED", "Required-check policy or Actions evidence unavailable", []
    if not required:
        return "UNVERIFIED", "No demonstrable required-check set; unrelated green runs do not suffice", []
    branch = (repo["metadata"] or {}).get("default_branch")
    latest: dict[str, Any] = {}
    identities = set()
    for run in workflows:
        identity = (run["id"], run.get("run_attempt", 1))
        drift.require(identity not in identities, "Duplicate workflow attempt evidence")
        identities.add(identity)
        if run.get("head_sha") != repo["commit"] or run.get("head_branch") != branch or run.get("event") not in {"push", "workflow_dispatch"}:
            continue
        key = run["path"]
        if key not in latest or (run["id"], run.get("run_attempt", 1)) > (latest[key]["id"], latest[key].get("run_attempt", 1)):
            latest[key] = run
    states, links = [], []
    for context in required:
        matches = [(run, job) for run in latest.values() for job in run.get("jobs", []) if job.get("name") == context]
        if len(matches) != 1:
            states.append("UNVERIFIED")
            continue
        run, job = matches[0]
        links.append(f"https://github.com/{repo['repository']}/actions/runs/{run['id']}/attempts/{run.get('run_attempt', 1)}")
        if run.get("status") != "completed" or job.get("status") != "completed":
            states.append("UNVERIFIED")
        elif job.get("conclusion") == "success" and run.get("conclusion") == "success":
            states.append("PASS")
        else:
            states.append("FAIL")
    state = "FAIL" if "FAIL" in states else "UNVERIFIED" if "UNVERIFIED" in states else "PASS"
    return state, "Exact-SHA required Actions jobs: " + ", ".join(required), sorted(set(links))


def release_state(repo: dict[str, Any], quality: dict[str, Any]) -> tuple[str, str, list[str]]:
    releases = quality.get("releases")
    if releases is None:
        return "UNVERIFIED", "Release collection unavailable", []
    published = [item for item in releases if item.get("draft") is False and item.get("prerelease") is False]
    if not published:
        return "UNRELEASED", "No published stable release observed; no certification claim", []
    latest = max(published, key=lambda item: (instant(item["published_at"]), item["id"]))
    if quality.get("observed_at") is None or instant(latest["published_at"]) > instant(quality["observed_at"]):
        return "UNVERIFIED", "Release publication postdates evidence capture", []
    tag = latest["tag_name"]
    link = f"https://github.com/{repo['repository']}/releases/tag/{quote(tag, safe='')}"
    sha = latest.get("resolved_commit")
    if not isinstance(sha, str) or not re.fullmatch(r"[0-9a-f]{40}", sha):
        return "UNVERIFIED", f"Published {tag}; tag commit unresolved; certification not assessed", [link]
    return "OBSERVED", f"Published {tag} at {sha}; publication is not certification", [link, f"https://github.com/{repo['repository']}/tree/{sha}"]


def specialist_claims(repo: dict[str, Any], as_of: datetime, days: int) -> list[dict[str, Any]]:
    claims = []
    for item in repo.get("specialist", []):
        drift.require(isinstance(item, dict), "Invalid specialist claim")
        for key in ("name", "score", "scale", "rationale", "evidence_path", "commit"):
            drift.require(key in item, "Incomplete specialist claim")
        path = item["evidence_path"]
        drift.require(isinstance(path, str) and path in repo["files"], "Specialist evidence must be captured")
        drift.require(isinstance(item["score"], (int, float, str)) and not isinstance(item["score"], bool), "Invalid specialist score")
        state = freshness(item.get("observed_at"), as_of, days)
        if item["commit"] != repo["commit"]:
            state = "UNVERIFIED"
        claims.append({**item, "status": "REPORTED" if state == "CURRENT" else state,
                       "evidence": f"https://github.com/{repo['repository']}/blob/{repo['commit']}/{quote(path, safe='/')}",
                       "scope": "Attributed specialist claim; not independently certified or included in an aggregate score"})
    return sorted(claims, key=lambda item: (item["name"], str(item["score"])))


def repository_report(repo: dict[str, Any], findings: list[dict[str, Any]], as_of: datetime, days: int) -> dict[str, Any]:
    quality = repo.get("quality", {})
    drift.require(isinstance(quality, dict), "Invalid quality evidence")
    for key in ("workflows", "releases"):
        value = quality.get(key)
        drift.require(value is None or isinstance(value, list) and all(isinstance(item, dict) for item in value),
                      f"Invalid quality {key}")
    for run in quality.get("workflows") or []:
        drift.require(isinstance(run.get("jobs"), list) and all(isinstance(job, dict) for job in run["jobs"]),
                      "Invalid workflow job evidence")
    observed = quality.get("observed_at")
    current = freshness(observed, as_of, days)
    evidence = f"https://github.com/{repo['repository']}/tree/{repo['commit']}"
    dimensions = []

    def add(name: str, state: str, detail: str, links: list[str] | None = None) -> None:
        dimensions.append({"dimension": name, "status": current if current != "CURRENT" else state,
                           "observed_status": state, "detail": detail, "evidence": links or [evidence]})

    config = drift.read_json(repo, drift.CONFIG) or {}
    adoption = next(row for row in findings if row["rule"] == "adoption")
    add("adoption", "ADOPT" if adoption["decision"] == "ADOPT" else adoption["status"], adoption["detail"])
    structural = [row for row in findings if row["rule"] not in {"adoption", "branch-protection", "release-tag-protection"}]
    states = {row["status"] for row in structural}
    conformance = "DRIFT" if "DRIFT" in states else "UNVERIFIED" if not states or "UNVERIFIED" in states else "PASS"
    add("structural-conformance", conformance, "Only #25 structural predicates; no runtime or implementation certification")
    for rule in ("branch-protection", "release-tag-protection"):
        state, detail = drift.protection(repo, "branch" if rule == "branch-protection" else "tag")
        add(rule, state, detail, [f"https://github.com/{repo['repository']}/rules"])
    add("required-workflow-health", *workflow_health(repo, quality))
    add("latest-stable-release", *release_state(repo, quality))
    rendered = []
    for row in findings:
        rendered.append({**row, "observed_status": row["status"],
                         "status": current if current != "CURRENT" else row["status"]})
    return {"repository": repo["repository"], "commit": repo["commit"], "observed_at": observed,
            "freshness": current, "recorded_contract": config.get("template_contract"),
            "recorded_profile": config.get("profile"), "dimensions": dimensions,
            "drift_findings": rendered, "specialist": specialist_claims(repo, as_of, days),
            "workflow_observations": quality.get("workflows"),
            "release_evidence": quality.get("releases"), "certification": "NOT_ASSESSED",
            "unavailable": repo.get("unavailable", {})}


def build_report(snapshot: dict[str, Any], as_of: str, max_age_days: int = 7) -> dict[str, Any]:
    drift.require(max_age_days > 0, "Maximum age must be positive")
    now = instant(as_of)
    raw = drift.build_report(snapshot)
    rows = [repository_report(repo, [row for row in raw["findings"] if row["repository"] == repo["repository"]], now, max_age_days)
            for repo in sorted(snapshot["repositories"], key=lambda item: item["repository"])]
    acceptable = {"PASS", "OBSERVED", "UNRELEASED", "NOT_APPLICABLE"}
    complete = all(dim["status"] in acceptable for repo in rows for dim in repo["dimensions"])
    return {"schema_version": 1, "status": "pass" if complete else "fail",
            "scope": "Evidence coverage and control observations, not a quality ranking or release certification",
            "as_of": now.isoformat(), "max_age_days": max_age_days,
            "snapshot_sha256": hashlib.sha256(json.dumps(snapshot, sort_keys=True, separators=(",", ":")).encode()).hexdigest(),
            "repositories": rows}


def markdown_report(report: dict[str, Any]) -> str:
    def cell(value: Any) -> str:
        return str(value).replace("|", "\\|").replace("\n", " ").replace("\r", " ")
    lines = ["# Portfolio quality and conformance", "", report["scope"], "",
             f"As of: {report['as_of']} · Maximum snapshot age: {report['max_age_days']} days", "",
             f"Snapshot SHA-256: `{report['snapshot_sha256']}`", ""]
    for repo in report["repositories"]:
        lines += [f"## {cell(repo['repository'])}", "",
                  f"Commit: `{repo['commit']}` · Observed: {cell(repo['observed_at'])} · Freshness: {repo['freshness']}", "",
                  f"Recorded contract: {cell(repo['recorded_contract'])} · Profile: {cell(repo['recorded_profile'])}", "",
                  "| Dimension | Status | Evidence / scope |", "| --- | --- | --- |"]
        for item in repo["dimensions"]:
            links = ", ".join(f"[evidence {index + 1}]({link})" for index, link in enumerate(item["evidence"]))
            lines.append(f"| {item['dimension']} | {item['status']} | {links} — {cell(item['detail'])} |")
        lines += ["", "### Decisions and structural findings", "",
                  "| Rule | Decision | Status | Rationale / evidence |", "| --- | --- | --- | --- |"]
        for row in repo["drift_findings"]:
            lines.append(f"| [{row['rule']}]({row['contract_rule']}) | {row['decision']} | {row['status']} | "
                         f"[source]({row['evidence']}) — {cell(row['reason'] or row['detail'])} |")
        for claim in repo["specialist"]:
            lines += ["", f"Specialist: {cell(claim['name'])} — {cell(claim['score'])} ({cell(claim['scale'])}), "
                      f"{claim['status']}; [source]({claim['evidence']}). {cell(claim['rationale'])}. {claim['scope']}."]
        lines += ["", "Release certification: NOT ASSESSED.", ""]
    return "\n".join(lines)


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--snapshot", type=Path, required=True)
    parser.add_argument("--as-of", required=True)
    parser.add_argument("--max-age-days", type=int, default=7)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    options = parser.parse_args(arguments)
    paths = [path.resolve() for path in (options.snapshot, options.output, options.summary) if path is not None]
    if len(paths) != len(set(paths)):
        parser.error("Input, JSON and Markdown paths must differ")
    return run_gate(options, build=lambda: build_report(json.loads(options.snapshot.read_text(encoding="utf-8")), options.as_of, options.max_age_days),
                    markdown=markdown_report, errors=(OSError, ValueError, TypeError, KeyError))


if __name__ == "__main__":
    raise SystemExit(main())
