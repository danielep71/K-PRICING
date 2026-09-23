#!/usr/bin/env python3
"""Plan exact GitHub settings changes; apply only a reviewed digest with a durable journal."""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import re
import tempfile
import time
from pathlib import Path
from typing import Any
from urllib.parse import quote
from urllib.request import Request, build_opener

from _gatelib import run_gate
from check_portfolio_drift import CONFIG, CORE_LABEL_NAMES, PROFILES, require
from check_template_contract import CONTRACT_RULE_SETS
from collect_portfolio_snapshot import NoRedirect

POLICY = ".github/provisioning-policy.json"
FEATURES = {"has_issues", "has_wiki", "has_projects", "has_discussions", "is_template",
            "allow_merge_commit", "allow_squash_merge", "allow_rebase_merge", "allow_auto_merge",
            "delete_branch_on_merge", "allow_update_branch"}
METHODS = {"allow_merge_commit", "allow_squash_merge", "allow_rebase_merge"}
FILES = (CONFIG, ".github/initialization.json", POLICY, ".github/labels.json")


def digest(value: Any) -> str:
    return hashlib.sha256(json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False).encode()).hexdigest()


class GitHub:
    def __init__(self, repository: str, token: str | None, apply: bool = False):
        require(bool(re.fullmatch(r"[A-Za-z0-9_-][A-Za-z0-9_.-]*/[A-Za-z0-9_-][A-Za-z0-9_.-]*", repository)), "Invalid repository")
        self.repository, self.token, self.apply = repository, token, apply

    def request(self, method: str, path: str = "", body: Any = None) -> Any:
        require(path == "" or path.startswith("/") and not path.startswith("//") and ".." not in path,
                "Invalid repository-relative API path")
        allowed_write = method == "PATCH" and (path == "" or path.startswith("/labels/")) or \
            method == "PUT" and path == "/topics" or method == "POST" and path in {"/labels", "/rulesets"}
        require(method == "GET" or self.apply and bool(self.token) and allowed_write, "Write outside approved API surface")
        headers = {"Accept": "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28"}
        if self.token:
            headers["Authorization"] = "Bearer " + self.token
        request = Request(f"https://api.github.com/repos/{self.repository}{path}",
                          data=None if body is None else json.dumps(body).encode(), headers=headers, method=method)
        with build_opener(NoRedirect).open(request, timeout=30) as response:
            data = response.read(10_000_001)
        require(len(data) <= 10_000_000, "Response size limit")
        return json.loads(data) if data else None


def collection(api: Any, path: str) -> list[Any]:
    rows = []
    for page in range(1, 101):
        values = api.request("GET", f"{path}?per_page=100&page={page}")
        require(isinstance(values, list), "Invalid collection")
        rows.extend(values)
        if len(values) < 100:
            return rows
    raise ValueError("Incomplete collection; refusing partial planning")


def source_json(api: Any, path: str, sha: str) -> Any:
    blob = api.request("GET", f"/contents/{path}?ref={sha}")
    require(blob.get("encoding") == "base64" and blob.get("type") == "file", "Expected source JSON file")
    return json.loads(base64.b64decode(blob["content"]).decode("utf-8"))


def capture(api: Any, sha: str) -> dict[str, Any]:
    meta = api.request("GET")
    require(meta.get("full_name", "").lower() == api.repository.lower(), "Repository identity changed")
    require(not meta.get("archived") and not meta.get("disabled"), "Repository is not writable")
    branch = meta["default_branch"]
    head = api.request("GET", "/commits/" + quote(branch, safe=""))["sha"]
    require(head == sha, "Default branch moved or source SHA is not the initialized default-branch candidate")
    documents = {path: source_json(api, path, sha) for path in FILES}
    topics = api.request("GET", "/topics")["names"]
    labels = [{"name": row["name"], "color": row["color"].upper(), "description": row.get("description") or ""}
              for row in collection(api, "/labels")]
    rulesets = [api.request("GET", f"/rulesets/{row['id']}") for row in collection(api, "/rulesets")]
    require(all(isinstance(row.get("bypass_actors"), list) for row in rulesets), "Ruleset bypass evidence unavailable")
    rulesets = [{key: row.get(key) for key in ("id", "name", "target", "enforcement", "bypass_actors", "conditions", "rules")}
                for row in rulesets]
    return {"repository": api.repository, "repository_id": meta["id"], "source_sha": sha,
            "default_branch": branch, "metadata": {key: meta.get(key) for key in sorted(FEATURES | {"description"})},
            "topics": sorted(topics), "labels": sorted(labels, key=lambda row: row["name"].lower()),
            "rulesets": sorted(rulesets, key=lambda row: row["id"]), "source": documents}


def desired_labels(manifest: dict[str, Any], config: dict[str, Any]) -> list[dict[str, Any]]:
    require(manifest.get("schema_version") == 1, "Unsupported label schema")
    require(CORE_LABEL_NAMES <= {label["name"] for label in manifest["core"]}, "Universal core labels removed")
    overlays = manifest["overlays"]
    labels = list(manifest["core"]) + overlays["profile"][config["profile"]]
    for domain in config.get("label_domains", []):
        labels += overlays["domain"][domain]
    names = set()
    for label in labels:
        require(set(label) == {"name", "color", "description"}, "Invalid label fields")
        name = label["name"]
        require(isinstance(name, str) and 0 < len(name) <= 50 and name == name.strip() and "\n" not in name,
                "Invalid label name")
        require(name.lower() not in names, "Duplicate selected label")
        names.add(name.lower())
        require(bool(re.fullmatch(r"[0-9A-F]{6}", label["color"])), "Invalid label color")
        require(isinstance(label["description"], str) and 0 < len(label["description"]) <= 100, "Invalid label description")
    require(bool(labels), "Empty label baseline")
    return sorted(labels, key=lambda row: row["name"].lower())


def validate(before: dict[str, Any], profile: str, version: str) -> dict[str, Any]:
    require(profile in PROFILES and version in CONTRACT_RULE_SETS, "Unsupported contract/profile")
    config, record, policy = (before["source"][path] for path in FILES[:3])
    require(config.get("mode") == "generated", "Only explicitly initialized generated repositories can be provisioned")
    require(config.get("repository") == before["repository"] and
            record.get("values", {}).get("REPOSITORY_PATH") == before["repository"], "Recorded repository differs from explicit target")
    for document in (config, record):
        require(document.get("profile") == profile,
                "Recorded repository/profile differs from explicit target")
        require(document.get("template_contract", {}).get("version") == version, "Recorded contract differs from requested version")
    require(config["template_contract"] == record["template_contract"], "Initialization contract mismatch")
    require(policy.get("schema_version") == 1 and policy.get("contract_version") == version == "1.2.0",
            "No provisioning policy adapter for this contract version")
    require(set(policy) == {"schema_version", "contract_version", "description", "topics", "features", "required_checks", "exceptions"},
            "Unsupported provisioning policy fields")
    policy = dict(policy)
    if policy["description"] == "@initialization:PROJECT_DESCRIPTION":
        policy["description"] = record["values"].get("PROJECT_DESCRIPTION")
    require(isinstance(policy["description"], str) and bool(policy["description"].strip()) and "{{" not in policy["description"],
            "Resolve the project description before provisioning")
    require(set(policy["features"]) == FEATURES and all(type(value) is bool for value in policy["features"].values()), "Invalid feature policy")
    require(policy["features"]["has_issues"] and not policy["features"]["allow_auto_merge"], "Universal issue/merge controls cannot be waived")
    require(any(policy["features"][key] for key in METHODS), "At least one merge method required")
    require(isinstance(policy["topics"], list) and all(isinstance(topic, str) and re.fullmatch(r"[a-z0-9][a-z0-9-]{0,49}", topic)
            for topic in policy["topics"]), "Invalid topics")
    require(isinstance(policy["required_checks"], list) and bool(policy["required_checks"]) and
            all(isinstance(name, str) and bool(name.strip()) for name in policy["required_checks"]), "Invalid required checks")
    require("Repository integrity" in policy["required_checks"], "Universal Repository integrity check cannot be removed")
    require(isinstance(policy["exceptions"], dict) and all(isinstance(reason, str) and bool(reason.strip())
            for reason in policy["exceptions"].values()), "Exceptions need reviewed rationale in source policy")
    return policy


def baseline_rules(policy: dict[str, Any]) -> list[dict[str, Any]]:
    common = [{"type": "deletion"}, {"type": "non_fast_forward"}]
    branch = {"name": "Template default branch", "target": "branch", "enforcement": "active", "bypass_actors": [],
              "conditions": {"ref_name": {"include": ["~DEFAULT_BRANCH"], "exclude": []}},
              "rules": common + [{"type": "pull_request", "parameters": {"required_approving_review_count": 0,
                  "dismiss_stale_reviews_on_push": False, "require_code_owner_review": False,
                  "require_last_push_approval": False, "required_review_thread_resolution": False}},
                  {"type": "required_status_checks", "parameters": {"strict_required_status_checks_policy": True,
                      "required_status_checks": [{"context": name} for name in sorted(set(policy["required_checks"]))]}}]}
    tag = {"name": "Template version tags", "target": "tag", "enforcement": "active", "bypass_actors": [],
           "conditions": {"ref_name": {"include": ["refs/tags/v*"], "exclude": []}},
           "rules": common + [{"type": "update", "parameters": {"update_allows_fetch_and_merge": False}}]}
    return [branch, tag]


def rule_covers(actual: dict[str, Any], desired: dict[str, Any]) -> bool:
    if actual.get("type") != desired["type"]:
        return False
    expected, params = desired.get("parameters", {}), actual.get("parameters", {})
    if desired["type"] == "required_status_checks":
        names = {row["context"] for row in params.get("required_status_checks", [])}
        return params.get("strict_required_status_checks_policy") is True and \
            {row["context"] for row in expected["required_status_checks"]} <= names
    if desired["type"] == "pull_request":
        return params.get("required_approving_review_count", -1) >= expected["required_approving_review_count"]
    if desired["type"] == "update" and expected == {"update_allows_fetch_and_merge": False}:
        # GitHub can omit the false update parameter when reading a ruleset back.
        return params.get("update_allows_fetch_and_merge", False) is False
    return all(params.get(key) == value for key, value in expected.items())


def covered(existing: list[dict[str, Any]], desired: dict[str, Any], branch: str) -> bool:
    accepted = {"~ALL", "~DEFAULT_BRANCH", f"refs/heads/{branch}"} if desired["target"] == "branch" else {"~ALL", "refs/tags/v*"}
    rules = []
    for item in existing:
        scope = (item.get("conditions") or {}).get("ref_name", {})
        if item.get("target") == desired["target"] and item.get("enforcement") == "active" and \
                not item.get("bypass_actors") and not scope.get("exclude") and accepted.intersection(scope.get("include", [])):
            rules += item.get("rules", [])
    return all(any(rule_covers(actual, rule) for actual in rules) for rule in desired["rules"])


def make_plan(before: dict[str, Any], profile: str, version: str) -> dict[str, Any]:
    policy = validate(before, profile, version)
    actions, kept, blocked = [], [], []
    exceptions = policy["exceptions"]
    metadata = {}
    for key, value in {"description": policy["description"], **policy["features"]}.items():
        old = before["metadata"][key]
        require(old is not None or key == "description", "Incomplete metadata; cannot plan " + key)
        if old == value:
            continue
        if key in METHODS and old is False and value is True:
            kept.append(f"Narrower existing merge method retained: {key}")
        elif key in {"has_wiki", "has_projects", "has_discussions", "is_template"} and old is True and value is False and "feature:" + key not in exceptions:
            blocked.append(f"Disabling {key} needs a recorded feature:{key} exception")
        else:
            metadata[key] = value
    if metadata:
        actions.append({"method": "PATCH", "path": "", "body": metadata})
    topics = sorted(set(before["topics"]) | set(policy["topics"]) | {"excel", "vba", profile})
    require(len(topics) <= 20, "Topic union exceeds GitHub limit; no local topics removed")
    if topics != before["topics"]:
        actions.append({"method": "PUT", "path": "/topics", "body": {"names": topics}})
    wanted = desired_labels(before["source"][FILES[3]], before["source"][CONFIG])
    live = {row["name"].lower(): row for row in before["labels"]}
    for label in wanted:
        old = live.get(label["name"].lower())
        if old == label:
            continue
        if old is not None and "label:" + old["name"] not in exceptions:
            blocked.append(f"Existing label differs; needs label:{old['name']} exception")
        else:
            actions.append({"method": "POST" if old is None else "PATCH",
                            "path": "/labels" if old is None else "/labels/" + quote(old["name"], safe=""),
                            "body": label if old is None else {"new_name": label["name"], "color": label["color"], "description": label["description"]}})
    kept += ["Extra label retained: " + row["name"] for key, row in live.items() if key not in {row["name"].lower() for row in wanted}]
    for desired in baseline_rules(policy):
        if covered(before["rulesets"], desired, before["default_branch"]):
            kept.append("Existing equal/stronger rules retained: " + desired["target"])
        elif any(row["name"] == desired["name"] for row in before["rulesets"]):
            blocked.append("Existing named ruleset differs; manual reviewed migration required: " + desired["name"])
        else:
            actions.append({"method": "POST", "path": "/rulesets", "body": desired})
    result = {"repository": before["repository"], "profile": profile, "contract_version": version,
              "source_sha": before["source_sha"], "before": before, "actions": actions,
              "keep": sorted(kept), "blocked": sorted(blocked), "exceptions": exceptions}
    return {**result, "plan_sha256": digest(result)}


def journal(path: Path, report: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, name = tempfile.mkstemp(dir=path.parent, prefix=".provision-")
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            json.dump(report, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.flush()
            os.fsync(stream.fileno())
        for attempt in range(6):
            try:
                os.replace(name, path)
                break
            except PermissionError:
                if attempt == 5:
                    raise
                time.sleep(0.02 * (2 ** attempt))
    finally:
        if os.path.exists(name):
            os.unlink(name)


def execute(api: Any, sha: str, profile: str, version: str, approval: str | None = None, receipt: Path | None = None) -> dict[str, Any]:
    before = capture(api, sha)
    plan = make_plan(before, profile, version)
    report: dict[str, Any] = {"status": "fail" if plan["blocked"] else "pass", "mode": "plan", "plan": plan,
                              "attempts": [], "verification": None, "scope": "Live settings only; not source, Excel or release certification"}
    if approval is None:
        return report
    if receipt is None:
        raise ValueError("Apply requires a journal path")
    require(not receipt.exists(), "Use a new journal path; earlier apply evidence must be retained")
    require(approval == plan["plan_sha256"], "Approved plan no longer matches live state or source policy")
    require(not plan["blocked"], "Blocked changes require policy review; no mutation performed")
    report.update(mode="apply", status="fail")
    journal(receipt, report)
    expected = before
    try:
        for action in plan["actions"]:
            current = capture(api, sha)
            require(digest(current) == digest(expected), "Concurrent change detected; stopped before next write")
            attempt = {"action": action, "outcome": "REQUEST_PENDING_OR_UNCERTAIN"}
            report["attempts"].append(attempt)
            journal(receipt, report)
            api.request(action["method"], action["path"], action["body"])
            attempt["outcome"] = "REQUEST_RETURNED"
            expected = capture(api, sha)
            journal(receipt, report)
        final = capture(api, sha)
        remaining = make_plan(final, profile, version)
        report["verification"] = {"after": final, "remaining_actions": remaining["actions"], "blocked": remaining["blocked"]}
        report["status"] = "fail" if remaining["actions"] or remaining["blocked"] else "pass"
    except (OSError, ValueError, KeyError, TypeError) as error:
        report["error"] = type(error).__name__ + "; inspect journal and collect a new plan; no automatic write retry or rollback"
    journal(receipt, report)
    return report


def markdown_report(report: dict[str, Any]) -> str:
    plan = report["plan"]
    lines = ["# Repository provisioning", "", f"Mode: {report['mode']} · Result: {report['status']}", "",
             report["scope"], "", f"Repository: {plan['repository']} · Profile: {plan['profile']} · Contract: {plan['contract_version']}", "",
             f"Source: `{plan['source_sha']}`", "", f"Plan digest: `{plan['plan_sha256']}`", "",
             "## Exact proposed requests", "", "```json", json.dumps(plan["actions"], indent=2, sort_keys=True), "```", "",
             "## Preserved controls", "", *["- " + item for item in plan["keep"]], "",
             "## Blockers", "", *["- " + item for item in plan["blocked"]], "",
             "## Apply and verification", "", "```json", json.dumps({key: report.get(key) for key in ("attempts", "verification", "error")}, indent=2, sort_keys=True), "```", ""]
    return "\n".join(lines)


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repository", required=True)
    parser.add_argument("--profile", required=True, choices=sorted(PROFILES))
    parser.add_argument("--contract-version", required=True)
    parser.add_argument("--source-sha", required=True)
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--approve-plan")
    parser.add_argument("--journal", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    options = parser.parse_args(arguments)
    if not re.fullmatch(r"[0-9a-f]{40}", options.source_sha):
        parser.error("source-sha must be an exact commit")
    if options.apply and (not options.approve_plan or not options.journal or not os.environ.get("GH_TOKEN")):
        parser.error("Apply requires --approve-plan, --journal and trusted GH_TOKEN")
    if not options.apply and (options.approve_plan or options.journal):
        parser.error("Approval and journal are apply-only options")
    paths = [path.resolve() for path in (options.output, options.summary, options.journal) if path is not None]
    if len(paths) != len(set(paths)):
        parser.error("Report and journal paths must differ")
    api = GitHub(options.repository, os.environ.get("GH_TOKEN"), options.apply)
    return run_gate(options, build=lambda: execute(api, options.source_sha, options.profile, options.contract_version,
                    options.approve_plan if options.apply else None, options.journal), markdown=markdown_report,
                    errors=(OSError, ValueError, KeyError, TypeError))


if __name__ == "__main__":
    raise SystemExit(main())
