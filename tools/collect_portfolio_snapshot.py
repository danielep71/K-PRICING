#!/usr/bin/env python3
"""Capture GitHub evidence using GET only; emit private snapshot JSON to stdout.

Supply optional GH_TOKEN with only the required read permissions. The collector
issues GET requests only; it cannot restrict the token's own granted scopes.
No inspected repository is modified.
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from typing import Any

from check_portfolio_drift import CONFIG, PATHS


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req: Any, fp: Any, code: int, msg: str,
                         headers: Any, newurl: str) -> None:
        return None


def get(path: str) -> Any:
    headers = {"Accept": "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28"}
    token = os.environ.get("GH_TOKEN")
    if token:
        headers["Authorization"] = "Bearer " + token
    request = urllib.request.Request("https://api.github.com/repos/" + path, headers=headers, method="GET")
    with urllib.request.build_opener(NoRedirect).open(request, timeout=30) as response:
        data = response.read(10_000_001)
    if len(data) > 10_000_000:
        raise ValueError("Response exceeds snapshot size limit")
    return json.loads(data)


def pages(path: str, key: str | None = None) -> list[Any]:
    result = []
    for page in range(1, 101):
        separator = "&" if "?" in path else "?"
        data = get(f"{path}{separator}per_page=100&page={page}")
        if key is not None:
            if data.get("total_count", 0) >= 1000:
                raise ValueError("Actions search limit; refusing potentially partial evidence")
            data = data[key]
        if not isinstance(data, list):
            raise ValueError("Expected paginated collection")
        result.extend(data)
        if len(data) < 100:
            return result
    raise ValueError("Pagination limit reached; refusing partial evidence")


def tag_commit(repository: str, tag: str) -> str:
    item = get(f"{repository}/git/ref/tags/{urllib.parse.quote(tag, safe='')}")["object"]
    for _ in range(10):
        if item["type"] == "commit":
            return str(item["sha"])
        if item["type"] != "tag":
            raise ValueError("Release tag does not resolve to a commit")
        item = get(f"{repository}/git/tags/{item['sha']}")["object"]
    raise ValueError("Release tag indirection limit")


def quality_evidence(repo: dict[str, Any]) -> dict[str, Any]:
    """Capture workflow attempts and latest stable release; never certify artifacts."""
    repository = repo["repository"]
    quality: dict[str, Any] = {"observed_at": repo.get("captured_at", datetime.now(timezone.utc).isoformat())}
    try:
        runs = pages(f"{repository}/actions/runs?head_sha={repo['commit']}", "workflow_runs")
        selected: dict[tuple[str, str, str], Any] = {}
        for run in runs:
            if run.get("head_sha") != repo["commit"]:
                continue
            key = (run["path"], run["event"], run["head_branch"])
            if key not in selected or run["id"] > selected[key]["id"]:
                selected[key] = run
        captured = []
        for run in sorted(selected.values(), key=lambda item: item["id"]):
            attempt = run.get("run_attempt", 1)
            jobs = pages(f"{repository}/actions/runs/{run['id']}/attempts/{attempt}/jobs", "jobs")
            item = {key: run.get(key) for key in ("id", "path", "head_sha", "head_branch", "event", "status", "conclusion", "run_attempt", "updated_at")}
            item["jobs"] = [{key: job.get(key) for key in ("id", "name", "status", "conclusion")} for job in jobs]
            captured.append(item)
        quality["workflows"] = captured
    except (urllib.error.HTTPError, ValueError) as error:
        quality["workflows"] = None
        repo["unavailable"]["workflows"] = type(error).__name__ + "; complete evidence unavailable"
    try:
        releases = pages(f"{repository}/releases")
        published = [item for item in releases if not item["draft"] and not item["prerelease"]]
        latest = max(published, key=lambda item: (item["published_at"], item["id"])) if published else None
        evidence = []
        if latest is not None:
            item = {key: latest[key] for key in ("id", "tag_name", "draft", "prerelease", "published_at", "body")}
            item["assets"] = [{key: asset.get(key) for key in ("name", "size", "digest", "browser_download_url")} for asset in latest["assets"]]
            try:
                item["resolved_commit"] = tag_commit(repository, item["tag_name"])
            except (urllib.error.HTTPError, ValueError):
                item["resolved_commit"] = None
                repo["unavailable"]["release_tag"] = "Published release tag could not be resolved"
            evidence.append(item)
        quality["releases"] = evidence
    except (urllib.error.HTTPError, ValueError) as error:
        quality["releases"] = None
        repo["unavailable"]["releases"] = type(error).__name__ + "; complete evidence unavailable"
    return quality


def collect(repository: str) -> dict[str, Any]:
    if not re.fullmatch(r"[\w.-]+/[\w.-]+", repository):
        raise ValueError("Repository must be owner/name")
    captured_at = datetime.now(timezone.utc).isoformat()
    metadata = get(repository)
    branch = urllib.parse.quote(metadata["default_branch"], safe="")
    commit = get(f"{repository}/commits/{branch}")["sha"]
    tree = get(f"{repository}/git/trees/{commit}?recursive=1")
    unavailable = {}
    blobs = {item["path"]: item["sha"] for item in tree["tree"] if item["type"] == "blob"}
    if tree.get("truncated"):
        unavailable["paths"] = "Recursive tree truncated"
    required = {path for paths in PATHS.values() for path in paths}
    required.update(path for path in blobs if path.startswith(".github/workflows/"))
    required.add(CONFIG)
    files = {}
    for path in sorted(required & set(blobs)):
        try:
            blob = get(f"{repository}/git/blobs/{blobs[path]}")
            files[path] = base64.b64decode(blob["content"]).decode("utf-8")
        except (urllib.error.HTTPError, UnicodeError, ValueError) as error:
            unavailable[path] = type(error).__name__
    live: dict[str, Any] = {}
    for resource in ("labels", "rulesets"):
        try:
            values = pages(f"{repository}/{resource}")
            live[resource] = ([get(f"{repository}/rulesets/{item['id']}") for item in values]
                              if resource == "rulesets" else values)
        except urllib.error.HTTPError as error:
            live[resource] = None
            unavailable[resource] = f"HTTP {error.code}; not evidence of absence"
    return {"repository": repository, "commit": commit, "captured_at": captured_at,
            "paths": None if tree.get("truncated") else sorted(blobs), "files": files,
            "metadata": {key: metadata.get(key) for key in
                         ("default_branch", "description", "has_issues", "allow_auto_merge")},
            **live, "unavailable": unavailable, "decisions": []}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("repositories", nargs="+")
    parser.add_argument("--contract-commit", required=True, help="Exact template evaluator commit")
    parser.add_argument("--quality", action="store_true", help="Also capture timestamped Actions and release evidence")
    options = parser.parse_args()
    try:
        if len(set(options.repositories)) != len(options.repositories):
            raise ValueError("Duplicate repository")
        if not re.fullmatch(r"[0-9a-f]{40}", options.contract_commit):
            raise ValueError("Invalid evaluator commit")
        data = {"schema_version": 1, "contract_source": "danielep71/EXCEL-VBA-PROJECT-TEMPLATE",
                "contract_commit": options.contract_commit,
                "repositories": [collect(name) for name in sorted(options.repositories)]}
        if options.quality:
            for repo in data["repositories"]:
                repo["quality"] = quality_evidence(repo)
        print(json.dumps(data, indent=2, sort_keys=True))
        return 0
    except (OSError, ValueError, KeyError) as error:
        print(f"Capture failed: {type(error).__name__}; no complete snapshot emitted", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
