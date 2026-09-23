#!/usr/bin/env python3
"""Validate post-release closeout from an exact Git candidate and captured provider facts."""

from __future__ import annotations

import argparse
import fnmatch
import hashlib
import json
import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from _gatelib import git_text

PROFILE_PATH = ".github/repository-profile.json"
RELEASE_POLICY_PATH = ".github/release-policy.json"
VERSION_PATH = "VERSION"
CHANGELOG_PATH = "CHANGELOG.md"
DEFAULT_WORKFLOW = ".github/workflows/static-checks.yml"
VERSION_PATTERN = re.compile(r"(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$")
SHA_PATTERN = re.compile(r"[0-9a-f]{40}$")


class CloseoutError(RuntimeError):
    """Captured closeout evidence cannot be evaluated deterministically."""


@dataclass(frozen=True)
class Options:
    root: Path
    snapshot: Path
    tag: str
    candidate_sha: str
    milestone_number: int
    workflow_path: str
    expect_prerelease: bool
    expect_latest: bool
    output: Path | None = None
    summary: Path | None = None


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CloseoutError(message)


def as_object(value: Any, name: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise CloseoutError(f"{name} must be an object")
    return value


def as_list(value: Any, name: str) -> list[Any]:
    if not isinstance(value, list):
        raise CloseoutError(f"{name} must be an array")
    return value


def as_string(value: Any, name: str) -> str:
    if not isinstance(value, str) or not value:
        raise CloseoutError(f"{name} must be a non-empty string")
    return value


def unique_object(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        require(key not in result, f"duplicate JSON key: {key}")
        result[key] = value
    return result


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=unique_object)
    except FileNotFoundError as error:
        raise CloseoutError(f"evidence file is missing: {path}") from error
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise CloseoutError(f"cannot read {path}: {error}") from error


def git_show(root: Path, candidate: str, relative: str) -> str:
    completed = git_text(root, "show", f"{candidate}:{relative}")
    if completed.returncode:
        detail = completed.stderr.strip() or "git show failed"
        raise CloseoutError(f"cannot read {relative} from {candidate}: {detail}")
    return completed.stdout


def candidate_json(root: Path, candidate: str, relative: str) -> dict[str, Any]:
    try:
        value = json.loads(git_show(root, candidate, relative), object_pairs_hook=unique_object)
    except json.JSONDecodeError as error:
        raise CloseoutError(f"{relative} at {candidate} is invalid JSON: {error}") from error
    return as_object(value, f"{relative} at {candidate}")


def finding(category: str, control: str, message: str) -> dict[str, str]:
    return {"category": category, "control": control, "message": message}


def workflow_runs(value: Any) -> tuple[list[dict[str, Any]], int | None]:
    pages = value if isinstance(value, list) else [value]
    rows: list[dict[str, Any]] = []
    total: int | None = None
    for raw_page in pages:
        page = as_object(raw_page, "workflow_runs page")
        current = as_list(page.get("workflow_runs"), "workflow_runs page.workflow_runs")
        rows.extend(row for row in current if isinstance(row, dict))
        count = page.get("total_count")
        if isinstance(count, int):
            total = count if total is None else max(total, count)
    return rows, total


def milestone_items(value: Any) -> list[dict[str, Any]]:
    raw = as_list(value, "milestone_items")
    if raw and all(isinstance(page, list) for page in raw):
        return [
            row
            for page in raw
            if isinstance(page, list)
            for row in page
            if isinstance(row, dict)
        ]
    return [row for row in raw if isinstance(row, dict)]


def source_identity(
    root: Path,
    candidate: str,
    tag: str,
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    version = git_show(root, candidate, VERSION_PATH).strip()
    if VERSION_PATTERN.fullmatch(version) is None:
        findings.append(
            finding("deterministic", "source-version", f"candidate VERSION is invalid: {version!r}")
        )
    if tag != f"v{version}":
        findings.append(
            finding(
                "deterministic",
                "source-version",
                f"tag {tag!r} disagrees with candidate VERSION {version!r}",
            )
        )

    profile_data = candidate_json(root, candidate, PROFILE_PATH)
    repository = as_string(profile_data.get("repository"), "candidate repository")
    require("/" in repository, "candidate repository must use owner/name form")
    mode_value = profile_data.get("mode")
    mode = as_string(mode_value, "candidate mode")
    if mode == "template":
        profile = "template"
    else:
        profile = as_string(profile_data.get("profile"), "candidate release profile")

    policy = candidate_json(root, candidate, RELEASE_POLICY_PATH)
    profiles = as_object(policy.get("profiles"), "release policy profiles")
    profile_policy = as_object(profiles.get(profile), f"release policy profile {profile!r}")
    raw_allowed = as_list(
        profile_policy.get("allowed_asset_globs"),
        f"release policy profile {profile!r}.allowed_asset_globs",
    )
    if not all(isinstance(item, str) and item for item in raw_allowed):
        raise CloseoutError("allowed_asset_globs must contain only non-empty strings")
    allowed = [str(item) for item in raw_allowed]

    changelog = git_show(root, candidate, CHANGELOG_PATH)
    heading = re.search(
        rf"^## \[{re.escape(version)}\] - (?P<date>\d{{4}}-\d{{2}}-\d{{2}})\s*$",
        changelog,
        re.MULTILINE,
    )
    if heading is None:
        findings.append(
            finding(
                "deterministic",
                "changelog",
                f"candidate changelog has no released section for {version}",
            )
        )

    release_link = re.search(
        rf"^\[{re.escape(version)}\]:\s+https://github\.com/{re.escape(repository)}/compare/"
        rf"(?P<base>v[^\s]+)\.\.\.{re.escape(tag)}\s*$",
        changelog,
        re.MULTILINE,
    )
    previous = release_link.group("base") if release_link else None
    if release_link is None:
        findings.append(
            finding(
                "deterministic",
                "comparison-link",
                f"candidate comparison link for {version} does not target {tag}",
            )
        )

    unreleased = re.search(
        rf"^\[Unreleased\]:\s+https://github\.com/{re.escape(repository)}/compare/"
        rf"{re.escape(tag)}\.\.\.HEAD\s*$",
        changelog,
        re.MULTILINE,
    )
    if unreleased is None:
        findings.append(
            finding(
                "deterministic",
                "comparison-link",
                f"Unreleased comparison link does not start from {tag}",
            )
        )

    return {
        "repository": repository,
        "mode": mode,
        "profile": profile,
        "version": version,
        "changelog_date": heading.group("date") if heading else None,
        "previous_tag": previous,
        "allowed_asset_globs": allowed,
    }


def check_tag(
    snapshot: dict[str, Any],
    tag: str,
    candidate: str,
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    ref = as_object(snapshot.get("tag_ref"), "snapshot.tag_ref")
    if ref.get("ref") != f"refs/tags/{tag}":
        findings.append(finding("deterministic", "tag", "provider tag ref does not match release tag"))
    obj = as_object(ref.get("object"), "snapshot.tag_ref.object")
    ref_type = obj.get("type")
    target: str | None = None
    tag_object_sha: str | None = None
    if ref_type == "commit":
        target = obj.get("sha") if isinstance(obj.get("sha"), str) else None
        findings.append(finding("deterministic", "tag", f"{tag} is lightweight; annotated tag required"))
    elif ref_type == "tag":
        tag_object_sha = obj.get("sha") if isinstance(obj.get("sha"), str) else None
        annotated = as_object(snapshot.get("tag_object"), "snapshot.tag_object")
        if annotated.get("sha") != tag_object_sha:
            findings.append(finding("deterministic", "tag", "annotated tag object SHA mismatches tag ref"))
        if annotated.get("tag") != tag:
            findings.append(finding("deterministic", "tag", "annotated tag object names another tag"))
        annotated_target = as_object(annotated.get("object"), "snapshot.tag_object.object")
        if annotated_target.get("type") != "commit":
            findings.append(finding("deterministic", "tag", "annotated tag does not target a commit"))
        target = (
            annotated_target.get("sha")
            if isinstance(annotated_target.get("sha"), str)
            else None
        )
    else:
        findings.append(finding("deterministic", "tag", f"unsupported tag object type: {ref_type!r}"))
    if target != candidate:
        findings.append(
            finding(
                "deterministic",
                "tag",
                f"tag resolves to {target!r}; expected certified candidate {candidate}",
            )
        )
    return {"ref_type": ref_type, "tag_object_sha": tag_object_sha, "target_sha": target}


def check_workflow(
    snapshot: dict[str, Any],
    path: str,
    tag: str,
    candidate: str,
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    rows, total = workflow_runs(snapshot.get("workflow_runs"))
    if total is not None and total > len(rows):
        findings.append(
            finding(
                "deterministic",
                "tag-ci",
                f"workflow capture is incomplete: {len(rows)} of {total} runs",
            )
        )
    matches = [
        row
        for row in rows
        if row.get("path") == path
        and row.get("head_branch") == tag
        and row.get("head_sha") == candidate
        and row.get("event") == "push"
    ]
    if not matches:
        findings.append(
            finding("deterministic", "tag-ci", f"no tag-triggered {path} run matches candidate")
        )
        return {"matches": 0, "run_id": None, "status": None, "conclusion": None}
    selected = max(
        matches,
        key=lambda row: (
            int(row.get("run_attempt") or 0),
            int(row.get("run_number") or 0),
            int(row.get("id") or 0),
        ),
    )
    status = selected.get("status")
    conclusion = selected.get("conclusion")
    if status != "completed" or conclusion != "success":
        findings.append(
            finding("deterministic", "tag-ci", f"tag CI is {status!r}/{conclusion!r}")
        )
    return {
        "matches": len(matches),
        "run_id": selected.get("id"),
        "status": status,
        "conclusion": conclusion,
    }


def check_release(
    snapshot: dict[str, Any],
    identity: dict[str, Any],
    expect_prerelease: bool,
    expect_latest: bool,
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    release = as_object(snapshot.get("release"), "snapshot.release")
    latest_value = snapshot.get("latest_release")
    if latest_value is None:
        require(
            not expect_latest,
            "snapshot.latest_release is required when latest status is expected",
        )
        latest: dict[str, Any] = {}
    else:
        latest = as_object(latest_value, "snapshot.latest_release")
    tag = f"v{identity['version']}"
    if release.get("tag_name") != tag:
        findings.append(finding("deterministic", "release", "GitHub Release uses another tag"))
    if release.get("draft") is not False or not release.get("published_at"):
        findings.append(finding("deterministic", "release", "GitHub Release is draft/unpublished"))
    prerelease = bool(release.get("prerelease"))
    if prerelease != expect_prerelease:
        findings.append(
            finding(
                "deterministic",
                "release",
                f"prerelease flag is {prerelease}; expected {expect_prerelease}",
            )
        )
    latest_matches = release.get("id") == latest.get("id")
    if latest_matches != expect_latest:
        findings.append(
            finding(
                "deterministic",
                "release",
                f"latest-release expectation is {expect_latest}; observed {latest_matches}",
            )
        )

    raw_assets = as_list(release.get("assets"), "snapshot.release.assets")
    names: list[str] = []
    for asset in raw_assets:
        if not isinstance(asset, dict):
            continue
        name = asset.get("name")
        if isinstance(name, str):
            names.append(name)
    if len(names) != len(raw_assets) or len(names) != len(set(names)):
        findings.append(finding("deterministic", "assets", "uploaded asset names are invalid/duplicate"))
    allowed = [str(item) for item in identity["allowed_asset_globs"]]
    unexpected = [
        name
        for name in names
        if not any(fnmatch.fnmatchcase(name, pattern) for pattern in allowed)
    ]
    if not allowed and names:
        findings.append(
            finding("deterministic", "assets", "source-only release profile contains uploaded assets")
        )
    if unexpected:
        findings.append(
            finding(
                "deterministic",
                "assets",
                "uploaded assets outside allowed_asset_globs: " + ", ".join(sorted(unexpected)),
            )
        )

    zip_url = release.get("zipball_url")
    tar_url = release.get("tarball_url")
    zip_exposed = isinstance(zip_url, str) and bool(zip_url)
    tar_exposed = isinstance(tar_url, str) and bool(tar_url)
    if not zip_exposed or not tar_exposed:
        findings.append(
            finding(
                "deterministic",
                "source-archives",
                "GitHub-generated source archive URLs are incomplete",
            )
        )
    archive_observation = as_object(
        snapshot.get("source_archives"), "snapshot.source_archives"
    )
    for kind in ("zip", "tar"):
        observed = archive_observation.get(kind)
        if observed != "pass":
            findings.append(
                finding(
                    "observation",
                    "source-archives",
                    f"{kind} source archive retrieval is {observed!r}",
                )
            )
    return {
        "release_id": release.get("id"),
        "published_at": release.get("published_at"),
        "draft": release.get("draft"),
        "prerelease": prerelease,
        "latest_expected": expect_latest,
        "latest_matches": latest_matches,
        "asset_mode": "source-only" if not allowed else "binary-capable",
        "uploaded_assets": sorted(names),
        "unexpected_assets": sorted(unexpected),
        "source_zip_exposed": zip_exposed,
        "source_tar_exposed": tar_exposed,
    }


def check_compare(
    snapshot: dict[str, Any],
    identity: dict[str, Any],
    candidate: str,
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    previous = identity.get("previous_tag")
    raw_compare = snapshot.get("compare")
    pages = raw_compare if isinstance(raw_compare, list) else [raw_compare]
    require(bool(pages), "snapshot.compare must contain at least one page")
    compare = as_object(pages[0], "snapshot.compare page 1")
    if previous is None:
        return {"previous_tag": None, "status": None, "candidate_seen": False}
    previous_tag = as_string(previous, "candidate previous tag")
    expected = f"/compare/{previous_tag}...v{identity['version']}"
    html_url = compare.get("html_url")
    if not isinstance(html_url, str) or not html_url.endswith(expected):
        findings.append(
            finding("deterministic", "comparison-link", "provider comparison range mismatches changelog")
        )
    status = compare.get("status")
    if status not in {"ahead", "identical"}:
        findings.append(
            finding("deterministic", "comparison-link", f"provider comparison status is {status!r}")
        )
    commits: list[dict[str, Any]] = []
    metadata_consistent = True
    for index, raw_page in enumerate(pages, start=1):
        page = as_object(raw_page, f"snapshot.compare page {index}")
        if page.get("html_url") != html_url or page.get("status") != status:
            metadata_consistent = False
        page_commits = as_list(page.get("commits"), f"snapshot.compare page {index}.commits")
        commits.extend(row for row in page_commits if isinstance(row, dict))
    if not metadata_consistent:
        findings.append(
            finding(
                "deterministic",
                "comparison-link",
                "paginated provider comparison metadata is inconsistent",
            )
        )
    candidate_seen = any(row.get("sha") == candidate for row in commits)
    if status == "ahead" and not candidate_seen:
        findings.append(
            finding(
                "deterministic",
                "comparison-link",
                "provider comparison does not contain the certified candidate",
            )
        )
    return {"previous_tag": previous_tag, "status": status, "candidate_seen": candidate_seen}


def check_milestone(
    snapshot: dict[str, Any],
    milestone_number: int,
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    milestone = as_object(snapshot.get("milestone"), "snapshot.milestone")
    items = milestone_items(snapshot.get("milestone_items"))
    if milestone.get("number") != milestone_number:
        findings.append(finding("deterministic", "milestone", "milestone number mismatches input"))
    if milestone.get("state") != "closed":
        findings.append(finding("deterministic", "milestone", "release milestone is not closed"))
    for row in items:
        row_milestone = row.get("milestone")
        if not isinstance(row_milestone, dict) or row_milestone.get("number") != milestone_number:
            findings.append(
                finding("deterministic", "milestone", "captured item belongs to another milestone")
            )
            break
    open_items = sorted(
        row["number"]
        for row in items
        if row.get("state") == "open" and isinstance(row.get("number"), int)
    )
    closed = [row for row in items if row.get("state") == "closed"]
    if any(row.get("state") not in {"open", "closed"} for row in items):
        findings.append(finding("deterministic", "milestone", "unsupported milestone item state"))
    if open_items:
        findings.append(
            finding(
                "deterministic",
                "milestone",
                "release milestone has open items: " + ", ".join(f"#{value}" for value in open_items),
            )
        )
    if milestone.get("open_issues") != len(open_items) or milestone.get("closed_issues") != len(closed):
        findings.append(
            finding(
                "deterministic",
                "milestone",
                "provider milestone counters disagree with actual captured membership/state",
            )
        )
    return {
        "number": milestone_number,
        "title": milestone.get("title"),
        "state": milestone.get("state"),
        "open_items": open_items,
        "closed_items": len(closed),
        "membership_count": len(items),
    }


def check_wiki(
    snapshot: dict[str, Any],
    identity: dict[str, Any],
    candidate: str,
    findings: list[dict[str, str]],
) -> dict[str, Any]:
    if identity["mode"] != "template":
        return {
            "applicable": False,
            "status": "not-applicable",
            "source_sha": None,
            "browser_review": "not-applicable",
        }
    wiki = as_object(snapshot.get("wiki"), "snapshot.wiki")
    status = wiki.get("status")
    source_sha = wiki.get("source_sha")
    if status != "pass":
        findings.append(
            finding("observation", "wiki", f"authoritative Wiki report is {status!r}")
        )
    if source_sha != candidate:
        findings.append(
            finding(
                "observation",
                "wiki",
                f"Wiki source SHA is {source_sha!r}; expected {candidate}",
            )
        )
    browser = as_object(snapshot.get("ui_observations"), "snapshot.ui_observations")
    browser_review = browser.get("wiki_browser_review")
    if browser_review != "pass":
        findings.append(
            finding(
                "observation",
                "wiki-browser-review",
                f"Wiki browser review is {browser_review!r}",
            )
        )
    return {
        "applicable": True,
        "status": status,
        "source_sha": source_sha,
        "browser_review": browser_review,
    }


def build_report(options: Options) -> dict[str, Any]:
    root = options.root.resolve()
    candidate = options.candidate_sha
    require(SHA_PATTERN.fullmatch(candidate) is not None, "candidate SHA must be full lowercase hex")
    require(
        git_text(root, "cat-file", "-e", f"{candidate}^{{commit}}").returncode == 0,
        f"candidate commit is unavailable locally: {candidate}",
    )
    snapshot = as_object(load_json(options.snapshot), "snapshot root")
    require(snapshot.get("schema_version") == 1, "unsupported closeout snapshot schema")

    findings: list[dict[str, str]] = []
    identity = source_identity(root, candidate, options.tag, findings)
    if snapshot.get("repository") != identity["repository"]:
        findings.append(
            finding("deterministic", "provider-scope", "snapshot repository mismatches candidate")
        )
    tag_state = check_tag(snapshot, options.tag, candidate, findings)
    workflow = check_workflow(
        snapshot, options.workflow_path, options.tag, candidate, findings
    )
    release = check_release(
        snapshot, identity, options.expect_prerelease, options.expect_latest, findings
    )
    comparison = check_compare(snapshot, identity, candidate, findings)
    milestone = check_milestone(snapshot, options.milestone_number, findings)
    wiki = check_wiki(snapshot, identity, candidate, findings)

    deterministic = [row for row in findings if row["category"] == "deterministic"]
    observations = [row for row in findings if row["category"] == "observation"]
    return {
        "schema_version": 1,
        "gate": "release-closeout",
        "status": "pass" if not findings else "fail",
        "deterministic_status": "pass" if not deterministic else "fail",
        "observation_status": "pass" if not observations else "incomplete",
        "repository": identity["repository"],
        "candidate_sha": candidate,
        "version": identity["version"],
        "tag": options.tag,
        "profile": identity["profile"],
        "snapshot_sha256": hashlib.sha256(options.snapshot.read_bytes()).hexdigest(),
        "source": {
            "version": identity["version"],
            "changelog_date": identity["changelog_date"],
            "previous_tag": identity["previous_tag"],
        },
        "tag_state": tag_state,
        "tag_workflow": workflow,
        "release": release,
        "comparison": comparison,
        "milestone": milestone,
        "wiki": wiki,
        "findings": findings,
        "scope_note": (
            "Read-only closeout validation over one exact Git candidate and captured GitHub REST "
            "facts. Uploaded assets are distinct from GitHub-generated source archives. Wiki "
            "publication bytes remain owned by check_wiki.py; browser and archive retrieval are "
            "explicit observations rather than inferred repository facts."
        ),
    }


def markdown_report(report: dict[str, Any]) -> str:
    controls = [
        (
            "Annotated tag / candidate",
            report["tag_state"]["ref_type"] == "tag"
            and report["tag_state"]["target_sha"] == report["candidate_sha"],
            report["tag_state"]["target_sha"],
        ),
        (
            "Tag-triggered workflow",
            report["tag_workflow"]["conclusion"] == "success",
            f"run {report['tag_workflow']['run_id']}",
        ),
        (
            "GitHub Release/latest",
            not report["release"]["draft"]
            and report["release"]["latest_matches"] == report["release"]["latest_expected"],
            (
                f"id {report['release']['release_id']}; "
                f"expected latest={report['release']['latest_expected']}; "
                f"observed={report['release']['latest_matches']}"
            ),
        ),
        (
            "Uploaded assets",
            not report["release"]["unexpected_assets"],
            f"{report['release']['asset_mode']}; {len(report['release']['uploaded_assets'])} uploaded",
        ),
        (
            "Comparison link",
            report["comparison"]["status"] in {"ahead", "identical"},
            f"{report['comparison']['previous_tag']} → {report['tag']}",
        ),
        (
            "Milestone membership",
            not report["milestone"]["open_items"] and report["milestone"]["state"] == "closed",
            f"{report['milestone']['membership_count']} items",
        ),
    ]
    if report["wiki"]["applicable"]:
        controls.append(
            (
                "Wiki source/read-back",
                report["wiki"]["status"] == "pass"
                and report["wiki"]["source_sha"] == report["candidate_sha"],
                report["wiki"]["source_sha"],
            )
        )
    lines = [
        "# Release closeout",
        "",
        f"- **Status:** {report['status'].upper()}",
        f"- **Deterministic checks:** {report['deterministic_status'].upper()}",
        f"- **Provider/UI observations:** {report['observation_status'].upper()}",
        f"- **Release:** `{report['tag']}` at `{report['candidate_sha']}`",
        f"- **Profile:** `{report['profile']}`",
        f"- **Snapshot SHA-256:** `{report['snapshot_sha256']}`",
        "",
        "| Control | Result | Evidence |",
        "| --- | --- | --- |",
    ]
    lines.extend(
        f"| {name} | {'PASS' if ok else 'FAIL'} | {detail or 'n/a'} |"
        for name, ok, detail in controls
    )
    lines.extend(["", f"**Findings:** {len(report['findings'])}"])
    if report["findings"]:
        lines.append("")
        lines.extend(
            f"- **{row['category']} / {row['control']}:** {row['message']}"
            for row in report["findings"]
        )
    lines.extend(["", report["scope_note"], ""])
    return "\n".join(lines)


def write_report(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", newline="\n")


def write_json(path: Path, value: Any) -> None:
    write_report(path, json.dumps(value, indent=2, sort_keys=True) + "\n")


def fixture_git(root: Path, *arguments: str) -> str:
    completed = subprocess.run(
        [
            "git",
            "-C",
            str(root),
            "-c",
            "user.name=Closeout Fixture",
            "-c",
            "user.email=fixture@example.invalid",
            *arguments,
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    return completed.stdout.strip()


def fixture(base: Path) -> tuple[Path, Path, str]:
    root = base / "repo"
    root.mkdir()
    fixture_git(root, "init", "-b", "main")
    (root / ".github").mkdir()
    repository = "owner/repo"
    version = "1.2.3"
    tag = f"v{version}"
    (root / VERSION_PATH).write_text(version + "\n", encoding="utf-8")
    (root / CHANGELOG_PATH).write_text(
        "# Changelog\n\n"
        "## [Unreleased]\n\n"
        f"## [{version}] - 2026-09-12\n\n"
        f"[Unreleased]: https://github.com/{repository}/compare/{tag}...HEAD\n"
        f"[{version}]: https://github.com/{repository}/compare/v1.2.2...{tag}\n",
        encoding="utf-8",
    )
    write_json(
        root / PROFILE_PATH,
        {"schema_version": 1, "mode": "template", "profile": None, "repository": repository},
    )
    write_json(
        root / RELEASE_POLICY_PATH,
        {"schema_version": 1, "profiles": {"template": {"allowed_asset_globs": []}}},
    )
    fixture_git(root, "add", "--all")
    fixture_git(root, "commit", "-m", "Fixture release")
    candidate = fixture_git(root, "rev-parse", "HEAD")
    tag_object_sha = "b" * 40
    release = {
        "id": 55,
        "tag_name": tag,
        "draft": False,
        "prerelease": False,
        "published_at": "2026-09-12T10:00:00Z",
        "assets": [],
        "zipball_url": f"https://api.github.com/repos/{repository}/zipball/{tag}",
        "tarball_url": f"https://api.github.com/repos/{repository}/tarball/{tag}",
    }
    snapshot = {
        "schema_version": 1,
        "repository": repository,
        "tag_ref": {
            "ref": f"refs/tags/{tag}",
            "object": {"type": "tag", "sha": tag_object_sha},
        },
        "tag_object": {
            "sha": tag_object_sha,
            "tag": tag,
            "object": {"type": "commit", "sha": candidate},
        },
        "workflow_runs": {
            "total_count": 1,
            "workflow_runs": [
                {
                    "id": 123,
                    "run_number": 7,
                    "run_attempt": 1,
                    "path": DEFAULT_WORKFLOW,
                    "head_branch": tag,
                    "head_sha": candidate,
                    "event": "push",
                    "status": "completed",
                    "conclusion": "success",
                }
            ],
        },
        "release": release,
        "latest_release": dict(release),
        "compare": {
            "html_url": f"https://github.com/{repository}/compare/v1.2.2...{tag}",
            "status": "ahead",
            "commits": [{"sha": candidate}],
        },
        "milestone": {
            "number": 4,
            "title": "v1.2.3",
            "state": "closed",
            "open_issues": 0,
            "closed_issues": 2,
        },
        "milestone_items": [
            {"number": 1, "state": "closed", "milestone": {"number": 4}},
            {"number": 2, "state": "closed", "milestone": {"number": 4}},
        ],
        "wiki": {"status": "pass", "source_sha": candidate, "findings": []},
        "source_archives": {"zip": "pass", "tar": "pass"},
        "ui_observations": {"wiki_browser_review": "pass"},
    }
    snapshot_path = base / "snapshot.json"
    write_json(snapshot_path, snapshot)
    return root, snapshot_path, candidate


def fixture_options(root: Path, snapshot: Path, candidate: str) -> Options:
    return Options(
        root=root,
        snapshot=snapshot,
        tag="v1.2.3",
        candidate_sha=candidate,
        milestone_number=4,
        workflow_path=DEFAULT_WORKFLOW,
        expect_prerelease=False,
        expect_latest=True,
    )


def exercise_review_regressions(failures: list[str]) -> None:
    with tempfile.TemporaryDirectory(prefix="release-closeout-optional-latest-") as raw:
        root, snapshot_path, candidate = fixture(Path(raw))
        value = as_object(load_json(snapshot_path), "fixture snapshot")
        value["release"]["prerelease"] = True
        value["latest_release"] = None
        write_json(snapshot_path, value)
        base_options = fixture_options(root, snapshot_path, candidate)
        options = Options(
            root=base_options.root,
            snapshot=base_options.snapshot,
            tag=base_options.tag,
            candidate_sha=base_options.candidate_sha,
            milestone_number=base_options.milestone_number,
            workflow_path=base_options.workflow_path,
            expect_prerelease=True,
            expect_latest=False,
        )
        report = build_report(options)
        summary = markdown_report(report)
        if report["status"] != "pass" or "| GitHub Release/latest | PASS |" not in summary:
            failures.append(
                "optional-latest: absent /releases/latest was not accepted and rendered consistently"
            )

    with tempfile.TemporaryDirectory(prefix="release-closeout-compare-pages-") as raw:
        root, snapshot_path, candidate = fixture(Path(raw))
        value = as_object(load_json(snapshot_path), "fixture snapshot")
        page_one = dict(value["compare"])
        page_two = dict(value["compare"])
        page_one["commits"] = [{"sha": "a" * 40}]
        page_two["commits"] = [{"sha": candidate}]
        value["compare"] = [page_one, page_two]
        write_json(snapshot_path, value)
        report = build_report(fixture_options(root, snapshot_path, candidate))
        if report["status"] != "pass" or not report["comparison"]["candidate_seen"]:
            failures.append("paginated-compare: candidate on a later page was not accepted")


def run_self_test() -> int:
    failures: list[str] = []

    def run_case(name: str, mutate: Any = None, expected: str | None = None) -> None:
        with tempfile.TemporaryDirectory(prefix="release-closeout-") as raw:
            root, snapshot_path, candidate = fixture(Path(raw))
            if mutate is not None:
                value = as_object(load_json(snapshot_path), "fixture snapshot")
                mutate(value, candidate)
                write_json(snapshot_path, value)
            report = build_report(fixture_options(root, snapshot_path, candidate))
            controls = {row["control"] for row in report["findings"]}
            expected_status = "pass" if expected is None else "fail"
            ok = report["status"] == expected_status
            ok = ok and (expected is None or expected in controls)
            if not ok:
                failures.append(f"{name}: unexpected report {report['findings']!r}")

    run_case("valid-closeout")
    run_case(
        "lightweight-tag",
        lambda value, candidate: value.update(
            {"tag_ref": {"ref": "refs/tags/v1.2.3", "object": {"type": "commit", "sha": candidate}}}
        ),
        "tag",
    )

    def moved(value: dict[str, Any], _candidate: str) -> None:
        value["tag_object"]["object"]["sha"] = "c" * 40

    run_case("moved-tag", moved, "tag")

    def failed_ci(value: dict[str, Any], _candidate: str) -> None:
        value["workflow_runs"]["workflow_runs"][0]["conclusion"] = "failure"

    run_case("failed-tag-ci", failed_ci, "tag-ci")

    def draft(value: dict[str, Any], _candidate: str) -> None:
        value["release"]["draft"] = True

    run_case("draft-release", draft, "release")

    def prerelease(value: dict[str, Any], _candidate: str) -> None:
        value["release"]["prerelease"] = True

    run_case("unexpected-prerelease", prerelease, "release")

    def latest(value: dict[str, Any], _candidate: str) -> None:
        value["latest_release"]["id"] = 99

    run_case("not-latest", latest, "release")
    exercise_review_regressions(failures)

    def asset(value: dict[str, Any], _candidate: str) -> None:
        value["release"]["assets"] = [{"name": "dist/unexpected.zip"}]

    run_case("unexpected-asset", asset, "assets")

    def archive_url(value: dict[str, Any], _candidate: str) -> None:
        value["release"]["zipball_url"] = None

    run_case("missing-archive-url", archive_url, "source-archives")

    def compare(value: dict[str, Any], _candidate: str) -> None:
        value["compare"]["status"] = "diverged"

    run_case("unresolved-compare", compare, "comparison-link")

    def open_item(value: dict[str, Any], _candidate: str) -> None:
        value["milestone_items"][0]["state"] = "open"
        value["milestone"]["open_issues"] = 1
        value["milestone"]["closed_issues"] = 1

    run_case("open-milestone", open_item, "milestone")

    def open_milestone(value: dict[str, Any], _candidate: str) -> None:
        value["milestone"]["state"] = "open"

    run_case("milestone-state", open_milestone, "milestone")

    def stale_count(value: dict[str, Any], _candidate: str) -> None:
        value["milestone"]["closed_issues"] = 99

    run_case("milestone-counter-drift", stale_count, "milestone")

    def wiki(value: dict[str, Any], _candidate: str) -> None:
        value["wiki"]["status"] = "fail"

    run_case("wiki-drift", wiki, "wiki")

    def archive(value: dict[str, Any], _candidate: str) -> None:
        value["source_archives"]["zip"] = "fail"

    run_case("archive-retrieval", archive, "source-archives")

    with tempfile.TemporaryDirectory(prefix="release-closeout-version-") as raw:
        root, snapshot_path, candidate = fixture(Path(raw))
        wrong = fixture_options(root, snapshot_path, candidate)
        wrong = Options(
            root=wrong.root,
            snapshot=wrong.snapshot,
            tag="v1.2.4",
            candidate_sha=wrong.candidate_sha,
            milestone_number=wrong.milestone_number,
            workflow_path=wrong.workflow_path,
            expect_prerelease=wrong.expect_prerelease,
            expect_latest=wrong.expect_latest,
        )
        report = build_report(wrong)
        controls = {row["control"] for row in report["findings"]}
        if "source-version" not in controls:
            failures.append("wrong-version: tag/VERSION mismatch was not rejected")

    if failures:
        print("SELF-TEST FAIL:", file=sys.stderr)
        for failure in failures:
            print(f"  - {failure}", file=sys.stderr)
        return 1
    print(
        "SELF-TEST PASS: VERSION/tag binding, annotated/lightweight/moved tags, tag CI, "
        "release state including optional latest, paginated comparison resolution, asset policy, "
        "milestone state/membership/counters, Wiki binding and source-archive observations passed."
    )
    return 0


def parse_arguments(argv: list[str]) -> tuple[Options | None, bool]:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--snapshot", type=Path)
    parser.add_argument("--tag")
    parser.add_argument("--candidate-sha")
    parser.add_argument("--milestone-number", type=int)
    parser.add_argument("--workflow-path", default=DEFAULT_WORKFLOW)
    parser.add_argument("--expect-prerelease", action="store_true")
    parser.add_argument("--allow-not-latest", dest="expect_latest", action="store_false")
    parser.set_defaults(expect_latest=True)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args(argv)
    if args.self_test:
        return None, True
    missing = [
        name
        for name in ("snapshot", "tag", "candidate_sha", "milestone_number")
        if getattr(args, name) in (None, "")
    ]
    if missing:
        parser.error(
            "operational closeout requires "
            + ", ".join("--" + name.replace("_", "-") for name in missing)
        )
    return (
        Options(
            root=args.root,
            snapshot=args.snapshot,
            tag=args.tag,
            candidate_sha=args.candidate_sha,
            milestone_number=args.milestone_number,
            workflow_path=args.workflow_path,
            expect_prerelease=args.expect_prerelease,
            expect_latest=args.expect_latest,
            output=args.output,
            summary=args.summary,
        ),
        False,
    )


def main(argv: list[str] | None = None) -> int:
    try:
        options, self_test = parse_arguments(sys.argv[1:] if argv is None else argv)
        if self_test:
            return run_self_test()
        if options is None:
            raise CloseoutError("operational options are unavailable")
        report = build_report(options)
        markdown = markdown_report(report)
        if options.output is not None:
            write_json(options.output, report)
        if options.summary is not None:
            write_report(options.summary, markdown)
        print(markdown, end="")
        return 0 if report["status"] == "pass" else 1
    except (CloseoutError, OSError, ValueError, TypeError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
