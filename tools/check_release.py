#!/usr/bin/env python3
"""Validate one evidence-backed Excel/VBA release candidate.

The checker is dependency-free and deliberately separates immutable candidate
source from evidence and staged release assets. Evidence may be retained by a
workflow, release draft, or review record without creating a commit-hash
self-reference inside the candidate it certifies.

Exit status is 0 for a valid candidate, 1 for policy findings, and 2 for an
operational error.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path, PurePosixPath
from typing import Any

from check_excel_evidence import release_findings as excel_release_findings
from release_provenance import validate as validate_provenance

SCHEMA_VERSION = 1
TOOL_NAME = "Canonical release integrity"
POLICY_PATH = ".github/release-policy.json"
PROFILE_PATH = ".github/repository-profile.json"
INITIALIZATION_PATH = ".github/initialization.json"
GENERATED_PROFILES = ("application", "library", "ui-component")
SUPPORTED_PROFILES = ("application", "library", "template", "ui-component")
SEMVER_PATTERN = re.compile(
    r"(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)"
    r"(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?"
    r"(?:\+([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?"
)
SHA256_PATTERN = re.compile(r"[0-9a-f]{64}")
COMMIT_PATTERN = re.compile(r"[0-9a-f]{40}")
CHECK_ID_PATTERN = re.compile(r"[a-z][a-z0-9]*(?:-[a-z0-9]+)*")
CHANGELOG_HEADING_PATTERN = re.compile(
    r"^## \[([^]\r\n]+)\] - ([0-9]{4}-[0-9]{2}-[0-9]{2})$", re.MULTILINE
)
TEXT_SUFFIXES = {
    ".bas", ".bat", ".cfg", ".cff", ".cjs", ".cls", ".cmd", ".csv",
    ".frm", ".ini", ".js", ".json", ".jsonc", ".md", ".mjs", ".ps1",
    ".psd1", ".psm1", ".py", ".r", ".reg", ".sh", ".svg", ".toml",
    ".tsv", ".txt", ".vbs", ".xml", ".yaml", ".yml",
}
TEXT_NAMES = {
    ".editorconfig", ".gitattributes", ".gitignore", "Dockerfile", "LICENSE",
    "Makefile", "VERSION",
}
TOP_LEVEL_EVIDENCE_KEYS = {
    "schema_version", "version", "tag", "candidate_sha", "profile",
    "distribution", "checks", "assets",
}
ASSET_KEYS = {"path", "sha256", "candidate_sha", "package_test"}
BASE_CHECK_FIELDS = {"status", "candidate_sha", "detail"}
CORE_CHECK_FIELDS = {
    "repository-integrity": {"run_url"},
    "vba-compile": {"environment"},
    "regression": {
        "entry_point", "environment", "cases", "assertions", "failures",
        "completeness", "cleanup",
    },
}


class OperationalError(RuntimeError):
    """The checker could not evaluate the requested candidate."""


def _finding(code: str, path: str, message: str) -> dict[str, str]:
    return {"code": code, "path": path, "message": message}


def _read_json(path: Path, label: str) -> tuple[object | None, list[dict[str, str]]]:
    try:
        text = path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return None, [_finding(f"missing-{label}", str(path), f"{label.replace('-', ' ')} is required")]
    except OSError as error:
        raise OperationalError(f"Could not read {path}: {error}") from error
    try:
        return json.loads(text), []
    except json.JSONDecodeError as error:
        return None, [_finding(
            f"invalid-{label}", str(path),
            f"invalid JSON at line {error.lineno}, column {error.colno}: {error.msg}",
        )]


def _load_configuration(root: Path) -> tuple[dict[str, Any] | None, list[dict[str, str]]]:
    value, findings = _read_json(root / PROFILE_PATH, "repository-profile")
    if findings:
        return None, findings
    if not isinstance(value, dict):
        return None, [_finding("invalid-repository-profile", PROFILE_PATH, "root must be an object")]
    return value, []


def _load_policy(root: Path) -> tuple[dict[str, Any] | None, list[dict[str, str]]]:
    value, findings = _read_json(root / POLICY_PATH, "release-policy")
    if findings:
        return None, findings
    if not isinstance(value, dict):
        return None, [_finding("invalid-release-policy", POLICY_PATH, "root must be an object")]
    required = {
        "schema_version", "evidence_schema_version", "provenance_signature_mode",
        "core_checks", "profiles", "source_scan_exclude_paths",
        "template_construction_markers",
    }
    if set(value) != required:
        missing = sorted(required - set(value))
        extra = sorted(set(value) - required)
        detail = []
        if missing:
            detail.append("missing " + ", ".join(missing))
        if extra:
            detail.append("unknown " + ", ".join(extra))
        return None, [_finding("invalid-release-policy", POLICY_PATH, "; ".join(detail))]
    if value.get("schema_version") != SCHEMA_VERSION or value.get("evidence_schema_version") != SCHEMA_VERSION:
        return None, [_finding("invalid-release-policy", POLICY_PATH, "unsupported schema version")]
    signature_mode = value.get("provenance_signature_mode")
    if not isinstance(signature_mode, str) or signature_mode not in {"none", "ssh"}:
        return None, [_finding(
            "invalid-release-policy", POLICY_PATH,
            "provenance_signature_mode must be none or ssh",
        )]
    core = value.get("core_checks")
    profiles = value.get("profiles")
    excludes = value.get("source_scan_exclude_paths")
    markers = value.get("template_construction_markers")
    if not isinstance(core, list) or not core or not all(isinstance(item, str) and CHECK_ID_PATTERN.fullmatch(item) for item in core):
        return None, [_finding("invalid-release-policy", POLICY_PATH, "core_checks must be a non-empty list of canonical identifiers")]
    if len(core) != len(set(core)):
        return None, [_finding("invalid-release-policy", POLICY_PATH, "core_checks contains duplicates")]
    if not isinstance(profiles, dict) or set(profiles) != set(SUPPORTED_PROFILES):
        return None, [_finding(
            "invalid-release-policy", POLICY_PATH,
            "profiles must define exactly application, library, template, and ui-component",
        )]
    for profile, specification in profiles.items():
        if not isinstance(specification, dict) or set(specification) != {"required_checks", "allowed_asset_globs"}:
            return None, [_finding("invalid-release-policy", POLICY_PATH, f"{profile} has an invalid policy shape")]
        checks = specification["required_checks"]
        globs = specification["allowed_asset_globs"]
        if not isinstance(checks, list) or not checks or not all(isinstance(item, str) and CHECK_ID_PATTERN.fullmatch(item) for item in checks):
            return None, [_finding("invalid-release-policy", POLICY_PATH, f"{profile}.required_checks is invalid")]
        if len(checks) != len(set(checks)) or set(checks) & set(core):
            return None, [_finding("invalid-release-policy", POLICY_PATH, f"{profile}.required_checks contains duplicates or core checks")]
        if not isinstance(globs, list) or not all(_safe_relative(item) for item in globs):
            return None, [_finding("invalid-release-policy", POLICY_PATH, f"{profile}.allowed_asset_globs is invalid")]
    if not isinstance(excludes, list) or not all(_safe_relative(item) for item in excludes):
        return None, [_finding("invalid-release-policy", POLICY_PATH, "source_scan_exclude_paths is invalid")]
    if not isinstance(markers, list) or not markers or not all(isinstance(item, str) and item.strip() for item in markers):
        return None, [_finding("invalid-release-policy", POLICY_PATH, "template_construction_markers is invalid")]
    return value, []


def _safe_relative(value: object) -> bool:
    if not isinstance(value, str) or not value or "\\" in value or "\0" in value:
        return False
    path = PurePosixPath(value)
    return not path.is_absolute() and ".." not in path.parts and path.as_posix() == value


def _git(root: Path, *arguments: str) -> subprocess.CompletedProcess[str]:
    try:
        return subprocess.run(
            ["git", "-C", str(root), *arguments], check=False,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
            encoding="utf-8", errors="replace",
        )
    except FileNotFoundError as error:
        raise OperationalError("git executable was not found on PATH") from error


def _git_output(root: Path, *arguments: str) -> str | None:
    result = _git(root, *arguments)
    return result.stdout.strip() if result.returncode == 0 else None


def _tracked_files(root: Path) -> list[str]:
    output = _git_output(root, "ls-files", "-z")
    if output is None:
        raise OperationalError(f"{root} is not a readable Git working tree")
    return sorted(item for item in output.split("\0") if item)


def _is_text(path: str) -> bool:
    item = PurePosixPath(path)
    return item.name in TEXT_NAMES or item.suffix.casefold() in TEXT_SUFFIXES


def _resolve_release_profile(
    root: Path, configuration: dict[str, Any]
) -> tuple[str | None, list[dict[str, str]]]:
    findings: list[dict[str, str]] = []
    mode = configuration.get("mode")
    configured_profile = configuration.get("profile")
    initialization = root / INITIALIZATION_PATH
    if mode == "generated" and configured_profile in GENERATED_PROFILES:
        release_profile = str(configured_profile)
        if not initialization.is_file():
            findings.append(_finding(
                "template-identity", INITIALIZATION_PATH,
                "generated release candidates require an initialization record",
            ))
            return release_profile, findings
        record, record_findings = _read_json(initialization, "initialization-record")
        findings.extend(record_findings)
        if isinstance(record, dict):
            if record.get("profile") != release_profile:
                findings.append(_finding(
                    "template-identity", INITIALIZATION_PATH,
                    "profile differs from repository policy",
                ))
            values = record.get("values")
            repository = configuration.get("repository")
            if not isinstance(values, dict) or values.get("REPOSITORY_PATH") != repository:
                findings.append(_finding(
                    "template-identity", INITIALIZATION_PATH,
                    "repository identity differs from repository policy",
                ))
        return release_profile, findings
    if mode == "template" and configured_profile is None:
        identity = configuration.get("identity")
        repository = configuration.get("repository")
        template_tokens = identity.get("template_tokens") if isinstance(identity, dict) else None
        contains_template_identity = (
            isinstance(repository, str)
            and isinstance(template_tokens, list)
            and any(
                isinstance(token, str) and token.casefold() in repository.casefold()
                for token in template_tokens
            )
        )
        if not contains_template_identity:
            findings.append(_finding(
                "template-identity", PROFILE_PATH,
                "template release candidates must retain a declared template identity token",
            ))
        if initialization.exists():
            findings.append(_finding(
                "template-identity", INITIALIZATION_PATH,
                "template release candidates must not contain a generated-project initialization record",
            ))
        return "template", findings
    findings.append(_finding(
        "template-identity", PROFILE_PATH,
        "release candidates must be an initialized generated profile or the canonical template profile",
    ))
    return None, findings


def _validate_candidate_git_state(
    root: Path, candidate_sha: str, tag: str, require_tag_ref: bool
) -> list[dict[str, str]]:
    findings: list[dict[str, str]] = []
    head = _git_output(root, "rev-parse", "HEAD")
    if head is None:
        raise OperationalError("git could not resolve HEAD")
    if head != candidate_sha:
        findings.append(_finding("candidate-sha-mismatch", "HEAD", f"HEAD is {head}, expected {candidate_sha}"))
    dirty = _git_output(root, "status", "--porcelain", "--untracked-files=no")
    if dirty is None:
        raise OperationalError("git could not inspect the working tree")
    if dirty:
        findings.append(_finding("dirty-candidate", ".", "tracked candidate files differ from HEAD"))
    if not require_tag_ref:
        return findings
    reference = f"refs/tags/{tag}"
    object_type = _git_output(root, "cat-file", "-t", reference)
    target = _git_output(root, "rev-list", "-n", "1", reference)
    if object_type is None or target is None:
        findings.append(_finding("missing-tag-ref", tag, "annotated tag is not available in this clone"))
        return findings
    if object_type != "tag":
        findings.append(_finding("lightweight-tag", tag, "release tag must be annotated"))
    if target != candidate_sha:
        findings.append(_finding("tag-target-mismatch", tag, f"tag targets {target}, expected {candidate_sha}"))
    return findings


def _validate_generated_source(
    root: Path, configuration: dict[str, Any], policy: dict[str, Any]
) -> list[dict[str, str]]:
    findings: list[dict[str, str]] = []
    tracked = _tracked_files(root)
    identity = configuration.get("identity")
    identity_excludes: set[str] = set()
    source_template_tokens: list[str] = []
    if isinstance(identity, dict):
        raw_excludes = identity.get("exclude_paths")
        raw_tokens = identity.get("template_tokens")
        if isinstance(raw_excludes, list):
            identity_excludes.update(item for item in raw_excludes if isinstance(item, str))
        if isinstance(raw_tokens, list):
            source_template_tokens.extend(item for item in raw_tokens if isinstance(item, str))
    excludes = set(policy["source_scan_exclude_paths"]) | identity_excludes
    placeholder_pattern = re.compile(r"\{\{[A-Z][A-Z0-9_]*\}\}")
    for relative in tracked:
        if relative in excludes or not _is_text(relative):
            continue
        path = root / relative
        try:
            text = path.read_text(encoding="utf-8", errors="strict")
        except (OSError, UnicodeError) as error:
            findings.append(_finding("unreadable-release-text", relative, str(error)))
            continue
        if placeholder_pattern.search(text) or "<!-- template:" in text:
            findings.append(_finding(
                "unresolved-template-token", relative,
                "unresolved placeholder or template block marker",
            ))
        folded = text.casefold()
        for token in source_template_tokens:
            if token.casefold() in folded:
                findings.append(_finding(
                    "template-identity", relative,
                    f"contains template identity token {token}",
                ))
                break
    try:
        changelog = (root / "CHANGELOG.md").read_text(encoding="utf-8")
    except OSError as error:
        findings.append(_finding("missing-changelog", "CHANGELOG.md", str(error)))
        changelog = ""
    for marker in policy["template_construction_markers"]:
        if str(marker).casefold() in changelog.casefold():
            findings.append(_finding(
                "template-construction-history", "CHANGELOG.md",
                f"contains construction marker {marker!r}",
            ))
    return findings


def _validate_source(
    root: Path,
    configuration: dict[str, Any],
    policy: dict[str, Any],
    candidate_sha: str,
    tag: str,
    require_tag_ref: bool,
) -> tuple[str | None, list[dict[str, str]]]:
    release_profile, findings = _resolve_release_profile(root, configuration)
    findings.extend(_validate_candidate_git_state(root, candidate_sha, tag, require_tag_ref))
    if release_profile in GENERATED_PROFILES:
        findings.extend(_validate_generated_source(root, configuration, policy))
    return release_profile, findings


def _validate_version_and_changelog(root: Path, tag: str) -> tuple[str | None, list[dict[str, str]]]:
    findings: list[dict[str, str]] = []
    try:
        raw = (root / "VERSION").read_text(encoding="utf-8")
    except OSError as error:
        return None, [_finding("missing-version", "VERSION", str(error))]
    version = raw.strip()
    if raw not in {version, version + "\n"} or SEMVER_PATTERN.fullmatch(version) is None:
        findings.append(_finding("invalid-version", "VERSION", "must contain one canonical semantic version"))
    if version == "0.0.0":
        findings.append(_finding("zero-version", "VERSION", "0.0.0 is a template sentinel, not a release version"))
    if tag != f"v{version}":
        findings.append(_finding("tag-version-mismatch", tag, f"expected v{version}"))

    try:
        changelog = (root / "CHANGELOG.md").read_text(encoding="utf-8")
    except OSError as error:
        return version, findings + [_finding("missing-changelog", "CHANGELOG.md", str(error))]
    matches = [match for match in CHANGELOG_HEADING_PATTERN.finditer(changelog) if match.group(1) == version]
    if len(matches) != 1:
        findings.append(_finding("missing-changelog-release", "CHANGELOG.md", f"expected one dated [{version}] release heading"))
    else:
        raw_date = matches[0].group(2)
        try:
            date.fromisoformat(raw_date)
        except ValueError:
            findings.append(_finding("invalid-changelog-date", "CHANGELOG.md", f"{raw_date} is not a calendar date"))
    return version, findings


def _validate_check(check_id: str, value: object, candidate_sha: str) -> list[dict[str, str]]:
    path = f"evidence.checks.{check_id}"
    if not isinstance(value, dict):
        return [_finding("invalid-evidence-check", path, "check record must be an object")]
    required = BASE_CHECK_FIELDS | CORE_CHECK_FIELDS.get(check_id, set())
    missing = sorted(required - set(value))
    if missing:
        return [_finding("invalid-evidence-check", path, "missing fields: " + ", ".join(missing))]
    findings: list[dict[str, str]] = []
    if value.get("status") != "PASS":
        findings.append(_finding("failed-evidence-check", path, "status must be PASS"))
    if value.get("candidate_sha") != candidate_sha:
        findings.append(_finding("evidence-sha-mismatch", path, "candidate_sha does not match the release candidate"))
    if not isinstance(value.get("detail"), str) or not value["detail"].strip():
        findings.append(_finding("invalid-evidence-check", path, "detail must be non-empty"))
    if check_id == "repository-integrity":
        run_url = value.get("run_url")
        if not isinstance(run_url, str) or not run_url.startswith("https://"):
            findings.append(_finding("invalid-evidence-check", path, "run_url must be HTTPS"))
    elif check_id == "vba-compile":
        if not isinstance(value.get("environment"), str) or not value["environment"].strip():
            findings.append(_finding("invalid-evidence-check", path, "environment must be non-empty"))
    elif check_id == "regression":
        if not isinstance(value.get("entry_point"), str) or not value["entry_point"].strip():
            findings.append(_finding("invalid-evidence-check", path, "entry_point must be non-empty"))
        if not isinstance(value.get("environment"), str) or not value["environment"].strip():
            findings.append(_finding("invalid-evidence-check", path, "environment must be non-empty"))
        cases = value.get("cases")
        assertions = value.get("assertions")
        failures = value.get("failures")
        if not isinstance(cases, int) or isinstance(cases, bool) or cases <= 0:
            findings.append(_finding("invalid-evidence-check", path, "cases must be a positive integer"))
        if not isinstance(assertions, int) or isinstance(assertions, bool) or assertions <= 0:
            findings.append(_finding("invalid-evidence-check", path, "assertions must be a positive integer"))
        if failures != 0:
            findings.append(_finding("failed-evidence-check", path, "failures must be zero"))
        if value.get("completeness") != "COMPLETE":
            findings.append(_finding("failed-evidence-check", path, "completeness must be COMPLETE"))
        if value.get("cleanup") != "PASS":
            findings.append(_finding("failed-evidence-check", path, "cleanup must be PASS"))
    return findings


def _parse_manifest(path: Path | None) -> tuple[dict[str, str] | None, list[dict[str, str]]]:
    if path is None:
        return None, []
    try:
        text = path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return None, [_finding("missing-asset-manifest", str(path), "asset manifest does not exist")]
    except OSError as error:
        raise OperationalError(f"Could not read {path}: {error}") from error
    entries: dict[str, str] = {}
    ordered: list[str] = []
    findings: list[dict[str, str]] = []
    for number, line in enumerate(text.splitlines(), 1):
        if not line:
            continue
        match = re.fullmatch(r"([0-9a-f]{64})  (.+)", line)
        if match is None or not _safe_relative(match.group(2)):
            findings.append(_finding("invalid-asset-manifest", f"{path}:{number}", "expected '<sha256>  <safe-relative-path>'"))
            continue
        digest, relative = match.groups()
        if relative in entries:
            findings.append(_finding("invalid-asset-manifest", f"{path}:{number}", f"duplicate asset {relative}"))
            continue
        entries[relative] = digest
        ordered.append(relative)
    if ordered != sorted(ordered):
        findings.append(_finding("invalid-asset-manifest", str(path), "entries must be sorted by path"))
    return entries, findings


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _validate_evidence_metadata(
    evidence: dict[str, Any],
    evidence_path: Path,
    policy: dict[str, Any],
    profile: str | None,
    version: str | None,
    tag: str,
    candidate_sha: str,
) -> list[dict[str, str]]:
    findings: list[dict[str, str]] = []
    if set(evidence) != TOP_LEVEL_EVIDENCE_KEYS:
        missing = sorted(TOP_LEVEL_EVIDENCE_KEYS - set(evidence))
        extra = sorted(set(evidence) - TOP_LEVEL_EVIDENCE_KEYS)
        detail = []
        if missing:
            detail.append("missing " + ", ".join(missing))
        if extra:
            detail.append("unknown " + ", ".join(extra))
        findings.append(_finding("invalid-release-evidence", str(evidence_path), "; ".join(detail)))
    expected = {
        "schema_version": policy["evidence_schema_version"],
        "version": version,
        "tag": tag,
        "candidate_sha": candidate_sha,
        "profile": profile,
    }
    for key, expected_value in expected.items():
        if evidence.get(key) != expected_value:
            code = "evidence-sha-mismatch" if key == "candidate_sha" else "evidence-metadata-mismatch"
            findings.append(_finding(code, f"evidence.{key}", f"expected {expected_value!r}"))
    if evidence.get("distribution") not in {"source-only", "binary"}:
        findings.append(_finding("invalid-release-evidence", "evidence.distribution", "must be source-only or binary"))
    return findings


def _validate_evidence_checks(
    evidence: dict[str, Any],
    policy: dict[str, Any],
    profile: str | None,
    candidate_sha: str,
) -> list[dict[str, str]]:
    findings: list[dict[str, str]] = []
    checks = evidence.get("checks")
    required_checks: set[str] = set(policy["core_checks"])
    if profile in SUPPORTED_PROFILES:
        required_checks.update(policy["profiles"][profile]["required_checks"])
    if not isinstance(checks, dict):
        return [_finding("invalid-release-evidence", "evidence.checks", "must be an object")]
    invalid_ids = sorted(
        key for key in checks
        if not isinstance(key, str) or CHECK_ID_PATTERN.fullmatch(key) is None
    )
    if invalid_ids:
        findings.append(_finding("invalid-release-evidence", "evidence.checks", "invalid check identifiers"))
    missing_checks = sorted(required_checks - set(checks))
    if missing_checks:
        findings.append(_finding(
            "missing-profile-evidence", "evidence.checks",
            "missing checks: " + ", ".join(missing_checks),
        ))
    for check_id in sorted(checks):
        findings.extend(_validate_check(check_id, checks[check_id], candidate_sha))
    return findings


def _validate_asset_record(
    root: Path,
    asset: object,
    item_path: str,
    allowed: list[str],
    candidate_sha: str,
    evidence_entries: dict[str, str],
) -> list[dict[str, str]]:
    findings: list[dict[str, str]] = []
    if not isinstance(asset, dict) or set(asset) != ASSET_KEYS:
        return [_finding("invalid-release-asset", item_path, "asset record has an invalid shape")]
    relative = asset.get("path")
    digest = asset.get("sha256")
    if not isinstance(relative, str) or not _safe_relative(relative):
        return [_finding("invalid-release-asset", item_path, "path must be safe and repository-relative")]
    if relative in evidence_entries:
        return [_finding("invalid-release-asset", item_path, f"duplicate asset {relative}")]
    if not isinstance(digest, str) or SHA256_PATTERN.fullmatch(digest) is None:
        return [_finding("invalid-release-asset", item_path, "sha256 must be 64 lowercase hexadecimal characters")]
    evidence_entries[relative] = digest
    if asset.get("candidate_sha") != candidate_sha:
        findings.append(_finding("asset-sha-binding-mismatch", item_path, "candidate_sha does not match the release candidate"))
    if asset.get("package_test") != "PASS":
        findings.append(_finding("failed-package-test", item_path, "package_test must be PASS"))
    if not any(PurePosixPath(relative).match(pattern) for pattern in allowed):
        findings.append(_finding("unapproved-binary", relative, "asset path is not approved for this profile"))
    target = root / relative
    try:
        resolved = target.resolve(strict=True)
        resolved.relative_to(root.resolve())
    except (OSError, ValueError):
        findings.append(_finding("missing-release-asset", relative, "asset is missing or escapes the repository root"))
        return findings
    if not resolved.is_file():
        findings.append(_finding("missing-release-asset", relative, "asset is not a regular file"))
        return findings
    actual = _sha256(resolved)
    if actual != digest:
        findings.append(_finding("asset-digest-mismatch", relative, f"actual SHA-256 is {actual}"))
    return findings


def _validate_binary_assets(
    root: Path,
    assets: list[object],
    manifest: dict[str, str] | None,
    manifest_path: Path | None,
    policy: dict[str, Any],
    profile: str | None,
    candidate_sha: str,
) -> list[dict[str, str]]:
    findings: list[dict[str, str]] = []
    if profile == "library":
        findings.append(_finding("unapproved-binary", "evidence.distribution", "library profile is source-only by default"))
    if not assets:
        findings.append(_finding("missing-release-assets", "evidence.assets", "binary distribution requires at least one asset"))
    if manifest is None:
        findings.append(_finding("missing-asset-manifest", str(manifest_path or "<not supplied>"), "binary distribution requires a SHA-256 manifest"))
        manifest = {}
    allowed = policy["profiles"].get(profile, {}).get("allowed_asset_globs", []) if profile else []
    evidence_entries: dict[str, str] = {}
    for index, asset in enumerate(assets):
        findings.extend(_validate_asset_record(
            root, asset, f"evidence.assets[{index}]", allowed, candidate_sha, evidence_entries
        ))
    if manifest != evidence_entries:
        findings.append(_finding("asset-manifest-mismatch", str(manifest_path), "manifest entries must exactly match evidence assets"))
    return findings


def _validate_evidence_and_assets(
    root: Path,
    evidence_path: Path,
    manifest_path: Path | None,
    policy: dict[str, Any],
    profile: str | None,
    version: str | None,
    tag: str,
    candidate_sha: str,
) -> list[dict[str, str]]:
    evidence, findings = _read_json(evidence_path, "release-evidence")
    manifest, manifest_findings = _parse_manifest(manifest_path)
    findings.extend(manifest_findings)
    if not isinstance(evidence, dict):
        if evidence is not None:
            findings.append(_finding("invalid-release-evidence", str(evidence_path), "root must be an object"))
        return findings
    findings.extend(_validate_evidence_metadata(
        evidence, evidence_path, policy, profile, version, tag, candidate_sha
    ))
    findings.extend(_validate_evidence_checks(evidence, policy, profile, candidate_sha))
    assets = evidence.get("assets")
    if not isinstance(assets, list):
        findings.append(_finding("invalid-release-evidence", "evidence.assets", "must be an array"))
        return findings
    distribution = evidence.get("distribution")
    if distribution == "source-only":
        if assets:
            findings.append(_finding("source-only-has-assets", "evidence.assets", "source-only releases cannot declare binary assets"))
        if manifest:
            findings.append(_finding("source-only-has-assets", str(manifest_path), "source-only manifest must be empty or omitted"))
        return findings
    if distribution == "binary":
        findings.extend(_validate_binary_assets(
            root, assets, manifest, manifest_path, policy, profile, candidate_sha
        ))
    return findings


def build_report(
    root: Path,
    tag: str,
    candidate_sha: str,
    evidence_path: Path,
    manifest_path: Path | None,
    require_tag_ref: bool,
    provenance_path: Path | None = None,
    signature_path: Path | None = None,
    excel_evidence_path: Path | None = None,
) -> dict[str, Any]:
    findings: list[dict[str, str]] = []
    configuration, configuration_findings = _load_configuration(root)
    policy, policy_findings = _load_policy(root)
    findings.extend(configuration_findings)
    findings.extend(policy_findings)
    if COMMIT_PATTERN.fullmatch(candidate_sha) is None:
        findings.append(_finding("invalid-candidate-sha", candidate_sha, "must be 40 lowercase hexadecimal characters"))
    version: str | None = None
    profile: str | None = None
    if configuration is not None and policy is not None and COMMIT_PATTERN.fullmatch(candidate_sha):
        version, version_findings = _validate_version_and_changelog(root, tag)
        findings.extend(version_findings)
        profile, source_findings = _validate_source(root, configuration, policy, candidate_sha, tag, require_tag_ref)
        findings.extend(source_findings)
        findings.extend(_validate_evidence_and_assets(
            root, evidence_path, manifest_path, policy, profile, version, tag, candidate_sha
        ))
        findings.extend(validate_provenance(
            root, configuration, candidate_sha, evidence_path, manifest_path,
            provenance_path, signature_path,
        ))
        findings.extend(excel_release_findings(root, candidate_sha, evidence_path, excel_evidence_path))
    findings.sort(key=lambda item: (item["code"], item["path"], item["message"]))
    return {
        "schema_version": SCHEMA_VERSION,
        "tool": TOOL_NAME,
        "status": "pass" if not findings else "fail",
        "candidate_sha": candidate_sha,
        "tag": tag,
        "version": version,
        "profile": profile,
        "tag_ref_required": require_tag_ref,
        "counts": {"findings": len(findings)},
        "findings": findings,
        "scope_note": (
            "Validates release metadata, candidate/tag identity, declared evidence, "
            "staged asset bytes, and manifest integrity; it does not itself run Excel "
            "or prove that a manually built Office binary was produced from source."
        ),
    }


def console_report(report: dict[str, Any]) -> str:
    lines = []
    for item in report["findings"]:
        lines.append(f"[FAIL] {item['code']}: {item['path']}: {item['message']}")
    lines.append(
        f"{str(report['status']).upper()}: release {report['tag']} at "
        f"{report['candidate_sha']}; {report['counts']['findings']} finding(s)."
    )
    lines.append(f"Scope: {report['scope_note']}")
    return "\n".join(lines)


def markdown_report(report: dict[str, Any]) -> str:
    status = str(report["status"]).upper()
    lines = [
        "# Release-integrity validation", "",
        f"- Result: **{status}**",
        f"- Tag: `{report['tag']}`",
        f"- Candidate: `{report['candidate_sha']}`",
        f"- Profile: `{report['profile'] or 'unresolved'}`",
        f"- Findings: **{report['counts']['findings']}**", "",
    ]
    if report["findings"]:
        lines.extend(["| Code | Path | Finding |", "| --- | --- | --- |"])
        for item in report["findings"]:
            message = str(item["message"]).replace("|", "\\|")
            lines.append(f"| `{item['code']}` | `{item['path']}` | {message} |")
        lines.append("")
    lines.extend(["> Scope boundary", ">", f"> {report['scope_note']}", ""])
    return "\n".join(lines)


def _write_atomic(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    temporary = Path(name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as stream:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def _fixture_configuration(profile: str) -> dict[str, Any]:
    return {
        "schema_version": 1,
        "mode": "template" if profile == "template" else "generated",
        "profile": None if profile == "template" else profile,
        "repository": (
            "example/TEMPLATE-IDENTITY"
            if profile == "template"
            else f"example/release-{profile}"
        ),
        "identity": {
            "template_tokens": ["TEMPLATE-IDENTITY"],
            "exclude_paths": [PROFILE_PATH],
        },
    }


def _fixture_evidence(
    profile: str,
    sha: str,
    policy: dict[str, Any],
    *,
    distribution: str = "source-only",
    assets: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    checks: dict[str, dict[str, Any]] = {}
    for check_id in policy["core_checks"] + policy["profiles"][profile]["required_checks"]:
        checks[check_id] = {
            "status": "PASS",
            "candidate_sha": sha,
            "detail": f"Synthetic {check_id} evidence.",
        }
    checks["repository-integrity"]["run_url"] = "https://example.invalid/actions/runs/1"
    checks["vba-compile"]["environment"] = "Microsoft Excel 16.0; Windows 64-bit; Office 64-bit"
    checks["regression"].update({
        "entry_point": "ProjectTests.RunProjectTests",
        "environment": "Microsoft Excel 16.0; Windows 64-bit; Office 64-bit",
        "cases": 4,
        "assertions": 6,
        "failures": 0,
        "completeness": "COMPLETE",
        "cleanup": "PASS",
    })
    return {
        "schema_version": 1,
        "version": "1.0.0",
        "tag": "v1.0.0",
        "candidate_sha": sha,
        "profile": profile,
        "distribution": distribution,
        "checks": checks,
        "assets": assets or [],
    }


def _fixture_repository(root: Path, profile: str, policy: dict[str, Any]) -> str:
    (root / ".github").mkdir(parents=True)
    (root / "src").mkdir()
    (root / PROFILE_PATH).write_text(json.dumps(_fixture_configuration(profile), indent=2) + "\n", encoding="utf-8")
    (root / POLICY_PATH).write_text(json.dumps(policy, indent=2) + "\n", encoding="utf-8")
    if profile != "template":
        record = {
            "schema_version": 1,
            "profile": profile,
            "values": {"REPOSITORY_PATH": f"example/release-{profile}"},
        }
        (root / INITIALIZATION_PATH).write_text(
            json.dumps(record, indent=2) + "\n", encoding="utf-8"
        )
    (root / "VERSION").write_text("1.0.0\n", encoding="utf-8")
    release_entry = (
        "- Canonical template construction history with {{PROJECT_NAME}} and TEMPLATE-IDENTITY.\n"
        if profile == "template"
        else "- Initial project release.\n"
    )
    (root / "CHANGELOG.md").write_text(
        "# Changelog\n\n## [Unreleased]\n\nNo unreleased changes.\n\n"
        "## [1.0.0] - 2026-09-04\n\n### Added\n\n" + release_entry,
        encoding="utf-8",
    )
    readme = (
        "# {{PROJECT_NAME}} template fixture\n\nTEMPLATE-IDENTITY\n"
        if profile == "template"
        else "# Synthetic release fixture\n"
    )
    (root / "README.md").write_text(readme, encoding="utf-8")
    (root / "src" / "Project.bas").write_text(
        'Attribute VB_Name = "Project"\nOption Explicit\n', encoding="utf-8"
    )
    _git(root, "init", "-b", "main")
    _git(root, "config", "user.name", "Release Self-Test")
    _git(root, "config", "user.email", "release-self-test@example.invalid")
    _git(root, "add", "--all")
    committed = _git(root, "commit", "-m", "Create release fixture")
    if committed.returncode != 0:
        raise OperationalError(committed.stderr.strip())
    sha = _git_output(root, "rev-parse", "HEAD")
    if sha is None:
        raise OperationalError("Could not resolve fixture commit")
    tagged = _git(root, "tag", "-a", "v1.0.0", "-m", "Synthetic release 1.0.0")
    if tagged.returncode != 0:
        raise OperationalError(tagged.stderr.strip())
    return sha


def _tree_digest(root: Path) -> str:
    digest = hashlib.sha256()
    for path in sorted(item for item in root.rglob("*") if item.is_file() and ".git" not in item.parts):
        digest.update(path.relative_to(root).as_posix().encode("utf-8"))
        digest.update(b"\0")
        digest.update(path.read_bytes())
        digest.update(b"\0")
    return digest.hexdigest()


def _run_self_test(root: Path, summary_path: Path | None) -> int:
    policy, findings = _load_policy(root)
    if policy is None:
        print("SELF-TEST FAIL: canonical release policy is invalid.", file=sys.stderr)
        for item in findings:
            print(f"  {item['code']}: {item['message']}", file=sys.stderr)
        return 1
    results: list[tuple[str, str, str]] = []
    with tempfile.TemporaryDirectory(prefix="release-self-test-") as temporary_name:
        temporary = Path(temporary_name)

        for profile in SUPPORTED_PROFILES:
            case = temporary / f"positive-{profile}"
            case.mkdir()
            sha = _fixture_repository(case, profile, policy)
            evidence = temporary / f"positive-{profile}.json"
            manifest: Path | None = None
            distribution = "source-only"
            assets: list[dict[str, Any]] = []
            if profile in {"application", "ui-component"}:
                distribution = "binary"
                asset = case / "dist" / f"fixture-{profile}.xlsm"
                asset.parent.mkdir()
                asset.write_bytes(f"synthetic {profile} asset\n".encode("utf-8"))
                digest = _sha256(asset)
                relative = asset.relative_to(case).as_posix()
                assets = [{
                    "path": relative,
                    "sha256": digest,
                    "candidate_sha": sha,
                    "package_test": "PASS",
                }]
                manifest = temporary / f"positive-{profile}.sha256"
                manifest.write_text(f"{digest}  {relative}\n", encoding="utf-8")
            evidence.write_text(
                json.dumps(_fixture_evidence(profile, sha, policy, distribution=distribution, assets=assets), indent=2) + "\n",
                encoding="utf-8",
            )
            before = _tree_digest(case)
            first = build_report(case, "v1.0.0", sha, evidence, manifest, True)
            second = build_report(case, "v1.0.0", sha, evidence, manifest, True)
            after = _tree_digest(case)
            passed = first["status"] == "pass" and first == second and before == after
            results.append((f"valid-{profile}", "accepted", "PASS" if passed else "FAIL"))

        def negative(name: str, mutate, expected_code: str, profile: str = "library") -> None:
            case = temporary / name
            case.mkdir()
            sha = _fixture_repository(case, profile, policy)
            evidence = temporary / f"{name}.json"
            manifest = temporary / f"{name}.sha256"
            evidence_data = _fixture_evidence(profile, sha, policy)
            evidence.write_text(json.dumps(evidence_data, indent=2) + "\n", encoding="utf-8")
            manifest_path: Path | None = None
            context: dict[str, Any] = {
                "root": case, "sha": sha, "evidence": evidence,
                "evidence_data": evidence_data, "manifest": manifest,
                "manifest_path": manifest_path, "tag": "v1.0.0",
            }
            mutate(context)
            before = _tree_digest(case)
            report = build_report(
                case, context["tag"], sha, context["evidence"],
                context.get("manifest_path"), True,
            )
            after = _tree_digest(case)
            codes = {item["code"] for item in report["findings"]}
            passed = report["status"] == "fail" and expected_code in codes and before == after
            results.append((name, f"rejected ({expected_code})", "PASS" if passed else "FAIL"))

        negative("version-tag-mismatch", lambda c: c.update(tag="v1.0.1"), "tag-version-mismatch")
        negative(
            "zero-version",
            lambda c: (c["root"] / "VERSION").write_text("0.0.0\n", encoding="utf-8"),
            "zero-version",
        )
        negative(
            "invalid-changelog-date",
            lambda c: (c["root"] / "CHANGELOG.md").write_text(
                (c["root"] / "CHANGELOG.md").read_text(encoding="utf-8").replace("2026-09-04", "2026-02-30"),
                encoding="utf-8",
            ),
            "invalid-changelog-date",
        )
        negative(
            "missing-changelog-release",
            lambda c: (c["root"] / "CHANGELOG.md").write_text("# Changelog\n\n## [Unreleased]\n", encoding="utf-8"),
            "missing-changelog-release",
        )
        negative(
            "unresolved-token",
            lambda c: (c["root"] / "README.md").write_text("# {{PROJECT_NAME}}\n", encoding="utf-8"),
            "unresolved-template-token",
        )
        negative(
            "template-identity",
            lambda c: (c["root"] / PROFILE_PATH).write_text(
                json.dumps({**_fixture_configuration("library"), "mode": "template", "profile": None}, indent=2) + "\n",
                encoding="utf-8",
            ),
            "template-identity",
        )
        negative(
            "construction-history",
            lambda c: (c["root"] / "CHANGELOG.md").write_text(
                (c["root"] / "CHANGELOG.md").read_text(encoding="utf-8") + "\nPortfolio audit carried forward.\n",
                encoding="utf-8",
            ),
            "template-construction-history",
        )
        negative(
            "missing-evidence",
            lambda c: c.update(evidence=c["evidence"].with_name("absent.json")),
            "missing-release-evidence",
        )
        negative(
            "missing-profile-evidence",
            lambda c: (
                c["evidence_data"]["checks"].pop("public-api"),
                c["evidence"].write_text(json.dumps(c["evidence_data"], indent=2) + "\n", encoding="utf-8"),
            ),
            "missing-profile-evidence",
        )
        negative(
            "evidence-sha-mismatch",
            lambda c: (
                c["evidence_data"].update(candidate_sha="f" * 40),
                c["evidence"].write_text(json.dumps(c["evidence_data"], indent=2) + "\n", encoding="utf-8"),
            ),
            "evidence-sha-mismatch",
        )

        def binary_mutation(context, *, approved: bool, digest_matches: bool) -> None:
            asset = context["root"] / "dist" / ("fixture.xlsm" if approved else "fixture.exe")
            asset.parent.mkdir()
            asset.write_bytes(b"synthetic release asset\n")
            actual = _sha256(asset)
            declared = actual if digest_matches else "0" * 64
            relative = asset.relative_to(context["root"]).as_posix()
            context["evidence_data"]["distribution"] = "binary"
            context["evidence_data"]["assets"] = [{
                "path": relative,
                "sha256": declared,
                "candidate_sha": context["sha"],
                "package_test": "PASS",
            }]
            context["evidence"].write_text(json.dumps(context["evidence_data"], indent=2) + "\n", encoding="utf-8")
            context["manifest"].write_text(f"{declared}  {relative}\n", encoding="utf-8")
            context["manifest_path"] = context["manifest"]

        negative(
            "unapproved-binary", lambda c: binary_mutation(c, approved=False, digest_matches=True),
            "unapproved-binary", profile="application",
        )
        negative(
            "incorrect-digest", lambda c: binary_mutation(c, approved=True, digest_matches=False),
            "asset-digest-mismatch", profile="application",
        )
        def missing_manifest_mutation(context) -> None:
            binary_mutation(context, approved=True, digest_matches=True)
            context.update(manifest_path=None)

        negative(
            "missing-asset-manifest",
            missing_manifest_mutation,
            "missing-asset-manifest", profile="application",
        )
        def asset_binding_mutation(context) -> None:
            binary_mutation(context, approved=True, digest_matches=True)
            context["evidence_data"]["assets"][0]["candidate_sha"] = "f" * 40
            context["evidence"].write_text(
                json.dumps(context["evidence_data"], indent=2) + "\n", encoding="utf-8"
            )

        negative(
            "asset-sha-binding", asset_binding_mutation,
            "asset-sha-binding-mismatch", profile="application",
        )
        negative(
            "library-binary", lambda c: binary_mutation(c, approved=True, digest_matches=True),
            "unapproved-binary", profile="library",
        )
        negative(
            "template-binary", lambda c: binary_mutation(c, approved=True, digest_matches=True),
            "unapproved-binary", profile="template",
        )

        def lightweight_tag(context) -> None:
            _git(context["root"], "tag", "-d", "v1.0.0")
            _git(context["root"], "tag", "v1.0.0")

        negative("lightweight-tag", lightweight_tag, "lightweight-tag")

        def moved_tag(context) -> None:
            _git(context["root"], "tag", "-d", "v1.0.0")
            _git(context["root"], "commit", "--allow-empty", "-m", "Move release target")
            _git(context["root"], "tag", "-a", "v1.0.0", "-m", "Moved release")

        negative("moved-tag", moved_tag, "tag-target-mismatch")

        failures = [name for name, _, result in results if result != "PASS"]
        lines = [
            "# Release-integrity self-test", "",
            f"- Positive fixtures: **{len(SUPPORTED_PROFILES)}**",
            f"- Negative fixtures: **{len(results) - len(SUPPORTED_PROFILES)}**", "",
            "| Fixture | Expected | Result |", "| --- | --- | --- |",
        ]
        lines.extend(f"| {name} | {expected} | {result} |" for name, expected, result in results)
        lines.append("")
        if failures:
            lines.append("FAIL: " + ", ".join(failures))
        else:
            lines.append(
                f"PASS: {len(results)} deterministic release fixtures satisfied; all executions were read-only."
            )
        report_text = "\n".join(lines) + "\n"
        print(report_text, end="")
        if summary_path is not None:
            _write_atomic(summary_path, report_text)
        return 1 if failures else 0


def parse_arguments(arguments: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path("."), help="repository root")
    parser.add_argument("--tag", help="candidate tag, for example v1.0.0")
    parser.add_argument("--candidate-sha", help="full 40-character candidate commit SHA")
    parser.add_argument("--evidence", type=Path, help="external release evidence JSON")
    parser.add_argument("--asset-manifest", type=Path, help="external SHA-256 asset manifest")
    parser.add_argument("--provenance", type=Path, help="external build provenance JSON")
    parser.add_argument("--provenance-signature", type=Path, help="detached SSH provenance signature")
    parser.add_argument("--excel-evidence", type=Path, help="optional retained host evidence JSON")
    parser.add_argument("--require-tag-ref", action="store_true", help="require an annotated local tag resolving to the candidate")
    parser.add_argument("--output", type=Path, help="write deterministic JSON report")
    parser.add_argument("--summary", type=Path, help="write Markdown report or self-test summary")
    parser.add_argument("--self-test", action="store_true", help="exercise positive and negative synthetic candidates")
    parsed = parser.parse_args(arguments)
    if not parsed.self_test:
        missing = [name for name in ("tag", "candidate_sha", "evidence") if getattr(parsed, name) is None]
        if missing:
            parser.error("candidate validation requires --" + ", --".join(name.replace("_", "-") for name in missing))
    return parsed


def main(arguments: list[str] | None = None) -> int:
    parsed = parse_arguments(sys.argv[1:] if arguments is None else arguments)
    root = parsed.root.resolve()
    try:
        if parsed.self_test:
            return _run_self_test(root, parsed.summary)
        evidence = parsed.evidence
        if not evidence.is_absolute():
            evidence = (Path.cwd() / evidence).resolve()
        manifest = parsed.asset_manifest
        if manifest is not None and not manifest.is_absolute():
            manifest = (Path.cwd() / manifest).resolve()
        report = build_report(
            root, parsed.tag, parsed.candidate_sha, evidence, manifest, parsed.require_tag_ref,
            parsed.provenance, parsed.provenance_signature, parsed.excel_evidence,
        )
        json_text = json.dumps(report, indent=2, sort_keys=True) + "\n"
        print(console_report(report))
        if parsed.output is not None:
            _write_atomic(parsed.output, json_text)
        if parsed.summary is not None:
            _write_atomic(parsed.summary, markdown_report(report))
        return 0 if report["status"] == "pass" else 1
    except OperationalError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
