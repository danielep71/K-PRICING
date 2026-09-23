#!/usr/bin/env python3
"""Build or verify deterministic, candidate-bound release certification bundles."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path, PurePosixPath
from typing import Any

from _gatelib import git_text
from release_provenance import github_signing_keys, policy_for

TOOL_VERSION = "1.2.0"
SCHEMA_VERSION = 1
REQUIRED_ROLES = {
    "excel-evidence",
    "external-links",
    "gate-evidence",
    "release-evidence",
    "release-integrity",
    "wiki-publication",
}
SHA256_RE = re.compile(r"[0-9a-f]{64}$")
SHA40_RE = re.compile(r"[0-9a-f]{40}$")
VERSION_RE = re.compile(r"(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$")
TAG_RE = re.compile(r"v(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$")
SAFE_ID_RE = re.compile(r"[a-z0-9]+(?:-[a-z0-9]+)*$")
CERTIFICATION_NAMESPACE_RE = re.compile(
    r"certification-v[^/]+(?:\.zip(?:\.sig)?|\.manifest\.json|\.sha256)$"
)
CERTIFICATION_SIGNATURE_NAMESPACE = "excel-vba-release-certification"
SECRET_PATTERNS = (
    re.compile(rb"-----BEGIN [A-Z ]*PRIVATE KEY-----"),
    re.compile(rb"\b" + b"gh" + rb"p_[A-Za-z0-9]{20,}\b"),
    re.compile(rb"\bgithub_pat_[A-Za-z0-9_]{20,}\b"),
    re.compile(rb"\bAKIA[A-Z0-9]{16}\b"),
)


class CertificationError(RuntimeError):
    """Certification input or retained bundle violates the contract."""


def require(condition: object, message: str) -> None:
    """Fail certification unless ``condition`` is truthy.

    Typed as ``object`` rather than ``bool`` because every call site passes the
    result of a truthiness test, not a coerced boolean.
    """
    if not condition:
        raise CertificationError(message)


def require_str(value: Any, message: str) -> str:
    """Return ``value`` when it is a string, else fail certification.

    Returning the value narrows it for the type checker, so downstream use does
    not depend on a separate ``require`` call the checker cannot see.
    """
    if not isinstance(value, str):
        raise CertificationError(message)
    return value


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def load_json(path: Path) -> dict[str, Any]:
    def unique(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"duplicate JSON key in {path}: {key}")
            result[key] = value
        return result

    try:
        value = json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=unique)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise CertificationError(f"cannot read {path}: {error}") from error
    require(isinstance(value, dict), f"{path} must contain one JSON object")
    return value


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")


def git_output(root: Path, *args: str) -> str:
    """Return trimmed stdout of a git command, or fail the certification.

    Subprocess handling belongs to ``_gatelib``; this wrapper only adds the
    certification-specific failure mode, so the shared helper is not
    reimplemented here.
    """
    completed = git_text(root, *args)
    if completed.returncode:
        raise CertificationError(f"git {' '.join(args)} failed")
    return completed.stdout.strip()


def is_within(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def safe_relative(value: str, field: str) -> str:
    path = PurePosixPath(value)
    require(value == path.as_posix(), f"{field} must use normalized POSIX separators")
    require(not path.is_absolute(), f"{field} must be relative")
    require(value not in {"", "."}, f"{field} must not be empty")
    require(all(part not in {"", ".", ".."} for part in path.parts), f"unsafe {field}: {value}")
    return value


def validate_identity(root: Path, specification: dict[str, Any]) -> dict[str, str]:
    require(specification.get("schema_version") == SCHEMA_VERSION, "unsupported specification schema")
    repository = require_str(specification.get("repository"), "repository must use owner/name form")
    version = require_str(specification.get("version"), "version must be SemVer core")
    tag = require_str(specification.get("tag"), "tag must equal v + version")
    candidate = require_str(specification.get("candidate_sha"), "candidate_sha must be full lowercase hex")
    require(re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repository) is not None, "repository must use owner/name form")
    require(VERSION_RE.fullmatch(version) is not None, "version must be SemVer core")
    require(tag == f"v{version}", "tag must equal v + version")
    require(SHA40_RE.fullmatch(candidate) is not None, "candidate_sha must be full lowercase hex")
    require(git_output(root, "rev-parse", "HEAD") == candidate, "working tree HEAD does not equal candidate_sha")
    require(not git_output(root, "status", "--porcelain"), "candidate working tree must be clean")
    require((root / "VERSION").read_text(encoding="utf-8").strip() == version, "candidate VERSION disagrees with specification")
    profile = load_json(root / ".github/repository-profile.json")
    require(profile.get("repository") == repository, "candidate repository identity disagrees with specification")
    return {
        "repository": repository,
        "version": version,
        "tag": tag,
        "candidate_sha": candidate,
    }


def validate_public_bytes(data: bytes, record_id: str) -> None:
    for pattern in SECRET_PATTERNS:
        require(pattern.search(data) is None, f"public record {record_id} contains secret-like material")


def validate_records(
    specification: dict[str, Any], evidence_dir: Path
) -> tuple[list[dict[str, Any]], dict[str, bytes]]:
    raw_records = specification.get("records")
    if not isinstance(raw_records, list) or not raw_records:
        raise CertificationError("records must be a non-empty array")
    seen_ids: set[str] = set()
    seen_roles: set[str] = set()
    seen_archive_paths: set[str] = set()
    manifest_records: list[dict[str, Any]] = []
    payloads: dict[str, bytes] = {}

    for index, raw in enumerate(raw_records):
        require(isinstance(raw, dict), f"records[{index}] must be an object")
        record_id = raw.get("id")
        role = raw.get("role")
        visibility = raw.get("visibility")
        require(isinstance(record_id, str) and SAFE_ID_RE.fullmatch(record_id) is not None, f"records[{index}].id must be kebab-case")
        require(record_id not in seen_ids, f"duplicate record id: {record_id}")
        require(isinstance(role, str) and SAFE_ID_RE.fullmatch(role) is not None, f"records[{index}].role must be kebab-case")
        require(role not in seen_roles, f"duplicate record role: {role}")
        require(visibility in {"public-file", "restricted-reference"}, f"unsupported visibility for {record_id}")
        seen_ids.add(record_id)
        seen_roles.add(role)

        if visibility == "public-file":
            require(raw.get("redaction_reviewed") is True, f"public record {record_id} must confirm redaction_reviewed=true")
            source = safe_relative(str(raw.get("path") or ""), f"{record_id}.path")
            public_name = safe_relative(str(raw.get("public_name") or ""), f"{record_id}.public_name")
            source_path = evidence_dir / source
            require(source_path.is_file() and not source_path.is_symlink(), f"public record file is unavailable or symlinked: {source}")
            archive_path = f"evidence/{public_name}"
            require(archive_path not in seen_archive_paths, f"duplicate archive path: {archive_path}")
            seen_archive_paths.add(archive_path)
            data = source_path.read_bytes()
            validate_public_bytes(data, record_id)
            payloads[archive_path] = data
            manifest_records.append(
                {
                    "id": record_id,
                    "role": role,
                    "visibility": visibility,
                    "archive_path": archive_path,
                    "sha256": sha256_bytes(data),
                    "size": len(data),
                }
            )
        else:
            digest = raw.get("sha256")
            reference = raw.get("reference")
            reason = raw.get("reason")
            require(isinstance(digest, str) and SHA256_RE.fullmatch(digest) is not None, f"restricted record {record_id} requires a lowercase SHA-256")
            require(isinstance(reference, str) and reference.strip(), f"restricted record {record_id} requires a reference")
            require(isinstance(reason, str) and reason.strip(), f"restricted record {record_id} requires a reason")
            require("\n" not in reference and "\r" not in reference, f"restricted record {record_id} reference must be single-line")
            require("\n" not in reason and "\r" not in reason, f"restricted record {record_id} reason must be single-line")
            manifest_records.append(
                {
                    "id": record_id,
                    "role": role,
                    "visibility": visibility,
                    "sha256": digest,
                    "reference": reference,
                    "reason": reason,
                }
            )

    missing = sorted(REQUIRED_ROLES - seen_roles)
    require(not missing, "required certification roles are missing: " + ", ".join(missing))
    manifest_records.sort(key=lambda row: (str(row["role"]), str(row["id"])))
    return manifest_records, payloads


def validate_manifest_records(raw_records: Any) -> list[dict[str, Any]]:
    """Validate the durable manifest independently from the build specification."""
    if not isinstance(raw_records, list) or not raw_records:
        raise CertificationError("manifest records must be a non-empty array")
    seen_ids: set[str] = set()
    seen_roles: set[str] = set()
    seen_archive_paths: set[str] = set()
    records: list[dict[str, Any]] = []

    for index, raw in enumerate(raw_records):
        require(isinstance(raw, dict), f"manifest records[{index}] must be an object")
        record_id = raw.get("id")
        role = raw.get("role")
        visibility = raw.get("visibility")
        require(
            isinstance(record_id, str) and SAFE_ID_RE.fullmatch(record_id) is not None,
            f"manifest records[{index}].id must be kebab-case",
        )
        require(record_id not in seen_ids, f"duplicate manifest record id: {record_id}")
        require(
            isinstance(role, str) and SAFE_ID_RE.fullmatch(role) is not None,
            f"manifest records[{index}].role must be kebab-case",
        )
        require(role not in seen_roles, f"duplicate manifest record role: {role}")
        require(
            visibility in {"public-file", "restricted-reference"},
            f"unsupported manifest visibility for {record_id}",
        )
        seen_ids.add(record_id)
        seen_roles.add(role)

        digest = raw.get("sha256")
        require(
            isinstance(digest, str) and SHA256_RE.fullmatch(digest) is not None,
            f"manifest record {record_id} requires a lowercase SHA-256",
        )
        if visibility == "public-file":
            archive_path_value = raw.get("archive_path")
            require(
                isinstance(archive_path_value, str),
                f"public manifest record {record_id} requires archive_path",
            )
            archive_path = safe_relative(
                archive_path_value,
                f"manifest record {record_id}.archive_path",
            )
            require(
                archive_path.startswith("evidence/"),
                f"public manifest record {record_id} must be stored under evidence/",
            )
            require(
                archive_path not in seen_archive_paths,
                f"duplicate manifest archive path: {archive_path}",
            )
            seen_archive_paths.add(archive_path)
            size = raw.get("size")
            require(
                isinstance(size, int) and not isinstance(size, bool) and size >= 0,
                f"public manifest record {record_id} requires a non-negative size",
            )
        else:
            reference = raw.get("reference")
            reason = raw.get("reason")
            require(
                isinstance(reference, str) and reference.strip(),
                f"restricted manifest record {record_id} requires a reference",
            )
            require(
                isinstance(reason, str) and reason.strip(),
                f"restricted manifest record {record_id} requires a reason",
            )
            require(
                "\n" not in reference and "\r" not in reference,
                f"restricted manifest record {record_id} reference must be single-line",
            )
            require(
                "\n" not in reason and "\r" not in reason,
                f"restricted manifest record {record_id} reason must be single-line",
            )
        records.append(raw)

    missing = sorted(REQUIRED_ROLES - seen_roles)
    require(
        not missing,
        "required certification roles are missing from manifest: " + ", ".join(missing),
    )
    return records


def zip_bytes(entries: dict[str, bytes]) -> bytes:
    import io

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_STORED) as archive:
        for name in sorted(entries):
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_STORED
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            archive.writestr(info, entries[name])
    return buffer.getvalue()


def build_bundle(root: Path, specification_path: Path, evidence_dir: Path, output_dir: Path) -> dict[str, Any]:
    root = root.resolve()
    specification_path = specification_path.resolve()
    evidence_dir = evidence_dir.resolve()
    output_dir = output_dir.resolve()
    require(root.is_dir(), "candidate root is unavailable")
    require(evidence_dir.is_dir() and not evidence_dir.is_symlink(), "evidence directory is unavailable or symlinked")
    require(not is_within(specification_path, root), "certification specification must stay outside the candidate tree")
    require(not is_within(evidence_dir, root), "certification evidence must stay outside the candidate tree")
    require(not is_within(output_dir, root), "certification output must stay outside the candidate tree")

    specification = load_json(specification_path)
    identity = validate_identity(root, specification)
    records, payloads = validate_records(specification, evidence_dir)
    base = f"certification-{identity['tag']}"
    bundle_name = base + ".zip"
    manifest_name = base + ".manifest.json"
    digest_name = base + ".sha256"
    tool_path = Path(__file__).resolve()
    manifest = {
        "schema_version": SCHEMA_VERSION,
        **identity,
        "distribution": "certification-evidence",
        "bundle_name": bundle_name,
        "builder": {
            "tool": "tools/release_certification.py",
            "version": TOOL_VERSION,
            "sha256": sha256_bytes(tool_path.read_bytes()),
        },
        "records": records,
        "scope_note": (
            "Certification evidence is retained separately from source/runtime distribution. "
            "restricted-reference records disclose identity/digest/reason without publishing bytes."
        ),
    }
    manifest_bytes = canonical_json(manifest)
    archive_entries = dict(payloads)
    archive_entries["certification-manifest.json"] = manifest_bytes
    bundle = zip_bytes(archive_entries)
    bundle_digest = sha256_bytes(bundle)
    digest_bytes = f"{bundle_digest}  {bundle_name}\n".encode("ascii")

    output_dir.mkdir(parents=True, exist_ok=True)
    for name in (bundle_name, manifest_name, digest_name):
        require(not (output_dir / name).exists(), f"refusing to overwrite existing output: {name}")
    (output_dir / bundle_name).write_bytes(bundle)
    (output_dir / manifest_name).write_bytes(manifest_bytes)
    (output_dir / digest_name).write_bytes(digest_bytes)
    return {
        "bundle": str(output_dir / bundle_name),
        "manifest": str(output_dir / manifest_name),
        "digest": str(output_dir / digest_name),
        "bundle_sha256": bundle_digest,
        "records": len(records),
    }


def verify_bundle(bundle: Path, manifest_path: Path, digest_path: Path) -> dict[str, Any]:
    bundle = bundle.resolve()
    manifest_path = manifest_path.resolve()
    digest_path = digest_path.resolve()
    require(bundle.is_file() and not bundle.is_symlink(), "bundle is unavailable or symlinked")
    require(manifest_path.is_file() and not manifest_path.is_symlink(), "manifest is unavailable or symlinked")
    require(digest_path.is_file() and not digest_path.is_symlink(), "digest file is unavailable or symlinked")
    manifest = load_json(manifest_path)
    require(manifest.get("schema_version") == SCHEMA_VERSION, "unsupported manifest schema")
    require(manifest.get("distribution") == "certification-evidence", "manifest distribution is not certification-evidence")
    require(manifest.get("bundle_name") == bundle.name, "manifest bundle_name mismatches bundle")
    expected_line = digest_path.read_text(encoding="ascii")
    match = re.fullmatch(r"([0-9a-f]{64})  ([^\r\n]+)\n", expected_line)
    if match is None:
        raise CertificationError("digest file must contain one canonical SHA-256 line")
    require(match.group(2) == bundle.name, "digest file names another bundle")
    actual_digest = sha256_bytes(bundle.read_bytes())
    require(match.group(1) == actual_digest, "bundle SHA-256 mismatch")

    records = validate_manifest_records(manifest.get("records"))
    expected_entries = {"certification-manifest.json": canonical_json(manifest)}
    for raw in records:
        if raw.get("visibility") != "public-file":
            continue
        archive_path = raw["archive_path"]
        expected_entries[str(archive_path)] = b""

    with zipfile.ZipFile(bundle, "r") as archive:
        names = archive.namelist()
        require(names == sorted(names), "bundle entries must be sorted")
        require(set(names) == set(expected_entries), "bundle entry set disagrees with manifest")
        require(archive.read("certification-manifest.json") == canonical_json(manifest), "internal manifest differs from published manifest")
        public_by_path = {
            str(raw["archive_path"]): raw
            for raw in records
            if raw.get("visibility") == "public-file"
        }
        for name, raw in public_by_path.items():
            data = archive.read(name)
            require(raw.get("sha256") == sha256_bytes(data), f"digest mismatch for {name}")
            require(raw.get("size") == len(data), f"size mismatch for {name}")
    return {
        "status": "pass",
        "bundle": bundle.name,
        "bundle_sha256": actual_digest,
        "candidate_sha": manifest.get("candidate_sha"),
        "tag": manifest.get("tag"),
        "records": len(records),
    }


def verify_ssh_signature(
    data: bytes,
    signature_path: Path,
    principal: str,
    public_keys: list[str],
) -> None:
    require(signature_path.is_file() and not signature_path.is_symlink(),
            "certification signature is unavailable or symlinked")
    require(principal.strip(), "certification signature principal must be nonempty")
    require(public_keys, "certification signature requires at least one trusted public key")
    with tempfile.TemporaryDirectory(prefix="certification-signature-") as directory:
        signers = Path(directory) / "allowed_signers"
        signers.write_text(
            "".join(f"{principal} {key.strip()}\n" for key in sorted(set(public_keys))),
            encoding="utf-8",
            newline="\n",
        )
        completed = subprocess.run(
            [
                "ssh-keygen", "-Y", "verify",
                "-f", str(signers),
                "-I", principal,
                "-n", CERTIFICATION_SIGNATURE_NAMESPACE,
                "-s", str(signature_path),
            ],
            input=data,
            capture_output=True,
            check=False,
            timeout=30,
        )
        require(completed.returncode == 0, "certification SSH signature verification failed")


def verify_certification_signature(
    root: Path,
    candidate_sha: str,
    bundle: Path,
    signature_path: Path,
) -> dict[str, Any]:
    require(SHA40_RE.fullmatch(candidate_sha) is not None,
            "candidate_sha must be full lowercase hex")
    root = root.resolve()
    require(git_output(root, "rev-parse", "HEAD") == candidate_sha,
            "signature verification root does not equal candidate_sha")
    configuration = load_json(root / ".github/repository-profile.json")
    if configuration.get("mode") != "template":
        return {
            "status": "not-applicable",
            "detail": "durable certification signing is required only for the canonical template",
        }
    try:
        policy = policy_for(root, candidate_sha, configuration)
        trust = policy["tag_signature"]["template"]
        require(trust.get("mode") == "ssh-github",
                "canonical certification signing requires ssh-github tag trust")
        principal = require_str(trust.get("principal"),
                                "canonical certification signing requires a principal")
        github_user = require_str(trust.get("github_user"),
                                  "canonical certification signing requires a GitHub user")
        public_keys = github_signing_keys(github_user)
    except (ValueError, KeyError, TypeError) as error:
        raise CertificationError(f"cannot resolve certification signing trust: {error}") from error

    verify_ssh_signature(bundle.read_bytes(), signature_path, principal, public_keys)
    return {
        "status": "pass",
        "namespace": CERTIFICATION_SIGNATURE_NAMESPACE,
        "principal": principal,
        "github_user": github_user,
        "signature": signature_path.name,
    }


def certification_asset_names(tag: str) -> dict[str, str]:
    require(TAG_RE.fullmatch(tag) is not None, "release tag must be v + SemVer core")
    base = f"certification-{tag}"
    return {
        "bundle": base + ".zip",
        "manifest": base + ".manifest.json",
        "digest": base + ".sha256",
        "signature": base + ".zip.sig",
    }


def plan_release_assets(release_path: Path, tag: str) -> dict[str, Any]:
    release = load_json(release_path)
    require(release.get("tag_name") == tag, "GitHub Release tag disagrees with requested tag")
    raw_assets = release.get("assets")
    if not isinstance(raw_assets, list):
        raise CertificationError("GitHub Release assets must be an array")
    expected = certification_asset_names(tag)
    expected_names = set(expected.values())
    seen_names: set[str] = set()
    certification: dict[str, dict[str, Any]] = {}
    product_assets: list[dict[str, Any]] = []

    for index, raw in enumerate(raw_assets):
        require(isinstance(raw, dict), f"release assets[{index}] must be an object")
        name = require_str(raw.get("name"), f"release assets[{index}].name must be a string")
        require(name not in seen_names, f"duplicate GitHub Release asset name: {name}")
        seen_names.add(name)
        if name in expected_names:
            api_url = require_str(raw.get("url"), f"release asset {name} requires an API URL")
            browser_url = require_str(
                raw.get("browser_download_url"),
                f"release asset {name} requires a browser download URL",
            )
            kind = next(key for key, value in expected.items() if value == name)
            asset_id = raw.get("id")
            require(isinstance(asset_id, int) and asset_id > 0, f"release asset {name} requires an asset id")
            certification[kind] = {
                "name": name,
                "asset_id": asset_id,
                "api_url": api_url,
                "browser_download_url": browser_url,
                "size": raw.get("size"),
                "provider_digest": raw.get("digest"),
            }
        elif CERTIFICATION_NAMESPACE_RE.fullmatch(name) is not None:
            raise CertificationError(f"unexpected certification evidence asset: {name}")
        else:
            product_assets.append(raw)

    kinds = ("bundle", "manifest", "digest", "signature")
    missing = [expected[kind] for kind in kinds if kind not in certification]
    require(not missing, "required certification evidence assets are missing: " + ", ".join(missing))
    ordered = [certification[kind] for kind in kinds]
    return {
        "schema_version": SCHEMA_VERSION,
        "status": "ready",
        "tag": tag,
        "retention_surface": "github-release-assets",
        "retention_lifetime": "published-release-lifetime",
        "certification_assets": ordered,
        "product_assets": product_assets,
    }


def verify_release_plan(
    root: Path,
    plan_path: Path,
    download_dir: Path,
    candidate_sha: str,
) -> dict[str, Any]:
    plan = load_json(plan_path)
    require(plan.get("schema_version") == SCHEMA_VERSION, "unsupported release-asset plan schema")
    require(plan.get("status") == "ready", "release-asset plan is not ready")
    tag = require_str(plan.get("tag"), "release-asset plan requires a tag")
    require(SHA40_RE.fullmatch(candidate_sha) is not None, "candidate_sha must be full lowercase hex")
    require(plan.get("retention_surface") == "github-release-assets", "unsupported retention surface")
    assets = plan.get("certification_assets")
    if not isinstance(assets, list) or len(assets) != 4:
        raise CertificationError("release-asset plan must contain four certification assets")
    by_name: dict[str, dict[str, Any]] = {}
    for raw in assets:
        require(isinstance(raw, dict), "certification asset plan entry must be an object")
        name = require_str(raw.get("name"), "certification asset plan entry requires a name")
        require(name not in by_name, f"duplicate certification asset plan entry: {name}")
        by_name[name] = raw
    names = certification_asset_names(tag)
    require(set(by_name) == set(names.values()), "release-asset plan names do not match the tag")
    download_dir = download_dir.resolve()
    bundle = download_dir / names["bundle"]
    manifest = download_dir / names["manifest"]
    digest = download_dir / names["digest"]
    signature = download_dir / names["signature"]
    verified = verify_bundle(bundle, manifest, digest)
    require(verified.get("candidate_sha") == candidate_sha, "published bundle candidate SHA mismatch")
    require(verified.get("tag") == tag, "published bundle tag mismatch")
    signature_result = verify_certification_signature(
        root,
        candidate_sha,
        bundle,
        signature,
    )
    return {
        **verified,
        "signature_verification": signature_result,
        "retention_surface": plan["retention_surface"],
        "retention_lifetime": plan.get("retention_lifetime"),
        "assets": [
            {
                "name": name,
                "browser_download_url": by_name[name].get("browser_download_url"),
            }
            for name in names.values()
        ],
    }


def fixture_git(root: Path, *args: str) -> str:
    completed = subprocess.run(
        ["git", "-C", str(root), "-c", "user.name=Fixture", "-c", "user.email=fixture@example.invalid", *args],
        check=True,
        capture_output=True,
        text=True,
    )
    return completed.stdout.strip()


def exercise_manifest_verification_controls(
    first_bundle: Path,
    manifest_path: Path,
    digest_path: Path,
    base: Path,
    failures: list[str],
) -> None:
    with zipfile.ZipFile(first_bundle, "r") as original_archive:
        original_payloads = {
            name: original_archive.read(name)
            for name in original_archive.namelist()
            if name != "certification-manifest.json"
        }

    def assert_manifest_mutation_rejected(name: str, mutate: Any) -> None:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        mutate(manifest)
        manifest_bytes = canonical_json(manifest)
        entries = {"certification-manifest.json": manifest_bytes}
        raw_manifest_records = manifest.get("records")
        if isinstance(raw_manifest_records, list):
            for row in raw_manifest_records:
                if not isinstance(row, dict) or row.get("visibility") != "public-file":
                    continue
                archive_path = row.get("archive_path")
                if isinstance(archive_path, str) and archive_path in original_payloads:
                    entries[archive_path] = original_payloads[archive_path]
        mutated_dir = base / ("manifest-" + name)
        mutated_dir.mkdir()
        mutated_bundle = mutated_dir / first_bundle.name
        mutated_manifest = mutated_dir / manifest_path.name
        mutated_digest = mutated_dir / digest_path.name
        bundle_bytes = zip_bytes(entries)
        mutated_bundle.write_bytes(bundle_bytes)
        mutated_manifest.write_bytes(manifest_bytes)
        mutated_digest.write_text(
            f"{sha256_bytes(bundle_bytes)}  {first_bundle.name}\n",
            encoding="ascii",
        )
        try:
            verify_bundle(mutated_bundle, mutated_manifest, mutated_digest)
            failures.append(f"{name} manifest was accepted")
        except CertificationError:
            pass

    assert_manifest_mutation_rejected(
        "empty-required-roles",
        lambda manifest: manifest.__setitem__("records", []),
    )

    def duplicate_role(manifest: dict[str, Any]) -> None:
        duplicate = dict(manifest["records"][0])
        duplicate["id"] = "duplicate-record"
        manifest["records"].append(duplicate)

    assert_manifest_mutation_rejected("duplicate-role", duplicate_role)

    def invalid_visibility(manifest: dict[str, Any]) -> None:
        manifest["records"][-1]["visibility"] = "unknown"

    assert_manifest_mutation_rejected("invalid-visibility", invalid_visibility)

    def invalid_restricted_structure(manifest: dict[str, Any]) -> None:
        manifest["records"][-1].pop("sha256", None)

    assert_manifest_mutation_rejected(
        "invalid-restricted-structure",
        invalid_restricted_structure,
    )


def run_self_test() -> int:
    failures: list[str] = []
    try:
        with tempfile.TemporaryDirectory(prefix="release-certification-") as raw:
            base = Path(raw)
            root = base / "repo"
            evidence = base / "evidence"
            root.mkdir()
            evidence.mkdir()
            fixture_git(root, "init", "-b", "main")
            (root / ".github").mkdir()
            (root / "VERSION").write_text("1.2.3\n", encoding="utf-8")
            (root / ".github/repository-profile.json").write_text(
                json.dumps({"repository": "owner/repo"}) + "\n", encoding="utf-8"
            )
            (root / "tools").mkdir()
            (root / "tools/release_certification.py").write_text("fixture\n", encoding="utf-8")
            fixture_git(root, "add", "--all")
            fixture_git(root, "commit", "-m", "fixture")
            candidate = fixture_git(root, "rev-parse", "HEAD")
            records = []
            for role in sorted(REQUIRED_ROLES):
                filename = role + ".json"
                (evidence / filename).write_text(json.dumps({"role": role}) + "\n", encoding="utf-8")
                records.append(
                    {
                        "id": role,
                        "role": role,
                        "visibility": "public-file",
                        "path": filename,
                        "public_name": filename,
                        "redaction_reviewed": True,
                    }
                )
            records.append(
                {
                    "id": "restricted-provider-record",
                    "role": "restricted-provider-record",
                    "visibility": "restricted-reference",
                    "sha256": "a" * 64,
                    "reference": "provider-record:123",
                    "reason": "provider access restriction",
                }
            )
            specification = base / "certification-input.json"
            specification.write_text(
                json.dumps(
                    {
                        "schema_version": 1,
                        "repository": "owner/repo",
                        "version": "1.2.3",
                        "tag": "v1.2.3",
                        "candidate_sha": candidate,
                        "records": records,
                    },
                    indent=2,
                    sort_keys=True,
                )
                + "\n",
                encoding="utf-8",
            )
            first = base / "out-1"
            second = base / "out-2"
            result_one = build_bundle(root, specification, evidence, first)
            result_two = build_bundle(root, specification, evidence, second)
            first_bundle = Path(result_one["bundle"])
            second_bundle = Path(result_two["bundle"])
            require(first_bundle.read_bytes() == second_bundle.read_bytes(), "bundle is not deterministic")
            verified = verify_bundle(
                first_bundle,
                Path(result_one["manifest"]),
                Path(result_one["digest"]),
            )
            require(verified["candidate_sha"] == candidate, "verification lost candidate binding")

            signing_key = base / "certification-signing-key"
            subprocess.run(
                ["ssh-keygen", "-q", "-t", "ed25519", "-N", "", "-f", str(signing_key)],
                check=True,
                capture_output=True,
                timeout=30,
            )
            subprocess.run(
                [
                    "ssh-keygen", "-Y", "sign",
                    "-f", str(signing_key),
                    "-n", CERTIFICATION_SIGNATURE_NAMESPACE,
                    str(first_bundle),
                ],
                check=True,
                capture_output=True,
                timeout=30,
            )
            signature_path = Path(str(first_bundle) + ".sig")
            public_parts = signing_key.with_suffix(".pub").read_text(
                encoding="utf-8"
            ).strip().split()
            fixture_public_key = " ".join(public_parts[:2])
            verify_ssh_signature(
                first_bundle.read_bytes(),
                signature_path,
                "fixture",
                [fixture_public_key],
            )
            try:
                verify_ssh_signature(
                    first_bundle.read_bytes() + b"x",
                    signature_path,
                    "fixture",
                    [fixture_public_key],
                )
                failures.append("certification signature accepted tampered bundle bytes")
            except CertificationError:
                pass

            exercise_manifest_verification_controls(
                first_bundle,
                Path(result_one["manifest"]),
                Path(result_one["digest"]),
                base,
                failures,
            )

            same_name_tamper = base / "tampered" / first_bundle.name
            same_name_tamper.parent.mkdir()
            same_name_tamper.write_bytes(first_bundle.read_bytes() + b"x")
            try:
                verify_bundle(
                    same_name_tamper,
                    Path(result_one["manifest"]),
                    Path(result_one["digest"]),
                )
                failures.append("same-filename tampered bundle was accepted")
            except CertificationError:
                pass

            release = {
                "tag_name": "v1.2.3",
                "assets": [
                    {
                        "id": 1,
                        "name": Path(result_one["bundle"]).name,
                        "url": "https://api.github.example/assets/1",
                        "browser_download_url": "https://github.example/download/1",
                        "size": Path(result_one["bundle"]).stat().st_size,
                    },
                    {
                        "id": 2,
                        "name": Path(result_one["manifest"]).name,
                        "url": "https://api.github.example/assets/2",
                        "browser_download_url": "https://github.example/download/2",
                        "size": Path(result_one["manifest"]).stat().st_size,
                    },
                    {
                        "id": 3,
                        "name": Path(result_one["digest"]).name,
                        "url": "https://api.github.example/assets/3",
                        "browser_download_url": "https://github.example/download/3",
                        "size": Path(result_one["digest"]).stat().st_size,
                    },
                    {
                        "id": 4,
                        "name": certification_asset_names("v1.2.3")["signature"],
                        "url": "https://api.github.example/assets/4",
                        "browser_download_url": "https://github.example/download/4",
                        "size": signature_path.stat().st_size,
                    },
                    {
                        "id": 5,
                        "name": "dist/example.xlsm",
                        "url": "https://api.github.example/assets/5",
                        "browser_download_url": "https://github.example/download/5",
                        "size": 123,
                    },
                ],
            }
            release_path = base / "release.json"
            release_path.write_text(json.dumps(release) + "\n", encoding="utf-8")
            plan = plan_release_assets(release_path, "v1.2.3")
            require(len(plan["certification_assets"]) == 4, "release plan lost certification assets")
            require(len(plan["product_assets"]) == 1, "release plan did not separate product assets")
            plan_path = base / "release-plan.json"
            plan_path.write_bytes(canonical_json(plan))
            downloaded = base / "downloaded"
            downloaded.mkdir()
            for key in ("bundle", "manifest", "digest"):
                source = Path(result_one[key])
                (downloaded / source.name).write_bytes(source.read_bytes())
            (downloaded / certification_asset_names("v1.2.3")["signature"]).write_bytes(
                signature_path.read_bytes()
            )
            retained = verify_release_plan(root, plan_path, downloaded, candidate)
            require(retained["bundle_sha256"] == result_one["bundle_sha256"], "retained verification lost bundle digest")

            (downloaded / first_bundle.name).write_bytes(first_bundle.read_bytes() + b"x")
            try:
                verify_release_plan(root, plan_path, downloaded, candidate)
                failures.append("tampered downloaded release asset was accepted")
            except CertificationError:
                pass
            (downloaded / first_bundle.name).write_bytes(first_bundle.read_bytes())

            missing_release = json.loads(release_path.read_text(encoding="utf-8"))
            missing_release["assets"] = missing_release["assets"][1:]
            missing_release_path = base / "release-missing.json"
            missing_release_path.write_text(json.dumps(missing_release) + "\n", encoding="utf-8")
            try:
                plan_release_assets(missing_release_path, "v1.2.3")
                failures.append("missing published certification asset was accepted")
            except CertificationError:
                pass

            missing_signature_release = json.loads(release_path.read_text(encoding="utf-8"))
            signature_name = certification_asset_names("v1.2.3")["signature"]
            missing_signature_release["assets"] = [
                asset for asset in missing_signature_release["assets"]
                if asset["name"] != signature_name
            ]
            missing_signature_path = base / "release-missing-signature.json"
            missing_signature_path.write_text(
                json.dumps(missing_signature_release) + "\n",
                encoding="utf-8",
            )
            try:
                plan_release_assets(missing_signature_path, "v1.2.3")
                failures.append("missing certification signature asset was accepted")
            except CertificationError:
                pass

            stale_release = json.loads(release_path.read_text(encoding="utf-8"))
            stale_release["assets"].append(
                {
                    "id": 5,
                    "name": "certification-v1.2.2.zip",
                    "url": "https://api.github.example/assets/5",
                    "browser_download_url": "https://github.example/download/5",
                    "size": 1,
                }
            )
            stale_release_path = base / "release-stale.json"
            stale_release_path.write_text(json.dumps(stale_release) + "\n", encoding="utf-8")
            try:
                plan_release_assets(stale_release_path, "v1.2.3")
                failures.append("unexpected certification namespace asset was accepted")
            except CertificationError:
                pass

            bad_spec = json.loads(specification.read_text(encoding="utf-8"))
            bad_spec["records"] = [row for row in bad_spec["records"] if row["role"] != "excel-evidence"]
            missing = base / "missing.json"
            missing.write_text(json.dumps(bad_spec), encoding="utf-8")
            try:
                build_bundle(root, missing, evidence, base / "out-missing")
                failures.append("missing required role was accepted")
            except CertificationError:
                pass

            secret_name = "release-evidence.json"
            (evidence / secret_name).write_text("-----BEGIN " + "PRIVATE KEY-----\n", encoding="utf-8")
            try:
                build_bundle(root, specification, evidence, base / "out-secret")
                failures.append("secret-like public evidence was accepted")
            except CertificationError:
                pass
    except (CertificationError, OSError, subprocess.SubprocessError, zipfile.BadZipFile) as error:
        failures.append(str(error))

    if failures:
        for failure in failures:
            print(f"[FAIL] {failure}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1
    print(
        "SELF-TEST PASS: deterministic build, exact-SHA binding, independent manifest schema/role "
        "verification, same-filename tamper, restricted-reference, secret-leakage, GitHub Release "
        "asset planning, product/evidence separation, retained-download verification and "
        "missing/stale asset controls passed."
    )
    return 0


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--build", type=Path, metavar="SPECIFICATION")
    parser.add_argument("--evidence-dir", type=Path)
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--verify-bundle", type=Path)
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--digest", type=Path)
    parser.add_argument("--verify-signature", type=Path, metavar="SIGNATURE")
    parser.add_argument("--bundle", type=Path)
    parser.add_argument("--plan-release-assets", type=Path, metavar="RELEASE_JSON")
    parser.add_argument("--release-tag")
    parser.add_argument("--verify-release-plan", type=Path, metavar="PLAN_JSON")
    parser.add_argument("--download-dir", type=Path)
    parser.add_argument("--candidate-sha")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        if options.self_test:
            return run_self_test()
        selected = sum(
            item is not None
            for item in (
                options.build,
                options.verify_bundle,
                options.verify_signature,
                options.plan_release_assets,
                options.verify_release_plan,
            )
        )
        require(selected == 1, "choose exactly one build or verification mode")
        if options.build is not None:
            require(options.evidence_dir is not None, "--build requires --evidence-dir")
            require(options.output_dir is not None, "--build requires --output-dir")
            result = build_bundle(options.root, options.build, options.evidence_dir, options.output_dir)
        elif options.verify_bundle is not None:
            require(options.manifest is not None and options.digest is not None, "--verify-bundle requires --manifest and --digest")
            require(options.evidence_dir is None and options.output_dir is None, "verify mode does not accept build directories")
            result = verify_bundle(options.verify_bundle, options.manifest, options.digest)
        elif options.verify_signature is not None:
            require(options.bundle is not None, "--verify-signature requires --bundle")
            require(options.candidate_sha is not None, "--verify-signature requires --candidate-sha")
            result = verify_certification_signature(
                options.root,
                options.candidate_sha,
                options.bundle,
                options.verify_signature,
            )
        elif options.plan_release_assets is not None:
            require(options.release_tag is not None, "--plan-release-assets requires --release-tag")
            result = plan_release_assets(options.plan_release_assets, options.release_tag)
        else:
            require(options.verify_release_plan is not None, "release plan path is unavailable")
            require(options.download_dir is not None, "--verify-release-plan requires --download-dir")
            require(options.candidate_sha is not None, "--verify-release-plan requires --candidate-sha")
            result = verify_release_plan(
                options.root,
                options.verify_release_plan,
                options.download_dir,
                options.candidate_sha,
            )
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except (CertificationError, OSError, UnicodeError, zipfile.BadZipFile) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
