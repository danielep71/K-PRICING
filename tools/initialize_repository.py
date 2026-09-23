#!/usr/bin/env python3
"""Deterministically initialize a repository created from this template.

Dry-run is the default. Pass --apply only after reviewing the complete plan.
The implementation validates every input and renders every affected file in
memory before changing the working tree. Replacements are atomic per file;
filesystem failures trigger a rollback attempt, not a repository-wide transaction.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import tempfile
from pathlib import Path, PurePosixPath
from typing import Any, Iterable

CONFIG_PATH = ".github/repository-profile.json"
RECORD_PATH = ".github/initialization.json"
CANONICAL_SOCIAL_PREVIEW_PATH = "assets/social-preview.png"
SUPPORTED_PROFILES = ("application", "library", "ui-component")
TOKEN_NAME_PATTERN = re.compile(r"[A-Z][A-Z0-9_]*")
REPOSITORY_PATTERN = re.compile(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+")
MARKER_PATTERN = re.compile(
    r"<!-- template:(remove|profile:(?:application|library|ui-component)|"
    r"optional:[A-Z][A-Z0-9_]*|repeatable:[A-Z][A-Z0-9_]*):(start|end) -->"
)
VBA_SUFFIXES = {".bas", ".cls", ".frm"}
EXECUTABLE_SUFFIXES = {
    ".bat",
    ".cjs",
    ".cmd",
    ".js",
    ".json",
    ".mjs",
    ".ps1",
    ".py",
    ".sh",
    ".vbs",
    ".yaml",
    ".yml",
} | VBA_SUFFIXES
TEXT_SUFFIXES = {
    ".cfg",
    ".cff",
    ".csv",
    ".ini",
    ".jsonc",
    ".md",
    ".reg",
    ".svg",
    ".toml",
    ".tsv",
    ".txt",
    ".xml",
} | EXECUTABLE_SUFFIXES
TEXT_NAMES = {
    ".editorconfig",
    ".gitattributes",
    ".gitignore",
    ".gitkeep",
    "Dockerfile",
    "LICENSE",
    "Makefile",
    "VERSION",
}


class InitializationError(Exception):
    """A deterministic validation or initialization failure."""


def _git(root: Path, *arguments: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["git", "-C", str(root), *arguments],
        check=check,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )


def _tracked_files(root: Path) -> tuple[str, ...]:
    completed = subprocess.run(
        ["git", "-C", str(root), "ls-files", "-z"],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    return tuple(
        sorted(
            item.decode("utf-8", errors="surrogateescape")
            for item in completed.stdout.split(b"\0")
            if item
        )
    )


def _is_text(path: str) -> bool:
    item = PurePosixPath(path)
    return item.name in TEXT_NAMES or item.suffix.casefold() in TEXT_SUFFIXES


def _decode(path: str, data: bytes) -> str:
    encoding = "cp1252" if PurePosixPath(path).suffix.casefold() in VBA_SUFFIXES else "utf-8"
    return data.decode(encoding)


def _encode(path: str, text: str) -> bytes:
    encoding = "cp1252" if PurePosixPath(path).suffix.casefold() in VBA_SUFFIXES else "utf-8"
    return text.encode(encoding)


def _load_config(root: Path) -> dict[str, Any]:
    try:
        document = json.loads((root / CONFIG_PATH).read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise InitializationError(f"Cannot read {CONFIG_PATH}: {error}") from error
    if not isinstance(document, dict):
        raise InitializationError(f"{CONFIG_PATH} must contain a JSON object.")
    placeholders = document.get("placeholders")
    if not isinstance(placeholders, dict):
        raise InitializationError(f"{CONFIG_PATH} has no placeholder schema.")
    try:
        token_pattern = re.compile(str(placeholders["pattern"]))
        catalogue = placeholders["catalogue"]
        markers = placeholders["block_markers"]
        template_only_paths = placeholders["template_only_paths"]
        excluded_paths = placeholders["exclude_paths"]
    except (KeyError, TypeError, re.error) as error:
        raise InitializationError(f"{CONFIG_PATH} has an invalid placeholder schema.") from error
    if not isinstance(catalogue, dict) or not isinstance(markers, dict):
        raise InitializationError(f"{CONFIG_PATH} has an invalid placeholder catalogue.")
    if token_pattern.groups != 1:
        raise InitializationError(f"{CONFIG_PATH} placeholder pattern must capture one token name.")
    if not isinstance(template_only_paths, list) or not isinstance(excluded_paths, list):
        raise InitializationError(f"{CONFIG_PATH} has invalid placeholder path lists.")
    return document


def _parse_assignments(entries: Iterable[str], option: str) -> dict[str, list[str]]:
    parsed: dict[str, list[str]] = {}
    for entry in entries:
        if "=" not in entry:
            raise InitializationError(f"{option} values must use NAME=value: {entry!r}")
        name, value = entry.split("=", 1)
        name = name.strip()
        value = value.strip()
        if not TOKEN_NAME_PATTERN.fullmatch(name):
            raise InitializationError(f"Invalid placeholder name for {option}: {name!r}")
        if not value:
            raise InitializationError(f"Empty value for {name} is not allowed.")
        if "\0" in value or "\n" in value or "\r" in value:
            raise InitializationError(f"{name} must be a single line.")
        if "{{" in value or "}}" in value or "<!-- template:" in value:
            raise InitializationError(f"{name} contains reserved template syntax.")
        parsed.setdefault(name, []).append(value)
    return parsed


def _validate_values(
    root: Path,
    tracked: set[str],
    catalogue: dict[str, dict[str, Any]],
    scalar_entries: Iterable[str],
    repeatable_entries: Iterable[str],
    *,
    require_preview_file: bool = True,
) -> tuple[dict[str, str], dict[str, list[str]]]:
    scalar_lists = _parse_assignments(scalar_entries, "--set")
    repeatable = _parse_assignments(repeatable_entries, "--add")
    known = set(catalogue)
    unknown = (set(scalar_lists) | set(repeatable)) - known
    if unknown:
        raise InitializationError("Unknown substitutions: " + ", ".join(sorted(unknown)))

    scalars: dict[str, str] = {}
    for name, values in scalar_lists.items():
        if len(values) != 1:
            raise InitializationError(f"{name} was supplied more than once with --set.")
        category = catalogue[name].get("category")
        if category not in {"required", "optional"}:
            raise InitializationError(f"{name} cannot be supplied with --set; category is {category}.")
        scalars[name] = values[0]

    for name in repeatable:
        category = catalogue[name].get("category")
        if category != "repeatable":
            raise InitializationError(f"{name} cannot be supplied with --add; category is {category}.")

    required = {
        name for name, specification in catalogue.items()
        if specification.get("category") == "required"
    }
    missing = required - set(scalars)
    if missing:
        raise InitializationError("Missing required substitutions: " + ", ".join(sorted(missing)))

    if not REPOSITORY_PATTERN.fullmatch(scalars.get("REPOSITORY_PATH", "")):
        raise InitializationError("REPOSITORY_PATH must use GitHub owner/name form.")
    contact = scalars.get("SUPPORT_CONTACT", "")
    if not (re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", contact) or contact.startswith("https://")):
        raise InitializationError("SUPPORT_CONTACT must be an email address or HTTPS URL.")
    length_limits = {
        "MAINTAINER_NAME": 100,
        "PROJECT_NAME": 100,
        "PROJECT_TAGLINE": 120,
        "PROJECT_DESCRIPTION": 280,
    }
    for name, limit in length_limits.items():
        if len(scalars[name]) > limit:
            raise InitializationError(f"{name} exceeds its {limit}-character limit.")
    if not re.fullmatch(r"20[0-9]{2}", scalars.get("COPYRIGHT_YEAR", "")):
        raise InitializationError("COPYRIGHT_YEAR must be a four-digit year from 2000 through 2099.")

    preview = scalars.get("SOCIAL_PREVIEW_PATH")
    if preview:
        item = PurePosixPath(preview)
        if item.is_absolute() or ".." in item.parts:
            raise InitializationError("SOCIAL_PREVIEW_PATH must name a repository-relative file.")
        if require_preview_file:
            if preview not in tracked:
                raise InitializationError(
                    "SOCIAL_PREVIEW_PATH must name a tracked repository-relative file."
                )
            if not (root / item).is_file():
                raise InitializationError("SOCIAL_PREVIEW_PATH does not exist in the working tree.")
    return scalars, repeatable


def _render_blocks(
    path: str,
    text: str,
    profile: str,
    scalars: dict[str, str],
    repeatable: dict[str, list[str]],
    catalogue: dict[str, dict[str, Any]],
) -> str:
    output: list[str] = []
    active: tuple[str, bool] | None = None
    for line_number, line in enumerate(text.splitlines(keepends=True), 1):
        match = MARKER_PATTERN.fullmatch(line.strip())
        if "<!-- template:" in line and match is None:
            raise InitializationError(f"{path}:{line_number}: invalid template block marker.")
        if match:
            marker, boundary = match.groups()
            if boundary == "start":
                if active is not None:
                    raise InitializationError(f"{path}:{line_number}: template blocks may not nest.")
                if marker == "remove":
                    keep = False
                elif marker.startswith("profile:"):
                    keep = marker.split(":", 1)[1] == profile
                elif marker.startswith("optional:"):
                    name = marker.split(":", 1)[1]
                    if catalogue.get(name, {}).get("category") != "optional":
                        raise InitializationError(f"{path}:{line_number}: {name} is not optional.")
                    keep = name in scalars
                else:
                    name = marker.split(":", 1)[1]
                    if catalogue.get(name, {}).get("category") != "repeatable":
                        raise InitializationError(f"{path}:{line_number}: {name} is not repeatable.")
                    keep = bool(repeatable.get(name))
                active = (marker, keep)
            else:
                if active is None or active[0] != marker:
                    raise InitializationError(f"{path}:{line_number}: unmatched template block end.")
                active = None
            continue
        if active is None or active[1]:
            output.append(line)
    if active is not None:
        raise InitializationError(f"{path}: unclosed template block {active[0]}.")
    return "".join(output)


def _replacement_values(
    profile: str,
    scalars: dict[str, str],
    repeatable: dict[str, list[str]],
    catalogue: dict[str, dict[str, Any]],
) -> dict[str, str]:
    values = dict(scalars)
    for name, specification in catalogue.items():
        category = specification.get("category")
        if category == "profile-specific":
            values[name] = str(specification["values"][profile])
        elif category == "repeatable":
            item_format = str(specification["item_format"])
            values[name] = "\n".join(
                item_format.replace("{value}", item) for item in repeatable.get(name, [])
            )
    return values


def _render_readme_badges(
    path: str, text: str, template_repository: str, repository: str
) -> str:
    """Retarget the template's live README presentation URLs for an adopter."""
    if path != "README.md":
        return text
    for prefix in (
        "https://github.com/",
        "https://img.shields.io/github/v/release/",
        "https://img.shields.io/github/issues/",
        "https://api.scorecard.dev/projects/github.com/",
        "https://scorecard.dev/viewer/?uri=github.com/",
    ):
        # Match a repository boundary, so similarly named repositories stay untouched.
        text = re.sub(
            re.escape(prefix + template_repository) + r"(?=[/?#)\s]|$)",
            lambda match: prefix + repository,
            text,
        )
    preview = re.search(
        r"(?m)^[ \t]*<!-- generated-social-preview: ([^<>\r\n]+) -->[ \t]*\r?\n?",
        text,
    )
    if preview is not None:
        selected_preview = preview.group(1).strip()
        text = text[: preview.start()] + text[preview.end() :]
        text = text.replace(
            f'src="{CANONICAL_SOCIAL_PREVIEW_PATH}"',
            f'src="{selected_preview}"',
            1,
        )
    return text


def _reset_changelog(text: str) -> str:
    start = text.find("## [Unreleased]")
    if start < 0:
        raise InitializationError("CHANGELOG.md has no Unreleased section to reset.")
    end = text.find("\n---", start)
    if end < 0:
        raise InitializationError("CHANGELOG.md has no boundary after Unreleased.")
    replacement = (
        "## [Unreleased]\n\n"
        "<!-- Add only user-visible changes made in this generated project. -->\n\n"
        "No unreleased changes recorded.\n"
    )
    return text[:start] + replacement + text[end:]


def _directory_readme(project_name: str, profile: str, directory: str) -> bytes:
    label = PurePosixPath(directory).name.replace("-", " ").title()
    content = (
        f"# 📂 {label}\n\n"
        f"This directory is reserved for {project_name}'s {profile} profile. "
        "Add authoritative exported source here when the project needs this "
        "component role. The required facade, core and regression starter "
        "already exist in their governed locations.\n"
    )
    return content.encode("utf-8")


def _record(
    profile: str,
    scalars: dict[str, str],
    repeatable: dict[str, list[str]],
    contract: dict[str, Any],
) -> bytes:
    values: dict[str, Any] = dict(sorted(scalars.items()))
    values.update({name: items for name, items in sorted(repeatable.items())})
    document = {
        "schema_version": 1,
        "profile": profile,
        "template_contract": contract,
        "values": values,
    }
    return (json.dumps(document, indent=2, ensure_ascii=False) + "\n").encode("utf-8")


def _already_initialized(
    root: Path,
    config: dict[str, Any],
    profile: str,
    scalars: dict[str, str],
    repeatable: dict[str, list[str]],
) -> bool:
    if config.get("mode") != "generated":
        return False
    try:
        existing = (root / RECORD_PATH).read_bytes()
    except OSError as error:
        raise InitializationError(f"Generated repository is missing {RECORD_PATH}.") from error
    if config.get("profile") != profile or config.get("repository") != scalars["REPOSITORY_PATH"]:
        raise InitializationError("Repository is already initialized with different profile or repository inputs.")
    if existing != _record(profile, scalars, repeatable, config.get("template_contract", {})):
        raise InitializationError("Repository is already initialized with different substitution inputs.")
    return True


def _reject_executable_placeholders(path: str, matches: list[Any]) -> None:
    if matches and PurePosixPath(path).suffix.casefold() in EXECUTABLE_SUFFIXES:
        raise InitializationError(
            f"Placeholders are prohibited in executable or VBA file {path}."
        )


def _strip_template_maintenance_workflow(path: str, text: str) -> str:
    if path != ".github/workflows/static-checks.yml":
        return text
    lines = text.splitlines(keepends=True)
    output: list[str] = []
    skipping = False
    for line in lines:
        if line.startswith("      - name: Exercise policy-branch coverage determinism"):
            skipping = True
            continue
        if skipping and line.startswith("      - name: Exercise positive and degraded checker paths"):
            skipping = False
        if skipping:
            continue
        if "test-results/policy-coverage." in line:
            continue
        if "POLICY_COVERAGE_" in line:
            continue
        if '"Policy-coverage self-test:' in line or '"Policy coverage:' in line:
            continue
        output.append(line)
    return "".join(output)


def _build_changes(
    root: Path,
    profile: str,
    scalar_entries: Iterable[str],
    repeatable_entries: Iterable[str],
) -> tuple[dict[str, bytes | None], dict[str, str]]:
    config = _load_config(root)
    tracked_files = _tracked_files(root)
    tracked = set(tracked_files)
    placeholders = config["placeholders"]
    catalogue = placeholders["catalogue"]
    scalars, repeatable = _validate_values(
        root,
        tracked,
        catalogue,
        scalar_entries,
        repeatable_entries,
        require_preview_file=config.get("mode") != "generated",
    )
    values = _replacement_values(profile, scalars, repeatable, catalogue)
    if _already_initialized(root, config, profile, scalars, repeatable):
        return {}, values
    if config.get("mode") != "template":
        raise InitializationError("Repository mode must be template or a matching prior initialization.")

    excluded = set(placeholders["exclude_paths"])
    configured_template_only = set(placeholders["template_only_paths"])
    missing_template_paths = configured_template_only - tracked
    if missing_template_paths:
        raise InitializationError(
            "Configured template-only paths are not tracked: "
            + ", ".join(sorted(missing_template_paths))
        )
    selected_preview = scalars.get("SOCIAL_PREVIEW_PATH", "")
    if (
        selected_preview in configured_template_only
        and selected_preview != CANONICAL_SOCIAL_PREVIEW_PATH
    ):
        raise InitializationError(
            "SOCIAL_PREVIEW_PATH may retain only "
            f"{CANONICAL_SOCIAL_PREVIEW_PATH} from template_only_paths."
        )
    retained_template_only = (
        {CANONICAL_SOCIAL_PREVIEW_PATH}
        if selected_preview == CANONICAL_SOCIAL_PREVIEW_PATH
        else set()
    )
    template_only = configured_template_only - retained_template_only
    token_pattern = re.compile(placeholders["pattern"])
    changes: dict[str, bytes | None] = {path: None for path in sorted(template_only)}
    seen: set[str] = set()

    for path in tracked_files:
        if path in template_only or path in excluded or not _is_text(path):
            continue
        source = (root / path).read_bytes()
        text = _decode(path, source)
        matches = list(token_pattern.finditer(text))
        _reject_executable_placeholders(path, matches)
        rendered = _render_blocks(path, text, profile, scalars, repeatable, catalogue)
        for match in token_pattern.finditer(rendered):
            name = match.group(1)
            if name not in catalogue:
                raise InitializationError(f"{path}: unknown placeholder {name}.")
            seen.add(name)
        for name, value in values.items():
            rendered = rendered.replace("{{" + name + "}}", value)
        rendered = _strip_template_maintenance_workflow(path, rendered)
        rendered = _render_readme_badges(
            path, rendered, config["repository"], scalars["REPOSITORY_PATH"]
        )
        if path == ".github/ISSUE_TEMPLATE/config.yml":
            template_security_url = (
                f"https://github.com/{config['repository']}/security/policy"
            )
            generated_security_url = (
                f"https://github.com/{scalars['REPOSITORY_PATH']}/security/policy"
            )
            if template_security_url not in rendered:
                raise InitializationError(
                    f"{path}: canonical template security URL is missing."
                )
            rendered = rendered.replace(template_security_url, generated_security_url)
        unresolved = sorted({match.group(1) for match in token_pattern.finditer(rendered)})
        if unresolved:
            raise InitializationError(f"{path}: unresolved placeholders: {', '.join(unresolved)}")
        if path == "CHANGELOG.md":
            rendered = _reset_changelog(rendered)
        elif path == "VERSION":
            rendered = "0.0.0\n"
        output = _encode(path, rendered)
        if output != source:
            changes[path] = output

    supplied = set(scalars) | set(repeatable)
    unused = supplied - seen
    if unused:
        raise InitializationError("Unused substitutions: " + ", ".join(sorted(unused)))
    expected_used = {
        name for name, specification in catalogue.items()
        if specification.get("category") in {"required", "profile-specific"}
    }
    missing_uses = expected_used - seen
    if missing_uses:
        raise InitializationError("Registered required/profile placeholders are unused: " + ", ".join(sorted(missing_uses)))

    generated_config = json.loads(json.dumps(config))
    generated_config["mode"] = "generated"
    generated_config["profile"] = profile
    generated_config["repository"] = scalars["REPOSITORY_PATH"]
    changes[CONFIG_PATH] = (
        json.dumps(generated_config, indent=2, ensure_ascii=False) + "\n"
    ).encode("utf-8")
    changes[RECORD_PATH] = _record(
        profile, scalars, repeatable, generated_config.get("template_contract", {})
    )

    profile_settings = generated_config["profiles"][profile]
    for directory in profile_settings["required_directories"]:
        prefix = directory.rstrip("/") + "/"
        if not any(
            path.startswith(prefix) and changes.get(path, b"present") is not None
            for path in tracked_files
        ):
            readme = prefix + "README.md"
            changes[readme] = _directory_readme(scalars["PROJECT_NAME"], profile, directory)
    return dict(sorted(changes.items())), values


def _sha256(data: bytes | None) -> str | None:
    return hashlib.sha256(data).hexdigest() if data is not None else None


def _plan(root: Path, profile: str, changes: dict[str, bytes | None]) -> dict[str, Any]:
    entries = []
    for path, after in changes.items():
        target = root / path
        before = target.read_bytes() if target.is_file() else None
        action = "delete" if after is None else ("create" if before is None else "update")
        entries.append(
            {
                "action": action,
                "path": path,
                "before_sha256": _sha256(before),
                "after_sha256": _sha256(after),
            }
        )
    return {
        "schema_version": 1,
        "status": "ready" if entries else "no-op",
        "mode": "dry-run",
        "profile": profile,
        "changes": entries,
    }


def _apply_changes(root: Path, changes: dict[str, bytes | None]) -> None:
    originals: dict[str, tuple[bytes, int] | None] = {}
    staged: dict[str, Path] = {}
    try:
        for path, after in changes.items():
            target = root / path
            originals[path] = (
                (target.read_bytes(), target.stat().st_mode) if target.is_file() else None
            )
            if after is None:
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            descriptor, temporary_name = tempfile.mkstemp(
                prefix=f".{target.name}.", suffix=".initialize", dir=target.parent
            )
            temporary = Path(temporary_name)
            with os.fdopen(descriptor, "wb") as stream:
                stream.write(after)
                stream.flush()
                os.fsync(stream.fileno())
            original = originals[path]
            mode = stat.S_IMODE(original[1]) if original is not None else 0o644
            os.chmod(temporary, mode)
            staged[path] = temporary

        for path, after in changes.items():
            target = root / path
            if after is None:
                if target.exists():
                    target.unlink()
            else:
                os.replace(staged[path], target)
    except Exception as error:
        for temporary in staged.values():
            temporary.unlink(missing_ok=True)
        for path, original in originals.items():
            target = root / path
            if original is None:
                target.unlink(missing_ok=True)
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(original[0])
                os.chmod(target, original[1])
        raise InitializationError(f"Apply failed; original files were restored: {error}") from error


def _tree_digest(root: Path) -> str:
    digest = hashlib.sha256()
    for path in sorted(item for item in root.rglob("*") if item.is_file() and ".git" not in item.parts):
        digest.update(path.relative_to(root).as_posix().encode("utf-8"))
        digest.update(b"\0")
        digest.update(path.read_bytes())
        digest.update(b"\0")
    return digest.hexdigest()


def _fixture_arguments(profile: str) -> tuple[list[str], list[str]]:
    label = profile.replace("-", " ").title()
    scalars = [
        "COPYRIGHT_YEAR=2026",
        "MAINTAINER_NAME=Example Maintainer",
        f"PROJECT_NAME=Fixture {label}",
        f"PROJECT_TAGLINE=Deterministic {label} fixture",
        f"PROJECT_DESCRIPTION=Generated fixture for the {profile} repository profile.",
        f"REPOSITORY_PATH=example/fixture-{profile}",
        "SUPPORT_CONTACT=security@example.invalid",
    ]
    repeatable = [
        "ADDITIONAL_TEST_COMMAND=python3 tools/check_repo.py --root .",
        f"KNOWN_LIMITATION=Excel execution is not available in the {profile} initializer fixture.",
    ]
    return scalars, repeatable


def _copy_fixture(source: Path, destination: Path) -> None:
    shutil.copytree(
        source,
        destination,
        ignore=shutil.ignore_patterns(".git", "__pycache__", "test-results"),
    )
    _git(destination, "init", "-b", "main")
    _git(destination, "config", "user.name", "Initializer Self-Test")
    _git(destination, "config", "user.email", "initializer@example.invalid")
    _git(destination, "add", "--all")
    _git(destination, "commit", "-m", "Create template fixture")


def _quality_report(
    root: Path,
) -> tuple[subprocess.CompletedProcess[str], dict[str, Any]]:
    report_path = root.parent / f"{root.name}-quality.json"
    try:
        completed = subprocess.run(
            [
                sys.executable,
                str(root / "tools/check_repo.py"),
                "--root",
                str(root),
                "--output",
                str(report_path),
            ],
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        report = json.loads(report_path.read_text(encoding="utf-8"))
        return completed, report
    finally:
        report_path.unlink(missing_ok=True)


def _make_component_variant(
    source: Path,
    destination: Path,
    remove_paths: tuple[str, ...],
) -> None:
    _copy_fixture(source, destination)
    config = json.loads((destination / CONFIG_PATH).read_text(encoding="utf-8"))
    components = config["vba"]["components"]
    removed_roles: set[str] = set()
    readme_path = destination / "README.md"
    readme = readme_path.read_text(encoding="utf-8")
    for relative in remove_paths:
        role = components.pop(relative)
        removed_roles.add(role)
        path = destination / relative
        path.unlink()
        placeholder = path.parent / "README.md"
        if not any(item.is_file() for item in path.parent.iterdir()):
            placeholder.write_text(
                "# Fixture placeholder\n\n"
                f"Instruction-only placeholder after removing the {role} component.\n",
                encoding="utf-8",
                newline="\n",
            )
        readme = readme.replace(
            f"]({relative})",
            f"]({path.parent.relative_to(destination).as_posix()}/README.md)",
        )
    readme_path.write_text(readme, encoding="utf-8", newline="\n")
    if "public" in removed_roles:
        config["vba"]["public_api_manifest"] = None
    (destination / CONFIG_PATH).write_text(
        json.dumps(config, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    _git(destination, "add", "--all")
    _git(destination, "commit", "-m", "Create component-removal fixture")


def _assert_contract_failure(root: Path, profile: str, roles: tuple[str, ...]) -> None:
    completed, report = _quality_report(root)
    failed = {
        result["id"]
        for result in report["rules"]
        if result["status"] == "fail"
    }
    if completed.returncode != 1 or failed != {"generated-vba-contract"}:
        raise AssertionError(
            f"{profile} contract fixture failed unexpected rules {sorted(failed)}:\n"
            f"{completed.stdout}{completed.stderr}"
        )
    contract = next(
        result for result in report["rules"]
        if result["id"] == "generated-vba-contract"
    )
    messages = "\n".join(item["message"] for item in contract["findings"])
    if f"profile '{profile}'" not in messages.casefold():
        raise AssertionError(
            f"{profile} contract failure did not identify the selected profile."
        )
    for role in roles:
        if f"'{role}'" not in messages:
            raise AssertionError(f"{profile} contract failure did not identify role {role!r}.")


def _assert_failure_without_change(
    root: Path,
    profile: str,
    scalars: list[str],
    repeatable: list[str],
    expected: str,
) -> None:
    before = _tree_digest(root)
    try:
        _build_changes(root, profile, scalars, repeatable)
    except InitializationError as error:
        if expected not in str(error):
            raise AssertionError(f"Expected {expected!r}, observed {error!r}") from error
    else:
        raise AssertionError(f"Expected initialization failure containing {expected!r}.")
    if _tree_digest(root) != before:
        raise AssertionError("Failed validation changed the fixture tree.")


def _make_unused_fixture(source: Path, destination: Path) -> None:
    _copy_fixture(source, destination)
    readme = destination / "README.md"
    text = readme.read_text(encoding="utf-8")
    start = "<!-- template:optional:SOCIAL_PREVIEW_PATH:start -->"
    end = "<!-- template:optional:SOCIAL_PREVIEW_PATH:end -->"
    first = text.find(start)
    last = text.find(end)
    if first < 0 or last < first:
        raise AssertionError("Optional social-preview fixture block is unavailable.")
    readme.write_text(text[:first] + text[last + len(end):], encoding="utf-8", newline="\n")
    preview = destination / "assets/social-preview.png"
    preview.write_bytes(b"fixture-preview\n")
    _git(destination, "add", "--all")
    _git(destination, "commit", "-m", "Create unused-substitution fixture")


def _assert_generated_cleanup(
    root: Path,
    profile: str,
    repository: str | None = None,
) -> None:
    for path in _tracked_files(root):
        if PurePosixPath(path).suffix.casefold() != ".md":
            continue
        text = (root / path).read_text(encoding="utf-8")
        if "<!-- template:" in text:
            raise AssertionError(f"{profile} retained a template marker in {path}.")
    issue_config = (root / ".github/ISSUE_TEMPLATE/config.yml").read_text(
        encoding="utf-8"
    )
    repository = repository or f"example/fixture-{profile}"
    expected_security_url = f"https://github.com/{repository}/security/policy"
    if expected_security_url not in issue_config:
        raise AssertionError(f"{profile} did not render its private-security URL.")


def _assert_fresh_generated_content(root: Path, profile: str) -> None:
    # Badge defaults are an initialization contract, not a maintained README policy.
    repository = f"example/fixture-{profile}"
    readme = (root / "README.md").read_text(encoding="utf-8")
    expected_badge_urls = (
        f"https://github.com/{repository}/actions/workflows/static-checks.yml/badge.svg",
        f"https://github.com/{repository}/actions/workflows/static-checks.yml)",
        f"https://img.shields.io/github/v/release/{repository}?",
        f"https://github.com/{repository}/releases)",
        f"https://img.shields.io/github/issues/{repository}/P1?",
        f"https://img.shields.io/github/issues/{repository}/P2?",
        f"https://img.shields.io/github/issues/{repository}/P3?",
        f"https://github.com/{repository}/issues?q=is%3Aissue%20is%3Aopen%20label%3AP1)",
        f"https://github.com/{repository}/issues?q=is%3Aissue%20is%3Aopen%20label%3AP2)",
        f"https://github.com/{repository}/issues?q=is%3Aissue%20is%3Aopen%20label%3AP3)",
    )
    if any(url not in readme for url in expected_badge_urls):
        raise AssertionError(f"{profile} did not retarget all README badges and links.")
    headings = {
        "application": "### Application commitments",
        "library": "### Library commitments",
        "ui-component": "### UI-component commitments",
    }
    for candidate, heading in headings.items():
        if (heading in readme) != (candidate == profile):
            raise AssertionError(f"{profile} retained an incorrect profile block: {candidate}.")
    changelog = (root / "CHANGELOG.md").read_text(encoding="utf-8")
    if "No unreleased changes recorded." not in changelog:
        raise AssertionError(f"{profile} did not reset generated changelog history.")
    if (root / "VERSION").read_text(encoding="utf-8") != "0.0.0\n":
        raise AssertionError(f"{profile} did not reset the generated version sentinel.")


def _make_evolved_generated_fixture(source: Path, destination: Path) -> None:
    _copy_fixture(source, destination)

    changelog = destination / "CHANGELOG.md"
    changelog_text = changelog.read_text(encoding="utf-8")
    sentinel = "No unreleased changes recorded."
    if sentinel not in changelog_text:
        raise AssertionError("Generated evolution fixture has no changelog sentinel.")
    changelog.write_text(
        changelog_text.replace(
            sentinel, "- Added a user-visible generated-project change.", 1
        ),
        encoding="utf-8",
        newline="\n",
    )

    preview = destination / "assets/social-preview.png"
    if not preview.is_file():
        raise AssertionError("Generated evolution fixture has no social-preview asset.")
    preview.unlink()

    readme = destination / "README.md"
    readme_text = readme.read_text(encoding="utf-8")
    profile = json.loads(
        (destination / RECORD_PATH).read_text(encoding="utf-8")
    )["profile"]
    profile_heading = {
        "application": "### Application commitments",
        "library": "### Library commitments",
        "ui-component": "### UI-component commitments",
    }[profile]
    if profile_heading not in readme_text:
        raise AssertionError("Generated evolution fixture has no profile heading.")
    readme_text = readme_text.replace(
        profile_heading, "### Maintained project commitments", 1
    )
    preview_pattern = re.compile(
        r'<p align="center">\n\s*<img src="assets/social-preview\.png".*?</p>\n\n---\n',
        re.DOTALL,
    )
    readme_text, replacements = preview_pattern.subn("", readme_text, count=1)
    if replacements != 1:
        raise AssertionError(
            "Generated evolution fixture could not remove its preview block."
        )
    # A maintained project owns its presentation: remove, replace and retarget badges.
    readme_text, removed = re.subn(
        r"^\[!\[Release\].*\n", "", readme_text, flags=re.MULTILINE
    )
    readme_text, replaced = re.subn(
        r"^\[!\[P1 issues\].*\n",
        "[Project issues](https://example.org/project/issues)\n",
        readme_text,
        count=1,
        flags=re.MULTILINE,
    )
    readme_text, removed_priorities = re.subn(
        r"^\[!\[P[23] issues\].*\n", "", readme_text, flags=re.MULTILINE
    )
    readme_text, retargeted = re.subn(
        r"^\[!\[Static checks\].*\n",
        "[CI status](https://example.org/project/ci)\n", readme_text,
        flags=re.MULTILINE,
    )
    if (removed, replaced, removed_priorities, retargeted) != (1, 1, 2, 1):
        raise AssertionError(
            "Generated evolution fixture did not customize release, priority, and CI badges."
        )
    readme.write_text(
        readme_text
        + "\n## Generated-project evolution fixture\n\n"
        + "This maintained section proves that initializer self-tests tolerate "
        + "normal README evolution.\n",
        encoding="utf-8",
        newline="\n",
    )
    _git(destination, "add", "--all")
    _git(destination, "commit", "-m", "Evolve generated project content")


def _record_arguments(root: Path) -> tuple[str, list[str], list[str]]:
    try:
        record = json.loads((root / RECORD_PATH).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise AssertionError(f"Generated repository has an invalid {RECORD_PATH}.") from error
    if (
        record.get("schema_version") != 1
        or record.get("profile") not in SUPPORTED_PROFILES
        or not isinstance(record.get("values"), dict)
        or not isinstance(record.get("template_contract"), dict)
    ):
        raise AssertionError(f"Generated repository has an invalid {RECORD_PATH} contract.")

    scalars: list[str] = []
    repeatable: list[str] = []
    for name, value in record["values"].items():
        if isinstance(value, str):
            scalars.append(f"{name}={value}")
        elif isinstance(value, list) and all(isinstance(item, str) for item in value):
            repeatable.extend(f"{name}={item}" for item in value)
        else:
            raise AssertionError(f"{RECORD_PATH} contains an invalid value for {name}.")
    return record["profile"], scalars, repeatable


def _assert_adopted_contract(source: Path, config: dict[str, Any]) -> None:
    """The adopted template contract survives initialization unchanged.

    ``repository`` becomes the generated project, but ``template_contract`` names
    the baseline it adopted and the template that published it. Substituting
    either would make the recorded baseline unusable for conformance.
    """
    contract = config.get("template_contract")
    if not isinstance(contract, dict) or set(contract) != {"version", "source"}:
        raise AssertionError(
            "Generated repository policy does not record a well-formed template_contract."
        )
    record = json.loads((source / RECORD_PATH).read_text(encoding="utf-8"))
    if record.get("template_contract") != contract:
        raise AssertionError(
            "Generated policy and initialization record disagree on the adopted contract."
        )
    if contract["source"] == config.get("repository"):
        raise AssertionError(
            "Generated repository claims to be its own template contract source."
        )
    completed = subprocess.run(
        [sys.executable, str(source / "tools" / "check_template_contract.py"), "--root", str(source)],
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode != 0:
        raise AssertionError(
            f"Generated repository failed the template-contract gate:\n"
            f"{completed.stdout}{completed.stderr}"
        )

def _generated_self_test(source: Path) -> None:
    config = _load_config(source)
    profile, scalars, repeatable = _record_arguments(source)
    repository = config.get("repository")
    if not isinstance(repository, str) or not repository:
        raise AssertionError("Generated repository policy has no repository identity.")
    tracked = set(_tracked_files(source))
    placeholders = config["placeholders"]
    scalar_values, repeatable_values = _validate_values(
        source,
        tracked,
        placeholders["catalogue"],
        scalars,
        repeatable,
        require_preview_file=False,
    )
    if not _already_initialized(
        source, config, profile, scalar_values, repeatable_values
    ):
        raise AssertionError(
            "Generated repository initialization record is not authoritative."
        )
    if config.get("profile") != profile or config.get("repository") != repository:
        raise AssertionError("Generated repository policy and initialization record disagree.")
    _assert_adopted_contract(source, config)
    _assert_generated_cleanup(source, profile, repository)

    documented = subprocess.run(
        [sys.executable, str(source / "tools" / "check_documentation.py"), "--root", str(source)],
        capture_output=True, text=True, check=False,
    )
    if documented.returncode != 0:
        raise AssertionError(
            f"Generated repository failed documentation contracts:\n{documented.stdout}{documented.stderr}"
        )

    completed, report = _quality_report(source)
    if completed.returncode != 0 or report.get("status") != "pass":
        raise AssertionError(
            f"Generated repository failed repository quality:\n{completed.stdout}{completed.stderr}"
        )
    print(
        f"PASS: generated {profile} repository initialization is recorded, clean, "
        "idempotent, and quality-valid."
    )


def self_test(source: Path) -> None:
    if _load_config(source).get("mode") == "generated":
        _generated_self_test(source)
        return
    with tempfile.TemporaryDirectory(prefix="repository-initializer-") as temporary:
        base = Path(temporary)
        for profile in SUPPORTED_PROFILES:
            fixture = base / profile
            _copy_fixture(source, fixture)
            scalars, repeatable = _fixture_arguments(profile)
            preview = fixture / "assets/social-preview.png"
            preview.write_bytes(b"fixture-preview\n")
            _git(fixture, "add", "--all")
            _git(fixture, "commit", "-m", "Add social-preview fixture asset")
            scalars.append("SOCIAL_PREVIEW_PATH=assets/social-preview.png")

            _assert_failure_without_change(
                fixture,
                profile,
                [item for item in scalars if not item.startswith("PROJECT_NAME=")],
                repeatable,
                "Missing required substitutions: PROJECT_NAME",
            )
            _assert_failure_without_change(
                fixture,
                profile,
                scalars + ["UNKNOWN_VALUE=not-registered"],
                repeatable,
                "Unknown substitutions: UNKNOWN_VALUE",
            )
            unused_fixture = base / f"{profile}-unused"
            _make_unused_fixture(source, unused_fixture)
            _assert_failure_without_change(
                unused_fixture,
                profile,
                [
                    item
                    for item in scalars
                    if not item.startswith("SOCIAL_PREVIEW_PATH=")
                ]
                + ["SOCIAL_PREVIEW_PATH=assets/social-preview.png"],
                repeatable,
                "Unused substitutions: SOCIAL_PREVIEW_PATH",
            )
            svg_scalars = [
                item
                for item in scalars
                if not item.startswith("SOCIAL_PREVIEW_PATH=")
            ] + ["SOCIAL_PREVIEW_PATH=assets/social-preview.svg"]
            _assert_failure_without_change(
                fixture,
                profile,
                svg_scalars,
                repeatable,
                "SOCIAL_PREVIEW_PATH may retain only assets/social-preview.png",
            )

            before = _tree_digest(fixture)
            changes, _ = _build_changes(fixture, profile, scalars, repeatable)
            if not changes or _tree_digest(fixture) != before:
                raise AssertionError(f"{profile} dry-run was empty or changed the tree.")
            repeated_changes, _ = _build_changes(fixture, profile, scalars, repeatable)
            if changes != repeated_changes or _plan(fixture, profile, changes) != _plan(
                fixture, profile, repeated_changes
            ):
                raise AssertionError(f"{profile} dry-run plan was not deterministic.")
            _apply_changes(fixture, changes)
            immediate_rerun, _ = _build_changes(fixture, profile, scalars, repeatable)
            if immediate_rerun:
                raise AssertionError(f"{profile} immediate second initialization was not idempotent.")
            _git(fixture, "add", "--all")
            _git(fixture, "commit", "-m", f"Initialize {profile} fixture")

            rerun, _ = _build_changes(fixture, profile, scalars, repeatable)
            if rerun:
                raise AssertionError(f"{profile} second initialization was not idempotent.")
            generated_config = json.loads(
                (fixture / CONFIG_PATH).read_text(encoding="utf-8")
            )
            retained_template_only = {CANONICAL_SOCIAL_PREVIEW_PATH}
            forbidden_template_only = set(
                generated_config["placeholders"]["template_only_paths"]
            ) - retained_template_only
            retained = sorted(
                path for path in forbidden_template_only if (fixture / path).exists()
            )
            if retained:
                raise AssertionError(
                    f"{profile} retained template-only files: {', '.join(retained)}"
                )
            if not (fixture / "assets/social-preview.png").is_file():
                raise AssertionError(f"{profile} removed its selected social preview.")
            _assert_generated_cleanup(fixture, profile)
            _assert_fresh_generated_content(fixture, profile)
            _generated_self_test(fixture)

            evolved = base / f"{profile}-evolved"
            _make_evolved_generated_fixture(fixture, evolved)
            evolved_profile, evolved_scalars, evolved_repeatable = _record_arguments(
                evolved
            )
            evolved_rerun, _ = _build_changes(
                evolved,
                evolved_profile,
                evolved_scalars,
                evolved_repeatable,
            )
            if evolved_rerun:
                raise AssertionError(
                    f"{profile} evolved generated-project rerun was not idempotent."
                )
            _generated_self_test(evolved)

            completed, report = _quality_report(fixture)
            if completed.returncode != 0:
                raise AssertionError(
                    f"{profile} generated fixture failed repository quality:\n{completed.stdout}{completed.stderr}"
                )
            contract = next(
                result for result in report["rules"]
                if result["id"] == "generated-vba-contract"
            )
            evidence = contract.get("evidence", {})
            if evidence.get("selected_profile") != profile or set(
                evidence.get("profiles", {})
            ) != {profile}:
                raise AssertionError(
                    f"{profile} quality evidence did not resolve only the selected profile."
                )

            mandatory = {
                "src/core/ProjectCore.bas": "internal",
                "src/modules/ProjectFacade.bas": "public",
                "tests/modules/ProjectTests.bas": "test",
            }
            readme_only = base / f"{profile}-readme-only"
            _make_component_variant(
                fixture,
                readme_only,
                tuple(
                    sorted(
                        json.loads(
                            (fixture / CONFIG_PATH).read_text(encoding="utf-8")
                        )["vba"]["components"]
                    )
                ),
            )
            _assert_contract_failure(
                readme_only,
                profile,
                ("internal", "public", "test"),
            )
            for path, role in mandatory.items():
                missing = base / f"{profile}-missing-{role}"
                _make_component_variant(fixture, missing, (path,))
                _assert_contract_failure(missing, profile, (role,))

            optional = base / f"{profile}-without-example"
            _make_component_variant(
                fixture,
                optional,
                ("examples/modules/ProjectExample.bas",),
            )
            optional_completed, optional_report = _quality_report(optional)
            if optional_completed.returncode != 0 or optional_report["status"] != "pass":
                raise AssertionError(
                    f"{profile} incorrectly required the optional example component:\n"
                    f"{optional_completed.stdout}{optional_completed.stderr}"
                )
            print(
                f"[PASS] {profile}: initialization, substantive contract, README-only rejection, "
                "mandatory-component removals, optional-component absence, and quality evidence"
            )
    print(
        f"PASS: {len(SUPPORTED_PROFILES)} profile fixtures initialized; "
        "12 mandatory/README-only removals rejected and 3 optional removals accepted."
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd(), help="Repository root.")
    parser.add_argument("--profile", choices=SUPPORTED_PROFILES, help="One repository profile.")
    parser.add_argument("--set", dest="scalar_entries", action="append", default=[], metavar="NAME=value")
    parser.add_argument("--add", dest="repeatable_entries", action="append", default=[], metavar="NAME=value")
    parser.add_argument("--apply", action="store_true", help="Apply the validated plan; dry-run is default.")
    parser.add_argument("--self-test", action="store_true", help="Exercise all profile and failure fixtures.")
    return parser


def main(argv: list[str] | None = None) -> int:
    arguments = build_parser().parse_args(argv)
    root = arguments.root.resolve()
    try:
        if arguments.self_test:
            self_test(root)
            return 0
        if arguments.profile is None:
            raise InitializationError("--profile is required unless --self-test is used.")
        status = _git(root, "status", "--porcelain", check=False)
        if status.returncode != 0:
            raise InitializationError("--root must be a Git working tree.")
        config = _load_config(root)
        if status.stdout and config.get("mode") == "template":
            raise InitializationError("Working tree must be clean before initialization.")
        changes, _ = _build_changes(
            root,
            arguments.profile,
            arguments.scalar_entries,
            arguments.repeatable_entries,
        )
        plan = _plan(root, arguments.profile, changes)
        if arguments.apply:
            _apply_changes(root, changes)
            plan["mode"] = "apply"
            plan["status"] = "applied" if changes else "no-op"
        print(json.dumps(plan, indent=2, ensure_ascii=False))
        return 0
    except (InitializationError, OSError, subprocess.SubprocessError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
