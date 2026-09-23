#!/usr/bin/env python3
"""Canonical, dependency-free repository-quality gate for Excel/VBA projects.

The checker validates facts available from the tracked repository tree. It does
not compile VBA, execute Excel, prove numerical accuracy, exercise UI state, or
certify release binaries. Those responsibilities belong to profile and project
gates that consume the generic baseline rather than weakening it.

Exit status is 0 for a clean repository, 1 for policy findings, and 2 for an
operational error.
"""

from __future__ import annotations

import argparse
import fnmatch
import hashlib
import json
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path, PurePosixPath
from typing import Any, Callable, Iterable
from urllib.parse import unquote, urlsplit

SCHEMA_VERSION = 1
CONFIG_PATH = ".github/repository-profile.json"
LABEL_MANIFEST_PATH = ".github/labels.json"
ISSUE_TEMPLATE_DIRECTORY = ".github/ISSUE_TEMPLATE"
TOOL_NAME = "Canonical repository quality"
MAX_XML_BYTES = 1024 * 1024
UNSAFE_XML_DECLARATION = re.compile(r"<!\s*(?:DOCTYPE|ENTITY)\b", re.IGNORECASE)
SUPPORTED_PROFILES = ("application", "library", "ui-component")
PLACEHOLDER_CATEGORIES = ("optional", "profile-specific", "repeatable", "required")
PLACEHOLDER_NAME_PATTERN = re.compile(r"[A-Z][A-Z0-9_]*")
VBA_SUFFIXES = {".bas", ".cls", ".frm"}
VBA_ROLES = ("example", "internal", "public", "test", "ui")
VBA_BASELINE_ROLES = ("internal", "public", "test")
PLACEHOLDER_PROHIBITED_SUFFIXES = {
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
VBA_COMPONENT_NAME_LIMIT = 31
VBA_HEADER_SCAN_LINES = 20

TEXT_SUFFIXES = {
    ".bas",
    ".bat",
    ".cfg",
    ".cff",
    ".cjs",
    ".cls",
    ".cmd",
    ".csv",
    ".frm",
    ".ini",
    ".js",
    ".json",
    ".jsonc",
    ".md",
    ".mjs",
    ".ps1",
    ".psd1",
    ".psm1",
    ".py",
    ".pyw",
    ".r",
    ".reg",
    ".sh",
    ".svg",
    ".toml",
    ".tsv",
    ".txt",
    ".vbs",
    ".xml",
    ".yaml",
    ".yml",
}
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
CROSS_PLATFORM_SUFFIXES = {
    ".cfg",
    ".cff",
    ".cjs",
    ".csv",
    ".js",
    ".json",
    ".jsonc",
    ".md",
    ".mjs",
    ".py",
    ".pyw",
    ".r",
    ".sh",
    ".svg",
    ".toml",
    ".tsv",
    ".txt",
    ".xml",
    ".yaml",
    ".yml",
}
WINDOWS_TEXT_SUFFIXES = {
    ".bas",
    ".bat",
    ".cls",
    ".cmd",
    ".frm",
    ".ini",
    ".ps1",
    ".psd1",
    ".psm1",
    ".reg",
    ".vbs",
}
OFFICE_BINARY_SUFFIXES = {
    ".accdb",
    ".accde",
    ".accdr",
    ".ade",
    ".doc",
    ".docm",
    ".docx",
    ".dot",
    ".dotm",
    ".dotx",
    ".mdb",
    ".mde",
    ".pot",
    ".potm",
    ".potx",
    ".ppam",
    ".pps",
    ".ppsm",
    ".ppsx",
    ".ppt",
    ".pptm",
    ".pptx",
    ".thmx",
    ".xla",
    ".xlam",
    ".xll",
    ".xls",
    ".xlsb",
    ".xlsm",
    ".xlsx",
    ".xlt",
    ".xltm",
    ".xltx",
    ".xlw",
}
SECRET_SUFFIXES = {".key", ".p12", ".pem", ".pfx", ".pvk"}
CONFIG_KEYS = {
    "schema_version",
    "mode",
    "profile",
    "repository",
    "template_contract",
    "label_domains",
    "required_paths",
    "required_directories",
    "profiles",
    "allowed_office_binary_globs",
    "placeholders",
    "identity",
    "vba",
}


class Repository:
    """Read-only view of a Git working tree and its tracked paths."""

    def __init__(self, root: Path):
        self.root = root.resolve()
        self.files = self._tracked_files()

    def _git(
        self, *arguments: str, check: bool = True
    ) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            ["git", "-C", str(self.root), *arguments],
            check=check,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

    def _tracked_files(self) -> tuple[str, ...]:
        completed = subprocess.run(
            ["git", "-C", str(self.root), "ls-files", "-z"],
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

    def path(self, relative: str) -> Path:
        return self.root / PurePosixPath(relative)

    def bytes(self, relative: str) -> bytes:
        return self.path(relative).read_bytes()

    def text(self, relative: str) -> str:
        data = self.bytes(relative)
        if PurePosixPath(relative).suffix.casefold() in VBA_SUFFIXES:
            return data.decode("cp1252")
        return data.decode("utf-8")

    def commit(self) -> str | None:
        completed = self._git("rev-parse", "HEAD", check=False)
        if completed.returncode != 0:
            return None
        return completed.stdout.strip() or None


def finding(
    path: str, message: str, line: int | None = None
) -> dict[str, Any]:
    item: dict[str, Any] = {"path": path, "message": message}
    if line is not None:
        item["line"] = line
    return item


def rule_result(
    rule_id: str,
    title: str,
    failures: list[dict[str, Any]],
    success_summary: str,
) -> dict[str, Any]:
    if failures:
        count = len(failures)
        return {
            "id": rule_id,
            "title": title,
            "status": "fail",
            "summary": f"{count} finding{'s' if count != 1 else ''}",
            "findings": failures,
        }
    return {
        "id": rule_id,
        "title": title,
        "status": "pass",
        "summary": success_summary,
        "findings": [],
    }


def line_number(text: str, offset: int) -> int:
    return text.count("\n", 0, offset) + 1


def is_text_file(path: str) -> bool:
    pure = PurePosixPath(path)
    return pure.suffix.casefold() in TEXT_SUFFIXES or pure.name in TEXT_NAMES


def is_under(path: str, roots: Iterable[str]) -> bool:
    return any(path == root or path.startswith(root.rstrip("/") + "/") for root in roots)


def _same_keys(value: object, expected: set[str]) -> bool:
    return isinstance(value, dict) and set(value) == expected


def _valid_relative_path(value: str) -> bool:
    pure = PurePosixPath(value)
    return (
        bool(value)
        and value == pure.as_posix()
        and not pure.is_absolute()
        and ".." not in pure.parts
        and "." not in pure.parts
    )


def _string_list(
    value: object,
    field: str,
    failures: list[dict[str, Any]],
    *,
    paths: bool = False,
) -> list[str]:
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        failures.append(finding(CONFIG_PATH, f"{field} must be an array of strings."))
        return []
    items = list(value)
    if len(items) != len(set(items)):
        failures.append(finding(CONFIG_PATH, f"{field} must not contain duplicates."))
    if items != sorted(items, key=lambda item: (item.casefold(), item)):
        failures.append(
            finding(CONFIG_PATH, f"{field} must be sorted case-insensitively.")
        )
    if paths:
        for item in items:
            if not _valid_relative_path(item):
                failures.append(
                    finding(CONFIG_PATH, f"{field} contains an invalid relative path: {item!r}.")
                )
    return items


def _validate_profile_contract(
    name: str, entry: object, failures: list[dict[str, Any]]
) -> None:
    if not _same_keys(entry, {"required_paths", "required_directories", "vba_contract"}):
        failures.append(finding(
            CONFIG_PATH,
            f"profiles.{name} must contain exactly required_paths, required_directories, and vba_contract.",
        ))
        return
    assert isinstance(entry, dict)
    _string_list(entry.get("required_paths"), f"profiles.{name}.required_paths", failures, paths=True)
    _string_list(entry.get("required_directories"), f"profiles.{name}.required_directories", failures, paths=True)
    contract = entry.get("vba_contract")
    field = f"profiles.{name}.vba_contract"
    if not _same_keys(contract, {"minimum_roles", "required_components"}):
        failures.append(finding(CONFIG_PATH, f"{field} must contain exactly minimum_roles and required_components."))
        return
    assert isinstance(contract, dict)
    minimum_roles = contract.get("minimum_roles")
    if not isinstance(minimum_roles, dict) or not minimum_roles:
        failures.append(finding(CONFIG_PATH, f"{field}.minimum_roles must be a non-empty object."))
        minimum_roles = {}
    else:
        role_names = list(minimum_roles)
        if role_names != sorted(role_names, key=lambda item: (item.casefold(), item)):
            failures.append(finding(CONFIG_PATH, f"{field}.minimum_roles keys must be sorted case-insensitively."))
        for role, minimum in minimum_roles.items():
            if role not in VBA_ROLES:
                failures.append(finding(CONFIG_PATH, f"{field}.minimum_roles has invalid role {role!r}."))
            if isinstance(minimum, bool) or not isinstance(minimum, int) or minimum < 1:
                failures.append(finding(CONFIG_PATH, f"{field}.minimum_roles.{role} must be a positive integer."))
    required_components = contract.get("required_components")
    if not isinstance(required_components, dict) or not required_components:
        failures.append(finding(CONFIG_PATH, f"{field}.required_components must be a non-empty object."))
        required_components = {}
    else:
        paths = list(required_components)
        if paths != sorted(paths, key=lambda item: (item.casefold(), item)):
            failures.append(finding(CONFIG_PATH, f"{field}.required_components keys must be sorted case-insensitively."))
        for path, role in required_components.items():
            if not isinstance(path, str) or not _valid_relative_path(path):
                failures.append(finding(CONFIG_PATH, f"{field}.required_components contains an invalid path: {path!r}."))
            if role not in VBA_ROLES:
                failures.append(finding(CONFIG_PATH, f"{field}.required_components.{path} has invalid role {role!r}."))
    component_roles = {role for role in required_components.values() if isinstance(role, str)}
    for role in VBA_BASELINE_ROLES:
        if minimum_roles.get(role, 0) < 1:
            failures.append(finding(CONFIG_PATH, f"{field}.minimum_roles must require the baseline role {role!r}."))
        if role not in component_roles:
            failures.append(finding(CONFIG_PATH, f"{field}.required_components must name a component with role {role!r}."))


def _validate_profiles(document: dict[str, Any], failures: list[dict[str, Any]]) -> None:
    profiles = document.get("profiles")
    if not isinstance(profiles, dict) or set(profiles) != set(SUPPORTED_PROFILES):
        failures.append(finding(CONFIG_PATH, "profiles must contain exactly application, library, and ui-component."))
        return
    for name in SUPPORTED_PROFILES:
        _validate_profile_contract(name, profiles[name], failures)


def _validate_placeholder_spec(
    name: object,
    specification: object,
    categories_seen: set[str],
    failures: list[dict[str, Any]],
) -> None:
    field = f"placeholders.catalogue.{name}"
    if not isinstance(name, str) or not PLACEHOLDER_NAME_PATTERN.fullmatch(name):
        failures.append(finding(CONFIG_PATH, f"{field} is not a canonical placeholder name."))
        return
    if not isinstance(specification, dict):
        failures.append(finding(CONFIG_PATH, f"{field} must be an object."))
        return
    category = specification.get("category")
    description = specification.get("description")
    if category not in PLACEHOLDER_CATEGORIES:
        failures.append(finding(CONFIG_PATH, f"{field}.category must be optional, profile-specific, repeatable, or required."))
        return
    categories_seen.add(category)
    if not isinstance(description, str) or not description.strip():
        failures.append(finding(CONFIG_PATH, f"{field}.description must be non-empty."))
    expected_keys = {"category", "description"}
    if category == "profile-specific":
        expected_keys.add("values")
        values = specification.get("values")
        if not isinstance(values, dict) or set(values) != set(SUPPORTED_PROFILES):
            failures.append(finding(CONFIG_PATH, f"{field}.values must cover exactly all supported profiles."))
        elif any(not isinstance(value, str) or not value.strip() for value in values.values()):
            failures.append(finding(CONFIG_PATH, f"{field}.values must all be non-empty strings."))
    elif category == "repeatable":
        expected_keys.add("item_format")
        item_format = specification.get("item_format")
        if not isinstance(item_format, str) or item_format.count("{value}") != 1:
            failures.append(finding(CONFIG_PATH, f"{field}.item_format must contain one {{value}} field."))
    if set(specification) != expected_keys:
        failures.append(finding(CONFIG_PATH, f"{field} has fields inconsistent with its category."))


def _validate_placeholders(document: dict[str, Any], failures: list[dict[str, Any]]) -> None:
    placeholders = document.get("placeholders")
    keys = {"pattern", "catalogue", "block_markers", "template_only_paths", "exclude_paths"}
    if not _same_keys(placeholders, keys):
        failures.append(finding(CONFIG_PATH, "placeholders must contain exactly pattern, catalogue, block_markers, template_only_paths, and exclude_paths."))
        return
    assert isinstance(placeholders, dict)
    pattern = placeholders.get("pattern")
    if not isinstance(pattern, str):
        failures.append(finding(CONFIG_PATH, "placeholders.pattern must be a string."))
    else:
        try:
            compiled = re.compile(pattern)
        except re.error as error:
            failures.append(finding(CONFIG_PATH, f"placeholders.pattern is invalid: {error}."))
        else:
            if compiled.groups != 1:
                failures.append(finding(CONFIG_PATH, "placeholders.pattern must contain exactly one capture group for the token name."))
    catalogue = placeholders.get("catalogue")
    categories_seen: set[str] = set()
    if not isinstance(catalogue, dict) or not catalogue:
        failures.append(finding(CONFIG_PATH, "placeholders.catalogue must be a non-empty object."))
    else:
        names = list(catalogue)
        if names != sorted(names, key=lambda item: (item.casefold(), item)):
            failures.append(finding(CONFIG_PATH, "placeholders.catalogue keys must be sorted case-insensitively."))
        for name, specification in catalogue.items():
            _validate_placeholder_spec(name, specification, categories_seen, failures)
        missing_categories = set(PLACEHOLDER_CATEGORIES) - categories_seen
        if missing_categories:
            failures.append(finding(CONFIG_PATH, "placeholders.catalogue does not exercise categories: " + ", ".join(sorted(missing_categories))))
    expected_markers = {
        "template_only": "template:remove",
        "profile": "template:profile:{profile}",
        "optional": "template:optional:{token}",
        "repeatable": "template:repeatable:{token}",
    }
    if placeholders.get("block_markers") != expected_markers:
        failures.append(finding(CONFIG_PATH, "placeholders.block_markers must use the canonical marker grammar."))
    _string_list(placeholders.get("template_only_paths"), "placeholders.template_only_paths", failures, paths=True)
    _string_list(placeholders.get("exclude_paths"), "placeholders.exclude_paths", failures, paths=True)


def _validate_identity(
    document: dict[str, Any], mode: object, repository: object, failures: list[dict[str, Any]]
) -> None:
    identity = document.get("identity")
    keys = {"forbidden_tokens", "template_tokens", "exclude_paths"}
    if not _same_keys(identity, keys):
        failures.append(finding(CONFIG_PATH, "identity must contain exactly forbidden_tokens, template_tokens, and exclude_paths."))
        return
    assert isinstance(identity, dict)
    forbidden = _string_list(identity.get("forbidden_tokens"), "identity.forbidden_tokens", failures)
    template = _string_list(identity.get("template_tokens"), "identity.template_tokens", failures)
    _string_list(identity.get("exclude_paths"), "identity.exclude_paths", failures, paths=True)
    if not template:
        failures.append(finding(CONFIG_PATH, "identity.template_tokens must not be empty."))
    if not isinstance(repository, str):
        return
    folded = repository.casefold()
    for token in forbidden:
        if token.casefold() in folded:
            failures.append(finding(CONFIG_PATH, f"repository contains a forbidden donor token: {token}"))
    contains_template = any(token.casefold() in folded for token in template)
    if mode == "template" and template and not contains_template:
        failures.append(finding(CONFIG_PATH, "Template mode repository must contain a declared template identity token."))
    elif mode == "generated" and contains_template:
        failures.append(finding(CONFIG_PATH, "Generated mode repository still contains a template identity token."))


def _validate_vba_configuration(document: dict[str, Any], failures: list[dict[str, Any]]) -> None:
    vba = document.get("vba")
    keys = {"source_roots", "test_roots", "components", "public_api_manifest"}
    if not _same_keys(vba, keys):
        failures.append(finding(CONFIG_PATH, "vba must contain exactly source_roots, test_roots, components, and public_api_manifest."))
        return
    assert isinstance(vba, dict)
    source_roots = _string_list(vba.get("source_roots"), "vba.source_roots", failures, paths=True)
    test_roots = _string_list(vba.get("test_roots"), "vba.test_roots", failures, paths=True)
    if set(source_roots).intersection(test_roots):
        failures.append(finding(CONFIG_PATH, "VBA source and test roots must not overlap."))
    components = vba.get("components")
    if not isinstance(components, dict):
        failures.append(finding(CONFIG_PATH, "vba.components must be an object."))
    else:
        component_keys = list(components)
        if component_keys != sorted(component_keys, key=lambda item: (item.casefold(), item)):
            failures.append(finding(CONFIG_PATH, "vba.components keys must be sorted case-insensitively."))
        for path, role in components.items():
            if not isinstance(path, str) or not _valid_relative_path(path):
                failures.append(finding(CONFIG_PATH, f"Invalid VBA component path: {path!r}."))
            if role not in VBA_ROLES:
                failures.append(finding(CONFIG_PATH, f"VBA component {path!r} has unsupported role {role!r}."))
    api_manifest = vba.get("public_api_manifest")
    if api_manifest is not None and (not isinstance(api_manifest, str) or not _valid_relative_path(api_manifest)):
        failures.append(finding(CONFIG_PATH, "vba.public_api_manifest must be null or a valid relative path."))


def _validate_configuration_root(
    document: dict[str, Any], failures: list[dict[str, Any]]
) -> tuple[object, object]:
    if document.get("schema_version") != SCHEMA_VERSION:
        failures.append(finding(CONFIG_PATH, f"schema_version must be {SCHEMA_VERSION}."))
    mode = document.get("mode")
    profile = document.get("profile")
    if mode not in {"template", "generated"}:
        failures.append(finding(CONFIG_PATH, "mode must be template or generated."))
    elif mode == "template" and profile is not None:
        failures.append(finding(CONFIG_PATH, "Template mode requires profile to be null."))
    elif mode == "generated" and profile not in SUPPORTED_PROFILES:
        failures.append(finding(CONFIG_PATH, "Generated mode requires profile to be application, library, or ui-component."))
    repository = document.get("repository")
    if not isinstance(repository, str) or not re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repository):
        failures.append(finding(CONFIG_PATH, "repository must use the owner/name form."))
    label_domains = _string_list(document.get("label_domains"), "label_domains", failures)
    for domain in label_domains:
        if not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", domain):
            failures.append(finding(CONFIG_PATH, f"label_domains contains a non-kebab-case name: {domain!r}."))
    if mode == "template" and label_domains:
        failures.append(finding(CONFIG_PATH, "Template mode requires label_domains to be empty."))
    _string_list(document.get("required_paths"), "required_paths", failures, paths=True)
    _string_list(document.get("required_directories"), "required_directories", failures, paths=True)
    _string_list(document.get("allowed_office_binary_globs"), "allowed_office_binary_globs", failures)
    return mode, repository


def load_configuration(
    repo: Repository,
) -> tuple[dict[str, Any] | None, dict[str, Any]]:
    failures: list[dict[str, Any]] = []
    try:
        document = json.loads(repo.text(CONFIG_PATH))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        failures.append(finding(CONFIG_PATH, f"Cannot load repository profile: {error}", getattr(error, "lineno", None)))
        return None, rule_result("configuration", "Repository profile configuration", failures, "")
    if not _same_keys(document, CONFIG_KEYS):
        failures.append(finding(CONFIG_PATH, "Root object must contain exactly the canonical configuration keys."))
    if not isinstance(document, dict):
        document = {}
    mode, repository = _validate_configuration_root(document, failures)
    _validate_profiles(document, failures)
    _validate_placeholders(document, failures)
    _validate_identity(document, mode, repository, failures)
    _validate_vba_configuration(document, failures)
    return (
        document if not failures else None,
        rule_result(
            "configuration",
            "Repository profile configuration",
            failures,
            "Versioned template/profile configuration is valid",
        ),
    )


def _effective_requirements(
    config: dict[str, Any],
) -> tuple[list[str], list[str]]:
    paths = list(config["required_paths"])
    directories = list(config["required_directories"])
    if config["mode"] == "generated":
        profile = config["profiles"][config["profile"]]
        paths.extend(profile["required_paths"])
        directories.extend(profile["required_directories"])
    return (
        sorted(set(paths), key=lambda item: (item.casefold(), item)),
        sorted(set(directories), key=lambda item: (item.casefold(), item)),
    )


def check_required_paths(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    tracked = set(repo.files)
    paths, directories = _effective_requirements(config)
    for path in paths:
        if path not in tracked:
            failures.append(finding(path, "Required file is not tracked."))
        elif not repo.path(path).is_file():
            failures.append(finding(path, "Required tracked file is absent from the worktree."))
    for directory in directories:
        prefix = directory.rstrip("/") + "/"
        members = [path for path in repo.files if path.startswith(prefix)]
        if not repo.path(directory).is_dir():
            failures.append(finding(directory, "Required directory is absent."))
        elif not members:
            failures.append(
                finding(
                    directory,
                    "Required directory has no tracked explanatory or substantive file.",
                )
            )
    return rule_result(
        "required-paths",
        "Required files and directories",
        failures,
        f"Validated {len(paths)} required files and {len(directories)} directories",
    )


def _markdown_link_label(text: str, end: int) -> bool:
    return end < len(text) and text[end] == "("


def check_placeholders(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    settings = config["placeholders"]
    pattern = re.compile(settings["pattern"])
    excluded = set(settings["exclude_paths"])
    allowed = {"{{" + name + "}}" for name in settings["catalogue"]}
    seen: set[str] = set()
    failures: list[dict[str, Any]] = []
    checked = 0
    for path in repo.files:
        if path in excluded or not is_text_file(path):
            continue
        try:
            text = repo.text(path)
        except (OSError, UnicodeError):
            continue
        checked += 1
        for match in pattern.finditer(text):
            if (
                PurePosixPath(path).suffix.casefold() == ".md"
                and _markdown_link_label(text, match.end())
            ):
                continue
            token = match.group(0)
            seen.add(token)
            if (
                PurePosixPath(path).suffix.casefold()
                in PLACEHOLDER_PROHIBITED_SUFFIXES
            ):
                failures.append(
                    finding(
                        path,
                        f"Template placeholders are prohibited in executable or VBA files: {token}",
                        line_number(text, match.start()),
                    )
                )
            if config["mode"] == "generated":
                failures.append(
                    finding(
                        path,
                        f"Unresolved template placeholder: {token}",
                        line_number(text, match.start()),
                    )
                )
            elif token not in allowed:
                failures.append(
                    finding(
                        path,
                        f"Template placeholder is not registered: {token}",
                        line_number(text, match.start()),
                    )
                )
    if config["mode"] == "template":
        for token in sorted(allowed):
            if token not in seen:
                failures.append(finding(CONFIG_PATH, f"Registered template placeholder is unused: {token}"))
    summary = (
        f"Validated {len(seen)} registered placeholders across {checked} text files"
        if config["mode"] == "template"
        else f"No unresolved placeholders in {checked} text files"
    )
    return rule_result(
        "placeholders",
        "Template placeholder governance",
        failures,
        summary,
    )


def _identity_scan_text(path: str, text: str, config: dict[str, Any], token: str) -> str:
    """Mask only immutable upstream workflow references, not donor branding."""
    contract = config.get("template_contract")
    if (
        config["mode"] != "generated"
        or token not in config["identity"]["template_tokens"]
        or token in config["identity"]["forbidden_tokens"]
        or not path.startswith(".github/workflows/")
        or Path(path).suffix not in {".yml", ".yaml"}
        or not isinstance(contract, dict)
        or not isinstance(contract.get("source"), str)
    ):
        return text
    source = re.escape(contract["source"])
    pattern = rf"(?<![\w/.-]){source}/\.github/workflows/[\w-]+\.ya?ml@[0-9a-f]{{40}}(?=$|[\s\"'])"
    return re.sub(pattern, lambda match: " " * len(match.group()), text)


def check_identity(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    settings = config["identity"]
    tokens = list(settings["forbidden_tokens"])
    if config["mode"] == "generated":
        tokens.extend(settings["template_tokens"])
    excluded = set(settings["exclude_paths"])
    failures: list[dict[str, Any]] = []
    checked = 0
    for path in repo.files:
        if path in excluded or not is_text_file(path):
            continue
        try:
            text = repo.text(path)
        except (OSError, UnicodeError):
            continue
        checked += 1
        for token in tokens:
            folded = _identity_scan_text(path, text, config, token).casefold()
            offset = folded.find(token.casefold())
            if offset >= 0:
                failures.append(
                    finding(
                        path,
                        f"Forbidden donor or template identity token is present: {token}",
                        line_number(text, offset),
                    )
                )
    return rule_result(
        "identity",
        "Donor and template identity isolation",
        failures,
        f"No forbidden identity tokens in {checked} text files",
    )


def _editorconfig_sections(text: str) -> dict[str, dict[str, str]]:
    sections: dict[str, dict[str, str]] = {"": {}}
    current = ""
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith(("#", ";")):
            continue
        if line.startswith("[") and line.endswith("]"):
            current = line[1:-1]
            sections.setdefault(current, {})
            continue
        if "=" in line:
            key, value = line.split("=", 1)
            sections[current][key.strip().casefold()] = value.strip().casefold()
    return sections


def _git_attribute(repo: Repository, attribute: str, probe: str) -> str | None:
    completed = repo._git("check-attr", attribute, "--", probe, check=False)
    if completed.returncode != 0:
        return None
    parts = completed.stdout.rstrip("\n").split(": ", 2)
    return parts[2] if len(parts) == 3 else None


def _is_ignored(repo: Repository, probe: str) -> bool:
    completed = repo._git(
        "check-ignore", "--no-index", "--quiet", "--", probe, check=False
    )
    return completed.returncode == 0


def check_dotfile_policy(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    del config
    failures: list[dict[str, Any]] = []
    try:
        sections = _editorconfig_sections(repo.text(".editorconfig"))
    except (OSError, UnicodeError) as error:
        sections = {}
        failures.append(finding(".editorconfig", f"Cannot read policy: {error}"))
    if sections.get("", {}).get("root") != "true":
        failures.append(finding(".editorconfig", "root must be true."))
    vba_section = sections.get("*.{bas,cls,frm}", {})
    if vba_section.get("end_of_line") != "crlf":
        failures.append(
            finding(
                ".editorconfig",
                "The VBA component section must require CRLF line endings.",
            )
        )
    if vba_section.get("insert_final_newline") != "true":
        failures.append(
            finding(
                ".editorconfig",
                "The VBA component section must require a final newline.",
            )
        )

    for suffix in ("bas", "cls", "frm"):
        probe = f"quality-probe.{suffix}"
        if _git_attribute(repo, "eol", probe) != "crlf":
            failures.append(
                finding(".gitattributes", f"{probe} must resolve to eol=crlf.")
            )
    for suffix in ("json", "md", "py", "yml"):
        probe = f"quality-probe.{suffix}"
        if _git_attribute(repo, "eol", probe) != "lf":
            failures.append(
                finding(".gitattributes", f"{probe} must resolve to eol=lf.")
            )
    if _git_attribute(repo, "text", "quality-probe.xlsm") != "unset":
        failures.append(
            finding(
                ".gitattributes",
                "Office packages must resolve to text=unset (binary).",
            )
        )

    ignored = (
        ".env",
        "__pycache__/quality.pyc",
        "private-key.pem",
        "quality.xlsm",
        "test-results/quality.json",
        "~$quality.xlsx",
    )
    for probe in ignored:
        if not _is_ignored(repo, probe):
            failures.append(
                finding(".gitignore", f"Required generated/secret probe is not ignored: {probe}")
            )
    if _is_ignored(repo, ".env.example"):
        failures.append(
            finding(".gitignore", ".env.example must remain trackable.")
        )

    return rule_result(
        "dotfile-policy",
        "Editor, attributes, and ignore policy",
        failures,
        "Canonical EditorConfig, Git attributes, and ignore probes pass",
    )


def _strip_yaml_comment(value: str) -> str:
    single = False
    double = False
    escaped = False
    for index, character in enumerate(value):
        if escaped:
            escaped = False
            continue
        if character == "\\" and double:
            escaped = True
        elif character == "'" and not double:
            single = not single
        elif character == '"' and not single:
            double = not double
        elif character == "#" and not single and not double:
            if index == 0 or value[index - 1].isspace():
                return value[:index].rstrip()
    return value.rstrip()


def _yaml_scalar_error(value: str) -> str | None:
    single = False
    double = False
    escaped = False
    stack: list[str] = []
    pairs = {"]": "[", "}": "{"}
    for character in value:
        if escaped:
            escaped = False
            continue
        if character == "\\" and double:
            escaped = True
        elif character == "'" and not double:
            single = not single
        elif character == '"' and not single:
            double = not double
        elif not single and not double and character in "[{":
            stack.append(character)
        elif not single and not double and character in "]}":
            if not stack or stack.pop() != pairs[character]:
                return "unbalanced flow brackets"
    if single or double:
        return "unterminated quoted scalar"
    if stack:
        return "unbalanced flow brackets"
    return None


def validate_yaml_subset(text: str) -> list[tuple[int, str]]:
    """Validate the conservative YAML dialect used by GitHub repository files."""

    errors: list[tuple[int, str]] = []
    block_parent_indent: int | None = None
    mapping = re.compile(
        r"^(?:[A-Za-z0-9_.$}{-]+|'[^']+'|\"[^\"]+\")\s*:(?:\s*(.*))?$"
    )
    for number, raw_line in enumerate(text.splitlines(), start=1):
        if not raw_line.strip() or raw_line.lstrip().startswith("#"):
            continue
        leading = raw_line[: len(raw_line) - len(raw_line.lstrip(" \t"))]
        if "\t" in leading:
            errors.append((number, "tab indentation is not allowed"))
            continue
        indent = len(leading)
        if block_parent_indent is not None:
            if indent > block_parent_indent:
                continue
            block_parent_indent = None
        if indent % 2:
            errors.append((number, "indentation must use two-space levels"))
            continue
        content = _strip_yaml_comment(raw_line[indent:])
        if not content:
            continue
        candidate = content
        if candidate == "-":
            continue
        if candidate.startswith("- "):
            candidate = candidate[2:].strip()
            if not candidate:
                continue
            if not mapping.match(candidate):
                scalar_error = _yaml_scalar_error(candidate)
                if scalar_error:
                    errors.append((number, scalar_error))
                continue
        match = mapping.match(candidate)
        if not match:
            errors.append((number, "expected a mapping entry or sequence item"))
            continue
        scalar = (match.group(1) or "").strip()
        if scalar in {"|", "|-", "|+", ">", ">-", ">+"}:
            block_parent_indent = indent
            continue
        scalar_error = _yaml_scalar_error(scalar)
        if scalar_error:
            errors.append((number, scalar_error))
    return errors


def _validate_xml_text(path: str, text: str) -> dict[str, Any] | None:
    size = len(text.encode("utf-8"))
    if size > MAX_XML_BYTES:
        return finding(
            path,
            f"XML input exceeds the {MAX_XML_BYTES}-byte structural-validation limit.",
        )
    match = UNSAFE_XML_DECLARATION.search(text)
    if match is not None:
        return finding(
            path,
            "XML DTD and entity declarations are prohibited by the portable checker.",
            line_number(text, match.start()),
        )
    try:
        ET.fromstring(text)  # noqa: S314 -- bounded input with DTD/entity declarations rejected above.
    except ET.ParseError as error:
        return finding(path, f"Invalid XML: {error}", error.position[0])
    return None


def check_structured_data(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    del config
    failures: list[dict[str, Any]] = []
    counts = {"json": 0, "yaml": 0, "xml": 0}
    for path in repo.files:
        suffix = PurePosixPath(path).suffix.casefold()
        if suffix == ".json":
            counts["json"] += 1
            try:
                json.loads(repo.text(path))
            except (OSError, UnicodeError, json.JSONDecodeError) as error:
                failures.append(
                    finding(path, f"Invalid JSON: {error}", getattr(error, "lineno", None))
                )
        elif suffix in {".yml", ".yaml"}:
            counts["yaml"] += 1
            try:
                text = repo.text(path)
            except (OSError, UnicodeError) as error:
                failures.append(finding(path, f"Cannot decode YAML as UTF-8: {error}"))
                continue
            for number, message in validate_yaml_subset(text):
                failures.append(
                    finding(path, f"Invalid YAML structure: {message}.", number)
                )
        elif suffix == ".xml":
            counts["xml"] += 1
            try:
                text = repo.text(path)
            except (OSError, UnicodeError) as error:
                failures.append(finding(path, f"Cannot decode XML as UTF-8: {error}"))
                continue
            xml_failure = _validate_xml_text(path, text)
            if xml_failure is not None:
                failures.append(xml_failure)
    summary = (
        f"Parsed {counts['json']} JSON, {counts['yaml']} YAML, "
        f"and {counts['xml']} XML files"
    )
    return rule_result(
        "structured-data",
        "JSON, YAML, and XML structure",
        failures,
        summary,
    )


def _markdown_destinations(text: str) -> Iterable[tuple[int, str]]:
    fenced = False
    fence = ""
    inline = re.compile(r"!?\[[^\]]*\]\(([^)]+)\)")
    reference = re.compile(r"^\s*\[[^\]]+\]:\s*(\S+)")
    code_fence = chr(96) * 3
    for number, line in enumerate(text.splitlines(), start=1):
        stripped = line.lstrip()
        if stripped.startswith((code_fence, "~~~")):
            marker = stripped[:3]
            if not fenced:
                fenced = True
                fence = marker
            elif marker == fence:
                fenced = False
            continue
        if fenced:
            continue
        for match in inline.finditer(line):
            yield number, match.group(1).strip()
        reference_match = reference.match(line)
        if reference_match:
            yield number, reference_match.group(1).strip()


def _split_destination(raw: str) -> tuple[str, str]:
    value = raw.strip()
    if value.startswith("<") and ">" in value:
        value = value[1 : value.index(">")]
    elif re.search(r"\s", value):
        value = value.split(None, 1)[0]
    parsed = urlsplit(value)
    return unquote(parsed.path), unquote(parsed.fragment)


def _github_slugs(text: str) -> set[str]:
    slugs: set[str] = set()
    counts: dict[str, int] = {}
    fenced = False
    fence = ""
    code_fence = chr(96) * 3
    for line in text.splitlines():
        stripped = line.lstrip()
        if stripped.startswith((code_fence, "~~~")):
            marker = stripped[:3]
            if not fenced:
                fenced = True
                fence = marker
            elif marker == fence:
                fenced = False
            continue
        if fenced:
            continue
        for anchor in re.finditer(
            r"<(?:a\s+(?:id|name)|[A-Za-z][^>]*\s+id)=[\"']([^\"']+)[\"']",
            line,
            re.IGNORECASE,
        ):
            slugs.add(anchor.group(1).casefold())
        match = re.match(r"^\s{0,3}#{1,6}\s+(.+?)\s*#*\s*$", line)
        if not match:
            continue
        heading = re.sub(r"<[^>]+>", "", match.group(1)).casefold()
        heading = re.sub(r"[*_~]", "", heading).replace(chr(96), "")
        heading = re.sub(r"[^\w\- ]", "", heading, flags=re.UNICODE)
        base = re.sub(r"\s+", "-", heading.strip())
        occurrence = counts.get(base, 0)
        counts[base] = occurrence + 1
        slugs.add(base if occurrence == 0 else f"{base}-{occurrence}")
    return slugs


def check_markdown_links(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    del config
    failures: list[dict[str, Any]] = []
    checked = 0
    slug_cache: dict[Path, set[str]] = {}
    for path in repo.files:
        if PurePosixPath(path).suffix.casefold() != ".md":
            continue
        try:
            text = repo.text(path)
        except (OSError, UnicodeError):
            continue
        source = repo.path(path)
        for number, raw in _markdown_destinations(text):
            parsed = urlsplit(raw.strip("<>"))
            if parsed.scheme or raw.startswith("//"):
                continue
            target_text, fragment = _split_destination(raw)
            if not target_text and not fragment:
                continue
            checked += 1
            target = source if not target_text else source.parent / target_text
            try:
                target = target.resolve()
                target.relative_to(repo.root)
            except (OSError, ValueError):
                failures.append(
                    finding(path, f"Relative link escapes the repository: {raw}", number)
                )
                continue
            if not target.exists():
                failures.append(
                    finding(path, f"Relative link target does not exist: {raw}", number)
                )
                continue
            relative_target = target.relative_to(repo.root).as_posix()
            if target.is_file() and relative_target not in repo.files:
                failures.append(
                    finding(path, f"Relative link target is not tracked: {raw}", number)
                )
                continue
            if target.is_dir() and not any(
                item.startswith(relative_target.rstrip("/") + "/")
                for item in repo.files
            ):
                failures.append(
                    finding(
                        path,
                        f"Relative directory link has no tracked content: {raw}",
                        number,
                    )
                )
                continue
            if fragment and target.is_file() and target.suffix.casefold() == ".md":
                if target not in slug_cache:
                    try:
                        slug_cache[target] = _github_slugs(
                            target.read_text(encoding="utf-8")
                        )
                    except (OSError, UnicodeError):
                        slug_cache[target] = set()
                if fragment.casefold() not in slug_cache[target]:
                    failures.append(
                        finding(path, f"Markdown heading does not exist: {raw}", number)
                    )
    return rule_result(
        "markdown-links",
        "Markdown relative links",
        failures,
        f"Resolved {checked} relative links and anchors",
    )


def check_text_integrity(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    del config
    failures: list[dict[str, Any]] = []
    conflict = re.compile(r"^(?:<{7}|={7}|>{7})(?:\s|$)", re.MULTILINE)
    private_key_marker = "-----BEGIN " + "PRIVATE KEY-----"
    github_token_prefix = "gh" + "p_"
    aws_key = re.compile("AK" + r"IA[0-9A-Z]{16}")
    checked = 0
    for path in repo.files:
        if not is_text_file(path):
            continue
        checked += 1
        try:
            data = repo.bytes(path)
        except OSError as error:
            failures.append(finding(path, f"Tracked text file cannot be read: {error}"))
            continue
        if b"\0" in data:
            failures.append(finding(path, "Tracked text file contains a NUL byte."))
            continue
        try:
            text = repo.text(path)
        except UnicodeError as error:
            failures.append(
                finding(path, f"Tracked text file has invalid encoding: {error}")
            )
            continue
        for match in conflict.finditer(text):
            failures.append(
                finding(
                    path,
                    "Merge-conflict marker is present.",
                    line_number(text, match.start()),
                )
            )
        for marker, label in (
            (private_key_marker, "private-key material"),
            (github_token_prefix, "GitHub token material"),
        ):
            offset = text.find(marker)
            if offset >= 0:
                failures.append(
                    finding(
                        path,
                        f"Possible {label} is tracked.",
                        line_number(text, offset),
                    )
                )
        aws_match = aws_key.search(text)
        if aws_match:
            failures.append(
                finding(
                    path,
                    "Possible AWS access key is tracked.",
                    line_number(text, aws_match.start()),
                )
            )
    return rule_result(
        "text-integrity",
        "Text integrity and secret markers",
        failures,
        f"Validated {checked} tracked text files",
    )


def check_forbidden_artifacts(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    allowed = config["allowed_office_binary_globs"]
    for path in repo.files:
        pure = PurePosixPath(path)
        lower = path.casefold()
        name = pure.name.casefold()
        suffix = pure.suffix.casefold()
        if name.startswith("~$") or suffix in {".laccdb", ".ldb"}:
            failures.append(finding(path, "Office lock file must not be tracked."))
        if suffix in OFFICE_BINARY_SUFFIXES and not any(
            fnmatch.fnmatchcase(path, pattern) for pattern in allowed
        ):
            failures.append(
                finding(
                    path,
                    "Office binary is not permitted by the profile allow-list.",
                )
            )
        if name == ".env" or (name.startswith(".env.") and name != ".env.example"):
            failures.append(
                finding(path, "Local environment or secret file must not be tracked.")
            )
        if suffix in SECRET_SUFFIXES:
            failures.append(
                finding(path, "Private key or certificate material must not be tracked.")
            )
        components = {component.casefold() for component in pure.parts}
        if components.intersection({"private", "private-review", "review-private"}):
            failures.append(finding(path, "Private review material must not be tracked."))
        if any(
            token in lower
            for token in ("private_review", "confidential-review", "internal-review")
        ):
            failures.append(finding(path, "Private review material must not be tracked."))
    return rule_result(
        "forbidden-artifacts",
        "Forbidden tracked artifacts",
        failures,
        "No unapproved Office binaries, locks, secrets, or private review files are tracked",
    )


def _has_bare_lf(data: bytes) -> bool:
    return bool(re.search(rb"(?<!\r)\n", data))


def _has_bare_cr(data: bytes) -> bool:
    return bool(re.search(rb"\r(?!\n)", data))


def check_line_endings(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    del config
    failures: list[dict[str, Any]] = []
    checked = 0
    for path in repo.files:
        if not is_text_file(path):
            continue
        pure = PurePosixPath(path)
        suffix = pure.suffix.casefold()
        if (
            suffix not in CROSS_PLATFORM_SUFFIXES | WINDOWS_TEXT_SUFFIXES
            and pure.name not in TEXT_NAMES
        ):
            continue
        checked += 1
        try:
            data = repo.bytes(path)
        except OSError:
            continue
        if data.startswith(b"\xef\xbb\xbf"):
            failures.append(
                finding(path, "UTF-8 BOM is not permitted by the encoding policy.")
            )
        if data and not data.endswith(b"\n"):
            failures.append(finding(path, "Text file must end with a newline."))
        if suffix in WINDOWS_TEXT_SUFFIXES:
            if _has_bare_lf(data) or _has_bare_cr(data):
                failures.append(
                    finding(path, "Windows/VBA source must use CRLF line endings only.")
                )
        elif b"\r" in data:
            failures.append(
                finding(path, "Cross-platform text must use LF line endings only.")
            )
    return rule_result(
        "line-endings",
        "Line endings and encoding",
        failures,
        f"Validated line-ending policy for {checked} tracked text files",
    )


def _compare_names(left: str, right: str) -> int:
    left_key = left.casefold()
    right_key = right.casefold()
    if left_key < right_key:
        return -1
    if left_key > right_key:
        return 1
    return (left > right) - (left < right)


def _validate_label_array(
    labels: object,
    location: str,
    seen: dict[str, str],
    failures: list[dict[str, Any]],
) -> None:
    if not isinstance(labels, list):
        failures.append(finding(LABEL_MANIFEST_PATH, f"{location} must be an array."))
        return
    names: list[str] = []
    for index, label in enumerate(labels):
        item = f"{location}[{index}]"
        if not _same_keys(label, {"name", "color", "description"}):
            failures.append(
                finding(
                    LABEL_MANIFEST_PATH,
                    f"{item} must contain exactly name, color, and description.",
                )
            )
            continue
        name = label.get("name")
        color = label.get("color")
        description = label.get("description")
        if (
            not isinstance(name, str)
            or not name
            or name != name.strip()
            or len(name) > 50
            or "\n" in name
            or "\r" in name
        ):
            failures.append(
                finding(
                    LABEL_MANIFEST_PATH,
                    f"{item}.name must be a trimmed single-line string of 1-50 characters.",
                )
            )
        else:
            key = name.casefold()
            if key in seen:
                failures.append(
                    finding(
                        LABEL_MANIFEST_PATH,
                        f"{item}.name duplicates {seen[key]} case-insensitively.",
                    )
                )
            else:
                seen[key] = item
            names.append(name)
        if not isinstance(color, str) or not re.fullmatch(r"[0-9A-F]{6}", color):
            failures.append(
                finding(
                    LABEL_MANIFEST_PATH,
                    f"{item}.color must be six uppercase hexadecimal characters.",
                )
            )
        if (
            not isinstance(description, str)
            or not description
            or len(description) > 100
            or "\n" in description
            or "\r" in description
        ):
            failures.append(
                finding(
                    LABEL_MANIFEST_PATH,
                    f"{item}.description must be a non-empty single-line string of at most 100 characters.",
                )
            )
    for previous, current in zip(names, names[1:]):
        if _compare_names(previous, current) >= 0:
            failures.append(
                finding(
                    LABEL_MANIFEST_PATH,
                    f"{location} must be sorted case-insensitively by name.",
                )
            )
            break


def check_label_manifest(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    try:
        document = json.loads(repo.text(LABEL_MANIFEST_PATH))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        return rule_result(
            "label-manifest",
            "Canonical label manifest",
            [
                finding(
                    LABEL_MANIFEST_PATH,
                    f"Cannot load label manifest: {error}",
                    getattr(error, "lineno", None),
                )
            ],
            "",
        )
    if not _same_keys(document, {"schema_version", "prune", "core", "overlays"}):
        failures.append(
            finding(
                LABEL_MANIFEST_PATH,
                "Root must contain exactly schema_version, prune, core, and overlays.",
            )
        )
    if not isinstance(document, dict):
        document = {}
    if document.get("schema_version") != 1:
        failures.append(finding(LABEL_MANIFEST_PATH, "schema_version must be 1."))
    if not isinstance(document.get("prune"), bool):
        failures.append(finding(LABEL_MANIFEST_PATH, "prune must be a boolean."))
    seen: dict[str, str] = {}
    core = document.get("core")
    _validate_label_array(core, "core", seen, failures)
    if isinstance(core, list) and not core:
        failures.append(finding(LABEL_MANIFEST_PATH, "core must not be empty."))
    overlays = document.get("overlays")
    if not _same_keys(overlays, {"profile", "domain"}):
        failures.append(
            finding(
                LABEL_MANIFEST_PATH,
                "overlays must contain exactly profile and domain.",
            )
        )
    else:
        assert isinstance(overlays, dict)
        profiles = overlays.get("profile")
        if not isinstance(profiles, dict) or set(profiles) != set(SUPPORTED_PROFILES):
            failures.append(
                finding(
                    LABEL_MANIFEST_PATH,
                    "overlays.profile must contain exactly the three supported profiles.",
                )
            )
        else:
            for profile in SUPPORTED_PROFILES:
                _validate_label_array(
                    profiles[profile],
                    f"overlays.profile.{profile}",
                    seen,
                    failures,
                )
        domains = overlays.get("domain")
        if not isinstance(domains, dict):
            failures.append(
                finding(LABEL_MANIFEST_PATH, "overlays.domain must be an object.")
            )
        else:
            for name in sorted(domains):
                if not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", name):
                    failures.append(
                        finding(
                            LABEL_MANIFEST_PATH,
                            f"Domain overlay name is not kebab-case: {name}",
                        )
                    )
                _validate_label_array(
                    domains[name],
                    f"overlays.domain.{name}",
                    seen,
                    failures,
                )
            for selected in config["label_domains"]:
                if selected not in domains:
                    failures.append(
                        finding(
                            CONFIG_PATH,
                            f"Selected label domain has no manifest overlay: {selected!r}.",
                        )
                    )
    count = len(seen)
    selected_profile = config["profile"] if config["mode"] == "generated" else None
    resolved_count = len(core) if isinstance(core, list) else 0
    if (
        selected_profile is not None
        and isinstance(overlays, dict)
        and isinstance(overlays.get("profile"), dict)
        and isinstance(overlays["profile"].get(selected_profile), list)
    ):
        resolved_count += len(overlays["profile"][selected_profile])
    if isinstance(overlays, dict) and isinstance(overlays.get("domain"), dict):
        for selected in config["label_domains"]:
            overlay = overlays["domain"].get(selected)
            if isinstance(overlay, list):
                resolved_count += len(overlay)
    result = rule_result(
        "label-manifest",
        "Canonical label manifest",
        failures,
        f"Validated {count} unique labels; resolved {resolved_count} for versioned selection",
    )
    result["evidence"] = {
        "profile": selected_profile,
        "domains": list(config["label_domains"]),
        "resolved_count": resolved_count,
    }
    return result


ISSUE_FORM_SPECS: dict[str, dict[str, Any]] = {
    "bug.yml": {
        "label": "bug",
        "title": "[Bug]: ",
        "required_ids": {
            "profile", "summary", "source", "reproduction", "expected",
            "actual", "environment", "evidence", "acknowledgements",
        },
    },
    "documentation.yml": {
        "label": "documentation",
        "title": "[Docs]: ",
        "required_ids": {
            "location", "problem_type", "problem", "correction", "source",
            "verification", "acknowledgements",
        },
    },
    "feature.yml": {
        "label": "enhancement",
        "title": "[Feature]: ",
        "required_ids": {
            "profile", "problem", "outcome", "acceptance", "non_goals",
            "compatibility", "alternatives", "validation", "acknowledgements",
        },
    },
}


def _yaml_header_scalar(text: str, key: str) -> str | None:
    match = re.search(rf"(?m)^{re.escape(key)}:\s*(.+?)\s*$", text)
    if not match:
        return None
    value = match.group(1).strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {'\"', "'"}:
        return value[1:-1]
    return value


def _yaml_flow_array(text: str, key: str) -> list[str] | None:
    raw = _yaml_header_scalar(text, key)
    if raw is None:
        return None
    try:
        value = json.loads(raw)
    except json.JSONDecodeError:
        return None
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        return None
    return value


def _issue_form_blocks(text: str) -> list[tuple[str, str | None]]:
    starts = list(re.finditer(r"(?m)^  - type:\s*([a-z]+)\s*$", text))
    blocks: list[tuple[str, str | None]] = []
    for index, match in enumerate(starts):
        end = starts[index + 1].start() if index + 1 < len(starts) else len(text)
        block = text[match.start():end]
        identifier = re.search(r"(?m)^    id:\s*([a-z][a-z0-9_]*)\s*$", block)
        blocks.append((match.group(1), identifier.group(1) if identifier else None))
    return blocks


def _issue_label_names(repo: Repository) -> set[object]:
    try:
        manifest = json.loads(repo.text(LABEL_MANIFEST_PATH))
    except (OSError, UnicodeError, json.JSONDecodeError):
        manifest = {}
    return {
        label.get("name")
        for label in manifest.get("core", [])
        if isinstance(label, dict) and isinstance(label.get("name"), str)
    }


def _validate_issue_form(
    repo: Repository,
    filename: str,
    specification: dict[str, Any],
    label_names: set[object],
    failures: list[dict[str, Any]],
) -> None:
    path = f"{ISSUE_TEMPLATE_DIRECTORY}/{filename}"
    try:
        text = repo.text(path)
    except (OSError, UnicodeError) as error:
        failures.append(finding(path, f"Cannot read canonical issue form: {error}"))
        return
    for key in ("name", "description"):
        value = _yaml_header_scalar(text, key)
        if value is None or not value.strip():
            failures.append(finding(path, f"Top-level {key} must be non-empty."))
    if _yaml_header_scalar(text, "title") != specification["title"]:
        failures.append(finding(path, f"Top-level title must be {specification['title']!r}."))
    labels = _yaml_flow_array(text, "labels")
    expected_label = specification["label"]
    if labels != [expected_label]:
        failures.append(finding(path, f"Top-level labels must be the JSON flow array [{expected_label!r}]."))
    elif expected_label not in label_names:
        failures.append(finding(path, f"Issue form label is absent from {LABEL_MANIFEST_PATH}: {expected_label}"))
    if _yaml_flow_array(text, "assignees") != []:
        failures.append(finding(path, "Top-level assignees must be an empty JSON flow array."))
    blocks = _issue_form_blocks(text)
    if not blocks or len(blocks) > 10:
        failures.append(finding(path, "Issue form body must contain between 1 and 10 elements."))
    invalid_types = sorted({kind for kind, _ in blocks} - {"checkboxes", "dropdown", "input", "markdown", "textarea"})
    if invalid_types:
        failures.append(finding(path, "Unsupported issue-form element types: " + ", ".join(invalid_types)))
    identifiers = [identifier for _, identifier in blocks if identifier is not None]
    if len(identifiers) != len(set(identifiers)):
        failures.append(finding(path, "Issue-form element IDs must be unique."))
    missing = sorted(set(specification["required_ids"]) - set(identifiers))
    if missing:
        failures.append(finding(path, "Required issue-form element IDs are missing: " + ", ".join(missing)))
    if "SECURITY.md" not in text or "private" not in text.casefold():
        failures.append(finding(path, "The opening guidance must route vulnerability details to SECURITY.md privately."))
    for identifier in set(specification["required_ids"]):
        if identifier == "acknowledgements":
            continue
        pattern = rf"(?ms)^    id:\s*{re.escape(identifier)}\s*$.*?(?=^  - type:|\Z)"
        match = re.search(pattern, text)
        if match and not re.search(r"(?m)^      required:\s*true\s*$", match.group(0)):
            failures.append(finding(path, f"Required evidence field {identifier!r} must be mandatory."))


def _validate_issue_intake_config(
    repo: Repository, repository: str, failures: list[dict[str, Any]]
) -> None:
    path = f"{ISSUE_TEMPLATE_DIRECTORY}/config.yml"
    try:
        text = repo.text(path)
    except (OSError, UnicodeError) as error:
        failures.append(finding(path, f"Cannot read issue-template configuration: {error}"))
        return
    if not re.search(r"(?m)^blank_issues_enabled:\s*false\s*$", text):
        failures.append(finding(path, "Blank issues must be disabled."))
    expected_url = f"https://github.com/{repository}/security/policy"
    url_match = re.search(r'(?m)^\s+url:\s*["\']?([^"\'\s]+)["\']?\s*$', text)
    actual_url = url_match.group(1) if url_match else None
    if actual_url != expected_url:
        failures.append(finding(path, f"Private-security contact URL must be {expected_url!r}."))
    if "private" not in text.casefold():
        failures.append(finding(path, "Security contact guidance must require private reporting."))


def check_issue_forms(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    label_names = _issue_label_names(repo)
    for filename, specification in ISSUE_FORM_SPECS.items():
        _validate_issue_form(repo, filename, specification, label_names, failures)
    _validate_issue_intake_config(repo, config["repository"], failures)
    return rule_result(
        "issue-forms",
        "Structured issue forms and private security routing",
        failures,
        "Validated three canonical forms, labels, evidence fields, and private security routing",
    )


WORKFLOW_PERMISSION_SCOPE = (
    r"(?:actions|attestations|checks|contents|deployments|discussions|id-token|issues|"
    r"models|packages|pages|pull-requests|repository-projects|security-events|statuses)"
)
WORKFLOW_PERMISSION_KEY = (
    rf"(?:{WORKFLOW_PERMISSION_SCOPE}|'(?:{WORKFLOW_PERMISSION_SCOPE})'|"
    rf'"(?:{WORKFLOW_PERMISSION_SCOPE})")'
)
WORKFLOW_WRITE_VALUE = r"(?:write|'write'|\"write\")"
WORKFLOW_WRITE_PERMISSION_RE = re.compile(
    rf"{WORKFLOW_PERMISSION_KEY}:\s*{WORKFLOW_WRITE_VALUE}(?:\s*#.*)?"
)
WORKFLOW_FLOW_WRITE_PERMISSION_RE = re.compile(
    rf"(?:^|,)\s*{WORKFLOW_PERMISSION_KEY}\s*:\s*"
    rf"{WORKFLOW_WRITE_VALUE}\s*(?=,|$)"
)


def _workflow_line_has_event(line: str, event: str) -> bool:
    escaped = re.escape(event)
    key = rf"(?:{escaped}|'{escaped}'|\"{escaped}\")"
    if re.fullmatch(rf"  {key}:\s*(?:#.*)?", line):
        return True
    stripped = _strip_yaml_comment(line).strip()
    if re.fullmatch(rf"on:\s*{key}\s*", stripped):
        return True
    sequence = re.fullmatch(r"on:\s*\[(.*)\]\s*", stripped)
    if sequence is not None:
        return re.search(
            rf"(?:^|,)\s*{key}\s*(?=,|$)", sequence.group(1)
        ) is not None
    mapping = re.fullmatch(r"on:\s*\{(.*)\}\s*", stripped)
    if mapping is not None:
        return re.search(rf"(?:^|,)\s*{key}\s*:", mapping.group(1)) is not None
    return False


def _workflow_has_event(lines: list[str], event: str) -> bool:
    return any(_workflow_line_has_event(line, event) for line in lines)


def _flow_permissions_request_write(stripped: str) -> bool:
    normalized = _strip_yaml_comment(stripped).strip()
    mapping = re.fullmatch(r"permissions:\s*\{(.*)\}\s*", normalized)
    return (
        mapping is not None
        and WORKFLOW_FLOW_WRITE_PERMISSION_RE.search(mapping.group(1)) is not None
    )


def _check_pull_request_target(
    path: str,
    lines: list[str],
    failures: list[dict[str, Any]],
) -> None:
    for number, line in enumerate(lines, start=1):
        if _workflow_line_has_event(line, "pull_request_target"):
            failures.append(
                finding(
                    path,
                    "pull_request_target is prohibited; untrusted pull requests must not execute privileged repository code.",
                    number,
                )
            )


def _check_workflow_level_pr_permissions(
    path: str,
    lines: list[str],
    failures: list[dict[str, Any]],
) -> None:
    in_permissions = False
    for number, line in enumerate(lines, start=1):
        indent = len(line) - len(line.lstrip(" "))
        stripped = line.strip()
        if indent == 0 and re.fullmatch(
            r"permissions:\s*write-all(?:\s*#.*)?", stripped
        ):
            failures.append(
                finding(
                    path,
                    "Workflow triggered by pull_request must not request write-capable token permissions.",
                    number,
                )
            )
        if indent == 0 and _flow_permissions_request_write(stripped):
            failures.append(
                finding(
                    path,
                    "Workflow triggered by pull_request must not request workflow-level write permissions.",
                    number,
                )
            )
        if indent == 0 and re.fullmatch(r"permissions:\s*(?:#.*)?", stripped):
            in_permissions = True
            continue
        if in_permissions and stripped and not stripped.startswith("#") and indent == 0:
            in_permissions = False
        if in_permissions and WORKFLOW_WRITE_PERMISSION_RE.fullmatch(stripped):
            failures.append(
                finding(
                    path,
                    "Workflow triggered by pull_request must not request workflow-level write permissions.",
                    number,
                )
            )


def _workflow_job_blocks(lines: list[str]) -> list[tuple[int, list[str]]]:
    jobs_index = next(
        (
            index
            for index, line in enumerate(lines)
            if re.fullmatch(r"jobs:\s*(?:#.*)?", line.strip())
            and not line.startswith(" ")
        ),
        None,
    )
    if jobs_index is None:
        return []
    starts = [
        index
        for index in range(jobs_index + 1, len(lines))
        if re.match(r"^  [A-Za-z0-9_-]+:\s*(?:#.*)?$", lines[index])
    ]
    return [
        (
            start,
            lines[
                start : starts[position + 1]
                if position + 1 < len(starts)
                else len(lines)
            ],
        )
        for position, start in enumerate(starts)
    ]


def _job_excludes_pull_request(block: list[str]) -> bool:
    exclusion = re.compile(
        r"^    if:\s*github\.event_name\s*!=\s*(['\"])pull_request\1\s*(?:#.*)?$"
    )
    return any(exclusion.match(line) for line in block)


def _check_job_pr_permissions(
    path: str,
    lines: list[str],
    failures: list[dict[str, Any]],
) -> None:
    for start, block in _workflow_job_blocks(lines):
        if _job_excludes_pull_request(block):
            continue
        in_permissions = False
        for offset, item in enumerate(block):
            number = start + offset + 1
            indent = len(item) - len(item.lstrip(" "))
            stripped = item.strip()
            if indent == 4 and re.fullmatch(
                r"permissions:\s*write-all(?:\s*#.*)?", stripped
            ):
                failures.append(
                    finding(
                        path,
                        "Job reachable from pull_request must not request write-all permissions.",
                        number,
                    )
                )
            if indent == 4 and _flow_permissions_request_write(stripped):
                failures.append(
                    finding(
                        path,
                        "Job reachable from pull_request must not request write-capable token permissions.",
                        number,
                    )
                )
            if indent == 4 and re.fullmatch(r"permissions:\s*(?:#.*)?", stripped):
                in_permissions = True
                continue
            if in_permissions and stripped and not stripped.startswith("#") and indent <= 4:
                in_permissions = False
            if in_permissions and WORKFLOW_WRITE_PERMISSION_RE.fullmatch(stripped):
                failures.append(
                    finding(
                        path,
                        "Job reachable from pull_request must not request write-capable token permissions.",
                        number,
                    )
                )


def _check_pr_workflow_permissions(
    path: str,
    lines: list[str],
    failures: list[dict[str, Any]],
) -> None:
    _check_pull_request_target(path, lines, failures)
    if not _workflow_has_event(lines, "pull_request"):
        return
    _check_workflow_level_pr_permissions(path, lines, failures)
    _check_job_pr_permissions(path, lines, failures)


def _check_external_action_references(
    path: str,
    lines: list[str],
    failures: list[dict[str, Any]],
) -> int:
    uses_line = re.compile(
        r"^\s*(?:-\s*)?uses:\s*([^\s#]+)(?:\s+#\s*(.+?))?\s*$"
    )
    full_sha = re.compile(r"^[0-9a-f]{40}$")
    version_comment = re.compile(r"^v\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$")
    checked = 0
    for number, line in enumerate(lines, start=1):
        if "uses:" not in line:
            continue
        match = uses_line.match(line)
        if not match:
            failures.append(finding(path, "Action reference cannot be parsed.", number))
            continue
        reference, comment = match.groups()
        if reference.startswith("./"):
            continue
        checked += 1
        if "@" not in reference:
            failures.append(
                finding(path, "External action must include a revision.", number)
            )
            continue
        action, revision = reference.rsplit("@", 1)
        if action.startswith("docker://") or not full_sha.fullmatch(revision):
            failures.append(
                finding(
                    path,
                    "External action must be pinned to a full lowercase 40-character commit SHA.",
                    number,
                )
            )
        if not comment or not version_comment.fullmatch(comment.strip()):
            failures.append(
                finding(
                    path,
                    "Pinned action must include an audited semantic-version comment.",
                    number,
                )
            )
    return checked


def check_workflow_actions(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    del config
    failures: list[dict[str, Any]] = []
    checked = 0
    for path in repo.files:
        pure = PurePosixPath(path)
        if (
            not path.startswith(".github/workflows/")
            or pure.suffix.casefold() not in {".yml", ".yaml"}
        ):
            continue
        try:
            lines = repo.text(path).splitlines()
        except (OSError, UnicodeError):
            continue
        _check_pr_workflow_permissions(path, lines, failures)
        checked += _check_external_action_references(path, lines, failures)
    return rule_result(
        "workflow-actions",
        "Immutable workflow actions",
        failures,
        f"Validated {checked} external action references",
    )


def check_version_changelog(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []
    try:
        version = repo.text("VERSION").strip()
    except (OSError, UnicodeError) as error:
        version = ""
        failures.append(finding("VERSION", f"Cannot read version: {error}"))
    semver = re.compile(
        r"^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)"
        r"(?:-[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?"
        r"(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$"
    )
    if version and not semver.fullmatch(version):
        failures.append(
            finding("VERSION", "VERSION must contain one SemVer value without a v prefix.")
        )
    try:
        changelog = repo.text("CHANGELOG.md")
    except (OSError, UnicodeError) as error:
        changelog = ""
        failures.append(finding("CHANGELOG.md", f"Cannot read changelog: {error}"))
    if changelog and not re.search(
        r"^## \[Unreleased\]\s*$", changelog, re.MULTILINE
    ):
        failures.append(
            finding("CHANGELOG.md", "Changelog must contain an Unreleased level-two heading.")
        )
    if (
        config["mode"] in {"generated", "template"}
        and version
        and version != "0.0.0"
        and changelog
        and not re.search(
            rf"^## \[{re.escape(version)}\] - \d{{4}}-\d{{2}}-\d{{2}}\s*$",
            changelog,
            re.MULTILINE,
        )
    ):
        failures.append(
            finding(
                "CHANGELOG.md",
                f"Changelog has no dated release heading for VERSION {version}.",
            )
        )
    return rule_result(
        "version-changelog",
        "Version and changelog structure",
        failures,
        f"VERSION {version or 'unavailable'} and changelog structure are consistent",
    )


def check_git_diff(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    del config
    completed = repo._git("diff", "--check", "HEAD", "--", check=False)
    failures: list[dict[str, Any]] = []
    if completed.returncode != 0:
        detail = (completed.stdout + completed.stderr).strip()
        failures.append(finding(".", detail or "git diff --check failed"))
    return rule_result(
        "git-diff-check",
        "Working-tree whitespace check",
        failures,
        "git diff --check passes",
    )


def _vba_paths(repo: Repository) -> list[str]:
    return [
        path
        for path in repo.files
        if PurePosixPath(path).suffix.casefold() in VBA_SUFFIXES
    ]


def check_vba_option_explicit(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    del config
    failures: list[dict[str, Any]] = []
    declaration = re.compile(
        r"^\s*Option\s+Explicit\b", re.IGNORECASE | re.MULTILINE
    )
    paths = _vba_paths(repo)
    for path in paths:
        try:
            text = repo.text(path)
        except (OSError, UnicodeError):
            continue
        if not declaration.search(text):
            failures.append(finding(path, "VBA source must declare Option Explicit."))
    return rule_result(
        "vba-option-explicit",
        "VBA explicit declarations",
        failures,
        f"All {len(paths)} VBA components declare Option Explicit",
    )


def check_vba_export_header(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    del config
    failures: list[dict[str, Any]] = []
    declaration = re.compile(r'^Attribute VB_Name = "([^"]*)"$')
    identifier = re.compile(r"^[A-Za-z][A-Za-z0-9_]*$")
    declared: dict[str, str] = {}
    paths = _vba_paths(repo)
    for path in paths:
        pure = PurePosixPath(path)
        try:
            lines = repo.text(path).splitlines()
        except (OSError, UnicodeError) as error:
            failures.append(finding(path, f"VBA source cannot be read: {error}"))
            continue
        matches = [
            (number, match.group(1))
            for number, line in enumerate(lines, start=1)
            if (match := declaration.match(line))
        ]
        if not matches:
            failures.append(
                finding(
                    path,
                    "VBA export must declare Attribute VB_Name in canonical VBE form.",
                )
            )
            continue
        if len(matches) > 1:
            failures.append(
                finding(
                    path,
                    "VBA export declares Attribute VB_Name more than once.",
                    matches[1][0],
                )
            )
        number, name = matches[0]
        if pure.suffix.casefold() == ".bas" and number != 1:
            failures.append(
                finding(path, "Attribute VB_Name must be line 1 of a module export.", number)
            )
        elif pure.suffix.casefold() in {".cls", ".frm"} and number > VBA_HEADER_SCAN_LINES:
            failures.append(
                finding(
                    path,
                    "Attribute VB_Name must appear in the leading export header.",
                    number,
                )
            )
        if name != pure.stem:
            failures.append(
                finding(
                    path,
                    f"Declared component name {name!r} does not match file name {pure.stem!r}.",
                    number,
                )
            )
        if not identifier.fullmatch(name):
            failures.append(
                finding(path, f"Component name {name!r} is not a VBA identifier.", number)
            )
        elif len(name) > VBA_COMPONENT_NAME_LIMIT:
            failures.append(
                finding(
                    path,
                    f"Component name exceeds the {VBA_COMPONENT_NAME_LIMIT}-character limit.",
                    number,
                )
            )
        previous = declared.get(name.casefold())
        if previous is not None:
            failures.append(
                finding(
                    path,
                    f"Component name {name!r} collides with {previous} case-insensitively.",
                    number,
                )
            )
        else:
            declared[name.casefold()] = path
    return rule_result(
        "vba-export-header",
        "VBA export headers and component names",
        failures,
        f"All {len(paths)} VBA components have unique matching export names",
    )


def _strip_vba_line(line: str) -> str:
    result: list[str] = []
    index = 0
    quoted = False
    while index < len(line):
        character = line[index]
        if character == '"':
            if quoted and index + 1 < len(line) and line[index + 1] == '"':
                index += 2
                continue
            quoted = not quoted
            result.append(" ")
        elif character == "'" and not quoted:
            break
        elif quoted:
            result.append(" ")
        else:
            result.append(character)
        index += 1
    return "".join(result)


def _handle_vba_directive(
    path: str,
    number: int,
    upper: str,
    directives: list[dict[str, bool]],
    failures: list[dict[str, Any]],
) -> bool:
    if upper.startswith("#IF "):
        condition = upper[4:]
        requires = "VBA7" in condition and "NOT VBA7" not in condition
        directives.append({"vba7": requires, "active": requires})
        return True
    if upper.startswith("#ELSEIF "):
        if not directives:
            failures.append(finding(path, "#ElseIf without #If.", number))
        else:
            condition = upper[8:]
            directives[-1]["active"] = "VBA7" in condition and "NOT VBA7" not in condition
        return True
    if upper.startswith("#ELSE"):
        if not directives:
            failures.append(finding(path, "#Else without #If.", number))
        else:
            directives[-1]["active"] = not directives[-1]["vba7"]
        return True
    if upper.startswith("#END IF"):
        if not directives:
            failures.append(finding(path, "#End If without #If.", number))
        else:
            directives.pop()
        return True
    return False


def _scan_vba_structure_component(
    path: str,
    lines: list[str],
    failures: list[dict[str, Any]],
) -> tuple[list[dict[str, bool]], list[tuple[str, str, int]], set[str], list[tuple[int, str]]]:
    opener = re.compile(
        r"^\s*(?:Public|Private|Friend)?\s*(?:Static\s+)?"
        r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\b",
        re.IGNORECASE,
    )
    closer = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
    label_re = re.compile(r"^\s*([A-Za-z_]\w*|\d+):\s*$")
    declare_re = re.compile(r"^\s*(?:Public|Private)?\s*Declare\s+(?:Function|Sub)\b", re.IGNORECASE)
    directives: list[dict[str, bool]] = []
    procedures: list[tuple[str, str, int]] = []
    labels: set[str] = set()
    executable: list[tuple[int, str]] = []
    for number, raw in enumerate(lines, start=1):
        upper = raw.strip().upper()
        if _handle_vba_directive(path, number, upper, directives, failures):
            continue
        code = _strip_vba_line(raw)
        if not code.strip():
            continue
        executable.append((number, code))
        match = label_re.match(code)
        if match:
            labels.add(match.group(1).casefold())
        if declare_re.match(code) and any(item["active"] for item in directives) and not re.search(r"\bPtrSafe\b", code, re.IGNORECASE):
            failures.append(finding(path, "Declare in an active VBA7 branch must include PtrSafe.", number))
        match = opener.match(code)
        if match and " declare " not in f" {code.casefold()} ":
            if procedures:
                kind, name, start = procedures[-1]
                failures.append(finding(path, f"{kind} {name} opened at line {start} has no closing statement.", number))
                procedures.clear()
            procedures.append((match.group(1), match.group(2), number))
            continue
        match = closer.match(code)
        if match:
            if not procedures:
                failures.append(finding(path, f"{match.group(0).strip()} has no opener.", number))
            else:
                procedures.pop()
    return directives, procedures, labels, executable


def _validate_vba_structure_tail(
    path: str,
    directives: list[dict[str, bool]],
    procedures: list[tuple[str, str, int]],
    labels: set[str],
    executable: list[tuple[int, str]],
    failures: list[dict[str, Any]],
) -> None:
    if directives:
        failures.append(finding(path, f"{len(directives)} conditional-compilation block(s) are unclosed."))
    for kind, name, start in procedures:
        failures.append(finding(path, f"{kind} {name} opened at line {start} is unclosed."))
    jump_re = re.compile(r"\b(?:GoTo|Resume)\s+([A-Za-z_]\w*|\d+|-1)\b", re.IGNORECASE)
    for number, code in executable:
        for match in jump_re.finditer(code):
            target = match.group(1)
            if target.casefold() in {"next", "0", "-1"}:
                continue
            if target.casefold() not in labels:
                failures.append(finding(path, f"Jump target is not defined: {target}", number))


def check_vba_structure(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    del config
    failures: list[dict[str, Any]] = []
    paths = _vba_paths(repo)
    for path in paths:
        try:
            lines = repo.text(path).splitlines()
        except (OSError, UnicodeError):
            continue
        state = _scan_vba_structure_component(path, lines, failures)
        _validate_vba_structure_tail(path, *state, failures)
    return rule_result(
        "vba-structure",
        "VBA structural safety",
        failures,
        f"Validated procedure, directive, jump, and PtrSafe structure in {len(paths)} components",
    )


def check_vba_visibility(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    settings = config["vba"]
    components = settings["components"]
    source_roots = settings["source_roots"]
    test_roots = settings["test_roots"]
    failures: list[dict[str, Any]] = []
    tracked_vba = set(_vba_paths(repo))
    governed = {
        path
        for path in tracked_vba
        if is_under(path, source_roots) or is_under(path, test_roots)
    }
    for path in sorted(governed - set(components)):
        failures.append(
            finding(path, "Tracked VBA component is not assigned a profile role.")
        )
    for path in sorted(set(components) - tracked_vba):
        failures.append(
            finding(path, "Configured VBA component is not tracked.")
        )
    private_re = re.compile(
        r"^\s*Option\s+Private\s+Module\b", re.IGNORECASE | re.MULTILINE
    )
    for path, role in components.items():
        if path not in tracked_vba:
            continue
        text = repo.text(path)
        private = bool(private_re.search(text))
        suffix = PurePosixPath(path).suffix.casefold()
        if role == "internal" and suffix == ".bas" and not private:
            failures.append(
                finding(path, "Internal standard module must declare Option Private Module.")
            )
        if role == "public" and suffix == ".bas" and private:
            failures.append(
                finding(path, "Public facade must not declare Option Private Module.")
            )
        if is_under(path, test_roots) and role != "test":
            failures.append(
                finding(path, "Component under a test root must use the test role.")
            )
        if is_under(path, source_roots) and role == "test":
            failures.append(
                finding(path, "A test-role component must not be stored under a source root.")
            )
    return rule_result(
        "vba-visibility",
        "VBA component roles and visibility",
        failures,
        "Validated "
        f"{len(components)} configured component roles "
        f"(public={sum(role == 'public' for role in components.values())}, "
        f"internal={sum(role == 'internal' for role in components.values())}, "
        f"test={sum(role == 'test' for role in components.values())}, "
        f"example={sum(role == 'example' for role in components.values())}, "
        f"ui={sum(role == 'ui' for role in components.values())})",
    )


def check_generated_vba_contract(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    """Require substantive facade, core, and test assets for each profile."""

    profiles = config["profiles"]
    components = config["vba"]["components"]
    tracked_vba = set(_vba_paths(repo))
    selected_profiles = (
        SUPPORTED_PROFILES
        if config["mode"] == "template"
        else (config["profile"],)
    )
    failures: list[dict[str, Any]] = []
    profile_evidence: dict[str, Any] = {}

    for profile in selected_profiles:
        contract = profiles[profile]["vba_contract"]
        minimum_roles = contract["minimum_roles"]
        required_components = contract["required_components"]
        observed_roles = {
            role: sum(
                configured_role == role and path in tracked_vba
                for path, configured_role in components.items()
            )
            for role in minimum_roles
        }
        profile_evidence[profile] = {
            "minimum_roles": dict(minimum_roles),
            "observed_roles": observed_roles,
            "required_components": dict(required_components),
        }

        for role, minimum in minimum_roles.items():
            observed = observed_roles[role]
            if observed < minimum:
                failures.append(
                    finding(
                        CONFIG_PATH,
                        f"Generated VBA contract for profile {profile!r} requires role "
                        f"{role!r} >= {minimum}; observed {observed}.",
                    )
                )

        for path, expected_role in required_components.items():
            actual_role = components.get(path)
            if actual_role is None:
                failures.append(
                    finding(
                        path,
                        f"Profile {profile!r} requires this {expected_role!r} "
                        "starter component to be registered.",
                    )
                )
            elif actual_role != expected_role:
                failures.append(
                    finding(
                        path,
                        f"Profile {profile!r} requires role {expected_role!r}; "
                        f"configured role is {actual_role!r}.",
                    )
                )
            elif path not in tracked_vba:
                failures.append(
                    finding(
                        path,
                        f"Profile {profile!r} requires this registered "
                        f"{expected_role!r} starter component to be tracked.",
                    )
                )

    result = rule_result(
        "generated-vba-contract",
        "Generated-profile VBA contract",
        failures,
        "Resolved substantive VBA contracts for profiles: "
        + ", ".join(selected_profiles),
    )
    result["evidence"] = {
        "mode": config["mode"],
        "selected_profile": config["profile"],
        "profiles": profile_evidence,
    }
    return result


def _public_surface(
    repo: Repository, components: dict[str, str]
) -> tuple[list[str], list[dict[str, Any]]]:
    declaration = re.compile(
        r"^\s*Public\s+(?:Static\s+)?"
        r"(Sub|Function|Property\s+(?:Get|Let|Set)|Enum|Type|Const)"
        r"\s+([A-Za-z_]\w*)\b",
        re.IGNORECASE,
    )
    surface: list[str] = []
    failures: list[dict[str, Any]] = []
    global_names: dict[str, str] = {}
    for path, role in components.items():
        if role != "public" or not repo.path(path).is_file():
            continue
        component = PurePosixPath(path).stem
        for number, raw in enumerate(repo.text(path).splitlines(), start=1):
            code = _strip_vba_line(raw)
            match = declaration.match(code)
            if not match:
                continue
            kind = " ".join(part.capitalize() for part in match.group(1).split())
            name = match.group(2)
            surface.append(f"{component}\t{kind}\t{name}")
            if PurePosixPath(path).suffix.casefold() == ".bas":
                key = name.casefold()
                previous = global_names.get(key)
                if previous is not None:
                    failures.append(
                        finding(
                            path,
                            f"Public standard-module member {name!r} collides with {previous}.",
                            number,
                        )
                    )
                else:
                    global_names[key] = path
    return sorted(set(surface), key=lambda item: (item.casefold(), item)), failures


def check_vba_public_api(
    repo: Repository, config: dict[str, Any]
) -> dict[str, Any]:
    settings = config["vba"]
    manifest = settings["public_api_manifest"]
    actual, failures = _public_surface(repo, settings["components"])
    if manifest is None:
        return rule_result(
            "vba-public-api",
            "VBA public API manifest",
            failures,
            f"Observed {len(actual)} public declarations; no manifest configured",
        )
    if manifest not in repo.files or not repo.path(manifest).is_file():
        failures.append(finding(manifest, "Configured public API manifest is not tracked."))
        recorded: list[str] = []
    else:
        try:
            recorded = sorted(
                {
                    line.strip()
                    for line in repo.text(manifest).splitlines()
                    if line.strip() and not line.lstrip().startswith("#")
                },
                key=lambda item: (item.casefold(), item),
            )
        except (OSError, UnicodeError) as error:
            failures.append(finding(manifest, f"Cannot read public API manifest: {error}"))
            recorded = []
    for removed in sorted(set(recorded) - set(actual), key=str.casefold):
        failures.append(
            finding(manifest, f"Recorded public declaration is missing: {removed}")
        )
    for added in sorted(set(actual) - set(recorded), key=str.casefold):
        failures.append(
            finding(manifest, f"Public declaration is not recorded: {added}")
        )
    return rule_result(
        "vba-public-api",
        "VBA public API manifest",
        failures,
        f"Public API manifest matches {len(actual)} declarations",
    )


Check = Callable[[Repository, dict[str, Any]], dict[str, Any]]
CHECKS: tuple[Check, ...] = (
    check_required_paths,
    check_placeholders,
    check_identity,
    check_dotfile_policy,
    check_structured_data,
    check_markdown_links,
    check_text_integrity,
    check_forbidden_artifacts,
    check_line_endings,
    check_label_manifest,
    check_issue_forms,
    check_workflow_actions,
    check_version_changelog,
    check_git_diff,
    check_vba_option_explicit,
    check_vba_export_header,
    check_vba_structure,
    check_vba_visibility,
    check_generated_vba_contract,
    check_vba_public_api,
)


def build_report(root: Path) -> dict[str, Any]:
    """Run the complete canonical rule set and return a deterministic report."""

    repo = Repository(root)
    config, configuration = load_configuration(repo)
    results = [configuration]
    if config is not None:
        for check in CHECKS:
            try:
                results.append(check(repo, config))
            except Exception as error:
                results.append(
                    rule_result(
                        check.__name__.removeprefix("check_").replace("_", "-"),
                        check.__name__.removeprefix("check_").replace("_", " ").title(),
                        [finding(".", f"Rule could not complete: {error}")],
                        "",
                    )
                )

    failed_rules = sum(result["status"] == "fail" for result in results)
    finding_count = sum(len(result["findings"]) for result in results)
    passed_rules = len(results) - failed_rules
    mode = config.get("mode") if config is not None else None
    profile = config.get("profile") if config is not None else None
    repository_name = config.get("repository") if config is not None else None
    return {
        "schema_version": SCHEMA_VERSION,
        "tool": TOOL_NAME,
        "repository": repository_name,
        "commit": repo.commit(),
        "mode": mode,
        "profile": profile,
        "scope_note": (
            "Repository evidence only; this gate does not execute Excel, compile VBA, "
            "prove numerical accuracy, exercise UI state, or certify release binaries."
        ),
        "status": "pass" if failed_rules == 0 else "fail",
        "counts": {
            "rules": len(results),
            "passed": passed_rules,
            "failed": failed_rules,
            "findings": finding_count,
        },
        "rules": results,
    }


def _markdown_escape(value: object) -> str:
    return str(value).replace("|", r"\|").replace("\n", " ")


def markdown_report(report: dict[str, Any]) -> str:
    status = str(report["status"]).upper()
    counts = report["counts"]
    lines = [
        "# Repository quality",
        "",
        f"**Status:** {status}",
        "",
        (
            f"Rules: {counts['passed']} passed, {counts['failed']} failed; "
            f"{counts['findings']} findings."
        ),
        "",
        str(report["scope_note"]),
        "",
        "| Rule | Status | Summary |",
        "| --- | --- | --- |",
    ]
    for result in report["rules"]:
        lines.append(
            "| "
            + " | ".join(
                (
                    _markdown_escape(result["id"]),
                    _markdown_escape(str(result["status"]).upper()),
                    _markdown_escape(result["summary"]),
                )
            )
            + " |"
        )
    failures = [
        result for result in report["rules"] if result["status"] == "fail"
    ]
    if failures:
        lines.extend(("", "## Findings", ""))
        for result in failures:
            lines.append(f"### {_markdown_escape(result['title'])}")
            lines.append("")
            for item in result["findings"]:
                location = str(item["path"])
                if "line" in item:
                    location += f":{item['line']}"
                quote = chr(96)
                lines.append(
                    f"- {quote}{_markdown_escape(location)}{quote} — "
                    f"{_markdown_escape(item['message'])}"
                )
            lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def console_report(report: dict[str, Any]) -> str:
    lines = []
    for result in report["rules"]:
        marker = "PASS" if result["status"] == "pass" else "FAIL"
        lines.append(f"[{marker}] {result['id']}: {result['summary']}")
        for item in result["findings"]:
            location = str(item["path"])
            if "line" in item:
                location += f":{item['line']}"
            lines.append(f"       {location}: {item['message']}")
    counts = report["counts"]
    lines.append(
        f"{str(report['status']).upper()}: {counts['passed']}/{counts['rules']} "
        f"rules passed; {counts['findings']} findings."
    )
    lines.append(f"Scope: {report['scope_note']}")
    return "\n".join(lines)


def _write_report(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", newline="\n")


def _write_fixture(path: Path, content: str | bytes, *, crlf: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if isinstance(content, bytes):
        path.write_bytes(content)
        return
    normalized = content.replace("\r\n", "\n")
    if crlf:
        path.write_bytes(normalized.replace("\n", "\r\n").encode("cp1252"))
    else:
        path.write_text(normalized, encoding="utf-8", newline="\n")


def _fixture_configuration() -> dict[str, Any]:
    forbidden_identity = "DONOR" + "-PROJECT"
    template_identity = "TEMPLATE" + "-IDENTITY"
    return {
        "schema_version": 1,
        "mode": "generated",
        "profile": "library",
        "repository": "example/fixture",
        "template_contract": {"version": "1.2.0", "source": "example/template"},
        "label_domains": [],
        "required_paths": [
            ".editorconfig",
            ".gitattributes",
            ".github/labels.json",
            ".github/workflows/static-checks.yml",
            ".gitignore",
            "CHANGELOG.md",
            "README.md",
            "VERSION",
        ],
        "required_directories": ["docs", "src", "tests"],
        "profiles": {
            "application": {
                "required_paths": [],
                "required_directories": [],
                "vba_contract": {
                    "minimum_roles": {"internal": 1, "public": 1, "test": 1},
                    "required_components": {
                        "src/core/QualityCore.bas": "internal",
                        "src/modules/Quality.bas": "public",
                        "tests/modules/QualityTests.bas": "test",
                    },
                },
            },
            "library": {
                "required_paths": [],
                "required_directories": ["src/core", "src/modules", "tests/modules"],
                "vba_contract": {
                    "minimum_roles": {"internal": 1, "public": 1, "test": 1},
                    "required_components": {
                        "src/core/QualityCore.bas": "internal",
                        "src/modules/Quality.bas": "public",
                        "tests/modules/QualityTests.bas": "test",
                    },
                },
            },
            "ui-component": {
                "required_paths": [],
                "required_directories": [],
                "vba_contract": {
                    "minimum_roles": {"internal": 1, "public": 1, "test": 1},
                    "required_components": {
                        "src/core/QualityCore.bas": "internal",
                        "src/modules/Quality.bas": "public",
                        "tests/modules/QualityTests.bas": "test",
                    },
                },
            },
        },
        "allowed_office_binary_globs": [],
        "placeholders": {
            "pattern": r"\{\{([A-Z][A-Z0-9_]*)\}\}",
            "catalogue": {
                "OPTIONAL_NOTE": {
                    "category": "optional",
                    "description": "Optional fixture value.",
                },
                "PROFILE_NOTE": {
                    "category": "profile-specific",
                    "description": "Profile-specific fixture value.",
                    "values": {
                        "application": "Application fixture",
                        "library": "Library fixture",
                        "ui-component": "UI fixture",
                    },
                },
                "REPEATABLE_NOTE": {
                    "category": "repeatable",
                    "description": "Repeatable fixture value.",
                    "item_format": "- {value}",
                },
                "REQUIRED_NOTE": {
                    "category": "required",
                    "description": "Required fixture value.",
                },
            },
            "block_markers": {
                "template_only": "template:remove",
                "profile": "template:profile:{profile}",
                "optional": "template:optional:{token}",
                "repeatable": "template:repeatable:{token}",
            },
            "template_only_paths": [],
            "exclude_paths": [CONFIG_PATH],
        },
        "identity": {
            "forbidden_tokens": [forbidden_identity],
            "template_tokens": [template_identity],
            "exclude_paths": [CONFIG_PATH],
        },
        "vba": {
            "source_roots": ["src"],
            "test_roots": ["tests"],
            "components": {
                "src/core/QualityCore.bas": "internal",
                "src/modules/Quality.bas": "public",
                "tests/modules/QualityTests.bas": "test",
            },
            "public_api_manifest": "docs/PUBLIC_API.txt",
        },
    }


def _fixture_labels() -> dict[str, Any]:
    return {
        "schema_version": 1,
        "prune": False,
        "core": [
            {
                "name": "bug",
                "color": "D73A4A",
                "description": "Something is not working",
            },
            {
                "name": "documentation",
                "color": "0075CA",
                "description": "Documentation correction",
            },
            {
                "name": "enhancement",
                "color": "A2EEEF",
                "description": "New capability",
            }
        ],
        "overlays": {
            "profile": {
                "application": [],
                "library": [],
                "ui-component": [],
            },
            "domain": {},
        },
    }


def _run_git(root: Path, *arguments: str) -> None:
    subprocess.run(
        ["git", "-C", str(root), *arguments],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )


def _initialize_fixture(root: Path) -> None:
    checkout_revision = "a" * 40
    _write_fixture(
        root / ".editorconfig",
        """root = true

[*]
charset = utf-8
end_of_line = lf
insert_final_newline = true

[*.{bas,cls,frm}]
charset = latin1
end_of_line = crlf
insert_final_newline = true
""",
    )
    _write_fixture(
        root / ".gitattributes",
        """* text=auto
*.bas text eol=crlf
*.cls text eol=crlf
*.frm text eol=crlf
*.json text eol=lf
*.md text eol=lf
*.py text eol=lf
*.yml text eol=lf
*.yaml text eol=lf
*.xlsm binary
""",
    )
    _write_fixture(
        root / ".gitignore",
        """.env
!.env.example
__pycache__/
*.pem
*.xlsm
test-results/
~$*
""",
    )
    _write_fixture(
        root / CONFIG_PATH,
        json.dumps(_fixture_configuration(), indent=2, ensure_ascii=False) + "\n",
    )
    _write_fixture(
        root / LABEL_MANIFEST_PATH,
        json.dumps(_fixture_labels(), indent=2, ensure_ascii=False) + "\n",
    )
    fixture_forms = {
        "bug.yml": ("Bug report", "[Bug]: ", "bug", [
            "profile", "summary", "source", "reproduction", "expected",
            "actual", "environment", "evidence", "acknowledgements",
        ]),
        "documentation.yml": ("Documentation request", "[Docs]: ", "documentation", [
            "location", "problem_type", "problem", "correction", "source",
            "verification", "acknowledgements",
        ]),
        "feature.yml": ("Feature request", "[Feature]: ", "enhancement", [
            "profile", "problem", "outcome", "acceptance", "non_goals",
            "compatibility", "alternatives", "validation", "acknowledgements",
        ]),
    }
    for filename, (name, title, label, identifiers) in fixture_forms.items():
        body = [
            f'name: "{name}"',
            'description: "Canonical fixture form."',
            f'title: "{title}"',
            f'labels: ["{label}"]',
            "assignees: []",
            "body:",
            "  - type: markdown",
            "    attributes:",
            "      value: Follow `SECURITY.md` and report vulnerabilities privately.",
        ]
        for identifier in identifiers:
            kind = "checkboxes" if identifier == "acknowledgements" else "textarea"
            body.extend([
                f"  - type: {kind}",
                f"    id: {identifier}",
                "    attributes:",
                f"      label: {identifier.replace('_', ' ').title()}",
            ])
            if kind == "checkboxes":
                body.extend([
                    "      options:",
                    "        - label: I confirm this fixture statement.",
                    "          required: true",
                ])
            else:
                body.extend([
                    "    validations:",
                    "      required: true",
                ])
        _write_fixture(root / ISSUE_TEMPLATE_DIRECTORY / filename, "\n".join(body) + "\n")
    _write_fixture(
        root / ISSUE_TEMPLATE_DIRECTORY / "config.yml",
        """blank_issues_enabled: false
contact_links:
  - name: Report a security vulnerability privately
    url: "https://github.com/example/fixture/security/policy"
    about: Use the private reporting channel.
""",
    )
    _write_fixture(
        root / ".github/workflows/static-checks.yml",
        f"""name: Static repository checks

on:
  pull_request:

permissions:
  contents: read

jobs:
  quality:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@{checkout_revision} # v4.2.2
      - name: Run gate
        run: python3 tools/check_repo.py --root .
""",
    )
    _write_fixture(
        root / "README.md",
        "# Fixture\n\nSee [details](docs/DETAILS.md#details).\n",
    )
    _write_fixture(
        root / "docs/DETAILS.md",
        "# Details\n\nFixture details.\n",
    )
    _write_fixture(
        root / "docs/PUBLIC_API.txt",
        "Quality\tFunction\tEcho\n",
    )
    _write_fixture(
        root / "CHANGELOG.md",
        "# Changelog\n\n## [Unreleased]\n\n- Fixture baseline.\n",
    )
    _write_fixture(root / "VERSION", "0.0.0\n")
    _write_fixture(
        root / "src/core/QualityCore.bas",
        """Attribute VB_Name = "QualityCore"
Option Explicit
Option Private Module

Public Function Normalize(ByVal value As String) As String
    Normalize = value
End Function
""",
        crlf=True,
    )
    _write_fixture(
        root / "src/modules/Quality.bas",
        """Attribute VB_Name = "Quality"
Option Explicit

Public Function Echo(ByVal value As String) As String
    Echo = value
End Function
""",
        crlf=True,
    )
    _write_fixture(
        root / "tests/modules/QualityTests.bas",
        """Attribute VB_Name = "QualityTests"
Option Explicit

Public Sub RunTests()
End Sub
""",
        crlf=True,
    )
    _write_fixture(
        root / "tools/check_repo.py",
        "# Fixture command placeholder.\n",
    )
    _run_git(root, "init", "-b", "main")
    _run_git(root, "config", "user.name", "Repository Quality Self-Test")
    _run_git(root, "config", "user.email", "quality-self-test@example.invalid")
    _run_git(root, "add", "--all")
    _run_git(root, "commit", "-m", "Create passing fixture")


def _update_fixture_json(root: Path, path: str, mutate: Callable[[dict], None]) -> None:
    document = json.loads((root / path).read_text(encoding="utf-8"))
    mutate(document)
    _write_fixture(
        root / path,
        json.dumps(document, indent=2, ensure_ascii=False) + "\n",
    )


def _degrade_configuration(root: Path) -> None:
    _update_fixture_json(
        root,
        CONFIG_PATH,
        lambda document: document.__setitem__("schema_version", 99),
    )


def _degrade_required_paths(root: Path) -> None:
    (root / "README.md").unlink()


def _degrade_placeholders(root: Path) -> None:
    token = "{{" + "UNKNOWN_TOKEN}}"
    _write_fixture(root / "README.md", f"# Fixture\n\n{token}\n")


def _degrade_identity(root: Path) -> None:
    token = "DONOR" + "-PROJECT"
    _write_fixture(root / "README.md", f"# Fixture\n\n{token}\n")


def _degrade_dotfile_policy(root: Path) -> None:
    text = (root / ".editorconfig").read_text(encoding="utf-8")
    _write_fixture(root / ".editorconfig", text.replace("root = true", "root = false"))


def _degrade_structured_data(root: Path) -> None:
    _write_fixture(root / LABEL_MANIFEST_PATH, "{\n")


def _degrade_structured_yaml(root: Path) -> None:
    path = root / ".github/workflows/invalid-yaml.yml"
    _write_fixture(
        path,
        """name: Invalid YAML
on: [push
jobs:
  test:
    runs-on: ubuntu-latest
""",
    )
    _run_git(root, "add", path.relative_to(root).as_posix())


def _degrade_structured_xml(root: Path) -> None:
    path = root / "docs/invalid.xml"
    _write_fixture(path, "<repository><unclosed></repository>\n")
    _run_git(root, "add", path.relative_to(root).as_posix())


def _degrade_structured_xml_encoding(root: Path) -> None:
    path = root / "docs/invalid-encoding.xml"
    _write_fixture(path, b"<repository>\xff</repository>\n")
    _run_git(root, "add", path.relative_to(root).as_posix())


def _degrade_structured_xml_doctype(root: Path) -> None:
    path = root / "docs/doctype.xml"
    _write_fixture(
        path,
        '<!DOCTYPE repository [<!ENTITY sample "value">]><repository>&sample;</repository>\n',
    )
    _run_git(root, "add", path.relative_to(root).as_posix())


def _degrade_structured_xml_oversize(root: Path) -> None:
    path = root / "docs/oversize.xml"
    padding = "x" * MAX_XML_BYTES
    _write_fixture(path, f"<repository>{padding}</repository>\n")
    _run_git(root, "add", path.relative_to(root).as_posix())


def _degrade_markdown_links(root: Path) -> None:
    _write_fixture(root / "README.md", "# Fixture\n\n[Missing](docs/MISSING.md)\n")


def _degrade_text_integrity(root: Path) -> None:
    marker = "<" * 7
    _write_fixture(root / "README.md", f"# Fixture\n\n{marker} HEAD\n")


def _degrade_forbidden_artifacts(root: Path) -> None:
    _write_fixture(root / "secret.pem", b"fixture")
    _run_git(root, "add", "-f", "secret.pem")


def _degrade_line_endings(root: Path) -> None:
    _write_fixture(root / "README.md", "# Fixture\r\n\r\nWrong endings.\r\n", crlf=True)


def _degrade_label_manifest(root: Path) -> None:
    def mutate(document: dict) -> None:
        document["core"][0]["color"] = "d73a4a"

    _update_fixture_json(root, LABEL_MANIFEST_PATH, mutate)


def _degrade_issue_forms(root: Path) -> None:
    path = root / ISSUE_TEMPLATE_DIRECTORY / "bug.yml"
    text = path.read_text(encoding="utf-8")
    _write_fixture(path, text.replace("assignees: []", 'assignees: ["owner"]'))


def _degrade_workflow_actions(root: Path) -> None:
    path = root / ".github/workflows/static-checks.yml"
    text = path.read_text(encoding="utf-8")
    text = re.sub(r"actions/checkout@[0-9a-f]{40}", "actions/checkout@v4", text)
    _write_fixture(path, text)


def _degrade_workflow_pull_request_target(root: Path) -> None:
    path = root / ".github/workflows/static-checks.yml"
    text = path.read_text(encoding="utf-8")
    _write_fixture(path, text.replace("on:\n", "on:\n  pull_request_target:\n", 1))


def _degrade_workflow_pull_request_target_scalar(root: Path) -> None:
    path = root / ".github/workflows/static-checks.yml"
    text = path.read_text(encoding="utf-8")
    _write_fixture(
        path,
        text.replace("on:\n  pull_request:\n", "on: pull_request_target\n", 1),
    )


def _degrade_workflow_pull_request_target_flow_map(root: Path) -> None:
    path = root / ".github/workflows/static-checks.yml"
    text = path.read_text(encoding="utf-8")
    _write_fixture(
        path,
        text.replace(
            "on:\n  pull_request:\n", "on: {pull_request_target: null}\n", 1
        ),
    )


def _degrade_workflow_pr_write(root: Path) -> None:
    path = root / ".github/workflows/static-checks.yml"
    text = path.read_text(encoding="utf-8")
    _write_fixture(
        path,
        text.replace("permissions:\n  contents: read", "permissions:\n  contents: write", 1),
    )


def _degrade_workflow_pr_scalar_flow_write(root: Path) -> None:
    path = root / ".github/workflows/static-checks.yml"
    text = path.read_text(encoding="utf-8")
    text = text.replace("on:\n  pull_request:\n", "on: pull_request\n", 1)
    _write_fixture(
        path,
        text.replace(
            "permissions:\n  contents: read", "permissions: {contents: write}", 1
        ),
    )


def _degrade_workflow_pr_flow_map_job_write(root: Path) -> None:
    path = root / ".github/workflows/static-checks.yml"
    text = path.read_text(encoding="utf-8")
    text = text.replace(
        "on:\n  pull_request:\n", "on: {pull_request: null}\n", 1
    )
    _write_fixture(
        path,
        text.replace(
            "  quality:\n    runs-on: ubuntu-latest",
            "  quality:\n    permissions: {contents: write}\n    runs-on: ubuntu-latest",
            1,
        ),
    )


def _degrade_version_changelog(root: Path) -> None:
    _write_fixture(root / "VERSION", "version-one\n")


def _degrade_git_diff(root: Path) -> None:
    _write_fixture(root / "README.md", "# Fixture  \n")


def _rewrite_vba(root: Path, relative: str, transform: Callable[[str], str]) -> None:
    path = root / relative
    text = path.read_bytes().decode("cp1252").replace("\r\n", "\n")
    _write_fixture(path, transform(text), crlf=True)


def _degrade_vba_option_explicit(root: Path) -> None:
    _rewrite_vba(
        root,
        "src/modules/Quality.bas",
        lambda text: text.replace("Option Explicit\n", ""),
    )


def _degrade_vba_export_header(root: Path) -> None:
    _rewrite_vba(
        root,
        "src/modules/Quality.bas",
        lambda text: text.replace('"Quality"', '"WrongName"', 1),
    )


def _degrade_vba_structure(root: Path) -> None:
    _rewrite_vba(
        root,
        "src/modules/Quality.bas",
        lambda text: text.replace("End Function\n", ""),
    )


def _degrade_vba_visibility(root: Path) -> None:
    def mutate(document: dict) -> None:
        document["vba"]["components"]["src/modules/Quality.bas"] = "internal"

    _update_fixture_json(root, CONFIG_PATH, mutate)


def _degrade_generated_vba_contract(root: Path) -> None:
    def mutate(document: dict) -> None:
        document["profiles"]["library"]["vba_contract"]["minimum_roles"][
            "internal"
        ] = 2

    _update_fixture_json(root, CONFIG_PATH, mutate)


def _degrade_vba_public_api(root: Path) -> None:
    _write_fixture(root / "docs/PUBLIC_API.txt", "Quality\tFunction\tMissing\n")


SELF_TEST_CASES: tuple[tuple[str, Callable[[Path], None]], ...] = (
    ("configuration", _degrade_configuration),
    ("required-paths", _degrade_required_paths),
    ("placeholders", _degrade_placeholders),
    ("identity", _degrade_identity),
    ("dotfile-policy", _degrade_dotfile_policy),
    ("structured-data", _degrade_structured_data),
    ("markdown-links", _degrade_markdown_links),
    ("text-integrity", _degrade_text_integrity),
    ("forbidden-artifacts", _degrade_forbidden_artifacts),
    ("line-endings", _degrade_line_endings),
    ("label-manifest", _degrade_label_manifest),
    ("issue-forms", _degrade_issue_forms),
    ("workflow-actions", _degrade_workflow_actions),
    ("version-changelog", _degrade_version_changelog),
    ("git-diff-check", _degrade_git_diff),
    ("vba-option-explicit", _degrade_vba_option_explicit),
    ("vba-export-header", _degrade_vba_export_header),
    ("vba-structure", _degrade_vba_structure),
    ("vba-visibility", _degrade_vba_visibility),
    ("generated-vba-contract", _degrade_generated_vba_contract),
    ("vba-public-api", _degrade_vba_public_api),
)

BRANCH_SELF_TEST_CASES: tuple[
    tuple[str, str, Callable[[Path], None]], ...
] = (
    ("structured-yaml", "structured-data", _degrade_structured_yaml),
    ("structured-xml", "structured-data", _degrade_structured_xml),
    ("structured-xml-encoding", "structured-data", _degrade_structured_xml_encoding),
    ("structured-xml-doctype", "structured-data", _degrade_structured_xml_doctype),
    ("structured-xml-oversize", "structured-data", _degrade_structured_xml_oversize),
    ("workflow-pull-request-target", "workflow-actions", _degrade_workflow_pull_request_target),
    (
        "workflow-pull-request-target-scalar",
        "workflow-actions",
        _degrade_workflow_pull_request_target_scalar,
    ),
    (
        "workflow-pull-request-target-flow-map",
        "workflow-actions",
        _degrade_workflow_pull_request_target_flow_map,
    ),
    ("workflow-pr-write", "workflow-actions", _degrade_workflow_pr_write),
    (
        "workflow-pr-scalar-flow-write",
        "workflow-actions",
        _degrade_workflow_pr_scalar_flow_write,
    ),
    (
        "workflow-pr-flow-map-job-write",
        "workflow-actions",
        _degrade_workflow_pr_flow_map_job_write,
    ),
)


def _tree_digest(root: Path) -> str:
    digest = hashlib.sha256()
    for path in sorted(
        item for item in root.rglob("*") if item.is_file() and ".git" not in item.parts
    ):
        relative = path.relative_to(root).as_posix()
        digest.update(relative.encode("utf-8"))
        digest.update(b"\0")
        digest.update(path.read_bytes())
        digest.update(b"\0")
    completed = subprocess.run(
        ["git", "-C", str(root), "status", "--porcelain=v1", "-z"],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    digest.update(completed.stdout)
    return digest.hexdigest()


def run_self_test() -> int:
    failures: list[str] = []
    with tempfile.TemporaryDirectory(prefix="repository-quality-") as temporary:
        root = Path(temporary)
        _initialize_fixture(root)
        before = _tree_digest(root)
        first = build_report(root)
        middle = _tree_digest(root)
        second = build_report(root)
        after = _tree_digest(root)
        first_json = json.dumps(first, indent=2, sort_keys=True) + "\n"
        second_json = json.dumps(second, indent=2, sort_keys=True) + "\n"
        if first["status"] != "pass":
            failures.append("positive fixture did not pass")
        if first["counts"]["rules"] != len(SELF_TEST_CASES):
            failures.append(
                f"positive fixture ran {first['counts']['rules']} rules; "
                f"expected {len(SELF_TEST_CASES)}"
            )
        if first_json != second_json:
            failures.append("JSON reports differ across identical runs")
        if markdown_report(first) != markdown_report(second):
            failures.append("Markdown reports differ across identical runs")
        if before != middle or middle != after:
            failures.append("checker changed the positive fixture")

    for expected_rule, degrade in SELF_TEST_CASES:
        with tempfile.TemporaryDirectory(
            prefix=f"repository-quality-{expected_rule}-"
        ) as temporary:
            root = Path(temporary)
            _initialize_fixture(root)
            degrade(root)
            before = _tree_digest(root)
            report = build_report(root)
            after = _tree_digest(root)
            results = {result["id"]: result for result in report["rules"]}
            actual = results.get(expected_rule)
            if actual is None:
                failures.append(f"{expected_rule}: expected rule did not run")
            elif actual["status"] != "fail":
                failures.append(f"{expected_rule}: degraded fixture was not rejected")
            if before != after:
                failures.append(f"{expected_rule}: checker changed the fixture")

    for case_name, expected_rule, degrade in BRANCH_SELF_TEST_CASES:
        with tempfile.TemporaryDirectory(
            prefix=f"repository-quality-{case_name}-"
        ) as temporary:
            root = Path(temporary)
            _initialize_fixture(root)
            degrade(root)
            before = _tree_digest(root)
            report = build_report(root)
            after = _tree_digest(root)
            results = {result["id"]: result for result in report["rules"]}
            actual = results.get(expected_rule)
            if actual is None:
                failures.append(f"{case_name}: expected rule did not run")
            elif actual["status"] != "fail":
                failures.append(f"{case_name}: degraded fixture was not rejected")
            if before != after:
                failures.append(f"{case_name}: checker changed the fixture")

    if failures:
        for message in failures:
            print(f"[FAIL] {message}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1
    print(
        f"SELF-TEST PASS: {len(SELF_TEST_CASES)} rules, one positive fixture, "
        f"{len(SELF_TEST_CASES)} degraded fixtures, deterministic JSON/Markdown, "
        f"{len(BRANCH_SELF_TEST_CASES)} branch fixtures, and read-only execution."
    )
    return 0


def parse_arguments(arguments: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        default=Path.cwd(),
        help="Git worktree to inspect (default: current directory).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Write the deterministic JSON report to this path.",
    )
    parser.add_argument(
        "--summary",
        type=Path,
        help="Write the deterministic Markdown summary to this path.",
    )
    parser.add_argument(
        "--self-test",
        action="store_true",
        help="Run positive, degraded, determinism, and read-only fixtures.",
    )
    return parser.parse_args(arguments)


def main(arguments: list[str] | None = None) -> int:
    options = parse_arguments(sys.argv[1:] if arguments is None else arguments)
    if options.self_test:
        try:
            return run_self_test()
        except (OSError, subprocess.SubprocessError, UnicodeError) as error:
            print(f"SELF-TEST ERROR: {error}", file=sys.stderr)
            return 2

    try:
        report = build_report(options.root)
        json_text = json.dumps(report, indent=2, sort_keys=True) + "\n"
        markdown = markdown_report(report)
        if options.output:
            _write_report(options.output, json_text)
        if options.summary:
            _write_report(options.summary, markdown)
        print(console_report(report))
        return 0 if report["status"] == "pass" else 1
    except (OSError, subprocess.SubprocessError, UnicodeError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
