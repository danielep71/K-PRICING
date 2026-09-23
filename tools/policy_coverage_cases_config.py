from __future__ import annotations

from pathlib import Path
from types import ModuleType
from typing import Callable

from policy_coverage_core import mutate_config

Case = tuple[str, str, str | None, Callable[[Path], None]]
Mutation = Callable[[dict], None]


def _assign(**values: object) -> Mutation:
    def mutate(document: dict) -> None:
        document.update(values)

    return mutate


def _case(
    cases: list[Case],
    module: ModuleType,
    name: str,
    mutation: Mutation,
    pattern: str | None = None,
) -> None:
    def apply(root: Path) -> None:
        mutate_config(module, root, mutation)

    cases.append((name, "configuration", pattern, apply))


def _root_and_profile_cases(module: ModuleType) -> list[Case]:
    cases: list[Case] = []

    def invalid_json(root: Path) -> None:
        module._write_fixture(root / module.CONFIG_PATH, "{\n")

    cases.append(
        (
            "config-invalid-json",
            "configuration",
            "Cannot load repository profile",
            invalid_json,
        )
    )
    entries: list[tuple[str, Mutation, str | None]] = [
        ("config-root-keys", lambda d: d.__setitem__("extra", True), "canonical configuration keys"),
        ("config-mode-invalid", lambda d: d.__setitem__("mode", "invalid"), "mode must be"),
        ("config-template-profile", _assign(mode="template", profile="library"), "Template mode requires profile"),
        ("config-generated-profile", lambda d: d.__setitem__("profile", "invalid"), "Generated mode requires profile"),
        ("config-repository-form", lambda d: d.__setitem__("repository", "invalid"), "owner/name"),
        ("config-string-list-type", lambda d: d.__setitem__("required_paths", None), "array of strings"),
        ("config-string-list-duplicate-order", lambda d: d.__setitem__("required_paths", ["z", "a", "a"]), None),
        ("config-invalid-relative-path", lambda d: d.__setitem__("required_paths", ["../bad"]), "invalid relative path"),
        ("config-label-domain-name", lambda d: d.__setitem__("label_domains", ["Bad_Name"]), "non-kebab-case"),
        (
            "config-template-domains",
            _assign(
                mode="template",
                profile=None,
                repository="example/TEMPLATE-IDENTITY",
                label_domains=["domain"],
            ),
            "Template mode requires label_domains",
        ),
        ("config-profiles-shape", lambda d: d.__setitem__("profiles", {}), "profiles must contain exactly"),
        ("config-profile-entry-shape", lambda d: d["profiles"]["library"].__setitem__("extra", True), "required_paths, required_directories"),
        ("config-contract-shape", lambda d: d["profiles"]["library"].__setitem__("vba_contract", {}), "minimum_roles and required_components"),
        ("config-minimum-roles-empty", lambda d: d["profiles"]["library"]["vba_contract"].__setitem__("minimum_roles", {}), "minimum_roles must be a non-empty"),
        ("config-minimum-roles-order", lambda d: d["profiles"]["library"]["vba_contract"].__setitem__("minimum_roles", {"test": 1, "public": 1, "internal": 1}), "minimum_roles keys must be sorted"),
        ("config-minimum-role-invalid", lambda d: d["profiles"]["library"]["vba_contract"]["minimum_roles"].__setitem__("wrong", 1), "invalid role"),
        ("config-minimum-value-invalid", lambda d: d["profiles"]["library"]["vba_contract"]["minimum_roles"].__setitem__("internal", 0), "positive integer"),
        ("config-required-components-empty", lambda d: d["profiles"]["library"]["vba_contract"].__setitem__("required_components", {}), "required_components must be a non-empty"),
    ]
    for name, mutation, pattern in entries:
        _case(cases, module, name, mutation, pattern)
    return cases


def _component_and_placeholder_cases(module: ModuleType) -> list[Case]:
    cases: list[Case] = []
    entries: list[tuple[str, Mutation, str | None]] = [
        ("config-required-components-order", lambda d: d["profiles"]["library"]["vba_contract"].__setitem__("required_components", {"tests/modules/QualityTests.bas": "test", "src/modules/Quality.bas": "public", "src/core/QualityCore.bas": "internal"}), "required_components keys must be sorted"),
        ("config-required-component-path", lambda d: d["profiles"]["library"]["vba_contract"]["required_components"].__setitem__("../bad.bas", "public"), "invalid path"),
        ("config-required-component-role", lambda d: d["profiles"]["library"]["vba_contract"]["required_components"].__setitem__("src/modules/Quality.bas", "wrong"), "invalid role"),
        ("config-baseline-minimum-role", lambda d: d["profiles"]["library"]["vba_contract"]["minimum_roles"].pop("public"), "baseline role"),
        ("config-baseline-component-role", lambda d: d["profiles"]["library"]["vba_contract"]["required_components"].pop("src/modules/Quality.bas"), "must name a component with role"),
        ("config-placeholders-shape", lambda d: d.__setitem__("placeholders", {}), "placeholders must contain exactly"),
        ("config-placeholder-pattern-type", lambda d: d["placeholders"].__setitem__("pattern", 7), "pattern must be a string"),
        ("config-placeholder-pattern-invalid", lambda d: d["placeholders"].__setitem__("pattern", "("), "pattern is invalid"),
        ("config-placeholder-pattern-groups", lambda d: d["placeholders"].__setitem__("pattern", r"\{\{[A-Z]+\}\}"), "exactly one capture group"),
        ("config-placeholder-catalogue-empty", lambda d: d["placeholders"].__setitem__("catalogue", {}), "catalogue must be a non-empty"),
        ("config-placeholder-catalogue-order", lambda d: d["placeholders"].__setitem__("catalogue", {"ZZZ": {"category": "required", "description": "z"}, **d["placeholders"]["catalogue"]}), "catalogue keys must be sorted"),
        ("config-placeholder-name", lambda d: d["placeholders"]["catalogue"].__setitem__("bad-name", {"category": "required", "description": "x"}), "canonical placeholder name"),
        ("config-placeholder-object", lambda d: d["placeholders"]["catalogue"].__setitem__("OPTIONAL_NOTE", "bad"), "must be an object"),
        ("config-placeholder-category", lambda d: d["placeholders"]["catalogue"]["OPTIONAL_NOTE"].__setitem__("category", "bad"), "category must be"),
        ("config-placeholder-description", lambda d: d["placeholders"]["catalogue"]["OPTIONAL_NOTE"].__setitem__("description", ""), "description must be non-empty"),
        ("config-placeholder-profile-values-shape", lambda d: d["placeholders"]["catalogue"]["PROFILE_NOTE"].__setitem__("values", {}), "values must cover exactly"),
        ("config-placeholder-profile-values-empty", lambda d: d["placeholders"]["catalogue"]["PROFILE_NOTE"]["values"].__setitem__("library", ""), "values must all be non-empty"),
        ("config-placeholder-repeatable-format", lambda d: d["placeholders"]["catalogue"]["REPEATABLE_NOTE"].__setitem__("item_format", "bad"), "item_format must contain one"),
    ]
    for name, mutation, pattern in entries:
        _case(cases, module, name, mutation, pattern)
    return cases


def _identity_and_vba_cases(module: ModuleType) -> list[Case]:
    cases: list[Case] = []
    entries: list[tuple[str, Mutation, str | None]] = [
        ("config-placeholder-extra-fields", lambda d: d["placeholders"]["catalogue"]["OPTIONAL_NOTE"].__setitem__("extra", True), "fields inconsistent"),
        ("config-placeholder-missing-category", lambda d: d["placeholders"]["catalogue"].pop("OPTIONAL_NOTE"), "does not exercise categories"),
        ("config-placeholder-markers", lambda d: d["placeholders"].__setitem__("block_markers", {}), "canonical marker grammar"),
        ("config-identity-shape", lambda d: d.__setitem__("identity", {}), "identity must contain exactly"),
        ("config-template-tokens-empty", lambda d: d["identity"].__setitem__("template_tokens", []), "template_tokens must not be empty"),
        ("config-repository-forbidden-token", lambda d: d.__setitem__("repository", "example/DONOR-PROJECT"), "forbidden donor token"),
        ("config-template-identity-missing", _assign(mode="template", profile=None, repository="example/plain"), "Template mode repository"),
        ("config-generated-template-identity", lambda d: d.__setitem__("repository", "example/TEMPLATE-IDENTITY"), "Generated mode repository"),
        ("config-vba-shape", lambda d: d.__setitem__("vba", {}), "vba must contain exactly"),
        ("config-vba-roots-overlap", lambda d: d["vba"].__setitem__("test_roots", ["src"]), "source and test roots must not overlap"),
        ("config-vba-components-type", lambda d: d["vba"].__setitem__("components", []), "components must be an object"),
        ("config-vba-components-order", lambda d: d["vba"].__setitem__("components", {"tests/modules/QualityTests.bas": "test", "src/modules/Quality.bas": "public", "src/core/QualityCore.bas": "internal"}), "components keys must be sorted"),
        ("config-vba-component-path", lambda d: d["vba"]["components"].__setitem__("../bad.bas", "public"), "Invalid VBA component path"),
        ("config-vba-component-role", lambda d: d["vba"]["components"].__setitem__("src/modules/Quality.bas", "wrong"), "unsupported role"),
        ("config-vba-api-manifest", lambda d: d["vba"].__setitem__("public_api_manifest", "../bad"), "public_api_manifest"),
    ]
    for name, mutation, pattern in entries:
        _case(cases, module, name, mutation, pattern)
    return cases


def configuration_cases(module: ModuleType) -> list[Case]:
    return [
        *_root_and_profile_cases(module),
        *_component_and_placeholder_cases(module),
        *_identity_and_vba_cases(module),
    ]
