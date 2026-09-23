from __future__ import annotations

import re
from pathlib import Path
from types import ModuleType
from typing import Callable

from policy_coverage_core import mutate_config, mutate_labels, rewrite_vba

Case = tuple[str, str, str | None, Callable[[Path], None]]


def _register(cases: list[Case], name: str, rule: str, pattern: str | None = None):
    def decorator(function: Callable[[Path], None]) -> Callable[[Path], None]:
        cases.append((name, rule, pattern, function))
        return function

    return decorator


def _case(cases: list[Case]):
    def register(name: str, rule: str, pattern: str | None = None):
        return _register(cases, name, rule, pattern)

    return register


def _label_cases(module: ModuleType) -> list[Case]:
    cases: list[Case] = []
    case = _case(cases)

    @case("labels-array-type", "label-manifest", "must be an array")
    def _(root: Path) -> None:
        mutate_labels(module, root, lambda d: d.__setitem__("core", {}))

    @case("labels-item-shape", "label-manifest", "exactly name, color, and description")
    def _(root: Path) -> None:
        mutate_labels(module, root, lambda d: d["core"][0].__setitem__("extra", True))

    @case("labels-name-invalid", "label-manifest", "name must be")
    def _(root: Path) -> None:
        mutate_labels(module, root, lambda d: d["core"][0].__setitem__("name", " bad "))

    @case("labels-name-duplicate", "label-manifest", "duplicates")
    def _(root: Path) -> None:
        mutate_labels(
            module,
            root,
            lambda d: d["core"][1].__setitem__("name", d["core"][0]["name"].upper()),
        )

    @case("labels-description-invalid", "label-manifest", "description must be")
    def _(root: Path) -> None:
        mutate_labels(module, root, lambda d: d["core"][0].__setitem__("description", ""))

    @case("labels-order-invalid", "label-manifest", "sorted")
    def _(root: Path) -> None:
        mutate_labels(module, root, lambda d: d["core"].reverse())

    @case("labels-root-shape", "label-manifest", "Root must contain exactly")
    def _(root: Path) -> None:
        mutate_labels(module, root, lambda d: d.__setitem__("extra", True))

    @case("labels-schema", "label-manifest", "schema_version")
    def _(root: Path) -> None:
        mutate_labels(module, root, lambda d: d.__setitem__("schema_version", 2))

    @case("labels-prune", "label-manifest", "prune must be")
    def _(root: Path) -> None:
        mutate_labels(module, root, lambda d: d.__setitem__("prune", "no"))

    @case("labels-core-empty", "label-manifest", "core must not be empty")
    def _(root: Path) -> None:
        mutate_labels(module, root, lambda d: d.__setitem__("core", []))

    @case("labels-overlays-shape", "label-manifest", "overlays must contain exactly")
    def _(root: Path) -> None:
        mutate_labels(module, root, lambda d: d.__setitem__("overlays", {}))

    @case("labels-profile-shape", "label-manifest", "overlays.profile")
    def _(root: Path) -> None:
        mutate_labels(module, root, lambda d: d["overlays"].__setitem__("profile", {}))

    @case("labels-domain-type", "label-manifest", "overlays.domain")
    def _(root: Path) -> None:
        mutate_labels(module, root, lambda d: d["overlays"].__setitem__("domain", []))

    @case("labels-domain-name", "label-manifest", "Domain overlay name")
    def _(root: Path) -> None:
        mutate_labels(
            module,
            root,
            lambda d: d["overlays"]["domain"].__setitem__("Bad_Name", []),
        )

    @case("labels-selected-domain-missing", "label-manifest", "Selected label domain")
    def _(root: Path) -> None:
        mutate_config(module, root, lambda d: d.__setitem__("label_domains", ["missing-domain"]))

    return cases


def _issue_form_cases(module: ModuleType) -> list[Case]:
    cases: list[Case] = []
    case = _case(cases)
    form_path = f"{module.ISSUE_TEMPLATE_DIRECTORY}/bug.yml"
    form_config = f"{module.ISSUE_TEMPLATE_DIRECTORY}/config.yml"

    @case("issue-form-unreadable", "issue-forms", "Cannot read canonical issue form")
    def _(root: Path) -> None:
        (root / form_path).unlink()

    @case("issue-form-name-empty", "issue-forms", "Top-level name")
    def _(root: Path) -> None:
        text = (root / form_path).read_text(encoding="utf-8").replace(
            'name: "Bug report"', 'name: ""'
        )
        module._write_fixture(root / form_path, text)

    @case("issue-form-title", "issue-forms", "Top-level title")
    def _(root: Path) -> None:
        text = (root / form_path).read_text(encoding="utf-8").replace(
            'title: "[Bug]: "', 'title: "[Wrong]: "'
        )
        module._write_fixture(root / form_path, text)

    @case("issue-form-labels", "issue-forms", "Top-level labels")
    def _(root: Path) -> None:
        text = (root / form_path).read_text(encoding="utf-8").replace(
            'labels: ["bug"]', 'labels: ["wrong"]'
        )
        module._write_fixture(root / form_path, text)

    @case("issue-form-body-count", "issue-forms", "between 1 and 10")
    def _(root: Path) -> None:
        text = (root / form_path).read_text(encoding="utf-8")
        module._write_fixture(
            root / form_path, re.sub(r"(?ms)^body:\n.*\Z", "body:\n", text)
        )

    @case("issue-form-unsupported-type", "issue-forms", "Unsupported issue-form")
    def _(root: Path) -> None:
        text = (root / form_path).read_text(encoding="utf-8").replace(
            "  - type: markdown", "  - type: invalid", 1
        )
        module._write_fixture(root / form_path, text)

    @case("issue-form-duplicate-id", "issue-forms", "IDs must be unique")
    def _(root: Path) -> None:
        text = (root / form_path).read_text(encoding="utf-8").replace(
            "    id: summary", "    id: profile"
        )
        module._write_fixture(root / form_path, text)

    @case("issue-form-missing-id", "issue-forms", "element IDs are missing")
    def _(root: Path) -> None:
        text = (root / form_path).read_text(encoding="utf-8").replace(
            "    id: summary\n", ""
        )
        module._write_fixture(root / form_path, text)

    @case("issue-form-security-guidance", "issue-forms", "route vulnerability")
    def _(root: Path) -> None:
        text = (root / form_path).read_text(encoding="utf-8").replace(
            "SECURITY.md", "SECURITY-X.md"
        )
        module._write_fixture(root / form_path, text)

    @case("issue-form-required-field", "issue-forms", "must be mandatory")
    def _(root: Path) -> None:
        text = (root / form_path).read_text(encoding="utf-8").replace(
            "      required: true", "      required: false", 1
        )
        module._write_fixture(root / form_path, text)

    @case("issue-config-unreadable", "issue-forms", "Cannot read issue-template")
    def _(root: Path) -> None:
        (root / form_config).unlink()

    @case("issue-config-blank", "issue-forms", "Blank issues")
    def _(root: Path) -> None:
        text = (root / form_config).read_text(encoding="utf-8").replace(
            "blank_issues_enabled: false", "blank_issues_enabled: true"
        )
        module._write_fixture(root / form_config, text)

    @case("issue-config-url", "issue-forms", "Private-security contact URL")
    def _(root: Path) -> None:
        text = (root / form_config).read_text(encoding="utf-8").replace(
            "https://github.com/example/fixture/security/policy",
            "https://example.invalid/security",
        )
        module._write_fixture(root / form_config, text)

    @case("issue-config-private", "issue-forms", "require private reporting")
    def _(root: Path) -> None:
        text = (root / form_config).read_text(encoding="utf-8").replace(
            "private", "restricted"
        )
        module._write_fixture(root / form_config, text)

    return cases


def _workflow_and_version_cases(module: ModuleType) -> list[Case]:
    cases: list[Case] = []
    case = _case(cases)
    workflow_path = ".github/workflows/static-checks.yml"

    @case("workflow-action-unparseable", "workflow-actions", "cannot be parsed")
    def _(root: Path) -> None:
        text = (root / workflow_path).read_text(encoding="utf-8")
        text = re.sub(
            r"(?m)^\s*- uses: actions/checkout@[^\n]+$", "      - uses:", text, count=1
        )
        module._write_fixture(root / workflow_path, text)

    @case("workflow-action-no-revision", "workflow-actions", "include a revision")
    def _(root: Path) -> None:
        text = (root / workflow_path).read_text(encoding="utf-8")
        text = re.sub(
            r"actions/checkout@[0-9a-f]{40}", "actions/checkout", text, count=1
        )
        module._write_fixture(root / workflow_path, text)

    @case("workflow-action-comment", "workflow-actions", "audited semantic-version comment")
    def _(root: Path) -> None:
        text = (root / workflow_path).read_text(encoding="utf-8").replace(
            " # v4.2.2", "", 1
        )
        module._write_fixture(root / workflow_path, text)

    @case(
        "workflow-pr-workflow-write-all",
        "workflow-actions",
        "Workflow triggered by pull_request must not request write-capable token permissions",
    )
    def _(root: Path) -> None:
        text = (root / workflow_path).read_text(encoding="utf-8")
        updated = text.replace("permissions:\n  contents: read", "permissions: write-all", 1)
        if updated == text:
            raise AssertionError("workflow-level permission mutation did not apply")
        module._write_fixture(root / workflow_path, updated)

    @case(
        "workflow-pr-job-write-all",
        "workflow-actions",
        "Job reachable from pull_request must not request write-all permissions",
    )
    def _(root: Path) -> None:
        text = (root / workflow_path).read_text(encoding="utf-8")
        prefix, jobs = text.split("jobs:\n", 1)
        jobs, count = re.subn(
            r"(?m)^(  [A-Za-z0-9_-]+:\s*)$",
            r"\1\n    permissions: write-all",
            jobs,
            count=1,
        )
        if count != 1:
            raise AssertionError("job-level write-all mutation did not apply")
        module._write_fixture(root / workflow_path, prefix + "jobs:\n" + jobs)

    @case(
        "workflow-pr-job-write-scope",
        "workflow-actions",
        "Job reachable from pull_request must not request write-capable token permissions",
    )
    def _(root: Path) -> None:
        text = (root / workflow_path).read_text(encoding="utf-8")
        prefix, jobs = text.split("jobs:\n", 1)
        jobs, count = re.subn(
            r"(?m)^(  [A-Za-z0-9_-]+:\s*)$",
            r"\1\n    permissions:\n      issues: write",
            jobs,
            count=1,
        )
        if count != 1:
            raise AssertionError("job-level scoped-write mutation did not apply")
        module._write_fixture(root / workflow_path, prefix + "jobs:\n" + jobs)

    @case("version-unreadable", "version-changelog", "Cannot read version")
    def _(root: Path) -> None:
        (root / "VERSION").unlink()

    @case("changelog-unreadable", "version-changelog", "Cannot read changelog")
    def _(root: Path) -> None:
        (root / "CHANGELOG.md").unlink()

    @case("changelog-missing-unreleased", "version-changelog", "Unreleased")
    def _(root: Path) -> None:
        module._write_fixture(root / "CHANGELOG.md", "# Changelog\n\nFixture.\n")

    @case("changelog-release-heading", "version-changelog", "no dated release heading")
    def _(root: Path) -> None:
        module._write_fixture(root / "VERSION", "1.2.3\n")

    return cases


def _vba_export_cases(module: ModuleType) -> list[Case]:
    cases: list[Case] = []
    case = _case(cases)
    path = "src/modules/Quality.bas"

    @case("vba-export-unreadable", "vba-export-header", "cannot be read")
    def _(root: Path) -> None:
        (root / path).unlink()

    @case("vba-export-no-name", "vba-export-header", "must declare Attribute VB_Name")
    def _(root: Path) -> None:
        rewrite_vba(
            module,
            root,
            path,
            "Option Explicit\nPublic Function Echo(ByVal value As String) As String\n"
            "Echo = value\nEnd Function\n",
        )

    @case("vba-export-duplicate-name", "vba-export-header", "more than once")
    def _(root: Path) -> None:
        text = (root / path).read_bytes().decode("cp1252").replace("\r\n", "\n")
        rewrite_vba(
            module,
            root,
            path,
            text.replace("Option Explicit", 'Attribute VB_Name = "Quality"\nOption Explicit'),
        )

    @case("vba-export-name-not-line1", "vba-export-header", "must be line 1")
    def _(root: Path) -> None:
        text = (root / path).read_bytes().decode("cp1252").replace("\r\n", "\n")
        rewrite_vba(module, root, path, "' comment\n" + text)

    @case("vba-export-name-not-leading", "vba-export-header", "leading export header")
    def _(root: Path) -> None:
        text = (root / path).read_bytes().decode("cp1252").replace("\r\n", "\n")
        lines = text.splitlines()
        rewrite_vba(
            module, root, path, "\n".join(["' pad"] * 25 + [lines[0]] + lines[1:]) + "\n"
        )

    @case("vba-export-invalid-identifier", "vba-export-header", "not a VBA identifier")
    def _(root: Path) -> None:
        text = (root / path).read_bytes().decode("cp1252").replace("\r\n", "\n")
        rewrite_vba(module, root, path, text.replace('"Quality"', '"Bad-Name"', 1))

    @case("vba-export-name-too-long", "vba-export-header", "character limit")
    def _(root: Path) -> None:
        text = (root / path).read_bytes().decode("cp1252").replace("\r\n", "\n")
        rewrite_vba(module, root, path, text.replace('"Quality"', '"' + "Q" * 32 + '"', 1))

    @case("vba-export-name-collision", "vba-export-header", "collides")
    def _(root: Path) -> None:
        rewrite_vba(
            module,
            root,
            "src/modules/Other.bas",
            'Attribute VB_Name = "quality"\nOption Explicit\nPublic Sub Other()\nEnd Sub\n',
        )
        module._run_git(root, "add", "src/modules/Other.bas")

    return cases


def _vba_structure_cases(module: ModuleType) -> list[Case]:
    cases: list[Case] = []
    case = _case(cases)
    path = "src/modules/Quality.bas"

    @case("vba-nested-opener", "vba-structure", "has no closing statement")
    def _(root: Path) -> None:
        rewrite_vba(module, root, path, 'Attribute VB_Name = "Quality"\nOption Explicit\nPublic Function Echo(ByVal value As String) As String\nPublic Function Second() As String\nEnd Function\n')

    @case("vba-elseif-without-if", "vba-structure", "#ElseIf without #If")
    def _(root: Path) -> None:
        rewrite_vba(module, root, path, 'Attribute VB_Name = "Quality"\nOption Explicit\n#ElseIf VBA7 Then\nPublic Function Echo(ByVal value As String) As String\n    Echo = value\nEnd Function\n')

    @case("vba-else-without-if", "vba-structure", "#Else without #If")
    def _(root: Path) -> None:
        rewrite_vba(module, root, path, 'Attribute VB_Name = "Quality"\nOption Explicit\n#Else\nPublic Function Echo(ByVal value As String) As String\n    Echo = value\nEnd Function\n')

    @case("vba-endif-without-if", "vba-structure", "#End If without #If")
    def _(root: Path) -> None:
        rewrite_vba(module, root, path, 'Attribute VB_Name = "Quality"\nOption Explicit\n#End If\nPublic Function Echo(ByVal value As String) As String\n    Echo = value\nEnd Function\n')

    @case("vba-unclosed-if", "vba-structure", "conditional-compilation block")
    def _(root: Path) -> None:
        rewrite_vba(module, root, path, 'Attribute VB_Name = "Quality"\nOption Explicit\n#If VBA7 Then\nPublic Function Echo(ByVal value As String) As String\n    Echo = value\nEnd Function\n')

    @case("vba-reachable-declare", "vba-structure", "must include PtrSafe")
    def _(root: Path) -> None:
        rewrite_vba(module, root, path, 'Attribute VB_Name = "Quality"\nOption Explicit\n#If VBA7 Then\nPrivate Declare Function Tick Lib "kernel32" () As Long\n#End If\nPublic Function Echo(ByVal value As String) As String\n    Echo = value\nEnd Function\n')

    @case("vba-end-without-opener", "vba-structure", "has no opener")
    def _(root: Path) -> None:
        rewrite_vba(module, root, path, 'Attribute VB_Name = "Quality"\nOption Explicit\nEnd Function\nPublic Function Echo(ByVal value As String) As String\n    Echo = value\nEnd Function\n')

    @case("vba-missing-jump", "vba-structure", "Jump target is not defined")
    def _(root: Path) -> None:
        rewrite_vba(module, root, path, 'Attribute VB_Name = "Quality"\nOption Explicit\nPublic Function Echo(ByVal value As String) As String\n    GoTo Missing\n    Echo = value\nEnd Function\n')

    return cases


def _vba_contract_cases(module: ModuleType) -> list[Case]:
    cases: list[Case] = []
    case = _case(cases)

    @case("vba-unconfigured-component", "vba-visibility", "not assigned a profile role")
    def _(root: Path) -> None:
        rewrite_vba(
            module,
            root,
            "src/modules/Extra.bas",
            'Attribute VB_Name = "Extra"\nOption Explicit\nPublic Sub X()\nEnd Sub\n',
        )
        module._run_git(root, "add", "src/modules/Extra.bas")

    @case("vba-configured-not-tracked", "vba-visibility", "Configured VBA component is not tracked")
    def _(root: Path) -> None:
        def mutation(document: dict) -> None:
            document["vba"]["components"]["src/modules/Missing.bas"] = "public"
            document["vba"]["components"] = dict(
                sorted(
                    document["vba"]["components"].items(),
                    key=lambda item: item[0].casefold(),
                )
            )

        mutate_config(module, root, mutation)

    @case("vba-internal-not-private", "vba-visibility", "Internal standard module")
    def _(root: Path) -> None:
        path = root / "src/core/QualityCore.bas"
        text = (
            path.read_bytes()
            .decode("cp1252")
            .replace("\r\n", "\n")
            .replace("Option Private Module\n", "")
        )
        rewrite_vba(module, root, "src/core/QualityCore.bas", text)

    @case("vba-public-private", "vba-visibility", "Public facade")
    def _(root: Path) -> None:
        path = root / "src/modules/Quality.bas"
        text = (
            path.read_bytes()
            .decode("cp1252")
            .replace("\r\n", "\n")
            .replace("Option Explicit\n", "Option Explicit\nOption Private Module\n")
        )
        rewrite_vba(module, root, "src/modules/Quality.bas", text)

    @case("vba-test-root-role", "vba-visibility", "under a test root")
    def _(root: Path) -> None:
        mutate_config(
            module,
            root,
            lambda d: d["vba"]["components"].__setitem__(
                "tests/modules/QualityTests.bas", "public"
            ),
        )

    @case("vba-source-root-test-role", "vba-visibility", "must not be stored under a source root")
    def _(root: Path) -> None:
        mutate_config(
            module,
            root,
            lambda d: d["vba"]["components"].__setitem__(
                "src/modules/Quality.bas", "test"
            ),
        )

    @case("generated-required-unregistered", "generated-vba-contract", "requires this")
    def _(root: Path) -> None:
        mutate_config(
            module,
            root,
            lambda d: d["vba"]["components"].pop("src/modules/Quality.bas"),
        )

    @case("generated-required-wrong-role", "generated-vba-contract", "requires role")
    def _(root: Path) -> None:
        mutate_config(
            module,
            root,
            lambda d: d["vba"]["components"].__setitem__(
                "src/modules/Quality.bas", "internal"
            ),
        )

    @case("generated-required-not-tracked", "generated-vba-contract", "starter component to be tracked")
    def _(root: Path) -> None:
        def mutation(document: dict) -> None:
            contract = document["profiles"]["library"]["vba_contract"]
            contract["required_components"].pop("src/modules/Quality.bas")
            contract["required_components"]["src/modules/Missing.bas"] = "public"
            contract["required_components"] = dict(
                sorted(contract["required_components"].items())
            )
            document["vba"]["components"]["src/modules/Missing.bas"] = "public"
            document["vba"]["components"] = dict(
                sorted(document["vba"]["components"].items())
            )

        mutate_config(module, root, mutation)

    @case("legacy-api-collision", "vba-public-api", "collides")
    def _(root: Path) -> None:
        rewrite_vba(
            module,
            root,
            "src/modules/Other.bas",
            'Attribute VB_Name = "Other"\nOption Explicit\nPublic Function Echo() As String\nEnd Function\n',
        )
        module._run_git(root, "add", "src/modules/Other.bas")

        def mutation(document: dict) -> None:
            document["vba"]["components"]["src/modules/Other.bas"] = "public"
            document["vba"]["components"] = dict(
                sorted(
                    document["vba"]["components"].items(),
                    key=lambda item: item[0].casefold(),
                )
            )

        mutate_config(module, root, mutation)

    @case("legacy-api-manifest-missing", "vba-public-api", "manifest is not tracked")
    def _(root: Path) -> None:
        (root / "docs/PUBLIC_API.txt").unlink()

    @case("legacy-api-manifest-invalid-encoding", "vba-public-api", "Cannot read public API manifest")
    def _(root: Path) -> None:
        (root / "docs/PUBLIC_API.txt").write_bytes(b"\xff\n")

    return cases


def quality_cases(module: ModuleType) -> list[Case]:
    return [
        *_label_cases(module),
        *_issue_form_cases(module),
        *_workflow_and_version_cases(module),
        *_vba_export_cases(module),
        *_vba_structure_cases(module),
        *_vba_contract_cases(module),
    ]
