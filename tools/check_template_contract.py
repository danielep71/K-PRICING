#!/usr/bin/env python3
"""Validate the adopted template contract recorded by a repository.

The template contract versions the *set of controls* a generated repository
adopts. It is deliberately independent of the project's own ``VERSION``: a
template release that changes no required control ships without a contract bump,
and a project may release any number of its own versions while its adopted
contract stays fixed.

This gate owns the contract's semantics. ``check_repo.py`` only requires the
``template_contract`` key to exist in the canonical configuration; every rule
about its content lives here, so the portable checker's policy-branch inventory
is unchanged.
"""

from __future__ import annotations

import json
import re
import sys
import tempfile
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from _gatelib import parse_report_args as parse_arguments
from _gatelib import run_gate, write_text

CONFIG_PATH = ".github/repository-profile.json"
RECORD_PATH = ".github/initialization.json"
NOTES_PATH = "docs/TEMPLATE_CONTRACT.md"
VERSION_PATH = "VERSION"

CONTRACT_KEYS = {"version", "source"}
CONTRACT_VERSION_PATTERN = re.compile(r"(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)$")
REPOSITORY_PATTERN = re.compile(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$")

# Each supported contract version maps to the rule set a repository adopting it
# must satisfy. Conformance tooling selects the entry for the *recorded*
# version, so an older adopter is never judged against newer required controls.
CONTRACT_RULE_SETS: dict[str, frozenset[str]] = {
    "1.0.0": frozenset(
        {
            "canonical-repository-gate",
            "deterministic-initializer",
            "label-policy",
            "placeholder-schema",
            "profile-model",
            "release-integrity",
        }
    ),
    "1.1.0": frozenset(
        {
            "canonical-repository-gate",
            "committed-whitespace",
            "complete-public-api",
            "deterministic-initializer",
            "label-drift-detection",
            "label-policy",
            "nested-conditional-compilation",
            "placeholder-schema",
            "procedure-scoped-jumps",
            "profile-model",
            "release-integrity",
            "repository-local-actions",
            "strict-release-semantics",
        }
    ),
    "1.2.0": frozenset(
        {
            "advanced-release-provenance",
            "canonical-repository-gate",
            "committed-whitespace",
            "complete-public-api",
            "controlled-dependency-updates",
            "documentation-drift",
            "deterministic-initializer",
            "label-drift-detection",
            "label-policy",
            "nested-conditional-compilation",
            "placeholder-schema",
            "procedure-scoped-jumps",
            "profile-model",
            "release-integrity",
            "repository-local-actions",
            "strict-release-semantics",
            "template-contract-version",
        }
    ),
}
SUPPORTED_VERSIONS = tuple(sorted(CONTRACT_RULE_SETS))


class ContractError(RuntimeError):
    """The gate could not complete: unreadable or unparsable input."""


def _read_json(root: Path, relative: str) -> Any:
    path = root / relative
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as error:
        raise ContractError(f"{relative} is missing.") from error
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ContractError(f"Cannot read {relative}: {error}") from error


def _finding(message: str) -> dict[str, str]:
    return {"message": message}


def _check_shape(contract: Any, findings: list[dict[str, str]]) -> tuple[str, str]:
    """Validate the contract object's shape; return recorded version and source."""
    if not isinstance(contract, dict):
        findings.append(_finding(f"{CONFIG_PATH}: template_contract must be an object."))
        return "", ""
    if set(contract) != CONTRACT_KEYS:
        findings.append(
            _finding(
                f"{CONFIG_PATH}: template_contract must contain exactly version and source."
            )
        )
    version = contract.get("version")
    source = contract.get("source")
    if not isinstance(version, str) or not CONTRACT_VERSION_PATTERN.fullmatch(version):
        findings.append(
            _finding(
                f"{CONFIG_PATH}: template_contract.version must be a canonical "
                f"MAJOR.MINOR.PATCH version without pre-release or build metadata; "
                f"found {version!r}."
            )
        )
        version = ""
    if not isinstance(source, str) or not REPOSITORY_PATTERN.fullmatch(source):
        findings.append(
            _finding(f"{CONFIG_PATH}: template_contract.source must use the owner/name form.")
        )
        source = ""
    return version, source


def _check_supported(version: str, findings: list[dict[str, str]]) -> frozenset[str]:
    """Resolve the rule set for a recorded version, or fail actionably."""
    if not version:
        return frozenset()
    rules = CONTRACT_RULE_SETS.get(version)
    if rules is None:
        findings.append(
            _finding(
                f"{CONFIG_PATH}: template_contract.version {version!r} is not a supported "
                f"contract. Supported versions: {', '.join(SUPPORTED_VERSIONS)}. Adopt one of "
                f"them, or upgrade the tooling that publishes the newer contract."
            )
        )
        return frozenset()
    return rules


def _check_mode(
    config: dict[str, Any], source: str, findings: list[dict[str, str]]
) -> None:
    mode = config.get("mode")
    repository = config.get("repository")
    if mode == "template" and source and source != repository:
        findings.append(
            _finding(
                f"{CONFIG_PATH}: in template mode template_contract.source must equal "
                f"repository; the template publishes its own contract."
            )
        )
    if mode == "generated" and source and source == repository:
        findings.append(
            _finding(
                f"{CONFIG_PATH}: in generated mode template_contract.source must name the "
                f"template that published the contract, not the generated repository. "
                f"Initialization preserves it verbatim."
            )
        )


def _check_notes(root: Path, findings: list[dict[str, str]]) -> None:
    """Every supported version must carry migration notes."""
    try:
        notes = (root / NOTES_PATH).read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        raise ContractError(f"Cannot read {NOTES_PATH}: {error}") from error
    for version in SUPPORTED_VERSIONS:
        if not re.search(rf"^###\s+{re.escape(version)}\b", notes, re.MULTILINE):
            findings.append(
                _finding(
                    f"{NOTES_PATH}: supported contract {version} has no migration-notes section."
                )
            )


def _check_record(
    root: Path, config: dict[str, Any], version: str, source: str,
    findings: list[dict[str, str]],
) -> str | None:
    """A generated repository must record the same contract it configures."""
    if config.get("mode") != "generated":
        return None
    record = _read_json(root, RECORD_PATH)
    if not isinstance(record, dict):
        findings.append(_finding(f"{RECORD_PATH}: initialization record must be an object."))
        return None
    adopted = record.get("template_contract")
    if not isinstance(adopted, dict) or set(adopted) != CONTRACT_KEYS:
        findings.append(
            _finding(
                f"{RECORD_PATH}: initialization record must record template_contract with "
                f"exactly version and source."
            )
        )
        return None
    if adopted.get("version") != version or adopted.get("source") != source:
        findings.append(
            _finding(
                f"{RECORD_PATH}: recorded adopted contract "
                f"{adopted.get('version')!r} from {adopted.get('source')!r} disagrees with "
                f"{CONFIG_PATH} ({version!r} from {source!r})."
            )
        )
    return str(adopted.get("version"))


def _project_version(root: Path) -> str | None:
    try:
        return (root / VERSION_PATH).read_text(encoding="utf-8").strip()
    except (OSError, UnicodeError):
        return None


def run_check(root: Path) -> dict[str, Any]:
    config = _read_json(root, CONFIG_PATH)
    if not isinstance(config, dict):
        raise ContractError(f"{CONFIG_PATH} must contain a JSON object.")
    findings: list[dict[str, str]] = []
    if "template_contract" not in config:
        findings.append(
            _finding(
                f"{CONFIG_PATH}: template_contract is missing. Record the adopted contract "
                f"version and the template that published it."
            )
        )
        version, source = "", ""
        rules: frozenset[str] = frozenset()
    else:
        version, source = _check_shape(config["template_contract"], findings)
        rules = _check_supported(version, findings)
        _check_mode(config, source, findings)
    _check_notes(root, findings)
    recorded = _check_record(root, config, version, source, findings)
    return {
        "gate": "template-contract",
        "status": "pass" if not findings else "fail",
        "mode": config.get("mode"),
        "contract_version": version or None,
        "contract_source": source or None,
        "recorded_adopted_version": recorded,
        "project_version": _project_version(root),
        "resolved_rule_set": sorted(rules),
        "supported_versions": list(SUPPORTED_VERSIONS),
        "findings": findings,
    }


def markdown_report(report: dict[str, Any]) -> str:
    lines = [
        "# Template contract",
        "",
        f"- **Status:** {report['status'].upper()}",
        f"- **Mode:** {report['mode']}",
        f"- **Adopted contract:** {report['contract_version']} "
        f"from `{report['contract_source']}`",
        f"- **Project version:** {report['project_version']} "
        f"(independent of the contract)",
        f"- **Resolved rule set:** {len(report['resolved_rule_set'])} controls",
        f"- **Supported contracts:** {', '.join(report['supported_versions'])}",
        "",
    ]
    if report["resolved_rule_set"]:
        lines.append("## Controls required by the adopted contract")
        lines.append("")
        lines.extend(f"- `{name}`" for name in report["resolved_rule_set"])
        lines.append("")
    lines.append(f"**Findings:** {len(report['findings'])}")
    if report["findings"]:
        lines.append("")
        lines.extend(f"- {item['message']}" for item in report["findings"])
    return "\n".join(lines) + "\n"


def _fixture(root: Path, config: dict[str, Any], *, record: Any = None) -> Path:
    root.mkdir(parents=True, exist_ok=True)
    (root / ".github").mkdir(parents=True, exist_ok=True)
    (root / "docs").mkdir(parents=True, exist_ok=True)
    write_text(root / CONFIG_PATH, json.dumps(config, indent=2) + "\n")
    write_text(root / VERSION_PATH, "9.9.9\n")
    notes = "# Notes\n\n" + "\n".join(
        f"### {version}\n\nMigration notes.\n" for version in SUPPORTED_VERSIONS
    )
    write_text(root / NOTES_PATH, notes)
    if record is not None:
        write_text(root / RECORD_PATH, json.dumps(record, indent=2) + "\n")
    return root


def _template_config(**overrides: Any) -> dict[str, Any]:
    config = {
        "mode": "template",
        "repository": "owner/template",
        "template_contract": {"version": SUPPORTED_VERSIONS[-1], "source": "owner/template"},
    }
    config.update(overrides)
    return config


def _generated_config(**overrides: Any) -> dict[str, Any]:
    config = {
        "mode": "generated",
        "repository": "owner/product",
        "template_contract": {"version": SUPPORTED_VERSIONS[-1], "source": "owner/template"},
    }
    config.update(overrides)
    return config


def _generated_record(**overrides: Any) -> dict[str, Any]:
    record = {
        "schema_version": 1,
        "profile": "library",
        "template_contract": {"version": SUPPORTED_VERSIONS[-1], "source": "owner/template"},
        "values": {},
    }
    record.update(overrides)
    return record


def _cases() -> list[tuple[str, dict[str, Any], Any, bool, str]]:
    """(name, config, record, expect_pass, expected_substring)."""
    latest = SUPPORTED_VERSIONS[-1]
    return [
        ("template-valid", _template_config(), None, True, ""),
        (
            "generated-valid",
            _generated_config(),
            _generated_record(),
            True,
            "",
        ),
        (
            "oldest-supported-contract-resolves",
            _template_config(
                template_contract={"version": SUPPORTED_VERSIONS[0], "source": "owner/template"}
            ),
            None,
            True,
            "",
        ),
        (
            "absent-contract",
            {"mode": "template", "repository": "owner/template"},
            None,
            False,
            "template_contract is missing",
        ),
        (
            "contract-not-an-object",
            _template_config(template_contract="1.2.0"),
            None,
            False,
            "must be an object",
        ),
        (
            "unsupported-version",
            _template_config(
                template_contract={"version": "9.9.9", "source": "owner/template"}
            ),
            None,
            False,
            "is not a supported contract",
        ),
        (
            "non-canonical-version",
            _template_config(
                template_contract={"version": "1.2.0-rc.1", "source": "owner/template"}
            ),
            None,
            False,
            "canonical MAJOR.MINOR.PATCH",
        ),
        (
            "bad-source-form",
            _template_config(
                template_contract={"version": latest, "source": "not-a-repo"}
            ),
            None,
            False,
            "owner/name form",
        ),
        (
            "extra-key",
            _template_config(
                template_contract={"version": latest, "source": "owner/template", "extra": 1}
            ),
            None,
            False,
            "exactly version and source",
        ),
        (
            "template-source-mismatch",
            _template_config(
                template_contract={"version": latest, "source": "other/elsewhere"}
            ),
            None,
            False,
            "must equal repository",
        ),
        (
            "generated-source-is-self",
            _generated_config(
                template_contract={"version": latest, "source": "owner/product"}
            ),
            _generated_record(
                template_contract={"version": latest, "source": "owner/product"}
            ),
            False,
            "not the generated repository",
        ),
        (
            "record-disagrees",
            _generated_config(),
            _generated_record(
                template_contract={"version": SUPPORTED_VERSIONS[0], "source": "owner/template"}
            ),
            False,
            "disagrees with",
        ),
        (
            "record-missing-contract",
            _generated_config(),
            _generated_record(template_contract=None),
            False,
            "exactly version and source",
        ),
    ]


def _run_case(
    base: Path, name: str, config: dict[str, Any], record: Any
) -> dict[str, Any]:
    root = _fixture(base / name, config, record=record)
    return run_check(root)


def run_self_test() -> int:
    failures: list[str] = []
    with tempfile.TemporaryDirectory() as raw:
        base = Path(raw)
        for name, config, record, expect_pass, expected in _cases():
            report = _run_case(base, name, config, record)
            passed = report["status"] == "pass"
            if passed != expect_pass:
                failures.append(
                    f"{name}: expected {'pass' if expect_pass else 'fail'}, got {report['status']}"
                )
                continue
            if expected:
                text = " ".join(item["message"] for item in report["findings"])
                if expected not in text:
                    failures.append(f"{name}: no finding mentioning {expected!r}")

        # Independence: the project VERSION differs from the contract version in
        # every fixture, and changing it must not alter the resolved contract.
        root = _fixture(base / "independence", _generated_config(), record=_generated_record())
        before = run_check(root)
        write_text(root / VERSION_PATH, "42.0.0\n")
        after = run_check(root)
        if before["contract_version"] != after["contract_version"]:
            failures.append("independence: project VERSION changed the adopted contract")
        if after["project_version"] == after["contract_version"]:
            failures.append("independence: fixture does not exercise differing versions")
        if after["status"] != "pass":
            failures.append("independence: contract stopped validating after a VERSION change")

        # Rule-set selection must follow the recorded version, not the newest.
        old = _run_case(
            base,
            "ruleset-selection",
            _template_config(
                template_contract={
                    "version": SUPPORTED_VERSIONS[0],
                    "source": "owner/template",
                }
            ),
            None,
        )
        if set(old["resolved_rule_set"]) != set(CONTRACT_RULE_SETS[SUPPORTED_VERSIONS[0]]):
            failures.append("ruleset-selection: recorded version did not select its own rule set")
        if "template-contract-version" in old["resolved_rule_set"]:
            failures.append(
                "ruleset-selection: an older contract was judged against a newer control"
            )

        # Determinism.
        first = _run_case(base, "determinism", _template_config(), None)
        second = _run_case(base, "determinism", _template_config(), None)
        if json.dumps(first, sort_keys=True) != json.dumps(second, sort_keys=True):
            failures.append("determinism: repeated runs differ")

    if failures:
        print("SELF-TEST FAIL:", file=sys.stderr)
        for item in failures:
            print(f"  - {item}", file=sys.stderr)
        return 1
    print(
        "SELF-TEST PASS: contract shape, supported-version resolution, template/generated "
        "source invariants, initialization-record agreement, migration-notes coverage, "
        "rule-set selection by recorded version, VERSION independence and determinism."
    )
    return 0


def main(argv: list[str] | None = None) -> int:
    options = parse_arguments(
        sys.argv[1:] if argv is None else argv,
        description="Validate the adopted template contract.",
    )
    return run_gate(
        options,
        build=lambda: run_check(options.root),
        markdown=markdown_report,
        errors=(ContractError, OSError, ValueError),
        self_test=run_self_test,
        self_test_error_prefix="SELF-TEST ERROR",
    )


if __name__ == "__main__":
    raise SystemExit(main())
