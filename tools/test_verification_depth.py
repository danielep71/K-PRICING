#!/usr/bin/env python3
"""Focused behavioral tests for release-critical Python verification depth."""

from __future__ import annotations

import ast
import base64
import json
import tempfile
import unittest
from email.message import Message
from pathlib import Path
from typing import Any
from unittest.mock import patch

import _gatelib as gatelib
import _release_closeout as closeout
import check_committed_whitespace as whitespace
import check_external_links as external_links
import check_local_actions as local_actions
import check_release as release
import check_release_semantics as semantics
import check_template_contract as contract
import check_wiki as wiki
import checker_development as checker_dev
import collect_portfolio_snapshot as snapshot
import initialize_repository as initializer
import policy_coverage_runner as coverage_runner
import provision_repository as provision


class InitializerDepthTests(unittest.TestCase):
    def test_load_config_failure_shapes(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            with self.assertRaises(initializer.InitializationError):
                initializer._load_config(root)
            (root / ".github").mkdir()
            config = root / initializer.CONFIG_PATH

            config.write_text("{", encoding="utf-8")
            with self.assertRaises(initializer.InitializationError):
                initializer._load_config(root)

            config.write_text("[]", encoding="utf-8")
            with self.assertRaises(initializer.InitializationError):
                initializer._load_config(root)

            documents: tuple[dict[str, object], ...] = (
                {},
                {"placeholders": []},
                {"placeholders": {"pattern": "(", "catalogue": {}, "block_markers": {},
                                  "template_only_paths": [], "exclude_paths": []}},
                {"placeholders": {"pattern": "(A)(B)", "catalogue": {}, "block_markers": {},
                                  "template_only_paths": [], "exclude_paths": []}},
                {"placeholders": {"pattern": "(A)", "catalogue": [], "block_markers": {},
                                  "template_only_paths": [], "exclude_paths": []}},
                {"placeholders": {"pattern": "(A)", "catalogue": {}, "block_markers": {},
                                  "template_only_paths": "x", "exclude_paths": []}},
            )
            for document in documents:
                config.write_text(json.dumps(document), encoding="utf-8")
                with self.assertRaises(initializer.InitializationError):
                    initializer._load_config(root)

    def test_parse_assignments_rejects_bad_values(self) -> None:
        self.assertEqual(
            initializer._parse_assignments(["NAME=value", "NAME=second"], "--set"),
            {"NAME": ["value", "second"]},
        )
        bad = ("NOVALUE", "bad-name=x", "NAME=", "NAME=a\nb", "NAME=" + "{" * 2 + "TOKEN" + "}" * 2)
        for entry in bad:
            with self.subTest(entry=entry):
                with self.assertRaises(initializer.InitializationError):
                    initializer._parse_assignments([entry], "--set")

    def test_validate_values_rejects_schema_and_value_errors(self) -> None:
        catalogue = {
            "REQ": {"category": "required"},
            "OPT": {"category": "optional"},
            "REP": {"category": "repeatable"},
            "PROF": {"category": "profile-specific"},
            "REPOSITORY_PATH": {"category": "required"},
            "SUPPORT_CONTACT": {"category": "required"},
            "COPYRIGHT_YEAR": {"category": "required"},
            "MAINTAINER_NAME": {"category": "required"},
            "PROJECT_NAME": {"category": "required"},
            "PROJECT_TAGLINE": {"category": "required"},
            "PROJECT_DESCRIPTION": {"category": "required"},
            "SOCIAL_PREVIEW_PATH": {"category": "optional"},
        }
        base = [
            "REQ=x",
            "REPOSITORY_PATH=owner/repo",
            "SUPPORT_CONTACT=security@example.invalid",
            "COPYRIGHT_YEAR=2026",
            "MAINTAINER_NAME=M",
            "PROJECT_NAME=P",
            "PROJECT_TAGLINE=T",
            "PROJECT_DESCRIPTION=D",
        ]
        root = Path(".")
        with self.assertRaisesRegex(initializer.InitializationError, "Unknown substitutions"):
            initializer._validate_values(root, set(), catalogue, base + ["UNKNOWN=x"], [])
        with self.assertRaisesRegex(initializer.InitializationError, "more than once"):
            initializer._validate_values(root, set(), catalogue, base + ["REQ=y"], [])
        with self.assertRaisesRegex(initializer.InitializationError, "cannot be supplied with --set"):
            initializer._validate_values(root, set(), catalogue, base + ["REP=x"], [])
        with self.assertRaisesRegex(initializer.InitializationError, "cannot be supplied with --add"):
            initializer._validate_values(root, set(), catalogue, base, ["OPT=x"])
        with self.assertRaisesRegex(initializer.InitializationError, "Missing required"):
            initializer._validate_values(root, set(), catalogue, base[1:], [])
        for replacement, expected in (
            ("REPOSITORY_PATH=bad", "owner/name"),
            ("SUPPORT_CONTACT=bad", "email address or HTTPS"),
            ("COPYRIGHT_YEAR=1999", "four-digit year"),
        ):
            values = [
                replacement if item.split("=")[0] == replacement.split("=")[0] else item
                for item in base
            ]
            with self.assertRaisesRegex(initializer.InitializationError, expected):
                initializer._validate_values(root, set(), catalogue, values, [])
        values = [item for item in base if not item.startswith("PROJECT_NAME=")]
        values.append("PROJECT_NAME=" + "x" * 101)
        with self.assertRaisesRegex(initializer.InitializationError, "100-character"):
            initializer._validate_values(root, set(), catalogue, values, [])
        with self.assertRaisesRegex(initializer.InitializationError, "tracked repository-relative"):
            initializer._validate_values(
                root, set(), catalogue, base + ["SOCIAL_PREVIEW_PATH=assets/x.png"], []
            )
        with self.assertRaisesRegex(initializer.InitializationError, "repository-relative"):
            initializer._validate_values(
                root, set(), catalogue, base + ["SOCIAL_PREVIEW_PATH=../x.png"], [],
                require_preview_file=False,
            )

    def test_render_blocks_behaviors_and_failures(self) -> None:
        marker = "<!-- " + "template:"
        catalogue = {
            "OPT": {"category": "optional"},
            "REP": {"category": "repeatable"},
            "BAD": {"category": "required"},
        }
        text = (
            "a\n"
            "" + marker + "remove:start -->\nremove\n" + marker + "remove:end -->\n"
            "" + marker + "profile:library:start -->\nlib\n" + marker + "profile:library:end -->\n"
            "" + marker + "optional:OPT:start -->\nopt\n" + marker + "optional:OPT:end -->\n"
            "" + marker + "repeatable:REP:start -->\nrep\n" + marker + "repeatable:REP:end -->\n"
        )
        rendered = initializer._render_blocks(
            "README.md", text, "library", {"OPT": "yes"}, {"REP": ["x"]}, catalogue
        )
        self.assertEqual(rendered, "a\nlib\nopt\nrep\n")
        cases = (
            ("bad " + marker + "oops -->\n", "invalid template block marker"),
            ("" + marker + "remove:start -->\n" + marker + "remove:start -->\n", "may not nest"),
            ("" + marker + "optional:BAD:start -->\n", "is not optional"),
            ("" + marker + "repeatable:BAD:start -->\n", "is not repeatable"),
            ("" + marker + "remove:end -->\n", "unmatched template block end"),
            ("" + marker + "remove:start -->\n", "unclosed template block"),
        )
        for text, message in cases:
            with self.subTest(message=message):
                with self.assertRaisesRegex(initializer.InitializationError, message):
                    initializer._render_blocks("README.md", text, "library", {}, {}, catalogue)

    def test_replacement_badges_changelog_and_helpers(self) -> None:
        catalogue: dict[str, dict[str, Any]] = {
            "PROFILE": {"category": "profile-specific", "values": {"library": "lib"}},
            "REP": {"category": "repeatable", "item_format": "- {value}"},
        }
        self.assertEqual(
            initializer._replacement_values("library", {"A": "x"}, {"REP": ["one", "two"]}, catalogue),
            {"A": "x", "PROFILE": "lib", "REP": "- one\n- two"},
        )
        readme = (
            "https://github.com/owner/template/actions "
            "https://img.shields.io/github/v/release/owner/template?x "
            "https://img.shields.io/github/issues/owner/template/P2?x "
            "https://api.scorecard.dev/projects/github.com/owner/template/badge "
            "https://scorecard.dev/viewer/?uri=github.com/owner/template "
            '<img src="assets/social-preview.png">\n'
            "<!-- generated-social-preview: assets/custom-preview.png -->\n"
        )
        retargeted = initializer._render_readme_badges(
            "README.md", readme, "owner/template", "owner/product"
        )
        self.assertNotIn("owner/template", retargeted)
        self.assertIn("owner/product", retargeted)
        self.assertIn('src="assets/custom-preview.png"', retargeted)
        self.assertNotIn("generated-social-preview", retargeted)
        self.assertEqual(
            initializer._render_readme_badges("OTHER.md", readme, "owner/template", "owner/product"),
            readme,
        )
        changelog = "# C\n\n## [Unreleased]\n\nold\n\n---\nrest\n"
        reset = initializer._reset_changelog(changelog)
        self.assertIn("No unreleased changes recorded.", reset)
        for bad in ("# C\n", "## [Unreleased]\nno boundary"):
            with self.assertRaises(initializer.InitializationError):
                initializer._reset_changelog(bad)
        self.assertIsNone(initializer._sha256(None))
        self.assertEqual(len(initializer._sha256(b"x") or ""), 64)

    def test_already_initialized_and_placeholder_protection(self) -> None:
        config = {"mode": "template"}
        self.assertFalse(initializer._already_initialized(Path("."), config, "library", {}, {}))
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            generated: dict[str, Any] = {
                "mode": "generated",
                "profile": "library",
                "repository": "owner/repo",
                "template_contract": {"version": "1.2.0", "source": "owner/template"},
            }
            scalars = {"REPOSITORY_PATH": "owner/repo"}
            with self.assertRaisesRegex(initializer.InitializationError, "missing"):
                initializer._already_initialized(root, generated, "library", scalars, {})
            (root / ".github").mkdir()
            (root / initializer.RECORD_PATH).write_bytes(
                initializer._record("library", scalars, {}, generated["template_contract"])
            )
            self.assertTrue(initializer._already_initialized(root, generated, "library", scalars, {}))
            with self.assertRaisesRegex(initializer.InitializationError, "different profile"):
                initializer._already_initialized(root, generated, "application", scalars, {})
            (root / initializer.RECORD_PATH).write_bytes(b"{}")
            with self.assertRaisesRegex(initializer.InitializationError, "different substitution"):
                initializer._already_initialized(root, generated, "library", scalars, {})
        with self.assertRaises(initializer.InitializationError):
            initializer._reject_executable_placeholders("tool.py", [object()])
        initializer._reject_executable_placeholders("README.md", [object()])

    def test_strip_workflow_plan_and_apply_rollback(self) -> None:
        workflow = (
            "before\n"
            "      - name: Exercise policy-branch coverage determinism\n"
            "        run: hidden\n"
            "      - name: Exercise positive and degraded checker paths\n"
            "keep\n"
            "test-results/policy-coverage.json\n"
            "POLICY_COVERAGE_OUTCOME=x\n"
            "\"Policy coverage:$X\"\n"
        )
        stripped = initializer._strip_template_maintenance_workflow(
            ".github/workflows/static-checks.yml", workflow
        )
        self.assertEqual(stripped, "before\n      - name: Exercise positive and degraded checker paths\nkeep\n")
        self.assertEqual(initializer._strip_template_maintenance_workflow("x.yml", workflow), workflow)
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            (root / "old.txt").write_bytes(b"old")
            plan = initializer._plan(
                root, "library", {"old.txt": b"new", "new.txt": b"x", "gone.txt": None}
            )
            self.assertEqual([row["action"] for row in plan["changes"]], ["update", "create", "delete"])
            initializer._apply_changes(root, {"old.txt": b"new", "new.txt": b"x"})
            self.assertEqual((root / "old.txt").read_bytes(), b"new")
            self.assertEqual((root / "new.txt").read_bytes(), b"x")
            before = (root / "old.txt").read_bytes()
            with patch.object(initializer.os, "replace", side_effect=OSError("boom")):
                with self.assertRaisesRegex(initializer.InitializationError, "original files were restored"):
                    initializer._apply_changes(root, {"old.txt": b"broken"})
            self.assertEqual((root / "old.txt").read_bytes(), before)


class ReleaseDepthTests(unittest.TestCase):
    def base_policy(self) -> dict:
        profile = {"required_checks": ["vba-compile", "regression"], "allowed_asset_globs": []}
        return {
            "schema_version": 1,
            "evidence_schema_version": 1,
            "provenance_signature_mode": "none",
            "core_checks": ["repository-integrity"],
            "profiles": {
                "application": dict(profile),
                "library": dict(profile),
                "template": dict(profile),
                "ui-component": dict(profile),
            },
            "source_scan_exclude_paths": ["README.md"],
            "template_construction_markers": ["template construction"],
        }

    def test_json_configuration_and_policy_failures(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            (root / ".github").mkdir()
            config_path = root / release.PROFILE_PATH
            value, findings = release._read_json(config_path, "repository-profile")
            self.assertIsNone(value)
            self.assertTrue(findings)
            config_path.write_text("{", encoding="utf-8")
            self.assertTrue(release._read_json(config_path, "repository-profile")[1])
            config_path.write_text("[]", encoding="utf-8")
            self.assertIsNone(release._load_configuration(root)[0])
            config_path.write_text("{}", encoding="utf-8")
            self.assertEqual(release._load_configuration(root), ({}, []))

            policy_path = root / release.POLICY_PATH
            good = self.base_policy()
            policy_path.write_text(json.dumps(good), encoding="utf-8")
            self.assertIsNotNone(release._load_policy(root)[0])
            variants = []
            value = dict(good)
            value.pop("schema_version")
            variants.append(value)
            value = dict(good)
            value["schema_version"] = 2
            variants.append(value)
            value = dict(good)
            value["provenance_signature_mode"] = "bad"
            variants.append(value)
            value = dict(good)
            value["core_checks"] = []
            variants.append(value)
            value = dict(good)
            value["core_checks"] = ["x", "x"]
            variants.append(value)
            value = dict(good)
            value["profiles"] = {}
            variants.append(value)
            value = json.loads(json.dumps(good))
            value["profiles"]["library"]["extra"] = 1
            variants.append(value)
            value = json.loads(json.dumps(good))
            value["profiles"]["library"]["required_checks"] = []
            variants.append(value)
            value = json.loads(json.dumps(good))
            value["profiles"]["library"]["required_checks"] = ["repository-integrity"]
            variants.append(value)
            value = json.loads(json.dumps(good))
            value["profiles"]["library"]["allowed_asset_globs"] = ["../bad"]
            variants.append(value)
            value = dict(good)
            value["source_scan_exclude_paths"] = ["../bad"]
            variants.append(value)
            value = dict(good)
            value["template_construction_markers"] = []
            variants.append(value)
            for variant in variants:
                with self.subTest(variant=variant):
                    policy_path.write_text(json.dumps(variant), encoding="utf-8")
                    self.assertIsNone(release._load_policy(root)[0])

    def test_safe_relative_and_git_operational_paths(self) -> None:
        for value, expected in (
            ("a/b", True), ("", False), ("/a", False), ("../a", False),
            ("a\\b", False), ("a\0b", False), (1, False),
        ):
            self.assertEqual(release._safe_relative(value), expected)
        with patch.object(release.subprocess, "run", side_effect=FileNotFoundError()):
            with self.assertRaises(release.OperationalError):
                release._git(Path("."), "status")
        with patch.object(release, "_git_output", return_value=None):
            with self.assertRaises(release.OperationalError):
                release._tracked_files(Path("."))

    def test_resolve_release_profile_modes(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            (root / ".github").mkdir()
            generated = {"mode": "generated", "profile": "library", "repository": "owner/product"}
            profile, findings = release._resolve_release_profile(root, generated)
            self.assertEqual(profile, "library")
            self.assertTrue(findings)

            record = {"profile": "application", "values": {"REPOSITORY_PATH": "owner/other"}}
            (root / release.INITIALIZATION_PATH).write_text(json.dumps(record), encoding="utf-8")
            _, findings = release._resolve_release_profile(root, generated)
            self.assertGreaterEqual(len(findings), 2)

            template = {
                "mode": "template", "profile": None, "repository": "owner/product",
                "identity": {"template_tokens": ["TEMPLATE"]},
            }
            _, findings = release._resolve_release_profile(root, template)
            self.assertGreaterEqual(len(findings), 2)

            profile, findings = release._resolve_release_profile(root, {"mode": "other"})
            self.assertIsNone(profile)
            self.assertTrue(findings)

    def test_candidate_git_state_variants(self) -> None:
        with patch.object(release, "_git_output", side_effect=["a" * 40, " M x"]):
            findings = release._validate_candidate_git_state(Path("."), "b" * 40, "v1.0.0", False)
            self.assertEqual({row["code"] for row in findings}, {"candidate-sha-mismatch", "dirty-candidate"})
        with patch.object(release, "_git_output", side_effect=["a" * 40, "", None, None]):
            findings = release._validate_candidate_git_state(Path("."), "a" * 40, "v1.0.0", True)
            self.assertEqual(findings[0]["code"], "missing-tag-ref")
        with patch.object(release, "_git_output", side_effect=["a" * 40, "", "commit", "b" * 40]):
            codes = {row["code"] for row in release._validate_candidate_git_state(Path("."), "a" * 40, "v1.0.0", True)}
            self.assertEqual(codes, {"lightweight-tag", "tag-target-mismatch"})

    def test_version_changelog_failures(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            version, findings = release._validate_version_and_changelog(root, "v1.0.0")
            self.assertIsNone(version)
            self.assertTrue(findings)
            (root / "VERSION").write_text("0.0.0\n", encoding="utf-8")
            (root / "CHANGELOG.md").write_text("## [0.0.0] - 2026-02-30\n", encoding="utf-8")
            _, findings = release._validate_version_and_changelog(root, "v9.9.9")
            codes = {row["code"] for row in findings}
            self.assertTrue({"zero-version", "tag-version-mismatch", "invalid-changelog-date"} <= codes)
            (root / "VERSION").write_text("1.0.0 extra\n", encoding="utf-8")
            (root / "CHANGELOG.md").write_text("# no release\n", encoding="utf-8")
            _, findings = release._validate_version_and_changelog(root, "v1.0.0")
            codes = {row["code"] for row in findings}
            self.assertTrue({"invalid-version", "tag-version-mismatch", "missing-changelog-release"} <= codes)

    def test_validate_check_field_failures(self) -> None:
        sha = "a" * 40
        self.assertEqual(release._validate_check("x", None, sha)[0]["code"], "invalid-evidence-check")
        self.assertTrue(release._validate_check("repository-integrity", {}, sha))
        base = {"status": "FAIL", "candidate_sha": "b" * 40, "detail": ""}
        findings = release._validate_check("x", base, sha)
        self.assertEqual({row["code"] for row in findings}, {"failed-evidence-check", "evidence-sha-mismatch", "invalid-evidence-check"})
        repo = {**base, "status": "PASS", "candidate_sha": sha, "detail": "x", "run_url": "http://x"}
        self.assertTrue(release._validate_check("repository-integrity", repo, sha))
        compile_check = {**base, "status": "PASS", "candidate_sha": sha, "detail": "x", "environment": ""}
        self.assertTrue(release._validate_check("vba-compile", compile_check, sha))
        regression = {
            "status": "PASS", "candidate_sha": sha, "detail": "x", "entry_point": "",
            "environment": "", "cases": True, "assertions": 0, "failures": 1,
            "completeness": "PARTIAL", "cleanup": "FAIL",
        }
        self.assertGreaterEqual(len(release._validate_check("regression", regression, sha)), 7)

    def test_manifest_paths_and_asset_validation(self) -> None:
        import hashlib
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            self.assertTrue(release._parse_manifest(root / "missing.txt")[1])
            manifest = root / "manifest.txt"
            manifest.write_text(
                "bad\n" + "a" * 64 + "  z.bin\n" + "b" * 64 + "  a.bin\n" + "c" * 64 + "  a.bin\n",
                encoding="utf-8",
            )
            _, findings = release._parse_manifest(manifest)
            self.assertGreaterEqual(len(findings), 3)
            asset_file = root / "dist.bin"
            asset_file.write_bytes(b"x")
            sha = "a" * 40
            self.assertTrue(release._validate_asset_record(root, {}, "asset", ["*.bin"], sha, {}))
            malformed = {"path": "../x", "sha256": "x", "candidate_sha": sha, "package_test": "PASS"}
            self.assertTrue(release._validate_asset_record(root, malformed, "asset", ["*.bin"], sha, {}))
            digest = hashlib.sha256(b"wrong").hexdigest()
            record = {"path": "dist.bin", "sha256": digest, "candidate_sha": "b" * 40, "package_test": "FAIL"}
            findings = release._validate_asset_record(root, record, "asset", ["*.zip"], sha, {})
            codes = {row["code"] for row in findings}
            self.assertTrue({"asset-sha-binding-mismatch", "failed-package-test", "unapproved-binary", "asset-digest-mismatch"} <= codes)

    def test_evidence_metadata_checks_and_binary_policy(self) -> None:
        policy = self.base_policy()
        sha = "a" * 40
        evidence = {
            "schema_version": 9, "version": "9.9.9", "tag": "bad", "candidate_sha": "b" * 40,
            "profile": "other", "distribution": "bad", "checks": {}, "assets": [], "extra": 1,
        }
        findings = release._validate_evidence_metadata(evidence, Path("e.json"), policy, "library", "1.0.0", "v1.0.0", sha)
        self.assertGreaterEqual(len(findings), 6)
        self.assertTrue(release._validate_evidence_checks({"checks": []}, policy, "library", sha))
        self.assertTrue(release._validate_evidence_checks({"checks": {"BAD ID": {}}}, policy, "library", sha))
        findings = release._validate_binary_assets(Path("."), [], None, None, policy, "library", sha)
        codes = {row["code"] for row in findings}
        self.assertTrue({"unapproved-binary", "missing-release-assets", "missing-asset-manifest"} <= codes)

    def test_reports_and_atomic_write(self) -> None:
        report = {
            "status": "fail", "tag": "v1", "candidate_sha": "a" * 40, "profile": "library",
            "counts": {"findings": 1},
            "findings": [{"code": "x", "path": "a", "message": "pipe | here"}],
            "scope_note": "scope",
        }
        self.assertIn("[FAIL]", release.console_report(report))
        self.assertIn("\\|", release.markdown_report(report))
        with tempfile.TemporaryDirectory() as name:
            target = Path(name) / "nested" / "report.txt"
            release._write_atomic(target, "hello")
            self.assertEqual(target.read_text(encoding="utf-8"), "hello")


class CloseoutDepthTests(unittest.TestCase):
    def test_type_helpers_and_json_errors(self) -> None:
        with self.assertRaises(closeout.CloseoutError):
            closeout.require(False, "x")
        for func, value in ((closeout.as_object, []), (closeout.as_list, {}), (closeout.as_string, "")):
            with self.subTest(func=func.__name__):
                with self.assertRaises(closeout.CloseoutError):
                    func(value, "x")
        self.assertEqual(closeout.unique_object([("a", 1)]), {"a": 1})
        with self.assertRaises(closeout.CloseoutError):
            closeout.unique_object([("a", 1), ("a", 2)])
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            with self.assertRaises(closeout.CloseoutError):
                closeout.load_json(root / "missing.json")
            path = root / "x.json"
            path.write_text('{"a":1,"a":2}', encoding="utf-8")
            with self.assertRaises(closeout.CloseoutError):
                closeout.load_json(path)
            path.write_text("{", encoding="utf-8")
            with self.assertRaises(closeout.CloseoutError):
                closeout.load_json(path)

    def test_workflow_runs_and_milestone_items(self) -> None:
        rows, total = closeout.workflow_runs([
            {"workflow_runs": [{"id": 1}], "total_count": 2},
            {"workflow_runs": [{"id": 2}], "total_count": 3},
        ])
        self.assertEqual((len(rows), total), (2, 3))
        self.assertEqual(closeout.milestone_items([[{"number": 1}], [{"number": 2}], ["bad"]]), [{"number": 1}, {"number": 2}])
        self.assertEqual(closeout.milestone_items([{"number": 1}, "bad"]), [{"number": 1}])

    def test_tag_variants(self) -> None:
        candidate = "a" * 40
        findings: list[dict[str, str]] = []
        state = closeout.check_tag({"tag_ref": {"ref": "bad", "object": {"type": "commit", "sha": "b" * 40}}}, "v1.0.0", candidate, findings)
        self.assertEqual(state["ref_type"], "commit")
        self.assertGreaterEqual(len(findings), 3)
        findings = []
        closeout.check_tag({
            "tag_ref": {"ref": "refs/tags/v1.0.0", "object": {"type": "tag", "sha": "t"}},
            "tag_object": {"sha": "wrong", "tag": "v2", "object": {"type": "blob", "sha": candidate}},
        }, "v1.0.0", candidate, findings)
        self.assertGreaterEqual(len(findings), 3)
        findings = []
        closeout.check_tag({"tag_ref": {"ref": "refs/tags/v1.0.0", "object": {"type": "tree"}}}, "v1.0.0", candidate, findings)
        self.assertTrue(findings)

    def test_workflow_selection_and_failure(self) -> None:
        candidate = "a" * 40
        findings: list[dict[str, str]] = []
        result = closeout.check_workflow({"workflow_runs": {"workflow_runs": [], "total_count": 1}}, "x.yml", "v1", candidate, findings)
        self.assertEqual(result["matches"], 0)
        self.assertGreaterEqual(len(findings), 2)
        rows = [
            {"id": 1, "path": "x.yml", "head_branch": "v1", "head_sha": candidate, "event": "push", "run_attempt": 1, "run_number": 1, "status": "completed", "conclusion": "failure"},
            {"id": 2, "path": "x.yml", "head_branch": "v1", "head_sha": candidate, "event": "push", "run_attempt": 2, "run_number": 2, "status": "completed", "conclusion": "success"},
        ]
        findings = []
        result = closeout.check_workflow({"workflow_runs": {"workflow_runs": rows, "total_count": 2}}, "x.yml", "v1", candidate, findings)
        self.assertEqual(result["run_id"], 2)
        self.assertEqual(findings, [])

    def test_release_asset_archive_and_latest_failures(self) -> None:
        identity = {"version": "1.0.0", "allowed_asset_globs": []}
        snapshot = {
            "release": {"id": 1, "tag_name": "bad", "draft": True, "published_at": None, "prerelease": True,
                        "assets": [{"name": "x.bin"}, {"name": "x.bin"}, "bad"], "zipball_url": None, "tarball_url": ""},
            "latest_release": {"id": 2},
            "source_archives": {"zip": "fail", "tar": "missing"},
        }
        findings: list[dict[str, str]] = []
        result = closeout.check_release(snapshot, identity, False, True, findings)
        self.assertEqual(result["asset_mode"], "source-only")
        self.assertTrue({"release", "assets", "source-archives"} <= {row["control"] for row in findings})

    def test_compare_milestone_and_wiki_failures(self) -> None:
        findings: list[dict[str, str]] = []
        result = closeout.check_compare({"compare": {}}, {"version": "1.0.0", "previous_tag": None}, "a" * 40, findings)
        self.assertIsNone(result["previous_tag"])
        self.assertEqual(findings, [])
        findings = []
        closeout.check_compare({"compare": {"html_url": "bad", "status": "behind", "commits": []}}, {"version": "1.0.0", "previous_tag": "v0.9.0"}, "a" * 40, findings)
        self.assertGreaterEqual(len(findings), 2)
        findings = []
        result = closeout.check_milestone({
            "milestone": {"number": 9, "state": "open", "open_issues": 0, "closed_issues": 0, "title": "m"},
            "milestone_items": [
                {"number": 1, "state": "open", "milestone": {"number": 8}},
                {"number": 2, "state": "mystery", "milestone": {"number": 9}},
            ],
        }, 4, findings)
        self.assertEqual(result["open_items"], [1])
        self.assertGreaterEqual(len(findings), 5)
        self.assertFalse(closeout.check_wiki({}, {"mode": "generated"}, "a" * 40, [])["applicable"])
        findings = []
        result = closeout.check_wiki({
            "wiki": {"status": "fail", "source_sha": "b" * 40},
            "ui_observations": {"wiki_browser_review": "fail"},
        }, {"mode": "template"}, "a" * 40, findings)
        self.assertTrue(result["applicable"])
        self.assertEqual(len(findings), 3)


class ContractAndWikiDepthTests(unittest.TestCase):
    def test_contract_read_shape_notes_and_record_errors(self) -> None:
        findings: list[dict[str, str]] = []
        self.assertEqual(contract._check_shape([], findings), ("", ""))
        self.assertTrue(findings)
        findings = []
        version, source = contract._check_shape({"version": "1.2.0-rc.1", "source": "bad", "extra": 1}, findings)
        self.assertEqual((version, source), ("", ""))
        self.assertGreaterEqual(len(findings), 3)
        findings = []
        self.assertEqual(contract._check_supported("", findings), frozenset())
        self.assertEqual(contract._check_supported("9.9.9", findings), frozenset())
        self.assertTrue(findings)
        findings = []
        contract._check_mode({"mode": "template", "repository": "owner/a"}, "owner/b", findings)
        contract._check_mode({"mode": "generated", "repository": "owner/a"}, "owner/a", findings)
        self.assertEqual(len(findings), 2)
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            with self.assertRaises(contract.ContractError):
                contract._read_json(root, "missing.json")
            (root / "docs").mkdir()
            (root / contract.NOTES_PATH).write_text("### 1.0.0\n", encoding="utf-8")
            findings = []
            contract._check_notes(root, findings)
            self.assertGreaterEqual(len(findings), 2)
            (root / ".github").mkdir()
            (root / contract.RECORD_PATH).write_text("[]", encoding="utf-8")
            findings = []
            self.assertIsNone(contract._check_record(root, {"mode": "generated"}, "1.2.0", "owner/template", findings))
            self.assertTrue(findings)

    def test_wiki_helpers_failure_boundaries(self) -> None:
        self.assertEqual(wiki.directories({"a/b/c.txt"}), {"a", "a/b"})
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            path = root / "x.json"
            path.write_text('{"a":1,"a":2}', encoding="utf-8")
            with self.assertRaises(ValueError):
                wiki.read_json(path)
            path.write_text("[]", encoding="utf-8")
            with self.assertRaises(ValueError):
                wiki.read_json(path)
            with self.assertRaises(ValueError):
                wiki.bundle(root, {"pages": []}, "bad")
            with self.assertRaises(ValueError):
                wiki.compare_bundle({}, root / "missing")
            outside = Path(name).parent / (Path(name).name + "-export")
            try:
                wiki.write_bundle({"x": b"1"}, outside, root)
                self.assertEqual((outside / "x").read_bytes(), b"1")
                with self.assertRaises(ValueError):
                    wiki.write_bundle({}, outside, root)
            finally:
                if outside.exists():
                    for item in outside.iterdir():
                        item.unlink()
                    outside.rmdir()

    def test_wiki_render_and_compare(self) -> None:
        text = "[remote](https://example.com) [anchor](#x) [page](Other.md#y) [file](../../README.md)"
        rendered = wiki.render_page(text, "Home.md", {"Home", "Other"}, "owner/repo", "a" * 40)
        self.assertIn("](Other#y)", rendered)
        self.assertIn("/blob/" + "a" * 40, rendered)
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            (root / "x").write_bytes(b"bad")
            differences = wiki.compare_bundle({"x": b"good", "y": b"new"}, root)
            self.assertEqual(len(differences), 2)


class ReleaseSemanticsDepthTests(unittest.TestCase):
    def valid_exception(self) -> dict[str, Any]:
        return {
            "base_tag": "v1.2.0",
            "commit": "a" * 40,
            "findings": ["merge-commit"],
            "review_ref": "https://github.com/owner/repo/issues/1",
            "reason": "Reviewed exception.",
        }

    def valid_record(self) -> dict[str, Any]:
        return {
            "release": "v1.2.0",
            "commit": "b" * 40,
            "pull_request": 7,
            "reason": "Historical record.",
        }

    def test_history_record_shape_validation(self) -> None:
        valid = self.valid_exception()
        variants: list[object] = [
            [],
            {},
            {**valid, "base_tag": "bad"},
            {**valid, "commit": "bad"},
            {**valid, "findings": []},
            {**valid, "findings": ["merge-commit", "merge-commit"]},
            {**valid, "review_ref": "bad"},
            {**valid, "reason": "   "},
        ]
        for value in variants:
            with self.subTest(value=value):
                with self.assertRaises(ValueError):
                    semantics._history_exception(value)
        self.assertEqual(semantics._history_exception(valid), valid)

        record = self.valid_record()
        record_variants: list[object] = [
            [],
            {},
            {**record, "release": "bad"},
            {**record, "commit": "bad"},
            {**record, "pull_request": True},
            {**record, "pull_request": 0},
            {**record, "reason": ""},
        ]
        for value in record_variants:
            with self.subTest(value=value):
                with self.assertRaises(ValueError):
                    semantics._historical_record(value)
        self.assertEqual(semantics._historical_record(record), record)

    def test_history_policy_loader_rejects_drift(self) -> None:
        valid = {
            "schema_version": 1,
            "rules": {
                "merge_commits": "block-unless-excepted",
                "duplicate_subjects": "block-unless-excepted",
            },
            "exceptions": [self.valid_exception()],
            "historical_records": [self.valid_record()],
        }
        variants: list[dict[str, Any]] = [
            {"schema_version": 1},
            {**valid, "schema_version": 2},
            {**valid, "rules": {}},
            {**valid, "exceptions": {}},
            {**valid, "historical_records": {}},
        ]
        for value in variants:
            with self.subTest(value=value):
                with patch.object(semantics, "_candidate_json", return_value=value):
                    with self.assertRaises(ValueError):
                        semantics._load_history_policy(Path("."), "c" * 40)
        with patch.object(semantics, "_candidate_json", return_value=valid):
            loaded = semantics._load_history_policy(Path("."), "c" * 40)
        self.assertEqual(loaded["exceptions"][0]["commit"], "a" * 40)
        self.assertEqual(loaded["historical_records"][0]["pull_request"], 7)

    def test_history_candidate_rows_and_exception_semantics(self) -> None:
        checkout = "c" * 40
        base = "b" * 40
        head = "a" * 40
        with patch.object(semantics, "_git_output", return_value=base):
            self.assertEqual(semantics._effective_history_candidate(Path("."), checkout), checkout)
        with patch.object(semantics, "_git_output", side_effect=[f"{base} {head}", "ordinary"]):
            self.assertEqual(semantics._effective_history_candidate(Path("."), checkout), checkout)
        subject = f"Merge {head} into {base}"
        with patch.object(semantics, "_git_output", side_effect=[f"{head} {base}", subject]):
            self.assertEqual(semantics._effective_history_candidate(Path("."), checkout), checkout)
        with patch.object(semantics, "_git_output", side_effect=[f"{base} {head}", subject]):
            self.assertEqual(semantics._effective_history_candidate(Path("."), checkout), head)

        with patch.object(semantics, "_git_output", side_effect=["parent", "not-a-tag"]):
            with self.assertRaises(ValueError):
                semantics._previous_release_tag(Path("."), head)
        with patch.object(semantics, "_git_output", side_effect=["parent", "v1.2.0"]):
            self.assertEqual(semantics._previous_release_tag(Path("."), head), "v1.2.0")

        raw = f"{head}\x1f{base}\x1fSubject\x1e"
        with patch.object(semantics, "_git_output", return_value=raw):
            rows = semantics._history_rows(Path("."), "v1.2.0", head)
        self.assertEqual(rows[0]["subject"], "Subject")
        with patch.object(semantics, "_git_output", return_value="x\x1fy\x1e"):
            with self.assertRaises(ValueError):
                semantics._history_rows(Path("."), "v1.2.0", head)

        first = "1" * 40
        second = "2" * 40
        rows = [
            {"commit": first, "parents": ["0" * 40], "subject": "Same"},
            {"commit": second, "parents": ["0" * 40, "f" * 40], "subject": "same"},
        ]
        conditions = semantics._history_conditions(rows)
        self.assertEqual(conditions[second], {"merge-commit", "duplicate-subject"})
        policy = {
            "exceptions": [
                {
                    "base_tag": "v1.2.0",
                    "commit": second,
                    "findings": ["merge-commit"],
                    "review_ref": "https://github.com/owner/repo/issues/1",
                    "reason": "reviewed",
                },
                {
                    "base_tag": "v1.2.0",
                    "commit": first,
                    "findings": ["merge-commit"],
                    "review_ref": "https://github.com/owner/repo/issues/2",
                    "reason": "overbroad",
                },
                {
                    "base_tag": "v1.2.0",
                    "commit": "3" * 40,
                    "findings": ["duplicate-subject"],
                    "review_ref": "https://github.com/owner/repo/issues/3",
                    "reason": "stale",
                },
            ]
        }
        findings, used = semantics._history_findings(rows, policy, "v1.2.0")
        self.assertEqual(used, [{"commit": second, "finding": "merge-commit"}])
        self.assertEqual(
            {item["code"] for item in findings},
            {"duplicate-commit-subject", "overbroad-history-exception", "stale-history-exception"},
        )

    def test_history_report_and_markdown_paths(self) -> None:
        checkout = "c" * 40
        with patch.object(semantics, "_git_output", return_value=checkout):
            report = semantics._release_history_report(
                Path("."), {"mode": "generated", "profile": "library"}
            )
        self.assertFalse(report["applicable"])

        with (
            patch.object(semantics, "_git_output", return_value=checkout),
            patch.object(semantics, "_effective_history_candidate", return_value="a" * 40),
            patch.object(
                semantics,
                "_load_history_policy",
                return_value={"exceptions": [], "historical_records": [self.valid_record()]},
            ),
            patch.object(semantics, "_previous_release_tag", return_value="v1.2.0"),
            patch.object(semantics, "_history_rows", return_value=[]),
            patch.object(semantics, "_history_findings", return_value=([], [])),
        ):
            report = semantics._release_history_report(
                Path("."), {"mode": "template", "profile": None}
            )
        self.assertTrue(report["applicable"])
        self.assertEqual(report["candidate_sha"], "a" * 40)

        rendered = semantics.markdown_report(
            {
                "status": "fail",
                "version": "1.3.0",
                "releases": [{"version": "1.3.0", "date": "2026-09-12", "line": 7}],
                "findings": [{"code": "x", "path": "CHANGELOG.md", "line": 7, "message": "bad"}],
                "history_policy": {
                    "applicable": True,
                    "previous_tag": "v1.2.0",
                    "candidate_sha": "a" * 40,
                    "checkout_sha": checkout,
                    "commits": [],
                    "exceptions_used": [],
                },
            }
        )
        self.assertIn("Synthetic PR merge checkout", rendered)
        self.assertIn("`x`", rendered)


class ExtendedReleaseAndCloseoutDepthTests(unittest.TestCase):
    def test_release_operational_source_and_cli_paths(self) -> None:
        with patch.object(Path, "read_text", side_effect=OSError("denied")):
            with self.assertRaises(release.OperationalError):
                release._read_json(Path("x.json"), "x")

        sha = "a" * 40
        with patch.object(release, "_git_output", return_value=None):
            with self.assertRaises(release.OperationalError):
                release._validate_candidate_git_state(Path("."), sha, "v1.0.0", False)
        with patch.object(release, "_git_output", side_effect=[sha, None]):
            with self.assertRaises(release.OperationalError):
                release._validate_candidate_git_state(Path("."), sha, "v1.0.0", False)

        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            placeholder = "{" + "{TOKEN}" + "}"
            (root / "token.txt").write_text(f"{placeholder} TemplateIdentity\n", encoding="utf-8")
            (root / "bad.txt").write_bytes(b"\xff")
            configuration = {
                "identity": {"exclude_paths": [], "template_tokens": ["TemplateIdentity"]}
            }
            policy = {
                "source_scan_exclude_paths": [],
                "template_construction_markers": ["construction marker"],
            }
            with patch.object(release, "_tracked_files", return_value=["token.txt", "bad.txt"]):
                findings = release._validate_generated_source(root, configuration, policy)
            codes = {item["code"] for item in findings}
            self.assertTrue(
                {"unresolved-template-token", "template-identity", "unreadable-release-text", "missing-changelog"}
                <= codes
            )
            (root / "CHANGELOG.md").write_text("construction marker\n", encoding="utf-8")
            with patch.object(release, "_tracked_files", return_value=[]):
                findings = release._validate_generated_source(root, configuration, policy)
            self.assertIn("template-construction-history", {item["code"] for item in findings})

            evidence = root / "evidence.json"
            evidence.write_text("{}", encoding="utf-8")
            output = root / "report.json"
            summary = root / "summary.md"
            args = [
                "--root", str(root), "--tag", "v1.2.0", "--candidate-sha", sha,
                "--evidence", str(evidence), "--output", str(output), "--summary", str(summary),
            ]
            with (
                patch.object(release, "build_report", return_value={"status": "pass"}),
                patch.object(release, "console_report", return_value="PASS"),
                patch.object(release, "markdown_report", return_value="summary\n"),
            ):
                self.assertEqual(release.main(args), 0)
            self.assertTrue(output.is_file())
            self.assertEqual(summary.read_text(encoding="utf-8"), "summary\n")
            with patch.object(release, "build_report", side_effect=release.OperationalError("boom")):
                with self.assertRaises(SystemExit):
                    release.main(args[:6])
            with patch.object(release, "_run_self_test", return_value=7):
                self.assertEqual(release.main(["--root", str(root), "--self-test"]), 7)

    def test_closeout_markdown_and_write_paths(self) -> None:
        candidate = "a" * 40
        report: dict[str, Any] = {
            "status": "fail",
            "deterministic_status": "pass",
            "observation_status": "fail",
            "tag": "v1.2.1",
            "candidate_sha": candidate,
            "profile": "template",
            "snapshot_sha256": "f" * 64,
            "tag_state": {"ref_type": "tag", "target_sha": candidate},
            "tag_workflow": {"conclusion": "success", "run_id": 11},
            "release": {
                "draft": False,
                "latest_expected": True,
                "latest_matches": True,
                "release_id": 22,
                "unexpected_assets": [],
                "asset_mode": "source-only",
                "uploaded_assets": [],
            },
            "comparison": {"status": "ahead", "previous_tag": "v1.2.0"},
            "milestone": {"open_items": [], "state": "closed", "membership_count": 8},
            "wiki": {"applicable": True, "status": "fail", "source_sha": "b" * 40},
            "findings": [
                {"category": "observation", "control": "wiki", "message": "read-back failed"}
            ],
            "scope_note": "scope",
        }
        rendered = closeout.markdown_report(report)
        self.assertIn("Wiki source/read-back", rendered)
        self.assertIn("read-back failed", rendered)
        self.assertIn("| Uploaded assets | PASS | source-only; 0 uploaded |", rendered)
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            text_path = root / "nested" / "report.md"
            json_path = root / "nested" / "report.json"
            closeout.write_report(text_path, rendered)
            closeout.write_json(json_path, {"status": "pass"})
            self.assertEqual(text_path.read_text(encoding="utf-8"), rendered)
            self.assertEqual(json.loads(json_path.read_text(encoding="utf-8"))["status"], "pass")

    def test_initializer_main_failure_and_apply_paths(self) -> None:
        self.assertEqual(initializer.main([]), 2)
        bad_git = type("Result", (), {"returncode": 1, "stdout": "", "stderr": "bad"})()
        with patch.object(initializer, "_git", return_value=bad_git):
            self.assertEqual(initializer.main(["--profile", "library"]), 2)

        dirty = type("Result", (), {"returncode": 0, "stdout": " M README.md", "stderr": ""})()
        with (
            patch.object(initializer, "_git", return_value=dirty),
            patch.object(initializer, "_load_config", return_value={"mode": "template"}),
        ):
            self.assertEqual(initializer.main(["--profile", "library"]), 2)

        clean = type("Result", (), {"returncode": 0, "stdout": "", "stderr": ""})()
        with (
            patch.object(initializer, "_git", return_value=clean),
            patch.object(initializer, "_load_config", return_value={"mode": "generated"}),
            patch.object(initializer, "_build_changes", return_value=({"x.txt": b"x"}, {})),
            patch.object(initializer, "_plan", return_value={"mode": "dry-run", "status": "planned", "changes": []}),
            patch.object(initializer, "_apply_changes") as apply_changes,
        ):
            self.assertEqual(initializer.main(["--profile", "library", "--apply"]), 0)
            apply_changes.assert_called_once()

    def test_initializer_generated_contract_failure_helpers(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            with self.assertRaises(AssertionError):
                initializer._record_arguments(root)
            (root / ".github").mkdir()
            (root / initializer.RECORD_PATH).write_text(
                json.dumps(
                    {
                        "schema_version": 1,
                        "profile": "library",
                        "template_contract": {},
                        "values": {"BAD": 1},
                    }
                ),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(AssertionError, "invalid value"):
                initializer._record_arguments(root)

            bad_marker = "<" + "!-- template:oops --" + ">"
            (root / "x.md").write_text(bad_marker + "\n", encoding="utf-8")
            with patch.object(initializer, "_tracked_files", return_value=["x.md"]):
                with self.assertRaisesRegex(AssertionError, "retained a template marker"):
                    initializer._assert_generated_cleanup(root, "library")

            (root / "x.md").write_text("clean\n", encoding="utf-8")
            issue = root / ".github/ISSUE_TEMPLATE/config.yml"
            issue.parent.mkdir(parents=True, exist_ok=True)
            issue.write_text("blank\n", encoding="utf-8")
            with patch.object(initializer, "_tracked_files", return_value=["x.md"]):
                with self.assertRaisesRegex(AssertionError, "private-security URL"):
                    initializer._assert_generated_cleanup(root, "library", "owner/repo")

            config: dict[str, Any] = {
                "repository": "owner/product",
                "template_contract": "bad",
            }
            with self.assertRaisesRegex(AssertionError, "well-formed template_contract"):
                initializer._assert_adopted_contract(root, config)

            config["template_contract"] = {"version": "1.2.0", "source": "owner/template"}
            (root / initializer.RECORD_PATH).write_text(
                json.dumps({"template_contract": {"version": "1.1.0", "source": "owner/template"}}),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(AssertionError, "disagree"):
                initializer._assert_adopted_contract(root, config)

            config["template_contract"] = {"version": "1.2.0", "source": "owner/product"}
            (root / initializer.RECORD_PATH).write_text(
                json.dumps({"template_contract": config["template_contract"]}), encoding="utf-8"
            )
            with self.assertRaisesRegex(AssertionError, "own template contract source"):
                initializer._assert_adopted_contract(root, config)

class WikiPublicationDepthTests(unittest.TestCase):
    def test_wiki_json_link_bundle_and_publication_paths(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            bad = root / "bad.json"
            bad.write_text('{"x": 1, "x": 2}', encoding="utf-8")
            with self.assertRaises(ValueError):
                wiki.read_json(bad)
            bad.write_text("[]", encoding="utf-8")
            with self.assertRaises(ValueError):
                wiki.read_json(bad)

            rendered = wiki.render_page(
                "[Guide](Other.md#part) [Root](../../README.md) [Remote](https://example.com/x) [A](#a)",
                "Home.md",
                {"Home", "Other"},
                "owner/repo",
                "a" * 40,
            )
            self.assertIn("](Other#part)", rendered)
            self.assertIn("/blob/" + "a" * 40 + "/README.md)", rendered)
            self.assertIn("https://example.com/x", rendered)
            self.assertIn("](#a)", rendered)
            with self.assertRaises(ValueError):
                wiki.render_page("[Escape](../../../outside.md)", "Home.md", {"Home"}, "owner/repo", "a" * 40)

            profile = root / ".github/repository-profile.json"
            profile.parent.mkdir(parents=True)
            profile.write_text(json.dumps({"repository": "owner/repo"}), encoding="utf-8")
            wiki_dir = root / "docs/wiki"
            wiki_dir.mkdir(parents=True)
            for page in ("Home.md", "_Sidebar.md"):
                (wiki_dir / page).write_text("# Page\n", encoding="utf-8")
            data = {"pages": [{"name": "Home"}]}
            with self.assertRaises(ValueError):
                wiki.bundle(root, data, "short")
            profile.write_text(json.dumps({"repository": "bad"}), encoding="utf-8")
            with self.assertRaises(ValueError):
                wiki.bundle(root, data, "a" * 40)
            profile.write_text(json.dumps({"repository": "owner/repo"}), encoding="utf-8")
            bundle = wiki.bundle(root, data, "a" * 40)
            self.assertIn("Wiki-Source.json", bundle)

            outside = root.parent / (root.name + "-wiki-export")
            if outside.exists():
                import shutil
                shutil.rmtree(outside)
            wiki.write_bundle({"Home.md": b"x"}, outside, root)
            self.assertEqual((outside / "Home.md").read_bytes(), b"x")
            with self.assertRaises(ValueError):
                wiki.write_bundle({"Home.md": b"x"}, outside, root)
            inside = root / "inside"
            with self.assertRaises(ValueError):
                wiki.write_bundle({"Home.md": b"x"}, inside, root)
            import shutil
            shutil.rmtree(outside)

    def test_wiki_build_report_export_and_compare(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            wiki_dir = root / "docs/wiki"
            wiki_dir.mkdir(parents=True)
            (wiki_dir / wiki.INVENTORY).write_text("inventory\n", encoding="utf-8")
            (wiki_dir / "_Sidebar.md").write_text("sidebar\n", encoding="utf-8")
            export_dir = root.parent / (root.name + "-export")
            published_dir = root.parent / (root.name + "-published")
            args = type("Args", (), {"root": root, "write_reference": False, "export_dir": export_dir, "published_dir": published_dir})()
            git_result = type("GitResult", (), {"stdout": "a" * 40 + "\n"})()
            clean_result = type("GitResult", (), {"stdout": ""})()
            with (
                patch.object(wiki, "tracked_files", return_value=set()),
                patch.object(wiki, "read_json", return_value={}),
                patch.object(wiki, "validate_catalogue"),
                patch.object(wiki, "inventory", return_value="inventory\n"),
                patch.object(wiki, "sidebar", return_value="sidebar\n"),
                patch.object(wiki, "git_text", side_effect=[git_result, clean_result]),
                patch.object(wiki, "bundle", return_value={"Home.md": b"x"}),
                patch.object(wiki, "write_bundle") as write_bundle,
                patch.object(wiki, "compare_bundle", return_value=["published bytes differ: Home.md"]) as compare,
            ):
                report = wiki.build_report(args)
            self.assertEqual(report["status"], "fail")
            write_bundle.assert_called_once()
            compare.assert_called_once()
            self.assertIn("published bytes differ", report["findings"][0])


class TemplateContractBoundaryTests(unittest.TestCase):
    def test_contract_io_notes_and_mode_boundaries(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            with self.assertRaises(contract.ContractError):
                contract._read_json(root, "missing.json")
            bad = root / "bad.json"
            bad.write_text("{", encoding="utf-8")
            with self.assertRaises(contract.ContractError):
                contract._read_json(root, "bad.json")
            with self.assertRaises(contract.ContractError):
                contract._check_notes(root, [])
            self.assertIsNone(contract._project_version(root))

            profile = root / contract.CONFIG_PATH
            profile.parent.mkdir(parents=True)
            profile.write_text("[]", encoding="utf-8")
            with self.assertRaises(contract.ContractError):
                contract.run_check(root)

            findings: list[dict[str, str]] = []
            self.assertEqual(contract._check_shape([], findings), ("", ""))
            self.assertTrue(findings)
            findings = []
            self.assertEqual(contract._check_supported("9.9.9", findings), frozenset())
            self.assertTrue(findings)
            findings = []
            contract._check_mode({"mode": "template", "repository": "owner/repo"}, "other/template", findings)
            contract._check_mode({"mode": "generated", "repository": "owner/repo"}, "owner/repo", findings)
            self.assertEqual(len(findings), 2)

            with patch.object(contract, "run_gate", return_value=7) as gate:
                self.assertEqual(contract.main([]), 7)
            gate.assert_called_once()


class LocalActionDepthTests(unittest.TestCase):
    def _action(self, text: str, tracked_extra: set[str] | None = None) -> list[dict[str, Any]]:
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        root = Path(temporary.name)
        action = root / ".github/actions/demo"
        action.mkdir(parents=True)
        (action / "action.yml").write_text(text, encoding="utf-8")
        files = {".github/actions/demo/action.yml"} | (tracked_extra or set())
        return local_actions.validate_action(root, files, ".github/workflows/ci.yml", 4, "./.github/actions/demo")

    def test_local_action_runtime_variants(self) -> None:
        prefix = "name: Demo\ndescription: Demo action\nruns:\n"
        findings = self._action(prefix + "  using: composite\n")
        self.assertTrue(any("runs.steps" in item["message"] for item in findings))
        findings = self._action(prefix + "  using: node20\n")
        self.assertTrue(any("runs.main" in item["message"] for item in findings))
        findings = self._action(prefix + "  using: docker\n")
        self.assertTrue(any("runs.image" in item["message"] for item in findings))
        findings = self._action(prefix + "  using: docker\n  image: ghcr.io/owner/image:latest\n")
        self.assertTrue(any("Dockerfile" in item["message"] for item in findings))
        findings = self._action(prefix + "  using: docker\n  image: Dockerfile\n")
        self.assertTrue(any("does not exist" in item["message"] for item in findings))

        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            action = root / ".github/actions/demo"
            action.mkdir(parents=True)
            (action / "action.yml").write_text(prefix + "  using: docker\n  image: Dockerfile\n", encoding="utf-8")
            (action / "Dockerfile").write_text("FROM scratch\n", encoding="utf-8")
            findings = local_actions.validate_action(root, {".github/actions/demo/action.yml"}, ".github/workflows/ci.yml", 4, "./.github/actions/demo")
            self.assertTrue(any("not tracked" in item["message"] for item in findings))

        findings = self._action(prefix + "  using: python99\n")
        self.assertTrue(any("Unsupported" in item["message"] for item in findings))


class PolicyCoverageRunnerDepthTests(unittest.TestCase):
    def test_policy_report_failure_matrix_and_markdown(self) -> None:
        core = {
            "source": "tools/check_repo.py",
            "sites": 2,
            "covered": [{"id": "covered", "function": "f", "line": 1, "cases": ["a"], "expression": "x"}],
            "uncovered": [{"id": "miss", "function": "g", "line": 2, "cases": [], "expression": "y"}],
            "unexpected_sites": [{"id": "z", "cases": ["b"]}],
            "cases": [{"id": "bad-case", "owner": "rule", "status": "fail", "detail": "bad"}],
        }
        failed_tool = {"tool": "tools/check_release.py", "status": "fail", "exit_code": 1, "summary": "bad"}
        operational = {"id": "top-level-operational-error", "status": "fail", "exit_code": 1, "expected_exit_code": 2}
        delegated = {"owner": "tools/test_workflow_validation.py", "execution": "delegated", "fixtures": ["x"], "status": "fail", "missing": ["x"], "terminally_required": False}
        with (
            patch.object(coverage_runner, "run_core_coverage", return_value=core),
            patch.object(coverage_runner, "run_tool_selftest", return_value=failed_tool),
            patch.object(coverage_runner, "operational_exit_fixture", return_value=operational),
            patch.object(coverage_runner, "delegated_workflow_fixtures", return_value=delegated),
        ):
            report = coverage_runner.build_report(Path("."))
        self.assertEqual(report["status"], "fail")
        self.assertGreaterEqual(len(report["failures"]), 6)
        rendered = coverage_runner.markdown_report(report)
        self.assertIn("Uncovered canonical finding sites", rendered)
        self.assertIn("Failed fixture assertions", rendered)
        self.assertIn("Coverage failures", rendered)

    def test_policy_selftest_and_main_failure_paths(self) -> None:
        report = {"status": "fail", "counts": {"canonical_sites": 0, "canonical_covered": 0, "canonical_uncovered": 0, "canonical_cases": 0, "focused_selftests": 0, "matrix_rows": 0, "exclusions": 0}, "uncovered": [], "canonical_cases": [], "failures": ["incomplete"], "exclusions": []}
        with patch.object(coverage_runner, "build_report", return_value=report):
            self.assertEqual(coverage_runner.run_self_test(Path(".")), 1)
        with patch.object(coverage_runner, "run_gate", return_value=9):
            self.assertEqual(coverage_runner.main([]), 9)


class SnapshotCollectorDepthTests(unittest.TestCase):
    def test_pages_and_tag_resolution_boundaries(self) -> None:
        with patch.object(snapshot, "get", return_value=[]):
            self.assertEqual(snapshot.pages("owner/repo/items"), [])
        with patch.object(snapshot, "get", return_value={"total_count": 1000, "items": []}):
            with self.assertRaises(ValueError):
                snapshot.pages("owner/repo/items", "items")
        with patch.object(snapshot, "get", return_value={}):
            with self.assertRaises(ValueError):
                snapshot.pages("owner/repo/items")
        with patch.object(snapshot, "get", return_value=[None] * 100):
            with self.assertRaises(ValueError):
                snapshot.pages("owner/repo/items")

        with patch.object(snapshot, "get", return_value={"object": {"type": "commit", "sha": "a" * 40}}):
            self.assertEqual(snapshot.tag_commit("owner/repo", "v1"), "a" * 40)
        with patch.object(snapshot, "get", return_value={"object": {"type": "blob", "sha": "a" * 40}}):
            with self.assertRaises(ValueError):
                snapshot.tag_commit("owner/repo", "v1")
        tag_obj = {"object": {"type": "tag", "sha": "a" * 40}}
        with patch.object(snapshot, "get", return_value=tag_obj):
            with self.assertRaises(ValueError):
                snapshot.tag_commit("owner/repo", "v1")

    def test_quality_collect_and_main_degraded_paths(self) -> None:
        repo = {"repository": "owner/repo", "commit": "a" * 40, "captured_at": "2026-09-12T00:00:00+00:00", "unavailable": {}}
        with patch.object(snapshot, "pages", side_effect=ValueError("unavailable")):
            quality = snapshot.quality_evidence(repo)
        self.assertIsNone(quality["workflows"])
        self.assertIsNone(quality["releases"])
        self.assertIn("workflows", repo["unavailable"])
        self.assertIn("releases", repo["unavailable"])

        with self.assertRaises(ValueError):
            snapshot.collect("bad repo")

        metadata = {"default_branch": "main", "description": "x", "has_issues": True, "allow_auto_merge": False}
        tree = {"tree": [], "truncated": True}
        with (
            patch.object(snapshot, "get", side_effect=[metadata, {"sha": "a" * 40}, tree]),
            patch.object(snapshot, "pages", side_effect=snapshot.urllib.error.HTTPError("u", 403, "no", Message(), None)),
        ):
            collected = snapshot.collect("owner/repo")
        self.assertIsNone(collected["paths"])
        self.assertIn("paths", collected["unavailable"])
        self.assertIn("labels", collected["unavailable"])
        self.assertIn("rulesets", collected["unavailable"])

        with patch.object(snapshot.sys, "argv", ["collector", "owner/repo", "owner/repo", "--contract-commit", "a" * 40]):
            self.assertEqual(snapshot.main(), 2)
        with patch.object(snapshot.sys, "argv", ["collector", "owner/repo", "--contract-commit", "bad"]):
            self.assertEqual(snapshot.main(), 2)
        with (
            patch.object(snapshot.sys, "argv", ["collector", "owner/repo", "--contract-commit", "a" * 40, "--quality"]),
            patch.object(snapshot, "collect", return_value=dict(repo)),
            patch.object(snapshot, "quality_evidence", return_value={"workflows": [], "releases": []}),
        ):
            self.assertEqual(snapshot.main(), 0)


class SharedGateAndWhitespaceDepthTests(unittest.TestCase):
    def test_shared_gate_tracked_files_and_whitespace_failures(self) -> None:
        failed = type("Result", (), {"returncode": 1, "stdout": "", "stderr": b"git failed"})()
        with patch.object(gatelib, "git_bytes", return_value=failed):
            with self.assertRaisesRegex(RuntimeError, "git failed"):
                gatelib.tracked_files(Path("."))

        failed_text = type("Result", (), {"returncode": 1, "stdout": "", "stderr": "bad revision"})()
        with patch.object(whitespace, "git", return_value=failed_text):
            with self.assertRaisesRegex(RuntimeError, "Cannot resolve commit"):
                whitespace.resolve_commit(Path("."), "missing")

        with (
            patch.object(whitespace, "resolve_commit", side_effect=["a" * 40, "b" * 40]),
            patch.object(whitespace, "git", return_value=failed_text),
        ):
            with self.assertRaisesRegex(RuntimeError, "Cannot compute merge base"):
                whitespace.resolve_committed_scope(Path("."), "head", "base")

        with self.assertRaisesRegex(ValueError, "--base"):
            whitespace.run_check(Path("."), "working-tree", base_revision="main")
        with self.assertRaisesRegex(ValueError, "--head"):
            whitespace.run_check(Path("."), "working-tree", head_revision="candidate")

        working = {
            "mode": "working-tree", "status": "fail", "basis": "working-tree",
            "base": "a", "head": None, "range": "working-tree",
            "findings": ["unstaged: trailing whitespace"],
        }
        committed = {
            "mode": "committed", "status": "pass", "basis": "merge-base",
            "base": "a", "head": "b", "range": "a..b", "findings": [],
        }
        self.assertIn("staged and unstaged", whitespace.markdown_report(working))
        self.assertIn("Inspected range", whitespace.markdown_report(committed))
        self.assertIn("trailing whitespace", whitespace.markdown_report(working))

        synthetic = {
            "status": "fail", "basis": "empty-tree", "base": "", "head": "", "findings": []
        }
        with patch.object(whitespace, "run_check", return_value=synthetic):
            self.assertEqual(whitespace.run_self_test(), 1)
        with patch.object(whitespace, "run_gate", return_value=7):
            self.assertEqual(whitespace.main(["--mode", "working-tree"]), 7)


class ExternalLinkBoundaryTests(unittest.TestCase):
    @staticmethod
    def policy() -> dict[str, Any]:
        return {
            "attempts": 3,
            "timeout_seconds": 1,
            "concurrency": 1,
            "redirects": 1,
            "max_links": 10,
            "domains": {"example.com": "fixture"},
            "exceptions": [],
            "classifications": [],
        }

    def test_destinations_status_transport_and_probe_boundaries(self) -> None:
        found = external_links.destinations(
            "<https://example.com/a>\n"
            "<a href='https://example.com/b'>x</a>\n"
            "```\n<https://example.com/ignored>\n```\n"
        )
        values = {value for _, value in found}
        self.assertIn("https://example.com/a", values)
        self.assertIn("https://example.com/b", values)
        self.assertNotIn("https://example.com/ignored", values)

        policy = self.policy()
        self.assertEqual(external_links.url_status("mailto:test@example.com", policy), "NOT_APPLICABLE")
        self.assertEqual(external_links.url_status("http://example.com/a", policy), "POLICY_BLOCKED")
        self.assertEqual(external_links.url_status("https://user@example.com/a", policy), "ACCESS_RESTRICTED")
        self.assertEqual(external_links.url_status("https://example.com/a?q=1", policy), "ACCESS_RESTRICTED")
        self.assertEqual(external_links.url_status("https://other.example/a", policy), "POLICY_BLOCKED")
        self.assertEqual(external_links.url_status("https://[broken", policy), "POLICY_BLOCKED")
        self.assertIsNone(external_links.url_status("https://example.com/a", policy))

        def value_error(_url: str, _timeout: int) -> tuple[int, str | None]:
            raise ValueError("blocked")

        def os_error(_url: str, _timeout: int) -> tuple[int, str | None]:
            raise OSError("offline")

        self.assertEqual(external_links.attempt("https://example.com/a", policy, value_error)[0], "POLICY_BLOCKED")
        self.assertEqual(external_links.attempt("https://example.com/a", policy, os_error)[0], "TRANSIENT_FAILURE")
        self.assertEqual(external_links.attempt("https://example.com/a", policy, lambda _u, _t: (302, None)), ("REDIRECT_FAILURE", 302))
        self.assertEqual(external_links.attempt("https://example.com/a", policy, lambda _u, _t: (302, "/a"))[0], "REDIRECT_FAILURE")
        self.assertEqual(external_links.attempt("https://example.com/a", policy, lambda _u, _t: (204, None)), ("OK", 204))
        self.assertEqual(external_links.attempt("https://example.com/a", policy, lambda _u, _t: (403, None)), ("ACCESS_RESTRICTED", 403))
        self.assertEqual(external_links.attempt("https://example.com/a", policy, lambda _u, _t: (429, None)), ("TRANSIENT_FAILURE", 429))
        self.assertEqual(external_links.attempt("https://example.com/a", policy, lambda _u, _t: (404, None)), ("PERMANENT_FAILURE", 404))

        responses = iter([(500, None), (200, None)])
        pauses: list[int] = []
        recovered = external_links.probe(
            "https://example.com/a", policy,
            lambda _u, _t: next(responses),
            pauses.append,
        )
        self.assertEqual(recovered["status"], "OK")
        self.assertEqual(recovered["attempts"], 2)
        self.assertEqual(pauses, [1])
        permanent = external_links.probe(
            "https://example.com/a", policy,
            lambda _u, _t: (404, None),
            lambda _n: None,
        )
        self.assertEqual(permanent["status"], "PERMANENT_FAILURE")
        self.assertEqual(permanent["attempts"], 3)

        class Raw:
            closed = False

            def close(self) -> None:
                self.closed = True

        class TLS:
            def wrap_socket(self, *_args: Any, **_kwargs: Any) -> Any:
                raise RuntimeError("tls")

        raw = Raw()
        connection = external_links.PinnedHTTPS("example.com", "8.8.8.8", 1)
        setattr(connection, "tls_context", TLS())
        with patch.object(external_links.socket, "create_connection", return_value=raw):
            with self.assertRaisesRegex(RuntimeError, "tls"):
                connection.connect()
        self.assertTrue(raw.closed)

        class Response:
            status = 200

            @staticmethod
            def getheader(_name: str) -> None:
                return None

        class Connection:
            closed = False

            def request(self, *_args: Any, **_kwargs: Any) -> None:
                return None

            @staticmethod
            def getresponse() -> Response:
                return Response()

            def close(self) -> None:
                self.closed = True

        fake = Connection()
        with (
            patch.object(external_links.socket, "getaddrinfo", return_value=[(None, None, None, None, ("8.8.8.8", 443))]),
            patch.object(external_links, "PinnedHTTPS", return_value=fake),
        ):
            self.assertEqual(external_links.request("https://example.com/a b", 1), (200, None))
        self.assertTrue(fake.closed)
        with patch.object(external_links.socket, "getaddrinfo", return_value=[(None, None, None, None, ("127.0.0.1", 443))]):
            with self.assertRaises(ValueError):
                external_links.request("https://example.com/a", 1)

    def test_collect_report_limit_and_cli_paths(self) -> None:
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            (root / "README.md").write_text(
                "<https://example.com/a#frag>\n<a src='https://example.com/b'>x</a>\n",
                encoding="utf-8",
            )
            (root / "plain.txt").write_text("<https://example.com/c>\n", encoding="utf-8")
            with patch.object(external_links, "tracked_files", return_value={"README.md", "plain.txt"}):
                links = external_links.collect(root)
            self.assertEqual(len(links), 2)
            self.assertTrue(all("#" not in row["url"] for row in links.values()))

        first, second, third = "a" * 64, "b" * 64, "c" * 64
        policy = self.policy()
        policy["exceptions"] = [{"id": second, "reason": "fixture", "expires": "2027-01-01"}]
        policy["classifications"] = [
            {"id": first, "kind": "restricted-historical", "reason": "fixture", "expires": "2027-01-01"}
        ]
        links = {
            first: {"url": "https://example.com/a", "locations": ["README.md:1"]},
            second: {"url": "https://example.com/b", "locations": ["README.md:2"]},
            third: {"url": "https://example.com/c", "locations": ["README.md:3"]},
        }
        with (
            patch.object(external_links, "load_policy", return_value={"network": policy}),
            patch.object(external_links, "collect", return_value=links),
        ):
            report = external_links.build_report(
                Path("."), external_links.date(2026, 9, 12),
                transport=lambda _u, _t: (200, None), pause=lambda _n: None,
            )
        self.assertEqual(report["status"], "fail")
        self.assertEqual(report["counts"]["restricted_historical"], 1)
        self.assertEqual({row["status"] for row in report["links"]}, {"RESTRICTED_HISTORICAL", "EXCEPTED", "OK"})
        self.assertIn("RESTRICTED_HISTORICAL", external_links.markdown(report))

        limited_policy = self.policy()
        limited_policy["max_links"] = 1
        with (
            patch.object(external_links, "load_policy", return_value={"network": limited_policy}),
            patch.object(external_links, "collect", return_value=links),
        ):
            limited = external_links.build_report(
                Path("."), external_links.date(2026, 9, 12),
                transport=lambda _u, _t: (200, None), pause=lambda _n: None,
            )
        self.assertTrue(limited["limit_exceeded"])
        self.assertEqual(limited["status"], "fail")

        with (
            patch("sys.argv", ["links", "--as-of", "2026-09-12"]),
            patch.object(external_links, "run_gate", return_value=7),
        ):
            self.assertEqual(external_links.main(), 7)


class TemplateContractFailureDepthTests(unittest.TestCase):
    def test_contract_markdown_and_selftest_failure_reporting(self) -> None:
        report = {
            "status": "fail",
            "mode": "generated",
            "contract_version": "1.2.0",
            "contract_source": "owner/template",
            "project_version": "9.9.9",
            "resolved_rule_set": ["one", "two"],
            "supported_versions": ["1.2.0"],
            "findings": [{"message": "fixture failure"}],
        }
        rendered = contract.markdown_report(report)
        self.assertIn("Controls required", rendered)
        self.assertIn("fixture failure", rendered)

        synthetic = {
            "status": "fail",
            "findings": [],
            "contract_version": "1.2.0",
            "project_version": "1.2.0",
            "resolved_rule_set": [],
        }
        with (
            patch.object(contract, "_cases", return_value=[("expected-pass", {}, None, True, "")]),
            patch.object(contract, "_run_case", return_value=synthetic),
            patch.object(contract, "_fixture", return_value=Path(".")),
            patch.object(contract, "run_check", return_value=synthetic),
            patch.object(contract, "write_text"),
        ):
            self.assertEqual(contract.run_self_test(), 1)


class SnapshotCollectorAdditionalDepthTests(unittest.TestCase):
    def test_get_pagination_tag_and_release_success_boundaries(self) -> None:
        class Response:
            def __enter__(self) -> "Response":
                return self

            def __exit__(self, *_args: Any) -> None:
                return None

            @staticmethod
            def read(_limit: int) -> bytes:
                return b"x" * 10_000_001

        class Opener:
            @staticmethod
            def open(*_args: Any, **_kwargs: Any) -> Response:
                return Response()

        with patch.object(snapshot.urllib.request, "build_opener", return_value=Opener()):
            with self.assertRaisesRegex(ValueError, "size limit"):
                snapshot.get("owner/repo")

        with patch.object(snapshot, "get", return_value={"total_count": 1000, "items": []}):
            with self.assertRaisesRegex(ValueError, "Actions search limit"):
                snapshot.pages("x", "items")
        with patch.object(snapshot, "get", return_value={"x": 1}):
            with self.assertRaisesRegex(ValueError, "Expected paginated"):
                snapshot.pages("x")
        with patch.object(snapshot, "get", return_value=[0] * 100):
            with self.assertRaisesRegex(ValueError, "Pagination limit"):
                snapshot.pages("x")

        with patch.object(snapshot, "get", return_value={"object": {"type": "blob", "sha": "a" * 40}}):
            with self.assertRaisesRegex(ValueError, "does not resolve"):
                snapshot.tag_commit("owner/repo", "v1")
        with patch.object(snapshot, "get", return_value={"object": {"type": "tag", "sha": "a" * 40}}):
            with self.assertRaisesRegex(ValueError, "indirection limit"):
                snapshot.tag_commit("owner/repo", "v1")

        repo = {
            "repository": "owner/repo",
            "commit": "a" * 40,
            "captured_at": "2026-09-12T00:00:00+00:00",
            "unavailable": {},
        }
        release_row = {
            "id": 7, "tag_name": "v1.2.0", "draft": False, "prerelease": False,
            "published_at": "2026-09-12T00:00:00Z", "body": "notes", "assets": [],
        }
        with (
            patch.object(snapshot, "pages", side_effect=[[], [release_row]]),
            patch.object(snapshot, "tag_commit", return_value="a" * 40),
        ):
            quality = snapshot.quality_evidence(repo)
        self.assertEqual(quality["releases"][0]["resolved_commit"], "a" * 40)

    def test_collect_blob_failure_is_explicit(self) -> None:
        config_path = snapshot.CONFIG
        metadata = {
            "default_branch": "main", "description": "x", "has_issues": True,
            "allow_auto_merge": False,
        }
        tree = {"tree": [{"path": config_path, "sha": "b" * 40, "type": "blob"}], "truncated": False}

        def fake_get(path: str) -> Any:
            if path == "owner/repo":
                return metadata
            if "/commits/" in path:
                return {"sha": "a" * 40}
            if "/git/trees/" in path:
                return tree
            if "/git/blobs/" in path:
                raise snapshot.urllib.error.HTTPError("u", 403, "no", Message(), None)
            raise AssertionError(path)

        with (
            patch.object(snapshot, "get", side_effect=fake_get),
            patch.object(snapshot, "pages", return_value=[]),
        ):
            report = snapshot.collect("owner/repo")
        self.assertIn(config_path, report["unavailable"])
        self.assertNotIn(config_path, report["files"])


class ProvisioningPrimitiveDepthTests(unittest.TestCase):
    def test_transport_collection_and_source_json_boundaries(self) -> None:
        with self.assertRaises(ValueError):
            provision.GitHub("bad", None)
        api = provision.GitHub("owner/repo", None)
        with self.assertRaises(ValueError):
            api.request("GET", "relative")
        with self.assertRaises(ValueError):
            api.request("POST", "/labels", {})

        class Response:
            def __init__(self, data: bytes):
                self.data = data

            def __enter__(self) -> "Response":
                return self

            def __exit__(self, *_args: Any) -> None:
                return None

            def read(self, _limit: int) -> bytes:
                return self.data

        class Opener:
            def __init__(self, data: bytes):
                self.data = data

            def open(self, *_args: Any, **_kwargs: Any) -> Response:
                return Response(self.data)

        with patch.object(provision, "build_opener", return_value=Opener(b'{"ok": true}')):
            self.assertEqual(api.request("GET"), {"ok": True})
        with patch.object(provision, "build_opener", return_value=Opener(b"x" * 10_000_001)):
            with self.assertRaises(ValueError):
                api.request("GET")

        class BadCollection:
            @staticmethod
            def request(*_args: Any, **_kwargs: Any) -> dict[str, int]:
                return {"bad": 1}

        with self.assertRaisesRegex(ValueError, "Invalid collection"):
            provision.collection(BadCollection(), "/labels")

        class FullCollection:
            @staticmethod
            def request(*_args: Any, **_kwargs: Any) -> list[int]:
                return [0] * 100

        with self.assertRaisesRegex(ValueError, "Incomplete collection"):
            provision.collection(FullCollection(), "/labels")

        class Source:
            def __init__(self, value: dict[str, Any]):
                self.value = value

            def request(self, *_args: Any, **_kwargs: Any) -> dict[str, Any]:
                return self.value

        with self.assertRaisesRegex(ValueError, "Expected source JSON"):
            provision.source_json(Source({"encoding": "utf-8", "type": "file"}), "x.json", "a" * 40)
        encoded = provision.base64.b64encode(b'{"x": 1}').decode("ascii")
        parsed = provision.source_json(
            Source({"encoding": "base64", "type": "file", "content": encoded}),
            "x.json", "a" * 40,
        )
        self.assertEqual(parsed, {"x": 1})

class CloseoutCliDepthTests(unittest.TestCase):
    def test_closeout_parse_and_main_paths(self) -> None:
        with self.assertRaises(SystemExit):
            closeout.parse_arguments([])
        options, self_test = closeout.parse_arguments(["--self-test"])
        self.assertIsNone(options)
        self.assertTrue(self_test)

        with patch.object(closeout, "run_self_test", return_value=6):
            self.assertEqual(closeout.main(["--self-test"]), 6)

        args = ["--snapshot", "snapshot.json", "--tag", "v1.2.1", "--candidate-sha", "a" * 40, "--milestone-number", "4", "--allow-not-latest", "--expect-prerelease"]
        parsed, self_test = closeout.parse_arguments(args)
        self.assertFalse(self_test)
        self.assertIsNotNone(parsed)
        assert parsed is not None
        self.assertFalse(parsed.expect_latest)
        self.assertTrue(parsed.expect_prerelease)

        with (patch.object(closeout, "build_report", return_value={"status": "pass"}), patch.object(closeout, "markdown_report", return_value="ok\n")):
            self.assertEqual(closeout.main(args), 0)
        with (patch.object(closeout, "build_report", return_value={"status": "fail"}), patch.object(closeout, "markdown_report", return_value="fail\n")):
            self.assertEqual(closeout.main(args), 1)
        with patch.object(closeout, "build_report", side_effect=closeout.CloseoutError("boom")):
            self.assertEqual(closeout.main(args), 2)

class RemainingCoverageDepthTests(unittest.TestCase):
    def test_semver_comparison_and_failure_boundaries(self) -> None:
        with self.assertRaises(ValueError):
            semantics.parse_semver("1.0")
        with self.assertRaises(ValueError):
            semantics.parse_semver("1.0.0-01")
        stable = semantics.parse_semver("1.0.0")
        prerelease = semantics.parse_semver("1.0.0-rc.1")
        self.assertEqual(semantics.compare(stable, stable), 0)
        self.assertEqual(semantics.compare(semantics.parse_semver("2.0.0"), stable), 1)
        self.assertEqual(semantics.compare(stable, prerelease), 1)
        self.assertEqual(semantics.compare(prerelease, stable), -1)
        self.assertEqual(
            semantics.compare(
                semantics.parse_semver("1.0.0-alpha.2"),
                semantics.parse_semver("1.0.0-alpha.1"),
            ),
            1,
        )
        self.assertEqual(
            semantics.compare(
                semantics.parse_semver("1.0.0-1"),
                semantics.parse_semver("1.0.0-alpha"),
            ),
            -1,
        )
        self.assertEqual(
            semantics.compare(
                semantics.parse_semver("1.0.0-alpha"),
                semantics.parse_semver("1.0.0-1"),
            ),
            1,
        )
        self.assertEqual(
            semantics.compare(
                semantics.parse_semver("1.0.0-beta"),
                semantics.parse_semver("1.0.0-alpha"),
            ),
            1,
        )
        self.assertEqual(
            semantics.compare(
                semantics.parse_semver("1.0.0-alpha"),
                semantics.parse_semver("1.0.0-alpha.1"),
            ),
            -1,
        )
        self.assertFalse(semantics.valid_date("2026-02-30"))
        with patch.object(semantics, "_history_self_test_cases", return_value=[("forced", False)]):
            self.assertEqual(semantics.run_self_test(), 1)

    def test_checker_development_ast_failure_boundaries(self) -> None:
        with patch.object(checker_dev.importlib.util, "spec_from_file_location", return_value=None):
            with self.assertRaises(checker_dev.ContractError):
                checker_dev.load_checker(Path("missing.py"))
        with self.assertRaises(checker_dev.ContractError):
            checker_dev.node_name(ast.Expr(value=ast.Constant(value="x")))

        tree = ast.parse("def first():\n    pass\n\ndef second():\n    pass\n")
        with patch.object(checker_dev, "SECTION_STARTS", (("one", "missing"),)):
            rows, failures = checker_dev.section_report("", tree)
        self.assertEqual(rows, [])
        self.assertTrue(failures)

        with patch.object(
            checker_dev,
            "SECTION_STARTS",
            (("one", "second"), ("two", "first")),
        ):
            rows, failures = checker_dev.section_report("", tree)
        self.assertEqual(rows, [])
        self.assertTrue(any("strictly ordered" in item for item in failures))

        with patch.object(checker_dev, "SECTION_STARTS", (("one", "second"),)):
            rows, failures = checker_dev.section_report("", tree)
        self.assertEqual(len(rows), 1)
        self.assertTrue(any("ownership covers" in item for item in failures))

        imports_tree = ast.parse("from .local import x\nimport definitely_not_stdlib\n")
        observed, failures = checker_dev.import_report(imports_tree)
        self.assertIn("definitely_not_stdlib", observed)
        self.assertTrue(any("relative imports" in item for item in failures))
        self.assertTrue(any("non-approved" in item for item in failures))

        fake_module: Any = type("FakeModule", (), {"CHECKS": (lambda: None,)})()
        functions, failures = checker_dev.check_ids(fake_module)
        self.assertEqual(functions, ["<lambda>"])
        self.assertTrue(failures)

    def test_whitespace_operational_and_selftest_failure_boundaries(self) -> None:
        failed = type("Result", (), {"returncode": 1, "stdout": "", "stderr": "git failed"})()
        passed = type("Result", (), {"returncode": 0, "stdout": "", "stderr": ""})()
        with patch.object(whitespace, "git", return_value=failed):
            with self.assertRaises(RuntimeError):
                whitespace.resolve_commit(Path("."), "bad")
        with (
            patch.object(whitespace, "resolve_commit", side_effect=["h" * 40, "b" * 40]),
            patch.object(whitespace, "git", return_value=failed),
        ):
            with self.assertRaises(RuntimeError):
                whitespace.resolve_committed_scope(Path("."), "HEAD", "base")
        with self.assertRaises(ValueError):
            whitespace.run_check(Path("."), "working-tree", base_revision="base")
        with self.assertRaises(ValueError):
            whitespace.run_check(Path("."), "working-tree", head_revision="other")

        rendered = whitespace.markdown_report(
            {
                "mode": "working-tree",
                "status": "fail",
                "basis": "working-tree",
                "base": "a" * 40,
                "head": None,
                "range": "working-tree",
                "findings": ["bad whitespace"],
            }
        )
        self.assertIn("staged and unstaged", rendered)
        self.assertIn("```text", rendered)

        fake_report = {
            "status": "pass",
            "basis": "first-parent",
            "base": "x",
            "head": "y",
            "findings": [],
        }
        with (
            patch.object(whitespace, "init_repo"),
            patch.object(whitespace, "commit_file", return_value="a" * 40),
            patch.object(whitespace, "git", return_value=passed),
            patch.object(whitespace, "run_check", return_value=fake_report),
        ):
            self.assertEqual(whitespace.run_self_test(), 1)

    def test_snapshot_success_paths_and_token_transport(self) -> None:
        run_old = {
            "id": 1,
            "path": ".github/workflows/ci.yml",
            "head_sha": "a" * 40,
            "head_branch": "main",
            "event": "push",
            "status": "completed",
            "conclusion": "success",
            "run_attempt": 1,
            "updated_at": "2026-09-12T00:00:00Z",
        }
        run_new = {**run_old, "id": 2}
        ignored = {**run_old, "id": 3, "head_sha": "b" * 40}
        job = {"id": 9, "name": "gate", "status": "completed", "conclusion": "success"}
        release_row = {
            "id": 5,
            "tag_name": "v1.2.0",
            "draft": False,
            "prerelease": False,
            "published_at": "2026-09-12T00:00:00Z",
            "body": "notes",
            "assets": [
                {
                    "name": "asset.zip",
                    "size": 1,
                    "digest": "sha256:x",
                    "browser_download_url": "https://example.invalid/asset.zip",
                }
            ],
        }

        def fake_pages(path: str, key: str | None = None) -> list[Any]:
            del key
            if "actions/runs?" in path:
                return [run_old, run_new, ignored]
            if "/attempts/" in path:
                return [job]
            if path.endswith("/releases"):
                return [release_row]
            raise AssertionError(path)

        repo: dict[str, Any] = {
            "repository": "owner/repo",
            "commit": "a" * 40,
            "unavailable": {},
        }
        with (
            patch.object(snapshot, "pages", side_effect=fake_pages),
            patch.object(snapshot, "tag_commit", return_value="a" * 40),
        ):
            quality = snapshot.quality_evidence(repo)
        self.assertEqual([item["id"] for item in quality["workflows"]], [2])
        self.assertEqual(quality["workflows"][0]["jobs"][0]["name"], "gate")
        self.assertEqual(quality["releases"][0]["resolved_commit"], "a" * 40)

        metadata = {
            "default_branch": "main",
            "description": "demo",
            "has_issues": True,
            "allow_auto_merge": False,
        }
        tree = {
            "tree": [{"path": snapshot.CONFIG, "sha": "blob", "type": "blob"}],
            "truncated": False,
        }

        def fake_get(path: str) -> Any:
            if path == "owner/repo":
                return metadata
            if path.endswith("/commits/main"):
                return {"sha": "c" * 40}
            if "/git/trees/" in path:
                return tree
            if path.endswith("/git/blobs/blob"):
                return {"content": base64.b64encode(b'{"profile":"library"}').decode("ascii")}
            raise AssertionError(path)

        with (
            patch.object(snapshot, "get", side_effect=fake_get),
            patch.object(snapshot, "pages", return_value=[]),
        ):
            collected = snapshot.collect("owner/repo")
        self.assertIn(snapshot.CONFIG, collected["files"])
        self.assertEqual(collected["labels"], [])
        self.assertEqual(collected["rulesets"], [])

        class Response:
            def __enter__(self) -> "Response":
                return self

            def __exit__(self, *args: Any) -> None:
                return None

            def read(self, size: int) -> bytes:
                self.assert_size = size
                return b"{}"

        class Opener:
            def open(self, request: Any, timeout: int) -> Response:
                del request, timeout
                return Response()

        with (
            patch.dict(snapshot.os.environ, {"GH_TOKEN": "token"}, clear=True),
            patch.object(snapshot.urllib.request, "build_opener", return_value=Opener()),
        ):
            self.assertEqual(snapshot.get("owner/repo"), {})

    def test_template_contract_selftest_failure_matrix(self) -> None:
        case_mismatch = {"status": "fail", "findings": []}
        case_missing = {"status": "fail", "findings": [{"message": "other"}]}
        old = {"resolved_rule_set": ["template-contract-version"]}
        deterministic_one = {"value": 1}
        deterministic_two = {"value": 2}
        before = {"contract_version": "1.0.0", "project_version": "9.9.9", "status": "pass"}
        after = {"contract_version": "2.0.0", "project_version": "2.0.0", "status": "fail"}
        cases: list[Any] = [
            ("status-mismatch", {}, None, True, ""),
            ("missing-finding", {}, None, False, "needle"),
        ]
        with (
            patch.object(contract, "_cases", return_value=cases),
            patch.object(
                contract,
                "_run_case",
                side_effect=[
                    case_mismatch,
                    case_missing,
                    old,
                    deterministic_one,
                    deterministic_two,
                ],
            ),
            patch.object(contract, "_fixture", return_value=Path(".")),
            patch.object(contract, "run_check", side_effect=[before, after]),
            patch.object(contract, "write_text"),
        ):
            self.assertEqual(contract.run_self_test(), 1)

        rendered = contract.markdown_report(
            {
                "status": "fail",
                "mode": "template",
                "contract_version": "1.2.0",
                "contract_source": "owner/template",
                "project_version": "9.9.9",
                "resolved_rule_set": [],
                "supported_versions": ["1.2.0"],
                "findings": [{"message": "bad contract"}],
            }
        )
        self.assertIn("bad contract", rendered)

    def test_local_action_parser_reporting_and_git_failures(self) -> None:
        self.assertIsNone(local_actions.uses_reference("run: echo no"))
        self.assertIsNone(local_actions.safe_relative("not-local"))
        self.assertIsNone(local_actions.safe_relative("./a/../b"))
        self.assertIsNone(local_actions.safe_relative("./a//b"))
        self.assertEqual(local_actions.unquote_scalar(' " value " '), "value")
        self.assertIsNone(local_actions.top_scalar("name: x\n", "description"))

        using, entrypoints, has_steps = local_actions.parse_runs_metadata(
            "runs:\n  using: 'docker'\n  image: Dockerfile\n  steps:\nnext: x\n"
        )
        self.assertEqual(using, "docker")
        self.assertEqual(entrypoints["image"], "Dockerfile")
        self.assertTrue(has_steps)

        report = {
            "status": "fail",
            "workflows": 1,
            "local_references": [{"reference": "./x"}],
            "findings": [{"path": "ci.yml", "line": 7, "message": "bad local action"}],
        }
        rendered = local_actions.markdown_report(report)
        self.assertIn("ci.yml:7", rendered)
        self.assertIn("bad local action", rendered)

        failed = type("Result", (), {"returncode": 1, "stdout": b"", "stderr": b"failed"})()
        passed = type("Result", (), {"returncode": 0, "stdout": b"", "stderr": b""})()
        with patch.object(local_actions, "git", return_value=failed):
            with self.assertRaises(RuntimeError):
                local_actions.fixture_report("./missing", None, (".github/workflows/test.yml",))
        with patch.object(local_actions, "git", side_effect=[passed, passed, passed, failed]):
            with self.assertRaises(RuntimeError):
                local_actions.fixture_report(
                    "./missing",
                    None,
                    (".github/workflows/test.yml",),
                )

    def test_policy_selftest_determinism_failure_branches(self) -> None:
        first = {
            "status": "pass",
            "counts": {
                "canonical_sites": 1,
                "canonical_covered": 1,
                "canonical_uncovered": 0,
                "canonical_cases": 0,
                "focused_selftests": 0,
                "matrix_rows": 0,
                "exclusions": 0,
            },
            "uncovered": [],
            "canonical_cases": [],
            "failures": [],
            "exclusions": [],
            "nonce": 1,
        }
        second = {**first, "nonce": 2}
        with patch.object(coverage_runner, "build_report", side_effect=[first, second]):
            self.assertEqual(coverage_runner.run_self_test(Path(".")), 1)


if __name__ == "__main__":
    unittest.main()
