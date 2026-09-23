"""Independent synthetic drift fixtures; never query or modify a live repository."""
from __future__ import annotations

import contextlib
import copy
import io
import json
import unittest
from email.message import Message
from unittest.mock import MagicMock, patch
from urllib.error import HTTPError

import check_portfolio_drift as drift
import collect_portfolio_snapshot as capture


def fixture(profile: str = "library", version: str = "1.2.0") -> dict:
    config = {"repository": "example/consumer", "profile": profile, "mode": "generated",
              "template_contract": {"version": version, "source": "example/template"},
              "placeholders": {"catalogue": {"NAME": {}}}}
    files = {path: "fixture\n" for paths in drift.PATHS.values() for path in paths}
    files[drift.CONFIG] = json.dumps(config)
    labels = [{"name": name, "color": "ff0000", "description": "Fixture"} for name in sorted(drift.CORE_LABEL_NAMES)]
    files[".github/labels.json"] = json.dumps({"core": labels})
    files[".github/release-policy.json"] = json.dumps({"core_checks": ["repository-integrity", "vba-compile", "regression"]})
    files[".github/workflows/static-checks.yml"] = """name: Integrity
on: push
permissions:
  contents: read
jobs:
  local-quality:
    runs-on: ubuntu-24.04
    steps:
      - run: echo fixture
"""
    files["docs/decision.md"] = "Reviewed local specialist control; no generic control waived.\n"
    branch = {"target": "branch", "enforcement": "active", "bypass_actors": [],
              "conditions": {"ref_name": {"include": ["~DEFAULT_BRANCH"], "exclude": []}},
              "rules": [{"type": name} for name in ("deletion", "non_fast_forward", "pull_request")] +
              [{"type": "required_status_checks", "parameters": {"strict_required_status_checks_policy": True,
                "required_status_checks": [{"context": "Stronger local checks"}]}}]}
    tag = {"target": "tag", "enforcement": "active", "bypass_actors": [],
           "conditions": {"ref_name": {"include": ["refs/tags/v*"], "exclude": []}},
           "rules": [{"type": name} for name in ("deletion", "non_fast_forward", "update")]}
    repo = {"repository": "example/consumer", "commit": "a" * 40, "paths": sorted(files), "files": files,
            "metadata": {"default_branch": "main", "description": "Fixture", "has_issues": True, "allow_auto_merge": False},
            "labels": labels,
            "rulesets": [branch, tag], "decisions": []}
    return {"schema_version": 1, "contract_source": "example/template", "contract_commit": "b" * 40, "repositories": [repo]}


class DriftTests(unittest.TestCase):
    def setUp(self) -> None:
        self.snapshot = fixture()
        self.repo = self.snapshot["repositories"][0]

    def row(self, rule: str) -> dict:
        return next(row for row in drift.build_report(self.snapshot)["findings"] if row["rule"] == rule)

    def decide(self, rule: str, decision: str) -> None:
        self.repo["decisions"].append({"rule": rule, "decision": decision, "reason": "Reviewed rationale",
            "commit": self.repo["commit"], "evidence_path": "docs/decision.md"})

    def remove(self, path: str) -> None:
        self.repo["paths"].remove(path)
        self.repo["files"].pop(path, None)

    def test_three_profiles(self) -> None:
        for profile in sorted(drift.PROFILES):
            with self.subTest(profile=profile):
                self.assertEqual(drift.build_report(fixture(profile))["status"], "pass")

    def test_missing_universal_path(self) -> None:
        for path in ("SECURITY.md", "tools/check_repo.py"):
            self.snapshot = fixture()
            self.repo = self.snapshot["repositories"][0]
            self.remove(path)
            self.assertEqual(self.row("canonical-repository-gate")["status"], "DRIFT")

    def test_defer_does_not_hide_failure(self) -> None:
        self.remove("SECURITY.md")
        self.decide("canonical-repository-gate", "DEFER")
        row = self.row("canonical-repository-gate")
        self.assertEqual((row["decision"], row["status"]), ("DEFER", "DRIFT"))

    def test_universal_not_applicable_rejected(self) -> None:
        self.decide("release-integrity", "NOT APPLICABLE")
        self.assertEqual(self.row("release-integrity")["status"], "DRIFT")

    def test_stronger_local_control_retained(self) -> None:
        self.decide("local-specialist", "KEEP")
        self.assertEqual(self.row("local-specialist")["decision"], "KEEP")
        self.assertEqual(self.row("branch-protection")["status"], "PASS")

    def test_profile_exception_is_not_universal(self) -> None:
        self.assertEqual(self.row("ui-evidence")["status"], "NOT_APPLICABLE")
        self.remove("RELEASING.md")
        self.assertEqual(self.row("release-integrity")["status"], "DRIFT")

    def test_unadopted_and_unavailable(self) -> None:
        self.remove(drift.CONFIG)
        self.repo["rulesets"] = None
        self.assertEqual(self.row("adoption")["decision"], "ADOPT")
        self.assertEqual(self.row("branch-protection")["status"], "UNVERIFIED")

    def test_ruleset_unavailable_not_absent(self) -> None:
        self.repo["rulesets"] = None
        self.assertEqual(self.row("branch-protection")["status"], "UNVERIFIED")
        self.repo["rulesets"] = []
        self.assertEqual(self.row("branch-protection")["status"], "DRIFT")

    def test_bypass_and_wrong_scope(self) -> None:
        self.repo["rulesets"][0]["bypass_actors"] = [{"actor_id": 5}]
        self.assertEqual(self.row("branch-protection")["status"], "DRIFT")
        self.repo["rulesets"][0]["conditions"]["ref_name"]["include"] = ["refs/heads/other"]
        self.assertEqual(self.row("branch-protection")["status"], "DRIFT")

    def test_label_missing_and_extra(self) -> None:
        self.repo["labels"].append({"name": "specialist"})
        self.assertEqual(self.row("label-policy")["status"], "PASS")
        self.repo["labels"] = []
        self.assertEqual(self.row("label-policy")["status"], "DRIFT")

    def test_manifest_cannot_remove_baseline(self) -> None:
        self.repo["files"][".github/labels.json"] = '{"core": []}'
        self.assertEqual(self.row("label-policy")["status"], "DRIFT")

    def test_release_core_removed(self) -> None:
        self.repo["files"][".github/release-policy.json"] = '{"core_checks": ["regression"]}'
        self.assertEqual(self.row("release-integrity")["status"], "DRIFT")

    def test_historical_rules(self) -> None:
        self.snapshot = fixture(version="1.0.0")
        self.repo = self.snapshot["repositories"][0]
        self.remove("docs/DEPENDENCY_UPDATES.md")
        self.assertNotIn("controlled-dependency-updates", [row["rule"] for row in drift.build_report(self.snapshot)["findings"]])

    def test_unsupported_version_and_foreign_source(self) -> None:
        for field, value in (("version", "99.0.0"), ("source", "other/template")):
            config = json.loads(fixture()["repositories"][0]["files"][drift.CONFIG])
            config["template_contract"][field] = value
            self.repo["files"][drift.CONFIG] = json.dumps(config)
            self.assertEqual(self.row("adoption")["status"], "UNVERIFIED")

    def test_decision_binding(self) -> None:
        self.decide("local-specialist", "KEEP")
        self.repo["decisions"][0]["commit"] = "b" * 40
        with self.assertRaises(ValueError):
            drift.build_report(self.snapshot)

    def test_deterministic_and_read_only(self) -> None:
        original = copy.deepcopy(self.snapshot)
        first = drift.build_report(self.snapshot)
        second = drift.build_report(self.snapshot)
        self.assertEqual(json.dumps(first, sort_keys=True), json.dumps(second, sort_keys=True))
        self.assertEqual(drift.markdown_report(first), drift.markdown_report(second))
        self.assertEqual(self.snapshot, original)
        self.assertTrue(all("a" * 40 in row["evidence"] for row in first["findings"]))

    def test_yaml_semantics(self) -> None:
        document = drift.yaml_document('on: push\npermissions: {contents: read}\njobs: {x: {steps: [{run: "permissions: write-all"}]}}')
        self.assertIn("on", document)
        self.assertFalse(drift.workflow_errors(document))
        document["permissions"] = {"contents": "write"}
        self.assertTrue(drift.workflow_errors(document))

    def test_yaml_ambiguous_constructs(self) -> None:
        for text in ('x: 1\nx: 2', 'x: &a [1]\ny: *a', 'x: !custom value', '!custom {x: 1}'):
            with self.subTest(text=text), self.assertRaises(ValueError):
                drift.yaml_document(text)

    def test_invalid_workflow_is_unverified(self) -> None:
        for source in ('jobs: [', 'permissions: {contents: read}\njobs: {x: {}}'):
            self.repo["files"][".github/workflows/static-checks.yml"] = source
            self.assertEqual(self.row("workflow-properties")["status"], "UNVERIFIED")

    def test_report_paths_cannot_collide(self) -> None:
        for arguments in (["--snapshot", "s.json", "--output", "s.json"],
                          ["--snapshot", "s.json", "--output", "out", "--summary", "out"]):
            with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as raised:
                drift.main(arguments)
            self.assertEqual(raised.exception.code, 2)

    def test_empty_portfolio(self) -> None:
        self.snapshot["repositories"] = []
        with self.assertRaises(ValueError):
            drift.build_report(self.snapshot)


class CaptureTests(unittest.TestCase):
    def test_get_only_and_no_redirect(self) -> None:
        opener = MagicMock()
        opener.open.return_value.__enter__.return_value.read.return_value = b'{}'
        with patch.object(capture.urllib.request, "build_opener", return_value=opener) as build:
            self.assertEqual(capture.get("example/consumer"), {})
        build.assert_called_once_with(capture.NoRedirect)
        request = opener.open.call_args.args[0]
        self.assertEqual(request.get_method(), "GET")
        self.assertEqual(request.full_url, "https://api.github.com/repos/example/consumer")
        capture.NoRedirect().redirect_request(None, None, 302, "", None, "https://other.invalid")

    def test_403_and_truncated_tree_stay_unknown(self) -> None:
        responses = [{"default_branch": "main"}, {"sha": "a" * 40}, {"tree": [], "truncated": True}]
        denied = HTTPError("https://api.github.com/", 403, "Denied", Message(), None)
        with patch.object(capture, "get", side_effect=responses), patch.object(capture, "pages", side_effect=[[], denied]):
            repo = capture.collect("example/consumer")
        self.assertIsNone(repo["paths"])
        self.assertIsNone(repo["rulesets"])
        self.assertEqual(repo["labels"], [])
        self.assertIn("403", repo["unavailable"]["rulesets"])

    def test_pagination_collects_all_or_fails(self) -> None:
        with patch.object(capture, "get", side_effect=[[{}] * 100, [{"last": True}]]):
            self.assertEqual(len(capture.pages("example/consumer/labels")), 101)
        with patch.object(capture, "get", return_value=[{}] * 100), self.assertRaises(ValueError):
            capture.pages("example/consumer/labels")

    def test_failed_capture_emits_no_snapshot(self) -> None:
        output = io.StringIO()
        with patch("sys.argv", ["capture", "example/consumer", "--contract-commit", "b" * 40]), \
                patch.object(capture, "collect", side_effect=OSError("Failed")), \
                contextlib.redirect_stdout(output), contextlib.redirect_stderr(io.StringIO()):
            self.assertEqual(capture.main(), 2)
        self.assertEqual(output.getvalue(), "")


if __name__ == "__main__":
    unittest.main()
