"""Independent report fixtures; synthetic data only, with no live repository writes."""
from __future__ import annotations

import contextlib
import copy
import io
import json
import unittest
from unittest.mock import patch

import collect_portfolio_snapshot as capture
import report_portfolio_quality as quality
from test_portfolio_drift import fixture

NOW = "2026-09-09T12:00:00Z"


def sample() -> dict:
    data = fixture()
    repo = data["repositories"][0]
    repo["unavailable"] = {}
    repo["quality"] = {"observed_at": NOW, "releases": [], "workflows": [
        {"id": 10, "path": ".github/workflows/local.yml", "event": "push", "head_branch": "main",
         "head_sha": repo["commit"], "status": "completed", "conclusion": "success", "run_attempt": 1,
         "jobs": [{"name": "Stronger local checks", "status": "completed", "conclusion": "success"}]}]}
    return data


class ReportTests(unittest.TestCase):
    def setUp(self) -> None:
        self.data = sample()
        self.repo = self.data["repositories"][0]

    def report(self, as_of: str = NOW) -> dict:
        return quality.build_report(self.data, as_of)

    def dimension(self, name: str, as_of: str = NOW) -> dict:
        return next(row for row in self.report(as_of)["repositories"][0]["dimensions"] if row["dimension"] == name)

    def test_complete_evidence_not_certification(self) -> None:
        report = self.report()
        self.assertEqual(report["status"], "pass")
        self.assertEqual(report["repositories"][0]["certification"], "NOT_ASSESSED")

    def test_seven_repositories_and_missing_adoption(self) -> None:
        self.repo["files"].pop(quality.drift.CONFIG)
        self.repo["paths"].remove(quality.drift.CONFIG)
        self.data["repositories"] = []
        for index in range(7):
            repo = copy.deepcopy(self.repo)
            repo["repository"] = f"example/repo-{index}"
            self.data["repositories"].append(repo)
        report = self.report()
        self.assertEqual(len(report["repositories"]), 7)
        self.assertTrue(all(repo["recorded_contract"] is None for repo in report["repositories"]))
        self.assertEqual(report["status"], "fail")

    def test_stale_never_passes(self) -> None:
        report = self.report("2026-09-20T12:00:00Z")
        repo = report["repositories"][0]
        self.assertTrue(all(item["status"] == "STALE" for item in repo["dimensions"]))
        self.assertTrue(all(item["status"] == "STALE" for item in repo["drift_findings"]))
        self.assertEqual(report["status"], "fail")

    def test_missing_future_and_boundary_timestamp(self) -> None:
        for timestamp in (None, "2026-09-10T12:00:00Z"):
            self.repo["quality"]["observed_at"] = timestamp
            self.assertEqual(self.dimension("adoption")["status"], "UNVERIFIED")
        self.repo["quality"]["observed_at"] = "2026-09-02T12:00:00Z"
        self.assertEqual(self.dimension("adoption")["status"], "PASS")

    def test_timezone_required(self) -> None:
        with self.assertRaises(ValueError):
            self.report("2026-09-09T12:00:00")

    def test_workflow_failure_skipped_and_pending(self) -> None:
        job = self.repo["quality"]["workflows"][0]["jobs"][0]
        for conclusion in ("failure", "cancelled", "skipped", "neutral", None):
            job["conclusion"] = conclusion
            self.assertEqual(self.dimension("required-workflow-health")["status"], "FAIL")
        job["status"] = "queued"
        self.assertEqual(self.dimension("required-workflow-health")["status"], "UNVERIFIED")

    def test_wrong_sha_branch_event_never_passes(self) -> None:
        for field, value in (("head_sha", "c" * 40), ("head_branch", "other"), ("event", "pull_request")):
            self.repo["quality"] = copy.deepcopy(sample()["repositories"][0]["quality"])
            self.repo["quality"]["workflows"][0][field] = value
            self.assertEqual(self.dimension("required-workflow-health")["status"], "UNVERIFIED")

    def test_latest_attempt_not_older_success(self) -> None:
        run = copy.deepcopy(self.repo["quality"]["workflows"][0])
        run["run_attempt"] = 2
        run["conclusion"] = "failure"
        self.repo["quality"]["workflows"].append(run)
        self.assertEqual(self.dimension("required-workflow-health")["status"], "FAIL")

    def test_ambiguous_job_name_not_pass(self) -> None:
        run = copy.deepcopy(self.repo["quality"]["workflows"][0])
        run.update(id=11, path=".github/workflows/other.yml")
        self.repo["quality"]["workflows"].append(run)
        self.assertEqual(self.dimension("required-workflow-health")["status"], "UNVERIFIED")

    def test_unknown_policy_vs_failed_control(self) -> None:
        self.repo["rulesets"] = None
        self.assertEqual(self.dimension("branch-protection")["status"], "UNVERIFIED")
        self.repo["rulesets"] = []
        self.assertEqual(self.dimension("branch-protection")["status"], "DRIFT")
        self.assertEqual(self.dimension("required-workflow-health")["status"], "UNVERIFIED")

    def test_provider_bound_and_ambiguous_scope_unknown(self) -> None:
        rule = self.repo["rulesets"][0]
        rule["rules"][-1]["parameters"]["required_status_checks"][0]["integration_id"] = 7
        self.assertIsNone(quality.required_checks(self.repo))
        rule["rules"][-1]["parameters"]["required_status_checks"][0].pop("integration_id")
        rule["conditions"]["ref_name"]["exclude"] = ["refs/heads/main"]
        self.assertIsNone(quality.required_checks(self.repo))

    def test_release_absent_unknown_and_observed(self) -> None:
        self.assertEqual(self.dimension("latest-stable-release")["status"], "UNRELEASED")
        self.repo["quality"]["releases"] = None
        self.assertEqual(self.dimension("latest-stable-release")["status"], "UNVERIFIED")
        self.repo["quality"]["releases"] = [{"id": 1, "tag_name": "v1.0.0", "published_at": NOW,
                                              "draft": False, "prerelease": False, "resolved_commit": "d" * 40}]
        row = self.dimension("latest-stable-release")
        self.assertEqual(row["status"], "OBSERVED")
        self.assertIn("d" * 40, row["detail"])
        self.assertIn("not certification", row["detail"])

    def test_unresolved_release_tag_not_certified(self) -> None:
        self.repo["quality"]["releases"] = [{"id": 1, "tag_name": "v1", "published_at": NOW, "draft": False, "prerelease": False}]
        self.assertEqual(self.dimension("latest-stable-release")["status"], "UNVERIFIED")

    def test_future_release_and_duplicate_attempt_rejected(self) -> None:
        self.repo["quality"]["releases"] = [{"id": 1, "tag_name": "v1", "published_at": "2026-09-10T12:00:00Z",
                                              "draft": False, "prerelease": False, "resolved_commit": "d" * 40}]
        self.assertEqual(self.dimension("latest-stable-release")["status"], "UNVERIFIED")
        self.repo["quality"]["workflows"].append(copy.deepcopy(self.repo["quality"]["workflows"][0]))
        with self.assertRaises(ValueError):
            self.report()

    def test_malformed_evidence_rejected(self) -> None:
        self.repo["quality"]["workflows"] = [None]
        with self.assertRaises(ValueError):
            self.report()

    def test_decisions_remain_visible(self) -> None:
        self.repo["decisions"] = [{"rule": "local-specialist", "decision": "KEEP", "reason": "Preserve local control",
                                  "commit": self.repo["commit"], "evidence_path": "docs/decision.md"}]
        markdown = quality.markdown_report(self.report())
        for value in ("KEEP", "NOT APPLICABLE", "Preserve local control"):
            self.assertIn(value, markdown)

    def test_specialist_claim_preserved_not_aggregated(self) -> None:
        self.repo["specialist"] = [{"name": "Numerical assurance", "score": 9.1, "scale": "0–10",
                                    "rationale": "External reviewer", "evidence_path": "docs/decision.md",
                                    "commit": self.repo["commit"], "observed_at": NOW}]
        claim = self.report()["repositories"][0]["specialist"][0]
        self.assertEqual((claim["score"], claim["status"]), (9.1, "REPORTED"))
        self.repo["specialist"][0]["commit"] = "d" * 40
        self.assertEqual(self.report()["repositories"][0]["specialist"][0]["status"], "UNVERIFIED")

    def test_popularity_never_changes_dimensions(self) -> None:
        baseline = self.report()["repositories"]
        self.repo["metadata"].update(stargazers_count=1000000, forks_count=50000, watchers_count=9999)
        self.assertEqual(baseline, self.report()["repositories"])

    def test_repeated_output_and_input_immutable(self) -> None:
        original = copy.deepcopy(self.data)
        first, second = self.report(), self.report()
        self.assertEqual(json.dumps(first, sort_keys=True), json.dumps(second, sort_keys=True))
        self.assertEqual(quality.markdown_report(first), quality.markdown_report(second))
        self.assertEqual(self.data, original)

    def test_no_empty_portfolio_or_output_collision(self) -> None:
        self.data["repositories"] = []
        with self.assertRaises(ValueError):
            self.report()
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            quality.main(["--snapshot", "a", "--output", "a", "--as-of", NOW])


class CollectionTests(unittest.TestCase):
    def test_actions_pagination_key_and_limit(self) -> None:
        with patch.object(capture, "get", return_value={"total_count": 1, "jobs": [{"id": 1}]}) as get:
            self.assertEqual(capture.pages("example/repo/jobs?filter=all", "jobs"), [{"id": 1}])
            self.assertIn("&per_page=", get.call_args.args[0])
        with patch.object(capture, "get", return_value={"total_count": 1000}), self.assertRaises(ValueError):
            capture.pages("example/repo/actions/runs", "workflow_runs")

    def test_annotated_tag_resolution(self) -> None:
        with patch.object(capture, "get", side_effect=[{"object": {"type": "tag", "sha": "b" * 40}},
                                                      {"object": {"type": "commit", "sha": "c" * 40}}]):
            self.assertEqual(capture.tag_commit("example/repo", "v1.0.0"), "c" * 40)

    def test_collector_records_exact_attempt(self) -> None:
        repo = sample()["repositories"][0]
        repo["captured_at"] = "2026-09-01T10:00:00Z"
        run = {**repo["quality"]["workflows"][0], "run_attempt": 3}
        with patch.object(capture, "pages", side_effect=[[run], run["jobs"], []]) as pages:
            result = capture.quality_evidence(repo)
        self.assertIn("/attempts/3/jobs", pages.call_args_list[1].args[0])
        self.assertEqual(result["workflows"][0]["run_attempt"], 3)
        self.assertEqual(result["releases"], [])
        self.assertEqual(result["observed_at"], repo["captured_at"])

    def test_capture_limits_are_not_empty_collections(self) -> None:
        repo = sample()["repositories"][0]
        with patch.object(capture, "pages", side_effect=ValueError("Incomplete")):
            result = capture.quality_evidence(repo)
        self.assertIsNone(result["workflows"])
        self.assertIsNone(result["releases"])


if __name__ == "__main__":
    unittest.main()
