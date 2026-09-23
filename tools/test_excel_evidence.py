"""Synthetic host-interface tests; none executes Office."""

from __future__ import annotations

import contextlib
import copy
import hashlib
import io
import json
import tempfile
import unittest
from pathlib import Path

import check_excel_evidence as host
import check_release as release

ROOT = Path(__file__).resolve().parents[1]


class HostEvidenceTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.area = Path(self.temporary.name)
        self.root = self.area / "candidate"
        self.bundle = self.area / "bundle"
        self.bundle.mkdir()
        self.path = self.bundle / "host.json"
        self.release_policy = json.loads((ROOT / release.POLICY_PATH).read_text())
        release._fixture_repository(self.root, "library", self.release_policy)
        config = release._fixture_configuration("library")
        config["template_contract"] = {"version": "1.2.0", "source": "example/template"}
        config["vba"] = {"components": {"src/Project.bas": "public"}}
        (self.root / release.PROFILE_PATH).write_text(json.dumps(config))
        self.policy = json.loads((ROOT / host.POLICY).read_text())
        (self.root / host.POLICY).write_text(json.dumps(self.policy))
        release._git(self.root, "add", "--all")
        release._git(self.root, "commit", "-m", "Synthetic host policy")
        self.sha = release._git_output(self.root, "rev-parse", "HEAD")
        self.log = "\n".join([
            "PROJECT TESTS", *("CASE=" + case for case in self.policy["cases"]),
            "CASES=4", "ASSERTIONS=6", "FAILURES=0", "CLEANUP=PASS; detail=synthetic state restored",
            "RESULT=PASS; completeness=COMPLETE; cases=4; assertions=6; failures=0; cleanup=PASS", "",
        ])
        (self.bundle / "host.log").write_text(self.log)
        log_ref = {"path": "host.log", "sha256": hashlib.sha256(self.log.encode()).hexdigest()}
        self.record = {
            "schema_version": 1, "repository": config["repository"], "candidate_sha": self.sha,
            "template_contract": config["template_contract"], "execution": "manual",
            "availability_reason": None, "started_at": "2026-09-09T09:00:00Z",
            "finished_at": "2026-09-09T09:01:00Z",
            "runner": {"class": "manual-interactive", "identity": "Synthetic operator/workstation", "workflow": None},
            "environment": {"excel_version": "16.0", "excel_build": "12345.67890",
                            "office_bitness": "64-bit", "os": "Windows 10", "os_architecture": "x64",
                            "runtime": "VBA7+", "macro_policy": "approved signed macros",
                            "vba_project_access": "disabled; manual import", "trust_changes": False},
            "sources": host.source_inventory(self.root, self.sha, config),
            "stages": {name: {"status": "PASS", "detail": "Synthetic assertion", "log": copy.deepcopy(log_ref)}
                       for name in host.STAGES},
            "harness": {"entry_point": self.policy["entry_point"], "cases": 4, "assertions": 6,
                        "failures": 0, "completeness": "COMPLETE", "expected_errors": [
                            {"case": "ratio.zero-denominator", "status": "PASS",
                             "detail": "Number/source/description assertions passed in complete synthetic suite"}]},
        }

    def evaluate(self):
        self.path.write_text(json.dumps(self.record))
        return host.evaluate(self.root, self.sha, self.path)

    def invalid(self):
        report = self.evaluate()
        self.assertIn("EVIDENCE_INVALID", report["outcomes"], report)

    def test_manual_pass_repeatability(self):
        first = self.evaluate()
        self.assertEqual(first["status"], "pass", first)
        self.assertEqual(first, self.evaluate())
        self.assertEqual(first["execution"], "manual")

    def test_automated_identity(self):
        self.record["execution"] = "automated"
        self.record["runner"] = {"class": "trusted-interactive", "identity": "Synthetic isolated desktop",
                                 "workflow": {"repository": "example/validation", "path": ".github/workflows/host.yml",
                                              "sha": "a" * 40, "run_id": 12, "run_attempt": 1}}
        self.assertEqual(self.evaluate()["status"], "pass")
        self.record["runner"]["workflow"]["sha"] = "main"
        self.invalid()

    def test_manual_cannot_claim_hosted_execution(self):
        self.record["runner"]["workflow"] = {"run_id": 123}
        self.invalid()
        self.record["runner"]["workflow"] = None
        self.record["runner"]["class"] = "trusted-interactive"
        self.invalid()

    def test_unavailable_is_non_green(self):
        self.record.update(execution="unavailable", availability_reason="No eligible trusted host",
                           runner=None, environment=None, harness=None, sources=[], stages={})
        report = self.evaluate()
        self.assertEqual(report["status"], "fail")
        self.assertEqual(report["outcomes"], ["UNAVAILABLE"])
        self.record["harness"] = {"cases": 4}
        self.invalid()

    def test_compile_failure_distinct(self):
        self.record["stages"]["compile"]["status"] = "FAIL"
        self.record["stages"]["regression"] = {"status": "NOT_RUN", "detail": "Compile failed", "log": None}
        self.record["harness"] = None
        self.assertEqual(self.evaluate()["outcomes"], ["COMPILE_FAILED", "INCOMPLETE"])

    def test_test_failure_distinct(self):
        self.record["stages"]["regression"]["status"] = "FAIL"
        self.record["harness"]["failures"] = 1
        self.assertEqual(self.evaluate()["outcomes"], ["TEST_FAILED"])

    def test_cleanup_failure_distinct(self):
        self.record["stages"]["cleanup"]["status"] = "FAIL"
        self.assertEqual(self.evaluate()["outcomes"], ["CLEANUP_FAILED"])

    def test_internal_harness_cleanup_failure(self):
        self.record["stages"]["cleanup"]["status"] = "FAIL"
        failed = self.log.replace("RESULT=PASS", "RESULT=FAIL").replace("cleanup=PASS", "cleanup=FAIL")
        (self.bundle / "failed.log").write_text(failed)
        self.record["stages"]["regression"]["log"] = {
            "path": "failed.log", "sha256": hashlib.sha256(failed.encode()).hexdigest()}
        self.assertEqual(self.evaluate()["outcomes"], ["CLEANUP_FAILED"])

    def test_import_failure_and_timeout(self):
        for status, outcome in (("FAIL", "IMPORT_FAILED"), ("TIMEOUT", "EXECUTION_TIMEOUT")):
            self.record["stages"]["import"]["status"] = status
            for name in ("compile", "regression"):
                self.record["stages"][name] = {"status": "NOT_RUN", "detail": "Import did not complete", "log": None}
            self.record["harness"] = None
            self.assertEqual(self.evaluate()["outcomes"], [outcome, "INCOMPLETE"])

    def test_sha_and_source_binding(self):
        original = copy.deepcopy(self.record)
        for field, value in (("candidate_sha", "f" * 40), ("template_contract", {}),
                             ("repository", "wrong/repo"), ("sources", [])):
            self.record = copy.deepcopy(original)
            self.record[field] = value
            self.invalid()

    def test_environment_and_trust_required(self):
        original = copy.deepcopy(self.record)
        for key in self.record["environment"]:
            self.record = copy.deepcopy(original)
            self.record["environment"][key] = True if key == "trust_changes" else ""
            self.invalid()

    def test_counts_and_expected_errors(self):
        original = copy.deepcopy(self.record)
        for field, value in (("cases", 3), ("assertions", True), ("failures", 1),
                             ("completeness", "INCOMPLETE"), ("expected_errors", [])):
            self.record = copy.deepcopy(original)
            self.record["harness"][field] = value
            self.invalid()
        self.record = copy.deepcopy(original)
        self.record["harness"]["expected_errors"][0]["status"] = "FAIL"
        self.invalid()

    def test_raw_log_tamper_missing_and_traversal(self):
        (self.bundle / "host.log").write_text("changed")
        self.invalid()
        (self.bundle / "host.log").unlink()
        self.invalid()
        self.record["stages"]["import"]["log"]["path"] = "../outside.log"
        self.invalid()

    def test_duplicate_or_incomplete_raw_summary(self):
        for raw in (self.log + self.log, self.log.replace("ASSERTIONS=6", "ASSERTIONS=5"),
                    self.log.replace("CASE=ratio.zero-denominator\n", "")):
            (self.bundle / "host.log").write_text(raw)
            for stage in self.record["stages"].values():
                stage["log"]["sha256"] = hashlib.sha256(raw.encode()).hexdigest()
            self.invalid()

    def test_release_binding_and_omission(self):
        self.evaluate()
        evidence = release._fixture_evidence("library", self.sha, self.release_policy)
        for key in ("vba-compile", "regression"):
            evidence["checks"][key]["environment"] = host.environment_summary(self.record)
        evidence["checks"]["excel-host-evidence"] = {
            "status": "PASS", "candidate_sha": self.sha, "detail": "Synthetic manual evidence",
            "sha256": hashlib.sha256(self.path.read_bytes()).hexdigest(), "execution": "manual"}
        path = self.area / "release.json"
        path.write_text(json.dumps(evidence))
        self.assertEqual(host.release_findings(self.root, self.sha, path, self.path), [])
        self.assertTrue(host.release_findings(self.root, self.sha, path, None))
        evidence["checks"]["excel-host-evidence"]["execution"] = "automated"
        path.write_text(json.dumps(evidence))
        self.assertTrue(host.release_findings(self.root, self.sha, path, self.path))

    def test_cli_unavailable_and_output_guard(self):
        self.record.update(execution="unavailable", availability_reason="No eligible trusted host",
                           runner=None, environment=None, harness=None, sources=[], stages={})
        self.evaluate()
        args = ["--root", str(self.root), "--candidate-sha", self.sha, "--evidence", str(self.path)]
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(host.main(args), 1)
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            host.main(args + ["--output", str(self.path)])

    def test_dirty_or_wrong_checkout(self):
        (self.root / "src/Project.bas").write_text("changed")
        self.invalid()
        release._git(self.root, "checkout", "--", "src/Project.bas")
        old = release._git_output(self.root, "rev-parse", "HEAD^")
        self.path.write_text(json.dumps(self.record))
        report = host.evaluate(self.root, old, self.path)
        self.assertIn("EVIDENCE_INVALID", report["outcomes"])

    def test_compile_failure_cannot_claim_test_execution(self):
        self.record["stages"]["compile"]["status"] = "FAIL"
        self.invalid()

    def test_symlinked_logs_rejected(self):
        original = self.bundle / "host.log"
        moved = self.area / "outside.log"
        original.rename(moved)
        original.symlink_to(moved)
        self.invalid()

    def test_timestamps_require_order_and_timezone(self):
        self.record["finished_at"] = "2026-09-09T08:00:00Z"
        self.invalid()
        self.record["finished_at"] = "2026-09-09T09:01:00"
        self.invalid()


if __name__ == "__main__":
    unittest.main()
