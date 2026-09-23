"""Offline drift and simulated external HTTP failure boundaries."""

from __future__ import annotations

import hashlib
import json
import subprocess
import tempfile
import threading
import time
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

import check_documentation as docs
import check_external_links as links
import check_repo as repo_checks
import initialize_repository as initializer

ROOT = Path(__file__).resolve().parents[1]
TODAY = date(2026, 9, 9)


class ProjectIdentityTests(unittest.TestCase):
    """Exercise the adopted identity policy, including reinitialization safety."""

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        (self.root / ".github").mkdir()
        for path in (initializer.CONFIG_PATH, initializer.RECORD_PATH):
            (self.root / path).write_bytes((ROOT / path).read_bytes())
        self.config = json.loads((self.root / initializer.CONFIG_PATH).read_text())
        self.sample = self.root / "README.md"
        self.sample.write_text("K-PRICING retains KPR_Dates_AddDays from danielep71/KPR.\n")
        subprocess.run(["git", "init", "-q", str(self.root)], check=True)
        subprocess.run(["git", "-C", str(self.root), "add", "--all"], check=True)

    def scan(self):
        return repo_checks.check_identity(repo_checks.Repository(self.root), self.config)

    def test_kpr_namespace_and_provenance_are_allowed(self):
        self.assertEqual(self.scan()["status"], "pass")

    def test_old_product_branding_is_rejected(self):
        for text in ("# KPR\n", "# 📈 KPR\n", "KPR is the product name.\n"):
            with self.subTest(text=text):
                self.sample.write_text(text)
                self.assertEqual(self.scan()["status"], "fail")

    def test_unrelated_donor_and_template_identities_remain_rejected(self):
        policy = self.config["identity"]
        self.assertEqual(len(policy["forbidden_tokens"]), 6)
        self.assertEqual(len(policy["template_tokens"]), 1)
        self.assertEqual(set(policy["exclude_paths"]),
                         {initializer.CONFIG_PATH, initializer.RECORD_PATH,
                          "evidence/setup-2026-09-23/host.json"})
        for token in policy["forbidden_tokens"] + policy["template_tokens"]:
            with self.subTest(token=token):
                self.sample.write_text(f"Project identity: {token.lower()}\n")
                result = self.scan()
                self.assertEqual(result["status"], "fail")
                self.assertTrue(any(item["path"] == "README.md"
                                    for item in result["findings"]))

    def test_repeat_initialization_preserves_evolved_identity_policy(self):
        before = (self.root / initializer.CONFIG_PATH).read_bytes()
        profile, scalars, repeatable = initializer._record_arguments(self.root)
        changes, _ = initializer._build_changes(self.root, profile, scalars, repeatable)
        self.assertEqual(changes, {})
        initializer._apply_changes(self.root, changes)
        self.assertEqual((self.root / initializer.CONFIG_PATH).read_bytes(), before)
        self.assertEqual(self.scan()["status"], "pass")


class DocumentationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        (self.root / ".github/workflows").mkdir(parents=True)
        (self.root / "tools").mkdir()
        self.policy = json.loads((ROOT / docs.POLICY).read_text())
        self.policy["references"] = []
        self.policy["historical_documents"] = {}
        self.policy["network"]["domains"] = {"example.org": "Synthetic documentation service"}
        self.policy["network"]["classifications"] = []
        (self.root / "README.md").write_text("# Fixture\n\npython3 tools/fixture.py --root .\n")
        (self.root / "tools/fixture.py").write_text("import argparse\np=argparse.ArgumentParser()\np.add_argument('--root')\n")
        subprocess.run(["git", "init", "-q", str(self.root)], check=True)
        self.save()

    def save(self):
        (self.root / docs.POLICY).write_text(json.dumps(self.policy))
        subprocess.run(["git", "-C", str(self.root), "add", "--all"], check=True)

    def test_documented_command_passes_without_execution(self):
        with patch.object(subprocess, "Popen", wraps=subprocess.Popen) as popen:
            self.assertEqual(docs.build_report(self.root)["status"], "pass")
            self.assertTrue(all(call.args[0][0] == "git" for call in popen.call_args_list))

    def test_readme_presentation_uses_repository_identity_and_selected_assets(self):
        readme = (ROOT / "README.md").read_text(encoding="utf-8")
        profile = json.loads((ROOT / ".github/repository-profile.json").read_text(encoding="utf-8"))
        if profile.get("mode") == "generated":
            record = json.loads((ROOT / ".github/initialization.json").read_text(encoding="utf-8"))
            self.assertEqual([line for line in readme.splitlines() if line.startswith("# ")],
                             [f"# ⚡ {record['values']['PROJECT_NAME']}"])
            repository = profile["repository"]
            self.assertIn(f"https://github.com/{repository}/actions/workflows/static-checks.yml", readme)
            self.assertNotIn("{{", readme)
            self.assertNotIn("<!-- template:", readme)
            preview = record["values"].get("SOCIAL_PREVIEW_PATH")
            if preview:
                self.assertTrue((ROOT / preview).is_file())
                self.assertIn(f"<!-- generated-social-preview: {preview} -->", readme)
            else:
                self.assertNotIn('src="assets/social-preview.png"', readme)
            return
        repository_token = "{" + "{REPOSITORY_PATH}" + "}"
        preview_token = "{" + "{SOCIAL_PREVIEW_PATH}" + "}"
        self.assertNotIn(f"https://github.com/{repository_token}", readme)
        self.assertNotIn(
            f"https://api.scorecard.dev/projects/github.com/{repository_token}",
            readme,
        )
        self.assertNotIn(
            f"https://scorecard.dev/viewer/?uri=github.com/{repository_token}",
            readme,
        )
        self.assertNotIn("securityscorecards.dev", readme)
        canonical_repository = "danielep71/" + "EXCEL-VBA-" + "PROJECT-TEMPLATE"
        self.assertIn(
            f"https://api.scorecard.dev/projects/github.com/{canonical_repository}/badge",
            readme,
        )
        self.assertIn('src="assets/social-preview.png"', readme)
        self.assertIn(
            f"<!-- generated-social-preview: {preview_token} -->",
            readme,
        )

    def test_utf8_repository_reads_do_not_depend_on_locale(self):
        workflow = self.root / ".github/workflows/fixture.yml"
        workflow.write_text(
            "name: UTF-8 workflow 🔐\njobs:\n  check:\n    name: UTF-8 context\n    runs-on: ubuntu-24.04\n",
            encoding="utf-8",
        )
        (self.root / "README.md").write_text(
            "# Fixture 🔐\n\nUTF-8 workflow 🔐\n\npython3 tools/fixture.py --root .\n",
            encoding="utf-8",
        )
        self.policy["references"] = [{
            "document": "README.md",
            "target": ".github/workflows/fixture.yml",
            "token": "UTF-8 workflow 🔐",
            "kind": "workflow-name",
        }]
        self.save()
        original = Path.read_text

        def require_explicit_utf8(path, *args, **kwargs):
            encoding = kwargs.get("encoding")
            if encoding is None and args:
                encoding = args[0]
            if encoding is None:
                raise UnicodeDecodeError(
                    "charmap", b"\x8f", 0, 1, "character maps to <undefined>"
                )
            self.assertEqual(encoding, "utf-8")
            return original(path, *args, **kwargs)

        with patch.object(Path, "read_text", require_explicit_utf8):
            self.assertEqual(docs.build_report(self.root)["status"], "pass")

    def test_renamed_command_detected(self):
        (self.root / "tools/fixture.py").rename(self.root / "tools/renamed.py")
        self.save()
        self.assertEqual(docs.build_report(self.root)["status"], "fail")

    def test_removed_cli_option_detected(self):
        (self.root / "tools/fixture.py").write_text("import argparse\n")
        self.assertEqual(docs.build_report(self.root)["status"], "fail")

    def test_shared_runner_does_not_grant_parser_flags(self):
        (self.root / "tools/fixture.py").write_text("from _gatelib import run_gate\n")
        (self.root / "tools/_gatelib.py").write_text("p.add_argument('--self-test')\n")
        self.assertNotIn("--self-test", docs.command_flags(self.root, "tools/fixture.py"))
        (self.root / "tools/fixture.py").write_text("from _gatelib import parse_report_args\n")
        self.assertIn("--self-test", docs.command_flags(self.root, "tools/fixture.py"))

    def test_multiline_and_inline_commands(self):
        text = "`python3 tools/a.py --root .`\npython3 tools/b.py \\\n  --output out.json\n"
        self.assertEqual(docs.commands(text), [("tools/a.py", {"--root"}), ("tools/b.py", {"--output"})])

    def test_workflow_and_context_renames_detected(self):
        workflow = self.root / ".github/workflows/fixture.yml"
        workflow.write_text("name: Example workflow\njobs:\n  check:\n    name: Example context\n    runs-on: ubuntu-24.04\n")
        (self.root / "README.md").write_text("Example workflow; Example context\n")
        self.policy["references"] = [
            {"document": "README.md", "target": ".github/workflows/fixture.yml", "token": "Example workflow", "kind": "workflow-name"},
            {"document": "README.md", "target": ".github/workflows/fixture.yml", "token": "Example context", "kind": "job-name", "job": "check"}]
        self.save()
        self.assertEqual(docs.build_report(self.root)["status"], "pass")
        workflow.write_text(workflow.read_text().replace("Example context", "Changed context"))
        self.assertEqual(docs.build_report(self.root)["status"], "fail")
        workflow.write_text(workflow.read_text().replace("Example workflow", "Changed workflow"))
        self.assertEqual(len(docs.build_report(self.root)["findings"]), 2)

    def test_filename_and_policy_value_drift(self):
        (self.root / "value.json").write_text('{"required": true}')
        (self.root / "README.md").write_text("Required value.json\n")
        self.policy["references"] = [{"document": "README.md", "target": "value.json", "token": "Required",
                                     "kind": "json-value", "pointer": "/required", "value": True}]
        self.save()
        self.assertEqual(docs.build_report(self.root)["status"], "pass")
        (self.root / "value.json").write_text('{"required": false}')
        self.assertEqual(docs.build_report(self.root)["status"], "fail")
        (self.root / "value.json").unlink()
        self.save()
        self.assertEqual(docs.build_report(self.root)["status"], "fail")

    def probe(self, responses):
        calls = []
        sleeps = []
        def transport(url, timeout):
            calls.append((url, timeout))
            value = responses[len(calls) - 1]
            if isinstance(value, Exception):
                raise value
            return value
        report = links.probe("https://example.org/page", self.policy["network"], transport, sleeps.append)
        return report, calls, sleeps

    def test_consistent_404_retried_and_reported(self):
        report, calls, sleeps = self.probe([(404, None)] * 3)
        self.assertEqual(report["status"], "PERMANENT_FAILURE")
        self.assertEqual(len(calls), 3)
        self.assertEqual(sleeps, [1, 2])

    def test_public_404_remains_deterministic_public_defect(self):
        url = "https://example.org/missing"
        (self.root / "README.md").write_text(f"[missing]({url})\n")
        self.save()
        report = links.build_report(self.root, TODAY, lambda *args: (404, None), lambda *args: None)
        self.assertEqual(report["status"], "fail")
        self.assertEqual(report["links"][0]["status"], "PERMANENT_FAILURE")
        self.assertEqual(report["counts"]["deterministic_public_defects"], 1)
        self.assertEqual(report["counts"]["restricted_historical"], 0)

    def test_restricted_historical_is_non_green_and_not_probed(self):
        url = "https://example.org/private-history"
        identifier = hashlib.sha256(url.encode()).hexdigest()
        (self.root / "README.md").write_text(f"[history]({url})\n")
        self.policy["network"]["classifications"] = [{
            "id": identifier,
            "kind": "restricted-historical",
            "reason": "Authenticated historical evidence",
            "expires": "2026-09-10",
        }]
        self.save()
        report = links.build_report(self.root, TODAY, lambda *args: self.fail("classified target must not be probed"))
        self.assertEqual(report["status"], "fail")
        self.assertEqual(report["links"][0]["status"], "RESTRICTED_HISTORICAL")
        self.assertEqual(report["counts"]["restricted_historical"], 1)
        self.assertEqual(report["counts"]["deterministic_public_defects"], 0)
        self.assertNotIn("private-history", json.dumps(report) + links.markdown(report))

    def test_current_private_target_is_non_green_without_anonymous_probe(self):
        url = "https://example.org/private-project"
        identifier = hashlib.sha256(url.encode()).hexdigest()
        (self.root / "README.md").write_text(f"[project]({url})\n")
        self.policy["network"]["classifications"] = [{
            "id": identifier,
            "kind": "access-restricted",
            "reason": "Current private repository requires authenticated verification",
            "expires": "2026-09-10",
        }]
        self.save()
        report = links.build_report(self.root, TODAY, lambda *args: self.fail("private target must not be probed"))
        self.assertEqual(report["status"], "fail")
        self.assertEqual(report["links"][0]["status"], "ACCESS_RESTRICTED")
        self.assertEqual(report["links"][0]["attempts"], 0)
        self.assertEqual(report["counts"]["access_restricted"], 1)
        self.assertEqual(report["counts"]["deterministic_public_defects"], 0)
        self.assertNotIn("private-project", json.dumps(report) + links.markdown(report))

    def test_pending_publication_is_distinct_non_green_classification(self):
        url = "https://example.org/compare/v1.0.0...v1.1.0"
        identifier = hashlib.sha256(url.encode()).hexdigest()
        (self.root / "README.md").write_text(f"[future-tag]({url})\n")
        self.policy["network"]["classifications"] = [{
            "id": identifier,
            "kind": "pending-publication",
            "reason": "Reviewed candidate link depends on a tag not published yet",
            "expires": "2026-09-10",
        }]
        self.save()
        report = links.build_report(self.root, TODAY, lambda *args: self.fail("pending target must not be probed"))
        self.assertEqual(report["status"], "fail")
        self.assertEqual(report["links"][0]["status"], "PENDING_PUBLICATION")
        self.assertEqual(report["counts"]["pending_publication"], 1)
        self.assertEqual(report["counts"]["deterministic_public_defects"], 0)

    def test_classification_cannot_override_local_url_policy(self):
        cases = (
            ("http://example.org/private-history", "POLICY_BLOCKED"),
            ("https://example.org/private-history?token=fixture", "ACCESS_RESTRICTED"),
        )
        for url, expected in cases:
            with self.subTest(url=url):
                identifier = hashlib.sha256(url.encode()).hexdigest()
                (self.root / "README.md").write_text(f"[classified]({url})\n")
                self.policy["network"]["classifications"] = [{
                    "id": identifier,
                    "kind": "restricted-historical",
                    "reason": "Reviewed historical classification",
                    "expires": "2026-09-10",
                }]
                self.save()
                report = links.build_report(
                    self.root,
                    TODAY,
                    lambda *args: self.fail("policy-rejected classified target must not be probed"),
                )
                self.assertEqual(report["links"][0]["status"], expected)
                self.assertEqual(report["links"][0]["attempts"], 0)
                self.assertEqual(report["status"], "fail")

    def test_markdown_exposes_every_json_count_category(self):
        rows = [
            {"status": "PERMANENT_FAILURE"},
            {"status": "RESTRICTED_HISTORICAL"},
            {"status": "PENDING_PUBLICATION"},
            {"status": "ACCESS_RESTRICTED"},
            {"status": "TRANSIENT_FAILURE"},
        ]
        counts = links._counts(rows)
        self.assertEqual(set(counts), set(links.COUNT_LABELS))
        report = {
            "status": "fail",
            "discovered": len(rows),
            "limit_exceeded": False,
            "counts": counts,
            "links": [],
            "scope_note": "fixture scope",
        }
        rendered = links.markdown(report)
        for key, label in links.COUNT_LABELS.items():
            self.assertIn(f"{label}: {counts[key]}", rendered)

    def test_transient_then_recovery(self):
        report, calls, _ = self.probe([(503, None), (200, None)])
        self.assertEqual(report["status"], "OK")
        self.assertEqual(len(calls), 2)

    def test_mixed_missing_transient_not_permanent(self):
        report, _, _ = self.probe([(404, None), TimeoutError("secret URL"), (404, None)])
        self.assertEqual(report["status"], "TRANSIENT_FAILURE")
        self.assertNotIn("secret", json.dumps(report))

    def test_access_denied_not_retried(self):
        report, calls, _ = self.probe([(403, None)])
        self.assertEqual(report["status"], "ACCESS_RESTRICTED")
        self.assertEqual(len(calls), 1)

    def test_redirect_checked_and_loop_bounded(self):
        calls = []
        def redirect(url, timeout):
            calls.append(url)
            return 302, "https://127.0.0.1/private"
        self.assertEqual(links.probe("https://example.org", self.policy["network"], redirect)["status"], "POLICY_BLOCKED")
        self.assertEqual(len(calls), 1)
        report = links.probe("https://example.org/page", self.policy["network"], lambda *args: (302, "/page"))
        self.assertEqual(report["status"], "REDIRECT_FAILURE")

    def test_redirect_queries_not_transmitted(self):
        calls = []
        def transport(url, timeout):
            calls.append(url)
            return 302, "/login?token=secret"
        report = links.probe("https://example.org", self.policy["network"], transport)
        self.assertEqual(report["status"], "ACCESS_RESTRICTED")
        self.assertEqual(len(calls), 1)
        self.assertNotIn("secret", json.dumps(report))

    def test_query_credentials_scheme_domain_not_requested(self):
        for url in ("https://example.org/?token=secret", "https://user:secret@example.org/",
                    "http://example.org", "https://other.example/page", "ftp://example.org"):
            with self.subTest(url=url), patch.object(links, "request") as transport:
                report = links.probe(url, self.policy["network"], transport)
                transport.assert_not_called()
                self.assertNotEqual(report["status"], "OK")

    def test_private_dns_never_connects(self):
        with patch.object(links.socket, "getaddrinfo", return_value=[(2, 1, 6, "", ("127.0.0.1", 443))]), \
                patch.object(links, "PinnedHTTPS") as connection:
            with self.assertRaises(ValueError):
                links.request("https://example.org", 2)
            connection.assert_not_called()

    def test_tls_connect_uses_checked_ip(self):
        with patch.object(links.socket, "create_connection") as connect:
            client = links.PinnedHTTPS("example.org", "93.184.215.14", 2)
            with patch.object(client.tls_context, "wrap_socket") as wrap:
                client.connect()
                connect.assert_called_once_with(("93.184.215.14", 443), 2)
                wrap.assert_called_once_with(connect.return_value, server_hostname="example.org")

    def test_exception_expiry_reason_and_limits(self):
        network = self.policy["network"]
        network["exceptions"] = [{"id": "a" * 64, "reason": "Temporary service outage", "expires": "2026-09-10"}]
        links.validate_policy(network, TODAY)
        with self.assertRaises(ValueError):
            links.validate_policy(network, date(2026, 9, 11))
        network["exceptions"] = []
        network["attempts"] = 100
        with self.assertRaises(ValueError):
            links.validate_policy(network, TODAY)

    def test_classification_expiry_kind_and_identity_fail_closed(self):
        network = self.policy["network"]
        network["classifications"] = [{
            "id": "b" * 64,
            "kind": "restricted-historical",
            "reason": "Reviewed private evidence",
            "expires": "2026-09-08",
        }]
        with self.assertRaises(ValueError):
            links.validate_policy(network, TODAY)
        network["classifications"][0]["expires"] = "2026-09-10"
        network["classifications"][0]["kind"] = "private-maybe"
        with self.assertRaises(ValueError):
            links.validate_policy(network, TODAY)
        network["classifications"][0]["kind"] = "restricted-historical"
        network["exceptions"] = [{
            "id": "b" * 64,
            "reason": "Conflicting exception",
            "expires": "2026-09-10",
        }]
        with self.assertRaises(ValueError):
            links.validate_policy(network, TODAY)

    def test_reports_redact_urls_and_deduplicate(self):
        url = "https://example.org/private?token=TOP_SECRET"
        (self.root / "README.md").write_text(f"[one]({url})\n[two]({url})\n")
        report = links.build_report(self.root, TODAY, lambda *args: self.fail("must not request"))
        self.assertEqual(report["discovered"], 1)
        self.assertEqual(len(report["links"][0]["locations"]), 2)
        rendered = json.dumps(report) + links.markdown(report)
        self.assertNotIn("TOP_SECRET", rendered)
        self.assertNotIn("/private", rendered)

    def test_active_exception_is_explicit(self):
        url = "https://example.org/page"
        (self.root / "README.md").write_text(f"[page]({url})")
        self.policy["network"]["exceptions"] = [{"id": hashlib.sha256(url.encode()).hexdigest(), "reason": "Reviewed temporary gap", "expires": "2026-09-10"}]
        self.save()
        report = links.build_report(self.root, TODAY, lambda *args: self.fail("excepted"))
        self.assertEqual(report["links"][0]["status"], "EXCEPTED")

    def test_total_limit_cannot_claim_complete(self):
        (self.root / "README.md").write_text("[a](https://example.org/a)\n[b](https://example.org/b)")
        self.policy["network"]["max_links"] = 1
        self.save()
        report = links.build_report(self.root, TODAY, lambda *args: (200, None))
        self.assertTrue(report["limit_exceeded"])
        self.assertEqual(report["status"], "fail")

    def test_worker_concurrency_bounded(self):
        (self.root / "README.md").write_text("\n".join(f"[page](https://example.org/{i})" for i in range(12)))
        self.policy["network"]["concurrency"] = 2
        self.save()
        active = peak = 0
        lock = threading.Lock()
        def transport(*args):
            nonlocal active, peak
            with lock:
                active += 1
                peak = max(peak, active)
            time.sleep(0.01)
            with lock:
                active -= 1
            return 200, None
        self.assertEqual(links.build_report(self.root, TODAY, transport)["status"], "pass")
        self.assertLessEqual(peak, 2)


if __name__ == "__main__":
    unittest.main()
