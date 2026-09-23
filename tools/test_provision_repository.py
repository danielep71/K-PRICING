"""Stateful simulated GitHub provisioning: no live mutation or credential required."""
from __future__ import annotations

import base64
import contextlib
import copy
import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from urllib.parse import unquote

import provision_repository as provision

SHA = "a" * 40
ROOT = Path(__file__).resolve().parents[1]


class FakeGitHub:
    def __init__(self, profile: str = "library"):
        self.repository = "example/generated"
        self.writes: list[dict] = []
        self.calls: list[tuple] = []
        self.fail_at = 0
        self.ignore = False
        policy = json.loads((ROOT / provision.POLICY).read_text())
        policy["description"] = "Example generated VBA project"
        contract = {"version": "1.2.0", "source": "example/template"}
        config = {"mode": "generated", "profile": profile, "repository": self.repository,
                  "template_contract": contract, "label_domains": []}
        record = {"schema_version": 1, "profile": profile, "template_contract": contract,
                  "values": {"REPOSITORY_PATH": self.repository}}
        self.files = dict(zip(provision.FILES, [config, record, policy,
                          json.loads((ROOT / ".github/labels.json").read_text())]))
        self.metadata = {**policy["features"], "id": 1, "full_name": self.repository,
                         "default_branch": "main", "description": None}
        self.topics: list[str] = []
        self.labels: list[dict] = []
        self.rulesets: list[dict] = []

    def request(self, method: str, path: str = "", body: dict | None = None) -> object:
        self.calls.append((method, path))
        if method == "GET":
            return copy.deepcopy(self.read(path))
        body = copy.deepcopy(body or {})
        self.writes.append({"method": method, "path": path, "body": body})
        if len(self.writes) == self.fail_at:
            raise OSError("Simulated connection loss; result uncertain")
        if self.ignore:
            return {}
        if path == "" and method == "PATCH":
            self.metadata.update(body)
        elif path == "/topics" and method == "PUT":
            self.topics = body["names"]
        elif path == "/labels" and method == "POST":
            self.labels.append(body)
        elif path.startswith("/labels/") and method == "PATCH":
            label = next(row for row in self.labels if row["name"] == unquote(path[8:]))
            label.update(name=body["new_name"], color=body["color"], description=body["description"])
        elif path == "/rulesets" and method == "POST":
            self.rulesets.append({"id": len(self.rulesets) + 1, **body})
        else:
            raise AssertionError((method, path))
        return {}

    def read(self, path: str) -> object:
        if path == "":
            return self.metadata
        if path == "/commits/main":
            return {"sha": SHA}
        if path.startswith("/contents/"):
            name = path[len("/contents/"):].split("?")[0]
            return {"encoding": "base64", "type": "file",
                    "content": base64.b64encode(json.dumps(self.files[name]).encode()).decode()}
        if path == "/topics":
            return {"names": self.topics}
        if path.startswith("/labels?"):
            return self.labels
        if path.startswith("/rulesets?"):
            return self.rulesets
        if path.startswith("/rulesets/"):
            row = copy.deepcopy(next(row for row in self.rulesets if row["id"] == int(path.rsplit("/", 1)[1])))
            for rule in row.get("rules", []):
                if rule.get("type") == "update" and rule.get("parameters") == {"update_allows_fetch_and_merge": False}:
                    rule.pop("parameters")
            return row
        raise AssertionError(path)


class ProvisionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.api = FakeGitHub()

    def plan(self) -> dict:
        return provision.execute(self.api, SHA, "library", "1.2.0")["plan"]

    def apply(self, path: Path, plan: dict | None = None) -> dict:
        return provision.execute(self.api, SHA, "library", "1.2.0", (plan or self.plan())["plan_sha256"], path)

    def test_plan_get_only_deterministic_complete(self) -> None:
        first, second = self.plan(), self.plan()
        self.assertEqual(first, second)
        self.assertFalse(self.api.writes)
        self.assertTrue(all(method == "GET" for method, _ in self.api.calls))
        self.assertEqual({action["path"] for action in first["actions"]}, {"", "/topics", "/labels", "/rulesets"})

    def test_description_from_real_initializer_record_shape(self) -> None:
        self.api.files[provision.POLICY]["description"] = "@initialization:PROJECT_DESCRIPTION"
        self.api.files[provision.FILES[1]]["values"]["PROJECT_DESCRIPTION"] = "Recorded project purpose"
        action = next(action for action in self.plan()["actions"] if action["path"] == "")
        self.assertEqual(action["body"]["description"], "Recorded project purpose")
        self.assertEqual(self.api.files[provision.POLICY]["description"], "@initialization:PROJECT_DESCRIPTION")

    def test_concurrent_change_between_writes_stops_apply(self) -> None:
        original = provision.capture
        calls = 0

        def changing(api: object, sha: str) -> dict:
            nonlocal calls
            calls += 1
            if calls == 4:
                self.api.topics.append("concurrent-topic")
            return original(api, sha)

        plan = self.plan()
        with tempfile.TemporaryDirectory() as directory, patch.object(provision, "capture", side_effect=changing):
            result = self.apply(Path(directory) / "journal.json", plan)
        self.assertEqual(result["status"], "fail")
        self.assertEqual(len(self.api.writes), 1)

    def test_retained_journal_not_overwritten(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "journal.json"
            provision.journal(path, {"retained": True})
            with self.assertRaises(ValueError):
                self.apply(path)
            self.assertEqual(json.loads(path.read_text()), {"retained": True})
        self.assertFalse(self.api.writes)

    def test_three_profiles_apply_and_second_apply_noop(self) -> None:
        for profile in sorted(provision.PROFILES):
            with self.subTest(profile=profile), tempfile.TemporaryDirectory() as directory:
                api = FakeGitHub(profile)
                first = provision.execute(api, SHA, profile, "1.2.0")["plan"]
                result = provision.execute(api, SHA, profile, "1.2.0", first["plan_sha256"], Path(directory) / "journal.json")
                self.assertEqual(result["status"], "pass")
                self.assertEqual(result["verification"]["remaining_actions"], [])
                second = provision.execute(api, SHA, profile, "1.2.0")["plan"]
                count = len(api.writes)
                again = provision.execute(api, SHA, profile, "1.2.0", second["plan_sha256"], Path(directory) / "second.json")
                self.assertEqual(again["status"], "pass")
                self.assertEqual(count, len(api.writes))
                self.assertIn(profile, api.topics)

    def test_update_rule_readback_omission_is_restrictive(self) -> None:
        desired = provision.baseline_rules(self.api.files[provision.POLICY])[1]
        actual = copy.deepcopy(desired)
        update = next(rule for rule in actual["rules"] if rule["type"] == "update")
        update.pop("parameters")
        self.assertTrue(provision.covered([{"id": 1, **actual}], desired, "main"))
        update["parameters"] = {"update_allows_fetch_and_merge": True}
        self.assertFalse(provision.covered([{"id": 1, **actual}], desired, "main"))

    def test_stale_approval_refused_before_write(self) -> None:
        plan = self.plan()
        self.api.topics.append("new-local-topic")
        with tempfile.TemporaryDirectory() as directory, self.assertRaises(ValueError):
            self.apply(Path(directory) / "journal.json", plan)
        self.assertFalse(self.api.writes)

    def test_tampered_plan_digest_refused(self) -> None:
        with tempfile.TemporaryDirectory() as directory, self.assertRaises(ValueError):
            self.apply(Path(directory) / "journal.json", {"plan_sha256": "0" * 64})
        self.assertFalse(self.api.writes)

    def test_mismatched_profile_identity_contract_and_head(self) -> None:
        for key, value in (("repository", "other/repo"), ("profile", "application"), ("mode", "template")):
            self.api = FakeGitHub()
            self.api.files[provision.CONFIG][key] = value
            with self.assertRaises(ValueError):
                self.plan()
        self.api = FakeGitHub()
        with self.assertRaises(ValueError):
            provision.execute(self.api, "b" * 40, "library", "1.2.0")
        with self.assertRaises(ValueError):
            provision.execute(self.api, SHA, "library", "1.1.0")
        self.assertFalse(self.api.writes)

    def test_stronger_controls_extra_labels_topics_retained(self) -> None:
        self.api.topics = ["specialist"]
        self.api.labels = [{"name": "local", "color": "FFFFFF", "description": "Local assurance"}]
        self.api.metadata["allow_rebase_merge"] = False
        self.api.rulesets = [{"id": index + 1, **row} for index, row in enumerate(provision.baseline_rules(self.api.files[provision.POLICY]))]
        self.api.rulesets[0]["rules"][2]["parameters"]["required_approving_review_count"] = 2
        plan = self.plan()
        self.assertFalse(any(action["path"] == "/rulesets" for action in plan["actions"]))
        self.assertFalse(any("allow_rebase_merge" in action["body"] for action in plan["actions"]))
        self.assertTrue(any("Extra label retained" in item for item in plan["keep"]))
        self.assertIn("specialist", next(action["body"]["names"] for action in plan["actions"] if action["path"] == "/topics"))

    def test_feature_disable_needs_recorded_exception(self) -> None:
        self.api.metadata["has_wiki"] = True
        self.assertTrue(self.plan()["blocked"])
        self.api.files[provision.POLICY]["exceptions"]["feature:has_wiki"] = "Reviewed unused generated wiki; no owned content"
        self.assertFalse(self.plan()["blocked"])

    def test_existing_label_needs_exception_and_correct_api_shape(self) -> None:
        self.api.labels = [{"name": "bug", "color": "FFFFFF", "description": "Old"}]
        self.assertTrue(self.plan()["blocked"])
        self.api.files[provision.POLICY]["exceptions"]["label:bug"] = "Reviewed default GitHub label replacement"
        action = next(action for action in self.plan()["actions"] if action["method"] == "PATCH" and action["path"] == "/labels/bug")
        self.assertIn("new_name", action["body"])
        with tempfile.TemporaryDirectory() as directory:
            self.assertEqual(self.apply(Path(directory) / "journal.json")["status"], "pass")

    def test_named_conflicting_ruleset_never_replaced(self) -> None:
        self.api.rulesets = [{"id": 1, "name": "Template default branch", "bypass_actors": [], "target": "branch", "enforcement": "disabled"}]
        self.assertTrue(self.plan()["blocked"])
        self.assertFalse(any(action["method"] == "DELETE" for action in self.plan()["actions"]))

    def test_unknown_bypass_evidence_refuses_plan(self) -> None:
        self.api.rulesets = [{"id": 1, "name": "Private"}]
        with self.assertRaises(ValueError):
            self.plan()

    def test_ignored_and_partial_writes_detected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            self.api.ignore = True
            result = self.apply(Path(directory) / "ignored.json")
            self.assertEqual(result["status"], "fail")
            self.assertTrue(result["verification"]["remaining_actions"])
            self.api = FakeGitHub()
            self.api.fail_at = 3
            receipt = Path(directory) / "partial.json"
            result = self.apply(receipt)
            self.assertEqual(result["status"], "fail")
            self.assertEqual(len(self.api.writes), 3)
            self.assertEqual(json.loads(receipt.read_text())["attempts"][-1]["outcome"], "REQUEST_PENDING_OR_UNCERTAIN")

    def test_journal_retries_transient_permission_error(self) -> None:
        original = provision.os.replace
        calls = 0

        def transient(source: str, target: str) -> None:
            nonlocal calls
            calls += 1
            if calls < 3:
                raise PermissionError(13, "Access is denied")
            original(source, target)

        with tempfile.TemporaryDirectory() as directory, \
                patch.object(provision.os, "replace", side_effect=transient), \
                patch.object(provision.time, "sleep") as sleep:
            path = Path(directory) / "journal.json"
            provision.journal(path, {"status": "pass"})
            self.assertEqual(json.loads(path.read_text()), {"status": "pass"})
            self.assertEqual(calls, 3)
            self.assertEqual(sleep.call_count, 2)

    def test_persistent_journal_permission_error_prevents_first_write(self) -> None:
        plan = self.plan()
        with tempfile.TemporaryDirectory() as directory, \
                patch.object(provision.os, "replace", side_effect=PermissionError(13, "Access is denied")), \
                patch.object(provision.time, "sleep") as sleep:
            with self.assertRaises(PermissionError):
                self.apply(Path(directory) / "journal.json", plan)
        self.assertFalse(self.api.writes)
        self.assertEqual(sleep.call_count, 5)

    def test_non_permission_journal_error_is_not_retried(self) -> None:
        with tempfile.TemporaryDirectory() as directory, \
                patch.object(provision.os, "replace", side_effect=OSError("Disk full")), \
                patch.object(provision.time, "sleep") as sleep:
            with self.assertRaises(OSError):
                provision.journal(Path(directory) / "journal.json", {"status": "fail"})
        sleep.assert_not_called()

    def test_journal_failure_prevents_first_write(self) -> None:
        with patch.object(provision, "journal", side_effect=OSError("Disk full")), self.assertRaises(OSError):
            self.apply(Path("unused.json"))
        self.assertFalse(self.api.writes)

    def test_transport_cannot_write_in_plan_or_delete(self) -> None:
        api = provision.GitHub("example/generated", "fixture-token")
        with self.assertRaises(ValueError):
            api.request("PATCH", "", {})
        api.apply = True
        for method, path in (("DELETE", "/rulesets/1"), ("POST", "/issues"), ("PATCH", "//other/repo")):
            with self.assertRaises(ValueError):
                api.request(method, path, {})

    def test_apply_requires_credentials_digest_and_journal(self) -> None:
        with patch.dict("os.environ", {}, clear=True), contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            provision.main(["--repository", "example/generated", "--profile", "library", "--contract-version", "1.2.0", "--source-sha", SHA, "--apply"])


if __name__ == "__main__":
    unittest.main()
