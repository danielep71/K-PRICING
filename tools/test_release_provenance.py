"""Synthetic release integration and real ephemeral SSH signature tests."""

from __future__ import annotations

import copy
import hashlib
import json
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import check_release as gate
import release_provenance as provenance

ROOT = Path(__file__).resolve().parents[1]


class ProvenanceTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.area = Path(self.temporary.name)
        self.root = self.area / "candidate"
        self.policy = json.loads((ROOT / gate.POLICY_PATH).read_text())
        self.trust = json.loads((ROOT / provenance.POLICY).read_text())
        self.evidence_path = self.area / "evidence.json"
        self.record_path = self.area / "provenance.json"
        self.manifest = self.area / "assets.sha256"
        self.signature = None
        self.tag_registry_keys: list[str] = []

    def new_key(self, name: str) -> Path:
        key = self.area / name
        subprocess.run(
            ["ssh-keygen", "-q", "-t", "ed25519", "-N", "", "-f", str(key)],
            check=True,
            capture_output=True,
        )
        return key

    def fixture(self, profile="application", binary=True, signed=False, tag_mode="none"):
        if signed:
            self.policy["provenance_signature_mode"] = "ssh"
        self.trust["tag_signature"] = {
            "generated": {"mode": "none"},
            "template": {"mode": "none"},
        }
        if tag_mode in {"unsigned", "signed"}:
            self.trust["tag_signature"]["template"] = {
                "mode": "ssh-github",
                "principal": "release-template",
                "github_user": "release-template",
            }
        gate._fixture_repository(self.root, profile, self.policy)
        self.configuration = gate._fixture_configuration(profile)
        self.configuration["template_contract"] = {"version": "1.2.0", "source": "example/template"}
        (self.root / gate.PROFILE_PATH).write_text(json.dumps(self.configuration))
        workflow = self.root / self.trust["workflow"]["path"]
        workflow.parent.mkdir(parents=True)
        workflow.write_text("name: Synthetic fixture\n")
        if signed:
            self.key = self.new_key("key")
            self.trust["signature"] = {
                "mode": "ssh",
                "principal": "release@example.invalid",
                "allowed_signers": ".github/release-signers",
            }
            (self.root / ".github/release-signers").write_text(
                "release@example.invalid " + self.key.with_suffix(".pub").read_text()
            )
        (self.root / provenance.POLICY).write_text(json.dumps(self.trust))
        gate._git(self.root, "add", "--all")
        self.assertEqual(gate._git(self.root, "commit", "-m", "Adopt provenance").returncode, 0)
        self.sha = gate._git_output(self.root, "rev-parse", "HEAD")
        self.assertIsNotNone(self.sha)
        if tag_mode == "signed":
            self.tag_key = self.new_key("tag-key")
            self.tag_registry_keys = [self.tag_key.with_suffix(".pub").read_text().strip()]
            gate._git(self.root, "tag", "-d", "v1.0.0")
            tagged = gate._git(
                self.root,
                "-c", "gpg.format=ssh",
                "-c", f"user.signingkey={self.tag_key}",
                "tag", "-s", "v1.0.0", "-m", "Synthetic signed candidate",
            )
            self.assertEqual(tagged.returncode, 0, tagged.stderr)
        else:
            gate._git(self.root, "tag", "-fa", "v1.0.0", "-m", "Synthetic candidate")
        assets = []
        if binary:
            asset = self.root / "dist/fixture.xlsm"
            asset.parent.mkdir()
            asset.write_bytes(b"Synthetic bytes; not an Excel workbook")
            digest = hashlib.sha256(asset.read_bytes()).hexdigest()
            assets = [{
                "path": "dist/fixture.xlsm",
                "sha256": digest,
                "candidate_sha": self.sha,
                "package_test": "PASS",
            }]
            self.manifest.write_text(f"{digest}  dist/fixture.xlsm\n")
        else:
            self.manifest = None
        self.evidence = gate._fixture_evidence(
            profile,
            self.sha,
            self.policy,
            distribution="binary" if binary else "source-only",
            assets=assets,
        )
        self.evidence_path.write_text(json.dumps(self.evidence))
        self.record = {
            "schema_version": 1,
            "repository": self.configuration["repository"],
            "candidate_sha": self.sha,
            "template_contract": self.configuration["template_contract"],
            "profile": profile,
            "tag": "v1.0.0",
            "distribution": self.evidence["distribution"],
            "digest_algorithm": "sha256",
            "evidence_sha256": hashlib.sha256(self.evidence_path.read_bytes()).hexdigest(),
            "manifest_sha256": hashlib.sha256(self.manifest.read_bytes()).hexdigest() if binary else None,
            "assets": [{"path": asset["path"], "sha256": asset["sha256"]} for asset in assets],
            "workflow": {
                "repository": self.configuration["repository"],
                "path": self.trust["workflow"]["path"],
                "sha": self.sha,
                "run_id": 123,
                "run_attempt": 1,
            },
            "environment": {
                "os": "Windows 10",
                "architecture": "x64",
                "host": "Excel",
                "host_version": "16.0",
                "office_bitness": "64",
                "runtime": "VBA7",
                "builder": "manual controlled workstation",
                "procedure": "synthetic fixture",
            },
        }
        self.save()
        if signed:
            self.sign()

    def save(self):
        self.record_path.write_text(json.dumps(self.record))

    def sign(self, namespace=provenance.NAMESPACE):
        self.signature = self.record_path.with_suffix(".json.sig")
        self.signature.unlink(missing_ok=True)
        subprocess.run(
            ["ssh-keygen", "-Y", "sign", "-f", str(self.key), "-n", namespace, str(self.record_path)],
            check=True,
            capture_output=True,
        )

    def report(self, include=True):
        with patch.object(provenance, "github_signing_keys", return_value=self.tag_registry_keys):
            return gate.build_report(
                self.root,
                "v1.0.0",
                self.sha,
                self.evidence_path,
                self.manifest,
                True,
                self.record_path if include else None,
                self.signature,
            )

    def passes(self, include=True):
        report = self.report(include)
        self.assertEqual(report["status"], "pass", report["findings"])

    def fails(self, include=True):
        report = self.report(include)
        self.assertIn("release-provenance", [item["code"] for item in report["findings"]], report)

    def test_application_binary_repeatability(self):
        self.fixture()
        self.passes()
        self.assertEqual(self.report(), self.report())

    def test_ui_binary(self):
        self.fixture("ui-component")
        self.passes()

    def test_generated_source_scan_allows_retained_initializer_grammar(self):
        gate._fixture_repository(self.root, "library", self.policy)
        tools = self.root / "tools"
        tools.mkdir()
        initializer = tools / "initialize_repository.py"
        marker_prefix = "<!-- " + "template" + ":"
        initializer.write_text(
            'MARKER_PATTERN = r"' + marker_prefix + '(remove):(start|end) -->"\n'
            'reserved = "' + marker_prefix + '"\n',
            encoding="utf-8",
        )
        gate._git(self.root, "add", "--all")
        self.assertEqual(
            gate._git(self.root, "commit", "-m", "Retain initializer grammar").returncode,
            0,
        )

        configuration = gate._fixture_configuration("library")
        findings = gate._validate_generated_source(self.root, configuration, self.policy)
        self.assertNotIn(
            "unresolved-template-token",
            [item["code"] for item in findings],
            findings,
        )

        unresolved = "# " + "{" * 2 + "PROJECT_NAME" + "}" * 2 + "\n"
        (self.root / "README.md").write_text(unresolved, encoding="utf-8")
        findings = gate._validate_generated_source(self.root, configuration, self.policy)
        self.assertIn(
            "unresolved-template-token",
            [item["code"] for item in findings],
            findings,
        )

    def test_source_only_profiles_without_provenance(self):
        for profile in gate.SUPPORTED_PROFILES:
            with self.subTest(profile=profile), tempfile.TemporaryDirectory() as directory:
                self.root = Path(directory) / "candidate"
                self.fixture(profile, binary=False)
                self.passes(False)

    def test_binary_requires_provenance(self):
        self.fixture()
        self.fails(False)

    def test_missing_extra_modified_payloads(self):
        self.fixture()
        asset = self.root / "dist/fixture.xlsm"
        original = asset.read_bytes()
        asset.unlink()
        self.fails()
        asset.write_bytes(original + b"modified")
        self.fails()
        asset.write_bytes(original)
        (self.root / "dist/undeclared.txt").write_text("extra")
        self.fails()

    def test_source_only_rejects_hidden_payload(self):
        self.fixture("library", binary=False)
        (self.root / "dist").mkdir()
        (self.root / "dist/hidden.zip").write_bytes(b"extra")
        self.fails(False)

    def test_symlinks_rejected(self):
        self.fixture()
        asset = self.root / "dist/fixture.xlsm"
        target = self.area / "payload"
        target.write_bytes(asset.read_bytes())
        asset.unlink()
        try:
            asset.symlink_to(target)
        except OSError as error:
            if getattr(error, "winerror", None) == 1314:
                self.skipTest("requires Windows symlink creation privilege")
            raise
        self.fails()

    def test_record_bindings_and_environment(self):
        self.fixture()
        original = copy.deepcopy(self.record)
        for field, value in [
            ("candidate_sha", "a" * 40),
            ("repository", "other/repo"),
            ("template_contract", {}),
            ("profile", "library"),
            ("tag", "v9.0.0"),
            ("distribution", "source-only"),
            ("digest_algorithm", "sha1"),
            ("assets", []),
            ("evidence_sha256", "b" * 64),
            ("manifest_sha256", "c" * 64),
            ("schema_version", True),
            ("environment", {}),
            ("workflow", {}),
        ]:
            with self.subTest(field=field):
                self.record = copy.deepcopy(original)
                self.record[field] = value
                self.save()
                self.fails()

    def test_workflow_policy_binding(self):
        self.fixture()
        original = copy.deepcopy(self.record)
        for field, value in [
            ("repository", "other/repo"),
            ("path", ".github/workflows/other.yml"),
            ("sha", "b" * 40),
            ("run_id", 0),
            ("run_attempt", True),
        ]:
            with self.subTest(field=field):
                self.record = copy.deepcopy(original)
                self.record["workflow"][field] = value
                self.save()
                self.fails()

    def test_exact_evidence_and_manifest_bytes(self):
        self.fixture()
        self.evidence_path.write_bytes(self.evidence_path.read_bytes() + b"\n")
        self.fails()
        self.evidence_path.write_text(json.dumps(self.evidence))
        self.manifest.write_bytes(self.manifest.read_bytes() + b"\n")
        self.fails()

    def test_duplicate_json_key(self):
        self.fixture()
        self.record_path.write_text(
            self.record_path.read_text().replace(
                '"schema_version": 1',
                '"schema_version": 1, "schema_version": 1',
            )
        )
        self.fails()

    def test_release_policy_requires_signature_mode(self):
        policy_root = self.area / "policy"
        (policy_root / ".github").mkdir(parents=True)
        missing = copy.deepcopy(self.policy)
        missing.pop("provenance_signature_mode")
        (policy_root / gate.POLICY_PATH).write_text(json.dumps(missing))
        loaded, findings = gate._load_policy(policy_root)
        self.assertIsNone(loaded)
        self.assertIn("invalid-release-policy", [item["code"] for item in findings])

    def test_release_policy_rejects_unsupported_signature_mode(self):
        policy_root = self.area / "policy"
        (policy_root / ".github").mkdir(parents=True)
        invalid = copy.deepcopy(self.policy)
        invalid["provenance_signature_mode"] = "gpg"
        (policy_root / gate.POLICY_PATH).write_text(json.dumps(invalid))
        loaded, findings = gate._load_policy(policy_root)
        self.assertIsNone(loaded)
        self.assertIn("invalid-release-policy", [item["code"] for item in findings])

    def test_committed_release_policy_selector_is_required_by_provenance(self):
        self.fixture("library", binary=False)
        release_policy = json.loads((self.root / gate.POLICY_PATH).read_text())
        release_policy.pop("provenance_signature_mode")
        (self.root / gate.POLICY_PATH).write_text(json.dumps(release_policy))
        gate._git(self.root, "add", gate.POLICY_PATH)
        gate._git(self.root, "commit", "-m", "Remove provenance selector")
        sha = gate._git_output(self.root, "rev-parse", "HEAD")
        findings = provenance.validate(
            self.root,
            self.configuration,
            sha,
            self.evidence_path,
            None,
            None,
            None,
        )
        self.assertTrue(findings)
        self.assertIn("requires provenance_signature_mode", findings[0]["message"])

    def test_provenance_mode_must_match_unsigned_release_policy(self):
        self.fixture("library", binary=False)
        trust = copy.deepcopy(self.trust)
        trust["signature"] = {
            "mode": "ssh",
            "principal": "release@example.invalid",
            "allowed_signers": ".github/release-signers",
        }
        (self.root / provenance.POLICY).write_text(json.dumps(trust))
        gate._git(self.root, "add", provenance.POLICY)
        gate._git(self.root, "commit", "-m", "Mismatch provenance mode")
        sha = gate._git_output(self.root, "rev-parse", "HEAD")
        findings = provenance.validate(
            self.root,
            self.configuration,
            sha,
            self.evidence_path,
            None,
            None,
            None,
        )
        self.assertTrue(findings)
        self.assertIn("differs from release policy", findings[0]["message"])

    def test_provenance_mode_must_match_signed_release_policy(self):
        self.fixture("library", binary=False)
        release_policy = json.loads((self.root / gate.POLICY_PATH).read_text())
        release_policy["provenance_signature_mode"] = "ssh"
        (self.root / gate.POLICY_PATH).write_text(json.dumps(release_policy))
        gate._git(self.root, "add", gate.POLICY_PATH)
        gate._git(self.root, "commit", "-m", "Require signed provenance")
        sha = gate._git_output(self.root, "rev-parse", "HEAD")
        findings = provenance.validate(
            self.root,
            self.configuration,
            sha,
            self.evidence_path,
            None,
            None,
            None,
        )
        self.assertTrue(findings)
        self.assertIn("differs from release policy", findings[0]["message"])

    @unittest.skipUnless(shutil.which("ssh-keygen"), "requires OpenSSH")
    def test_real_ssh_signature_and_tamper(self):
        self.fixture(signed=True)
        self.passes()
        self.record["environment"]["builder"] = "changed assertion"
        self.save()
        self.fails()

    @unittest.skipUnless(shutil.which("ssh-keygen"), "requires OpenSSH")
    def test_wrong_namespace(self):
        self.fixture(signed=True)
        self.sign("wrong-namespace")
        self.fails()

    @unittest.skipUnless(shutil.which("ssh-keygen"), "requires OpenSSH")
    def test_untrusted_key(self):
        self.fixture(signed=True)
        self.key = self.new_key("other-key")
        self.sign()
        self.fails()

    @unittest.skipUnless(shutil.which("ssh-keygen"), "requires OpenSSH")
    def test_missing_signature_and_cannot_disable_in_worktree(self):
        self.fixture(signed=True)
        self.signature = None
        (self.root / provenance.POLICY).write_text('{"signature":{"mode":"none"}}')
        self.configuration["template_contract"]["version"] = "1.1.0"
        (self.root / gate.PROFILE_PATH).write_text(json.dumps(self.configuration))
        self.fails()

    @unittest.skipUnless(shutil.which("ssh-keygen"), "requires OpenSSH")
    def test_signature_mandatory_even_source_only(self):
        self.fixture("library", binary=False, signed=True)
        self.passes()
        self.fails(False)

    def test_signature_without_enabled_policy_rejected(self):
        self.fixture()
        self.signature = self.area / "not-a-signature"
        self.fails()

    def test_invalid_committed_policy_cannot_be_repaired_by_external_record(self):
        self.fixture()
        for value in ({"mode": "unknown"}, {"mode": "ssh"}, {"mode": "none", "extra": True}):
            with self.subTest(value=value):
                self.trust["signature"] = value
                (self.root / provenance.POLICY).write_text(json.dumps(self.trust))
                gate._git(self.root, "add", provenance.POLICY)
                gate._git(self.root, "commit", "-m", "Synthetic invalid policy")
                sha = gate._git_output(self.root, "rev-parse", "HEAD")
                findings = provenance.validate(
                    self.root,
                    self.configuration,
                    sha,
                    self.evidence_path,
                    self.manifest,
                    None,
                    None,
                )
                self.assertTrue(findings)

    def test_missing_committed_policy(self):
        self.fixture("library", binary=False)
        gate._git(self.root, "rm", provenance.POLICY)
        gate._git(self.root, "commit", "-m", "Synthetic missing policy")
        sha = gate._git_output(self.root, "rev-parse", "HEAD")
        self.assertTrue(
            provenance.validate(
                self.root,
                self.configuration,
                sha,
                self.evidence_path,
                None,
                None,
                None,
            )
        )

    def test_unknown_contract_rejected(self):
        self.fixture("library", binary=False)
        self.configuration["template_contract"]["version"] = "99.0.0"
        (self.root / gate.PROFILE_PATH).write_text(json.dumps(self.configuration))
        gate._git(self.root, "add", gate.PROFILE_PATH)
        gate._git(self.root, "commit", "-m", "Synthetic future contract")
        sha = gate._git_output(self.root, "rev-parse", "HEAD")
        self.assertTrue(
            provenance.validate(
                self.root,
                self.configuration,
                sha,
                self.evidence_path,
                None,
                None,
                None,
            )
        )

    @unittest.skipUnless(shutil.which("ssh-keygen"), "requires OpenSSH")
    def test_verifier_unavailable_or_times_out(self):
        self.fixture(signed=True)
        real_run = subprocess.run
        for error in (FileNotFoundError("ssh-keygen"), subprocess.TimeoutExpired("ssh-keygen", 30)):
            def run(args, **kwargs):
                if args[0] == "ssh-keygen":
                    raise error
                return real_run(args, **kwargs)

            with self.subTest(error=type(error).__name__), patch.object(provenance.subprocess, "run", run):
                self.fails()

    @unittest.skipUnless(shutil.which("ssh-keygen"), "requires OpenSSH")
    def test_signed_template_tag_accepts_current_github_signing_key(self):
        self.fixture("template", binary=False, tag_mode="signed")
        self.passes(False)

    @unittest.skipUnless(shutil.which("ssh-keygen"), "requires OpenSSH")
    def test_unsigned_template_tag_is_rejected(self):
        self.fixture("template", binary=False, tag_mode="unsigned")
        self.fails(False)

    @unittest.skipUnless(shutil.which("ssh-keygen"), "requires OpenSSH")
    def test_wrong_github_signing_key_is_rejected(self):
        self.fixture("template", binary=False, tag_mode="signed")
        other = self.new_key("wrong-tag-key")
        self.tag_registry_keys = [other.with_suffix(".pub").read_text().strip()]
        self.fails(False)

    @unittest.skipUnless(shutil.which("ssh-keygen"), "requires OpenSSH")
    def test_corrupted_tag_signature_is_rejected(self):
        self.fixture("template", binary=False, tag_mode="signed")
        result = subprocess.run(
            ["git", "-C", str(self.root), "cat-file", "tag", "refs/tags/v1.0.0"],
            check=True,
            capture_output=True,
        )
        raw = bytearray(result.stdout)
        marker = b"-----BEGIN SSH SIGNATURE-----\n"
        start = raw.index(marker) + len(marker)
        while raw[start] in b"\r\n":
            start += 1
        raw[start] = ord("A") if raw[start] != ord("A") else ord("B")
        written = subprocess.run(
            ["git", "-C", str(self.root), "hash-object", "-t", "tag", "-w", "--stdin"],
            input=bytes(raw),
            check=True,
            capture_output=True,
            text=False,
        )
        object_sha = written.stdout.decode().strip()
        subprocess.run(
            ["git", "-C", str(self.root), "update-ref", "refs/tags/v1.0.0", object_sha],
            check=True,
            capture_output=True,
        )
        self.fails(False)

    @unittest.skipUnless(shutil.which("ssh-keygen"), "requires OpenSSH")
    def test_signed_moved_tag_is_rejected(self):
        self.fixture("template", binary=False, tag_mode="signed")
        self.assertEqual(
            gate._git(self.root, "commit", "--allow-empty", "-m", "Move tag target").returncode,
            0,
        )
        gate._git(self.root, "tag", "-d", "v1.0.0")
        tagged = gate._git(
            self.root,
            "-c", "gpg.format=ssh",
            "-c", f"user.signingkey={self.tag_key}",
            "tag", "-s", "v1.0.0", "-m", "Moved signed tag",
        )
        self.assertEqual(tagged.returncode, 0, tagged.stderr)
        self.fails(False)

    @unittest.skipUnless(shutil.which("ssh-keygen"), "requires OpenSSH")
    def test_tag_signer_rotation_and_revocation(self):
        self.fixture("template", binary=False, tag_mode="signed")
        original = self.tag_registry_keys[0]
        replacement = self.new_key("replacement-tag-key").with_suffix(".pub").read_text().strip()
        self.tag_registry_keys = [original, replacement]
        self.passes(False)
        self.tag_registry_keys = [replacement]
        self.fails(False)

    def test_generated_release_does_not_inherit_template_tag_requirement(self):
        self.fixture("library", binary=False, tag_mode="unsigned")
        self.passes(False)


if __name__ == "__main__":
    unittest.main()
