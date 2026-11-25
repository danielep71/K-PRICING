"""Tests for tools/check_repository.py, in temporary Git repositories.

The run itself is tested with small checks injected in place of the real
list, so that the test suite never runs itself.
"""

from __future__ import annotations

import contextlib
import importlib.util
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import check_repository  # noqa: E402
from check_repository import Check, Context  # noqa: E402
from gitfixture import git, init  # noqa: E402

# The checks that read YAML need PyYAML, which CI always installs.
NEEDS_YAML = unittest.skipIf(
    importlib.util.find_spec("yaml") is None and not os.environ.get("GITHUB_ACTIONS"),
    "PyYAML is not installed: python -m pip install -r .github/requirements.txt")


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class RepositoryTest(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="check-repository-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        init(self.root)

    def write(self, files: dict[str, str | bytes]) -> None:
        for name, content in files.items():
            path = self.root / name
            path.parent.mkdir(parents=True, exist_ok=True)
            if isinstance(content, bytes):
                path.write_bytes(content)
            else:
                path.write_text(content, encoding="utf-8")

    def commit(self, files: dict[str, str | bytes] | None = None) -> None:
        self.write(files or {})
        git(self.root, "add", "-A")
        git(self.root, "commit", "-qm", "Fixture", "--allow-empty")

    def context(self, **values: object) -> Context:
        base: dict[str, object] = {"repository": "owner/project", "default_branch": "main",
                                   "is_template": False}
        base.update(values)
        return Context(self.root, **base)  # type: ignore[arg-type]


class FileCheckTest(RepositoryTest):
    def test_template_only_content(self) -> None:
        self.write({"AGENTS.md": "<!-- template-only:start -->\n"})
        self.assertEqual(len(check_repository.template_only(self.context())), 1)
        self.assertEqual(check_repository.template_only(self.context(is_template=True)), [])
        self.write({"AGENTS.md": "# Instructions\n"})
        self.assertEqual(check_repository.template_only(self.context()), [])
        # The README's template-only sections count as much as those of AGENTS.md.
        self.write({"README.md": "# Project\n\n<!-- template-only:start -->\nx\n"
                                 "<!-- template-only:end -->\n"})
        self.assertEqual([path for path, _ in check_repository.template_only(self.context())],
                         ["README.md"])

    def test_contact_links(self) -> None:
        self.write({".github/ISSUE_TEMPLATE/config.yml":
                    'contact_links:\n  - name: Security\n'
                    '    url: "https://github.com/Owner/Project/security"\n'
                    "  - name: Other\n    url: https://github.com/other/template/discussions\n"})
        findings = check_repository.contact_links(self.context())
        self.assertEqual(len(findings), 1)
        self.assertIn("https://github.com/other/template/discussions", findings[0][1])

    def test_traffic_badges(self) -> None:
        badge = "https://raw.githubusercontent.com/{}/traffic-history/badges/views.svg"
        self.write({"README.md": f"![a]({badge.format('owner/project')})\n"})
        self.assertEqual(check_repository.traffic_badges(self.context()), [])
        self.write({"README.md": f"![a]({badge.format('other/template')})\n"})
        self.assertIn("reads other/template",
                      check_repository.traffic_badges(self.context())[0][1])

    def test_the_badge_check_is_listed_while_traffic_or_a_badge_is_there(self) -> None:
        def listed() -> bool:
            return "readme-traffic-badges" in [check.name for check in
                                               check_repository.checks(self.context())]
        self.commit({"README.md": "# Project\n"})
        self.assertFalse(listed())
        self.write({check_repository.TRAFFIC_WORKFLOW: "name: Traffic history\n"})
        self.assertTrue(listed())
        (self.root / check_repository.TRAFFIC_WORKFLOW).unlink()
        badge = "https://raw.githubusercontent.com/other/template/traffic-history/badges/x.json"
        self.write({"README.md": f"![a]({badge})\n"})
        self.assertTrue(listed())

    def origin(self, *branches: str) -> Path:
        """A bare repository with main and these branches, as origin would be."""
        self.commit({"README.md": "x\n"})
        bare = self.root.parent / f"{self.root.name}-origin.git"
        self.addCleanup(shutil.rmtree, bare, ignore_errors=True)
        git(self.root, "init", "-q", "--bare", str(bare))
        for branch in ("main", *branches):
            git(self.root, "push", "-q", str(bare), f"HEAD:refs/heads/{branch}")
        return bare

    def single_branch_clone(self, *branches: str) -> Path:
        """A clone of main only: no remote-tracking branch for the other branches."""
        clone = self.root.parent / f"{self.root.name}-clone"
        self.addCleanup(shutil.rmtree, clone, ignore_errors=True)
        subprocess.run(["git", "clone", "-q", "--single-branch", "--branch", "main",
                        str(self.origin(*branches)), str(clone)], capture_output=True,
                       check=True)
        for branch in branches:
            self.assertEqual(subprocess.run(["git", "rev-parse", "--verify", "--quiet",
                                             f"refs/remotes/origin/{branch}"], cwd=clone,
                                            capture_output=True, check=False).returncode, 1)
        (clone / ".github").mkdir()
        (clone / ".github/dependabot.yml").write_text(
            "updates:\n  - package-ecosystem: pip\n    target-branch: release/1.0.0\n",
            encoding="utf-8")
        return clone

    def test_dependabot_target_branches(self) -> None:
        git(self.root, "remote", "add", "origin", str(self.origin()))
        git(self.root, "update-ref", "refs/remotes/origin/release/1.0.0", "HEAD")
        self.write({".github/dependabot.yml":
                    "updates:\n  - package-ecosystem: pip\n    target-branch: release/1.0.0\n"
                    '  - package-ecosystem: npm\n    target-branch: "release/0.9.0"\n'})
        findings = check_repository.dependabot_branches(self.context())
        self.assertEqual([message for _, message in findings],
                         ["target-branch release/0.9.0 does not exist; set it to the active "
                          "release branch or remove it."])

    def test_dependabot_target_branch_only_on_origin_passes(self) -> None:
        clone = self.single_branch_clone("release/1.0.0")
        self.assertEqual(check_repository.dependabot_branches(Context(clone)), [])

    def test_dependabot_target_branch_nowhere_fails(self) -> None:
        # A longer name that ends with the branch's name is not the branch. git ls-remote
        # returns the second one for the pattern refs/heads/release/1.0.0, so only the
        # exact-name filter rejects it.
        clone = self.single_branch_clone("old/release/1.0.0", "x/refs/heads/release/1.0.0")
        self.assertEqual(check_repository.dependabot_branches(Context(clone)),
                         [(".github/dependabot.yml", "target-branch release/1.0.0 does not "
                           "exist; set it to the active release branch or remove it.")])

    @NEEDS_YAML
    def test_codeql_default_branch(self) -> None:
        self.write({".github/workflows/codeql.yml":
                    "on:\n  push:\n    branches: [main]\n  pull_request:\n"
                    "    branches: [main, 'release/**']\n"})
        self.assertEqual(check_repository.codeql_default_branch(self.context()), [])
        findings = check_repository.codeql_default_branch(self.context(default_branch="trunk"))
        self.assertEqual(len(findings), 2)
        self.assertIn("excludes the default branch trunk", findings[0][1])

    @NEEDS_YAML
    def test_integrity_push_branches_name_the_default_branch(self) -> None:
        integrity = ".github/workflows/repository-integrity.yml"
        self.write({integrity:
                    "on:\n  pull_request:\n  push:\n    branches: [main, 'release/**']\n"})
        self.assertEqual(check_repository.codeql_default_branch(self.context()), [])
        self.assertEqual(check_repository.codeql_default_branch(self.context(default_branch="trunk")),
                         [(integrity, "push filter {'branches': ['main', 'release/**']} "
                                      "excludes the default branch trunk; make it match.")])

    @NEEDS_YAML
    def test_minimum_python_push_branches_name_the_default_branch(self) -> None:
        minimum = ".github/workflows/minimum-python.yml"
        self.write({minimum: "on:\n  pull_request:\n    paths: ['**/*.py']\n"
                             "  push:\n    branches: [main]\n"})
        self.assertEqual(check_repository.codeql_default_branch(self.context()), [])
        self.assertEqual(check_repository.codeql_default_branch(self.context(default_branch="trunk")),
                         [(minimum, "push filter {'branches': ['main']} "
                                    "excludes the default branch trunk; make it match.")])

    def test_branch_patterns_follow_github_rules(self) -> None:
        cases = [("main", "main", True), ("main", "mainline", False),
                 ("release/*", "release/1.1.0", True), ("release/*", "release/a/b", False),
                 ("release/**", "release/a/b", True), ("**", "feature/x", True),
                 ("v[0-9]+", "v12", True), ("v[0-9]+", "v", False), ("ma?in", "min", True),
                 ("ma?in", "maain", False), ("[a-c]x", "bx", True), (r"a\*b", "a*b", True),
                 (r"a\*b", "axb", False), ("feat.x", "featxx", False),
                 ("**/main", "main", True), ("**/main", "a/b/main", True),
                 ("**/main", "xmain", False), ("a/**/b", "a/b", True), ("a/**/b", "a/x/y/b", True)]
        for pattern, branch, expected in cases:
            with self.subTest(pattern=pattern, branch=branch):
                self.assertEqual(bool(check_repository.branch_pattern(pattern).fullmatch(branch)),
                                 expected)
        for malformed in ("a\\", "?a", "**+", "[", "[]", "[z-a]", "[a_b]"):
            with self.subTest(malformed=malformed), self.assertRaises(ValueError):
                check_repository.branch_pattern(malformed)

    @NEEDS_YAML
    def test_required_workflows_read_filters_as_github_does(self) -> None:
        codeql = ".github/workflows/codeql.yml"
        cases = {"push:\n    branches: ['**']\n": True,
                 "push:\n    branches: ['**', '!main']\n": False,
                 "push:\n    branches: ['!main', '**']\n": True,
                 "push:\n    branches-ignore: [main]\n": False,
                 "push:\n    branches-ignore: ['release/**']\n": True,
                 "push:\n    branches: ['**', '!**/main']\n": False,
                 "push:\n    branches-ignore: ['**/main']\n": False,
                 "push:\n    tags: ['v*']\n": False,
                 "push:\n    tags: ['v*']\n    branches: [main]\n": True,
                 "pull_request_target:\n    branches: ['release/**']\n": False,
                 "pull_request:\n": True}
        for trigger, runs in cases.items():
            with self.subTest(trigger=trigger):
                self.write({codeql: "on:\n  " + trigger})
                findings = check_repository.codeql_default_branch(self.context())
                self.assertEqual(findings == [], runs, findings)
                if not runs:
                    self.assertIn("excludes the default branch main", findings[0][1])

    @NEEDS_YAML
    def test_every_workflow_branch_filter_must_be_well_formed(self) -> None:
        other = ".github/workflows/other.yml"
        malformed = {"push:\n    branches: [main]\n    branches-ignore: [x]\n": "sets both",
                     "push:\n    branches: ['!main']\n": "lists only negations",
                     "pull_request:\n    branches-ignore: ['!main']\n": "cannot negate",
                     "pull_request_target:\n    branches: ['v[0-9']\n": "malformed or unclosed",
                     "push:\n    branches: []\n": "non-empty list",
                     "push:\n    branches: [1]\n": "non-empty list"}
        for trigger, reason in malformed.items():
            with self.subTest(trigger=trigger):
                self.write({other: "on:\n  " + trigger})
                findings = check_repository.codeql_default_branch(self.context())
                self.assertEqual(len(findings), 1, findings)
                self.assertEqual(findings[0][0], other)
                self.assertIn(reason, findings[0][1])
        # A workflow made for release branches only is valid: only CodeQL,
        # Minimum Python and Repository integrity must run on the default branch.
        self.write({other: "on:\n  push:\n    branches: ['release/**']\n"})
        self.assertEqual(check_repository.codeql_default_branch(self.context()), [])

    def test_secret_files(self) -> None:
        failing = {".env": ".env", "config/.env.local": ".env.*", ".ENV.production": ".env.*",
                   "certs/server.pem": "*.pem", "tls.key": "*.key", "client.p12": "*.p12",
                   "sign.PFX": "*.pfx", "code.pvk": "*.pvk", "store.jks": "*.jks",
                   "release.keystore": "*.keystore", "keys/id_rsa": "id_rsa*",
                   "id_rsa_deploy": "id_rsa*", "home/.ssh/id_ed25519": "id_ed25519*",
                   "vault.kdbx": "*.kdbx", "tests/fixtures/CASE-01/base/.env": ".env"}
        passing = (".env.example", "config/.env.example", "keys/id_rsa.pub",
                   "id_ed25519.pub", "id_rsa_deploy.pub", ".envrc", "env.md", "keys.md",
                   "docs/pem.txt", "tests/fixtures/CASE-01/base/input.csv")
        # git add -f, as a force-added file bypasses .gitignore.
        self.write(dict.fromkeys([*failing, *passing], "x\n"))
        git(self.root, "add", "-f", ".")
        findings = check_repository.secret_files(self.context())
        self.assertEqual(sorted(path for path, _ in findings), sorted(failing))
        for path, message in findings:
            with self.subTest(path=path):
                self.assertIn(f"({failing[path]})", message)
                self.assertIn("remove it from the repository, rotate the secret", message)
                self.assertIn("Handling a finding in docs/repo_governance/DISTRIBUTION.md",
                              message)

    def test_gitignore_ignores_the_secret_file_names(self) -> None:
        def ignored(path: str) -> bool:
            return subprocess.run(["git", "-C", str(ROOT), "check-ignore", "-q", "--no-index",
                                   path], check=False).returncode == 0

        for path in (".env", ".env.local", "a.pem", "a.key", "a.p12", "a.pfx", "a.pvk",
                     "a.jks", "a.keystore", "id_rsa", "id_rsa_deploy", "id_ed25519", "a.kdbx"):
            with self.subTest(path=path):
                self.assertIsNotNone(check_repository.secret_pattern(path))
                self.assertTrue(ignored(f"config/{path}"))
        for path in (".env.example", "id_rsa.pub", "id_ed25519.pub"):
            with self.subTest(path=path):
                self.assertIsNone(check_repository.secret_pattern(path))
                self.assertFalse(ignored(f"config/{path}"))

    def test_repository_tracks_no_secret_file(self) -> None:
        self.assertEqual(check_repository.secret_files(Context(ROOT)), [])

    def test_required_documents(self) -> None:
        self.write({"README.md": "x\n", "LICENSE": "",
                    "docs/repo_governance/SETTINGS.md": "x\n"})
        self.assertEqual(check_repository.required_documents(self.context()),
                         [("LICENSE", "is missing or empty")])


class WhitespaceTest(RepositoryTest):
    def outcomes(self) -> dict[str, int]:
        return {name: subprocess.run(command, cwd=self.root, capture_output=True,
                                     check=False).returncode
                for name, command in check_repository.whitespace_commands(self.root)}

    def test_committed_trailing_white_space_fails_in_a_clean_checkout(self) -> None:
        self.commit({"notes.txt": "trailing \n"})
        self.assertEqual(git(self.root, "status", "--porcelain"), "")
        self.assertNotEqual(self.outcomes()["whitespace-committed"], 0)
        self.assertEqual(self.outcomes()["whitespace-local"], 0)

    def test_clean_text_crlf_and_binary_pass(self) -> None:
        self.commit({".gitattributes": "*.cmd text eol=crlf\n*.bin binary\n",
                     "notes.txt": "clean\n", "run.cmd": b"echo clean\r\n",
                     "data.bin": b"\x00binary \n"})
        self.assertEqual(set(self.outcomes().values()), {0})

    def test_byte_preserved_files_pass_under_the_repository_attributes(self) -> None:
        # Source copies keep their issuer's bytes; a form keeps the VBE header,
        # whose Begin line ends with a blank (tools/check_vba_source.py checks
        # the code after it).
        header = ["VERSION 5.00", "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRun ",
                  "   Caption         =   \"Run\"   ", "\tClientHeight    =   3015", "End",
                  "Attribute VB_Name = \"frmRun\"", "Option Explicit"]
        self.commit({".gitattributes": (ROOT / ".gitattributes").read_bytes(),
                     "docs/methodology/sources/SRC-01_rules.csv": b"a,b \r\nc,d\t\r\n",
                     "docs/methodology/sources/SRC-02_rules.txt": b"no final newline  ",
                     "src/forms/frmRun.frm": "".join(f"{line}\r\n" for line in header)
                     .encode("ascii")})
        self.assertEqual(git(self.root, "status", "--porcelain"), "")
        self.assertEqual(self.outcomes(), {"whitespace-committed": 0, "whitespace-local": 0})

    def test_ordinary_text_beside_byte_preserved_files_is_checked(self) -> None:
        self.commit({".gitattributes": (ROOT / ".gitattributes").read_bytes()})
        for path in ("docs/methodology/sources/README.md", "docs/notes.md",
                     "src/core/CORE_Probe.bas"):
            with self.subTest(path=path):
                self.commit({path: "trailing  \n"})
                self.assertNotEqual(self.outcomes()["whitespace-committed"], 0)
                git(self.root, "rm", "-q", path)
                self.commit()

    def test_local_white_space_fails(self) -> None:
        self.commit({"notes.txt": "clean\n"})
        self.write({"notes.txt": "bad\t\n"})
        self.assertNotEqual(self.outcomes()["whitespace-local"], 0)


class RunTest(RepositoryTest):
    def run_checks(self, listed: list[Check], strict: bool = False,
                   summary: Path | None = None) -> tuple[int, str, dict[str, Any]]:
        with contextlib.redirect_stdout(io.StringIO()) as printed, \
                mock.patch.dict(os.environ, {"PR_HEAD_SHA": "f" * 40, "GITHUB_ACTIONS": ""}):
            code = check_repository.run(self.context(), listed,
                                        self.root / "test-results/integrity", strict, summary)
        report = json.loads((self.root / "test-results/integrity/report.json")
                            .read_text(encoding="utf-8"))
        return code, printed.getvalue(), report

    def test_report_logs_and_summary(self) -> None:
        self.commit({"README.md": "x\n"})
        self.write({"untracked.txt": "y\n"})
        listed = [Check("ok", [sys.executable, "-c", "print('fine')"]),
                  Check("broken", [sys.executable, "-c", "import sys; sys.exit('bad input')"]),
                  Check("finding", function=lambda context: [("README.md", "is wrong")])]
        step_summary = self.root / "step-summary.md"
        step_summary.write_text("before\n", encoding="utf-8")
        code, printed, report = self.run_checks(listed, summary=step_summary)
        self.assertEqual(code, 1)
        self.assertIn("broken: FAIL (exit code 1)", printed)
        self.assertIn("bad input", printed)
        self.assertEqual(report["status"], "fail")
        self.assertEqual(report["commit"], git(self.root, "rev-parse", "HEAD").strip())
        self.assertIs(report["dirty"], True)
        self.assertEqual(report["pull_request_head"], "f" * 40)
        self.assertEqual(set(report["files_sha256"]), {"README.md"})
        statuses = {check["name"]: (check["status"], check["exit_code"])
                    for check in report["checks"]}
        self.assertEqual(statuses, {"ok": ("pass", 0), "broken": ("fail", 1),
                                    "finding": ("fail", 1)})
        out = self.root / "test-results/integrity"
        self.assertIn("fine", (out / "ok.log").read_text(encoding="utf-8"))
        self.assertIn("FAIL: README.md: is wrong", (out / "finding.log").read_text("utf-8"))
        text = step_summary.read_text(encoding="utf-8")
        self.assertTrue(text.startswith("before\n## Repository integrity: FAIL"))
        self.assertIn("| `broken` | ❌ fail | exit code 1 |", text)
        self.assertIn("no Excel run", text)
        self.assertEqual((out / "summary.md").read_text(encoding="utf-8"),
                         text.removeprefix("before\n"))

    @NEEDS_YAML
    def test_a_check_that_raises_fails_and_the_run_continues(self) -> None:
        self.commit({".github/workflows/codeql.yml": "on:\n  push: {}\n  push: {}\n"})
        listed = [Check("codeql-default-branch", function=check_repository.codeql_default_branch),
                  Check("after", [sys.executable, "-c", "print('ran')"])]
        code, printed, report = self.run_checks(listed)
        self.assertEqual(code, 1)
        statuses = {check["name"]: check["status"]
                    for check in report["checks"]}
        self.assertEqual(statuses, {"codeql-default-branch": "fail", "after": "pass"})
        log = (self.root / "test-results/integrity/codeql-default-branch.log").read_text("utf-8")
        self.assertIn("duplicate key", log)
        self.assertIn("codeql-default-branch: FAIL (raised an error)", printed)

    def sleeper(self, name: str, seconds: float, writes: str | None = None) -> Check:
        """A command check that sleeps, then writes ``writes`` into the working tree,
        if given, and records when it finished, outside the tree, for
        ``finish_order``."""
        if not hasattr(self, "finished"):
            directory = tempfile.TemporaryDirectory(prefix="check-repository-finished-")
            self.addCleanup(directory.cleanup)
            self.finished = Path(directory.name)
        code = ("import pathlib, sys, time; time.sleep(float(sys.argv[1])); "
                "[pathlib.Path(name).write_text('late\\n') for name in sys.argv[3:]]; "
                "pathlib.Path(sys.argv[2]).write_text(str(time.time_ns()))")
        return Check(name, [sys.executable, "-c", code, str(seconds),
                            str(self.finished / name), *([writes] if writes else [])])

    def finish_order(self) -> list[str]:
        times = {path.name: int(path.read_text()) for path in self.finished.iterdir()}
        return sorted(times, key=times.__getitem__)

    def run_in_parallel(self, listed: list[Check]) -> tuple[int, str, dict[str, Any]]:
        """run_checks with four workers, however many CPUs this machine has."""
        with mock.patch("os.cpu_count", return_value=4):
            return self.run_checks(listed)

    def test_the_report_keeps_the_listed_order_when_checks_finish_out_of_order(self) -> None:
        self.commit({".gitignore": "/test-results/\n"})
        listed = [self.sleeper("slowest", 1.2), self.sleeper("slow", 0.8),
                  Check("function", function=lambda context: []),
                  Check("missing", ["absent-tool"], missing="absent-tool is not on PATH"),
                  self.sleeper("quick", 0.4), self.sleeper("quickest", 0)]
        code, printed, report = self.run_in_parallel(listed)
        self.assertEqual(code, 0, printed)
        # The command checks ran together, so the quickest finished first ...
        self.assertEqual(self.finish_order(), ["quickest", "quick", "slow", "slowest"])
        # ... yet the console, the report and the summary follow the list.
        names = ["slowest", "slow", "function", "missing", "quick", "quickest"]
        self.assertEqual([line.split(":")[0] for line in printed.splitlines()][:-1], names)
        self.assertEqual([check["name"] for check in report["checks"]], names)
        self.assertEqual([check["status"] for check in report["checks"]],
                         ["pass", "pass", "pass", "skip", "pass", "pass"])
        summary = (self.root / "test-results/integrity/summary.md").read_text("utf-8")
        self.assertEqual(re.findall(r"^\| `([^`]+)` \|", summary, re.M), names)

    def test_clean_working_tree_runs_after_every_other_check(self) -> None:
        self.commit({".gitignore": "/test-results/\n", "README.md": "x\n"})
        clean = check_repository.checks(self.context(ci=True))[-1]
        self.assertEqual((clean.name, clean.alone), ("clean-working-tree", True))
        listed = [self.sleeper("late-writer", 0.8, "late.txt"), self.sleeper("quick", 0), clean]
        code, printed, report = self.run_in_parallel(listed)
        self.assertEqual(self.finish_order(), ["quick", "late-writer"])
        # The file written last, by the slowest check, is still seen.
        self.assertEqual(code, 1, printed)
        self.assertEqual([(check["name"], check["status"]) for check in report["checks"]],
                         [("late-writer", "pass"), ("quick", "pass"),
                          ("clean-working-tree", "fail")])
        self.assertIn("FAIL: late.txt: differs from the commit (git status ??)", printed)

    def test_a_command_check_that_raises_fails_and_the_run_continues(self) -> None:
        self.commit({".gitignore": "/test-results/\n"})

        def broken(context: Context) -> list[check_repository.Finding]:
            raise RuntimeError("cannot read")

        listed = [self.sleeper("slow", 0.4),
                  # A null byte in an argument cannot reach the operating system.
                  Check("cannot-start", [sys.executable, "-c", "print('never')\0"]),
                  Check("raises", function=broken), self.sleeper("after", 0)]
        code, printed, report = self.run_in_parallel(listed)
        self.assertEqual(code, 1, printed)
        self.assertEqual(self.finish_order(), ["after", "slow"])
        self.assertEqual([(check["name"], check["status"], check["detail"])
                          for check in report["checks"]],
                         [("slow", "pass", ""), ("cannot-start", "fail", "raised an error"),
                          ("raises", "fail", "raised an error"), ("after", "pass", "")])
        log = (self.root / "test-results/integrity/cannot-start.log").read_text("utf-8")
        self.assertIn("ValueError", log)
        self.assertIn("RuntimeError: cannot read", printed)

    def test_every_check_that_runs_node_is_skipped_without_it(self) -> None:
        # check_issues.py --check-config validates the catalogue with the label script.
        with mock.patch("shutil.which",
                        side_effect=lambda name: None if name == "node" else name):
            listed = check_repository.checks(self.context())
        needs_node = {check.name: check.missing for check in listed
                      if check.command and (check.command[0] == "node"
                                            or "tools/check_issues.py" in check.command)}
        self.assertEqual(set(needs_node), {"labels-sync-self-test", "labels-catalogue",
                                           "labels-drift-self-test", "issue-configuration"})
        self.assertEqual(set(needs_node.values()), {"node is not on PATH"})

    def test_missing_context_is_skipped_unless_strict(self) -> None:
        self.commit()
        listed = [Check("needs-node", ["node"], missing="node is not on PATH")]
        code, printed, report = self.run_checks(listed)
        self.assertEqual(code, 0)
        self.assertIn("needs-node: SKIP (node is not on PATH)", printed)
        self.assertEqual(report["status"], "pass")
        code, _, report = self.run_checks(listed, strict=True)
        self.assertEqual(code, 1)
        self.assertEqual(report["status"], "fail")

    def test_dependabot_target_branches_without_origin_are_skipped_unless_strict(self) -> None:
        self.commit({".github/dependabot.yml":
                     "updates:\n  - package-ecosystem: pip\n    target-branch: release/1.0.0\n"})
        listed = [check for check in check_repository.checks(self.context())
                  if check.name == "dependabot-target-branches"]
        code, printed, report = self.run_checks(listed)
        self.assertEqual(code, 0, printed)
        # The rest is Git's own message, which a locale may translate.
        self.assertIn("dependabot-target-branches: SKIP (origin cannot be asked whether branch "
                      "release/1.0.0 exists (git ls-remote: ", printed)
        self.assertEqual(report["checks"][0]["status"], "skip")
        code, printed, report = self.run_checks(listed, strict=True)
        self.assertEqual(code, 1, printed)
        self.assertEqual(report["checks"][0]["status"], "fail")

    def test_codeql_default_branch_without_pyyaml_is_skipped_unless_strict(self) -> None:
        self.commit({".github/workflows/codeql.yml": "on:\n  push:\n    branches: [main]\n"})
        reason = "PyYAML is not installed (.github/requirements.txt)"
        # Block PyYAML, and the loader that imports it, as if neither were installed.
        with mock.patch.dict(sys.modules, {"yaml": None}):
            sys.modules.pop("strict_yaml", None)
            listed = [check for check in check_repository.checks(self.context())
                      if check.name == "codeql-default-branch"]
            code, printed, report = self.run_checks(listed)
            self.assertEqual(code, 0, printed)
            self.assertIn(f"codeql-default-branch: SKIP ({reason})", printed)
            self.assertEqual(report["checks"][0]["status"], "skip")
            code, printed, report = self.run_checks(listed, strict=True)
            self.assertEqual(code, 1, printed)
            self.assertEqual(report["checks"][0]["status"], "fail")
            unknown = {check.name: check.missing
                       for check in check_repository.checks(self.context(default_branch=None))}
            self.assertEqual(unknown["codeql-default-branch"],
                             f"the default branch is unknown (--default-branch); {reason}")

    def test_ci_fails_a_working_tree_that_differs_from_the_commit(self) -> None:
        self.commit({".gitignore": "/test-results/\n", "README.md": "x\n"})
        listed = [Check("clean-working-tree", function=check_repository.clean_working_tree)]
        code, printed, report = self.run_checks(listed)
        self.assertEqual((code, report["dirty"]), (0, False), printed)
        self.write({"README.md": "changed\n", "notes/new.txt": "y\n"})
        code, printed, report = self.run_checks(listed)
        self.assertEqual((code, report["dirty"], report["status"]), (1, True, "fail"))
        self.assertIn("clean-working-tree: FAIL (2 finding(s))", printed)
        self.assertIn("FAIL: README.md: differs from the commit (git status M)", printed)
        self.assertIn("FAIL: notes/new.txt: differs from the commit (git status ??)", printed)

    def test_an_advisory_check_warns_without_failing_the_run(self) -> None:
        self.commit()
        listed = [Check("advice", [sys.executable, "-c", "import sys; sys.exit('a finding')"],
                        advisory=True)]
        code, printed, report = self.run_checks(listed)
        self.assertEqual(code, 0)
        self.assertIn("advice: WARN (advisory, not a failure: exit code 1)", printed)
        self.assertIn("a finding", printed)
        self.assertEqual(report["status"], "pass")
        self.assertEqual(report["checks"][0]["status"], "warn")
        summary = (self.root / "test-results/integrity/summary.md").read_text(encoding="utf-8")
        self.assertIn("| `advice` | ⚠️ warn | advisory, not a failure: exit code 1 |", summary)

    def fake_actionlint(self, probe: tuple[int, str], workflows: int) -> list[Check]:
        """The real actionlint check, run by a fake actionlint on PATH.

        The fake answers the probe with an exit code and an output, and the
        lint of the workflows, which it runs without a file argument, with an
        exit code.
        """
        folder = self.root.parent / f"{self.root.name}-bin"
        folder.mkdir()
        self.addCleanup(shutil.rmtree, folder, ignore_errors=True)
        fake = folder / "actionlint"
        fake.write_text(f"#!{sys.executable}\nimport sys\n"
                        f"if {check_repository.ACTIONLINT_PROBE!r} in sys.argv:\n"
                        f"    print({probe[1]!r})\n    sys.exit({probe[0]})\n"
                        f"print('linted the workflows')\nsys.exit({workflows})\n",
                        encoding="utf-8")
        fake.chmod(0o755)
        with mock.patch.dict(os.environ, {"PATH": f"{folder}{os.pathsep}{os.environ['PATH']}"}):
            listed = [check for check in check_repository.checks(self.context())
                      if check.name == "actionlint"]
        self.assertEqual(listed[0].command, [str(fake), "-no-color"])
        self.commit()
        return listed

    REJECTED = (1, f"{check_repository.ACTIONLINT_PROBE}:19:9: "
                   f"{check_repository.ACTIONLINT_PROBE_FINDING} for step [syntax-check]")

    @unittest.skipIf(os.name == "nt", "the fake actionlint is a script with a shebang line")
    def test_actionlint_that_accepts_the_probe_fails_without_linting(self) -> None:
        listed = self.fake_actionlint((0, "no problems"), workflows=0)
        code, printed, report = self.run_checks(listed)
        self.assertEqual(code, 1, printed)
        self.assertIn("actionlint: FAIL (probe not rejected: exited 0)", printed)
        self.assertIn("the tool must exit non-zero and report 'unexpected key \"bogus-key\"'",
                      printed)
        self.assertNotIn("linted the workflows", printed)
        self.assertEqual((report["checks"][0]["status"], report["checks"][0]["exit_code"]),
                         ("fail", None))

    @unittest.skipIf(os.name == "nt", "the fake actionlint is a script with a shebang line")
    def test_actionlint_that_fails_the_probe_without_the_finding_fails(self) -> None:
        listed = self.fake_actionlint((1, "cannot read the file"), workflows=0)
        code, printed, _ = self.run_checks(listed)
        self.assertEqual(code, 1, printed)
        self.assertIn("actionlint: FAIL (probe not rejected: exited 1 without reporting "
                      "'unexpected key \"bogus-key\"')", printed)
        self.assertNotIn("linted the workflows", printed)

    @unittest.skipIf(os.name == "nt", "the fake actionlint is a script with a shebang line")
    def test_actionlint_that_rejects_the_probe_lints_the_workflows(self) -> None:
        listed = self.fake_actionlint(self.REJECTED, workflows=0)
        code, printed, report = self.run_checks(listed)
        self.assertEqual(code, 0, printed)
        self.assertIn("actionlint: PASS", printed)
        self.assertEqual(report["checks"][0]["exit_code"], 0)
        log = (self.root / "test-results/integrity/actionlint.log").read_text("utf-8")
        fake = self.root.parent / f"{self.root.name}-bin/actionlint"
        self.assertIn(f"probe: $ {fake} -no-color "
                      f"{check_repository.ACTIONLINT_PROBE}\n{self.REJECTED[1]}\n"
                      "probe: rejected as expected (exit code 1)\n", log)
        self.assertTrue(log.endswith("linted the workflows\n"), log)

    @unittest.skipIf(os.name == "nt", "the fake actionlint is a script with a shebang line")
    def test_actionlint_that_rejects_the_workflows_fails(self) -> None:
        listed = self.fake_actionlint(self.REJECTED, workflows=1)
        code, printed, report = self.run_checks(listed)
        self.assertEqual(code, 1, printed)
        self.assertIn("actionlint: FAIL (exit code 1)", printed)
        self.assertIn("linted the workflows", printed)
        self.assertEqual(report["checks"][0]["exit_code"], 1)

    def test_the_tests_check_fails_on_a_failing_or_broken_test(self) -> None:
        # The real check, run in a fixture whose tests/ holds only the module below.
        check = next(check for check in check_repository.checks(self.context())
                     if check.name == "tests")
        passing = ("import unittest\n\n\nclass ProbeTest(unittest.TestCase):\n"
                   "    def test_probe(self) -> None:\n        self.assertEqual(1, 1)\n")
        cases = [("passes", passing, "pass"),
                 ("fails", passing.replace("assertEqual(1, 1)", "assertEqual(1, 2)"), "fail"),
                 ("cannot be imported", "import a_module_that_does_not_exist\n", "fail")]
        if sys.version_info >= (3, 12):
            # From Python 3.12, a run that finds no test exits with 5.
            cases.append(("holds no test", "", "fail"))
        for index, (name, source, status) in enumerate(cases):
            with self.subTest(module=name):
                # A new module name each time, so that no cached bytecode is read.
                shutil.rmtree(self.root / "tests", ignore_errors=True)
                self.write({f"tests/test_probe_{index}.py": source})
                result = check_repository.run_check(check, self.context(), strict=True)
                self.assertEqual(result.status, status, result.log)

    def test_github_annotations(self) -> None:
        with mock.patch.dict(os.environ, {"GITHUB_ACTIONS": "true"}):
            self.assertEqual(check_repository.annotate(("a.md", "bad")), "::error file=a.md::bad")

    def test_main_detects_the_repository_and_takes_arguments(self) -> None:
        self.commit()
        git(self.root, "remote", "add", "origin", "https://github.com/Owner/Project.git")
        seen: list[Context] = []

        def listed(context: Context) -> list[Check]:
            seen.append(context)
            return [Check("ok", [sys.executable, "-c", "pass"])]

        with contextlib.redirect_stdout(io.StringIO()):
            code = check_repository.main(["--root", str(self.root), "--template", "false"],
                                         listed)
        self.assertEqual(code, 0)
        self.assertEqual((seen[0].repository, seen[0].default_branch, seen[0].is_template,
                          seen[0].ci), ("Owner/Project", None, False, False))
        with contextlib.redirect_stdout(io.StringIO()):
            check_repository.main(["--root", str(self.root), "--repository", "a/b",
                                   "--default-branch", "trunk", "--template", "", "--ci"],
                                  listed)
        self.assertEqual((seen[1].repository, seen[1].default_branch, seen[1].is_template,
                          seen[1].ci), ("a/b", "trunk", None, True))

    def test_real_check_list_names_each_check_once(self) -> None:
        self.commit({check_repository.TRAFFIC_WORKFLOW: "name: Traffic history\n"})
        names = [check.name for check in check_repository.checks(self.context())]
        self.assertEqual(len(names), len(set(names)))
        for expected in ("tests", "ruff", "mypy", "actionlint", "binaries", "secret-files",
                         "vba-format", "vba-errors", "vba-calls", "data-contract", "test-cases",
                         "derivations", "independent-implementations", "workbook-protection",
                         "changelog", "generated-docs", "restated-counts", "template-wording",
                         "markdown-links", "documented-commands", "verification-depth",
                         "codeql-default-branch", "whitespace-committed"):
            self.assertIn(expected, names)
        advisory = {check.name for check in check_repository.checks(self.context())
                    if check.advisory}
        self.assertNotIn("vba-errors", advisory)
        # Only --ci checks the working tree, after every other check.
        self.assertNotIn("clean-working-tree", names)
        in_ci = check_repository.checks(self.context(ci=True))
        self.assertEqual([check.name for check in in_ci], [*names, "clean-working-tree"])
        without = check_repository.checks(self.context(repository=None, is_template=None,
                                                       default_branch=None))
        missing = {check.name for check in without if check.missing}
        self.assertTrue({"issue-contact-links", "readme-traffic-badges", "template-only-content",
                         "codeql-default-branch"} <= missing)


if __name__ == "__main__":
    unittest.main()
