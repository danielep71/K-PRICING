"""Tests for the closing keyword check (tools/check_closing_keywords.py)."""

from __future__ import annotations

import contextlib
import io
import json
import os
import shutil
import sys
import tempfile
import unittest
import unittest.mock
import urllib.error
from email.message import Message
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import check_closing_keywords as closing  # noqa: E402
from check_closing_keywords import Reference, Request  # noqa: E402
from gitfixture import git, head, init  # noqa: E402

REPOSITORY = "example-org/sample-project"
KEYWORDS = ("close", "closes", "closed", "fix", "fixes", "fixed", "resolve", "resolves",
            "resolved")


def numbers(text: str, where: str = "abc1234") -> list[int]:
    return [reference.number for reference in closing.references(text, where, REPOSITORY)]


def run_main(*args: str, env: dict[str, str] | None = None) -> tuple[int, str, str]:
    output, errors = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(output), contextlib.redirect_stderr(errors), \
            unittest.mock.patch.dict(os.environ, env or {}, clear=False):
        try:
            code = closing.main(list(args))
        except SystemExit as stop:
            code = int(stop.code or 0)
    return code, output.getvalue(), errors.getvalue()


class ReferenceTests(unittest.TestCase):
    def test_every_keyword_form_with_and_without_a_colon_in_any_case(self) -> None:
        for keyword in KEYWORDS:
            for written in (keyword, keyword.upper(), keyword.capitalize()):
                for separator in (" ", ": ", ":"):
                    with self.subTest(text=f"{written}{separator}#12"):
                        self.assertEqual(numbers(f"{written}{separator}#12"), [12])

    def test_the_keyword_is_reported_as_written(self) -> None:
        found = closing.references("docs: x\n\nThis RESOLVES: #7 for now.", "abc1234",
                                   REPOSITORY)
        self.assertEqual(found, [Reference("abc1234", "RESOLVES: #7", 7)])

    def test_a_keyword_anywhere_in_the_message_counts(self) -> None:
        message = "fix: Resolve #4 in the subject\n\nThe body also closes #5 mid-sentence.\n"
        self.assertEqual(numbers(message), [4, 5])

    def test_each_reference_needs_its_own_keyword(self) -> None:
        # The two shapes of a lower-case subject and of a sentence in a body.
        self.assertEqual(numbers("docs: tidy the guide and resolve #31, #32"), [31])
        self.assertEqual(numbers("x\n\nThis resolves #31 and #32 together."), [31])
        self.assertEqual(numbers("Fixes #1, fixes #2"), [1, 2])

    def test_no_keyword_or_no_number_is_no_reference(self) -> None:
        for text in ("Refs #12", "Fixes the parser", "fixes #", "Fixes issue 12", "fixes#12",
                     "unfixed #3", "prefixes #3", "fixed-income #3", "Closes #12a",
                     "Fixes\n#12"):
            with self.subTest(text=text):
                self.assertEqual(numbers(text), [])

    def test_a_reference_to_this_repository_counts_and_another_is_ignored(self) -> None:
        self.assertEqual(numbers(f"Fixes {REPOSITORY}#9"), [9])
        self.assertEqual(numbers("Fixes Example-Org/Sample-Project#9"), [9])
        self.assertEqual(numbers("Fixes other-org/other-project#9"), [])
        self.assertEqual(numbers("Fixes example-org/sample-project-2#9"), [])


class AcknowledgementTests(unittest.TestCase):
    def test_lines_list_the_issues_to_reopen(self) -> None:
        body = ("Refs #3\n\nReopen after merge: #12\n- reopen after merge: #13, #14\n"
                "* Reopen after merge:#15\n")
        self.assertEqual(closing.acknowledged(body), {12, 13, 14, 15})

    def test_other_lines_and_html_comments_are_not_acknowledgements(self) -> None:
        body = ("<!--\nReopen after merge: #20\n-->\nPlease reopen after merge: #21\n"
                "Reopen after merge: #22 later\nReopen after merge: #N\n")
        self.assertEqual(closing.acknowledged(body), set())

    def test_crlf_line_endings_are_read(self) -> None:
        self.assertEqual(closing.acknowledged("Refs #1\r\nReopen after merge: #12\r\n"), {12})


class JudgeTests(unittest.TestCase):
    STATES = {1: ("open", False), 2: ("closed", False), 3: ("open", True), 4: ("open", False)}

    def judge(self, mode: str, found: list[Reference], body: str = "") -> list[tuple[str, int]]:
        request = Request(mode, REPOSITORY, None, None, "", body)
        outcomes = closing.judge(request, found, self.STATES.__getitem__)
        return [(outcome.level, outcome.reference.number) for outcome in outcomes]

    def test_on_a_release_branch_every_keyword_fails_without_reading_states(self) -> None:
        found = [Reference("abc1234", "Fixes #2", 2), Reference("title", "fixes #9", 9)]
        outcomes = closing.judge(Request("release", REPOSITORY, None, None, "", ""), found,
                                 lambda number: self.fail(f"read the state of #{number}"))
        self.assertEqual([(o.level, o.reference.number) for o in outcomes],
                         [("FAIL", 2), ("FAIL", 9)])

    def test_an_open_issue_fails_and_a_closed_issue_or_pull_request_is_a_note(self) -> None:
        found = [Reference("abc1234", "Fixes #1", 1), Reference("abc1234", "Fixes #2", 2),
                 Reference("abc1234", "Fixes #3", 3)]
        self.assertEqual(self.judge("integration", found),
                         [("FAIL", 1), ("NOTE", 2), ("NOTE", 3)])

    def test_an_acknowledged_commit_keyword_is_a_note(self) -> None:
        found = [Reference("abc1234", "Fixes #1", 1), Reference("abc1234", "Fixes #4", 4)]
        self.assertEqual(self.judge("integration", found, "Reopen after merge: #1\n"),
                         [("NOTE", 1), ("FAIL", 4)])

    def test_a_keyword_in_the_title_or_body_is_not_acknowledged(self) -> None:
        found = [Reference("title", "Fixes #1", 1), Reference("body", "fixes #1", 1)]
        self.assertEqual(self.judge("integration", found, "fixes #1\nReopen after merge: #1\n"),
                         [("FAIL", 1), ("FAIL", 1)])

    def test_a_finding_names_the_commit_the_keyword_and_the_issue(self) -> None:
        request = Request("integration", REPOSITORY, None, None, "", "")
        [outcome] = closing.judge(request, [Reference("abc1234", "resolves: #1", 1)],
                                  self.STATES.__getitem__)
        self.assertIn('commit abc1234: "resolves: #1"', outcome.line())
        self.assertIn("open issue #1", outcome.line())
        self.assertIn("Reopen after merge: #1", outcome.line())


class StateTests(unittest.TestCase):
    def test_the_api_answer_gives_the_state_and_whether_it_is_a_pull_request(self) -> None:
        answers = {"/issues/5": {"number": 5, "state": "closed"},
                   "/issues/6": {"number": 6, "state": "open", "pull_request": {"url": "x"}}}
        calls = []

        def fetch(url: str, token: str) -> Any:
            calls.append((url, token))
            return answers[url[url.rindex("/issues/"):]]

        with unittest.mock.patch.object(closing, "fetch", fetch):
            read = closing.live_states(REPOSITORY, "token-value")
            self.assertEqual((read(5), read(6), read(5)), (("closed", False), ("open", True),
                                                           ("closed", False)))
        self.assertEqual(calls, [(f"{closing.API}/repos/{REPOSITORY}/issues/5", "token-value"),
                                 (f"{closing.API}/repos/{REPOSITORY}/issues/6", "token-value")])

    def test_an_api_error_cannot_be_read(self) -> None:
        error = urllib.error.HTTPError("https://api.github.com", 404, "Not Found", Message(),
                                       None)
        with unittest.mock.patch("urllib.request.urlopen", side_effect=error), \
                self.assertRaisesRegex(closing.StateError, "#5.*404"):
            closing.live_states(REPOSITORY, "token-value")(5)

    def test_without_a_token_the_state_cannot_be_read(self) -> None:
        with unittest.mock.patch.dict(os.environ, {"GITHUB_TOKEN": ""}), \
                self.assertRaisesRegex(closing.StateError, "GITHUB_TOKEN"):
            closing.token_states(REPOSITORY)(5)

    def test_an_input_file_gives_the_states_and_refuses_a_missing_number(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "issues.json"
            path.write_text(json.dumps([{"number": 1, "state": "OPEN"},
                                        {"number": 2, "state": "closed", "pull_request": {}}]),
                            encoding="utf-8")
            read = closing.file_states(path)
            self.assertEqual((read(1), read(2)), (("open", False), ("closed", True)))
            with self.assertRaisesRegex(closing.StateError, "#3"):
                read(3)
            path.write_text("{", encoding="utf-8")
            with self.assertRaises(closing.StateError):
                closing.file_states(path)


class ModeTests(unittest.TestCase):
    def test_the_mode_follows_the_base_and_head(self) -> None:
        self.assertEqual(closing.mode_for("release/1.2.0", "feat/x", "main")[0], "release")
        self.assertEqual(closing.mode_for("main", "release/1.2.0", "main")[0], "integration")
        mode, reason = closing.mode_for("main", "feat/x", "main")
        self.assertEqual(mode, "none")
        self.assertIn("not checked", reason)
        self.assertEqual(closing.mode_for("develop", "release/1.2.0", "main")[0], "none")


@unittest.skipUnless(shutil.which("git"), "needs Git")
class RepositoryTests(unittest.TestCase):
    """base..head in a temporary repository: base on main, head on a release branch."""

    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory(prefix="closing-keywords-")
        self.root = Path(self.temporary.name)
        init(self.root)
        (self.root / "a.txt").write_text("a\n", encoding="utf-8")
        git(self.root, "add", "-A")
        git(self.root, "commit", "-qm", "chore: start\n\nFixes #100")
        self.base = head(self.root)
        git(self.root, "checkout", "-qb", "release/1.2.0")
        self.commit("feat: add b (#8)\n\nThe body resolves #1 and #2.\n\nRefs #9", "b.txt")
        git(self.root, "checkout", "-qb", "feat/side")
        self.commit("docs: side change\n\nCloses: #3", "c.txt")
        git(self.root, "checkout", "-q", "release/1.2.0")
        self.commit("fix: other", "d.txt")
        git(self.root, "merge", "--no-ff", "-q", "feat/side", "-m",
            "Merge pull request #10 from feat/side\n\nFIXED #4")
        self.head = head(self.root)
        self.states = self.root.parent / f"{self.root.name}-states.json"
        self.addCleanup(self.states.unlink, missing_ok=True)
        self.addCleanup(self.temporary.cleanup)

    def commit(self, message: str, name: str) -> None:
        (self.root / name).write_text(name, encoding="utf-8")
        git(self.root, "add", "-A")
        git(self.root, "commit", "-qm", message)

    def write_states(self, items: list[dict[str, Any]]) -> str:
        self.states.write_text(json.dumps(items), encoding="utf-8")
        return str(self.states)

    def main(self, mode: str, *extra: str, env: dict[str, str] | None = None,
             base: str | None = None) -> tuple[int, str, str]:
        return run_main("--root", str(self.root), "--repository", REPOSITORY, "--mode", mode,
                        "--base", base or self.base, "--head", self.head, *extra, env=env)

    def test_the_range_lists_every_commit_with_merge_commits_and_not_the_base(self) -> None:
        listed = closing.commits(self.root, self.base, self.head)
        self.assertEqual(len(listed), 4)
        self.assertTrue(listed[0][1].startswith("Merge pull request #10"))
        found = closing.collect(Request("release", REPOSITORY, self.base, self.head, "", ""),
                                self.root)[0]
        self.assertEqual(sorted(reference.number for reference in found), [1, 3, 4])
        self.assertTrue(all(len(reference.where) == 7 for reference in found))

    def test_into_a_release_branch_every_keyword_fails(self) -> None:
        code, output, _ = self.main("release", "--title", "Fix things", "--body",
                                    "Resolves #5", env={"GITHUB_TOKEN": ""})
        self.assertEqual(code, 1)
        self.assertIn("Commits read: 4", output)
        self.assertIn("findings: 4", output)
        sha = git(self.root, "rev-parse", "--short=7", "HEAD").strip()
        self.assertIn(f'commit {sha}: "FIXED #4"', output)
        self.assertIn('pull request body: "Resolves #5"', output)

    def test_an_integration_fails_on_an_open_issue_unless_acknowledged(self) -> None:
        states = self.write_states([{"number": 1, "state": "open"},
                                    {"number": 3, "state": "closed"},
                                    {"number": 4, "state": "open", "pull_request": {}}])
        code, output, _ = self.main("integration", "--input", states, "--body", "Refs #9")
        self.assertEqual(code, 1)
        self.assertIn("findings: 1", output)
        self.assertIn("would close open issue #1", output)
        code, output, _ = self.main("integration", "--input", states, "--body",
                                    "Refs #9\n\nReopen after merge: #1\n")
        self.assertEqual(code, 0, output)
        self.assertIn("findings: 0", output)
        self.assertIn("NOTE", output)

    def test_an_empty_range_passes(self) -> None:
        code, output, _ = self.main("integration", base=self.head, env={"GITHUB_TOKEN": ""})
        self.assertEqual(code, 0)
        self.assertIn("Commits read: 0", output)
        self.assertIn("Closing keywords found: 0", output)

    def test_an_unreadable_state_or_revision_exits_2(self) -> None:
        code, _, errors = self.main("integration", env={"GITHUB_TOKEN": ""})
        self.assertEqual(code, 2)
        self.assertIn("GITHUB_TOKEN", errors)
        code, _, errors = self.main("release", base="no-such-revision")
        self.assertEqual(code, 2)
        self.assertIn("no-such-revision", errors)

    def test_the_summary_receives_the_report(self) -> None:
        summary = self.root / "summary.md"
        code, output, _ = self.main("release", "--summary", str(summary))
        self.assertEqual(code, 1)
        self.assertEqual(summary.read_text(encoding="utf-8"), output)

    def event(self, base_ref: str, head_ref: str, body: str = "",
              head_repository: str | None = REPOSITORY, **pull: Any) -> str:
        path = self.root.parent / f"{self.root.name}-event.json"
        self.addCleanup(path.unlink, missing_ok=True)
        path.write_text(json.dumps({
            "repository": {"full_name": REPOSITORY, "default_branch": "main"},
            "pull_request": {"title": "Integrate", "body": body,
                             "base": {"ref": base_ref, "sha": self.base},
                             "head": {"ref": head_ref, "sha": self.head,
                                      "repo": ({"full_name": head_repository}
                                               if head_repository else None)}, **pull},
        }), encoding="utf-8")
        return path.as_posix()

    def test_an_event_gives_the_mode_range_and_text(self) -> None:
        states = self.write_states([{"number": n, "state": "closed"} for n in (1, 3, 4)])
        code, output, _ = run_main("--root", str(self.root), "--input", states,
                                   "--event", self.event("main", "release/1.2.0"))
        self.assertEqual((code, "Mode: integration" in output), (0, True), output)
        code, output, _ = run_main("--root", str(self.root),
                                   "--event", self.event("release/1.2.0", "feat/x"))
        self.assertEqual((code, "Mode: release" in output), (1, True), output)

    def test_a_sync_event_reads_only_the_title_and_body(self) -> None:
        code, output, _ = run_main("--root", str(self.root),
                                   "--event", self.event("release/1.2.0", "main"))
        self.assertEqual(code, 0, output)
        self.assertIn("commits not checked", output.lower())
        code, output, _ = run_main("--root", str(self.root),
                                   "--event", self.event("release/1.2.0", "main", "Fixes #7"))
        self.assertEqual(code, 1, output)

    def test_a_fork_is_never_a_sync_or_an_integration(self) -> None:
        # A fork's main into a release branch: its commits are read, as any task's.
        code, output, _ = run_main("--root", str(self.root), "--event",
                                   self.event("release/1.2.0", "main",
                                              head_repository="someone/fork"))
        self.assertEqual((code, "Mode: release" in output), (1, True), output)
        self.assertNotIn("commits not checked", output.lower())
        # A fork's release branch into the default branch is no integration.
        for repository in ("someone/fork", None):
            with self.subTest(repository=repository):
                code, output, _ = run_main("--root", str(self.root), "--event",
                                           self.event("main", "release/1.2.0",
                                                      head_repository=repository))
                self.assertEqual((code, "Mode: none" in output), (0, True), output)

    def test_another_pull_request_is_reported_as_not_checked(self) -> None:
        code, output, _ = run_main("--root", str(self.root),
                                   "--event", self.event("main", "feat/x", "Fixes #7"))
        self.assertEqual(code, 0)
        self.assertIn("Mode: none", output)
        self.assertIn("not checked", output.lower())
        self.assertNotIn("Commits read", output)

    def test_an_event_without_a_pull_request_exits_2(self) -> None:
        path = self.root / "push.json"
        path.write_text(json.dumps({"ref": "refs/heads/main"}), encoding="utf-8")
        code, _, errors = run_main("--root", str(self.root), "--event", str(path))
        self.assertEqual(code, 2)
        self.assertIn("not a pull_request event", errors)


class UsageTests(unittest.TestCase):
    def test_bad_usage_exits_2(self) -> None:
        for args in ((), ("--repository", REPOSITORY, "--mode", "release", "--base", "a"),
                     ("--repository", "not-a-repository", "--mode", "release", "--base", "a",
                      "--head", "b"),
                     ("--repository", REPOSITORY, "--mode", "task", "--base", "a",
                      "--head", "b"),
                     ("--event", "event.json", "--mode", "release")):
            with self.subTest(args=args):
                code, _, errors = run_main(*args)
                self.assertEqual(code, 2)
                self.assertIn("error:", errors)


if __name__ == "__main__":
    unittest.main()
