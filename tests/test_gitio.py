"""Tests for tools/gitio.py, the one way the tools run Git, in temporary repositories."""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import gitio  # noqa: E402
from gitfixture import commit, init  # noqa: E402
from gitio import GitError  # noqa: E402


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class GitTest(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="gitio-")
        self.addCleanup(directory.cleanup)
        self.root = init(Path(directory.name) / "work")

    def write(self, name: str, data: bytes = b"x\n") -> None:
        path = self.root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)

    def test_success_returns_the_output(self) -> None:
        self.write("a.txt", b"line\r\n")
        commit(self.root)
        self.assertEqual(gitio.git(self.root, "show", "HEAD:a.txt"), b"line\r\n")
        # As text, line ends read as Python's text mode reads them.
        self.assertEqual(gitio.text(self.root, "show", "HEAD:a.txt"), "line\n")
        self.assertEqual(gitio.text(self.root, "hash-object", "--stdin", input=""),
                         "e69de29bb2d1d6434b8b29ae775ad8c2e48c5391\n")
        self.assertEqual(gitio.git(self.root, "cat-file", "--batch-check",
                                   input=b"HEAD:a.txt\n").split()[1], b"blob")

    def test_failure_raises_git_error_with_stderr(self) -> None:
        with self.assertRaises(GitError) as raised:
            gitio.git(self.root, "show", "HEAD:missing.txt")
        error = raised.exception
        self.assertIsInstance(error, RuntimeError)
        self.assertEqual(error.command[:4], ["git", "-C", str(self.root), "show"])
        self.assertEqual(error.returncode, 128)
        self.assertTrue(error.stderr.strip())
        self.assertEqual(str(error), error.stderr.strip())
        self.assertIn("HEAD", str(error))
        self.assertIsNone(gitio.text_or_none(self.root, "show", "HEAD:missing.txt"))

    def test_a_failure_without_stderr_names_the_command(self) -> None:
        self.write("a.txt")
        # check-ignore -q says "not ignored" by its exit status alone.
        with self.assertRaises(GitError) as raised:
            gitio.git(self.root, "check-ignore", "-q", "--", "a.txt")
        self.assertEqual(raised.exception.stderr, "")
        self.assertEqual(raised.exception.returncode, 1)
        self.assertIn("check-ignore -q -- a.txt exited with 1", str(raised.exception))

    def test_git_that_cannot_start_raises_git_error(self) -> None:
        with mock.patch.object(subprocess, "run", side_effect=FileNotFoundError("no git")):
            with self.assertRaises(GitError) as raised:
                gitio.git(self.root, "status")
            self.assertIsNone(raised.exception.returncode)
            self.assertIn("no git", str(raised.exception))
            self.assertIsNone(gitio.text_or_none(self.root, "status"))

    def test_tracked_lists_sorted_paths_with_non_ascii_names(self) -> None:
        for name in ("z.txt", "docs/Übersicht.md", "docs/a b.md", "dir/naïve.bas"):
            self.write(name)
        self.write("untracked.txt")
        gitio.git(self.root, "add", "--", "z.txt", "docs", "dir")
        self.assertEqual(gitio.tracked(self.root),
                         ["dir/naïve.bas", "docs/a b.md", "docs/Übersicht.md", "z.txt"])
        self.assertEqual(gitio.tracked(self.root, "*.md"), ["docs/a b.md", "docs/Übersicht.md"])
        self.assertEqual(gitio.tracked(self.root, "dir"), ["dir/naïve.bas"])

    def test_tracked_raises_outside_a_repository(self) -> None:
        with tempfile.TemporaryDirectory() as folder, self.assertRaises(GitError):
            gitio.tracked(Path(folder) / "missing")

    def test_tracked_or_scanned_scans_without_git(self) -> None:
        self.write("a.md")
        self.write("notes/b.md")
        gitio.git(self.root, "add", "a.md")
        self.assertEqual(gitio.tracked_or_scanned(self.root, "*.md"), ["a.md"])
        with tempfile.TemporaryDirectory() as folder:
            plain = Path(folder)
            for name in ("a.md", "notes/b.md", ".git/c.md", "d.txt"):
                (plain / name).parent.mkdir(parents=True, exist_ok=True)
                (plain / name).write_text("x\n", encoding="utf-8")
            (plain / "folder.md").mkdir()
            self.assertEqual(gitio.tracked_or_scanned(plain, "*.md"), ["a.md", "notes/b.md"])
            self.assertEqual(gitio.tracked_or_scanned(plain),
                             ["a.md", "d.txt", "notes/b.md"])


if __name__ == "__main__":
    unittest.main()
