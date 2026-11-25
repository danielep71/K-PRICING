"""Tests for tools/make_mutant.py, in temporary Git repositories."""

from __future__ import annotations

import contextlib
import hashlib
import io
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import make_mutant  # noqa: E402
from gitfixture import git, init  # noqa: E402

MODULE = "src/modules/M_Sample.bas"
CRLF_MODULE = (b'Attribute VB_Name = "M_Sample"\r\n'
               b"Option Explicit\r\n"
               b"\r\n"
               b"Public Function Total(ByVal a As Double, ByVal b As Double) As Double\r\n"
               b"    If a > 0 Then\r\n"
               b"        Total = a + b\r\n"
               b"    End If\r\n"
               b"End Function\r\n")


class Fixture(unittest.TestCase):
    """A temporary Git repository that ignores /test-results/ and holds MODULE."""

    def setUp(self) -> None:
        temporary = tempfile.TemporaryDirectory(prefix="mutant-")
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name).resolve()
        (self.root / ".gitignore").write_text("/test-results/\n", encoding="utf-8")
        self.put(MODULE, CRLF_MODULE)
        init(self.root)
        git(self.root, "add", "-A")
        git(self.root, "commit", "-q", "-m", "fixture")

    def put(self, name: str, data: bytes) -> Path:
        path = self.root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        return path

    def run_tool(self, *args: str, file: str = MODULE) -> tuple[int, str, str]:
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            try:
                code = make_mutant.main([str(self.root / file), "--root", str(self.root), *args])
            except SystemExit as exit_:
                code = int(exit_.code or 0)
        return code, out.getvalue(), err.getvalue()

    @property
    def mutants(self) -> Path:
        return self.root / "test-results" / "mutants"

    def assert_nothing_written(self) -> None:
        self.assertFalse(self.mutants.exists())
        self.assertEqual((self.root / MODULE).read_bytes(), CRLF_MODULE)


@unittest.skipUnless(shutil.which("git"), "needs Git")
class MakeMutantTest(Fixture):
    def test_one_occurrence_is_replaced_under_the_same_relative_path(self) -> None:
        code, out, err = self.run_tool("--find", "a > 0", "--replace", "a >= 0")
        self.assertEqual(code, 0, err)
        mutant = self.mutants / MODULE
        self.assertEqual(mutant.read_bytes(), CRLF_MODULE.replace(b"a > 0", b"a >= 0"))
        self.assertIn("mutant: test-results/mutants/src/modules/M_Sample.bas\n", out)
        self.assertIn(f"changed at {MODULE} line 5:\n"
                      "-     If a > 0 Then\n"
                      "+     If a >= 0 Then\n", out)
        self.assertEqual((self.root / MODULE).read_bytes(), CRLF_MODULE)

    def test_zero_occurrences_are_refused(self) -> None:
        code, _, err = self.run_tool("--find", "a < 0", "--replace", "a > 0")
        self.assertEqual(code, 2)
        self.assertIn("occurs 0 times", err)
        self.assert_nothing_written()

    def test_two_occurrences_are_refused(self) -> None:
        code, _, err = self.run_tool("--find", "Total", "--replace", "Sum")
        self.assertEqual(code, 2)
        self.assertIn("occurs 2 times", err)
        self.assert_nothing_written()

    def test_overlapping_occurrences_are_refused(self) -> None:
        with self.assertRaisesRegex(make_mutant.Refusal, "occurs 2 times"):
            make_mutant.mutate(b"aaa", b"aa", b"X", False)
        with self.assertRaisesRegex(make_mutant.Refusal, "occurs 3 times"):
            make_mutant.mutate(b"abababa", b"aba", b"X", False)
        self.put("notes/run.txt", b"x = aaa\n")
        code, _, err = self.run_tool("--find", "aa", "--replace", "b", file="notes/run.txt")
        self.assertEqual(code, 2)
        self.assertIn("occurs 2 times", err)
        self.assertFalse(self.mutants.exists())

    def test_an_unchanged_replacement_is_refused(self) -> None:
        code, _, err = self.run_tool("--find", "a > 0", "--replace", "a > 0")
        self.assertEqual(code, 2)
        self.assertIn("would not differ", err)
        self.assert_nothing_written()

    def test_an_empty_text_to_find_is_a_usage_error(self) -> None:
        code, _, err = self.run_tool("--find", "", "--replace", "x")
        self.assertEqual(code, 2)
        self.assert_nothing_written()

    def test_crlf_is_kept_byte_for_byte_and_multi_line_text_follows_it(self) -> None:
        find = self.put("find.txt", b"    If a > 0 Then\n        Total = a + b\n")
        replace = self.put("replace.txt", b"        Total = a - b\n")
        code, out, err = self.run_tool("--find-file", str(find), "--replace-file", str(replace))
        self.assertEqual(code, 0, err)
        expected = CRLF_MODULE.replace(b"    If a > 0 Then\r\n        Total = a + b\r\n",
                                       b"        Total = a - b\r\n")
        mutant = (self.mutants / MODULE).read_bytes()
        self.assertEqual(mutant, expected)
        self.assertEqual(mutant.count(b"\n"), mutant.count(b"\r\n"))
        self.assertIn("line 5:\n-     If a > 0 Then\n-         Total = a + b\n"
                      "+         Total = a - b\n", out)

    def test_lf_text_on_one_line_keeps_an_lf_file_lf(self) -> None:
        source = self.put("notes/plain.txt", b"one\ntwo\nthree\n")
        code, _, err = self.run_tool("--find", "two", "--replace", "2", file="notes/plain.txt")
        self.assertEqual(code, 0, err)
        self.assertEqual((self.mutants / "notes/plain.txt").read_bytes(), b"one\n2\nthree\n")
        self.assertEqual(source.read_bytes(), b"one\ntwo\nthree\n")

    def test_multi_line_text_in_a_file_with_mixed_line_endings_is_refused(self) -> None:
        self.put("notes/mixed.txt", b"one\r\ntwo\nthree\n")
        find = self.put("find.txt", b"one\ntwo\n")
        code, _, err = self.run_tool("--find-file", str(find), "--replace", "x",
                                     file="notes/mixed.txt")
        self.assertEqual(code, 2)
        self.assertIn("mixes CRLF and LF", err)
        self.assertFalse(self.mutants.exists())

    def test_vb_name_is_kept(self) -> None:
        code, _, err = self.run_tool("--find", "Option Explicit", "--replace", "Option Base 1")
        self.assertEqual(code, 0, err)
        self.assertTrue((self.mutants / MODULE).read_bytes()
                        .startswith(b'Attribute VB_Name = "M_Sample"\r\n'))

    def test_a_replacement_that_touches_vb_name_is_refused(self) -> None:
        cases = {"changed": ('"M_Sample"', '"M_Other"'),
                 "removed": ('Attribute VB_Name = "M_Sample"', "' no name"),
                 "added": ("Option Explicit", 'Option Explicit\nAttribute VB_Name = "M_Two"')}
        for label, (find, replace) in cases.items():
            with self.subTest(label):
                code, _, err = self.run_tool("--find", find, "--replace", replace)
                self.assertEqual(code, 2)
                self.assertIn("Attribute VB_Name", err)
                self.assert_nothing_written()

    def test_a_component_without_vb_name_is_refused(self) -> None:
        self.put("src/modules/M_Bare.bas", b"Option Explicit\r\nPublic X As Long\r\n")
        code, _, err = self.run_tool("--find", "Long", "--replace", "Integer",
                                     file="src/modules/M_Bare.bas")
        self.assertEqual(code, 2)
        self.assertIn("no Attribute VB_Name", err)

    def test_a_non_ascii_replacement_in_vba_is_refused(self) -> None:
        code, _, err = self.run_tool("--find", "a + b", "--replace", "a − b")
        self.assertEqual(code, 2)
        self.assertIn("not ASCII", err)
        self.assert_nothing_written()

    def test_a_form_mutant_gets_its_frx(self) -> None:
        form = b'VERSION 5.00\r\nAttribute VB_Name = "F_Input"\r\nPrivate X As Long\r\n'
        self.put("src/forms/F_Input.frm", form)
        self.put("src/forms/F_Input.frx", b"\0\1binary")
        code, out, err = self.run_tool("--find", "Long", "--replace", "Integer",
                                       file="src/forms/F_Input.frm")
        self.assertEqual(code, 0, err)
        self.assertEqual((self.mutants / "src/forms/F_Input.frx").read_bytes(), b"\0\1binary")
        self.assertIn("copied src/forms/F_Input.frx unchanged beside it", out)

    def test_printed_hashes_are_those_of_the_source_and_the_mutant(self) -> None:
        code, out, err = self.run_tool("--find", "a + b", "--replace", "a * b")
        self.assertEqual(code, 0, err)
        mutant = (self.mutants / MODULE).read_bytes()
        self.assertIn(f"source SHA-256 {hashlib.sha256(CRLF_MODULE).hexdigest()}\n", out)
        self.assertIn(f"mutant SHA-256 {hashlib.sha256(mutant).hexdigest()}\n", out)
        self.assertNotEqual(mutant, CRLF_MODULE)

    def test_running_again_replaces_the_mutant(self) -> None:
        self.assertEqual(self.run_tool("--find", "a + b", "--replace", "a * b")[0], 0)
        self.assertEqual(self.run_tool("--find", "a + b", "--replace", "a / b")[0], 0)
        self.assertIn(b"a / b", (self.mutants / MODULE).read_bytes())

    def test_an_output_folder_git_does_not_ignore_is_refused(self) -> None:
        code, _, err = self.run_tool("--find", "a + b", "--replace", "a * b",
                                     "--output", str(self.root / "scratch"))
        self.assertEqual(code, 2)
        self.assertIn("Git does not ignore scratch/src/modules/M_Sample.bas", err)
        self.assertFalse((self.root / "scratch").exists())

    def test_an_output_folder_in_a_source_directory_is_refused(self) -> None:
        # Even when a local rule ignores it, a source directory never holds a mutant.
        (self.root / ".git" / "info" / "exclude").write_text("/src/scratch/\n", encoding="utf-8")
        code, _, err = self.run_tool("--find", "a + b", "--replace", "a * b",
                                     "--output", str(self.root / "src" / "scratch"))
        self.assertEqual(code, 2)
        self.assertIn("in the source directory src/", err)
        self.assertFalse((self.root / "src" / "scratch").exists())

    def test_an_output_folder_holding_tracked_files_is_refused(self) -> None:
        self.put("test-results/kept/evidence.txt", b"kept\n")
        git(self.root, "add", "-f", "test-results/kept/evidence.txt")
        code, _, err = self.run_tool("--find", "a + b", "--replace", "a * b",
                                     "--output", str(self.root / "test-results" / "kept"))
        self.assertEqual(code, 2)
        self.assertIn("holds tracked files", err)
        self.assertFalse((self.root / "test-results/kept/src").exists())

    def test_a_file_outside_the_repository_is_refused(self) -> None:
        with tempfile.TemporaryDirectory(prefix="outside-") as other:
            outside = Path(other) / "M_Out.bas"
            outside.write_bytes(CRLF_MODULE)
            code, _, err = self.run_tool("--find", "a + b", "--replace", "a * b",
                                         file=str(outside))
        self.assertEqual(code, 2)
        self.assertIn("outside the repository", err)
        self.assertFalse(self.mutants.exists())


def symlink(link: Path, target: Path, directory: bool = False) -> bool:
    """Create ``link`` to ``target``; False where the platform or user cannot."""
    link.parent.mkdir(parents=True, exist_ok=True)
    try:
        link.symlink_to(target, target_is_directory=directory)
    except (OSError, NotImplementedError):
        return False
    return True


@unittest.skipUnless(shutil.which("git"), "needs Git")
class SymlinkDestinationTest(Fixture):
    """A link left at the mutant's path never redirects the write."""

    def link_or_skip(self, link: Path, target: Path, directory: bool = False) -> None:
        if not symlink(link, target, directory):
            self.skipTest("cannot create symbolic links here")

    def assert_refused(self, code: int, err: str) -> None:
        self.assertEqual(code, 2)
        self.assertIn("symbolic link", err)

    def test_a_link_at_the_mutant_path_to_the_source_is_refused(self) -> None:
        self.link_or_skip(self.mutants / MODULE, self.root / MODULE)
        code, _, err = self.run_tool("--find", "a + b", "--replace", "a * b")
        self.assertEqual((self.root / MODULE).read_bytes(), CRLF_MODULE)
        self.assert_refused(code, err)

    def test_a_link_at_the_mutant_path_elsewhere_is_refused(self) -> None:
        elsewhere = self.put("notes/elsewhere.txt", b"kept\n")
        self.link_or_skip(self.mutants / MODULE, elsewhere)
        code, _, err = self.run_tool("--find", "a + b", "--replace", "a * b")
        self.assertEqual(elsewhere.read_bytes(), b"kept\n")
        self.assert_refused(code, err)

    def test_a_linked_folder_under_the_output_is_refused(self) -> None:
        self.link_or_skip(self.mutants / "src", self.root / "src", directory=True)
        code, _, err = self.run_tool("--find", "a + b", "--replace", "a * b")
        self.assertEqual((self.root / MODULE).read_bytes(), CRLF_MODULE)
        self.assert_refused(code, err)

    def test_a_link_at_the_frx_path_is_refused_and_nothing_is_written(self) -> None:
        form = b'VERSION 5.00\r\nAttribute VB_Name = "F_Input"\r\nPrivate X As Long\r\n'
        self.put("src/forms/F_Input.frm", form)
        frx = self.put("src/forms/F_Input.frx", b"\0\1binary")
        self.link_or_skip(self.mutants / "src/forms/F_Input.frx", frx)
        code, _, err = self.run_tool("--find", "Long", "--replace", "Integer",
                                     file="src/forms/F_Input.frm")
        self.assertEqual(frx.read_bytes(), b"\0\1binary")
        self.assertFalse((self.mutants / "src/forms/F_Input.frm").exists())
        self.assert_refused(code, err)


def hard_link(link: Path, target: Path) -> bool:
    """Create ``link`` as a hard link to ``target``; False where that is impossible."""
    link.parent.mkdir(parents=True, exist_ok=True)
    try:
        os.link(target, link)
    except (OSError, NotImplementedError):
        return False
    return True


@unittest.skipUnless(shutil.which("git"), "needs Git")
class HardLinkDestinationTest(Fixture):
    """A hard link left at the mutant's path is replaced, never written through."""

    def link_or_skip(self, link: Path, target: Path) -> None:
        if not hard_link(link, target):
            self.skipTest("cannot create hard links here")

    def test_a_hard_link_to_the_source_is_replaced_and_the_source_kept(self) -> None:
        mutant = self.mutants / MODULE
        self.link_or_skip(mutant, self.root / MODULE)
        code, _, err = self.run_tool("--find", "a + b", "--replace", "a * b")
        self.assertEqual((self.root / MODULE).read_bytes(), CRLF_MODULE)
        self.assertEqual(code, 0, err)
        self.assertEqual(mutant.read_bytes(), CRLF_MODULE.replace(b"a + b", b"a * b"))
        self.assertFalse(os.path.samefile(mutant, self.root / MODULE))

    def test_a_hard_link_to_the_frx_is_replaced(self) -> None:
        self.put("src/forms/F_Input.frm",
                 b'VERSION 5.00\r\nAttribute VB_Name = "F_Input"\r\nPrivate X As Long\r\n')
        frx = self.put("src/forms/F_Input.frx", b"\0\1binary")
        copy = self.mutants / "src/forms/F_Input.frx"
        self.link_or_skip(copy, frx)
        code, _, err = self.run_tool("--find", "Long", "--replace", "Integer",
                                     file="src/forms/F_Input.frm")
        self.assertEqual(code, 0, err)
        self.assertEqual(frx.read_bytes(), b"\0\1binary")
        self.assertEqual(copy.read_bytes(), b"\0\1binary")
        self.assertFalse(os.path.samefile(copy, frx))

    def test_no_temporary_file_is_left_behind(self) -> None:
        self.assertEqual(self.run_tool("--find", "a + b", "--replace", "a * b")[0], 0)
        self.assertEqual(sorted(p.name for p in (self.mutants / MODULE).parent.iterdir()),
                         ["M_Sample.bas"])


@unittest.skipUnless(shutil.which("git"), "needs Git")
class DefaultOutputTest(unittest.TestCase):
    def test_the_default_output_folder_is_ignored_in_this_repository(self) -> None:
        path = (make_mutant.DEFAULT_OUTPUT / MODULE).as_posix()
        ignored = subprocess.run(["git", "check-ignore", "-q", "--", path], cwd=ROOT,
                                 capture_output=True, check=False)
        self.assertEqual(ignored.returncode, 0, f"{path} is not ignored by Git")
        self.assertNotIn(make_mutant.DEFAULT_OUTPUT.parts[0], make_mutant.SOURCE_DIRECTORIES)


if __name__ == "__main__":
    unittest.main()
