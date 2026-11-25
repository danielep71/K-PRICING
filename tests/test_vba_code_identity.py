"""Positive and negative fixtures for vba_code_identity.py.

1. Comments and layout are not code: a changed or added apostrophe comment, a
   Rem line, a Rem after a colon, a comment continued with " _", indentation,
   trailing blanks, blank lines, CRLF line ends, a logical line joined or
   re-wrapped with continuations, runs of spaces between tokens and the letter
   case of code leave the code stream unchanged.
2. Code is code: each listed edit of a literal, a string, a statement, a
   directive or a statement inside #If is a code change, and so is white space
   removed between two tokens.
3. A component that uses [ or ] outside a string literal is never carried,
   even unchanged; brackets inside a string literal do not count.
4. The comparison fails closed on a non-ASCII byte.
5. On a real repository: comment-only edits of a .bas, a .cls and a .frm are
   carried; a code edit, an unknown commit, an empty commit name, a deleted
   component and a .frx are refused.
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from typing import ClassVar

SOURCE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SOURCE / "tools"))
sys.path.insert(0, str(SOURCE / "tests"))

import vba_code_identity as identity  # noqa: E402
from gitfixture import git, head, init  # noqa: E402

BASE = (
    'Attribute VB_Name = "CORE_Probe"\n'
    "Option Explicit\n"
    "' Module comment\n"
    "#Const FAST = 1\n"
    "Public Function Run(ByVal X As Double) As Double\n"
    "    Dim Total As Double, Remainder As Double   ' locals\n"
    '    Const LABEL As String = "it\'s  a \'quoted\' label"\n'
    "    Rem an old-style comment\n"
    "#If FAST Then\n"
    "    Total = X * 2#\n"
    "#Else\n"
    "    Total = X + X\n"
    "#End If\n"
    "    Remainder = 1: Rem trailing remark\n"
    "    Run = Total + _\n"
    "        Remainder\n"
    "End Function\n"
)

# Edits that keep the code.
SAME = {
    "a changed comment": BASE.replace("' Module comment", "' Another comment entirely"),
    "an added comment line": BASE.replace("Option Explicit\n", "Option Explicit\n' new\n"),
    "a removed trailing comment": BASE.replace("   ' locals", ""),
    "a changed Rem line": BASE.replace("Rem an old-style comment", "Rem rewritten"),
    "a changed Rem after a colon": BASE.replace("Rem trailing remark", "Rem other"),
    "a comment continued with _": BASE.replace(
        "' Module comment", "' Module comment _\n    Total = 99 is still comment"),
    "indentation": BASE.replace("    Total = X * 2#", "\t\tTotal = X * 2#"),
    "trailing blanks": BASE.replace("Option Explicit\n", "Option Explicit   \n"),
    "blank lines": BASE.replace("Option Explicit\n", "Option Explicit\n\n\n"),
    "CRLF line ends": BASE.replace("\n", "\r\n"),
    "a joined continuation": BASE.replace("Run = Total + _\n        Remainder",
                                          "Run = Total + Remainder"),
    "a new continuation": BASE.replace("Dim Total As Double, Remainder",
                                       "Dim Total As Double, _\n        Remainder"),
    "runs of spaces between tokens": BASE.replace("Total = X + X", "Total  =  X   +   X"),
    "the letter case of code (HB-01)": BASE.replace("Total = X + X", "total = x + X"),
}

BRACKETS = BASE.replace("        Remainder\n", "        Remainder + [A1 : B2]\n")
# Edits that change the code: the new text, or (new, old) when the old is not BASE.
DIFFERENT: dict[str, str | tuple[str, str]] = {
    "a changed numeric literal": BASE.replace("X * 2#", "X * 3#"),
    "a changed string literal": BASE.replace("label", "lAbel"),
    "whitespace inside a string": BASE.replace("it's  a", "it's a"),
    "an apostrophe inside a string": BASE.replace("it's", "its"),
    "the text after an apostrophe in a string": BASE.replace("'quoted' label", "'quoted' LABEL"),
    "a statement starting with Rem...": BASE.replace("Remainder = 1:", "Remainder = 2:"),
    "a statement inside #If": BASE.replace("Total = X + X", "Total = X + X + 0"),
    "an #If condition": BASE.replace("#If FAST Then", "#If Not FAST Then"),
    "a #Const value": BASE.replace("#Const FAST = 1", "#Const FAST = 0"),
    "a statement commented out": BASE.replace("    Total = X * 2#", "    ' Total = X * 2#"),
    "a statement turned into Rem": BASE.replace("    Total = X * 2#", "    Rem Total = X * 2#"),
    "a comment uncommented": BASE.replace("' Module comment", "Dim Leak As Long"),
    "a [...] expression": (BRACKETS.replace("[A1 : B2]", "[A1:B2]"), BRACKETS),
    "whitespace in a string inside [...]": (BRACKETS.replace("[A1 : B2]", '["]  "]'),
                                            BRACKETS.replace("[A1 : B2]", '["] "]')),
    "an added [...] expression": BRACKETS,
    "whitespace after a string that holds a bracket": (
        BASE.replace('"it\'s  a \'quoted\' label"', '"[" & "a  b"'),
        BASE.replace('"it\'s  a \'quoted\' label"', '"[" & "a b"')),
    "a name Rem inside [...]": (BRACKETS.replace("[A1 : B2]", "[A1:Rem + 2]"),
                                BRACKETS.replace("[A1 : B2]", "[A1:Rem + 1]")),
    "whitespace removed between tokens": BASE.replace("Total = X + X", "Total=X+X"),
    "two statements swapped": BASE.replace(
        "    Total = X * 2#\n#Else\n    Total = X + X\n",
        "    Total = X + X\n#Else\n    Total = X * 2#\n"),
    "a removed statement": BASE.replace("    Remainder = 1: Rem trailing remark\n", ""),
    "an added statement": BASE.replace("End Function\n",
                                       "End Function\nPublic Sub Probe()\nEnd Sub\n"),
    "a continuation that hides code in a comment": BASE.replace(
        "    Dim Total As Double, Remainder As Double   ' locals\n",
        "    Dim Total As Double   ' locals _\n    Dim Remainder As Double\n"),
}


def text(label: str) -> str:
    """The new text of the code change ``label``, made to BASE."""
    change = DIFFERENT[label]
    assert isinstance(change, str), label
    return change


def difference(new: str, old: str = BASE) -> str | None:
    return identity.first_difference(old.encode("ascii"), new.encode("ascii"))


class CodeStreamTest(unittest.TestCase):
    def test_comments_and_layout_are_not_code(self) -> None:
        for label, text in SAME.items():
            with self.subTest(label):
                self.assertNotEqual(text, BASE)
                self.assertIsNone(difference(text))

    def test_code_is_code(self) -> None:
        for label, change in DIFFERENT.items():
            new, old = change if isinstance(change, tuple) else (change, BASE)
            with self.subTest(label):
                self.assertNotEqual(new, old)
                self.assertIsNotNone(difference(new, old))

    def test_the_difference_names_the_lines_and_the_statements(self) -> None:
        found = difference(BASE.replace("X * 2#", "X * 3#"))
        self.assertEqual(found, "line 10 at the run, line 10 now: 'Total = X * 2#' became "
                                "'Total = X * 3#'")
        found = difference(SAME["blank lines"].replace("X * 2#", "X * 3#"))
        self.assertEqual(found, "line 10 at the run, line 12 now: 'Total = X * 2#' became "
                                "'Total = X * 3#'")
        self.assertIn("2 statement(s) only in the current version, the first at line 18",
                      difference(text("an added statement")) or "")

    def test_explicit_fixtures(self) -> None:
        # Rem after a colon ends the statement; a statement named Rem... does not.
        stream = identity.code_stream(b"x = 1: Rem note\nRemainder = 2\n")
        self.assertEqual([item.code for item in stream], ["x = 1:", "Remainder = 2"])
        # A continuation joins the lines; a continued comment stays comment.
        stream = identity.code_stream(b"x = 1 + _\n    2\ny = 3 ' note _\n  z = 4\n")
        self.assertEqual([item.code for item in stream], ["x = 1 + 2", "y = 3"])
        # An apostrophe inside a string is not a comment.
        stream = identity.code_stream(b'MsgBox "it\'s": x = 1 \' note\n')
        self.assertEqual([item.code for item in stream], ['MsgBox "it\'s": x = 1'])
        # A statement inside #If is code, and the directives are kept as such.
        stream = identity.code_stream(b"#If Win64 Then\n    x = 1\n#End If\n")
        self.assertEqual([(item.kind, item.code) for item in stream],
                         [("directive", "#If Win64 Then"), ("code", "x = 1"),
                          ("directive", "#End If")])
        self.assertIsNotNone(identity.first_difference(b"#If Win64 Then\n    x = 1\n#End If\n",
                                                       b"#If Win64 Then\n    x = 2\n#End If\n"))
        # A ] inside quotes is a string, not a [...] expression.
        quoted = b'x = "]" \' note\n'
        self.assertIsNone(identity.first_difference(quoted, b'x = "]" \' other note\n'))
        self.assertIsNotNone(identity.first_difference(b"x = [A1] ' note\n",
                                                       b"x = [A1] ' other note\n"))

    def test_re_cased_identifiers_are_reported(self) -> None:
        old, new = BASE.encode("ascii"), SAME["the letter case of code (HB-01)"].encode("ascii")
        self.assertEqual(identity.recased(old, new), ["Total -> total", "X -> x"])
        self.assertEqual(identity.recased(old, old), [])


class BracketTest(unittest.TestCase):
    def test_an_unchanged_component_with_brackets_is_refused(self) -> None:
        self.assertIn("never carried", difference(BRACKETS, BRACKETS) or "")

    def test_a_comment_edit_beside_brackets_is_refused(self) -> None:
        self.assertIsNotNone(difference(BRACKETS.replace("' Module comment", "' edited"),
                                        BRACKETS))

    def test_brackets_inside_a_string_literal_do_not_count(self) -> None:
        text = BASE.replace('"it\'s  a', '"[it\'s]  a')
        self.assertIsNone(difference(text.replace("' Module comment", "' edited"), text))


class FailClosedTest(unittest.TestCase):
    def test_a_non_ascii_byte_is_refused(self) -> None:
        raw = BASE.replace("Module comment", "Module commént").encode("latin-1")
        with self.assertRaisesRegex(identity.CarryRefused, "non-ASCII byte at offset"):
            identity.code_stream(raw)
        with self.assertRaises(identity.CarryRefused):
            identity.first_difference(BASE.encode("ascii"), raw)


MODULE = "src/core/CORE_Probe.bas"
CLASS = "src/core/CProbe.cls"
FORM = "src/forms/frmProbe.frm"
RESOURCE = "src/forms/frmProbe.frx"
CLASS_TEXT = ("VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\n"
              'Attribute VB_Name = "CProbe"\nOption Explicit\n\' A class\n'
              "Public Value As Double\n")
FORM_TEXT = ('VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProbe\n'
             '   OleObjectBlob   =   "frmProbe.frx":0000\nEnd\n'
             'Attribute VB_Name = "frmProbe"\nOption Explicit\n\' A form\n')


@unittest.skipUnless(shutil.which("git"), "needs Git")
class RepositoryTest(unittest.TestCase):
    root: ClassVar[Path]
    executed: ClassVar[str]

    @classmethod
    def setUpClass(cls) -> None:
        cls.root = Path(tempfile.mkdtemp(prefix="code-identity-"))
        init(cls.root)
        cls.executed = cls.commit({MODULE: BASE, CLASS: CLASS_TEXT, FORM: FORM_TEXT,
                                   RESOURCE: "\x00\x01 binary"}, on_run=False)

    @classmethod
    def tearDownClass(cls) -> None:
        shutil.rmtree(cls.root, ignore_errors=True)

    @classmethod
    def commit(cls, files: dict[str, str | None], on_run: bool = True) -> str:
        """A commit on top of the run commit (the first commit, without ``on_run``)
        that writes each file of ``files``, or deletes it when its text is None."""
        if on_run:
            git(cls.root, "checkout", "-q", "--detach", cls.executed)
        for name, content in files.items():
            path = cls.root / name
            if content is None:
                path.unlink()
                continue
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(content.encode("ascii"))
        git(cls.root, "add", "-A")
        git(cls.root, "commit", "-qm", "Change", "--allow-empty")
        return head(cls.root)

    def refused(self, path: str, run: str, commit: str) -> str | None:
        try:
            identity.carried(self.root, path, run, commit)
        except identity.CarryRefused as error:
            return str(error)
        return None

    def test_comment_only_edits_are_carried(self) -> None:
        target = self.commit({MODULE: SAME["a changed comment"].replace("\n", "\r\n"),
                              CLASS: CLASS_TEXT.replace("' A class", "' Another class"),
                              FORM: FORM_TEXT.replace("' A form", "' Another form")})
        for path in (MODULE, CLASS, FORM):
            with self.subTest(path):
                self.assertIsNone(self.refused(path, self.executed, target))
        self.assertIsNone(self.refused(MODULE, self.executed[:7], target))
        target = self.commit({MODULE: SAME["the letter case of code (HB-01)"]})
        self.assertEqual(identity.carried(self.root, MODULE, self.executed, target),
                         ["Total -> total", "X -> x"])

    def test_a_code_edit_is_refused_and_named(self) -> None:
        target = self.commit({MODULE: text("a changed numeric literal")})
        found = self.refused(MODULE, self.executed, target) or ""
        self.assertIn("code changed: line 10 at the run", found)
        self.assertIn("'Total = X * 3#'", found)

    def test_what_cannot_be_read_is_refused(self) -> None:
        target = self.commit({MODULE: None})
        self.assertIn("cannot be read", self.refused(MODULE, self.executed, target) or "")
        self.assertIn("cannot be read", self.refused(MODULE, "f" * 40, self.executed) or "")
        self.assertIn("no valid commit", self.refused(MODULE, "", self.executed) or "")
        self.assertIn("no valid commit", self.refused(MODULE, self.executed, "") or "")
        self.assertIn("no valid commit", self.refused(MODULE, "--all", self.executed) or "")
        self.assertIn("not a VBA component", self.refused(RESOURCE, self.executed, self.executed) or "")
        self.assertIn("cannot be read", self.refused("src/core/Missing.bas", self.executed,
                                                     self.executed) or "")


if __name__ == "__main__":
    unittest.main()
