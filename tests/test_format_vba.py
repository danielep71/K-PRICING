"""Tests for tools/format_vba.py: layout rules, the tracked modules and the baseline.

Each rule case states one line or a few lines as the formatter must leave
them, so that a failure names the rule. The literal and apostrophe cases are
the ones a line-based formatter is most likely to get wrong. Every rule case
also checks that formatting keeps each logical line, with its continuations
joined, and every token outside white space. The command-line cases run in
temporary Git repositories.
"""

from __future__ import annotations

import contextlib
import io
import json
import re
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import format_vba  # noqa: E402
import vba_baseline  # noqa: E402
import vba_lexer  # noqa: E402
from gitfixture import git, init  # noqa: E402
from test_check_vba_layout import module as headed_module  # noqa: E402

HEAD = ['Attribute VB_Name = "CORE_Probe"', "Option Explicit", "Option Private Module", ""]


def module(*body: str) -> str:
    return "".join(f"{line}\r\n" for line in [*HEAD, *body])


def body_of(text: str) -> list[str]:
    return text.split("\r\n")[len(HEAD):-1]


# A token is a string literal, a comment without its trailing white space, or a
# run of code characters; white space and line breaks separate tokens.
TOKEN = re.compile(r'"(?:[^"\r\n]|"")*"|\'[^\r\n]*|[^\s"\']+')


def tokens(text: str) -> list[str]:
    return [token.rstrip() if token.startswith("'") else token
            for token in TOKEN.findall(text)]


def logical_lines(text: str) -> list[tuple[str, list[str]]]:
    """Each logical line that tools/vba_lexer.py reads, as its statements.

    Continuations, in code and in comments, are joined and comments removed,
    so a blank line added to or removed from a logical line, or comment text
    turned into code, changes the result. Blank logical lines are left out.
    """
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    return [(kind, vba_lexer.split_statements(code))
            for _, _, kind, code in vba_lexer.logical_units(lines) if code.strip()]


class RuleTest(unittest.TestCase):
    def assertFormats(self, before: list[str], after: list[str]) -> None:
        text = module(*before)
        formatted = format_vba.format_text(text)
        self.assertEqual(body_of(formatted), after)
        self.assertEqual(format_vba.format_text(formatted), formatted, "not idempotent")
        self.assertEqual(logical_lines(formatted), logical_lines(text), "logical lines changed")
        self.assertEqual(tokens(formatted), tokens(text), "tokens changed")

    def test_formatted_module_is_unchanged(self) -> None:
        text = module("Private Sub Run()", "    Dim Count As Long", "End Sub")
        self.assertEqual(format_vba.format_text(text), text)

    def test_line_endings_and_end_of_file(self) -> None:
        lf = "\n".join([*HEAD, "Private Sub Run()", "End Sub", "", "", ""])
        self.assertEqual(format_vba.format_text(lf),
                         module("Private Sub Run()", "End Sub"))
        no_final = module("Private Sub Run()", "End Sub").removesuffix("\r\n")
        self.assertEqual(format_vba.format_text(no_final),
                         module("Private Sub Run()", "End Sub"))

    def test_trailing_white_space(self) -> None:
        self.assertFormats(["Private Sub Run()  ", "    Count = 1\t ", "End Sub"],
                           ["Private Sub Run()", "    Count = 1", "End Sub"])

    def test_indentation_tabs(self) -> None:
        self.assertFormats(["Private Sub Run()", "\tIf Ready Then", "\t\tCount = 1",
                            "\tEnd If", "End Sub"],
                           ["Private Sub Run()", "    If Ready Then", "        Count = 1",
                            "    End If", "End Sub"])

    def test_outer_level_lines(self) -> None:
        self.assertFormats(["    Public Function Total() As Double", "Clean_Exit:",
                            "    Exit Function", "  Fail:  ' handler", "    End Function"],
                           ["Public Function Total() As Double", "Clean_Exit:",
                            "    Exit Function", "Fail:  ' handler", "End Function"])
        self.assertFormats(["  Friend Property Get Name() As String", "  End Property"],
                           ["Friend Property Get Name() As String", "End Property"])
        self.assertFormats(["   Private Static Sub Tick()", " End Sub"],
                           ["Private Static Sub Tick()", "End Sub"])

    def test_continuation_lines_keep_their_indentation(self) -> None:
        self.assertFormats(["Public Function Total( _", "    ByVal Amount As Double, _",
                            "    ByVal Rate As Double) As Double", "End Function"],
                           ["Public Function Total( _", "    ByVal Amount As Double, _",
                            "    ByVal Rate As Double) As Double", "End Function"])
        # A continued statement whose next line looks like a label is left alone.
        self.assertFormats(["Private Sub Run()", "    Debug.Print _",
                            "        Done:", "End Sub"],
                           ["Private Sub Run()", "    Debug.Print _",
                            "        Done:", "End Sub"])

    def test_a_continued_comment_is_never_split_or_moved(self) -> None:
        # A comment that ends with " _" continues onto the next line: no blank
        # line may end it early, or its text would become code.
        for end in ("End Sub", "End Function", "End Property"):
            with self.subTest(end=end):
                self.assertFormats([f"{end} ' explanation _", "    Debug.Print 1",
                                    "Private Sub Two()", "End Sub"],
                                   [f"{end} ' explanation _", "    Debug.Print 1", "",
                                    "Private Sub Two()", "End Sub"])
        # Comment text that looks like a label or an End statement stays put,
        # after a space or a tab before the underscore.
        for continued in ("    ' note _", "    ' note\t_", "    Debug.Print 1 ' note _"):
            for text in ("        Done:", "        End Sub"):
                with self.subTest(continued=continued, text=text):
                    lines = ["Private Sub Run()", continued, text, "End Sub"]
                    self.assertFormats(lines, lines)
        # A blank line ends the continuation; it is the one blank line after End.
        self.assertFormats(["End Sub ' a _", "    more", "", "Private Sub Two()", "End Sub"],
                           ["End Sub ' a _", "    more", "", "Private Sub Two()", "End Sub"])

    def test_statement_keywords_are_not_labels(self) -> None:
        # A keyword and a colon separator, such as Else: or Loop:, keeps its
        # indentation; a label still moves to the outer level.
        lines = ["Private Sub Run()", "    If Ready Then", "        Count = 1", "    Else:",
                 "        Count = 2", "    End If", "    Do", "    Loop:", "End Sub"]
        self.assertFormats(lines, lines)
        for word in sorted(format_vba.NOT_LABELS):
            for spelled in (word, word.capitalize(), word.upper()):
                with self.subTest(keyword=spelled):
                    lines = ["Private Sub Run()", f"    {spelled}:", "End Sub"]
                    self.assertFormats(lines, lines)
        self.assertFormats(["Private Sub Run()", "    Elsewhere:", "    Done:", "End Sub"],
                           ["Private Sub Run()", "Elsewhere:", "Done:", "End Sub"])

    # Each row: the input lines and the lines the formatter must leave.
    PROBE_ROWS = [
        ("a continued comment after End",
         ["End Sub ' explanation _", "    Debug.Print 1", "Private Sub Two()"],
         ["End Sub ' explanation _", "    Debug.Print 1", "", "Private Sub Two()"]),
        ("comment continuation text that looks like a label",
         ["    ' note _", "        Done:"], ["    ' note _", "        Done:"]),
        ("a tab before the underscore",
         ["    Debug.Print\t_", "        Done:"], ["    Debug.Print\t_", "        Done:"]),
        ("a blank line that ends a continuation at the end of the file",
         ["End Sub ' a _", ""], ["End Sub ' a _", ""]),
        ("statement keywords with a colon", ["    Else:", "    Loop:"],
         ["    Else:", "    Loop:"]),
        ("indented comment headings", ["'Declare", "    'Invert", "  ' text"],
         ["'Declare", "    'Invert", "  ' text"]),
    ]

    def test_probe_rows(self) -> None:
        for label, before, after in self.PROBE_ROWS:
            with self.subTest(label):
                self.assertFormats(before, after)

    def test_blank_lines_inside_a_logical_line_are_kept(self) -> None:
        # A blank line that a continuation carries over ends that logical line:
        # it is neither removed nor joined to what follows, after an End
        # statement or at the end of the file included.
        unchanged = [
            ["Private Sub One()", "    X = 1 _", "", "    Y = 2", "End Sub"],
            ["Private Sub One()", "    ' a _", "", "End Sub"],
            ["Private Sub One()", "End Sub ' a _", "", "", "Private Sub Two()", "End Sub"],
            ["Private Sub One()", "    X = 1 _", ""],
        ]
        for lines in unchanged:
            with self.subTest(lines=lines):
                self.assertFormats(lines, lines)
        self.assertFormats(["Private Sub One()", "End Sub ' a _", "", "", "Private Sub Two()",
                            "End Sub"],
                           ["Private Sub One()", "End Sub ' a _", "", "", "Private Sub Two()",
                            "End Sub"])
        self.assertFormats(["Private Sub One()", "End Sub ' a _", "", "", ""],
                           ["Private Sub One()", "End Sub ' a _", ""])
        self.assertFormats(["Private Const A As Long = 1 _", "    + 1", "", "  ", ""],
                           ["Private Const A As Long = 1 _", "    + 1"])

    def test_one_blank_line_between_procedures(self) -> None:
        self.assertFormats(["Private Sub One()", "End Sub", "'Header of Two",
                            "Private Sub Two()", "End Sub"],
                           ["Private Sub One()", "End Sub", "", "'Header of Two",
                            "Private Sub Two()", "End Sub"])
        self.assertFormats(["Private Sub One()", "End Sub", "", "", "", "",
                            "Private Sub Two()", "End Sub"],
                           ["Private Sub One()", "End Sub", "",
                            "Private Sub Two()", "End Sub"])

    def test_other_lines_are_not_moved(self) -> None:
        lines = ["Private Sub Run()", "    If Ready Then Exit Sub", "    Select Case Mode",
                 "        Case Else: Count = 0", "    End Select", "    End",
                 "    Debug.Print Count: Count = 1", "End Sub"]
        self.assertFormats(lines, lines)


class LiteralTest(unittest.TestCase):
    """No rule changes a string literal, or mistakes one for code."""

    CASES = [
        ("a label name in a literal is data", '    Message = "Fail:"'),
        ("a procedure keyword in a literal is data", '    Message = "End Sub"'),
        ("an apostrophe in a literal does not start a comment",
         "    Const Quote As String = \"a 'b' c\""),
        ("a doubled quote does not unbalance the scan",
         "    Const Said As String = \"he said \"\"go\"\" 'x' now\""),
        ("a tab inside a literal is kept", '    Message = "a\tb"'),
        ("spaces inside a literal are kept", '    Message = "a     b"'),
    ]

    def test_literals_are_unchanged(self) -> None:
        for label, line in self.CASES:
            with self.subTest(label):
                text = module("Private Sub Run()", line, "End Sub")
                self.assertIn(line, body_of(format_vba.format_text(text)))

    def test_comment_after_a_literal_with_an_apostrophe(self) -> None:
        line = "Done:  ' after \"a 'quoted' word\""
        self.assertEqual(format_vba.split_code_comment(line),
                         ("Done:  ", "' after \"a 'quoted' word\""))
        self.assertEqual(format_vba.split_code_comment("    X = \"it's\" ' note"),
                         ('    X = "it\'s" ', "' note"))


class HeaderTest(unittest.TestCase):
    def test_vbe_header_is_kept_byte_for_byte(self) -> None:
        header = ["VERSION 5.00", "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRun ",
                  "   Caption         =   \"Run\"   ", "\tClientHeight    =   3015",
                  "   Begin VB.CommandButton cmdOk ", "      Caption = \"OK\" ", "   End",
                  "End", "Attribute VB_Name = \"frmRun\"", "Attribute VB_Exposed = False"]
        text = "".join(f"{line}\r\n" for line in [*header, "Option Explicit", "",
                                                    "  Private Sub cmdOk_Click()  ",
                                                    "  End Sub"])
        lines = format_vba.format_text(text).split("\r\n")
        self.assertEqual(lines[:len(header)], header)
        self.assertEqual(lines[len(header):-1],
                         ["Option Explicit", "", "Private Sub cmdOk_Click()", "End Sub"])

    def test_class_header(self) -> None:
        header = ["VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "END",
                  'Attribute VB_Name = "CResult"', "Attribute VB_PredeclaredId = False"]
        lines = [*header, "Option Explicit"]
        self.assertEqual(format_vba.format_body([]), [])
        self.assertEqual(format_vba.header_length(lines), len(header))


def deform(text: str) -> str:
    """Module text with every kind of drift that the formatter must undo.

    The VBE header is kept as it stands, as the formatter keeps it, and a line
    that a continuation carries over is never read as an End statement.
    """
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if lines and lines[-1] == "":
        lines.pop()
    split = format_vba.header_length(lines)
    out: list[str] = lines[:split]
    continued = False
    for line in lines[split:]:
        carried = continued
        code = format_vba.split_code_comment(line.strip())[0].strip()
        if format_vba.outer_level(line) and not carried:
            line = "  " + line                          # outer-level line indented
        elif line.startswith("    "):
            line = "\t" + line[4:]                      # tab indentation
        if code and not code.startswith("Attribute"):
            line += " \t"                               # trailing white space
        out.append(line)
        continued = format_vba.continues(line.rstrip())
        if not carried and not continued and format_vba.PROCEDURE_END.match(code):
            out.append("")                              # an extra blank line after End
    return "\n".join(out) + "\n\n\n"                    # LF endings, blank lines at the end


def baselined(root: Path) -> set[str]:
    """The modules the root's default layout baseline lists; none without the file."""
    path = root / vba_baseline.BASELINE
    return set(vba_baseline.load_baseline(path, vba_baseline.BASELINE)) if path.exists() else set()


class DeformTest(unittest.TestCase):
    """The drift helper on module shapes the project may track."""

    CLASS = ("VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
             'Attribute VB_Name = "CResult"\r\nOption Explicit\r\n\r\n'
             "Public Sub Run()\r\n    X = 1\r\nEnd Sub\r\n")
    STANDARD = 'Attribute VB_Name = "A"\r\nOption Explicit\r\n\r\nSub Run()\r\n    X = 1\r\nEnd Sub\r\n'
    CONTINUED_COMMENT = ('Attribute VB_Name = "A"\r\nSub Run()\r\n    X = 1 \' note _\r\n'
                         "End Sub\r\nEnd Sub\r\n")

    def assert_undone(self, text: str) -> None:
        drifted = deform(text)
        self.assertNotEqual(drifted, text)
        formatted = format_vba.format_text(drifted)
        self.assertEqual(formatted, format_vba.format_text(text))
        self.assertEqual(tokens(formatted), tokens(text))

    def test_a_class_header_is_kept(self) -> None:
        self.assertEqual(format_vba.format_text(self.CLASS), self.CLASS)
        self.assert_undone(self.CLASS)
        self.assertTrue(deform(self.CLASS).startswith("VERSION 1.0 CLASS\nBEGIN\n  MultiUse"))

    def test_lf_endings_and_a_missing_final_line_break(self) -> None:
        for text in (self.STANDARD.replace("\r\n", "\n"), self.STANDARD[:-2]):
            with self.subTest(repr(text[-12:])):
                self.assert_undone(text)

    def test_a_continued_comment_before_end_is_not_an_end(self) -> None:
        self.assert_undone(self.CONTINUED_COMMENT)


class TrackedModuleTest(unittest.TestCase):
    """The formatter on the text of this repository's tracked modules, read only."""

    modules: dict[str, str]
    baseline: set[str]

    @classmethod
    def setUpClass(cls) -> None:
        cls.modules = {name: format_vba.read(ROOT / name)
                       for name in format_vba.tracked_modules(ROOT)}
        cls.baseline = baselined(ROOT)

    def test_modules_are_in_layout(self) -> None:
        # A baselined module may stay out of layout; the CLI's ratchet, run on this
        # repository by MainTest.test_repository_passes, fails it once it is in layout.
        for name, text in self.modules.items():
            if name in self.baseline:
                continue
            with self.subTest(name):
                self.assertEqual(format_vba.format_text(text), text)

    def test_drift_is_undone_without_changing_code(self) -> None:
        for name, text in self.modules.items():
            with self.subTest(name):
                drifted = deform(text)
                self.assertNotEqual(drifted, text)
                formatted = format_vba.format_text(drifted)
                self.assertEqual(formatted, format_vba.format_text(text))
                self.assertEqual(tokens(formatted), tokens(text))
                self.assertEqual(logical_lines(formatted), logical_lines(text))
                self.assertEqual(format_vba.format_text(formatted), formatted)

    def test_token_check_detects_a_change(self) -> None:
        self.assertNotEqual(tokens('X = "a  b"'), tokens('X = "a b"'))
        self.assertNotEqual(tokens("X = a b"), tokens("X = ab"))
        self.assertEqual(tokens("  X = 1 ' note  \r\n"), tokens("X  =  1\t' note"))
        self.assertNotEqual(logical_lines("End Sub ' a _\r\n    Debug.Print 1\r\n"),
                            logical_lines("End Sub ' a _\r\n\r\n    Debug.Print 1\r\n"))


class MainTest(unittest.TestCase):
    """The command line, in temporary Git repositories."""

    GOOD = module("Private Sub Run()", "End Sub")
    BAD = module("  Private Sub Run()", "End Sub")

    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="format-vba-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        init(self.root)

    def track(self, files: dict[str, str], *, add: bool = True) -> None:
        for name, text in files.items():
            (self.root / name).parent.mkdir(parents=True, exist_ok=True)
            (self.root / name).write_bytes(text.encode("latin-1"))
        if add:
            git(self.root, "add", "--all")

    def baseline(self, files: object, **fields: object) -> None:
        path = self.root / vba_baseline.BASELINE
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps({"schema_version": 1, "files": files, **fields}),
                        encoding="utf-8")

    def run_main(self, *args: str) -> tuple[int, str, str]:
        with contextlib.redirect_stdout(io.StringIO()) as printed, \
                contextlib.redirect_stderr(io.StringIO()) as errors:
            code = format_vba.main([*args, "--root", str(self.root)])
        return code, printed.getvalue(), errors.getvalue()

    def test_no_source_passes(self) -> None:
        code, printed, _ = self.run_main("--check")
        self.assertEqual(code, 0)
        self.assertIn("vba-format: PASS (no VBA source found)", printed)

    def test_repository_passes(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()) as printed:
            self.assertEqual(format_vba.main(["--check", "--root", str(ROOT)]), 0)
        self.assertIn("vba-format: PASS (", printed.getvalue())

    def test_check_and_write(self) -> None:
        files = {"src/core/CORE_Probe.bas": self.BAD, "src/workbook/frmRun.frx": "\x00\x01",
                 "docs/notes.bas": self.BAD}
        self.track(files)
        code, printed, _ = self.run_main("--check")
        self.assertEqual(code, 1)
        self.assertIn("FAIL: src/core/CORE_Probe.bas: not in house-style layout", printed)
        self.assertNotIn("docs/notes.bas", printed)
        self.assertIn("vba-format: FAIL (1 file(s), 0 baselined", printed)
        code, printed, _ = self.run_main("--write")
        self.assertEqual(code, 0)
        self.assertIn("formatted: src/core/CORE_Probe.bas", printed)
        self.assertIn("Re-import each rewritten module into the VBE", printed)
        self.assertEqual((self.root / "src/core/CORE_Probe.bas").read_bytes().decode("latin-1"),
                         self.GOOD)
        self.assertEqual((self.root / "docs/notes.bas").read_bytes().decode("latin-1"),
                         self.BAD)

    def test_only_tracked_modules_are_read(self) -> None:
        self.track({"src/A.bas": self.GOOD, "tests/modules/B.cls": self.GOOD,
                    "examples/modules/C.frm": self.GOOD, "notes/D.bas": self.BAD})
        self.track({"src/Untracked.bas": self.BAD}, add=False)
        code, printed, _ = self.run_main("--check")
        self.assertEqual(code, 0, printed)
        self.assertIn("(3 file(s), 0 baselined", printed)
        self.assertNotIn("Untracked", printed)

    def test_not_a_git_repository(self) -> None:
        with tempfile.TemporaryDirectory(prefix="format-vba-plain-") as plain:
            with contextlib.redirect_stdout(io.StringIO()), \
                    contextlib.redirect_stderr(io.StringIO()) as errors:
                code = format_vba.main(["--check", "--root", plain])
        self.assertEqual(code, 2)
        self.assertIn("vba-format: ERROR: cannot list the tracked files", errors.getvalue())

    def test_baselined_module_out_of_layout_passes(self) -> None:
        self.track({"src/A.bas": self.BAD, "src/B.bas": self.GOOD})
        self.baseline(["src/A.bas"])
        code, printed, _ = self.run_main("--check")
        self.assertEqual(code, 0, printed)
        self.assertIn("baselined: src/A.bas: not in house-style layout, listed in "
                      "tools/vba_layout_baseline.json", printed)
        self.assertIn("vba-format: PASS (2 file(s), 1 baselined", printed)

    def test_tracked_module_test_reads_the_baseline(self) -> None:
        self.assertEqual(baselined(self.root), set())
        self.track({"src/A.bas": self.BAD, "src/B.bas": self.GOOD})
        self.baseline(["src/A.bas"])
        self.assertEqual(baselined(self.root), {"src/A.bas"})
        self.assertEqual(self.run_main("--check")[0], 0)

    def test_unlisted_module_out_of_layout_fails(self) -> None:
        self.track({"src/A.bas": self.BAD, "src/B.bas": self.BAD})
        self.baseline(["src/A.bas"])
        code, printed, _ = self.run_main("--check")
        self.assertEqual(code, 1)
        self.assertIn("FAIL: src/B.bas: not in house-style layout; run python "
                      "tools/format_vba.py --write src/B.bas", printed)
        self.assertNotIn("FAIL: src/A.bas", printed)

    def test_listed_module_without_headers_stays_listed_after_write(self) -> None:
        # In statement layout but without banners and headers: the entry stays,
        # because tools/check_vba_layout.py still needs it.
        self.track({"src/A.bas": self.BAD, "src/B.bas": self.GOOD})
        self.baseline(["src/A.bas"])
        code, printed, _ = self.run_main("--write", "src/A.bas")
        self.assertEqual(code, 0)
        self.assertIn("src/A.bas stays in tools/vba_layout_baseline.json until its banners "
                      "and headers pass python tools/check_vba_layout.py", printed)
        self.assertNotIn("Remove src/A.bas", printed)
        code, printed, _ = self.run_main("--check")
        self.assertEqual(code, 0, printed)
        self.assertIn("baselined: src/A.bas", printed)

    def test_listed_module_in_layout_fails_until_removed(self) -> None:
        bad = headed_module().replace("\n", "\r\n").replace("Public Sub Run()",
                                                             "  Public Sub Run()")
        good = headed_module().replace("\n", "\r\n")
        self.track({"src/A.bas": bad, "src/B.bas": good})
        self.baseline(["src/A.bas"])
        code, printed, _ = self.run_main("--write", "src/A.bas")
        self.assertEqual(code, 0)
        self.assertIn("formatted: src/A.bas", printed)
        self.assertIn("Re-import each rewritten module into the VBE", printed)
        self.assertIn("Remove src/A.bas from tools/vba_layout_baseline.json", printed)
        code, printed, _ = self.run_main("--check")
        self.assertEqual(code, 1)
        self.assertIn("FAIL: src/A.bas: in layout but still listed in "
                      "tools/vba_layout_baseline.json; remove its entry", printed)
        self.baseline([])
        self.assertEqual(self.run_main("--check")[0], 0)
        (self.root / vba_baseline.BASELINE).unlink()
        self.assertEqual(self.run_main("--check")[0], 0)

    def test_listed_path_that_is_not_a_tracked_module_fails(self) -> None:
        self.track({"src/A.bas": self.GOOD, "docs/C.bas": self.BAD, "src/D.txt": self.BAD})
        self.track({"src/Untracked.bas": self.BAD}, add=False)
        for name in ("src/Gone.bas", "src/Untracked.bas", "docs/C.bas", "src/D.txt"):
            with self.subTest(name):
                self.baseline([name])
                code, printed, _ = self.run_main("--check")
                self.assertEqual(code, 1)
                self.assertIn(f"FAIL: {name}: listed in tools/vba_layout_baseline.json but not "
                              "a tracked VBA module", printed)

    def test_no_baseline_checks_every_module(self) -> None:
        self.track({"src/A.bas": self.BAD})
        self.baseline(["src/A.bas"])
        code, printed, _ = self.run_main("--check", "--no-baseline")
        self.assertEqual(code, 1)
        self.assertIn("FAIL: src/A.bas: not in house-style layout", printed)
        self.assertIn("0 baselined", printed)
        # --no-baseline ignores a baseline file, even a malformed one.
        (self.root / vba_baseline.BASELINE).write_text("{", encoding="utf-8")
        self.assertEqual(self.run_main("--check", "--no-baseline")[0], 1)

    def test_explicit_baseline(self) -> None:
        self.track({"src/A.bas": self.BAD})
        (self.root / "layout.json").write_text(
            json.dumps({"schema_version": 1, "files": ["src/A.bas"]}), encoding="utf-8")
        code, printed, _ = self.run_main("--check", "--baseline", "layout.json")
        self.assertEqual(code, 0, printed)
        self.assertIn("baselined: src/A.bas: not in house-style layout, listed in layout.json",
                      printed)
        code, _, errors = self.run_main("--check", "--baseline", "missing.json")
        self.assertEqual(code, 2)
        self.assertIn("vba-format: ERROR: cannot read baseline missing.json", errors)

    def test_baseline_errors(self) -> None:
        self.track({"src/A.bas": self.BAD, "src/B.bas": self.GOOD})
        valid = {"schema_version": 1, "files": ["src/A.bas"]}
        cases: dict[str, object] = {
            "not JSON": "{",
            "not an object": json.dumps(["src/A.bas"]),
            "wrong schema_version": json.dumps({**valid, "schema_version": 2}),
            "no schema_version": json.dumps({"files": ["src/A.bas"]}),
            "files not a list": json.dumps({**valid, "files": "src/A.bas"}),
            "a path not a string": json.dumps({**valid, "files": [1]}),
            "unsorted": json.dumps({**valid, "files": ["src/B.bas", "src/A.bas"]}),
            "duplicated": json.dumps({**valid, "files": ["src/A.bas", "src/A.bas"]}),
            "absolute": json.dumps({**valid, "files": ["/src/A.bas"]}),
            "drive letter": json.dumps({**valid, "files": ["C:/src/A.bas"]}),
            "backslash": json.dumps({**valid, "files": ["src\\A.bas"]}),
            "parent": json.dumps({**valid, "files": ["src/../src/A.bas"]}),
            "dot": json.dumps({**valid, "files": ["./src/A.bas"]}),
            "empty": json.dumps({**valid, "files": [""]}),
        }
        path = self.root / vba_baseline.BASELINE
        path.parent.mkdir(parents=True, exist_ok=True)
        for label, content in cases.items():
            for mode in ("--check", "--write"):
                with self.subTest(label, mode=mode):
                    path.write_text(str(content), encoding="utf-8")
                    code, _, errors = self.run_main(mode)
                    self.assertEqual(code, 2)
                    self.assertIn("vba-format: ERROR: ", errors)
                    self.assertIn("tools/vba_layout_baseline.json", errors)
        # An error never rewrites a module.
        self.assertEqual((self.root / "src/A.bas").read_bytes().decode("latin-1"), self.BAD)

    def test_bytes_outside_ascii_round_trip(self) -> None:
        text = module("Private Sub Run()", "    ' caf\xe9", "End Sub")
        self.track({"src/core/CORE_Probe.bas": text})
        code, _, _ = self.run_main("--write")
        self.assertEqual(code, 0)
        self.assertEqual((self.root / "src/core/CORE_Probe.bas").read_bytes(),
                         text.encode("latin-1"))

    def test_paths_outside_the_root(self) -> None:
        # Every file is processed and reported, by its absolute path.
        with tempfile.TemporaryDirectory(prefix="format-vba-outside-") as outside:
            files = [Path(outside) / name for name in ("A.bas", "B.bas")]
            for path in files:
                path.write_bytes(self.BAD.encode("latin-1"))
            for mode in ("--check", "--write"):
                with self.subTest(mode=mode):
                    code, printed, _ = self.run_main(mode, *map(str, files))
                    for path in files:
                        self.assertIn(path.resolve().as_posix(), printed)
            self.assertEqual(code, 0)
            for path in files:
                self.assertEqual(path.read_bytes().decode("latin-1"), self.GOOD)

    def test_folder_paths(self) -> None:
        self.track({"elsewhere/a/CORE_Probe.bas": self.BAD,
                    "elsewhere/b/notes.txt": "  not VBA\r\n"}, add=False)
        code, printed, _ = self.run_main("--write", "elsewhere")
        self.assertEqual(code, 0, printed)
        self.assertIn("formatted: elsewhere/a/CORE_Probe.bas", printed)
        self.assertEqual((self.root / "elsewhere/b/notes.txt").read_bytes(), b"  not VBA\r\n")
        code, _, errors = self.run_main("--check", "missing")
        self.assertEqual(code, 2)
        self.assertIn("missing is not a file or folder", errors)

    def test_explicit_paths(self) -> None:
        self.track({"src/A.bas": self.GOOD})
        self.track({"elsewhere/CORE_Probe.bas": self.BAD}, add=False)
        code, printed, _ = self.run_main("--check", "elsewhere/CORE_Probe.bas")
        self.assertEqual(code, 1)
        self.assertIn("FAIL: elsewhere/CORE_Probe.bas", printed)
        self.assertIn("(1 file(s)", printed)
        # A named path is checked against the baseline; entries outside the
        # named paths are not reported.
        self.baseline(["elsewhere/CORE_Probe.bas", "src/Gone.bas"])
        code, printed, _ = self.run_main("--check", "elsewhere/CORE_Probe.bas")
        self.assertEqual(code, 0, printed)
        self.assertIn("baselined: elsewhere/CORE_Probe.bas", printed)
        self.assertEqual(self.run_main("--check", "src/A.bas")[0], 0)


if __name__ == "__main__":
    unittest.main()
