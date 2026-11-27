"""Valid and invalid fixtures for the layout check (tools/check_vba_layout.py).

Each fixture is a throw-away Git repository, built as in test_check_vba.py,
holding one synthetic VBA module. The valid module follows the layout of
docs/VBA_HOUSE_STYLE.md; each invalid fixture changes one thing that the
check must report.
"""

from __future__ import annotations

import contextlib
import io
import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import check_vba_layout  # noqa: E402
import vba_baseline  # noqa: E402
import vba_body_layout  # noqa: E402
import vba_header_layout  # noqa: E402
from test_check_vba import repository  # noqa: E402

PATH = "tests/modules/TEST_Fixture.bas"
BANNER = "'" + "-" * 78
FIELDS = ("PURPOSE", "WHY THIS EXISTS", "INPUTS", "RETURNS", "BEHAVIOR", "ERROR POLICY",
          "DEPENDENCIES", "NOTES", "UPDATED")


def header(kind: str, fields: tuple[str, ...] = FIELDS, banner: str = BANNER) -> str:
    """A component or procedure header: a title line, then each field with text."""
    title = f"' {kind}" if ":" in kind else "'" + " " * ((78 - len(kind)) // 2) + kind
    lines = ["'", banner, title, banner]
    for field in fields:
        text = ("2026-11-21" if field == "UPDATED"
                else "Nothing" if field == "RETURNS" and ":" not in kind
                else "None at module level" if field == "INPUTS" and ":" in kind else "None")
        lines += [f"' {field}", f"'   {text}", "'"]
    lines[-1] = banner
    return "\n".join(lines)


def module(component: str | None = None, procedure: str | None = None) -> str:
    """A valid module, or one with its component or procedure header replaced."""
    component = header("MODULE: TEST_Fixture") if component is None else component
    procedure = header("RUN") if procedure is None else procedure
    return (f'Attribute VB_Name = "TEST_Fixture"\n\nOption Explicit\nOption Private Module\n\n'
            f"{component}\n\n{BANNER}\n' PRIVATE CONSTANTS\n{BANNER}\n"
            "    Private Const LIMIT As Long = 1 'Largest value\n\n"
            f"Public Sub Run()\n\n{procedure}\n\n{section('DECLARE')}\n"
            '    Const PROC_NAME     As String = "TEST_Fixture.Run"\n'
            "    Dim Count           As Long\n\nEnd Sub\n")


def section(title: str) -> str:
    """A section banner inside a procedure."""
    return f"{BANNER}\n' {title}\n{BANNER}"


def findings(text: str) -> list[str]:
    found, _ = vba_header_layout.check_text(PATH, text)
    return [str(finding) for finding in found]


class LayoutTest(unittest.TestCase):
    def test_valid_module_passes(self) -> None:
        self.assertEqual(findings(module()), [])

    def test_a_signature_on_several_lines_is_read(self) -> None:
        text = module().replace("Public Sub Run()\n", "Public Sub Run( _\n    ByVal A As Long, _\n"
                                "    ByVal B As Long)\n")
        self.assertEqual(findings(text), [])
        broken = text.replace(header("RUN"), "' just a comment")
        self.assertTrue(any("procedure Run" in line for line in findings(broken)))

    def test_procedure_attribute_lines_come_before_the_header(self) -> None:
        text = module().replace("Public Sub Run()\n", "Public Sub Run()\n"
                                'Attribute Run.VB_Description = "Runs"\n'
                                "Attribute Run.VB_UserMemId = 0\n")
        self.assertEqual(findings(text), [])
        broken = text.replace(header("RUN"), "' just a comment")
        self.assertTrue(any("procedure Run" in line for line in findings(broken)))

    def test_a_declaration_in_a_continued_comment_is_not_a_procedure(self) -> None:
        text = module().replace("    Dim Count As Long\n",
                                "    ' Example declaration _\nPublic Sub Example()\n"
                                "    Dim Count As Long\n")
        self.assertEqual(findings(text), [])
        found, procedures = vba_header_layout.check_text(PATH, text)
        self.assertEqual(procedures, 1)

    def test_equals_banner_is_allowed_for_a_component_header(self) -> None:
        equals = "'" + "=" * 78
        self.assertEqual(findings(module(header("MODULE: TEST_Fixture", banner=equals))), [])

    def test_invalid_fixtures_fail(self) -> None:
        short = "'" + "-" * 77
        cases = {
            "a short banner": (module(header("MODULE: TEST_Fixture", banner=short)),
                               "banner is 78 characters, not 79"),
            "no component header": (module("' only a note"), "component header missing"),
            "a wrong component name": (module(header("MODULE: Other")),
                                       "expected MODULE: TEST_Fixture"),
            "a missing field": (module(header("MODULE: TEST_Fixture",
                                              tuple(f for f in FIELDS if f != "NOTES"))),
                                "component: header missing NOTES"),
            "fields out of order": (module(header("MODULE: TEST_Fixture",
                                                  (*FIELDS[1::-1], *FIELDS[2:]))),
                                    "out of order or repeated"),
            "an unknown field": (module(header("MODULE: TEST_Fixture").replace(
                "' NOTES", "' REFERENCE")), "unknown header field REFERENCE"),
            "a procedure without a header": (module(procedure="' note"),
                                             "procedure Run: header missing"),
            "a procedure with a missing field": (
                module(procedure=header("RUN", tuple(f for f in FIELDS if f != "RETURNS"))),
                "procedure Run: header missing RETURNS"),
            "a field without text": (
                module(procedure=header("RUN").replace("' ERROR POLICY\n'   None",
                                                       "' ERROR POLICY\n'")),
                "ERROR POLICY has no text"),
            "an UPDATED that is not a date": (
                module(procedure=header("RUN").replace("2026-11-21", "yesterday")),
                "UPDATED is not one YYYY-MM-DD date"),
        }
        for label, (text, expected) in cases.items():
            with self.subTest(label):
                self.assertTrue(any(expected in line for line in findings(text)),
                                findings(text))

    def test_a_form_header_written_by_the_vbe_is_exempt(self) -> None:
        text = ("VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRun \n"
                '   Caption         =   "Run"\nEnd\nAttribute VB_Name = "frmRun"\n'
                "Attribute VB_Exposed = False\n\nOption Explicit\n\n"
                f"{header('MODULE: frmRun')}\n")
        self.assertEqual(findings(text), [])

    def test_main_exits_0_when_valid_and_1_when_invalid(self) -> None:
        for files, expected in (({PATH: module()}, 0),
                                ({PATH: module("' only a note")}, 1)):
            with repository(files) as root, contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(check_vba_layout.main(["--root", str(root)]), expected)
                verdict = output.getvalue().splitlines()[-1]
                self.assertIn("PASS" if expected == 0 else "FAIL", verdict)
                self.assertIn("static check, nothing compiled or run", verdict)


def procedure_header(title: str, inputs: tuple[str, ...] = (), returns: str = "Nothing") -> str:
    """A procedure header whose INPUTS name ``inputs`` and whose RETURNS is ``returns``."""
    text = header(title)
    named = "\n".join(f"'   {name}\n'     The {name}." for name in inputs) or "'   None"
    text = text.replace("' INPUTS\n'   None", f"' INPUTS\n{named}")
    return text.replace("' RETURNS\n'   Nothing", f"' RETURNS\n'   {returns}")


FULL = "\n".join([
    'Attribute VB_Name = "TEST_Fixture"', "", "Option Explicit", "", header("MODULE: TEST_Fixture"),
    "", "Private Const LIMIT As Long = 1", "", "", "'", BANNER, "'", "'" + " " * 34 + "PROCEDURES",
    "'", BANNER, "'", "", "Public Function Count(ByVal Text As String) As Long",
    "", procedure_header("COUNT", ("Text",), "The count."), "", section("DECLARE"),
    '    Const PROC_NAME     As String = "TEST_Fixture.Count"   \'Name in errors',
    "    Dim Found           As Long                            'Matches so far", "",
    section("INITIALIZE"), "    'Trap errors", "        On Error GoTo Failed", "",
    section("COUNT"), "    'Count the matches",
    '        Found = InStr(1, Text, "a", vbBinaryCompare)', "        Select Case Found",
    "            Case 0", "                Count = 0", "            Case Else",
    "                Count = 1", "        End Select", "", section("EXIT PROCEDURE"),
    "    'Leave", "        Exit Function", "", section("ERROR HANDLER"), "Failed:",
    "    'Report the failure", "        Count = -1", "", "End Function", ""])


def body(text: str) -> list[str]:
    return [str(finding) for finding in vba_body_layout.check_text(PATH, text)]


class BodyTest(unittest.TestCase):
    """The body and statement rules (tools/vba_body_layout.py), one change at a time."""

    def test_the_valid_procedure_passes(self) -> None:
        self.assertEqual(body(FULL), [])
        self.assertEqual(body(module()), [])

    def test_a_compact_private_procedure_skips_the_section_rules(self) -> None:
        compact = FULL.replace("Public Function", "Private Function").replace(
            section("INITIALIZE"), section("SET UP"))
        self.assertEqual(body(compact), [])
        self.assertTrue(body(FULL.replace(section("INITIALIZE"), section("SET UP"))))

    def test_string_contents_do_not_count_towards_the_line_length(self) -> None:
        text = FULL.replace('"a"', '"' + "a" * 150 + '"')
        self.assertEqual(body(text), [])

    def test_invalid_fixtures_fail(self) -> None:
        def drop(old: str, count: int = 1) -> str:
            return FULL.replace(old, "", count)

        title = "'" + " " * 36 + "COUNT"
        multi = FULL.replace("Public Function Count(ByVal Text As String) As Long\n",
                             "Public Function Count( _\n    ByVal Text As String) As Long\n")
        cases = {
            "a title not centred": (FULL.replace(title, "' COUNT"),
                                    "title 'COUNT' is not centred"),
            "no blank after a one-line signature": (
                FULL.replace("As Long\n\n'\n", "As Long\n'\n", 1),
                "0 blank line(s) between the signature and its header, expected 1"),
            "a blank after a multi-line signature": (
                multi, "1 blank line(s) between the signature and its header, expected 0"),
            "no blank before a section": (
                FULL.replace("'Matches so far\n\n", "'Matches so far\n"),
                "0 blank line(s) before the section INITIALIZE, expected 1"),
            "a blank after a section banner": (
                FULL.replace(section("COUNT") + "\n", section("COUNT") + "\n\n"),
                "a blank line after the banner of the section COUNT"),
            "DECLARE without PROC_NAME": (
                drop('    Const PROC_NAME     As String = "TEST_Fixture.Count"   \'Name in errors\n'),
                "DECLARE does not start with Const PROC_NAME"),
            "a wrong PROC_NAME": (FULL.replace('"TEST_Fixture.Count"', '"TEST_Fixture.Run"'),
                                  'PROC_NAME is "TEST_Fixture.Run"'),
            "no DECLARE section": (
                module().replace(section("DECLARE") + "\n" + '    Const PROC_NAME     As String = '
                                 '"TEST_Fixture.Run"\n', ""),
                "no DECLARE section for its local variables"),
            "sections out of order": (
                FULL.replace(section("EXIT PROCEDURE"), section("SWAP")).replace(
                    section("COUNT"), section("EXIT PROCEDURE")).replace(
                    section("SWAP"), section("COUNT")),
                "section COUNT is out of order"),
            "the handler outside INITIALIZE": (
                FULL.replace("    'Trap errors\n        On Error GoTo Failed\n", "    'Start\n"
                             "        Found = 0\n").replace(
                    "    'Count the matches\n", "    'Count the matches\n"
                    "        On Error GoTo Failed\n"),
                "On Error GoTo Failed is not in INITIALIZE"),
            "no Exit before the label": (FULL.replace("        Exit Function\n",
                                                      "        Count = Count\n"),
                                         "no Exit Function in EXIT PROCEDURE before Failed:"),
            "a misaligned As": (FULL.replace("Dim Found           As", "Dim Found  As"),
                                "As in column 16, not in column 25"),
            "a misaligned trailing comment": (
                FULL.replace("As Long                            'Matches",
                             "As Long  'Matches"),
                "trailing comment in column 34, not in column 60"),
            "a Hungarian local name": (FULL.replace("Found", "lngFound"),
                                       "local variable lngFound is not a descriptive"),
            "INPUTS that do not match": (FULL.replace("'   Text\n", "'   Value\n"),
                                         "INPUTS lists Value; the signature has Text"),
            "component INPUTS": (FULL.replace("None at module level", "None"),
                                 "component INPUTS is not 'None at module level'"),
            "RETURNS of a Sub": (module().replace("'   Nothing", "'   A count."),
                                 "RETURNS of a Sub is not Nothing"),
            "a line over 120": (FULL.replace("        Count = -1",
                                             "        Count = -1 + " + " + ".join(["1"] * 40)),
                                "over 120"),
            "InStr without a compare mode": (FULL.replace(', vbBinaryCompare)', ')'),
                                             "InStr without an explicit compare mode"),
            "Select Case without Case Else": (
                FULL.replace("            Case Else\n", "            Case 1\n"),
                "Select Case without Case Else"),
            "no blank before End": (FULL.replace("Count = -1\n\nEnd", "Count = -1\nEnd"),
                                    "0 blank line(s) before End Function, expected 1"),
            "one blank after the declarations": (
                FULL.replace("= 1\n\n\n'", "= 1\n\n'"),
                "group banner: 1 blank line(s) after the declarations, expected 2"),
            "no blank after a group banner": (
                FULL.replace("PROCEDURES\n'\n" + BANNER + "\n'\n\n", "PROCEDURES\n'\n" + BANNER
                             + "\n'\n"),
                "0 blank line(s) between a group banner and the signature"),
        }
        for label, (text, expected) in cases.items():
            with self.subTest(label):
                self.assertNotEqual(text, FULL if label not in ("no DECLARE section",
                                                                "RETURNS of a Sub") else module())
                self.assertTrue(any(expected in line for line in body(text)), body(text))

    def test_a_group_banner_after_an_end_wants_one_blank_line(self) -> None:
        second = FULL.replace('Attribute VB_Name = "TEST_Fixture"\n\nOption Explicit\n\n'
                              + header("MODULE: TEST_Fixture") + "\n\nPrivate Const LIMIT As Long = 1"
                              "\n\n\n", "")
        text = FULL + "\n" + second.replace("Count", "Other").replace("COUNT", "OTHER")
        self.assertEqual(body(text), [])
        self.assertIn("group banner: 2 blank line(s) after the End above it, expected 1",
                      " ".join(body(FULL + "\n\n" + second.replace("Count", "Other").replace(
                          "COUNT", "OTHER"))))


OTHER = "tests/modules/TEST_Other.bas"
BASELINE_NAME = vba_baseline.BASELINE


def baseline_file(files: object, **fields: object) -> str:
    return json.dumps({"schema_version": 1, "files": files, **fields})


class BaselineTest(unittest.TestCase):
    """The baseline that the layout check shares with tools/format_vba.py."""

    GOOD = module()
    BAD = module("' only a note")

    def run_check(self, files: dict[str, str], *args: str) -> tuple[int, str, str]:
        with repository(files) as root, contextlib.redirect_stdout(io.StringIO()) as out, \
                contextlib.redirect_stderr(io.StringIO()) as errors:
            code = check_vba_layout.main(["--root", str(root), *args])
        return code, out.getvalue(), errors.getvalue()

    def test_a_listed_module_out_of_layout_is_baselined(self) -> None:
        code, printed, _ = self.run_check({PATH: self.BAD, OTHER: self.GOOD,
                                           BASELINE_NAME: baseline_file([PATH])})
        self.assertEqual(code, 0, printed)
        self.assertIn(f"baselined: {PATH}: not in house-style layout, listed in "
                      f"{BASELINE_NAME}", printed)
        self.assertIn("vba-layout: PASS (2 VBA component(s), 1 baselined,", printed)

    def test_an_unlisted_module_out_of_layout_fails(self) -> None:
        code, printed, _ = self.run_check({PATH: self.BAD, OTHER: self.BAD,
                                           BASELINE_NAME: baseline_file([PATH])})
        self.assertEqual(code, 1)
        self.assertIn(f"FAIL: {OTHER}:", printed)
        self.assertNotIn(f"FAIL: {PATH}", printed)

    def test_a_listed_module_in_layout_fails_until_removed(self) -> None:
        code, printed, _ = self.run_check({PATH: self.GOOD,
                                           BASELINE_NAME: baseline_file([PATH])})
        self.assertEqual(code, 1)
        self.assertIn(f"FAIL: {PATH}: in layout but still listed in {BASELINE_NAME}; "
                      "remove its entry", printed)
        code, printed, _ = self.run_check({PATH: self.GOOD, BASELINE_NAME: baseline_file([])})
        self.assertEqual(code, 0, printed)

    def test_a_listed_module_in_statement_layout_only_stays_listed(self) -> None:
        # Fine for tools/format_vba.py, but without its header: still baselined.
        code, printed, _ = self.run_check({PATH: self.BAD, BASELINE_NAME: baseline_file([PATH])})
        self.assertEqual(code, 0, printed)
        self.assertNotIn("remove its entry", printed)

    def test_a_listed_path_that_is_not_a_tracked_module_fails(self) -> None:
        for name in ("tests/modules/Gone.bas", "docs/C.bas", "tests/modules/D.txt"):
            with self.subTest(name):
                code, printed, _ = self.run_check({
                    PATH: self.GOOD, "docs/C.bas": self.BAD, "tests/modules/D.txt": "x",
                    BASELINE_NAME: baseline_file([name])})
                self.assertEqual(code, 1)
                self.assertIn(f"FAIL: {name}: listed in {BASELINE_NAME} but not a tracked VBA "
                              "module under the source directories", printed)

    def test_a_module_outside_the_source_directories_is_not_read(self) -> None:
        code, printed, _ = self.run_check({PATH: self.GOOD, "docs/C.bas": self.BAD})
        self.assertEqual(code, 0, printed)
        self.assertIn("(1 VBA component(s), 0 baselined,", printed)

    def test_no_baseline_checks_every_module(self) -> None:
        files = {PATH: self.BAD, BASELINE_NAME: "{"}
        code, printed, _ = self.run_check(files, "--no-baseline")
        self.assertEqual(code, 1)
        self.assertIn(f"FAIL: {PATH}:", printed)
        self.assertIn("0 baselined", printed)

    def test_an_explicit_baseline_file(self) -> None:
        files = {PATH: self.BAD, "layout.json": baseline_file([PATH])}
        code, printed, _ = self.run_check(files, "--baseline", "layout.json")
        self.assertEqual(code, 0, printed)
        self.assertIn("listed in layout.json", printed)
        code, _, errors = self.run_check({PATH: self.GOOD}, "--baseline", "missing.json")
        self.assertEqual(code, 2)
        self.assertIn("vba-layout: ERROR: cannot read baseline missing.json", errors)

    def test_baseline_errors_exit_2(self) -> None:
        cases = {
            "not JSON": "{",
            "wrong schema": json.dumps({"schema_version": 2, "files": []}),
            "not a list": baseline_file("x"),
            "an absolute path": baseline_file(["/tests/modules/A.bas"]),
            "a backslash": baseline_file(["tests\\modules\\A.bas"]),
            "unsorted": baseline_file(["tests/modules/B.bas", "tests/modules/A.bas"]),
            "duplicated": baseline_file(["tests/modules/A.bas", "tests/modules/A.bas"]),
        }
        for label, text in cases.items():
            with self.subTest(label):
                code, _, errors = self.run_check({PATH: self.GOOD, BASELINE_NAME: text})
                self.assertEqual(code, 2)
                self.assertIn("vba-layout: ERROR:", errors)
                self.assertIn(BASELINE_NAME, errors)


if __name__ == "__main__":
    unittest.main()
