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
import vba_header_layout  # noqa: E402
from test_check_vba import repository  # noqa: E402

PATH = "tests/modules/TEST_Fixture.bas"
BANNER = "'" + "-" * 78
FIELDS = ("PURPOSE", "WHY THIS EXISTS", "INPUTS", "RETURNS", "BEHAVIOR", "ERROR POLICY",
          "DEPENDENCIES", "NOTES", "UPDATED")


def header(kind: str, fields: tuple[str, ...] = FIELDS, banner: str = BANNER) -> str:
    """A component or procedure header: a title line, then each field with text."""
    title = f"' {kind}" if ":" in kind else "'" + kind.center(78).rstrip()
    lines = ["'", banner, title, banner]
    for field in fields:
        text = "2026-11-21" if field == "UPDATED" else "None"
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
            f"Public Sub Run()\n{procedure}\n\n    Dim Count As Long\nEnd Sub\n")


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
