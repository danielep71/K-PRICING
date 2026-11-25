"""Valid and invalid fixtures for the VBA static checks (tools/check_vba_*.py).

Each fixture is a throw-away Git repository holding synthetic VBA components,
written with CRLF as the VBE exports them and stored under this repository's
.gitattributes. The components are fixtures for the checks only: no
calculation, product or regulatory value is implied by them. The adapted
files and their license are listed in docs/licenses/THIRD_PARTY_NOTICES.txt.
"""

from __future__ import annotations

import contextlib
import io
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from collections.abc import Callable, Iterator, Mapping
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_vba_calls  # noqa: E402
import check_vba_conditionals  # noqa: E402
import check_vba_dependencies  # noqa: E402
import check_vba_errors  # noqa: E402
import check_vba_identifiers  # noqa: E402
import check_vba_jumps  # noqa: E402
import check_vba_public_api  # noqa: E402
import check_vba_source  # noqa: E402
from vba_lexer import (  # noqa: E402
    Finding,
    logical_units,
    read_component,
    strip_code,
    tracked_components,
)

CORE = '''Attribute VB_Name = "CORE_Validation"
Option Explicit
Option Private Module

Public Function IsPositive( _
    ByVal Value As Double) _
    As Boolean

        On Error GoTo Failed
        IsPositive = Value > 0
        Exit Function

Failed:
        Err.Raise Err.Number, "CORE_Validation.IsPositive", Err.Description
End Function
'''
FACADE = '''Attribute VB_Name = "DEMO_Functions"
Option Explicit

Public Const DEMO_FORMAT_TEXT As String = "0.00"


Public Function DEMO_IsPositive( _
    ByVal Value As Double) _
    As Variant

        DEMO_IsPositive = CORE_Validation.IsPositive(Value)
End Function
'''
CLASS_ROWS = (
    "CCalculationResult\tProperty Get\tAmount\n"
    "# SIG\tCCalculationResult\tProperty Get\tAmount\tPublic Property Get Amount() As Double\n"
    "CCalculationResult\tProperty Let\tAmount\n"
    "# SIG\tCCalculationResult\tProperty Let\tAmount\t"
    "Public Property Let Amount(ByVal Value As Double)\n"
)
INVENTORY = (
    "# Public VBA declarations\n"
    "[project-public]\n"
    "CORE_Validation\tFunction\tIsPositive\n"
    "# SIG\tCORE_Validation\tFunction\tIsPositive\t"
    "Public Function IsPositive(ByVal Value As Double) As Boolean\n"
    + CLASS_ROWS +
    "[supported]\n"
    "DEMO_Functions\tConst\tDEMO_FORMAT_TEXT\n"
    '# SIG\tDEMO_Functions\tConst\tDEMO_FORMAT_TEXT\tPublic Const DEMO_FORMAT_TEXT As String = "0.00"\n'
    "DEMO_Functions\tFunction\tDEMO_IsPositive\n"
    "# SIG\tDEMO_Functions\tFunction\tDEMO_IsPositive\t"
    "Public Function DEMO_IsPositive(ByVal Value As Double) As Variant\n"
)
CLASS = '''VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCalculationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mAmount As Double  'Synthetic value of the fixture


Public Property Get Amount() As Double
        Amount = mAmount
End Property


Public Property Let Amount(ByVal Value As Double)
        mAmount = Value
End Property
'''
WORKBOOK = '''VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
'''
FORM = '''VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRunOptions
   Caption         =   "Run options"
   OleObjectBlob   =   "frmRunOptions.frx":0000
End
Attribute VB_Name = "frmRunOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''
HARNESS = '''Attribute VB_Name = "TEST_Harness"
Option Explicit
Option Private Module

Public Sub RunAll()
        Debug.Print "It's ""quoted"": GoTo Nowhere", ['Sheet 1'!A1]
        Rem A comment with an open " quote
End Sub
'''
EXAMPLE = '''Attribute VB_Name = "EX_Inputs"
Option Explicit

Public Sub Run()
        Debug.Print DEMO_Functions.DEMO_IsPositive(1)
End Sub
'''
VALID: dict[str, str | bytes] = {
    "src/core/CORE_Validation.bas": CORE,
    "src/modules/DEMO_Functions.bas": FACADE,
    "src/classes/CCalculationResult.cls": CLASS,
    "src/workbook/ThisWorkbook.cls": WORKBOOK,
    "src/forms/frmRunOptions.frm": FORM,
    "src/forms/frmRunOptions.frx": b"\x00\x01\x02\x03",
    "tests/modules/TEST_Harness.bas": HARNESS,
    "examples/modules/EX_Inputs.bas": EXAMPLE,
    "docs/PUBLIC_API.txt": INVENTORY,
}


# The .git of an empty repository, made by one git init and copied into each
# fixture, which takes about half the time of a git init per fixture.
TEMPLATE: tempfile.TemporaryDirectory[str] | None = None


def remove_template() -> None:
    global TEMPLATE
    if TEMPLATE is not None:
        TEMPLATE.cleanup()
        TEMPLATE = None


def empty_git() -> Path:
    """The template .git, made on first use and removed once the tests of the
    module that used it have run."""
    global TEMPLATE
    if TEMPLATE is None:
        TEMPLATE = tempfile.TemporaryDirectory(prefix="vba-check-template-")
        subprocess.run(["git", "init", "-q", TEMPLATE.name], check=True)
        # Git never runs a sample hook, so the copies do without them.
        for sample in Path(TEMPLATE.name, ".git", "hooks").glob("*.sample"):
            sample.unlink()
        unittest.addModuleCleanup(remove_template)
    return Path(TEMPLATE.name) / ".git"


@contextlib.contextmanager
def repository(files: Mapping[str, str | bytes], *, attributes: bool = True) -> Iterator[Path]:
    """A Git repository holding ``files``: text written with CRLF, then staged."""
    with tempfile.TemporaryDirectory(prefix="vba-check-") as temporary:
        root = Path(temporary)
        shutil.copytree(empty_git(), root / ".git")
        if attributes:
            shutil.copyfile(ROOT / ".gitattributes", root / ".gitattributes")
        for relative, content in files.items():
            path = root / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            if isinstance(content, bytes):
                data = content
            elif relative.endswith((".bas", ".cls", ".frm")):
                data = content.replace("\n", "\r\n").encode("ascii", errors="surrogateescape")
            else:
                data = content.encode("utf-8")
            path.write_bytes(data)
        subprocess.run(["git", "-C", str(root), "-c", "core.autocrlf=false", "-c",
                        "core.safecrlf=false", "add", "--all"], check=True)
        yield root


DEMO_STYLE = "| Public facade module | `DEMO_` plus subject | `DEMO_Functions` |\n"


def run_check(check: Callable[[Path], tuple[list[Finding], *tuple[int, ...]]],
              files: Mapping[str, str | bytes], *, attributes: bool = True) -> list[str]:
    with repository(files, attributes=attributes) as root:
        return [str(finding) for finding in check(root)[0]]


def changed(path: str, old: str, new: str, files: Mapping[str, str | bytes] = VALID
            ) -> dict[str, str | bytes]:
    text = files[path]
    assert isinstance(text, str) and old in text, (path, old)
    return {**files, path: text.replace(old, new, 1)}


def standalone(source: str) -> dict[str, str | bytes]:
    """One test module holding ``source``, plus an empty inventory."""
    return {"tests/modules/TEST_Fixture.bas": source, "docs/PUBLIC_API.txt": "# None\n"}


class SourceTest(unittest.TestCase):
    def findings(self, files: Mapping[str, str | bytes], **options: bool) -> list[str]:
        return run_check(check_vba_source.check, files, **options)

    def assertFinding(self, files: Mapping[str, str | bytes], text: str) -> None:
        findings = self.findings(files)
        self.assertTrue(any(text in item for item in findings), findings)

    def test_valid_project_passes(self) -> None:
        self.assertEqual(self.findings(VALID), [])

    def test_empty_repository_passes(self) -> None:
        self.assertEqual(self.findings({"README.md": "# Empty\n"}), [])

    def test_component_outside_its_homes(self) -> None:
        self.assertFinding({"src/CORE_Stray.bas": CORE.replace("CORE_Validation", "CORE_Stray")},
                           "outside its homes")
        self.assertFinding({"src/core/sub/CORE_Validation.bas": CORE}, "outside its homes")
        self.assertFinding({"tools/CORE_Validation.bas": CORE}, "outside its homes")

    def test_home_without_that_component_kind(self) -> None:
        self.assertFinding({"src/core/CORE_Validation.cls": CORE}, "holds no .cls components")
        self.assertFinding({"src/modules/frmRunOptions.frm": FORM}, "holds no .frm components")

    def test_role_prefixes(self) -> None:
        cases = {
            "src/core/Validation.bas": "CORE_<Subject>",
            "src/modules/Functions.bas": "<REGIME>_<Subject>",
            "src/classes/CalculationResult.cls": "C<Noun>",
            "src/workbook/Inputs.cls": "ThisWorkbook or sh<Noun>",
            "src/workbook/Main.bas": "M_<Subject>",
            "tests/modules/Harness.bas": "TEST_<Subject>",
            "examples/modules/Inputs.bas": "EX_<Subject>",
        }
        for path, rule in cases.items():
            with self.subTest(path=path):
                source = CORE.replace("CORE_Validation", Path(path).stem)
                self.assertFinding({path: source}, f"is named {rule}")

    def test_form_resource_rules(self) -> None:
        without = {k: v for k, v in VALID.items() if not k.endswith(".frx")}
        self.assertFinding(without, "form resource frmRunOptions.frx is not tracked")
        self.assertFinding({**VALID, "src/forms/frmRunOptions.frx": b""},
                           "resource offset is outside")
        self.assertFinding(changed("src/forms/frmRunOptions.frm",
                                   '"frmRunOptions.frx":0000', '"Other"'),
                           "no form resource")
        self.assertFinding({"src/workbook/frmLost.frx": b"\x00"}, "beside its form")

    def test_vb_name_must_equal_file_name(self) -> None:
        self.assertFinding(changed("src/core/CORE_Validation.bas", 'VB_Name = "CORE_Validation"',
                                   'VB_Name = "CORE_Checks"'),
                           'exactly one Attribute VB_Name = "CORE_Validation"')

    def test_bas_starts_with_vb_name(self) -> None:
        self.assertFinding(changed("src/core/CORE_Validation.bas", "Attribute VB_Name",
                                   "'Banner\nAttribute VB_Name"),
                           "starts with its Attribute VB_Name line")

    def test_names_are_case_insensitive(self) -> None:
        files = {"src/core/CORE_Ab.bas": CORE.replace("CORE_Validation", "CORE_Ab"),
                 "src/core/CORE_AB.bas": CORE.replace("CORE_Validation", "CORE_AB")}
        self.assertFinding(files, "component name also used by")

    def test_name_length_limit(self) -> None:
        name = "CORE_" + "A" * 27
        self.assertFinding({f"src/core/{name}.bas": CORE.replace("CORE_Validation", name)},
                           "longer than 31 characters")

    def test_module_options(self) -> None:
        self.assertFinding(changed("src/modules/DEMO_Functions.bas", "Option Explicit\n", ""),
                           "missing Option Explicit")
        for path in ("src/core/CORE_Validation.bas", "tests/modules/TEST_Harness.bas"):
            with self.subTest(path=path):
                self.assertFinding(changed(path, "Option Private Module\n", ""),
                                   "missing Option Private Module")
        for base in ("0", "1"):
            with self.subTest(base=base):
                self.assertFinding(changed("src/core/CORE_Validation.bas", "Option Explicit\n",
                                           f"Option Explicit\nOption Base {base}\n"),
                                   "Option Base; declare explicit array bounds")

    def test_options_in_comments_do_not_count(self) -> None:
        core = "src/core/CORE_Validation.bas"
        continued = "Option Explicit\n' An example _\nOption Base 1\n"
        self.assertFalse(any("Option Base" in item for item in
                             self.findings(changed(core, "Option Explicit\n", continued))))
        commented = "' Note _\nOption Explicit\n"
        self.assertFinding(changed(core, "Option Explicit\n", commented),
                           "missing Option Explicit")
        self.assertFinding(changed(core, "Option Private Module\n",
                                   "' Note _\nOption Private Module\n"),
                           "missing Option Private Module")
        self.assertFinding(changed(core, "Option Explicit\n", "' Option Explicit\n"),
                           "missing Option Explicit")

    def test_ascii_only(self) -> None:
        self.assertFinding(changed("src/core/CORE_Validation.bas", "Err.Description",
                                   "Err.Description & \"\udce9\""), "non-ASCII byte 0xE9")
        bom = {"src/core/CORE_Validation.bas": b"\xef\xbb\xbf" + CORE.encode("ascii")}
        self.assertFinding(bom, "byte-order mark")

    def test_tabs_and_trailing_white_space(self) -> None:
        self.assertFinding(changed("src/core/CORE_Validation.bas", "        IsPositive =",
                                   "\tIsPositive ="), "tab character")
        self.assertFinding(changed("src/core/CORE_Validation.bas", "Exit Function",
                                   "Exit Function  "), "trailing white space")

    def test_form_header_keeps_the_bytes_the_vbe_writes(self) -> None:
        # The VBE ends a form's Begin line with a blank and may write other
        # trailing blanks or a tab there; tools/format_vba.py keeps them.
        header = changed("src/forms/frmRunOptions.frm", "frmRunOptions\n   Caption",
                         "frmRunOptions \n\tClientHeight    =   3015\n   Caption")
        header = changed("src/forms/frmRunOptions.frm", '"Run options"', '"Run options"   ',
                         header)
        self.assertEqual(self.findings(header), [])
        code = changed("src/forms/frmRunOptions.frm", "Option Explicit\n",
                       "Option Explicit \n\tPrivate X As Long\n", header)
        self.assertEqual(self.findings(code), [
            "src/forms/frmRunOptions.frm:12: trailing white space",
            "src/forms/frmRunOptions.frm:13: tab character; indent with spaces"])
        # A module's header has no such exemption.
        self.assertFinding(changed("src/core/CORE_Validation.bas", '"CORE_Validation"',
                                   '"CORE_Validation" '), "CORE_Validation.bas:1: trailing")

    def test_unclosed_string_literal(self) -> None:
        self.assertFinding(changed("tests/modules/TEST_Harness.bas", 'Debug.Print "It',
                                   'Debug.Print "Open\n        Debug.Print "It'),
                           "string literal not closed")

    def test_comment_continuation_is_comment(self) -> None:
        source = changed("tests/modules/TEST_Harness.bas", "End Sub",
                         "        ' A comment that continues _\n        on \" this line\nEnd Sub")
        self.assertEqual(self.findings(source), [])

    def test_storage_needs_lf_and_crlf_checkout(self) -> None:
        findings = self.findings({"src/core/CORE_Validation.bas": CORE}, attributes=False)
        self.assertTrue(any("stored with crlf" in item for item in findings), findings)
        self.assertTrue(any("check it out with CRLF" in item for item in findings), findings)

    def test_workbooks_are_not_committed(self) -> None:
        self.assertFinding({"src/workbook/Workbook.xlsm": b"PK\x03\x04"}, "workbook committed")
        source = {"docs/methodology/sources/Annex.xlsx": b"PK\x03\x04"}
        self.assertFalse(any("workbook committed" in item for item in self.findings(source)))

    def test_main_exit_codes(self) -> None:
        for files, expected in ((VALID, 0), ({"src/Stray.bas": CORE}, 1)):
            with self.subTest(expected=expected), repository(files) as root, \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(check_vba_source.main(["--root", str(root)]), expected)
                self.assertIn("static check, nothing compiled or run", output.getvalue())


CONDITIONAL_VALID = {
    "nested": '''#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function F64 Lib "kernel32" () As LongPtr
    #ElseIf Win32 Then
        Private Declare PtrSafe Function F32 Lib "kernel32" () As Long
    #End If
#Else
    Private Declare Function FLegacy Lib "kernel32" () As Long
#End If
''',
    "inactive-legacy": '''#If Not VBA7 Then
    #If Win64 Then
        Private Declare Function Impossible Lib "kernel32" () As Long
    #End If
#End If
''',
    "boolean-expression": '''#If Not VBA7 Then
    Private Declare Function Legacy Lib "kernel32" () As Long
#ElseIf VBA7 And (Win64 Or Win32) Then
    Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
#End If
''',
    "continued-declare": '''#If VBA7 Then
    Private Declare PtrSafe Function Modern Lib "kernel32" _
        () As LongPtr
#Else
    Private Declare Function Legacy Lib "kernel32" _
        () As Long
#End If
''',
    "colon-separated-ptrsafe": '''Private Const N As Long = 1: Private Declare PtrSafe Function B Lib "kernel32" () As LongPtr
''',
    "colon-separated-legacy-under-vba6": '''#If Not VBA7 Then
    Private Const N As Long = 1: Private Declare Function B Lib "kernel32" () As Long
#End If
''',
    "literals": '''#If True And Not 0 Then
    Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
#End If
''',
    "module-const-false": '''#Const WRITE_STATUS_BAR = False
Private Sub Status(ByVal Message As String)
#If WRITE_STATUS_BAR Then
        Application.StatusBar = Message
#End If
End Sub
#If WRITE_STATUS_BAR Then
Private Declare Function Bad Lib "kernel32" () As Long
#End If
''',
    "module-const-expression": '''#Const LEGACY = Not VBA7
#Const MODERN = Not LEGACY And (Win64 Or Win32) And -1 <> 0
#If LEGACY Then
    Private Declare Function Legacy Lib "kernel32" () As Long
#ElseIf MODERN Then
    Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
#End If
''',
    "module-const-is-case-insensitive": '''#Const Write_Status_Bar = False
#If WRITE_STATUS_BAR Then
    Private Declare Function Bad Lib "kernel32" () As Long
#End If
''',
    "module-const-per-environment": '''#If VBA7 Then
    #Const HOST = True
#Else
    #Const HOST = False
#End If
#If HOST Then
    Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
#End If
''',
    # HOST exists only under VBA7; VBA6 never enters the outer branch, so it
    # never evaluates the nested #If.
    "const-used-within-its-defining-branch": '''#If VBA7 Then
    #Const HOST = True
    #If HOST Then
        Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
    #End If
#Else
    Private Declare Function Legacy Lib "kernel32" () As Long
#End If
''',
    # LEGACY exists only under VBA6; under VBA7 the first branch is taken, so
    # the #ElseIf is not evaluated there.
    "elseif-not-evaluated-after-a-taken-branch": '''#If Not VBA7 Then
    #Const LEGACY = True
#End If
#If VBA7 Then
    Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
#ElseIf LEGACY Then
    Private Declare Function Legacy Lib "kernel32" () As Long
#End If
''',
}
CONDITIONAL_INVALID = {
    "declare-without-ptrsafe": ('''#If VBA7 And Win64 Then
    Private Declare Function Bad Lib "kernel32" () As Long
#End If
''', "without PtrSafe in vba7-win64"),
    "unconditional-declare": ('Private Declare Function Bad Lib "kernel32" () As Long\n',
                              "without PtrSafe in vba7-win32, vba7-win64"),
    "win32-is-true-on-win64": ('''#If Win32 Then
    Private Declare Function WrongOrder Lib "kernel32" () As Long
#ElseIf Win64 Then
    Private Declare PtrSafe Function NeverReached Lib "kernel32" () As LongPtr
#End If
''', "vba7-win64"),
    "colon-separated-declare": ('Private Const N As Long = 1: Private Declare Function B Lib "kernel32" '
                                '() As Long\n', "without PtrSafe in vba7-win32, vba7-win64"),
    "declare-after-two-colons": ('Private Const N As Long = 1: Private Const M As Long = 2: '
                                 'Private Declare Sub C Lib "kernel32" ()\n',
                                 "without PtrSafe in vba7-win32, vba7-win64"),
    "unknown-symbol": ("#If Mac Then\n#End If\n", "unsupported conditional-compilation symbol"),
    "vba6-is-true-under-vba7": ('''#If VBA6 Then
    Private Declare Function Legacy Lib "kernel32" () As Long
#End If
''', "without PtrSafe in vba7-win32, vba7-win64"),
    "integer-literal": ("#If 2 Then\n#End If\n", "-1 and 0"),
    "missing-then": ("#If VBA7\n#End If\n", "must end with Then"),
    "module-const-true-reaches-declare": ('''#Const ENABLED = True
#If ENABLED Then
    Private Declare Function Bad Lib "kernel32" () As Long
#End If
''', "without PtrSafe in vba7-win32, vba7-win64"),
    "const-used-before-definition": ("#If LATER Then\n#End If\n#Const LATER = True\n",
                                     "unsupported conditional-compilation symbol 'LATER'"),
    "const-defined-in-one-environment-only": ("#If VBA7 Then\n#Const HOST = True\n#End If\n"
                                              "#If HOST Then\n#End If\n",
                                              "unsupported conditional-compilation symbol "
                                              "'HOST'"),
    # A #Const defined in an inactive branch does not leak: VBA6 reaches the
    # #ElseIf (Win64 is False there) without HOST.
    "elseif-uses-const-from-another-environment": ("#If VBA7 Then\n#Const HOST = True\n#End If\n"
                                                   "#If Win64 Then\n#ElseIf HOST Then\n#End If\n",
                                                   "indeterminate #ElseIf directive: unsupported "
                                                   "conditional-compilation symbol 'HOST'"),
    "nested-if-uses-const-outside-its-branch": ("#If VBA7 Then\n#Const HOST = True\n#End If\n"
                                                "#If Not VBA7 Then\n#If HOST Then\n#End If\n"
                                                "#End If\n",
                                                "indeterminate #If directive: unsupported "
                                                "conditional-compilation symbol 'HOST'"),
    "const-redefines-compiler-constant": ("#Const Win64 = False\n",
                                          "#Const Win64 redefines a compiler constant"),
    "duplicate-const": ("#Const FLAG = True\n#Const FLAG = False\n",
                        "#Const FLAG is defined twice in one environment"),
    "duplicate-const-in-a-branch": ("#Const FLAG = True\n#If VBA7 Then\n#Const FLAG = False\n"
                                    "#End If\n", "#Const FLAG is defined twice"),
    "malformed-const-without-value": ("#Const FLAG\n", "malformed #Const"),
    "malformed-const-without-name": ("#Const = True\n", "malformed #Const"),
    "const-named-like-an-operator": ("#Const Not = True\n", "malformed #Const"),
    "const-not-boolean": ("#Const LEVEL = 2\n", "indeterminate #Const LEVEL: only the Boolean"),
    "const-without-expression": ("#Const FLAG =\n", "indeterminate #Const FLAG: empty"),
    "const-unknown-symbol": ("#Const FLAG = PROJECT_ARG\n",
                             "indeterminate #Const FLAG: unsupported conditional-compilation "
                             "symbol 'PROJECT_ARG'"),
    "unclosed": ("#If VBA7 Then\n", "not closed"),
    "stray-elseif": ("#ElseIf VBA7 Then\n", "#ElseIf without #If"),
    "elseif-after-else": ("#If VBA7 Then\n#Else\n#ElseIf Win64 Then\n#End If\n",
                          "#ElseIf after #Else"),
    "duplicate-else": ("#If VBA7 Then\n#Else\n#Else\n#End If\n", "duplicate #Else"),
    "stray-end-if": ("#End If\n", "#End If without #If"),
    "unsupported-directive": ("#Region Main\n", "unsupported conditional-compilation directive"),
}


def module(body: str) -> str:
    return f'Attribute VB_Name = "TEST_Fixture"\nOption Explicit\nOption Private Module\n{body}'


class ConditionalTest(unittest.TestCase):
    def test_valid_fixtures_pass(self) -> None:
        self.assertEqual(run_check(check_vba_conditionals.check, VALID), [])
        for name, body in CONDITIONAL_VALID.items():
            with self.subTest(name=name):
                self.assertEqual(run_check(check_vba_conditionals.check,
                                           standalone(module(body))), [])

    def test_invalid_fixtures_fail(self) -> None:
        for name, (body, text) in CONDITIONAL_INVALID.items():
            with self.subTest(name=name):
                findings = run_check(check_vba_conditionals.check, standalone(module(body)))
                self.assertTrue(any(text in item for item in findings), findings)

    def test_reachable_sources_blank_inactive_lines(self) -> None:
        sources, findings = check_vba_conditionals.reachable_sources(
            "x.bas", "A\n#If VBA7 Then\nB\n#Else\nC\n#End If")
        self.assertEqual(findings, [])
        self.assertEqual(sources["vba7-win64"].split("\n"), ["A", "", "B", "", "", ""])
        self.assertEqual(sources["vba6-win32"].split("\n"), ["A", "", "", "", "C", ""])
        sources, findings = check_vba_conditionals.reachable_sources(
            "x.bas", "#Const ON = Win64\n#If ON Then\nB\n#End If")
        self.assertEqual(findings, [])
        self.assertEqual(sources["vba7-win64"].split("\n"), ["", "", "B", ""])
        self.assertEqual(sources["vba7-win32"].split("\n"), ["", "", "", ""])

    def test_reachable_sources_returns_independent_results_for_the_same_input(self) -> None:
        text = "A\n#If VBA7 Then\nB\n#End If\n#If\n"
        first, first_findings = check_vba_conditionals.reachable_sources("y.bas", text)
        expected = (dict(first), list(first_findings))
        self.assertTrue(first_findings)
        first["vba7-win64"] = "changed"
        first_findings.clear()
        second, second_findings = check_vba_conditionals.reachable_sources("y.bas", text)
        self.assertEqual((second, second_findings), expected)
        self.assertIsNot(second, first)

    def test_module_const_false_blanks_its_branch_everywhere(self) -> None:
        sources, findings = check_vba_conditionals.reachable_sources(
            "x.bas", "A\n#Const WRITE_STATUS_BAR = False\n#If WRITE_STATUS_BAR Then\nB\n"
                     "#End If\nC")
        self.assertEqual(findings, [])
        for name, source in sources.items():
            with self.subTest(environment=name):
                self.assertEqual(source.split("\n"), ["A", "", "", "", "", "C"])

    def test_branch_local_const_selects_per_environment(self) -> None:
        sources, findings = check_vba_conditionals.reachable_sources(
            "x.bas", "#If VBA7 Then\n#Const HOST = True\n#If HOST Then\nB\n#End If\n#End If")
        self.assertEqual(findings, [])
        for name in ("vba7-win32", "vba7-win64"):
            self.assertEqual(sources[name].split("\n")[3], "B")
        self.assertEqual(sources["vba6-win32"].split("\n")[3], "")
        sources, findings = check_vba_conditionals.reachable_sources(
            "x.bas", "#If VBA7 Then\n#Const HOST = True\n#Else\n#Const HOST = False\n#End If\n"
                     "#If HOST Then\nA\n#Else\nB\n#End If")
        self.assertEqual(findings, [])
        self.assertEqual(sources["vba7-win64"].split("\n")[6:9], ["A", "", ""])
        self.assertEqual(sources["vba7-win32"].split("\n")[6:9], ["A", "", ""])
        self.assertEqual(sources["vba6-win32"].split("\n")[6:9], ["", "", "B"])
        sources, findings = check_vba_conditionals.reachable_sources(
            "x.bas", "#If Not VBA7 Then\n#Const LEGACY = True\n#End If\n"
                     "#If VBA7 Then\nA\n#ElseIf LEGACY Then\nB\n#End If")
        self.assertEqual(findings, [])
        self.assertEqual(sources["vba7-win64"].split("\n")[4:7], ["A", "", ""])
        self.assertEqual(sources["vba6-win32"].split("\n")[4:7], ["", "", "B"])

    def test_an_undefined_symbol_is_reported_once_with_its_line(self) -> None:
        body, text = CONDITIONAL_INVALID["const-defined-in-one-environment-only"]
        findings = run_check(check_vba_conditionals.check, standalone(module(body)))
        self.assertEqual(findings, ["tests/modules/TEST_Fixture.bas:7: indeterminate #If directive: "
                                    "unsupported conditional-compilation symbol 'HOST'; supported "
                                    "are VBA6, VBA7, Win32, Win64 and a #Const defined earlier in "
                                    "the module"])

    def test_module_const_passes_the_checks_that_share_the_model(self) -> None:
        # A module whose only conditional directives are a valid #Const and
        # the #If that reads it.
        body = ("#Const WRITE_STATUS_BAR = False\n\nPublic Sub Report(ByVal Message As String)\n"
                "#If WRITE_STATUS_BAR Then\n        Application.StatusBar = Message\n#End If\n"
                "End Sub\n")
        files = standalone(module(body))
        for check in (check_vba_conditionals.check, check_vba_identifiers.check,
                      check_vba_jumps.check, check_vba_public_api.check,
                      check_vba_errors.check, check_vba_calls.check):
            with self.subTest(check=check.__module__):
                self.assertEqual(run_check(check, files), [])
        facade = API_FACADE + ("\n#Const WRITE_STATUS_BAR = False\n"
                               "#If WRITE_STATUS_BAR Then\n"
                               "Private Sub Status()\n        Application.StatusBar = False\n"
                               "End Sub\n#End If\n")
        self.assertEqual(run_check(check_vba_public_api.check, api_files(facade)), [])

    def test_the_other_checks_read_the_branch_the_const_selects(self) -> None:
        body = ("#Const KEEP_TOTAL = {}\n\nPublic Sub Run()\n#If KEEP_TOTAL Then\n"
                "        Dim Total As Long\n#End If\n        Total = 1\nEnd Sub\n")
        self.assertEqual(run_check(check_vba_identifiers.check,
                                   standalone(module(body.format("True")))), [])
        findings = run_check(check_vba_identifiers.check, standalone(module(body.format("False"))))
        self.assertTrue(any("Total is assigned in Run but not declared" in item
                            and "vba6-win32, vba7-win32, vba7-win64" in item
                            for item in findings), findings)
        jump = ("#Const USE_HANDLER = {}\n\nPublic Sub Run()\n        On Error GoTo Failed\n"
                "        Exit Sub\n#If USE_HANDLER Then\nFailed:\n#End If\nEnd Sub\n")
        self.assertEqual(run_check(check_vba_jumps.check,
                                   standalone(module(jump.format("True")))), [])
        self.assertNotEqual(run_check(check_vba_jumps.check,
                                      standalone(module(jump.format("False")))), [])


JUMPS_VALID = {
    "local-handler": '''Public Sub First(ByVal Value As Long)
        On Error GoTo Handler
        GoSub Worker
        Resume 100
Worker:
        Return
100:    Exit Sub
Handler:
        Resume Next
End Sub
''',
    # Resume 0 is VBA's spelling of Resume: it names no label.
    "resume-zero": '''Public Sub First()
        On Error GoTo Handler
        Exit Sub
Handler:
        Resume 0
End Sub
''',
    "same-label-in-two-procedures": '''Private Sub First()
Shared:
        Exit Sub
End Sub

Private Sub Second()
Shared:
        Exit Sub
End Sub
''',
    "continued-comment": '''Private Sub First()
        ' A comment that continues _
          GoTo Missing
        Exit Sub
End Sub
''',
    "numbered-label": '''Private Sub First()
        GoTo 100
100     Exit Sub
End Sub
''',
    "special-error-controls": '''Private Sub First()
        On Error GoTo 0
        On Error GoTo -1
        On Error Resume Next
        Resume Next
End Sub
''',
    "quoted-jump-text": '''Private Sub First()
        Debug.Print "GoSub Missing", "GoTo Lost", "Resume Nowhere"
        Debug.Print "He said ""GoSub Missing""; it's text"
        Debug.Print "GoSub Missing" & _
            "Resume Lost"
        Debug.Print "GoSub Missing": GoSub Worker
        Exit Sub
Worker:
        Return
End Sub
''',
    "keyword-before-colon": '''Private Sub First(ByVal Flag As Boolean)
        If Flag Then
            Debug.Print 1
        Else: Debug.Print 2
        End If
        If Not Flag Then
            Debug.Print 3
        Else: Debug.Print 4
        End If
End Sub
''',
    "rem-after-separator": '''Private Sub First()
        Debug.Print 1: Rem GoTo Missing
        Debug.Print 2:Rem GoSub Lost
End Sub
''',
    "one-line-procedure": '''Private Sub First(): Debug.Print 1: End Sub

Private Sub Second(): Debug.Print 2: End Sub
''',
    "rem-after-line-number": '''Private Sub First()
100 Rem GoTo Missing
        Exit Sub
End Sub
''',
    "numbered-comment-line-as-target": '''Private Sub First()
        GoTo 100
100 Rem target
        GoTo 200
200 ' target
End Sub
''',
    "bracketed-sheet-name": '''Private Sub First(): Debug.Print ['Sheet 1'!A1]: End Sub

Private Sub Second()
        Debug.Print [A1:B2]: GoTo Done
Done:
End Sub
''',
    "jump-keyword-in-a-sheet-name": '''Private Sub First()
        Debug.Print ['GoTo Missing'!A1], [GoSub!B2]: GoTo Done
Done:
End Sub
''',
    "computed-jump": '''Private Sub First(ByVal Choice As Long)
        On Choice GoTo One, Two
        On Choice GoSub One, Two
One:
Two:
End Sub
''',
    "colons-in-a-date-literal": '''Private Sub First()
        Debug.Print #12:30:00 PM#: GoTo Done
Done:
End Sub
''',
    "exclusive-branches": '''Private Sub First()
#If VBA7 Then
        GoTo Modern
Modern:
#Else
        GoTo Legacy
Legacy:
#End If
End Sub
''',
    "member-calls-named-like-jumps": '''Private Sub First()
        Application.Goto Reference:=ThisWorkbook.Worksheets("Inputs").Range("A1"), Scroll:=True
        With Application
            .Goto ThisWorkbook.Worksheets("Inputs").Range("A1")
        End With
End Sub
''',
}
JUMPS_INVALID = {
    "cross-procedure": ('''Private Sub First()
        GoTo Shared
End Sub

Private Sub Second()
Shared:
        Exit Sub
End Sub
''', "GoTo target 'Shared' is not a label of procedure First"),
    "duplicate-label": ('''Private Sub First()
Again:
Again:
        Exit Sub
End Sub
''', "duplicate label 'Again' in procedure First; first defined at line"),
    "jump-after-string": ('''Private Sub First()
        Debug.Print "it's ""quoted""": GoSub Missing
End Sub
''', "GoSub target 'Missing'"),
    "resume-to-missing-label": ('''Private Sub First()
        On Error GoTo Handler
        Exit Sub
Handler:
        Resume Retry
End Sub
''', "Resume target 'Retry'"),
    "label-in-other-branch": ('''Private Sub First()
#If VBA7 Then
        GoTo Legacy
#Else
Legacy:
#End If
End Sub
''', "[vba7-win32, vba7-win64]"),
    "procedure-not-closed": ("Private Sub First()\n        Exit Sub\n", "is not closed"),
    "procedure-opened-twice": ("Private Sub First()\nPrivate Sub Second()\nEnd Sub\n",
                               "is not closed before Second"),
    "stray-end": ("End Sub\n", "End statement outside any procedure"),
    "computed-jump-missing-target": ("Private Sub First(ByVal Choice As Long)\n"
                                     "        On Choice GoTo Present, Missing\nPresent:\nEnd Sub\n",
                                     "GoTo target 'Missing'"),
    "jump-after-bracketed-sheet-name": ("Private Sub First()\n"
                                        "        Debug.Print ['Sheet 1'!A1]: GoTo Lost\nEnd Sub\n",
                                        "GoTo target 'Lost'"),
    "one-line-procedure-missing-label": ("Private Sub First(): GoTo Lost: End Sub\n",
                                         "GoTo target 'Lost'"),
    "implicit-line-number-jump-missing-target": ("Private Sub First(ByVal Flag As Boolean)\n"
                                                 "        If Flag Then 100 Else 200\n"
                                                 "100     Exit Sub\nEnd Sub\n",
                                                 "GoTo target '200'"),
    "jump-after-member-call": ("Private Sub First()\n"
                               "        Application.Goto Reference:=\"R1C1\": GoTo Lost\nEnd Sub\n",
                               "GoTo target 'Lost'"),
}
# A handler that saves the error and resumes to the cleanup, which raises the
# saved error again; RAISE marks the line before the raise.
RERAISE = '''Private Sub First()
        Dim SavedNumber As Long
        On Error GoTo Handler
        Debug.Print 1 / 0
Cleanup:
        Debug.Print "restore"
        ' RAISE
        If SavedNumber <> 0 Then Err.Raise SavedNumber, "TEST_Fixture.First"
        Exit Sub
Handler:
        SavedNumber = Err.Number
        Resume Cleanup
End Sub
'''
RAISE = "        ' RAISE\n"
TRAPPED = ("Err.Raise after Resume Cleanup is trapped again by handler 'Handler' of procedure "
           "First; put On Error GoTo 0 before the raise")


def reraise(before_raise: str = "", fixture: str = RERAISE) -> dict[str, str | bytes]:
    """The re-raise fixture with ``before_raise`` lines inserted before the raise."""
    return standalone(module(fixture.replace(RAISE, before_raise)))


def enclosed(blocks: str) -> dict[str, str | bytes]:
    """A procedure whose handler resumes to ``Cleanup:``, a label inside ``blocks``."""
    return standalone(module(f'''Private Sub First(ByVal Flag As Boolean)
        Dim Index As Long
        On Error GoTo Handler
        Debug.Print 1 / 0
{blocks}        Exit Sub
Handler:
        Resume Cleanup
End Sub
'''))


SAFE = "On Error GoTo 0: Err.Raise 5"
# Valid cleanups after "Resume Cleanup": every raise runs with the handler
# disabled, or is not trapped by a local handler. None may be reported.
SAFE_CLEANUPS = {
    "plain": f"Cleanup:\n        {SAFE}\n",
    "separate-lines": "Cleanup:\n        On Error GoTo 0\n        Err.Raise 5\n",
    "goto-minus-one-then-zero": "Cleanup:\n        On Error GoTo -1\n        On Error GoTo 0\n"
                                "        Err.Raise 5\n",
    "block-if": f"Cleanup:\n        If Flag Then\n            {SAFE}\n        End If\n",
    "if-else": f"Cleanup:\n        If Flag Then\n            {SAFE}\n        Else\n"
               "            Debug.Print 0\n        End If\n",
    "if-elseif": f"Cleanup:\n        If Flag Then\n            Debug.Print 0\n"
                 f"        ElseIf Index = 1 Then\n            {SAFE}\n        End If\n",
    "else-colon": f"Cleanup:\n        If Flag Then\n            Debug.Print 0\n"
                  f"        Else: {SAFE}\n        End If\n",
    "one-line-if": f"Cleanup:\n        If Flag Then {SAFE}\n",
    "one-line-if-else": f"Cleanup:\n        If Flag Then Debug.Print 0 Else {SAFE}\n",
    "nested-one-line-if": f"Cleanup:\n        If Flag Then If Index = 1 Then {SAFE}\n",
    "one-line-if-with-a-colon": f"Cleanup:\n        If Flag Then: {SAFE}\n",
    "select-case": f"Cleanup:\n        Select Case Index\n        Case 1\n            {SAFE}\n"
                   "        Case Else\n            Debug.Print 0\n        End Select\n",
    "case-colon": f"Cleanup:\n        Select Case Index\n        Case 1: {SAFE}\n"
                  "        End Select\n",
    "for": f"Cleanup:\n        For Index = 1 To 2\n            {SAFE}\n        Next Index\n",
    "for-each-combined-next": "Cleanup:\n        For Each Item In Items\n"
                              f"            For Each Other In Item\n                {SAFE}\n"
                              "        Next Other, Item\n",
    "do-while": f"Cleanup:\n        Do While Flag\n            {SAFE}\n        Loop\n",
    "do-until": f"Cleanup:\n        Do Until Flag\n            {SAFE}\n        Loop\n",
    "loop-until": f"Cleanup:\n        Do\n            {SAFE}\n        Loop Until Flag\n",
    "bare-loop": f"Cleanup:\n        Do\n            If Flag Then Exit Do\n        Loop\n"
                 f"        {SAFE}\n",
    "while-wend": f"Cleanup:\n        While Flag\n            {SAFE}\n        Wend\n",
    "exit-for": f"Cleanup:\n        For Index = 1 To 2\n            If Flag Then Exit For\n"
                f"        Next Index\n        {SAFE}\n",
    "with": f"Cleanup:\n        With Application\n            {SAFE}\n        End With\n",
    "straight-subroutine": f"Cleanup:\n        GoSub Restore\n        {SAFE}\n        Exit Sub\n"
                           "Restore:\n        Debug.Print 0\n        Return\n",
    "branching-subroutine": "Cleanup:\n        GoSub Disarm\n        Err.Raise 5\n        Exit Sub\n"
                            "Disarm:\n        If Flag Then Exit Sub\n        On Error GoTo 0\n"
                            "        Return\n",
    "computed-gosub": "Cleanup:\n        On Index GoSub Disarm\n        Err.Raise 5\n"
                      "        Exit Sub\nDisarm:\n        On Error GoTo 0\n        Return\n",
    "computed-goto": f"Cleanup:\n        On Index GoTo Finish\n        {SAFE}\nFinish:\n",
    "second-label": f"Cleanup:\n        Debug.Print 0\nRestore:\n        {SAFE}\n",
    "label-with-code": f"Cleanup: {SAFE}\n",
    "exit-before": "Cleanup:\n        If Flag Then Exit Sub\n        Stop\n"
                   f"        {SAFE}\n",
    "goto-skips": "Cleanup:\n        GoTo Finish\n        Err.Raise 5\nFinish:\n",
    "conditional-compilation": f"Cleanup:\n#If VBA7 Then\n        {SAFE}\n#Else\n"
                               f"        {SAFE}\n#End If\n",
    "continued-line": "Cleanup:\n        On Error GoTo 0: Err.Raise 5, _\n"
                      "            \"TEST_Fixture.First\"\n",
    "named-arguments": "Cleanup:\n        On Error GoTo 0\n        VBA.Err.Raise Number:=5, "
                       "Description:=\"Err.Raise\"\n",
    "call-raise": "Cleanup:\n        On Error GoTo 0\n        Call Err.Raise(5)\n",
    "error-statement": "Cleanup:\n        On Error GoTo 0\n        Error 5\n",
    "resume-next": "Cleanup:\n        On Error Resume Next\n        Err.Raise 5\n",
    "raise-in-text": "Cleanup:\n        Debug.Print \"Err.Raise 5\" ' Err.Raise 6\n",
    "nested-blocks": "Cleanup:\n        Select Case Index\n        Case 1\n"
                     "            For Index = 1 To 2\n                If Flag Then\n"
                     f"                    {SAFE}\n                End If\n"
                     "            Next Index\n        End Select\n",
}


class JumpTest(unittest.TestCase):
    def test_safe_cleanups_are_not_reported(self) -> None:
        for name, blocks in SAFE_CLEANUPS.items():
            with self.subTest(name=name):
                self.assertEqual(run_check(check_vba_jumps.check, enclosed(blocks)), [])
        with self.subTest(name="line-numbers"):
            numbered = '''Private Sub First()
        On Error GoTo 900
        Debug.Print 1 / 0
100     On Error GoTo 0: Err.Raise 5
        Exit Sub
900     Resume 100
End Sub
'''
            self.assertEqual(run_check(check_vba_jumps.check, standalone(module(numbered))), [])

    def test_valid_fixtures_pass(self) -> None:
        self.assertEqual(run_check(check_vba_jumps.check, VALID), [])
        for name, body in JUMPS_VALID.items():
            with self.subTest(name=name):
                self.assertEqual(run_check(check_vba_jumps.check, standalone(module(body))), [])

    def test_invalid_fixtures_fail(self) -> None:
        for name, (body, text) in JUMPS_INVALID.items():
            with self.subTest(name=name):
                findings = run_check(check_vba_jumps.check, standalone(module(body)))
                self.assertTrue(any(text in item for item in findings), findings)

    def test_counts_procedures(self) -> None:
        with repository(VALID) as root:
            _, components, procedures = check_vba_jumps.check(root)
        self.assertEqual((components, procedures), (7, 6))

    def test_reraise_trapped_by_the_local_handler_is_reported(self) -> None:
        every = " [vba6-win32, vba7-win32, vba7-win64]"
        self.assertEqual(run_check(check_vba_jumps.check, reraise()),
                         [f"tests/modules/TEST_Fixture.bas:10: {TRAPPED}{every}"])
        block = RERAISE.replace(
            "        If SavedNumber <> 0 Then Err.Raise SavedNumber, \"TEST_Fixture.First\"\n",
            "        If SavedNumber <> 0 Then\n"
            "            Err.Raise SavedNumber, \"TEST_Fixture.First\"\n"
            "        End If\n")
        cases = {
            # Enabling the handler again after disabling it restores the loop.
            "handler-enabled-again": "        On Error GoTo 0\n        On Error GoTo Handler\n",
            # GoTo -1 ends the error-handling state but keeps the handler enabled.
            "goto-minus-one": "        On Error GoTo -1\n",
            # The exit in one branch leaves the other branch on the path.
            "exit-in-one-branch": "        If SavedNumber = 0 Then\n            Exit Sub\n"
                                  "        End If\n",
            # A one-line If keeps its branch's state across its ":" statements.
            "handler-enabled-in-a-one-line-if":
                "        On Error GoTo 0\n"
                "        If SavedNumber <> 0 Then On Error GoTo Handler: Err.Raise 5\n",
            "handler-enabled-in-a-one-line-else":
                "        On Error GoTo 0\n        If SavedNumber = 0 Then Debug.Print 0 "
                "Else On Error GoTo Handler: Err.Raise 5\n",
        }
        for name, before in cases.items():
            with self.subTest(name=name):
                self.assertEqual([item.split(": ", 1)[1] for item in
                                  run_check(check_vba_jumps.check, reraise(before))],
                                 [TRAPPED + every])
        with self.subTest(name="raise-in-a-block-if"):
            self.assertEqual([item.split(": ", 1)[1] for item in
                              run_check(check_vba_jumps.check, reraise(fixture=block))],
                             [TRAPPED + every])

    def test_reraise_with_the_local_handler_disabled_passes(self) -> None:
        handler_raises = RERAISE.replace("        Resume Cleanup\n",
                                         "        Err.Raise Err.Number\n        Resume Cleanup\n")
        cases = {
            "goto-zero": reraise("        On Error GoTo 0\n"),
            "goto-zero-in-the-raising-branch": reraise(fixture=RERAISE.replace(
                "Then Err.Raise", "Then On Error GoTo 0: Err.Raise")),
            "handler-enabled-in-the-other-one-line-branch": reraise(
                "        On Error GoTo 0\n", RERAISE.replace(
                    "If SavedNumber <> 0 Then Err.Raise",
                    "If SavedNumber = 0 Then On Error GoTo Handler: Debug.Print 0 Else Err.Raise")),
            "goto-zero-in-a-block-if": reraise(fixture=RERAISE.replace(
                "        If SavedNumber <> 0 Then Err.Raise SavedNumber, \"TEST_Fixture.First\"\n",
                "        If SavedNumber <> 0 Then\n            On Error GoTo 0\n"
                "            Err.Raise SavedNumber, \"TEST_Fixture.First\"\n        End If\n")),
            # While the handler runs, a raise goes to the caller, not back to the handler.
            "raise-in-the-handler": reraise("        On Error GoTo 0\n", handler_raises),
            "exit-before-the-raise": reraise("        Exit Sub\n"),
            "no-local-handler": standalone(module('''Private Sub First()
Cleanup:
        Err.Raise 5, "TEST_Fixture.First"
End Sub
''')),
        }
        for name, files in cases.items():
            with self.subTest(name=name):
                self.assertEqual(run_check(check_vba_jumps.check, files), [])

    def test_resume_target_inside_blocks_keeps_them(self) -> None:
        reported = {
            "raise-after-the-if": '''        If Flag Then
Cleanup:
            Debug.Print 0
        Else
            On Error GoTo 0
        End If
        Err.Raise 5
''',
            "raise-after-the-select": '''        Select Case Index
        Case 1
Cleanup:
            Debug.Print 0
        Case Else
            On Error GoTo 0
        End Select
        Err.Raise 5
''',
            "raise-after-the-inner-if": '''        If Flag Then
            If Index = 1 Then
Cleanup:
                Debug.Print 0
            Else
                On Error GoTo 0
            End If
            Err.Raise 5
        Else
            On Error GoTo 0
        End If
''',
            "raise-after-the-loop": '''        For Index = 1 To 2
Cleanup:
            If Flag Then Exit For
        Next Index
        Err.Raise 5
''',
            "raise-after-a-combined-next": '''        If Flag Then
            For Index = 1 To 2
                For Inner = 1 To 2
Cleanup:
                    Debug.Print Inner
                Next Inner, Index
        Else
            On Error GoTo 0
        End If
        Err.Raise 5
''',
        }
        passing = {
            "raise-in-the-else-branch": '''        If Flag Then
Cleanup:
            On Error GoTo 0
        Else
            On Error GoTo Handler: Err.Raise 5
        End If
''',
            "raise-in-the-next-case": '''        Select Case Index
        Case 1
Cleanup:
            On Error GoTo 0
        Case 2
            Err.Raise 5
        End Select
''',
            "raise-in-the-outer-else-branch": '''        If Flag Then
            If Index = 1 Then
Cleanup:
                On Error GoTo 0
            Else
                Err.Raise 5
            End If
        Else
            Err.Raise 6
        End If
''',
            # The next pass disables the handler before the target.
            "loop-disables-the-handler": '''        For Index = 1 To 2
            On Error GoTo 0
Cleanup:
            Debug.Print Index
        Next Index
        Err.Raise 5
''',
            "exit-the-loop-disabled": '''        Do While Index < 2
Cleanup:
            On Error GoTo 0
            If Flag Then Exit Do
        Loop
        Err.Raise 5
''',
            # Next Inner, Index closes both loops, so the Else stays a sibling.
            "raise-in-the-else-after-a-combined-next": '''        If Flag Then
            For Index = 1 To 2
                For Inner = 1 To 2
Cleanup:
                    On Error GoTo 0
                Next Inner, Index
        Else
            On Error GoTo Handler: Err.Raise 5
        End If
''',
        }
        every = " [vba6-win32, vba7-win32, vba7-win64]"
        for name, blocks in reported.items():
            with self.subTest(name=name):
                self.assertEqual([item.split(": ", 1)[1] for item in
                                  run_check(check_vba_jumps.check, enclosed(blocks))],
                                 [TRAPPED + every])
        for name, blocks in passing.items():
            with self.subTest(name=name):
                self.assertEqual(run_check(check_vba_jumps.check, enclosed(blocks)), [])

    def test_reraise_rule_reads_each_construct(self) -> None:
        reported = {
            # The raise after "Then:" belongs to the one-line If.
            "one-line-if-with-a-colon": ("Cleanup:\n        If Flag Then: Err.Raise 5\n",
                                         "Err.Raise"),
            "subroutine-without-on-error": ('''Cleanup:
        GoSub Report
        Err.Raise 5
        Exit Sub
Report:
        Debug.Print 0
        Return
''', "Err.Raise"),
            "error-statement": ("Cleanup:\n        Error 5\n", "Error"),
            "qualified-raise": ("Cleanup:\n        VBA.Err.Raise Number:=5\n", "Err.Raise"),
            "called-raise": ("Cleanup:\n        Call Err.Raise(5, \"TEST_Fixture.First\")\n",
                             "Err.Raise"),
            # A branch that propagates an error ends there and takes no part
            # in the merge after End If.
            "raise-after-a-branch-that-raises": ('''Cleanup:
        If Flag Then
            On Error GoTo 0: Err.Raise 5
        End If
        Err.Raise 6
''', "Err.Raise"),
            # Do ... Loop While runs its body at least once.
            "raise-after-a-post-tested-loop": ('''Cleanup:
        On Error GoTo 0
        Do
            On Error GoTo Handler
        Loop While Index < 0
        Err.Raise 5
''', "Err.Raise"),
            "raise-after-a-bare-loop-left-by-exit-do": ('''Cleanup:
        Do
            If Flag Then Exit Do
        Loop
        Err.Raise 5
''', "Err.Raise"),
            # The path ends at Exit Sub inside the enclosing loop, but the
            # Exit Do it passed first still leaves the loop.
            "raise-after-an-enclosing-loop-left-before-the-path-ends": ('''        Do
Cleanup:
            If Flag Then Exit Do
            Exit Sub
        Loop
        Err.Raise 5
''', "Err.Raise"),
            # Nothing jumps to Unused, so the branch stays ended after Exit Sub.
            "raise-after-an-untargeted-label-in-an-ended-branch": ('''Cleanup:
        If Flag Then
            Exit Sub
Unused:
            On Error GoTo 0
        End If
        Err.Raise 5
''', "Err.Raise"),
            # A line number kept for Erl is no jump target either.
            "raise-after-an-erl-line-number-in-an-ended-branch": ('''Cleanup:
        If Flag Then
            Exit Sub
100         On Error GoTo 0
        End If
        Err.Raise 5
''', "Err.Raise"),
            # Under On Error Resume Next a raise falls through.
            "raise-after-a-raise-under-resume-next": (
                "Cleanup:\n        On Error Resume Next: Err.Raise 5: "
                "On Error GoTo Handler: Err.Raise 6\n", "Err.Raise"),
            "raise-after-a-branch-with-an-error-statement": ('''Cleanup:
        If Flag Then
            On Error GoTo 0: Error 5
        End If
        Err.Raise 6
''', "Err.Raise"),
        }
        passing = {
            "one-line-if-with-a-colon-disabled":
                "Cleanup:\n        If Flag Then: On Error GoTo 0: Err.Raise 5\n",
            # "If ... Then:" opens no block, so the Else stays a sibling.
            "one-line-if-with-a-colon-in-a-branch": '''        If Flag Then
Cleanup:
            On Error GoTo 0
            If Index = 1 Then: Debug.Print 0
        Else
            On Error GoTo Handler: Err.Raise 5
        End If
''',
            "subroutine-disables-the-handler": '''Cleanup:
        GoSub Disarm
        Err.Raise 5
        Exit Sub
Disarm:
        On Error GoTo 0
        Return
''',
            # Exit Do leaves the Do, not the For it stands in, with the handler off.
            "exit-do-inside-a-for": '''Cleanup:
        Do While Flag
            On Error GoTo 0
            For Index = 1 To 2
                If Index = 2 Then Exit Do
            Next Index
            On Error GoTo Handler
            Exit Sub
        Loop
        Err.Raise 5
''',
            "error-function": "Cleanup:\n        Debug.Print Error$(5), Error(5)\n",
            # A bare Do ... Loop without Exit Do never reaches the raise after it.
            "raise-after-a-bare-loop-without-exit-do": '''Cleanup:
        Do
            If Flag Then Exit Sub
            On Error GoTo 0: Err.Raise 6
        Loop
        Err.Raise 5
''',
            "raise-after-an-enclosing-bare-loop-without-exit-do": '''        Do
Cleanup:
            Do
                Exit Do
            Loop
            If Flag Then Exit Sub
        Loop
        Err.Raise 5
''',
            "raise-after-an-enclosing-loop-left-disabled": '''        Do
Cleanup:
            If Flag Then On Error GoTo 0: Exit Do
            Exit Sub
        Loop
        Err.Raise 5
''',
            # The Exit Do after Exit Sub is never reached: the path leaves the
            # procedure or loops forever, so nothing runs after the loop.
            "unreached-exit-of-an-enclosing-bare-loop": '''        Do
Cleanup:
            If Flag Then
                Exit Sub
                Exit Do
            End If
        Loop
        Err.Raise 5
''',
            # Documented false negative: an Exit Do reached only in a later pass.
            "exit-before-the-target-in-an-enclosing-bare-loop": '''        Do
            If Flag Then Exit Do
Cleanup:
            Debug.Print 0
        Loop
        Err.Raise 5
''',
            # Once the path ends, a label another jump reaches is not followed.
            "targeted-label-after-the-path-ends-in-a-loop": '''        If Index = 2 Then GoTo Again
        Do
Cleanup:
            Exit Sub
Again:
            On Error GoTo Handler
            Err.Raise 5
        Loop
''',
            # A Do While may run no pass, so the state before it also continues.
            "raise-after-a-pre-tested-loop": '''Cleanup:
        On Error GoTo 0
        Do While Flag
            On Error GoTo Handler
        Loop
        Err.Raise 5
''',
            # A jump may reach Disarm, so the state after End If is unknown.
            "raise-after-a-targeted-label-in-an-ended-branch": '''        If Index = 2 Then GoTo Disarm
Cleanup:
        If Flag Then
            Exit Sub
Disarm:
            On Error GoTo 0
        End If
        Err.Raise 5
''',
            # "If ... Then 100" jumps to line 100.
            "raise-after-an-implicitly-targeted-line-number": '''        If Index = 2 Then 100
Cleanup:
        If Flag Then
            Exit Sub
100         On Error GoTo 0
        End If
        Err.Raise 5
''',
            # "If Flag Then 100" jumps, so the raise after it on that line never runs.
            "implicit-line-number-jump-ends-the-branch":
                "Cleanup:\n        If Flag Then 100: Err.Raise 6\n        Exit Sub\n"
                "100     Exit Sub\n",
            "implicit-line-number-jumps-in-both-branches":
                "Cleanup:\n        If Flag Then 100 Else 200\n        Err.Raise 6\n"
                "100     Exit Sub\n200     Exit Sub\n",
            "raise-after-an-implicit-jump-disabled":
                "Cleanup:\n        If Flag Then 100\n        On Error GoTo 0\n        Err.Raise 6\n"
                "100     Exit Sub\n",
            # The subroutine branches before its Return, so the state is unknown.
            "branching-subroutine-disables-the-handler": '''Cleanup:
        GoSub Disarm
        Err.Raise 5
        Exit Sub
Disarm:
        If Flag Then Exit Sub
        On Error GoTo 0
        Return
''',
            "nested-subroutine-disables-the-handler": '''Cleanup:
        GoSub Outer
        Err.Raise 5
        Exit Sub
Outer:
        GoSub Disable
        Return
Disable:
        On Error GoTo 0
        Return
''',
            "raise-after-a-raise-under-resume-next-disabled":
                "Cleanup:\n        On Error Resume Next: Err.Raise 5: On Error GoTo 0: Err.Raise 6\n",
        }
        every = " [vba6-win32, vba7-win32, vba7-win64]"
        for name, (blocks, raised) in reported.items():
            with self.subTest(name=name):
                self.assertEqual([item.split(": ", 1)[1] for item in
                                  run_check(check_vba_jumps.check, enclosed(blocks))],
                                 [TRAPPED.replace("Err.Raise", raised, 1) + every])
        for name, blocks in passing.items():
            with self.subTest(name=name):
                self.assertEqual(run_check(check_vba_jumps.check, enclosed(blocks)), [])
        with self.subTest(name="line-numbers"):
            numbered = standalone(module('''Private Sub First()
        On Error GoTo 900
        Debug.Print 1 / 0
100     Err.Raise 5
        Exit Sub
900     Resume 100
End Sub
'''))
            self.assertEqual(run_check(check_vba_jumps.check, numbered),
                             ["tests/modules/TEST_Fixture.bas:7: Err.Raise after Resume 100 is "
                              "trapped again by handler '900' of procedure First; put On Error "
                              "GoTo 0 before the raise" + every])

    def test_reraise_is_checked_in_each_environment(self) -> None:
        findings = run_check(check_vba_jumps.check,
                             reraise("#If VBA7 Then\n        On Error GoTo 0\n#End If\n"))
        self.assertEqual(findings,
                         [f"tests/modules/TEST_Fixture.bas:13: {TRAPPED} [vba6-win32]"])

    def test_conditional_resume_starts_no_path(self) -> None:
        retry = '''Private Sub First()
        Dim Attempts As Long
Retry:
        On Error GoTo Handler
        If Attempts < 1 Then Err.Raise 5, "TEST_Fixture.First"
        Exit Sub
Handler:
        Attempts = Attempts + 1
        If Attempts < 3 Then Resume Retry
        Err.Raise Err.Number, "TEST_Fixture.First"
End Sub
'''
        block_retry = retry.replace("        If Attempts < 3 Then Resume Retry\n",
                                    "        If Attempts < 3 Then\n            Resume Retry\n"
                                    "        End If\n")
        # Documented false negative: a cleanup reached only under a condition.
        conditional_cleanup = RERAISE.replace(
            "        Resume Cleanup\n",
            "        If SavedNumber <> 0 Then\n            Resume Cleanup\n        End If\n")
        for name, body in {"bounded-retry": retry, "bounded-retry-in-a-block-if": block_retry,
                           "conditional-cleanup": conditional_cleanup}.items():
            with self.subTest(name=name):
                self.assertEqual(run_check(check_vba_jumps.check, standalone(module(body))), [])
        with self.subTest(name="unconditional-retry"):
            findings = run_check(check_vba_jumps.check, standalone(module(
                retry.replace("If Attempts < 3 Then Resume Retry", "Resume Retry"))))
            self.assertEqual([item.split(": ", 1)[1] for item in findings],
                             [TRAPPED.replace("Cleanup", "Retry")
                              + " [vba6-win32, vba7-win32, vba7-win64]"])

    def test_resume_without_a_label_starts_no_path(self) -> None:
        # Resume and Resume Next continue at the statement that failed, or
        # after it, which the check cannot locate; only Resume <label> starts a path.
        for resume in ("Resume", "Resume Next"):
            with self.subTest(resume=resume):
                files = reraise(fixture=RERAISE.replace("Resume Cleanup", resume))
                self.assertEqual(run_check(check_vba_jumps.check, files), [])


API_FACADE = '''Attribute VB_Name = "DEMO_Api"
Option Explicit
Public Const DEMO_LIMIT As Long = 7
Private Declare PtrSafe Function Tick Lib "kernel32" () As Long
Private mStatus As Long
Public Enum DemoMode
    dmFirst = 1
End Enum
Public Type DemoPair
    Left As Long
    Right As Long
End Type
Private mHidden As Long

Public Function DEMO_Echo( _
    ByVal Value As Long) _
    As Long
        DEMO_Echo = Value
End Function

Public Property Get DEMO_Current() As Long
        DEMO_Current = mStatus
End Property

Public Property Let DEMO_Current(ByVal Value As Long)
        mStatus = Value
End Property

Private Sub Helper()
End Sub
'''
API_ROWS = [
    ("Const", "DEMO_LIMIT", "Public Const DEMO_LIMIT As Long = 7"),
    ("Enum", "DemoMode", "Public Enum DemoMode | dmFirst = 1 | End Enum"),
    ("Type", "DemoPair", "Public Type DemoPair | Left As Long | Right As Long | End Type"),
    ("Function", "DEMO_Echo", "Public Function DEMO_Echo(ByVal Value As Long) As Long"),
    ("Property Get", "DEMO_Current", "Public Property Get DEMO_Current() As Long"),
    ("Property Let", "DEMO_Current", "Public Property Let DEMO_Current(ByVal Value As Long)"),
]


INTERNAL_ROWS = ("[project-public]\nCORE_Internal\tFunction\tInside\n"
                 "# SIG\tCORE_Internal\tFunction\tInside\tPublic Function Inside() As Long\n")


def api_inventory(rows: list[tuple[str, str, str]], extra: str = "",
                  internal: str = INTERNAL_ROWS) -> str:
    """The facade's ``rows`` and ``extra`` under [supported], then ``internal``."""
    lines = ["[supported]"]
    for kind, name, signature in rows:
        lines += [f"DEMO_Api\t{kind}\t{name}", f"# SIG\tDEMO_Api\t{kind}\t{name}\t{signature}"]
    return "\n".join(lines) + "\n" + extra + internal


def api_files(facade: str = API_FACADE, inventory: str | None = None,
              **others: str) -> dict[str, str | bytes]:
    files: dict[str, str | bytes] = {
        "src/modules/DEMO_Api.bas": facade,
        "src/core/CORE_Internal.bas": 'Attribute VB_Name = "CORE_Internal"\nOption Explicit\n'
                                      "Option Private Module\nPublic Function Inside() As Long\n"
                                      "End Function\n",
        "tests/modules/TEST_Api.bas": 'Attribute VB_Name = "TEST_Api"\nOption Explicit\n'
                                      "Option Private Module\nPublic Sub RunAll()\nEnd Sub\n",
        "docs/PUBLIC_API.txt": api_inventory(API_ROWS) if inventory is None else inventory,
    }
    files.update({f"src/modules/{name}.bas": text for name, text in others.items()})
    return files


MODERN = 'Public Const DEMO_HOST As String = "vba7"'
LEGACY = 'Public Const DEMO_HOST As String = "vba6"'
HOST_ROW = "DEMO_Api\tConst\tDEMO_HOST\n"
MODERN_ROW = f"# SIG\tDEMO_Api\tConst\tDEMO_HOST\t{MODERN}\n"
LEGACY_ROW = f"# SIG\tDEMO_Api\tConst\tDEMO_HOST\t{LEGACY}\n"
CONDITIONAL = API_FACADE + f"\n#If VBA7 Then\n{MODERN}\n#Else\n{LEGACY}\n#End If\n"


class PublicApiTest(unittest.TestCase):
    def findings(self, files: Mapping[str, str | bytes]) -> list[str]:
        return run_check(check_vba_public_api.check, files)

    def assertFinding(self, files: Mapping[str, str | bytes], text: str) -> None:
        findings = self.findings(files)
        self.assertTrue(any(text in item for item in findings), findings)

    def test_valid_fixtures_pass(self) -> None:
        self.assertEqual(self.findings(VALID), [])
        self.assertEqual(self.findings(api_files()), [])

    def test_repository_inventory_matches_its_source(self) -> None:
        # The inventory and the source of this repository together, so the test holds
        # whether or not the inventory lists a declaration yet.
        findings, _, _, _ = check_vba_public_api.check(ROOT)
        self.assertEqual([str(finding) for finding in findings], [])
        # Each row names a declaration of its tier, and each declaration has a row.
        # One declaration may be compiled in several #If branches under one row.
        source = {(check_vba_public_api.tier_of(path), declaration.key.casefold())
                  for path in tracked_components(ROOT)
                  if check_vba_public_api.tier_of(path)
                  for declaration in check_vba_public_api.scan_component(
                      path, read_component(ROOT, path))[0]}
        inventory, errors = check_vba_public_api.read_inventory(ROOT)
        self.assertEqual(errors, [])
        self.assertTrue(inventory.tiered)
        listed = {(tier, folded) for folded, (tier, _) in inventory.rows.items()}
        self.assertEqual(listed, source)

    def test_implicit_visibility_is_rejected_everywhere(self) -> None:
        self.assertFinding(api_files(API_FACADE.replace("Public Function DEMO_Echo",
                                                        "Function DEMO_Echo")),
                           "implicit public procedure: DEMO_Echo")
        hidden = 'Attribute VB_Name = "CORE_Implicit"\nOption Explicit\nOption Private Module\n' \
                 "Sub Hidden()\nEnd Sub\nEnum Mode\n    mA = 1\nEnd Enum\n"
        findings = self.findings({"src/core/CORE_Implicit.bas": hidden,
                                  "docs/PUBLIC_API.txt": "# None\n"})
        self.assertTrue(any("implicit public procedure: Hidden" in f for f in findings), findings)
        self.assertTrue(any("implicit public module-level declaration: Enum Mode" in f
                            for f in findings), findings)

    def test_module_level_dim_and_const_state_their_visibility(self) -> None:
        unstated = 'Attribute VB_Name = "CORE_Unstated"\nOption Explicit\nOption Private Module\n' \
                   "Dim mCount As Long\nConst LIMIT As Long = 5\n" \
                   "Private Sub Work()\n        Dim Local As Long\n        Const STEP As Long = 1\n" \
                   "End Sub\n"
        findings = self.findings({"src/core/CORE_Unstated.bas": unstated,
                                  "docs/PUBLIC_API.txt": "# None\n"})
        for text in ("Dim mCount As Long", "Const LIMIT As Long = 5"):
            with self.subTest(text=text):
                self.assertTrue(any(f"module-level declaration without Private or Public: {text}"
                                    in f for f in findings), findings)
        # Declarations inside a procedure need no visibility.
        self.assertEqual(len(findings), 2, findings)

    def test_one_identifier_per_public_statement(self) -> None:
        self.assertFinding(api_files(API_FACADE.replace("Private mHidden As Long",
                                                        "Public gFirst As Long, gOther")),
                           "Public variable statement must declare one identifier")
        self.assertFinding(api_files(API_FACADE.replace("= 7", "= 7, DEMO_MORE As Long = 8")),
                           "Public Const statement must declare one identifier")

    def test_inventory_must_list_every_declaration(self) -> None:
        self.assertFinding(api_files(inventory=api_inventory(API_ROWS[1:])),
                           "public declaration not listed: DEMO_Api\tConst\tDEMO_LIMIT")
        rows = api_inventory(API_ROWS).replace("DEMO_Api\tFunction\tDEMO_Echo\n", "")
        self.assertFinding(api_files(inventory=rows), "not listed: DEMO_Api\tFunction\tDEMO_Echo")

    def test_signature_records(self) -> None:
        rows = api_inventory(API_ROWS).replace("(ByVal Value As Long) As Long", "() As Long")
        self.assertFinding(api_files(inventory=rows), "signature mismatch for DEMO_Api\tFunction")
        missing = "\n".join(line for line in api_inventory(API_ROWS).splitlines()
                            if "SIG\tDEMO_Api\tFunction" not in line) + "\n"
        self.assertFinding(api_files(inventory=missing), "signature record missing")
        duplicate = api_inventory(API_ROWS, api_inventory(API_ROWS[:1]).splitlines()[2] + "\n")
        self.assertFinding(api_files(inventory=duplicate), "duplicate signature record")
        self.assertFinding(api_files(inventory=api_inventory(API_ROWS, "# SIG\tDEMO_Api\n")),
                           "malformed # SIG record")

    def test_stale_and_malformed_rows(self) -> None:
        stale = api_inventory(API_ROWS, "DEMO_Api\tFunction\tGone\n"
                                        "# SIG\tDEMO_Api\tFunction\tGone\tPublic Function Gone()\n")
        findings = self.findings(api_files(inventory=stale))
        self.assertTrue(any("listed declaration not in src/modules/" in f for f in findings))
        self.assertTrue(any("stale signature record" in f for f in findings))
        self.assertFinding(api_files(inventory=api_inventory(API_ROWS, "DEMO_Api\tFunction\n")),
                           "needs component, kind and name")
        self.assertFinding(api_files(inventory=api_inventory(API_ROWS, "DEMO_Api\tConst\tdemo_limit\n")),
                           "duplicate inventory row")

    def test_only_the_facade_is_supported_api(self) -> None:
        # CORE_Internal's rows moved from [project-public] to [supported].
        rows = api_inventory(API_ROWS, INTERNAL_ROWS.split("\n", 1)[1], internal="")
        findings = self.findings(api_files(inventory=rows))
        for text in ("public declaration listed under [supported], not [project-public]: "
                     "CORE_Internal\tFunction\tInside",
                     "signature record listed under [supported], not [project-public]: "
                     "CORE_Internal\tFunction\tInside"):
            with self.subTest(text=text):
                self.assertTrue(any(text in item for item in findings), findings)
        self.assertEqual(len(findings), 2, findings)

    def test_missing_inventory(self) -> None:
        files = {k: v for k, v in api_files().items() if k != "docs/PUBLIC_API.txt"}
        self.assertFinding(files, "public API inventory is missing")

    def test_collisions(self) -> None:
        twice = API_FACADE + "\nPublic Function DEMO_Echo(ByVal Value As String) As String\n" \
                             "End Function\n"
        self.assertFinding(api_files(twice), "public declaration appears more than once")
        other = 'Attribute VB_Name = "DEMO_Other"\nOption Explicit\n' \
                "Public Function DEMO_Echo(ByVal Value As Long) As Long\nEnd Function\n"
        self.assertFinding(api_files(DEMO_Other=other), "public name 'DEMO_Echo' collides with "
                                                       "DEMO_Api.DEMO_Echo (Function)")
        # Every standard module shares the project's namespace, not only the facade.
        core = 'Attribute VB_Name = "CORE_Twin"\nOption Explicit\nOption Private Module\n' \
               "Public Function Inside() As Long\nEnd Function\n"
        self.assertFinding({**api_files(), "src/core/CORE_Twin.bas": core},
                           "public name 'Inside' collides with CORE_Internal.Inside (Function)")
        echo = core.replace("Inside", "DEMO_Echo")
        self.assertFinding({**api_files(), "src/core/CORE_Twin.bas": echo},
                           "collides with")

    def test_no_public_globals_or_declares_in_standard_modules(self) -> None:
        self.assertFinding(api_files(API_FACADE.replace("Private mHidden As Long",
                                                        "Public gState As Long")),
                           "public variable gState in a standard module")
        self.assertFinding(api_files(API_FACADE.replace("Private mHidden As Long",
                                                        "Global gState As Long")),
                           "public variable gState in a standard module")
        self.assertFinding(api_files(API_FACADE.replace("Private Declare", "Public Declare")),
                           "public Declare Tick; declare native procedures Private")
        core = 'Attribute VB_Name = "CORE_State"\nOption Explicit\nOption Private Module\n' \
               "Public gCounter As Long\n"
        self.assertFinding({"src/core/CORE_State.bas": core, "docs/PUBLIC_API.txt": "# None\n"},
                           "public variable gCounter in a standard module")
        # A class may expose a public field; it is listed under [project-public].
        field = CLASS.replace("Private mAmount As Double", "Public Label As String")
        listed = ("[project-public]\n" + CLASS_ROWS + "CCalculationResult\tVariable\tLabel\n"
                  "# SIG\tCCalculationResult\tVariable\tLabel\tPublic Label As String\n")
        self.assertEqual(self.findings({"src/classes/CCalculationResult.cls": field,
                                        "docs/PUBLIC_API.txt": listed}), [])

    def test_conditional_variants(self) -> None:
        variants = api_inventory(API_ROWS, HOST_ROW + MODERN_ROW + LEGACY_ROW)
        self.assertEqual(self.findings(api_files(CONDITIONAL, variants)), [])
        self.assertFinding(api_files(CONDITIONAL, api_inventory(API_ROWS, HOST_ROW + MODERN_ROW)),
                           "signature mismatch")
        self.assertFinding(api_files(API_FACADE + "\n" + MODERN + "\n", variants),
                           "signature mismatch")
        self.assertFinding(api_files(CONDITIONAL + MODERN + "\n", variants),
                           "appears more than once")
        dead = API_FACADE + f"\n#If False Then\n{MODERN.replace('HOST', 'DEAD')}\n#End If\n"
        self.assertEqual(self.findings(api_files(dead)), [])
        self.assertFinding(api_files(CONDITIONAL.replace("#If VBA7", "#If Unknown"), variants),
                           "indeterminate #If")

    def test_exclusive_names_across_modules(self) -> None:
        modern_only = API_FACADE + f"\n#If VBA7 Then\n{MODERN}\n#End If\n"
        legacy = f'Attribute VB_Name = "DEMO_Legacy"\nOption Explicit\n#If Not VBA7 Then\n{LEGACY}\n' \
                 "#End If\n"
        rows = api_inventory(API_ROWS, HOST_ROW + MODERN_ROW + "DEMO_Legacy\tConst\tDEMO_HOST\n"
                             + LEGACY_ROW.replace("DEMO_Api", "DEMO_Legacy"))
        self.assertEqual(self.findings(api_files(modern_only, rows, DEMO_Legacy=legacy)), [])
        self.assertFinding(api_files(API_FACADE + "\n" + MODERN + "\n", rows, DEMO_Legacy=legacy),
                           "collides")

    def test_signatures_keep_string_literals(self) -> None:
        facade = API_FACADE.replace("Private Sub Helper()", "Public Function DEMO_Pad( _\n"
                                    '    Optional ByVal Text As String = "A  ( B )") _\n'
                                    "    As String\nEnd Function\n\nPrivate Sub Helper()")
        exact = api_inventory([*API_ROWS, ("Function", "DEMO_Pad", 'Public Function DEMO_Pad('
                                           'Optional ByVal Text As String = "A  ( B )") '
                                           "As String")])
        self.assertEqual(self.findings(api_files(facade, exact)), [])
        collapsed = exact.replace('"A  ( B )"', '"A (B)"')
        self.assertFinding(api_files(facade, collapsed), "signature mismatch for DEMO_Api\tFunction"
                                                         "\tDEMO_Pad")

    def test_declare_inside_a_literal_is_text(self) -> None:
        facade = API_FACADE.replace("Private Sub Helper()", "Public Function DEMO_Say("
                                    'Optional ByVal Text As String = " declare ") As String\n'
                                    "End Function\n\nPrivate Sub Helper()")
        rows = api_inventory([*API_ROWS, ("Function", "DEMO_Say", "Public Function DEMO_Say("
                                          'Optional ByVal Text As String = " declare ") As String')])
        self.assertEqual(self.findings(api_files(facade, rows)), [])

    def test_colon_separated_and_commented_statements(self) -> None:
        facade = API_FACADE.replace("Private mHidden As Long",
                                    "Private mHidden As Long: Public gExtra As Long ' note: x")
        self.assertFinding(api_files(facade), "public variable gExtra in a standard module")
        dated = API_FACADE.replace("Private mHidden As Long",
                                   "Private Const START_DATE As Date = #1/2/2026#: Rem : Public x")
        self.assertEqual(self.findings(api_files(dated)), [])

    def test_project_public_tier_lists_every_other_public_declaration(self) -> None:
        self.assertEqual(self.findings(project_files()), [])
        with repository(project_files()) as root:
            _, _, supported, project = check_vba_public_api.check(root)
        self.assertEqual((supported, project), (len(API_ROWS), 1 + len(PROJECT_ROWS)))
        for component, kind, name, _ in PROJECT_ROWS:
            key = f"{component}\t{kind}\t{name}"
            with self.subTest(key=key):
                rows = project_inventory().replace(f"{key}\n", "")
                self.assertFinding(project_files(rows),
                                   f"public declaration not listed: {key}; list it under "
                                   "[project-public]")
        self.assertFinding(project_files(project_inventory().replace(
            "Seam(ByVal Value As Long) As Long", "Seam() As Long")),
            "signature mismatch for CORE_Fixture\tFunction\tSeam")
        self.assertFinding(project_files(project_inventory().replace(
            "# SIG\tCFixture\tSub\tAdd\tPublic Sub Add(ByVal Amount As Double)\n", "")),
            "signature record missing: CFixture\tSub\tAdd")
        stale = project_inventory() + "CORE_Fixture\tSub\tGone\n" \
                                      "# SIG\tCORE_Fixture\tSub\tGone\tPublic Sub Gone()\n"
        findings = self.findings(project_files(stale))
        self.assertIn("docs/PUBLIC_API.txt: listed declaration not in src/ outside "
                      "src/modules/: CORE_Fixture\tSub\tGone", findings)
        self.assertIn("docs/PUBLIC_API.txt: stale signature record: core_fixture\tsub\tgone",
                      findings)
        self.assertEqual(len(findings), 2, findings)

    def test_tests_and_examples_belong_to_neither_tier(self) -> None:
        listed = "TEST_Api\tSub\tRunAll\n# SIG\tTEST_Api\tSub\tRunAll\tPublic Sub RunAll()\n"
        self.assertFinding(project_files(project_inventory() + listed),
                           "listed declaration not in src/ outside src/modules/: TEST_Api")
        self.assertFinding(project_files(api_inventory(API_ROWS, listed,
                                                       INTERNAL_ROWS + rows_text(PROJECT_ROWS))),
                           "listed declaration not in src/modules/: TEST_Api")

    def test_declaration_listed_in_the_wrong_tier(self) -> None:
        limit = "DEMO_Api\tConst\tDEMO_LIMIT"
        moved = api_inventory(API_ROWS[1:], internal=INTERNAL_ROWS + rows_text(PROJECT_ROWS)
                              + rows_text([("DEMO_Api", *API_ROWS[0])]))
        findings = self.findings(project_files(moved))
        self.assertEqual(sorted(findings), sorted([
            f"docs/PUBLIC_API.txt: public declaration listed under [project-public], "
            f"not [supported]: {limit}",
            f"docs/PUBLIC_API.txt: signature record listed under [project-public], "
            f"not [supported]: {limit}"]))
        # A class member in [supported]: only its row is in the wrong tier.
        add = "CFixture\tSub\tAdd"
        rows = project_inventory().replace(f"{add}\n", "").replace("[project-public]",
                                                                   f"{add}\n[project-public]")
        self.assertEqual(self.findings(project_files(rows)), [
            f"docs/PUBLIC_API.txt: public declaration listed under [supported], "
            f"not [project-public]: {add}"])

    def test_legacy_inventory_without_tier_headings(self) -> None:
        legacy = "# Supported VBA declarations\n" + rows_text(
            [("DEMO_Api", *row) for row in API_ROWS])
        facade_only = {k: v for k, v in api_files(inventory=legacy).items()
                       if k != "src/core/CORE_Internal.bas"}
        self.assertEqual(self.findings(facade_only), [])
        # With a public declaration elsewhere in src/, a legacy inventory fails.
        findings = self.findings(api_files(inventory=legacy))
        self.assertIn("docs/PUBLIC_API.txt: public declaration not listed: "
                      "CORE_Internal\tFunction\tInside; list it under [project-public]", findings)
        self.assertTrue(any("has no tier heading" in item for item in findings), findings)
        # A heading inside a comment is a comment: the rows after it stay [supported].
        commented = legacy + "# " + INTERNAL_ROWS
        self.assertFinding(api_files(inventory=commented),
                           "listed under [supported], not [project-public]: CORE_Internal")

    def test_tier_headings(self) -> None:
        empty = {"tests/modules/TEST_Api.bas": api_files()["tests/modules/TEST_Api.bas"],
                 "docs/PUBLIC_API.txt": "# Nothing yet\n[supported]\n\n[project-public]\n"}
        self.assertEqual(self.findings(empty), [])
        unknown = api_inventory(API_ROWS, internal=INTERNAL_ROWS.replace("[project-public]",
                                                                         "[internal]"))
        findings = self.findings(api_files(inventory=unknown))
        self.assertTrue(any("unknown tier heading [internal]" in item for item in findings),
                        findings)
        self.assertTrue(any("inventory entry outside a [supported] or [project-public] tier"
                            in item for item in findings), findings)
        self.assertFinding(api_files(inventory=api_inventory(API_ROWS, "[supported]\n")),
                           "tier heading appears more than once: [supported]")
        early = "DEMO_Api\tConst\tDEMO_EARLY\n" + api_inventory(API_ROWS)
        self.assertFinding(api_files(inventory=early),
                           "docs/PUBLIC_API.txt:1: inventory entry outside a [supported] or "
                           "[project-public] tier")

    def test_conditional_variants_in_project_public(self) -> None:
        # A class member whose header differs between #If branches: one row, one
        # # SIG record per distinct signature.
        variants = rows_text([("CFixture", "Sub", "Attach",
                               "Public Sub Attach(ByVal Handle As LongPtr)")]) + \
            "# SIG\tCFixture\tSub\tAttach\tPublic Sub Attach(ByVal Handle As Long)\n"
        source = project_sources()[CLASS_PATH]
        assert isinstance(source, str)
        conditional = {CLASS_PATH: source + "\n#If VBA7 Then\n"
                       "Public Sub Attach(ByVal Handle As LongPtr)\n#Else\n"
                       "Public Sub Attach(ByVal Handle As Long)\n#End If\nEnd Sub\n"}
        self.assertEqual(self.findings(project_files(project_inventory() + variants,
                                                     conditional)), [])
        one = variants.rsplit("# SIG", 1)[0]
        self.assertFinding(project_files(project_inventory() + one, conditional),
                           "signature mismatch for CFixture\tSub\tAttach")
        # A core constant compiled in one branch only, and its row.
        core = project_sources()[CORE_PATH]
        assert isinstance(core, str)
        modern = {CORE_PATH: core + f"\n#If VBA7 Then\n{MODERN.replace('DEMO_', 'CORE_')}\n"
                                    "#End If\n"}
        host = rows_text([("CORE_Fixture", "Const", "CORE_HOST",
                           MODERN.replace("DEMO_", "CORE_"))])
        self.assertEqual(self.findings(project_files(project_inventory() + host, modern)), [])
        self.assertFinding(project_files(None, modern),
                           "not listed: CORE_Fixture\tConst\tCORE_HOST; list it under "
                           "[project-public]")


def rows_text(rows: list[tuple[str, str, str, str]]) -> str:
    """Inventory rows with their # SIG records."""
    return "".join(f"{component}\t{kind}\t{name}\n# SIG\t{component}\t{kind}\t{name}\t"
                   f"{signature}\n" for component, kind, name, signature in rows)


PROJECT_ROWS = [
    ("CORE_Fixture", "Const", "CORE_STEP", "Public Const CORE_STEP As Long = 1"),
    ("CORE_Fixture", "Enum", "CoreMode", "Public Enum CoreMode | cmFirst = 1 | End Enum"),
    ("CORE_Fixture", "Function", "Seam", "Public Function Seam(ByVal Value As Long) As Long"),
    ("CFixture", "Event", "Changed", "Public Event Changed(ByVal Total As Double)"),
    ("CFixture", "Variable", "Label", "Public Label As String"),
    ("CFixture", "Property Get", "Total", "Public Property Get Total() As Double"),
    ("CFixture", "Sub", "Add", "Public Sub Add(ByVal Amount As Double)"),
    ("ThisWorkbook", "Sub", "Refresh", "Public Sub Refresh()"),
    ("frmRunOptions", "Property Get", "Confirmed",
     "Public Property Get Confirmed() As Boolean"),
]


def project_inventory() -> str:
    """Both tiers listing every public declaration of project_files()."""
    return api_inventory(API_ROWS, internal=INTERNAL_ROWS + rows_text(PROJECT_ROWS))


def project_sources() -> dict[str, str | bytes]:
    """A public declaration in every home of src/, and in a test and an example."""
    return {
        CORE_PATH: core_module("Public Const CORE_STEP As Long = 1\nPublic Enum CoreMode\n"
                               "    cmFirst = 1\nEnd Enum\n\nPublic Function Seam( _\n"
                               "    ByVal Value As Long) _\n    As Long\n        Seam = Value\n"
                               "End Function\n"),
        CLASS_PATH: class_module("Public Event Changed(ByVal Total As Double)\n"
                                 "Public Label As String\nPrivate mTotal As Double\n\n\n"
                                 "Public Property Get Total() As Double\n"
                                 "        Total = mTotal\nEnd Property\n\n\n"
                                 "Public Sub Add(ByVal Amount As Double)\n"
                                 "        mTotal = mTotal + Amount\nEnd Sub\n"),
        "src/workbook/ThisWorkbook.cls": WORKBOOK + "\n\nPublic Sub Refresh()\nEnd Sub\n",
        "src/forms/frmRunOptions.frm": FORM + "\n\nPublic Property Get Confirmed() As Boolean\n"
                                              "End Property\n",
        "src/forms/frmRunOptions.frx": b"\x00\x01\x02\x03",
        EXAMPLE_PATH: example_module("Public Sub Run()\nEnd Sub\n"),
    }


def project_files(inventory: str | None = None,
                  sources: Mapping[str, str] | None = None) -> dict[str, str | bytes]:
    """api_files() plus project_sources(), with ``sources`` replacing them by path."""
    files = api_files(inventory=project_inventory() if inventory is None else inventory)
    files.update(project_sources())
    files.update(sources or {})
    return files


def core_module(body: str) -> str:
    return f'Attribute VB_Name = "CORE_Fixture"\nOption Explicit\nOption Private Module\n{body}'


def class_module(body: str) -> str:
    return CLASS.split("Option Explicit")[0].replace("CCalculationResult", "CFixture") + \
        f"Option Explicit\n{body}"


def example_module(body: str) -> str:
    return f'Attribute VB_Name = "EX_Fixture"\nOption Explicit\n{body}'


CORE_PATH, CLASS_PATH = "src/core/CORE_Fixture.bas", "src/classes/CFixture.cls"
EXAMPLE_PATH = "examples/modules/EX_Fixture.bas"

DEPENDENCY_VALID = {
    "pure-core": (CORE_PATH, core_module('''Private Const LIMIT As Long = 3

Private Enum Slot
    [_First] = 0
    slA = 0
End Enum

Private Type Pair
    Low As Double
    High As Double
End Type

Public Function Count(ByVal Items As Object) As Long
        ' Application, MsgBox and [A1] in a comment are not code.
        Dim Seen As Object
        Set Seen = CreateObject("Scripting.Dictionary")
        Set Seen = CreateObject( "Scripting.Dictionary" )
        CreateObject "Scripting.Dictionary"
        Debug.Print "Range and MsgBox inside a string"
        Count = Items.Count + Items.Rows + LIMIT
End Function
''')),
    "alternative-procedure-headers": (CORE_PATH, core_module('''#If Win64 Then
Private Sub Work(ByVal Value As LongPtr)
#Else
Private Sub Work(ByVal Value As Long)
#End If
        Dim Local As Long
        Local = 1
End Sub
''')),
    "class-instance-state": (CLASS_PATH, class_module('''Private mTotal As Double


Public Property Get Total() As Double
        Total = mTotal
End Property


Public Property Let Total(ByVal Value As Double)
        mTotal = Value
End Property
''')),
    "example-uses-facade": (EXAMPLE_PATH, example_module('''Public Sub Run()
        Debug.Print DEMO_Functions.DEMO_IsPositive(1)
End Sub
''')),
    "object-calls-in-strings-are-text": (CORE_PATH, core_module('''Public Sub Work()
        Debug.Print "GetObject(", "CreateObject(""Excel.Application"")"
End Sub
''')),
    "lower-case-member-prefix-is-an-identifier": (CLASS_PATH, class_module('''Private m_Total As Double


Public Sub Add(ByVal Amount As Double)
        m_Total = m_Total + Amount
End Sub
''')),
    "single-line-if-without-io": (CORE_PATH, core_module('''Public Sub Work(ByVal Flag As Boolean)
        If Flag Then Debug.Print "Kill" Else Debug.Print "Open"
End Sub
''')),
    "ordinary-names-and-early-bound-dictionary": (CORE_PATH, core_module('''Public Function Pick(ByVal Inputs As Long, ByVal Location As Long) As String
        Dim Seen As Scripting.Dictionary
        Set Seen = New Scripting.Dictionary
100     Debug.Print Date, Inputs + Location
        Pick = VBA.Strings.Left("ab", 1)
        Call Helper(1)
End Function

Private Sub Helper(ByVal Value As Long)
End Sub
''')),
    "members-named-like-statements": (CORE_PATH, core_module('''Private Type Quote
    Name As String
    Close As Double
    Lock As Boolean
    Range(1 To 2) As Double
    [Open] As Double
End Type

Private Enum Field
    Reset = 1
    Load = 2
    Dir = 3
End Enum
''')),
    "continued-comment": (CORE_PATH, core_module('''Public Sub Work()
        ' A comment that continues _
          MsgBox 1
End Sub
''')),
    "tests-and-facade-are-not-restricted": ("tests/modules/TEST_Fixture.bas", module('''
Public Sub RunAll()
        Debug.Print Application.Name, CORE_Validation.IsPositive(1)
End Sub
''')),
}


def in_procedure(body: str) -> str:
    return core_module(f"Public Sub Work()\n{body}\nEnd Sub\n")


DEPENDENCY_INVALID = {
    "application": (CORE_PATH, in_procedure("        Application.Run \"Macro\""),
                    "Application reaches the Excel host"),
    "msgbox": (CORE_PATH, in_procedure("        MsgBox \"x\""), "MsgBox reaches the Excel host"),
    "worksheets": (CORE_PATH, in_procedure("        Debug.Print Worksheets.Count"),
                   "Worksheets reaches the Excel host"),
    "excel-type": (CORE_PATH, in_procedure("        Dim Target As Worksheet"),
                   "As Worksheet is an Excel or UI type"),
    "qualified-excel-type": (CORE_PATH, in_procedure("        Dim Target As Excel.Range"),
                             "Excel reaches the Excel host"),
    "bracket-evaluate": (CORE_PATH, in_procedure("        Debug.Print [A1]"),
                         "evaluated by Excel"),
    "open-file": (CORE_PATH, in_procedure('        Open "data.txt" For Input As #1'),
                  "file, registry or process access"),
    "print-to-file": (CORE_PATH, in_procedure("        Print #1, 1"),
                      "file, registry or process access"),
    "environment": (CORE_PATH, in_procedure('        Debug.Print Environ("TEMP")'),
                    "file, registry or process access"),
    "shell": (CORE_PATH, in_procedure('        Debug.Print Shell("cmd")'),
              "file, registry or process access"),
    "call-by-name": (CORE_PATH, in_procedure('        Debug.Print CallByName(Me, "X", VbGet)'),
                     "CallByName: a dynamic call"),
    "evaluate": (CORE_PATH, in_procedure('        Debug.Print Evaluate("1+1")'),
                 "Evaluate: a dynamic call"),
    "create-host-object": (CORE_PATH,
                           in_procedure('        Debug.Print CreateObject("Word.Application")'),
                           "CreateObject creates an object no check can see"),
    "create-object-from-variable": (CORE_PATH,
                                    in_procedure("        Debug.Print CreateObject(ProgId)"),
                                    "CreateObject creates an object no check can see"),
    "create-object-concatenated": (CORE_PATH, in_procedure(
        '        Set Seen = CreateObject("Scripting.Dictionary" & Suffix)'),
        "CreateObject creates an object no check can see"),
    "create-object-statement-concatenated": (CORE_PATH, in_procedure(
        '        CreateObject "Scripting.Dictionary" & Suffix'),
        "CreateObject creates an object no check can see"),
    "create-object-on-a-server": (CORE_PATH, in_procedure(
        '        Set Seen = CreateObject("Scripting.Dictionary", Server)'),
        "CreateObject creates an object no check can see"),
    "create-object-statement": (CORE_PATH, in_procedure('        CreateObject "Word.Application"'),
                                "CreateObject creates an object no check can see"),
    "get-object-statement": (CORE_PATH, in_procedure('        GetObject , "Excel.Application"'),
                             "GetObject creates an object no check can see"),
    "call-create-object": (CLASS_PATH, class_module('Public Sub Go()\n        Call CreateObject("Word.Application")\n'
                                                    'End Sub\n'),
                           "CreateObject creates an object no check can see"),
    "qualified-create-object-statement": (CORE_PATH, in_procedure(
        '        VBA.Interaction.CreateObject "Word.Application"'),
        "CreateObject creates an object no check can see"),
    "create-object-statement-from-variable": (CORE_PATH, in_procedure("        CreateObject ProgId"),
                                              "CreateObject creates an object no check can see"),
    "get-object": (CORE_PATH, in_procedure('        Debug.Print GetObject(, "Excel.Sheet")'),
                   "GetObject creates an object no check can see"),
    "core-calls-facade": (CORE_PATH, in_procedure("        Debug.Print DEMO_Functions.DEMO_IsPositive(1)"),
                          "DEMO_Functions names a facade, host, test or example component"),
    "class-calls-adapter": (CLASS_PATH, class_module("Public Sub Go()\n        M_Run.Start\nEnd Sub\n"),
                            "M_Run names a facade, host, test or example component"),
    "class-reaches-host": (CLASS_PATH, class_module("Public Sub Go()\n        Range(\"A1\").Value = 1\n"
                                                    "End Sub\n"),
                           "Range reaches the Excel host"),
    "module-level-private": (CORE_PATH, core_module("Private mCount As Long\n"),
                             "module-level variable; core holds no ambient mutable run state"),
    "module-level-dim": (CORE_PATH, core_module("Dim mCount As Long\n"), "module-level variable"),
    "static-variable": (CORE_PATH, in_procedure("        Static Calls As Long"),
                        "Static variable; core holds no run state"),
    "static-procedure": (CORE_PATH, core_module("Private Static Function Next1() As Long\n"
                                                "End Function\n"),
                         "Static procedure Next1 keeps state between calls"),
    "module-variable-between-alternative-headers": (CORE_PATH, core_module('''#If Win64 Then
Private Sub Work()
#Else
Private mCounter As Long
Private Sub Work()
#End If
        Debug.Print 1
End Sub
'''), "module-level variable; core holds no ambient mutable run state"),
    "static-in-one-branch-procedure": (CORE_PATH, core_module('''#If VBA7 Then
Private Sub Work()
        Static Calls As Long
#Else
Private Sub Work()
#End If
        Debug.Print 1
End Sub
'''), "Static variable; core holds no run state"),
    "inactive-branch-is-read": (CORE_PATH, in_procedure("#If VBA7 Then\n        Debug.Print 1\n#Else\n"
                                                        "        MsgBox \"legacy\"\n#End If"),
                                "MsgBox reaches the Excel host"),
    "core-calls-facade-in-lower-case": (CORE_PATH, in_procedure("        Debug.Print demo_functions.Calculate(1)"),
                                        "demo_functions names a facade, host, test or example component"),
    "single-line-if-kill": (CORE_PATH, in_procedure('        If Flag Then Kill "f.txt"'),
                            "file, registry or process access"),
    "single-line-else-kill": (CORE_PATH, in_procedure('        If Flag Then Debug.Print 1 Else Kill "f.txt"'),
                              "file, registry or process access"),
    "nested-single-line-if": (CORE_PATH, in_procedure('        If A Then If B Then Open "f" For Input As #1'),
                              "file, registry or process access"),
    "block-else-on-same-line": (CORE_PATH, in_procedure('        If A Then\n        Else Kill "f.txt"\n        End If'),
                                "file, registry or process access"),
    "file-length": (CORE_PATH, in_procedure('        Debug.Print FileLen("data.txt")'),
                    "file, registry or process access"),
    "file-date": (CORE_PATH, in_procedure('        Debug.Print FileDateTime("data.txt")'),
                  "file, registry or process access"),
    "get-attributes": (CORE_PATH, in_procedure('        Debug.Print GetAttr("data.txt")'),
                       "file, registry or process access"),
    "set-attributes": (CORE_PATH, in_procedure('        SetAttr "data.txt", vbHidden'),
                       "file, registry or process access"),
    "end-of-file": (CORE_PATH, in_procedure("        Debug.Print EOF(1)"),
                    "file, registry or process access"),
    "command-line": (CORE_PATH, in_procedure("        Debug.Print Command$"),
                     "file, registry or process access"),
    "set-system-date": (CORE_PATH, in_procedure("        Date = #1/2/2026#"),
                        "file, registry or process access"),
    "unload-form": (CLASS_PATH, class_module("Public Sub Go()\n        Unload Me\nEnd Sub\n"),
                    "file, registry or process access"),
    "numbered-kill": (CORE_PATH, in_procedure('100     Kill "data.txt"'),
                      "file, registry or process access"),
    "numbered-static": (CORE_PATH, in_procedure("100     Static Calls As Long"),
                        "Static variable; core holds no run state"),
    "numbered-single-line-if": (CORE_PATH, in_procedure('100     If Flag Then Kill "data.txt"'),
                                "file, registry or process access"),
    "qualified-kill": (CORE_PATH, in_procedure('        VBA.FileSystem.Kill "data.txt"'),
                       "file, registry or process access"),
    "vba-qualified-kill": (CORE_PATH, in_procedure('        VBA.Kill "data.txt"'),
                           "file, registry or process access"),
    "qualified-shell": (CORE_PATH, in_procedure('        Debug.Print VBA.Interaction.Shell("cmd")'),
                        "file, registry or process access"),
    "qualified-msgbox": (CORE_PATH, in_procedure('        VBA.Interaction.MsgBox "x"'),
                         "MsgBox reaches the Excel host"),
    "call-save-setting": (CORE_PATH, in_procedure('        Call SaveSetting("App", "Section", "Key", "1")'),
                          "file, registry or process access"),
    "call-qualified-save-setting": (CORE_PATH, in_procedure(
        '        Call VBA.Interaction.SaveSetting("App", "Section", "Key", "1")'),
        "file, registry or process access"),
    "call-kill": (CLASS_PATH, class_module('Public Sub Go()\n        Call Kill("data.txt")\nEnd Sub\n'),
                  "file, registry or process access"),
    "call-after-then": (CORE_PATH, in_procedure('        If Flag Then Call MkDir("out")'),
                        "file, registry or process access"),
    "numbered-call": (CORE_PATH, in_procedure('100     Call FileCopy("a", "b")'),
                      "file, registry or process access"),
    "close-all-files": (CORE_PATH, in_procedure("        Close"), "file, registry or process access"),
    "close-without-hash": (CORE_PATH, in_procedure("        Close 1"), "file, registry or process access"),
    "get-without-hash": (CORE_PATH, in_procedure("        Get 1, , Value"),
                         "file, registry or process access"),
    "seek-without-hash": (CORE_PATH, in_procedure("        Seek 1, 1"), "file, registry or process access"),
    "native-declare": (CORE_PATH, core_module('Private Declare PtrSafe Function GetTickCount Lib '
                                              '"kernel32" () As Long\n'),
                       "a native Declare; core and classes call no external library"),
    "file-system-object": (CORE_PATH, in_procedure("        Dim Fso As Scripting.FileSystemObject"),
                           "Scripting.FileSystemObject is an external object model"),
    "new-file-system-object": (CLASS_PATH, class_module("Private mFso As New FileSystemObject\n"),
                               "FileSystemObject is a file-system type"),
    "standalone-new-file-system-object": (CORE_PATH, in_procedure(
        "        Dim Fso As Object\n        Set Fso = New FileSystemObject\n"
        '        Fso.CreateTextFile "output.txt"'), "FileSystemObject is a file-system type"),
    "excel-type-in-a-type-member": (CORE_PATH, core_module("Private Type Holder\n"
                                                           "    Target As Worksheet\nEnd Type\n"),
                                    "As Worksheet is an Excel or UI type"),
    "statement-after-a-type-block": (CORE_PATH, core_module("Private Type Holder\n"
                                                            "    Value As Double\nEnd Type\n\n"
                                                            "Public Sub Work()\n        Close\n"
                                                            "End Sub\n"),
                                     "file, registry or process access"),
    "database-object-model": (CORE_PATH, in_procedure("        Dim Connection As ADODB.Connection"),
                              "ADODB.Connection is an external object model"),
    "example-behind-facade-in-lower-case": (EXAMPLE_PATH, example_module("Public Sub Run()\n"
                                                                         "        Debug.Print core_validation.IsPositive(1)\n"
                                                                         "End Sub\n"),
                                            "core_validation is behind the facade"),
    "example-behind-facade": (EXAMPLE_PATH, example_module("Public Sub Run()\n"
                                                           "        Debug.Print CORE_Validation.IsPositive(1)\n"
                                                           "End Sub\n"),
                              "CORE_Validation is behind the facade"),
}


class DependencyTest(unittest.TestCase):
    def test_valid_fixtures_pass(self) -> None:
        self.assertEqual(run_check(check_vba_dependencies.check, VALID), [])
        for name, (path, source) in DEPENDENCY_VALID.items():
            with self.subTest(name=name):
                self.assertEqual(run_check(check_vba_dependencies.check, {path: source}), [])

    def test_invalid_fixtures_fail(self) -> None:
        for name, (path, source, text) in DEPENDENCY_INVALID.items():
            with self.subTest(name=name):
                findings = run_check(check_vba_dependencies.check,
                                     {path: source, "docs/VBA_HOUSE_STYLE.md": DEMO_STYLE})
                self.assertTrue(any(text in item for item in findings), findings)

    def test_facade_prefix_comes_from_the_house_style(self) -> None:
        calls = {CORE_PATH: in_procedure("        Debug.Print DEMO_IsPositive(1)")}
        # While the house style keeps the <REGIME> placeholder, an unqualified
        # facade procedure is unknown; tracked facade names are still matched.
        placeholder = "| Public facade module | `<REGIME>_` plus subject | `<REGIME>_Functions` |\n"
        self.assertEqual(run_check(check_vba_dependencies.check,
                                   {**calls, "docs/VBA_HOUSE_STYLE.md": placeholder}), [])
        findings = run_check(check_vba_dependencies.check,
                             {**calls, "docs/VBA_HOUSE_STYLE.md": DEMO_STYLE})
        self.assertTrue(any("DEMO_IsPositive names a facade" in item for item in findings),
                        findings)

    def test_tracked_component_names_match_in_any_case(self) -> None:
        adapter = 'Attribute VB_Name = "M_Run"\nOption Explicit\nPublic Sub Start()\nEnd Sub\n'
        files = {CLASS_PATH: class_module("Public Sub Go()\n        m_run.Start\nEnd Sub\n"),
                 "src/workbook/M_Run.bas": adapter}
        findings = run_check(check_vba_dependencies.check, files)
        self.assertTrue(any("m_run names a facade" in item for item in findings), findings)
        workbook_module = files["src/workbook/M_Run.bas"].replace("M_Run", "M_Inputs")
        hidden = {EXAMPLE_PATH: example_module("Public Sub Run()\n        m_inputs.Load\nEnd Sub\n"),
                  "src/workbook/M_Inputs.bas": workbook_module}
        findings = run_check(check_vba_dependencies.check, hidden)
        self.assertTrue(any("m_inputs is behind the facade" in item for item in findings), findings)

    def test_examples_do_not_use_model_classes(self) -> None:
        example = example_module("Public Sub Run()\n        Dim Result As New CCalculationResult\n"
                                 "        Result.Amount = 10\nEnd Sub\n")
        files = {EXAMPLE_PATH: example, "src/classes/CCalculationResult.cls": CLASS}
        findings = run_check(check_vba_dependencies.check, files)
        self.assertTrue(any("CCalculationResult is behind the facade" in item for item in findings),
                        findings)
        lower = {**files, EXAMPLE_PATH: example.replace("New CCalculationResult",
                                                        "New ccalculationresult")}
        findings = run_check(check_vba_dependencies.check, lower)
        self.assertTrue(any("ccalculationresult is behind the facade" in item for item in findings),
                        findings)
        counted = {EXAMPLE_PATH: example_module("Public Sub Run()\n        Debug.Print Count\n"
                                                "End Sub\n"),
                   "src/classes/CCalculationResult.cls": CLASS}
        self.assertEqual(run_check(check_vba_dependencies.check, counted), [])

    def test_keywords_inside_ordinary_identifiers_are_not_reported(self) -> None:
        # Every keyword the I/O, host and dynamic-call rules look for, used as
        # the start of an ordinary identifier, in each statement position.
        keywords = ("Open Kill FileCopy MkDir RmDir ChDir ChDrive SetAttr Reset SaveSetting "
                    "DeleteSetting SendKeys AppActivate Load Unload Name Close Print Input Write "
                    "Width Get Put Seek Lock Unlock Date Time Dir Shell GetSetting GetAllSettings "
                    "Environ CurDir FreeFile FileLen FileDateTime GetAttr Command FileAttr EOF LOF "
                    "Loc InputB Call Declare CallByName Evaluate AddressOf CreateObject GetObject "
                    + " ".join(check_vba_dependencies.HOST_NAMES)).split()
        body = "".join(f"        {word}Amount = 1\n        Total = {word}Amount\n"
                       f"        Call {word}Rate(1)\n        If Flag Then {word}Count = 2\n"
                       for word in keywords)
        helpers = "".join(f"Private Sub {word}Rate(ByVal Value As Long)\nEnd Sub\n" for word in keywords)
        source = core_module("Public Sub Work(ByVal Flag As Boolean)\n"
                             f"        Dim Total As Long\n{body}End Sub\n{helpers}")
        self.assertEqual(run_check(check_vba_dependencies.check, {CORE_PATH: source}), [])
        hashless = in_procedure("        Get 1, , Value\n        Get #1, , Value\n        Lock 1\n"
                                "        Unlock #1\n        Put 1, , Value")
        findings = run_check(check_vba_dependencies.check, {CORE_PATH: hashless})
        self.assertEqual(len(findings), 5, findings)

    def test_counts_only_restricted_homes(self) -> None:
        with repository(VALID) as root:
            _, components = check_vba_dependencies.check(root)
        self.assertEqual(components, 3)


class LexerTest(unittest.TestCase):
    def test_comments_end_the_code(self) -> None:
        cases = {
            "x = 1 ' note": "x = 1",
            "Rem whole line": "",
            "  Rem": "",
            "Debug.Print 1: Rem GoTo Missing": "Debug.Print 1:",
            'Debug.Print "a: Rem b"': 'Debug.Print "a: Rem b"',
            "Call Run(Remainder:=1)": "Call Run(Remainder:=1)",
            "Report Name:= Rem": "Report Name:= Rem",
            "Remove 1": "Remove 1",
            "100 Rem GoTo Missing": "100",
            "  10 x = 1 ' note": "  10 x = 1",
            "10: Rem note": "10:",
            "x = ['Sheet 1'!A1] ' note": "x = ['Sheet 1'!A1]",
            "x = [A1] ' note": "x = [A1]",
            "x = [A1:Rem] ' note": "x = [A1:Rem]",
            "x = [A1 : Rem + 1]: Rem note": "x = [A1 : Rem + 1]:",
            "x = 1: Remark = 2": "x = 1: Remark = 2",
            '"it\'s": GoTo Label': '"it\'s": GoTo Label',
        }
        for raw, expected in cases.items():
            with self.subTest(raw=raw):
                self.assertEqual(strip_code(raw), expected)
        self.assertEqual(strip_code('Debug.Print "GoTo x": GoTo y', blank_strings=True),
                         "Debug.Print         : GoTo y")
        self.assertEqual(strip_code("x = ['GoTo z'!A1]: GoTo y", blank_strings=True),
                         "x = [" + " " * len("'GoTo z'!A1") + "]: GoTo y")

    def test_continued_comment_lines_are_comment(self) -> None:
        lines = ["x = 1 ' note _", "  GoTo Missing _", "  MsgBox 1", "y = 2", "' plain _"]
        self.assertEqual(logical_units(lines), [(1, 3, "code", "x = 1"), (4, 4, "code", "y = 2"),
                                                (5, 5, "code", "")])
        # A continuation in code still joins the lines.
        self.assertEqual(logical_units(["x = 1 + _", "  2"]), [(1, 2, "code", "x = 1 + 2")])


IDENTIFIERS_VALID = {
    "declarations-of-every-kind": """Private mCount As Long, mNames() As String
Private Const LIMIT As Long = 3
Private Enum Slot
    [_First] = 0
    slA = 1
End Enum

Public Function Total( _
    ByVal Value As Long, _
    Optional ByRef Label As String = "a, b", _
    Optional ByVal Extra As Variant) _
    As Long

        Dim i As Long, Items As Collection
        Static Calls As Long
        Const STEP_SIZE As Long = 1
        ReDim Buffer(1 To LIMIT) As Long
        ReDim Preserve Buffer(1 To LIMIT + 1), Spare$(2)
        ReDim Items.Values(1 To 3), Items!Keys(1)
        ReDim Ratios!(1 To 3)
        Ratios!(1) = 1
        Items!Count = 3
        Call Items(1).Refresh
        Call Items!First.Refresh
        Items(1).Refresh
        With Items
            ReDim .Values(1 To 4)
        End With
        Dim Fixed As String * 8, Range2 As Excel.Range, Made As New Collection
        Set Items = New Collection
        For i = 1 To LIMIT Step STEP_SIZE: Buffer(i) = i: Next i
        For Each Extra In Items
        Next Extra
        If Value > 0 Then mCount = slA Else mCount = LIMIT
        If Value > 1 Then mCount = 1: mCount = 2 Else mCount = 3: Calls = 0
        Dim Settings As Long, Letter As Long, Reset As Long
        Settings = 1: Letter = 2
        Set Items = Nothing
        Mid$(Label, 1, 1) = "z"
        Calls = Calls + 1
        Total = Value
End Function
""",
    "calls-of-every-kind": """Public Sub Run()
        Helper 1
        Call Helper(2)
        Call TEST_Other.Thing(3)
        TEST_Other.Thing 4
        Debug.Print "Undeclared = 1"
        Err.Raise 5
        MsgBox "x"
        DoEvents
        Shared
Retry:
        On Error GoTo Retry
End Sub

Private Sub Helper(ByVal Value As Long)
        With Application
            .StatusBar = Value
        End With
        Application.StatusBar = False
End Sub
""",
    "a-declaration-in-each-active-branch": """Public Sub Run()
#If VBA7 Then
        Dim Handle As LongPtr
#Else
        Dim Handle As Long
#End If
        Handle = 0
End Sub
""",
}
IDENTIFIERS_INVALID = {
    "undeclared-assignment": ("Public Sub Run()\n        Total = 1\nEnd Sub\n",
                              "Total is assigned in Run but not declared"),
    "undeclared-set": ("Public Sub Run()\n        Set Items = Nothing\nEnd Sub\n",
                       "Items is assigned in Run but not declared"),
    "undeclared-loop-counter": ("Public Sub Run()\n        For i = 1 To 2\n        Next i\n"
                                "End Sub\n", "i is assigned in Run but not declared"),
    "assignment-after-then": ("Public Sub Run()\n        If True Then Missing = 1\nEnd Sub\n",
                              "Missing is assigned in Run but not declared"),
    "assignment-after-else-past-a-colon": (
        "Public Sub Run()\n        Dim x As Long\n"
        "        If True Then x = 1: x = 2 Else Missing = 3\nEnd Sub\n",
        "Missing is assigned in Run but not declared"),
    "call-after-else-past-a-colon": (
        "Public Sub Run()\n        Dim x As Long\n"
        "        If True Then x = 1: x = 2 Else Gone 3\nEnd Sub\n",
        "Gone is called in Run"),
    "a-declaration-only-in-an-inactive-branch": (
        "Public Sub Run()\n#If Not VBA7 Then\n        Dim Legacy As Long\n#End If\n"
        "        Legacy = 1\nEnd Sub\n",
        "Legacy is assigned in Run but not declared [vba7-win32, vba7-win64]"),
    "a-renamed-private-helper": (
        "Public Sub Run()\n        OldHelper 1\nEnd Sub\n\n"
        "Private Sub NewHelper(ByVal Value As Long)\nEnd Sub\n",
        "OldHelper is called in Run but no procedure of that name is reachable"),
    "an-explicit-call-to-nothing": ("Public Sub Run()\n        Call Gone(1)\nEnd Sub\n",
                                    "Gone is called in Run"),
    "an-explicit-call-without-arguments": ("Public Sub Run()\n        Call Gone\nEnd Sub\n",
                                           "Gone is called in Run"),
    "dim-empty": ("Public Sub Run()\n        Dim Empty() As Long\nEnd Sub\n",
                  "variable Empty is named like the VBA keyword Empty"),
    "keyword-parameter": ("Public Sub Run(ByVal Date As Long)\nEnd Sub\n",
                          "parameter Date is named like the VBA keyword Date"),
    "keyword-module-name": ("Private Name As String\n",
                            "variable Name is named like the VBA keyword Name"),
    "keyword-enum-member": ("Private Enum Slot\n    Error = 1\nEnd Enum\n",
                            "Enum member Error is named like the VBA keyword Error"),
    "unreadable-parameters": ("Public Sub Run(ByVal 1st As Long)\nEnd Sub\n",
                              "cannot read the parameter"),
    "unreadable-declaration": ("Public Sub Run()\n        Dim 2nd As Long\nEnd Sub\n",
                               "cannot read the declaration"),
    "unreadable-redim": ("Public Sub Run()\n        ReDim 3rd(1)\nEnd Sub\n",
                         "cannot read the declaration"),
    "single-line-if-without-then": ("Public Sub Run()\n        If True x = 1\nEnd Sub\n",
                                    "cannot read the If statement"),
}
OTHER = """Attribute VB_Name = "TEST_Other"
Option Explicit
Option Private Module

Public Shared As Long

Public Sub Thing(ByVal Value As Long)
End Sub

Public Sub Shared()
End Sub
"""


class IdentifierTest(unittest.TestCase):
    def test_valid_fixtures_pass(self) -> None:
        self.assertEqual(run_check(check_vba_identifiers.check, VALID), [])
        for name, body in IDENTIFIERS_VALID.items():
            with self.subTest(name=name):
                files = {**standalone(module(body)), "tests/modules/TEST_Other.bas": OTHER}
                self.assertEqual(run_check(check_vba_identifiers.check, files), [])

    def test_invalid_fixtures_fail(self) -> None:
        for name, (body, text) in IDENTIFIERS_INVALID.items():
            with self.subTest(name=name):
                findings = run_check(check_vba_identifiers.check, standalone(module(body)))
                self.assertTrue(any(text in item for item in findings), findings)

    def test_public_names_of_other_standard_modules_resolve(self) -> None:
        caller = module("Public Sub Run()\n        Shared = 1\n        Thing 2\nEnd Sub\n")
        files = {**standalone(caller), "tests/modules/TEST_Other.bas": OTHER}
        self.assertEqual(run_check(check_vba_identifiers.check, files), [])
        private = OTHER.replace("Public Sub Thing", "Private Sub Thing")
        files["tests/modules/TEST_Other.bas"] = private
        self.assertEqual(run_check(check_vba_identifiers.check, files),
                         ["tests/modules/TEST_Fixture.bas:6: Thing is called in Run but no "
                          "procedure of that name is reachable [vba6-win32, vba7-win32, "
                          "vba7-win64]"])

    def test_form_controls_and_class_properties_count_as_declared(self) -> None:
        form = FORM.replace("   OleObjectBlob", "   Begin MSForms.TextBox txtAmount\n   End\n"
                            "   OleObjectBlob") + ("\nPrivate Sub UserForm_Initialize()\n"
                                                    "        txtAmount = \"0\"\n        Hide\n"
                                                    "        Me!txtAmount = \"1\"\n"
                                                    "        Call Me!txtAmount.SetFocus\n"
                                                    "End Sub\n")
        files = changed("src/forms/frmRunOptions.frm", FORM, form)
        cls = CLASS + "\nPublic Sub Reset()\n        Amount = 0\nEnd Sub\n"
        files = changed("src/classes/CCalculationResult.cls", CLASS, cls, files)
        self.assertEqual(run_check(check_vba_identifiers.check, files), [])

    def test_host_properties_are_assignable_in_their_module_kind(self) -> None:
        form = FORM + ('\nPrivate Sub UserForm_Initialize()\n        Caption = "Options"\n'
                       '        Width = 200\n        Range("A1") = 1\nEnd Sub\n')
        sheet = WORKBOOK + ('\nPrivate Sub Workbook_Open()\n        Saved = True\n'
                            '        Cells(1, 1) = 2\n        RefreshAll\n        Save\n'
                            'End Sub\n')
        files = changed("src/workbook/ThisWorkbook.cls", WORKBOOK, sheet,
                        changed("src/forms/frmRunOptions.frm", FORM, form))
        self.assertEqual(run_check(check_vba_identifiers.check, files), [])
        # A form property is not a name of a standard module, and a typo is
        # still undeclared in a form.
        standard = module("Public Sub Run()\n        Caption = \"x\"\n        Range(\"A1\") = 1\n"
                          "End Sub\n")
        self.assertEqual(run_check(check_vba_identifiers.check, standalone(standard)),
                         ["tests/modules/TEST_Fixture.bas:5: Caption is assigned in Run but not "
                          "declared [vba6-win32, vba7-win32, vba7-win64]"])
        hidden = module("Public Sub Run()\n        Hide\nEnd Sub\n")
        self.assertIn("Hide is called in Run", run_check(check_vba_identifiers.check,
                                                         standalone(hidden))[0])
        typo = form.replace("Caption = ", "Captoin = ")
        found = run_check(check_vba_identifiers.check,
                          changed("src/forms/frmRunOptions.frm", FORM, typo))
        self.assertEqual(len(found), 1)
        self.assertIn("Captoin is assigned in UserForm_Initialize but not declared", found[0])

    def test_counts_components_and_procedures(self) -> None:
        with repository(VALID) as root:
            _, components, procedures = check_vba_identifiers.check(root)
        self.assertEqual((components, procedures), (7, 6))


CALLEE = """Attribute VB_Name = "TEST_Callee"
Option Explicit
Option Private Module

Public Sub Pair(ByVal First As Long, ByVal Second As Long)
End Sub

Public Function Scaled(ByVal Value As Double, Optional ByVal Factor As Double = 1) As Double
        Scaled = Value * Factor
End Function

Public Sub Collect(ByVal Label As String, ParamArray Values() As Variant)
End Sub

Public Function Grid() As Variant
        Grid = Array(1, 2)
End Function

Public Property Get Size(ByVal Index As Long) As Long
        Size = Index
End Property

Private Sub Hidden(ByVal Value As Long)
End Sub

Private Function Secret() As Long
End Function
"""
CALLS_VALID = """
Private Totals(1 To 3) As Double


Public Sub Run()
        Dim Values(1 To 2) As Double
        Dim Items As Collection
        Dim Pair(1 To 2) As Long
        Pair(1) = 1
        TEST_Callee.Pair 1, 2
        Call TEST_Callee.Pair(1, Second:=2)
        Debug.Print Scaled(1), Scaled(1, 2), TEST_Callee.Scaled(Value:=1, Factor:=2)
        Collect "a, b"
        Collect "a", 1, , "c", 3
        Call Collect(Label:="a")
        Values(1) = Scaled(Values(2))
        Totals(1) = Values(1)
        If Scaled(1) > 0 Then Collect "x", 1 Else Helper 1
        Debug.Print Grid(1), TEST_Callee.Size(1, 2)
        Set Items = New Collection
        With Items
            .Add Scaled(2), "Key"
            .Add 1, 2, 3, 4, 5
            .Scaled = 1
        End With
        Items.Scaled 1, 2, 3
        Helper Value:=1
        TEST_Fixture.Helper 1
        Debug.Print "Scaled(1, 2, 3)"
End Sub


Private Sub Helper(ByVal Value As Long)
End Sub
"""
CALLS_INVALID = {
    "too-few": ("TEST_Callee.Pair 1", "call to TEST_Callee.Pair in Run: required argument(s) "
                                      "Second not passed"),
    "too-many": ("Pair 1, 2, 3", "call to TEST_Callee.Pair in Run: 3 argument(s) passed, at "
                                 "most 2 accepted"),
    "omitted-required": ('Pair , "2"', "required argument(s) First not passed"),
    "too-many-in-an-expression": ("Debug.Print Scaled(1, 2, 3)",
                                  "call to TEST_Callee.Scaled in Run: 3 argument(s) passed"),
    "too-few-nested": ("Debug.Print Scaled(Scaled())",
                       "call to TEST_Callee.Scaled in Run: required argument(s) Value"),
    "unknown-named": ("Call Pair(1, Second:=2, Third:=3)", "no parameter is named Third"),
    "paramarray-without-its-required-argument": ("Collect", "required argument(s) Label"),
    "private-of-another-module": ("TEST_Callee.Hidden 1", "call to TEST_Callee.Hidden in Run: "
                                                          "Hidden is Private in TEST_Callee"),
    "private-in-an-expression": ("Debug.Print TEST_Callee.Hidden(1)",
                                 "Hidden is Private in TEST_Callee"),
    "private-without-parameters": ("Debug.Print TEST_Callee.Secret()",
                                   "Secret is Private in TEST_Callee"),
}


class CallTest(unittest.TestCase):
    def files(self, body: str) -> dict[str, str | bytes]:
        return {**standalone(module(body)), "tests/modules/TEST_Callee.bas": CALLEE}

    def test_valid_calls_pass(self) -> None:
        self.assertEqual(run_check(check_vba_calls.check, self.files(CALLS_VALID)), [])
        self.assertEqual(run_check(check_vba_calls.check, VALID), [])

    def test_invalid_calls_are_reported(self) -> None:
        for name, (statement, text) in CALLS_INVALID.items():
            with self.subTest(name=name):
                body = f"\nPublic Sub Run()\n        {statement}\nEnd Sub\n"
                findings = run_check(check_vba_calls.check, self.files(body))
                self.assertEqual(len(findings), 1, findings)
                self.assertIn(text, findings[0])
                self.assertTrue(findings[0].startswith("tests/modules/TEST_Fixture.bas:6: "))

    def test_date_literals_and_type_suffixes(self) -> None:
        # A comma inside a date literal separates no arguments; a # type
        # suffix opens no date literal, so the arguments after it still count.
        valid = ("\nPublic Sub Run()\n        Dim Amount As Double\n"
                 "        Pair #January 1, 2020#, 2\n        Pair Amount#, 2#\n"
                 "        Debug.Print Scaled(#1/2/2020#, 1#)\nEnd Sub\n")
        self.assertEqual(run_check(check_vba_calls.check, self.files(valid)), [])
        invalid = valid.replace("Pair Amount#, 2#", "Pair Amount#, 2#, #3/4/2020#")
        findings = run_check(check_vba_calls.check, self.files(invalid))
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("3 argument(s) passed, at most 2 accepted", findings[0])

    def test_redim_bounds_are_expressions(self) -> None:
        # The bounds of a ReDim are evaluated at run time; its target is an
        # array, even when it is named like a procedure or indexed past the
        # parameters of one.
        valid = ("\nPublic Sub Run()\n        Dim Values() As Double\n        Dim Pair() As Long\n"
                 "        ReDim Values(1 To Scaled(2, 3))\n"
                 "        ReDim Preserve Values(Scaled(1), 0 To TEST_Callee.Scaled(4)), "
                 "Pair(1, 2, 3)\nEnd Sub\n")
        self.assertEqual(run_check(check_vba_calls.check, self.files(valid)), [])
        cases = {"ReDim Values(Scaled(1, 2, 3))": "3 argument(s) passed, at most 2 accepted",
                 "ReDim Preserve Values(1 To 2, 0 To TEST_Callee.Hidden(1))":
                     "Hidden is Private in TEST_Callee"}
        for statement, text in cases.items():
            with self.subTest(statement=statement):
                body = ("\nPublic Sub Run()\n        Dim Values() As Double\n"
                        f"        {statement}\nEnd Sub\n")
                findings = run_check(check_vba_calls.check, self.files(body))
                self.assertEqual(len(findings), 1, findings)
                self.assertIn(text, findings[0])

    def test_with_block_members_are_not_calls(self) -> None:
        body = ("\nPublic Sub Run()\n        With New Collection\n            .Pair = 1\n"
                "            .Pair 1, 2, 3\n            .Hidden\n"
                "            Debug.Print .Scaled(1, 2, 3)\n        End With\nEnd Sub\n")
        self.assertEqual(run_check(check_vba_calls.check, self.files(body)), [])

    def test_repository_vba_has_no_findings(self) -> None:
        findings, components, resolved = check_vba_calls.check(ROOT)
        self.assertEqual(findings, [])
        if components:
            self.assertGreater(resolved, 0)

    def test_main_exits_1_on_a_finding(self) -> None:
        body = "\nPublic Sub Run()\n        TEST_Callee.Hidden 1\nEnd Sub\n"
        for files, expected in ((self.files(CALLS_VALID), 0), (self.files(body), 1)):
            with self.subTest(expected=expected), repository(files) as root, \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(check_vba_calls.main(["--root", str(root)]), expected)
                self.assertIn("resolved call(s); static check", output.getvalue())


def handler(statements: str) -> str:
    """A procedure whose handler, from line 9, runs ``statements``; the
    module header takes lines 1 to 3."""
    return ("Private Sub First(ByVal Retry As Boolean)\n"
            "        Dim Number As Long, Source As String, Text As String\n"
            "        On Error GoTo Handler\n"
            "        Exit Sub\n"
            "Handler:\n" + statements + "End Sub\n\n\n"
            "Private Sub Helper()\nEnd Sub\n\n\n"
            "Private Sub Report(ByVal Value As Variant)\nEnd Sub\n")


def guarded(statements: str) -> str:
    """A procedure that runs ``statements`` from line 7, after On Error Resume
    Next at line 6."""
    return ("Private Sub First(ByVal Retry As Boolean)\n"
            "        Dim Number As Long, Flag As Boolean, Text As String\n"
            "        On Error Resume Next\n" + statements + "End Sub\n\n\n"
            "Private Sub Helper()\nEnd Sub\n\n\n"
            "Private Sub Report(ByVal Value As Variant)\nEnd Sub\n")


def fetch(result: str, statements: str) -> str:
    """A function declared with ``result``, whose handler, from line 9, runs
    ``statements``."""
    return (f"Private Function Fetch(ByVal Retry As Boolean){result}\n"
            "        Dim Number As Long\n        On Error GoTo Handler\n"
            "        Exit Function\nHandler:\n" + statements + "End Function\n")


ERRORS_VALID = {
    "copied-before-any-call": handler(
        "        Number = Err.Number\n        Source = Err.Source\n"
        "        Text = Err.Description\n        Helper\n        Debug.Print Text\n"
        "        Err.Raise Number, Source, Text\n"),
    "arguments-read-before-the-call": handler(
        '        Report "failed: " & Err.Description\n'),
    "raise-again-first": handler("        Err.Raise Err.Number, Err.Source, Err.Description\n"),
    "pure-conversions": handler(
        "        Number = CLng(Err.Number)\n"
        '        Source = IIf(Len(Err.Source) > 0, Err.Source, "First")\n'
        "        Text = CStr(Err.Description)\n        Helper\n"),
    "plain-copies-and-literals": handler(
        "        Retry = False\n        Number = -1\n        Text = Source\n"
        "        Number = Err.Number\n"),
    "concatenation-in-a-handler": handler(
        '        Text = "failed: " & Source\n        Number = Err.Number\n'),
    "resume-next-guard": guarded(
        "        Helper\n        If Err.Number <> 0 Then Exit Sub\n        On Error GoTo 0\n"),
    "resume-next-copy-after-the-call": guarded(
        "        Helper\n        Number = Err.Number\n        Text = Err.Description\n"
        "        On Error GoTo 0\n"),
    "resume-next-concatenation-is-the-operation": guarded(
        '        Text = "failed: " & Text\n        Number = Err.Number\n'),
    "block-if-ends-the-segment": handler(
        "        If Retry Then\n            Helper\n        End If\n"
        "        Report Err.Description\n"),
    "single-line-if-ends-the-segment": handler(
        "        If Retry Then Helper\n        Report Err.Source\n"),
    "loop-ends-the-segment": handler(
        "        Do While Retry\n            Helper\n        Loop\n        Report Err.Number\n"),
    "label-starts-a-segment": handler(
        "        Helper\nAgain:\n        Report Err.Number\n"),
    "err-in-strings-and-comments": handler(
        "        Helper\n        Debug.Print \"Err.Number\" ' Err.Description\n"),
    "array-element-in-a-handler": handler(
        "        Dim Values(1 To 2) As Long\n        Number = Values(1)\n"
        "        Text = Err.Description\n"),
    "values-without-parentheses": (
        "Private Const LIMIT As Long = 3\nPrivate Enum Mode\n    mdFirst = 1\nEnd Enum\n"
        "Private mCount As Long\n\n\n" + handler(
            "        Number = LIMIT\n        Number = mCount\n        Number = mdFirst\n"
            "        Text = vbNullString\n        Text = Source\n        Number = First\n"
            "        Number = Err.Number\n")),
    "vba-qualified-pure-functions": handler(
        "        Number = VBA.CLng(Err.Number)\n        Number = VBA.Len(Err.Source)\n"
        "        Text = vba.cstr(Err.Description)\n        Source = Err.Source\n"),
    "redim-and-erase-are-statements": handler(
        "        Dim Values() As Long\n        ReDim Values(1 To 2)\n"
        "        ReDim Preserve Values(1 To 3)\n        Erase Values\n"
        "        Number = Err.Number\n"),
    "statement-targets-and-file-statements": handler(
        "        Dim Book As Object\n        Set Book = Nothing\n"
        '        Mid$(Text, 1, 1) = "x"\n        LSet Text = Source\n        Let Number = 1\n'
        "        Open Text For Input Access Read Shared As #1\n        Line Input #1, Text\n"
        "        Print #1, Spc(2); Text\n        Close #1\n        Number = Err.Number\n"),
    "resume-next-operation-evaluated-before-the-read": guarded(
        "        Text = CStr(Number / 2) & Err.Description\n"),
    "err-writes-are-not-reads": handler(
        '        Err.Clear\n        Err.Description = "custom"\n        Err.Number = 5\n'
        '        Let Err.Source = "First"\n        VBA.Err.HelpContext = 0\n'),
    "context-added-then-raised-again": handler(
        '        Err.Description = "context: " & Err.Description\n'
        "        Err.Raise Err.Number, Err.Source, Err.Description\n"),
    "with-err-context-added-then-raised-again": handler(
        '        With Err\n            .Description = "context: " & .Description\n'
        "            .Raise .Number, .Source, .Description\n        End With\n"),
    "with-err-copied-first": handler(
        "        With Err\n            Number = .Number\n            Text = .Description\n"
        "        End With\n        Helper\n"),
    "read-before-a-call-in-the-same-statement": handler(
        "        Number = Err.Number + Describe()\n")
    + "\n\nPrivate Function Describe() As Long\nEnd Function\n",
    "read-as-an-argument-before-a-call": handler(
        "        Report Err.Number, Describe()\n")
    + "\n\nPrivate Function Describe() As Long\nEnd Function\n",
    "single-line-if-branch-reads-first": handler("        If Retry Then Report Err.Number\n"),
    "single-line-if-branches-exclude-each-other": handler(
        "        If Retry Then Helper Else Report Err.Number\n"),
    "nested-single-line-if-else-belongs-to-the-inner-if": handler(
        "        If Retry Then If Number = 0 Then Helper Else Report Err.Number\n"),
    "own-name-is-the-result-variable": fetch(
        " As Long", "        Fetch = 1\n        Number = Fetch\n        Number = Err.Number\n"),
    "own-name-indexes-an-array-result": fetch(
        " As Long()", "        Number = Fetch(1)\n        Number = Err.Number\n"),
    "own-name-indexes-a-variant-result": fetch(
        "", "        Number = Fetch(1)\n        Number = Err.Number\n"),
    "other-err-properties-are-data": handler(
        "        Text = Err.HelpFile\n        Number = Err.HelpContext\n"
        "        Number = Err.LastDllError\n        Number = VBA.Err.HelpContext\n"
        "        Number = Err.Number\n"),
    "context-added-then-raised-again-with-help": handler(
        '        Err.Description = "context: " & Err.Description\n'
        "        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, "
        "Err.HelpContext\n"),
    "with-err-other-properties": handler(
        "        With Err\n            Text = .HelpFile\n            Number = .HelpContext\n"
        "            Number = .Number\n        End With\n"),
    "help-properties-are-not-checked-for-staleness": handler(
        "        Helper\n        Text = Err.HelpFile\n        Number = Err.HelpContext\n"),
    "resume-next-qualified-conversion-can-fail": guarded(
        '        Number = VBA.CLng("bad")\n        Number = Err.Number\n'),
    "resume-next-qualified-len-can-fail": guarded(
        "        Number = VBA.Len(Text)\n        Number = Err.Number\n"),
    "qualified-vba-names-are-vba": handler(
        "        Text = VBA.vbCrLf & Err.Description\n        VBA.Beep\n"
        '        VBA.Mid$(Text, 1, 1) = "x"\n        Number = VBA.Conversion.CLng(Err.Number)\n'
        "        Source = Err.Source\n"),
}
# Procedures that the fixtures call without parentheses or assign to.
DESCRIBE = "\n\nPrivate Function Describe() As Long\nEnd Function\n"
LABEL = "\n\nPrivate Property Get Label() As String\nEnd Property\n"
STATUS = "\n\nPrivate Property Let Status(ByVal Value As String)\nEnd Property\n"
OWNER = "\n\nPrivate Property Set Owner(ByVal Value As Object)\nEnd Property\n"
ERRORS_INVALID = {
    "after-a-call": (handler("        Helper\n        Report Err.Description\n"), 10,
                     "Err.Description read in First after line 9, which may have changed Err "
                     "with a call or member access"),
    "after-debug-print": (handler('        Debug.Print "failed"\n'
                                  "        Err.Raise Err.Number, Err.Source, Err.Description\n"),
                          10, "Err.Number, Err.Source, Err.Description read in First after line 9"),
    "after-a-host-property": (handler("        Application.StatusBar = False\n"
                                      "        Number = Err.Number\n"), 10,
                              "Err.Number read in First after line 9, which may have changed"),
    "after-err-clear": (handler("        Err.Clear\n        Number = Err.Number\n"), 10,
                        "Err.Number read in First after line 9"),
    "after-a-function-result": (handler("        Text = Format$(Now, \"yyyy\") & Describe(1)\n"
                                        "        Source = Err.Source\n"), 10,
                                "Err.Source read in First after line 9"),
    "after-on-error-in-a-handler": (handler("        On Error GoTo 0\n"
                                            "        Report Err.Description\n"), 10,
                                    "Err.Description read in First after line 9, which cleared "
                                    "Err with an On Error statement"),
    "after-a-plain-copy-only": (handler("        Helper\n        Retry = False\n"
                                        "        Number = Err.Number\n"), 11,
                                "Err.Number read in First after line 9"),
    "inside-a-branch": (handler("        If Retry Then\n            Helper\n"
                                "            Report Err.Number\n        End If\n"), 11,
                        "Err.Number read in First after line 10"),
    "on-the-line-of-the-handler-label": (
        handler("").replace("Handler:\n", "Handler: Helper: Number = Err.Number\n"), 8,
        "Err.Number read in First after line 8"),
    "guard-after-on-error-goto-0": (guarded("        Helper\n        On Error GoTo 0\n"
                                            "        If Err.Number <> 0 Then Exit Sub\n"), 9,
                                    "Err.Number read in First after line 8, which cleared Err"),
    "resume-next-before-any-operation": (guarded("        Report Err.Description\n"), 7,
                                         "Err.Description read in First after line 6"),
    "resume-next-declaration-blank-line-and-comment-run-nothing": (
        guarded("        Dim Local As Long\n\n        ' a comment\n        Number = Err.Number\n"),
        10, "Err.Number read in First after line 6"),
    "resume-next-err-write-runs-nothing": (
        guarded("        Err.Number = 0\n        Number = Err.Number\n"), 8,
        "Err.Number read in First after line 6"),
    "resume-next-err-clear": (
        guarded("        Helper\n        Err.Clear\n        Text = Err.Description\n"), 9,
        "Err.Description read in First after line 8, which cleared Err with Err.Clear"),
    "resume-next-call-err-clear": (
        guarded("        Helper\n        Call Err.Clear\n        Text = Err.Description\n"), 9,
        "Err.Description read in First after line 8, which cleared Err with Err.Clear"),
    "resume-next-call-err-clear-with-parentheses": (
        guarded("        Helper\n        Call VBA.Err.Clear()\n        Number = Err.Number\n"), 9,
        "Err.Number read in First after line 8, which cleared Err with Err.Clear"),
    "select-case-after-on-error": (guarded("        Helper\n        On Error GoTo 0\n"
                                           "        Select Case Err.Number\n"
                                           "        End Select\n"), 9,
                                   "Err.Number read in First after line 8"),
    "after-a-function-without-parentheses": (
        handler("        Number = Describe\n        Text = Err.Description\n") + DESCRIBE, 10,
        "Err.Description read in First after line 9, which may have changed Err"),
    "after-a-property-without-parentheses": (
        handler("        Text = Label\n        Number = Err.Number\n") + LABEL, 10,
        "Err.Number read in First after line 9"),
    "after-a-declare-without-parentheses": (
        'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long\n'
        + handler("        Number = GetTickCount\n        Source = Err.Source\n"), 11,
        "Err.Source read in First after line 10"),
    "after-a-host-global-without-parentheses": (
        handler("        Text = ActiveCell\n        Number = Err.Number\n"), 10,
        "Err.Number read in First after line 9"),
    "after-a-property-let": (
        handler('        Status = "failed"\n        Number = Err.Number\n') + STATUS, 10,
        "Err.Number read in First after line 9, which may have changed Err"),
    "after-a-property-let-with-let": (
        handler('        Let Status = "failed"\n        Text = Err.Description\n') + STATUS, 10,
        "Err.Description read in First after line 9"),
    "after-a-property-set": (
        handler("        Set Owner = Nothing\n        Source = Err.Source\n") + OWNER, 10,
        "Err.Source read in First after line 9"),
    "after-a-function-in-redim-bounds": (
        handler("        Dim Values() As Long\n        ReDim Values(1 To Describe)\n"
                "        Number = Err.Number\n") + DESCRIBE, 11,
        "Err.Number read in First after line 10"),
    "after-a-with-block-member": (
        handler("        Dim Book As Object\n        With Book\n            .Saved = True\n"
                "        End With\n        Number = Err.Number\n"), 13,
        "Err.Number read in First after line 11"),
    "resume-next-read-inside-a-conversion": (
        guarded("        Number = CLng(Err.Number)\n"), 7,
        "Err.Number read in First after line 6, which cleared Err"),
    "resume-next-read-inside-len-in-a-condition": (
        guarded("        If Len(Err.Description) > 0 Then Exit Sub\n"), 7,
        "Err.Description read in First after line 6"),
    "resume-next-read-inside-iif": (
        guarded('        Text = IIf(Len(Err.Description) > 0, Err.Description, "none")\n'), 7,
        "Err.Description read in First after line 6"),
    "resume-next-read-inside-a-call-argument": (
        guarded("        Call Report(CStr(Err.Number))\n"), 7,
        "Err.Number read in First after line 6"),
    "resume-next-argument-read-before-a-later-one-calls": (
        guarded("        Report Err.Number, Describe()\n") + DESCRIBE, 7,
        "Err.Number read in First after line 6"),
    "with-err-member-read-after-a-call": (
        handler("        Helper\n        With Err\n            Number = .Number\n"
                "        End With\n"), 11,
        "Err.Number read in First after line 9"),
    "single-line-if-call-then-read-in-the-same-branch": (
        handler("        If Retry Then Helper: Report Err.Number\n"), 9,
        "Err.Number read in First after line 9, which may have changed Err"),
    "single-line-if-call-then-read-in-the-else-branch": (
        handler("        If Retry Then Report 1 Else Helper: Report Err.Source\n"), 9,
        "Err.Source read in First after line 9"),
    "single-line-if-condition-call-then-read": (
        handler("        If Describe() Then Report Err.Number\n") + DESCRIBE, 9,
        "Err.Number read in First after line 9"),
    "recursive-call": (
        fetch(" As Long", "        Number = Fetch(True)\n        Number = Err.Number\n"), 10,
        "Err.Number read in Fetch after line 9, which may have changed Err"),
    "module-name-that-is-a-parameter": (
        handler("        Number = Strings.Len(1)\n        Number = Err.Number\n").replace(
            "ByVal Retry As Boolean", "ByVal Strings As Object"), 10,
        "Err.Number read in First after line 9"),
    "call-before-a-read-in-the-same-statement": (
        handler("        Number = Describe() + Err.Number\n") + DESCRIBE, 9,
        "Err.Number read in First after a call or member access that the same statement "
        "evaluates first"),
    "call-argument-before-a-read": (
        handler("        Report Describe(), Err.Number\n") + DESCRIBE, 9,
        "Err.Number read in First after a call or member access"),
    "member-access-before-a-read": (
        handler("        Text = Application.Name & Err.Description\n"), 9,
        "Err.Description read in First after a call or member access"),
    "err-write-reads-its-value-after-a-call": (
        handler('        Helper\n        Err.Description = "context: " & Err.Description\n'), 10,
        "Err.Description read in First after line 9"),
}
ALL_ENVIRONMENTS = "[vba6-win32, vba7-win32, vba7-win64]"
# VBA statements and functions that run no project code, with literals and
# declared variables as operands: none may make the next read of Err stale.
STATEMENT_SWEEP = [
    "Open Text For Input As #1", "Open Text For Append Shared As FileNumber",
    "Open Text For Random Access Read Write Lock Read Write As #FileNumber Len = 128",
    "Close #1", "Close", "Reset", "Get #1, , Number", "Put #1, 1, Text", "Seek #1, 1",
    "Lock #1", "Unlock #1, 1 To 2", "Width #1, 80", "Write #1, Text, Number",
    "Input #1, Text, Number", "Line Input #1, Text", "Print #1, Text; Number",
    "Print #1, Spc(2); Tab(4); Text", "Name Text As Text", "Erase Values",
    "Erase Values, Other", "ReDim Values(1 To 2)", "ReDim Values(1 To 2) As Long",
    "ReDim Preserve Values(1 To Number), Other(2) As String",
    "ReDim Values(0 To 1, 1 To 3) As Long", "Date = Now", "Time = #12:00:00 PM#",
    'Mid(Text, 1, 1) = "x"', "Mid$(Text, 2) = Text", 'MidB(Text, 1, 2) = "ab"',
    'LSet Text = "x"', "RSet Text = Text", "Let Number = 1", "Set Item = Nothing",
    "Static Count As Long", "Dim Local As Long", "Dim Fixed As String * 10",
    "Dim Things As New Collection", "Const LIMIT As Long = 3", "Stop",
    "Number = Sqr(4) + Abs(-1) + Rnd", "FileNumber = FreeFile", "Number = LOF(1) + Loc(1)",
    "Flag = EOF(1)", 'Text = Format(Date, "yyyy") & Dir(Text) & Environ("TEMP")',
    'Number = DateDiff("d", Date, Now)', "If TypeOf Item Is Collection Then Number = 1",
    "If TypeOf Item Is VBA.Collection Then Number = 1",
    # Error raises: in a handler the read after it is never reached, and after
    # On Error Resume Next it reads the error Error raised, so no claim.
    "Error 5",
]
# One call of each VBA function of PURE_FUNCTIONS, with arguments it accepts;
# Spc and Tab exist only in Print. The test checks that this table and the
# check's set name the same functions.
PURE_CALLS = {
    **{name: f"Result = {name}(Number)" for name in """
        Abs Atn CBool CByte CCur CDbl CDec Chr ChrB ChrW CInt CLng CLngLng CLngPtr Cos CSng CStr
        CVar Exp Fix Hex Int IsEmpty IsError IsNull IsNumeric IsObject Log Oct QBColor Sgn Sin
        Space Sqr Str Tan TypeName VarType MonthName WeekdayName
        """.split()},
    **{name: f"Result = {name}(Text)" for name in """
        Asc AscB AscW CurDir Dir FileDateTime FileLen GetAttr IsDate LCase Len LenB LTrim RTrim
        Split StrReverse Trim UCase Val
        """.split()},
    **{name: f"Result = {name}(Now)" for name in """
        CDate CVDate Day FormatDateTime Hour Minute Month Second Weekday Year
        """.split()},
    **{name: f"Result = {name}" for name in "Date Now Rnd Time Timer FreeFile".split()},
    "Array": "Result = Array(1, 2)", "Choose": "Result = Choose(1, Text)",
    "CVErr": "Result = CVErr(2007)", "DateAdd": 'Result = DateAdd("d", 1, Now)',
    "DateDiff": 'Result = DateDiff("d", Now, Now)', "DatePart": 'Result = DatePart("d", Now)',
    "DateSerial": "Result = DateSerial(2026, 1, 1)",
    "DateValue": 'Result = DateValue("2026-01-01")', "DDB": "Result = DDB(1000, 100, 5, 1)",
    "Environ": 'Result = Environ("TEMP")', "EOF": "Result = EOF(1)", "Loc": "Result = Loc(1)",
    "LOF": "Result = LOF(1)", "Filter": "Result = Filter(Other, Text)",
    "Format": 'Result = Format(Number, "0.00")',
    "FormatCurrency": "Result = FormatCurrency(Number)",
    "FormatNumber": "Result = FormatNumber(Number, 2)",
    "FormatPercent": "Result = FormatPercent(Number)", "FV": "Result = FV(0.1, 12, -100)",
    "GetSetting": 'Result = GetSetting("App", "Section", "Key")',
    "IIf": "Result = IIf(Flag, 1, 2)", "Input": "Result = Input(1, #1)",
    "InputB": "Result = InputB(1, #1)", "InStr": 'Result = InStr(Text, "a")',
    "InStrB": 'Result = InStrB(Text, "a")', "InStrRev": 'Result = InStrRev(Text, "a")',
    "IPmt": "Result = IPmt(0.1, 1, 12, 1000)", "IRR": "Result = IRR(Amounts)",
    "IsArray": "Result = IsArray(Values)", "IsMissing": "Result = IsMissing(Result)",
    "Join": "Result = Join(Other)", "LBound": "Result = LBound(Values)",
    "UBound": "Result = UBound(Values)", "Left": "Result = Left(Text, 1)",
    "LeftB": "Result = LeftB(Text, 1)", "Right": "Result = Right(Text, 1)",
    "RightB": "Result = RightB(Text, 1)", "Mid": "Result = Mid(Text, 1, 1)",
    "MidB": "Result = MidB(Text, 1, 1)", "MIRR": "Result = MIRR(Amounts, 0.1, 0.1)",
    "NPer": "Result = NPer(0.1, -100, 1000)", "NPV": "Result = NPV(0.1, Amounts)",
    "Pmt": "Result = Pmt(0.1, 12, 1000)", "PPmt": "Result = PPmt(0.1, 1, 12, 1000)",
    "PV": "Result = PV(0.1, 12, -100)", "Rate": "Result = Rate(12, -100, 1000)",
    "Replace": 'Result = Replace(Text, "a", "b")', "RGB": "Result = RGB(1, 2, 3)",
    "Round": "Result = Round(Number, 2)", "SLN": "Result = SLN(1000, 100, 5)",
    "Spc": "Print #1, Spc(2); Text", "Tab": "Print #1, Tab(4); Text",
    "StrComp": "Result = StrComp(Text, Text)", "StrConv": "Result = StrConv(Text, vbUpperCase)",
    "String": 'Result = String(3, "x")', "Switch": "Result = Switch(Flag, 1)",
    "SYD": "Result = SYD(1000, 100, 5, 1)", "TimeSerial": "Result = TimeSerial(1, 2, 3)",
    "TimeValue": 'Result = TimeValue("12:00")',
}
# Procedures of the VBA library that are written as statements, with the
# module of the library that holds them, as in VBA.Interaction.Beep.
LIBRARY_STATEMENTS = [
    ("Beep", "", "Interaction"),
    ("SaveSetting", '"App", "Section", "Key", Text', "Interaction"),
    ("DeleteSetting", '"App", "Section", "Key"', "Interaction"),
    ("AppActivate", "Text", "Interaction"), ("SendKeys", '"{ENTER}", True', "Interaction"),
    ("Kill", "Text", "FileSystem"), ("MkDir", "Text", "FileSystem"),
    ("RmDir", "Text", "FileSystem"), ("ChDir", "Text", "FileSystem"),
    ("ChDrive", "Text", "FileSystem"), ("FileCopy", "Text, Text", "FileSystem"),
    ("SetAttr", "Text, vbNormal", "FileSystem"), ("Randomize", "", "Math"),
    ("Randomize", "42", "Math"),
]


# Statements that run code, so that after On Error Resume Next each can fail
# and be the error a read of Err checks.
RESUME_NEXT_OPERATIONS = [
    "Matches = Text Like Pattern", "Flag = A = B", "Flag = Not Flag", "Flag = A And B",
    "Number = Text", "Set Item = Other", "Number = 1", "If A > B Then Number = 1",
    "Number = A Mod B", "Flag = Item Is Nothing", "Flag = False",
]


# A standard module of public error codes and records, as the house style
# prescribes, and a procedure whose handler runs a statement at line 24 and
# then reads Err.Number at line 25.
CODES = """Attribute VB_Name = "TEST_Codes"
Option Explicit
Option Private Module

Public Const LIMIT As Long = 3
Public Enum RunStatus
    rsOk = 0
    rsFailed = 1
End Enum
Public Type SavedError
    Number As Long
    Source As String
    Inner As InnerRecord
End Type
Public Type InnerRecord
    Field As String
End Type
Public LastError As SavedError
Private mHidden As Long


Public Function Describe() As Long
End Function
"""


def qualified(statement: str) -> str:
    """``statement`` in a handler at line 24, then a read of Err.Number."""
    return module("""Private Const OWN As Long = 1
Private Enum Kind
    kOne = 1
End Enum
Private Type Record
    Count As Long
    Items(1 To 3) As Long
    Owner As Object
    Things(1 To 2) As Object
    Data As Variant
End Type
Private mRecord As Record


Private Sub First(ByRef Given As SavedError, ByVal Book As Object, ByVal Loose As Variant)
        Dim Number As Long, Text As String, Saved As SavedError, Many(1 To 2) As SavedError
        Dim Index As Long, Grown() As SavedError
        On Error GoTo Handler
        Exit Sub
Handler:
        """ + statement + """
        Number = Err.Number
End Sub
""")


# Each reads or writes a value, not a member: no finding.
QUALIFIED_VALUES = [
    "Number = RunStatus.rsFailed", "Number = TEST_Codes.RunStatus.rsFailed",
    "Number = TEST_Codes.rsFailed", "Number = Kind.kOne", "Number = TEST_Codes.LIMIT",
    "Number = TEST_Fixture.OWN + TEST_Fixture.kOne", "Number = TEST_Codes.LastError.Number",
    "Err.Raise RunStatus.rsFailed + 0, , Err.Description: Exit Sub",
    "Saved.Number = Err.Number: Saved.Source = Err.Source", "Number = Saved.Number",
    "Saved.Inner.Field = Err.Description", "Many(1).Number = Err.Number",
    "Many(Index).Source = Err.Source", "Number = Given.Number",
    "mRecord.Count = Err.Number: mRecord.Items(1) = 2", "LastError.Source = Err.Source",
    "TEST_Codes.LastError.Number = Err.Number",
    "With Saved: .Number = Err.Number: .Source = Err.Source: End With",
    "Number = mRecord.Items(2)", "ReDim Grown(1 To 2): Grown(1).Number = Err.Number",
    "Text = Saved.Source: Saved.Source = Err.Source",
]
# Each is a member access or a call before the read: one finding.
QUALIFIED_MEMBERS = [
    "Number = Book.Count", "Text = Application.Name", "Number = mRecord.Owner.Count",
    "Number = TEST_Codes.Describe", "Number = Kind.kTwo", "Number = Saved.Missing",
    "Number = TEST_Codes.mHidden", "Number = Many(TEST_Codes.Describe()).Number",
    # An index on anything but an array may call a default property, as
    # Owner(1) does on a Collection; a Variant that holds an array too.
    "Number = mRecord.Owner(1)", "mRecord.Owner(1) = 3", "Number = mRecord.Data(1)",
    "Number = Loose(1).Number", "Number = Saved(1).Number",
    # Reading or assigning an object or a Variant calls its default property,
    # and even Set may create an object declared As New.
    "Number = mRecord.Owner", "mRecord.Owner = 3", "Number = mRecord.Data",
    "Set Book = mRecord.Owner", "Set mRecord.Owner = Nothing",
    "Set Book = mRecord.Things(1)",
]


# The same name declared at two scope levels, with a different kind at each:
# a variable against a property or function, or an Enum against a variable.
# The nearer level decides, as in VBA. Each case is
# (module-level declarations, parameters, local declaration, declaration of
# another standard module, statement in the handler, whether the read of
# Err.Number after it is stale).
SAVED_PROPERTY = "Private Property Get Saved() As Object\nEnd Property"
SAVED_FUNCTION = "Private Function Saved() As Long\nEnd Function"
PRECEDENCE = {
    "local-variable-over-own-property": (SAVED_PROPERTY, "", "Dim Saved As SavedError", "",
                                         "Number = Saved.Number", False),
    "own-property-over-other-variable": (SAVED_PROPERTY, "", "", "Public Saved As SavedError",
                                         "Number = Saved.Number", True),
    "own-property-over-other-variable-bare": (SAVED_PROPERTY, "", "",
                                              "Public Saved As SavedError", "Number = Saved",
                                              True),
    "own-variable-over-other-function": ("Private Saved As SavedError", "", "",
                                         "Public Function Saved() As Object\nEnd Function",
                                         "Number = Saved.Number", False),
    "own-variable-over-other-function-bare": ("Private Saved As Long", "", "",
                                              "Public Function Saved() As Long\nEnd Function",
                                              "Number = Saved", False),
    "parameter-over-own-function": (SAVED_FUNCTION, "ByRef Saved As SavedError", "", "",
                                    "Number = Saved.Number", False),
    "parameter-over-own-function-bare": (SAVED_FUNCTION, "ByVal Saved As Long", "", "",
                                         "Number = Saved", False),
    "own-function-over-other-variable": (SAVED_FUNCTION, "", "", "Public Saved As SavedError",
                                         "Number = Saved.Number", True),
    "own-enum-over-other-variable": ("Private Enum Kind\n    kOne = 1\nEnd Enum", "", "",
                                     "Public Kind As Object", "Number = Kind.kOne", False),
    "own-variable-over-other-enum": ("Private Kind As Object", "", "",
                                     "Public Enum Kind\n    kOne = 1\nEnd Enum",
                                     "Number = Kind.kOne", True),
    "local-variable-over-other-enum": ("", "", "Dim Kind As Object",
                                       "Public Enum Kind\n    kOne = 1\nEnd Enum",
                                       "Number = Kind.kOne", True),
}


def precedence_files(module_level: str, parameters: str, local: str, other: str,
                     statement: str) -> dict[str, str | bytes]:
    """The fixture of one PRECEDENCE case, with TEST_Codes for its Type."""
    body = (f"{module_level}\n\n\nPrivate Sub First({parameters})\n"
            f"        Dim Number As Long\n        {local}\n        On Error GoTo Handler\n"
            f"        Exit Sub\nHandler:\n        {statement}\n        Number = Err.Number\n"
            "End Sub\n")
    files = {**standalone(module(body)), "tests/modules/TEST_Codes.bas": CODES}
    if other:
        files["tests/modules/TEST_Other.bas"] = ('Attribute VB_Name = "TEST_Other"\n'
                                                 "Option Explicit\n" + other + "\n")
    return files


# A Type, Enum or public variable that two components declare may mean
# either: a dotted name through it is a member access, as is a type
# qualified by anything but the component that declares it. Each case is
# (module-level declarations, local declaration, statement in the handler,
# whether TEST_Second is present, whether the read of Err.Number after the
# statement is stale).
LIBRARY = """Attribute VB_Name = "TEST_Library"
Option Explicit
Public Type Payload
    Item As Object
End Type
Public Type Unique
    Item As Long
End Type
Public Saved As Payload
Public Kept As Unique
Public Enum Kind
    kOne = 1
End Enum
"""
SECOND = """Attribute VB_Name = "TEST_Second"
Option Explicit
Public Type Payload
    Item As Long
End Type
Public Enum Kind
    kOne = 1
End Enum
Public Kept As Long
"""
LOCAL_PAYLOAD = "Private Type Payload\n    Item As Long\nEnd Type"
SHADOWING = {
    "qualified-type-shadowed-in-the-module": (LOCAL_PAYLOAD, "Dim Rec As TEST_Library.Payload",
                                              "Number = Rec.Item", False, True),
    "variable-of-a-type-declared-twice": (LOCAL_PAYLOAD, "", "Number = TEST_Library.Saved.Item",
                                          False, True),
    "type-declared-in-two-components": ("", "Dim Rec As Payload", "Number = Rec.Item", True,
                                        True),
    "enum-declared-in-two-components": ("", "", "Number = Kind.kOne", True, True),
    "qualified-enum-declared-twice": ("", "", "Number = TEST_Library.Kind.kOne", True, True),
    "public-variable-in-two-components": ("", "", "Number = Kept.Item", True, True),
    "library-qualified-type": ("", "Dim Rec As VBA.Collection", "Number = Rec.Count", False,
                               True),
    "qualified-unique-type": ("", "Dim Rec As TEST_Library.Unique", "Number = Rec.Item", False,
                              False),
    "unqualified-unique-type": ("", "Dim Rec As Unique", "Number = Rec.Item", False, False),
    "enum-declared-once": ("", "", "Number = Kind.kOne", False, False),
}


def shadowing_files(module_level: str, local: str, statement: str, second: bool
                    ) -> dict[str, str | bytes]:
    """The fixture of one SHADOWING case."""
    body = (f"{module_level}\n\n\nPrivate Sub First()\n        Dim Number As Long\n"
            f"        {local}\n        On Error GoTo Handler\n        Exit Sub\nHandler:\n"
            f"        {statement}\n        Number = Err.Number\nEnd Sub\n")
    files = {**standalone(module(body)), "tests/modules/TEST_Library.bas": LIBRARY}
    if second:
        files["tests/modules/TEST_Second.bas"] = SECOND
    return files


# A Function or Property Get holds its result in its own name: a field of a
# result of a user-defined Type is a value there. Each case is (procedure
# header, its End keyword, statement in the handler, whether the read of
# Err.Number after it is stale).
OWN_RESULT = {
    "function-result-written": ("Public Function TryRun() As SavedError", "Function",
                                "TryRun.Number = Err.Number: TryRun.Source = Err.Source", False),
    "function-result-read": ("Public Function TryRun() As SavedError", "Function",
                             "Number = TryRun.Number", False),
    "property-get-result-written": ("Public Property Get Last() As SavedError", "Property",
                                    "Last.Number = Err.Number: Last.Source = Err.Source",
                                    False),
    "object-result-member": ("Public Function TryRun() As Object", "Function",
                             "Number = TryRun.Owner", True),
    "recursive-call": ("Public Function TryRun(ByVal Retry As Long) As Long", "Function",
                       "Number = TryRun(1)", True),
    "recursive-call-of-a-record": ("Public Function TryRun(ByVal Retry As Long) As SavedError",
                                   "Function", "Number = TryRun(1).Number", True),
}


def own_result(header: str, end: str, statement: str) -> dict[str, str | bytes]:
    """A procedure ``header`` whose handler runs ``statement``, then reads Err.Number."""
    body = (f"{header}\n        Dim Number As Long\n        On Error GoTo Fail\n"
            f"        Exit {end}\nFail:\n        {statement}\n        Number = Err.Number\n"
            f"End {end}\n")
    return {**standalone(module(body)), "tests/modules/TEST_Codes.bas": CODES}


def resume_next_operation(statement: str) -> str:
    """``statement`` after On Error Resume Next, then a read of Err.Number."""
    return ("Private Sub First()\n"
            "        Dim Matches As Boolean, Text As String, Pattern As String, Flag As Boolean\n"
            "        Dim A As Long, B As Long, Number As Long, Item As Object, Other As Object\n"
            "        On Error Resume Next\n"
            f"        {statement}\n        Number = Err.Number\nEnd Sub\n")


def sweep_statements() -> list[str]:
    """Each statement of the sweep, one call of each VBA function, and each
    library statement written plain, with Call, and qualified by VBA. and by
    its module."""
    statements = list(STATEMENT_SWEEP) + list(PURE_CALLS.values())
    for name, arguments, library in LIBRARY_STATEMENTS:
        for qualifier in ("", "VBA.", f"VBA.{library}."):
            statements.append(f"{qualifier}{name} {arguments}".rstrip())
            statements.append(f"Call {qualifier}{name}({arguments})" if arguments
                              else f"Call {qualifier}{name}")
    return statements


def sweep_handler(statement: str) -> str:
    """A handler that runs ``statement`` at line 10, then reads Err.Number
    at line 11; the module header takes lines 1 to 3."""
    return ("Private Sub First()\n"
            "        Dim Number As Long, Text As String, Flag As Boolean, Item As Object\n"
            "        Dim Values() As Long, Other() As String, FileNumber As Integer, "
            "Result As Variant, Amounts() As Double\n"
            "        On Error GoTo Handler\n        Exit Sub\nHandler:\n"
            f"        {statement}\n        Number = Err.Number\nEnd Sub\n")


class ErrorTest(unittest.TestCase):
    def test_valid_fixtures_pass(self) -> None:
        self.assertEqual(run_check(check_vba_errors.check, VALID), [])
        for name, body in ERRORS_VALID.items():
            with self.subTest(name=name):
                self.assertEqual(run_check(check_vba_errors.check, standalone(module(body))), [])

    def test_invalid_fixtures_fail(self) -> None:
        for name, (body, line, text) in ERRORS_INVALID.items():
            with self.subTest(name=name):
                findings = run_check(check_vba_errors.check, standalone(module(body)))
                self.assertEqual(len(findings), 1, findings)
                self.assertTrue(findings[0].startswith(f"tests/modules/TEST_Fixture.bas:{line}: "),
                                findings)
                self.assertIn(text, findings[0])
                self.assertTrue(findings[0].endswith(ALL_ENVIRONMENTS), findings)

    def test_qualified_values_and_type_fields_are_data(self) -> None:
        for statement in QUALIFIED_VALUES:
            with self.subTest(statement=statement):
                files = {**standalone(qualified(statement)), "tests/modules/TEST_Codes.bas": CODES}
                self.assertEqual(run_check(check_vba_errors.check, files), [])
        for statement in QUALIFIED_MEMBERS:
            with self.subTest(statement=statement):
                files = {**standalone(qualified(statement)), "tests/modules/TEST_Codes.bas": CODES}
                findings = run_check(check_vba_errors.check, files)
                self.assertEqual(len(findings), 1, findings)
                self.assertTrue(findings[0].startswith(
                    "tests/modules/TEST_Fixture.bas:25: Err.Number read in First after line 24, "
                    "which may have changed Err with a call or member access"), findings)

    def test_set_of_an_object_declared_as_new_is_a_member_access(self) -> None:
        # Reading a variable declared As New creates it on first use, which
        # runs Class_Initialize, here a call of Err.Clear: even Set reads it.
        owner = ('VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  \'True\nEND\n'
                 'Attribute VB_Name = "COwner"\nOption Explicit\n\n'
                 "Private Sub Class_Initialize()\n        Err.Clear\nEnd Sub\n")
        body = ("Private mOwner As New COwner\n\n\nPrivate Sub First()\n"
                "        Dim Number As Long, Book As Object\n        On Error GoTo Handler\n"
                "        Exit Sub\nHandler:\n        Set Book = TEST_Fixture.mOwner\n"
                "        Number = Err.Number\nEnd Sub\n")
        files = {**standalone(module(body)), "src/classes/COwner.cls": owner}
        findings = run_check(check_vba_errors.check, files)
        self.assertEqual(len(findings), 1, findings)
        self.assertTrue(findings[0].startswith(
            "tests/modules/TEST_Fixture.bas:13: Err.Number read in First after line 12, "
            "which may have changed Err with a call or member access"), findings)

    def test_names_resolve_by_scope_level(self) -> None:
        for name, (module_level, parameters, local, other, statement, stale) in \
                PRECEDENCE.items():
            with self.subTest(name=name):
                files = precedence_files(module_level, parameters, local, other, statement)
                findings = run_check(check_vba_errors.check, files)
                self.assertEqual(bool(findings), stale, findings)
                if stale:
                    self.assertEqual(len(findings), 1, findings)
                    self.assertIn("TEST_Fixture.bas:", findings[0])
                    self.assertIn("Err.Number read in First after line ", findings[0])

    def test_names_declared_twice_are_member_accesses(self) -> None:
        for name, (module_level, local, statement, second, stale) in SHADOWING.items():
            with self.subTest(name=name):
                files = shadowing_files(module_level, local, statement, second)
                findings = run_check(check_vba_errors.check, files)
                self.assertEqual(bool(findings), stale, findings)

    def test_own_result_is_a_variable(self) -> None:
        for name, (header, end, statement, stale) in OWN_RESULT.items():
            with self.subTest(name=name):
                findings = run_check(check_vba_errors.check, own_result(header, end, statement))
                self.assertEqual(bool(findings), stale, findings)

    def test_every_statement_can_fail_after_resume_next(self) -> None:
        for statement in RESUME_NEXT_OPERATIONS:
            with self.subTest(statement=statement):
                files = standalone(module(resume_next_operation(statement)))
                self.assertEqual(run_check(check_vba_errors.check, files), [])
        # Nothing but On Error itself comes before this read.
        findings = run_check(check_vba_errors.check,
                             standalone(module(resume_next_operation("").replace(
                                 "        \n", ""))))
        self.assertEqual(len(findings), 1, findings)
        self.assertIn("Err.Number read in First after line 7, which cleared Err with an On "
                      "Error statement", findings[0])

    def test_vba_statements_and_functions_change_nothing(self) -> None:
        # A table of every VBA statement and function the check knows to run
        # no project code: in a handler, none makes the read after it stale.
        files = {f"tests/modules/TEST_Sweep{index}.bas":
                 module(sweep_handler(statement)).replace("TEST_Fixture", f"TEST_Sweep{index}")
                 for index, statement in enumerate(sweep_statements())}
        findings = run_check(check_vba_errors.check, {**files, "docs/PUBLIC_API.txt": "# None\n"})
        statements = sweep_statements()
        for finding in findings:
            index = int(finding.split(".bas")[0].removeprefix("tests/modules/TEST_Sweep"))
            with self.subTest(statement=statements[index]):
                self.fail(finding)
        self.assertGreater(len(statements), 250)

    def test_sweep_covers_what_the_check_knows(self) -> None:
        # The sweep is checked against the check's own lists, so that a name
        # removed from them, or added without a case, fails here.
        self.assertEqual({name.casefold() for name in PURE_CALLS}, check_vba_errors.PURE_FUNCTIONS)
        statements = sweep_statements()
        for keyword in (*check_vba_errors.BUILTIN_STATEMENTS,
                        *check_vba_errors.BUILTIN_TARGETS):
            with self.subTest(keyword=keyword):
                words = r"\s+".join(keyword.split())
                pattern = re.compile(rf"^(?:Call\s+)?(?:VBA\.(?:\w+\.)?)?{words}\b",
                                     re.IGNORECASE)
                self.assertTrue(any(pattern.match(item) for item in statements), keyword)

    def test_project_declarations_shadow_vba_built_ins(self) -> None:
        # A procedure of the project named like a VBA function, statement,
        # constant or statement target is called instead; VBA.Format still
        # names the VBA function.
        shadows = ("\n\nPrivate Function Format(ByVal Value As Variant) As String\nEnd Function\n"
                   "\n\nPrivate Sub Beep()\nEnd Sub\n"
                   "\n\nPrivate Function vbLimit() As Long\nEnd Function\n"
                   "\n\nPrivate Property Let Mid(ByVal Value As String)\nEnd Property\n")
        other = ('Attribute VB_Name = "TEST_Shadow"\nOption Explicit\nOption Private Module\n'
                 "Public Function Trim(ByVal Value As String) As String\nEnd Function\n")
        # A project module named like a module of the VBA library.
        math = ('Attribute VB_Name = "Math"\nOption Explicit\nOption Private Module\n'
                "Public Function Abs(ByVal Value As Double) As Double\nEnd Function\n")
        cases = {
            "        Text = Format(1)\n        Number = Err.Number\n": 10,
            "        Beep\n        Number = Err.Number\n": 10,
            "        Number = vbLimit\n        Number = Err.Number\n": 10,
            '        Mid = "x"\n        Number = Err.Number\n': 10,
            "        Text = Trim(Text)\n        Number = Err.Number\n": 10,
            "        Text = VBA.Format(1)\n        Number = Err.Number\n": None,
            "        Text = VBA.Trim$(Text)\n        Number = Err.Number\n": None,
            "        Number = Math.Abs(1)\n        Number = Err.Number\n": 10,
            "        Number = VBA.Math.Abs(1)\n        Number = Err.Number\n": None,
            "        Call Beep\n        Number = Err.Number\n": 10,
            "        Call VBA.Beep\n        Number = Err.Number\n": None,
        }
        for statements, line in cases.items():
            with self.subTest(statements=statements):
                files = {**standalone(module(handler(statements) + shadows)),
                         "tests/modules/TEST_Shadow.bas": other, "tests/modules/Math.bas": math}
                findings = run_check(check_vba_errors.check, files)
                if line is None:
                    self.assertEqual(findings, [])
                else:
                    self.assertEqual(len(findings), 1, findings)
                    self.assertTrue(findings[0].startswith(
                        f"tests/modules/TEST_Fixture.bas:{line}: Err.Number read in First after "
                        f"line {line - 1}, which may have changed Err"), findings)

    def test_each_environment_is_analyzed_separately(self) -> None:
        # The call that makes the read stale exists in one #If branch only;
        # the finding names the environments of that branch, once.
        body = handler("#If VBA7 Then\n        Helper\n#Else\n        Number = 0\n#End If\n"
                       "        Report Err.Description\n")
        findings = run_check(check_vba_errors.check, standalone(module(body)))
        self.assertEqual(len(findings), 1, findings)
        self.assertTrue(findings[0].startswith("tests/modules/TEST_Fixture.bas:14: "), findings)
        self.assertIn("after line 10, which may have changed Err", findings[0])
        self.assertTrue(findings[0].endswith("[vba7-win32, vba7-win64]"), findings)
        legacy = body.replace("#If VBA7 Then", "#If Not VBA7 Then")
        findings = run_check(check_vba_errors.check, standalone(module(legacy)))
        self.assertEqual(len(findings), 1, findings)
        self.assertTrue(findings[0].endswith("[vba6-win32]"), findings)

    def test_counts_and_exit_codes(self) -> None:
        invalid = standalone(module(ERRORS_INVALID["after-a-call"][0]))
        for files, expected, counted in (({}, 0, "0 VBA component(s), 0 procedure(s)"),
                                         (VALID, 0, "7 VBA component(s), 6 procedure(s)"),
                                         (invalid, 1, "1 VBA component(s), 3 procedure(s)")):
            with self.subTest(counted=counted), repository(files) as root, \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(check_vba_errors.main(["--root", str(root)]), expected)
                verdict = output.getvalue().splitlines()[-1]
                self.assertIn("PASS" if expected == 0 else "FAIL", verdict)
                self.assertIn(f"({counted}; static check, nothing compiled or run)", verdict)

    def test_repository_vba_has_no_findings(self) -> None:
        findings, _, _ = check_vba_errors.check(ROOT)
        self.assertEqual(findings, [])


class MainTest(unittest.TestCase):
    def test_every_gate_exits_0_when_valid_and_1_when_invalid(self) -> None:
        # Each gate has its own defect here: no Option Explicit, an unknown
        # #If symbol, a jump to a missing label, an implicit public function,
        # a core module that reaches the host and an undeclared assignment.
        invalid = standalone(module("Private Sub First()\n        GoTo Lost\n"
                                    "        Undeclared = 1\nEnd Sub\n"
                                    "Function Implicit()\nEnd Function\n#If Mac Then\n#End If\n"))
        invalid["tests/modules/TEST_Fixture.bas"] = \
            str(invalid["tests/modules/TEST_Fixture.bas"]).replace("Option Explicit\n", "")
        invalid["src/core/CORE_Fixture.bas"] = core_module(
            'Private Sub Warn()\n        MsgBox "x"\nEnd Sub\n')
        mains = (check_vba_source.main, check_vba_conditionals.main, check_vba_jumps.main,
                 check_vba_public_api.main, check_vba_dependencies.main,
                 check_vba_identifiers.main)
        for files, expected in ((VALID, 0), (invalid, 1)):
            with repository(files) as root:
                for main in mains:
                    with self.subTest(main=main.__module__, expected=expected), \
                            contextlib.redirect_stdout(io.StringIO()) as output:
                        self.assertEqual(main(["--root", str(root)]), expected)
                        verdict = output.getvalue().splitlines()[-1]
                        self.assertIn("PASS" if expected == 0 else "FAIL", verdict)
                        self.assertIn("static check, nothing compiled or run", verdict)


if __name__ == "__main__":
    unittest.main()
