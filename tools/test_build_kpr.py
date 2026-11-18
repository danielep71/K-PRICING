#!/usr/bin/env python3
"""Tests for tools/Build-KPR.ps1 and Build-KPR.cmd: shared helpers, roles and launcher."""

from __future__ import annotations

import json
import re
import unittest

import workbook_template as wt
from inject_ribbon import ROOT

TOOLS = ROOT / "tools"
BUILD = (TOOLS / "Build-KPR.ps1").read_text(encoding="ascii").replace("\r\n", "\n")
CHECK = (TOOLS / "Check-KPR.ps1").read_text(encoding="ascii").replace("\r\n", "\n")
LAUNCHER = (TOOLS / "Build-KPR.cmd").read_text(encoding="ascii").replace("\r\n", "\n")

# The Excel helpers both scripts need. A script run through the one-line
# command is a single file, so the helpers are copied, and kept identical here.
SHARED = (
    "Say",
    "Get-RegistryValue",
    "Get-OfficeVersion",
    "Confirm-VbaProjectAccess",
    "Get-ExcelProcessId",
    "Test-WorkbookRibbon",
    "New-FromTemplate",
    "Invoke-Compile",
)


def function_block(text: str, name: str) -> str:
    match = re.search(rf"^function {re.escape(name)} \{{\n.*?^\}}\n", text, re.S | re.M)
    if match is None:
        raise AssertionError(f"function {name} not found")
    return match.group(0)


def constant(text: str, name: str) -> str:
    match = re.search(rf"^\${name} = (.+)$", text, re.M)
    if match is None:
        raise AssertionError(f"${name} not found")
    return match.group(1).strip()


class BuildScriptTests(unittest.TestCase):
    def test_shared_helpers_are_identical_in_both_scripts(self) -> None:
        for name in SHARED:
            with self.subTest(function=name):
                self.assertEqual(function_block(BUILD, name), function_block(CHECK, name))

    def test_shared_constants_agree(self) -> None:
        for name in ("xlOpenXMLWorkbookMacroEnabled", "compileControlId", "templateCell"):
            with self.subTest(constant=name):
                self.assertEqual(constant(BUILD, name), constant(CHECK, name))
        self.assertEqual(constant(BUILD, "templateCell"), f'"{wt.COMMIT_CELL}"')

    def test_roles_follow_the_repository_profile(self) -> None:
        profile = json.loads((ROOT / ".github" / "repository-profile.json").read_text(encoding="utf-8"))
        every = set(profile["vba"]["components"].values())
        match = re.search(
            r'\$roles = \$\(if \(\$Mode -eq "Development"\) \{ "([a-z,]+)" \} else \{ "([a-z,]+)" \}\)', BUILD
        )
        assert match is not None
        self.assertEqual(set(match.group(1).split(",")), every)
        self.assertEqual(set(match.group(2).split(",")), every - {"test"})

    def test_parameters_and_modes(self) -> None:
        self.assertIn('[ValidateSet("Production", "Development")]', BUILD)
        for name in ("Mode", "RepoRoot", "TimeoutMinutes", "Open", "NoRibbonCheck", "NoPause"):
            self.assertRegex(BUILD, rf"\[(?:string|int|switch)\]\${name}\b")

    def test_no_loop_variable_shadows_a_parameter(self) -> None:
        # PowerShell variables ignore case: `foreach ($open ...)` would overwrite `-Open`.
        parameters = {name.casefold() for name in re.findall(r"\]\$(\w+)\b", BUILD.split(")\n\nSet-StrictMode")[0])}
        loops = {name.casefold() for name in re.findall(r"foreach \(\$(\w+) in", BUILD)}
        self.assertTrue(parameters)
        self.assertFalse(parameters & loops, parameters & loops)

    def test_the_build_reads_the_clone_and_runs_no_regression(self) -> None:
        self.assertIn("KPR_Dev_ImportModules", BUILD)
        self.assertIn("KPR_Demo_BuildDates", BUILD)
        self.assertIn('Join-Path $root "src\\workbook\\KPR_Template.xlsx"', BUILD)
        self.assertNotIn('"KPR_Test_RunAll"', BUILD)
        self.assertNotIn("KPR_Dev_CheckBuild", BUILD)
        self.assertNotIn("raw.githubusercontent.com", BUILD)
        self.assertNotIn("api.github.com", BUILD)

    def test_launcher_runs_the_script_with_the_clone_and_both_modes(self) -> None:
        self.assertIn('set "SCRIPT=%~dp0Build-KPR.ps1"', LAUNCHER)
        self.assertIn('-File "%SCRIPT%" -Mode "%MODE%" -RepoRoot "%REPO_ROOT%" -NoPause', LAUNCHER)
        self.assertIn('set "MODE=Production"', LAUNCHER)
        for mode in ("Production", "Development"):
            self.assertIn(f'if /i not "%MODE%"=="{mode}"', LAUNCHER)

    def test_files_are_ascii_with_crlf_in_the_working_tree(self) -> None:
        for name in ("Build-KPR.ps1", "Build-KPR.cmd"):
            with self.subTest(file=name):
                raw = (TOOLS / name).read_bytes()
                raw.decode("ascii")
                self.assertEqual(raw.count(b"\n"), raw.count(b"\r\n"))


if __name__ == "__main__":
    unittest.main()
