"""What tools/devbuild/excel_run.py does without Excel.

The command runs only on Windows with Excel. These tests check the rest: the
host record it writes is accepted by check_excel_run.py and names the
project's references, the record of a run that stops early, the Windows
description, the Excel bitness and Trust access it reads, the evidence folder
it empties, the PowerShell script that drives Excel, which is faked, and how
it reports the end of the Excel instance it started.
"""

from __future__ import annotations

import base64
import contextlib
import io
import json
import re
import shutil
import subprocess
import sys
import tempfile
import types
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))
sys.path.insert(0, str(ROOT / "tools"))

import check_excel_run  # noqa: E402
import excel_run  # noqa: E402

POLICY = check_excel_run.working_policy(ROOT)
# What the PowerShell script prints once Excel has started.
EXCEL_OUTPUT = ("EXCEL-VERSION=16.0\nEXCEL-BUILD=17928\n"
                "EXCEL-PATH=C:\\Program Files\\Microsoft Office\\root\\Office16\n"
                "EXCEL-AUTOMATION-SECURITY=2\n")
TIME = re.compile(r"^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d[+-]\d\d:\d\d$")


def pe_header(machine: int) -> bytes:
    """The start of a Windows executable whose PE header names ``machine``."""
    return b"MZ" + bytes(0x3A) + (0x40).to_bytes(4, "little") + b"PE\0\0" + \
        machine.to_bytes(2, "little")


def fake_winreg(values: dict[str, object], failures: dict[str, type[OSError]]) -> object:
    """A stand-in for winreg: each value is found by a part of its path
    (hive\\key\\name); a path holding a part of ``failures`` raises that error, and
    any other path is missing (FileNotFoundError), as winreg reports them."""
    def query(handle: str, name: str) -> tuple[object, int]:
        path = f"{handle}\\{name}"
        for part, error in failures.items():
            if part in path:
                raise error(5, "Access is denied")
        for part, value in values.items():
            if part in path:
                return value, 4
        raise FileNotFoundError(2, "The system cannot find the file specified")
    return types.SimpleNamespace(
        HKEY_CURRENT_USER="HKEY_CURRENT_USER", HKEY_LOCAL_MACHINE="HKEY_LOCAL_MACHINE",
        OpenKey=lambda hive, key: contextlib.nullcontext(f"{hive}\\{key}"), QueryValueEx=query)


def read_host(path: Path) -> dict[str, str]:
    return dict(line.split("=", 1) for line in path.read_text(encoding="utf-8").splitlines())


class HostRecordTest(unittest.TestCase):
    def test_the_built_references_are_the_policy_references(self) -> None:
        self.assertEqual(sorted(excel_run.built_references()), sorted(POLICY.references))

    def test_the_record_passes_the_check_only_with_identity_by_rebuild(self) -> None:
        record = excel_run.host_record({
            "operator": "someone", "started": "2026-11-03T14:05:09+01:00",
            "finished": "2026-11-03T14:07:41+01:00",
            "windows": "Windows 11 Pro, version 24H2, OS build 26100.4061",
            "excel": "16.0 build 17928", "bitness": "64-bit", "automation_security": "ByUI",
            "vba_project_access": "NOT_TRUSTED",
            "references": "; ".join(excel_run.built_references()),
            "compile": check_excel_run.COMPILE_ON_DEMAND, "outcome": "PASS"})
        self.assertIn("compile=ON_DEMAND", record)
        self.assertEqual([line.split("=")[0] for line in record.splitlines()],
                         list(check_excel_run.HOST_FIELDS))
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            (folder / POLICY.host).write_text(record, encoding="utf-8")
            for rebuilt, expected in ((True, []), (False, ["compile=ON_DEMAND"])):
                with self.subTest(rebuilt=rebuilt):
                    evidence = check_excel_run.Evidence()
                    check_excel_run.check_host(folder, evidence, POLICY, rebuilt)
                    self.assertEqual(len(evidence.findings), len(expected), evidence.findings)
                    for text, finding in zip(expected, evidence.findings):
                        self.assertIn(text, finding)

    def test_windows_is_described_as_winver_shows_it(self) -> None:
        values = {"ProductName": "Windows 10 Pro", "DisplayVersion": "24H2",
                  "CurrentBuild": "26100", "UBR": 4061}
        self.assertEqual(excel_run.describe_windows(values),
                         "Windows 11 Pro, version 24H2, OS build 26100.4061")
        older = {"ProductName": "Windows 10 Enterprise", "DisplayVersion": "22H2",
                 "CurrentBuild": "19045", "UBR": 5247}
        self.assertEqual(excel_run.describe_windows(older),
                         "Windows 10 Enterprise, version 22H2, OS build 19045.5247")
        self.assertEqual(excel_run.describe_windows({"ProductName": "Windows 10 Pro"}),
                         "Windows 10 Pro")

    def test_a_value_not_read_is_unknown(self) -> None:
        record = excel_run.host_record({"operator": "someone", "outcome": "UNAVAILABLE"})
        self.assertIn("excel=UNKNOWN\n", record)
        self.assertIn("automation_security=UNKNOWN\n", record)

    def test_the_bitness_is_read_from_the_executable_header(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "EXCEL.EXE"
            for data, expected in ((pe_header(0x8664), "64-bit"), (pe_header(0x014C), "32-bit"),
                                   (pe_header(0x0200), "UNKNOWN"), (b"not a program", "UNKNOWN")):
                with self.subTest(expected=expected):
                    path.write_bytes(data)
                    self.assertEqual(excel_run.executable_bitness(path), expected)
            self.assertEqual(excel_run.executable_bitness(Path(temporary) / "absent.exe"),
                             "UNKNOWN")

    def test_trust_access_is_read_with_a_policy_first(self) -> None:
        user = r"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM"
        policy = (r"HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\Excel"
                  r"\Security\AccessVBOM")
        for values, expected in (({}, "NOT_TRUSTED"), ({user: 1}, "TRUSTED"),
                                 ({user: 0}, "NOT_TRUSTED"), ({user: 1, policy: 0}, "NOT_TRUSTED")):
            with self.subTest(values=values):
                def read(hive: str, key: str, name: str,
                         values: dict[str, int] = values) -> object | None:
                    return values.get(f"{hive}\\{key}\\{name}")
                self.assertEqual(excel_run.vba_project_access("16.0", read), expected)

    def test_a_failed_registry_read_is_unknown_never_a_lower_setting(self) -> None:
        # Only a missing key or value falls through to the next location; a read that
        # fails at any of them leaves the setting unknown.
        for failing in ("HKEY_CURRENT_USER\\Software\\Policies", "HKEY_LOCAL_MACHINE",
                        "HKEY_CURRENT_USER\\Software\\Microsoft"):
            with self.subTest(failing=failing):
                with mock.patch.dict(sys.modules, {"winreg": fake_winreg(
                        {r"HKEY_CURRENT_USER\Software\Microsoft": 1},
                        {failing: PermissionError})}):
                    self.assertEqual(excel_run.vba_project_access("16.0"), "UNKNOWN")
        with mock.patch.dict(sys.modules, {"winreg": fake_winreg({}, {"HKEY": PermissionError})}):
            self.assertEqual(excel_run.vba_project_access("16.0"), "UNKNOWN")
        with mock.patch.dict(sys.modules, {"winreg": fake_winreg(
                {r"HKEY_CURRENT_USER\Software\Microsoft": 1}, {})}):
            self.assertEqual(excel_run.vba_project_access("16.0"), "TRUSTED")

    def test_a_failed_windows_read_is_unknown_not_a_shorter_description(self) -> None:
        values = {"ProductName": "Windows 10 Pro", "CurrentBuild": "26100", "UBR": 4061}
        with mock.patch.dict(sys.modules, {"winreg": fake_winreg(values, {})}):
            # A value that is absent is left out, as on a Windows without DisplayVersion.
            self.assertEqual(excel_run.windows_found(), "Windows 11 Pro, OS build 26100.4061")
        with mock.patch.dict(sys.modules, {"winreg": fake_winreg(
                values, {"UBR": PermissionError})}):
            self.assertEqual(excel_run.windows_found(), "UNKNOWN")

    def test_the_excel_fields_come_from_what_the_script_prints(self) -> None:
        with mock.patch.object(excel_run, "executable_bitness", return_value="64-bit") as bits, \
                mock.patch.object(excel_run, "vba_project_access", return_value="NOT_TRUSTED"):
            fields = excel_run.excel_host("noise\n" + EXCEL_OUTPUT)
        self.assertEqual(fields, {"excel": "16.0 build 17928", "bitness": "64-bit",
                                  "automation_security": "ByUI",
                                  "vba_project_access": "NOT_TRUSTED"})
        self.assertEqual(bits.call_args.args[0].name, "EXCEL.EXE")
        self.assertEqual(excel_run.excel_host("EXCEL-AUTOMATION-SECURITY=9"),
                         {"excel": "", "bitness": "", "automation_security": ""})

    def test_how_excel_ended_comes_from_what_the_script_prints(self) -> None:
        for result in excel_run.QUIT_RESULTS:
            with self.subTest(result=result):
                output = EXCEL_OUTPUT + f"EXCEL-QUIT={result}\r\nEXCEL-RUN=DONE\n"
                self.assertEqual(excel_run.excel_quit(output), result)
                self.assertTrue(excel_run.quit_line(result).startswith(
                    f"excel-quit: {result} ("))
        # No line (Excel did not start, or PowerShell stopped first) or another value.
        self.assertEqual(excel_run.excel_quit(EXCEL_OUTPUT), "UNKNOWN")
        self.assertEqual(excel_run.excel_quit("EXCEL-QUIT=KILLED"), "UNKNOWN")
        self.assertIn("close it yourself", excel_run.quit_line("SURVIVED"))
        # The line is not a field of the host record.
        self.assertEqual(excel_run.excel_host("EXCEL-QUIT=STOPPED"),
                         {"excel": "", "bitness": "", "automation_security": ""})

    def test_the_report_holds_how_excel_ended_after_the_outcome(self) -> None:
        report = "Candidate commit: abc\nOutcome: PASS\nBinding: VERIFIED\nexcel-run: PASS"
        self.assertEqual(excel_run.with_quit(report, "STOPPED").splitlines(),
                         ["Candidate commit: abc", "Outcome: PASS",
                          excel_run.quit_line("STOPPED"), "Binding: VERIFIED",
                          "excel-run: PASS"])

    def test_the_tool_reads_the_macro_settings_and_never_writes_them(self) -> None:
        source = (ROOT / "tools/devbuild/excel_run.py").read_text(encoding="utf-8")
        for call in ("SetValue", "CreateKey", "DeleteValue", "DeleteKey"):
            self.assertNotIn(call, source)


class EvidenceFolderTest(unittest.TestCase):
    def test_a_folder_under_test_results_is_emptied(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            folder = root / "test-results" / "excel-run"
            (folder / "exported").mkdir(parents=True)
            (folder / "examples.log").write_text("old", encoding="utf-8")
            excel_run.prepare(folder, root)
            self.assertEqual(list(folder.iterdir()), [])

    def test_another_non_empty_folder_is_refused_and_kept(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            folder = root / "elsewhere"
            folder.mkdir()
            (folder / "keep.txt").write_text("mine", encoding="utf-8")
            with self.assertRaisesRegex(excel_run.RunError, "not empty"):
                excel_run.prepare(folder, root)
            self.assertTrue((folder / "keep.txt").is_file())
            new = root / "new"
            excel_run.prepare(new, root)
            self.assertTrue(new.is_dir())


class PowerShellTest(unittest.TestCase):
    script = excel_run.powershell_script(Path(r"C:\Dev\it's\Workbook-development.xlsm"),
                                         Path(r"C:\Dev\run"), Path(r"C:\Dev\tests\fixtures"),
                                         [f"TEST_Examples.{POLICY.entry_point}"],
                                         "TEST_Harness.LastResult", 300)

    def test_the_workbook_is_opened_read_only_and_never_saved(self) -> None:
        self.assertIn("$excel.Workbooks.Open($workbookPath, 0, $true)", self.script)
        self.assertIn("$book.Close($false)", self.script)
        self.assertIn("$excel.Quit()", self.script)
        self.assertNotIn(".Save", self.script)
        # Nothing reaches the VBA project, so no Trust access is needed.
        self.assertNotIn("VBProject", self.script)

    def test_no_macro_setting_changes_and_a_person_enables_the_macros(self) -> None:
        # A policy may open the workbook with macros disabled (HB-02); the script never
        # overrides it, and waits for Enable Content instead. It only reads the setting.
        self.assertIsNone(re.search(r"AutomationSecurity\s*=", self.script, re.I))
        self.assertIn('Write-Output ("EXCEL-AUTOMATION-SECURITY=" + [int]$excel.'
                      'AutomationSecurity)', self.script)
        self.assertNotIn("AccessVBOM", self.script)
        probe = self.script.index("[void]$excel.Run($prefix + 'TEST_Harness.LastResult')")
        opened = self.script.index("$excel.Workbooks.Open(")
        run = self.script.index(f"$excel.Run($prefix + 'TEST_Examples.{POLICY.entry_point}'")
        self.assertLess(opened, probe)
        self.assertLess(probe, run)
        self.assertIn("$deadline = (Get-Date).AddSeconds(300)", self.script)
        self.assertIn("Start-Sleep -Seconds 2", self.script)
        self.assertIn('$failure = "NOT-ENABLED"', self.script)
        self.assertIn("exit 4", self.script)

    def test_each_way_the_run_ends_has_its_exit_code(self) -> None:
        for code in excel_run.OUTCOME_BY_EXIT:
            if code:
                self.assertIn(f"exit {code}", self.script)
        self.assertEqual(set(excel_run.OUTCOME_BY_EXIT.values()) - set(check_excel_run.OUTCOMES),
                         set())
        self.assertIn('if ($null -eq $excel) { "UNAVAILABLE" } else { "INCOMPLETE" }',
                      self.script)

    def test_the_process_of_the_instance_it_starts_is_recorded(self) -> None:
        # From the window of the instance New-Object created, never by a process name,
        # and before the workbook opens.
        created = self.script.index("$excel = New-Object -ComObject Excel.Application")
        window = self.script.index("$window = [IntPtr][int64]$excel.Hwnd")
        found = self.script.index(
            "[void][ExcelRunWindow]::GetWindowThreadProcessId($window, [ref]$windowProcess)")
        start = self.script.index(
            "$excelStart = (Get-Process -Id ([int]$windowProcess)).StartTime")
        opened = self.script.index("$excel.Workbooks.Open(")
        self.assertLess(created, window)
        self.assertLess(window, found)
        self.assertLess(found, start)
        self.assertLess(start, opened)
        self.assertIn("public static extern uint GetWindowThreadProcessId(IntPtr window, "
                      "out uint processId);", self.script)
        # A failure to read them leaves them unknown and does not stop the run.
        self.assertIn("} catch {\n        $excelId = $null\n        $excelStart = $null\n"
                      "    }", self.script)

    def test_after_quit_only_the_recorded_process_is_ended_past_the_wait(self) -> None:
        quit_ = self.script.index("$excel.Quit()")
        released = self.script.index("FinalReleaseComObject($excel)")
        wait = self.script.index(f"$process.WaitForExit({excel_run.QUIT_WAIT * 1000})")
        guard = self.script.index("if ($process.ProcessName -ne 'EXCEL' -or "
                                  "$process.StartTime -ne $excelStart) {")
        stop = self.script.index("Stop-Process -InputObject $process -Force")
        self.assertLess(quit_, released)
        self.assertLess(released, wait)
        self.assertLess(wait, guard)
        self.assertLess(guard, stop)
        self.assertIn("$process = Get-Process -Id $excelId -ErrorAction SilentlyContinue",
                      self.script)
        self.assertIn("$process.Refresh()", self.script)
        self.assertIn(f"$ended = $process.WaitForExit({excel_run.STOP_WAIT * 1000})",
                      self.script)
        # The workbook is released too, so that no reference keeps Excel running.
        self.assertIn("FinalReleaseComObject($book)", self.script)
        # The wait runs even when closing failed.
        self.assertIn("    } finally {\n        # Runs even when closing failed.", self.script)
        # One process, by its id: never by a name, never all Excel processes.
        self.assertEqual(self.script.count("Stop-Process"), 1)
        for call in ("-Name", "taskkill", "Get-Process EXCEL", "Get-Process -ProcessName",
                     ".Kill("):
            self.assertNotIn(call, self.script)
        self.assertIn('if ($started) { Write-Output ("EXCEL-QUIT=" + $quit) }', self.script)

    def test_the_wait_after_quit_is_bounded_by_its_argument(self) -> None:
        script = excel_run.powershell_script(Path(r"C:\Dev\Workbook-development.xlsm"),
                                             Path(r"C:\Dev\run"), Path(r"C:\Dev\fixtures"),
                                             ["TEST_Examples.CertifyExamples"],
                                             "TEST_Harness.LastResult", 300, quit_wait=5)
        self.assertIn("$process.WaitForExit(5000)", script)
        self.assertNotIn(f"WaitForExit({excel_run.QUIT_WAIT * 1000})", script)

    def test_the_guides_state_the_wait_after_quit(self) -> None:
        for guide in ("SECURITY.md", "INSTALLATION.md"):
            with self.subTest(guide=guide):
                text = " ".join((ROOT / guide).read_text(encoding="utf-8").split())
                self.assertIn(f"still running {excel_run.QUIT_WAIT} seconds", text)

    def test_every_result_the_script_prints_is_known(self) -> None:
        tail = self.script[self.script.index('$quit = "UNKNOWN"'):]
        printed = set(re.findall(r'\$quit = "([A-Z]+)"', tail)) | set(
            re.findall(r'\{ "([A-Z]+)" \}', tail))
        self.assertEqual(printed, set(excel_run.QUIT_RESULTS))

    def test_progress_records_do_not_reach_the_output(self) -> None:
        self.assertIn("$ProgressPreference = 'SilentlyContinue'", self.script)

    def test_the_entry_point_runs_with_both_folders(self) -> None:
        self.assertIn(f"$excel.Run($prefix + 'TEST_Examples.{POLICY.entry_point}', $evidence, "
                      "$fixtures)", self.script)
        self.assertIn(r"$evidence = 'C:\Dev\run'", self.script)
        self.assertIn(r"$fixtures = 'C:\Dev\tests\fixtures'", self.script)
        # A quote in a path is doubled, so a path cannot end the literal.
        self.assertIn(r"$workbookPath = 'C:\Dev\it''s\Workbook-development.xlsm'", self.script)

    def test_every_suite_runs_even_after_one_fails(self) -> None:
        script = excel_run.powershell_script(Path(r"C:\Dev\Workbook-development.xlsm"),
                                             Path(r"C:\Dev\run"), Path(r"C:\Dev\fixtures"),
                                             ["TEST_Examples.CertifyExamples",
                                              "TEST_Project.CertifyProject"],
                                             "TEST_Harness.LastResult", 300)
        first = script.index("$excel.Run($prefix + 'TEST_Examples.CertifyExamples'")
        second = script.index("$excel.Run($prefix + 'TEST_Project.CertifyProject'")
        self.assertLess(first, second)
        # Each call has its own catch, and only the first failure is kept.
        self.assertEqual(script.count("if (-not $failure) { $failure = "), 2)
        # Added after the suites the working policy already lists, whatever they are.
        data = json.loads((ROOT / check_excel_run.POLICY).read_text(encoding="utf-8"))
        listed = [f"{Path(item['source']).stem}.{item['certification_entry_point']}"
                  for item in data["suites"]]
        data["suites"].append({
            "name": "added-project", "source": "tests/modules/TEST_AddedProject.bas",
            "certification_entry_point": "CertifyAddedProject", "expected_cases": 1,
            "expected_assertions": 1, "log": "added-project.log",
            "certification_log": "added-project-certification.log"})
        policy = check_excel_run.parse_policy(json.dumps(data))
        self.assertEqual(excel_run.entry_points(policy),
                         [f"{Path(POLICY.suite_source).stem}.{POLICY.entry_point}", *listed,
                          "TEST_AddedProject.CertifyAddedProject"])

    def test_the_command_is_encoded_and_skips_the_profile(self) -> None:
        text = excel_run.encoded(self.script)
        self.assertEqual(base64.b64decode(text).decode("utf-16-le"), self.script)
        with mock.patch.object(subprocess, "run") as run:
            run.return_value = subprocess.CompletedProcess([], 0, "EXCEL-RUN=DONE", "")
            self.assertEqual(excel_run.run_excel(self.script, 5), (0, "EXCEL-RUN=DONE"))
        command = run.call_args.args[0]
        self.assertEqual(command[:4], ["powershell.exe", "-NoProfile", "-NonInteractive",
                                       "-EncodedCommand"])
        self.assertNotIn("Bypass", command)


class CommandTest(unittest.TestCase):
    def stopped(self, run: mock.Mock, error: str, outcome: str,
                quit: str = "UNKNOWN") -> dict[str, str]:
        """The host record a run that stops early leaves, with PowerShell faked by ``run``;
        the output says how Excel ended (``quit``)."""
        built = mock.Mock(commit="a" * 40, clean=True)
        folder = ROOT / "test-results" / "excel-run-unit-test"
        try:
            with mock.patch("excel_run.build_workbook.build", return_value=built), \
                    mock.patch.object(subprocess, "run", run), \
                    mock.patch.object(excel_run, "windows_version",
                                      return_value="Windows 11 Pro, version 24H2"), \
                    mock.patch.object(excel_run, "executable_bitness", return_value="64-bit"), \
                    mock.patch.object(excel_run, "vba_project_access",
                                      return_value="NOT_TRUSTED"), \
                    mock.patch("excel_run.check_excel_run.check") as check, \
                    redirect_stdout(io.StringIO()) as output:
                with self.assertRaisesRegex(excel_run.RunError, error):
                    excel_run.record(ROOT, folder, 600, 300, "Operator Name")
            check.assert_not_called()
            self.assertIn("click Enable Content", output.getvalue())
            self.assertIn(f"outcome={outcome}", output.getvalue())
            self.assertIn(excel_run.quit_line(quit), output.getvalue())
            host = read_host(folder / POLICY.host)
            evidence = check_excel_run.Evidence()
            check_excel_run.check_host(folder, evidence, POLICY, rebuilt=True)
        finally:
            shutil.rmtree(folder, ignore_errors=True)
        self.assertEqual(list(host), list(check_excel_run.HOST_FIELDS))
        self.assertEqual((host["outcome"], host["operator"]), (outcome, "Operator Name"))
        self.assertRegex(host["started"], TIME)
        self.assertRegex(host["finished"], TIME)
        self.assertEqual(host["windows"], "Windows 11 Pro, version 24H2")
        # The record is valid, and evidence of its outcome only.
        self.assertEqual(evidence.findings, [f"host.txt: outcome={outcome}: the record is "
                                             f"evidence of a run that ended {outcome}, never "
                                             "of a pass"])
        return host

    def test_macros_never_enabled_leave_a_record_of_that_outcome(self) -> None:
        run = mock.Mock(return_value=subprocess.CompletedProcess(
            [], excel_run.NOT_ENABLED,
            EXCEL_OUTPUT + "EXCEL-QUIT=EXITED\nEXCEL-RUN=NOT-ENABLED", ""))
        host = self.stopped(run, "not enabled within 300 s", "NOT_ENABLED", "EXITED")
        self.assertEqual((host["excel"], host["bitness"], host["automation_security"],
                          host["vba_project_access"]),
                         ("16.0 build 17928", "64-bit", "ByUI", "NOT_TRUSTED"))

    def test_a_timeout_leaves_a_record_of_that_outcome(self) -> None:
        run = mock.Mock(side_effect=subprocess.TimeoutExpired([], 600, output=EXCEL_OUTPUT))
        host = self.stopped(run, "did not finish within 600 s", "EXECUTION_TIMEOUT")
        self.assertEqual(host["automation_security"], "ByUI")

    def test_excel_or_powershell_unavailable_leaves_a_record(self) -> None:
        missing = mock.Mock(side_effect=FileNotFoundError("powershell.exe"))
        host = self.stopped(missing, "PowerShell could not be started", "UNAVAILABLE")
        self.assertEqual(host["excel"], "UNKNOWN")
        no_excel = mock.Mock(return_value=subprocess.CompletedProcess(
            [], 5, "EXCEL-RUN=UNAVAILABLE; class not registered", ""))
        self.stopped(no_excel, "exit 5, outcome UNAVAILABLE", "UNAVAILABLE")
        opened = mock.Mock(return_value=subprocess.CompletedProcess(
            [], 6, EXCEL_OUTPUT + "EXCEL-RUN=INCOMPLETE; cannot open", ""))
        self.stopped(opened, "exit 6, outcome INCOMPLETE", "INCOMPLETE")

    def test_a_completed_run_prints_and_saves_how_excel_ended(self) -> None:
        built = mock.Mock(commit="a" * 40, clean=True)
        folder = ROOT / "test-results" / "excel-run-quit-unit-test"
        saved = folder.parent / f"{folder.name}-report.txt"
        run = mock.Mock(return_value=subprocess.CompletedProcess(
            [], 0, EXCEL_OUTPUT + "EXCEL-QUIT=STOPPED\nEXCEL-RUN=DONE", ""))
        evidence = check_excel_run.Evidence(host={"outcome": "PASS"}, binding="VERIFIED")
        try:
            with mock.patch("excel_run.build_workbook.build", return_value=built), \
                    mock.patch.object(subprocess, "run", run), \
                    mock.patch.object(excel_run, "windows_version", return_value="Windows 11"), \
                    mock.patch.object(excel_run, "executable_bitness", return_value="64-bit"), \
                    mock.patch.object(excel_run, "vba_project_access",
                                      return_value="NOT_TRUSTED"), \
                    mock.patch("excel_run.check_excel_run.check", return_value=evidence), \
                    redirect_stdout(io.StringIO()) as output:
                self.assertEqual(excel_run.record(ROOT, folder, 600, 300, "Operator Name"), 0)
            report = saved.read_text(encoding="utf-8").splitlines()
        finally:
            shutil.rmtree(folder, ignore_errors=True)
            saved.unlink(missing_ok=True)
        line = excel_run.quit_line("STOPPED")
        self.assertIn(line, output.getvalue())
        self.assertEqual(report[report.index(line) - 1], "Outcome: PASS")
        self.assertTrue(report[-1].startswith("excel-run: PASS"))

    def test_the_wait_for_enable_content_is_shorter_than_the_run(self) -> None:
        with redirect_stdout(io.StringIO()), mock.patch("sys.stderr", io.StringIO()), \
                self.assertRaises(SystemExit):
            excel_run.main(["--timeout", "300", "--enable-timeout", "300"])

    @unittest.skipIf(sys.platform == "win32", "checks the refusal off Windows")
    def test_off_windows_it_records_nothing(self) -> None:
        output = io.StringIO()
        with redirect_stdout(output):
            self.assertEqual(excel_run.main([]), 2)
        self.assertIn("SKIP", output.getvalue())

    def test_a_build_that_is_not_evidence_grade_stops_before_excel(self) -> None:
        built = mock.Mock(commit="a" * 40, clean=False)
        with mock.patch("excel_run.build_workbook.build", return_value=built), \
                mock.patch.object(excel_run, "run_excel") as run, \
                redirect_stdout(io.StringIO()):
            with self.assertRaisesRegex(excel_run.RunError, "local changes"):
                excel_run.record(ROOT, ROOT / "test-results" / "excel-run-test", 5)
        run.assert_not_called()


if __name__ == "__main__":
    unittest.main()
