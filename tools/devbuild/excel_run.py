#!/usr/bin/env python3
"""Record an Excel run of the harness certification with one command, on Windows.

Usage: python tools/devbuild/excel_run.py [--evidence FOLDER] [--timeout SECONDS]
                                         [--enable-timeout SECONDS] [--operator NAME]

Run it in a clean checkout of the candidate commit, on Windows with Excel
installed. It replaces the steps of a run by hand
(docs/repo_governance/EXCEL_RUN.md), and never needs Trust access to the VBA
project object model:

1. builds the development workbook with build_workbook.py, and stops if the
   build is not evidence-grade (a change in the checkout);
2. empties the evidence folder (default test-results/excel-run, which Git
   ignores; a folder under test-results/ is emptied, any other must be empty
   or absent);
3. opens the workbook read-only in a visible Excel through COM, from
   PowerShell started with -NoProfile and -EncodedCommand. It changes no
   macro setting: where Excel opens it with macros disabled, as under a
   policy that makes programs follow the user's macro setting
   (VALIDATION_PLAN.md, HB-02), you click Enable Content once, which enables
   them without reopening the workbook (HB-04), and the command waits for
   that, up to --enable-timeout seconds. Then it runs the certification entry
   point of the example suite and of each project suite the policy lists, in
   order, qualified with the workbook's name so that Application.Run reaches
   an Option Private Module module (HB-03), with the evidence folder and
   tests/fixtures, closes the workbook without saving and quits Excel. It
   records the process id and start time of the Excel instance it started,
   and if that instance is still running QUIT_WAIT seconds after it quits,
   it ends that process, never another one: a process whose name or start
   time differs is left alone. It prints how Excel ended (excel-quit:), and
   the report records it;
4. writes host.txt, the host record of docs/methodology/TEST_HARNESS.md:
   the operator (--operator, by default the Windows user), when the run
   started and finished, Windows from the registry, the Excel build, its
   bitness (from the EXCEL.EXE header), AutomationSecurity as Excel reports
   it (read, never set), the Trust access to the VBA project as the registry
   holds it, the references the build gives the project, compile=ON_DEMAND,
   because no automated run can compile the whole project without Trust
   access, so VBA compiled only the modules the run called (check_excel_run.py
   accepts it only with identity by rebuild, as step 5 checks it), and the
   outcome. It writes the record also when the run stops early: Excel
   unavailable, macros not enabled in time, a timeout, or a run that stopped
   before its end;
5. runs tools/check_excel_run.py with --workbook and --require-binding: with
   no exported folder it proves identity by rebuilding the commit and
   comparing the workbook's parts. It prints the report, and saves it
   beside the evidence folder as <folder>-report.txt, with how Excel ended.

It exits 0 when the check passes, 1 when it fails or the run could not
complete, and 2 for an invalid option, or when it is not run on Windows.
Standard library only.
It reads the macro setting and the Trust access only; it changes neither.
"""

from __future__ import annotations

import argparse
import base64
import getpass
import os
import shutil
import subprocess
import sys
from collections.abc import Callable, Mapping, Sequence
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(ROOT / "tools"))

import build_workbook  # noqa: E402
import check_excel_run  # noqa: E402
import vba_project  # noqa: E402

DEFAULT_EVIDENCE = Path("test-results") / "excel-run"
WORKBOOK = Path("test-results") / "build" / "Workbook-development.xlsm"
FIXTURES = Path("tests") / "fixtures"
# References every VBA project in Excel has, besides those the build writes.
IMPLICIT_REFERENCES = ("Visual Basic For Applications", "Microsoft Excel 16.0 Object Library")
WINDOWS_KEY = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion"
# Windows 11 keeps "Windows 10" in ProductName; its builds start at 22000.
WINDOWS_11_BUILD = 22000
# Seconds to wait, after Quit, for the Excel instance the run started to exit
# before its process is ended; then seconds to wait for it to end.
QUIT_WAIT = 30
STOP_WAIT = 10
# How the Excel instance the run started ended, as the script prints it
# (EXCEL-QUIT=): what the command then tells you.
QUIT_RESULTS = {
    "EXITED": "Excel exited after Quit",
    "STOPPED": (f"Excel was still running {QUIT_WAIT} s after Quit, so the run ended the "
                "process of the instance it started, and no other"),
    "SURVIVED": ("the Excel instance this run started is still running and could not be "
                 "ended: close it yourself"),
    "UNKNOWN": ("whether the Excel instance this run started has exited is not known: if "
                "an Excel window of this run is left open, close it"),
}


class RunError(Exception):
    """The run could not be recorded."""


class ExcelStopped(RunError):
    """The run stopped before PowerShell returned: its outcome and what it printed."""

    def __init__(self, message: str, outcome: str, output: str = "") -> None:
        super().__init__(message)
        self.outcome = outcome
        self.output = output


def built_references() -> list[str]:
    """The references of the project build_workbook.py writes, by their names in Excel."""
    return [*IMPLICIT_REFERENCES,
            *(libid.rsplit("#", 1)[1] for _, libid in vba_project.REFERENCES)]


def describe_windows(values: Mapping[str, object]) -> str:
    """The text winver shows, from the CurrentVersion registry values."""
    build = str(values.get("CurrentBuild", "")).strip()
    product = str(values.get("ProductName", "Windows")).strip()
    if build.isdigit() and int(build) >= WINDOWS_11_BUILD:
        product = product.replace("Windows 10", "Windows 11")
    version = str(values.get("DisplayVersion", "")).strip()
    patch = values.get("UBR")
    full = f"{build}.{patch}" if build and patch is not None else build
    return ", ".join(part for part in (product, f"version {version}" if version else "",
                                       f"OS build {full}" if full else "") if part)


def windows_version() -> str:
    """The Windows description; a value that is absent is left out, and a read
    that fails otherwise raises OSError, so that it is not mistaken for one."""
    import winreg  # Windows only

    values: dict[str, object] = {}
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, WINDOWS_KEY) as key:  # type: ignore[attr-defined]
        for name in ("ProductName", "DisplayVersion", "CurrentBuild", "UBR"):
            try:
                values[name] = winreg.QueryValueEx(key, name)[0]  # type: ignore[attr-defined]
            except FileNotFoundError:
                continue
    return describe_windows(values)


def host_record(fields: Mapping[str, str]) -> str:
    """The host record: one key=value line per field of check_excel_run.HOST_FIELDS,
    UNKNOWN for a value the run could not read."""
    return "".join(f"{key}={fields.get(key) or check_excel_run.UNKNOWN}\n"
                   for key in check_excel_run.HOST_FIELDS)


def now() -> str:
    """The local time with its UTC offset, to the second."""
    return datetime.now().astimezone().isoformat(timespec="seconds")


# The Machine field of a PE header, for the bitness of EXCEL.EXE.
PE_MACHINES = {0x8664: "64-bit", 0x014C: "32-bit"}


def executable_bitness(path: Path) -> str:
    """The bitness of a Windows executable, from its PE header; UNKNOWN when it
    cannot be read."""
    try:
        with path.open("rb") as file:
            header = file.read(64)
            file.seek(int.from_bytes(header[0x3C:0x40], "little"))
            signature = file.read(6)
    except (OSError, ValueError):
        return check_excel_run.UNKNOWN
    if header[:2] != b"MZ" or signature[:4] != b"PE\0\0":
        return check_excel_run.UNKNOWN
    return PE_MACHINES.get(int.from_bytes(signature[4:6], "little"), check_excel_run.UNKNOWN)


def registry_value(hive: str, key: str, name: str) -> object | None:
    """A registry value, read only; None when the key or the value is absent
    (FileNotFoundError). Any other failure, such as a denied read, raises
    OSError: an unread value is not an absent one."""
    import winreg  # Windows only

    try:
        with winreg.OpenKey(getattr(winreg, hive), key) as handle:  # type: ignore[attr-defined]
            value: object = winreg.QueryValueEx(handle, name)[0]  # type: ignore[attr-defined]
    except FileNotFoundError:
        return None
    return value


def vba_project_access(version: str,
                       read: Callable[[str, str, str], object | None] = registry_value) -> str:
    """Trust access to the VBA project object model, as the AccessVBOM value
    holds it: a policy value before the user's own setting, off when none is set.
    UNKNOWN when a location could not be read, rather than a setting of lower
    precedence."""
    security = rf"Microsoft\Office\{version}\Excel\Security"
    for hive, key in (("HKEY_CURRENT_USER", rf"Software\Policies\{security}"),
                      ("HKEY_LOCAL_MACHINE", rf"SOFTWARE\Policies\{security}"),
                      ("HKEY_CURRENT_USER", rf"Software\{security}")):
        try:
            value = read(hive, key, "AccessVBOM")
        except OSError:
            return check_excel_run.UNKNOWN
        if value is not None:
            return "TRUSTED" if value == 1 else "NOT_TRUSTED"
    return "NOT_TRUSTED"


def excel_host(output: str) -> dict[str, str]:
    """The Excel fields of the host record, from the EXCEL-<NAME>= lines the
    script prints once Excel has started."""
    found = {}
    for line in output.splitlines():
        name, separator, value = line.strip().partition("=")
        if separator and name.startswith("EXCEL-"):
            found[name.removeprefix("EXCEL-")] = value.strip()
    version, build, path = (found.get(name, "") for name in ("VERSION", "BUILD", "PATH"))
    security = found.get("AUTOMATION-SECURITY", "")
    names = check_excel_run.AUTOMATION_SECURITY
    fields = {"excel": f"{version} build {build}" if version and build else "",
              "bitness": executable_bitness(Path(path) / "EXCEL.EXE") if path else "",
              "automation_security": names[int(security) - 1]
              if security.isdigit() and 1 <= int(security) <= len(names) else ""}
    if version:
        try:
            fields["vba_project_access"] = vba_project_access(version)
        except ImportError:
            pass
    return fields


def excel_quit(output: str) -> str:
    """How the Excel instance the run started ended, from the EXCEL-QUIT= line the
    script prints: a key of QUIT_RESULTS, UNKNOWN when the line is absent (Excel did
    not start, or PowerShell stopped first) or holds another value."""
    result = check_excel_run.UNKNOWN
    for line in output.splitlines():
        name, separator, value = line.strip().partition("=")
        if separator and name == "EXCEL-QUIT":
            result = value.strip()
    return result if result in QUIT_RESULTS else check_excel_run.UNKNOWN


def quit_line(result: str) -> str:
    """The line that says how Excel ended, for the output and the report."""
    return f"excel-quit: {result} ({QUIT_RESULTS[result]})"


def with_quit(report: str, result: str) -> str:
    """The report, with how Excel ended after its Outcome line."""
    lines = report.splitlines()
    at = next((index + 1 for index, line in enumerate(lines) if line.startswith("Outcome: ")),
              len(lines))
    return "\n".join([*lines[:at], quit_line(result), *lines[at:]])


def windows_found() -> str:
    try:
        return windows_version()
    except (OSError, ImportError):
        return check_excel_run.UNKNOWN


def prepare(folder: Path, root: Path) -> None:
    """An empty evidence folder. Only a folder under test-results/ is emptied."""
    scratch = (root / "test-results").resolve()
    folder = folder.resolve()
    if folder.exists() and any(folder.iterdir()):
        if scratch not in folder.parents:
            raise RunError(f"{folder} is not empty, and only a folder under {scratch} is "
                           "emptied: choose an empty or new folder")
        shutil.rmtree(folder)
    folder.mkdir(parents=True, exist_ok=True)


def quoted(text: str) -> str:
    """A PowerShell single-quoted literal."""
    return "'" + text.replace("'", "''") + "'"


def powershell_script(workbook: Path, folder: Path, fixtures: Path,
                      entry_points: Sequence[str], probe: str, enable_timeout: int,
                      quit_wait: int = QUIT_WAIT) -> str:
    """Open the workbook read-only, wait until its macros run, run each certification
    entry point (``Module.Procedure``) in order, close without saving, quit. Every
    entry point runs even when an earlier one fails, so each suite leaves its logs; the
    first failure is reported. Each name is qualified with the workbook's name (``$prefix``),
    which lets Application.Run reach an Option Private Module module (HB-03). It never
    changes a macro setting: when the workbook opens with macros disabled (HB-02), a
    person enables them with Enable Content, without reopening it (HB-04). Once Excel
    has started it prints its version, build, folder and AutomationSecurity, which it
    only reads, for the host record. It records the process id of that instance, from
    its window, and the process start time. After Quit it waits up to ``quit_wait``
    seconds for that process to exit, then ends it, but only while its name is EXCEL
    and its start time is the one recorded, so that a process that reuses the id is
    never touched; it prints EXCEL-QUIT= with the result (QUIT_RESULTS). Its exit
    code says how the run ended (OUTCOME_BY_EXIT)."""
    calls = "\n".join(
        f"""        try {{
            [void]$excel.Run($prefix + {quoted(entry)}, $evidence, $fixtures)
        }} catch {{
            if (-not $failure) {{ $failure = {quoted(entry + ": ")} + $_.Exception.Message }}
        }}""" for entry in entry_points)
    return f"""$ErrorActionPreference = 'Stop'
# Progress records would reach the caller as CLIXML on standard error.
$ProgressPreference = 'SilentlyContinue'
$workbookPath = {quoted(str(workbook))}
$evidence = {quoted(str(folder))}
$fixtures = {quoted(str(fixtures))}
$excel = $null
$book = $null
$failure = $null
$stopped = $null
$excelId = $null
$excelStart = $null
try {{
    $excel = New-Object -ComObject Excel.Application
    # The process of this instance, from its window, and when it started: after
    # Quit, only a process with this id, name and start time is ever ended.
    try {{
        if (-not ('ExcelRunWindow' -as [type])) {{
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ExcelRunWindow {{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}}
'@
        }}
        [uint32]$windowProcess = 0
        $window = [IntPtr][int64]$excel.Hwnd
        [void][ExcelRunWindow]::GetWindowThreadProcessId($window, [ref]$windowProcess)
        if ($windowProcess -ne 0) {{
            $excelStart = (Get-Process -Id ([int]$windowProcess)).StartTime
            $excelId = [int]$windowProcess
        }}
    }} catch {{
        $excelId = $null
        $excelStart = $null
    }}
    # Read only, for the host record: the macro policy is recorded, never set.
    Write-Output ("EXCEL-VERSION=" + $excel.Version)
    Write-Output ("EXCEL-BUILD=" + $excel.Build)
    Write-Output ("EXCEL-PATH=" + $excel.Path)
    Write-Output ("EXCEL-AUTOMATION-SECURITY=" + [int]$excel.AutomationSecurity)
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    # UpdateLinks 0, ReadOnly: the workbook on disk stays the build.
    $book = $excel.Workbooks.Open($workbookPath, 0, $true)
    [void]$book.Activate()
    $prefix = "'" + $book.Name.Replace("'", "''") + "'!"
    # Wait until the workbook's macros run: at once where Excel enables them, or
    # after a person clicks Enable Content. The probe only reads harness state.
    $deadline = (Get-Date).AddSeconds({enable_timeout})
    $enabled = $false
    while (-not $enabled) {{
        try {{
            [void]$excel.Run($prefix + {quoted(probe)})
            $enabled = $true
        }} catch {{
            if ((Get-Date) -gt $deadline) {{ break }}
            Start-Sleep -Seconds 2
        }}
    }}
    if (-not $enabled) {{
        $failure = "NOT-ENABLED"
    }} else {{
{calls}
    }}
}} catch {{
    # Excel did not start, or the run stopped before the certification ended.
    $stopped = if ($null -eq $excel) {{ "UNAVAILABLE" }} else {{ "INCOMPLETE" }}
    $failure = $_.Exception.Message
}} finally {{
    $started = $null -ne $excel
    try {{
        if ($null -ne $book) {{
            $book.Close($false)
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
            $book = $null
        }}
        if ($null -ne $excel) {{
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
            $excel = $null
        }}
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
    }} finally {{
        # Runs even when closing failed. The instance this script started has
        # {quit_wait} s to exit; past that its process is ended, but only while the
        # process with its id still has its name and its start time.
        $quit = "UNKNOWN"
        if ($null -ne $excelId -and $null -ne $excelStart) {{
            try {{
                $process = Get-Process -Id $excelId -ErrorAction SilentlyContinue
                if ($null -eq $process -or $process.WaitForExit({quit_wait * 1000})) {{
                    $quit = "EXITED"
                }} else {{
                    $process.Refresh()
                    if ($process.ProcessName -ne 'EXCEL' -or $process.StartTime -ne $excelStart) {{
                        # Another process has the id now: this instance exited.
                        $quit = "EXITED"
                    }} else {{
                        try {{ Stop-Process -InputObject $process -Force }} catch {{ }}
                        $ended = $process.WaitForExit({STOP_WAIT * 1000})
                        $quit = if ($ended) {{ "STOPPED" }} else {{ "SURVIVED" }}
                    }}
                }}
            }} catch {{
                $quit = "UNKNOWN"
            }}
        }}
        if ($started) {{ Write-Output ("EXCEL-QUIT=" + $quit) }}
    }}
}}
if ($failure -eq "NOT-ENABLED") {{
    Write-Output "EXCEL-RUN=NOT-ENABLED"
    exit 4
}}
if ($stopped -eq "UNAVAILABLE") {{
    Write-Output "EXCEL-RUN=UNAVAILABLE; $failure"
    exit 5
}}
if ($stopped) {{
    Write-Output "EXCEL-RUN=INCOMPLETE; $failure"
    exit 6
}}
if ($failure) {{
    Write-Output "EXCEL-RUN=ERROR; $failure"
    exit 3
}}
Write-Output "EXCEL-RUN=DONE"
"""


def entry_points(policy: check_excel_run.Policy) -> list[str]:
    """The certification entry point of the example suite, then of each project suite,
    as ``Module.Procedure``."""
    return [f"{Path(policy.suite_source).stem}.{policy.entry_point}",
            *(f"{Path(item.source).stem}.{item.entry_point}" for item in policy.project_suites)]


def encoded(script: str) -> str:
    """The script as -EncodedCommand takes it: Base64 of UTF-16LE."""
    return base64.b64encode(script.encode("utf-16-le")).decode("ascii")


def run_excel(script: str, timeout: int) -> tuple[int, str]:
    command = ["powershell.exe", "-NoProfile", "-NonInteractive", "-EncodedCommand",
               encoded(script)]
    try:
        completed = subprocess.run(command, capture_output=True, text=True, timeout=timeout,
                                   check=False)
    except subprocess.TimeoutExpired as error:
        printed = error.stdout or ""
        raise ExcelStopped(f"Excel did not finish within {timeout} s; close it, check for a "
                           "dialog, then run TEST_Harness.ResetTests if it reports recovery",
                           "EXECUTION_TIMEOUT",
                           printed.decode(errors="replace") if isinstance(printed, bytes)
                           else printed) from error
    except OSError as error:
        raise ExcelStopped(f"PowerShell could not be started: {error}", "UNAVAILABLE") \
            from error
    return completed.returncode, (completed.stdout + completed.stderr).strip()


# The exit code of the PowerShell script when the macros were never enabled.
NOT_ENABLED = 4
# The outcome of the host record for each exit code of the PowerShell script:
# every entry point returned, one raised, the macros were never enabled, Excel
# did not start, or the run stopped before its end. Any other code is INCOMPLETE.
OUTCOME_BY_EXIT = {0: "PASS", 3: "FAIL", NOT_ENABLED: "NOT_ENABLED", 5: "UNAVAILABLE",
                   6: "INCOMPLETE"}
ENABLE_HINT = ("Excel opens the workbook now. If it shows a yellow bar SECURITY WARNING, "
               "Macros have been disabled, click Enable Content (or File > Info > Enable "
               "Content > Enable All Content). Change nothing else, and leave Excel alone "
               "until it closes.")


def record(root: Path, folder: Path, timeout: int, enable_timeout: int = 300,
           operator: str | None = None) -> int:
    policy = check_excel_run.working_policy(root)
    workbook = root / WORKBOOK
    built = build_workbook.build(root, check_excel_run.REBUILT_CONFIGURATION, workbook)
    print(f"built {workbook}: commit {built.commit}")
    if not built.clean:
        raise RunError("the checkout has local changes, so the build is not evidence-grade: "
                       "commit or remove them, or use a clean clone")
    prepare(folder, root)
    probe = f"{Path(policy.harness).stem}.LastResult"
    script = powershell_script(workbook, folder.resolve(), (root / FIXTURES).resolve(),
                               entry_points(policy), probe, enable_timeout)
    print(f"running {', '.join(entry_points(policy))} in Excel ...")
    print(ENABLE_HINT)
    host = {"operator": operator or getpass.getuser(), "started": now(),
            "references": "; ".join(built_references()),
            "compile": check_excel_run.COMPILE_ON_DEMAND}
    stopped: ExcelStopped | None = None
    try:
        code, output = run_excel(script, timeout)
    except ExcelStopped as error:
        code, output, stopped = -1, error.output, error
    host["finished"] = now()
    if output:
        print(output)
    host["outcome"] = stopped.outcome if stopped else OUTCOME_BY_EXIT.get(code, "INCOMPLETE")
    host.update(excel_host(output), windows=windows_found())
    # Written however the run ended, so that a run that stopped early is recorded.
    (folder / policy.host).write_text(host_record(host), encoding="utf-8")
    print(f"{policy.host}: outcome={host['outcome']}")
    quit_result = excel_quit(output)
    print(quit_line(quit_result))
    if stopped:
        raise stopped
    if code == NOT_ENABLED:
        raise RunError(f"the workbook's macros were not enabled within {enable_timeout} s, so "
                       "nothing ran: run the command again and click Enable Content in Excel")
    if code not in (0, 3):
        raise RunError(f"PowerShell or Excel failed (exit {code}, outcome "
                       f"{host['outcome']}); see the output above")
    evidence = check_excel_run.check(root, built.commit, folder, workbook, require_binding=True)
    report = check_excel_run.report(built.commit, evidence)
    print(report)
    saved = folder.parent / f"{folder.name}-report.txt"
    saved.write_text(with_quit(report, quit_result) + "\n", encoding="utf-8")
    print(f"report saved to {saved}")
    return 1 if evidence.findings else 0


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--evidence", type=Path, default=None,
                        help=f"the evidence folder (default: {DEFAULT_EVIDENCE.as_posix()})")
    parser.add_argument("--timeout", type=int, default=600,
                        help="seconds to wait for Excel in all (default: 600)")
    parser.add_argument("--enable-timeout", type=int, default=300,
                        help="seconds to wait for Enable Content (default: 300)")
    parser.add_argument("--operator", default=None,
                        help="who runs it, for the host record (default: the Windows user)")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    if args.enable_timeout >= args.timeout:
        parser.error("--enable-timeout must be shorter than --timeout")
    if os.name != "nt":
        print("excel-run: SKIP (needs Windows with Excel; this records no evidence)")
        return 2
    root = args.root.resolve()
    folder = args.evidence if args.evidence is not None else root / DEFAULT_EVIDENCE
    try:
        return record(root, folder, args.timeout, args.enable_timeout, args.operator)
    except (RunError, build_workbook.BuildError, check_excel_run.PolicyError, OSError) as error:
        print(f"FAIL: {error}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
