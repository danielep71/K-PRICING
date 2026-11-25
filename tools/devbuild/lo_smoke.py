#!/usr/bin/env python3
"""Run the harness certification in headless LibreOffice as an early warning.

Usage: python3 tools/devbuild/lo_smoke.py [WORKBOOK]
Without WORKBOOK it builds the development workbook first.

It opens the workbook in LibreOffice with VBA compatibility and runs the
example suite's certification entry point into a temporary folder, with the
checkout's tests/fixtures as the folder of the test cases. The entry point,
the run logs of the modes, the certification log and the results file are the
names that .github/excel-evidence-policy.json gives. Then it prints a JSON
summary: the modules LibreOffice loaded, the error the certification raised,
each run log's RESULT and failures, the certification log, the results file,
and the failures that LibreOffice is not known to cause. It exits 0 when there is none, 1 when
there is one, the certification wrote no log, the build of the development
workbook fails (lo-smoke: FAIL (build: ...)), or LibreOffice could not open
the workbook or run it, and 2 when an option is unknown, WORKBOOK does not
exist, the evidence policy is missing or malformed, or LibreOffice or its
Python UNO bridge (python3-uno) is missing. Run it with the Python that
provides the bridge.

LibreOffice is not a supported host, and this is never Excel evidence. It
only catches some runtime defects before a run on the supported host. It runs
valid VBA differently in known ways:
- vbObjectError in a constant expression stops the whole project from
  compiling, so the harness writes its error numbers as literals;
- Debug.Print raises when it runs headless, so the harness prints through
  TEST_Harness.Echo;
- Application.Build does not exist, so ENVIRONMENT has an empty build;
- after a call that raised under On Error Resume Next, Err.Number reads 0,
  so the assertions on raised errors fail (KNOWN_FAILURES below).
"""

from __future__ import annotations

import argparse
import importlib
import importlib.util
import json
import shutil
import socket
import subprocess
import sys
import tempfile
import time
from collections.abc import Mapping
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(ROOT / "tools"))

import check_excel_run  # noqa: E402

# Assertions of the example suite that fail in LibreOffice only, because
# Err.Number reads 0 after a call that raised under On Error Resume Next.
KNOWN_FAILURES = frozenset({
    "raised-error/negative-input",
    "cleanup-keeps-primary-error/primary-error",
    "cleanup-partial-capture/capture-error",
    "cleanup-change-failure/change-error",
    "cleanup-primary-only/primary-error",
    "cleanup-each-restoration-failure/calculation-cleanup-error",
    "cleanup-each-restoration-failure/screen-cleanup-error",
    "cleanup-only-failure/cleanup-error",
})
# The failure each deliberate failure mode is designed to produce, by mode.
DESIGNED = {"INJECTED_FAILURE": frozenset({"exact-double/half-plus-quarter"}),
            "SKIPPED_CASE": frozenset({"harness.completeness"})}


def entry_point(policy: check_excel_run.Policy) -> str:
    """The example suite's certification entry point, as Module.Procedure."""
    return f"{Path(policy.suite_source).stem}.{policy.entry_point}"


def log_names(policy: check_excel_run.Policy) -> tuple[str, ...]:
    """The run log of each mode, in run order, then the certification log."""
    return (*policy.logs.values(), policy.certification)


def designed(policy: check_excel_run.Policy) -> dict[str, frozenset[str]]:
    """The failures each deliberate failure mode is designed to produce, by its log."""
    return {policy.logs[mode]: failures for mode, failures in DESIGNED.items()}


def free_port() -> int:
    with socket.socket() as probe:
        probe.bind(("127.0.0.1", 0))
        return int(probe.getsockname()[1])


def property_value(name: str, value: object) -> Any:
    item = importlib.import_module("uno").createUnoStruct("com.sun.star.beans.PropertyValue")
    item.Name = name
    item.Value = value
    return item


def connect(port: int) -> Any:
    local = importlib.import_module("uno").getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(120):
        try:
            return resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;"
                                    "StarOffice.ComponentContext")
        except Exception:  # noqa: BLE001 - the office is still starting
            time.sleep(0.5)
    raise RuntimeError("LibreOffice did not accept a connection")


def enable_vba(context: Any) -> None:
    provider = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.configuration.ConfigurationProvider", context)
    access = provider.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationUpdateAccess",
        (property_value("nodepath", "/org.openoffice.Office.Calc/Filter/Import/VBA"),))
    access.setPropertyValue("Load", True)
    access.setPropertyValue("Executable", True)
    access.commitChanges()


def certify(context: Any, workbook: Path, folder: Path, procedure: str) -> dict[str, object]:
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",
                                                              context)
    try:
        document = desktop.loadComponentFromURL(
            workbook.resolve().as_uri(), "_blank", 0,
            (property_value("Hidden", True), property_value("MacroExecutionMode", 4)))
    except Exception:
        # Let the office exit on its own; killing the launcher leaves it running.
        desktop.terminate()
        raise
    summary: dict[str, object] = {"workbook": str(workbook)}
    try:
        summary["modules"] = sorted(document.BasicLibraries.getByName("VBAProject")
                                    .getElementNames())
        script = document.getScriptProvider().getScript(
            f"vnd.sun.star.script:VBAProject.{procedure}?language=Basic&location=document")
        try:
            script.invoke((str(folder), str(ROOT / "tests" / "fixtures")), (), ())
            summary["error"] = None
        except Exception as error:  # noqa: BLE001 - reported, not raised
            summary["error"] = str(error)
    finally:
        document.close(True)
        try:
            desktop.terminate()
        except Exception:  # noqa: BLE001 - the office may already be gone
            pass
    return summary


def read_logs(folder: Path, policy: check_excel_run.Policy) -> dict[str, object]:
    found: dict[str, object] = {}
    for name in log_names(policy):
        matches = sorted(folder.rglob(name)) + sorted(folder.parent.glob(f"*{name}"))
        if not matches:
            found[name] = None
            continue
        lines = matches[0].read_text(encoding="utf-8", errors="replace").splitlines()
        if name == policy.certification:
            found[name] = lines
        else:
            found[name] = {"result": lines[-1] if lines else "",
                           "failures": [line for line in lines
                                        if line.startswith(("FAILURE=", "ERROR="))]}
    return found


def read_results(folder: Path, name: str) -> object:
    """The results file, parsed; its text when it is not JSON; None when it is missing.
    The workbook's build stamp names the working tree's HEAD, so it is shown, not checked."""
    matches = sorted(folder.rglob(name)) + sorted(folder.parent.glob(f"*{name}"))
    if not matches:
        return None
    text = matches[0].read_text(encoding="utf-8", errors="replace")
    try:
        return json.loads(text)
    except ValueError:
        return text


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("workbook", nargs="?", type=Path,
                        help="the workbook to run; without it, the development workbook is built")
    args = parser.parse_args(argv)
    # Checked before LibreOffice starts, which would otherwise wait on a load
    # that cannot succeed.
    if args.workbook is not None and not args.workbook.is_file():
        print(f"lo-smoke: ERROR (no workbook at {args.workbook})", file=sys.stderr)
        return 2
    try:
        policy = check_excel_run.working_policy(ROOT)
    except check_excel_run.PolicyError as error:
        print(f"lo-smoke: ERROR ({error})", file=sys.stderr)
        return 2
    soffice = shutil.which("soffice")
    if soffice is None or importlib.util.find_spec("uno") is None:
        print("lo-smoke: SKIP (needs LibreOffice and python3-uno; this is never Excel evidence)")
        return 2
    with tempfile.TemporaryDirectory(prefix="lo-smoke-") as temporary:
        base = Path(temporary)
        if args.workbook is not None:
            workbook = args.workbook
        else:
            import build_workbook
            workbook = base / "Workbook-development.xlsm"
            try:
                build_workbook.build(ROOT, "development", workbook)
            except (build_workbook.BuildError, ValueError, OSError) as error:
                # A form, a module too dense to compress or an unreadable source.
                print(f"lo-smoke: FAIL (build: {error})")
                return 1
        folder = base / "evidence"
        folder.mkdir()
        port = free_port()
        office = subprocess.Popen(
            [soffice, "--headless", "--invisible", "--norestore", "--nologo",
             f"-env:UserInstallation={(base / 'profile').as_uri()}",
             f"--accept=socket,host=127.0.0.1,port={port};urp;"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        failure = None
        try:
            context = connect(port)
            enable_vba(context)
            summary = certify(context, workbook, folder, entry_point(policy))
        except Exception as error:  # noqa: BLE001 - a UNO error: reported, not raised
            failure = f"{type(error).__name__}: {error}"
        finally:
            try:
                office.wait(timeout=60)
            except subprocess.TimeoutExpired:
                office.kill()
        if failure is not None:
            print(f"lo-smoke: FAIL (LibreOffice could not run the workbook: {failure})")
            return 1
        summary["logs"] = read_logs(folder, policy)
        summary["results"] = read_results(folder, policy.results)
    summary["unexpected"] = unexpected(summary["logs"], designed(policy))  # type: ignore[arg-type]
    if not isinstance(summary["results"], dict) or summary["results"].get("kind") != "test-results":
        summary["unexpected"].append(f"{policy.results}: not written as a test-results file")  # type: ignore[attr-defined]
    summary["note"] = "LibreOffice is not a supported host; this is never Excel evidence"
    print(json.dumps(summary, indent=1, default=str))
    return 1 if summary["unexpected"] else 0


def unexpected(logs: dict[str, Any], expected: Mapping[str, frozenset[str]]) -> list[str]:
    """Failures and errors that LibreOffice is not known to cause, and missing logs;
    ``expected`` holds, by log, the failures a deliberate failure mode produces."""
    found = []
    for name, content in logs.items():
        if content is None:
            found.append(f"{name}: not written")
        elif isinstance(content, dict):
            for line in content["failures"]:
                key, _, rest = line.partition("=")
                identifier = rest.split(";", 1)[0]
                if key == "ERROR" or identifier not in KNOWN_FAILURES | expected.get(name, set()):
                    found.append(f"{name}: {line}")
    return found


if __name__ == "__main__":
    sys.exit(main())
