#!/usr/bin/env python3
"""Cross-check the test cases against the independent implementations that adapters run.

An expected value of class independent cites an IMPL- source of the source
register: an external implementation by another author
(docs/methodology/TEST_CASES.md, Independent implementations). An adapter lets
the check run that implementation. tests/independent/adapters.json registers
at most one adapter per IMPL- source, as the command that runs it.

The check runs each adapter once per case, with the argument a derivation
script receives: the fixture, or for a file-based case each variant's
directory. The adapter translates it into the implementation's input, runs the
implementation and prints one JSON object:

- {"outputs": [...]}: items with the quantity and scope of an expected output
  and its value or text;
- {"unsupported": "<reason>"}: the implementation cannot represent the case;
- {"unavailable": "<reason>"}: the implementation is not installed here.

Each printed output that the expected file holds is compared with it, as
tools/check_derivations.py compares: a number within
max(absolute, relative * |expected|), a text exactly. Every expected value
that cites the adapter's source must be printed. A case the implementation
cannot represent is listed with its reason and fails nothing.

An adapter whose program is not on PATH, or that answers unavailable, is
skipped with a note, so the check runs where the implementation is not
installed. With no adapter registered, the check passes and says so.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from check_derivations import key, label, matches  # noqa: E402
from check_test_cases import LOAD_FAILED, ROW, Report, declared, load_json  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
REGISTRY = Path("tests/independent/adapters.json")
FIXTURES = Path("tests/fixtures")
EXPECTED = Path("tests/expected")
REGISTRY_KEYS = frozenset({"schema_version", "kind", "adapters"})
ADAPTER_KEYS = frozenset({"source", "command"})
REPLIES = frozenset({"outputs", "unsupported", "unavailable"})
# A whole value with ASCII digits (fullmatch), as in check_test_cases.py.
SOURCE_ID = re.compile(r"IMPL-[0-9]+")
TIMEOUT_SECONDS = 300


@dataclass(frozen=True)
class Adapter:
    source: str
    command: list[str]


@dataclass
class Run:
    """One input of a case: a JSON fixture, or one variant of a file-based case."""

    where: str
    name: str
    argument: Path
    outputs: list[Any]  # as read: an entry may be malformed


@dataclass
class Outcome:
    findings: list[str] = field(default_factory=list)
    # Adapters not run here, and the runs an implementation cannot represent.
    skipped: list[str] = field(default_factory=list)
    unsupported: list[str] = field(default_factory=list)
    adapters: int = 0
    compared: int = 0


def read_registry(root: Path, report: Report) -> list[Adapter]:
    """The registered adapters; none when the registry is absent or unusable."""
    path = root / REGISTRY
    where = REGISTRY.as_posix()
    if not path.is_file():
        return []
    data = load_json(path, root, report)
    if data is LOAD_FAILED:
        return []
    if not isinstance(data, dict) or data.keys() != REGISTRY_KEYS or \
            data["schema_version"] != 1 or data["kind"] != "independent-adapters" or \
            not isinstance(data["adapters"], list):
        report.add(where, "envelope is not schema_version 1, kind 'independent-adapters' "
                          "and an adapters list")
        return []
    sources = declared(root, "SOURCES.md", "Register", ROW, report)
    adapters: list[Adapter] = []
    seen: set[str] = set()
    for index, item in enumerate(data["adapters"]):
        at = f"{where}: adapters[{index}]"
        if not isinstance(item, dict) or item.keys() != ADAPTER_KEYS:
            report.add(at, f"keys are not {', '.join(sorted(ADAPTER_KEYS))}")
            continue
        source, command = item["source"], item["command"]
        problems = []
        if not isinstance(source, str) or not SOURCE_ID.fullmatch(source):
            problems.append("source is not an IMPL- ID")
        elif source not in sources:
            problems.append(f"source {source} is not in the register (SOURCES.md)")
        elif source in seen:
            problems.append(f"source {source} has a second adapter")
        if not isinstance(command, list) or not command or \
                not all(isinstance(word, str) and word.strip() for word in command):
            problems.append("command is not a non-empty list of words")
        for problem in problems:
            report.add(at, problem)
        if not problems:
            seen.add(source)
            adapters.append(Adapter(source, list(command)))
    return adapters


def case_runs(root: Path) -> list[Run]:
    """Every input an adapter is given, with the expected outputs it is compared with."""
    runs = []
    for path in sorted((root / EXPECTED).glob("*.json")):
        where = path.relative_to(root).as_posix()
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (UnicodeDecodeError, ValueError):
            continue  # check_test_cases.py reports a malformed file
        if not isinstance(data, dict):
            continue
        variants = data.get("variants")
        if not isinstance(variants, dict):
            outputs = data.get("outputs")
            runs.append(Run(where, path.stem, root / FIXTURES / path.name,
                            outputs if isinstance(outputs, list) else []))
            continue
        for variant, outputs in sorted(variants.items()):
            directory = root / FIXTURES / path.stem / variant
            if directory.is_dir():  # check_test_cases.py reports a variant without files
                runs.append(Run(f"{where}: variants.{variant}", f"{path.stem}/{variant}",
                                directory, outputs if isinstance(outputs, list) else []))
    return runs


def executable(command: Sequence[str]) -> tuple[list[str] | None, str]:
    """The command to run, with python as this interpreter; None and why when absent."""
    first = command[0]
    if first == "python":
        return [sys.executable, *command[1:]], ""
    found = shutil.which(first)
    if found is None:
        return None, f"{first} is not on PATH"
    return [found, *command[1:]], ""


def reply(command: list[str], shown: str, root: Path,
          argument: Path) -> tuple[dict[str, Any] | None, str]:
    """The adapter's reply for one input, or None and why it is unusable."""
    try:
        completed = subprocess.run([*command, str(argument)], cwd=root, capture_output=True,
                                   text=True, check=False, stdin=subprocess.DEVNULL,
                                   timeout=TIMEOUT_SECONDS)
    except subprocess.TimeoutExpired:
        return None, f"{shown} did not finish within {TIMEOUT_SECONDS} seconds"
    except OSError as error:
        return None, f"{shown} could not start ({error})"
    if completed.returncode != 0:
        detail = (completed.stderr.strip().splitlines() or ["no message"])[-1]
        return None, f"{shown} exited with {completed.returncode}: {detail}"
    try:
        data = json.loads(completed.stdout)
    except ValueError as error:
        return None, f"{shown} did not print JSON ({error})"
    if not isinstance(data, dict) or len(data) != 1 or not data.keys() <= REPLIES:
        return None, f"{shown} must print an object with one of {', '.join(sorted(REPLIES))}"
    value = next(iter(data.values()))
    if "outputs" in data:
        if not isinstance(value, list) or not all(isinstance(item, dict) for item in value):
            return None, f"{shown}: outputs is not a list of objects"
    elif not isinstance(value, str) or not value.strip():
        return None, f"{shown}: {next(iter(data))} is not a reason"
    return data, ""


def cites(output: object, source: str) -> bool:
    reference = output.get("reference") if isinstance(output, dict) else None
    return isinstance(reference, dict) and reference.get("source") == source


def compare(adapter: Adapter, run: Run, printed: list[dict[str, Any]],
            outcome: Outcome) -> None:
    by_key = {key(item): item for item in printed}
    for output in run.outputs:
        if not isinstance(output, dict):
            continue
        actual = by_key.get(key(output))
        if actual is None:
            if cites(output, adapter.source):
                outcome.findings.append(f"{run.where}: {adapter.source} gives no "
                                        f"{label(key(output))}, which cites it")
            continue
        tolerance = output.get("tolerance")
        if "text" not in output and not ("value" in output and isinstance(tolerance, dict)
                                         and {"absolute", "relative"} <= tolerance.keys()):
            continue  # check_test_cases.py reports an output without a value or tolerance
        outcome.compared += 1
        if reason := matches(output, actual):
            outcome.findings.append(f"{run.where}: {label(key(output))} from "
                                    f"{adapter.source}: {reason}")


def cross_check(root: Path, adapter: Adapter, runs: Sequence[Run], outcome: Outcome) -> None:
    command, missing = executable(adapter.command)
    if command is None:
        outcome.skipped.append(f"{adapter.source} skipped: {missing}")
        return
    for run in runs:
        data, problem = reply(command, " ".join(adapter.command), root, run.argument)
        if data is None:
            outcome.findings.append(f"{run.where}: {adapter.source}: {problem}")
        elif "unavailable" in data:
            outcome.skipped.append(f"{adapter.source} skipped: {data['unavailable'].strip()}")
            return
        elif "unsupported" in data:
            unchecked = sum(cites(output, adapter.source) for output in run.outputs)
            note = f" ({unchecked} value(s) citing it not recomputed)" if unchecked else ""
            outcome.unsupported.append(f"{adapter.source} cannot represent {run.name}: "
                                       f"{data['unsupported'].strip()}{note}")
        else:
            compare(adapter, run, data["outputs"], outcome)


def check(root: Path) -> Outcome:
    report = Report()
    adapters = read_registry(root, report)
    outcome = Outcome(findings=report.findings, adapters=len(adapters))
    if adapters:
        runs = case_runs(root)
        for adapter in adapters:
            cross_check(root, adapter, runs, outcome)
    return outcome


def summary(outcome: Outcome) -> str:
    if not outcome.adapters:
        return (f"no adapter registered in {REGISTRY.as_posix()}; skipped, nothing "
                "cross-checked")
    return (f"{outcome.adapters} adapter(s), {len(outcome.skipped)} skipped; "
            f"{outcome.compared} value(s) compared; {len(outcome.unsupported)} input(s) "
            "not representable")


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    outcome = check(args.root.resolve())
    for line in outcome.skipped + outcome.unsupported:
        print(f"NOTE: {line}")
    for finding in outcome.findings:
        print(f"FAIL: {finding}")
    status = "FAIL" if outcome.findings else "PASS"
    print(f"independent-implementations: {status} ({summary(outcome)})")
    return 1 if outcome.findings else 0


if __name__ == "__main__":
    sys.exit(main())
