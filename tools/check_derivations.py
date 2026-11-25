#!/usr/bin/env python3
"""Run the derivation scripts of the test cases and compare their results.

An expected value may name, in reference.script, a script under
tests/derivations/ that computes it from the documented method
(docs/methodology/TEST_CASES.md, Derivation scripts). This check runs each
named script once per case, with the case's fixture as its only argument, and
compares every value the expected file attributes to it:

- a number matches when |actual - expected| <= max(absolute, relative * |expected|);
- a text matches exactly.

A file-based case runs each script once per variant instead, with the
variant's directory as the argument, and compares that variant's values.

A match shows that the script and the expected file agree. It validates
neither: a value derived within the project stays illustrative. With no
script named, the check passes and says so.
"""

from __future__ import annotations

import argparse
import json
import math
import subprocess
import sys
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = Path("tests/fixtures")
EXPECTED = Path("tests/expected")
DERIVATIONS = "tests/derivations/"
TIMEOUT_SECONDS = 60

Key = tuple[str, str | None, str | None]


def key(item: Mapping[str, Any]) -> Key:
    """An output's quantity and scope, comparable across the two files."""
    scope = item.get("scope")
    if isinstance(scope, dict):
        return str(item.get("quantity")), str(scope.get("dataset")), str(scope.get("record"))
    return str(item.get("quantity")), None, None


def label(output_key: Key) -> str:
    quantity, dataset, record = output_key
    return quantity if dataset is None else f"{quantity} ({dataset}/{record})"


def matches(expected: Mapping[str, Any], actual: Mapping[str, Any]) -> str | None:
    """None when the derived output matches the expected one, else the reason."""
    if "text" in expected:
        if actual.get("text") != expected["text"]:
            return f"text {actual.get('text')!r}, expected {expected['text']!r}"
        return None
    value = actual.get("value")
    if not isinstance(value, (int, float)) or isinstance(value, bool) or not math.isfinite(value):
        return f"value {value!r} is not a finite number"
    want = expected["value"]
    tolerance = expected["tolerance"]
    bound = max(tolerance["absolute"], tolerance["relative"] * abs(want))
    if abs(value - want) > bound:
        return f"value {value!r}, expected {want!r} within {bound!r}"
    return None


def run_script(root: Path, script: str, fixture: Path) -> tuple[dict[Key, Any] | None, str]:
    """Run one script on one fixture; return its outputs by key, or None and the reason."""
    if not script.startswith(DERIVATIONS) or ".." in script or not (root / script).is_file():
        return None, f"script {script!r} is not an existing file under {DERIVATIONS}"
    try:
        completed = subprocess.run([sys.executable, "-I", str(root / script), str(fixture)],
                                   cwd=root, capture_output=True, text=True, check=False,
                                   timeout=TIMEOUT_SECONDS)
    except subprocess.TimeoutExpired:
        return None, f"{script} did not finish within {TIMEOUT_SECONDS} seconds"
    if completed.returncode != 0:
        detail = (completed.stderr.strip().splitlines() or ["no message"])[-1]
        return None, f"{script} exited with {completed.returncode}: {detail}"
    try:
        data = json.loads(completed.stdout)
    except ValueError as error:
        return None, f"{script} did not print JSON ({error})"
    outputs = data.get("outputs") if isinstance(data, dict) else None
    if not isinstance(outputs, list) or not all(isinstance(item, dict) for item in outputs):
        return None, f"{script} must print an object with an outputs list"
    return {key(item): item for item in outputs}, ""


def finite(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def comparable(output: Mapping[str, Any]) -> bool:
    """A text output, or a number with an absolute and a relative tolerance."""
    if "text" in output:
        return True
    tolerance = output.get("tolerance")
    return finite(output.get("value")) and isinstance(tolerance, dict) \
        and finite(tolerance.get("absolute")) and finite(tolerance.get("relative"))


def compare(root: Path, where: str, outputs: object, fixture: Path) -> tuple[list[str], int]:
    """Run the scripts that one list of outputs names on one fixture, and compare."""
    by_script: dict[str, list[dict[str, Any]]] = {}
    for output in outputs if isinstance(outputs, list) else []:
        reference = output.get("reference") if isinstance(output, dict) else None
        script = reference.get("script") if isinstance(reference, dict) else None
        if not isinstance(script, str):
            continue
        if not comparable(output):
            continue  # check_test_cases.py reports an output without a value or tolerance
        by_script.setdefault(script, []).append(output)
    findings: list[str] = []
    compared = 0
    for script, wanted in sorted(by_script.items()):
        derived, problem = run_script(root, script, fixture)
        if derived is None:
            findings.append(f"{where}: {problem}")
            continue
        for output in wanted:
            compared += 1
            actual = derived.get(key(output))
            if actual is None:
                findings.append(f"{where}: {script} gives no {label(key(output))}")
            elif reason := matches(output, actual):
                findings.append(f"{where}: {label(key(output))} from {script}: {reason}")
    return findings, compared


def check_case(root: Path, path: Path) -> tuple[list[str], int]:
    """Check one expected file; return the findings and the number of values compared."""
    where = path.relative_to(root).as_posix()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, ValueError):
        return [], 0  # check_test_cases.py reports a malformed file
    if not isinstance(data, dict):
        return [], 0
    variants = data.get("variants")
    if not isinstance(variants, dict):
        return compare(root, where, data.get("outputs"), root / FIXTURES / path.name)
    findings: list[str] = []
    compared = 0
    for variant, outputs in sorted(variants.items()):
        directory = root / FIXTURES / path.stem / variant
        if not directory.is_dir():
            continue  # check_test_cases.py reports an expected variant without files
        variant_findings, variant_compared = compare(
            root, f"{where}: variants.{variant}", outputs, directory)
        findings += variant_findings
        compared += variant_compared
    return findings, compared


def check(root: Path) -> tuple[list[str], int]:
    findings: list[str] = []
    compared = 0
    for path in sorted((root / EXPECTED).glob("*.json")):
        case_findings, case_compared = check_case(root, path)
        findings += case_findings
        compared += case_compared
    return findings, compared


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, compared = check(args.root.resolve())
    for finding in findings:
        print(f"FAIL: {finding}")
    status = "FAIL" if findings else "PASS"
    found = f"{compared} derived value(s) compared" if compared else "no derivation script named"
    print(f"derivations: {status} ({found}; agreement only, not validation)")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
