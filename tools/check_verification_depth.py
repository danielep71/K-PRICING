#!/usr/bin/env python3
"""Fail when a release-blocking check has no executable negative proof.

A check that accepts everything stays green, and so does a check whose
negative tests have been deleted. tools/verification_depth.json records, for
each check that can fail Repository integrity, the tests that prove it can
fail. Each entry of its "checks" list has these fields, none empty:

- check: the check's name, as tools/check_repository.py lists it;
- checker: what the check runs, as the integrity-checks table of
  CONTRIBUTING.md shows it: the command, with OWNER/REPOSITORY for the
  repository, or tools/check_repository.py:<function> for a built-in check;
- proofs: unittest ids, tests.<module>.<Class>.<test_method>, of tests that
  show the check rejecting a bad input, or "probe" for the check's Probe,
  which the run makes the tool reject before the check may pass;
- negative_claim: the violations the proofs show rejected, at least
  MIN_CLAIM characters long;
- state: one of STATES.

The checks that need an entry are not a second list kept by hand: they are
those of check_repository.checks() as CI runs them, with --ci, less the
advisory ones, which cannot fail the run. The check fails when:

- a release-blocking check has no entry, or an entry names a check that does
  not exist or is advisory;
- an entry is listed twice, has a field missing, empty or unknown, or its
  checker differs from what the check runs;
- a state is not in the closed list, or a negative claim is too short to be
  specific;
- a proof is neither "probe" nor a unittest id, names the probe of a check
  that has none, or names a test module, class or method that
  python -m unittest discover -s tests, which the tests check runs, does not
  collect.

It does not run the proofs: the tests check runs every collected test. To
write a new proof, make the defective input with tools/make_mutant.py
(CONTRIBUTING.md, Negative controls).
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import tempfile
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

import check_data_contract  # noqa: E402
import check_repository  # noqa: E402
import render_docs  # noqa: E402
from check_repository import Check, Context  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
INVENTORY = Path("tools/verification_depth.json")
TESTS = "tests"
SCHEMA_VERSION = 1
ENVELOPE = frozenset({"schema_version", "checks"})
FIELDS = ("check", "checker", "proofs", "negative_claim", "state")
STATES = {
    "green": "the check passes, and its proofs run wherever the tests run",
    "green_in_ci": "the check passes; a proof is skipped where a tool it needs is missing, "
                   "and runs in CI, which installs the tool",
}
# Shorter than this, a claim cannot say which violations are rejected.
MIN_CLAIM = 40
PROBE = "probe"
TEST_ID = re.compile(r"tests\.(test\w*)\.([A-Za-z_]\w*)\.(test\w*)")
TIMEOUT_SECONDS = 300
# Collects the tests as python -m unittest discover -s tests does, without
# running them, and writes their ids and the modules that failed to import.
DISCOVER = """\
import json, sys, unittest
loader = unittest.TestLoader()
ids, failed = [], {}
def walk(suite):
    for test in suite:
        if isinstance(test, unittest.TestSuite):
            walk(test)
        elif type(test).__name__ == "_FailedTest":
            failed[test._testMethodName] = str(getattr(test, "_exception", "")).strip()
        else:
            ids.append(test.id())
walk(loader.discover(sys.argv[1]))
with open(sys.argv[2], "w", encoding="utf-8") as handle:
    json.dump({"ids": ids, "failed": failed}, handle)
"""


@dataclass
class Collected:
    """The test ids that discovery collects, without the tests. prefix, and the failures."""

    ids: set[str] = field(default_factory=set)
    # Module name: why it cannot be imported.
    failed: dict[str, str] = field(default_factory=dict)


def listed_in_ci(root: Path) -> list[Check]:
    """The checks of Repository integrity as CI runs them, advisory ones included."""
    context = Context(root, repository="OWNER/REPOSITORY", ci=True)
    return check_repository.checks(context)


def checker_of(check: Check) -> str:
    """What a check runs, as the integrity-checks table shows it."""
    if check.command is not None:
        return render_docs.shown_command(check.command)
    assert check.function is not None
    return f"tools/check_repository.py:{check.function.__name__}"


def collect(root: Path) -> Collected:
    """The tests that the tests check runs, collected in a separate process."""
    with tempfile.TemporaryDirectory(prefix="verification-depth-") as directory:
        out = Path(directory) / "collected.json"
        try:
            completed = subprocess.run([sys.executable, "-c", DISCOVER, TESTS, str(out)],
                                       cwd=root, capture_output=True, text=True,
                                       timeout=TIMEOUT_SECONDS, check=False)
        except (OSError, subprocess.TimeoutExpired) as error:
            raise RuntimeError(f"the tests cannot be collected: {error}") from None
        if completed.returncode or not out.is_file():
            lines = (completed.stderr or completed.stdout).strip().splitlines()
            raise RuntimeError("the tests cannot be collected: "
                               + (lines[-1] if lines else f"exit code {completed.returncode}"))
        data = json.loads(out.read_text(encoding="utf-8"))
    return Collected(set(data["ids"]), dict(data["failed"]))


def load(path: Path) -> tuple[Any, list[str]]:
    """The inventory as read, and why it cannot be read."""
    try:
        return check_data_contract.strict_json(path.read_text(encoding="utf-8")), []
    except FileNotFoundError:
        return None, [f"{INVENTORY.as_posix()} is missing"]
    except (OSError, UnicodeDecodeError, ValueError) as error:
        return None, [f"{INVENTORY.as_posix()} is not valid JSON: {error}"]


def proof_problem(name: str, proof: str, check: Check, collected: Collected) -> str | None:
    """Why a proof does not prove the check can fail, or None if it does."""
    if proof == PROBE:
        return None if check.probe is not None else \
            f"{name}: proof {PROBE!r}, but the check has no probe"
    match = TEST_ID.fullmatch(proof)
    if match is None:
        return (f"{name}: proof {proof!r} is neither {PROBE!r} nor a unittest id "
                f"{TESTS}.<module>.<Class>.<test_method>")
    module, case, _ = match.groups()
    if proof.removeprefix(f"{TESTS}.") in collected.ids:
        return None
    where = f"{name}: proof {proof}"
    if module in collected.failed:
        return f"{where}: {TESTS}/{module}.py cannot be imported: {collected.failed[module]}"
    if not any(test.startswith(f"{module}.") for test in collected.ids):
        return (f"{where}: module {TESTS}/{module}.py is not collected by python -m unittest "
                f"discover -s {TESTS}")
    if not any(test.startswith(f"{module}.{case}.") for test in collected.ids):
        return f"{where}: class {case} is not collected from {TESTS}/{module}.py"
    return f"{where}: method is not collected from class {case}"


def entry_problems(name: str, entry: Mapping[str, Any], check: Check,
                   collected: Collected) -> list[str]:
    """The problems of one entry of a release-blocking check."""
    problems = []
    checker, state, claim = entry.get("checker"), entry.get("state"), entry.get("negative_claim")
    if isinstance(checker, str) and checker.strip() and checker != checker_of(check):
        problems.append(f"{name}: checker {checker!r} is not what the check runs, "
                        f"{checker_of(check)!r}")
    if isinstance(state, str) and state.strip() and state not in STATES:
        problems.append(f"{name}: state {state!r} is not one of {', '.join(STATES)}")
    if isinstance(claim, str) and claim.strip() and len(claim.strip()) < MIN_CLAIM:
        problems.append(f"{name}: negative_claim is too short to be specific "
                        f"({len(claim.strip())} characters, at least {MIN_CLAIM})")
    proofs = entry.get("proofs")
    if not isinstance(proofs, list) or not proofs:
        return problems
    seen: set[str] = set()
    for proof in proofs:
        if not isinstance(proof, str) or not proof.strip():
            problems.append(f"{name}: a proof is not a non-empty text")
        elif proof in seen:
            problems.append(f"{name}: proof {proof} is listed twice")
        else:
            seen.add(proof)
            problem = proof_problem(name, proof, check, collected)
            if problem:
                problems.append(problem)
    return problems


def field_problems(name: str, entry: Mapping[str, Any]) -> list[str]:
    """Fields that are missing, empty or unknown."""
    problems = []
    for key in FIELDS[1:]:
        value = entry.get(key)
        if not value or (isinstance(value, str) and not value.strip()):
            problems.append(f"{name}: {key} is missing or empty")
        elif not isinstance(value, list if key == "proofs" else str):
            problems.append(f"{name}: {key} is not a {'list' if key == 'proofs' else 'text'}")
    unknown = sorted(set(entry) - set(FIELDS))
    if unknown:
        problems.append(f"{name}: unknown field(s) {', '.join(unknown)}")
    return problems


def validate(data: Any, listed: Sequence[Check], collected: Collected) -> list[str]:
    """Every problem of the inventory, against the checks and the collected tests."""
    if not isinstance(data, dict) or set(data) != ENVELOPE \
            or type(data["schema_version"]) is not int \
            or data["schema_version"] != SCHEMA_VERSION or not isinstance(data["checks"], list):
        return [f"the inventory must be an object with exactly schema_version "
                f"{SCHEMA_VERSION} and a checks list"]
    blocking = {check.name: check for check in listed if not check.advisory}
    advisory = {check.name for check in listed if check.advisory}
    problems: list[str] = []
    seen: dict[str, int] = {}
    for index, entry in enumerate(data["checks"]):
        if not isinstance(entry, dict):
            problems.append(f"checks[{index}] is not an object")
            continue
        name = entry.get("check")
        if not isinstance(name, str) or not name.strip():
            problems.append(f"checks[{index}]: check is missing or empty")
            continue
        if name in seen:
            problems.append(f"{name}: listed twice, as checks[{seen[name]}] and checks[{index}]")
            continue
        seen[name] = index
        problems += field_problems(name, entry)
        if name in advisory:
            problems.append(f"{name}: the check is advisory and cannot fail the run; remove "
                            "its entry")
        elif name not in blocking:
            problems.append(f"{name}: no such check in tools/check_repository.py; remove its "
                            "entry")
        else:
            problems += entry_problems(name, entry, blocking[name], collected)
    problems += [f"{name}: release-blocking check has no entry in {INVENTORY.as_posix()}"
                 for name in blocking if name not in seen]
    return problems


def main(argv: Sequence[str] | None = None,
         listed: Callable[[Path], list[Check]] = listed_in_ci) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    root = args.root.resolve()
    data, problems = load(root / INVENTORY)
    checks = listed(root)
    if not problems:
        try:
            problems = validate(data, checks, collect(root))
        except RuntimeError as error:
            problems = [str(error)]
    for problem in problems:
        print(f"FAIL: {problem}")
    blocking = sum(not check.advisory for check in checks)
    print(f"verification-depth: {'FAIL' if problems else 'PASS'} ({blocking} release-blocking "
          "check(s); the tests check runs the proofs)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
