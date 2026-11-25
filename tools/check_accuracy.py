#!/usr/bin/env python3
"""Check the accuracy contracts of the test cases, and measure them against an Excel run.

An accuracy contract (docs/methodology/TEST_CASES.md, Accuracy contracts) is the
worst error a calculation guarantees for one regime, over the outputs of its
quantities in two lists of cases: the calibration cases, which may be used to
set the threshold, and the holdout cases, which never are. The registry is
tests/accuracy/contracts.json. The threshold is the guarantee, which blocks;
an optional target, no larger, is the accuracy aimed for, which is only
reported. The registry may also declare invariants: properties that the
results of a calculation's cases always have, whatever any contract's status,
and whether or not a contract covers them:
- bounds: each number of a quantity is at least a minimum, at most a
  maximum, or both, with no tolerance (a domain rule);
- monotone: over an ordered list of cases or variants, a quantity does not
  decrease (non-decreasing) or does not increase (non-increasing) by more
  than the two outputs' own tolerances; a smaller reversal is only counted;
- sum: within each case, variant and scope, declared quantities add up to
  another quantity or to a constant, within the sum of their tolerances.

Without a results file, the check reads the registry against the cases:
- every contract is well formed, and its target, when it has one, is a finite
  number of at least 0 and no larger than its threshold;
- its calculation is declared by the calculation contracts;
- each of its cases tests that calculation and has an expected number for each
  of its quantities, and no case is both calibration and holdout;
- a frozen contract names the decision that froze it, is not illustrative, and
  has at least min_holdout_cases holdout cases, every expected value of which
  is published or independent;
- its design document, when it names one, is a file under docs/methodology/
  or tests/accuracy/;
- a contract frozen at the newest vX.Y.Z tag that HEAD contains is still
  frozen, and keeps its guarantee unless it names another decision: the same
  calculation, regime and metric, a threshold no larger, and every quantity
  and case it had. Its target may change under the same decision;
- every invariant is well formed, its calculation is declared, each of its
  cases tests that calculation and has the expected numbers it needs, and the
  expected values themselves keep it.

With --results, a results file of an Excel run, it also fails; repeated, it
measures the results files of several suites together:
- results of another revision: their commit is the full SHA of the candidate,
  --commit when given, HEAD otherwise;
- a result that two results files both hold;
- an output with no result;
- a result outside its expected output's own tolerance;
- a contract whose worst error, among its calibration or among its holdout
  outputs, exceeds its threshold;
- an output an invariant reads with no result, and a result that breaks an
  invariant.
It prints, for each split, the worst error, the threshold and the target, and
marks the target missed when the worst error is above the target but within
the threshold: that passes. It prints, for each invariant, what it checked,
and for a monotone one the reversals within the allowance.

The holdout exposure record, tests/accuracy/holdout_exposures.json, lists
each time a frozen contract's holdout was first measured under a decision.
A repository with a registry keeps it, with an empty list until a holdout is
first measured. A contract's holdout is measured only once its exposure under
its decision is recorded:
- a draft contract's holdout is sealed: --results measures its calibration
  cases only and prints that the holdout is sealed;
- a frozen contract with no exposure under its decision is sealed too, and
  with --results that is a finding: measure it once with --record-holdout;
- --record-holdout ACC-<n> writes the contract's exposure, then measures as
  --results does. The entry holds the contract ID, its decision, the number
  of exposures of its holdout so far, the commit and the SHA-256 of each
  results file, the SHA-256 of the canonical contract entry, its threshold,
  the SHA-256 of each holdout case's fixture and expected file, and the byte
  length and SHA-256 of the contract's design document when it names one. It
  refuses, writing nothing and measuring nothing, a draft contract, a
  contract already exposed under its decision, a missing, malformed or
  inconsistent record, a registry check with any finding, and results that
  are unusable or lack an output of the holdout. A record is never reset;
- any later --results run, such as tools/check_excel_run.py on a later
  candidate, measures a recorded holdout as a regression, prints its number
  of exposures and records nothing.
The check also fails, without results, on a record that is malformed or
inconsistent, an exposure whose contract is gone, an exposure recorded at
the newest vX.Y.Z tag that the record no longer starts with, and, for a
frozen contract under the decision of its last exposure, an entry, holdout
fixture or expected file that differs from the record, a lower threshold,
or a design document that no longer begins with the recorded bytes. A new
decision, which states the earlier exposures, ends these comparisons until
its own exposure is recorded.

The error is |a - e| for an absolute contract and |a - e| / |e| for a relative
one; a relative error at e = 0 is 0 when a = 0 and infinite otherwise. A
contract proves only what its cases and their references prove: an
illustrative contract validates nothing.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import sys
from collections.abc import Callable, Collection, Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

from check_data_contract import SNAKE  # noqa: E402
from check_test_cases import (  # noqa: E402
    LOAD_FAILED,
    Report,
    is_number,
    load_json,
    read_methodology,
)
from gitio import text_or_none  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
REGISTRY = Path("tests/accuracy/contracts.json")
FIXTURES = Path("tests/fixtures")
EXPECTED = Path("tests/expected")
MANIFEST = "case.json"
REGISTRY_KEYS = frozenset({"schema_version", "kind", "min_holdout_cases", "contracts"})
# Keys the registry may leave out: a registry without them keeps schema_version 1.
OPTIONAL_REGISTRY_KEYS = frozenset({"invariants"})
CONTRACT_KEYS = frozenset({"id", "calculation", "regime", "quantities", "metric", "threshold",
                           "provenance", "status", "decision", "calibration_cases",
                           "holdout_cases"})
# Keys a contract may leave out: a registry without them keeps schema_version 1.
OPTIONAL_CONTRACT_KEYS = frozenset({"target", "design"})
# Where a contract's design document may be: what tools/check_excel_run.py copies.
DESIGN_HOMES = ("docs/methodology/", "tests/accuracy/")
# The holdout exposure record, beside the registry.
EXPOSURES = REGISTRY.parent / "holdout_exposures.json"
EXPOSURES_KEYS = frozenset({"schema_version", "kind", "exposures"})
EXPOSURE_KEYS = frozenset({"contract", "decision", "exposures", "commit", "results",
                           "entry_sha256", "threshold", "holdout", "design"})
# An invariant's keys, by kind.
INVARIANT_COMMON = frozenset({"id", "kind", "calculation", "basis"})
INVARIANT_KEYS = {
    "bounds": INVARIANT_COMMON | {"quantity", "cases", "minimum", "maximum"},
    "monotone": INVARIANT_COMMON | {"quantity", "scope", "direction", "sequence"},
    "sum": INVARIANT_COMMON | {"parts", "total", "cases"},
}
DIRECTIONS = frozenset({"non-decreasing", "non-increasing"})
POINT_KEYS = frozenset({"case", "variant"})
SCOPE_KEYS = frozenset({"dataset", "record"})
RESULTS_KEYS = frozenset({"schema_version", "kind", "commit", "results"})
RESULT_KEYS = frozenset({"case", "variant", "quantity", "scope", "value"})
# Each pattern must match a whole value (fullmatch): "$" would also accept a
# trailing line break. Digits are ASCII: \d would also accept other scripts' digits.
CONTRACT_ID = re.compile(r"ACC-[0-9]+")
INVARIANT_ID = re.compile(r"INV-[0-9]+")
INDEPENDENT = frozenset({"published", "independent"})
METRICS = frozenset({"absolute", "relative"})
STATUSES = frozenset({"draft", "frozen"})
TAG = re.compile(r"v[0-9]+\.[0-9]+\.[0-9]+")
SHA = re.compile(r"[0-9a-f]{40}")
DIGEST = re.compile(r"[0-9a-f]{64}")

# A result's identity: case, variant (None for a JSON fixture), quantity and scope.
Key = tuple[str, str | None, str, str | None]


@dataclass(frozen=True)
class Output:
    """One expected number of a case."""

    key: Key
    value: float
    absolute: float
    relative: float
    reference: str

    def bound(self) -> float:
        """How far a result may lie from the expected value: its own tolerance."""
        return max(self.absolute, self.relative * abs(self.value))


@dataclass
class Contract:
    id: str
    calculation: str
    regime: str
    quantities: list[str]
    metric: str
    threshold: float
    status: str
    decision: str | None
    calibration: list[str]
    holdout: list[str]
    # Reported only: the worst error aimed for, no larger than the threshold.
    target: float | None = None
    # The committed design or preregistration document, a repository path.
    design: str | None = None
    # The SHA-256 of the canonical entry, which an exposure record binds.
    entry_sha256: str = ""


@dataclass
class Invariant:
    """A property the results of a calculation's cases always have; only the
    fields of its kind are set."""

    id: str
    kind: str
    calculation: str
    # bounds and sum: the cases, in order; monotone: the cases of its points.
    cases: list[str]
    # bounds and monotone: the one quantity; sum: the parts, then a total quantity.
    quantities: list[str]
    minimum: float | None = None
    maximum: float | None = None
    direction: str | None = None
    scope: str | None = None
    # monotone: the ordered points, each a case and its variant.
    sequence: list[tuple[str, str | None]] = field(default_factory=list)
    parts: list[str] = field(default_factory=list)
    # sum: a quantity name, or a constant.
    total: str | float | None = None


@dataclass
class Registry:
    contracts: list[Contract] = field(default_factory=list)
    min_holdout: int = 1
    invariants: list[Invariant] = field(default_factory=list)


def scope_text(scope: object) -> str | None:
    if isinstance(scope, dict):
        return f"{scope.get('dataset')}/{scope.get('record')}"
    return None


def place(key: Key) -> str:
    """Where a result is: its case, variant and scope."""
    case, variant, _, scope = key
    where = case if variant is None else f"{case}[{variant}]"
    return where + (f" ({scope})" if scope else "")


def label(key: Key) -> str:
    case, variant, quantity, scope = key
    where = case if variant is None else f"{case}[{variant}]"
    return f"{where} {quantity}" + (f" ({scope})" if scope else "")


def case_outputs(root: Path, case: str, report: Report) -> tuple[str | None, list[Output]]:
    """A case's calculation and its expected numbers; (None, []) when it has no fixture."""
    fixture = root / FIXTURES / f"{case}.json"
    manifest = root / FIXTURES / case / MANIFEST
    path = fixture if fixture.is_file() else manifest
    data = load_json(path, root, report) if path.is_file() else LOAD_FAILED
    expected = load_json(root / EXPECTED / f"{case}.json", root, report) \
        if (root / EXPECTED / f"{case}.json").is_file() else LOAD_FAILED
    if not isinstance(data, dict) or not isinstance(expected, dict):
        return None, []
    calculation = data.get("calculation")
    variants = expected.get("variants")
    # A malformed file gives no outputs here: check_test_cases.py reports it.
    groups: dict[str | None, Any] = {None: expected.get("outputs")} if "outputs" in expected \
        else dict(variants) if isinstance(variants, dict) else {}
    outputs = []
    for variant, items in groups.items():
        for item in items if isinstance(items, list) else []:
            tolerance = item.get("tolerance") if isinstance(item, dict) else None
            if not isinstance(item, dict) or not is_number(item.get("value")) or \
                    not isinstance(tolerance, dict) or not is_number(tolerance.get("absolute")) or \
                    not is_number(tolerance.get("relative")):
                continue
            reference = item.get("reference")
            key = (case, variant, str(item.get("quantity")), scope_text(item.get("scope")))
            outputs.append(Output(key, float(item["value"]), float(tolerance["absolute"]),
                                  float(tolerance["relative"]),
                                  str(reference.get("class") if isinstance(reference, dict) else None)))
    return calculation if isinstance(calculation, str) else None, outputs


def one_of(value: object, choices: Collection[str]) -> bool:
    """Whether a JSON value is one of the texts in ``choices``; an array or an object,
    which cannot be hashed, is not one of them, and is a finding rather than an error."""
    return isinstance(value, str) and value in choices


def target_problems(data: Mapping[str, Any]) -> list[str]:
    """What is wrong with a contract's optional target, which is only reported: a
    target looser than the threshold, which blocks, would aim for nothing."""
    if "target" not in data:
        return []
    target = data["target"]
    if not is_number(target) or target < 0:
        return ["target is not a finite number of at least 0"]
    threshold = data.get("threshold")
    if is_number(threshold) and target > threshold:
        return [f"target {target!r} is larger than the threshold {threshold!r}"]
    return []


def is_design_path(value: object) -> bool:
    """A repository path, with forward slashes and no "..", to a file below one of
    DESIGN_HOMES."""
    if not isinstance(value, str) or "\\" in value or value != value.strip():
        return False
    parts = value.split("/")
    return any(value.startswith(home) for home in DESIGN_HOMES) and \
        all(part not in ("", ".", "..") for part in parts)


def contract_problems(data: Mapping[str, Any]) -> list[str]:
    """What is wrong with one contract's own fields."""
    problems = []
    if not isinstance(data.get("id"), str) or not CONTRACT_ID.fullmatch(data["id"]):
        problems.append("id is not ACC-<number>")
    if not isinstance(data.get("regime"), str) or not data["regime"].strip():
        problems.append("regime is not a text")
    quantities = data.get("quantities")
    if not isinstance(quantities, list) or not quantities or \
            not all(isinstance(item, str) and SNAKE.fullmatch(item) for item in quantities):
        problems.append("quantities is not a non-empty list of lower snake case names")
    elif len(set(quantities)) != len(quantities):
        problems.append("quantities lists a name twice")
    if not one_of(data.get("metric"), METRICS):
        problems.append("metric is not absolute or relative")
    if not is_number(data.get("threshold")) or data["threshold"] < 0:
        problems.append("threshold is not a finite number of at least 0")
    problems += target_problems(data)
    if "design" in data and not is_design_path(data["design"]):
        problems.append("design is not a path to a file under "
                        f"{' or '.join(DESIGN_HOMES)}")
    if not isinstance(data.get("provenance"), str) or not data["provenance"].strip():
        problems.append("provenance is not a text")
    if not one_of(data.get("status"), STATUSES):
        problems.append("status is not draft or frozen")
    decision = data.get("decision")
    if data.get("status") == "frozen" and (not isinstance(decision, str) or not decision.strip()):
        problems.append("a frozen contract names the decision that froze it")
    if data.get("status") == "draft" and decision is not None:
        problems.append("a draft contract has no decision; decision is null")
    for name in ("calibration_cases", "holdout_cases"):
        cases = data.get(name)
        if not isinstance(cases, list) or not all(isinstance(item, str) for item in cases):
            problems.append(f"{name} is not a list of case IDs")
        elif len(set(cases)) != len(cases):
            # Each case counts once, so a repeated ID would inflate the holdout.
            problems.append(f"{name} lists a case twice")
    return problems


def is_name_list(value: object) -> bool:
    return isinstance(value, list) and bool(value) and \
        all(isinstance(item, str) and SNAKE.fullmatch(item) for item in value)


def case_list_problems(data: Mapping[str, Any]) -> list[str]:
    cases = data.get("cases")
    if not isinstance(cases, list) or not cases or \
            not all(isinstance(item, str) for item in cases):
        return ["cases is not a non-empty list of case IDs"]
    if len(set(cases)) != len(cases):
        return ["cases lists a case twice"]
    return []


def bound_problems(data: Mapping[str, Any]) -> list[str]:
    """A bounds invariant's limits: numbers or null, at least one a number."""
    problems = [f"{name} is not a finite number or null" for name in ("minimum", "maximum")
                if data.get(name) is not None and not is_number(data[name])]
    if problems:
        return problems
    minimum, maximum = data.get("minimum"), data.get("maximum")
    if minimum is None and maximum is None:
        return ["minimum and maximum are both null; a bounds invariant states at least one"]
    if minimum is not None and maximum is not None and minimum > maximum:
        return [f"minimum {minimum!r} is larger than the maximum {maximum!r}"]
    return []


def scope_problems(scope: object) -> list[str]:
    if scope is None:
        return []
    if not isinstance(scope, dict) or set(scope) != SCOPE_KEYS or \
            not all(isinstance(value, str) and value.strip() for value in scope.values()):
        return ["scope is not null or an object with exactly a dataset and a record text"]
    return []


def sequence_problems(sequence: object) -> list[str]:
    """A monotone invariant's ordered points: at least two, each a case and its
    variant (null for a JSON fixture), none twice."""
    if not isinstance(sequence, list) or len(sequence) < 2 or not all(
            isinstance(item, dict) and set(item) == POINT_KEYS and
            isinstance(item["case"], str) and isinstance(item["variant"], (str, type(None)))
            for item in sequence):
        return ["sequence is not a list of at least two points, each an object with a case "
                "text and a variant text or null"]
    points = [(item["case"], item["variant"]) for item in sequence]
    if len(set(points)) != len(points):
        return ["sequence lists a point twice"]
    return []


def invariant_problems(data: Mapping[str, Any]) -> list[str]:
    """What is wrong with one invariant's own fields, its keys being those of its kind."""
    problems = []
    if not isinstance(data.get("id"), str) or not INVARIANT_ID.fullmatch(data["id"]):
        problems.append("id is not INV-<number>")
    if not isinstance(data.get("calculation"), str):
        problems.append("calculation is not a text")
    if not isinstance(data.get("basis"), str) or not data["basis"].strip():
        problems.append("basis is not a text")
    kind = data.get("kind")
    if kind in ("bounds", "monotone") and \
            (not isinstance(data.get("quantity"), str) or not SNAKE.fullmatch(data["quantity"])):
        problems.append("quantity is not a lower snake case name")
    if kind in ("bounds", "sum"):
        problems += case_list_problems(data)
    if kind == "bounds":
        problems += bound_problems(data)
    elif kind == "monotone":
        problems += scope_problems(data.get("scope"))
        if not one_of(data.get("direction"), DIRECTIONS):
            problems.append("direction is not non-decreasing or non-increasing")
        problems += sequence_problems(data.get("sequence"))
    elif kind == "sum":
        parts, total = data.get("parts"), data.get("total")
        if not isinstance(parts, list) or not is_name_list(parts):
            problems.append("parts is not a non-empty list of lower snake case names")
        elif len(set(parts)) != len(parts):
            problems.append("parts lists a name twice")
        if isinstance(total, str) and SNAKE.fullmatch(total):
            if isinstance(parts, list) and total in parts:
                problems.append(f"total {total} is also a part")
        elif not is_number(total):
            problems.append("total is not a lower snake case name or a finite number")
    return problems


def read_invariant(item: Mapping[str, Any]) -> Invariant:
    """An invariant from fields that invariant_problems() found well formed."""
    kind = item["kind"]
    if kind == "monotone":
        sequence = [(point["case"], point["variant"]) for point in item["sequence"]]
        return Invariant(item["id"], kind, item["calculation"],
                         list(dict.fromkeys(case for case, _ in sequence)), [item["quantity"]],
                         direction=item["direction"], scope=scope_text(item["scope"]),
                         sequence=sequence)
    if kind == "bounds":
        return Invariant(item["id"], kind, item["calculation"], list(item["cases"]),
                         [item["quantity"]],
                         minimum=None if item["minimum"] is None else float(item["minimum"]),
                         maximum=None if item["maximum"] is None else float(item["maximum"]))
    total = item["total"]
    return Invariant(item["id"], kind, item["calculation"], list(item["cases"]),
                     list(item["parts"]) + ([total] if isinstance(total, str) else []),
                     parts=list(item["parts"]),
                     total=total if isinstance(total, str) else float(total))


def read_invariants(items: list[Any], report: Report, where: str) -> list[Invariant]:
    """The well-formed invariants; each malformed one is a finding."""
    invariants = []
    seen: set[str] = set()
    for index, item in enumerate(items):
        name = item.get("id") if isinstance(item, dict) else None
        at = f"{where}: invariants[{index}]" + (f" {name}" if isinstance(name, str) else "")
        kind = item.get("kind") if isinstance(item, dict) else None
        if not isinstance(kind, str) or kind not in INVARIANT_KEYS:
            report.add(at, "is not an object whose kind is bounds, monotone or sum")
            continue
        assert isinstance(item, dict)
        if set(item) != INVARIANT_KEYS[kind]:
            report.add(at, f"keys of a {kind} invariant are not "
                           f"{', '.join(sorted(INVARIANT_KEYS[kind]))}")
            continue
        problems = invariant_problems(item)
        if isinstance(item["id"], str):
            if item["id"] in seen:
                problems.append("id appears twice")
            seen.add(item["id"])
        for problem in problems:
            report.add(at, problem)
        if not problems:
            invariants.append(read_invariant(item))
    return invariants


def read_registry(root: Path, text: str | None, report: Report, where: str) -> Registry | None:
    """The registry in ``text`` (the working tree's when None); None when absent or unusable."""
    if text is None:
        path = root / REGISTRY
        if not path.is_file():
            return None
        data = load_json(path, root, report)
    else:
        try:
            data = json.loads(text)
        except ValueError:
            return None
    if data is LOAD_FAILED:
        return None
    if not isinstance(data, dict) or \
            not REGISTRY_KEYS <= data.keys() <= REGISTRY_KEYS | OPTIONAL_REGISTRY_KEYS or \
            data["schema_version"] != 1 or data["kind"] != "accuracy-contracts":
        report.add(where, "envelope is not schema_version 1, kind 'accuracy-contracts', "
                          "min_holdout_cases and contracts, and optionally invariants")
        return None
    minimum = data["min_holdout_cases"]
    if not isinstance(minimum, int) or isinstance(minimum, bool) or minimum < 1:
        report.add(where, "min_holdout_cases is not a whole number of at least 1")
        return None
    if not isinstance(data["contracts"], list):
        report.add(where, "contracts is not a list")
        return None
    if not isinstance(data.get("invariants", []), list):
        report.add(where, "invariants is not a list")
        return None
    registry = Registry(min_holdout=minimum)
    seen: set[str] = set()
    for index, item in enumerate(data["contracts"]):
        name = item.get("id") if isinstance(item, dict) else None
        at = f"{where}: contracts[{index}]" + (f" {name}" if isinstance(name, str) else "")
        if not isinstance(item, dict) or \
                not CONTRACT_KEYS <= set(item) <= CONTRACT_KEYS | OPTIONAL_CONTRACT_KEYS:
            report.add(at, f"keys are not {', '.join(sorted(CONTRACT_KEYS))}, and "
                           f"optionally {', '.join(sorted(OPTIONAL_CONTRACT_KEYS))}")
            continue
        problems = contract_problems(item)
        # A malformed id is already a problem; only a text id can repeat.
        if isinstance(item["id"], str):
            if item["id"] in seen:
                problems.append("id appears twice")
            seen.add(item["id"])
        for problem in problems:
            report.add(at, problem)
        if problems:
            continue
        registry.contracts.append(Contract(
            item["id"], str(item["calculation"]), item["regime"], list(item["quantities"]),
            item["metric"],
            float(item["threshold"]), item["status"], item["decision"],
            list(item["calibration_cases"]), list(item["holdout_cases"]),
            float(item["target"]) if "target" in item else None,
            item.get("design"), canonical_sha256(item)))
    registry.invariants = read_invariants(data.get("invariants", []), report, where)
    return registry


def contract_outputs(root: Path, contract: Contract, calculations: Mapping[str, bool],
                     report: Report, where: str) -> dict[str, list[Output]]:
    """Each split's expected outputs of the contract, after checking its cases."""
    if contract.calculation not in calculations:
        report.add(where, f"calculation {contract.calculation} is not declared by the "
                          "calculation contracts")
    if calculations.get(contract.calculation) and contract.status == "frozen":
        report.add(where, f"calculation {contract.calculation} is illustrative, so the "
                          "contract stays draft")
    both = set(contract.calibration) & set(contract.holdout)
    for case in sorted(both):
        report.add(where, f"case {case} is both calibration and holdout")
    if not contract.calibration and not contract.holdout:
        report.add(where, "lists no case")
    splits: dict[str, list[Output]] = {}
    for split, cases in (("calibration", contract.calibration), ("holdout", contract.holdout)):
        splits[split] = []
        for case in cases:
            calculation, outputs = case_outputs(root, case, report)
            if calculation is None:
                report.add(where, f"case {case} has no fixture and expected file")
                continue
            if calculation != contract.calculation:
                report.add(where, f"case {case} tests {calculation}, not "
                                  f"{contract.calculation}")
            mine = [output for output in outputs if output.key[2] in contract.quantities]
            for quantity in contract.quantities:
                if not any(output.key[2] == quantity for output in mine):
                    report.add(where, f"case {case} has no expected number for {quantity}")
            splits[split] += mine
    return splits


def invariant_cases(root: Path, invariant: Invariant, calculations: Mapping[str, bool],
                    report: Report, where: str) -> dict[str, list[Output]]:
    """Each case's expected outputs of the invariant's quantities, after checking its
    calculation and that the case exists and tests it."""
    if invariant.calculation not in calculations:
        report.add(where, f"calculation {invariant.calculation} is not declared by the "
                          "calculation contracts")
    found: dict[str, list[Output]] = {}
    for case in invariant.cases:
        calculation, outputs = case_outputs(root, case, report)
        if calculation is None:
            report.add(where, f"case {case} has no fixture and expected file")
            continue
        if calculation != invariant.calculation:
            report.add(where, f"case {case} tests {calculation}, not {invariant.calculation}")
        found[case] = [output for output in outputs if output.key[2] in invariant.quantities]
    return found


def monotone_terms(invariant: Invariant, found: Mapping[str, list[Output]], report: Report,
                   where: str) -> list[list[Output]]:
    """Each two consecutive points of the sequence, both with an expected number."""
    points: list[Output | None] = []
    for case, variant in invariant.sequence:
        key = (case, variant, invariant.quantities[0], invariant.scope)
        point = next((output for output in found.get(case, []) if output.key == key), None)
        if point is None and case in found:
            report.add(where, f"point {label(key)} has no expected number")
        points.append(point)
    return [[before, after] for before, after in zip(points, points[1:])
            if before is not None and after is not None]


def sum_terms(invariant: Invariant, found: Mapping[str, list[Output]], report: Report,
              where: str) -> list[list[Output]]:
    """The parts, then a total quantity, at each variant and scope of each case where
    one of them has an expected number; every one of them must have one there."""
    terms = []
    for case, outputs in found.items():
        groups: dict[tuple[str | None, str | None], dict[str, Output]] = {}
        for output in outputs:
            groups.setdefault((output.key[1], output.key[3]), {})[output.key[2]] = output
        if not groups:
            report.add(where, f"case {case} has no expected number for "
                              f"{', '.join(invariant.quantities)}")
        for (variant, scope), group in groups.items():
            missing = [name for name in invariant.quantities if name not in group]
            for name in missing:
                report.add(where, f"{place((case, variant, name, scope))} has no expected "
                                  f"number for {name}")
            if not missing:
                terms.append([group[name] for name in invariant.quantities])
    return terms


def invariant_terms(root: Path, invariant: Invariant, calculations: Mapping[str, bool],
                    report: Report, where: str) -> list[list[Output]]:
    """The groups of expected outputs that the invariant relates, after checking its
    cases: each output of a bounds invariant alone; each two consecutive points of a
    monotone one; the parts, then a total quantity, of each case, variant and scope
    of a sum."""
    found = invariant_cases(root, invariant, calculations, report, where)
    if invariant.kind == "monotone":
        return monotone_terms(invariant, found, report, where)
    if invariant.kind == "sum":
        return sum_terms(invariant, found, report, where)
    for case, outputs in found.items():
        if not outputs:
            report.add(where, f"case {case} has no expected number for "
                              f"{invariant.quantities[0]}")
    return [[output] for outputs in found.values() for output in outputs]


def describe(invariant: Invariant) -> str:
    """What the invariant states, for its measurement line."""
    quantity = invariant.quantities[0]
    if invariant.kind == "bounds":
        low = "" if invariant.minimum is None else f"{invariant.minimum!r} <= "
        high = "" if invariant.maximum is None else f" <= {invariant.maximum!r}"
        return f"{low}{quantity}{high}"
    if invariant.kind == "monotone":
        scope = f" ({invariant.scope})" if invariant.scope else ""
        return f"{quantity}{scope} {invariant.direction} over " \
               f"{len(invariant.sequence)} point(s)"
    total = invariant.total if isinstance(invariant.total, str) else repr(invariant.total)
    return f"{' + '.join(invariant.parts)} = {total}"


@dataclass
class Verdict:
    """What an invariant found on one set of values."""

    missing: list[Output] = field(default_factory=list)
    violations: list[str] = field(default_factory=list)
    checked: int = 0
    # monotone: reversals within the allowance, which are only counted.
    reversals: int = 0


def evaluate(invariant: Invariant, terms: list[list[Output]],
             value: Callable[[Output], float | None]) -> Verdict:
    """The invariant on the values that ``value`` gives each output: the expected
    values, or the results. A term with an output that has no value is not checked."""
    verdict = Verdict()
    for term in terms:
        values = [value(output) for output in term]
        absent = [output for output, number in zip(term, values) if number is None]
        if absent:
            verdict.missing += [output for output in absent if output not in verdict.missing]
            continue
        numbers = [number for number in values if number is not None]
        verdict.checked += 1
        if invariant.kind == "bounds":
            (output,), (number,) = term, numbers
            if invariant.minimum is not None and number < invariant.minimum:
                verdict.violations.append(f"{label(output.key)} is {number!r}, below the "
                                          f"minimum {invariant.minimum!r}")
            if invariant.maximum is not None and number > invariant.maximum:
                verdict.violations.append(f"{label(output.key)} is {number!r}, above the "
                                          f"maximum {invariant.maximum!r}")
        elif invariant.kind == "monotone":
            (before, after), (first, second) = term, numbers
            allowance = before.bound() + after.bound()
            back = first - second if invariant.direction == "non-decreasing" else second - first
            if back > allowance:
                change = "decreases" if invariant.direction == "non-decreasing" \
                    else "increases"
                verdict.violations.append(
                    f"{invariant.quantities[0]} {change} from {first!r} at "
                    f"{place(before.key)} to {second!r} at {place(after.key)}, by {back!r}, "
                    f"beyond the allowance {allowance!r}")
            elif back > 0:
                verdict.reversals += 1
        else:
            count = len(invariant.parts)
            parts = math.fsum(numbers[:count])
            total = numbers[count] if isinstance(invariant.total, str) \
                else float(invariant.total or 0.0)
            allowance = math.fsum(output.bound() for output in term)
            difference = abs(parts - total)
            if difference > allowance:
                verdict.violations.append(
                    f"at {place(term[0].key)}, {' + '.join(invariant.parts)} is {parts!r} and "
                    f"the total is {total!r}: the difference {difference!r} is beyond the "
                    f"allowance {allowance!r}")
    return verdict


def invariant_line(invariant: Invariant, verdict: Verdict) -> str:
    """One invariant's measurement: what it states, what it checked, and what broke it."""
    unit = {"bounds": "result(s)", "monotone": "step(s)", "sum": "group(s)"}[invariant.kind]
    line = f"{invariant.id} {invariant.kind}: {describe(invariant)}, {verdict.checked} " \
           f"{unit} checked"
    if invariant.kind == "monotone":
        line += f", {verdict.reversals} reversal(s) within the allowance"
    if verdict.violations:
        line += f", {len(verdict.violations)} " + {
            "bounds": "outside the bounds", "monotone": "reversal(s) beyond the allowance",
            "sum": "beyond the allowance"}[invariant.kind]
    return line


def frozen_problems(contract: Contract, holdout: list[Output], minimum: int) -> list[str]:
    if contract.status != "frozen":
        return []
    problems = []
    if len(contract.holdout) < minimum:
        problems.append(f"a frozen contract needs at least {minimum} holdout case(s), has "
                        f"{len(contract.holdout)}")
    weak = sorted({label(output.key) for output in holdout
                   if output.reference not in INDEPENDENT})
    problems += [f"holdout output {item} is not published or independent" for item in weak]
    return problems


def newest_tag(root: Path) -> str | None:
    """The newest vX.Y.Z tag that HEAD contains, by version."""
    tags = [tag for tag in (text_or_none(root, "tag", "--merged", "HEAD") or "").split()
            if TAG.fullmatch(tag)]
    return max(tags, key=lambda tag: tuple(int(part) for part in tag[1:].split(".")),
               default=None)


def canonical_sha256(value: object) -> str:
    """The SHA-256 of a JSON value written canonically: keys sorted, no white space,
    UTF-8. The record binds a contract entry by it."""
    text = json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def bytes_sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def fixture_sha256(root: Path, case: str) -> str | None:
    """The SHA-256 of a case's fixture: of the bytes of a JSON fixture or, for a
    file-based case, of a listing of every file in its folder, sorted by path, one
    line each: the file's SHA-256, two spaces, its path in the folder and a line
    feed. None when the case has no fixture."""
    fixture = root / FIXTURES / f"{case}.json"
    if fixture.is_file():
        return bytes_sha256(fixture.read_bytes())
    folder = root / FIXTURES / case
    if not (folder / MANIFEST).is_file():
        return None
    files = sorted((path.relative_to(folder).as_posix(), path) for path in folder.rglob("*")
                   if path.is_file())
    listing = "".join(f"{bytes_sha256(path.read_bytes())}  {name}\n" for name, path in files)
    return bytes_sha256(listing.encode("utf-8"))


def expected_sha256(root: Path, case: str) -> str | None:
    path = root / EXPECTED / f"{case}.json"
    return bytes_sha256(path.read_bytes()) if path.is_file() else None


@dataclass(frozen=True)
class Exposure:
    """One recorded first measurement of a frozen contract's holdout under a decision."""

    contract: str
    decision: str
    # The number of exposures of the contract's holdout, this one included.
    count: int
    commit: str
    entry_sha256: str
    threshold: float
    # By holdout case: the SHA-256 of its fixture and of its expected file.
    holdout: dict[str, tuple[str, str]]
    # The design document: its path, the number of bytes recorded and their SHA-256.
    design: tuple[str, int, str] | None


@dataclass
class Record:
    exposures: list[Exposure] = field(default_factory=list)
    # The entries as the file holds them: the record only grows.
    raw: list[Any] = field(default_factory=list)

    def of(self, contract: str) -> list[Exposure]:
        return [item for item in self.exposures if item.contract == contract]


def is_digest(value: object) -> bool:
    return isinstance(value, str) and DIGEST.fullmatch(value) is not None


def is_count(value: object, least: int) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value >= least


def exposure_problems(item: Mapping[str, Any]) -> list[str]:
    """What is wrong with one exposure entry's own fields."""
    problems = []
    if not isinstance(item.get("contract"), str) or not CONTRACT_ID.fullmatch(item["contract"]):
        problems.append("contract is not ACC-<number>")
    if not isinstance(item.get("decision"), str) or not item["decision"].strip():
        problems.append("decision is not a text")
    if not is_count(item.get("exposures"), 1):
        problems.append("exposures is not a whole number of at least 1")
    if not isinstance(item.get("commit"), str) or not SHA.fullmatch(item["commit"]):
        problems.append("commit is not a full lower-case commit SHA")
    results = item.get("results")
    if not isinstance(results, list) or not results or not all(
            isinstance(entry, dict) and set(entry) == {"file", "sha256"} and
            isinstance(entry["file"], str) and entry["file"].strip() and
            is_digest(entry["sha256"]) for entry in results):
        problems.append("results is not a non-empty list of objects with a file name and "
                        "its sha256")
    if not is_digest(item.get("entry_sha256")):
        problems.append("entry_sha256 is not a lower-case SHA-256")
    if not is_number(item.get("threshold")) or item["threshold"] < 0:
        problems.append("threshold is not a finite number of at least 0")
    holdout = item.get("holdout")
    if not isinstance(holdout, list) or not holdout or not all(
            isinstance(entry, dict) and
            set(entry) == {"case", "fixture_sha256", "expected_sha256"} and
            isinstance(entry["case"], str) and is_digest(entry["fixture_sha256"]) and
            is_digest(entry["expected_sha256"]) for entry in holdout):
        problems.append("holdout is not a non-empty list of objects with a case and the "
                        "sha256 of its fixture and of its expected file")
    elif len({entry["case"] for entry in holdout}) != len(holdout):
        problems.append("holdout lists a case twice")
    design = item.get("design")
    if design is not None and not (
            isinstance(design, dict) and set(design) == {"path", "bytes", "sha256"} and
            is_design_path(design["path"]) and is_count(design["bytes"], 0) and
            is_digest(design["sha256"])):
        problems.append("design is not null or an object with a path, its bytes and their "
                        "sha256")
    return problems


def read_exposure(item: Mapping[str, Any]) -> Exposure:
    """An exposure from fields that exposure_problems() found well formed."""
    design = item["design"]
    return Exposure(item["contract"], item["decision"], item["exposures"], item["commit"],
                    item["entry_sha256"], float(item["threshold"]),
                    {entry["case"]: (entry["fixture_sha256"], entry["expected_sha256"])
                     for entry in item["holdout"]},
                    None if design is None else
                    (design["path"], design["bytes"], design["sha256"]))


def read_record(root: Path, text: str | None, report: Report, where: str) -> Record | None:
    """The exposure record in ``text`` (the working tree's when None); None when it is
    malformed or inconsistent, which is a finding: a record is never reset."""
    if text is None:
        data = load_json(root / EXPOSURES, root, report)
    else:
        try:
            data = json.loads(text)
        except ValueError:
            return None
    if data is LOAD_FAILED:
        return None
    if not isinstance(data, dict) or set(data) != EXPOSURES_KEYS or \
            data["schema_version"] != 1 or data["kind"] != "holdout-exposures" or \
            not isinstance(data["exposures"], list):
        report.add(where, "envelope is not schema_version 1, kind 'holdout-exposures' and an "
                          "exposures list")
        return None
    record = Record(raw=list(data["exposures"]))
    usable = True
    counts: dict[str, int] = {}
    decisions: set[tuple[str, str]] = set()
    for index, item in enumerate(data["exposures"]):
        name = item.get("contract") if isinstance(item, dict) else None
        at = f"{where}: exposures[{index}]" + (f" {name}" if isinstance(name, str) else "")
        if not isinstance(item, dict) or set(item) != EXPOSURE_KEYS:
            report.add(at, f"keys are not {', '.join(sorted(EXPOSURE_KEYS))}")
            usable = False
            continue
        problems = exposure_problems(item)
        if not problems:
            contract, decision = item["contract"], item["decision"]
            counts[contract] = counts.get(contract, 0) + 1
            if item["exposures"] != counts[contract]:
                problems.append(f"exposures is {item['exposures']}, but this is exposure "
                                f"{counts[contract]} of {contract} in the record")
            if (contract, decision) in decisions:
                problems.append(f"records a second first exposure of {contract} under "
                                f"decision {decision}")
            decisions.add((contract, decision))
        for problem in problems:
            report.add(at, problem)
        if problems:
            usable = False
            continue
        record.exposures.append(read_exposure(item))
    return record if usable else None


def exposed_under(contract: Contract, exposures: Sequence[Exposure]) -> Exposure | None:
    """The exposure of a frozen contract's holdout under its current decision, if any."""
    if contract.status != "frozen" or not exposures or \
            exposures[-1].decision != contract.decision:
        return None
    return exposures[-1]


def reused_decision(contract: Contract, exposures: Sequence[Exposure]) -> Exposure | None:
    """An exposure, other than the latest, under the frozen contract's current decision: a
    decision used again after a later one, which cannot state the exposures since."""
    if contract.status != "frozen" or exposed_under(contract, exposures) is not None:
        return None
    return next((item for item in exposures if item.decision == contract.decision), None)


def drift_problems(root: Path, contract: Contract, exposures: Sequence[Exposure]) -> list[str]:
    """How a frozen contract differs from its exposure under the same decision: its
    entry, a lower threshold, a holdout fixture or expected file, or a design
    document that no longer begins with the bytes recorded. A new decision ends the
    comparison."""
    last = exposed_under(contract, exposures)
    if last is None:
        reused = reused_decision(contract, exposures)
        return [] if reused is None else [
            f"decision {reused.decision} was already used for exposure {reused.count} "
            f"(results of commit {reused.commit[:12]}), before decision "
            f"{exposures[-1].decision}: a new exposure needs a new decision that states the "
            "earlier exposures"]
    since = f"since its holdout was exposed under decision {last.decision} (exposure " \
            f"{last.count}, results of commit {last.commit[:12]})"
    problems = []
    if contract.entry_sha256 != last.entry_sha256:
        problems.append(f"its entry changed {since}: a change after an exposure needs a new "
                        "decision that states the earlier exposures")
    if contract.threshold < last.threshold:
        problems.append(f"threshold {contract.threshold!r} is lower than {last.threshold!r}, "
                        f"the threshold when its holdout was exposed: a threshold lowered "
                        "after an exposure needs a new decision")
    for case, (fixture, expected) in last.holdout.items():
        if fixture_sha256(root, case) != fixture:
            problems.append(f"the fixture of holdout case {case} changed {since}")
        if expected_sha256(root, case) != expected:
            problems.append(f"the expected file of holdout case {case} changed {since}")
    if last.design is not None:
        path, size, digest = last.design
        data = (root / path).read_bytes() if (root / path).is_file() else b""
        if len(data) < size or bytes_sha256(data[:size]) != digest:
            problems.append(f"design {path} no longer begins with the {size} bytes recorded "
                            f"{since}: it was changed, not amended by appending")
    return problems


def sealed_reason(contract: Contract, exposures: Sequence[Exposure],
                  record: Record | None) -> str | None:
    """Why a contract's holdout is not measured; None when it is."""
    if contract.status != "frozen":
        return "sealed while the contract is draft"
    if record is None:
        return "sealed, because the exposure record is missing or malformed"
    if exposed_under(contract, exposures) is None:
        return f"sealed until its exposure under decision {contract.decision} is recorded"
    return None


def record_problems(root: Path, record: Record, contracts: Sequence[Contract],
                    widening: bool) -> list[str]:
    """An exposure whose contract is gone, and, against the newest vX.Y.Z tag that
    HEAD contains, a record that no longer starts with the exposures recorded there."""
    known = {contract.id for contract in contracts}
    problems = [f"{name} has recorded holdout exposures but no contract in "
                f"{REGISTRY.as_posix()}: a contract whose holdout was exposed is never removed"
                for name in dict.fromkeys(item.contract for item in record.exposures
                                          if item.contract not in known)]
    tag = newest_tag(root) if widening else None
    text = text_or_none(root, "show", f"{tag}:{EXPOSURES.as_posix()}") \
        if tag is not None else None
    tagged = read_record(root, text, Report(), "") if text is not None else None
    if tagged is not None and record.raw[:len(tagged.raw)] != tagged.raw:
        problems.append(f"the exposures recorded at {tag} are not all kept, unchanged, at the "
                        "start of the record: the record only grows")
    return problems


def widening_problems(root: Path, registry: Registry, report: Report) -> list[str]:
    """A contract frozen at the newest tag stays frozen, and is widened only by a new decision."""
    tag = newest_tag(root)
    if tag is None:
        return []
    text = text_or_none(root, "show", f"{tag}:{REGISTRY.as_posix()}")
    tagged = read_registry(root, text, Report(), f"{REGISTRY.as_posix()} at {tag}") \
        if text is not None else None
    if tagged is None:
        return []
    current = {contract.id: contract for contract in registry.contracts}
    problems = []
    for before in tagged.contracts:
        if before.status != "frozen":
            continue
        now = current.get(before.id)
        if now is None or now.status != "frozen":
            problems.append(f"{before.id} was frozen at {tag}; a frozen contract is neither "
                            "removed nor returned to draft")
        elif now.decision == before.decision:
            problems += [f"{before.id} {change} at {tag}; a weaker or different guarantee "
                         "needs a new decision" for change in weakened(before, now)]
    return problems


def weakened(before: Contract, now: Contract) -> list[str]:
    """How ``now`` changes or weakens the guarantee ``before`` froze; tightening is no change.
    The target is only reported, so a change of target is no change of the guarantee."""
    changes = [f"{name} is {getattr(now, name)!r}, not {getattr(before, name)!r}"
               for name in ("calculation", "regime", "metric")
               if getattr(now, name) != getattr(before, name)]
    if now.threshold > before.threshold:
        changes.append(f"threshold {now.threshold!r} is larger than {before.threshold!r}")
    for name, label in (("quantities", "quantities"), ("calibration", "calibration cases"),
                        ("holdout", "holdout cases")):
        dropped = sorted(set(getattr(before, name)) - set(getattr(now, name)))
        if dropped:
            changes.append(f"drops {label} {', '.join(dropped)}")
    return changes


def error(metric: str, actual: float, expected: float) -> float:
    if metric == "absolute":
        return abs(actual - expected)
    if expected == 0:
        return 0.0 if actual == 0 else math.inf
    return abs(actual - expected) / abs(expected)


def read_results(path: Path, root: Path, candidate: str | None,
                 report: Report) -> dict[Key, float] | None:
    """The results of the candidate's run; None when the file is unusable or of another commit."""
    where = str(path)
    data = load_json(path, root, report)
    if data is LOAD_FAILED:
        return None
    if not isinstance(data, dict) or data.keys() != RESULTS_KEYS or \
            data["schema_version"] != 1 or data["kind"] != "test-results" or \
            not isinstance(data["results"], list):
        report.add(where, "envelope is not schema_version 1, kind 'test-results', commit "
                          "and a results list")
        return None
    commit = data["commit"]
    if not isinstance(commit, str) or not SHA.fullmatch(commit):
        report.add(where, "commit is not a full lower-case commit SHA")
        return None
    if candidate is None:
        report.add(where, "the candidate commit is unknown: pass --commit, or run in a Git "
                          "checkout")
        return None
    if commit != candidate:
        report.add(where, f"results are of commit {commit[:12]}, not of the candidate "
                          f"{candidate[:12]}")
        return None
    results: dict[Key, float] = {}
    for index, item in enumerate(data["results"]):
        if not isinstance(item, dict) or item.keys() != RESULT_KEYS or \
                not is_number(item["value"]):
            report.add(where, f"results[{index}] is not case, variant, quantity, scope and a "
                              "finite value")
            continue
        if not isinstance(item["case"], str) or not isinstance(item["quantity"], str) or \
                not isinstance(item["variant"], (str, type(None))):
            report.add(where, f"results[{index}] case and quantity are not text, or its "
                              "variant is not text or null")
            continue
        key = (item["case"], item["variant"], item["quantity"],
               scope_text(item["scope"]))
        if key in results:
            report.add(where, f"results[{index}] repeats {label(key)}")
        results[key] = float(item["value"])
    return results


def merged_results(paths: Sequence[Path], root: Path, candidate: str | None,
                   report: Report) -> dict[Key, float] | None:
    """The results of every file together; None when one is unusable. A result that
    two files both hold is a finding: each output is computed by one suite."""
    merged: dict[Key, float] = {}
    origin: dict[Key, Path] = {}
    usable = True
    for path in paths:
        results = read_results(path, root, candidate, report)
        if results is None:
            usable = False
            continue
        for key, value in results.items():
            if key in origin:
                report.add(str(path), f"{label(key)} is also in {origin[key]}")
                continue
            origin[key] = path
            merged[key] = value
    return merged if usable else None


def summary(contract: Contract, split: str, worst: float, exposures: int | None = None) -> str:
    """One split's measurement: the worst error, the threshold, which blocks, and the
    target, which is only reported; for the holdout, its number of exposures."""
    target = "no target" if contract.target is None else f"target {contract.target!r}"
    line = (f"{contract.id} {split}: worst {contract.metric} error {worst!r}, "
            f"threshold {contract.threshold!r}, {target}")
    if exposures is not None:
        line += f", exposures {exposures}"
    if worst > contract.threshold:
        return f"{line}: threshold exceeded"
    if contract.target is not None and worst > contract.target:
        return f"{line}: target missed"
    return line


def measure(contract: Contract, splits: Mapping[str, list[Output]], results: Mapping[Key, float],
            report: Report, where: str, exposures: int = 0,
            sealed: str | None = None) -> list[str]:
    """Each output against its own tolerance and the contract; one summary line per split.
    Only the threshold blocks: a worst error above the target but within the threshold
    passes, and its line says that the target was missed. A sealed holdout is not
    measured: its line says why, with its number of exposures."""
    lines = []
    for split, outputs in splits.items():
        if split == "holdout" and sealed is not None:
            if contract.holdout:
                lines.append(f"{contract.id} holdout: {sealed}, exposures {exposures}")
            continue
        worst = 0.0
        for output in outputs:
            actual = results.get(output.key)
            if actual is None:
                report.add(where, f"no result for {label(output.key)}")
                continue
            bound = output.bound()
            if abs(actual - output.value) > bound:
                report.add(where, f"{label(output.key)} is {actual!r}, expected "
                                  f"{output.value!r} within {bound!r}")
            worst = max(worst, error(contract.metric, actual, output.value))
        if outputs and worst > contract.threshold:
            report.add(where, f"worst {contract.metric} error {worst!r} over {split} cases "
                              f"exceeds the threshold {contract.threshold!r}")
        if outputs:
            lines.append(summary(contract, split, worst,
                                 exposures if split == "holdout" else None))
    return lines


def check_invariant(root: Path, invariant: Invariant, calculations: Mapping[str, bool],
                    results: Mapping[Key, float] | None, report: Report, where: str) -> list[str]:
    """The invariant on the expected values, then on the results when there are any;
    its measurement line. A break is a finding whatever any contract's status."""
    terms = invariant_terms(root, invariant, calculations, report, where)
    for violation in evaluate(invariant, terms, lambda output: output.value).violations:
        report.add(where, f"the expected values break it: {violation}")
    if results is None:
        return []
    verdict = evaluate(invariant, terms, lambda output: results.get(output.key))
    for output in verdict.missing:
        report.add(where, f"no result for {label(output.key)}")
    for violation in verdict.violations:
        report.add(where, violation)
    return [invariant_line(invariant, verdict)]


def check_contract(root: Path, contract: Contract, min_holdout: int,
                   calculations: Mapping[str, bool], record: Record | None,
                   results: Mapping[Key, float] | None, report: Report) -> list[str]:
    """One contract against its cases and its exposures, and against the results when
    there are any; its measurement lines."""
    at = f"{REGISTRY.as_posix()}: {contract.id}"
    splits = contract_outputs(root, contract, calculations, report, at)
    for problem in frozen_problems(contract, splits["holdout"], min_holdout):
        report.add(at, problem)
    if contract.design is not None and not (root / contract.design).is_file():
        report.add(at, f"design {contract.design} is not a file")
    exposures = record.of(contract.id) if record is not None else []
    for problem in drift_problems(root, contract, exposures):
        report.add(at, problem)
    if results is None:
        return []
    sealed = sealed_reason(contract, exposures, record)
    if sealed is not None and contract.status == "frozen" and record is not None:
        report.add(at, f"its holdout has no exposure recorded under decision "
                       f"{contract.decision}, so it was not measured: measure it once with "
                       f"tools/check_accuracy.py --record-holdout {contract.id}")
    return measure(contract, splits, results, report, at, len(exposures), sealed)


def check(root: Path, results_path: Path | Sequence[Path] | None = None,
          commit: str | None = None,
          widening: bool = True) -> tuple[list[str], int, list[str]]:
    """The findings, the number of contracts and the measurement lines: those of the
    contracts, then those of the invariants.

    ``results_path`` is one results file, or the files of several suites,
    measured together. ``commit`` is the candidate the results must be of;
    HEAD when None.
    ``widening`` False skips the comparison with the newest release tag, for a
    copy of a commit's cases without its Git history.
    Nothing here writes the exposure record: a measurement of a recorded holdout
    is a regression measurement, not an exposure."""
    report = Report()
    paths = [results_path] if isinstance(results_path, Path) else list(results_path or ())
    where = REGISTRY.as_posix()
    present = (root / REGISTRY).is_file()
    registry = read_registry(root, None, report, where)
    record = read_record(root, None, report, EXPOSURES.as_posix()) \
        if (root / EXPOSURES).is_file() else None
    if registry is None:
        if not present and widening:
            # A deleted registry still may not drop a released frozen contract.
            for problem in widening_problems(root, Registry(), report):
                report.add(where, problem)
        if not present and record is not None:
            # Nor may it drop a contract whose holdout was exposed.
            for problem in record_problems(root, record, [], widening):
                report.add(EXPOSURES.as_posix(), problem)
        if paths and not present:
            report.add(where, "is missing, so there is nothing to measure")
        return report.findings, 0, []
    if not (root / EXPOSURES).is_file():
        report.add(EXPOSURES.as_posix(), "is missing: a repository with accuracy contracts "
                                         "keeps it, with an empty exposures list until a "
                                         "holdout is first measured. Restore it from Git: a "
                                         "record is never reset")
    calculations = read_methodology(root, Report()).calculations
    candidate = commit or (text_or_none(root, "rev-parse", "HEAD") or "").strip() or None
    results = merged_results(paths, root, candidate, report) if paths else None
    lines: list[str] = []
    for contract in registry.contracts:
        lines += check_contract(root, contract, registry.min_holdout, calculations, record,
                                results, report)
    for invariant in registry.invariants:
        lines += check_invariant(root, invariant, calculations, results, report,
                                 f"{where}: {invariant.id}")
    for problem in widening_problems(root, registry, report) if widening else []:
        report.add(where, problem)
    if record is not None:
        for problem in record_problems(root, record, registry.contracts, widening):
            report.add(EXPOSURES.as_posix(), problem)
    return report.findings, len(registry.contracts), lines


def exposure_entry(root: Path, contract: Contract, earlier: int, candidate: str,
                   paths: Sequence[Path]) -> dict[str, Any]:
    """The record of the first measurement of a frozen contract's holdout under its
    decision, after ``earlier`` exposures under other decisions."""
    design = None
    if contract.design is not None:
        data = (root / contract.design).read_bytes()
        design = {"path": contract.design, "bytes": len(data), "sha256": bytes_sha256(data)}
    return {
        "contract": contract.id, "decision": contract.decision, "exposures": earlier + 1,
        "commit": candidate,
        "results": [{"file": path.name, "sha256": bytes_sha256(path.read_bytes())}
                    for path in paths],
        "entry_sha256": contract.entry_sha256, "threshold": contract.threshold,
        "holdout": [{"case": case, "fixture_sha256": fixture_sha256(root, case),
                     "expected_sha256": expected_sha256(root, case)}
                    for case in contract.holdout],
        "design": design,
    }


def record_holdout(root: Path, names: Sequence[str], paths: Sequence[Path],
                   commit: str | None = None,
                   widening: bool = True) -> tuple[list[str], list[str]]:
    """Record the first exposure of the holdout of each named frozen contract under its
    decision, before anything measures it: the refusals, when nothing was written,
    else what was recorded. It refuses a missing, malformed or inconsistent record
    and never resets one."""
    where = EXPOSURES.as_posix()
    if not paths:
        return ["--record-holdout needs --results: the results files that measure the "
                "holdout"], []
    if not (root / REGISTRY).is_file():
        return [f"{REGISTRY.as_posix()} is missing, so there is no contract to record"], []
    if not (root / EXPOSURES).is_file():
        return [f"{where} is missing: restore it from Git. A missing record is never taken "
                "to mean that no holdout was exposed"], []
    findings = check(root, (), commit, widening)[0]
    if findings:
        return [f"the check fails, so no holdout is exposed: {item}" for item in findings], []
    registry = read_registry(root, None, Report(), REGISTRY.as_posix())
    record = read_record(root, None, Report(), where)
    assert registry is not None and record is not None  # else check() had a finding
    contracts = {contract.id: contract for contract in registry.contracts}
    refusals: list[str] = []
    chosen: list[Contract] = []
    for name in dict.fromkeys(names):
        contract = contracts.get(name)
        earlier = record.of(name)
        if contract is None:
            refusals.append(f"{name} is not a contract of {REGISTRY.as_posix()}")
        elif contract.status != "frozen":
            refusals.append(f"{name} is draft: its holdout stays sealed while its threshold "
                            "may still change. Freeze it first")
        elif (reused := reused_decision(contract, earlier)) is not None:
            refusals.append(f"exposure of {name} under decision {contract.decision} refused: "
                            f"that decision was already used for exposure {reused.count}, "
                            f"before decision {earlier[-1].decision}. Another exposure needs a "
                            "new decision that states the earlier exposures")
        elif exposed_under(contract, earlier) is not None:
            refusals.append(f"second first exposure of {name} refused: its holdout was "
                            f"exposed under decision {contract.decision} with the results of "
                            f"commit {earlier[-1].commit[:12]} (exposures {earlier[-1].count}). "
                            "Another exposure needs a new decision that states the earlier "
                            "exposures")
        else:
            chosen.append(contract)
    candidate = commit or (text_or_none(root, "rev-parse", "HEAD") or "").strip() or None
    read = Report()
    results = merged_results(paths, root, candidate, read)
    refusals += [f"the results are unusable, so nothing is measured: {item}"
                 for item in read.findings]
    for contract in chosen:
        missing = [label(output.key) for case in contract.holdout
                   for output in case_outputs(root, case, Report())[1]
                   if output.key[2] in contract.quantities and output.key not in (results or {})]
        if missing and results is not None:
            refusals.append(f"{contract.id}: the results hold no value for holdout output(s) "
                            f"{', '.join(missing)}, so nothing is measured")
    if refusals:
        return refusals, []
    assert candidate is not None  # else the results were unusable
    entries = [exposure_entry(root, contract, len(record.of(contract.id)), candidate, paths)
               for contract in chosen]
    document = {"schema_version": 1, "kind": "holdout-exposures",
                "exposures": [*record.raw, *entries]}
    (root / EXPOSURES).write_bytes(
        (json.dumps(document, indent=2, ensure_ascii=False) + "\n").encode("utf-8"))
    return [], [f"exposure {entry['exposures']} of the holdout of {entry['contract']} under "
                f"decision {entry['decision']}, with the results of commit {candidate[:12]}, "
                f"in {where}: commit it" for entry in entries]


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT,
                        help="repository root (default: this checkout)")
    parser.add_argument("--results", type=Path, action="append",
                        help="results file of an Excel run, to measure the contracts; "
                             "repeat it for the results file of each suite")
    parser.add_argument("--commit",
                        help="full SHA of the candidate the results must be of (default: HEAD)")
    parser.add_argument("--record-holdout", action="append", metavar="ACC-ID",
                        help="record the first exposure of this frozen contract's holdout "
                             f"under its decision in {EXPOSURES.as_posix()}, then measure "
                             "with --results; repeat it for each contract")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    recorded: list[str] = []
    if args.record_holdout:
        refusals, recorded = record_holdout(root, args.record_holdout, args.results or (),
                                            args.commit)
        if refusals:
            for refusal in refusals:
                print(f"REFUSED: {refusal}")
            print(f"accuracy: REFUSED (nothing measured; {EXPOSURES.as_posix()} not written)")
            return 2
    findings, contracts, lines = check(root, args.results or (), args.commit)
    registry = read_registry(root, None, Report(), REGISTRY.as_posix())
    invariants = len(registry.invariants) if registry is not None else 0
    for line in lines:
        print(line)
    for finding in findings:
        print(f"FAIL: {finding}")
    for line in recorded:
        print(f"Recorded: {line}")
    status = "FAIL" if findings else "PASS"
    measured = "measured against the results" if args.results else \
        "registry only, nothing measured without --results"
    print(f"accuracy: {status} ({contracts} contract(s), {invariants} invariant(s); "
          f"{measured})")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
