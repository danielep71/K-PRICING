"""The completeness policy of the date layer's worksheet-function cases.

docs/methodology/VALIDATION_PLAN.md, Completeness policy, states which case
categories each function needs and which conditions each of its arguments, and
the function itself, must have a case for. This test reads those tables and the
cases, attributes each condition a case expects to the argument that fails
first at that position, and fails on anything the policy requires that no case
covers. The attribution uses the reference model of
tests/derivations/date_layer_reference.py to parse each argument by its kind.
"""

from __future__ import annotations

import importlib.util
import json
import re
import sys
import unittest
from collections import defaultdict
from pathlib import Path
from types import ModuleType

ROOT = Path(__file__).resolve().parents[1]
PLAN = ROOT / "docs/methodology/VALIDATION_PLAN.md"
FIXTURES = ROOT / "tests/fixtures"
EXPECTED = ROOT / "tests/expected"
MODEL = ROOT / "tests/derivations/date_layer_reference.py"
HOST_FUNCTION = "KPR_Dates_HostDateSystem"
HOST_CONTEXTS = frozenset({"direct", "ws1900", "ws1904"})
CATEGORIES = ("nominal", "boundary", "invalid")
# Conditions of the call, whatever argument holds the shape.
CALL_CONDITIONS = frozenset({"HOST_DATE1904", "SHAPE_UNSUPPORTED", "SHAPE_MISMATCH",
                             "CAPACITY_EXCEEDED"})
CALL = "call"
RESULT = "result"
CODE = re.compile(r"`([^`]+)`")
POSITION = re.compile(r"r([0-9]+)c([0-9]+)")


def load_model() -> ModuleType:
    spec = importlib.util.spec_from_file_location("date_layer_reference", MODEL)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


M = load_model()


def section() -> str:
    text = PLAN.read_text(encoding="utf-8")
    start = text.index('<a id="date-layer-completeness"></a>')
    return text[start:text.index("\n## ", start)]


def rows(header: str) -> list[list[str]]:
    """The rows of the policy table whose header starts with ``header``."""
    lines = section().splitlines()
    first = next(index for index, line in enumerate(lines) if line.startswith(header))
    found = []
    for line in lines[first + 2:]:
        if not line.startswith("|"):
            break
        found.append([cell.strip() for cell in line.strip().strip("|").split("|")])
    return found


def argument_policy() -> dict[str, tuple[str, list[str]]]:
    """Each argument name: its kind and the conditions it needs."""
    policy = {}
    for names, kind, conditions in rows("| Arguments |"):
        for name in CODE.findall(names):
            policy[name] = (kind, CODE.findall(conditions))
    return policy


def result_policy() -> dict[str, list[str]]:
    """The result conditions each function needs."""
    return {CODE.findall(function)[0]: CODE.findall(conditions)
            for function, conditions in rows("| Function | Result conditions |")}


def required() -> dict[str, set[tuple[str, str]]]:
    """Each value-taking function: the (argument or level, condition) pairs it needs."""
    arguments = argument_policy()
    results = result_policy()
    needs: dict[str, set[tuple[str, str]]] = {}
    for function, spec in M.FUNCTIONS.items():
        pairs = {(CALL, condition) for condition in
                 ("HOST_DATE1904", "SHAPE_UNSUPPORTED", "CAPACITY_EXCEEDED")}
        if len(spec.params) > 1:
            pairs.add((CALL, "SHAPE_MISMATCH"))
        for name, _ in (*spec.params, *spec.controls):
            pairs |= {(name, condition) for condition in arguments[name][1]}
        pairs |= {(RESULT, condition) for condition in results.get(function, [])}
        needs[function] = pairs
    return needs


def first_failure(spec: object, items: list[object]) -> tuple[str, str] | None:
    """The first value argument that fails at one position, with its condition."""
    for (name, kind), item in zip(spec.params, items):  # type: ignore[attr-defined]
        parsed = M.parse_value(kind, item)
        if isinstance(parsed, M.Fail):
            return name, str(parsed.condition)
    return None


def attributed(function: str, arguments: list[str], status: str) -> set[tuple[str, str]]:
    """The (argument or level, condition) pairs a case's _status text covers."""
    spec = M.FUNCTIONS[function]
    values = [M.decode(text) for text in arguments][:len(spec.params)]
    pairs = set()
    for part in status.split(";"):
        place, _, condition = part.rpartition("=")
        if condition in CALL_CONDITIONS:
            pairs.add((CALL, condition))
        elif condition.startswith("CONTROL_"):
            pairs |= {(name, condition) for name, _ in spec.controls}
        else:
            match = POSITION.fullmatch(place)
            row, col = (int(match.group(1)) - 1, int(match.group(2)) - 1) if match else (0, 0)
            items = [M.element_at(M.unwrap(value), row, col) for value in values]
            failure = first_failure(spec, items)
            if failure is not None and failure[1] == condition:
                pairs.add(failure)
            elif failure is None:
                # Every argument parsed: the rule that computes the result failed.
                pairs.add((RESULT, condition))
    return pairs


def covered() -> tuple[dict[str, set[tuple[str, str]]], dict[str, dict[str, int]],
                       set[str]]:
    """What the cases cover: the pairs and categories by function, and the callers of
    the host function."""
    pairs: dict[str, set[tuple[str, str]]] = defaultdict(set)
    categories: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    host: set[str] = set()
    for path in sorted(FIXTURES.glob("*.json")):
        fixture = json.loads(path.read_text(encoding="utf-8"))
        if fixture.get("kind") != "test-call-fixture":
            continue
        function = fixture["function"]
        categories[function][fixture["category"]] += 1
        if function == HOST_FUNCTION:
            host.add(fixture["context"])
            continue
        expected = json.loads((EXPECTED / path.name).read_text(encoding="utf-8"))
        for output in expected["outputs"]:
            if output["quantity"].endswith("_status"):
                pairs[function] |= attributed(function, fixture["arguments"], output["text"])
    return pairs, categories, host


def missing() -> list[str]:
    """Everything the policy requires that no case covers."""
    pairs, categories, host = covered()
    gaps = []
    for function, needs in sorted(required().items()):
        for category in CATEGORIES:
            if not categories[function][category]:
                gaps.append(f"{function}: no {category} case")
        for target, condition in sorted(needs - pairs[function]):
            gaps.append(f"{function}: no case for {condition} at {target}")
    for context in sorted(HOST_CONTEXTS - host):
        gaps.append(f"{HOST_FUNCTION}: no case for the {context} caller")
    return gaps


class CompletenessPolicyTest(unittest.TestCase):
    def test_the_policy_names_every_argument_with_the_model_kind(self) -> None:
        policy = argument_policy()
        for function, spec in M.FUNCTIONS.items():
            for name, kind in spec.params:
                self.assertIn(name, policy, function)
                self.assertEqual(policy[name][0], kind, f"{function} {name}")
            for name, _ in spec.controls:
                self.assertTrue(policy[name][0].endswith("control"), f"{function} {name}")

    def test_the_result_table_names_known_functions(self) -> None:
        self.assertLessEqual(set(result_policy()), set(M.FUNCTIONS))

    def test_every_function_meets_the_policy(self) -> None:
        self.assertEqual(missing(), [])

    def test_a_condition_counts_for_the_argument_that_fails_first(self) -> None:
        # Both arguments are malformed: the first, StartDate, decides.
        self.assertEqual(attributed("KPR_Dates_DateFromPillar",
                                    ['str:"x"', 'str:"3X"'], "DATE_TEXT_FORMAT"),
                         {("StartDate", "DATE_TEXT_FORMAT")})
        self.assertEqual(attributed("KPR_Dates_DateFromPillar",
                                    ["date:2024-01-01", 'str:"3X"'], "PILLAR_TOKEN_MALFORMED"),
                         {("Pillar", "PILLAR_TOKEN_MALFORMED")})
        self.assertEqual(attributed("KPR_Dates_AddDays",
                                    ["arr2:1x2[date:2024-01-01;num:1]", "lng:1"],
                                    "r1c2=DATE_WINDOW"),
                         {("DateIn", "DATE_WINDOW")})
        self.assertEqual(attributed("KPR_Dates_AddDays", ["date:9999-12-31", "lng:1"],
                                    "RESULT_WINDOW"),
                         {(RESULT, "RESULT_WINDOW")})


if __name__ == "__main__":
    unittest.main()
