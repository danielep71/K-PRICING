"""Tests for tools/check_data_contract.py: the contract and the reference validator.

The contract and the input files are synthetic, written here rather than read
from the repository, so the pinned finding codes hold whatever datasets a
repository's own contract declares.
"""

from __future__ import annotations

import contextlib
import copy
import io
import json
import sys
import tempfile
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_data_contract  # noqa: E402


def field(name: str, kind: str, required: bool = False, unit: object = None,
          allowed: object = None) -> dict[str, Any]:
    return {"name": name, "type": kind, "unit": unit, "required": required,
            "allowed": allowed, "notes": ""}


DATASET: dict[str, Any] = {
    "id": "DS-02", "illustrative": False, "description": "Positions", "provider": "A system",
    "frequency": "Daily", "granularity": "One row per position", "source": "positions.csv",
    "encoding": "UTF-8", "delimiter": ",", "decimal_separator": ".",
    "fields": [
        field("position_id", "text", required=True),
        field("side", "text", required=True, allowed={"values": ["BUY", "SELL"]}),
        field("currency", "currency"),
        field("notional", "decimal", required=True, unit={"currency_field": "currency"},
              allowed={"min": 0}),
        field("trade_count", "integer", allowed={"min": 0, "max": 100}),
        field("start_date", "date"),
        field("netted", "boolean"),
        field("weight", "decimal", unit="Decimal fraction", allowed={"min": 0, "max": 1}),
    ],
}
CONTRACT: dict[str, Any] = {"schema_version": 1, "kind": "data-contract",
                            "datasets": [DATASET]}
HEADER = "position_id,side,currency,notional\n"


def one(column: str, value: str) -> str:
    """A file with one valid row and one more column."""
    return f"position_id,side,currency,notional,{column}\nP1,BUY,EUR,0,{value}\n"

Edit = Callable[[dict[str, Any]], None]


def parsed(edit: Edit | None = None) -> tuple[list[check_data_contract.Dataset], list[str]]:
    """Parse the synthetic contract after edit(contract), through its JSON text."""
    contract = copy.deepcopy(CONTRACT)
    if edit is not None:
        edit(contract)
    return check_data_contract.parse(check_data_contract.strict_json(json.dumps(contract)))


def dataset(**changes: object) -> check_data_contract.Dataset:
    datasets, problems = parsed(lambda c: c["datasets"][0].update(changes))
    assert not problems, problems
    return datasets[0]


def codes(data: bytes | str, target: check_data_contract.Dataset | None = None
          ) -> list[tuple[str, int | None]]:
    raw = data.encode("utf-8") if isinstance(data, str) else data
    findings = check_data_contract.validate_bytes(raw, target or dataset())
    return [(finding.code, finding.line) for finding in findings]


class ValidFileTest(unittest.TestCase):
    def test_valid_files_have_no_finding(self) -> None:
        semicolon = dataset(delimiter=";", decimal_separator=",")
        utf16 = dataset(encoding="UTF-16")
        files: list[tuple[str, bytes, check_data_contract.Dataset | None]] = [
            ("lf", (HEADER + "P1,BUY,EUR,600000.5\n").encode(), None),
            ("crlf", (HEADER + "P1,BUY,EUR,600000.5\n").replace("\n", "\r\n").encode(), None),
            ("no final newline", (HEADER + "P1,SELL,EUR,0").encode(), None),
            ("header only", HEADER.encode(), None),
            ("any column order, every type",
             b"netted,notional,start_date,position_id,weight,trade_count,side,currency\n"
             b"true,1,2026-01-02,P1,0.25,100,BUY,EUR\nfalse,2,,P2,,0,SELL,USD\n", None),
            ("optional cells empty", (HEADER + "P1,BUY,EUR,1\n").encode().replace(
                b"currency,", b"currency,trade_count,").replace(b"EUR,", b"EUR,,"), None),
            ("quoted cell", (HEADER + '"P,1",BUY,EUR,"2"\n').encode(), None),
            ("semicolon and decimal comma", b"position_id;side;currency;notional\n"
                                            b"P1;BUY;EUR;600000,5\n", semicolon),
            ("UTF-16 with its byte order mark", (HEADER + "P1,BUY,EUR,1\n").encode("utf-16"),
             utf16),
        ]
        for name, data, target in files:
            with self.subTest(name=name):
                self.assertEqual(codes(data, target), [])


class InvalidFileTest(unittest.TestCase):
    # Each file and the codes, with their lines, that it must report.
    PINNED: list[tuple[str, bytes | str, list[tuple[str, int | None]]]] = [
        ("not UTF-8", b"position_id\n\xff\n", [("DC-ENCODING", None)]),
        ("empty", "", [("DC-HEADER", None)]),
        ("quote left open", HEADER + 'P1,BUY,EUR,"1\n', [("DC-SYNTAX", 2)]),
        ("text after a quote", HEADER + 'P1,"BUY"X,EUR,1\n', [("DC-SYNTAX", 2)]),
        ("duplicate column", "position_id,side,notional,side\n",
         [("DC-DUPLICATE-COLUMN", 1)]),
        ("unknown column", "position_id,side,notional,colour\n", [("DC-UNKNOWN-COLUMN", 1)]),
        ("byte order mark kept by UTF-8", "﻿" + HEADER,
         [("DC-UNKNOWN-COLUMN", 1), ("DC-MISSING-COLUMN", 1)]),
        ("missing required column", "position_id,notional\n", [("DC-MISSING-COLUMN", 1)]),
        ("blank line and short row", HEADER + "P1,BUY,EUR,1\n\nP2,SELL\n",
         [("DC-CELL-COUNT", 3), ("DC-CELL-COUNT", 4)]),
        ("required cell empty", HEADER + "P1,,EUR,1\n", [("DC-REQUIRED", 2)]),
        ("exponent", HEADER + "P1,BUY,EUR,1e3\n", [("DC-TYPE", 2)]),
        ("thousands separator", HEADER + 'P1,BUY,EUR,"1,000"\n', [("DC-TYPE", 2)]),
        ("surrounding blank", HEADER + "P1,BUY,EUR, 1\n", [("DC-TYPE", 2)]),
        ("trailing blank", HEADER + "P1,BUY,EUR,1 \n", [("DC-TYPE", 2)]),
        ("digits of another script", HEADER + "P1,BUY,EUR,\u0661\n", [("DC-TYPE", 2)]),
        # A quoted cell may hold a line break, which ends the record on line 3.
        ("decimal with a quoted line break", HEADER + 'P1,BUY,EUR,"1\n"\n', [("DC-TYPE", 3)]),
        ("currency with a quoted line break", HEADER + 'P1,BUY,"EUR\n",1\n',
         [("DC-TYPE", 3)]),
        ("currency with blanks", HEADER + 'P1,BUY," EUR ",1\n', [("DC-TYPE", 2)]),
        ("integer with a quoted line break", one("trade_count", '"1\n"'), [("DC-TYPE", 3)]),
        ("date with a quoted line break", one("start_date", '"2026-01-02\n"'),
         [("DC-TYPE", 3)]),
        ("boolean with a quoted line break", one("netted", '"true\n"'), [("DC-TYPE", 3)]),
        ("allowed value with a quoted line break", HEADER + 'P1,"BUY\n",EUR,1\n',
         [("DC-VALUE", 3)]),
        ("decimal comma in a decimal point file", HEADER + 'P1,BUY,EUR,"1,5"\n',
         [("DC-TYPE", 2)]),
        ("currency in lower case", HEADER + "P1,BUY,eur,1\n", [("DC-TYPE", 2)]),
        ("integer with a fraction", one("trade_count", "1.5"), [("DC-TYPE", 2)]),
        ("impossible date", one("start_date", "2026-02-30"), [("DC-TYPE", 2)]),
        ("boolean as yes", one("netted", "yes"), [("DC-TYPE", 2)]),
        ("value not allowed", HEADER + "P1,HOLD,EUR,1\n", [("DC-VALUE", 2)]),
        ("negative notional", HEADER + "P1,BUY,EUR,-1\n", [("DC-RANGE", 2)]),
        ("count above its maximum", one("trade_count", "101"), [("DC-RANGE", 2)]),
        ("weight above one", one("weight", "1.01"), [("DC-RANGE", 2)]),
        ("amount with an empty currency", HEADER + "P1,BUY,,1\n", [("DC-CURRENCY", 2)]),
        ("amount without a currency column", "position_id,side,notional\nP1,BUY,1\n",
         [("DC-CURRENCY", 2)]),
    ]

    def test_pinned_codes(self) -> None:
        for name, data, expected in self.PINNED:
            with self.subTest(name=name):
                self.assertEqual(codes(data), expected)

    def test_every_code_is_pinned_and_documented(self) -> None:
        pinned = {code for _, _, expected in self.PINNED for code, _ in expected}
        self.assertEqual(pinned, set(check_data_contract.CODES))

    def test_a_finding_names_its_line_code_and_value(self) -> None:
        findings = check_data_contract.validate_bytes((HEADER + "P1,BUY,EUR,-1\n").encode(),
                                                      dataset())
        self.assertEqual([str(finding) for finding in findings],
                         ["line 2: DC-RANGE: field notional value '-1' is not ≥ 0"])

    def test_a_row_reports_each_cell(self) -> None:
        self.assertEqual(codes(HEADER + "P1,HOLD,eur,-1\n"),
                         [("DC-VALUE", 2), ("DC-TYPE", 2), ("DC-RANGE", 2)])


class ContractTest(unittest.TestCase):
    def assertProblem(self, edit: Edit, text: str) -> None:
        datasets, problems = parsed(edit)
        self.assertEqual(datasets, [])
        self.assertTrue(any(text in problem for problem in problems), problems)

    def test_the_synthetic_contract_is_valid(self) -> None:
        datasets, problems = parsed()
        self.assertEqual(problems, [])
        self.assertEqual([field.name for field in datasets[0].fields][:4],
                         ["position_id", "side", "currency", "notional"])

    def test_envelope(self) -> None:
        self.assertProblem(lambda c: c.update(kind="contract"), "envelope does not match")
        self.assertProblem(lambda c: c.update(extra=1), "unknown extra")
        self.assertProblem(lambda c: c.update(datasets=[]), "datasets is not a non-empty list")
        self.assertEqual(check_data_contract.parse([]), ([], ["is not an object"]))

    def test_datasets(self) -> None:
        def ds(**changes: object) -> Edit:
            return lambda c: c["datasets"][0].update(changes)
        cases: list[tuple[Edit, str]] = [
            (ds(id="DS-2a"), "is not a DS- ID"),
            (ds(id="DS-2\n"), "is not a DS- ID"),
            (ds(illustrative="yes"), "illustrative is not true or false"),
            (ds(description="a | b"), "description is not a text for one table cell"),
            (ds(source="`x`.csv"), "source holds a backtick"),
            (ds(encoding="klingon"), "is not a Python codec name"),
            (ds(source="Inputs", delimiter=None, decimal_separator=None),
             "encoding is a codec name for a file, null for a sheet"),
            (ds(delimiter=";;"), "delimiter is one character"),
            (ds(delimiter='"'), "delimiter is one character"),
            (ds(delimiter=",", decimal_separator=","), "decimal_separator is . or ,"),
            (ds(delimiter=None), "decimal_separator is null without a delimiter"),
            (ds(fields=[]), "fields is not a non-empty list"),
            (ds(extra=1), "unknown extra"),
            (lambda c: c["datasets"].append(copy.deepcopy(DATASET)), "id DS-02 appears in two"),
            (lambda c: c["datasets"].append({**copy.deepcopy(DATASET), "id": "DS-03"}),
             "source positions.csv appears in two"),
        ]
        for edit, text in cases:
            with self.subTest(text=text):
                self.assertProblem(edit, text)

    def test_a_sheet_and_a_file_of_another_kind(self) -> None:
        for source in ("Inputs", "positions.xlsx"):
            with self.subTest(source=source):
                target = dataset(source=source, encoding=None, delimiter=None,
                                 decimal_separator=None)
                self.assertFalse(target.validated)
                self.assertEqual(target.is_file, source.endswith(".xlsx"))

    def test_fields(self) -> None:
        def fd(index: int, **changes: object) -> Edit:
            return lambda c: c["datasets"][0]["fields"][index].update(changes)
        cases: list[tuple[Edit, str]] = [
            (fd(0, name="PositionId"), "is not lower snake case"),
            (fd(0, name="position_id\n"), "is not lower snake case"),
            (fd(0, type="money"), "type 'money' is not one of"),
            (fd(0, required="yes"), "required is not true or false"),
            (fd(0, unit=""), "unit is not null, a text or"),
            (fd(0, notes="a\nb"), "notes is not a text for one table cell"),
            (fd(0, allowed={"values": []}), "not a non-empty list of distinct texts"),
            (fd(0, allowed={"values": ["A", "A"]}), "not a non-empty list of distinct texts"),
            (fd(3, allowed={"values": ["1"]}), "allowed values need a type of"),
            (fd(0, allowed={"min": 0}), "a range needs the type integer or decimal"),
            (fd(3, allowed={"min": "0"}), "min and max are not numbers"),
            (fd(3, allowed={"min": 2, "max": 1}), "min is above its max"),
            (fd(3, allowed={"minimum": 0}), "allowed is not null"),
            (fd(3, unit={"currency_field": "side"}), "which is not a currency field"),
            (fd(1, name="position_id"), "field position_id appears twice"),
            (lambda c: c["datasets"][0]["fields"].append("x"), "is not an object"),
            (lambda c: c["datasets"][0]["fields"][0].pop("notes"), "missing notes"),
        ]
        for edit, text in cases:
            with self.subTest(text=text):
                self.assertProblem(edit, text)

    def test_strict_json(self) -> None:
        for raw in ('{"a": 1, "a": 2}', '{"a": NaN}', '{"a": -Infinity}'):
            with self.subTest(raw=raw), self.assertRaises(ValueError):
                check_data_contract.strict_json(raw)

    def test_load(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self.assertEqual(check_data_contract.load(root),
                             ([], ["docs/methodology/data-contract.json: is missing"]))
            path = root / check_data_contract.CONTRACT
            path.parent.mkdir(parents=True)
            path.write_text('{"a": 1, "a": 2}', encoding="utf-8")
            self.assertIn("duplicate key", check_data_contract.load(root)[1][0])
            path.write_text(json.dumps({**CONTRACT, "kind": "x"}), encoding="utf-8")
            self.assertEqual(check_data_contract.load(root)[1],
                             ["docs/methodology/data-contract.json: envelope does not match: "
                              "schema_version 1, kind 'data-contract'"])


class RepositoryTest(unittest.TestCase):
    def test_the_repository_contract_is_valid(self) -> None:
        datasets, problems = check_data_contract.load(ROOT)
        self.assertEqual(problems, [])
        self.assertTrue(datasets)


class MainTest(unittest.TestCase):
    def run_main(self, *args: str, contract: dict[str, Any] | None = None,
                 files: dict[str, bytes] | None = None) -> tuple[int, str]:
        with tempfile.TemporaryDirectory() as directory, \
                contextlib.redirect_stdout(io.StringIO()) as printed:
            root = Path(directory)
            path = root / check_data_contract.CONTRACT
            path.parent.mkdir(parents=True)
            path.write_text(json.dumps(contract or CONTRACT), encoding="utf-8")
            for name, data in (files or {}).items():
                (root / name).write_bytes(data)
            names = [str(root / name) for name in files or {}]
            code = check_data_contract.main(["--root", directory, *args, *names])
        return code, printed.getvalue().replace(directory + "/", "")

    def test_the_contract_alone(self) -> None:
        self.assertEqual(self.run_main(),
                         (0, "data-contract: PASS (1 dataset(s), 0 file(s) validated)\n"))
        code, printed = self.run_main(contract={**CONTRACT, "kind": "x"})
        self.assertEqual(code, 1)
        self.assertIn("FAIL: docs/methodology/data-contract.json: envelope", printed)

    def test_files(self) -> None:
        good, bad = (HEADER + "P1,BUY,EUR,1\n").encode(), (HEADER + "P1,BUY,EUR,-1\n").encode()
        self.assertEqual(self.run_main("--dataset", "DS-02", files={"good.csv": good})[0], 0)
        code, printed = self.run_main("--dataset", "DS-02",
                                      files={"good.csv": good, "bad.csv": bad})
        self.assertEqual(code, 1)
        self.assertIn("FAIL: bad.csv: line 2: DC-RANGE: field notional value '-1'", printed)
        self.assertNotIn("good.csv", printed)
        self.assertIn("data-contract: FAIL (1 dataset(s), 2 file(s) validated)", printed)
        code, printed = self.run_main("--dataset", "DS-09", files={"good.csv": good})
        self.assertEqual(code, 1)
        self.assertIn("DS-09: is not a delimited-file dataset", printed)

    def test_dataset_and_files_go_together(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            check_data_contract.main(["--dataset", "DS-02"])


if __name__ == "__main__":
    unittest.main()
