#!/usr/bin/env python3
"""Check the machine-readable data contract, and input files against it.

docs/methodology/data-contract.json lists the datasets the workbook reads and
the fields of each: type, unit or currency field, whether it is required, and
its allowed values or range (docs/methodology/DATA_CONTRACT.md, Machine-readable
contract). tools/render_docs.py renders the Datasets table and the field tables
of DATA_CONTRACT.md from it.

With no file named, this checks the contract itself. With --dataset DS-NN and
one or more files, it also validates each file against that dataset: it is the
reference validator for delimited text files, and every finding carries one of
the stable codes in CODES. tools/check_test_cases.py runs the same validator on
every variant file of a file-based case outside an invalid case.

The validator checks the layout and the values of a file, never what they
mean: no calculation runs and no total is reconciled.
"""

from __future__ import annotations

import argparse
import codecs
import csv
import datetime
import io
import json
import re
import sys
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
CONTRACT = Path("docs/methodology/data-contract.json")
DEFAULT_ENCODING = "UTF-8"

# Each pattern must match a whole value (fullmatch): a quoted cell may hold a
# line break, which "$" would accept at its end. Digits are ASCII: \d and
# Decimal() would also accept other scripts' digits.
DATASET_ID = re.compile(r"DS-[0-9]+")
SNAKE = re.compile(r"[a-z][a-z0-9]*(?:_[a-z0-9]+)*")
CURRENCY = re.compile(r"[A-Z]{3}")
DATE = re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}")
INTEGER = re.compile(r"[+-]?[0-9]+")

CONTRACT_KEYS = frozenset({"schema_version", "kind", "datasets"})
DATASET_KEYS = frozenset({"id", "illustrative", "description", "provider", "frequency",
                          "granularity", "source", "encoding", "delimiter",
                          "decimal_separator", "fields"})
FIELD_KEYS = frozenset({"name", "type", "unit", "required", "allowed", "notes"})
# The texts rendered into the cells of a Markdown table.
CELL_KEYS = ("description", "provider", "frequency", "granularity", "source")
TYPES = frozenset({"text", "integer", "decimal", "date", "boolean", "currency"})
VALUE_TYPES = frozenset({"text", "integer", "date", "currency"})
RANGE_TYPES = frozenset({"integer", "decimal"})
DECIMAL_SEPARATORS = frozenset({".", ","})
# How each type is shown in the field tables of DATA_CONTRACT.md.
TYPE_NAMES = {"text": "Text", "integer": "Integer", "decimal": "Decimal",
              "date": "Date, `YYYY-MM-DD`", "boolean": "`true` or `false`",
              "currency": "Currency code, ISO 4217"}

# The finding codes of the validator. A code is never reused for another meaning.
CODES: dict[str, str] = {
    "DC-ENCODING": "The file does not decode in the dataset's encoding",
    "DC-SYNTAX": "The file is not readable as delimited text, such as a quote left open",
    "DC-HEADER": "The file has no header row",
    "DC-DUPLICATE-COLUMN": "The header names a column twice",
    "DC-UNKNOWN-COLUMN": "The header names a column that is not a field of the dataset",
    "DC-MISSING-COLUMN": "A required field has no column",
    "DC-CELL-COUNT": "A row, a blank line included, has more or fewer cells than the header",
    "DC-REQUIRED": "A required field is empty",
    "DC-TYPE": "A value is not of its field's type",
    "DC-VALUE": "A value is not one of its field's allowed values",
    "DC-RANGE": "A number is outside its field's range",
    "DC-CURRENCY": "An amount is given without a currency in its currency field",
}


@dataclass(frozen=True)
class Field:
    name: str
    type: str
    # A unit stated as text, or the field that holds the currency of an amount.
    unit: str | None
    currency_field: str | None
    required: bool
    values: tuple[str, ...] | None
    minimum: Decimal | None
    maximum: Decimal | None
    notes: str


@dataclass(frozen=True)
class Dataset:
    id: str
    illustrative: bool
    description: str
    provider: str
    frequency: str
    granularity: str
    # A file name, with an extension, or a sheet name, without one.
    source: str
    encoding: str | None
    # None for a sheet, or for a file that is not delimited text.
    delimiter: str | None
    decimal_separator: str | None
    fields: tuple[Field, ...]

    @property
    def is_file(self) -> bool:
        return "." in self.source

    @property
    def validated(self) -> bool:
        """Whether the reference validator reads it: a delimited text file."""
        return self.delimiter is not None


@dataclass(frozen=True)
class Finding:
    code: str
    line: int | None
    message: str

    def __str__(self) -> str:
        where = f"line {self.line}: " if self.line is not None else ""
        return f"{where}{self.code}: {self.message}"


# ---------------------------------------------------------------------------
# The contract.

def strict_json(text: str) -> Any:
    """Parse JSON without duplicate keys or non-finite numbers; decimals stay exact."""

    def no_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"duplicate key {key!r}")
            result[key] = value
        return result

    def no_constants(name: str) -> None:
        raise ValueError(f"non-standard number {name}")

    def finite(text: str) -> Decimal:
        value = Decimal(text)
        if not value.is_finite():
            raise ValueError(f"number {text} is not finite")
        return value

    return json.loads(text, object_pairs_hook=no_duplicates, parse_constant=no_constants,
                      parse_float=finite)


def keys_problem(data: Mapping[str, Any], allowed: frozenset[str]) -> str | None:
    missing = sorted(allowed - data.keys())
    unknown = sorted(data.keys() - allowed)
    parts = ([f"missing {', '.join(missing)}"] if missing else []) + (
        [f"unknown {', '.join(unknown)}"] if unknown else [])
    return "; ".join(parts) or None


def is_cell_text(value: object) -> bool:
    """A non-empty text that fits in one Markdown table cell."""
    return isinstance(value, str) and value.strip() != "" and not set(value) & set("|\r\n")


def is_number(value: object) -> bool:
    return isinstance(value, (int, Decimal)) and not isinstance(value, bool)


def parse_unit(unit: object, where: str,
               problems: list[str]) -> tuple[str | None, str | None]:
    """A unit text, or {"currency_field": name}; return (text, currency field)."""
    if unit is None:
        return None, None
    if is_cell_text(unit):
        assert isinstance(unit, str)
        return unit, None
    if isinstance(unit, dict) and set(unit) == {"currency_field"} and \
            isinstance(unit["currency_field"], str):
        return None, unit["currency_field"]
    problems.append(f"{where}: unit is not null, a text or {{\"currency_field\": name}}")
    return None, None


def parse_allowed(allowed: object, kind: object, where: str, problems: list[str]
                  ) -> tuple[tuple[str, ...] | None, Decimal | None, Decimal | None]:
    """Allowed values ({"values": [...]}) or a range ({"min": x, "max": y}), or null."""
    if allowed is None:
        return None, None, None
    if isinstance(allowed, dict) and set(allowed) == {"values"}:
        values = allowed["values"]
        if kind not in VALUE_TYPES:
            problems.append(f"{where}: allowed values need a type of "
                            f"{', '.join(sorted(VALUE_TYPES))}")
        elif not isinstance(values, list) or not values or \
                not all(is_cell_text(value) for value in values) or \
                len(set(values)) != len(values):
            problems.append(f"{where}: allowed values are not a non-empty list of distinct "
                            "texts")
        else:
            return tuple(values), None, None
        return None, None, None
    if isinstance(allowed, dict) and allowed and set(allowed) <= {"min", "max"}:
        bounds = [allowed.get("min"), allowed.get("max")]
        if kind not in RANGE_TYPES:
            problems.append(f"{where}: a range needs the type integer or decimal")
        elif not all(bound is None or is_number(bound) for bound in bounds):
            problems.append(f"{where}: the range's min and max are not numbers")
        else:
            low, high = (None if bound is None else Decimal(bound) for bound in bounds)
            if low is not None and high is not None and low > high:
                problems.append(f"{where}: the range's min is above its max")
            return None, low, high
        return None, None, None
    problems.append(f"{where}: allowed is not null, {{\"values\": [...]}} or "
                    "{\"min\": x, \"max\": y}")
    return None, None, None


def parse_field(data: object, where: str, problems: list[str]) -> Field | None:
    if not isinstance(data, dict):
        problems.append(f"{where}: is not an object")
        return None
    if problem := keys_problem(data, FIELD_KEYS):
        problems.append(f"{where}: {problem}")
        return None
    before = len(problems)
    name, kind = data["name"], data["type"]
    if not isinstance(name, str) or not SNAKE.fullmatch(name):
        problems.append(f"{where}: name {name!r} is not lower snake case")
    if kind not in TYPES:
        problems.append(f"{where}: type {kind!r} is not one of {', '.join(sorted(TYPES))}")
    unit, currency_field = parse_unit(data["unit"], where, problems)
    if not isinstance(data["required"], bool):
        problems.append(f"{where}: required is not true or false")
    values, minimum, maximum = parse_allowed(data["allowed"], kind, where, problems)
    notes = data["notes"]
    if not isinstance(notes, str) or set(notes) & set("|\r\n"):
        problems.append(f"{where}: notes is not a text for one table cell")
    if len(problems) > before:
        return None
    return Field(name, kind, unit, currency_field, data["required"], values, minimum, maximum,
                 notes)


def format_problems(data: Mapping[str, Any], file: bool, where: str) -> list[str]:
    """The encoding, delimiter and decimal separator of a dataset's file or sheet."""
    problems = []
    encoding, delimiter, separator = (data["encoding"], data["delimiter"],
                                      data["decimal_separator"])
    if encoding is not None and (not isinstance(encoding, str) or not file):
        problems.append(f"{where}: encoding is a codec name for a file, null for a sheet")
    elif encoding is not None:
        try:
            codecs.lookup(encoding)
        except LookupError:
            problems.append(f"{where}: encoding {encoding!r} is not a Python codec name")
    if delimiter is None:
        if separator is not None:
            problems.append(f"{where}: decimal_separator is null without a delimiter")
        return problems
    if not file or not isinstance(delimiter, str) or len(delimiter) != 1 or \
            delimiter in "`\"\r\n":
        problems.append(f"{where}: delimiter is one character for a delimited file, or null")
    elif separator not in DECIMAL_SEPARATORS or separator == delimiter:
        problems.append(f"{where}: decimal_separator is . or , and differs from the delimiter")
    return problems


def currency_problems(fields: Sequence[Field], where: str) -> list[str]:
    """Each currency field an amount names is a currency field of the dataset."""
    kinds = {field.name: field.type for field in fields}
    return [f"{where}: field {field.name} names the currency field {field.currency_field!r}, "
            "which is not a currency field of the dataset"
            for field in fields if field.currency_field is not None
            and kinds.get(field.currency_field) != "currency"]


def parse_dataset(data: object, where: str, problems: list[str]) -> Dataset | None:
    if not isinstance(data, dict):
        problems.append(f"{where}: is not an object")
        return None
    if problem := keys_problem(data, DATASET_KEYS):
        problems.append(f"{where}: {problem}")
        return None
    before = len(problems)
    if not isinstance(data["id"], str) or not DATASET_ID.fullmatch(data["id"]):
        problems.append(f"{where}: id {data['id']!r} is not a DS- ID")
    if not isinstance(data["illustrative"], bool):
        problems.append(f"{where}: illustrative is not true or false")
    problems += [f"{where}: {key} is not a text for one table cell" for key in CELL_KEYS
                 if not is_cell_text(data[key])]
    if "`" in str(data["source"]):
        problems.append(f"{where}: source holds a backtick")
    problems += format_problems(data, "." in str(data["source"]), where)
    items = data["fields"]
    if not isinstance(items, list) or not items:
        problems.append(f"{where}: fields is not a non-empty list")
        items = []
    fields = [parse_field(item, f"{where}.fields[{index}]", problems)
              for index, item in enumerate(items)]
    valid = [field for field in fields if field is not None]
    names = [field.name for field in valid]
    problems += [f"{where}: field {name} appears twice" for name in sorted(set(names))
                 if names.count(name) > 1]
    problems += currency_problems(valid, where)
    if len(problems) > before:
        return None
    return Dataset(data["id"], data["illustrative"], data["description"], data["provider"],
                   data["frequency"], data["granularity"], data["source"], data["encoding"],
                   data["delimiter"], data["decimal_separator"], tuple(valid))


def parse(data: object) -> tuple[list[Dataset], list[str]]:
    """The datasets of a parsed contract, and its problems; no dataset when it has any."""
    if not isinstance(data, dict):
        return [], ["is not an object"]
    if problem := keys_problem(data, CONTRACT_KEYS):
        return [], [problem]
    problems = []
    if data["schema_version"] != 1 or data["kind"] != "data-contract":
        problems.append("envelope does not match: schema_version 1, kind 'data-contract'")
    items = data["datasets"]
    if not isinstance(items, list) or not items:
        return [], [*problems, "datasets is not a non-empty list"]
    datasets = [parse_dataset(item, f"datasets[{index}]", problems)
                for index, item in enumerate(items)]
    valid = [dataset for dataset in datasets if dataset is not None]
    for key in ("id", "source"):
        seen = [getattr(dataset, key) for dataset in valid]
        problems += [f"{key} {value} appears in two datasets" for value in sorted(set(seen))
                     if seen.count(value) > 1]
    return ([] if problems else valid), problems


def load(root: Path) -> tuple[list[Dataset], list[str]]:
    """Read the contract under root; each problem is prefixed with its path."""
    path = root / CONTRACT
    where = CONTRACT.as_posix()
    if not path.is_file():
        return [], [f"{where}: is missing"]
    try:
        data = strict_json(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, ValueError) as error:
        return [], [f"{where}: not valid UTF-8 JSON ({error})"]
    datasets, problems = parse(data)
    return datasets, [f"{where}: {problem}" for problem in problems]


def unit_text(field: Field) -> str:
    """A field's unit or currency as the field tables show it."""
    if field.currency_field is not None:
        return f"Currency in field `{field.currency_field}`"
    return field.unit or "—"


def allowed_text(field: Field) -> str:
    """A field's allowed values or inclusive range as the field tables show it."""
    if field.values is not None:
        return ", ".join(f"`{value}`" for value in field.values)
    if field.minimum is not None and field.maximum is not None:
        return f"{field.minimum} to {field.maximum}"
    if field.minimum is not None:
        return f"≥ {field.minimum}"
    if field.maximum is not None:
        return f"≤ {field.maximum}"
    return "Any"


# ---------------------------------------------------------------------------
# The reference validator.

def is_date(value: object) -> bool:
    """Whether the value is a real calendar date written YYYY-MM-DD."""
    if not isinstance(value, str) or not DATE.fullmatch(value):
        return False
    try:
        datetime.date.fromisoformat(value)
    except ValueError:
        return False
    return True


def number(text: str, field: Field, separator: str) -> Decimal | None:
    """The value of an integer or decimal cell, or None when it is not one."""
    if field.type == "integer":
        return Decimal(text) if INTEGER.fullmatch(text) else None
    # Only a full match reaches Decimal(), which would accept surrounding blanks.
    pattern = rf"[+-]?[0-9]+(?:{re.escape(separator)}[0-9]+)?"
    return Decimal(text.replace(separator, ".")) if re.fullmatch(pattern, text) else None


def value_findings(text: str, field: Field, separator: str, line: int) -> list[Finding]:
    """Check one non-empty cell against its field's type and allowed values or range."""
    shown = f"field {field.name} value {text!r}"
    if field.type in RANGE_TYPES:
        value = number(text, field, separator)
        if value is None:
            return [Finding("DC-TYPE", line, f"{shown} is not {field.type}")]
        if (field.minimum is not None and value < field.minimum) or \
                (field.maximum is not None and value > field.maximum):
            return [Finding("DC-RANGE", line, f"{shown} is not {allowed_text(field)}")]
    valid = {"date": is_date(text), "boolean": text in ("true", "false"),
             "currency": bool(CURRENCY.fullmatch(text))}.get(field.type, True)
    if not valid:
        return [Finding("DC-TYPE", line, f"{shown} is not {field.type}")]
    if field.values is not None and text not in field.values:
        return [Finding("DC-VALUE", line, f"{shown} is not one of "
                                          f"{', '.join(field.values)}")]
    return []


def header_findings(header: Sequence[str], dataset: Dataset) -> list[Finding]:
    names = {field.name for field in dataset.fields}
    findings = [Finding("DC-DUPLICATE-COLUMN", 1, f"column {name!r} appears twice")
                for name in sorted(set(header)) if header.count(name) > 1]
    findings += [Finding("DC-UNKNOWN-COLUMN", 1, f"column {name!r} is not a field of "
                                                 f"{dataset.id}")
                 for name in dict.fromkeys(header) if name not in names]
    findings += [Finding("DC-MISSING-COLUMN", 1, f"required field {field.name} has no column")
                 for field in dataset.fields if field.required and field.name not in header]
    return findings


def row_findings(row: Mapping[str, str], dataset: Dataset, line: int) -> list[Finding]:
    """Check the cells of one row, by field name."""
    findings = []
    separator = dataset.decimal_separator or "."
    required = {field.name for field in dataset.fields if field.required}
    for field in dataset.fields:
        text = row.get(field.name, "")
        if text == "":
            if field.required and field.name in row:
                findings.append(Finding("DC-REQUIRED", line, f"required field {field.name} "
                                                             "is empty"))
            continue
        findings += value_findings(text, field, separator, line)
        currency = field.currency_field
        if currency is not None and currency not in required and row.get(currency, "") == "":
            findings.append(Finding("DC-CURRENCY", line, f"field {field.name} has no currency "
                                                         f"in field {currency}"))
    return findings


def validate_text(text: str, dataset: Dataset) -> list[Finding]:
    """Validate decoded delimited text: a header row of field names, then one row per record.

    Columns may come in any order; an empty cell is a missing value. Line
    endings, LF or CRLF, are both read; a cell may be quoted with double quotes.
    """
    assert dataset.delimiter is not None, "only a delimited text file is validated"
    reader = csv.reader(io.StringIO(text, newline=""), delimiter=dataset.delimiter,
                        strict=True)
    findings: list[Finding] = []
    header: list[str] | None = None
    try:
        for cells in reader:
            line = reader.line_num
            if header is None:
                header = cells
                findings += header_findings(header, dataset)
            elif len(cells) != len(header):
                findings.append(Finding("DC-CELL-COUNT", line, f"{len(cells)} cell(s), the "
                                                               f"header has {len(header)}"))
            else:
                columns = dict(reversed(list(zip(header, cells, strict=True))))
                findings += row_findings(columns, dataset, line)
    except csv.Error as error:
        findings.append(Finding("DC-SYNTAX", reader.line_num, str(error)))
    if header is None and not any(finding.code == "DC-SYNTAX" for finding in findings):
        findings.append(Finding("DC-HEADER", None, "the file is empty"))
    return findings


def validate_bytes(data: bytes, dataset: Dataset) -> list[Finding]:
    """Decode a file in its dataset's encoding, then validate it."""
    encoding = dataset.encoding or DEFAULT_ENCODING
    try:
        text = data.decode(encoding)
    except UnicodeDecodeError as error:
        return [Finding("DC-ENCODING", None, f"is not {encoding} ({error.reason} at byte "
                                             f"{error.start})")]
    return validate_text(text, dataset)


# ---------------------------------------------------------------------------
# Command line.

def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("files", nargs="*", type=Path, metavar="FILE",
                        help="input files to validate against --dataset")
    parser.add_argument("--dataset", metavar="DS-NN", help="the dataset the files belong to")
    parser.add_argument("--root", type=Path, default=ROOT,
                        help="repository root (default: this checkout)")
    args = parser.parse_args(argv)
    if bool(args.files) != bool(args.dataset):
        parser.error("--dataset and at least one FILE go together")
    datasets, problems = load(args.root.resolve())
    findings = list(problems)
    by_id = {dataset.id: dataset for dataset in datasets}
    if args.files and not problems:
        dataset = by_id.get(args.dataset)
        if dataset is None or not dataset.validated:
            findings.append(f"{args.dataset}: is not a delimited-file dataset of "
                            f"{CONTRACT.as_posix()}")
        else:
            for path in args.files:
                try:
                    data = path.read_bytes()
                except OSError as error:
                    findings.append(f"{path}: cannot be read ({error.strerror})")
                    continue
                findings += [f"{path}: {finding}" for finding in validate_bytes(data, dataset)]
    for finding in findings:
        print(f"FAIL: {finding}")
    status = "FAIL" if findings else "PASS"
    print(f"data-contract: {status} ({len(datasets)} dataset(s), {len(args.files)} file(s) "
          "validated)")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
