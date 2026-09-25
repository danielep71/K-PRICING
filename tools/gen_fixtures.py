#!/usr/bin/env python3
"""Generate the independent date-layer fixtures and the generated VBA fixture module.

The expected results are computed by a reference model written from
docs/DATE_LAYER_CONTRACT.md with Python standard-library date arithmetic. The
model never imports, executes or translates the VBA implementation. The
canonical TSV is written first; the VBA module is then derived from the parsed
TSV, so the TSV remains the reviewed source of truth.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import re
import shutil
import sys
import tempfile
from dataclasses import dataclass, field
from decimal import ROUND_FLOOR, Decimal, InvalidOperation
from pathlib import Path
from typing import Callable, Iterable, NamedTuple, Union

TSV_PATH = "tests/fixtures/date_layer_fixtures.tsv"
VBA_PATH = "tests/modules/KPR_Test_Fixtures_Generated.bas"
CONTRACT_PATH = "docs/DATE_LAYER_CONTRACT.md"
MODULE_NAME = "KPR_Test_Fixtures_Generated"
GENERATOR_UPDATED = "2026-09-25"
BLOCK_SIZE = 40
MAX_VBA_LINE = 1000

COLUMNS = (
    "id", "suite", "function", "context", "arg1", "arg2", "arg3", "arg4", "arg5",
    "input_kind", "expect_type", "expected", "condition", "level", "rationale",
)
CONTEXTS = ("direct", "ws1900", "ws1904")
LEVELS = ("", "element", "call")

MIN_DATE = dt.date(1900, 3, 1)
MAX_DATE = dt.date(9999, 12, 31)
MIN_ORD = MIN_DATE.toordinal()
MAX_ORD = MAX_DATE.toordinal()
EPOCH_ORD = dt.date(1899, 12, 30).toordinal()
LONG_MIN = -2147483648
LONG_MAX = 2147483647
CAPACITY = 100000

ERRORS = {
    "#NULL!": 2000, "#DIV/0!": 2007, "#VALUE!": 2015, "#REF!": 2023,
    "#NAME?": 2029, "#NUM!": 2036, "#N/A": 2042,
}
ERROR_CONSTANTS = {
    "#NULL!": "FX_ERR_NULL", "#DIV/0!": "FX_ERR_DIV0", "#VALUE!": "FX_ERR_VALUE",
    "#REF!": "FX_ERR_REF", "#NAME?": "FX_ERR_NAME", "#NUM!": "FX_ERR_NUM", "#N/A": "FX_ERR_NA",
}

# Conditions the registry defines but no deterministic fixture can express.
UNFIXTURABLE = {
    "HOST_UNRESOLVED": "requires a worksheet host whose date system cannot be read; "
                       "only failure injection can produce it",
}


class FixtureError(Exception):
    """A fixture, encoding or model inconsistency."""


# ---------------------------------------------------------------------------
# Value model and canonical text encoding
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Val:
    """One VBA value: a scalar, a 1-D array, a 2-D array or a filled 2-D array."""

    kind: str
    payload: str = ""
    rows: int = 0
    cols: int = 0
    items: tuple[Val, ...] = ()


SCALAR_KINDS = {"date", "datetime", "num", "lng", "str", "bool", "empty", "null", "err", "missing"}
ARRAY_KINDS = {"arr1", "arr2", "fill2"}
NUMBER = re.compile(r"-?[0-9]+(?:\.[0-9]+)?(?:E[+-][0-9]+)?")


def D(text: str) -> Val:
    dt.date.fromisoformat(text)
    return Val("date", text)


def DT(text: str) -> Val:
    dt.datetime.fromisoformat(text)
    return Val("datetime", text)


def N(text: str) -> Val:
    if not NUMBER.fullmatch(text):
        raise FixtureError(f"non-canonical number literal: {text!r}")
    return Val("num", text)


def L(value: int) -> Val:
    if not LONG_MIN <= value <= LONG_MAX:
        raise FixtureError(f"Long literal out of range: {value}")
    return Val("lng", str(value))


def S(text: str) -> Val:
    if any(ord(ch) < 32 or ord(ch) > 126 for ch in text):
        raise FixtureError(f"fixture text must be printable ASCII: {text!r}")
    return Val("str", text)


def B(value: bool) -> Val:
    return Val("bool", "TRUE" if value else "FALSE")


def E(code: str) -> Val:
    if code not in ERRORS:
        raise FixtureError(f"unknown Excel error: {code}")
    return Val("err", code)


def A2(rows: int, cols: int, items: Iterable[Val]) -> Val:
    values = tuple(items)
    if rows < 1 or cols < 1 or len(values) != rows * cols:
        raise FixtureError("2-D array item count must equal rows * cols")
    return Val("arr2", rows=rows, cols=cols, items=values)


def A1(items: Iterable[Val]) -> Val:
    return Val("arr1", items=tuple(items))


def F2(rows: int, cols: int, item: Val) -> Val:
    return Val("fill2", rows=rows, cols=cols, items=(item,))


EMPTY = Val("empty")
NULL = Val("null")
MISSING = Val("missing")


def encode(value: Val) -> str:
    kind = value.kind
    if kind in {"empty", "null", "missing"}:
        return kind
    if kind == "str":
        return "str:" + json.dumps(value.payload, ensure_ascii=True)
    if kind in SCALAR_KINDS:
        return f"{kind}:{value.payload}"
    inner = ";".join(encode(item) for item in value.items)
    if kind == "arr1":
        return f"arr1:[{inner}]"
    return f"{kind}:{value.rows}x{value.cols}[{inner}]"


class _Decoder:
    def __init__(self, text: str) -> None:
        self.text = text
        self.pos = 0

    def value(self) -> Val:
        for word in ("empty", "null", "missing"):
            if self.text.startswith(word, self.pos) and not self.text.startswith(word + ":", self.pos):
                self.pos += len(word)
                return {"empty": EMPTY, "null": NULL, "missing": MISSING}[word]
        colon = self.text.find(":", self.pos)
        if colon < 0:
            raise FixtureError(f"malformed value at {self.pos}: {self.text!r}")
        kind = self.text[self.pos:colon]
        self.pos = colon + 1
        if kind == "str":
            payload, end = json.JSONDecoder().raw_decode(self.text, self.pos)
            self.pos = end
            return S(payload)
        if kind in ARRAY_KINDS:
            return self.array(kind)
        match = re.compile(r"[^;\]]*").match(self.text, self.pos)
        assert match is not None
        self.pos = match.end()
        return self.scalar(kind, match.group(0))

    @staticmethod
    def scalar(kind: str, payload: str) -> Val:
        builders: dict[str, Callable[[str], Val]] = {
            "date": D, "datetime": DT, "num": N, "lng": lambda p: L(int(p)),
            "bool": lambda p: B({"TRUE": True, "FALSE": False}[p]), "err": E,
        }
        if kind not in builders:
            raise FixtureError(f"unknown value kind: {kind!r}")
        return builders[kind](payload)

    def array(self, kind: str) -> Val:
        rows = cols = 0
        if kind != "arr1":
            match = re.compile(r"([0-9]+)x([0-9]+)").match(self.text, self.pos)
            if match is None:
                raise FixtureError(f"missing dimensions for {kind}")
            rows, cols = int(match.group(1)), int(match.group(2))
            self.pos = match.end()
        self.expect("[")
        items: list[Val] = []
        if not self.text.startswith("]", self.pos):
            items.append(self.value())
            while self.text.startswith(";", self.pos):
                self.pos += 1
                items.append(self.value())
        self.expect("]")
        if kind == "arr1":
            return A1(items)
        if kind == "arr2":
            return A2(rows, cols, items)
        if len(items) != 1:
            raise FixtureError("fill2 carries exactly one item")
        return F2(rows, cols, items[0])

    def expect(self, token: str) -> None:
        if not self.text.startswith(token, self.pos):
            raise FixtureError(f"expected {token!r} at {self.pos} in {self.text!r}")
        self.pos += len(token)


def decode(text: str) -> Val:
    decoder = _Decoder(text)
    value = decoder.value()
    if decoder.pos != len(text):
        raise FixtureError(f"trailing text in value: {text!r}")
    return value


# ---------------------------------------------------------------------------
# Reference model of the date-layer contract
# ---------------------------------------------------------------------------


class Fail(NamedTuple):
    error: str
    condition: str


Parsed = Union[dt.date, int, tuple[int, int], Fail]


def fail(error: str, condition: str) -> Fail:
    return Fail(error, condition)


def in_window(ordinal: int) -> bool:
    return MIN_ORD <= ordinal <= MAX_ORD


def days_in_month(year: int, month: int) -> int:
    if month == 12:
        return 31
    return (dt.date(year, month + 1, 1) - dt.date(year, month, 1)).days


def is_leap(year: int) -> bool:
    return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)


ISO = re.compile(r"([0-9]{4})-([0-9]{2})-([0-9]{2})")
LOCALE = re.compile(r"[0-9]{1,2}[/.][0-9]{1,2}[/.][0-9]{2,4}")
NUMERIC_TEXT = re.compile(r"[+-]?[0-9]+(?:\.[0-9]+)?")


def parse_date(value: Val) -> dt.date | Fail:
    """Section 3.1: the date-value matrix under the direct 1900 serial contract."""
    kind = value.kind
    if kind == "err":
        return fail(value.payload, "INPUT_ERROR_PROPAGATED")
    if kind == "empty":
        return fail("#VALUE!", "INPUT_BLANK_REQUIRED")
    if kind in {"null", "bool"}:
        return fail("#VALUE!", "DATE_TYPE_REJECTED")
    if kind in {"date", "datetime"}:
        ordinal = dt.datetime.fromisoformat(value.payload).date().toordinal()
    elif kind in {"num", "lng"}:
        serial = Decimal(value.payload).to_integral_value(rounding=ROUND_FLOOR)
        if serial < MIN_ORD - EPOCH_ORD or serial > MAX_ORD - EPOCH_ORD:
            return fail("#NUM!", "DATE_WINDOW")
        ordinal = EPOCH_ORD + int(serial)
    elif kind == "str":
        return parse_date_text(value.payload)
    else:
        raise FixtureError(f"unsupported date input kind: {kind}")
    if not in_window(ordinal):
        return fail("#NUM!", "DATE_WINDOW")
    return dt.date.fromordinal(ordinal)


def parse_date_text(text: str) -> dt.date | Fail:
    match = ISO.fullmatch(text)
    if match is None:
        if LOCALE.fullmatch(text):
            return fail("#VALUE!", "DATE_TEXT_LOCALE")
        if NUMERIC_TEXT.fullmatch(text):
            return fail("#VALUE!", "DATE_TEXT_NUMERIC")
        return fail("#VALUE!", "DATE_TEXT_FORMAT")
    year, month, day = (int(part) for part in match.groups())
    if year < 1900:
        return fail("#NUM!", "DATE_WINDOW")
    if not 1 <= month <= 12 or not 1 <= day <= days_in_month(year, month):
        return fail("#VALUE!", "DATE_TEXT_IMPOSSIBLE")
    parsed = dt.date(year, month, day)
    if not in_window(parsed.toordinal()):
        return fail("#NUM!", "DATE_WINDOW")
    return parsed


def parse_long(value: Val) -> int | Fail:
    """Section 3.2: range precedes integrality; no truncation or rounding."""
    kind = value.kind
    if kind == "err":
        return fail(value.payload, "INPUT_ERROR_PROPAGATED")
    if kind == "empty":
        return fail("#VALUE!", "INPUT_BLANK_REQUIRED")
    if kind not in {"num", "lng"}:
        return fail("#VALUE!", "INTEGER_TYPE_REJECTED")
    number = Decimal(value.payload)
    if number < LONG_MIN or number > LONG_MAX:
        return fail("#NUM!", "INTEGER_RANGE")
    if number != number.to_integral_value():
        return fail("#VALUE!", "INTEGER_FRACTION")
    return int(number)


DOMAINS = {
    "year": (1900, 9999, "DOMAIN_YEAR"),
    "month": (1, 12, "DOMAIN_MONTH"),
    "weekday": (1, 7, "DOMAIN_WEEKDAY"),
    "occurrence": (1, 5, "DOMAIN_OCCURRENCE"),
}
PILLAR_ALIASES = {"ON": 1, "O/N": 1, "TN": 2, "T/N": 2}
PILLAR_BODY = re.compile(r"(?:[0-9]+[YMWD]){1,4}")
PILLAR_COMPONENT = re.compile(r"([0-9]+)([YMWD])")


def parse_pillar(value: Val) -> tuple[int, int] | Fail:
    """Section 3.4: returns the signed (month delta, day delta) of a token."""
    kind = value.kind
    if kind == "err":
        return fail(value.payload, "INPUT_ERROR_PROPAGATED")
    if kind == "empty":
        return fail("#VALUE!", "INPUT_BLANK_REQUIRED")
    if kind != "str":
        return fail("#VALUE!", "PILLAR_TYPE_REJECTED")
    token = value.payload.strip(" ").upper()
    if token in PILLAR_ALIASES:
        return (0, PILLAR_ALIASES[token])
    sign = 1
    body = token
    if token[:1] in {"+", "-"}:
        if token[1:] in PILLAR_ALIASES:
            return fail("#VALUE!", "PILLAR_ALIAS_SIGNED")
        sign = -1 if token[0] == "-" else 1
        body = token[1:]
    if not PILLAR_BODY.fullmatch(body):
        return fail("#VALUE!", "PILLAR_TOKEN_MALFORMED")
    quantities: dict[str, int] = {}
    for amount, unit in PILLAR_COMPONENT.findall(body):
        if unit in quantities:
            return fail("#VALUE!", "PILLAR_DUPLICATE_UNIT")
        quantities[unit] = int(amount)
    months = 12 * quantities.get("Y", 0) + quantities.get("M", 0)
    days = 7 * quantities.get("W", 0) + quantities.get("D", 0)
    return (sign * months, sign * days)


def parse_value(kind: str, value: Val) -> Parsed:
    if kind == "date":
        return parse_date(value)
    if kind == "pillar":
        return parse_pillar(value)
    parsed = parse_long(value)
    if isinstance(parsed, Fail) or kind == "int":
        return parsed
    low, high, condition = DOMAINS[kind]
    if not low <= parsed <= high:
        return fail("#VALUE!", condition)
    return parsed


def parse_control(kind: str, value: Val) -> bool | str | Fail:
    """Section 3.3: omitted and Empty select the default; others are strict."""
    if value.kind == "err":
        return fail(value.payload, "CONTROL_ERROR_PROPAGATED")
    if kind == "rounding":
        if value.kind in {"missing", "empty"}:
            return "NEAREST"
        if value.kind != "str":
            return fail("#VALUE!", "CONTROL_TYPE_REJECTED")
        token = value.payload.strip(" ").upper()
        if token not in {"NEAREST", "FLOOR", "CEILING"}:
            return fail("#VALUE!", "CONTROL_TOKEN_UNKNOWN")
        return token
    if value.kind in {"missing", "empty"}:
        return kind == "bool-true"
    if value.kind != "bool":
        return fail("#VALUE!", "CONTROL_TYPE_REJECTED")
    return value.payload == "TRUE"


def add_months(start: dt.date, months: int, keep_eom: bool) -> dt.date | None:
    """Clip-mode month shift; None when the result leaves the supported window."""
    total = start.year * 12 + (start.month - 1) + months
    year, month = divmod(total, 12)
    month += 1
    if not 1900 <= year <= 9999:
        return None
    last = days_in_month(year, month)
    day = min(start.day, last)
    if keep_eom and start.day == days_in_month(start.year, start.month):
        day = last
    result = dt.date(year, month, day)
    return result if in_window(result.toordinal()) else None


def shift_days(start: dt.date, days: int) -> dt.date | None:
    ordinal = start.toordinal() + days
    return dt.date.fromordinal(ordinal) if in_window(ordinal) else None


def dated(result: dt.date | None, condition: str = "RESULT_WINDOW") -> Val | Fail:
    if result is None:
        return fail("#NUM!", condition)
    return D(result.isoformat())


def gated(result: dt.date) -> Val | Fail:
    return dated(result if in_window(result.toordinal()) else None)


def weekday_number(day: dt.date, monday_base: bool) -> int:
    iso = day.isoweekday()
    return iso if monday_base else iso % 7 + 1


def iso_weekday(index: int, monday_base: bool) -> int:
    return index if monday_base else (7 if index == 1 else index - 1)


def nth_weekday(year: int, month: int, index: int, n: int, monday_base: bool) -> Val | Fail:
    first = dt.date(year, month, 1)
    offset = (iso_weekday(index, monday_base) - first.isoweekday()) % 7
    day = 1 + offset + 7 * (n - 1)
    if day > days_in_month(year, month):
        return fail("#NUM!", "OCCURRENCE_ABSENT")
    return gated(dt.date(year, month, day))


def last_weekday(year: int, month: int, index: int, monday_base: bool) -> Val | Fail:
    last = dt.date(year, month, days_in_month(year, month))
    offset = (last.isoweekday() - iso_weekday(index, monday_base)) % 7
    return gated(last - dt.timedelta(days=offset))


class Anchor(NamedTuple):
    ordinal: int
    unit: str
    quantity: int


def pillar_label(anchor: Anchor, negative: bool) -> str:
    if anchor.unit == "W":
        text = f"{anchor.quantity}W"
    else:
        years, months = divmod(anchor.quantity, 12)
        text = (f"{years}Y" if years else "") + (f"{months}M" if months else "")
    return ("-" if negative else "") + text


def pillar_candidates(start: dt.date, end: dt.date, sign: int) -> list[Anchor]:
    """Section 8.4: 1W-3W anchors plus the floor and ceiling month anchors."""
    candidates = [
        Anchor(start.toordinal() + sign * 7 * weeks, "W", weeks)
        for weeks in (1, 2, 3)
        if in_window(start.toordinal() + sign * 7 * weeks)
    ]
    floor_count = 0
    count = 1
    while True:
        anchor = add_months(start, sign * count, False)
        if anchor is None or (anchor > end if sign > 0 else anchor < end):
            break
        floor_count = count
        count += 1
    for months in (floor_count, floor_count + 1):
        if months < 1:
            continue
        anchor = add_months(start, sign * months, False)
        if anchor is not None:
            candidates.append(Anchor(anchor.toordinal(), "M", months))
    month_dates = {c.ordinal for c in candidates if c.unit == "M"}
    return [c for c in candidates if c.unit == "M" or c.ordinal not in month_dates]


def pillar_from_dates(start: dt.date, end: dt.date, mode: str) -> Val | Fail:
    delta = end.toordinal() - start.toordinal()
    if delta == 0:
        return S("0D")
    if abs(delta) < 7:
        return S(f"{delta}D")
    sign = 1 if delta > 0 else -1
    target = end.toordinal()
    candidates = pillar_candidates(start, end, sign)
    reach = [c for c in candidates if sign * (c.ordinal - target) >= 0]
    short = [c for c in candidates if sign * (c.ordinal - target) <= 0]
    month_first = {"M": 1, "W": 0}
    if mode == "FLOOR":
        chosen = max(short, key=lambda c: (abs(c.ordinal - start.toordinal()), month_first[c.unit]))
    elif mode == "CEILING":
        if not reach:
            return fail("#NUM!", "RESULT_WINDOW")
        chosen = min(reach, key=lambda c: (abs(c.ordinal - start.toordinal()), -month_first[c.unit]))
    else:
        chosen = min(candidates, key=lambda c: (abs(c.ordinal - target), -month_first[c.unit], -c.quantity))
    return S(pillar_label(chosen, sign < 0))


def date_from_pillar(start: dt.date, pillar: tuple[int, int]) -> Val | Fail:
    months, days = pillar
    if abs(months) > 12 * 9999 or abs(days) > MAX_ORD:
        return fail("#NUM!", "PILLAR_AGGREGATE_RANGE")
    shifted = add_months(start, months, False)
    if shifted is None:
        return fail("#NUM!", "PILLAR_AGGREGATE_RANGE")
    return dated(shift_days(shifted, days), "PILLAR_AGGREGATE_RANGE")


# Element semantics: parsed value arguments plus resolved controls -> result.
Compute = Callable[[list[Parsed], list[Union[bool, str]]], Union[Val, Fail]]


def _d(values: list[Parsed], index: int) -> dt.date:
    value = values[index]
    assert isinstance(value, dt.date)
    return value


def _i(values: list[Parsed], index: int) -> int:
    value = values[index]
    assert isinstance(value, int)
    return value


def _quarter_start(day: dt.date) -> dt.date:
    return dt.date(day.year, 3 * ((day.month - 1) // 3) + 1, 1)


def _quarter_end(day: dt.date) -> dt.date:
    month = 3 * ((day.month - 1) // 3) + 3
    return dt.date(day.year, month, days_in_month(day.year, month))


def _eom(day: dt.date) -> dt.date:
    return dt.date(day.year, day.month, days_in_month(day.year, day.month))


def _months(values: list[Parsed], controls: list[bool | str], factor: int) -> Val | Fail:
    months = _i(values, 1) * factor
    if not LONG_MIN <= months <= LONG_MAX:
        return fail("#NUM!", "RESULT_WINDOW")
    return dated(add_months(_d(values, 0), months, bool(controls[0])))


def _pillar_value(values: list[Parsed]) -> tuple[int, int]:
    value = values[1]
    if isinstance(value, Fail) or not isinstance(value, tuple):
        raise FixtureError("Pillar argument was not parsed")
    return value


@dataclass(frozen=True)
class Spec:
    params: tuple[tuple[str, str], ...]
    compute: Compute
    controls: tuple[tuple[str, str], ...] = ()


FUNCTIONS: dict[str, Spec] = {
    "KPR_Dates_DayOfWeek": Spec(
        (("DateIn", "date"),),
        lambda v, c: L(weekday_number(_d(v, 0), bool(c[0]))),
        (("Opt_WeekBaseMonday", "bool-true"),)),
    "KPR_Dates_DaysInMonth": Spec(
        (("DateIn", "date"),), lambda v, c: L(days_in_month(_d(v, 0).year, _d(v, 0).month))),
    "KPR_Dates_DaysInYear": Spec(
        (("YearIn", "year"),), lambda v, c: L(366 if is_leap(_i(v, 0)) else 365)),
    "KPR_Dates_BeginOfMonth": Spec((("DateIn", "date"),), lambda v, c: gated(_d(v, 0).replace(day=1))),
    "KPR_Dates_EndOfMonth": Spec((("DateIn", "date"),), lambda v, c: gated(_eom(_d(v, 0)))),
    "KPR_Dates_BeginOfQuarter": Spec((("DateIn", "date"),), lambda v, c: gated(_quarter_start(_d(v, 0)))),
    "KPR_Dates_EndOfQuarter": Spec((("DateIn", "date"),), lambda v, c: gated(_quarter_end(_d(v, 0)))),
    "KPR_Dates_BeginOfYear": Spec((("DateIn", "date"),), lambda v, c: gated(dt.date(_d(v, 0).year, 1, 1))),
    "KPR_Dates_EndOfYear": Spec((("DateIn", "date"),), lambda v, c: gated(dt.date(_d(v, 0).year, 12, 31))),
    "KPR_Dates_IsMonthEnd": Spec((("DateIn", "date"),), lambda v, c: B(_d(v, 0) == _eom(_d(v, 0)))),
    "KPR_Dates_IsQuarterEnd": Spec((("DateIn", "date"),), lambda v, c: B(_d(v, 0) == _quarter_end(_d(v, 0)))),
    "KPR_Dates_IsYearEnd": Spec(
        (("DateIn", "date"),), lambda v, c: B((_d(v, 0).month, _d(v, 0).day) == (12, 31))),
    "KPR_Dates_IsLeapYear": Spec((("YearIn", "year"),), lambda v, c: B(is_leap(_i(v, 0)))),
    "KPR_Dates_AddDays": Spec(
        (("DateIn", "date"), ("nDays", "int")), lambda v, c: dated(shift_days(_d(v, 0), _i(v, 1)))),
    "KPR_Dates_AddWeeks": Spec(
        (("DateIn", "date"), ("nWeeks", "int")),
        lambda v, c: dated(shift_days(_d(v, 0), 7 * _i(v, 1)))),
    "KPR_Dates_AddMonths": Spec(
        (("DateIn", "date"), ("nMonths", "int")), lambda v, c: _months(v, c, 1),
        (("Opt_KeepEOM", "bool-false"),)),
    "KPR_Dates_AddYears": Spec(
        (("DateIn", "date"), ("nYears", "int")), lambda v, c: _months(v, c, 12),
        (("Opt_KeepEOM", "bool-false"),)),
    "KPR_Dates_NthWeekdayOfMonth": Spec(
        (("YearIn", "year"), ("MonthIn", "month"), ("WdIndex", "weekday"), ("n", "occurrence")),
        lambda v, c: nth_weekday(_i(v, 0), _i(v, 1), _i(v, 2), _i(v, 3), bool(c[0])),
        (("Opt_WeekBaseMonday", "bool-true"),)),
    "KPR_Dates_LastWeekdayOfMonth": Spec(
        (("YearIn", "year"), ("MonthIn", "month"), ("WdIndex", "weekday")),
        lambda v, c: last_weekday(_i(v, 0), _i(v, 1), _i(v, 2), bool(c[0])),
        (("Opt_WeekBaseMonday", "bool-true"),)),
    "KPR_Dates_PillarFromDates": Spec(
        (("StartDate", "date"), ("EndDate", "date")),
        lambda v, c: pillar_from_dates(_d(v, 0), _d(v, 1), str(c[0])),
        (("Opt_Rounding", "rounding"),)),
    "KPR_Dates_DateFromPillar": Spec(
        (("StartDate", "date"), ("Pillar", "pillar")),
        lambda v, c: date_from_pillar(_d(v, 0), _pillar_value(v))),
}
HOST_FUNCTION = "KPR_Dates_HostDateSystem"
ALL_FUNCTIONS = tuple(FUNCTIONS) + (HOST_FUNCTION,)


@dataclass(frozen=True)
class Outcome:
    value: Val
    condition: str
    level: str


def call_level(error: str, condition: str) -> Outcome:
    return Outcome(E(error), condition, "call")


def shape_of(value: Val) -> tuple[int, int] | None:
    """Section 5.1: None for a scalar, otherwise (rows, columns)."""
    if value.kind == "arr1":
        return None if len(value.items) == 1 else (1, len(value.items))
    if value.kind in {"arr2", "fill2"}:
        return None if (value.rows, value.cols) == (1, 1) else (value.rows, value.cols)
    return None


def unwrap(value: Val) -> Val:
    if value.kind in ARRAY_KINDS and shape_of(value) is None:
        return value.items[0]
    return value


def element_at(value: Val, row: int, col: int) -> Val:
    if value.kind == "arr2":
        return value.items[row * value.cols + col]
    if value.kind == "arr1":
        return value.items[col]
    if value.kind == "fill2":
        return value.items[0]
    return value


def evaluate_element(spec: Spec, values: list[Val], controls: list[bool | str]) -> Val | Fail:
    parsed: list[Parsed] = []
    for (_, kind), value in zip(spec.params, values):
        item = parse_value(kind, value)
        if isinstance(item, Fail):
            return item
        parsed.append(item)
    return spec.compute(parsed, controls)


def as_outcome(result: Val | Fail) -> Outcome:
    if isinstance(result, Fail):
        return Outcome(E(result.error), result.condition, "element")
    return Outcome(result, "", "")


def evaluate(function: str, args: list[Val], context: str) -> Outcome:
    """Apply the section 5.2 public-boundary stages, then per-element semantics."""
    if function == HOST_FUNCTION:
        return Outcome(L(1904 if context == "ws1904" else 1900), "", "")
    spec = FUNCTIONS[function]
    if context == "ws1904":
        return call_level("#N/A", "HOST_DATE1904")
    width = len(spec.params)
    values = args[:width]
    controls_in = args[width:] + [MISSING] * (len(spec.controls) - len(args[width:]))
    if any(v.kind == "arr1" and not v.items for v in values):
        return call_level("#VALUE!", "SHAPE_UNSUPPORTED")
    for control in controls_in:
        if control.kind in ARRAY_KINDS and shape_of(control) is not None:
            return call_level("#VALUE!", "CONTROL_NOT_SCALAR")
    controls: list[bool | str] = []
    for (_, kind), control in zip(spec.controls, controls_in):
        resolved = parse_control(kind, unwrap(control))
        if isinstance(resolved, Fail):
            return call_level(resolved.error, resolved.condition)
        controls.append(resolved)
    shapes = {shape for shape in map(shape_of, values) if shape is not None}
    if len(shapes) > 1:
        return call_level("#VALUE!", "SHAPE_MISMATCH")
    if not shapes:
        return as_outcome(evaluate_element(spec, [unwrap(v) for v in values], controls))
    rows, cols = shapes.pop()
    if rows * cols > CAPACITY:
        return call_level("#NUM!", "CAPACITY_EXCEEDED")
    return evaluate_array(spec, values, controls, rows, cols)


def evaluate_array(spec: Spec, values: list[Val], controls: list[bool | str],
                   rows: int, cols: int) -> Outcome:
    if all(v.kind == "fill2" or shape_of(v) is None for v in values):
        single = as_outcome(evaluate_element(spec, [element_at(unwrap(v), 0, 0) for v in values], controls))
        condition = f"all={single.condition}" if single.condition else ""
        return Outcome(F2(rows, cols, single.value), condition, single.level)
    cells: list[Val] = []
    conditions: list[str] = []
    for row in range(rows):
        for col in range(cols):
            items = [element_at(unwrap(v), row, col) for v in values]
            if any(item.kind in ARRAY_KINDS for item in items):
                return call_level("#VALUE!", "SHAPE_UNSUPPORTED")
            outcome = as_outcome(evaluate_element(spec, items, controls))
            cells.append(outcome.value)
            if outcome.condition:
                conditions.append(f"r{row + 1}c{col + 1}={outcome.condition}")
    return Outcome(A2(rows, cols, cells), ";".join(conditions), "element" if conditions else "")


# ---------------------------------------------------------------------------
# Fixture authoring
# ---------------------------------------------------------------------------


@dataclass
class Case:
    id: str
    suite: str
    function: str
    context: str
    args: list[Val]
    rationale: str
    outcome: Outcome = field(default_factory=lambda: Outcome(EMPTY, "", ""))


class Book:
    def __init__(self) -> None:
        self.cases: list[Case] = []
        self.ids: set[str] = set()

    def add(self, suite: str, slug: str, function: str, args: list[Val], rationale: str,
            context: str = "direct") -> Case:
        case_id = f"{suite}.{slug}"
        if case_id in self.ids:
            raise FixtureError(f"duplicate fixture id: {case_id}")
        if function not in ALL_FUNCTIONS or context not in CONTEXTS or len(args) > 5:
            raise FixtureError(f"invalid fixture definition: {case_id}")
        self.ids.add(case_id)
        case = Case(case_id, suite, function, context, list(args), rationale)
        case.outcome = evaluate(function, case.args, context)
        self.cases.append(case)
        return case


def slug(*parts: object) -> str:
    words = [str(part) for part in parts]
    words = ["minus" + w[1:] if re.fullmatch(r"-[0-9]+(?:\.[0-9]+)?", w) else w for w in words]
    text = "-".join(words).lower()
    text = re.sub(r"[^a-z0-9]+", "-", text).strip("-")
    return text or "blank"


def text_slug(text: str) -> str:
    """A collision-free slug for fixture text: punctuation and spaces are named."""
    names = {" ": "sp", "/": "sl", ".": "dot", ":": "col", "+": "plus", "-": "-"}
    return slug(*(names.get(ch, ch) for ch in text)) if text else "empty"


def token_slug(token: str) -> str:
    """A collision-free slug for a pillar token, where signs are significant."""
    names = {" ": "sp", "/": "sl", ".": "dot", "+": "plus", "-": "minus"}
    words = (names.get(ch, "lc" + ch if ch.islower() else ch) for ch in token)
    return slug(*words) if token else "empty"


F = "KPR_Dates_"
ERR_ORDER = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A")


def author_date_input(book: Book) -> None:
    suite = "date-input"
    fn = F + "EndOfMonth"
    accepted = [
        ("vba-date", D("2024-02-10"), "native VBA Date"),
        ("vba-datetime", DT("2024-02-10T23:59:59"), "native Date with a time component; time removed"),
        ("serial", N("45332"), "native numeric serial under the direct 1900 contract"),
        ("serial-fraction", N("45332.75"), "numeric serial with time fraction; fraction removed"),
        ("serial-long", L(45332), "Long serial"),
        ("iso-text", S("2024-02-10"), "exact YYYY-MM-DD text"),
        ("min-date", D("1900-03-01"), "window lower bound as a VBA Date"),
        ("min-text", S("1900-03-01"), "window lower bound as ISO text"),
        ("min-serial", N("61"), "serial 61 is 1900-03-01, the first supported serial"),
        ("min-serial-fraction", N("61.5"), "serial 61 with time"),
        ("max-date", D("9999-12-31"), "window upper bound as a VBA Date"),
        ("max-text", S("9999-12-31"), "window upper bound as ISO text"),
        ("max-serial", N("2958465"), "serial 2958465 is 9999-12-31"),
        ("max-serial-fraction", N("2958465.9"), "last supported serial with time"),
    ]
    for name, value, why in accepted:
        book.add(suite, f"accept-{name}", fn, [value], f"Section 3.1: {why}.")
    rejected = [
        ("empty", EMPTY, "required blank or Empty"),
        ("null", NULL, "Null is a prohibited type"),
        ("bool-true", B(True), "Boolean is a prohibited type"),
        ("bool-false", B(False), "Boolean is a prohibited type"),
        ("serial-60", N("60"), "serial 60 is the fictitious 1900-02-29"),
        ("serial-60-fraction", N("60.5"), "serial 60 with time is still below the window"),
        ("serial-59", N("59"), "serial below the window"),
        ("serial-0", N("0"), "serial below the window"),
        ("serial-negative", N("-1"), "negative serial"),
        ("serial-above-max", N("2958466"), "serial above 9999-12-31"),
        ("serial-huge", N("1E+20"), "numeric far above the window"),
        ("vba-date-1900-02-28", D("1900-02-28"), "VBA Date below the window"),
        ("vba-date-1899-12-31", D("1899-12-31"), "VBA Date below the window"),
        ("text-1900-02-28", S("1900-02-28"), "exact ISO date below the window"),
        ("text-1899-12-31", S("1899-12-31"), "year below 1900 is DATE_WINDOW without component checks"),
        ("text-0001-01-01", S("0001-01-01"), "year below 1900"),
        ("text-1800-02-30", S("1800-02-30"), "year below 1900 is DATE_WINDOW before component checks"),
    ]
    for name, value, why in rejected:
        book.add(suite, f"reject-{name}", fn, [value], f"Section 3.1: {why}.")
    texts = [
        ("impossible", ["2025-02-29", "2023-04-31", "2024-13-01", "2024-00-10", "2024-01-00",
                        "2024-02-30", "1900-02-29"], "names a date that does not exist"),
        ("format", ["", " 2024-02-10", "2024-02-10 ", "2024-2-10", "2024-02-1", "2024/02/10",
                    "2024.02.10", "20240210", "2024-02-10T00:00:00", "2024-02-10 00:00", "24-02-10",
                    "abc", "2024-02-10Z"], "is not exactly YYYY-MM-DD"),
        ("locale", ["31/12/2026", "12/31/2026", "10/02/2024", "10.02.2024"], "is locale-formatted"),
        ("numeric", ["61", "45292", "45292.5", "2958465"], "is numeric-looking text, never a serial"),
    ]
    for group, samples, why in texts:
        for sample in samples:
            book.add(suite, f"text-{group}-{text_slug(sample)}", fn, [S(sample)],
                     f"Section 3.1: {sample!r} {why}.")


def author_integer_input(book: Book) -> None:
    suite = "integer-input"
    fn = F + "AddDays"
    start = D("2024-01-31")
    samples = [
        ("one", N("1"), "integral numeric"),
        ("one-point-zero", N("1.0"), "integral numeric written with a fraction digit"),
        ("minus-one", N("-1"), "negative shift"),
        ("zero", N("0"), "zero shift"),
        ("long-one", L(1), "Long subtype"),
        ("fraction-half", N("0.5"), "fractional value is never truncated"),
        ("fraction-minus-half", N("-0.5"), "negative fraction"),
        ("fraction-one-half", N("1.5"), "fraction is never rounded"),
        ("fraction-near-max", N("2147483646.5"), "fraction inside the Long range"),
        ("range-above", N("2147483648"), "one above the Long range"),
        ("range-below", N("-2147483649"), "one below the Long range"),
        ("range-fraction-above", N("2147483648.5"), "range precedes integrality"),
        ("range-fraction-just-above", N("2147483647.5"), "above Long max, so range precedes integrality"),
        ("range-huge", N("1E+20"), "far outside the Long range"),
        ("long-max", N("2147483647"), "Long max parses; the shifted date leaves the window"),
        ("long-min", N("-2147483648"), "Long min parses; the shifted date leaves the window"),
        ("type-bool", B(True), "Boolean is rejected"),
        ("type-date", D("2024-01-01"), "a Date is rejected at an integer position"),
        ("type-text", S("1"), "numeric-looking text is rejected"),
        ("type-empty-text", S(""), "empty text is rejected"),
        ("type-null", NULL, "Null is rejected"),
        ("blank", EMPTY, "blank required value"),
    ]
    for name, value, why in samples:
        book.add(suite, name, fn, [start, value], f"Section 3.2: {why}.")


def author_controls(book: Book) -> None:
    suite = "control"
    sunday = D("2024-02-11")
    base = [
        ("omitted", None), ("empty", EMPTY), ("true", B(True)), ("false", B(False)),
        ("num-0", N("0")), ("num-1", N("1")), ("text-true", S("TRUE")), ("null", NULL),
        ("date", D("2024-01-01")), ("error", E("#DIV/0!")), ("array-1x1", A2(1, 1, [B(False)])),
        ("array-1x2", A2(1, 2, [B(True), B(False)])), ("array-1d-single", A1([B(True)])),
    ]
    for name, value in base:
        args = [sunday] if value is None else [sunday, value]
        book.add(suite, f"weekbase-{name}", F + "DayOfWeek", args,
                 "Section 3.3: Opt_WeekBaseMonday accepts omitted, Empty or native Boolean only.")
    eom = D("2023-02-28")
    keep = [
        ("omitted", None), ("empty", EMPTY), ("true", B(True)), ("false", B(False)),
        ("num-1", N("1")), ("text", S("TRUE")), ("error", E("#NAME?")),
        ("array-2x1", A2(2, 1, [B(True), B(True)])),
    ]
    for name, value in keep:
        args = [eom, N("1")] if value is None else [eom, N("1"), value]
        book.add(suite, f"keepeom-{name}", F + "AddMonths", args,
                 "Section 3.3: Opt_KeepEOM accepts omitted, Empty or native Boolean only.")
    start, end = D("2024-01-15"), D("2024-02-10")
    rounding = [
        ("omitted", None), ("empty", EMPTY), ("nearest", S("NEAREST")), ("lower", S("nearest")),
        ("floor-padded", S("  FLOOR  ")), ("ceiling-mixed", S("Ceiling")), ("unknown", S("ROUND")),
        ("unknown-two-words", S("FLOOR NEAREST")), ("num", N("1")), ("bool", B(True)), ("null", NULL),
        ("error", E("#REF!")), ("array-1x1", A2(1, 1, [S("floor")])),
        ("array-1x2", A2(1, 2, [S("FLOOR"), S("CEILING")])),
    ]
    for name, value in rounding:
        args = [start, end] if value is None else [start, end, value]
        book.add(suite, f"rounding-{name}", F + "PillarFromDates", args,
                 "Section 3.3: Opt_Rounding is trimmed text NEAREST, FLOOR or CEILING.")
    book.add(suite, "precedence-control-before-element", F + "AddMonths",
             [S("bad"), N("1"), N("1")],
             "Section 5.2: control validation is call-level and precedes element evaluation.")


BOUNDARY_DATES = (
    "1900-03-01", "1900-03-31", "1900-04-15", "1900-12-31", "1901-01-01", "1999-12-31",
    "2000-02-29", "2000-03-31", "2023-02-28", "2024-02-28", "2024-02-29", "2024-03-31",
    "2024-06-30", "2024-07-01", "2024-09-30", "2024-11-30", "2024-12-31", "2100-02-28",
    "9999-01-01", "9999-12-31",
)
BOUNDARY_FUNCTIONS = (
    "DaysInMonth", "BeginOfMonth", "EndOfMonth", "BeginOfQuarter", "EndOfQuarter",
    "BeginOfYear", "EndOfYear", "IsMonthEnd", "IsQuarterEnd", "IsYearEnd",
)


def author_boundaries(book: Book) -> None:
    for day in BOUNDARY_DATES:
        for name in BOUNDARY_FUNCTIONS:
            book.add("boundary", slug(name, day), F + name, [D(day)],
                     "Section 8.1: Gregorian boundaries; a boundary is a result and is window-gated.")
        book.add("weekday", slug("monday-base", day), F + "DayOfWeek", [D(day)],
                 "Section 8.1: Monday base, Monday is 1 and Sunday is 7.")
        book.add("weekday", slug("sunday-base", day), F + "DayOfWeek", [D(day), B(False)],
                 "Section 8.1: Sunday base, Sunday is 1 and Saturday is 7.")


def author_years(book: Book) -> None:
    years = [N(str(y)) for y in (1900, 1901, 1904, 2000, 2023, 2024, 2100, 2400, 9996, 9999)]
    extra = [
        ("long-2024", L(2024)), ("domain-1899", N("1899")), ("domain-10000", N("10000")),
        ("domain-0", N("0")), ("domain-negative", N("-1")), ("text-2024", S("2024")),
        ("iso-text", S("2024-01-01")), ("vba-date", D("2024-01-01")), ("bool", B(True)),
        ("blank", EMPTY), ("null", NULL), ("fraction", N("2024.5")), ("range", N("1E+20")),
        ("error", E("#N/A")),
    ]
    for name in ("DaysInYear", "IsLeapYear"):
        for value in years:
            book.add("year", slug(name, value.payload), F + name, [value],
                     "Section 8.1: Gregorian leap rule on a calendar year; no window gate.")
        for label, value in extra:
            book.add("year", slug(name, label), F + name, [value],
                     "Sections 3.2 and 8.1: strict integer year, domain 1900-9999.")


def author_arithmetic(book: Book) -> None:
    rows = [
        ("AddDays", "2024-02-28", "1", None), ("AddDays", "2023-02-28", "1", None),
        ("AddDays", "2024-01-01", "-1", None), ("AddDays", "1900-03-02", "-1", None),
        ("AddDays", "1900-03-01", "-1", None), ("AddDays", "9999-12-30", "1", None),
        ("AddDays", "9999-12-31", "1", None), ("AddDays", "2024-01-01", "366", None),
        ("AddDays", "1900-03-01", "2958404", None), ("AddDays", "1900-03-01", "2958405", None),
        ("AddWeeks", "2024-02-22", "1", None), ("AddWeeks", "1900-03-08", "-1", None),
        ("AddWeeks", "1900-03-07", "-1", None), ("AddWeeks", "9999-12-24", "1", None),
        ("AddWeeks", "9999-12-25", "1", None), ("AddWeeks", "2024-01-01", "2147483647", None),
        ("AddWeeks", "2024-01-01", "-2147483648", None),
        ("AddMonths", "2024-01-31", "1", None), ("AddMonths", "2023-01-31", "1", None),
        ("AddMonths", "2024-03-31", "-1", None), ("AddMonths", "2024-05-31", "-1", None),
        ("AddMonths", "2024-01-15", "1", None), ("AddMonths", "2024-12-15", "1", None),
        ("AddMonths", "2024-01-15", "-1", None), ("AddMonths", "2024-02-29", "12", None),
        ("AddMonths", "2024-01-15", "0", None), ("AddMonths", "2023-02-28", "1", True),
        ("AddMonths", "2024-04-30", "1", True), ("AddMonths", "2024-02-29", "12", True),
        ("AddMonths", "2023-02-28", "12", True), ("AddMonths", "2024-01-30", "1", True),
        ("AddMonths", "2024-01-31", "1", True), ("AddMonths", "2024-02-29", "1", True),
        ("AddMonths", "2024-02-29", "1", False), ("AddMonths", "1900-03-31", "-1", None),
        ("AddMonths", "1900-04-30", "-1", True), ("AddMonths", "9999-12-01", "1", None),
        ("AddMonths", "9999-11-30", "1", True), ("AddMonths", "2024-01-15", "2147483647", None),
        ("AddMonths", "2024-01-15", "-2147483648", None),
        ("AddYears", "2024-02-29", "1", None), ("AddYears", "2024-02-29", "4", None),
        ("AddYears", "2023-02-28", "1", True), ("AddYears", "2023-02-28", "1", False),
        ("AddYears", "2024-02-29", "-1", None), ("AddYears", "1900-03-01", "-1", None),
        ("AddYears", "9999-01-01", "1", None), ("AddYears", "2024-06-30", "-124", None),
        ("AddYears", "2024-06-30", "7975", None), ("AddYears", "2024-01-15", "200000000", None),
        ("AddYears", "2024-01-15", "2147483647", None),
    ]
    for name, start, amount, keep in rows:
        args = [D(start), N(amount)] + ([] if keep is None else [B(keep)])
        book.add("arithmetic", slug(name, start, amount, "" if keep is None else f"eom-{keep}"),
                 F + name, args,
                 "Section 8.2: exact day/week shifts and clip-mode month shifts; every result is "
                 "window-gated and overflow returns #NUM!.")


def author_locators(book: Book) -> None:
    nth = [
        (2024, 9, 1, 1, None), (2024, 1, 3, 3, None), (2024, 3, 5, 5, None), (2024, 2, 1, 5, None),
        (2024, 2, 4, 5, None), (2024, 9, 1, 1, False), (2024, 9, 7, 1, False), (2024, 9, 7, 1, True),
        (2024, 12, 2, 4, None), (1900, 1, 1, 1, None), (1900, 2, 1, 2, None), (1900, 3, 4, 1, None),
        (1900, 3, 5, 5, None), (9999, 12, 5, 5, None), (9999, 12, 5, 1, False), (2023, 2, 3, 4, None),
    ]
    for year, month, index, n, base in nth:
        args = [N(str(year)), N(str(month)), N(str(index)), N(str(n))] + ([] if base is None else [B(base)])
        book.add("locator", slug("nth", year, month, index, n, "" if base is None else base),
                 F + "NthWeekdayOfMonth", args,
                 "Section 8.3: occurrence n of WdIndex; absent occurrence is OCCURRENCE_ABSENT and a "
                 "result before 1900-03-01 is RESULT_WINDOW.")
    last = [
        (2024, 2, 1, None), (2024, 5, 5, None), (2024, 3, 1, False), (2024, 8, 7, False),
        (1900, 2, 3, None), (1900, 3, 6, None), (9999, 12, 5, None), (9999, 12, 7, None),
    ]
    for year, month, index, base in last:
        args = [N(str(year)), N(str(month)), N(str(index))] + ([] if base is None else [B(base)])
        book.add("locator", slug("last", year, month, index, "" if base is None else base),
                 F + "LastWeekdayOfMonth", args, "Section 8.3: final occurrence of WdIndex.")
    domains = [
        ("year-1899", ["1899", "3", "1", "1"]), ("year-10000", ["10000", "3", "1", "1"]),
        ("month-0", ["2024", "0", "1", "1"]), ("month-13", ["2024", "13", "1", "1"]),
        ("weekday-0", ["2024", "3", "0", "1"]), ("weekday-8", ["2024", "3", "8", "1"]),
        ("occurrence-0", ["2024", "3", "1", "0"]), ("occurrence-6", ["2024", "3", "1", "6"]),
        ("occurrence-fraction", ["2024", "3", "1", "1.5"]),
    ]
    for name, raw in domains:
        book.add("locator", f"nth-domain-{name}", F + "NthWeekdayOfMonth", [N(v) for v in raw],
                 "Section 3.2: locator argument domains return #VALUE!.")
    book.add("locator", "nth-month-text", F + "NthWeekdayOfMonth",
             [N("2024"), S("3"), N("1"), N("1")], "Section 3.2: text is rejected at MonthIn.")
    book.add("locator", "last-domain-weekday-8", F + "LastWeekdayOfMonth",
             [N("2024"), N("3"), N("8")], "Section 3.2: WdIndex domain is 1 through 7.")
    book.add("locator", "nth-first-failing-year", F + "NthWeekdayOfMonth",
             [N("1899"), N("13"), N("1"), N("1")],
             "Section 5.2: the first failing argument in signature order determines the result.")


PILLAR_STARTS = {"s15": "2024-01-15", "s31": "2024-01-31", "s31n": "2023-01-31"}


def author_pillar_parse(book: Book) -> None:
    suite = "pillar-parse"
    fn = F + "DateFromPillar"
    valid = [
        "ON", "O/N", "TN", "T/N", "on", "o/n", "t/n", "1D", "1d", "7D", "1W", "2W", "1M", "1m", "3M",
        "12M", "1Y", "1Y6M", "6M1Y", "2W3D", "3D2W", "1Y2M3W4D", "4D3W2M1Y", "+1M", "-1M", "-1Y6M",
        "-2W3D", "0D", "0M", "0Y0M0W0D", "00001M", "  1M  ", "1M ", " -1M",
    ]
    for token in valid:
        book.add(suite, "s15-" + token_slug(token), fn, [D(PILLAR_STARTS["s15"]), S(token)],
                 "Section 3.4: accepted grammar; months apply first with clip semantics, then days.")
    clip = [("s31", "1M"), ("s31n", "1M"), ("s31", "1M1D"), ("s31", "1D1M"), ("s31", "1Y1M"), ("s31", "-2M")]
    for start, token in clip:
        book.add(suite, start + "-" + token_slug(token), fn, [D(PILLAR_STARTS[start]), S(token)],
                 "Section 3.4: month delta first with clip, then the exact day delta.")
    rejected = {
        "signed-alias": ["-ON", "+ON", "-O/N", "+TN", "-T/N", "+t/n"],
        "duplicate": ["1M2M", "3D4D", "1Y1Y", "1W2W", "1M1D1M"],
        "malformed": ["1 M", "1M 2W", "M", "1", "1.5M", "1X", "-", "+", "", "1M+2W", "ON1D", "1MM",
                      "+-1M", "--1M", "1M-", "   "],
        "aggregate": ["10000Y", "99999999999999999999M", "3000000000D", "2147483648M"],
    }
    for group, tokens in rejected.items():
        for token in tokens:
            book.add(suite, group + "-" + token_slug(token), fn, [D(PILLAR_STARTS["s15"]), S(token)],
                     f"Section 3.4: {group} token is rejected with its registry condition.")
    edges = [
        ("9999-12-15", "1M"), ("9999-12-31", "1D"), ("9999-12-31", "ON"), ("1900-03-01", "-1D"),
        ("1900-03-31", "-1M"), ("1900-03-01", "ON"), ("1900-03-01", "0D"),
    ]
    for start, token in edges:
        book.add(suite, slug("edge", start) + "-" + token_slug(token), fn, [D(start), S(token)],
                 "Section 3.4: a valid token whose result leaves the window is PILLAR_AGGREGATE_RANGE.")
    types = [("num", N("12")), ("num-1", N("1")), ("bool", B(True)), ("date", D("2024-01-01")),
             ("null", NULL), ("blank", EMPTY), ("error", E("#N/A"))]
    for name, value in types:
        book.add(suite, f"type-{name}", fn, [D(PILLAR_STARTS["s15"]), value],
                 "Sections 3.4 and 7: non-text payloads; a blank required value is INPUT_BLANK_REQUIRED.")
    book.add(suite, "start-before-pillar", fn, [S("bad"), S("1X")],
             "Section 5.2: StartDate is evaluated before Pillar.")


TRANSITION_STARTS = ("2024-01-15", "2024-01-31", "2023-01-31", "2024-04-15", "2023-02-01")
MODES = ("NEAREST", "FLOOR", "CEILING")


def author_pillar_format(book: Book) -> None:
    suite = "pillar-format"
    fn = F + "PillarFromDates"
    base = dt.date(2024, 1, 15)
    for days in [0, *range(1, 7), *range(-6, 0), 7, 14, 21, -7, -14, -21]:
        end = base + dt.timedelta(days=days)
        book.add(suite, slug("short", days), fn, [D(base.isoformat()), D(end.isoformat())],
                 "Section 8.4: 0D, exact signed days below seven, exact 1W-3W anchors.")
    for start in TRANSITION_STARTS:
        origin = dt.date.fromisoformat(start)
        for days in range(22, 33):
            end = origin + dt.timedelta(days=days)
            for mode in MODES:
                book.add(suite, slug("transition", start, days, mode), fn,
                         [D(start), D(end.isoformat()), S(mode)],
                         "Section 8.4: 3W/1M transition derived from calendar-day distance; ties take "
                         "the month.")
    origin = dt.date(2024, 3, 15)
    for days in range(22, 33):
        end = origin - dt.timedelta(days=days)
        for mode in MODES:
            book.add(suite, slug("backward", days, mode), fn,
                     [D(origin.isoformat()), D(end.isoformat()), S(mode)],
                     "Section 8.4: negative intervals round the magnitude and restore the sign.")
    pairs = [
        ("2024-01-15", "2024-03-01"), ("2024-01-15", "2024-02-29"), ("2023-01-15", "2023-03-01"),
        ("2024-01-15", "2025-01-15"), ("2024-01-15", "2025-07-15"), ("2024-01-15", "2034-01-29"),
        ("2024-01-15", "2025-01-14"), ("2024-01-15", "2026-01-15"), ("2025-07-15", "2024-01-15"),
        ("2024-01-31", "2024-02-29"), ("2024-01-31", "2024-03-31"), ("2024-03-31", "2024-02-29"),
        ("9999-12-20", "9999-12-31"), ("1900-03-12", "1900-03-01"), ("2024-01-15", "2024-02-15"),
    ]
    for first, second in pairs:
        for mode in MODES:
            book.add(suite, slug("interval", first, second, mode), fn, [D(first), D(second), S(mode)],
                     "Section 8.4: month-family labels, year formatting, exact clip anchors, ties to "
                     "the larger month count and window-excluded anchors.")
    book.add(suite, "error-start-text", fn, [S("bad"), D("2024-01-15")],
             "Section 5.2: StartDate parse failure.")
    book.add(suite, "error-end-blank", fn, [D("2024-01-15"), EMPTY], "Section 3.1: blank EndDate.")


ROUNDTRIP_PAIRS = (
    ("2024-01-15", "2024-02-10"), ("2024-01-15", "2024-02-08"), ("2024-01-15", "2034-01-29"),
    ("2024-01-15", "2024-02-15"), ("2024-01-31", "2024-02-29"), ("2024-01-15", "2025-07-15"),
)


def author_pillar_roundtrip(book: Book) -> None:
    for start, end in ROUNDTRIP_PAIRS:
        emitted = evaluate(F + "PillarFromDates", [D(start), D(end)], "direct").value
        if emitted.kind != "str":
            raise FixtureError("round-trip pair must format successfully")
        back = evaluate(F + "DateFromPillar", [D(start), emitted], "direct").value
        relation = "equals" if back.payload == end else "does not equal"
        book.add("pillar-roundtrip", slug(start, end, emitted.payload), F + "DateFromPillar",
                 [D(start), emitted],
                 f"Section 8.4: re-parsing the emitted NEAREST token {emitted.payload!r} {relation} "
                 f"EndDate {end}; rounded intervals are deliberately not invariant.")


def author_shapes(book: Book) -> None:
    suite = "shape"
    feb = D("2024-02-10")
    eom = F + "EndOfMonth"
    add = F + "AddDays"
    rows = [
        ("scalar-1x1", eom, [A2(1, 1, [feb])], "a 1x1 array is scalar and returns a scalar"),
        ("scalar-1d-single", eom, [A1([feb])], "a one-element 1-D array is scalar"),
        ("row", eom, [A2(1, 3, [D("2024-01-10"), S("2024-02-10"), N("45332")])], "row shape preserved"),
        ("column", eom, [A2(3, 1, [D("2024-01-10"), D("2024-02-10"), D("2023-02-10")])],
         "column shape preserved"),
        ("rectangle-mixed", eom,
         [A2(2, 3, [D("2024-01-10"), S("bad"), E("#N/A"), EMPTY, S("2024-02-30"), N("60")])],
         "element-level errors stay at their positions"),
        ("one-dimensional", eom, [A1([D("2024-01-10"), D("2024-02-10"), D("2024-03-10")])],
         "a 1-D array is a 1xN row"),
        ("broadcast-first", add, [A2(2, 2, [D("2024-01-31"), D("2024-02-28"), D("2023-02-28"),
                                            D("2024-12-31")]), N("1")], "scalar expands to 2x2"),
        ("broadcast-second", add, [D("2024-01-31"), A2(1, 3, [N("1"), N("2"), N("3")])],
         "scalar expands to 1x3"),
        ("elementwise", add, [A2(2, 2, [D("2024-01-31")] * 4), A2(2, 2, [N("1"), N("-1"), N("0"), N("29")])],
         "identical shapes evaluate element for element"),
        ("oned-with-row", add, [A1([D("2024-01-31")] * 3), A2(1, 3, [N("1"), N("2"), N("3")])],
         "a 1-D array and a 1x3 array share the 1x3 shape"),
        ("mismatch-row-column", add, [A2(1, 3, [D("2024-01-31")] * 3), A2(3, 1, [N("1")] * 3)],
         "no row/column outer product"),
        ("mismatch-width", add, [A2(1, 3, [D("2024-01-31")] * 3), A2(1, 2, [N("1")] * 2)],
         "non-scalar shapes must be identical"),
        ("error-order", add, [A2(1, 2, [E("#REF!"), D("2024-01-31")]), A2(1, 2, [E("#DIV/0!"), E("#N/A")])],
         "the first error in signature order propagates at each position"),
        ("failure-order", add, [A2(1, 2, [S("bad"), D("2024-01-31")]), A2(1, 2, [E("#N/A"), S("x")])],
         "the first failing argument determines each element"),
        ("control-not-scalar", F + "AddMonths",
         [A2(1, 2, [D("2024-01-31")] * 2), N("1"), A2(1, 2, [B(True), B(False)])],
         "a multi-element control is CONTROL_NOT_SCALAR"),
        ("control-1x1", F + "AddMonths", [A2(1, 2, [D("2023-02-28"), D("2024-01-15")]), N("1"),
                                          A2(1, 1, [B(True)])], "a 1x1 control is scalar"),
        ("control-1d-pair", F + "DayOfWeek", [A2(1, 2, [feb, feb]), A1([B(False), B(True)])],
         "a two-element 1-D control is not scalar"),
        ("empty-array", eom, [A1([])], "an empty array is SHAPE_UNSUPPORTED"),
        ("jagged", eom, [A1([A1([D("2024-01-10"), D("2024-02-10")]), feb])],
         "a jagged array is SHAPE_UNSUPPORTED, found during materialization"),
        ("scalar-with-1x1", add, [feb, A2(1, 1, [N("5")])], "all-scalar call returns a scalar"),
        ("four-argument-broadcast", F + "NthWeekdayOfMonth",
         [N("2024"), A2(1, 3, [N("1"), N("2"), N("3")]), N("1"), N("1")], "one non-scalar argument"),
        ("pillar-vector", F + "PillarFromDates",
         [A2(1, 3, [D("2024-01-15"), D("2024-01-20"), D("2023-12-15")]), D("2024-02-10"), S("FLOOR")],
         "vectorized PillarFromDates"),
        ("pillar-token-vector", F + "DateFromPillar",
         [D("2024-01-15"), A2(1, 3, [S("1M"), S("ON"), S("1X")])], "vectorized DateFromPillar"),
        ("year-rectangle", F + "IsLeapYear", [A2(2, 2, [N("1900"), N("2000"), N("2024"), N("2023")])],
         "vectorized year predicate"),
    ]
    for name, fn, args, why in rows:
        book.add(suite, name, fn, args, f"Section 5: {why}.")


def author_capacity(book: Book) -> None:
    suite = "capacity"
    feb = D("2024-02-10")
    eom = F + "EndOfMonth"
    rows = [
        ("row-100000", eom, [F2(1, 100000, feb)], "exactly 100,000 elements are permitted"),
        ("column-100000", eom, [F2(100000, 1, feb)], "exactly 100,000 elements as a column"),
        ("rectangle-100000", eom, [F2(250, 400, feb)], "exactly 100,000 elements as 250x400"),
        ("row-100001", eom, [F2(1, 100001, feb)], "100,001 elements exceed the cap"),
        ("rectangle-100001", eom, [F2(11, 9091, feb)], "11x9091 is 100,001 elements"),
        ("invalid-100001", eom, [F2(1, 100001, S("bad"))], "the cap is decided before any element is read"),
        ("broadcast-100000", F + "AddDays", [feb, F2(1, 100000, N("1"))], "broadcast target of 100,000"),
        ("broadcast-100001", F + "AddDays", [F2(1, 100001, feb), N("1")], "broadcast target of 100,001"),
        ("control-before-capacity", F + "AddMonths", [F2(1, 100001, feb), N("1"), N("1")],
         "control validation precedes the capacity check"),
        ("mismatch-before-capacity", F + "AddDays", [F2(1, 100000, feb), F2(1, 100001, N("1"))],
         "shape resolution precedes the capacity check"),
        ("propagated-100000", F + "DaysInMonth", [F2(1, 100000, E("#N/A"))],
         "incoming errors propagate at every position"),
    ]
    for name, fn, args, why in rows:
        book.add(suite, name, fn, args, f"Section 5.2: {why}.")


def author_host(book: Book) -> None:
    suite = "host"
    feb = D("2024-02-10")
    rows = [
        ("diagnostic-direct", HOST_FUNCTION, [], "direct", "no worksheet host: 1900 serial contract"),
        ("diagnostic-ws1900", HOST_FUNCTION, [], "ws1900", "identifiable 1900 worksheet caller"),
        ("diagnostic-ws1904", HOST_FUNCTION, [], "ws1904", "identifiable 1904 worksheet caller"),
        ("value-ws1900", F + "EndOfMonth", [feb], "ws1900", "1900 worksheet evaluates normally"),
        ("value-ws1904", F + "EndOfMonth", [feb], "ws1904", "1904 worksheet is refused at call level"),
        ("array-ws1904", F + "EndOfMonth", [A2(1, 3, [feb] * 3)], "ws1904", "one call-level #N/A, no array"),
        ("invalid-ws1904", F + "EndOfMonth", [S("bad")], "ws1904", "the host guard precedes parsing"),
        ("shift-ws1904", F + "AddDays", [feb, N("1")], "ws1904", "every value-taking function is guarded"),
        ("year-ws1904", F + "IsLeapYear", [N("2024")], "ws1904", "year functions are value-taking too"),
        ("control-ws1904", F + "AddMonths", [feb, N("1"), N("1")], "ws1904", "the host guard precedes controls"),
        ("propagated-na-direct", F + "EndOfMonth", [E("#N/A")], "direct",
         "a propagated #N/A is value-identical to host #N/A; provenance is the condition"),
        ("propagated-na-ws1900", F + "EndOfMonth", [E("#N/A")], "ws1900",
         "a 1900 worksheet leaves propagation as the source of #N/A"),
        ("propagated-na-ws1904", F + "EndOfMonth", [E("#N/A")], "ws1904",
         "in a 1904 worksheet the same input is host refusal, not propagation"),
    ]
    for name, fn, args, context, why in rows:
        book.add(suite, name, fn, args, f"Sections 4 and 6: {why}.", context)


def author_propagation(book: Book) -> None:
    suite = "propagation"
    feb = D("2024-02-10")
    for code in ERR_ORDER:
        name = slug(code) or "error"
        book.add(suite, f"date-{name}", F + "EndOfMonth", [E(code)],
                 "Section 5.2: incoming errors propagate verbatim at a date position.")
        book.add(suite, f"integer-{name}", F + "AddDays", [feb, E(code)],
                 "Section 5.2: incoming errors propagate verbatim at an integer position.")
        book.add(suite, f"pillar-{name}", F + "DateFromPillar", [feb, E(code)],
                 "Section 5.2: incoming errors propagate verbatim at a Pillar position.")
        book.add(suite, f"start-{name}", F + "PillarFromDates", [E(code), feb],
                 "Section 5.2: incoming errors propagate verbatim at StartDate.")
    rows = [
        ("first-error-wins", F + "AddDays", [E("#REF!"), E("#DIV/0!")], "the first error in signature order"),
        ("parse-before-error", F + "AddDays", [S("bad"), E("#N/A")], "the first failing argument wins"),
        ("locator-first-error", F + "NthWeekdayOfMonth", [E("#NUM!"), N("13"), N("0"), N("9")],
         "an incoming error at YearIn precedes later domain failures"),
        ("control-error-call", F + "AddMonths", [E("#N/A"), N("1"), E("#NAME?")],
         "an error control propagates at call level before elements"),
    ]
    for name, fn, args, why in rows:
        book.add(suite, name, fn, args, f"Section 5.2: {why}.")


SURFACE_GOOD: dict[str, list[Val]] = {
    "DayOfWeek": [S("2024-02-10")], "DaysInMonth": [S("2024-02-10")], "DaysInYear": [N("2024")],
    "BeginOfMonth": [S("2024-02-10")], "EndOfMonth": [S("2024-02-10")],
    "BeginOfQuarter": [S("2024-02-10")], "EndOfQuarter": [S("2024-02-10")],
    "BeginOfYear": [S("2024-02-10")], "EndOfYear": [S("2024-02-10")],
    "IsMonthEnd": [S("2024-02-29")], "IsQuarterEnd": [S("2024-03-31")], "IsYearEnd": [S("2024-12-31")],
    "IsLeapYear": [N("2024")], "AddDays": [S("2024-02-10"), N("20")], "AddWeeks": [S("2024-02-10"), N("3")],
    "AddMonths": [S("2024-01-31"), N("1")], "AddYears": [S("2024-02-29"), N("1")],
    "NthWeekdayOfMonth": [N("2024"), N("2"), N("5"), N("1")],
    "LastWeekdayOfMonth": [N("2024"), N("2"), N("5")],
    "PillarFromDates": [S("2024-01-15"), S("2024-07-15")], "DateFromPillar": [S("2024-01-15"), S("6M")],
}
SURFACE_BAD = {"date": S("2024/02/10"), "int": S("1"), "year": S("2024"), "month": S("2"),
               "weekday": S("5"), "occurrence": S("1"), "pillar": S("1X")}


def author_surface(book: Book) -> None:
    for name, good in SURFACE_GOOD.items():
        fn = F + name
        book.add("surface", slug(name, "ok"), fn, good, "Section 2: supported surface, ISO text inputs.")
        for index, (param, kind) in enumerate(FUNCTIONS[fn].params):
            args = list(good)
            args[index] = SURFACE_BAD[kind]
            book.add("surface", slug(name, "bad", param), fn, args,
                     f"Section 3: an invalid {param} is an element-level #VALUE!.")


AUTHORS = (
    author_date_input, author_integer_input, author_controls, author_boundaries, author_years,
    author_arithmetic, author_locators, author_pillar_parse, author_pillar_format,
    author_pillar_roundtrip, author_shapes, author_capacity, author_host, author_propagation,
    author_surface,
)


def build_cases() -> list[Case]:
    book = Book()
    for author in AUTHORS:
        author(book)
    return book.cases


# ---------------------------------------------------------------------------
# TSV serialization
# ---------------------------------------------------------------------------


def expect_type(value: Val) -> str:
    return {"lng": "Long", "date": "Date", "bool": "Boolean", "str": "String", "err": "Error",
            "arr2": "Array", "fill2": "Array"}[value.kind]


def input_kind(args: list[Val]) -> str:
    kinds = []
    for arg in args:
        kinds.append(arg.kind if arg.kind in SCALAR_KINDS else f"{arg.kind}({len(arg.items)})"
                     if arg.kind == "arr1" else f"{arg.kind}({arg.rows}x{arg.cols})")
    return ",".join(kinds) or "none"


def case_row(case: Case) -> list[str]:
    args = [encode(arg) for arg in case.args] + ["-"] * (5 - len(case.args))
    outcome = case.outcome
    return [case.id, case.suite, case.function, case.context, *args, input_kind(case.args),
            expect_type(outcome.value), encode(outcome.value), outcome.condition, outcome.level,
            case.rationale]


def render_tsv(cases: list[Case]) -> str:
    lines = ["\t".join(COLUMNS)]
    for case in cases:
        row = case_row(case)
        if any("\t" in cell or "\n" in cell for cell in row):
            raise FixtureError(f"fixture {case.id} contains a tab or newline")
        lines.append("\t".join(row))
    return "\n".join(lines) + "\n"


@dataclass(frozen=True)
class Row:
    id: str
    suite: str
    function: str
    context: str
    args: tuple[Val, ...]
    expect_type: str
    expected: Val
    condition: str
    level: str


def parse_tsv(text: str) -> list[Row]:
    lines = text.split("\n")
    if lines[-1] != "" or lines[0] != "\t".join(COLUMNS):
        raise FixtureError("TSV header or final newline is not canonical")
    rows: list[Row] = []
    seen: set[str] = set()
    for number, line in enumerate(lines[1:-1], 2):
        cells = line.split("\t")
        if len(cells) != len(COLUMNS):
            raise FixtureError(f"TSV line {number} has {len(cells)} cells")
        record = dict(zip(COLUMNS, cells))
        args = tuple(decode(cell) for cell in cells[4:9] if cell != "-")
        if record["id"] in seen or record["context"] not in CONTEXTS or record["level"] not in LEVELS:
            raise FixtureError(f"TSV line {number} has an invalid id, context or level")
        seen.add(record["id"])
        rows.append(Row(record["id"], record["suite"], record["function"], record["context"], args,
                        record["expect_type"], decode(record["expected"]), record["condition"],
                        record["level"]))
    return rows


# ---------------------------------------------------------------------------
# VBA emission (derived from parsed TSV rows only)
# ---------------------------------------------------------------------------

RULE = "'" + "-" * 78
BANNER = "'" + "=" * 78


def vba_string(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def vba_number(payload: str) -> str:
    """A literal the VBE keeps verbatim on import and export.

    Integral values drop a redundant fraction, and an integral value beyond the
    Long range carries the Double suffix the VBE would otherwise add.
    """
    if "E" in payload:
        if Decimal(payload).copy_abs() < Decimal("1E+15"):
            raise FixtureError(f"exponent literal {payload} would be rewritten by the VBE")
        return payload
    number = Decimal(payload)
    if number != number.to_integral_value():
        return payload
    text = str(int(number))
    return text + "#" if number.copy_abs() > LONG_MAX else text


def vba_value(value: Val) -> str:
    kind = value.kind
    if kind == "date":
        day = dt.date.fromisoformat(value.payload)
        return f"DateSerial({day.year}, {day.month}, {day.day})"
    if kind == "datetime":
        moment = dt.datetime.fromisoformat(value.payload)
        return (f"DateSerial({moment.year}, {moment.month}, {moment.day}) + "
                f"TimeSerial({moment.hour}, {moment.minute}, {moment.second})")
    simple = {
        "num": lambda: f"CDbl({vba_number(value.payload)})", "lng": lambda: f"CLng({value.payload})",
        "str": lambda: vba_string(value.payload), "bool": lambda: value.payload.title(),
        "empty": lambda: "Empty", "null": lambda: "Null",
        "err": lambda: f"CVErr({ERROR_CONSTANTS.get(value.payload, '')})",
    }
    if kind in simple:
        return simple[kind]()
    items = ", ".join(vba_value(item) for item in value.items)
    if kind == "arr1":
        return f"Array({items})"
    if kind == "arr2":
        return f"M2({value.rows}, {value.cols}, Array({items}))"
    if kind == "fill2":
        return f"Fill2({value.rows}, {value.cols}, {items})"
    raise FixtureError(f"cannot emit value kind {kind}")


def split_arguments(text: str) -> list[str]:
    parts: list[str] = []
    depth = 0
    quoted = False
    start = 0
    for index, char in enumerate(text):
        if char == '"':
            quoted = not quoted
        elif not quoted and char == "(":
            depth += 1
        elif not quoted and char == ")":
            depth -= 1
        elif not quoted and depth == 0 and char == ",":
            parts.append(text[start:index + 1])
            start = index + 1
    parts.append(text[start:])
    return [part.strip() for part in parts]


def case_lines(number: int, block: str, row: Row) -> list[str]:
    args = ", ".join(vba_value(arg) for arg in row.args)
    call = (f"Fx({vba_string(row.id)}, {vba_string(row.suite)}, {vba_string(row.function)}, "
            f"{vba_string(row.context)}, {len(row.args)}, Array({args}), {vba_string(row.expect_type)}, "
            f"{vba_value(row.expected)}, {vba_string(row.condition)}, {vba_string(row.level)})")
    head = f"        Case {number}: {block} = "
    if len(head) + len(call) <= MAX_VBA_LINE:
        return [head + call]
    inner = call[len("Fx("):-1]
    parts = split_arguments(inner)
    lines = [head + "Fx( _"]
    for part in parts[:-1]:
        lines.append("                " + part + " _")
    lines.append("                " + parts[-1] + ")")
    if len(lines) > 24 or any(len(line) > MAX_VBA_LINE for line in lines):
        raise FixtureError(f"fixture {row.id} cannot be emitted within VBA line limits")
    return lines


def procedure_banner(name: str, sections: list[tuple[str, list[str]]]) -> list[str]:
    lines = ["'", BANNER, "'" + name.center(78).rstrip(), RULE]
    for title, body in sections:
        lines.append(f"' {title}")
        lines.extend(f"'   {text}" if text else "'" for text in body)
        lines.append("'")
    lines[-1] = BANNER
    return lines


def module_header(rows: list[Row], digest: str) -> list[str]:
    return [
        f'Attribute VB_Name = "{MODULE_NAME}"',
        BANNER,
        f"' MODULE: {MODULE_NAME}",
        RULE,
        "' GENERATED FILE - DO NOT EDIT",
        "'   Regenerate with: python3 tools/gen_fixtures.py --write",
        f"'   Canonical source: {TSV_PATH}",
        f"'   Source SHA-256: {digest}",
        f"'   Fixture cases: {len(rows)}",
        "'",
        "' PURPOSE",
        "'   Independently generated expected results for the KPR date-layer",
        "'   contract. Expectations were computed outside VBA from",
        "'   docs/DATE_LAYER_CONTRACT.md; this module only replays them as VBA values.",
        "'",
        "' PUBLIC SURFACE (test infrastructure, not supported API)",
        "'   KPR_Fixtures_Count       number of fixture cases",
        "'   KPR_Fixtures_Case        one case record, indexed 1 to Count",
        "'   KPR_Fixtures_SourceHash  SHA-256 of the canonical TSV",
        "'   KPR_FX_*                 field indexes of a case record",
        "'",
        "' CASE RECORD",
        "'   A zero-based Variant array: id, suite, function, context (direct,",
        "'   ws1900 or ws1904), supplied argument count, argument array, expected",
        "'   result type, expected value, originating condition and level. Omitted",
        "'   optional controls are trailing arguments beyond the supplied count.",
        "'   Array conditions list 1-based positions such as r1c2=DATE_TEXT_FORMAT;",
        "'   all=<condition> applies to every element of a filled array.",
        "'",
        "' DEPENDENCIES",
        "'   None. VBA built-ins only; no KPR production module is referenced.",
        "'",
        "' STATE OWNERSHIP",
        "'   None. Every call builds a fresh record.",
        "'",
        "' ERROR POLICY",
        "'   An index outside 1 to Count raises run-time error 9.",
        "'",
        "' UPDATED",
        f"'   {GENERATOR_UPDATED}",
        "'",
        "' AUTHOR",
        "'   Daniele Penza (generated by tools/gen_fixtures.py)",
        BANNER,
        "",
        RULE,
        "' MODULE SETTINGS",
        RULE,
        "    Option Explicit         'Force explicit variable declarations",
        "    Option Private Module   'Visible to this VBA project only",
        "",
        RULE,
        "' MODULE CONSTANTS",
        RULE,
        "    'Field indexes of a case record returned by KPR_Fixtures_Case",
        "        Public Const KPR_FX_ID            As Long = 0   'Stable semantic case identifier",
        "        Public Const KPR_FX_SUITE         As Long = 1   'Fixture suite name",
        "        Public Const KPR_FX_FUNCTION      As Long = 2   'Public function under test",
        "        Public Const KPR_FX_CONTEXT       As Long = 3   'direct, ws1900 or ws1904",
        "        Public Const KPR_FX_ARGC          As Long = 4   'Number of supplied arguments",
        "        Public Const KPR_FX_ARGS          As Long = 5   'Zero-based argument array",
        "        Public Const KPR_FX_EXPECT_TYPE   As Long = 6   'Long, Date, Boolean, String, Error, Array",
        "        Public Const KPR_FX_EXPECTED      As Long = 7   'Expected value or 1-based 2-D array",
        "        Public Const KPR_FX_CONDITION     As Long = 8   'Registry condition, blank on success",
        "        Public Const KPR_FX_LEVEL         As Long = 9   'element, call or blank",
        "",
        "    'Generated data identity",
        f"        Private Const FX_CASE_COUNT       As Long = {len(rows)}   'Number of fixture cases",
        f"        Private Const FX_BLOCK_SIZE       As Long = {BLOCK_SIZE}   'Cases per block procedure",
        f'        Private Const FX_SOURCE_HASH      As String = "{digest}"   \'Canonical TSV SHA-256',
        "",
        "    'Native Excel error numbers used with CVErr",
        "        Private Const FX_ERR_NULL         As Long = 2000   '#NULL!",
        "        Private Const FX_ERR_DIV0         As Long = 2007   '#DIV/0!",
        "        Private Const FX_ERR_VALUE        As Long = 2015   '#VALUE!",
        "        Private Const FX_ERR_REF          As Long = 2023   '#REF!",
        "        Private Const FX_ERR_NAME         As Long = 2029   '#NAME?",
        "        Private Const FX_ERR_NUM          As Long = 2036   '#NUM!",
        "        Private Const FX_ERR_NA           As Long = 2042   '#N/A",
    ]


def entry_points(block_count: int) -> list[str]:
    lines = ["", "", "'", RULE, "'", "'" + "FIXTURE ACCESS".center(78).rstrip(), "'", RULE, "'", ""]
    lines += ["Public Function KPR_Fixtures_Count() As Long"]
    lines += procedure_banner("KPR_Fixtures_Count", [
        ("PURPOSE", ["Returns the number of generated fixture cases."]),
        ("UPDATED", [GENERATOR_UPDATED]),
    ])
    lines += ["", RULE, "' RETURN COUNT", RULE, "    'The count is fixed at generation time.",
              "        KPR_Fixtures_Count = FX_CASE_COUNT", "", "End Function", "", ""]
    lines += ["Public Function KPR_Fixtures_SourceHash() As String"]
    lines += procedure_banner("KPR_Fixtures_SourceHash", [
        ("PURPOSE", ["Returns the SHA-256 of the canonical TSV this module was derived",
                     "from, so evidence can bind a run to one fixture set."]),
        ("UPDATED", [GENERATOR_UPDATED]),
    ])
    lines += ["", RULE, "' RETURN HASH", RULE, "    'The hash is fixed at generation time.",
              "        KPR_Fixtures_SourceHash = FX_SOURCE_HASH", "", "End Function", "", ""]
    lines += ["Public Function KPR_Fixtures_Case( _", "    ByVal Index As Long) _", "    As Variant"]
    lines += procedure_banner("KPR_Fixtures_Case", [
        ("PURPOSE", ["Returns one fixture case record."]),
        ("INPUTS", ["Index    1-based case number, 1 to KPR_Fixtures_Count."]),
        ("RETURNS", ["A zero-based Variant array indexed by the KPR_FX_* constants."]),
        ("ERROR POLICY", ["Raises run-time error 9 for an index outside the fixture range."]),
        ("UPDATED", [GENERATOR_UPDATED]),
    ])
    lines += ["", RULE, "' GUARD INDEX", RULE, "    'Reject an index outside the generated range.",
              "        If Index < 1 Or Index > FX_CASE_COUNT Then",
              f'            Err.Raise 9, "{MODULE_NAME}.KPR_Fixtures_Case", "Fixture index out of range"',
              "        End If", "", RULE, "' DISPATCH TO BLOCK", RULE,
              "    'Cases are split into fixed-size blocks to stay within VBA procedure limits.",
              "        Select Case (Index - 1) \\ FX_BLOCK_SIZE"]
    for block in range(block_count):
        lines.append(f"            Case {block}: KPR_Fixtures_Case = FixtureBlock{block + 1:03d}(Index)")
    lines += ["        End Select", "", "End Function"]
    return lines


def helpers() -> list[str]:
    lines = ["", "", "'", RULE, "'", "'" + "RECORD BUILDERS".center(78).rstrip(), "'", RULE, "'", ""]
    lines += ["Private Function Fx( _", "    ByVal CaseId As String, _", "    ByVal Suite As String, _",
              "    ByVal FunctionName As String, _", "    ByVal Context As String, _",
              "    ByVal ArgCount As Long, _", "    ByVal Args As Variant, _",
              "    ByVal ExpectType As String, _", "    ByVal Expected As Variant, _",
              "    ByVal Condition As String, _", "    ByVal Level As String) _", "    As Variant"]
    lines += procedure_banner("Fx", [
        ("PURPOSE", ["Packs one case into the zero-based record read by KPR_Fixtures_Case."]),
        ("UPDATED", [GENERATOR_UPDATED]),
    ])
    lines += ["", RULE, "' DECLARE", RULE,
              "    Dim Record(0 To 9)     As Variant   'Case record indexed by KPR_FX_*", "",
              RULE, "' PACK RECORD", RULE, "    'Store every field in its fixed position.",
              "        Record(KPR_FX_ID) = CaseId", "        Record(KPR_FX_SUITE) = Suite",
              "        Record(KPR_FX_FUNCTION) = FunctionName", "        Record(KPR_FX_CONTEXT) = Context",
              "        Record(KPR_FX_ARGC) = ArgCount", "        Record(KPR_FX_ARGS) = Args",
              "        Record(KPR_FX_EXPECT_TYPE) = ExpectType", "        Record(KPR_FX_EXPECTED) = Expected",
              "        Record(KPR_FX_CONDITION) = Condition", "        Record(KPR_FX_LEVEL) = Level",
              "        Fx = Record", "", "End Function", "", ""]
    lines += ["Private Function M2( _", "    ByVal RowCount As Long, _", "    ByVal ColumnCount As Long, _",
              "    ByVal Items As Variant) _", "    As Variant"]
    lines += procedure_banner("M2", [
        ("PURPOSE", ["Builds a 1-based two-dimensional Variant array from row-major items."]),
        ("UPDATED", [GENERATOR_UPDATED]),
    ])
    lines += ["", RULE, "' DECLARE", RULE,
              "    Dim Output()           As Variant   'Two-dimensional result",
              "    Dim RowIndex           As Long      'Current output row",
              "    Dim ColumnIndex        As Long      'Current output column", "",
              RULE, "' FILL ROW-MAJOR", RULE, "    'Items are listed row by row, matching the TSV encoding.",
              "        ReDim Output(1 To RowCount, 1 To ColumnCount)",
              "        For RowIndex = 1 To RowCount", "            For ColumnIndex = 1 To ColumnCount",
              "                Output(RowIndex, ColumnIndex) = Items(LBound(Items) + " +
              "(RowIndex - 1) * ColumnCount + ColumnIndex - 1)",
              "            Next ColumnIndex", "        Next RowIndex", "        M2 = Output", "",
              "End Function", "", ""]
    lines += ["Private Function Fill2( _", "    ByVal RowCount As Long, _", "    ByVal ColumnCount As Long, _",
              "    ByVal Item As Variant) _", "    As Variant"]
    lines += procedure_banner("Fill2", [
        ("PURPOSE", ["Builds a 1-based two-dimensional Variant array holding one value in",
                     "every position; used for the capacity fixtures."]),
        ("UPDATED", [GENERATOR_UPDATED]),
    ])
    lines += ["", RULE, "' DECLARE", RULE,
              "    Dim Output()           As Variant   'Two-dimensional result",
              "    Dim RowIndex           As Long      'Current output row",
              "    Dim ColumnIndex        As Long      'Current output column", "",
              RULE, "' FILL", RULE, "    'Every position receives the same value.",
              "        ReDim Output(1 To RowCount, 1 To ColumnCount)",
              "        For RowIndex = 1 To RowCount", "            For ColumnIndex = 1 To ColumnCount",
              "                Output(RowIndex, ColumnIndex) = Item",
              "            Next ColumnIndex", "        Next RowIndex", "        Fill2 = Output", "",
              "End Function"]
    return lines


def blocks(rows: list[Row]) -> list[str]:
    lines = ["", "", "'", RULE, "'", "'" + "GENERATED CASE BLOCKS".center(78).rstrip(), "'", RULE, "'"]
    for block_index in range(0, len(rows), BLOCK_SIZE):
        number = block_index // BLOCK_SIZE + 1
        name = f"FixtureBlock{number:03d}"
        first, last = block_index + 1, min(block_index + BLOCK_SIZE, len(rows))
        lines += ["", "", f"Private Function {name}( _", "    ByVal Index As Long) _", "    As Variant"]
        lines += procedure_banner(name, [("PURPOSE", [f"Returns fixture cases {first} to {last}."]),
                                         ("UPDATED", [GENERATOR_UPDATED])])
        lines += ["", RULE, "' SELECT CASE", RULE, "    'One generated record per case number.",
                  "        Select Case Index"]
        for offset, row in enumerate(rows[block_index:block_index + BLOCK_SIZE]):
            lines += case_lines(block_index + offset + 1, name, row)
        lines += ["        End Select", "", "End Function"]
    return lines


def render_vba(tsv_text: str) -> str:
    rows = parse_tsv(tsv_text)
    digest = hashlib.sha256(tsv_text.encode("ascii")).hexdigest()
    block_count = (len(rows) + BLOCK_SIZE - 1) // BLOCK_SIZE
    lines = module_header(rows, digest) + entry_points(block_count) + helpers() + blocks(rows)
    text = "\r\n".join(lines) + "\r\n"
    text.encode("ascii")
    return text


# ---------------------------------------------------------------------------
# Contract cross-checks
# ---------------------------------------------------------------------------


def contract_registry(root: Path) -> list[str]:
    text = (root / CONTRACT_PATH).read_text(encoding="utf-8")
    section = text.split("## 🆔 7. Condition identifier registry", 1)[1].split("### 🛡️ 7.1", 1)[0]
    return re.findall(r"^\| `([A-Z0-9_]+)` \|", section, re.M)


def contract_functions(root: Path) -> list[str]:
    text = (root / CONTRACT_PATH).read_text(encoding="utf-8")
    return re.findall(r"^Public Function (KPR_Dates_\w+)\(", text, re.M)


def coverage_problems(root: Path, cases: list[Case]) -> list[str]:
    problems: list[str] = []
    if contract_functions(root) != list(ALL_FUNCTIONS):
        problems.append("generator function list differs from the contract public surface")
    covered_functions = {case.function for case in cases}
    problems += [f"no fixture for {name}" for name in ALL_FUNCTIONS if name not in covered_functions]
    conditions: set[str] = set()
    for case in cases:
        for part in case.outcome.condition.split(";"):
            if part:
                conditions.add(part.split("=")[-1])
    registry = contract_registry(root)
    if len(registry) < 30:
        problems.append("could not read the contract condition registry")
    for condition in registry:
        if condition not in conditions and condition not in UNFIXTURABLE:
            problems.append(f"registry condition {condition} has no fixture")
    problems += [f"condition {c} is not in the contract registry" for c in sorted(conditions - set(registry))]
    sizes = {(arg.rows * arg.cols) for case in cases for arg in case.args if arg.kind == "fill2"}
    problems += [f"no capacity fixture of {n} elements" for n in (CAPACITY, CAPACITY + 1) if n not in sizes]
    return problems


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------


def generate(root: Path) -> tuple[str, str]:
    cases = build_cases()
    problems = coverage_problems(root, cases)
    if problems:
        raise FixtureError("; ".join(problems))
    tsv = render_tsv(cases)
    return tsv, render_vba(tsv)


def normalized(path: Path) -> str | None:
    if not path.is_file():
        return None
    return path.read_bytes().decode("ascii").replace("\r\n", "\n")


def check(root: Path) -> list[str]:
    tsv, vba = generate(root)
    stale = []
    if normalized(root / TSV_PATH) != tsv:
        stale.append(TSV_PATH)
    if normalized(root / VBA_PATH) != vba.replace("\r\n", "\n"):
        stale.append(VBA_PATH)
    return stale


def write(root: Path) -> int:
    tsv, vba = generate(root)
    (root / TSV_PATH).parent.mkdir(parents=True, exist_ok=True)
    (root / TSV_PATH).write_bytes(tsv.encode("ascii"))
    (root / VBA_PATH).write_bytes(vba.encode("ascii"))
    print(f"Wrote {TSV_PATH} and {VBA_PATH} ({tsv.count(chr(10)) - 1} cases).")
    return 0


# Contract statements, each asserted against the reference model by --self-test.
EXAMPLES: tuple[tuple[str, list[Val], str, str], ...] = (
    (F + "IsLeapYear", [N("1900")], "bool:FALSE", ""),
    (F + "DaysInYear", [N("1900")], "lng:365", ""),
    (F + "EndOfMonth", [S("1900-02-29")], "err:#VALUE!", "DATE_TEXT_IMPOSSIBLE"),
    (F + "EndOfMonth", [S("2025-02-29")], "err:#VALUE!", "DATE_TEXT_IMPOSSIBLE"),
    (F + "EndOfMonth", [S("31/12/2026")], "err:#VALUE!", "DATE_TEXT_LOCALE"),
    (F + "EndOfMonth", [S("45292.5")], "err:#VALUE!", "DATE_TEXT_NUMERIC"),
    (F + "EndOfMonth", [N("60")], "err:#NUM!", "DATE_WINDOW"),
    (F + "AddDays", [D("2024-01-01"), N("2147483648.5")], "err:#NUM!", "INTEGER_RANGE"),
    (F + "BeginOfYear", [D("1900-06-01")], "err:#NUM!", "RESULT_WINDOW"),
    (F + "EndOfQuarter", [D("1900-03-15")], "date:1900-03-31", ""),
    (F + "BeginOfQuarter", [D("1900-05-15")], "date:1900-04-01", ""),
    (F + "EndOfYear", [D("1900-03-15")], "date:1900-12-31", ""),
    (F + "AddMonths", [D("2024-02-29"), N("1"), B(True)], "date:2024-03-31", ""),
    (F + "AddMonths", [D("2024-02-29"), N("1")], "date:2024-03-29", ""),
    (F + "DateFromPillar", [D("2024-01-31"), S("1M")], "date:2024-02-29", ""),
    (F + "DateFromPillar", [D("2024-01-15"), S("-ON")], "err:#VALUE!", "PILLAR_ALIAS_SIGNED"),
    (F + "DateFromPillar", [D("2024-01-15"), S("1M2M")], "err:#VALUE!", "PILLAR_DUPLICATE_UNIT"),
    (F + "DateFromPillar", [D("2024-01-15"), S("1 M")], "err:#VALUE!", "PILLAR_TOKEN_MALFORMED"),
    (F + "PillarFromDates", [D("2024-01-15"), D("2024-02-08")], 'str:"3W"', ""),
    (F + "PillarFromDates", [D("2024-01-15"), D("2024-02-11")], 'str:"1M"', ""),
    (F + "PillarFromDates", [D("2024-01-15"), D("2024-02-09")], 'str:"3W"', ""),
    (F + "PillarFromDates", [D("2024-01-15"), D("2024-02-10")], 'str:"1M"', ""),
    (F + "PillarFromDates", [D("2024-01-31"), D("2024-02-25")], 'str:"1M"', ""),
    (F + "PillarFromDates", [D("2024-01-15"), D("2034-01-29")], 'str:"10Y"', ""),
    (F + "PillarFromDates", [D("9999-12-20"), D("9999-12-31"), S("NEAREST")], 'str:"1W"', ""),
    (F + "PillarFromDates", [D("9999-12-20"), D("9999-12-31"), S("FLOOR")], 'str:"1W"', ""),
    (F + "PillarFromDates", [D("9999-12-20"), D("9999-12-31"), S("CEILING")], "err:#NUM!", "RESULT_WINDOW"),
    (F + "PillarFromDates", [D("2024-01-01"), D("2024-01-12")], 'str:"2W"', ""),
    (F + "EndOfMonth", [F2(1, 100001, S("bad"))], "err:#NUM!", "CAPACITY_EXCEEDED"),
    (F + "AddDays", [A2(1, 3, [D("2024-01-01")] * 3), A2(3, 1, [N("1")] * 3)], "err:#VALUE!",
     "SHAPE_MISMATCH"),
)


def self_test(root: Path) -> int:
    failures: list[str] = []
    for function, args, expected, condition in EXAMPLES:
        outcome = evaluate(function, args, "direct")
        if encode(outcome.value) != expected or outcome.condition != condition:
            failures.append(f"{function}{[encode(a) for a in args]} -> {encode(outcome.value)} "
                            f"{outcome.condition}; contract expects {expected} {condition}")
    ws = evaluate(F + "EndOfMonth", [E("#N/A")], "ws1904")
    direct = evaluate(F + "EndOfMonth", [E("#N/A")], "direct")
    if ws.value != direct.value or ws.condition == direct.condition:
        failures.append("host #N/A and propagated #N/A must be value-identical with distinct conditions")
    tsv, vba = generate(root)
    if (tsv, vba) != generate(root):
        failures.append("generation is not deterministic")
    rows = parse_tsv(tsv)
    if [encode(v) for r in rows for v in r.args] != [encode(decode(encode(v))) for r in rows for v in r.args]:
        failures.append("value encoding does not round-trip")
    for number, line in enumerate(vba.split("\r\n"), 1):
        if len(line) > MAX_VBA_LINE:
            failures.append(f"VBA line {number} exceeds {MAX_VBA_LINE} characters")
    if re.search(r"\bKPR_(?:Core_|DATES_DAYS\b|Dates_)", vba.split("GENERATED CASE BLOCKS")[1]
                 .replace('"KPR_Dates_', '"')):
        failures.append("generated case blocks reference a production module")
    for name, value in (("single", S("1M")), ("month-first", S("1D1M"))):
        token = evaluate(F + "DateFromPillar", [D("2024-01-31"), value], "direct")
        if token.value.kind != "date":
            failures.append(f"accepted pillar example {name} did not parse")
    with tempfile.TemporaryDirectory() as temporary:
        scratch = Path(temporary)
        (scratch / "docs").mkdir()
        shutil.copy(root / CONTRACT_PATH, scratch / CONTRACT_PATH)
        write_quietly(scratch)
        if check(scratch):
            failures.append("--check reports fresh output as stale")
        stale = scratch / TSV_PATH
        stale.write_bytes(stale.read_bytes().replace(b"\tDate\t", b"\tLong\t", 1))
        if check(scratch) != [TSV_PATH]:
            failures.append("--check does not detect a stale TSV")
        write_quietly(scratch)
        module = scratch / VBA_PATH
        module.write_bytes(module.read_bytes().replace(b"CLng(", b"CInt(", 1))
        if check(scratch) != [VBA_PATH]:
            failures.append("--check does not detect a stale VBA module")
    for failure in failures:
        print(f"FAIL {failure}")
    print(f"{'FAIL' if failures else 'PASS'} fixture generator self-test: "
          f"{len(EXAMPLES)} contract examples, {len(rows)} generated cases.")
    return 1 if failures else 0


def write_quietly(root: Path) -> None:
    tsv, vba = generate(root)
    (root / TSV_PATH).parent.mkdir(parents=True, exist_ok=True)
    (root / VBA_PATH).parent.mkdir(parents=True, exist_ok=True)
    (root / TSV_PATH).write_bytes(tsv.encode("ascii"))
    (root / VBA_PATH).write_bytes(vba.encode("ascii"))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=Path.cwd())
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true", help="regenerate the TSV and VBA module")
    mode.add_argument("--check", action="store_true", help="fail if committed output is stale")
    mode.add_argument("--self-test", action="store_true", help="exercise the model and the check mode")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    try:
        if args.write:
            return write(root)
        if args.self_test:
            return self_test(root)
        stale = check(root)
    except (FixtureError, InvalidOperation, OSError, ValueError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    for path in stale:
        print(f"STALE {path}: run python3 tools/gen_fixtures.py --write and commit the result.")
    if not stale:
        print(f"PASS generated fixtures are current: {TSV_PATH}, {VBA_PATH}.")
    return 1 if stale else 0


if __name__ == "__main__":
    sys.exit(main())
