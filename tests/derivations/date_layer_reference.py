#!/usr/bin/env python3
"""Derive the expected results of the date-layer test cases from the contracts.

A reference model of the 22 KPR_Dates_* worksheet functions, written with
Python standard-library date arithmetic from the calculation contracts CALC-02
to CALC-10 and the data contract's Function arguments. It never imports, runs
or translates the VBA. Run on one call fixture, it prints the result of the
call and, when the call fails, its condition, as docs/methodology/TEST_CASES.md
(Derivation scripts and Worksheet-function cases) defines:

    python -I tests/derivations/date_layer_reference.py tests/fixtures/<case-id>.json

The model is the oracle of the project's earlier date-layer fixture generator,
carried over with one correction: a whole number given as a Double is checked
in its binary value, as VBA holds it. Its values are illustrative: the project
derived them.
"""

from __future__ import annotations

import datetime as dt
import json
import math
import re
import sys
from dataclasses import dataclass
from decimal import ROUND_FLOOR, Decimal
from pathlib import Path
from typing import Callable, Iterable, NamedTuple, Union

MIN_DATE = dt.date(1900, 3, 1)
MAX_DATE = dt.date(9999, 12, 31)
MIN_ORD = MIN_DATE.toordinal()
MAX_ORD = MAX_DATE.toordinal()
EPOCH_ORD = dt.date(1899, 12, 30).toordinal()
LONG_MIN = -2147483648
LONG_MAX = 2147483647
CAPACITY = 100000

ERRORS = frozenset({"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"})

# The calculation contract of each function: its output quantity and the
# _status quantity that holds a failing call's condition.
CONTRACTS: dict[str, tuple[str, str, str]] = {
    "KPR_Dates_DayOfWeek": ("CALC-02", "weekday_number", "day_of_week_status"),
    "KPR_Dates_DaysInMonth": ("CALC-03", "days_in_month", "month_year_length_status"),
    "KPR_Dates_DaysInYear": ("CALC-03", "days_in_year", "month_year_length_status"),
    "KPR_Dates_IsLeapYear": ("CALC-03", "leap_year_flag", "month_year_length_status"),
    "KPR_Dates_BeginOfMonth": ("CALC-04", "period_begin_date", "period_boundary_status"),
    "KPR_Dates_EndOfMonth": ("CALC-04", "period_end_date", "period_boundary_status"),
    "KPR_Dates_BeginOfQuarter": ("CALC-04", "period_begin_date", "period_boundary_status"),
    "KPR_Dates_EndOfQuarter": ("CALC-04", "period_end_date", "period_boundary_status"),
    "KPR_Dates_BeginOfYear": ("CALC-04", "period_begin_date", "period_boundary_status"),
    "KPR_Dates_EndOfYear": ("CALC-04", "period_end_date", "period_boundary_status"),
    "KPR_Dates_IsMonthEnd": ("CALC-05", "period_end_flag", "period_end_status"),
    "KPR_Dates_IsQuarterEnd": ("CALC-05", "period_end_flag", "period_end_status"),
    "KPR_Dates_IsYearEnd": ("CALC-05", "period_end_flag", "period_end_status"),
    "KPR_Dates_AddDays": ("CALC-06", "shifted_date", "date_shift_status"),
    "KPR_Dates_AddWeeks": ("CALC-06", "shifted_date", "date_shift_status"),
    "KPR_Dates_AddMonths": ("CALC-06", "shifted_date", "date_shift_status"),
    "KPR_Dates_AddYears": ("CALC-06", "shifted_date", "date_shift_status"),
    "KPR_Dates_NthWeekdayOfMonth": ("CALC-07", "weekday_date", "weekday_locator_status"),
    "KPR_Dates_LastWeekdayOfMonth": ("CALC-07", "weekday_date", "weekday_locator_status"),
    "KPR_Dates_DateFromPillar": ("CALC-08", "pillar_date", "date_from_pillar_status"),
    "KPR_Dates_PillarFromDates": ("CALC-09", "pillar_token", "pillar_from_dates_status"),
    "KPR_Dates_HostDateSystem": ("CALC-10", "host_date_system", "host_date_system_status"),
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
# Reference model of the date-layer contracts
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
# The VBA parser's LooksNumeric rule: an optional sign, ASCII digits and at
# most one decimal point, with at least one digit ("1.", ".5" and "+.5" count).
NUMERIC_TEXT = re.compile(r"[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)")


def parse_date(value: Val) -> dt.date | Fail:
    """A date value (data contract, Date values) under the 1900 serial contract."""
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


INTEGER_TOLERANCE = 1e-9


def parse_long(value: Val) -> int | Fail:
    """A whole number (data contract): range precedes integrality; a residue within 1E-9 rounds."""
    kind = value.kind
    if kind == "err":
        return fail(value.payload, "INPUT_ERROR_PROPAGATED")
    if kind == "empty":
        return fail("#VALUE!", "INPUT_BLANK_REQUIRED")
    if kind not in {"num", "lng"}:
        return fail("#VALUE!", "INTEGER_TYPE_REJECTED")
    if kind == "lng":
        return int(value.payload)
    # A num is a Double: its range and its distance from the nearest whole
    # number are those of its binary value, not of the decimal literal, so
    # 3.000000001, about 1.0000000827E-9 above 3, is a fraction.
    number = float(value.payload)
    if number < LONG_MIN or number > LONG_MAX:
        return fail("#NUM!", "INTEGER_RANGE")
    nearest = math.floor(number + 0.5)
    if abs(number - nearest) > INTEGER_TOLERANCE:
        return fail("#VALUE!", "INTEGER_FRACTION")
    return int(nearest)


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
    """A pillar token (CALC-08): the signed (month delta, day delta) it names."""
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
    """A control value (data contract): omitted and Empty select the default; others are strict."""
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
    """Calendar-month shift: clip mode, or EOM-preserving mode when keep_eom is set
    and the start is a month end; None when the result leaves the supported window."""
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
    """CALC-09: the 1W to 3W anchors plus the floor and ceiling month anchors."""
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
    """Arrays and evaluation: None for a scalar, otherwise (rows, columns)."""
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
    """Apply the call-level stages of Arrays and evaluation, then each element."""
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


def derive(fixture: dict[str, object]) -> list[dict[str, object]]:
    """The outputs of one call fixture: the result and, on failure, the condition."""
    function = fixture["function"]
    context = fixture["context"]
    arguments = fixture["arguments"]
    if not isinstance(function, str) or function not in CONTRACTS:
        raise FixtureError(f"unknown function: {function!r}")
    if not isinstance(context, str) or not isinstance(arguments, list):
        raise FixtureError("context must be a text and arguments a list")
    outcome = evaluate(function, [decode(str(item)) for item in arguments], context)
    _, quantity, status = CONTRACTS[function]
    outputs: list[dict[str, object]] = [
        {"quantity": quantity, "scope": None, "text": encode(outcome.value)}]
    if outcome.condition:
        outputs.append({"quantity": status, "scope": None, "text": outcome.condition})
    return outputs


def main(argv: list[str]) -> int:
    if len(argv) != 1:
        print("usage: date_layer_reference.py <fixture.json>", file=sys.stderr)
        return 2
    fixture = json.loads(Path(argv[0]).read_text(encoding="utf-8"))
    print(json.dumps({"outputs": derive(fixture)}))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
