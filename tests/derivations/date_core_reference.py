#!/usr/bin/env python3
"""Reference computations for the expected values of tests/modules/TEST_CoreDates.bas.

The calendar rules use Python's own calendar, which is independent of the VBA. The
pillar format is a second implementation of the calculation contract CALC-09
(docs/methodology/CALCULATION_CONTRACTS.md), written from the contract. The author of the VBA
wrote it too, so it is a cross-check that catches a slip in either, not an independent
reference: its values stay illustrative (docs/methodology/TEST_CASES.md, Reference classes).

A serial is the Excel 1900 date system's day number, which is the number of days since
1899-12-30 from 1900-03-01 on, the first day of the supported window.

Run it to print the worked examples of CALC-09.
"""

from __future__ import annotations

import calendar
import datetime as dt

EPOCH = dt.date(1899, 12, 30)
WINDOW_FIRST = dt.date(1900, 3, 1)
WINDOW_LAST = dt.date(9999, 12, 31)

# One candidate of a pillar: (is a month anchor, its count, its anchor date).
Candidate = tuple[bool, int, dt.date]


def serial(day: dt.date) -> int:
    return (day - EPOCH).days


def parse(text: str) -> dt.date:
    year, month, day = (int(part) for part in text.split("-"))
    return dt.date(year, month, day)


def in_window(day: dt.date) -> bool:
    return WINDOW_FIRST <= day <= WINDOW_LAST


def month_length(year: int, month: int) -> int:
    return calendar.monthrange(year, month)[1]


def weekday_number(day: dt.date, monday_base: bool) -> int:
    """Monday is 1 and Sunday is 7, or Sunday is 1 and Saturday is 7."""
    iso = day.isoweekday()
    return iso if monday_base else (iso % 7) + 1


def python_weekday(number: int, monday_base: bool) -> int:
    """Python's weekday, Monday 0 to Sunday 6, of a weekday number in a base."""
    return (number - 1) if monday_base else (number - 2) % 7


def add_days(day: dt.date, count: int) -> dt.date | None:
    """The shifted date, or None outside the window."""
    try:
        shifted = day + dt.timedelta(days=count)
    except OverflowError:
        return None
    return shifted if in_window(shifted) else None


def add_months(day: dt.date, count: int, keep_month_end: bool) -> dt.date | None:
    """The date shifted by months, clipped to a short month, or None outside the window."""
    year, month = divmod(day.year * 12 + (day.month - 1) + count, 12)
    month += 1
    if year < WINDOW_FIRST.year or year > WINDOW_LAST.year:
        return None
    is_month_end = day.day == month_length(day.year, day.month)
    last = month_length(year, month)
    shifted = dt.date(year, month, last if (keep_month_end and is_month_end) else min(day.day, last))
    return shifted if in_window(shifted) else None


def weekdays_of_month(year: int, month: int, number: int, monday_base: bool) -> list[dt.date]:
    wanted = python_weekday(number, monday_base)
    days = (dt.date(year, month, k) for k in range(1, month_length(year, month) + 1))
    return [day for day in days if day.weekday() == wanted]


def nth_weekday(year: int, month: int, number: int, occurrence: int,
                monday_base: bool) -> tuple[str, dt.date | None]:
    """('ok', date), ('absent', None) or ('window', None)."""
    days = weekdays_of_month(year, month, number, monday_base)
    if occurrence > len(days):
        return "absent", None
    found = days[occurrence - 1]
    return ("ok", found) if in_window(found) else ("window", None)


def last_weekday(year: int, month: int, number: int,
                 monday_base: bool) -> tuple[str, dt.date | None]:
    """('ok', date) or ('window', None)."""
    found = weekdays_of_month(year, month, number, monday_base)[-1]
    return ("ok", found) if in_window(found) else ("window", None)


def candidates(start: dt.date, end: dt.date, direction: int) -> list[Candidate]:
    """The week anchors 1W to 3W, then the floor and ceiling month anchors, in the window."""
    def passes(anchor: dt.date) -> bool:
        return direction * (anchor - end).days > 0

    found: list[Candidate] = []
    for weeks in (1, 2, 3):
        anchor = add_days(start, direction * 7 * weeks)
        if anchor is not None:
            found.append((False, weeks, anchor))
    floor = 0
    while True:
        anchor = add_months(start, direction * (floor + 1), False)
        if anchor is None or passes(anchor):
            break
        floor += 1
    if floor >= 1:
        floor_anchor = add_months(start, direction * floor, False)
        assert floor_anchor is not None
        found.append((True, floor, floor_anchor))
    ceiling = add_months(start, direction * (floor + 1), False)
    if ceiling is not None:
        found.append((True, floor + 1, ceiling))
    return found


def token(is_month: bool, count: int, sign: str) -> str:
    if not is_month:
        return f"{sign}{count}W"
    years, months = divmod(count, 12)
    if years == 0:
        return f"{sign}{months}M"
    return f"{sign}{years}Y" if months == 0 else f"{sign}{years}Y{months}M"


def format_pillar(start: dt.date, end: dt.date, mode: str) -> tuple[str | None, str | None]:
    """(token, None), or (None, 'RESULT_WINDOW') when no candidate satisfies the mode.

    mode is NEAREST, FLOOR or CEILING."""
    delta = (end - start).days
    direction = 1 if delta >= 0 else -1
    sign = "-" if direction < 0 else ""
    if delta == 0:
        return "0D", None
    if abs(delta) < 7:
        return f"{sign}{abs(delta)}D", None

    def passes(anchor: dt.date) -> bool:
        return direction * (anchor - end).days > 0

    def better_tie(new: Candidate, old: Candidate) -> bool:
        """A month beats a week, and between months the larger count wins."""
        return new[0] and (not old[0] or new[1] > old[1])

    best: Candidate | None = None
    best_key = 0
    for candidate in candidates(start, end, direction):
        anchor = candidate[2]
        reach = direction * (anchor - start).days
        distance = abs((anchor - end).days)
        if mode == "FLOOR":
            if passes(anchor):
                continue
            key = -reach
        elif mode == "CEILING":
            if not (passes(anchor) or distance == 0):
                continue
            key = reach
        else:
            key = distance
        if best is None or key < best_key or (key == best_key and better_tie(candidate, best)):
            best, best_key = candidate, key
    if best is None:
        return None, "RESULT_WINDOW"
    return token(best[0], best[1], sign), None


def main() -> None:
    start = parse("2026-01-01")
    for days in (24, 25, 26, 27):
        print(days, format_pillar(start, start + dt.timedelta(days=days), "NEAREST"))
    january_31 = parse("2026-01-31")
    print("31 January, 25 days", format_pillar(january_31, january_31 + dt.timedelta(days=25), "NEAREST"))
    print("under a month, CEILING", format_pillar(start, parse("2026-01-28"), "CEILING"))
    last_fortnight = parse("9999-12-20"), parse("9999-12-31")
    print("end of 9999", [format_pillar(*last_fortnight, mode) for mode in ("NEAREST", "FLOOR", "CEILING")])


if __name__ == "__main__":
    main()
