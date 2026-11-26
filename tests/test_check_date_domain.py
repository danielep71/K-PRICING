"""tools/check_date_domain.py: each date-domain rule accepts the production source and
rejects a defect made in a copy of it.

The defects are made in memory from the real components, so a rule cannot pass on a
fixture that only resembles the code it guards. A rule that rejects nothing is
caught here, because every test names the message its defect must produce.
"""

from __future__ import annotations

import contextlib
import io
import sys
import unittest
from collections.abc import Mapping
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import check_date_domain  # noqa: E402
from test_check_vba import repository  # noqa: E402

CORE_DATES = "src/core/CORE_Dates.bas"
CORE_PARSE = check_date_domain.CORE_PARSE
FACADE = check_date_domain.FACADE
HOST = check_date_domain.HOST
RANGES = "src/workbook/M_Ranges.bas"
GUARD = "        If Not M_Host.TryCheckHost(Condition) Then"
VOLATILE = "        Application.Volatile True"


def real() -> dict[str, str]:
    return check_date_domain.production(ROOT)


def replaced(path: str, old: str, new: str) -> dict[str, str]:
    """The production source with ``old`` replaced once, which must occur exactly once."""
    sources = real()
    assert sources[path].count(old) == 1, (path, old, sources[path].count(old))
    sources[path] = sources[path].replace(old, new)
    return sources


def inserted(path: str, anchor: str, statement: str, *, before: bool = False) -> dict[str, str]:
    """The production source with a statement line added after (or before) a unique line."""
    return replaced(path, anchor, f"{statement}\n{anchor}" if before else f"{anchor}\n{statement}")


def messages(sources: Mapping[str, str]) -> list[str]:
    return [str(finding) for finding in check_date_domain.analyze(sources)[0]]


class ProductionTest(unittest.TestCase):
    def test_the_production_source_passes_every_rule(self) -> None:
        findings, functions = check_date_domain.analyze(real())
        self.assertEqual([str(finding) for finding in findings], [])
        self.assertEqual(functions, 22)

    def test_text_in_a_comment_or_a_string_is_not_code(self) -> None:
        sources = inserted(CORE_DATES, "        BeginOfMonth = VBA.DateSerial(VBA.Year(DateIn), "
                           "VBA.Month(DateIn), 1)",
                           "    'DateSerial(2026, 1, 0) IsDate(x) Application.Volatile True\n"
                           '        Text = "IsDate(x) DateSerial(y, m, 0) 2958465 ActiveWorkbook"')
        self.assertEqual(messages(sources), [])


class DayZeroTest(unittest.TestCase):
    ANCHOR = "        BeginOfMonth = VBA.DateSerial(VBA.Year(DateIn), VBA.Month(DateIn), 1)"

    def test_a_day_of_zero_is_rejected(self) -> None:
        sources = replaced(CORE_DATES, self.ANCHOR, self.ANCHOR.replace("), 1)", "), 0)"))
        self.assertTrue(any("day of zero" in message for message in messages(sources)))

    def test_the_last_day_taken_from_the_next_month_is_rejected(self) -> None:
        sources = inserted(CORE_DATES, self.ANCHOR,
                           "        Last = VBA.DateSerial(YearPart, MonthPart + 1, 0)")
        self.assertTrue(any("day of zero" in message for message in messages(sources)))

    def test_a_negative_day_is_rejected(self) -> None:
        sources = inserted(CORE_DATES, self.ANCHOR,
                           "        Last = DateSerial(YearPart, MonthPart, -1)")
        self.assertTrue(any("day of zero or below" in message for message in messages(sources)))

    def test_a_call_continued_over_lines_is_read_whole(self) -> None:
        sources = replaced(CORE_DATES, self.ANCHOR,
                           "        BeginOfMonth = VBA.DateSerial(VBA.Year(DateIn), _\n"
                           "            VBA.Month(DateIn), 0)")
        self.assertTrue(any("day of zero" in message for message in messages(sources)))

    def test_a_day_in_any_vba_number_notation_is_read(self) -> None:
        for day in ("0&", "0#", "0%", "0!", "0@", "-1&", "- 1#", "0.0", "0.", ".0", "0E0", "0D0",
                    "&H0", "&O0", "(0)", "((-1))", "+0", "00"):
            with self.subTest(day=day):
                sources = replaced(CORE_DATES, self.ANCHOR,
                                   self.ANCHOR.replace("), 1)", f"), {day})"))
                self.assertTrue(any("day of zero or below" in message
                                    for message in messages(sources)), day)

    def test_a_positive_literal_day_is_accepted(self) -> None:
        for day in ("1&", "28#", "0.5", "&H1", "(1)", "1E0", "(1) + (0)", "Day0", "0 + Day"):
            with self.subTest(day=day):
                sources = replaced(CORE_DATES, self.ANCHOR,
                                   self.ANCHOR.replace("), 1)", f"), {day})"))
                self.assertEqual(messages(sources), [], day)

    def test_a_computed_day_is_not_judged(self) -> None:
        sources = inserted(CORE_DATES, self.ANCHOR,
                           "        Last = DateSerial(YearPart, MonthPart, DaysInMonth(Y, M))")
        self.assertEqual(messages(sources), [])


class WindowTest(unittest.TestCase):
    MIN_SERIAL = "    Public Const WINDOW_MIN_SERIAL  As Double = 61#"

    def test_a_changed_edge_is_rejected(self) -> None:
        sources = replaced(CORE_PARSE, self.MIN_SERIAL, self.MIN_SERIAL.replace("61#", "60#"))
        self.assertTrue(any("WINDOW_MIN_SERIAL is no longer 61" in message
                            for message in messages(sources)))

    def test_every_year_and_serial_is_pinned(self) -> None:
        for name, old, new in (("WINDOW_MAX_SERIAL", "2958465#", "2958466#"),
                               ("WINDOW_MIN_YEAR", "= 1900 ", "= 1901 "),
                               ("WINDOW_MAX_YEAR", "= 9999 ", "= 9998 ")):
            with self.subTest(name=name):
                sources = real()
                lines = sources[CORE_PARSE].split("\n")
                index = next(i for i, line in enumerate(lines) if f"Const {name} " in line)
                lines[index] = lines[index].replace(old, new, 1)
                sources[CORE_PARSE] = "\n".join(lines)
                self.assertTrue(any(f"{name} is no longer" in message
                                    for message in messages(sources)), messages(sources))

    def test_an_initializer_with_more_than_the_edge_is_rejected(self) -> None:
        for name, old, new in (("WINDOW_MIN_SERIAL", "61#", "61# + 1#"),
                               ("WINDOW_MAX_SERIAL", "2958465#", "2958465# - 1#"),
                               ("WINDOW_MIN_YEAR", "= 1900", "= 1900 + 1"),
                               ("WINDOW_MAX_YEAR", "= 9999", "= 9999 - 1")):
            with self.subTest(name=name):
                sources = real()
                lines = sources[CORE_PARSE].split("\n")
                index = next(i for i, line in enumerate(lines) if f"Const {name} " in line)
                code, _, comment = lines[index].partition("'")
                lines[index] = f"{code.rstrip().replace(old, new, 1)} '{comment}"
                sources[CORE_PARSE] = "\n".join(lines)
                self.assertTrue(any(f"{name} is no longer" in message
                                    for message in messages(sources)), messages(sources))

    def test_a_missing_constant_is_rejected(self) -> None:
        sources = replaced(CORE_PARSE, self.MIN_SERIAL, "")
        self.assertTrue(any("WINDOW_MIN_SERIAL is not declared" in message
                            for message in messages(sources)))

    def test_a_constant_declared_again_is_rejected(self) -> None:
        anchor = "        BeginOfMonth = VBA.DateSerial(VBA.Year(DateIn), VBA.Month(DateIn), 1)"
        for path, statement in ((CORE_DATES, "        Const WINDOW_MIN_SERIAL As Double = 0#"),
                                (CORE_DATES, "    Private Const WINDOW_MAX_YEAR As Long = 1"),
                                (RANGES, "    Public Const WINDOW_MIN_YEAR As Long = 1900")):
            with self.subTest(path=path, statement=statement):
                sources = real()
                sources[path] = sources[path].replace("Option Explicit", "Option Explicit\n"
                                                      + statement, 1) if path == RANGES else \
                    replaced(path, anchor, f"{statement}\n{anchor}")[path]
                self.assertTrue(any("is declared again" in message
                                    for message in messages(sources)), messages(sources))

    def test_a_copy_of_an_edge_is_rejected(self) -> None:
        for literal in ("Limit = 2958465#", "Limit = #12/31/9999#", "Limit = #3/1/1900#"):
            with self.subTest(literal=literal):
                sources = inserted(CORE_DATES, "        BeginOfMonth = VBA.DateSerial(VBA.Year"
                                   "(DateIn), VBA.Month(DateIn), 1)", f"        {literal}")
                self.assertTrue(any("edge of the supported window" in message
                                    for message in messages(sources)))


class HostGuardTest(unittest.TestCase):
    CALL = "        KPR_Dates_DaysInMonth = EvaluateCall(dfDaysInMonth, Values)"

    def test_a_function_that_skips_the_evaluator_is_rejected(self) -> None:
        sources = replaced(FACADE, self.CALL, "        KPR_Dates_DaysInMonth = Empty")
        self.assertTrue(any("KPR_Dates_DaysInMonth must call EvaluateCall exactly once" in message
                            for message in messages(sources)))

    def test_the_facade_is_read_without_regard_to_letter_case(self) -> None:
        sources = replaced(FACADE, "Public Function KPR_Dates_DaysInMonth(",
                           "Public Function kpr_dates_daysinmonth(")
        sources[FACADE] = sources[FACADE].replace(self.CALL, "        kpr_dates_daysinmonth = Empty")
        self.assertTrue(any("must call EvaluateCall exactly once" in message
                            for message in messages(sources)), messages(sources))

    def test_a_function_that_calls_the_evaluator_twice_is_rejected(self) -> None:
        sources = inserted(FACADE, self.CALL, "        Spare = EvaluateCall(dfDaysInMonth, Values)")
        self.assertTrue(any("exactly once" in message for message in messages(sources)))

    def test_a_function_that_applies_the_host_rule_itself_is_rejected(self) -> None:
        sources = inserted(FACADE, self.CALL, "        Ok = M_Host.TryCheckHost(Condition)", before=True)
        self.assertTrue(any("must leave the host rule to EvaluateCall" in message
                            for message in messages(sources)))

    def test_an_argument_read_before_the_host_rule_is_rejected(self) -> None:
        sources = inserted(FACADE, GUARD, "        Ok = M_Ranges.TryReadScalar(Values(1), Spare)",
                           before=True)
        self.assertTrue(any("reads an argument before TryCheckHost" in message
                            for message in messages(sources)))

    def test_a_second_host_rule_in_the_evaluator_is_rejected(self) -> None:
        sources = inserted(FACADE, GUARD, "        Ok = M_Host.TryCheckHost(Condition)")
        self.assertTrue(any("EvaluateCall must call TryCheckHost exactly once" in message
                            for message in messages(sources)))

    def test_a_missing_host_rule_is_rejected(self) -> None:
        sources = replaced(FACADE, GUARD, "        If Not True Then")
        self.assertTrue(any("EvaluateCall must call TryCheckHost exactly once" in message
                            for message in messages(sources)))

    def test_the_diagnostic_must_not_apply_the_host_rule(self) -> None:
        sources = inserted(FACADE, "        If Not M_Host.TryReadHostDateSystem(DateSystem, "
                           "Condition) Then", "        Ok = M_Host.TryCheckHost(Condition)",
                           before=True)
        self.assertTrue(any("KPR_Dates_HostDateSystem must call TryReadHostDateSystem exactly "
                            "once" in message for message in messages(sources)))

    def test_the_diagnostic_must_read_the_date_system(self) -> None:
        sources = replaced(FACADE, "        If Not M_Host.TryReadHostDateSystem(DateSystem, "
                           "Condition) Then", "        If Not True Then")
        self.assertTrue(any("KPR_Dates_HostDateSystem must call TryReadHostDateSystem" in message
                            for message in messages(sources)))

    def test_another_component_applying_the_host_rule_is_rejected(self) -> None:
        sources = inserted(RANGES, "Option Explicit", "    '", before=False)
        sources[RANGES] = sources[RANGES].replace("Option Explicit", "Option Explicit\n"
                                                  "Private Sub Stray()\n"
                                                  "        Ok = M_Host.TryCheckHost(Spare)\n"
                                                  "End Sub", 1)
        self.assertTrue(any("applied outside the facade" in message
                            for message in messages(sources)))


class VolatilityTest(unittest.TestCase):
    def test_a_function_that_is_volatile_is_rejected(self) -> None:
        sources = inserted(FACADE, "        KPR_Dates_DaysInMonth = EvaluateCall(dfDaysInMonth, "
                           "Values)", VOLATILE, before=True)
        self.assertTrue(any("KPR_Dates_DaysInMonth must not be volatile" in message
                            for message in messages(sources)))

    def test_a_missing_volatile_call_is_rejected(self) -> None:
        sources = replaced(FACADE, VOLATILE, "        Dummy = 0")
        self.assertTrue(any("must call Application.Volatile exactly once" in message
                            for message in messages(sources)))

    def test_a_volatile_call_that_is_not_first_is_rejected(self) -> None:
        sources = replaced(FACADE, f"{VOLATILE}\n    'Contain an unexpected error\n",
                           "    'Contain an unexpected error\n        Dummy = 0\n"
                           f"{VOLATILE}\n")
        self.assertTrue(any("must be the first statement" in message
                            for message in messages(sources)))

    def test_a_volatile_call_with_false_is_rejected(self) -> None:
        sources = replaced(FACADE, VOLATILE, VOLATILE.replace("True", "False"))
        self.assertTrue(any("must be called with True" in message
                            for message in messages(sources)))

    def test_a_volatile_argument_that_is_more_than_true_is_rejected(self) -> None:
        for argument in ("True And False", "True Or False", "True And Flag", "Not True",
                         "True = False", "Flag", "1"):
            with self.subTest(argument=argument):
                sources = replaced(FACADE, VOLATILE, f"{VOLATILE.split('True')[0]}{argument}")
                self.assertTrue(any("must be called with True" in message
                                    for message in messages(sources)), argument)

    def test_volatile_true_in_parentheses_is_accepted(self) -> None:
        sources = replaced(FACADE, VOLATILE, VOLATILE.replace("True", "(True)"))
        self.assertEqual(messages(sources), [])

    def test_a_second_volatile_call_is_rejected(self) -> None:
        sources = inserted(FACADE, VOLATILE, VOLATILE)
        self.assertTrue(any("exactly once" in message for message in messages(sources)))

    def test_a_volatile_call_in_another_component_is_rejected(self) -> None:
        sources = inserted(HOST, "    'Without a worksheet host the 1900 serial contract applies",
                           VOLATILE, before=True)
        self.assertTrue(any("TryReadHostDateSystem must not be volatile" in message
                            for message in messages(sources)))


class HostAuthorityTest(unittest.TestCase):
    CALLER = "        Set CallerObject = Application.Caller"

    def test_reading_the_caller_elsewhere_is_rejected(self) -> None:
        sources = inserted(RANGES, "Option Explicit", "", before=False)
        sources[RANGES] = sources[RANGES].replace(
            "Option Explicit", "Option Explicit\nPrivate Sub Stray()\n"
            "        Set Spare = Application.Caller\nEnd Sub", 1)
        self.assertTrue(any("Application.Caller is read outside" in message
                            for message in messages(sources)))

    def test_reading_the_caller_twice_in_the_host_is_rejected(self) -> None:
        sources = inserted(HOST, self.CALLER, self.CALLER)
        self.assertTrue(any("Application.Caller must be read exactly once" in message
                            for message in messages(sources)))

    def test_a_workbook_fallback_is_rejected(self) -> None:
        for fallback in ("Set Book = ActiveWorkbook", "Set Book = ThisWorkbook",
                         "Set Sheet = ActiveSheet", "Set Book = Application.ActiveWorkbook"):
            with self.subTest(fallback=fallback):
                sources = inserted(HOST, self.CALLER, f"        {fallback}")
                self.assertTrue(any("consults the active or the containing workbook" in message
                                    for message in messages(sources)))

    def test_registration_may_name_its_own_workbook(self) -> None:
        self.assertEqual(messages(real()), [])
        self.assertIn("ThisWorkbook", real()["src/workbook/M_Register.bas"])


class LocaleTest(unittest.TestCase):
    ANCHOR = "        BeginOfMonth = VBA.DateSerial(VBA.Year(DateIn), VBA.Month(DateIn), 1)"

    def test_a_regional_conversion_is_rejected(self) -> None:
        for call in ("Ok = IsDate(Text)", "D = DateValue(Text)", "D = CVDate(Text)",
                     "Ok = IsNumeric(Text)", "T = Format(D, Spare)", "T = Format$(D, Spare)",
                     "T = FormatDateTime(D)", "Ok = VBA.IsDate(Text)"):
            with self.subTest(call=call):
                sources = inserted(CORE_DATES, self.ANCHOR, f"        {call}")
                self.assertTrue(any("depends on regional settings" in message
                                    for message in messages(sources)))

    def test_cdate_of_text_is_rejected(self) -> None:
        for call in ('D = CDate("2026-01-01")', "D = CDate(Text)", "D = VBA.CDate(DateText)"):
            with self.subTest(call=call):
                sources = inserted(CORE_DATES, self.ANCHOR, f"        {call}")
                self.assertTrue(any("CDate may only convert a serial" in message
                                    for message in messages(sources)))

    def test_cdate_of_a_nested_text_conversion_is_rejected(self) -> None:
        for call in ("D = VBA.CDate(Trim$(CStr(DateIn)))", "D = CDate((Text))",
                     "D = CDate(Left$(Text, 10))"):
            with self.subTest(call=call):
                sources = inserted(CORE_DATES, self.ANCHOR, f"        {call}")
                self.assertTrue(any("CDate may only convert a serial" in message
                                    for message in messages(sources)))

    def test_cdate_of_a_serial_is_accepted(self) -> None:
        sources = inserted(CORE_DATES, self.ANCHOR, "        D = VBA.CDate(ResultSerial)")
        self.assertEqual(messages(sources), [])

    def test_a_longer_name_is_not_a_regional_call(self) -> None:
        sources = inserted(CORE_DATES, self.ANCHOR, "        Ok = IsDateInWindow(DateIn)")
        self.assertEqual(messages(sources), [])


class MainTest(unittest.TestCase):
    def run_main(self, files: Mapping[str, str]) -> tuple[int, str]:
        with repository(dict(files)) as root, contextlib.redirect_stdout(io.StringIO()) as out:
            return check_date_domain.main(["--root", str(root)]), out.getvalue()

    def test_exit_status_follows_the_findings(self) -> None:
        status, output = self.run_main(real())
        self.assertEqual(status, 0, output)
        self.assertIn("date-domain: PASS", output.splitlines()[-1])
        self.assertIn("static check, nothing compiled or run", output.splitlines()[-1])
        status, output = self.run_main(replaced(CORE_PARSE, WindowTest.MIN_SERIAL,
                                                WindowTest.MIN_SERIAL.replace("61#", "60#")))
        self.assertEqual(status, 1)
        self.assertIn("date-domain: FAIL", output.splitlines()[-1])


if __name__ == "__main__":
    unittest.main()
