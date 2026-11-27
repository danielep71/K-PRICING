"""The content of the Dates Demo sheet, which tools/gen_workbook_template.py lays out.

The sheet shows each of the 22 KPR_Dates_* worksheet functions on inputs the
reader can edit, beside a formula of Excel's own functions that should give
the same result, with a status per row. A side panel counts the statuses,
lists inputs that must give an Excel error value, spills one formula over
many dates and shows the array limits.

Formulas are written as Excel shows them, starting with "=". A name such as
DateIn is one of INPUTS. {IN} in a case stands for the case's input cell, and
{SRC} in a spill column for the source spill. A row's status is OK when the
function's result equals the reference, or both are the same error value;
DIFFERS when they differ while the row's condition is TRUE, a documented
difference by design; FAIL otherwise; NO VBA while the functions are not
loaded (#NAME?). Standard library only.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass

# A cell value: a number, a Boolean, text, a date, an Excel error value
# ("#VALUE!", written as ErrorValue), or None for an empty cell.


@dataclass(frozen=True)
class ErrorValue:
    """An Excel error value, such as #VALUE!."""

    text: str


Value = float | int | bool | str | datetime.date | ErrorValue | None


@dataclass(frozen=True)
class Input:
    label: str
    name: str
    value: Value
    note: str


@dataclass(frozen=True)
class Compare:
    formula: str
    reference: str
    note: str
    date: bool = False
    differs_when: str = ""


@dataclass(frozen=True)
class Section:
    title: str
    rows: tuple[Compare, ...]


@dataclass(frozen=True)
class Case:
    value: Value
    formula: str
    expected: Value
    rule: str


@dataclass(frozen=True)
class CaseTable:
    title: str
    cases: tuple[Case, ...]


@dataclass(frozen=True)
class SpillColumn:
    header: str
    formula: str
    reference: str
    date: bool = False


@dataclass(frozen=True)
class Spill:
    title: str
    header: str
    formula: str
    rows: int
    columns: tuple[SpillColumn, ...]


TITLE = "Dates Demo"
SUBTITLE = ("The 22 date functions beside Excel's own formulas. Edit the orange cells; every "
            "status recalculates.")
INPUTS_TITLE = "Inputs (edit these)"
CHECKS_TITLE = "Checks"

INPUTS = (
    Input("DateIn", "DateIn", datetime.date(2024, 4, 30), "Any date from 1900-03-01 to 9999-12-31"),
    Input("nDays", "nDays", 2, "Whole number of days"),
    Input("nWeeks", "nWeeks", 3, "Whole number of weeks"),
    Input("nMonths", "nMonths", 1, "Whole number of months"),
    Input("nYears", "nYears", 2, "Whole number of years"),
    Input("Keep EOM", "KeepEOM", True, "TRUE keeps a month end at month end"),
    Input("Weekday index", "WdIndex", 4, "1 to 7 in the weekday base below"),
    Input("Occurrence n", "nOccur", 2, "1 to 5"),
    Input("WeekBase Monday", "WeekBaseMonday", True, "TRUE: Monday = 1; FALSE: Sunday = 1"),
)

# Formula pieces the references share.
LEAP = "OR(MOD(YEAR(DateIn),400)=0,AND(MOD(YEAR(DateIn),4)=0,MOD(YEAR(DateIn),100)<>0))"
FIRST_OF_MONTH = "DATE(YEAR(DateIn),MONTH(DateIn),1)"
WEEKDAY_MODE = "IF(WeekBaseMonday,2,1)"
KEEP_EOM_APPLIES = "AND(KeepEOM,DateIn=EOMONTH(DateIn,0))"
NTH = (f"{FIRST_OF_MONTH}+MOD(WdIndex-WEEKDAY({FIRST_OF_MONTH},{WEEKDAY_MODE}),7)"
       "+7*(nOccur-1)")
LAST = f"EOMONTH(DateIn,0)-MOD(WEEKDAY(EOMONTH(DateIn,0),{WEEKDAY_MODE})-WdIndex,7)"

SECTIONS = (
    Section("Day primitives", (
        Compare("=KPR_Dates_DayOfWeek(DateIn,WeekBaseMonday)", f"=WEEKDAY(DateIn,{WEEKDAY_MODE})",
                "Monday = 1 by default; WeekBaseMonday = FALSE matches WEEKDAY's Sunday = 1."),
        Compare("=KPR_Dates_DaysInMonth(DateIn)", "=DAY(EOMONTH(DateIn,0))",
                "Gregorian month length."),
        Compare("=KPR_Dates_DaysInYear(YEAR(DateIn))", f"=IF({LEAP},366,365)",
                "Takes a year, not a date."),
    )),
    Section("Period boundaries", (
        Compare("=KPR_Dates_BeginOfMonth(DateIn)", "=EOMONTH(DateIn,-1)+1",
                "First day of the month.", True),
        Compare("=KPR_Dates_EndOfMonth(DateIn)", "=EOMONTH(DateIn,0)",
                "Last day of the month.", True),
        Compare("=KPR_Dates_BeginOfQuarter(DateIn)",
                "=DATE(YEAR(DateIn),3*INT((MONTH(DateIn)-1)/3)+1,1)",
                "Calendar quarters start in January, April, July and October.", True),
        Compare("=KPR_Dates_EndOfQuarter(DateIn)",
                "=EOMONTH(DATE(YEAR(DateIn),3*INT((MONTH(DateIn)-1)/3)+3,1),0)",
                "Last day of the calendar quarter.", True),
        Compare("=KPR_Dates_BeginOfYear(DateIn)", "=DATE(YEAR(DateIn),1,1)", "1 January.", True),
        Compare("=KPR_Dates_EndOfYear(DateIn)", "=DATE(YEAR(DateIn),12,31)", "31 December.",
                True),
    )),
    Section("Predicates", (
        Compare("=KPR_Dates_IsMonthEnd(DateIn)", "=DateIn=EOMONTH(DateIn,0)",
                "TRUE on the last day of a month."),
        Compare("=KPR_Dates_IsQuarterEnd(DateIn)",
                "=AND(DateIn=EOMONTH(DateIn,0),MOD(MONTH(DateIn),3)=0)",
                "TRUE on 31 March, 30 June, 30 September and 31 December."),
        Compare("=KPR_Dates_IsYearEnd(DateIn)", "=DateIn=DATE(YEAR(DateIn),12,31)",
                "TRUE on 31 December."),
        Compare("=KPR_Dates_IsLeapYear(YEAR(DateIn))", f"={LEAP}",
                "Takes a year. 1900 is not a leap year, whatever Excel's DATE says."),
    )),
    Section("Date arithmetic", (
        Compare("=KPR_Dates_AddDays(DateIn,nDays)", "=DateIn+nDays", "Exact calendar days.",
                True),
        Compare("=KPR_Dates_AddWeeks(DateIn,nWeeks)", "=DateIn+7*nWeeks",
                "Weeks of seven days.", True),
        Compare("=KPR_Dates_AddMonths(DateIn,nMonths,KeepEOM)",
                f"=IF({KEEP_EOM_APPLIES},EOMONTH(DateIn,nMonths),EDATE(DateIn,nMonths))",
                "Clips to a shorter month end; Keep EOM moves a month end to the target month "
                "end.", True),
        Compare("=KPR_Dates_AddYears(DateIn,nYears,KeepEOM)",
                f"=IF({KEEP_EOM_APPLIES},EOMONTH(DateIn,12*nYears),EDATE(DateIn,12*nYears))",
                "Twelve months per year: 29 February clips to 28 February.", True),
    )),
    Section("Month-end policy against EDATE and EOMONTH", (
        Compare("=KPR_Dates_AddMonths(DateIn,nMonths,KeepEOM)", "=EDATE(DateIn,nMonths)",
                "EDATE always clips. DIFFERS when Keep EOM moves a month end further.", True,
                KEEP_EOM_APPLIES),
        Compare("=KPR_Dates_AddMonths(DateIn,nMonths,KeepEOM)", "=EOMONTH(DateIn,nMonths)",
                "EOMONTH always lands on a month end. Equal only when Keep EOM applies.", True,
                f"NOT({KEEP_EOM_APPLIES})"),
    )),
    Section("Weekday locators", (
        Compare("=KPR_Dates_NthWeekdayOfMonth(YEAR(DateIn),MONTH(DateIn),WdIndex,nOccur,"
                "WeekBaseMonday)",
                f"=IF(MONTH({NTH})=MONTH(DateIn),{NTH},SQRT(-1))",
                "The n-th weekday of the month of DateIn; #NUM! when the month has no such "
                "occurrence.", True),
        Compare("=KPR_Dates_LastWeekdayOfMonth(YEAR(DateIn),MONTH(DateIn),WdIndex,"
                "WeekBaseMonday)", f"={LAST}",
                "The last such weekday of the month of DateIn.", True),
    )),
    Section("Pillars", (
        Compare('=KPR_Dates_DateFromPillar(DateIn,"ON")', "=DateIn+1",
                "Overnight: one calendar day.", True),
        Compare('=KPR_Dates_DateFromPillar(DateIn,"TN")', "=DateIn+2",
                "Tom-next: two calendar days. Business days arrive with calendars.", True),
        Compare('=KPR_Dates_DateFromPillar(DateIn,"1W")', "=DateIn+7", "One week.", True),
        Compare('=KPR_Dates_DateFromPillar(DateIn,"3M")', "=EDATE(DateIn,3)",
                "Months clip like EDATE.", True),
        Compare('=KPR_Dates_DateFromPillar(DateIn,"1Y6M")', "=EDATE(DateIn,18)",
                "Years and months combine.", True),
        Compare('=KPR_Dates_DateFromPillar(DateIn,"-1D")', "=DateIn-1",
                "A sign applies to the whole token.", True),
        Compare("=KPR_Dates_PillarFromDates(DateIn,DateIn+7)", '="1W"',
                "Seven days format as one week."),
        Compare("=KPR_Dates_PillarFromDates(DateIn,EDATE(DateIn,3))", '="3M"',
                "An exact month count formats as months."),
        Compare("=KPR_Dates_PillarFromDates(DateIn,EDATE(DateIn,18))", '="1Y6M"',
                "Twelve months or more format as years and months."),
        Compare("=KPR_Dates_PillarFromDates(DateIn,DateIn-7)", '="-1W"',
                "An end before the start gives a minus sign."),
        Compare('=KPR_Dates_PillarFromDates("2024-01-15","2024-02-10","NEAREST")', '="1M"',
                "Rounding NEAREST: 26 days are nearer one month than three weeks."),
        Compare('=KPR_Dates_PillarFromDates("2024-01-15","2024-02-10","FLOOR")', '="3W"',
                "Rounding FLOOR: the largest pillar not after the end."),
        Compare('=KPR_Dates_PillarFromDates("2024-01-15","2024-02-10","CEILING")', '="1M"',
                "Rounding CEILING: the smallest pillar not before the end."),
    )),
)

NOTES = (
    "Dates are passed as cells or as YYYY-MM-DD text, so nothing here depends on the regional "
    "date format.",
    "In a workbook on the 1904 date system every value-taking function returns #N/A instead of "
    "a shifted date.",
)

VALUE, NUM, DIV0 = ErrorValue("#VALUE!"), ErrorValue("#NUM!"), ErrorValue("#DIV/0!")

CASES = CaseTable("Invalid input gives an Excel error value", (
    Case("aa", "=KPR_Dates_DaysInMonth({IN})", VALUE,
         "DaysInMonth: text that is not YYYY-MM-DD gives #VALUE!."),
    Case("32/04/2020", "=KPR_Dates_BeginOfMonth({IN})", VALUE,
         "BeginOfMonth: regional date text is never guessed."),
    Case('"2024-04-30"', "=KPR_Dates_IsMonthEnd({IN})", VALUE,
         "IsMonthEnd: quotes inside the text make it invalid."),
    Case(True, "=KPR_Dates_DaysInYear({IN})", VALUE, "DaysInYear: TRUE is not a year."),
    Case(1899, "=KPR_Dates_IsLeapYear({IN})", VALUE, "IsLeapYear: years run from 1900 to 9999."),
    Case(-4, "=KPR_Dates_DayOfWeek({IN})", NUM,
         "DayOfWeek: a serial before 1900-03-01 gives #NUM!."),
    Case(0, "=KPR_Dates_EndOfMonth({IN})", NUM,
         "EndOfMonth: serial 0 is outside the supported window."),
    Case(32, "=KPR_Dates_IsQuarterEnd({IN})", NUM,
         "IsQuarterEnd: 1900-02-01 is before the window."),
    Case("9999-12-31", "=KPR_Dates_AddDays({IN},1)", NUM,
         "AddDays: a result after 9999-12-31 gives #NUM!."),
    Case(1.5, '=KPR_Dates_AddDays("2024-01-31",{IN})', VALUE,
         "AddDays: a count is never truncated or rounded."),
    Case("yes", '=KPR_Dates_AddMonths("2024-01-31",1,{IN})', VALUE,
         "AddMonths: Keep EOM takes TRUE or FALSE only."),
    Case(5, "=KPR_Dates_NthWeekdayOfMonth(2026,2,1,{IN})", NUM,
         "NthWeekdayOfMonth: February 2026 has four Mondays."),
    Case("-ON", '=KPR_Dates_DateFromPillar("2024-01-31",{IN})', VALUE,
         "DateFromPillar: an alias never carries a sign."),
    Case("1M2M", '=KPR_Dates_DateFromPillar("2024-01-31",{IN})', VALUE,
         "DateFromPillar: a unit may appear once."),
    Case("ROUND", '=KPR_Dates_PillarFromDates("2024-01-15","2024-04-15",{IN})', VALUE,
         "PillarFromDates: rounding is NEAREST, FLOOR or CEILING."),
    Case(DIV0, "=KPR_Dates_EndOfMonth({IN})", DIV0,
         "EndOfMonth: an error in the input passes through unchanged."),
    Case(None, "=KPR_Dates_HostDateSystem()", 1900,
         "HostDateSystem: this workbook uses the 1900 date system."),
))

# EOMONTH refuses a spill or range reference as an array, so +0 turns it into values.
SPILL = Spill("One formula, many dates", "Date", "=SEQUENCE(30,1,DateIn,25)", 30, (
    SpillColumn("Days in month", "=KPR_Dates_DaysInMonth({SRC})", "DAY(EOMONTH({SRC}+0,0))"),
    SpillColumn("Begin of month", "=KPR_Dates_BeginOfMonth({SRC})", "EOMONTH({SRC}+0,-1)+1",
                True),
    SpillColumn("Is leap year", "=KPR_Dates_IsLeapYear(YEAR({SRC}))",
                "((MOD(YEAR({SRC}),400)=0)+(MOD(YEAR({SRC}),4)=0)*(MOD(YEAR({SRC}),100)<>0)>0)"),
))

LIMITS = CaseTable("Array limits", (
    Case("100,000 dates", "=ROWS(KPR_Dates_AddDays(SEQUENCE(100000,1,45000),0))", 100000,
         "Up to 100,000 elements per call are accepted."),
    Case("100,001 dates", "=KPR_Dates_AddDays(SEQUENCE(100001,1,45000),0)", NUM,
         "One element more gives a single #NUM!."),
    Case("2x1 and 1x3", '=KPR_Dates_AddDays({"2024-01-01";"2024-01-02"},{1,2,3})', VALUE,
         "Array arguments of different shapes give #VALUE!."),
))
