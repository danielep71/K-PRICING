# 🧮 Calculation contracts

A contract states what one calculation must produce, from which inputs, under
which rules. The VBA implements the contract, the
[validation plan](VALIDATION_PLAN.md) tests against it, and a change to a
result it defines is a `behavior-change`.

## Common rules

These rules hold for every date-layer contract, `CALC-02` to `CALC-10`. A
rule that cites `Decision (#4)` is a design decision, recorded under
[Decisions](#decisions); the [data contract](DATA_CONTRACT.md#function-arguments)
says how each argument is read and which error each rejection returns.

- **Units and signs:** no amount, currency or rate. A date is a day of the
  proleptic Gregorian calendar ([SRC-02](SOURCES.md#register)): on the
  worksheet an Excel serial under the 1900 date system
  ([SRC-03](SOURCES.md#register)), in VBA a `Date`
  ([SRC-11](SOURCES.md#register)). From serial 61, 1900-03-01, the two
  numberings name the same day; below it they differ by one, because Excel
  counts a 1900-02-29 that never existed ([SRC-04](SOURCES.md#register)).
  Days, weeks, months, years, weekdays and occurrences are whole numbers
  (`Long`, [SRC-12](SOURCES.md#register)). A shift is signed: a positive one
  moves forward in time.
- **Dates:** the supported window is 1900-03-01 to 9999-12-31, both included
  (Decision (#4)); 9999-12-31 is the latest date Excel calculates
  ([SRC-08](SOURCES.md#register)) and VBA holds
  ([SRC-11](SOURCES.md#register)). Every accepted date and every date result
  is a whole day, with no time. A date result outside the window is never
  returned: it is `RESULT_WINDOW`. No date is built with a day of zero or
  below, as `DateSerial(year, month + 1, 0)` for a month's last day: after
  December 9999 that overflows the date range, so a month's last day comes
  from the month's length (`CALC-03` step 2). No calendar, holiday, business
  day, day count or roll convention applies: those are later milestones.
- **Precision and rounding:** whole-day arithmetic, with no rounding. The
  only exceptions are stated where they apply: the time of an accepted date
  is removed, a whole-number argument within `1E-9` of a whole number is read
  as that number ([data contract](DATA_CONTRACT.md#integer-values)), and
  `CALC-09` rounds an interval to a pillar.
- **Currency conversion:** none.
- **Failure:** a function returns an Excel error, never raises one to the
  worksheet, shows a dialog or changes Excel's state
  ([error values](DATA_CONTRACT.md#error-values)). The functions write no
  cell, so there are no earlier outputs to withdraw: Excel recalculates the
  result. Each contract's *Failure behaviour* names one `_status` quantity,
  whose texts are the [condition](DATA_CONTRACT.md#conditions) identifiers:
  those of its arguments' parsers, those of the rules below, and its own.
- **Tolerances:** every comparison is exact, with `absolute` and `relative`
  both 0 ([tolerance](TEST_CASES.md#tolerance)): a date is compared as its
  serial, a whole number, a Boolean and a text as themselves (Decision (#4)).

<a id="host-and-date-system"></a>

### Host and date system

Each public call reads `Application.Caller` once, at its start, before it
reads any argument ([SRC-05](SOURCES.md#register)). Only a worksheet caller
decides the date system (Decision (#4)):

1. When `Application.Caller` is a worksheet `Range`, the call reads
   `Date1904` of that Range's own workbook ([SRC-06](SOURCES.md#register)).
2. When that workbook uses the 1900 date system, the call evaluates normally.
3. When it uses the 1904 date system, whose serials are 1,462 lower for the
   same day ([SRC-03](SOURCES.md#register)), every function with a value
   argument returns one call-level `#N/A`, `HOST_DATE1904`: never a value
   shifted by 1,462 days.
4. When a worksheet `Range` was identified but its workbook's date system
   cannot be read, the call returns one call-level `#N/A`,
   `HOST_UNRESOLVED`. Only that sequence gives it: a failure while reading or
   classifying `Application.Caller` itself means that no worksheet host was
   identified, and step 5 applies.
5. When no worksheet host is identified, the 1900 serial contract applies.
   No other workbook, such as `ActiveWorkbook` or `ThisWorkbook`, is ever
   consulted.

Direct VBA calls, the Immediate window, `Application.Run` and the test
harness are the supported callers without a worksheet host; the caller owns
the meaning of the values it builds, and a Range passed as an argument from
VBA does not make the call a worksheet call. A caller that is not a Range is
not proof of direct VBA: a macro attached to a shape can give a text caller,
and data validation, a chart series or a defined name can give an error or
another form. Those contexts are probed in Excel
([host behaviour](VALIDATION_PLAN.md#host-behaviour)) and are not supported.

Only `KPR_Dates_HostDateSystem` is volatile (`CALC-10`); the other 21
functions never call `Application.Volatile`.

<a id="arrays-and-evaluation"></a>

### Arrays and evaluation

Every function except `KPR_Dates_HostDateSystem` accepts arrays at its value
arguments (Decision (#4)):

- A scalar, a one-cell Range or a one-element array is a scalar.
- A one-dimensional array is read as one row of N columns.
- A two-dimensional array or a single-area Range keeps its rows, columns and
  orientation. A Range is read at its full size, never trimmed to the used
  range.
- A scalar value argument is repeated over the shape of the non-scalar ones.
  Several non-scalar value arguments must have the same rows and columns
  (`SHAPE_MISMATCH`). There is no outer product, broadcasting of a row against
  a column, jagged traversal, multi-area Range or third dimension
  (`SHAPE_UNSUPPORTED`).
- A call whose arguments are all scalar returns a scalar, not a 1 by 1 array.
  Any other call returns a two-dimensional array, even of one row or one
  column.

Each call runs these stages once:

1. read the caller and apply the [host rule](#host-and-date-system);
2. classify the shape of every value and control argument from its type and
   dimensions alone, reading no content, and reject an unsupported one;
3. validate the controls and resolve the result's shape;
4. reject a result of 100,001 elements or more (`CAPACITY_EXCEEDED`), from
   the dimensions alone, before any Range is read; 100,000 is allowed;
5. evaluate the elements.

An array whose elements are themselves arrays can be found only by reading
its elements, which stage 4 must precede, so it is reported as
`SHAPE_UNSUPPORTED` while the elements are read, not at stage 2.

Elements are evaluated by rows: every column of the first row, then the
second. At each position the value arguments are read in signature order,
and the first that fails decides that element, whether it holds an Excel
error or breaks a rule. A failed element never stops its neighbours. A
call-level failure comes before every element, because no element is then
evaluated.

<a id="excel-versions"></a>

### Excel versions

A single-cell call is claimed only for the Excel versions on which a
certification run passed. A multi-cell result is claimed only for Excel with
dynamic arrays, which spills it ([SRC-10](SOURCES.md#register)); no
Ctrl+Shift+Enter array formula is claimed. The functions detect no Excel
version and never return `#SPILL!`: placing the result is Excel's
(Decision (#4)).

## Contract template

Copy this section for each calculation.

### *Illustrative:* `CALC-01` Example calculation

**Purpose:** what the figure is, and where it is reported.

**Inputs:**

| Input | Dataset and field | Unit |
| --- | --- | --- |
| Example amount | `DS-01.notional` | Currency of `DS-01.currency` |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | Example: sum the notional amounts | `SRC-01`, Article 1(1) |

- **Step:** a whole number, once per step. A test case's `rule` and a
  traceability row name step 2 as `CALC-NN.2`, and
  `tools/check_test_cases.py` rejects a step that this table does not list.

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Example total | `illustrative_total_notional_amount` | Reporting currency | Whole calculation | None |

- **Quantity:** the lower snake case name that the test cases' expected
  files and the accuracy contracts use for the output
  ([expected-result format](TEST_CASES.md#expected-format)). A `_status`
  quantity, which reports a rejection, is not listed here: it belongs to
  *Failure behaviour*.

**Limits and non-goals:** cases the calculation does not cover, and what
happens if they occur in the input.

**Failure behaviour:** which conditions stop it, and what it reports: each
`_status` quantity, in lower snake case ending in `_status`, with the exact
texts it can hold, which the `invalid` test cases expect
([invalid inputs](TEST_CASES.md#invalid-inputs)).

**Validation:** the checks in the [validation plan](VALIDATION_PLAN.md) that
cover it, and the rows of its steps in its
[traceability table](VALIDATION_PLAN.md#traceability).

<a id="date-layer-contracts"></a>

## Date-layer contracts

The 22 `KPR_Dates_*` worksheet functions, by contract. Each returns a
`Variant`: its result, or an Excel error. The supported API is exactly these
22 names.

| Function | Arguments | Result | Contract |
| --- | --- | --- | --- |
| `KPR_Dates_DayOfWeek` | `DateIn`, `[Opt_WeekBaseMonday=TRUE]` | Whole number 1 to 7 | `CALC-02` |
| `KPR_Dates_DaysInMonth` | `DateIn` | Whole number 28 to 31 | `CALC-03` |
| `KPR_Dates_DaysInYear` | `YearIn` | 365 or 366 | `CALC-03` |
| `KPR_Dates_IsLeapYear` | `YearIn` | Boolean | `CALC-03` |
| `KPR_Dates_BeginOfMonth` | `DateIn` | Date | `CALC-04` |
| `KPR_Dates_EndOfMonth` | `DateIn` | Date | `CALC-04` |
| `KPR_Dates_BeginOfQuarter` | `DateIn` | Date | `CALC-04` |
| `KPR_Dates_EndOfQuarter` | `DateIn` | Date | `CALC-04` |
| `KPR_Dates_BeginOfYear` | `DateIn` | Date | `CALC-04` |
| `KPR_Dates_EndOfYear` | `DateIn` | Date | `CALC-04` |
| `KPR_Dates_IsMonthEnd` | `DateIn` | Boolean | `CALC-05` |
| `KPR_Dates_IsQuarterEnd` | `DateIn` | Boolean | `CALC-05` |
| `KPR_Dates_IsYearEnd` | `DateIn` | Boolean | `CALC-05` |
| `KPR_Dates_AddDays` | `DateIn`, `nDays` | Date | `CALC-06` |
| `KPR_Dates_AddWeeks` | `DateIn`, `nWeeks` | Date | `CALC-06` |
| `KPR_Dates_AddMonths` | `DateIn`, `nMonths`, `[Opt_KeepEOM=FALSE]` | Date | `CALC-06` |
| `KPR_Dates_AddYears` | `DateIn`, `nYears`, `[Opt_KeepEOM=FALSE]` | Date | `CALC-06` |
| `KPR_Dates_NthWeekdayOfMonth` | `YearIn`, `MonthIn`, `WdIndex`, `n`, `[Opt_WeekBaseMonday=TRUE]` | Date | `CALC-07` |
| `KPR_Dates_LastWeekdayOfMonth` | `YearIn`, `MonthIn`, `WdIndex`, `[Opt_WeekBaseMonday=TRUE]` | Date | `CALC-07` |
| `KPR_Dates_DateFromPillar` | `StartDate`, `Pillar` | Date | `CALC-08` |
| `KPR_Dates_PillarFromDates` | `StartDate`, `EndDate`, `[Opt_Rounding="NEAREST"]` | Text | `CALC-09` |
| `KPR_Dates_HostDateSystem` | None | 1900 or 1904 | `CALC-10` |

Arguments in brackets are optional, with their default. Every function except
`KPR_Dates_HostDateSystem` accepts arrays at its value arguments
([arrays and evaluation](#arrays-and-evaluation)).

### `CALC-02` Day of week

**Purpose:** the number of a date's weekday, Monday-based by default.

**Inputs:**

| Input | Argument and reading | Unit |
| --- | --- | --- |
| Date | `DateIn`, a [date value](DATA_CONTRACT.md#date-values) | Day |
| Weekday base | `Opt_WeekBaseMonday`, a [control](DATA_CONTRACT.md#control-values), default `TRUE` | — |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | With `Opt_WeekBaseMonday` `TRUE`, Monday is 1 and Sunday is 7 | `SRC-02`, day of the week, Monday as day 1 |
| 2 | With `FALSE`, Sunday is 1 and Saturday is 7, as `WEEKDAY(date)` with its default return type | `SRC-09`, return type 1 |
| 3 | The default is Monday-based, the reverse of `WEEKDAY`'s default; the same base numbers `WdIndex` in `CALC-07` | Decision (#4) |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Weekday number | `weekday_number` | Whole number, 1 to 7 | Each element | None |

**Limits and non-goals:** no other weekday base, such as `WEEKDAY`'s return
types 3 and 11 to 17; a control that is not a Boolean is
`CONTROL_TYPE_REJECTED`.

**Failure behaviour:** `day_of_week_status` holds the condition: the
[common](#common-rules) and [date value](DATA_CONTRACT.md#date-values)
conditions, and the control conditions.

**Validation:** the [validation plan](VALIDATION_PLAN.md#date-layer-coverage);
traceability rows follow the implementation.

### `CALC-03` Month and year lengths

**Purpose:** the number of days in a date's month or in a year, and whether a
year is a leap year.

**Inputs:**

| Input | Argument and reading | Unit |
| --- | --- | --- |
| Date | `DateIn` of `KPR_Dates_DaysInMonth`, a [date value](DATA_CONTRACT.md#date-values) | Day |
| Year | `YearIn` of `KPR_Dates_DaysInYear` and `KPR_Dates_IsLeapYear`, an [integer value](DATA_CONTRACT.md#integer-values), 1900 to 9999 | Calendar year |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | A year is a leap year when it is divisible by 4, except a year divisible by 100 and not by 400 | `SRC-02`, Gregorian calendar, leap year |
| 2 | `KPR_Dates_DaysInMonth` returns the length of the date's month: 31 for January, March, May, July, August, October and December; 30 for April, June, September and November; for February 29 in a leap year, 28 otherwise | `SRC-02`, Gregorian calendar |
| 3 | `KPR_Dates_DaysInYear` returns 366 for a leap year and 365 otherwise | `SRC-02`, Gregorian calendar, leap year |
| 4 | `KPR_Dates_IsLeapYear` returns `TRUE` for a leap year and `FALSE` otherwise | `SRC-02`, Gregorian calendar, leap year |
| 5 | The two year functions take a calendar year, never a date, and build no date, so the window does not apply: `KPR_Dates_DaysInYear(1900)` is 365 and `KPR_Dates_IsLeapYear(1900)` is `FALSE` | Decision (#4) |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Days in the month | `days_in_month` | Days, 28 to 31 | Each element | None |
| Days in the year | `days_in_year` | Days, 365 or 366 | Each element | None |
| Leap year | `leap_year_flag` | Boolean | Each element | None |

**Limits and non-goals:** each function takes the smallest input that
decides its answer: a month needs its year, so `KPR_Dates_DaysInMonth` takes a
date; a year takes a whole number.

**Failure behaviour:** `month_year_length_status` holds the condition: the
date value or integer value conditions of its argument and `DOMAIN_YEAR`.
`DATE_WINDOW` and `RESULT_WINDOW` cannot arise for the year functions.

**Validation:** the [validation plan](VALIDATION_PLAN.md#date-layer-coverage).

### `CALC-04` Period boundaries

**Purpose:** the first and last day of a date's month, quarter and year.

**Inputs:**

| Input | Argument and reading | Unit |
| --- | --- | --- |
| Date | `DateIn`, a [date value](DATA_CONTRACT.md#date-values) | Day |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | Quarters are the calendar quarters January to March, April to June, July to September and October to December | Decision (#4) |
| 2 | `BeginOfMonth`, `BeginOfQuarter` and `BeginOfYear` return the first day of the month, quarter or year that holds the date | Decision (#4) |
| 3 | `EndOfMonth`, `EndOfQuarter` and `EndOfYear` return its last day, with month lengths as in `CALC-03` step 2 | `SRC-02`, Gregorian calendar |
| 4 | A boundary is a result and leaves the window as one: `BeginOfYear` of any 1900 date and `BeginOfQuarter` of a date in the first quarter of 1900 name 1900-01-01 and return `#NUM!`, `RESULT_WINDOW`; `EndOfQuarter` and `EndOfYear` of the same dates return 1900-03-31 and 1900-12-31, and `BeginOfQuarter` of a second-quarter 1900 date returns 1900-04-01 | Decision (#4) |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| First day of the period | `period_begin_date` | Date | Each element | None |
| Last day of the period | `period_end_date` | Date | Each element | None |

**Limits and non-goals:** no fiscal year, no quarter starting in another
month, no week boundary.

**Failure behaviour:** `period_boundary_status` holds the condition: the date
value conditions and `RESULT_WINDOW`.

**Validation:** the [validation plan](VALIDATION_PLAN.md#date-layer-coverage).

### `CALC-05` Period-end tests

**Purpose:** whether a date is the last day of its month, quarter or year.

**Inputs:**

| Input | Argument and reading | Unit |
| --- | --- | --- |
| Date | `DateIn`, a [date value](DATA_CONTRACT.md#date-values) | Day |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | `IsMonthEnd`, `IsQuarterEnd` and `IsYearEnd` return `TRUE` when the date is the last day of its month, quarter or year as `CALC-04` defines it, and `FALSE` otherwise | Decision (#4) |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Period end | `period_end_flag` | Boolean | Each element | None |

**Limits and non-goals:** no test for the first day of a period: compare with
`CALC-04`'s result.

**Failure behaviour:** `period_end_status` holds the condition: the date
value conditions.

**Validation:** the [validation plan](VALIDATION_PLAN.md#date-layer-coverage).

### `CALC-06` Date shifts

**Purpose:** a date moved by a signed number of days, weeks, months or years.

**Inputs:**

| Input | Argument and reading | Unit |
| --- | --- | --- |
| Date | `DateIn`, a [date value](DATA_CONTRACT.md#date-values) | Day |
| Shift | `nDays`, `nWeeks`, `nMonths` or `nYears`, an [integer value](DATA_CONTRACT.md#integer-values) | Days, weeks, months or years |
| Keep month end | `Opt_KeepEOM` of `AddMonths` and `AddYears`, a [control](DATA_CONTRACT.md#control-values), default `FALSE` | — |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | `AddDays` adds the shift in calendar days | Decision (#4) |
| 2 | `AddWeeks` adds seven times the shift in calendar days | Decision (#4) |
| 3 | `AddMonths` moves the date's month by the shift and keeps its day when the target month has it; otherwise it returns the target month's last day (clip) | Decision (#4) |
| 4 | With `Opt_KeepEOM` `TRUE`, a date that is its month's last day moves to the target month's last day; any other date follows step 3 | Decision (#4) |
| 5 | `AddYears` is `AddMonths` with twelve times the shift, so clipping, 29 February and `Opt_KeepEOM` behave the same | Decision (#4) |
| 6 | Every intermediate value and the result are checked against the window before a VBA `Date` is built; an overflow or a result outside the window is `#NUM!`, `RESULT_WINDOW`, never a runtime error or a rollover | Decision (#4) |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Shifted date | `shifted_date` | Date | Each element | None |

**Limits and non-goals:** calendar days only: no business day, holiday or
roll convention. `AddMonths(AddMonths(d, n), -n)` need not return `d`, because
step 3 clips.

**Failure behaviour:** `date_shift_status` holds the condition: the date and
integer value conditions, the control conditions and `RESULT_WINDOW`.

**Validation:** the [validation plan](VALIDATION_PLAN.md#date-layer-coverage).

### `CALC-07` Weekday locators

**Purpose:** the date of the nth, or the last, given weekday of a month.

**Inputs:**

| Input | Argument and reading | Unit |
| --- | --- | --- |
| Year | `YearIn`, an [integer value](DATA_CONTRACT.md#integer-values), 1900 to 9999 | Calendar year |
| Month | `MonthIn`, an integer value, 1 to 12 | Month |
| Weekday | `WdIndex`, an integer value, 1 to 7, in the weekday base | Weekday number |
| Occurrence | `n` of `NthWeekdayOfMonth`, an integer value, 1 to 5 | Occurrence |
| Weekday base | `Opt_WeekBaseMonday`, a [control](DATA_CONTRACT.md#control-values), default `TRUE` | — |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | `WdIndex` is numbered as `CALC-02` numbers weekdays under the same base | Decision (#4) |
| 2 | `NthWeekdayOfMonth` returns the date of occurrence `n` of that weekday in the month | Decision (#4) |
| 3 | A valid request for an occurrence the month does not have, such as a fifth weekday it lacks, is `#NUM!`, `OCCURRENCE_ABSENT` | Decision (#4) |
| 4 | `LastWeekdayOfMonth` returns the date of the last occurrence of that weekday in the month | Decision (#4) |
| 5 | A result before 1900-03-01 is `#NUM!`, `RESULT_WINDOW` | Decision (#4) |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Weekday's date | `weekday_date` | Date | Each element | None |

**Limits and non-goals:** no occurrence counted from the end other than the
last, and no occurrence beyond the fifth.

**Failure behaviour:** `weekday_locator_status` holds the condition: the
integer value conditions, `DOMAIN_YEAR`, `DOMAIN_MONTH`, `DOMAIN_WEEKDAY`,
`DOMAIN_OCCURRENCE`, `OCCURRENCE_ABSENT`, `RESULT_WINDOW` and the control
conditions.

**Validation:** the [validation plan](VALIDATION_PLAN.md#date-layer-coverage).

### `CALC-08` Date from pillar

**Purpose:** the date that a pillar token, such as `3M` or `1Y6M`, names from
a start date.

**Inputs:**

| Input | Argument and reading | Unit |
| --- | --- | --- |
| Start date | `StartDate`, a [date value](DATA_CONTRACT.md#date-values) | Day |
| Pillar | `Pillar`, a [pillar token](DATA_CONTRACT.md#pillar-tokens) | — |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | `ON` and `O/N` are `1D`; `TN` and `T/N` are `2D`, in calendar days | Decision (#4) |
| 2 | The year and month components form one signed month shift, twelve months a year; the week and day components form one signed day shift, seven days a week; the token's sign applies to both | Decision (#4) |
| 3 | The month shift is applied first, with the clip of `CALC-06` step 3 and without `Opt_KeepEOM`, then the day shift exactly | Decision (#4) |
| 4 | A shift outside the `Long` range, or a result outside the window, is `#NUM!`, `PILLAR_AGGREGATE_RANGE` | Decision (#4) |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Pillar date | `pillar_date` | Date | Each element | None |

**Limits and non-goals:** no business-day meaning for `ON` or `TN`: that
belongs to the calendar milestones, through calendar-aware functions, so a
result of this function never changes when they arrive.

**Failure behaviour:** `date_from_pillar_status` holds the condition: the date
value conditions and the pillar conditions.

**Validation:** the [validation plan](VALIDATION_PLAN.md#date-layer-coverage).

### `CALC-09` Pillar from dates

**Purpose:** the pillar token that labels the interval from a start date to an
end date, rounded to a curve bucket.

**Inputs:**

| Input | Argument and reading | Unit |
| --- | --- | --- |
| Start date | `StartDate`, a [date value](DATA_CONTRACT.md#date-values) | Day |
| End date | `EndDate`, a date value | Day |
| Rounding | `Opt_Rounding`, a [control](DATA_CONTRACT.md#control-values), default `NEAREST` | — |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | Equal dates give `0D` | Decision (#4) |
| 2 | An interval shorter than seven days, in absolute value, gives its exact signed day count, such as `5D` or `-3D` | Decision (#4) |
| 3 | A longer interval chooses among anchors generated from `StartDate` in the direction of `EndDate`: the `1W`, `2W` and `3W` anchors, the floor month anchor (the largest positive month count whose anchor does not pass `EndDate`) and the ceiling month anchor (that count plus one). When no positive count qualifies, as from 2026-01-01 to 2026-01-28, the floor count is 0: it has no anchor and is not a candidate, and the ceiling anchor is `1M`. Month anchors use the clip of `CALC-06` step 3, with signed counts, so that a negative token reads back the same way. Week tokens beyond `3W` are never candidates | Decision (#4) |
| 4 | An anchor outside the window is not a candidate: it is excluded, never approximated | Decision (#4) |
| 5 | `FLOOR` chooses the largest candidate whose anchor does not pass `EndDate`; `CEILING` the smallest whose anchor reaches or passes it; `NEAREST` the one closest to `EndDate` in calendar days. Rounding works on the absolute interval, then restores the minus sign of a negative one | Decision (#4) |
| 6 | Ties are broken the same way in every mode: between a week and a month anchor, the month; between two month anchors, the larger count. An exact anchor wins outright | Decision (#4) |
| 7 | When no candidate reaches `EndDate` under `CEILING`, the call is `#NUM!`, `RESULT_WINDOW`, never the largest anchor that exists | Decision (#4) |
| 8 | A month count of 1 to 11 is written `nM`; 12 or more as years and the months left, `nY` or `nYmM`. The written tokens are `[-] (0D \| nD \| 1W \| 2W \| 3W \| nM \| nY \| nYmM)`, with no `+`, alias, blank, zero component or mixed weeks and months | Decision (#4) |

- **Week cap:** a whole-week anchor always lands within three days of any
  end date, while a month anchor can be fifteen days away, so without the cap
  nearly every long interval would be weeks: ten years and two weeks would be
  `524W`. `4W` and longer are therefore never written, and the cap applies
  in every mode.
- **3W to 1M:** under `NEAREST` the boundary follows the distance rule. With
  a `1M` anchor in the window, 24 days give `3W` and 27 days `1M`; 25 and 26
  days go to the closer anchor, a tie to the month. From a start date whose
  `1M` anchor is 31 days away, 25 days give `3W` and 26 `1M`; from 31 January,
  whose `1M` anchor clips to 28 February, 25 days already give `1M`.
- **Near 9999-12-31:** an eleven-day interval gives `2W` under `NEAREST`
  except in the last fortnight of 9999, where the `2W` and `1M` anchors leave
  the window: `NEAREST` and `FLOOR` give `1W`, and `CEILING` has no candidate
  (step 7). `FLOOR` and `NEAREST` always have one, since the `1W` anchor lies
  between two dates in the window seven or more days apart.
- **Round trip:** every written token reads back with `CALC-08`, but
  `DateFromPillar(StartDate, PillarFromDates(StartDate, EndDate))` equals
  `EndDate` only when the interval is an exact anchor: tests never assume it
  for a rounded one.

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Pillar | `pillar_token` | Text | Each element | Rules 3 to 7 |

**Limits and non-goals:** no business-day pillar, no week token beyond `3W`.

**Failure behaviour:** `pillar_from_dates_status` holds the condition: the
date value conditions, the control conditions and `RESULT_WINDOW`.

**Validation:** the [validation plan](VALIDATION_PLAN.md#date-layer-coverage).

### `CALC-10` Host date system

**Purpose:** which date system the calling worksheet uses, so that a caller
can tell a library `#N/A` from one passed in.

**Inputs:** none: the caller, through `Application.Caller`.

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | `KPR_Dates_HostDateSystem` calls `Application.Volatile True` first, so that an ordinary recalculation evaluates it again after the caller or the workbook's date system changes | `SRC-07` |
| 2 | It returns 1900 for a worksheet caller on the 1900 date system, 1904 for one on the 1904 date system, and 1900 when no worksheet host is identified, by the [host rule](#host-and-date-system) | `SRC-05`, `SRC-06` |
| 3 | It returns `#N/A`, `HOST_UNRESOLVED`, only when a worksheet host was identified and its date system cannot be read; a 1904 workbook is a result here, not a refusal | Decision (#4) |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Date system | `host_date_system` | 1900 or 1904 | Whole call | None |

**Limits and non-goals:** no array; the only volatile function.

**Failure behaviour:** `host_date_system_status` holds `HOST_UNRESOLVED`.

**Validation:** the host checks of the
[validation plan](VALIDATION_PLAN.md#date-layer-coverage), in Excel only.

## Decisions

| Date | Decision | Reason | Issue |
| --- | --- | --- | --- |
| *Illustrative* | How a rule the source leaves open is applied | Why | Link |
| 2026-11-25 | The date layer's contracts start at `CALC-02`; `CALC-01` stays illustrative | The illustrative test cases test `CALC-01` until [#11](https://github.com/danielep71/K-PRICING/issues/11) replaces them, and an identifier is never reused | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-25 | A date-layer contract's *Inputs* name the worksheet argument and how it is read, not a dataset field | The functions read no dataset ([data contract](DATA_CONTRACT.md#function-arguments)) | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-25 | Only a worksheet caller decides the date system; a 1904 workbook gets `#N/A`, never a shifted value, and no other workbook is consulted | A value shifted by 1,462 days looks valid; `ActiveWorkbook` can be any workbook | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-25 | Arrays keep their shape, a scalar repeats over the others, a result is capped at 100,000 elements before anything is read, and a failed element does not stop its neighbours | One rule for every function, a bounded cost per call, and a result that maps cell for cell to its inputs | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-25 | `DayOfWeek` and `WdIndex` are Monday-based by default; `Opt_WeekBaseMonday=FALSE` gives `WEEKDAY`'s numbering | Monday = 1 is the ISO 8601 and market convention | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-25 | `AddMonths` clips to the target month's last day; `Opt_KeepEOM` keeps a month end; `AddYears` is twelve months | One month rule for every shift and for pillars | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-25 | `PillarFromDates` writes weeks only up to `3W`, chooses among week and month anchors under `FLOOR`, `CEILING` and `NEAREST` with fixed tie-breaks, excludes anchors outside the window, and fails rather than fall back under `CEILING` | Weeks label only the short end of a curve; a label must be repeatable and never silently another mode | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-26 | No date is built with a day of zero or below; the window is declared once, in `CORE_Parse`; the host rule runs once per call, before any argument is read, in `EvaluateCall` only; only `KPR_Dates_HostDateSystem` is volatile; only `M_Host` reads `Application.Caller`; no date or number is read by a conversion that depends on regional settings. `tools/check_date_domain.py` enforces them | Each is already a rule of the guides above; a source check keeps a later edit from breaking one without a test that fails first. A day of zero in the month after December 9999 is the one defect that no case reaches | [#15](https://github.com/danielep71/K-PRICING/issues/15) |
| 2026-11-25 | Every date-layer comparison is exact | The results are whole days, whole numbers, Booleans and texts | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-26 | The values that validate `CALC-02` to `CALC-07` are Excel's own worksheet functions, `IMPL-01`, as the [Excel oracle](EXCEL_ORACLE.md) computed them in a passing Excel run; the holdout of `ACC-02` to `ACC-20` is every fifth, by ID, of the function's cases the oracle covers, and every other case calibrates; the contracts stay `draft`, and whether each freezes is decided with the release certification | Excel's functions are an implementation by another author that the date layer never calls; a holdout of `independent` values is what a frozen contract needs, and the release certification runs the candidate the guarantee would cover. `ACC-16` and `ACC-18` hold fewer than `min_holdout_cases` such cases, and `ACC-21` to `ACC-23` none | [#48](https://github.com/danielep71/K-PRICING/issues/48) |
