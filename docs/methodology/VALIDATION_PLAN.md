# ✅ Validation plan

This plan says how each [calculation contract](CALCULATION_CONTRACTS.md) is
shown to hold, and what evidence a result needs before anyone relies on it.

## Responsibilities

Implementing a calculation, checking it independently and accepting it are
separate steps, recorded separately even when one person does all three.
Code review and passing repository checks are necessary, but they are not
evidence that a figure is right: the static checks never run Excel or a
calculation.

## Reference values

An expected result is never taken from the workbook. Two kinds of value
validate a rule:

- **Published:** printed in a registered [source](SOURCES.md), such as a
  worked example of the regulation.
- **Independent:** computed by an external implementation by another author,
  registered in the source register with an `IMPL-` ID.

A value the project computes itself, by hand, in a separate spreadsheet that
shares no formulas with the workbook, or with a Python script under
`tests/derivations/`, is an `illustrative` cross-check: it catches defects,
but it validates no rule. Record how each value was derived and who derived
it. A reference script is a test tool, never a runtime dependency of the
workbook.

Each comparison has an absolute and a relative tolerance, with the reason for
them: the source's own rounding, or the precision of `Double` arithmetic.

The [numerical test cases](TEST_CASES.md) record each expected value with its
reference class (`published`, `independent` or `illustrative`), its source,
its derivation and its tolerance
([reference classes](TEST_CASES.md#reference-classes)). Only `published` and
`independent` values validate a rule: a rule reaches *Validated* in the
[traceability table](#traceability) only through one of them.

## Validation matrix

| Calculation | Reference values | Evidence | Status |
| --- | --- | --- | --- |
| *Illustrative:* `CALC-01` | `illustrative` cross-checks only: two `nominal` cases, a JSON fixture and a file-based case, whose total of two synthetic notional amounts a derivation script computes; empty input and invalid input rejected are still to cover; no `published` or `independent` value yet | Evidence record per release candidate | Not implemented |
| `CALC-02` Day of week | The VBA suite `TEST_CoreDates`, whose values come from a Python reference and are not `published` or `independent`; the worksheet-function cases of [#11](https://github.com/danielep71/K-PRICING/issues/11), `illustrative`, run in Excel in [#13](https://github.com/danielep71/K-PRICING/issues/13); the [Excel oracle](EXCEL_ORACLE.md) of [#14](https://github.com/danielep71/K-PRICING/issues/14) recomputes 48 of them with Excel's worksheet functions, all confirmed in the Excel run of `573884baeb42` ([report](https://github.com/danielep71/K-PRICING/issues/14#issuecomment-7199627903)); those values are `independent`, computed by `IMPL-01` ([#48](https://github.com/danielep71/K-PRICING/issues/48)), and the rule is *Validated* only once they pass on an Excel run of the release candidate | Evidence record per release candidate | Not validated |
| `CALC-03` Month and year lengths | The VBA suite `TEST_CoreDates`, whose values come from a Python reference and are not `published` or `independent`; the worksheet-function cases of [#11](https://github.com/danielep71/K-PRICING/issues/11), `illustrative`, run in Excel in [#13](https://github.com/danielep71/K-PRICING/issues/13); the [Excel oracle](EXCEL_ORACLE.md) of [#14](https://github.com/danielep71/K-PRICING/issues/14) recomputes 55 of them with Excel's worksheet functions, all confirmed in the Excel run of `573884baeb42` ([report](https://github.com/danielep71/K-PRICING/issues/14#issuecomment-7199627903)); those values are `independent`, computed by `IMPL-01` ([#48](https://github.com/danielep71/K-PRICING/issues/48)), and the rule is *Validated* only once they pass on an Excel run of the release candidate | Evidence record per release candidate | Not validated |
| `CALC-04` Period boundaries | The VBA suite `TEST_CoreDates`, whose values come from a Python reference and are not `published` or `independent`; the worksheet-function cases of [#11](https://github.com/danielep71/K-PRICING/issues/11), `illustrative`, run in Excel in [#13](https://github.com/danielep71/K-PRICING/issues/13); the [Excel oracle](EXCEL_ORACLE.md) of [#14](https://github.com/danielep71/K-PRICING/issues/14) recomputes 144 of them with Excel's worksheet functions, all confirmed in the Excel run of `573884baeb42` ([report](https://github.com/danielep71/K-PRICING/issues/14#issuecomment-7199627903)); those values are `independent`, computed by `IMPL-01` ([#48](https://github.com/danielep71/K-PRICING/issues/48)), and the rule is *Validated* only once they pass on an Excel run of the release candidate | Evidence record per release candidate | Not validated |
| `CALC-05` Period-end tests | The VBA suite `TEST_CoreDates`, whose values come from a Python reference and are not `published` or `independent`; the worksheet-function cases of [#11](https://github.com/danielep71/K-PRICING/issues/11), `illustrative`, run in Excel in [#13](https://github.com/danielep71/K-PRICING/issues/13); the [Excel oracle](EXCEL_ORACLE.md) of [#14](https://github.com/danielep71/K-PRICING/issues/14) recomputes 72 of them with Excel's worksheet functions, all confirmed in the Excel run of `573884baeb42` ([report](https://github.com/danielep71/K-PRICING/issues/14#issuecomment-7199627903)); those values are `independent`, computed by `IMPL-01` ([#48](https://github.com/danielep71/K-PRICING/issues/48)), and the rule is *Validated* only once they pass on an Excel run of the release candidate | Evidence record per release candidate | Not validated |
| `CALC-06` Date shifts | The VBA suite `TEST_CoreDates`, whose values come from a Python reference and are not `published` or `independent`; the worksheet-function cases of [#11](https://github.com/danielep71/K-PRICING/issues/11), `illustrative`, run in Excel in [#13](https://github.com/danielep71/K-PRICING/issues/13); the [Excel oracle](EXCEL_ORACLE.md) of [#14](https://github.com/danielep71/K-PRICING/issues/14) recomputes 57 of them with Excel's worksheet functions, all confirmed in the Excel run of `573884baeb42` ([report](https://github.com/danielep71/K-PRICING/issues/14#issuecomment-7199627903)); those values are `independent`, computed by `IMPL-01` ([#48](https://github.com/danielep71/K-PRICING/issues/48)), and the rule is *Validated* only once they pass on an Excel run of the release candidate | Evidence record per release candidate | Not validated |
| `CALC-07` Weekday locators | The VBA suite `TEST_CoreDates`, whose values come from a Python reference and are not `published` or `independent`; the worksheet-function cases of [#11](https://github.com/danielep71/K-PRICING/issues/11), `illustrative`, run in Excel in [#13](https://github.com/danielep71/K-PRICING/issues/13); the [Excel oracle](EXCEL_ORACLE.md) of [#14](https://github.com/danielep71/K-PRICING/issues/14) recomputes 40 of them with Excel's worksheet functions, all confirmed in the Excel run of `573884baeb42` ([report](https://github.com/danielep71/K-PRICING/issues/14#issuecomment-7199627903)); those values are `independent`, computed by `IMPL-01` ([#48](https://github.com/danielep71/K-PRICING/issues/48)), and the rule is *Validated* only once they pass on an Excel run of the release candidate | Evidence record per release candidate | Not validated |
| `CALC-08` Date from pillar | The VBA suite `TEST_CoreDates`, whose values come from a Python reference and are not `published` or `independent`; the worksheet-function cases of [#11](https://github.com/danielep71/K-PRICING/issues/11), `illustrative`, run in Excel in [#13](https://github.com/danielep71/K-PRICING/issues/13); no external implementation defines pillars ([open decisions](#open-decisions)) | Evidence record per release candidate | Not validated |
| `CALC-09` Pillar from dates | The VBA suite `TEST_CoreDates`, whose values come from a Python reference and are not `published` or `independent`; the worksheet-function cases of [#11](https://github.com/danielep71/K-PRICING/issues/11), `illustrative`, run in Excel in [#13](https://github.com/danielep71/K-PRICING/issues/13); no external implementation defines pillars ([open decisions](#open-decisions)) | Evidence record per release candidate | Not validated |
| `CALC-10` Host date system | Three worksheet-function cases of [#11](https://github.com/danielep71/K-PRICING/issues/11), one per caller, `illustrative`, run in Excel in [#13](https://github.com/danielep71/K-PRICING/issues/13); Excel host checks ([host behaviour](#host-behaviour)) | Evidence record per release candidate | Not implemented |

The *Status* of a calculation is the least advanced
[traceability](#traceability) status of its rules: *Validated* only when every
rule is.

Cover, for each calculation, with the case categories of the
[test-case format](TEST_CASES.md#categories):

- typical cases (`nominal`) and the boundaries of every rule (`boundary`);
- invalid and missing input, which the [data contract](DATA_CONTRACT.md)
  rejects (`invalid`).

Cover with VBA tests, which add no category
([relation, end-to-end and volume tests](TEST_CASES.md#other-tests)):

- failure: an operation that fails clears its earlier results and marks them
  `FAILED` ([status and staleness](WORKBOOK_CONTRACT.md#status-and-staleness));
- repeatability: the same input gives the same output.

A relation test or the end-to-end case runner of the
[harness guide](TEST_HARNESS.md) can check either.

<a id="date-layer-coverage"></a>

### Date-layer coverage

The cases of the 22 `KPR_Dates_*` functions ([#11](https://github.com/danielep71/K-PRICING/issues/11)) cover each contract
with the [case categories](TEST_CASES.md#categories):

| Category | Covers, for the date layer |
| --- | --- |
| `nominal` | Dates across the window and in every month, quarter and weekday; leap and common years; both weekday bases; `Opt_KeepEOM` both ways; every `Opt_Rounding` mode; every pillar unit, alias and order |
| `boundary` | The window's edges: 1900-03-01, 9999-12-31, serial 61, and one day beyond each; 1900-02-28 and `1900-02-29` as text; leap days in 2000, 1900 and 2100; month ends of 28, 29, 30 and 31 days; quarter and year ends; shifts whose result lands on an edge and one day past it; `CALC-04` step 4 in 1900; fifth weekdays present and absent; `CALC-07` results in early 1900; pillar intervals of 6 and 7 days and of 24 to 27 days, a 31 January start, and the last fortnight of 9999; whole numbers at `1E-9` from an integer and just beyond, and at the edges of `Long` |
| `invalid` | Every [condition](DATA_CONTRACT.md#conditions) of every argument a function takes, each with at least one case, as a [worksheet-function case](TEST_CASES.md#worksheet-function-cases) where it can express the input, otherwise as a VBA test |

The worksheet-function cases hold, as encoded values, a VBA `Date`,
`Empty`, `Null`, an Excel error, an empty array, a jagged array and the
100,000 and 100,001 element calls, and the array rules of
[arrays and evaluation](CALCULATION_CONTRACTS.md#arrays-and-evaluation): a
kept shape, a scalar repeated over an array, a one-dimensional array read as a
row, `SHAPE_MISMATCH`, an element-level error beside valid neighbours and the
first failing argument at a position. Inputs that they cannot express are VBA
tests, run by the harness ([#13](https://github.com/danielep71/K-PRICING/issues/13)):
an object, a Range of several areas, a `Decimal`, and the
order of evaluation.

<a id="date-layer-completeness"></a>

#### Completeness policy

Each value-taking function has at least one `nominal`, one `boundary` and one
`invalid` worksheet-function case, and an `invalid` case for each condition
below, at the argument or level it names ([#12](https://github.com/danielep71/K-PRICING/issues/12)).
`KPR_Dates_HostDateSystem` takes no argument and needs one `nominal` case per
caller. `tests/test_date_layer_completeness.py` reads these tables and fails on
any function, argument or condition without a case: it attributes a case's
condition to the first argument that fails at that position, as the
evaluation order of the [calculation contracts](CALCULATION_CONTRACTS.md#arrays-and-evaluation) gives.

Every argument needs a case for each condition of its kind:

| Arguments | Kind | Conditions |
| --- | --- | --- |
| `DateIn`, `StartDate`, `EndDate` | date | `DATE_TEXT_FORMAT`, `DATE_TEXT_LOCALE`, `DATE_TEXT_NUMERIC`, `DATE_TEXT_IMPOSSIBLE`, `DATE_TYPE_REJECTED`, `DATE_WINDOW`, `INPUT_BLANK_REQUIRED`, `INPUT_ERROR_PROPAGATED` |
| `YearIn` | year | `INTEGER_TYPE_REJECTED`, `INTEGER_FRACTION`, `INTEGER_RANGE`, `INPUT_BLANK_REQUIRED`, `INPUT_ERROR_PROPAGATED`, `DOMAIN_YEAR` |
| `MonthIn` | month | `INTEGER_TYPE_REJECTED`, `INTEGER_FRACTION`, `INTEGER_RANGE`, `INPUT_BLANK_REQUIRED`, `INPUT_ERROR_PROPAGATED`, `DOMAIN_MONTH` |
| `WdIndex` | weekday | `INTEGER_TYPE_REJECTED`, `INTEGER_FRACTION`, `INTEGER_RANGE`, `INPUT_BLANK_REQUIRED`, `INPUT_ERROR_PROPAGATED`, `DOMAIN_WEEKDAY` |
| `n` | occurrence | `INTEGER_TYPE_REJECTED`, `INTEGER_FRACTION`, `INTEGER_RANGE`, `INPUT_BLANK_REQUIRED`, `INPUT_ERROR_PROPAGATED`, `DOMAIN_OCCURRENCE` |
| `nDays`, `nWeeks`, `nMonths`, `nYears` | int | `INTEGER_TYPE_REJECTED`, `INTEGER_FRACTION`, `INTEGER_RANGE`, `INPUT_BLANK_REQUIRED`, `INPUT_ERROR_PROPAGATED` |
| `Pillar` | pillar | `PILLAR_TYPE_REJECTED`, `PILLAR_TOKEN_MALFORMED`, `PILLAR_DUPLICATE_UNIT`, `PILLAR_ALIAS_SIGNED`, `INPUT_BLANK_REQUIRED`, `INPUT_ERROR_PROPAGATED` |
| `Opt_WeekBaseMonday`, `Opt_KeepEOM` | bool-control | `CONTROL_TYPE_REJECTED`, `CONTROL_NOT_SCALAR`, `CONTROL_ERROR_PROPAGATED` |
| `Opt_Rounding` | rounding-control | `CONTROL_TYPE_REJECTED`, `CONTROL_NOT_SCALAR`, `CONTROL_ERROR_PROPAGATED`, `CONTROL_TOKEN_UNKNOWN` |

Every value-taking function also needs `HOST_DATE1904`, `SHAPE_UNSUPPORTED`
and `CAPACITY_EXCEEDED`, a function of two or more value arguments
`SHAPE_MISMATCH`, and each function below the conditions of its result:

| Function | Result conditions |
| --- | --- |
| `KPR_Dates_BeginOfQuarter` | `RESULT_WINDOW` |
| `KPR_Dates_BeginOfYear` | `RESULT_WINDOW` |
| `KPR_Dates_AddDays` | `RESULT_WINDOW` |
| `KPR_Dates_AddWeeks` | `RESULT_WINDOW` |
| `KPR_Dates_AddMonths` | `RESULT_WINDOW` |
| `KPR_Dates_AddYears` | `RESULT_WINDOW` |
| `KPR_Dates_NthWeekdayOfMonth` | `OCCURRENCE_ABSENT`, `RESULT_WINDOW` |
| `KPR_Dates_LastWeekdayOfMonth` | `RESULT_WINDOW` |
| `KPR_Dates_DateFromPillar` | `PILLAR_AGGREGATE_RANGE` |
| `KPR_Dates_PillarFromDates` | `RESULT_WINDOW` |

`HOST_UNRESOLVED` needs a worksheet host whose date system cannot be read, and
`ERR_UNMAPPED_CONDITION` a defect: no case can hold either, so they are VBA
tests ([#13](https://github.com/danielep71/K-PRICING/issues/13)).

<a id="date-domain-static-rules"></a>

#### Date-domain static rules

`python tools/check_date_domain.py`, which **Repository integrity** runs as
`date-domain`, reads the production source under `src/` and enforces the rules
below, which the generic [VBA checks](../VBA_HOUSE_STYLE.md#static-checks) do
not cover ([#15](https://github.com/danielep71/K-PRICING/issues/15)). Each rule
has self-tests in `tests/test_check_date_domain.py` that make the defect in a
copy of the production source and show the check rejecting it.

| Rule | It fails when | Enforces |
| --- | --- | --- |
| `day-zero` | a `DateSerial` call has a literal day of zero or below, in any VBA number notation (`0&`, `0#`, `0.0`, `0E0`, `&H0`, `(0)`, a negative number) | [Common rules](CALCULATION_CONTRACTS.md#common-rules), Dates: a month's last day comes from the month's length, `CALC-03` step 2 |
| `window` | `CORE_Parse` does not declare `WINDOW_MIN_SERIAL` as `61#`, `WINDOW_MAX_SERIAL` as `2958465#`, `WINDOW_MIN_YEAR` as `1900` and `WINDOW_MAX_YEAR` as `9999`, or another place declares one of them again or writes an edge of the window out | [Common rules](CALCULATION_CONTRACTS.md#common-rules), Dates: the supported window is 1900-03-01 to 9999-12-31 |
| `host-guard` | a worksheet function of the facade other than `KPR_Dates_HostDateSystem` does not call `EvaluateCall` exactly once, or applies the host rule itself; `EvaluateCall` does not call `M_Host.TryCheckHost` exactly once, or calls a public procedure of `M_Ranges`, `CORE_Array`, `CORE_Parse` or `CORE_Dates`, qualified or not, before it; the diagnostic does not call `M_Host.TryReadHostDateSystem` exactly once; another component applies the host rule | [Host and date system](CALCULATION_CONTRACTS.md#host-and-date-system): each public call reads the host once, at its start, before any argument |
| `volatility` | `Application.Volatile` is called anywhere but as the first statement of `KPR_Dates_HostDateSystem`, with an argument other than the literal `True`, or not exactly once in production | [Host and date system](CALCULATION_CONTRACTS.md#host-and-date-system), and `CALC-10` step 1: only `KPR_Dates_HostDateSystem` is volatile |
| `host-authority` | `Application.Caller` is read outside `M_Host` or not exactly once there, or `M_Host` or the facade names `ActiveWorkbook`, `ThisWorkbook`, `ActiveSheet`, `ActiveCell`, `Selection` or `Workbooks` | [Host and date system](CALCULATION_CONTRACTS.md#host-and-date-system), steps 1 and 5: only the caller's own workbook decides, and no other is consulted |
| `locale` | production calls `IsDate`, `DateValue`, `CVDate`, `IsNumeric`, `Format`, `Format$` or `FormatDateTime`, or `CDate` on anything but a variable named `...Serial` | [Date values](DATA_CONTRACT.md#date-values): no parse uses a conversion that depends on regional settings |

The check joins a statement continued over lines, reads comments out and blanks
string literals; it does not compile
or run VBA, follows no call through a variable, and judges a computed day or a
`CDate` argument only by its name, so a defect written another way, such as a
day held in a variable that happens to be zero, stays with the cases and the
code review.

[Relation tests](TEST_HARNESS.md#relation-tests) check properties, each with
its corruption check, over the dates whose inputs and intermediate results
all stay in the window. Each relation holds only there: `BeginOfYear` of a
1900 date, `BeginOfQuarter` of a date in the first quarter of 1900, `d + 7`
beyond 9999-12-31 and a shift that leaves the window return their documented
error (`CALC-04` step 4, `RESULT_WINDOW`, `DATE_WINDOW`), and those edges are
`boundary` cases that expect that error, not relation tests:

- `DayOfWeek(d + 7) = DayOfWeek(d)`;
- `BeginOfMonth(d) ≤ d ≤ EndOfMonth(d)`, and the same for quarters and years;
- `IsMonthEnd(EndOfMonth(d))`, `IsQuarterEnd(EndOfQuarter(d))` and
  `IsYearEnd(EndOfYear(d))` are `TRUE`;
- `DaysInMonth(d) = EndOfMonth(d) − BeginOfMonth(d) + 1`, and
  `DaysInYear(y) = 365 + IsLeapYear(y)`, counting `TRUE` as 1;
- `AddWeeks(d, n) = AddDays(d, 7n)` and `AddYears(d, n) = AddMonths(d, 12n)`,
  with the same `Opt_KeepEOM`;
- `DateFromPillar(s, PillarFromDates(s, e)) = e` when `e` is an exact anchor
  of `s`, and never assumed otherwise (`CALC-09`).

The host checks run in Excel only, as probes that record the
[host behaviour](#host-behaviour) entries `HB-08` to `HB-14`: a worksheet
caller in a 1900 and in a 1904 workbook, a call from VBA, from
`Application.Run` and from the Immediate window, the contexts that give a
caller that is not a Range, the recalculation of `KPR_Dates_HostDateSystem`,
and a multi-cell result spilled by dynamic-array Excel.

<a id="regression-coverage"></a>

#### Regression coverage

Each regression area of the date layer maps onto a certified suite. A check
on a worksheet function whose inputs the
[case format](TEST_CASES.md#worksheet-function-cases) can express is a
`regression-<area>-<label>` case, with its result and condition derived by the
reference model; every other check is a literal assertion of a VBA suite.

| Regression area | Covered by |
| --- | --- |
| Date types, date text, date window, integers | `core-err-parse`: the parser cases, with `date-regression` and `long-regression` for numeric types, text shapes and tolerances beyond the ordinary ones |
| Controls | `core-err-parse` (`bool-control`, `rounding-control`), `core-array` (`unwrap-control`, `regression-shapes`), `dates` (`controls`: a forwarded `Missing` and a tab-padded `Opt_Rounding`) |
| Boundaries | `core-err-parse` (`date-window`) and 12 `regression-boundary-*` cases |
| Conditions and their errors | `core-err-parse` (`err-constructors`, `condition-registry`, `unmapped-conditions`) |
| Host | 3 `regression-host-*` cases and the `KPR_Dates_HostDateSystem` cases; `excel-host` (`date-system-toggle`: one workbook from 1900 to 1904 and back, on an ordinary recalculation) |
| Pillars | `core-dates` (`parse-pillar`, `add-pillar`, `format-pillar`, `pillar-regression`) and 55 `regression-pillar-*` cases |
| Function surface | 49 `regression-surface-*` cases |
| Shapes | `core-array`, with `regression-shapes` for an error scalar, the 250 by 400 block and a one-cell two-dimensional control; `ranges`, and `excel-host` (`range-block`: a block of a workbook's cells with a blank and an error, a column, a Range beyond the used range, the selection unchanged) |
| Array and scalar parity | 10 `regression-parity-*` cases; every array case states each element's value, derived element by element |
| Worksheet-function cases | `date-cases` for a VBA caller; `worksheet-cases` for a worksheet caller in either date system |
| Multi-cell results | `excel-host` (`spill`: a dynamic-array formula spills one result per row, an element's error beside its neighbours, and nothing under 1904) |
| Function registration | `excel-host` (`registration`: registering and unregistering, each twice, reach the 22 functions and leave the session unchanged) |

No case holds a `LongLong`: one cannot be built on a 32-bit host, so a suite
that held one would not run there.

## Traceability

Each rule is traced from the paragraph of its source to the step of the
calculation contract that states it, to the VBA procedure that implements it
and to the test cases that test it, with its validation status:

| Source | Paragraph | Calculation step | VBA procedure | Test cases | Status |
| --- | --- | --- | --- | --- | --- |
| *Illustrative:* `SRC-01` | Article 1(1) | `CALC-01.1` | — | `illustrative-format-example`, `illustrative-file-import-example` | Not implemented |
| `SRC-09` | Return type 1 | `CALC-02.2` | `CORE_Dates.DayOfWeek` | — | Not validated |
| `SRC-02` | Gregorian calendar, leap year | `CALC-03.1` | `CORE_Dates.IsLeapYear` | — | Not validated |
| `SRC-02` | Gregorian calendar | `CALC-03.2` | `CORE_Dates.DaysInMonth` | — | Not validated |
| `SRC-02` | Gregorian calendar, leap year | `CALC-03.3` | `CORE_Dates.DaysInYear` | — | Not validated |
| `SRC-02` | Gregorian calendar, leap year | `CALC-03.4` | `CORE_Dates.IsLeapYear` | — | Not validated |
| `SRC-02` | Gregorian calendar | `CALC-04.3` | `CORE_Dates.EndOfMonth`, `CORE_Dates.EndOfQuarter`, `CORE_Dates.EndOfYear` | — | Not validated |

- **Source and paragraph:** a `SRC-` ID of the [source register](SOURCES.md)
  and the exact paragraph, as in the step's *Source* column.
- **Calculation step:** `CALC-NN.<step>`, a step of the *Rules* table of
  that [calculation contract](CALCULATION_CONTRACTS.md). A rule implemented
  in several steps has one row per step.
- **VBA procedure:** `Module.Procedure` of a component under `src/`, or `—`
  while none implements the rule. Name every procedure that applies the rule.
- **Test cases:** the IDs of the [test cases](TEST_CASES.md) that test the
  step, or `—`. Each case tests the row's calculation.
- **Status:** one of the four below.

| Status | Meaning | The row names |
| --- | --- | --- |
| Validated | A `published` or `independent` case passes for the rule on an Excel run of the release candidate ([completeness policy](TEST_CASES.md#completeness-policy)) | A procedure, and a case with a `published` or `independent` output whose `rule` is the row's step |
| Exercised | Cases run the rule, but none is a `published` or `independent` case that passed: they are `illustrative`, or the reference cases have not passed yet | A procedure and at least one case |
| Not validated | Implemented, but no case tests it | A procedure and no case |
| Not implemented | No procedure implements it yet; its cases may come first | No procedure |

`python tools/check_test_cases.py`, which **Repository integrity** runs,
fails a row whose source is not in the register, whose step is not one that
the *Rules* table of a declared calculation lists, whose procedure is not in
a component under `src/`, whose case does not exist or tests another
calculation, or whose status contradicts what the row names. It cannot see an Excel run: whether a
`Validated` rule's case passed on the release candidate is the
[evidence record](#evidence-record)'s to show.

## Evidence record

Each run of a check in Excel records:

- the commit of the source and the workbook it was built into;
- the case, its expected value, that value's provenance and the tolerance;
- the observed value and the outcome;
- the Excel build, its bitness and the Windows version;
- who ran it, who reviewed it, and when.

A check that was not run is recorded as `NOT RUN` with the reason. It is
never a pass.

An Excel run of the harness keeps the host, who ran it, when and how it
ended in its [host record](TEST_HARNESS.md#host-record), also when it stopped
early. `tools/check_excel_run.py` rejects a record that lacks one of them or
leaves who ran it, when or how it ended `UNKNOWN`, and accepts an outcome
other than `PASS` only as evidence of that outcome.

### What an accuracy contract proves

A frozen [accuracy contract](TEST_CASES.md#accuracy-contracts) whose results
pass on an Excel run of the candidate shows that, on that host, the worst
error of the calculation over its calibration and holdout cases stays within
the threshold. The threshold is the guarantee a release may publish. A
target, when the contract declares one, is only reported: a worst error
above it but within the threshold passes, marked `target missed`, and shows
that the calculation is less accurate than aimed for, not that the guarantee
fails. The holdout cases played no part in setting the threshold,
and their values are `published` or `independent`. The evidence for the
first is the [holdout exposure record](TEST_CASES.md#holdout-exposure-record):
the holdout stays sealed while the contract is draft, and the record binds
its first measurement under the contract's decision to the contract entry,
its threshold, the holdout's fixtures and expected files, and the design
document if any. The check fails when any of them changes, or the threshold
is lowered, under the same decision. The record shows what was measured with
the tools, not what someone computed by hand: a threshold is only as
independent of its holdout as the people who set it kept it. It proves
nothing outside those cases, regime and host. A draft contract, a frozen
contract with no recorded exposure, and any contract on `illustrative`
values, validate nothing. Record each contract's worst errors with the
evidence, with its threshold, its target and the number of exposures of its
holdout.

An [invariant](TEST_CASES.md#invariants) of the registry that the same run
passes shows that, on that host, the results of its cases keep its bounds,
its order or its sum, whatever the status of the contracts over them. It
proves only what its `basis` states, on those cases, and holds for no other
input. Record each invariant's line with the evidence, with the monotone
reversals it counted within the allowance.

<a id="host-behaviour"></a>

## Host behaviour

The code relies on how Excel and the VBE behave: which settings an operation
changes, how a name is re-cased, what an import needs, how a window or an
event behaves. Such behaviour can differ by Excel build, channel, bitness and
policy. Record each behaviour that the code or a guide relies on here, with
whether it was measured:

| ID | Behaviour | Status | Probe | Host | Date | Result | Issue |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `HB-01` | The VBE re-cases every use of a name to its latest spelling, so an exported module can differ from its source only in the letter case of code outside strings and comments | `assumed` | The next Excel run by hand: the `Note:` lines of `tools/check_excel_run.py` | — | — | Not measured: no Excel run has recorded a re-cased export yet. `tools/check_excel_run.py` tolerates it: it compares that code without letter case and reports each re-cased identifier as a note | — |
| `HB-02` | Excel started through COM can follow the user's macro setting (`AutomationSecurity` is `msoAutomationSecurityByUI`, set by policy). A workbook it opens then has its macros disabled, and `Application.Run` fails with *Excel 4.0 (XLM) macros have been disabled*, because the VBA procedure is not found | `measured` | By hand: in PowerShell, open the built workbook through COM and `Run` `'<workbook>'!TEST_Harness.LastResult` | Microsoft 365 Excel 16.0 build 20430, channel not recorded, 64-bit; Windows 11 Enterprise 25H2, build 26200.9711 | 2026-11-03 | `AutomationSecurity` was `ByUI`, and `Run` failed with that message. `tools/devbuild/excel_run.py` changes no macro setting: it waits for **Enable Content** | — |
| `HB-03` | `Application.Run` reaches a public procedure of an `Option Private Module` module when the name is qualified with the workbook (`'Book.xlsm'!Module.Procedure`) | `measured` | As `HB-02`, in a session whose macros were enabled | As `HB-02` | 2026-11-03 | `'Workbook-development.xlsm'!TEST_Harness.LastResult` ran | — |
| `HB-04` | In a workbook opened through COM with its macros disabled, **Enable Content** enables them without reopening, and `Application.Run` then reaches its procedures | `measured` | `python tools/devbuild/excel_run.py` on a host with `HB-02`: click **Enable Content** in the yellow bar | As `HB-02` | 2026-11-03 | Excel showed the yellow bar; after the click, the certification ran to `excel-run: PASS` without reopening the workbook | — |
| `HB-05` | The project that `tools/devbuild/vba_project.py` writes holds no compiled code (`_VBA_PROJECT` version `0xFFFF`, no performance cache), so Excel reads it from source and compiles each module when it is first used: the code that runs is the source the build wrote, and no Trust access to the VBA project object model is needed | `assumed` | By hand: open a development build in Excel, check that `? TEST_Harness.LastResult` runs in the Immediate window, which compiles that module on its first use, then check that **Debug → Compile VBAProject**, which compiles the other modules, completes | — | — | Not measured by a probe of its own. `tools/devbuild/build_workbook.py` relies on it, and so does the identity by rebuild of `tools/check_excel_run.py`, which proves that the workbook Excel ran holds the source of the commit, not that Excel compiled it | — |
| `HB-06` | **End** in an error dialog, or **Run → Reset** in the VBE, resets the project and clears every module-level variable, so the harness's active run and recovery flag are lost and the stop is not detected | `assumed` | By hand: set a breakpoint on the line `TEST_Harness.BeginCase "text-exact"` of `TEST_Examples.CaseExactText`, type `TEST_Examples.RunExampleTests "<fixtures>"` with the full path of `tests/fixtures/` in the Immediate window, and when the run pauses there, choose **Run → Reset**; then check that `? TEST_Harness.LastResult` is empty and that the next run is not refused | — | — | Not measured. The harness guides tell you to check Excel's state by hand after such a stop | — |
| `HB-07` | The Immediate window keeps only its last 200 lines | `assumed` | By hand: print 250 numbered lines with `Debug.Print`, then check that the first line left is line 51 | — | — | Not measured. The installation guide tells you to save the whole log with `TEST_Harness.SaveLog` | — |
| `HB-08` | In a worksheet formula, `Application.Caller` is the calling Range, and its `Parent.Parent` is the workbook that holds the formula, even when another workbook is active | `assumed` | A probe of [#13](https://github.com/danielep71/K-PRICING/issues/13): a formula calling a test UDF in two open workbooks, the other one active | — | — | Not measured here. The [host rule](CALCULATION_CONTRACTS.md#host-and-date-system) relies on it ([SRC-05](SOURCES.md#register)) | — |
| `HB-09` | `Workbook.Date1904` reads that workbook's date system, and the 1904 system numbers the same day 1,462 lower | `assumed` | A probe of [#13](https://github.com/danielep71/K-PRICING/issues/13): the same date entered in a 1900 and in a 1904 workbook, and `Date1904` of each read through the caller | — | — | Not measured here ([SRC-03](SOURCES.md#register), [SRC-06](SOURCES.md#register)) | — |
| `HB-10` | Called from VBA, from the Immediate window or through `Application.Run`, `Application.Caller` is not a Range, and reading it raises no error | `assumed` | A probe of [#13](https://github.com/danielep71/K-PRICING/issues/13) that records the caller's `TypeName` in each context | — | — | Not measured here. The [host rule](CALCULATION_CONTRACTS.md#host-and-date-system) treats each as no worksheet host | — |
| `HB-11` | A macro attached to a shape, data validation, a chart series and a defined name give a caller that is not a Range: text, an error or another form | `assumed` | By hand: call a probe UDF from each context and record the caller's `TypeName` | — | — | Not measured here. These contexts are not supported | — |
| `HB-12` | A function that calls `Application.Volatile True` is evaluated again by an ordinary recalculation, F9, after the workbook's date system changes, without a full rebuild | `assumed` | By hand: enter `=KPR_Dates_HostDateSystem()`, switch the workbook to the 1904 date system, press F9, and read the cell | — | — | Not measured here ([SRC-07](SOURCES.md#register)) | — |
| `HB-13` | Excel's 1900 date system and VBA's `Date` name the same day from serial 61, 1900-03-01, on; below it they differ by one day, and Excel's serial 60 is 1900-02-29 | `assumed` | A probe of [#13](https://github.com/danielep71/K-PRICING/issues/13): `CDbl(DateSerial(1900, 3, 1))` in VBA against `DATE(1900,3,1)` and the serials 59 to 61 in a cell | — | — | Not measured here ([SRC-04](SOURCES.md#register), [SRC-11](SOURCES.md#register)). The supported window starts at 1900-03-01 because of it | — |
| `HB-14` | In dynamic-array Excel, a two-dimensional `Variant` array returned by a worksheet function spills into the cells beside its formula, and a blocked spill shows `#SPILL!` | `assumed` | By hand: enter `=KPR_Dates_EndOfMonth(A1:A3)` with free cells below, then with one blocked | — | — | Not measured here ([SRC-10](SOURCES.md#register)) | — |

- **Status:** `measured` when a probe ran on the named host and its output is
  kept; `assumed` otherwise.
- **Probe:** a `TEST_` procedure under `tests/modules/` that observes the
  behaviour and prints what it saw, run with the VBA test harness. A
  measurement found by hand names the steps instead.
- **Host:** the Excel product, version and build, its channel and bitness,
  and the Windows version, as in the [evidence record](#evidence-record).
- **Result:** what was observed, and the model it supports. A measurement
  holds for the named host only; it is never generalised to others.

A guide that states host behaviour the code relies on cites its `HB-` entry.
An `assumed` entry is listed under *Limits* in the
[release certification](../release/RELEASE_CERTIFICATION.md) until it is
measured.

## Open decisions

| Decision | Options | Issue |
| --- | --- | --- |
| How `CALC-08` and `CALC-09`, whose rules no external implementation defines, can reach *Validated* | Accept them as *Exercised* with `illustrative` values and list them as not validated in the release certification; register an external implementation of the same grammar, if one is found | [#12](https://github.com/danielep71/K-PRICING/issues/12) |
| How a traceability row cites a rule whose source is `Decision (#4)`, since the check accepts only an `SRC-` ID | Register each methodology guide's decisions as a source; let the check accept a decision reference. Decided with the rows that list the cases once they run in Excel | [#13](https://github.com/danielep71/K-PRICING/issues/13) |
