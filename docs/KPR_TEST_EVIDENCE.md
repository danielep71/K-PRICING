# K-PRICING test evidence

This document defines the durable VBA test runner's interface and the
structured evidence record it writes (#39). The committed schema is
[`kpr-test-evidence.schema.json`](kpr-test-evidence.schema.json).
[`tools/check_test_evidence.py`](../tools/check_test_evidence.py) validates a
record against that schema and against the semantic rules below. The validator
does not execute Excel.

The runner was verified in Windows Excel for #39: two runs on candidate
`9a978cd` (Excel 16.0 build 20326, 64-bit) each passed 17 suites and 1,577
assertions with state restored, and the two records matched outside the
declared nondeterministic fields. The #40 regression matrix and `worksheet-state` suite were verified the same
way on candidate `33b07cd` (18 suites, 1,953 assertions). The #41
`worksheet-oracle` suite was verified the same way on candidate `7d3f359` (19
suites, 9,177 assertions). The #42 `worksheet-registration` suite was verified
the same way on candidate `aa4aa0d` (20 suites, 9,263 assertions); a first run
on `e4e0c75` passed every assertion but failed state restoration because the
status bar read back as the text `FALSE`, and the runner now hands the bar back
with a literal `False`, retried after `DoEvents`.

## Runner interface

`tests/modules/KPR_REGRESSION_TESTS.bas` is the single hand-maintained runner,
assertion and suite module. It adds two durable entry points:

```vb
KPR_Test_RunAll(ByVal SourceSha As String, ByVal OutputFolder As String) As Boolean
KPR_Test_RunSuite(ByVal SuiteName As String, ByVal SourceSha As String, ByVal OutputFolder As String) As Boolean
```

Both are callable through `Application.Run` and the Immediate window. They are
test infrastructure, not supported production API, and are absent from
`docs/PUBLIC_API.txt`.

| Input | Contract |
| --- | --- |
| `SourceSha` | Full 40-character lowercase commit SHA of the imported source; anything else is refused before a suite runs |
| `OutputFolder` | Folder outside tracked source; it and any missing parent folders are created. A drive root or UNC share must already exist. The runner writes `kpr-test-evidence.json` there and overwrites an earlier file of that name, so use a fresh folder per run |
| `SuiteName` | One name from the registry below, compared case-insensitively; an unknown name runs nothing and writes nothing |

The return value is `True` only when every selected case passed, the caller's
state was restored and the evidence file was written. A refused input, a
failed case, a restoration failure or an unwritten file returns `False`. The
runner reports to the Immediate window with `Debug.Print` and never shows a
`MsgBox`. It refuses a worksheet caller, because its worksheet suites add and
close a scratch workbook and it writes a file.

From the Immediate window:

```vb
? KPR_Test_RunAll("0123456789abcdef0123456789abcdef01234567", "C:\kpr-evidence\run-1")
? KPR_Test_RunSuite("fixtures", "0123456789abcdef0123456789abcdef01234567", "C:\kpr-evidence\run-2")
```

Replace the example SHA with the exact candidate you imported.

## Suite registry

`TestRegistry` in the runner module is the single ordered registry;
`KPR_Test_RunAll` runs it in order and the validator reads it from source.

| Suite | Kind | Contents |
| --- | --- | --- |
| `date-type` … `parity` | pure | The twelve migrated suites, in `KPR_Tests_RunAll("all")` order |
| `fixtures` | fixture | Every direct-VBA case from `KPR_Test_Fixtures_Generated` (#38), one assertion per case |
| `worksheet-host` | worksheet | `KPR_Tests_RunHost`: 1900/1904 worksheet caller and `HostDateSystem()` refresh |
| `worksheet-shape` | worksheet | `KPR_Tests_RunShape`: Range orientation, multi-area, blanks, errors, `UsedRange` independence |
| `worksheet-array` | worksheet | `KPR_Tests_RunArray`: dynamic-array spill and 1904 call-level `#N/A` |
| `worksheet-fixtures` | worksheet | `KPR_Tests_RunFixtureHost`: the generated cases whose context is a 1900 or 1904 worksheet caller |
| `worksheet-oracle` | worksheet | `KPR_Tests_RunOracle`: the Excel cross-oracle cases in `KPR_Test_Oracle` (#41) |
| `worksheet-registration` | worksheet | `KPR_Tests_RunRegistration`: the MacroOptions manifest and register / clean-up lifecycle in `KPR_REGISTER_PUBLIC_UDFS` (#42) |
| `worksheet-state` | worksheet | `KPR_Tests_RunStateCheck`: runs the durable runner on `date-type` twice, once with a deliberately injected failure, and proves the caller's state is restored after both |

`KPR_Tests_RunAll("all")`, `KPR_Tests_RunEvidence` and the
[Excel host-evidence policy](EXCEL_EVIDENCE.md) keep the migrated twelve-suite
baseline unchanged. `fixtures` can also be run alone through
`KPR_Tests_RunSuite "fixtures"`.

Direct calls use the documented direct-VBA 1900 contract with no worksheet
caller. The fixture suites compare results with one typed assertion covering
values, dates and serials, Booleans, strings, native Excel error codes and
array shape and content. Direct-VBA results must have the exact semantic
subtype. A worksheet result carries no VBA subtype, so worksheet `Long` and
`Date` results compare by numeric value. A failed assertion records its case
and the run continues. An unexpected runtime error is recorded against the
suite it interrupted, and the remaining suites are marked `NOT_RUN`.

## Regression matrix (#40)

| #40 criterion | Where it is asserted |
| --- | --- |
| Positive, edge and invalid-domain cases for every value-taking function | Generated `matrix` cases: for each of the 21 functions a valid call, and for each value argument its lowest and highest supported values and two invalid-domain values |
| Scalar, 1x1 and array elements compared | `matrix` `array-1x1` cases equal the scalar call; row, column and rectangle cases reuse the same element values; the migrated `parity` suite |
| Row, column and rectangle orientation asserted | Every array expectation is a 1-based two-dimensional array of the exact shape |
| Mixed arrays keep valid neighbours | `matrix` `array-row` and `array-rectangle` cases mix valid, invalid and error elements |
| `#VALUE!`, `#NUM!`, host `#N/A` and propagated errors, scalar and array | `matrix` invalid-domain, `propagated`, array and `ws1904-array` cases; `propagation` and `shape` fixtures |
| Host and propagated `#N/A` provenance | `host` fixtures keep separate case IDs and conditions for the same Excel value |
| Direct VBA under the 1900 contract | The `fixtures` suite calls with no worksheet caller |
| 21 functions succeed in 1900 and return one call-level `#N/A` in 1904 | `matrix` `ws1900`, `ws1904` and `ws1904-array` cases, replayed by `worksheet-fixtures` |
| `HostDateSystem()` 1900/1904 and ordinary-recalculation refresh | `host` fixtures and `worksheet-host` |
| State restoration after passing and failing cases | `worksheet-state` |
| Two consecutive runs identical | Two `KPR_Test_RunAll` records compared with `check_test_evidence.py --compare` |
| No `_Spill`, calendar, weekend-mask, holiday or business-day case | Fixtures may call only the 22 contract names (`gen_fixtures.py`), and the public surface is pinned by `check_kpr_contract.py` |

`gen_fixtures.py` fails generation if any matrix case is missing for any
value-taking function or value argument.

## Excel cross-oracle (#41)

`tests/modules/KPR_Test_Oracle.bas` compares the date primitives with native
Excel worksheet functions only where the two contracts overlap. It evaluates
`EOMONTH`, `EDATE`, `WEEKDAY`, `DAY`, `YEAR` and `MONTH` on a scratch workbook
whose date system is 1900. `check_kpr_contract.py` (`kpr-oracle-scope`)
rejects any other worksheet function in an oracle formula, including
`WORKDAY.INTL` and `NETWORKDAYS.INTL`. It recognises `Xl` calls in any letter
case and requires every formula to reach Excel through `Xl`, so no other
procedure may call `Evaluate`, `ExecuteExcel4Macro`, bracket evaluation
(`[...]`) or `WorksheetFunction`, touch cells, names or their formulas and
values, recalculate, or run a macro. Outside `Xl` the module is also held to an
allowlist: every bare name is declared in the module, a label, a `KPR_Dates_`
call or a listed VBA or Excel name, and every member access is one of the
few the scratch workbook needs, so `ActiveCell`, `Selection` and similar
implicit cell writes fail. Excel object variables and `Application` may appear
only in the few forms that open, set up and close the scratch workbook, so no
default-member call such as `mSheet("A1") = ...` can write a cell, and an Excel
object may be stored only in a variable declared with an Excel object type,
never in a Variant alias. The module may declare only numeric, `String`,
`Boolean`, `Date`, `Variant`, `Collection`, `Workbook`, `Worksheet` and
`XlCalculation` values, and `KPR_Oracle_RunCases(ByRef Checks As Long, ByVal
Failures As Collection)` is its only public procedure, so no caller can hand
it an `Application` or other Excel object. The failure `Collection` is held to
the same rule: the oracle may only `.Add` to it, so nothing can be read back out
of it. `Xl` itself holds only
`Xl = mSheet.Evaluate(Formula)`, with `Formula` a required `ByVal` String. Each `Xl` formula must be built only from string literals and
`CStr` of arithmetic over numeric literals and variables declared with a
numeric type in scope (a local declaration shadows a module one, and an untyped
name is a Variant). No text variable reaches a formula, and the formula text may
hold only numbers, `+ - * /`, commas, parentheses and the permitted functions,
so no defined name or cell reference can hide another formula.

Every case label states the identity it tests, for example
`EndOfMonth(d) = EOMONTH(d,0)`. The identities cover month, quarter and year
boundaries and predicates, `DaysInMonth`, leap years and `DaysInYear`,
`DayOfWeek` in both weekday bases, `AddMonths` and `AddYears` against `EDATE`,
`EndOfMonth` of `AddMonths` against `EOMONTH`, `AddDays` and `AddWeeks`
against serial arithmetic, and both weekday locators.

Samples are 22 boundary dates (window edges, month ends, leap days and
century years), each with a forward and a backward parameter set, followed by
300 samples from a Park-Miller generator with a fixed seed. Case order and
values are identical on every run: 344 samples of 21 assertions, 7,224 in all.

Documented exclusions are asserted, never skipped:

- where Excel's answer lies outside the supported window, or Excel refuses,
  the contract requires `#NUM!` and the case asserts `#NUM!`;
- Excel's February 1900 has a fictitious 29th day, so `IsLeapYear(1900)` and
  `DaysInYear(1900)` are asserted against the Gregorian `FALSE` and 365;
- across that fictitious day Excel's `EOMONTH(d,-1)+1` does not give
  1-Mar-1900, so `BeginOfMonth` is asserted against `d-DAY(d)+1` for March
  1900, and the weekday locators always take the month start as `d-DAY(d)+1`;
- text or coerced inputs, fractional serials, serials before 1900-03-01 and
  the 1904 date system are not compared, because Excel's behaviour there is
  not the KPR contract.

## MacroOptions registration (#42)

`src/modules/KPR_REGISTER_PUBLIC_UDFS.bas` registers the Function Wizard
description, category and argument help of the 22 supported functions. It is
unsupported infrastructure (`Option Private Module`, role `internal`) and
contains no date algorithm. Its manifest is one `AddRecord` statement per
function: name, argument names in signature order, description and one
argument description per argument. `check_kpr_contract.py`
(`kpr-registration-manifest`) reads the records as data and requires exactly
one record per supported function, argument names that match the public
signature, complete argument descriptions that begin with the argument name
(optional ones with `(optional, default ...)`), descriptions within Excel's
255-character limit that state scalar and dynamic-array behaviour, 1-based
argument arrays, the single category `KPR Dates`, and no MacroOptions call in
any other module. Its JSON report carries the parsed manifest as
`registration_manifest` (name and arguments only) for later completeness
checks.

The `worksheet-registration` suite checks the manifest shape at run time, then
calls the workbook-qualified entry points through `Application.Run`: register
twice, clean up twice and register again. Every call must report all 22
functions with an empty failure report, and `ThisWorkbook.Saved`, calculation,
events, screen updating, alerts, the active workbook and sheet and the
selection must be exactly as found. Excel cannot read MacroOptions metadata
back, so repeatability is shown by identical results rather than by reading
the registered text. The suite leaves the functions registered. It has 86
assertions, so the #42 `all` run had 20 suites and 9,263 assertions. The
near-integer tolerance added five generated fixtures and eight `integer`
assertions, so an `all` run now has 20 suites and 9,276 assertions.

Excel has no operation that unregisters a VBA function. Clean-up blanks each
description and argument description and moves the function to Excel's
built-in "User Defined" category; the functions stay callable while the
project is loaded.

## Caller-state restoration

Before any suite runs, the runner captures the calculation mode, events,
screen updating, alerts, the status bar, the active workbook, the active sheet
and the selection. During the run it suppresses events, screen updating and
alerts, uses manual calculation and shows its progress in the status bar.
On every exit path, including a runtime error, it restores the captured state
and then reads each item back. A Range selection is compared by address; a
chart, shape or other selected object is compared by type and, when it has
one, by name. A failed activation or re-selection, or any item that does not
match, is recorded in `state_restoration` and fails the run. Each worksheet suite closes its own
scratch workbook with `SaveChanges:=False`.

## Evidence record

The runner writes one ASCII JSON document. Every non-ASCII character is written
as a JSON `\u` escape.

| Field | Contents |
| --- | --- |
| `schema`, `schema_version` | `kpr-test-evidence`, `1` |
| `source_sha` | The supplied exact commit SHA |
| `runner` | Module, entry point, selection (`all` or one suite), fixture TSV SHA-256 and fixture case count |
| `environment` | Excel version and build, Office bitness, VBA version, operating system, locale (country, separators, date order), harness workbook date system and caller context |
| `timing` | Local start and finish time and total elapsed milliseconds |
| `nondeterministic_fields` | `/environment`, `/timing`, `/suites/*/elapsed_ms` |
| `suites` | Per suite: name, kind, status (`PASS`, `FAIL`, `NOT_RUN`), assertions, failures, elapsed milliseconds |
| `failures` | Every failing case: suite, case label and detail |
| `totals` | Suite, assertion and failure totals |
| `state_restoration` | `PASS` or `FAIL` with the items that did not restore |
| `certification` | The nine #52 outcomes below |
| `result` | `PASS` only when every suite passed, no failure is listed and state was restored |

A case is one assertion. Passing assertions are counted per suite; each failing
assertion is listed with the label that identifies it. A suite that executes no
assertion fails with the case `<suite>/runner`. For the fixture suites,
that label is the generated fixture ID.

### Determinism

Two runs of the same source on the same host must produce records that are
identical after removing the fields listed in `nondeterministic_fields`.
Suite order, case order, counts, failure labels and failure details are
deterministic. Failure details render numbers and dates without the host's
locale separators.

### Certification outcomes

The record carries every structured #52 certification outcome, so the
certification record is this record and not a separate ad hoc document:

| Key | Outcome |
| --- | --- |
| `source_import` | Import of every production, test and demo export |
| `vba_compile` | VBA project compilation |
| `regression` | This run; written by the runner |
| `cross_oracle` | Excel cross-oracle suite (#41) |
| `macro_options` | MacroOptions registration and cleanup (#42) |
| `ribbonx` | RibbonX lifecycle |
| `commandbars` | CommandBars lifecycle |
| `demo_generation` | Deterministic demo generation |
| `source_round_trip` | Normalized VBA export round trip |

Each outcome has a `status` of `PASS`, `FAIL`, `NOT_RUN` or `NOT_APPLICABLE`.
`PASS` and `FAIL` carry a nonempty `detail` and a null `reason`. `NOT_RUN` and
`NOT_APPLICABLE` carry a nonempty `reason` and a null `detail`.

The runner writes `regression` from the run itself, `cross_oracle` from the
`worksheet-oracle` suite and `macro_options` from the `worksheet-registration`
suite: `PASS` or `FAIL` when that suite ran, `NOT_RUN` when it was not
selected. Every other outcome is `NOT_RUN`, because the runner
cannot observe it. The certification operator
completes those outcomes in the same file after observing them, and never
edits a field the runner wrote.

## Validation

```bash
python3 tools/check_test_evidence.py --root . --self-test
python3 tools/check_test_evidence.py --root . \
  --evidence ../kpr-evidence/run-1/kpr-test-evidence.json \
  --candidate-sha FULL_CANDIDATE_SHA
python3 tools/check_test_evidence.py --root . \
  --evidence ../kpr-evidence/run-1/kpr-test-evidence.json \
  --compare ../kpr-evidence/run-2/kpr-test-evidence.json
python3 tools/check_test_evidence.py --root . \
  --evidence ../kpr-evidence/certification/kpr-test-evidence.json \
  --candidate-sha FULL_CANDIDATE_SHA --certification
```

Run the validator from a checkout of the same candidate. It checks the record
against the schema, then:

- the source SHA matches `--candidate-sha` when given;
- the fixture SHA-256 and case count match `tests/fixtures/date_layer_fixtures.tsv`;
- an `all` selection lists the complete `TestRegistry` in order, and a single
  selection lists exactly that suite;
- both timestamps are real local dates and times, in order;
- suite kinds, statuses, failure counts and totals agree with the failure list,
  and a passing suite executed at least one assertion;
- `result` and the `regression` outcome follow from the suites, failures and
  state restoration;
- every certification outcome follows the detail and reason rules;
  `cross_oracle` matches the `worksheet-oracle` suite and `macro_options` the
  `worksheet-registration` suite, each `NOT_RUN` when its suite was not
  selected;
- with `--compare`, the second record passes the same schema and semantic
  checks, and both records are identical outside the declared
  nondeterministic fields;
- with `--certification`, the record is a `KPR_Test_RunAll` record and every
  outcome is `PASS`. `ribbonx`, `commandbars`, `demo_generation` and
  `source_round_trip` may instead be `NOT_APPLICABLE`.

Exit 0 means a valid record of a passing run. Exit 1 means an invalid record
or a valid record of a failing run. Exit 2 means a usage or input error.

Keep run output outside tracked source. Retain the evidence folder with the
issue or release evidence it supports, and preserve failing records until the
finding is resolved.
