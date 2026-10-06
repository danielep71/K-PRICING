# 💡 Examples

`examples/` contains reproducible, non-production examples that teach the supported API and can be rebuilt from committed source.

Appropriate contents include:

- example VBA modules;
- source-controlled demo builders;
- synthetic input data;
- minimal integration snippets; and
- instructions for creating an example `.xlsm` from the tagged source.

Examples must use supported behavior, synthetic or redistributable data, and the same installation path documented for users. They must not become an undocumented second implementation or a substitute for regression tests.

## Minimal migrated example

After importing the production modules, optionally import
`modules/KPR_Demo_DirectVBA.bas` and run `KPR_Demo_DirectVBA.RunDateExample`.
It calls the supported `KPR_Dates_DaysInMonth` entry point from direct VBA with
an ISO date and writes the result to the Immediate window. Direct VBA follows
the frozen no-worksheet-host 1900 serial contract; worksheet callers remain
subject to the caller workbook's date-system rules.

The example is intentionally minimal.

## Demonstration sheets

Demo sheets are built from tracked source by a shared builder; no generated
workbook is committed.

- `modules/KPR__Demo_Builder.bas` is the shared engine: layout, house style,
  status rules, the checks summary and the Excel state around a build. It also
  holds the catalog of demos (`KPR_Demo_Catalog`).
- `modules/KPR_Demo_Dates.bas` is the Date Primitives demo, a short list of
  builder calls.

Import both after the production modules and run, in the Immediate window:

```text
? KPR_Demo_BuildDates()
```

It returns `TRUE` and adds a **KPR Dates Demo** sheet to the active workbook,
named `KPR Dates Demo (2)`, `(3)` … when the name is taken. No existing sheet
is changed, replaced or deleted. When no ordinary workbook is active, or it is
a 1904 workbook, read-only or structure-protected, the sheet goes into a new
unsaved workbook instead. Ctrl+Z cannot remove a sheet a macro added: delete
the sheet to remove the demo. `? KPR_Demo_BuildDates("C:\path\demo.xlsx")`
builds into a new `.xlsx` file instead and never overwrites an existing file;
`? KPR_Demo_LastReport()` explains a `FALSE`.

The sheet has editable orange inputs (DateIn, nDays, nWeeks, nMonths, nYears,
Keep EOM, weekday index, occurrence and weekday base), each a sheet-level name
the formulas use. Each section row shows the KPR formula as text, its live
result, a native-Excel reference formula that reproduces the contract for any
valid input, and a status:

| Status | Meaning |
| --- | --- |
| `OK` | The KPR result equals the reference, or both are the same error |
| `DIFFERS` | A documented difference by design, explained in the row's note |
| `FAIL` | Any other difference |
| `NO VBA` | The K-PRICING project is not loaded |

Sections cover day primitives, period boundaries, predicates, date
arithmetic, the month-end policy against `EDATE` and `EOMONTH`, weekday
locators and pillars. The side panel holds the checks summary, fixed
native-error cases, dynamic-array spills and the 100,000-element limit (the
last two need dynamic-array Excel). At the default inputs the sheet shows one
`DIFFERS` row and no `FAIL`. The live results need the K-PRICING project to be
open, because the formulas call its functions; when it is not an add-in the
formulas carry its workbook name. The regression harness's `worksheet-demo`
suite rebuilds the demo and checks it. The sheet demonstrates documented
behaviour; it does not certify accuracy, production readiness or any untested
environment.

### Adding a demo

Write a content module with one public entry point
`Function KPR_Demo_Build<Name>(Optional ByVal OutputPath As String = "") As Boolean`
that calls `Demo_Begin`, the writers (`Demo_Inputs`, `Demo_Input`,
`Demo_Section`, `Demo_Compare`, `Demo_Note`, `Demo_Cases`, `Demo_Case`,
`Demo_Spill`, `Demo_SpillColumn`) and `Demo_Finish`, with `Demo_Abort` in its
error handler, and add one line to `KPR_Demo_Catalog`. A feature moves into
the builder only when a second demo needs it.

## Minimal worksheet example

This synthetic example uses the same imported production modules in a new,
blank macro-enabled workbook that uses the default **1900** date system. The
expected results are taken from
[`docs/DATE_LAYER_CONTRACT.md`](../docs/DATE_LAYER_CONTRACT.md). The
[issue #17](https://github.com/danielep71/K-PRICING/issues/17) parity run
exercised destination worksheet calls (date-system, Range-shape and
dynamic-array probes) on one Windows 64-bit Excel host, but it did not execute
these exact formulas. Their results below are therefore contract expectations,
not recorded destination runs.

Date-valued results are VBA `Date` values; apply a date number format if a cell
shows a serial number. Date arguments may be real dates, date serials or text
in exactly `YYYY-MM-DD` form.

| Formula | Expected result | Contract rule shown |
| --- | --- | --- |
| `=KPR_Dates_DaysInMonth("2026-02-15")` | `28` | ISO text input |
| `=KPR_Dates_EndOfMonth("2024-02-10")` | 2024-02-29 | Gregorian month end |
| `=KPR_Dates_AddMonths("2024-02-29", 1)` | 2024-03-29 | Clip mode keeps the day when it exists |
| `=KPR_Dates_AddMonths("2024-02-29", 1, TRUE)` | 2024-03-31 | `Opt_KeepEOM` maps month end to month end |
| `=KPR_Dates_DateFromPillar("2024-01-31", "1M")` | 2024-02-29 | Pillar month delta with clip semantics |
| `=KPR_Dates_PillarFromDates("2024-01-15", "2024-04-15")` | `3M` | Canonical emitted pillar |
| `=KPR_Dates_IsLeapYear(1900)` | `FALSE` | Year-taking function; Gregorian rule |
| `=KPR_Dates_DaysInMonth("31/12/2026")` | `#VALUE!` | Locale-formatted text is rejected |
| `=KPR_Dates_HostDateSystem()` | `1900` | Volatile date-system diagnostic |

### Dynamic arrays

On dynamic-array Excel, put three dates in `A2:A4` and enter:

```text
=KPR_Dates_EndOfMonth(A2:A4)      spills a 3x1 result
=KPR_Dates_AddMonths(A2:A4, 1)    the scalar 1 expands to the 3x1 shape
```

Each invalid element returns its own error without affecting its neighbours.
Non-scalar arguments of different shapes return one `#VALUE!`, and a result of
more than 100,000 elements returns one `#NUM!`. Multi-cell results are claimed
only on dynamic-array Excel; no Ctrl+Shift+Enter behavior is claimed.

### Date-system boundary

The date system comes only from the workbook that contains the calling cell.
In a workbook using the **1904** date system, every value-taking function
returns `#N/A` rather than a date shifted by 1,462 days, and
`=KPR_Dates_HostDateSystem()` returns `1904`. Direct VBA calls, including
`RunDateExample`, have no worksheet caller and use the 1900 serial contract.

Use `demo/` instead only when an interactive demo is itself a distinct project deliverable, an established public path must remain stable, or packaging automation requires that profile. Document the reason in the root README and do not maintain both `examples/` and `demo/` for the same purpose.

Do not commit opaque generated workbooks here unless the repository's release policy explicitly treats them as reviewed source artifacts. Published binaries normally belong to GitHub Releases.
