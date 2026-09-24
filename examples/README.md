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
`modules/KPR_DateExample.bas` and run `KPR_DateExample.RunDateExample`.
It calls the supported `KPR_Dates_DaysInMonth` entry point from direct VBA with
an ISO date and writes the result to the Immediate window. Direct VBA follows
the frozen no-worksheet-host 1900 serial contract; worksheet callers remain
subject to the caller workbook's date-system rules.

The example is intentionally minimal. The planned full demo remains outside
v0.0.2 and is carried forward through the migration backlog.

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

Delete this README only if real examples and equivalent instructions make the directory's role equally explicit.
