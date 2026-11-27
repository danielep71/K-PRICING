# 📘 K-PRICING user manual, version 0.2.0

This manual is for you if you received the K-PRICING workbook and want to use
its 22 date functions in your own worksheets. It describes the workbook of
this version and nothing else. It needs no knowledge of VBA or of the
repository. The links to `docs/` files lead to the repository's methodology,
which holds the exact rules; nothing here depends on them.

- [What the workbook is](#what-the-workbook-is)
- [Open it and enable macros](#open-it-and-enable-macros)
- [Check that it works](#check-that-it-works)
- [Use a function](#use-a-function)
- [The 22 functions](#the-22-functions)
- [Error values](#error-values)
- [The Ribbon tab and the toolbar](#the-ribbon-tab-and-the-toolbar)
- [The Dates Demo sheet](#the-dates-demo-sheet)
- [Limits](#limits)
- [Troubleshooting](#troubleshooting)
- [Report a problem](#report-a-problem)

## What the workbook is

An Excel workbook with macros (`.xlsm`) that holds 22 worksheet functions
named `KPR_Dates_*`. They work on calendar dates of the Gregorian calendar:
the weekday of a date, the first and last day of its month, quarter or year,
the date a number of days, weeks, months or years later, the nth weekday of a
month, and curve pillars such as `3M`. Every function returns a value or an
Excel error value. None of them changes a sheet: they read their arguments and
the date system of the workbook that calls them.

The workbook has two sheets:

| Sheet | What it holds |
| --- | --- |
| `Home` | The build stamp, what the workbook provides and does not cover, how to use it, and a link to each sheet |
| `Dates Demo` | A worked example of each function on inputs you can edit, beside Excel's own formulas, with a status per row |

It also adds a **K-PRICING** tab to the Ribbon, described
[below](#the-ribbon-tab-and-the-toolbar).

The functions live in this workbook. They work in a worksheet of this
workbook while it is open with macros enabled; a function written in another
workbook shows `#NAME?` there.

**Supported host.** Microsoft 365 Excel on Windows, desktop version, 64-bit.
Mac, Excel for the web and older perpetual versions are not supported. The
[installation guide](../../../INSTALLATION.md#%EF%B8%8F-host-policy) records
the environments on which this code was run.

## Open it and enable macros

Nothing here needs administrator rights, a lower macro security level, or a
change to Excel's *Trust Center*. If your organization's policy prevents a
step, the workbook does not run on your machine: ask your administrator
rather than work around the policy.

1. **Save the file** to a folder you can write to, such as *Documents*.
   Download or copy the file once; do not run it from an e-mail or a browser
   preview.
2. **Check it** if you received a checksum with it. In Windows PowerShell:

   ```powershell
   Get-FileHash .\FILE.xlsm -Algorithm SHA256
   ```

   Write your file's name in place of `FILE.xlsm`. The hash must equal the
   one you were given. If it differs, do not open
   the file.
3. **Unblock it only if your policy allows.** A file that came from the
   internet or another computer can carry a mark that makes Excel block its
   macros. If Excel says so, open the file's **Properties** in Explorer and
   tick **Unblock** where your organization permits it; otherwise ask your
   administrator.
4. **Open it** in Excel and choose **Enable Content** in the yellow bar. The
   macro setting can stay at *Disable all macros with notification*: you
   enable this one workbook, not every macro.
5. The **K-PRICING** tab appears in the Ribbon. If it does not, close the
   file and open it again with macros enabled.

The workbook contains no tests and no examples of its own code; a build for
developers does, and also shows a *Development* group in the tab.

## Check that it works

1. Open `Home`. The **Commit**, **Commit time** and **Configuration** cells
   are the build stamp: the revision of the source this workbook was built
   from. Note them when you report a problem. A configuration of
   `production` is the one to use; a stamp that says *NOT evidence-grade*
   marks a workbook built with local changes.
2. In any cell of the workbook, type `=KPR_Dates_HostDateSystem()`. The
   result is `1900`. A result of `1904` means that the workbook uses the 1904
   date system, which the date functions refuse; see
   [Troubleshooting](#troubleshooting).
3. Open `Dates Demo`. In the checks panel, the count of `FAIL` statuses is
   `0`. A `NO VBA` status means that macros are not enabled.

## Use a function

Type `=KPR_Dates_` in a cell and pick a function, or use the
[Ribbon](#the-ribbon-tab-and-the-toolbar). Examples, with fixed dates:

| Formula | Result |
| --- | --- |
| `=KPR_Dates_DayOfWeek("2026-11-27")` | `5`, a Friday, with Monday as 1 |
| `=KPR_Dates_EndOfMonth("2026-02-10")` | 28 February 2026 |
| `=KPR_Dates_AddMonths("2026-02-28", 1, TRUE)` | 31 March 2026 |
| `=KPR_Dates_NthWeekdayOfMonth(2026, 11, 4, 4)` | 26 November 2026, the fourth Thursday |
| `=KPR_Dates_DateFromPillar("2026-01-31", "1M")` | 28 February 2026 |
| `=KPR_Dates_PillarFromDates("2026-01-01", "2026-04-01")` | `3M` |

A date result is an Excel date: format the cell as a date if it shows a
number.

### How an argument is read

- **Dates** (`DateIn`, `StartDate`, `EndDate`). Give a cell holding a date, a
  number, which is read as an Excel serial under the 1900 date system, or
  text in exactly the form `YYYY-MM-DD`. A time of day is dropped. Text such
  as `31/12/2026`, `10.02.2024` or the text `45292` is refused, so that a result never
  depends on your regional settings. Dates run from 1 March 1900 to
  31 December 9999. A blank cell is refused.
- **Whole numbers** (`nDays`, `nWeeks`, `nMonths`, `nYears`, `YearIn`,
  `MonthIn`, `WdIndex`, `n`). A number with no fraction. `3` and `3.0000000001`
  are 3; `2.5` is refused, never rounded. Text, a Boolean and a date are
  refused. A year is the number `2026`, never a date: write `YEAR(A1)` when
  you hold a date.
- **Options** (`Opt_WeekBaseMonday`, `Opt_KeepEOM`, `Opt_Rounding`). Optional,
  always last. Leave one out, or leave it blank, to get its default.
  `Opt_WeekBaseMonday` and `Opt_KeepEOM` take `TRUE` or `FALSE` and nothing
  else: `1`, `0` and `"TRUE"` are refused. `Opt_Rounding` takes the text
  `NEAREST`, `FLOOR` or `CEILING`, in any letter case. An option is a single
  value, never a range.
- **Ranges and arrays.** Every function except `KPR_Dates_HostDateSystem`
  accepts a range or an array at each value argument, and returns a result of
  the same shape, which spills in Excel with dynamic arrays. A single value
  next to a range is repeated for each cell. Several ranges must have the
  same rows and columns. A result of more than 100,000 cells is refused. A
  cell that holds an error gives that error at its own position and does not
  stop the others.

## The 22 functions

Arguments in brackets are optional, with their default. *Date errors*,
*whole-number errors* and the others are defined under
[Error values](#error-values); the last column lists only what a function adds
to them.

### Weekday and lengths

| Function | Arguments | Result | Adds |
| --- | --- | --- | --- |
| `KPR_Dates_DayOfWeek` | `DateIn`, `[Opt_WeekBaseMonday=TRUE]` | 1 to 7. With `TRUE`, Monday is 1 and Sunday 7; with `FALSE`, Sunday is 1 and Saturday 7, as Excel's `WEEKDAY` | Date errors, option errors |
| `KPR_Dates_DaysInMonth` | `DateIn` | 28 to 31, the length of the date's month | Date errors |
| `KPR_Dates_DaysInYear` | `YearIn` | 365 or 366 | Whole-number errors; `#VALUE!` for a year outside 1900 to 9999 |
| `KPR_Dates_IsLeapYear` | `YearIn` | `TRUE` or `FALSE` | Same as `KPR_Dates_DaysInYear` |

The two year functions take a year, not a date, and build no date, so
`KPR_Dates_DaysInYear(1900)` is 365 and `KPR_Dates_IsLeapYear(1900)` is
`FALSE`.

### Period boundaries and period-end tests

Quarters are the calendar quarters, January to March and so on.

| Function | Arguments | Result | Adds |
| --- | --- | --- | --- |
| `KPR_Dates_BeginOfMonth` | `DateIn` | First day of the date's month | Date errors |
| `KPR_Dates_EndOfMonth` | `DateIn` | Last day of the date's month | Date errors |
| `KPR_Dates_BeginOfQuarter` | `DateIn` | First day of the date's quarter | Date errors; `#NUM!` for a date of the first quarter of 1900 |
| `KPR_Dates_EndOfQuarter` | `DateIn` | Last day of the date's quarter | Date errors |
| `KPR_Dates_BeginOfYear` | `DateIn` | First day of the date's year | Date errors; `#NUM!` for every date of 1900 |
| `KPR_Dates_EndOfYear` | `DateIn` | Last day of the date's year | Date errors |
| `KPR_Dates_IsMonthEnd` | `DateIn` | `TRUE` when the date is the last day of its month | Date errors |
| `KPR_Dates_IsQuarterEnd` | `DateIn` | `TRUE` when it is the last day of its quarter | Date errors |
| `KPR_Dates_IsYearEnd` | `DateIn` | `TRUE` when it is the last day of its year | Date errors |

The `#NUM!` of the two *Begin* functions is their result, 1 January 1900,
falling before the first supported date, 1 March 1900.

### Date shifts

| Function | Arguments | Result | Adds |
| --- | --- | --- | --- |
| `KPR_Dates_AddDays` | `DateIn`, `nDays` | The date `nDays` calendar days later; negative goes back | Whole-number errors; `#NUM!` when the result leaves 1 March 1900 to 31 December 9999 |
| `KPR_Dates_AddWeeks` | `DateIn`, `nWeeks` | The date seven times `nWeeks` days later | Same |
| `KPR_Dates_AddMonths` | `DateIn`, `nMonths`, `[Opt_KeepEOM=FALSE]` | The same day `nMonths` later; when that month is shorter, its last day | Same, and option errors |
| `KPR_Dates_AddYears` | `DateIn`, `nYears`, `[Opt_KeepEOM=FALSE]` | `KPR_Dates_AddMonths` by twelve times `nYears` | Same |

With `Opt_KeepEOM` set to `TRUE`, the last day of a month moves to the last
day of the target month: 28 February 2026 plus one month is 31 March 2026,
where `FALSE` gives 28 March 2026. Adding a month and then taking it away does
not always return the first date: `KPR_Dates_AddMonths` clips 31 January to
28 February, and going back gives 28 January. These are calendar days only:
there is no business day, holiday or roll convention.

### Weekdays of a month

`WdIndex` numbers the weekday as `KPR_Dates_DayOfWeek` does under the same
`Opt_WeekBaseMonday`: with the default, Monday is 1 and Sunday 7.

| Function | Arguments | Result | Adds |
| --- | --- | --- | --- |
| `KPR_Dates_NthWeekdayOfMonth` | `YearIn`, `MonthIn`, `WdIndex`, `n`, `[Opt_WeekBaseMonday=TRUE]` | The date of occurrence `n` (1 to 5) of that weekday in the month | Whole-number errors; `#VALUE!` for a month outside 1 to 12, a weekday outside 1 to 7 or an `n` outside 1 to 5; `#NUM!` when the month has no such occurrence, such as a fifth Monday of February 2026 |
| `KPR_Dates_LastWeekdayOfMonth` | `YearIn`, `MonthIn`, `WdIndex`, `[Opt_WeekBaseMonday=TRUE]` | The date of the last occurrence of that weekday | Same, without `n` |

A year before 1900 or a result before 1 March 1900 is refused.

### Pillars

A pillar is a short text for an interval, such as `3M` or `1Y6M`.

| Function | Arguments | Result | Adds |
| --- | --- | --- | --- |
| `KPR_Dates_DateFromPillar` | `StartDate`, `Pillar` | The date the pillar names from `StartDate` | Date errors; `#VALUE!` for a malformed pillar; `#NUM!` when the shift or the result is out of range |
| `KPR_Dates_PillarFromDates` | `StartDate`, `EndDate`, `[Opt_Rounding="NEAREST"]` | The pillar that labels the interval, as text | Date errors, option errors; `#NUM!` when `CEILING` has no bucket that reaches `EndDate` |

**Writing a pillar.** One to four components, each a whole number and a unit,
each unit used at most once: `Y` years, `M` months, `W` weeks, `D` days. They
may come in any order, `2W3D` and `3D2W` being the same. A leading `+` or `-`
applies to the whole token. The aliases `ON` and `O/N` mean `1D`, and `TN` and
`T/N` mean `2D`, always calendar days; an alias stands alone and takes no sign.
Letter case is ignored. A blank inside, a decimal or a repeated unit, such as
`1 M` or `1M2M`, is refused.

**How a date moves.** The years and months are applied first, with the clip of
`KPR_Dates_AddMonths`, then the weeks and days exactly: `1M` from 31 January
2026 is 28 February 2026.

**How an interval is labelled.** Equal dates give `0D`. A gap of less than
seven days gives its signed day count, such as `5D` or `-3D`. A longer one
gives `1W`, `2W`, `3W`, a number of months, or years and months (`11M`, `1Y`,
`1Y6M`), never `4W` or more. `NEAREST` takes the closest of these anchors to
`EndDate`, and a tie goes to the month; `FLOOR` takes the largest that does not
pass `EndDate`; `CEILING` the smallest that reaches it. The label rounds, so
`KPR_Dates_DateFromPillar` of the label gives `EndDate` only when the gap is
an exact bucket.

### Host date system

| Function | Arguments | Result | Adds |
| --- | --- | --- | --- |
| `KPR_Dates_HostDateSystem` | None | `1900` or `1904`, the date system of the workbook that calls it | `#N/A` when the date system cannot be read |

This is the only function that recalculates every time Excel does, so that it
follows a change of the workbook's date system. It accepts no array.

## Error values

A function returns an Excel error value, never a message or a guessed result.

| Error | Means |
| --- | --- |
| `#VALUE!` | The input cannot be read or breaks a rule: text that is not exact `YYYY-MM-DD`, a date that does not exist, a blank, a rejected type, a fraction where a whole number is required, a value outside an argument's range, a bad option or pillar, ranges of different sizes, or an unsupported shape |
| `#NUM!` | The input is well formed but its value or result is outside what is supported: a date outside 1 March 1900 to 31 December 9999, a whole number beyond about ±2.1 billion, a shift whose result leaves that window, a fifth weekday the month does not have, or a result of more than 100,000 cells |
| `#N/A` | The workbook is on the 1904 date system, or its date system cannot be read |
| Any error you pass in | Returned unchanged |

By argument:

| Argument | Error | When |
| --- | --- | --- |
| A date | `#VALUE!` | Text that is not exact `YYYY-MM-DD`, with locale-shaped text, number-shaped text and a date that does not exist such as `2025-02-29` all refused; a blank; a Boolean |
| A date | `#NUM!` | A date or serial outside 1 March 1900 to 31 December 9999 |
| A whole number | `#VALUE!` | A fraction, text, a Boolean, a date or a blank; `YearIn` outside 1900 to 9999, `MonthIn` outside 1 to 12, `WdIndex` outside 1 to 7, `n` outside 1 to 5 |
| A whole number | `#NUM!` | A number beyond the range of a 32-bit integer |
| An option | `#VALUE!` | The wrong type or a value that is not one of its tokens; more than one cell; an unsupported shape. The whole call fails |
| Any | `#N/A` | The calling workbook is on the 1904 date system: the date functions return `#N/A` for the whole call |

An element-level error leaves the other cells of an array result alone. An
option, shape or capacity error, and the date system, fail the whole call.

An `#N/A` that the library returns looks the same as one you passed in. To tell
them apart, put `=KPR_Dates_HostDateSystem()` in the same workbook: `1904`
shows that the workbook refuses the date functions.

The [data contract](../../methodology/DATA_CONTRACT.md#conditions) names every
condition behind these errors with a stable identifier, which the test
evidence quotes.

## The Ribbon tab and the toolbar

The **K-PRICING** tab appears when the workbook is open with macros enabled.

| Control | What it does |
| --- | --- |
| **Functions** (*Date functions* group) | A menu of the 22 functions. Choosing one puts `=<function>()` in the active cell and opens Excel's **Function Arguments** dialog for it. Cancelling the dialog gives the cell back its content. It refuses a cell inside an Excel table, where Excel would copy the formula to the whole column, and a locked cell on a protected sheet |
| **Register function help** (*Setup* group) | Shows the 22 functions in the **Insert Function** dialog under the *KPR Dates* category, with a description of each argument. Excel keeps this until it closes |
| **Clear function help** (*Setup* group) | Returns the 22 functions to the *User Defined* category with no description. They stay callable |
| **Run tests** (*Development* group) | Only in a development build: runs the project's test suites. It is not an evidence run |

A temporary **KPR Dates** toolbar offers the same actions. Press **Alt+F8**,
select `InstallToolbar` and choose **Run** to show it, and `RemoveToolbar` to
hide it. Excel discards the toolbar when it closes.

Each action reports a failure in a message and leaves Excel's own settings, such
as screen updating and calculation, as it found them.

## The Dates Demo sheet

`Dates Demo` shows each function at work on inputs you can change:

- **Inputs.** The orange cells at the top of the main column, one for each
  argument name, are the only cells you can edit. Change `DateIn`, `nDays` and
  the others and the whole sheet follows.
- **Rows.** Each row shows the formula as text, the same formula live, a
  formula of Excel's own functions that should give the same result, a status
  and a note.
- **Status.** `OK`: the function equals the reference, or both are the same
  error. `DIFFERS`: they differ in a case the note explains as a difference by
  design. `NO VBA`: the functions are not loaded, because macros are disabled.
  `FAIL`: anything else; none should appear.
- **Side panel.** It counts the statuses, lists inputs that must give an error
  value with the error expected, spills one formula over 30 dates, and shows the
  limits of arrays.

The spill and the array limits need an Excel with dynamic arrays; elsewhere
their statuses are `FAIL`. The sheet calculates when the workbook opens.

## Limits

- The date functions cover Gregorian dates from 1 March 1900, the first date on
  which Excel's 1900 system agrees with the real calendar, to 31 December 9999.
- A workbook on the 1904 date system is refused, never read with its serials
  shifted by 1,462 days.
- The workbook does not cover business calendars and holidays, business-day
  arithmetic, day-count conventions, schedules, valuation or pricing. Later
  versions are planned for some of them; none is provided here.
- It performs no calculation on your data apart from the functions above, and
  it stores none.
- The results follow the repository's calculation contracts. Check them
  against your own requirements before you rely on them.

## Troubleshooting

| You see | Check |
| --- | --- |
| `#NAME?` for a `KPR_Dates_` function | Macros are disabled, or the formula is in another workbook. Enable content in this workbook |
| `#N/A` from every function that has a value argument | The workbook is on the 1904 date system: **File → Options → Advanced**, under *When calculating this workbook*, shows *Use 1904 date system* ticked. Excel does not convert the dates already in a workbook when you change it, so do not untick it to make the error go away; use the functions in a workbook on the 1900 date system |
| `#VALUE!` on a date you typed | Excel may have stored it as text in another form. Write it as `YYYY-MM-DD` text, or enter a real date |
| `#VALUE!` on `Opt_KeepEOM` or another option | It must be `TRUE` or `FALSE`, not `1` or text |
| A number instead of a date | Format the cell as a date |
| The K-PRICING tab is missing | The workbook was opened without macros. Close and reopen it with macros enabled |
| `Dates Demo` shows `NO VBA` | Macros are not enabled |
| Excel blocks the macros of a downloaded file | See [Open it](#open-it-and-enable-macros), step 3 |

## Report a problem

Open an issue on the
[K-PRICING repository](https://github.com/danielep71/K-PRICING/issues). Give:

- the build stamp from `Home`: the commit and the configuration;
- your Excel build, from **File → Account → About Excel**, and whether it is
  64-bit;
- the formula, with synthetic inputs, the result you got and the result you
  expected.

Never attach real positions, counterparty or client data.
