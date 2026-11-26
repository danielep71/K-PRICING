# 🧾 Data contract

This contract defines every input the workbook accepts: where it comes from,
its layout and units, what makes it invalid, and how an import is traced.
The [calculation contracts](CALCULATION_CONTRACTS.md) refer to its datasets
and fields by name.

The date layer reads no dataset: the arguments of its worksheet functions are
defined under [Function arguments](#function-arguments). The illustrative
dataset `DS-01` stays until the illustrative test cases that use it are
replaced ([Decisions](#decisions)).

## Datasets

This table and the field tables are rendered from the
[machine-readable contract](#machine-readable-contract).

<!-- BEGIN generated: data-contract-datasets -->
| ID | Description | Provider | Frequency | Granularity | File or sheet | Encoding |
| --- | --- | --- | --- | --- | --- | --- |
| *Illustrative:* `DS-01` | Example positions | Upstream system or manual input | Per reporting date | One row per position | `positions.csv` | `UTF-8` |
<!-- END generated: data-contract-datasets -->

A name with an extension is a file; one without is a sheet. *Encoding* is the
encoding the workbook reads a file in, as a Python codec name such as `UTF-8`
or `cp1252`. For a file, an empty cell (`null` in the contract) means UTF-8; a
sheet has no encoding, and its cell is always empty. The
[file-based test cases](TEST_CASES.md#file-based-cases) hold only these files.
Outside an `invalid` case, each decodes in its dataset's encoding; an
`invalid` case may hold one that does not, because that is what it tests.

## File format

State, for each delimited text file the workbook imports, its encoding,
delimiter and decimal separator in `data-contract.json`. The rest of the
format is fixed, so that the [reference validator](#reference-validator) and
the workbook read a file the same way. In outline:

- dates are ISO 8601, `YYYY-MM-DD`;
- the first row is a header that names each column by its field name, and
  the columns may come in any order;
- a missing value is an empty cell, kept distinct from zero;
- a cell may be quoted with double quotes.

The [reference validator](#reference-validator) section gives the complete
rules, including how numbers and booleans are written, the line endings, a
blank line and a byte order mark.

A provider file in another format is converted before it reaches the
workbook, and the conversion is recorded as a *Decision* below.

## Fields

One table per dataset, after the file or sheet it is read from:

<!-- BEGIN generated: data-contract-fields -->
### *Illustrative:* `DS-01` Example positions

Read from `positions.csv`, delimited text in `UTF-8`, with `,` between cells and `.` as the decimal separator.

| Field | Type | Unit or currency | Required | Allowed values or range | Notes |
| --- | --- | --- | --- | --- | --- |
| *Illustrative:* `position_id` | Text | — | Yes | Any | Identifies the position |
| *Illustrative:* `currency` | Currency code, ISO 4217 | — | Yes | Any | Currency of the position's amounts |
| *Illustrative:* `notional` | Decimal | Currency in field `currency` | Yes | ≥ 0 | Absolute amount, never negative |
<!-- END generated: data-contract-fields -->

A dataset's identifier field is its required `text` field whose name ends in
`_id`, such as the illustrative `position_id`. A dataset whose records a
[JSON fixture](TEST_CASES.md#fixture-format) lists has exactly one, and its
value is the record's `id`.

State the unit of every amount, rate and period explicitly: a percentage, a
decimal fraction and basis points are different fields, not different
formats.

<a id="machine-readable-contract"></a>

## Machine-readable contract

[`data-contract.json`](data-contract.json), beside this guide, holds the
datasets and their fields. Edit it, never the generated tables, then run
`python tools/render_docs.py --write`. **Repository integrity** fails while a
table differs from it, and `python tools/check_data_contract.py`, which it
also runs, checks the file itself.

| Key | Holds |
| --- | --- |
| `schema_version`, `kind` | `1` and `data-contract` |
| `datasets` | One object per dataset, holding the columns of the *Datasets* table as `id`, `illustrative` (`true` or `false`), `description`, `provider`, `frequency`, `granularity`, `source` (the file or sheet) and `encoding` (a Python codec name for a file, `null` for a sheet), then `delimiter`, `decimal_separator` and `fields` |
| `delimiter`, `decimal_separator` | For a delimited text file, the one character between cells, and `.` or `,`; both `null` for a sheet or for a file of another kind |
| `fields` | One object per field: `name` (lower snake case, unique in the dataset), `type`, `unit`, `required` (`true` or `false`), `allowed` and `notes` |
| `type` | `text`, `integer`, `decimal`, `date` (ISO `YYYY-MM-DD`), `boolean` (`true` or `false`) or `currency` (three upper-case letters, ISO 4217) |
| `unit` | `null`, a text such as `Decimal fraction` or `Years`, or `{"currency_field": "<name>"}` for an amount whose currency is in a `currency` field of the same dataset |
| `allowed` | `null`; `{"values": [...]}`, texts matched exactly, for a `text`, `integer`, `date` or `currency` field; or `{"min": x, "max": y}`, either bound optional and both inclusive, for an `integer` or `decimal` field |

The file is strict JSON, with no duplicate key and no non-finite number, and
a text shown in a table cell holds no `|` and no line break.

<a id="reference-validator"></a>

### Reference validator

`python tools/check_data_contract.py --dataset DS-NN FILE...` validates input
files of a delimited text dataset against the contract.
`tools/check_test_cases.py` runs the same validator on every variant file of
a [file-based case](TEST_CASES.md#file-based-cases), except in an `invalid`
case. It reads a file this way:

- The file decodes in the dataset's encoding. A byte order mark stays part of
  the first column's name unless the encoding, such as `utf-8-sig` or
  `UTF-16`, removes it.
- The first row is a header that names each column by its field name, in any
  order. Every required field has a column.
- Each further row has one cell per column; a blank line is a row with no
  cell. LF and CRLF line endings are both read, and a cell may be quoted with
  double quotes.
- An empty cell is a missing value. Each other cell is matched whole: a
  number is written in ASCII digits, with an optional sign and the dataset's
  decimal separator, and no thousands separator, exponent, surrounding blank
  or line break; a boolean is `true` or `false`, in lower case; and a quoted
  line break makes a code, date or boolean invalid too.

It checks the layout and the values of a file, never what they mean: test a
key that must be unique, a reconciliation against control totals or a rule of
the next section in the import itself. Each finding carries one of these
codes, which keep their meaning:

<!-- BEGIN generated: data-contract-codes -->
| Code | Finding |
| --- | --- |
| `DC-ENCODING` | The file does not decode in the dataset's encoding |
| `DC-SYNTAX` | The file is not readable as delimited text, such as a quote left open |
| `DC-HEADER` | The file has no header row |
| `DC-DUPLICATE-COLUMN` | The header names a column twice |
| `DC-UNKNOWN-COLUMN` | The header names a column that is not a field of the dataset |
| `DC-MISSING-COLUMN` | A required field has no column |
| `DC-CELL-COUNT` | A row, a blank line included, has more or fewer cells than the header |
| `DC-REQUIRED` | A required field is empty |
| `DC-TYPE` | A value is not of its field's type |
| `DC-VALUE` | A value is not one of its field's allowed values |
| `DC-RANGE` | A number is outside its field's range |
| `DC-CURRENCY` | An amount is given without a currency in its currency field |
<!-- END generated: data-contract-codes -->

The codes belong to the validator. The rules below are the workbook's, and the
workbook reports its own rule IDs.

## Validation and invalid input

| Rule | Rejects | Effect |
| --- | --- | --- |
| *Illustrative:* `DV-01` | A required field that is empty | The import stops and lists every row that fails |

- Each rule has an identifier that the workbook reports with its finding.
- Invalid input is reported, never silently corrected, rounded or defaulted.
- State whether a rule stops the whole import or excludes the row, and where
  excluded rows are listed.
- After a failed import, the workbook shows no results from earlier data as if
  they were current.

<a id="function-arguments"></a>

## Function arguments

The date layer reads no dataset. Its inputs are the arguments of the 22
`KPR_Dates_*` worksheet functions, whose rules are the
[calculation contracts](CALCULATION_CONTRACTS.md#date-layer-contracts)
`CALC-02` to `CALC-10`. This section defines how each argument is read, the
domain it must fall in, and the Excel error each rejection returns. A rule
that cites `Decision (#4)` is a design decision, recorded under
[Decisions](#decisions).

### Terms

- A **value argument** is a required argument; its values may decide the
  shape of the result. Every function except `KPR_Dates_HostDateSystem` has at
  least one.
- A **control argument** is an optional argument whose name starts with
  `Opt_`. It never decides or expands the shape of the result.
- A **scalar** is a value that is not an array, or a one-cell Range or
  one-element array; a **non-scalar** is a Range of more than one cell or an
  array of more than one element.
- A **call-level** error is one Excel error returned for the whole call, with
  no array. An **element-level** error is an Excel error at one position of the
  result, while the other positions are still evaluated.

Every public function is declared `As Variant`, so that it can return either
its result or an Excel error, and takes every argument `ByVal` as a `Variant`.
Required arguments come first and controls last.

<a id="date-values"></a>

### Date values

A date argument (`DateIn`, `StartDate`, `EndDate`) is read the same way
whatever the Windows or Excel locale: no parse uses `CDate`, `IsDate` or
another conversion that depends on regional settings (Decision (#4)).

| Value at one position | Result | Condition |
| --- | --- | --- |
| An Excel error | The same error, unchanged | `INPUT_ERROR_PROPAGATED` |
| A VBA `Date` | Accepted; its time is removed, then the window applies | — |
| A number | Read as an Excel serial under the 1900 date system ([SRC-03](SOURCES.md#register)); its time fraction is removed, then the window applies | — |
| Text of exactly ten characters `YYYY-MM-DD` ([SRC-02](SOURCES.md#register), calendar date, extended format) naming a date that exists | Accepted, then the window applies | — |
| Text shaped like a locale date, such as `31/12/2026`, `12/31/2026` or `10.02.2024` | `#VALUE!` | `DATE_TEXT_LOCALE` |
| Text that looks like a number, such as `61`, `45292` or `45292.5` | `#VALUE!`; never read as a serial | `DATE_TEXT_NUMERIC` |
| Any other text that is not exact `YYYY-MM-DD`, empty text included | `#VALUE!` | `DATE_TEXT_FORMAT` |
| Exact `YYYY-MM-DD` text, year 1900 to 9999, naming a date that does not exist, such as `2025-02-29` or `1900-02-29` | `#VALUE!` | `DATE_TEXT_IMPOSSIBLE` |
| A blank cell or `Empty` | `#VALUE!` | `INPUT_BLANK_REQUIRED` |
| `Null`, a Boolean, or another type | `#VALUE!` | `DATE_TYPE_REJECTED` |
| An accepted date outside 1900-03-01 to 9999-12-31, serial 60 and below included | `#NUM!` | `DATE_WINDOW` |

- **Time:** a number with a fraction, such as a cell holding `NOW()`, is
  accepted: the fraction is removed first, and the window applies to the
  date that is left. Integer arguments never remove a fraction
  ([integer values](#integer-values)).
- **Exact ISO text:** ASCII digits and hyphens only, every zero written. No
  other separator, no time suffix, no blank before or after, no other digits,
  and no rollover: a day beyond its month's length is never carried into the
  next month. Padded or time-suffixed text is `DATE_TEXT_FORMAT`.
- **Locale-shaped text:** exactly three runs of ASCII digits, of one or two,
  one or two, and two to four digits, separated by `/` or `.`. Any other text
  with those separators, such as the year-first `2024/02/10`, is
  `DATE_TEXT_FORMAT`.
- **Number-shaped text:** an optional sign, ASCII digits and at most one
  decimal point, with at least one digit, such as `61`, `45292.5`, `1.` or
  `.5`.
- **Order of the checks:** a date's existence is checked before the window.
  Exact text whose year is 1900 to 9999 but which names no date is
  `DATE_TEXT_IMPOSSIBLE`, even when the date would also fall outside the
  window: `1900-02-29` names the day that Excel's 1900 date system counts but
  that never existed ([SRC-04](SOURCES.md#register)), and calling it merely
  out of range would hide that. A year outside 1900 to 9999 is `DATE_WINDOW`
  with no other check.

<a id="integer-values"></a>

### Integer values

`nDays`, `nWeeks`, `nMonths`, `nYears`, `YearIn`, `MonthIn`, `WdIndex` and `n`
are read by one strict parser at each position (Decision (#4)):

| Value | Result | Condition |
| --- | --- | --- |
| A number inside the VBA `Long` range ([SRC-12](SOURCES.md#register)) that is whole, or within `1E-9` of a whole number | Accepted as that whole number, a `Long` | — |
| A number inside the `Long` range more than `1E-9` from any whole number | `#VALUE!`; never truncated or rounded | `INTEGER_FRACTION` |
| A number outside the `Long` range, with or without a fraction | `#NUM!` | `INTEGER_RANGE` |
| A Boolean, a `Date`, text (number-shaped text included), `Null` or an object | `#VALUE!` | `INTEGER_TYPE_REJECTED` |
| A blank cell or `Empty` | `#VALUE!` | `INPUT_BLANK_REQUIRED` |
| An Excel error | The same error, unchanged | `INPUT_ERROR_PROPAGATED` |

- **Range first:** the `Long` range is checked before the fraction, so a
  value outside it is `INTEGER_RANGE` even when it also has a fraction.
- **Tolerance:** `1E-9` is an absolute distance. It absorbs the residue of
  binary arithmetic, so that `(0.1+0.2)*10`, which is `2.9999999999999996`,
  reads as `3`, and is far smaller than any fraction a user means:
  `3.0000000001` is `3`, while `3.00000001` and `2.5` are `INTEGER_FRACTION`.
  A VBA `Decimal` is measured at its own precision, not after conversion to
  `Double`, so `3.000000000999999999` is within the tolerance.
- **Domains**, checked after the parse:

| Argument | Domain | Outside it | Condition |
| --- | --- | --- | --- |
| `YearIn` | 1900 to 9999 | `#VALUE!` | `DOMAIN_YEAR` |
| `MonthIn` | 1 to 12 | `#VALUE!` | `DOMAIN_MONTH` |
| `WdIndex` | 1 to 7, numbered by the call's weekday base (`CALC-02`) | `#VALUE!` | `DOMAIN_WEEKDAY` |
| `n` | 1 to 5 | `#VALUE!` | `DOMAIN_OCCURRENCE` |
| `nDays`, `nWeeks`, `nMonths`, `nYears` | Any `Long`: negative, zero or positive | A shift whose result leaves the window is `#NUM!` | `RESULT_WINDOW` |

A year is a calendar year, never a date: every whole number from 1900 to 9999
is also a valid serial, so an argument that took both would have to guess. A
caller that holds a date passes `YEAR(A1)`. A date, a serial typed as a `Date`,
or ISO text at a year position is `INTEGER_TYPE_REJECTED`.

<a id="control-values"></a>

### Control values

A control may be omitted, a scalar, or a one-cell Range or array. Omitting it,
passing `Empty` (a blank cell included), and forwarding an omitted VBA
`Optional` argument (`Missing`) all select its default; this is the one place
where `Empty` is valid. A control never vectorizes (Decision (#4)).

| Control | Accepted values | Default |
| --- | --- | --- |
| `Opt_WeekBaseMonday` | A VBA Boolean | `TRUE` |
| `Opt_KeepEOM` | A VBA Boolean | `FALSE` |
| `Opt_Rounding` | Text `NEAREST`, `FLOOR` or `CEILING`, compared without letter case after removing blanks before and after | `NEAREST` |

| Control value | Result | Condition |
| --- | --- | --- |
| More than one cell or element | Call-level `#VALUE!` | `CONTROL_NOT_SCALAR` |
| An unsupported shape: an object other than a Range, a multi-area Range, or an empty or unsupported array | Call-level `#VALUE!`, before its type is read | `SHAPE_UNSUPPORTED` |
| Of another type: a number such as `0` or `1`, Boolean-looking text such as `"TRUE"`, `Null` or a date for a Boolean control; anything but text for `Opt_Rounding` | Call-level `#VALUE!` | `CONTROL_TYPE_REJECTED` |
| `Opt_Rounding` text that is not one of the three tokens | Call-level `#VALUE!` | `CONTROL_TOKEN_UNKNOWN` |
| An Excel error | The same error, for the whole call | `CONTROL_ERROR_PROPAGATED` |

<a id="pillar-tokens"></a>

### Pillar tokens

A `Pillar` argument of `KPR_Dates_DateFromPillar` is text in this grammar,
compared without letter case (Decision (#4)):

```text
(ON | O/N | TN | T/N) | [+|-] [0-9]+[YMWD]   1 to 4 components, each unit at most once
```

- Blanks before and after are removed; a blank inside the token is malformed,
  so `1 M` and `1M 2W` are rejected.
- A sign applies to the whole token.
- `ON` and `O/N` mean `1D`, and `TN` and `T/N` mean `2D`: calendar days, in
  every `KPR_Dates_*` function. An alias stands alone, with no other component
  and no sign: `-ON` and `+T/N` are rejected. A backward shift of one day is
  `-1D`.
- Units may come in any order: `2W3D` and `3D2W` are the same interval. A unit
  may appear once: `1M2M` is rejected, never summed into `3M`.
- A quantity is whole and written; a decimal or missing quantity, an unknown
  unit or a partial parse is rejected.

How a token moves a date is `CALC-08`, and the subset that
`KPR_Dates_PillarFromDates` writes is `CALC-09`.

| Token | Result | Condition |
| --- | --- | --- |
| Not text | `#VALUE!` | `PILLAR_TYPE_REJECTED` |
| Outside the grammar: a blank inside, a missing or decimal quantity, an unknown unit, a sign with no body, empty text, a partial parse | `#VALUE!` | `PILLAR_TOKEN_MALFORMED` |
| A unit repeated, such as `1M2M` or `3D4D` | `#VALUE!` | `PILLAR_DUPLICATE_UNIT` |
| An alias with a sign, such as `-ON` or `+T/N` | `#VALUE!` | `PILLAR_ALIAS_SIGNED` |
| A valid token whose total or resulting date leaves the numeric or date domain | `#NUM!` | `PILLAR_AGGREGATE_RANGE` |

<a id="error-values"></a>

### Error values

A function returns an Excel error value ([SRC-13](SOURCES.md#register)),
never a message, a dialog or a plausible fallback value (Decision (#4)):

| Error | Means |
| --- | --- |
| `#VALUE!` | The input cannot be read or breaks the contract: text that is not exact ISO, a date that does not exist, a rejected type, a blank required value, a fraction where a whole number is required, a value outside an argument's domain, an invalid control or token, an unsupported shape, or a shape mismatch |
| `#NUM!` | A well-formed input whose value or result falls outside the supported domain: a date outside 1900-03-01 to 9999-12-31, a whole number outside `Long`, a shift that cannot give a supported date, a fifth weekday that the month does not have, a pillar total out of range, or a result of more than 100,000 elements |
| `#N/A` from the library | The result is not available in the calling worksheet's configuration: a workbook on the 1904 date system, or one whose date system cannot be read |
| Any error passed in | Returned unchanged, under the call-level rules of the [calculation contracts](CALCULATION_CONTRACTS.md#arrays-and-evaluation) |

An `#N/A` the library returns and an `#N/A` passed in are the same Excel value,
and nothing in the value tells them apart. The condition that produced an
error is carried by its identifier below, which test evidence records; in a
worksheet, `KPR_Dates_HostDateSystem()` returning `1904` shows that the
calling workbook refuses the date layer (`CALC-10`).

<a id="conditions"></a>

### Conditions

Each condition that returns an error has a stable identifier, so that test
cases and evidence can name what the Excel value cannot (Decision (#4)):

- The identifier, the Excel error and the level are three separate facts;
  none is derived from another.
- Identifiers are words, never numbers. One is never renamed or reused; a
  retired identifier stays retired.
- A new condition needs a new identifier, added to this table in the change
  that introduces it.
- Success has no identifier.

| Condition | When | Error | Level |
| --- | --- | --- | --- |
| `DATE_TEXT_FORMAT` | Text at a date position that is not exact `YYYY-MM-DD` and neither locale- nor number-shaped | `#VALUE!` | Element |
| `DATE_TEXT_LOCALE` | Locale-shaped date text | `#VALUE!` | Element |
| `DATE_TEXT_NUMERIC` | Number-shaped text at a date position | `#VALUE!` | Element |
| `DATE_TEXT_IMPOSSIBLE` | Exact ISO text naming a date that does not exist | `#VALUE!` | Element |
| `DATE_TYPE_REJECTED` | `Null`, a Boolean or another rejected type at a date position | `#VALUE!` | Element |
| `DATE_WINDOW` | An accepted date or serial outside 1900-03-01 to 9999-12-31, serial 60 and below included | `#NUM!` | Element |
| `INPUT_BLANK_REQUIRED` | A blank cell or `Empty` at a value position | `#VALUE!` | Element |
| `INPUT_ERROR_PROPAGATED` | An Excel error at a value position, returned unchanged | The error passed in | Element |
| `INTEGER_FRACTION` | A number inside `Long` more than `1E-9` from a whole number, at an integer position | `#VALUE!` | Element |
| `INTEGER_TYPE_REJECTED` | A Boolean, `Date`, text, `Null` or object at an integer position | `#VALUE!` | Element |
| `INTEGER_RANGE` | A number outside `Long`, with or without a fraction | `#NUM!` | Element |
| `DOMAIN_YEAR` | `YearIn` outside 1900 to 9999 | `#VALUE!` | Element |
| `DOMAIN_MONTH` | `MonthIn` outside 1 to 12 | `#VALUE!` | Element |
| `DOMAIN_WEEKDAY` | `WdIndex` outside 1 to 7 | `#VALUE!` | Element |
| `DOMAIN_OCCURRENCE` | `n` outside 1 to 5 | `#VALUE!` | Element |
| `OCCURRENCE_ABSENT` | A valid request for an occurrence that the month does not have | `#NUM!` | Element |
| `RESULT_WINDOW` | A valid operation whose result or intermediate value leaves the window, arithmetic overflow included | `#NUM!` | Element |
| `PILLAR_TYPE_REJECTED` | A `Pillar` that is not text | `#VALUE!` | Element |
| `PILLAR_TOKEN_MALFORMED` | A `Pillar` outside the grammar | `#VALUE!` | Element |
| `PILLAR_DUPLICATE_UNIT` | A unit repeated in one token | `#VALUE!` | Element |
| `PILLAR_ALIAS_SIGNED` | An alias with a sign | `#VALUE!` | Element |
| `PILLAR_AGGREGATE_RANGE` | A valid token whose total or resulting date leaves the domain | `#NUM!` | Element |
| `CONTROL_TYPE_REJECTED` | A control of a rejected type | `#VALUE!` | Call |
| `CONTROL_TOKEN_UNKNOWN` | `Opt_Rounding` text that is not `NEAREST`, `FLOOR` or `CEILING` | `#VALUE!` | Call |
| `CONTROL_NOT_SCALAR` | A control of more than one cell or element | `#VALUE!` | Call |
| `CONTROL_ERROR_PROPAGATED` | An Excel error passed as a control, returned unchanged | The error passed in | Call |
| `SHAPE_UNSUPPORTED` | A multi-area Range, a jagged, empty or higher-dimensional array, or an object that is not a Range | `#VALUE!` | Call |
| `SHAPE_MISMATCH` | Non-scalar value arguments of different dimensions | `#VALUE!` | Call |
| `CAPACITY_EXCEEDED` | A result of 100,001 elements or more | `#NUM!` | Call |
| `HOST_DATE1904` | A worksheet caller in a workbook on the 1904 date system | `#N/A` | Call |
| `HOST_UNRESOLVED` | A worksheet caller whose workbook's date system cannot be read | `#N/A` | Call |

<a id="internal-failures"></a>

### Internal failures

No condition covers an unexpected failure. A handler that catches one only
contains it: reaching it in a call that follows this contract is a defect,
which fails the run that meets it, and is never recorded as an expected
`#VALUE!`, `#NUM!` or `#N/A` (Decision (#4)).

The core reports each condition above by returning its code, never by raising
an error (Decision (#4)). It raises one error, in the project's range
([VBA house style](../VBA_HOUSE_STYLE.md)):

| Constant | Number | Source | Message |
| --- | --- | --- | --- |
| `ERR_UNMAPPED_CONDITION` | `vbObjectError + 2048` | `CORE_Err.ErrForCondition` | `Condition has no worksheet error mapping; this is an internal defect.` |

It is raised when a condition code with no Excel error reaches the mapping: the
success code, a propagation condition, whose error is the one passed in, or an
unknown code. The numbers of the condition codes in VBA are internal: test
cases and evidence name the identifiers.

## Point-in-time data

Use only data that was known on the reporting date. State, for each dataset,
which date decides that: the reporting date, a market data cut-off, or the
date a record was booked.

## Volume

Excel worksheets hold at most 1,048,576 rows and 16,384 columns. State the
expected and the largest supported volume, and how the workbook imports data
beyond a sheet's limits, if it needs to.

## Lineage

Each import records, where the workbook can read it later:

- the file name and its SHA-256;
- the number of rows read, accepted and rejected;
- the reporting date and the import time;
- the commit of the workbook's source, when the build records it.

## Synthetic data

The repository holds synthetic data only: no real positions, counterparties,
clients or amounts. Test fixtures say they are synthetic and how they were
made.

## Decisions

| Date | Decision | Reason | Issue |
| --- | --- | --- | --- |
| *Illustrative* | How a provider file with another date format is converted to ISO 8601 before import | Why | Link |
| 2026-11-25 | The date layer's inputs are worksheet-function arguments, defined under [Function arguments](#function-arguments); `data-contract.json` keeps the illustrative `DS-01` until [#11](https://github.com/danielep71/K-PRICING/issues/11) replaces the illustrative cases and decides the datasets that the date-layer cases use | The machine-readable contract describes imported files and sheets; its types have no Excel error, no Variant and no array, so the arguments would not fit it | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-25 | Dates are read without regard to locale: a number as a 1900-system serial, a `Date` as itself, and text only as exact `YYYY-MM-DD`; no parse uses `CDate` or `IsDate` | The same cell must give the same date on every host; `CDate` reads `01/02/2026` by the regional settings | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-25 | The supported window is 1900-03-01 to 9999-12-31 | Below serial 61, Excel's 1900 system counts a 1900-02-29 that never existed, so Excel and VBA disagree by a day; 9999-12-31 is the latest date both Excel and VBA hold | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-25 | A whole-number argument accepts a value within `1E-9` of a whole number and rejects any other fraction; it never truncates or rounds | A computed count such as `(0.1+0.2)*10` must read as 3, and a real fraction such as 2.5 must not be silently changed | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-25 | Every error has a semantic condition identifier, separate from its Excel error and level | Excel shows only the error value; tests and evidence need the cause, such as a library `#N/A` against one passed in | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-25 | The core returns a condition code for each contract condition and raises only `ERR_UNMAPPED_CONDITION`, a defect; the facade maps codes to Excel errors | A condition is an expected outcome of one element among up to 100,000, not a failure: raising and trapping it per element costs time and hides the cause; the house style's documented errors remain for defects | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-25 | Pillar tokens accept the aliases `ON`, `O/N`, `TN`, `T/N` as 1 and 2 calendar days, one occurrence of each unit, and no sign on an alias | Callers write these forms; business-day meanings belong to the calendar milestones, so a result never changes silently when they arrive; a repeated unit is more likely a typing error than a sum | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-26 | The date layer's test cases are [worksheet-function cases](TEST_CASES.md#worksheet-function-cases), which hold the arguments of one call and use no dataset; `DS-01` stays with the illustrative cases until those are retired | The arguments of a call are worksheet values, which the datasets' typed fields cannot hold; one dataset per function would restate the signatures | [#11](https://github.com/danielep71/K-PRICING/issues/11) |
