# 🧾 Data contract

This contract defines every input the workbook accepts: where it comes from,
its layout and units, what makes it invalid, and how an import is traced.
The [calculation contracts](CALCULATION_CONTRACTS.md) refer to its datasets
and fields by name.

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
