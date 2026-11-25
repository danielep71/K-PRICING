# 📊 Workbook contract

This contract defines what the workbook looks like:

- its sheets and their code names;
- the defined names the code addresses;
- which areas are inputs and which the code writes;
- how each operation reads and writes them;
- how a failed run marks earlier results as no longer current;
- what the build creates, and how each sheet is protected.

The [data contract](DATA_CONTRACT.md) defines the inputs and the
[calculation contracts](CALCULATION_CONTRACTS.md) the results; this contract
says where they sit in the workbook. The
[repository structure](../REPOSITORY_STRUCTURE.md) says which components may
touch a sheet, and the [house style](../VBA_HOUSE_STYLE.md) how they are
named.

## Principles

These defaults hold unless a decision below replaces one:

- **Built from source, never committed.** The build creates the workbook from
  the exported source of one commit, as the
  [installation guide](../../INSTALLATION.md) describes, and saves it outside
  the checkout or under `test-results/`. Only a macro-free
  `src/workbook/Template.xlsx` may be tracked.
- **Datasets where the data contract says.** The *File or sheet* column of the
  data contract's *Datasets* table says whether a dataset is read from a file
  into memory or typed on a sheet. A dataset read from a file is never copied
  onto a sheet.
- **No calculation formulas on sheets.** VBA writes every calculated value,
  through the public facade, as a value. Input areas may use data validation;
  conditional formats and charts may read written values.
- **Adapters own the sheets.** Only the Excel adapters (`M_` modules) and the
  document modules in `src/workbook/` read or write a worksheet. They address
  a sheet by its code name and a range by its defined name, never by tab name
  or a hard-coded address. Core, classes and the facade never touch a
  workbook object.
- **No partial results.** An operation validates first, calculates in memory,
  then clears and writes all of its result areas. On failure it writes no
  new result: it clears its result areas and marks them `FAILED`
  ([status and staleness](#status-and-staleness)).

## Sheets

Tab order is the order of the table. A sheet is visible unless a decision
below says why it is hidden.

| # | Tab name | Code name | Role | Code writes | Content |
| ---: | --- | --- | --- | --- | --- |
| *Illustrative:* 1 | `Readme` | `shReadme` | Reference | Build only | What the workbook is, its limits, and the build stamp |
| *Illustrative:* 2 | `Inputs` | `shInputs` | Input | No | Reporting date and run settings ([areas](#areas)) |

- **Code name:** `sh` plus the tab name in PascalCase, as the house style
  requires. The workbook's own document module is `ThisWorkbook`.
- **Role:**
  - `Reference`: what the workbook is, its limits and its build;
  - `Input`: typed by users;
  - `Parameter`: values a calibration writes and later runs read;
  - `Output`: results;
  - `Diagnostics`: findings, logs and test results;
  - `Test`: present only in a development build.
- **Code writes:**
  - `No`: users edit the sheet, and code never writes it;
  - `Yes`: code replaces its content, and users never edit it;
  - `Append only`: code adds rows, and never changes or deletes them;
  - `Build only`: only the build writes it.

## Layout

State the layout every sheet shares. A common default:

- The first rows carry the workbook's title band and the sheet title. Every
  sheet that code writes has a status line, named `<Sheet>_Status`.
- A table's first header cell carries the name `<Sheet>_<Table>`, and its data
  starts on the row below. Tables that grow downward sit side by side, so
  that none can overwrite another.
- Input cells look different from cells that code writes, and are the only
  unlocked cells ([protection](#protection)).
- A table that would exceed a worksheet's 1,048,576 rows fails the operation
  before anything is written; it is never truncated.

## Defined names

| Name | Sheet | Refers to | Written by |
| --- | --- | --- | --- |
| *Illustrative:* `build_commit` | `shReadme` | The commit the workbook was built from | The build |
| *Illustrative:* `build_time` | `shReadme` | The commit's committer time | The build |
| *Illustrative:* `build_configuration` | `shReadme` | `production` or `development` | The build |
| *Illustrative:* `cfg_reporting_date` | `shInputs` | The value cell of the reporting date | Users |
| *Illustrative:* `Inputs_Settings` | `shInputs` | The first header cell of the settings area ([areas](#areas)) | Users, in the value column below it |

- Every name is workbook-scoped and visible, and refers to one cell of this
  workbook by an absolute reference. External and relative references are
  not allowed.
- **Settings:** `cfg_<code>`, on the value cell of each setting.
- **Table anchors:** `<Sheet>_<Table>`, on a table's first header cell. No
  table is named `Status`, so an anchor never collides with a status line.
- **State:** the status line `<Sheet>_Status` of each sheet that code writes,
  and `Run_Status`, the state of the last run, on one of those sheets. The
  illustrative sheets are written by users or the build only, so they have
  neither; declare them here with the first sheet that code writes.
- **Build stamp:** `build_commit`, `build_time` and `build_configuration`, on
  the reference sheet.
- The build creates every name of the table. Before an operation, the
  adapter checks that each name it uses exists and refers to the expected
  sheet; a missing or moved name stops the operation before any calculation.

## Areas

| Area | Sheet | Kind | Written by | Holds |
| --- | --- | --- | --- | --- |
| *Illustrative:* `Inputs_Settings` | `shInputs` | Input | Users, in the value column | Below a header row at `B4:E4`, one row per setting in columns B to E: code, value, unit and the contract that uses it. It has one row, the reporting date, so its value cell is `C5`, the cell of `cfg_reporting_date` |

An area is named by its table anchor, which refers to its first header cell
only ([defined names](#defined-names)). Code that reads, writes or restores
an area derives it from the anchor: the rows below it, with the columns of
its header. The contract fixes each area's size, as in the table above, or
says that its number of rows varies.

- In a settings area, users type only the value column. The code, unit and
  contract columns are labels that the template fills and keeps locked,
  with the header row ([protection](#protection)); the label of the
  reporting date is its code cell, `B5`.

- **Input** areas hold what users type: settings, parameters chosen by hand,
  and datasets the data contract places on a sheet. Each cell's type, unit
  and valid range come from the data contract or the calculation contract
  that reads it. A value that fails validation stops the operation; it is
  never corrected.
- **Parameter** areas hold values one operation writes and later operations
  read, with the operation and run that wrote them, such as the data an
  import reads from a file and its lineage.
- **Result** areas hold the outputs of the calculation contracts, in the units
  and scopes they define. Every row names the scope and the run it belongs
  to.
- **Diagnostics** areas hold findings, the operation log and, in a development
  build, test results.

## Operations

| Entry point | Reads | Writes | Requires |
| --- | --- | --- | --- |
| *Illustrative:* `M_Run.RunCalculation` | `Inputs_Settings` and the datasets of `CALC-01` | Its result areas, findings and log | A valid reporting date |

- Each entry point is an Excel adapter procedure that calls the facade; it
  holds no regulatory formula.
- It captures and restores the Excel settings it changes, and reports both
  the operation and the cleanup outcome
  ([house style](../VBA_HOUSE_STYLE.md#operation-and-cleanup-outcomes)).
- A cancelled selection, such as a closed file dialog, changes nothing.

## Status and staleness

`Run_Status` and each `<Sheet>_Status` hold `CURRENT`, `STALE` or `FAILED`,
with the run ID and the time of the operation that set it.

| Area kind | After success | After failure |
| --- | --- | --- |
| Input | Unchanged: no operation writes it | Unchanged |
| Parameter | An import replaces its data and their lineage; another operation appends one row or set | Unchanged: nothing replaced or appended |
| Result | Replaced and `CURRENT` | Cleared and `FAILED` |
| Diagnostics | Findings rewritten; one log row appended | Findings rewritten, so those that explain the failure are shown; one log row appended |

- A successful import marks every result that depends on it `STALE`. A
  rejected import keeps the previous data and still marks those results
  `STALE`, so no result from earlier data is shown as current, as the data
  contract requires.
- A result records a fingerprint of the inputs it used. Each operation
  compares it with the current inputs and marks the result `STALE` on any
  difference. Without worksheet event handlers, an edit made after a run is
  flagged at the next operation, not when it is typed; state which applies.

## Build

The build writes, from one commit:

- the sheets of [Sheets](#sheets), in order, with their tab and code names;
- every defined name;
- once a decision here defines them, the formats, the data validation of the
  input areas and the print areas;
- the build stamp on the reference sheet: the commit, its committer time
  (not the clock, so that a rebuild is byte-identical) and the configuration.
  A working tree with changes, or a build whose `--override` replaces
  modules, is stamped as not evidence-grade;
- the VBA components of that commit, as the
  [installation guide](../../INSTALLATION.md) lists them.

It writes no data and no results. A production build has no `Test` sheet, test
module or example; a development build adds them.

`tools/devbuild/build_workbook.py` builds the sheets, code names, defined
names, stamp and VBA project. It reads them from
`tools/devbuild/workbook.json`, whose sheets, roles and names must match the
tables above. It leaves the `Test` sheets and their names out of every
configuration but `development`. It builds the sheets in one of two ways:

- **Blank sheets**, while there is no `src/workbook/Template.xlsx`. Each
  sheet holds only the stamp, a title on the reference sheet and, to the
  left of each stamp and setting cell, its label; a table anchor has none. Formats, data validation and print areas are not
  built on this path.
- **From the template**, once `src/workbook/Template.xlsx` exists. The build
  starts from a copy of it, so the formats, data validation, print areas and
  protection are the template's. The template holds the sheets of the table
  other than `Test`, in order, and defines their names on the same cells; the
  build stops otherwise. The build sets the code names, writes the stamp
  values into the stamp cells, keeping their style, makes the workbook
  macro-enabled and adds the VBA project. A development build appends each
  `Test` sheet blank, with its names, so `Test` sheets come last in the
  table. Every other part of the template is copied unchanged, and two builds
  of the same commit and working tree stay byte-identical.

Save the template in Excel as a macro-free `.xlsx` workbook without personal
document metadata: **Repository integrity** fails on macros, or on an
author, last editor, company or manager in it. The
[installation guide](../../INSTALLATION.md#building-a-workbook-from-source)
says how to run the build.

## Protection

The template carries the protection of each sheet and the lock flag of each
cell. Record them in the protection matrix,
`docs/methodology/workbook-protection.json`, added with the template. For
the illustrative sheets above, it could read:

```json
{
  "schema_version": 1,
  "sheets": {
    "shReadme": {"protection": {}, "output": ["A1:B5"]},
    "shInputs": {
      "protection": {"sheet": true, "objects": true, "scenarios": true},
      "input": ["C5"],
      "output": ["A1:B5", "C4:E4", "D5:E5"]
    }
  }
}
```

- **Sheets:** one entry for each sheet of the template, by code name. `Test`
  sheets are not in the template and have no entry.
- **`protection`:** the options of the sheet's protection, named as the
  attributes of the SpreadsheetML `sheetProtection` element. `sheet: true`
  turns protection on; every other option set to `true` blocks its action, or
  protects its part, while the sheet is protected. An option left out takes
  the standard's default: `false` for `sheet`, `objects`, `scenarios`,
  `selectLockedCells` and `selectUnlockedCells`, `true` for the others. `{}`
  is an unprotected sheet. Never put a password in the matrix.
- **`input`:** the cells and ranges, such as `C5` or `C5:D10`, that users
  type in. They must be unlocked.
- **`output`:** the cells and ranges that users must not change, such as
  results and labels. They must be locked. An input range never overlaps an
  output range; a cell in neither is not checked.

`tools/check_workbook_protection.py`, which **Repository integrity** runs,
compares the matrix with the template's sheets and styles. A cell takes the
style of its own cell, or else of its row when the row is formatted, or else
of its column, as Excel shows it; a style is locked unless it unlocks the
cell. A lock flag that differs from the matrix fails the check, and so does
a protection option, compared when the matrix or the template protects the
sheet; the options of a sheet that neither protects are not compared. Without the template the check reads the matrix alone, and
without either it passes as inactive; a template without a matrix fails. It
compares neither the password nor protection that VBA sets at run time with
`Worksheet.Protect`.

## Decisions

| Date | Decision | Reason | Issue |
| --- | --- | --- | --- |
| *Illustrative* | Whether a sheet is hidden | Why | Link |
