# 📊 Workbook contract

This contract defines what the workbook looks like:

- its sheets and their code names;
- how a user moves between the sheets;
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
  the checkout or under `test-results/`. Only the macro-free
  `src/workbook/Template.xlsx` is tracked, and it is generated from source
  too ([build](#build)).
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
  or a hard-coded address. `M_Ranges` reads only the Range objects that a
  worksheet function receives as arguments, and `M_Host` only
  `Application.Caller` and the `Date1904` of the caller's own workbook;
  `M_Register` only gives the 22 functions of its own workbook their
  Insert Function descriptions (`Application.MacroOptions`); none addresses a
  sheet. `M_Actions` writes only the cell the user made active, one formula
  at a time ([operations](#operations)), `M_CommandBars` only its own toolbar
  and `M_Ribbon` nothing. Core, classes and the facade never touch a
  workbook object: the facade passes a Range argument to `M_Ranges` unread,
  and its one call to the Excel object model is `Application.Volatile` in
  `KPR_Dates_HostDateSystem`.
- **No partial results.** An operation validates first, calculates in memory,
  then clears and writes all of its result areas. On failure it writes no
  new result: it clears its result areas and marks them `FAILED`
  ([status and staleness](#status-and-staleness)).

## Sheets

Tab order is the order of the table. A sheet is visible unless a decision
below says why it is hidden.

| # | Tab name | Code name | Role | Code writes | Content |
| ---: | --- | --- | --- | --- | --- |
| 1 | `Home` | `shHome` | Reference | Build only | The build stamp, what the workbook provides and does not cover, how to use it, and a link to every sheet |

The workbook's 22 `KPR_Dates_*` worksheet functions return their results to
the cells that call them and write no sheet, so it has no input, parameter,
output or diagnostics sheet yet. A sheet added later takes the next row of
the table, and the [navigation](#navigation) and the template follow it.

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

Every sheet shares this layout, except a `Test` sheet, which the build
appends blank to a development workbook:

- **Columns:** A and B are margins 4.625 wide, C holds labels and section
  titles (26), D text and values (185.375), E is a margin (4.625), and the
  columns after E are hidden. Every column is centred vertically, and
  gridlines are off.
- **Title band:** rows 1 to 3, frozen, and nothing else frozen. Row 1, 24
  points high, is the black band of `B1:D1`, with the workbook name,
  `K-PRICING`, in gold in `B1`. Row 2, 26.1 points, is the blue band
  (`#265A99`) of `B2:D2`, with the sheet's tab name in bold white in `B2`.
  Row 3, 30 points, holds a grey italic subtitle in `C3` that says what the
  sheet is for.
- **Sections:** from row 4, each section is a title row, 30 points high, in
  bold navy (`#1F3864`) in column C, then a table of one row per entry, 20.1
  points high: the label in bold in C, the text in D, wrapped. A thin line
  boxes the table, and dashed lines separate its rows and its two columns.
  A row 9.95 points high separates a table from the next section.
- **Print:** landscape, one page wide.
- Every sheet that code writes has a status line, named `<Sheet>_Status`.
- A table's first header cell carries the name `<Sheet>_<Table>`, and its data
  starts on the row below. Tables that grow downward sit side by side, so
  that none can overwrite another.
- Input cells look different from cells that code writes, and are the only
  unlocked cells ([protection](#protection)).
- A table that would exceed a worksheet's 1,048,576 rows fails the operation
  before anything is written; it is never truncated.

## Navigation

- **Home lists every sheet.** Its *Sheets* section has one entry per sheet of
  the [table](#sheets), in tab order: the tab name in column C and what the
  sheet holds in column D. The tab name of each sheet other than Home is a
  link to that sheet's cell `A1`.
- **Every other sheet links back.** Cell `D1`, in the black band, holds the
  link `Home` to Home's cell `A1`. A `Test` sheet is exempt, as in
  [Layout](#layout).
- Links are cell hyperlinks to a location in this workbook: they work while
  the sheet is protected, and need no shape, macro or external address.

## Defined names

| Name | Sheet | Refers to | Written by |
| --- | --- | --- | --- |
| `build_commit` | `shHome` | `D5`: the commit the workbook was built from | The build |
| `build_time` | `shHome` | `D6`: the commit's committer time | The build |
| `build_configuration` | `shHome` | `D7`: `production` or `development` | The build |

The three names sit in Home's *This workbook* section, each with its label
in column C: `Commit`, `Commit time` and `Configuration`.

- Every name is workbook-scoped and visible, and refers to one cell of this
  workbook by an absolute reference. External and relative references are
  not allowed.
- **Settings:** `cfg_<code>`, on the value cell of each setting.
- **Table anchors:** `<Sheet>_<Table>`, on a table's first header cell. No
  table is named `Status`, so an anchor never collides with a status line.
- **State:** the status line `<Sheet>_Status` of each sheet that code writes,
  and `Run_Status`, the state of the last run, on one of those sheets. Home
  is written by the build only, so the workbook has neither; declare them
  here with the first sheet that code writes.
- **Build stamp:** `build_commit`, `build_time` and `build_configuration`, on
  the reference sheet.
- The build creates every name of the table. Before an operation, the
  adapter checks that each name it uses exists and refers to the expected
  sheet; a missing or moved name stops the operation before any calculation.

## Areas

The workbook has no area yet. The first sheet that holds one lists it here,
in a table with the columns *Area*, *Sheet*, *Kind*, *Written by* and
*Holds*.

An area is named by its table anchor, which refers to its first header cell
only ([defined names](#defined-names)). Code that reads, writes or restores
an area derives it from the anchor: the rows below it, with the columns of
its header. The contract fixes each area's size in its table, or says that
its number of rows varies.

- In a settings area, users type only the value column. Its other columns,
  such as the code, the unit and the contract that uses the setting, are
  labels that the template fills and keeps locked, with the header row
  ([protection](#protection)).

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

The worksheet functions are evaluated by Excel and write no sheet. The
operations are the actions of the menus: the **K-PRICING** Ribbon tab and the
**KPR Dates** toolbar call the same entry points of `M_Actions`, with the
action text that each control stores.

| Entry point | Reads | Writes | Requires |
| --- | --- | --- | --- |
| `M_Actions.InsertFunction` | The active cell | One formula in the active cell, `=<function>()`, through the Function Arguments dialog; a cancelled dialog gives the cell back its content | A cell selected, not locked on a protected sheet |
| `M_Actions.RegisterHelp` | Nothing | The category, description and argument descriptions that Excel keeps for the 22 functions until it closes | Nothing |
| `M_Actions.ClearHelp` | Nothing | The same options, returned to the User Defined category with no description | Nothing |
| `M_Actions.RunTests` | The project suites | Scratch workbooks that each suite closes, and each suite's log in the Immediate window | A development build |
| `M_CommandBars.InstallToolbar`, `RemoveToolbar` | Nothing | The **KPR Dates** toolbar, temporary: Excel discards it when it closes | Nothing |


- Each entry point is an Excel adapter procedure that calls the facade; it
  holds no regulatory formula.
- It captures and restores the Excel settings it changes, and reports both
  the operation and the cleanup outcome
  ([house style](../VBA_HOUSE_STYLE.md#operation-and-cleanup-outcomes)).
- A cancelled selection, such as a closed file dialog, changes nothing.

## Status and staleness

These rules apply from the first sheet that code writes; Home is written by
the build only. `Run_Status` and each `<Sheet>_Status` hold `CURRENT`,
`STALE` or `FAILED`, with the run ID and the time of the operation that set
it.

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
  [installation guide](../../INSTALLATION.md) lists them;
- the Ribbon tab, `src/workbook/customUI14.xml`, as the package's
  `customUI/customUI14.xml` part, in every configuration. Its Development
  group shows only in a development build.

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
- **From the template**, the path this workbook takes, since
  `src/workbook/Template.xlsx` exists. The build
  starts from a copy of it, so the formats, data validation, print areas and
  protection are the template's. The template holds the sheets of the table
  other than `Test`, in order, and defines their names on the same cells; the
  build stops otherwise. The build sets the code names, writes the stamp
  values into the stamp cells, keeping their style, makes the workbook
  macro-enabled and adds the VBA project. A development build appends each
  `Test` sheet blank, with its names, so `Test` sheets come last in the
  table. Every other part of the template is copied unchanged, and two builds
  of the same commit and working tree stay byte-identical.

`tools/gen_workbook_template.py` generates the template from its source: the
sheets and names of `tools/devbuild/workbook.json`, laid out as
[Layout](#layout) and [Navigation](#navigation) say, with the text of Home
held in the script. It writes no VBA project, no document properties and
no external link, and the same source always gives the same bytes. Never
edit the template in Excel: change the script or `workbook.json`, then run

```sh
python tools/gen_workbook_template.py --write
```

**Repository integrity** runs it with `--check`, which fails when the
template differs from what the script generates. It also fails on macros
in the template, or on an author, last editor, company or manager in it. The
[installation guide](../../INSTALLATION.md#building-a-workbook-from-source)
says how to run the build.

## Protection

The template carries the protection of each sheet and the lock flag of each
cell. The protection matrix, `docs/methodology/workbook-protection.json`,
records them. For Home it reads:

```json
{
  "schema_version": 1,
  "sheets": {
    "shHome": {
      "protection": {},
      "output": ["A1:E20"]
    }
  }
}
```

No sheet is protected unless a decision below says so, so Home is not
protected. Its cells keep Excel's default lock flag, which takes effect only
if a later decision protects the sheet; the matrix still lists them as
output, so that such a decision cannot leave a label or the stamp unlocked.

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
| 2026-11-27 | The workbook has one sheet, Home, of role Reference; no sheet is hidden | The 22 worksheet functions write no sheet, so there is nothing to input, store or show beyond what the workbook is and how it was built | [#16](https://github.com/danielep71/K-PRICING/issues/16) |
| 2026-11-27 | `tools/gen_workbook_template.py` generates `src/workbook/Template.xlsx`; it is never saved from Excel | The same source always gives the same bytes, a change is reviewed as a diff of the script, and no personal document metadata can enter it | [#16](https://github.com/danielep71/K-PRICING/issues/16) |
| 2026-11-27 | Navigation is by cell hyperlinks: Home lists every sheet, and every other sheet links back to Home from `D1` | Cell links work on a protected sheet and need no shape, macro or drawing part, and the template's source states every link | [#16](https://github.com/danielep71/K-PRICING/issues/16) |
| 2026-11-27 | The menus are a Ribbon tab and a toolbar that run the same actions of `M_Actions`; the toolbar is installed from the macro list, never on open | The tab needs Office 2010 or later and a policy that allows custom tabs; the toolbar covers the rest, and installing it on open would show the same buttons twice where the tab works | [#17](https://github.com/danielep71/K-PRICING/issues/17) |
| 2026-11-27 | No sheet is protected unless a decision here says so | Protection is a choice I make per sheet, with its reason; by default users may type anywhere, including the functions on Home | [#16](https://github.com/danielep71/K-PRICING/issues/16) |
