# 📜 Changelog

Notable changes to this repository, from a user's point of view. The format
follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions
follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html). It is
neither a commit log nor release evidence. **No version has been released.**

## 🧭 Maintenance policy

- Add a material change under `[Unreleased]` in the pull request that makes
  it. A change nobody would notice, such as a typo or a refactor with no
  effect, needs no entry.
- Describe the observable result: what exists, what a user or contributor
  can now do, what changed in a contract, and what to do about it.
- Link the issue or pull request that holds the detail.
- Record only evidence that was produced. A static check is not an Excel run:
  say what was and was not run in Excel.
- Once a version has been released, edit an existing entry of the next
  version rather than adding one that describes a change to its unreleased
  work: an entry describes what the release contains. The first release's
  section also records, as **Changed** and **Fixed**, how its additions
  changed while it was prepared.
- Entries move into a dated section only when a release is prepared, as
  [RELEASING.md](RELEASING.md) describes. After that, until the tag exists,
  a fix for a release gate that failed on the candidate records its change in
  the dated section being released
  ([step 2](RELEASING.md#2-integrate-and-freeze-the-candidate)). A published section is not
  rewritten; a demonstrable error is corrected in review, with a line
  `Corrected after release on YYYY-MM-DD: <reason>` in that section.

<a id="date-and-version-rules"></a>

### Date and version rules

`python tools/check_changelog.py` checks the headings, dates, order, link
references and `VERSION`, and that this introduction, the *Current status* of
[RELEASING.md](RELEASING.md#current-status) and the README describe the same
release state; **Repository integrity** also checks that every link reference
points to this repository. Once a release's tag exists, its section must equal
that section at the tag, unless it holds the dated correction line above; a
correction is reported as a note. What a version number means is checked
in review.

| Rule | Standard |
| --- | --- |
| Version | `MAJOR.MINOR.PATCH`, without a leading `v` in headings |
| Release heading | `## [X.Y.Z] - YYYY-MM-DD`, with a real calendar date |
| Order | `[Unreleased]` first, then releases from newest to oldest, by version and date |
| Link references | One for `[Unreleased]` and one for every release; a release's points to `.../releases/tag/vX.Y.Z` |
| `0.y.z` | Initial development: no compatibility is promised between versions |
| Patch, minor, major | From `1.0.0`: a compatible correction, a compatible capability, an incompatible change to a calculation result, an input or output layout, or the supported VBA API |

| Category | Use for |
| --- | --- |
| **Added** | New capabilities, files, checks or tests |
| **Changed** | Changes to existing behaviour, contracts, tooling or documentation |
| **Fixed** | Corrected defects |
| **Security** | Safely disclosed security corrections |
| **Validation** | Evidence actually produced |
| **Known limitations** | Deliberate, unresolved boundaries |

Use only the categories an entry needs.

## [Unreleased]

### Added

- The date-layer methodology, ahead of its code: how the arguments of the 22
  `KPR_Dates_*` worksheet functions are read and which Excel error each
  rejection returns, with a stable condition identifier
  ([data contract](docs/methodology/DATA_CONTRACT.md#function-arguments)); the
  calculation contracts `CALC-02` to `CALC-10`, with the host, array and
  tolerance rules they share
  ([calculation contracts](docs/methodology/CALCULATION_CONTRACTS.md#date-layer-contracts));
  their sources; and the cases, relation tests and Excel host checks that
  will validate them
  ([validation plan](docs/methodology/VALIDATION_PLAN.md#date-layer-coverage)).
  No function exists yet and nothing was run in Excel
  ([#4](https://github.com/danielep71/K-PRICING/issues/4)).
- The date layer's error and parsing core, `CORE_Err` and `CORE_Parse` in
  `src/core/`: the native `#VALUE!`, `#NUM!` and `#N/A` values, the 31
  conditions of the data contract with their identifiers and errors, the one
  raised error `ERR_UNMAPPED_CONDITION`, the supported date window, and the
  strict parsers of dates, whole numbers, Boolean controls and `Opt_Rounding`.
  They use no Excel object model and are not supported API. A VBA suite,
  `TEST_CoreErrParse` (467 assertions), and a Python test that keeps the code
  equal to the data contract check them. Only the static checks and, as an early
  warning that is not Excel evidence, a LibreOffice run were done: 457 of the
  467 assertions passed, and the ten that failed need a `Decimal`, which
  LibreOffice lacks. Nothing was run in Excel
  ([#5](https://github.com/danielep71/K-PRICING/issues/5)).
- The date layer's calendar core, `CORE_Dates` in `src/core/`: the window
  predicate and argument domains, the leap and month-length rules, the weekday,
  the period boundaries and tests, shifts by days and months, the weekday
  locators, and the parse, shift and format of pillar tokens, one routine for
  each rule of `CALC-02` to `CALC-09`. Its tests, `TEST_CoreDates` (2,083
  assertions), take their expected values from `date_core_reference.py`. Only
  the static checks and, as an early warning that is not Excel evidence, a
  LibreOffice run were done: all 2,083 assertions passed, and so did a
  comparison of the pillar format with the reference on 9,000 random intervals.
  Nothing was run in Excel
  ([#6](https://github.com/danielep71/K-PRICING/issues/6)).
- The date layer's shape engine, split in two: `CORE_Array` in `src/core/`,
  which applies the array rules (shape, capacity, broadcasting, traversal) to
  plain values with no Excel object, and the adapter `M_Ranges` in
  `src/workbook/`, which describes a Range from its counts and reads it only
  after its size passes the cap. An element that is an object or an array
  is now refused on every path, where the engine it replaces read an object
  element through its default member and returned a nested array as a scalar;
  a worksheet cannot pass either. Their tests are `TEST_CoreArray` (122
  assertions) and `TEST_Ranges` (62). Only the static checks and, as an early
  warning that is not Excel evidence, a LibreOffice run were done: all 62 range
  assertions and 121 of the 122 core ones passed, the one left being an
  unallocated array that LibreOffice ranks 1, where VBA ranks it 0; and,
  before those two corrections, the split engine gave the same results as the
  engine it replaces on 24 argument shapes. Nothing was run in Excel
  ([#7](https://github.com/danielep71/K-PRICING/issues/7)).
- The 22 `KPR_Dates_*` worksheet functions, in the facade `KPR_DATES_DAYS` in
  `src/modules/`, with the names, signatures, argument order and defaults
  that `docs/PUBLIC_API.txt` lists as the supported API. Each function checks
  the calling workbook (the adapter `M_Host` refuses the 1904 date system),
  reads its arguments in the contract's order and returns its result or an
  Excel error cell for cell; only `KPR_Dates_HostDateSystem` is volatile. Its
  tests are `TEST_Dates` (81 assertions). Only the static checks and, as an
  early warning that is not Excel evidence, a LibreOffice run were done: 72 of
  the 81 assertions passed, the other 9 meeting LibreOffice's handling of
  optional arguments, and the functions gave the same results as the ones they
  replace on 29,440 combinations of arguments. Nothing was run in Excel
  ([#8](https://github.com/danielep71/K-PRICING/issues/8)).
- The registration of the 22 worksheet functions, the adapter `M_Register` in
  `src/workbook/`: `RegisterFunctions` gives each `KPR_Dates_*` function the
  category `KPR Dates`, its description and one description per argument for
  the Insert Function dialog, and `UnregisterFunctions` returns them to the User
  Defined category with no description; both can run again without adding an
  entry. The texts are those of the facade's documentation, and a Python test
  keeps them equal to it and to the signatures of the supported API. Only the
  static checks were run. Nothing was run in Excel, so what the Insert Function
  dialog shows is not yet verified
  ([#10](https://github.com/danielep71/K-PRICING/issues/10)).
- The date layer's 1,361 test cases, converted from the project's earlier
  fixture set into a third case form, the worksheet-function case: one call of
  a `KPR_Dates_*` function, its caller (VBA, a 1900 or a 1904 worksheet) and
  its arguments as encoded values, which can hold an Excel error, `Empty`, a
  `Date`, an array or a 100,001-element call
  ([worksheet-function cases](docs/methodology/TEST_CASES.md#worksheet-function-cases)).
  The expected file holds the result and, for a failing call, its condition.
  `tests/derivations/date_layer_reference.py`, the earlier reference model
  carried over, recomputes every expected value from the contracts without VBA
  or Excel, so the values are `illustrative`; it now reads a number given as a
  `Double`, a date serial or a whole number, in its binary value, as VBA
  does, refuses an empty or nested array at a control or a value as
  `SHAPE_UNSUPPORTED`, and reads a pillar token from left to right, its first
  fault deciding its condition, as the data contract now states; no
  correction changes a converted case; `tools/check_test_cases.py`
  checks the new form, and `tools/check_derivations.py`, which now runs several
  cases at a time, compares all of them. The cases are 601 `nominal`, 246
  `boundary` and 514 `invalid`. None was run in Excel: the runner is
  [#13](https://github.com/danielep71/K-PRICING/issues/13) and the `independent` values from
  Excel are [#14](https://github.com/danielep71/K-PRICING/issues/14)
  ([#11](https://github.com/danielep71/K-PRICING/issues/11)).
- The date layer's accuracy contracts and completeness policy. An accuracy
  contract can now use the `exact` metric, for a text output such as an
  encoded date or Excel error: a result is the expected text or it is wrong,
  with threshold and target 0, and a results file holds a `text` for it.
  `ACC-02` to `ACC-23`, one for each `KPR_Dates_*` function, are `exact` and
  draft until the `independent` values of
  [#14](https://github.com/danielep71/K-PRICING/issues/14). The completeness
  policy ([validation plan](docs/methodology/VALIDATION_PLAN.md#date-layer-completeness))
  lists the case categories each function needs and the conditions each of
  its arguments, its call and its result need, and
  `tests/test_date_layer_completeness.py` fails on any that no case covers;
  145 cases were added to meet it, so the date layer has 1,506. Nothing was
  run in Excel
  ([#12](https://github.com/danielep71/K-PRICING/issues/12)).
- The project suite `date-cases`, `tests/modules/TEST_DateCases.bas`, which
  replays in Excel the 1,561 worksheet-function cases whose caller is VBA: each
  is one assertion named after the case, comparing the encoded result with its
  expected file, and one text result for the `exact` contracts, recorded with
  the harness's new `RecordTextResult`. The suite is generated from the cases
  by `tools/gen_date_case_suite.py`, and a test fails while it is stale; it is
  listed in the evidence policy, so an Excel run certifies it. Not yet run in
  Excel; in LibreOffice the facade cannot read a `Date` argument, so it is no
  early warning there
  ([#13](https://github.com/danielep71/K-PRICING/issues/13)).
- The suites `core-err-parse`, `core-dates`, `core-array`, `ranges` and
  `dates` are now certified: the evidence policy lists them with their
  certification entry points, so an Excel run records them as evidence. Their
  assertions were made in helper procedures, which the certification does not
  read; each helper is now a `Probe*` procedure that keeps the values to
  compare, and the case asserts them under literal names. The names and their
  order equal those of the Excel run of 2026-11-26, assertion for assertion
  (467, 2,083, 122, 62 and 81), and `tests/test_project_suites.py` fails when a
  project suite's source no longer reads as its counts state
  ([#13](https://github.com/danielep71/K-PRICING/issues/13)).
- The `date-domain` check, `tools/check_date_domain.py`, which **Repository
  integrity** runs: six rules of the date layer that the generic VBA checks
  do not cover, each tied to the methodology section it enforces and each with
  a self-test that makes the defect in a copy of the production source: no
  `DateSerial` day of zero or below, the window pinned in `CORE_Parse` and
  nowhere copied, the host rule once per call and before any argument,
  `Application.Volatile` only first in `KPR_Dates_HostDateSystem`,
  `Application.Caller` only in `M_Host` with no workbook fallback, and no
  conversion that depends on regional settings. The production source already
  satisfies all of them; no VBA changed
  ([#15](https://github.com/danielep71/K-PRICING/issues/15)).
- The date demo: a **Dates Demo** sheet in every build, linked from Home. It
  shows each of the 22 functions on inputs you can edit, beside a formula of
  Excel's own functions, with a status per row (`OK`, `DIFFERS` for a
  documented difference, `FAIL`, or `NO VBA` while the functions are not
  loaded). A side panel counts the statuses and lists inputs that must give
  an Excel error value. It also spills one formula over 30 dates and shows
  the array limits. `tools/gen_workbook_template.py` generates it into
  `src/workbook/Template.xlsx` from `tools/demo_dates.py`, with live
  dynamic-array formulas that Excel calculates when it opens the workbook. It
  is the one sheet with formulas, as the
  [workbook contract](docs/methodology/WORKBOOK_CONTRACT.md#dates-demo)
  records. The project suite `ui` reads its checks in Excel (case
  `demo-sheet`). The demo was opened in LibreOffice only, where Excel's own
  formulas calculate; it has not run in Excel
  ([#18](https://github.com/danielep71/K-PRICING/issues/18)).
- The workbook's menus: a **K-PRICING** Ribbon tab,
  `src/workbook/customUI14.xml`, which every build now writes into the
  workbook, and a **KPR Dates** toolbar for an Excel that does not show the tab,
  installed and removed from the macro list. Both offer the same actions,
  through `M_Actions`: insert one of the 22 functions in the active cell with
  Excel's Function Arguments dialog, register or clear their help in the
  Insert Function dialog and, in a development build only, run the project
  suites. The callbacks are in `M_Ribbon`, the toolbar in `M_CommandBars`.
  **Repository integrity** fails when the Ribbon names a callback that
  `M_Ribbon` does not declare with Office's signature, or a tag that is no
  action of `M_Actions` (`ribbon`). A new project suite, `ui`, checks the
  menus' list of functions, their actions, the cells where a function is not
  inserted (a cell of a table among them), the Ribbon's menu and the
  toolbar's lifecycle; it is in the evidence policy and has not run in Excel.
  No button has been clicked in Excel yet; the certification run does that
  ([#17](https://github.com/danielep71/K-PRICING/issues/17)).
- The workbook's layout. The [workbook contract](docs/methodology/WORKBOOK_CONTRACT.md)
  now states its one sheet, **Home**: the build stamp, what the workbook
  provides and does not cover, how to use it, and a link to every sheet. It
  also states the columns, the title band and the bordered sections, the
  navigation by cell links, the three stamp names, and that no sheet is
  protected unless a decision says so. `src/workbook/Template.xlsx`
  is generated from source by `tools/gen_workbook_template.py`, with no VBA
  project and no document properties; the build starts from it, and
  **Repository integrity** fails when it differs from its source
  (`workbook-template`) or from its protection matrix,
  `docs/methodology/workbook-protection.json`. The template was not opened in
  Excel; the certification run does that
  ([#16](https://github.com/danielep71/K-PRICING/issues/16)).
- `IMPL-01` in the source register: Excel's own worksheet functions, as the
  Excel oracle calls them, in Excel 16.0 build 20527, 64-bit. The 416 expected
  values the oracle confirmed in Excel are `independent`: each cites
  `IMPL-01`, its formula, the passing run and the step of its calculation
  contract. The holdout of `ACC-02` to `ACC-20` is now every fifth, by ID, of
  the function's covered cases, so it holds only `independent` values; the
  contracts stay `draft` until the release certification. `ACC-16` and
  `ACC-18` hold fewer than three such cases. A test keeps the values, the
  oracle and the holdouts in step
  ([#48](https://github.com/danielep71/K-PRICING/issues/48)).
- The Excel oracle, the project suite `excel-oracle`,
  `tests/modules/TEST_ExcelOracle.bas`: 416 of the 1,635 worksheet-function
  cases computed a second time with Excel's own worksheet functions only
  (`DATE`, `YEAR`, `MONTH`, `DAY`, `EOMONTH`, `EDATE`, `WEEKDAY`, `INT`, `MOD`,
  `IF`) and compared with their expected files; it calls no procedure of the
  date layer. The other 1,219 are excluded, each for a stated reason, as
  [EXCEL_ORACLE.md](docs/methodology/EXCEL_ORACLE.md) lists them. Generated
  by `tools/gen_excel_oracle_suite.py`; a test fails while it is stale, when
  a formula calls another function and when it names the date layer. Checked
  against an emulation of those functions only; not yet run in Excel, and no
  expected value changes class
  ([#14](https://github.com/danielep71/K-PRICING/issues/14)).
- The project suite `worksheet-cases`, `tests/modules/TEST_WorksheetCases.bas`,
  which replays in Excel the 74 worksheet-function cases whose caller is a
  worksheet: each is a formula in a scratch workbook of the case's date
  system, on arguments written to cells, compared as encoded text with its
  expected file and recorded for the `exact` contracts; each group of cases
  asserts that its workbook was closed and the calculation mode restored.
  Generated by `tools/gen_worksheet_case_suite.py`, and a test fails while it
  is stale. The project suite `excel-host`, `tests/modules/TEST_Host.bas`:
  one workbook changing date system on an ordinary recalculation, a
  dynamic-array formula spilling one result per row and nothing under 1904, a
  block of a workbook's cells read with its blank and its error, and the
  registration of the 22 functions, twice each way, leaving the session
  unchanged. Both are in the evidence policy; neither has run in Excel yet
  ([#13](https://github.com/danielep71/K-PRICING/issues/13)).
- Regression coverage of the date layer: 129 `regression-*`
  worksheet-function cases, so the date layer has 1,635; the cases
  `date-regression` and `long-regression` of `core-err-parse` (165
  assertions), `pillar-regression` of `core-dates` (28),
  `regression-shapes` of `core-array` (10), and four control assertions of
  `dates`. The
  [validation plan](docs/methodology/VALIDATION_PLAN.md#regression-coverage)
  maps each regression area to the suites that cover it; the calibration and
  holdout cases of `ACC-02` to `ACC-23` are recomputed by the same rule. Not
  yet run in Excel
  ([#13](https://github.com/danielep71/K-PRICING/issues/13)).

- A [user manual](docs/user-manual/0.2.0/MANUAL.md) for the people who
  receive the workbook, which needs neither the repository nor VBA: how to
  open and enable it without administrator rights and without bypassing a
  macro policy, how to check it, how each argument is read, the 22
  `KPR_Dates_*` functions with their arguments, results and error values,
  the Ribbon tab and the toolbar, and the Dates Demo sheet. The
  [installation guide](INSTALLATION.md#opening-the-workbook) gains the same
  opening steps, and the README links both. The steps were not followed in
  Excel in this change; the certification run does that
  ([#19](https://github.com/danielep71/K-PRICING/issues/19)).

### Changed

- The facade module of the 22 date functions is `KPR_DATES_DAYS`, no longer
  `KPR_Dates`, so that the planned `KPR_DATES_BUSINESS_DAYS` and
  `KPR_DATES_CALENDARS` sit beside it. The worksheet functions keep their
  names and signatures; a VBA caller that names the module, as in
  `KPR_Dates.KPR_Dates_AddDays`, uses the new name. The
  [house style](docs/VBA_HOUSE_STYLE.md#naming-and-visibility) states the
  facade naming rule, and `tools/check_vba_source.py` enforces it. Nothing was
  run in Excel ([#61](https://github.com/danielep71/K-PRICING/issues/61)).
- The public API inventory is reconciled with the code: the 22 `KPR_Dates_*`
  functions are the `[supported]` tier, with their exact signatures; the other
  60 public declarations of the core and the adapters stay `[project-public]`,
  and the inventory says why the ones that only the tests call are not
  private. The repository structure no longer says that nothing is supported
  yet. No procedure changed
  ([#9](https://github.com/danielep71/K-PRICING/issues/9)).
- The [VBA house style](docs/VBA_HOUSE_STYLE.md#layout) states the full
  layout of a procedure body: centred titles, the blank lines around
  signatures, sections, group banners and `End`, `PROC_NAME` first in
  `DECLARE`, the section order, the first error handler in `INITIALIZE`,
  declarations aligned to the first of their run, `Case Else` and explicit
  compare modes, and a table of which tool checks each rule.
  `tools/check_vba_layout.py` now reports each of those rules, and every VBA
  module, hand-written or generated, is reformatted to pass. Only layout
  changed: no procedure, parameter or public name was renamed, and
  `PUBLIC_API.txt` is unchanged. The modules were not run in Excel after the
  reformat ([#54](https://github.com/danielep71/K-PRICING/issues/54)).

### Validation

- Excel run of `573884baeb4258a53e8bc7bbe43cbb7c8c44ce5f` on 2026-11-26,
  the same host as below: every suite of the evidence policy certified,
  `excel-oracle` included, `excel-run: PASS`, `Binding: VERIFIED`. Excel's own
  worksheet functions give the expected value of all 416 cases the oracle
  covers; the values stay `illustrative` until
  [#48](https://github.com/danielep71/K-PRICING/issues/48) registers Excel
  as an independent source
  ([#14](https://github.com/danielep71/K-PRICING/issues/14)).
- Excel run of `10706ac88fd7fe04f1905f0aabe59c17e50a4695` on 2026-11-26,
  recorded with `tools/devbuild/excel_run.py` in Excel for Microsoft 365
  (build 16.0.20527, 64-bit) on Windows 11: every suite of the evidence
  policy certified, `examples`, `core-err-parse`, `core-dates`, `core-array`,
  `ranges`, `dates`, `date-cases`, `worksheet-cases` and `excel-host`;
  `excel-run: PASS`, `Binding: VERIFIED`, compile `ON_DEMAND`. The
  calibration cases of `ACC-02` to `ACC-23` have an exact error of 0; their
  holdouts stay sealed while the contracts are draft. The
  [host policy](INSTALLATION.md#%EF%B8%8F-host-policy) records the
  environment
  ([#13](https://github.com/danielep71/K-PRICING/issues/13)).
- The five date-layer suites ran in Excel for Microsoft 365 (build
  16.0.20527, 64-bit) on the development workbook, after a clean compile:
  `TEST_CoreErrParse` 467, `TEST_CoreDates` 2,083, `TEST_CoreArray` 122,
  `TEST_Ranges` 62 and `TEST_Dates` 81 assertions, all passed, with no
  failure and the Excel state unchanged. This includes the 9 `TEST_Dates`
  assertions that failed in LibreOffice. Four worksheet formulas gave the
  expected values, in the 1900 and in the 1904 date system. The suites are
  not certified and are not in the evidence policy; the cases of the 22
  functions are converted in
  [#11](https://github.com/danielep71/K-PRICING/issues/11).

[Unreleased]: https://github.com/danielep71/K-PRICING/commits/main
