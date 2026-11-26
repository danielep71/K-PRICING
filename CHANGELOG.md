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
- The 22 `KPR_Dates_*` worksheet functions, in the facade `KPR_Dates` in
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

### Changed

- The public API inventory is reconciled with the code: the 22 `KPR_Dates_*`
  functions are the `[supported]` tier, with their exact signatures; the other
  60 public declarations of the core and the adapters stay `[project-public]`,
  and the inventory says why the ones that only the tests call are not
  private. The repository structure no longer says that nothing is supported
  yet. No procedure changed
  ([#9](https://github.com/danielep71/K-PRICING/issues/9)).

### Validation

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
