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

[Unreleased]: https://github.com/danielep71/K-PRICING/commits/main
