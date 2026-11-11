<div align="center">

# 📜 Changelog

### Release history for K-PRICING

[![Format](https://img.shields.io/badge/Format-Keep_a_Changelog-0969da?style=flat-square)](https://keepachangelog.com/en/1.1.0/)
[![Versioning](https://img.shields.io/badge/Versioning-SemVer-6f42c1?style=flat-square)](https://semver.org/spec/v2.0.0.html)
[![Dates](https://img.shields.io/badge/Dates-YYYY--MM--DD-217346?style=flat-square)](#date-and-version-rules)
[![Staging](https://img.shields.io/badge/Staging-Unreleased_first-d97706?style=flat-square)](#unreleased)
[![Contributing](https://img.shields.io/badge/Changes-Contribution_guide-2ea44f?style=flat-square)](CONTRIBUTING.md)

<br>

**User-visible history · Explicit compatibility · Reproducible evidence · Immutable releases**

</div>

---

All notable changes to **K-PRICING** are documented here.

This changelog follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and [Semantic Versioning](https://semver.org/spec/v2.0.0.html). It records
released behavior and material unreleased changes; it is not a commit log, issue
tracker, or substitute for release evidence.

Define the versioned public surface—API, behavior, defaults, errors, data
formats, compatibility, and supported environments—in maintained project
documentation before the first functional release.

---

## 🧭 Maintenance policy

- Add material changes under **Unreleased** in the same pull request as the
  behavior or documentation they describe.
- Write from the user's perspective: describe the observable result, contract,
  compatibility impact, and migration need.
- Link the owning issue or pull request when it contains useful engineering
  detail.
- Keep entries concise; do not duplicate implementation notes already preserved
  in source, issues, or technical documentation.
- Record only validation actually performed. State skipped environments and
  known limitations plainly.
- Move Unreleased entries into a dated version section on the release-section
  cut/freeze date defined below.
- Do not edit a published release entry except to correct a demonstrable factual
  or link error; annotate material corrections instead of rewriting history.
- Never claim that a tag, binary, workbook, hash, test run, or environment was
  certified unless the evidence binds it to the released source.

See [CONTRIBUTING.md](CONTRIBUTING.md) for change and evidence requirements and
[SECURITY.md](SECURITY.md) for private vulnerability reporting.

<a id="date-and-version-rules"></a>

### Date and version rules

| Rule | Standard |
|---|---|
| Version | `MAJOR.MINOR.PATCH`, without the leading `v` in headings |
| Release heading | `## [X.Y.Z] - YYYY-MM-DD` |
| Date | Gregorian ISO `YYYY-MM-DD` release-section cut/freeze date; it may precede tag creation or publication |
| Ordering | Unreleased first; released versions newest to oldest by SemVer precedence; cut/freeze dates do not move backward |
| Comparison | Unreleased → latest tag; each later release → preceding tag; initial release → release tag |
| Patch | Backward-compatible correction or hardening |
| Minor | Backward-compatible capability |
| Major | Incompatible public-contract change |
| Pre-release | Strict SemVer identifiers; numeric identifiers have no leading zeros |

A repository may remain below `1.0.0` while its supported surface is still
forming. Pre-release status does not excuse undocumented breaking changes.

<details>
<summary><strong>Entry categories</strong></summary>

<br>

| Category | Use for |
|---|---|
| **Added** | New supported capabilities, APIs, files, or tests |
| **Changed** | Changes to existing behavior, contracts, tooling, or documentation |
| **Deprecated** | Supported behavior scheduled for removal |
| **Removed** | Removed capabilities or compatibility |
| **Fixed** | Corrected defects |
| **Security** | Safely disclosed security corrections |
| **Documentation** | Material documentation-only changes |
| **Validation** | Evidence actually produced |
| **Compatibility** | Upgrade or migration effects |
| **Known limitations** | Deliberate, unresolved boundaries |

Use only the categories needed by a release.

</details>

---

<a id="unreleased"></a>

## [Unreleased]

<!-- Add only user-visible changes made in this generated project. -->

## [0.0.4] - 2026-11-10

The first functional release of K-PRICING: the Date Primitives layer. It
provides exactly 22 supported `KPR_Dates_*` worksheet and VBA functions in one
scalar/array-capable surface, with the **KPR Dates Demo** sheet, the
**K-PRICING** Ribbon tab and the classic **KPR Dates** toolbar. The supported
behavior is defined by [`docs/DATE_LAYER_CONTRACT.md`](docs/DATE_LAYER_CONTRACT.md)
and the supported names by [`docs/PUBLIC_API.txt`](docs/PUBLIC_API.txt).
v0.0.1 (repository setup) and v0.0.2 (repository migration) were not
functional releases and have no tag. Certification of this candidate in
Windows Excel is recorded in #52.

### Added

- The v0.0.4 release candidate (#51): `VERSION` is `0.0.4`, this section is cut
  from the reviewed Unreleased entries, and the README, installation guide and
  security policy describe the 22-function surface, its compatibility claim and
  its boundaries.

- Static checks and milestone-register drift monitoring for the v0.0.4
  candidate (#50). The deterministic `kpr-candidate-tree` rule of
  `tools/check_kpr_contract.py` requires exactly the 16 approved VBA components
  and their roles, the candidate's RibbonX, template, fixtures, contracts,
  roadmap, register and tooling, classified RibbonX callbacks, parsable RibbonX
  and evidence schema, and generated fixtures marked as generated; it refuses
  tracked run output, built workbooks and calendar, business-day or `_Spill`
  placeholders (15 degraded self-tests, one removing `docs/ROADMAP.md`).
  Separately, `tools/milestone_register.py` renders milestone v0.0.4 into
  `docs/MILESTONE_REGISTER.md` and checks it against the live issues: `CURRENT`,
  `DRIFT` or `UNVERIFIED`, where an API, authentication, network or
  malformed-response failure is never drift. The new `Milestone register`
  workflow runs that check on relevant pushes, issue events and on demand, with
  read-only permissions and no `pull_request` trigger; the static gate runs only
  its frozen-snapshot tests. A fresh `CURRENT` is required at candidate assembly
  and immediately before tagging (`RELEASING.md`).

- `docs/API_INVENTORY.txt` classifies all 150 technically public VBA members
  (#49): the 22 supported `KPR_Dates_*` functions with their vectorization, and
  128 unsupported infrastructure members in the `internal`, `registration`,
  `ui`, `test`, `development` and `demo` classes. It records `KPR_Cal_*` as
  reserved for v0.0.5 and the calendar-aware v0.0.6 boundary. The new
  `kpr-api-inventory` rule of `tools/check_kpr_contract.py` keeps it complete
  and checks that the supported set, vectorization and signatures agree across
  the contract, the facade and `docs/PUBLIC_API.txt`. Every name of a
  comma-separated public declaration is inventoried, and `KPR_Cal_*` names are
  refused at any visibility; 15 degraded self-tests cover the rule.

- The K-PRICING workbook is built from a template, as in K-SACCR.
  `src/workbook/KPR_Template.xlsx` is the macro-free template (a Home sheet and
  the K-PRICING Ribbon), generated by `tools/workbook_template.py` and checked
  byte for byte in CI; it is the one tracked Office package. `tools/Check-KPR.ps1`
  now runs the regression in a development workbook saved from it, then builds
  `K-PRICING-<commit>.xlsm` in the `build` folder of your clone (`-RepoRoot`,
  remembered for later runs; `Documents\K-PRICING` without one): the
  production and example modules, compiled, the commit on Home and the KPR
  Dates Demo sheet built once by `KPR_Demo_BuildDates`.

- `tools/Check-KPR.ps1` runs the Excel check of a branch with no manual
  step: one line pasted into Windows Run (Win+R), the desktop shortcut it
  creates on its first run, or `tools\Check-KPR.cmd` from a clone. It starts
  its own Excel, builds a fresh `KPR.xlam` (since replaced by the template
  build, above) from the branch's exact commit with
  the development loader, compiles it, runs `KPR_Test_RunAll` twice and loads
  the Ribbon, then zips the evidence and copies it to the clipboard. It asks
  before turning on *Trust access to the VBA project object model*, and a
  watchdog closes its Excel after 30 minutes. `KPR_Dev_CheckBuild` gains an
  optional `Unattended` argument (no message box or Explorer window) and
  `KPR_Dev_LastCheckFolder` returns the folder it wrote.

- Added a one-click Excel check to the development loader
  (`tests/modules/KPR_Dev_Import.bas`): the macro `KPR_Dev_CheckLatest`
  (Alt+F8) resolves `main` to its exact commit, imports it into `KPR.xlam`,
  compiles and saves it, runs `KPR_Test_RunAll` twice, checks on Windows that
  Excel loads the K-PRICING Ribbon in a throwaway copy, and leaves the two
  evidence files and a `summary.txt` in an evidence folder (first
  `Documents\KPR-evidence`; since #118 the clone's `build\evidence` or
  `%LOCALAPPDATA%\K-PRICING\evidence`, see Changed).
  `KPR_Dev_CheckBuild` does the same for any branch, commit or build, and the
  Ribbon module gains `KPR_UI_RibbonLoaded` for that check. Run in Windows
  Excel at `ba250a7` (#47, #48).
- Added the classic **KPR Dates** toolbar (#48) and the shared menu commands it
  calls. `src/modules/KPR_UI_Actions.bas` holds the commands: insert a
  supported function into the active cell with Excel's Function Arguments
  dialog, build a catalog demo sheet, register or clear the Function Wizard
  help, and, only when the regression modules are loaded, the pure and durable
  regression runs. `src/modules/KPR_UI_CommandBars.bas` adds
  `KPR_UI_InstallCommandBar` and `KPR_UI_RemoveCommandBar`, which are safe to
  run repeatedly in any order, create temporary controls only and remove only
  K-PRICING's own; the macro list (Alt+F8) offers them as
  `KPR_UI_ToolbarInstall` and `KPR_UI_ToolbarRemove`. An inserted function is
  written through `Formula2` where Excel has it, so a vectorized call spills,
  and cancelling the Function Arguments dialog gives the cell back its exact
  previous content. Demo catalog entries gain a group, so the menus show one
  submenu per functional module. The new `worksheet-commandbars` regression
  suite sets the evidence record's `commandbars` outcome. Verified in Windows
  Excel at `ba250a7`: `worksheet-commandbars` PASS.
- Added the **K-PRICING** Ribbon tab (#47). `src/ribbon/customUI14.xml` (Office
  2010 schema, stable `KPR_` ids, built-in icons) has a **Date primitives**
  group whose Functions and Demo sheets menus are rebuilt every time they open
  from the registration manifest and the demo catalog, a **Setup** group, and
  a **Development** group shown only when the regression modules are loaded.
  The callbacks in `src/modules/KPR_UI_Ribbon.bas` only hand the control's tag
  to the shared commands or build a menu; `onLoad` keeps the Ribbon object for
  `KPR_UI_RibbonRefresh`, the only invalidation. `tools/inject_ribbon.py
  --target` writes the tab into one explicitly named, untracked `.xlsm` or
  `.xlam` after validating the package: every other entry is copied
  byte-for-byte, exactly one ribbon relationship is kept, repeated runs give
  identical bytes, and the file is replaced atomically after an exclusive
  backup. `--check`, now part of Repository integrity, validates the RibbonX
  against the Office 2010 signatures of its VBA callbacks. The new
  `worksheet-ribbon` suite parses the menu XML the callbacks return; an `all`
  run now has 23 suites. Verified in Windows Excel at `ba250a7`:
  `worksheet-ribbon` PASS, and the tab loaded in the workbook built from
  `src/workbook/KPR_Template.xlsx`.
- Added the development loader `tests/modules/KPR_Dev_Import.bas`:
  `KPR_Dev_ImportModules(Source, TargetWorkbook)` replaces the K-PRICING
  modules of an open workbook or add-in with those of one commit, downloaded
  from GitHub or read from a local clone, using the commit's repository
  profile as the module list. Repository modules the commit does not list
  are removed, so an older commit's set is restored; other modules of the
  target are left alone and the target is not saved. Used by every Windows
  Excel run of the one-click check, including `ba250a7`.
- Portable native developer commands for logical demo-workbook snapshots and
  cell/property differences, clean-candidate source inventories, supplied VBE
  export round trips and read-only diagnostics (#46, #50, #52). Existing source
  validators remain authoritative. The candidate inventory includes the current
  demo modules; the historical host-evidence inventory is unchanged. Formula
  caches are explicitly unverified; automated host builds and final release
  certification remain outstanding.
- Added demo sheets built from tracked source (#46): the shared
  `examples/modules/KPR_Demo__Builder.bas` owns layout, house style, status
  rules, the checks summary, the Excel state and the demo catalog, and
  `examples/modules/KPR_Demo_Dates.bas` describes the Date Primitives sheet
  through it. `KPR_Demo_BuildDates()` adds a **KPR Dates Demo** sheet to the
  active workbook (never changing, replacing or deleting an existing sheet, and
  falling back to a new workbook for a 1904, read-only or protected one), or
  builds into a new `.xlsx` file that must not exist yet. Each row compares a
  live KPR formula, driven by editable named inputs, with a native-Excel
  reference and shows `OK`, `DIFFERS` (by design), `FAIL` or `NO VBA`; the side
  panel adds native-error cases, dynamic-array spills and the 100,000-element
  limit. The new `worksheet-demo` regression suite checks file and
  active-workbook builds and sets the evidence record's `demo_generation`
  outcome. Verified in Windows Excel (64-bit, build 20326) at `cd0e1a1`: two
  matching `all` runs with 21 suites, 9,320 assertions and 0 failures.
- Added MacroOptions registration (#42): the internal
  `KPR_REGISTER_PUBLIC_UDFS` module holds a machine-readable manifest of the 22
  supported functions and registers them under the single **KPR Dates**
  category with a description and complete argument help
  (`KPR_Register_PublicUDFs`), or blanks that metadata and returns them to
  **User Defined** (`KPR_Register_ClearPublicUDFs`). Both are idempotent,
  callable through `Application.Run`, preserve the workbook's saved state and
  never change a calculation. `check_kpr_contract.py` checks the manifest
  against the façade, and the new `worksheet-registration` suite sets the
  evidence record's `macro_options` outcome. Verified in Windows Excel; the
  durable runner also hands the status bar back with a literal `False`.
- Added the Excel cross-oracle checks (#41): `KPR_Test_Oracle` compares the
  date primitives with native `EOMONTH`, `EDATE`, `WEEKDAY`, `DAY`, `YEAR` and
  `MONTH` where the contracts overlap, on 22 boundary dates and 300 fixed-seed
  samples (7,224 assertions), asserting `#NUM!` where Excel's answer leaves the
  supported window and the Gregorian answers for Excel's fictitious
  29-Feb-1900. It runs as the `worksheet-oracle` suite and sets the evidence
  record's `cross_oracle` outcome; `check_kpr_contract.py` rejects any other
  worksheet function in an oracle formula. Verified in Windows Excel.
- Added the regression matrix (#40): the fixture generator now emits a
  systematic `matrix` family for each of the 21 value-taking functions and
  each value argument (valid, edge, invalid-domain, propagated, 1x1, row,
  column and rectangle cases, and 1900/1904 worksheet calls), bringing the
  fixtures to 1,353 cases, and fails generation if any matrix case is missing.
  The durable runner adds a `worksheet-state` suite that proves caller-state
  restoration after a deliberately failing and a passing nested run. Verified
  in Windows Excel on `33b07cd` (18 suites, 1,953 assertions).
- Added the durable VBA test runner and structured evidence schema (#39):
  `KPR_Test_RunAll` and `KPR_Test_RunSuite` run a 17-suite registry, including
  the 981 generated fixture cases (direct VBA and 1900/1904 worksheet callers)
  and the worksheet runners, restore the caller's Excel state, and write
  `kpr-test-evidence.json` bound to an exact source SHA with the #52
  certification outcomes. `docs/kpr-test-evidence.schema.json` defines the
  record and `tools/check_test_evidence.py` validates it. Verified in Windows
  Excel on `9a978cd`: two runs of 17 suites and 1,577 assertions.
- Added the independent date-layer fixture generator (#38): `tools/gen_fixtures.py`
  computes 981 expected results from the contract with Python standard-library
  date arithmetic, writes the canonical `tests/fixtures/date_layer_fixtures.tsv`
  and derives the generated VBA module `KPR_Test_Fixtures_Generated` from it.
  Hosted CI runs its self-test and a drift check; the fixtures run in Excel
  through the durable runner (#39).
- Added `docs/ROADMAP.md`, the authoritative forward roadmap: v0.0.3 test
  hardening and registration (#38–#42), v0.0.4 date primitives and the first
  functional release (#46–#52), and planned v0.0.5 business calendars and
  v0.0.6 business-day arithmetic, with dependencies and scope transitions.
- Recorded v0.0.2 migration acceptance in `docs/MIGRATION_COMPLETION.md`, with
  the accepted candidate, inventory, evidence, limitations and the v0.0.3 next
  delivery scope (#38–#42); no release, tag or version change.
- Retained the #17 exact-source Excel parity evidence for candidate
  `db98e506358ab74893ab28d52183fa2e216352b9`: the frozen KPR source compiled
  and passed its own 557 checks, the destination passed 568 assertions, and
  source and destination observations matched exactly on one Windows 64-bit
  Excel host; the #32 correction is registered as a known source difference.
- Added a synthetic worksheet usage example for the migrated date API, covering
  ISO inputs, month-end clipping, pillars, dynamic-array spills and the 1904
  date-system boundary; its results are contract expectations, since the #17
  parity run did not execute those exact formulas.
- Added the KPR migration handover register: source-to-destination work
  disposition, destination-only corrections, the carried-forward roadmap and
  the reference policy for the frozen source.
- Imported the frozen KPR date and VBE export contracts with verified source
  hashes and license attribution (#11).
- Imported the frozen KPR date layer (four internal cores and the 22-function
  `KPR_DATES_DAYS` façade), the migrated regression harness, a direct-VBA
  example, the public API manifest and KPR contract static checks (#12, #13);
  defined the migration regression/parity protocol and evidence validator (#14).
- Accepted manual neutral-starter Excel evidence for candidate
  `c8f0b6ee147d07549484ec83743f5b7dfc0ea1f2` on one Windows/64-bit Office host; no pricing
  or application-release certification.
- Project-specific developer setup, neutral Excel runbook, dated GitHub verification
  and a frozen-source migration inventory with issue traceability.
- Initialized the K-PRICING application scaffold with project identity, private
  reporting contact, profile directories and reproducible initialization inputs.

### Changed

- The Excel check keeps its evidence on the computer: `tools/Check-KPR.ps1`
  passes `<clone>\build\evidence` to `KPR_Dev_CheckBuild`, whose new optional
  `EvidenceRoot` argument receives it, and without a clone the evidence goes
  to `%LOCALAPPDATA%\K-PRICING\evidence`. `Documents`, often synced to
  OneDrive or SharePoint, is no longer used. A commit whose loader predates
  the argument still runs with three arguments.

- The workbook template's Home sheet follows the K-SACCR template: Aptos, a
  black navigation column with Home and a link to KPR Dates Demo, the black
  band with the gold wordmark, the blue sheet-title band, a grey subtitle,
  frozen rows 1 to 3 and column A, and navy sections with labels in column C
  and text in column D (what K-PRICING does, how to use it, its sheets and the
  commit, now written in `D22`).

- Label synchronization no longer prunes a label that an issue or pull request,
  open or closed, still carries: before any change it checks every label it
  would delete (a paginated scan for names with a comma, which the labels
  filter cannot express) and fails on a used label, a malformed response or an
  exhausted page limit. The reconcile job declares `pull-requests: read` for
  that check. The label sync and drift summaries share one Markdown helper that
  fences code spans longer than any backtick inside the value, escapes table
  pipes and turns line breaks into spaces. The PR-milestone scan reads each
  removal back with a separate GET and lists every pull request and removed
  milestone title in its job summary.

- Quality gates follow the template more closely. mypy adds `warn_unreachable`
  to strict mode (no existing finding). Static checks cancel only a superseded
  pull-request head: every push, tag, manual or reusable run has a concurrency
  group of its own and is never cancelled. `tools/check_documentation.py`
  checks documented commands per subcommand, including subcommands named from
  a literal loop, short options, `unittest` options and the `py`, `py -3` and
  `-m tools.<module>` forms, and fails on a parser it cannot read statically.
  `.gitignore` also excludes Java and password-manager keystores (`*.jks`,
  `*.keystore`, `*.kdbx`) and SSH private keys (`id_rsa*`, `id_ecdsa*`,
  `id_ed25519*`), keeping their `.pub` halves trackable.

- The repository's Python tooling moves from Python 3.10, which reached end of
  life in October 2026 and no longer receives security fixes, to 3.13: CI pins
  3.13.18, `pyproject.toml` targets 3.13 for Ruff and mypy, and the hash lock in
  `tools/requirements-quality-ci.txt` now holds the CPython 3.13 wheels of the
  same tool versions (`tomli`, needed only before 3.11, is dropped). Running the
  tools locally needs Python 3.13 or newer. This unblocks Dependabot's tool
  updates, starting with coverage 7.17, which requires Python 3.11 or newer.
  The VBA source and the Excel requirements are unchanged.

- Issue governance now follows the reference template ([issue triage
  guide](docs/ISSUE_TRIAGE.md)). The issue forms set a default priority (`P2`
  for bugs, `P3` for features and documentation), assign the maintainer and ask
  for a target end date and dependencies; the documentation form also asks for
  acceptance criteria. `P1` is reserved for confirmed calculation bugs (#52
  moved to `P2`). A read-only **Issue triage** workflow audits every open issue
  on issue and milestone events with `tools/check_issues.py`, and a daily
  **Keep milestones issue-only** scan removes milestones from pull requests.
  The pull request template records the reviewed head SHA and requires a new
  review after each relevant push.

- Corrected stale, inaccurate and template-inherited comments across the VBA
  modules, Python tooling and configuration (no behaviour change). Issue
  numbers carried over from the migrated source repository now point to this
  repository's issues, the issue forms ask for the affected area instead of a
  template profile, and the demo summary no longer refers to a menu that does
  not exist yet.
- Relicensed K-PRICING from the MIT License to the Mozilla Public License
  2.0 (MPL-2.0). `LICENSE` holds the unmodified MPL 2.0 text; the copyright
  notice is in the README. Commits published before the change remain
  available under MIT.
- Demo modules share one naming rule: every demo module is named
  `KPR_Demo_*`, and the shared builder is `KPR_Demo__Builder` (double
  underscore) so it sorts first. `KPR_Demo_Builder` became `KPR_Demo__Builder`
  and the direct-VBA example `KPR_DateExample` became `KPR_Demo_DirectVBA`;
  its macro is still `RunDateExample`.
- Integer arguments now accept a number within `1E-9` of a whole number as
  that whole number, so binary floating-point residue such as
  `2.9999999999999996` no longer returns `#VALUE!`; a real fraction such as
  `2.5` or `3.00000001` still returns `INTEGER_FRACTION`. The contract also
  fixes three v0.0.4 decisions: `DayOfWeek` keeps Monday = 1 (the reverse of
  `WEEKDAY`), `ON`/`TN` pillars stay calendar days in `KPR_Dates_*` (business
  days arrive through v0.0.6 calendar-aware functions), and performance is
  measured with an external VBA performance add-in. The fixtures grow to
  1,358 cases.
- Hosted static checks now measure statement coverage of the Python tooling in
  a parallel **Python tooling coverage** job and fail below the `fail_under`
  ratchet (80%, measured baseline 81%). This replaces the unenforced 95%
  setting, which pointed at a workflow removed at initialization.
- Tightened the Python tooling quality gates to the reference template:
  Ruff's McCabe complexity ceiling drops from 20 to 15, with the 16 functions
  above it split into named helpers (no behaviour change), and mypy now runs in
  `strict` mode over every `tools/` module instead of `_gatelib` alone. Three
  unittest modules relax only the annotate-every-definition rules.
- The weekly external-link observation now probes static badge queries on
  `img.shields.io`, an approved query domain, instead of reporting all 37
  styled badges as access restricted. The README CI badge no longer carries a
  query, the actionlint checksum reference points to its public release page,
  and the private vulnerability-reporting form is a dated exception. The
  workflow had failed on every run without finding a broken link.
- Adapted the date-layer contract's version labels to the K-PRICING milestones:
  the contract targets v0.0.4 and `KPR_Cal_*` is reserved for v0.0.5. Behavior,
  signatures, defaults and errors are unchanged; provenance records the
  adaptation.
- The migration handover register, completion record and README now map the
  former unscheduled demo/UI, classification, controls, assembly and
  certification items to #46–#52 and defer current planning to the roadmap.
- Refreshed stale Markdown across the repository: the README states the fixed
  application profile instead of the inherited profile taxonomy; removed
  neutral-starter examples and template-only procedures from current guidance;
  marked the neutral-starter runbook and initialization procedure as historical
  or reference material; recorded #17, #18 and #38 as complete; and added the
  generated fixture module to the Excel evidence import inventory.
- Updated the migration plan to record completed migration issues with their
  merge commits, the source implementation-plan disposition and the #17/#18
  outcomes; the production source guide now states which modules are
  byte-identical to the frozen KPR source and which carry destination deltas.
- Reset the development version to `0.0.0` and removed template-maintenance
  history, assets and tools through the deterministic initializer.

### Security

- A manual dispatch of the label-sync workflow from any branch other than the
  default branch now skips reconciliation, so an unmerged label manifest or
  script can no longer change the live labels. The repository gate rejects any
  workflow job with `issues: write` that lacks this default-branch guard.
- CodeQL also analyzes the workflows themselves (Actions language). CI pins
  the Python and Node.js versions instead of the runner's choice (first Python
  3.10.22, now 3.13.18; Node.js 24.22.2), and
  Dependabot also proposes updates to the hash-locked Python quality tools;
  its pull requests carry a `P3` priority and an assignee.

- Enable public-repository CodeQL/Scorecard paths, private vulnerability reporting
  and branch/tag rulesets; retain conditional safeguards for private visibility.
- Record the maintainer's 2026-09-24 decision to retain the existing email and
  host details publicly, as documented in setup verification and issue #29.

### Fixed

- Before certification, a second audit aligned the remaining texts and models:
  the fixture generator's numeric-text rule now matches the parser (`1.`,
  `.5` and `+.5` are `DATE_TEXT_NUMERIC`; three fixture cases, 1,361 in all,
  and a `-1.` regression assertion), and contract section 3.1 states the
  locale-date and numeric-text shapes both use. The forwarded-`Opt_Rounding`
  regression test uses two spans (10 and 11 days) that, together, only `NEAREST`
  maps to `1W` and `2W`.
  `tools/Check-KPR.ps1` chooses the loader's arguments from its source instead
  of retrying after any error, so a failed run is reported, not repeated. The
  candidate's README, installation guide, test-evidence notes and provenance
  tables, and the internal documentation, are corrected.

- Date text is reported as `DATE_TEXT_LOCALE` only when it is shaped like a
  locale-formatted date (one or two digits, one or two digits, then two to four
  digits, separated by `/` or `.`, as in `31/12/2026`), as contract section 3.1
  and the fixture generator define it. Year-first text such as `2024/02/10`,
  a dotted time suffix and `1.2.3` are `DATE_TEXT_FORMAT`; both still return
  `#VALUE!`, so no worksheet value changes. Eight regression assertions pin the
  boundary.
- A direct VBA caller that forwards its own omitted `Optional` argument to
  `Opt_WeekBaseMonday` or `Opt_KeepEOM` now gets the documented default, as it
  already did for `Opt_Rounding`, instead of the forwarded `Missing` error value.
  Three regression assertions cover the three controls. The contract states the
  rule, and section 3.3 now agrees with section 7 that an object, multi-area
  Range or empty array control is `SHAPE_UNSUPPORTED`.
- Documentation consistency before certification: the release and
  contribution guides name `KPR_Test_RunAll` as the certification entry point;
  stale "not yet run in Excel" notes in this section, `0.0.0` version
  statements, module counts, import orders, layout tables, roadmap status and
  milestone titles, security risk surfaces and the README's claim about the
  milestone register are corrected; `tools/README.md` describes the CI tools.
- The dependency matrix of `tools/check_kpr_contract.py` also covers
  `KPR_Demo_DirectVBA` (facade only) and `KPR_Dev_Import` (no module calls),
  with two degraded self-test cases (96 in the whole self-test). Repository
  integrity now runs the
  `release_certification.py` and `_release_closeout.py` self-tests, and
  `.editorconfig` no longer keeps Markdown trailing whitespace the
  whitespace gate rejects.
- The one-click Excel check's Ribbon step failed on the first real run ("Excel
  did not load the tab", and a bare "400" box) although the build passed: the
  loader opened a copy from inside a running macro, and Excel loads a
  workbook's Ribbon only while no macro runs. Unattended, `KPR_Dev_CheckBuild`
  now leaves that step to `tools/Check-KPR.ps1`, which reopens the built
  workbook once the macros have returned and requires its K-PRICING tab and
  Functions menu. The Alt+F8 path keeps the copy check.

- The one-click Excel check works where a policy forbids PowerShell or Excel
  from starting `powershell.exe` ("Access is denied"). Its watchdog runs in a
  runspace of the same process, it can be pasted into an open PowerShell
  window as `&([scriptblock]::Create((irm <url>)))`, and when the loader's
  Ribbon check cannot start PowerShell it opens the copy as it is: a workbook
  built from the template already carries the commit's tab, and the result
  says so.

- `KPR_Dates_DateFromPillar` now records `PILLAR_AGGREGATE_RANGE`, as the
  contract and fixtures specify, when a valid pillar shifts the result outside
  the supported window; it recorded `RESULT_WINDOW` (#87). The result is
  `#NUM!` either way, so no worksheet or VBA value changes.
- The durable runner's status-bar hand-back now records each step and what
  the bar read afterwards in the state-restoration detail and, when Excel keeps the text `FALSE`
  after the retries, tries two more hand-backs that need no Excel 4.0 macros
  or Trust Center change (#77). A text `FALSE` still fails the run. Verified in
  Windows Excel at `732db13`: two matching `all` runs, 0 failures, state restored.
- Corrected the Function Wizard error text of `KPR_Dates_DaysInYear` and
  `KPR_Dates_IsLeapYear`: a year outside 1900-9999 returns `#VALUE!`
  (`DOMAIN_YEAR`), and only a year beyond the `Long` range returns `#NUM!`.
  `TryPillar_Format` gained the local error handler every other `Try*`
  routine has, the month-shift window gate uses its named constant, and two
  misleading source comments were corrected. No calculation changed.
- Classify grammatically valid pillar quantities and aggregates that exceed the
  parser's numeric domain as `PILLAR_AGGREGATE_RANGE` / `#NUM!`, while
  preserving parser outputs on failure (#32).
- Reconcile completed setup status, generated-project dependency guidance and
  closeout inputs; enforce retained verification-depth tests in CI and remove
  unused template dependency locks.
- Stop release tagging/pushing after any failed prerequisite or check; reconcile
  generated-mode, private-security and release-closeout documentation.
- Documentation presentation checks now validate generated-project identity and
  selected assets instead of requiring the upstream template's public badge and
  banner after initialization.
- Retained validation suites no longer import removed template-maintenance
  tools. Their inapplicable consumer-generation, portfolio, Wiki and checker
  maintenance cases are removed; operational validator coverage remains.
- Replace stale pull-request setup fields with K-PRICING identity, current source
  authorities, validation commands and entry points.
- Remove obsolete private-link classifications after the public-visibility
  read-back; retain honest anonymous reachability results.
- Reconcile visibility records, documentation authorities, ASCII source rules
  and the migrated regression harness visibility exception.
- Allow the retained `KPR_` VBA namespace through the project identity policy, while preserving unrelated donor and
  template restrictions ([#2](https://github.com/danielep71/K-PRICING/issues/2)).

### Compatibility

- Scalar calls are claimed only for the Excel builds exercised in the
  exact-source Windows certification of this candidate (#52). Multi-cell calls
  are claimed only on dynamic-array Excel; there is no Ctrl+Shift+Enter or other
  legacy multi-cell claim and no `_Spill` twin.
- Worksheet calls use the 1900 date system of the calling workbook; a 1904
  workbook gets a call-level `#N/A` from every value-taking function, never a
  shifted value. Direct VBA, `Application.Run` and the Immediate window use the
  documented 1900 serial contract, with no active-workbook fallback.

### Known limitations

- The tested hosts are Windows 64-bit Microsoft 365 Excel only. 32-bit Office,
  other builds and locales, Mac and Excel for the web are untested and not
  claimed.
- Business calendars, weekend masks and holidays (`KPR_Cal_*`) are v0.0.5 scope;
  business-day arithmetic and roll conventions are v0.0.6 scope. `KPR_Dates_*`
  functions count calendar days only.
- Registration, Ribbon, toolbar, demo and test entry points are unsupported
  infrastructure ([`docs/API_INVENTORY.txt`](docs/API_INVENTORY.txt)).
- The accepted neutral-starter Excel run remains setup history only and does not
  certify the date layer.

---

<!--
Release procedure:
1. Move applicable Unreleased entries under ## [X.Y.Z] - YYYY-MM-DD using the release-section cut/freeze date.
2. Remove empty categories.
3. Set [Unreleased] to compare the new latest tag to HEAD.
4. For the initial release, link directly to its release tag. For every later
   release, compare the preceding tag to the new release tag.
5. Recreate an empty Unreleased section at the top.
-->

[Unreleased]: https://github.com/danielep71/K-PRICING/compare/v0.0.4...HEAD
[0.0.4]: https://github.com/danielep71/K-PRICING/releases/tag/v0.0.4
