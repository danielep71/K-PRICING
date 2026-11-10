<div align="center">

# ⚡ K-PRICING

### Financial analytics and instrument pricing for Excel/VBA

**K-PRICING is an Excel/VBA project for practitioners developing transparent financial analytics and instrument pricing.**

<br>

[![Excel VBA](https://img.shields.io/badge/Excel_VBA-source--first-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](#requirements)
[![Profile](https://img.shields.io/badge/Profile-application-6f42c1?style=for-the-badge)](#application-profile)
[![Version](https://img.shields.io/badge/Version-VERSION_file-0969da?style=for-the-badge)](VERSION)
[![License](https://img.shields.io/badge/License-MPL--2.0-2ea44f?style=for-the-badge)](LICENSE)

<br>

[![Static checks](https://github.com/danielep71/K-PRICING/actions/workflows/static-checks.yml/badge.svg)](https://github.com/danielep71/K-PRICING/actions/workflows/static-checks.yml)

<br>

**Source-first VBA · Explicit contracts · Deterministic evidence**

[Quick start](#quick-start)
&nbsp;·&nbsp;
[Profile](#application-profile)
&nbsp;·&nbsp;
[Structure](#repository-shape)
&nbsp;·&nbsp;
[Validation](#validation)
&nbsp;·&nbsp;
[Documentation](#documentation)
&nbsp;·&nbsp;
[Security](SECURITY.md)
&nbsp;·&nbsp;
[Supply chain](docs/SUPPLY_CHAIN_ASSURANCE.md)

</div>

---

<p align="center">
  <a href="assets/kpr-social-preview.png">
    <img src="assets/kpr-social-preview.png"
         alt="KPR — multi-screen market quotations, rate curves and instrument-pricing analytics for Excel/VBA"
         title="KPR — financial analytics and instrument pricing for Excel/VBA"
         width="100%">
  </a>
</p>


---



## ✨ What this project is

**Status: version `0.0.4`, the first functional release candidate: the
[Date Primitives](#date-primitives) layer.** It is certified only by the
exact-source Windows Excel run recorded in #52; until that evidence is complete
and the `v0.0.4` tag exists, it is a candidate, not a release. The earlier
milestones were not functional releases: v0.0.1 set up the repository (the
neutral starter passed a manual Windows 64-bit Excel run at commit
`c8f0b6ee147d07549484ec83743f5b7dfc0ea1f2`, see
[the setup completion record](docs/SETUP_COMPLETION.md)) and v0.0.2 migrated
the frozen KPR date layer (exact-source parity on one Windows 64-bit host in #17,
[evidence](evidence/migration-2026-09-24/session.txt); accepted in #18,
[completion record](docs/MIGRATION_COMPLETION.md)). See the
[initialization status](docs/INITIALIZATION_STATUS.md) for provenance and setup boundaries.

The product and repository name is **K-PRICING**; **KPR** is its short form,
used for the **`KPR_`** VBA module and function namespace.

K-PRICING is a source-first Excel/VBA project. Exported VBA, tests,
versioned policy and documentation are the reviewable source of truth; Office
packages are generated or release artifacts unless an exact path is explicitly
governed.

K-PRICING separates five evidence layers that must not be conflated:
repository integrity, VBA compilation, regression execution, specialist
assurance, and release certification.

| Principle | Project contract |
| --- | --- |
| Source identity | Exported text source and exact Git revision |
| Public surface | Explicitly documented and machine-checked |
| State ownership | Caller/host state is changed and restored deliberately |
| Failure behavior | Invalid inputs, errors and cleanup are part of the contract |
| Evidence | Claims name the exact candidate and environment tested |
| Portability | Supported Excel/Office environments are stated, not inferred |

<a id="quick-start"></a>

## ⚡ Quick start


### 1. Review the source contract

The migrated date layer contains:

- four internal [`KPR_Core_*`](src/core/) modules for errors, parsing, dates and arrays;
- [`KPR_DATES_DAYS`](src/modules/KPR_DATES_DAYS.bas) — the supported 22-function worksheet façade;
- [`KPR_REGRESSION_TESTS`](tests/modules/KPR_REGRESSION_TESTS.bas) — the imported focused regression harness;
- [`KPR_Test_Fixtures_Generated`](tests/modules/KPR_Test_Fixtures_Generated.bas) — independently generated fixtures (see [`tests/fixtures/`](tests/fixtures/README.md)); and
- [`KPR_Demo_DirectVBA`](examples/modules/KPR_Demo_DirectVBA.bas) — a minimal direct-VBA consumer example; and
- [`KPR_Demo__Builder`](examples/modules/KPR_Demo__Builder.bas) — the shared builder of demonstration sheets, and [`KPR_Demo_Dates`](examples/modules/KPR_Demo_Dates.bas) — the Date Primitives demo built with it.

The source is migrated from the frozen KPR candidate identified in
[`docs/MIGRATION_PROVENANCE.md`](docs/MIGRATION_PROVENANCE.md). The exact
supported calculation surface is recorded in
[`docs/PUBLIC_API.txt`](docs/PUBLIC_API.txt).

### 2. Validate locally

```bash
python3 tools/check_repo.py --root . --self-test
python3 tools/check_repo.py --root . \
  --output test-results/static-checks.json \
  --summary test-results/static-checks.md
python3 tools/check_kpr_contract.py --root . --self-test
python3 tools/check_kpr_contract.py --root . \
  --output test-results/kpr-contract.json \
  --summary test-results/kpr-contract.md
python3 tools/gen_fixtures.py --root . --check
python3 tools/check_release.py --root . --self-test \
  --summary test-results/release-self-test.md
```


Then import the applicable VBA components into a Windows Excel host, run
**Debug → Compile VBAProject**, and execute the migrated regression entry points.
The primary imported runner is `KPR_Tests_Run`; focused host, shape and
dynamic-array runners are also retained. Compilation and parity for the migrated
candidate were recorded in #17 ([evidence](evidence/migration-2026-09-24/session.txt));
record any new host run against its exact candidate.

### 3. Run the migrated date layer in Excel

Follow the [developer setup](docs/DEVELOPER_SETUP.md) and
[Excel evidence interface](docs/EXCEL_EVIDENCE.md) (with the
[migration regression protocol](docs/MIGRATION_REGRESSION.md) for parity runs).
Initialization is complete.
Current GitHub controls and limitations are recorded in
[setup verification](docs/SETUP_VERIFICATION.md); the completed migration is
recorded in [the migration completion record](docs/MIGRATION_COMPLETION.md), and
future work is sequenced in [the roadmap](docs/ROADMAP.md).

<a id="date-primitives"></a>

## 📅 Date Primitives (v0.0.4)

The supported surface is exactly these 22 functions of
[`KPR_DATES_DAYS`](src/modules/KPR_DATES_DAYS.bas), usable from a worksheet
cell and from VBA. [`docs/DATE_LAYER_CONTRACT.md`](docs/DATE_LAYER_CONTRACT.md)
is the normative behavior and [`docs/PUBLIC_API.txt`](docs/PUBLIC_API.txt) the
checked list of names and signatures.

| Group | Functions |
| --- | --- |
| Period boundaries | `KPR_Dates_BeginOfMonth`, `KPR_Dates_EndOfMonth`, `KPR_Dates_BeginOfQuarter`, `KPR_Dates_EndOfQuarter`, `KPR_Dates_BeginOfYear`, `KPR_Dates_EndOfYear` |
| Predicates and counts | `KPR_Dates_IsMonthEnd`, `KPR_Dates_IsQuarterEnd`, `KPR_Dates_IsYearEnd`, `KPR_Dates_IsLeapYear`, `KPR_Dates_DaysInMonth`, `KPR_Dates_DaysInYear`, `KPR_Dates_DayOfWeek` |
| Arithmetic | `KPR_Dates_AddDays`, `KPR_Dates_AddWeeks`, `KPR_Dates_AddMonths`, `KPR_Dates_AddYears` |
| Weekday locators | `KPR_Dates_NthWeekdayOfMonth`, `KPR_Dates_LastWeekdayOfMonth` |
| Pillars | `KPR_Dates_PillarFromDates`, `KPR_Dates_DateFromPillar` |
| Diagnostic | `KPR_Dates_HostDateSystem` (the one volatile function) |

The rules every function follows (contract sections 3 to 6 and 9):

- **One surface for scalars and arrays.** A scalar call returns a scalar; a
  Range or array argument returns a result of the same shape, and a scalar
  argument is broadcast to it. There are no separate array or spill
  functions. Multi-cell results are claimed only on dynamic-array Excel; no
  Ctrl+Shift+Enter behavior is claimed.
- **Supported dates:** 1900-03-01 through 9999-12-31. A date can be an Excel
  date value or text in exactly the ISO form `YYYY-MM-DD`; locale-formatted or
  numeric-looking text is refused, never guessed. A time part of a date value
  is ignored.
- **Optional arguments** (`Opt_...`) may be omitted or left blank to take their
  default, and must be a single value: a multi-cell control is a call-level
  `#VALUE!`.
- **Capacity:** a result of up to 100,000 elements; a larger one is a
  call-level `#NUM!`, decided before any cell is read.
- **Errors are native Excel errors:** `#VALUE!` for input that cannot be
  interpreted, `#NUM!` for a valid value outside the supported domain, a
  library `#N/A` when the calling workbook's date system cannot be used, and
  any incoming error passed through unchanged at its position.
- **Date system:** a worksheet call uses the calling workbook's own date
  system. A 1904 workbook gets a call-level `#N/A` from every value-taking
  function, never a value shifted by 1,462 days; `KPR_Dates_HostDateSystem()`
  reports 1900 or 1904. Direct VBA, `Application.Run` and the Immediate window
  use the 1900 serial contract; the active workbook is never consulted.
- **Calendar days only.** Business calendars, weekend masks and holidays
  (`KPR_Cal_*`) are v0.0.5 scope, business-day arithmetic and roll conventions
  v0.0.6 scope ([roadmap](docs/ROADMAP.md)).

To try it, build the workbook with the one-click Excel check
([`INSTALLATION.md`](INSTALLATION.md#one-click-excel-check)): it opens with the
**K-PRICING** Ribbon tab, whose **Functions** menu inserts any of the 22
functions and whose **Demo sheets** menu builds the **KPR Dates Demo** sheet,
comparing each function with Excel's own formulas
([`examples/README.md`](examples/README.md)). Registration, Ribbon, toolbar,
demo and test entry points are unsupported infrastructure, classified in
[`docs/API_INVENTORY.txt`](docs/API_INVENTORY.txt).

<a id="application-profile"></a>

## 🧭 Application profile

K-PRICING is an **application**: an end-to-end Excel workbook or add-in
solution that will own its deployment and lifecycle. The profile is fixed in
[`.github/repository-profile.json`](.github/repository-profile.json). Profile
rules add to the repository baseline and never weaken source integrity,
security, action pinning or release provenance.

Today the source is the migrated date layer, its test infrastructure, the
Ribbon (`src/ribbon/`) and the macro-free workbook template
(`src/workbook/KPR_Template.xlsx`). There are no class, form or
workbook-document modules, and no tracked workbook or add-in package: the
workbook is built from the template by `tools/Check-KPR.ps1`.

Profile exception: Office packages are never tracked, except
`src/workbook/KPR_Template.xlsx`, re-included by its exact path in
`.gitignore` and allowed by `allowed_office_binary_globs`. It is generated
by `tools/workbook_template.py` from tracked source, holds no VBA and no
document properties, and CI fails when it differs from its source. Owner: the
repository owner. Removing it means building the workbook from an empty one
again in `tools/Check-KPR.ps1`.

### Application commitments

As application components are added, document startup, shutdown,
configuration, data boundaries, deployment, upgrade and rollback. A functional
release must provide the startup, shutdown, upgrade, recovery, packaging and
end-to-end smoke evidence required by the release policy, and a distributable
workbook or add-in also needs provenance and post-package smoke evidence; see
[`RELEASING.md`](RELEASING.md) and
[`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md).

<a id="repository-shape"></a>

## 🏗️ Repository shape

| Path | Responsibility |
| --- | --- |
| [`src/`](src/) | Authoritative exported production VBA |
| [`tests/`](tests/) | Regression source, stable fixtures and evidence instructions |
| [`examples/`](examples/) | Minimal examples using supported APIs |
| [`assets/`](assets/) | Versioned documentation/visual assets |
| [`docs/`](docs/) | Durable specialized contracts |
| [`tools/`](tools/) | Deterministic validation and release tooling |
| [`.github/`](.github/) | Workflows, intake forms and declarative repository policy |

Directory ownership, VBA façade/core/UI/test separation, export format and
legitimate profile-specific alternatives are authoritative in
[`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md).

<a id="requirements"></a>

## 🖥️ Requirements

The initial target is Microsoft 365 Excel desktop on Windows, with 32-bit and
64-bit Office evaluated separately. The accepted v0.0.1 neutral-starter run is
historical setup evidence only; it does not certify the migrated KPR date layer.
The migrated date layer compiled and matched the frozen source on one Windows
64-bit Microsoft 365 Excel host (build 16.0.20326.20072, Italian regional format)
in issue #17. The most recent full regression ran twice on Windows 11 Enterprise
with 64-bit Microsoft 365 Excel (build 16.0.20430.20180) at commit `ba250a7`:
23 suites, 9,372 assertions, no failures (#47, #48). That is development
evidence; the candidate is certified only by its own run in #52. 32-bit Office
and other builds and locales remain untested.
Mac, Excel for the web, older builds and a packaged deployment are not claimed
as supported.

Do not infer compatibility from source inspection or from one successful host.
Installation, import, upgrade and removal procedures are authoritative in
[`INSTALLATION.md`](INSTALLATION.md).

<a id="validation"></a>

## ✅ Validation

The hosted `Repository integrity` workflow checks source/repository facts and
fails closed when a required validator or evidence report does not complete. It
does **not** compile VBA or execute Excel.


Generated inputs are regenerated, never edited by hand:
`python3 tools/gen_fixtures.py --root . --write` rewrites the regression
fixtures ([`tests/fixtures/README.md`](tests/fixtures/README.md)),
`python3 tools/workbook_template.py --write` the workbook template, and
`python3 tools/milestone_register.py --write` the milestone register. The
static checks fail when the fixtures or the template differ from their source;
the register is compared with the live issues by the separate
`Milestone register` workflow and must be `CURRENT` before tagging. The demo sheet is not tracked: it is
built in Excel each time ([`examples/README.md`](examples/README.md)).

For a release candidate, use [`RELEASING.md`](RELEASING.md); certification
runs `KPR_Test_RunAll` twice on the exact candidate in Windows Excel and
validates the evidence ([`docs/KPR_TEST_EVIDENCE.md`](docs/KPR_TEST_EVIDENCE.md)). SemVer/changelog
semantics and release-evidence schemas are maintained separately in
[`RELEASE_SEMANTICS.md`](docs/RELEASE_SEMANTICS.md) and
[`RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md).

## 🛡️ Engineering boundaries

- Treat exported source as authoritative; never use an opaque workbook as the
  only record of a code change.
- Restore only Excel/Windows state the component successfully acquired or
  changed and still owns.
- Define invalid-input, error, cleanup and partial-success behavior explicitly.
- Use independent expected results for numerical or behavioral verification.
- Record skips and untested environments as limitations, not passes.
- Keep stronger project-specific numerical, UI, lifecycle, performance or
  packaging gates additive to the generic repository baseline.

Detailed source ownership belongs to
[`REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md); contributor workflow
belongs to [`CONTRIBUTING.md`](CONTRIBUTING.md).

<a id="documentation"></a>

## 📚 Documentation

The canonical authority map is [`docs/README.md`](docs/README.md). Start with the
document that owns your task:

| Task | Authority |
| --- | --- |
| Set up a developer checkout | [`docs/DEVELOPER_SETUP.md`](docs/DEVELOPER_SETUP.md) |
| Verify repository setup | [`docs/SETUP_VERIFICATION.md`](docs/SETUP_VERIFICATION.md) |
| Review planned milestones and sequencing | [`docs/ROADMAP.md`](docs/ROADMAP.md) |
| Review the completed migration | [`docs/MIGRATION_COMPLETION.md`](docs/MIGRATION_COMPLETION.md) |
| Understand source/repository structure | [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) |
| Install, upgrade or remove | [`INSTALLATION.md`](INSTALLATION.md) |
| Contribute or review a change | [`CONTRIBUTING.md`](CONTRIBUTING.md) |
| Report a vulnerability | [`SECURITY.md`](SECURITY.md) |
| Prepare a release | [`RELEASING.md`](RELEASING.md) |
| Define release evidence | [`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md) |
| Check release semantics | [`docs/RELEASE_SEMANTICS.md`](docs/RELEASE_SEMANTICS.md) |
| Provision repository settings | [`docs/POST_CREATION_CHECKLIST.md`](docs/POST_CREATION_CHECKLIST.md) |

## ⚠️ Known limitations

These are the current limitations. The
[initialization record](.github/initialization.json) preserves the original
2026-09-23 setup inputs and is provenance, not a live copy of this section.

- Windows Excel evidence exists for 64-bit Microsoft 365 Excel only; 32-bit Office, other builds and other locales, Mac and Excel for the web are untested and not claimed.
- v0.0.4 is a candidate until its exact-source certification and the `v0.0.4` tag and GitHub pre-release are complete (#52). It is distributed as tagged source; the workbook is built from it by `tools/Check-KPR.ps1`, and no workbook or add-in package is tracked.
- Only the 22 `KPR_Dates_*` date primitives are supported. Calendars and holidays are planned for v0.0.5, business-day arithmetic and roll conventions for v0.0.6; broader pricing capabilities remain unscheduled ([roadmap](docs/ROADMAP.md)).
- The historical neutral-starter run certifies repository setup only.

The general evidence boundaries still apply: static inspection is not Excel
execution, and one tested environment does not certify another.

## 🔐 Security and conduct

Never include credentials, client/personal data, proprietary workbooks or other
restricted material. Suspected vulnerabilities must be reported privately using
[`SECURITY.md`](SECURITY.md). Participation is governed by
[`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md).

## 📄 License and maintainer

Copyright (c) 2026 Daniele Penza. Distributed under the
[Mozilla Public License 2.0](LICENSE) (MPL-2.0). You may combine K-PRICING
with your own code in any workbook or product, including a closed one, and
your own files keep your terms. Whenever you distribute K-PRICING's files,
modified or not and whether as source or inside a workbook or add-in, you must
make their source available under MPL 2.0 and keep their notices. This summary
does not replace the licence text. Commits before the
licence change remain available under the MIT License they were published
with. Maintained by **Daniele Penza**.

---

**Project principle:** keep source reviewable, contracts explicit, evidence
bounded to what was actually tested, and detailed policy in one authoritative
location.
