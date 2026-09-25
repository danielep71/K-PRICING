<div align="center">

# 📦 Installation and Upgrade Guide

### Install, validate, upgrade, troubleshoot, and remove K-PRICING

[![Deployment](https://img.shields.io/badge/Deployment-Source--first-0969da?style=flat-square)](#deployment-model)
[![Validation](https://img.shields.io/badge/Validation-Required-d97706?style=flat-square)](#validation)
[![Security](https://img.shields.io/badge/Security-Private_reporting-d73a49?style=flat-square)](SECURITY.md)
[![Version](https://img.shields.io/badge/Version-VERSION_file-6f42c1?style=flat-square)](VERSION)

<br>

**One identifiable source version · Clean import · Compile · Validate · Preserve caller state**

</div>

---

This document is authoritative for **installation, import, export, upgrade,
recovery and removal**. Source layout is owned by
[`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md), vulnerability
handling by [`SECURITY.md`](SECURITY.md), and release publication/provenance by
[`RELEASING.md`](RELEASING.md).

> [!IMPORTANT]
> VBA executes with the user's Office permissions. Review the exact source or use
> a trusted release, follow organizational macro policy, and never enable macros
> in an untrusted workbook.

## 🧭 Support baseline

The current installation path is manual source import into a disposable `.xlsm`
for development. There is no supported workbook/add-in package or functional
release. The initial target and evidence status are:

| Target | Status |
| --- | --- |
| Microsoft 365 Excel desktop on Windows, 64-bit Office | Migrated date layer compiled and matched the frozen source on one host (Version 2608, Build 16.0.20326.20072, Italian regional format) in #17; other builds and locales untested |
| Microsoft 365 Excel desktop on Windows, 32-bit Office | Intended; untested separately |
| Mac, Excel for the web, older desktop builds | No support claim |
| References | Built-in VBA/Excel only; no additional external dependency for the migrated date layer |

Use [developer setup](docs/DEVELOPER_SETUP.md) for a checkout and
[the Excel evidence interface](docs/EXCEL_EVIDENCE.md) for the host procedure
(with [the migration regression protocol](docs/MIGRATION_REGRESSION.md) for parity runs).
Compatibility claims apply only to environments actually validated.

| Item | Source of truth |
| --- | --- |
| Current version | [`VERSION`](VERSION) |
| User-visible changes | [`CHANGELOG.md`](CHANGELOG.md) |
| Source/component ownership | [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) |
| Security handling | [`SECURITY.md`](SECURITY.md) |
| Published release evidence | [`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md) |

K-PRICING uses the **application** profile (see the [README](README.md#application-profile)).

### Complete repository checkout

Use a **Git clone** when you intend to validate, contribute to, or release the
project. The repository's `.gitattributes` deliberately marks
repository-plumbing paths such as `.github/`, `.gitignore`, `.editorconfig`, and
`.gitattributes` as `export-ignore`; GitHub source archives are generated with
`git archive`, so **Code → Download ZIP is not a complete maintainer checkout**.
In particular, a ZIP snapshot can omit workflows and
`.github/repository-profile.json`, and repository gates that require those files
will fail by design.

```text
git clone https://github.com/danielep71/K-PRICING.git
cd K-PRICING
```

Use a GitHub ZIP/tar source archive only when you intentionally need the
consumable exported source subset and do not intend to run repository governance
or initialization tooling. Do not diagnose missing workflow/profile files from
such an archive as repository defects.

<a id="deployment-model"></a>

## 🎯 Deployment model

Use one documented deployment model per supported installation path:

| Model | Use when | Identity boundary |
| --- | --- | --- |
| Embedded source | Components travel inside a workbook/add-in | Destination project contains the reviewed exports |
| Tagged source | Consumer imports/builds the project | Tag/commit and exported files define identity |
| Published binary | Project ships a workbook/add-in asset | Tag binding, hash and packaged smoke evidence are required |
| Development source | Contribution/testing work | Not a supported release unless explicitly stated |

Never mix components from different tags, commits, local exports or release
assets.

## 📂 Migrated date-layer import order

Import the migrated production dependency set in this order:

1. `src/core/KPR_Core_Err.bas`
2. `src/core/KPR_Core_Parse.bas`
3. `src/core/KPR_Core_Dates.bas`
4. `src/core/KPR_Core_Array.bas`
5. `src/modules/KPR_DATES_DAYS.bas`
6. `tests/modules/KPR_Test_Fixtures_Generated.bas` — development/regression
   only; generated fixture data with no dependencies
7. `tests/modules/KPR_REGRESSION_TESTS.bas` — development/regression only
8. `examples/modules/KPR_DateExample.bas` — optional consumer example

The four core modules are project-internal; `KPR_DATES_DAYS` owns the supported
22-function calculation API. Source provenance and path adaptation are recorded
in [`docs/MIGRATION_PROVENANCE.md`](docs/MIGRATION_PROVENANCE.md).

> [!CAUTION]
> A `.frm` and adjacent `.frx` are one logical UserForm component. Import the
> `.frm`; never import or edit the `.frx` as text.

## 🚀 Fresh installation

1. Back up the destination workbook/add-in and user data.
2. Obtain one exact supported source version or verified release.
3. Remove any same-named standard module, class module or UserForm from the
   intended VBA project after preserving local changes, then use
   **File → Import File** for each of those required components. Import does
   not replace an existing component and can silently suffix its name.
   Host-bound document modules under `src/workbook/` follow the separate
   procedure below; never import them.
4. Configure only documented references, callbacks and host integrations.
5. Run **Debug → Compile VBAProject**.
6. Save in the required macro-capable format.
7. Close and reopen the host when the project requires clean-session validation.
8. Run the applicable smoke/regression checks below.

Do not paste exported source into arbitrarily named modules when a governed VBE
export is available. Component identity and form resources are part of a
reproducible installation.

### Workbook and worksheet document modules

`ThisWorkbook` and worksheet modules are bound to host objects. **File → Import
File** does not restore that binding; it creates an ordinary class module
instead. For each component under `src/workbook/`:

1. Identify the existing host object it belongs to (`ThisWorkbook` or the named
   worksheet) and preserve any local code in that module.
2. Replace the code in that existing module with the exported procedure text.
   Omit the export header (`VERSION`, `BEGIN`…`END` and `Attribute` lines); the
   host object already owns those properties.
3. Do not rename, delete or re-create the host object to force a match.

The migrated date layer has no document modules. Any future component that adds
one must name its host object here and follow the
[repository structure](docs/REPOSITORY_STRUCTURE.md) and
[Excel evidence](docs/EXCEL_EVIDENCE.md) rules for document modules.

## Export from Excel

1. Select the owned component in the VBE Project Explorer and choose
   **File → Export File**.
2. Use the role directory defined by
   [repository structure](docs/REPOSITORY_STRUCTURE.md), preserving the component
   name and matching filename. Keep a UserForm's `.frm` and `.frx` together.
3. Review the complete diff against the
   [VBE format contract](docs/VBE_EXPORT.md), including ASCII source, line endings,
   hidden attributes and identity, before committing.
4. For release certification, import the exact candidate into a clean project,
   export without edits and compare normalized source. Retain the candidate SHA
   and comparison under [release evidence](docs/RELEASE_EVIDENCE.md).

<a id="validation"></a>

## ✅ Validation

A successful import is not certification. Validate the exact installed source in
a supported host.

For the migrated date layer, compile the complete project and run
`KPR_Tests_Run`. The imported harness also exposes `KPR_Tests_RunSuite`,
`KPR_Tests_RunAll`, `KPR_Tests_RunHost`, `KPR_Tests_RunShape`, and
`KPR_Tests_RunArray` for focused execution. `KPR_Tests_RunEvidence` emits the
machine-readable host-evidence log, and `KPR_Tests_RunMigrationEvidence` is
retained for the historical #17 parity run.

Run `KPR_DateExample.RunDateExample` separately for the minimal direct-VBA
consumer smoke. Historical KPR results are source evidence only. Destination
parity for the migrated candidate is recorded in
[`evidence/migration-2026-09-24`](evidence/migration-2026-09-24/session.txt) (#17);
record counts, failures/skips, environment and cleanup for any new run against
its exact candidate and owning issue.

Record at least:

```text
Source tag / full commit SHA:
VERSION:
Files imported:
Excel / Office version and build:
Office bitness:
Operating system:
Compile:
Smoke / regression:
Specialist checks:
Cleanup:
Skipped or unverified:
```

A skipped, incomplete or cleanup-failed run is not a pass. Repository/static
checks cannot substitute for Excel-host execution.

## ⬆️ Upgrade

Before upgrading:

1. read the complete version-to-version changelog;
2. back up the host and any supported user configuration;
3. identify the complete production component/package set for the target release;
4. review migration, compatibility and known limitations; and
5. stop/clean active project state where applicable.

Replace one coherent release, compile again and repeat the full installation
validation. Do not infer backward compatibility merely from successful VBA
compilation.

### Local modifications

Treat a locally modified copy as a fork. Diff it against the old and new exported
source, reapply modifications deliberately and retest them. Do not overwrite a
local fork and assume behavior survived.

## 🧯 Troubleshooting

| Symptom | Check first |
| --- | --- |
| Compile error / missing procedure | Confirm all required components come from one version and dependencies are present. |
| Ambiguous name | Remove duplicate or legacy components. |
| Form controls missing/corrupt | Re-import the exact `.frm` with its adjacent `.frx`. |
| Workbook-dependent behavior | Check explicit workbook/worksheet ownership, references, locale and date-system assumptions. |
| 32/64-bit failure | Confirm supported bitness and applicable conditional declarations. |
| Excel state remains altered | Use the documented cleanup/recovery path; do not blindly overwrite global state. |
| Security warning | Verify source origin/release provenance and organizational macro settings. |
| Result differs from reference | Confirm exact version, inputs, configuration, environment and independent acceptance rule. |

If recovery is uncertain, preserve user data, close Excel, reopen a clean session
and reproduce with a minimal sanitized workbook before changing code.

Suspected security problems follow [`SECURITY.md`](SECURITY.md), not public
troubleshooting channels.

## 🗑️ Removal

1. Run any project-owned shutdown/cleanup procedure.
2. Remove the production components and optional integrations the project owns.
3. Compile the remaining VBA project.
4. Close/reopen and verify the host no longer depends on removed project state.

Removing modules does not automatically remove formulas, Ribbon XML, names,
links, connections, add-in registration, callbacks or other host integrations.
Remove only state the project owns and document anything intentionally retained.

## 🔐 Security boundary

Installation guidance does not redefine the security policy. Obtain source and
artifacts only from official channels, use synthetic/non-sensitive validation
data, and follow [`SECURITY.md`](SECURITY.md) for trust boundaries, secrets,
private vulnerability reporting and safe use.

## 📚 Related authorities

- [`README.md`](README.md) — overview, application profile and first-use navigation
- [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) — source/layout contract
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — change/review workflow
- [`RELEASING.md`](RELEASING.md) — maintainer release sequence
- [`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md) — release evidence schema
- [`SECURITY.md`](SECURITY.md) — vulnerability/security policy
- [`docs/README.md`](docs/README.md) — complete documentation authority map

---

**Installation principle:** install one identifiable source version, compile it,
exercise its real host behavior, and retain evidence of what was and was not
validated.

The accepted neutral-starter run is recorded in
[SETUP_COMPLETION.md](docs/SETUP_COMPLETION.md), with its exact tested SHA and limits.
