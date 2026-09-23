<div align="center">

# 📦 Installation and Upgrade Guide

### Install, validate, upgrade, troubleshoot, and remove {{PROJECT_NAME}}

[![Deployment](https://img.shields.io/badge/Deployment-Source--first-0969da?style=flat-square)](#deployment-model)
[![Validation](https://img.shields.io/badge/Validation-Required-d97706?style=flat-square)](#validation)
[![Security](https://img.shields.io/badge/Security-Private_policy-d73a49?style=flat-square)](SECURITY.md)
[![Version](https://img.shields.io/badge/Version-VERSION_file-6f42c1?style=flat-square)](VERSION)

<br>

**One identifiable source version · Clean import · Compile · Validate · Preserve caller state**

</div>

---

This document is authoritative for **installation, import, upgrade, recovery and
removal**. Source layout is owned by
[`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md), vulnerability
handling by [`SECURITY.md`](SECURITY.md), and release publication/provenance by
[`RELEASING.md`](RELEASING.md).

> [!IMPORTANT]
> VBA executes with the user's Office permissions. Review the exact source or use
> a trusted release, follow organizational macro policy, and never enable macros
> in an untrusted workbook.

## 🧭 Support baseline

Before a release is installable, document the actual supported Excel/Office
versions, Windows scope, Office bitness, references/dependencies and deployment
model. Compatibility claims apply only to environments actually certified for
that release.

| Item | Source of truth |
| --- | --- |
| Current version | [`VERSION`](VERSION) |
| User-visible changes | [`CHANGELOG.md`](CHANGELOG.md) |
| Source/component ownership | [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) |
| Security handling | [`SECURITY.md`](SECURITY.md) |
| Published release evidence | [`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md) |

This repository is currently rendered for the **{{PROFILE_NAME}}** profile.

### Complete repository checkout

Use a **Git clone** when you intend to initialize, validate, contribute to, or
release the project. The repository's `.gitattributes` deliberately marks
repository-plumbing paths such as `.github/`, `.gitignore`, `.editorconfig`, and
`.gitattributes` as `export-ignore`; GitHub source archives are generated with
`git archive`, so **Code → Download ZIP is not a complete maintainer checkout**.
In particular, a ZIP snapshot can omit workflows and
`.github/repository-profile.json`, and repository gates that require those files
will fail by design.

```text
git clone https://github.com/{{REPOSITORY_PATH}}.git
cd <repository-directory>
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

## 📂 Neutral starter import order

For the baseline starter:

1. `src/core/ProjectCore.bas`
2. `src/modules/ProjectFacade.bas`
3. `tests/modules/ProjectTests.bas` — development/regression only
4. `examples/modules/ProjectExample.bas` — optional consumer example

A generated project may replace or extend this layout. The authoritative source
and component-role rules are in
[`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md).

> [!CAUTION]
> A `.frm` and adjacent `.frx` are one logical UserForm component. Import the
> `.frm`; never import or edit the `.frx` as text.

## 🚀 Fresh installation

1. Back up the destination workbook/add-in and user data.
2. Obtain one exact supported source version or verified release.
3. Import every required component into the intended VBA project.
4. Configure only documented references, callbacks and host integrations.
5. Run **Debug → Compile VBAProject**.
6. Save in the required macro-capable format.
7. Close and reopen the host when the project requires clean-session validation.
8. Run the applicable smoke/regression checks below.

Do not paste exported source into arbitrarily named modules when a governed VBE
export is available. Component identity and form resources are part of a
reproducible installation.

<a id="validation"></a>

## ✅ Validation

A successful import is not certification. Validate the exact installed source in
a supported host.

For the neutral baseline, run `ProjectTests.RunProjectTests`. A passing run ends
with:

```text
RESULT=PASS; completeness=COMPLETE; cases=4; assertions=6; failures=0; cleanup=PASS
```

Run `ProjectExample.RunProjectExample` separately for the minimal consumer smoke.
A generated project must replace these baseline expectations when it changes the
starter contract.

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

- [`README.md`](README.md) — overview, profiles and first-use navigation
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
