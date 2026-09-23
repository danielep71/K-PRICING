<div align="center">

# ⚡ K-PRICING

### Financial analytics and instrument pricing for Excel/VBA

**K-PRICING is an Excel/VBA project for practitioners developing transparent financial analytics and instrument pricing.**

<br>

[![Excel VBA](https://img.shields.io/badge/Excel_VBA-source--first-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](#requirements)
[![Profile](https://img.shields.io/badge/Profile-see_contract-6f42c1?style=for-the-badge)](#supported-profiles)
[![Version](https://img.shields.io/badge/Version-VERSION_file-0969da?style=for-the-badge)](VERSION)
[![License](https://img.shields.io/badge/License-MIT-2ea44f?style=for-the-badge)](LICENSE)

<br>

[![Static checks](https://github.com/danielep71/K-PRICING/actions/workflows/static-checks.yml/badge.svg?branch=main)](https://github.com/danielep71/K-PRICING/actions/workflows/static-checks.yml)

<br>

**Source-first VBA · Explicit contracts · Deterministic evidence**

[Quick start](#quick-start)
&nbsp;·&nbsp;
[Profiles](#supported-profiles)
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



## ✨ What this project is

**Status: initialized private application scaffold, version `0.0.0`.**
The pricing implementation has not yet been migrated and no Excel runtime
certification or installable product is claimed. See the
[initialization status](docs/INITIALIZATION_STATUS.md) for provenance and setup boundaries.

The product and repository name is **K-PRICING**. The planned migration from
[`KPR`](https://github.com/danielep71/KPR) retains the **`KPR_`** VBA namespace
to preserve existing module and function names.

K-PRICING is a source-first Excel/VBA project. Exported VBA, tests,
versioned policy and documentation are the reviewable source of truth; Office
packages are generated or release artifacts unless an exact path is explicitly
governed.

The template separates five evidence layers that must not be conflated:
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

The neutral starter contains:

- [`ProjectCore`](src/core/ProjectCore.bas) — internal implementation;
- [`ProjectFacade`](src/modules/ProjectFacade.bas) — supported public façade;
- [`ProjectTests`](tests/modules/ProjectTests.bas) — regression harness; and
- [`ProjectExample`](examples/modules/ProjectExample.bas) — minimal consumer example.

The starter proves the repository shape; it is not project-specific business
logic. Rename or replace it only as one coherent change across source, tests,
examples, repository policy and the [public API manifest](docs/PUBLIC_API.txt).

### 2. Validate locally

```bash
python3 tools/check_repo.py --root . --self-test
python3 tools/check_repo.py --root . \
  --output test-results/static-checks.json \
  --summary test-results/static-checks.md
python3 tools/check_release.py --root . --self-test \
  --summary test-results/release-self-test.md
```


Then import the applicable VBA components into a supported Excel host, run
**Debug → Compile VBAProject**, and execute the documented regression or smoke
entry point. The neutral baseline is `ProjectTests.RunProjectTests`.

### 3. Provision GitHub settings

Template generation copies files, not labels, rulesets, topics, merge settings
or security settings. Apply and verify the authoritative
[`POST_CREATION_CHECKLIST.md`](docs/POST_CREATION_CHECKLIST.md) after generation.

<a id="supported-profiles"></a>

## 🧭 Supported profiles

Choose one profile. Specialist controls may be added, but a profile never
weakens source integrity, security, action pinning or release provenance.

| Profile | Use when | Additional evidence |
| --- | --- | --- |
| `library` | Reusable callable VBA with no owned end-user shell | Public API, caller contract and focused regression |
| `ui-component` | An embeddable bounded interactive surface | UI state, cleanup, recovery, DPI/accessibility and lifecycle evidence |
| `application` | An end-to-end workbook or add-in solution | Startup, shutdown, upgrade, recovery, packaging and smoke evidence |

### Selected profile contract

This repository is an **application**: an end-to-end workbook or add-in solution that owns deployment and lifecycle.
Its source contract covers modules, classes and the applicable workbook, form, Ribbon or host-lifecycle exports.
At minimum, retain startup, shutdown, upgrade, recovery, packaging and end-to-end smoke evidence.



### Application commitments

Document startup, shutdown, configuration, data boundaries, deployment,
upgrade and rollback. A distributable package requires provenance and
post-package smoke evidence.

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

A generated project must state its actual support matrix before release:

- supported Excel/Microsoft 365 versions/builds;
- Windows/operating-system scope;
- Office bitness and VBA generation;
- required references, native APIs or external dependencies; and
- supported deployment model.

Do not infer compatibility from source inspection or from one successful host.
Installation, import, upgrade and removal procedures are authoritative in
[`INSTALLATION.md`](INSTALLATION.md).

<a id="validation"></a>

## ✅ Validation

The hosted `Repository integrity` workflow checks source/repository facts and
fails closed when a required validator or evidence report does not complete. It
does **not** compile VBA or execute Excel.


For a release candidate, use [`RELEASING.md`](RELEASING.md). SemVer/changelog
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
| Initialize a generated repository | [`docs/INITIALIZATION.md`](docs/INITIALIZATION.md) |
| Understand source/repository structure | [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) |
| Install, upgrade or remove | [`INSTALLATION.md`](INSTALLATION.md) |
| Contribute or review a change | [`CONTRIBUTING.md`](CONTRIBUTING.md) |
| Report a vulnerability | [`SECURITY.md`](SECURITY.md) |
| Prepare a release | [`RELEASING.md`](RELEASING.md) |
| Define release evidence | [`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md) |
| Check release semantics | [`docs/RELEASE_SEMANTICS.md`](docs/RELEASE_SEMANTICS.md) |
| Provision repository settings | [`docs/POST_CREATION_CHECKLIST.md`](docs/POST_CREATION_CHECKLIST.md) |

## ⚠️ Known limitations

Do not hide an unresolved support or assurance boundary.

- Repository initialization only: the neutral starter is not a pricing implementation; migration of existing pricing work is pending.
- No supported workbook or add-in, functional release, or exact-source Excel runtime certification is available.

If no project-specific limitation is rendered, the general evidence boundaries
above still apply: static inspection is not Excel execution, and one tested
environment does not certify another.

## 🔐 Security and conduct

Never include credentials, client/personal data, proprietary workbooks or other
restricted material. Suspected vulnerabilities must be reported privately using
[`SECURITY.md`](SECURITY.md). Participation is governed by
[`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md).

## 📄 License and maintainer

Distributed under the [MIT License](LICENSE). Maintained by
**Daniele Penza**.

---

**Project principle:** keep source reviewable, contracts explicit, evidence
bounded to what was actually tested, and detailed policy in one authoritative
location.
