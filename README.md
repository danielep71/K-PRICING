<div align="center">

# ⚡ K-PRICING

### Financial analytics and instrument pricing for Excel/VBA

**Source-first VBA · Explicit contracts · Evidence tied to one commit · Governed releases**

<br>

[![Excel VBA](https://img.shields.io/badge/Excel_VBA-exported_source-217346?style=for-the-badge)](#-vba-development)
[![Windows](https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge)](INSTALLATION.md)
[![Python](https://img.shields.io/badge/Tooling-Python_3.11+-3776AB?style=for-the-badge&logo=python&logoColor=white)](#-quality-gates)
[![GitHub Actions](https://img.shields.io/badge/CI-GitHub_Actions-2088FF?style=for-the-badge&logo=githubactions&logoColor=white)](#-quality-gates)
[![License](https://img.shields.io/badge/License-MPL_2.0-2ea44f?style=for-the-badge)](#-license)
<br>[![Clones, including CI](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/danielep71/K-PRICING/traffic-history/data/badges/clones.json&style=flat-square)](docs/repo_governance/TRAFFIC.md#badges)
[![Unique visitors, last 14 days](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/danielep71/K-PRICING/traffic-history/data/badges/visitors.json&style=flat-square)](docs/repo_governance/TRAFFIC.md#badges)
[![Daily unique visitors since tracking began](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/danielep71/K-PRICING/traffic-history/data/badges/visits.json&style=flat-square)](docs/repo_governance/TRAFFIC.md#badges)

<br>

**K-PRICING is an Excel/VBA project for practitioners who want transparent financial analytics and instrument pricing, reviewable line by line**

[What it is](#-what-k-pricing-is)
&nbsp;·&nbsp;
[What it provides](#-what-the-repository-provides)
&nbsp;·&nbsp;
[Find your way](#-find-your-way)
&nbsp;·&nbsp;
[Lifecycle](#-from-source-to-release)
&nbsp;·&nbsp;
[Quality gates](#-quality-gates)
&nbsp;·&nbsp;
[Documentation](#-documentation)
&nbsp;·&nbsp;
[Status](#-status)

---

</div>

## ✨ What K-PRICING is

K-PRICING is a source-first Excel/VBA project. Exported VBA, tests, contracts
and documentation are the reviewable source of truth; a workbook is built from
the source of one commit and is never the only record of a change. **KPR** is
the short form of the name, and the prefix of its VBA namespace.

I build it for practitioners who need to see how a figure is computed: every
function states its inputs, its errors and its limits, and every claim of
correctness names the commit and the Excel environment that produced it.

I am building it in layers, each released on its own evidence: the 22
`KPR_Dates_*` date primitives first, then calendars and business days, day-count
conventions and schedule generators ([plan](#-status)). Valuation and pricing
are not scheduled yet.

K-PRICING keeps five kinds of evidence apart:

| Evidence | What it shows |
| --- | --- |
| Repository integrity | The source, the documents and the workflows agree with the rules |
| VBA compilation | The source compiles in Excel |
| Regression execution | Every case and assertion ran and passed in Excel |
| Specialist assurance | An independent reference agrees with the results |
| Release certification | A tagged commit carries all of the above |

| Principle | Project contract |
| --- | --- |
| Source identity | Exported text source and an exact Git revision |
| Public surface | Documented and machine-checked |
| State ownership | Caller and host state is changed and restored deliberately |
| Failure behavior | Invalid input, errors and cleanup are part of the contract |
| Evidence | A claim names the exact candidate and environment tested |
| Portability | Supported Excel and Office environments are stated, not inferred |

## 🎯 What the repository provides

| Area | What the repository provides |
| --- | --- |
| Source control for Excel | Editing and Git policies for workbooks, exported VBA source and byte-preserved regulatory source copies |
| VBA quality | Static checks of exported VBA (layout, names, `#If` structure, jump targets, visibility, dependency direction, identifiers, `Err` copies, native declarations, calls) and a layout formatter, run on every change |
| Reproducible builds | A macro-enabled workbook built from the source of one commit, without *Trust access to the VBA project object model*, with the export bound to the workbook it came from |
| Testing and evidence | A VBA test harness with a certification entry point, numerical test cases and accuracy contracts, and an offline check that the evidence of an Excel run matches one commit |
| Methodology | Guides for the sources, the data contract, the calculation and workbook contracts and the validation plan, with a traceability table from each rule to its tests |
| Python quality gates | Ruff, strict mypy and actionlint, plus tests that hold the workflows to pinned actions, explicit permissions and the required check |
| Links | Relative links, anchors and HTTPS on every change; external links weekly |
| Issues and pull requests | Issue labels, issue forms, triage automation, a pull request template with a final-head review, and a workflow that keeps milestones on issues only |
| GitHub settings | A documented settings baseline (general settings, security features, rulesets, Actions settings) that a script previews, applies and verifies, writing a dated evidence record |
| Security | A security policy, a review to record before distribution, CodeQL scanning of the Python, JavaScript and workflow code, an OpenSSF Scorecard measure, and Dependabot updates for the pinned Actions and the hash-locked Python dependencies |
| Releases | A release process with a changelog, a certification record, a milestone closeout and a reproducible, attested source bundle |
| Traffic | Optional daily collection of clones and visitors, with the badges above |

## 📍 Find your way

| I want to… | Read |
| --- | --- |
| Move VBA source between the repository and Excel, or record an Excel run | [Installation](INSTALLATION.md) |
| Write VBA that passes the static checks | [VBA house style](docs/VBA_HOUSE_STYLE.md) |
| Know where a component belongs and what it may call | [Repository structure](docs/REPOSITORY_STRUCTURE.md) |
| Record the sources, the data and the calculation, and how they are tested | [Methodology](docs/methodology/README.md) |
| Propose a change and run the checks locally | [Contributing](CONTRIBUTING.md) |
| Label, triage or plan an issue | [Issue labels](docs/repo_governance/ISSUE_LABELS.md) |
| Check or reapply the GitHub settings | [Repository setup](docs/setup/REPOSITORY_SETUP.md) |
| Prepare, certify and publish a release | [Releasing](RELEASING.md) |

## 🔄 From source to release

Each step leaves something a later step, or a reviewer, can check:

```mermaid
flowchart LR
    A["Exported VBA source<br/>and contracts"] -->|pull request| B["Repository integrity<br/>static checks and tests"]
    B --> C["Build the workbook<br/>from one commit"]
    C --> D["Excel run<br/>harness certification"]
    D --> E["Evidence checked<br/>against the commit"]
    E --> F["Release<br/>certified and attested"]
```

| Step | Tool or record | What it proves |
| --- | --- | --- |
| Change | Pull request, with the [final-head review](docs/repo_governance/PULL_REQUESTS.md) | The reviewed commit is the one merged |
| Check | `python tools/check_repository.py` | Every **Repository integrity** check passes locally as in CI |
| Build | `tools/devbuild/build_workbook.py` | The workbook comes from the source of one commit, byte for byte reproducible |
| Run | `tools/devbuild/excel_run.py` | The harness certification ran in Excel, with a host record |
| Evidence | `python tools/check_excel_run.py` | The run's logs and results match the commit and the evidence policy |
| Release | [RELEASING.md](RELEASING.md), `python tools/release.py` | The tag, the changelog, the certification and the published assets agree |

## ✅ Quality gates

| Workflow | What it checks |
| --- | --- |
| **Repository integrity** | The required check: Python, VBA, workflow, link, changelog, test-case and governance checks, with the unit tests |
| **Minimum Python** | The unit tests on the minimum supported Python |
| **CodeQL** | Code scanning of the Python, JavaScript and workflow code |
| **OpenSSF Scorecard** | The supply-chain posture of the default branch; a run on any other branch skips the analysis |
| **Link check** | External links |
| **Sync issue labels** · **Detect issue-label drift** | The live labels against `.github/labels.json`, never deleting one in use |
| **Detect settings drift** | The live rulesets and the other settings a read-only token can read against `docs/setup/repository-settings.json`; the rest is listed as not checked |
| **Issue triage** | Every open issue against the [triage rules](docs/repo_governance/ISSUE_LABELS.md#triage-rules) |
| **Closing keywords** | The closing keywords a pull request would carry to the default branch: none into a release branch, none for an open issue in an integration |
| **Keep milestones issue-only** | Removes any milestone set on a pull request, and lists each removal |
| **Release** | Verifies the tag, builds a reproducible source bundle, attests it where the repository allows it and drafts the release |
| **Traffic history** | Optional: collects clones and visitors for the badges above |

Each workflow runs on the events of its `on:` block:

<!-- BEGIN generated: workflow-triggers -->
| Workflow | When it runs |
| --- | --- |
| [**Closing keywords**](.github/workflows/closing-keywords.yml) | Pull requests (opened, synchronize, reopened, edited) |
| [**CodeQL**](.github/workflows/codeql.yml) | Pushes to `main`; pull requests into `main`, `release/**` unless they change only `**/*.md`, `docs/**`; weekly on Wednesday at 04:29 UTC; manual runs |
| [**Issue triage**](.github/workflows/issue-triage.yml) | Issue events (opened, edited, reopened, closed, labeled, unlabeled, assigned, unassigned, milestoned, demilestoned); milestone events (edited, closed, opened, deleted); manual runs |
| [**Detect issue-label drift**](.github/workflows/labels-drift.yml) | Weekly on Tuesday at 05:23 UTC; manual runs |
| [**Sync issue labels**](.github/workflows/labels-sync.yml) | Pushes that change `.github/labels.json`, `.github/scripts/labels-sync.mjs`, `.github/workflows/labels-sync.yml`; manual runs |
| [**Link check**](.github/workflows/link-check.yml) | Weekly on Monday at 06:43 UTC; manual runs |
| [**Minimum Python**](.github/workflows/minimum-python.yml) | Pull requests that change `**/*.py`, `tests/**`, `tools/**`, `.github/scripts/**`, `pyproject.toml`, `.github/requirements.txt`, `.github/workflows/minimum-python.yml`; pushes to `main`; manual runs |
| [**Keep milestones issue-only**](.github/workflows/pr-milestones.yml) | Pushes that change `.github/workflows/pr-milestones.yml`; pull requests (opened, reopened, edited, milestoned), run from the base branch; weekly on Friday at 03:37 UTC; manual runs |
| [**Release**](.github/workflows/release.yml) | Pushes of the tags `v*.*.*`; manual runs |
| [**Repository integrity**](.github/workflows/repository-integrity.yml) | Every pull request; pushes to `main`, `release/**`; manual runs |
| [**OpenSSF Scorecard**](.github/workflows/scorecard.yml) | Branch protection rule changes; every push; weekly on Thursday at 05:41 UTC; manual runs |
| [**Detect settings drift**](.github/workflows/settings-drift.yml) | Daily at 04:41 UTC; manual runs |
| [**Traffic history**](.github/workflows/traffic-history.yml) | Daily at 06:17 UTC; manual runs |
<!-- END generated: workflow-triggers -->

## 🧩 VBA development

- The [repository structure](docs/REPOSITORY_STRUCTURE.md) says where each VBA
  component lives and which may call which.
- The [house style](docs/VBA_HOUSE_STYLE.md) covers naming, contracts, errors
  and cleanup.
- The [public API inventory](docs/PUBLIC_API.txt) lists the supported entry
  points, and every other public declaration under `src/`.
- The [installation guide](INSTALLATION.md) moves source between the
  repository and Excel, and covers upgrades, troubleshooting and removal.

## 📦 Repository structure

```text
.github/   workflows, issue forms, labels, Dependabot and the Excel evidence policy
docs/      methodology, release records, governance, setup, house style, licenses
tests/     VBA harness and example suite, test cases, accuracy contracts, Python tests
tools/     static checks, workbook build and Excel run, release and setup
src/       the workbook's VBA source and macro-free template (reserved: added with the date primitives)
```

The full layout and its rules are in
[`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md).

## 📚 Documentation

| Document | Covers |
| --- | --- |
| [`docs/methodology/`](docs/methodology/README.md) | Sources, data contract, calculation and workbook contracts, test cases, test harness, validation plan |
| [`INSTALLATION.md`](INSTALLATION.md) | Moving source between the repository and Excel, recording an Excel run, upgrades, troubleshooting and removal |
| [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) | Source layout, reserved paths, dependency direction |
| [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md) | VBA naming, contracts, errors, cleanup and static checks |
| [`docs/repo_governance/`](docs/repo_governance/README.md) | Settings, issue labels, pull requests, traffic, distribution |
| [`docs/setup/REPOSITORY_SETUP.md`](docs/setup/REPOSITORY_SETUP.md) | Setting up a new repository and what to keep, adapt or remove |
| [`RELEASING.md`](RELEASING.md) | How a version is prepared, certified and published, and the current status |
| [`docs/release/`](docs/release/CLOSEOUT.md) | The milestone [closeout](docs/release/CLOSEOUT.md) and the [release certification](docs/release/RELEASE_CERTIFICATION.md) records, and the [maintenance audit](docs/release/MAINTENANCE_AUDIT.md) done before each release |
| [`CONTRIBUTING.md`](CONTRIBUTING.md) | Workflow, checks, rules of evidence and pull requests |
| [`SECURITY.md`](SECURITY.md) | Private vulnerability reporting, runtime risk surface and data handling |
| [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md) | Participant behaviour and conduct reporting |
| [`CHANGELOG.md`](CHANGELOG.md) | Change history |

## 📌 Status

No version has been released: the repository is being set up (v0.1.0). The
release state is recorded in one place, the
[*Current status*](RELEASING.md#current-status) of `RELEASING.md`, which
**Repository integrity** checks against [`CHANGELOG.md`](CHANGELOG.md).
Development continues on the active release branch through task branches and
pull requests; only a release is versioned and tagged.

| Milestone | Scope |
| --- | --- |
| v0.1.0 | Repository setup: settings baseline, README and About fields |
| v0.2.0 | Date primitives: the 22 `KPR_Dates_*` worksheet functions on a core that does not depend on Excel |
| v0.3.0 | Holiday calendars and business days, with roll conventions |
| v0.4.0 | Day-count conventions |
| v0.5.0 | Schedule generators |

The [milestones](https://github.com/danielep71/K-PRICING/milestones) hold the
scope, the exit criteria and the due dates.

## 🖥️ Requirements and limitations

- The target is Microsoft 365 Excel desktop on Windows. 32-bit and 64-bit
  Office are evaluated separately.
- No Excel run has been recorded for K-PRICING yet, so no environment is
  claimed as supported. Mac, Excel for the web and older builds are not
  targets.
- Do not infer compatibility from source inspection or from one successful
  host: a claim names the exact commit and Excel environment it was run on.
- The first version is certified only by a run of the whole suite in Excel on
  its exact candidate, recorded in the release certification.

## 🔐 Data and security

No real portfolio data: never commit or attach real trades, positions,
counterparty data, or any client or employer material. VBA runs with the
user's Office permissions: review the source and follow your organisation's
macro policy. Report suspected vulnerabilities privately as described in
[`SECURITY.md`](SECURITY.md).

## 🧭 Contributing

To take part, read the [contributing guide](CONTRIBUTING.md) and the
[code of conduct](CODE_OF_CONDUCT.md).

## 👤 Author

**Daniele Penza** — [@danielep71](https://github.com/danielep71)

## 📄 License

Copyright (c) 2026 Daniele Penza. Distributed under the
[Mozilla Public License 2.0](LICENSE) (MPL-2.0). You may combine K-PRICING with
your own code in any workbook or product, including a closed one, and your own
files keep your terms. Whenever you distribute K-PRICING's files, modified or
not, as source or inside a workbook or add-in, you must make their source
available under MPL 2.0 and keep their notices. This summary does not replace
the licence text. The files listed in
[`docs/licenses/`](docs/licenses/THIRD_PARTY_NOTICES.txt) keep the license the
notice states for them, MPL-2.0 or MIT.
