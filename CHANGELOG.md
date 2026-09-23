<div align="center">

# 📜 Changelog

### Release history for {{PROJECT_NAME}}

[![Format](https://img.shields.io/badge/Format-Keep_a_Changelog-0969da?style=flat-square)](https://keepachangelog.com/en/1.1.0/)
[![Versioning](https://img.shields.io/badge/Versioning-SemVer-6f42c1?style=flat-square)](https://semver.org/spec/v2.0.0.html)
[![Dates](https://img.shields.io/badge/Dates-YYYY--MM--DD-217346?style=flat-square)](#date-and-version-rules)
[![Staging](https://img.shields.io/badge/Staging-Unreleased_first-d97706?style=flat-square)](#unreleased)
[![Contributing](https://img.shields.io/badge/Changes-Contribution_guide-2ea44f?style=flat-square)](CONTRIBUTING.md)

<br>

**User-visible history · Explicit compatibility · Reproducible evidence · Immutable releases**

</div>

---

All notable changes to **{{PROJECT_NAME}}** are documented here.

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

## [1.2.1] - 2026-09-22

### Changed

- Added pinned CodeQL analysis for maintained Python/JavaScript, weekly GitHub Actions update proposals through Dependabot, and trusted-context OpenSSF Scorecard publication; pull-request CodeQL now runs as a separate read-only, non-publishing analysis path and `pull_request_target` remains prohibited by repository validation.
- Hash-locked the hosted CPython 3.10 quality/coverage/portfolio installs to reviewed wheel SHA-256 digests with `--only-binary=:all: --require-hashes`, while retaining cross-platform local version pins and mirror checks.
- Refreshed all canonical `github/codeql-action` occurrences together from v4.38.0 to the official v4.38.1 commit `1c5b675653bb5c22dbe9b12b556ec555138e09fd`; the release adds experimental per-language bundle support without changing this repository's CodeQL permissions, triggers, query set, or analysis categories.
- Added reproducible maintainer Python coverage with subprocess measurement, a 95% statement floor, and focused behavioral failure-path fixtures for release, template-contract, Wiki, initializer, closeout, and reusable-workflow tooling.
- Added a read-only post-release closeout workflow that binds the annotated tag, tag-triggered CI, GitHub Release state, candidate-bound asset policy, comparison range, actual milestone membership, Wiki read-back, and provider-generated source archive observations to one certified SHA; uploaded assets remain distinct from GitHub source archives.
- Defined changelog release dates as the reviewed release-section cut/freeze date rather than tag or publication timestamps. Release-semantic evidence now names that meaning explicitly, accepts same-day releases, and rejects backward cut/freeze-date ordering without relying on wall-clock or provider timestamps.
- Made provenance-record signature mode an explicit release-policy selector. Contract 1.2.0 candidates now fail closed when the selector is missing or unsupported, or when the committed provenance trust policy attempts to use a different mode; Git-tag signing remains a separate control.
- Added a reproducible local validation environment for template maintainers: `tools/requirements-dev.txt` pins the Python tooling to the versions hosted CI installs, and `tools/dev_check.sh` runs the locally reproducible gates in the hosted order. It fails when a pin no longer mirrors the workflow that owns it or when an installed version does not match the pin, and it reports coverage, Excel evidence, live GitHub state and — unless `actionlint` is on `PATH` — workflow validation as skipped rather than passed.
- Replaced the disposable per-task workflow pattern with `.github/workflows/maintenance.yml`, one permanent dispatch-only home for maintenance tasks that must run in the pinned environment and commit their result. The task list is a closed `choice` passed to the shell through `env`, write scope is granted on the job rather than the workflow, the result is validated before it is committed, and a task that changes nothing exits without an empty commit. Both files are template-only and are removed from generated projects by initialization.
- Added canonical-template release-history validation from the previous release tag to the candidate SHA. Unapproved merge commits and duplicate commit subjects now block release semantics; base-tag/SHA-scoped reviewed exceptions remain auditable, while v1.2.0 merge ancestry is preserved as historical evidence and generated projects do not inherit this policy.

### Fixed

- Harden workflow-policy recognition so scalar and simple flow-mapping pull-request triggers and flow-mapping write permissions cannot bypass the documented read-only and prohibited-trigger controls.
- Apply external-link URL policy before historical or pending-publication classifications, and render the same five aggregate outcome counts in JSON and Markdown, including access-restricted links.
- Report active release-history exceptions whose base tag no longer matches the candidate's current previous-release tag, while preserving durable prior-release ancestry only as descriptive historical evidence.
- Removed the obsolete development-only duplicate-history exception after the v1.2.1 release branch was squash-integrated, so release semantics evaluate only commits that remain in the canonical mainline release range.
- Revalidate required certification evidence roles and manifest record structure during independent bundle verification, so a self-consistent but incomplete durable bundle cannot pass post-publication closeout.
- Sign the canonical durable certification ZIP with a detached OpenSSH signature bound to the current trusted GitHub signing-key registry and verify that signature again from downloaded GitHub Release assets.
- Make release closeout expectation-aware for deliberately non-latest releases and first prereleases, capture every page of GitHub comparison commits, and render the latest-release control consistently with the configured expectation.
- Keep the canonical template README presentation live by using real CodeQL, OpenSSF Scorecard and social-preview URLs while preserving deterministic repository and preview retargeting during initialization.
- Split the README's aggregate open-issue badge into live P1, P2 and P3 issue badges that link directly to the corresponding open-priority queues.
- Update the OpenSSF Scorecard badge and viewer links to the current official `api.scorecard.dev` and `scorecard.dev` endpoints.
- Make Scorecard publication fail closed: keep OpenSSF-forbidden global/job environment settings out of the publishing path, verify the exact-SHA public API record after the Action completes, validate the actual badge SVG so error badges such as `invalid repo path` cannot pass silently, and use Git file enumeration so file-based Scorecard checks see the reviewed repository tree.

## [1.2.0] - 2026-09-10

### Added

- Added a discovered CLI self-test interface audit with explicit alternative-suite exclusions and negative controls for missing or stale declarations. Documentation and wiki gates retain their dedicated offline suites; runner ownership remains independently enforced.
- Added reviewed wiki source for the complete new-repository journey, a checked file/directory inventory with actual initializer lifecycle, and deterministic navigation/export/read-back tooling. Offline source checks remain separate from weekly/manual published-wiki observations; publication and pilot evidence are tracked independently from source preparation.
- Added deterministic documentation command/reference drift checks and a separate weekly/manual external-link workflow with bounded retries, redirects, timeouts and concurrency. Versioned domain approvals and expiring exceptions distinguish missing pages, transient failures and restricted access; reports omit raw URLs and query tokens.
- Added an optional Windows/Excel job interface and shared manual evidence schema, with exact-source/log bindings, environment and trust records, explicit expected-error results, and distinct import, compile, test, cleanup and unavailable outcomes. The validator does not execute Office.
- Added contract 1.2.0 release provenance: complete staged-payload checks, source/workflow/build-environment records, and optional SSH signature verification using committed trust policy. Source-only releases remain valid without binary artifacts.
- Added optional versioned provisioning policy and a template-maintenance provisioner that defaults to a read-only exact-state plan. Explicit apply requires the approved plan digest, initialized target SHA, trusted credentials and a durable journal; source changes, stale plans and partial writes cannot be reported as verified. Extra topics/labels and stronger rules are preserved, with simulated coverage for all three profiles and no live portfolio mutations.
- Added timestamped portfolio quality/conformance reports with explicit freshness, exact-SHA required workflow observations, branch/tag protection, published-release evidence and attributed specialist scores. Missing adoption and inaccessible evidence remain visible; popularity never contributes to quality, and publication never implies certification. Existing portfolio repositories remain read-only.
- Added a GET-only portfolio snapshot collector and deterministic structural drift evaluator. Adopted contract versions select required controls; missing adoption remains `ADOPT`, inaccessible evidence remains `UNVERIFIED`, and `DEFER` or profile exceptions cannot erase universal findings. Documented local specialist controls remain separate from generic compliance. Synthetic hosted fixtures exercise all three profiles without modifying portfolio repositories.
- Defined controlled dependency updates for Actions, reusable workflows, validation tools and execution runtimes. The policy requires manual monitoring and approval, verified source/release evidence, explicit trust-boundary review, exact candidate tests and a recorded rollback target; the PR template carries the evidence block. Existing dependency versions are unchanged, and no automatic updater or merge workflow is installed.
- Exposed the existing static checks as reusable workflow interface v1, with an optional generated-profile assertion and exact checked-commit output. Consumers pin workflow orchestration by full commit SHA while retaining their reviewed local gate scripts and specialist jobs. Added reproducible consumer fixtures and compatibility, deprecation and rollback guidance in `docs/REUSABLE_WORKFLOWS.md`.
- Versioned the template contract independently from the generated project's product version. `.github/repository-profile.json` now records `template_contract` with the adopted contract version and the template repository that published it; `VERSION` continues to describe only the project, and changing it never rewrites the adopted baseline. `docs/TEMPLATE_CONTRACT.md` is the authority for the contract's SemVer policy and carries migration notes for every supported version, classified as breaking, required, optional or not applicable.
- Added `tools/check_template_contract.py`, the focused gate owning the contract's semantics. It resolves the rule set registered for the *recorded* version, so a repository is evaluated against the contract it adopted rather than the newest one; rejects unsupported or non-canonical versions with a message naming the supported set; requires migration notes for every supported version; and enforces the source invariants — a template publishes its own contract, a generated repository names the template it adopted and never itself. Initialization now records the adopted contract and the initializer self-test proves it survives generation unchanged for all three profiles.

### Changed

- Retained the canonical template's live external-link observation, complete clean-room maintainer journey, exact-source Wiki publication/read-back, and browser navigation review as explicit pre-tag release blockers in the maintained release authority.
- Reconciled the documentation and source comments against the v1.2.0 candidate: clarified initialization commit/no-op order, advisory Python line length, runner ownership, historical pilot scope, token permissions, parser behavior and complete Wiki publication. Improved directory-guide presentation and removed a personal portfolio-size assumption.
- Applied the maintainer's VBA house style to all four starter modules: ordered module banners, procedure contracts, aligned declarations with inline annotations, and explanatory comments beneath execution-phase banners and before key steps. Added `docs/VBA_HOUSE_STYLE.md` as the reusable presentation reference. Executable statements, public API, error contracts and regression output are preserved.
- Consolidated the focused-gate CLI orchestration that was provably identical across gates into the typed `run_gate` runner in `tools/_gatelib.py`: `--self-test` dispatch, canonical JSON serialization, Markdown summary writing, console output and the pass/findings/could-not-complete exit mapping. The initial eight consumers were extended by the later v1.2.0 gates; `checker_development.py` records the complete consumer/exclusion registry. `check_repo.py` remains a self-contained single file that never imports the helper. Each gate keeps its own semantic rules, fixtures, report schema, Markdown renderer and operational-exception tuple, so no gate's exception handling was widened and programming errors still raise instead of being reported as an operational exit.
- Extended the checker-development contract to own the shared-runner boundary: every tool defining a top-level `main` must be a declared `run_gate` consumer or a documented exclusion, no tool may redefine the shared helpers locally, and sixteen independent unit tests exercise the runner's CLI flags, defaults and `--help`, self-test dispatch, both self-test diagnostic prefixes, pass/fail/operational exits, deterministic evidence, report-write failures and exception propagation.

### Fixed

- Recognize chained CLI guard comparisons while preserving negation and rejecting contradictory or literal-false chains.
- Preserve Boolean guard polarity in CLI discovery so import-only and impossible branches do not trigger the self-test interface audit.
- Discover executable CLI guards independently of entry-function names in the self-test interface audit, including module-qualified calls and unittest scripts.
- Initialized projects may customize or remove README badges without failing
  initializer self-tests; default badge URLs are checked only in fresh template fixtures.
- The public template README now uses working repository badge URLs; initialization
  retargets the status, release and issue badges and links to the generated repository.
- Distinguish local reusable-workflow job calls from local action steps, and require tracked workflow files.
- Ignore quoted VBA text when checking jump targets, while retaining checks for executable jumps on the same line.
- Validate public API declarations across the supported conditional-compilation environments; accept mutually exclusive variants and require every distinct signature in the manifest.
- Restore bounded exponential retry delays in label synchronization and drift checks when `Retry-After` is absent or empty.
- Run the checker-development contract and the label-drift and label-validation fixtures on `release/**` pull requests, and run the checker-development contract on `release/**` pushes, so milestone work developed on a release branch is gated by the same hosted checks as `main`.
- Keep automatic live label reconciliation on trusted path-filtered pushes to `main`; manual dispatch remains an explicit live path. Label-drift detection uses scheduled/manual live reads and offline-only pull-request fixtures. Release-branch pushes alone do not mutate labels.

## [1.1.0] - 2026-09-05

### Added

- Added committed-candidate whitespace validation with explicit commit-range,
  root-commit, staged, unstaged, and committed-versus-working-tree fixtures.
- Added procedure-scoped VBA jump validation and nested conditional-compilation
  validation across the supported VBA6/VBA7 and Win32/Win64 environments.
- Added complete public-API extraction with normalized signature records,
  paired-property handling, collision detection, and strict explicit visibility.
- Added repository-local GitHub Action containment, tracked-state, metadata, and
  entrypoint validation alongside the pinned authoritative workflow parser.
- Added strict release-semantics validation for SemVer precedence, prerelease
  identifiers, changelog ordering, Gregorian dates, VERSION agreement, and
  comparison-link policy.
- Added deterministic semantic policy-branch assurance proving every canonical
  blocking finding site is exercised; the v1.1.0 branch baseline covers 175/175
  production finding sites.
- Added an independent checker-development contract that preserves
  `tools/check_repo.py` as a single-file, standard-library-only distributable
  while testing parser/reporting boundaries, CLI behavior, canonical check order,
  and artifact identity.
- Added read-only live issue-label drift detection with deterministic
  create/update/delete evidence, canonical-plan cross-checking, scheduled/manual
  monitoring, and retained JSON/Markdown evidence without a mutation path.
- Added an explicit Python 3.10 tooling baseline with pinned Ruff and mypy checks
  in the hosted repository-integrity workflow, including retained lint/type
  evidence and fail-closed terminal enforcement.

### Changed

- Extended the hosted repository-integrity gate so focused hardening checks have
  deterministic self-tests, exact-candidate evidence, artifact retention, and
  fail-closed terminal enforcement.
- Consolidated documentation around one authoritative owner per evolving
  contract. The root README is now a shorter first-use/navigation surface while
  installation, contribution, security, conduct, release, initialization,
  repository structure, checker development, and release evidence remain in
  their specialized maintained documents.
- Required generated repositories to retain the read-only label-drift control in
  addition to the existing trusted label-reconciliation workflow.
- Removed the completed temporary implementation plan and redirected durable
  governance references to maintained contracts and historical evidence.

### Fixed

- Disarmed the regression runner's procedure-level error handler before cleanup
  and summary reporting so a cleanup/reporting fault cannot re-enter `CleanExit`
  indefinitely through `RunFailed`.
- Made regression cleanup evidence substantive by comparing the relevant Excel
  application state with the pre-run snapshot rather than merely re-reading an
  immediately cleared re-entry flag.
- Consolidated duplicated focused-gate Git, report-output, tracked-file, and
  common CLI mechanics into the private standard-library-only `tools/_gatelib.py`,
  while keeping `tools/check_repo.py` explicitly self-contained.
- Removed residual unused imports exposed by the enforced Ruff baseline.
- Reduced Python checker complexity under a permanently enforced McCabe ceiling
  of 20 and normalized the VBA public-API checker so no style exception remains.
- Hardened stdlib XML validation with a bounded input size and fail-closed
  rejection of DTD/entity declarations before parsing.
- Split template-maintainer checker-development and semantic policy-coverage
  tooling from the operational tool payload retained by generated repositories.

### Compatibility

- The neutral starter VBA public API and profile architecture are unchanged from
  v1.0.0. The release hardens repository governance, validation, documentation,
  and automation without intentionally changing the starter consumer contract.

## [1.0.0] - 2026-09-04

### Added

- Added a dependency-free release-integrity gate, versioned profile policy,
  external evidence contract, SHA-256 asset-manifest validation, annotated-tag
  verification, and deterministic positive and negative fixtures for generated
  projects and the canonical template itself.
- Added a classified double-brace token schema and deterministic,
  dry-run-first repository initializer with atomic validation and all-profile
  fixtures.
- Added a profile-driven, dependency-free repository-quality gate with
  deterministic JSON and Markdown evidence plus positive and degraded self-tests.
- Added the canonical root README with profile selection, generated-repository
  initialization, structure, source policy and quality boundaries.
- Added a neutral, importable VBA façade/core starter, a fixed public API
  manifest, a state-safe example, and a deterministic four-case regression
  harness with equality, tolerance, expected-error, environment, completeness,
  and cleanup reporting.
- Added per-profile substantive VBA contracts and full-tree fixtures proving
  that README-only trees or missing façade/core/test components fail while
  optional example removal remains valid.
- Added structured bug, feature and documentation forms, private security
  routing, a post-creation provisioning checklist and static form fixtures.
- Added content-pinned authoritative GitHub Actions validation with positive,
  malformed-YAML, duplicate-schema and local-action fixtures, plus direct XML
  and YAML branches in the portable checker self-test.

### Changed

- Renamed the canonical repository to `EXCEL-VBA-PROJECT-TEMPLATE` and aligned
  its profile identity, security route, evidence links and template-identity
  rejection rule with the new URL.
- Replaced the progress-oriented implementation plan with an exact-snapshot code
  review, weighted template-readiness score, P1/P2/P3 finding register and
  findings-driven certification sequence.
- Rebuilt the canonical README from a seven-repository benchmark, combining a
  premium identity block, quick navigation, profile-aware initialization,
  architecture, assurance boundaries, recovery, security and release guidance.
- Upgraded the canonical static-check workflow with bounded execution,
  non-persistent checkout credentials, deterministic JSON and Markdown
  artifacts, rerun-safe evidence names and an explicit terminal verdict.
- Made profile and domain label selection a versioned repository policy that
  both the checker and trusted reconciliation workflow validate and consume.

[Unreleased]: https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/compare/v1.2.1...HEAD
[1.2.1]: https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/compare/v1.2.0...v1.2.1
[1.2.0]: https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/danielep71/EXCEL-VBA-PROJECT-TEMPLATE/releases/tag/v1.0.0

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
