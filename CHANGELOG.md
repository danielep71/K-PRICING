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

### Added

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
  hashes, license attribution and explicit pending destination runtime evidence.
- Accepted manual neutral-starter Excel evidence for candidate
  `c8f0b6ee147d07549484ec83743f5b7dfc0ea1f2` on one Windows/64-bit Office host; no pricing
  or application-release certification.
- Project-specific developer setup, neutral Excel runbook, dated GitHub verification
  and a frozen-source migration inventory with issue traceability.
- Initialized the K-PRICING application scaffold with project identity, private
  reporting contact, profile directories and reproducible initialization inputs.

### Changed

- Updated the migration plan to record completed migration issues with their
  merge commits, the source implementation-plan disposition and the pending
  #17/#18 gates; the production source guide now states which modules are
  byte-identical to the frozen KPR source and which carry destination deltas.
- Reset the development version to `0.0.0` and removed template-maintenance
  history, assets and tools through the deterministic initializer.

### Security

- Enable public-repository CodeQL/Scorecard paths, private vulnerability reporting
  and branch/tag rulesets; retain conditional safeguards for private visibility.
- Record the maintainer's 2026-09-24 decision to retain the existing email and
  host details publicly, as documented in setup verification and issue #29.

### Fixed

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
  authorities, validation commands, and neutral-starter entry points.
- Remove obsolete private-link classifications after the public-visibility
  read-back; retain honest anonymous reachability results.
- Reconcile visibility records, documentation authorities, ASCII source rules
  and the migrated regression harness visibility exception.
- Allow the retained `KPR_` VBA namespace through the project identity policy, while preserving unrelated donor and
  template restrictions ([#2](https://github.com/danielep71/K-PRICING/issues/2)).

### Known limitations

- The frozen KPR date-layer source, regression harness and project-specific static
  contract guard are migrated. Exact-source Excel compilation and parity passed
  on one Windows 64-bit host only (#17); 32-bit Office, other builds and locales
  are untested, and release packaging remains pending.
- The accepted neutral-starter Excel run remains setup history only and does not
  certify the migrated pricing/date implementation.

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
