# 🔖 Release Semantics Contract

[![SemVer: strict](https://img.shields.io/badge/SemVer-strict-3f4551)](https://semver.org/spec/v2.0.0.html)
[![Changelog: ordered](https://img.shields.io/badge/changelog-ordered-0969da)](../CHANGELOG.md)
[![Dates: cut%2Ffreeze](https://img.shields.io/badge/dates-cut%2Ffreeze-217346)](../CHANGELOG.md#date-and-version-rules)
[![Gate: fail closed](https://img.shields.io/badge/gate-fail%20closed-success)](../tools/check_release_semantics.py)

This document defines release-only version and changelog semantics for the
canonical template and initialized repositories. The generic repository checker
keeps only broad structural checks; release correctness is enforced separately
by `tools/check_release_semantics.py` and `tools/check_release.py`.

## SemVer contract

`VERSION` contains one SemVer 2.0.0 value without a leading `v`.

- Major, minor, and patch numeric identifiers never contain leading zeros.
- Pre-release identifiers follow SemVer precedence exactly.
- Numeric pre-release identifiers never contain leading zeros (`rc.01` is
  invalid; `rc.1` is valid).
- Build metadata is accepted but does not affect precedence.
- A stable release has higher precedence than a pre-release with the same core
  version.

The tag is formed by adding the lower-case `v` prefix to `VERSION`.

## Changelog contract

`CHANGELOG.md` contains exactly one `## [Unreleased]` heading. Dated release
headings use exactly:

```text
## [MAJOR.MINOR.PATCH] - YYYY-MM-DD
```

Pre-release versions use the same heading form with their valid SemVer suffix.
Every date must be a real Gregorian calendar date.

### Changelog date semantic

The date in a released changelog heading is the **release-section cut/freeze
date**: the calendar date on which the reviewed `Unreleased` content is moved
into that versioned section and the release section is frozen for the candidate.
It is not the Git tag creation date, GitHub Release publication date, first
installation date, or a timestamp inferred from Git history.

Tagging and publication may occur later than the cut/freeze date. That delay does
not make the changelog heading stale and must not cause a published historical
section to be rewritten. The canonical v1.2.0 record is therefore valid with a
`2026-09-10` changelog cut/freeze date and tag/publication on `2026-09-12`.

If a candidate is abandoned and the release section is materially reopened, a
later reviewed candidate may be cut/frozen again with a new date. That is a new
release-preparation event, not a cosmetic attempt to match a future publication
date.

Across newest-to-oldest release headings, cut/freeze dates must not move
backward. Same-day releases are valid. This ordering is deterministic and is
enforced by `tools/check_release_semantics.py` in addition to Gregorian-date
validation.

The gate deliberately does **not** compare changelog dates with commit, tag, or
provider publication timestamps. Those timestamps do not prove when the
release section was frozen, and wall-clock equality would make otherwise valid
release preparation brittle and timezone-dependent. The maintainer release
sequence records the cut/freeze event; the static gate proves the parts that are
available from candidate source alone.

Dated releases appear newest to oldest by **full SemVer precedence**, not by
lexical text or date alone. Duplicate release versions are invalid. When
`VERSION` is not the generated-project development sentinel `0.0.0`, it must
match the newest dated release heading during a release candidate.

## Comparison-link policy

When at least one release exists:

- `[Unreleased]` compares the latest release tag to `HEAD`:

  ```text
  [Unreleased]: https://github.com/OWNER/REPOSITORY/compare/vLATEST...HEAD
  ```

- the initial release links to its immutable release tag:

  ```text
  [1.0.0]: https://github.com/OWNER/REPOSITORY/releases/tag/v1.0.0
  ```

- every later release compares the immediately preceding release tag to the new
  release tag:

  ```text
  [1.1.0]: https://github.com/OWNER/REPOSITORY/compare/v1.0.0...v1.1.0
  ```

Missing, duplicated, stale, or mismatched comparison links fail release
semantics.

## Canonical-template history policy

The canonical template additionally applies a squash-by-default history rule to
the Git range from the previous reachable release tag to the candidate SHA. This
branch is **template-maintainer policy only**: initialized/generated repositories
receive no blocking history verdict unless they deliberately adopt an equivalent
local rule.

In the canonical template, the committed authority is
`.github/release-history-policy.json`; that file is template-only and is not
present in K-PRICING. The gate reads
that file from the exact candidate Git object rather than trusting a mutable
working-copy override. Two conditions are blocking by default:

- a multi-parent commit in the candidate release range; and
- a duplicate normalized commit subject in that same range.

A deliberate history-preserving merge can be approved by adding an exception
that binds all of the following before release certification:

- the previous release `base_tag` that scopes the exception;
- the exact 40-character commit SHA;
- the permitted finding type (`merge-commit` and/or `duplicate-subject`);
- a GitHub issue or pull-request review reference; and
- a non-empty reason explaining why ancestry preservation is required.

Every entry in `exceptions` is active release policy and must name the **current
previous-release base tag** for the candidate being inspected. An entry carrying
any other `base_tag` is stale and produces
`superseded-history-exception-base`; it is never silently ignored. When an
exception's base is superseded, remove the active exception or, if the event must
remain documented, preserve it as descriptive evidence under
`historical_records`. Historical records never grant an exception to a later
release range.

Exceptions also fail closed when a current-base entry points outside the
inspected range or permits a condition that is not actually present. These are
reported separately from a superseded-base entry so the gate distinguishes
range staleness from base-scope staleness and prevents a standing or wildcard
exception from silently weakening future releases.

The same policy file retains historical records for the v1.2.0 stabilization
merges in PRs #68, #71, and #74. Those records are descriptive only: they keep
published ancestry visible without exempting any later candidate range or
requiring a rewrite of the immutable v1.2.0 history.

Release notes remain curated from the changelog, issues, and reviewed release
evidence. Raw commit subjects are never treated as an authoritative release-note
source, especially when an approved history-preserving exception exists.

## Validation

Run the deterministic policy fixtures:

```bash
python3 tools/check_release_semantics.py --root . --self-test
```

Validate the current tree and retain evidence:

```bash
python3 tools/check_release_semantics.py \
  --root . \
  --output test-results/release-semantics.json \
  --summary test-results/release-semantics.md
```

The generated evidence names the changelog date semantic explicitly as
`release-section-cut-freeze-date`. For the canonical template it also records the
previous release tag, inspected commit range, reviewed exceptions actually used,
and retained historical records. Generated repositories report the history branch
as not applicable.

The self-test covers valid stable and pre-release versions, numeric pre-release
leading zeros, SemVer precedence, duplicate and out-of-order releases,
impossible dates, same-day cut/freeze dates, backward cut/freeze-date ordering,
`VERSION`/heading disagreement, missing or incorrect comparison links, compliant
squash history, an approved current-base exception, a superseded-base exception,
an unapproved merge commit, and duplicate-subject rejection.

A release candidate is not eligible for tagging unless this gate, the executable
release-integrity gate, the repository gates, and all applicable runtime evidence
are green for the same candidate source.
