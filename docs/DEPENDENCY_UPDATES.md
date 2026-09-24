# 🔄 Controlled Dependency Updates

[![Policy: reviewed updates](https://img.shields.io/badge/policy-reviewed%20updates-217346)](#review-and-merge)
[![Pins: immutable](https://img.shields.io/badge/actions-full%20SHA-1D76DB)](#provenance-review)
[![Merge: manual](https://img.shields.io/badge/merge-manual-6f42c1)](#review-and-merge)

This document owns dependency discovery, update evidence, approval and rollback.
[REUSABLE_WORKFLOWS.md](REUSABLE_WORKFLOWS.md) owns the reusable interface;
[TEMPLATE_CONTRACT.md](TEMPLATE_CONTRACT.md) owns contract compatibility.
The policy applies to template maintenance and initialized repositories alike.

## 🎯 Monitoring scope and ownership

The repository maintainer reviews the following sources weekly and before a
release candidate is selected. Security advisories are triaged when received,
without waiting for the weekly review. Record the review date, reviewer, sources
checked, candidates, disposition and next action in a tracking issue; a no-change
review must say so. Dependabot now performs weekly proposal discovery for GitHub
Actions. It never auto-approves or auto-merges; provenance review and the manual
merge decision below remain mandatory. Other dependency surfaces retain their
explicit maintainer review cadence.

| Dependency surface | Authoritative locations / discovery source | Update unit |
| --- | --- | --- |
| External Actions | Every `uses:` in `.github/workflows/`; official `actions/checkout`, `actions/setup-python`, `actions/upload-artifact` releases and advisories | Full commit SHA plus audited semantic-version comment; update all intended occurrences |
| GitHub Actions update discovery | `.github/dependabot.yml`; Dependabot pull requests are proposals only | Weekly candidate proposal; preserve immutable SHA pins and manual approval |
| Security analyzers | `.github/workflows/codeql.yml`, `.github/workflows/scorecard.yml`, and [SUPPLY_CHAIN_ASSURANCE.md](SUPPLY_CHAIN_ASSURANCE.md) | Reviewed Action pin, query/analyzer behavior, permissions, triggers, and publication boundary |
| Reusable workflows | Caller job-level `uses:`, [published interface pin](REUSABLE_WORKFLOWS.md), template source and its compatibility notes | Full provider SHA plus interface revision; review coupled caller-local scripts separately |
| actionlint | `ACTIONLINT_VERSION` and archive digest in the static workflow; `EXPECTED_ACTIONLINT_VERSION` in `tools/test_workflow_validation.py`; official `rhysd/actionlint` releases | Version, platform-specific archive SHA-256 and expected validator version together |
| Python quality tools | `RUFF_VERSION` and `MYPY_VERSION` in the static workflow; `tools/requirements-quality-ci.txt`; official PyPI distribution metadata and upstream release notes | Exact direct/transitive versions plus reviewed wheel SHA-256 hashes for hosted CPython 3.10/Ubuntu x64; local installation follows the retained developer setup; no separate cross-platform lock is retained |
| Runtime and execution environment | `PYTHON_VERSION`, `runs-on`, Action `runs.using`, `pyproject.toml` target/lint/type configuration and workflow install commands | Explicit compatibility change, not an incidental pin refresh |
| Local validation code and fixtures | Retained `tools/` scripts and their operational fixtures; canonical template-only maintenance tools are absent here | Reviewed source commit; not a separately fetched package |

Search all tracked workflows, scripts, documentation examples and fixture
generators, not just the static workflow. Record deliberately unchanged
occurrences and historical evidence. Never rewrite published provenance to make
an old release appear to use the new dependency.

Python `3.10` selects a runtime series, and `ubuntu-24.04` selects a hosted image
family; neither freezes every patch or preinstalled utility. Hosted Python tool
installs are additionally constrained by reviewed CPython-3.10/Linux-x64 wheel
hashes in `tools/requirements-quality-ci.txt`, installed with
`--only-binary=:all: --require-hashes`. That materially narrows package
substitution risk but still does not make the runner image or network path
hermetic. Capture actual versions in hosted evidence. This generated repository
retains no separate local requirements file or local-to-CI lock comparison.
Follow [DEVELOPER_SETUP.md](DEVELOPER_SETUP.md) for portable standard-library
checks and the matching Linux quality-tool environment. This policy adds no
package manager to VBA runtime code.

<a id="provenance-review"></a>

## 🔎 Provenance review before execution

For every proposed change, complete the dependency block in the
[PR template](../.github/PULL_REQUEST_TEMPLATE.md). Each dependency needs its own
old/new identifiers, evidence and rollback target, including non-SHA tool
versions. Fields must contain facts or `NOT APPLICABLE` with a reason, not an
unchecked assertion that a bot looked after provenance.

1. Start with the official upstream owner/repository or distribution index.
   Resolve the advertised release tag to its commit; peel annotated tags.
   Verify the commit belongs to that upstream, not merely a fork. Record the
   release URL, immutable commit URL, verification date and reviewer. A version
   comment or matching 40-character shape is not proof of origin.
2. Read release notes, advisories and the old-to-new source diff. Check ownership
   or maintainer changes, repository transfers, renamed/discontinued packages,
   release-source changes, license changes and support/runtime requirements.
   Unknown provenance blocks the update; do not silently switch mirrors.
3. Inspect Action metadata and the shipped executable/bundle, not just source
   TypeScript. For downloaded binaries, record the exact asset and expected
   digest from the official release and verify the downloaded bytes. Record
   signature/attestation verification when supplied; absence is a documented
   limitation, never a fabricated verification result. A checksum served by the
   same compromised source is not an independent authenticity proof.
4. Compare token permissions, secrets, checkout refs, triggers, runner type,
   artifact trust, network/install steps and reusable-workflow inputs. Surface
   every change explicitly. Never expand execution authority as an incidental
   dependency update or run proposed code with production credentials.
5. Record the reason: security fix, defect correction, required compatibility or
   maintenance. Newer is not sufficient by itself. If deferred, record why and
   when to revisit; handle sensitive vulnerabilities through
   [SECURITY.md](../SECURITY.md).

The SHA/origin and least-privilege principles follow GitHub's
[Secure use reference](https://docs.github.com/en/actions/reference/security/secure-use).

## 🧩 Grouping and compatibility

- Group repeated occurrences of the same Action pin in one coherent update.
- Keep an actionlint version, archive digest and validator expectation together.
- Couple a reusable-workflow pin with local CLI changes only when required by
  its documented interface migration; explain and test the pair.
- Separate unrelated dependencies, major upgrades, ownership changes, new
  permissions, new download sources and runner/runtime migrations. A shared
  tooling compatibility constraint may justify grouping, with a stated reason.
- Do not opportunistically change VBA behavior, formatting policy or required
  checks while refreshing pins. Contract-affecting changes need classification
  and migration notes under [TEMPLATE_CONTRACT.md](TEMPLATE_CONTRACT.md).

<a id="review-and-merge"></a>

## ✅ Validation, review and merge

Before executing a candidate, finish the source/trust review above. Then:

1. Run `python3 tools/check_repo.py --root .` and `git diff --check`. The
   canonical gate checks full lowercase SHA pins and version-comment syntax.
2. Run the complete hosted static workflow, not only the changed tool. It must
   pass lint, typing, authoritative workflow validation, focused/self-tests,
   initializer checks, evidence upload and the final verdict. Record the exact
   tested commit and hosted run; new commits invalidate old candidate evidence.
3. In K-PRICING, run the retained operational fixtures, including
   `python3 tools/test_verification_depth.py -v`, the recorded-profile initializer
   self-test and applicable specialist jobs. Checker-development, policy-coverage
   and three-profile consumer generation belong to canonical template maintenance;
   their tools are absent here and are not K-PRICING merge requirements.
   For reusable-workflow changes, apply the retained interface and specialist
   boundaries in [REUSABLE_WORKFLOWS.md](REUSABLE_WORKFLOWS.md).
4. Assess Excel, numerical, UI and packaging impact. Run affected specialist
   tests; otherwise record why they are not required. A pin-only CI change is
   not a new Excel compilation or runtime certification.
5. Have the maintainer review the evidence block and full diff, then record an
   explicit approval bound to the final candidate SHA. With one maintainer this
   is a recorded maintainer decision, not a claim of independent approval.
6. Merge manually through the reviewed PR workflow. Verify the active branch/tag
   rulesets in
   [SETUP_VERIFICATION.md](SETUP_VERIFICATION.md). Keep repository automatic
   merging disabled; do not introduce auto-approve, auto-merge, bypass tokens or
   unattended pin-changing workflows. If stronger review rules are configured,
   satisfy them. Never bypass failing or missing required checks.

These controls have different enforcement boundaries. The static gate verifies
pin syntax; hosted jobs verify tested behavior. **Release provenance, completeness
of the PR evidence, maintainer approval and trust-boundary review are human
merge gates**, not assertions parsed or certified by the static checker.
Generated repositories must configure and verify their own live merge settings
using [POST_CREATION_CHECKLIST.md](POST_CREATION_CHECKLIST.md); files alone cannot
enforce GitHub account settings. Dependabot is the installed proposal bot for
GitHub Actions only; it must retain immutable pins and manual approval. CodeQL
and OpenSSF Scorecard are subject to the private/public eligibility policy in
[SUPPLY_CHAIN_ASSURANCE.md](SUPPLY_CHAIN_ASSURANCE.md), not dependency approval.

## ↩️ Rollback and failed updates

Record both the previous **dependency identifier** and the **repository commit**
containing the last passing combination before proposing an update. For Actions
and reusable workflows the rollback identifier is the old full SHA; for tools
it is the old exact version and, where applicable, archive digest. Retain the
previous evidence and any coupled configuration needed to reproduce it.

- Before merge: reject or revise the proposal. Do not deploy it to unblock CI.
- After merge: create a reviewed revert commit on the active development
  branch. Revert the update commit, or revert the complete grouped migration in
  reverse order. A merge commit requires deliberate mainline selection; do not
  guess a `git revert -m` parent. Resolve conflicts explicitly.
- Compare restored dependency values and coupled configuration against the
  recorded baseline. Rerun generic and affected specialist checks on the revert
  commit; restoration is not itself fresh test evidence. Record the revert SHA
  and successful run before manual merge.
- Do not force-push, move a release tag, edit published evidence or disable a
  required check. If the former version is compromised or unavailable, stop and
  select a separately reviewed safe baseline; there is no automatic fallback.

The dependency PR record is the durable update ledger. It preserves old/new
identifiers, source review, limitations, candidate evidence, approval and the
rollback path in one place. This policy establishes the process; it does not
retroactively certify older dependency changes or claim the current versions
are the latest available.
