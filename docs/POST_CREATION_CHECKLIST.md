# ✅ Post-Creation Repository Checklist

[![Phase: after initialization](https://img.shields.io/badge/phase-after%20initialization-217346)](#confirm-initialized-source)
[![Live settings: required](https://img.shields.io/badge/live%20settings-required-6f42c1)](#repository-identity)
[![Evidence: read-back](https://img.shields.io/badge/evidence-read--back-success)](#capture-read-back-evidence)
[![Tags: immutable](https://img.shields.io/badge/tags-immutable-1D76DB)](#protect-releases)

Use this checklist immediately after generating and initializing a repository.
File generation does not inherit GitHub labels, rulesets, features, topics,
merge settings, security settings, or protected-tag policy.

Record the generated repository, selected profile, exact source SHA, operator,
and completion date before applying settings. Evidence must come from API or UI
read-back after each change, not from the setup command or intended values.

The optional `.github/provisioning-policy.json` records the initialized
description, topics, feature/merge defaults, required checks and reviewed
exceptions for template-maintenance provisioning tooling. Review and commit it
before generating a live setup plan. A plan alone changes no GitHub settings.

<a id="confirm-initialized-source"></a>

## 1. 🔎 Confirm the Initialized Source

- [ ] `.github/repository-profile.json` has `mode: generated`, the correct
  `repository`, one selected `profile`, and the intended `label_domains`.
- [ ] `.github/initialization.json` records the same profile and repository.
- [ ] `python3 tools/check_repo.py --root . --self-test` passes.
- [ ] `python3 tools/check_repo.py --root .` passes on the committed tree.
- [ ] The issue chooser renders bug, feature, and documentation forms; blank
  issues are disabled and the security link opens this repository's policy.

## 2. 🏷️ Resolve, Reconcile, and Monitor Labels

The versioned selection authority is `.github/repository-profile.json`:

- `profile` selects exactly one profile overlay in generated mode;
- `label_domains` lists zero or more kebab-case domain overlays; and
- `.github/labels.json` supplies the core, profile, and domain catalogues.

### Trusted reconciliation

On the first trusted push to `main`, or by a deliberate manual dispatch, run
`Sync issue labels`. The workflow must resolve policy itself; do not supply an
unrecorded profile or domain at dispatch time.

- [ ] The synchronization summary names `.github/repository-profile.json` as the
  policy source and lists the complete resolved core, profile, domain, and
  combined label sets.
- [ ] Post-run verification reports an exact match.
- [ ] API read-back contains every resolved label with the expected name,
  uppercase color, and description, and contains no label pruned by policy.

### Read-only drift detection

`Detect issue-label drift` is separate from reconciliation. Pull requests run
only deterministic offline fixtures. Scheduled and manual live checks use only
`contents: read` and `issues: read`, compare the live catalogue with the same
versioned policy, and have no mutation path.

The detector cross-checks every local create/update/delete difference against the
canonical reconciliation planner before reporting it. A no-drift run exits
green. Drift exits non-zero and records the exact observed and desired label
values in deterministic JSON and Markdown evidence retained for 30 days.

- [ ] The drift workflow is enabled on the default branch after generation.
- [ ] Its offline no-drift and simulated create/update/delete fixtures pass.
- [ ] A live no-drift check reports zero differences.
- [ ] The workflow permissions are read-only; it has no `issues: write` grant.
- [ ] Routine label mutation remains in the trusted `Sync issue labels` workflow.
  Initial provisioning may also use the template-maintenance provisioner with
  an explicitly approved exact-state plan and trusted credentials. It adds
  missing labels, requires recorded exceptions for replacements and never
  deletes extra labels. If extra labels are retained, add them to a selected
  overlay or review `prune: false` before routine synchronization runs.

<a id="repository-identity"></a>

## 3. 🪪 Repository Identity and Maintained Features

Set and then read back:

| Setting | Canonical expectation |
| --- | --- |
| Description | One specific sentence describing the supported Excel/VBA outcome and audience |
| Topics | `excel`, `vba`, the selected profile, and only maintained project-domain topics |
| Template repository | Enabled only for a repository intentionally distributed as a template |
| Issues | Enabled |
| Private vulnerability reporting | Enabled |
| Discussions | Disabled unless a maintained support/community process is documented |
| Wiki | Disabled unless it has an owner and version-drift control |
| Projects | Disabled unless a maintained board is part of project governance |
| Sponsorships | Disabled unless deliberately configured by the owner |

Never publish a private reporting address as an issue-form assignee or public
issue body. The issue chooser points to `SECURITY.md`, which owns the reporting
channels and disclosure process.

## 4. 🔀 Merge Policy

The canonical single-maintainer baseline is:

- [ ] merge commits, squash merging, and rebasing are enabled;
- [ ] automatic merging is disabled; dependency updates require the manual
  approval and merge process in [DEPENDENCY_UPDATES.md](DEPENDENCY_UPDATES.md);
- [ ] head branches are deleted after merge;
- [ ] contributors may update pull-request branches; and
- [ ] commit-message/title defaults are reviewed for the repository's release
  and attribution policy.

A project may narrow merge methods, but it must record the reason. It may not
weaken protected-branch, exact-SHA evidence, or release-provenance requirements.

## 5. 🛡️ Protect `main`

Create or verify one ruleset targeting the default branch:

- [ ] active enforcement with no default bypass actor;
- [ ] deletion and non-fast-forward updates blocked;
- [ ] changes require a pull request;
- [ ] zero mandatory approvals for the canonical single-maintainer baseline;
- [ ] required status check is exactly `Repository integrity` and is strict;
- [ ] conversation resolution and other review controls are enabled only when
  the repository can satisfy them without bypass; and
- [ ] the rule targets only the intended default branch pattern.

Stronger controls are allowed. A plan limitation that prevents rulesets must be
recorded explicitly; it does not justify weakening the public baseline.

<a id="protect-releases"></a>

## 6. 🔐 Protect Releases

Before the first stable release, create or verify an active ruleset targeting
`refs/tags/v*`:

- [ ] tag deletion is blocked;
- [ ] non-fast-forward tag updates are blocked;
- [ ] all updates of existing version tags are blocked, including fast-forward
  movement, using an `update` rule with no bypass;
- [ ] tag creation remains available so the documented pre-tag release gate can
  publish a new version; and
- [ ] no bypass silently permits a published tag to move.

Restricting creation with an empty bypass list would also prevent legitimate
new releases. `RELEASING.md` and the pre-tag release gate therefore own creation
authorization; the ruleset makes every matching tag immutable after creation.
Do not create the first stable tag until the release gate, selected-profile
pilot or equivalent deployment proof, and exact-SHA Excel evidence are
complete. The canonical template itself additionally requires all three
generated-profile pilots. Follow the
[`RELEASE_EVIDENCE.md`](RELEASE_EVIDENCE.md) contract for the final candidate.

## 7. 🧩 Profile-Specific Controls

| Profile | Required additional review |
| --- | --- |
| `library` | Public API, caller contract, compatibility, and focused regression evidence; no UI or workbook assets unless actually used |
| `ui-component` | UI lifecycle, state restoration, recovery, DPI/accessibility, callbacks, and any form/Ribbon resources actually used |
| `application` | Startup, shutdown, upgrade, recovery, packaging, workbook/add-in lifecycle, and end-to-end smoke evidence |

Profile differences add controls; they never remove the common façade, core,
test, issue-intake, branch, tag, or evidence baseline.

<a id="capture-read-back-evidence"></a>

## 8. 📸 Capture Read-Back Evidence

Preserve links or JSON responses for:

- repository metadata and enabled features;
- resolved labels and the successful trusted reconciliation run;
- the latest successful read-only label-drift check, or the exact drift report
  when a difference is detected;
- default-branch ruleset and exact required status context;
- protected-tag ruleset;
- private vulnerability-reporting state;
- issue chooser rendering and security routing; and
- the exact initialized commit and repository-quality run.

The record must identify collection time, repository, default branch, selected
profile, selected domains, and the account or automation that performed the
verification. Redact credentials and private contact details. Re-read the live
settings after any correction and retain the final state only as certification
evidence.

---

**Provisioning principle:** configure deliberately, detect drift read-only, read
back independently, and retain only final verified state as certification
evidence.
