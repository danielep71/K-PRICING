# Repository setup verification

This is the single authority for dated repository settings and visibility
read-backs. [Initialization status](INITIALIZATION_STATUS.md) owns historical
provenance; [setup completion](SETUP_COMPLETION.md) owns the accepted starter
checkpoint. Neither settings checks nor tooling scans execute Excel.

## Current read-back — 2026-09-24

Repository: `danielep71/K-PRICING`; default branch `main`; application profile,
no domain overlay. Collected through GitHub API and authenticated maintainer UI
as Daniele Penza. The repository was already public when inspected. This
correction did not change visibility. The actor, time and intent of the earlier
visibility change were not established; this read-back is not retroactive
approval of publication. The maintainer subsequently confirmed on 2026-09-24 that the existing email
and host details should remain public. That decision is recorded in
[K-PRICING #29](https://github.com/danielep71/K-PRICING/issues/29).

| Setting | Current observation |
| --- | --- |
| Visibility / default branch | Public / `main` |
| Branch ruleset | Active `Protect main` (23919706), default branch; no bypass actors; PR required, zero required approvals, all review conversations resolved; strict `Repository integrity` from GitHub Actions; deletion and force-push blocked |
| Tag ruleset | Active `Protect version tags` (23919838), `refs/tags/v*`; no bypass actors; all updates, deletions and force-pushes blocked; creation allowed |
| Private vulnerability reporting | Enabled; authenticated UI offers Disable; reporting route in [SECURITY.md](../SECURITY.md) |
| CodeQL | Public eligibility active; advanced setup; Python and JavaScript/TypeScript analysis, not VBA |
| Scorecard | Public default-branch publication and verification active |
| Dependency graph | Enabled |
| Dependabot alerts / security updates | Disabled; configured version-update PRs are a separate control |
| Issues / template / wiki / discussions / projects | Enabled / disabled / disabled / disabled / disabled |
| Merge / squash / rebase | Enabled; auto-merge disabled; delete merged branches and suggest updates enabled |

The zero-approval baseline supports one maintainer; it does not claim independent
approval. Re-read controls before release and whenever visibility or plan changes.
Successful main tooling runs at `01f50be937be181e32a21b4e7b8acd1415c49bf1`:
[static checks](https://github.com/danielep71/K-PRICING/actions/runs/35956382063),
[CodeQL](https://github.com/danielep71/K-PRICING/actions/runs/35956382017), and
[Scorecard](https://github.com/danielep71/K-PRICING/actions/runs/35956382038).
These are exact-revision results, not destination Excel certification.

### Public exposure and maintainer disposition

Tracked files, Git history and Actions logs are public; retained Actions artifacts
are available to readers subject to GitHub's sign-in and retention requirements.
The existing `evidence/` records include operator, run time and host/software
configuration. Personal contact text also remains in historical initialization
inputs, conduct policy and Git history. The current security-reporting route is
GitHub private reporting. On 2026-09-24, Daniele Penza explicitly chose to keep
the existing email and host details public, including their existing historical
records. This resolves the disclosure decision in issue #29; no redaction,
history rewrite or visibility reversal is requested. This decision concerns the
existing material; future evidence must still be reviewed for public disclosure.

## Historical setup observation — 2026-09-23

The repository was private at this checkpoint. The historical observations below
are retained to explain the accepted setup and former platform limitations;
current controls are recorded only in the table above. Exact starter Excel
results were supplied by the maintainer, not independently rerun here.

## Source and review evidence

The original generated source is `18b1713499d7ba8fe3accd598a817c42a7928f4d`;
the adopted template revision is `cb1f75bc20102f3a39478cfa4dfea04bbea696ee`.
Both have tree `0beb2b89b292e643ff236177083679391395d7c8`.
Contract version `1.2.0` does not assert that this snapshot is a release tag.
Initialization merged at `1d33ea3f66e1d3c3ff6c848c448e8c3e910c3679` in PR 1,
with successful main static run `35825795771`.
The namespace/branding correction merged in PR 19 at
`3047a761af316759110ccdb5141d7a62819edcd0`, after successful PR static run
`35829333395`. Both actionable bot threads were answered with fix/evidence and
resolved. This does not claim that the bot reviewed the final revisions.

## Historical post-creation read-back

This table records the 2026-09-23 setup checkpoint against
[the provisioning checklist](POST_CREATION_CHECKLIST.md). Formerly unavailable
controls are not marked passed retrospectively.

| Checklist area | Observed state / disposition |
| --- | --- |
| Initialized identity | Generated application profile, K-PRICING identity, 20 selected labels, no domains; initialization record agrees |
| Visibility/default branch | PASS: private / `main` |
| Description/topics | PASS: financial analytics and instrument pricing for Excel/VBA; `excel`, `vba`, `application`, `financial-analytics`, `instrument-pricing` |
| Features | PASS: issues enabled; template, wiki, discussions, projects and sponsorships disabled |
| Issue chooser | PASS: bug, documentation and feature forms render; security routes to this repository's policy; blank issues disabled for ordinary users (GitHub still offers maintainer-only blank issues) |
| Merge policy | PASS: merge, squash and rebase enabled; auto-merge disabled; delete merged branches and suggest branch updates enabled; GitHub default commit-message choices retained |
| Labels | Trusted sync performs reconciliation and exact post-read verification; live read-only drift result is recorded in issue 4 before closure |
| Branch/tag rulesets | UNAVAILABLE: API returned HTTP 403, “Upgrade to GitHub Pro or make this repository public to enable this feature.” No visibility or plan change made |
| Private vulnerability-reporting UI | UNAVAILABLE in the observed settings; use the existing private email route in SECURITY.md |
| Dependency graph/alerts/security updates | DISABLED in observed settings; version-update PR configuration is separate |
| Private CodeQL opt-in | NOT ENABLED: Actions repository/environment variables empty; eligibility not established |
| Application lifecycle/packaging | DEFERRED: no workbook/add-in, UI or packaging implementation in this setup milestone |
| Neutral starter Excel test | ACCEPTED: issue #8 closed; exact candidate `c8f0b6ee147d07549484ec83743f5b7dfc0ea1f2`, four cases, six assertions, zero failures, cleanup PASS; [retained evidence](SETUP_COMPLETION.md) |
| Issue milestones | Setup issues 2–10 assigned to v0.0.1; migration execution issues 11–18 assigned to v0.0.2 |

At that checkpoint, manual PR/check review was a process control without
server enforcement. The public-visibility recheck above supersedes those settings.
Application lifecycle and exact-candidate host evidence remain release prerequisites.

## Retained automation and permissions

All third-party actions remain pinned to full commit hashes. Hosted logs and
retained artifacts belong to a public repository; do not include confidential
inputs. No publishing token is supplied to source validation.

| Workflow | Trigger and execution boundary | Permissions / evidence |
| --- | --- | --- |
| Static repository checks | PR, main/release push, version tag, manual, reusable caller; applicable | Contents read; exact SHA, Python 3.10, Ruff 0.16.6, mypy 2.3.1, actionlint 1.7.12; reports 30 days |
| Sync issue labels | Relevant PRs run offline fixtures; relevant trusted main push/manual reconciles | PR contents read; trusted issues write; exact post-read verification |
| Detect issue-label drift | Relevant PR fixtures; daily 05:17 UTC, manual and changes to its workflow on main run live | Contents/ issues read only; no mutation; reports 30 days |
| External documentation links | Monday 06:17 UTC, manual and checker/policy/workflow changes on main | Anonymous probes, contents read; reports 14 days; restricted/transient observations remain non-green |
| CodeQL | Existing PR/push/schedule/manual paths gated by visibility/explicit private opt-in | Public jobs active; tooling-language scans cover Python/JavaScript, not VBA; PR analysis is read-only without SARIF upload; trusted upload uses security-events write |
| Scorecard | Existing paths require a public repository | Active here; publication verification requires successful publication job; the public path remains excluded for private repositories |
| Release closeout | Manual verification of an already published release | Not applicable to setup; actions/contents/issues read only; validates candidate, tag, release and milestone facts without publishing |

`Repository integrity` runs the retained verification-depth Python suite as an
enforced step and retains its output. It also runs documentation tests, release/provenance/Excel
record fixtures, release semantics, initializer fixtures, committed-whitespace,
VBA jump/conditional/public-API checks, local Action checks, template-contract
checks and the 21-rule repository gate. Step-level collection does not mask
failures: final enforcement fails on unsuccessful required outcomes. The source
workflow is the authority for exact commands and pins; use
[developer setup](DEVELOPER_SETUP.md) to reproduce them. Retained tests import
only retained operational modules; hosted imports and full CI verify this.

Dependabot checks GitHub Actions weekly on Monday at 06:00 Europe/Rome, up to
five PRs, labeled `ci` and `tests`. Python lock updates require deliberate review
and regeneration; there is no automatic pip-update configuration. Auto-merge
remains disabled. Follow [dependency review](DEPENDENCY_UPDATES.md).

The four stale private-target classifications were removed after the visibility
read-back. Anonymous probes now determine their status. Restricted, transient or
policy-blocked results remain non-green; public visibility alone is not proof
that every external link passes. See [documentation checks](DOCUMENTATION_CHECKS.md).

## Upstream lessons and safe updates

| Existing upstream issue | Local correction retained |
| --- | --- |
| `Upstream #121` | README checks use initialized project identity and permit valid profile changes |
| `Upstream #122` | Orphan maintenance imports removed; complete generated-project CI retained and executed |
| `Upstream #123` | Public Scorecard excluded for private repositories; CodeQL private opt-in requires confirmed eligibility |

The upstream repository identity is recorded in
[the initialization record](../.github/initialization.json); setup issue 7 retains
the direct issue links. These existing upstream issues belong to its v1.2.1 milestone; local setup does
not wait for upstream completion. The additional local identity policy allows
the retained `KPR_` namespace and historical source links while rejecting old
product branding and other forbidden identities. Preserve these deviations.

For a template update, pin the proposed upstream SHA, compare it with the adopted
revision, and identify applicable source/policy changes in a focused PR. Preserve
K-PRICING identity, visibility-dependent safeguards, selected labels and product
contracts. Review schema/version changes explicitly; do not rerun initialization
to overwrite the project. Run all retained CI and initializer fixtures, review
the final diff, then record the new adopted baseline and reconciled deviations.

## Completion boundary

v0.0.1 is a **milestone-only setup checkpoint**. Keep `VERSION` at `0.0.0` and
changes under Unreleased; no tag or release publication is authorized by the
milestone name. All setup issues #2–#10 and the milestone are closed. The
final setup correction merged at `715eda63365ba4a94e619976acc4721411c875f0`
in PR #23; issue #10 closed on 2026-09-23 at 20:22:54 UTC. The later Python
fixture fix in PR #24 does not change the accepted runtime source. Migration
follows [MIGRATION_PLAN.md](MIGRATION_PLAN.md).
No migrated numerical result, supported package or functional release is claimed.
