# Repository setup verification

**Closeout update:** the subsequent manual starter run passed. See
[SETUP_COMPLETION.md](SETUP_COMPLETION.md) for the accepted candidate, evidence
and limits. Current documentation follow-up and final milestone acceptance are
tracked in [issue 10](https://github.com/danielep71/K-PRICING/issues/10).
Pending-run statements below describe the earlier setup checkpoint.


Observed on **2026-09-23** for private `danielep71/K-PRICING`, default branch
`main`, application profile, no domain overlay. Verification used the GitHub
API and the authenticated maintainer UI (Daniele Penza). Final candidate/run
links are retained in setup issues 4, 5 and 10 so the record does not attempt
to embed its own future commit hash. No Excel execution is claimed here.

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

## Live post-creation read-back

This table is the K-PRICING completion record for
[the provisioning checklist](POST_CREATION_CHECKLIST.md). Unavailable controls
are not marked passed.

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
| Neutral starter Excel test | PENDING issue 8; instructions prepared, no runtime evidence yet |
| Issue milestones | Setup issues 2–10 assigned to v0.0.1; migration execution issues 11–18 assigned to v0.0.2 |

Manual PR/check review is the current process control and **is not server-enforced
protection**. Daniele Penza owns resolving branch/tag protection before a
functional release and rechecking reporting/security eligibility whenever the
plan or visibility changes. Private visibility is deliberate. Application
lifecycle and exact-candidate host evidence remain release prerequisites.

## Retained automation and permissions

All third-party actions remain pinned to full commit hashes. Hosted artifacts
and logs remain within this private repository; do not copy them to a public
service. No publishing token is supplied to source validation.

| Workflow | Trigger and execution boundary | Permissions / evidence |
| --- | --- | --- |
| Static repository checks | PR, main/release push, version tag, manual, reusable caller; applicable | Contents read; exact SHA, Python 3.10, Ruff 0.16.6, mypy 2.3.1, actionlint 1.7.12; reports 30 days |
| Sync issue labels | Relevant PRs run offline fixtures; relevant trusted main push/manual reconciles | PR contents read; trusted issues write; exact post-read verification |
| Detect issue-label drift | Relevant PR fixtures; daily 05:17 UTC, manual and changes to its workflow on main run live | Contents/ issues read only; no mutation; reports 30 days |
| External documentation links | Monday 06:17 UTC, manual and checker/policy/workflow changes on main | Anonymous probes, contents read; reports 14 days; restricted/transient observations remain non-green |
| CodeQL | Existing PR/push/schedule/manual paths gated by visibility/explicit private opt-in | Private jobs skipped; eligible tooling-language scans cover Python/JavaScript, not VBA; trusted upload uses security-events write |
| Scorecard | Existing paths require a public repository | Skipped here; publication verification requires successful publication job; private source/results cannot enter its public publication path |
| Release closeout | Manual verification of an already published release | Not applicable to setup; actions/contents/issues read only; validates candidate, tag, release and milestone facts without publishing |

`Repository integrity` retains documentation tests, release/provenance/Excel
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

Private repository URLs are classified individually by fragment-free URL hash
as `ACCESS_RESTRICTED`, with a reason and expiry. They are not interpreted as
public 404 defects or successful reachability checks. Anonymous public probes
remain enabled. Revisit these classifications on visibility changes or expiry;
authenticated issue/settings verification above remains separate evidence.

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
K-PRICING identity, private visibility conditions, selected labels and product
contracts. Review schema/version changes explicitly; do not rerun initialization
to overwrite the project. Run all retained CI and initializer fixtures, review
the final diff, then record the new adopted baseline and reconciled deviations.

## Completion boundary

v0.0.1 is a **milestone-only setup checkpoint**. Keep `VERSION` at `0.0.0` and
changes under Unreleased; no tag or release publication is authorized by the
milestone name. Issue 10 can close after the exact starter evidence and all setup
issues are accepted. Migration follows [MIGRATION_PLAN.md](MIGRATION_PLAN.md).
No migrated numerical result, supported package or functional release is claimed.
