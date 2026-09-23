# Initialization status

Initialized on 2026-09-23 for `danielep71/K-PRICING`, using the application
profile. The repository is private. This is a development scaffold, not a
functional pricing release.

## Source provenance

- Original generated commit: `18b1713499d7ba8fe3accd598a817c42a7928f4d`.
- Upstream template revision: `cb1f75bc20102f3a39478cfa4dfea04bbea696ee`.
- Both revisions have the same source tree:
  `0beb2b89b292e643ff236177083679391395d7c8`.
- Adopted template contract: `1.2.0`; this is the contract version, not a claim
  that the source snapshot is the tagged release.
- Exact initialization inputs and upstream repository identity are recorded in
  [the initialization record](../.github/initialization.json).

The deterministic initializer was reviewed in dry-run mode before application.
The generic facade, core, tests and example remain unchanged so that migration
can replace them together with their API and regression contracts.

## Project identity and future migration

**K-PRICING** is the product and repository name. **`KPR_`** remains the VBA
module/function namespace for the planned migration from
[`danielep71/KPR`](https://github.com/danielep71/KPR). This preserves existing
caller names without treating the source repository as unrelated donor branding.
The identity policy permits these references while retaining all other donor
and template identity restrictions and the existing scan scope.

The initialization record preserves the original inputs. Re-running the
initializer with those same inputs returns a no-op and preserves the evolved
identity policy; different inputs are rejected. Future template updates must
preserve this project-specific policy and pass its regression tests. The neutral
starter remains in place until source, tests, examples and API policy can be
migrated together.

## Validation boundary

The pre-initialization self-test passed all three profile fixtures, including
12 mandatory-component/README-only rejection cases and three optional-component
absence cases. The initialized source passes the 21-rule portable repository
gate. These are repository checks, not VBA compilation or Excel execution.

No pricing implementation, workbook/add-in package, functional release or
exact-candidate Windows Excel certification is included in this initialization.

## Private automation

Public Scorecard publication is excluded while the repository is private.
Private CodeQL requires confirmed eligibility and explicit enablement; no
successful security scan is claimed. See
[Supply-chain assurance](SUPPLY_CHAIN_ASSURANCE.md).

## Live setup

Read back on 2026-09-23 through the authenticated maintainer account:

| Setting | Observed state |
| --- | --- |
| Visibility / default branch | Private / `main` |
| Description / topics | Project description; `excel`, `vba`, `application`, `financial-analytics`, `instrument-pricing` |
| Issues / template mode | Enabled / disabled |
| Wiki / Discussions / Projects | Disabled |
| Merge / squash / rebase | All enabled |
| Auto-merge | Disabled |
| Delete merged branches / suggest branch updates | Enabled / enabled |
| Branch and tag rulesets | Unavailable: GitHub API returned HTTP 403 requiring GitHub Pro for this private repository |
| Private vulnerability-reporting UI | Not offered in the repository security settings; the documented private email contact remains the reporting route |

Ruleset enforcement is an unresolved platform limitation. No protected-branch,
immutable-tag or mandatory-status-check enforcement is claimed. Development
uses reviewed pull requests and verifies `Repository integrity` before merge;
this process does not replace server-side enforcement. Resolve that limitation
before a functional release without changing visibility implicitly.

The selected label policy has 20 labels and no domain overlay. The trusted push
after integration owns live reconciliation; its run and exact read-back evidence
are retained with the initialization pull request. Pull-request label validation
does not itself prove live reconciliation.

The [post-creation checklist](POST_CREATION_CHECKLIST.md) remains the detailed
setup authority. CI and final integration evidence belong to the
[initialization pull request](https://github.com/danielep71/K-PRICING/pull/1).
