# Initialization status

This page records the historical initialization on 2026-09-23, when the
repository was private. It is the authority for initialization provenance, not
live settings. [Setup verification](SETUP_VERIFICATION.md) owns the current
visibility and controls; [setup completion](SETUP_COMPLETION.md) records the
accepted starter run and closed v0.0.1 checkpoint. Migration is tracked in
[MIGRATION_PLAN.md](MIGRATION_PLAN.md).

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
At initialization, the generic facade, core, tests and example were unchanged.
Migration replaces them together with their API and regression contracts.

## Project identity and future migration

**K-PRICING** is the product and repository name. **`KPR_`** is the VBA
module/function namespace of the migrated date layer. The identity policy
permits that namespace while retaining all other donor and template identity
restrictions and the existing scan scope.
The checker rejects standalone old-product headings and explicit old product-name
declarations; documentation tests also bind the README title to the recorded
project name. Historical references remain legitimate, not blanket scan exclusions.

The initialization record preserves the original inputs, including the original
absence of exact-source Excel evidence. The README limitations evolved after
the accepted starter run; the original input is intentionally not rewritten. Re-running the
initializer with those same inputs returns a no-op and preserves the evolved
identity policy; different inputs are rejected. Future template updates must
preserve this project-specific policy and pass its regression tests. The
initialization checkpoint did not migrate source, tests, examples or API
policy; their current status belongs to the migration plan.

## Validation boundary

The pre-initialization self-test passed all three profile fixtures, including
12 mandatory-component/README-only rejection cases and three optional-component
absence cases. The initialized source passes the 21-rule portable repository
gate. These are repository checks, not VBA compilation or Excel execution.

Initialization itself did not execute Excel. The subsequent neutral-starter
run was accepted on its exact candidate; see [SETUP_COMPLETION.md](SETUP_COMPLETION.md).
Pricing implementation, workbook/add-in packaging and functional-release
certification remain outside this completed setup milestone.

## Settings and automation

The original private-repository observations and the subsequent public settings
read-back are retained in [SETUP_VERIFICATION.md](SETUP_VERIFICATION.md). Recheck
that authority whenever visibility or the account plan changes; do not change
visibility implicitly to satisfy a control.

The selected label policy has 20 labels and no domain overlay. Trusted label
reconciliation and read-back evidence belong to the
[initialization pull request](https://github.com/danielep71/K-PRICING/pull/1).

## Completed setup and next work

[SETUP_VERIFICATION.md](SETUP_VERIFICATION.md) records the observed repository
settings and accepted starter result. [SETUP_COMPLETION.md](SETUP_COMPLETION.md)
records the final setup baseline and closure of issues #3, #6 and #10.
Use [DEVELOPER_SETUP.md](DEVELOPER_SETUP.md) for ongoing development and
[EXCEL_SETUP_RUNBOOK.md](EXCEL_SETUP_RUNBOOK.md) for evidence capture on a new
candidate, together with the active validation issue's scope and commands.
The migration inventory and issue mapping are in [MIGRATION_PLAN.md](MIGRATION_PLAN.md).
v0.0.1 is a completed milestone-only checkpoint; VERSION remains 0.0.0 and no
release tag was created.
