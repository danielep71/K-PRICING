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

GitHub settings and labels are separate from file initialization. Their final
read-back status is recorded here after the initialization change is integrated.
The [post-creation checklist](POST_CREATION_CHECKLIST.md) remains the detailed
setup authority.
