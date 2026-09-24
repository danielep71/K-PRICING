# v0.0.2 migration completion

Repository migration **v0.0.2** accepts the frozen KPR date-layer candidate into
K-PRICING. This is a milestone record, not a published release: `VERSION`
remains `0.0.0`, no tag or package is created, and no production-ready pricing
product is claimed. Issue #18 owns this acceptance.

## Accepted candidate

| Item | Value |
| --- | --- |
| Frozen source | danielep71/KPR `f26450d1fa7b11261162e901dedba062f21c99a7` |
| Runtime-tested destination candidate | `db98e506358ab74893ab28d52183fa2e216352b9` |
| Main at acceptance review | `b0794e2dab1bab9ff0cc948cf1d5734863c6bece` (PR #37) |

Between the tested candidate and the acceptance review, no file under `src/`,
`tests/` or `examples/modules/` changed; only documentation and the retained
evidence were added. The #17 runtime evidence therefore applies to the
accepted production and test source.

## Inventory and deltas

All six migrated components match the destination blob IDs recorded in
[MIGRATION_PROVENANCE.md](MIGRATION_PROVENANCE.md):

| Component | Relationship to the frozen source |
| --- | --- |
| `KPR_Core_Err`, `KPR_Core_Parse`, `KPR_Core_Array` | Byte-for-byte imports |
| `KPR_Core_Dates` | #32 correction: oversized valid pillar quantities return `PILLAR_AGGREGATE_RANGE` (`#NUM!`) |
| `KPR_DATES_DAYS` | Format-only: seven trailing spaces removed |
| `KPR_REGRESSION_TESTS` | Three destination-only test adaptations: `Win64` guard, `KPR_Tests_RunEvidence`, and the migration observation instrumentation |

`KPR_DateExample` is a destination-only example. The supported calculation API
is exactly the 22 functions in [PUBLIC_API.txt](PUBLIC_API.txt).

## Evidence

- **Runtime parity (#17):** [`evidence/migration-2026-09-24`](../evidence/migration-2026-09-24/session.txt).
  - The frozen source compiled and passed its own 557 checks.
  - The destination passed 12 cases and 568 assertions with 0 failures.
  - Source and destination migration observation streams were equal (31
    observations, cleanup PASS).
  - #32 is registered as the known difference.
  - Both `check_excel_evidence.py` and `check_migration_evidence.py` pass at
    the tested candidate.
- **Hosted CI at acceptance review (`b0794e2`):**
  [Static checks 36053281729](https://github.com/danielep71/K-PRICING/actions/runs/36053281729),
  [CodeQL 36053281678](https://github.com/danielep71/K-PRICING/actions/runs/36053281678),
  [Scorecard 36053281663](https://github.com/danielep71/K-PRICING/actions/runs/36053281663)
  and [label sync 36053281642](https://github.com/danielep71/K-PRICING/actions/runs/36053281642)
  passed.
- **Execution history:** the [migration plan](MIGRATION_PLAN.md#migration-status)
  records each issue with its merge commit.

## Limitations

- Runtime parity was verified on one Windows 64-bit host only: Microsoft 365
  Excel Version 2608, Build 16.0.20326.20072, Italian (Italy) regional format.
  32-bit Office, other builds and other locales are untested.
- The #32 source-side probe result is maintainer-reported and not
  validator-bound; the destination correction is covered by the bound
  568-assertion harness.
- No VBE export round trip was performed.
- The worksheet examples in `examples/README.md` are contract expectations; the
  #17 run did not execute those exact formulas.
- Registration, generated fixtures, demo/UI, packaging, application lifecycle
  and pricing capabilities are not part of v0.0.2.

## Metadata audit

On 2026-09-24 all 37 issues and pull requests (#1–#37) had a milestone, exactly
one P1/P2/P3 priority label and assignee `danielep71`. The successor issues
#38–#42 were created with the same metadata.

## Next delivery scope

The maintainer selected independent test hardening first, then registration,
in milestone **v0.0.3 - Test Hardening and Registration**:

| Issue | Scope | Source |
| --- | --- | --- |
| #38 | Independent fixture generator and generated VBA fixture module | danielep71/KPR#19 |
| #39 | Durable regression runner and structured evidence schema | danielep71/KPR#20 |
| #40 | Complete contract/shape/parity/error/state regression matrix | danielep71/KPR#21 |
| #41 | Native Excel cross-oracle checks | danielep71/KPR#22 |
| #42 | MacroOptions registration and register/unregister lifecycle | danielep71/KPR#18 |

danielep71/KPR#23–#29 (demo, Ribbon, CommandBars, final inventory, candidate
assembly and certification) remain unscheduled in the
[handover register](MIGRATION_HANDOVER.md).

## Source repository

danielep71/KPR remains the historical reference and recovery source. Its issues
are not closed or rewritten by this migration, and archiving, redirecting or
changing its visibility is a separate maintainer decision.
