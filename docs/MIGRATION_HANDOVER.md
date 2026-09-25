# KPR migration handover register

This register preserves the engineering decisions, evidence boundary and unfinished
roadmap carried into K-PRICING with the frozen date-layer source. It supports
K-PRICING issue #16 and the v0.0.2 repository-migration milestone.

The migration baseline is the frozen KPR source commit
`f26450d1fa7b11261162e901dedba062f21c99a7`. K-PRICING is the only maintained
home of this work; this register records what was carried in, not a second
backlog. Current milestones and sequencing are owned by the
[roadmap](ROADMAP.md).

## Scope boundary

v0.0.2 migrates the implemented KPR date-layer candidate. It does not expand the
milestone into registration, independently generated fixtures, the full future
regression architecture, cross-oracles, the deterministic demo builder, Ribbon,
CommandBars, package injection or a functional product release.

K-PRICING is the active implementation destination after migration acceptance.

## Completed source implementation mapped to K-PRICING

| Source capability | Source result | K-PRICING disposition |
| --- | --- | --- |
| Behavioral contract | Frozen 22-name date-layer behavioral contract: strict ISO parsing, supported date window, host/date-system policy, array semantics, 100,000-element cap, native-error taxonomy and pillar modes | Contract and source provenance frozen by K-PRICING #11. The migrated implementation/API/static controls are owned by #12/#13. Destination wording corrections remain explicit rather than rewriting source history. |
| VBE export identity | Deterministic VBE export identity and component-name rules | VBE contract migrated by #11; generic VBE/API/static controls and project-specific KPR checks retained by #13. Windows import and compile evidence for one 64-bit host is retained under K-PRICING #17; no VBE export round trip was performed. |
| Layered architecture | Five-module production architecture | Four cores plus `KPR_DATES_DAYS` imported by K-PRICING #12. The regression harness is migrated separately as test infrastructure. Destination paths are recorded in migration provenance. |
| Strict parsing | Strict scalar date/integer/control parsing and native-error propagation | Migrated by #12; parsing/window/error invariants are enforced by #13. |
| Host and date system | Worksheet/VBA date-system guard and volatile `HostDateSystem` | Migrated by #12; host-guard, volatility and no-workbook-fallback rules are enforced by #13. Destination 1900/1904 worksheet-caller behavior was observed on one Windows 64-bit host in K-PRICING #17. |
| Pillar grammar | Pillar grammar and `NEAREST/FLOOR/CEILING` policy | Migrated by #12. K-PRICING #32 corrects an inherited oversized-valid-pillar range-classification defect and records it as an explicit destination divergence. |
| Public surface | Complete element-correct 22-function date surface | Migrated by #12; exact names/signatures/ownership are enforced by destination public-API and KPR contract gates. |
| Array engine | Array shape/broadcasting services | Migrated by #12; dependency, purity and capacity structure are guarded by #13. |
| Array façade | Array-capable public surface plus focused source regression evidence | Implementation migrated by #12. The historical 557-check pure run and seven-check dynamic-array run are retained as source evidence only, not destination certification. K-PRICING #14 defines destination parity evidence and #17 owns actual Windows execution. |

## Destination-only corrections

The following are intentional K-PRICING divergences from the frozen source and
must never be back-attributed to it.

1. **Integer-range contract wording.** The frozen implementation tests the VBA
   `Long` range before integrality, so an out-of-range fractional numeric is
   `INTEGER_RANGE/#NUM!`; `INTEGER_FRACTION` applies only to in-range
   fractional numerics. K-PRICING clarifies the contract text to match the
   frozen implementation.
2. **Regression-harness 32-bit portability.** The two `CLngLng` test cases are
   guarded with `Win64` rather than the frozen source's `VBA7` condition.
   Production behavior is unchanged.
3. **Retained-evidence instrumentation.** K-PRICING adds destination test
   adapters required by its evidence contract: `KPR_Tests_RunEvidence` for the
   retained host record and, from #14, `KPR_Tests_RunMigrationEvidence` with
   serialized `OBS`/`CLEANUP` records from the stateful runners. They add no
   assertions and do not become supported calculation API.
4. **Pillar range classification (#32).** K-PRICING validates the entire pillar
   grammar before numeric conversion and maps grammatically valid numeric
   component/aggregate overflow to `PILLAR_AGGREGATE_RANGE/#NUM!`. The frozen
   source is unchanged.

The exact file hashes and destination blob identities are maintained in
[MIGRATION_PROVENANCE.md](MIGRATION_PROVENANCE.md).

## Roadmap carried forward without completion claims

This section is historical: it records what migration carried forward and where
each item was scheduled. Current planning lives in the [roadmap](ROADMAP.md).

| Roadmap item | Remaining scope | Destination disposition |
| --- | --- | --- |
| Registration | MacroOptions manifest and registration/cleanup lifecycle | Not implemented by migration. Scheduled in v0.0.3 as #42. |
| Independent fixtures | Independent Python fixture generator, canonical TSV and generated VBA fixture module | Not implemented. Current migrated tests are not claimed to be an independent fixture oracle. Scheduled in v0.0.3 as #38. |
| Durable runner | Final durable regression interface and broad structured evidence schema | Partially overlapped by migration-specific K-PRICING evidence work; not complete. Scheduled in v0.0.3 as #39. |
| Regression matrix | Complete contract/shape/parity/error/state regression matrix | Focused migrated suites exist, but the complete matrix remains future work. Scheduled in v0.0.3 as #40. |
| Cross-oracle | Native Excel cross-oracle module | Not implemented. Scheduled in v0.0.3 as #41. |
| Demo builder | Deterministic date-demo builder | Not implemented. The destination carries only a minimal direct-VBA consumer example. Scheduled in v0.0.4 as #46. |
| Ribbon | Ribbon callbacks, RibbonX and safe package injection | Not implemented. Scheduled in v0.0.4 as #47. |
| CommandBars | Idempotent classic CommandBars lifecycle | Not implemented. Scheduled in v0.0.4 as #48. |
| Member classification | Final supported/infrastructure public-member classification | Partially satisfied for the migrated 22-function calculation API by `docs/PUBLIC_API.txt` and destination static checks. Future registration/UI/demo infrastructure remains absent, so the item is not complete. Scheduled in v0.0.4 as #49. |
| Inventory controls | Final architecture inventory checks and live milestone-register drift monitoring | Partially overlapped by K-PRICING repository/KPR static gates. Final fixture/oracle/UI inventory and live register workflow remain future work. Scheduled in v0.0.4 as #50. |
| Candidate assembly | Functional candidate documentation, VERSION/CHANGELOG assembly | Not a v0.0.2 functional-release commitment. K-PRICING v0.0.2 is a repository-migration milestone. Scheduled in v0.0.4 as #51. |
| Certification | Full roadmap Windows certification and publication | Not completed by migration. K-PRICING #17 certifies only the implemented migrated candidate/parity boundary; future registration/UI/demo/oracle/package certification remains separate. Scheduled in v0.0.4 as #52. |

Every carried-forward item is now scheduled. Detailed acceptance criteria are
recorded in the destination issues, and the [roadmap](ROADMAP.md) is the
authority for current milestones and sequencing.

## Preserved dependency chain

The carried-forward dependency structure maps onto destination issues as
follows. The [roadmap](ROADMAP.md) owns any later change to it.

- registration foundation: registration (#42), after all of test hardening
  (#38–#41);
- independent test hardening: independent fixtures (#38) → durable runner
  (#39) → regression matrix (#40), with cross-oracle (#41) dependent on the
  durable runner (#39);
- demo/UI delivery: the demo builder (#46) depends on registration (#42);
  Ribbon (#47) and CommandBars (#48) depend on registration (#42), the durable
  runner (#39) and the demo builder (#46);
- member classification (#49) depends on Ribbon (#47) and CommandBars (#48);
- inventory controls (#50) depend on the regression matrix (#40), cross-oracle
  (#41) and member classification (#49);
- candidate assembly (#51) depends on inventory controls (#50);
- certification (#52) depends on candidate assembly (#51).

At v0.0.2 acceptance (#18) the maintainer scheduled independent test hardening
(#38–#41) followed by registration (#42) in milestone v0.0.3. The remaining
items were later scheduled in milestone v0.0.4 — Date Primitives as #46–#52,
ending in the first functional release. Calendars (v0.0.5) and business-day
arithmetic (v0.0.6) are new destination scope recorded only in the
[roadmap](ROADMAP.md).

## Evidence disposition

Evidence is classified by where and when it was produced.

- Historical KPR regression results remain **source evidence** and do
  not certify K-PRICING.
- Hosted K-PRICING static checks prove only their documented static properties.
- Migration regression/parity evidence follows K-PRICING #14 and
  [MIGRATION_REGRESSION.md](MIGRATION_REGRESSION.md), merged in PR #31.
- Exact Windows Excel import, compile and source/destination parity are owned by
  K-PRICING #17.
- Packaging, registration and UI certification are outside the migrated
  candidate and therefore cannot be inferred from #17.

Any source/destination difference must be either a documented migration
adaptation with its own evidence or a blocking defect; matching the source is
not itself proof of correctness.

## Development handover policy

- New implementation belongs in K-PRICING.
- The frozen source commit `f26450d1fa7b11261162e901dedba062f21c99a7` is the
  reference baseline for this migration; its identity is recorded by commit and
  blob IDs in [MIGRATION_PROVENANCE.md](MIGRATION_PROVENANCE.md).
- Do not present destination corrections as frozen-source behavior.
- No roadmap item is pulled into v0.0.2 solely to empty the backlog.
- K-PRICING is public. The maintainer's K-PRICING #29 decision permits the already
  reviewed email/host evidence to remain public; it does not authorize
  publishing new sensitive material.

## Exit from migration handover

K-PRICING #16 closed on 2026-09-24 with its working register in the issue
body. This document is the durable repository copy of that register. It is
linked from the [documentation authority map](README.md), references the
merged #14 evidence protocol by path, and is the register that documentation
issue #15 consumed. K-PRICING #18 accepted the migration.
