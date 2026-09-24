# KPR migration handover register

This register preserves the engineering decisions, evidence boundary and unfinished
roadmap carried from [danielep71/KPR](https://github.com/danielep71/KPR) into
K-PRICING. It supports K-PRICING issue #16 and the v0.0.2 repository-migration
milestone.

The migration baseline is the frozen KPR source commit
`f26450d1fa7b11261162e901dedba062f21c99a7`. KPR remains the historical source
of issue discussion and source evidence. A mapping in this document does not
close, rewrite or transfer ownership of a KPR issue.

## Scope boundary

v0.0.2 migrates the implemented KPR date-layer candidate. It does not expand the
milestone into registration, independently generated fixtures, the full future
regression architecture, cross-oracles, the deterministic demo builder, Ribbon,
CommandBars, package injection or a functional product release.

K-PRICING is the active implementation destination after migration acceptance.
The source repository remains the rollback/reference baseline until the
maintainer makes a separate archival or redirect decision.

## Completed source implementation mapped to K-PRICING

| KPR issue | Source result | K-PRICING disposition |
| --- | --- | --- |
| [KPR#9](https://github.com/danielep71/KPR/issues/9) | Frozen 22-name date-layer behavioral contract: strict ISO parsing, supported date window, host/date-system policy, array semantics, 100,000-element cap, native-error taxonomy and pillar modes | Contract and source provenance frozen by K-PRICING #11. The migrated implementation/API/static controls are owned by #12/#13. Destination wording corrections remain explicit rather than rewriting KPR history. |
| [KPR#10](https://github.com/danielep71/KPR/issues/10) | Deterministic VBE export identity and component-name rules | VBE contract migrated by #11; generic VBE/API/static controls and project-specific KPR checks retained by #13. Actual Windows import/round-trip evidence remains K-PRICING #17. |
| [KPR#11](https://github.com/danielep71/KPR/issues/11) | Five-module production architecture | Four cores plus `KPR_DATES_DAYS` imported by K-PRICING #12. The regression harness is migrated separately as test infrastructure. Destination paths are recorded in migration provenance. |
| [KPR#12](https://github.com/danielep71/KPR/issues/12) | Strict scalar date/integer/control parsing and native-error propagation | Migrated by #12; parsing/window/error invariants are enforced by #13. |
| [KPR#13](https://github.com/danielep71/KPR/issues/13) | Worksheet/VBA date-system guard and volatile `HostDateSystem` | Migrated by #12; host-guard, volatility and no-workbook-fallback rules are enforced by #13. Destination worksheet execution is still a #17 runtime claim. |
| [KPR#14](https://github.com/danielep71/KPR/issues/14) | Pillar grammar and `NEAREST/FLOOR/CEILING` policy | Migrated by #12. K-PRICING #32 corrects an inherited oversized-valid-pillar range-classification defect and records it as an explicit destination divergence. |
| [KPR#15](https://github.com/danielep71/KPR/issues/15) | Complete element-correct 22-function date surface | Migrated by #12; exact names/signatures/ownership are enforced by destination public-API and KPR contract gates. |
| [KPR#16](https://github.com/danielep71/KPR/issues/16) | Array shape/broadcasting services | Migrated by #12; dependency, purity and capacity structure are guarded by #13. |
| [KPR#17](https://github.com/danielep71/KPR/issues/17) | Array-capable public surface plus focused source regression evidence | Implementation migrated by #12. The historical 557-check pure run and seven-check dynamic-array run are retained as source evidence only, not destination certification. K-PRICING #14 defines destination parity evidence and #17 owns actual Windows execution. |

## Destination-only corrections

The following are intentional K-PRICING divergences from the frozen source and
must never be back-attributed to KPR.

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
   component/aggregate overflow to `PILLAR_AGGREGATE_RANGE/#NUM!`. The source
   repository remains unchanged.

The exact file hashes and destination blob identities are maintained in
[MIGRATION_PROVENANCE.md](MIGRATION_PROVENANCE.md).

## Open KPR roadmap carried forward without completion claims

| KPR issue | Remaining source scope | Destination disposition |
| --- | --- | --- |
| [KPR#18](https://github.com/danielep71/KPR/issues/18) | MacroOptions manifest and registration/cleanup lifecycle | Not implemented by migration. Future registration delivery. |
| [KPR#19](https://github.com/danielep71/KPR/issues/19) | Independent Python fixture generator, canonical TSV and generated VBA fixture module | Not implemented. Current migrated tests are not claimed to be an independent fixture oracle. |
| [KPR#20](https://github.com/danielep71/KPR/issues/20) | Final durable regression interface and broad structured evidence schema | Partially overlapped by migration-specific K-PRICING evidence work; the full source issue is not complete. |
| [KPR#21](https://github.com/danielep71/KPR/issues/21) | Complete contract/shape/parity/error/state regression matrix | Focused migrated suites exist, but the complete source-roadmap matrix remains future work. |
| [KPR#22](https://github.com/danielep71/KPR/issues/22) | Native Excel cross-oracle module | Not implemented. |
| [KPR#23](https://github.com/danielep71/KPR/issues/23) | Deterministic date-demo builder | Not implemented. The destination carries only a minimal direct-VBA consumer example. |
| [KPR#24](https://github.com/danielep71/KPR/issues/24) | Ribbon callbacks, RibbonX and safe package injection | Not implemented. |
| [KPR#25](https://github.com/danielep71/KPR/issues/25) | Idempotent classic CommandBars lifecycle | Not implemented. |
| [KPR#26](https://github.com/danielep71/KPR/issues/26) | Final supported/infrastructure public-member classification | Partially satisfied for the migrated 22-function calculation API by `docs/PUBLIC_API.txt` and destination static checks. Future registration/UI/demo infrastructure remains absent, so KPR#26 is not complete. |
| [KPR#27](https://github.com/danielep71/KPR/issues/27) | Final architecture inventory checks and live milestone-register drift monitoring | Partially overlapped by K-PRICING repository/KPR static gates. Final fixture/oracle/UI inventory and live register workflow remain future work. |
| [KPR#28](https://github.com/danielep71/KPR/issues/28) | Functional candidate documentation, VERSION/CHANGELOG assembly | Not imported as a v0.0.2 functional-release commitment. K-PRICING v0.0.2 is a repository-migration milestone. |
| [KPR#29](https://github.com/danielep71/KPR/issues/29) | Full source-roadmap Windows certification and publication | Not completed by migration. K-PRICING #17 certifies only the implemented migrated candidate/parity boundary; future registration/UI/demo/oracle/package certification remains separate. |

The complete acceptance criteria for unfinished work remain in the linked KPR
issue bodies. This avoids creating a second mutable copy of their detailed
requirements.

## Preserved dependency chain

Future delivery retains the source dependency structure unless a reviewed
destination plan explicitly changes it:

- registration foundation: KPR#18;
- independent test hardening: KPR#19 → KPR#20 → KPR#21, with KPR#22 dependent
  on KPR#20;
- demo/UI delivery: KPR#23 depends on #18; KPR#24 and #25 depend on
  #18/#20/#23;
- final supported/infrastructure inventory: KPR#26 depends on #24/#25;
- final static/register controls: KPR#27 depends on #21/#22/#26;
- functional candidate assembly: KPR#28 depends on #27;
- full source-roadmap certification/publication: KPR#29 depends on #28.

These are unscheduled future delivery phases in K-PRICING. When a phase is
activated, create or map destination issues with an explicit milestone and
fully qualified KPR source links. Do not invent release commitments to clear
this register.

## Evidence disposition

Evidence is classified by where and when it was produced.

- Historical KPR regression results remain **source evidence** and do not
  certify K-PRICING.
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

## Development handover and rollback policy

- New implementation belongs in K-PRICING after migration acceptance.
- Use fully qualified source references such as `danielep71/KPR#14`; use
  K-PRICING issue numbers only for destination work.
- The frozen KPR commit
  `f26450d1fa7b11261162e901dedba062f21c99a7` is the rollback/reference
  baseline for this migration.
- Do not rewrite source history to make destination corrections appear native to
  KPR.
- No open KPR feature is pulled into v0.0.2 solely to empty the source backlog.
- KPR remains available through migration acceptance. Archiving, redirecting or
  changing its visibility is a separate maintainer action after K-PRICING #18;
  this register grants no such action.
- K-PRICING is public. The maintainer's #29 decision permits the already
  reviewed email/host evidence to remain public; it does not authorize
  publishing new sensitive material.

## Exit from migration handover

K-PRICING #16 closed on 2026-09-24 with its working register in the issue
body. This document is the durable repository copy of that register. It is
linked from the [documentation authority map](README.md), references the
merged #14 evidence protocol by path, and is the register that documentation
issue #15 consumes. K-PRICING #18 remains the final migration-acceptance gate.
