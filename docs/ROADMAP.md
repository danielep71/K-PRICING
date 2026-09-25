# K-PRICING delivery roadmap

This is the authoritative forward delivery roadmap for K-PRICING. It records
the planned milestones, their boundaries, the issue dependency sequence and the
scope that moves between milestones. GitHub milestones and issues hold the
detailed acceptance criteria and live state; this document holds the plan they
implement.

Historical records stay where they are:
[migration completion](MIGRATION_COMPLETION.md) records the v0.0.2 acceptance
and [migration handover](MIGRATION_HANDOVER.md) records what was carried into
K-PRICING. Neither is a current plan.

## Status and authority

- The closed milestones v0.0.1 and v0.0.2 are delivered and listed for context;
  their records are the setup and migration completion documents.
- Every open or future milestone listed here is planned, not delivered. Its
  delivery is recorded only by its closed issues, retained evidence and, for a
  release, a published tag.
- The [date-layer contract](DATE_LAYER_CONTRACT.md) is normative for behavior;
  this roadmap only schedules work against it.
- Changing a milestone boundary, dependency or scope transition requires a
  reviewed update to this document together with the affected issues.
- `VERSION` stays `0.0.0` until a release candidate is assembled (#51).

## Milestones

| Milestone | Scope | Issues | Release |
| --- | --- | --- | --- |
| v0.0.1 | Repository setup | Closed | Setup record, not a functional release |
| v0.0.2 | Repository migration of the frozen date layer | Closed | Migration record, not a functional release |
| v0.0.3 — Test Hardening and Registration | Independent fixtures, durable runner and evidence schema, complete regression suites, Excel cross-oracle checks, MacroOptions registration | #38 (done, PR #54), #39–#42 | No release |
| v0.0.4 — Date Primitives | Demo builder, Ribbon and CommandBars, public/infrastructure classification, final static controls and drift monitoring, candidate assembly, exact-source certification | #46–#52 | First functional release: tag `v0.0.4` |
| v0.0.5 — Business Calendars and Holidays | Calendars, weekend masks and holiday sets in the reserved `KPR_Cal_*` namespace | Not yet created | Planned |
| v0.0.6 — Business-Day Arithmetic and Roll Conventions | Business-day arithmetic and roll conventions built on v0.0.5 calendars | Not yet created | Planned |

## v0.0.3 — Test Hardening and Registration

Test hardening comes first and registration second, as decided at v0.0.2
acceptance (#18).

```text
#38 fixture generator (done, PR #54)
  └─> #39 runner and evidence schema
        ├─> #40 regression suites
        └─> #41 Excel cross-oracle checks
#38, #39, #40, #41 ─> #42 MacroOptions registration
```

Boundary: test and evidence infrastructure plus registration. No change to the
supported 22-function calculation API or its behavior.

## v0.0.4 — Date Primitives

```text
#42 (v0.0.3) ─> #46 demo builder
#39, #42, #46 ─> #47 Ribbon, #48 CommandBars
#47, #48 ─> #49 public/infrastructure classification
#40, #41, #49 ─> #50 final static controls and drift monitoring
#50 ─> #51 candidate assembly
#51 ─> #52 exact-source certification and release
```

Boundary: the release certifies exactly the 22 supported `KPR_Dates_*`
functions in one scalar/array-capable surface, with multi-cell use claimed only
on dynamic-array Excel. Registration, UI, test and demo entry points are
unsupported infrastructure. #52 requires every v0.0.3 issue and every other
v0.0.4 issue to be complete.

## v0.0.5 — Business Calendars and Holidays

Planned. Introduces calendars, weekend masks and holiday sets in the
`KPR_Cal_*` namespace reserved by the date-layer contract. Its issues and
acceptance criteria are created when the milestone is activated.

## v0.0.6 — Business-Day Arithmetic and Roll Conventions

Planned. Adds business-day arithmetic and following, preceding, modified and
other roll conventions on top of the v0.0.5 calendars. Its issues and
acceptance criteria are created when the milestone is activated.

## Scope transitions

| Capability | Excluded from | Planned in |
| --- | --- | --- |
| Calendars, weekend masks, holiday sets | v0.0.4 | v0.0.5 |
| Business-day arithmetic and roll conventions | v0.0.4, v0.0.5 | v0.0.6 |
| `_Spill` twins or `KPR_Dates_Spill.bas` | All milestones | Not planned |
| Ctrl+Shift+Enter or legacy multi-cell compatibility | All milestones | Not planned |
| Committed Office binaries | All milestones | Release assets only |
| Broader pricing capabilities | v0.0.3–v0.0.6 | Not yet scheduled |
