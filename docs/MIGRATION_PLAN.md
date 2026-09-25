# KPR migration plan

Prepared on 2026-09-23 for setup issue 9. Execution belongs to **v0.0.2 - Repo
Migration**, issues 11–18, plus destination issues #29 and #32 added during
execution. The plan itself transferred no runtime certification;
execution now imports the frozen candidate while preserving that evidence
boundary. Product identity remains K-PRICING; VBA names retain `KPR_`.

## Migration status

Status as of 2026-09-24. The frozen source inventory and 22 source signatures
were rechecked on that date. [Migration provenance](MIGRATION_PROVENANCE.md)
records the source hashes, license, destination adaptations and blob
identities; the [handover register](MIGRATION_HANDOVER.md) records the current
source-to-destination disposition.

| Issue | Outcome | Merged as |
| --- | --- | --- |
| #11 | [Date-layer contract](DATE_LAYER_CONTRACT.md), [VBE export contract](VBE_EXPORT.md) and source provenance imported | PR #27, `01f50be937be181e32a21b4e7b8acd1415c49bf1` |
| #12, #13 | Five production modules, focused regression harness, example, API manifest, specialist static rules and retained-evidence adapter | PR #28, `70ba37ac9f8df4049b44cabd6bc013559b8f618f` |
| #32 | Destination correction: oversized valid pillar quantities classify as `PILLAR_AGGREGATE_RANGE` | PR #33, `a8ffa458186c51ec28ee778d27a86d45ed9fad49` |
| #29 | Public visibility, live controls and documentation authorities reconciled | PR #30, `47aa4c181306a4668383631888d2267589ea5906` |
| #14 | [Regression/parity protocol](MIGRATION_REGRESSION.md), migration instrumentation and evidence validator | PR #31, `118555ddf5c2cd8f944688f02fc9813713bf30c2` |
| #16 | [Handover register](MIGRATION_HANDOVER.md) of source work history and carried-forward roadmap | Issue closed; register merged in PR #34, `184e4b6976a51a3b29e17460ba06f7e919e7fbc6` |
| #15 | Documentation, examples and destination status reconciliation | PR #35, `80e992df73c2a3209dc104045aeed14119f22534` |
| #17 | Exact-source Windows Excel compilation and source/destination parity on candidate `db98e506358ab74893ab28d52183fa2e216352b9`, one Windows 64-bit Excel host | Evidence in [`evidence/migration-2026-09-24`](../evidence/migration-2026-09-24/session.txt); both validators PASS; PR #37, `b0794e2dab1bab9ff0cc948cf1d5734863c6bece` |
| #18 | Migration acceptance and handover of the next delivery backlog | [Completion record](MIGRATION_COMPLETION.md); next scope in milestone v0.0.3 (#38–#42); current plan in the [roadmap](ROADMAP.md) |

#29 and #32 were added to the milestone during execution; they are
destination work, not source history.

## Frozen source and evidence boundary

Frozen source commit: `f26450d1fa7b11261162e901dedba062f21c99a7`.
Source tree: `013653658c9c9fb1e61e5958d2d4277ab8c44191`.
Always read/import this immutable revision; newer source requires a reviewed
baseline change and a new comparison.

The calculation implementation is complete at the frozen revision. The source
record reports 557 checks with zero failures and seven focused array checks at
`a750cd5a935529b807c43631e92cd9f1b15ee8b3`, on one dynamic-array Excel host.
Earlier issue-level logs cover scalar/shape and host behavior. These are
historical, narrower results at their named revisions, not certification of the
later selected source or the destination. The destination now contains the frozen migrated
date candidate; exact-source destination runtime evidence for one Windows
64-bit host is retained under #17.

## File inventory and intended disposition

The SHA column contains **Git blob object IDs**, not SHA-256 evidence digests.
Record new SHA-256 Git-byte digests when preparing host evidence.

| Source path | Git blob SHA | Intended destination/disposition |
| --- | --- | --- |
| `docs/DATE_LAYER_CONTRACT.md` | `6984e9354cd2a32986f6504ac7ae4f95b0a472e2` | docs/DATE_LAYER_CONTRACT.md |
| `docs/IMPLEMENTATION_PLAN.md` | `0355dc71cd59d9dab08e11d456172d922e7b38be` | Not imported; reconciled through the [handover register](MIGRATION_HANDOVER.md) (see below) |
| `docs/VBE_EXPORT.md` | `f37eba3de6f3feb28a0d686bbf68b5fea5ff044e` | docs/VBE_EXPORT.md |
| `src/modules/KPR_Core_Array.bas` | `a6c98c192dc2746a59a418bf2190a40f1c13c233` | src/core/KPR_Core_Array.bas |
| `src/modules/KPR_Core_Dates.bas` | `7e0929474e1ae1311df2081eb7a64fe126f1b688` | src/core/KPR_Core_Dates.bas |
| `src/modules/KPR_Core_Err.bas` | `a6ac71c00979dd9dd13ed01264ae5a375da14c28` | src/core/KPR_Core_Err.bas |
| `src/modules/KPR_Core_Parse.bas` | `ec08bbb0056cfe6e27ed1115f3d6431e08a37b98` | src/core/KPR_Core_Parse.bas |
| `src/modules/KPR_DATES_DAYS.bas` | `37e997107c4c030d69bfa9b16d8aae3d22c97fc7` | src/modules/KPR_DATES_DAYS.bas |
| `test/modules/KPR_REGRESSION_TESTS.bas` | `b339d4b932f390c143973fe5fcc106e79ffbcad1` | tests/modules/KPR_REGRESSION_TESTS.bas |
| `tools/check_repo.py` | `80a5d078d9bb8f7114132e3985b6997f80a95517` | tools/check_kpr_contract.py (adapt; keep generic checker) |

All remaining source files are governance, presentation or metadata:
`.editorconfig`, `.gitattributes`, `.github/ISSUE_TEMPLATE/bug_report.md`, `.github/ISSUE_TEMPLATE/config.yml`, `.github/ISSUE_TEMPLATE/feature_request.md`, `.github/PULL_REQUEST_TEMPLATE.md`, `.github/labels.json`, `.github/scripts/labels-sync.mjs`, `.github/workflows/labels-sync.yml`, `.github/workflows/static-checks.yml`, `.gitignore`, `CHANGELOG.md`, `CODE_OF_CONDUCT.md`, `CONTRIBUTING.md`, `INSTALLATION.md`, `LICENSE`, `README.md`, `RELEASING.md`, `SECURITY.md`, `VERSION`, `assets/social-preview.png`.

Compare these for relevant intent and license attribution; retain the destination's
identity, workflows, 20-label policy, editor/export controls and stricter template
gates. Rewrite user documentation for the destination. Do not transplant source
VERSION, tags, old milestone status or visual branding. No generated fixture
module, demo package, registration implementation, Ribbon or CommandBars source
exists in the selected tree; these are open plans, not missing migrated files.

## Public surface and behavioral constraints

The source contract declares exactly 22 supported names:

- `KPR_Dates_DayOfWeek`
- `KPR_Dates_DaysInMonth`
- `KPR_Dates_DaysInYear`
- `KPR_Dates_BeginOfMonth`
- `KPR_Dates_EndOfMonth`
- `KPR_Dates_BeginOfQuarter`
- `KPR_Dates_EndOfQuarter`
- `KPR_Dates_BeginOfYear`
- `KPR_Dates_EndOfYear`
- `KPR_Dates_IsMonthEnd`
- `KPR_Dates_IsQuarterEnd`
- `KPR_Dates_IsYearEnd`
- `KPR_Dates_IsLeapYear`
- `KPR_Dates_AddDays`
- `KPR_Dates_AddWeeks`
- `KPR_Dates_AddMonths`
- `KPR_Dates_AddYears`
- `KPR_Dates_NthWeekdayOfMonth`
- `KPR_Dates_LastWeekdayOfMonth`
- `KPR_Dates_PillarFromDates`
- `KPR_Dates_DateFromPillar`
- `KPR_Dates_HostDateSystem`

The first 21 accept value inputs with shape-preserving arrays and scalar
expansion; `HostDateSystem` remains scalar-only. Preserve Variant returns,
ByVal arguments, optional controls and their defaults. The authoritative
source contract governs strict ISO parsing, the inclusive Gregorian window
1900-03-01 through 9999-12-31, error propagation and precedence, 1900/1904
boundaries, pillar grammar/rounding, exact shape rules and the 100,000-element
cap. Preserve the singular `DateFromPillar` name, with no old plural alias or
spill twins. No calendar/business-day extension is part of parity.

## Coherent replacement and acceptance gates

1. **#11 — provenance/contracts:** import and review the frozen date and VBE
   export contracts, preserve amendments and
   attribution, and record source hashes. Destination migration milestone
   naming must not silently redefine source functional-release scope.
2. **#12 — implementation:** import the five production modules without
   semantic rewrites. Coordinate the starter replacement with test/example,
   profile component inventory, API manifest and evidence-policy changes in the
   same reviewed candidate. Delete obsolete starter artifacts only when their
   complete replacements are included; no knowingly broken intermediate main.
3. **#13 and #14 — static/API and regression adaptation:** #13 is complete in
   PR #28: the generic `tools/check_repo.py` remains authoritative for generic
   repository policy and `tools/check_kpr_contract.py` adds the migrated KPR
   rules and negative fixtures. #14 is governed by
   [MIGRATION_REGRESSION.md](MIGRATION_REGRESSION.md): preserve the existing
   focused suites, distinguish native/common runners from the destination-only
   evidence adapter, and define exact-source parity without claiming the
   unfinished independent-fixture, runner, matrix and cross-oracle regression
   architecture is already implemented.
4. **#15 and #16 — documentation/history:** publish accurate developer/import
   guidance and reconcile every source work item against the register below.
   Preserve original evidence revisions, errors/corrections,
   decisions and unresolved conditions. Record the actual destination PR and
   SHA for each migrated item, rather than treating a completed source item as
   an automatically closed destination task.
5. **#17 — fresh Excel parity:** freeze the destination SHA after review and
   green retained/specialist CI. Import into a fresh Windows Excel host,
   compile, run the complete migrated harness and compare the documented
   scalar, array, expected-error, cap and cleanup behavior. Exercise relevant
   1900/1904 host boundaries; record locale, build, bitness and dynamic-array
   capabilities. Preserve full logs/source hashes and run again in the same
   session to detect leaked state. Separate untested environments and future
   coverage from successful results.
6. **#18 — acceptance:** reconcile implementation, contracts, API, evidence,
   docs and issue register at the reviewed candidate. Record limitations and
   deferred product work. Do not publish a package or rename APIs as a side
   effect of migration.

If moving a module requires a format/export correction, record it separately
from behavior changes and verify the diff. If starter replacement spans review
PRs, prepare dependent branches and integrate the coherent set together after
all gates pass. Never replace generic governance with the older source checks.

## Source work register

This is the preparation-time register of the frozen source's work items, read on
2026-09-23; destination issue numbers refer to K-PRICING. The current
disposition, carried-forward roadmap and handover policy are maintained in
[MIGRATION_HANDOVER.md](MIGRATION_HANDOVER.md); where the two differ, the
handover register is authoritative. Completed source work is implementation
history. Open work remains open unless separately implemented and evidenced.

| Source work item | State and decision | Destination owner/trace |
| --- | --- | --- |
| Remove inherited implementation | Complete; historical setup only; keep destination baseline | #16 |
| Repository identity and links | Complete; reconcile naming; do not overwrite destination | #15, #16 |
| Editor, ignore and export policy | Complete; compare source export rules, retain stricter destination | #12, #16 |
| Labels | Complete; retain destination 20-label policy, not source catalog | #16 |
| Static checks | Complete; port specialist rules alongside generic CI | #13 |
| GitHub controls | Complete; source settings are not inherited or evidence for this destination repository | #16 |
| Visual identity | Complete; defer source image reuse; product branding must stay K-PRICING | #15 |
| Repository baseline release | Complete; historical release only, no tag/version transfer | #16 |
| Frozen date contract | Complete; preserve selected revision and amendments | #11 |
| VBE export contract | Complete; preserve guidance; source round trip remains uncertified | #11, #17 |
| Layered architecture | Complete; preserve module identities with mapped paths | #12 |
| Strict parsing | Complete; migrate implemented behavior unchanged | #12, #14 |
| Host and date system | Complete; retain contract; historical focused host evidence only | #12, #17 |
| Pillar grammar and modes | Complete; preserve supported grammar/errors | #12, #14 |
| 22 public functions | Complete; preserve exact signatures and supported surface | #12, #13 |
| Array shape engine | Complete; preserve shapes, caps and failure rules | #12, #14 |
| Array facade integration | Complete; preserve implementation; repeat exact-candidate parity | #12, #17 |
| Function registration | Open; defer new MacroOptions category/argument registration | #16; now #42 |
| Independent generated fixtures | Open; defer generator/TSV/generated-module implementation | #16; now #38 |
| Full runner and evidence | Open; preserve current harness first; final runner interface still future | #14, #16; now #39 |
| Complete regression matrix | Open; preserve implemented tests; broader coverage remains future | #14, #16; now #40 |
| Native Excel cross-oracle | Open; defer independent overlap checks, not a current certification | #16; now #41 |
| Deterministic demo builder | Open; no workbook builder to import | #16 |
| Ribbon integration | Open; no XML/callback/package implementation to import | #16 |
| CommandBars lifecycle | Open; no install/remove implementation to import | #16 |
| API/classification manifest | Open; migration records implemented subset; future UI/register surface deferred | #13, #16 |
| Final static/live issue register | Open; port existing checks; future full inventory/live monitor deferred | #13, #16 |
| Final documentation/version/candidate | Open; rewrite destination docs; do not claim release completion | #15, #16 |
| Exact-source certification/release | Open; historical runs do not satisfy a final release | #17, #18 |

Issue 16 owned detailed reconciliation and successor traceability, not feature
implementation. A successor feature issue is created only with an explicit
future milestone; no unassigned backlog items are created by this inventory.
Daniele Penza owns selecting follow-on functional scope.

Preserve the source's recorded corrections as history. In particular, the
layered-architecture work records an interim non-atomic case/name change and its
correction; do not rewrite that history as continuously green. The VBE export
contract is implemented documentation, not proof of an actual VBE source round
trip. The host/date-system and array-façade records contain focused host
results with narrower scope than full certification. The final documentation
work item explicitly retains unverified final-state conditions.

## Source implementation plan disposition

The source `docs/IMPLEMENTATION_PLAN.md` at the frozen revision is the source's
own roadmap and milestone register. It is not copied into K-PRICING. Its
completed work is mapped to destination issues, and its unfinished work, with
its dependency chain, is carried forward without completion claims in the
[handover register](MIGRATION_HANDOVER.md). K-PRICING's own roadmap is its
milestones and issues; future phases get destination issues with an explicit
milestone only when they are scheduled.

## Exit from preparation

Preparation ended with issue #9: this inventory and executable sequence. The
[migration status](#migration-status) table above records execution. See
[developer setup](DEVELOPER_SETUP.md) and
[setup verification](SETUP_VERIFICATION.md) for the destination baseline.

## Test-module visibility decision

For issue #14, preserve the source harness exception: `KPR_REGRESSION_TESTS`
omits `Option Private Module` to keep its existing regression entry points
callable. The starter's private test module is not the migrated convention.
[House style](VBA_HOUSE_STYLE.md) owns the rule; the supported production API
remains the 22 entries in [PUBLIC_API.txt](PUBLIC_API.txt).
