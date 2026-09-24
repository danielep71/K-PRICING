# KPR migration plan

Prepared on 2026-09-23 for setup issue 9. Execution belongs to **v0.0.2 - Repo
Migration**, issues 11–18. The plan itself transferred no runtime certification;
execution now imports the frozen candidate while preserving that evidence
boundary. Product identity remains K-PRICING; VBA names retain `KPR_`.

## Contract migration started

The frozen source inventory and 22 source signatures were rechecked on
2026-09-24. [Migration provenance](MIGRATION_PROVENANCE.md) records the
source hashes, licence and editorial adaptations for issue #11. The
[date-layer contract](DATE_LAYER_CONTRACT.md) and
[VBE export contract](VBE_EXPORT.md) are imported. PR #28 merged the five
production modules, focused regression harness, API manifest, specialist static
rules and retained-evidence adapter as `70ba37ac9f8df4049b44cabd6bc013559b8f618f`.
The [regression/parity protocol](MIGRATION_REGRESSION.md) owns issue #14's
execution and comparison method; destination Excel parity remains pending in
#17.

## Frozen source and evidence boundary

Source repository: [danielep71/KPR](https://github.com/danielep71/KPR).
Selected commit: `f26450d1fa7b11261162e901dedba062f21c99a7`.
Source tree: `013653658c9c9fb1e61e5958d2d4277ab8c44191`.
Always read/import this immutable revision; newer source requires a reviewed
baseline change and a new comparison.

Calculation implementation through source issue 17 is present. Source issue 17
records 557 checks with zero failures and seven focused array checks at
`a750cd5a935529b807c43631e92cd9f1b15ee8b3`, on one dynamic-array Excel host.
Earlier issue-level logs cover scalar/shape and host behavior. These are
historical, narrower results at their named revisions, not certification of the
later selected source or the destination. Source final release issue 29 remains
a source-repository obligation. The destination now contains the frozen migrated
date candidate; exact-source destination runtime evidence remains pending.

## File inventory and intended disposition

The SHA column contains **Git blob object IDs**, not SHA-256 evidence digests.
Record new SHA-256 Git-byte digests when preparing host evidence.

| Source path | Git blob SHA | Intended destination/disposition |
| --- | --- | --- |
| `docs/DATE_LAYER_CONTRACT.md` | `6984e9354cd2a32986f6504ac7ae4f95b0a472e2` | docs/DATE_LAYER_CONTRACT.md |
| `docs/IMPLEMENTATION_PLAN.md` | `0355dc71cd59d9dab08e11d456172d922e7b38be` | Reconcile into this migration plan and issue register |
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
   export contracts, qualify source issue links, preserve amendments and
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
   unfinished source #19-#22 regression architecture is already implemented.
4. **#15 and #16 — documentation/history:** publish accurate developer/import
   guidance and reconcile every source issue/comment against the mapping below.
   Preserve qualified links, original evidence revisions, errors/corrections,
   decisions and unresolved conditions. Record the actual destination PR and
   SHA for each migrated item, rather than treating a closed source issue as
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
   deferred product work. Keep KPR accessible; do not archive or delete it,
   publish a package or rename APIs as a side effect of migration.

If moving a module requires a format/export correction, record it separately
from behavior changes and verify the diff. If starter replacement spans review
PRs, prepare dependent branches and integrate the coherent set together after
all gates pass. Never replace generic governance with the older source checks.

## Source issue register

States were read from the source on 2026-09-23; destination issue numbers below
refer to K-PRICING. Closed source work is implementation history. Open work
remains open unless separately implemented and evidenced.

| Source | Subject | State and decision | Destination owner/trace |
| --- | --- | --- | --- |
| [KPR#1](https://github.com/danielep71/KPR/issues/1) | Remove inherited implementation | Closed; historical setup only; keep destination baseline | #16 |
| [KPR#2](https://github.com/danielep71/KPR/issues/2) | Repository identity and links | Closed; reconcile naming and qualified links; do not overwrite destination | #15, #16 |
| [KPR#3](https://github.com/danielep71/KPR/issues/3) | Editor, ignore and export policy | Closed; compare source export rules, retain stricter destination | #12, #16 |
| [KPR#4](https://github.com/danielep71/KPR/issues/4) | Labels | Closed; retain destination 20-label policy, not source catalogue | #16 |
| [KPR#5](https://github.com/danielep71/KPR/issues/5) | Static checks | Closed; port specialist rules alongside generic CI | #13 |
| [KPR#6](https://github.com/danielep71/KPR/issues/6) | GitHub controls | Closed; source settings are not inherited or evidence for this private repo | #16 |
| [KPR#7](https://github.com/danielep71/KPR/issues/7) | Visual identity | Closed; defer source image reuse; product branding must stay K-PRICING | #15 |
| [KPR#8](https://github.com/danielep71/KPR/issues/8) | Repository baseline release | Closed; historical release only, no tag/version transfer | #16 |
| [KPR#9](https://github.com/danielep71/KPR/issues/9) | Frozen date contract | Closed; preserve selected revision and amendments | #11 |
| [KPR#10](https://github.com/danielep71/KPR/issues/10) | VBE export contract | Closed; preserve guidance; source round trip remains uncertified | #11, #17 |
| [KPR#11](https://github.com/danielep71/KPR/issues/11) | Layered architecture | Closed; preserve module identities with mapped paths | #12 |
| [KPR#12](https://github.com/danielep71/KPR/issues/12) | Strict parsing | Closed; migrate implemented behavior unchanged | #12, #14 |
| [KPR#13](https://github.com/danielep71/KPR/issues/13) | Host and date system | Closed; retain contract; historical focused host evidence only | #12, #17 |
| [KPR#14](https://github.com/danielep71/KPR/issues/14) | Pillar grammar and modes | Closed; preserve supported grammar/errors | #12, #14 |
| [KPR#15](https://github.com/danielep71/KPR/issues/15) | 22 public functions | Closed; preserve exact signatures and supported surface | #12, #13 |
| [KPR#16](https://github.com/danielep71/KPR/issues/16) | Array shape engine | Closed; preserve shapes, caps and failure rules | #12, #14 |
| [KPR#17](https://github.com/danielep71/KPR/issues/17) | Array facade integration | Closed; preserve implementation; repeat exact-candidate parity | #12, #17 |
| [KPR#18](https://github.com/danielep71/KPR/issues/18) | Function registration | Open; defer new MacroOptions category/argument registration | Source #18; track in #16 |
| [KPR#19](https://github.com/danielep71/KPR/issues/19) | Independent generated fixtures | Open; defer generator/TSV/generated-module implementation | Source #19; track in #16 |
| [KPR#20](https://github.com/danielep71/KPR/issues/20) | Full runner and evidence | Open; preserve current harness first; final runner interface still future | Source #20; #14, #16 |
| [KPR#21](https://github.com/danielep71/KPR/issues/21) | Complete regression matrix | Open; preserve implemented tests; broader coverage remains future | Source #21; #14, #16 |
| [KPR#22](https://github.com/danielep71/KPR/issues/22) | Native Excel cross-oracle | Open; defer independent overlap checks, not a current certification | Source #22; track in #16 |
| [KPR#23](https://github.com/danielep71/KPR/issues/23) | Deterministic demo builder | Open; no workbook builder to import | Source #23; track in #16 |
| [KPR#24](https://github.com/danielep71/KPR/issues/24) | Ribbon integration | Open; no XML/callback/package implementation to import | Source #24; track in #16 |
| [KPR#25](https://github.com/danielep71/KPR/issues/25) | CommandBars lifecycle | Open; no install/remove implementation to import | Source #25; track in #16 |
| [KPR#26](https://github.com/danielep71/KPR/issues/26) | API/classification manifest | Open; migration records implemented subset; future UI/register surface deferred | #13, #16; source #26 |
| [KPR#27](https://github.com/danielep71/KPR/issues/27) | Final static/live issue register | Open; port existing checks; future full inventory/live monitor deferred | #13, #16; source #27 |
| [KPR#28](https://github.com/danielep71/KPR/issues/28) | Final documentation/version/candidate | Open; rewrite destination docs, do not claim source release completion | #15, #16; source #28 |
| [KPR#29](https://github.com/danielep71/KPR/issues/29) | Exact-source certification/release | Open; historical runs do not satisfy final source or destination release | #17, #18; source #29 |

Issue 16 owns detailed reconciliation and successor traceability, not feature
implementation. Before creating any successor feature issue, assign an explicit
future milestone; no unassigned backlog items are created by this inventory.
Daniele Penza owns selecting that follow-on functional scope. Existing source
issues remain the trace for deferred work until then.

Preserve source comments as evidence, including corrections. In particular,
source issue 11 records an interim non-atomic case/name change and its correction;
do not rewrite that history as continuously green. Source issue 10's export
contract is implemented documentation, not proof of an actual VBE source round
trip. Issues 13 and 17 contain focused host results with narrower scope than
the open full certification work. Source issue 28 explicitly retains unverified
final-state conditions.

## Exit from preparation

The migration backlog already exists in milestone v0.0.2: issues 11–18. Setup
preparation ends with this inventory and executable sequence; it does not close
those execution issues. See [developer setup](DEVELOPER_SETUP.md) and
[setup verification](SETUP_VERIFICATION.md) for the destination baseline.
