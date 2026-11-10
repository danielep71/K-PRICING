# 🗂️ Milestone register: v0.0.4 — Date primitives

This register mirrors the GitHub issues of milestone **v0.0.4** so that the
release plan is reviewable in the repository and drift between the plan and
the live issues is visible (#50). The region between the two marker comments
is generated: never edit it by hand.

```bash
python3 tools/milestone_register.py --write   # regenerate from the live milestone, then commit
python3 tools/milestone_register.py --check   # CURRENT (0), DRIFT (1) or UNVERIFIED (2)
```

The `Milestone register` workflow (`.github/workflows/milestone-register.yml`)
runs `--check` on relevant pushes to `main`, on issue events and on demand. It
reads live GitHub state, so it stays outside the deterministic static gate and
its result is not a required status; a fresh `CURRENT` is required at candidate
assembly (#51) and immediately before tagging ([RELEASING.md](../RELEASING.md)).
`UNVERIFIED` means GitHub could not be read, never drift. Nothing here claims
that VBA imports, compiles or runs in Excel.

Issue bodies are kept as written, with line ends normalized and trailing blanks
removed.

<!-- BEGIN GENERATED MILESTONE REGISTER: tools/milestone_register.py --write -->
Milestone **v0.0.4 — Date primitives**: 9 issues, 2 open, 7 closed.

### #46 Add the deterministic date-demo builder

| Field | Value |
| --- | --- |
| State | closed |
| Assignees | danielep71 |
| Labels | documentation, enhancement, P2, tests |
| Milestone | v0.0.4 — Date primitives |
| URL | https://github.com/danielep71/K-PRICING/issues/46 |

````text
## Objective

Provide a reproducible demonstration workbook for the single scalar/array-capable KPR date surface without committing an Office binary or making unsupported claims.

## Deliverable

Add a shared deterministic demo builder under `examples/modules/`, with `KPR_Demo_Dates.bas` exposing the Date Primitives demo. The builder may add a new demo sheet to a suitable active workbook or create a new workbook, and file output is written only to an explicit `.xlsx` path.

The tracked builder is authoritative. A generated workbook may be attached later as a release asset after exact-source certification.

## Acceptance criteria

- [x] The builder creates the same sheets, labels, formulas, examples, widths, styles and named areas on repeated runs.
- [x] Examples use one supported surface: the 21 value-taking functions keep the same names for scalar and multi-cell calls, `HostDateSystem` is shown as scalar-only, and no `_Spill` function is shown.
- [x] Scalar examples are clearly distinguished from dynamic-array-only multi-cell examples.
- [x] Examples cover native-error categories, representative array shapes, the 100,000-element policy and documented pillar behavior; the capacity limit is explained without materializing an excessive 100,000-cell demo spill.
- [x] The workbook labels calendars and business-day calculations as future scope.
- [x] No example claims CSE compatibility, certified accuracy, production readiness or compatibility beyond the agreed evidence.
- [x] The builder preserves Excel state and never overwrites an existing path without a controlled error.
- [x] No `.xlsm`, `.xlam`, `.xlsx` or other generated Office binary is committed.
- [x] The #52 exact-source certification scope explicitly includes deterministic demo generation.

## Dependencies

- [x] #42 — function registration

The migrated date implementation, regression/cross-oracle hardening, MacroOptions registration lifecycle and one-host parity baseline were complete before implementation. #46 is now complete, unblocking #47 and #48.

## Scope boundary

Demo infrastructure only; no new supported calculation API.

## Completion record — 2026-10-06

Implemented by PR #75 and squash-merged as `abf0be9e5afb2940a47cf1c730ebb5d94f87246e`. The final PR head was `3ab71222d5c16cddf8bbbc9da3d270a1a2c779ed`; the Windows Excel runtime evidence was produced on `cd0e1a1767f98367e9c4c708867c638a50d70090`, with the later commit changing documentation only.

The implementation uses the repository-authorized `examples/modules/` location rather than creating a parallel `demo/` tree. `KPR_Demo_Builder.bas` provides the shared deterministic builder and `KPR_Demo_Dates.bas` defines the Date Primitives demo. Two full Windows Excel runs passed with **21 suites, 9,320 assertions, 0 failures**, state restored, and `demo_generation = PASS`; the focused `worksheet-demo` suite passed 44 assertions. Repeated builds were fingerprinted for deterministic content, existing output files are refused, generated Office binaries remain untracked, and #52 retains deterministic demo generation as a certification gate.

One intermittent status-bar read-back issue observed in the durable runner is tracked separately as #77; it is not a demo-builder failure and does not invalidate the passing #46 evidence.


## Current naming — 2026-10-26

PR #79 subsequently renamed the shared builder to `examples/modules/KPR_Demo__Builder.bas`; `KPR_Demo_Dates.bas` remains the Date Primitives module. The completion record above retains the names and source identities actually used for #75. Current source still contains the demo and determinism suite; final-candidate Excel certification remains in #52.


## Schedule

- Start date: 2026-09-25
- Target date: 2026-10-06

Retrospective dates recorded on 2026-11-03: Start date uses the issue creation date and Target date uses the recorded closure date (Europe/Rome). These are historical metadata proxies, not evidence of the actual work start or an original planned deadline.
````

### #47 Add Ribbon callbacks, RibbonX and safe package injection

| Field | Value |
| --- | --- |
| State | closed |
| Assignees | danielep71 |
| Labels | enhancement, P2, repository, tests |
| Milestone | v0.0.4 — Date primitives |
| URL | https://github.com/danielep71/K-PRICING/issues/47 |

````text
## Objective

Add a modern Excel Ribbon surface whose callbacks delegate to registration, demo and test infrastructure without becoming supported calculation API.

## Deliverables
- `src/modules/KPR_UI_Ribbon.bas`
- `src/ribbon/customUI14.xml`
- `tools/inject_ribbon.py`, a deterministic package-injection helper that reads and writes only explicit untracked workbook/add-in paths

## Ribbon contract
Use the Office 2010+ `customUI/2009/07` schema, stable KPR IDs, built-in `imageMso` icons and exact callback signatures. The release surface provides registration and demo actions; any regression action is development-only.

## Validation boundary
This issue owns deterministic XML/package tests and production callback/injection implementation. #50 later folds RibbonX/callback consistency into repository-wide static inventory; #52 owns exact-candidate Windows load and exercise.

## Acceptance criteria
- [x] Callback names/signatures in RibbonX and VBA match exactly.
- [x] `onLoad` caches the Ribbon object safely and invalidation is controlled.
- [x] Repeated loads/actions create no duplicate controls or registration.
- [x] Callbacks contain no date algorithm and delegate to the owning module.
- [x] Package injection validates ZIP structure, content types and relationships before changing an untracked target.
- [x] Repeated injection replaces the KPR customUI part idempotently and creates no duplicate relationships.
- [x] Failed injection leaves the original target recoverable and produces no committed binary.
- [x] Focused deterministic tests cover valid, repeated and failed injection without mutating source artifacts.
- [x] Injection uses explicit paths and atomic replacement or recoverable backup; it never discovers or rewrites a tracked Office binary.
- [x] #52 exact-source certification includes RibbonX loading and exercise.

## Dependencies
- [x] #42
- [x] #39
- [x] #46

## Scope boundary
Excel UI/package infrastructure only; no change to the supported 22-function calculation API.

## Current status — 2026-10-06

All listed dependencies are complete. This issue is ready for implementation and may proceed in parallel with #48. The durable-runner status-bar follow-up (#77 / PR #78) is also complete, so UI lifecycle tests can use the current hardened caller-state restoration path.


## Current status — 2026-11-09 (closed)

Complete. Implementation: #96/#97 and the one-click check #101, #106–#111. Windows Excel evidence at `ba250a7` (Microsoft 365 16.0.20430, 64-bit): two runs of 23 suites, 9,372 assertions, 0 failures, `worksheet-ribbon` PASS and the K-PRICING tab loaded in the built workbook; both evidence records validated by `check_test_evidence.py`. Evidence: https://github.com/danielep71/K-PRICING/issues/47#issuecomment-6751840736

## Schedule

- Start date: 2026-11-03
- Target date: 2026-11-10

Planned on 2026-11-03. Dates follow the roadmap dependency sequence and precede the v0.0.4 milestone due date of 2026-11-30. Completion remains subject to the existing acceptance criteria and, where applicable, Windows Excel certification.


---
_Generated by [Claude Code](https://claude.ai/code)_
````

### #48 Add idempotent classic CommandBars integration

| Field | Value |
| --- | --- |
| State | closed |
| Assignees | danielep71 |
| Labels | enhancement, P2, tests |
| Milestone | v0.0.4 — Date primitives |
| URL | https://github.com/danielep71/K-PRICING/issues/48 |

````text
## Objective
Provide a dedicated classic Excel toolbar/menu lifecycle for environments and users that rely on CommandBars.

## Lifecycle
Expose technically public infrastructure entry points to install and remove a temporary KPR Dates bar/menu. Use a stable unique name and tags, remove stale KPR-owned controls before building, and bind actions through fully qualified `OnAction` targets. The release bar provides registration and demo commands; regression commands are development-only.

## Acceptance criteria
- [x] Install/build and teardown entry points are safe to call repeatedly in any order.
- [x] A second install produces no duplicate bar, menu or control.
- [x] Teardown removes only KPR-owned controls and tolerates already-absent UI.
- [x] Temporary controls do not persist uncontrolled Excel state after the host closes.
- [x] Callbacks delegate to registration, demo or test modules and contain no date calculations.
- [x] Failures restore alerts, events, status bar, active workbook/sheet and selection.
- [x] #52 exact-source certification covers install, repeat install, action dispatch, teardown and repeat teardown.

## Dependencies
- [x] #42
- [x] #39
- [x] #46

## Scope boundary
Classic Excel UI infrastructure only; no change to the supported 22-function calculation API.

## Current status — 2026-10-06

All listed dependencies are complete. This issue is ready for implementation and may proceed in parallel with #47. The durable-runner status-bar follow-up (#77 / PR #78) is also complete, so UI lifecycle tests can use the current hardened caller-state restoration path.


## Current status — 2026-11-09 (closed)

Complete. Implementation: #96/#97 and the one-click check #101, #106–#111. Windows Excel evidence at `ba250a7` (Microsoft 365 16.0.20430, 64-bit): two runs of 23 suites, 9,372 assertions, 0 failures, `worksheet-commandbars` PASS (37 assertions) in both runs; both evidence records validated by `check_test_evidence.py`. Evidence: https://github.com/danielep71/K-PRICING/issues/48#issuecomment-6751841078

## Schedule

- Start date: 2026-11-03
- Target date: 2026-11-10

Planned on 2026-11-03. Dates follow the roadmap dependency sequence and precede the v0.0.4 milestone due date of 2026-11-30. Completion remains subject to the existing acceptance criteria and, where applicable, Windows Excel certification.


---
_Generated by [Claude Code](https://claude.ai/code)_
````

### #49 Freeze the public-API manifest and classify infrastructure members

| Field | Value |
| --- | --- |
| State | closed |
| Assignees | danielep71 |
| Labels | documentation, P2, repository |
| Milestone | v0.0.4 — Date primitives |
| URL | https://github.com/danielep71/K-PRICING/issues/49 |

````text
## Objective
Create one auditable inventory that separates the supported single public surface from internal and host-required procedures.

## Deliverable
Maintain a machine-checkable public-API manifest and readable documentation covering exactly 22 supported `KPR_Dates_*` functions; vectorization classification; internal `KPR_Core_*` procedures; MacroOptions registration entry points; Ribbon/CommandBars callbacks; regression-runner and demo-builder entry points.

Reserve `KPR_Cal_*` without creating placeholder modules or functions. Do not reserve or document a `KPR_Dates_*_Spill` surface.

## Acceptance criteria
- [x] Exactly 22 functions are classified as supported v0.0.4 API.
- [x] Every technically public infrastructure/internal member is explicitly marked unsupported.
- [x] Names, signatures, vectorization and argument ordering agree with `docs/PUBLIC_API.txt`, `docs/DATE_LAYER_CONTRACT.md` and the #42 MacroOptions manifest.
- [x] No supported function is implemented in a core, UI, registration, test or demo module.
- [x] No unclassified public procedure remains.
- [x] No `_Spill` name or `KPR_Dates_Spill.bas` appears.
- [x] The manifest is suitable for static completeness and drift checks.
- [x] Future calendar/business-day namespaces and milestone boundaries are recorded.

## Dependencies
- [x] #47
- [x] #48

## Current status — 2026-11-09 (closed)

Complete through #112 (`50320b9`): `docs/API_INVENTORY.txt` and the `kpr-api-inventory` rule. Evidence: https://github.com/danielep71/K-PRICING/issues/49#issuecomment-6752693930

## Schedule

- Start date: 2026-11-11
- Target date: 2026-11-13

Planned on 2026-11-03. Dates follow the roadmap dependency sequence and precede the v0.0.4 milestone due date of 2026-11-30. Completion remains subject to the existing acceptance criteria and, where applicable, Windows Excel certification.



---
_Generated by [Claude Code](https://claude.ai/code)_
````

### #50 Extend deterministic static checks and add milestone-register drift monitoring

| Field | Value |
| --- | --- |
| State | closed |
| Assignees | danielep71 |
| Labels | ci, P2, repository, tests |
| Milestone | v0.0.4 — Date primitives |
| URL | https://github.com/danielep71/K-PRICING/issues/50 |

````text
## Objective
Add two deliberately separate controls: extend the deterministic repository-integrity gate for the final single-surface KPR date architecture, and add a live milestone-register drift monitor whose result is explicitly external-state dependent. Neither control may claim that VBA imports, compiles or executes in Excel.

## Track A — deterministic tree checks
- Require all production, test, fixture, RibbonX, demo, contract and supporting tooling files for the v0.0.4 candidate.
- Enforce VBE export-name fidelity, case-insensitive component uniqueness, `Option Explicit`, private-core visibility and facade ownership.
- Keep module inventory and public/core/UI/test boundaries aligned with the approved architecture.
- Require the public-API manifest, #42 MacroOptions manifest and source declarations to contain the same 22 supported names/signatures.
- Resolve every RibbonX callback to a classified callback procedure.
- Parse RibbonX and evidence schemas successfully.
- Require #38 fixture generation `--check` to report no drift and classify generated source correctly.
- Require the regression harness, generated fixtures and cross-oracle infrastructure without obsolete duplicate runner/assertion/suite modules.
- Unit-test the milestone-register renderer against a frozen issue snapshot.
- Reject generated Office binaries, test-results output, `_Spill` API and calendar/business-day production placeholders.
- Preserve all existing repository checks.

The deterministic static workflow must not call the GitHub Issues API or depend on mutable live issue state.

## Track B — live milestone-register drift monitor
Maintain a canonical renderer for the v0.0.4 milestone register.

### Renderer contract
- Render milestone v0.0.4 issues in ascending issue-number order.
- Include number, title, state, assignees, labels, milestone, URL and complete body.
- Replace only a stable bounded generated region.
- `--write` fetches live milestone state and rewrites only that region.
- `--check` returns `CURRENT` (0), `DRIFT` (1), or `UNVERIFIED` (2).
- Drift output names the first mismatched issue/field and regeneration command.
- Authentication, rate-limit, network, API and malformed-response failures are `UNVERIFIED`, never `DRIFT`.

### Separate workflow
- Keep the live fetch outside the deterministic static gate.
- Trigger on relevant pushes to `main`, relevant issue events and `workflow_dispatch`; not on `pull_request`.
- Use `contents: read` and `issues: read`.
- Keep the live workflow out of the deterministic required-status gate, but require a fresh `CURRENT` result for candidate assembly and tagging.
- Pin external actions immutably.

## Acceptance criteria
- [x] Deterministic positive checks pass on the complete candidate tree without network access.
- [x] Focused negative self-tests fail for every new deterministic rule with actionable messages.
- [x] Removing the authoritative roadmap document, `docs/ROADMAP.md`, fails the required-file rule.
- [x] Frozen-snapshot tests prove byte-identical renderer output and bounded replacement.
- [x] The deterministic workflow never calls the live GitHub Issues API.
- [x] A verified live issue edit produces `DRIFT` until regeneration.
- [x] API/authentication/network/malformed-response cases produce `UNVERIFIED`, not `DRIFT`.
- [x] The live workflow is separate, read-only, pinned and absent from `pull_request` triggers.
- [x] Relevant push and issue events compare against the default-branch plan.
- [x] Workflow summaries distinguish `CURRENT`, `DRIFT` and `UNVERIFIED`.
- [x] A fresh `CURRENT` result is required at candidate assembly and immediately before tagging.
- [x] Static checks and drift monitoring make no Excel-runtime claim.
- [x] Deterministic reports include all new tree-rule results.
- [x] Required-file policy does not require future `KPR_Cal*` modules.
- [x] Final tree checks preserve export fidelity, component uniqueness and facade-role ownership.

## Dependencies
- [x] #40
- [x] #41
- [x] #49

## Schedule

- Start date: 2026-11-16
- Target date: 2026-11-20

Planned on 2026-11-03. Dates follow the roadmap dependency sequence and precede the v0.0.4 milestone due date of 2026-11-30. Completion remains subject to the existing acceptance criteria and, where applicable, Windows Excel certification.



---
_Generated by [Claude Code](https://claude.ai/code)_
````

### #51 Reconcile documentation, VERSION and CHANGELOG and assemble the candidate

| Field | Value |
| --- | --- |
| State | open |
| Assignees | danielep71 |
| Labels | documentation, P2, release |
| Milestone | v0.0.4 — Date primitives |
| URL | https://github.com/danielep71/K-PRICING/issues/51 |

````text
## Objective
Commit the complete v0.0.4 date-primitives release-candidate documentation and metadata before the exact-source Windows run.

## Scope
- Update README status, installation/compatibility guidance, contract links, public API and demo instructions.
- Document exactly 22 public names in one scalar/array-capable surface with no `_Spill` twins.
- Claim multi-cell use only on dynamic-array Excel; make no CSE compatibility claim.
- Document the supported date window, strict ISO parsing, optional-argument rule, 100,000-element cap and native-error categories.
- Document worksheet-Range date-system handling and the direct-VBA 1900 caller contract without active-workbook fallback.
- State that business calendars/holidays are v0.0.5 scope and business-day arithmetic/roll conventions are v0.0.6 scope.
- Preserve v0.0.1 as repository setup and v0.0.2 as repository migration; neither is a functional date/business-day release.
- Reconcile `CHANGELOG.md` and prepare the v0.0.4 entry.
- Set `VERSION` to `0.0.4` only as part of the complete candidate.
- Document fixture/demo regeneration and certification.
- Keep generated Office binaries and final run output untracked.

## Acceptance criteria
- [ ] Documentation and manifests describe exactly the implemented 22-name surface.
- [ ] No page documents a spill twin or unsupported legacy multi-cell behavior.
- [ ] No page claims calendar or business-day support in v0.0.4.
- [ ] Historical descriptions of v0.0.1 and v0.0.2 remain accurate.
- [ ] `VERSION`, changelog and candidate scope agree.
- [ ] All repository and Markdown checks pass at the candidate commit.
- [ ] The candidate SHA is recorded for #52.
- [ ] No tracked change is planned after certification except a rerun-triggering correction.

## Dependencies
- [x] #50

## Schedule

- Start date: 2026-11-23
- Target date: 2026-11-24

Planned on 2026-11-03. Dates follow the roadmap dependency sequence and precede the v0.0.4 milestone due date of 2026-11-30. Completion remains subject to the existing acceptance criteria and, where applicable, Windows Excel certification.
````

### #52 Run exact-source Windows certification and publish v0.0.4

| Field | Value |
| --- | --- |
| State | open |
| Assignees | danielep71 |
| Labels | blocked, ci, P2, release, repository, tests |
| Milestone | v0.0.4 — Date primitives |
| URL | https://github.com/danielep71/K-PRICING/issues/52 |

````text
## Objective
Certify the exact v0.0.4 date-primitives candidate in real Windows Excel and publish only if every gate passes.

## Required certification
- Start from a clean checkout of the recorded candidate SHA.
- Import every production, test and demo VBE export into supported Windows Excel and compile with no error/missing reference.
- Run direct-VBA scalar cases under the documented 1900 contract.
- Exercise worksheet `Range` calls in 1900 and 1904 workbooks, including `KPR_Dates_HostDateSystem()` and ordinary recalculation refresh.
- On dynamic-array Excel, run complete shape, broadcasting, optional-argument, per-element error and capacity suites; make no CSE claim.
- Run the full deterministic regression harness and cross-oracle suite.
- Run MacroOptions registration/cleanup twice and record residual Excel-owned metadata honestly.
- Load/exercise RibbonX and CommandBars lifecycle.
- Generate the demo deterministically.
- Export VBA source again and compare the normalized round trip.
- Record Excel/Windows environment, exact SHA, caller contexts, case totals and failures using the committed evidence schema.
- Keep run records outside tracked source and attach them to this issue/release evidence.

## Caller-context probe
Record `Application.Caller` type/value and `HostDateSystem()` behavior for worksheet `Range`, direct VBA/Immediate/`Application.Run`/regression harness, macro-attached shape, and other controlled Excel host contexts where reproducible. Untestable contexts must be reported as not testable, not silently omitted.

## Exit gate
- [x] Every issue in milestone v0.0.3 is complete.
- [ ] Every other issue in milestone v0.0.4 is complete.
- [ ] Hosted static checks pass at the exact candidate SHA.
- [ ] Fresh import and VBA compilation pass in Windows Excel.
- [ ] All deterministic regression and allowed cross-oracle cases pass.
- [ ] Scalar/array parity, shapes, capacity, native errors and caller-state restoration pass.
- [ ] 1900 worksheet behavior, 1904 refusal for value-taking functions and `HostDateSystem()` behavior pass.
- [ ] Caller-context probe results and untestable contexts are attached.
- [ ] Release documentation states unsupported non-Range host contexts explicitly.
- [ ] MacroOptions, Ribbon and CommandBars lifecycles are idempotent.
- [ ] Demo and source round-trip evidence are attached.
- [ ] Compatibility claims distinguish scalar support from dynamic-array-only multi-cell support.
- [ ] Calendars and business-day arithmetic are explicitly out of v0.0.4 scope.
- [ ] No tracked file changes after the run; any change requires a new exact-source run.
- [ ] Create protected tag `v0.0.4` and GitHub pre-release only after evidence is complete.
- [ ] Do not rewrite or move prior tags.
- [ ] Generated Office artifacts are release assets only, never committed.

## Dependencies
- [ ] #51

## Loader validation carried forward — 2026-10-27

PR #80 merged as `d29d7f3e51151fff19705b3d11d8da6d34b5302f` without a newly recorded Excel execution result in its discussion. Retain its specific validation scope alongside final-candidate certification:

- [ ] Import `KPR_Dev_Import` into `PERSONAL.XLSB` and use it to load the exact recorded candidate into the intended target workbook.
- [ ] Compile the target VBA project and run `KPR_Test_RunAll` twice; retain both results and environment/source identity.
- [ ] Confirm role-filtered loading preserves unselected modules and retirement removes only repository-owned modules absent from the selected source.

Passing hosted CI, local static checks and Python tests do not complete these Excel checks. Earlier runtime evidence remains attributed to its original SHA.


## Schedule

- Start date: 2026-11-25
- Target date: 2026-11-27

Planned on 2026-11-03. Dates follow the roadmap dependency sequence and precede the v0.0.4 milestone due date of 2026-11-30. Completion remains subject to the existing acceptance criteria and, where applicable, Windows Excel certification.
````

### #77 Durable runner sometimes reads the status bar back as the text FALSE

| Field | Value |
| --- | --- |
| State | closed |
| Assignees | danielep71 |
| Labels | bug, P2, tests |
| Milestone | v0.0.4 — Date primitives |
| URL | https://github.com/danielep71/K-PRICING/issues/77 |

````text
### Affected repository profile

application (development/regression harness only; no supported surface)

### Defect summary and impact

The durable runner (`KPR_Test_RunAll` / `KPR_Test_RunSuite` in `tests/modules/KPR_REGRESSION_TESTS.bas`) sometimes fails its caller-state check after a run whose suites all passed:

```text
FAIL  state : not restored: status bar (expected Excel control, read 'FALSE')
```

Excel reads the status bar back as the **text** `FALSE` instead of handing it back to Excel (a Boolean `False`). The run is then reported as `RESULT=FAIL` and needs repeating, so the failure costs Excel certification runs.

### Exact source identity

- `cd0e1a1767f98367e9c4c708867c638a50d70090` (PR #75): run-17 failed this way; run-18 and run-19 on the same commit passed.
- The same failure happened in run-9 during #65. `aa4aa0d` (hand back a literal `False` and retry three times after `DoEvents`) made it rarer but did not remove it.

### Minimal reproduction

Intermittent. In both cases it happened on the first full run of a session, right after other macros had run (the demo build in #75; registration runs in #65). The next run passed.

### Expected behavior

After every run the status bar is back under Excel's control (`VarType(Application.StatusBar) = vbBoolean`), as it was before the run.

### Observed behavior

`Application.StatusBar` still reads as the String `"FALSE"` after the hand-back and its retries.

### Environment

Windows, Excel 16.0 build 20326, 64-bit, VBA7, DMY locale with `,` as decimal separator.

### Evidence, mitigation, and recovery

The root cause remains unconfirmed because the intermittent failure did not recur during the #78 verification runs. PR #78 implemented the diagnostic and recovery work without requiring Excel 4.0 macros:

1. The runner records the status bar's type and value after every hand-back step and includes that trace in the state-restoration detail.
2. If the literal `False` plus retries still leaves text, the runner sets a different text and hands the bar back again, then hides/shows the status bar while preserving the caller's `DisplayStatusBar` setting and retries the hand-back.
3. The state check remains strict: a text `FALSE` is still a failure and is never accepted as restored.

Two full Windows Excel runs on the #78-tested source passed with 21 suites, 9,320 assertions, 0 failures and state restored. If the intermittent failure recurs, reopen this issue with the new full hand-back trace.

### Submission checks

- [x] I searched open and closed issues and did not find an equivalent report.
- [x] I reproduced this against the exact source identity stated above.
- [x] This report contains no vulnerability details, secrets, private links, or restricted data.


---
_Generated by [Claude Code](https://claude.ai/code)_

### Completion record — 2026-10-06

Addressed by PR #78, squash-merged as `7927be0b0f23fbb9bf8f26e99f3a7180e801ca23`. The durable-runner source verified in Windows Excel was `732db137b21d9fc8d0a6ebf27bf761cc3c178bec`; the final PR head retained the same runner blob. Two matching full runs passed. The non-XLM recovery design replaced the initially proposed `MESSAGE(FALSE)` fallback after review, so no Trust Center weakening is required.


## Schedule

- Start date: 2026-10-06
- Target date: 2026-10-06

Retrospective dates recorded on 2026-11-03: Start date uses the issue creation date and Target date uses the recorded closure date (Europe/Rome). These are historical metadata proxies, not evidence of the actual work start or an original planned deadline.
````

### #87 DateFromPillar records RESULT_WINDOW where the contract requires PILLAR_AGGREGATE_RANGE

| Field | Value |
| --- | --- |
| State | closed |
| Assignees | danielep71 |
| Labels | bug, P3 |
| Milestone | v0.0.4 — Date primitives |
| URL | https://github.com/danielep71/K-PRICING/issues/87 |

````text
### Affected repository profile

Date library (src/)

### Defect summary and impact

`KPR_Dates_DateFromPillar` records the wrong internal condition when a valid pillar pushes the result outside the supported window. `Elem_DateFromPillar` in `src/modules/KPR_DATES_DAYS.bas` returns `ErrForCondition(KPR_COND_RESULT_WINDOW)` in two places:
- when the month shift leaves the window (`TryAddMonths` fails);
- when the day shift leaves the window.

`docs/DATE_LAYER_CONTRACT.md` assigns `PILLAR_AGGREGATE_RANGE` to "a valid token whose aggregate or resulting date falls outside the numerical or supported date domain" (section 3.4 and the condition registry in section 7). The generated fixtures expect it too, for example `pillar-parse.edge-9999-12-31-1-d` and `matrix.datefrompillar-startdate-edge-high`.

Impact: none on worksheet or VBA results. Both conditions map to `#NUM!`, and the element functions are private, so no caller and no test can tell them apart. The code simply disagrees with its contract and fixtures.

### Exact source identity

`main` @ `206a9636589c479033aebcee290d824d63ee8ff0`.

### Minimal reproduction

Static: read `Elem_DateFromPillar`. `=KPR_Dates_DateFromPillar("9999-12-31","1D")` returns `#NUM!` (correct value), but internally through `RESULT_WINDOW` rather than `PILLAR_AGGREGATE_RANGE`.

### Expected behavior

Both out-of-window paths use `PILLAR_AGGREGATE_RANGE`, as the contract and fixtures specify.

### Observed behavior

Both use `RESULT_WINDOW`.

### Environment

Static source finding from a comment audit; no Excel run needed to see it.

### Evidence, workaround, and recovery

Found during the repository comment audit (PR #86). Fix: change the two condition codes. No workaround needed, because values are unaffected.

### Submission checks

- [x] I searched open and closed issues and did not find an equivalent report.
- [x] I reproduced this against the exact source identity stated above.
- [x] This report contains no vulnerability details, secrets, private links, or restricted data.


---
_Generated by [Claude Code](https://claude.ai/code)_

### Closure verification — 2026-10-26

Implemented by #88, merged as `b52ec1d5fb9a21ec368338cc1a27eaa5524187d1`. Both the failed month shift and the out-of-window day result in `Elem_DateFromPillar` now use `KPR_COND_PILLAR_AGGREGATE_RANGE`; this remains present on main at `a9ecbaf3d79a9da9af2c08ed5fa852e1e2c33cfc`. This is verified source/contract alignment, not Excel execution. The PR still records compilation and `KPR_Test_RunAll` as not run; final exact-source Windows certification remains required under #52.


## Schedule

- Start date: 2026-10-21
- Target date: 2026-10-21

Retrospective dates recorded on 2026-11-03: Start date uses the issue creation date and Target date uses the recorded closure date (Europe/Rome). These are historical metadata proxies, not evidence of the actual work start or an original planned deadline.
````

<!-- END GENERATED MILESTONE REGISTER -->
