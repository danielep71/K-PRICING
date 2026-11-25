# ✅ Validation plan

This plan says how each [calculation contract](CALCULATION_CONTRACTS.md) is
shown to hold, and what evidence a result needs before anyone relies on it.

## Responsibilities

Implementing a calculation, checking it independently and accepting it are
separate steps, recorded separately even when one person does all three.
Code review and passing repository checks are necessary, but they are not
evidence that a figure is right: the static checks never run Excel or a
calculation.

## Reference values

An expected result is never taken from the workbook. Two kinds of value
validate a rule:

- **Published:** printed in a registered [source](SOURCES.md), such as a
  worked example of the regulation.
- **Independent:** computed by an external implementation by another author,
  registered in the source register with an `IMPL-` ID.

A value the project computes itself, by hand, in a separate spreadsheet that
shares no formulas with the workbook, or with a Python script under
`tests/derivations/`, is an `illustrative` cross-check: it catches defects,
but it validates no rule. Record how each value was derived and who derived
it. A reference script is a test tool, never a runtime dependency of the
workbook.

Each comparison has an absolute and a relative tolerance, with the reason for
them: the source's own rounding, or the precision of `Double` arithmetic.

The [numerical test cases](TEST_CASES.md) record each expected value with its
reference class (`published`, `independent` or `illustrative`), its source,
its derivation and its tolerance
([reference classes](TEST_CASES.md#reference-classes)). Only `published` and
`independent` values validate a rule: a rule reaches *Validated* in the
[traceability table](#traceability) only through one of them.

## Validation matrix

| Calculation | Reference values | Evidence | Status |
| --- | --- | --- | --- |
| *Illustrative:* `CALC-01` | `illustrative` cross-checks only: two `nominal` cases, a JSON fixture and a file-based case, whose total of two synthetic notional amounts a derivation script computes; empty input and invalid input rejected are still to cover; no `published` or `independent` value yet | Evidence record per release candidate | Not implemented |

The *Status* of a calculation is the least advanced
[traceability](#traceability) status of its rules: *Validated* only when every
rule is.

Cover, for each calculation, with the case categories of the
[test-case format](TEST_CASES.md#categories):

- typical cases (`nominal`) and the boundaries of every rule (`boundary`);
- invalid and missing input, which the [data contract](DATA_CONTRACT.md)
  rejects (`invalid`).

Cover with VBA tests, which add no category
([relation, end-to-end and volume tests](TEST_CASES.md#other-tests)):

- failure: an operation that fails clears its earlier results and marks them
  `FAILED` ([status and staleness](WORKBOOK_CONTRACT.md#status-and-staleness));
- repeatability: the same input gives the same output.

A relation test or the end-to-end case runner of the
[harness guide](TEST_HARNESS.md) can check either.

## Traceability

Each rule is traced from the paragraph of its source to the step of the
calculation contract that states it, to the VBA procedure that implements it
and to the test cases that test it, with its validation status:

| Source | Paragraph | Calculation step | VBA procedure | Test cases | Status |
| --- | --- | --- | --- | --- | --- |
| *Illustrative:* `SRC-01` | Article 1(1) | `CALC-01.1` | — | `illustrative-format-example`, `illustrative-file-import-example` | Not implemented |

- **Source and paragraph:** a `SRC-` ID of the [source register](SOURCES.md)
  and the exact paragraph, as in the step's *Source* column.
- **Calculation step:** `CALC-NN.<step>`, a step of the *Rules* table of
  that [calculation contract](CALCULATION_CONTRACTS.md). A rule implemented
  in several steps has one row per step.
- **VBA procedure:** `Module.Procedure` of a component under `src/`, or `—`
  while none implements the rule. Name every procedure that applies the rule.
- **Test cases:** the IDs of the [test cases](TEST_CASES.md) that test the
  step, or `—`. Each case tests the row's calculation.
- **Status:** one of the four below.

| Status | Meaning | The row names |
| --- | --- | --- |
| Validated | A `published` or `independent` case passes for the rule on an Excel run of the release candidate ([completeness policy](TEST_CASES.md#completeness-policy)) | A procedure, and a case with a `published` or `independent` output whose `rule` is the row's step |
| Exercised | Cases run the rule, but none is a `published` or `independent` case that passed: they are `illustrative`, or the reference cases have not passed yet | A procedure and at least one case |
| Not validated | Implemented, but no case tests it | A procedure and no case |
| Not implemented | No procedure implements it yet; its cases may come first | No procedure |

`python tools/check_test_cases.py`, which **Repository integrity** runs,
fails a row whose source is not in the register, whose step is not one that
the *Rules* table of a declared calculation lists, whose procedure is not in
a component under `src/`, whose case does not exist or tests another
calculation, or whose status contradicts what the row names. It cannot see an Excel run: whether a
`Validated` rule's case passed on the release candidate is the
[evidence record](#evidence-record)'s to show.

## Evidence record

Each run of a check in Excel records:

- the commit of the source and the workbook it was built into;
- the case, its expected value, that value's provenance and the tolerance;
- the observed value and the outcome;
- the Excel build, its bitness and the Windows version;
- who ran it, who reviewed it, and when.

A check that was not run is recorded as `NOT RUN` with the reason. It is
never a pass.

An Excel run of the harness keeps the host, who ran it, when and how it
ended in its [host record](TEST_HARNESS.md#host-record), also when it stopped
early. `tools/check_excel_run.py` rejects a record that lacks one of them or
leaves who ran it, when or how it ended `UNKNOWN`, and accepts an outcome
other than `PASS` only as evidence of that outcome.

### What an accuracy contract proves

A frozen [accuracy contract](TEST_CASES.md#accuracy-contracts) whose results
pass on an Excel run of the candidate shows that, on that host, the worst
error of the calculation over its calibration and holdout cases stays within
the threshold. The threshold is the guarantee a release may publish. A
target, when the contract declares one, is only reported: a worst error
above it but within the threshold passes, marked `target missed`, and shows
that the calculation is less accurate than aimed for, not that the guarantee
fails. The holdout cases played no part in setting the threshold,
and their values are `published` or `independent`. The evidence for the
first is the [holdout exposure record](TEST_CASES.md#holdout-exposure-record):
the holdout stays sealed while the contract is draft, and the record binds
its first measurement under the contract's decision to the contract entry,
its threshold, the holdout's fixtures and expected files, and the design
document if any. The check fails when any of them changes, or the threshold
is lowered, under the same decision. The record shows what was measured with
the tools, not what someone computed by hand: a threshold is only as
independent of its holdout as the people who set it kept it. It proves
nothing outside those cases, regime and host. A draft contract, a frozen
contract with no recorded exposure, and any contract on `illustrative`
values, validate nothing. Record each contract's worst errors with the
evidence, with its threshold, its target and the number of exposures of its
holdout.

An [invariant](TEST_CASES.md#invariants) of the registry that the same run
passes shows that, on that host, the results of its cases keep its bounds,
its order or its sum, whatever the status of the contracts over them. It
proves only what its `basis` states, on those cases, and holds for no other
input. Record each invariant's line with the evidence, with the monotone
reversals it counted within the allowance.

<a id="host-behaviour"></a>

## Host behaviour

The code relies on how Excel and the VBE behave: which settings an operation
changes, how a name is re-cased, what an import needs, how a window or an
event behaves. Such behaviour can differ by Excel build, channel, bitness and
policy. Record each behaviour that the code or a guide relies on here, with
whether it was measured:

| ID | Behaviour | Status | Probe | Host | Date | Result | Issue |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `HB-01` | The VBE re-cases every use of a name to its latest spelling, so an exported module can differ from its source only in the letter case of code outside strings and comments | `assumed` | The next Excel run by hand: the `Note:` lines of `tools/check_excel_run.py` | — | — | Not measured: no Excel run has recorded a re-cased export yet. `tools/check_excel_run.py` tolerates it: it compares that code without letter case and reports each re-cased identifier as a note | #245 |
| `HB-02` | Excel started through COM can follow the user's macro setting (`AutomationSecurity` is `msoAutomationSecurityByUI`, set by policy). A workbook it opens then has its macros disabled, and `Application.Run` fails with *Excel 4.0 (XLM) macros have been disabled*, because the VBA procedure is not found | `measured` | By hand: in PowerShell, open the built workbook through COM and `Run` `'<workbook>'!TEST_Harness.LastResult` | Microsoft 365 Excel 16.0 build 20430, channel not recorded, 64-bit; Windows 11 Enterprise 25H2, build 26200.9711 | 2026-11-03 | `AutomationSecurity` was `ByUI`, and `Run` failed with that message. `tools/devbuild/excel_run.py` changes no macro setting: it waits for **Enable Content** | #204 |
| `HB-03` | `Application.Run` reaches a public procedure of an `Option Private Module` module when the name is qualified with the workbook (`'Book.xlsm'!Module.Procedure`) | `measured` | As `HB-02`, in a session whose macros were enabled | As `HB-02` | 2026-11-03 | `'Workbook-development.xlsm'!TEST_Harness.LastResult` ran | #204 |
| `HB-04` | In a workbook opened through COM with its macros disabled, **Enable Content** enables them without reopening, and `Application.Run` then reaches its procedures | `measured` | `python tools/devbuild/excel_run.py` on a host with `HB-02`: click **Enable Content** in the yellow bar | As `HB-02` | 2026-11-03 | Excel showed the yellow bar; after the click, the certification ran to `excel-run: PASS` without reopening the workbook | #204 |
| `HB-05` | The project that `tools/devbuild/vba_project.py` writes holds no compiled code (`_VBA_PROJECT` version `0xFFFF`, no performance cache), so Excel reads it from source and compiles each module when it is first used: the code that runs is the source the build wrote, and no Trust access to the VBA project object model is needed | `assumed` | By hand: open a development build in Excel, check that `? TEST_Harness.LastResult` runs in the Immediate window, which compiles that module on its first use, then check that **Debug → Compile VBAProject**, which compiles the other modules, completes | — | — | Not measured by a probe of its own. `tools/devbuild/build_workbook.py` relies on it, and so does the identity by rebuild of `tools/check_excel_run.py`, which proves that the workbook Excel ran holds the source of the commit, not that Excel compiled it | #285 |
| `HB-06` | **End** in an error dialog, or **Run → Reset** in the VBE, resets the project and clears every module-level variable, so the harness's active run and recovery flag are lost and the stop is not detected | `assumed` | By hand: set a breakpoint on the line `TEST_Harness.BeginCase "text-exact"` of `TEST_Examples.CaseExactText`, type `TEST_Examples.RunExampleTests "<fixtures>"` with the full path of `tests/fixtures/` in the Immediate window, and when the run pauses there, choose **Run → Reset**; then check that `? TEST_Harness.LastResult` is empty and that the next run is not refused | — | — | Not measured. The harness guides tell you to check Excel's state by hand after such a stop | #300 |
| `HB-07` | The Immediate window keeps only its last 200 lines | `assumed` | By hand: print 250 numbered lines with `Debug.Print`, then check that the first line left is line 51 | — | — | Not measured. The installation guide tells you to save the whole log with `TEST_Harness.SaveLog` | #300 |

- **Status:** `measured` when a probe ran on the named host and its output is
  kept; `assumed` otherwise.
- **Probe:** a `TEST_` procedure under `tests/modules/` that observes the
  behaviour and prints what it saw, run with the VBA test harness. A
  measurement found by hand names the steps instead.
- **Host:** the Excel product, version and build, its channel and bitness,
  and the Windows version, as in the [evidence record](#evidence-record).
- **Result:** what was observed, and the model it supports. A measurement
  holds for the named host only; it is never generalised to others.

A guide that states host behaviour the code relies on cites its `HB-` entry.
An `assumed` entry is listed under *Limits* in the
[release certification](../release/RELEASE_CERTIFICATION.md) until it is
measured.

## Open decisions

| Decision | Options | Issue |
| --- | --- | --- |
| *Illustrative:* where the reference values that validate `CALC-01` come from | A worked example of a registered source (`published`); an external implementation by another author (`independent`, `IMPL-`) | Link |
