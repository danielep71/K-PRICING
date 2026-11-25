# 🧪 VBA test harness

The harness runs the workbook's VBA test suites in Excel and writes a complete,
line-based log of each run. A certification runs a suite in every mode,
writes the evidence of the run, and raises only afterwards.
`tools/check_excel_run.py` then checks that evidence against one commit,
without Excel. [Recording an Excel run](../../INSTALLATION.md#recording-an-excel-run)
gives the procedure; this guide defines the modules, the outcomes and the
records.

## Modules

| Module | Holds |
| --- | --- |
| `tests/modules/TEST_Harness.bas` | The run lifecycle (`BeginRun`, `BeginCase`, `RecordCaseError`, `EndRun`), the assertions (`AssertEqualDouble`, `AssertNear`, `AssertEqualText`, `AssertTrue`, `AssertErrorValue`, `AssertRaised`, with `WithinTolerance` as the comparison rule), the log (`Say`, `Echo`, `SaveLog`, `LastResult`), the [results](#results) (`RecordResult`, `SaveResults`), recovery (`ResetTests`) and certification (`BeginCertification`, `CertifyRun`, `EndCertification`) |
| `tests/modules/TEST_Examples.bas` | The example suite: one case per kind of assertion, one case that reads the illustrative [test cases](TEST_CASES.md) and records their totals, and the reference implementation of the [cleanup pattern](../VBA_HOUSE_STYLE.md#operation-and-cleanup-outcomes) with its failure paths. It holds no calculation of the workbook |
| `tests/modules/TEST_Relations.bas` | The illustrative [relation tests](#relation-tests), `RunRelationTests`: four relations of a stand-in total, each with its corruption check. It is not certified, and holds no calculation of the workbook |

All three use `Option Private Module`, so they never become part of the
workbook's external surface. A production workbook contains none of them.

A project suite follows the example suite: one `Case` procedure per case,
each starting with `BeginCase` and ending in an error handler that calls
`RecordCaseError`; a runner that calls every case once, between `BeginRun`
and `EndRun`; and the expected case and assertion counts as constants. To be
certified, it keeps the shape that `tools/check_excel_run.py` reads from the
source, because it compares every case and assertion of the log with it:

- the runner, `RunSuite` unless the suite's `suites` entry names another as
  `runner`, calls each case procedure by its bare name, `Case<Name>`, alone
  on its line;
- each case calls `TEST_Harness.BeginCase "<case-id>"` with a literal id,
  and each assertion has a literal name, such as
  `TEST_Harness.AssertNear "holds", ...`. The check reads them only in the
  body of the case procedure itself: an assertion in a procedure that the
  case calls is not counted, and the certification fails;
- the counts are `Private Const EXPECTED_CASES As Long = <n>` and
  `Private Const EXPECTED_ASSERTIONS As Long = <n>`.

The expected values of a numerical case come from its
[test case](TEST_CASES.md), with the tolerance that the case states. A case
that computes an output of an accuracy contract also records it with
`RecordResult`, and the suite's certification entry point saves those
results in the suite's own results file.

The example suite's entry points take the full path of `tests/fixtures/` in
the checkout, from which its results case reads the illustrative test cases:
`RunExampleTests "<fixtures>"`, and so on.

## Run outcomes

`EndRun` reports four outcomes separately:

| Outcome | Values | Meaning |
| --- | --- | --- |
| `operation` | `PASS`, `FAIL`, `ERROR`, `NOT_RUN` | What ran: every assertion passed, an assertion or a case identifier failed, or an unexpected error occurred; `NOT_RUN` when the start was refused |
| `completeness` | `COMPLETE`, `INCOMPLETE` | `COMPLETE` only when both expected counts were met exactly. A skipped case is `INCOMPLETE`, whatever the cases that ran did ([completeness policy](TEST_CASES.md#completeness-policy)) |
| `cleanup` | `PASS`, `FAIL`, `NOT_RUN` | Whether Excel's state is as it was at `BeginRun`: `Calculation`, `DisplayAlerts`, `EnableEvents`, `ScreenUpdating`, the number of open workbooks and the active window; `NOT_RUN` when the start was refused |
| `recovery` | `NONE`, `REQUIRED` | `REQUIRED` when cleanup failed or could not be checked, or when a run was interrupted while the project kept its state (see below) |

`RESULT=PASS` needs `operation=PASS`, `COMPLETE`, cleanup `PASS` and no
recovery. Any other run raises `ERR_TEST_FAILED`, or the unexpected error that
escaped the runner, only after its summary is logged.

A run that starts while another is active, or while recovery is required, is
refused with `ERR_TEST_DIRTY_START`. Check Excel's state by hand, then type
`TEST_Harness.ResetTests`; it restores nothing itself.

The harness keeps its state in module variables, which a project reset
clears ([HB-06](VALIDATION_PLAN.md#host-behaviour)). A run stopped with **End** in an error dialog, or with **Run →
Reset** in the VBE, therefore leaves no active run and no recovery flag: the
stop is not detected, and the next run starts normally and takes Excel's
state as it finds it as its baseline. Every runner traps its own errors, so
a run that stops early has usually been stopped this way. After such a stop,
check by hand, before the next run, the four settings the cleanup compares
(`Calculation`, `DisplayAlerts`, `EnableEvents`, `ScreenUpdating`), the open
workbooks and the active window, and put back anything the stopped run
changed ([running the harness](../../INSTALLATION.md#running-the-harness)).

## Modes and certification

The example suite runs in each of these modes, in the order of `modes` in the
[evidence policy](#evidence-policy):

| Mode | Log | Must end |
| --- | --- | --- |
| `NORMAL` | `examples.log` | `RESULT=PASS`, `COMPLETE` |
| `INJECTED_FAILURE` | `injected.log` | `RESULT=FAIL`, `operation=FAIL`, `COMPLETE`, one failure: the designated assertion |
| `SKIPPED_CASE` | `skipped.log` | `RESULT=FAIL`, `operation=PASS`, `INCOMPLETE`, one failure: `harness.completeness` |

`TEST_Examples.CertifyExamples "<folder>", "<fixtures>"` is the certification
entry point:

1. `BeginCertification` refuses to start while a run or a certification is
   active, or while recovery is required.
2. Each mode runs. Its error is recorded, not raised, and its log is saved,
   so that every log is written whatever the run did. After the normal run,
   `SaveResults` saves its results as `results.json`.
3. `CertifyRun` records how the run reached its caller. The normal run must
   pass and be complete. Each deliberate failure must reach its caller as
   `ERR_TEST_FAILED` with a description. This is the self-test that a failed
   run is never silent.
4. `EndCertification` writes `certification.log`, then raises
   `ERR_TEST_FAILED` if the certification did not pass.

A project suite is certified in its normal run only. Its certification entry
point takes the same two arguments and follows the same steps with one mode:

```vb
Public Sub CertifyProject(ByVal Folder As String, ByVal Fixtures As String)
        TEST_Harness.BeginCertification SUITE_NAME
        ' Run the suite in NORMAL mode, capturing its error as CertifyMode does,
        TEST_Harness.SaveLog Folder & "\project.log"
        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription
        TEST_Harness.SaveResults Folder & "\project-results.json"
        TEST_Harness.EndCertification Folder & "\project-certification.log"
End Sub
```

List the suite under `suites` in the [evidence policy](#evidence-policy), with
the names of its two logs and, when it records results, of its results file.
Each suite saves its results in a file of its own: `BeginRun` clears the
results, so `SaveResults` writes only those of its own suite's run, and a
project suite writing `results.json` would replace the example suite's file.
A suite that records no result calls no `SaveResults` and names no results
file. Each suite declares its counts for the Office bitness the policy
supports: a case or assertion inside `#If Win64` counts only where that
branch is compiled, and the check reads the suite the same way.

## Relation tests

A relation test checks a property of the calculation that needs no external
expected value: additivity, symmetry under a reordering of the inputs,
monotonicity, or a zero input giving zero. It complements the
[test cases](TEST_CASES.md), whose expected values come from a reference.
`TEST_Relations.bas` shows the pattern on a stand-in total:

- **One case per relation.** The case computes the outputs the relation
  links, such as the total of two lists together and the totals of each, and
  asserts that the relation holds, as `holds`: an equality with `AssertNear`
  and its tolerance, an order such as monotonicity with a comparison.
- **A corruption check in every case.** The same case shifts one of those
  outputs far outside the tolerance and asserts, as `detects-corruption`,
  that the same comparison rejects it. A relation whose comparison accepts the
  corrupted output too, for example because its tolerance is too wide or it
  compares a value with itself, would pass whatever the calculation returned;
  the corruption check makes that a failure.
- **Fixed inputs.** The inputs are written in the case, so each run computes
  the same values and the expected counts hold.

To test a calculation of the workbook, replace `IllustrativeTotal` with a
call to its core procedure and the relations with those its
[calculation contract](CALCULATION_CONTRACTS.md) implies, derived from the
contract and never from the code under test. Run the suite
with `TEST_Relations.RunRelationTests`. It is not listed in the evidence
policy, so an Excel run does not certify it; to certify it, give it a
certification entry point and a `suites` entry as above, with
`"runner": "RunRelationTests"`: it already has the shape the check reads. A relation that holds
validates no figure: only a `published` or `independent` case does
([completeness policy](TEST_CASES.md#completeness-policy)).

## End-to-end case runner

A case runner tests the workbook as a user runs it: it writes a test case's
inputs to the input sheets, runs the calculation through its public entry
point, compares the outputs with the case's expected values within the case's
tolerance, and restores the sheets. It needs the input and result areas of the
[workbook contract](WORKBOOK_CONTRACT.md), which the repository does not lay
out yet, so it ships no runner. The suite's runner calls one case procedure per
test case, `Case<Name>`, which follows this skeleton; it uses the contract's
illustrative `Inputs_Settings` area and `M_Run.RunCalculation` entry point,
and the constant `SETTING_ROWS`, the number of rows the contract fixes for
that area (one):

```vb
Private Sub CaseExample()
    Dim Anchor           As Range      'The first header cell of the area
    Dim Area             As Range      'The input area the case fills
    Dim Saved            As Variant    'Its values before the case
    Dim RestoreDue       As Boolean    'True once Saved holds them
    Dim Detail           As String     'Why the area could not be restored
    Dim ErrNumber        As Long       'Error of the case; 0 for none
    Dim ErrSource        As String
    Dim ErrDescription   As String

        On Error GoTo Failed
        TEST_Harness.BeginCase "<case-id>"
        Set Anchor = ThisWorkbook.Names("Inputs_Settings").RefersToRange
        Set Area = Anchor.Offset(1, 1).Resize(SETTING_ROWS, 1)
        Saved = Area.Value
        RestoreDue = True
        ' Write the case's inputs to Area, then run the entry point:
        M_Run.RunCalculation
        ' For each output of tests/expected/<case-id>.json, read it where the
        ' workbook contract places it, then call TEST_Harness.AssertNear with
        ' a literal name, its value and tolerance, and
        ' TEST_Harness.RecordResult with its scope.

CleanUp:
        On Error GoTo 0
        If RestoreDue Then
            Detail = RestoreInputs(Area, Saved)
        End If
        TEST_Harness.AssertTrue "inputs-restored", Len(Detail) = 0, Detail
        If ErrNumber <> 0 Then
            On Error GoTo CaseFailed
            Err.Raise ErrNumber, ErrSource, ErrDescription
        End If
        Exit Sub

Failed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Resume CleanUp

CaseFailed:
        TEST_Harness.RecordCaseError
End Sub
```

- **Certification.** Each test case has its own `Case<Name>` procedure,
  because the check reads the case id and the assertion names as literals of
  the case procedure ([modules](#modules)). Every assertion is written there:
  one per output of the case's expected file, and `inputs-restored`. A
  runner that passes the case id to one shared procedure, or an assertion
  made in a procedure the case calls, does not have that shape, and the
  certification fails. Keep the shared steps, such as filling and restoring
  the area, in helpers that return what the case asserts, as `RestoreInputs`
  returns what failed.
- **The inputs.** The runner stands in for the user, in a development build
  only. It writes only `Input` areas, which code otherwise never writes, and
  finds each by its defined name. A dataset that the
  [data contract](DATA_CONTRACT.md) reads from a file is never written to a
  sheet: the runner gives the import the case's files instead, the variant
  directory of a [file-based case](TEST_CASES.md#file-based-cases).
- **The area.** The defined name of an area is its table anchor, which
  refers to one cell, the area's first header cell
  ([areas](WORKBOOK_CONTRACT.md#areas)). The runner derives the cells it
  writes from it: in a settings area, only the value column of the rows
  below the anchor, which users type and the sheet leaves unlocked, sized by
  the rows the workbook contract fixes, as above; in an area whose number of
  rows varies and that blank cells separate from its neighbours,
  `Anchor.CurrentRegion` less its header row, limited to its input columns.
  State in the suite which one it uses. Capturing the
  anchor alone would restore one cell, and `inputs-restored` would pass while
  the rest of the area kept the case's inputs.
- **The sheets.** It captures an area before changing it and restores it on
  the success and the error path, through a helper such as `RestoreInputs`
  that contains its own error and returns what failed, as the
  [cleanup pattern](../VBA_HOUSE_STYLE.md#operation-and-cleanup-outcomes)
  requires. The case asserts `inputs-restored` on that result on both paths,
  so the counts hold.
- **Excel's state.** The entry point owns the settings it changes, such as
  `Calculation` and `ScreenUpdating`, and the runner never restores them.
  `EndRun` compares them with their values at `BeginRun`: a setting the
  entry point left changed fails cleanup and requires recovery, and the next
  run is refused until you check Excel and type `TEST_Harness.ResetTests`
  ([run outcomes](#run-outcomes)). That check does not see cell values, so a
  failed `inputs-restored` is the only record of a sheet left changed: check
  the input sheets by hand before the next run.
- **The case's error.** It is saved before cleanup, raised again once the
  sheets are restored, and recorded by `RecordCaseError`, so the next case
  still runs.

## Volume benchmarks

A volume benchmark times the calculation on a large synthetic input. Its
timing depends on the host and its load, so it is never evidence that a
figure is right, and it stays outside the certification. A benchmark module
follows these conventions; none ships with the harness:

- **Name.** `TEST_Benchmark` plus a subject, such as
  `TEST_BenchmarkTotals`, in `tests/modules/`, so that a development build
  holds it and a production build does not.
- **Input.** Built in memory by a fixed rule, such as the record's index
  setting its amount, never read from a committed file, so that every run
  times the same input. It respects the data contract, as a fixture's inputs
  do.
- **Run.** Between `BeginRun` and `EndRun`, under the benchmark's own suite
  name, so that the harness refuses a dirty start and checks that the run
  left Excel's state as it found it. Its assertions check what volume can
  break, such as the number of results and that each is finite, never a time
  limit. It prints the number of records and the elapsed seconds, measured
  with `Timer`, with `TEST_Harness.Echo`.
- **Evidence.** Never listed in the [evidence policy](#evidence-policy):
  `tools/check_excel_run.py` rejects a policy that names a `TEST_Benchmark`
  module as a suite. Save its log, if you keep it, outside the evidence
  folder, for example under `test-results/benchmark/`, and record the host as
  the [evidence record](VALIDATION_PLAN.md#evidence-record) describes.

## Log format

One record per line, `KEY=value`, with further fields separated by `; `.
`RUN` and `ENVIRONMENT` come first, except after a refused start (below);
`CASE`, `ASSERT`, `ERROR` and `FAILURE`
follow as the run reaches them; the summary records come last, in the order
of the table:

| Record | Fields |
| --- | --- |
| `RUN=<suite>` | `mode` (`NORMAL` or a deliberate failure mode), `expected_cases`, `expected_assertions`, `started` |
| `ENVIRONMENT=` | `host`, `version`, `build` (`Application.Build`), `os`, `office` (`64-bit` or `32-bit`), `runtime` (`VBA7` or `VBA6`) |
| `CASE=<case-id>` | One per case, in run order |
| `ASSERT=<case-id>/<name>` | `status` (`PASS` or `FAIL`); on failure, the expected and actual values |
| `FAILURE=<id>` | One per failure: an assertion, `<case-id>.unexpected-error`, `runner.unexpected-error`, `harness.case-id`, `harness.no-case` or `harness.completeness`, with its detail |
| `ERROR=<case-id>` or `ERROR=run` | `number`, `source`, `description` of an unexpected error |
| `CASES=<n>`, `ASSERTIONS=<n>` | `expected` |
| `FAILURES=<n>` | `errors` |
| `PRIMARY_FAILURE=<id>` | The first failure of the run, when there is one |
| `CLEANUP=PASS` or `FAIL` | `detail`: the state compared, or what changed |
| `RECOVERY=NONE` or `REQUIRED` | `reason`, `action`, when required |
| `RESULT=PASS` or `FAIL` | `suite`, `operation`, `completeness`, `cases`, `assertions`, `failures`, `cleanup`, `recovery` |

`RESULT` is always the last record of a run. A start that `BeginRun`
refuses writes only two records, after those the log already holds: a
`RECOVERY=REQUIRED` with its `reason` and `action`, then
`RESULT=FAIL; suite=<suite>; operation=NOT_RUN; completeness=INCOMPLETE; cases=0; assertions=0; failures=0; cleanup=NOT_RUN; recovery=REQUIRED`,
with no `RUN` or `ENVIRONMENT`. A start that `BeginCertification` refuses
writes no record. `ResetTests` writes `RESET=DONE`. Case identifiers and
assertion names are fixed in the suite's source, so two logs of the same
commit can be compared line by line, apart from `started` and
`ENVIRONMENT`.

## Certification log

`certification.log` holds one record per line, in this order:

| Record | Fields |
| --- | --- |
| `CERTIFY=<suite>` | `started` |
| `MODE=<mode>` | One per mode, in run order: `expected` (`PASS` or `FAIL`), `error` (the number that reached the caller, `0` for none), `result` (`PASS`, `FAIL`, or `NONE` when the run wrote no `RESULT`), `completeness` (`COMPLETE`, `INCOMPLETE` or `NONE`), `status` (`PASS` when the run ended as its mode requires) |
| `CERTIFICATION=PASS` or `FAIL` | `suite`, `completeness` (`COMPLETE` when every run that must pass was complete), `modes`, `failures` |

## Results

`RecordResult` keeps one result of the running suite: its case, its variant
(an empty string for a JSON fixture), its quantity, the dataset and record of
its scope (two empty strings for a result of the whole calculation) and its
value. It raises `ERR_TEST_USAGE` outside a run.
The results belong to the run, like its log: `BeginRun` and `ResetTests`
clear them.

`SaveResults "<path>"` writes them as the results file of the
[accuracy contracts](TEST_CASES.md#accuracy-contracts): ASCII JSON with LF
line endings, an empty variant or scope written as `null`, a record scope as
`{"dataset": …, "record": …}`, and each value
with the 15 significant digits of `Str$`, whatever the regional settings.
Its `commit` is the value of the defined name `build_commit` in the
workbook that holds the harness, the build stamp of
`tools/devbuild/build_workbook.py`. A workbook without that name writes an
empty commit, and a build with local changes writes its marked stamp; the
accuracy check rejects both. This is the only part of a workbook the harness
reads, and it writes none.

Each certification saves the results of its own normal run: the example
suite as `results.json`, a project suite in the file its `suites` entry names.
`check_excel_run.py` measures the accuracy contracts on all these files
together, and rejects a result that two of them hold: each output is computed
by one suite.

## Evidence folder

| File | Written by |
| --- | --- |
| `examples.log`, `injected.log`, `skipped.log` | `CertifyExamples`, one per mode |
| `results.json` | `SaveResults`, after the example suite's normal run |
| `certification.log` | `EndCertification` |
| A log and a certification log per project suite, and its results file when it names one, as `suites` names them | Its certification entry point |
| `exported/` | By hand only: every VBA component exported back from the workbook after the run, each form with its `.frx`. Without it, `check_excel_run.py --workbook` proves identity by rebuilding the commit |
| `host.txt` | You, or `tools/devbuild/excel_run.py`: the [host record](#host-record), also for a run that stopped early |
| `export-manifest.json` | By hand only: `tools/devbuild/export_manifest.py`, the SHA-256 of the workbook and of every exported file, which binds the export to the workbook |

<a id="host-record"></a>

### Host record

`host.txt` holds one `key=value` line per field, in this order:

| Field | Value |
| --- | --- |
| `operator` | Who ran it |
| `started`, `finished` | When the run started and finished, in ISO 8601 with the UTC offset, such as `2026-11-03T14:05:09+01:00` |
| `windows` | The Windows version, as `winver` shows it |
| `excel` | `Application.Version` and `Application.Build`, such as `16.0 build 17928` |
| `bitness` | `64-bit` or `32-bit`: the bitness of Excel |
| `automation_security` | `Application.AutomationSecurity` as found: `Low`, `ByUI` or `ForceDisable` (1, 2 or 3). It is read, never set |
| `vba_project_access` | *Trust access to the VBA project object model* as found: `TRUSTED` or `NOT_TRUSTED` |
| `references` | The ticked references, separated by `; `. A run of `tools/devbuild/excel_run.py` records the references the build gives the project instead, which it does not read back in Excel |
| `compile` | `PASS` when **Debug → Compile VBAProject** completed, or what the compile reported. `ON_DEMAND` when VBA compiled only the modules the run called, as `tools/devbuild/excel_run.py` records it: no automated run can compile the whole project without Trust access. The check accepts `ON_DEMAND` only with identity by rebuild (`--workbook` and no `exported/` folder). A person compiles with **Debug → Compile VBAProject**, which needs no Trust access, and writes `PASS` |
| `outcome` | How the run ended: `PASS`, `FAIL`, `IMPORT_FAILED`, `COMPILE_FAILED`, `EXECUTION_TIMEOUT`, `INCOMPLETE`, `NOT_ENABLED` or `UNAVAILABLE` |

`PASS` means every certification entry point returned, and `FAIL` that one
raised. The other outcomes say where the run stopped: the components could
not be imported, the project did not compile, Excel did not finish in time,
the run stopped before its end, the macros were not enabled, or Excel could
not be started. A value that could not be read, such as a registry value whose
read was denied, is `UNKNOWN`, never a default or a setting of lower
precedence. Only the record of a run that did not pass may hold it, whether it
ended `FAIL` or stopped early; the check rejects it in the record of a `PASS`. No record may leave unread who ran it,
when or how it ended: the check rejects `UNKNOWN` in `operator`, `started`,
`finished` and `outcome` whatever the outcome.

`check_excel_run.py` rejects a record that lacks a field or holds a value
outside this table, a time without its UTC offset, a finish before the start,
or, for a `PASS`, references other than exactly those of the
[evidence policy](#evidence-policy), a `compile` other than `PASS` (or
`ON_DEMAND` with identity by rebuild), or an `excel` or `bitness` other than
the logs' `ENVIRONMENT`.
It accepts any other outcome only as evidence of that outcome: the check then
fails, whatever the rest of the folder holds.

## Evidence policy

[`.github/excel-evidence-policy.json`](../../.github/excel-evidence-policy.json)
holds what an Excel run must leave as evidence:

- `harness` and `suite`: the harness module, the example suite's name, source
  module and certification entry point, and the case and assertion counts it
  declares as `EXPECTED_CASES` and `EXPECTED_ASSERTIONS`;
- `suites`: the project suites, each with its name, source module,
  certification entry point, counts, `log` and `certification_log`,
  `results` when it records results, and `runner` when its runner is not
  `RunSuite`, run in `NORMAL` mode after the example
  suite; empty until the project adds its own, and never a
  [volume benchmark](#volume-benchmarks);
- `modes`: the example suite's modes in run order, the log each one saves and whether
  it must pass (only `NORMAL`);
- `evidence`: the names of the certification log, the export folder, the host
  record, the export manifest and the results file in the evidence folder;
- `host`: the exact references and the required Office bitness.

`tools/check_excel_run.py` reads the policy from the candidate commit, so a
run is judged by the policy of the commit it names, and rejects a policy that
is missing or malformed. It reads the harness and every suite at the exact
paths the policy names, as the policy's Office bitness compiles them, and
fails when the commit has no file there, or when a suite's counts or entry
point differ from the policy. Every log of every suite must come from one
host. A policy of version 1, which older commits hold, has no `suites`. The same policy is read
by `tools/devbuild/excel_run.py`, for the harness, the certification entry
point of each suite and the name of the host record, by
`tools/devbuild/export_manifest.py`, for the names of the export folder and
the manifest, by `tools/devbuild/lo_smoke.py`, for the example suite's
certification entry point and the names of its logs and results file, and by
`tools/render_docs.py`, for the expected results in the installation guide. `tests/test_check_excel_run.py` keeps the policy equal to
the suite, to the host policy of the installation guide and to the evidence
files the guide names. To change a mode, a file name, a reference or the
bitness, change the policy together with the suite, the installation guide
and this guide. A reference also changes `REFERENCES` in
`tools/devbuild/vba_project.py`, which the build writes into the project:
`tests/test_excel_run.py` fails while the references the build gives the
project differ from the policy.

## What it proves

The example suite exercises the harness itself: assertions, tolerances, error
values, raised errors, completeness, cleanup, certification and the results
file. Its results case reads the illustrative test cases, whose contract is
illustrative, so a passing run validates no calculation of the workbook.
Without Excel, `tests/test_vba_harness.py` checks what the source can show:
- the example suite's counts equal its cases and assertions;
- every case runs once with a unique identifier;
- every record in this guide is written;
- `EndCertification` writes the log before it raises;
- the results case records exactly the outputs of the contracts' cases, from
  files that exist;
- the harness reads the stamp name the build writes;
- the relation suite's counts equal its cases and assertions, and every
  relation case asserts both `holds` and `detects-corruption`.

`check_excel_run.py` proves that an evidence folder is consistent with a
commit, and measures the commit's accuracy contracts on its results files. It
cannot see Excel: the host record, from the operator to the outcome, is the
runner's own record.
