# 🧾 Excel run

An Excel run is the evidence that a commit's VBA was imported, compiled and
run on the supported host. The
[installation guide](../../INSTALLATION.md#recording-an-excel-run) records one
with one command. This page is the reference behind it: a run recorded by
hand, what the evidence check rejects, and when a run counts for another
commit. The [harness guide](../methodology/TEST_HARNESS.md) defines the
outcomes, the logs and the host record.

## By hand

1. **Fix the candidate.** In a clone with no local changes
   (`git status --short` prints nothing), type `git switch --detach <commit>`,
   then `git rev-parse HEAD` for the full SHA.
2. **Make the evidence folder**, for example `test-results\excel-run\` in the
   checkout (ignored by Git), with an empty `exported\` folder inside.
3. **Build a development workbook** from that checkout with
   `python tools/devbuild/build_workbook.py`, as in
   [building a workbook from source](../../INSTALLATION.md#building-a-workbook-from-source).
   Use
   [importing VBA into Excel](../../INSTALLATION.md#%EF%B8%8F-importing-vba-into-excel)
   only if the project has forms or the build fails. Then only the source
   identity is checked, not the binding of steps 7 and 9. Such a workbook has
   no `build_commit` stamp, so `SaveResults` writes an empty commit, which the
   [accuracy check](../methodology/TEST_HARNESS.md#results) rejects: while the commit holds
   `tests/accuracy/contracts.json`, as it does from the start, the run of a
   workbook imported by hand ends `excel-run: FAIL` with *commit is not a full
   lower-case commit SHA*, whatever else it shows.
4. **Compile** with **Debug → Compile VBAProject**.
5. **Certify.** In the Immediate window type, with the full path of the
   evidence folder and of `tests\fixtures` in the same checkout, both with no
   trailing backslash:

   ```text
   TEST_Examples.CertifyExamples "C:\path\to\test-results\excel-run", "C:\path\to\checkout\tests\fixtures"
   ```

   It refuses to start if an earlier run left state behind. It runs the
   suite normally, with the injected failure and with the skipped case, and
   saves each complete log as `examples.log`, `injected.log` and
   `skipped.log`. After the normal run it saves `results.json`: the
   results that run recorded, with the commit of the workbook's
   `build_commit` stamp. Then it writes `certification.log`. Each run prints a
   `SAVED=` line: the normal run with `error=0`, the two failure runs with
   the error *Suite examples did not pass*. Those errors are recorded, not
   raised, so that every log is written. The certification ends with
   `CERTIFICATION=PASS`; otherwise it raises *Certification of examples did
   not pass* after writing its log. Then type the certification entry point
   of each project suite that `.github/excel-evidence-policy.json` lists under
   `suites`, in order, with the same two arguments; each saves its log, its
   certification log and, when its entry names one, its results file.
6. **Export the modules back.** In the VBE, right-click each component →
   **Export File**, into `exported\`, under its own name. Export every VBA
   component of the commit, the test modules included; a form exports as its
   `.frm` with its `.frx`. Do not export a document module that the build
   generated and that has no file in `src/workbook/`. Do not edit the
   exported files.
7. **Bind the export to the workbook.** Close the workbook, then record the
   SHA-256 of the workbook and of every exported file:

   ```sh
   python tools/devbuild/export_manifest.py test-results/build/Workbook-development.xlsm test-results/excel-run
   ```

   It writes `export-manifest.json` in the folder. Do not open or save the
   workbook again before step 9.
8. **Write `host.txt`** in the folder:

   ```text
   operator=<your name>
   started=2026-11-03T14:05:09+01:00
   finished=2026-11-03T14:21:40+01:00
   windows=Windows 11 Pro, version 24H2, OS build 26100.4061
   excel=16.0 build 17928
   bitness=64-bit
   automation_security=Low
   vba_project_access=NOT_TRUSTED
   references=Visual Basic For Applications; Microsoft Excel 16.0 Object Library; OLE Automation; Microsoft Office 16.0 Object Library
   compile=PASS
   outcome=PASS
   ```

   The [host record](../methodology/TEST_HARNESS.md#host-record) defines
   each field:
   - `started` and `finished` are when you began step 4 and ended step 6,
     with your UTC offset.
   - `windows` is what `winver` shows.
   - `excel` is what `? Application.Version & " build " & Application.Build`
     prints in the Immediate window, and `bitness` what **File → Account →
     About Excel** shows.
   - `automation_security` is what `? Application.AutomationSecurity` prints:
     1 is `Low`, 2 `ByUI`, 3 `ForceDisable`. Only read it; never change it.
   - `vba_project_access` is `TRUSTED` when **File → Options → Trust Center →
     Trust Center Settings → Macro Settings** ticks *Trust access to the VBA
     project object model*, `NOT_TRUSTED` otherwise. Change nothing there.
   - `references` lists the boxes ticked in **Tools → References**, separated
     by `; `.
   - Write `compile=PASS` only if step 4 completed with no error; otherwise
     write what it reported. Never write `ON_DEMAND`: the check accepts it
     only with identity by rebuild, never with `exported\`.
   - `outcome` is `PASS` when step 5 ended with `CERTIFICATION=PASS` for
     every suite, and `FAIL` when a certification raised. When you stop
     early, write the record anyway, with `IMPORT_FAILED` (step 3),
     `COMPILE_FAILED` (step 4), `NOT_ENABLED`, `INCOMPLETE` or `UNAVAILABLE`
     (Excel would not start). In a record whose outcome is not `PASS`, write
     `UNKNOWN` for a value you could not read, but never for `operator`,
     `started`, `finished` or `outcome`.
9. **Check** the folder against the commit, with Python 3.11 or later and
   Git:

   ```sh
   python tools/check_excel_run.py test-results/excel-run --commit <commit> \
     --workbook test-results/build/Workbook-development.xlsm --require-binding
   ```

   It must end with `Binding: VERIFIED` and `excel-run: PASS`. It prints
   the commit, the host and the Git blob id of every imported component. When
   the commit holds accuracy contracts, an `Accuracy:` line gives each
   contract's worst error over its calibration and its holdout cases, with
   its threshold and its target, and says when the target was missed; the
   holdout line gives the number of its exposures, or says that the holdout
   is sealed: while the contract is draft, or until its first exposure is
   [recorded](../methodology/TEST_CASES.md#holdout-exposure-record), which
   fails the run for a frozen contract;
   another gives what each invariant of the registry checked. For
   a workbook imported by hand, omit `--workbook` and `--require-binding`:
   the binding is then `UNVERIFIED`, which a release does not accept, and
   while the commit holds accuracy contracts the run fails, because its
   results name no commit (step 3).
10. **Record.** Copy the report into the issue, the pull request or the
    [release certification](../release/RELEASE_CERTIFICATION.md), and add a
    row to the tested environments of the
    [host policy](../../INSTALLATION.md#%EF%B8%8F-host-policy). The folder
    stays outside Git: its exported modules would be second copies of the VBA
    components.

## What the check rejects

`check_excel_run.py` reads the evidence folder and the commit and, with
`--workbook`, the workbook; with `--workbook` and no `exported\` folder, it
also rebuilds the commit in a temporary Git worktree:

| It fails when | Because |
| --- | --- |
| An exported module differs from the commit's file line for line, a form's `.frx` differs byte for byte, or one is missing or extra. Code outside strings and comments is compared without letter case, because the VBE re-cases a name ([HB-01](../methodology/VALIDATION_PLAN.md#host-behaviour)): each re-cased identifier is reported on a `Note:` line instead. In a document module of `src/workbook/`, the attributes Excel writes itself (`VB_Base`, `VB_GlobalNameSpace`, `VB_Creatable`, `VB_PredeclaredId`, `VB_Exposed`, `VB_TemplateDerived`, `VB_Customizable`) are not compared; `VB_Name` is | The workbook was not built from that commit |
| A log is missing or empty, does not end with `RESULT`, holds more than one run, or a line is not a `KEY=value` record | The log is incomplete |
| A record is missing, extra or out of the order in which the commit's suite writes it, or an unexpected error (`ERROR`) is logged | The run did not execute the suite as written |
| The normal run is `INCOMPLETE` | A mandatory case or assertion did not run |
| The normal run is not `RESULT=PASS`, or its cases and assertions differ from the commit's suite in name or order | The run did not pass as written |
| `CLEANUP` is not `PASS` or `RECOVERY` is not `NONE` in any log of any suite | Cleanup failed |
| The two failure runs do not end as [documented](../../INSTALLATION.md#running-the-harness), or fail anything but their own failure: the injected run its one designated assertion, with the expected value the suite injects and the same values in each record, the skipped-case run only `harness.completeness` | The failure paths did not behave |
| `certification.log` is missing, a mode did not end as it must, a failure run did not reach its caller as `ERR_TEST_FAILED`, or it does not end with `CERTIFICATION=PASS` and `COMPLETE` | The certification did not pass |
| A project suite's log or certification log, as `suites` names them, is missing; its normal run is not `RESULT=PASS` and `COMPLETE`, or its cases and assertions differ from the commit's suite; or its certification log does not end with `CERTIFICATION=PASS` | The project suite did not pass |
| `ENVIRONMENT` lacks a field, differs between the logs, or Office is not 64-bit | The host is unknown or not the supported target |
| `host.txt` lacks a field or holds a value the [host record](../methodology/TEST_HARNESS.md#host-record) does not allow, a time lacks its UTC offset, the run finished before it started, the references are not exactly the four of the [host policy](../../INSTALLATION.md#%EF%B8%8F-host-policy), `compile` is not `PASS` (or `ON_DEMAND` with identity by rebuild: `--workbook` and no `exported\` folder, as the one command records it), `operator`, `started`, `finished` or `outcome` is `UNKNOWN`, another field is `UNKNOWN` in the record of a `PASS`, or `excel` or `bitness` differs from the logs' `ENVIRONMENT` | The host record is incomplete |
| `outcome` in `host.txt` is not `PASS` | The run stopped early or failed: the record is evidence of that outcome, never of a pass |
| Without `exported\`, a part of the workbook given with `--workbook` differs from a rebuild of the commit, or a part is missing or extra | It was saved after the build, or not built from that commit (`Binding: FAIL`) |
| `export-manifest.json` does not match the exported files or the workbook, or the workbook's `build_commit` is not the commit of a clean build | The export is not bound to a workbook built from the commit (`Binding: FAIL`) |
| With `--require-binding`, no manifest or no `--workbook` | The binding is `UNVERIFIED` |
| The commit holds `tests/accuracy/contracts.json`, and `results.json` or the results file a project suite names is missing or names another commit, a build with local changes or, for a workbook imported by hand, no commit; or two files hold the same result, the files together lack a result, a result lies outside its output's tolerance, or a contract's worst error exceeds its threshold | The run did not measure the contracts, or they do not hold ([accuracy contracts](../methodology/TEST_CASES.md#accuracy-contracts)) |
| With `--equivalent-to`, a [run input](#a-run-for-another-commit) has another Git blob at the two commits, or exists at only one of them | The run does not show what the other commit would run |
| With `--carried-to`, a VBA component differs in code, holds a non-ASCII byte or uses `[` or `]` outside a string, or another run input has another Git blob at the two commits or exists at only one of them ([carried evidence](#carried-evidence)) | The run does not show what the other commit would run |

It cannot see Excel: the host record, from the operator to the outcome, is
the runner's own record. A PASS says the evidence is consistent
with the commit. It validates a calculation only as far as its accuracy
contracts and their reference values do; an illustrative contract, such as
`ACC-01`, validates none.

## A run for another commit

A run is evidence for the commit the workbook was built from. It also counts
for a later commit that changes none of its run inputs, such as a commit
that changes only a guide outside `docs/methodology/` or a workflow. Check it
from a clean checkout of that later commit, with `--commit` naming the commit
of the run and `--equivalent-to` the full SHA of the commit to certify:

```sh
python tools/check_excel_run.py test-results/excel-run --commit RUN_SHA \
  --equivalent-to CANDIDATE_SHA \
  --workbook test-results/build/Workbook-development.xlsm --require-binding
```

The evidence and the binding are checked against `RUN_SHA`, as above: the
workbook's `build_commit` and the results still name it. Then every run input
must have the same Git blob at both commits, and none may be added or
removed. The run inputs are:

- every VBA component (`.bas`, `.cls`, `.frm`) and form resource (`.frx`)
  of the tree;
- every file in a directory of any configuration of
  `tools/devbuild/workbook.json`, at either commit, and in `src/workbook/`;
- every file in `docs/methodology/`, `tests/fixtures/`, `tests/expected/` and
  `tests/accuracy/`, which the suites or the accuracy check read, except the
  holdout exposure record `tests/accuracy/holdout_exposures.json`: it states
  which exposures happened, each bound by its hashes to run inputs, and the
  accuracy check reads it from the commit to certify;
- every `.gitattributes` file, at the root or in any directory, because each
  one can change the bytes Git checks these files out with;
- `tools/devbuild/workbook.json`, `src/workbook/Template.xlsx`,
  `.github/excel-evidence-policy.json`, and the build tools
  `tools/devbuild/build_workbook.py` and `tools/devbuild/vba_project.py`.

The report then names `CANDIDATE_SHA` as the candidate and `RUN_SHA` as the
run commit, and adds
`Equivalent: CANDIDATE_SHA (N run inputs identical to RUN_SHA)`. Any
difference is a `FAIL` line that names the path; record a new run for that
commit instead. Record both commits wherever you cite the run.

For a release, `tools/release.py evidence` runs this check with
`--require-binding`, for the candidate or, with `--run-commit`, for an
earlier run, and prints the `Excel evidence:` line of the tag message
([RELEASING.md step 4](../../RELEASING.md#excel-evidence-gate)).

### Carried evidence

A comment does not change what Excel runs. For a pull request, a run also
counts for a later commit whose VBA components differ from the run's only in
comments or layout, when you check it with `--carried-to` in place of
`--equivalent-to`:

```sh
python tools/check_excel_run.py test-results/excel-run --commit RUN_SHA \
  --carried-to CANDIDATE_SHA \
  --workbook test-results/build/Workbook-development.xlsm --require-binding
```

The evidence and the binding are checked against `RUN_SHA`, and every run
input must have the same Git blob at both commits, as above, with one
exception: a VBA component (`.bas`, `.cls`, `.frm`) at both commits may
differ, if `tools/vba_code_identity.py` finds the same code in both. It reads
each version with the lexer of the VBA source checks and compares their
logical lines without comments:

- comments are dropped: from an apostrophe outside a string, from `Rem` where
  a statement starts, after a colon too, and the lines of a comment continued
  with ` _`;
- continuation lines are joined, and blank lines are dropped;
- conditional-compilation directives (`#If`, `#Const`) are kept, so a
  changed directive, or a changed statement inside an `#If` block, is a code
  change;
- outside string literals, a run of white space counts as one space and the
  letter case does not count, because the VBE re-cases names
  ([HB-01](../methodology/VALIDATION_PLAN.md#host-behaviour)); each
  re-cased identifier is reported on a `Note:` line. White space between two
  tokens is never added or removed: `a=1` and `a = 1` are different code.

Anything else is a code change, and a `FAIL` line names the component and its
first changed statement, with its line at both commits. The check fails
closed: a non-ASCII byte, a component that uses `[` or `]` outside a string
literal (changed or not), a component added, removed or renamed, a form's
`.frx` that differs by one byte, or any other run input that differs is a
`FAIL`. The report then adds

`Carried from RUN_SHA: COMPONENTS differ only in comments or layout; the
other N run inputs are identical. Carried evidence never certifies a release`

with the 12-character `RUN_SHA`. The check derives this line each time it
runs and never writes it into the evidence folder, which stays the record of
`RUN_SHA`. Cite the run as `run at <SHA>, carried`, with that line.

Carried evidence is not evidence of the candidate itself. A release candidate
needs a run of its exact commit, or one that `--equivalent-to` proves
identical ([releasing](../../RELEASING.md#4-certify-the-workbook)).
