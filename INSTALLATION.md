# 📥 Installation: VBA import and export

This guide moves VBA source between the repository and Excel, in both
directions, without losing line endings, metadata or form resources. It also
covers building a workbook from source, running the VBA test harness,
recording an Excel run with one command, upgrading to a new revision,
troubleshooting and removal. The
[repository structure](docs/REPOSITORY_STRUCTURE.md) says where each component
lives, and the [house style](docs/VBA_HOUSE_STYLE.md) says how it is written.
The repository's own checks are in the
[contributing guide](CONTRIBUTING.md#-local-checks).

## 🖥️ Host policy

Adapt this section to the project and record the decision in an issue.

- **Target host:** Microsoft 365 Excel desktop on Windows, 64-bit. Windows
  32-bit, older perpetual versions, Mac and Excel for the web are not
  supported until a recorded decision and evidence add them.
- **References:** only the defaults of the target installation: Visual Basic
  For Applications, the Microsoft Excel 16.0 Object Library, OLE Automation
  and the Microsoft Office 16.0 Object Library. Any other reference needs a
  reviewed decision, and a change to the host references of
  `.github/excel-evidence-policy.json`. Before compiling, check **Tools → References**
  for entries marked *MISSING*.
- **Python** is used for repository tooling and for deriving reference values,
  never by the workbook.
- **Macro security:** follow your organization's policy. Nothing here needs
  a lower security level, or *Trust access to the VBA project object model*.

Record each host on which a commit has been compiled and run:

| Tested environment | Commit | Evidence |
| --- | --- | --- |
| Microsoft 365 Excel 16.0 build 20430, 64-bit; Windows 11 Enterprise 25H2, build 26200.9711; `tools/devbuild/excel_run.py`, `compile=ON_DEMAND` | `50c065859c7f566ef37946cca9c6206447951145` | #204: `excel-run: PASS`, `Binding: VERIFIED`, ACC-01 worst error 0.0 |

## 🔖 Start from an exact revision

```sh
git clone https://github.com/OWNER/REPOSITORY.git
cd REPOSITORY
git switch BRANCH
git rev-parse HEAD
git status --short
```

Record the full SHA, and start only when `git status --short` prints nothing.

Import from a Git checkout. `.gitattributes` checks VBA source out with CRLF
line endings, which the VBE needs. A raw download or a source archive holds
the stored LF files and skips that conversion. Never mix files from two
revisions.

## ⬇️ Importing VBA into Excel

1. **Workbook.** Create a blank workbook, or copy `src/workbook/Template.xlsx`
   if the project has one. Save it as `.xlsm` outside the checkout or under
   `test-results/`, which Git ignores.
2. **Project.** Open the VBE (**Alt+F11**) and select that workbook's
   project explicitly, never whichever one is active. Check the references.
3. **Modules and classes.** Import the `.bas` and `.cls` files of
   `src/core/`, `src/classes/` and `src/modules/`, then the standard `.bas`
   adapters of `src/workbook/`.
   - Import in file name order, so that every import is the same.
   - Skip a directory that does not exist.
   - Start from a fresh workbook, so that no obsolete or duplicate component
     remains.
4. **Forms.** Import each `.frm` of `src/forms/` with its `.frx` beside it;
   never import a `.frx` on its own.
5. **Document modules.** Do not import the `.cls` file of `ThisWorkbook` or a
   sheet as an ordinary class.
   - Find the existing document component with the same code name. Create a
     sheet only where the build instructions call for one.
   - Paste the file's code into that component, without its `VERSION`,
     `BEGIN` and `END` header and without its `Attribute` lines.
   - Hidden attributes need a reviewed round trip before such a component is
     accepted; never drop them silently.
6. **Tests.** A development workbook also imports `tests/modules/` and,
   when the project has it, `examples/modules/`, the reserved location of
   synthetic VBA examples. A production workbook never contains tests or
   examples.
7. **Compile and record.** Compile only once every component is present:
   **Debug → Compile VBAProject**. Then save the workbook and run the
   documented tests. Record the full SHA, the Excel build and bitness, the
   Windows version, the references, and what you ran with its output, as the
   [validation plan](docs/methodology/VALIDATION_PLAN.md#evidence-record)
   describes.

Manual import needs no *Trust access to the VBA project object model*.

## ⬆️ Exporting changes

1. **Export.** Export each changed component into a temporary folder first,
   keeping the VBE header and `Attribute` lines. Check that each file name
   stem equals its `Attribute VB_Name`.
2. **Document modules.** Export the component as `.cls` and compare it with
   the matching file in `src/workbook/`. Review differences in the code name
   and in hidden attributes, then update the file. Never replace the
   component by importing the `.cls`.
3. **Forms.** Review a changed `.frm` and its `.frx` together and copy them
   as a pair. Never edit a `.frx` by hand.
4. **Compare.** Compare the exports with the exact checkout. Reject non-ASCII
   source, unexpected changes in letter case and unexpected metadata changes.
   Then copy the intended changes into their source paths; Git normalizes the
   line endings when you add them.
5. **Check.** Run the [local checks](CONTRIBUTING.md#-local-checks) before
   committing. They check the repository, not VBA compilation or the results
   of a calculation.

<a id="building-a-workbook-from-source"></a>

## 🏗️ Building a workbook from source

`tools/devbuild/build_workbook.py` builds the workbook of steps 1 to 6 above,
apart from forms (step 4), in one command, from the checkout:

```sh
python tools/devbuild/build_workbook.py                            # development
python tools/devbuild/build_workbook.py --configuration production
```

It writes `test-results/build/Workbook-<configuration>.xlsm`, which Git
ignores, with:

- the sheets, code names and defined names of the
  [workbook contract](docs/methodology/WORKBOOK_CONTRACT.md), read from
  `tools/devbuild/workbook.json`;
- the layout of `src/workbook/Template.xlsx`, when the project has one: the
  build starts from a copy of it and keeps its formats and protection
  ([build](docs/methodology/WORKBOOK_CONTRACT.md#build)). Without it, the
  sheets are blank;
- the build stamp on the reference sheet: the commit, its committer time and
  the configuration. A tracked change, an untracked file in a source
  directory, or `--override` (below) adds *NOT evidence-grade* to the commit;
- a VBA project written from source. The production configuration takes the
  standard and class modules of `src/`. The development configuration adds
  `tests/modules/` and `examples/modules/`. Each sheet's document module
  comes from `src/workbook/<code name>.cls`, when that file exists.

The project holds no compiled code, so Excel reads it from source and
compiles each module when it is first used
([HB-05](docs/methodology/VALIDATION_PLAN.md#host-behaviour)): the build
needs no *Trust access to the VBA project object model*. Two builds of
the same commit and working tree are byte-identical. The build stops on a form,
which needs a designer the builder does not write; import forms by hand
(step 4). It also stops on a module too dense to fit the VBA compression
format; import that module by hand.

**Replacing modules for a negative control.** To show that a test fails
without a VBA fix, you need a workbook that differs from the commit by exactly
the modules under test, built the same way. Put each replacement module, as
the VBE exports it, in a folder that Git ignores, such as
`test-results/mutants/`, and build with `--override`:

```sh
python tools/devbuild/build_workbook.py --override test-results/mutants
```

- Each `.bas` or `.cls` file in the folder or its subfolders replaces the
  module of the configuration with the same `Attribute VB_Name`. A `.bas`
  replaces a standard module, and a `.cls` a class or document module. Other
  files are ignored.
- The build stops and writes nothing when a file names no module of the
  configuration, when two files name the same module (VBA names ignore case),
  when a name differs in case from its module, and on a `.frm`. It also stops
  on a file without exactly one `Attribute VB_Name` line, one that does not
  start with that line after a class's header, one with a byte outside ASCII,
  and one with a carriage return without a line feed. An override never adds
  or drops a module.
- The build prints each replaced module with its file and SHA-256. The
  stamp's commit records the folder, relative to the checkout when it lies
  inside it, and each replaced module with its SHA-256, and says
  *NOT evidence-grade*.

A workbook built this way is never evidence for its commit.
`tools/check_excel_run.py` reads the stamp of the workbook you give it with
`--workbook`, and the commit in each results file of the run. When either
records overrides, it reports the run as built with overrides and not
evidence for the commit, never as `Binding: VERIFIED` or `excel-run: PASS`.
Release evidence needs `Binding: VERIFIED`, so it never comes from such a
workbook. The replacement files are not in Git: keep them with the run's
notes if you need to repeat it.

1. Unblock the file if it was copied from another machine, open it, enable
   macros, and check **Tools → References**.
2. Compile with **Debug → Compile VBAProject**.
3. Changes made in the VBE are not read back: export them as
   [exporting changes](#%EF%B8%8F-exporting-changes) describes, then build
   again.

If Excel rejects or strips the embedded project, import by hand as in
[importing VBA into Excel](#%EF%B8%8F-importing-vba-into-excel), and record
the failure in an issue.

**Early warning in LibreOffice (optional).** Where LibreOffice and its
Python UNO bridge are installed, `python3 tools/devbuild/lo_smoke.py` builds
the development workbook and runs the harness certification in headless
LibreOffice; `python3 tools/devbuild/lo_smoke.py WORKBOOK` runs a workbook
you built. It reports only failures that LibreOffice is not known to cause,
and exits 0 when there is none, 1 on such a failure, when the build fails or
when LibreOffice cannot run the workbook, and 2 for an unknown option, a
missing workbook, a missing or malformed evidence policy, or a missing
LibreOffice or UNO bridge. LibreOffice is not a supported host, and this is
never Excel evidence.

<a id="running-the-harness"></a>

## 🧪 Running the harness

In a **development** workbook built as above, with the modules of
`tests/modules/`, open the Immediate window (**Ctrl+G**) and type, with the
full path of `tests\fixtures` in your checkout and no trailing backslash:

```text
TEST_Examples.RunExampleTests "C:\path\to\checkout\tests\fixtures"
```

It runs the example suite through `TEST_Harness` and writes one record per
line. One case reads the illustrative test cases from that folder and
records their totals as results. The entry points below that show the
failure paths take the same argument. A passing run ends with:

<!-- BEGIN generated: example-suite-results -->
```text
RESULT=PASS; suite=examples; operation=PASS; completeness=COMPLETE; cases=14; assertions=34; failures=0; cleanup=PASS; recovery=NONE
```

Anything else, or the error *Suite examples did not pass*, is a failure.
Two more entry points show the failure paths and are expected to fail:

| Entry point | Expected end |
| --- | --- |
| `TEST_Examples.RunExampleTestsWithInjectedFailure` | `RESULT=FAIL`, `operation=FAIL`, `completeness=COMPLETE`, `failures=1`, then *Suite examples did not pass* |
| `TEST_Examples.RunExampleTestsWithSkippedCase` | `RESULT=FAIL`, `operation=PASS`, `completeness=INCOMPLETE`, `cases=13`, `assertions=32`, `failures=1`, then *Suite examples did not pass* |
<!-- END generated: example-suite-results -->

- **Complete log.** The Immediate window keeps only its last 200 lines
  ([HB-07](docs/methodology/VALIDATION_PLAN.md#host-behaviour)). To
  keep the whole run, type `TEST_Harness.SaveLog "<folder>\examples.log"`
  after it, with a folder outside the checkout or the ignored
  `test-results/` directory.
- **Recovery.** A run that left Excel's state changed, or that stopped half
  way while the project kept its state, leaves `recovery=REQUIRED`, and the
  next run is refused. Check **File → Options → Formulas** (calculation), the
  open workbooks and the other settings the log names, then type
  `TEST_Harness.ResetTests`. A run stopped with **End** or **Reset** leaves
  no flag: see the next item.
- **End or Reset.** **End** in an error dialog, or **Run → Reset** in the
  VBE, clears the harness's state with the rest of the project's
  ([HB-06](docs/methodology/VALIDATION_PLAN.md#host-behaviour)), so the
  stop is not detected: the next run starts normally and takes Excel's state
  as it finds it as its baseline. After such a stop, check Excel by hand
  before the next run. In the Immediate window,
  `? Application.Calculation, Application.ScreenUpdating, Application.DisplayAlerts, Application.EnableEvents`
  prints the four settings the harness compares (`-4105` is automatic
  calculation), and the open workbooks and the active window must be those
  you had before the run. Put back anything the stopped run changed.
- **What it proves.** The suite exercises the harness itself and validates no
  calculation of the workbook. The [harness guide](docs/methodology/TEST_HARNESS.md)
  defines the outcomes and the log format.

<a id="recording-an-excel-run"></a>

## 🧾 Recording an Excel run

This is the evidence that a commit's VBA was imported, compiled and run on the
supported host, for a release and for any issue or pull request that cites
an Excel run. It needs no self-hosted runner, no change to the
macro-security level and no *Trust access to the VBA project object model*.

### One command

On Windows with Excel, Python 3.11 or later and Git, in a clone with no local
changes (`git status --short` prints nothing), at the candidate commit:

```sh
python tools/devbuild/excel_run.py
```

It builds the development workbook, and stops if the build is not
evidence-grade. Then it:

1. empties `test-results\excel-run\`;
2. opens the workbook read-only in Excel. If Excel shows the yellow bar
   *SECURITY WARNING Macros have been disabled*, click **Enable Content**
   (or **File → Info → Enable Content**). This happens where a policy makes
   programs follow your macro setting
   ([HB-02](docs/methodology/VALIDATION_PLAN.md#host-behaviour)). The command
   changes no macro setting, and waits up to `--enable-timeout` seconds (300
   by default) for that click, which enables the macros without reopening
   the workbook ([HB-04](docs/methodology/VALIDATION_PLAN.md#host-behaviour));
3. runs the certification of step 5 of a
   [run by hand](docs/repo_governance/EXCEL_RUN.md#by-hand), with that folder
   and `tests\fixtures`;
4. closes the workbook without saving, and quits Excel. If the Excel
   instance it started is still running 30 seconds later, it ends that
   process, and only that one: it checks the process name and the start
   time it recorded when it started Excel, so it never ends another Excel;
5. writes `host.txt`, the
   [host record](docs/methodology/TEST_HARNESS.md#host-record): you as the
   operator (your Windows user name, or `--operator "<name>"`), the start and
   finish times, Windows from the registry, the Excel build and bitness, the
   `AutomationSecurity` and the Trust access it found, which it reads and
   never changes, the four references the build gives the project,
   `compile=ON_DEMAND` and the outcome;
6. runs the check of step 9 of a run by hand with `--workbook` and
   `--require-binding`.

It prints the report and saves it as `test-results\excel-run-report.txt`.
It must end with `Binding: VERIFIED` and `excel-run: PASS`; record it as
step 10 of a run by hand says. The
[Excel run reference](docs/repo_governance/EXCEL_RUN.md#what-the-check-rejects)
lists what the check rejects. The command also prints an `excel-quit:` line, which the
saved report holds after its outcome, to say how Excel ended: `EXITED`;
`STOPPED` when the command ended its process; `SURVIVED` when it could
not, so close that Excel yourself; or `UNKNOWN` when it could not tell,
so close any Excel window the run left open.

There is no export: the check proves identity by rebuilding the commit in a
temporary worktree. The workbook Excel ran must hold the same parts as that
build, byte for byte once decompressed, the VBA project included, so it holds
the commit's source. That project holds no compiled code, so Excel reads it
from source and compiles each module when it is first used
([HB-05](docs/methodology/VALIDATION_PLAN.md#host-behaviour)): the code that
ran is the commit's.

Without Trust access no automated run can compile the whole project. VBA
then compiles only the modules the run calls, each when it is first used, and
`compile=ON_DEMAND` says so; a module the run never calls is not compiled.
`check_excel_run.py` accepts `ON_DEMAND` only with identity by rebuild,
`--workbook` and no `exported\` folder, as this command checks it. A run by
hand compiles with **Debug → Compile VBAProject**, which needs no Trust
access, and records `compile=PASS`.

Excel opens visibly while the run lasts. Apart from **Enable Content**, leave
it alone: a dialog it shows stops the run until you close it, and the command
gives up after `--timeout` seconds (600 by default). A run that stops early
still writes `host.txt`, with the outcome `NOT_ENABLED` when the macros were
not enabled in time, `EXECUTION_TIMEOUT` after `--timeout`, `UNAVAILABLE`
when Excel could not be started or `INCOMPLETE` when the run stopped before
its end. That record is evidence of how the run stopped, never of a pass;
run the command again.

Record the run [by hand](docs/repo_governance/EXCEL_RUN.md#by-hand) instead
when the project has forms, when the build fails, or when you need a full
compile. A workbook you import
by hand instead of building it has no build stamp, so the check fails while
the commit holds accuracy contracts. A run can also count for a later commit
that changes none of its inputs
([a run for another commit](docs/repo_governance/EXCEL_RUN.md#a-run-for-another-commit)).

## 🔄 Upgrading to a new revision

A workbook is never patched in place. To move to a new revision, build a new
workbook from it and carry your inputs across:

1. **Read the changes.** Read the changelog between the two revisions, and the
   `behavior-change` entries first. They name changed results, input layouts
   and output layouts, so you know which differences to expect.
2. **Keep the old workbook.** Keep it unchanged and record its commit, so you
   can compare results.
3. **Build the new one.** Import every component from the new revision into a
   fresh workbook, as [Importing VBA into Excel](#%EF%B8%8F-importing-vba-into-excel)
   describes. Never import some components from one revision and some from
   another, and never paste new source beside old components.
4. **Carry the inputs across.**
   - Re-import the input files the [data contract](docs/methodology/DATA_CONTRACT.md)
     defines.
   - Copy the values of each input area that the
     [workbook contract](docs/methodology/WORKBOOK_CONTRACT.md) lists, area by
     area, by its defined name. A sheet, area or name that the new revision
     renamed or removed is listed in the changelog. Never copy results:
     recalculate them.
5. **Compile and compare.** Compile, run the documented tests, and compare the
   results with the old workbook. Explain every difference with a changelog
   entry, or treat it as a defect. Record the run as for a new installation.

A copy with local changes is a fork. Export its changes first, compare them
with both revisions, and reapply them deliberately; never assume they survive
a rebuild.

## 🩺 Troubleshooting

| Symptom | Check |
| --- | --- |
| A compile error or a missing procedure | Every component comes from one revision and was imported; the references match the [host policy](#%EF%B8%8F-host-policy) |
| *Ambiguous name detected* | An obsolete or duplicate component remains. Start again from a fresh workbook |
| A reference marked *MISSING* | The workbook was built on another host. Remove it, or record a decision for it |
| A form without its controls | Its `.frm` was imported without the matching `.frx` beside it |
| Source with LF line endings, or one long line | It came from a raw download or an archive instead of a [Git checkout](#-start-from-an-exact-revision) |
| A document module duplicated as a class | Its `.cls` was imported instead of pasted ([step 5](#%EF%B8%8F-importing-vba-into-excel)) |
| Excel left with screen updating, events or calculation changed after an error | Read the operation's cleanup outcome ([house style](docs/VBA_HOUSE_STYLE.md#operation-and-cleanup-outcomes)); never reset settings to assumed defaults |
| Results marked `STALE` or `FAILED` | Read the findings and the log; the [workbook contract](docs/methodology/WORKBOOK_CONTRACT.md#status-and-staleness) says what each status means |
| A result differs from its expected value | Check the commit, the inputs and the tolerance of the [test case](docs/methodology/TEST_CASES.md) before changing code |

If Excel's state is uncertain, save your data separately, close Excel, and
reproduce the problem in a fresh session with synthetic data before changing
code.

## 🗑️ Removing the workbook's components

1. Run any documented cleanup, so that the workbook leaves Excel's settings as
   it found them.
2. Remove every imported module, class and form, and any test or example
   module.
3. Delete the code pasted into `ThisWorkbook` and each sheet's document module
   ([step 5](#%EF%B8%8F-importing-vba-into-excel)), keeping the components
   themselves. Otherwise their event handlers stay active and call
   procedures that no longer exist.
4. Compile what remains, then reopen the workbook and check that it opens with
   no error.

Removing components does not remove:

- data on worksheets, defined names or sheets the build created;
- references added by hand;
- trusted locations, macro settings or anything else outside the workbook.

Remove only what this workbook created, and record anything you keep.

## 🧹 Git hygiene

These commands check the line-ending, binary and ignore rules on probe paths.
The paths are probes; the files need not exist.

```sh
git check-attr text eol diff merge -- src/core/CORE_Example.bas src/forms/frmExample.frx
git check-ignore -v --no-index test-results/run.log sample.xlsm
```

Expected results:

- `.bas` is `text` with `eol: crlf` and `diff: vba`;
- `.frx` is binary, with `text` unset;
- both probe files are reported as ignored.

To check that the stored files follow the line-ending rules, start from a
clean checkout and run:

```sh
git add --renormalize .
git diff --cached --name-only
```

The second command prints nothing when every file follows the rules. A file it
lists would be normalized: inspect it, then either commit the normalization as
a change of its own or undo the staging with `git restore --staged .`. Before
any commit, stage only the intended changes and run `git diff --cached --check`.

Ignore rules do not remove files that Git already tracks, and they are not a
secret scanner.

To show procedure names in the hunk headers of VBA diffs, configure each
clone once with the `diff.vba.xfuncname` setting given in `.gitattributes`.
