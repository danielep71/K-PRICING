# Native developer operations

`tools/kpr.py` coordinates existing checks and supplies portable workbook and
source comparisons. It uses Python 3.10+ and the standard library. No xlflow
code, package, formatter, assertion framework or runtime dependency is used.
The workflow inspiration is its CLI, inspection and diff concepts, described
in the [xlflow README](https://github.com/harumiWeb/xlflow); the implementation
and contracts here are independent and deliberately smaller.

## Current scope and delivery order

The initial portable implementation was prepared against `b415175` on 1 October
2026 and reconciled with `dab4e5c` on 9 October. The demo builder (#46, PR #75)
is now complete; PR #78 hardened the runner's status-bar restoration, PR #79
renamed the demo modules and PR #81 adopted MPL-2.0. Those changes are preserved.
The authoritative [roadmap](ROADMAP.md) and remaining dependency order are unchanged.

| Capability | Existing authority | This increment / remaining work |
| --- | --- | --- |
| Deterministic demo (#46, complete) | `KPR_Demo__Builder`, `KPR_Demo_Dates` and the `worksheet-demo` suite | Add saved-content comparison to the existing builder/fingerprint tests; a fresh final-candidate host run is still required by #52 |
| Staged build (#47, #52) | Host inventory, release provenance, package gates | Reuse candidate inventory and compare supplied VBE exports; Excel staging, Ribbon injection, save/reopen, smoke and promotion remain unimplemented |
| Command interface (#50) | Existing focused validators | Thin `check`, `doctor`, `inventory`, `round-trip`, `snapshot`, `compare` commands |
| Diagnostics | Host evidence environment contract | Read-only Git/OS/output-path observations; Excel, references, VBProject, locks and live state remain unknown |
| Recovery | Existing host failure outcomes | No execution adapter in this increment; recovery procedure below is a requirement for its future implementation |
| Scriptable actions (#46–#49) | Existing `KPR_Demo_BuildDates(OutputPath)` and `KPR_Demo_LastReport`; remaining UI acceptance criteria | Reuse the existing explicit-path operation; Ribbon/CommandBars and final classification remain open |

These tools complement the completed #46 and support #50 and #52 without
closing either release issue. #47,
#48 and #49 remain prerequisites for the final release inventory. #51 still
owns candidate assembly; `VERSION` is unchanged. Optional automation does not
become an additional dependency for the current functional release: manual
exact-source certification remains available.

The separate `tools/dev-module-import` branch was reviewed at `576e72c` during
reconciliation. Its `KPR_Dev_ImportModules` loader mutates a selected open VBA
project and does not save it; these portable tools only inspect committed
source and supplied exports. This change neither merges that branch nor adds
a second importer. A later host adapter may reuse reviewed import tooling,
but must independently establish the owned staging workbook, exact-source
alignment, compilation, save/reopen, promotion and recovery controls below.

## Commands and outcomes

Run from the checkout; global options precede the subcommand:

```bash
python3 tools/kpr.py --help
python3 tools/kpr.py check
python3 tools/kpr.py doctor --output-path ../kpr-runs/demo-1.xlsx
python3 tools/kpr.py inventory --candidate-sha FULL_CANDIDATE_SHA
python3 tools/kpr.py round-trip --candidate-sha FULL_CANDIDATE_SHA --exports ../kpr-runs/exports
python3 tools/kpr.py --output ../kpr-runs/demo-1.json snapshot ../kpr-runs/demo-1.xlsx
python3 tools/kpr.py compare ../kpr-runs/demo-1.xlsx ../kpr-runs/demo-2.xlsx
python3 tools/test_native_tools.py -v
```

`FULL_CANDIDATE_SHA` is the actual full lowercase commit ID, not a branch or tag.
Create the external report directory first. Every operation prints JSON; help
and argparse usage errors use the normal CLI text. `--output` creates a **new**
JSON file exclusively. An existing file, missing parent or denied write returns
2 without replacing earlier evidence. No overwrite switch is provided. A new
report interrupted during its write may be partial; preserve it and use a fresh
path. These are supplementary observations, not release-evidence records.
CLI reports also identify the Python version and hashes of the executing
entry-point and snapshot-reader scripts.

| Exit | Meaning |
| --- | --- |
| 0 | Requested observation/comparison completed, or all selected source gates passed |
| 1 | Workbook/export differences, or an authoritative source gate reported findings |
| 2 | Invalid input, inaccessible output or an operation/gate could not complete |

An exit 0 from `doctor` means observations were collected, **not preflight
approval**. `check` runs a documented convenience subset of source gates;
the full inventory remains in `static-checks.yml`. It forwards each validator's
exit code and output, and never reimplements its policy. Interrupted child
checks are errors. No command launches Excel, executes a macro, imports code,
compiles, changes trust settings or promotes an Office artifact.

`doctor` distinguishes tracked/untracked Git changes from the unobserved saved
workbook and live session. Output access is an `os.access` hint only; it does
not create a probe, test exclusive access or prove that a future write will
succeed. On Linux, local Windows Excel automation is `UNAVAILABLE`; on Windows
it is `NOT_TESTED`. Version/build/bitness, references, VBA-project access,
workbook locks, unsaved session changes and workbook/source alignment are
unknown until an eligible host adapter or operator observes them.

## Workbook snapshot contract, version 1

The reader accepts saved transitional OOXML worksheet workbooks (`.xlsx` or
`.xlsm` content). It is not an Excel file-format validator, calculation engine,
`.xls`/`.xlsb` reader or package smoke test. Non-worksheet sheets and unsupported
namespaces fail closed. ZIP members are read in memory, never extracted;
unsafe paths, duplicate members, encryption, DTD/entities and oversized packages
are refused. The expanded package limit is 64 MiB; each XML part also uses the
existing repository XML size limit.

The machine-readable record has `schema: kpr-workbook-snapshot` and
`schema_version: 1`, with three distinct sections:

- `structure`: sheet order/names/visibility; workbook properties, including the
  date system and calculation settings; defined names; cell formulas, attributes
  and fixed inputs; array/shared-formula attributes and ranges; row/column
  settings, formatting, merges, validations, styles and retained package parts.
- `cached_results`: saved values of formula cells and members of declared
  array ranges. Blank/missing caches, Boolean/numeric/text types, original error
  tokens and known native error codes remain distinct. Numeric XML values are
  retained lexically, without rounding or tolerance. Unknown error tokens are
  retained with a null native-code mapping. Array references are preserved, not
  flattened. Missing spill cells are not invented.
- `observation`: calculation is always `NOT_OBSERVED`, Excel environment is
  null, and excluded parts are listed. This reader cannot know whether a cache
  came from Excel, is current or was calculated against the candidate.

The comparison record separately exposes `structural_match`,
`cached_results_match`, and JSON-pointer differences naming sheet, cell or
property. Its overall `status` fails if either section differs. Even equal
caches have `calculated_result_comparison: NOT_RUN`. Comparing two saved files
does not establish that the demo was generated twice.

### Explicit normalization boundary

Only these incidental representations are normalized or excluded:

1. ZIP order, compression and timestamps are ignored by comparing content.
2. XML namespace prefixes, attribute ordering and indentation between elements
   are normalized. Leaf text, formula spelling, significant tail text and child
   order are retained.
3. Shared strings are resolved into their rich-text content; inline and shared
   cell text use the same representation. Relationship wiring is still compared.
4. `docProps/` metadata and `xl/calcChain.xml` are excluded. References to a
   calculation-chain part in relationships/content types are still compared
   conservatively; there is no general relationship-ID rewrite.
5. `xl/vbaProject.bin` is excluded from logical comparison: supplied **text
   exports** must be checked separately. Equal workbook content never proves
   equal VBA. Opaque auxiliary parts use exact digests and XML auxiliary parts
   use structural content; neither is an Office-file hash equivalence test.

Nothing else is masked: cell style indexes, style tables, relationship IDs,
sheet IDs, workbook views, calculation IDs, formula whitespace, external links
and paths can produce conservative differences. Review those differences; do
not expand normalization merely to make a run pass. Shared-formula layouts are
preserved, not algebraically expanded. This deliberately narrow comparison is
suited to repeated output of the same builder, not arbitrary equivalent files.

## Source verification contract, version 1

`inventory` requires HEAD equal to the supplied SHA and a clean checkout,
including nonignored untracked files. It calls the existing
`check_excel_evidence.source_inventory` against the **committed** repository
profile with the explicit `include_examples=True` option. The native commands
include **all** configured components and applicable `.frx` companions, including
`KPR_Demo__Builder`, `KPR_Demo_Dates` and `KPR_Demo_DirectVBA` (12 components at
the reconciliation baseline). The same helper's default still excludes examples
for the historical host-evidence contract; existing records and validators do
not change. The report identifies its all-components scope. It adds component identity, role and normalized
hashes, and records Python/Git/OS versions. Missing candidate components,
incorrect `VB_Name` and duplicate case-insensitive identities are errors.

`round-trip` expects one flat directory containing exactly the candidate's
export filenames (and resources). Export only imported candidate components;
record the staging workbook's default document modules separately. Missing,
extra, stale and differently named components fail. Text must remain ASCII;
only CRLF versus LF is normalized. Comments, whitespace, final newline, casing,
attributes and declarations are compared exactly; resources remain byte-exact.
Neither command modifies source or a workbook. Candidate cleanliness is checked
again after collection.

A supplied SHA is metadata. Even matching exports do not authenticate their
origin: `workbook_alignment` remains `NOT_OBSERVED`. An operator/adapter must
retain the association with the owned workbook, its saved state, export time
and process. This report supplements the existing
[test evidence](KPR_TEST_EVIDENCE.md), [host evidence](EXCEL_EVIDENCE.md) and
[release provenance](RELEASE_PROVENANCE.md); their schema versions and statuses
are not changed, and these tools do not manufacture certification outcomes.

## Remaining Windows implementation and validation

Use the approved interactive Windows 64-bit Excel host. Other environments,
including 32-bit Office, remain untested unless independently executed. Before
claiming a complete deterministic build workflow, implement and validate:

1. **Staging and identity.** Freeze a clean candidate, retain the inventory,
   create a disposable workbook in an owned Excel process and import that exact
   inventory. Record component names/types and resolved reference identities.
   Re-export immediately and compare. A stale or missing import must prevent
   compilation/testing. Never use or overwrite a development workbook.
2. **Compilation and tests.** Observe compilation (not a simulated click).
   Call `KPR_Test_RunAll(SourceSha, OutputFolder)` or the existing
   `KPR_Test_RunSuite` with an explicit selection and a fresh output folder.
   Validate using `check_test_evidence.py`; retain environment and caller-state
   results. Inject a compile defect and a failed assertion into disposable
   fixtures: they must retain `COMPILE_FAILED` and `TEST_FAILED` host outcomes,
   respectively, and prevent artifact promotion.
3. **Demo twice.** Reuse the existing `KPR_Demo_BuildDates(OutputPath)` with two
   explicit absolute `.xlsx` paths in the controlled host. It returns a Boolean
   and exposes diagnostics through `KPR_Demo_LastReport`; do not call the empty-path
   interactive/active-workbook mode for automation. Its existing
   `worksheet-demo` suite remains authoritative for builder behavior. For the
   final-candidate comparison, explicitly calculate, wait for calculation completion,
   record Excel version/build/bitness, locale/date system and reference inventory,
   then save and snapshot both. Bind each saved file's digest to its observed
   calculation and candidate. Check formula/input structure separately from
   freshly calculated typed results, including native errors and array shapes.
   The portable cached-result match alone cannot complete this step.
4. **UI and packaging.** Implement #47/#48 callbacks as thin delegates to the
   same operations, with observable cancellation. Apply the candidate RibbonX
   using its reviewed package injector. Save, close, reopen and smoke-test the
   package before atomic promotion to the requested output. Inject a locked
   output and failed smoke test: the previous successful artifact must survive.
   Retain build inputs, tools, environment and stage logs using existing contracts.
5. **Interrupted execution.** Persist operation identity and owned process/
   workbook before running VBA. Record `EXECUTION_TIMEOUT` without assuming VBA
   stopped. Block another run against that workbook until the owned process has
   verifiably exited or the adapter verifies idle state, reconciles saved versus
   live contents, re-exports source, and completes cleanup. Failed cleanup is
   `CLEANUP_FAILED`; interrupted/partial results remain `INCOMPLETE`. Exercise
   timeout and failed cleanup on the real host; never kill unrelated Excel.

Keep failure bundles. Do not retry into the same evidence directory or substitute
an old passing record. Fill #52's `demo_generation` and `source_round_trip`
outcomes only from these observed operations. Static and synthetic tests here
are portable engineering evidence, never Excel execution or compatibility proof.
