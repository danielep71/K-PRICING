# 🏅 Release certification

The record that one candidate commit was checked, built, run and accepted
before a version is published. Copy the template below into an issue for the
release, filed under the milestone after the released one
([release records](../../RELEASING.md#release-records)), fill it in from the
runs themselves, and close the issue only when every row holds evidence or a
recorded deviation. [RELEASING.md](../../RELEASING.md)
says when each part is due.

Write `NOT RUN` with the reason for anything that was not done. A row left
empty or marked `NOT RUN` is never a pass.

## Template

```markdown
## Release certification: vX.Y.Z

### Candidate

| Item | Value |
| --- | --- |
| Version | X.Y.Z |
| Candidate commit | Full 40-character SHA of the integration merge commit |
| Excel run commit | The candidate commit, or the full SHA of the earlier commit the workbook was built from, with the report's `Equivalent: CANDIDATE_SHA (N run inputs identical to RUN_SHA)` line ([RELEASING.md step 4](../../RELEASING.md#4-certify-the-workbook)). A report with a `Carried from` line (`--carried-to`) does not certify a release |
| Integration and sync pull requests | #N, #N |
| Changelog section | `## [X.Y.Z] - YYYY-MM-DD`, matches the candidate: yes or no |
| Clean checkout | `git status --short` printed nothing: yes or no |

### Environment

| Item | Value |
| --- | --- |
| Excel | Product, version and build, from **File → Account → About Excel** |
| Bitness | 64-bit or 32-bit |
| Windows | Edition, version and build, from `winver` |
| References | The boxes ticked in **Tools → References** for a run by hand, or the `references` line of `host.txt` for a run of `tools/devbuild/excel_run.py` |
| Python, Git | Versions used for the repository checks |

### Static gates

| Gate | Result | Evidence |
| --- | --- | --- |
| Repository integrity on the candidate | PASS or FAIL | Run link, and its `integrity-<sha>-…` artifact |
| Repository checks from a clean clone (RELEASING.md step 3) | PASS or FAIL | `report.json`: commit, `dirty: false`, status |
| CodeQL alerts open on the default branch | None, or each with its assessment | Alert links |

### Regression evidence

| Item | Result | Evidence |
| --- | --- | --- |
| Debug → Compile VBAProject | PASS, the error reported, or `ON_DEMAND`, which the check accepts only with identity by rebuild (`--workbook` and no `exported/` folder), as `tools/devbuild/excel_run.py` records it: VBA compiles only the modules the run calls, so list the modules it did not call as a limit | Host record |
| Harness certification ([recording an Excel run](../../INSTALLATION.md#recording-an-excel-run)) | `CERTIFICATION=PASS` or `FAIL`, `COMPLETE` or `INCOMPLETE` | `certification.log` |
| Workbook built from the Excel run commit (`tools/devbuild/build_workbook.py`) | Its `build_commit` stamp, which must be the Excel run commit without *NOT evidence-grade* | For a run by hand, the workbook's SHA-256 in `export-manifest.json`; for a run of `tools/devbuild/excel_run.py`, `Binding: VERIFIED` by rebuild in the check report, with the SHA-256 of `test-results/build/Workbook-development.xlsm` |
| `tools/check_excel_run.py --workbook … --require-binding` against the candidate, or against the Excel run commit with `--equivalent-to` the candidate | `Binding: VERIFIED` and `excel-run: PASS`, with the `Equivalent:` line for an earlier commit, or each `FAIL` line | Its report, with the Git blob id of every component |
| Excel evidence gate: `tools/release.py evidence CANDIDATE_SHA --excel-evidence … --workbook …`, with `--run-commit` the Excel run commit when it is not the candidate ([RELEASING.md step 4](../../RELEASING.md#excel-evidence-gate)) | `release-evidence: PASS`, or each `FAIL` line; carried evidence is never accepted | Its `Excel evidence digest:` and `Tag message line:` lines; the tag message carries the second |
| Accuracy contracts, measured on `results.json` and the results file of each project suite that names one | Each contract's worst calibration and holdout error against its threshold, which blocks, and its target, if any, which is only reported, with each `target missed`, the number of exposures of each holdout, each holdout sealed because its contract is draft, and what each invariant checked and its reversals within the allowance, from the report's `Accuracy:` lines; `frozen` or `draft`. A frozen contract's holdout is measured only once its first exposure is recorded ([RELEASING.md step 4](../../RELEASING.md#4-certify-the-workbook)) | The report, those results files, and `tests/accuracy/holdout_exposures.json` at the candidate |
| Regression cases (validation plan) | Number run, passed, failed | Log or evidence folder |
| Failure paths: invalid input, cleanup after an error | PASS or FAIL | Log lines |

### Manual validation

What was checked by hand, by whom, against which independent values. Write
`None` if nothing was.

### Artifacts

| File | SHA-256 | Built from |
| --- | --- | --- |
| `<repository>-X.Y.Z-source.tar.gz` and `SHA256SUMS`, from the **Release** workflow, with their provenance attestation, or a deviation that names its absence | | The tag |
| `<repository>-X.Y.Z-user-manual.pdf`, or None when the repository keeps no [user manual](../../RELEASING.md#user-manual) | | `docs/user-manual/X.Y.Z/` at the candidate, with the PDF tool and its version |
| None, or each other file attached to the GitHub Release, such as the production workbook | | Candidate commit |

### Published release

| Check | Result | Evidence |
| --- | --- | --- |
| `tools/release.py verify-published vX.Y.Z --repository OWNER/REPOSITORY --milestone N`, with an `--asset` for each other file of *Artifacts* (RELEASING.md step 7) | `release-verify-published: PASS`, or each `FAIL` line; a `NOTE` on an unattested bundle is a deviation | Its output |

### Published workbook

The checks of RELEASING.md step 8, after publication. Write
`NOT APPLICABLE` when the release attaches no workbook.

| Check | Result | Evidence |
| --- | --- | --- |
| Build stamp of the attached workbook | The candidate's full SHA and `production`, without *NOT evidence-grade* | The `build_commit` and `build_configuration` values |
| Rebuild of the tag (`build_workbook.py --configuration production`) | SHA-256 equal to the attached workbook's, or every part identical | Both SHA-256 values, or the parts compared |
| Downloaded copy: SHA-256 before unblocking | Matches *Artifacts* | The value |
| **Data → Edit Links** | No link to another workbook, or each link listed | Screenshot or list |
| **Formulas → Name Manager** | The workbook contract's names, none referring to `#REF!` | Screenshot or list |
| Run, save and reopen a copy | Each operation run on synthetic data; reopens with no error, links and names unchanged | Operations run, and any message |
| User manual PDF, downloaded | Opens and names version X.Y.Z, or `NOT APPLICABLE` | Its SHA-256 against *Artifacts* |

### Deviations

`None`, or each release invariant not met, why, and my decision to accept it.
Without Excel evidence for a candidate that holds a VBA component, the tag
message lists the same `- Invariant 4: …` item under `Deviations:`.

### Limits

What this certification does not cover, such as hosts not tested,
calculations without independent reference values, rules that the validation
plan's [traceability table](../methodology/VALIDATION_PLAN.md#traceability)
does not mark `Validated`, sources not yet verified, or host behaviour still
`assumed` in its
[host behaviour register](../methodology/VALIDATION_PLAN.md#host-behaviour).

### Sign-off

- Certified by: name, date
- Distribution review: #N, closed
- Decision to publish: recorded here, with the date
```
