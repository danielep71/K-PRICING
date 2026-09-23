# 🪟 Windows and Excel Evidence

[![Execution: optional](https://img.shields.io/badge/execution-optional-6f42c1)](#-eligibility)
[![Binding: exact SHA](https://img.shields.io/badge/binding-exact%20SHA-217346)](#-job-interface)
[![Fallback: manual](https://img.shields.io/badge/fallback-manual-1D76DB)](#-manual-fallback)

This documented job interface lets a project retain Office-host evidence from
an eligible automated environment or an interactive manual session. It does
not install a self-hosted workflow, change macro trust, or execute Excel from
the portable CI gates. `tools/check_excel_evidence.py` validates the records.
The [base release evidence contract](RELEASE_EVIDENCE.md) still requires actual
compile and regression evidence for every release.

## 🛡️ Eligibility

An automated host must have a licensed Excel installation, an approved
interactive Windows session, an isolated disposable workspace, a named owner,
and a reviewed adapter for that project's imports and deterministic harness.
Approve access to the exact candidate before it reaches the workstation. Use a
separate restricted validation repository or equivalent controlled execution
boundary; do not attach a privileged workstation to public pull-request jobs.
GitHub documents the [risks of self-hosted runners in public repositories](https://docs.github.com/en/actions/reference/security/secure-use).

Microsoft documents limitations of [unattended non-interactive Office automation](https://support.microsoft.com/en-us/visio/considerations-for-server-side-automation-of-office).
Installing Excel on a service account or selecting a Windows runner label does
not establish eligibility. If a reviewed interactive adapter is unavailable,
use the manual procedure below. No automated Excel run was performed to validate
this interface; the committed tests use synthetic records.

Record the existing macro policy and VBA-project-object-model access setting.
The job must not change either setting, disable protections, or add a trusted
location. Programmatic import requires preapproved access; otherwise import
manually. Use no repository write token or publishing key. Keep any required
adapter credentials outside logs and limit them to the isolated host operation.

## 📥 Job Interface

| Input | Contract |
| --- | --- |
| Repository and candidate | Reviewed repository identity and full 40-character commit SHA; no branch/tag-only input |
| Harness policy | Candidate-committed `.github/excel-evidence-policy.json` |
| Import inventory | All candidate-configured VBA components except examples, plus tracked `.frx` companions; sorted paths and SHA-256 of exact Git bytes |
| Adapter | Reviewed procedure/version capable of the project's module, document-module, form, reference and host-lifecycle requirements |
| Output directory | Fresh external directory for `host.json` and UTF-8 logs; no previous results reused |
| Timeout | Explicit finite session timeout chosen for the harness; timeout is non-green and cleanup must still be attempted |

The policy declares the entry point, ordered case IDs, exact expected assertion
count and expected-error case IDs. Change it with the harness, before freezing
the candidate. An empty expected-error list is permitted only when the suite
has no expected-error cases; document that design in the project tests.

The implementation order for an adapter is:

1. Check out the exact candidate in an isolated workspace; verify
   `git rev-parse HEAD` equals the input and `git diff --quiet CANDIDATE_SHA --`
   succeeds. Exclude unrelated untracked sources from the import workspace.
2. Compute the import inventory from that candidate. Import standard/class
   modules and forms with their resources into a fresh test workbook. Populate
   workbook/worksheet document modules in their corresponding host objects;
   do not import them as ordinary classes. Record reference resolution and any
   host-specific setup. The neutral starter imports core, facade and tests.
3. Compile with an adapter capable of observing the actual compile outcome.
   Do not equate a successful import or a button click with compilation. A
   complete successful harness run also demonstrates that its executed VBA
   compiled; record that basis without inventing a separate compile run.
4. Execute the exact policy entry point, retain its entire Immediate-window or
   equivalent log, and collect case/assertion/failure counts and expected-error
   results. Import or compile failure prevents regression execution.
5. Always perform cleanup. Record the harness cleanup result and restore any
   adapter-owned state, close only its test workbook/owned Excel process, and
   release handles. Never kill unrelated Excel processes. A successful harness
   followed by failed adapter cleanup is `CLEANUP_FAILED`.
6. Write the record after cleanup, compute each retained log's SHA-256, and
   run the validator. Preserve non-green records and logs too. Reject any
   changed candidate tree at the end. Never translate `NOT_RUN`, `TIMEOUT` or
   `UNAVAILABLE` into PASS.

The adapter must implement these operations for its host and certify its own
failure controls before use. This template supplies the interface and validator,
not a generic COM controller or a claimed automated compile implementation.

## 🧾 Shared Record Schema

Automated and manual execution use exactly the same top-level fields. This
example is a shape guide, not usable runtime evidence; replace illustrative
values, all digest markers, environment and timestamps with observed values:

```json
{
  "schema_version": 1,
  "repository": "owner/project",
  "candidate_sha": "0123456789abcdef0123456789abcdef01234567",
  "template_contract": {"version": "1.2.0", "source": "owner/template"},
  "execution": "manual",
  "availability_reason": null,
  "started_at": "2026-09-09T09:00:00Z",
  "finished_at": "2026-09-09T09:01:00Z",
  "runner": {"class": "manual-interactive", "identity": "Operator and workstation identifier", "workflow": null},
  "environment": {
    "excel_version": "16.0",
    "excel_build": "REPLACE_WITH_OBSERVED_BUILD",
    "office_bitness": "64-bit",
    "os": "Windows 10; observed build",
    "os_architecture": "x64",
    "runtime": "VBA7+",
    "macro_policy": "Describe the existing approved macro policy",
    "vba_project_access": "Disabled; manual import",
    "trust_changes": false
  },
  "sources": [
    {"path": "src/core/ProjectCore.bas", "sha256": "REPLACE_WITH_SOURCE_DIGEST"},
    {"path": "src/modules/ProjectFacade.bas", "sha256": "REPLACE_WITH_SOURCE_DIGEST"},
    {"path": "tests/modules/ProjectTests.bas", "sha256": "REPLACE_WITH_SOURCE_DIGEST"}
  ],
  "stages": {
    "import": {"status": "PASS", "detail": "Imported exact inventory into fresh test project", "log": {"path": "session.log", "sha256": "REPLACE_WITH_LOG_DIGEST"}},
    "compile": {"status": "PASS", "detail": "Record observed compile outcome and basis", "log": {"path": "session.log", "sha256": "REPLACE_WITH_LOG_DIGEST"}},
    "regression": {"status": "PASS", "detail": "Complete deterministic harness", "log": {"path": "harness.log", "sha256": "REPLACE_WITH_LOG_DIGEST"}},
    "cleanup": {"status": "PASS", "detail": "Harness state verified and owned test workbook closed", "log": {"path": "session.log", "sha256": "REPLACE_WITH_LOG_DIGEST"}}
  },
  "harness": {
    "entry_point": "ProjectTests.RunProjectTests",
    "cases": 4,
    "assertions": 6,
    "failures": 0,
    "completeness": "COMPLETE",
    "expected_errors": [{"case": "ratio.zero-denominator", "status": "PASS", "detail": "Number, source and description assertions passed in the complete suite"}]
  }
}
```

For automation set `execution: "automated"`, `runner.class: "trusted-interactive"`
and replace `runner.workflow: null` with this exact shape:

```json
{"repository": "owner/validation", "path": ".github/workflows/excel.yml", "sha": "0123456789abcdef0123456789abcdef01234567", "run_id": 123, "run_attempt": 1}
```

The workflow SHA identifies the exact adapter workflow revision; the top-level
candidate SHA identifies the VBA source being tested. Manual records must have
no workflow object and cannot claim the automated runner class. The validator
checks identities and bindings, not whether a service really executed that run.

Every executed stage requires a nonempty detail and retained-log reference.
`NOT_RUN` requires `log: null`. When regression did not run, `harness` is null;
otherwise keep the observed counts and one result per expected-error case, in
policy order, with status `PASS`, `FAIL` or `NOT_RUN`. For an interrupted run,
preserve observed counts and mark completeness `INCOMPLETE`.

For a passing regression, the raw log must contain exactly one ordered `CASE=`
line per policy case, one `CASES=`, `ASSERTIONS=` and `FAILURES=` line matching
the record, and one complete `RESULT=` line in the starter harness format.
An adapter for a different harness must emit that documented format while
retaining its native log as supporting evidence. The starter's complete
four-case/six-assertion PASS proves its three expected-error assertions passed;
the JSON result records that inference, not a separate instrumented observation.
Never infer an expected-error PASS from a `CASE=` line alone or an incomplete run.

## 🧑‍💻 Manual Fallback

1. Use a fresh checkout of the reviewed SHA, with no modified tracked files.
   Save the SHA and repository identity in `session.log`; use the candidate's
   policy and import inventory, not a workbook left over from an earlier test.
2. Open a fresh test workbook in an interactive Excel session. Record **File →
   Account → About Excel** version/build and bitness, Windows version/build and
   architecture, and the existing Trust Center settings. Do not change trust
   settings to make the test run. Record start time with timezone.
3. In the VBA editor, import the exact source and test modules. Review references,
   compile, and run `ProjectTests.RunProjectTests`. Preserve the complete
   Immediate-window report as `harness.log`. If it fails, retain the failure;
   do not replace it with the last passing report.
4. Record the compile basis, expected-error assertion result and cleanup in
   `session.log`; close the owned test workbook. Record finish time. Populate
   the shared JSON with `execution: "manual"` and `workflow: null`.
5. Validate the bundle and retain it with the release evidence. A manual PASS
   is eligible evidence but is never described as a hosted execution.

## 🚦 Outcomes and Release Binding

```bash
python3 tools/check_excel_evidence.py --root . \
  --candidate-sha FULL_CANDIDATE_SHA --evidence ../host-bundle/host.json \
  --output test-results/excel-evidence.json --summary test-results/excel-evidence.md
```

Reports must be outside the retained evidence directory. Exit 0 means the
assertions/log bindings passed validation; exit 1 means a non-passing outcome;
exit 2 means a CLI/output failure. Multiple outcomes may be retained together.

| Outcome | Meaning |
| --- | --- |
| `PASS` | Complete consistent record, passing stages and expected-error assertions |
| `IMPORT_FAILED` | Source import failed |
| `COMPILE_FAILED` | Compilation failed; regression must not claim PASS |
| `TEST_FAILED` | Harness failed |
| `CLEANUP_FAILED` | Harness or adapter cleanup failed |
| `EXECUTION_TIMEOUT` | An executed stage timed out |
| `INCOMPLETE` | At least one stage did not run |
| `UNAVAILABLE` | No eligible host execution was attempted |
| `EVIDENCE_INVALID` | Missing, contradictory, stale-source or malformed record/logs |

To record unavailability use the same identity/timestamp fields, set
`execution: "unavailable"`, give a nonempty `availability_reason`, set `runner`,
`environment` and `harness` to null, `sources` to `[]`, and `stages` to `{}`.
Do not invent environment or stage successes. This record exits 1. Availability
reporting does not waive the base release compile/regression requirement.

When adopting this interface for a release, add the following additional check
to the base evidence, replacing the digest with SHA-256 of exact `host.json`
bytes and setting execution to the actual mode:

```json
{"status": "PASS", "candidate_sha": "0123456789abcdef0123456789abcdef01234567", "detail": "Validated retained manual host bundle", "sha256": "REPLACE_WITH_HOST_JSON_DIGEST", "execution": "manual"}
```

Its key is `excel-host-evidence`. Supply `--excel-evidence ../host-bundle/host.json`
to `check_release.py`. The gate requires both the file and binding check, rejects
non-green host outcomes, and cross-checks base regression counts and entry point.
Copy the validator JSON report's `release_environment` string verbatim into
both base `vba-compile.environment` and `regression.environment`. This binds the
base claims to the same detailed environment and manual/automated designation;
an inconsistent environment is rejected.
For signed provenance, finalize this binding before hashing/signing the base
evidence. Existing release evidence without this optional check remains valid.

Keep logs free of credentials and client workbook contents. Use a short explicit
CI retention period appropriate to review (for example 30 days), then retain
the certified bundle with the durable release record before temporary artifacts
expire. Preserve failed-run logs until the finding is resolved. Treat source
hashes and log digests as identity controls: they do not prove what a workbook
contained or authenticate a human's or adapter's assertions.

## 🧪 Interface Validation

Run `python3 tools/test_excel_evidence.py -v`. Portable CI exercises synthetic
manual and automated records, exact-source/log bindings, expected-error results,
and distinct failure/unavailable outcomes. A real host adapter must separately
demonstrate import, compilation, regression and cleanup failure controls before
its automated evidence is trusted. This documented interface can be delivered
without certifying an Excel runner that is not available.
