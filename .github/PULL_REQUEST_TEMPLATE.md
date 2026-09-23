<!--
  Keep this pull request focused on one coherent outcome.
  Complete every common section. Delete optional profile blocks that do not apply.
  Use NOT RUN or NOT APPLICABLE with a reason; never manufacture PASS evidence.
  Record only checks and environments exercised against the exact candidate.
  Report vulnerabilities privately through SECURITY.md; do not disclose secrets,
  exploitable details, confidential workbooks, or restricted data in a pull request.
-->

<div align="center">

# 🔀 [PROJECT_NAME] Pull Request

### [PROJECT PROFILE] · Exact evidence · Reviewable change · Honest boundaries

[![Contributing](https://img.shields.io/badge/guide-CONTRIBUTING-217346?style=flat-square)](../CONTRIBUTING.md)
[![Security](https://img.shields.io/badge/security-private%20reporting-d73a49?style=flat-square)](../SECURITY.md)
[![Release](https://img.shields.io/badge/release-RELEASING-6f42c1?style=flat-square)](../RELEASING.md)
[![Changelog](https://img.shields.io/badge/changes-Unreleased-d97706?style=flat-square)](../CHANGELOG.md)

</div>

---

> [!IMPORTANT]
> Replace every square-bracket field and delete every profile block that the project cannot exercise before the first pull request.

## 📌 Summary

<!-- State the observable outcome and why it is needed. Prefer one precise purpose. -->

## 🔗 Related issues

```text
Closes #
Related to #
```

Use a closing keyword only when this pull request satisfies the issue's complete acceptance criteria.

## 🧭 Change classification

- [ ] Defect correction
- [ ] Backward-compatible capability
- [ ] Breaking API, behavior, deployment, or migration change
- [ ] Internal refactor with no intended supported-behavior change
- [ ] Test, fixture, reference-data, or validation change
- [ ] Performance change
- [ ] Security or trust-boundary hardening
- [ ] Documentation-only change
- [ ] Repository tooling, workflow, or governance change
- [ ] Packaging or release preparation
- [ ] [PROJECT-SPECIFIC CHANGE TYPE]

## 🎚️ Affected surface

- [ ] [PUBLIC API OR USER SURFACE]
- [ ] [INTERNAL OR CORE ENGINE]
- [ ] [EXCEL HOST OR UI INTEGRATION]
- [ ] [TEST OR EVIDENCE SYSTEM]
- [ ] No runtime or supported surface — documentation/repository-only

---

## 📐 Scope and contract impact

### In scope

- <!-- Deliberate outcome -->

### Out of scope

- <!-- Reasonable adjacent work deliberately deferred -->

### Supported behavior and compatibility

```text
Supported behavior changed:       Yes / No
Backward compatible:              Yes / No / Uncertain
Suggested release impact:         none / patch / minor / major / uncertain
New supported members:
Removed or renamed members:
Changed signatures or defaults:
Changed results, errors, state, or side effects:
Migration required:
Known limitation introduced or retained:
```

Assess compatibility against documented behavior, not merely the VBA `Public` keyword. Infrastructure callbacks, Ribbon entry points, test seams, and `Application.Run` targets are not automatically supported API.

### Production source and package

[AUTHORITATIVE PRODUCTION SOURCE MANIFEST OR LINK]

- [ ] Required source files and import order are unchanged.
- [ ] Required source files or order changed and `INSTALLATION.md` was updated.
- [ ] No production source/package impact.

## 🔧 Implementation notes

### Dependency update evidence

For dependency, runtime or workflow-pin changes, follow
[`docs/DEPENDENCY_UPDATES.md`](../docs/DEPENDENCY_UPDATES.md).
Repeat this block per dependency; otherwise write `NOT APPLICABLE` with a reason.

```text
Dependency and every affected file:
Old full SHA / exact version / archive digest:
New full SHA / exact version / archive digest:
Audited semantic release / interface revision:
Official release and immutable source URLs:
Tag-to-commit / asset-digest verification, date and reviewer:
Release-note, source/bundle and advisory review:
Ownership, maintainer or release-source changes:
Permissions, secrets, triggers, runner and installation changes:
Reason for update; grouping and compatibility rationale:
Candidate SHA and all applicable hosted/profile/specialist evidence:
Previous passing repository commit and run:
Rollback dependency SHA / version / digest and coupled changes:
Rollback verification or limitations:
Maintainer approval bound to final candidate SHA; manual merge:
```

Missing provenance, failed checks or unreviewed authority expansion blocks merge.
Automatic approval and merging are prohibited; the evidence review is a human
merge gate, not a result the static checker infers from this template.

```text
Approach and key invariant:
Alternatives considered:
New dependency, reference, or generated input:
State ownership and cleanup:
Failure behavior:
```

Explain decisions a future reviewer cannot safely infer from the diff.

---

## ✅ Verification

### Candidate identity

| Evidence | Result |
| --- | --- |
| Exact PR HEAD SHA | <!-- Full 40-character SHA --> |
| Base branch and base SHA | <!-- Branch + full SHA --> |
| Working tree used locally | <!-- clean / dirty; explain --> |
| Source or package tested | <!-- Exact candidate source / artifact / N/A --> |

Evidence from another commit does not certify this candidate.

### Static and repository checks

- [STATIC CHECK COMMAND]
- `git diff --check`

| Check | Result / evidence |
| --- | --- |
| Hosted required checks | <!-- PASS / FAIL / NOT RUN + workflow URL --> |
| Local static command | <!-- Command + PASS / FAIL / NOT RUN --> |
| Formatting / `git diff --check` | <!-- PASS / FAIL --> |
| Machine-readable artifact | <!-- Name / URL / not produced --> |

### Excel and VBA execution

- [ ] Required and completed against the exact PR HEAD.
- [ ] Required but incomplete — reason and merge/release consequence stated.
- [ ] Not required — documentation/repository-only change with no executable or packaging impact.

Relevant entry points:

- [COMPLETE REGRESSION OR CERTIFICATION ENTRY POINT]
- [OPTIONAL UI OR MANUAL SMOKE ENTRY POINT]

| Evidence | Result |
| --- | --- |
| Tested commit SHA | <!-- Full SHA or N/A --> |
| `Debug → Compile VBAProject` | <!-- PASS / FAIL / NOT RUN / N/A --> |
| Regression/certification entry point | <!-- Exact procedure --> |
| Completion state | <!-- PASS / FAIL / INCOMPLETE / NOT RUN --> |
| Cases / assertions / failures | <!-- Counts or N/A --> |
| Skipped / cleanup outcome | <!-- Counts and state or N/A --> |
| Focused and manual checks | <!-- Scenarios + result --> |
| Evidence file or workflow | <!-- Name / URL / N/A --> |

### Validation environment

```text
Excel product, version, and build:
Office bitness:                    32-bit / 64-bit
Windows version/build:
Workbook or add-in host:
Deployment model:
[PROJECT-SPECIFIC HOST OR TOOL VERSION]
```

Record only tested environments. Source inspection does not constitute host execution, and one Office bitness does not execute the other conditional branch.

### Regression coverage

- [ ] Existing tests cover the changed success path.
- [ ] New or amended tests cover each corrected defect.
- [ ] Boundary, invalid-input, failure, fallback, and cleanup paths are covered as applicable.
- [ ] Test entry points and inventory/count metadata remain synchronized.
- [ ] Expected results come from the contract or an independent reference.
- [ ] No regression change is needed — rationale recorded below.

```text
Coverage rationale and new test names:
Unexecuted or deferred coverage:
```

---

## ⚠️ Risk, rollback, and recovery

- [ ] Low — documentation, metadata, or mechanically verified change.
- [ ] Medium — bounded runtime, tooling, or compatibility impact.
- [ ] High — numerical integrity, shared Excel state, native API, security, release, or breaking impact.

```text
Principal failure modes:
Residual risk after validation:
Rollback or revert procedure:
Excel-process, workbook, data, or artifact recovery:
Conditions that make rollback unsafe:
```

## 🔐 Security, data, and provenance

- [ ] No credential, secret, signing material, internal URL, or personal path is included.
- [ ] No client, employer, counterparty, student, personal, or restricted production data is included.
- [ ] Test data is synthetic, anonymized, or explicitly redistributable.
- [ ] External algorithms, code, datasets, and market/vendor data have attributable provenance and compatible licensing.
- [ ] Formula, command, path, callback, deserialization, and external-content injection surfaces were assessed.
- [ ] No security-sensitive detail belongs in private disclosure instead of this pull request.
- [ ] Generated evidence identifies its inputs, tool/runtime version, candidate SHA, and limitations.

```text
Security or privacy impact:
Source/data provenance:
New trust boundary:
```

## 📚 Documentation and release hygiene

- [ ] `README.md` reflects supported behavior and examples.
- [ ] `INSTALLATION.md` reflects paths, dependencies, import order, validation, upgrades, and removal.
- [ ] `CONTRIBUTING.md` reflects development and evidence requirements.
- [ ] `CHANGELOG.md` records material change under `[Unreleased]`.
- [ ] `SECURITY.md` reflects supported versions or trust boundaries.
- [ ] `RELEASING.md` reflects certification, package, provenance, or recovery changes.
- [ ] Source headers, API references, demos, Wiki pages, and counts remain synchronized.
- [ ] Version markers remain unchanged unless this is the deliberate release-stamp change.
- [ ] No documentation change is required — reason recorded below.

```text
Documentation impact:
Release, artifact, or migration impact:
```

---

## 🧩 Project-specific review

<details>
<summary><strong>🧮 Numerical or financial profile</strong></summary>

Keep only for calculations, models, dates, curves, statistics, or other accuracy-sensitive behavior.

- [ ] Parameterization, units, signs, dates, conventions, defaults, and admissible domain are explicit.
- [ ] The algorithm or derivation and its provenance are documented.
- [ ] Independent reference system, version, precision, grid, tolerance, and worst discrepancy are recorded.
- [ ] Central, boundary, tail, overflow, underflow, non-convergence, and invalid-input behavior are covered as applicable.
- [ ] Expected values were produced independently of the implementation.

</details>
<details>
<summary><strong>🪟 Excel UI or host-state profile</strong></summary>

Keep only for workbooks, UserForms, RibbonX, WinAPI, events, timers, shapes, or shared Excel state.

- [ ] Caller-owned Excel state is recorded before mutation and restored on success and failure.
- [ ] Every registration, callback, timer, hook, form, or shape has an ownership and teardown rule.
- [ ] 32-bit and 64-bit declarations and supported host boundaries are assessed.
- [ ] Multi-window, reset, cancellation, partial-success, and recovery paths are covered as applicable.
- [ ] Interactive checks name the exact scenario and environment.

</details>
<details>
<summary><strong>⚡ Performance profile</strong></summary>

Keep only when the pull request changes performance or makes a performance claim.

- [ ] Correctness and numerical stability remain primary.
- [ ] Workload, warm-up, timer, sample count, baseline, dispersion, and host state are recorded.
- [ ] The measured candidate and comparison baseline are identified by commit.
- [ ] The result reports magnitude and uncertainty, not only “faster” or “slower”.

</details>
<details>
<summary><strong>📦 Packaging and release profile</strong></summary>

Keep only for release preparation, workbook/add-in packaging, provenance, or distribution.

- [ ] Artifact inputs come from the exact candidate source.
- [ ] The packaged artifact was reopened and tested after its final save.
- [ ] Filename, size, SHA-256, build environment, and candidate commit are recorded.
- [ ] Version, changelog, tag, release notes, and installation guidance agree.
- [ ] No artifact was modified after hashing or silently substituted.

</details>

---

## 👀 Reviewer focus

```text
Highest-risk decision:
Files and procedures to inspect first:
Evidence to challenge:
Known boundary not proved by this pull request:
Unresolved question or accepted trade-off:
```

## ☑️ Final author check

- [ ] The title describes the observable outcome.
- [ ] The pull request has one coherent purpose and no unrelated churn.
- [ ] Linked issue acceptance criteria are met or remaining work is explicit.
- [ ] Compatibility and release impact are assessed.
- [ ] Evidence belongs to the exact candidate claimed.
- [ ] Required checks are terminal and passing; incomplete work is not presented as PASS.
- [ ] Executable VBA was compiled and tested when required.
- [ ] Failure, cleanup, and recovery behavior were reviewed.
- [ ] The complete diff, including comments, metadata, binary companions, and documentation, was reviewed.
- [ ] No merge marker, stale placeholder, unexplained N/A, accidental binary, or private material remains.

---

**Review principle:** approve the smallest coherent change whose contract, evidence, risk, and recovery can all be explained from this pull request.
