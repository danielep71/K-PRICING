# ✅ Run the quality gates

> **Guide, not policy:** [tools/README.md](../../tools/README.md) is authoritative.

**Working directory:** initialized project root. Use its local operational tools.
Stage intentional new files before tracked-file checks; otherwise those files
may be absent from the inspected inventory. Review the staged set first.

## 1. Validate source contracts

```bash
git diff --check
python tools/check_committed_whitespace.py --root . --mode working-tree
python tools/check_repo.py --root .
python tools/check_vba_public_api.py --root .
python tools/check_vba_jumps.py --root .
python tools/check_vba_conditionals.py --root .
python tools/check_template_contract.py --root .
python tools/check_documentation.py --root .
python tools/check_release_semantics.py --root .
```

Expected: every applicable gate passes. The canonical gate has 21 rules; focused
gates add checks beyond that baseline. Exit 1 means findings; exit 2 generally
means the tool could not complete. Read the particular tool's report.

Use self-tests when validating a checker or reproducing its CI gate:

```bash
python tools/check_repo.py --root . --self-test
python tools/check_release.py --root . --self-test
```

A release self-test uses synthetic fixtures. It is not certification of your
actual release. Run the candidate release gate later with real evidence.

## 2. Read the hosted run

Push the reviewed branch and open a PR. In **Actions** and the PR checks, open
**Repository integrity**. Read failing steps, the job summary and retained
`static-checks-RUN_ID-RUN_ATTEMPT` artifact. The workflow also checks Python
lint/types, workflow syntax, fixture behavior and release semantics.

A PR run may check GitHub's merge candidate rather than the head branch commit.
Record the actual checked SHA. Do not assume a previous green run applies after
a new push. Required-check settings must match the actual emitted job context.

Generated projects do not retain the template's checker-development, portfolio
fixtures or wiki maintenance workflows. Their disappearance is intentional.

## 3. Run Excel and project-specific checks

Import the candidate exports and run the documented regression entry point.
The starter uses `ProjectTests.RunProjectTests`. A complete PASS includes
environment, four cases, six assertions, zero failures and cleanup PASS.
Successful execution establishes compilation of the exercised project; retain
the full output. Static checks do not execute Office.

A UI component or application needs its profile-specific lifecycle and state
evidence as well. Add numerical, performance or other specialist checks when
the supported contract needs them.

## 4. Handle links separately

The external-link workflow runs on its schedule or manual dispatch, not as a
network dependency of PR validation. It reports inaccessible/transient targets
separately from permanent failures under
[Documentation Checks](../DOCUMENTATION_CHECKS.md).

To run it deliberately from the project root, replace the UTC observation date:

```bash
python tools/check_external_links.py --root . --as-of YYYY-MM-DD --output test-results/external-links.json --summary test-results/external-links.md
```

An unavailable target is not a clean bill of health. Resolve deterministic
source defects locally; review external observations and expiring exceptions
according to their authority.

---
[Home](Home.md) · [Previous](Configure-GitHub.md) · [Next](Develop-safely.md)
