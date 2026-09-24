# v0.0.1 setup completion

Repository setup **v0.0.1 is complete and its milestone is closed**. The
subsequent documentation audit temporarily reopened issues #3, #6 and #10;
all three were closed again after PR #23 merged at
`715eda63365ba4a94e619976acc4721411c875f0` (tree
`b59908b04e2c3ef333c2d2fbac0e724579b32eb3`) and
[merged-main CI passed](https://github.com/danielep71/K-PRICING/actions/runs/35915390210).
Issue #10 closed on 2026-09-23 at 20:22:54 UTC and retains the
[acceptance and milestone read-back](https://github.com/danielep71/K-PRICING/issues/10).

PR #24 subsequently fixed a Python fixture's locale-dependent text read at
`2e2d3dba198c1641468af277553dea13a42d6411`, with
[main CI passing](https://github.com/danielep71/K-PRICING/actions/runs/35918139650).
These documentation/tooling changes do not change the runtime-tested candidate
below. This is a milestone checkpoint, not a published release: VERSION remains
0.0.0 and no release tag was created. v0.0.2 owns migration execution.

## Accepted candidate and host

Tested source: `c8f0b6ee147d07549484ec83743f5b7dfc0ea1f2`, tree
`90c3b4680539fb18bf5b645c0afa3a6164a15f1c`. The closeout commit adds evidence and
documentation only; the runtime claim remains bound to the tested candidate.

Daniele Penza confirmed the SHA, a fresh workbook, successful compilation and
normal workbook closure. The regression reports four cases, six assertions,
zero failures, COMPLETE and cleanup PASS with unchanged Excel state. The
consumer example reports `ProjectRatio(12, 4) = 3`.

Host: Excel for Microsoft 365 MSO Version 2608, Build 16.0.20326.20072, 64-bit;
Windows (64-bit) NT 10.00; VBA7+. Checked references were Visual Basic For
Applications, Microsoft Excel 16.0 Object Library, OLE Automation and Microsoft
Office 16.0 Object Library. The supplied Trust Center screenshot shows Disable
VBA macros with notification selected (options grayed out) and VBA-project
object-model access checked. These existing settings were recorded, not changed.

The reported session time was approximately 2026-09-23 21:10 Europe/Rome.
The schema's two timestamps use that same reported minute; no measured duration
or separately observed start/end times are claimed. Exact Windows build was not
separately captured. This limitation is preserved in the session record.

## Retained evidence

- [Manual host record](../evidence/setup-2026-09-23/host.json)
- [Supplied regression output](../evidence/setup-2026-09-23/harness.txt)
- [Supplied example output](../evidence/setup-2026-09-23/example.txt)
- [Session observations and precision limits](../evidence/setup-2026-09-23/session.txt)
- [Evidence validation result](../evidence/setup-2026-09-23/excel-validation.md)
- [Machine-readable validation](../evidence/setup-2026-09-23/excel-validation.json)

The candidate's existing validator returned PASS from a clean checkout of the
exact tested SHA, checking the canonical Git-byte source inventory, retained log
hashes and expected harness results. Logs were transcribed from the maintainer's
provided text. Validation checks consistency and bindings; it does not independently
authenticate or repeat the Excel execution. Screenshot license/session identifiers
are not retained.

## Repository checks and remaining limits

Tested-candidate repository integrity run `35830974147` passed. Live label drift
run `35830974169` passed with zero differences. External-link observation run
`35830974162` found 89 OK and 33 access-restricted observations, with zero
deterministic public defects; that workflow remains non-green. CodeQL and
Scorecard were skipped under the private-repository policy.

[Initialization status](INITIALIZATION_STATUS.md) owns initialization provenance.
[Setup verification](SETUP_VERIFICATION.md) owns control read-back and upstream
lessons, including the public-visibility recheck
on 2026-09-24. The private-plan limitations above describe the historical setup
checkpoint. Daniele Penza owns remaining functional-release prerequisites.

This validates only the neutral starter on one Windows/64-bit Office host.
32-bit Office, migrated KPR parity, pricing accuracy, application lifecycle and
distributable packaging are not certified. Migration is prepared in
[MIGRATION_PLAN.md](MIGRATION_PLAN.md) and milestone v0.0.2 issues 11–18.

The exact `evidence/setup-2026-09-23/host.json` path is excluded from the branding
scan because its validated schema must retain the adopted template repository
identity. This is a provenance-record exception, not a directory-wide exclusion;
the identity regression test binds the complete three-path allowlist.
