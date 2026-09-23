# v0.0.1 setup completion

The setup milestone has met its technical acceptance criteria. Final issue and
milestone closure are recorded in issues 8 and 10. This is a milestone checkpoint,
not a published release: VERSION remains 0.0.0 and no release tag is created.

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

[Setup verification](SETUP_VERIFICATION.md) owns source/template provenance,
control read-back and upstream lessons. Branch/tag rulesets remain unavailable
under the private plan; manual review is not server-enforced protection. Daniele
Penza owns resolving release prerequisites before a functional release.

This validates only the neutral starter on one Windows/64-bit Office host.
32-bit Office, migrated KPR parity, pricing accuracy, application lifecycle and
distributable packaging are not certified. Migration is prepared in
[MIGRATION_PLAN.md](MIGRATION_PLAN.md) and milestone v0.0.2 issues 11–18.

The exact `evidence/setup-2026-09-23/host.json` path is excluded from the branding
scan because its validated schema must retain the adopted template repository
identity. This is a provenance-record exception, not a directory-wide exclusion;
the identity regression test binds the complete three-path allowlist.
