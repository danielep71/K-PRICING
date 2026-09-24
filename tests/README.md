# 🧪 Tests and Evidence

`tests/` is the canonical home for all verification source and stable test data.

Use subdirectories only when they contain real material:

| Location | Contents |
| --- | --- |
| `tests/modules/` | Exported VBA regression modules and release-certification entry points |
| `tests/fixtures/` | Deterministic inputs, manifests, and reusable test workbooks |
| `tests/expected/` | Reviewed expected outputs or golden files |

## Migrated KPR regression harness

Import `modules/KPR_REGRESSION_TESTS.bas` after all five production modules and
compile the VBA project. The imported harness retains these entry points:
`KPR_Tests_Run`, `KPR_Tests_RunSuite`, `KPR_Tests_RunAll`,
`KPR_Tests_RunHost`, `KPR_Tests_RunShape`, and `KPR_Tests_RunArray`.

The source expectations and condition identifiers are preserved. Historical KPR
run counts are not destination certification: exact-candidate execution,
failures/skips, environment and source-versus-destination parity are recorded
under migration issues #14 and #17.

The optional [Windows/Excel evidence interface](../docs/EXCEL_EVIDENCE.md)
records this harness output, source identity and host environment with explicit
manual/automated execution. Its validator does not execute Excel.

## Rules

- Test modules are never part of the production import set.
- Name the complete regression and release-certification entry points in `CONTRIBUTING.md` and `RELEASING.md`.
- Keep fixtures synthetic, anonymized, or explicitly redistributable.
- Bind numerical or platform-sensitive evidence to the exact candidate commit and environment.
- Do not commit transient output, logs, caches, or locally generated workbooks; use ignored output directories or workflow artifacts.
- New repositories use `tests/`, not `test/`.

A legacy `test/` directory is legitimate only when an existing public path, build script, or release contract makes migration materially disruptive. Document that exception and never keep both `test/` and `tests/`.

Delete this README only if the real harness and equivalent test documentation make the directory's role equally explicit.
