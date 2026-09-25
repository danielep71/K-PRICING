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
`KPR_Tests_RunHost`, `KPR_Tests_RunShape`, and `KPR_Tests_RunArray`. K-PRICING
also adds `KPR_Tests_RunEvidence`, a destination-only adapter that delegates to
`KPR_Tests_RunAll("all")` and emits the structured retained-evidence log format.

The source expectations and condition identifiers are preserved. Historical KPR
run counts are not destination certification. The authoritative
[regression/parity protocol](../docs/MIGRATION_REGRESSION.md) documents the
entry-point roles, existing coverage and gaps, exact source-versus-destination
comparison and the additional host/shape/array logs required by #17.

## Independent generated fixtures

`fixtures/date_layer_fixtures.tsv` holds independently generated expected
results for the 22-function date layer, computed from the contract by
`tools/gen_fixtures.py` without running or translating the VBA implementation.
`modules/KPR_Test_Fixtures_Generated.bas` replays the same cases as VBA values
through `KPR_Fixtures_Count`, `KPR_Fixtures_Case` and `KPR_Fixtures_SourceHash`.
It references no production module and is test infrastructure, not supported
API. Both files are generated: regenerate them with the tool and never edit
them by hand. [`fixtures/README.md`](fixtures/README.md) documents the schema;
executing the fixtures through the durable runner is issue #39.

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


`KPR_Tests_RunMigrationEvidence` is the migration-only common observation adapter used by #17.
