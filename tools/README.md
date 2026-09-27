# 🛠️ Tools and Validation Guide

[![Runtime: Python 3.10](https://img.shields.io/badge/runtime-Python%203.10-1D76DB)](#python-presentation-and-lint-policy)
[![Evidence: bounded](https://img.shields.io/badge/evidence-explicit%20scope-217346)](#canonical-repository-quality-gate)

`tools/` contains repository validators, evidence collectors and setup tooling.

The [local Action contract](LOCAL_ACTIONS.md) owns repository-local Action and
reusable-workflow containment checks. `test_verification_depth.py` exercises
retained validator failure paths and is required by `Repository integrity`:

```bash
python3 tools/test_verification_depth.py -v
```

Its output includes expected diagnostics from negative fixtures; the unittest
exit status determines success. These are Python tests, not Excel execution.

Local source checks are deterministic. Network collectors and live setup tools
have separately documented observation, credential and mutation boundaries.

<a id="python-presentation-and-lint-policy"></a>

## 📐 Python presentation and lint policy

[`pyproject.toml`](../pyproject.toml) sets Python 3.10 as the compatibility target
and 100 columns as a formatting target. CI runs `ruff check tools` with exactly
`E4`, `E7`, `E9`, `F`, `I001`, `C90`, and `S314`; `I001` rejects unsorted or
unformatted import blocks, and `C90` uses a McCabe ceiling of 20. CI also runs
`mypy` across `tools/`; `_gatelib` is subject to the pinned mypy 2.3.1 strict
bundle through its per-module override while the rest of the tree remains on the
established whole-tree baseline. E501 is not selected and CI does not run
`ruff format --check`; line length and formatter output are therefore not
blocking rules. Prefer readable wrapping without changing literals or churning
unrelated code. Editor indentation and line endings are defined in
[`.editorconfig`](../.editorconfig).


Comments and docstrings explain purpose, inputs, ownership, error handling and
limits when these are not clear from the code. Check them whenever behavior
changes. Distinguish synthetic fixtures, structural predicates, live observations
and executed host evidence; a token variable does not constrain granted scopes.


`check_documentation.py` checks literal documented Python commands and registered
file/workflow/policy references without executing them. `check_external_links.py`
produces separate bounded anonymous HTTP observations. Both use `run_gate`;
`test_documentation.py` supplies offline failure fixtures. See
[documentation checks](../docs/DOCUMENTATION_CHECKS.md) for policy and scope.

`release_provenance.py` extends `check_release.py` with contract 1.2.0 build
records, complete payload inventory and optional SSH verification. It has no
separate CLI. See [the provenance contract](../docs/RELEASE_PROVENANCE.md);
`test_release_provenance.py` exercises both the integrated gate and real
ephemeral signatures without Office or publishing credentials.

`check_excel_evidence.py` validates the optional
[Windows/Excel host interface](../docs/EXCEL_EVIDENCE.md), using the shared
`run_gate` runner. Its `test_excel_evidence.py` fixtures exercise synthetic
manual/automated records and non-green outcomes; neither tool runs Office.

Appropriate contents include:

- VBA static checks and exported-source validators;
- formatting and documentation-link checks;
- release-provenance and hashing utilities;
- fixture or report generators whose inputs and outputs are documented; and
- local wrappers that reproduce a CI gate.

### Shared focused-gate infrastructure

`_gatelib.py` is the private, standard-library-only owner of Git, report-output, tracked-file, and common focused-gate CLI primitives. `check_repo.py` deliberately does not import it: the canonical checker remains a self-contained distributable artifact. The canonical template also carries checker-development and semantic policy-coverage harnesses; initialization strips those maintainer-only files while retaining the operational gates needed by generated projects.

`_gatelib.run_gate` additionally owns the orchestration shared by focused gates: `--self-test` dispatch, canonical JSON serialization, Markdown summary writing, console output, and the `0` (pass) / `1` (findings) / `2` (could not complete) exit mapping. Each gate keeps its own semantic checks, fixtures, report schema, Markdown renderer and operational-exception tuple; the runner never widens exception handling, so a programming error still raises rather than being reported as exit `2`.

Focused report gates use the shared runner. Distinct CLI contracts retain their
own entry points: release validation has atomic evidence writes and separate
console rendering; workflow validation emits text-only evidence; initialization
provisions source. Disposable fixture creation and snapshot collection are
canonical maintainer tooling removed from this generated repository. The canonical checker remains self-contained.


## Canonical repository-quality gate

`check_repo.py` is the dependency-free baseline gate. Its versioned policy lives
in `.github/repository-profile.json`, where K-PRICING records `mode: generated`
and `profile: application`. Update the declared paths and VBA component roles
there instead of editing checker logic.

Run the portable commands locally:

```bash
python3 tools/check_repo.py --root . --self-test
python3 tools/check_repo.py --root . \
  --output test-results/static-checks.json \
  --summary test-results/static-checks.md
```

The first command exercises a passing fixture, one deliberately degraded
fixture for each canonical rule, malformed YAML/XML, prohibited XML DTD/entity
declarations, oversized XML, deterministic JSON and Markdown rendering, and
read-only execution.
The second command validates the current tracked tree, prints a readable result,
and writes optional machine-readable evidence.

The `generated-vba-contract` rule resolves the applicable profile contract and
requires its registered, tracked façade, core, and test assets. Its JSON evidence
records the selected profile, role minima, observed role counts, and mandatory
component paths. In template mode it validates all three supported contracts.
The `issue-forms` rule validates the three canonical intake forms, their
manifest-backed labels, required evidence fields, empty reusable assignees,
blank-issue policy and repository-specific private-security route.

Exit status `0` means every applicable rule passed, `1` means policy findings
were reported, and `2` means the checker could not complete. Reports contain no
timestamps, so identical commits and configurations produce identical bytes.

This gate validates repository evidence and exported VBA structure. It does not
execute Excel, compile a VBA project, prove numerical accuracy, exercise UI
state, or certify a release package. Profile and project gates retain those
responsibilities.

## Committed and working-tree whitespace

`check_committed_whitespace.py` separates two different Git checks that must not
be confused:

- **committed mode** is the CI/release-facing gate. It runs `git diff --check`
  over a committed candidate range. With `--base`, the range begins at the
  merge base of that revision and `--head`; without `--base`, it checks the
  first-parent delta, or the empty tree for a root commit;
- **working-tree mode** is local feedback. It checks both staged and unstaged
  changes without making those mutable files part of committed-candidate
  evidence.

Run the deterministic fixtures and local mode with:

```bash
python3 tools/check_committed_whitespace.py --root . --self-test
python3 tools/check_committed_whitespace.py --root . --mode working-tree
```

To reproduce the hosted committed check explicitly:

```bash
python3 tools/check_committed_whitespace.py \
  --root . \
  --mode committed \
  --head HEAD \
  --output test-results/committed-whitespace.json \
  --summary test-results/committed-whitespace.md
```

For pull requests, CI supplies the target-branch revision with `--base` and
records the resolved merge base, exact head SHA, inspected range, and findings
in JSON and Markdown evidence. The self-test proves that a committed
trailing-whitespace defect fails from a clean checkout, a defective root commit
fails against the empty tree, a valid commit passes, and staged/unstaged defects
remain detectable only through the local working-tree path.

## Procedure-scoped VBA jump validation

`check_vba_jumps.py` is the authoritative hardening gate for `GoTo`, `GoSub`,
and `Resume` target ownership. It parses logical VBA statements, including
continued procedure declarations, and associates every label and jump with one
owning Sub, Function, or Property procedure.

The gate:

- resolves named and numbered labels only inside the owning procedure;
- accepts equivalent label names in different procedures without collision;
- rejects duplicate labels within one procedure;
- ignores jump-like text inside strings, including escaped quotes and continuations,
  while still checking executable jumps following a string on the same line;
- treats `On Error GoTo 0`, `On Error GoTo -1`, bare `Resume`, and
  `Resume Next` as control forms rather than label references; and
- reports component, procedure, source line, operation, and unresolved target
  in deterministic JSON and Markdown evidence.

Run the focused fixtures and repository check with:

```bash
python3 tools/check_vba_jumps.py --root . --self-test
python3 tools/check_vba_jumps.py \
  --root . \
  --output test-results/vba-jumps.json \
  --summary test-results/vba-jumps.md
```

The fixture matrix proves valid local handlers, `GoSub`/`Resume`, deliberate
cross-procedure rejection, same-name labels in separate procedures, duplicate
local labels, numbered labels, line continuations, and special error-control
forms. This dedicated gate is authoritative for procedure-scoped target
resolution; the broader `vba-structure` rule remains a compatibility and
structural check. The hosted terminal verdict requires both gates, so the
broader rule cannot make a cross-procedure target green in CI.

## VBA conditional-compilation validation

`check_vba_conditionals.py` is the authoritative hardening gate for reachable
VBA `Declare` statements under the supported host model. It evaluates three
explicit environments: `vba6-win32`, `vba7-win32`, and `vba7-win64`.

The checker maintains a full nested conditional stack containing parent
activity, branch selection, current activity, and `#Else` state. It evaluates
`#If`, `#ElseIf`, `#Else`, and `#End If` consistently, so inactive descendants
cannot accidentally become active when an outer branch is false. Supported
expressions use `VBA6`, `VBA7`, `Win32`, `Win64`, Boolean literals, integer
literals, parentheses, `Not`, `And`, `Or`, `=`, and `<>`.

The boundary is deliberately conservative:

- every `Declare` reachable in either supported VBA7 environment must include
  `PtrSafe`;
- VBA6-only declarations may retain legacy syntax;
- unknown or project-defined symbols fail closed rather than being guessed;
- `#Const` is rejected because project-defined compilation constants are outside
  the reusable baseline; and
- malformed, duplicate, or unbalanced branch directives produce actionable
  diagnostics.

Run the focused fixtures and repository check with:

```bash
python3 tools/check_vba_conditionals.py --root . --self-test
python3 tools/check_vba_conditionals.py \
  --root . \
  --output test-results/vba-conditionals.json \
  --summary test-results/vba-conditionals.md
```

The fixtures cover nested VBA6/VBA7 and Win32/Win64 branches, `#ElseIf`
selection, inactive nesting, reachable non-`PtrSafe` failures in each VBA7
bitness, continued declares, unsupported symbols, and unbalanced directives.
This dedicated gate is authoritative for reachable conditional-compilation
semantics; the broader `vba-structure` rule remains a compatibility and
structural check. Both remain required in hosted CI, so the broader check cannot
hide a reachable declaration defect.

## Complete VBA public API validation

`check_vba_public_api.py` is the authoritative hardening gate for the supported
VBA surface and its checked-in `docs/PUBLIC_API.txt` manifest. Every generated
profile requires that manifest from initialization onward because the manifest
is a global required path and every profile requires a public-role component.
There is no maturity-stage exemption.

The reusable policy deliberately prohibits implicit public visibility. Supported
API declarations must use explicit `Public` visibility; `Global` is accepted for
legacy public variables. The gate normalizes line continuations and covers
public Subs, Functions, Property Get/Let/Set members, constants, events,
Declare Function/Sub members, variables, Enums, and Types. Public variable
statements contain one identifier each so signatures remain unambiguous.

`PUBLIC_API.txt` remains the single manifest. Its traditional three-column rows
preserve compatibility with the canonical repository checker, while `# SIG`
comment records bind every row to a normalized declaration signature. Those
comments include meaningful VBA distinctions such as property direction,
parameter modifiers and order, return types, Declare metadata, constant
definitions, and Enum/Type bodies. The dedicated gate detects missing, stale,
duplicate, or changed signatures and case-insensitive public-name collisions in
standard modules.

Conditional declarations are collected separately for the three supported
compilation environments described above. Mutually exclusive variants share one
three-column manifest row and require one `# SIG` record for each distinct
reachable signature. Duplicate records, missing or stale variants, and declarations
that collide in any shared environment fail. Unknown conditions fail closed;
these static models do not certify execution in Excel.

Run the focused fixtures and repository check with:

```bash
python3 tools/check_vba_public_api.py --root . --self-test
python3 tools/check_vba_public_api.py \
  --root . \
  --output test-results/vba-public-api.json \
  --summary test-results/vba-public-api.md
```

The fixtures cover every supported declaration family, continued declarations,
implicit-public rejection, signature drift, name collisions, and the
single-public-variable rule. This dedicated gate is authoritative for complete
public-surface extraction and signature binding; the broader `vba-public-api`
rule remains a compatibility check. Both are required in hosted CI, so the
compatibility view cannot hide an unsupported or unrecorded public declaration.

## KPR migrated date-layer contract

`check_kpr_contract.py` is the project-specific additive guard for the migrated
KPR date implementation. It does not replace the generic repository or complete
public-API validators. It pins the six migrated component roles and the frozen
date-layer architecture: exact 22-function facade surface, Variant return types,
core dependency direction, strict locale-independent parsing, date-window
constants, caller/date-system guard placement, volatility scope, caller-workbook
authority, array-engine purity, required in-project members, and boundary-safe
date construction.

Run the focused positive/degraded fixtures and the candidate check with:

```bash
python3 tools/check_kpr_contract.py --root . --self-test
python3 tools/check_kpr_contract.py \
  --root . \
  --output test-results/kpr-contract.json \
  --summary test-results/kpr-contract.md
```

The hosted Repository integrity workflow retains both reports and treats either
specialist-check failure as terminal. Excel compilation and runtime behavior
remain separate Excel evidence; the v0.0.2 source/destination parity run (#17)
passed and is recorded in `docs/MIGRATION_COMPLETION.md` (evidence under
`evidence/migration-2026-09-24`).

The same checker keeps the generated fixture module independent: it may not
reference any production module, and only the regression harness may depend on
it.

Its `kpr-oracle-scope` rule limits the Excel cross-oracle (`KPR_Test_Oracle`,
#41) to formulas that call only `EOMONTH`, `EDATE`, `WEEKDAY`, `DAY`, `YEAR`
and `MONTH`, and rejects `WORKDAY.INTL` and `NETWORKDAYS.INTL`. Every formula
must reach Excel through the `Xl` helper, whose calls are matched in any letter
case; no other procedure may evaluate a formula, through bracket evaluation,
`WorksheetFunction`, cell or name formulas, recalculation or a macro call.
`Xl` itself holds only `Xl = mSheet.Evaluate(Formula)`. Outside `Xl`, every bare name and
member access must be on the rule's allowlist.
Each formula is built only from literals and `CStr` of arithmetic over numeric
literals and numeric-typed variables, so it can be inspected, and its text holds
only numbers, arithmetic and the permitted functions (no names or references). The oracle may
depend on `KPR_DATES_DAYS` alone.

## Independent date-layer fixtures

`gen_fixtures.py` is the independent fixture generator for issue #38. A
reference model written from
[`docs/DATE_LAYER_CONTRACT.md`](../docs/DATE_LAYER_CONTRACT.md) computes every
expected result with Python standard-library date arithmetic; it never imports,
executes or translates the VBA implementation. The generator writes the
canonical `tests/fixtures/date_layer_fixtures.tsv` and then derives
`tests/modules/KPR_Test_Fixtures_Generated.bas` from the parsed TSV.

```bash
python3 tools/gen_fixtures.py --root . --write
python3 tools/gen_fixtures.py --root . --check
python3 tools/gen_fixtures.py --root . --self-test
```

`--check` fails when either committed file differs from a fresh generation.
`--self-test` asserts the model against worked examples quoted from the
contract, confirms deterministic output and encoding round trips, checks VBA
line limits and production independence, and proves that `--check` detects a
stale TSV and a stale VBA module. Generation also fails if a contract function
or registry condition has no fixture. Hosted Repository integrity runs
`--self-test` and `--check`. The schema, encoding and contract interpretation
notes are in [`tests/fixtures/README.md`](../tests/fixtures/README.md).

## Structured test evidence

`check_test_evidence.py` validates the `kpr-test-evidence.json` record written
by the durable VBA runner (`KPR_Test_RunAll`, `KPR_Test_RunSuite`, issue #39).
It checks the record against
[`docs/kpr-test-evidence.schema.json`](../docs/kpr-test-evidence.schema.json)
with a standard-library subset validator, reads the suite registry from
`tests/modules/KPR_REGRESSION_TESTS.bas`, binds the fixture SHA-256 and case
count to the canonical TSV, and checks counts, statuses, the result and every
certification outcome for consistency.

```bash
python3 tools/check_test_evidence.py --root . --self-test
python3 tools/check_test_evidence.py --root . --evidence ../kpr-evidence/run-1/kpr-test-evidence.json --candidate-sha FULL_CANDIDATE_SHA
```

`--compare` checks that two records differ only in their declared
nondeterministic fields, and `--certification` requires a complete #52
certification record. `--self-test` exercises synthetic records and confirms
that every schema property is written by the VBA runner. Hosted Repository
integrity runs `--self-test`. The record contract is
[`docs/KPR_TEST_EVIDENCE.md`](../docs/KPR_TEST_EVIDENCE.md).

## KPR migration evidence binding

`check_migration_evidence.py` validates the retained v0.0.2 source-versus-
destination parity bundle defined by
[`docs/MIGRATION_REGRESSION.md`](../docs/MIGRATION_REGRESSION.md). It does not
run Excel. Given a completed `migration.json` manifest, it:

- hard-binds the frozen KPR source SHA and derives the exact K-PRICING
  destination SHA from a clean tracked checkout;
- binds the exact candidate `KPR_REGRESSION_TESTS.bas` instrumentation bytes;
- requires one recorded 64-bit Windows/Excel environment for exact parity;
- hashes every required supplemental compile, observation and host-record file,
  requires canonical PASS import/compile/native-run records, invokes the
  authoritative `check_excel_evidence.py` validation logic on the exact bound
  destination host record, and cross-checks that validated record against the
  same candidate SHA and complete parity environment;
- requires the baseline `OBS` IDs and `CLEANUP` records from both common
  observation hosts while allowing additional common probe IDs; and
- rejects different source/destination observations or non-PASS
  host/shape/array cleanup while requiring known source corrections such as #32
  to be explicitly registered rather than silently waived.

Run the deterministic positive/degraded fixtures with:

```bash
python3 tools/check_migration_evidence.py --self-test
```

Validate a real retained bundle, such as the v0.0.2 bundle under
`evidence/migration-2026-09-24`:

```bash
python3 tools/check_migration_evidence.py \
  --root . \
  --manifest /path/to/evidence/migration.json \
  --output test-results/migration-evidence.json
```

The validator authenticates file bindings and parity records, requires every
stateful runner summary to report zero failures, and binds the destination to
the checked-out candidate rather than a user-supplied SHA. It does not
authenticate the human or host that produced them. The destination host-record
schema and its source/log/harness bindings are validated through the existing
`check_excel_evidence.py` authority rather than duplicated here.
`--output` must name a path outside the manifest's evidence directory; the
validator refuses to write its report into the retained bundle.

## Adopted template contract

`check_template_contract.py` owns the semantics of the template contract: the
versioned identity of the control set a repository adopts, recorded as
`template_contract` in `.github/repository-profile.json`.

The canonical checker only requires the key to exist among the configuration
keys; every rule about its content lives in this gate, so the portable checker's
policy-branch inventory is unchanged.

The gate requires exactly `version` and `source`, rejects a non-canonical or
unsupported version with a message naming the supported set, resolves the rule
set registered for the *recorded* version so an older adopter is never judged
against newer controls, requires migration notes for every supported version in
`docs/TEMPLATE_CONTRACT.md`, and checks the template/generated source
invariants: a template publishes its own contract, a generated repository names
the template it adopted and never itself.

The contract version is independent of the project `VERSION`; the self-test
proves that changing `VERSION` does not alter the adopted contract.

```bash
python3 tools/check_template_contract.py --root .
python3 tools/check_template_contract.py --root . --self-test
```

`docs/TEMPLATE_CONTRACT.md` is the authority for the SemVer policy and the
migration notes.

## Authoritative workflow validation

The hosted gate complements the portable YAML subset check with
[actionlint 1.7.12](https://github.com/rhysd/actionlint/releases/tag/v1.7.12),
whose release tag resolves to upstream commit
[`914e7df21a07ef503a81201c76d2b11c789d3fca`](https://github.com/rhysd/actionlint/commit/914e7df21a07ef503a81201c76d2b11c789d3fca).
The workflow downloads only
`actionlint_1.7.12_linux_amd64.tar.gz` and verifies the upstream-published
SHA-256
`8aca8db96f1b94770f1b0d72b6dddcb1ebb8123cb3712530b08cc387b349a3d8`
before extraction or execution. The version and digest are explicit workflow
constants; a failed download, digest, version, current-workflow check, fixture,
report, or artifact fails the terminal verdict.

`test_workflow_validation.py` requires that exact actionlint version. It
accepts the tracked workflows and a valid local composite action, and requires
rejection of malformed YAML, duplicate job IDs, an invalid job key, missing
local-action metadata, and a missing local Node entry point. Run it locally with
an independently verified binary:

```bash
python3 tools/test_workflow_validation.py \
  --root . \
  --actionlint /verified/path/to/actionlint \
  --summary test-results/workflow-validation.md
```

The release archive and checksum catalogue are authoritative upstream assets;
review both the [release notes](https://github.com/rhysd/actionlint/releases/tag/v1.7.12)
and [published checksums](https://github.com/rhysd/actionlint/releases/download/v1.7.12/actionlint_1.7.12_checksums.txt)
before changing either pin.

## Release-integrity gate

`check_release.py` validates one initialized generated-project candidate or the
canonical template itself against the versioned `.github/release-policy.json`.
It requires version, dated changelog, tag, Git source, external evidence,
profile assurance, and any staged assets to name the same full candidate SHA.
Reports are deterministic and written atomically.

Exercise all three generated profiles, the template release profile, and every
named failure path with:

```bash
python3 tools/check_release.py --root . --self-test \
  --summary test-results/release-self-test.md
```

Candidate validation requires `--tag`, `--candidate-sha`, and `--evidence`.
Binary UI/application distributions also require `--asset-manifest`; a
source-only release omits it. After creating the local annotated tag, add
`--require-tag-ref` to prove its object type and target. See
[`docs/RELEASE_EVIDENCE.md`](../docs/RELEASE_EVIDENCE.md) for the exact JSON,
manifest, profile, and command contracts.

The hosted repository gate runs the complete self-test and fails if the tool or
its report is missing. The tool validates recorded evidence but does not execute
Excel or claim that a manually packaged Office file was built from source.

## Canonical repository initializer

`initialize_repository.py` converts a clean generated repository from template
mode to one explicit profile. It validates the complete substitution set,
defaults to a deterministic dry-run, applies only with `--apply`, removes
non-applicable/template-only content, resets inherited changelog history, and
resets the generated version to the `0.0.0` sentinel, and supports an idempotent
second run.

See [`docs/INITIALIZATION.md`](../docs/INITIALIZATION.md) for the token catalogue,
profile commands, optional and repeatable values, and the transparent manual
fallback. In this generated repository the self-test validates the recorded
application profile, identity, cleanup, quality checks and repeat-run safety.
Only canonical template mode generates fixtures for all three profiles. Run:

```bash
python3 tools/initialize_repository.py --root . --self-test
```

## Rules

- Tools must fail clearly and return a non-zero status for a blocking result.
- Pin or document material runtime dependencies.
- Separate generated output from the script and keep transient output ignored.
- Record the exact command used by CI and release certification.
- Never embed credentials, personal paths, private data, or workstation-specific assumptions.
- Do not place production VBA, regression modules, examples, or GitHub workflow definitions here.

Workflow orchestration belongs under `.github/workflows/`; `tools/` contains the reusable logic those workflows call.
