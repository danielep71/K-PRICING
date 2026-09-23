# 🧩 Checker Development Contract

[![Runtime: single file](https://img.shields.io/badge/runtime-single--file-217346)](../tools/check_repo.py)
[![Dependencies: stdlib only](https://img.shields.io/badge/dependencies-stdlib%20only-success)](../tools/checker_development.py)
[![Coverage: semantic policy](https://img.shields.io/badge/coverage-semantic%20policy-6F42C1)](../tools/check_policy_coverage.py)

`tools/check_repo.py` remains the canonical portable checker delivered to generated repositories. **The single-file identity rule applies to that canonical checker only.** Its reviewed source is the distributable artifact, so there is no bundle transform that can drift from source. Focused sibling gates are development/runtime tools within the repository and may share private standard-library infrastructure.

## 🔧 Shared focused-gate primitives

`tools/_gatelib.py` owns the small cross-tool mechanics that are genuinely identical: Git subprocess wrappers, tracked-file enumeration, deterministic UTF-8/LF report writes, the common `--root` / `--output` / `--summary` / `--self-test` parser, and `run_gate`, the typed runner that owns the shared console, evidence and exit-code contract. Focused gates import those primitives instead of maintaining copies. Tool-specific `run_check`, `build_report`, `run_self_test`, semantic rules, fixtures and Markdown renderers remain local because their behavior and evidence schemas differ.

### `run_gate` ownership

`run_gate` centralizes only the orchestration that was provably identical across gates: self-test dispatch, report construction, canonical JSON serialization, Markdown summary writing, console output, and the `0` / `1` / `2` exit mapping. It never widens a gate's exception handling. Each caller passes its own operational-exception tuple, so a programming error still surfaces as a traceback instead of being reported as exit code `2`.

Gates that historically evaluated `--self-test` outside their operational handler reported failures as `SELF-TEST ERROR`; gates that evaluated it inside reported `ERROR`. `run_gate` preserves both wordings through `self_test_error_prefix`.

The complete registry is `GATE_RUNNER_CONSUMERS` and `GATE_RUNNER_EXCLUSIONS` in
[`checker_development.py`](../tools/checker_development.py). Its report lists
every consumer and each exclusion's reason. Keep that executable registry as
the maintained list rather than duplicating a table that misses later gates.

`ownership_scan` checks focused tools other than `_gatelib.py` and the canonical
checker. Every top-level `main` in that scope must be a declared consumer or
exclusion; adding a gate without updating the declaration fails the contract,
as does an excluded tool quietly adopting the runner. The canonical checker is
excluded separately by its self-contained import contract. Sixteen independent
runner tests cover CLI flags, defaults/help, self-test dispatch, diagnostic
prefixes, exit mapping, deterministic evidence, write failures and propagation
of non-operational exceptions.

The canonical-template merge-history control introduced after v1.2.0 does **not**
add another focused-gate entry point. Its ownership sits inside the already
registered `check_release_semantics.py` consumer: the normal SemVer/changelog
report gains a template-only history branch, while generated repositories report
that branch as not applicable. The deterministic history fixtures therefore run
through `check_release_semantics.py --self-test`, preserving the existing
shared-runner registry and avoiding a parallel orchestration path.

`tools/check_repo.py` must never import `_gatelib.py`. Generated repositories retain `_gatelib.py` for the focused operational gates, while the canonical checker remains independently copyable and executable as one standard-library-only file. `checker_development.py` enforces this ownership boundary in the canonical template. The checker-development workflow, this document, and the `policy_coverage_*` semantic-coverage harness are template-maintainer assets and are removed by initialization rather than shipped into generated projects.

The public API gate also uses `check_vba_conditionals.py` to evaluate the same three supported compilation environments. It checks name collisions only where declarations can coexist, and requires one manifest declaration row plus a `# SIG` record for every distinct reachable signature. Unknown conditions fail closed. These are static models, not evidence of Excel runtime certification.

### Self-test interface coverage

Runner ownership and `--self-test` support are separate contracts. The development
report discovers CLI definitions, imported-main wrappers and executable
`__main__` guards (including reversed equality), then checks each
script's `--help`, and requires either an advertised `--self-test` flag or a
non-empty reason in `SELF_TEST_EXCLUSIONS`. Missing declarations, failed help,
stale exclusions, and exclusions for tools that now advertise the flag fail.
The report lists every inspected CLI and its alternative test command or limitation.
Guard discovery preserves `not`, `and` and `or` polarity and literal truth values.
Chained comparisons are modeled as conjunctions of adjacent comparisons; literal
equality and inequality are resolved without executing source.
A guarded branch must be possible in script mode and impossible on import; this
also recognizes a main-only `else` branch. Unknown operands remain unknown, so
this is a bounded syntactic model, not general Python control-flow analysis.

`check_documentation.py` and `check_wiki.py` intentionally use dedicated offline
suites: `python tools/test_documentation.py -v` and `python tools/test_wiki.py -v`.
CI runs them in `static-checks.yml` and `wiki-checks.yml`, respectively. They remain
runner consumers; their self-test exclusions do not exempt them from ownership checks.
The registry also names the separate suites for the other operational CLIs without
the flag and explicitly discloses the provisioning fixture utility's lack of a
dedicated offline suite.

This check verifies interface declarations, not exhaustive test coverage. Advertising
the flag does not prove that its fixtures ran or that every branch is covered.
Executable unittest scripts are included with explicit exclusions identifying their
normal fixture invocation. Non-CLI helper modules remain outside this interface
inventory; suite execution and semantic coverage remain separate. No host execution or release
certification is inferred from synthetic fixtures.

### Supported invocation mode

The supported focused-tool interface is **path execution from the repository checkout**, for example `python3 tools/check_vba_public_api.py --root . --self-test`. Python places the executed script's directory on `sys.path`, which is the declared mechanism by which focused sibling gates resolve the private `_gatelib.py` module.

`python -m tools.<module>` is **not** part of the supported contract: `tools/` is not a public Python package and no package-installation interface is promised. If module-mode execution is added later, it must be introduced deliberately with package-aware imports and fixtures for both invocation modes rather than relying on incidental interpreter path behavior. The canonical `check_repo.py` remains unaffected because it has no sibling import.

### Python import and typing ratchet

Ruff `I001` is part of the maintained Python lint baseline. Import blocks must
therefore stay deterministically ordered alongside the existing E4/E7/E9/F,
C90 and S314 checks. This is an ordering rule, not formatter adoption: CI still
does not run `ruff format`, and the 100-column target remains advisory because
E501 is deliberately not selected.

Mypy remains incremental. The global configuration continues to cover the whole
`tools/` tree at the established baseline. `_gatelib` is the first promoted
strict module because it is imported by the focused-gate stack. In
`pyproject.toml`, its per-module override spells out the exact strictness flags
enabled by pinned mypy 2.3.1 rather than using the `strict = true` meta-option;
this keeps strictness scoped to `_gatelib` and prevents unrelated modules from
being promoted accidentally.

The reproducible strict check for the boundary is:

```bash
mypy --strict tools/_gatelib.py
```

The normal hosted `mypy` invocation enforces the same pinned strict bundle on
`_gatelib` through that per-module override while retaining the established
whole-tree baseline elsewhere. Before the v1.2.1 release candidate is frozen,
the exact strict command is also exercised as hosted evidence. Once a module is
promoted, do not weaken its strict settings merely to make a later change green.
Migrate additional modules only after they pass the pinned strict contract;
record any narrow temporary relaxation explicitly rather than adding blanket
ignores.

Private-member debt is deliberately a separate architecture concern. The
post-v1.2.0 inventory contains three recurring families: checker-development
introspection of private `check_repo` parser/rule helpers, semantic policy-coverage
harness access to canonical check internals, and focused-tool reuse of narrowly
scoped private parsing helpers such as Markdown destination extraction. This
patch does not enable Ruff `SLF` rules or redesign those boundaries. Only a
private access that prevents the strict `_gatelib` boundary from passing belongs
here; broader ownership/API cleanup remains later architecture work.

## 🧭 Internal boundaries

`tools/checker_development.py` parses the checker with Python AST and requires the following ordered ownership boundaries:

| Section | Sentinel | Responsibility |
| --- | --- | --- |
| Runtime core | `Repository` | repository I/O, findings and common primitives |
| Configuration | `_same_keys` | repository-profile schema and effective requirements |
| Repository policy | `check_required_paths` | generic repository, document, workflow and metadata rules |
| VBA policy | `_vba_paths` | VBE exports, VBA structure, roles, generated contracts and public surface |
| Reporting | `build_report` | deterministic report model, Markdown and console serialization |
| Fixtures | `_write_fixture` | positive/degraded synthetic repository fixtures and self-test |
| CLI | `parse_arguments` | supported command-line surface and exit behavior |

The development check fails if a boundary disappears, changes order, leaves a top-level definition outside the ordered sections, or changes the canonical policy-check sequence without an explicit update to the contract.

## 🧪 Independent tests

Maintainers can exercise parser and reporter behavior without running the full synthetic repository matrix:

```bash
python3 tools/checker_development.py --root . --self-test
```

The contract directly tests representative YAML, GitHub-style Markdown anchors, EditorConfig parsing, VBA lexical stripping, Markdown/console serialization, CLI flags and operational exit-code mapping. The full `check_repo.py --self-test` and semantic policy-coverage matrix remain separate higher-level gates. Hosted CI keeps both template-maintainer contracts together in `checker-development.yml`; the operational `static-checks.yml` deliberately remains free of template-only policy-coverage tooling so the generated workflow is self-contained.

Release-history fixtures are owned by the existing release-semantics suite rather
than this AST/interface harness. `python3 tools/check_release_semantics.py --root . --self-test`
proves compliant squash history, reviewed merge exceptions, unapproved merge
rejection and duplicate-subject rejection using temporary Git repositories.

## 🖥️ Local pre-push validation

Hosted CI is authoritative, but a maintainer should be able to reach the same
verdict before pushing. Two files make that reproducible.

`tools/requirements-dev.txt` pins the cross-platform local Python tooling to
the exact direct versions used by hosted CI. The hosted CPython 3.10 / Ubuntu
x64 installs are additionally locked by `tools/requirements-quality-ci.txt`,
`tools/requirements-coverage-ci.txt`, and
`tools/requirements-portfolio-ci.txt`, which carry the reviewed wheel hashes
consumed by `static-checks.yml`, `checker-development.yml`, and
`portfolio-drift.yml` respectively. `tools/dev_check.sh` verifies the local
version pins still mirror those hosted lock files and the workflow-declared
Ruff/mypy versions. `actionlint` is a Go binary rather than a Python package,
so it is documented in `tools/requirements-dev.txt` but not installable from it.

`tools/dev_check.sh` runs the locally reproducible gates in the hosted order and
prints a pass/fail/skip verdict:

```bash
python3 -m pip install -r tools/requirements-dev.txt
tools/dev_check.sh              # full local set
FAST=1 tools/dev_check.sh       # skip the initializer and policy-coverage gates
```

The script guards its own premise in two steps before running any gate. It first
checks that the local direct pins still match the hosted lock files, that the
Ruff/mypy lock entries still match `RUFF_VERSION` and `MYPY_VERSION` in
`static-checks.yml`, that both hosted PyYAML lock files agree with the local
pin, and that the `actionlint` version/digest still match
`static-checks.yml`. A lock or workflow pin that moves without the other
surfaces therefore cannot turn a clean local run into a false negative. It then
checks that every installed local tool version matches the pin. Both are
failures rather than warnings, because a local run on different versions can
disagree with CI in both directions. A mismatched PyYAML is the clearest case:
`check_portfolio_drift.py` refuses to run on any other version, and the
resulting failure names the portfolio suites rather than the real cause.

Workflow validation is the one gate that needs a non-Python tool.
`tools/requirements-dev.txt` carries the command that installs the same pinned,
digest-verified `actionlint` binary `static-checks.yml` uses; with it on `PATH`
the local run covers every tracked workflow and reports no skips.

The script is explicitly not a claim of CI equivalence. Workflow validation
without `actionlint` on `PATH`, the statement-coverage floor, Excel evidence,
live GitHub state and the hosted terminal verdicts are reported as SKIPPED
rather than counted as passes, and remain the hosted gates' responsibility.

## 🧰 Registered maintenance tasks

Some maintenance work must run in the pinned hosted environment and commit its
result — regenerating the wiki inventory reference is the current example.
`.github/workflows/maintenance.yml` is the one permanent, dispatch-only home for
that work:

```bash
gh workflow run maintenance.yml --ref <branch> -f task=regenerate-wiki-reference
```

That command works only after the file has reached the default branch. GitHub
registers a `workflow_dispatch` workflow from `main` alone, so while this file
lives only on a release branch the workflow is not dispatchable at all — the
Actions UI does not offer it and the REST dispatch endpoint answers 404. Naming
the branch with `--ref` does not work around it. Once the file is on `main`,
`--ref` may name any branch that carries it, and the task commits to that
branch.

It replaces the pattern of adding a disposable workflow per task. Each such
workflow triggered on its own push, granted itself `contents: write`, ran at the
moment it was introduced and then deleted itself, which left a permanent entry
in the Actions tab and no window in which it could be reviewed.

The permanent workflow keeps these rules, and a change to it is a reviewed edit:

- `workflow_dispatch` only; it never triggers itself and never on push.
- The task list is a closed `choice`. An arbitrary-command input would be the
  script-injection hole Scorecard's Dangerous-Workflow check flags, so adding a
  task means editing this file.
- The selected task reaches the shell through `env`, never through inline
  `${{ }}` interpolation.
- Write scope is granted on the job, not the workflow.
- The result is validated before it is committed, and a task that changes
  nothing exits cleanly instead of creating an empty commit.
- It never deletes itself.

## 📦 Portability and artifact identity

The runtime checker may import only Python standard-library modules and must not use relative imports. No `pip`, package manager, virtual environment, generated package tree or network dependency is required in a generated VBA repository.

Every development-contract report records the SHA-256 of `tools/check_repo.py`. Because the reviewed source and shipped artifact are the same bytes, reproducibility is an identity operation (`build_transform = none`) rather than a hidden build step.

## 🔁 Change procedure

1. Change `tools/check_repo.py`, `_gatelib.py`, and/or the focused gate that owns the affected behavior. Keep `check_repo.py` independent of `_gatelib.py`.
2. Run `python3 tools/checker_development.py --root . --self-test`.
3. Run `python3 tools/check_repo.py --root . --self-test`.
4. Run `python3 tools/check_policy_coverage.py --root . --self-test` so every canonical blocking finding remains exercised.
5. Run `tools/dev_check.sh` to reproduce the locally checkable part of hosted CI before pushing.
6. Run the normal repository gate and require successful hosted terminal verdicts from both `Checker development` and `Repository integrity`.
7. If a deliberate internal boundary, CLI contract or canonical check order changes, update this document and `checker_development.py` in the same reviewed change.

## 🚫 Non-goals

This contract does not replace authoritative GitHub Actions validation, Excel/VBA compilation, runtime regression tests, numerical assurance, UI-state tests, release evidence or profile-specific specialist controls.
