# 🛠️ Contributing

Thank you for helping. This guide is the entry point; the detailed rules live
in the guides it links, and this file does not repeat them.

By taking part you agree to the [code of conduct](CODE_OF_CONDUCT.md). Report a
suspected vulnerability privately, as [SECURITY.md](SECURITY.md) describes, never
in an issue or pull request.

## 🐞 Issues

Open issues with the issue forms. **Issue triage** checks each open issue
against the [triage rules](docs/repo_governance/ISSUE_LABELS.md#triage-rules).
The [issue guide](docs/repo_governance/ISSUE_LABELS.md) also explains the
labels, the priorities and how an issue is closed, with a *Resolution*
section.

## 🌿 Branches and pull requests

- While a release branch such as `release/1.2.0` is active, start task branches
  from it and open the pull request against it; otherwise use the default
  branch. The [release branches](docs/repo_governance/PULL_REQUESTS.md#-release-branches)
  section gives the bases, issue keywords and merge methods.
- Fill in the pull request template. Its *Calculation impact* section names
  the source and the class of the expected values behind a changed
  calculation, or says there is none. Its *Validated* section lists the checks
  you ran, and for a defect fix its [negative control](#-negative-controls);
  write NOT RUN or NOT APPLICABLE with a reason instead of claiming
  evidence that was not produced.
- Add the `CHANGELOG.md` entry under `[Unreleased]` in the same pull request,
  as its [maintenance policy](CHANGELOG.md#-maintenance-policy) describes; a
  fix for a failed release gate before the tag goes in the dated section
  being released instead.
  `tools/check_changelog.py` checks the format, not that the change is
  recorded.
- Every pull request needs a passing **Repository integrity** check and a
  review of its final head, as the
  [pull request guide](docs/repo_governance/PULL_REQUESTS.md) describes.

## ✍️ Commit messages

- The subject says what changed, after a type prefix such as `feat:`, `fix:`,
  `docs:`, `ci:` or `chore:`.
- The body says why, and what was validated: the checks you ran, and what was
  not run, such as an Excel run.
- The last line references the issue with `Refs #N`.
- A closing keyword, such as `Resolves #N`, `Fixes #N` or `Closes #N`, belongs
  only in a pull request into the default branch, or a commit that lands there,
  while no release branch is active, and only for an issue that it completes.
  GitHub closes an issue when a commit whose message names it with a keyword
  reaches the default branch. A commit merged into a release branch reaches it
  at the integration, whatever the state of the issue by then, and it cannot be
  changed after the merge. The
  [release branches](docs/repo_governance/PULL_REQUESTS.md#-release-branches)
  section explains how the **Closing keywords** workflow checks this.

A squashed pull request becomes one commit, and its message follows the same
rules: its title and body become the message, so they use `Refs #N` too.

## 📏 Rules of evidence

- Expected values come from the documented method and its sources, never from
  running the code under test, and a value derived within the project stays
  `illustrative` ([reference classes](docs/methodology/TEST_CASES.md#reference-classes),
  [derivation scripts](docs/methodology/TEST_CASES.md#derivation-scripts)).
- A skipped test is not a pass. A run that does not execute every expected
  case and assertion is incomplete, even if everything that ran passed
  ([completeness policy](docs/methodology/TEST_CASES.md#completeness-policy)).
- A static check is not an Excel run. Record only evidence that was produced,
  on the commit it was produced for.
- A test that passes before and after a fix proves nothing about the defect.
  A fix shows its test failing on the defective code first
  ([negative controls](#-negative-controls)).

## 🔬 Negative controls

A pull request that fixes a defect shows each new or changed test that proves
the fix failing on the defective code, then passing on the fix. Make the
defective code in one of two ways:

- **The base.** Check the base out in a detached worktree, for example with
  `git worktree add --detach ../base-check origin/main`, copy only the new or
  changed test files into it, and run them there. Remove it afterwards with
  `git worktree remove ../base-check`.
- **A mutant.** When the base cannot run the test, for example because the
  test needs code that the fix adds, put the defect back into a copy of the
  fixed file. `tools/make_mutant.py` writes it under `test-results/mutants/`
  and prints its SHA-256; its docstring says what it refuses:

  ```sh
  python tools/make_mutant.py src/core/CORE_Example.bas --find "a - b" --replace "a + b"
  ```

  Run a mutated VBA component in a development workbook in place of its
  source component, built with `--override`
  ([building a workbook from source](INSTALLATION.md#building-a-workbook-from-source))
  or [imported by hand](INSTALLATION.md#%EF%B8%8F-importing-vba-into-excel),
  and build the workbook from source again for the run on the fix. For another
  file, copy the mutant over its source in a detached worktree of your branch
  and run the test there.

Record in *Validated* which way you used, the command, and the failing result
on the defective code; for a mutant, also its SHA-256. When no test can fail on
the defect, write NOT APPLICABLE with the reason.

A negative control proves one fix. `tools/verification_depth.json` keeps the
standing record: for each check of **Repository integrity** that can fail the
run, the tests that prove it can fail. The `verification-depth` check fails
while such a check has no entry; the docstring of
`tools/check_verification_depth.py` gives the fields and what it rejects. A
pull request that adds a check, or renames or deletes a proof, updates the
inventory in the same change.

## 📐 Contract changes

A contract change alters what a documented contract promises its users:

- a dataset, field, layout, unit or validation rule of the
  [data contract](docs/methodology/DATA_CONTRACT.md);
- the inputs, rules, sources or results of a
  [calculation contract](docs/methodology/CALCULATION_CONTRACTS.md);
- a sheet, code name, defined name or written area of the
  [workbook contract](docs/methodology/WORKBOOK_CONTRACT.md);
- a declaration in the [supported VBA API](docs/PUBLIC_API.txt);
- a frozen [accuracy contract](docs/methodology/TEST_CASES.md#accuracy-contracts).

The pull request that makes it updates the contract, its
tests and the documents that describe it together, and its changelog entry
says what changed and what a user must do. A change to computed results or to
an input or output layout carries the `behavior-change` label; from `1.0.0`,
an incompatible one needs a major version
([version rules](CHANGELOG.md#date-and-version-rules)). A change that alters
none of these is not a contract change.

## 🧪 Local checks

Run every check of **Repository integrity** before opening a pull request:

```sh
python -m pip install --require-hashes -r .github/requirements.txt
python -m pip install --only-binary=:all: --require-hashes -r tools/requirements-quality-ci.txt
python tools/check_repository.py
```

- **Tools.** The first command installs the pinned configuration parser, which
  the tests that read YAML need. The second installs Ruff and mypy, whose
  hashes cover CPython 3.13 on Linux x86_64. The tests also need Node.js, for
  the label scripts.
- **actionlint.** Download the release that `ACTIONLINT_VERSION` names in
  `.github/workflows/repository-integrity.yml` and put it on your `PATH`. On
  Linux x86_64, first check the archive against `ACTIONLINT_LINUX_AMD64_SHA256`
  there, as CI does.
- **Results.** A check whose tool is missing is reported as skipped. The run
  writes a log for each check, `report.json` and `summary.md` under
  `test-results/integrity/`, which Git ignores; the docstring of
  `tools/check_repository.py` says how it runs.
- **Exit status.** Every tool in `tools/` exits 0 when it passed, 1 when it
  ran and found problems, and 2 when it could not run: an invalid argument, an
  unreadable input, a refusal or a missing prerequisite. A red run that exits
  2 needs a different command, not a code fix.

It runs these checks, in this order. `OWNER/REPOSITORY` is the repository that
`--repository` names, as CI passes it, or else the GitHub repository of the
`origin` remote.

<!-- BEGIN generated: integrity-checks -->
| Check | Runs |
| --- | --- |
| `tests` | `python -m unittest discover -s tests -v` |
| `ruff` | `python -m ruff check --no-cache tools tests` |
| `mypy` | `python -m mypy --no-incremental` |
| `actionlint` | `actionlint -no-color`, once the probe `actionlint -no-color .github/actionlint/invalid-workflow.yml` has failed with `unexpected key "bogus-key"` |
| `labels-sync-self-test` | `node .github/scripts/labels-sync.mjs --self-test` |
| `labels-catalogue` | `node .github/scripts/labels-sync.mjs --mode validate` |
| `labels-drift-self-test` | `node .github/scripts/labels-drift.mjs --self-test` |
| `issue-configuration` | `python tools/check_issues.py --check-config` |
| `template-only-content` | Built in: Only the template keeps the template-only sections of AGENTS.md and README.md. |
| `issue-contact-links` | Built in: The issue forms' contact links point to this repository. |
| `readme-traffic-badges` | Built in: The README's traffic badges read this repository's traffic. |
| `binaries` | `python tools/check_binaries.py` |
| `secret-files` | Built in: No tracked file is named as a secret or private key file. |
| `vba-format` | `python tools/format_vba.py --check` |
| `vba-source` | `python tools/check_vba_source.py` |
| `vba-layout` | `python tools/check_vba_layout.py` |
| `vba-conditionals` | `python tools/check_vba_conditionals.py` |
| `vba-jumps` | `python tools/check_vba_jumps.py` |
| `vba-public-api` | `python tools/check_vba_public_api.py` |
| `vba-dependencies` | `python tools/check_vba_dependencies.py` |
| `vba-identifiers` | `python tools/check_vba_identifiers.py` |
| `vba-errors` | `python tools/check_vba_errors.py` |
| `vba-declares` | `python tools/check_vba_declares.py` |
| `vba-calls` | `python tools/check_vba_calls.py` (advisory: findings warn, never fail the run) |
| `data-contract` | `python tools/check_data_contract.py` |
| `test-cases` | `python tools/check_test_cases.py` |
| `derivations` | `python tools/check_derivations.py` |
| `independent-implementations` | `python tools/check_independent.py` |
| `accuracy` | `python tools/check_accuracy.py` |
| `workbook-protection` | `python tools/check_workbook_protection.py` |
| `changelog` | `python tools/check_changelog.py --repository OWNER/REPOSITORY` |
| `generated-docs` | `python tools/render_docs.py --check` |
| `restated-counts` | `python tools/check_restated_counts.py` |
| `template-wording` | `python tools/check_template_wording.py` |
| `markdown-links` | `python tools/check_links.py` |
| `documented-commands` | `python tools/check_doc_commands.py` |
| `verification-depth` | `python tools/check_verification_depth.py` |
| `dependabot-target-branches` | Built in: Every Dependabot target-branch exists. |
| `codeql-default-branch` | Built in: CodeQL, Minimum Python and Repository integrity run on the default branch, and every workflow's branch filters are well formed. |
| `required-documents` | Built in: The documents the repository requires exist and are not empty. |
| `whitespace-committed` | `git diff --check 4b825dc642cb6eb9a060e54bf8d69288fbee4904 HEAD` |
| `whitespace-local` | `git diff --check HEAD` |
<!-- END generated: integrity-checks -->

A table between `BEGIN generated` and `END generated` markers is rendered from
the code: edit its renderer in `tools/render_docs.py`, never the block, and run
`python tools/render_docs.py --write`. Outside a block, link a count's source
instead of copying the count (`restated-counts`), and word each sentence for
the repository it is in (`template-wording`); each tool's docstring lists what
it rejects.

To run one check, use its own command, such as
`python tools/format_vba.py --check`, or
`python -m unittest tests.test_check_links` for one test module.
`python tools/format_vba.py --write` fixes what `vba-format` reports;
re-import a rewritten module into the VBE before committing
([house style](docs/VBA_HOUSE_STYLE.md#source-representation)). To run Ruff
and strict mypy on their own, see the
[setup guide](docs/setup/REPOSITORY_SETUP.md#-maintain-the-baseline).

## 📄 Data and licensing

- Use synthetic data only: no real portfolio, client, counterparty or
  credential data in code, tests, issues, screenshots or workbooks.
- Contributions are licensed under the repository's [LICENSE](LICENSE). Files
  listed in the [third-party notice](docs/licenses/THIRD_PARTY_NOTICES.txt) keep
  the license the notice states for them.
