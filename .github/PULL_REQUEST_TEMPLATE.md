<!--
  Base: the default branch, or the active release branch when development
  happens on one. State which applies, and the merge method, on the Base line.
  Write NOT RUN or NOT APPLICABLE with a reason; never claim evidence that
  was not produced. Synthetic data only; no secrets or vulnerability details.
-->

Base: <!-- the default branch, merged with a merge commit; or `release/X.Y.Z`: a task PR squashed or merged with a merge commit, a sync PR merged with a merge commit -->

## 🔗 Linked issues

<!-- While a release branch is active, use only Refs # here, in the title and
in every commit message: a closing keyword in any of them closes its issue
when the release branch reaches the default branch. Use Resolves # only into
the default branch while no release branch is active, for an issue this PR
completes. Integration and sync PRs use Refs # for the issues and preceding
PRs they carry; an integration lists each open issue that a merged commit
would close on a line "Reopen after merge: #N". Release branches section of
docs/repo_governance/PULL_REQUESTS.md: bases, keywords and merge methods.
Milestones belong to issues only. -->

Refs #

## 🧾 Changes

-

## 💡 Why

## 🧮 Calculation impact

<!-- For a change to a calculation's results, rules or expected values: the
source each changed rule cites (its `SRC-` ID and locator) and the class of
the expected values that test it: `published`, `independent` or
`illustrative`, as the reference classes of docs/methodology/TEST_CASES.md
define them. Otherwise replace both lines with "No calculation change". -->

- Source:
- Class of the expected values:

## ✅ Validated

<!-- Commands and checks actually run, with results. Hosted checks run on this pull request.
Negative controls: CONTRIBUTING.md. -->

- [ ] `python tools/check_repository.py`, the local checks of CONTRIBUTING.md; name any check it skipped
- [ ] Ruff, strict mypy and actionlint pass in `Repository integrity`
- [ ] Negative control, for a defect fix: the new or changed test fails on the defective code (a detached worktree of the base, or a mutant from `tools/make_mutant.py`) and passes on the fix. How, and the failing result:
- [ ] Other checks specific to this change:

### 🧪 Excel run

<!-- Keep this block when the pull request changes VBA source, a workbook
layout or a test case that runs in Excel; delete it otherwise. Only PASS with
every case run counts as a pass. Rules: docs/repo_governance/PULL_REQUESTS.md. -->

| Item | Result |
| --- | --- |
| Exact commit | <!-- Full SHA of the source imported; or "run at <SHA>, equivalent" with the Equivalent: line of check_excel_run.py --equivalent-to; or "run at <SHA>, carried" with its Carried from line of --carried-to, when only VBA comments or layout changed --> |
| Excel build and bitness, Windows version | |
| `Debug → Compile VBAProject` | <!-- PASS / FAIL / ON_DEMAND (the one-command run) / NOT RUN, with a reason --> |
| Suite and mode | |
| Completion state | <!-- PASS / FAIL / INCOMPLETE / NOT RUN; for a run that stopped early, the host record's outcome --> |
| Evidence | <!-- Where the log or evidence folder is --> |

## 🔎 Final-head review

<!-- After the last push, compare the reviewed commit with the current head.
Procedure and exceptions: docs/repo_governance/PULL_REQUESTS.md. Bot pull
requests do not use this template; record my review of their final head in a
comment on the pull request. -->

- Final PR head SHA:
- Review type and reviewed commit SHA:
- Review result and link:
- `Repository integrity` result and run link:

- [ ] Review covers the current PR head; all findings are fixed or answered and applicable threads resolved.
- [ ] `Repository integrity` passes for the current revision and GitHub merge requirements are satisfied.
- [ ] Any push after this record is followed by renewed review and CI verification.

## ⚠️ Not validated

<!-- State plainly what was not checked and why. Write "Nothing" only if that is true. -->

## ☑️ Checklist

- [ ] The base is the default branch or the active release branch, as stated above.
- [ ] Each linked issue meets the *Triage rules* of `docs/repo_governance/ISSUE_LABELS.md`; this PR has no milestone, links the relevant issues and carries appropriate labels.
- [ ] Scope is focused: no unrelated changes, local artifacts or generated output.
- [ ] Documentation affected by the change is updated.
- [ ] `CHANGELOG.md` has an entry under `[Unreleased]` for this change (in the dated section being released, for a fix of a failed release gate before the tag), or the change is one its maintenance policy says needs none.
- [ ] No tokens, credentials, real portfolio or client data.
