# 🔎 Pull request review

Every change reaches the default branch or a release branch through a pull
request that passes **Repository integrity**, the required check. The
data-only `traffic-history` branch is the exception: the optional Traffic
history workflow pushes to it directly ([TRAFFIC.md](TRAFFIC.md)). This guide covers the evidence I
record before merging: which commit was reviewed, by what, and what was not.
The pull request template has a *Final-head review* section for it.

## ✅ Final-head review

A review covers only the commit it read. When Codex reviews pull requests in
this repository, it runs when a pull request is opened or a draft is marked
ready, not after later pushes. A review needs no Codex environment for the
repository. If Codex answers a pull request with a request to create an
environment instead of reviewing, comment `@codex review`: the review then
runs without one. Before merging a pull request I authored:

1. Push the last fixes and note the current head commit SHA.
2. If no completed review covers that SHA, for example because fixes were
   pushed after the opening review, wait until any running review has
   finished and post `@codex review` as a new top-level comment on the pull
   request. If the opening review already covers the head, go to step 4.
3. Wait until the Codex review summary shows **Completed** for that head SHA.
   A running review, a reaction, or a completed review of an older commit is
   not enough. If the summary shows a shortened SHA, open the commit and
   compare it with the head.
4. Fix or answer every finding and resolve the threads that apply. Any further
   push, a branch update included, makes the review stale: repeat from step 1.
5. Check that **Repository integrity** passes on the current head and that
   GitHub's merge requirements are met. Record the reviewed SHA, the review
   link and the check result in the pull request, then merge.

If the review fails, is unavailable or does not finish, I leave the pull
request open and record what blocks it.

<a id="excel-run"></a>

## 🧪 Excel run

CI cannot run Excel. A pull request that changes VBA source, a workbook layout
or a test case that runs in Excel keeps the *Excel run* block of the pull
request template, and records a run on the exact commit of the pull request:

- the commit, the Excel build and bitness, and the Windows version;
- whether **Debug → Compile VBAProject** passed, or `ON_DEMAND` as the
  one-command run records it;
- the suite and its mode, and where the log is;
- the completion state, with the `outcome` of the
  [host record](../methodology/TEST_HARNESS.md#host-record) when the run
  stopped early:

| State | Meaning |
| --- | --- |
| `PASS` | Every case and assertion ran and passed |
| `FAIL` | At least one case or assertion failed |
| `INCOMPLETE` | Some case did not run, for example because it was skipped or the run stopped. It is not a pass, even if every case that ran passed ([completeness policy](../methodology/TEST_CASES.md#completeness-policy)) |
| `NOT RUN` | No Excel run was made; the *Not validated* section says why |

A run recorded as [recording an Excel run](../../INSTALLATION.md#recording-an-excel-run)
describes, and checked with `tools/check_excel_run.py` against the commit, is
the strongest evidence; cite its report.

A push to VBA source after the run makes it stale, like a review. Record a new
run, or state `NOT RUN` for the new head. A run on an earlier commit of the
pull request still counts for the head when
`tools/check_excel_run.py --equivalent-to` proves their
[run inputs](EXCEL_RUN.md#a-run-for-another-commit) identical;
record it as `run at <SHA>, equivalent`, with the report's `Equivalent:`
line. When the later commits change only comments or layout in VBA
components, `--carried-to` carries the run to the head instead
([carried evidence](EXCEL_RUN.md#carried-evidence)); record it as
`run at <SHA>, carried`, with the report's `Carried from` line. Carried
evidence counts for a pull request, never for a release. A pull request that
changes none of these files deletes the block.

## 🌿 Release branches

While a release branch such as `release/1.2.0` is active, the work for that
release reaches the default branch through it. Without one, task pull requests
target the default branch directly and the integration and sync steps do not
apply.

| Pull request | Base ← head | Issue links | Merge method |
| --- | --- | --- | --- |
| Task | release branch ← task branch created from it | `Refs #N` | merge commit or squash |
| Integration | default branch ← release branch | `Refs #` for the issues and pull requests it carries | merge commit |
| Sync | release branch ← default branch | `Refs #` for the integration pull request | merge commit |

- GitHub closes an issue when a closing keyword that names it reaches the
  default branch: in the title or body of a pull request that merges there, or
  in the message of any commit that the merge brings there. The keywords are
  `close`, `fix` and `resolve`, with their `-s` and `-d` forms, in any letter
  case. The integration brings every commit of the release branch, so a
  keyword in any of them closes its issue then, even one whose acceptance
  criteria still wait for evidence. While a release branch is active, use only
  `Refs #N` in commit messages, squash messages, and the titles and bodies of
  task pull requests. `Resolves #N` stays for a pull request into the default
  branch while no release branch is active, for an issue that it completes.
- After a task pull request merges into the release branch, close its issue by
  hand with the *Resolution* that
  [ISSUE_LABELS.md](ISSUE_LABELS.md#-closing-issues) requires, and move it to
  the milestone of that release branch's version.
- An integration pull request uses `Refs #`, never a closing keyword: a keyword
  there would close every named issue, including one whose acceptance criteria
  still wait for evidence.
- The **Closing keywords** workflow runs `tools/check_closing_keywords.py` on
  each pull request. Each finding names the commit, the keyword as written and
  the issue. Into a release branch, any keyword in a
  commit between the base and the head, in the title or in the body fails; a
  sync pull request has only its title and body read, because its commits are
  already on the default branch. Into the default branch from a release branch,
  a keyword that names an open issue fails, and one that names a closed issue
  or a pull request is a note. Any other pull request is reported as not
  checked. It is not a required check: check that it passes before merging.
- A merged commit cannot be changed. When an integration carries a commit
  keyword that names an open issue, list the issue in the integration pull
  request's body on a line of its own, `Reopen after merge: #N`, or
  `Reopen after merge: #N, #M` for several. The check then reports it as a
  note; reopen each listed issue right after the merge. A keyword in the title
  or body is fixed by editing them instead.
- To check a task branch before opening its pull request, run
  `python tools/check_closing_keywords.py --repository OWNER/REPOSITORY --mode release --base origin/release/1.2.0 --head HEAD`.
  With `--mode integration`, the check reads issue states through the API with
  the token in `GITHUB_TOKEN`, or from a file passed with `--input`.
- Open the sync pull request right after each integration. It brings the
  integration's merge commit back, changes no file, and lets the next task
  branch start from the same history as the default branch.
- Integration and sync pull requests follow the final-head review above, like
  any other pull request, with one exception. A sync pull request needs no new
  automated review when its head is the integration merge commit, the diff
  between the release branch and that head is empty, and the integration pull
  request's final-head review covered the same tree. Its *Final-head review*
  then names the integration pull request and its review in place of a new
  one. **Repository integrity** must still pass on the sync head. A sync pull
  request that changes any file, for example to resolve a conflict, is
  reviewed as usual.
- A sync pull request under that exemption may have auto-merge enabled, with
  the merge-commit method, as soon as it is opened: GitHub merges it once
  every required check and rule passes. Merge no task pull request into the
  release branch until it has merged: the release-branch ruleset requires a
  head that contains its base, and the sync head is the default branch, which
  cannot be updated from the release branch. Other work can go on. Its body
  may replace the template with this short record:

  ```text
  Sync after #N. Head: <integration merge commit SHA>, the integration merge commit.
  Diff against <release branch>: empty; tree <tree SHA>.
  Final-head review: the review of #N at <SHA> (<link>) covered the same tree.
  The required Repository integrity check gates the merge (auto-merge).
  Refs #N
  ```

  If the check fails, auto-merge waits: fix the cause on the default branch,
  not in the sync pull request. If a task pull request merged first, close
  the sync pull request. Open a new one from a branch created from the release
  branch, with the default branch merged into it. It changes no file, but its
  head is not the integration merge commit, so it is reviewed as usual.
- `.github/dependabot.yml` names the active release branch in `target-branch`.
  Update it when a new release branch opens, or remove it when none is active.
- GitHub reads some files only from the default branch: the issue forms, this
  pull request template, `.github/dependabot.yml`, and the workflows started by
  issue, milestone, schedule and `pull_request_target` events. A change to one of them that is
  merged only into the release branch takes effect after the next integration;
  integrate early when it must apply sooner.

## 🤖 Dependency updates

Dependabot proposes the version updates that `.github/dependabot.yml`
configures. Each waits 14 days after its release (`cooldown`), so a package
withdrawn soon after publication is never proposed.

- A **security update** does not wait, and it opens against the default
  branch whatever `target-branch` says. While a release branch is active, make
  the same change in a task pull request into the release branch, and close
  the bot pull request with a link to it.
- **actionlint** is outside Dependabot: its version and the SHA-256 of its
  archive are set in `.github/workflows/repository-integrity.yml`. Update both
  together in a task pull request.
- Before merging an update, read the release notes between the two versions;
  for an Action, check that the new SHA belongs to the tag in its comment, and
  that the pin and its version comment change together. Record the review as
  *Bot pull requests* below describes.

## ⚖️ Recorded exceptions

Record these in the *Final-head review* section, in place of the steps above,
or in a comment on the pull request when it has no such section:

- **Identical reviewed content.** A pull request whose head has the same tree
  as a commit already reviewed against the same base may cite that review.
  Record both SHAs and the shared tree, and do not claim a review of the head
  itself.
- **Merge authorized without a final-head review.** I may decide to merge
  without one. Record my decision word for word, what changed after the last
  reviewed commit, and that no review of the final head is claimed.
- **Bot pull requests.** Codex does not review pull requests opened by
  Dependabot or another bot, and their description does not use the pull
  request template. Before merging, record my own review of the final diff and
  head SHA in a comment on the pull request, with **Repository integrity**
  passing, and do not claim a Codex review.

Self-review, an automated review and CI are separate kinds of evidence. This
procedure is a merge rule I apply, not an additional required status check.
