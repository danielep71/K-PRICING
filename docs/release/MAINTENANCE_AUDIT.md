# 🧹 Maintenance audit

A review of the repository's issues, pull requests, review discussions and
notifications, done before every release. It checks that the record a
release freezes is true: issues closed only when their criteria are met,
open issues current, review findings answered, and no claim about a change,
a test or a review that its evidence does not support.

Run it during preparation, before
[RELEASING.md step 2](../../RELEASING.md#2-integrate-and-freeze-the-candidate)
freezes the candidate. Post its report as a comment on the
[milestone closeout](CLOSEOUT.md) issue. A finding that needs a code or
documentation change is fixed through a task pull request before the
candidate is frozen, or tracked in an issue and accepted as a deviation in the
closeout record.

## 🧭 Rules

- **Read first.** Read the repository's governance and contribution rules
  before changing anything: [CONTRIBUTING.md](../../CONTRIBUTING.md), the
  [issue rules](../repo_governance/ISSUE_LABELS.md) and the
  [pull request guide](../repo_governance/PULL_REQUESTS.md).
- **Complete coverage.** Inspect every applicable item, following every page
  of a paginated list. If pagination, permissions or missing evidence stop
  you, say so in the report; do not claim a complete audit.
- **Evidence, not appearance.** Verify each claim against the current code,
  the merged pull requests, CI and test evidence, linked discussions and
  dependencies. A title, a label, a ticked box, a merged pull request or a
  green run alone proves nothing about an acceptance criterion.
- **What the audit changes.** Correct issue and pull request metadata,
  descriptions, checkboxes and comments written on my behalf directly when
  the evidence supports it.
- **What the audit never does.** It merges no pull request, changes no
  repository setting, deletes no branch, publishes no release and changes no
  code. A finding that needs one of these goes in the report, and in an issue
  when it needs work.
- **Unclear items.** Leave an ambiguous item, or one that needs a decision,
  as it is, and report exactly what is uncertain.
- **No noise.** Make no speculative change and post no duplicate comment.

## 1️⃣ Closed issues in open milestones

For every closed issue assigned to an open milestone, check that its scope
and acceptance criteria are really met. Review:

- the title and description;
- the labels, including the priority and `blocked`;
- the assignee: an actionable issue is assigned to me;
- the milestone: an issue whose change merged into `release/X.Y.Z` belongs
  to milestone `vX.Y.Z`;
- the acceptance criteria and their checkboxes;
- every comment, linked pull request and piece of evidence;
- its dependencies and any follow-up it promised.

Correct stale or inaccurate metadata and text. Reopen an issue closed too
early, and say precisely what remains. Keep completed work, duplicates,
rejected proposals and work explicitly deferred or abandoned apart, each with
its own closing reason. Never tick a criterion you have not verified.

## 2️⃣ Open issues

For every open issue, decide whether it is still outstanding work, from its
title, body, labels, assignee, milestone, checkboxes, comments, dependencies
and any implementation evidence.

- Update a stale description or progress record.
- Keep `blocked` only while a named prerequisite is unresolved, and remove it
  once all are resolved, as the
  [dependency section](../repo_governance/ISSUE_LABELS.md#dependency-section)
  rules say.
- Fix broken or obsolete links.
- Close an issue only on evidence, with the right reason. For a duplicate,
  name the canonical issue and carry over any requirement only the duplicate
  had.

## 3️⃣ Pull request comments and review discussions

Review the comments and review threads of open and closed pull requests for
wrong claims, stale instructions, unresolved findings and misleading
statements about the implementation, the tests or the review status.

- Check review evidence against the commit SHA it names. Keep self-review,
  automated review, CI and manual testing apart.
- Correct a comment written on my behalf where it is wrong. A statement that
  was accurate when written stays; add what changed later instead of
  rewriting it.
- Never edit another contributor's words or a bot's status summary so that
  it implies an approval or a completion.
- Resolve a review thread only when its concern is fixed or explicitly
  accepted, with the evidence in the thread. A finding that still needs a
  code change goes in the report, not into a resolved thread.

## 4️⃣ Notifications

Read every one of my GitHub notifications for this repository, with the discussions and
review threads it links.

- Handle what this audit covers.
- Mark a notification done only when no action remains, or its action is
  complete.
- Leave an actionable or blocked notification pending, and say in the report
  what remains.
- Leave other repositories' notifications alone.

## 5️⃣ Repository scope and wording

Check the repository's documentation, configuration, templates, issues, pull
request descriptions and comments for references copied from other
repositories.

- Remove or correct a reference that is irrelevant, incoherent or wrong.
- Keep legitimate dependencies, source attribution, provenance, licence
  notices, comparisons and cross-repository links.
- Correct text written on my behalf where it departs from the repository's
  writing rules.
- Keep authorship, responsibility and approval records. Never invent a
  statement or an approval in my name, and never alter a quotation or another
  contributor's words.

## 6️⃣ Report

Post one comment on the closeout issue with:

- the repository audited and the coverage: the counts of issues, pull
  requests, comments, review threads and notifications examined;
- the changes made, grouped by category, each with a link and a short
  reason;
- the issues reopened or closed, the dependencies corrected and the review
  threads resolved;
- the notifications marked done, and those left pending with what remains;
- the unresolved findings, with the issue that tracks each, and the
  recommended next actions;
- every access limit, and every item not examined.

Keep verified facts apart from assumptions.
