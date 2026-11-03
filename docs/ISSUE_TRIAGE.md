# 🏷️ Issue triage and planning

[![Audit: read-only](https://img.shields.io/badge/audit-read--only-1D76DB)](#automation)
[![Priorities: P1 reserved](https://img.shields.io/badge/P1-confirmed%20calculation%20bugs-B60205)](#priorities)

This guide owns the rules for issue metadata: priorities, milestones, target
end dates and dependencies. [`.github/labels.json`](../.github/labels.json)
owns the label catalogue; [CONTRIBUTING](../CONTRIBUTING.md) owns the
contribution workflow.

<a id="automation"></a>

## ⚙️ Automation

| Workflow | Purpose |
| --- | --- |
| Issue triage (`.github/workflows/issue-triage.yml`) | Audits every open issue on issue and milestone events and on manual runs. Reports a missing assignee, priority, open milestone, target end date or dependency section, a target end date not before the milestone's due date, labels outside the catalogue, and inconsistent `blocked` status. Read-only: it never changes an issue. |
| Keep milestones issue-only (`.github/workflows/pr-milestones.yml`) | Removes any milestone set on a pull request, in a daily scan of every pull request, on a manual run of the default branch, and when the workflow changes on `main`. It never checks out repository code. |
| Static repository checks | Runs the triage tests and checks that the issue forms set the default priority, assign the maintainer and ask for the target end date and dependencies. |

The repository gate prohibits `pull_request_target`, so milestone cleanup does
not react to each pull-request event: a milestone set on a pull request lasts
until the next daily scan.

## 📋 Rules for every open issue

- Assigned to `danielep71`, the maintainer the issue forms assign.
- Exactly one priority: `P1`, `P2` or `P3`.
- An open milestone with a due date.
- A target end date earlier than that due date.
- A dependency section, and `blocked` exactly while a prerequisite it names is
  unresolved.
- Only labels from the catalogue.

<a id="priorities"></a>

### Priorities

| Label | Use |
| --- | --- |
| `P1` | Reserved for a confirmed calculation bug: wrong results for valid inputs. Only together with `bug`; never for release, setup, documentation or tooling work, however urgent. |
| `P2` | Material correctness, robustness or maintainability defect. Default of the bug form. |
| `P3` | Non-blocking quality, testing, documentation or technical-debt work. Default of the feature and documentation forms and of Dependabot pull requests. |

A form sets the default priority; triage confirms or changes it. Raise a bug to
`P1` only once the calculation defect is confirmed. Use the milestone, not the
priority, to say that work blocks a release.

### Target end date

An ISO `YYYY-MM-DD` calendar date (Europe/Rome) earlier than the due date of
the issue's milestone. The forms ask for it in the *Target end date* field. On
other issues, add it during triage as a line in the issue body, usually in a
*Schedule* section next to the start date:

```text
- Start date: YYYY-MM-DD
- Target date: YYYY-MM-DD
```

`- Target end date: YYYY-MM-DD (Europe/Rome)` is accepted too.

When both are present, that line takes precedence over the form field, so edit
the line to reschedule. Triage reports a missing or invalid date
(`target-date`), and a milestone without a due date or a target end date on or
after it (`milestone-due`).

### Dependencies

Give every issue a section whose heading contains *Dependencies*. Name each
prerequisite as an issue link (`#12`) or on a line starting
`External condition:`; write `None` when nothing blocks the work. List only the
issues this one depends on, never the issues that depend on it. A named pull
request is a reference, not a prerequisite. Add `blocked` while a named
prerequisite is open or an external condition applies, and remove it once all
are resolved.

### Milestones and pull requests

Milestones belong to issues only. A pull request carries the labels of the work
it delivers and links its issue; the issue carries the milestone. Dependabot
pull requests carry `ci`, `tests` and `P3` and an assignee; link each to an
issue when it belongs to planned work.

## 🖥️ Local audit

The audit needs Python 3.13+ and Node.js, because it resolves the label
catalogue with the synchronization script. To audit the live issues with a
GitHub CLI login:

```sh
GITHUB_TOKEN="$(gh auth token)" python3 tools/check_issues.py --repository danielep71/K-PRICING
```

To audit an export instead, pass a JSON array of issues and pull requests, open
and closed, in GitHub REST format, and name the repository so dependency URLs
resolve:

```sh
python3 tools/check_issues.py --input issues.json --repo-name danielep71/K-PRICING
```

To check only the issue forms and the label catalogue, offline:

```sh
python3 tools/check_issues.py --check-config
```
