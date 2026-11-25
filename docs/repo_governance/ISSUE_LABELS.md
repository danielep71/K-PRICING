# 🏷️ Issue labels & automation

![Labels](https://img.shields.io/badge/labels-20-1D76DB)

This repository defines 20 issue labels, each with a name, description, and colour. `.github/labels.json` is authoritative; edit it through a pull request. Its catalogue has a `core` list and optional `overlays`: named label sets that `selected_overlays` enables. Synchronization and drift detection resolve the core plus the selected overlays; every name is unique across the core and every overlay, and an unknown or duplicated selection is a validation error. In this repository the overlays are empty and none is selected, so the resolved catalogue is the core. `prune: true` deletes live labels outside the catalogue, but only unused ones. Before it changes anything, synchronization checks every label it would delete against all issues and pull requests, open and closed. It stops if one still carries the label. A name that contains a comma cannot be passed to the API's label filter, so for that name it reads the labels of every issue and pull request, up to 10,000, and stops if a response is malformed or the limit is reached. Remove a label that is still used from its issues and pull requests, or add it to the catalogue or select its overlay, then rerun. A dry run, `node .github/scripts/labels-sync.mjs --mode reconcile --dry-run`, changes nothing but reads the same usage: its summary counts the deletions still in use, marks each one as a deletion a real run would refuse, and ends with the refusal text, while still exiting 0. For each such deletion it names up to ten of the issues and pull requests that use the label, in number order, and adds `and others` when more do, such as `used by #3, #7 and others`; a real run's refusal names them too. A plain name costs one request, a page of the eleven oldest items that carry it; for a name with a comma, the numbers come from the same read of every issue and pull request. A response item without a positive whole number stops the run like a malformed label. `tests/test_labels_sync.py` runs both against a local fake API.

## ⚙️ Automation

| Workflow | Purpose |
| --- | --- |
| Sync issue labels | Reconciles labels after a catalogue change lands on the default branch or a manual run on the default branch; pull requests are validated by **Repository integrity**, which runs the script's self-test and validates the catalogue. Refuses to delete a label in use. Verifies the exact live result. |
| Detect issue-label drift | Checks weekly and manually; **Repository integrity** runs its offline fixtures on every pull request. Reports drift without modifying labels. Each report is cross-checked against the reconciler's plan on the same snapshot: the script writes the live labels it read to a temporary file, runs `labels-sync.mjs --mode plan` on it, and fails (status `fail`) if the two plans differ; the evidence then carries both plans under `cross_check`. |
| Issue triage | Audits open issues against the [triage rules](#triage-rules) on issue and milestone events or manual runs. Changes no issue. |
| Keep milestones issue-only | Removes any milestone set on a pull request when the pull request is opened, reopened, edited, or given a milestone, and in a weekly and manual scan of every pull request. It reads each pull request again after the change and fails if the milestone is still set; its job summary lists each pull request and the milestone removed. If GitHub gives a pull-request run a read-only token, the run warns instead of failing, and the next weekly or manual scan removes the milestone. It never checks out pull-request code. |
| Dependabot (`.github/dependabot.yml`) | Opens weekly grouped pull requests for the SHA-pinned GitHub Actions, the hash-locked configuration parser, and the hash-locked Python quality tools, labelled `ci` and `P3` and assigned to me. Nothing merges automatically. |

Among these workflows, only synchronization has `issues: write` and only *Keep milestones issue-only* has `pull-requests: write`; outside them, only the optional traffic-alert job writes issues ([TRAFFIC.md](TRAFFIC.md)). Synchronization also has `pull-requests: read`: without it the issues API omits pull requests, and the check before a deletion would miss a label only pull requests carry. `tests/test_workflow_controls.py` fails if a job outside its `WRITE_PERMISSIONS` list gains a write permission, or a listed job gains one the list does not give it; the list also lets code scanning, the Scorecard analysis, the release workflow's attestation and draft, and traffic collection write what each needs. PR validation is read-only. Actions are pinned to full commit SHAs.

## 🚀 Setup and ownership

1. Run **Sync issue labels** from the default branch to bring the live labels to the catalogue: once Actions is enabled, and again whenever the live labels must be brought back to it.
2. To add labels for this repository, add them as an overlay in `.github/labels.json`, such as one overlay per regulatory regime: an entry of `overlays` whose name maps to a label array sorted case-insensitively by name, with the same fields as the core. List the overlay's name in `selected_overlays`, sorted case-insensitively; an overlay that is not selected adds nothing. Then add each selected label to the catalogue and usage tables below and update the label counts, the badge included, to the resolved total. Leave the `core` list as it is. **Repository integrity** checks all three.
3. My GitHub account, `danielep71`, is the assignee in the four issue forms (`assignees`) and in `.github/dependabot.yml`; the triage check reads it from `bug.yml`. If someone else maintains this repository, replace my account in those five files. **Repository integrity** fails if they disagree.
4. Create project milestones with due dates, and give each open issue what the [triage rules](#triage-rules) require.
5. Label validation and the offline drift fixtures run in **Repository integrity** on every pull request, to any branch. Live synchronization runs only on the current default branch, including after a rename.

**Keep** the label manifest, scripts, all four workflows, issue forms, triage checker/tests, and this guide after bootstrap. They provide ongoing maintenance and must not be deleted with one-time repository setup files.

## 🧭 Using the labels

Every issue has exactly one priority and the type labels that fit the work. Pull requests carry the labels of the work they deliver; milestones belong to issues only, and *Keep milestones issue-only* removes any milestone set on a pull request.

Pull requests opened by Dependabot or another bot carry the labels of the work and no milestone. During triage, link each one to an existing issue or a new tracking issue, and give that issue the milestone. Before merging one, record my review of its final diff as [PULL_REQUESTS.md](PULL_REQUESTS.md) describes.

### Priorities

| Label | Use |
| --- | --- |
| `P1` | Reserved for a confirmed calculation bug: wrong results for valid inputs. Only together with `bug`; never for setup, documentation, or tooling work. |
| `P2` | Material correctness, robustness, or maintainability defect. Default of the bug and private security contact forms. |
| `P3` | Non-blocking quality, testing, documentation, setup, or technical-debt work. Default of the enhancement and documentation forms and of Dependabot pull requests. |

A form sets the default priority; triage confirms or changes it. Raise a bug to `P1` only once the calculation defect is confirmed.

Every open issue also needs a start date and a target end date: ISO `YYYY-MM-DD` calendar dates (Europe/Rome). The start date is when the work is planned to start, not after the target end date; the target end date is earlier than the due date of the issue's milestone. The bug, enhancement, and documentation forms ask for them in the *Start date* and *Target end date* fields. On other issues, including private security contacts, add them during triage as lines in the issue body:

```text
- Start date: YYYY-MM-DD (Europe/Rome)
- Target end date: YYYY-MM-DD (Europe/Rome)
```

When both a line and a form field are present, the line takes precedence, so edit the line to reschedule. Each date takes one line. The `start-date`, `target-date` and `milestone-due` [triage rules](#triage-rules) check the dates.

### Types

| Label | Use |
| --- | --- |
| `bug` | A confirmed defect. The bug form sets it. |
| `enhancement` | New capability or improvement. The enhancement form sets it. |
| `documentation` | Documentation work. The documentation form sets it. |
| `behavior-change` | Add to work that changes externally observable behavior, such as computed results or input and output layouts. |
| `refactor` | Restructuring that preserves supported behavior. If supported behavior changes, use `behavior-change` instead. |
| `tests` | Regression coverage, test harnesses, fixtures, or manual validation. |
| `ci` | Workflows, automated checks, and dependency updates, including every Dependabot pull request. Release certification belongs to `release`. |
| `repository` | Repository structure, governance, metadata, or tooling. |
| `release` | Release preparation, certification, evidence, packaging, or publication. |
| `security` | Hardening and supply-chain work, and private security contact requests, which never contain vulnerability details. Never for an undisclosed vulnerability: report it as `SECURITY.md` describes. |

### Status and triage

| Label | Use |
| --- | --- |
| `blocked` | A prerequisite named in the [dependency section](#dependency-section), as an issue link or an `External condition:` line, is unresolved. Remove it once all are resolved; ready issues carry no status label. In the same edit, delete the `External condition:` line of each resolved condition, or replace the section with "None" if nothing else is named: **Issue triage** treats every such line as unresolved. |
| `question` | I need information or a decision before the work can proceed. |
| `help wanted` | I welcome outside investigation, implementation, validation, or review. |
| `good first issue` | Well-scoped task suitable for a first contribution. |
| `duplicate` | Close the issue and link the issue or pull request that already covers the work. |
| `invalid` | Close the issue: it is incomplete, unsupported, or not actionable as submitted. |
| `wontfix` | Close the issue: I have decided not to do this work. |

### Dependency section

**Issue triage** reads the section under the first level-2 or level-3 heading whose text contains "Dependencies", up to the next level-2 or level-3 heading. In it:

- name a prerequisite issue as `#12` or by its URL in this repository; it blocks the work while it is open;
- write an external blocker on a line that starts with `External condition:`, in any letter case, optionally after a `-` or `*` list bullet, followed by what you are waiting for; it counts as unresolved until you delete the line;
- a pull request named there is a known reference, never a prerequisite: to wait for one, name the issue it resolves;
- write `None` when nothing blocks the work.

A section with nothing but blank lines, or only `_No response_`, which GitHub writes for an empty form field, is a `dependencies` finding: write `None`, link each prerequisite or add an `External condition:` line.

Open issues whose dependency sections name each other in a cycle can never be unblocked by finishing work. **Issue triage** reports each issue of the cycle (`dependency-cycle`) until one dependency is removed or replaced with an `External condition:` line.

A blocked issue, for example:

```markdown
## Dependencies and blocking status

- #12, which defines the input layout
- External condition: waiting for access to an Excel host
- Related: #34, the pull request that started this work
```

<a id="triage-rules"></a>

## 🔍 Triage rules

**Issue triage** audits every open issue on issue and milestone events and on manual runs. It is read-only: a failed run is the signal, and I resolve each finding by hand. Each finding names one of these rules, which `tools/check_issues.py` defines:

<!-- BEGIN generated: triage-rules -->
| Rule | What every open issue needs |
| --- | --- |
| `assignee` | An assignee: the account that the issue forms assign, read from `bug.yml`. |
| `priority` | Exactly one of the priority labels `P1`, `P2` and `P3`. |
| `p1-reserved` | `P1` only together with `bug`: it is reserved for a confirmed calculation bug. |
| `milestone` | A milestone, and an open one. |
| `target-date` | One valid ISO `YYYY-MM-DD` target end date (Europe/Rome), in a planning line or the issue form's field. |
| `milestone-due` | A milestone with a due date, and a target end date before it. |
| `start-date` | One valid ISO `YYYY-MM-DD` start date (Europe/Rome), not after the target end date. |
| `dependencies` | A dependency section that is not empty, in which every issue or pull request named exists. |
| `blocked` | The `blocked` label exactly while the dependency section names an open prerequisite issue or an `External condition:` line. |
| `catalogue` | Only labels of the catalogue that `.github/labels.json` resolves: the core and the selected overlays. |
| `dependency-cycle` | No cycle among the dependency sections of open issues. |
<!-- END generated: triage-rules -->

## ✅ Closing issues

Close an issue as completed only when linked evidence meets its acceptance criteria, and check an acceptance box only when that evidence demonstrates it. When the issue is closed, add a **Resolution** section to its body with:

- the change: the pull request and the merge commit;
- the checks actually run, with their results, including the Excel run and its case and assertion counts when the change touched VBA source, a workbook layout or a test case that runs in Excel ([Excel run](PULL_REQUESTS.md#-excel-run));
- what was not validated, any deviation from the proposed fix, and any accepted limitation.

An issue closed after its change merged into a release branch `release/X.Y.Z` goes into milestone `vX.Y.Z`, the version that contains it, whatever milestone it was planned in; say so in the Resolution when you move it.

Close for another reason (duplicate, not planned, or superseded) only with an explicit explanation; the `duplicate`, `invalid`, and `wontfix` rows above say which label applies. Record material scope or policy decisions in the issue, pull request, or commit description.

## 🎨 Catalogue

The table below lists the descriptions defined in `.github/labels.json`, which synchronization uses as its source of truth: the `core` list plus every selected overlay.

| Label | Colour | Description |
| --- | --- | --- |
| `behavior-change` | `#C24E00` | Intentional externally observable behavior change |
| `blocked` | `#B60205` | Cannot proceed until a named dependency is complete |
| `bug` | `#D73A4A` | Confirmed defect in behavior, correctness, error handling, or a documented contract |
| `ci` | `#24292F` | Automated quality checks, workflow maintenance, or dependency updates |
| `documentation` | `#0075CA` | Documentation for users or for my repository maintenance |
| `duplicate` | `#CFD3D7` | An existing issue or pull request already covers this work |
| `enhancement` | `#A2EEEF` | New capability or meaningful improvement to supported behavior |
| `good first issue` | `#7057FF` | Well-scoped task suitable for a first contribution |
| `help wanted` | `#008672` | I welcome external investigation, implementation, validation, or review |
| `invalid` | `#E4E669` | Report is incomplete, unsupported, or not actionable as submitted |
| `P1` | `#B60205` | Confirmed calculation bug: wrong results for valid inputs (reserved for this case) |
| `P2` | `#D93F0B` | Material correctness, robustness, or maintainability defect |
| `P3` | `#FBCA04` | Non-blocking quality, testing, documentation, setup or technical-debt work |
| `question` | `#D876E3` | I need further information or need to make a decision |
| `refactor` | `#6F42C1` | Internal restructuring intended to preserve supported behavior |
| `release` | `#B60276` | Release preparation, certification, evidence, packaging, or publication |
| `repository` | `#1D76DB` | Repository structure, governance, metadata, tooling, or maintenance |
| `security` | `#8B0000` | Security boundary, vulnerability handling, supply chain, or hardening |
| `tests` | `#006B75` | Regression coverage, test harnesses, fixtures, or manual validation |
| `wontfix` | `#6E7781` | I have decided not to implement or continue this work |

## Local issue audit

The audit needs Python 3.11+ and Node.js, because it validates the catalogue with the synchronization script. To audit the live issues with your GitHub CLI login:

```sh
GITHUB_TOKEN="$(gh auth token)" python tools/check_issues.py --repository OWNER/REPOSITORY
```

To audit an export instead, pass a JSON array of issues and pull requests, open and closed, in GitHub REST format, and name the repository explicitly so dependency URLs resolve:

```sh
python tools/check_issues.py --input issues.json --repo-name OWNER/REPOSITORY
```

Triage accepts only the labels that `.github/labels.json` resolves: the core plus the selected overlays.
