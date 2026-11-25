# Repository instructions

## Documentation voice

- Write my decisions and responsibilities in first person: “I”, “me”, and “my”.
- Do not describe me in third person as “the maintainer” or “the owner”.
- Address instructions for contributors and template users directly to the reader.
- These rules also apply to issue and pull request text and comments written on my behalf: do not call me "the owner", "the user" or by name there; use direct, impersonal or first-person wording.

## Repository scope

- Describe this repository's current files, configuration, and behavior.
- Do not insert references to other repositories, copied-from history, unrelated projects, or conversation context into operational documentation.
- Make reusable guidance self-contained, so that it stays accurate in every repository that holds a copy of it.
- Keep necessary platform/dependency documentation links and third-party license notices. Put licensing attribution in dedicated license files, not operational guides.
- Do not invent project history, capabilities, or dependencies.

## Branch names

Use descriptive prefixes such as `docs/`, `fix/`, `feat/`, or `chore/`. Do not use `codex/` or `claude/`.

## Commit authorship

Author every commit as me, not under an assistant's identity. Before committing, set the repository's Git `user.name` and `user.email` to the name and GitHub noreply address (`<id>+<login>@users.noreply.github.com`) that my own commits in this repository use. Do not add a `Co-Authored-By` trailer that credits an assistant. Never rewrite existing history to change authors.

## Audience separation

Keep editorial rules, assistant instructions, branch-naming preferences, and collaboration conventions in this file. Reader-facing documentation must explain the repository and its usage, not our writing process. Apply the requested voice naturally; do not explain that prose is written “in my voice”.

## Review findings

When a review of this repository finds something to fix, record each finding, or each group of findings that one change will fix, as an issue before or instead of fixing it silently. Every such issue must meet the [triage rules](docs/repo_governance/ISSUE_LABELS.md#triage-rules) of **Issue triage**. Give it the type labels that fit and the open milestone with the earliest due date, and write its dates as the `- Start date:` and `- Target end date:` lines of that guide.

Describe the finding with its file and line, the evidence, why it matters, and the proposed fix, so the issue can be resolved without the review that found it.

## Domain defaults

This repository holds, or provides the starting point for, a regulatory risk calculation workbook such as SA-CCR, SFT, FRTB, LCR, or NSFR. Excel workbooks, exported VBA source, regulatory source copies under `docs/methodology/sources/`, the P1 definition for confirmed calculation bugs, the `traffic-history` branch, and related reserved paths are intended defaults. Do not report them as leftovers from another project.

The repository is licensed under MPL-2.0 through the root `LICENSE` unless that file says otherwise. Keep `docs/licenses/` for the files listed in its third-party notice.

## Reusable defaults

Preserve useful policies and assumptions for file types, tooling, and reserved paths even when this repository does not use them yet. Adapt project-specific names and historical references without stripping reusable capabilities. Describe reserved paths as conventions, not as files or directories that already exist.

<!-- template-only:start -->
## Template maintenance

This section applies only to the template repository. In a repository generated from the template, delete it together with its markers; **Repository integrity** fails until you do.

- This repository is the template that new workbook repositories are created from. Keep its files generic: each one must still be correct after it is copied into a new repository.
- Word each file so that it is true both here and in a repository created from the template: write "empty until the project adds its own", not "empty in the template". Put a step that only the template needs in `docs/setup/REPOSITORY_SETUP.md` or in this section, or let the initializer adapt the sentence and list the edit. The `template-wording` check of **Repository integrity** fails on the phrases its tool lists.
- When releasing this repository, run the repository checks of `RELEASING.md` with `--template true`, not `--template false`.
- My reference repositories are `danielep71/K-SFTs` (primary) and `danielep71/K-SACCR` (secondary). Compare repository settings, issue labels, issue forms, and automation with them before changing the template.
- My other workbook repositories, `danielep71/K-FRTB-ASA`, `danielep71/K-FRTB-SSA`, `danielep71/K-LCR`, and `danielep71/K-NSFR`, are also compared: adopt what they do better, and report where they lag behind the template.
- My portfolio hub, `danielep71/K-Risk-Engines`, holds no workbook but shares labels, traffic collection, link checks and review rules with the template; compare those parts with it too.
- Port improvements in generic form. Do not copy their names, links, issue numbers, or history into template files; only this section and `docs/licenses/` may name them.
- This repository tracks its own work with milestones, starting with `v0.1.0 - Initial setup`. Issues assigned to a milestone follow the same triage rules as in generated repositories, including a target end date before the milestone's due date.
- The rule that links each bot pull request to a tracking issue applies only in generated repositories. Here, merge a Dependabot pull request once **Repository integrity** passes and my review of its final diff is recorded in a comment on it, as `docs/repo_governance/PULL_REQUESTS.md` describes.
- From a session on this repository, change only this repository. Read the reference repositories and any other repository only to compare them; for changes they need, write a message for a session dedicated to that repository instead of committing, pushing, or posting there.
<!-- template-only:end -->
