# 📜 Changelog

Notable changes to this repository, from a user's point of view. The format
follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions
follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html). It is
neither a commit log nor release evidence.

## 🧭 Maintenance policy

- Add a material change under `[Unreleased]` in the pull request that makes
  it. A change nobody would notice, such as a typo or a refactor with no
  effect, needs no entry.
- Describe the observable result: what exists, what a user or contributor
  can now do, what changed in a contract, and what to do about it.
- Link the issue or pull request that holds the detail.
- Record only evidence that was produced. A static check is not an Excel run:
  say what was and was not run in Excel.
- Once a version has been released, edit an existing entry of the next
  version rather than adding one that describes a change to its unreleased
  work: an entry describes what the release contains. The first release's
  section also records, as **Changed** and **Fixed**, how its additions
  changed while it was prepared.
- Entries move into a dated section only when a release is prepared, as
  [RELEASING.md](RELEASING.md) describes. After that, until the tag exists,
  a fix for a release gate that failed on the candidate records its change in
  the dated section being released
  ([step 2](RELEASING.md#2-integrate-and-freeze-the-candidate)). A published section is not
  rewritten; a demonstrable error is corrected in review, with a line
  `Corrected after release on YYYY-MM-DD: <reason>` in that section.

<a id="date-and-version-rules"></a>

### Date and version rules

`python tools/check_changelog.py` checks the headings, dates, order, link
references and `VERSION`, and that this introduction, the *Current status* of
[RELEASING.md](RELEASING.md#current-status) and the README describe the same
release state; **Repository integrity** also checks that every link reference
points to this repository. Once a release's tag exists, its section must equal
that section at the tag, unless it holds the dated correction line above; a
correction is reported as a note. What a version number means is checked
in review.

| Rule | Standard |
| --- | --- |
| Version | `MAJOR.MINOR.PATCH`, without a leading `v` in headings |
| Release heading | `## [X.Y.Z] - YYYY-MM-DD`, with a real calendar date |
| Order | `[Unreleased]` first, then releases from newest to oldest, by version and date |
| Link references | One for `[Unreleased]` and one for every release; a release's points to `.../releases/tag/vX.Y.Z` |
| `0.y.z` | Initial development: no compatibility is promised between versions |
| Patch, minor, major | From `1.0.0`: a compatible correction, a compatible capability, an incompatible change to a calculation result, an input or output layout, or the supported VBA API |

| Category | Use for |
| --- | --- |
| **Added** | New capabilities, files, checks or tests |
| **Changed** | Changes to existing behaviour, contracts, tooling or documentation |
| **Fixed** | Corrected defects |
| **Security** | Safely disclosed security corrections |
| **Validation** | Evidence actually produced |
| **Known limitations** | Deliberate, unresolved boundaries |

Use only the categories an entry needs.

## [Unreleased]

### Added

- `tools/check_vba_layout.py`, run by **Repository integrity** as `vba-layout`,
  checks the new *Layout* section of the VBA house style on every component:
  banners of exactly 79 characters, and a header with nine fields, in order
  and each with text, on the component and on every procedure. VBA written
  from now on follows that layout, and every repository created from the
  template inherits the guide and the check. Static check only: no Excel run
  (#591). It reads the same optional baseline as `format_vba.py --check`,
  `tools/vba_layout_baseline.json`, so that a project importing existing modules
  brings them into layout one at a time: a listed module may stay out of
  layout, an unlisted one fails, and a listed module that both tools find in
  layout fails until its entry is removed. Both tools read the modules under
  the source directories of the house style (#593).
- The **Closing keywords** workflow runs `tools/check_closing_keywords.py` on
  each pull request. Into a release branch, it fails on any closing keyword,
  such as `Resolves #N`, in a commit, the title or the body. Into the default
  branch from a release branch, it fails on a keyword that names an open
  issue, unless the body lists the issue on a `Reopen after merge: #N` line;
  each finding names the commit, the keyword as written and the issue. It
  reads issue states with the workflow token and is not a required check
  (#497).
- `tools/check_excel_run.py --carried-to SHA` carries the evidence of an
  Excel run to a later commit whose VBA components differ from the run's
  only in comments or layout, so that correcting a comment no longer needs a
  new Excel run for a pull request. The new `tools/vba_code_identity.py`
  compares the components' logical lines without comments, with each run of
  white space outside strings read as one space and the letter case of code
  ignored, and keeps conditional-compilation directives. It refuses non-ASCII input and any
  component that uses `[` or `]` outside a string. Every other run input,
  each form's `.frx` included, must keep its Git blob. The report labels the
  result `Carried from <run commit>`, with the components carried, and
  writes nothing into the evidence folder. Carried evidence never certifies
  a release: `--equivalent-to` keeps its blob identity, and a release still
  needs a run of the candidate or one `--equivalent-to` proves identical.
  The VBA lexer no longer reads `Rem` or a colon inside a `[...]` expression
  as a comment or a statement separator (#499).
- `python tools/format_vba.py --check` reads an optional baseline,
  `tools/vba_layout_baseline.json`, that lists the modules not yet in
  layout, so that a project importing existing modules can bring them into
  layout one at a time. The check fails on an unlisted module out of layout,
  on a listed module back in layout until its entry is removed, and on a
  listed path that is not a tracked module; a missing or malformed baseline
  is an error with exit status 2. `--baseline FILE` reads another file and
  `--no-baseline` ignores it. `--write` says to re-import each rewritten
  module and to remove its baseline entry. The template has no baseline,
  since its modules are in layout; the house style describes how to create
  one (#502).
- **Repository integrity** runs a new check, `verification-depth`, which
  fails while a check that can fail the run lacks an executable negative
  proof. `tools/verification_depth.json` lists each such check with what it
  runs, the unittest ids of the tests that show it rejecting a bad input, or
  its probe, the violations they show rejected and a state from a closed
  list. The checks come from `tools/check_repository.py` as CI runs them,
  less the advisory ones, so a pull request that adds a check without an
  entry, renames or deletes a proof test, or leaves an entry for a removed
  check fails. New tests prove that the `tests` check fails on a failing or
  unimportable test, and that the three label checks fail on a defective
  script or catalogue (#503).
- The **Repository integrity** check `restated-counts` runs the new
  `tools/check_restated_counts.py`. It fails when a tracked Markdown file
  other than this changelog states, outside a generated block, a count that
  has a single source: the example suite's cases, assertions, modes and
  failure entry points, the integrity checks and the static checks of VBA
  source, the data contract's datasets and the validator's finding codes. It
  matches a count in digits or in words before the phrase that names the
  source, whatever its value, so a stale copy fails too, and names the file,
  line, phrase and source. A reviewed restatement stays only in its list of
  allowed restatements, with a reason. The installation guide, the harness
  guide, the VBA house style and the setup guide no longer restate these
  counts (#504).
- An accuracy contract may declare an optional `target` beside its
  `threshold`: a finite number of at least 0, no larger than the threshold,
  that `tools/check_accuracy.py` reports but never blocks on. The threshold
  stays the guarantee that fails a run, and a registry without a target keeps
  `schema_version` 1. Each measurement line, which `tools/check_excel_run.py`
  passes on as an `Accuracy:` line, now names the threshold and the target,
  or `no target`, and ends with `target missed` when the worst error is above
  the target but within the threshold, or `threshold exceeded` when it fails.
  A malformed target, or one larger than the threshold, is a finding. A
  frozen contract's target may change under the same decision; its threshold
  keeps its rules. The illustrative `ACC-01` declares a target of 0 (#505).
- The accuracy registry may declare invariants in an optional `invariants`
  list, and keeps `schema_version` 1 without it. A `bounds` invariant keeps a
  quantity within a minimum, a maximum or both, with no tolerance; a
  `monotone` one keeps a quantity from falling, or from rising, along an
  ordered list of cases or variants by more than the two outputs' own
  tolerances, and counts a smaller reversal without failing; a `sum` one
  keeps declared parts adding up to another quantity, or to a constant, in
  each case, variant and scope, within the sum of their tolerances.
  `tools/check_accuracy.py` reports a malformed invariant, an unknown
  calculation, case or quantity, and expected values that break an
  invariant. With `--results`, a result that breaks one, or a missing result
  it reads, fails whatever any contract's status, and each invariant prints a
  line, which `tools/check_excel_run.py` passes on as an `Accuracy:` line.
  The summary line counts the invariants beside the contracts. The template
  declares one illustrative invariant, `INV-01`: the illustrative total is
  at least 0 (#505).
- `tests/accuracy/holdout_exposures.json` records each first measurement of
  a frozen accuracy contract's holdout under its decision: the contract, the
  decision, the number of exposures, the commit and SHA-256 of each results
  file, the SHA-256 of the canonical contract entry and its threshold, the
  SHA-256 of each holdout case's fixture and expected file, and the length
  and SHA-256 of an optional `design` document the contract may name. With
  `--results`, `tools/check_accuracy.py` no longer measures or prints a draft
  contract's holdout, and seals a frozen contract's holdout, as a finding,
  until `--record-holdout ACC-NN` records its exposure; that mode refuses a
  draft contract, a second first exposure under the same decision, a
  missing, malformed or inconsistent record, a failing registry check and
  incomplete results, and writes nothing then. The check fails on a
  malformed record, an exposed contract that is removed, exposures recorded
  at the newest tag that are dropped, and, under the decision of the last
  exposure, a changed entry, a lower threshold, a changed holdout fixture or
  expected file, or a design document that was rewritten rather than
  appended to. Holdout lines print the number of exposures;
  `tools/check_excel_run.py` measures a recorded holdout as a regression and
  records nothing. The template ships an empty record: its illustrative
  `ACC-01` stays draft, so its holdout stays sealed (#505).
- The holdout exposure record is not a run input of
  `tools/check_excel_run.py --equivalent-to` or `--carried-to`: the
  accuracy measurement reads it from the commit the run certifies, so the
  run that exposed a holdout also certifies the commit that records it,
  when every other run input is identical (#526).
- The **Repository integrity** check `template-wording` runs the new
  `tools/check_template_wording.py`. It fails, with `--template true` and
  `false` alike, when a tracked text file says "in the template", "the
  template's", "template repository" or "generated repository", phrases that
  describe the template rather than the repository the file is in. It names
  the file, line and phrase. The changelog, the license texts, the setup
  guide, the initializer and the check itself are not read, nor is the
  template-only section of `AGENTS.md`; the reviewed uses that stay, such as
  "the template's" meaning the workbook template, are listed with their
  reasons, and one that matches nothing fails the check (#491).
- `python tools/check_workbook_privacy.py FILE...` reports what an Office
  Open XML package says about where it was saved and by whom: a OneDrive or
  SharePoint path or URL, an absolute path under a user profile, a network
  share, a `file:` URL or any `x15ac:absPath`, in any XML part or in the
  text of the VBA project, stored or compressed, in its declared code page,
  multi-byte ones such as cp936 included, or as UTF-16LE, for profile names
  in any script, Chinese characters included; a person or organization in
  the document properties, `fileSharing`, threaded-comment people, comment
  authors and shared-workbook revisions, each part found by its usual path
  or by its content type; a Microsoft Purview sensitivity
  label (`MSIP_Label_*` properties, `docMetadata/` parts); and SharePoint
  library properties and schemas (`docProps/custom.xml`, `customXml/`). Each
  finding names the file, the part, the kind and a hint, never the value. It
  exits with 1 on a finding and 2 on a file it cannot read. **Repository
  integrity** now fails a committed Office package with any finding, through
  `tools/check_binaries.py`, except the regulatory source copies, and an XML
  part that declares a DTD is no longer parsed. A part is read as XML when its
  name ends in `.xml` or `.rels`, its declared content type is XML or it has
  one of these roles, so a renamed part hides nothing. `tools/release.py
  verify-published` fails when an attached `--asset` that is an Office
  package, or is named like one, has a finding, whatever its SHA-256. A
  committed or attached workbook that carries such data fails until it is
  cleaned (#546).
- `tools/initialize_repository.py` takes two options and does three more
  things. `--remove-traffic` removes the optional traffic collection: its
  workflow, guide, tool and test, the README's badges, the ruleset of the
  `traffic-history` branch in `docs/setup/repository-settings.json` (one
  already created on GitHub is left, and the preview says so), its files in
  the license notices, its `tools/verification_depth.json` entry and the
  rows and sentences of the guides that describe it, and it warns about a
  mention it cannot place; `--remove-traffic-badges` still removes only the
  badges. `--assignee LOGIN`, a validated GitHub login, becomes the assignee
  of the four issue forms and of `.github/dependabot.yml`, and the account
  `docs/repo_governance/ISSUE_LABELS.md` names. A template-only section may
  now hold its replacement in a `<!-- template-only:replacement ... -->`
  comment: the README's *Start a new repository* section and the template's
  wiki links are template-only, and give way to a table of the main guides.
  After a successful `--apply`, the initializer renders the generated blocks
  afresh and, last, deletes itself and its test; the preview lists that
  deletion. `readme-traffic-badges` runs only while the traffic workflow or a
  badge is there, `template-only-content` also checks `README.md`, and the
  settings tests no longer rely on the traffic ruleset (#571).

### Changed

- The VBA house style separates procedures with one blank line instead of
  two, and gives each component and procedure the same nine-field header
  under banners of 79 characters; `python tools/format_vba.py` applies the
  one blank line. The modules of `tests/modules/` are reformatted to the new
  layout with their code unchanged. Their blob ids change, so the Excel
  evidence of the harness and the example suite needs a new run, unless
  `--carried-to` carries it over the comment-only change (#591).
- **Repository integrity** runs on every pull request and on pushes to `main`
  and `release/**` only, no longer a second time on each push to a task
  branch. The `codeql-default-branch` check also fails while the workflow's
  push branches omit the default branch (#538).
- A sync pull request whose head is the integration merge commit, with an
  empty diff against the release branch and a tree the integration's
  final-head review covered, needs no new automated review;
  **Repository integrity** still runs on it (#539).
- The VBA source analysis and the Excel-run source comparison are faster,
  with the same findings: the conditional-compilation pass of a component is
  computed once and shared by the checks that read it, identical exported and
  committed modules are not compared line by line, and the commit's
  components are read by one `git cat-file --batch` instead of one `git show`
  each. The test suite runs in about three quarters of its time (#551).
- The settings baseline allows auto-merge (`allow_auto_merge: true`); it
  bypasses no required check or rule. A sync pull request under the review
  exemption may use it, and may carry a short record in place of the full
  template. No task pull request merges into the release branch while it
  waits (#552).
- Every open issue needs a start date (Europe/Rome), not after its target end
  date. The bug, enhancement and documentation forms ask for it; on other
  issues it is a `- Start date:` line. **Issue triage** reports a missing or
  invalid one, or one after the target end date (`start-date`), and the
  traffic alert writes `- Start date:` in place of `- Planned start:` (#557).
- **Issue triage** reports an open issue with two `- Start date:` lines, or
  two `- Target end date:` lines, with different dates (`start-date`,
  `target-date`). A line that differs from the form field is still the
  documented way to reschedule (#561).
- The `codeql-default-branch` check reads branch filters as GitHub does:
  patterns, `!` negations in order, `branches-ignore`, `pull_request_target`
  and a push filtered by tags alone. CodeQL and **Repository integrity** must
  still run on the default branch, and a malformed branch filter in any
  workflow is now reported (#562).
- While a release branch is active, commit messages, squash messages and task
  pull requests use only `Refs #N`, as `CONTRIBUTING.md`, the pull request
  guide and the pull request template now say: a closing keyword in a commit
  merged into a release branch closes its issue at the integration, whatever
  its state. `Resolves #N` stays for a pull request into the default branch
  while no release branch is active (#497).
- *Start a new repository* in the README says that **Use this template**
  copies the template's default branch, not a release tag, and asks to record
  the template commit a new repository came from; the setup guide's
  prerequisites say the same (#462). A new step leads to the setup guide's
  handover table for the files the initializer leaves as they are, such as
  the README's title, banner and wiki links, the social preview and the
  repository's **About** fields (#463).
- The journal of `tools/setup_repository.py --apply` and the planned changes
  of its `--evidence` record name the ruleset of each ruleset write, so that
  the writes to the shared `rulesets` endpoint can be told apart (#469).
- The setup guide and the README run the settings baseline on a task branch,
  since applying it protects the default branch: the journal, report and
  evidence it writes are committed there and merged through a pull request
  (#470).
- The setup guide and the README quote the `--evidence` path that holds
  `{date}`, so that the documented command also works in PowerShell (#471).
- The journal of `tools/setup_repository.py --apply` also records the ID
  that GitHub returns for each ruleset it creates or replaces (#479).
- The setup commands of the setup guide and the README are each on one line,
  without `sh` line continuations, so that they can be pasted into
  PowerShell (#478).
- The pull request review guide and README step 8 no longer say that Codex
  needs an environment for the repository: a repository generated from the
  template had its first pull request reviewed without one. If Codex asks
  for an environment instead of reviewing, an `@codex review` comment
  produces the review (#477).
- `tools/initialize_repository.py` also removes the template's evidence,
  which cites the template's own commits and issues: the rows of the
  *Tested environment* table of `INSTALLATION.md`, and the *Issue* column of
  the *Host behaviour* table of the validation plan, which becomes a dash.
  The validation plan's `HB-04` result no longer names a commit (#485).
- The traffic guide creates the `traffic-history` branch in a separate
  worktree with commands that work alike in PowerShell and `sh`, instead of
  `sh`-only commands with a temporary directory and a shell variable (#484).
- The social preview and README banner name the risk areas the template
  is for, market, credit, operational, counterparty, IRRBB and liquidity,
  instead of a list of regimes, and show one formula for each (#492).
- **Issue triage** reports each open issue in a cycle of open issues'
  dependency sections, which no finished work can unblock (#506).
- **Issue triage** reports an empty dependency section, one with nothing but
  blank lines or `_No response_`, as a `dependencies` finding: write `None`,
  link each prerequisite or add an `External condition:` line. An open issue
  with such a section now fails triage until it is filled in (#531).
- An issue's *Resolution* names the merge commit, the Excel run with its
  case and assertion counts when the change needed one, and any deviation
  from the proposed fix. An issue closed through `release/X.Y.Z` goes into
  milestone `vX.Y.Z`; the pull request guide, the closeout and the
  maintenance audit say the same (#507).
- A release tag whose commit holds a VBA component needs its Excel run.
  `tools/release.py evidence` checks the evidence folder and the workbook
  with `tools/check_excel_run.py --require-binding`, for the candidate or
  for an earlier commit with identical run inputs, and prints an
  `Excel evidence: <digest> (run <sha>).` line for the tag message. It never
  accepts carried evidence. `tools/release.py verify` fails unless the tag
  message has that line or an `- Invariant 4:` deviation. Run with
  `--excel-evidence` and `--workbook` before the tag is pushed, it also checks
  the evidence against the line. The **Release** workflow checks only the
  line's form: the evidence and `host.txt` stay on the machine of the run,
  so CI cannot check the run itself (#498).
- `tools/check_vba_conditionals.py` accepts a module-level `#Const`, such as
  `#Const WRITE_STATUS_BAR = False`, instead of rejecting every `#Const`. Its
  value is `True`, `False`, `-1`, `0` or an expression with `Not`, `And`,
  `Or`, `=`, `<>` and parentheses over `VBA6`, `VBA7`, `Win32`, `Win64` and
  the module's earlier `#Const` names. It is defined from its line on in each
  compilation environment that reaches it, and a condition is not evaluated
  where its branch cannot be entered. The VBA checks that read each
  environment's source, such as `vba-identifiers`, `vba-jumps` and
  `vba-public-api`, read the branch the `#Const` selects. Any other symbol,
  such as a project's conditional-compilation argument or a `#Const` used
  before its definition, stays a finding, as does a malformed, non-Boolean or
  duplicate `#Const` or one that redefines a compiler constant. Static check
  only; nothing was compiled or run in Excel (#500).
- `tools/check_links.py` reads Markdown the way GitHub renders it. Text in
  code spans that continue over line breaks, in indented code blocks and in
  HTML comments is no longer read as a link, nor is escaped text such as
  `\[text](path)`. Backslash escapes and character references, such as
  `&amp;`, are decoded in destinations, `href` and `src` values and
  `<a id>` anchors. Fences, headings, setext underlines and thematic breaks
  are recognised inside blockquotes and list items, a closing fence may have
  nothing after its run, and a heading's anchor drops the `_` runs of its
  emphasis, while the anchor that keeps them is still accepted. A wrong
  report can be listed with its reason in the optional
  `.github/link-check-exclusions.json`; **Repository integrity** fails while
  the file is malformed or an entry matches no report, and a plain `http://`
  link cannot be listed. With `--external`, a link to a page of this
  repository's wiki is requested as the page's raw Markdown while the
  repository answers anonymous requests, and skipped otherwise (#501).
- `tools/format_vba.py`, run without paths, reads the modules that Git tracks
  under the source directories; an untracked file there is no longer
  formatted or checked, and a run outside a Git repository is an error with
  exit status 2 (#502).
- The **Sync issue labels** dry run names, for each deletion a real run would
  refuse, up to ten of the issues and pull requests that still use the label,
  in number order, and adds `and others` when more do, such as
  `used by #3, #7 and others`; the refusal of a real run names them too. A
  plain name still costs one request, now a page of eleven; a name with a
  comma collects the numbers from every page it reads. An item without a
  valid number stops the run. A real run refuses the same deletions as
  before (#534).
- The guides and configuration comments describe the repository they are
  in, so a repository created from the template no longer rewrites them by
  hand. The `VERSION` row of the structure guide says that the file is
  present once the first release is prepared; the Dependabot header says
  that `target-branch`, when set, names the active release branch; and the
  guides say that a register or list is empty until the project adds its
  own. Steps
  only the template needs moved: running the release checks with
  `--template true` to the template-only section of `AGENTS.md`, and the
  new-repository steps of the label and traffic guides to the setup guide,
  which already held them. `tools/initialize_repository.py` also replaces
  the house style's request to replace `<REGIME>` with a paragraph that
  names the prefix, and lists that edit. The README's *Start a new
  repository* links the setup guide's *Template cleanup* for the list of
  edits instead of repeating it (#491).
- Fewer workflow runs repeat what **Repository integrity** already checks.
  **Sync issue labels** and **Detect issue-label drift** no longer run on
  pull requests, whose label self-tests, catalogue validation and drift
  fixtures **Repository integrity** runs. **CodeQL** analyzes pushes to the
  default branch only, and skips a pull request that changes only Markdown
  files and files under `docs/`. The minimum-Python test job moved to its
  own workflow, `minimum-python.yml`, which runs on pull requests that
  change Python code, the tests, the tools, the label scripts or the Python
  requirements, on pushes to the default branch and on manual dispatch; the
  `codeql-default-branch` check also requires it to run on the default
  branch. The label-drift and pull-request milestone scans run weekly
  instead of daily. The README's workflow table and the setup guide state
  when each workflow runs (#567).
- The tests and the repository checks run faster, with the same tests and
  the same results. `tools/check_repository.py` runs the command checks in
  parallel, up to four at a time, and still prints, logs and reports every
  check in the listed order; `clean-working-tree` still runs last and alone,
  and a command check that cannot start now fails instead of stopping the
  run. `tools/check_excel_run.py` copies the test cases of a commit with one
  `git cat-file` instead of one `git show` per file. The tests build the
  Excel-run fixture repository when they run, not on import, so that
  importing them leaves no temporary directory; copy one empty `.git` into
  each VBA fixture instead of running `git init` for each; run mypy once over
  its probes; stop their HTTP test server without the half-second poll; and
  keep the milestone workflow's output out of the test log. On a four-CPU
  machine, the test suite went from about 77 s to about 60 s and
  `tools/check_repository.py` from about 98 s to about 64 s (#568).
- The tools share one source for their common building blocks, with the
  same results. The new `tools/gitio.py` runs Git for every tool, raises one
  `GitError` that carries the command and Git's message, and lists the
  tracked files; each tool keeps its own errors, messages and exit codes.
  The documentation checks that scan the folder when Git cannot list the
  files do so through one documented function. The VBA component suffixes,
  the VBA string-literal pattern and the data checks' `SNAKE`, `is_number`
  and `is_date` are each defined once. The duplicate-key YAML loader moved
  from `tests/strict_yaml.py` to `tools/strict_yaml.py`, which
  `tools/check_repository.py` imports directly. The tests create their Git
  repositories through one helper, `tests/gitfixture.py`, which ignores the
  user's and the system's Git configuration (#569).
- Each repeated rule is written down once, with the same rules. The issue
  triage rules live in a registry in `tools/check_issues.py`, which refuses a
  finding with an unregistered code; `tools/render_docs.py` renders them as
  the *Triage rules* table of `docs/repo_governance/ISSUE_LABELS.md`, and the
  other guides, the pull request template and the triage workflow link it.
  The README's workflow triggers are rendered from each workflow's `on:`
  block. The setup guide's handover table keeps only the rows that need an
  action. Recording an Excel run by hand, a run for another commit and
  carried evidence moved from `INSTALLATION.md` to the new
  `docs/repo_governance/EXCEL_RUN.md`, and `RELEASING.md` step 4 keeps only
  the release gate. The tool details of *Negative controls* and *Local checks*
  in `CONTRIBUTING.md` and the link checker's exclusions and exceptions
  moved to the tools' docstrings. The release checklist now links the
  closeout checklist (#570).
- Every tool in `tools/` follows one exit status: 0 passed, 1 problems found,
  2 could not run. `check_excel_run.py` (an invalid or unknown commit),
  `check_accuracy.py --record-holdout` (a refusal), `make_mutant.py` (a
  refusal) and `release.py evidence` (an unknown commit) exited 1 in those
  cases and now exit 2, so a failed run says whether to fix the code or the
  command. `CONTRIBUTING.md` states the convention (#582).
- Dependabot has one `pip` block for `/.github` and `/tools`, with one group,
  so a week with updates in both directories opens a single pull request
  instead of two (#581).

### Fixed

- `tools/initialize_repository.py` writes no file when it only lists the
  changes, not even the bytecode of the tools it imports, as the README says
  (#464).
- `tools/format_vba.py` no longer splits a comment that continues with ` _`
  after an `End` statement, which turned its continuation into code, and no
  longer moves continuation text that looks like a label or an `End`
  statement. A folder named as a path is formatted, and a missing path is an
  error with exit status 2 instead of a traceback (#496).
- `tools/format_vba.py` no longer moves a statement keyword followed by a
  colon, such as `Else:` or `Loop:`, to column one as if it were a label. It
  keeps a blank line that a continuation carries over, after an `End`
  statement and at the end of the file too, so that no blank line is added
  to or removed from a logical line (#502).
- **Issue triage** reports a milestone without a due date together with a
  missing or invalid target end date, in the same run, instead of only after
  the date is fixed (#530).
- The plan digest of `tools/setup_repository.py` binds the rules of each
  live ruleset that covers a configured one, so an approval goes stale and
  apply refuses it when a covering ruleset changes while it still covers.
  A plan in which no ruleset is covered keeps its earlier digest; preview
  again before applying a plan in which one is (#532).
- `tools/setup_repository.py --apply` refuses before any write when a
  managed Actions setting that applies to the destination cannot be read,
  naming each one as unknown, instead of making every write and then asking
  you to set a value it never saw. `--check` is unchanged (#533).

## [1.0.0] - 2026-11-11

Changes made before this changelog existed are described by the closed
milestones and their pull requests.

### Added

- `tools/check_doc_commands.py` checks the commands that the tracked
  Markdown cites, in inline code and in shell or unmarked fenced blocks:
  each `python tools/...` script and each `python -m` module of the
  repository must be tracked, and each subcommand and option must exist in
  the script's `argparse` parser, which it reads from the source without
  importing or running it. It joins continued lines, skips placeholders
  such as `<folder>` and `X.Y.Z`, and reports a parser it cannot read
  instead of passing it. **Repository integrity** runs it, so a renamed
  option or a removed tool fails until the guides follow; only the
  commands of a tool marked **Remove once run**, once it is deleted, are
  skipped. The test that compares the setup guide's *Remove once run* rows
  with the tool is skipped once the guide is deleted, so a repository that
  retired the setup tooling still passes (#273, #298).
- **Repository integrity** has a `secret-files` check: it fails on a tracked
  file named as a secret or key file (`.env` and `.env.*` except
  `.env.example`; `.pem`, `.key`, `.p12`, `.pfx`, `.pvk`, `.jks`,
  `.keystore`; `id_rsa*` and `id_ed25519*` except their `.pub` keys;
  `.kdbx`), also one force-added past `.gitignore`, and says to remove the
  file, rotate the secret and follow *Handling a finding* in
  `DISTRIBUTION.md`. `.gitignore` now ignores every one of these names
  (#276).
- `tools/release.py verify-published vX.Y.Z --repository OWNER/REPOSITORY
  [--milestone N] [--asset FILE=SHA256]` checks a published release through
  `gh`, read-only, and prints one line per fact: the tag is annotated and
  points at the commit its `Candidate` line names, **Repository integrity**
  succeeded on it, the release is published and not a draft, its assets are
  exactly the source bundle, `SHA256SUMS` and each `--asset`, with the
  SHA-256 values that `SHA256SUMS` and `--asset` record, `gh attestation
  verify` passes on the bundle, and the milestone is closed with no open
  issue. It fails on any of these, except the attestation: when the release
  notes state that the bundle is not attested, it prints a `NOTE` instead.
  `RELEASING.md` step 7, the milestone closeout and the release
  certification now use it, each command with the tag and `--repository`;
  the download, the unblock and the Excel checks of a published workbook
  stay manual (#274, #317, #320).

- `tools/initialize_repository.py` adapts a repository created from the
  template in one diff: it deletes the template-only section of `AGENTS.md`,
  points the contact links and the traffic badges at the new repository, or
  removes the badges, sets or removes Dependabot's `target-branch`, replaces
  the changelog's entries and link references with an empty `[Unreleased]`
  section and restores its "No version has been released" sentence, deletes
  `VERSION`, sets the release rows of the *Current status* of `RELEASING.md`
  to their state before a first release, and replaces `<REGIME>` in the
  house style. Without `--apply` it only lists the changes; with it, it
  refuses a working tree that differs from the commit. It changes nothing on
  GitHub, and a second run changes nothing. A test initializes a copy of the
  template and runs **Repository integrity** on it with `--template false`.
  README step 2 and the setup guide name the command, with the manual edits
  as the fallback: they list every edit it makes and say which of them
  **Repository integrity** checks. The initializer and its test can be
  deleted once run (#272, #302, #318, #320).
- `tools/check_vba_calls.py` checks VBA calls that resolve to a procedure of
  the project, in statements and in expressions, qualified or not: too many
  or too few arguments, counting `Optional`, `ParamArray` and named
  arguments, a named argument the procedure does not declare, and a
  qualified call to a `Private` procedure of another module. It reports
  nothing for a `With`-block member, an object member, an array element, a
  `Property` procedure or a built-in. **Repository integrity** runs it as
  advisory: a finding is a warning, not a failure, until the house style's
  rule makes it blocking (#249).
- The host record of an Excel run, `host.txt`, names the operator, the start
  and finish times with their UTC offset, the Excel build and bitness, the
  `AutomationSecurity` and the Trust access to the VBA project as found, and
  an outcome (`PASS`, `FAIL`, `IMPORT_FAILED`, `COMPILE_FAILED`,
  `EXECUTION_TIMEOUT`, `INCOMPLETE`, `NOT_ENABLED` or `UNAVAILABLE`), beside
  Windows, the references and the compile mode. `excel_run.py` writes it also
  when the run stops early, and reads the macro settings without changing
  them. `check_excel_run.py` rejects a record that lacks a field, accepts
  `UNKNOWN` only in a record that did not pass and never in `operator`,
  `started`, `finished` or `outcome`, and accepts an outcome other than
  `PASS` only as evidence of that outcome; a record written by hand in the
  earlier three-line form now fails (#244, #284, #300).
- `tools/check_independent.py`, which **Repository integrity** runs: it runs
  each independent implementation registered for an `IMPL-` source in
  `tests/independent/adapters.json`, a file a project creates when it adds
  one, on the test cases, through an adapter that
  translates each case into the implementation's input, and compares the
  results with the expected values within their tolerance. It lists the cases
  an implementation cannot represent, with the reason, and skips an adapter
  whose implementation is not installed, with a note. The template registers
  none, so the check passes and says that nothing was cross-checked (#248,
  #318).
- The VBA harness has patterns for three more kinds of test.
  `tests/modules/TEST_Relations.bas` holds illustrative relation tests
  (additivity, symmetry, monotonicity, zero input) on a stand-in total, each
  with a check that a deliberately corrupted output fails it; it is not
  certified, and `tests/test_vba_harness.py` keeps its counts and checks
  consistent. The harness guide gives the skeleton of an end-to-end case
  runner, which restores the input area it writes, derived from the workbook
  contract's settings anchor, and the conventions of a volume benchmark,
  which `check_excel_run.py` refuses as an evidence suite. The case runner is
  a case procedure that writes every assertion, `inputs-restored` included,
  itself, with shared steps in helpers that return what failed, because the
  check reads assertions only in a case procedure. Static checks only: the
  new module and the skeleton have not been run in Excel (#246, #286, #301,
  #319).
- The methodology guides trace each rule. The methodology overview has a
  support matrix that marks each feature as supported (in scope, with each
  rule's status in the traceability table), not supported or entered by the
  user. The validation plan has a traceability table that
  maps each `SRC-` paragraph to its `CALC-` step, its VBA procedure, its test
  cases and a status: Validated, Exercised, Not validated or Not implemented.
  Each table has one illustrative row. `check_test_cases.py` fails a row whose
  source, procedure under `src/` or case does not exist, whose step is not
  one that its calculation contract's *Rules* table lists, whose case tests
  another calculation, or whose status contradicts what the row names. The
  validation matrix uses the same statuses (#247, #286, #301, #319).
- `RELEASING.md` step 8 checks the published workbook, with a *Published
  workbook* section in the release certification: the attached workbook's
  build stamp names the candidate and `production`, a rebuild of the tag
  gives the same SHA-256 or the same parts, and a downloaded, unblocked copy
  shows no link to another workbook and no `#REF!` name, and survives run,
  save and reopen. An optional user manual for workbook recipients, by
  convention one folder per released version under `docs/user-manual/`, is
  built to a PDF with a tool the release records (Pandoc as the example; the
  template installs none) and listed among the certified artifacts (#250).
- With `src/workbook/Template.xlsx`, `tools/devbuild/build_workbook.py` builds
  the workbook from a copy of the template, keeping its formats, print areas
  and protection: it sets the code names, writes the stamp and adds the VBA
  project, and two builds stay byte-identical. Without a template it writes
  blank sheets as before. A protection matrix, by convention
  `docs/methodology/workbook-protection.json` and added with the template,
  states each sheet's protection options, an option left out taking the
  standard's default, and its input (unlocked) and output (locked) ranges;
  `tools/check_workbook_protection.py` fails **Repository integrity** when
  the template differs from it, and passes as inactive while neither exists.
  The guides' protection example follows the workbook contract's
  illustrative settings area (#243, #286, #301).
- A machine-readable data contract, `docs/methodology/data-contract.json`,
  holds each dataset's file or sheet, encoding and delimited-text format, and
  each field's type, unit or currency field, required flag and allowed values
  or range. The *Datasets* table and the field tables of `DATA_CONTRACT.md`
  are rendered from it, and **Repository integrity** fails while they differ.
  `tools/check_data_contract.py`, which it runs, checks the contract and, with
  `--dataset`, validates input files against it, reporting each finding with a
  stable code such as `DC-TYPE` or `DC-MISSING-COLUMN`.
  `tools/check_test_cases.py` now runs that validator on every variant file of
  a file-based case outside an `invalid` case. The guide's *File format*
  points at the reference validator for the complete rules, numbers and
  booleans included; an empty *Encoding* means UTF-8 and applies to a file
  only, and a file of an `invalid` case need not decode. The illustrative
  dataset `DS-01` lists the three columns its example file already holds
  (#238, #301, #319).
- The evidence of an Excel run can certify project suites beside the example
  suite: the evidence policy, now version 2, lists them under `suites`, each
  with its certification entry point, counts and two logs, and a suite may
  name its own runner (`runner`, `RunSuite` by default), so that the relation
  tests can be certified with `"runner": "RunRelationTests"`. `excel_run.py`
  runs every entry point, `check_excel_run.py` checks each suite's normal run
  and certification from the same host, and every suite is read as the
  policy's Office bitness compiles it, so `#If Win64` cases count only there.
  The harness guide states the suite shape the check reads, and the
  installation guide's rejection table covers the project suites (#237,
  #300).
- **Repository integrity** fails a committed Office package, outside the
  regulatory source copies, that carries macros (a `vbaProject` part, a
  reference to one or a macro-enabled content type) or personal metadata (an
  author, last editor, company or manager in `docProps/`) (#241).
- Dependabot version updates wait 14 days after a release (`cooldown`), and a
  workflow control test requires it. The pull request guide describes how
  security updates, the manual actionlint update and the files GitHub reads
  only from the default branch behave with a release branch (#239).
- `tools/devbuild/excel_run.py` records an Excel run with one command on
  Windows. It builds the workbook, runs the certification in Excel through
  COM, closes without saving and writes `host.txt`, then checks the
  evidence. It needs no Trust access: with `--workbook` and no export,
  `tools/check_excel_run.py` proves identity by rebuilding the commit and
  comparing the workbook's parts byte for byte, and accepts `compile=ON_DEMAND`,
  because only the modules the run calls are compiled. It changes no macro
  setting: where Excel opens the workbook with macros disabled, it waits for
  **Enable Content**. It records the build's references in `host.txt`, and
  the release certification accepts its evidence for the workbook built from
  the candidate (`Binding: VERIFIED` by rebuild, with the workbook's
  SHA-256) and for the references (#204, #285, #287, #300).
- Accuracy contracts (`tests/accuracy/contracts.json`, checked by
  `tools/check_accuracy.py` in **Repository integrity**). A contract states
  the worst error a calculation guarantees for one regime, over separate
  calibration and holdout cases. Freezing it needs enough `published` or
  `independent` holdout cases, and a frozen threshold grows only with a new
  decision. With `--results`, the check measures the results of an Excel run.
  The template's illustrative contract stays a draft (#162). An Excel run now
  writes those results: a suite records them with `TEST_Harness.RecordResult`,
  and its certification saves them with the commit of the workbook's build
  stamp, the example suite as `results.json` and a project suite in the
  `results` file its entry in `.github/excel-evidence-policy.json` names. The
  example suite records the illustrative cases' totals from `tests/fixtures/`,
  so its entry points take that folder. `tools/check_excel_run.py` requires
  every results file when the commit holds contracts, measures them together,
  rejects a result that two files hold, and fails the run when the contracts
  do not hold. A workbook imported by hand has no build stamp, so its run
  fails while contracts exist (#184, #284).
- A **Release** workflow: on a pushed `vX.Y.Z` tag it verifies the tag against
  the changelog, `VERSION`, the default branch, the tag message and
  **Repository integrity**. It builds a reproducible source bundle with its
  checksums, attests their build provenance, and drafts the GitHub Release.
  A private repository attests only on GitHub Enterprise Cloud with
  `RELEASE_ATTESTATIONS` set to `enabled`. Without the variable, the draft
  says that the bundle is not attested; with it but without Enterprise Cloud,
  the attestation fails and no draft is created. A manual run rehearses a
  ref without attesting or releasing. `tools/release.py` runs the steps
  locally: `verify` reports each failed check, a missing tag included, as a
  finding, and `notes` and `bundle` report a failure on a `FAIL` line
  (#161, #299).
- `tools/render_docs.py`, which renders documentation facts the code already
  knows into marked blocks: the list of **Repository integrity** checks in
  `CONTRIBUTING.md` and the example suite's expected results in
  `INSTALLATION.md`. **Repository integrity** fails while a block is stale
  (#163).
- An OpenSSF Scorecard workflow, which scores the default branch's
  supply-chain posture weekly and on each push to it. It publishes to code
  scanning and the Scorecard service on public repositories only; the weekly
  run reads the visibility from the API, since a scheduled event carries no
  repository in its payload (#164, #299).
- `tools/devbuild/build_workbook.py`, which builds a macro-enabled workbook
  from the source of one commit, with the sheets, code names and defined
  names of the workbook contract, a build stamp, and a VBA project written
  from source, so it needs no *Trust access to the VBA project object model*.
  A rebuild is byte-identical, and a working tree with changes is stamped as
  not evidence-grade. `tools/devbuild/export_manifest.py` and
  `check_excel_run.py --workbook` bind an export to the workbook it came from
  (`VERIFIED`, `FAIL` or `UNVERIFIED`), and a release requires `VERIFIED`. An
  optional LibreOffice smoke test runs the harness certification as an early
  warning, never as evidence; it reads the example suite's entry point and
  the names of its logs and results file from the evidence policy. It exits
  2 for an unknown option, a missing workbook or a missing or malformed
  policy before it starts LibreOffice, and 1 without a traceback when
  LibreOffice cannot run the workbook. A form stops the build: import it by
  hand (#100, #285, #299, #316, #366).
- A VBA test harness and an example suite in `tests/modules/`, with a
  certification entry point. It refuses state left by an earlier run,
  reports completeness apart from the verdict, and checks that Excel's state,
  the open workbooks and the active window are unchanged. It writes every log
  before it raises, and checks that each deliberate failure reaches its
  caller. The example suite has 14 cases and 34 assertions; in
  `cleanup-primary-only`, not yet run in Excel, only the operation fails, its
  error reaches the caller and both settings are restored. A refused start
  writes two records, with `operation=NOT_RUN` and `cleanup=NOT_RUN`.
  **End** or **Reset** in the VBE clears the run state, so that stop is not
  detected; the guides say what to check by hand afterwards, and the
  installation guide's recovery item covers only a run that kept the
  project's state. `tools/check_excel_run.py` checks the evidence of an Excel
  run against one commit, without Excel, and rejects an incomplete run. The
  [harness guide](docs/methodology/TEST_HARNESS.md) and
  [recording an Excel run](INSTALLATION.md#recording-an-excel-run) describe
  them, and the check is a release gate (#79, #285, #319).
- Seven static checks of exported VBA source, which **Repository integrity**
  runs, the call check as advisory:
  - placement, names, encoding, options, layout and storage;
  - `#If` structure and `PtrSafe` in each compilation environment;
  - procedure-local jump targets;
  - visibility and the public API inventory, including a public name in two
    modules;
  - the dependency direction;
  - identifiers, in each compilation environment: a name a procedure assigns
    that nothing declares, a statement that calls a procedure no module
    provides, or a declaration named like a VBA keyword, such as
    `Dim Empty()`. A statement it cannot read is a finding (#173);
  - the calls to the project's own procedures, as the
    `tools/check_vba_calls.py` entry above describes (#249).

  The [house style](docs/VBA_HOUSE_STYLE.md#static-checks) says what the call
  and identifier checks each resolve and lists what the checks cannot prove
  (#70, #285, #300).
- `tools/check_repository.py`, one command that runs every **Repository
  integrity** check locally and in CI. It writes a log per check and a
  `report.json` with the commit, the dirty state and the SHA-256 of every
  tracked file. A summary of each check appears on the run page, and CI keeps
  the reports as a 30-day artifact. White space is now checked in every
  committed file, not only in the pushed changes. Without PyYAML it skips
  `codeql-default-branch`, and `--strict` fails it; the unit tests that read
  YAML are skipped too (#73, #287).
- `tools/format_vba.py`, which brings exported VBA source into the
  house-style layout with `--write`, and which **Repository integrity** runs
  with `--check`. It changes only line endings, trailing white space,
  indentation tabs, the level of procedure and label lines and the blank
  lines between procedures, never a string literal (#130).
- A release process: [RELEASING.md](RELEASING.md), this changelog, a release
  certification record and a milestone closeout record under `docs/release/`,
  and `tools/check_changelog.py`, which **Repository integrity** runs. A
  release's certification, distribution review and milestone closeout are
  issues of the next milestone. Milestone N closes in step 7, before the
  GitHub Release is published, once the closeout checklist is complete for
  the candidate that step 2 freezes. A fix for a release gate that failed on
  the candidate records its change in the dated section being released, as
  the changelog policy, the contributing guide and the pull request template
  say (#69, #302, #317).
- A numerical test-case format, `docs/methodology/TEST_CASES.md`, and
  `tools/check_test_cases.py`, which **Repository integrity** runs (#75).
  A case is a JSON fixture, or a file-based case whose variants hold input
  files named and encoded as in the data contract, stored byte for byte
  (#117). A case is `nominal`, `boundary` or `invalid`; failure and
  repeatability are left to VBA tests. A fixture record's `id` carries the
  dataset's identifier field, its required text field ending in `_id`. An
  output's `rule` is a step that its calculation contract's *Rules* table
  lists. Values that the project computes by hand, in a spreadsheet or with
  its own script are `illustrative` cross-checks, and a rule is validated
  only by `published` or `independent` values, as the validation plan also
  says. An illustrative case of each form is under `tests/fixtures/` and
  `tests/expected/`; its expected values cite rule `CALC-01.1` (#286, #301,
  #319).
- `tools/check_binaries.py`, which **Repository integrity** runs: a tracked
  binary file outside the approved paths fails unless
  `.github/binary-allowlist.json` lists it with its current SHA-256 (#113).
- `tools/check_derivations.py`, which **Repository integrity** runs: it runs
  the derivation scripts named by expected values, which are Python
  scripts, and compares their results within tolerance, once per variant for
  a file-based case, with an illustrative script for each example case
  (#115, #117, #301).

- A host behaviour register in `docs/methodology/VALIDATION_PLAN.md`: each
  Excel or VBE behaviour that the code relies on is `assumed` or `measured`,
  with the probe, the host and the result, and the guides and tools cite the
  entries they rely on. It holds `HB-01` to `HB-07`; the assumed ones include
  Excel compiling a project written from source when the workbook opens
  (`HB-05`), **End** or **Reset** clearing the harness's state (`HB-06`,
  whose probe pauses the run at a breakpoint) and the Immediate window
  keeping 200 lines (`HB-07`). Assumed entries are listed as a limit in the
  release certification (#133, #285, #300, #319).
- `docs/methodology/WORKBOOK_CONTRACT.md`, a workbook contract to fill in: the
  sheets and code names, defined names, areas, operations, status after a
  failure, and what the build creates. Code names are `sh` plus the tab name
  in PascalCase, as in the house style. The illustrative example declares
  `Inputs_Settings` as a one-cell table anchor, which
  `tools/devbuild/workbook.json` builds on `Inputs!B4`, typed only in its
  value column `C5`, with its size in the contract; `Run_Status` and the
  status lines are declared with the first sheet that code writes, and the
  reporting-date label stays in the example's output. The status table uses
  the area kinds the contract defines, and the contract lists `sheet` among
  the protection options that default to `false`. It states one failure
  behaviour: no new result, the result areas cleared and marked `FAILED`
  (#118, #286, #301, #319).
- The label catalogue `.github/labels.json` (schema 3) has a `core` list and
  optional `overlays`: named label sets that `selected_overlays` enables, for
  example one per regulatory regime in a generated repository. Synchronization,
  drift detection and the triage check resolve the core plus the selected
  overlays, refuse an unknown or duplicated selection or a name repeated across
  the core and an overlay, and show the core, overlay, selection and resolved
  counts; the template defines no overlay, so its live labels are unchanged.
  `labels-sync.mjs --mode plan --plan-output PATH` writes the planned changes
  as JSON, and every drift report, live or on a fixture, is cross-checked
  against that plan on the same snapshot: the run fails and the evidence
  carries both plans under `cross_check` if they differ (#354, #355).
- An eighth static check of exported VBA source, `tools/check_vba_errors.py`,
  which **Repository integrity** runs as `vba-errors`. In each compilation
  environment, it reports a read of `Err.Number`, `Err.Source` or
  `Err.Description` that follows, on straight-line code, an `On Error`
  statement or `Err.Clear`, also written `Call Err.Clear`, or, in an error
  handler, a call or member access, so that a handler cannot report or raise
  again an error that another statement has cleared or replaced. An Enum
  member or constant, also qualified as in `RunStatus.rsFailed` or
  `Module.LIMIT`, and a field of a variable of a user-defined `Type`, as in
  `Saved.Number = Err.Number`, are values, not member accesses. The
  [house style](docs/VBA_HOUSE_STYLE.md#static-checks) lists what it cannot
  see (#396, #419, #422, #427, #429).
- **Repository integrity** runs a ninth static check of exported VBA,
  `vba-declares` (`tools/check_vba_declares.py`), on the new register of
  native declarations, `docs/native-declarations.json`, which is empty in the
  template. In each compilation environment it reports, with file, line and
  environment, a `Declare` whose export and library the register does not
  list, a `Declare` of `GetLastError`, a `Declare Sub` registered as
  ambiguous on zero, a read of `Err.LastDllError` other than
  `<local> = Err.LastDllError` directly after a statement that is one native
  call, such as a read after a call inside a single-line `If` or a condition,
  or `.LastDllError` in a `With Err` block, and a call to an export registered
  as ambiguous on zero that is not directly preceded by `SetLastError 0`,
  whose whole result is not assigned to a local variable, or that is not
  directly followed by that capture. A repository with a `Declare` must
  register its export. An entry may also record the export's ABI, its
  `arguments` (each type, or `pointer`, and `byval`) and `returns` (a type,
  `pointer` or `null`); the check then reports, in each environment, a
  `Declare` whose parameter count, parameter types, `ByVal` or `ByRef`,
  `Sub` or `Function`, or return type differs from it, with a `pointer`
  expected as `LongPtr` under VBA7 and `Long` under VBA6 (#399). The house
  style, the structure guide and the setup guide document the rules and the
  register; nothing was run in Excel (#398).
- `tools/check_vba_jumps.py` reports an `Err.Raise` or `Error` statement
  reached after an unconditional `Resume <label>` while the procedure's own
  `On Error GoTo <label>` handler is enabled, in each compilation environment:
  the handler would trap the raise again, and a cleanup that raises the saved
  error can loop and hang Excel. The check also requires the line number in
  `If x Then 100`, an implicit `GoTo`, to be a label of the procedure. The
  [house style](docs/VBA_HOUSE_STYLE.md) now says to disable the handler with
  `On Error GoTo 0` before raising the saved error again. The check is static;
  nothing was run in Excel (#397).
- A pull request that fixes a defect now shows its new or changed test
  failing on the defective code, then passing on the fix, and records how in
  the *Validated* section of the pull request template
  ([negative controls](CONTRIBUTING.md#-negative-controls)).
  `python tools/make_mutant.py FILE --find TEXT --replace TEXT` writes the
  defective variant of a file reproducibly: a copy with the one occurrence of
  the text replaced, under the ignored `test-results/mutants/`, with every
  other byte, CRLF line endings and `Attribute VB_Name` kept, and the SHA-256
  of the source and the mutant printed. It refuses a text that does not occur
  exactly once, overlapping occurrences included, an output folder that Git
  does not ignore, and a symbolic link at the mutant's path (#400).
- The daily **Detect settings drift** workflow compares the live rulesets with
  `docs/setup/repository-settings.json` using only the job's read-only token,
  and fails on drift. It runs the new
  `tools/setup_repository.py --check --token-scope workflow`, which compares
  what that token can read: each ruleset without its bypass list, the
  visibility, the general settings GitHub returns to it and, on a public
  repository, private vulnerability reporting. Each value it cannot read is
  listed as not checked, never as drift or a match; with no drift the result
  is `PARTIAL` and the run passes. `--summary PATH` appends a Markdown summary
  in every mode. The [settings reference](docs/repo_governance/SETTINGS.md#drift)
  lists what the workflow checks and why (#403).
- `tools/check_excel_run.py --equivalent-to SHA` lets one Excel run certify a
  later commit that changes none of the run's inputs, such as a fix to a
  workflow or a guide. The evidence and the binding are still checked against
  `--commit`, the commit the workbook was built from. Then every run input
  must have the same Git blob at both commits: the VBA components and form
  resources, the files of the build configuration's directories, the
  methodology guides, fixtures, expected results and accuracy contracts, the
  build configuration, the layout template, the evidence policy, every
  `.gitattributes` file at any depth and the build tools. A difference fails and names each
  path; a match prints
  `Equivalent: <SHA> (N run inputs identical to <commit>)`, and the report
  names both commits. `INSTALLATION.md`,
  `RELEASING.md` step 4, the release certification record and the pull
  request guide and template say when such a run counts and that its record
  names both commits (#402).
- The `actionlint` check of **Repository integrity** first lints
  `.github/actionlint/invalid-workflow.yml`, a workflow with a known error
  that never runs, and fails unless actionlint rejects it with that finding.
  Only then does it lint the workflows, so a broken or stubbed actionlint that
  accepts everything can no longer pass the check (#404).
- `tools/devbuild/build_workbook.py --override DIR` builds a workbook in which
  each `.bas` or `.cls` file of `DIR` and its subfolders replaces the module
  with the same `Attribute VB_Name`, for the negative control of a VBA fix. A
  file that names no module or the same module as another, a file without one
  `VB_Name` or outside the source rules, and a `.frm` stop the build. The build
  prints each replaced module with its SHA-256, and the stamp's commit records
  the folder and those SHA-256 values as *NOT evidence-grade*.
  `tools/check_excel_run.py` reports a workbook built this way, or a results
  file it wrote, as not evidence for the commit, never as
  `Binding: VERIFIED`. [INSTALLATION.md](INSTALLATION.md#building-a-workbook-from-source)
  describes the option and its limits (#406).
- Every release starts with a
  [maintenance audit](docs/release/MAINTENANCE_AUDIT.md) of the repository's
  issues, pull requests, review discussions and notifications. It checks that
  closed issues meet their acceptance criteria, that open issues and `blocked`
  labels are current, that review threads are resolved only on evidence, and
  that no text claims a change, test or review its evidence does not support.
  Its report goes on the milestone closeout issue before the candidate is
  frozen; a release invariant, the release checklist and the closeout
  checklist require it (#438).
- `docs/assets/social-preview.png`, exported from its editable source
  `social-preview.svg`, is the repository's social preview and the banner at
  the top of the README. It is allowlisted in `.github/binary-allowlist.json`
  with its SHA-256. The setup guide tells a generated repository to replace
  or remove both, and to set its own **About** description, website, topics
  and social preview, which GitHub does not copy from a template (#434).

### Changed

- `tools/setup_repository.py` manages all twelve general repository settings
  of `docs/setup/repository-settings.json`, now schema 4. It writes the merge
  methods before the rulesets and the other settings after them, so merged
  branches are deleted automatically only once the release branches are
  protected. The configuration may hold any number of uniquely named
  rulesets; loading refuses one that allows a merge method the settings
  disable, a required check without a context and an integration ID, or a
  repeated rule type. An optional `visibility` makes a different visibility
  drift, which apply refuses and never changes. Preview and `--check` need no
  administrator access: a security feature or Actions setting that GitHub
  does not report is unknown, with its endpoint and status, and so is a
  ruleset whose bypass list GitHub omits for a token that cannot write it;
  an unknown value counts as drift, and apply refuses while a managed
  setting, security feature or ruleset is unknown. Each difference is printed as a `DIFF` line, a ruleset's by the
  path of the parameter that differs. `--evidence PATH` writes the check as
  JSON with the configuration's SHA-256, a `{date}` in the path becoming the
  UTC date. Five GitHub Actions settings are checked and never written: an
  apply whose read-back finds one different names it, to set in
  **Settings → Actions → General**. The plan ends with its SHA-256 over the
  repository, the configuration, the observed managed values and the exact
  writes, recorded in the evidence; `--apply` requires `--approve-plan` with
  that digest, refusing a plan that changed since the review, and `--journal`
  with a new path, written atomically before the first write and after each,
  with each write's status and the read-back result. A configured ruleset
  with no ruleset of its name that active rulesets without bypass actors or
  excluded refs enforce at least as strictly, under any names and as the
  plan leaves them, is `covered`, as is a named one with stronger
  parameters: not drift, and never written. Missing rulesets are created
  before differing ones are replaced (#332, #333, #334, #335, #352, #353).
- The label guide states the dependency-section syntax that **Issue
  triage** reads: an issue link or URL, an `External condition:` line in any
  letter case after an optional list bullet, and a pull request reference,
  which never counts as a prerequisite. A test reads its example as triage
  does (#329).
- *Keep milestones issue-only* reads each pull request again after removing
  its milestone and fails if GitHub still reports one, then writes a job
  summary with its scope and each pull request and milestone it removed, or
  says that nothing was removed.
  When GitHub gives the run of one pull request a read-only token, it warns
  and passes, because the next daily or manual scan removes the milestone; a
  scan still fails on any error (#328, #345).
- The **Traffic alert** job reads the `traffic-history` commit that its
  collection pushed, or merged onto when nothing changed, which the collect
  job outputs as `history-sha`, not the branch's current tip; it stops when
  that output is missing, and a workflow control test keeps the binding
  (#304).
- External links must use HTTPS: `tools/check_links.py` fails a plain
  `http://` link in a Markdown link, a reference definition, an autolink or an
  HTML attribute, outside code, in **Repository integrity**, and the weekly
  **Link check** requests links over HTTPS only and fails a redirect to
  another protocol. A 403, 429 or 5xx answer stays a warning, but one listed
  in the optional `.github/link-exceptions.json`, with a reason and an
  `expires` date, fails once that date has passed; **Repository integrity**
  fails while the file is malformed or lists a URL that no Markdown file
  links (#275).
- **Repository integrity** cancels the run of a pull request commit that a
  newer commit replaces; a push run is never cancelled. CI runs
  `tools/check_repository.py` with `--ci`, which fails when, after the
  checks, the working tree differs from the commit. mypy checks `tests/` as
  well as `tools/`; three unannotated test files need no annotations, but
  their code is checked (#366). The bug form has an optional field for evidence, a
  workaround and recovery, and the enhancement form one for alternatives
  considered (#251).
- `tools/check_excel_run.py` compares the modules exported by hand with the
  commit without the letter case of code outside strings and comments, as the
  VBE re-cases a name (`HB-01`, still `assumed`), and reports each re-cased
  identifier on a `Note:` line. In a document module it no longer compares the
  attributes Excel writes itself, such as `VB_Base`; `VB_Name`, string
  literals, comments and any other difference still fail (#245).
- `SECURITY.md` describes the runtime risk surface: what the test harness,
  the example suite, `tools/devbuild/excel_run.py` and
  `tools/devbuild/lo_smoke.py` read, write and delete,
  the Excel settings they need (enabled macros, no Trust access, no change to
  a macro setting), which access the dependency rules forbid in core and
  classes, and what a generated repository states for its own entry points.
  It adds a data-and-secrets section (what a workbook can hide, and that a
  leaked secret is revoked or rotated before it is reported) and a safe
  harbor for good-faith research. The setup guide tells a generated
  repository to adapt the runtime section (#242, #287, #300, #316).
- The pull request template asks for the calculation impact of a change: the
  source of each changed rule and the class of the expected values that test
  it, or "No calculation change". Its checklist asks for the `[Unreleased]`
  changelog entry. `CONTRIBUTING.md` states the rules of evidence (expected
  values never come from the code under test; a skipped test is not a pass),
  what a commit message says, and what a contract change is and what it
  updates (#240).
- mypy also reports unreachable code (`warn_unreachable`), such as a branch
  that type narrowing has already excluded; strict mode alone does not (#191).
- Label synchronization no longer deletes a label that an issue or pull request,
  open or closed, still carries: with `prune: true`, it checks every label it
  would delete and stops before any change; a dry run (`--dry-run`) reads
  the same usage without any write, and its summary names each deletion a
  real run would refuse. A name with a comma, which the API's label filter
  cannot query, is checked against the labels of every issue and pull
  request, up to 10,000; a malformed response or the limit stops it. The
  reconcile job has `pull-requests: read`, without which the issues API
  would omit pull requests from that check (#188, #305, #316, #336, #360).
- `tools/check_changelog.py` freezes released sections: once the tag
  `vX.Y.Z` exists, the section must equal its content at the tag. A section
  with a reviewed line `Corrected after release on YYYY-MM-DD: <reason>` may
  differ, and the difference is reported as a note. Deleting or renaming a
  tagged section fails. A release without a tag yet, or a copy without Git
  history, is a note (#172).
- The workflow control tests check every `uses:` reference a workflow can
  reach, not only those of job steps. That covers job-level calls of reusable
  workflows, local actions followed into their metadata, and Docker images,
  which need an `@sha256:` digest. Each remote pin also needs a version comment
  (#174).
- What an Excel run must leave as evidence moves from constants in
  `tools/check_excel_run.py` to `.github/excel-evidence-policy.json`: the
  harness, the suite with its counts and certification entry point, the modes
  and their logs, the evidence files, the references and the Office bitness.
  The check reads the policy of the candidate commit, rejects a missing or
  malformed one, and fails when the suite disagrees with it (#160).
- `tools/check_changelog.py` also fails when the prose describes another
  release state than the changelog: the "No version has been released"
  sentence, the *Current status* of `RELEASING.md`, placeholder lines such as
  `TBD`, and a "Release Candidate" mention or unchecked task in the README
  (#131). A status row must name the newest release as a whole version
  number, so `10.2.0` does not pass for `0.2.0` (#142).
- The pull request template has an *Excel run* block for changes that need
  Excel: the commit, host, compile result (`compile=ON_DEMAND` included),
  suite, evidence, and a completion state that is `PASS`, `FAIL`,
  `INCOMPLETE` or `NOT RUN`, with the host record's `outcome` for a run that
  stopped early.
  [PULL_REQUESTS.md](docs/repo_governance/PULL_REQUESTS.md#excel-run) defines
  the states (#134, #285).
- `INSTALLATION.md` explains how to upgrade to a new revision by rebuilding
  the workbook and carrying its inputs across, how to diagnose common import
  and run failures, and what removing the workbook's components leaves
  behind. Its import steps include `examples/modules/` (#132, #285).
- `docs/licenses/provenance.json` records the source repository, license and
  revision of each adapted file, or why the revision is not recorded; a test
  keeps it in step with the third-party notice (#116).
- **Repository integrity** proves that actionlint rejects an invalid workflow,
  and the workflow control tests reject duplicate YAML keys in the workflows,
  Dependabot's configuration and the issue forms (#114).
- The [public API inventory](docs/PUBLIC_API.txt) has two tiers.
  `[supported]` lists the supported API in `src/modules/`, as before;
  `[project-public]` lists every other `Public` declaration under `src/`,
  including the public members of classes, forms and workbook modules.
  `tools/check_vba_public_api.py` matches each tier with the source one to
  one, with normalized signatures, and reports a declaration listed in the
  wrong tier; test and example modules belong to neither tier. An inventory
  without tier headings is read as `[supported]` alone, so it still passes
  only while `src/` has no public declaration outside `src/modules/`. In a
  generated repository, add the `[supported]` and `[project-public]` headings
  and list each public declaration of core, classes, workbook modules and
  forms under `[project-public]`, as the
  [repository structure](docs/REPOSITORY_STRUCTURE.md#public-api-inventory)
  guide describes (#405).
- The template repository is `danielep71/K-TEMPLATE`. The README title, its
  badges and wiki links, the issue forms' contact links and the changelog's
  link references use the new name; GitHub redirects links that use the old
  one (#433).
- `tools/format_vba.py` and `tests/test_format_vba.py` are licensed under the
  Mozilla Public License 2.0 of the root `LICENSE`, like the repository's own
  files. The third-party notice and the provenance register no longer list
  them (#440).

### Fixed

- The guides describe the repository's files. The setup guide's handover
  table covers every tracked file, names both traffic tests and
  `tests/test_excel_run.py`, and describes each test of
  `tests/test_repository_consistency.py`; the structure guide covers the
  root files, `VERSION`, all of `.github/`, `docs/release/`,
  `docs/licenses/` and the harness guide. The label guide matches the
  workflows' write-permission list. `.gitattributes`, `.editorconfig` and the
  guides name the source register in `SOURCES.md`. The calculation
  contracts' *Outputs* template names each output's quantity, their common
  rules hold the default tolerances, and a contract lists its `_status`
  quantities and their texts under *Failure behaviour*. The rules for filling
  in the guides name the guides with a *Decisions* table and the two tables
  that differ (#286, #287, #301, #302, #318, #319, #366, #371).
- `tools/setup_repository.py` no longer reports paused Dependabot security
  updates as on: they count as drift, the setup record shows `on (paused)`,
  and an apply that leaves them paused fails its read-back (#190).
- A backtick, pipe or line break in a label name, colour or API error no
  longer breaks the code span or table row of the label synchronization and
  drift summaries (#189).
- A malformed expected file no longer stops the derivation, accuracy and
  Excel-run checks with a traceback: they leave its malformed outputs out,
  and the test-case check reports them. A malformed evidence policy, a suite
  entry point the suite source lacks, or a suite source that does not exist
  is reported as a finding by the generated-documentation check instead of a
  traceback. Without Node.js, a local run of **Repository integrity** skips
  `issue-configuration`, as it skips the label checks, instead of failing it
  (#367, #371).
- The tests that a generated repository keeps no longer depend on the
  template's illustrative content: they still pass once the setup guide's
  steps replace the illustrative cases and the contract `ACC-01`, change the
  example suite's counts, delete or replace the relation tests, rename the
  illustrative sheets or list a project suite in the evidence policy, once
  `docs/PUBLIC_API.txt` lists a declaration, including one declared in both
  branches of an `#If`, and once `tests/accuracy/contracts.json` is removed
  with the outputs that the example suite's results case records. The tests
  of a tool use synthetic cases, contracts, modules and build configuration
  of their own, and the tests of the repository derive their expected values
  from its files. No test comment cites a template issue (#368, #371, #373).
- The [repository structure](docs/REPOSITORY_STRUCTURE.md) guide now claims
  only what the checks enforce: the static checks enforce the dependency
  direction for core, classes and examples, and the facade's and adapters'
  direction is checked in review; only the standard modules in `src/modules/`
  can hold supported API, so a class is never listed in the inventory; and
  only an Office Open XML package is inspected for macros and personal
  metadata, which the [distribution guide](docs/repo_governance/DISTRIBUTION.md)
  now also says (#382).
- The test-case checks now enforce what `TEST_CASES.md` and the validation
  plan state. An invalid file-based case scopes a result only to a dataset of
  its manifest. An invalid JSON fixture may break or omit its calculation
  inputs, but a key outside the fixture keys fails. A code, ID, quantity or
  rule must be the whole value with ASCII digits, so `"EUR\n"` and a rule
  step in another script's digits fail, in `check_test_cases.py`,
  `check_accuracy.py` and `check_independent.py`. A *Traceability* row may
  name a procedure of a form under `src/`. `check_independent.py` fails on
  any expected value citing an adapter's source that the adapter does not
  print, not only an `independent` one. The setup guide's `DS-01` row now
  says to remove both illustrative cases and their *Traceability* row, and
  to replace `ACC-01`, in the same change (#384).
- `tools/release.py verify` and the **Release** workflow accept a tag
  message whose `Deviations:` line is followed by a `- ` or `* ` list on the
  next lines, as well as by text such as `none` on the same line, and reject
  `Deviations:` with nothing after it, with an error that names both forms.
  `RELEASING.md` shows both forms, defines a released version as one whose
  section is dated and a published one as one with a tag and a GitHub
  Release, and says that `RELEASE_ATTESTATIONS` without GitHub Enterprise
  Cloud fails the attestation and creates no draft. The changelog policy,
  `SECURITY.md` and the setup and structure guides match (#385).
- **Repository integrity** no longer reports a Dependabot `target-branch` as
  missing in a single-branch or shallow clone: `dependabot-target-branches`
  asks the `origin` remote for a branch that has no remote-tracking branch,
  and is skipped, or fails under `--strict`, when `origin` cannot answer.
  `tools/initialize_repository.py` asks the same way before it warns that
  `--release-branch` is missing, and says so when `origin` cannot answer. The
  check table in `CONTRIBUTING.md` shows the `changelog` check with
  `--repository OWNER/REPOSITORY`, as CI runs it (#381, #394).
- **Repository integrity** accepts the bytes that `.gitattributes`,
  `.editorconfig` and the formatter keep. `git diff --check` reports no
  white-space error in the regulatory source copies under
  `docs/methodology/sources/`, nor trailing blanks in a `.frm` file, and
  `tools/check_vba_source.py` skips a form's VBE header, which keeps the
  blank the VBE writes at the end of its `Begin` line; the code after it is
  still checked. Editors trim trailing white space in Markdown, which
  `git diff --check` reports, and leave registry exports (`*.reg`) byte for
  byte (#380).
- `tools/devbuild/lo_smoke.py` reports a failed build of the development
  workbook, such as one stopped by a form, as `lo-smoke: FAIL (build: ...)`
  and exits 1, without a traceback. The Excel-run guides, `SECURITY.md` and
  the tools' docstrings now match the build and evidence tools: Excel reads
  the built project from source and compiles each module when it is first
  used (`HB-05`); `tools/check_excel_run.py` also reads the workbook and
  rebuilds the commit in a temporary Git worktree; `SECURITY.md` lists that
  worktree, the second build and the `Visible` and `DisplayAlerts` settings
  of the Excel instance `tools/devbuild/excel_run.py` starts; the evidence
  policy and the harness guide name every tool that reads the policy, and a
  changed reference also changes `REFERENCES` in
  `tools/devbuild/vba_project.py` (#383).
- Label, settings and traffic messages and guides now match what the tools
  do (#386). Label synchronization reports a label as outside the catalogue,
  the core plus the selected overlays, and advises adding it to the catalogue
  or selecting its overlay; **Issue triage** names the same catalogue. An
  apply with no steps reads the repository back, so its journal records a
  fresh observation. The *Traffic alert* job decides from the history before
  it looks up a milestone, so a run with no alert due needs none, and it
  compares due dates with the collection date instead of the job's date;
  `docs/repo_governance/TRAFFIC.md` lists the milestone among the enabling
  steps. `docs/setup/REPOSITORY_SETUP.md` says that only an existing managed
  ruleset is unknown without write access.
- The traffic alert tests in `tests/test_workflow_controls.py` skip when
  `.github/workflows/traffic-history.yml` does not exist, so the traffic
  removal that `docs/setup/REPOSITORY_SETUP.md` describes leaves them passing,
  and the guide says so. `SECURITY.md` lists the temporary folder in which
  `tools/check_excel_run.py` writes the commit's test cases to measure the
  accuracy contracts. The Dependabot `target-branch` test uses a longer name
  that `git ls-remote` returns for the branch, so it fails if the check stops
  comparing exact names (#423).
- The Excel instance that `tools/devbuild/excel_run.py` starts no longer
  outlives the run. The command records the process id of that instance,
  from its window, and the process start time. After it quits Excel, it waits
  up to 30 seconds for that process to exit and then ends it, but only while
  the process with that id is named `EXCEL` and has the recorded start time,
  so it never ends another Excel. It prints how Excel ended in an
  `excel-quit:` line (`EXITED`, `STOPPED`, `SURVIVED` or `UNKNOWN`), which
  the saved report also holds; `SECURITY.md` and the installation guide say
  so (#401).
- The tests accept a non-empty register of native declarations: a
  repository that registers the exports its `Declare` statements import, as
  the structure and setup guides say, no longer fails the `tests` check
  (#445).

### Validation

- First Excel run of the template, with `tools/devbuild/excel_run.py` at
  50c0658 on Microsoft 365 Excel 16.0 build 20430, 64-bit, Windows 11
  Enterprise 25H2: `excel-run: PASS`, binding `VERIFIED` by rebuild, ACC-01
  worst error 0.0, `compile=ON_DEMAND`. The host behaviours `HB-02` to `HB-04`
  of the validation plan are measured on that host (#204, #211).

[Unreleased]: https://github.com/danielep71/K-TEMPLATE/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/danielep71/K-TEMPLATE/releases/tag/v1.0.0
