<div align="center">

# 🚀 Repository Setup

**From template files to a configured GitHub repository**

[![Scope](https://img.shields.io/badge/Scope-12_settings_%C2%B7_5_security_features_%C2%B7_5_rulesets-00338D?style=flat-square)](#%EF%B8%8F-what-this-applies)
[![Mode](https://img.shields.io/badge/Default-Preview_only-005EB8?style=flat-square)](#%EF%B8%8F-launch-sequence)
[![Python](https://img.shields.io/badge/Python-3.11%2B-3776AB?style=flat-square&logo=python&logoColor=white)](#-before-you-start)
[![Verification](https://img.shields.io/badge/Verification-Read_back_after_apply-00875A?style=flat-square)](#-verification-and-recovery)

[Prepare](#-before-you-start) · [Launch](#%EF%B8%8F-launch-sequence) · [Labels](#%EF%B8%8F-initialize-issue-labels) · [Coverage](#%EF%B8%8F-what-this-applies) · [Clean up](#-after-launch-keep-adapt-or-remove) · [Maintain](#-maintain-the-baseline)

</div>

---

> [!IMPORTANT]
> GitHub's **Use this template** copies files, not repository settings. Run this
> setup for each new destination after its first integrity workflow succeeds.
> No workflow applies the baseline: **Detect settings drift** only checks it,
> with a read-only token, and fails until you apply it. The script never
> creates or stores credentials.

## 🧭 Before you start

| Prerequisite | Ready when… |
| --- | --- |
| **Runtime** | Python **3.11+** and a recent [GitHub CLI](https://cli.github.com/) are installed. |
| **Destination** | The new repository exists and contains the template files. **Use this template** copies the template's default branch as it is then, not a release tag: record the template commit it came from. |
| **Template cleanup** | The template-only sections of `AGENTS.md` and `README.md`, between their `template-only` markers, are deleted, or replaced with the text each holds for a new repository; the two `contact_links` URLs in `.github/ISSUE_TEMPLATE/config.yml` and the README traffic badges name the new repository, or the badges are removed; `target-branch` in `.github/dependabot.yml` names a branch of the new repository, or is removed; and every link reference of `CHANGELOG.md` points to the new repository. `Repository integrity` fails in a generated repository until all five are done. Replace the changelog's entries with an empty `[Unreleased]` section, and restore the introduction's "No version has been released" sentence if it is missing; delete `VERSION` if the template has one, since it exists once a template release is prepared; reset the *Current status* of `RELEASING.md` to no release; replace `<REGIME>` in the naming table of `docs/VBA_HOUSE_STYLE.md`, and the paragraph above the table, which asks for that, with one that names the prefix; and remove the template's evidence, which cites its own commits and issues: the rows of the *Tested environment* table of `INSTALLATION.md`, and the *Issue* column of the *Host behaviour* table of `docs/methodology/VALIDATION_PLAN.md`, which becomes a dash. Once the changelog holds no dated release, `Repository integrity` fails while `VERSION` exists, the sentence is missing or *Current status* does not say that no version is released and that `VERSION` does not exist; it does not check the changelog's entries or `<REGIME>`. In a clean clone, `python tools/initialize_repository.py --repository OWNER/NEW-REPO --regime PREFIX` lists all these edits and makes them with `--apply`; see *Recommended handover* for its options. Without it, make them by hand. |
| **Identity** | `gh auth status --hostname github.com` shows an account that can read the destination; `--apply` needs one with administrator access to it. Sign in with `gh auth login --hostname github.com`, or set `GH_TOKEN` to a fine-grained token. |
| **Token permissions** | Preview and `--check` only read. A fine-grained token needs destination repository **Metadata: read**, and **Administration: read** for Dependabot alerts, Dependabot security updates and the five Actions settings. GitHub reports secret scanning and push protection only to administrators of the repository, and a ruleset's bypass list only to a token that can write the ruleset; without it, each managed ruleset that exists is unknown, and a missing one is reported as missing. A value the token cannot read is shown as unknown, with the endpoint and HTTP status, and counts as drift. `--apply` needs administrator access and a token with **Administration: write**, **Contents: read** and **Checks: read**. A workflow `GITHUB_TOKEN` cannot administer these rulesets; `--check --token-scope workflow` compares what it can read ([scheduled drift check](../repo_governance/SETTINGS.md#drift)). Never commit tokens. |
| **Integrity check** | Actions is enabled and `Repository integrity` has succeeded on the **current default-branch HEAD**. |
| **Workflow coverage** | The required job runs on every relevant PR without path filters. The supplied workflow runs on every pull request, on pushes to the default branch and `release/**`, and on manual dispatch. Its push filter names the default branch: after renaming it, update the filter, or the `codeql-default-branch` check fails. |
| **Availability** | The destination's plan and visibility support repository rulesets. |

## ▶️ Launch sequence

Run from the copied repository directory. Replace `OWNER/NEW-REPO` with the
**new destination**, not the source template. Each command is on one line, so
it can be pasted into PowerShell or `sh` alike.

Run it on a task branch created from the default branch, such as
`docs/settings-evidence`:

```sh
git switch -c docs/settings-evidence
```

Step ② writes its journal, and step ③ its report and evidence, into
`docs/repo_governance/`, as files to commit. Step ② also creates `Protect main`,
after which the default branch accepts changes only through a pull request.
Commit the three files on the task branch and merge them through a pull request.

The baseline is what `docs/setup/repository-settings.json` holds once the
initializer has run: `--remove-traffic` takes a ruleset out of it, which step ②
then never creates. A ruleset the configuration does not hold, such as one
created before the initializer ran, is listed as unmanaged and stays active: no
step changes or deletes it, so delete it in **Settings → Rules → Rulesets** if
nothing needs it.

### ① Preview the changes

```sh
python tools/setup_repository.py --repo OWNER/NEW-REPO
```

Read-only by default, and it needs no administrator access. Review the printed
destination, visibility and default branch, then one `DIFF` line for each managed
value that differs from the configuration, and the proposed API changes that
follow them, before applying them:

```text
DIFF setting has_projects: observed true, configured false
DIFF security vulnerability_alerts: observed false, configured true [GET vulnerability-alerts (HTTP 404: off)]
DIFF ruleset "Protect main" (12345) rules.pull_request.allowed_merge_methods: observed ["merge", "squash"], configured ["merge"]
```

A security feature or Actions setting that GitHub does not report to your token is
`null`, with the endpoint and status in brackets; it counts as drift.

A configured ruleset that no ruleset of its name exists for, but that other
active rulesets already enforce at least as strictly, under any names, is
**covered**: the preview prints a `NOTE` line that names them instead of a `DIFF`,
it is not drift, and nothing is planned or written for it. A ruleset counts
towards coverage only when it has the same target, a reported and empty bypass
list, no excluded refs and every configured ref in its include (`~ALL`, or
`~DEFAULT_BRANCH` for the default branch); each configured rule must be present in
one of them with parameters at least as strict: more required approvals, fewer
allowed merge methods, a superset of the required checks. Different rules may be
covered by different rulesets. Coverage is judged on the rulesets as the plan
leaves them: a named ruleset that the plan replaces counts with its configured
definition, not with what its live definition enforces beyond it, and missing
rulesets are created before differing ones are replaced. A named ruleset with
stronger parameters covers itself and is `covered`, not `differs`; a weaker one
is `differs` and is replaced whatever other rulesets enforce. Otherwise the named
ruleset is `matches`, `missing` or `unknown`, as the `DIFF` lines show.

The plan ends with its digest, `PLAN SHA-256: <hex>`: the SHA-256 of the
repository name, the configuration's SHA-256, every observed managed value, the
definition of each live ruleset that covers a configured one, and the exact
planned writes. A covering ruleset that changes while it still covers, for
example by gaining a rule, changes the digest; a ruleset that covers none does
not. The same unchanged repository and configuration give the
same digest; `--check` prints it too, and `--report` and `--evidence` record it.
Apply needs it.

### ② Apply the baseline

```sh
python tools/setup_repository.py --repo OWNER/NEW-REPO --apply --approve-plan PLAN_SHA256 --journal docs/repo_governance/settings-journal-YYYY-MM-DD.json
```

Replace `PLAN_SHA256` with the digest that the preview printed. Apply reads the
repository again, rebuilds the plan and computes its digest before any write; when
it differs from the approved one, because the repository or the configuration
changed since the reviewed plan, apply refuses and asks you to preview again. An
approved plan with no steps is a valid apply: the journal is written, nothing is
written to GitHub, and the read-back runs as usual.

`--journal PATH` names a new file; an existing one is refused before any write.
The journal is JSON with the plan digest, the repository, the configuration's
SHA-256, the start time, each planned write with its method, endpoint and status
(`pending`, `done` or `failed`, with the error text), and the read-back result.
A ruleset write also names its ruleset, and once done records the ID that GitHub
returned for it. The journal is written atomically before the first write and
again after each write's result, so it exists before anything reaches GitHub. A
failing write stops the sequence and is recorded; nothing is retried or rolled
back. Keep the journal with the setup record, for example beside
`SETTINGS_EVIDENCE.md` as the path above does.

The script first sets the three merge methods where they differ, because GitHub
may reject a ruleset that allows a disabled one. It then creates missing rulesets,
replaces differing ones when you allow it, sets the security features GitHub offers for the destination's
visibility, patches the other differing general settings, and reads back the
result. Protection is installed before automatic branch deletion is enabled, so a
merged pull request cannot delete a release branch before its deletion rule exists.

Apply refuses before any write when the plan's digest differs from the approved
one, the journal path exists, you lack administrator access, the destination
is archived or disabled, its visibility differs from a `visibility` the
configuration sets, or a managed general setting, security feature, ruleset or
Actions setting is unknown. An Actions setting that does not apply to the
destination, such as the fork pull request approval policy of a private
repository, is never unknown.
It never changes visibility: follow the
[distribution review](../repo_governance/DISTRIBUTION.md) first.

The script never writes the five Actions settings. When one still differs after
the rest is applied, the read-back fails and names each: set it in the
repository's **Settings → Actions → General**, then rerun `--check`.

### ③ Verify the result

```sh
python tools/setup_repository.py --repo OWNER/NEW-REPO --check --report docs/repo_governance/SETTINGS_EVIDENCE.md --evidence 'docs/repo_governance/settings-evidence-{date}.json'
```

A verified match returns **0**. `--report` writes the repository's **setup
record**: a dated Markdown file with the plan digest, each managed general setting, security feature
and Actions setting as configured and as GitHub reports it, each managed ruleset with its ID and status,
the rulesets that cover it included, and any unmanaged or inherited rulesets. `docs/repo_governance/SETTINGS_EVIDENCE.md`
is the conventional place to commit it in a generated repository; add anything the
script cannot read, such as decisions and exceptions, below it. `--report` works with
preview, `--check` and `--apply`; after apply it records the read-back state.
Running apply again against matching settings makes no writes.

`--evidence PATH` writes the same check as JSON, for a machine to read, whether or
not you also pass `--report`; a `{date}` in the path becomes the UTC date of the
check (`YYYY-MM-DD`). Keep such a path in single quotes, as above: PowerShell
does not pass the braces of an unquoted `{date}` through. The record names the
mode, repository, default branch and result, the configuration's path and
SHA-256, and each managed value as configured
and observed, with the source of each security feature and Actions setting, each
managed ruleset with its differences and the rulesets that cover it, the unmanaged
and inherited rulesets, the method and endpoint of each planned change, with the
ruleset's name for a ruleset change but without its payload, and the plan digest
(`plan_sha256`). It works in the same modes as `--report`. A public-only feature
or Actions setting on another visibility is recorded as
`not applied (public repositories only)` and does not count as drift.

`--summary PATH` appends a short Markdown summary of the comparison to a file,
such as `$GITHUB_STEP_SUMMARY`, in every mode.

The **Detect settings drift** workflow runs the check daily with its own
read-only token:

```sh
python tools/setup_repository.py --repo OWNER/NEW-REPO --check --token-scope workflow
```

`--token-scope workflow` compares only what a workflow `GITHUB_TOKEN` with
`contents: read` can read: each ruleset without its bypass list, the visibility,
the general settings GitHub returns to it and, on a public repository, private
vulnerability reporting. It never calls the endpoints that need **Administration:
read**. Each value it cannot read is printed as `NOT CHECKED` and is neither drift
nor a match: a value GitHub leaves out of its answer, or refuses with HTTP 401,
403 or 404. Any other failed request, such as a server error, is an API failure
that returns **2**. Drift in
what it reads returns **1**; otherwise it returns **0** with the result `PARTIAL`
while something is not checked. It plans nothing, prints no
digest and takes neither `--report` nor `--evidence`: the setup record comes only
from a check with an administrator's token. The
[settings reference](../repo_governance/SETTINGS.md#drift) lists what each section
needs.

Dependabot security updates count as on only when GitHub reports them enabled and
not paused. GitHub pauses them when Dependabot pull requests go unanswered. Paused
updates are drift: the plan prints a note, the record shows `on (paused)`, and an
apply that does not resume them fails its read-back.

<details>
<summary><strong>⚙️ Advanced: reconcile an existing named ruleset</strong></summary>

A named ruleset that differs and is not covered blocks apply before any writes.
After reviewing the preview, explicitly allow replacement of its entire definition:

```sh
python tools/setup_repository.py --repo OWNER/NEW-REPO --apply --update-existing --approve-plan PLAN_SHA256 --journal docs/repo_governance/settings-journal-YYYY-MM-DD.json
```

Replacement removes additional rules or bypasses inside that named ruleset; a
named ruleset that is `covered` is never replaced. Rulesets with other names and
inherited organization rules remain unchanged. Duplicate managed names are rejected.

When you adopt an existing repository, its rulesets may protect the same branches
and tags under other names. One that enforces a configured ruleset with no ruleset
of its name at least as strictly covers it, and the preview says so. The preview also names every unmanaged
ruleset with its target, ref patterns and enforcement. It stays active after
apply, with its own rules and bypasses, so delete the ones that a managed ruleset
replaces in **Settings → Rules → Rulesets**. The script never edits or deletes an
unmanaged ruleset.

The destination is always explicit; it is never inferred from the current directory.
The default JSON path is relative to the script, so it also works from another
working directory. Use `--config PATH` for a reviewed baseline with the same schema.

</details>

## 🏷️ Initialize issue labels

After creating a repository, run **Actions → Sync issue labels → Run workflow** on the default branch. This installs the 20 labels defined in this repository, including exact descriptions and colours. See [labels and automation](../repo_governance/ISSUE_LABELS.md) for triage rules and instructions for replacing my default assignee account in a generated repository; the initializer's `--assignee LOGIN` does it. Keep the label configuration, scripts, workflows, forms, and triage tests after setup; they are ongoing automation, not bootstrap leftovers.

## 🛡️ What this applies

The [JSON configuration](repository-settings.json) is the executable counterpart
of the [settings reference](../repo_governance/SETTINGS.md).

| Layer | Managed configuration |
| --- | --- |
| **General · 12 settings** | The three merge methods, auto-merge, branch update suggestions, deletion of merged head branches, web commit sign-off, and the Issues, Projects, Wiki, Discussions and legacy downloads features. Five differ from GitHub's defaults: allow auto-merge, delete merged head branches, suggest PR branch updates, disable Projects and the legacy downloads flag. |
| **Security · 5 features** | Dependabot alerts and security updates on every repository; private vulnerability reporting, secret scanning and push protection on a public repository (skipped with a note otherwise). |
| **Actions · 5 settings, checked only** | SHA pinning required, read-only default workflow permissions, no pull request approval by workflows, approval for every external contributor's fork pull request workflows (public repository only) and 90-day retention. Reported and verified, never written. |
| **Protection · 5 rulesets** | Default branch, release branches, independent release-branch deletion protection, release tags, and traffic history. The configuration may hold any number of uniquely named rulesets. Active rulesets without bypass actors or excluded refs that enforce a configured one at least as strictly, under any names, cover it: nothing is written for it. |
| **Visibility · optional** | `null` in the template: any visibility. Set to `public` or `private`, a different visibility is drift that apply refuses and never changes. |
| **Enforcement** | Active rulesets with no bypass actors or ref exclusions. `Protect main` targets `~DEFAULT_BRANCH`, not a fixed branch name. |
| **Verification** | All pages are read; full rule definitions are compared without response metadata; apply writes only the plan whose digest was approved, journaling each write; managed settings are read back after apply. |

**Outside this baseline:** template mode, the other Actions policies, Pages,
secrets, environments, labels, collaborators and integrations. Configure these
separately where required. Code scanning is configured by the CodeQL workflow,
not by the script; see *Recommended handover*.
The script creates no branches, tags, releases or PRs. Destination defaults and
inherited policies may impose additional restrictions.

## 🧹 After launch: keep, adapt or remove

> [!TIP]
> **Nothing must be deleted after setup.** Keeping the script and configuration
> lets you run `--check` later and reapply an approved baseline. Removing local
> setup files does not undo settings already saved on GitHub.

### Recommended handover

| File or setting | Action in the new repository | Reason / dependency |
| --- | --- | --- |
| `docs/assets/social-preview.svg`, `docs/assets/social-preview.png` | **Replace or remove** | The template's social preview and README banner, which name the template. Edit the SVG with the project's name and subtitle, export a 1280×640 PNG or larger at 2:1, under 1 MB, and update its SHA-256 in `.github/binary-allowlist.json`. Or delete both, the allowlist entry and the image at the top of `README.md`. |
| Repository **About** and **Social preview** | **Set** | GitHub copies neither from a template. Set the description, website and topics in the repository page's **About** settings, and upload the PNG in **Settings → General → Social preview**; GitHub offers no API for the preview. |
| `tools/initialize_repository.py`, `tests/test_initialize_repository.py` | **Removes itself once run** | Makes the *Template cleanup* edits of *Before you start*: it also replaces the changelog's entries and `<REGIME>`, with the house style's request to replace it, restores the changelog's "No version has been released" sentence if it is missing, deletes `VERSION` if it exists, resets the *Current status* of `RELEASING.md`, removes the template's evidence rows, and puts in place of the README's template-only sections, the *Start a new repository* section and the template's wiki links, the table of the main guides they hold, as one diff to review and commit. Without `--apply` it lists them and writes nothing. With `--apply` it refuses a working tree that differs from the commit, writes the edits, renders the generated blocks afresh and, last, once all that has succeeded, deletes itself and its test; the list without `--apply` names that deletion too. It never changes GitHub or commits. By default it repoints the traffic badges, removes Dependabot's `target-branch`, links the changelog's `[Unreleased]` to `main` and keeps the assignee; `--remove-traffic-badges`, `--release-branch BRANCH` (an existing branch), `--default-branch BRANCH` and `--assignee LOGIN` change that. `--assignee LOGIN` (a GitHub login) assigns the four issue forms and the updates of `.github/dependabot.yml` to that account, and names it in `docs/repo_governance/ISSUE_LABELS.md`. `--remove-traffic` removes the optional traffic collection, everything its row below lists, so it implies `--remove-traffic-badges`, which removes only the badges and keeps the collection; it takes the ruleset of the `traffic-history` branch out of `docs/setup/repository-settings.json` only, and its list says to delete one already created on GitHub by hand, since `tools/setup_repository.py` never deletes an unmanaged ruleset. A mention of the collection it cannot place is listed as a warning, to edit by hand. It does not change the label manifest, so the overlays you select stay as you set them. A template-only section may hold its replacement in a comment, `<!-- template-only:replacement` ... `-->`, just before its end marker; the initializer's docstring gives the format. Its test initializes copies of the template, runs **Repository integrity** on one with `--template false`, and is skipped once the template-only sections are gone. |
| `AGENTS.md` | **Adapt** | Delete the template-only section first; keep the remaining instructions that still apply to the project. |
| `CONTRIBUTING.md`, `CODE_OF_CONDUCT.md` | **Keep and adapt** | The contribution entry point and the Contributor Covenant 2.1 by reference, both shown in GitHub's community profile. Adapt the checks and reporting contact to the project. |
| `SECURITY.md` | **Keep and adapt** | Names the private reporting channel, response targets and safe harbor. Adjust supported versions and scope to the project. Rewrite *Runtime risk surface* for the project's macros as they are added: for each entry point of the production workbook, what it reads, writes, saves or deletes, and whether the facade, workbook or form code starts a process, creates an external object, calls a native library or reaches the network. |
| `docs/repo_governance/DISTRIBUTION.md` | **Keep and adapt** | The review to record before making the repository public, sharing a copy or publishing a release, with the full-history scan. Add the project's own search patterns. |
| `README.md` | **Adapt** | Replace the template's title, introduction and banner with the project's purpose, installation and usage. The initializer replaces the *Start a new repository* section and the template's wiki links, which are template-only sections, with a table of the main guides; keep it current as guides are added. |
| `LICENSE` | **Keep or replace** | MPL-2.0 by default. The integrity workflow requires a non-empty `LICENSE`. |
| Default branch other than `main` | **Adapt** | Rename `main` in the `push` branch lists of `.github/workflows/repository-integrity.yml`, `.github/workflows/minimum-python.yml` and `.github/workflows/codeql.yml`, and in the `pull_request` branch list of `codeql.yml`. **Repository integrity** fails while they omit the default branch. |
| Code scanning and release attestations on a private repository | **Set** | Leave GitHub's code scanning default setup off: it rejects the results of `.github/workflows/codeql.yml`. On a private repository, CodeQL needs GitHub Code Security and the repository variable `CODE_SCANNING` set to `enabled`, and release attestations need GitHub Enterprise Cloud and the variable `RELEASE_ATTESTATIONS` set to `enabled`; without them both are skipped ([settings reference](../repo_governance/SETTINGS.md)). |
| `docs/VBA_HOUSE_STYLE.md`, `docs/REPOSITORY_STRUCTURE.md`, `docs/PUBLIC_API.txt`, `INSTALLATION.md` | **Keep and adapt** | VBA naming, layout, contracts, errors and cleanup; component locations and dependency direction; the inventory of the supported and project-public VBA declarations; import, export, upgrade, troubleshooting and removal. Replace `<REGIME>` with the project's short name, and adapt the host policy. The initializer removes the template's tested environments: record the runs of the new repository. The static checks of VBA source read the facade prefix from the house style's naming table. If you import existing modules that are not in the house-style layout, list them in the [layout baseline](../VBA_HOUSE_STYLE.md#source-representation), which the formatter and the layout check both read, instead of reformatting them all at once. |
| `RELEASING.md`, `CHANGELOG.md`, `docs/release/`, `tools/check_changelog.py`, `tests/test_check_changelog.py` | **Keep and adapt** | The release process, the changelog, and the certification and closeout records. Replace the changelog's entries and its link references with the new repository's: **Repository integrity** fails while a link reference points to another repository. Keep or restore the introduction's "No version has been released" sentence, delete `VERSION` if the template has one (it exists once a template release is prepared), and reset *Current status* in `RELEASING.md` to no release; **Repository integrity** fails while these disagree with a changelog that holds no dated release. Create `VERSION` when you prepare the first release. |
| `docs/methodology/` | **Keep and adapt** | Sources, data contract, calculation contracts, validation plan, the numerical test-case format, the workbook contract and the VBA test harness guide. Replace each illustrative row with the project's entries as decisions are taken. The *Datasets* table and the field tables of `DATA_CONTRACT.md` are rendered from `docs/methodology/data-contract.json`: replace its illustrative dataset `DS-01` there with the project's datasets, each with its file or sheet, encoding, delimiter and decimal separator and its fields, then run `python tools/render_docs.py --write`. In the same change, remove both illustrative cases and their row of the *Traceability* table, and replace `ACC-01`, as the next row says. **Repository integrity** fails while a table differs from the file, and `python tools/check_data_contract.py --dataset DS-NN FILE` validates a real input file against it. The host behaviour entry `HB-01` concerns `tools/check_excel_run.py`, `HB-02` to `HB-04` concern `tools/devbuild/excel_run.py` and hold for the host they name, `HB-05` concerns `tools/devbuild/build_workbook.py` and the identity by rebuild of `tools/check_excel_run.py`, and `HB-06` and `HB-07` concern the harness and installation guides; keep them, delete the template's issue and commit references from their *Issue* and *Result* columns, which this repository's history cannot resolve, and record your own host when you measure it. |
| `tests/fixtures/`, `tests/expected/`, `tests/derivations/`, `tests/accuracy/`, `tools/check_test_cases.py`, `tools/check_data_contract.py`, `tools/check_derivations.py`, `tools/check_accuracy.py` and their tests | **Keep and adapt** | The numerical test cases, their format check, the data contract's reference validator, which the format check runs on every variant file of a file-based case, the comparison with the derivation scripts and the accuracy contracts, which **Repository integrity** runs. Remove the `illustrative-format-example` and `illustrative-file-import-example` cases, their scripts, `tests/derivations/illustrative_total.py` and `illustrative_file_total.py`, and the row of the *Traceability* table in `docs/methodology/VALIDATION_PLAN.md` that names them, when you replace the illustrative entries of the methodology guides they use, and replace the illustrative contract `ACC-01` with the project's. In the same change, point `CaseIllustrativeResults` in `tests/modules/TEST_Examples.bas`, which reads those fixtures, at the replacement cases, or replace it and update the suite's counts in `.github/excel-evidence-policy.json` and the guides; `tests/test_vba_harness.py` fails while it reads a fixture that is gone, or while the results it records differ from the outputs of the contracts' cases. |
| `tools/check_independent.py`, `tests/test_check_independent.py` | **Keep; register an adapter when wanted** | Runs each independent implementation registered in `tests/independent/adapters.json` on the test cases and compares its results with the expected values; with no adapter, as in the template, it passes and says that nothing was cross-checked. To register one: give the implementation an `IMPL-` entry in `docs/methodology/SOURCES.md`; write an adapter script under `tests/independent/` that translates a fixture into the implementation's input, runs it and prints its outputs, or why it cannot represent the case or is not installed; then create `tests/independent/adapters.json`, which the template does not ship, or add to it the adapter's `source` and `command` ([independent implementations](../methodology/TEST_CASES.md#independent-implementations)). An adapter whose implementation is not installed is skipped with a note: install the implementation in the **Repository integrity** workflow for the comparison to run in CI. |
| `docs/native-declarations.json`, `tools/check_vba_declares.py`, `tests/test_check_vba_declares.py` | **Keep and adapt** | The [register of native declarations](../REPOSITORY_STRUCTURE.md#native-declarations) and the static check that reads it. Add an entry for each export that a `Declare` imports, in the pull request that adds it, with its ABI in the optional `arguments` and `returns` keys; **Repository integrity** fails on a `Declare` the register does not list, or that differs from its export's ABI in a compilation environment. The register is empty in the template, and the check passes while there is no VBA source. |
| `tests/modules/TEST_Relations.bas` | **Keep and adapt** | The illustrative [relation tests](../methodology/TEST_HARNESS.md#relation-tests), each with its corruption check, on a stand-in total. Replace `IllustrativeTotal` with a call to the core procedure of your calculation and the relations with those its calculation contract implies, keeping a corruption check in every case; `tests/test_vba_harness.py` keeps the counts and the checks consistent, and skips them once the module is gone. Delete the module, and its row in the harness guide's *Modules* table, if no relation applies. |
| `tools/devbuild/`, `tests/test_devbuild.py`, `tests/test_excel_run.py` | **Keep and adapt** | Builds the workbook from source ([building a workbook from source](../../INSTALLATION.md#building-a-workbook-from-source)), binds an export to it, and records an Excel run with one command (`excel_run.py`, [recording an Excel run](../../INSTALLATION.md#recording-an-excel-run)). Replace the illustrative sheets and names of `tools/devbuild/workbook.json` together with the workbook contract; the tests keep them equal. |
| `tools/check_restated_counts.py`, `tests/test_check_restated_counts.py` | **Keep and adapt** | Fails, in **Repository integrity**, when a Markdown file restates by hand a count that has a single source; its docstring lists the sources. Declare the project's own counted records in `SOURCES`, and keep a reviewed restatement only in `ALLOWED`, with its reason. |
| `.github/PULL_REQUEST_TEMPLATE.md` | **Keep and adapt** | Records linked issues, validation, final-head review and what was not validated. Add the project's own checks to *Validated*. |
| `.github/workflows/traffic-history.yml`, `tools/traffic_history.py`, `tests/test_traffic_history.py`, `docs/repo_governance/TRAFFIC.md` | **Keep; enable if wanted** | Optional traffic collection. It does nothing until you follow *Enabling* in `TRAFFIC.md`. Point the README's traffic badges at the new repository, or remove them; **Repository integrity** fails while they read another repository. If you never want it, initialize with `--remove-traffic`. After initialization, remove by hand, in one change: the four files and the badges; the files' entries in `docs/licenses/THIRD_PARTY_NOTICES.txt` and `docs/licenses/provenance.json`; the `Protect traffic history` ruleset of `docs/setup/repository-settings.json`, and the live one in **Settings → Rules → Rulesets**; the `readme-traffic-badges` entry of `tools/verification_depth.json`, since `tools/check_repository.py` lists that check only while the workflow or a badge is there; the rows and sentences of the guides that describe it, and every link to `TRAFFIC.md`, which `python tools/check_links.py` lists; then run `python tools/render_docs.py --write`. The traffic tests of `tests/test_workflow_controls.py` skip once the workflow is gone. |
| `.github/workflows/scorecard.yml` | **Keep, or remove** | Scores the default branch's supply-chain posture with OpenSSF Scorecard weekly and on each push to it, and publishes the results to the Scorecard service and to code scanning. It is skipped on a private repository and is not a required check. If you never want the results published, delete it together with its `WRITE_PERMISSIONS` entry and its test in `tests/test_workflow_controls.py`. |
| `docs/repo_governance/README.md` | **Keep and adapt** | The index of the governance guides and the repository controls. Update its links when you add, retire or move a guide. |
| `docs/repo_governance/PULL_REQUESTS.md` | **Keep and adapt** | Final-head review procedure that the pull request template refers to. Adapt it if the project uses another automated reviewer. |
| GitHub **Template repository** setting | **Review separately** | Choose whether the new project should itself be reusable as a template; the setup script does not change this flag. |

Every other file is kept as it is; its guide or its docstring says what it
does.

### As the project grows

- **Checks.** Add each project check to `tools/check_repository.py`, which the
  **Repository integrity** workflow runs, with its entry in
  `tools/verification_depth.json`. A check without a negative test needs one
  first; update the entry when you rename or delete a proof test.
- **Generated facts.** Add a renderer to `tools/render_docs.py` for each fact
  the code already knows, and run `python tools/render_docs.py --write`.
- **Workflows.** Update `tests/test_workflow_controls.py` in the same change as
  the workflow it checks.
- **Suites.** Add the project's VBA suites beside the example suite, each with
  a certification entry point and an entry under `suites` in
  `.github/excel-evidence-policy.json`
  ([harness guide](../methodology/TEST_HARNESS.md#modes-and-certification)).
  Keep the example suite: its certification is the release evidence that the
  harness works on the host.
- **Binaries.** Add a reviewed binary, such as a documentation image, to
  `.github/binary-allowlist.json` with its SHA-256 and reason.
- **Cross-file checks.** Adapt a check of `tests/test_repository_consistency.py`
  only together with the files it compares.
- **Git and editor rules.** Change `.gitattributes`, `.gitignore` and
  `.editorconfig` only with a layout change that needs it.
- **Third-party files.** If you remove a file listed in
  `docs/licenses/THIRD_PARTY_NOTICES.txt`, remove it from the notice and from
  `docs/licenses/provenance.json` too. Record there the revision of any file
  you adapt, or why it is unknown.

### Optional: retire the setup tooling

Use this route only in a **generated project** where you no longer want the setup
utilities. Keep them in the source template so future repositories receive them.

1. Run `--check` successfully and preserve the result and adopted configuration in
   the setup record.
2. Create a normal task branch and PR for cleanup. Protection remains active after
   setup; do not rely on direct pushes to the default branch.
3. In that same PR, add meaningful project checks to the workflow. Keep the job
   name **`Repository integrity`**, its coverage of all relevant PRs, the steps
   that set up Node.js and install the configuration parser, the quality tools
   and actionlint, the **Run the repository checks** step and the step that
   keeps its reports. Add project checks to `tools/check_repository.py`, and
   keep every check it already runs, as the generated table of
   [CONTRIBUTING.md](../../CONTRIBUTING.md#-local-checks) lists them, the
   template-cleanup checks included. Keep `.github/workflows/codeql.yml`,
   `.github/workflows/minimum-python.yml` and
   `.github/workflows/scorecard.yml`: they are not part of the setup tooling.
   Every test in `tests/` other than `tests/test_setup_repository.py` stays
   after the setup tests are removed; without those setup and install steps,
   the issue triage, workflow control and quality-gate tests fail. The test
   that compares the *Removes itself once run* rows of this guide with
   `tools/check_doc_commands.py` is skipped once the guide is deleted.
4. Remove only the retired setup files, together:

   ```text
   tools/setup_repository.py
   tests/test_setup_repository.py
   docs/setup/repository-settings.json
   docs/setup/REPOSITORY_SETUP.md
   .github/workflows/settings-drift.yml
   ```

   In the same change, remove
   `test_settings_drift_checks_on_a_schedule_with_the_read_only_token` from
   `tests/test_workflow_controls.py`.

5. Update `README.md`, `docs/repo_governance/README.md`,
   `docs/repo_governance/SETTINGS.md` and `CONTRIBUTING.md`: remove links and
   commands that reference deleted files, and explain how settings will be
   maintained going forward. Preserve the actual settings reference. In
   `CONTRIBUTING.md`, replace the link to this guide with the Ruff and mypy
   commands of *Maintain the baseline* below. In `docs/release/CLOSEOUT.md`,
   replace criterion 5, which runs the setup script, with the way you now
   verify the settings. `python tools/check_links.py` lists any link left
   behind, and `python tools/check_doc_commands.py` any command of a deleted
   tool.
6. Verify that the cleanup PR reports a passing **`Repository integrity`** check
   with the intended project checks, then merge it through the protected PR process.

> [!WARNING]
> Do not delete the integrity workflow or rename its required job during cleanup.
> GitHub would continue waiting for that check and block merges. Deleting only the
> script or JSON would also break the existing setup tests. Remove individual files,
> not entire `tools/`, `tests/`, `docs/` or `.github/` directories.

Cleanup is documented as a reviewed PR because replacing template tests with useful
project validation depends on the new project. There is no automatic deletion step.

## 🔎 Verification and recovery

| Exit code | Meaning | Next action |
| --- | --- | --- |
| **0** | Successful preview, verified match in apply/check mode, or no drift in what `--token-scope workflow` read | Read the mode-specific message: preview success does not mean settings were applied, and `PARTIAL` means that the values listed as `NOT CHECKED` were not compared. |
| **1** | Drift found by `--check`, an unknown value included; with `--token-scope workflow`, drift in a value it read | Preview the differences and decide whether to reapply, or grant the access the unknown value needs. |
| **2** | Invalid configuration, a plan whose digest is not the approved one, an existing journal path, refused or failed apply prerequisite, unavailable feature, API failure or read-back mismatch, an Actions setting to set by hand included | Resolve the reported problem, then preview and retry. |

API operations are **not transactional**. If a later request fails, earlier changes
remain. There is no silent skip, no success claim for partial application and no
automatic rollback of protection. The journal records which writes were done and
which one failed; rerun preview to see what remains.

A passing check on the current HEAD is a prerequisite, not proof of future workflow
coverage or destructive enforcement testing. Verification covers the documented
settings, not every GitHub setting or inherited policy.

## 🔄 Maintain the baseline

Update the JSON and Markdown together when changing policy, review the preview,
and run the [local checks](../../CONTRIBUTING.md#-local-checks) before applying
it:

```sh
python tools/check_repository.py
```

Once the tools they need are installed, they require neither credentials nor
network access.

**Toolchain baseline.** The workflows set exact versions: Python 3.13.18, Python
3.11.18 in `minimum-python.yml` (the `python_version` in `pyproject.toml`), and
Node.js 24.22.2 without the package-manager cache. Dependabot does not update
these values: change them together, in every workflow, in one pull request. The
workflow control tests reject floating versions and mixed versions.

**Action pins.** Every `uses:` reference reachable from a workflow is
immutable:

- a remote action or reusable workflow, called from a step or from a job, is
  pinned to a full commit SHA;
- a Docker image (`docker://`) is pinned to an `@sha256:` digest;
- each pin carries its version as a comment on the same line
  (`# v7.0.1`), which Dependabot and reviewers read;
- a local reusable workflow must exist in `.github/workflows/`, and is checked
  as a workflow in its own right;
- a local action (`./path`) is followed into its `action.yml` or `action.yaml`,
  whose steps and image follow the same rules.

`tests/test_workflow_controls.py` fails on any other form, and on a reference
it cannot read with its comment, such as a flow mapping. A job that calls a
reusable workflow is exempt from the `timeout-minutes` rule, which GitHub does
not allow there.

**Repository integrity** also runs Ruff and strict mypy on `tools/` and
`tests/` (both configured in `pyproject.toml`, with `warn_unreachable` on top of
strict mode) and actionlint on the workflows, once actionlint has rejected
`.github/actionlint/invalid-workflow.yml`. A test file with a `# mypy:`
comment that turns off the annotation rules need not annotate its functions,
but mypy still checks their code. To run Ruff and mypy locally on
CPython 3.13 for Linux x86_64, the platform the hashes in
`tools/requirements-quality-ci.txt` cover:

```sh
python -m pip install --only-binary=:all: --require-hashes -r tools/requirements-quality-ci.txt
ruff check tools tests
mypy
```

**Repository integrity** checks every relative link and anchor, and that every
external link uses HTTPS; the weekly **Link check** workflow requests the
external links with `python tools/check_links.py --external`. The docstring of
[`tools/check_links.py`](../../tools/check_links.py) says how Markdown is read,
and how to list a wrongly reported link in `.github/link-check-exclusions.json`
or a busy external site in `.github/link-exceptions.json`, with an expiry date.

Once setup tooling has been retired from a generated project, use that project's
replacement validation and settings-maintenance process instead.

---

**References:** [Repository updates](https://docs.github.com/en/rest/repos/repos#update-a-repository)
· [Rulesets](https://docs.github.com/en/rest/repos/rules)
· [Actions permissions](https://docs.github.com/en/rest/actions/permissions)
· [GitHub CLI API](https://cli.github.com/manual/gh_api)

[↑ Back to top](#-repository-setup)
