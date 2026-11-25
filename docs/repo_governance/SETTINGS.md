# ⚙️ Repository settings · Non-default configuration

[![Scope](https://img.shields.io/badge/Scope-Non--default_settings-00338D?style=flat-square)](#scope)
[![Format](https://img.shields.io/badge/Format-Portable-005EB8?style=flat-square)](#reuse)
[![Comparison](https://img.shields.io/badge/Comparison-Documented_defaults-483698?style=flat-square)](#sources)

**Only deliberate differences · No repository identifiers · Reusable configuration**

<a id="scope"></a>

## 🧭 Comparison basis

This document records only settings that differ from the comparison baseline:
a new standalone GitHub.com repository, without inherited organization rules,
an applied security configuration or copied template customizations. General
options use GitHub's documented REST parameter defaults; those defaults are
not a guarantee of every account's repository-creation UI behavior.

Custom active rulesets are compared with an otherwise unprotected repository,
not with individual checkboxes preselected inside the ruleset editor. In
particular, deletion protection may be preselected when creating a ruleset;
creating and activating that targeted ruleset is still a customization.

Reference documentation checked: **2026-10-29**. Repository-specific evidence,
observation dates and rule IDs belong in the repository's setup record described
in the [setup guide](../setup/REPOSITORY_SETUP.md), not in this portable file.
Copying this file does not apply settings or verify another repository.

[General differences](#general) · [Security features](#security) · [Actions settings](#actions) · [Custom protection](#rulesets) · [Reuse](#reuse) · [Sources](#sources)

<a id="general"></a>

## 🏷️ General settings — changed values only

| Setting | Documented baseline | Configured value | Effect / qualification |
| --- | --- | --- | --- |
| Allow auto-merge | Off (`allow_auto_merge: false`) | **On** | Lets a pull request merge on its own once every required check and rule passes; it bypasses none of them. Using it stays optional. |
| Automatically delete merged head branches | Off (`delete_branch_on_merge: false`) | **On** | Cleans up eligible task branches; protected retained branches remain subject to their deletion rules. |
| Always suggest updating PR branches | Off (`allow_update_branch: false`) | **On** | Offers updates even without a freshness requirement; local governance still governs use on shared integration branches. |
| Repository Projects feature | On (`has_projects: true`) | **Off** | Disables the repository feature; this does not describe account-level project boards. Organization defaults may differ. |
| Legacy downloads flag | On (`has_downloads: true`) | **Off** | API metadata difference; not a prohibition on release assets or source downloads. |

Source: GitHub's [repository REST documentation][repo-api]. These values are
configuration differences, not blanket recommendations for every project.

<a id="security"></a>

## 🔐 Security features — always enabled

Their initial state depends on the account, the repository's visibility and
any applied security configuration, so they are not classified as differences
from a default. The setup script enables each one where GitHub offers it and
verifies it on every check.

| Feature | Configured value | Availability / effect |
| --- | --- | --- |
| Private vulnerability reporting | **On** | Public repositories only. Adds **Security → Report a vulnerability**, the private channel that `SECURITY.md` names. |
| Dependabot alerts | **On** | Every repository. Alerts on vulnerable dependencies. |
| Dependabot security updates | **On** | Every repository; requires Dependabot alerts. Opens pull requests that fix vulnerable dependencies, separately from the scheduled version updates in `.github/dependabot.yml`. |
| Secret scanning | **On** | Public repositories; private repositories need a paid licence and are skipped with a note. |
| Push protection | **On** | Same availability as secret scanning, which it requires. Blocks pushes that contain a detected secret. |

Sources: [private vulnerability reporting][pvr], [Dependabot alerts][dependabot-alerts],
[Dependabot security updates][dependabot-updates] and [secret scanning][secret-scanning].

<a id="actions"></a>

## ⚡ Actions settings — checked, never applied

Their initial state depends on the account and on any organization or
enterprise policy, so they are not classified as differences from a default.
The setup script reads each one and reports it as configured and observed, but
never writes it: set them in the repository's **Settings → Actions → General**.

| Setting | Configured value | Why |
| --- | --- | --- |
| Require actions to be pinned to a full-length commit SHA (`sha_pinning_required`) | **On** | GitHub refuses to run an action that is not pinned to a commit SHA. The workflows already pin every action, and the workflow control tests check it. |
| Workflow permissions (`default_workflow_permissions`) | **Read** | The `GITHUB_TOKEN` starts with read access only; each job declares the write permissions it needs. |
| Allow GitHub Actions to create and approve pull requests (`can_approve_pull_request_reviews`) | **Off** | A workflow cannot approve a pull request; no workflow here needs to create one. |
| Approval for running fork pull request workflows (`fork_pr_approval_policy`) | **All external contributors** | Public repositories only. A workflow on a pull request from anyone who is not a member of the repository waits for approval. |
| Check, workflow run, status, artifact and log retention (`retention_days`) | **90 days** | The longest period a public repository allows. A workflow can keep an artifact for less, as the integrity reports do. |

Sources: [Actions permissions REST API][actions-api] and
[Managing GitHub Actions settings for a repository][actions-settings].

<a id="rulesets"></a>

## 🛡️ Custom active protection

The baseline has no corresponding custom active rulesets. The rulesets below
are explicit additions. All are **Active**, have **no excluded refs**
and **no bypass actors**. Default-valued supporting parameters below are
included only to describe each custom rule completely; they are not counted
as separate deviations from GitHub defaults.

| Custom ruleset | Target | Added restrictions |
| --- | --- | --- |
| Protect main | `~DEFAULT_BRANCH` | Require PR, resolved conversations and strict up-to-date `Repository integrity` from GitHub Actions; allow merge commits only; block deletion and force pushes. |
| Protect release branches | `refs/heads/release/**` | Require PR, resolved conversations and the same strict required check; allow squash or merge commits; block deletion and force pushes. |
| Prevent release branch deletion | `refs/heads/release/**` | Independent deletion restriction. |
| Protect release tags | `refs/tags/v*` | Block updates, deletion and force pushes. |
| Protect traffic history | `refs/heads/traffic-history` | Block deletion and force pushes. Direct pushes stay allowed for the optional Traffic history workflow ([TRAFFIC.md](TRAFFIC.md)). |

### 🔬 Supporting parameters for the two PR rulesets

| Parameter | Default-branch rule | Release-branch rule |
| --- | --- | --- |
| Required status context | `Repository integrity` | `Repository integrity` |
| Expected check integration | GitHub Actions (`15368`) | GitHub Actions (`15368`) |
| Up-to-date requirement | Strict | Strict |
| Required check enforced on creation | Yes | No — branch bootstrap exception |
| Required approving reviews | 0 | 0 |
| Extra approval when Copilot opens a PR without a human collaborator | Yes | Yes |
| Required review-thread resolution | Yes | Yes |
| Permitted PR merge methods | Merge | Squash, merge |
| Administrator/app bypass | None | None |

**Operational implications:** zero approvals does not establish independent
review. The extra-approval parameter adds one required review only when the
required count is above zero, so it has no effect here until that count is
raised; GitHub's documentation does not state its default. The release
bootstrap exception does not exempt subsequent changes. Synchronize
integration branches through the approved PR process; do not rely on an
administrator bypass. Tag protection does not itself restrict tag creation or
authorize publishing a release.

The check name and branch patterns are reusable conventions, not built-in
GitHub jobs or automatically existing branches. Confirm that the required job
reports on every relevant PR before adopting these rules.

<a id="reuse"></a>

## ♻️ Reuse and maintenance

- Keep repository names, usernames, URLs, issue numbers and rule IDs out of this file.
- Verify the destination's actual default and inherited policies before copying a value.
- Include a setting only when both its configured value and comparison default are known.
- Remove rows that match the destination's default; retain only its actual differences.
- Keep unverified settings and historical evidence in the setup record.
- Recheck this comparison after platform-default changes or an approved configuration change.

> [!NOTE]
> This is intentionally not a complete settings inventory. Security features
> whose initial state depends on account, visibility or an applied security
> configuration are not classified as non-default without a verified comparator.
> Omission does not mean a feature is disabled or unimportant.

<a id="sources"></a>

## 📚 Default and ruleset references

| Reference | Supports |
| --- | --- |
| [Repository REST API][repo-api] | Declared option defaults and parameter meanings |
| [Private vulnerability reporting][pvr] | Availability on public repositories and how the setting is enabled |
| [Dependabot alerts][dependabot-alerts] and [security updates][dependabot-updates] | What the two Dependabot switches do and how they depend on each other |
| [Secret scanning][secret-scanning] | Availability by visibility and the push protection it enables |
| [Creating rulesets][ruleset-create] | Custom targets, bypass configuration and activation; the editor initially uses Disabled enforcement |
| [Available rules][ruleset-rules] | Meaning of selected protection rules, including editor defaults |
| [Actions permissions REST API][actions-api] and [Actions settings][actions-settings] | What each Actions setting does, where it is set and the endpoints that report it |
| [Security configuration defaults][security-defaults] | Why organization-applied defaults must be distinguished from a standalone baseline |

[repo-api]: https://docs.github.com/en/rest/repos/repos
[pvr]: https://docs.github.com/en/code-security/how-tos/report-and-fix-vulnerabilities/configure-vulnerability-reporting/configure-for-a-repository
[dependabot-alerts]: https://docs.github.com/en/code-security/concepts/supply-chain-security/dependabot-alerts
[dependabot-updates]: https://docs.github.com/en/code-security/concepts/supply-chain-security/dependabot-security-updates
[secret-scanning]: https://docs.github.com/en/code-security/concepts/secret-security/secret-scanning
[actions-api]: https://docs.github.com/en/rest/actions/permissions
[actions-settings]: https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/enabling-features-for-your-repository/managing-github-actions-settings-for-a-repository
[ruleset-create]: https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-rulesets/creating-rulesets-for-a-repository
[ruleset-rules]: https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-rulesets/available-rules-for-rulesets
[security-defaults]: https://docs.github.com/en/enterprise-cloud@latest/code-security/how-tos/secure-at-scale/configure-enterprise-security/establish-complete-coverage/create-custom-configuration

---

[↑ Back to comparison basis](#scope)

## Applying this baseline

The executable configuration is [repository-settings.json](../setup/repository-settings.json).
See [REPOSITORY_SETUP.md](../setup/REPOSITORY_SETUP.md) for preview, apply and verification commands.
The configuration holds all twelve general settings that the script manages, so
it also keeps the eight that match GitHub's defaults from drifting; only the four
that differ are listed above. The script applies them, the five security features
and the rulesets documented above; the configuration may hold any number of
uniquely named rulesets. It applies the public-only features only on a public
destination and names them in a note for any other visibility. It reports the five
Actions settings and never writes them: after an apply, its read-back fails while
one differs, naming each to set by hand. The configuration's `visibility` is
optional: `null` accepts any visibility, while `public` or `private` makes a
different visibility drift that the script reports and never changes; follow the
[distribution review](DISTRIBUTION.md) before changing it. Code scanning
is not part of this baseline: the CodeQL workflow
(`.github/workflows/codeql.yml`) runs it with advanced setup, so leave GitHub's
default setup off, which would reject the workflow's results. On a private
repository the workflow needs GitHub Code Security and the repository variable
`CODE_SCANNING` set to `enabled`; without them it skips the analysis.
The OpenSSF Scorecard workflow (`.github/workflows/scorecard.yml`) scores the
default branch's supply-chain posture weekly and on each push to it. It
publishes the results to the Scorecard service and to code scanning, and it is
skipped on a private repository; it reads the visibility from the API, because
a scheduled run carries no repository in its event. A low check there, such as branch protection
or signed releases, points to a setting of this baseline or of the release
process; it is not a required check.
The **Release** workflow (`.github/workflows/release.yml`) attests the build
provenance of its source bundle on a public repository. A private one needs
GitHub Enterprise Cloud for attestations and the repository variable
`RELEASE_ATTESTATIONS` set to `enabled`. Without the variable the workflow
skips the attestation and still drafts the release, whose notes say so; with
it but without Enterprise Cloud the attestation fails and no draft is created.
`Protect main` targets `~DEFAULT_BRANCH`, so it follows the default branch under
any name. The script matches each configured ruleset by name; one with no ruleset of
its name is covered when other active rulesets with no bypass actors or excluded
refs, as the plan leaves them, enforce it at least as strictly under any names,
and a named one with stronger parameters covers itself. A covered ruleset is not
drift and is never written, so renaming one in the configuration creates nothing
while the old one covers it; a weaker named one is replaced. Apply writes only the plan whose digest, printed by the preview,
I pass with `--approve-plan`, and journals each write to a new `--journal` file.

<a id="drift"></a>

### Scheduled drift check

The **Detect settings drift** workflow (`.github/workflows/settings-drift.yml`)
runs `tools/setup_repository.py --check --token-scope workflow` daily and on
demand, so a ruleset changed or deleted in the GitHub settings does not wait for
the next manual check. It uses the job's `GITHUB_TOKEN` with `contents: read`
only. I give it no administrator token: nothing it runs can change a setting,
and no stored credential is exposed to it. That token can read only part of the
baseline:

| Section | Compared by the workflow | Why |
| --- | --- | --- |
| Rulesets | Yes, except each bypass list | Reading rulesets needs only **Metadata: read**, which every workflow token has. GitHub returns a ruleset's bypass list only to a token that can write the ruleset. |
| Visibility | Yes | GitHub reports it to every token that can read the repository. |
| General settings | Those GitHub returns to the token, such as the Issues, Projects, Wiki, Discussions and downloads switches | GitHub returns the merge-related settings only to a token that can also write contents. |
| Private vulnerability reporting | Yes, on a public repository | Its endpoint needs only **Metadata: read**. |
| Dependabot alerts and security updates | No | Their endpoints need **Administration: read**, which no workflow permission grants. The workflow scope never calls them. |
| Secret scanning and push protection | No | GitHub reports them only to administrators of the repository. |
| Actions settings | No | Their endpoints need **Administration: read**. |

The run fails on drift in any value it reads: a managed ruleset that is missing,
disabled or changed in a rule, a condition or its target, a different
visibility, or a different general setting. A ruleset that other rulesets cover
counts as covered, as in `--check`. A value it cannot read is never drift and
never a match: the output prints it as `NOT CHECKED`, and the job summary lists
it under **Not checked**. When everything it reads matches, the result is
`PARTIAL` and the run passes; the list names what only `--check` with an
administrator's token verifies, as the [release closeout](../release/CLOSEOUT.md)
and the [setup guide](../setup/REPOSITORY_SETUP.md#-verify-the-result) run it.
A value is not checked only when GitHub leaves it out of its answer, or refuses
to read it with HTTP 401, 403 or 404. Any other failed request, such as a server
error or no answer, is an API failure, as is a refused read of the repository or
its rulesets, such as rulesets that the account's GitHub plan does not offer for the
repository: it fails the run with status 2. In a new repository the run fails until the baseline is applied.

Sources: [repository REST API][repo-api] (merge-related settings and
`security_and_analysis`), [ruleset REST API][rules-api] (bypass list),
[permissions of fine-grained tokens][token-permissions] and
[workflow token permissions][workflow-permissions].

[rules-api]: https://docs.github.com/en/rest/repos/rules
[token-permissions]: https://docs.github.com/en/rest/authentication/permissions-required-for-fine-grained-personal-access-tokens
[workflow-permissions]: https://docs.github.com/en/actions/reference/workflows-and-actions/workflow-syntax#permissions
