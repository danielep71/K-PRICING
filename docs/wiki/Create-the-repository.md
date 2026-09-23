# 🌱 Create the repository

> **Guide, not policy:** [docs/INITIALIZATION.md](../INITIALIZATION.md) is authoritative.

## 1. Confirm the starting edition

Open the template and inspect its **default branch**, release notes and
`.github/repository-profile.json`. Record the source revision and adopted
contract version. A wiki published from a development branch does not change
what GitHub's template button copies.

Use the normal creation path after the intended edition reaches the default
branch. A local development pilot may instead use an explicitly reviewed source
snapshot, but must label that departure from the live template path.

## 2. Create a new GitHub repository

Use **Use this template → Create a new repository**. Select the intended owner,
choose a descriptive repository name and visibility, and leave **Include all
branches** unchecked. Do not overwrite or repurpose an existing project.

Before creating it, confirm that the target account/plan and chosen visibility
expose repository rulesets if you intend to use the maintained provisioner. A
private repository can return a platform-capability `403` for rulesets even when
normal Git authentication works. If rulesets are unavailable, choose a supported
visibility/plan or document a separately reviewed governance path before
initialization; do not treat missing ruleset access as an empty ruleset state.

GitHub copies files from the template; labels, rulesets, secrets, settings and
the separate Wiki repository need their own setup. Template branches also have
unrelated histories: importing all branches is not a release workflow.
See [GitHub's template model](https://docs.github.com/en/repositories/creating-and-managing-repositories/creating-a-template-repository).

## 3. Clone the new repository

The commands in this guide use **Git Bash**, including on Windows. Replace
`OWNER` and `NEW-REPOSITORY` with the new repository's values before running:

```bash
git clone https://github.com/OWNER/NEW-REPOSITORY.git
cd NEW-REPOSITORY
git remote -v
git branch --show-current
git status --short
git rev-parse HEAD
python --version
```

Expected: the remote is your new repository, the default branch is as intended,
and status prints nothing. Keep the starting SHA in your setup record outside
the checkout. Do not create a notes file inside the repository before the
initializer's clean-tree check.

If authentication is requested, sign in through Git's credential manager or the
GitHub web interface. If Git or Python is not found, repair PATH and reopen the
terminal before continuing.

Provisioning API access is a separate prerequisite from clone/push
authentication. Plan mode is GET-only, but authenticated API reads can still be
required because anonymous public-repository metadata may omit merge-policy
fields needed for deterministic planning. Resolve that read access before the
live provisioning step rather than defaulting missing metadata.

## 4. Work on a setup branch

```bash
git switch -c setup/initialize
```

This gives the initialized diff a reviewable branch. After initialization,
commit and push that branch, review it in a PR and merge it through the
repository's configured process. Configure the live protection baseline before
normal development begins; do not bypass an organization policy to bootstrap.

Continue only when the source contains the initializer and the contract edition
this guide describes. A missing file is a version mismatch to resolve, not a
reason to copy isolated scripts from a newer branch.

---
[Home](Home.md) · [Previous](Choose-a-profile.md) · [Next](Initialize-the-project.md)
