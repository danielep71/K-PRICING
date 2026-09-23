# 🚀 Initialize the project

> **Guide, not policy:** [docs/INITIALIZATION.md](../INITIALIZATION.md) is authoritative.

**Working directory:** the clean root of the newly created repository.
The [initialization authority](../INITIALIZATION.md) defines the complete token
catalogue, optional values and manual fallback.

## 1. Prepare values

Choose the profile and seven required values. Use a real public project name,
specific description, repository path and monitored security contact. The values
are recorded in `.github/initialization.json`; they must not contain secrets.
`VERSION` becomes `0.0.0`, a pre-release setup sentinel, not a published version.

## 2. Preview

In Git Bash, edit the example values once in this array. Each argument remains
quoted even when a name or description contains spaces:

```bash
init_args=(
  --profile library
  --set "PROJECT_NAME=Example Project"
  --set "PROJECT_TAGLINE=Reusable Excel calculations"
  --set "PROJECT_DESCRIPTION=Reusable scalar calculations for Excel VBA callers."
  --set "REPOSITORY_PATH=OWNER/NEW-REPOSITORY"
  --set "MAINTAINER_NAME=Example Maintainer"
  --set "SUPPORT_CONTACT=security@example.com"
  --set "COPYRIGHT_YEAR=2026"
)
python tools/initialize_repository.py "${init_args[@]}"
```

Expected: a deterministic plan of creates, updates and removals with content
digests; no files changed. Inspect the chosen profile, security routing,
description, template-only deletions and retained operational tools.
Stop on missing/unknown values, a dirty tree or a wrong target repository.

Omitting the optional social preview removes the template's sample artwork.
To retain the supplied PNG, add
`--set "SOCIAL_PREVIEW_PATH=assets/social-preview.png"` to the array before both
preview and apply. A custom preview must already be tracked in the clean tree.
Repeatable test commands and limitations follow the authority's `--add` rules.

## 3. Apply the reviewed inputs

```bash
python tools/initialize_repository.py "${init_args[@]}" --apply
git diff --stat
git diff
git diff --check
git add -A
python tools/check_repo.py --root .
python tools/check_vba_public_api.py --root .
```

Inspect the staged tree before committing. Expected: generated mode, one
profile, your identity, the fixed starter VBA names, an initialization record,
no template-maintenance wiki/tooling, and a passing gate. The whole initial diff
belongs to this isolated setup branch; do not use blanket staging in a checkout
that contains unrelated work.

## 4. Commit and prove idempotency

```bash
git commit -m "Initialize the selected project profile"
python tools/initialize_repository.py "${init_args[@]}"
git status --short
git push -u origin setup/initialize
```

Expected: the same inputs yield `no-op`; status is empty. Repeating different
inputs must fail. Commit before this check so the initializer sees a clean tree.
If a command fails, stop and read its finding; do not edit the initialization
record to make mismatched values appear accepted.

Review and merge the setup PR, then return to the default branch:

```bash
git switch main
git pull --ff-only
git rev-parse HEAD
```

Record this initialized default-branch SHA for GitHub provisioning. Keep a
separate checkout of the original template if using its provisioner: that tool
is intentionally removed from the generated repository.

---
[Home](Home.md) · [Previous](Create-the-repository.md) · [Next](Configure-GitHub.md)
