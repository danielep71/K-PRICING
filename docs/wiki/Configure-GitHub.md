# ⚙️ Configure GitHub

> **Guide, not policy:** [docs/POST_CREATION_CHECKLIST.md](../POST_CREATION_CHECKLIST.md) is authoritative.

**Prerequisite:** initialized source is reviewed, merged and available at an
exact default-branch SHA. Source initialization changes files; it does not prove
that GitHub settings are correct.

## 1. Review the desired state in source

In the new repository, review `.github/provisioning-policy.json`,
`.github/repository-profile.json` and `.github/labels.json`. Commit any intended
policy change before planning. Keep selected domain overlays and local labels
explicit. Routine label synchronization may prune unlisted labels; reconcile
the overlay or reviewed pruning policy before running it.

## 2. Choose the setup path

For a guided manual setup, follow each checkbox in
[Post-Creation Checklist](../POST_CREATION_CHECKLIST.md) in the new repository.

For the guarded provisioner, use a **separate template-maintenance checkout**.
The tool is removed during project initialization. Replace the repository,
profile and full initialized default-branch SHA:

```bash
python tools/provision_repository.py --repository OWNER/NEW-REPOSITORY --profile library --contract-version 1.2.0 --source-sha FULL_INITIALIZED_SHA --output ../setup-plan.json --summary ../setup-plan.md
```

This reads GitHub; it does not create a repository or change settings. Inspect
actions, KEEP entries, blockers, exceptions and the plan digest. Planning exit
zero means a plan was produced, not that setup is complete.

Only after reviewing that exact plan, in a trusted environment with the required
credential available through `GH_TOKEN`, run:

```bash
python tools/provision_repository.py --repository OWNER/NEW-REPOSITORY --profile library --contract-version 1.2.0 --source-sha FULL_INITIALIZED_SHA --apply --approve-plan REVIEWED_PLAN_SHA256 --journal ../setup-apply-journal.json --output ../setup-apply.json --summary ../setup-apply.md
```

Use fresh evidence filenames, store no token in them, and stop on any failure.
A changed plan digest requires another review. Partial changes are not rolled
back automatically. [Provisioning](../PROVISIONING.md) owns the exact supported
operations, permissions and recovery behavior.

## 3. Complete the live checklist

| Area | Action and read-back |
| --- | --- |
| Identity and features | Check description, topics, Issues and the owner of every enabled feature |
| Issue chooser | Open New issue; confirm three forms, no blank issues and private security routing |
| Labels | Verify resolved names, colors and descriptions after trusted synchronization |
| Default branch | Confirm the intended branch; source branch names alone do not set GitHub's default |
| Merge policy | Review methods, manual merge, automatic branch deletion and message defaults |
| Branch protection | Active PR requirement, strict `Repository integrity`, deletion/force-push protection, intended scope and no unintended bypass |
| Version tags | Protect existing `v*` tags against update and deletion while allowing new tag creation |
| Security | Enable private vulnerability reporting and check the monitored reporting route |
| Wiki | Enable only with a maintainer and a version/drift procedure |

## 4. Retain verified evidence

Read metadata, labels and rulesets back from the API or UI after applying them.
Keep repository, profile, SHA, operator, time and observed values. An inaccessible
ruleset is unverified; it is not proof that protection is absent or satisfied.
The provisioner deliberately does not certify private reporting, rendered forms,
all merge-message settings, source quality or runtime behavior. Complete these
manual checks before calling setup complete.

---
[Home](Home.md) · [Previous](Initialize-the-project.md) · [Next](Run-the-quality-gates.md)
