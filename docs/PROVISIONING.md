# 🛠️ Guarded Repository Provisioning

[![Default: plan](https://img.shields.io/badge/default-read%20only%20plan-217346)](#plan)
[![Apply: explicit](https://img.shields.io/badge/apply-reviewed%20digest-1D76DB)](#apply)

The provisioner configures one explicitly initialized generated repository.
It does not create repositories, initialize source, merge branches, publish
releases or migrate the existing portfolio. This milestone develops the tooling
and tests it against a simulated GitHub API; no existing portfolio is modified.
The tool, guide and fixtures are template-maintenance files removed during
generation. The optional `.github/provisioning-policy.json` remains in generated
source and resolves the project description from the immutable initialization
record using `@initialization:PROJECT_DESCRIPTION`, or an explicit reviewed
description string. No executable/configuration placeholder exception is added.

## 📋 Versioned source authority

Inputs are explicit repository, profile, adopted contract version and exact
default-branch SHA. The tool reads four JSON files at that SHA:

| File | Purpose |
| --- | --- |
| `.github/repository-profile.json` | Generated mode, repository identity, adopted contract and label selection |
| `.github/initialization.json` | Initializer's matching profile, contract and recorded repository value |
| `.github/provisioning-policy.json` | Description, topics, feature/merge settings, required checks and exception rationale |
| `.github/labels.json` | Core plus selected profile/domain labels |

The schema-1 provisioning adapter supports contract **1.2.0**. Older or unknown
versions fail explicitly; they are never silently upgraded. The versioned tool
defines active branch and tag rule payloads for this adapter, including no bypass,
strict `Repository integrity`, and immutable existing version tags. Stronger
requirements can be listed in `required_checks`; the universal job remains.

## 🧭 Preservation and supported operations

| Area | Plan/apply behavior |
| --- | --- |
| Description | Set the initialized versioned policy value |
| Topics | Add the union of existing, declared and Excel/VBA/profile topics; never remove local topics |
| Features | Apply declared booleans; disabling an enabled wiki, projects, discussions or template flag requires a recorded exception |
| Merge methods | Preserve an already narrower method instead of enabling it; at least one desired method must remain enabled |
| Labels | Add missing resolved labels; replacement needs a recorded exception; retain extra labels |
| Rulesets | Preserve equal/stronger applicable rules; add missing baseline rulesets; never edit or delete existing rulesets |

An existing ruleset with a baseline name but insufficient controls blocks apply
and requires a separately reviewed migration. Rules with bypass or exclusions
cannot prove coverage. Adding baseline protection does not remove existing
reviewers, specialist checks, scope restrictions or stronger controls.

Examples of exception keys in the committed policy:

```json
{
  "feature:has_wiki": "Reviewed unused generated wiki; no maintained content",
  "label:bug": "Reviewed replacement of the default GitHub label metadata"
}
```

An exception is a rationale, not approval by itself: committing it changes the
source SHA and plan digest, requiring review again. It cannot waive enabled
issues, disabled auto-merge, universal core labels or `Repository integrity`.
Destructive operations such as deleting labels/rulesets, changing visibility,
renaming/deleting a repository, bypass grants and source/ref writes are outside
the supported API surface. Even an exception does not enable these operations;
they require separate reviewed work. Review label overlays or `prune: false`
before the existing sync workflow runs if extra local labels should survive.

<a id="plan"></a>

## 1. 🔎 Generate and review a plan

First complete source initialization, commit the intended provisioning policy,
and obtain the default-branch commit SHA.

Before the first live plan, verify that the target repository/account plan and
chosen visibility expose repository rulesets and that the operator can read the
complete repository metadata required for deterministic planning. Git clone or
push authentication is not proof of those API capabilities. A private
repository can return a platform-capability `403` for rulesets, while anonymous
reads of a public repository can omit merge-policy fields such as
`allow_merge_commit`. Plan mode remains GET-only, but it may still require an
authenticated API context. If rulesets or required metadata cannot be read,
resolve the visibility/plan/access limitation first; the provisioner fails
closed and never defaults missing state.

From the template-maintenance checkout:

```bash
python3 tools/provision_repository.py --repository owner/generated-project --profile library --contract-version 1.2.0 --source-sha FULL_INITIALIZED_SHA --output plan.json --summary plan.md
```

This performs GET requests only. Inspect the complete `before` snapshot, exact
request bodies, KEEP entries, blockers, exceptions and `plan_sha256`. A plan
with changes may exit 0: that means planning succeeded, not that settings were
applied. A blocked plan exits 1. Missing access, truncated collections, missing
source or mismatched identity prevents an actionable plan and exits 2.

The digest covers repository ID/name, source SHA, version/profile, source policy,
observed supported settings and exact proposed actions. Volatile API timestamps
are excluded. Repeated unchanged observations produce the same plan.

<a id="apply"></a>

## 2. 🔐 Apply the exact approved plan

Use a trusted environment with an appropriately scoped `GH_TOKEN`, kept out of
arguments and files. The operator explicitly approves the displayed digest:

```bash
python3 tools/provision_repository.py --repository owner/generated-project --profile library --contract-version 1.2.0 --source-sha FULL_INITIALIZED_SHA --apply --approve-plan REVIEWED_PLAN_SHA256 --journal apply-journal.json --output apply.json --summary apply.md
```

Apply rebuilds the plan from live reads and refuses a changed digest. It writes
and flushes an atomic journal before the first request and before every later
write. Use a fresh journal filename; earlier receipts cannot be overwritten.
Report/summary/journal destinations must differ. GETs are repeated before each
write to detect intervening changes; all supported settings are independently
read back afterward. A second approved no-change plan applies with zero writes.

There is no GitHub transaction spanning these APIs. A concurrent change can
still race between a read and write; the tool does not claim atomic remote
application. An API response is not proof: final verification must contain no
remaining actions or blockers. A failed or uncertain write stops the sequence
and is retained in the journal. Writes are never automatically retried and
partial changes are never automatically rolled back. Collect a new plan,
inspect what actually applied, and approve any recovery separately.

## 3. 📸 Evidence and limits

Keep the before/plan, apply journal and independent final verification together.
Redact private data before sharing. API responses and source inputs never
include the token. Redirects are refused and the transport targets only the
explicit repository on `api.github.com`. Apply permits only supported metadata,
topic, label and ruleset-create requests; plan mode has no write permission in
the transport.

[GitHub repository settings](https://docs.github.com/en/rest/repos/repos#update-a-repository)
and [ruleset creation](https://docs.github.com/en/rest/repos/rules#create-a-repository-ruleset)
require the corresponding trusted write permissions, including Administration
for rulesets. Contents/metadata/labels/ruleset reads must also be available.
Unavailable ruleset bypass data is not treated as an empty bypass list. Plan
limitations and missing access must be resolved by the maintainer outside this
tool; it never escalates credentials, changes visibility or requests an upgrade.

Private vulnerability reporting, issue chooser rendering, sponsorships,
commit-message defaults, ownership of enabled features, source quality and
runtime/release certification remain manual or separate checks under the
[post-creation checklist](POST_CREATION_CHECKLIST.md). Provisioning success is
limited to supported settings and never certifies that entire checklist.

## 🧪 Validation

```bash
python3 tools/test_provision_repository.py -v
```

Hosted tests use a stateful simulated API and temporary journals. They cover
all profiles, exact request scope, plan repeatability, idempotence, stale
approval, mismatched target/version/SHA, preserved stronger controls, reviewed
exceptions, missing bypass evidence, ignored/partial writes and journal failure.
They do not demonstrate a live administrative deployment or Excel execution.
No production credentials are used by the fixture workflow.
