# 🔗 Versioned Reusable Workflows

[![Interface: v1](https://img.shields.io/badge/interface-v1-1D76DB)](#interface-v1)
[![Permissions: read only](https://img.shields.io/badge/permissions-read%20only-217346)](#trust-boundary)

This document owns the callable workflow interface, adoption and compatibility
policy. [TEMPLATE_CONTRACT.md](TEMPLATE_CONTRACT.md) owns the versioned required
control set. Adoption is optional for contract `1.2.0`: keeping the copied
workflow is supported and does not weaken conformance.

## 🎯 What is reusable

`.github/workflows/static-checks.yml` is both the existing push/PR workflow and
the `workflow_call` entry point. There is one implementation, not a second copy
of the job. It runs the canonical repository gate, focused source/contract/release
semantics gates, their self-tests, Python quality checks, authoritative workflow
validation, evidence upload and the fail-closed final verdict.

Excel compilation, numerical regression, UI-state restoration, workbook builds,
packaging and release certification remain separate caller-owned jobs. The
generic result never claims to certify any of those activities.

<a id="interface-v1"></a>

## 📋 Interface v1

The initial interface revision is `1.0.0`, independent of the product release
and template contract. Annotate its immutable caller pin with `# v1.0.0`.
A version comment describes the pin; it never substitutes for the full SHA.

### Published immutable pin

| Interface | Provider commit | Publication |
| --- | --- | --- |
| `1.0.0` | `466c48d9f4a5f984ae58561d1441089ef09975d1` | Published by commit SHA on 2026-09-08 after all three hosted consumer fixtures passed |

This is commit-based publication of the workflow interface, **not** publication
of the v1.2.0 product release, and no product tag was created or moved. The
verified development snapshot was promoted unchanged; its historical fixture
comments say `v1.0.0-dev`. Use the exact SHA above, not the current branch head.
Completion evidence is maintained in upstream issue #23. Fixture test refs are
disposable and are never consumer pins or merge candidates.

| Surface | Contract |
| --- | --- |
| Input `expected-profile` | Optional string, default empty; otherwise exactly `application`, `library` or `ui-component`. A nonempty value requires matching generated mode and profile. Invalid or mismatched values fail. |
| Output `candidate-sha` | String: the exact checked-out caller commit. Consume only after successful completion; an output alone is not a passing verdict. |
| Permissions | Caller grants `contents: read`; no write permission is requested. |
| Secrets | None declared or required. Do not use `secrets: inherit`, environments or a personal access token. Checkout uses the scoped automatic GitHub token and does not persist credentials. |
| Runtime | GitHub-hosted `ubuntu-24.04`; tool/runtime/action pins are versioned in the called YAML. |
| Evidence | Existing JSON/Markdown reports, job summary and `static-checks-RUN_ID-RUN_ATTEMPT` artifact, retained 30 days. |
| Invocation limit | One invocation per workflow run: artifact names are run-scoped. Use separate runs for separate consumers, not a caller matrix. |
| Supported caller | Initialized contract `1.2.0` repository with its complete gate scripts and quality configuration. Normal `push` or `pull_request` events, or manual dispatch. No privileged `pull_request_target` use. |

<a id="trust-boundary"></a>

## 🛡️ Trust boundary and identity

Pinning this YAML pins **orchestration**, not a centrally downloaded checker
package. Checkout reads the **caller repository at `github.sha`**, including its
reviewed local Python scripts, policy, manifest and dependencies. For a PR this
is normally GitHub's test merge commit, not just the head branch SHA. The job
checks checkout identity explicitly. It does not check out or run template
scripts from a floating provider branch.

Keep required gate scripts and their CLI contracts when adopting the workflow.
Review changes to both the caller scripts and the workflow pin. Read-only token
permissions and no inherited secrets limit exposure, but local scripts still
execute code: this is a CI check, not an independent tamper-proof attestation.
Do not invoke it on a privileged runner with credentials or private network
access. Jobs that require secrets belong in a separate trusted workflow.

GitHub documents the job-level calling syntax, full-SHA references and permission
restrictions in [Reuse workflows](https://docs.github.com/en/actions/how-tos/reuse-automations/reuse-workflows).

## 🚀 Adoption, step by step

1. Initialize a supported profile and pass the existing local gates first.
2. Select the published immutable provider commit above, with green hosted
   evidence in issue #23 of the repository recorded by `template_contract.source`.
   Never use `main`, a release branch or a moving
   major tag. Even with a named release, resolve and pin its full commit SHA.
3. Replace the generic job in the caller's `static-checks.yml` with a job-level
   call to `SOURCE/.github/workflows/static-checks.yml@SHA`, where `SOURCE` is
   `template_contract.source`,
   substituting that exact 40-character SHA. Preserve caller triggers, branch
   filters and concurrency. Set workflow permissions to `contents: read`.
   The identity gate permits this exact upstream reference in workflow files;
   floating refs, other sources and leftover template branding remain rejected.
4. Under the call's `with`, set `expected-profile` to your initialized profile.
   Keep specialist jobs local, with their own runner and stricter commands.
   Use `needs: generic` when they depend on the generic result; do not use
   `continue-on-error` to turn a required failure into a pass.
5. Open a PR. Verify the generic job, evidence artifact and specialist jobs.
   Update required-check settings to the actual emitted job name before retiring
   the old required check; never leave a protection gap. Do not merge on output
   SHA presence alone.


## 🔄 Compatibility, deprecation and rollback

Discovery, provenance evidence and manual approval for pin updates follow
[DEPENDENCY_UPDATES.md](DEPENDENCY_UPDATES.md). The interface-specific rules below
remain authoritative for compatibility and coupled caller-tool migrations.

- Workflow revisions keep the interface-v1 inputs/output meanings and the
  caller-local tool boundary. New optional inputs may be additive; existing
  calls must continue to work at their original pins.
- A removed input/output, changed default or meaning, newly required input,
  changed tool ownership, or incompatible local CLI requirement is breaking.
  Publish a new interface major and template-contract version, extend
  `CONTRACT_RULE_SETS`, and document exact migration steps in the contract
  history. Never silently retarget an existing pin.
- Deprecation is announced in this document and the changelog at least one
  released minor version before removal from the recommended interface.
  Historical commits remain addressable; deprecated does not mean patched.
  Security advisories may require urgent migration.
- Updates arrive as reviewed caller PRs containing the old/new workflow SHA,
  compatibility notes and hosted evidence. No scheduled job moves a pin.
- Roll back by reverting the caller pin and any coupled local-tool migration
  together, then rerun generic and specialist checks. Keep the previous passing
  SHA and evidence until the replacement is accepted. If the provider is
  unavailable, restore the previously reviewed copied workflow and verify check
  protection names; do not fall back automatically or bypass required checks.
