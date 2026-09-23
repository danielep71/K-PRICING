# 🔄 Maintenance and migration

> **Guide, not policy:** [docs/TEMPLATE_CONTRACT.md](../TEMPLATE_CONTRACT.md) is authoritative.

## Distinguish three versions

| Version | Meaning |
| --- | --- |
| Project `VERSION` | Your product's release version |
| `template_contract.version` | The required control baseline this repository adopted |
| Reusable workflow pin/interface | Reviewed orchestration revision and its callable interface |

Changing one does not automatically migrate the others. Read
[Template Contract](../TEMPLATE_CONTRACT.md) before assigning a new adopted
baseline. Older adopters are evaluated against their recorded version.

## Upgrade deliberately

1. Record the current source, adopted baseline and passing local/host evidence.
2. Compare the target migration notes and classify required, optional, breaking
   and not-applicable changes.
3. Apply reviewed source/configuration changes on a branch. Keep specialist
   checks and stronger live protections.
4. Run affected gates and profile tests. Review operational dependencies,
   permissions and CLI compatibility.
5. Update the adoption record only when the target contract is actually met.
6. Retain the previous tested revision and a reviewed rollback procedure.

The initializer is not an upgrade engine. Do not rerun it with different values
or replace a mature repository with a fresh template tree.

## Reusable workflow adoption

Keeping the copied workflow is supported. If you adopt the optional reusable
workflow, follow [Reusable Workflows](../REUSABLE_WORKFLOWS.md), including its
published full-SHA pin, permissions, expected profile, candidate output and
required-check naming. A pinned provider selects orchestration; the caller's
local scripts still run and need review. Never replace the full pin with a
floating branch or broad tag.

## Dependencies and drift

Use [Dependency Updates](../DEPENDENCY_UPDATES.md) for manual monitoring,
provenance, compatibility review and rollback. Do not enable automatic merges
as a shortcut around release evidence.

Template-maintenance portfolio tools can collect and compare read-only
observations. Missing adoption is an adoption decision; missing permissions are
unverified evidence. A conformance result is not proof of numerical correctness
or live Excel behavior. This guide does not authorize changes to existing
portfolio repositories.

## Keep documentation synchronized

Update the governing document with a changed contract, then reconcile its wiki
summary and inventory description. The wiki publication procedure compares
rendered bytes with the exact source commit. A newer candidate does not inherit
an older publication's certification merely because its page titles match.
Use the final documentation audit to reconcile all links and terminology.

---
[Home](Home.md) · [Previous](Workflow-and-tool-reference.md) · [Next](Troubleshooting.md)
