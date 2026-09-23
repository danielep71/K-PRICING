# 🛠️ Develop safely

> **Guide, not policy:** [CONTRIBUTING.md](../../CONTRIBUTING.md) is authoritative.

## 1. Define the change before editing

Start a focused branch from the protected development baseline. Record supported
callers, inputs, outputs, errors, compatibility and state ownership. Open an
issue for a nontrivial API or architecture change. Keep source formatting and
semantic changes reviewable.

## 2. Put each component in its owned location

| Role | Location | Boundary |
| --- | --- | --- |
| Public facade | `src/modules` | Supported external API and caller-facing errors |
| Internal core | `src/core` | In-project implementation; no dependency back to tests or examples |
| Classes/forms/workbook exports | Applicable `src/` directory | Only when the selected design actually owns that role |
| Regression harness | `tests/modules` | Deterministic cases, failure reporting and cleanup |
| Examples | `examples/modules` | Demonstrations through the supported facade |

Register components and roles in `.github/repository-profile.json`.
Keep the mandatory substantive starter until its replacement contract is
reviewed; deleting it and leaving directory instructions fails the gate.

## 3. Keep exported source reviewable

Use [VBA House Style](../VBA_HOUSE_STYLE.md): procedure contracts, inline variable
annotations, comments below execution-phase banners and explicit cleanup intent.
Preserve the first export attribute, encoding and CRLF working-tree endings.
Keep `.frm` and `.frx` companions together; do not paste binary data into text.

Follow [Installation](../../INSTALLATION.md) when replacing modules in a workbook.
Remove/replace the old component deliberately so importing does not silently
create a second suffixed module. Export the updated source back to its governed
path and review the diff before committing.

## 4. Update contracts and tests together

A public declaration change needs matching rows and signature records in
`docs/PUBLIC_API.txt`, caller documentation and regression coverage. A change to
an error, default, side effect or state lifetime can be an API change even if
the procedure name is unchanged.

Tests must make omissions observable. Keep case/argument boundaries, expected
errors, repeatability and cleanup evidence. Do not change expected counts just
to make an incomplete run pass. If the harness intentionally evolves, review
its evidence policy and release consumers in the same change.

## 5. Review and merge

Run source and host checks appropriate to the change; add user-visible history
under Unreleased. Inspect the full staged diff and the PR evidence template.
State which environments were tested and which were not. Resolve conversations
and use the configured manual merge process. Reconcile candidate identity if the
merge changes the source SHA before any release certification.

---
[Home](Home.md) · [Previous](Run-the-quality-gates.md) · [Next](Prepare-and-publish-a-release.md)
