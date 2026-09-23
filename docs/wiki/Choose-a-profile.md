# 🧭 Choose a profile

> **Guide, not policy:** [docs/REPOSITORY_STRUCTURE.md](../REPOSITORY_STRUCTURE.md) is authoritative.

Decide what the project owns before initializing. Initialization selects exactly
one profile; it is not a command to rerun casually when the product changes.

| Question | Profile | Additional evidence |
| --- | --- | --- |
| Do callers use reusable functions without an owned user interface? | `library` | Public API and caller contract |
| Does another workbook embed a bounded interactive component? | `ui-component` | UI state, cleanup, recovery, DPI/accessibility and lifecycle |
| Does this solution own deployment and workbook/add-in lifecycle? | `application` | Startup, shutdown, upgrade, recovery, packaging and end-to-end smoke |

## What all three receive

The required starter is the same: `src/modules/ProjectFacade.bas`,
`src/core/ProjectCore.bas`, and `tests/modules/ProjectTests.bas`.
The optional example is `examples/modules/ProjectExample.bas`.
Profile selection does not generate UI controls, financial calculations or an
application shell. A README in a directory cannot replace substantive VBA.

The public facade is the supported external boundary. Core, tests and examples
remain project-private. Keep that dependency direction when adding functionality.

## Additional structure

- **Library:** core, modules and test-module directories; classes only if needed.
- **UI component:** also reserves `src/classes`; add forms, binary companions
  and Ribbon resources only when the design actually uses them.
- **Application:** also reserves `src/classes` and `src/workbook`; implement
  host startup/shutdown and deployment deliberately.

The initializer may create directory README files. They explain reserved
locations and are not claims that those roles have been implemented.

## Record the decision

Write a short project boundary: supported callers, owned state, lifecycle,
public entry points and supported Excel environments. Choose the closest
existing profile and add stronger project checks where needed. Do not weaken
common gates or invent a new profile by editing a string alone.

If the boundary later changes, review the source layout, profile policy,
release evidence and migration consequences together in a separate PR.

---
[Home](Home.md) · [Previous](Home.md) · [Next](Create-the-repository.md)
