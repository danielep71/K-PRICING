# 🧩 Production Source

`src/` contains the authoritative production VBA source. A user must be able to reconstruct the supported workbook or add-in from this directory and the instructions in the root documentation.

## Canonical separation

Create only the subdirectories the project actually needs:

| Location | Contents | Excludes |
| --- | --- | --- |
| `src/modules/` | Public standard modules: supported procedures, worksheet functions, constants, and thin facades | Test harnesses and private implementation engines |
| `src/core/` | Internal standard modules: parsing, calculation, validation, and other implementation details | Supported public entry points |
| `src/classes/` | Production class modules, including state managers, event sinks, and UI hooks | Test doubles used only by the harness |
| `src/forms/` | Production UserForms; each `.frm` stays beside its required `.frx` | Screenshots and distributable workbooks |

A small project may keep production components directly in `src/` when further subdivision would add no clarity. If it does, document each component's role in `INSTALLATION.md`.

## Neutral starter

| Import order | Path | Component | Role |
| ---: | --- | --- | --- |
| 1 | `core/ProjectCore.bas` | `ProjectCore` | Internal implementation guarded by `Option Private Module` |
| 2 | `modules/ProjectFacade.bas` | `ProjectFacade` | Supported public façade recorded in `docs/PUBLIC_API.txt` |

The starter implements one stateless ratio operation only to prove the
façade/core, error, import, and test contracts. Replace it with real project
behavior before release, or document why the sample remains supported.

## Rules

- Preserve exported VBE component names, headers, and text encoding.
- Keep public facades thin and move reusable implementation logic into `core/`.
- Mark every production class as public-surface or internal in its header or architecture documentation.
- Keep a UserForm's `.frm` and `.frx` together and import only the `.frm` through the VBE.
- Do not place tests, examples, release binaries, generated evidence, or local workbooks here.
- Document the exact production manifest and import order in `INSTALLATION.md`.

Delete this README only if real source files and equivalent project documentation make the directory's purpose equally explicit.
