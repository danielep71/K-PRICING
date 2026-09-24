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

## Migrated KPR date layer

| Import order | Path | Component | Role |
| ---: | --- | --- | --- |
| 1 | `core/KPR_Core_Err.bas` | `KPR_Core_Err` | Internal native-error boundary |
| 2 | `core/KPR_Core_Parse.bas` | `KPR_Core_Parse` | Internal strict scalar/control parsing |
| 3 | `core/KPR_Core_Dates.bas` | `KPR_Core_Dates` | Internal Gregorian/date-domain core |
| 4 | `core/KPR_Core_Array.bas` | `KPR_Core_Array` | Internal shape/materialization engine |
| 5 | `modules/KPR_DATES_DAYS.bas` | `KPR_DATES_DAYS` | Supported 22-function worksheet façade |

The five production modules are migrated from the frozen source revision recorded
in `docs/MIGRATION_PROVENANCE.md`. The four cores remain `Option Private Module`;
the façade alone defines the supported calculation API in `docs/PUBLIC_API.txt`.

Provenance boundary against that frozen source:

| Component | Relationship to the frozen source |
| --- | --- |
| `KPR_Core_Err`, `KPR_Core_Parse`, `KPR_Core_Array` | Byte-for-byte imports; source Git blob IDs unchanged |
| `KPR_Core_Dates` | Destination-adapted: issue #32 corrects the pillar range classification so oversized valid quantities return `PILLAR_AGGREGATE_RANGE` (`#NUM!`) |
| `KPR_DATES_DAYS` | Format-only delta: seven trailing spaces removed from comment banners; no executable change |

The frozen source repository is unchanged; these are destination deltas, not
source history. `docs/MIGRATION_PROVENANCE.md` holds the exact source and
destination blob IDs and the justification for each delta.

## Rules

- Preserve exported VBE component names, headers, and text encoding.
- Keep public facades thin and move reusable implementation logic into `core/`.
- Mark every production class as public-surface or internal in its header or architecture documentation.
- Keep a UserForm's `.frm` and `.frx` together and import only the `.frm` through the VBE.
- Do not place tests, examples, release binaries, generated evidence, or local workbooks here.
- Document the exact production manifest and import order in `INSTALLATION.md`.

Delete this README only if real source files and equivalent project documentation make the directory's purpose equally explicit.
