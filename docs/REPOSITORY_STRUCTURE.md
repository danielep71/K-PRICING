# 🗂️ Canonical Repository Structure

[![Model: source first](https://img.shields.io/badge/model-source--first-217346)](#canonical-root)
[![Directories: purposeful](https://img.shields.io/badge/directories-purposeful-success)](#directory-rules)
[![VBA layout: explicit](https://img.shields.io/badge/VBA-layout%20explicit-6f42c1)](#vba-source-separation)
[![Profiles: 3](https://img.shields.io/badge/profiles-3-1D76DB)](#-profile-specific-alternatives)

This document defines where every durable project artifact belongs in a
repository created from the canonical template. The goal is one authoritative
location per responsibility, without empty decoration or competing directory
conventions.

<a id="canonical-root"></a>
## 📁 Canonical root

Every generated repository starts with these six directories:

```text
src/       authoritative production VBA source
tests/     regression source, fixtures, and expected results
examples/  reproducible examples and demo builders
assets/    versioned non-code project assets
docs/      durable technical and contract documentation
tools/     deterministic validation, packaging, and evidence tooling
```

Git does not track empty directories. Until a directory contains project material, its instructional `README.md` explains why it exists and what may replace it. Do not use `.gitkeep` files.

| Directory | Owns | Must not own |
| --- | --- | --- |
| `src/` | Production VBA components required by the supported product | Tests, demos, generated workbooks |
| `tests/` | Test modules, stable fixtures, expected results | Production entry points, transient results |
| `examples/` | Supported-API examples and reproducible demo builders | Release evidence, duplicated tests |
| `assets/` | Source visuals, icons, diagrams, other versioned static inputs | Generated release packages |
| `docs/` | Contracts, architecture, methods, migrations, maintained plans | Duplicate copies of root governance files |
| `tools/` | Deterministic repository checks, packaging and provenance scripts | Product logic and workflow YAML |

<a id="vba-source-separation"></a>
## 🧩 VBA source separation

Create these `src/` subdirectories only when the project has corresponding components:

| Location | Canonical responsibility |
| --- | --- |
| `src/modules/` | Public standard modules and thin supported facades |
| `src/core/` | Internal standard modules for calculations, parsing, validation, and shared implementation |
| `src/classes/` | Production class modules, state managers, event sinks, and UI hook classes |
| `src/forms/` | Production UserForms with each `.frm` adjacent to its `.frx` resource |

### Public modules

Public worksheet functions, macros, supported enums, and entry points belong in `src/modules/`. Public modules should validate the call contract and delegate substantial work rather than duplicating algorithms.

### Internal and core modules

Private algorithms, parsers, numerical kernels, and host-independent helpers belong in `src/core/`. Internal modules may be imported with the product but are not automatically part of the supported public API.

### Classes and UI

Production `.cls` files belong in `src/classes/`, whether they implement public objects or internal state, events, or UI hooks. State the status in the component header or architecture documentation. UserForms belong in `src/forms/`; keep binary `.frx` companions adjacent and never import an `.frx` separately.

### Test modules

All regression and certification modules belong under `tests/`, normally `tests/modules/`. Test doubles and fixtures used only by the harness also remain under `tests/`. No test component belongs in the production installation manifest.

A very small project may keep production components directly under `src/` and test components directly under `tests/`. Subdivision is required only when it materially clarifies ownership.

## 🧪 Examples versus tests

An example demonstrates supported use and may prioritize clarity. A test establishes a repeatable assertion and must report failure deterministically. Do not make one file silently serve both roles.

- Put runnable learning material and demo builders in `examples/`.
- Put assertions, fixtures, oracles, and release-certification entry points in `tests/`.
- Put a reusable test-data generator in `tools/`, with generated fixtures written to `tests/fixtures/` only when they are reviewed and stable.

## 🧭 Profile-specific alternatives

Canonical names are the default for new projects. An alternative is legitimate only when it expresses a real project profile or preserves a material compatibility contract.

| Alternative | Legitimate when | Required control |
| --- | --- | --- |
| `demo/` instead of `examples/` | The interactive demo is a distinct distributable profile, an established public path, or a packaging input | Explain the profile in the root README; do not also use `examples/` for the same material |
| `test/` instead of `tests/` | An existing repository has stable scripts, links, or release evidence bound to `test/` | Record it as a compatibility exception; new repositories still use `tests/`; never keep both |
| `images/` instead of `assets/` | A renderer, package format, or stable URL requires the name | Document the constraint; do not split equivalent visuals across both locations |
| `dist/` | Local or CI-generated packages need a staging location | Keep it ignored and publish through GitHub Releases; track it only under an explicit provenance-enforced release contract |

Directory names are not style variants. Do not introduce an alternative because it looks more familiar.

<a id="directory-rules"></a>
## 🚦 Directory rules

1. Give each durable artifact one authoritative home.
2. Create a subdirectory only with its first real file or an instructional README.
3. Do not use `.gitkeep`, empty decorative folders, or parallel legacy/canonical locations.
4. Keep generated, cache, test-output, and distribution directories ignored unless an explicit repository contract says otherwise.
5. If a profile exception is necessary, document the reason, owner, and migration constraint in the root README.
6. Update `INSTALLATION.md`, `CONTRIBUTING.md`, static checks, and release tooling together when a path becomes contractual.
7. Preserve published links and package contracts deliberately; do not retain unexplained history forever.

## ✅ Generated-repository acceptance gate

A repository created from this template passes the structure gate when:

- all six canonical directories exist and contain either project material or their instructional README;
- the selected profile's `vba_contract` resolves at least one registered and tracked public façade, internal core, and test component;
- every path named by `vba_contract.required_components` exists with its declared role;
- every production component has one documented source location;
- public modules, internal/core modules, classes, forms, and tests are distinguishable;
- `INSTALLATION.md` lists the exact production manifest and import order;
- tests and examples do not share an ambiguous home;
- no pair of `tests/` and `test/`, `examples/` and `demo/`, or `assets/` and `images/` serves the same purpose;
- generated packages and transient evidence are not mistaken for source; and
- every profile-specific exception is explicit and justified.

An instructional README can preserve a purposeful directory, but it cannot
stand in for any mandatory VBA role. Classes, forms, Ribbon XML, workbook
modules, UI components, and examples are optional unless the selected profile
contract names them; the common baseline does not invent unused UI assets.

## 📚 Directory guides

- [`../src/README.md`](../src/README.md)
- [`../tests/README.md`](../tests/README.md)
- [`../examples/README.md`](../examples/README.md)
- [`../assets/README.md`](../assets/README.md)
- [`README.md`](README.md)
- [`../tools/README.md`](../tools/README.md)

---

**Structure principle:** create a location because it owns real responsibility, and keep exactly one location for that responsibility.
