<div align="center">

# 🤝 Contributing to {{PROJECT_NAME}}

### {{PROJECT_DESCRIPTION}}

[![Contributions](https://img.shields.io/badge/Contributions-Welcome-2ea44f?style=flat-square)](#ways-to-contribute)
[![Conduct](https://img.shields.io/badge/Conduct-Required-6f42c1?style=flat-square)](CODE_OF_CONDUCT.md)
[![Security](https://img.shields.io/badge/Security-Private_reporting-d73a49?style=flat-square)](SECURITY.md)
[![Evidence](https://img.shields.io/badge/Evidence-Exact_source-0969da?style=flat-square)](#validation-and-evidence)

<br>

**Focused scope · Reviewable source · Reproducible evidence · Honest limitations**

</div>

---

This document is authoritative for the **contribution and review workflow**.
Repository layout is owned by
[`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md), vulnerability
handling by [`SECURITY.md`](SECURITY.md), installation by
[`INSTALLATION.md`](INSTALLATION.md), and publication by
[`RELEASING.md`](RELEASING.md). The complete documentation authority map is
[`docs/README.md`](docs/README.md).

Participation is governed by the [Code of Conduct](CODE_OF_CONDUCT.md).
Suspected vulnerabilities must never be disclosed in a public issue or pull
request.

<a id="ways-to-contribute"></a>

## 🌱 Ways to contribute

| Contribution | First action |
| --- | --- |
| 🐛 Reproducible defect | Use the bug form with exact source, environment, expected/observed behavior and evidence. |
| ✨ Feature or API change | Use the feature form to define users, observable contract, compatibility, non-goals and validation. |
| 🧪 Test/evidence improvement | Explain provenance, independence, coverage and what failure the test detects. |
| 📖 Documentation | Use the documentation form and identify the authoritative contract being corrected. |
| ⚙️ Tooling/governance | Explain failure behavior, portability, trust boundary and maintenance cost. |
| 🔐 Security concern | Follow [`SECURITY.md`](SECURITY.md) privately. |

Open an issue before a non-trivial public-API change, dependency, architectural
change, compatibility break or broad refactor. Small documentation corrections
and narrowly obvious fixes may go directly to a focused pull request.

## 🌿 Development workflow

1. Start from the current protected development baseline and create one focused
   branch unless the repository's maintainer workflow explicitly says otherwise.
2. Reproduce the current behavior before changing it.
3. Define the observable contract, affected callers, compatibility impact and
   validation plan.
4. Make the smallest coherent change; avoid unrelated formatting, generated
   output or opportunistic refactoring.
5. Keep exported VBA and repository policy synchronized with the actual tree.
6. Run the applicable local static, compile, regression and specialist checks.
7. Update the authoritative documentation and changelog surface affected by the
   change.
8. Review the complete diff, then open a pull request with evidence and explicit
   limitations.

Use imperative, specific commit subjects. Reference the issue when one exists.
Do not place credentials, private links, attribution boilerplate or unverifiable
test claims in commit messages.

## 📦 Source-change discipline

Dependency, validation-tool, runtime and reusable-workflow updates follow
[`docs/DEPENDENCY_UPDATES.md`](docs/DEPENDENCY_UPDATES.md). Complete the PR's
dependency evidence block and retain a tested rollback target. Merge manually
only after provenance, trust-boundary and candidate-evidence review.

The exact source/storage contract is maintained in
[`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md). VBA presentation
follows [`docs/VBA_HOUSE_STYLE.md`](docs/VBA_HOUSE_STYLE.md). For contribution
work, the practical rules are:

- exported VBA remains the reviewable source of truth;
- component filenames and VBE identities stay coherent;
- `.frm`/`.frx` pairs remain together and `.frx` stays binary;
- production, tests, examples and generated artifacts remain in their governed
  locations;
- public-surface changes update [`docs/PUBLIC_API.txt`](docs/PUBLIC_API.txt) and
  their regression coverage; and
- a workbook/add-in is never the only record of a source change.

Do not weaken a project-specific numerical, UI, lifecycle, performance or
packaging gate merely because the generic repository gate passes.

For Python comments, docstrings, wrapping and enforced lint rules, follow
[the tooling presentation policy](tools/README.md#python-presentation-and-lint-policy).
Apply the same accuracy review to workflow comments, configuration explanations
and embedded examples whenever their owning behavior changes.

## 🔄 Compatibility and state ownership

A change to documented procedures, functions, classes, enums, parameters,
defaults, return values, errors, side effects, file formats or supported
platforms is a contract change.

Such a contribution must identify callers and migration impact, update permanent
regression coverage, update the authoritative user documentation and state the
release impact.

Excel/Windows state belongs to the caller or host unless the component
explicitly owns it. Capture before mutation, restore only state successfully
changed and still owned, and never let cleanup conceal the original failure.

<a id="validation-and-evidence"></a>

## 🧪 Validation and evidence

Validation must be reproducible from the exact source under review. Record:

```text
Source
------
Commit / tag:
Files or components changed:

Environment
-----------
Excel / Office build:
Office bitness:
Operating system:
Relevant deployment / locale / date system:

Checks
------
Compile:
Repository/static checks:
Focused regression:
Full regression:
Specialist/manual checks:
Cleanup:

Limitations
-----------
Skipped or unverified:
Follow-up:
```

Use only applicable fields, but never omit a material limitation. Static source
inspection cannot substitute for Excel execution. A skipped check is not a pass.
Numerical or reference evidence must be independent of the implementation under
test.

<!-- template:remove:start -->
For checker changes, additionally follow
[`docs/CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md).
<!-- template:remove:end -->
For release-evidence schemas and exact-SHA binding, use
[`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md).

<a id="pull-requests"></a>

## 🚀 Pull requests

A pull request should answer five questions:

1. What problem does this solve?
2. What observable contract changes?
3. What remains compatible?
4. How was the exact source validated?
5. What remains unverified?

### Review checklist

- [ ] Scope is focused and the related issue is linked where applicable.
- [ ] Public API, compatibility and release impact are assessed.
- [ ] Exported VBA and required binary companions are synchronized.
- [ ] Relevant compile, static, regression and specialist checks are recorded.
- [ ] Error, boundary, recovery and cleanup paths are covered where affected.
- [ ] Caller-owned state and platform/bitness concerns are addressed.
- [ ] The authoritative documentation and changelog are updated.
- [ ] No confidential, restricted, accidental binary or generated material is added.
- [ ] Unverified environments and skipped checks are stated plainly.
- [ ] Final diff contains no unrelated formatting or local artifacts.

Reviewers may request changes to scope, tests, contracts, compatibility,
documentation or evidence. Discussion remains technical and respectful under the
[Code of Conduct](CODE_OF_CONDUCT.md).

## 📚 Where detailed rules live

| Need | Authority |
| --- | --- |
| Repository/source ownership | [`docs/REPOSITORY_STRUCTURE.md`](docs/REPOSITORY_STRUCTURE.md) |
| Installation and upgrade | [`INSTALLATION.md`](INSTALLATION.md) |
| Vulnerability reporting | [`SECURITY.md`](SECURITY.md) |
| Release procedure | [`RELEASING.md`](RELEASING.md) |
| SemVer/changelog semantics | [`docs/RELEASE_SEMANTICS.md`](docs/RELEASE_SEMANTICS.md) |
| Release evidence | [`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md) |
<!-- template:remove:start -->
| Checker changes | [`docs/CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md) |
<!-- template:remove:end -->
| Full authority map | [`docs/README.md`](docs/README.md) |

## 📄 Licensing and maintainer

This project is distributed under the [MIT License](LICENSE). Contributors must
have the right to submit every part of a contribution, including code, tests,
data, images and generated material.

Maintained by **{{MAINTAINER_NAME}}**.

---

**Contribution principle:** make the contract explicit, keep the diff focused,
and leave evidence another person can reproduce.
