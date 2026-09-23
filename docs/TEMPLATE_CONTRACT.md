# 🧬 Template Contract Versioning

[![Contract: 1.2.0](https://img.shields.io/badge/contract-1.2.0-1D76DB)](#version-history)
[![Versioning: SemVer](https://img.shields.io/badge/versioning-SemVer-6f42c1)](https://semver.org/spec/v2.0.0.html)
[![Independent of VERSION](https://img.shields.io/badge/independent%20of-project%20VERSION-217346)](#independence-from-the-project-version)

This document is the authority for the **template contract version**: the
versioned identity of the set of controls a generated repository adopts.

It is not the version of your project. A project's own product version lives in
[`../VERSION`](../VERSION) and its history in [`../CHANGELOG.md`](../CHANGELOG.md).

---

## 1. 🧭 What the contract version identifies

The template contract is the set of controls a generated repository is expected
to carry: required files and directories, the profile model, the placeholder
schema, label policy, the quality gates that must pass, and the release-evidence
rules.

The contract version identifies **that set**, so a repository can state which
baseline it adopted and tooling can evaluate it against the correct rules rather
than against whatever the template happens to look like today.

It is recorded in `.github/repository-profile.json`:

```json
"template_contract": {
  "version": "1.2.0",
  "source": "owner/template-repository"
}
```

`version` is the adopted contract. `source` is the template repository that
published it, and it is preserved verbatim through initialization — a generated
project's own `repository` field changes, `template_contract.source` does not.

<a id="semver-policy"></a>

## 2. 📏 SemVer policy

The contract follows [Semantic Versioning 2.0.0](https://semver.org/spec/v2.0.0.html)
with `MAJOR.MINOR.PATCH` core versions only. Pre-release and build metadata are
not used: an adopter either holds a published contract or does not.

| Change | Level | Meaning for an adopter |
| --- | --- | --- |
| A required control is removed, or its meaning changes so that a previously conforming repository stops conforming | **MAJOR** | Migration required before the repository conforms again |
| A required control is added, or an existing one is strengthened in a way a conforming repository can satisfy without redesign | **MINOR** | Adoption expected; migration notes state the exact steps |
| Wording, documentation, evidence formatting or an internal fix with no change to any required control | **PATCH** | No adopter action |

Two rules keep the level honest:

- classify by **effect on an adopting repository**, not by the size of the
  template diff;
- a control that becomes blocking where it was previously advisory is **MINOR**
  at least, never PATCH.

<a id="independence-from-the-project-version"></a>

## 3. 🔀 Independence from the project version

The contract version and a project's `VERSION` are deliberately unrelated:

- a template release that changes no required control ships **without** a
  contract bump, so a v1.3.0 template can still publish contract `1.2.0`;
- a generated project may release any number of its own versions while its
  adopted contract stays fixed;
- changing a project's `VERSION` never rewrites `template_contract`;
- adopting a new contract is a deliberate, reviewed change, not a side effect of
  updating the template.

The template's own release history and the contract history therefore diverge
over time. They currently share numbers only because every release so far has
changed the required control set.

<a id="version-history"></a>

## 4. 🗂️ Version history and migration notes

Each entry classifies its changes as **breaking**, **required**, **optional** or
**not applicable**, so an adopter can decide what to do.

### 1.2.0 — current

| Change | Class | Adopter action |
| --- | --- | --- |
| `template_contract` is recorded in `.github/repository-profile.json` | Required | Add the object with the adopted version and the template source |
| Deterministic documentation drift checks | Required | Adopt `.github/documentation-policy.json`, the documented-command/reference gate and its fixtures. This extends the pending 1.2.0 contract. [External-link observations](DOCUMENTATION_CHECKS.md) run separately on schedule or dispatch and are not a required PR check. |
| Profile-aware release provenance | Required | Adopt `.github/release-provenance.json`, `tools/release_provenance.py`, its fixtures and [RELEASE_PROVENANCE.md](RELEASE_PROVENANCE.md). Binary releases need build records; all releases enforce complete `dist/` inventory. Source-only releases need no artificial assets. SSH signatures are optional until enabled in committed policy. This extends the pending 1.2.0 contract; published rule sets remain unchanged. |
| Controlled dependency updates and `docs/DEPENDENCY_UPDATES.md` | Required | Adopt the policy and PR evidence block; record provenance and rollback, run applicable fixtures, and keep approval/merge manual. This is a minor-level addition to the pending 1.2.0 contract; earlier rule sets are unchanged. |
| `docs/TEMPLATE_CONTRACT.md` is a required path | Required | Adopt this document so the recorded baseline is explainable |
| `tools/check_template_contract.py` validates the recorded contract | Required | Run it locally and in CI alongside the canonical checker |
| Focused-gate CLI orchestration consolidated into `tools/_gatelib.py` | Not applicable | Internal to the template's tooling; no public CLI changed |
| Static repository checks expose reusable interface v1 | Optional | Follow [REUSABLE_WORKFLOWS.md](REUSABLE_WORKFLOWS.md) to adopt an exact workflow commit; copied workflows remain supported |
| Windows/Excel host evidence interface | Optional execution | The template retains the policy, validator, fixtures and [EXCEL_EVIDENCE.md](EXCEL_EVIDENCE.md). Use a reviewed eligible host or the same-schema manual fallback; unavailable never means PASS. Existing base compile/regression evidence stays required, and no Windows runner becomes a universal requirement. |
| Generated workflow identity checks recognize immutable references to the recorded template source | Optional | Update the canonical checker before adopting a reusable workflow; no broad identity-path exclusion is needed |

Reusable interface v1 adds no required control and does not change the resolved
rule sets above. Its compatibility policy is owned by
[REUSABLE_WORKFLOWS.md](REUSABLE_WORKFLOWS.md). A future breaking interface change
must introduce a new interface major and a new template-contract version with a
corresponding `CONTRACT_RULE_SETS` entry and migration note; do not redefine a
published contract or move a published pin.

### 1.1.0

| Change | Class | Adopter action |
| --- | --- | --- |
| Committed-whitespace, procedure-scoped VBA jump, nested conditional-compilation, complete public-API, repository-local Action and strict release-semantics gates became required | Required | Adopt the gates and wire them into the terminal CI verdict |
| `docs/PUBLIC_API.txt` required for every profile, with explicit visibility enforced | Required | Declare visibility explicitly and maintain the manifest |
| Read-only label-drift detection separated from trusted reconciliation | Required | Adopt the drift workflow; keep reconciliation on the default branch |
| Documentation authority map introduced in `docs/README.md` | Optional | Recommended for repositories with substantial documentation |

### 1.0.0

| Change | Class | Adopter action |
| --- | --- | --- |
| Initial certified baseline: profile model, placeholder schema, deterministic initializer, canonical repository gate, label policy, release-integrity gate and live governance | Required | This is the floor; earlier states are not contracts |

---

## 5. ✅ Validation

`tools/check_template_contract.py` is the focused gate for this contract. It:

- requires exactly `version` and `source`, both well formed;
- rejects a version outside the supported set with a message naming the
  supported versions, rather than failing obscurely;
- resolves the rule set registered for the recorded version, so conformance
  tooling evaluates a repository against the contract it adopted;
- requires every supported version to be documented here with migration notes;
- asserts that the recorded contract is independent of `VERSION`.

Run it directly:

```bash
python3 tools/check_template_contract.py --root .
python3 tools/check_template_contract.py --root . --self-test
```

Unknown or unsupported versions fail with an actionable message naming the
recorded value and the supported set.

---

## 6. 🔗 Related contracts

| Concern | Authority |
| --- | --- |
| Project product version | [`../VERSION`](../VERSION) |
| User-visible project history | [`../CHANGELOG.md`](../CHANGELOG.md) |
| SemVer and changelog ordering for releases | [`RELEASE_SEMANTICS.md`](RELEASE_SEMANTICS.md) |
| Template initialization and profile rendering | [`INITIALIZATION.md`](INITIALIZATION.md) |
| Live provisioning after generation | [`POST_CREATION_CHECKLIST.md`](POST_CREATION_CHECKLIST.md) |
