# 📚 Documentation Hub and Authority Map

[![Status: maintained](https://img.shields.io/badge/status-maintained-217346)](../README.md)
[![Authority: one contract](https://img.shields.io/badge/authority-one%20contract%20%2F%20one%20owner-success)](#authority-map)
[![Profiles: 3](https://img.shields.io/badge/profiles-3-6f42c1)](INITIALIZATION.md)
[![Evidence: exact SHA](https://img.shields.io/badge/evidence-exact%20SHA-1D76DB)](RELEASE_EVIDENCE.md)

`docs/` contains durable contracts that are too specialized for the root README.
The root README is an overview and navigation surface, not a second copy of the
operational policies below.

<a id="authority-map"></a>

## 🏛️ Authority map

Each evolving normative contract has one maintained authority. Other documents
may summarize it, but must link here rather than restating a competing rule.

| Contract | Authoritative location | Other documents may contain |
| --- | --- | --- |
| Project purpose, supported profiles, quick start and navigation | [`../README.md`](../README.md) | Short orientation only |
| Template initialization, token categories, profile rendering and manual fallback | [`INITIALIZATION.md`](INITIALIZATION.md) | A minimal initializer command |
| Repository/source layout and VBA component ownership | [`REPOSITORY_STRUCTURE.md`](REPOSITORY_STRUCTURE.md) | A small directory summary |
| VBA banners, procedure contracts, indentation and formatting review | [`VBA_HOUSE_STYLE.md`](VBA_HOUSE_STYLE.md) | A link to the source presentation rules |
| Installation, import, export, upgrade, recovery and removal | [`../INSTALLATION.md`](../INSTALLATION.md) | Links and release prerequisites |
| Contribution workflow, review discipline and PR evidence | [`../CONTRIBUTING.md`](../CONTRIBUTING.md) | Contributor links only |
| Participant conduct and enforcement | [`../CODE_OF_CONDUCT.md`](../CODE_OF_CONDUCT.md) | Conduct link only |
| Vulnerability scope, private reporting, disclosure and safe harbor | [`../SECURITY.md`](../SECURITY.md) | Private-reporting link and brief warning |
| Release sequence and maintainer decision points | [`../RELEASING.md`](../RELEASING.md) | Release navigation only |
| SemVer/changelog ordering and comparison-link semantics | [`RELEASE_SEMANTICS.md`](RELEASE_SEMANTICS.md) | Reference to the semantic gate |
| External release-evidence JSON, profile evidence and asset manifests | [`RELEASE_EVIDENCE.md`](RELEASE_EVIDENCE.md) | Reference to evidence requirements |
| Release build records, workflow identity and optional signatures | [`RELEASE_PROVENANCE.md`](RELEASE_PROVENANCE.md) | Verification commands and trust boundaries |
| Optional Windows/Excel execution interface and manual host evidence | [`EXCEL_EVIDENCE.md`](EXCEL_EVIDENCE.md) | Host prerequisites and validator commands |
| Maintained documentation drift and separate external-link observations | [`DOCUMENTATION_CHECKS.md`](DOCUMENTATION_CHECKS.md) | Local gate commands and network report interpretation |
| Adopted template contract version, its SemVer policy and migration notes | [`TEMPLATE_CONTRACT.md`](TEMPLATE_CONTRACT.md) | The recorded version; never a competing policy |
| Reusable workflow interface, immutable adoption, compatibility and rollback | [`REUSABLE_WORKFLOWS.md`](REUSABLE_WORKFLOWS.md) | A pinned caller example |
| Dependency monitoring, provenance review, update evidence and manual approval | [`DEPENDENCY_UPDATES.md`](DEPENDENCY_UPDATES.md) | A PR evidence block or link only |
| Private/public security-analysis eligibility and assurance gaps | [`SUPPLY_CHAIN_ASSURANCE.md`](SUPPLY_CHAIN_ASSURANCE.md) | Eligibility summary and result links |
| Post-publication provider checks, asset partition and retained closeout evidence | [`RELEASE_CLOSEOUT.md`](RELEASE_CLOSEOUT.md) | Closeout navigation only |
| Repository-local Action containment and tracked entry points | [`../tools/LOCAL_ACTIONS.md`](../tools/LOCAL_ACTIONS.md) | Link to the local Action gate contract |
| Migrated date-layer behavior (runtime parity verified on one Windows 64-bit host, #17) | [`DATE_LAYER_CONTRACT.md`](DATE_LAYER_CONTRACT.md) | Behavioral summaries only |
| VBA exchange format and character set | [`VBE_EXPORT.md`](VBE_EXPORT.md) | Link to format; procedures remain in Installation |
| Frozen migration source and attribution | [`MIGRATION_PROVENANCE.md`](MIGRATION_PROVENANCE.md) | Source reference and migration status |
| Migrated regression coverage, evidence boundary and source/destination parity procedure | [`MIGRATION_REGRESSION.md`](MIGRATION_REGRESSION.md) | Links to execution/parity requirements only |
| Source-to-destination work disposition and handover policy (historical) | [`MIGRATION_HANDOVER.md`](MIGRATION_HANDOVER.md) | Disposition only; acceptance criteria live in destination issues |
| v0.0.2 migration acceptance, limitations and next delivery scope | [`MIGRATION_COMPLETION.md`](MIGRATION_COMPLETION.md) | Link to the acceptance record |
| Forward delivery roadmap: milestones, boundaries, dependency sequence and scope transitions | [`ROADMAP.md`](ROADMAP.md) | Link to the roadmap; issues hold acceptance criteria |
| Public VBA surface | [`PUBLIC_API.txt`](PUBLIC_API.txt) | Human-readable API summary only |
| Live repository provisioning and verification after generation | [`POST_CREATION_CHECKLIST.md`](POST_CREATION_CHECKLIST.md) | One reminder that settings are not inherited |
| Historical initialization provenance and original inputs | [`INITIALIZATION_STATUS.md`](INITIALIZATION_STATUS.md) | Provenance link only |
| Dated live settings, visibility and control read-back | [`SETUP_VERIFICATION.md`](SETUP_VERIFICATION.md) | Link to the latest dated observation |
| Current version | [`../VERSION`](../VERSION) | Display/read the value; never redefine it |
| User-visible release history | [`../CHANGELOG.md`](../CHANGELOG.md) | Link or current Unreleased summary only |
| License terms | [`../LICENSE`](../LICENSE) | License name/link only |


The project-specific authorities are [developer setup](DEVELOPER_SETUP.md),
[dated setup verification](SETUP_VERIFICATION.md),
[the neutral Excel runbook](EXCEL_SETUP_RUNBOOK.md), and
[the delivery roadmap](ROADMAP.md) for current milestones and sequencing.
[The staged migration plan](MIGRATION_PLAN.md) is the historical v0.0.2
execution record.

## 🧭 First-use path

A new maintainer should normally need only this sequence:

1. Read the root [`README.md`](../README.md).
2. Follow [`DEVELOPER_SETUP.md`](DEVELOPER_SETUP.md); this repository is already initialized.
3. Import/compile/test using [`INSTALLATION.md`](../INSTALLATION.md).
4. Read the dated [`SETUP_VERIFICATION.md`](SETUP_VERIFICATION.md) and the
   [`EXCEL_SETUP_RUNBOOK.md`](EXCEL_SETUP_RUNBOOK.md) and
   accepted [`SETUP_COMPLETION.md`](SETUP_COMPLETION.md).
   Use [`ROADMAP.md`](ROADMAP.md) for current milestone status and sequencing.
5. For changes, use [`CONTRIBUTING.md`](../CONTRIBUTING.md).
6. For publication, use [`RELEASING.md`](../RELEASING.md) together with the
   release-semantics and release-evidence contracts.

Security reports always follow [`SECURITY.md`](../SECURITY.md), regardless of
where a problem is discovered.

## ✍️ Documentation rule

When a requirement changes, edit its authoritative document first. In every
secondary document:

- keep only the context needed for the reader's current task;
- link to the authority for exact rules or schemas;
- do not copy checklists, evidence schemas, support matrices or release rules;
- preserve historical evidence as historical evidence; and
- keep local links valid after every profile-specific initializer removal.

A Wiki may provide navigation or explanatory material, but versioned behavior
and governance remain in the repository.

---

**Documentation principle:** one maintained authority per contract, concise
navigation everywhere else, and no evidence claim broader than the source that
proves it.
