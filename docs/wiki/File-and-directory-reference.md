# 🗂️ File and directory reference

> **Guide, not policy:** [Repository Structure](../REPOSITORY_STRUCTURE.md) and the linked authorities govern.

This inventory is generated from the reviewed catalogue and checked against `git ls-files`.
Each tracked file and each ancestor directory appears once. Group rules apply to every row
of that kind; the row supplies its specific purpose and authoritative document.

**Lifecycle order: application / library / UI component.** R = retained byte-for-byte;
T = transformed; X = removed. Values come from the actual initializer dry run with its
standard synthetic values. Supplying optional/repeatable values changes rendered text.
The optional `assets/social-preview.png` is retained when selected; otherwise removed.
An untracked directory is not a Git artifact. Directory retention follows its children.

## Editing and validation rules by kind

| Kind / owner | When to edit | Invariants and unsafe edits | Validation |
| --- | --- | --- | --- |
| asset / Documentation maintainer | When replacing maintained preview artwork. | Keep formats and referenced paths coherent; no confidential screenshots or unrelated binaries. | Canonical artifact/path checks and visual review. |
| directory / Role owner under Repository Structure | When adding a component of this directory role. | Keep production, examples, tests, metadata and outputs in their owned locations. | Canonical required-directory/role checks; file-specific gates below. |
| document / Document maintainer | When the supported contract or instructions change. | Keep one authority; preserve historical evidence and private reporting boundaries. | Canonical links/identity plus documentation drift; review prose. |
| git-policy / Repository maintainer | For reviewed editor, normalization, exclusion or archive policy changes. | Do not convert binary companions to text, normalize VBA incorrectly or commit local secrets. | Canonical dotfile/line-ending/whitespace policy probes. |
| policy / Repository maintainer | Only for a reviewed control or profile change. | Do not weaken rules to silence findings or hand-edit adopted evidence. | Canonical structured-data and the owning focused gate; hosted integrity. |
| script / Tool maintainer | For a scoped CLI/validation implementation change with failure fixtures. | Preserve exit/report contracts, boundaries and downstream compatibility; no embedded credentials. | Ruff/mypy, named tool fixtures and owning hosted workflow; Tools README. |
| vba / VBA component maintainer | For reviewed behavior or source-presentation changes. | Preserve export identity/CRLF, component visibility, API and cleanup; annotate variables and phases. | Canonical and focused VBA gates, generated profiles and actual Excel regression. |
| wiki / Template wiki maintainer | When the guided journey or a linked contract changes. | Keep authority notices, ordered navigation and exact-source publication; no independent online policy edits. | Wiki source/fixtures, canonical links, documentation drift and separate published comparison. |
| workflow / Automation maintainer | When triggers, permissions, pins or job contracts intentionally change. | Use reviewed immutable action pins, least permissions and honest failure handling. | Canonical action/YAML checks, pinned actionlint and workflow fixtures; see workflow reference. |

## Tracked files

| Path / kind | Purpose | A / L / UI | Authority |
| --- | --- | --- | --- |
| `.editorconfig` / git-policy | Editor defaults for indentation, final newlines and per-type line endings. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `.gitattributes` / git-policy | Git normalization, binary treatment, language statistics and archive exclusions. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `.github/ISSUE_TEMPLATE/bug.yml` / policy | Structured reproducible-defect intake with environment and evidence. | R / R / R | [docs/POST_CREATION_CHECKLIST.md](../../docs/POST_CREATION_CHECKLIST.md) |
| `.github/ISSUE_TEMPLATE/config.yml` / policy | Disable blank issues and route vulnerabilities to the project security policy. | T / T / T | [docs/POST_CREATION_CHECKLIST.md](../../docs/POST_CREATION_CHECKLIST.md) |
| `.github/ISSUE_TEMPLATE/documentation.yml` / policy | Documentation defect intake identifying the authority and proposed correction. | R / R / R | [docs/DOCUMENTATION_CHECKS.md](../../docs/DOCUMENTATION_CHECKS.md) |
| `.github/ISSUE_TEMPLATE/feature.yml` / policy | Feature/change intake with scope, compatibility and acceptance criteria. | R / R / R | [docs/POST_CREATION_CHECKLIST.md](../../docs/POST_CREATION_CHECKLIST.md) |
| `.github/PULL_REQUEST_TEMPLATE.md` / document | PR review/evidence checklist, including dependency-update information. | R / R / R | [CONTRIBUTING.md](../../CONTRIBUTING.md) |
| `.github/dependabot.yml` / policy | Weekly proposal-only GitHub Actions dependency discovery. | R / R / R | [docs/DEPENDENCY_UPDATES.md](../../docs/DEPENDENCY_UPDATES.md) |
| `.github/documentation-policy.json` / policy | Registered documentation references, HTTP bounds, allowed domains and expiring exceptions. | R / R / R | [docs/DOCUMENTATION_CHECKS.md](../../docs/DOCUMENTATION_CHECKS.md) |
| `.github/excel-evidence-policy.json` / policy | Starter entry point, case registry, assertion count and expected-error case. | R / R / R | [docs/EXCEL_EVIDENCE.md](../../docs/EXCEL_EVIDENCE.md) |
| `.github/labels.json` / policy | Canonical core/profile/domain label definitions and reconciliation policy. | R / R / R | [docs/POST_CREATION_CHECKLIST.md](../../docs/POST_CREATION_CHECKLIST.md) |
| `.github/provisioning-policy.json` / policy | Reviewed metadata/features/merge/required-check goals and exception rationale. | R / R / R | [docs/PROVISIONING.md](../../docs/PROVISIONING.md) |
| `.github/release-history-policy.json` / policy | Canonical-template merge-history rules, exact-SHA reviewed exceptions and preserved historical records. | X / X / X | [docs/RELEASE_SEMANTICS.md](../../docs/RELEASE_SEMANTICS.md) |
| `.github/release-policy.json` / policy | Required candidate checks, allowed profile assets and construction-history exclusions. | R / R / R | [docs/RELEASE_EVIDENCE.md](../../docs/RELEASE_EVIDENCE.md) |
| `.github/release-provenance.json` / policy | Expected source/workflow identity and optional committed SSH signing trust. | R / R / R | [docs/RELEASE_PROVENANCE.md](../../docs/RELEASE_PROVENANCE.md) |
| `.github/repository-profile.json` / policy | Mode, profile, identity, roles, paths, placeholder schema and adopted control version. | T / T / T | [docs/INITIALIZATION.md](../../docs/INITIALIZATION.md) |
| `.github/scripts/labels-drift.mjs` / script | Read-only label diff with planner cross-check and offline self-tests. | R / R / R | [docs/POST_CREATION_CHECKLIST.md](../../docs/POST_CREATION_CHECKLIST.md) |
| `.github/scripts/labels-sync.mjs` / script | Validate label schema, resolve overlays, reconcile on trusted events and verify result. | R / R / R | [docs/POST_CREATION_CHECKLIST.md](../../docs/POST_CREATION_CHECKLIST.md) |
| `.github/workflows/checker-development.yml` / workflow | Template-only independent checker contract and coverage execution. | X / X / X | [.github/workflows/checker-development.yml](../../.github/workflows/checker-development.yml) |
| `.github/workflows/codeql.yml` / workflow | Trusted-context CodeQL analysis for maintained Python and JavaScript source. | R / R / R | [docs/SUPPLY_CHAIN_ASSURANCE.md](../../docs/SUPPLY_CHAIN_ASSURANCE.md) |
| `.github/workflows/external-links.yml` / workflow | Separate weekly/manual anonymous external-link observation with retained reports. | R / R / R | [.github/workflows/external-links.yml](../../.github/workflows/external-links.yml) |
| `.github/workflows/labels-drift.yml` / workflow | PR label fixtures and separate daily/manual read-only live drift detection. | R / R / R | [.github/workflows/labels-drift.yml](../../.github/workflows/labels-drift.yml) |
| `.github/workflows/labels-sync.yml` / workflow | Read-only PR validation and trusted main/manual label reconciliation. | R / R / R | [.github/workflows/labels-sync.yml](../../.github/workflows/labels-sync.yml) |
| `.github/workflows/maintenance.yml` / workflow | Template-only dispatch-only home for registered maintenance tasks that commit their result. | X / X / X | [docs/CHECKER_DEVELOPMENT.md](../../docs/CHECKER_DEVELOPMENT.md) |
| `.github/workflows/portfolio-drift.yml` / workflow | Template-only synthetic drift, quality and provisioner acceptance fixtures. | X / X / X | [.github/workflows/portfolio-drift.yml](../../.github/workflows/portfolio-drift.yml) |
| `.github/workflows/release-closeout.yml` / workflow | Manual read-only capture and deterministic post-release closeout evidence. | R / R / R | [docs/RELEASE_CLOSEOUT.md](../../docs/RELEASE_CLOSEOUT.md) |
| `.github/workflows/scorecard.yml` / workflow | Trusted-context OpenSSF Scorecard publication and retained SARIF evidence. | R / R / R | [docs/SUPPLY_CHAIN_ASSURANCE.md](../../docs/SUPPLY_CHAIN_ASSURANCE.md) |
| `.github/workflows/static-checks.yml` / workflow | Canonical/focused gates, quality tools, self-tests, artifacts and reusable interface. | R / R / R | [.github/workflows/static-checks.yml](../../.github/workflows/static-checks.yml) |
| `.github/workflows/wiki-checks.yml` / workflow | Template-only offline wiki checks, publication fixtures and exact-source export artifact. | X / X / X | [.github/workflows/wiki-checks.yml](../../.github/workflows/wiki-checks.yml) |
| `.github/workflows/wiki-drift.yml` / workflow | Separate weekly/manual anonymous wiki fetch and deterministic publication comparison. | X / X / X | [.github/workflows/wiki-drift.yml](../../.github/workflows/wiki-drift.yml) |
| `.gitignore` / git-policy | Exclude transient reports, editor state, credentials, Office locks and build outputs. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `CHANGELOG.md` / document | User-visible release history, version comparison links and Unreleased changes. | T / T / T | [docs/RELEASE_SEMANTICS.md](../../docs/RELEASE_SEMANTICS.md) |
| `CODE_OF_CONDUCT.md` / document | Participant expectations and private enforcement route. | T / T / T | [CODE_OF_CONDUCT.md](../../CODE_OF_CONDUCT.md) |
| `CONTRIBUTING.md` / document | Change scope, review discipline, source ownership and validation evidence. | T / T / T | [CONTRIBUTING.md](../../CONTRIBUTING.md) |
| `INSTALLATION.md` / document | Import, upgrade, recovery and removal procedure for exported VBA. | T / T / T | [INSTALLATION.md](../../INSTALLATION.md) |
| `LICENSE` / document | MIT reuse terms and copyright attribution. | T / T / T | [LICENSE](../../LICENSE) |
| `README.md` / document | Project identity, profile overview, quick start and main navigation. | T / T / T | [README.md](../../README.md) |
| `RELEASING.md` / document | Ordered release preparation, certification, tagging and publication procedure. | T / T / T | [RELEASING.md](../../RELEASING.md) |
| `SECURITY.md` / document | Vulnerability scope, private reporting, disclosure and response. | T / T / T | [SECURITY.md](../../SECURITY.md) |
| `VERSION` / policy | Single product-version value; initializer resets it to the 0.0.0 sentinel. | T / T / T | [docs/RELEASE_SEMANTICS.md](../../docs/RELEASE_SEMANTICS.md) |
| `assets/README.md` / document | Rules for maintained visual/support assets. | T / T / T | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `assets/social-preview.png` / asset | Optional raster social preview removed unless explicitly selected. | X / X / X | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `assets/social-preview.svg` / asset | Editable template preview source; removed during generation. | X / X / X | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `docs/CHECKER_DEVELOPMENT.md` / document | Independent checker-development and policy-coverage requirements. | X / X / X | [docs/CHECKER_DEVELOPMENT.md](../../docs/CHECKER_DEVELOPMENT.md) |
| `docs/DEPENDENCY_UPDATES.md` / document | Manual dependency monitoring, provenance review and rollback. | T / T / T | [docs/DEPENDENCY_UPDATES.md](../../docs/DEPENDENCY_UPDATES.md) |
| `docs/DOCUMENTATION_CHECKS.md` / document | Offline command/reference validation and separate external-link policy. | R / R / R | [docs/DOCUMENTATION_CHECKS.md](../../docs/DOCUMENTATION_CHECKS.md) |
| `docs/EXCEL_EVIDENCE.md` / document | Optional Windows/Excel runner interface and manual retained-log schema. | R / R / R | [docs/EXCEL_EVIDENCE.md](../../docs/EXCEL_EVIDENCE.md) |
| `docs/INITIALIZATION.md` / document | Authoritative initialization inputs, profile rendering and no-op contract. | R / R / R | [docs/INITIALIZATION.md](../../docs/INITIALIZATION.md) |
| `docs/PILOT_CERTIFICATION.md` / document | Historical template pilot/certification evidence; not a fresh acceptance run. | X / X / X | [docs/PILOT_CERTIFICATION.md](../../docs/PILOT_CERTIFICATION.md) |
| `docs/PORTFOLIO_DRIFT.md` / document | Read-only snapshot and structural adoption/conformance comparison contract. | X / X / X | [docs/PORTFOLIO_DRIFT.md](../../docs/PORTFOLIO_DRIFT.md) |
| `docs/PORTFOLIO_QUALITY.md` / document | Timestamped quality observations, freshness and missing-evidence interpretation; no overall score. | X / X / X | [docs/PORTFOLIO_QUALITY.md](../../docs/PORTFOLIO_QUALITY.md) |
| `docs/POST_CREATION_CHECKLIST.md` / document | Live settings and read-back controls that template creation does not transfer. | R / R / R | [docs/POST_CREATION_CHECKLIST.md](../../docs/POST_CREATION_CHECKLIST.md) |
| `docs/PROVISIONING.md` / document | Guarded setup planner, approved-digest apply and durable journal contract. | X / X / X | [docs/PROVISIONING.md](../../docs/PROVISIONING.md) |
| `docs/PUBLIC_API.txt` / document | Complete public declaration inventory with normalized signature records. | R / R / R | [docs/PUBLIC_API.txt](../../docs/PUBLIC_API.txt) |
| `docs/README.md` / document | Documentation hub and single-authority mapping. | T / T / T | [docs/README.md](../../docs/README.md) |
| `docs/RELEASE_CLOSEOUT.md` / document | Canonical post-publication release closeout procedure and evidence boundaries. | R / R / R | [docs/RELEASE_CLOSEOUT.md](../../docs/RELEASE_CLOSEOUT.md) |
| `docs/RELEASE_EVIDENCE.md` / document | Candidate-bound release JSON and profile/asset evidence schemas. | R / R / R | [docs/RELEASE_EVIDENCE.md](../../docs/RELEASE_EVIDENCE.md) |
| `docs/RELEASE_PROVENANCE.md` / document | Complete payload inventory, build records and optional SSH trust/signatures. | R / R / R | [docs/RELEASE_PROVENANCE.md](../../docs/RELEASE_PROVENANCE.md) |
| `docs/RELEASE_SEMANTICS.md` / document | Strict SemVer, changelog ordering and comparison-link rules. | R / R / R | [docs/RELEASE_SEMANTICS.md](../../docs/RELEASE_SEMANTICS.md) |
| `docs/REPOSITORY_STRUCTURE.md` / document | Source layout, component roles, exports and distribution ownership. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `docs/REUSABLE_WORKFLOWS.md` / document | Immutable workflow interface pin, caller contract and adoption procedure. | T / T / T | [docs/REUSABLE_WORKFLOWS.md](../../docs/REUSABLE_WORKFLOWS.md) |
| `docs/SUPPLY_CHAIN_ASSURANCE.md` / document | Canonical CodeQL, Dependabot and Scorecard permission/evidence boundaries. | R / R / R | [docs/SUPPLY_CHAIN_ASSURANCE.md](../../docs/SUPPLY_CHAIN_ASSURANCE.md) |
| `docs/TEMPLATE_CONTRACT.md` / document | Independent adopted control versions and migration classification. | R / R / R | [docs/TEMPLATE_CONTRACT.md](../../docs/TEMPLATE_CONTRACT.md) |
| `docs/VBA_HOUSE_STYLE.md` / document | Procedure contracts, declaration annotations and explanatory body comments. | R / R / R | [docs/VBA_HOUSE_STYLE.md](../../docs/VBA_HOUSE_STYLE.md) |
| `docs/WIKI_PUBLICATION.md` / document | Reviewed wiki source, complete inventory, exact-SHA publication and drift. | X / X / X | [docs/WIKI_PUBLICATION.md](../../docs/WIKI_PUBLICATION.md) |
| `docs/wiki/Choose-a-profile.md` / wiki | Guided choose a profile page; summarizes and links to the owning repository contracts. | X / X / X | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `docs/wiki/Configure-GitHub.md` / wiki | Guided configure github page; summarizes and links to the owning repository contracts. | X / X / X | [docs/POST_CREATION_CHECKLIST.md](../../docs/POST_CREATION_CHECKLIST.md) |
| `docs/wiki/Create-the-repository.md` / wiki | Guided create the repository page; summarizes and links to the owning repository contracts. | X / X / X | [docs/INITIALIZATION.md](../../docs/INITIALIZATION.md) |
| `docs/wiki/Develop-safely.md` / wiki | Guided develop safely page; summarizes and links to the owning repository contracts. | X / X / X | [CONTRIBUTING.md](../../CONTRIBUTING.md) |
| `docs/wiki/File-and-directory-reference.md` / wiki | Generated exact tracked-file/directory inventory with per-profile initializer lifecycle. | X / X / X | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `docs/wiki/Glossary.md` / wiki | Guided glossary page; summarizes and links to the owning repository contracts. | X / X / X | [docs/README.md](../../docs/README.md) |
| `docs/wiki/Home.md` / wiki | Guided start here page; summarizes and links to the owning repository contracts. | X / X / X | [docs/README.md](../../docs/README.md) |
| `docs/wiki/Initialize-the-project.md` / wiki | Guided initialize the project page; summarizes and links to the owning repository contracts. | X / X / X | [docs/INITIALIZATION.md](../../docs/INITIALIZATION.md) |
| `docs/wiki/Maintenance-and-migration.md` / wiki | Guided maintenance and migration page; summarizes and links to the owning repository contracts. | X / X / X | [docs/TEMPLATE_CONTRACT.md](../../docs/TEMPLATE_CONTRACT.md) |
| `docs/wiki/Prepare-and-publish-a-release.md` / wiki | Guided prepare and publish a release page; summarizes and links to the owning repository contracts. | X / X / X | [RELEASING.md](../../RELEASING.md) |
| `docs/wiki/Publish-the-wiki.md` / wiki | Guided publish the wiki page; summarizes and links to the owning repository contracts. | X / X / X | [docs/WIKI_PUBLICATION.md](../../docs/WIKI_PUBLICATION.md) |
| `docs/wiki/Run-the-quality-gates.md` / wiki | Guided run the quality gates page; summarizes and links to the owning repository contracts. | X / X / X | [tools/README.md](../../tools/README.md) |
| `docs/wiki/Troubleshooting.md` / wiki | Guided troubleshooting page; summarizes and links to the owning repository contracts. | X / X / X | [docs/README.md](../../docs/README.md) |
| `docs/wiki/Workflow-and-tool-reference.md` / wiki | Guided workflow and tool reference page; summarizes and links to the owning repository contracts. | X / X / X | [tools/README.md](../../tools/README.md) |
| `docs/wiki/_Sidebar.md` / wiki | Generated ordered Wiki navigation; edit the page registry rather than this output. | X / X / X | [docs/WIKI_PUBLICATION.md](../../docs/WIKI_PUBLICATION.md) |
| `docs/wiki/catalogue.json` / wiki | Ordered page registry, per-path purpose/authority and shared inventory editing rules. | X / X / X | [docs/WIKI_PUBLICATION.md](../../docs/WIKI_PUBLICATION.md) |
| `examples/README.md` / document | Ownership and scope of reproducible usage examples. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `examples/modules/ProjectExample.bas` / vba | Optional project-private example that prints one supported facade result. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `pyproject.toml` / policy | Python compatibility target, advisory line length, enforced lint selection, complexity ceiling and typing scope. | R / R / R | [tools/README.md](../../tools/README.md) |
| `src/README.md` / document | Production export organization and public/internal boundaries. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `src/core/ProjectCore.bas` / vba | Stateless checked division and the internal zero-denominator error code. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `src/modules/ProjectFacade.bas` / vba | Supported scalar ratio entry point and normalized caller-facing error source. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `tests/README.md` / document | Regression entry point, output contract, repeatability and cleanup expectations. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `tests/modules/ProjectTests.bas` / vba | Four-case, six-assertion harness with host-state comparison and full reporting. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `tools/LOCAL_ACTIONS.md` / document | Repository-local composite/Node/Docker action validation boundaries. | R / R / R | [tools/LOCAL_ACTIONS.md](../../tools/LOCAL_ACTIONS.md) |
| `tools/README.md` / document | Operational tool commands, reports and focused-gate responsibilities. | T / T / T | [tools/README.md](../../tools/README.md) |
| `tools/_gatelib.py` / script | Shared Git, tracked-file, report and focused CLI orchestration primitives; imported helper. | R / R / R | [tools/README.md](../../tools/README.md) |
| `tools/_release_closeout.py` / script | Private deterministic validator for retained post-release provider snapshots. | R / R / R | [docs/RELEASE_CLOSEOUT.md](../../docs/RELEASE_CLOSEOUT.md) |
| `tools/check_committed_whitespace.py` / script | Validate staged/unstaged feedback or an explicit committed Git range. | R / R / R | [tools/README.md](../../tools/README.md) |
| `tools/check_documentation.py` / script | Parse literal documented Python commands and registered contract references. | R / R / R | [docs/DOCUMENTATION_CHECKS.md](../../docs/DOCUMENTATION_CHECKS.md) |
| `tools/check_excel_evidence.py` / script | Validate exact-source host records, stage outcomes and retained-log hashes. | R / R / R | [docs/EXCEL_EVIDENCE.md](../../docs/EXCEL_EVIDENCE.md) |
| `tools/check_external_links.py` / script | Bounded anonymous HTTPS observations with categorized access and transient failures. | R / R / R | [docs/DOCUMENTATION_CHECKS.md](../../docs/DOCUMENTATION_CHECKS.md) |
| `tools/check_local_actions.py` / script | Resolve local action metadata, implementation files and supported invocation forms. | R / R / R | [tools/LOCAL_ACTIONS.md](../../tools/LOCAL_ACTIONS.md) |
| `tools/check_policy_coverage.py` / script | Independent policy-branch coverage entry point for the canonical checker. | X / X / X | [docs/CHECKER_DEVELOPMENT.md](../../docs/CHECKER_DEVELOPMENT.md) |
| `tools/check_portfolio_drift.py` / script | Compare captured repository structure with the explicitly adopted control set. | X / X / X | [docs/PORTFOLIO_DRIFT.md](../../docs/PORTFOLIO_DRIFT.md) |
| `tools/check_release.py` / script | Validate an actual release candidate/evidence/assets/tag or synthetic self-tests. | R / R / R | [docs/RELEASE_EVIDENCE.md](../../docs/RELEASE_EVIDENCE.md) |
| `tools/check_release_semantics.py` / script | Enforce strict version, dated changelog, comparison and canonical-template history semantics. | R / R / R | [docs/RELEASE_SEMANTICS.md](../../docs/RELEASE_SEMANTICS.md) |
| `tools/check_repo.py` / script | Self-contained dependency-free canonical 21-rule repository-quality gate. | R / R / R | [tools/README.md](../../tools/README.md) |
| `tools/check_template_contract.py` / script | Resolve adopted contract versions and reject invalid or unsupported identities. | R / R / R | [docs/TEMPLATE_CONTRACT.md](../../docs/TEMPLATE_CONTRACT.md) |
| `tools/check_vba_conditionals.py` / script | Validate conditional compilation and reachable PtrSafe declarations in three host models. | R / R / R | [tools/README.md](../../tools/README.md) |
| `tools/check_vba_jumps.py` / script | Resolve jumps and labels within their owning VBA procedures. | R / R / R | [tools/README.md](../../tools/README.md) |
| `tools/check_vba_public_api.py` / script | Compare complete public VBA declarations and normalized manifest signatures. | R / R / R | [tools/README.md](../../tools/README.md) |
| `tools/check_wiki.py` / script | Validate complete wiki inventory/navigation, export pinned pages or compare a fetched publication. | X / X / X | [docs/WIKI_PUBLICATION.md](../../docs/WIKI_PUBLICATION.md) |
| `tools/checker_development.py` / script | Independent canonical-checker architecture, self-test and shared-runner contract. | X / X / X | [docs/CHECKER_DEVELOPMENT.md](../../docs/CHECKER_DEVELOPMENT.md) |
| `tools/collect_portfolio_snapshot.py` / script | GET-only capture of explicit repositories and source/live observations. | X / X / X | [docs/PORTFOLIO_DRIFT.md](../../docs/PORTFOLIO_DRIFT.md) |
| `tools/create_reusable_workflow_fixture.py` / script | Create a disposable initialized caller pinned to an explicit provider commit. | X / X / X | [docs/REUSABLE_WORKFLOWS.md](../../docs/REUSABLE_WORKFLOWS.md) |
| `tools/dev_check.sh` / script | Run the locally reproducible part of hosted CI with the pinned tool versions. | X / X / X | [docs/CHECKER_DEVELOPMENT.md](../../docs/CHECKER_DEVELOPMENT.md) |
| `tools/initialize_repository.py` / script | Validate values, preview transformations, apply once and verify identical-input no-op. | R / R / R | [docs/INITIALIZATION.md](../../docs/INITIALIZATION.md) |
| `tools/policy_coverage_cases_config.py` / script | Synthetic configuration and policy-validation branch cases; imported helper. | X / X / X | [docs/CHECKER_DEVELOPMENT.md](../../docs/CHECKER_DEVELOPMENT.md) |
| `tools/policy_coverage_cases_quality.py` / script | Synthetic content/source quality branch cases; imported helper. | X / X / X | [docs/CHECKER_DEVELOPMENT.md](../../docs/CHECKER_DEVELOPMENT.md) |
| `tools/policy_coverage_cases_repo.py` / script | Synthetic repository/Git branch cases; imported helper. | X / X / X | [docs/CHECKER_DEVELOPMENT.md](../../docs/CHECKER_DEVELOPMENT.md) |
| `tools/policy_coverage_core.py` / script | Coverage fixture data model and shared branch-observation primitives; imported helper. | X / X / X | [docs/CHECKER_DEVELOPMENT.md](../../docs/CHECKER_DEVELOPMENT.md) |
| `tools/policy_coverage_runner.py` / script | Aggregate semantic branch observations into reproducible policy-coverage reports. | X / X / X | [docs/CHECKER_DEVELOPMENT.md](../../docs/CHECKER_DEVELOPMENT.md) |
| `tools/provision_repository.py` / script | Plan live new-repository settings or apply a reviewed digest with a durable journal. | X / X / X | [docs/PROVISIONING.md](../../docs/PROVISIONING.md) |
| `tools/release_certification.py` / script | Build or verify deterministic, candidate-bound release certification bundles. | R / R / R | [docs/RELEASE_EVIDENCE.md](../../docs/RELEASE_EVIDENCE.md) |
| `tools/release_provenance.py` / script | Release helper for complete payloads, build identity and optional SSH verification. | R / R / R | [docs/RELEASE_PROVENANCE.md](../../docs/RELEASE_PROVENANCE.md) |
| `tools/report_portfolio_quality.py` / script | Produce freshness-aware quality/conformance reports from captured observations. | X / X / X | [docs/PORTFOLIO_QUALITY.md](../../docs/PORTFOLIO_QUALITY.md) |
| `tools/requirements-coverage-ci.txt` / policy | Hash-locked CPython 3.10/Linux x64 wheels for hosted coverage and PyYAML checks. | R / R / R | [docs/DEPENDENCY_UPDATES.md](../../docs/DEPENDENCY_UPDATES.md) |
| `tools/requirements-dev.txt` / policy | Local development pins mirroring the versions hosted CI installs. | X / X / X | [docs/CHECKER_DEVELOPMENT.md](../../docs/CHECKER_DEVELOPMENT.md) |
| `tools/requirements-portfolio-ci.txt` / policy | Hash-locked CPython 3.10/Linux x64 PyYAML wheel for portfolio fixtures. | R / R / R | [docs/DEPENDENCY_UPDATES.md](../../docs/DEPENDENCY_UPDATES.md) |
| `tools/requirements-quality-ci.txt` / policy | Hash-locked CPython 3.10/Linux x64 wheels for hosted Ruff and mypy quality checks. | R / R / R | [docs/DEPENDENCY_UPDATES.md](../../docs/DEPENDENCY_UPDATES.md) |
| `tools/test_documentation.py` / script | Offline command/reference and HTTP failure fixtures; no live website certification. | R / R / R | [docs/DOCUMENTATION_CHECKS.md](../../docs/DOCUMENTATION_CHECKS.md) |
| `tools/test_excel_evidence.py` / script | Synthetic host-record and stage-failure fixtures; does not execute Excel. | R / R / R | [docs/EXCEL_EVIDENCE.md](../../docs/EXCEL_EVIDENCE.md) |
| `tools/test_portfolio_drift.py` / script | Synthetic structural/adoption drift fixtures for all supported profiles. | X / X / X | [docs/PORTFOLIO_DRIFT.md](../../docs/PORTFOLIO_DRIFT.md) |
| `tools/test_portfolio_quality.py` / script | Synthetic quality, missing-evidence and freshness-boundary fixtures. | X / X / X | [docs/PORTFOLIO_QUALITY.md](../../docs/PORTFOLIO_QUALITY.md) |
| `tools/test_provision_repository.py` / script | Stateful simulated API fixtures for planning, guarded writes and partial failure. | X / X / X | [docs/PROVISIONING.md](../../docs/PROVISIONING.md) |
| `tools/test_release_provenance.py` / script | Release-payload fixtures and ephemeral SSH signing/verification controls. | R / R / R | [docs/RELEASE_PROVENANCE.md](../../docs/RELEASE_PROVENANCE.md) |
| `tools/test_verification_depth.py` / script | Focused behavioral failure-path coverage for release, contract, Wiki, initializer and closeout tooling. | R / R / R | [tools/README.md](../../tools/README.md) |
| `tools/test_wiki.py` / script | Offline publication link, source identity, missing/extra/modified page and safe-export fixtures. | X / X / X | [docs/WIKI_PUBLICATION.md](../../docs/WIKI_PUBLICATION.md) |
| `tools/test_workflow_validation.py` / script | Authoritative pinned-actionlint validation of workflows and invalid fixtures. | R / R / R | [tools/README.md](../../tools/README.md) |

## Tracked directories

| Directory / kind | Purpose | A / L / UI | Authority |
| --- | --- | --- | --- |
| `.github/` / directory | GitHub metadata, policies, forms and workflow automation. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `.github/ISSUE_TEMPLATE/` / directory | Governed public intake forms and private-security routing. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `.github/scripts/` / directory | Reusable label policy and live reconciliation/detection implementations. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `.github/workflows/` / directory | Hosted triggers, permissions, orchestration and evidence retention. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `assets/` / directory | Maintained project visuals and optional social preview. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `docs/` / directory | Normative contracts, source references and template-maintenance guides. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `docs/wiki/` / directory | Canonical reviewed Wiki source, ordered navigation and catalogue. | X / X / X | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `examples/` / directory | Reproducible demonstrations outside the production API. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `examples/modules/` / directory | Project-private example module exports. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `src/` / directory | Production VBA exports grouped by responsibility. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `src/core/` / directory | Internal implementation behind the facade. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `src/modules/` / directory | Public facade and standard production modules. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `tests/` / directory | Regression contracts and source-controlled test harnesses. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `tests/modules/` / directory | VBA regression module exports. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |
| `tools/` / directory | Operational validators and template-maintenance tooling. | R / R / R | [docs/REPOSITORY_STRUCTURE.md](../../docs/REPOSITORY_STRUCTURE.md) |

## Files created only during initialization

| Generated path | Profiles | Purpose and editing rule |
| --- | --- | --- |
| `.github/initialization.json` | application, library, ui-component | Immutable initialization inputs; do not hand-edit to conceal a mismatch. |
| `src/classes/README.md` | application, ui-component | Directory guidance; replace with applicable exported source when that role is implemented. |
| `src/workbook/README.md` | application | Directory guidance; replace with applicable exported source when that role is implemented. |
