# ⚙️ Workflow and tool reference

> **Guide, not policy:** [tools/README.md](../../tools/README.md) is authoritative.

The file inventory gives every configuration and script its purpose, owner,
lifecycle and safe-edit rules. This page explains how the automation fits
together. YAML and script source links in the inventory are the execution
authority; the documents own the user-facing contracts.

## Hosted workflows

| Workflow file | Trigger | Permissions / inputs | Output and evidence |
| --- | --- | --- | --- |
| `static-checks.yml` | Push main/release branches/version tags; PR to main/release; manual; reusable call | Contents read; optional `expected-profile` for reusable calls | `Repository integrity`, candidate SHA output, JSON/Markdown and 30-day artifacts |
| `checker-development.yml` | Release push; main/release PR; manual | Contents read; exact PR head/source | Checker contract, independent self-tests and policy coverage; template only |
| `portfolio-drift.yml` | Release push; main/release PR; manual | Contents read; synthetic API fixtures | Structural drift, quality and provisioning fixture logs; template only |
| `labels-sync.yml` | Path-filtered PR; trusted path-filtered main push; manual | PR read-only; trusted reconciliation adds issues write; policy resolves profile/domains | Validate or reconcile labels; post-write verification |
| `labels-drift.yml` | Path-filtered PR; daily 05:17 UTC; manual | PR fixtures; live contents/issues read | Label differences and 30-day evidence; no writes |
| `external-links.yml` | Monday 06:17 UTC; manual | Contents read; anonymous bounded HTTP | Categorized observations and 14-day artifacts; separate from PR source checks |
| `wiki-checks.yml` | Main/release PR and push; manual | Contents read; no network comparison | Wiki source, inventory, navigation, publication fixtures; template only |
| `wiki-drift.yml` | Monday 06:47 UTC; manual | Contents read; anonymous wiki clone | Separate network availability and deterministic published-byte comparison; template only |

Schedules start from the default branch. A release-branch file does not activate
its schedule. Path-filtered label workflows need not run for unrelated changes;
absence on an unrelated PR is not a failed fixture.

## Operational command families

Run commands from the repository root. Use `python SCRIPT_PATH --help`
for that script's actual arguments; do not assume every script has `--self-test`.

| Family | Scripts | Inputs and outputs / authority |
| --- | --- | --- |
| Canonical source | `check_repo.py` | Profile policy and tracked files → 21-rule console/JSON/Markdown; Repository Structure |
| VBA contracts | `check_vba_public_api.py`, `check_vba_jumps.py`, `check_vba_conditionals.py` | Source, roles and API manifest → focused findings; Tools README |
| Whitespace and actions | `check_committed_whitespace.py`, `check_local_actions.py` | Working/committed Git range or local action metadata → findings; Tools README / LOCAL_ACTIONS |
| Contract and version | `check_template_contract.py`, `check_release_semantics.py` | Recorded adoption and version/changelog → findings; respective contracts |
| Documentation | `check_documentation.py`, `check_external_links.py` | Commands/references or anonymous HTTP observations → separate reports; Documentation Checks |
| Release | `check_release.py`, `check_excel_evidence.py` | Candidate, actual external records/logs/assets → validation; release/host contracts |
| Initialization | `initialize_repository.py` | Clean source and explicit values → plan, optional apply, immutable input record |
| Portfolio | `collect_portfolio_snapshot.py`, `check_portfolio_drift.py`, `report_portfolio_quality.py` | GET-only capture → structural comparison and timestamped quality report; template only |
| Provisioning | `provision_repository.py` | Explicit new target, source SHA, reviewed digest → plan or guarded writes/journal; template only |
| Reusable fixtures | `create_reusable_workflow_fixture.py` | Explicit profile, destination and immutable provider SHA → disposable caller; template only |
| Wiki | `check_wiki.py` | Catalogue/source → offline check, explicit reference refresh, pinned export or read-only comparison; template only |

## Helpers and tests

`_gatelib.py` supplies shared Git/report/CLI primitives.
`release_provenance.py` extends the release gate; neither is a standalone CLI.
The `policy_coverage_*` modules and `checker_development.py` maintain the
template checker, not generated application behavior.

The `test_*.py` files exercise failure boundaries for their named tool.
`test_workflow_validation.py` additionally needs the verified pinned actionlint
binary; the static workflow installs it. Fixture PASS does not establish that
your repository was provisioned, your wiki was published, or Excel was run.

## Configurations

Repository-profile JSON selects roles, paths, placeholders and adopted contract;
initialization JSON records applied inputs. Label JSON defines core/profile/domain
sets. Provisioning JSON describes supported live setup. Release policy defines
required evidence and asset globs; provenance policy defines source/workflow and
signing trust; Excel policy fixes the starter harness contract; documentation
policy owns references, approved domains, bounds and expiring exceptions.
The wiki catalogue records ordered pages and inventory descriptions.

Edit the owning configuration only for a reviewed contract change. A policy
change that makes a red result green needs a reason and relevant tests; it is
not ordinary troubleshooting.

---
[Home](Home.md) · [Previous](File-and-directory-reference.md) · [Next](Maintenance-and-migration.md)
