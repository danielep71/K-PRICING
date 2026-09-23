# 🧭 Read-only Portfolio Drift

[![Mode: read only](https://img.shields.io/badge/mode-read%20only-217346)](#capture)
[![Evidence: explicit gaps](https://img.shields.io/badge/evidence-explicit%20gaps-1D76DB)](#decisions)

This is a structural conformance detector, not a numerical, Excel, source-code
correctness or release certification tool. It never runs inspected code and
never changes inspected files, branches, issues, settings, profiles or versions.
The collector uses GET only. The evaluator has no network access and consumes a
captured JSON snapshot. Tooling and fixtures are template-only and are removed
from generated repositories.

## 🧬 Version selection

The evaluator reads the captured `.github/repository-profile.json` and selects
`CONTRACT_RULE_SETS` for its recorded version. It never substitutes the newest
version. `contract_source` and `contract_commit` in the snapshot identify the
template evaluator/documentation being used; they do not assign a contract to
any inspected repository. Use the actual evaluator commit when collecting.

Missing adoption is `ADOPT`; unavailable content or an unsupported version is
`UNVERIFIED`. Repositories without adoption are not scored against an invented
profile or baseline. Their live governance remains unscored. This is expected
for established repositories that have not adopted the new template.

<a id="decisions"></a>

## ⚖️ Decisions and observations are separate

| Decision | Meaning |
| --- | --- |
| `REQUIRED` | A recorded baseline control applies |
| `ADOPT` | Adoption needs a maintainer decision; the tool assigns nothing |
| `KEEP` | Preserve the documented local control; not a waiver of universal controls |
| `DEFER` | Record a postponement while retaining the observed finding |
| `NOT APPLICABLE` | Only for an optional profile-specific control; never a universal waiver |

Observations are `PASS`, `DRIFT`, `UNVERIFIED` or `NOT_APPLICABLE`. A `PASS` means
the documented **structural predicate below** passed, not that the named checker
is correct or has run. Missing runtime proof for a retained specialist control
remains `UNVERIFIED`. Both drift and unknown evidence produce a nonzero overall
verdict. An empty portfolio cannot pass.

Decisions are operator-supplied snapshot annotations with `rule`, `decision`,
`reason`, exact inspected `commit` and `evidence_path`. The evidence file must be
captured. The evaluator validates binding and presence, not the truth of a human
claim of semantic equivalence. Stale, duplicate and unknown decisions are
rejected. No annotation suppresses a universal finding. Generic workflow location
can be selected through `policy.workflow`; stronger local job names are retained.

## 📋 Predicate register

The registry selects which versioned rows apply. Metadata, generic-workflow and
protection predicates are the universal governance floor described by the
post-creation contract. Required paths come from the evaluator, not a consumer's
editable required-path list. Additional local files, labels and specialist jobs
are not replaced merely because they differ from the template.

| Rule | Predicate / limitation |
| --- | --- |
| <a id="adoption"></a>`adoption` | Captured supported version and matching source; no fallback |
| <a id="canonical-repository-gate"></a>`canonical-repository-gate` | README, contribution, security, licence and canonical checker paths present; does not prove a checker implementation |
| <a id="profile-model"></a>`profile-model` | Generated mode, supported profile and matching repository identity |
| <a id="placeholder-schema"></a>`placeholder-schema` | Nonempty placeholder catalogue |
| <a id="label-policy"></a>`label-policy` | Baseline core label names retained; declared core name/colour/description match live labels; extra labels retained |
| <a id="release-integrity"></a>`release-integrity` | Release docs/version/changelog/policy paths and universal repository/compile/regression evidence requirements retained; does not certify a release |
| <a id="deterministic-initializer"></a>`deterministic-initializer` | Initializer path present; determinism itself is not proven here |
| <a id="committed-whitespace"></a>`committed-whitespace` | Focused checker path present |
| <a id="complete-public-api"></a>`complete-public-api` | Public-API manifest and checker paths present |
| <a id="label-drift-detection"></a>`label-drift-detection` | Label-drift workflow path present |
| <a id="nested-conditional-compilation"></a>`nested-conditional-compilation` | Conditional checker path present |
| <a id="procedure-scoped-jumps"></a>`procedure-scoped-jumps` | Jump checker path present |
| <a id="repository-local-actions"></a>`repository-local-actions` | Local-Action checker path present |
| <a id="strict-release-semantics"></a>`strict-release-semantics` | Release-semantics checker path present |
| <a id="template-contract-version"></a>`template-contract-version` | Contract authority and checker paths present |
| <a id="controlled-dependency-updates"></a>`controlled-dependency-updates` | Dependency policy path present; human review remains outside automated proof |
| <a id="documentation-drift"></a>`documentation-drift` | Documentation policy, guide and checkers present; command/reference correctness and external availability are evaluated separately |
| <a id="advanced-release-provenance"></a>`advanced-release-provenance` | Provenance policy, implementation, fixtures and documentation paths present; artifact and signature verification belongs to the release gate |
| <a id="metadata"></a>`metadata` | Description, enabled issue intake and disabled auto-merge |
| <a id="workflow-properties"></a>`workflow-properties` | Selected generic workflow has explicit contents-read-only permissions and immutable external references; no job secrets/environment; execution, triggers and transitive actions are not certified |
| <a id="branch-protection"></a>`branch-protection` | Active rules cover default branch, prohibit deletion/non-fast-forward, require PR and a strict named check, with no bypass |
| <a id="release-tag-protection"></a>`release-tag-protection` | Active rules cover version tags and prohibit update/deletion/non-fast-forward with no bypass |
| <a id="local-specialist"></a>`local-specialist` | Retain documented specialist requirement; runtime proof remains unverified |
| <a id="ui-evidence"></a>`ui-evidence` | Optional UI evidence is inapplicable outside UI profile; universal release requirements still apply |

Ambiguous YAML (aliases, merges, duplicate keys or custom tags) is not guessed.
Ruleset exclusions and unavailable evidence remain unverified. Unsupported scope
patterns cannot establish coverage. Disabled/nonmatching rulesets do not count.
These deliberate limitations avoid claiming equivalence from a text match or
from a successful unrelated workflow run.

<a id="capture"></a>

## 📸 Capture and reproduce

Install `PyYAML==6.0.3` in the template-maintenance environment, then:

```bash
python3 tools/collect_portfolio_snapshot.py --contract-commit FULL_EVALUATOR_SHA owner/repository > snapshot.json
python3 tools/check_portfolio_drift.py --snapshot snapshot.json --output drift.json --summary drift.md
python3 tools/test_portfolio_drift.py -v
```

The collector accepts multiple repository names and optional `GH_TOKEN` with only
the read permissions needed for those repositories (contents, metadata and issue
labels; ruleset access may require an account capability the token cannot grant).
Never put tokens in arguments, snapshots or reports. Redirects are refused, and
credentials are sent only to `api.github.com`. A 403 for a private repository's
plan-restricted rulesets is captured as unavailable, not absent or compliant.
No upgrade or public-visibility change is performed.

Snapshot schema 1 contains `contract_source`, `contract_commit` and nonempty
`repositories`. Each entry has `repository`, exact `commit`, complete `paths`
(or null for unavailable/truncated trees), captured UTF-8 `files`, `metadata`,
`labels`, `rulesets`, and `decisions`. Metadata is a selected REST repository
response; labels are complete paginated label objects; rulesets are full details,
not collection summaries. Unavailable collections are null; an observed empty
collection is `[]`. `unavailable` records collection limitations. Missing file
content is not equivalent to missing a file from a complete tree.

Code is read at an immutable commit. Live settings cannot be atomically bound to
that commit: snapshots record observations, not signed attestations. Reports
include a snapshot digest and links to source and evaluator-rule evidence. For
reproduction, retain the snapshot privately and use the same evaluator commit.
Identical input produces byte-identical JSON/Markdown; reports add no current
timestamp. Capture failures never emit a deceptively complete partial portfolio.

Collector exits 0 on capture or 2 on failure. Evaluator exits 0 for all structural
predicates passing/inapplicable, 1 for drift or evidence gaps, and 2 for malformed
input/operational failure. Explicit report outputs may be written; input snapshots
cannot be overwritten. Keep private repository snapshots/reports private, and
do not commit them to the public template. The hosted workflow runs only synthetic
fixtures and has no portfolio credentials or private artifacts.

## 🔒 Parser dependency review

PyYAML 6.0.3 is a new template-maintenance dependency, not a VBA runtime dependency.
Its official [PyPI release](https://pypi.org/project/PyYAML/6.0.3/) identifies the
maintainers, MIT licence, upstream source and Python >=3.8 compatibility.
[Parser documentation](https://pyyaml.org/wiki/PyYAMLDocumentation) describes its
node/event APIs. This evaluator composes nodes and never constructs Python
objects; aliases/custom tags are refused and input size/depth is bounded.
Only the dedicated fixture workflow installs it. No new permissions, secret
inheritance or automatic updates are introduced. The installation is version
pinned, not a hash-locked transitive build. Rollback is the complete #25 commit
on top of the prior passing #22 baseline; remove the parser-dependent workflow
and tools together. Future parser updates follow the dependency-update policy.
