# 🔐 Supply-Chain Assurance

This document owns the repository's supply-chain analysis layer.
It complements the deterministic repository gates and the controlled dependency
update policy; it does not replace either one.

## Private repository policy

This repository is private. All Scorecard scan/publication jobs are skipped
while it remains private, and the README does not request public Scorecard,
release or issue-count badges for it. No public Scorecard result is claimed.

CodeQL jobs are skipped for private repositories unless the maintainer has
confirmed the applicable private-code license/entitlement and sets the repository
variable `ENABLE_PRIVATE_CODEQL=true`. This variable is not enabled during
initialization. Skipped security analysis is an explicit assurance gap, never a
successful scan. The deterministic repository, workflow, VBA and release gates
remain active. Before a functional release, resolve and record the required
security-analysis coverage for the chosen private distribution model.

The public-publication paths described below apply only if the owner later
explicitly changes the repository visibility; initialization does not do so.

## Controls

| Control | Purpose | Trusted execution boundary |
| --- | --- | --- |
| CodeQL | Static security analysis of maintained Python and JavaScript source | `push` to `main` / `release/**`, schedule, or explicit manual dispatch |
| Dependabot | Weekly discovery of GitHub Actions updates | Proposal-only pull requests; no automatic approval or merge |
| OpenSSF Scorecard | Exact-commit release-candidate analysis plus public default-branch posture publication | CLI scan on trusted non-default branch push/manual dispatch; published Action on `main`, branch-protection changes, schedule, or main-branch manual dispatch |

All external Actions remain pinned to full immutable commit SHAs with audited
semantic-version comments. Dependabot may propose a new Actions revision, but a
proposal is not trusted merely because a bot opened it: the source/tag must be
reviewed under [DEPENDENCY_UPDATES.md](DEPENDENCY_UPDATES.md), the resulting
workflow must retain immutable pins, and merge remains manual.

The published Scorecard path uses `ossf/scorecard-action` v2.4.4, resolved to its
reviewed commit. OpenSSF restricts published workflows: they may not define
top-level `env` or `defaults`, the publishing job may not define job-level
`env` or `defaults`, and that job may use only OpenSSF-approved Actions.
The release-candidate CLI version/digest are therefore scoped only to the
non-publishing job. The published Action uses `file_mode: git` so file-based
checks enumerate the reviewed repository rather than depending on the provider
source-tarball path; this is important because an empty archive view can otherwise
make workflows, Dependabot and SAST appear absent while API-based checks still
succeed. After the publishing Action completes, a separate read-only job
retrieves the public `api.scorecard.dev` record for the exact GitHub SHA and
then retrieves the badge SVG itself. Missing, stale, mismatched, non-SVG, or
error-bearing badge data (including `invalid repo path`) makes the workflow
non-green.

OpenSSF also restricts Action publication on `push` and `schedule` to the
repository default branch. Release candidates therefore use the official
Scorecard CLI v5.5.0 with `--commit` bound to the exact candidate SHA. The Linux
amd64 release archive is verified against its recorded SHA-256 before execution;
the resulting JSON and CLI version are retained as workflow evidence.

## Permission model

The Scorecard workflow does not use `pull_request` or
`pull_request_target`. CodeQL uses two distinct trust paths: trusted
push/schedule/manual runs receive `contents: read` plus
`security-events: write` for Code Scanning publication, while
`pull_request` runs use a separate read-only job with `contents: read` only,
set `upload: never`, and retain SARIF as an Actions artifact. No
`pull_request_target` path exists and untrusted pull-request code receives no
write-capable token or secrets. The Scorecard release-candidate
CLI job receives only `contents: read`. The default-branch publication job
receives `contents: read`, `security-events: write`, and `id-token: write`; the
last permission is used only for authenticated Scorecard result publication.
The repository-quality gate rejects `pull_request_target` and rejects
write-capable permissions in any workflow that runs on `pull_request`. Event
detection, workflow-level permission analysis, job-level permission analysis,
and immutable Action pin validation remain separate checker routines so each
finding is independently testable and the canonical gate stays within its
enforced complexity ceiling.

Checkout credentials are not persisted. No supply-chain workflow has a source,
release, label, issue, branch, or repository mutation step.

## Evidence boundary

A successful trusted CodeQL run means the configured CodeQL queries completed
against the exact checked source and were eligible for Code Scanning publication.
A successful pull-request CodeQL run means the same query set analyzed the
GitHub pull-request merge commit under a read-only token and retained SARIF
without privileged publication. A successful release-candidate Scorecard CLI run means
the official scanner completed against the requested commit and its retained
JSON is candidate-bound; some Scorecard checks necessarily observe current
repository settings rather than historical settings. The Scorecard Action can
complete while reporting publication rejection as a warning, so its job result
alone does **not** prove public publication. Public publication is claimed only
when the separate verification job retrieves an exact-SHA record from
`api.scorecard.dev`. A Dependabot proposal means only that GitHub detected a
candidate dependency update.

None of those outcomes proves VBA compilation, Excel runtime behavior,
numerical accuracy, UI cleanup, packaging integrity, or release certification.
Those claims remain owned by their existing specialist and release evidence.
Likewise, a public Scorecard score is not the repository's project-quality
score.

## Review and maintenance

Pull-request review uses the existing read-only deterministic gates; analyzer
publication is deferred until the reviewed change reaches the default branch.
Release-branch Scorecard evidence is an exact-commit CLI scan, not a claim that
the public Scorecard service published that non-default candidate.

Before merging a CodeQL, Scorecard, or dependency-workflow change:

1. resolve every proposed Action release tag to its upstream commit and record
   the immutable SHA; verify downloaded CLI assets against recorded official
   release digests;
2. review permissions, triggers, checkout behavior, new network activity, and
   publication surfaces;
3. run repository integrity, authoritative workflow validation and the retained
   supply-chain fixtures; checker-development tooling applies only to canonical
   template maintenance and is removed from initialized projects;
4. preserve manual dependency approval and rollback under
   [DEPENDENCY_UPDATES.md](DEPENDENCY_UPDATES.md);
5. retain exact-revision analyzer results when eligible under the private policy
   above. Public Scorecard scan/publication and exact-SHA publication verification
   apply only to a public repository. For this private repository, record skipped
   analysis as unavailable coverage and retain the successful deterministic CI
   result separately. Resolve required coverage before a functional release;
   neither a skip nor static CI substitutes for a security scan.

If a required security analyzer is unavailable, misconfigured, or denied its
required permissions, its workflow is non-green when selected to run. The private eligibility skips above
remain documented limitations, not completed analysis. Do not convert an unavailable
result into a pass or weaken deterministic gates to recover a public score.
