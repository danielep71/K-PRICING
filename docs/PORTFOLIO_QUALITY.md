# 🧭 Portfolio Quality and Conformance

[![Scope: read only](https://img.shields.io/badge/scope-read%20only-217346)](#capture)
[![Ranking: none](https://img.shields.io/badge/ranking-none-1D76DB)](#dimensions)

This template-maintenance report combines the [structural drift model](PORTFOLIO_DRIFT.md)
with timestamped Actions, governance and published-release observations. It does
not change portfolio repositories, assign adoption, execute inspected code, or
certify releases. Generated projects do not inherit the reporter or its fixtures.

<a id="dimensions"></a>

## 📋 Evidence dimensions

| Dimension | Evidence and boundary |
| --- | --- |
| Adoption | Recorded contract source/version and project profile, or explicit ADOPT; no inferred replacement |
| Structural conformance | Version-selected #25 predicates; file presence is not implementation correctness |
| Required workflow health | Exact inspected SHA, default branch, latest push/dispatch run per workflow and its captured attempt; all required named jobs and their runs must succeed |
| Branch and tag protection | Observed active rulesets and #25 scope predicates; unavailable policy is not a passing control |
| Latest stable release | Most recently published non-draft, non-prerelease release, resolved tag commit, release text and asset metadata; publication is not certification |
| Specialist claims | Original score, scale, rationale, source file and SHA; attributed as REPORTED, never normalized into a portfolio ranking |

There is no overall numerical score. Stars, forks, traffic, downloads, watchers
and other popularity metrics are not quality inputs. A report's `status: pass`
means all six main evidence dimensions have acceptable observations, **not** that
the software or a release is certified. `UNRELEASED` is an observed absence of a
stable release, not a defect. `OBSERVED` records publication without endorsing it.
Certification is always `NOT_ASSESSED` here; release certification uses the
separate [release-evidence contract](RELEASE_EVIDENCE.md).

`KEEP`, `DEFER` and `NOT APPLICABLE` remain visible with rationale and the original
drift findings. Universal findings cannot be waived by profile exceptions.
Specialist claims require captured file evidence and retain their original scale;
the reporter checks binding/freshness, not the truth of an external review.

## ⏱️ Freshness and reproducibility

The evaluator requires explicit `--as-of`; it never consults the current clock.
The default maximum snapshot age is seven days, configurable with positive
`--max-age-days`. Exactly seven days is current. Older evidence becomes STALE;
missing or future observation times become UNVERIFIED. Timezones are mandatory.
The report retains `observed_status` for historical context but never displays
stale observations as a current PASS. Every repository shows its inspected SHA
and evidence timestamp. Specialist claims have their own timestamps and binding.

Given identical snapshot bytes, evaluator and command options, JSON and Markdown
are byte-identical. Retain the snapshot and its SHA-256 digest: GitHub settings
pages are live links, not immutable historical evidence. The captured snapshot
is the reproducible evidence for settings; source findings also link to commits.
Changing `--as-of` intentionally changes the report and may change freshness.

<a id="capture"></a>

## 📸 Capture and render

Use a template-maintenance environment with the parser dependency described in
[the drift guide](PORTFOLIO_DRIFT.md#capture). No new dependency is introduced.

```bash
python3 tools/collect_portfolio_snapshot.py --quality --contract-commit FULL_EVALUATOR_SHA owner/repo-one owner/repo-two > snapshot.json
python3 tools/report_portfolio_quality.py --snapshot snapshot.json --as-of 2026-09-09T12:00:00Z --output quality.json --summary quality.md
python3 tools/test_portfolio_quality.py -v
```

Supply every repository in the explicitly agreed scope of the portfolio brief.
The tool accepts arbitrary explicit repository lists; it neither guesses a
portfolio from account popularity nor drops repositories with missing adoption.
The public CI workflow exercises synthetic fixtures only, without portfolio
credentials, remote changes or public uploads of private snapshots.

`--quality` preserves snapshot schema 1 and adds optional `quality` evidence to
each repository. Without that flag, #25's existing evaluator remains usable.
The reporter accepts older snapshots, but absent timestamps and quality evidence
remain UNVERIFIED rather than being filled from the current clock.

The collector remains GET-only. It needs repository metadata, contents, labels,
read-only Actions and the available ruleset/release read access. Existing access
or plan restrictions are recorded, not bypassed. Optional `GH_TOKEN` stays in
the environment; never place credentials in arguments or captured files.
Administration-only or provider-specific evidence may remain unavailable.

## 🧬 Snapshot extension

Each repository's optional `quality` object contains:

- `observed_at`: timezone-aware start of repository capture (`captured_at`);
- `workflows`: complete selected run observations, or null if unavailable;
  each includes run ID, workflow path, SHA, branch, event, status, conclusion,
  attempt number and the jobs captured for that exact attempt;
- `releases`: latest published stable release and asset metadata, `[]` for an
  observed absence, or null for unavailable collection; `resolved_commit` comes
  from resolving the tag object, never from the mutable `target_commitish` name.

An optional repository-level `specialist` list contains `name`, `score`, `scale`,
`rationale`, `evidence_path`, exact `commit` and `observed_at`. Preserve the
underlying review file in `files`/`paths`, using the #25 snapshot rules. No score
is invented when a specialist review is absent. These operator annotations
do not assert that the repository adopted the template.

Collection is not atomic: code comes from the captured commit, while settings
are observations made during capture. Timestamps are not signed attestations.
The capture start is retained, so later API calls do not make earlier evidence
look newer. Read errors and collection limits never become empty successful
collections. A fatal capture failure emits no complete snapshot.

## 🛡️ Workflow and release limits

Required job names come from demonstrably applicable active default-branch
rulesets. Missing requirements are UNVERIFIED, even if unrelated workflows are
green. Ambiguous scopes/exclusions and check-provider-bound requirements remain
UNVERIFIED: Actions job evidence alone cannot authenticate a different provider.
Legacy branch-protection-only requirements are not inferred from rulesets.

PR merge-SHA runs, other branches, older candidate SHAs, duplicate job-name
matches and missing jobs cannot establish PASS. Latest attempts supersede older
successes. Pending runs are UNVERIFIED; failed, cancelled, skipped or neutral
required jobs do not pass. A captured failed run also prevents its job from
establishing PASS. Captured workflow observations remain in JSON even when the
required-check policy is unavailable.

The collector paginates [Actions runs](https://docs.github.com/rest/actions/workflow-runs)
and [attempt-specific jobs](https://docs.github.com/en/rest/actions/workflow-jobs).
It conservatively refuses Actions collections at the 1,000-result search limit.
It resolves annotated release tags with bounded indirection. It does not download
assets, validate archive contents, verify signatures or substitute publication
for the profile-aware release-evidence gate. A missing/unresolved tag is explicit.

Exit 0 means acceptable main observations; exit 1 means drift, failure, missing
adoption, stale or unverified main evidence; exit 2 means malformed input or an
operational error. Report paths cannot overwrite input or each other. Keep
private portfolio snapshots and reports private; publish only approved redacted
summaries. Do not commit them into the public template.
