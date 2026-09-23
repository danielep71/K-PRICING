# Post-release closeout

This procedure closes the gap between pre-tag certification and the provider state that exists only after a release is published. It is **read-only with respect to repository and release state**: the workflow captures provider facts, validates them, and retains evidence; it does not move tags, alter a GitHub Release, close issues, edit milestones, publish the Wiki, or upload release assets.

## Mode applicability

K-PRICING is an initialized application. Its closeout uses the selected generated
profile; canonical certification-bundle and Wiki controls are not applicable.
The workflow records that scope explicitly. A generated project must separately
adopt and implement any stronger local policy; canonical requirements are not
automatically inherited. The canonical template procedure below includes its
additional certification and Wiki checks.

## Procedure

After the annotated tag, tag-triggered static checks, GitHub Release, release milestone, and (for the canonical template) Wiki publication are complete, run the **Release closeout** workflow manually from the released repository. Supply:

- `tag`: the published tag, for example `v1.2.1`;
- `candidate_sha`: the exact 40-character SHA that was certified before tagging;
- `milestone_number`: the GitHub milestone number used for that release;
- `wiki_browser_reviewed`: `true` only after a human has opened the published Wiki and checked Home, sidebar/navigation, and representative links in a browser;
- `expect_prerelease`: normally `false`; set it only for an intentionally prerelease publication;
- `allow_not_latest`: normally `false`; set it only when the release is intentionally not expected to be GitHub's current latest release.

The workflow checks out the exact candidate, captures GitHub REST facts, performs the existing Wiki read-back comparison, tests GitHub-generated source archive retrieval, builds one retained snapshot, and evaluates that snapshot with `tools/_release_closeout.py`. The terminal workflow verdict is green only when the closeout report, readable summary, and retained evidence are all produced successfully.

## Deterministic provider controls

The closeout validator binds all deterministic checks to the same candidate SHA. It verifies:

1. the remote tag ref is an **annotated tag object**, the tag object has the expected name, and it resolves to the certified commit rather than a moved or lightweight tag;
2. the canonical static-check workflow has a completed successful `push` run whose `head_branch` is the release tag and whose `head_sha` is the candidate;
3. the GitHub Release exists for that tag, is published rather than draft, has the expected prerelease flag, and matches the expected latest-release state;
4. the filtered **product assets** are compared with the selected candidate profile's `allowed_asset_globs` from `.github/release-policy.json`;
5. a source-only profile has no product assets; binary-capable profiles may contain only names allowed by their candidate-bound patterns;
6. `VERSION`, the released CHANGELOG heading, the released comparison link, and the `Unreleased` comparison link remain coherent with the tag;
7. the provider comparison range resolves to the release range and contains the certified candidate when the range is ahead;
8. milestone closure is evaluated from the actual captured milestone membership and item states, with the provider's open/closed counters reconciled to those items rather than trusted as a UI percentage.

An unexpected uploaded asset, lightweight or moved tag, failed/missing tag CI, draft or incorrectly classified Release, wrong VERSION/tag relationship, unresolved comparison, open milestone, or stale milestone counters is non-green.

## Uploaded assets are not source archives

GitHub's generated `zipball_url` and `tarball_url` are provider-generated source archives. They are **not** entries in the Release API `assets` array and are never interpreted as runtime binaries or uploaded release payloads.

The closeout report therefore records these separately:

- raw Release API `assets`: all explicitly uploaded files, retained in provider metadata;
- canonical certification assets: the ZIP, detached `.zip.sig`, manifest and SHA-256
  file, partitioned by `release_certification.py` and independently downloaded,
  hash-checked and signature-verified against the candidate/current trust policy;
- `uploaded_assets` in the normalized closeout report: the remaining product assets,
  governed by `allowed_asset_globs`; in generated mode all raw assets remain product assets;
- source ZIP/tar exposure and retrieval: provider observations confirming that GitHub's generated source archives are available.

For the canonical template's source-only profile, product `uploaded_assets` must
remain empty while the four certification assets and generated source archives
are required. Certification files are not exempt from verification: the workflow
binds the separate certification verdict to its terminal result. The helper
evaluates the filtered snapshot; helper success alone is not certification success.

## Wiki and UI observations

Wiki publication remains owned by `tools/check_wiki.py`; the closeout helper does not implement a second Wiki policy. For template-mode releases, the workflow clones the published Wiki, invokes the existing checker against the exact candidate source, and includes that checker's result in the closeout snapshot.

Some evidence is necessarily an **observation**, not a repository fact. The report labels these separately from deterministic provider controls:

- successful retrieval of GitHub-generated ZIP/tar source archives;
- the explicit browser review of the Wiki Home/sidebar/navigation.

A missing or failed observation remains non-green, but its category is preserved so the report does not imply that a browser review or network retrieval was derived from Git history.

## Evidence retained

The workflow retains, for 90 days:

- captured provider metadata, including the unfiltered Release asset inventory;
- `certification-plan.json`: certification/product partition or explicit not-applicable scope;
- `certification-verification.json`: independent verification verdict or not-applicable scope;
- `snapshot.json`: the normalized input facts used by the validator;
- the authoritative Wiki comparison JSON/Markdown when applicable;
- `release-closeout.json`: machine-readable closeout result;
- `release-closeout.md`: concise human-readable closeout result.

`release-closeout.json` records the SHA-256 of `snapshot.json`, the candidate/tag/profile identity, deterministic and observation status, tag/CI/Release/asset/comparison/milestone/Wiki state, and categorized findings. Retaining the snapshot makes the conclusion replayable without relying on a maintainer workstation.

The 90-day Actions reports are diagnostic retention. Canonical certification
assets remain attached unchanged to the GitHub Release for the lifetime of that
release, as required by [RELEASE_EVIDENCE.md](RELEASE_EVIDENCE.md); an expiring
Actions artifact does not replace that durable record.

## Offline fixture contract

The helper has a deterministic self-test:

```bash
python3 tools/_release_closeout.py --self-test
```

The fixture matrix covers the valid path plus VERSION/tag mismatch, lightweight and moved tags, failed tag CI, draft/unexpected-prerelease/not-latest Release state, unexpected source-only assets, missing source-archive exposure, unresolved comparison, open milestone membership/state, stale milestone counters, Wiki drift, and source-archive retrieval failure. These fixtures do not call GitHub and do not mutate live provider state.

## Scope boundary

Post-release closeout supplements, but does not replace, pre-tag `check_release.py` certification, provenance validation, Excel runtime evidence, or the Wiki publication checker. A green closeout report means the captured post-publication state is coherent with the certified candidate under this contract; it is not a new claim about Excel execution, numerical accuracy, UI runtime behavior, or release-binary provenance.
