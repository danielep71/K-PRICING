# 🚀 K-PRICING Release Guide

[![Release model: exact source](https://img.shields.io/badge/release-exact%20source-0969da)](#release-invariants)
[![SemVer contract](https://img.shields.io/badge/versioning-SemVer-3f4551)](docs/RELEASE_SEMANTICS.md)
[![Evidence contract](https://img.shields.io/badge/evidence-required-success)](docs/RELEASE_EVIDENCE.md)
[![Security policy](https://img.shields.io/badge/security-private-d73a49)](SECURITY.md)

This document is authoritative for the **maintainer release sequence**. Strict
version/changelog semantics are owned by
[`docs/RELEASE_SEMANTICS.md`](docs/RELEASE_SEMANTICS.md); external evidence JSON,
profile evidence and asset-manifest schemas are owned by
[`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md).

> [!IMPORTANT]
> Generated-project releases must be fully initialized. The canonical template's
> own release is the documented exception: it preserves registered template
> tokens and uses the `template` release-evidence profile.

## 🧭 Release identity

| Property | Authority |
| --- | --- |
| Project/profile | K-PRICING / application |
| Current version | [`VERSION`](VERSION) |
| User-visible history | [`CHANGELOG.md`](CHANGELOG.md) |
| SemVer/changelog policy | [`docs/RELEASE_SEMANTICS.md`](docs/RELEASE_SEMANTICS.md) |
| Installation contract | [`INSTALLATION.md`](INSTALLATION.md) |
| Evidence schema | [`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md) |
| Vulnerability handling | [`SECURITY.md`](SECURITY.md) |

<a id="release-invariants"></a>

## 🔒 Release invariants

A release is valid only when:

1. one exact candidate SHA is frozen and reviewable;
2. version/changelog/tag semantics pass the strict release-semantic gate;
3. repository/static checks pass on that candidate;
4. VBA compile and applicable regression/specialist checks pass on that candidate;
5. every distributed artifact is derived from and tested against that candidate;
6. external evidence and optional asset hashes bind to the candidate;
7. the annotated lower-case `v*` tag targets the certified commit; for the
   canonical template, that tag also carries a verified SSH signature from the
   signer trust policy committed in the candidate;
8. the canonical template's durable certification ZIP is signed in the dedicated
   certification namespace and verified against the same committed/current
   GitHub SSH signing trust before publication; and
9. post-publication retrieval/installation checks pass.

If source changes after certification, the affected evidence is stale and must be
rerun. Never compensate by manually editing an already-tested artifact.

## 1. Freeze and identify the candidate

Verify branch and tag protection before a functional release. Re-read the active
controls recorded in
[setup verification](docs/SETUP_VERIFICATION.md). Do not treat this setup
checkpoint as release authorization.

Start from the repository's reviewed release path, freeze scope, and record the
exact base/candidate revisions.

```bash
git fetch --tags --prune
git rev-parse HEAD
git status --short
git diff --stat <previous-tag>...HEAD
```

A dirty tree, unexplained generated file or unreviewed binary delta is blocking.

## 2. Synchronize version and user-visible change surfaces

Update the applicable release surfaces in one reviewable change:

- `VERSION`;
- the dated `CHANGELOG.md` release section and comparison links;
- user-facing documentation/examples affected by the release; and
- package metadata where the project actually has one.

When moving `Unreleased` entries into the versioned section, use the calendar
date on which that reviewed release section is **cut/frozen for the candidate**.
That is the changelog release date. It is independent of the later annotated-tag
creation date and GitHub Release publication date. Do not rewrite a frozen or
published section merely because tagging or publication occurs on a later day.
If a candidate is abandoned and the release section is materially reopened, a
later reviewed cut/freeze may deliberately use a new date.

Do not duplicate the remaining SemVer/order/link rules here. Run the authoritative
semantic contract:

```bash
python3 tools/check_release_semantics.py --root . --self-test
python3 tools/check_release_semantics.py --root .
```

The gate verifies Gregorian validity and non-backward cut/freeze-date ordering;
it deliberately does not infer the cut/freeze event from Git, tag, or provider
timestamps. Historical changelog sections and immutable evidence remain
historical.

## 3. Verify documentation and installation

From a clean environment:

- follow [`INSTALLATION.md`](INSTALLATION.md);
- verify source paths, component names, prerequisites and supported upgrade path;
- verify README examples and the supported public surface;
- confirm the security and license links; and
- remove stale compatibility or evidence claims.

For a generated-project release, verify no unresolved template state remains.
For the canonical template's own release, preserve registered template state and
use the `template` evidence profile.

## 4. Run repository and release gates

At minimum:

```bash
python3 tools/check_repo.py --root . --self-test
python3 tools/check_repo.py --root .
python3 tools/check_release.py --root . --self-test
```

Run every project-specific numerical, UI, lifecycle, performance or packaging
gate as well. A stronger specialist gate is additive; the generic repository
gate never replaces it.


## 5. Certify in Excel

Use the exact candidate source in each advertised Excel environment:

1. import only candidate-controlled exports;
2. run **Debug → Compile VBAProject**;
3. execute the documented regression entry point;
4. run applicable UI/lifecycle/platform/manual checks; and
5. record environment, counts, failures, completeness and cleanup.

For the migrated date layer, the primary human-facing regression entry point is
`KPR_Tests_Run`, with focused host, shape and dynamic-array runners retained in
`KPR_REGRESSION_TESTS`. Retained machine-readable host evidence uses
`KPR_Tests_RunEvidence`, which delegates to the same all-suite dispatcher and
only adapts its result into the evidence log format. Release evidence must bind
the exact candidate to the actual suite/case results, failures/skips and tested
Excel environment; historical KPR runs are not destination certification.

Source inspection is not Excel execution. If code changes, recertify.

## 6. Build and test release artifacts

Source-only libraries do not need an artificial binary asset. When the project
ships a workbook/add-in/package:

1. build from a clean location using only candidate-controlled inputs;
2. preserve required binary companions such as `.frx` files;
3. exclude development-only material unless promised;
4. reopen and smoke/regression-test the packaged artifact;
5. record filename, size and SHA-256; and
6. never edit the artifact after hashing.

The exact manifest format and profile-specific evidence requirements are defined
only in [`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md).

## 7. Create and validate external evidence

Keep candidate-binding release evidence outside the candidate tree to avoid
self-referential commit hashes. Prepare the evidence JSON and, when applicable,
the sorted asset manifest according to
[`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md).

The evidence must identify the exact candidate, executed checks, environment and
material limitations. A release note summarizes evidence; it does not replace it.

## 8. Review and merge the release candidate

The release review should make these facts easy to verify:

- target version and previous tag;
- candidate SHA and final diff;
- semantic/repository/release gate results;
- Excel and specialist evidence;
- artifact manifest/hashes when applicable;
- compatibility/migration/security notes; and
- remaining limitations.

### Merge convention

Use **Squash and merge** for release PRs and focused stabilization PRs into
`main`. Each PR should leave one commit describing the resulting change;
intermediate planning, progress and fixup commits remain in the PR history.
This means one commit per PR, not necessarily one commit per version when
stabilization spans several PRs.

Choose the merge method in the PR before merging. The squash commit title and
description should explain the delivered behavior and link the relevant issues
and evidence; do not copy the intermediate commit log as the description.
Required checks must pass, and the merge must use the reviewed head SHA.

A history-preserving merge is an exception when retaining individual commit
ancestry serves a concrete integration or provenance need. Record the reason
and maintainer decision in the PR before merging. Availability of multiple merge
methods in GitHub settings does not override this convention; this is a review
policy, not a claim that repository settings enforce squash-only merging.


An implementation PR may merge before release certification for stabilization.
That does not authorize tagging or publication. Keep incomplete acceptance work
open, then certify the final `main` commit before tagging. A squash or merge
creates a new source identity: retain original evidence attribution and obtain
the required final-candidate evidence rather than silently rebinding old results.


Do not force-push shared `main` or rewrite published tags to make historical
merges conform retroactively. Apply this convention to future merges.


## 9. Create and verify the release tag

Tag only the certified commit and run the release-integrity checker before and
after creating the local tag. The default generated-project contract remains an
annotated tag; generated repositories do not inherit the canonical SSH-signing
requirement unless they explicitly adopt an equivalent local policy.

### Initialized generated project

Use this sequence for an initialized generated project under the default tag
policy. Run the entire subshell block; each failure stops subsequent commands
and returns a nonzero status without exiting an interactive parent shell:

```bash
(
git switch main || exit $?
git pull --ff-only || exit $?
candidate_sha="$(git rev-parse HEAD)" || exit $?
release_version="$(tr -d '\r\n' < VERSION)" || exit $?
release_tag="v${release_version}"

python3 tools/check_release.py \
  --root . \
  --tag "$release_tag" \
  --candidate-sha "$candidate_sha" \
  --evidence ../release-evidence.json \
  --output test-results/release-integrity.json \
  --summary test-results/release-integrity.md || exit $?

git tag -a "$release_tag" "$candidate_sha" -m "K-PRICING ${release_version}" || exit $?

python3 tools/check_release.py \
  --root . \
  --tag "$release_tag" \
  --candidate-sha "$candidate_sha" \
  --evidence ../release-evidence.json \
  --require-tag-ref || exit $?

git push origin "refs/tags/$release_tag:refs/tags/$release_tag" || exit $?
)
```


Add `--asset-manifest ../release-assets.sha256` to both applicable release-gate
invocations when the release distributes binary assets. For contract 1.2.0 add
the build record and any required provenance-record signature using
[the provenance procedure](docs/RELEASE_PROVENANCE.md). The provenance-record
signature is an independent assertion signature; it never substitutes for the
canonical template's Git-tag signature.

Do not push the tag if either release check fails. An incorrect local tag that
has **not** been pushed may be deleted and recreated after the candidate and
signing setup are corrected, followed by a complete post-tag verification. Once
a tag has been pushed or published, never move, delete, recreate, or replace it
to hide an error; publish a corrected patch release instead.

## 10. Publish the GitHub Release

If using the optional [host evidence interface](docs/EXCEL_EVIDENCE.md), include
its `excel-host-evidence` check and `--excel-evidence` in both pre-tag and
post-tag validations. A manual run remains explicitly manual; an unavailable
runner is not compile or regression evidence.

After verifying the required protection controls, create the release from the
certified annotated tag. Include:

- user-facing summary/highlights;
- upgrade or migration notes;
- supported platform statement;
- known limitations;
- installation link;
- artifact/hash table when applicable;
- changelog comparison link; and
- security-reporting link.

Curate the release notes from `CHANGELOG.md`, the owning issues/PRs and retained
release evidence. Do not generate authoritative notes from raw commit subjects:
an approved ancestry-preserving exception can legitimately make Git history
contain subjects that are unsuitable or duplicated as user-facing release text.

Before creating the canonical template's GitHub Release, build and verify the
durable certification set using
[the release-evidence procedure](docs/RELEASE_EVIDENCE.md), sign the exact
certification ZIP with the externally held canonical SSH signing key in the
`excel-vba-release-certification` namespace, and verify that detached signature
against the candidate's committed/current GitHub signing trust. Upload the
resulting ZIP, `.zip.sig`, manifest, and SHA-256 file unchanged. This detached
signature authenticates the durable certification ZIP; it is separate from the
SSH-signed Git tag and from any optional provenance-record signature.

Upload the already-tested, already-hashed and, where required, already-signed
artifacts. Do not rebuild or re-sign between certification/tagging and
publication.

## 11. Verify after publication

- [ ] Tag resolves to the certified SHA.
- [ ] `VERSION` and changelog agree with the tag.
- [ ] Published assets download and hashes match.
- [ ] Installation and documentation links work.
- [ ] Packaged artifact, when present, passes its published smoke test.
- [ ] Source archive contains the expected release tree.
- [ ] Default branch is ready for the next Unreleased cycle.

Do not announce broad availability until these checks pass.

## 🧯 Recovery

Before publication, repair the candidate and rerun every affected gate. An
unpushed incorrect tag may be deleted and recreated only after correcting the
candidate/signing setup and repeating post-tag verification. After a public
release, never silently replace assets, move the tag, or recreate its signature:
document the problem and publish a corrected patch release. Vulnerability
handling follows [`SECURITY.md`](SECURITY.md).

## 📚 Related authorities

- [`docs/RELEASE_SEMANTICS.md`](docs/RELEASE_SEMANTICS.md) — exact version/changelog/history semantics
- [`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md) — evidence and asset-manifest schema
- [`docs/RELEASE_PROVENANCE.md`](docs/RELEASE_PROVENANCE.md) — build-record and signature trust policy
- [`INSTALLATION.md`](INSTALLATION.md) — clean install/upgrade validation
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — change/review workflow
- [`SECURITY.md`](SECURITY.md) — vulnerability handling
- [`docs/README.md`](docs/README.md) — complete documentation authority map

---

**Release principle:** certify one exact source revision, derive artifacts from it
once, and publish only evidence-backed output bound to that revision.
