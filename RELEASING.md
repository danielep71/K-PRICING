# 🚀 {{PROJECT_NAME}} Release Guide

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
| Project/profile | {{PROJECT_NAME}} / {{PROFILE_NAME}} |
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

Start from the repository's protected release path, freeze scope, and record the
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

<!-- template:remove:start -->
For changes to checker behavior in the canonical template, also run the
checker-development and semantic policy-coverage contracts documented in
[`docs/CHECKER_DEVELOPMENT.md`](docs/CHECKER_DEVELOPMENT.md).
<!-- template:remove:end -->

## 5. Certify in Excel

Use the exact candidate source in each advertised Excel environment:

1. import only candidate-controlled exports;
2. run **Debug → Compile VBAProject**;
3. execute the documented regression entry point;
4. run applicable UI/lifecycle/platform/manual checks; and
5. record environment, counts, failures, completeness and cleanup.

The neutral starter baseline is `ProjectTests.RunProjectTests`; until replaced by
the generated project's own contract it reports four cases, six assertions,
zero failures, complete execution and passing cleanup.

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

<!-- template:remove:start -->
For the canonical template, release certification also checks this convention
against the complete Git range from the previous release tag to the candidate.
An intentional history-preserving merge must be registered in
`.github/release-history-policy.json` before certification with the previous
`base_tag`, exact commit SHA, permitted finding type, GitHub issue/PR review
reference, and a concrete reason. Stale, out-of-range or overbroad exceptions
are blocking. This machine-check is template-maintainer policy; generated
repositories do not inherit it unless they deliberately adopt an equivalent
local control.
<!-- template:remove:end -->

An implementation PR may merge before release certification for stabilization.
That does not authorize tagging or publication. Keep incomplete acceptance work
open, then certify the final `main` commit before tagging. A squash or merge
creates a new source identity: retain original evidence attribution and obtain
the required final-candidate evidence rather than silently rebinding old results.

<!-- template:remove:start -->
**Historical exception:** the v1.1.0 release PR used squash merging. The v1.2.0
implementation entered `main` through PR #49 at
`ac78ddca5de9de1fbfbf89d504b8ba93b06220c4` using a history-preserving merge
during the move to stabilization on `main`. The later v1.2.0 clean-room
stabilization merge commits from PRs #68, #71 and #74 are retained in
`.github/release-history-policy.json` as descriptive historical records. Those
records preserve the published ancestry but do not exempt any future release
range. Keep all published v1.2.0 ancestry intact.
<!-- template:remove:end -->

Do not force-push shared `main` or rewrite published tags to make historical
merges conform retroactively. Apply this convention to future merges.

<!-- template:remove:start -->
### Canonical-template release certification

Before creating a release tag for the canonical template, complete and retain
all of these additional checks against the same exact candidate SHA:

- run `python3 tools/check_release_semantics.py --root . --self-test` and the
  current-tree release-semantics validation; the previous-release-to-candidate
  history range must contain no unapproved merge commit, duplicate commit
  subject, stale exception or overbroad exception;
- run the live external-link observation defined by
  [`docs/DOCUMENTATION_CHECKS.md`](docs/DOCUMENTATION_CHECKS.md); deterministic
  documentation defects must be zero, while restricted or transient network
  outcomes remain explicitly reported and are never converted to `PASS`;
- complete a clean-room maintainer journey from live GitHub template creation
  through initialization, live repository provisioning, Excel validation and a
  first release; retain numbered steps, any gaps/corrections, and elapsed-step
  evidence;
- export the complete Wiki from the exact candidate SHA, publish it, freshly
  clone/read back the publication, byte-compare it with zero drift, and record
  both the source SHA and resulting Wiki commit; and
- review the published Wiki in a browser, confirming Home, the sidebar and the
  complete page-navigation set render and navigate as intended.

These are tag blockers, not optional observations. A missing execution, an
unresolved deterministic defect, publication drift or an incomplete browser
review prevents tag creation. Network restrictions and transient failures remain
non-success observations until separately resolved or explicitly reported under
the documentation policy.
<!-- template:remove:end -->

## 9. Create and verify the release tag

Tag only the certified commit and run the release-integrity checker before and
after creating the local tag. The default generated-project contract remains an
annotated tag; generated repositories do not inherit the canonical SSH-signing
requirement unless they explicitly adopt an equivalent local policy.

### Initialized generated project

Use this sequence for an initialized generated project under the default tag
policy:

```bash
git switch main
git pull --ff-only
candidate_sha="$(git rev-parse HEAD)"
release_version="$(tr -d '\r\n' < VERSION)"
release_tag="v${release_version}"

python3 tools/check_release.py \
  --root . \
  --tag "$release_tag" \
  --candidate-sha "$candidate_sha" \
  --evidence ../release-evidence.json \
  --output test-results/release-integrity.json \
  --summary test-results/release-integrity.md

git tag -a "$release_tag" -m "{{PROJECT_NAME}} ${release_version}"

python3 tools/check_release.py \
  --root . \
  --tag "$release_tag" \
  --candidate-sha "$candidate_sha" \
  --evidence ../release-evidence.json \
  --require-tag-ref

git push origin "$release_tag"
```

<!-- template:remove:start -->
### Canonical-template SSH-signed tag

**Do not use the generated-project `git tag -a` command above for the canonical
template.** The maintainer must have an SSH signing key whose private half
remains outside the repository and whose public half is registered on GitHub
**as an SSH signing key** for the `github_user` named by the candidate's
`.github/release-provenance.json`. For the current canonical policy, that account
and verification principal are both `danielep71`.

Set `RELEASE_SIGNING_KEY` to the local private-key path. Do not commit the private
key, a copy of it, an agent socket, or signing credentials. Then run:

```bash
git switch main
git pull --ff-only
candidate_sha="$(git rev-parse HEAD)"
release_version="$(tr -d '\r\n' < VERSION)"
release_tag="v${release_version}"
: "${RELEASE_SIGNING_KEY:?set RELEASE_SIGNING_KEY to the canonical SSH signing private key}"

python3 tools/check_release.py \
  --root . \
  --tag "$release_tag" \
  --candidate-sha "$candidate_sha" \
  --evidence ../release-evidence.json \
  --output test-results/release-integrity.json \
  --summary test-results/release-integrity.md

git -c gpg.format=ssh -c user.signingkey="$RELEASE_SIGNING_KEY" \
  tag -s "$release_tag" -m "{{PROJECT_NAME}} ${release_version}"

python3 tools/check_release.py \
  --root . \
  --tag "$release_tag" \
  --candidate-sha "$candidate_sha" \
  --evidence ../release-evidence.json \
  --require-tag-ref

git push origin "$release_tag"
```

The authoritative post-tag gate retrieves only the configured GitHub account's
public SSH signing keys and builds a temporary OpenSSH allowed-signers file. It
requires all three facts before the tag can be pushed: the ref is an annotated
tag object, it resolves to the certified candidate SHA, and its SSH signature
verifies against current trusted signer material. Failure to read the registry,
no usable registered signing key, an unsigned tag, another key, a corrupted
signature, or a moved/recreated tag is blocking. Pre-tag candidate validation
remains network-independent because tag-signature verification starts only after
the local tag exists.

#### Canonical signer rotation and revocation

For planned rotation, register the replacement public key on GitHub as an SSH
signing key before retiring the old key. During the overlap, either registered
key can satisfy current trust; sign the next candidate with the replacement and
verify it through the release gate before removing the old key. For compromise,
remove the affected public key from the GitHub signing-key registry immediately
and stop release publication until a replacement is registered and verified.

The registry is deliberately a **current-trust** source. After a key is removed,
a fresh verification of an older tag signed only by that key becomes non-green.
Retain the exact successful release-gate evidence from publication as historical
evidence; do not re-add a compromised key merely to make an old verification
green.
<!-- template:remove:end -->

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

Create the release from the protected annotated tag. Include:

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
<!-- template:remove:start -->
- [ ] Canonical-template tag signature still verifies against the current trusted GitHub SSH signing-key registry.
- [ ] Canonical-template certification ZIP signature still verifies against the same current trusted GitHub SSH signing-key registry.
<!-- template:remove:end -->
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
