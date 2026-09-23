# 🔐 Release Evidence Contract

[![Binding: exact SHA](https://img.shields.io/badge/binding-exact%20SHA-217346)](#-evidence-json)
[![Profiles: 4](https://img.shields.io/badge/profiles-4-6f42c1)](#-policy)
[![Tags: template SSH-signed](https://img.shields.io/badge/tags-template%20SSH--signed-1D76DB)](#-commands)
[![Assets: digest verified](https://img.shields.io/badge/assets-digest%20verified-success)](#-source--only-and-binary-distributions)
[![Certification: retained release assets](https://img.shields.io/badge/certification-release%20assets-0969da)](#-durable-canonical-template-certification)

This contract defines the machine-readable evidence consumed by
`tools/check_release.py`. It binds version, changelog, tag, source checks,
Excel evidence, profile-specific assurance, and any staged binary assets to one
full candidate commit SHA. For the canonical template it also defines the
durable certification bundle published beside the GitHub Release and verified
again after publication.

## 🧩 Why Evidence Stays Outside the Candidate

Candidate-bound evidence is deliberately supplied separately from the candidate
Git tree. A committed file cannot contain the hash of the commit that contains
that file: changing the recorded hash changes the commit again. Build and retain
the final evidence outside the candidate source tree.

For canonical-template releases, the authoritative durable retention surface is
the **published GitHub Release asset set**, not a short-lived Actions artifact.
The deterministic certification ZIP, its standalone manifest, its SHA-256
file, and a detached SSH signature over the ZIP are uploaded to the Release and
are expected to remain available for the lifetime of that published release. The release closeout workflow downloads
those published bytes again and verifies them independently. Actions artifacts
may retain additional operational evidence, but their expiry is not the public
retention contract.

## 📜 Policy

`.github/release-policy.json` is the versioned release-evidence policy. Every
profile requires:

- `repository-integrity` with an HTTPS run URL;
- `vba-compile` with the tested Excel environment; and
- `regression` with its entry point, environment, positive case and assertion
  counts, zero failures, complete execution, and passing cleanup.

Additional required checks are profile-specific:

| Profile | Required evidence |
| --- | --- |
| `library` | Public API and caller-contract evidence |
| `ui-component` | UI state, cleanup, recovery, DPI/accessibility, and lifecycle evidence |
| `application` | Startup, shutdown, upgrade, recovery, packaging, and end-to-end smoke evidence |
| `template` | All three generated-profile pilots, live branch/tag governance evidence, the separately verified canonical Git-tag trust policy once the release tag exists, and the durable certification bundle described below |

All checks use `status: "PASS"`, a non-empty `detail`, and the same
`candidate_sha`. Project-specific checks may be added with lower-case,
hyphen-separated identifiers and the same three base fields.

## 🧾 Evidence JSON

The top-level object contains exactly these fields:

> **Example values:** replace the version, tag, candidate SHA, profile,
> distribution, checks, and assets with the exact release candidate values.

```json
{
  "schema_version": 1,
  "version": "1.0.0",
  "tag": "v1.0.0",
  "candidate_sha": "0123456789abcdef0123456789abcdef01234567",
  "profile": "template",
  "distribution": "source-only",
  "checks": {},
  "assets": []
}
```

The complete `checks` object depends on the selected profile. A regression
record has this mandatory shape in addition to `status`, `candidate_sha`, and
`detail`:

```json
{
  "entry_point": "ProjectTests.RunProjectTests",
  "environment": "host=Microsoft Excel; version=16.0; os=Windows (64-bit) NT 10.00; office=64-bit; runtime=VBA7+",
  "cases": 4,
  "assertions": 6,
  "failures": 0,
  "completeness": "COMPLETE",
  "cleanup": "PASS"
}
```

The record reports evidence; it does not manufacture it. Copy counts and
environment details from the exact Excel run, and retain its raw output beside
the checked JSON.

The optional [Windows/Excel interface](EXCEL_EVIDENCE.md) adds a structured host
record and retained-log validation. When adopted, declare an
`excel-host-evidence` check and supply `--excel-evidence`: neither can be omitted
while retaining the other. An unavailable host record cannot satisfy release
compile or regression evidence; the documented manual fallback remains valid.

## 📦 Source-Only and Binary Distributions

A library or template release is source-only by default. Set `distribution` to
`source-only`, keep the **product/runtime** `assets` array empty, and omit the
binary asset manifest. UI-component and application profiles may do the same.

Canonical-template certification attachments are not runtime/product assets.
The four `certification-vX.Y.Z.*` files described below are retained evidence
and therefore do **not** turn a source-only template release into a binary
distribution. The post-release closeout partitions those files from product
assets before applying `allowed_asset_globs`.

The template profile deliberately retains its registered placeholders,
template-only blocks, canonical repository identity, and construction history.
Its `repository-integrity`, `generated-profile-pilots`, and `live-governance`
evidence proves those controls are intentional and usable. Generated project
profiles remain subject to the stricter unresolved-token, template-identity,
and inherited-construction-history rejection rules.

Optional UI/application binaries must match an allowed `dist/` glob in the
policy. Each asset record contains exactly:

```json
{
  "path": "dist/example.xlsm",
  "sha256": "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
  "candidate_sha": "0123456789abcdef0123456789abcdef01234567",
  "package_test": "PASS"
}
```

Binary distribution requires a separate UTF-8 manifest, sorted by path, with
two spaces between each lower-case digest and safe repository-relative path:

```text
aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa  dist/example.xlsm
```

The checker requires exact equality among the evidence asset list, manifest,
and bytes staged under the candidate root. A digest establishes download
identity. It does not prove that a manually built workbook was produced from
the exported source; compile, regression, package-test, and build-process
evidence carry that separate claim.

## 🗄️ Durable canonical-template certification

`tools/release_certification.py` builds one deterministic, candidate-bound
certification package outside the candidate tree. The bundle must account for
these six roles:

| Role | Typical retained evidence |
| --- | --- |
| `excel-evidence` | Exact-source compile/regression log or structured host record |
| `external-links` | Final external-link observation/classification report |
| `gate-evidence` | Hosted/static/checker-development evidence selected for certification |
| `release-evidence` | The exact external release-evidence JSON consumed by `check_release.py` |
| `release-integrity` | The final pre/post-tag release-integrity report |
| `wiki-publication` | Exact-source Wiki publication/read-back evidence |

A record is either `public-file` or `restricted-reference`:

- `public-file` publishes the reviewed bytes inside the ZIP and records their
  archive path, size, and SHA-256. The input record must explicitly state
  `redaction_reviewed: true`; secret-like material is rejected.
- `restricted-reference` publishes **no protected bytes**. It records only the
  evidence identity/reference, its SHA-256, and the concrete reason public
  retrieval is unavailable. Restricted evidence must never be silently omitted
  or exposed merely to make the public bundle complete.

A minimal external specification has this shape:

```json
{
  "schema_version": 1,
  "repository": "owner/repository",
  "version": "1.2.1",
  "tag": "v1.2.1",
  "candidate_sha": "0123456789abcdef0123456789abcdef01234567",
  "records": [
    {
      "id": "release-evidence",
      "role": "release-evidence",
      "visibility": "public-file",
      "path": "release-evidence.json",
      "public_name": "release-evidence.json",
      "redaction_reviewed": true
    },
    {
      "id": "restricted-provider-record",
      "role": "restricted-provider-record",
      "visibility": "restricted-reference",
      "sha256": "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
      "reference": "provider-record:123",
      "reason": "provider access restriction"
    }
  ]
}
```

The complete specification must contain one unique record for every required
role. Additional restricted-reference roles are allowed when useful for the
release review.

Build twice only for deterministic comparison/testing; publish one accepted
output set. The canonical names are derived from the release tag:

```text
certification-v1.2.1.zip
certification-v1.2.1.zip.sig
certification-v1.2.1.manifest.json
certification-v1.2.1.sha256
```

Build and locally verify them after the candidate is frozen and the final
release evidence is available:

```bash
candidate_sha="$(git rev-parse HEAD)"
release_version="$(tr -d '\r\n' < VERSION)"
release_tag="v${release_version}"

python3 tools/release_certification.py \
  --root . \
  --build ../certification-input.json \
  --evidence-dir ../certification-evidence \
  --output-dir ../certification-output

python3 tools/release_certification.py \
  --verify-bundle "../certification-output/certification-${release_tag}.zip" \
  --manifest "../certification-output/certification-${release_tag}.manifest.json" \
  --digest "../certification-output/certification-${release_tag}.sha256"

: "${RELEASE_SIGNING_KEY:?set RELEASE_SIGNING_KEY to the canonical SSH signing private key}"
ssh-keygen -Y sign \
  -f "$RELEASE_SIGNING_KEY" \
  -n excel-vba-release-certification \
  "../certification-output/certification-${release_tag}.zip"

python3 tools/release_certification.py \
  --root . \
  --verify-signature "../certification-output/certification-${release_tag}.zip.sig" \
  --bundle "../certification-output/certification-${release_tag}.zip" \
  --candidate-sha "$candidate_sha"
```

The signing key is the same externally held SSH identity trusted for the
canonical release tag; the private key never enters the repository or
certification ZIP. Signature verification resolves the candidate's committed
`ssh-github` trust policy and the configured GitHub user's current public SSH
signing-key registry, then verifies the exact ZIP bytes in the dedicated
`excel-vba-release-certification` namespace.

The bundle contains its own canonical manifest; the same manifest is published
as a standalone Release asset. The `.sha256` file contains one canonical line
for the ZIP. The verifier requires the published manifest, internal manifest,
entry set, per-record sizes/digests, bundle filename, tag, and candidate binding
to agree.

### Publication and retention contract

For each canonical-template GitHub Release:

1. upload the four certification files above **unchanged** as GitHub Release
   assets, in addition to any separately allowed product/runtime assets;
2. do not rebuild, re-sign, rename, replace, or silently delete them after publication;
3. include the ZIP filename and SHA-256 in the release notes, plus a statement
   that the manifest, digest, and detached SSH signature are attached to the same Release;
4. treat the GitHub Release page and each asset's browser download URL as the
   stable public reference; and
5. retain those assets for the lifetime of the published release. Short-lived
   Actions artifacts may duplicate them operationally but are not the durable
   public copy.

If an evidence error is discovered after publication, do not replace the bytes
under the same immutable release claim. Document the defect and publish a
corrected patch release where necessary.

### Independent post-publication retrieval

The dispatch-only `Release closeout` workflow is the mandatory canonical
post-publication retrieval check. It reads the published Release asset metadata,
requires exactly the four certification names for the tag, separates them from
product/runtime assets, downloads each through the GitHub Release asset API, and
runs `release_certification.py` against the downloaded bytes. The closeout is
non-green if an asset is missing/duplicated/stale, the downloaded bytes differ
from the published digest/manifest, the detached signature does not verify
against the candidate's current trusted GitHub SSH signing-key registry, or the
embedded tag/candidate differs from the certified release.

The workflow retains the closeout snapshot and verification report as
short-lived operational evidence; the Release assets themselves are the durable
public evidence source.

## 🛠️ Commands

First certify the release checker and certification-bundle tool themselves:

```bash
python3 tools/check_release.py --root . --self-test \
  --summary test-results/release-self-test.md
python3 tools/release_certification.py --self-test
```

Before tagging, validate an external evidence file and any staged product
assets:

```bash
candidate_sha="$(git rev-parse HEAD)"
python3 tools/check_release.py \
  --root . \
  --tag "v$(tr -d '\r\n' < VERSION)" \
  --candidate-sha "$candidate_sha" \
  --evidence ../release-evidence.json \
  --output test-results/release-integrity.json \
  --summary test-results/release-integrity.md
```

Add `--asset-manifest ../release-assets.sha256` for a binary distribution.
Contract 1.2.0 also requires `--provenance ../release-provenance.json` for
binary distributions. Follow [RELEASE_PROVENANCE.md](RELEASE_PROVENANCE.md)
for build records, the optional provenance-record SSH signature, and the
canonical template's distinct Git-tag trust policy. Keep only deliverable
payloads in `dist/`: the gate rejects undeclared files there, including
non-binary files.

After creating the local tag, repeat the command with `--require-tag-ref`. For
all profiles that mode requires an annotated tag object resolving to the exact
candidate SHA. For the canonical `template` profile, the committed provenance
trust policy additionally requires the tag's SSH signature to verify against a
public SSH signing key currently registered to the explicitly trusted GitHub
account. Generated profiles retain annotation/target verification only unless
they deliberately adopt a stronger local policy.

The canonical tag signature, the detached certification-ZIP signature, and the
optional detached signature on an external provenance record are separate
controls. The tag authenticates the Git ref; the certification signature
authenticates the exact durable evidence ZIP; the optional provenance-record
signature authenticates assertions in that external record. They use dedicated
signature namespaces/purposes and none can silently satisfy another. Any
non-zero required verification blocks publication.

After the GitHub Release is published with the certification assets, run the
`Release closeout` workflow for that exact tag, candidate SHA, and milestone.
For the canonical template, successful durable-certification retrieval is part
of the terminal closeout verdict.

## 🚫 What the Gates Reject

The deterministic release self-test covers all four release profiles and
rejects:

- a version/tag mismatch or `0.0.0` sentinel;
- an invalid or missing dated changelog release;
- unresolved template syntax, template identity, or inherited construction
  history in a generated project release;
- a template release carrying a generated-project initialization record or
  missing its pilot/governance evidence;
- missing evidence or missing profile-specific checks;
- evidence or product assets bound to another candidate SHA;
- unapproved library or profile binaries;
- an absent product-asset manifest or incorrect asset digest; and
- a lightweight or moved tag.

The certification-bundle self-test additionally rejects missing required roles,
secret-like public evidence, non-deterministic output, a same-filename bundle
whose bytes no longer match the digest, incomplete/stale GitHub Release
certification asset sets, invalid/tampered detached certification signatures,
and downloaded Release bytes whose candidate/tag or manifest/digest contract
disagrees.

The provenance integration suite adds the canonical-template trust negatives:
an unsigned annotated tag, signature from the wrong public key, corrupted SSH
signature, signed tag moved to a different commit, and a tag whose former signer
has been removed from current trust. It also proves planned signer overlap and
that generated projects do not inherit the canonical signed-tag requirement.

An active, verified `v*` ruleset provides complementary server-side
immutability. This repository currently cannot enable branch/tag rulesets on
its account plan, as recorded in [setup verification](SETUP_VERIFICATION.md).
Protection must be resolved and verified before a functional release; no live
tag immutability enforcement is claimed by this setup milestone. A valid signature authenticates the tag signer; it does not
prove Excel/runtime behavior, build provenance, or the truth of external
evidence.

---

**Release principle:** certify one immutable candidate, retain its reviewable
evidence outside that candidate, and publish only output whose identity can be
independently re-verified later.
