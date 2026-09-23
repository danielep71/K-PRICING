# 🔐 Release Provenance

[![Contract: 1.2.0](https://img.shields.io/badge/contract-1.2.0-217346)](TEMPLATE_CONTRACT.md)
[![Digest: SHA-256](https://img.shields.io/badge/digest-SHA--256-1D76DB)](#-minimum-levels)
[![Signatures: policy-selected](https://img.shields.io/badge/signatures-policy--selected-6f42c1)](#-signature-policy)

This is the authority for build records and signature verification under template
contract 1.2.0. The existing [release evidence contract](RELEASE_EVIDENCE.md)
continues to own runtime checks and the SHA-256 asset manifest. The release gate
calls `tools/release_provenance.py`; there is no second publication verdict.

## 📋 Minimum Levels

| Profile | Source-only distribution | Generated binary assets |
| --- | --- | --- |
| `library` | Base release evidence; empty `assets`; no manifest or build record required | Disallowed by the base release policy |
| `template` | Base release evidence, including pilots and governance; canonical Git-tag signature required once the tag exists | Disallowed by the base release policy |
| `ui-component` | Base release evidence; no artificial workbook required | Base evidence, SHA-256 manifest and build record |
| `application` | Base release evidence; no artificial workbook required | Base evidence, SHA-256 manifest and build record |

Selecting SSH verification for the **provenance record** raises every
distribution's minimum to a signed build record, including source-only releases.
That selection is explicit in `.github/release-policy.json`; omission, an
unsupported value, or disagreement with `.github/release-provenance.json` fails
closed. The canonical template's Git-tag signature is a separate authenticity
control and does not change whether an external provenance record is required.
Contracts 1.0.0 and 1.1.0 retain their previous evidence rules; advanced inputs
require 1.2.0. Unsupported future contracts fail instead of inheriting today's
policy.

`dist/` is the complete payload boundary. Its regular files must exactly equal
the evidence asset list and digests. Missing, additional, changed, duplicate or
symlinked payloads fail. Keep reports, provenance and signatures outside `dist/`.
For source-only releases, `dist/` must be empty or absent.

## 🧭 Trust Policy

Commit `.github/release-provenance.json` before freezing the candidate. The
default identifies `.github/workflows/static-checks.yml` in the candidate's own
repository. `@repository` and `@candidate` resolve to the committed repository
identity and exact release SHA. They are policy references, not initialization
placeholders.

Set `workflow.path` to the workflow responsible for the retained release
validation or build record. For a reusable provider, set `workflow.repository`
and `workflow.sha` to its actual repository and full immutable commit. Do not
use a branch, tag, or a workflow that did not participate. Record the actual
invocation's run ID and attempt. A local default workflow must exist in the
candidate; a remote provider's content and run must be reviewed separately.

The gate reads policy and any committed provenance-signing keys from candidate
Git objects. Editing the working copy or an external record cannot disable
required verification. The reviewed candidate SHA and trusted verifier
installation are trust roots: review changes to that policy as carefully as
changes to the workflow itself.

The same policy file also owns `tag_signature`, which has separate `generated`
and `template` scopes. Generated repositories use `mode: "none"` by default and
therefore retain the existing annotated-tag contract unless they deliberately
adopt an equivalent stronger local control. The canonical template uses:

```json
{
  "mode": "ssh-github",
  "principal": "danielep71",
  "github_user": "danielep71"
}
```

When the canonical local release tag exists, the verifier reads the public SSH
**signing-key** registry for the configured GitHub user, constructs a temporary
OpenSSH allowed-signers file using the configured principal, and invokes
`git verify-tag` with SSH verification. No private key, token, or signing
credential is read from the repository. The registry lookup is deliberately
live current-trust evidence; inability to retrieve it or the absence of usable
registered signing keys is non-green.

## 🔏 Signature Policy

Three independent SSH signature controls may exist in the canonical release
process. They must not be conflated.

### Git-tag signature

The canonical template's `tag_signature.template.mode = "ssh-github"` requires
an SSH-signed annotated Git tag once the tag exists locally. Verification proves
both that the tag resolves to the certified candidate and that its signature is
valid under a public SSH signing key currently registered to the explicitly
trusted GitHub account. An unsigned tag, a different key, a damaged signature,
a moved/recreated tag, or unavailable current trust fails publication.

Pre-tag candidate validation remains possible and network-independent: the tag
signature check begins only once `refs/tags/<release-tag>` exists. The
post-tag `check_release.py --require-tag-ref` invocation is therefore the
publication boundary. Existing published v1.2.0 history remains immutable and
is not retrofitted.

### Durable certification-bundle signature

The canonical template additionally signs the deterministic durable
certification ZIP described in [RELEASE_EVIDENCE.md](RELEASE_EVIDENCE.md).
That signature uses the same externally held signing identity and current
GitHub signing-key trust source as the canonical Git tag, but a distinct
`excel-vba-release-certification` SSH namespace. The exact ZIP and its detached
`.zip.sig` file are retained together as GitHub Release evidence and are
re-verified after publication. This authenticates the certification ZIP bytes;
it does not authenticate the Git ref and does not replace either the tag
signature or an enabled provenance-record signature.

### Provenance-record signature

`.github/release-policy.json` is the selector for the **external provenance
record** signature requirement. Its top-level `provenance_signature_mode` field
is mandatory and supports exactly:

- `"none"` — detached provenance signing is disabled; a supplied provenance
  signature is rejected rather than silently ignored;
- `"ssh"` — a provenance record and detached OpenSSH signature are mandatory,
  including for source-only releases.

`.github/release-provenance.json` supplies the candidate-bound trust details and
must declare the **same** mode in its `signature` object. The duplication is
intentional: the release policy selects the required assurance level, while the
provenance trust policy supplies the verification configuration. Neither file may
silently raise or lower the other. A missing selector, unsupported selector, or
mode mismatch is a blocking release finding.

The current canonical baseline selects `"none"` explicitly for the provenance
record. That does not weaken or satisfy the separate canonical Git-tag signature
requirement. Conversely, a valid provenance-record signature cannot substitute
for an unsigned or untrusted canonical release tag.

## 🧾 Build Record

After producing and testing assets, finalize the external evidence JSON and
manifest. Hash their exact bytes with SHA-256, then create an external UTF-8
`release-provenance.json`. Replace every example value with observed facts:

```json
{
  "schema_version": 1,
  "repository": "owner/project",
  "candidate_sha": "0123456789abcdef0123456789abcdef01234567",
  "template_contract": {"version": "1.2.0", "source": "owner/template"},
  "profile": "application",
  "tag": "v1.0.0",
  "distribution": "binary",
  "digest_algorithm": "sha256",
  "evidence_sha256": "REPLACE_WITH_EXACT_EVIDENCE_DIGEST",
  "manifest_sha256": "REPLACE_WITH_EXACT_MANIFEST_DIGEST",
  "assets": [{"path": "dist/project.xlsm", "sha256": "REPLACE_WITH_ASSET_DIGEST"}],
  "workflow": {
    "repository": "owner/project",
    "path": ".github/workflows/static-checks.yml",
    "sha": "0123456789abcdef0123456789abcdef01234567",
    "run_id": 123,
    "run_attempt": 1
  },
  "environment": {
    "os": "Windows 10",
    "architecture": "x64",
    "host": "Microsoft Excel",
    "host_version": "16.0",
    "office_bitness": "64",
    "runtime": "VBA7+",
    "builder": "Controlled workstation identifier",
    "procedure": "Retained build log identifier and exact import/package procedure"
  }
}
```

The shape is exact; duplicate JSON keys fail. Copy `template_contract` from the
candidate. List assets in path order. For optional source-only records use
`assets: []`, `distribution: "source-only"`, and `manifest_sha256: null` if the
manifest is omitted. All environment fields must be nonempty; use an explicit
`not applicable: source-only` explanation for unused Office/build fields.
Keep runtime test environments in the separate base evidence even if identical.

## ✍️ Optional Provenance-Record Signatures

This implementation supports detached OpenSSH signatures for external provenance
records, without a signing-service dependency. It does not claim SLSA or verify
vendor attestations. Use an approved release signing key kept outside the
repository. Commit its public key in an allowed-signers file, for example
`.github/release-signers`:

```text
release@example.org ssh-ed25519 REPLACE_WITH_APPROVED_PUBLIC_KEY
```

To enable provenance-record signing, change the release selector before freezing
the candidate:

```json
{"provenance_signature_mode": "ssh"}
```

and configure the matching provenance trust policy:

```json
{"mode": "ssh", "principal": "release@example.org", "allowed_signers": ".github/release-signers"}
```

Both changes belong in the reviewed candidate. Changing only one file is a
policy mismatch and fails before signature verification.

Sign the finalized record with the dedicated namespace:

```bash
ssh-keygen -Y sign -f /secure/path/release-key -n excel-vba-release ../release-provenance.json
```

The gate invokes `ssh-keygen -Y verify` with the committed allowed signers,
configured principal, namespace and exact record bytes. See the
[OpenSSH manual](https://man.openbsd.org/ssh-keygen.1) for key and allowed-signers
formats. An absent tool, unsupported operation, timeout, missing signature,
wrong key/namespace, changed record, or policy mismatch blocks publication.
Supplying a provenance signature while both provenance policies say `none` also
fails: no signature is silently left unchecked.

## ✅ Verification and Retention

1. Check out the reviewed candidate; retain the exact source SHA and both release
   and provenance trust policies.
2. Stage only the approved downloadable payloads in `dist/`. Build/test them
   using the recorded environment and keep the actual logs.
3. Finalize base evidence, manifest and build record; sign the provenance record
   last if that independent control is enabled.
4. Create the release tag according to [`RELEASING.md`](../RELEASING.md), then
   run the integrated gate with `--require-tag-ref`. For the canonical template,
   this post-tag run verifies the SSH-signed tag through current GitHub signer
   trust. Omit `--provenance-signature` when provenance-record signing is
   disabled:

```bash
python3 tools/check_release.py --root . --tag v1.0.0 \
  --candidate-sha FULL_CANDIDATE_SHA \
  --evidence ../release-evidence.json \
  --asset-manifest ../release-assets.sha256 \
  --provenance ../release-provenance.json \
  --provenance-signature ../release-provenance.json.sig \
  --require-tag-ref
```

5. Retain the checked evidence, manifest, record, provenance signature if enabled,
   tag-verification gate output, build logs, and workflow run/attempt alongside
   the release. Recheck downloaded payloads using the same candidate and those
   exact bytes.

Checksums prove byte identity. A verified tag signature authenticates the trusted
signer's Git tag; a verified certification signature authenticates the durable
certification ZIP; a verified provenance-record signature authenticates the
approved signer's assertions in that separate record. Neither independently
proves that Excel imported the recorded source, that an environment description
is truthful, or that a named workflow ran successfully. Review exact source,
logs and outcomes separately. The default static workflow supplies validation
identity, not an automated Excel builder. Optional Windows/Excel automation is a
separate contract.

## ↩️ Rollback and Key Changes

If validation fails before publication, stop, correct the candidate or rebuild
assets, and regenerate every affected digest/signature. An incorrect local tag
that has not been pushed may be deleted and recreated only after the candidate
or signing setup is corrected and the complete post-tag verification is rerun.
After publication, never replace assets silently or move/recreate the tag;
publish corrections under a new version with fresh evidence.

For canonical Git-tag signer rotation, register the replacement public key on
GitHub as an SSH signing key before retiring the old one. Overlap allows a
controlled transition; sign and verify the next release with the replacement,
then remove the old key. For compromise, remove the affected public signing key
from GitHub immediately. Fresh verification of historical tags signed only by a
removed key then becomes non-green by design: the GitHub registry represents
**current** trust. Preserve the exact successful publication-time gate evidence
as historical evidence and never re-enable a compromised key merely to make an
old tag green.

Provenance-record signer changes are separate: when that optional control is
enabled, commit and review its allowed-signers rotation before freezing the new
candidate. An old candidate's embedded provenance key is historical trust, not
proof that the key remains approved today.

## 🧪 Validation Scope

`python3 tools/test_release_provenance.py -v` exercises synthetic candidates,
payload tampering, explicit unsigned provenance policy, missing/unsupported
selectors, both provenance policy-mismatch directions, and real ephemeral SSH
signatures. The signed-tag matrix separately proves a valid current trusted key,
unsigned tag, wrong key, corrupted signature, signed moved tag, signer overlap,
revocation, and generated-project non-inheritance. Tests use ephemeral private
keys only and never request or expose the maintainer's real signing key.

The suite runs with the existing release self-test in repository CI. It does not
build or execute Office files, use the maintainer's live signing credentials, or
publish a release. The actual canonical post-tag release check performs the live
public GitHub signing-key lookup when the release tag exists.
