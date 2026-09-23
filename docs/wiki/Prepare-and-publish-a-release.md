# 🚀 Prepare and publish a release

> **Guide, not policy:** [RELEASING.md](../../RELEASING.md) is authoritative.

**Working directory:** the project's clean repository root.
[RELEASING.md](../../RELEASING.md) owns the sequence;
[Release Evidence](../RELEASE_EVIDENCE.md) owns schemas and required checks.

## 1. Choose the version and freeze source

Replace the initializer's `0.0.0` sentinel with the intended nonzero version.
Move the relevant Unreleased changes into a dated release section and update
comparison links. Use the date on which that reviewed release section is
**cut/frozen for the candidate**. The changelog date is not required to equal the
later tag-creation or GitHub Release publication date. The first release has no
previous-tag diff; review the complete initial project tree instead.

Merge the release changes through the protected process. If that merge changes
the candidate SHA, certify the merged commit before tagging. From the exact
candidate:

```bash
git switch main
git pull --ff-only
git status --short
candidate_sha="$(git rev-parse HEAD)"
release_version="$(tr -d '\r\n' < VERSION)"
release_tag="v${release_version}"
python tools/check_repo.py --root .
python tools/check_release_semantics.py --root .
```

Stop on a dirty tree, a sentinel version, failed gates or an unreviewed delta.
The release-semantic gate verifies Gregorian cut/freeze dates and rejects a
newer release whose cut/freeze date moves backward relative to an older release;
same-day releases remain valid.

## 2. Collect actual evidence

Retain a successful hosted integrity run for this candidate. Import the exact
exports into Excel, compile/run the regression suite and collect the full host
output. Perform the selected profile's additional checks from
`.github/release-policy.json`. A four-case starter PASS is not proof of an
application's startup, upgrade or packaging behavior.

Build any deliverable once from the candidate, reopen it and run its package
tests. Never hash a workbook and then edit it.

## 3. Prepare the external bundle

Keep evidence outside the candidate tree. Start with the schema in
[Release Evidence](../RELEASE_EVIDENCE.md) and populate every required check with
actual results, details and the same candidate SHA. Keep raw logs alongside it.
An empty example `checks` object will fail; do not fill it with invented PASSes.

| Distribution | Additional files |
| --- | --- |
| Source-only | Evidence JSON; `assets: []`; no artificial workbook or manifest |
| Allowed UI/application binaries | Complete `dist/` payload, sorted SHA-256 manifest, asset/package evidence and contract-1.2.0 provenance record |
| Optional structured Excel interface | Host record plus retained import/compile/regression/cleanup logs; matching base evidence check |
| Policy requires signing | Detached signature and committed trust policy under the provenance contract |

Libraries have no allowed binary asset globs in the baseline. UI/application
profiles may release source-only, but still need their required assurance checks.
See [Release Provenance](../RELEASE_PROVENANCE.md) and
[Excel Evidence](../EXCEL_EVIDENCE.md) for the exact optional/required branches.

## 4. Validate, create the local tag, validate again

This is the source-only command. Add the asset/provenance/signature/Excel flags
from the authorities when that branch applies:

```bash
python tools/check_release.py --root . --tag "$release_tag" --candidate-sha "$candidate_sha" --evidence ../release-evidence.json
git tag -a "$release_tag" "$candidate_sha" -m "Release ${release_version}"
python tools/check_release.py --root . --tag "$release_tag" --candidate-sha "$candidate_sha" --evidence ../release-evidence.json --require-tag-ref
```

Run each line only if the preceding line passed. The second check verifies the
annotated tag and target. If it fails, do not push. Repair the unpublished
candidate through review and regenerate affected evidence.

## 5. Publish and retrieve

```bash
git push origin "refs/tags/${release_tag}"
```

In GitHub Releases, draft the release from that existing tag. Summarize behavior,
installation, compatibility, supported environments and known limitations.
Upload the tested assets, checksums and retained evidence as appropriate, then
publish. Do not let the release form create a different tag target, and do not
rewrite the changelog date merely because publication occurs later than the
release-section cut/freeze date.

Download the published assets, compare hashes, verify the tag target and follow
installation in a clean environment. For source-only releases, inspect the source
archive and its expected files. Record the result before announcing availability.
Never move a public tag or silently replace an asset to hide a defect; follow
the documented patch-release recovery procedure.

---
[Home](Home.md) · [Previous](Develop-safely.md) · [Next](File-and-directory-reference.md)
