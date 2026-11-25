# 📚 Sources

Every rule in the [calculation contracts](CALCULATION_CONTRACTS.md) cites an
entry of this register by its identifier and a locator, such as an article,
paragraph, table or template cell. A rule without a verified source is a
documented assumption, and its contract says so.

## Register

| ID | Title | Issuer | Version or date | Locator | Local copy | SHA-256 of the copy | Status | Verified by | Verified on |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| *Illustrative:* `SRC-01` | Example regulation | Example authority | 2026-01-01 | Official publication reference or URL | — | — | `unverified` | — | — |

- **ID:** `SRC-` followed by a number, never reused. An independent
  implementation used to compute reference values, such as another author's
  tool, is registered too, with an `IMPL-` ID
  ([test cases](TEST_CASES.md#reference-classes)).
- **Version or date:** the consolidated version, publication or application
  date that the workbook implements.
- **Locator:** where the issuer publishes it, precise enough to find the same
  version again.
- **Status:** `unverified`, `verified`, or `superseded by SRC-NN`.

## Local copies

A copy of a source, such as a regulation, a reporting template or the
issuer's own validation rules, is stored byte for byte under
`docs/methodology/sources/`. Git neither normalizes those files nor reports
a white-space error in them, and editors leave them unchanged. A `README.md`
in that directory, by convention, is ordinary documentation that Git and
editors treat as text; the register stays here.

- Name each copy `<ID>_<short-title>.<extension>`, for example
  `SRC-01_example-regulation.pdf`.
- Record its SHA-256 in the register (`sha256sum`, or `certutil -hashfile
  FILE SHA256` on Windows).
- Store a copy only if its terms allow redistribution: the copy goes with
  every clone of the repository.

## Verifying an entry

1. Open the source from the issuer's own publication, not from a copy.
2. Confirm the version or date, and that it applies on the reporting dates
   the workbook supports.
3. If there is a local copy, recompute its SHA-256 and compare it with the
   register.
4. Check each locator that a calculation contract cites against the text.
5. Set the status to `verified`, with your name and the date.

When a source is replaced, keep its row, set it to `superseded by` the new
entry, and update every contract that cites it in the same change.

## Decisions

| Date | Decision | Reason | Issue |
| --- | --- | --- | --- |
| *Illustrative* | Which version of a source the workbook implements | Why that version | Link |
