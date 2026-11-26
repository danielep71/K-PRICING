# 📚 Sources

Every rule in the [calculation contracts](CALCULATION_CONTRACTS.md) cites an
entry of this register by its identifier and a locator, such as an article,
paragraph, table or template cell. A rule without a verified source is a
documented assumption, and its contract says so.

## Register

| ID | Title | Issuer | Version or date | Locator | Local copy | SHA-256 of the copy | Status | Verified by | Verified on |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| *Illustrative:* `SRC-01` | Example regulation | Example authority | 2026-01-01 | Official publication reference or URL | — | — | `unverified` | — | — |
| `SRC-02` | ISO 8601-1, Date and time — Representations for information interchange — Part 1: Basic rules | ISO | ISO 8601-1:2019 | https://www.iso.org/standard/70907.html | — | — | `unverified` | — | — |
| `SRC-03` | Date systems in Excel | Microsoft | Support article, read 2026-11-25 | https://support.microsoft.com/en-us/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487 | — | — | `unverified` | — | — |
| `SRC-04` | Excel incorrectly assumes that the year 1900 is a leap year | Microsoft | Microsoft Learn article, read 2026-11-25 | https://learn.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year | — | — | `unverified` | — | — |
| `SRC-05` | Application.Caller property (Excel) | Microsoft | Microsoft Learn VBA reference, read 2026-11-25 | https://learn.microsoft.com/en-us/office/vba/api/excel.application.caller | — | — | `unverified` | — | — |
| `SRC-06` | Workbook.Date1904 property (Excel) | Microsoft | Microsoft Learn VBA reference, read 2026-11-25 | https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.date1904 | — | — | `unverified` | — | — |
| `SRC-07` | Application.Volatile method (Excel) | Microsoft | Microsoft Learn VBA reference, read 2026-11-25 | https://learn.microsoft.com/en-us/office/vba/api/excel.application.volatile | — | — | `unverified` | — | — |
| `SRC-08` | Excel specifications and limits | Microsoft | Support article, read 2026-11-25 | https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3 | — | — | `unverified` | — | — |
| `SRC-09` | WEEKDAY function | Microsoft | Support article, read 2026-11-25 | https://support.microsoft.com/en-us/office/weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a | — | — | `unverified` | — | — |
| `SRC-10` | Dynamic array formulas and spilled array behavior | Microsoft | Support article, read 2026-11-25 | https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531 | — | — | `unverified` | — | — |
| `SRC-11` | Date data type (VBA) | Microsoft | Microsoft Learn VBA language reference, read 2026-11-25 | https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/date-data-type | — | — | `unverified` | — | — |
| `SRC-12` | Long data type (VBA) | Microsoft | Microsoft Learn VBA language reference, read 2026-11-25 | https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/long-data-type | — | — | `unverified` | — | — |
| `SRC-13` | Cell error values (Excel VBA) | Microsoft | Microsoft Learn VBA reference, read 2026-11-25 | https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/cell-error-values | — | — | `unverified` | — | — |
| `IMPL-01` | Excel worksheet functions `DATE`, `YEAR`, `MONTH`, `DAY`, `EOMONTH`, `EDATE`, `WEEKDAY`, `INT`, `MOD` and `IF` | Microsoft | Microsoft 365 Excel 16.0 build 20527, 64-bit, as run on 2026-11-26 | https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188 | — | — | `unverified` | — | — |

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
| 2026-11-25 | The date layer cites ISO 8601-1:2019 for the Gregorian calendar and the `YYYY-MM-DD` text form, and Microsoft's documentation, as read on 2026-11-25, for Excel's date systems, its limits and the VBA host; the Microsoft pages carry no version, so the date read is their version | ISO 8601-1:2019 is the edition in force; the Microsoft pages are the issuer's own statement of the host behaviour the functions rely on, and the [host behaviour](VALIDATION_PLAN.md#host-behaviour) entries measure it | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-25 | A date-layer rule with no external source, such as the supported window, the strict parsers, the condition registry, the array rules and the pillar grammar and rounding, cites `Decision (#4)` in its *Source* column: it is a documented design decision, recorded in the *Decisions* table of the guide that states it | No issuer defines them; registering the project's own decisions as a source would make the register cite itself | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-25 | `SRC-01` stays illustrative, and the date layer's entries start at `SRC-02` | The illustrative test cases still cite the illustrative entries until [#11](https://github.com/danielep71/K-PRICING/issues/11) replaces them, and an identifier is never reused | [#4](https://github.com/danielep71/K-PRICING/issues/4) |
| 2026-11-26 | `IMPL-01` is Excel's own worksheet functions, as the Excel oracle calls them, in the build that computed the date layer's `independent` values | Microsoft wrote them and the date layer's code never calls them, so they are an implementation by another author; a value is `independent` only for the build that computed it in a passing run | [#48](https://github.com/danielep71/K-PRICING/issues/48) |
