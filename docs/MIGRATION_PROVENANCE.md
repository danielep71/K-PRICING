# KPR migration provenance

Issue [K-PRICING #11](https://github.com/danielep71/K-PRICING/issues/11).
Verified on 2026-09-24 against the immutable source, before importing VBA.

## Source and destination boundary

- Source: [danielep71/KPR](https://github.com/danielep71/KPR).
- Selected commit: `f26450d1fa7b11261162e901dedba062f21c99a7`; unchanged from the preparation plan.
- Tree: `013653658c9c9fb1e61e5958d2d4277ab8c44191`.
- Destination base: `77983a398620805bd19469a3d6b18a29d52372cb`.
- This step imports only the date-layer and VBE export contracts. Production
  modules, regression harness and specialist checker remain planned imports.
- Source date-contract v0.0.2 is independent of the destination migration
  milestone v0.0.2 and destination VERSION 0.0.0. No release is created.

## Verified source bytes and path map

Git blob IDs and SHA-256 below are independently calculated from `git show`
bytes at the selected commit, before working-tree newline conversion. All ten
blob IDs match the preparation inventory. These are source hashes, not hashes
of the editorially adapted destination contracts or future imported modules.

| Source path | Git blob SHA-1 | SHA-256 of source Git bytes | Destination or disposition |
| --- | --- | --- | --- |
| `docs/DATE_LAYER_CONTRACT.md` | `6984e9354cd2a32986f6504ac7ae4f95b0a472e2` | `13bf70fe3eacd97ff2579142cce21423083cd86bfa5cce44a01e8fba2c3f6bc8` | docs/DATE_LAYER_CONTRACT.md |
| `docs/IMPLEMENTATION_PLAN.md` | `0355dc71cd59d9dab08e11d456172d922e7b38be` | `9ef26f8bc3df055a37fae4901356874fd5d334ad18c46285d95758984d09aace` | Reconcile into this migration plan and issue register |
| `docs/VBE_EXPORT.md` | `f37eba3de6f3feb28a0d686bbf68b5fea5ff044e` | `a97e7cc742d544e86304d5d80b31f991d7537d760b218969f42893ee39d93ffd` | docs/VBE_EXPORT.md |
| `src/modules/KPR_Core_Array.bas` | `a6c98c192dc2746a59a418bf2190a40f1c13c233` | `6b9468db0d7a9be9d41a331cc6d8be302e508136ec80f3d32d5c8515de4235b4` | src/core/KPR_Core_Array.bas |
| `src/modules/KPR_Core_Dates.bas` | `7e0929474e1ae1311df2081eb7a64fe126f1b688` | `a28f3c8c0af78c4a06236a8553e7db95c76406fe406f34911a0a4221fbcf5f94` | src/core/KPR_Core_Dates.bas |
| `src/modules/KPR_Core_Err.bas` | `a6ac71c00979dd9dd13ed01264ae5a375da14c28` | `91b6c613da8939aab60da84dbd614434829ded62f4b0096cc1cff95a364ed5d3` | src/core/KPR_Core_Err.bas |
| `src/modules/KPR_Core_Parse.bas` | `ec08bbb0056cfe6e27ed1115f3d6431e08a37b98` | `8b5e252554fb9786b4ac2513c56e6bc614c6eaa19483222159df9c347ea7826f` | src/core/KPR_Core_Parse.bas |
| `src/modules/KPR_DATES_DAYS.bas` | `37e997107c4c030d69bfa9b16d8aae3d22c97fc7` | `eb1d4ba1144c1f56fc06ac035cd408fddff5635f49428fe1fc16444d634c3985` | src/modules/KPR_DATES_DAYS.bas |
| `test/modules/KPR_REGRESSION_TESTS.bas` | `b339d4b932f390c143973fe5fcc106e79ffbcad1` | `5ab84d316390eb2b3c61eb805add87938ba9001c5d327b279a3906934f7939de` | tests/modules/KPR_REGRESSION_TESTS.bas |
| `tools/check_repo.py` | `80a5d078d9bb8f7114132e3985b6997f80a95517` | `5379ed7e044705e2fa611116b6a65636de039cf051056f56ce596f0bbf098e17` | tools/check_kpr_contract.py (adapt; keep generic checker) |

The [migration plan](MIGRATION_PLAN.md) owns the rest of the source-file
inventory and disposition. Detailed source implementation-plan reconciliation
belongs to issue #16; the source plan is not copied as a competing authority.

## Contract preservation and adaptations

The [date-layer contract](DATE_LAYER_CONTRACT.md) preserves source behaviour, including amendments, defaults, errors, caller
rules, date window and array semantics. Editorial adaptations include a migration-scope notice,
fully qualified source issue references and the implementation-plan link
redirected to the destination migration plan, plus the explicit registry
clarification below. All 22 declaration signatures
in its public-surface block match the source facade declarations after joining
VBA line continuations and normalizing whitespace, including whitespace after
an opening parenthesis; the KPR_ namespace is unchanged. The destination
`PUBLIC_API.txt` now records exactly the same 22 supported facade functions;
internal core and regression members remain project infrastructure rather than
supported calculation API.

The [VBE export contract](VBE_EXPORT.md) preserves format and round-trip
requirements. Editorial changes add source/destination evidence boundaries,
qualify the source certification issue, map source `test/` to `tests/` and
`demo/` to `examples/`, and split production destinations into core and facade.
KPR-specific static rules are retained additively in
`tools/check_kpr_contract.py` and wired into Repository integrity; generic
K-PRICING repository, VBE, API-manifest and release controls remain authoritative
for their existing scopes. References to MacroOptions and demo infrastructure
remain requirements for future source-roadmap work, not claims of existing
implementation.

## Reviewed source-contract clarification

PR review identified an inherited contradiction: section 3.2 explicitly gives
range checking precedence over integrality, but the section 7 registry limited
INTEGER_RANGE to integral values. The destination registry now includes all
out-of-range numeric values, including fractions; INTEGER_FRACTION is explicitly
limited to values inside the Long range in both the matrix and registry.
This matches the frozen KPR_Core_Parse.TryParseLongScalar implementation, which
checks the Long bounds before its Int comparison. For example, 2147483648.5
therefore yields INTEGER_RANGE/#NUM!, as section 3.2 already requires. This is
a documented source wording correction, not a changed algorithm or new runtime
result. Source files remain untouched; downstream fixture review belongs to
issues #13–#14 and source-history reconciliation to #16.

## Licence and attribution

Source is MIT, copyright (c) 2026 Daniele Penza. Source LICENSE blob
`51f4dbcee64abe24d6f607a9a7a8f491e11b5c57` is byte-identical to the destination
[LICENSE](../LICENSE); its copyright, permission and warranty text are retained.
Original contract URLs are preserved in each imported document's notice.

## Historical runtime evidence

[KPR #17's recorded result](https://github.com/danielep71/KPR/issues/17#issuecomment-5514012069)
reports `KPR_Tests_Run`: 557 checks, zero failures, and
`KPR_Tests_RunArray`: seven checks, zero failures with dynamic-array API
SUPPORTED, at `a750cd5a935529b807c43631e92cd9f1b15ee8b3`.
This verifies what the source record reports; it is not a new execution or
certification of the later frozen source. It describes one dynamic-array host
and focused 1900/1904 behaviour. Source final certification remains owned by
[KPR #29](https://github.com/danielep71/KPR/issues/29).

## Destination source import status

The four core modules are imported byte-for-byte from the frozen source: their
destination Git blob IDs remain the same as the source inventory.

`KPR_DATES_DAYS.bas` has one documented format-only destination delta required
by K-PRICING's committed-whitespace gate: seven trailing spaces on
section-banner comment lines were removed. Its source blob is
`37e997107c4c030d69bfa9b16d8aae3d22c97fc7`; the normalized destination blob
is `07656d21f01770e2d74d9de77d5f956949ff1301`. No executable statement,
declaration, attribute, signature, default, error rule or algorithm changed.

`KPR_REGRESSION_TESTS.bas` has two destination-only test-infrastructure
adaptations identified during review. First, the two `CLngLng` regression
cases are guarded by `#If Win64 Then` rather than the frozen source's
`#If VBA7 Then`. On 32-bit Office with VBA7, `VBA7` is true while
`CLngLng` is unavailable, so the source guard can prevent the regression
project from compiling. Second, K-PRICING adds the public macro
`KPR_Tests_RunEvidence`. It delegates the actual pure regression execution to
`KPR_Tests_RunAll("all")` and translates only the returned count/failure
result into the stable `CASE=/CASES=/ASSERTIONS=/FAILURES=/RESULT=` format
required by the destination retained-evidence validator. It does not add test
cases, duplicate suite dispatch, change expectations or touch production code.

The destination harness blob after the PR #28 adaptations is
`a67bda4bd15d00ddf564efb5ce353f31d9882055`. Issue #32 then adds two focused
pillar-range assertions for conversion overflow, aggregate overflow and grammar
precedence including later grammar errors after an overflowing component,
taking the retained-evidence totals to 566 assertions on 32-bit Office and 568
on 64-bit Office; the two-count difference remains the
LongLong cases compiled only under `Win64`.

Issue #32 also corrects one inherited production defect in
`KPR_Core_Dates.TryPillar_Parse`: a grammatically valid digits+unit component
whose numeric quantity cannot be represented as `Double`, or whose Y/M or W/D
aggregate overflows `Double`, is now classified as
`PILLAR_AGGREGATE_RANGE` rather than falling through as
`PILLAR_TOKEN_MALFORMED`. Aggregate values are computed into local temporaries
before the ByRef outputs are assigned, preserving the parser's
outputs-on-success-only contract on failure. This is a destination-only
behavioral correction required by the already-migrated date contract; the
frozen source repository remains unchanged.

After the #32 correction, the destination blob IDs are
`fedc43b3e38f706d528186cf09d3af12086f52f9` for
`src/core/KPR_Core_Dates.bas` (source baseline blob
`7e0929474e1ae1311df2081eb7a64fe126f1b688`) and
`c44101222bbd1f5a7de58452ee53c5fb27e9ea57` for
`tests/modules/KPR_REGRESSION_TESTS.bas` (source baseline blob
`b339d4b932f390c143973fe5fcc106e79ffbcad1`).

Repository-path adaptations remain `src/modules/KPR_Core_*` to
`src/core/KPR_Core_*` and `test/` to `tests/`. The neutral starter modules
were removed, and a new minimal `KPR_DateExample` consumer was added as
destination-only example code.

**Destination runtime verification remains pending** in
[K-PRICING #17](https://github.com/danielep71/K-PRICING/issues/17).
The accepted neutral-starter run remains setup history and is not evidence for
the migrated date layer.
