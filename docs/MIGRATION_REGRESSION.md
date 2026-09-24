# KPR migration regression and parity protocol

Issue [#14](https://github.com/danielep71/K-PRICING/issues/14) defines how the
migrated date-layer regression harness is executed and how source-versus-
destination parity is judged. Actual Windows Excel execution is retained under
[#17](https://github.com/danielep71/K-PRICING/issues/17); this document defines
the method and evidence boundary and does not claim that a run has occurred.

## Baselines and evidence boundary

The frozen source baseline is
[danielep71/KPR@`f26450d1fa7b11261162e901dedba062f21c99a7`](https://github.com/danielep71/KPR/tree/f26450d1fa7b11261162e901dedba062f21c99a7).
The destination candidate is the exact K-PRICING commit selected by #17.
[Migration provenance](MIGRATION_PROVENANCE.md) owns the source/destination
inventory, hashes and justified source deltas.

Historical KPR results are context, not destination evidence. KPR issue #17
records a 557-check pure run and a seven-check dynamic-array run at
`a750cd5a935529b807c43631e92cd9f1b15ee8b3`; those results neither execute
the later frozen source nor certify K-PRICING.

For a destination host run, the retained record follows
[EXCEL_EVIDENCE.md](EXCEL_EVIDENCE.md). The validated pure-run entry point is
`KPR_Tests_RunEvidence`. It delegates to `KPR_Tests_RunAll("all")`, so the
adapter formats evidence but does not own test dispatch or expectations.

## Entry points

| Entry point | Invocation | Role | Evidence use |
| --- | --- | --- | --- |
| `KPR_Tests_Run` | Alt+F8 or direct VBA, no arguments | Runs all 12 pure suites and prints the native report | Human-readable smoke and diagnostic output |
| `KPR_Tests_RunSuite "suite"` | Direct VBA / Immediate window | Runs one named pure suite and prints the native report | Focused investigation |
| `KPR_Tests_RunAll([suite])` | Direct VBA or programmatic call; optional suite name | Returns a 2-D summary/failure array; omitted/Empty means all suites | Common source/destination parity interface |
| `KPR_Tests_RunEvidence` | Macro / direct VBA, no arguments | Destination-only adapter over `RunAll("all")`; emits ordered `CASE=`, counts and `RESULT=` lines | Retained K-PRICING host evidence |
| `KPR_Tests_RunMigrationEvidence` | Macro / direct VBA, no arguments | Common migration instrumentation; emits deterministic direct/host/shape/array `OBS` records plus cleanup status | Imported unchanged into both source-production and destination observation hosts |
| `KPR_Tests_RunHost` | Macro-only | Creates a scratch workbook; tests worksheet caller/date-system behavior and state restoration | Additional parity/certification log |
| `KPR_Tests_RunShape` | Macro-only | Creates a scratch workbook; tests real Range shapes, errors, UsedRange independence and state restoration | Additional parity/certification log |
| `KPR_Tests_RunArray` | Macro-only | Uses the late-bound dynamic-array API; tests spill behavior and 1904 call-level refusal | Additional parity/certification log |

The three stateful runners are deliberately outside `RunSuite("all")`.
A passing `KPR_Tests_RunEvidence` record therefore does not prove that the
host, Range-shape and dynamic-array macro runners were executed.
`KPR_Tests_RunMigrationEvidence` is the migration-only observation adapter.
It emits stable `OBS<TAB>id<TAB>payload` and
`CLEANUP<TAB>runner<TAB>status` records for the three stateful runners and
representative direct/worksheet probes. Those records serialize passing
observations as well as failures, including native error numbers, value/type
representations, dimensions and state-restoration outcomes. It deliberately
does **not** make pure-suite check counts the parity oracle: the frozen source
and destination regression harnesses have justified test-infrastructure and
correctness differences. Pure-suite results remain separately retained and
bound. The adapter is test instrumentation, not supported production API.

For parity, #17 imports the **same candidate regression module** containing that
adapter into both observation hosts. The frozen source's own regression module
is still compiled and retained unchanged in the separate exact-source compile
host; it is not rewritten to produce destination evidence. A missing
dynamic-array API remains a non-passing limitation rather than a successful
skip.

`KPR_Tests_RunAll` preserves the source harness failure semantics: a recognized
suite returns a 2-D array whose first row contains `checks: N` and
`failures: N`, with one label/detail row per assertion failure. Assertions
continue after an ordinary failed case. An unknown suite name returns a plain
explanatory string rather than an empty pass, and an unexpected harness runtime
error likewise returns an explanatory string. `KPR_Tests_RunEvidence` treats a
non-array native result as evidence failure, retains each native failure
label/detail, and emits `RESULT=FAIL` whenever the native failure count is
nonzero. It therefore formats the source semantics; it does not redefine them.

## Pure-suite coverage

The dispatcher owns 12 suites in this exact order.

| Suite | Existing migrated coverage |
| --- | --- |
| `date-type` | Accepted/rejected scalar date types, native errors and object rejection |
| `date-text` | Exact ISO text, locale-shaped text, numeric-looking text and malformed/impossible dates |
| `date-window` | `1900-03-01 .. 9999-12-31` and adjacent/out-of-window inputs |
| `integer` | Integer subtypes, fractions, type rejection and range-before-integrality |
| `control` | Optional Boolean/default handling, scalar-only controls and invalid control forms |
| `boundary` | Public-facade window/result boundaries, propagation and composition |
| `mapper` | Condition-to-native-error mapping and refusal cases |
| `host` | Direct-VBA 1900 caller contract and value identity of library/propagated `#N/A` |
| `pillar` | Grammar, signed/duplicate-unit refusals, rounding modes, 3W cap and format/parse consistency |
| `surface` | Frozen 22-name surface, YearIn behavior, singular pillar name and lower-window boundaries |
| `shape` | In-memory scalar/1-D/2-D shapes, broadcasting, control unwrapping, unsupported forms, capacity and allocation |
| `parity` | Every value argument of all 21 value-taking functions, vectorized element-for-element against scalar calls |

The existing shape/parity coverage includes the exact 100,000-element acceptance
and 100,001-element refusal boundaries, shape mismatches, scalar expansion,
multi-cell optional-control rejection, mixed valid/blank/error arrays, 1x1
scalar wrappers and output orientation.

The separate macro runners add the worksheet-specific coverage that pure direct
VBA cannot establish: 1900/1904 worksheet callers and ordinary recalculation,
real Range row/column/rectangle/multi-area behavior, and dynamic-array spill
behavior including 1904 call-level `#N/A`.

## Known limitations and independent expectations

The migrated harness is a focused implementation candidate, not the unfinished
full KPR regression architecture.

- Independent generated fixtures from danielep71/KPR#19 are not implemented.
  Existing expected values are reviewed test expectations, but they are not a
  complete external oracle.
- The broader final runner/evidence and complete error/state matrix planned in
  danielep71/KPR#20 and #21 are not represented as completed work. The
  destination `KPR_Tests_RunEvidence` adapter is intentionally narrower: it
  formats the already-landed pure suites for K-PRICING evidence.
- Excel cross-oracles planned in danielep71/KPR#22 are not implemented.
- Platform support, Excel build compatibility and dynamic-array availability are
  claims only for environments actually executed in #17.
- No CSE/legacy multi-cell compatibility claim is made.

One inherited source-contract wording defect is already documented in
[MIGRATION_PROVENANCE.md](MIGRATION_PROVENANCE.md): an out-of-`Long` numeric
is `INTEGER_RANGE/#NUM!` even when fractional, because the frozen parser tests
range before integrality. The destination contract text was corrected to match
the frozen implementation. A parity match to contradictory prose would not be
accepted as correctness.

The frozen source regression harness also has a real 32-bit portability defect:
its two `CLngLng` cases are guarded by `VBA7`, although `CLngLng` is
available only on 64-bit VBA. K-PRICING changes those two test guards to
`Win64`. This is a test-only destination correction, not a production
algorithm change.

## Exact source-versus-destination parity procedure

Use separate clean workbook hosts in the same Windows/Excel environment. Never
import source and destination production modules into one VBA project: their
component names intentionally collide and Excel would rename duplicates.

1. Freeze and record the exact K-PRICING candidate SHA and the frozen KPR source
   SHA above. Reconcile the source/destination inventory against
   `MIGRATION_PROVENANCE.md`.
2. Use the same Excel version/build, Office bitness, Windows build, locale,
   references and macro policy for every parity host. Record start/finish
   timestamps and the operator/runner identity.
3. Use a **64-bit Office host** for exact frozen-source parity. The frozen
   source's `VBA7`/LongLong guard makes an unmodified 32-bit source regression
   project an invalid exact-source baseline. A 32-bit K-PRICING run may be
   retained separately as portability evidence.
4. **Exact-source compile host A:** import the frozen KPR production modules and
   frozen regression module exactly as inventoried, compile the complete project
   and retain the compile/native-run log. This host establishes that the source
   baseline itself was not rewritten for evidence collection.
5. **Source observation host B:** import the frozen KPR production modules, but
   use the exact K-PRICING candidate's `KPR_REGRESSION_TESTS.bas` as the
   measurement harness. Record that harness's SHA-256. This is an instrumentation
   host, not a claim that the destination harness belonged to KPR.
6. **Destination observation host C:** import the exact K-PRICING candidate
   production modules and the same candidate regression module used in host B.
   Compile it and retain the compile result.
7. In hosts B and C run `KPR_Tests_RunMigrationEvidence`. Retain its complete
   output. The adapter emits serialized passing/failing observations for
   host/date-system, Range shape/type/error behavior, representative direct and
   worksheet calls, dynamic-array behavior and cleanup/state restoration. The
   two observation streams must match exactly.
8. Exercise any additional representative worksheet/direct-VBA calls required
   by #17 that are not already serialized by the adapter. If such a probe is
   added, add it to the common instrumentation harness so both sides emit the
   same observation ID and payload format; do not compare handwritten notes.
9. Investigate every unexplained difference. The destination-only evidence
   adapter and the `Win64` test guard are test-infrastructure differences, not
   production parity differences. Any production observation difference remains
   open until explained and regression evidence is added.
10. Bind destination retained host evidence with `check_excel_evidence.py`.
    Then create `migration.json` and validate the whole parity bundle with:

    ```bash
    python3 tools/check_migration_evidence.py \
      --root . \
      --manifest /path/to/evidence/migration.json \
      --source-sha f26450d1fa7b11261162e901dedba062f21c99a7 \
      --destination-sha FULL_DESTINATION_SHA \
      --output test-results/migration-evidence.json
    ```

    The migration validator verifies both SHAs, the instrumentation digest,
    every required supplemental-log digest, the 64-bit parity environment,
    source/destination observation equality, the required baseline observation IDs
    (while allowing additional IDs emitted by the same common instrumentation),
    explicit PASS cleanup records for the host, shape and array runners, and the
    manifest's explicit register of
    known source/destination correctness differences. Issue #32 must remain
    named as a known source defect/destination correction rather than being
    hidden inside an equality exception.

## Minimum #17 evidence bundle

The retained bundle contains a validated `migration.json` manifest plus these
five bound artifacts:

- `source-exact-compile`: exact frozen-source import/compile/native-run log;
- `destination-compile`: exact destination candidate import/compile log;
- `source-observations`: complete `KPR_Tests_RunMigrationEvidence` output
  from frozen-source production modules plus the common candidate
  instrumentation harness;
- `destination-observations`: the same adapter output from destination
  production modules and the same instrumentation harness;
- `destination-host-record`: the JSON retained host record validated by
  `check_excel_evidence.py`.

The manifest identifies both repositories/SHAs, the candidate instrumentation
path and SHA-256, the common Windows/Excel environment, the path plus SHA-256 of
every artifact above, and an explicit `known_differences` register. For v0.0.2
that register must include #32: the frozen source can classify oversized valid
pillar quantities incorrectly, while the destination intentionally fixes the
contract. `check_migration_evidence.py` rejects missing, extra, stale,
symlinked or digest-mismatched files and rejects unequal common observation
streams or non-PASS cleanup records. Known differences are documentary
dispositions, not permission to suppress an observation mismatch. This makes
the supplemental parity evidence part of the acceptance boundary rather than
unbound files placed beside a valid host record.

K-PRICING's shared host record remains governed by
[EXCEL_EVIDENCE.md](EXCEL_EVIDENCE.md). The migration manifest is an additional
v0.0.2 parity binding; it does not expand the generic release-evidence schema or
turn unfinished KPR #19-#22 work into completed scope.

## Handover to later work

Issue #17 executes this protocol. Issue #16 maps the unfinished KPR fixture,
runner-expansion and cross-oracle work to future destination backlog. Issue #15
may describe only the behavior and environments that #17 actually verifies.

**Acceptance principle:** parity means the same observable behavior on the same
environment and inputs, with independent contract expectations still governing
known source defects; matching a source defect is not proof of correctness.
