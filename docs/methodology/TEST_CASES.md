# 🧪 Numerical test cases

How the numerical test cases are written: the files, the inputs and their
units, the expected results and their provenance, tolerances, case
categories, invalid inputs, and when a rule counts as validated. A case
tests one [calculation contract](CALCULATION_CONTRACTS.md), with inputs laid
out as the [data contract](DATA_CONTRACT.md) defines, against values whose
provenance the [source register](SOURCES.md) records. The
[validation plan](VALIDATION_PLAN.md) says which cases each calculation needs.

`python tools/check_test_cases.py` checks the format, and **Repository
integrity** runs it ([checking](#checking)). It never runs a calculation.

## Files

A case is one fixture and one expected-result file. The fixture takes one of
two forms:

| File | Holds |
| --- | --- |
| `tests/fixtures/<case-id>.json` | A JSON fixture: synthetic inputs as records, grouped by dataset |
| `tests/fixtures/<case-id>/` | A [file-based case](#file-based-cases): a `case.json` manifest and one directory per variant, holding input files |
| `tests/expected/<case-id>.json` | Expected results, each with its provenance |
| `tests/derivations/<name>.py` | Optional: a script that computes expected values from the documented method, separately from the code under test |
| `tests/accuracy/contracts.json` | Optional: the [accuracy contracts](#accuracy-contracts), each stating the worst error a calculation guarantees over a set of cases |
| `tests/accuracy/holdout_exposures.json` | Beside the contracts: the [holdout exposure record](#holdout-exposure-record), each first measurement of a frozen contract's holdout |
| `tests/independent/adapters.json` | Optional: the [adapters](#independent-implementations) that run an independent implementation on the cases, with their scripts beside it |

- Use a **JSON fixture** to test a calculation: its inputs are values, already
  read. Use a **file-based case** to test the import of the files the
  [data contract](DATA_CONTRACT.md) defines, where the file itself matters:
  delimiters, headers, encodings, decimal separators, line endings, a
  malformed line, or a reconciliation against control totals.
- Case IDs are lowercase kebab-case, equal to the file or directory names,
  and never reused. A case has one form, never both.
- A case whose expected values are all `illustrative` starts with
  `illustrative-`, and only such a case may.
- Every fixture has an expected file, and every expected file a fixture.
- JSON files are UTF-8, stored with LF. All data is **synthetic**: never real
  positions, counterparties, clients, amounts or market data.

`illustrative-format-example` and `illustrative-file-import-example` show the
two forms and hold no regulatory value. They use the illustrative entries of
the methodology guides; remove them when you replace those.

<a id="fixture-format"></a>

## Fixture format

```json
{
  "schema_version": 1,
  "kind": "test-fixture",
  "id": "illustrative-format-example",
  "synthetic": true,
  "description": "What the case exercises, in one sentence",
  "category": "nominal",
  "calculation": "CALC-01",
  "reporting_date": "2026-01-02",
  "reporting_currency_code": "EUR",
  "datasets": {
    "DS-01": [
      {"id": "P1", "inputs": {"currency_code": "EUR", "notional_amount": 600000}},
      {"id": "P2", "inputs": {"currency_code": "EUR", "notional_amount": 400000}}
    ]
  }
}
```

| Rule | Detail |
| --- | --- |
| Envelope | `schema_version` (1), `kind` (`test-fixture`), `id`, `synthetic` (`true`), `description`, `category`, `calculation`, `reporting_date`, `reporting_currency_code` and `datasets` are required, and no other key is allowed. An `invalid` case relaxes this for the calculation inputs ([invalid inputs](#invalid-inputs)) |
| Calculation | A `CALC-` ID declared by a heading of the [calculation contracts](CALCULATION_CONTRACTS.md) |
| Datasets | A non-empty object whose keys are `DS-` IDs declared in the *Datasets* table of the [data contract](DATA_CONTRACT.md); each holds a non-empty list of records |
| Records | Each record has an `id`, unique in the fixture, and an `inputs` object. The `id` is the value of the dataset's identifier field, its required `text` field whose name ends in `_id`, such as the illustrative `position_id` ([fields](DATA_CONTRACT.md#fields)), so `inputs` leaves that field out; it holds every other field the case needs |
| Inputs | Keys are lower snake case and end in a unit suffix (below). The field tables of the data contract decide which fields exist; a key is the field's name with the suffix of its unit, unless the name already ends in one, so the illustrative `notional` and `currency` are `notional_amount` and `currency_code`. Input files of a [file-based case](#file-based-cases) keep the field names of the data contract |
| Currencies | `reporting_currency_code` is the currency of results. A record whose `inputs` hold an `_amount` key also holds `currency_code`, the currency of all its amounts, so a currency mismatch is always visible |
| Illustrative entries | A `CALC-` or `DS-` ID marked *Illustrative* in its guide may be used only by an `illustrative-` case |

**Unit suffixes.** A key without one of these suffixes is rejected:

| Suffix | Value |
| --- | --- |
| `_amount` | Number in the record's `currency_code`, unscaled |
| `_rate`, `_weight`, `_factor` | Decimal number: `0.05`, never `5` meaning 5 % |
| `_years` | Number of years |
| `_count` | Whole number, at least 0 |
| `_date` | Text, ISO `YYYY-MM-DD` |
| `_flag` | `true` or `false` |
| `_code` | Text code; one ending `currency_code` is three upper-case letters (ISO 4217) |
| `_id` | Text identifier |

`null` means not applicable. Files are strict JSON, with no duplicate key,
and numbers are finite in every category: `NaN`, `Infinity` and a number that
overflows to infinity, such as `1e400`, are rejected. Test a non-finite input
in a VBA test module instead.

<a id="file-based-cases"></a>

## File-based cases

```text
tests/fixtures/illustrative-file-import-example/
├── case.json
├── crlf-line-endings/
│   └── positions.csv
└── lf-line-endings/
    └── positions.csv
```

- **Manifest.** `case.json` holds the keys of a [JSON fixture](#fixture-format)
  with `kind` `test-file-fixture`, and with `datasets` as a list of the `DS-`
  IDs of the file datasets the case reads instead of records. Its `id` is the
  directory name. Every key is checked in every category, `invalid` included.
- **Variants.** Each other entry is a directory, one import of the case, named
  in lowercase kebab-case. It holds only files that the *Datasets* table of
  the data contract names for the manifest's datasets. A variant may omit a
  file, for example to test a missing input.
- **Encoding.** Each file decodes in the encoding the *Datasets* table gives
  its dataset, UTF-8 by default. An `invalid` case may hold a file that does
  not, because that is what it tests.
- **Content.** Outside an `invalid` case, each delimited text file passes the
  [reference validator](DATA_CONTRACT.md#reference-validator) of the
  machine-readable data contract: its header names the dataset's fields, and
  each value has its field's type and lies within its allowed values or range.
- **Bytes.** Git stores variant files byte for byte and reports no
  white-space error in them, and editors leave them as they are
  (`.gitattributes`, `.editorconfig`). A file keeps its line endings, final
  newline and trailing blanks, so a case can test them.
- **Expected file.** It has `variants` instead of `outputs`: an object with one
  key per variant directory, each holding a non-empty list of
  [outputs](#expected-format). A variant without expected outputs, or expected
  outputs without a variant, fails. A record `scope` names a dataset of the
  manifest and a text locating the record, such as a line number or a key,
  because the records are rows of the files.

```json
{
  "schema_version": 1,
  "kind": "test-expected",
  "fixture": "illustrative-file-import-example",
  "variants": {
    "crlf-line-endings": [{"quantity": "illustrative_total_notional_amount", "...": "..."}],
    "lf-line-endings": [{"quantity": "illustrative_total_notional_amount", "...": "..."}]
  }
}
```

The other rules are the same for both forms. The category applies to every
variant, so an `invalid` case needs a `_status` text output in each variant.

<a id="expected-format"></a>

## Expected-result format

```json
{
  "schema_version": 1,
  "kind": "test-expected",
  "fixture": "illustrative-format-example",
  "outputs": [
    {
      "quantity": "illustrative_total_notional_amount",
      "scope": null,
      "rule": "CALC-01.1",
      "value": 1000000,
      "unit": "EUR",
      "tolerance": {"absolute": 0.000001, "relative": 0.000000001},
      "reference": {
        "class": "illustrative",
        "source": null,
        "locator": null,
        "derivation": "Why the value exists and how it was obtained",
        "script": null,
        "derived_by": "Name",
        "reviewed_by": null,
        "date": "2026-11-02"
      }
    }
  ]
}
```

| Rule | Detail |
| --- | --- |
| Envelope | `schema_version` (1), `kind` (`test-expected`), `fixture` (the case ID) and a non-empty `outputs` list are required |
| Scope | `null` for a result of the whole calculation, or `{"dataset": "DS-NN", "record": "<id>"}` for a result of one record of the fixture |
| Quantity | Lower snake case, the *Quantity* of one of the calculation contract's *Outputs*, or a `_status` quantity that its *Failure behaviour* lists. Only an `illustrative-` case may use a name starting with `illustrative_` |
| Number | `value`, `unit` and `tolerance`. `unit` is a currency code for an amount, `1` for a rate, weight or factor, or `years` |
| Text | `text`, compared exactly, with no `unit` or `tolerance`. A rejection is a text output whose quantity ends in `_status`, with one of the texts the contract's *Failure behaviour* lists for it |
| Rule | A step of the case's calculation, `CALC-NN.<step>`, such as `CALC-01.2` for step 2 of its *Rules* table. It may be `null` only for an `illustrative` value |
| Reference | One per output; see [reference classes](#reference-classes) |

<a id="reference-classes"></a>

## Reference classes

Every expected value carries exactly one class. Only the first two are
numerical-validation references.

| Class | Meaning | Required fields |
| --- | --- | --- |
| `published` | Printed in a registered source, for example a worked example | `source` (a `SRC-` register ID) and an exact `locator` (article, paragraph, table, printed page) |
| `independent` | Computed by an external implementation by another author, registered in the source register with an `IMPL-` ID | `source` (`IMPL-…`), `locator`, `derivation`, `derived_by` (the implementation and its version) |
| `illustrative` | Shows the format or a behaviour; **not** a validated figure | `derivation`, saying why it exists and how the value was obtained |

- A `published` value never cites an `IMPL-` source, and an `independent`
  value always does. A source marked *Illustrative* in the register supports
  only an `illustrative` value.
- A value derived within the project, even from a registered text, stays
  `illustrative`, including one computed by a script under
  `tests/derivations/`.
- `script`, when not `null`, names a file under `tests/derivations/` that
  computed the value; it must exist.
- `reviewed_by` may be `null`, but record who reviewed a value when someone
  other than its author did.
- **Examples are never benchmarks** unless their source is registered and the
  value is printed there. Test output reports `illustrative` results
  separately and never as validation.

<a id="derivation-scripts"></a>

## Derivation scripts

A script under `tests/derivations/` computes expected values from the
documented method, written from the contracts and never from the code under
test. An output names it in `reference.script`, and
`python tools/check_derivations.py`, which **Repository integrity** runs,
compares the two:

- it runs each named script once per case, as
  `python -I tests/derivations/<name>.py tests/fixtures/<case-id>.json`, or for
  a file-based case once per variant, with
  `tests/fixtures/<case-id>/<variant>` as the argument, comparing that
  variant's outputs;
- the script prints one JSON object, `{"outputs": [...]}`, whose items have the
  `quantity` and `scope` of an expected output and its `value` or `text`;
- every output that names the script must appear in what it prints, and match
  within the output's [tolerance](#tolerance), or exactly for a text.

A failing script, output that is not JSON, a missing quantity or a mismatch
fails the check. A match shows only that the script and the expected file
agree: the value stays `illustrative`, because the project derived it. The
scripts use the Python standard library only and are never a runtime
dependency of the workbook. `tests/derivations/illustrative_total.py` and
`illustrative_file_total.py` show the format for the two example cases.

<a id="independent-implementations"></a>

## Independent implementations

An `independent` value cites an `IMPL-` source: an implementation by another
author, such as an R or Python package. An adapter lets the repository run
that implementation on the cases and compare its results with the expected
values. `python tools/check_independent.py`, which **Repository integrity**
runs, does so for each adapter registered in `tests/independent/adapters.json`:

```json
{
  "schema_version": 1,
  "kind": "independent-adapters",
  "adapters": [
    {"source": "IMPL-01", "command": ["Rscript", "tests/independent/impl_01.R"]}
  ]
}
```

| Field | Rule |
| --- | --- |
| `source` | An `IMPL-` ID of the [source register](SOURCES.md); a source has one adapter at most |
| `command` | The words of the command that runs the adapter, from the repository root. A first word `python` is the interpreter that runs the check; any other is looked up on `PATH` |

An adapter is a script under `tests/independent/`, written for one
implementation. The check runs it once per case, with the argument a
[derivation script](#derivation-scripts) receives: the fixture, or for a
file-based case each variant's directory in turn. The adapter translates that
input into the implementation's own, runs the implementation, and prints one
JSON object:

| Object | When |
| --- | --- |
| `{"outputs": [...]}` | The implementation computed the case. The items have the `quantity` and `scope` of an expected output and its `value` or `text`, as a derivation script prints them |
| `{"unsupported": "<reason>"}` | The implementation cannot represent the case, for example an input it has no field for |
| `{"unavailable": "<reason>"}` | The implementation is not installed here |

The check then:

- compares each printed output that the expected file holds, whatever its
  reference class, within the output's [tolerance](#tolerance), or exactly
  for a text, and ignores a printed quantity that the file does not hold;
- fails on a mismatch, on an output citing the adapter's source that is not
  printed, and on an adapter that fails or prints anything else;
- lists each case or variant the implementation cannot represent, with the
  reason and the number of values citing the source that stay unchecked, and
  fails nothing for it;
- skips the adapter with a note when the first word of its command is not on
  `PATH` or it answers `unavailable`, so the check runs where the
  implementation is not installed. A skipped adapter compares nothing:
  install the implementation in the **Repository integrity** workflow for the
  comparison to run on every pull request.

With no registry, or no adapter in it, as before the project registers one,
the check passes and says that nothing was cross-checked. An adapter and its
implementation are test tools, never a runtime dependency of the workbook. A
match shows that the implementation and the expected file agree; what class a
value has is still decided by the [reference classes](#reference-classes).

<a id="tolerance"></a>

## Tolerance

An actual result `a` matches the expected value `e` when

```text
|a − e| ≤ max(absolute, relative × |e|)
```

Both bounds are stated on every numeric output, as finite numbers of at least
0; there is no hidden default. Both 0 means an exact comparison. Record the
project's default bounds in the *Common rules* of the
[calculation contracts](CALCULATION_CONTRACTS.md#common-rules) once they are
decided, with the decision in that guide's *Decisions* table. A common
starting point:

| Reference | `absolute` | `relative` |
| --- | --- | --- |
| `independent` or `illustrative`, full precision | `1e-6` currency units, `1e-9` for a rate, weight or factor | `1e-9` |
| `published`, rounded | Half a unit of the last printed digit: printed in thousands → `500` | `0` |

- **Near zero** the absolute bound governs: an expected `0` is matched within
  `absolute` only, so an `absolute` of 0 demands an exact zero.
- **Domain rules apply before tolerance:** a quantity that must not be
  negative fails if it is negative, however small. Declare such a rule as a
  `bounds` [invariant](#invariants), which checks it on the results.
- Results are compared unrounded, unless the contract rounds them.

<a id="accuracy-contracts"></a>

## Accuracy contracts

A case's tolerance decides whether that case passes. An accuracy contract
states the guarantee a calculation gives for one regime: the worst error
allowed over the outputs of its quantities in a set of cases. It is what a
release may publish as the calculation's accuracy. The contract's
`threshold` is that guarantee, and blocks. An optional `target`, no larger,
is the worst error the project aims for: it is measured and reported, never
blocks, and so can be tighter than what a release may promise.
`tests/accuracy/contracts.json` holds the contracts:

```json
{
  "schema_version": 1,
  "kind": "accuracy-contracts",
  "min_holdout_cases": 3,
  "contracts": [
    {
      "id": "ACC-01",
      "calculation": "CALC-01",
      "regime": "illustrative",
      "quantities": ["illustrative_total_notional_amount"],
      "metric": "absolute",
      "threshold": 0.000001,
      "target": 0,
      "provenance": "Why the threshold and the target are what they are",
      "status": "draft",
      "decision": null,
      "calibration_cases": ["illustrative-format-example"],
      "holdout_cases": ["illustrative-file-import-example"]
    }
  ]
}
```

| Field | Rule |
| --- | --- |
| `id` | `ACC-` and a number, unique |
| `calculation` | A `CALC-` ID declared by the calculation contracts; every case of the contract tests it |
| `regime` | A free text label, such as the regulation or its version |
| `quantities` | The outputs the contract covers; every case has an expected number for each |
| `metric`, `threshold` | `absolute`: the absolute difference between the result and the expected value; `relative`: that difference divided by the expected value's magnitude, which is 0 at an expected 0 only when the result is 0 too. The threshold, the guarantee, is a finite number of at least 0 |
| `target` | Optional: the worst error aimed for, in the same metric, a finite number of at least 0 and no larger than the threshold. Only reported; leave it out for no target. No rule derives the threshold from it: declare each one |
| `provenance` | Why the threshold, and the target if any, are what they are |
| `status`, `decision` | `draft`, with `decision` `null`, or `frozen`, with the issue or record that froze it |
| `calibration_cases` | The cases that may be used to set the threshold, each listed once |
| `holdout_cases` | The cases that never are, each listed once; no case is in both lists |
| `design` | Optional: the committed design or preregistration document of the contract, a path under `docs/methodology/` or `tests/accuracy/`, where `tools/check_excel_run.py` copies it from; later amendments only append to it. Leave it out when the project keeps none |

- **Frozen.** A contract may be frozen only with at least `min_holdout_cases`
  holdout cases, and only when every expected value of its quantities in
  them is `published` or `independent`. A contract on an illustrative
  calculation stays `draft`, as the illustrative `ACC-01` does.
- **Frozen is kept.** Once a release tag holds a frozen contract, it is never
  removed or returned to `draft`, even by deleting the registry. With the
  same `decision`, it may only get stricter: a lower threshold, or added
  quantities and cases. A change of calculation, regime or metric, a larger
  threshold, or a dropped quantity or case needs a new `decision`. The check
  compares with the newest `vX.Y.Z` tag of the branch. The target is not
  part of the guarantee: it may be added, changed or removed under the same
  `decision`, within the threshold.
- **Measured on the workbook.** The contracts measure the workbook's results,
  which exist only after an Excel run. `python tools/check_accuracy.py`, which
  **Repository integrity** runs, checks the registry against the cases. With
  `--results <file>`, the results file of an Excel run, it also fails on the
  items below. Repeat `--results` to measure the results files of several
  suites together:
  - results of another revision: their `commit` must be the full SHA of the
    candidate, given with `--commit`, or HEAD by default;
  - a result that two results files both hold;
  - an output without a result;
  - a result outside its own tolerance;
  - a worst error above the threshold, which it computes separately for the
    calibration and the holdout cases; the holdout only once its
    [exposure is recorded](#holdout-exposure-record);
  - a frozen contract whose holdout has no exposure recorded under its
    decision;
  - an output that an [invariant](#invariants) reads without a result, and a
    result that breaks an invariant.

  For each split it prints the worst error, the threshold and the target, or
  `no target`, and for the holdout the number of its exposures. A worst
  error above the target but within the threshold passes, and its line ends
  with `target missed`; one above the threshold ends with
  `threshold exceeded` and fails:

  ```text
  ACC-02 calibration: worst relative error 4e-06, threshold 0.0001, target 1e-05
  ACC-02 holdout: worst relative error 5e-05, threshold 0.0001, target 1e-05, exposures 1: target missed
  ```

```json
{
  "schema_version": 1,
  "kind": "test-results",
  "commit": "<full commit SHA of the workbook's source>",
  "results": [
    {"case": "illustrative-format-example", "variant": null,
     "quantity": "illustrative_total_notional_amount", "scope": null, "value": 1000000}
  ]
}
```

A result names its case, its variant (`null` for a JSON fixture), its
quantity and its scope as in the expected file. In an Excel run, a suite
records each result with `TEST_Harness.RecordResult`, and its certification
saves them with `TEST_Harness.SaveResults` in the evidence folder, with the
commit of the workbook's build stamp
([harness guide](TEST_HARNESS.md#results)): the example suite as
`results.json`, a project suite in the file its entry in the
[evidence policy](TEST_HARNESS.md#evidence-policy) names.
`tools/check_excel_run.py` then measures the commit's contracts on all these
files together, and fails the run when they do not hold. A contract whose
outputs a project suite computes is therefore measured on that suite's file.

<a id="holdout-exposure-record"></a>

### Holdout exposure record

A holdout shows how the calculation does on cases that played no part in
setting the threshold only if nobody saw its errors while the threshold could
still change. The record `tests/accuracy/holdout_exposures.json` is the
evidence: it lists each time a frozen contract's holdout was first measured
under a decision. A repository with a registry keeps the record, with an
empty `exposures` list until a holdout is first measured; the check fails
when it is missing, and nothing ever resets it.

- **Sealed while draft.** With `--results`, a draft contract's calibration
  cases are measured; its holdout is not compared, not printed and needs no
  result. Its line says it is sealed. The harness still asserts each case
  within its own tolerance, and the [invariants](#invariants) still read
  every case they name: sealing hides the contract's worst error over the
  holdout, the number a threshold could be tuned to.
- **Sealed until recorded.** A frozen contract's holdout is sealed in the same
  way until its exposure under its current `decision` is recorded, and with
  `--results` that is a finding:

  ```text
  ACC-01 holdout: sealed while the contract is draft, exposures 0
  ACC-02 holdout: sealed until its exposure under decision #N is recorded, exposures 0
  FAIL: tests/accuracy/contracts.json: ACC-02: its holdout has no exposure recorded under decision #N, so it was not measured: measure it once with tools/check_accuracy.py --record-holdout ACC-02
  ```

- **The first exposure.** Freeze the contract, run Excel on a commit that
  holds it frozen, then record the exposure from a checkout of that commit,
  with the results files of the run:

  ```sh
  python tools/check_accuracy.py --results <folder>/results.json --commit <commit> --record-holdout ACC-02
  ```

  It writes the exposure into the record first, then measures as `--results`
  does, holdout included, and prints a `Recorded:` line. Commit the record:
  it is part of what [validates](VALIDATION_PLAN.md#what-an-accuracy-contract-proves)
  the contract. The exposure is recorded even when the holdout then fails:
  its errors were seen. The record is not a run input of `tools/check_excel_run.py`
  ([a run for another commit](../repo_governance/EXCEL_RUN.md#a-run-for-another-commit)):
  the exposure run certifies the commit that records it through
  `--equivalent-to`, reading that commit's record, when nothing else
  changed.
- **Refusals.** `--record-holdout` writes and measures nothing, and ends
  `accuracy: REFUSED`, for a draft contract, a contract already exposed under
  its decision (a second first exposure), a missing, malformed or
  inconsistent record, any finding of the check without results, and
  results that are unusable or lack an output of the holdout:

  ```text
  REFUSED: second first exposure of ACC-02 refused: its holdout was exposed under decision #N with the results of commit 3f2c1e0a9b8d (exposures 1). Another exposure needs a new decision that states the earlier exposures
  accuracy: REFUSED (nothing measured; tests/accuracy/holdout_exposures.json not written)
  ```

- **Regression, not exposure.** Every later measurement of a recorded
  holdout, such as `tools/check_excel_run.py` on a later candidate, measures
  it as a regression of a threshold already fixed, prints the number of
  exposures and writes nothing.

Each exposure holds:

| Field | Holds |
| --- | --- |
| `contract`, `decision` | The contract's ID and the decision it was frozen under |
| `exposures` | The number of exposures of the contract's holdout, this one included: earlier exposures under other decisions count |
| `commit`, `results` | The commit the results are of, and each results file's name and SHA-256 |
| `entry_sha256`, `threshold` | The SHA-256 of the contract's entry, written as JSON with sorted keys, no white space and UTF-8, and its threshold |
| `holdout` | Each holdout case with the SHA-256 of its fixture and of its expected file. A file-based fixture's is the SHA-256 of a listing of its folder: one line per file, sorted by path, with the file's SHA-256, two spaces and its path |
| `design` | `null`, or the contract's `design` document: its path, its length in bytes and their SHA-256 |

The record is checked with or without results. The check fails on an
exposure that is malformed, numbered out of order, or a second one of a
contract under the same decision; on an exposure whose contract is no longer
in the registry; and on a record that no longer starts with the exposures
recorded at the newest `vX.Y.Z` tag. It fails on a frozen contract whose
`decision` is that of an earlier exposure but not of its last: a decision
used again after a later one cannot state the exposures since, and
`--record-holdout` refuses it. For a frozen contract whose `decision` is
still that of its last exposure, it also fails on:

- an entry that differs from the one recorded, the target included;
- a threshold lower than the one recorded: a threshold lowered after the
  holdout errors were seen was set with them;
- a holdout fixture or expected file that differs from the one recorded;
- a design document that no longer begins with the bytes recorded: an
  amendment appends, never rewrites.

```text
FAIL: tests/accuracy/contracts.json: ACC-02: its entry changed since its holdout was exposed under decision #N (exposure 1, results of commit 3f2c1e0a9b8d): a change after an exposure needs a new decision that states the earlier exposures
FAIL: tests/accuracy/contracts.json: ACC-02: threshold 5e-05 is lower than 0.0001, the threshold when its holdout was exposed: a threshold lowered after an exposure needs a new decision
FAIL: tests/accuracy/contracts.json: ACC-02: the expected file of holdout case portfolio-large changed since its holdout was exposed under decision #N (exposure 1, results of commit 3f2c1e0a9b8d)
```

To change a frozen contract after an exposure, name a new `decision` and
state in it the earlier exposures that the record lists. The comparisons
then stop, and the holdout is sealed again until its exposure under the new
decision is recorded, which counts the earlier ones. A contract returned to
draft keeps its exposures; it is sealed and prints them.

A contract measured before the record existed is recorded once, with its
next run; its decision should state that the holdout was measured before.
While the only contract is the illustrative `ACC-01`, the record holds no
exposure: an illustrative contract stays draft and its holdout stays sealed.
The record gets its first exposure when the project freezes its first
contract.

<a id="invariants"></a>

### Invariants

An invariant is a property that the results of a calculation's cases always
have, such as a quantity that is never negative, or parts that add up to
their total. Unlike a contract it has no threshold and no status: a result
that breaks it fails, whatever the status of any contract over the same
cases, and whether or not a contract covers them. Declare invariants in an
optional `invariants` list of the registry, beside `contracts`; a registry
without it keeps `schema_version` 1. The example shows one invariant of each
kind; `INV-02` and `INV-03` are on a calculation `CALC-02` and cases the
template does not have:

```json
"invariants": [
  {
    "id": "INV-01",
    "kind": "bounds",
    "calculation": "CALC-01",
    "basis": "The domain rule that makes it hold",
    "quantity": "illustrative_total_notional_amount",
    "cases": ["illustrative-format-example", "illustrative-file-import-example"],
    "minimum": 0,
    "maximum": null
  },
  {
    "id": "INV-02",
    "kind": "monotone",
    "calculation": "CALC-02",
    "basis": "Why the quantity cannot fall along the sequence",
    "quantity": "total_exposure_amount",
    "scope": null,
    "direction": "non-decreasing",
    "sequence": [
      {"case": "portfolio-small", "variant": null},
      {"case": "portfolio-large", "variant": null}
    ]
  },
  {
    "id": "INV-03",
    "kind": "sum",
    "calculation": "CALC-02",
    "basis": "Why the parts add up to the total",
    "parts": ["secured_exposure_amount", "unsecured_exposure_amount"],
    "total": "total_exposure_amount",
    "cases": ["portfolio-small", "portfolio-large"]
  }
]
```

| Field | Rule |
| --- | --- |
| `id` | `INV-` and a number, unique among the invariants |
| `kind` | `bounds`, `monotone` or `sum`; it decides the other fields, and no other field is allowed |
| `calculation` | A `CALC-` ID declared by the calculation contracts; every case of the invariant tests it |
| `basis` | The domain rule, or the property of the calculation, that makes the invariant hold |
| `quantity` | `bounds` and `monotone`: the quantity checked |
| `cases` | `bounds` and `sum`: the cases checked, each listed once |
| `minimum`, `maximum` | `bounds`: the limits, inclusive, each a finite number or `null` for none; at least one is a number, and the minimum is no larger than the maximum |
| `scope`, `direction`, `sequence` | `monotone`: the scope of the quantity, `null` or a `dataset` and a `record` as in the expected file; `non-decreasing` or `non-increasing`; at least two points in order, each a `case` and its `variant`, `null` for a JSON fixture, none listed twice |
| `parts`, `total` | `sum`: the quantities that add up, each named once; the quantity they add up to, which is not one of the parts, or a constant number |

- **`bounds`:** each number of the quantity, in every variant and scope of
  each case, lies within the limits exactly, with no tolerance: this is how
  a domain rule applies before tolerance.
- **`monotone`:** from each point to the next, the quantity does not move
  against the direction by more than the two outputs' own tolerances added
  together. A smaller reversal is counted in the measurement line, and fails
  nothing.
- **`sum`:** in each case, at every variant and scope where one of the
  quantities has an expected number, every one of them has one, and the
  parts differ from the total by no more than the tolerances of all these
  outputs added together; a constant total has no tolerance of its own.

Without `--results`, the check reads each invariant against the cases: a
calculation that is not declared, a case that does not exist or tests another
calculation, a case or point with no expected number for a quantity the
invariant reads, and expected values that themselves break the invariant are
findings. With `--results`, it also checks the invariant on the results, and
an output it reads without a result is a finding. It prints a line for each
invariant, after the lines of the contracts:

```text
INV-01 bounds: 0.0 <= illustrative_total_notional_amount, 3 result(s) checked
INV-02 monotone: total_exposure_amount non-decreasing over 2 point(s), 1 step(s) checked, 1 reversal(s) within the allowance
INV-03 sum: secured_exposure_amount + unsecured_exposure_amount = total_exposure_amount, 2 group(s) checked, 1 beyond the allowance
```

On results that break them, each break is a `FAIL` line, and the
invariant's line ends with the number of breaks: `outside the bounds`,
`reversal(s) beyond the allowance` or `beyond the allowance`:

```text
FAIL: tests/accuracy/contracts.json: INV-01: illustrative-format-example illustrative_total_notional_amount is -0.5, below the minimum 0.0
FAIL: tests/accuracy/contracts.json: INV-02: total_exposure_amount decreases from 120.0 at portfolio-small to 100.0 at portfolio-large, by 20.0, beyond the allowance 0.02
FAIL: tests/accuracy/contracts.json: INV-03: at portfolio-large, secured_exposure_amount + unsecured_exposure_amount is 100.5 and the total is 100.0: the difference 0.5 is beyond the allowance 0.03
```

An invariant does not repeat the per-case check: every output of a contract's
cases already fails outside its own tolerance. When the expected values keep
a `monotone` or `sum` invariant exactly, results within their own tolerances
keep it too, so such an invariant adds most for the outputs that no contract
measures. A `bounds` invariant is stricter than any tolerance. Until the
project adds its own, the registry declares only the illustrative `INV-01`:
the illustrative total is a sum of notional amounts, which the illustrative
data contract requires to be at least 0. Like the contract on the same cases, it validates nothing.

<a id="categories"></a>

## Case categories

Each fixture declares one `category`:

| Category | Covers |
| --- | --- |
| `nominal` | Typical inputs within the domain |
| `boundary` | Values on a domain edge: a zero amount, a single record, a value at a cap, a result that must be exactly zero |
| `invalid` | Inputs that must be rejected: missing or unknown fields, a value outside its range, an unknown code or currency |

<a id="invalid-inputs"></a>

### Invalid inputs

- An `invalid` case is expected to be **rejected**, never calculated: its
  expected file holds at least one `_status` text output, and no numeric
  output for what it rejects.
- An `invalid` case separates the harness metadata, which must always be
  valid, from the calculation inputs, which may break any input rule, because
  that is what the case tests:
  - **Always valid:** `schema_version`, `kind`, `id`, `synthetic`,
    `description`, `category` and `reporting_date`, and the expected file.
  - **May be invalid, missing or unknown:** `calculation`,
    `reporting_currency_code` and `datasets`, with their records, keys, unit
    suffixes and values.
  - The `description` names the input rule the case breaks.
- The `_status` quantities and their texts are fixed by the calculation
  contract, under its *Failure behaviour*. A rejection never substitutes a
  default value.
- Inputs that JSON cannot express, such as non-finite numbers or wrong VBA
  types, are tested in a VBA test module instead of a fixture.

<a id="other-tests"></a>

## Relation, end-to-end and volume tests

Three kinds of VBA test complement the cases. Each is described in the
[harness guide](TEST_HARNESS.md), and none adds a category:

| Test | Uses | Shows |
| --- | --- | --- |
| [Relation test](TEST_HARNESS.md#relation-tests) | No fixture or expected file: inputs written in the test module | A property of the calculation, such as additivity, symmetry, monotonicity or a zero input giving zero, with a corruption check that a deliberately corrupted output fails it |
| [End-to-end case runner](TEST_HARNESS.md#end-to-end-case-runner) | A case's fixture and expected file, written to the input sheets | That the workbook, run through its public entry point, gives the case's expected values within its tolerance |
| [Volume benchmark](TEST_HARNESS.md#volume-benchmarks) | A large synthetic input built in memory, never committed | How long the calculation takes on one host; never a reference value and never certification evidence |

<a id="completeness-policy"></a>

## Completeness policy

- A rule is **validated** only when at least one `published` or
  `independent` case passes for it. `illustrative` cases, relation tests and
  volume benchmarks never count.
- Each implemented rule also needs at least one `boundary` case, and each
  input domain at least one `invalid` case, as a fixture or a VBA test.
- A run that does not execute every expected case and assertion is
  **incomplete, not a pass**.
- A calculation's accuracy is guaranteed only by a **frozen** accuracy
  contract whose results were measured on an Excel run of the release
  candidate.
- A release states which rules are validated, under which version of their
  sources, and lists every rule that is implemented but not validated
  ([release certification](../release/RELEASE_CERTIFICATION.md)).

<a id="checking"></a>

## Checking

`tools/check_test_cases.py` checks the files under `tests/fixtures/` and
`tests/expected/`, and the traceability table that names them. With no case
and no traceability row present it passes and says so.

It checks:

- strict JSON, with no duplicate key and only finite numbers, in every
  category;
- the harness metadata of every fixture;
- for a file-based case:
  - the whole manifest;
  - the variant names, and that each variant holds only files the data
    contract names for the manifest's datasets;
  - outside an `invalid` case, that each file decodes in its encoding, and
    that a delimited text file passes the
    [reference validator](DATA_CONTRACT.md#reference-validator), whose
    findings it reports with their codes;
  - that the expected file's variants and the variant directories agree;
- outside an `invalid` case, the calculation inputs:
  - the calculation and the datasets, against the methodology guides;
  - record identifiers;
  - keys, unit suffixes and value types;
  - currencies;
- each expected file's name and envelope;
- its outputs:
  - their scope against the fixture;
  - their form, unit and tolerance;
  - their rule, against the case's calculation and the steps that its
    *Rules* table lists;
- the reference class, its source against the register, and its script;
- the `illustrative` naming rules, and a `_status` text output in every
  `invalid` case;
- the [traceability table](VALIDATION_PLAN.md#traceability) of the validation
  plan: each row's source, VBA procedures and cases exist, its calculation
  step is one that the *Rules* table of its calculation contract lists, each
  case tests the row's calculation, and the status agrees with what the row
  names.

It does not check:

- whether an input key or a quantity name belongs to the contracts;
- the content of a variant file in an `invalid` case beyond its name, or of a
  file that is not delimited text beyond its name and encoding;
- what the values of a variant file mean: the import under test checks them;
- that an `invalid` case has no numeric output for what it rejects;
- any value: nothing is computed or compared. `tools/check_derivations.py`
  compares the values that a [derivation script](#derivation-scripts) computes,
  `tools/check_independent.py` those that an
  [independent implementation](#independent-implementations) computes, and
  running the cases on the workbook is the job of a VBA test harness in Excel.
