# Date-layer fixtures

`date_layer_fixtures.tsv` is the canonical, independently generated fixture
set for the 22-function date layer (issue #38). `tools/gen_fixtures.py`
computes every expected result from
[`docs/DATE_LAYER_CONTRACT.md`](../../docs/DATE_LAYER_CONTRACT.md) with Python
standard-library date arithmetic. It never imports, runs or translates the
VBA implementation, so the fixtures are an oracle that can disagree with it.

The generated VBA module
[`tests/modules/KPR_Test_Fixtures_Generated.bas`](../modules/KPR_Test_Fixtures_Generated.bas)
is derived from the parsed TSV and replays the same cases as VBA values. Both
files are generated source: never edit them by hand.

## Regenerate and check

```bash
python3 tools/gen_fixtures.py --write       # regenerate the TSV and the VBA module
python3 tools/gen_fixtures.py --check       # fail if either committed file is stale
python3 tools/gen_fixtures.py --self-test   # contract examples, round trips, drift detection
```

Hosted CI runs `--self-test` and `--check` on every change. Generation is
deterministic: case IDs, order and bytes change only when the generator, its
authored cases or the contract changes. A change to expected values therefore
needs a reviewed contract or generator change, never a hand edit.

The generator also cross-checks the contract itself: it reads the public
surface from section 2 and the condition registry from section 7, and fails
if a function or registry condition has no fixture. The one exclusion is
recorded in the generator: `HOST_UNRESOLVED` needs a worksheet host whose date
system cannot be read, which only failure injection can produce.

## Columns

| Column | Meaning |
| --- | --- |
| `id` | Stable semantic case identifier, `<suite>.<slug>` |
| `suite` | Fixture family, for example `pillar-format` or `capacity` |
| `function` | Public function under test |
| `context` | `direct` (direct VBA, 1900 serial contract), `ws1900` or `ws1904` worksheet caller |
| `arg1`–`arg5` | Encoded arguments; `-` means not supplied, so omitted controls are trailing |
| `input_kind` | Encoded kinds of the supplied arguments, for review and filtering |
| `expect_type` | `Long`, `Date`, `Boolean`, `String`, `Error` or `Array` |
| `expected` | Encoded expected value, native error or array |
| `condition` | Registry condition ID; blank on success |
| `level` | `element`, `call` or blank on success |
| `rationale` | Contract section and the rule the case demonstrates |

Condition ID, expected error and level are separate fields, as section 7
requires. For an array result, `condition` lists 1-based positions such as
`r1c2=DATE_TEXT_FORMAT;r2c1=INPUT_BLANK_REQUIRED`. A filled array uses
`all=<condition>` when every element shares one outcome. Host-generated and
propagated `#N/A` keep separate case IDs and conditions even though their Excel
values are identical.

## Value encoding

| Encoding | VBA value |
| --- | --- |
| `date:2024-02-29` | Native `Date` |
| `datetime:2024-02-29T13:30:00` | Native `Date` with a time component |
| `num:45292.5` | `Double`, the worksheet numeric type |
| `lng:5` | `Long` |
| `str:"text"` | `String`, JSON-quoted so outer spaces are visible |
| `bool:TRUE`, `bool:FALSE` | `Boolean` |
| `empty`, `null` | `Empty`, `Null` |
| `err:#N/A` | Native Excel error (`#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`) |
| `arr1:[a;b;c]` | One-dimensional array (a `1xN` row) |
| `arr2:2x3[a;b;c;d;e;f]` | Two-dimensional array, items row-major |
| `fill2:1x100000[a]` | Two-dimensional array with one value in every position |

`fill2` keeps the 100,000- and 100,001-element capacity cases reviewable. The
VBA module builds those arrays at run time.

## Contract interpretation notes

Two contract passages disagree. The fixtures follow the section 7 registry,
whose identifiers fixtures are required to cite, and the difference is
recorded here for a separate contract wording review:

- **Blank integer argument.** The section 3.2 integer matrix lists blank and
  `Empty` under `INTEGER_TYPE_REJECTED`. The registry defines
  `INPUT_BLANK_REQUIRED` as a blank or `Empty` at any required value position
  and limits `INTEGER_TYPE_REJECTED` to Boolean, `Date`, text, `Null` or
  object. Fixtures use `INPUT_BLANK_REQUIRED`; both give `#VALUE!`.
- **Blank Pillar argument.** Section 3.4 has no blank row. Fixtures apply the
  same registry rule and use `INPUT_BLANK_REQUIRED`.

## Scope

The fixtures use in-memory values only. Real Range behavior (multi-area
Ranges, `UsedRange` independence and worksheet caller state) belongs to the
stateful runners in `KPR_REGRESSION_TESTS`. The durable runner (#39) executes
these fixtures: its `fixtures` suite replays the direct cases with strict
result types, and its `worksheet-fixtures` suite replays the 1900 and 1904
worksheet-caller cases as Range formulas, comparing worksheet numbers by value
([KPR_TEST_EVIDENCE.md](../../docs/KPR_TEST_EVIDENCE.md)). The facade returns
Excel values only, so replay asserts the expected value and error code; the
condition ID and level are provenance for review. Calendar, weekend-mask,
holiday and business-day fixtures are outside v0.0.4.
