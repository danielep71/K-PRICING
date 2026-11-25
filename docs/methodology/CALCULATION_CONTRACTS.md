# 🧮 Calculation contracts

A contract states what one calculation must produce, from which inputs, under
which rules. The VBA implements the contract, the
[validation plan](VALIDATION_PLAN.md) tests against it, and a change to a
result it defines is a `behavior-change`.

## Common rules

State once, for every calculation:

- **Units and signs:** the currency, unit and sign convention of every amount,
  and whether rates are fractions or percentages.
- **Dates:** the reporting date, the day-count conventions, and how
  maturities and residual terms are measured.
- **Precision and rounding:** VBA `Double` arithmetic throughout, rounding
  only where a source requires it, with the method and the number of digits.
- **Currency conversion:** which rates, from which date and source.
- **Failure:** a calculation that cannot complete reports why, and withdraws
  its earlier outputs instead of leaving them as if current.
- **Tolerances:** the default absolute and relative bounds of the test cases'
  comparisons, once decided ([tolerance](TEST_CASES.md#tolerance)), with the
  decision in the *Decisions* table below.

## Contract template

Copy this section for each calculation.

### *Illustrative:* `CALC-01` Example calculation

**Purpose:** what the figure is, and where it is reported.

**Inputs:**

| Input | Dataset and field | Unit |
| --- | --- | --- |
| Example amount | `DS-01.notional` | Currency of `DS-01.currency` |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | Example: sum the notional amounts | `SRC-01`, Article 1(1) |

- **Step:** a whole number, once per step. A test case's `rule` and a
  traceability row name step 2 as `CALC-NN.2`, and
  `tools/check_test_cases.py` rejects a step that this table does not list.

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Example total | `illustrative_total_notional_amount` | Reporting currency | Whole calculation | None |

- **Quantity:** the lower snake case name that the test cases' expected
  files and the accuracy contracts use for the output
  ([expected-result format](TEST_CASES.md#expected-format)). A `_status`
  quantity, which reports a rejection, is not listed here: it belongs to
  *Failure behaviour*.

**Limits and non-goals:** cases the calculation does not cover, and what
happens if they occur in the input.

**Failure behaviour:** which conditions stop it, and what it reports: each
`_status` quantity, in lower snake case ending in `_status`, with the exact
texts it can hold, which the `invalid` test cases expect
([invalid inputs](TEST_CASES.md#invalid-inputs)).

**Validation:** the checks in the [validation plan](VALIDATION_PLAN.md) that
cover it, and the rows of its steps in its
[traceability table](VALIDATION_PLAN.md#traceability).

## Decisions

| Date | Decision | Reason | Issue |
| --- | --- | --- | --- |
| *Illustrative* | How a rule the source leaves open is applied | Why | Link |
