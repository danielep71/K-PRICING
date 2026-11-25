# 🗂️ Repository structure

This guide says where each kind of file lives, which components may call
which, and how the public VBA declarations are recorded. The
[house style](VBA_HOUSE_STYLE.md) covers how the code is written; the
[methodology guides](methodology/README.md) cover what it must calculate.

Paths marked *reserved* are conventions: create each directory with its first
real file, never as an empty placeholder. The Git, editor and ignore rules for
them are already set in `.gitattributes`, `.editorconfig` and `.gitignore`.

## Locations

| Location | Holds | In this repository |
| --- | --- | --- |
| Repository root | The README; the contribution, conduct, security, installation and release guides (`CONTRIBUTING.md`, `CODE_OF_CONDUCT.md`, `SECURITY.md`, `INSTALLATION.md`, `RELEASING.md`); `CHANGELOG.md`, `AGENTS.md` and `LICENSE`; the Ruff and mypy settings in `pyproject.toml`; and the Git and editor rules in `.gitattributes`, `.gitignore` and `.editorconfig` | Present |
| `VERSION` | The newest released version, published or not, one `X.Y.Z` line, created when the first release is prepared ([releasing](../RELEASING.md#current-status)) | Present once the first release is prepared; [current status](../RELEASING.md#current-status) says whether it exists and, if so, the version it holds |
| `.github/` | Workflows (`workflows/`) and the label scripts they run (`scripts/`); issue forms (`ISSUE_TEMPLATE/`) and the pull request template; the label catalogue `labels.json`; the Dependabot configuration and the hash-locked configuration parser of the workflow tests (`requirements.txt`); the binary allowlist `binary-allowlist.json`; the Excel evidence policy `excel-evidence-policy.json`; and, when needed, the expiring link exceptions `link-exceptions.json` and the reviewed link-check exclusions `link-check-exclusions.json` | Present, without `link-exceptions.json` or `link-check-exclusions.json` |
| `docs/` | Governance, setup, structure, style and the API inventory; the templates of the release certification and milestone closeout records and the pre-release maintenance audit (`docs/release/`); and the license texts, the third-party notice and its machine-readable provenance (`docs/licenses/`) | Present |
| `docs/assets/` | The repository's social preview image, `social-preview.png`, shown at the top of the README, and its editable source `social-preview.svg` | Present |
| `docs/native-declarations.json` | The [register of native declarations](#native-declarations): each export a `Declare` imports, with its library, whether its zero result is ambiguous and, optionally, its ABI | Present, with no declaration |
| `docs/methodology/` | The [index](methodology/README.md) of the guides: sources, data contract and its machine-readable form `data-contract.json`, calculation contracts, validation plan, test-case format, [workbook contract](methodology/WORKBOOK_CONTRACT.md) and [VBA test harness](methodology/TEST_HARNESS.md) | Present, as guides to fill in, with the illustrative dataset `DS-01` |
| `docs/methodology/sources/` | Byte-preserved copies of regulatory sources | Reserved |
| `docs/methodology/workbook-protection.json` | The [protection matrix](methodology/WORKBOOK_CONTRACT.md#protection) of `src/workbook/Template.xlsx`, added with the template | Reserved |
| `docs/user-manual/` | Optional manual for workbook recipients, one folder per released version (`docs/user-manual/X.Y.Z/MANUAL.md`), built to the PDF attached to the release ([user manual](../RELEASING.md#user-manual)) | Reserved |
| `src/core/` | Calculations and pure validation, host-independent standard modules | Reserved |
| `src/classes/` | Host-independent classes, if the design needs them | Reserved |
| `src/modules/` | The public facade: the supported entry points | Reserved |
| `src/workbook/` | Excel adapters, `ThisWorkbook` and sheet document modules, and the optional macro-free `Template.xlsx`, laid out as the [workbook contract](methodology/WORKBOOK_CONTRACT.md) defines | Reserved |
| `src/forms/` | UserForms, each `.frm` with its `.frx` | Reserved |
| `tests/` | Python tests of the repository tooling (`test_*.py`) | Present |
| `tests/modules/` | VBA test modules: the [test harness](methodology/TEST_HARNESS.md), its example suite, relation tests, case runners and volume benchmarks | Present, with the harness, the example suite and the illustrative relation tests |
| `tests/fixtures/`, `tests/expected/` | Synthetic inputs and expected results in the [test-case format](methodology/TEST_CASES.md) | Present, with one illustrative case of each form |
| `tests/derivations/` | Scripts that compute expected values from the documented method ([derivation scripts](methodology/TEST_CASES.md#derivation-scripts)) | Present, with one illustrative script for each case |
| `tests/accuracy/` | The [accuracy contracts](methodology/TEST_CASES.md#accuracy-contracts) of the calculations | Present, with the illustrative draft contract `ACC-01` |
| `tests/independent/` | The adapters that run an [independent implementation](methodology/TEST_CASES.md#independent-implementations) on the test cases, and their registry `adapters.json` | Reserved |
| `tools/` | Repository checks, setup and build tooling, in Python; `tools/devbuild/` builds a workbook from source ([building a workbook from source](../INSTALLATION.md#building-a-workbook-from-source)) | Present |
| `examples/modules/` | Synthetic VBA examples that use only the supported API | Reserved |
| `test-results/` | Disposable output of local and CI runs, and the mutants that `tools/make_mutant.py` writes under `test-results/mutants/` | Ignored by Git |

- **Source first.** The exported source is authoritative. A change made in a
  workbook is exported and reviewed like any other change.
- **No built workbooks.** Workbooks are ignored, except
  `src/workbook/Template.xlsx`, a macro-free template, and the copies under
  `docs/methodology/sources/`. A built workbook is saved outside the checkout
  or under `test-results/`.
- **No unreviewed binaries.** **Repository integrity** runs
  `tools/check_binaries.py`, which fails on any tracked file that
  `.gitattributes` treats as binary, or that contains a NUL byte, outside those
  paths, the variant directories of file-based test cases
  (`tests/fixtures/<case-id>/<variant>/`) and the `.frx` beside its `.frm`. To keep another binary, such as an
  image for the documentation, review it and add its path, SHA-256 and reason
  to `.github/binary-allowlist.json`; a later change to the file fails until
  the entry is updated. An Office Open XML package outside
  `docs/methodology/sources/` also fails while it carries macros or any
  finding of `tools/check_workbook_privacy.py`: a OneDrive or SharePoint
  path, a user name, a sensitivity label or SharePoint metadata.
- **Forms.** A form's `.frx` stays beside its `.frm`.

## Dependencies

The Excel adapter calls the public facade. The facade validates its contract
and delegates to the core, which may use host-independent classes. Nothing
calls back the other way, so there are no cycles.

- Core and classes have no Excel object model, worksheet, range, user
  interface, file or network access, and no mutable run state.
- The facade orchestrates and validates; it does not repeat formulas. It
  returns structured results and errors.
- Workbook and form code maps cells and events to facade inputs and displays
  results; it contains no regulatory formula. It addresses sheets by code
  name and ranges by defined name, as the
  [workbook contract](methodology/WORKBOOK_CONTRACT.md) lists them.
- Tests may call the core directly. Examples use only the supported facade.
- The workbook never calls Python. Reference values are derived from the
  methodology, not by translating the VBA
  ([validation plan](methodology/VALIDATION_PLAN.md#reference-values)).
- Repository tooling is not a second implementation of the calculation.

The [static checks](VBA_HOUSE_STYLE.md#static-checks) enforce the locations
and the role prefixes, and match the inventory below against `src/`.
They enforce this dependency direction by name for core, classes and
examples only; the direction of the facade and the adapters is checked in
review.

## Public API inventory

[PUBLIC_API.txt](PUBLIC_API.txt) lists every `Public` declaration under
`src/`, in two tiers; both are empty until the project adds its own.

- **`[supported]`:** the supported API, the public declarations of the
  standard modules in `src/modules/`. A `Public` declaration is supported
  only if it is listed here.
- **`[project-public]`:** every other `Public` declaration under `src/`, in
  core, classes, workbook modules and forms. Other components of the project
  may call them; no caller outside the project may rely on them. A class is
  reached through the facade, so its public members are listed here, never
  under `[supported]`.

The tiers let a reviewer tell a deliberate seam from a member made `Public`
by mistake. `Option Private Module` limits visibility between projects; it is
not a security boundary. Test and example modules are never part of a
production workbook, so their public declarations belong to neither tier.

Each tier starts with its heading on a line of its own. Each declaration is
recorded under its tier as two tab-separated lines:

```text
[supported]
<component>	<kind>	<name>
# SIG	<component>	<kind>	<name>	<normalized declaration>

[project-public]
<component>	<kind>	<name>
# SIG	<component>	<kind>	<name>	<normalized declaration>
```

- **Kinds:** `Function`, `Sub`, `Property Get`, `Property Let`,
  `Property Set`, `Const`, `Enum` or `Type`. Under `[project-public]`, a
  class, form or workbook module can also expose an `Event`, a `Variable` or
  a `WithEvents Variable`.
- **Normalized declaration:** the declaration on one line, outside string
  literals with its white space reduced to single spaces and none just inside
  parentheses, so `Total( ByVal Amount As Double )` is recorded as
  `Total(ByVal Amount As Double)`. An `Enum` or `Type` is recorded with its
  whole block, each line normalized and joined with ` | `, for example
  `Public Enum DemoMode | dmFirst = 1 | End Enum`.
- **Conditional compilation:** a declaration compiled in several `#If`
  branches has one row. Its `# SIG` lines give each distinct signature once,
  so a declaration whose parameters differ between branches has one `# SIG`
  line per branch signature.
- Blank lines and other lines starting with `#` are ignored, so
  `# [supported]` is a comment, not a heading.
- **Matching:** each tier must list exactly the declarations of its
  components. A declaration listed under the other tier is a finding, and so
  are an unknown or repeated heading and a row or `# SIG` line before the
  first heading.
- **Inventory without headings:** an inventory with no tier heading is read
  as `[supported]` alone. It passes only while `src/` holds no `Public`
  declaration outside `src/modules/`; each such declaration is reported as
  not listed. Add both headings and list those declarations under
  `[project-public]`.

A change to a `[project-public]` declaration updates its row in the same pull
request. A change to the supported API needs an issue and an update, in the
same pull request, to:

- the inventory;
- the [calculation contract](methodology/CALCULATION_CONTRACTS.md);
- the tests.

There are no supported entry points until these are reviewed together with
the implementation.

## Native declarations

`docs/native-declarations.json` registers each native export that a
`Declare` imports; it lists none until the project adds a `Declare`.
`python tools/check_vba_declares.py`, which **Repository integrity** runs,
reports a `Declare` whose export it does not list or whose signature differs
from the export's ABI, and holds every call to the error-capture rules of
the [house style](VBA_HOUSE_STYLE.md#naming-and-visibility).

```json
{
  "schema_version": 1,
  "declarations": [
    {
      "export": "TlsGetValue",
      "library": "kernel32",
      "zero_is_ambiguous": true,
      "arguments": [{"type": "Long", "byval": true}],
      "returns": "pointer"
    },
    {"export": "SetLastError", "library": "kernel32", "zero_is_ambiguous": false}
  ]
}
```

- **`export`:** the name the library exports: the `Alias` of the `Declare`
  if it has one, else its name, or an ordinal such as `#12`. Case is
  ignored.
- **`library`:** the `Lib` of the `Declare`, such as `kernel32`. Case and a
  `.dll` suffix are ignored.
- **`zero_is_ambiguous`:** `true` when a zero result can mean success as well
  as failure, so that only the error tells them apart. Each call then
  clears the error with `SetLastError 0` and captures it.
- Each export and library has one entry. `GetLastError` has none, because it
  is never declared, and a `Declare Sub`, which returns nothing, is never
  ambiguous on zero.
- An entry that no `Declare` uses is not reported: remove it with the last
  `Declare` of its export.
- **`arguments`** (optional): the export's parameters in order, each
  `{"type": ..., "byval": true|false}`. `type` is a VBA type, such as `Long`
  or `String`, with `()` for an array, or `pointer` for a pointer-sized value
  such as a handle; never `LongPtr`, which VBA6 does not have. `byval` is
  `false` for a parameter passed `ByRef`, which is also a parameter written
  with neither keyword.
- **`returns`** (optional): the type a `Function` returns, in the same
  form, or `null` for a `Sub`.
- With `arguments` or `returns`, each `Declare` of the export must match them
  in every compilation environment: the number of parameters, each one's type
  and passing, and `Sub` or `Function` with its type. A `pointer` must be
  `LongPtr` where VBA7 is true and `Long` where it is not; every other type
  must be the same, ignoring case. A type suffix, as in `dwIndex&`, gives a
  type, and a name with neither suffix nor `As` is a `Variant`. An entry
  without them is not compared.
- Any other key fails.

Add the entry in the pull request that adds the `Declare`, after checking
the export's documentation for what a zero result means and, when you record
its ABI, for the width of each argument and of the result.

## Input and output contracts

Before the first facade entry point exists, its inputs and outputs are defined
in the [data contract](methodology/DATA_CONTRACT.md) and the
[calculation contracts](methodology/CALCULATION_CONTRACTS.md). For every field
they give:

- the type, unit, currency and date basis;
- the valid range and sign convention;
- how a missing value is treated.

A contract also names:

- the approach and the version of the rules it implements;
- where the inputs come from, and the reporting and calculation dates;
- the validation failures and the result status;
- the output units and the intermediate values kept for audit.

It distinguishes invalid input, unsupported scope and numerical failure. It
never substitutes a default silently, turns missing values into zero, or
presents a partial result as complete.
