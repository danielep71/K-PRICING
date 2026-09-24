# K-PRICING developer setup

K-PRICING is an initialized **application** repository, version `0.0.0`.
Initialization is complete and the frozen KPR date-layer candidate has been
migrated into the canonical source layout. Do not initialize the repository again.
Destination Excel compilation and source parity passed on one Windows 64-bit
host under [issue #17](https://github.com/danielep71/K-PRICING/issues/17);
other environments are not certified.

## Architecture and source ownership

| Layer | Current content | Intended responsibility |
| --- | --- | --- |
| `src/core/` | `KPR_Core_Err.bas`, `KPR_Core_Parse.bas`, `KPR_Core_Dates.bas`, `KPR_Core_Array.bas` | Internal error, parsing, calendar and array/shape implementation |
| `src/modules/` | `KPR_DATES_DAYS.bas` | Supported 22-function `KPR_Dates_*` worksheet façade |
| `tests/modules/` | `KPR_REGRESSION_TESTS.bas` | Migrated pure and stateful regression harness |
| `examples/modules/` | `KPR_DateExample.bas` | Minimal direct-VBA consumer using the supported façade |
| Future workbook/add-in/UI | Not implemented | Registration, startup/shutdown, configuration, packaging and recovery, separated from calculations |

Exported text in Git is authoritative. Preserve component names, export
attributes, encoding and line endings. A form and its resource companion are one
component; Ribbon XML and callbacks must be reviewed together when introduced.
Keep generated workbooks/add-ins, credentials, local environments and client data
out of source control. The [repository structure](REPOSITORY_STRUCTURE.md) owns
the full layout and export rules. Preserve existing MIT attribution in
[`LICENSE`](../LICENSE); migration must retain applicable source attribution.

## Clean checkout and tools

Use Git to clone this public repository. A source ZIP omits governance
files because of export rules and is not a complete developer checkout.

```bash
git clone https://github.com/danielep71/K-PRICING.git
cd K-PRICING
git status --short
git rev-parse HEAD
```

The hosted baseline is Ubuntu 24.04 and Python 3.10. The hash-pinned quality-tool
lock is for that CI environment; do not remove hashes to force a Windows install.
On a matching Linux environment:

```bash
python3.10 -m venv .venv
. .venv/bin/activate
python -m pip install --disable-pip-version-check --only-binary=:all: --require-hashes -r tools/requirements-quality-ci.txt
ruff check tools
mypy
```

Portable checks use Python 3.10 or newer and the standard library. On Windows,
use the corresponding installed Python launcher, such as `py -3.10`, in place
of `python3`. Node.js is required only for the label scripts; their exact runtime
is the GitHub runner's Node installation. The workflow installs and verifies
the pinned actionlint binary itself.

```bash
python3 tools/check_repo.py --root . --self-test
python3 tools/check_repo.py --root .
python3 tools/check_kpr_contract.py --root . --self-test
python3 tools/check_kpr_contract.py --root .
python3 tools/test_documentation.py -v
python3 tools/test_verification_depth.py -v
python3 tools/check_documentation.py --root .
python3 tools/check_release.py --root . --self-test
python3 tools/check_release_semantics.py --root . --self-test
python3 tools/check_release_semantics.py --root .
node .github/scripts/labels-sync.mjs --self-test
node .github/scripts/labels-drift.mjs --self-test
```

These are the first local checks, not the entire hosted gate. The exact complete
command inventory, pinned tool versions, installation hashes and required
outcomes are maintained in
[`static-checks.yml`](../.github/workflows/static-checks.yml).
`Repository integrity` must pass on the reviewed candidate before merge. Its
artifacts retain the validator reports and actual quality-tool versions for
30 days. An installation failure or missing report is a failure, not a skip.

## Windows Excel validation

The intended initial host is **Microsoft 365 Excel desktop on Windows**, with
32-bit and 64-bit Office treated as separate targets. The accepted neutral-starter
run in [the completion record](SETUP_COMPLETION.md) is setup history only; it does
not certify the migrated KPR source. Destination compilation and parity passed on
one Windows 64-bit host in issue #17; 32-bit Office is untested. Mac, Excel for the web and older desktop
builds are not claimed as supported. The authoritative import order is in
[`INSTALLATION.md`](../INSTALLATION.md).

The migrated date layer uses built-in VBA and Excel object models. Its stateful
host/shape regressions create controlled scratch workbooks, while dynamic-array
members are reached late-bound by the imported harness. Use a fresh macro-enabled
workbook and the existing organizational macro policy; manual import does not
require enabling programmatic VBA-project access. Follow
[the Excel runbook](EXCEL_SETUP_RUNBOOK.md) for exact-source validation and
evidence capture. Linux CI does not compile or execute VBA.

## Change and review

Create a focused branch and PR, update contracts and tests together, and preserve
the distinction between source validation and host evidence. The maintainer
manually verifies the exact CI revision and resolves review findings. Active GitHub
rulesets enforce the documented branch/tag baseline. See
[the dated setup verification](SETUP_VERIFICATION.md) for the live settings,
security controls and the reviewed template-update process.
