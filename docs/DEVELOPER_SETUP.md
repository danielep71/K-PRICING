# K-PRICING developer setup

K-PRICING is an initialized **application** scaffold, version `0.0.0`.
Initialization is complete. Keep the neutral starter until the coordinated
replacement in [the migration plan](MIGRATION_PLAN.md); do not initialize again.

## Architecture and source ownership

| Layer | Current content | Intended responsibility |
| --- | --- | --- |
| `src/core/` | `ProjectCore.bas` | Calculation implementation; migrated `KPR_Core_*` modules remain internal |
| `src/modules/` | `ProjectFacade.bas` | Documented callable surface; retain the existing `KPR_` names during migration |
| `tests/modules/` | `ProjectTests.bas` | Deterministic regression and error/cleanup evidence |
| `examples/modules/` | `ProjectExample.bas` | Small consumers using only supported APIs |
| Future workbook/add-in/UI | Not implemented | Registration, startup/shutdown, configuration, packaging and recovery, separated from calculations |

Exported text in Git is authoritative. Preserve component names, export
attributes, encoding and line endings. A form and its resource companion are one
component; Ribbon XML and callbacks must be reviewed together when introduced.
Keep generated workbooks/add-ins, credentials, local environments and client data
out of source control. The [repository structure](REPOSITORY_STRUCTURE.md) owns
the full layout and export rules. Preserve existing MIT attribution in
[`LICENSE`](../LICENSE); migration must retain applicable source attribution.

## Clean checkout and tools

Use Git with your existing private-repository access. A source ZIP omits governance
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
python3 tools/test_documentation.py -v
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
32-bit and 64-bit Office as separate intended targets. The neutral starter has passed on the single 64-bit host in
[the completion record](SETUP_COMPLETION.md); 32-bit Office remains untested. Mac, Excel for the web and older desktop
builds are not claimed as supported. The authoritative support state and import
order are in [`INSTALLATION.md`](../INSTALLATION.md).

The starter uses built-in VBA and Excel object models, with no additional
third-party reference, native API, network service or workbook fixture. Use a
fresh macro-enabled workbook and the existing organizational macro policy;
manual import does not require enabling programmatic VBA-project access.
Follow [the setup runbook](EXCEL_SETUP_RUNBOOK.md) for the exact test, expected
summary and evidence capture. Linux CI does not compile or execute this VBA.

## Change and review

Create a focused branch and PR, update contracts and tests together, and preserve
the distinction between source validation and host evidence. The maintainer
manually verifies the exact CI revision and resolves review findings. GitHub
rulesets are currently unavailable under the private plan, so this is a process
control, not enforced branch protection. See
[the dated setup verification](SETUP_VERIFICATION.md) for that limitation,
security controls and the reviewed template-update process.
