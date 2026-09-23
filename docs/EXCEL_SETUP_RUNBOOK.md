# Neutral starter: Windows Excel setup run

This reusable procedure produced the accepted run for closed
[setup issue 8](https://github.com/danielep71/K-PRICING/issues/8); see
[the retained completion record](SETUP_COMPLETION.md).
It validates only the neutral starter. Pricing behavior, migrated date-layer
parity, application lifecycle and a distributable package remain outside this run.

## Before opening Excel

1. Use the full 40-character candidate SHA pinned in issue 8 after setup PRs
   merge. Obtain that revision with Git, record `git rev-parse HEAD`, and check
   that `git status --short` is empty. Do not import files from a moving branch
   or mix versions.
2. Create a fresh evidence directory outside the checkout. Record operator,
   UTC date/time, SHA, `VERSION`, Windows version, Excel version/build and Office
   bitness (Excel **File > Account > About Excel**).
3. Record the current macro policy and **Trust access to the VBA project object
   model** setting without changing either. Follow organizational policy. Use
   a clean disposable workbook containing no personal or client data.
4. Set a 15-minute session limit. If import, compilation or execution stalls,
   record where it stopped; a timeout or incomplete run is not a pass.

## Import and compile

Create a blank workbook and save it locally as `K-PRICING-setup.xlsm`. Open the
VBE with **Alt+F11** and select this workbook's project. Use **File > Import File**
in this order:

1. `src/core/ProjectCore.bas`
2. `src/modules/ProjectFacade.bas`
3. `tests/modules/ProjectTests.bas`
4. `examples/modules/ProjectExample.bas`

Confirm the four component names match those filenames. Record the enabled
references from **Tools > References**, including any `MISSING` reference.
No additional reference should be required beyond built-in VBA/Excel references.
Run **Debug > Compile VBAProject** and record the observed outcome. If it fails,
capture the exact error, component and highlighted statement and stop the run.
Do not change the frozen source just to make this evidence pass.

## Execute and retain the complete output

Open the Immediate window with **Ctrl+G**, clear old output, and execute:

```vb
ProjectTests.RunProjectTests
```

Copy the entire output, including its environment header and every case, into
a UTF-8 `regression.log` file. The expected final line is:

```text
RESULT=PASS; completeness=COMPLETE; cases=4; assertions=6; failures=0; cleanup=PASS
```

The four case IDs are `ratio.exact`, `ratio.tolerance`,
`ratio.zero-denominator` and `ratio.repeatability`. Preserve their actual results;
the zero-denominator case exercises the expected error. Do not substitute this
example summary for the observed log.

Then run:

```vb
ProjectExample.RunProjectExample
```

Save its output separately as `example.log`; expect `ProjectRatio(12, 4) = 3`.
Confirm Excel remains usable and the harness reports unchanged calculation,
display-alert, event and screen-updating state. Close only the disposable
workbook after saving the logs; record cleanup and any unexpected effects.
If interrupted, record the failure before considering the documented
`ProjectTests.ResetProjectTests` recovery command. Recovery is not a passing run.

## Evidence handoff

Return the complete two logs and the following observed facts to issue 8 or the
maintainer assisting with the run:

```text
Operator and UTC start/end:
Full candidate SHA and VERSION:
Windows version:
Excel version and build:
Office bitness / VBA generation:
Enabled references / missing references:
Macro policy and existing VBOM access setting:
Import outcome:
Compile outcome and basis:
Regression outcome:
Example outcome:
Cleanup outcome / host usable:
Failures, deviations and untested environments:
```

The maintainer will bind the observations to the candidate's Git-byte component
hashes and log hashes in a manual `host.json`, following
[the Excel evidence schema](EXCEL_EVIDENCE.md). The required import inventory
contains core, facade and tests; the example log is additional smoke evidence.
Retain unsuccessful records too. Validate the assembled record using
`tools/check_excel_evidence.py` and its documented command before linking it as
accepted evidence. Never invent missing observations or infer a second bitness
from the tested host. For future candidates, accept evidence only after this validation. The original
issue 8 run is already accepted; current setup closeout is tracked in
[issue 10](https://github.com/danielep71/K-PRICING/issues/10).
