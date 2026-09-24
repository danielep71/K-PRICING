# 💡 Examples

`examples/` contains reproducible, non-production examples that teach the supported API and can be rebuilt from committed source.

Appropriate contents include:

- example VBA modules;
- source-controlled demo builders;
- synthetic input data;
- minimal integration snippets; and
- instructions for creating an example `.xlsm` from the tagged source.

Examples must use supported behavior, synthetic or redistributable data, and the same installation path documented for users. They must not become an undocumented second implementation or a substitute for regression tests.

## Minimal migrated example

After importing the production modules, optionally import
`modules/KPR_DateExample.bas` and run `KPR_DateExample.RunDateExample`.
It calls the supported `KPR_Dates_DaysInMonth` entry point from direct VBA with
an ISO date and writes the result to the Immediate window. Direct VBA follows
the frozen no-worksheet-host 1900 serial contract; worksheet callers remain
subject to the caller workbook's date-system rules.

The example is intentionally minimal. The planned full demo remains outside
v0.0.2 and is carried forward through the migration backlog.

Use `demo/` instead only when an interactive demo is itself a distinct project deliverable, an established public path must remain stable, or packaging automation requires that profile. Document the reason in the root README and do not maintain both `examples/` and `demo/` for the same purpose.

Do not commit opaque generated workbooks here unless the repository's release policy explicitly treats them as reviewed source artifacts. Published binaries normally belong to GitHub Releases.

Delete this README only if real examples and equivalent instructions make the directory's role equally explicit.
