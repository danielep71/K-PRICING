# 💡 Examples

`examples/` contains reproducible, non-production examples that teach the supported API and can be rebuilt from committed source.

Appropriate contents include:

- example VBA modules;
- source-controlled demo builders;
- synthetic input data;
- minimal integration snippets; and
- instructions for creating an example `.xlsm` from the tagged source.

Examples must use supported behavior, synthetic or redistributable data, and the same installation path documented for users. They must not become an undocumented second implementation or a substitute for regression tests.

## Neutral example

After importing the production modules, optionally import
`modules/ProjectExample.bas` and run `ProjectExample.RunProjectExample`. It calls
the supported `ProjectFacade.ProjectRatio` entry point with explicit scalar
inputs and writes `ProjectRatio(12, 4) = 3` to the Immediate window. It does
not read or mutate workbook, worksheet, selection, calculation, event, or UI
state.

Use `demo/` instead only when an interactive demo is itself a distinct project deliverable, an established public path must remain stable, or packaging automation requires that profile. Document the reason in the root README and do not maintain both `examples/` and `demo/` for the same purpose.

Do not commit opaque generated workbooks here unless the repository's release policy explicitly treats them as reviewed source artifacts. Published binaries normally belong to GitHub Releases.

Delete this README only if real examples and equivalent instructions make the directory's role equally explicit.
