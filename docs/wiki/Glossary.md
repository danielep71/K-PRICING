# 📖 Glossary

> **Guide, not policy:** [docs/README.md](../README.md) is authoritative.

| Term | Meaning |
| --- | --- |
| Application | Workbook/add-in solution owning deployment and lifecycle |
| Library | Reusable callable VBA without an owned end-user shell |
| UI component | Embeddable component owning a bounded interactive surface |
| Facade | Supported caller-facing boundary; delegates internal computation |
| Core | Internal implementation that does not depend back on examples or tests |
| Project-private | Accessible as permitted inside the VBA project, excluded from the supported external surface |
| VBE export | Reviewable component source exported by the VBA editor, including required attributes and companions |
| Policy | Versioned statement of required behavior or controls |
| Gate | Check that reports findings and fails when its contract is not met |
| Fixture | Deliberately constructed input used to test a checker or workflow |
| Pilot | Recorded attempt to follow a real operational journey, with its scope and gaps stated |
| Evidence | Retained observations supporting a claim, including source and environment |
| Exact SHA | Full Git commit identifier for the particular source under review |
| Merge candidate | GitHub's synthesized PR test commit; may differ from the branch head |
| Ruleset | GitHub server-side restrictions applied to specified branches or tags |
| Annotated tag | Tag object with metadata pointing to the certified release commit |
| Provenance | Recorded relationship between source, workflow, build environment and artifacts |
| Checksum | Digest identifying bytes; it does not independently prove how they were produced |
| Conformance | Satisfaction of the controls belonging to the recorded adopted baseline |
| Drift | Difference between declared/reviewed state and observed source, settings or publication |
| Idempotence | Repeating an operation with the same inputs leaves the same resulting state |
| Dry run | Validation and planned changes without applying them |
| Cleanup | Releasing/restoring owned state while preserving relevant failure evidence |
| Source-only | Distribution without artificial binary assets; profile evidence still applies |
| Template contract | Version of the reusable repository control baseline, independent of product version |
| Normative authority | The single maintained document or configuration that owns a rule |

---
[Home](Home.md) · [Previous](Troubleshooting.md) · [Next](Publish-the-wiki.md)
