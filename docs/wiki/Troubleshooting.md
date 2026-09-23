# 🧯 Troubleshooting

> **Guide, not policy:** [docs/README.md](../README.md) is authoritative.

Work from the first diagnostic and retain the exact command, source SHA and
sanitized output. Do not fix a failure by weakening an unrelated gate.

| Symptom | Check and next action |
| --- | --- |
| `gh` is not recognized | GitHub CLI is optional; use the web interface and Git credential manager |
| Git/Python is not found | Verify installation and PATH; reopen Git Bash and check versions |
| Bash continuation/array errors | Use Git Bash, not Command Prompt or PowerShell; those shells have different syntax |
| Dirty-tree initialization refusal | Review changes; commit intended setup inputs or use a fresh clean checkout; do not discard unrelated work |
| Unknown/missing/unused substitution | Compare the exact catalogue and argument categories in Initialization |
| Second initialization fails | Use identical recorded inputs from a clean committed tree; different inputs are not an upgrade |
| Expected maintenance tools disappeared | Check X entries in the inventory: template-only removal is intentional |
| New file not inspected | Stage the intended file; canonical inventory uses tracked paths |
| README-only source is rejected | Keep the mandatory facade, core and substantive tests; empty roles do not meet the contract |
| CRLF or encoding failure | Follow attributes and export rules; preserve CP1252-compatible VBA without BOM |
| VBE imports a duplicate module | Replace the existing component through the documented installation procedure; verify its exported name |
| API manifest mismatch | Compare normalized public signatures and declared roles; update the contract only for an intentional API change |
| Test run says dirty start | Stop the interrupted execution, use `ProjectTests.ResetProjectTests`, then run the complete suite |
| Counts or cleanup fail | Keep the entire output and diagnose the first failure; do not relabel an incomplete result PASS |
| Required check is waiting | Compare branch/path triggers and actual job context with the ruleset; preserve protection while repairing configuration |
| Ruleset read gives 403 | Record unverified access; obtain authorized administration access or a UI read-back |
| Provisioning digest changed | Re-read the new plan; concurrent/source changes invalidate the old approval |
| Provisioning partially applied | Inspect the durable journal and actual live state; create a new reviewed recovery plan |
| External link timeout/429/5xx | Treat as a transient observation; keep it separate from deterministic broken references |
| Wiki source check fails | Refresh catalogue entries and generated reference/sidebar, stage all intended files and rerun |
| Published wiki differs | Compare source SHA, manifest and actual page bytes; republish the reviewed bundle rather than editing online |
| Release rejects `0.0.0` | Finish release preparation with a nonzero version and dated changelog |
| Release evidence SHA mismatch | Certify the actual final candidate; do not retarget old run records by changing hashes |
| Asset mismatch or undeclared `dist/` file | Rebuild/retest when needed, inventory the complete payload and hash final bytes |
| Existing public tag is wrong | Follow the patch-release recovery policy; never force-move a public version tag |

## What to include in a help request

Use the appropriate bug, feature or documentation form. Include the repository,
profile, source revision, exact non-secret command, expected/actual result and
relevant environment. Redact private workbook data, credentials and personal
paths. Security issues go through the private route in
[SECURITY.md](../../SECURITY.md), not a public troubleshooting thread.

---
[Home](Home.md) · [Previous](Maintenance-and-migration.md) · [Next](Glossary.md)
