# 🧭 Repository governance

![Governance](https://img.shields.io/badge/Governance-Versioned-00338D?style=flat-square)
![Labels](https://img.shields.io/badge/Issue_labels-20-1D76DB?style=flat-square)

Ongoing repository rules and maintenance guidance live here. One-time initialization instructions and the executable settings baseline live in [setup](../setup/REPOSITORY_SETUP.md).

| Guide | Purpose |
| --- | --- |
| [⚙️ Settings](SETTINGS.md) | Repository settings and branch/tag protection baseline. |
| [🏷️ Issue labels](ISSUE_LABELS.md) | Label catalogue, synchronization, drift detection, and issue triage. |
| [🔎 Pull request review](PULL_REQUESTS.md) | Final-head review before merging, and the exceptions to record. |
| [🧾 Excel run](EXCEL_RUN.md) | An Excel run recorded by hand, what the evidence check rejects, and when a run counts for another commit. |
| [📈 Traffic history](TRAFFIC.md) | Optional daily traffic collection on the `traffic-history` branch, with alerts. |
| [🚀 Repository setup](../setup/REPOSITORY_SETUP.md) | Preview, apply and verify the settings baseline, and what to keep, adapt or remove. |
| [📋 Settings configuration](../setup/repository-settings.json) | Machine-readable baseline used by the setup script. |
| [🔒 Security policy](../../SECURITY.md) | Private vulnerability reporting, response targets, and scope. |
| [📦 Distribution review](DISTRIBUTION.md) | What to check and record before making the repository public, sharing a copy or publishing a release. |
| [🛠️ Contributing](../../CONTRIBUTING.md) | How to open issues and pull requests, which branch they target, and local checks. |
| [🤝 Code of conduct](../../CODE_OF_CONDUCT.md) | Expected behaviour and how to report a problem. |

## Repository controls

Settings are defined in the [settings configuration](../setup/repository-settings.json). Label configuration, issue forms, and automation live under `.github/`.

Changes follow the protected pull-request process and must pass **Repository integrity**, which also checks every relative link and anchor in the Markdown files and that every external link uses HTTPS. The weekly **Link check** workflow requests the external links. The daily **Detect settings drift** workflow compares the live rulesets, and the other settings its read-only token can read, with the settings configuration ([scheduled drift check](SETTINGS.md#drift)). The **Closing keywords** workflow checks that no pull request into a release branch, and no integration of one, carries a closing keyword that would close an open issue ([release branches](PULL_REQUESTS.md#-release-branches)). The **CodeQL** workflow scans the Python, JavaScript and workflow code on pushes to the default branch, on pull requests into the default and release branches that change more than Markdown files and files under `docs/`, and weekly; its availability on private repositories is in the setup guide's [handover table](../setup/REPOSITORY_SETUP.md#recommended-handover). The [setup guide](../setup/REPOSITORY_SETUP.md) explains initialization, verification, and which files to retain after launch.
