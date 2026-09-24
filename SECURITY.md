<div align="center">

# 🔒 K-PRICING Security Policy

### Financial analytics and instrument pricing for Excel/VBA

[![Reporting](https://img.shields.io/badge/Reporting-Private-d97706?style=for-the-badge)](#reporting-a-vulnerability)
[![Support](https://img.shields.io/badge/Support-Latest_release-217346?style=for-the-badge)](#supported-versions)
[![Scope](https://img.shields.io/badge/Scope-Source_%7C_Releases_%7C_Automation-0969da?style=for-the-badge)](#security-scope)
[![Disclosure](https://img.shields.io/badge/Disclosure-Coordinated-6f42c1?style=for-the-badge)](#coordinated-disclosure)

<br>

**Protect users · Minimize exposure · Preserve evidence · Coordinate disclosure**

</div>

---

This document is authoritative for **vulnerability scope, private reporting,
security triage, coordinated disclosure and safe harbor**. Contribution workflow
is owned by [`CONTRIBUTING.md`](CONTRIBUTING.md); release sequence and provenance
are owned by [`RELEASING.md`](RELEASING.md) and
[`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md).

> [!IMPORTANT]
> A security policy does not make macros, workbooks, add-ins or release artifacts
> inherently trustworthy. Establish provenance and apply organizational security
> controls before enabling executable content.

## 🧭 Security model

The project assumes Microsoft Excel, the operating system and VBA runtime are
trusted; the user is authorized to run the project; macros are enabled through
an approved trust mechanism; and source/artifacts come from an official channel.

These are trust boundaries, not guarantees. VBA projects running in the same
Excel process are not isolated security sandboxes.

The migrated date layer parses and calculates dates, inspects the calling
workbook's date system and returns scalar or array results. The regression
harness creates disposable host fixtures and must restore the Excel state it
owns. See [the date contract](docs/DATE_LAYER_CONTRACT.md) and
[test guidance](tests/README.md). Repository automation separately accesses
GitHub and approved tool-download endpoints. Reassess this surface when changing
host integration or introducing packaging/UI.

The repository is public. CodeQL analyzes tooling languages, not VBA; Scorecard
publishes the default-branch posture. See the dated
[setup verification](docs/SETUP_VERIFICATION.md) for observed security settings.

<a id="supported-versions"></a>

## 📦 Supported versions

| Source state | Security support |
| --- | --- |
| Latest tagged functional release | None exists yet; no production-supported version |
| Release candidate | ⚠️ Testing / best-effort remediation |
| `main` | ⚠️ Development code / best effort |
| Older tagged releases | ❌ Normally unsupported; upgrade first |
| Modified copies / unofficial mirrors | ❌ Unsupported unless reproduced in official supported source |

If no functional release exists yet, the project is pre-release and has no
production-supported security version. Reports must identify an exact release
tag or full commit SHA; relative descriptions such as “latest” are insufficient.

<a id="reporting-a-vulnerability"></a>

## 📣 Reporting a vulnerability

Do **not** disclose a suspected vulnerability in a public issue, discussion,
pull request, commit message, Wiki page, sample workbook, screenshot or release
thread.

Use [GitHub private vulnerability reporting](https://github.com/danielep71/K-PRICING/security/advisories/new).
It was enabled and read back on 2026-09-24. Do not put vulnerability details in
public issues or pull requests.

Include only the information needed to assess the issue:

| Evidence | Requested detail |
| --- | --- |
| Identity | Repository, exact tag/SHA, component/procedure and affected artifact |
| Environment | Excel/Office build, bitness, Windows version and deployment model |
| Impact | Confidentiality, integrity, availability, execution or supply-chain consequence |
| Reproduction | Minimal steps using synthetic data |
| Exploitability | Preconditions, privileges, user interaction and persistence |
| Mitigation | Tested workaround/containment if known |
| Evidence | Sanitized diagnostics, hashes or proof of concept |

Never send client, employer, counterparty, student, production or personal
workbooks. Remove credentials, personal data, internal paths, links,
connections, document metadata, cached values and unrelated content.

If a secret has been exposed, revoke/rotate it immediately before improving the
report.

## ⏱️ Response process

This project may be maintained by one person; targets are best-effort, not a
contractual SLA.

| Stage | Target |
| --- | --- |
| Acknowledge | Within 5 business days |
| Initial scope/severity assessment | Within 10 business days after sufficient evidence |
| Active-investigation update | At least every 14 days |
| Remediation/disclosure | Proportionate to severity, exploitability and validation needs |

The normal path is reproduce → scope affected versions/artifacts → contain risk →
fix → add regression/fault evidence → validate in the relevant host → publish a
corrected release/advisory when appropriate.

## 🎯 Security issue or ordinary defect?

When uncertain, report privately. Security-relevant reports include credible
risk of:

- unintended code execution or trust-boundary crossing;
- unauthorized reading, modification, deletion or disclosure of data;
- persistent/exploitable loss of availability;
- credential, token, signing-key, runner or automation compromise;
- malicious/substituted official release artifacts;
- provenance or validation bypass that can misrepresent unsafe output as trusted;
  or
- a correctness defect deliberately exploitable to defeat a security/integrity
  boundary.

An incorrect result, compatibility problem, bounded performance regression,
documentation error or recoverable UI defect is normally an ordinary bug unless
it creates concrete security impact.

<a id="security-scope"></a>

## 🛠️ Security scope

### In scope

- official source and committed executable/macro-enabled artifacts;
- official releases, archives, checksums, manifests and provenance claims;
- repository-owned build, test, validation, packaging and release tooling;
- GitHub Actions workflows, permissions, dependencies and project-managed
  credentials;
- documented runtime integrations/trust boundaries; and
- security/integrity behavior introduced by project code.

### Project-specific risk surfaces

Before release, replace or extend this list with verified project facts:

- **Runtime surface** — workbook/application/native/file/network/UI risks.
- **Artifact surface** — executable workbooks, add-ins or other distributed assets.
- **Automation surface** — runners, credentials, release jobs and third-party dependencies.

### Out of scope

- vulnerabilities in Microsoft Excel, Office, Windows, GitHub, Python or VBA
  themselves;
- organization-controlled endpoint, macro, access or deployment policy;
- malicious VBA already trusted in the same Excel process;
- unrelated workbooks, add-ins, dependencies or infrastructure;
- unsupported modified copies/mirrors/historical snapshots;
- compromised user credentials not exposed by this project; and
- ordinary defects without concrete security impact.

Upstream vulnerabilities belong with the responsible vendor/platform.

<a id="data-and-secrets"></a>

## 🔐 Data and secret handling

Never commit, upload, log or attach:

- passwords, API keys, personal access tokens, signing keys, certificates or
  connection strings;
- client/employer/counterparty/student/employee/personal data;
- proprietary source, models, workbooks, production extracts or licensed data;
- internal URLs, machine-specific paths, environment dumps or unredacted
  screenshots; or
- exploit material beyond what is necessary to establish the issue.

Use synthetic data and a minimal reproduction. Excel files can contain sensitive
material outside visible cells, including document properties, names, hidden
sheets, VBA, cached values, queries, links and connections.

Repository secrets must use least privilege, remain unavailable to untrusted
pull-request code and be rotated after suspected exposure.

## 📦 Supply-chain boundary

Trusted distribution is limited to the official repository and its published
releases. Release actions, evidence schemas, checksums and tag-binding rules are
**not duplicated here**; maintainers must follow [`RELEASING.md`](RELEASING.md)
and [`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md).

Security-sensitive workflow changes require least-privilege permissions,
immutable dependency pins and explicit review. Do not run untrusted code on a
persistent credentialed Excel/Windows runner. Treat logs, screenshots,
workbooks, test artifacts and environment metadata as potentially sensitive.

## ✅ Safe-use guidance

Users should:

- preserve organization-approved macro security/deployment controls;
- obtain source/artifacts only from official channels and establish provenance;
- test with synthetic data in a controlled environment before production use;
- apply the project's documented supported-version and installation contract;
  and
- understand that numerical, pricing, timing or UI output is not itself an
  authentication, authorization, cryptographic or safety control.

<a id="coordinated-disclosure"></a>

## 📣 Coordinated disclosure

Avoid public disclosure while exploitability is being assessed, a fix is being
prepared, users have not had reasonable time to update, an exposed secret remains
valid, or a malicious artifact/runner remains reachable.

The maintainer and reporter should agree a plan based on severity, active
exploitation, remediation complexity, workarounds and validation time. The
maintainer may request a sanitized reproduction, more environment detail,
confirmation against a candidate fix or a reasonable embargo.

When remediation is available, the project may publish a GitHub Security
Advisory, corrected release, mitigation guidance and agreed reporter credit.

<a id="safe-harbor"></a>

## 🛡️ Good-faith research and safe harbor

Good-faith research is welcome when it:

- stays within project-owned source, artifacts and documented integrations;
- avoids privacy violations, destructive actions, persistence, social
  engineering and unnecessary access;
- stops after establishing the minimum required evidence;
- reports privately and promptly; and
- allows reasonable investigation/remediation time.

The project will not initiate or recommend legal action solely for research
conducted in good faith and consistently with this policy. This does not
authorize testing third-party systems or bind Microsoft, GitHub, an employer,
client or other third party.

No paid bug bounty is offered unless stated otherwise in writing.

## 📚 Related authorities

- [`CONTRIBUTING.md`](CONTRIBUTING.md) — contribution/review workflow
- [`CODE_OF_CONDUCT.md`](CODE_OF_CONDUCT.md) — participant behavior
- [`INSTALLATION.md`](INSTALLATION.md) — safe installation/upgrade/removal
- [`RELEASING.md`](RELEASING.md) — release sequence
- [`docs/RELEASE_EVIDENCE.md`](docs/RELEASE_EVIDENCE.md) — provenance/evidence schema
- [`docs/README.md`](docs/README.md) — complete documentation authority map

Conduct complaints and vulnerability reports are different: use the Code of
Conduct for participant behavior and this policy for software/security risk.

---

<div align="center">

### Security principle

**Trust deliberately · Run minimally · Protect secrets · Preserve evidence · Disclose responsibly**

<br>

Maintained by **Daniele Penza**

</div>
