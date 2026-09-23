# 🏠 From template to first release

> **Guide, not policy:** [docs/README.md](../README.md) is authoritative.

This guide is for a maintainer creating a **new** Excel/VBA repository. Work through
the numbered journey in the sidebar. It explains what to do and where the precise
rules live; it does not replace those rules.

> **Version scope:** this edition documents the reviewed v1.2.0 development
> contract. While that work is on a release branch, GitHub's default-branch
> template may still contain an earlier edition. Check the source link at the
> foot of the published page and the template's default branch before adopting.

## What you will produce

A project with one declared profile, a public facade, internal core, substantive
regression tests, reviewed GitHub settings, passing source gates, and a release
whose evidence identifies the exact source tested.

The template does not supply your application's business logic or a prebuilt
workbook. The ratio example is a working starting point, not a pricing library.

## Before starting

| Requirement | Verify |
| --- | --- |
| Git and Git Bash on Windows | `git --version`; commands below use Bash syntax |
| Python 3.10 or later | `python --version`; use the same interpreter throughout |
| Excel desktop for host validation | Check Office version, build and bitness |
| GitHub access | Create the intended repository and administer its settings |
| Editor and VBE access | Edit text exports and import modules into a test workbook |
| Private security contact | A monitored address or private reporting route |

The GitHub CLI is optional. If `gh` is not installed, use the GitHub web interface
and Git's credential manager. Never paste tokens into commands, issues or this wiki.

## Journey and checkpoints

1. **Choose** a profile and its owner/state boundary.
2. **Create** only the new repository; record the template source.
3. **Initialize** from a clean tree: preview, review, apply, commit, verify no-op.
4. **Configure GitHub** and read settings back.
5. **Run gates**, then implement and test the project's own behavior.
6. **Certify the release**, tag the exact candidate and verify downloads.

Stop at the first failed checkpoint. A skipped network check or unavailable
Excel installation is not a passing result.

Use the file reference when a filename is unfamiliar, the tool reference to
understand automation, and troubleshooting to diagnose failures. The publication
page is for maintainers of this template's wiki; generated projects remove that
template-maintenance source rather than inheriting a second documentation system.

---
[Home](Home.md) · [Next](Choose-a-profile.md)
