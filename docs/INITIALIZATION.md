# 🚀 Repository Initialization

[![Mode: dry-run first](https://img.shields.io/badge/mode-dry--run%20first-217346)](#-safety-model)
[![Profiles: 3](https://img.shields.io/badge/profiles-3-6f42c1)](#-initialize-one-profile)
[![Writes: staged](https://img.shields.io/badge/writes-staged-217346)](#-deterministic-transformations)
[![Verification: self-tested](https://img.shields.io/badge/verification-self--tested-1D76DB)](#-verification)

This document is the authoritative contract for turning a clean repository
created from this template into one initialized project. Initialization changes
versioned files only. It does not configure GitHub labels, metadata, secrets,
rulesets, environments, or other live settings.

## 🛡️ Safety Model

`tools/initialize_repository.py` is dependency-free and dry-run-first. It:

1. requires a clean Git working tree;
2. validates the complete input set before rendering any file;
3. renders all changes in memory and reports content digests;
4. changes files only when `--apply` is present;
5. attempts to restore original files if a filesystem write fails; and
6. records the exact non-secret initialization inputs in
   `.github/initialization.json`.

Missing, unknown, duplicated, category-incompatible, and unused substitutions
are errors. Values may not contain line breaks or reserved template syntax.

Replacements are staged and applied per file; this is not a repository-wide
filesystem transaction. Rollback itself can fail if the filesystem remains
unwritable, and newly created directories may remain. After any apply failure,
inspect the complete working tree against the clean starting commit before
retrying. Do not infer successful restoration from an interrupted operation.

## 🧬 Canonical Token Grammar

A token is two opening braces, one uppercase name matching
`[A-Z][A-Z0-9_]*`, and two closing braces. Square brackets retain their normal
Markdown, checklist, and changelog meanings; they are not template tokens.

The machine-readable catalogue is
[`.github/repository-profile.json`](../.github/repository-profile.json). Every
catalogued token has exactly one category and a description explaining why it
exists.

| Category | Behavior |
| --- | --- |
| `required` | Supply exactly once with `--set NAME=value`; initialization fails if absent. |
| `optional` | Supply at most once with `--set`; omitting it removes its complete optional block. |
| `profile-specific` | Do not supply it; the initializer derives it from the selected profile. |
| `repeatable` | Supply zero or more times with `--add NAME=value`; omitting it removes its complete repeatable block. |

### 📋 Catalogue

| Name | Category | Purpose |
| --- | --- | --- |
| `PROJECT_NAME` | Required | Human-readable name in documentation and release text |
| `PROJECT_TAGLINE` | Required | Short identity line below the project name |
| `PROJECT_DESCRIPTION` | Required | One-sentence supported problem and audience |
| `REPOSITORY_PATH` | Required | GitHub `owner/name` used by clone commands, badges, and links |
| `MAINTAINER_NAME` | Required | Person or organization responsible for maintained decisions |
| `SUPPORT_CONTACT` | Required | Private-reporting email address or maintained HTTPS URL |
| `COPYRIGHT_YEAR` | Required | Four-digit MIT-licence copyright year |
| `SOCIAL_PREVIEW_PATH` | Optional | Tracked repository-relative banner image |
| `PROFILE_NAME` | Profile-specific | Human-readable selected profile |
| `PROFILE_PURPOSE` | Profile-specific | Selected profile's ownership boundary |
| `PROFILE_SOURCE_CONTRACT` | Profile-specific | Selected profile's expected production structure |
| `PROFILE_EVIDENCE` | Profile-specific | Selected profile's minimum runtime evidence |
| `ADDITIONAL_TEST_COMMAND` | Repeatable | Additional project-specific validation command |
| `KNOWN_LIMITATION` | Repeatable | Honest user-visible limitation rendered as a list item |

Tokens are permitted only in documentation and licence text. They are
prohibited in VBA exports, identifiers, workflows, executable scripts, and
structured configuration. The initializer rewrites the canonical issue
chooser's template-repository security URL to the generated repository without
placing a token in YAML. VBA components therefore use fixed, compile-safe
identifiers; a project may rename them later as an explicit source change.

## 🧭 Initialize One Profile

Run this command from a clean repository root. Dry-run is the default:

```bash
python3 tools/initialize_repository.py --profile library \
  --set PROJECT_NAME="Example Project" \
  --set PROJECT_TAGLINE="A concise project identity" \
  --set PROJECT_DESCRIPTION="One sentence describing the supported problem and audience." \
  --set REPOSITORY_PATH="owner/repository" \
  --set MAINTAINER_NAME="Example Maintainer" \
  --set SUPPORT_CONTACT="security@example.com" \
  --set COPYRIGHT_YEAR="2026"
```

Use exactly one profile:

| Profile | `--profile` value | Boundary |
| --- | --- | --- |
| Library | `library` | Reusable callable VBA without an owned end-user shell |
| UI component | `ui-component` | Embeddable component with a bounded interactive surface |
| Application | `application` | End-to-end workbook or add-in owning deployment and lifecycle |

Review every planned create, update, and delete operation and its before/after
SHA-256 digest. Repeat the identical command with `--apply` only when that plan
is correct. Review and stage the applied changes, run the repository gate, and
commit the initialized tree before repeating the command. A second run with
the same arguments from that clean committed tree returns `no-op`; different
inputs fail rather than silently rewriting an initialized repository.

### ➕ Optional and Repeatable Values

Add a tracked social-preview image only when it already exists:

```bash
--set SOCIAL_PREVIEW_PATH="assets/social-preview.png"
```

Repeat list inputs in command order:

```bash
--add ADDITIONAL_TEST_COMMAND="python3 tools/check_repo.py --root ." \
--add KNOWN_LIMITATION="Excel for macOS has not been tested."
```

Omitted optional and repeatable values remove their complete marked blocks;
empty placeholder prose is never retained.

## ⚙️ Deterministic Transformations

An applied initialization:

- sets `mode` to `generated`, selects one profile, and records the repository;
- substitutes every required and derived profile value;
- retains only the selected profile block;
- retains supplied optional and repeatable blocks and removes unused ones;
- deletes every path declared under `placeholders.template_only_paths`, subject
  only to the documented social-preview retention exception; template-maintainer
  checker-development and semantic policy-coverage tooling is deliberately in
  that set, while operational repository/release/VBA gates remain in generated
  projects;
- resets the changelog's `Unreleased` section so template-construction history
  is not attributed to the generated project;
- resets `VERSION` to the `0.0.0` development sentinel;
- creates explanatory files in currently empty profile-required directories;
- writes `.github/initialization.json`; and
- leaves the initializer available for idempotence verification.

Explanatory profile-directory files are structural guidance only. They never
satisfy the generated VBA contract: every selected profile must retain the
registered public façade, internal core, and regression module declared by its
`vba_contract`. Profile-specific classes, forms, Ribbon XML, workbook modules,
and examples remain optional unless the selected contract explicitly adds them.

## 🧰 Manual Fallback

The script is authoritative, but the transformation remains transparent and can
be reproduced manually:

1. Start from a clean clone and save the pre-initialization commit SHA.
2. Read the placeholder catalogue and profile values in
   `.github/repository-profile.json`.
3. Replace every required token consistently and copy the selected profile's
   derived values.
4. Keep only the selected profile blocks. Remove all other profile blocks,
   omitted optional/repeatable blocks, template-only blocks, and their marker
   lines.
5. Delete every path listed under `template_only_paths`, except
   `assets/social-preview.png` when that exact path was deliberately supplied
   as `SOCIAL_PREVIEW_PATH`. No other template-only path may be retained through
   that placeholder.
6. Reset `CHANGELOG.md` under `Unreleased` to project-owned content only and set
   `VERSION` to the `0.0.0` development sentinel.
7. Set configuration mode, profile, and repository; create
   `.github/initialization.json` using the same schema as the initializer.
8. Add an explanatory or substantive tracked file to every directory required
   by the selected profile.
9. Search for the configured token grammar, template identity, marker prefix,
   and deleted template-only paths. All searches must be empty outside the
   policy files that define those checks.
10. Review `git diff --check`, stage the candidate tree, and run the checker.

Manual initialization is incomplete if its resulting tree differs in policy or
content from what the deterministic initializer would produce for the same
inputs.

## ✅ Verification

```bash
python3 tools/initialize_repository.py --root . --self-test
python3 tools/check_repo.py --root . --self-test
python3 tools/check_repo.py --root . \
  --output test-results/static-checks.json \
  --summary test-results/static-checks.md
```

The initializer self-test exercises missing, unknown, and unused inputs; dry-run
immutability; application; second-run idempotence; template-only cleanup; and a
green generated tree for every profile. It also verifies that template-maintainer
checker-development/policy-coverage files are absent after generation while the
operational gate set remains. For each profile it proves that a
README-only tree and removal of the façade, core, or test module fail only the
named `generated-vba-contract` rule, while removal of the optional example still
passes.

After initialization, configure the live repository settings that a GitHub
template cannot inherit. Follow
[`POST_CREATION_CHECKLIST.md`](POST_CREATION_CHECKLIST.md) and preserve read-back
evidence of the applied state.

---

**Initialization principle:** validate everything, preview every mutation, and apply one reproducible profile from a clean tree.
