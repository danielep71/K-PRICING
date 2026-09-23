# 📚 Wiki Source and Publication Contract

[![Source: reviewed](https://img.shields.io/badge/source-reviewed-217346)](wiki/Home.md)
[![Publication: exact SHA](https://img.shields.io/badge/publication-exact%20SHA-0969da)](#publication)

This is the authority for this template's wiki source, inventory, publication
and drift procedure. The wiki is a guided reading path. Normative project rules
remain in the repository documents linked from every page. The owner is the
template maintainer; generated repositories remove this maintenance system.

## Source and inventory

`docs/wiki/catalogue.json` schema 1 records the ordered page set and one entry
for each tracked file and ancestor directory. Names are stable ASCII slugs;
`Home` is first, and `_Sidebar.md` follows the same order. Each page has its own
authority notice. Do not rename a published page without reviewing incoming
links and a migration note.

Each inventory entry supplies purpose, kind and authority. Kind definitions
provide ownership, when to edit, invariants/unsafe edits, and validation. The
rendered reference includes all those fields without repeating the same rule
paragraph for every helper file. Per-profile lifecycle is computed by calling
the real initializer's dry-run renderer with its synthetic fixture values:
retained, transformed or removed. Generated-only paths are a separate table.
The social-preview retention exception is stated explicitly.

From the template root, stage intended new files, update the catalogue and run:

```bash
python tools/check_wiki.py --root . --write-reference
git add docs/wiki
python tools/check_wiki.py --root .
python tools/test_wiki.py -v
python tools/check_documentation.py --root .
python tools/check_repo.py --root .
```

The default check is read-only. It rejects missing/extra inventory paths, stale
reference or sidebar, missing pages, duplicate names/JSON keys, missing
authorities and missing authority notices. The canonical gate validates local
links/anchors. The documentation gate validates literal tool commands and
registered references. Review prose and intent separately; a passing parser
does not certify the whole guide's accuracy.

`--write-reference` rewrites only the two generated Markdown files. Review their
diff and commit with the source change. It does not write to GitHub.

<a id="publication"></a>

## Controlled publication

1. Finish the source review and local tests. Commit the complete candidate and
   wait for its applicable hosted checks. Preserve the source SHA and run URLs.
2. Use a clean checkout of that exact source. Run the export below into a new,
   nonexistent directory outside the source repository.
3. Review the exported pages. Relative repository links become full-SHA links;
   page links become GitHub Wiki slugs. Every page identifies the source commit.
   `Wiki-Source.json` records the source, repository and SHA-256 of every page,
   including the sidebar. It deliberately does not hash itself.
4. Clone the separate wiki repository and record its current default-branch
   head. Inspect existing pages before replacing content. Preserve unrelated
   material or reconcile it explicitly; the verifier rejects extra files.
5. Copy the reviewed bundle into the wiki checkout, inspect the staged diff,
   commit and push without force to its existing default branch.
   Include `Wiki-Source.json` as well as every Markdown file. File-manager
   selections filtered to `*.md` omit the manifest and fail publication checks.
6. Fetch a fresh clone and compare its actual bytes using `--published-dir`.
   Record source SHA, wiki SHA, hosted runs, comparison result and web-page
   read-back. A successful push alone is not publication verification.

Commands use Git Bash and run from the clean **template source root**. Replace
the repository path and review commit before running; the CLI itself never
fetches, commits, pushes or supplies credentials:

```bash
git status --short
git rev-parse HEAD
python tools/check_wiki.py --root . --export-dir ../wiki-export
git clone https://github.com/OWNER/TEMPLATE-REPOSITORY.wiki.git ../wiki-checkout
git -C ../wiki-checkout branch --show-current
git -C ../wiki-checkout rev-parse HEAD
```

After reviewing existing content, copy the exported files with your file manager
or the following commands. This first-publication example assumes the existing
wiki contains only its reviewed Home placeholder. For an established wiki,
review removed pages separately; do not delete all files blindly.

```bash
cp ../wiki-export/*.md ../wiki-export/Wiki-Source.json ../wiki-checkout/
python tools/check_wiki.py --root . --published-dir ../wiki-checkout
git -C ../wiki-checkout add -- '*.md' Wiki-Source.json
git -C ../wiki-checkout diff --cached --check
git -C ../wiki-checkout diff --cached --stat
git -C ../wiki-checkout commit -m "Publish wiki from REVIEWED_SOURCE_SHA"
git -C ../wiki-checkout push origin HEAD
git clone https://github.com/OWNER/TEMPLATE-REPOSITORY.wiki.git ../wiki-readback
python tools/check_wiki.py --root . --published-dir ../wiki-readback
git -C ../wiki-readback rev-parse HEAD
```

Stop on the first failed command. If authentication fails, leave the prepared
bundle and reviewed commit available; use authorized Git credentials or a
maintainer-operated push. Never put a token in the URL or source. The GitHub
connector's source-repository write capability does not imply Wiki Git access.
If no wiki repository exists, an authorized maintainer enables Wiki and creates
its initial Home page first. See
[GitHub's wiki Git workflow](https://docs.github.com/en/communities/documenting-your-project-with-wikis/adding-or-editing-wiki-pages).

## Drift and reliability

`wiki-checks.yml` blocks its own source-check job on deterministic findings and
runs without live Wiki access on pushes/PRs. It does not change the published
reusable-workflow pin or add a new required branch context automatically.

`wiki-drift.yml` is a separate weekly/manual observation. It clones the public
wiki anonymously under a bounded job timeout, then compares against the clean
source revision checked out by that run. Clone failure is reported as
`UNAVAILABLE`, with no claim about page correctness. A successful fetch followed
by a byte/path/SHA mismatch is `DRIFT`; it fails that observation job. A match is
`PASS`. None of these network outcomes replaces the deterministic PR checks.

The comparison is intentionally strict about source identity: a publication
from another source SHA is stale relative to the requested source, even if most
prose is unchanged. Before final release, publish from the final reviewed
default-branch commit. A development edition remains visibly tied to its source
and must be refreshed after the final merge. Operators can reproduce an older
publication by checking out its recorded source commit and comparing there.

Outbound source-page links also use the separate policies in
[Documentation Checks](DOCUMENTATION_CHECKS.md). A wiki-byte match says nothing
about internet availability. Do not convert timeout/access failures to PASS.

## Pilot and final acceptance

Retain a clean-room record of starting revision, profile, numbered guide steps,
commands, elapsed time, outputs, gaps/corrections and final verdict. Local fresh
repositories can prove preview/apply/idempotence and source checks for all three
profiles; simulated fixtures cannot prove live GitHub creation, governance,
Excel execution or a published first release. Keep those distinctions explicit.

The final Markdown audit in #44 reconciles this guide against the finished
v1.2.0 tree. Refresh the complete inventory when any tracked file is added or
removed, including deletion of the temporary implementation plan. Record final
publication read-back against that audited source before closing final wiki
acceptance. Do not edit wiki pages online as an independent policy stream.
