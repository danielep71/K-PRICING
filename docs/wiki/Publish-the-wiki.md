# 📤 Publish and maintain this wiki

> **Guide, not policy:** [Wiki Publication](../WIKI_PUBLICATION.md) is authoritative.

This page is for the template maintainer. New generated projects deliberately
remove this wiki source, catalogue, tooling and maintenance workflows. They
retain their operational documentation and may design their own documentation
system later.

## Source review

Edit the source pages in `docs/wiki/` on the reviewed development branch. Update
the catalogue when a page, tracked file or directory changes. Every page links
to a repository authority; clarify the authority first when changing a rule.
Keep the ordered sidebar and previous/next links coherent.

From the template root:

```bash
python tools/check_wiki.py --root . --write-reference
git add docs/wiki
python tools/check_wiki.py --root .
python tools/test_wiki.py -v
```

Review the generated diff, run documentation/repository checks, commit and wait
for the hosted results. The inventory uses the actual initializer to classify
profile transformations. It is not a manually guessed list of copied files.

## Export the reviewed revision

From a clean source checkout, choose a new directory outside that repository:

```bash
python tools/check_wiki.py --root . --export-dir ../wiki-export
```

The output contains every page, the sidebar and `Wiki-Source.json`. Repository
links name the exact source SHA; page links use stable Wiki slugs. The exporter
refuses an existing destination instead of overwriting earlier evidence.

## Publish and read back

Follow the complete clone/copy/review/commit/push procedure in
[Wiki Publication](../WIKI_PUBLICATION.md). Use the Wiki's existing default
branch and an authorized Git login. Review existing material before replacing
it; never force-push a documentation history.

Copy `Wiki-Source.json` alongside `Home.md` and the other Markdown files.
Selecting only Markdown files leaves out the source manifest and fails the
publication check even when every visible page is correct.

Fetch a fresh Wiki clone, then compare it from the same source checkout:

```bash
python tools/check_wiki.py --root . --published-dir ../wiki-readback
```

Record both the source and Wiki commit SHAs and the result. Open Home, the
sidebar and representative pages on GitHub to check navigation and rendering.
Do not treat local export or successful upload as read-back verification.

## Respond to drift

The weekly/manual observation distinguishes an unavailable network fetch from
deterministic mismatched pages. Fix sources through review and republish the
complete bundle; an online-only edit will be reported as drift. Source-SHA
changes also require a refreshed publication when comparing with the new
candidate. For a historical publication, check out its recorded source first.

## Complete acceptance

Retain a pilot of the adoption journey and its unresolved live/host steps.
After the final Markdown audit, regenerate the inventory for the final tree,
publish the audited revision and retain fresh read-back evidence. A development
edition is not a claim that the v1.2.0 release has already shipped.

---
[Home](Home.md) · [Previous](Glossary.md)
