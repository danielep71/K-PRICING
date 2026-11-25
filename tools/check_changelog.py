#!/usr/bin/env python3
"""Check CHANGELOG.md and VERSION against the rules of RELEASING.md.

It never tags or publishes a release. CHANGELOG.md starts with an ``## [Unreleased]`` section; each release
heading is ``## [X.Y.Z] - YYYY-MM-DD`` with a real date, newest first by
version and date, with a link reference to its tag's release page. VERSION
does not exist before the first dated release; from then on it holds one
``X.Y.Z`` line, the newest release. With --repository, every link reference
to GitHub, not only those of the version headings, must also point into that
repository.

It also checks that the prose agrees with that state:

- the introduction of CHANGELOG.md, before its first level-2 heading, says
  "No version has been released" exactly while it has no dated release, and
  no version section holds a placeholder line such as TBD;
- the *Current status* table of RELEASING.md, when present, names no release
  and no VERSION before the first release, and the newest one after it;
- README.md, when present, holds no "Release Candidate" and no unchecked
  task-list item.

A released section is frozen: once its tag vX.Y.Z exists, the section, from
its heading to the next level-2 heading, must equal that section in
CHANGELOG.md at the tag. Only line endings are normalised. A section that
holds a dated note "Corrected after release on YYYY-MM-DD: <reason>", with a
real calendar date, may differ, and the difference is reported as a note, so
a reviewed correction stays visible. A
release whose tag does not exist yet is a note, since the tag is created
after the changelog is dated; so is a directory that is not a Git checkout,
where nothing can be compared. A vX.Y.Z tag on a commit that HEAD contains,
whose CHANGELOG.md released [X.Y.Z], fails when the section is no longer in
the changelog, so a released section cannot be deleted or renamed; a newer
tag does not apply to an older revision.
"""

from __future__ import annotations

import argparse
import datetime
import itertools
import re
import sys
from collections.abc import Sequence
from pathlib import Path

from gitio import text_or_none

ROOT = Path(__file__).resolve().parents[1]
CHANGELOG = "CHANGELOG.md"
VERSION = "VERSION"
SEMVER = r"(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)"
VERSION_HEADING = re.compile(r"^## \[([^\]]+)\](.*)$", re.M)
# A level-2 heading that looks like a version, bracketed or not; policy headings
# start with a word and are not matched.
VERSION_LIKE = re.compile(r"^##[ \t]+(?:\[|v?\d|unreleased\b).*$", re.M | re.I)
RELEASE_SUFFIX = re.compile(r" - (\d{4}-\d{2}-\d{2})")
RELEASING = "RELEASING.md"
README = "README.md"
NO_RELEASE = "No version has been released"
# A list item or paragraph holding only a placeholder, with optional emphasis.
PLACEHOLDER = re.compile(r"^\s*(?:[-*]\s+)?[_*]*(?:tbd|todo|nothing yet|\.\.\.)[_*.]*\s*$",
                         re.I | re.M)
UNCHECKED_TASK = re.compile(r"^\s*[-*+]\s+\[ \]\s", re.M)
# A whole version number, optionally after "v", and not part of a longer one
# such as 10.2.0, 0.2.0.1 or 0.2.0-rc.1; a full stop after it ends a sentence.
VERSION_TOKEN = re.compile(rf"(?<![\w.])v?({SEMVER})(?![\w-]|\.\d)")
# The line that marks a reviewed change to a released section.
CORRECTION = re.compile(r"^Corrected after release on (\d{4}-\d{2}-\d{2}): \S.*$")
STATUS_ROW = re.compile(r"^\|\s*(Released versions|`VERSION`)\s*\|(.*)\|\s*$", re.M)


def valid_date(text: str) -> bool:
    try:
        datetime.date.fromisoformat(text)
    except ValueError:
        return False
    return True


def link_problems(text: str, label: str, repository: str | None) -> list[str]:
    """A release's link reference points at its tag's release page."""
    match = re.search(rf"^\[{re.escape(label)}\]: (\S+)$", text, re.M)
    if not match:
        return [f"{CHANGELOG}: missing link reference for [{label}]"]
    url = match.group(1)
    problems = []
    if repository and not url.lower().startswith(f"https://github.com/{repository.lower()}/"):
        problems.append(f"{CHANGELOG}: the link reference for [{label}] does not point to "
                        f"{repository}: {url}")
    if label != "Unreleased" and not url.endswith(f"/releases/tag/v{label}"):
        problems.append(f"{CHANGELOG}: the link reference for [{label}] must end with "
                        f"/releases/tag/v{label}")
    return problems


LINK_REFERENCE = re.compile(r"^\[([^\]]+)\]:[ \t]*(\S+)", re.M)
GITHUB = "https://github.com/"


def foreign_references(text: str, repository: str, skip: set[str]) -> list[str]:
    """Link references to GitHub outside ``repository``, such as a copied issue link.
    The references in ``skip``, those of the version headings, are left to
    link_problems."""
    own = f"{GITHUB}{repository.lower()}/"
    return [f"{CHANGELOG}: the link reference for [{label}] does not point to {repository}: {url}"
            for label, url in LINK_REFERENCE.findall(text)
            if label not in skip and url.lower().startswith(GITHUB)
            and not url.lower().startswith(own)]


def released_versions(text: str) -> list[str]:
    return [label for label, _ in VERSION_HEADING.findall(text) if label != "Unreleased"]


def check_changelog(text: str, repository: str | None = None) -> list[str]:
    headings = VERSION_HEADING.findall(text)
    findings = [f"{CHANGELOG}: version heading must be '## [Unreleased]' or "
                f"'## [X.Y.Z] - YYYY-MM-DD': {line}"
                for line in VERSION_LIKE.findall(text) if not VERSION_HEADING.fullmatch(line)]
    if not headings or headings[0] != ("Unreleased", ""):
        findings.append(f"{CHANGELOG}: the first version heading must be '## [Unreleased]'")
    seen: set[str] = set()
    releases: list[tuple[tuple[int, ...], str]] = []
    for label, suffix in headings:
        if label in seen:
            findings.append(f"{CHANGELOG}: duplicate heading [{label}]")
        seen.add(label)
        if label == "Unreleased":
            if suffix:
                findings.append(f"{CHANGELOG}: '## [Unreleased]' takes no date or suffix")
            continue
        dated = RELEASE_SUFFIX.fullmatch(suffix)
        if not re.fullmatch(SEMVER, label) or not dated:
            findings.append(f"{CHANGELOG}: release heading must be '## [X.Y.Z] - YYYY-MM-DD': "
                            f"[{label}]{suffix}")
            continue
        if not valid_date(dated.group(1)):
            findings.append(f"{CHANGELOG}: [{label}] date {dated.group(1)} is not a calendar date")
        releases.append((tuple(int(part) for part in label.split(".")), dated.group(1)))
        findings.extend(link_problems(text, label, repository))
    for newer, older in itertools.pairwise(releases):
        if newer[0] <= older[0] or newer[1] < older[1]:
            findings.append(f"{CHANGELOG}: releases must be listed newest first, by version "
                            "and date")
            break
    if "Unreleased" in seen:
        findings.extend(link_problems(text, "Unreleased", repository))
    if repository:
        findings.extend(foreign_references(text, repository, {label for label, _ in headings}))
    return findings


def check_version(root: Path, released: list[str]) -> list[str]:
    """VERSION exists from the first release on and names the newest one."""
    path = root / VERSION
    if not path.is_file():
        return [f"{VERSION} is missing; {CHANGELOG} releases [{released[0]}]"] if released else []
    raw = path.read_text(encoding="utf-8")
    if not re.fullmatch(SEMVER + r"\n", raw):
        return [f"{VERSION} must hold one X.Y.Z line ending in a newline"]
    if not released:
        return [f"{VERSION} exists but {CHANGELOG} has no dated release; it is created "
                "when a release is prepared (RELEASING.md)"]
    if raw.strip() != released[0]:
        return [f"{VERSION} {raw.strip()} differs from the newest {CHANGELOG} release "
                f"[{released[0]}]"]
    return []


def changelog_state_problems(text: str, released: list[str]) -> list[str]:
    """The introduction's release statement, and placeholders in version sections.

    Only the introduction is read for the statement, so that an entry or a
    policy section may quote it without contradicting a release.
    """
    problems = []
    introduction = re.split(r"^## ", text, maxsplit=1, flags=re.M)[0]
    if released and NO_RELEASE in introduction:
        problems.append(f"{CHANGELOG}: the introduction says '{NO_RELEASE}' but the changelog "
                        f"releases [{released[0]}]")
    if not released and NO_RELEASE not in introduction:
        problems.append(f"{CHANGELOG}: has no dated release but its introduction does not say "
                        f"'{NO_RELEASE}'")
    first = VERSION_HEADING.search(text)
    sections = text[first.start():] if first else ""
    for line in PLACEHOLDER.findall(sections):
        problems.append(f"{CHANGELOG}: placeholder line in a version section: {line.strip()}")
    return problems


def releasing_problems(root: Path, released: list[str]) -> list[str]:
    """The Current status rows of RELEASING.md agree with the changelog."""
    path = root / RELEASING
    if not path.is_file():
        return []
    section = re.search(r"^##[^#\n]*Current status[^\n]*\n(.*?)(?=^## |\Z)",
                        path.read_text(encoding="utf-8"), re.M | re.S)
    if section is None:
        return [f"{RELEASING}: has no Current status section"]
    rows = {name: cell.strip() for name, cell in STATUS_ROW.findall(section.group(1))}
    problems = []
    for name in ("Released versions", "`VERSION`"):
        cell = rows.get(name)
        if cell is None:
            problems.append(f"{RELEASING}: Current status has no '{name}' row")
        elif released and released[0] not in VERSION_TOKEN.findall(cell):
            problems.append(f"{RELEASING}: Current status '{name}' does not name the newest "
                            f"release [{released[0]}]: {cell}")
        elif not released and name == "Released versions" and cell != "None":
            problems.append(f"{RELEASING}: Current status '{name}' must be None before the "
                            f"first release: {cell}")
        elif not released and name == "`VERSION`" and not cell.startswith("Does not exist"):
            problems.append(f"{RELEASING}: Current status '{name}' must say it does not exist "
                            f"before the first release: {cell}")
    return problems


def readme_problems(root: Path) -> list[str]:
    """README.md describes no unfinished release."""
    path = root / README
    if not path.is_file():
        return []
    text = path.read_text(encoding="utf-8")
    problems = []
    if re.search(r"release[ _-]candidate", text, re.I):
        problems.append(f"{README}: mentions a Release Candidate")
    if UNCHECKED_TASK.search(text):
        problems.append(f"{README}: holds an unchecked task-list item")
    return problems


def sections(text: str) -> dict[str, tuple[int, list[str]]]:
    """Each version section's first line number and its lines, from its heading to
    the next level-2 heading, without trailing blank lines or link references."""
    lines = text.replace("\r\n", "\n").split("\n")
    found: dict[str, tuple[int, list[str]]] = {}
    label: str | None = None
    for number, line in enumerate(lines, start=1):
        if line.startswith("## "):
            match = VERSION_HEADING.fullmatch(line)
            label = match.group(1) if match else None
            if label is not None and label not in found:
                found[label] = (number, [])
            elif label is not None:
                label = None  # A duplicate heading: check_changelog reports it.
        if label is not None:
            found[label][1].append(line)
    for _, body in found.values():
        while body and (not body[-1].strip() or LINK_REFERENCE.match(body[-1])):
            body.pop()
    return found


def frozen_problems(root: Path, text: str, released: list[str]) -> tuple[list[str], list[str]]:
    """Each released section against its content at its tag: findings and notes.

    A tagged release whose section the changelog no longer holds is a finding
    too, so that deleting or renaming a section cannot escape the comparison.
    Only tags on commits that HEAD contains count: an older revision, checked
    out after a newer tag was created, never held the newer section."""
    if text_or_none(root, "rev-parse", "--git-dir") is None:
        return [], [f"{CHANGELOG}: not a Git checkout, so released sections were not compared "
                    "with their tags"] if released else []
    findings: list[str] = []
    notes: list[str] = []
    current = sections(text)
    for tag in sorted((text_or_none(root, "tag", "--list", "v*") or "").split()):
        label = tag[1:]
        if not re.fullmatch(SEMVER, label) or label in released:
            continue
        if text_or_none(root, "merge-base", "--is-ancestor", f"{tag}^{{commit}}",
                        "HEAD") is None:
            continue
        tagged_text = text_or_none(root, "show", f"{tag}:{CHANGELOG}")
        if tagged_text is not None and label in released_versions(tagged_text):
            findings.append(f"{CHANGELOG}: [{label}] was released with tag {tag} but its "
                            "section is gone; a released section is frozen")
    for label in released:
        tag = f"v{label}"
        if text_or_none(root, "rev-parse", "--verify", "--quiet",
                        f"refs/tags/{tag}^{{commit}}") is None:
            notes.append(f"{CHANGELOG}: [{label}] has no tag {tag} yet, so it is not frozen")
            continue
        tagged_text = text_or_none(root, "show", f"{tag}:{CHANGELOG}")
        tagged = sections(tagged_text).get(label) if tagged_text is not None else None
        if tagged is None:
            findings.append(f"{CHANGELOG}: [{label}] is not in {CHANGELOG} at its tag {tag}")
            continue
        start, lines = current[label]
        if lines == tagged[1]:
            continue
        correction = next((line for line in lines if (match := CORRECTION.fullmatch(line))
                           and valid_date(match.group(1))), None)
        if correction is not None:
            notes.append(f"{CHANGELOG}: [{label}] differs from its tag {tag}: {correction}")
            continue
        offset = next((index for index, (now, then) in enumerate(zip(lines, tagged[1],
                                                                       strict=False))
                       if now != then), min(len(lines), len(tagged[1])))
        findings.append(f"{CHANGELOG}:{start + offset}: [{label}] differs from its content at "
                        f"tag {tag}; a released section is frozen (add 'Corrected after "
                        "release on YYYY-MM-DD: <reason>' only for a reviewed correction)")
    return findings, notes


def check(root: Path, repository: str | None = None) -> tuple[list[str], int]:
    findings, released, _ = check_with_notes(root, repository)
    return findings, released


def check_with_notes(root: Path,
                     repository: str | None = None) -> tuple[list[str], int, list[str]]:
    path = root / CHANGELOG
    if not path.is_file():
        return [f"{CHANGELOG} is missing"], 0, []
    text = path.read_text(encoding="utf-8")
    released = released_versions(text)
    frozen, notes = frozen_problems(root, text, released)
    findings = (check_changelog(text, repository) + check_version(root, released)
                + changelog_state_problems(text, released) + releasing_problems(root, released)
                + readme_problems(root) + frozen)
    return findings, len(released), notes


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--repository", metavar="OWNER/REPOSITORY",
                        help="require every link reference to GitHub to point into this repository")
    args = parser.parse_args(argv)
    findings, released, notes = check_with_notes(args.root.resolve(), args.repository)
    for note in notes:
        print(f"NOTE: {note}")
    for finding in findings:
        print(f"FAIL: {finding}")
    status = "FAIL" if findings else "PASS"
    print(f"changelog: {status} ({released} dated release(s); format, VERSION, release state "
          "and frozen released sections, nothing published)")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
