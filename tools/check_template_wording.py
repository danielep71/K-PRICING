#!/usr/bin/env python3
"""Fail when a file describes the template instead of the repository it is in.

Every tracked file is copied into each repository created from the template,
so each must stay true there. A sentence such as "in the template it lists
none" or "the template's own contract" is true where it was written and, once
copied, describes a repository the reader is not in. Reword it so that it is
true in both, move a step that only the template needs to
docs/setup/REPOSITORY_SETUP.md or to the template-only section of AGENTS.md,
or let tools/initialize_repository.py adapt it.

A finding is one of PHRASES, matched case-insensitively. Its words may be
separated by white space, a line break, and the comment and Markdown marks
``#``, ``/``, ``*``, ``>``, ``'`` and backticks, so a phrase wrapped across
the lines of a comment or a paragraph is found. A phrase that is part of a
longer word, such as "in the template-only section", is not.

Scope: every tracked file that is UTF-8 text without a NUL byte, in the
template and in every repository created from it alike, except the paths of
EXCLUDED, each with its reason. The template-only sections of AGENTS.md and
README.md, between their markers, are left out: only the template keeps
them, and the replacement a section holds is read once
tools/initialize_repository.py has put it in their place.

ALLOWED lists the reviewed uses that stay, each as a file, a phrase of
PHRASES and the reason, such as "the template's" meaning the workbook
template. An entry whose phrase does not occur in its file, or that has no
reason, is itself a finding.

Usage: python tools/check_template_wording.py [--root .]
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

from check_repository import TEMPLATE_ONLY_FILES
from gitio import GitError, tracked

ROOT = Path(__file__).resolve().parents[1]
# What may stand between the words of a phrase.
GAP = r"[\s#/*>'`]+"
PHRASES: dict[str, re.Pattern[str]] = {
    "in the template": re.compile(rf"\bin{GAP}the{GAP}template(?![\w'’-])", re.I),
    "the template's": re.compile(rf"\bthe{GAP}template['’]s\b", re.I),
    "template repository": re.compile(rf"\btemplate{GAP}repositor(?:y|ies)\b", re.I),
    "generated repository": re.compile(rf"\bgenerated{GAP}repositor(?:y|ies)\b", re.I),
}
TEMPLATE_ONLY = re.compile(r"<!-- template-only:start -->.*?<!-- template-only:end -->", re.S)
ADVICE = ("reword it so that it is true in a repository created from the template, move it "
          "to docs/setup/REPOSITORY_SETUP.md or the template-only section of AGENTS.md, or "
          "add it to ALLOWED in tools/check_template_wording.py with a reason")


@dataclass(frozen=True)
class Allowed:
    """A reviewed use of a phrase in a file, with the reason it stays."""

    path: str
    # A key of PHRASES.
    phrase: str
    reason: str


# Paths that are not read, each with its reason. A path ending in "/" covers a folder.
EXCLUDED: dict[str, str] = {
    "CHANGELOG.md": "history: its entries record each change as it was made, and "
                    "tools/initialize_repository.py replaces them in a new repository",
    "docs/licenses/": "third-party notices and license texts, kept as their authors wrote "
                      "them",
    "docs/setup/REPOSITORY_SETUP.md": "the handover guide: it describes creating a "
                                      "repository from the template, and a repository keeps "
                                      "it as the record of its setup or retires it",
    "tools/initialize_repository.py": "template-only tooling, deleted once it has run",
    "tests/test_initialize_repository.py": "the test of the template-only initializer, "
                                           "deleted with it",
    "tools/check_template_wording.py": "this check, which names the phrases",
    "tests/test_check_template_wording.py": "the test of this check, whose fixtures hold the "
                                            "phrases",
}

WORKBOOK = "the workbook template, src/workbook/Template.xlsx, not the repository template"
# Reviewed uses that stay. Keep this list short: prefer rewording.
ALLOWED: tuple[Allowed, ...] = (
    Allowed("docs/methodology/WORKBOOK_CONTRACT.md", "in the template", WORKBOOK),
    Allowed("docs/methodology/WORKBOOK_CONTRACT.md", "the template's", WORKBOOK),
    Allowed("tools/check_workbook_protection.py", "the template's", WORKBOOK),
    Allowed("tools/devbuild/build_workbook.py", "the template's", WORKBOOK),
    Allowed("tests/test_devbuild.py", "the template's", WORKBOOK),
)


@dataclass(frozen=True)
class Match:
    path: str
    line: int
    phrase: str


def excluded(path: str, scope: dict[str, str]) -> bool:
    return any(path == name or (name.endswith("/") and path.startswith(name))
               for name in scope)


def text_of(root: Path, path: str) -> str | None:
    """The file's text, without its template-only sections; None if not text."""
    file = root / path
    if not file.is_file():
        return None
    data = file.read_bytes()
    if b"\0" in data:
        return None
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError:
        return None
    if path in TEMPLATE_ONLY_FILES:
        text = TEMPLATE_ONLY.sub(lambda section: re.sub(r"[^\n]", " ", section.group(0)), text)
    return text


def matches(root: Path, scope: dict[str, str]) -> tuple[list[Match], int]:
    """Every phrase found, and the number of files read."""
    found: list[Match] = []
    read = 0
    for path in tracked(root):
        if excluded(path, scope):
            continue
        text = text_of(root, path)
        if text is None:
            continue
        read += 1
        for phrase, pattern in PHRASES.items():
            found += [Match(path, text.count("\n", 0, match.start()) + 1, phrase)
                      for match in pattern.finditer(text)]
    return sorted(found, key=lambda m: (m.path, m.line, m.phrase)), read


def findings(root: Path, scope: dict[str, str] | None = None,
             allowed: Sequence[Allowed] = ALLOWED) -> tuple[list[str], int]:
    """The problems, and the number of files read."""
    problems: list[str] = []
    for entry in allowed:
        if entry.phrase not in PHRASES:
            problems.append(f"allowed use {entry.phrase!r} in {entry.path} is not one of the "
                            f"phrases: {', '.join(map(repr, PHRASES))}")
        if not entry.reason.strip():
            problems.append(f"allowed use {entry.phrase!r} in {entry.path} has no reason")
    try:
        found, read = matches(root, EXCLUDED if scope is None else scope)
    except (OSError, GitError) as error:
        return [f"the tracked files cannot be listed: {error}"], 0
    used: set[Allowed] = set()
    for match in found:
        kept = next((entry for entry in allowed
                     if (entry.path, entry.phrase) == (match.path, match.phrase)), None)
        if kept is not None:
            used.add(kept)
            continue
        problems.append(f"{match.path}:{match.line}: {match.phrase!r} describes the template, "
                        f"not the repository the file is in; {ADVICE}")
    problems += [f"allowed use {entry.phrase!r} in {entry.path} matches nothing; remove it "
                 "from ALLOWED in tools/check_template_wording.py"
                 for entry in allowed if entry not in used and entry.phrase in PHRASES]
    return problems, read


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    problems, read = findings(args.root.resolve())
    for problem in problems:
        print(f"FAIL: {problem}")
    status = "FAIL" if problems else "PASS"
    print(f"template-wording: {status} ({read} file(s) read, {len(PHRASES)} phrase(s), "
          f"{len(ALLOWED)} allowed use(s))")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
