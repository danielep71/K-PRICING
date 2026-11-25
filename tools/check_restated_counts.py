#!/usr/bin/env python3
"""Fail when a guide restates a count that has a single source.

Some counts live in one place: a policy file, a contract or the code of a
tool. tools/render_docs.py generates the ones a guide shows into blocks
between ``BEGIN generated`` and ``END generated`` markers, and Repository
integrity fails while a block is stale. A count copied into prose outside a
block has no such check, and goes stale silently when its source changes.

Each source is declared once in SOURCES, with a reader that returns its
current value and the noun phrases that name it:

- the cases and the assertions of the example suite, read from
  ``suite.expected_cases`` and ``suite.expected_assertions`` of
  .github/excel-evidence-policy.json;
- the modes of the example suite, and its failure entry points, the modes
  that must not pass, read from ``modes`` of the same policy;
- the Repository integrity checks that tools/check_repository.py lists, as
  the generated table of CONTRIBUTING.md shows them, and the static checks
  of exported VBA source among them, those that run a tools/check_vba_*.py;
- the datasets of docs/methodology/data-contract.json and the finding codes
  of tools/check_data_contract.py.

The issue label count is not a source here: tests/test_repository_consistency.py
already holds every label count the guides state equal to .github/labels.json.

A finding is a declared noun phrase preceded by a count, in digits (``14``,
``1,024``) or in words (``three``, ``thirty-four``, up to ninety-nine),
whatever its value: a stale copy is reported as well as a current one. The
number and the phrase may be separated by white space, a line break, and the
Markdown marks of emphasis and code spans, and nothing else, so
"34 ms", "issue #34", "v3.4", "step 6 checks" or a commit hash that contains
34 are not findings. A number that is part of a longer token, a version, an
issue reference or a path (preceded by a letter, digit, ``.``, ``#``, ``/``,
``$`` or ``-``) is not a count.

Scope: every tracked Markdown file except CHANGELOG.md, which records counts
as they were released, with the text of each generated block left out.

ALLOWED lists the restatements that are reviewed and kept, each as a file, the
phrase as the finding prints it, and the reason. An entry that matches nothing,
or that has no reason, is itself a finding.

Usage: python tools/check_restated_counts.py [--root .]
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

import check_data_contract  # noqa: E402
import check_repository  # noqa: E402
import render_docs  # noqa: E402
from gitio import GitError, tracked_or_scanned  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
POLICY = ".github/excel-evidence-policy.json"
EXCLUDED = frozenset({"CHANGELOG.md"})
ADVICE = "link the source or add a renderer to tools/render_docs.py"

UNITS = ("one", "two", "three", "four", "five", "six", "seven", "eight", "nine")
TEENS = ("ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen",
         "seventeen", "eighteen", "nineteen")
TENS = ("twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety")
WORD_VALUES: dict[str, int] = {
    **{word: index + 1 for index, word in enumerate(UNITS)},
    **{word: index + 10 for index, word in enumerate(TEENS)},
    **{word: (index + 2) * 10 for index, word in enumerate(TENS)},
}
# Marks that may stand between a count and its phrase, or between the words of
# a phrase: white space, line breaks, Markdown emphasis and code spans.
GAP = r"[\s*_`]+"
WORDS = (rf"(?:{'|'.join(TENS)})(?:[-\s](?:{'|'.join(UNITS)}))?"
         rf"|{'|'.join(TEENS)}|{'|'.join(UNITS)}")
# A count in digits, grouped or not, or in words; not part of a longer token,
# a decimal, a version, an issue reference or a path.
NUMBER = (rf"(?<![A-Za-z0-9.#/$,-])(?P<number>\d{{1,3}}(?:,\d{{3}})+(?![\d,])|\d+(?![\d.,]\d)"
          rf"|(?:{WORDS})(?![A-Za-z-]))")


@dataclass(frozen=True)
class Source:
    """A count with a single source, and the noun phrases that name it."""

    # What is counted, as a finding names it.
    name: str
    # Where the count lives.
    where: str
    # Plural noun phrases that follow a count of it, as in "14 cases".
    phrases: tuple[str, ...]
    read: Callable[[Path], int]


@dataclass(frozen=True)
class Allowed:
    """A reviewed restatement that stays, with the reason it stays."""

    path: str
    # The phrase as a finding prints it, such as "three modes".
    phrase: str
    reason: str


def policy(root: Path) -> dict[str, Any]:
    data = json.loads((root / POLICY).read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"{POLICY}: not a JSON object")
    return data


def count(value: object, what: str) -> int:
    if not isinstance(value, int) or isinstance(value, bool) or value < 0:
        raise ValueError(f"{what} is not a count: {value!r}")
    return value


def suite_count(key: str) -> Callable[[Path], int]:
    def read(root: Path) -> int:
        suite = policy(root).get("suite")
        if not isinstance(suite, dict):
            raise ValueError(f"{POLICY}: suite is not an object")
        return count(suite.get(key), f"{POLICY}: suite.{key}")
    return read


def modes(root: Path) -> list[Any]:
    listed = policy(root).get("modes")
    if not isinstance(listed, list):
        raise ValueError(f"{POLICY}: modes is not a list")
    return listed


def failure_modes(root: Path) -> int:
    return sum(1 for mode in modes(root) if not (isinstance(mode, dict) and mode.get("must_pass")))


def integrity_checks(root: Path) -> list[check_repository.Check]:
    # As tools/render_docs.py lists them in CONTRIBUTING.md: without --ci.
    context = check_repository.Context(root, repository="OWNER/REPOSITORY")
    return check_repository.checks(context)


VBA_CHECK = re.compile(r"(?:^|/)tools/check_vba_\w+\.py$")


def vba_checks(root: Path) -> int:
    return sum(1 for check in integrity_checks(root)
               if check.command and any(VBA_CHECK.search(word) for word in check.command))


def datasets(root: Path) -> int:
    loaded, problems = check_data_contract.load(root)
    if problems:
        raise ValueError(problems[0])
    return len(loaded)


SOURCES: tuple[Source, ...] = (
    Source("cases of the example suite", f"suite.expected_cases in {POLICY}",
           ("cases", "expected cases"), suite_count("expected_cases")),
    Source("assertions of the example suite", f"suite.expected_assertions in {POLICY}",
           ("assertions", "expected assertions"), suite_count("expected_assertions")),
    Source("modes of the example suite", f"modes in {POLICY}", ("modes",),
           lambda root: len(modes(root))),
    Source("failure entry points of the example suite",
           f"the modes in {POLICY} that must not pass",
           ("entry points", "failure entry points", "failure modes"), failure_modes),
    Source("Repository integrity checks", "the checks tools/check_repository.py lists",
           ("integrity checks", "Repository integrity checks", "repository checks"),
           lambda root: len(integrity_checks(root))),
    Source("static checks of exported VBA source",
           "the checks tools/check_repository.py runs from tools/check_vba_*.py",
           ("static checks", "VBA checks", "static VBA checks"), vba_checks),
    Source("datasets of the data contract", check_data_contract.CONTRACT.as_posix(),
           ("datasets", "data sets"), datasets),
    Source("finding codes of the data-contract validator",
           "CODES in tools/check_data_contract.py",
           ("finding codes", "validator codes"), lambda root: len(check_data_contract.CODES)),
)

# Reviewed restatements that stay. Keep this list short: prefer a link to the
# source or a generated block.
ALLOWED: tuple[Allowed, ...] = ()


def number_value(text: str) -> int:
    """The value of a count as written, in digits or in words."""
    if text[0].isdigit():
        return int(text.replace(",", ""))
    return sum(WORD_VALUES[word] for word in re.split(r"[-\s]+", text.lower()))


def phrase_pattern(phrases: Iterable[str]) -> re.Pattern[str]:
    alternatives = sorted({GAP.join(re.escape(word) for word in phrase.split())
                           for phrase in phrases}, key=len, reverse=True)
    return re.compile(rf"{NUMBER}{GAP}(?:{'|'.join(alternatives)})(?![\w-])", re.I)


def shown(text: str) -> str:
    """A matched phrase as a finding prints it."""
    return " ".join(re.sub(GAP, " ", text).split()).lower()


def without_generated_blocks(text: str, path: str) -> str:
    """The text with each generated block blanked, keeping its line breaks."""
    spans, _ = render_docs.blocks_in(text, path)
    for begin, end in spans.values():
        text = text[:begin] + re.sub(r"[^\n]", " ", text[begin:end]) + text[end:]
    return text


@dataclass(frozen=True)
class Match:
    path: str
    line: int
    phrase: str
    source: Source
    value: int


def matches(root: Path, sources: Sequence[Source]) -> list[Match]:
    found: list[Match] = []
    patterns = [(source, phrase_pattern(source.phrases)) for source in sources]
    for path in tracked_or_scanned(root, "*.md"):
        if Path(path).name in EXCLUDED or not (root / path).is_file():
            continue
        text = without_generated_blocks((root / path).read_text(encoding="utf-8"), path)
        for source, pattern in patterns:
            for match in pattern.finditer(text):
                found.append(Match(path, text.count("\n", 0, match.start()) + 1,
                                   shown(match.group(0)), source,
                                   number_value(match.group("number"))))
    return sorted(found, key=lambda m: (m.path, m.line, m.phrase))


def findings(root: Path, sources: Sequence[Source] = SOURCES,
             allowed: Sequence[Allowed] = ALLOWED) -> list[str]:
    problems: list[str] = []
    values: dict[str, int] = {}
    for source in sources:
        try:
            values[source.name] = source.read(root)
        except (OSError, ValueError, KeyError, GitError) as error:
            problems.append(f"source {source.name} ({source.where}) cannot be read: {error}")
    for entry in allowed:
        if not entry.reason.strip():
            problems.append(f"allowed restatement {entry.phrase!r} in {entry.path} has no "
                            "reason")
    used: set[Allowed] = set()
    for match in matches(root, sources):
        kept = next((entry for entry in allowed if entry.path == match.path
                     and shown(entry.phrase) == match.phrase), None)
        if kept is not None:
            used.add(kept)
            continue
        current = values.get(match.source.name)
        held = "" if current is None else f", which holds {current}"
        stale = "" if current is None or current == match.value else " (a stale copy)"
        problems.append(f"{match.path}:{match.line}: {match.phrase!r} restates the "
                        f"{match.source.name}{stale}; its source is {match.source.where}"
                        f"{held}; {ADVICE}")
    problems += [f"allowed restatement {entry.phrase!r} in {entry.path} matches nothing; "
                 "remove it from ALLOWED in tools/check_restated_counts.py"
                 for entry in allowed if entry not in used]
    return problems


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    problems = findings(args.root.resolve())
    for problem in problems:
        print(f"FAIL: {problem}")
    status = "FAIL" if problems else "PASS"
    print(f"restated-counts: {status} ({len(SOURCES)} source(s), {len(ALLOWED)} allowed "
          "restatement(s))")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
