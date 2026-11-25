#!/usr/bin/env python3
"""Check the closing keywords that a pull request's commits carry to the default branch.

GitHub closes an issue when a closing keyword that names it reaches the
default branch, in a pull request's title or body or in the message of any
commit the merge brings there. Commits merged into a release branch are
immutable, and the integration merge carries all of them to the default
branch, so a keyword written in one of them closes its issue at
integration, even when the issue still waits for evidence.

A keyword is close, closes, closed, fix, fixes, fixed, resolve, resolves or
resolved, in any letter case, optionally followed by a colon, then by spaces
and one reference: #N, or OWNER/REPOSITORY#N naming this repository. As in
GitHub's syntax, each reference needs its own keyword: in "fixes #1, #2" only
#1 is closed. A reference to another repository is ignored.

The mode follows the pull request's base and head branches:

- release: into a release branch (release/...). Any keyword, in a commit
  between base and head, the title or the body, is a finding. With the
  default branch as head (a sync pull request), only the title and body are
  read: its commits are already on the default branch.
- integration: into the default branch from a release branch. A keyword
  naming an open issue is a finding; one naming a closed issue or a pull
  request is a note. A commit cannot be changed after it is merged, so a
  commit keyword is accepted when the body names its issue on a line
  ``Reopen after merge: #N`` (several numbers may be listed, separated by
  commas); reopen each such issue after the merge. A keyword in the title or
  body is fixed by editing them instead.
- Otherwise, nothing is checked and the output says why.

Issue states are read through the REST API with GITHUB_TOKEN, or from --input,
a JSON list of REST issue items such as tools/check_issues.py reads.

Exit codes: 0 when there is no finding, 1 when there is one, 2 on a usage
error or when a commit range or an issue state cannot be read.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import urllib.error
import urllib.request
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any

API = "https://api.github.com"
RELEASE_PREFIX = "release/"
REPOSITORY_NAME = re.compile(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+")
# A closing keyword and the one reference it applies to.
CLOSING = re.compile(
    r"\b(?P<keyword>close[sd]?|fix(?:e[sd])?|resolve[sd]?)(?::[ \t]*|[ \t]+)"
    r"(?P<repository>[A-Za-z0-9][A-Za-z0-9-]*/[A-Za-z0-9_.-]+)?#(?P<number>\d+)\b",
    re.IGNORECASE)
ACKNOWLEDGED = re.compile(
    r"^[ \t]*(?:[-*][ \t]+)?Reopen after merge:[ \t]*(?P<numbers>#\d+(?:[ \t]*,[ \t]*#\d+)*)"
    r"[ \t]*$", re.IGNORECASE | re.MULTILINE)
HTML_COMMENT = re.compile(r"<!--.*?-->", re.DOTALL)


class UsageError(Exception):
    """The arguments or the event cannot be used; exit code 2."""


class StateError(Exception):
    """A commit range or an issue state cannot be read; exit code 2."""


@dataclass(frozen=True)
class Reference:
    """One closing keyword, where it was written and the issue it names."""

    where: str  # a short commit SHA, "title" or "body"
    text: str  # the keyword and reference as written
    number: int

    @property
    def in_commit(self) -> bool:
        return self.where not in ("title", "body")


@dataclass(frozen=True)
class Outcome:
    level: str  # "FAIL" or "NOTE"
    reference: Reference
    message: str

    def line(self) -> str:
        where = (f"commit {self.reference.where}" if self.reference.in_commit
                 else f"pull request {self.reference.where}")
        return f"{self.level}: {where}: \"{self.reference.text}\": {self.message}"


@dataclass(frozen=True)
class Request:
    """What to check: the mode, the commit range and the pull request text."""

    mode: str  # "release", "integration" or "none"
    repository: str
    base: str | None
    head: str | None
    title: str
    body: str
    # Why nothing, or only part, is checked; printed with the result.
    reason: str = ""


def references(text: str, where: str, repository: str) -> list[Reference]:
    """Every closing keyword in ``text`` that names an issue of ``repository``."""
    found = []
    for match in CLOSING.finditer(text):
        named = match["repository"]
        if named and named.lower() != repository.lower():
            continue
        found.append(Reference(where, match.group(0), int(match["number"])))
    return found


def acknowledged(body: str) -> set[int]:
    """The issues the body lists on ``Reopen after merge:`` lines, outside HTML comments."""
    numbers: set[int] = set()
    text = HTML_COMMENT.sub("", body.replace("\r\n", "\n").replace("\r", "\n"))
    for match in ACKNOWLEDGED.finditer(text):
        numbers |= {int(number) for number in re.findall(r"\d+", match["numbers"])}
    return numbers


def resolve(root: Path, revision: str) -> str:
    """The full SHA of a commit revision."""
    verify = subprocess.run(["git", "rev-parse", "--verify", "--quiet", f"{revision}^{{commit}}"],
                            cwd=root, capture_output=True, text=True, check=False)
    if verify.returncode:
        raise StateError(f"{revision} is not a commit in {root}")
    return verify.stdout.strip()


def commits(root: Path, base: str, head: str) -> list[tuple[str, str]]:
    """The (SHA, full message) of each commit in base..head, merge commits included."""
    base, head = resolve(root, base), resolve(root, head)
    listed = subprocess.run(["git", "-c", "log.showSignature=false", "log", "-z",
                             "--format=%H%n%B", f"{base}..{head}", "--"], cwd=root,
                            capture_output=True, text=True, check=False)
    if listed.returncode:
        raise StateError(f"git log {base}..{head} failed: {listed.stderr.strip()}")
    entries = []
    for record in listed.stdout.split("\0"):
        if record.strip():
            sha, _, message = record.lstrip("\n").partition("\n")
            entries.append((sha, message))
    return entries


def collect(request: Request, root: Path) -> tuple[list[Reference], int]:
    """The references in the commits, title and body, and the number of commits read."""
    found: list[Reference] = []
    listed = commits(root, request.base, request.head) if request.base and request.head else []
    for sha, message in listed:
        found += references(message, sha[:7], request.repository)
    found += references(request.title, "title", request.repository)
    found += references(request.body, "body", request.repository)
    return found, len(listed)


StateReader = Callable[[int], tuple[str, bool]]


def judge(request: Request, found: Sequence[Reference],
          read_state: StateReader) -> list[Outcome]:
    """The finding or note for each reference; reads issue states in integration mode."""
    if request.mode == "release":
        return [Outcome("FAIL", reference,
                        f"a closing keyword for #{reference.number} on a release branch; "
                        "write Refs #N instead" + (" and edit the pull request" if not
                                                   reference.in_commit else ""))
                for reference in found]
    accepted = acknowledged(request.body)
    outcomes = []
    for reference in found:
        state, is_pull = read_state(reference.number)
        if is_pull:
            outcomes.append(Outcome("NOTE", reference,
                                    f"#{reference.number} is a pull request; nothing is closed"))
        elif state != "open":
            outcomes.append(Outcome("NOTE", reference,
                                    f"issue #{reference.number} is already {state}"))
        elif not reference.in_commit:
            outcomes.append(Outcome("FAIL", reference,
                                    f"would close open issue #{reference.number}; write "
                                    "Refs #N in the pull request instead"))
        elif reference.number in accepted:
            outcomes.append(Outcome("NOTE", reference,
                                    f"will close open issue #{reference.number}; the body "
                                    "says to reopen it after the merge"))
        else:
            outcomes.append(Outcome("FAIL", reference,
                                    f"would close open issue #{reference.number} at the "
                                    f"merge; add the line \"Reopen after merge: "
                                    f"#{reference.number}\" to the pull request body and "
                                    "reopen the issue after the merge"))
    return outcomes


# ---------------------------------------------------------------------------
# Issue states.

def fetch(url: str, token: str) -> Any:
    request = urllib.request.Request(url, headers={
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "X-GitHub-Api-Version": "2022-11-28",
    })
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.load(response)


def state_of(item: Mapping[str, Any]) -> tuple[str, bool]:
    return str(item["state"]).lower(), "pull_request" in item


def live_states(repository: str, token: str) -> StateReader:
    cache: dict[int, tuple[str, bool]] = {}

    def read(number: int) -> tuple[str, bool]:
        if number not in cache:
            try:
                cache[number] = state_of(fetch(f"{API}/repos/{repository}/issues/{number}",
                                               token))
            except (OSError, ValueError, KeyError, urllib.error.URLError) as error:
                raise StateError(f"cannot read the state of #{number}: {error}") from None
        return cache[number]
    return read


def file_states(path: Path) -> StateReader:
    try:
        items = json.loads(path.read_text(encoding="utf-8"))
        known = {int(item["number"]): state_of(item) for item in items}
    except (OSError, ValueError, KeyError, TypeError) as error:
        raise StateError(f"cannot read {path}: {error}") from None

    def read(number: int) -> tuple[str, bool]:
        if number not in known:
            raise StateError(f"cannot read the state of #{number}: it is not in {path}")
        return known[number]
    return read


def token_states(repository: str) -> StateReader:
    token = os.environ.get("GITHUB_TOKEN", "")

    def read(number: int) -> tuple[str, bool]:
        if not token:
            raise StateError(f"cannot read the state of #{number}: set GITHUB_TOKEN, or pass "
                             "--input")
        return reader(number)
    reader = live_states(repository, token)
    return read


# ---------------------------------------------------------------------------
# The request, from a pull_request event or from arguments.

def mode_for(base: str, head: str, default_branch: str) -> tuple[str, str]:
    """The mode for a pull request from head into base, and why when it checks nothing."""
    if base.startswith(RELEASE_PREFIX):
        return "release", ""
    if base == default_branch and head.startswith(RELEASE_PREFIX):
        return "integration", ""
    return "none", (f"not checked: a pull request from {head} into {base} is neither into a "
                    f"release branch nor an integration of one into {default_branch}")


def from_event(path: Path) -> Request:
    try:
        event = json.loads(path.read_text(encoding="utf-8"))
        pull = event["pull_request"]
        repository = str(event["repository"]["full_name"])
        default_branch = str(event["repository"]["default_branch"])
        base_ref, head_ref = str(pull["base"]["ref"]), str(pull["head"]["ref"])
        base, head = str(pull["base"]["sha"]), str(pull["head"]["sha"])
        title, body = str(pull.get("title") or ""), str(pull.get("body") or "")
        # A head in another repository, such as a fork, or one whose repository
        # was deleted, is never this repository's default or release branch.
        head_repository = (pull["head"].get("repo") or {}).get("full_name")
    except (OSError, ValueError, KeyError, TypeError, AttributeError) as error:
        raise UsageError(f"{path} is not a pull_request event: {error!r}") from None
    own_head = str(head_repository or "").lower() == repository.lower()
    if not own_head and not base_ref.startswith(RELEASE_PREFIX):
        source = head_repository or "a deleted repository"
        return Request("none", repository, base, head, title, body,
                       f"not checked: a pull request from {source}:{head_ref} into "
                       f"{base_ref} is not an integration of this repository's release branch")
    mode, reason = mode_for(base_ref, head_ref, default_branch)
    if mode == "release" and own_head and head_ref == default_branch:
        return Request(mode, repository, None, None, title, body,
                       f"commits not checked: the head {head_ref} is the default branch, "
                       "whose commits are already there")
    return Request(mode, repository, base, head, title, body, reason)


def parse_args(argv: Sequence[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1],
                        help="the Git repository (default: this one)")
    parser.add_argument("--event", type=Path,
                        help="a pull_request event, such as $GITHUB_EVENT_PATH; it gives "
                             "everything below")
    parser.add_argument("--repository", metavar="OWNER/REPOSITORY")
    parser.add_argument("--mode", choices=("release", "integration"))
    parser.add_argument("--base", help="base revision of the range base..head")
    parser.add_argument("--head", help="head revision of the range base..head")
    parser.add_argument("--title", default="", help="the pull request title")
    parser.add_argument("--body", default="", help="the pull request body")
    parser.add_argument("--input", type=Path,
                        help="issue states as a JSON list of REST issue items, instead of the "
                             "API")
    parser.add_argument("--summary", type=Path, help="also append the report to this file")
    args = parser.parse_args(argv)
    manual = ("repository", "mode", "base", "head")
    if args.event:
        if any(getattr(args, name) for name in (*manual, "title", "body")):
            parser.error("--event gives the repository, mode, range, title and body; do not "
                         "pass them too")
    else:
        missing = [f"--{name}" for name in manual if not getattr(args, name)]
        if missing:
            parser.error(f"without --event, pass {', '.join(missing)}")
        if not REPOSITORY_NAME.fullmatch(args.repository):
            parser.error("--repository must have the form OWNER/REPOSITORY")
    return args


def report(request: Request, outcomes: Sequence[Outcome], checked: int) -> str:
    failures = sum(outcome.level == "FAIL" for outcome in outcomes)
    lines = [f"## Closing keywords: {'FAIL' if failures else 'PASS'}", "",
             f"- Mode: {request.mode}"]
    if request.mode != "none" and request.base and request.head:
        lines.append(f"- Commits read: {checked} ({request.base}..{request.head})")
    if request.reason:
        lines.append(f"- {request.reason[0].upper()}{request.reason[1:]}")
    if request.mode != "none":
        lines.append(f"- Closing keywords found: {len(outcomes)}, findings: {failures}")
    lines += ["", *[f"- {outcome.line()}" for outcome in outcomes]] if outcomes else []
    return "\n".join(lines) + "\n"


def run(request: Request, root: Path, read_state: StateReader) -> tuple[str, int]:
    if request.mode == "none":
        return report(request, [], 0), 0
    if request.base and request.head:
        request = replace(request, base=resolve(root, request.base)[:7],
                          head=resolve(root, request.head)[:7])
    found, checked = collect(request, root)
    outcomes = judge(request, found, read_state)
    failed = any(outcome.level == "FAIL" for outcome in outcomes)
    return report(request, outcomes, checked), 1 if failed else 0


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        request = from_event(args.event) if args.event else Request(
            args.mode, args.repository, args.base, args.head, args.title, args.body)
        read_state = file_states(args.input) if args.input else token_states(request.repository)
        text, code = run(request, args.root.resolve(), read_state)
    except (UsageError, StateError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    print(text, end="")
    if args.summary:
        with args.summary.open("a", encoding="utf-8") as handle:
            handle.write(text)
    return code


if __name__ == "__main__":
    sys.exit(main())
