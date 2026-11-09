#!/usr/bin/env python3
"""Render the v0.0.4 milestone register from live GitHub issues, or check it for drift.

docs/MILESTONE_REGISTER.md holds every issue of milestone v0.0.4 between two
marker lines: number, title, state, assignees, labels, milestone, URL and the
complete body, in ascending issue-number order. --write fetches the live
milestone and rewrites only that region; --check fetches it and compares:

  CURRENT     (exit 0) the register matches the live milestone;
  DRIFT       (exit 1) it does not; the first mismatched issue and field and
              the regeneration command are printed;
  UNVERIFIED  (exit 2) the live state could not be read: authentication,
              rate limit, network, API error or a malformed response. Never
              reported as DRIFT.

--snapshot reads a saved issue list instead of GitHub, for the frozen tests.
The result depends on live, mutable GitHub state, so the deterministic static
gate never runs this tool; .github/workflows/milestone-register.yml does. It
makes no claim about VBA, Excel or the source tree.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.request
from collections.abc import Callable, Sequence
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
DOCUMENT = ROOT / "docs" / "MILESTONE_REGISTER.md"
DEFAULT_REPOSITORY = "danielep71/K-PRICING"
MILESTONE_PREFIX = "v0.0.4"
API = "https://api.github.com"
BEGIN = "<!-- BEGIN GENERATED MILESTONE REGISTER: tools/milestone_register.py --write -->"
END = "<!-- END GENERATED MILESTONE REGISTER -->"
REGENERATE = "python3 tools/milestone_register.py --write"
CURRENT, DRIFT, UNVERIFIED = 0, 1, 2
RESULTS = {CURRENT: "CURRENT", DRIFT: "DRIFT", UNVERIFIED: "UNVERIFIED"}
PER_PAGE = 100
MAX_PAGES = 20

Fetch = Callable[[str], Any]


class Unverified(Exception):
    """The live milestone could not be read; the register is neither current nor drifted."""


class RegionError(Exception):
    """The document has no single, ordered generated region."""


def github_fetch(token: str | None) -> Fetch:
    """GET a GitHub API path and parse its JSON; every failure is Unverified."""

    def fetch(path: str) -> Any:
        request = urllib.request.Request(API + path, headers={
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "kpr-milestone-register",
            **({"Authorization": f"Bearer {token}"} if token else {}),
        })
        try:
            with urllib.request.urlopen(request, timeout=30) as response:  # noqa: S310 -- fixed https API host
                payload = response.read()
        except urllib.error.HTTPError as error:
            remaining = error.headers.get("x-ratelimit-remaining") if error.headers else None
            hint = " (rate limit exhausted)" if remaining == "0" or error.code == 429 else ""
            raise Unverified(f"GitHub returned HTTP {error.code}{hint} for {path}") from error
        except (urllib.error.URLError, TimeoutError, OSError) as error:
            raise Unverified(f"GitHub could not be reached for {path}: {error}") from error
        try:
            return json.loads(payload.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as error:
            raise Unverified(f"GitHub returned malformed JSON for {path}") from error

    return fetch


def snapshot_fetch(snapshot: dict[str, Any]) -> Fetch:
    """Serve a saved {"milestones": [...], "issues": [...]} snapshot as the API would."""

    def fetch(path: str) -> Any:
        page = int(path.rsplit("page=", 1)[1]) if "page=" in path else 1
        key = "milestones" if path.startswith("/repos/") and "/milestones" in path else "issues"
        items = snapshot.get(key)
        if not isinstance(items, list):
            raise Unverified(f"the snapshot has no {key} list")
        return items if page == 1 else []

    return fetch


def _pages(fetch: Fetch, path: str) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    for page in range(1, MAX_PAGES + 1):
        batch = fetch(f"{path}&per_page={PER_PAGE}&page={page}")
        if not isinstance(batch, list) or not all(isinstance(item, dict) for item in batch):
            raise Unverified(f"GitHub returned a malformed list for {path}")
        items.extend(batch)
        if len(batch) < PER_PAGE:
            return items
    raise Unverified(f"{path} has more than {MAX_PAGES} pages")


def _text(item: dict[str, Any], key: str, number: object) -> str:
    value = item.get(key)
    if not isinstance(value, str):
        raise Unverified(f"issue {number}: field {key} is missing or not text")
    return value


def _names(item: dict[str, Any], key: str, field: str, number: object) -> list[str]:
    values = item.get(key)
    if not isinstance(values, list) or not all(
            isinstance(value, dict) and isinstance(value.get(field), str) for value in values):
        raise Unverified(f"issue {number}: field {key} is malformed")
    return sorted((str(value[field]) for value in values), key=str.casefold)


def normal_issue(item: dict[str, Any]) -> dict[str, Any]:
    """One register record from one API issue; a malformed issue is Unverified."""
    number = item.get("number")
    if not isinstance(number, int) or isinstance(number, bool):
        raise Unverified("an issue has no integer number")
    milestone = item.get("milestone")
    if not isinstance(milestone, dict) or not isinstance(milestone.get("title"), str):
        raise Unverified(f"issue {number}: field milestone is malformed")
    body = item.get("body")
    if body is not None and not isinstance(body, str):
        raise Unverified(f"issue {number}: field body is not text")
    # Line ends are normalized and trailing blanks dropped, so the register
    # stays a clean tracked text file; the wording is kept as written.
    lines = [line.rstrip() for line in (body or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    while lines and not lines[-1]:
        lines.pop()
    return {
        "number": number,
        "title": _text(item, "title", number),
        "state": _text(item, "state", number),
        "assignees": _names(item, "assignees", "login", number),
        "labels": _names(item, "labels", "name", number),
        "milestone": str(milestone["title"]),
        "url": _text(item, "html_url", number),
        "body": "\n".join(lines),
    }


def fetch_register(fetch: Fetch, repository: str, prefix: str) -> tuple[str, list[dict[str, Any]]]:
    """The milestone title and its issues (pull requests excluded), by ascending number."""
    milestones = [item for item in _pages(fetch, f"/repos/{repository}/milestones?state=all")
                  if isinstance(item.get("title"), str) and str(item["title"]).startswith(prefix)]
    if len(milestones) != 1 or not isinstance(milestones[0].get("number"), int):
        raise Unverified(f"expected one milestone titled {prefix}..., found {len(milestones)}")
    milestone = milestones[0]
    issues = _pages(fetch, f"/repos/{repository}/issues?milestone={milestone['number']}&state=all")
    records = [normal_issue(item) for item in issues if "pull_request" not in item]
    return str(milestone["title"]), sorted(records, key=lambda record: int(record["number"]))


def _cell(value: str) -> str:
    return value.replace("\\", "\\\\").replace("|", "\\|") if value else "none"


def _fence(body: str) -> str:
    # Markdown lets a closing fence be indented, so a run counts after any
    # leading blanks; counting deeper indents too only lengthens the fence.
    runs = (line.lstrip(" \t") for line in body.split("\n"))
    longest = max((len(run) - len(run.lstrip("`")) for run in runs), default=0)
    return "`" * max(4, longest + 1)


def render(title: str, records: Sequence[dict[str, Any]]) -> str:
    """The generated region, without its marker lines; byte-identical for the same input."""
    closed = sum(record["state"] == "closed" for record in records)
    lines = [f"Milestone **{_cell(title)}**: {len(records)} issues, {len(records) - closed} open, {closed} closed.", ""]
    for record in records:
        fence = _fence(record["body"])
        lines += [
            f"### #{record['number']} {_cell(record['title'])}", "",
            "| Field | Value |", "| --- | --- |",
            f"| State | {record['state']} |",
            f"| Assignees | {_cell(', '.join(record['assignees']))} |",
            f"| Labels | {_cell(', '.join(record['labels']))} |",
            f"| Milestone | {_cell(record['milestone'])} |",
            f"| URL | {record['url']} |", "",
            f"{fence}text", *([record["body"]] if record["body"] else []), fence, "",
        ]
    return "\n".join(lines)


def split_region(document: str) -> tuple[str, str, str]:
    """(before, region, after); the markers must each appear once, as whole lines, in order."""
    lines = document.split("\n")
    if lines.count(BEGIN) != 1 or lines.count(END) != 1 or lines.index(BEGIN) > lines.index(END):
        raise RegionError("the register needs exactly one BEGIN and one END marker line, in that order")
    start, stop = lines.index(BEGIN), lines.index(END)
    return "\n".join(lines[: start + 1]), "\n".join(lines[start + 1: stop]), "\n".join(lines[stop:])


def marker_collision(records: Sequence[dict[str, Any]]) -> None:
    """A body line equal to a marker would split the region on the next read; refuse it before writing."""
    for record in records:
        if {BEGIN, END} & set(record["body"].split("\n")):
            raise RegionError(f"issue #{record['number']}: field body has a line equal to a generated-region "
                              "marker; the register cannot hold it, so nothing was written")


def replace_region(document: str, region: str) -> str:
    before, _, after = split_region(document)
    return f"{before}\n{region}\n{after}"


def first_mismatch(committed: str, live: str) -> str:
    """The first issue and field where the committed region differs from the live one."""
    old, new = committed.split("\n"), live.split("\n")
    # Issue lines first: a changed issue also changes the summary counts above it.
    index = next((i for i, (a, b) in enumerate(zip(old, new, strict=False)) if i and a != b),
                 min(len(old), len(new)))
    if index == min(len(old), len(new)) and len(old) == len(new):
        return "the milestone summary: field summary differs"
    issue, field, closing = "the milestone summary", "summary", ""
    for line in new[: index + 1]:
        if closing:
            # Inside a body every line is body text, whatever it looks like;
            # only the fence that opened it, longer than any fence inside, ends it.
            field = "body"
            if line == closing:
                closing = ""
        elif line.startswith("### #"):
            issue, field = "issue " + line[4:].split(" ", 1)[0], "title"
        elif line.startswith("| ") and not line.startswith(("| Field", "| ---")):
            field = line[2:].split(" |", 1)[0].casefold()
        elif re.fullmatch(r"`{4,}text", line):
            field, closing = "body", line[: -len("text")]
    if index >= len(new) or index >= len(old):
        return f"{issue}: the register has {'more' if len(old) > len(new) else 'fewer'} entries than the milestone"
    return f"{issue}: field {field} differs"


def run(mode: str, document: Path, fetch: Fetch, repository: str, prefix: str) -> tuple[int, str]:
    try:
        title, records = fetch_register(fetch, repository, prefix)
        text = document.read_text(encoding="utf-8")
        marker_collision(records)
        live = render(title, records)
        if mode == "write":
            document.write_text(replace_region(text, live), encoding="utf-8")
            return CURRENT, f"Wrote {len(records)} issues of {title} into {document.name}"
        committed = split_region(text)[1]
    except (Unverified, RegionError, OSError) as error:
        return UNVERIFIED, str(error)
    if committed == live:
        return CURRENT, f"{document.name} matches the {len(records)} issues of {title}"
    return DRIFT, f"{first_mismatch(committed, live)}; run `{REGENERATE}` and commit the result"


def summary(code: int, detail: str) -> str:
    meaning = {
        CURRENT: "The register matches the live milestone.",
        DRIFT: "The register differs from the live milestone.",
        UNVERIFIED: "The live milestone could not be read; this is not a drift result.",
    }[code]
    return (f"# Milestone register\n\nResult: **{RESULTS[code]}**. {meaning}\n\n- {detail}\n\n"
            "This compares GitHub issue state with docs/MILESTONE_REGISTER.md; it makes no claim "
            "about VBA, Excel or the source tree.\n")


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--check", action="store_true", help="compare with the live milestone (default)")
    mode.add_argument("--write", action="store_true", help="rewrite the generated region from the live milestone")
    parser.add_argument("--document", type=Path, default=DOCUMENT)
    parser.add_argument("--repository", default=os.environ.get("GITHUB_REPOSITORY") or DEFAULT_REPOSITORY)
    parser.add_argument("--milestone", default=MILESTONE_PREFIX, help="milestone title prefix")
    parser.add_argument("--snapshot", type=Path, help="saved issue list instead of GitHub")
    parser.add_argument("--summary", type=Path, help="append a Markdown summary here")
    args = parser.parse_args(argv)
    if args.snapshot is not None:
        try:
            data = json.loads(args.snapshot.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as error:
            print(f"UNVERIFIED: cannot read the snapshot: {error}", file=sys.stderr)
            return UNVERIFIED
        fetch = snapshot_fetch(data if isinstance(data, dict) else {})
    else:
        fetch = github_fetch(os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN"))
    code, detail = run("write" if args.write else "check", args.document, fetch, args.repository, args.milestone)
    if args.write and code == CURRENT:
        print(detail)
    else:
        print(f"{RESULTS[code]}: {detail}", file=sys.stdout if code == CURRENT else sys.stderr)
    if args.summary is not None:
        with args.summary.open("a", encoding="utf-8") as handle:
            handle.write(summary(code, detail))
    return code


if __name__ == "__main__":
    raise SystemExit(main())
