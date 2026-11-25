#!/usr/bin/env python3
"""Adapt the files of a repository created from the template, once, in one diff.

It makes the edits that the *Template cleanup* row of
docs/setup/REPOSITORY_SETUP.md lists, those that **Repository integrity**
requires with --template false among them:

- AGENTS.md and README.md: replace each template-only section with the
  replacement it holds, or delete it, with its markers (format below);
- .github/ISSUE_TEMPLATE/config.yml: point the contact links at the new
  repository;
- README.md: point the traffic badges at the new repository, or remove them
  with --remove-traffic-badges;
- .github/dependabot.yml: set target-branch to --release-branch, or remove it;
- CHANGELOG.md: keep the introduction and the policy sections, and replace the
  entries and link references with an empty [Unreleased] section of the new
  repository; with them go VERSION and the release rows of the *Current
  status* of RELEASING.md, which describe the template's releases;
- docs/VBA_HOUSE_STYLE.md: replace <REGIME> in the naming table with
  --regime, and the paragraph above the table, which asks for that, with
  one that names the prefix;
- INSTALLATION.md and docs/methodology/VALIDATION_PLAN.md: with the
  changelog's entries go the template's evidence, which cites its own commits
  and issues: the rows of the Tested environment table are removed, and the
  Issue column of the Host behaviour table becomes a dash.

Two options make more edits:

- --remove-traffic removes the optional traffic collection: its workflow,
  guide, tool and test, the README's badges (it implies
  --remove-traffic-badges, which removes only the badges), the ruleset of its
  branch in docs/setup/repository-settings.json, its files in the license
  notices, the readme-traffic-badges entry of tools/verification_depth.json,
  which tools/check_repository.py no longer lists without the collection, and
  the rows and sentences of the guides that describe it. A ruleset already
  created on GitHub stays: tools/setup_repository.py reports a ruleset the
  configuration does not hold as unmanaged and never deletes it, so the
  preview says to delete it by hand. A mention it cannot place is listed as a
  warning, to edit by hand; the Python code that tests the collection only
  while it exists stays, as do the setup guide's lines that describe this
  initializer and CHANGELOG.md.
- --assignee LOGIN makes LOGIN the assignee of the issue forms and of
  .github/dependabot.yml, and names it where docs/repo_governance/ISSUE_LABELS.md
  names the current one.

A template-only section, in AGENTS.md or README.md, reads:

    <!-- template-only:start -->
    What only the template keeps.
    <!-- template-only:replacement
    What takes its place in a new repository, hidden in the template.
    -->
    <!-- template-only:end -->

The replacement is optional: without it, the section is deleted. When the
start marker is alone on its line, the section is whole lines, and one blank
line is kept around what replaces it. Otherwise it is text within a line,
such as a table cell or a paragraph, and the replacement is put in exactly as
written between "template-only:replacement" and "-->".

Without --apply it only lists the changes. With --apply it writes them,
renders the generated blocks afresh as tools/render_docs.py --write does,
and, last, once all that has succeeded, deletes itself and its test. It
refuses a working tree that differs from the commit, so that the result is
one reviewable diff. It changes nothing on GitHub and commits nothing. A file
already adapted is left as it is, so a second run from elsewhere, with
--root, changes nothing.
"""

from __future__ import annotations

import argparse
import contextlib
import io
import json
import re
import sys
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
# The preview writes nothing, not even the bytecode of the tools it imports.
sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parent))

import gitio  # noqa: E402
import render_docs  # noqa: E402
from check_changelog import NO_RELEASE, VERSION_HEADING, foreign_references  # noqa: E402
from check_issues import FORM_PRIORITIES, flow_list  # noqa: E402
from check_repository import (  # noqa: E402
    BADGE,
    CONTACT_URL,
    TARGET_BRANCH,
    TEMPLATE_ONLY_FILES,
    TRAFFIC_WORKFLOW,
    MissingContext,
    origin_has_branch,
)
from gitio import GitError  # noqa: E402
from vba_lexer import FACADE_ROW, HOUSE_STYLE, RESERVED_PREFIXES  # noqa: E402

REPOSITORY = re.compile(r"[\w.-]+/[\w.-]+")
REGIME = re.compile(r"[A-Z][A-Z0-9]*")
BRANCH = re.compile(r"[\w.-]+(?:/[\w.-]+)*")
# A GitHub login: letters, digits and single hyphens, neither first nor last.
LOGIN = re.compile(r"[A-Za-z0-9](?:-?[A-Za-z0-9])*")
LOGIN_LENGTH = 39
SECTION_START = "<!-- template-only:start -->"
SECTION_END = "<!-- template-only:end -->"
REPLACEMENT = "<!-- template-only:replacement"
GITHUB_URL = re.compile(r"https://github\.com/[\w.-]+/[\w.-]+/")
# An update's directory, or its directories with the list items below the key.
DIRECTORY = re.compile(r"^([ \t]*)(?:directory:[^\n]*|directories:[ \t]*(?:\n\1[ \t]*-[^\n]*)+)$",
                       re.M)
INSTALLATION = "INSTALLATION.md"
VALIDATION_PLAN = "docs/methodology/VALIDATION_PLAN.md"
TESTED_HEADER = "| Tested environment |"
HOST_HEADER = "| ID | Behaviour |"
# The house style's request to replace <REGIME>, up to the blank line after it.
REGIME_REQUEST = re.compile(r"^Replace `<REGIME>` with the project's short name\b"
                            r".*?(?=\n[ \t]*\n)", re.M | re.S)
LAST_CELL = re.compile(r"\|[^|]*\|[ \t]*$")
NO_EVIDENCE = "—"
STATUS_ROWS = {
    "Released versions": "None",
    "`VERSION`": "Does not exist yet; it is created when the first release is prepared",
}
ISSUE_FORMS = ".github/ISSUE_TEMPLATE"
DEPENDABOT = ".github/dependabot.yml"
ISSUE_LABELS = "docs/repo_governance/ISSUE_LABELS.md"
# The assignee list of each update, and the indentation of its first item.
DEPENDABOT_ASSIGNEES = re.compile(r"^([ \t]*assignees:[ \t]*\n)([ \t]*)-[^\n]*\n"
                                  r"(?:[ \t]*-[^\n]*\n)*", re.M)
SETUP_GUIDE = "docs/setup/REPOSITORY_SETUP.md"
SETTINGS_CONFIG = "docs/setup/repository-settings.json"
VERIFICATION_DEPTH = "tools/verification_depth.json"
NOTICE = "docs/licenses/THIRD_PARTY_NOTICES.txt"
PROVENANCE = "docs/licenses/provenance.json"

# The initializer and its test, which it deletes once it has run, and the
# words of the guides that name it among the files of the repository.
INITIALIZER = ("tools/initialize_repository.py", "tests/test_initialize_repository.py")
INITIALIZER_PHRASES: tuple[tuple[str, str, str], ...] = (
    ("README.md", "release, setup and initializer", "release and setup"),
)

# The optional traffic collection that --remove-traffic removes.
TRAFFIC_FILES = (TRAFFIC_WORKFLOW, "docs/repo_governance/TRAFFIC.md",
                 "tools/traffic_history.py", "tests/test_traffic_history.py")
TRAFFIC_REF = "refs/heads/traffic-history"
TRAFFIC_CHECK = "readme-traffic-badges"
# The table rows that describe it, each found by a text it holds; the
# generated rows are those tools/render_docs.py would no longer render.
TRAFFIC_ROWS: tuple[tuple[str, str], ...] = (
    ("README.md", "| Traffic |"),
    ("README.md", "| **Traffic history** |"),
    ("README.md", f"]({TRAFFIC_WORKFLOW})"),
    ("CONTRIBUTING.md", f"| `{TRAFFIC_CHECK}` |"),
    ("docs/repo_governance/README.md", "](TRAFFIC.md)"),
    ("docs/repo_governance/SETTINGS.md", f"`{TRAFFIC_REF}`"),
    (SETUP_GUIDE, f"| `{TRAFFIC_WORKFLOW}`"),
)
# The words that mention it within a sentence, and what replaces them.
TRAFFIC_PHRASES: tuple[tuple[str, str, str], ...] = (
    ("AGENTS.md", " the `traffic-history` branch,", ""),
    ("README.md", "pull requests, traffic, distribution", "pull requests, distribution"),
    ("docs/repo_governance/DISTRIBUTION.md", " The `traffic-history` branch,\nwhen it exists, "
     "holds aggregate counts only; it is scanned with the rest.", ""),
    ("docs/repo_governance/PULL_REQUESTS.md", " The\ndata-only `traffic-history` branch is the "
     "exception: the optional Traffic\nhistory workflow pushes to it directly "
     "([TRAFFIC.md](TRAFFIC.md)).", ""),
    (ISSUE_LABELS, "; outside them, only the optional traffic-alert job writes issues "
     "([TRAFFIC.md](TRAFFIC.md))", ""),
    (ISSUE_LABELS, "the Scorecard analysis, the release workflow's attestation and draft, and "
     "traffic collection write", "the Scorecard analysis and the release workflow's "
     "attestation and draft write"),
    (SETUP_GUIDE, "protection, release tags, and traffic history.",
     "protection and release tags."),
    (NOTICE, "newest-observation tracking in the traffic\nhistory merge, ", ""),
)
MENTION = re.compile(r"traffic", re.I)
# The initializer's own options, which name the collection without describing it.
OPTION = re.compile(r"--remove-traffic(?:-badges)?\b")
SCANNED = (".md", ".json", ".yml", ".yaml", ".txt")


class InitializeError(Exception):
    pass


@dataclass(frozen=True)
class Options:
    repository: str
    regime: str
    release_branch: str | None = None
    remove_traffic_badges: bool = False
    default_branch: str = "main"
    remove_traffic: bool = False
    assignee: str | None = None


@dataclass
class Change:
    """A file's new content (None deletes it) and what changes in it.

    A last change is made only once every other change is written and the
    generated blocks are rendered afresh.
    """

    path: str
    content: str | None
    notes: list[str] = field(default_factory=list)
    last: bool = False


# ---------------------------------------------------------------------------
# One function per file: the adapted text and what changed, or no notes.

Adapted = tuple[str, list[str]]
Adapter = Callable[[str, Options], Adapted]


def replace_sections(text: str, path: str) -> Adapted:
    """Each template-only section replaced by its replacement, or deleted."""
    replaced = deleted = 0
    while SECTION_START in text:
        before, rest = text.split(SECTION_START, 1)
        if SECTION_END not in rest:
            raise InitializeError(f"{path}: {SECTION_START} has no {SECTION_END}")
        section, after = rest.split(SECTION_END, 1)
        if SECTION_START in section:
            raise InitializeError(f"{path}: {SECTION_START} inside a template-only section")
        _, marker, replacement = section.partition(REPLACEMENT)
        body = ""
        if marker:
            closed = replacement.rstrip()
            if not closed.endswith("-->") or SECTION_START in closed:
                raise InitializeError(f"{path}: {REPLACEMENT} is not closed by --> just before "
                                      f"{SECTION_END}")
            body = closed[:-3]
            replaced += 1
        else:
            deleted += 1
        whole_lines = ((not before or before.endswith("\n")) and section.startswith("\n")
                       and section.endswith("\n") and (not after or after.startswith("\n")))
        if whole_lines:
            parts = [part for part in (before.rstrip("\n"), body.strip("\n"),
                                       after.lstrip("\n").rstrip("\n")) if part.strip()]
            text = "\n\n".join(parts) + "\n"
        else:
            text = before + body + after
    notes = []
    if replaced:
        notes.append(f"replace {replaced} template-only section(s) with the replacement each "
                     "holds, with their markers")
    if deleted:
        notes.append(f"delete {deleted} template-only section(s), with their markers")
    return text, notes


def template_only(path: str) -> Adapter:
    def adapt(text: str, options: Options) -> Adapted:
        return replace_sections(text, path)
    return adapt


def adapt_contact_links(text: str, options: Options) -> Adapted:
    own = f"https://github.com/{options.repository}/"
    notes = []
    for url in CONTACT_URL.findall(text):
        if url.lower().startswith(own.lower()):
            continue
        if not GITHUB_URL.match(url):
            raise InitializeError(f".github/ISSUE_TEMPLATE/config.yml: contact link {url} is "
                                  f"not a GitHub repository link; point it at "
                                  f"{options.repository} by hand")
        new = GITHUB_URL.sub(own, url, count=1)
        text = text.replace(url, new)
        notes.append(f"point contact link {url} at {new}")
    return text, notes


def adapt_badges(text: str, options: Options) -> Adapted:
    foreign = sorted({name for name in BADGE.findall(text)
                      if name.lower() != options.repository.lower()})
    if options.remove_traffic_badges or options.remove_traffic:
        lines = text.split("\n")
        badges = sum(1 for line in lines if BADGE.search(line))
        if not badges:
            return text, []
        kept: list[str] = []
        removed = False
        for line in lines:
            if BADGE.search(line):
                removed = True
            # Of the blank lines around a removed badge block, keep one.
            elif not (removed and not line.strip() and kept and not kept[-1].strip()):
                removed = removed and not line.strip()
                kept.append(line)
        return "\n".join(kept), [f"remove the {badges} traffic badge line(s)"]
    if not foreign:
        return text, []
    adapted = BADGE.sub(f"raw.githubusercontent.com/{options.repository}/traffic-history/",
                        text)
    return adapted, [f"point the traffic badges at {options.repository}, not "
                     f"{', '.join(foreign)}"]


def adapt_dependabot(text: str, options: Options) -> Adapted:
    branch = options.release_branch
    current = TARGET_BRANCH.findall(text)
    if branch is None:
        if not current:
            return text, []
        adapted = re.sub(r"^\s*target-branch:.*\n", "", text, flags=re.M)
        return adapted, [f"remove target-branch ({', '.join(sorted(set(current)))}), so that "
                         "updates target the default branch"]
    if current:
        if set(current) == {branch}:
            return text, []
        adapted = re.sub(r"^(\s*target-branch:\s*).*$", rf"\g<1>{branch}", text, flags=re.M)
        return adapted, [f"set target-branch to {branch}, not "
                         f"{', '.join(sorted(set(current)))}"]
    adapted = DIRECTORY.sub(lambda match: f"{match.group(0)}\n{match.group(1)}target-branch: "
                            f"{branch}", text)
    return adapted, [f"add target-branch {branch} to each update"]


def adapt_changelog(text: str, options: Options) -> Adapted:
    # A link reference to another repository marks the template's changelog.
    foreign = foreign_references(text, options.repository, set())
    if not foreign:
        return text, []
    first = VERSION_HEADING.search(text)
    kept = text[:first.start()] if first else text
    introduction, heading, rest = kept.partition("\n## ")
    if NO_RELEASE not in introduction:
        introduction = f"{introduction.rstrip()} **{NO_RELEASE}.**\n"
    kept = (introduction + heading + rest).rstrip()
    adapted = (f"{kept}\n\n## [Unreleased]\n\n[Unreleased]: "
               f"https://github.com/{options.repository}/commits/{options.default_branch}\n")
    entries = len(VERSION_HEADING.findall(text))
    return adapted, [f"replace its {entries} version section(s) and {len(foreign)} link "
                     "reference(s) to another repository with an empty [Unreleased] section; "
                     "the introduction and the policy stay"]


def adapt_releasing(text: str, options: Options) -> Adapted:
    notes = []
    for name, state in STATUS_ROWS.items():
        row = re.compile(rf"^(\|\s*{re.escape(name)}\s*\|)(.*)\|\s*$", re.M)
        match = row.search(text)
        if match and match.group(2).strip() != state:
            text = row.sub(lambda found: f"{found.group(1)} {state} |", text, count=1)
            notes.append(f"set the Current status row {name} to: {state}")
    return text, notes


def regime_paragraph(regime: str) -> str:
    """The house style's paragraph above the naming table, once the prefix is set."""
    reserved = [f"`{prefix}_`" for prefix in RESERVED_PREFIXES]
    return (f"The project's short name is `{regime}`. The [static checks](#static-checks)\n"
            f"read the facade prefix, `{regime}_`, from this table; the reserved prefixes\n"
            f"are {', '.join(reserved[:-1])} and {reserved[-1]}.")


def adapt_house_style(text: str, options: Options) -> Adapted:
    lines = text.split("\n")
    rows = [index for index, line in enumerate(lines)
            if line.startswith("|") and "<REGIME>" in line]
    if not rows:
        match = FACADE_ROW.search(text)
        if match and match.group(1) != options.regime:
            raise InitializeError(f"{HOUSE_STYLE}: the facade prefix is already "
                                  f"{match.group(1)}_, not {options.regime}_; change it by hand")
    notes = []
    for index in rows:
        lines[index] = lines[index].replace("<REGIME>", options.regime)
    if rows:
        notes.append(f"replace <REGIME> with {options.regime} in {len(rows)} row(s) of the "
                     "naming table")
    adapted, requests = REGIME_REQUEST.subn(lambda _: regime_paragraph(options.regime),
                                            "\n".join(lines), count=1)
    if requests:
        notes.append("replace the paragraph that asks for <REGIME> to be replaced with one that "
                     f"names the prefix {options.regime}_")
    return adapted, notes


def table_rows(lines: list[str], header: str) -> list[int]:
    """The indexes of the data rows of the table whose header row starts with header."""
    start = next((index for index, line in enumerate(lines) if line.startswith(header)), None)
    if start is None:
        return []
    rows = []
    for index in range(start + 2, len(lines)):
        if not lines[index].startswith("|"):
            break
        rows.append(index)
    return rows


def adapt_tested_environments(text: str, options: Options) -> Adapted:
    lines = text.split("\n")
    rows = table_rows(lines, TESTED_HEADER)
    if not rows:
        return text, []
    del lines[rows[0]:rows[-1] + 1]
    return "\n".join(lines), [f"remove the {len(rows)} row(s) of the Tested environment table: "
                              "they record the template's Excel runs of its own commits"]


def adapt_host_behaviour(text: str, options: Options) -> Adapted:
    lines = text.split("\n")
    changed = 0
    for index in table_rows(lines, HOST_HEADER):
        row = LAST_CELL.sub(f"| {NO_EVIDENCE} |", lines[index])
        if row != lines[index]:
            lines[index] = row
            changed += 1
    if not changed:
        return text, []
    return "\n".join(lines), [f"replace the Issue column of {changed} Host behaviour row(s) "
                              f"with {NO_EVIDENCE}: they cite the template's issues"]


def replace_phrase(old: str, new: str, what: str) -> Adapter:
    """Replace old, where it is still there, with new."""
    def adapt(text: str, options: Options) -> Adapted:
        if old not in text:
            return text, []
        return text.replace(old, new), [what]
    return adapt


def remove_rows(key: str) -> Adapter:
    """Remove each table row that holds key."""
    def adapt(text: str, options: Options) -> Adapted:
        lines = text.split("\n")
        kept = [line for line in lines if not (line.startswith("|") and key in line)]
        if len(kept) == len(lines):
            return text, []
        return "\n".join(kept), [f"remove the table row that holds {key}"]
    return adapt


def adapt_form_assignee(text: str, options: Options) -> Adapted:
    login = str(options.assignee)
    current = flow_list(text, "assignees")
    if current == [login]:
        return text, []
    adapted, count = re.subn(r"^assignees:.*$", f'assignees: ["{login}"]', text, count=1,
                             flags=re.M)
    if not count:
        raise InitializeError(f"an issue form has no top-level assignees: list; add "
                              f'assignees: ["{login}"] by hand')
    return adapted, [f"assign {login}, not {', '.join(current) or 'nobody'}"]


def adapt_dependabot_assignee(text: str, options: Options) -> Adapted:
    login = str(options.assignee)

    def one(match: re.Match[str]) -> str:
        return f"{match.group(1)}{match.group(2)}- {login}\n"
    adapted, count = DEPENDABOT_ASSIGNEES.subn(one, text)
    if adapted == text:
        return text, []
    return adapted, [f"assign {login} in the {count} update(s)"]


def adapt_json(path: str, edit: Callable[[Any], list[str]]) -> Adapter:
    """Edit a JSON file's data, and write it as the repository formats it."""
    def adapt(text: str, options: Options) -> Adapted:
        try:
            data = json.loads(text)
        except json.JSONDecodeError as error:
            raise InitializeError(f"{path}: {error}") from error
        notes = edit(data)
        if not notes:
            return text, []
        return json.dumps(data, indent=2, ensure_ascii=False) + "\n", notes
    return adapt


def without_traffic_ruleset(config: Any) -> list[str]:
    rulesets = config.get("rulesets") if isinstance(config, dict) else None
    if not isinstance(rulesets, list):
        return []
    traffic = [ruleset for ruleset in rulesets if isinstance(ruleset, dict) and TRAFFIC_REF in
               ((ruleset.get("conditions") or {}).get("ref_name") or {}).get("include", [])]
    if not traffic:
        return []
    config["rulesets"] = [ruleset for ruleset in rulesets if ruleset not in traffic]
    names = ", ".join(f'"{ruleset.get("name")}"' for ruleset in traffic)
    return [f"remove the ruleset {names} of {TRAFFIC_REF} from the configuration",
            f"a ruleset {names} already created on GitHub is left to you: "
            "tools/setup_repository.py reports it as unmanaged and never deletes it; delete it "
            "in Settings > Rules > Rulesets if it exists"]


def without_traffic_check(inventory: Any) -> list[str]:
    entries = inventory.get("checks") if isinstance(inventory, dict) else None
    if not isinstance(entries, list):
        return []
    kept = [entry for entry in entries
            if not (isinstance(entry, dict) and entry.get("check") == TRAFFIC_CHECK)]
    if len(kept) == len(entries):
        return []
    inventory["checks"] = kept
    return [f"remove the entry of {TRAFFIC_CHECK}, which is no longer listed"]


def without_traffic_files(register: Any) -> list[str]:
    removed = 0
    for source in register.get("sources", []) if isinstance(register, dict) else []:
        files = source.get("files") if isinstance(source, dict) else None
        if isinstance(files, list):
            kept = [name for name in files if name not in TRAFFIC_FILES]
            removed += len(files) - len(kept)
            source["files"] = kept
    return [f"remove {removed} entry(ies) of the traffic files"] if removed else []


def adapt_notice(text: str, options: Options) -> Adapted:
    lines = text.split("\n")
    kept = [line for line in lines if line.strip() not in TRAFFIC_FILES]
    if len(kept) == len(lines):
        return text, []
    return "\n".join(kept), [f"remove {len(lines) - len(kept)} line(s) that list the traffic "
                             "files"]


def ruleset_count(before: int, after: int) -> Adapter:
    """The setup guide's counts of the configured rulesets, after a ruleset is removed."""
    count = re.compile(rf"(?<![\d.]){before}(?=[_ ]rulesets\b)")

    def adapt(text: str, options: Options) -> Adapted:
        adapted, found = count.subn(str(after), text)
        return adapted, [f"state {after} rulesets, not {before}, in {found} place(s)"] \
            if found else []
    return adapt


# The template's evidence, which goes with its changelog entries.
EVIDENCE: tuple[tuple[str, Adapter], ...] = (
    (INSTALLATION, adapt_tested_environments),
    (VALIDATION_PLAN, adapt_host_behaviour),
)

ADAPTERS: tuple[tuple[str, Adapter], ...] = (
    *((path, template_only(path)) for path in TEMPLATE_ONLY_FILES),
    (".github/ISSUE_TEMPLATE/config.yml", adapt_contact_links),
    ("README.md", adapt_badges),
    (DEPENDABOT, adapt_dependabot),
    ("CHANGELOG.md", adapt_changelog),
    (HOUSE_STYLE, adapt_house_style),
)


# ---------------------------------------------------------------------------
# Planning and applying.

def read(root: Path, path: str) -> str | None:
    file = root / path
    # Bytes, so that line endings are kept as they are on every platform.
    return file.read_bytes().decode("utf-8") if file.is_file() else None


class Draft:
    """The planned content of each file, in the order the plan first changes it."""

    def __init__(self, root: Path, options: Options) -> None:
        self.root = root
        self.options = options
        self.changes: dict[str, Change] = {}

    def text(self, path: str) -> str | None:
        if path in self.changes:
            return self.changes[path].content
        return read(self.root, path)

    def record(self, path: str, content: str | None, notes: list[str],
               last: bool = False) -> None:
        change = self.changes.setdefault(path, Change(path, content, last=last))
        change.content = content
        change.notes += notes

    def adapt(self, path: str, adapt: Adapter) -> list[str]:
        text = self.text(path)
        if text is None:
            return []
        adapted, notes = adapt(text, self.options)
        if notes:
            self.record(path, adapted, notes)
        return notes

    def delete(self, path: str, note: str, last: bool = False) -> None:
        if self.text(path) is not None:
            self.record(path, None, [note], last)


def plan_traffic_removal(draft: Draft) -> None:
    for path in TRAFFIC_FILES:
        draft.delete(path, "delete it: --remove-traffic removes the traffic collection")
    for path, key in TRAFFIC_ROWS:
        draft.adapt(path, remove_rows(key))
    for path, old, new in TRAFFIC_PHRASES:
        shown = " ".join(old.split())
        draft.adapt(path, replace_phrase(old, new, f'remove "{shown}"' if not new
                                         else f'replace "{shown}" with "{new}"'))
    config = draft.text(SETTINGS_CONFIG)
    before = len(json.loads(config).get("rulesets", [])) if config else 0
    draft.adapt(SETTINGS_CONFIG, adapt_json(SETTINGS_CONFIG, without_traffic_ruleset))
    config = draft.text(SETTINGS_CONFIG)
    after = len(json.loads(config).get("rulesets", [])) if config else 0
    if after != before:
        draft.adapt(SETUP_GUIDE, ruleset_count(before, after))
    draft.adapt(VERIFICATION_DEPTH, adapt_json(VERIFICATION_DEPTH, without_traffic_check))
    draft.adapt(NOTICE, adapt_notice)
    draft.adapt(PROVENANCE, adapt_json(PROVENANCE, without_traffic_files))


def plan_assignee(draft: Draft) -> None:
    login = str(draft.options.assignee)
    bug = draft.text(f"{ISSUE_FORMS}/bug.yml")
    current = flow_list(bug, "assignees") if bug else []
    for name in sorted(FORM_PRIORITIES):
        draft.adapt(f"{ISSUE_FORMS}/{name}", adapt_form_assignee)
    draft.adapt(DEPENDABOT, adapt_dependabot_assignee)
    if len(current) == 1 and current[0] != login:
        draft.adapt(ISSUE_LABELS, replace_phrase(f"`{current[0]}`", f"`{login}`",
                                                 f"name {login} as the assignee, not "
                                                 f"{current[0]}"))


def plan(root: Path, options: Options) -> list[Change]:
    draft = Draft(root, options)
    for path, adapt in ADAPTERS:
        notes = draft.adapt(path, adapt)
        if path == "CHANGELOG.md" and notes:
            # The template's release state goes with its entries.
            draft.delete("VERSION", "delete it: no version of this repository has been "
                                    "released")
            draft.adapt("RELEASING.md", adapt_releasing)
            # So does its evidence, which cites the template's commits and issues.
            for evidence, adapt_evidence in EVIDENCE:
                draft.adapt(evidence, adapt_evidence)
    if options.remove_traffic:
        plan_traffic_removal(draft)
    if options.assignee is not None:
        plan_assignee(draft)
    # Last, the initializer removes itself, and the guides stop naming it.
    if any(read(root, path) is not None for path in INITIALIZER):
        for path, old, new in INITIALIZER_PHRASES:
            draft.adapt(path, replace_phrase(old, new, "stop naming the initializer, which "
                                                       "removes itself"))
        for path in INITIALIZER:
            draft.delete(path, "delete it, last, once every other change is written: the "
                               "initializer removes itself and its test", last=True)
    return list(draft.changes.values())


def traffic_mentions(root: Path, changes: Sequence[Change]) -> list[str]:
    """Where a file still mentions the traffic collection once the changes are made.

    CHANGELOG.md keeps its history, Python code keeps the tests that run only
    while the collection exists, and the lines of the setup guide that describe
    this initializer keep naming what it does.
    """
    planned = {change.path: change.content for change in changes}
    found = []
    for path in gitio.tracked_or_scanned(root):
        if path == "CHANGELOG.md" or not path.endswith(SCANNED):
            continue
        text = planned[path] if path in planned else read(root, path)
        if text is None:
            continue
        for number, line in enumerate(text.split("\n"), 1):
            if path == SETUP_GUIDE and INITIALIZER[0] in line:
                continue
            if MENTION.search(OPTION.sub("", line)):
                found.append(f"{path}:{number}")
    return found


def dirty(root: Path) -> list[str]:
    try:
        status = gitio.text(root, "status", "--porcelain", "--untracked-files=all")
    except GitError as error:
        raise InitializeError(f"{root} is not a Git working tree") from error
    return [line for line in status.splitlines() if line]


def warnings(root: Path, options: Options) -> list[str]:
    """Whether --release-branch exists on origin, asked as dependabot-target-branches asks it.

    The check looks for the remote-tracking branch, and else asks origin, so a
    single-branch or shallow clone of a repository that has the branch needs
    no warning.
    """
    branch = options.release_branch
    if not branch:
        return []
    try:
        if origin_has_branch(root, branch):
            return []
    except MissingContext as missing:
        return [f"{missing}: make sure origin has it before Repository integrity runs, or "
                "the dependabot-target-branches check fails there"]
    return [f"origin has no branch {branch}: push it before Repository integrity runs, or the "
            "dependabot-target-branches check fails"]


def apply(root: Path, changes: Sequence[Change]) -> None:
    for change in changes:
        path = root / change.path
        try:
            if change.content is None:
                path.unlink()
            else:
                path.write_bytes(change.content.encode("utf-8"))
        except OSError as error:
            raise InitializeError(f"{change.path} cannot be written: {error}; the files "
                                  "already written stay, as git status shows, and so does "
                                  "the initializer") from error


def render(root: Path) -> list[str]:
    """Render the generated blocks afresh; return what was rendered, or fail."""
    printed = io.StringIO()
    with contextlib.redirect_stdout(printed):
        problems = render_docs.run(root, write=True)
    if problems:
        raise InitializeError("the generated blocks cannot be rendered, so the initializer "
                              "stays; fix them, then run python tools/render_docs.py --write:"
                              "\n  " + "\n  ".join(problems))
    return printed.getvalue().splitlines()


def validate(options: Options) -> None:
    if not REPOSITORY.fullmatch(options.repository):
        raise InitializeError(f"--repository must be OWNER/NAME, not {options.repository!r}")
    if not REGIME.fullmatch(options.regime) or options.regime in RESERVED_PREFIXES:
        raise InitializeError(f"--regime must be upper-case letters and digits, starting with "
                              f"a letter, and not one of {', '.join(RESERVED_PREFIXES)}: "
                              f"{options.regime!r}")
    for name, branch in (("--release-branch", options.release_branch),
                         ("--default-branch", options.default_branch)):
        if branch is not None and not BRANCH.fullmatch(branch):
            raise InitializeError(f"{name} is not a branch name: {branch!r}")
    login = options.assignee
    if login is not None and not (LOGIN.fullmatch(login) and len(login) <= LOGIN_LENGTH):
        raise InitializeError(f"--assignee must be a GitHub login: letters, digits and single "
                              f"hyphens, neither first nor last, at most {LOGIN_LENGTH} "
                              f"characters, not {login!r}")


def run(root: Path, options: Options, write: bool) -> int:
    validate(options)
    if write and (changed := dirty(root)):
        raise InitializeError("the working tree differs from the commit; commit or discard "
                              "these first, so that the initialization is one diff:\n  "
                              + "\n  ".join(changed))
    changes = plan(root, options)
    found = [*warnings(root, options)]
    if options.remove_traffic:
        found += [f"{place}: still mentions the traffic collection; edit it by hand"
                  for place in traffic_mentions(root, changes)]
    for warning in found:
        print(f"WARNING: {warning}")
    if not changes:
        print(f"initialize-repository: nothing to change; the files are already initialized "
              f"for {options.repository}")
        return 0
    for change in changes:
        for note in change.notes:
            print(f"- {change.path}: {note}")
    count = f"{len(changes)} file(s)"
    if not write:
        print("- generated blocks: rendered afresh from the edited files, as "
              "tools/render_docs.py --write does, before the initializer removes itself")
        print(f"initialize-repository: dry run, nothing written; --apply changes {count}.")
        return 0
    apply(root, [change for change in changes if not change.last])
    for line in render(root):
        print(f"- generated blocks: {line}")
    apply(root, [change for change in changes if change.last])
    print(f"initialize-repository: changed {count} for {options.repository}; review them "
          "with git diff and commit them. Nothing on GitHub was changed.")
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--repository", required=True, metavar="OWNER/NAME",
                        help="the new repository on GitHub")
    parser.add_argument("--regime", required=True, metavar="PREFIX",
                        help="the project's short name for the VBA facade prefix, such as LCR")
    parser.add_argument("--release-branch", metavar="BRANCH",
                        help="Dependabot's target-branch (default: none, so the default branch)")
    parser.add_argument("--default-branch", default="main",
                        help="the branch whose commits the changelog's [Unreleased] link shows "
                             "(default: main)")
    parser.add_argument("--remove-traffic-badges", action="store_true",
                        help="remove the README's traffic badges instead of repointing them, "
                             "and keep the traffic collection")
    parser.add_argument("--remove-traffic", action="store_true",
                        help="remove the optional traffic collection, its badges included")
    parser.add_argument("--assignee", metavar="LOGIN",
                        help="the GitHub account the issue forms and Dependabot assign "
                             "(default: the one they assign now)")
    parser.add_argument("--apply", action="store_true", help="write the changes")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    options = Options(args.repository, args.regime, args.release_branch,
                      args.remove_traffic_badges, args.default_branch, args.remove_traffic,
                      args.assignee)
    try:
        return run(args.root.resolve(), options, args.apply)
    except InitializeError as error:
        print(f"initialize-repository: FAIL: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
