#!/usr/bin/env python3
r"""Check the links in the repository's Markdown files.

Offline, every relative link must name an existing directory or non-empty file
inside the repository, and every fragment must name a heading or an explicit anchor of
the Markdown file it points to. An external link must use HTTPS: a plain http://
link is an error, whether it is an inline link, a reference definition, an
autolink or an HTML attribute.

Links are read the way GitHub renders Markdown. A destination, an href or src
value and an <a id> or <a name> anchor are read with their backslash escapes
removed and their character references decoded (&amp; is &), then URL-decoded
(%20 is a space). Text in code spans, which may continue over line breaks, in
fenced and indented code blocks and in HTML comments is not a link, and neither
is text that only looks like one, such as \[text](path). Blockquote markers are
removed, and list-item markers read, before fences and headings are recognised.
A heading's anchor follows GitHub's slug rules, with the "_" runs that delimit
emphasis removed; the slug that keeps them is accepted too. The checker
approximates GitHub's renderer; it is not a full CommonMark parser.

A report the checker gets wrong is listed, with the reason, in
.github/link-check-exclusions.json. The file is optional; when present, its
form is

    {"schema_version": 1,
     "exclusions": [{"path": "docs/guide.md",
                     "link": "notes.md#part",
                     "reason": "GitHub renders this text as code"}]}

where path is the Markdown file and link is the link as the report prints it.
An entry that matches no report, or a malformed file, is an error, so the list
cannot go stale. A plain http:// link cannot be excluded.

With --external, each https link is also requested once, over HTTPS only and
without following a redirect to another protocol. Absolute links into this
repository itself are not requested, because a private repository answers 404
to anonymous requests; write links to its files as relative links, which the
offline check covers. A link to a page of the repository's wiki is the
exception: when the repository's own page answers an anonymous request, so the
repository is public, the page is requested as its raw Markdown, which answers
404 when the page does not exist; otherwise wiki links are skipped too. The
repository is --repository, else GITHUB_REPOSITORY, else the GitHub repository
of the origin remote, so a local run handles the same links as the workflow.

A 403, 429 or 5xx answer, or no answer at all, is a warning, not an error,
because the site may only be busy or refuse automated requests. To record that
such a link is known and still wanted, list it in .github/link-exceptions.json
with a reason and an expiry date; after that date the warning becomes an error,
so the exception is reviewed again. The file is optional; when present, its
form is

    {"schema_version": 1,
     "exceptions": [{"url": "https://example.org/page",
                     "reason": "Answers 403 to automated requests",
                     "expires": "2027-06-30"}]}

where url is the link without its #fragment, and the expires date itself
still keeps the warning. The offline check fails while the file is malformed
or lists a URL that no Markdown file links.
"""

from __future__ import annotations

import argparse
import html
import json
import os
import re
import subprocess
import sys
import unicodedata
from collections.abc import Iterable, Sequence
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import UTC, date, datetime
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlsplit

from gitio import text_or_none, tracked_or_scanned

ROOT = Path(__file__).resolve().parents[1]
GITHUB_REMOTE = re.compile(r"github\.com[:/]([\w.-]+/[\w.-]+?)(?:\.git)?/?$")
ATTRIBUTE = r"""\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+))"""
EXCEPTIONS = Path(".github/link-exceptions.json")
EXCEPTION_KEYS = {"url", "reason", "expires"}
EXCLUSIONS = Path(".github/link-check-exclusions.json")
EXCLUSION_KEYS = ("path", "link", "reason")
# curl's exit status when a redirect leads to a protocol --proto-redir refuses.
CURL_UNSUPPORTED_PROTOCOL = 1

FENCE = re.compile(r" {0,3}(`{3,}|~{3,})")
REFERENCE = re.compile(r"&(?:#[0-9]{1,7}|#[xX][0-9a-fA-F]{1,6}|[A-Za-z][A-Za-z0-9]{1,31});")
ESCAPABLE = re.compile(r"[!-/:-@\[-`{-~]")  # ASCII punctuation, escapable with a backslash
LIST_ITEM = re.compile(r"( *)([-*+]|\d{1,9}[.)])( {1,4}|$)")
QUOTE = re.compile(r" {0,3}>[ \t]?")
CODE_OR_COMMENT = re.compile(r"<!--|`+")
THEMATIC = re.compile(r" {0,3}([-*_])(?:[ \t]*\1){2,}[ \t]*$")
SETEXT = re.compile(r" {0,3}(?:=+|-+)[ \t]*$")
ATX = re.compile(r" {0,3}#{1,6}(?:[ \t]+(.*?))?(?:[ \t]+#+)?[ \t]*$")
# A line that starts a block of its own, so it never continues a paragraph lazily.
BLOCK_START = re.compile(r" {0,3}(?:`{3,}|~{3,}|#{1,6}(?:[ \t]|$)|<!--"
                         r"|(?:[-*+]|\d{1,9}[.)])(?:[ \t]|$)|([-*_])(?:[ \t]*\1){2,}[ \t]*$)")


@dataclass(frozen=True)
class LinkException:
    """Why a URL's unverifiable answer is tolerated, and until when (inclusive)."""

    reason: str
    expires: date


# ---------------------------------------------------------------------------
# Destinations and anchors.

def decode_references(text: str) -> str:
    """Decode complete character references (ending in ';'), as Markdown and HTML
    attributes do: notes&amp;examples.md names notes&examples.md."""
    return REFERENCE.sub(lambda match: html.unescape(match.group(0)), text)


def decode_destination(text: str) -> str:
    """A Markdown link destination as the renderer reads it: backslash escapes of
    ASCII punctuation removed, then character references decoded."""
    return decode_references(re.sub(r"\\(" + ESCAPABLE.pattern + ")", r"\1", text))


def github_slug(heading: str, emphasis: bool = True) -> str:
    """Anchor GitHub generates for a Markdown heading, before duplicate suffixes.

    GitHub slugs the rendered text. "*" and "~" are punctuation, which the slug
    drops anyway; "_" is kept, so with emphasis (the default) the "_" runs that
    delimit emphasis are removed first: "__API__" gives "api", while
    "snake_case" and "`__init__`" keep their underscores.
    """
    text = re.sub(r"!?\[([^\]]*)\]\([^)]*\)", r"\1", heading)
    if emphasis:
        text = strip_underscore_emphasis(text)
    # Decode only complete character references (ending in ';'), as Markdown does.
    text = decode_references(re.sub(r"<[^>]+>", "", text))
    return "".join(char for char in text.lower() if slug_character(char)).replace(" ", "-")


def strip_underscore_emphasis(text: str) -> str:
    """Remove the "_" runs that delimit emphasis, outside code spans.

    A run can open when it is followed by a non-space and not preceded by a
    letter or digit, and close when it is preceded by a non-space and not
    followed by a letter or digit (CommonMark's rule for "_": never inside a
    word). An opener pairs with the next closer of the same length; unpaired
    runs, and a "_" escaped with a backslash, are literal.
    """
    parts = re.split(r"((`+).+?\2)", text)
    # re.split gives [text, code span, its backticks, text, ...]; only text changes.
    for index in range(0, len(parts), 3):
        parts[index] = strip_segment_emphasis(parts[index])
    return "".join(part for index, part in enumerate(parts) if index % 3 != 2)


def strip_segment_emphasis(segment: str) -> str:
    def side(position: int) -> str:
        return segment[position] if 0 <= position < len(segment) else " "

    # A backslash escapes the first "_" of a run, which stays literal text.
    runs = [(run.start() + escaped(segment, run.start()), run.end())
            for run in re.finditer(r"_+", segment)]
    runs = [(start, end) for start, end in runs if start < end]
    removed: set[int] = set()
    for number, (start, end) in enumerate(runs):
        if number in removed or side(end).isspace() or side(start - 1).isalnum():
            continue
        for later in range(number + 1, len(runs)):
            close_start, close_end = runs[later]
            if (later not in removed and close_end - close_start == end - start
                    and not side(close_start - 1).isspace()
                    and not side(close_end).isalnum()):
                removed |= {number, later}
                break
    kept, position = [], 0
    for number, (start, end) in enumerate(runs):
        if number in removed:
            kept.append(segment[position:start])
            position = end
    return "".join(kept) + segment[position:]


def slug_character(char: str) -> bool:
    """GitHub keeps word characters (letters, marks, decimal digits and connector
    punctuation, such as the variation selector U+FE0F after many emoji), "-" and spaces."""
    return char in "- " or unicodedata.category(char)[:1] in ("L", "M") \
        or unicodedata.category(char) in ("Nd", "Pc")


def heading_anchor(slug: str, used: dict[str, int]) -> str:
    """Make a heading slug unique the way GitHub does: -1, -2... until unused."""
    anchor = slug
    while anchor in used:
        used[slug] += 1
        anchor = f"{slug}-{used[slug]}"
    used[anchor] = 0
    return anchor


def heading_anchors(heading: str, used: dict[str, int]) -> set[str]:
    """The anchor of a heading, plus its slug with "_" emphasis left in place, which
    is accepted too: where GitHub's rendering is ambiguous, no error is reported."""
    return {heading_anchor(github_slug(heading), used), github_slug(heading, emphasis=False)}


def inline_destinations(text: str) -> list[str]:
    r"""Destinations of [text](destination "title") and ![alt](destination), allowing
    <...> and balanced parentheses. A "](" counts only after an unescaped "[" in the
    same block of text, so \[text](path) is plain text."""
    found = []
    openers = 0
    index = 0
    while index < len(text):
        char = text[index]
        if char == "\\" and ESCAPABLE.fullmatch(text[index + 1:index + 2]):
            index += 2
            continue
        if char == "[":
            openers += 1
        elif char == "]" and openers:
            openers -= 1
            if text.startswith("(", index + 1):
                destination = link_destination(text[index + 2:])
                if destination:
                    found.append(destination)
        index += 1
    return found


def link_destination(rest: str) -> str:
    """The destination at the start of the text after "](", or "" if there is none."""
    rest = rest.lstrip(" \t")
    if rest.startswith("<"):
        end = rest.find(">")
        return rest[1:end] if end > 1 and "\n" not in rest[:end] else ""
    depth, destination, index = 0, "", 0
    while index < len(rest):
        char = rest[index]
        if char == "\\" and ESCAPABLE.fullmatch(rest[index + 1:index + 2]):
            # An escaped character, such as \) or \(, never delimits.
            destination += rest[index:index + 2]
            index += 2
            continue
        if char.isspace() or (char == ")" and depth == 0):
            break
        depth += {"(": 1, ")": -1}.get(char, 0)
        destination += char
        index += 1
    return destination


def escaped(text: str, position: int) -> bool:
    """Whether the character at position follows an odd number of backslashes."""
    start = position
    while start > 0 and text[start - 1] == "\\":
        start -= 1
    return (position - start) % 2 == 1


def strip_code_and_comments(text: str) -> str:
    """Remove the code spans and HTML comments of one block of text, which may
    span several lines.

    They are read left to right, as in CommonMark, so whichever starts first
    wins: a backtick inside a comment is part of the comment, and "<!--" inside
    a code span is code. A backtick run opens a span only if it is not escaped
    (a backslash escapes its first backtick only) and an equally long run
    follows in the same block; "<!--" opens a comment only if it is not escaped
    and "-->" follows. Otherwise they are literal. The newlines inside a removed
    span or comment are kept, so the text keeps its lines.
    """
    kept, position, index = [], 0, 0
    while match := CODE_OR_COMMENT.search(text, index):
        opener = match.start()
        if escaped(text, opener):
            index = opener + 1
            continue
        if match.group().startswith("<"):
            found = text.find("-->", match.end())
            close = found + 3 if found >= 0 else None
        else:
            length = len(match.group())
            closer = re.compile(rf"(?<!`)`{{{length}}}(?!`)").search(text, match.end())
            close = closer.end() if closer else None
        if close is None:
            index = match.end()
            continue
        kept += [text[position:opener], "\n" * text.count("\n", opener, close)]
        position = index = close
    return "".join(kept) + text[position:]


def line_links(line: str) -> tuple[list[str], set[str]]:
    """Reference definitions, autolinks, HTML links and explicit anchors on one line
    whose code spans and comments are already removed, decoded the way GitHub reads them."""
    anchors = {decode_references("".join(value)) for value in
               re.findall(r"<a\s[^>]*?(?<![\w-])(?:id|name)" + ATTRIBUTE, line)}
    links = [decode_destination("".join(value)) for value in
             re.findall(r"^ {0,3}\[(?!\^)[^\]]+\]:\s*(?:<([^>]*)>|(\S+))", line)]
    links += re.findall(r"<(https?://[^>\s]+)>", line)
    links += [decode_references("".join(value)) for value in
              re.findall(r"(?<![\w-])(?:href|src)" + ATTRIBUTE, line)]
    return links, anchors


def block_links(lines: Sequence[str]) -> tuple[list[str], set[str]]:
    """Links and explicit anchors of one block of text lines, outside its code spans
    and HTML comments."""
    text = strip_code_and_comments("\n".join(lines))
    links = [decode_destination(link) for link in inline_destinations(text)]
    anchors: set[str] = set()
    for line in text.split("\n"):
        found, explicit = line_links(line)
        links += found
        anchors |= explicit
    return links, anchors


# ---------------------------------------------------------------------------
# Blocks.

def indentation(line: str) -> int:
    """Leading columns of a line, with tabs to the next multiple of 4."""
    expanded = line.expandtabs(4)
    return len(expanded) - len(expanded.lstrip(" "))


def strip_quotes(line: str, limit: int | None = None) -> tuple[int, str]:
    """Blockquote depth of a line, and the line without its ">" markers; at most
    limit markers are read, so a ">" inside a quoted code block stays text."""
    depth = 0
    while (limit is None or depth < limit) and (match := QUOTE.match(line)):
        depth, line = depth + 1, line[match.end():]
    return depth, line


def setext_candidate(text: str) -> bool:
    """Whether a line of a paragraph can be part of a setext heading's text."""
    return bool(text.strip()) and not re.match(r" {0,3}([-*+]\s|\d+[.)]\s|[>|<])", text)


class MarkdownReader:
    """Read a Markdown document line by line, collecting its links and anchors.

    Blockquote markers are removed, and list-item markers read, before fences,
    headings, setext underlines and thematic breaks are recognised. Indented
    code follows CommonMark conservatively: a line indented four columns beyond
    the content of the open list item (or beyond the margin, outside lists) is
    code unless it continues a paragraph, which code cannot interrupt.
    """

    def __init__(self) -> None:
        self.links: list[str] = []
        self.anchors: set[str] = set()
        self.used: dict[str, int] = {}
        self.fence: tuple[str, int, int] | None = None  # (opening run, content column, quote depth)
        self.comment = False         # inside an HTML comment block
        self.quote = 0               # blockquote depth of the previous line
        self.paragraph: list[str] = []  # candidate setext heading text
        self.block: list[str] = []      # lines whose code spans and links are read together
        self.items: list[int] = []      # content columns of the open list items
        self.previous = "blank"         # blank, text, code or closed (heading, fence, break)

    def flush(self) -> None:
        found, explicit = block_links(self.block)
        self.links.extend(found)
        self.anchors.update(explicit)
        self.block.clear()

    def close(self, previous: str) -> None:
        """End the current block and the paragraph."""
        self.flush()
        self.paragraph, self.previous = [], previous

    def read(self, raw: str) -> None:
        depth, line = strip_quotes(raw, self.fence[2] if self.fence else None)
        lazy = (depth < self.quote and self.previous == "text" and not self.fence
                and bool(line.strip()) and not BLOCK_START.match(line))
        if lazy:
            depth = self.quote   # a lazy continuation line of the quoted paragraph
        if depth != self.quote:
            # A new or closed blockquote: its blocks, lists and fences end here.
            self.flush()
            if self.fence and depth < self.fence[2]:
                self.fence = None
            self.items, self.paragraph, self.previous, self.quote = [], [], "blank", depth
        if self.fence and not self.read_fenced(line):
            return
        if self.comment:
            # An HTML comment block ends on the line that contains "-->".
            if "-->" in line:
                self.comment, self.previous = False, "closed"
            return
        if not line.strip():
            self.close("blank")
            return
        self.read_content(line, lazy)

    def read_fenced(self, line: str) -> bool:
        """Read a line inside a fenced code block; True when the line ended the
        list item that holds the fence and must be read as ordinary text."""
        assert self.fence is not None
        run, column, _ = self.fence
        expanded = line.expandtabs(4)
        indent = indentation(line)
        if column and line.strip() and indent < column:
            # The list item, and with it the fence, ends left of its content.
            self.fence = None
            return True
        relative = expanded[min(indent, column):]
        match = FENCE.match(relative)
        if (match and match.group(1)[0] == run[0] and len(match.group(1)) >= len(run)
                and not relative[match.end():].strip()):
            self.fence, self.previous = None, "closed"
        return False

    def read_content(self, line: str, lazy: bool) -> None:
        expanded = line.expandtabs(4)
        indent = indentation(line)
        base = self.items[-1] if self.items else 0
        thematic = None if lazy else THEMATIC.match(expanded[min(indent, base):])
        item = None if thematic or lazy or indent >= base + 4 else LIST_ITEM.match(expanded)
        starts_block = indent < base + 4 and BLOCK_START.match(expanded.lstrip(" "))
        if item or self.previous != "text" or starts_block:
            # A list item ends where a line starts left of its content.
            while self.items and indent < self.items[-1]:
                self.items.pop()
            base = self.items[-1] if self.items else 0
        if not item and indent >= base + 4 and self.previous != "text":
            self.close("code")
            return
        if item:
            self.flush()
            # The item's content is read as a block of its own: a fence or a heading
            # may start on the marker's line.
            base = len(item.group(0)) + (0 if item.group(3) else 1)
            self.items.append(base)
            relative = expanded[base:]
        else:
            relative = expanded[min(indent, base):]
        if not self.opens_block(relative, base):
            self.read_text(relative, item is not None, bool(thematic),
                           lazy or indent < base)

    def opens_block(self, relative: str, base: int) -> bool:
        """Open a fenced code block or an HTML comment block at the line, if it starts one."""
        match = FENCE.match(relative)
        if match and not (match.group(1)[0] == "`" and "`" in relative[match.end():]):
            self.close("closed")
            self.fence = (match.group(1), base, self.quote)
            return True
        if re.match(r" {0,3}<!--", relative):
            self.close("closed")
            self.comment = "-->" not in relative[relative.index("<!--") + 4:]
            return True
        return False

    def read_text(self, relative: str, item: bool, thematic: bool, outside: bool) -> None:
        """Read a line of text: a heading, a thematic break or a paragraph line.
        outside is True for a line left of its container's content, which can
        continue a paragraph but cannot underline it."""
        if self.paragraph and not item and not outside and SETEXT.match(relative):
            # Setext heading: the paragraph above an === or --- underline.
            self.anchors |= heading_anchors(" ".join(self.paragraph), self.used)
            self.close("closed")
            return
        heading = ATX.match(relative)
        if heading or thematic:
            self.flush()
            self.block.append(relative)
            if heading:
                self.anchors |= heading_anchors(heading.group(1) or "", self.used)
            self.close("closed")
            return
        if not setext_candidate(relative):
            self.paragraph = []
        elif item:
            self.paragraph = [relative.strip()]
        else:
            self.paragraph = [*self.paragraph, relative.strip()]
        self.block.append(relative)
        self.previous = "text"


def markdown_links(text: str) -> tuple[list[str], set[str]]:
    """Return (links, anchors) of a Markdown document, ignoring fenced and indented
    code blocks, code spans and HTML comments."""
    reader = MarkdownReader()
    for line in text.splitlines():
        reader.read(line)
    reader.flush()
    return reader.links, reader.anchors


# ---------------------------------------------------------------------------
# Offline check.

def local_error(root: Path, source: Path, link: str, anchors: dict[Path, set[str]]) -> str | None:
    url = urlsplit(link)
    if url.path.startswith("/"):
        return f"root-relative link: {link}"
    target = (source.parent / unquote(url.path)).resolve() if url.path else source.resolve()
    if not target.is_relative_to(root.resolve()):
        return f"path outside the repository: {link}"
    if not target.exists():
        return f"missing target: {link}"
    if target.is_file() and target.stat().st_size == 0:
        return f"empty target: {link}"
    if url.fragment and target in anchors and unquote(url.fragment) not in anchors[target]:
        return f"missing anchor: {link}"
    return None


def exception_problem(entry: Any, seen: dict[str, LinkException]) -> str | None:
    """Why an entry of the exceptions file is malformed, or None."""
    if not isinstance(entry, dict) or set(entry) != EXCEPTION_KEYS:
        return "needs exactly url, reason and expires"
    url, reason, expires = entry["url"], entry["reason"], entry["expires"]
    if not isinstance(url, str) or not url.startswith("https://") or not urlsplit(url).netloc \
            or "#" in url or url != url.strip():
        return f"url must be an https URL without a fragment: {url!r}"
    if url in seen:
        return f"{url} is listed twice"
    if not isinstance(reason, str) or not reason.strip():
        return f"{url} needs a reason"
    if not isinstance(expires, str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", expires):
        return f"{url} needs expires as YYYY-MM-DD"
    try:
        date.fromisoformat(expires)
    except ValueError:
        return f"{url} has an expires date that does not exist: {expires}"
    return None


def read_list_file(root: Path, relative: Path, key: str) -> tuple[list[Any], list[str]]:
    """Return (entries, problems) of an optional {"schema_version": 1, key: [...]} file."""
    path, name = root / relative, relative.as_posix()
    if not path.is_file():
        return [], []
    try:
        data: Any = json.loads(path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, ValueError) as error:
        return [], [f"{name}: not valid UTF-8 JSON ({error})"]
    if not isinstance(data, dict) or set(data) != {"schema_version", key} \
            or type(data["schema_version"]) is not int or data["schema_version"] != 1 \
            or not isinstance(data[key], list):
        return [], [f"{name}: needs exactly schema_version 1 and an {key} list"]
    return data[key], []


def load_exceptions(root: Path) -> tuple[dict[str, LinkException], list[str]]:
    """Return (exceptions by URL, problems) of the optional exceptions file."""
    entries, problems = read_list_file(root, EXCEPTIONS, "exceptions")
    exceptions: dict[str, LinkException] = {}
    for index, entry in enumerate(entries):
        problem = exception_problem(entry, exceptions)
        if problem:
            problems.append(f"{EXCEPTIONS.as_posix()}: exception {index + 1}: {problem}")
        else:
            exceptions[entry["url"]] = LinkException(entry["reason"].strip(),
                                                     date.fromisoformat(entry["expires"]))
    return exceptions, problems


def exclusion_problem(entry: Any, seen: dict[tuple[str, str], str]) -> str | None:
    """Why an entry of the exclusions file is malformed, or None."""
    if not isinstance(entry, dict) or set(entry) != set(EXCLUSION_KEYS):
        return "needs exactly path, link and reason"
    for key in EXCLUSION_KEYS:
        value = entry[key]
        if not isinstance(value, str) or not value.strip() or value != value.strip():
            return f"needs a non-empty {key} without surrounding spaces"
    if urlsplit(entry["link"]).scheme.lower() == "http":
        return f"a plain HTTP link cannot be excluded: {entry['link']}"
    if (entry["path"], entry["link"]) in seen:
        return f"{entry['path']}: {entry['link']} is listed twice"
    return None


def load_exclusions(root: Path) -> tuple[dict[tuple[str, str], str], list[str]]:
    """Return (reasons by (Markdown file, link), problems) of the optional exclusions file."""
    entries, problems = read_list_file(root, EXCLUSIONS, "exclusions")
    exclusions: dict[tuple[str, str], str] = {}
    for index, entry in enumerate(entries):
        problem = exclusion_problem(entry, exclusions)
        if problem:
            problems.append(f"{EXCLUSIONS.as_posix()}: exclusion {index + 1}: {problem}")
        else:
            exclusions[(entry["path"], entry["link"])] = entry["reason"]
    return exclusions, problems


def check(root: Path = ROOT) -> tuple[list[str], set[str], int]:
    """Return (errors, external URLs, number of links checked) for the Markdown files.

    A report listed in the exclusions file is not an error; an exclusion that
    matches no report is.
    """
    errors: list[str] = []
    documents = {}
    for path in (root / name for name in tracked_or_scanned(root, "*.md")):
        if not path.is_file():
            # Tracked by Git but deleted from the working tree without git rm.
            errors.append(f"{path.relative_to(root).as_posix()}: tracked but missing from the working tree")
            continue
        documents[path.resolve()] = markdown_links(path.read_text(encoding="utf-8"))
    anchors = {path: found for path, (_, found) in documents.items()}
    exclusions, exclusion_problems = load_exclusions(root)
    matched: set[tuple[str, str]] = set()
    external: set[str] = set()
    count = 0
    for source, (links, _) in sorted(documents.items()):
        name = source.relative_to(root.resolve()).as_posix()
        for link in links:
            count += 1
            url = urlsplit(link)
            if url.scheme == "http":
                errors.append(f"{name}: plain HTTP link, use https: {link}")
                continue
            if url.scheme == "https" or (not url.scheme and url.netloc):
                external.add(("https:" + link if link.startswith("//") else link).split("#")[0])
                continue
            if url.scheme in ("mailto", "tel", "data"):
                continue
            problem = (f"unsupported URL scheme: {link}" if url.scheme
                       else local_error(root, source, link, anchors))
            if problem and (name, link) in exclusions:
                matched.add((name, link))
            elif problem:
                errors.append(f"{name}: {problem}")
    exceptions, problems = load_exceptions(root)
    errors += problems
    errors += [f"{EXCEPTIONS.as_posix()}: {url} is not linked from any Markdown file; "
               "remove its exception" for url in sorted(set(exceptions) - external)]
    errors += exclusion_problems
    errors += [f"{EXCLUSIONS.as_posix()}: {path}: {link} is no longer reported; "
               "remove its exclusion" for path, link in sorted(set(exclusions) - matched)]
    return errors, external, count


# ---------------------------------------------------------------------------
# External check.

def detect_repository(root: Path) -> str:
    """owner/name of the origin remote on GitHub, or "" when there is none."""
    url = text_or_none(root, "remote", "get-url", "origin")
    match = GITHUB_REMOTE.search(url.strip()) if url is not None else None
    return match.group(1) if match else ""


def own_repository(url: str, repository: str) -> bool:
    """Whether a URL points into this repository; such URLs are not requested."""
    if not repository:
        return False
    prefixes = (f"https://github.com/{repository}", f"https://raw.githubusercontent.com/{repository}/")
    lowered = url.lower()
    return any(lowered == prefix.lower() or lowered.startswith(prefix.lower().rstrip("/") + "/")
               for prefix in prefixes)


def wiki_page(url: str, repository: str) -> str | None:
    """The raw Markdown URL of a page of this repository's wiki, which answers 404
    when the page does not exist, or None when the URL names no wiki page. The
    wiki's home page, /wiki, is Home; the wiki's own views, such as _pages and
    _history, are not pages."""
    if not repository:
        return None
    prefix = f"https://github.com/{repository}/wiki".lower()
    path = url.split("#")[0].split("?")[0]
    if path.lower() != prefix and not path.lower().startswith(prefix + "/"):
        return None
    page = re.fullmatch(r"/?([^/]*)/?", path[len(prefix):])
    if page is None or page.group(1).startswith("_"):
        return None
    return f"https://raw.githubusercontent.com/wiki/{repository}/{page.group(1) or 'Home'}.md"


def check_external(url: str) -> tuple[str, str]:
    # GET follows redirects; HEAD is rejected by some otherwise working sites.
    result = subprocess.run([
        "curl", "--silent", "--show-error", "--location", "--output", os.devnull,
        "--write-out", "%{http_code}", "--connect-timeout", "10", "--max-time", "25",
        "--retry", "2", "--retry-delay", "2", "--proto", "=https",
        "--proto-redir", "=https", url,
    ], capture_output=True, text=True, timeout=95, check=False)
    code = result.stdout.strip()
    if result.returncode == CURL_UNSUPPORTED_PROTOCOL:
        return "ERROR", f"{url}: redirects to a URL that is not HTTPS"
    if result.returncode or code in ("403", "429") or code.startswith("5"):
        return "WARN", f"{url}: unavailable for verification (HTTP {code or 'unknown'})"
    if not code.startswith("2"):
        return "ERROR", f"{url}: HTTP {code}"
    return "PASS", url


def check_request(url: str, target: str) -> tuple[str, str]:
    """Request target for the link url, and report the answer under the link."""
    status, message = check_external(target)
    if target == url:
        return status, message
    return status, message.replace(target, f"{url} (requested as {target})", 1)


def apply_exception(status: str, message: str, exception: LinkException | None,
                    today: date) -> tuple[str, str]:
    """Keep a tolerated answer a warning until its exception expires, then fail it."""
    if status != "WARN" or exception is None:
        return status, message
    if today > exception.expires:
        return "ERROR", (f"{message}; its exception in {EXCEPTIONS.as_posix()} expired on "
                         f"{exception.expires.isoformat()} ({exception.reason})")
    return status, (f"{message}; tolerated until {exception.expires.isoformat()} "
                    f"({exception.reason})")


def wiki_requests(urls: Iterable[str], repository: str) -> tuple[list[tuple[str, str]],
                                                                  list[tuple[str, str, str]]]:
    """Return ((link, raw page) to request, (link, status, message) not requested)
    for links to this repository's wiki. They are requested only when the
    repository answers an anonymous request: a private repository's wiki answers
    404 even where the page exists."""
    pages = [(url, page) for url in urls if (page := wiki_page(url, repository))]
    if not pages:
        return [], []
    home = f"https://github.com/{repository}"
    status, message = check_external(home)
    if status == "PASS":
        return pages, []
    if status == "WARN":
        return [], [(url, "WARN", f"{url}: not requested, because {message}")
                    for url, _ in pages]
    return [], [(url, "SKIP", f"{url}: {home} does not answer anonymous requests ({message}), "
                 "so its wiki is not checked") for url, _ in pages]


def external_errors(urls: Iterable[str], repository: str,
                    exceptions: dict[str, LinkException] | None = None,
                    today: date | None = None) -> list[str]:
    links = sorted(urls)
    own = [url for url in links if own_repository(url, repository)]
    requests, results = wiki_requests(own, repository)
    wiki = {url for url, _ in requests} | {url for url, _, _ in results}
    for url in own:
        if url not in wiki:
            print(f"SKIP: {url}: a link into this repository, which is not requested")
    requests += [(url, url) for url in links if url not in own]
    requests.sort()
    listed = exceptions or {}
    day = today or datetime.now(UTC).date()
    errors = []
    with ThreadPoolExecutor(max_workers=3) as pool:
        answers = pool.map(lambda request: check_request(*request), requests)
        results += [(url, *answer) for (url, _), answer in zip(requests, answers, strict=True)]
    for url, result_status, result_message in sorted(results):
        status, message = apply_exception(result_status, result_message, listed.get(url), day)
        print(f"{status}: {message}")
        if status == "ERROR":
            errors.append(message)
    return errors


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--external", action="store_true",
                        help="also request every external https link")
    parser.add_argument("--repository", default=os.environ.get("GITHUB_REPOSITORY", ""),
                        help="owner/name whose own URLs are skipped, except its wiki pages "
                             "while it is public (default: GITHUB_REPOSITORY, else the origin "
                             "remote)")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    errors, external, count = check(root)
    if args.external:
        repository = args.repository or detect_repository(root)
        errors += external_errors(external, repository, load_exceptions(root)[0])
    for error in errors:
        print(f"ERROR: {error}")
    print(f"Checked {count} link(s) in Markdown; {len(errors)} error(s).")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
