"""Compare two versions of an exported VBA component by its code, without comments.

A comment does not change what Excel runs. This module decides whether two
versions of a VBA component (.bas, .cls or .frm) hold the same code, so that
tools/check_excel_run.py --carried-to can carry the evidence of an Excel run
across a commit that changes only comments or layout. The code of a version
is its code stream: its logical lines as tools/vba_lexer.py reads them, the
lexer the VBA source checks use, with

- comments removed: from an apostrophe outside a string, or from Rem where a
  statement starts, and the lines of a comment continued with " _";
- continuation lines joined into their logical line, and blank lines dropped;
- conditional-compilation directives (#If, #Const ...) kept, and marked as
  directives, so that a changed directive or a changed statement inside an
  #If block is a code change;
- outside string literals, each run of white space made one space, and white
  space at the start and end of a line dropped. Inside a string literal every
  character is kept. White space is never inserted or removed between two
  tokens, so `a=1` and `a = 1` are different code;
- outside string literals, letters compared without case, because the VBE
  re-cases every use of a name (HB-01, as tools/check_excel_run.py compares an
  export with its commit). The re-cased identifiers are reported, not refused.

Any other difference is a code change. The comparison fails closed: a
non-ASCII byte, a commit or a file that cannot be read, or a file that is not
a VBA component refuses it (CarryRefused). A component that uses [ or ]
outside a string literal is always refused, changed or not: inside a [...]
expression an apostrophe, a colon, Rem and a quote do not mean what they mean
in a statement, so its comments cannot be told from its code with certainty.
A form's .frx resource is binary and never compared here; it must stay byte
for byte identical.
"""

from __future__ import annotations

import re
import subprocess
from dataclasses import dataclass
from pathlib import Path

import vba_lexer
from vba_lexer import VBA_SUFFIXES

IDENTIFIER = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")
# How much of a statement a message quotes.
QUOTED = 80


class CarryRefused(Exception):
    """The two versions cannot be shown to hold the same code."""


@dataclass(frozen=True)
class Statement:
    """One logical line of code: the physical line it starts on, its kind
    ("code" or "directive"), its text with white space collapsed, and that
    text with the letters outside string literals in lower case."""

    line: int
    kind: str
    code: str
    key: str


def decode(raw: bytes) -> str:
    """The component's text with LF line ends; refuses any byte that is not ASCII."""
    try:
        text = raw.decode("ascii")
    except UnicodeDecodeError as error:
        raise CarryRefused(f"non-ASCII byte at offset {error.start}") from None
    return text.replace("\r\n", "\n").replace("\r", "\n")


def normalize(code: str) -> str:
    """``code`` with each run of white space outside string literals made one space,
    and none at either end."""
    out: list[str] = []
    in_string = pending_space = False
    for char in code:
        if in_string:
            out.append(char)
            # A doubled quote closes the literal and opens it again: same result.
            in_string = char != '"'
        elif char.isspace():
            pending_space = True
        else:
            if pending_space and out:
                out.append(" ")
            pending_space = False
            out.append(char)
            in_string = char == '"'
    return "".join(out)


def case_key(code: str) -> str:
    """``code`` with the letters outside string literals in lower case."""
    out: list[str] = []
    in_string = False
    for char in code:
        if char == '"':
            in_string = not in_string
        out.append(char if in_string or not "A" <= char <= "Z" else char.lower())
    return "".join(out)


def outside_strings(code: str) -> str:
    """``code`` without its string literals."""
    out: list[str] = []
    in_string = False
    for char in code:
        if char == '"':
            in_string = not in_string
        elif not in_string:
            out.append(char)
    return "".join(out)


def code_stream(raw: bytes) -> list[Statement]:
    """The comment-free logical lines of a component, in order."""
    stream = []
    for first, _last, kind, code in vba_lexer.logical_units(decode(raw).split("\n")):
        normal = normalize(code)
        if normal:
            stream.append(Statement(first, kind, normal, case_key(normal)))
    return stream


def bracket_statement(stream: list[Statement]) -> Statement | None:
    """The first statement that uses [ or ] outside a string literal, or None."""
    for statement in stream:
        plain = outside_strings(statement.code)
        if "[" in plain or "]" in plain:
            return statement
    return None


def quoted(statement: Statement) -> str:
    return repr(statement.code[:QUOTED])


def first_difference(executed: bytes, current: bytes) -> str | None:
    """None when the two versions hold the same code, else where the code first
    differs, with the line of each version. CarryRefused on a non-ASCII byte."""
    old, new = code_stream(executed), code_stream(current)
    for stream, label in ((old, "run"), (new, "current")):
        found = bracket_statement(stream)
        if found is not None:
            return (f"the {label} version uses [ or ] outside a string at line {found.line}, "
                    f"which is never carried: {quoted(found)}")
    for a, b in zip(old, new, strict=False):
        if (a.kind, a.key) != (b.kind, b.key):
            return f"line {a.line} at the run, line {b.line} now: {quoted(a)} became {quoted(b)}"
    if len(old) != len(new):
        longer, label = (old, "run") if len(old) > len(new) else (new, "current")
        extra = longer[min(len(old), len(new))]
        return (f"{abs(len(old) - len(new))} statement(s) only in the {label} version, "
                f"the first at line {extra.line}: {quoted(extra)}")
    return None


def recased(executed: bytes, current: bytes) -> list[str]:
    """The identifiers whose letter case differs between two versions of the same
    code, as "run spelling -> current spelling", each once."""
    found: dict[str, None] = {}
    for a, b in zip(code_stream(executed), code_stream(current), strict=False):
        if a.code == b.code or a.key != b.key:
            continue
        for match in IDENTIFIER.finditer(b.code):
            spelling = a.code[match.start():match.end()]
            if spelling != match.group():
                found[f"{spelling} -> {match.group()}"] = None
    return list(found)


def read_blob(root: Path, commit: str, path: str) -> bytes:
    """The file ``path`` as the commit stores it; CarryRefused when the commit or the
    file cannot be read."""
    if not commit or commit.startswith("-"):
        raise CarryRefused(f"no valid commit named for {path}: {commit!r}")
    try:
        completed = subprocess.run(["git", "-C", str(root), "cat-file", "blob",
                                    f"{commit}:{path}"], capture_output=True, check=False)
    except OSError as error:
        raise CarryRefused(f"git is not available: {error}") from None
    if completed.returncode:
        raise CarryRefused(f"{path} cannot be read at commit {commit}")
    return completed.stdout


def carried(root: Path, path: str, run_commit: str, commit: str) -> list[str]:
    """Prove that the VBA component ``path`` holds at ``commit`` the code it held at
    ``run_commit``, the commit an Excel run executed. Returns the identifiers
    re-cased between them; raises CarryRefused with the reason otherwise."""
    if not path.lower().endswith(VBA_SUFFIXES):
        raise CarryRefused(f"{path} is not a VBA component ({', '.join(VBA_SUFFIXES)})")
    executed = read_blob(root, run_commit, path)
    current = read_blob(root, commit, path)
    difference = first_difference(executed, current)
    if difference is not None:
        raise CarryRefused(f"code changed: {difference}")
    return recased(executed, current)
