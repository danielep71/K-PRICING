#!/usr/bin/env python3
"""Write a mutant of one source file: a copy with one piece of text replaced.

A mutant is the defective variant of a file, made reproducibly for a negative
control: the test that proves a fix must fail on it and pass on the fix
(CONTRIBUTING.md, *Negative controls*). The tool replaces the text that
--find gives with the text that --replace gives and changes nothing else.

    python tools/make_mutant.py FILE --find TEXT --replace TEXT [--output DIR]

The text is literal: no escape sequence and no pattern is interpreted. For a
text that spans lines, put it in a file and pass --find-file or
--replace-file instead; the whole file is the text, a final line ending
included. Line endings in the text are converted to those of FILE, so the
same text matches a CRLF and an LF checkout.

The mutant keeps FILE's path relative to the repository root, under the
output folder, test-results/mutants/ of the repository by default. FILE is
read and the mutant written as bytes, so every byte outside the replaced text,
CRLF line endings included, stays as it was. For an exported VBA component (a
.bas, .cls or .frm file), the Attribute VB_Name line must exist and stay the
same, and the replacement must be ASCII; a form's .frx is copied beside its
mutant, since the VBE imports a form only with it.

It refuses, and writes nothing, when:

- the text to find is not in FILE exactly once, overlapping occurrences
  included, so "aa" in "aaa" counts twice (it reports the count);
- the replacement would leave FILE unchanged;
- the replacement would change, remove or add an Attribute VB_Name line;
- the text spans lines and FILE mixes CRLF and LF line endings;
- the mutant's path, a form's .frx beside it or a folder between them and
  the output folder is a symbolic link, or the path resolves outside the
  output folder, so that a link left there never redirects the write;
- the output folder holds a tracked file, lies in a source directory (src/,
  tests/, tools/, docs/, .github/ or examples/), or Git does not ignore the
  mutant's path, so that a mutant can never be committed or replace a source;
- FILE or the output folder is outside the repository.

On success it prints the changed lines, with the line number where the
replaced text starts, and the SHA-256 of the source and of the mutant.
Each output is written to a new file and renamed over its path, so a file
already there, even a hard link to the source, is replaced and never
written through. Running it again replaces the mutant. Exit status: 0 when the mutant is
written, 2 when it refuses or on a usage error.
"""

from __future__ import annotations

import argparse
import hashlib
import os
import re
import sys
import tempfile
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

import gitio
from gitio import GitError
from vba_lexer import VBA_SUFFIXES

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = Path("test-results") / "mutants"
# The directories that hold authoritative source, as .gitignore lists them.
SOURCE_DIRECTORIES = ("src", "tests", "tools", "docs", ".github", "examples")
VB_NAME = re.compile(rb'^Attribute VB_Name = "[^"]*"[ \t]*\r?$', re.M)


class Refusal(Exception):
    """The mutant cannot be written as asked; nothing was written."""


@dataclass(frozen=True)
class Mutant:
    data: bytes
    line: int
    before: bytes
    after: bytes


def line_ending(data: bytes) -> bytes | None:
    """CRLF or LF when FILE uses only that line ending, None when it mixes them."""
    crlf = data.count(b"\r\n")
    if crlf == data.count(b"\n") and crlf:
        return b"\r\n"
    if crlf == 0:
        return b"\n"
    return None


def as_file_text(text: bytes, data: bytes) -> bytes:
    """``text`` with its line endings converted to those of ``data``."""
    if b"\n" not in text:
        return text
    ending = line_ending(data)
    if ending is None:
        raise Refusal("the file mixes CRLF and LF line endings, so a text that spans "
                      "lines cannot match it reliably")
    return text.replace(b"\r\n", b"\n").replace(b"\n", ending)


def occurrences(data: bytes, find: bytes) -> int:
    """How often ``find`` occurs in ``data``, overlapping occurrences included."""
    count, position = 0, data.find(find)
    while position >= 0:
        count += 1
        position = data.find(find, position + 1)
    return count


def mutate(data: bytes, find: bytes, replace: bytes, vba: bool) -> Mutant:
    """Replace the only occurrence of ``find`` in ``data`` with ``replace``."""
    find, replace = as_file_text(find, data), as_file_text(replace, data)
    count = occurrences(data, find)
    if count != 1:
        raise Refusal(f"the text to find occurs {count} times, overlapping occurrences "
                      f"included; it must occur exactly once")
    if find == replace:
        raise Refusal("the replacement equals the text to find; the mutant would not differ")
    start = data.index(find)
    end = start + len(find)
    result = data[:start] + replace + data[end:]
    if vba:
        check_vba(data, result, replace)
    # The changed lines run from the start of the first to the end of the last.
    first = data.rfind(b"\n", 0, start) + 1
    stop = data.find(b"\n", end)
    tail = b"" if find.endswith(b"\n") else data[end:] if stop < 0 else data[end:stop]
    return Mutant(result, data.count(b"\n", 0, start) + 1,
                  data[first:end] + tail, data[first:start] + replace + tail)


def check_vba(data: bytes, result: bytes, replace: bytes) -> None:
    names = VB_NAME.findall(data)
    if not names:
        raise Refusal("the VBA component has no Attribute VB_Name line; export it from the VBE")
    if VB_NAME.findall(result) != names:
        raise Refusal("the replacement would change, remove or add an Attribute VB_Name line")
    if not replace.isascii():
        raise Refusal("the replacement is not ASCII; exported VBA source is ASCII only")


def inside(path: Path, folder: Path) -> bool:
    return path == folder or folder in path.parents


def check_output(root: Path, output: Path, target: Path) -> None:
    """Refuse an output folder where the mutant could be committed or replace a source."""
    if not inside(output, root):
        raise Refusal(f"the output folder {output} is outside the repository {root}")
    folder = output.relative_to(root)
    for name in SOURCE_DIRECTORIES:
        if inside(output, root / name):
            raise Refusal(f"the output folder {folder.as_posix()} is in the source "
                          f"directory {name}/")
    try:
        tracked = gitio.tracked(root, folder.as_posix())
    except GitError as error:
        raise Refusal(f"Git cannot list the files of {root}: {error.stderr.strip()}") from error
    if tracked:
        raise Refusal(f"the output folder {folder.as_posix()} holds tracked files")
    relative = target.relative_to(root).as_posix()
    try:
        gitio.git(root, "check-ignore", "-q", "--", relative)
    except GitError as error:
        if error.returncode == 1:
            raise Refusal(f"Git does not ignore {relative}; choose an ignored output "
                          "folder") from error
        raise Refusal(f"Git cannot check whether {relative} is ignored: "
                      f"{error.stderr.strip()}") from error


def check_destination(output: Path, path: Path, source: Path) -> None:
    """Refuse a destination that a symbolic link could redirect, or that is the source."""
    current = output
    for part in path.relative_to(output).parts:
        current = current / part
        if current.is_symlink():
            raise Refusal(f"{current} is a symbolic link; delete it, since writing "
                          f"through it could change another file")
    if path.exists() and not path.is_file():
        raise Refusal(f"{path} exists and is not a file")
    resolved = path.resolve()
    if resolved == source or not inside(resolved, output):
        raise Refusal(f"the mutant path {path} resolves to {resolved}, outside {output}")


def write_new(path: Path, data: bytes) -> None:
    """Write ``data`` to a new file and move it to ``path``.

    The file is created fresh in ``path``'s folder and then renamed over
    ``path``, which replaces the directory entry. An existing file at ``path``
    is never opened or truncated, so a hard link left there to the source or
    to any other file leaves that file unchanged.
    """
    descriptor, temporary = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp",
                                             dir=path.parent)
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(data)
        os.replace(temporary, path)
    except BaseException:
        Path(temporary).unlink(missing_ok=True)
        raise


def read_text(value: str | None, path: Path | None) -> bytes:
    if path is not None:
        return path.read_bytes()
    assert value is not None
    return value.encode("utf-8")


def show(text: bytes, sign: str) -> list[str]:
    return [f"{sign} {line}" for line in text.decode("utf-8", "replace").splitlines()]


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def write(root: Path, source: Path, output: Path, find: bytes, replace: bytes) -> list[str]:
    """Write the mutant and return the report; raise Refusal before writing anything."""
    if not source.is_file():
        raise Refusal(f"{source} is not a file")
    if not inside(source, root):
        raise Refusal(f"{source} is outside the repository {root}")
    relative = source.relative_to(root)
    target = output / relative
    if target == source:
        raise Refusal(f"the mutant would overwrite {relative.as_posix()}")
    check_destination(output, target, source)
    check_output(root, output, target)
    data = source.read_bytes()
    vba = source.suffix.lower() in VBA_SUFFIXES
    mutant = mutate(data, find, replace, vba)
    companion = source.with_suffix(".frx") if source.suffix.lower() == ".frm" else None
    extra = companion if companion is not None and companion.is_file() else None
    if extra is not None:
        check_destination(output, target.with_suffix(extra.suffix), extra.resolve())
    target.parent.mkdir(parents=True, exist_ok=True)
    write_new(target, mutant.data)
    report = [f"mutant: {target.relative_to(root).as_posix()}",
              f"changed at {relative.as_posix()} line {mutant.line}:",
              *show(mutant.before, "-"), *show(mutant.after, "+")]
    if extra is not None:
        copy = target.with_suffix(extra.suffix)
        write_new(copy, extra.read_bytes())
        report.append(f"copied {extra.relative_to(root).as_posix()} unchanged beside it")
    report += [f"source SHA-256 {digest(data)}",
               f"mutant SHA-256 {digest(mutant.data)}"]
    return report


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Write a copy of FILE with one occurrence of a text replaced, "
                    "under an ignored output folder.")
    parser.add_argument("file", type=Path, metavar="FILE", help="the source file to mutate")
    finding = parser.add_mutually_exclusive_group(required=True)
    finding.add_argument("--find", metavar="TEXT", help="the literal text to replace")
    finding.add_argument("--find-file", type=Path, metavar="PATH",
                         help="a file whose whole content is the text to replace")
    replacing = parser.add_mutually_exclusive_group(required=True)
    replacing.add_argument("--replace", metavar="TEXT", help="the literal replacement")
    replacing.add_argument("--replace-file", type=Path, metavar="PATH",
                           help="a file whose whole content is the replacement")
    parser.add_argument("--output", type=Path, metavar="DIR",
                        help="the ignored output folder (default: test-results/mutants/ "
                             "of the repository)")
    parser.add_argument("--root", type=Path, default=ROOT,
                        help="the repository root (default: the checkout that holds this tool)")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    output = (args.output if args.output is not None else root / DEFAULT_OUTPUT).resolve()
    if args.find == "":
        parser.error("the text to find is empty")
    try:
        find = read_text(args.find, args.find_file)
        replace = read_text(args.replace, args.replace_file)
        if not find:
            raise Refusal(f"{args.find_file} is empty; there is no text to find")
        report = write(root, args.file.resolve(), output, find, replace)
    except (Refusal, OSError) as exc:
        print(f"make_mutant: refused: {exc}", file=sys.stderr)
        return 2
    print("\n".join(report))
    return 0


if __name__ == "__main__":
    sys.exit(main())
