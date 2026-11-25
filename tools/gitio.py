"""Running Git for the tools: one way to call it, one error, one file listing.

Every tool that reads Git runs ``git -C <root> ...`` through :func:`git`. A
failure, including Git not being installed, raises :class:`GitError`, which
carries the command, the exit status and Git's standard error; a tool that
reports its own error type catches it at its boundary and converts it.
"""

from __future__ import annotations

import subprocess
from collections.abc import Sequence
from pathlib import Path


class GitError(RuntimeError):
    """A Git command that failed or could not start.

    The message is Git's own standard error, stripped, so that a tool that prints
    the error prints what Git said; ``command``, ``returncode`` (None when Git
    could not start) and the unstripped ``stderr`` are kept for callers that need
    them.
    """

    def __init__(self, command: Sequence[str], returncode: int | None, stderr: str) -> None:
        self.command = list(command)
        self.returncode = returncode
        self.stderr = stderr
        status = "could not start" if returncode is None else f"exited with {returncode}"
        super().__init__(stderr.strip() or f"{' '.join(self.command)} {status}")


def git(root: Path, *args: str, input: bytes | None = None) -> bytes:
    """The standard output of ``git -C root args``, as bytes; GitError on failure."""
    command = ["git", "-C", str(root), *args]
    try:
        completed = subprocess.run(command, input=input, capture_output=True, check=False)
    except OSError as error:
        raise GitError(command, None, str(error)) from error
    if completed.returncode:
        raise GitError(command, completed.returncode,
                       completed.stderr.decode("utf-8", errors="replace"))
    return completed.stdout


def text(root: Path, *args: str, input: str | None = None) -> str:
    """The standard output of ``git -C root args`` as UTF-8 text, with line ends
    translated to LF as Python's text mode does; GitError on failure, and
    UnicodeDecodeError when the output is not UTF-8."""
    data = git(root, *args,
               input=None if input is None else input.encode("utf-8", "surrogateescape"))
    return data.decode("utf-8").replace("\r\n", "\n").replace("\r", "\n")


def tracked(root: Path, *pathspec: str) -> list[str]:
    """The paths Git tracks (staged ones included) that match ``pathspec``, all
    when none is given: sorted, with "/" separators, a byte that is not UTF-8
    kept by surrogateescape. GitError when Git cannot list them."""
    listed = git(root, "ls-files", "-z", "--", *pathspec)
    return sorted(item.decode("utf-8", errors="surrogateescape")
                  for item in listed.split(b"\0") if item)


def tracked_or_scanned(root: Path, pattern: str = "*") -> list[str]:
    """The tracked files that match ``pattern``, or, when Git cannot list them,
    the files found by scanning ``root``, outside any .git folder.

    The documentation checks (doc commands, links and restated counts) read the
    Markdown of whatever tree they are given, including one that is not a Git
    checkout, such as an unpacked source archive: there no file is tracked, so
    every file present is read instead. In a checkout only tracked files count,
    so that ignored folders such as a virtual environment are never read. Both
    lists are sorted paths with "/" separators.
    """
    try:
        return tracked(root, pattern)
    except GitError:
        return sorted(path.relative_to(root).as_posix() for path in root.rglob(pattern)
                      if path.is_file() and ".git" not in path.relative_to(root).parts)


def text_or_none(root: Path, *args: str) -> str | None:
    """The output of :func:`text`, or None when Git fails or is not installed:
    for a check that reads what Git knows when it can, and says so otherwise."""
    try:
        return text(root, *args)
    except GitError:
        return None
