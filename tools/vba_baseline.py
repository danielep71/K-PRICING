"""The layout baseline that tools/format_vba.py and tools/check_vba_layout.py share.

A project that imports existing modules adopts the house layout one module at
a time: tools/vba_layout_baseline.json lists, as a sorted, unique list of
root-relative POSIX paths under "files" with "schema_version" 1, the modules
that are not yet fully in layout. Both checks read it. A listed module may stay
out of layout, so neither check fails it; a module out of layout that the file
does not list fails; and a listed module that is fully in layout, in the
statement layout of format_vba.py and in the banners and headers of
check_vba_layout.py, is stale and fails until its entry is removed. A listed
path that is not a tracked module under the source directories fails too.
"""

from __future__ import annotations

import json
from collections.abc import Sequence
from pathlib import Path, PurePosixPath

BASELINE = "tools/vba_layout_baseline.json"
BASELINE_SCHEMA = 1

OK, BASELINED, STALE, FAIL = "ok", "baselined", "stale", "fail"


class BaselineError(ValueError):
    """The baseline file is missing, malformed or lists an invalid path."""


def load_baseline(path: Path, shown_as: str) -> list[str]:
    """The sorted, unique root-relative POSIX paths that a baseline file lists."""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        reason = error.strerror if isinstance(error, OSError) and error.strerror else error
        raise BaselineError(f"cannot read baseline {shown_as}: {reason}") from error
    if not isinstance(data, dict) or data.get("schema_version") != BASELINE_SCHEMA:
        raise BaselineError(f"{shown_as}: expected an object with schema_version "
                            f"{BASELINE_SCHEMA}")
    files = data.get("files")
    if not isinstance(files, list) or not all(isinstance(name, str) for name in files):
        raise BaselineError(f"{shown_as}: 'files' must be a list of paths")
    for name in files:
        parts = PurePosixPath(name).parts
        if (not name or name.startswith("/") or "\\" in name or ":" in name
                or ".." in parts or "." in name.split("/") or name.endswith("/")):
            raise BaselineError(f"{shown_as}: {name!r} is not a root-relative POSIX path")
    if files != sorted(set(files)):
        raise BaselineError(f"{shown_as}: 'files' must be sorted and unique")
    return files


def shown(path: Path, root: Path) -> str:
    """A path relative to the root when it is inside it, else absolute."""
    resolved = path.resolve()
    if resolved.is_relative_to(root):
        return resolved.relative_to(root).as_posix()
    return resolved.as_posix()


def read_baseline(root: Path, explicit: Path | None, disabled: bool) -> tuple[list[str], str]:
    """The listed paths and the baseline's name as shown in messages.

    The default baseline is optional; one named with --baseline must exist.
    ``disabled`` is --no-baseline: nothing is read, not even a malformed file.
    """
    path = explicit or Path(BASELINE)
    path = path if path.is_absolute() else root / path
    name = shown(path, root)
    if disabled or not (explicit or path.exists()):
        return [], name
    return load_baseline(path, name), name


def judge(name: str, listed: set[str], own_in_layout: bool, fully_in_layout: bool) -> str:
    """What a module's layout means for a check.

    ``own_in_layout`` is the verdict of the check that asks; ``fully_in_layout``
    is the verdict of both layout checks together.
    """
    if name in listed:
        return STALE if fully_in_layout else BASELINED
    return OK if own_in_layout else FAIL


def stale_message(baseline_name: str) -> str:
    return f"in layout but still listed in {baseline_name}; remove its entry"


def baselined_message(baseline_name: str) -> str:
    return f"not in house-style layout, listed in {baseline_name}"


def unknown_entries(baseline: Sequence[str], tracked: Sequence[str]) -> list[str]:
    """The listed paths that are not tracked modules under the source directories."""
    known = set(tracked)
    return [name for name in baseline if name not in known]


def unknown_message(baseline_name: str) -> str:
    return (f"listed in {baseline_name} but not a tracked VBA module under the source "
            "directories")
