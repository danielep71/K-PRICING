#!/usr/bin/env python3
"""Reject tracked binary files outside the approved paths and the allowlist.

A file is binary when .gitattributes unsets its text attribute (the
office-doc, opaque and vba-form-bin rules, the byte-preserved regulatory
sources and the variant files of file-based test cases), or when a file Git
would treat as text contains a NUL byte anywhere, also past the first 8000
bytes, which are all that Git inspects. A binary file passes only if it is:

- under docs/methodology/sources/, the byte-preserved regulatory sources;
- in a variant directory of a file-based test case,
  tests/fixtures/<case-id>/<variant>/, where tools/check_test_cases.py admits
  only the files the data contract names;
- src/workbook/Template.xlsx, the macro-free workbook template;
- a UserForm .frx with its .frm beside it; or
- listed in .github/binary-allowlist.json with its current SHA-256.

.gitignore keeps such files out by default; this check catches one added on
purpose, for example with git add -f.

An Office Open XML package that passes (a .xlsx, .xltx, .docx or similar
file, as a ZIP with [Content_Types].xml) also fails when it carries macros, a
vbaProject.bin part or a reference to one, or any finding of
tools/check_workbook_privacy.py: a OneDrive or SharePoint path or URL, an
absolute path under a user profile, a person or organization in its
metadata, comments or revisions, a Purview sensitivity label or SharePoint
metadata. Each part is parsed as XML in its declared encoding, so a UTF-16
part or a character reference hides nothing, and an XML part that cannot be
read fails too. The byte-preserved regulatory sources under
docs/methodology/sources/ are exempt: they stay as published. It never
changes a file.
"""

from __future__ import annotations

import argparse
import fnmatch
import hashlib
import json
import re
import sys
from collections.abc import Sequence
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

sys.path.insert(0, str(Path(__file__).resolve().parent))

import check_workbook_privacy as privacy  # noqa: E402
import gitio  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
ALLOWLIST = Path(".github/binary-allowlist.json")
# fnmatch's * also matches "/": the second pattern needs a variant directory.
APPROVED_PATHS = ("docs/methodology/sources/*", "tests/fixtures/*/*/*",
                  "src/workbook/Template.xlsx")
# Read in chunks: Git looks at the first 8000 bytes only, so a NUL byte later in
# the file would otherwise stay unseen.
CHUNK_BYTES = 1 << 16
# Office packages are inspected for macros and private data, except the
# byte-preserved regulatory sources.
SOURCE_COPIES = "docs/methodology/sources/"
# A VBA part, its content type or relationship, or a macro-enabled main part.
MACRO_PART = re.compile(r"vbaProject|macroEnabled", re.I)


def text_attributes(root: Path, names: Sequence[str]) -> dict[str, str]:
    """The value of the text attribute for each path: set, unset, auto or unspecified."""
    fields = gitio.text(root, "check-attr", "-z", "--stdin", "text",
                        input="\0".join(names) + "\0").split("\0")
    return {fields[index]: fields[index + 2] for index in range(0, len(fields) - 2, 3)}


def has_nul(path: Path) -> bool:
    """Whether the file holds a NUL byte anywhere."""
    with path.open("rb") as handle:
        while chunk := handle.read(CHUNK_BYTES):
            if b"\0" in chunk:
                return True
    return False


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load_allowlist(root: Path, findings: list[str]) -> dict[str, str]:
    """Map each allowlisted path to its SHA-256; report a malformed file."""
    path = root / ALLOWLIST
    if not path.is_file():
        return {}
    try:
        data: Any = json.loads(path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, ValueError) as error:
        findings.append(f"{ALLOWLIST.as_posix()}: not valid UTF-8 JSON ({error})")
        return {}
    files = data.get("files") if isinstance(data, dict) else None
    if not isinstance(data, dict) or data.get("schema_version") != 1 or not isinstance(files, dict):
        findings.append(f"{ALLOWLIST.as_posix()}: needs schema_version 1 and a files object")
        return {}
    allowed: dict[str, str] = {}
    for name, entry in files.items():
        digest = entry.get("sha256") if isinstance(entry, dict) else None
        reason = entry.get("reason") if isinstance(entry, dict) else None
        if not isinstance(digest, str) or len(digest) != 64 or not isinstance(reason, str) \
                or not reason.strip():
            findings.append(f"{ALLOWLIST.as_posix()}: {name} needs a 64-character sha256 "
                            "and a reason")
            continue
        allowed[name] = digest.lower()
    return allowed


def macro_reference(part: ElementTree.Element) -> bool:
    """Whether a name, attribute value or text in the part refers to macros."""
    return any(MACRO_PART.search(text)
               for element in part.iter()
               for text in (element.tag, element.text or "", *element.attrib.values()))


def package_problems(path: Path) -> list[str]:
    """Macros or private data in an Office Open XML package; none for other files.

    The private data are the findings of tools/check_workbook_privacy.py, which
    also reports an XML part that cannot be read.
    """
    try:
        data = path.read_bytes()
        parts = privacy.read_package(data)
    except (OSError, privacy.NotAPackage):
        return []
    problems = []
    # The content types and relationships, which name a VBA part or a macro-enabled type.
    references = [privacy.read_part(content) for name, content in parts.items()
                  if content is not None
                  and (name == privacy.CONTENT_TYPES or name.endswith(".rels"))]
    if any(MACRO_PART.search(name) for name in parts) or \
            any(part is not None and macro_reference(part) for part in references):
        problems.append("carries macros (a vbaProject part, a reference to one or a "
                        "macro-enabled content type); commit "
                        "the macro-free package and keep the code as exported source")
    problems += [f"{finding}; clear it, as DISTRIBUTION.md describes"
                 for finding in privacy.scan(data)]
    return problems


def approved(name: str, tracked: set[str]) -> bool:
    if any(fnmatch.fnmatchcase(name, pattern) for pattern in APPROVED_PATHS):
        return True
    return name.endswith(".frx") and f"{name[:-4]}.frm" in tracked


def check(root: Path) -> tuple[list[str], int]:
    """Return the findings and the number of binary files found."""
    findings: list[str] = []
    names = gitio.tracked(root)
    tracked = set(names)
    allowed = load_allowlist(root, findings)
    attributes = text_attributes(root, names) if names else {}
    binaries = 0
    for name in names:
        path = root / name
        if not path.is_file():
            continue
        if not name.startswith(SOURCE_COPIES):
            findings.extend(f"{name}: {problem}" for problem in package_problems(path))
        # An allowlisted file is checked against its recorded hash whatever its
        # content is now, so a replacement never passes by no longer looking binary.
        if name in allowed:
            binaries += 1
            if sha256(path) != allowed[name]:
                findings.append(f"{name}: content changed since it was allowlisted; review it "
                                f"and update its sha256 in {ALLOWLIST.as_posix()}")
            continue
        declared = attributes.get(name) == "unset"
        if not declared and not has_nul(path):
            continue
        binaries += 1
        if approved(name, tracked):
            continue
        kind = "binary file" if declared else "binary content in a file treated as text"
        findings.append(f"{name}: {kind} outside the approved paths; remove it, or review it "
                        f"and add it to {ALLOWLIST.as_posix()}")
    for name in sorted(set(allowed) - tracked):
        findings.append(f"{ALLOWLIST.as_posix()}: {name} is not a tracked file")
    return findings, binaries


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, binaries = check(args.root.resolve())
    for finding in findings:
        print(f"FAIL: {finding}")
    status = "FAIL" if findings else "PASS"
    print(f"binaries: {status} ({binaries} binary file(s) tracked)")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
