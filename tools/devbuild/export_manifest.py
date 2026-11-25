#!/usr/bin/env python3
"""Record which workbook an export of VBA components came from.

Usage: python tools/devbuild/export_manifest.py WORKBOOK EVIDENCE_FOLDER

Run it right after exporting the components into EVIDENCE_FOLDER/exported/,
before the workbook is opened or saved again. It writes
EVIDENCE_FOLDER/export-manifest.json with the SHA-256 of the workbook and of
every exported file. Both names come from the evidence section of
.github/excel-evidence-policy.json, the policy that tools/check_excel_run.py
reads, so the producer and the checker agree; --policy names another copy. tools/check_excel_run.py --workbook then binds the
export to that workbook, and the workbook's build stamp to the commit. A
manifest is a record, not a signature: it binds an export to a workbook,
but it does not attest who made either. Standard library only.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from collections.abc import Sequence
from pathlib import Path

FORMAT = "vba-export-1"
POLICY = Path(__file__).resolve().parents[2] / ".github/excel-evidence-policy.json"


def layout(policy: Path = POLICY) -> tuple[str, str]:
    """The export folder and the manifest file that the evidence policy names."""
    try:
        evidence = json.loads(policy.read_text(encoding="utf-8"))["evidence"]
        names = (evidence["exported"], evidence["export_manifest"])
    except (OSError, ValueError, KeyError, TypeError) as error:
        raise ValueError(f"{policy}: no evidence layout ({error!r})") from error
    for name in names:
        if not isinstance(name, str) or not name or "/" in name or "\\" in name \
                or name in (".", ".."):
            raise ValueError(f"{policy}: {name!r} is not a plain file name")
    return names


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def manifest(workbook: Path, folder: Path, policy: Path = POLICY) -> dict[str, object]:
    exported = folder / layout(policy)[0]
    files = sorted(path for path in exported.iterdir() if path.is_file()) \
        if exported.is_dir() else []
    if not files:
        raise ValueError(f"{exported} holds no exported file")
    return {"format": FORMAT,
            "workbook": {"file": workbook.name, "sha256": sha256(workbook)},
            "files": [{"file": path.name, "sha256": sha256(path)} for path in files]}


def write(workbook: Path, folder: Path, policy: Path = POLICY) -> Path:
    target = folder / layout(policy)[1]
    record = manifest(workbook, folder, policy)
    target.write_text(json.dumps(record, indent=2) + "\n", encoding="utf-8")
    return target


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("workbook", type=Path, help="the workbook the components were exported from")
    parser.add_argument("evidence", type=Path, help="the evidence folder holding exported/")
    parser.add_argument("--policy", type=Path, default=POLICY,
                        help="the evidence policy that names the folder and the manifest")
    args = parser.parse_args(argv)
    if not args.workbook.is_file():
        print(f"FAIL: {args.workbook}: not a file")
        return 1
    try:
        target = write(args.workbook, args.evidence, args.policy)
    except (OSError, ValueError) as error:
        print(f"FAIL: {error}")
        return 1
    print(f"wrote {target}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
