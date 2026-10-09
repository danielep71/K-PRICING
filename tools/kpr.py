#!/usr/bin/env python3
"""Thin portable developer operations; owning validators retain their rules."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import re
import subprocess
import sys
from pathlib import Path
from typing import Any
from zipfile import BadZipFile

from _gatelib import git_text
from check_excel_evidence import source_inventory
from release_provenance import SHA, committed, decode, relative, require
from workbook_snapshot import compare, differences, snapshot

PROFILE = ".github/repository-profile.json"
CHECKS = (
    ("check_repo.py",), ("check_vba_jumps.py",), ("check_vba_conditionals.py",),
    ("check_vba_public_api.py",), ("check_kpr_contract.py",),
    ("gen_fixtures.py", "--check"), ("check_documentation.py",),
    ("check_committed_whitespace.py", "--mode", "working-tree"),
)


def git(root: Path, *arguments: str) -> str:
    result = git_text(root, *arguments)
    require(result.returncode == 0, result.stderr.strip() or "Git operation failed")
    return result.stdout.strip()


def candidate(root: Path, sha: str) -> dict[str, Any]:
    require(SHA.fullmatch(sha), "candidate requires a full lowercase 40-character SHA")
    require(git(root, "rev-parse", "HEAD") == sha, "HEAD differs from candidate")
    require(not git(root, "status", "--porcelain", "--untracked-files=all"),
            "candidate is dirty (tracked changes or nonignored untracked files)")
    return decode(committed(root, sha, PROFILE))


def normalize_source(raw: bytes, suffix: str) -> bytes:
    if suffix == ".frx":
        return raw
    text = raw.decode("ascii")
    require("\r" not in text.replace("\r\n", ""), "bare CR in source export")
    # Only the documented Git LF / VBE CRLF transport distinction is incidental.
    return text.replace("\r\n", "\n").encode("ascii")


def inventory(root: Path, sha: str) -> dict[str, Any]:
    config = candidate(root, sha)
    entries = source_inventory(root, sha, config, include_examples=True)
    identities: set[str] = set()
    for entry in entries:
        path = entry["path"]
        require(relative(path), "unsafe component path")
        raw = committed(root, sha, path)
        suffix = Path(path).suffix
        normalized = normalize_source(raw, suffix)
        entry["normalized_sha256"] = hashlib.sha256(normalized).hexdigest()
        if suffix != ".frx":
            names = re.findall(rb'^Attribute VB_Name = "([A-Za-z][A-Za-z0-9_]*)"$',
                               normalized, re.MULTILINE)
            require(len(names) == 1, f"missing or ambiguous component identity: {path}")
            name = names[0].decode("ascii")
            require(name == Path(path).stem, f"component/filename mismatch: {path}")
            require(name.casefold() not in identities, f"duplicate component identity: {name}")
            identities.add(name.casefold())
            entry["component"] = name
        entry["role"] = config["vba"]["components"].get(path, "resource")
    # Recheck after reading; neither invocation trusts a supplied SHA as import proof.
    candidate(root, sha)
    return {"schema": "kpr-source-inventory", "schema_version": 1, "status": "pass",
            "repository": config["repository"], "candidate_sha": sha, "sources": entries,
            "scope": "all-configured-components-including-examples",
            "import": "NOT_RUN", "compile": "NOT_RUN",
            "tool_environment": {"python": platform.python_version(),
                                 "git": git(root, "--version"), "os": platform.platform()}}


def round_trip(root: Path, sha: str, directory: Path) -> dict[str, Any]:
    report = inventory(root, sha)
    require(directory.is_dir() and not directory.is_symlink(), "exports must be a real directory")
    expected = {Path(entry["path"]).name: entry for entry in report["sources"]}
    require(len(expected) == len(report["sources"]), "ambiguous flat export inventory")
    observed = {path.name: path for path in directory.iterdir()}
    findings: list[dict[str, Any]] = []
    for name in sorted(expected.keys() | observed.keys()):
        if name not in expected or name not in observed:
            findings.append({"path": name, "reason": "extra" if name in observed else "missing"})
            continue
        path = observed[name]
        require(path.is_file() and not path.is_symlink(), f"not a regular export: {name}")
        actual = normalize_source(path.read_bytes(), path.suffix)
        original = normalize_source(committed(root, sha, expected[name]["path"]), path.suffix)
        if actual != original:
            findings.append({"path": name, "reason": "source differs",
                             "actual_sha256": hashlib.sha256(actual).hexdigest(),
                             "differences": differences(original.decode("ascii").splitlines(),
                                                        actual.decode("ascii").splitlines())
                             if path.suffix != ".frx" else []})
    candidate(root, sha)
    return {"schema": "kpr-source-round-trip", "schema_version": 1,
            "status": "fail" if findings else "pass", "candidate_sha": sha,
            "inventory": report, "findings": findings,
            "workbook_alignment": "NOT_OBSERVED",
            "limitation": "Compares supplied exports; cannot authenticate their workbook or session origin."}


def doctor(root: Path, output_path: Path | None) -> dict[str, Any]:
    sha = git(root, "rev-parse", "HEAD")
    dirty = git(root, "status", "--porcelain", "--untracked-files=all")
    output: dict[str, Any] = {"status": "NOT_TESTED"}
    if output_path is not None:
        path = output_path.absolute()
        output = {"path": str(path), "exists": path.exists(), "symlink": path.is_symlink(),
                  "parent_exists": path.parent.is_dir(),
                  "parent_access_hint": os.access(path.parent, os.W_OK),
                  "exclusive_create": "NOT_TESTED", "lock": "UNKNOWN"}
    return {"schema": "kpr-doctor", "schema_version": 1, "status": "observed",
            "source": {"head": sha, "clean": not dirty, "changes": dirty.splitlines()},
            "host": {"os": platform.platform(), "python": platform.python_version(),
                     "excel_automation": "UNAVAILABLE" if sys.platform != "win32" else "NOT_TESTED",
                     "excel_version": None, "excel_build": None, "office_bitness": None,
                     "references": None, "vba_project_access": "UNKNOWN",
                     "workbook_locks": "UNKNOWN", "live_session": "NOT_OBSERVED",
                     "saved_workbook_alignment": "NOT_OBSERVED"},
            "output_path": output, "compile": "NOT_RUN", "runtime": "NOT_RUN"}


def check(root: Path) -> dict[str, Any]:
    results = []
    for name, *arguments in CHECKS:
        result = subprocess.run([sys.executable, str(root / "tools" / name), "--root", str(root),
                                 *arguments], capture_output=True, text=True, check=False)
        results.append({"validator": name, "exit_code": result.returncode,
                        "stdout": result.stdout, "stderr": result.stderr})
    code = 2 if any(item["exit_code"] not in (0, 1) for item in results) else int(
        any(item["exit_code"] == 1 for item in results))
    return {"schema": "kpr-check", "schema_version": 1,
            "status": ("pass", "fail", "error")[code], "validators": results,
            "scope": "Source checks only; not the full hosted gate or Excel certification."}


def parser() -> argparse.ArgumentParser:
    cli = argparse.ArgumentParser(description=__doc__)
    cli.add_argument("--root", type=Path, default=Path.cwd())
    cli.add_argument("--output", type=Path, help="new JSON report path; never overwrite an existing file")
    commands = cli.add_subparsers(dest="command", required=True)
    commands.add_parser("check", help="coordinate existing portable source validators")
    diagnostic = commands.add_parser("doctor", help="read-only source and host availability observations")
    diagnostic.add_argument("--output-path", type=Path, help="inspect a prospective output without writing it")
    for name in ("inventory", "round-trip"):
        operation = commands.add_parser(name, help="exact-candidate inventory or supplied export comparison")
        operation.add_argument("--candidate-sha", required=True)
        if name == "round-trip":
            operation.add_argument("--exports", type=Path, required=True)
    snap = commands.add_parser("snapshot", help="saved OOXML content; calculation is not observed")
    snap.add_argument("workbook", type=Path)
    diff = commands.add_parser("compare", help="compare saved workbooks, separating structure and caches")
    diff.add_argument("left", type=Path)
    diff.add_argument("right", type=Path)
    return cli


def dispatch(args: argparse.Namespace) -> dict[str, Any]:
    if args.command == "check":
        return check(args.root)
    if args.command == "doctor":
        return doctor(args.root, args.output_path)
    if args.command == "inventory":
        return inventory(args.root, args.candidate_sha)
    if args.command == "round-trip":
        return round_trip(args.root, args.candidate_sha, args.exports)
    if args.command == "snapshot":
        return snapshot(args.workbook)
    return compare(args.left, args.right)


def main(arguments: list[str] | None = None) -> int:
    args = parser().parse_args(arguments)
    try:
        report = dispatch(args)
        report["tool"] = {"entry_point": "tools/kpr.py",
                          "sha256": hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
                          "snapshot_sha256": hashlib.sha256(
                              Path(__file__).with_name("workbook_snapshot.py").read_bytes()).hexdigest(),
                          "python": platform.python_version()}
        payload = json.dumps(report, indent=2, sort_keys=True) + "\n"
        if args.output is not None:
            # No overwrite option: a locked/existing output cannot destroy earlier evidence.
            with args.output.open("x", encoding="utf-8", newline="\n") as stream:
                stream.write(payload)
        print(payload, end="")
        return {"fail": 1, "error": 2}.get(report.get("status", "pass"), 0)
    except (OSError, ValueError, KeyError, BadZipFile, subprocess.SubprocessError) as error:
        print(json.dumps({"schema": "kpr-operation-error", "schema_version": 1,
                          "status": "error", "detail": str(error)}))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
