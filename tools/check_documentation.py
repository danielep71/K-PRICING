#!/usr/bin/env python3
"""Offline documentation contracts; parses commands without executing them."""

from __future__ import annotations

import argparse
import ast
import json
import re
from pathlib import Path
from typing import Any

from _gatelib import run_gate, tracked_files
from release_provenance import decode, nonempty, relative, require

POLICY = ".github/documentation-policy.json"


def load_policy(root: Path) -> dict[str, Any]:
    policy = decode((root / POLICY).read_bytes())
    require(isinstance(policy, dict) and policy.get("schema_version") == 1,
            "unsupported documentation policy")
    history = policy.get("historical_documents")
    require(isinstance(history, dict) and all(relative(p) and nonempty(reason) for p, reason in history.items()),
            "historical exclusions require safe paths and reasons")
    require(isinstance(policy.get("references"), list), "references must be an array")
    return policy


def command_flags(root: Path, path: str, seen: set[str] | None = None) -> set[str]:
    seen = set() if seen is None else seen
    if path in seen:
        return set()
    seen.add(path)
    tree = ast.parse((root / path).read_text(encoding="utf-8"))
    flags = {"--help"}
    for node in ast.walk(tree):
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute) and node.func.attr == "add_argument":
            flags.update(arg.value for arg in node.args if isinstance(arg, ast.Constant)
                         and isinstance(arg.value, str) and arg.value.startswith("--"))
        if isinstance(node, ast.ImportFrom) and (node.module == "policy_coverage_runner" or
                (node.module == "_gatelib" and any(item.name == "parse_report_args" for item in node.names))):
            flags.update(command_flags(root, f"tools/{node.module}.py", seen))
    return flags


def commands(text: str) -> list[tuple[str, set[str]]]:
    # Literal Python invocations, in fenced examples or inline code; shell
    # variables, pipes, substitutions and the command itself are never executed.
    text = re.sub(r"\\\r?\n\s*", " ", text)
    result = []
    for match in re.finditer(r"\bpython3?\s+(tools/[A-Za-z0-9_./-]+\.py)([^\n`\"<>]*)", text):
        result.append((match[1], set(re.findall(r"(?<!\S)(--[a-z][a-z0-9-]*)", match[2]))))
    return result


def reference_check(root: Path, files: set[str], record: Any) -> None:
    require(isinstance(record, dict), "reference must be an object")
    document, target, token = record.get("document"), record.get("target"), record.get("token")
    require(relative(document) and relative(target) and nonempty(token), "invalid reference fields")
    require(document in files and target in files, f"missing reference document or target: {document} -> {target}")
    require(token in (root / document).read_text(encoding="utf-8"),
            f"document no longer states registered reference: {document}")
    kind = record.get("kind")
    if kind == "file":
        return
    text = (root / target).read_text(encoding="utf-8")
    if kind in ("workflow-name", "job-name"):
        if kind == "job-name":
            job = record.get("job")
            require(nonempty(job), "job-name reference requires job ID")
            match = re.search(rf"^  {re.escape(job)}:\s*\n((?:^    .*\n|^\s*\n)*)", text, re.MULTILINE)
            require(match is not None, f"workflow job ID changed: {target}")
            assert match is not None
            text = match[1]
            pattern = r"^    name:\s*(.+)$"
        else:
            pattern = r"^name:\s*(.+)$"
        match = re.search(pattern, text, re.MULTILINE)
        require(match is not None and match[1].strip().strip("\"'") == token,
                f"documented workflow/context name differs: {target}")
    elif kind == "json-value":
        pointer = record.get("pointer")
        require(isinstance(pointer, str) and pointer.startswith("/"), "invalid JSON pointer")
        value = json.loads(text)
        for segment in pointer[1:].split("/"):
            value = value[segment.replace("~1", "/").replace("~0", "~")]
        require(value == record.get("value"), f"documented policy value differs: {target}{pointer}")
    else:
        raise ValueError("unknown documentation reference kind")


def build_report(root: Path) -> dict[str, Any]:
    findings = []
    count = 0
    try:
        policy = load_policy(root)
        files = tracked_files(root)
        for document in sorted(files):
            if not document.endswith(".md") or document in policy["historical_documents"]:
                continue
            for path, flags in commands((root / document).read_text(encoding="utf-8")):
                count += 1
                if path not in files:
                    findings.append(f"{document}: documented command missing: {path}")
                else:
                    unsupported = flags - command_flags(root, path)
                    if unsupported:
                        findings.append(f"{document}: {path} has undocumented CLI options: {', '.join(sorted(unsupported))}")
        for reference in policy["references"]:
            try:
                reference_check(root, files, reference)
            except (ValueError, KeyError, TypeError, OSError) as error:
                findings.append(str(error))
    except (ValueError, OSError, SyntaxError, RuntimeError) as error:
        findings.append(str(error))
    return {"status": "fail" if findings else "pass", "commands": count, "findings": sorted(findings),
            "scope_note": "Checks literal Python commands and registered file/workflow/policy references; does not execute commands or infer every prose claim."}


def markdown(report: dict[str, Any]) -> str:
    return (f"# Documentation drift\n\nResult: {report['status'].upper()}; commands checked: {report['commands']}\n\n"
            + "\n".join(f"- {item}" for item in report["findings"]) + "\n\n" + report["scope_note"] + "\n")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    args = parser.parse_args()
    return run_gate(args, build=lambda: build_report(args.root), markdown=markdown, errors=(OSError,))


if __name__ == "__main__":
    raise SystemExit(main())
