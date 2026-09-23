#!/usr/bin/env python3
"""Prove blocking policy-branch fixture coverage for the repository gates.

The coverage contract is semantic: every production ``finding(...)`` site in the
canonical repository checker must be exercised by at least one deterministic
fixture, or the coverage gate fails. Focused hardening gates and authoritative
workflow validation are delegated to their own self-tests and must also pass.
"""

from __future__ import annotations

import ast
import importlib.util
import json
import re
import sys
from pathlib import Path
from types import ModuleType
from typing import Any, Callable

TOOL_NAME = "Policy branch coverage"
CORE_TOOL = "tools/check_repo.py"


class CoverageError(RuntimeError):
    pass


def load_module(path: Path, name: str) -> ModuleType:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise CoverageError(f"Cannot load Python module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def production_finding_sites(source_path: Path) -> dict[str, dict[str, Any]]:
    source = source_path.read_text(encoding="utf-8")
    tree = ast.parse(source, filename=str(source_path))
    fixture_boundary = None
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name == "_write_fixture":
            fixture_boundary = node.lineno
            break
    if fixture_boundary is None:
        raise CoverageError("Cannot locate the canonical fixture boundary in check_repo.py")
    boundary = fixture_boundary

    sites: dict[str, dict[str, Any]] = {}

    class Visitor(ast.NodeVisitor):
        def __init__(self) -> None:
            self.functions: list[str] = []

        def _visit_function(
            self, node: ast.FunctionDef | ast.AsyncFunctionDef
        ) -> None:
            self.functions.append(node.name)
            self.generic_visit(node)
            self.functions.pop()

        def visit_FunctionDef(self, node: ast.FunctionDef) -> None:
            self._visit_function(node)

        def visit_AsyncFunctionDef(self, node: ast.AsyncFunctionDef) -> None:
            self._visit_function(node)

        def visit_Call(self, node: ast.Call) -> None:
            if (
                node.lineno < boundary
                and isinstance(node.func, ast.Name)
                and node.func.id == "finding"
            ):
                function = self.functions[-1] if self.functions else "<module>"
                key = f"{function}:{node.lineno}"
                expression = ast.get_source_segment(source, node) or "finding(...)"
                sites[key] = {
                    "id": key,
                    "function": function,
                    "line": node.lineno,
                    "expression": re.sub(r"\s+", " ", expression).strip(),
                }
            self.generic_visit(node)

    Visitor().visit(tree)
    return dict(sorted(sites.items(), key=lambda item: (item[1]["line"], item[0])))


def rule_by_id(report: dict[str, Any], rule_id: str) -> dict[str, Any] | None:
    for result in report.get("rules", []):
        if result.get("id") == rule_id:
            return result
    return None


def write_json(module: ModuleType, root: Path, relative: str, document: object) -> None:
    module._write_fixture(root / relative, json.dumps(document, indent=2, ensure_ascii=False) + "\n")


def config_document(root: Path, module: ModuleType) -> dict[str, Any]:
    return json.loads((root / module.CONFIG_PATH).read_text(encoding="utf-8"))


def mutate_config(module: ModuleType, root: Path, mutation: Callable[[dict], None]) -> None:
    document = config_document(root, module)
    mutation(document)
    write_json(module, root, module.CONFIG_PATH, document)


def mutate_labels(module: ModuleType, root: Path, mutation: Callable[[dict], None]) -> None:
    document = json.loads((root / module.LABEL_MANIFEST_PATH).read_text(encoding="utf-8"))
    mutation(document)
    write_json(module, root, module.LABEL_MANIFEST_PATH, document)


def add_force(module: ModuleType, root: Path, relative: str) -> None:
    module._run_git(root, "add", "-f", relative)


def rewrite_vba(module: ModuleType, root: Path, relative: str, text: str) -> None:
    module._write_fixture(root / relative, text, crlf=True)
