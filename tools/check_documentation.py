#!/usr/bin/env python3
"""Offline documentation contracts; parses commands without executing them."""

from __future__ import annotations

import argparse
import ast
import json
import re
import shlex
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from _gatelib import run_gate, tracked_files
from release_provenance import decode, nonempty, relative, require

POLICY = ".github/documentation-policy.json"


def load_policy(root: Path) -> dict[str, Any]:
    policy: dict[str, Any] = decode((root / POLICY).read_bytes())
    require(isinstance(policy, dict) and policy.get("schema_version") == 1,
            "unsupported documentation policy")
    history = policy.get("historical_documents")
    require(isinstance(history, dict) and all(relative(p) and nonempty(reason) for p, reason in history.items()),
            "historical exclusions require safe paths and reasons")
    require(isinstance(policy.get("references"), list), "references must be an array")
    return policy


# Options that take no value. Every other option consumes the next token.
NO_VALUE_ACTIONS = frozenset({"store_true", "store_false", "store_const", "append_const",
                              "count", "help", "version"})
HELP_OPTIONS = {"--help": False, "-h": False}
# Options of unittest.main, for test modules run as scripts.
UNITTEST_OPTIONS = {"-v": False, "--verbose": False, "-q": False, "--quiet": False,
                    "--locals": False, "-f": False, "--failfast": False, "-c": False,
                    "--catch": False, "-b": False, "--buffer": False, "-k": True,
                    "--durations": True}
SHELL_OPERATORS = frozenset({"|", "||", "&", "&&", ";", ">", ">>", "<", "(", ")"})
NUMBER = re.compile(r"-\d")
PLACEHOLDER = re.compile(r"<[^<>\s`]+>")
COMMAND = re.compile(
    r"(?<![\w.-])(?:python3?|py)(?:\s+-3(?:\.\d+)?)?\s+"
    r"(?:(tools/[A-Za-z0-9_./-]+\.py)|-m\s+(tools\.[A-Za-z0-9_.]+))([^\n`]*)")


@dataclass
class Cli:
    """The command line a tools/ script declares, read from its AST without importing it."""

    options: dict[str, bool] = field(default_factory=lambda: dict(HELP_OPTIONS))
    subcommands: dict[str, dict[str, bool]] | None = None


@dataclass
class _Scope:
    """Loop variables resolved to the literal names they iterate over."""

    loops: dict[str, list[str]] = field(default_factory=dict)


class _CliReader:
    def __init__(self, path: str, tree: ast.Module) -> None:
        self.path = path
        self.cli = Cli()
        self.unittest = False
        self.literals: dict[str, list[str]] = {}
        # Parser variable -> (loop variable it was named by, the names it can hold).
        self.parsers: dict[str, tuple[str | None, list[str]]] = {}
        for node in ast.walk(tree):
            if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
                names = self.literal_names(node.value)
                if names is not None:
                    self.literals[node.targets[0].id] = names
        self.statements(tree.body, _Scope())

    @staticmethod
    def literal_names(node: ast.expr) -> list[str] | None:
        items = node.keys if isinstance(node, ast.Dict) else (
            node.elts if isinstance(node, (ast.List, ast.Tuple, ast.Set)) else None)
        names = [item.value for item in items or []
                 if isinstance(item, ast.Constant) and isinstance(item.value, str)]
        return names if items is not None and len(names) == len(items) else None

    def iterated(self, node: ast.expr) -> list[str] | None:
        if (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
                and node.func.attr in ("items", "keys") and not node.args):
            node = node.func.value
        if isinstance(node, ast.Name):
            return self.literals.get(node.id)
        return self.literal_names(node)

    def statements(self, body: list[ast.stmt], scope: _Scope) -> None:
        for statement in body:
            if isinstance(statement, ast.For):
                self.loop(statement, scope)
            elif isinstance(statement, ast.If):
                self.branch(statement, scope)
            else:
                self.simple(statement, scope)

    def loop(self, statement: ast.For, scope: _Scope) -> None:
        target = statement.target
        if isinstance(target, ast.Tuple) and target.elts:
            target = target.elts[0]
        names = self.iterated(statement.iter)
        inner = _Scope(dict(scope.loops))
        if isinstance(target, ast.Name) and names is not None:
            inner.loops[target.id] = names
        self.calls(statement.iter, scope)
        self.statements(statement.body, inner)
        self.statements(statement.orelse, scope)

    def branch(self, statement: ast.If, scope: _Scope) -> None:
        test = statement.test
        then, other = scope, scope
        if (isinstance(test, ast.Compare) and isinstance(test.left, ast.Name) and len(test.ops) == 1
                and isinstance(test.ops[0], ast.Eq) and isinstance(test.comparators[0], ast.Constant)
                and test.left.id in scope.loops):
            names, value = scope.loops[test.left.id], test.comparators[0].value
            then = _Scope({**scope.loops, test.left.id: [name for name in names if name == value]})
            other = _Scope({**scope.loops, test.left.id: [name for name in names if name != value]})
        self.calls(test, scope)
        self.statements(statement.body, then)
        self.statements(statement.orelse, other)

    def simple(self, statement: ast.stmt, scope: _Scope) -> None:
        nested = [getattr(statement, name, None) for name in ("body", "orelse", "finalbody")]
        if any(isinstance(block, list) for block in nested):
            for block in nested:
                if isinstance(block, list):
                    self.statements(block, scope)
            for handler in getattr(statement, "handlers", []):
                self.statements(handler.body, scope)
            return
        if (isinstance(statement, ast.Assign) and len(statement.targets) == 1
                and isinstance(statement.targets[0], ast.Name) and self.is_call(statement.value, "add_parser")):
            assert isinstance(statement.value, ast.Call)
            self.parsers[statement.targets[0].id] = self.add_parser(statement.value, scope)
        self.calls(statement, scope)

    @staticmethod
    def is_call(node: ast.AST, attribute: str) -> bool:
        return (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
                and node.func.attr == attribute)

    def add_parser(self, call: ast.Call, scope: _Scope) -> tuple[str | None, list[str]]:
        name = call.args[0] if call.args else None
        if isinstance(name, ast.Constant) and isinstance(name.value, str):
            result: tuple[str | None, list[str]] = (None, [name.value])
        elif isinstance(name, ast.Name) and name.id in scope.loops:
            result = (name.id, list(scope.loops[name.id]))
        else:
            raise ValueError(f"{self.path}:{call.lineno}: subcommand name is not readable statically")
        assert self.cli.subcommands is not None
        for subcommand in result[1]:
            self.cli.subcommands.setdefault(subcommand, dict(HELP_OPTIONS))
        return result

    def calls(self, node: ast.AST, scope: _Scope) -> None:
        for call in ast.walk(node):
            if not isinstance(call, ast.Call) or not isinstance(call.func, ast.Attribute):
                continue
            if call.func.attr == "add_subparsers" and self.cli.subcommands is None:
                self.cli.subcommands = {}
            elif call.func.attr == "add_parser" and self.cli.subcommands is not None:
                self.add_parser(call, scope)
            elif call.func.attr == "add_argument":
                self.add_argument(call, scope)
            elif (call.func.attr == "main" and isinstance(call.func.value, ast.Name)
                  and call.func.value.id == "unittest"):
                self.unittest = True

    def add_argument(self, call: ast.Call, scope: _Scope) -> None:
        flags = [arg.value for arg in call.args if isinstance(arg, ast.Constant)
                 and isinstance(arg.value, str) and arg.value.startswith("-")]
        action = next((keyword.value.value for keyword in call.keywords if keyword.arg == "action"
                       and isinstance(keyword.value, ast.Constant)), None)
        takes_value = action not in NO_VALUE_ACTIONS
        receiver = call.func.value if isinstance(call.func, ast.Attribute) else None
        targets = [self.cli.options]
        if isinstance(receiver, ast.Name) and receiver.id in self.parsers and self.cli.subcommands is not None:
            loop, names = self.parsers[receiver.id]
            if loop is not None and loop in scope.loops:
                names = scope.loops[loop]
            targets = [self.cli.subcommands[name] for name in names]
        for options in targets:
            options.update(dict.fromkeys(flags, takes_value))


def command_cli(root: Path, path: str, seen: set[str] | None = None) -> Cli:
    """Read a script's parser statically; a parser that cannot be read raises ValueError."""
    seen = set() if seen is None else seen
    if path in seen:
        return Cli()
    seen.add(path)
    tree = ast.parse((root / path).read_text(encoding="utf-8"))
    reader = _CliReader(path, tree)
    cli = reader.cli
    if reader.unittest:
        cli.options.update(UNITTEST_OPTIONS)
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and (node.module == "policy_coverage_runner" or
                (node.module == "_gatelib" and any(item.name == "parse_report_args" for item in node.names))):
            cli.options.update(command_cli(root, f"tools/{node.module}.py", seen).options)
    return cli


def command_flags(root: Path, path: str) -> set[str]:
    """The long options a script accepts before any subcommand."""
    return {flag for flag in command_cli(root, path).options if flag.startswith("--")}


def command_tokens(rest: str) -> list[str]:
    """The arguments of one documented command, up to a comment or shell operator."""
    lexer = shlex.shlex(PLACEHOLDER.sub("PLACEHOLDER", rest), posix=True, punctuation_chars=True)
    lexer.whitespace_split = True
    try:
        tokens = list(lexer)
    except ValueError:
        tokens = rest.split()
    for index, token in enumerate(tokens):
        if token in SHELL_OPERATORS or token.startswith(("|", ">", "&", ";")) or token.startswith("$("):
            return tokens[:index]
    return tokens


def commands(text: str) -> list[tuple[str, list[str]]]:
    # Literal Python invocations, in fenced examples or inline code: python,
    # python3, py and py -3 launchers, a tools/ script or a -m tools. module.
    # Shell variables, pipes, substitutions and the command itself are never
    # executed.
    text = re.sub(r"\\\r?\n\s*", " ", text)
    result = []
    for match in COMMAND.finditer(text):
        path = match[1] or match[2].replace(".", "/") + ".py"
        rest = match[3]
        # A command quoted as an argument of another command ends at its quote.
        quote = text[match.start() - 1] if match.start() > 0 else ""
        if quote in ("'", '"'):
            rest = rest.split(quote, 1)[0]
        result.append((path, command_tokens(rest)))
    return result


def command_problems(cli: Cli, tokens: list[str]) -> list[str]:
    """Options the parser does not declare, and subcommands it does not have."""
    problems = []
    scope, chosen = cli.options, False
    index = 0
    while index < len(tokens):
        token = tokens[index]
        index += 1
        if token == "--":
            break
        if token.startswith("-") and len(token) > 1 and not NUMBER.match(token):
            flag = token.split("=", 1)[0]
            if flag not in scope:
                problems.append(f"undocumented CLI option {flag}")
            elif scope[flag] and "=" not in token:
                index += 1
        elif cli.subcommands and not chosen:
            if token not in cli.subcommands:
                problems.append(f"unknown subcommand {token}")
                break
            scope, chosen = cli.subcommands[token], True
    return problems


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
    parsers: dict[str, Cli] = {}
    try:
        policy = load_policy(root)
        files = tracked_files(root)
        for document in sorted(files):
            if not document.endswith(".md") or document in policy["historical_documents"]:
                continue
            for path, tokens in commands((root / document).read_text(encoding="utf-8")):
                count += 1
                if path not in files:
                    findings.append(f"{document}: documented command missing: {path}")
                    continue
                try:
                    if path not in parsers:
                        parsers[path] = command_cli(root, path)
                    problems = command_problems(parsers[path], tokens)
                except (ValueError, SyntaxError) as error:
                    problems = [f"parser not readable: {error}"]
                findings.extend(f"{document}: {path} {problem}" for problem in dict.fromkeys(problems))
        for reference in policy["references"]:
            try:
                reference_check(root, files, reference)
            except (ValueError, KeyError, TypeError, OSError) as error:
                findings.append(str(error))
    except (ValueError, OSError, SyntaxError, RuntimeError) as error:
        findings.append(str(error))
    return {"status": "fail" if findings else "pass", "commands": count, "findings": sorted(findings),
            "scope_note": "Checks literal Python commands (options and subcommands, read statically) and registered file/workflow/policy references; does not execute commands or infer every prose claim."}


def markdown(report: dict[str, Any]) -> str:
    scope_note: str = report["scope_note"]
    return (f"# Documentation drift\n\nResult: {report['status'].upper()}; commands checked: {report['commands']}\n\n"
            + "\n".join(f"- {item}" for item in report["findings"]) + "\n\n" + scope_note + "\n")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    args = parser.parse_args()
    return run_gate(args, build=lambda: build_report(args.root), markdown=markdown, errors=(OSError,))


if __name__ == "__main__":
    raise SystemExit(main())
