#!/usr/bin/env python3
"""Check that the commands cited in the Markdown guides exist in the tools.

In each tracked Markdown file it reads the inline code spans and the fenced
blocks marked as a shell (``sh``, ``bash``, ``powershell`` and the like) or
not marked at all; fenced blocks of another language are skipped. There it
finds each command that runs a script under tools/ (``python tools/X.py``,
``python3``, ``py`` or a path to an interpreter, with interpreter options
such as ``-I``) and each ``python -m MODULE`` whose module is one of this
repository's. A line ending in ``\\`` or a PowerShell backtick continues on
the next line. It reports:

- a script or module that is not a tracked file;
- an option, long or short, that the script's argparse parser does not
  define, written exactly: an abbreviation argparse would accept is reported
  too, because adding an option can make it ambiguous;
- a subcommand that the parser does not define, and an option after it that
  the subcommand does not define;
- a script whose parser cannot be read, with the reason. The command is then
  not checked, so it fails rather than passes.

It reads each parser from the script's syntax tree, from its entry point
(the ``if __name__ == "__main__":`` block) through the functions of the
script it calls: ``ArgumentParser``, ``add_argument``, argument groups,
``add_subparsers`` and ``add_parser``. It never imports or runs a script. A
parser it cannot follow is reported, not guessed: names that are not string
literals, ``parents=``, ``prefix_chars=``, a parser handed to a function
defined elsewhere, or more than one parser that parses arguments. A script
that builds no parser accepts only arguments that are not options.

Values, positional arguments and placeholders such as ``<folder>``,
``X.Y.Z`` or ``OWNER/REPOSITORY`` are not checked. An option's name is
always checked, even when its value is attached as a placeholder
(``--name=<value>``, ``-x$VALUE``); only a token whose name is itself a
placeholder, such as ``--<option>`` or ``<options>``, is skipped. A
placeholder where a subcommand belongs ends the check of that command. A
shell operator (``|``, ``&&``, ``;``, a redirection), with or without
spaces around it, or a ``#`` comment ends a command. The dated sections of CHANGELOG.md are skipped: they are
not rewritten once released. A tool that the setup guide marks **Remove
once run** is checked while it is tracked; once it is deleted, the commands
that cite it are skipped, because the guides keep recording how the
repository was set up.
"""

from __future__ import annotations

import argparse
import ast
import re
import sys
from collections.abc import Iterator, Sequence
from dataclasses import dataclass, field
from pathlib import Path

from gitio import tracked_or_scanned

ROOT = Path(__file__).resolve().parents[1]
FENCE = re.compile(r"\s*(`{3,}|~{3,})\s*([^\s`]*)")
SHELL_FENCES = frozenset({"", "sh", "bash", "shell", "zsh", "console", "terminal",
                          "shell-session", "powershell", "pwsh", "ps1", "ps", "cmd", "bat",
                          "batch"})
BLOCK_START = re.compile(r"\s*$|\s*(#|>|\||[-*+]\s|\d+[.)]\s)")
CODE_SPAN = re.compile(r"(?<!`)(`+)(?!`)(.+?)(?<!`)\1(?!`)", re.S)
# A word, quoted or escaped parts and placeholders such as ``<folder>``
# included, or an unquoted shell operator, which needs no spaces around it (``--check;``,
# ``2>&1``).
TOKEN = re.compile(r"""(?:"[^"]*"|'[^']*'|\\[^\s]|<[^\s<>|&;]+>|[^\s"'|&;<>])+|[|&;<>]+\d?""")
OPERATOR = re.compile(r"[|&;<>]+\d?")
INTERPRETER = re.compile(r"(?:python(?:3(?:\.\d+)?)?|py)(?:\.exe)?", re.I)
# Interpreter options that take a value.
INTERPRETER_VALUES = frozenset({"-X", "-W"})
CONTINUATIONS = ("\\", "`")
PLACEHOLDER = re.compile(r"[<>{}\[\]$*…]|\.\.\.|^[A-Z][A-Z0-9_./-]*$")
# Released sections are not rewritten (the changelog's maintenance policy),
# so the commands they cite stay as released.
HISTORY = {"CHANGELOG.md": re.compile(r"^## \[\d", re.M)}
TOOLS = "tools/"
# The tools that docs/setup/REPOSITORY_SETUP.md marks **Removes itself once run**.
REMOVED_ONCE_RUN = frozenset({"tools/initialize_repository.py"})
NO_VALUE = frozenset({"store_true", "store_false", "store_const", "append_const", "count",
                      "help", "version"})
PARSE = frozenset({"parse_args", "parse_known_args", "parse_intermixed_args",
                   "parse_known_intermixed_args"})
GROUPS = frozenset({"add_argument_group", "add_mutually_exclusive_group"})
MAX_DEPTH = 25


@dataclass(frozen=True)
class Finding:
    path: str
    line: int
    message: str

    def __str__(self) -> str:
        return f"{self.path}:{self.line}: {self.message}"


@dataclass
class Option:
    # How many values follow it: a count, or argparse's "?", "*", "+" or "...".
    takes: int | str


@dataclass
class Parser:
    options: dict[str, Option] = field(default_factory=dict)
    # Positional arguments, by the add_argument call that defined them.
    positionals: set[int] = field(default_factory=set)
    # Subcommands, once add_subparsers is called, and how many positional
    # arguments come before them.
    commands: dict[str, Parser] | None = None
    before_commands: int = 0
    parses: bool = False


@dataclass
class Subparsers:
    owner: Parser


Value = Parser | Subparsers | None


class Unreadable(Exception):
    """The parser cannot be read statically; the message says why."""


@dataclass
class Script:
    """A script's parser, or why it cannot be read (problem); neither: no parser."""

    parser: Parser | None = None
    problem: str | None = None


# ---------------------------------------------------------------------------
# Reading an argparse parser from a script's syntax tree.

def literal_string(node: ast.expr) -> str | None:
    return node.value if isinstance(node, ast.Constant) and isinstance(node.value, str) \
        else None


def keyword(call: ast.Call, name: str) -> ast.expr | None:
    return next((item.value for item in call.keywords if item.arg == name), None)


def called_name(node: ast.expr) -> str:
    return node.attr if isinstance(node, ast.Attribute) else \
        node.id if isinstance(node, ast.Name) else ""


class Reader:
    """Follows the script from its module body, through the functions it calls."""

    def __init__(self, tree: ast.Module) -> None:
        self.functions = {node.name: node for node in tree.body
                          if isinstance(node, ast.FunctionDef)}
        self.parsers: dict[int, Parser] = {}
        self.roots: list[Parser] = []
        self.stack: list[str] = []
        self.returned: list[Value] = []
        # Functions called with no parser among their arguments run once.
        self.done: dict[str, Value] = {}
        self.globals: dict[str, Value] = {}
        self.run(tree.body, self.globals)

    def run(self, body: Sequence[ast.stmt], scope: dict[str, Value]) -> None:
        for statement in body:
            self.statement(statement, scope)

    def statement(self, node: ast.stmt, scope: dict[str, Value]) -> None:
        if isinstance(node, ast.Assign | ast.AnnAssign | ast.AugAssign):
            assigned = self.evaluate(node.value, scope) if node.value is not None else None
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            for target in targets:
                if isinstance(target, ast.Name):
                    scope[target.id] = assigned
        elif isinstance(node, ast.Expr):
            self.evaluate(node.value, scope)
        elif isinstance(node, ast.Return | ast.Raise):
            value = node.value if isinstance(node, ast.Return) else node.exc
            result = self.evaluate(value, scope) if value is not None else None
            if result is not None and self.returned:
                self.returned[-1] = result
        elif not isinstance(node, ast.FunctionDef | ast.AsyncFunctionDef | ast.ClassDef):
            for name in ("body", "orelse", "finalbody"):
                self.run(getattr(node, name, []), scope)
            for handler in getattr(node, "handlers", []):
                self.run(handler.body, scope)
            for case in getattr(node, "cases", []):
                self.run(case.body, scope)

    def evaluate(self, node: ast.expr, scope: dict[str, Value]) -> Value:
        if isinstance(node, ast.Name):
            return scope[node.id] if node.id in scope else self.globals.get(node.id)
        if isinstance(node, ast.NamedExpr):
            scope[node.target.id] = self.evaluate(node.value, scope)
            return scope[node.target.id]
        if not isinstance(node, ast.Call):
            return None
        receiver = self.evaluate(node.func.value, scope) \
            if isinstance(node.func, ast.Attribute) else None
        arguments = [self.evaluate(item, scope) for item in node.args]
        named = {item.arg: self.evaluate(item.value, scope) for item in node.keywords
                 if item.arg is not None}
        name = called_name(node.func)
        if isinstance(receiver, Parser):
            return self.parser_method(receiver, name, node)
        if isinstance(receiver, Subparsers):
            return self.subcommand(receiver, node) if name == "add_parser" else None
        if name == "ArgumentParser":
            return self.new_parser(node, root=True)
        handed = [value for value in [*arguments, *named.values()] if value is not None]
        if isinstance(node.func, ast.Name) and name in self.functions:
            return self.call(self.functions[name], arguments, named)
        if handed:
            raise Unreadable(f"line {node.lineno} passes the parser to {ast.unparse(node.func)},"
                             " which the script does not define")
        return None

    def call(self, function: ast.FunctionDef, arguments: list[Value],
             named: dict[str, Value]) -> Value:
        handed = any(value is not None for value in [*arguments, *named.values()])
        if function.name in self.stack or len(self.stack) >= MAX_DEPTH:
            if handed:
                raise Unreadable(f"{function.name} is called recursively with the parser")
            return None
        if not handed and function.name in self.done:
            return self.done[function.name]
        parameters = [*function.args.posonlyargs, *function.args.args]
        scope: dict[str, Value] = {parameter.arg: value
                                   for parameter, value in zip(parameters, arguments)}
        scope.update({key: value for key, value in named.items()
                      if any(parameter.arg == key for parameter in
                             [*parameters, *function.args.kwonlyargs])})
        self.stack.append(function.name)
        self.returned.append(None)
        try:
            self.run(function.body, scope)
        finally:
            self.stack.pop()
        result = self.returned.pop()
        if not handed:
            self.done[function.name] = result
        return result

    def new_parser(self, node: ast.Call, root: bool) -> Parser:
        for unsupported in ("parents", "prefix_chars"):
            if keyword(node, unsupported) is not None:
                raise Unreadable(f"line {node.lineno} sets {unsupported}=, which is not read")
        if id(node) not in self.parsers:
            parser = self.parsers[id(node)] = Parser()
            add_help = keyword(node, "add_help")
            if not (isinstance(add_help, ast.Constant) and add_help.value is False):
                parser.options.update({"-h": Option(0), "--help": Option(0)})
            if root:
                self.roots.append(parser)
        return self.parsers[id(node)]

    def parser_method(self, parser: Parser, name: str, node: ast.Call) -> Value:
        if name == "add_argument":
            add_argument(parser, node)
        elif name in GROUPS:
            return parser
        elif name == "add_subparsers":
            if parser.commands is None:
                parser.commands = {}
                parser.before_commands = len(parser.positionals)
            return Subparsers(parser)
        elif name in PARSE:
            parser.parses = True
        return None

    def subcommand(self, subparsers: Subparsers, node: ast.Call) -> Parser:
        names = [literal_string(node.args[0])] if node.args else []
        aliases = keyword(node, "aliases")
        if isinstance(aliases, ast.List | ast.Tuple):
            names += [literal_string(item) for item in aliases.elts]
        elif aliases is not None:
            names.append(None)
        if not names or None in names:
            raise Unreadable(f"line {node.lineno} names a subcommand with something other than "
                             "a string literal")
        parser = self.new_parser(node, root=False)
        assert subparsers.owner.commands is not None
        subparsers.owner.commands.update({str(name): parser for name in names})
        return parser


def add_argument(parser: Parser, node: ast.Call) -> None:
    names = [literal_string(item) for item in node.args]
    if not names or None in names:
        raise Unreadable(f"line {node.lineno}: add_argument with names that are not string "
                         "literals")
    flags = [str(name) for name in names if str(name).startswith("-")]
    if not flags:
        parser.positionals.add(id(node))
        return
    action, nargs = keyword(node, "action"), keyword(node, "nargs")
    takes: int | str = 1
    if isinstance(nargs, ast.Constant) and isinstance(nargs.value, int | str):
        takes = nargs.value
    if action is not None and literal_string(action) in NO_VALUE:
        takes = 0
    if action is not None and called_name(action) == "BooleanOptionalAction":
        takes = 0
        flags += ["--no-" + flag[2:] for flag in flags if flag.startswith("--")]
    parser.options.update({flag: Option(takes) for flag in flags})


def read_script(path: Path) -> Script:
    """The parser a script builds, read from its syntax tree without running it."""
    try:
        source = path.read_text(encoding="utf-8")
        reader = Reader(ast.parse(source, filename=str(path)))
    except (OSError, UnicodeDecodeError, SyntaxError) as error:
        return Script(problem=f"it cannot be parsed: {error}")
    except Unreadable as error:
        return Script(problem=str(error))
    roots = reader.roots
    if len(roots) > 1:
        roots = [parser for parser in roots if parser.parses]
    if len(roots) > 1:
        return Script(problem=f"it builds {len(roots)} parsers that parse arguments")
    if not roots and "ArgumentParser" in source:
        return Script(problem="its ArgumentParser is not reached from the script's entry point")
    return Script(parser=roots[0] if roots else None)


# ---------------------------------------------------------------------------
# Checking a documented command against a parser.

def is_option(token: str) -> bool:
    return token.startswith("-") and token != "-" and not re.match(r"-\d", token)


def skip_values(option: Option, tokens: Sequence[str], index: int) -> int:
    """The index after the values that follow an option."""
    if isinstance(option.takes, int):
        return min(index + option.takes, len(tokens))
    if option.takes == "...":
        return len(tokens)
    end = index
    while end < len(tokens) and not is_option(tokens[end]):
        end += 1
        if option.takes == "?":
            break
    return end


def option_name(token: str) -> str:
    """The name an option token is written with, without a value attached to it:
    --name=VALUE gives --name, and a short -xVALUE gives -x."""
    return token.split("=", 1)[0] if token.startswith("--") else token[:2]


def find_option(parser: Parser, token: str) -> tuple[Option | None, bool]:
    """The option a token names, and whether its value is attached to it."""
    name, attached = token.split("=", 1)[0], "=" in token
    if name in parser.options:
        return parser.options[name], attached
    if token.startswith("--") or token[:2] not in parser.options:
        return None, False
    option = parser.options[token[:2]]
    rest = token[2:]
    while rest and option.takes == 0:
        # Short options combined, such as -vq: each takes no value but the last.
        if f"-{rest[0]}" not in parser.options:
            return None, False
        option, rest = parser.options[f"-{rest[0]}"], rest[1:]
    return option, bool(rest)


def check_arguments(parser: Parser | None, tokens: Sequence[str]) -> list[str]:
    """What in a command the parser does not define; an empty list if nothing."""
    problems: list[str] = []
    positionals, index, lost = 0, 0, False
    while index < len(tokens):
        token = tokens[index].lstrip("[(").rstrip("])|,")
        index += 1
        if tokens[index - 1] == "--":
            break
        if is_option(token):
            if PLACEHOLDER.search(option_name(token)):
                # The name itself is a placeholder, such as --<option>.
                continue
            if parser is None:
                problems.append(f"{token}: the script defines no argparse option")
                continue
            option, attached = find_option(parser, token)
            if option is None:
                problems.append(f"unknown option {token.split('=', 1)[0]}")
                lost = True
            elif not attached:
                index = skip_values(option, tokens, index)
            continue
        if parser is None or parser.commands is None or lost \
                or positionals != parser.before_commands:
            positionals += 1
            continue
        if PLACEHOLDER.search(token):
            break
        if token not in parser.commands:
            known = ", ".join(sorted(parser.commands)) or "none"
            problems.append(f"unknown subcommand {token} (defined: {known})")
            break
        parser, positionals = parser.commands[token], 0
    return problems


# ---------------------------------------------------------------------------
# Finding commands in Markdown.

def fenced_lines(lines: list[str], start: int, end: int) -> Iterator[tuple[int, str]]:
    """Lines start..end-1 of a shell block, each continued line joined to the next."""
    held: list[str] = []
    first = start
    for number in range(start, end):
        line = lines[number].rstrip()
        if not held:
            first = number
        if line.endswith(CONTINUATIONS) and (len(line) < 2 or line[-2].isspace()):
            held.append(line[:-1])
            continue
        yield first + 1, " ".join([*held, line])
        held = []
    if held:
        yield first + 1, " ".join(held)


def code_spans(paragraph: list[tuple[int, str]]) -> Iterator[tuple[int, str]]:
    """The inline code spans of a paragraph, with the line each starts on."""
    text = "\n".join(line for _, line in paragraph)
    for match in CODE_SPAN.finditer(text):
        yield paragraph[text.count("\n", 0, match.start())][0], match.group(2).replace("\n", " ")


def snippets(text: str) -> Iterator[tuple[int, str]]:
    """(line, code) for each inline code span and each line of a shell block."""
    lines = text.splitlines()
    paragraph: list[tuple[int, str]] = []
    index = 0
    while index < len(lines):
        line = lines[index]
        opening = FENCE.match(line)
        if opening:
            yield from code_spans(paragraph)
            paragraph = []
            marker = opening.group(1)
            close = re.compile(rf"\s*{re.escape(marker[0])}{{{len(marker)},}}\s*$")
            end = next((number for number in range(index + 1, len(lines))
                        if close.match(lines[number])), len(lines))
            if opening.group(2).lower() in SHELL_FENCES:
                yield from fenced_lines(lines, index + 1, end)
            index = end + 1
            continue
        if BLOCK_START.match(line):
            yield from code_spans(paragraph)
            paragraph = []
        if line.strip():
            paragraph.append((index + 1, line))
        index += 1
    yield from code_spans(paragraph)


class Operator(str):
    """An unquoted shell operator; a quoted ``";"`` is a word, not one of these."""


def split_tokens(code: str) -> list[str]:
    """The words and operators of a snippet, without quotes and comments.

    A ``#`` starts a comment only where a word starts, outside quotes; the
    comment ends with its line. Unquoted operators are returned as
    ``Operator``, so a word that is an operator only once its quotes are
    removed does not end a command.
    """
    tokens: list[str] = []
    comment_end = -1
    for match in TOKEN.finditer(code):
        token, start = match.group(), match.start()
        if start < comment_end:
            continue
        if token.startswith("#") and (start == 0 or code[start - 1].isspace()
                                      or code[start - 1] in "|&;<>"):
            newline = code.find("\n", start)
            comment_end = len(code) if newline < 0 else newline
            continue
        if OPERATOR.fullmatch(token):
            tokens.append(Operator(token))
        elif token not in CONTINUATIONS:
            tokens.append(token.replace('"', "").replace("'", ""))
    return tokens


def command_end(tokens: Sequence[str], start: int) -> int:
    for index in range(start, len(tokens)):
        if isinstance(tokens[index], Operator):
            return index
    return len(tokens)


def commands(code: str) -> Iterator[tuple[str, str, list[str]]]:
    """(kind, target, arguments) for each Python command: kind is "script" or "module"."""
    tokens = split_tokens(code)
    for index, token in enumerate(tokens):
        if not INTERPRETER.fullmatch(token.replace("\\", "/").rsplit("/", 1)[-1]):
            continue
        at = index + 1
        while at < len(tokens) and is_option(tokens[at]) and tokens[at] not in ("-m", "-c"):
            at += 2 if tokens[at] in INTERPRETER_VALUES else 1
        if at >= len(tokens) or tokens[at] == "-c":
            continue
        if tokens[at] == "-m":
            if at + 1 < len(tokens):
                yield "module", tokens[at + 1], tokens[at + 2:command_end(tokens, at + 2)]
            continue
        target = tokens[at].replace("\\", "/").removeprefix("./")
        if target.startswith(TOOLS) and target.endswith(".py") and not PLACEHOLDER.search(target):
            yield "script", target, tokens[at + 1:command_end(tokens, at + 1)]


# ---------------------------------------------------------------------------
# The check.

def module_script(module: str, tracked: set[str]) -> tuple[bool, str | None]:
    """Whether a module is this repository's, and the tracked file that runs it."""
    parts = module.split(".")
    if not any(name == f"{parts[0]}.py" or name.startswith(f"{parts[0]}/") for name in tracked):
        return False, None
    base = "/".join(parts)
    return True, next((name for name in (f"{base}.py", f"{base}/__main__.py")
                       if name in tracked), None)


def markdown_documents(root: Path, tracked: set[str]) -> Iterator[tuple[str, str]]:
    for name in sorted(item for item in tracked if item.endswith(".md")):
        path = root / name
        if not path.is_file():
            continue
        text = path.read_text(encoding="utf-8")
        history = HISTORY.get(name)
        match = history.search(text) if history else None
        yield name, text[:match.start()] if match else text


def command_findings(root: Path, tracked: set[str], scripts: dict[str, Script], name: str,
                     line: int, command: tuple[str, str, list[str]]) -> list[Finding]:
    kind, target, arguments = command
    shown = " ".join([f"python -m {target}" if kind == "module" else f"python {target}",
                      *arguments])
    script: str | None = target
    if kind == "module":
        script = module_script(target, tracked)[1]
        if script is None:
            return [Finding(name, line, f"`{shown}`: no tracked file runs module {target}")]
    assert script is not None
    if script not in tracked:
        return [Finding(name, line, f"`{shown}`: {script} is not a tracked file")]
    if script not in scripts:
        scripts[script] = read_script(root / script)
    parsed = scripts[script]
    if parsed.problem:
        return [Finding(name, line, f"`{shown}`: the parser of {script} cannot be read "
                                    f"statically: {parsed.problem}")]
    return [Finding(name, line, f"`{shown}`: {problem}")
            for problem in check_arguments(parsed.parser, arguments)]


def check(root: Path = ROOT) -> tuple[list[Finding], int, int]:
    """Return (findings, commands checked, Markdown files read)."""
    tracked = set(tracked_or_scanned(root))
    scripts: dict[str, Script] = {}
    findings: list[Finding] = []
    count = files = 0
    for name, text in markdown_documents(root, tracked):
        files += 1
        for line, code in snippets(text):
            for command in commands(code):
                if command[0] == "module" and not module_script(command[1], tracked)[0]:
                    continue
                if command[1] in REMOVED_ONCE_RUN and command[1] not in tracked:
                    continue
                count += 1
                findings += command_findings(root, tracked, scripts, name, line, command)
    return findings, count, files


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, count, files = check(args.root.resolve())
    for finding in findings:
        print(f"FAIL: {finding}")
    status = "FAIL" if findings else "PASS"
    print(f"documented-commands: {status} ({count} command(s) in {files} Markdown file(s); "
          "read from each tool's argparse parser, nothing run)")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
