#!/usr/bin/env python3
"""Check native Declare statements against their register, and how their errors are captured.

docs/native-declarations.json is the register of the native exports the
project declares. ``Err.LastDllError`` holds the error of the most recent
``Declare`` call; ``Err.Clear``, any ``On Error`` statement and the next
native call replace it, so it is trustworthy only when it is copied the
moment the call returns. The rules, applied in each compilation environment
of check_vba_conditionals.py, so that one branch of an ``#If VBA7`` split
cannot drift while another stays correct:

- every ``Declare`` names an export of the register: its ``Alias`` if it has
  one, else its own name, compared ignoring case, from the library the
  register gives for it;
- a ``Declare`` whose export has an ABI in the register, its ``arguments``
  or its ``returns``, matches it: the number of parameters, each one's type
  and ``ByVal`` or ``ByRef`` (``ByRef`` when neither is written), ``Sub`` for
  a ``returns`` of null, else ``Function`` and its type. A ``pointer`` is
  ``LongPtr`` where VBA7 is true and ``Long`` before it; every other type
  matches exactly, ignoring case. A type suffix gives a type, and a name
  without one or ``As`` is a ``Variant``;
- ``GetLastError`` is never declared: it does not see the error that VBA has
  already copied into ``Err.LastDllError``; and a ``Declare Sub``, which
  returns nothing, is never registered as ambiguous on zero;
- ``Err.LastDllError`` is read only as ``<local> = Err.LastDllError``, into a
  ``Dim`` or ``Static`` variable of the procedure, in the statement directly
  after one that is exactly a call to a ``Declare`` (``Name args``,
  ``Call Name(args)`` or ``<variable> = Name(args)``, the name alone or
  qualified by the module's name, the parentheses optional without
  arguments), not a call inside a condition, a single-line If, an expression
  or an argument, nor one assigned to an array element or a member; a named
  label in between counts as a statement, a line number does not; a
  ``.LastDllError`` read inside a ``With Err`` block, or through another
  qualifier, is rejected, since the ``With`` statement already stands
  between the call and the read;
- a call to an export the register marks ``zero_is_ambiguous``, whose zero
  result may mean success or failure, assigns its result, and nothing else,
  to a local variable in the call statement, is directly preceded by
  ``SetLastError 0`` (or ``Call SetLastError(0)``) through a ``Declare`` of
  the ``SetLastError`` export, and is directly followed by that capture.

A ``Declare`` is private to its module (check_vba_public_api.py rejects a
public one in a standard module, and VBA one in a class), so a call resolves
to a ``Declare`` of the calling module, by its name alone or qualified by the
module's name, the component's file name. The check reads statements in
order, not control flow: a jump into the middle of a sequence, and an error
replaced inside the call itself, are not seen. docs/VBA_HOUSE_STYLE.md,
Static checks, lists its limits.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections.abc import Iterator, Sequence
from dataclasses import dataclass, field
from pathlib import Path

from check_vba_conditionals import ENVIRONMENTS, reachable_sources
from check_vba_identifiers import (
    END_PROC,
    NAME,
    PROC,
    STATEMENT_WORDS,
    SUFFIX,
    closing_parenthesis,
    declared_names,
    top_level_split,
)
from vba_lexer import (
    Finding,
    logical_units,
    read_component,
    report,
    split_statements,
    tracked_components,
)

ROOT = Path(__file__).resolve().parents[1]
REGISTER = "docs/native-declarations.json"
REGISTER_KEYS = frozenset({"schema_version", "declarations"})
ENTRY_KEYS = frozenset({"export", "library", "zero_is_ambiguous"})
# The export's signature, its ABI: optional, and checked when present.
ABI_KEYS = frozenset({"arguments", "returns"})
ARGUMENT_KEYS = frozenset({"type", "byval"})
# A VBA type, possibly an array, or "pointer" for a pointer-sized value:
# LongPtr under VBA7 and Long before it.
POINTER = "pointer"
ABI_TYPE = re.compile(rf"{NAME}(?:\(\))?")
# The type a suffix gives a name, as in GetText$.
SUFFIX_TYPES = {"%": "Integer", "&": "Long", "!": "Single", "#": "Double", "@": "Currency",
                "$": "String", "^": "LongLong"}
PARAMETER = re.compile(rf"^\s*(?:Optional\s+)?(?:(?P<passing>ByVal|ByRef)\s+)?"
                       rf"(?:ParamArray\s+)?(?P<name>{NAME})(?P<suffix>[%&!#@$^])?"
                       rf"(?P<array>\s*\(\s*\))?(?:\s+As\s+(?P<type>{NAME}))?\s*(?:=.*)?$",
                       re.I)
RETURNS = re.compile(rf"^\s*(?:As\s+(?P<type>{NAME})(?P<array>\s*\(\s*\))?)?\s*$", re.I)
# An exported name, or an ordinal such as #12.
EXPORT = re.compile(r"[A-Za-z_][A-Za-z0-9_@$?]*|#[0-9]+")
DECLARE = re.compile(rf"^\s*(?:(?:Public|Private|Global)\s+)?Declare\s+(?:PtrSafe\s+)?"
                     rf"(?P<kind>Sub|Function)\s+(?P<name>{NAME})"
                     r"(?P<suffix>[%&!#@$^](?![A-Za-z_]))?", re.I)
# CDecl, which only Office for Mac reads, may stand before Lib.
LIBRARY = re.compile(r'(?:\s+CDecl)?\s+Lib\s+"(?P<library>[^"]*)"'
                     r'(?:\s+Alias\s+"(?P<alias>[^"]*)")?', re.I)
LOCAL = re.compile(r"^\s*(?:Dim|Static)\s+(?P<list>.*)$", re.I)
CONST = re.compile(r"^\s*Const\s", re.I)
LABEL = re.compile(rf"^\s*(?:(?P<name>{NAME})|\d+):(?!=)")
LINE_NUMBER = re.compile(r"^\s*\d+(?=\s)")
# Any read of a LastDllError member: Err.LastDllError, .LastDllError in a
# With Err block, or one through another qualifier such as VBA.Err.
LAST_DLL_ERROR = re.compile(r"\.\s*LastDllError\b", re.I)
PLAIN_LAST_DLL_ERROR = re.compile(r"(?<![\w.!])Err\s*\.\s*LastDllError\b", re.I)
CAPTURE = re.compile(rf"^\s*(?:Let\s+)?(?P<local>{NAME}){SUFFIX}\s*=\s*Err\s*\.\s*LastDllError"
                     r"\s*$", re.I)
ON_ERROR = re.compile(r"^\s*On\s+Error\b", re.I)
# A name, alone or after one qualifier, that is not itself the member of an
# object or of a With block: Name, or Module.Name.
WORD = re.compile(rf"(?<![\w.!])(?:(?P<qualifier>{NAME})\s*\.\s*)?(?P<name>{NAME})")
ASSIGNMENT = re.compile(rf"^\s*(?:Let\s+)?(?P<variable>{NAME}){SUFFIX}\s*=\s*", re.I)
EXPLICIT_CALL = re.compile(r"^\s*Call\s+", re.I)


@dataclass(frozen=True)
class Argument:
    """One parameter of an export: its type, or "pointer", and how it is passed."""

    type: str
    byval: bool


# The value of an entry's returns when the register does not record it.
NOT_RECORDED = "<not recorded>"


@dataclass(frozen=True)
class Entry:
    """One export of the register."""

    export: str
    library: str
    zero_is_ambiguous: bool
    # None when the register records no arguments for the export.
    arguments: tuple[Argument, ...] | None = None
    # A type, "pointer", None for a Sub, or NOT_RECORDED.
    returns: str | None = NOT_RECORDED


def export_key(export: str, library: str) -> tuple[str, str]:
    """How an export is matched: ignoring case, and a library's .dll suffix."""
    library = library.strip().casefold()
    return export.casefold(), library.removesuffix(".dll")


def read_register(root: Path) -> tuple[dict[tuple[str, str], Entry], list[Finding]]:
    """The register's entries by export and library, and what is wrong with it."""
    path = root / REGISTER
    if not path.is_file():
        return {}, [Finding(REGISTER, None, "is missing; restore it with an empty "
                                            "declarations list")]
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, ValueError) as error:
        return {}, [Finding(REGISTER, None, f"is not valid JSON: {error}")]
    if not isinstance(data, dict) or data.keys() != REGISTER_KEYS or \
            data["schema_version"] != 1 or not isinstance(data["declarations"], list):
        return {}, [Finding(REGISTER, None, "needs exactly schema_version 1 and a "
                                            "declarations list")]
    entries: dict[tuple[str, str], Entry] = {}
    findings: list[Finding] = []
    for index, item in enumerate(data["declarations"]):
        where = f"{REGISTER}: declarations[{index}]"
        problems = entry_problems(item)
        if not problems:
            arguments = item.get("arguments")
            entry = Entry(item["export"], item["library"], item["zero_is_ambiguous"],
                          None if arguments is None else
                          tuple(Argument(value["type"], value["byval"]) for value in arguments),
                          item.get("returns", NOT_RECORDED))
            key = export_key(entry.export, entry.library)
            if key in entries:
                problems.append(f"{entry.export} from {entry.library} is listed twice")
            elif key[0] == "getlasterror":
                problems.append("GetLastError is never declared, so it has no entry")
            else:
                entries[key] = entry
        findings.extend(Finding(where, None, problem) for problem in problems)
    return entries, findings


def entry_problems(item: object) -> list[str]:
    """What is wrong with one entry of the declarations list."""
    if not isinstance(item, dict):
        return ["is not an object"]
    problems = [f"unknown key {key!r}" for key in sorted(item.keys() - ENTRY_KEYS - ABI_KEYS)]
    problems += [f"{key} is missing" for key in sorted(ENTRY_KEYS - item.keys())]
    if problems:
        return problems
    if not isinstance(item["export"], str) or not EXPORT.fullmatch(item["export"]):
        problems.append("export is not an export name or an ordinal such as #12")
    library = item["library"]
    if not isinstance(library, str) or not library.strip() or library != library.strip() \
            or '"' in library:
        problems.append("library is not a library name, such as kernel32")
    if not isinstance(item["zero_is_ambiguous"], bool):
        problems.append("zero_is_ambiguous is not true or false")
    elif item["zero_is_ambiguous"] and "returns" in item and item["returns"] is None:
        problems.append("zero_is_ambiguous is true for an export that returns nothing")
    return problems + abi_problems(item)


def type_problem(value: object, what: str) -> str | None:
    """What is wrong with a type of the ABI, or None."""
    if not isinstance(value, str) or not ABI_TYPE.fullmatch(value):
        return f"{what} is not a VBA type or {POINTER}"
    if value.casefold().startswith("longptr"):
        return f"{what} is {value}; write {POINTER} for a pointer-sized value"
    return None


def abi_problems(item: dict[str, object]) -> list[str]:
    """What is wrong with the arguments and returns of one entry."""
    problems: list[str] = []
    arguments = item.get("arguments")
    if "arguments" in item and not isinstance(arguments, list):
        problems.append("arguments is not a list")
    for index, value in enumerate(arguments if isinstance(arguments, list) else []):
        where = f"arguments[{index}]"
        if not isinstance(value, dict) or value.keys() != ARGUMENT_KEYS:
            problems.append(f"{where} does not have exactly the keys byval and type")
            continue
        problem = type_problem(value["type"], f"{where}.type")
        if problem:
            problems.append(problem)
        if not isinstance(value["byval"], bool):
            problems.append(f"{where}.byval is not true or false")
    if item.get("returns") is not None:
        problem = type_problem(item["returns"], "returns")
        if problem:
            problems.append(problem.replace(f" or {POINTER}", f", {POINTER} or null"))
    return problems


@dataclass(frozen=True)
class Parameter:
    """One parameter of a Declare, as written."""

    name: str
    type: str
    byval: bool


@dataclass(frozen=True)
class Declared:
    """One Declare statement: its VBA name, the export it imports and its
    signature, whose parameters are None when they cannot be read."""

    line: int
    name: str
    export: str
    library: str
    sub: bool
    parameters: tuple[Parameter, ...] | None = None
    # The Function's type; None for a Sub.
    returns: str | None = None


@dataclass
class Procedure:
    name: str
    locals: set[str] = field(default_factory=set)
    # Executable statements, and labels as "<label>:", in source order.
    body: list[tuple[int, str]] = field(default_factory=list)


@dataclass
class Module:
    # The component's name, which may qualify a call to its own Declare.
    name: str
    declared: dict[str, Declared] = field(default_factory=dict)
    procedures: list[Procedure] = field(default_factory=list)
    findings: list[Finding] = field(default_factory=list)


def declaration(path: str, lines: list[str], start: int, end: int,
                module: Module) -> None:
    """Record the Declare statement of the logical line ``start``..``end``.

    The statements were read with string literals blanked; the library and
    the alias are read again from the source, where they are kept."""
    for _, _, kind, code in logical_units(lines[start - 1:end]):
        for statement in split_statements(code) if kind == "code" else []:
            head = DECLARE.match(statement)
            if head is None:
                continue
            names = LIBRARY.match(statement, head.end())
            if names is None:
                module.findings.append(Finding(path, start, "cannot read the Lib clause of "
                                                            f"Declare {head.group('name')}"))
                continue
            alias = names.group("alias")
            sub = head.group("kind").casefold() == "sub"
            parameters, returns = signature(statement[names.end():], head.group("suffix"))
            module.declared[head.group("name").casefold()] = Declared(
                start, head.group("name"), head.group("name") if alias is None else alias,
                names.group("library"), sub, parameters, None if sub else returns)


def written_type(declared: str | None, suffix: str | None, array: str | None) -> str:
    """The type a declaration gives: As, else a type suffix, else Variant."""
    name = declared or SUFFIX_TYPES.get(suffix or "", "Variant")
    return name + ("()" if array else "")


def signature(rest: str, suffix: str | None) -> tuple[tuple[Parameter, ...] | None, str | None]:
    """The parameters and the return type written after a Declare's Lib and
    Alias clauses; None for the parameters when they cannot be read."""
    rest = rest.strip()
    inner = ""
    if rest.startswith("("):
        end = closing_parenthesis(rest, 0)
        if end is None:
            return None, None
        inner, rest = rest[1:end], rest[end + 1:]
    returned = RETURNS.match(rest)
    if returned is None:
        return None, None
    parameters = []
    for item in top_level_split(inner) if inner.strip() else []:
        match = PARAMETER.match(item)
        if match is None:
            return None, None
        parameters.append(Parameter(
            match.group("name"),
            written_type(match.group("type"), match.group("suffix"), match.group("array")),
            (match.group("passing") or "ByRef").casefold() == "byval"))
    return tuple(parameters), written_type(returned.group("type"), suffix,
                                           returned.group("array"))


def scan(path: str, source: str) -> Module:
    """The Declare statements and the procedure bodies of one environment's source."""
    lines = source.split("\n")
    module = Module(Path(path).stem)
    current: Procedure | None = None
    for start, end, kind, code in logical_units(lines, blank_strings=True):
        if kind == "directive":
            continue
        # A named label is a jump target that may sit between a call and its
        # capture. A line number is not kept: code numbered for Erl numbers
        # every line, and a GoTo to a number is rare.
        number = LINE_NUMBER.match(code)
        code = code[number.end():] if number else code
        label = LABEL.match(code)
        name = label.group("name") if label else None
        if label and (name is None or name.casefold() not in STATEMENT_WORDS):
            if current is not None and name is not None:
                current.body.append((start, name + ":"))
            code = code[label.end():]
        for statement in split_statements(code):
            if not statement:
                continue
            opener = PROC.match(statement)
            if DECLARE.match(statement):
                declaration(path, lines, start, end, module)
            elif opener:
                current = Procedure(opener.group("name"))
                module.procedures.append(current)
            elif END_PROC.match(statement):
                current = None
            elif current is not None:
                local_or_statement(current, start, statement)
    return module


def local_or_statement(procedure: Procedure, line: int, statement: str) -> None:
    """Record a Dim or Static as locals; skip a Const; keep any other statement."""
    local = LOCAL.match(statement)
    if local:
        procedure.locals |= {name.casefold() for name in declared_names(local.group("list"))
                             or []}
    elif not CONST.match(statement):
        procedure.body.append((line, statement))


@dataclass(frozen=True)
class NativeCall:
    """A statement that is exactly one call to a Declare."""

    item: Declared
    # The variable the result is assigned to, or None for a statement call.
    variable: str | None
    arguments: str


def declare_at(text: str, module: Module) -> tuple[Declared, str] | None:
    """The Declare that ``text`` starts with, as ``Name`` or ``<module>.Name``,
    and the text after it."""
    match = WORD.match(text)
    if match is None or (match.group("qualifier") and
                         match.group("qualifier").casefold() != module.name.casefold()):
        return None
    item = module.declared.get(match.group("name").casefold())
    rest = text[match.end():]
    suffix = re.match(SUFFIX, rest)
    rest = rest[suffix.end():] if suffix else rest
    # A longer name, a member or a ! access follows a name that is not the call.
    if item is None or re.match(r"[\w.!]", rest):
        return None
    return item, rest


def native_call(statement: str | None, module: Module) -> NativeCall | None:
    """The call, when ``statement`` is exactly one call to a Declare of the module:
    ``Name args``, ``Call Name(args)`` or ``<variable> = Name(args)``, the name
    alone or qualified by the module's name. A call inside a condition, a
    single-line If or a larger expression is not one."""
    if statement is None:
        return None
    assignment = ASSIGNMENT.match(statement)
    explicit = EXPLICIT_CALL.match(statement)
    head = assignment or explicit
    found = declare_at(statement[head.end():] if head else statement.lstrip(), module)
    if found is None:
        return None
    item, rest = found
    if head is None:
        # A statement call: its arguments, if any, follow a space or a parenthesis.
        if rest.strip() and (not rest[:1].isspace() and rest[:1] != "("
                             or rest.lstrip().startswith("=")):
            return None
        return NativeCall(item, None, rest.strip())
    rest = rest.strip()
    arguments = ""
    if rest.startswith("("):
        end = closing_parenthesis(rest, 0)
        if end is None:
            return None
        arguments, rest = rest[1:end], rest[end + 1:]
    if rest.strip():
        return None
    return NativeCall(item, assignment.group("variable") if assignment else None,
                      arguments.strip())


def called(statement: str, module: Module) -> list[Declared]:
    """The Declares a statement names, alone or qualified by the module's name,
    each once, in order."""
    found: dict[str, Declared] = {}
    for match in WORD.finditer(statement):
        if match.group("qualifier") and \
                match.group("qualifier").casefold() != module.name.casefold():
            continue
        item = module.declared.get(match.group("name").casefold())
        if item is not None:
            found.setdefault(item.name.casefold(), item)
    return list(found.values())


def is_clear(statement: str | None, module: Module) -> bool:
    """Whether a statement is SetLastError 0, through any Declare of that export."""
    call = native_call(statement, module)
    return call is not None and call.variable is None and \
        call.item.export.casefold() == "setlasterror" and \
        re.fullmatch(r"\(?\s*0&?\s*\)?", call.arguments) is not None


def reads_error(call: NativeCall | None) -> bool:
    """Whether a capture may follow this call: one to a Declare other than SetLastError."""
    return call is not None and call.item.export.casefold() != "setlasterror"


def capture_problem(procedure: Procedure, index: int, module: Module) -> str | None:
    """What is wrong with the Err.LastDllError read of ``procedure.body[index]``."""
    statement = procedure.body[index][1]
    if len(LAST_DLL_ERROR.findall(statement)) != len(PLAIN_LAST_DLL_ERROR.findall(statement)):
        return (f"{procedure.name} reads LastDllError through a With block or a qualifier "
                "other than Err; write it as <local> = Err.LastDllError")
    capture = CAPTURE.match(statement)
    if capture is None:
        return (f"{procedure.name} reads Err.LastDllError inside a larger statement; copy "
                "it by itself to a local variable (<local> = Err.LastDllError)")
    if capture.group("local").casefold() not in procedure.locals:
        return (f"{procedure.name} copies Err.LastDllError to {capture.group('local')}, "
                "which is not a Dim or Static variable of the procedure")
    if index == 0:
        return f"{procedure.name} reads Err.LastDllError before any native call"
    line, before = procedure.body[index - 1]
    if before.endswith(":"):
        return (f"{procedure.name} reads Err.LastDllError after the label of line {line}, "
                "which a jump can reach without the native call")
    if ON_ERROR.match(before):
        return (f"{procedure.name} reads Err.LastDllError after the On Error statement of "
                f"line {line}, which clears it; read it directly after the native call")
    if reads_error(native_call(before, module)):
        return None
    inner = [item.name for item in called(before, module)
             if item.export.casefold() != "setlasterror"]
    if inner:
        return (f"{procedure.name} reads Err.LastDllError after the statement of line "
                f"{line}, which calls {inner[0]} but is not a call statement of its own, "
                "such as a call inside a condition, a single-line If, an expression or an "
                "argument, or one assigned to an array element or a member; write Name args, "
                "Call Name(args) or <variable> = Name(args), then capture")
    return (f"{procedure.name} reads Err.LastDllError after the statement of line "
            f"{line}, not directly after a native call")


def ambiguous_problems(procedure: Procedure, index: int, item: Declared,
                       module: Module) -> list[str]:
    """What is wrong around a call to an export whose zero result is ambiguous."""
    body = procedure.body
    statement = body[index][1]
    before = body[index - 1][1] if index else None
    after = body[index + 1][1] if index + 1 < len(body) else None
    call = f"{procedure.name} calls {item.name} ({item.export}), whose zero result is ambiguous,"
    result = native_call(statement, module)
    problems = []
    if result is None or result.item != item or result.variable is None or \
            result.variable.casefold() not in procedure.locals:
        problems.append(f"{call} without assigning its result to a local variable in the "
                        "call statement")
    if not is_clear(before, module):
        problems.append(f"{call} without SetLastError 0 directly before it")
    capture = CAPTURE.match(after or "")
    if capture is None or capture.group("local").casefold() not in procedure.locals:
        problems.append(f"{call} without <local> = Err.LastDllError directly after it")
    return problems


def procedure_findings(path: str, procedure: Procedure, module: Module,
                       register: dict[tuple[str, str], Entry]) -> Iterator[Finding]:
    """The capture findings of one procedure in one environment."""
    for index, (line, statement) in enumerate(procedure.body):
        if LAST_DLL_ERROR.search(statement):
            problem = capture_problem(procedure, index, module)
            if problem:
                yield Finding(path, line, problem)
        for item in called(statement, module):
            entry = register.get(export_key(item.export, item.library))
            if entry is not None and entry.zero_is_ambiguous:
                yield from (Finding(path, line, problem)
                            for problem in ambiguous_problems(procedure, index, item, module))


def expected_type(value: str, environment: str) -> str:
    """The VBA type a type of the ABI stands for in one environment."""
    if value.casefold() != POINTER:
        return value
    return "LongPtr" if ENVIRONMENTS[environment]["VBA7"] else "Long"


def shown(value: str, environment: str) -> str:
    """A type of the ABI as a message shows it."""
    resolved = expected_type(value, environment)
    return f"{resolved} ({POINTER})" if resolved != value else value


def return_findings(item: Declared, entry: Entry, environment: str, label: str) -> list[str]:
    """How a Declare's Sub or Function and return type differ from the ABI."""
    if entry.returns == NOT_RECORDED:
        return []
    if entry.returns is None:
        return [] if item.sub else [f"{label} is a Function returning {item.returns}; the "
                                    "register expects a Sub (returns null)"]
    expected = shown(entry.returns, environment)
    if item.sub:
        return [f"{label} is a Sub; the register expects a Function returning {expected}"]
    if (item.returns or "").casefold() != expected_type(entry.returns, environment).casefold():
        return [f"{label} returns {item.returns}; the register expects {expected}"]
    return []


def abi_findings(item: Declared, entry: Entry, environment: str) -> list[str]:
    """How a Declare differs from its export's registered ABI in one environment."""
    if entry.arguments is None and entry.returns == NOT_RECORDED:
        return []
    label = f"Declare {item.name} ({item.export})"
    if item.parameters is None:
        return [f"{label}: cannot read its parameters or return type to compare them with "
                f"the ABI in {REGISTER}"]
    problems = return_findings(item, entry, environment, label)
    if entry.arguments is None:
        return problems
    if len(item.parameters) != len(entry.arguments):
        return problems + [f"{label} has {len(item.parameters)} parameter(s); the register "
                           f"expects {len(entry.arguments)}"]
    for number, (parameter, argument) in enumerate(zip(item.parameters, entry.arguments,
                                                       strict=True), start=1):
        where = f"{label}: parameter {number} ({parameter.name})"
        if parameter.type.casefold() != expected_type(argument.type, environment).casefold():
            problems.append(f"{where} is {parameter.type}; the register expects "
                            f"{shown(argument.type, environment)}")
        if parameter.byval != argument.byval:
            problems.append(f"{where} is passed {'ByVal' if parameter.byval else 'ByRef'}; "
                            f"the register expects {'ByVal' if argument.byval else 'ByRef'}")
    return problems


def analyze_source(path: str, source: str, register: dict[tuple[str, str], Entry],
                   environment: str) -> list[Finding]:
    """The findings of one environment's source."""
    module = scan(path, source)
    findings = list(module.findings)
    for item in module.declared.values():
        if item.export.casefold() == "getlasterror":
            findings.append(Finding(path, item.line, f"Declare {item.name} imports "
                                                     "GetLastError; copy Err.LastDllError "
                                                     "directly after the native call instead"))
        entry = register.get(export_key(item.export, item.library))
        if entry is not None and entry.zero_is_ambiguous and item.sub:
            findings.append(Finding(path, item.line, f"Declare {item.name} is a Sub, which "
                                                     f"returns nothing, but {REGISTER} marks "
                                                     f"{item.export} as ambiguous on zero"))
        if entry is not None:
            findings.extend(Finding(path, item.line, problem)
                            for problem in abi_findings(item, entry, environment))
        if entry is None:
            findings.append(Finding(path, item.line, f"Declare {item.name} imports "
                                                     f"{item.export} from {item.library}, "
                                                     f"which {REGISTER} does not list"))
    for procedure in module.procedures:
        findings.extend(procedure_findings(path, procedure, module, register))
    return findings


def analyze(path: str, text: str, register: dict[tuple[str, str], Entry]
            ) -> tuple[list[Finding], set[int]]:
    """Findings of every environment, each listed once with the environments it
    occurs in; and the lines of the Declare statements found."""
    # Unreadable directives are vba-conditionals findings, not repeated here.
    sources, _ = reachable_sources(path, text)
    environments: dict[Finding, list[str]] = {}
    declares: set[int] = set()
    for environment, source in sources.items():
        declares |= {item.line for item in scan(path, source).declared.values()}
        for finding in analyze_source(path, source, register, environment):
            names = environments.setdefault(finding, [])
            if environment not in names:
                names.append(environment)
    return [Finding(item.path, item.line, f"{item.message} [{', '.join(names)}]")
            for item, names in environments.items()], declares


def check(root: Path) -> tuple[list[Finding], int, int, int]:
    register, findings = read_register(root)
    components = tracked_components(root)
    declares = 0
    for path in components:
        found, lines = analyze(path, read_component(root, path), register)
        findings.extend(found)
        declares += len(lines)
    return findings, len(components), declares, len(register)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components, declares, registered = check(args.root.resolve())
    return report("vba-declares", findings,
                  f"{components} VBA component(s), {declares} Declare statement(s), "
                  f"{registered} registered export(s)")


if __name__ == "__main__":
    sys.exit(main())
