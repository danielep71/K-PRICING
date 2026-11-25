#!/usr/bin/env python3
"""Check VBA error handling: Err is copied before any other statement can change it.

A read of ``Err.Number``, ``Err.Source`` or ``Err.Description`` is reported
when, on the straight-line path that leads to it, a statement has already
cleared or may have replaced the error it means to read:

- an ``On Error`` statement or ``Err.Clear``, which clears ``Err``;
- in an error handler, a call or a member access (``Helper``, ``x = Helper(1)``,
  ``x = Helper`` without parentheses, ``obj.Method``, ``Debug.Print``,
  ``Application.StatusBar = ...``, ``New``), which can run code that clears
  or replaces it. A name without parentheses is a call unless it
  is a variable, a constant, an Enum member or a component: a function,
  property or Declare of the project, or a host global such as ActiveCell.
  Names that start with ``vb`` or ``xl`` are the constants of VBA and Excel.
  A name is resolved against the project first: a procedure, property or
  Declare of the module, or a public one of a standard module, shadows a VBA
  function, statement, statement target or constant of its name, as in VBA,
  while ``VBA.Format`` still names the VBA function. ``Err`` and ``Erl``
  always mean VBA's. Within the project, a name, bare or first in a dotted
  name, means what the nearest scope level declares (Scope.resolve): the
  procedure's parameters, variables and own name (the result of a
  Function or Property Get, of its declared type); then any declaration of
  its module; then a public one of another component; then a component. A
  procedure there wins over a variable of a farther level, and the reverse.
  When that is not certain, a dotted name is a member access: when it
  starts, unqualified, with a variable that two components declare
  publicly, when it goes through a Type or Enum that two components
  declare, or when its type is qualified by anything but the component
  that declares it (``As VBA.Collection``). A name qualified by its
  component, as ``Module.Kept``, is resolved in that component.
  A dotted name is a member access unless it reads a value (Scope.value): a
  field of a variable declared ``As`` a user-defined Type of the project
  (``Saved.Number``, ``Saved(i).Inner.Field``, ``.Number`` inside
  ``With Saved``), an Enum member (``RunStatus.rsFailed``,
  ``Module.RunStatus.rsFailed``), or a constant, Enum member or variable of
  a standard module that its name qualifies (``Module.LIMIT``). A variable
  of a class or object type, a class or document module and ``Me`` stay
  member accesses. In a dotted name, an index is an element only on a name
  declared as an array (``Name() As T``, ``Name(1 To 3) As T``, a
  ``ReDim``): on an object, a Variant or an undeclared type it may call a
  default property, as ``Owner(1)`` does on a Collection, and it stays a
  member access; an index before a dot that holds a parenthesis, as in
  ``Saved(Count(1)).Number``, is one too. A bare ``C(1)`` is read as an
  element, even on a Collection, whose Item it calls. The part
  a dotted name reaches must hold data: a value type (Long, String, Date,
  ...), an Enum or a user-defined Type. Reading or assigning an object, a
  class or a Variant calls its default property, so ``Number = rec.Owner``
  stays a member access, even when the Variant holds a plain value. So does
  ``Set Book = rec.Owner``: a variable declared ``As New`` runs its class's
  Class_Initialize when it is first read.
  In the same way, an assignment (``=``, ``Let``, ``Set``, ``LSet``,
  ``RSet``) to anything but a variable, such a field or value, an array
  element or a statement target (``Date``, ``Mid(Text, 1, 1)``) sets a
  property, a call made once the value is read: a Property Let or Set of
  the project, a member (``.Value = x`` in a ``With`` block) or a property
  of the host.

After ``On Error Resume Next``, and outside a handler, ``Err`` belongs to the
most recent statement that runs anything, since any of them can fail, even
``Number = 1`` (a type mismatch) or ``Matches = Text Like Pattern``: a read
after one checks it and is not reported. Only these leave ``Err`` as the
``On Error`` statement or ``Err.Clear`` cleared it: a declaration (``Dim``,
``Static``, ``Const``), a label, a blank line or a comment, ``With Err`` and
``End With``, and a write of a property of ``Err`` whose value runs nothing,
as in ``Err.Number = 0``.

In a handler, the VBA functions of PURE_FUNCTIONS (``CLng``, ``CStr``,
``Len``, ..., also written ``VBA.CLng``), operators, an array element and
the VBA statements of BUILTIN_STATEMENT (``ReDim``, ``Erase``, ``Open``,
``Print #``, ...) change nothing: they can fail but run no code that clears
``Err``. A VBA statement may be written with ``Call``, as in
``Call SaveSetting(...)``. Its operands alone are read as expressions,
without the type that ``ReDim`` declares (``As Long``), and the type in
``TypeOf x Is Thing`` is not a value. A statement that starts with any other
name is a call, ``Load`` and ``Unload`` included, since they run the form's
event procedures.

Within a statement, the check follows VBA's evaluation order: a read inside
the arguments of a call, as in ``CLng(Err.Number)``, runs before that call.
Only what runs before a read counts for it, under the same rules as the
statements before: after ``On Error Resume Next`` any name, call or operator
other than a literal, a constant or a property of ``Err`` refreshes ``Err``, and
in a handler a call or member access makes the read stale, as in
``Report Helper(), Err.Number``. An operator written before a read may run
after it, as ``&`` in ``a & b / Err.Number``; it counts as running before.
A name qualified by the VBA library or one of its modules (``VBA.CLng``,
``VBA.Strings.Len``, ``VBA.vbCrLf``, ``VBA.Beep``, ``VBA.Mid``) is VBA's, as
unqualified, whatever the project declares; a module name alone, as in
``Strings.Len``, only when no value, component or procedure of the project
has that name.
Reading or writing any property of ``Err`` (``Number``, ``Source``,
``Description``, ``HelpFile``, ``HelpContext``, ``LastDllError``) is data:
it changes nothing, and only a read of the first three is checked. Assigning
one (``Err.Description = x``) reads nothing of it, and ``.Number`` inside
``With Err`` reads ``Err.Number``. ``Err.Raise`` is a call.
``Err`` read through its default member (``n = Err``), ``Err`` passed to a
procedure and ``Erl`` are not followed.

A segment starts at each label and after each branch or loop statement:
``If``, ``Select Case``, ``Do``, ``For``, ``While``, ``Exit``, ``Resume``,
``GoTo``, ``Return`` and the statements that close them end the segment
being tracked. A single-line ``If`` is read as its condition, then each
branch on its own, as the statements it runs from the state after the
condition (an ``Else`` belongs to the nearest ``If``); the segment ends
after the line. What follows a branch is never claimed, so a finding is a
read that is stale on every path to it.

Inside a Function or Property Get, its own name is the result variable,
and with an argument list, as in ``Fetch(True)``, a call of itself, unless
the declared result is an array or a Variant, which an index may read: then
it counts as an element of the result.

Each compilation environment of check_vba_conditionals.py is analyzed
separately, and a finding names the environments it occurs in.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Sequence
from dataclasses import dataclass, field, replace
from pathlib import Path

from check_vba_calls import dates_as_values
from check_vba_conditionals import reachable_sources
from check_vba_identifiers import Module, closing_parenthesis, scan, statements, top_level_split
from check_vba_jumps import PROC_CLOSE, is_procedure, label_at_start
from vba_lexer import (
    Finding,
    logical_units,
    read_component,
    report,
    split_statements,
    tracked_components,
)

ROOT = Path(__file__).resolve().parents[1]
ERR_READ = re.compile(r"(?<![\w.!])(?:VBA\s*\.\s*)?Err\s*\.\s*(Number|Source|Description)\b",
                      re.IGNORECASE)
# Every property of Err: reading or writing one is data, not a call. Only
# Number, Source and Description are checked for being read stale.
ERR_PROPERTY = re.compile(r"(?<![\w.!])(?:VBA\s*\.\s*)?Err\s*\.\s*"
                          r"(?:Number|Source|Description|HelpFile|HelpContext|LastDllError)\b",
                          re.IGNORECASE)
ON_ERROR = re.compile(r"^On\s+Error\b", re.IGNORECASE)
RESUME_NEXT = re.compile(r"^On\s+Error\s+Resume\s+Next\b", re.IGNORECASE)
HANDLER = re.compile(r"\bOn\s+Error\s+GoTo\s+([A-Za-z_]\w*|\d+)", re.IGNORECASE)
# A single-line If: its condition and the statements after Then, which
# may hold an Else and the statements after it.
SINGLE_LINE_IF = re.compile(r"^If\b(?P<condition>.*?)\bThen\b(?P<tail>.*)$", re.IGNORECASE)
IF_OR_ELSE = re.compile(r"\b(?:If|Else)\b", re.IGNORECASE)
# The types that a suffix gives a name, as Fetch$ is a String.
SUFFIX_TYPES = {"%": "integer", "&": "long", "!": "single", "#": "double", "@": "currency",
                "$": "string"}
# The declared result of a function or Property Get, after its parameters.
RESULT_TYPE = re.compile(r"^\s*As\s+(?P<type>[A-Za-z_][\w.]*)\s*(?P<array>\(\s*\))?",
                         re.IGNORECASE)
# Any property of Err as an assignment target, as in Err.Description = x.
ERR_CLEAR = re.compile(r"^(?:VBA\s*\.\s*)?Err\s*\.\s*Clear\b", re.IGNORECASE)
ERR_TARGET = re.compile(r"(?:VBA\s*\.\s*)?Err\s*\.\s*[A-Za-z_]\w*", re.IGNORECASE)
WITH_ERR = re.compile(r"^With\s+(?:VBA\s*\.\s*)?Err\s*$", re.IGNORECASE)
# A member of the object of a With block, as in .Number; not x.y, f(1).y or 1.5.
WITH_MEMBER = re.compile(r"(?<![\w)\].!])\.(?=\s*[A-Za-z_])")
# What an operand is an argument or operand of, before it: the name of an
# enclosing call (the last name only, as in .Method of obj.Method(), since
# obj is read first) and an operator, which both run after the operand.
# The VBA library and its modules, as in VBA.CLng, VBA.Strings.Len or
# Strings.Len: what VBA. qualifies is VBA's, whatever the project declares; a
# module name alone is VBA's only when nothing of the project has that name
# (library_qualifier).
VBA_MODULES = (r"Constants|Conversion|DateTime|FileSystem|Financial|Information|Interaction"
               r"|Math|Strings")
VBA_QUALIFIER = (rf"(?:VBA\s*\.\s*(?:(?:{VBA_MODULES})\s*\.\s*)?"
                 rf"|(?P<module>{VBA_MODULES})\s*\.\s*)")
VBA_HEAD = re.compile(rf"^{VBA_QUALIFIER}", re.IGNORECASE)
ENCLOSING_CALL = re.compile(rf"(?:(?<![\w.!]){VBA_QUALIFIER}|\.\s*)?[A-Za-z_]\w*[%&!#@$]?\s*$",
                            re.IGNORECASE)
TRAILING_OPERATOR = re.compile(r"(?:[-&+*/\\^<>=]|\b(?:Mod|Not|And|Or|Xor|Eqv|Imp|Like|Is)\b)\s*$",
                               re.IGNORECASE)
# The loop variable of a For, assigned once its bounds are read.
FOR_HEAD = re.compile(r"^For\s+(?:Each\s+[A-Za-z_]\w*\s+In\b|[A-Za-z_]\w*[%&!#@$]?\s*=)",
                      re.IGNORECASE)
NAMED_ARGUMENT = re.compile(r"\b[A-Za-z_]\w*\s*:=?\s*$")
LABEL_PREFIX = re.compile(r"^\s*(?:[A-Za-z_]\w*:(?!=)|\d+:?)")
# Statements that end the tracked segment once their own expression is read:
# that expression runs straight after the statement before.
BRANCH = re.compile(r"^(?:If|Select|Do|For|While|Exit|Resume|GoTo|GoSub|Return|Stop|End"
                    r"(?!\s+With\b))\b|^On\s+(?!Error\b)", re.IGNORECASE)
# Statements reached from elsewhere, which end the segment before anything is read.
JOIN = re.compile(r"^(?:ElseIf|Else|Case|Loop|Next|Wend|End\s+If|End\s+Select)\b", re.IGNORECASE)
DECLARATION = re.compile(r"^(?:Dim|Static|Const)\b", re.IGNORECASE)
WITH = re.compile(r"^With\b", re.IGNORECASE)
END_WITH = re.compile(r"^End\s+With\b", re.IGNORECASE)
# Statements of the VBA language that can fail but run no project code: only
# their operands are expressions, and the words of their syntax are not names.
BUILTIN_STATEMENTS = (
    "ReDim", "Erase", "Open", "Close", "Print", "Write", "Input", "Line Input", "Get", "Put",
    "Seek", "Lock", "Unlock", "Kill", "MkDir", "RmDir", "ChDir", "ChDrive", "FileCopy", "Name",
    "SetAttr", "Randomize", "Beep", "Reset", "Width", "Error", "AppActivate", "SaveSetting",
    "DeleteSetting", "SendKeys")
BUILTIN_STATEMENT = re.compile(
    r"^(?:ReDim\s+Preserve|" + "|".join(name.replace(" ", r"\s+") for name in BUILTIN_STATEMENTS)
    + r")\b", re.IGNORECASE)
CALL_KEYWORD = re.compile(r"^Call\s+", re.IGNORECASE)
# The type of a variable that ReDim declares, as in ReDim a(1 To n) As Long,
# As New Thing or As String * 10: a name of a type, not a value or a call.
TYPE_CLAUSE = re.compile(r"\s+As\s+(?:New\s+)?[A-Za-z_][\w.]*(?:\s*\*\s*\w+)?\s*$",
                         re.IGNORECASE)
# The type that TypeOf x Is Thing compares with.
TYPEOF_TYPE = re.compile(r"(\bTypeOf\b[^,()]*?\bIs\s+)[A-Za-z_][\w.]*", re.IGNORECASE)
STATEMENT_SYNTAX = re.compile(r"\b(?:Access|Append|Binary|Input|Lock|Output|Random|Read|Shared"
                              r"|Write)\b", re.IGNORECASE)
# Assignment targets that are VBA statements: Date = x sets the system date,
# Mid(Text, 1, 1) = x replaces characters.
BUILTIN_TARGETS = frozenset({"date", "time", "mid", "midb"})
ASSIGNMENT_KEYWORD = re.compile(r"^(?:Let|Set|LSet|RSet)\s+", re.IGNORECASE)
TARGET = re.compile(r"\.?[A-Za-z_]\w*[%&!#@$]?(?:\(\))?(?:[.!][A-Za-z_]\w*[%&!#@$]?(?:\(\))?)*")
PLAIN_TARGET = re.compile(r"[A-Za-z_]\w*[%&!#@$]?")
HEAD = re.compile(r"^(?:Call\s+)?\.?[A-Za-z_]\w*[%&!#@$]?(?:\s*[.!]\s*[A-Za-z_]\w*[%&!#@$]?)*",
                  re.IGNORECASE)
CALL = re.compile(r"(?<![\w.!#])([A-Za-z_]\w*)[%&!#@$]?\s*\(")
# A name not followed by an argument list or by := (a named argument).
NAME = re.compile(r"(?<![\w.!#])([A-Za-z_]\w*)\b(?![%&!#@$]?\s*\()(?!\s*:=)")
# A ! before a name is a member access, as in Me!txtAmount; x! is a Single.
MEMBER = re.compile(r"\.\s*[A-Za-z_]|![A-Za-z_]")
NEW = re.compile(r"\bNew\s+[A-Za-z_]", re.IGNORECASE)
NUMBER_LITERAL = re.compile(r"&[HO][0-9A-F]+&?", re.IGNORECASE)
# Any operator: it computes, and so it can fail, as Like does with error 93.
OPERATOR = re.compile(r"[-+*/\\^&=<>]|\b(?:And|Eqv|Imp|Is|Like|Mod|Not|Or|Xor)\b",
                      re.IGNORECASE)
# VBA functions that run no project code, so they leave Err as it is unless
# they fail; extend the set when a copy of Err goes through another one.
PURE_FUNCTIONS = frozenset("""
    abs array asc ascb ascw atn cbool cbyte ccur cdate cdbl cdec choose chr chrb chrw cint clng
    clnglng clngptr cos csng cstr curdir cvar cvdate cverr date dateadd datediff datepart
    dateserial datevalue day ddb dir environ eof exp filedatetime filelen filter fix format
    formatcurrency formatdatetime formatnumber formatpercent freefile fv getattr getsetting hex
    hour iif input inputb instr instrb instrrev int ipmt irr isarray isdate isempty iserror
    ismissing isnull isnumeric isobject join lbound lcase left leftb len lenb loc lof log ltrim
    mid midb minute mirr month monthname now nper npv oct pmt ppmt pv qbcolor rate replace rgb
    right rightb rnd round rtrim second sgn sin sln space spc split sqr str strcomp strconv
    string strreverse switch syd tab tan time timer timeserial timevalue trim typename ubound
    ucase val vartype weekday weekdayname year
""".split())
# The same functions qualified by their library, as in VBA.CLng(x), and the
# constants of VBA so qualified, as in VBA.vbCrLf: neither a member access of
# an object nor a procedure of the project of that name.
VBA_PURE = re.compile(rf"(?<![\w.!]){VBA_QUALIFIER}(?:{'|'.join(sorted(PURE_FUNCTIONS))})\b"
                      r"[%&!#@$]?", re.IGNORECASE)
VBA_CONSTANT = re.compile(rf"(?<![\w.!]){VBA_QUALIFIER}vb\w*", re.IGNORECASE)
# A dotted name, each part with an optional index, as in Saved(i).Inner.Field or
# Module.Enum.member; an index holds no parentheses of its own.
CHAIN = re.compile(r"(?<![\w.!#])[A-Za-z_]\w*[%&!#@$]?(?:\s*\([^()]*\))?"
                   r"(?:\s*\.\s*[A-Za-z_]\w*[%&!#@$]?(?:\s*\([^()]*\))?)+")
CHAIN_PART = re.compile(r"([A-Za-z_]\w*)[%&!#@$]?|\(([^()]*)\)")
# What a variable, parameter or field is declared as: its type, in lower
# case with its qualifier (codes.payload) and "" when undeclared, and whether
# it is an array, as Name() or Name(1 To 3) declares it.
Declared = tuple[str, bool]
# A part of a dotted name, and whether an index follows it, as in Many(i).
Link = tuple[str, bool]
# The types whose value is data: reading or assigning one calls nothing, as an
# object or a Variant may through its default property.
VALUE_TYPES = frozenset("""
    boolean byte currency date decimal double integer long longlong longptr single string
""".split())
# A declared name and its type, in a variable or parameter list or a Type:
# Name, Name(1 To 3) As Long, ByRef Name As Record, WithEvents Name As Thing.
DECLARED = re.compile(r"^\s*(?:(?:Optional|ByVal|ByRef|ParamArray|WithEvents)\s+)*"
                      r"(?P<name>[A-Za-z]\w*)[%&!#@$]?\s*(?P<array>\(.*\))?\s*"
                      r"(?:As\s+(?:New\s+)?(?P<type>[A-Za-z_][\w.]*))?", re.IGNORECASE)
MODULE_VARIABLES = re.compile(r"^\s*(?P<vis>Dim|Public|Private|Global)\s+(?P<list>.*)$",
                              re.IGNORECASE)
LOCAL_VARIABLES = re.compile(r"^\s*(?P<keyword>Dim|Static|ReDim(?:\s+Preserve)?)\s+"
                             r"(?P<list>.*)$", re.IGNORECASE)
NOT_VARIABLES = re.compile(r"^\s*(?:(?:Public|Private|Global|Friend)\s+)?(?:Static\s+)?"
                           r"(?:Const|Declare|Enum|Type|Event|Sub|Function|Property|Implements)\b",
                           re.IGNORECASE)
BLOCK_START = re.compile(r"^\s*(?:(?P<vis>Public|Private)\s+)?(?P<kind>Enum|Type)\s+"
                         r"(?P<name>[A-Za-z]\w*)\s*$", re.IGNORECASE)
BLOCK_END = re.compile(r"^\s*End\s+(?:Enum|Type)\b", re.IGNORECASE)
# Prefixes of the constants of VBA (vbCrLf) and Excel (xlUp), which are values.
CONSTANT_PREFIXES = ("vb", "xl")
# Words that may stand in an expression without being a value that can fail.
WORDS = frozenset("""
    and as byref byval call case do each else elseif empty end eqv erl err error exit false for
    function gosub goto if imp in is let like loop lset me mod new next not nothing null on or
    property resume return rset select set step stop sub then to true typeof until wend while
    with xor
""".split())


@dataclass
class Segment:
    """What is known of Err at one point of a procedure's straight-line code."""

    in_handler: bool = False
    resume_next: bool = False
    # The line and the reason that made Err stale, or None while it is current.
    stale: tuple[int, str] | None = None

    @property
    def follows_last_failure(self) -> bool:
        """Whether Err belongs to the most recent statement that can fail."""
        return self.resume_next or not self.in_handler

    def restart(self, in_handler: bool | None = None) -> None:
        if in_handler is not None:
            self.in_handler = in_handler
            self.resume_next = False
        self.stale = None


@dataclass
class Procedure:
    name: str
    line: int
    header: str = ""
    units: list[tuple[int, str | None, list[str]]] = field(default_factory=list)

    def variables(self) -> dict[str, Declared]:
        """Its parameters, its own name when it holds its result (result) and
        the variables it declares, with their types; ReDim declares an array,
        keeping the type a Dim gave it."""
        found: dict[str, Declared] = {}
        if (own := self.result()) is not None:
            found[self.name.casefold()] = own
        opener = is_procedure(self.header)
        rest = self.header[opener.end():].lstrip("%&!#@$") if opener else ""
        end = closing_parenthesis(rest, 0) if rest.startswith("(") else None
        if end is not None and rest[1:end].strip():
            found.update(declared(rest[1:end]))
        for _, _, unit in self.units:
            for statement in unit:
                local = LOCAL_VARIABLES.match(statement)
                if local and local.group("keyword").casefold().startswith("redim"):
                    for name, (kind, _) in declared(local.group("list")).items():
                        found[name] = (kind or found.get(name, ("", True))[0], True)
                elif local:
                    found.update(declared(local.group("list")))
        return found

    def result(self) -> Declared | None:
        """What the own name of a Function or Property Get, the variable that
        holds its result, is declared as: its result type, or the type of its
        suffix (Fetch$), else a Variant (""); None for a Sub or a Property
        Let or Set, whose name holds nothing."""
        opener = is_procedure(self.header)
        kind = opener.group(1).casefold().split() if opener else []
        if kind not in (["function"], ["property", "get"]) or opener is None:
            return None
        rest = self.header[opener.end():]
        suffix = SUFFIX_TYPES.get(rest[:1], "")
        rest = rest[1:] if suffix else rest
        start = rest.find("(")
        end = closing_parenthesis(rest, start) if start >= 0 and not rest[:start].strip() else None
        result = RESULT_TYPE.match(rest[end + 1:] if end is not None else rest)
        if result is None:
            return suffix, False
        return result.group("type").casefold(), result.group("array") is not None

    @property
    def calls_itself_with_arguments(self) -> bool:
        """Whether its own name with an argument list, as in Fetch(True), is a
        call of itself rather than an element of its result: a Function or
        Property Get whose declared result is neither an array nor a Variant,
        which may hold one. Without a declared type, the result is a Variant,
        unless a type suffix, as in Fetch$, gives one."""
        result = self.result()
        return result is not None and not result[1] and result[0] not in {"", "variant"}


@dataclass
class Declarations:
    """What one component declares at module level that a dotted name can
    reach without a call: its Enums and their members, its user-defined
    Types and their fields, its variables and their types, and its values
    (variables, constants, Enum members), each with the subset another
    component sees. Types are lower-case names, "" when undeclared."""

    standard: bool
    enums: dict[str, frozenset[str]] = field(default_factory=dict)
    types: dict[str, dict[str, Declared]] = field(default_factory=dict)
    variables: dict[str, Declared] = field(default_factory=dict)
    public: set[str] = field(default_factory=set)
    values: frozenset[str] = frozenset()
    public_values: frozenset[str] = frozenset()
    callables: frozenset[str] = frozenset()
    public_callables: frozenset[str] = frozenset()

    def declares(self, name: str, public: bool) -> str | None:
        """What ``name`` is here: "variable", "enum", "callable" or "value"
        (a constant, an Enum member, a Type, a control); with ``public``, only
        as another component sees it: a variable, value or procedure of a
        standard module, or a public Enum of any module."""
        if public and name not in self.public and name not in self.public_values \
                and name not in self.public_callables:
            return None
        if name in self.enums:
            return "enum"
        if public and not self.standard:
            return None
        if name in self.variables:
            return "variable"
        if name in (self.public_callables if public else self.callables):
            return "callable"
        return "value" if name in (self.public_values if public else self.values) else None


def declared(items: str) -> dict[str, Declared]:
    """Each name of a declaration or parameter list, with what it is declared as."""
    found = {}
    for item in top_level_split(items):
        match = DECLARED.match(item)
        if match:
            found[match.group("name").casefold()] = \
                ((match.group("type") or "").casefold(),
                 match.group("array") is not None)
    return found


def declarations(source: str, module: Module) -> Declarations:
    """The module-level declarations of one component's source."""
    result = Declarations(module.standard, values=frozenset(module.names - module.callables),
                          public_values=frozenset(module.public - module.public_callables),
                          callables=frozenset(module.callables),
                          public_callables=frozenset(module.public_callables))
    block: tuple[str, str] | None = None
    members: dict[str, Declared] = {}
    inside = False
    for _, statement in statements(source):
        if is_procedure(statement):
            inside = True
        elif PROC_CLOSE.match(statement):
            inside = False
        elif inside:
            continue
        elif block is not None:
            if BLOCK_END.match(statement):
                if block[0] == "enum":
                    result.enums[block[1]] = frozenset(members)
                else:
                    result.types[block[1]] = members
                block, members = None, {}
            elif block[0] == "enum":
                members[statement.split("=")[0].strip().strip("[]").casefold()] = ("", False)
            else:
                members.update(declared(statement))
        elif opened := BLOCK_START.match(statement):
            block = (opened.group("kind").casefold(), opened.group("name").casefold())
            if (opened.group("vis") or "Public").casefold() == "public":
                result.public.add(block[1])
        elif (variables := MODULE_VARIABLES.match(statement)) and \
                not NOT_VARIABLES.match(statement):
            found = declared(variables.group("list"))
            result.variables.update(found)
            if variables.group("vis").casefold() in {"public", "global"}:
                result.public.update(found)
    return result


@dataclass(frozen=True)
class Scope:
    """The declarations a procedure resolves a dotted name against: its own
    variables and parameters (``local``), its component (``home``) and the
    other components of the environment (``project``, by lower-case name)."""

    home: str
    local: dict[str, Declared]
    project: dict[str, Declarations]

    def resolve(self, name: str) -> tuple[str, Declarations | None]:
        """What the unqualified ``name`` means here, as VBA resolves it, by
        scope level: a parameter, variable or the procedure's own name; then
        a declaration of the module, whatever its kind; then a public one of
        another component; then a component. The kind is "local",
        "variable", "enum", "callable", "value" or "component", with the
        component that declares it; "" when nothing does."""
        if name in self.local:
            return "local", None
        home = self.project.get(self.home)
        if home is not None and (kind := home.declares(name, False)):
            return kind, home
        public = [(kind, found) for found in self.project.values()
                  if found is not home and (kind := found.declares(name, True))]
        if len(public) > 1:
            return "ambiguous", None  # declared public in two components
        if public:
            return public[0]
        return ("component", self.project[name]) if name in self.project else ("", None)

    def unique(self, table: str, name: str) -> bool:
        """Whether one component only declares the Type (``table`` "types")
        or Enum ("enums") ``name``, so that no other one can shadow it."""
        return sum(name in getattr(found, table) for found in self.project.values()) == 1

    def variable(self, name: str, owner: str | None = None) -> Declared | None:
        """What variable ``name`` is declared as, or None when it is no
        variable or another declaration hides it; with ``owner``, only a
        variable of that component."""
        if owner is not None:
            found = self.project.get(owner)
            visible = found is not None and (owner == self.home or name in found.public)
            return found.variables.get(name) if found is not None and visible else None
        kind, found = self.resolve(name)
        if kind == "local":
            return self.local[name]
        return found.variables[name] if kind == "variable" and found is not None else None

    def owners(self, name: str, owner: str | None = None) -> list[Declarations]:
        """The components whose declaration ``name`` this procedure sees: its
        own first, then the others that make it public; with ``owner``, only
        that component."""
        homes = [owner or self.home] + ([] if owner else list(self.project))
        return [found for home in homes if (found := self.project.get(home)) is not None
                and (home == self.home or name in found.public)]

    def members(self, name: str, owner: str | None = None) -> frozenset[str] | None:
        """The members of the Enum ``name``, if this procedure sees one that
        nothing else of that name hides and no other component declares; a
        qualified name, as codes.kind, is looked up in that component only."""
        qualifier, _, name = name.rpartition(".")
        if not self.unique("enums", name):
            return None
        if qualifier:
            found = self.project.get(qualifier)
            return found.enums.get(name) if found is not None else None
        if owner is None:
            kind, found = self.resolve(name)
            return found.enums[name] if kind == "enum" and found is not None else None
        return next((found.enums[name] for found in self.owners(name, owner)
                     if name in found.enums), None)

    def fields(self, name: str) -> dict[str, Declared] | None:
        """The fields of the user-defined Type ``name``, if this procedure sees
        one and no other component declares one of that name; a qualified
        name, as codes.payload, is looked up in that component only, and
        VBA.Collection or Excel.Range, which name no component, are no Type."""
        qualifier, _, name = name.rpartition(".")
        if not self.unique("types", name):
            return None
        if qualifier:
            found = self.project.get(qualifier)
            return found.types.get(name) if found is not None else None
        return next((found.types[name] for found in self.owners(name) if name in found.types),
                    None)

    def walk(self, declaration: Declared, links: list[Link]) -> Declared | None:
        """What the dotted name ``links`` reaches, from its first part,
        declared as ``declaration``, through the fields of user-defined Types;
        None when a part is not such a field. An index is an element only on
        an array: on anything else, an object or a Variant, it may call a
        default property, as Owner(1) does on a Collection."""
        for position, (_, indexed) in enumerate(links):
            kind, array = declaration
            if indexed != array and (indexed or position < len(links) - 1):
                return None
            if position == len(links) - 1:
                return kind, array and not indexed
            fields = self.fields(kind)
            if fields is None or links[position + 1][0] not in fields:
                return None
            declaration = fields[links[position + 1][0]]
        return None

    def record(self, links: list[Link]) -> bool:
        """Whether ``links`` name a variable of a user-defined Type, or a
        field of one that is itself one, as With Saved needs."""
        declaration = self.variable(links[0][0])
        reached = self.walk(declaration, links) if declaration else None
        return reached is not None and not reached[1] and self.fields(reached[0]) is not None

    def data(self, reached: Declared | None) -> bool:
        """Whether a part declared as ``reached`` is read or assigned without
        a call: a value type, an Enum or a user-defined Type, whose copy runs
        no code. An object, a class or a Variant is not: its value is its
        default property, and even Set may create it, if declared As New."""
        if reached is None:
            return False
        kind = reached[0]
        return kind in VALUE_TYPES or self.members(kind) is not None \
            or self.fields(kind) is not None

    def value(self, links: list[Link], names: Names) -> bool:
        """Whether the dotted name ``links`` reads a value without a call: a
        field of a variable of a user-defined Type (Saved.Number,
        Saved(i).Inner.Field), an Enum member (Kind.kOne), or a value of a
        standard module named by it (Module.LIMIT, Module.Kind.kOne,
        Module.Saved.Number). The part it reaches must hold data (data).
        Anything else is a member access."""
        (root, indexed), rest = links[0], links[1:]
        kind, home = self.resolve(root)
        declaration = self.variable(root)
        if declaration is not None:
            return self.data(self.walk(declaration, links))
        if kind in {"callable", "value", "local", "variable"} or indexed:
            return False
        if kind == "component" and home is not None and home.standard and rest:
            declaration = self.variable(rest[0][0], root)
            if declaration is not None:
                return self.data(self.walk(declaration, rest))
            if len(rest) == 1:
                own = root == self.home
                return not rest[0][1] and rest[0][0] in (home.values if own else home.public_values)
            members = self.members(rest[0][0], root)
            return len(rest) == 2 and members is not None and not any(
                link[1] for link in rest) and rest[1][0] in members
        members = self.members(root)
        return len(rest) == 1 and members is not None and not rest[0][1] and rest[0][0] in members


def links(text: str) -> list[Link]:
    """The parts of the dotted name ``text``, each with whether an index follows it."""
    parts = CHAIN_PART.findall(text)
    return [(name.casefold(), position + 1 < len(parts) and bool(parts[position + 1][1]))
            for position, (name, _) in enumerate(parts) if name]


@dataclass(frozen=True)
class Names:
    """The names a procedure resolves before VBA's own: ``values``, read
    without a call, and ``callables``, the procedures, properties and Declares
    of the project it can call unqualified. A project declaration shadows a
    VBA function, statement or constant of the same name, as in VBA."""

    values: frozenset[str] = frozenset()
    callables: frozenset[str] = frozenset()
    # Values whose name with an argument list calls a procedure: the
    # procedure's own name when it calls itself, as in Fetch(True).
    calls: frozenset[str] = frozenset()
    # The declarations that dotted names resolve against, if known.
    scope: Scope | None = None

    def value_chain(self, text: str) -> list[str] | None:
        """The indices of the dotted name ``text`` when it reads a value
        without a call (Scope.value), or None."""
        chain = links(text)
        if self.scope is None or len(chain) < 2 or not self.scope.value(chain, self):
            return None
        return [index for _, index in CHAIN_PART.findall(text) if index]


@dataclass(frozen=True)
class Effect:
    """Whether code runs anything, so that it can fail and leave its own
    error in Err after On Error Resume Next, and whether it can run code that
    changes Err in a handler: a call or a member access."""

    fails: bool = False
    changes: bool = False

    def __or__(self, other: Effect) -> Effect:
        return Effect(self.fails or other.fails, self.changes or other.changes)


def library_qualifier(match: re.Match[str] | None, names: Names) -> bool:
    """Whether ``match``, of a pattern built on VBA_QUALIFIER, names VBA's
    library: always after VBA., and for a module name alone, such as
    Strings, only when no value, component or procedure of the project has
    that name."""
    if match is None:
        return False
    module = match.group("module")
    return module is None or module.casefold() not in names.values | names.callables


def effect(text: str, names: Names) -> Effect:
    """What evaluating the expression ``text`` can do to Err; strings are blanked.

    A name followed by parentheses is an array element when it is a value,
    and a call otherwise; a bare name that is not a value is a call too: a
    function or property of the project, a Declare, or a member of the host's
    global object, such as ActiveSheet, called without parentheses. Only a
    name the project does not declare can be a VBA function or constant."""
    text = ERR_PROPERTY.sub(lambda match: "0" * len(match.group()), text)
    text = TYPEOF_TYPE.sub(r"\g<1>", text)

    def library(match: re.Match[str]) -> str:
        return "0" * len(match.group()) if library_qualifier(match, names) else match.group()

    # A qualified VBA function can fail, as CLng("x") does, but runs no project code.
    qualified = any(library_qualifier(match, names) for match in VBA_PURE.finditer(text))
    text = VBA_CONSTANT.sub(library, VBA_PURE.sub(library, text))
    # A dotted name that reads a value, as Saved.Number or Kind.kOne, is read
    # like a variable: only its indices are expressions.
    indices: list[str] = []
    chains = 0

    def value(match: re.Match[str]) -> str:
        nonlocal chains
        found = names.value_chain(match.group())
        if found is None:
            return match.group()
        indices.extend(found)
        chains += 1
        return "0" * len(match.group())

    text = CHAIN.sub(value, text)
    text = NUMBER_LITERAL.sub(lambda match: "0" * len(match.group()), dates_as_values(text))
    # A literal, a constant and a property of Err run nothing; any other
    # name, call or operator runs code that can fail.
    result = Effect(fails=qualified or chains > 0 or OPERATOR.search(text) is not None)
    for index in indices:
        result = result | effect(index, names)
    if MEMBER.search(text) or NEW.search(text):
        result = Effect(True, True)
    for match in CALL.finditer(text):
        name = match.group(1).casefold()
        if name in names.calls:
            result = Effect(True, True)
        elif name in names.values:
            result = result | Effect(fails=True)
        elif name in names.callables:
            result = Effect(True, True)
        elif name not in WORDS:
            result = result | Effect(True, name not in PURE_FUNCTIONS)
    for match in NAME.finditer(text):
        name = match.group(1).casefold()
        if name in names.callables and name not in names.values:
            result = Effect(True, True)
        elif not (name in WORDS or name.startswith(CONSTANT_PREFIXES)):
            result = result | Effect(True, name not in names.values | PURE_FUNCTIONS)
    return result


def assignment(statement: str) -> tuple[str, str] | None:
    """``(target, value)`` of an assignment, or None for any other statement."""
    keyword = ASSIGNMENT_KEYWORD.match(statement)
    body = statement[keyword.end():] if keyword else statement
    depth = 0
    for index, char in enumerate(body):
        depth += {"(": 1, ")": -1}.get(char, 0)
        if char == "=" and not depth:
            target = body[:index].strip()
            collapsed, position = "", 0
            while position < len(target):
                end = closing_parenthesis(target, position) if target[position] == "(" else None
                collapsed += "()" if end is not None else target[position]
                position = end + 1 if end is not None else position + 1
            return (target, body[index + 1:]) if TARGET.fullmatch(collapsed) else None
    return None


def statement_effect(statement: str, names: Names) -> tuple[str, Effect]:
    """The expressions a statement evaluates before its own operation, in
    evaluation order, and what the whole statement can do to Err. A procedure
    runs after its arguments are read, and a property is set after its value
    is read. Every statement runs code that can fail, even Number = 1 with a
    type mismatch, except a declaration, With Err and End With, and a write
    of a property of Err, as in Err.Description = x, which reads and clears
    nothing of it and fails only when its value runs code."""
    if DECLARATION.match(statement) or END_WITH.match(statement) or WITH_ERR.match(statement):
        return "", Effect()
    if WITH.match(statement):
        return statement[4:], effect(statement[4:], names) | Effect(fails=True)
    pair = assignment(statement)
    if pair is not None:
        target, value = pair
        if ERR_TARGET.fullmatch(target):
            return value, effect(value, names)
        index = target.find("(")
        reads = value + (" , " + target[index:] if index > 0 else "")
        return reads, effect(value, names) | target_effect(target, names) | Effect(fails=True)
    # A procedure of the project named like a VBA statement, such as Beep,
    # shadows it, unless the statement is qualified, as in VBA.Beep. Call
    # SaveSetting(...) is SaveSetting ... written as a call.
    called = CALL_KEYWORD.match(statement)
    body = statement[called.end():] if called else statement
    qualifier = VBA_HEAD.match(body)
    if not library_qualifier(qualifier, names):
        qualifier = None
    unqualified = body[qualifier.end():] if qualifier else body
    builtin = BUILTIN_STATEMENT.match(unqualified)
    if builtin and (qualifier or unqualified.split(None, 1)[0].casefold() not in names.callables):
        operands = STATEMENT_SYNTAX.sub(" ", unqualified[builtin.end():])
        if builtin.group().casefold().startswith("redim"):
            operands = ",".join(TYPE_CLAUSE.sub("", item) for item in top_level_split(operands))
        return operands, effect(operands, names) | Effect(fails=True)
    head_match = HEAD.match(statement)
    arguments = statement[head_match.end():] if head_match else statement
    return arguments, effect(arguments, names) | Effect(True, True)


def evaluated_before(text: str, position: int) -> str:
    """The part of the expression ``text`` that VBA evaluates before the
    operand at ``position``: the text before it, without the calls that it is
    an argument of and the operators that it is an operand of, which run
    after it. An operator earlier in the text may also run after it, as & in
    a & b / Err.Number; it is kept, so that a read is never claimed stale on
    the strength of a guess."""
    before = list(text[:position])
    opened: list[int] = []
    for index, char in enumerate(text[:position]):
        if char == "(":
            opened.append(index)
        elif char == ")" and opened:
            opened.pop()
    for index in opened:
        call = ENCLOSING_CALL.search(text[:index])
        start = call.start() if call else index
        before[start:index + 1] = " " * (index + 1 - start)
    result = "".join(before).rstrip()
    # An operator, or the name of a named argument (Number:=), runs after it.
    while (operator := TRAILING_OPERATOR.search(result) or NAMED_ARGUMENT.search(result)):
        result = result[:operator.start()].rstrip()
    return result


def target_effect(target: str, names: Names) -> Effect:
    """What assigning to ``target`` can do to Err, once the value is read.

    A variable of the procedure or the module (and the procedure's own name)
    is a plain copy, an array element can fail, and a statement target such
    as Date or Mid(Text, 1, 1) can fail. Any other target is a property: a
    Property Let or Set of the project, which shadows a statement target of
    its name, a member (x.y = 1, .y = 1 in a With block), or a property of
    the host, such as Caption, which runs code. VBA.Mid(Text, 1, 1) is Mid,
    and a field of a variable of a user-defined Type, as in Saved.Number or
    Saved(i).Number, or a value of a standard module, is a variable."""
    if (chain := CHAIN.fullmatch(target.strip())) is not None:
        chain_indices = names.value_chain(chain.group())
        if chain_indices is not None:
            found = Effect()
            for chain_index in chain_indices:
                found = found | effect(chain_index, names) | Effect(fails=True)
            return found
    qualifier = VBA_HEAD.match(target)
    if qualifier and library_qualifier(qualifier, names):
        rest = target[qualifier.end():]
        index = rest.find("(")
        if (rest[:index] if index > 0 else rest).rstrip("%&!#@$ ").casefold() in BUILTIN_TARGETS:
            return (effect(rest[index:], names) if index > 0 else Effect()) | Effect(fails=True)
    index = target.find("(")
    name = (target[:index] if index > 0 else target).rstrip("%&!#@$ ").casefold()
    indices = effect(target[index:], names) if index > 0 else Effect()
    if PLAIN_TARGET.fullmatch(target) or (index > 0 and PLAIN_TARGET.fullmatch(name)):
        if name in names.values:
            return indices | Effect(fails=index > 0)
        if name in BUILTIN_TARGETS and name not in names.callables:
            return indices | Effect(fails=True)
    return indices | Effect(True, True)


def handler_labels(procedure: Procedure) -> set[str]:
    """The labels that an On Error GoTo of the procedure names."""
    return {match.group(1).casefold() for _, _, statements in procedure.units
            for statement in statements for match in HANDLER.finditer(statement)
            if match.group(1) not in {"0"}}


def stale_message(procedure: str, reads: list[re.Match[str]], cause: str) -> str:
    fields = list(dict.fromkeys(f"Err.{match.group(1).capitalize()}" for match in reads))
    return (f"{', '.join(fields)} read in {procedure} after {cause}; "
            "save Err.Number, Err.Source and Err.Description into variables first")


def stale_reads(reads: str, segment: Segment, names: Names) -> list[re.Match[str]]:
    """The reads of Err in ``reads`` that are stale, each judged on what the
    statement evaluates before it as well as on ``segment``.

    After On Error Resume Next or outside a handler, code that runs before a
    read and can fail may be the error it checks: the read is stale only when
    Err already was and no such code runs. In a handler, a read is stale when
    Err already was, or when a call or member access runs before it in the
    same statement, as in Report Helper(), Err.Number. A call that a read is
    an argument of, as in CLng(Err.Number), runs after it."""
    stale = []
    for match in ERR_READ.finditer(reads):
        before = effect(evaluated_before(reads, match.start()), names)
        if segment.follows_last_failure:
            if segment.stale and not before.fails:
                stale.append(match)
        elif segment.stale or before.changes:
            stale.append(match)
    return stale


def judge(path: str, procedure: Procedure, line: int, reads: str, total: Effect,
          segment: Segment, names: Names) -> list[Finding]:
    """Report the stale reads of ``reads``, then apply ``total`` to ``segment``."""
    findings: list[Finding] = []
    stale = stale_reads(reads, segment, names)
    if stale:
        cause = (f"line {segment.stale[0]}, {segment.stale[1]}" if segment.stale else
                 "a call or member access that the same statement evaluates first")
        findings.append(Finding(path, line, stale_message(procedure.name, stale, cause)))
    if segment.follows_last_failure:
        if total.fails:
            segment.stale = None
    elif total.changes and segment.stale is None:
        segment.stale = (line, "which may have changed Err with a call or member access")
    return findings


def check_statement(path: str, procedure: Procedure, line: int, statement: str,
                    segment: Segment, names: Names) -> list[Finding]:
    """Read one statement against ``segment``, then apply its effect on Err."""
    if JOIN.match(statement):
        segment.restart()
        return []
    if ON_ERROR.match(statement):
        segment.stale = (line, "which cleared Err with an On Error statement")
        segment.resume_next = bool(RESUME_NEXT.match(statement))
        return []
    if ERR_CLEAR.match(CALL_KEYWORD.sub("", statement, count=1)):
        segment.stale = (line, "which cleared Err with Err.Clear")
        return []
    if BRANCH.match(statement):
        # Its expression runs first; what follows is reached from elsewhere too.
        single, loop = SINGLE_LINE_IF.match(statement), FOR_HEAD.match(statement)
        reads = (single.group("condition") if single else
                 statement[loop.end():] if loop else statement)
        findings = judge(path, procedure, line, reads, Effect(), segment, names)
        segment.restart()
        return findings
    reads, total = statement_effect(statement, names)
    return judge(path, procedure, line, reads, total, segment, names)


def split_else(tail: str) -> tuple[str, str | None]:
    """The statements after Then of a single-line If, and those after its
    Else, if any. An Else belongs to the nearest If before it that has none,
    as in If a Then If b Then x Else y, where y runs when b is False."""
    open_ifs = 0
    for match in IF_OR_ELSE.finditer(tail):
        if match.group().casefold() == "if":
            open_ifs += 1
        elif open_ifs:
            open_ifs -= 1
        else:
            return tail[:match.start()], tail[match.end():]
    return tail, None


def check_single_line_if(path: str, procedure: Procedure, line: int, text: str,
                         segment: Segment, names: Names) -> list[Finding]:
    """Read a single-line If: its condition, then each branch as the
    statements it runs, from the state after the condition. The branches
    exclude each other, and the segment ends after the If."""
    match = SINGLE_LINE_IF.match(text)
    if match is None:
        return check_statement(path, procedure, line, text, segment, names)
    condition = match.group("condition")
    findings = judge(path, procedure, line, condition, effect(condition, names), segment, names)
    for branch in split_else(match.group("tail")):
        if branch is not None:
            findings += check_statements(path, procedure, line, split_statements(branch),
                                         replace(segment), names, [])
    segment.restart()
    return findings


def with_object(statement: str, names: Names) -> str | None:
    """What .x means inside the With block ``statement`` opens, when it is
    data rather than a member access: Err, as in With Err, or a variable of
    a user-defined Type, as in With Saved or With Saved(i).Inner."""
    if WITH_ERR.match(statement):
        return "Err."
    target = statement[4:].strip()
    chain = CHAIN.fullmatch(target) or re.fullmatch(r"[A-Za-z_]\w*(?:\s*\([^()]*\))?", target)
    if chain and names.scope is not None and names.scope.record(links(target)):
        return target + "."
    return None


def check_statements(path: str, procedure: Procedure, line: int, statements: list[str],
                     segment: Segment, names: Names, with_objects: list[str | None]
                     ) -> list[Finding]:
    """Read the statements of one line in order; ``with_objects`` holds, for
    each open With block, what its .x reads (with_object), if it is data."""
    findings: list[Finding] = []

    def resolved(statement: str) -> str:
        if with_objects and with_objects[-1]:
            return WITH_MEMBER.sub(with_objects[-1], statement)
        return statement

    for index, statement in enumerate(statements):
        statement = resolved(statement)
        single = SINGLE_LINE_IF.match(statement)
        if single and (single.group("tail").strip() or index < len(statements) - 1):
            # The rest of the line belongs to the If's branches.
            rest = [resolved(item) for item in statements[index + 1:]]
            return findings + check_single_line_if(path, procedure, line,
                                                   " : ".join([statement, *rest]), segment,
                                                   names)
        findings += check_statement(path, procedure, line, statement, segment, names)
        if WITH.match(statement):
            with_objects.append(with_object(statement, names))
        elif END_WITH.match(statement) and with_objects:
            with_objects.pop()
    return findings


def check_procedure(path: str, procedure: Procedure, names: Names) -> list[Finding]:
    handlers = handler_labels(procedure)
    segment = Segment()
    findings: list[Finding] = []
    with_objects: list[str | None] = []
    for line, label, unit in procedure.units:
        if label is not None:
            segment.restart(in_handler=label.casefold() in handlers)
        findings += check_statements(path, procedure, line, unit, segment, names, with_objects)
    return findings


def procedures(source: str) -> list[Procedure]:
    """Each procedure's logical lines: ``(line, label, statements)``, strings blanked."""
    found: list[Procedure] = []
    current: Procedure | None = None
    for line, _, kind, code in logical_units(source.split("\n"), blank_strings=True):
        if kind == "directive":
            continue
        label = label_at_start(code) if current is not None else None
        if label is not None:
            prefix = LABEL_PREFIX.match(code)
            code = code[prefix.end():] if prefix else code
        owner, statements = current, []
        for statement in split_statements(code):
            opener = is_procedure(statement)
            if opener:
                current = owner = Procedure(opener.group(2), line, statement)
                found.append(current)
            elif PROC_CLOSE.match(statement):
                current = None
            elif current is not None and statement:
                statements.append(statement)
        if owner is not None and (statements or label is not None):
            owner.units.append((line, label, statements))
    return found


def procedure_names(module: Module, project: Names, components: frozenset[str],
                    line: int) -> Names:
    """The names the procedure at ``line`` resolves, by scope level as VBA
    does, each to what the first level that declares it makes it: its
    parameters, variables and own name (its result) are values; then the
    module's variables, constants and Enum members are values and its
    procedures, properties and Declares are callables; then the public ones
    of the standard modules, in ``project``; then ``components``, values."""
    procedure = next((item for item in module.procedures if item.line == line), None)
    seen = set(procedure.names | {procedure.name.casefold()} if procedure else set())
    values: set[str] = set(seen)
    callables: set[str] = set()
    for level_values, level_callables in ((module.names - module.callables, module.callables),
                                          (project.values, project.callables)):
        values |= level_values - seen
        callables |= level_callables - seen
        seen |= level_values | level_callables
    return Names(frozenset(values | (components - seen)), frozenset(callables))


def analyze(modules: dict[str, str]) -> tuple[list[Finding], int]:
    """The findings of one environment's sources, and how many procedures it has."""
    scanned = {path: scan(path, source) for path, source in modules.items()}
    standard = [module for module in scanned.values() if module.standard]
    components = {Path(path).stem.casefold(): declarations(source, scanned[path])
                  for path, source in modules.items()}
    project = Names(
        frozenset(name for module in standard for name in module.public - module.public_callables),
        frozenset(name for module in standard for name in module.public_callables))
    stems = frozenset(Path(path).stem.casefold() for path in modules)
    findings: list[Finding] = []
    count = 0
    for path, source in modules.items():
        for procedure in procedures(source):
            count += 1
            names = procedure_names(scanned[path], project, stems, procedure.line)
            names = replace(names, scope=Scope(Path(path).stem.casefold(),
                                               procedure.variables(), components))
            if procedure.calls_itself_with_arguments:
                names = replace(names, calls=frozenset({procedure.name.casefold()}))
            findings += check_procedure(path, procedure, names)
    return findings, count


def check(root: Path) -> tuple[list[Finding], int, int]:
    components = tracked_components(root)
    by_environment: dict[str, dict[str, str]] = {}
    for path in components:
        # Unreadable directives are vba-conditionals findings, not repeated here.
        sources, _ = reachable_sources(path, read_component(root, path))
        for environment, source in sources.items():
            by_environment.setdefault(environment, {})[path] = source
    environments: dict[Finding, list[str]] = {}
    count = 0
    for environment, modules in by_environment.items():
        found, procedures_seen = analyze(modules)
        count = max(count, procedures_seen)
        for finding in found:
            names = environments.setdefault(finding, [])
            if environment not in names:
                names.append(environment)
    findings = [Finding(item.path, item.line, f"{item.message} [{', '.join(names)}]")
                for item, names in environments.items()]
    return findings, len(components), count


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components, count = check(args.root.resolve())
    return report("vba-errors", findings, f"{components} VBA component(s), {count} procedure(s)")


if __name__ == "__main__":
    sys.exit(main())
