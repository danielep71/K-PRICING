#!/usr/bin/env python3
"""Check the dependency direction of docs/REPOSITORY_STRUCTURE.md, Dependencies.

src/core/ and src/classes/ must stay host-independent: no Excel object model,
worksheet, range or UI, no file, registry or process access, and no dynamic
call whose target no check can see; neither may name a facade, adapter, form,
test or example component. src/core/ standard modules also hold no ambient
mutable run state: no module-level or Static variable. examples/modules/ use
only the supported facade. These are name-based static rules over every
#If branch at once, and the state rules in each compilation environment of
check_vba_conditionals.py; they are not a resolution of calls; the *Static
checks* section of docs/VBA_HOUSE_STYLE.md lists their limits.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Sequence
from pathlib import Path

from check_vba_conditionals import reachable_sources
from check_vba_jumps import PROC_CLOSE, is_procedure
from vba_lexer import (
    STRING,
    Finding,
    facade_prefix,
    logical_units,
    read_component,
    report,
    split_statements,
    tracked_components,
)

ROOT = Path(__file__).resolve().parents[1]
CORE = "src/core/"
CLASSES = "src/classes/"
EXAMPLES = "examples/modules/"
INDEPENDENT = (CORE, CLASSES)

# Excel and UI globals; an unqualified use reaches the host. Reusing one of
# these names as a local identifier is rejected too, since it shadows the host.
HOST_NAMES = ("Application", "ActiveWorkbook", "ActiveSheet", "ActiveCell", "ActiveChart",
              "ActiveWindow", "ThisWorkbook", "Workbooks", "Worksheets", "Sheets", "Charts",
              "Range", "Cells", "Rows", "Columns", "Selection", "Names", "WorksheetFunction",
              "Excel", "MSForms", "UserForms", "MsgBox", "InputBox", "DoEvents")
HOST = re.compile(rf"(?<![\w.])({'|'.join(HOST_NAMES)})\b", re.I)
HOST_TYPE = re.compile(r"\bAs\s+(?:New\s+)?(Worksheet|Workbook|Chart|Window|Shape|ListObject|"
                       r"PivotTable|Name|UserForm|Control)\b", re.I)
# Bracketed expressions such as [A1] are evaluated by Excel; an Enum's hidden
# member such as [_First] is not.
BRACKET = re.compile(r"\[(?!_\w*\])")
# VBA's file, registry, process, form and system-clock statements, read at
# the start of a statement once labels and VBA. qualifiers are removed.
FILE_STATEMENT = re.compile(r"^(?:Open\b.*\bFor\b|Kill|FileCopy|MkDir|RmDir|ChDir|ChDrive|SetAttr|"
                            r"Reset|SaveSetting|DeleteSetting|SendKeys|AppActivate|Load|Unload|"
                            r"Name\b.*\bAs)\b|"
                            r"^(?:Print|Input|Line\s+Input|Write|Width)\s*#|^Close\b|"
                            r"^(?:Get|Put|Seek|Lock|Unlock)\b\s*[#\w(]|^(?:Date|Time)\$?\s*=", re.I)
# VBA's file, registry, environment and process functions. Those that are
# also ordinary words match only as calls.
FILE_FUNCTION = re.compile(r"(?<![\w.])(?:(Dir|Shell|GetSetting|GetAllSettings|Environ|CurDir|"
                           r"FreeFile|FileLen|FileDateTime|GetAttr|Command)\b\$?|"
                           r"(FileAttr|EOF|LOF|Loc|Seek|Input|InputB)\$?\s*\()", re.I)
# Native libraries and external object models: a Declare, a Scripting type
# other than Dictionary, or a library-qualified name.
DECLARE = re.compile(r"^(?:(?:Public|Private|Global)\s+)?Declare\b", re.I)
EXTERNAL = re.compile(r"(?<![\w.])(Scripting\.(?!Dictionary\b)\w+|(?:ADODB|DAO|MSXML2|WinHttp|"
                      r"WScript|IWshRuntimeLibrary|Shell32|MSHTML|Outlook|Word|PowerPoint|Access|"
                      r"Office)\.\w+)", re.I)
# Scripting Runtime file-system types, declared or created with a standalone
# New, as in Set Fso = New FileSystemObject.
EXTERNAL_TYPE = re.compile(r"\b(?:As\s+(?:New\s+)?|New\s+)(FileSystemObject|TextStream|Drive|"
                           r"Folder|File)\b", re.I)
# A numeric line label, an explicit Call, and the VBA library qualifiers that
# reach a built-in under another name, as in Call VBA.FileSystem.Kill(Path).
LINE_LABEL = re.compile(r"^\s*\d+(?=\s|$)\s*")
CALL = re.compile(r"^Call\s+", re.I)
VBA_QUALIFIER = re.compile(r"(?<![\w.])(?:VBA\.)?(?:FileSystem|Interaction|Strings|Information|"
                           r"Conversion|DateTime|Math|Financial)\.|(?<![\w.])VBA\.", re.I)
DYNAMIC = re.compile(r"(?<![\w.])(CallByName|Evaluate|ExecuteExcel4Macro|AddressOf)\b", re.I)
# CreateObject and GetObject in any call form: with or without parentheses,
# as VBA allows for a call whose result is not used, with the literal argument
# only when it is the whole argument list: a concatenated ProgID or a server
# name leaves the object type to run time.
OBJECT = re.compile(r'(?<![\w.])(CreateObject|GetObject)\b\s*\(?\s*(?:"([^"]*)"\s*(?:\)|$))?',
                    re.I)
ALLOWED_OBJECTS = frozenset({"scripting.dictionary"})
# Components that src/core/ and src/classes/ may not name: the facade and the
# host adapters above them, and the tests and examples beside them. VBA
# resolves names ignoring case, so tracked component names, the facade prefix
# and TEST_ match in any case; the M_, EX_, sh and frm prefixes match only as
# the house style writes them, since in lower case they are ordinary
# identifiers such as m_Total. The facade prefix comes from the house style;
# while it is the <REGIME> placeholder, only tracked facade names are known.
UPPER_HOMES = ("src/modules/", "src/workbook/", "src/forms/", "tests/modules/", EXAMPLES)
UPWARD_ANY_CASE = re.compile(r"(?<![\w.])(TEST_\w+)", re.I)
UPWARD = re.compile(r"(?<![\w.])(M_\w+|EX_\w+|sh[A-Z]\w*|frm[A-Z]\w*)")
# Components that examples may not name: everything but the facade. Model
# classes match only by their tracked names: a C prefix pattern would also
# match ordinary identifiers such as Count.
HIDDEN_HOMES = (CORE, CLASSES, "src/workbook/", "src/forms/", "tests/modules/")
HIDDEN_ANY_CASE = re.compile(r"(?<![\w.])((?:CORE|TEST)_\w+)", re.I)
HIDDEN = re.compile(r"(?<![\w.])(M_\w+|sh[A-Z]\w*|frm[A-Z]\w*)")
# The statements after Then and Else of a single-line If, as in
# If Flag Then Kill Path Else Exit Sub.
SINGLE_LINE_IF = re.compile(r"^If\b.*?\bThen\b\s*(\S.*)$", re.I)
ELSE = re.compile(r"\bElse\b", re.I)
MODULE_VARIABLE = re.compile(r"^(?:Dim|Private|Public|Global|Static)\s+(?!(?:Const|Declare|Enum|"
                             r"Type|Sub|Function|Property|Event)\b)(?:WithEvents\s+)?[A-Za-z_]",
                             re.I)
STATIC = re.compile(r"^Static\s+[A-Za-z_]", re.I)
# The members of a Type or Enum block are names, not statements: Close As
# Double is a field and Reset = 1 a member, also in brackets as [Close]. Only
# what follows the name is checked.
BLOCK_OPEN = re.compile(r"^(?:(?:Public|Private)\s+)?(?:Type|Enum)\s+[A-Za-z_]", re.I)
BLOCK_CLOSE = re.compile(r"^End\s+(?:Type|Enum)\b", re.I)
MEMBER_NAME = re.compile(r"^(?:[A-Za-z_]\w*|\[[^\]]*\])")


def names(components: Sequence[str]) -> re.Pattern[str] | None:
    """Any of these component names, ignoring case, or None for no component."""
    stems = sorted({Path(path).stem for path in components}, key=len, reverse=True)
    return re.compile(rf"(?<![\w.])({'|'.join(map(re.escape, stems))})\b", re.I) if stems \
        else None


def references(code: str, patterns: Sequence[re.Pattern[str] | None]) -> list[str]:
    """Distinct names matched by any of ``patterns``, in order of appearance."""
    found = sorted((match.start(), match.group(1)) for pattern in patterns if pattern
                   for match in pattern.finditer(code))
    return list(dict.fromkeys(name for _, name in found))


def executable_parts(code: str) -> list[str]:
    """The statement and, for a single-line If, the statements after Then and Else."""
    parts: list[str] = []
    pending = [code.strip()]
    while pending:
        for piece in (normalize(item) for item in ELSE.split(pending.pop())):
            parts.append(piece)
            branch = SINGLE_LINE_IF.match(piece)
            if branch:
                pending.append(branch.group(1))
    return parts


def object_calls(raw: str) -> list[re.Match[str]]:
    """CreateObject and GetObject calls outside string literals."""
    literals = [match.span() for match in STRING.finditer(raw)]
    return [match for match in OBJECT.finditer(raw)
            if not any(start <= match.start() < end for start, end in literals)]


def upward_any_case(prefix: str | None) -> re.Pattern[str]:
    """Names above core in any case: the facade prefix, when set, and TEST_."""
    if not prefix:
        return UPWARD_ANY_CASE
    return re.compile(rf"(?<![\w.])((?:{re.escape(prefix)}|TEST)_\w+)", re.I)


def host_independent(path: str, line: int, code: str, raw: str,
                     upper: re.Pattern[str] | None,
                     upward: re.Pattern[str] = UPWARD_ANY_CASE) -> list[Finding]:
    """Host, file, process, dynamic-call and upward references in one statement."""
    findings: list[Finding] = []

    def add(message: str) -> None:
        findings.append(Finding(path, line, message))

    for match in HOST.finditer(code):
        add(f"{match.group(1)} reaches the Excel host; core and classes are host-independent")
    for match in HOST_TYPE.finditer(code):
        add(f"As {match.group(1)} is an Excel or UI type; core and classes are host-independent")
    if BRACKET.search(STRING.sub('""', raw)):
        add("a [...] expression is evaluated by Excel; core and classes are host-independent")
    if any(FILE_STATEMENT.match(part) for part in executable_parts(code)) or \
            FILE_FUNCTION.search(code):
        add("file, registry or process access; core and classes do no I/O")
    if DECLARE.match(code):
        add("a native Declare; core and classes call no external library")
    for match in EXTERNAL.finditer(code):
        add(f"{match.group(1)} is an external object model; core and classes are "
            "host-independent")
    for match in EXTERNAL_TYPE.finditer(code):
        add(f"{match.group(1)} is a file-system type; core and classes do no I/O")
    for match in DYNAMIC.finditer(code):
        add(f"{match.group(1)}: a dynamic call that no check can resolve")
    for match in object_calls(raw):
        if match.group(1).casefold() == "getobject" or \
                (match.group(2) or "").casefold() not in ALLOWED_OBJECTS:
            add(f"{match.group(1)} creates an object no check can see; only "
                'CreateObject("Scripting.Dictionary") is allowed')
    for name in references(code, (upper, upward, UPWARD)):
        add(f"{name} names a facade, host, test or example component; core and "
            "classes depend on nothing above them")
    return findings


def statements(text: str) -> list[tuple[int, str, str]]:
    """``(line, code, raw)`` per statement, code with strings blanked; directives
    are skipped, so every #If branch is read."""
    lines = text.split("\n")
    blanked = logical_units(lines, blank_strings=True)
    kept = logical_units(lines)
    result: list[tuple[int, str, str]] = []
    for (line, _, kind, code), (_, _, _, raw) in zip(blanked, kept, strict=True):
        if kind == "directive" or code.lstrip().startswith(("Attribute ", "Option ")):
            continue
        result.extend((line, normalize(part), normalize(raw_part)) for part, raw_part
                      in zip(split_statements(code), split_statements(raw), strict=False))
    return result


def normalize(statement: str) -> str:
    """The statement without a numeric line label, an explicit Call or VBA
    library qualifiers, so that 100 Kill Path, Call Kill(Path) and
    VBA.FileSystem.Kill Path all read as a Kill statement."""
    return VBA_QUALIFIER.sub("", CALL.sub("", LINE_LABEL.sub("", statement).strip()))


def analyze(path: str, text: str, upper: re.Pattern[str] | None = None,
            hidden: re.Pattern[str] | None = None,
            upward: re.Pattern[str] = UPWARD_ANY_CASE) -> list[Finding]:
    """Findings for one component; ``upper`` and ``hidden`` match the tracked
    components that core and classes, or examples, may not name."""
    findings: list[Finding] = []
    member = False
    for line, statement, raw_statement in statements(text):
        code, raw = statement, raw_statement
        # A procedure header also ends a block, so that a missing End Type
        # cannot hide the statements after it.
        if BLOCK_CLOSE.match(code) or is_procedure(code):
            member = False
        elif member:
            code, raw = MEMBER_NAME.sub("", code), MEMBER_NAME.sub("", raw)
        elif BLOCK_OPEN.match(code):
            member = True
        if path.startswith(INDEPENDENT):
            findings.extend(host_independent(path, line, code, raw, upper, upward))
        elif path.startswith(EXAMPLES):
            findings.extend(Finding(path, line, f"{name} is behind the facade; examples use only "
                                                "the supported API")
                            for name in references(code, (hidden, HIDDEN_ANY_CASE, HIDDEN)))
    if path.startswith(CORE) and path.lower().endswith(".bas"):
        # Procedure scope depends on which #If branches compile, so the state
        # rules read each compilation environment on its own: a procedure
        # opened in one branch must not hide a declaration in another.
        for source in reachable_sources(path, text)[0].values():
            findings.extend(core_state(path, source))
    return list(dict.fromkeys(findings))


def core_state(path: str, source: str) -> list[Finding]:
    """Module-level and Static state in one environment's source of a core module."""
    findings: list[Finding] = []
    inside = False
    for line, code, _ in statements(source):
        header = is_procedure(code)
        if header and re.match(r"^\s*(?:\w+\s+)?Static\b", code, re.I):
            findings.append(Finding(path, line, f"Static procedure {header.group(2)} keeps state "
                                                "between calls; core holds no run state"))
        elif not inside and MODULE_VARIABLE.match(code.strip()):
            findings.append(Finding(path, line, "module-level variable; core holds no ambient "
                                                "mutable run state"))
        elif inside and STATIC.match(code.strip()):
            findings.append(Finding(path, line, "Static variable; core holds no run state"))
        if header:
            inside = True
        elif PROC_CLOSE.match(code):
            inside = False
    return findings


def check(root: Path) -> tuple[list[Finding], int]:
    tracked = tracked_components(root)
    upper = names([path for path in tracked if path.startswith(UPPER_HOMES)])
    hidden = names([path for path in tracked if path.startswith(HIDDEN_HOMES)])
    components = [path for path in tracked if path.startswith((*INDEPENDENT, EXAMPLES))]
    upward = upward_any_case(facade_prefix(root))
    findings: list[Finding] = []
    for path in components:
        findings.extend(analyze(path, read_component(root, path), upper, hidden, upward))
    return findings, len(components)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components = check(args.root.resolve())
    return report("vba-dependencies", findings,
                  f"{components} component(s) in {', '.join((*INDEPENDENT, EXAMPLES))}")


if __name__ == "__main__":
    sys.exit(main())
