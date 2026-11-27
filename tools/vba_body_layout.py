"""The body and statement rules of docs/VBA_HOUSE_STYLE.md (*Layout*), on one module's text.

tools/check_vba_layout.py runs them with the header rules of vba_header_layout.py.
They read the source only: nothing is compiled or run. A rule a procedure meets
in a way these readings cannot see, such as a label inside a #If branch, is
left to review.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

from vba_header_layout import ATTRIBUTE, BANNER, COMPONENT, PROCEDURE, VB_NAME, vbe_header_length
from vba_lexer import Finding, logical_units, strip_code

MAX_LINE = 120
COMPACT_MAX = 15    # code lines up to which a Private procedure may keep a compact body
TITLE_WIDTH = 78
END = re.compile(r"^End\s+(Sub|Function|Property)\b", re.I)
SECTION = re.compile(r"^' ([A-Z0-9][A-Z0-9 ,'()/&#!-]*[A-Z0-9)])$")
RANKS = {"DECLARE": 0, "INITIALIZE": 1, "EXIT PROCEDURE": 3, "ERROR HANDLER": 4}
TASK_RANK = 2
PROC_NAME = re.compile(r'^Const\s+PROC_NAME\s+As\s+String\s*=\s*"([^"]*)"$', re.I)
ON_ERROR = re.compile(r"^On\s+Error\s+GoTo\s+([A-Za-z]\w*)$", re.I)
EXIT = re.compile(r"^Exit\s+(?:Sub|Function|Property)$", re.I)
LABEL = re.compile(r"^([A-Za-z]\w*):$")
LOCAL = re.compile(r"^(?:Dim|Static)\s+(.+)$", re.I)
PASCAL = re.compile(r"^[A-Z][A-Za-z0-9]*$")
HUNGARIAN = re.compile(r"^(?:str|lng|int|obj|bln|dbl|var|rng|ws|wb)[A-Z]")
# A one-line declaration whose "As" lines up with its neighbours.
DECLARATION = re.compile(r"^(\s*(?:(?:Public|Private|Global|Friend|Dim|Static)\s+(?:Const\s+)?"
                         r"|Const\s+)(?:WithEvents\s+)?\w+(?:\([^)]*\))?\s+)(As)\s", re.I)
COMPARED = re.compile(r"(?<![\w.])(?:VBA\.)?(InStr|StrComp)\s*\(", re.I)
SELECT = re.compile(r"^Select\s+Case\b", re.I)
CASE_ELSE = re.compile(r"^Case\s+Else\b", re.I)
END_SELECT = re.compile(r"^End\s+Select\b", re.I)
PARAMETER = re.compile(r"^(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?(?:ParamArray\s+)?(\w+)", re.I)


@dataclass(frozen=True)
class Procedure:
    name: str
    kind: str
    first: int          # 0-based index of the signature's first line
    last: int           # 0-based index of its last line
    end: int            # 0-based index of the End line
    code: str           # the signature as one logical line


def code_of(line: str) -> str:
    """A physical line's code without its comment, string contents blanked."""
    return strip_code(line, blank_strings=True).strip()


def blank_run(lines: Sequence[str], index: int) -> int:
    """The number of blank lines right above ``index``."""
    count = 0
    while index - count - 1 >= 0 and not lines[index - count - 1].strip():
        count += 1
    return count


def procedures(lines: Sequence[str], start: int) -> list[Procedure]:
    found: list[Procedure] = []
    units = logical_units(list(lines), blank_strings=True)
    for position, (first, last, kind, code) in enumerate(units):
        if first - 1 < start or kind != "code":
            continue
        match = PROCEDURE.match(code.strip())
        if match is None:
            continue
        end = next((later[0] - 1 for later in units[position + 1:]
                    if later[2] == "code" and END.match(later[3].strip())), len(lines) - 1)
        word = code.strip().split("(")[0]
        found.append(Procedure(match.group(1), "Sub" if re.search(r"\bSub\b", word, re.I)
                               else "Function" if re.search(r"\bFunction\b", word, re.I)
                               else "Property", first - 1, last - 1, end, code.strip()))
    return found


def top_level(text: str) -> list[str]:
    """The parts of ``text`` between commas outside parentheses."""
    depth, current, parts = 0, "", []
    for char in text:
        depth += (char == "(") - (char == ")")
        if char == "," and depth == 0:
            parts.append(current)
            current = ""
        else:
            current += char
    return [*parts, current]


def parameters(code: str) -> list[str]:
    """The parameter names of a signature."""
    opened = code.find("(")
    if opened < 0:
        return []
    depth, current, parts = 0, "", []
    for char in code[opened:]:
        if char == "(":
            depth += 1
            if depth == 1:
                continue
        elif char == ")":
            depth -= 1
            if depth == 0:
                break
        if char == "," and depth == 1:
            parts.append(current)
            current = ""
        else:
            current += char
    parts.append(current)
    names = []
    for part in parts:
        match = PARAMETER.match(part.strip())
        if match and part.strip():
            names.append(match.group(1))
    return names


def header_span(lines: Sequence[str], procedure: Procedure) -> tuple[int, int]:
    """The first line of a procedure's header and the line after it (0-based)."""
    index = procedure.last + 1
    while index < procedure.end and ATTRIBUTE.match(lines[index]):
        index += 1
    while index < procedure.end and not lines[index].strip():
        index += 1
    stop = index
    while stop < procedure.end and lines[stop].startswith("'"):
        stop += 1
    return index, stop


def sections(lines: Sequence[str], first: int, stop: int) -> list[tuple[str, int]]:
    """(title, 0-based index of the title line) of each section banner in lines[first:stop]."""
    found = []
    for index in range(first + 1, stop - 1):
        match = SECTION.match(lines[index].rstrip())
        if match and BANNER.match(lines[index - 1]) and BANNER.match(lines[index + 1]):
            found.append((match.group(1), index))
    return found


def field_text(block: Sequence[str], field: str) -> list[str]:
    """The text lines of one header field."""
    try:
        start = next(index for index, line in enumerate(block) if line.rstrip() == f"' {field}")
    except StopIteration:
        return []
    text = []
    for line in block[start + 1:]:
        if re.match(r"^' [A-Z]", line) or BANNER.match(line):
            break
        text.append(line.rstrip())
    return text


def header_findings(path: str, lines: Sequence[str], procedure: Procedure) -> list[Finding]:
    """The title, the spacing around the signature, and INPUTS and RETURNS against it."""
    findings: list[Finding] = []
    what = f"procedure {procedure.name}"
    first, stop = header_span(lines, procedure)
    block = list(lines[first:stop])
    if not block:
        return findings
    blanks = first - procedure.last - 1 - sum(
        1 for index in range(procedure.last + 1, first) if ATTRIBUTE.match(lines[index]))
    expected = 1 if procedure.first == procedure.last else 0
    if blanks != expected:
        findings.append(Finding(path, procedure.first + 1, f"{what}: {blanks} blank line(s) "
                                f"between the signature and its header, expected {expected}"))
    if len(block) > 3 and BANNER.match(block[1]) and BANNER.match(block[3]):
        title = block[2][1:].strip()
        lead = len(block[2]) - 1 - len(block[2][1:].lstrip(" "))
        want = max(0, (TITLE_WIDTH - len(title)) // 2)
        if title and lead != want:
            findings.append(Finding(path, first + 3, f"{what}: title {title!r} is not centred: "
                                    f"{lead} leading spaces, expected {want}"))
    returns = [line[4:].strip() for line in field_text(block, "RETURNS") if line.strip("'")]
    if procedure.kind == "Sub" and returns and not re.match(r"Nothing\b", returns[0]):
        findings.append(Finding(path, first + 1, f"{what}: RETURNS of a Sub is not Nothing"))
    inputs = [line for line in field_text(block, "INPUTS") if line.strip("'").strip()]
    names = [line[4:] for line in inputs if re.match(r"^'   \S", line)]
    expected_names = parameters(procedure.code)
    if names == ["None"]:
        if expected_names:
            findings.append(Finding(path, first + 1, f"{what}: INPUTS is None, but the "
                                    f"signature has {', '.join(expected_names)}"))
    elif names and all(re.fullmatch(r"\w+", name) for name in names) \
            and [name.casefold() for name in names] != \
            [name.casefold() for name in expected_names]:
        findings.append(Finding(path, first + 1, f"{what}: INPUTS lists {', '.join(names)}; "
                                f"the signature has {', '.join(expected_names) or 'none'}"))
    return findings


def alignment_findings(path: str, lines: Sequence[str], first: int,
                       stop: int) -> list[Finding]:
    """Declarations whose As or trailing comment is not in the column of the first
    declaration of their run; a blank, comment or other line ends a run."""
    findings: list[Finding] = []
    runs: list[list[int]] = []
    run: list[int] = []
    for index in range(first, stop):
        code = code_of(lines[index])
        if DECLARATION.match(lines[index]) and not code.endswith(" _"):
            run.append(index)
        elif run:
            runs.append(run)
            run = []
    if run:
        runs.append(run)
    for run in runs:
        columns = {}
        comments = {}
        for index in run:
            match = DECLARATION.match(lines[index])
            assert match is not None
            columns[index] = match.start(2)
            code = strip_code(lines[index])
            if len(code) < len(lines[index].rstrip()):
                comments[index] = lines[index].index("'", len(code.rstrip()))
        for label, found in (("As", columns), ("trailing comment", comments)):
            if not found:
                continue
            reference = found[min(found)]
            wrong = next((index for index in sorted(found) if found[index] != reference), None)
            if wrong is not None:
                findings.append(Finding(path, wrong + 1, f"{label} in column "
                                        f"{found[wrong] + 1}, not in column {reference + 1} "
                                        "of the first declaration of its run"))
    return findings


@dataclass(frozen=True)
class Body:
    """One procedure's body: its sections and its logical statements, continuation lines
    joined, each at the 0-based index of its first line."""
    path: str
    what: str
    procedure: Procedure
    found: list[tuple[str, int]]
    statements: list[tuple[int, str]]

    def section_of(self, index: int) -> str:
        """The title of the section that holds line ``index``."""
        title = ""
        for name, at in self.found:
            if at > index:
                break
            title = name
        return title

    def finding(self, index: int, message: str) -> Finding:
        return Finding(self.path, index + 1, f"{self.what}: {message}")


def spacing_findings(body: Body, lines: Sequence[str]) -> list[Finding]:
    """One blank line before each section banner and before End, none after a banner."""
    findings = []
    for title, index in body.found:
        if blank_run(lines, index - 1) != 1:
            findings.append(body.finding(index - 1, f"{blank_run(lines, index - 1)} blank "
                                         f"line(s) before the section {title}, expected 1"))
        if index + 2 < body.procedure.end and not lines[index + 2].strip():
            findings.append(body.finding(index + 2, "a blank line after the banner of the "
                                         f"section {title}"))
    end = body.procedure.end
    if blank_run(lines, end) != 1:
        findings.append(body.finding(end, f"{blank_run(lines, end)} blank line(s) before "
                                     f"{lines[end].strip()}, expected 1"))
    rank = -1
    for title, index in body.found:
        current = RANKS.get(title, TASK_RANK)
        if current < rank:
            findings.append(body.finding(index, f"section {title} is out of order; the order "
                                         "is DECLARE, INITIALIZE, the task sections, EXIT "
                                         "PROCEDURE, ERROR HANDLER"))
            break
        rank = current
    return findings


def declare_findings(body: Body, lines: Sequence[str], component: str) -> list[Finding]:
    """PROC_NAME first in DECLARE and the alignment there, or a missing DECLARE."""
    declare = next((at for title, at in body.found if title == "DECLARE"), None)
    if declare is None:
        need = next((index for index, code in body.statements
                     if LOCAL.match(code) or ON_ERROR.match(code)), None)
        return [] if need is None else [body.finding(
            need, "no DECLARE section for its local variables or its error handler")]
    findings = []
    index = next((index for index, _ in body.statements if index > declare), declare)
    named = PROC_NAME.match(strip_code(lines[index]).strip()) \
        if body.section_of(index) == "DECLARE" else None
    expected = f"{component}.{body.procedure.name}"
    if named is None:
        findings.append(body.finding(index, "DECLARE does not start with "
                                     f'Const PROC_NAME As String = "{expected}"'))
    elif named.group(1).casefold() != expected.casefold():
        findings.append(body.finding(index, f'PROC_NAME is "{named.group(1)}", '
                                     f'expected "{expected}"'))
    stop = next((at for _, at in body.found if at > declare), body.procedure.end)
    findings += [body.finding(item.line - 1, item.message)
                 for item in alignment_findings(body.path, lines, declare + 2, stop)
                 if item.line is not None]
    return findings


def handler_findings(body: Body) -> list[Finding]:
    """The first On Error GoTo in INITIALIZE, and a clean Exit in EXIT PROCEDURE before
    its label."""
    handler = next(((index, match.group(1)) for index, code in body.statements
                    if (match := ON_ERROR.match(code))), None)
    if handler is None:
        return []
    findings = []
    index, label = handler
    if body.section_of(index) != "INITIALIZE":
        findings.append(body.finding(index, f"On Error GoTo {label} is not in INITIALIZE"))
    at = next((index for index, code in body.statements
               if (match := LABEL.match(code)) and match.group(1).casefold()
               == label.casefold()), None)
    if at is not None and not any(EXIT.match(code) and body.section_of(index) == "EXIT PROCEDURE"
                                  for index, code in body.statements if index < at):
        findings.append(body.finding(at, f"no Exit {body.procedure.kind} in EXIT PROCEDURE "
                                     f"before {label}:"))
    return findings


def name_findings(body: Body) -> list[Finding]:
    """Local variables that are not descriptive PascalCase names."""
    findings = []
    for index, code in body.statements:
        local = LOCAL.match(code)
        for part in top_level(local.group(1)) if local else ():
            name = part.strip().split("(")[0].split()[0] if part.strip() else ""
            if name and (not PASCAL.match(name) or HUNGARIAN.match(name)):
                findings.append(body.finding(index, f"local variable {name} is not a "
                                             "descriptive PascalCase name"))
    return findings


def body_findings(path: str, component: str, lines: Sequence[str],
                  procedure: Procedure) -> list[Finding]:
    """The sections of one procedure: their spacing and order, PROC_NAME, the error handler
    in INITIALIZE with the clean exit before it, the alignment in DECLARE and local names."""
    _, first = header_span(lines, procedure)
    body = Body(path, f"procedure {procedure.name}", procedure,
                sections(lines, first, procedure.end),
                [(first + start - 1, code.strip()) for start, _, kind, code
                 in logical_units(list(lines[first:procedure.end]), blank_strings=True)
                 if kind == "code" and code.strip()])
    return (spacing_findings(body, lines) + declare_findings(body, lines, component)
            + handler_findings(body) + name_findings(body))


def group_banner_findings(path: str, lines: Sequence[str], start: int,
                          found: Sequence[Procedure]) -> list[Finding]:
    """Two blank lines between the declarations and the first group banner, one after an
    End, and one between a group banner and the signature below it."""
    findings: list[Finding] = []
    inside = {index for procedure in found for index in range(procedure.first,
                                                               procedure.end + 1)}
    index = start
    seen_code = False
    while index < len(lines):
        if index in inside:
            index += 1
            continue
        if lines[index].rstrip() != "'" or index + 1 >= len(lines) \
                or not BANNER.match(lines[index + 1]):
            seen_code = seen_code or bool(code_of(lines[index]))
            index += 1
            continue
        last = index
        while last + 1 < len(lines) and lines[last + 1].startswith("'") and last + 1 not in inside:
            last += 1
        if sum(1 for at in range(index, last + 1) if BANNER.match(lines[at])) < 2 \
                or not any(lines[at].strip("' ") and not BANNER.match(lines[at])
                           for at in range(index, last + 1)):
            index = last + 1
            continue
        above = index - blank_run(lines, index) - 1
        blanks = blank_run(lines, index)
        if above >= 0 and END.match(lines[above].strip()):
            if blanks != 1:
                findings.append(Finding(path, index + 1, f"group banner: {blanks} blank line(s) "
                                        "after the End above it, expected 1"))
        elif seen_code and blanks != 2:
            findings.append(Finding(path, index + 1, f"group banner: {blanks} blank line(s) "
                                    "after the declarations, expected 2"))
        below = last + 1
        while below < len(lines) and not lines[below].strip():
            below += 1
        if below < len(lines) and PROCEDURE.match(lines[below]) and below - last - 1 != 1:
            findings.append(Finding(path, below + 1, f"{below - last - 1} blank line(s) between "
                                    "a group banner and the signature below it, expected 1"))
        index = last + 1
    return findings


def length(line: str) -> int:
    """A line's length without the contents of its string literals, and without the trailing
    comment of a declaration, which keeps its column."""
    code = strip_code(line, blank_strings=True)
    if DECLARATION.match(line) and len(code) < len(line.rstrip()):
        line = code
    return len(re.sub(r'"[^"]*"', '""', line.rstrip()))


def statement_findings(path: str, lines: Sequence[str], start: int) -> list[Finding]:
    """Lines over MAX_LINE, InStr and StrComp without a compare mode, and Select Case
    blocks without Case Else."""
    findings: list[Finding] = []
    for index in range(start, len(lines)):
        if length(lines[index]) > MAX_LINE:
            findings.append(Finding(path, index + 1, f"line is {length(lines[index])} characters "
                                    f"without its string contents, over {MAX_LINE}"))
    selects: list[tuple[int, bool]] = []
    for first, _, kind, code in logical_units(list(lines), blank_strings=True):
        if first - 1 < start or kind != "code":
            continue
        code = code.strip()
        for match in COMPARED.finditer(code):
            depth, arguments, position = 0, 1, match.end() - 1
            for char in code[position:]:
                depth += (char == "(") - (char == ")")
                if depth == 0:
                    break
                if char == "," and depth == 1:
                    arguments += 1
            inside = code[position:]
            named = re.search(r"\bCompare\s*:=", inside, re.I)
            short = (arguments < 4 if match.group(1).lower() == "instr" else arguments < 3)
            if short and not named:
                findings.append(Finding(path, first, f"{match.group(1)} without an explicit "
                                        "compare mode"))
        if SELECT.match(code):
            selects.append((first, False))
        elif CASE_ELSE.match(code) and selects:
            selects[-1] = (selects[-1][0], True)
        elif END_SELECT.match(code) and selects:
            at, has_else = selects.pop()
            if not has_else:
                findings.append(Finding(path, at, "Select Case without Case Else"))
    return findings


def module_header_end(lines: Sequence[str], start: int, stop: int) -> int:
    """The line after the module (or class) header, or ``start`` without one."""
    for index in range(start, stop):
        if COMPONENT.match(lines[index]):
            end = index
            while end < stop and lines[end].startswith("'"):
                end += 1
            return end
    return start


def component_findings(path: str, lines: Sequence[str], start: int,
                       stop: int) -> list[Finding]:
    """A component header whose INPUTS is not "None at module level" (or class level)."""
    for index in range(start, stop):
        match = COMPONENT.match(lines[index])
        if match is None:
            continue
        end = module_header_end(lines, start, stop)
        inputs = [line for line in field_text(lines[index:end], "INPUTS") if line.strip("'")]
        want = f"None at {match.group(1).lower()} level"
        if inputs and inputs[0][4:].strip().rstrip(".") != want:
            at = index + next(offset for offset, line in enumerate(lines[index:end])
                              if line.rstrip() == "' INPUTS")
            return [Finding(path, at + 2, f"component INPUTS is not {want!r}")]
        return []
    return []


def check_text(path: str, text: str) -> list[Finding]:
    """The findings of the body and statement rules in one component's text."""
    lines = text.replace("\r\n", "\n").split("\n")
    start = vbe_header_length(lines)
    named = VB_NAME.search(text)
    component = named.group(1) if named else Path(path).stem
    found = procedures(lines, start)
    findings: list[Finding] = []
    for procedure in found:
        first, stop = header_span(lines, procedure)
        if not any(BANNER.match(line) for line in lines[first:stop]):
            continue
        findings += header_findings(path, lines, procedure)
        code = sum(1 for line in lines[stop:procedure.end]
                   if line.strip() and not line.lstrip().startswith("'"))
        if not (procedure.code.lower().startswith("private ") and code <= COMPACT_MAX):
            findings += body_findings(path, component, lines, procedure)
    first = found[0].first if found else len(lines)
    findings += component_findings(path, lines, start, first)
    findings += alignment_findings(path, lines, start, first)
    findings += group_banner_findings(path, lines, module_header_end(lines, start, first), found)
    findings += statement_findings(path, lines, start)
    return sorted(findings, key=lambda item: (item.line or 0, item.message))
