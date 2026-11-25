#!/usr/bin/env python3
"""Check that every VBA jump target is a label of the procedure that jumps.

``GoTo``, ``GoSub``, ``Resume <label>`` and ``On Error GoTo <label>``, and
the line number in ``If x Then 100``, which VBA reads as ``GoTo 100``, must
name a label defined in the same procedure, labels must be unique within a
procedure, and every procedure must be closed. Each compilation environment
of check_vba_conditionals.py is checked separately, so that mutually
exclusive ``#If`` branches are never combined.

A raise reached after a ``Resume <label>`` while one of the procedure's own
``On Error GoTo <label>`` handlers is enabled is reported: that handler
traps the raise again, so a handler that resumes to the same code loops,
and the error never reaches the caller. Raise a saved error after
``On Error GoTo 0``. A raise is ``Err.Raise`` or ``VBA.Err.Raise``, with or
without ``Call`` and named arguments, or the ``Error`` statement; a
``.Raise`` inside ``With Err`` is not seen. The model of the handler state:

- ``On Error GoTo <label>`` enables that local handler, ``On Error GoTo 0``
  disables it and ``On Error Resume Next`` replaces it. ``On Error GoTo -1``
  only ends the error-handling state, so the enabled handler stays enabled.
- A ``Resume <label>`` belongs to the handler whose label precedes it most
  closely, and re-enables that handler, which trapped the error. A
  ``Resume`` before any handler label is not followed.
- Only a ``Resume <label>`` that the handler runs unconditionally starts a
  path. One inside an ``If``, ``Select Case`` or loop of the handler, or in
  a one-line ``If``, runs only under a condition, as in a bounded retry, so
  it starts no path, and a raise it would loop through is not reported.
- From the ``Resume`` target, the check follows the text of the procedure.
  ``If``, ``Select Case`` and loop blocks merge the states of their
  branches. A one-line ``If ... Then a: b Else c: d`` is read as a block
  ``If`` whose branches hold the statements in order, up to the end of the
  line, and so is ``If x Then: a``. A combined ``Next j, i`` closes one
  loop per name, and ``Exit Do`` or ``Exit For`` leaves the innermost loop
  of its kind. Where the branches disagree, the state is unknown and
  nothing is reported.
- After a loop, the path continues from its ``Exit Do`` or ``Exit For``
  statements; from the end of its body if the loop tests a condition
  (``For``, ``While``, ``Do While`` or ``Until``, ``Loop While`` or
  ``Until``); and from before the loop if it tests first and may run no
  pass. A bare ``Do ... Loop`` with no ``Exit Do`` never continues. If the
  body may change the handler and an ``Exit`` leaves the loop, a later pass
  may leave in another state, so the state is unknown. Raises in the body
  are checked in the first pass only.
- A ``GoSub`` or ``On ... GoSub`` keeps the state only if its subroutine
  runs straight to an unconditional ``Return``: no label, block, one-line
  ``If``, jump, ``Exit``, ``End``, raise or ``On Error`` statement before
  it. Any other subroutine, or a missing one, makes the state unknown. A
  raise inside a subroutine is not checked.
- A target inside blocks stays inside them: at the next ``ElseIf``,
  ``Else`` or ``Case`` of an enclosing block, the path skips to that block's
  end. At the end of an enclosing loop, the path continues after the loop
  as above; if the loop's body changes the handler, the state is unknown
  there, and a raise before the target in a later pass is not checked. An
  enclosing bare ``Do ... Loop`` continues only from an ``Exit Do`` the path
  reached, so one reached only in a later pass, such as an ``Exit Do``
  before the target, is not followed.
- The path ends at an ``Exit Sub``, ``Exit Function``, ``Exit Property``,
  ``End``, ``GoTo``, ``Resume`` or ``Return`` outside any block, and at the
  end of the procedure; inside a block, that branch ends there and takes no
  part in the merge. If the path ends inside an enclosing loop, an
  ``Exit Do`` or ``Exit For`` it passed before still continues after that
  loop; nothing else does. A raise ends the path in the same way when the handler
  is disabled or a local handler is enabled. Under ``On Error Resume Next``
  the raise falls through, and in an unknown state it falls through too, so
  that the branch keeps its unknown state in the merge and cannot produce a
  finding. ``Stop`` only pauses, so the path goes on.
  Jumps are not followed: the path goes on after an
  ``On ... GoTo`` list, which falls through when no target is chosen, and a
  label on the path does not merge the states of other jumps to it. In a
  branch the path has left, a label that some statement jumps to makes the
  state unknown; a label nothing jumps to, such as a line number kept for
  ``Erl``, leaves the branch ended.
- ``Resume`` and ``Resume Next`` continue at a statement that cannot be known
  statically, so no path starts from them. A raise inside a handler before
  its ``Resume`` is not trapped by that handler, so it is not reported.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path

from check_vba_conditionals import reachable_sources
from vba_lexer import (
    Finding,
    logical_units,
    read_component,
    report,
    split_statements,
    tracked_components,
)

ROOT = Path(__file__).resolve().parents[1]
PROC_OPEN = re.compile(
    r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\b", re.IGNORECASE)
PROC_CLOSE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
NAMED_LABEL = re.compile(r"^\s*([A-Za-z_]\w*|\d+):")
# A line number, followed by a statement or by nothing once a comment such
# as "100 Rem target" or "100 ' target" has been stripped.
NUMBERED_LABEL = re.compile(r"^\s*(\d+)(?=\s|$)")
# A keyword after a dot is a member call, such as Application.Goto or .Goto
# in a With block, not a jump.
JUMP = re.compile(r"(?<![\w.])(GoTo|GoSub|Resume)\b(?:\s+([A-Za-z_]\w*|\d+|-1))?", re.IGNORECASE)
ON_ERROR = re.compile(r"\bon\s+error\s*$", re.IGNORECASE)
# On <expression> GoTo|GoSub <label>, <label>, ...: every listed target.
LABEL = r"(?:[A-Za-z_]\w*|\d+)"
COMPUTED_JUMP = re.compile(rf"\bOn\b(?!\s+Error\b).*?(?<![\w.])(GoTo|GoSub)\s+"
                           rf"({LABEL}(?:\s*,\s*{LABEL})*)", re.IGNORECASE)
# Statements that can stand alone before a colon, as in "Else: x = 1"; a
# keyword is never a label.
KEYWORDS = frozenset({"beep", "do", "else", "end", "loop", "next", "resume", "return", "stop",
                      "wend"})
# The handler-state model of the re-raise rule; see the module docstring.
ON_ERROR_SET = re.compile(r"(?<![\w.])On\s+Error\s+(?:GoTo\s+(-1|[A-Za-z_]\w*|\d+)|Resume\s+Next)\b",
                          re.IGNORECASE)
ERR_RAISE = re.compile(r"(?<![\w.])(?:VBA\s*\.\s*)?Err\s*\.\s*Raise\b", re.IGNORECASE)
# The Error statement, as in "Error 5"; Error$ and Error(...) in an
# expression are functions, and never start a statement.
ERROR_STATEMENT = re.compile(r"^Error\b(?!\s*\$)\s*\S", re.IGNORECASE)
# A line number alone after Then or Else, as in "If x Then 100", is a GoTo.
LINE_NUMBER = re.compile(r"^\d+$")
IMPLICIT = "<implicit>"
COMBINED_NEXT =re.compile(r"^Next\s+(\w+(?:\s*,\s*\w+)+)$", re.IGNORECASE)
# A jump or exit that keeps a subroutine from running straight to its Return.
BRANCH_JUMP = re.compile(r"(?<![\w.])(?:GoTo|GoSub|Resume|Return)\b|^(?:Exit|End)\b",
                         re.IGNORECASE)
ONE_LINE_IF = re.compile(r"^(If\b.*?\bThen)\s+(\S.*)$", re.IGNORECASE)
ONE_LINE_ELSE = re.compile(r"(?<![\w.])Else\b", re.IGNORECASE)
CASE_ELSE = re.compile(r"^Case\s+Else$", re.IGNORECASE)
BLOCK_IF = re.compile(r"^If\b.*\bThen$", re.IGNORECASE)
SELECT = re.compile(r"^Select\s+Case\b", re.IGNORECASE)
BRANCH = re.compile(r"^(?:ElseIf\b|Else$|Case\b)", re.IGNORECASE)
BRANCH_ELSE = re.compile(r"^(?:Else|Case\s+Else)$", re.IGNORECASE)
BLOCK_END = re.compile(r"^End\s+(?:If|Select)$", re.IGNORECASE)
LOOP_OPEN = re.compile(r"^(Do|For|While)\b", re.IGNORECASE)
LOOP_CLOSE = re.compile(r"^(?:Loop|Next|Wend)\b", re.IGNORECASE)
LOOP_EXIT = re.compile(r"^Exit\s+(Do|For)$", re.IGNORECASE)
# A loop that tests before its body may run it no time; Do ... Loop While or
# Until runs it at least once; a bare Do ... Loop is left only by Exit Do.
PRETESTED_LOOP = re.compile(r"^(?:For|While|Do\s+(?:While|Until))\b", re.IGNORECASE)
POSTTESTED_CLOSE = re.compile(r"^Loop\s+(?:While|Until)\b", re.IGNORECASE)
PATH_END = re.compile(r"^(?:Exit\s+(?:Sub|Function|Property)|End|GoTo\b.*|Resume\b.*|Return)$",
                      re.IGNORECASE)
# Handler states other than an enabled local handler, which is its label.
OFF, RESUME_NEXT, UNKNOWN, DEAD = "<off>", "<resume next>", "<unknown>", "<dead>"


@dataclass
class Step:
    """A label or a statement of a procedure, in source order."""
    line: int
    label: str | None = None
    code: str = ""


@dataclass
class Block:
    """An If, Select Case or loop block open on a re-raise path. An outer
    block was already open at the ``Resume`` target; ``start`` is the index
    of its opening step. ``kind`` is ``"do"``, ``"for"`` or ``"while"``
    for a loop, which ``Exit Do`` or ``Exit For`` leaves."""
    loop: bool
    entry: str
    kind: str = ""
    ends: list[str] = field(default_factory=list)
    has_else: bool = False
    started: bool = False
    outer: bool = False
    start: int = 0


@dataclass
class Procedure:
    name: str
    line: int
    labels: dict[str, int] = field(default_factory=dict)
    jumps: list[tuple[int, str, str]] = field(default_factory=list)
    steps: list[Step] = field(default_factory=list)


def is_procedure(code: str) -> re.Match[str] | None:
    """A procedure header; a Declare statement never matches, since its
    keyword stands where PROC_OPEN expects Sub, Function or Property."""
    return PROC_OPEN.match(code)


def label_at_start(code: str) -> str | None:
    match = NAMED_LABEL.match(code) or NUMBERED_LABEL.match(code)
    return match.group(1) if match and match.group(1).casefold() not in KEYWORDS else None


def statement_jumps(code: str) -> list[tuple[str, str]]:
    """``(operation, target)`` for each jump to a label in one statement.

    ``Resume``, ``Resume 0``, ``Resume Next`` and ``On Error GoTo 0`` or ``-1`` name no
    label;
    ``On <expression> GoTo`` and ``GoSub`` name every target of their list.
    """
    jumps: list[tuple[str, str]] = []
    computed = COMPUTED_JUMP.search(code)
    if computed:
        jumps.extend((computed.group(1), target.strip())
                     for target in computed.group(2).split(","))
        code = code[:computed.start()] + " " * (computed.end() - computed.start()) + \
            code[computed.end():]
    for match in JUMP.finditer(code):
        operation, target = match.group(1), match.group(2)
        if target is None or (operation.casefold() == "resume"
                              and target.casefold() in {"next", "0"}):
            continue
        if operation.casefold() == "goto" and target in {"0", "-1"} \
                and ON_ERROR.search(code[:match.start()]):
            continue
        jumps.append((operation, target))
    return jumps


def scan(path: str, text: str) -> tuple[list[Procedure], list[Finding]]:
    """Collect each procedure's labels and jumps; report structural defects."""
    procedures: list[Procedure] = []
    findings: list[Finding] = []
    current: Procedure | None = None
    for line, _, _, code in logical_units(text.split("\n"), blank_strings=True):
        # A label starts the logical line; the statements after it, split at
        # ":" separators, may open or close a procedure and may jump.
        label = label_at_start(code) if current is not None else None
        if current is not None and label is not None:
            if label.casefold() in current.labels:
                findings.append(Finding(path, line, f"duplicate label {label!r} in procedure "
                                                    f"{current.name}; first defined at line "
                                                    f"{current.labels[label.casefold()]}"))
            else:
                current.labels[label.casefold()] = line
            current.steps.append(Step(line, label=label))
        # The statements of each procedure on this line, recorded once the
        # line, or the procedure, ends, so that a one-line If spans them all.
        owner, texts = current, []
        for index, statement in enumerate(split_statements(code)):
            # The label opens the first statement, as in "Cleanup" or "100 Exit Sub".
            text = statement[len(label):].strip() if index == 0 and label else statement
            before = current
            current = scan_statement(path, line, statement, current, procedures, findings)
            if before is not None and current is before:
                texts.append(text)
            else:
                record_steps(owner, line, texts)
                owner, texts = current, []
        record_steps(owner, line, texts)
    if current is not None:
        findings.append(Finding(path, current.line, f"procedure {current.name} is not closed"))
    return procedures, findings


def scan_statement(path: str, line: int, statement: str, current: Procedure | None,
                   procedures: list[Procedure], findings: list[Finding]) -> Procedure | None:
    """Apply one statement; return the procedure open after it."""
    opener = is_procedure(statement)
    if opener:
        if current is not None:
            findings.append(Finding(path, line, f"procedure {current.name} opened at line "
                                                f"{current.line} is not closed before "
                                                f"{opener.group(2)}"))
        current = Procedure(opener.group(2), line)
        procedures.append(current)
    elif PROC_CLOSE.match(statement):
        if current is None:
            findings.append(Finding(path, line, "End statement outside any procedure"))
        current = None
    elif current is not None:
        current.jumps.extend((line, operation, target)
                             for operation, target in statement_jumps(statement))
    return current


def record_steps(procedure: Procedure | None, line: int, statements: list[str]) -> None:
    """Append the statements of one line to ``procedure``, in block form; a
    combined ``Next j, i`` closes one loop per name. An implicit line-number
    jump is also a jump of the procedure, whose target must be a label."""
    if procedure is None:
        return
    codes = block_form([statement for statement in statements if statement])
    procedure.jumps.extend((line, "GoTo", code.split()[1]) for code in codes
                           if code.startswith(IMPLICIT))
    for code in codes:
        code = code.removeprefix(IMPLICIT)
        combined = COMBINED_NEXT.match(code)
        names = combined.group(1).split(",") if combined else []
        procedure.steps.extend(Step(line, code=f"Next {name.strip()}") for name in names)
        if not combined:
            procedure.steps.append(Step(line, code=code))


def one_line_parts(statement: str) -> list[tuple[str, str]]:
    """``(kind, code)`` parts of a statement: ``"if"`` for the head of a
    one-line If, ``"else"`` for its Else, ``"code"`` for anything else."""
    match = ONE_LINE_IF.match(statement)
    if match:
        return [("if", match.group(1)), *one_line_parts(match.group(2))]
    pieces = ONE_LINE_ELSE.split(statement)
    if len(pieces) == 1 or CASE_ELSE.match(statement):
        return [("code", statement)]
    parts: list[tuple[str, str]] = []
    for index, piece in enumerate(pieces):
        if index:
            parts.append(("else", "Else"))
        if piece.strip():
            parts.extend(one_line_parts(piece.strip()))
    return parts


def block_form(statements: list[str]) -> list[str]:
    """The statements of one line, with each one-line ``If ... Then a: b
    Else c`` written as a block ``If`` that ends with the line. An Else
    belongs to the nearest open one-line If, as in VBA. In ``If x Then: y``,
    the statements after the colon belong to that one-line If. A line number
    alone after ``Then`` or ``Else`` is a jump there: it is returned as
    ``GoTo <n>`` behind the IMPLICIT prefix, which record_steps removes."""
    parts = [part for index, statement in enumerate(statements)
             for part in ([("if", statement)]
                          if BLOCK_IF.match(statement) and index < len(statements) - 1
                          else one_line_parts(statement))]

    def branch(index: int, nested: bool) -> tuple[list[str], int]:
        codes: list[str] = []
        while index < len(parts):
            kind, code = parts[index]
            if kind == "else" and nested:
                return codes, index
            if kind == "if":
                then, index = branch(index + 1, True)
                codes += [code, *then]
                if index < len(parts):
                    other, index = branch(index + 1, True)
                    codes += ["Else", *other]
                codes.append("End If")
                continue
            # A block If's Else, as in "Else: x = 1", stays a statement.
            codes.append(f"{IMPLICIT}GoTo {code}" if nested and LINE_NUMBER.match(code)
                         else code)
            index += 1
        return codes, index

    return branch(0, False)[0]


def enabled(state: str) -> bool:
    """Whether ``state`` is a local handler, named by its label."""
    return not state.startswith("<")


def handler_after(code: str, state: str) -> str:
    """The handler state after the ``On Error`` statements in ``code``."""
    for match in ON_ERROR_SET.finditer(code):
        target = match.group(1)
        if target is None:
            state = RESUME_NEXT
        elif target == "0":
            state = OFF
        elif target != "-1":
            state = target
    return state


def merge(states: list[str]) -> str:
    """The state where branches meet: DEAD if none reaches it, UNKNOWN if they disagree."""
    live = [state for state in states if state != DEAD]
    if not live:
        return DEAD
    return live[0] if len({state.casefold() for state in live}) == 1 else UNKNOWN


def changes_handler(steps: list[Step]) -> bool:
    """Whether ``steps`` run an ``On Error`` statement, directly or in a subroutine."""
    return any(ON_ERROR_SET.search(step.code) or subroutine_changes_handler(steps, step.code)
               for step in steps)


def loop_has_exit(steps: list[Step], start: int, close: int) -> bool:
    """Whether the loop opened at ``steps[start]`` and closed at
    ``steps[close]`` holds an ``Exit Do`` or ``Exit For`` that leaves it."""
    opener = LOOP_OPEN.match(steps[start].code)
    kind = opener.group(1).casefold() if opener else ""
    inner: list[str] = []
    for step in steps[start + 1:close]:
        nested = LOOP_OPEN.match(step.code)
        leaving = LOOP_EXIT.match(step.code)
        if nested:
            inner.append(nested.group(1).casefold())
        elif LOOP_CLOSE.match(step.code) and inner:
            inner.pop()
        elif leaving and leaving.group(1).casefold() == kind and kind not in inner:
            return True
    return False


def after_loop(steps: list[Step], block: Block, close: int, state: str) -> str:
    """The state after the loop ``block`` closes at ``steps[close]``, with
    ``state`` at the end of the body. Leaving through an ``Exit`` gives
    ``block.ends``; a loop that tests its condition also ends after the
    body, and one that tests it first may run no pass. A later pass that
    may change the handler before an ``Exit`` makes the state unknown. An
    enclosing bare ``Do ... Loop`` continues only from the exits the path
    reached, never from one that is present but not reached."""
    may_skip = bool(PRETESTED_LOOP.match(steps[block.start].code))
    tested = may_skip or bool(POSTTESTED_CLOSE.match(steps[close].code))
    if not tested and not loop_has_exit(steps, block.start, close):
        return DEAD
    body = steps[block.start + 1:close]
    if block.outer:
        # Another pass runs the body before the target too.
        if state != DEAD and changes_handler(body):
            return UNKNOWN
        return merge([*block.ends, *([state] if tested else [])])
    if merge(block.ends) != DEAD and changes_handler(body):
        return UNKNOWN
    return merge([*block.ends, *([state] if tested else []),
                  *([block.entry] if may_skip else [])])


def block_statement(steps: list[Step], index: int, state: str, blocks: list[Block]
                    ) -> str | None:
    """The state after ``steps[index]`` if it opens, divides, closes or
    leaves a block, or None for any other statement."""
    code = steps[index].code
    if BLOCK_IF.match(code) or SELECT.match(code):
        blocks.append(Block(loop=False, entry=state, started=bool(BLOCK_IF.match(code))))
        return state
    opener = LOOP_OPEN.match(code)
    if opener:
        blocks.append(Block(loop=True, entry=state, kind=opener.group(1).casefold(),
                            start=index))
        return state
    leaving = LOOP_EXIT.match(code)
    left = [block for block in blocks if leaving and block.kind == leaving.group(1).casefold()]
    if left:
        left[-1].ends.append(state)
        return DEAD
    if not blocks:
        return None
    block = blocks[-1]
    if BRANCH.match(code) and not block.loop:
        if block.started:
            block.ends.append(state)
        block.started = True
        block.has_else = block.has_else or bool(BRANCH_ELSE.match(code))
        return block.entry
    if BLOCK_END.match(code) and not block.loop:
        blocks.pop()
        return merge([*block.ends, state, *([] if block.has_else else [block.entry])])
    if LOOP_CLOSE.match(code) and block.loop:
        blocks.pop()
        return after_loop(steps, block, index, state)
    return None


def enclosing_blocks(steps: list[Step], target: int) -> list[Block]:
    """The blocks open at ``steps[target]``, outermost first."""
    blocks: list[Block] = []
    for index, step in enumerate(steps[:target]):
        opener = LOOP_OPEN.match(step.code)
        if BLOCK_IF.match(step.code) or SELECT.match(step.code) or opener:
            blocks.append(Block(loop=bool(opener), entry=UNKNOWN, outer=True, start=index,
                                kind=opener.group(1).casefold() if opener else ""))
        elif blocks and (LOOP_CLOSE.match(step.code) if blocks[-1].loop
                         else BLOCK_END.match(step.code)):
            blocks.pop()
    return blocks


def outer_statement(steps: list[Step], index: int, state: str, blocks: list[Block]
                    ) -> tuple[str, bool] | None:
    """``(state, skip)`` after a statement that divides or closes the
    innermost block when that block encloses the ``Resume`` target, or None.
    ``skip`` means the path leaves the block's other branches unrun."""
    if not blocks or not blocks[-1].outer:
        return None
    block, code = blocks[-1], steps[index].code
    if not block.loop and BRANCH.match(code):
        return state, True
    if not block.loop and BLOCK_END.match(code):
        blocks.pop()
        return state, False
    if block.loop and LOOP_CLOSE.match(code):
        blocks.pop()
        return after_loop(steps, block, index, state), False
    return None


def straight_subroutine(steps: list[Step], label: int) -> bool:
    """Whether the subroutine at ``steps[label]`` runs straight to a
    ``Return``: no label, block, one-line If, jump, exit, raise or
    ``On Error`` statement before it."""
    for step in steps[label + 1:]:
        code = step.code
        if code.casefold() == "return":
            return True
        if step.label is not None or BLOCK_IF.match(code) or SELECT.match(code) \
                or BRANCH.match(code) or BLOCK_END.match(code) or LOOP_OPEN.match(code) \
                or LOOP_CLOSE.match(code) or BRANCH_JUMP.search(code) \
                or ON_ERROR_SET.search(code) or ERR_RAISE.search(code) \
                or ERROR_STATEMENT.match(code):
            return False
    return False


def subroutine_changes_handler(steps: list[Step], code: str) -> bool:
    """Whether a ``GoSub`` in ``code`` may leave the handler in another
    state: only a subroutine that runs straight to its ``Return`` keeps it,
    and one whose label is missing does not."""
    labels = {step.label.casefold(): index for index, step in enumerate(steps)
              if step.label is not None}
    return any(target.casefold() not in labels
               or not straight_subroutine(steps, labels[target.casefold()])
               for operation, target in statement_jumps(code)
               if operation.casefold() == "gosub")


def jump_targets(steps: list[Step]) -> set[str]:
    """The labels that some statement may jump to: ``GoTo``, ``GoSub``,
    ``On ... GoTo`` and ``GoSub``, ``Resume <label>``, ``On Error GoTo
    <label>``, and a line number alone after ``Then`` or ``Else``, which
    the scan records as ``GoTo``."""
    return {target.casefold() for step in steps for _, target in statement_jumps(step.code)}


def skipped(code: str, depth: int, blocks: list[Block]) -> int | None:
    """The depth of nested blocks after ``code`` while the path skips the
    other branches of the innermost enclosing block, or None once its end
    closes that block."""
    if BLOCK_IF.match(code) or SELECT.match(code) or LOOP_OPEN.match(code):
        return depth + 1
    if BLOCK_END.match(code) or LOOP_CLOSE.match(code):
        if depth == 0:
            blocks.pop()
            return None
        return depth - 1
    return depth


def trapped_raises(steps: list[Step], start: int, handler: str
                   ) -> list[tuple[int, str, str]]:
    """``(line, handler, raise)`` for each ``Err.Raise`` or ``Error``
    statement that a local handler traps on the path from ``steps[start]``,
    entered with ``handler`` enabled."""
    found: list[tuple[int, str, str]] = []
    state = handler
    blocks = enclosing_blocks(steps, start)
    targeted = jump_targets(steps)
    # While the path skips the other branches of an enclosing block: the
    # depth of the blocks nested in them.
    skipping: int | None = None
    # Whether the path ended inside an enclosing block; only the Exit Do or
    # Exit For it collected for an enclosing loop can carry it on.
    ended = False
    for index in range(start, len(steps)):
        step = steps[index]
        if skipping is not None:
            skipping = skipped(step.code, skipping, blocks)
            continue
        outer = outer_statement(steps, index, state, blocks)
        if outer is not None:
            state, skip = outer
            skipping = 0 if skip else None
            ended = ended and state == DEAD
            if ended and not any(block.outer for block in blocks):
                break
            continue
        if step.label is not None:
            # Another jump may reach a label in a branch this path left; a
            # label nothing jumps to, such as a line number for Erl, cannot.
            if state == DEAD and not ended and step.label.casefold() in targeted:
                state = UNKNOWN
            continue
        after = block_statement(steps, index, state, blocks)
        if after is not None:
            state = after
            continue
        if state == DEAD:
            continue
        raised = "Err.Raise" if ERR_RAISE.search(step.code) else \
            "Error" if ERROR_STATEMENT.match(step.code) else None
        if raised and enabled(state):
            found.append((step.line, state, raised))
        if subroutine_changes_handler(steps, step.code):
            state = UNKNOWN
        elif (raised and (state == OFF or enabled(state))) or PATH_END.match(step.code):
            # A raise goes to the caller with no handler or, as reported
            # above, to the local handler. Under On Error Resume Next it
            # falls through; in an unknown state too, since ending that
            # branch would let the other branches alone decide the merge.
            if not blocks:
                break
            ended = blocks[-1].outer
            state = DEAD
        else:
            state = handler_after(step.code, state)
    return found


def reraise_findings(path: str, procedure: Procedure) -> list[Finding]:
    """Each ``Err.Raise`` or ``Error`` statement reached after a
    ``Resume <label>`` while a local handler of ``procedure`` is enabled."""
    handlers = {match.group(1).casefold() for step in procedure.steps
                for match in ON_ERROR_SET.finditer(step.code)
                if match.group(1) not in {None, "0", "-1"}}
    targets = {step.label.casefold(): index for index, step in enumerate(procedure.steps)
               if step.label is not None}
    messages: dict[tuple[int, str], str] = {}
    trapping: str | None = None
    # The blocks open in the handler; a Resume inside one is conditional.
    depth = 0
    for step in procedure.steps:
        if step.label is not None and step.label.casefold() in handlers:
            trapping, depth = step.label, 0
        if BLOCK_IF.match(step.code) or SELECT.match(step.code) or LOOP_OPEN.match(step.code):
            depth += 1
        elif BLOCK_END.match(step.code) or LOOP_CLOSE.match(step.code):
            depth = max(depth - 1, 0)
        for operation, target in statement_jumps(step.code):
            if trapping is None or depth or operation.casefold() != "resume" \
                    or target.casefold() not in targets:
                continue
            for line, handler, raised in trapped_raises(procedure.steps,
                                                        targets[target.casefold()], trapping):
                messages.setdefault((line, handler.casefold()), (
                    f"{raised} after Resume {target} is trapped again by handler "
                    f"{handler!r} of procedure {procedure.name}; put On Error GoTo 0 "
                    f"before the raise"))
    return [Finding(path, line, message) for (line, _), message in sorted(messages.items())]


def analyze_source(path: str, text: str) -> list[Finding]:
    procedures, findings = scan(path, text)
    for procedure in procedures:
        findings.extend(Finding(path, line, f"{operation} target {target!r} is not a label of "
                                            f"procedure {procedure.name}")
                        for line, operation, target in procedure.jumps
                        if target.casefold() not in procedure.labels)
        findings.extend(reraise_findings(path, procedure))
    return findings


def analyze(path: str, text: str) -> list[Finding]:
    """Findings of every environment, each listed once with the environments it occurs in."""
    sources, findings = reachable_sources(path, text)
    environments: dict[Finding, list[str]] = {}
    for environment, source in sources.items():
        for finding in analyze_source(path, source):
            environments.setdefault(finding, []).append(environment)
    findings.extend(Finding(item.path, item.line, f"{item.message} [{', '.join(names)}]")
                    for item, names in environments.items())
    return findings


def check(root: Path) -> tuple[list[Finding], int, int]:
    components = tracked_components(root)
    findings: list[Finding] = []
    procedures = 0
    for path in components:
        text = read_component(root, path)
        findings.extend(analyze(path, text))
        procedures += sum(bool(is_procedure(statement)) for _, _, _, code
                          in logical_units(text.split("\n"), blank_strings=True)
                          for statement in split_statements(code))
    return findings, len(components), procedures


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components, procedures = check(args.root.resolve())
    return report("vba-jumps", findings,
                  f"{components} VBA component(s), {procedures} procedure(s)")


if __name__ == "__main__":
    sys.exit(main())
