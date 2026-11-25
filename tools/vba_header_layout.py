"""The banner and header rules of docs/VBA_HOUSE_STYLE.md (*Layout*), on one module's text.

tools/check_vba_layout.py runs them over the tracked modules, and
tools/format_vba.py asks them whether a baselined module is now fully in
layout. They read comments only: nothing is compiled or run.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from pathlib import Path

from vba_lexer import Finding, logical_units

FIELDS = ("PURPOSE", "WHY THIS EXISTS", "INPUTS", "RETURNS", "BEHAVIOR", "ERROR POLICY",
          "DEPENDENCIES", "NOTES", "UPDATED")
BANNER_WIDTH = 79
# A comment line made of one repeated dash or equals sign, ten or more times.
BANNER = re.compile(r"^'([-=])\1{9,}\s*$")
LABEL = re.compile(r"^' ([A-Z][A-Z ]*[A-Z])\s*$")
COMPONENT = re.compile(r"^' (MODULE|CLASS): (\S+)\s*$")
FIELD_TEXT = re.compile(r"^'   \S")
DATE = re.compile(r"^'   \d{4}-\d{2}-\d{2}\s*$")
VBE_LINE = re.compile(r"^(?:VERSION\s|Object\s*=|Attribute\s)", re.I)
VB_NAME = re.compile(r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"', re.I | re.M)
PROCEDURE = re.compile(r"^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
                       r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)", re.I)
ATTRIBUTE = re.compile(r"^Attribute\s+\w+\.VB_", re.I)


def vbe_header_length(lines: Sequence[str]) -> int:
    """The number of leading lines the VBE writes: VERSION, BEGIN ... END, Attribute."""
    depth = 0
    for index, line in enumerate(lines):
        word = line.strip().split(" ", 1)[0].upper()
        if word == "BEGIN":
            depth += 1
        elif word == "END" and depth:
            depth -= 1
        elif not depth and not VBE_LINE.match(line):
            return index
    return len(lines)


def comment_block(lines: Sequence[str], start: int) -> list[tuple[int, str]]:
    """The consecutive comment lines from ``start``, with their 0-based indexes."""
    block = []
    index = start
    while index < len(lines) and lines[index].startswith("'"):
        block.append((index, lines[index].rstrip()))
        index += 1
    return block


def header_findings(path: str, block: Sequence[tuple[int, str]], what: str,
                    first_line: int) -> list[Finding]:
    """The findings on the fields of one header."""
    findings: list[Finding] = []
    labels: list[tuple[int, str]] = []
    for position, (index, line) in enumerate(block):
        match = LABEL.match(line)
        if not match:
            continue
        label = match.group(1)
        if label not in FIELDS:
            findings.append(Finding(path, index + 1, f"{what}: unknown header field {label}"))
            continue
        labels.append((position, label))
    found = [label for _, label in labels]
    if found != list(FIELDS):
        missing = [field for field in FIELDS if field not in found]
        detail = f"missing {', '.join(missing)}" if missing else \
            f"fields out of order or repeated: {', '.join(found)}"
        findings.append(Finding(path, first_line, f"{what}: header {detail}"))
        return findings
    ends = [position for position, _ in labels[1:]] + [len(block)]
    for (position, label), end in zip(labels, ends, strict=True):
        text = [line for _, line in block[position + 1:end] if not BANNER.match(line)]
        index = block[position][0] + 1
        if not any(FIELD_TEXT.match(line) for line in text):
            findings.append(Finding(path, index, f"{what}: {label} has no text; "
                                                 "write None"))
        elif label == "UPDATED" and not (len([t for t in text if FIELD_TEXT.match(t)]) == 1
                                         and DATE.match(next(t for t in text
                                                             if FIELD_TEXT.match(t)))):
            findings.append(Finding(path, index, f"{what}: UPDATED is not one "
                                                 "YYYY-MM-DD date"))
    return findings


def check_text(path: str, text: str) -> tuple[list[Finding], int]:
    """The findings in one component's text and the number of procedures read."""
    lines = text.replace("\r\n", "\n").split("\n")
    findings: list[Finding] = []
    start = vbe_header_length(lines)
    for index in range(start, len(lines)):
        line = lines[index].rstrip()
        if BANNER.match(line) and len(line) != BANNER_WIDTH:
            findings.append(Finding(path, index + 1, f"banner is {len(line)} characters, "
                                                     f"not {BANNER_WIDTH}"))
    # The component header: after blank lines and Option statements only.
    index = start
    while index < len(lines) and (not lines[index].strip()
                                  or lines[index].lstrip().lower().startswith("option ")):
        index += 1
    block = comment_block(lines, index)
    named = [COMPONENT.match(line) for _, line in block]
    kinds = ("MODULE", "CLASS") if path.lower().endswith(".cls") else ("MODULE",)
    name_match = VB_NAME.search(text)
    name = name_match.group(1) if name_match else Path(path).stem
    found = next((match for match in named if match), None)
    if not block or found is None:
        findings.append(Finding(path, index + 1, "component header missing: after the "
                                                 "Option statements, expected a comment "
                                                 f"block naming ' {kinds[0]}: {name}"))
    else:
        if found.group(1) not in kinds or found.group(2) != name:
            findings.append(Finding(path, index + 1, f"component header names "
                                                     f"{found.group(1)}: {found.group(2)}, "
                                                     f"expected {kinds[0]}: {name}"))
        findings += header_findings(path, block, "component", index + 1)
    # Each procedure's header. A declaration is found among the logical lines,
    # so a line continued with " _" is one declaration and a comment continued
    # with " _" never holds one.
    procedures = 0
    resume = start
    for first, last, kind, code in logical_units(lines, blank_strings=True):
        if first - 1 < resume or kind != "code":
            continue
        match = PROCEDURE.match(code.strip())
        if not match:
            continue
        procedures += 1
        # The VBE writes a procedure's Attribute lines right below its signature.
        following = last
        while following < len(lines) and ATTRIBUTE.match(lines[following]):
            following += 1
        if following < len(lines) and not lines[following].strip():
            following += 1
        block = comment_block(lines, following)
        what = f"procedure {match.group(1)}"
        if not block:
            findings.append(Finding(path, first, f"{what}: header missing below its "
                                                 "declaration"))
        else:
            findings += header_findings(path, block, what, first)
        resume = following + len(block)
    return findings, procedures
