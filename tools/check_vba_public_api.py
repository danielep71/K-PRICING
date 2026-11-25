#!/usr/bin/env python3
"""Check VBA visibility and docs/PUBLIC_API.txt against the exported source.

Every procedure and module-level declaration states its visibility. The
inventory has two tiers. [supported] lists the public declarations of
src/modules/ (the facade); [project-public] lists every other public
declaration under src/, including the public members of classes, forms and
workbook modules. Each tier and the source must match one to one, with the
same normalized signatures, and no two public names may collide
(docs/REPOSITORY_STRUCTURE.md, Public API inventory). An inventory without a
tier heading is read as [supported] alone. A standard module declares no
public variable (docs/VBA_HOUSE_STYLE.md: no mutable public globals) and no
public Declare, so the supported kinds are the procedures, Const, Enum and
Type. Test and example modules belong to neither tier. Each compilation
environment of check_vba_conditionals.py is read separately.
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
    STRING,
    Finding,
    logical_units,
    read_component,
    report,
    split_statements,
    tracked_components,
)

ROOT = Path(__file__).resolve().parents[1]
INVENTORY = "docs/PUBLIC_API.txt"
SIG_PREFIX = "# SIG\t"
# Only src/modules/ is supported API; the rest of src/ is project-public, and
# test and example code is in neither tier (docs/REPOSITORY_STRUCTURE.md,
# Public API inventory).
PUBLIC_HOME = "src/modules/"
PROJECT_HOME = "src/"
SUPPORTED, PROJECT_PUBLIC = "supported", "project-public"
TIERS = (SUPPORTED, PROJECT_PUBLIC)
TIER_HOMES = {SUPPORTED: PUBLIC_HOME, PROJECT_PUBLIC: "src/ outside src/modules/"}
# A tier heading is a whole line in brackets; a line starting with # stays a comment.
HEADING = re.compile(r"^\[(?P<tier>[^\]]*)\]$")
HOMES = ("src/core/", "src/modules/", "src/classes/", "src/workbook/", "src/forms/",
         "tests/modules/", "examples/modules/")
PROPERTY_KINDS = frozenset({"property get", "property let", "property set"})

PROC = re.compile(
    r"^\s*(?:(?P<vis>Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+(?P<name>[A-Za-z_]\w*)\b", re.I)
END_PROC = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.I)
BLOCK = re.compile(r"^\s*Public\s+(?P<kind>Enum|Type)\s+(?P<name>[A-Za-z_]\w*)\b", re.I)
END_BLOCK = {"enum": re.compile(r"^\s*End\s+Enum\b", re.I),
             "type": re.compile(r"^\s*End\s+Type\b", re.I)}
DECLARE = re.compile(r"^\s*(?:Public|Global)\s+Declare\s+(?:PtrSafe\s+)?"
                     r"(?P<kind>Function|Sub)\s+(?P<name>[A-Za-z_]\w*)\b", re.I)
EVENT = re.compile(r"^\s*(?:Public|Global)\s+Event\s+(?P<name>[A-Za-z_]\w*)\b", re.I)
CONST = re.compile(r"^\s*(?:Public|Global)\s+Const\s+(?P<name>[A-Za-z_]\w*)\b(?P<rest>.*)$",
                   re.I)
VARIABLE = re.compile(r"^\s*(?:Public|Global)\s+(?P<withevents>WithEvents\s+)?"
                      r"(?P<name>[A-Za-z_]\w*)\b(?P<rest>.*)$", re.I)
IMPLICIT = re.compile(r"^\s*(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set)|Enum|Type|"
                      r"Event|Declare\s+(?:PtrSafe\s+)?(?:Function|Sub))\b", re.I)
# Module-level Dim and Const are private by default, but state no visibility.
UNSTATED = re.compile(r"^\s*(?:Dim|Const)\b", re.I)


@dataclass(frozen=True)
class Declaration:
    component: str
    kind: str
    name: str
    signature: str
    path: str
    line: int

    @property
    def key(self) -> str:
        return f"{self.component}\t{self.kind}\t{self.name}"



def norm(text: str) -> str:
    """White space as single spaces, none just inside parentheses, outside
    string literals: a signature split over continuation lines reads as it
    would on one line, and a literal such as a default keeps its value."""
    def code(part: str) -> str:
        return re.sub(r"\(\s+", "(", re.sub(r"\s+\)", ")", re.sub(r"\s+", " ", part)))
    parts: list[str] = []
    position = 0
    for literal in STRING.finditer(text):
        parts += [code(text[position:literal.start()]), literal.group()]
        position = literal.end()
    return ("".join(parts) + code(text[position:])).strip()


def statements(text: str) -> list[tuple[int, str]]:
    return [(line, part) for line, _, _, code in logical_units(text.split("\n"))
            for part in split_statements(code)]


def top_level_comma(text: str) -> bool:
    depth = 0
    in_string = False
    index = 0
    while index < len(text):
        char = text[index]
        if char == '"':
            if in_string and text[index:index + 2] == '""':
                index += 2
                continue
            in_string = not in_string
        elif not in_string and char == "(":
            depth += 1
        elif not in_string and char == ")" and depth:
            depth -= 1
        elif not in_string and char == "," and not depth:
            return True
        index += 1
    return False


def module_declaration(code: str) -> tuple[str, str] | str | None:
    """``(kind, name)`` of a module-level public declaration, a message for a
    prohibited form, or None for anything else."""
    match = DECLARE.match(code)
    if match:
        return f"Declare {match.group('kind').capitalize()}", match.group("name")
    match = EVENT.match(code)
    if match:
        return "Event", match.group("name")
    match = CONST.match(code)
    if match:
        if top_level_comma(match.group("rest")):
            return "a Public Const statement must declare one identifier"
        return "Const", match.group("name")
    match = VARIABLE.match(code)
    if match:
        if top_level_comma(match.group("rest")):
            return "a Public variable statement must declare one identifier"
        return ("WithEvents Variable" if match.group("withevents") else "Variable"), \
            match.group("name")
    return None


@dataclass
class Scan:
    """Declarations and findings of one component in one environment."""

    path: str
    declarations: list[Declaration] = field(default_factory=list)
    findings: list[Finding] = field(default_factory=list)

    def declare(self, kind: str, name: str, signature: str, line: int) -> None:
        self.declarations.append(Declaration(Path(self.path).stem, kind, name, signature,
                                             self.path, line))

    def reject(self, line: int, message: str) -> None:
        self.findings.append(Finding(self.path, line, message))


def scan_block(scan: Scan, match: re.Match[str], items: list[tuple[int, str]],
               index: int) -> int:
    """Record a Public Enum or Type with its body; return the index after it."""
    kind, name = match.group("kind").capitalize(), match.group("name")
    line = items[index][0]
    body = [norm(items[index][1])]
    end = index + 1
    while end < len(items) and not END_BLOCK[kind.casefold()].match(items[end][1]):
        if items[end][1].strip():
            body.append(norm(items[end][1]))
        end += 1
    if end == len(items):
        scan.reject(line, f"Public {kind} {name} is not closed")
        return index + 1
    body.append(norm(items[end][1]))
    scan.declare(kind, name, " | ".join(body), line)
    return end + 1


def scan_statement(scan: Scan, line: int, code: str) -> None:
    """A module-level statement other than a procedure or a Public Enum or Type."""
    if IMPLICIT.match(code):
        scan.reject(line, f"implicit public module-level declaration: {norm(code)}")
        return
    if UNSTATED.match(code):
        scan.reject(line, f"module-level declaration without Private or Public: {norm(code)}")
        return
    outcome = module_declaration(code)
    standard = Path(scan.path).suffix.lower() == ".bas"
    if isinstance(outcome, str):
        scan.reject(line, outcome)
    elif outcome is not None and standard and outcome[0].endswith("Variable"):
        scan.reject(line, f"public variable {outcome[1]} in a standard module; no mutable "
                          "public globals")
    elif outcome is not None and standard and outcome[0].startswith("Declare"):
        scan.reject(line, f"public Declare {outcome[1]}; declare native procedures Private")
    elif outcome is not None:
        scan.declare(outcome[0], outcome[1], norm(code), line)


def scan_source(path: str, text: str) -> Scan:
    scan = Scan(path)
    items = statements(text)
    inside = False
    index = 0
    while index < len(items):
        line, code = items[index]
        procedure = PROC.match(code)
        index += 1
        if not code.strip() or code.lstrip().startswith(("Attribute ", "Option ", "#")):
            continue
        if procedure:
            visibility = (procedure.group("vis") or "").casefold()
            if not visibility:
                scan.reject(line, f"implicit public procedure: {procedure.group('name')}")
            elif visibility == "public":
                kind = " ".join(part.capitalize() for part in procedure.group("kind").split())
                scan.declare(kind, procedure.group("name"), norm(code), line)
            inside = True
        elif END_PROC.match(code):
            inside = False
        elif not inside and (block := BLOCK.match(code)):
            index = scan_block(scan, block, items, index - 1)
        elif not inside:
            scan_statement(scan, line, code)
    return scan


def scan_component(path: str, text: str) -> tuple[dict[Declaration, list[str]], list[Finding]]:
    """Declarations with the environments that reach them, and the findings."""
    sources, findings = reachable_sources(path, text)
    declarations: dict[Declaration, list[str]] = {}
    for environment, source in sources.items():
        scan = scan_source(path, source)
        findings.extend(item for item in scan.findings if item not in findings)
        for declaration in scan.declarations:
            declarations.setdefault(declaration, []).append(environment)
    return declarations, findings


def property_pair(first: Declaration, second: Declaration) -> bool:
    """A Property Get with its Let or Set in the same component shares one name."""
    return (first.component.casefold() == second.component.casefold()
            and first.kind.casefold() in PROPERTY_KINDS
            and second.kind.casefold() in PROPERTY_KINDS
            and first.kind.casefold() != second.kind.casefold())


def collisions(declarations: dict[Declaration, list[str]]) -> list[Finding]:
    """Same key twice, or one public name in two standard modules, in a shared environment."""
    findings: list[Finding] = []
    seen: list[Declaration] = []
    for declaration, environments in declarations.items():
        for previous in seen:
            if not set(declarations[previous]) & set(environments):
                continue
            if previous.key.casefold() == declaration.key.casefold():
                findings.append(Finding(declaration.path, declaration.line,
                                        f"public declaration appears more than once: "
                                        f"{declaration.key}"))
            elif (previous.name.casefold() == declaration.name.casefold()
                  and Path(declaration.path).suffix.lower() == ".bas"
                  and Path(previous.path).suffix.lower() == ".bas"
                  and not property_pair(previous, declaration)):
                findings.append(Finding(declaration.path, declaration.line,
                                        f"public name {declaration.name!r} collides with "
                                        f"{previous.component}.{previous.name} "
                                        f"({previous.kind})"))
        seen.append(declaration)
    return findings


def tier_of(path: str) -> str | None:
    """The inventory tier of a component's public declarations, if any."""
    if path.startswith(PUBLIC_HOME):
        return SUPPORTED
    if path.startswith(PROJECT_HOME):
        return PROJECT_PUBLIC
    return None


@dataclass
class Inventory:
    """Rows and signature records by case-folded key, each with its tier."""

    rows: dict[str, tuple[str, str]] = field(default_factory=dict)
    signatures: dict[str, dict[str, set[str]]] = field(default_factory=dict)
    tiered: bool = False


def read_inventory(root: Path) -> tuple[Inventory, list[Finding]]:
    """The inventory's rows and records; without a tier heading, all are [supported]."""
    path = root / INVENTORY
    inventory = Inventory()
    if not path.is_file():
        return inventory, [Finding(INVENTORY, None, "public API inventory is missing")]
    lines = path.read_text(encoding="utf-8").splitlines()
    headings = [HEADING.match(raw.strip()) for raw in lines]
    inventory.tiered = any(headings)
    tier: str | None = None if inventory.tiered else SUPPORTED
    seen: set[str] = set()
    findings: list[Finding] = []
    for number, (raw, heading) in enumerate(zip(lines, headings), 1):
        if heading:
            name = heading.group("tier")
            if name not in TIERS:
                findings.append(Finding(INVENTORY, number, f"unknown tier heading [{name}]; "
                                        "use [supported] or [project-public]"))
            elif name in seen:
                findings.append(Finding(INVENTORY, number,
                                        f"tier heading appears more than once: [{name}]"))
            seen.add(name)
            tier = name if name in TIERS else None
            continue
        record = raw.startswith(SIG_PREFIX)
        if not record and (not raw.strip() or raw.lstrip().startswith("#")):
            continue
        if tier is None:
            findings.append(Finding(INVENTORY, number, "inventory entry outside a "
                                    "[supported] or [project-public] tier"))
            continue
        if record:
            fields = raw[len(SIG_PREFIX):].split("\t", 3)
            if len(fields) != 4:
                findings.append(Finding(INVENTORY, number, "malformed # SIG record"))
                continue
            variants = inventory.signatures.setdefault(
                "\t".join(fields[:3]).casefold(), {}).setdefault(tier, set())
            if fields[3] in variants:
                findings.append(Finding(INVENTORY, number, "duplicate signature record: "
                                                           + "\t".join(fields[:3])))
            variants.add(fields[3])
            continue
        fields = raw.split("\t")
        if len(fields) != 3 or not all(fields):
            findings.append(Finding(INVENTORY, number,
                                    "an inventory row needs component, kind and name"))
        elif raw.casefold() in inventory.rows:
            findings.append(Finding(INVENTORY, number, f"duplicate inventory row: {raw}"))
        else:
            inventory.rows[raw.casefold()] = (tier, raw)
    return inventory, findings


def inventory_findings(declarations: dict[Declaration, list[str]],
                       inventory: Inventory) -> list[Finding]:
    """Each tier and the source must list the same declarations with the same signatures."""
    actual: dict[str, tuple[str, str, set[str]]] = {}
    for declaration in declarations:
        tier = tier_of(declaration.path)
        if tier is not None:
            entry = actual.setdefault(declaration.key.casefold(),
                                      (declaration.key, tier, set()))
            entry[2].add(declaration.signature)
    findings: list[Finding] = []
    for folded, (key, tier, found) in actual.items():
        row = inventory.rows.get(folded)
        if row is None:
            findings.append(Finding(INVENTORY, None,
                                    f"public declaration not listed: {key}; list it under "
                                    f"[{tier}]"))
        elif row[0] != tier:
            findings.append(Finding(INVENTORY, None, f"public declaration listed under "
                                                     f"[{row[0]}], not [{tier}]: {key}"))
        records = inventory.signatures.get(folded, {})
        findings.extend(Finding(INVENTORY, None, f"signature record listed under [{other}], "
                                                 f"not [{tier}]: {key}")
                        for other in records if other != tier)
        if tier in records and records[tier] != found:
            findings.append(Finding(INVENTORY, None, f"signature mismatch for {key}: source "
                                                     f"{sorted(found)!r}, inventory "
                                                     f"{sorted(records[tier])!r}"))
        elif not records:
            findings.append(Finding(INVENTORY, None, f"signature record missing: {key}"))
    findings.extend(Finding(INVENTORY, None, f"listed declaration not in {TIER_HOMES[tier]}: "
                                             f"{row}")
                    for folded, (tier, row) in inventory.rows.items() if folded not in actual)
    findings.extend(Finding(INVENTORY, None, f"stale signature record: {folded}")
                    for folded in inventory.signatures if folded not in actual)
    if not inventory.tiered and any(tier == PROJECT_PUBLIC for _, tier, _ in actual.values()):
        findings.append(Finding(INVENTORY, None, "the inventory has no tier heading, so it is "
                                                 "read as [supported] alone; list the other "
                                                 "public declarations of src/ under "
                                                 "[project-public]"))
    return findings


def check(root: Path) -> tuple[list[Finding], int, int, int]:
    components = [path for path in tracked_components(root) if path.startswith(HOMES)]
    declarations: dict[Declaration, list[str]] = {}
    findings: list[Finding] = []
    for path in components:
        found, errors = scan_component(path, read_component(root, path))
        declarations.update(found)
        findings.extend(errors)
    # Collisions span every home, since Excel compiles them into one project;
    # only the facade's declarations are supported API.
    findings.extend(collisions(declarations))
    inventory, errors = read_inventory(root)
    findings.extend(errors)
    findings.extend(inventory_findings(declarations, inventory))
    counts = {tier: len({declaration.key.casefold() for declaration in declarations
                         if tier_of(declaration.path) == tier}) for tier in TIERS}
    return findings, len(components), counts[SUPPORTED], counts[PROJECT_PUBLIC]


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components, supported, project = check(args.root.resolve())
    return report("vba-public-api", findings, f"{components} VBA component(s), "
                  f"{supported} supported declaration(s) in {PUBLIC_HOME}, "
                  f"{project} project-public declaration(s) in the rest of src/")


if __name__ == "__main__":
    sys.exit(main())
