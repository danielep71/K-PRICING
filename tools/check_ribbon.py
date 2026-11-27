#!/usr/bin/env python3
"""Check the workbook's Ribbon, src/workbook/customUI14.xml, against its VBA.

The build writes the XML into every workbook as its Ribbon part
(tools/devbuild/build_workbook.py). Excel loads it only when it opens the
workbook, and then shows no error for a callback it cannot find: the button
just does nothing. This check reads the XML and the VBA without Excel:

- the XML is well formed, has no DOCTYPE or entity declaration, and its root
  is the customUI element of Office 2010's namespace, customUI/2009/07;
- every control id matches KPR_ followed by letters, digits and underscores,
  and is used once; a control has no custom image (image, loadImage), only a
  built-in imageMso icon;
- every callback the XML names (onLoad, onAction, onChange and every get...
  attribute) is a Public Sub of src/workbook/M_Ribbon.bas with the parameters
  Office passes, and every Public Sub of that module with a Ribbon parameter
  is named by the XML; the onAction that M_Ribbon writes into the menus it
  builds, its Private Const ON_ACTION, is one of those Subs too;
- every control with an onAction has a tag, and the tag is the value of a
  Public Const ACTION_ of src/workbook/M_Actions.bas, the texts that
  M_Actions.RunAction runs. ACTION_FUNCTION, a prefix that the menu built at
  run time completes, is not a whole action.

Without src/workbook/customUI14.xml the check passes and says so. It never
changes a file, and it compiles and runs nothing.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree

from vba_lexer import Finding, logical_units, report

ROOT = Path(__file__).resolve().parents[1]
CUSTOM_UI = "src/workbook/customUI14.xml"
CALLBACKS = "src/workbook/M_Ribbon.bas"
ACTIONS = "src/workbook/M_Actions.bas"
GATE = "ribbon"

NAMESPACE = "http://schemas.microsoft.com/office/2009/07/customui"
MAX_BYTES = 1024 * 1024
UNSAFE = re.compile(r"<!\s*(?:DOCTYPE|ENTITY)\b", re.I)
CONTROL_ID = re.compile(r"KPR_[A-Za-z0-9_]+")
VBA_NAME = re.compile(r"[A-Za-z][A-Za-z0-9_]*")
CUSTOM_IMAGES = ("image", "loadImage")
GETTER = re.compile(r"get[A-Z][A-Za-z]*")
# Item getters receive the item's index between the control and the returned value.
ITEM_GETTERS = frozenset({"getItemCount", "getItemLabel", "getItemID", "getItemImage",
                          "getItemScreentip", "getItemSupertip"})
TOGGLES = frozenset({"toggleButton", "checkBox"})
SELECTIONS = frozenset({"dropDown", "gallery"})
# The function prefix completed at run time; no static control may carry it alone.
FUNCTION_PREFIX = "ACTION_FUNCTION"

PROCEDURE = re.compile(
    r"^(?:(?P<scope>Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"(?P<name>[A-Za-z][A-Za-z0-9_]*)\s*(?P<rest>.*)$", re.I)
PARAMETER = re.compile(
    r"^(?:(?P<optional>Optional)\s+)?(?:(?P<mode>ByVal|ByRef)\s+)?"
    r"(?:(?P<paramarray>ParamArray)\s+)?(?P<name>[A-Za-z][A-Za-z0-9_]*)"
    r"\s*(?P<array>\(\s*\))?\s*(?:As\s+(?P<type>[A-Za-z][A-Za-z0-9_.]*))?\s*(?:=.*)?$", re.I)
CONSTANT = re.compile(r'^(?P<scope>Public|Private|Global)?\s*Const\s+(?P<name>\w+)\s+'
                      r'As\s+String\s*=\s*"(?P<value>(?:[^"]|"")*)"$', re.I)


@dataclass(frozen=True)
class Slot:
    """One parameter of a callback: how it reads, the types it may have, and
    whether Office returns a value through it."""

    text: str
    types: frozenset[str | None]
    returned: bool = False


@dataclass(frozen=True)
class Parameter:
    name: str
    mode: str
    type: str | None
    special: bool

    @property
    def type_key(self) -> str | None:
        return self.type.rpartition(".")[2].casefold() if self.type else None


@dataclass(frozen=True)
class Procedure:
    name: str
    kind: str
    scope: str
    parameters: tuple[Parameter, ...] | None
    line: int

    @property
    def ribbon_shaped(self) -> bool:
        return any(item.type_key in ("iribboncontrol", "iribbonui")
                   for item in self.parameters or ())


@dataclass(frozen=True)
class Use:
    """A callback that the XML names: on which attribute of which control."""

    name: str
    attribute: str
    tag: str
    control: str | None

    def where(self) -> str:
        target = f"<{self.tag} id={self.control!r}>" if self.control else f"<{self.tag}>"
        return f"{self.attribute} of {target}"


CONTROL = Slot("control As IRibbonControl", frozenset({"iribboncontrol"}))
RETURNED = Slot("ByRef returnedVal As Variant", frozenset({"variant", None}), True)
INDEX = Slot("index As Long", frozenset({"integer", "long"}))
ON_LOAD = (Slot("ribbon As IRibbonUI", frozenset({"iribbonui"})),)
BUTTON = (CONTROL,)
TOGGLE = (CONTROL, Slot("pressed As Boolean", frozenset({"boolean"})))
SELECTION = (CONTROL, Slot("selectedId As String", frozenset({"string"})),
             Slot("selectedIndex As Long", frozenset({"integer", "long"})))
CHANGE = (CONTROL, Slot("text As String", frozenset({"string"})))


def local_name(tag: str) -> str:
    return tag.rpartition("}")[2]


def shape(use: Use) -> tuple[Slot, ...]:
    """The parameters Office 2010 passes to a callback of that attribute and control."""
    if use.attribute == "onLoad":
        return ON_LOAD
    if use.attribute == "onAction":
        return TOGGLE if use.tag in TOGGLES else SELECTION if use.tag in SELECTIONS else BUTTON
    if use.attribute == "onChange":
        return CHANGE
    if use.attribute in ITEM_GETTERS and use.attribute != "getItemCount":
        return (CONTROL, INDEX, RETURNED)
    return (CONTROL, RETURNED)


def is_callback(attribute: str) -> bool:
    return attribute in ("onLoad", "onAction", "onChange") or bool(GETTER.fullmatch(attribute))


def load_xml(root: Path) -> tuple[ElementTree.Element | None, list[Finding]]:
    data = (root / CUSTOM_UI).read_bytes()
    if len(data) > MAX_BYTES:
        return None, [Finding(CUSTOM_UI, None, f"is over {MAX_BYTES} bytes")]
    text = data.decode("utf-8", errors="replace")
    if UNSAFE.search(text):
        return None, [Finding(CUSTOM_UI, None, "declares a DOCTYPE or an entity")]
    try:
        element = ElementTree.fromstring(data)
    except ElementTree.ParseError as error:
        return None, [Finding(CUSTOM_UI, None, f"is not well-formed XML ({error})")]
    if element.tag != f"{{{NAMESPACE}}}customUI":
        return None, [Finding(CUSTOM_UI, None, f"its root is {element.tag}, not customUI of "
                              f"{NAMESPACE}")]
    return element, []


def xml_findings(element: ElementTree.Element) -> tuple[list[Finding], list[Use], list[str]]:
    """The findings on ids and images, the callbacks named, and the static action tags."""
    findings: list[Finding] = []
    uses: list[Use] = []
    tags: list[str] = []
    seen: dict[str, int] = {}
    for item in element.iter():
        tag = local_name(item.tag)
        control = item.get("id")
        if control is not None:
            seen[control] = seen.get(control, 0) + 1
            if not CONTROL_ID.fullmatch(control):
                findings.append(Finding(CUSTOM_UI, None, f"<{tag}> id {control!r} does not "
                                        f"match {CONTROL_ID.pattern}"))
        findings += [Finding(CUSTOM_UI, None, f"<{tag} id={control!r}> uses {name}; use a "
                             "built-in imageMso icon") for name in CUSTOM_IMAGES if name in item.attrib]
        uses += [Use(value, attribute, tag, control) for attribute, value in item.attrib.items()
                 if is_callback(attribute)]
        if "onAction" in item.attrib:
            if item.get("tag") is None:
                findings.append(Finding(CUSTOM_UI, None, f"<{tag} id={control!r}> has an "
                                        "onAction but no tag naming its action"))
            else:
                tags.append(item.get("tag", ""))
    findings += [Finding(CUSTOM_UI, None, f"control id {control!r} is used {count} times")
                 for control, count in seen.items() if count > 1]
    return findings, uses, tags


def code_lines(text: str) -> list[tuple[int, str]]:
    """The first line and the code of each logical line of a module, comments removed."""
    return [(first, code.strip()) for first, _, kind, code in
            logical_units(text.replace("\r\n", "\n").split("\n")) if kind == "code"]


def split_top_level(text: str) -> list[str]:
    pieces: list[str] = []
    depth = 0
    current: list[str] = []
    for char in text:
        depth += (char == "(") - (char == ")")
        if char == "," and depth == 0:
            pieces.append("".join(current))
            current = []
        else:
            current.append(char)
    pieces.append("".join(current))
    return [piece.strip() for piece in pieces]


def parameters(rest: str) -> tuple[Parameter, ...] | None:
    """The parameter list after a procedure's name; None when it cannot be read."""
    rest = rest.strip()
    if not rest.startswith("("):
        return ()
    depth = 0
    for index, char in enumerate(rest):
        depth += (char == "(") - (char == ")")
        if depth == 0:
            inner = rest[1:index]
            break
    else:
        return None
    if not inner.strip():
        return ()
    found = []
    for piece in split_top_level(inner):
        match = PARAMETER.match(piece)
        if match is None:
            return None
        found.append(Parameter(match.group("name"), (match.group("mode") or "").casefold(),
                               match.group("type"), bool(match.group("optional")
                                                         or match.group("paramarray")
                                                         or match.group("array"))))
    return tuple(found)


def procedures(text: str) -> list[Procedure]:
    found = []
    for line, code in code_lines(text):
        match = PROCEDURE.match(code)
        if match is None:
            continue
        scope = (match.group("scope") or "public").casefold()
        found.append(Procedure(match.group("name"), " ".join(match.group("kind").split()).casefold(),
                               "public" if scope == "global" else scope,
                               parameters(match.group("rest")), line))
    return found


def constants(text: str) -> dict[str, tuple[str, str]]:
    """The String constants declared at module level: name -> (scope, value)."""
    found = {}
    for _, code in code_lines(text):
        match = CONSTANT.match(code)
        if match is not None:
            found[match.group("name")] = ((match.group("scope") or "private").casefold(),
                                          match.group("value").replace('""', '"'))
    return found


def slot_problem(position: int, slot: Slot, parameter: Parameter) -> str | None:
    label = f"parameter {position} ({parameter.name})"
    if parameter.special:
        return f"{label} must not be Optional, ParamArray or an array"
    if slot.returned and parameter.mode == "byval":
        return f"{label} must be ByRef: Office returns the value through it"
    if parameter.type_key not in slot.types:
        return f"{label} must be declared {slot.text}, not {parameter.type or 'untyped'}"
    return None


def signature_findings(use: Use, procedure: Procedure) -> list[Finding]:
    title = f"{procedure.name}, named by {use.where()},"
    expected = f"Public Sub {use.name}({', '.join(slot.text for slot in shape(use))})"
    if procedure.kind != "sub" or procedure.scope != "public":
        return [Finding(CALLBACKS, procedure.line, f"{title} must be a Public Sub: {expected}")]
    if procedure.parameters is None:
        return [Finding(CALLBACKS, procedure.line, f"{title} has a parameter list this check "
                        "cannot read")]
    if len(procedure.parameters) != len(shape(use)):
        return [Finding(CALLBACKS, procedure.line, f"{title} takes {len(procedure.parameters)} "
                        f"parameter(s): {expected}")]
    problems = [slot_problem(position, slot, parameter) for position, (slot, parameter)
                in enumerate(zip(shape(use), procedure.parameters, strict=True), 1)]
    return [Finding(CALLBACKS, procedure.line, f"{title} {problem}: {expected}")
            for problem in problems if problem]


def callback_findings(uses: Sequence[Use], declared: Sequence[Procedure]) -> list[Finding]:
    findings: list[Finding] = []
    by_name: dict[str, Procedure] = {}
    for procedure in declared:
        if procedure.name.casefold() in by_name:
            findings.append(Finding(CALLBACKS, procedure.line, f"{procedure.name} is declared "
                                    "twice"))
        by_name.setdefault(procedure.name.casefold(), procedure)
    named: set[str] = set()
    for use in uses:
        if not VBA_NAME.fullmatch(use.name):
            findings.append(Finding(CUSTOM_UI, None, f"{use.where()} names {use.name!r}, not a "
                                    "plain VBA procedure name"))
            continue
        named.add(use.name.casefold())
        target = by_name.get(use.name.casefold())
        if target is None:
            findings.append(Finding(CUSTOM_UI, None, f"{use.where()} names {use.name}, which "
                                    f"is no procedure of {CALLBACKS}"))
        else:
            findings += signature_findings(use, target)
    findings += [Finding(CALLBACKS, procedure.line, f"{procedure.name} takes a Ribbon "
                         f"parameter, but no attribute of {CUSTOM_UI} names it")
                 for procedure in declared if procedure.scope == "public"
                 and procedure.ribbon_shaped and procedure.name.casefold() not in named]
    return findings


def check(root: Path) -> tuple[list[Finding], str]:
    if not (root / CUSTOM_UI).is_file():
        return [], f"no {CUSTOM_UI}: nothing to check"
    element, findings = load_xml(root)
    if element is None:
        return findings, f"{CUSTOM_UI} not read"
    findings, uses, tags = xml_findings(element)
    for path in (CALLBACKS, ACTIONS):
        if not (root / path).is_file():
            return [*findings, Finding(path, None, "is missing")], "callbacks not read"
    ribbon = (root / CALLBACKS).read_bytes().decode("ascii", errors="replace")
    declared = procedures(ribbon)
    # The menus that M_Ribbon builds at run time name their onAction through ON_ACTION.
    built = constants(ribbon).get("ON_ACTION")
    menu_uses = [Use(built[1], "onAction", "button", None)] if built else []
    findings += callback_findings([*uses, *menu_uses], declared)
    actions = {value for name, (scope, value) in
               constants((root / ACTIONS).read_bytes().decode("ascii", errors="replace")).items()
               if scope == "public" and name.startswith("ACTION_") and name != FUNCTION_PREFIX}
    findings += [Finding(CUSTOM_UI, None, f"tag {tag!r} is no action text of {ACTIONS} "
                         f"(its Public Const ACTION_ values: {sorted(actions)})")
                 for tag in tags if tag not in actions]
    return findings, (f"{len(uses)} callback use(s), {len(tags)} action tag(s) of {CUSTOM_UI} "
                      f"against {CALLBACKS} and {ACTIONS}")


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, counted = check(args.root.resolve())
    return report(GATE, findings, counted)


if __name__ == "__main__":
    sys.exit(main())
