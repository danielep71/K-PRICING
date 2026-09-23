#!/usr/bin/env python3
"""Validate explicit VBA public API declarations and the checked-in manifest."""

from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

from _gatelib import git_bytes as git
from _gatelib import parse_report_args as parse_args
from _gatelib import run_gate
from check_vba_conditionals import reachable_sources

CONFIG_PATH = ".github/repository-profile.json"
TOOL_NAME = "VBA public API"
VBA_SUFFIXES = {".bas", ".cls", ".frm"}
SIG_PREFIX = "# SIG\t"
PROPERTY_KINDS = {"property get", "property let", "property set"}

PROC = re.compile(
    r"^\s*(?:(?P<vis>Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"(?P<name>[A-Za-z_]\w*)\b",
    re.I,
)
END_PROC = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.I)
BLOCK = re.compile(r"^\s*Public\s+(?P<kind>Enum|Type)\s+(?P<name>[A-Za-z_]\w*)\b", re.I)
END_BLOCK = {
    "enum": re.compile(r"^\s*End\s+Enum\b", re.I),
    "type": re.compile(r"^\s*End\s+Type\b", re.I),
}
DECLARE = re.compile(
    r"^\s*(?:Public|Global)\s+Declare\s+(?:PtrSafe\s+)?"
    r"(?P<kind>Function|Sub)\s+(?P<name>[A-Za-z_]\w*)\b",
    re.I,
)
EVENT = re.compile(r"^\s*(?:Public|Global)\s+Event\s+(?P<name>[A-Za-z_]\w*)\b", re.I)
CONST = re.compile(
    r"^\s*(?:Public|Global)\s+Const\s+(?P<name>[A-Za-z_]\w*)\b(?P<rest>.*)$",
    re.I,
)
VARIABLE = re.compile(
    r"^\s*(?:Public|Global)\s+(?P<withevents>WithEvents\s+)?"
    r"(?P<name>[A-Za-z_]\w*)\b(?P<rest>.*)$",
    re.I,
)
IMPLICIT = re.compile(
    r"^\s*(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set)|"
    r"Enum|Type|Event|Declare\s+(?:PtrSafe\s+)?(?:Function|Sub))\b",
    re.I,
)


def tracked_vba(root: Path) -> list[str]:
    completed = git(root, "ls-files", "-z")
    if completed.returncode:
        raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
    return sorted(
        path.decode("utf-8", errors="surrogateescape")
        for path in completed.stdout.split(b"\0")
        if path
        and Path(path.decode("utf-8", errors="surrogateescape")).suffix.casefold()
        in VBA_SUFFIXES
    )


def strip_vba(raw: str) -> str:
    output: list[str] = []
    in_string = False
    index = 0
    while index < len(raw):
        character = raw[index]
        if character == '"':
            if in_string and index + 1 < len(raw) and raw[index + 1] == '"':
                output.extend(('"', '"'))
                index += 2
                continue
            in_string = not in_string
        elif character == "'" and not in_string:
            break
        output.append(character)
        index += 1
    text = "".join(output)
    return "" if re.match(r"^\s*Rem(?:\s|$)", text, re.I) else text.rstrip()


def logical(lines: list[str]) -> list[tuple[int, int, str]]:
    result: list[tuple[int, int, str]] = []
    buffer: list[str] = []
    start = 0
    for number, raw in enumerate(lines, 1):
        code = strip_vba(raw)
        if not buffer:
            start = number
        if re.search(r"\s_\s*$", code):
            buffer.append(re.sub(r"\s_\s*$", " ", code))
            continue
        buffer.append(code)
        result.append((start, number, " ".join(item.strip() for item in buffer)))
        buffer.clear()
    if buffer:
        result.append((start, len(lines), " ".join(item.strip() for item in buffer)))
    return result


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip())


def key(component: str, kind: str, name: str) -> str:
    return f"{component}\t{kind}\t{name}"


def top_comma(text: str) -> bool:
    depth = 0
    in_string = False
    index = 0
    while index < len(text):
        character = text[index]
        if character == '"':
            if in_string and index + 1 < len(text) and text[index + 1] == '"':
                index += 2
                continue
            in_string = not in_string
        elif not in_string:
            if character == "(":
                depth += 1
            elif character == ")" and depth:
                depth -= 1
            elif character == "," and depth == 0:
                return True
        index += 1
    return False


def record(
    output: list[dict[str, Any]],
    supported: bool,
    component: str,
    kind: str,
    name: str,
    signature: str,
    path: str,
    line: int,
) -> None:
    if supported:
        output.append(
            {
                "component": component,
                "kind": kind,
                "name": name,
                "signature": signature,
                "path": path,
                "line": line,
            }
        )


def _parse_active_component(
    path: str, text: str, supported: bool
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    component = Path(path).stem
    declarations: list[dict[str, Any]] = []
    findings: list[dict[str, Any]] = []
    statements = logical(text.splitlines())
    inside_procedure = False
    index = 0
    while index < len(statements):
        line, _, code = statements[index]
        stripped = code.strip()
        if not stripped or stripped.startswith(("Attribute ", "Option ", "#")):
            index += 1
            continue

        match = PROC.match(code)
        if match and " declare " not in f" {code.casefold()} ":
            visibility = (match.group("vis") or "").casefold()
            if not visibility:
                findings.append(
                    {
                        "path": path,
                        "line": line,
                        "message": (
                            "Implicit public procedure is prohibited: "
                            f"{match.group('name')}."
                        ),
                    }
                )
            if visibility == "public":
                kind = " ".join(item.capitalize() for item in match.group("kind").split())
                record(
                    declarations,
                    supported,
                    component,
                    kind,
                    match.group("name"),
                    norm(code),
                    path,
                    line,
                )
            inside_procedure = True
            index += 1
            continue
        if END_PROC.match(code):
            inside_procedure = False
            index += 1
            continue
        if inside_procedure:
            index += 1
            continue

        if IMPLICIT.match(code):
            findings.append(
                {
                    "path": path,
                    "line": line,
                    "message": (
                        "Implicit public module-level declaration is prohibited: "
                        f"{norm(code)}"
                    ),
                }
            )
            index += 1
            continue

        match = BLOCK.match(code)
        if match:
            kind = match.group("kind").capitalize()
            name = match.group("name")
            body = [norm(code)]
            close = END_BLOCK[kind.casefold()]
            end_index = index + 1
            while end_index < len(statements) and not close.match(statements[end_index][2]):
                if statements[end_index][2].strip():
                    body.append(norm(statements[end_index][2]))
                end_index += 1
            if end_index >= len(statements):
                findings.append(
                    {
                        "path": path,
                        "line": line,
                        "message": f"Public {kind} {name} is not closed.",
                    }
                )
                index += 1
                continue
            body.append(norm(statements[end_index][2]))
            record(
                declarations,
                supported,
                component,
                kind,
                name,
                " | ".join(body),
                path,
                line,
            )
            index = end_index + 1
            continue

        match = DECLARE.match(code)
        if match:
            kind = (
                "Declare Function"
                if match.group("kind").casefold() == "function"
                else "Declare Sub"
            )
            record(
                declarations,
                supported,
                component,
                kind,
                match.group("name"),
                norm(code),
                path,
                line,
            )
            index += 1
            continue

        match = EVENT.match(code)
        if match:
            record(
                declarations,
                supported,
                component,
                "Event",
                match.group("name"),
                norm(code),
                path,
                line,
            )
            index += 1
            continue

        match = CONST.match(code)
        if match:
            if top_comma(match.group("rest")):
                findings.append(
                    {
                        "path": path,
                        "line": line,
                        "message": (
                            "Public Const declarations must contain one identifier per statement."
                        ),
                    }
                )
            else:
                record(
                    declarations,
                    supported,
                    component,
                    "Const",
                    match.group("name"),
                    norm(code),
                    path,
                    line,
                )
            index += 1
            continue

        match = VARIABLE.match(code)
        if match:
            if top_comma(match.group("rest")):
                findings.append(
                    {
                        "path": path,
                        "line": line,
                        "message": (
                            "Public variable declarations must contain one identifier per statement."
                        ),
                    }
                )
            else:
                kind = "WithEvents Variable" if match.group("withevents") else "Variable"
                record(
                    declarations,
                    supported,
                    component,
                    kind,
                    match.group("name"),
                    norm(code),
                    path,
                    line,
                )
            index += 1
            continue

        index += 1
    return declarations, findings


def parse_component(
    path: str, text: str, supported: bool
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    sources, findings = reachable_sources(path, text)
    declarations: dict[tuple[int, str, str, str], dict[str, Any]] = {}
    for environment, source in sources.items():
        parsed, errors = _parse_active_component(path, source, supported)
        for error in errors:
            if error not in findings:
                findings.append(error)
        for item in parsed:
            identity = (item["line"], item["kind"], item["name"], item["signature"])
            if identity not in declarations:
                declarations[identity] = {**item, "environments": []}
            declarations[identity]["environments"].append(environment)
    return list(declarations.values()), findings


def read_manifest(
    root: Path, relative: str
) -> tuple[set[str], dict[str, set[str]], list[dict[str, Any]]]:
    path = root / relative
    if not path.is_file():
        return set(), {}, [
            {"path": relative, "message": "Configured public API manifest is missing."}
        ]
    rows: set[str] = set()
    signatures: dict[str, set[str]] = {}
    findings: list[dict[str, Any]] = []
    for number, raw in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        if not raw.strip():
            continue
        if raw.startswith(SIG_PREFIX):
            fields = raw[len(SIG_PREFIX) :].split("\t", 3)
            if len(fields) != 4:
                findings.append(
                    {"path": relative, "line": number, "message": "Malformed # SIG record."}
                )
                continue
            declaration_key = key(fields[0], fields[1], fields[2])
            canonical = next((existing for existing in signatures
                              if existing.casefold() == declaration_key.casefold()), declaration_key)
            variants = signatures.setdefault(canonical, set())
            if fields[3] in variants:
                findings.append({"path": relative, "line": number,
                                 "message": f"Duplicate signature record: {declaration_key}"})
            variants.add(fields[3])
            continue
        if raw.lstrip().startswith("#"):
            continue
        fields = raw.split("\t")
        if len(fields) != 3 or any(not item for item in fields):
            findings.append(
                {
                    "path": relative,
                    "line": number,
                    "message": (
                        "Manifest declaration rows require component, kind, and name."
                    ),
                }
            )
            continue
        declaration_key = key(*fields)
        if any(existing.casefold() == declaration_key.casefold() for existing in rows):
            findings.append(
                {
                    "path": relative,
                    "line": number,
                    "message": f"Duplicate manifest declaration: {declaration_key}",
                }
            )
        rows.add(declaration_key)
    return rows, signatures, findings


def is_property(kind: str) -> bool:
    return kind.casefold() in PROPERTY_KINDS


def property_pair(first: dict[str, Any], second: dict[str, Any]) -> bool:
    return (
        str(first["component"]).casefold() == str(second["component"]).casefold()
        and is_property(str(first["kind"]))
        and is_property(str(second["kind"]))
        and str(first["kind"]).casefold() != str(second["kind"]).casefold()
    )


def run_check(root: Path) -> dict[str, Any]:
    config = json.loads((root / CONFIG_PATH).read_text(encoding="utf-8"))
    vba = config["vba"]
    components: dict[str, str] = vba["components"]
    manifest = vba["public_api_manifest"]
    findings: list[dict[str, Any]] = []

    if manifest not in config["required_paths"]:
        findings.append(
            {
                "path": CONFIG_PATH,
                "message": (
                    "public_api_manifest must be a global required_path for every profile."
                ),
            }
        )
    for profile, entry in config["profiles"].items():
        if entry["vba_contract"]["minimum_roles"].get("public", 0) < 1:
            findings.append(
                {
                    "path": CONFIG_PATH,
                    "message": (
                        f"Profile {profile!r} must require at least one public-role component."
                    ),
                }
            )

    declarations: list[dict[str, Any]] = []
    tracked = set(tracked_vba(root))
    for path, role in sorted(components.items()):
        if path not in tracked or not (root / path).is_file():
            continue
        parsed, errors = parse_component(
            path,
            (root / path).read_bytes().decode("cp1252"),
            role == "public",
        )
        declarations.extend(parsed)
        findings.extend(errors)

    keys: dict[str, list[dict[str, Any]]] = {}
    names: dict[str, list[dict[str, Any]]] = {}
    for declaration in declarations:
        declaration_key = key(
            str(declaration["component"]),
            str(declaration["kind"]),
            str(declaration["name"]),
        )
        canonical = next((existing for existing in keys
                          if existing.casefold() == declaration_key.casefold()), declaration_key)
        variants = keys.setdefault(canonical, [])
        if any(set(previous["environments"]) & set(declaration["environments"])
               for previous in variants):
            findings.append({
                "path": declaration["path"], "line": declaration["line"],
                "message": "Public declaration appears more than once: " + declaration_key,
            })
        variants.append(declaration)
        if Path(str(declaration["path"])).suffix.casefold() == ".bas":
            name_key = str(declaration["name"]).casefold()
            prior = names.setdefault(name_key, [])
            for previous in prior:
                if (set(previous["environments"]) & set(declaration["environments"])
                        and not property_pair(previous, declaration)):
                    findings.append(
                        {
                            "path": declaration["path"],
                            "line": declaration["line"],
                            "message": (
                                "Public standard-module name "
                                f"{declaration['name']!r} collides with "
                                f"{previous['component']}.{previous['name']} "
                                f"({previous['kind']})."
                            ),
                        }
                    )
            prior.append(declaration)

    rows, signatures, errors = read_manifest(root, manifest)
    findings.extend(errors)
    actual = {item.casefold(): item for item in keys}
    row_map = {item.casefold(): item for item in rows}
    signature_map = {item.casefold(): item for item in signatures}
    for folded, declaration_key in actual.items():
        if folded not in row_map:
            findings.append(
                {
                    "path": manifest,
                    "message": f"Public declaration is not recorded: {declaration_key}",
                }
            )
        signature_key = signature_map.get(folded)
        if signature_key is None:
            findings.append(
                {
                    "path": manifest,
                    "message": (
                        "Normalized signature record is missing: "
                        f"{declaration_key}"
                    ),
                }
            )
        elif signatures[signature_key] != {item["signature"] for item in keys[declaration_key]}:
            findings.append(
                {
                    "path": manifest,
                    "message": (
                        f"Signature mismatch for {declaration_key}: expected "
                        f"{sorted({item['signature'] for item in keys[declaration_key]})!r}, recorded "
                        f"{sorted(signatures[signature_key])!r}"
                    ),
                }
            )
    for folded, declaration_key in row_map.items():
        if folded not in actual:
            findings.append(
                {
                    "path": manifest,
                    "message": (
                        "Manifest declaration is not present in supported source: "
                        f"{declaration_key}"
                    ),
                }
            )
    for folded, declaration_key in signature_map.items():
        if folded not in actual:
            findings.append(
                {
                    "path": manifest,
                    "message": f"Stale signature record: {declaration_key}",
                }
            )

    evidence = sorted(
        declarations,
        key=lambda declaration: (
            str(declaration["component"]).casefold(),
            str(declaration["name"]).casefold(),
            str(declaration["kind"]).casefold(),
        ),
    )
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": "pass" if not findings else "fail",
        "manifest": manifest,
        "policy": {
            "required_profiles": sorted(config["profiles"]),
            "required_from": "initialization",
            "implicit_visibility": "prohibited",
            "const_and_variable_form": "one identifier per statement",
            "property_accessors": "same-name Get/Let/Set allowed only within one component",
        },
        "declarations": evidence,
        "findings": findings,
    }


def markdown_report(report: dict[str, Any]) -> str:
    lines = [
        "## VBA public API",
        "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **Manifest:** `{report['manifest']}`",
        f"- **Declarations:** {len(report['declarations'])}",
        "- **Profiles:** "
        + ", ".join(f"`{item}`" for item in report["policy"]["required_profiles"]),
        "- **Manifest required from:** initialization",
        "- **Implicit public visibility:** prohibited",
        "- **Public Const/variable form:** one identifier per statement",
        "- **Property accessors:** same-name Get/Let/Set allowed within one component",
        f"- **Findings:** {len(report['findings'])}",
    ]
    if report["declarations"]:
        lines += [
            "",
            "| Component | Kind | Name | Signature |",
            "| --- | --- | --- | --- |",
        ]
        for declaration in report["declarations"]:
            signature = str(declaration["signature"]).replace("|", "\\|")
            lines.append(
                f"| {declaration['component']} | {declaration['kind']} | "
                f"{declaration['name']} | `{signature}` |"
            )
    if report["findings"]:
        lines += ["", "### Findings", ""]
        for item in report["findings"]:
            location = str(item.get("path", "."))
            if item.get("line"):
                location += f":{item['line']}"
            lines.append(f"- `{location}` — {item['message']}")
    return "\n".join(lines) + "\n"


def init_fixture(
    root: Path, facade: str, manifest: list[str], other: str | None = None
) -> None:
    for path in (
        root / ".github",
        root / "src/modules",
        root / "src/core",
        root / "tests/modules",
        root / "docs",
    ):
        path.mkdir(parents=True, exist_ok=True)
    components = {
        "src/core/Core.bas": "internal",
        "src/modules/Facade.bas": "public",
        "tests/modules/Tests.bas": "test",
    }
    if other is not None:
        components["src/modules/Other.bas"] = "public"
    config = {
        "required_paths": ["docs/PUBLIC_API.txt"],
        "profiles": {
            profile: {"vba_contract": {"minimum_roles": {"public": 1}}}
            for profile in ("application", "library", "ui-component")
        },
        "vba": {
            "components": components,
            "public_api_manifest": "docs/PUBLIC_API.txt",
        },
    }
    (root / CONFIG_PATH).write_text(
        json.dumps(config, indent=2) + "\n", encoding="utf-8"
    )
    files = {
        "src/modules/Facade.bas": facade,
        "src/core/Core.bas": (
            'Attribute VB_Name = "Core"\n'
            "Option Explicit\n"
            "Option Private Module\n"
            "Public Function InternalOnly() As Long\n"
            "End Function\n"
        ),
        "tests/modules/Tests.bas": (
            'Attribute VB_Name = "Tests"\n'
            "Option Explicit\n"
            "Public Sub RunTests()\n"
            "End Sub\n"
        ),
    }
    if other is not None:
        files["src/modules/Other.bas"] = other
    for relative, text in files.items():
        (root / relative).write_bytes(text.replace("\n", "\r\n").encode("cp1252"))
    (root / "docs/PUBLIC_API.txt").write_text(
        "\n".join(manifest) + "\n", encoding="utf-8"
    )
    for command in (
        ("init", "-b", "main"),
        ("config", "user.name", "API Self-Test"),
        ("config", "user.email", "api@example.invalid"),
        ("add", "--all"),
    ):
        completed = git(root, *command)
        if completed.returncode:
            raise RuntimeError(
                completed.stderr.decode("utf-8", errors="replace").strip()
            )


def fixture(
    facade: str, manifest: list[str], other: str | None = None
) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix="vba-api-") as temporary:
        root = Path(temporary)
        init_fixture(root, facade, manifest, other)
        return run_check(root)


def run_self_test() -> int:
    facade = '''Attribute VB_Name = "Facade"
Option Explicit
Public Const C As Long = 7
Public Event Changed(ByVal value As Long)
Public Declare PtrSafe Function Tick Lib "kernel32" () As Long
Public Number As Long
Public WithEvents Source As Object
Public Enum Mode
    ModeA = 1
End Enum
Public Type Pair
    Left As Long
    Right As Long
End Type
Public Function Echo( _
    ByVal value As Long) As Long
    Echo = value
End Function
Public Property Get Current() As Long
    Current = Number
End Property
Public Property Let Current(ByVal value As Long)
    Number = value
End Property
'''
    manifest = [
        "# Supported VBA declarations: <component><TAB><kind><TAB><name>",
        "Facade\tConst\tC",
        "# SIG\tFacade\tConst\tC\tPublic Const C As Long = 7",
        "Facade\tEvent\tChanged",
        "# SIG\tFacade\tEvent\tChanged\tPublic Event Changed(ByVal value As Long)",
        "Facade\tDeclare Function\tTick",
        '# SIG\tFacade\tDeclare Function\tTick\tPublic Declare PtrSafe Function Tick Lib "kernel32" () As Long',
        "Facade\tVariable\tNumber",
        "# SIG\tFacade\tVariable\tNumber\tPublic Number As Long",
        "Facade\tWithEvents Variable\tSource",
        "# SIG\tFacade\tWithEvents Variable\tSource\tPublic WithEvents Source As Object",
        "Facade\tEnum\tMode",
        "# SIG\tFacade\tEnum\tMode\tPublic Enum Mode | ModeA = 1 | End Enum",
        "Facade\tType\tPair",
        "# SIG\tFacade\tType\tPair\tPublic Type Pair | Left As Long | Right As Long | End Type",
        "Facade\tFunction\tEcho",
        "# SIG\tFacade\tFunction\tEcho\tPublic Function Echo( ByVal value As Long) As Long",
        "Facade\tProperty Get\tCurrent",
        "# SIG\tFacade\tProperty Get\tCurrent\tPublic Property Get Current() As Long",
        "Facade\tProperty Let\tCurrent",
        "# SIG\tFacade\tProperty Let\tCurrent\tPublic Property Let Current(ByVal value As Long)",
    ]
    failures: list[str] = []
    tests: list[tuple[str, dict[str, Any], str, str | None]] = []
    tests.append(("positive", fixture(facade, manifest), "pass", None))
    tests.append(
        (
            "implicit",
            fixture(facade.replace("Public Function Echo", "Function Echo"), manifest),
            "fail",
            "Implicit public procedure",
        )
    )
    tests.append(
        (
            "missing-signature",
            fixture(
                facade,
                [
                    item
                    for item in manifest
                    if not item.startswith("# SIG\tFacade\tFunction\tEcho")
                ],
            ),
            "fail",
            "Normalized signature record is missing",
        )
    )
    tests.append(
        (
            "multi-var",
            fixture(
                facade.replace(
                    "Public Number As Long",
                    "Public Number As Long, Other As Long",
                ),
                manifest,
            ),
            "fail",
            "Public variable declarations must contain one identifier",
        )
    )
    tests.append(
        (
            "multi-const",
            fixture(
                facade.replace(
                    "Public Const C As Long = 7",
                    "Public Const C As Long = 7, D As Long = 8",
                ),
                manifest,
            ),
            "fail",
            "Public Const declarations must contain one identifier",
        )
    )
    duplicate = facade + "\nPublic Function Echo(ByVal value As String) As String\nEnd Function\n"
    tests.append(
        ("same-component-collision", fixture(duplicate, manifest), "fail", "collides")
    )
    other = '''Attribute VB_Name = "Other"
Option Explicit
Public Function Echo(ByVal value As Long) As Long
    Echo = value
End Function
'''
    tests.append(
        (
            "cross-component-collision",
            fixture(facade, manifest, other),
            "fail",
            "collides",
        )
    )
    modern = 'Public Declare PtrSafe Function Tick Lib "kernel32" () As Long'
    legacy = modern.replace("PtrSafe ", "")
    conditional = facade.replace(modern, f"#If VBA7 Then\n{modern}\n#Else\n{legacy}\n#End If")
    variant_manifest = manifest + ["# SIG\tFacade\tDeclare Function\tTick\t" + legacy]
    tests.append(("conditional-declare", fixture(conditional, variant_manifest), "pass", None))
    tests.append(("conditional-missing-signature", fixture(conditional, manifest), "fail", "Signature mismatch"))
    tests.append(("conditional-stale-signature", fixture(facade, variant_manifest), "fail", "Signature mismatch"))
    tests.append(("duplicate-signature", fixture(facade, manifest + [manifest[6]]), "fail", "Duplicate signature"))
    overlap = conditional + "\n" + modern + "\n"
    tests.append(("reachable-collision", fixture(overlap, variant_manifest), "fail", "collides"))
    nested = facade.replace(modern, f"#If VBA7 Then\n#If Win64 Then\n{modern}\n#Else\n{modern}\n#End If\n#ElseIf VBA6 Then\n{legacy}\n#End If")
    tests.append(("nested-elseif", fixture(nested, variant_manifest), "pass", None))
    dead = facade + "\n#If False Then\n" + modern + "\n#End If\n"
    tests.append(("unreachable-declaration", fixture(dead, manifest), "pass", None))
    tests.append(("unknown-condition", fixture(conditional.replace("#If VBA7 Then", "#If Unknown Then"), variant_manifest), "fail", "Indeterminate"))
    tests.append(("unclosed-condition", fixture(conditional.replace("#End If", ""), variant_manifest), "fail", "unclosed"))
    alternate = 'Attribute VB_Name = "Other"\n#If VBA6 Then\n' + legacy + '\n#End If\n'
    complementary = facade.replace(modern, f"#If VBA7 Then\n{modern}\n#End If")
    cross_manifest = manifest + ["Other\tDeclare Function\tTick", "# SIG\tOther\tDeclare Function\tTick\t" + legacy]
    tests.append(("exclusive-cross-component", fixture(complementary, cross_manifest, alternate), "pass", None))
    tests.append(("overlapping-cross-component", fixture(facade, cross_manifest, alternate), "fail", "collides"))
    for name, report, expected, needle in tests:
        if report["status"] != expected:
            failures.append(f"{name}: expected {expected}, got {report['status']}")
        if needle and not any(
            needle in item["message"] for item in report["findings"]
        ):
            failures.append(f"{name}: missing diagnostic {needle!r}")
    if failures:
        for failure in failures:
            print(f"[FAIL] {failure}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1
    print(
        "SELF-TEST PASS: procedures, paired properties, constants, events, declares, "
        "variables, WithEvents, enums, types, continuations, implicit-public rejection, "
        "signature drift, conditional variants, reachable collisions, unknown conditions, and one-identifier Const/variable policy passed."
    )
    return 0


def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    return run_gate(
        options,
        build=lambda: run_check(options.root),
        markdown=markdown_report,
        errors=(
            OSError,
            UnicodeError,
            RuntimeError,
            subprocess.SubprocessError,
            json.JSONDecodeError,
        ),
        self_test=run_self_test,
    )

if __name__ == "__main__":
    raise SystemExit(main())
