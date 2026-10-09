"""Read saved OOXML demo content; never execute Excel or trust formula caches.

This is a conservative comparison, not a general Excel equivalence engine.
See docs/NATIVE_DEVELOPER_TOOLS.md for the versioned normalization boundary.
"""

from __future__ import annotations

import copy
import hashlib
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any
from zipfile import ZipFile

from check_repo import MAX_XML_BYTES, _validate_xml_text
from release_provenance import relative, require

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
MAX_PACKAGE = 64 * 1024 * 1024
ERROR_CODES = {"#NULL!": 2000, "#DIV/0!": 2007, "#VALUE!": 2015,
               "#REF!": 2023, "#NAME?": 2029, "#NUM!": 2036, "#N/A": 2042,
               "#GETTING_DATA": 2043, "#SPILL!": 2045, "#CALC!": 2050}


def xml(zip_file: ZipFile, name: str) -> ET.Element:
    require(zip_file.getinfo(name).file_size <= MAX_XML_BYTES, f"oversized XML: {name}")
    text = zip_file.read(name).decode("utf-8-sig")
    error = _validate_xml_text(name, text)
    require(error is None, f"unsafe or malformed XML: {name}: {error}")
    return ET.fromstring(text)  # noqa: S314 -- existing bounded DTD/entity validator above.


def tree(node: ET.Element) -> dict[str, Any]:
    """Ignore XML indentation, prefixes and attribute order; keep leaf text exact."""
    return {"tag": node.tag, "attributes": dict(sorted(node.attrib.items())),
            "text": node.text if not len(node) or (node.text or "").strip() else "",
            "tail": node.tail if (node.tail or "").strip() else None,
            "children": [tree(child) for child in node]}


def coordinate(address: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]{1,3})([1-9][0-9]*)", address)
    require(match is not None, f"invalid cell address: {address}")
    assert match is not None
    column = 0
    for letter in match[1]:
        column = column * 26 + ord(letter) - 64
    row = int(match[2])
    require(column <= 16384 and row <= 1048576, f"cell outside Excel bounds: {address}")
    return row, column


def bounds(reference: str) -> tuple[int, int, int, int]:
    ends = reference.split(":")
    require(len(ends) in (1, 2), f"unsupported array reference: {reference}")
    start, end = coordinate(ends[0]), coordinate(ends[-1])
    require(start[0] <= end[0] and start[1] <= end[1], "reversed array reference")
    return start[0], start[1], end[0], end[1]


def value(cell: ET.Element, strings: list[ET.Element]) -> dict[str, Any]:
    kind = cell.get("t", "n")
    raw = cell.find(NS + "v")
    text = None if raw is None else raw.text
    if kind == "s":
        require(text is not None and text.isdigit(), "invalid shared string index")
        index = int(str(text))
        require(index < len(strings), "shared string index outside table")
        content = copy.deepcopy(strings[index])
        content.tag = NS + "is"
        return {"type": "string", "content": tree(content)}
    if kind == "inlineStr":
        inline = cell.find(NS + "is")
        require(inline is not None, "missing inline string")
        assert inline is not None
        return {"type": "string", "content": tree(inline)}
    if kind == "e":
        return {"type": "error", "token": text, "native_code": ERROR_CODES.get(str(text)),
                "cache_present": raw is not None}
    return {"type": kind, "value": text, "cache_present": raw is not None}


def sheet_snapshot(node: ET.Element, strings: list[ET.Element]) -> tuple[dict, dict]:
    cells = node.findall(f"{NS}sheetData/{NS}row/{NS}c")
    arrays = []
    for cell in cells:
        formula = cell.find(NS + "f")
        if formula is not None and formula.get("t") == "array":
            arrays.append(bounds(formula.get("ref", "")))
    structure: dict[str, Any] = {}
    results: dict[str, Any] = {}
    for cell in cells:
        address = cell.get("r", "")
        row, column = coordinate(address)
        require(address not in structure, f"duplicate cell: {address}")
        formula = cell.find(NS + "f")
        calculated = formula is not None or any(
            r1 <= row <= r2 and c1 <= column <= c2 for r1, c1, r2, c2 in arrays)
        attributes = {key: val for key, val in cell.attrib.items() if key not in ("r", "t")}
        other = [tree(child) for child in cell if child.tag not in
                 {NS + "v", NS + "is", NS + "f"}]
        record: dict[str, Any] = {"attributes": attributes, "other": other}
        if calculated:
            record["formula"] = None if formula is None else tree(formula)
            results[address] = value(cell, strings)
        else:
            record["input"] = value(cell, strings)
        structure[address] = record
    # Retain row heights/visibility, column widths, styles, merges, conditional
    # formatting, validations, views and extensions without guessing semantics.
    metadata = copy.deepcopy(node)
    for row_node in metadata.findall(f"{NS}sheetData/{NS}row"):
        for cell in list(row_node):
            if cell.tag == NS + "c":
                row_node.remove(cell)
    return {"cells": structure, "properties": tree(metadata)}, results


def worksheet_parts(zip_file: ZipFile, workbook: ET.Element) -> list[tuple[str, dict]]:
    relations = xml(zip_file, "xl/_rels/workbook.xml.rels")
    mapping = {}
    for item in relations:
        key = item.get("Id")
        require(key not in mapping, "duplicate workbook relationship ID")
        mapping[key] = item
    result = []
    names: set[str] = set()
    for sheet in workbook.findall(f"{NS}sheets/{NS}sheet"):
        name = sheet.get("name", "")
        require(name and name.casefold() not in names, "missing or duplicate sheet name")
        names.add(name.casefold())
        relation = mapping.get(sheet.get(REL + "id"))
        require(relation is not None, f"missing sheet relationship: {name}")
        assert relation is not None
        require(relation.get("Type") == REL[1:-1] + "/worksheet"
                and relation.get("TargetMode", "Internal") == "Internal",
                f"unsupported sheet relationship: {name}")
        target = relation.get("Target", "")
        part = target.lstrip("/") if target.startswith("/") else "xl/" + target
        require(relative(part), f"unsafe worksheet path: {part}")
        result.append((part, dict(sheet.attrib)))
    require(bool(result), "workbook contains no worksheets")
    require(len({part for part, _ in result}) == len(result), "duplicate worksheet target")
    return result


def snapshot(path: Path) -> dict[str, Any]:
    require(path.stat().st_size <= MAX_PACKAGE, "workbook exceeds package limit")
    with ZipFile(path) as package:
        names = package.namelist()
        require(len(names) == len(set(names)), "duplicate ZIP member")
        require(all(relative(name.rstrip("/")) for name in names), "unsafe ZIP member")
        require(sum(info.file_size for info in package.infolist()) <= MAX_PACKAGE,
                "expanded workbook exceeds package limit")
        require(all(not info.flag_bits & 1 for info in package.infolist()),
                "encrypted ZIP members are unsupported")
        workbook = xml(package, "xl/workbook.xml")
        require(workbook.tag == NS + "workbook", "unsupported workbook namespace")
        parts = worksheet_parts(package, workbook)
        strings = list(xml(package, "xl/sharedStrings.xml")) if "xl/sharedStrings.xml" in names else []
        sheets, results = {}, {}
        for part, attributes in parts:
            name = attributes["name"]
            node = xml(package, part)
            require(node.tag == NS + "worksheet", f"unsupported worksheet: {part}")
            sheets[name], results[name] = sheet_snapshot(node, strings)
        consumed = {part for part, _ in parts} | {"xl/workbook.xml", "xl/sharedStrings.xml"}
        excluded = {name for name in names if name.startswith("docProps/")
                    or name in {"xl/calcChain.xml", "xl/vbaProject.bin"} or name.endswith("/")}
        # Unknown parts are retained, not silently normalized away. XML is
        # compared canonically; opaque parts conservatively use exact digests.
        auxiliary = {}
        for name in sorted(set(names) - consumed - excluded):
            auxiliary[name] = (tree(xml(package, name)) if name.endswith((".xml", ".rels"))
                               else {"sha256": hashlib.sha256(package.read(name)).hexdigest()})
        return {"schema": "kpr-workbook-snapshot", "schema_version": 1,
                "structure": {"workbook": tree(workbook), "sheet_order": list(sheets),
                              "sheets": sheets, "package_parts": auxiliary},
                "cached_results": results,
                "observation": {"calculation": "NOT_OBSERVED", "excel_environment": None,
                                "excluded_parts": sorted(excluded)}}


def differences(left: Any, right: Any, path: str = "") -> list[dict[str, Any]]:
    """Return deterministic JSON-pointer findings, including sheet and cell keys."""
    if type(left) is type(right) and left == right:
        return []
    if isinstance(left, dict) and isinstance(right, dict):
        result = []
        for key in sorted(left.keys() | right.keys()):
            pointer = path + "/" + str(key).replace("~", "~0").replace("/", "~1")
            if key not in left or key not in right:
                result.append({"path": pointer, "left_present": key in left,
                               "right_present": key in right,
                               "left": left.get(key), "right": right.get(key)})
            else:
                result.extend(differences(left[key], right[key], pointer))
        return result
    if isinstance(left, list) and isinstance(right, list) and len(left) == len(right):
        return [item for index, (a, b) in enumerate(zip(left, right))
                for item in differences(a, b, f"{path}/{index}")]
    return [{"path": path, "left": left, "right": right}]


def compare(left: Path, right: Path) -> dict[str, Any]:
    a, b = snapshot(left), snapshot(right)
    structural = differences(a["structure"], b["structure"], "/structure")
    cached = differences(a["cached_results"], b["cached_results"], "/cached_results")
    return {"schema": "kpr-workbook-comparison", "schema_version": 1,
            "status": "pass" if not structural and not cached else "fail",
            "structural_match": not structural, "cached_results_match": not cached,
            "structure_differences": structural, "cached_result_differences": cached,
            "calculated_result_comparison": "NOT_RUN",
            "observations": {"left": a["observation"], "right": b["observation"]}}
