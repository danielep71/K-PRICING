"""A synthetic, macro-free layout template, written as Excel writes one.

The builder and protection tests write it to src/workbook/Template.xlsx of a
temporary repository, with its own build configuration, build_configuration(),
as tools/devbuild/workbook.json: its sheets and names are synthetic, so these
tests do not depend on the sheets of the repository's workbook contract. Its
parts declare the namespaces that Excel declares, including one that only
mc:Ignorable names, so a test can see that the build keeps them. It carries no
personal document metadata, so it passes tools/check_binaries.py.
"""

from __future__ import annotations

import io
import json
import zipfile
from dataclasses import dataclass, field
from typing import Any

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE = "http://schemas.openxmlformats.org/package/2006/relationships"
ROOT_NAMESPACES = (f'xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}" '
                   'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
                   'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" '
                   'mc:Ignorable="x14ac"')
HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
# Cell formats (cellXfs): 0 the default, locked; 1 unlocked, for input cells;
# 2 explicitly locked; 3 unlocked but with applyProtection="0", so locked by its style.
STYLES = (f'{HEADER}<styleSheet {ROOT_NAMESPACES}>'
          '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
          '<fills count="2"><fill><patternFill patternType="none"/></fill>'
          '<fill><patternFill patternType="gray125"/></fill></fills>'
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border>'
          '</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" '
          'borderId="0"/></cellStyleXfs><cellXfs count="4">'
          '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
          '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">'
          '<protection locked="0"/></xf>'
          '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">'
          '<protection locked="1"/></xf>'
          '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="0">'
          '<protection locked="0"/></xf></cellXfs>'
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
          '</styleSheet>')
UNLOCKED, LOCKED, IGNORED = 1, 2, 3


@dataclass
class Sheet:
    tab: str
    # Cell reference -> (style index, shared string index or None).
    cells: dict[str, tuple[int, int | None]] = field(default_factory=dict)
    protection: dict[str, str] | None = None
    # (first column, last column, style index) of each <col>.
    columns: tuple[tuple[int, int, int], ...] = ()


def default_sheets() -> list[Sheet]:
    """The sheets of build_configuration(), laid out with styles and a title."""
    return [Sheet("Readme", {"A1": (0, 0), "A3": (0, 1), "B3": (LOCKED, None), "A4": (0, 1)}),
            Sheet("Inputs", {"A1": (0, 2), "B5": (0, 3), "C5": (UNLOCKED, None)},
                  {"sheet": "1", "objects": "1", "scenarios": "1"})]


DEFAULT_NAMES = {"build_commit": "Readme!$B$3", "build_time": "Readme!$B$4",
                 "build_configuration": "Readme!$B$5", "cfg_reporting_date": "Inputs!$C$5",
                 "Inputs_Settings": "Inputs!$B$4"}
STRINGS = ("Synthetic workbook", "Commit", "Inputs", "Reporting date")


def build_configuration() -> dict[str, Any]:
    """A build configuration whose sheets and names are those of default_sheets()
    and DEFAULT_NAMES, in the format of tools/devbuild/workbook.json; a new copy
    on each call, so a test may change it."""
    return {
        "schema_version": 1,
        "description": "Synthetic build configuration of tests/workbook_template.py.",
        "workbook_code_name": "ThisWorkbook",
        "sheets": [{"tab": "Readme", "code_name": "shReadme", "role": "Reference"},
                   {"tab": "Inputs", "code_name": "shInputs", "role": "Input"}],
        "names": [
            {"name": "build_commit", "sheet": "shReadme", "cell": "B3", "label": "Commit"},
            {"name": "build_time", "sheet": "shReadme", "cell": "B4", "label": "Commit time"},
            {"name": "build_configuration", "sheet": "shReadme", "cell": "B5",
             "label": "Configuration"},
            {"name": "cfg_reporting_date", "sheet": "shInputs", "cell": "C5",
             "label": "Reporting date"},
            {"name": "Inputs_Settings", "sheet": "shInputs", "cell": "B4"}],
        "stamp": {"commit": "build_commit", "time": "build_time",
                  "configuration": "build_configuration"},
        "configurations": {
            "production": ["src/core", "src/classes", "src/modules", "src/workbook"],
            "development": ["src/core", "src/classes", "src/modules", "src/workbook",
                            "tests/modules", "examples/modules"]},
    }


def build_configuration_bytes() -> bytes:
    """build_configuration() as the file tools/devbuild/workbook.json holds it."""
    return json.dumps(build_configuration(), indent=2).encode()


def split(reference: str) -> tuple[str, int]:
    letters = reference.rstrip("0123456789")
    return letters, int(reference[len(letters):])


def sheet_xml(sheet: Sheet) -> str:
    rows: dict[int, list[str]] = {}
    for reference, (style, text) in sorted(
            sheet.cells.items(), key=lambda item: (split(item[0])[1], len(item[0]), item[0])):
        rows.setdefault(split(reference)[1], []).append(
            f'<c r="{reference}" s="{style}"/>' if text is None else
            f'<c r="{reference}" s="{style}" t="s"><v>{text}</v></c>')
    body = "".join(f'<row r="{row}" spans="1:3" x14ac:dyDescent="0.25">{"".join(cells)}</row>'
                   for row, cells in sorted(rows.items()))
    columns = "".join(f'<col min="{low}" max="{high}" width="12" style="{style}"/>'
                      for low, high, style in sheet.columns)
    protection = "" if sheet.protection is None else "<sheetProtection " + " ".join(
        f'{name}="{value}"' for name, value in sheet.protection.items()) + "/>"
    return (f'{HEADER}<worksheet {ROOT_NAMESPACES}><dimension ref="A1:C5"/>'
            '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
            + (f"<cols>{columns}</cols>" if columns else "")
            + f"<sheetData>{body}</sheetData>{protection}"
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
            'footer="0.3"/></worksheet>')


def template(sheets: list[Sheet] | None = None,
             names: dict[str, str] | None = None) -> bytes:
    """The template package, by default with the sheets and names of
    build_configuration()."""
    sheets = default_sheets() if sheets is None else sheets
    names = DEFAULT_NAMES if names is None else names
    count = len(sheets)
    defined = "".join(f'<definedName name="{name}">{reference}</definedName>'
                      for name, reference in sorted(names.items()))
    defined += '<definedName name="_xlnm.Print_Area" localSheetId="1">Inputs!$A$1:$C$5' \
               "</definedName>"
    parts = {
        "[Content_Types].xml": (
            f'{HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
            'content-types"><Default Extension="rels" ContentType="application/'
            'vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" '
            'ContentType="application/xml"/><Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.'
            'sheet.main+xml"/>' + "".join(
                f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/'
                'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                for index in range(1, count + 1))
            + '<Override PartName="/xl/styles.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/'
            'vnd.openxmlformats-package.core-properties+xml"/></Types>'),
        "_rels/.rels": (
            f'{HEADER}<Relationships xmlns="{PACKAGE}"><Relationship Id="rId3" '
            'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/'
            'core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" '
            f'Type="{RELATIONSHIPS}/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>"),
        "xl/workbook.xml": (
            f'{HEADER}<workbook {ROOT_NAMESPACES}><fileVersion appName="xl" lastEdited="7"/>'
            '<workbookPr defaultThemeVersion="166925"/><bookViews><workbookView xWindow="0" '
            'yWindow="0" windowWidth="28800" windowHeight="12300"/></bookViews><sheets>'
            + "".join(f'<sheet name="{sheet.tab}" sheetId="{index}" r:id="rId{index}"/>'
                      for index, sheet in enumerate(sheets, 1))
            + f'</sheets><definedNames>{defined}</definedNames><calcPr calcId="191029"/>'
            "</workbook>"),
        "xl/_rels/workbook.xml.rels": (
            f'{HEADER}<Relationships xmlns="{PACKAGE}">' + "".join(
                f'<Relationship Id="rId{index}" Type="{RELATIONSHIPS}/worksheet" '
                f'Target="worksheets/sheet{index}.xml"/>' for index in range(1, count + 1))
            + f'<Relationship Id="rId{count + 1}" Type="{RELATIONSHIPS}/styles" '
            f'Target="styles.xml"/><Relationship Id="rId{count + 2}" '
            f'Type="{RELATIONSHIPS}/sharedStrings" Target="sharedStrings.xml"/>'
            "</Relationships>"),
        "xl/styles.xml": STYLES,
        "xl/sharedStrings.xml": (
            f'{HEADER}<sst xmlns="{MAIN}" count="{len(STRINGS)}" '
            f'uniqueCount="{len(STRINGS)}">'
            + "".join(f"<si><t>{text}</t></si>" for text in STRINGS) + "</sst>"),
        "docProps/core.xml": (
            f'{HEADER}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/'
            '2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">'
            "<dc:title>Synthetic layout</dc:title></cp:coreProperties>"),
    }
    for index, sheet in enumerate(sheets, 1):
        parts[f"xl/worksheets/sheet{index}.xml"] = sheet_xml(sheet)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, text in parts.items():
            archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0)), text)
    return buffer.getvalue()
