"""Tests for tools/check_workbook_protection.py, with synthetic layout templates.

The fixtures pair tests/workbook_template.py's template and build configuration
with a matrix of their sheets; only the repository tests read this repository's
own configuration, matrix and template.
"""

from __future__ import annotations

import contextlib
import copy
import io
import json
import sys
import tempfile
import time
import unittest
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))

import build_workbook  # noqa: E402
import check_workbook_protection  # noqa: E402
import workbook_template  # noqa: E402

MATRIX: dict[str, Any] = {
    "schema_version": 1,
    "sheets": {
        "shReadme": {"protection": {}, "output": ["A1:B5"]},
        "shInputs": {"protection": {"sheet": True, "objects": True, "scenarios": True},
                     "input": ["C5"], "output": ["A1:B5", "C1:C4"]},
    },
}


def run(matrix: object = MATRIX, template: bytes | None = b"default",
        ) -> tuple[list[str], str]:
    """Check a repository with the synthetic template's build configuration and the
    given files."""
    with tempfile.TemporaryDirectory(prefix="protection-") as directory:
        root = Path(directory)
        (root / "tools/devbuild").mkdir(parents=True)
        (root / "tools/devbuild/workbook.json").write_bytes(
            workbook_template.build_configuration_bytes())
        if matrix is not None:
            path = root / check_workbook_protection.MATRIX
            path.parent.mkdir(parents=True)
            path.write_text(matrix if isinstance(matrix, str) else json.dumps(matrix),
                            encoding="utf-8")
        if template is not None:
            path = root / check_workbook_protection.TEMPLATE
            path.parent.mkdir(parents=True)
            path.write_bytes(workbook_template.template() if template == b"default"
                             else template)
        return check_workbook_protection.check(root)


def changed(sheet: str, **entry: Any) -> dict[str, Any]:
    """The matrix with one sheet's entry updated."""
    matrix = copy.deepcopy(MATRIX)
    matrix["sheets"][sheet].update(entry)
    return matrix


class ProtectionCheckTest(unittest.TestCase):
    def assertFinding(self, findings: list[str], text: str) -> None:
        self.assertTrue(any(text in finding for finding in findings), findings)

    def test_repository_passes(self) -> None:
        self.assertEqual(check_workbook_protection.check(ROOT)[0], [])

    def test_the_contracts_example_matrix_fits_the_built_layout(self) -> None:
        # The example in WORKBOOK_CONTRACT.md: each setting's value cell is an
        # input, and its label and each table anchor are locked outputs.
        text = (ROOT / "docs/methodology/WORKBOOK_CONTRACT.md").read_text(encoding="utf-8")
        section = text.split("## Protection\n", 1)[1]
        example = json.loads(section.split("```json\n", 1)[1].split("```", 1)[0])
        layout = build_workbook.load_layout(ROOT)
        roles = {sheet.code_name: sheet.role for sheet in layout.sheets}
        entries = {}
        for code, item in example["sheets"].items():
            entry, problems = check_workbook_protection.entry_problems(code, item, roles)
            self.assertEqual(problems, [], code)
            assert entry is not None
            entries[code] = entry

        def kinds(sheet: str, row: int, column: int) -> set[str]:
            return {area.kind for area in entries[sheet].areas
                    if area.top <= row <= area.bottom and area.left <= column <= area.right}

        for name in layout.names:
            row, column = build_workbook.cell_position(name.cell)
            with self.subTest(name=name.name):
                if name.name.startswith("cfg_"):
                    self.assertEqual(kinds(name.sheet, row, column), {"input"})
                # A name's label, or a table anchor's own header cell, is never typed in.
                cell = (row, column - 1) if name.label else (row, column)
                self.assertEqual(kinds(name.sheet, *cell), {"output"})

    def test_without_a_template_or_a_matrix_the_check_is_inactive(self) -> None:
        self.assertEqual(run(None, None), ([], "inactive: no src/workbook/Template.xlsx and no "
                                               "docs/methodology/workbook-protection.json"))

    def test_a_template_matching_its_matrix_passes(self) -> None:
        findings, compared = run()
        self.assertEqual(findings, [])
        self.assertEqual(compared, "2 sheet(s) compared with src/workbook/Template.xlsx")

    def test_a_differing_lock_flag_fails(self) -> None:
        cases = {
            "1 cell(s) of input range B5 are locked, first B5": changed(
                "shInputs", input=["B5"], output=["A1:A5", "C1:C5"]),
            "1 cell(s) of output range C1:C5 are unlocked, first C5": changed(
                "shInputs", input=[], output=["A1:B5", "C1:C5"]),
        }
        for expected, matrix in cases.items():
            with self.subTest(expected=expected):
                self.assertFinding(run(matrix)[0], expected)
        sheets = workbook_template.default_sheets()
        sheets[1].cells["C5"] = (workbook_template.LOCKED, None)
        self.assertFinding(run(template=workbook_template.template(sheets))[0],
                           "sheet shInputs: 1 cell(s) of input range C5 are locked")
        # applyProtection="0": the cell style decides, and it is locked.
        sheets[1].cells["C5"] = (workbook_template.IGNORED, None)
        self.assertFinding(run(template=workbook_template.template(sheets))[0],
                           "input range C5 are locked")

    def test_an_empty_cell_takes_the_lock_flag_of_its_column(self) -> None:
        sheets = workbook_template.default_sheets()
        sheets[1].columns = ((4, 5, workbook_template.UNLOCKED),)
        template = workbook_template.template(sheets)
        self.assertEqual(run(changed("shInputs", input=["C5", "D1:E9"]), template)[0], [])
        self.assertFinding(run(changed("shInputs", output=["A1:B5", "C1:C4", "D9"]),
                               template)[0], "output range D9 are unlocked")

    def test_a_full_height_range_is_checked_without_visiting_each_cell(self) -> None:
        sheets = workbook_template.default_sheets()
        sheets[1].columns = ((4, 5, workbook_template.UNLOCKED),)
        template = workbook_template.template(sheets)
        matrix = changed("shInputs", input=[], output=["A1:T1048576"])
        matrix["sheets"]["shReadme"] = {"protection": {}, "input": ["A1:T1048576"]}
        started = time.perf_counter()
        findings = run(matrix, template)[0]
        self.assertLess(time.perf_counter() - started, 5)
        # Every Readme cell is locked; on Inputs, C5 and the columns D:E are unlocked.
        self.assertFinding(findings, "sheet shReadme: 20971520 cell(s) of input range "
                                     "A1:T1048576 are locked, first A1")
        self.assertFinding(findings, f"sheet shInputs: {2 * 1048576 + 1} cell(s) of output "
                                     "range A1:T1048576 are unlocked, first D1")
        self.assertEqual(len(findings), 2, findings)

    def test_a_differing_protection_option_fails(self) -> None:
        cases = {
            "protection option scenarios is true, but the matrix says false": changed(
                "shInputs", protection={"sheet": True, "objects": True}),
            "protection option formatCells is true, but the matrix says false": changed(
                "shInputs", protection={"sheet": True, "objects": True, "scenarios": True,
                                        "formatCells": False}),
            "sheet shReadme: protection option sheet is false, but the matrix says true": changed(
                "shReadme", protection={"sheet": True}),
        }
        for expected, matrix in cases.items():
            with self.subTest(expected=expected):
                self.assertFinding(run(matrix)[0], expected)
        sheets = workbook_template.default_sheets()
        assert sheets[1].protection is not None
        sheets[1].protection["selectLockedCells"] = "1"
        self.assertFinding(run(template=workbook_template.template(sheets))[0],
                           "selectLockedCells is true, but the matrix says false")

    def test_a_template_needs_a_matrix_and_a_matrix_is_read_alone(self) -> None:
        self.assertFinding(run(None)[0], "has no protection matrix")
        self.assertEqual(run(template=None), (
            [], "2 sheet(s) in the matrix; no src/workbook/Template.xlsx to compare"))

    def test_a_malformed_matrix_fails(self) -> None:
        unknown = copy.deepcopy(MATRIX)
        unknown["sheets"]["shOther"] = {}
        missing = copy.deepcopy(MATRIX)
        del missing["sheets"]["shReadme"]
        cases: dict[str, object] = {
            "not valid UTF-8 JSON": "{",
            "needs schema_version 1": {"sheets": {}},
            "is not a sheet of tools/devbuild/workbook.json": unknown,
            "has no entry for sheet Readme": missing,
            "protection needs true or false": changed("shInputs", protection={"locked": True}),
            "range 'C:C' is not a cell": changed("shInputs", input=["C:C"]),
            "range C5:B4 does not start at its top left cell": changed("shInputs",
                                                                         input=["C5:B4"]),
            "input range C5 overlaps output range A1:C9": changed("shInputs",
                                                                  output=["A1:C9"]),
            "needs an object with protection, input and output only": changed(
                "shInputs", password="secret"),
        }
        for expected, matrix in cases.items():
            with self.subTest(expected=expected):
                self.assertFinding(run(matrix)[0], expected)

    def test_main_reports_the_result(self) -> None:
        with tempfile.TemporaryDirectory() as directory, \
                contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(check_workbook_protection.main(["--root", directory]), 0)
        self.assertIn("workbook-protection: PASS (inactive", output.getvalue())


if __name__ == "__main__":
    unittest.main()
