"""Tests for tools/gen_workbook_template.py and the committed template.

The template must equal what the script generates from the repository's
build configuration, carry no macro and no personal metadata, satisfy the
builder's template rules and the protection matrix, and lay out the Home
sheet as docs/methodology/WORKBOOK_CONTRACT.md says. The negative tests work
on copies of the configuration in a temporary directory.
"""

from __future__ import annotations

import contextlib
import dataclasses
import io
import json
import re
import shutil
import sys
import tempfile
import unittest
import unittest.mock
import zipfile
from collections.abc import Iterator
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))

import build_workbook  # noqa: E402
import check_workbook_privacy  # noqa: E402
import check_workbook_protection  # noqa: E402
import demo_dates  # noqa: E402
import gen_workbook_template  # noqa: E402

LAYOUT = build_workbook.load_layout(ROOT)


@contextlib.contextmanager
def copy_root(template: bytes | None = None) -> Iterator[Path]:
    """A directory with the repository's build configuration and, optionally, a template."""
    with tempfile.TemporaryDirectory(prefix="template-") as directory:
        root = Path(directory)
        (root / "tools/devbuild").mkdir(parents=True)
        shutil.copy(ROOT / build_workbook.CONFIG, root / build_workbook.CONFIG)
        if template is not None:
            (root / "src/workbook").mkdir(parents=True)
            (root / build_workbook.TEMPLATE).write_bytes(template)
        yield root


def run_main(*args: str) -> tuple[int, str]:
    with contextlib.redirect_stdout(io.StringIO()) as output:
        status = gen_workbook_template.main(list(args))
    return status, output.getvalue()


def parts() -> dict[str, bytes]:
    with zipfile.ZipFile(ROOT / build_workbook.TEMPLATE) as archive:
        return {info.filename: archive.read(info) for info in archive.infolist()}


class CommittedTemplateTest(unittest.TestCase):
    def test_the_template_matches_its_source(self) -> None:
        status, output = run_main("--check", "--root", str(ROOT))
        self.assertEqual(status, 0, output)
        self.assertIn("workbook-template: PASS", output)

    def test_generation_is_byte_identical(self) -> None:
        self.assertEqual(gen_workbook_template.generate(ROOT),
                         gen_workbook_template.generate(ROOT))

    def test_the_build_accepts_it(self) -> None:
        template = build_workbook.read_template(ROOT / build_workbook.TEMPLATE)
        self.assertEqual(build_workbook.template_problems(LAYOUT, template), [])
        self.assertEqual(build_workbook.main_part_problems(template), [])

    def test_it_carries_no_macro_and_no_personal_metadata(self) -> None:
        names = list(parts())
        self.assertFalse([name for name in names if name.startswith("docProps/")])
        self.assertFalse([name for name in names if "vba" in name.casefold()])
        status = check_workbook_privacy.main([str(ROOT / build_workbook.TEMPLATE)])
        self.assertEqual(status, 0)

    def test_it_meets_its_protection_matrix(self) -> None:
        findings, compared = check_workbook_protection.check(ROOT)
        self.assertEqual(findings, [])
        self.assertIn("2 sheet(s) compared", compared)

    def test_the_matrix_covers_every_row_home_writes(self) -> None:
        matrix = json.loads((ROOT / check_workbook_protection.MATRIX).read_text(
            encoding="utf-8"))
        last = gen_workbook_template.last_row(LAYOUT)
        self.assertEqual(matrix["sheets"]["shHome"]["output"], [f"A1:E{last}"])
        self.assertNotIn("input", matrix["sheets"]["shHome"])
        self.assertEqual(matrix["sheets"]["shHome"]["protection"],
                         gen_workbook_template.PROTECTION)

    def test_the_contract_shows_the_matrix(self) -> None:
        text = (ROOT / "docs/methodology/WORKBOOK_CONTRACT.md").read_text(encoding="utf-8")
        section = text.split("## Protection\n", 1)[1]
        example = json.loads(section.split("```json\n", 1)[1].split("```", 1)[0])
        matrix = json.loads((ROOT / check_workbook_protection.MATRIX).read_text(
            encoding="utf-8"))
        self.assertEqual(example, matrix)


class HomeLayoutTest(unittest.TestCase):
    rows = {row.number: row for row in gen_workbook_template.home_rows(LAYOUT)}

    def text(self, reference: str) -> str:
        column, row = reference[0], int(reference[1:])
        return next(cell.text for cell in self.rows[row].cells if cell.column == column)

    def test_the_title_band_names_the_workbook_and_the_sheet(self) -> None:
        self.assertEqual(self.text("B1"), "K-PRICING")
        self.assertEqual(self.text("B2"), "Home")
        self.assertTrue(self.text("C3"))
        self.assertEqual([self.rows[number].height for number in (1, 2, 3)],
                         list(gen_workbook_template.BAND_HEIGHTS))
        sheet = parts()["xl/worksheets/sheet1.xml"].decode()
        self.assertIn('<pane ySplit="3" topLeftCell="A4"', sheet)

    def test_each_stamp_cell_has_its_label_and_stays_empty(self) -> None:
        for item in (item for item in LAYOUT.names if item.sheet == "shHome"):
            row, column = build_workbook.cell_position(item.cell)
            with self.subTest(name=item.name):
                self.assertEqual(self.text(f"C{row}"), item.label)
                self.assertEqual(self.text(item.cell), "")
                self.assertEqual(column, 4)

    def test_the_sheets_section_lists_every_sheet(self) -> None:
        sheets = [sheet.tab for sheet in gen_workbook_template.template_sheets(LAYOUT)]
        start = next(number for number, row in self.rows.items()
                     if row.cells and row.cells[0].text == "Sheets")
        listed = [self.rows[number].cells[0].text for number in sorted(self.rows)
                  if number > start and self.rows[number].cells]
        self.assertEqual(listed, sheets)

    def test_another_sheet_gets_a_link_from_home(self) -> None:
        notes = build_workbook.Sheet("Notes", "shNotes", "Output")
        layout = dataclasses.replace(LAYOUT, sheets=(*LAYOUT.sheets, notes))
        texts = {**gen_workbook_template.SHEET_TEXT, "shNotes": "Notes."}
        with unittest.mock.patch.dict(gen_workbook_template.SHEET_TEXT, texts):
            rows = gen_workbook_template.home_rows(layout)
            xml = gen_workbook_template.sheet_xml(rows).decode()
        last = rows[-1].cells[0]
        self.assertEqual((last.text, last.link), ("Notes", "'Notes'!A1"))
        self.assertIn(f'<hyperlink ref="C{rows[-1].number}" location="\'Notes\'!A1" '
                      'display="Notes"/>', xml)
        demo = rows[-2].cells[0]
        self.assertEqual((demo.text, demo.link), ("Dates Demo", "'Dates Demo'!A1"))
        home = rows[-3].cells[0]
        self.assertEqual((home.text, home.link), ("Home", ""))

    def test_the_sheet_is_not_protected(self) -> None:
        # No sheet is protected unless a decision of the workbook contract says
        # so, which also lets a function be typed on Home, the only sheet.
        self.assertEqual(gen_workbook_template.PROTECTION, {})
        for name in ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"):
            self.assertNotIn("sheetProtection", parts()[name].decode())

    def test_a_protection_decision_reaches_the_sheet(self) -> None:
        options = {"sheet": True, "objects": True, "scenarios": True}
        with unittest.mock.patch.dict(gen_workbook_template.PROTECTION, options):
            sheet = gen_workbook_template.sheet_xml(
                gen_workbook_template.home_rows(LAYOUT)).decode()
        self.assertIn('<sheetProtection sheet="1" objects="1" scenarios="1"/>', sheet)
        self.assertNotIn("password", sheet.casefold())


class DemoSheetTest(unittest.TestCase):
    rows = gen_workbook_template.demo_rows(LAYOUT)
    cells = {f"{cell.column}{row.number}": cell for row in rows for cell in row.cells}

    def formulas(self) -> list[str]:
        return [cell.formula for cell in self.cells.values() if cell.formula]

    def test_every_function_is_called(self) -> None:
        source = (ROOT / "docs/PUBLIC_API.txt").read_text(encoding="utf-8")
        functions = set(re.findall(r"^KPR_Dates\tFunction\t(\w+)$", source, re.M))
        self.assertEqual(len(functions), 22)
        called = {name for formula in self.formulas()
                  for name in re.findall(r"\b(KPR_Dates_\w+)\(", formula)}
        self.assertEqual(called, functions)

    def test_each_row_shows_the_formula_it_calculates(self) -> None:
        shown = [(number, cell.text) for number, cell in
                 ((int(ref[1:]), cell) for ref, cell in self.cells.items() if ref[0] == "C")
                 if cell.style == gen_workbook_template.DEMO_FORMULA_TEXT]
        self.assertEqual(len(shown), sum(len(item.rows) for item in demo_dates.SECTIONS))
        for number, text in shown:
            with self.subTest(row=number):
                self.assertEqual(self.cells[f"D{number}"].formula, text)
                self.assertNotIn("KPR_", self.cells[f"E{number}"].formula)
                self.assertIn(f"D{number}", self.cells[f"F{number}"].formula)

    def test_every_input_is_named_unlocked_and_used(self) -> None:
        names = {item.name: item.cell for item in LAYOUT.names if item.sheet == "shDatesDemo"}
        text = " ".join(self.formulas())
        for item in demo_dates.INPUTS:
            with self.subTest(name=item.name):
                cell = self.cells[names[item.name]]
                self.assertIn(cell.style, (gen_workbook_template.DEMO_INPUT,
                                           gen_workbook_template.DEMO_INPUT_DATE))
                self.assertRegex(text, rf"\b{item.name}\b")
        self.assertNotRegex(text, r"\{IN\}|\{SRC\}")

    def test_the_checks_count_every_status(self) -> None:
        statuses = gen_workbook_template.demo_statuses(LAYOUT)
        self.assertEqual(len(statuses), 57)
        self.assertEqual(self.cells["J9"].formula, f"={len(statuses)}")
        for cell in statuses:
            with self.subTest(cell=cell):
                self.assertIn(f'({cell}="OK")', self.cells["J6"].formula)
        self.assertEqual(self.cells["I5"].text, "Check")
        source = (ROOT / "tests/modules/TEST_Ui.bas").read_text(encoding="ascii")
        self.assertRegex(source, rf"DEMO_CHECKS\s+As Long = {len(statuses)}\b")

    def test_the_matrix_records_the_inputs_and_the_rest(self) -> None:
        matrix = json.loads((ROOT / check_workbook_protection.MATRIX).read_text(
            encoding="utf-8"))
        entry = matrix["sheets"]["shDatesDemo"]
        expected = gen_workbook_template.demo_protection(LAYOUT)
        self.assertEqual((entry["input"], entry["output"]), (expected["input"],
                                                             expected["output"]))
        self.assertEqual(entry["protection"], {})

    def test_every_formula_is_a_dynamic_array_formula_calculated_on_open(self) -> None:
        sheet = parts()["xl/worksheets/sheet2.xml"].decode()
        self.assertEqual(sheet.count("<f "), len(self.formulas()))
        self.assertEqual(sheet.count('cm="1"><f t="array"'), len(self.formulas()))
        self.assertIn('<f t="array" ref="I35:I64">_xlfn.SEQUENCE(30,1,DateIn,25)</f>', sheet)
        self.assertIn("KPR_Dates_DaysInMonth(_xlfn.ANCHORARRAY(I35))", sheet)
        self.assertIn('fullCalcOnLoad="1"', parts()["xl/workbook.xml"].decode())
        self.assertIn("XLDAPR", parts()["xl/metadata.xml"].decode())
        self.assertIn('<hyperlink ref="D1" location="Home!A1" display="Home"/>', sheet)

    def test_file_formulas(self) -> None:
        self.assertEqual(gen_workbook_template.file_formula("=ROWS(SEQUENCE(2))+ROWS(I36#)"),
                         "ROWS(_xlfn.SEQUENCE(2))+ROWS(_xlfn.ANCHORARRAY(I36))")

    def test_a_moved_input_or_anchor_is_refused(self) -> None:
        for name, cell in (("nDays", "D20"), ("DatesDemo_Checks", "I4")):
            names = tuple(dataclasses.replace(item, cell=cell) if item.name == name else item
                          for item in LAYOUT.names)
            with self.subTest(name=name), self.assertRaisesRegex(build_workbook.BuildError,
                                                                 name):
                gen_workbook_template.demo_rows(dataclasses.replace(LAYOUT, names=names))


class RefusalTest(unittest.TestCase):
    def test_a_changed_template_fails(self) -> None:
        good = gen_workbook_template.generate(ROOT)
        with copy_root(good + b"\0") as root:
            status, output = run_main("--root", str(root))
        self.assertEqual(status, 1)
        self.assertIn("differs from its source", output)

    def test_a_missing_template_fails(self) -> None:
        with copy_root() as root:
            status, output = run_main("--check", "--root", str(root))
        self.assertEqual(status, 1)
        self.assertIn("is missing", output)

    def test_write_then_check_passes(self) -> None:
        with copy_root() as root:
            self.assertEqual(run_main("--write", "--root", str(root))[0], 0)
            self.assertEqual(run_main("--root", str(root))[0], 0)
            self.assertEqual((root / build_workbook.TEMPLATE).read_bytes(),
                             (ROOT / build_workbook.TEMPLATE).read_bytes())

    def test_a_sheet_without_a_description_is_refused(self) -> None:
        demo = build_workbook.Sheet("Demo", "shDemo", "Output")
        layout = dataclasses.replace(LAYOUT, sheets=(*LAYOUT.sheets, demo))
        problems = gen_workbook_template.layout_problems(layout)
        self.assertTrue(any("sheet shDemo has no text in SHEET_TEXT" in problem
                            for problem in problems), problems)
        with self.assertRaises(build_workbook.BuildError):
            gen_workbook_template.parts(layout)

    def test_a_first_sheet_other_than_home_is_refused(self) -> None:
        demo = build_workbook.Sheet("Demo", "shDemo", "Output")
        layout = dataclasses.replace(LAYOUT, sheets=(demo, *LAYOUT.sheets))
        self.assertTrue(any("the first sheet must be shHome" in problem
                            for problem in gen_workbook_template.layout_problems(layout)))

    def test_a_stamp_name_off_its_row_is_refused(self) -> None:
        names = tuple(dataclasses.replace(item, cell="D9") if item.name == "build_time"
                      else item for item in LAYOUT.names)
        layout = dataclasses.replace(LAYOUT, names=names)
        with self.assertRaisesRegex(build_workbook.BuildError, "consecutive cells of column D"):
            gen_workbook_template.home_rows(layout)

    def test_a_configuration_the_script_cannot_lay_out_fails_the_check(self) -> None:
        with copy_root(gen_workbook_template.generate(ROOT)) as root:
            path = root / build_workbook.CONFIG
            data = json.loads(path.read_text(encoding="utf-8"))
            data["sheets"].append({"tab": "Demo", "code_name": "shDemo", "role": "Output"})
            path.write_text(json.dumps(data), encoding="utf-8")
            status, output = run_main("--root", str(root))
        self.assertEqual(status, 1)
        self.assertIn("workbook-template: FAIL", output)


if __name__ == "__main__":
    unittest.main()
