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
        self.assertIn("1 sheet(s) compared", compared)

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
        for item in LAYOUT.names:
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
        demo = build_workbook.Sheet("Dates Demo", "shDatesDemo", "Output")
        layout = dataclasses.replace(LAYOUT, sheets=(*LAYOUT.sheets, demo))
        texts = {**gen_workbook_template.SHEET_TEXT, "shDatesDemo": "A demo."}
        with unittest.mock.patch.dict(gen_workbook_template.SHEET_TEXT, texts):
            rows = gen_workbook_template.home_rows(layout)
            xml = gen_workbook_template.sheet_xml(rows).decode()
        last = rows[-1].cells[0]
        self.assertEqual((last.text, last.link), ("Dates Demo", "'Dates Demo'!A1"))
        self.assertIn(f'<hyperlink ref="C{rows[-1].number}" location="\'Dates Demo\'!A1" '
                      'display="Dates Demo"/>', xml)
        home = rows[-2].cells[0]
        self.assertEqual((home.text, home.link), ("Home", ""))

    def test_the_steps_say_where_a_function_can_be_typed(self) -> None:
        # Home is protected with every cell locked, and it is the only sheet: the
        # step that uses a function must send the reader to another sheet.
        steps = dict(dict(gen_workbook_template.ABOUT)["How to use"])
        self.assertTrue(gen_workbook_template.PROTECTION["sheet"])
        self.assertIn("Add a worksheet, because Home is protected", steps["2. Use a function"])

    def test_the_sheet_is_protected_without_a_password(self) -> None:
        sheet = parts()["xl/worksheets/sheet1.xml"].decode()
        self.assertIn('<sheetProtection sheet="1" objects="1" scenarios="1"/>', sheet)
        self.assertNotIn("password", sheet.casefold())


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
