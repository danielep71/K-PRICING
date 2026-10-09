"""Synthetic OOXML and temporary Git tests; these never constitute Excel evidence."""

from __future__ import annotations

import contextlib
import io
import json
import subprocess
import tempfile
import unittest
import warnings
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

import kpr
from workbook_snapshot import NS, REL, compare, snapshot


def workbook(path: Path, *, formula: str = "1+1", result: str = "2",
             input_value: str = "7", date1904: str = "0", hidden: str = "visible",
             number_format: str = "0.00", defined_name: str = "Demo!$A$1",
             result_type: str = "n", array_ref: str = "B1:B2", reverse: bool = False,
             created: str = "first", shared: bool = False) -> None:
    sheets = [('Demo', '1'), ('Notes', '2')]
    if reverse:
        sheets.reverse()
    sheets_xml = ''.join(f'<sheet name="{name}" sheetId="{sid}" r:id="rId{sid}" '
                         f'state="{hidden}"/>' for name, sid in sheets)
    label = '<c r="C1" t="inlineStr"><is><t>Label</t></is></c>'
    if shared:
        label = '<c r="C1" t="s"><v>0</v></c>'
    parts = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        'xl/workbook.xml': f'<workbook xmlns="{NS[1:-1]}" xmlns:r="{REL[1:-1]}">'
                           f'<workbookPr date1904="{date1904}"/><sheets>{sheets_xml}</sheets>'
                           f'<definedNames><definedName name="Inputs">{defined_name}</definedName>'
                           '</definedNames></workbook>',
        'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                    f'<Relationship Id="rId1" Type="{REL[1:-1]}/worksheet" Target="worksheets/sheet1.xml"/>'
                                    f'<Relationship Id="rId2" Type="{REL[1:-1]}/worksheet" Target="worksheets/sheet2.xml"/>'
                                    '</Relationships>',
        'xl/worksheets/sheet1.xml': f'<worksheet xmlns="{NS[1:-1]}"><cols><col min="1" max="3" width="20"/></cols>'
                                  f'<sheetData><row r="1"><c r="A1"><v>{input_value}</v></c>'
                                  f'<c r="B1" t="{result_type}"><f t="array" ref="{array_ref}">{formula}</f>'
                                  f'<v>{result}</v></c>{label}</row>'
                                  '<row r="2"><c r="B2"><v>3</v></c></row></sheetData></worksheet>',
        'xl/worksheets/sheet2.xml': f'<worksheet xmlns="{NS[1:-1]}"><sheetData/></worksheet>',
        'xl/styles.xml': f'<styleSheet xmlns="{NS[1:-1]}"><numFmts><numFmt numFmtId="164" formatCode="{number_format}"/></numFmts></styleSheet>',
        'docProps/core.xml': f'<properties><created>{created}</created></properties>',
    }
    if shared:
        parts['xl/sharedStrings.xml'] = f'<sst xmlns="{NS[1:-1]}"><si><t>Label</t></si></sst>'
    with ZipFile(path, 'w') as package:
        for name, contents in parts.items():
            package.writestr(name, contents)


class WorkbookTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.left = Path(self.temp.name) / "left.xlsx"
        self.right = Path(self.temp.name) / "right.xlsx"
        workbook(self.left)

    def test_incidental_properties_and_string_storage(self) -> None:
        workbook(self.right, created="second", shared=True)
        self.assertEqual(compare(self.left, self.right)["status"], "pass")
        self.assertNotEqual(self.left.read_bytes(), self.right.read_bytes())

    def test_changed_formula_points_to_sheet_cell(self) -> None:
        workbook(self.right, formula="1+2")
        report = compare(self.left, self.right)
        self.assertFalse(report["structural_match"])
        self.assertTrue(report["cached_results_match"])
        self.assertIn("/sheets/Demo/cells/B1/formula/text",
                      report["structure_differences"][0]["path"])

    def test_cache_change_is_separate_and_unverified(self) -> None:
        workbook(self.right, result="3")
        report = compare(self.left, self.right)
        self.assertTrue(report["structural_match"])
        self.assertFalse(report["cached_results_match"])
        self.assertEqual(report["calculated_result_comparison"], "NOT_RUN")
        self.assertEqual(report["observations"]["left"]["calculation"], "NOT_OBSERVED")

    def test_fixed_inputs_names_format_date_visibility_order_and_shape(self) -> None:
        mutations = ({"input_value": "8"}, {"defined_name": "Demo!$A$2"},
                     {"number_format": "yyyy-mm-dd"}, {"date1904": "1"},
                     {"hidden": "veryHidden"}, {"reverse": True}, {"array_ref": "B1:C1"})
        for change in mutations:
            with self.subTest(change=change):
                workbook(self.right, **change)
                self.assertFalse(compare(self.left, self.right)["structural_match"])

    def test_native_error_and_array_shape_preserved(self) -> None:
        workbook(self.right, result="#N/A", result_type="e")
        report = snapshot(self.right)
        self.assertEqual(report["cached_results"]["Demo"]["B1"],
                         {"type": "error", "token": "#N/A", "native_code": 2042})
        self.assertEqual(report["structure"]["sheets"]["Demo"]["cells"]["B1"]
                         ["formula"]["attributes"]["ref"], "B1:B2")
        self.assertNotIn("input", report["structure"]["sheets"]["Demo"]["cells"]["B2"])

    def test_unknown_error_keeps_token(self) -> None:
        workbook(self.right, result="#FUTURE!", result_type="e")
        self.assertEqual(snapshot(self.right)["cached_results"]["Demo"]["B1"]["token"], "#FUTURE!")

    def test_malformed_unsafe_and_duplicate_package_refused(self) -> None:
        for contents in ('<!DOCTYPE x [<!ENTITY y "a">]><x>&y;</x>', '<broken>'):
            with self.subTest(contents=contents), ZipFile(self.right, "w") as package:
                package.writestr("xl/workbook.xml", contents)
            with self.assertRaises(ValueError):
                snapshot(self.right)
        with ZipFile(self.right, "w") as package:
            package.writestr("../bad.xml", "<x/>")
        with self.assertRaisesRegex(ValueError, "unsafe ZIP"):
            snapshot(self.right)
        with warnings.catch_warnings(), ZipFile(self.right, "w") as package:
            warnings.simplefilter("ignore", UserWarning)
            package.writestr("duplicate.xml", "<x/>")
            package.writestr("duplicate.xml", "<y/>")
        with self.assertRaisesRegex(ValueError, "duplicate ZIP"):
            snapshot(self.right)

    def test_size_bound_and_non_worksheet_fail_closed(self) -> None:
        with patch("workbook_snapshot.MAX_PACKAGE", 1):
            with self.assertRaisesRegex(ValueError, "package limit"):
                snapshot(self.left)
        with ZipFile(self.left) as package:
            parts = {name: package.read(name) for name in package.namelist()}
        parts["xl/_rels/workbook.xml.rels"] = parts["xl/_rels/workbook.xml.rels"].replace(
            b"/worksheet", b"/chartsheet")
        with ZipFile(self.right, "w") as package:
            for name, raw in parts.items():
                package.writestr(name, raw)
        with self.assertRaisesRegex(ValueError, "unsupported sheet"):
            snapshot(self.right)

    def test_cli_snapshot_and_compare_machine_readable_outcomes(self) -> None:
        output_path = Path(self.temp.name) / "snapshot.json"
        with contextlib.redirect_stdout(io.StringIO()) as output:
            code = kpr.main(["--output", str(output_path), "snapshot", str(self.left)])
        self.assertEqual(code, 0)
        self.assertEqual(output_path.read_text(), output.getvalue())
        self.assertIn("sha256", json.loads(output.getvalue())["tool"])
        workbook(self.right, formula="2+2")
        with contextlib.redirect_stdout(io.StringIO()) as output:
            code = kpr.main(["compare", str(self.left), str(self.right)])
        self.assertEqual(code, 1)
        self.assertFalse(json.loads(output.getvalue())["structural_match"])

    def test_existing_output_preserved_and_errors_structured(self) -> None:
        report_path = Path(self.temp.name) / "report.json"
        report_path.write_text("previous success")
        with contextlib.redirect_stdout(io.StringIO()) as output:
            code = kpr.main(["--output", str(report_path), "snapshot", str(self.left)])
        self.assertEqual(code, 2)
        self.assertEqual(report_path.read_text(), "previous success")
        self.assertEqual(json.loads(output.getvalue())["status"], "error")

    def test_locked_output_failure_preserves_workbooks(self) -> None:
        before = self.left.read_bytes()
        with patch.object(Path, "open", side_effect=PermissionError("synthetic locked output")):
            with contextlib.redirect_stdout(io.StringIO()) as output:
                code = kpr.main(["--output", "locked.json", "snapshot", str(self.left)])
        self.assertEqual(code, 2)
        self.assertIn("locked output", output.getvalue())
        self.assertEqual(before, self.left.read_bytes())


class SourceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name) / "repo"
        self.root.mkdir()
        self.exports = Path(self.temp.name) / "exports"
        self.exports.mkdir()
        self.source = b'Attribute VB_Name = "Example"\nOption Explicit\nPublic Sub Run()\nEnd Sub\n'
        (self.root / "Example.bas").write_bytes(self.source)
        (self.root / ".github").mkdir()
        (self.root / kpr.PROFILE).write_text(json.dumps({
            "repository": "owner/test", "vba": {"components": {"Example.bas": "test"}}}))
        for arguments in (("init", "-q"), ("config", "user.name", "Synthetic"),
                          ("config", "user.email", "synthetic@example.invalid"), ("add", "."),
                          ("commit", "-qm", "Synthetic fixture")):
            kpr.git(self.root, *arguments)
        self.sha = kpr.git(self.root, "rev-parse", "HEAD")
        (self.exports / "Example.bas").write_bytes(self.source.replace(b"\n", b"\r\n"))

    def test_inventory_and_round_trip_reuse_candidate_bytes(self) -> None:
        report = kpr.round_trip(self.root, self.sha, self.exports)
        self.assertEqual(report["status"], "pass")
        self.assertEqual(report["workbook_alignment"], "NOT_OBSERVED")
        self.assertEqual(report["inventory"]["sources"][0]["component"], "Example")

    def test_missing_extra_stale_and_wrong_identity(self) -> None:
        path = self.exports / "Example.bas"
        path.unlink()
        self.assertEqual(kpr.round_trip(self.root, self.sha, self.exports)["findings"][0]["reason"], "missing")
        for raw in (self.source.replace(b"Run", b"OldRun"), self.source.replace(b"Example", b"Wrong")):
            path.write_bytes(raw)
            self.assertEqual(kpr.round_trip(self.root, self.sha, self.exports)["status"], "fail")
        path.write_bytes(self.source)
        (self.exports / "Stale.bas").write_bytes(self.source)
        self.assertEqual(kpr.round_trip(self.root, self.sha, self.exports)["status"], "fail")

    def test_dirty_wrong_sha_and_untracked_sources_refused(self) -> None:
        with self.assertRaisesRegex(ValueError, "HEAD differs"):
            kpr.inventory(self.root, "0" * 40)
        (self.root / "untracked.bas").write_bytes(self.source)
        with self.assertRaisesRegex(ValueError, "dirty"):
            kpr.inventory(self.root, self.sha)
        (self.root / "untracked.bas").unlink()
        (self.root / "Example.bas").write_bytes(self.source + b"' changed\n")
        with self.assertRaisesRegex(ValueError, "dirty"):
            kpr.inventory(self.root, self.sha)

    def test_missing_candidate_component_refused(self) -> None:
        profile = json.loads((self.root / kpr.PROFILE).read_text())
        profile["vba"]["components"]["Missing.bas"] = "test"
        (self.root / kpr.PROFILE).write_text(json.dumps(profile))
        kpr.git(self.root, "add", ".")
        kpr.git(self.root, "commit", "-qm", "Missing synthetic component")
        sha = kpr.git(self.root, "rev-parse", "HEAD")
        with self.assertRaisesRegex(ValueError, "does not contain Missing.bas"):
            kpr.inventory(self.root, sha)

    def test_source_normalization_does_not_hide_edits(self) -> None:
        self.assertNotEqual(kpr.normalize_source(self.source, ".bas"),
                            kpr.normalize_source(self.source.rstrip(), ".bas"))
        with self.assertRaisesRegex(ValueError, "bare CR"):
            kpr.normalize_source(b"line\rnext", ".bas")
        with self.assertRaises(UnicodeError):
            kpr.normalize_source(b"\xff", ".bas")
        self.assertEqual(kpr.normalize_source(b"\xff\r\n", ".frx"), b"\xff\r\n")

    def test_cli_inventory_round_trip_and_doctor(self) -> None:
        for command in (("inventory", "--candidate-sha", self.sha),
                        ("round-trip", "--candidate-sha", self.sha, "--exports", str(self.exports)),
                        ("doctor",)):
            with contextlib.redirect_stdout(io.StringIO()) as output:
                code = kpr.main(["--root", str(self.root), *command])
            self.assertEqual(code, 0)
            self.assertIn("schema_version", json.loads(output.getvalue()))

    def test_read_only_diagnostics_do_not_claim_success(self) -> None:
        prospective = Path(self.temp.name) / "new.json"
        report = kpr.doctor(self.root, prospective)
        self.assertFalse(prospective.exists())
        self.assertTrue(report["source"]["clean"])
        self.assertEqual(report["compile"], "NOT_RUN")
        self.assertEqual(report["host"]["vba_project_access"], "UNKNOWN")
        self.assertEqual(report["output_path"]["lock"], "UNKNOWN")

    def test_authoritative_gate_failure_and_interruption_not_green(self) -> None:
        for exit_code, expected in ((0, "pass"), (1, "fail"), (2, "error"), (-15, "error")):
            with self.subTest(exit_code=exit_code), patch.object(kpr.subprocess, "run", return_value=
                    subprocess.CompletedProcess([], exit_code, "details", "")):
                self.assertEqual(kpr.check(self.root)["status"], expected)


if __name__ == "__main__":
    unittest.main()
