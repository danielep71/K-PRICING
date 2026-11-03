"""Tests for the RibbonX package injector and callback check (tools/inject_ribbon.py).

Every package is a synthetic .xlsm/.xlam built in a temporary directory; no
tracked file is read or changed.
"""

from __future__ import annotations

import contextlib
import io
import os
import random
import shutil
import subprocess
import tempfile
import unittest
import unittest.mock
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import inject_ribbon

CT_NS = inject_ribbon.CONTENT_TYPES_NS
REL_NS = inject_ribbon.RELATIONSHIPS_NS
OFFICE_DOCUMENT = inject_ribbon.OFFICE_DOCUMENT_TYPES[0]
CORE_PROPERTIES = ("http://schemas.openxmlformats.org/package/2006/relationships/metadata/"
                   "core-properties")
DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
XML_DEFAULT = '<Default Extension="xml" ContentType="application/xml"/>'
VBA_PROJECT = random.Random(47).randbytes(4096)
WORKBOOK = (DECLARATION + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml'
            '/2006/main"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" xmlns:r="http://'
            'schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets></workbook>')
WORKBOOK_RELS = (DECLARATION + f'<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" '
                 'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                 'Target="vbaProject.bin"/></Relationships>')
RIBBON = f'Type="{inject_ribbon.RIBBON_REL_TYPE}" Target="customUI/customUI14.xml"'

CUSTOMUI = """<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"
          onLoad="KPR_Ribbon_OnLoad">
  <ribbon>
    <tabs>
      <tab id="KPR_Tab" label="K-PRICING" getVisible="KPR_Ribbon_GetVisible">
        <group id="KPR_Group_Dates" label="Dates">
          <button id="KPR_Btn_Register" label="Register" imageMso="FunctionWizard"
                  onAction="KPR_Ribbon_OnAction" getEnabled="KPR_Ribbon_GetEnabled"/>
          <toggleButton id="KPR_Tgl_Demo" getLabel="KPR_Ribbon_GetLabel"
                        getPressed="KPR_Ribbon_GetPressed" onAction="KPR_Ribbon_OnToggle"/>
          <dropDown id="KPR_Drp_Suite" getItemCount="KPR_Ribbon_GetItemCount"
                    getItemLabel="KPR_Ribbon_GetItemLabel" getItemID="KPR_Ribbon_GetItemID"
                    getSelectedItemIndex="KPR_Ribbon_GetSelectedItemIndex"
                    onAction="KPR_Ribbon_OnSelect"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"""

# One block per procedure, so a negative case can replace exactly one of them.
CALLBACKS = {
    "OnLoad": "Public Sub KPR_Ribbon_OnLoad(ByVal ribbon As IRibbonUI)\n"
              "    Set mRibbon = ribbon\nEnd Sub",
    "OnAction": "Public Sub KPR_Ribbon_OnAction(control As IRibbonControl)\n"
                '    MsgBox "Public Sub KPR_Ribbon_InString(control As IRibbonControl)"\n'
                "    HandleClick control ' Public Sub KPR_Ribbon_InComment(c As IRibbonControl)\n"
                "End Sub",
    "OnToggle": "Public Sub KPR_Ribbon_OnToggle(control As Office.IRibbonControl, _\n"
                "                               pressed As Boolean)\nEnd Sub",
    "OnSelect": "Public Sub KPR_Ribbon_OnSelect(control As IRibbonControl, _\n"
                "    selectedId As String, _\n    selectedIndex As Integer)\nEnd Sub",
    "GetVisible": "Public Sub KPR_Ribbon_GetVisible(control As IRibbonControl, ByRef returnedVal)"
                  "\n    returnedVal = True\nEnd Sub",
    "GetEnabled": "Public Sub kpr_ribbon_getenabled(ByVal control As IRibbonControl, "
                  "returnedVal As Variant)\nEnd Sub",
    "GetLabel": "Sub KPR_Ribbon_GetLabel(control As IRibbonControl, ByRef label As Variant)\n"
                "End Sub",
    "GetPressed": "Public Sub KPR_Ribbon_GetPressed(control As IRibbonControl, ByRef pressed)\n"
                  "End Sub",
    "GetItemCount": "Public Sub KPR_Ribbon_GetItemCount(control As IRibbonControl, count)\n"
                    "End Sub",
    "GetItemLabel": "Public Sub KPR_Ribbon_GetItemLabel(control As IRibbonControl, "
                    "index As Integer, ByRef returnedVal)\nEnd Sub",
    "GetItemID": "Public Sub KPR_Ribbon_GetItemID(control As IRibbonControl, ByVal index As Long, "
                 "ByRef returnedVal As Variant)\nEnd Sub",
    "GetSelectedItemIndex": "Public Sub KPR_Ribbon_GetSelectedItemIndex(control As IRibbonControl,"
                            " ByRef returnedVal)\nEnd Sub",
    "Helper": "Private Sub HandleClick(ByVal control As IRibbonControl)\nEnd Sub\n\n"
              "Public Function KPR_Ribbon_Version() As Long\nEnd Function",
}
MODULE_HEADER = ('Attribute VB_Name = "KPR_UI_Ribbon"\nOption Explicit\n'
                 "' Ribbon callbacks – “Office 2010” signatures\n"
                 "Private mRibbon As IRibbonUI\n")


def module(**changes: str) -> bytes:
    """The callback module as Excel exports it (cp1252, CRLF), with blocks replaced."""
    blocks = {**CALLBACKS, **changes}
    text = MODULE_HEADER + "\n" + "\n\n".join(b for b in blocks.values() if b) + "\n"
    return text.replace("\n", "\r\n").encode("cp1252")


def content_types(extension: str = ".xlsm", xml_default: bool = True) -> str:
    return (DECLARATION + f'<Types xmlns="{CT_NS}">'
            '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
            '<Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            + (XML_DEFAULT if xml_default else "")
            + '<Override PartName="/xl/workbook.xml" '
            f'ContentType="{inject_ribbon.WORKBOOK_CONTENT_TYPES[extension]}"/></Types>')


def package_rels(extra: str = "") -> str:
    return (DECLARATION + f'<Relationships xmlns="{REL_NS}">'
            f'<Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="xl/workbook.xml"/>'
            f'<Relationship Id="rId2" Type="{CORE_PROPERTIES}" Target="docProps/core.xml"/>'
            f"{extra}</Relationships>")


class Unseekable(io.BytesIO):
    """Makes zipfile stream its entries with data descriptors, as some writers do."""

    def seek(self, *args: object) -> int:
        raise OSError("unseekable")


def build_package(path: Path, parts: dict[str, str | bytes | None] | None = None,
                  descriptors: bool = False) -> bytes:
    """Write a minimal macro-enabled package; a None part is left out."""
    extension = path.suffix if path.suffix in inject_ribbon.WORKBOOK_CONTENT_TYPES else ".xlsm"
    entries: dict[str, str | bytes | None] = {
        "[Content_Types].xml": content_types(extension),
        "_rels/.rels": package_rels(),
        "xl/workbook.xml": WORKBOOK,
        "xl/_rels/workbook.xml.rels": WORKBOOK_RELS,
        "xl/vbaProject.bin": VBA_PROJECT,
        "docProps/core.xml": DECLARATION + "<cp:coreProperties xmlns:cp=\"urn:x\"/>",
        **(parts or {}),
    }
    stream = Unseekable() if descriptors else io.BytesIO()
    with zipfile.ZipFile(stream, "w") as archive:
        for name, payload in entries.items():
            if payload is None:
                continue
            method = zipfile.ZIP_STORED if name.endswith(".bin") else zipfile.ZIP_DEFLATED
            info = zipfile.ZipInfo(name, (2026, 11, 3, 10, 20, 30))
            info.compress_type = method
            archive.writestr(info, payload)
    data = stream.getvalue()
    path.write_bytes(data)
    return data


def ribbon_relationships(data: bytes) -> list[ET.Element]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        root = ET.fromstring(archive.read("_rels/.rels"))  # noqa: S314 -- synthetic fixture.
    return [rel for rel in root if rel.get("Type") == inject_ribbon.RIBBON_REL_TYPE]


def xml_defaults(data: bytes) -> list[str | None]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        root = ET.fromstring(archive.read("[Content_Types].xml"))  # noqa: S314 -- synthetic fixture.
    return [d.get("ContentType") for d in root.iter(inject_ribbon.qualified(CT_NS, "Default"))
            if d.get("Extension") == "xml"]


def run(*argv: str) -> tuple[int, str, str]:
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        code = inject_ribbon.main(list(argv))
    return code, out.getvalue(), err.getvalue()


class Fixture(unittest.TestCase):
    def setUp(self) -> None:
        self.work = Path(tempfile.mkdtemp(prefix="kpr-ribbon-"))
        self.addCleanup(shutil.rmtree, self.work)
        self.box = self.work / "box"
        self.box.mkdir()
        self.customui = self.work / "customUI14.xml"
        self.customui.write_text(CUSTOMUI, encoding="utf-8")

    def inject(self, target: Path, *extra: str) -> tuple[int, str, str]:
        return run("--target", str(target), "--customui", str(self.customui), *extra)

    def assert_untouched(self, target: Path, original: bytes, code: int, err: str) -> None:
        self.assertEqual(code, 1, err)
        self.assertTrue(err.startswith("ERROR: "), err)
        self.assertEqual(target.read_bytes(), original)
        self.assertEqual(sorted(os.listdir(target.parent)), [target.name])


class InjectionTests(Fixture):
    def test_injects_xlsm_and_xlam(self) -> None:
        for extension in (".xlsm", ".xlam"):
            with self.subTest(extension=extension):
                target = self.box / f"book{extension}"
                original = build_package(target)
                code, out, err = self.inject(target)
                self.assertEqual(code, 0, err)
                self.assertIn("customUI/customUI14.xml: added", out)
                self.assertIn("added relationship rIdKprRibbon", out)
                data = target.read_bytes()
                with zipfile.ZipFile(io.BytesIO(data)) as archive:
                    self.assertIsNone(archive.testzip())
                    self.assertEqual(archive.read("customUI/customUI14.xml"),
                                     self.customui.read_bytes())
                rels = ribbon_relationships(data)
                self.assertEqual([(r.get("Id"), r.get("Target")) for r in rels],
                                 [("rIdKprRibbon", "customUI/customUI14.xml")])
                self.assertEqual(xml_defaults(data), ["application/xml"])
                backup = target.with_name(target.name + ".bak")
                self.assertEqual(backup.read_bytes(), original)
                self.assertEqual(sorted(os.listdir(self.box)), [target.name, backup.name])
                os.remove(target)
                os.remove(backup)

    def test_repeat_injection_is_byte_identical(self) -> None:
        first, second = self.box / "a.xlsm", self.work / "b.xlsm"
        build_package(first)
        shutil.copyfile(first, second)
        self.assertEqual(self.inject(first)[0], 0)
        self.assertEqual(self.inject(second)[0], 0)
        injected = first.read_bytes()
        self.assertEqual(injected, second.read_bytes())
        code, out, _ = self.inject(first)
        self.assertEqual(code, 0)
        self.assertIn("nothing written", out)
        self.assertEqual(first.read_bytes(), injected)
        self.assertEqual(len(ribbon_relationships(injected)), 1)

    def test_new_xml_replaces_the_part_and_reuses_the_relationship(self) -> None:
        target = self.box / "book.xlsm"
        build_package(target)
        self.assertEqual(self.inject(target)[0], 0)
        injected = target.read_bytes()
        self.customui.write_text(CUSTOMUI.replace('label="Dates"', 'label="Days"'),
                                 encoding="utf-8")
        # The default backup now holds the pristine package, so it is not overwritten.
        self.assertEqual(self.inject(target)[0], 1)
        self.assertEqual(target.read_bytes(), injected)
        second_backup = self.work / "book.second.bak"
        code, out, err = self.inject(target, "--backup", str(second_backup))
        self.assertEqual(code, 0, err)
        self.assertIn("customUI/customUI14.xml: replaced", out)
        self.assertIn("kept relationship rIdKprRibbon", out)
        self.assertEqual(second_backup.read_bytes(), injected)
        self.assertEqual(len(ribbon_relationships(target.read_bytes())), 1)

    def test_duplicate_relationships_collapse_to_the_first_id(self) -> None:
        target = self.box / "book.xlsm"
        duplicates = (f'<Relationship Id="rId7" {RIBBON}/>'
                      f'<Relationship Id="rId9" Type="{inject_ribbon.RIBBON_REL_TYPE}" '
                      'Target="/customUI/customUI14.xml"/>')
        build_package(target, {"_rels/.rels": package_rels(duplicates)})
        code, out, err = self.inject(target)
        self.assertEqual(code, 0, err)
        self.assertIn("removed 1 duplicate", out)
        with zipfile.ZipFile(target) as archive:
            root = ET.fromstring(archive.read("_rels/.rels"))  # noqa: S314 -- synthetic fixture.
        self.assertEqual([(r.get("Id"), r.get("Type")) for r in root],
                         [("rId1", OFFICE_DOCUMENT), ("rId2", CORE_PROPERTIES),
                          ("rId7", inject_ribbon.RIBBON_REL_TYPE)])

    def test_new_relationship_id_avoids_existing_ids(self) -> None:
        target = self.box / "book.xlsm"
        taken = f'<Relationship Id="rIdKprRibbon" Type="{CORE_PROPERTIES}" Target="x.xml"/>'
        build_package(target, {"_rels/.rels": package_rels(taken)})
        self.assertEqual(self.inject(target)[0], 0)
        rels = ribbon_relationships(target.read_bytes())
        self.assertEqual([r.get("Id") for r in rels], ["rIdKprRibbon2"])

    def test_xml_default_is_added_once_and_never_duplicated(self) -> None:
        for present in (True, False):
            with self.subTest(present=present):
                target = self.box / "book.xlsm"
                original = build_package(
                    target, {"[Content_Types].xml": content_types(xml_default=present)})
                code, out, _ = self.inject(target, "--backup", str(self.work / f"{present}.bak"))
                self.assertEqual(code, 0)
                self.assertEqual(xml_defaults(target.read_bytes()), ["application/xml"])
                self.assertEqual("added Default xml" in out, not present)
                if present:
                    before = inject_ribbon.load_package(original).find("[Content_Types].xml")
                    after = inject_ribbon.load_package(target.read_bytes()).find(
                        "[Content_Types].xml")
                    assert before is not None and after is not None
                    self.assertEqual(before.raw, after.raw)
                os.remove(target)

    def test_other_entries_are_preserved_byte_for_byte(self) -> None:
        for descriptors in (False, True):
            with self.subTest(descriptors=descriptors):
                target = self.box / "book.xlsm"
                original = build_package(target, descriptors=descriptors)
                code, _, err = self.inject(target, "--backup", str(self.work / "x.bak"))
                self.assertEqual(code, 0, err)
                before = inject_ribbon.load_package(original).members
                after = inject_ribbon.load_package(target.read_bytes()).members
                written = {"_rels/.rels", "customUI/customUI14.xml"}
                kept = [m for m in before if m.name not in written]
                self.assertEqual([m.name for m in after if m.name not in written],
                                 [m.name for m in kept])
                for old, new in zip(kept, (m for m in after if m.name not in written)):
                    self.assertEqual(old.raw, new.raw, old.name)
                    self.assertEqual(old.info.compress_type, new.info.compress_type)
                    self.assertEqual(old.info.date_time, new.info.date_time)
                    self.assertEqual(old.info.flag_bits, new.info.flag_bits)
                self.assertEqual(after[-1].name, "customUI/customUI14.xml")
                self.assertEqual(after[-1].info.date_time, inject_ribbon.FIXED_DATE_TIME)
                vba = next(m for m in after if m.name == "xl/vbaProject.bin")
                self.assertEqual(vba.info.compress_type, zipfile.ZIP_STORED)
                with zipfile.ZipFile(target) as archive:
                    self.assertEqual(archive.read("xl/vbaProject.bin"), VBA_PROJECT)
                os.remove(target)
                os.remove(self.work / "x.bak")

    def test_backup_restores_the_original(self) -> None:
        target = self.box / "book.xlam"
        original = build_package(target)
        backup = self.work / "kept" / "book.original.xlam"
        backup.parent.mkdir()
        self.assertEqual(self.inject(target, "--backup", str(backup))[0], 0)
        self.assertNotEqual(target.read_bytes(), original)
        shutil.copyfile(backup, target)
        self.assertEqual(target.read_bytes(), original)
        # An existing backup that already equals the target is reused, not rewritten.
        code, out, _ = self.inject(target, "--backup", str(backup))
        self.assertEqual(code, 0)
        self.assertIn("already present, identical", out)

    def test_backup_must_differ_from_target(self) -> None:
        target = self.box / "book.xlsm"
        original = build_package(target)
        self.assert_untouched(target, original, *self.inject(target, "--backup", str(target))[::2])


class RefusalTests(Fixture):
    def assert_refused(self, target: Path, original: bytes, *extra: str) -> str:
        code, _, err = self.inject(target, *extra)
        self.assert_untouched(target, original, code, err)
        return err

    def test_not_a_zip(self) -> None:
        target = self.box / "book.xlsm"
        target.write_bytes(b"not a zip package")
        self.assertIn("not a ZIP", self.assert_refused(target, b"not a zip package"))

    def test_missing_package_parts(self) -> None:
        for part in ("_rels/.rels", "[Content_Types].xml", "xl/workbook.xml"):
            with self.subTest(part=part):
                target = self.box / "book.xlsm"
                original = build_package(target, {part: None})
                self.assertIn(f"no {part} part", self.assert_refused(target, original))
                os.remove(target)

    def test_unsafe_or_malformed_part_xml(self) -> None:
        doctype = DECLARATION + '<!DOCTYPE x [<!ENTITY a "b">]><workbook/>'
        for part, payload, message in (
                ("xl/workbook.xml", doctype, "DTD and entity"),
                ("_rels/.rels", "<Relationships", "invalid XML"),
                ("[Content_Types].xml", "<Types/>", "no OPC Types root")):
            with self.subTest(part=part):
                target = self.box / "book.xlsm"
                original = build_package(target, {part: payload})
                self.assertIn(message, self.assert_refused(target, original))
                os.remove(target)

    def test_workbook_type_must_match_the_extension(self) -> None:
        target = self.box / "book.xlam"
        original = build_package(target, {"[Content_Types].xml": content_types(".xlsm")})
        self.assertIn("addin.macroEnabled", self.assert_refused(target, original))

    def test_invalid_customui(self) -> None:
        cases = {
            "root": CUSTOMUI.replace("<customUI", "<ribbonUI").replace("</customUI", "</ribbonUI"),
            "namespace": CUSTOMUI.replace("2009/07/customui", "2006/01/customui"),
            "malformed": CUSTOMUI.replace("</customUI>", ""),
            "doctype": '<!DOCTYPE customUI SYSTEM "x.dtd">' + CUSTOMUI,
        }
        for name, text in cases.items():
            with self.subTest(case=name):
                target = self.box / "book.xlsm"
                original = build_package(target)
                self.customui.write_text(text, encoding="utf-8")
                self.assert_refused(target, original)
                os.remove(target)

    def test_wrong_extension(self) -> None:
        for name in ("book.xlsx", "book.zip", "book.xlsm.txt"):
            with self.subTest(name=name):
                target = self.box / name
                original = build_package(target)
                self.assertIn(".xlsm and .xlam", self.assert_refused(target, original))
                os.remove(target)

    def test_existing_different_backup_is_never_overwritten(self) -> None:
        target = self.box / "book.xlsm"
        original = build_package(target)
        backup = self.work / "old.bak"
        backup.write_bytes(b"someone's earlier backup")
        self.assertIn("already exists", self.assert_refused(target, original, "--backup",
                                                            str(backup)))
        self.assertEqual(backup.read_bytes(), b"someone's earlier backup")

    @unittest.skipUnless(shutil.which("git"), "git is not installed")
    def test_tracked_target_is_refused(self) -> None:
        repository = self.box
        subprocess.run(["git", "init", "-q", str(repository)], check=True)
        target = repository / "Book.xlsm"
        original = build_package(target)
        subprocess.run(["git", "-C", str(repository), "add", "Book.xlsm"], check=True)
        code, _, err = self.inject(target)
        self.assertEqual(code, 1)
        self.assertIn("tracked by git", err)
        self.assertEqual(target.read_bytes(), original)
        self.assertEqual(sorted(os.listdir(repository)), [".git", "Book.xlsm"])
        # The same package next to it, but untracked, is accepted.
        untracked = repository / "Copy.xlsm"
        shutil.copyfile(target, untracked)
        self.assertEqual(self.inject(untracked)[0], 0)

    def test_failure_at_replacement_keeps_target_and_removes_temp(self) -> None:
        target = self.box / "book.xlsm"
        original = build_package(target)
        with unittest.mock.patch("inject_ribbon.os.replace",
                                 side_effect=PermissionError("locked by Excel")):
            code, _, err = self.inject(target)
        self.assertEqual(code, 1)
        self.assertIn("locked by Excel", err)
        self.assertEqual(target.read_bytes(), original)
        backup = target.with_name("book.xlsm.bak")
        self.assertEqual(sorted(os.listdir(self.box)), ["book.xlsm", backup.name])
        self.assertEqual(backup.read_bytes(), original)  # identical, so a rerun reuses it
        self.assertEqual(self.inject(target)[0], 0)

    def test_failed_revalidation_keeps_target_and_removes_temp(self) -> None:
        target = self.box / "book.xlsm"
        original = build_package(target)
        failure = inject_ribbon.RibbonError("re-validation failed")
        with unittest.mock.patch.object(inject_ribbon, "verify_rebuilt", side_effect=failure):
            self.assert_untouched(target, original, *self.inject(target)[::2])

    def test_usage_errors_exit_2(self) -> None:
        for argv in ([], ["--check", "--backup", "x.bak"], ["--target", "x.xlsm", "--callbacks",
                                                           "x.bas"],
                     ["--check", "--target", "x.xlsm"]):
            with self.subTest(argv=argv), contextlib.redirect_stderr(io.StringIO()), \
                    self.assertRaises(SystemExit) as error:
                inject_ribbon.main(argv)
            self.assertEqual(error.exception.code, 2)


class CallbackTests(unittest.TestCase):
    def setUp(self) -> None:
        self.work = Path(tempfile.mkdtemp(prefix="kpr-ribbon-check-"))
        self.addCleanup(shutil.rmtree, self.work)

    def rules(self, xml: str = CUSTOMUI, **blocks: str) -> list[str]:
        customui, callbacks = self.work / "customUI14.xml", self.work / "KPR_UI_Ribbon.bas"
        customui.write_text(xml, encoding="utf-8")
        callbacks.write_bytes(module(**blocks))
        return [f.rule for f in inject_ribbon.check_ribbon(customui, callbacks)]

    def test_consistent_ribbon_passes(self) -> None:
        self.assertEqual(self.rules(), [])
        code, out, err = run("--check", "--customui", str(self.work / "customUI14.xml"),
                             "--callbacks", str(self.work / "KPR_UI_Ribbon.bas"))
        self.assertEqual((code, out, err), (0, "ribbon-check: PASS\n", ""))

    def test_missing_callback(self) -> None:
        self.assertEqual(self.rules(GetPressed=""), ["missing-callback"])

    def test_wrong_parameter_count(self) -> None:
        self.assertEqual(self.rules(OnToggle="Public Sub KPR_Ribbon_OnToggle(control As "
                                    "IRibbonControl)\nEnd Sub"), ["callback-signature"])
        self.assertEqual(self.rules(OnAction="Public Sub KPR_Ribbon_OnAction()\nEnd Sub"),
                         ["callback-signature"])

    def test_wrong_parameter_type(self) -> None:
        for block in ("Public Sub KPR_Ribbon_OnLoad(ribbon As IRibbonControl)\nEnd Sub",
                      "Public Sub KPR_Ribbon_OnLoad(ribbon As Object)\nEnd Sub",
                      "Public Sub KPR_Ribbon_OnLoad(ribbon)\nEnd Sub"):
            with self.subTest(block=block):
                self.assertEqual(self.rules(OnLoad=block), ["callback-signature"])
        self.assertEqual(self.rules(GetItemLabel="Public Sub KPR_Ribbon_GetItemLabel(control As "
                                    "IRibbonControl, index As String, returnedVal)\nEnd Sub"),
                         ["callback-signature"])
        self.assertEqual(self.rules(GetVisible="Public Sub KPR_Ribbon_GetVisible(control As "
                                    "IRibbonControl, ByRef returnedVal As Boolean)\nEnd Sub"),
                         ["callback-signature"])

    def test_returned_value_passed_byval(self) -> None:
        block = ("Public Sub KPR_Ribbon_GetVisible(control As IRibbonControl, _\n"
                 "    ByVal returnedVal As Variant)\nEnd Sub")
        self.assertEqual(self.rules(GetVisible=block), ["callback-signature"])

    def test_private_or_function_callback(self) -> None:
        self.assertEqual(self.rules(OnAction="Private Sub KPR_Ribbon_OnAction(control As "
                                    "IRibbonControl)\nEnd Sub"), ["callback-signature"])
        self.assertEqual(self.rules(OnAction="Public Function KPR_Ribbon_OnAction(control As "
                                    "IRibbonControl) As Long\nEnd Function"),
                         ["callback-signature"])

    def test_dead_callback(self) -> None:
        unused = "Public Sub KPR_Ribbon_Unused(control As IRibbonControl)\nEnd Sub"
        self.assertEqual(self.rules(Helper=unused), ["dead-callback"])

    def test_duplicate_and_unstable_ids(self) -> None:
        duplicate = CUSTOMUI.replace('id="KPR_Tgl_Demo"', 'id="KPR_Btn_Register"')
        self.assertEqual(self.rules(duplicate), ["duplicate-id"])
        for bad in ("btnRegister", "KPR_", "KPR_Btn-Register", "kpr_Btn"):
            with self.subTest(id=bad):
                self.assertEqual(self.rules(CUSTOMUI.replace("KPR_Btn_Register", bad)),
                                 ["control-id"])

    def test_custom_image(self) -> None:
        xml = CUSTOMUI.replace('imageMso="FunctionWizard"', 'image="kpr_logo"')
        self.assertEqual(self.rules(xml), ["custom-image"])

    def test_qualified_callback_name(self) -> None:
        xml = CUSTOMUI.replace('"KPR_Ribbon_OnAction"', '"KPR_UI_Ribbon.KPR_Ribbon_OnAction"')
        self.assertEqual(self.rules(xml), ["callback-name", "dead-callback"])

    def test_duplicate_procedure(self) -> None:
        twice = CALLBACKS["OnAction"] + "\n\n" + CALLBACKS["OnAction"].upper()
        self.assertEqual(self.rules(OnAction=twice), ["duplicate-procedure"])

    def test_findings_reach_stderr_with_exit_1(self) -> None:
        self.rules(GetPressed="")
        code, out, err = run("--check", "--customui", str(self.work / "customUI14.xml"),
                             "--callbacks", str(self.work / "KPR_UI_Ribbon.bas"))
        self.assertEqual((code, out), (1, ""))
        self.assertIn("FAIL: [missing-callback] KPR_Ribbon_GetPressed", err)
        self.assertIn("ribbon-check: FAIL (1 finding(s))", err)

    def test_missing_inputs_are_clear_errors(self) -> None:
        missing = str(self.work / "absent.xml")
        code, _, err = run("--check", "--customui", missing)
        self.assertEqual(code, 1)
        self.assertIn(f"customUI file not found: {missing}", err)
        (self.work / "customUI14.xml").write_text(CUSTOMUI, encoding="utf-8")
        code, _, err = run("--check", "--customui", str(self.work / "customUI14.xml"),
                           "--callbacks", str(self.work / "absent.bas"))
        self.assertEqual(code, 1)
        self.assertIn("VBA callback module not found", err)

    def test_vba_parsing_details(self) -> None:
        text = ("Public Sub A(ByVal x As Long, Optional y As String = \"a, b\")\r\n"
                "Sub B _\r\n  (ByRef c As Office.IRibbonControl, _\r\n   n%)\r\n"
                "' Public Sub C(control As IRibbonControl)\r\n"
                "Rem Public Sub D(control As IRibbonControl)\r\n"
                "Public Declare PtrSafe Sub Sleep Lib \"kernel32\" (ByVal ms As Long)\r\n")
        procedures = inject_ribbon.vba_procedures(text)
        self.assertEqual([(p.name, p.scope, p.line) for p in procedures],
                         [("A", "public", 1), ("B", "public", 2)])
        second = procedures[1].params
        assert second is not None
        self.assertEqual([(p.name, p.mode, p.type) for p in second],
                         [("c", "byref", "Office.IRibbonControl"), ("n", "", "Integer")])
        first = procedures[0].params
        assert first is not None
        self.assertEqual([p.special for p in first], [False, True])


if __name__ == "__main__":
    unittest.main()
