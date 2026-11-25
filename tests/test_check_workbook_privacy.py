"""Tests for tools/check_workbook_privacy.py, on packages built in memory.

Every fixture names the person "Avery" and the organization "Contoso", so a
test can check that no finding repeats a value it found.
"""

from __future__ import annotations

import codecs
import contextlib
import io
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import build_workbook  # noqa: E402
import check_workbook_privacy as privacy  # noqa: E402
import vba_project  # noqa: E402
import workbook_template  # noqa: E402

CONTENT_TYPES = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.'
                 'openxmlformats.org/package/2006/content-types"><Default Extension="xml" '
                 'ContentType="application/xml"/>{extra}</Types>')
WORKBOOK = ('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            'xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac">'
            '{inner}<sheets><sheet name="Inputs" sheetId="1"/></sheets></workbook>')
ABS_PATH = ('<mc:AlternateContent><mc:Choice Requires="x15"><x15ac:absPath url="{url}"/>'
            '</mc:Choice></mc:AlternateContent>')
RELATIONSHIPS = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                 'relationships">{inner}</Relationships>')
EXTERNAL = ('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/'
            '2006/relationships/externalLinkPath" Target="{target}" TargetMode="External"/>')
CORE = ('<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/'
        'metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">'
        '<dc:title>Synthetic</dc:title>{inner}</cp:coreProperties>')
APP = ('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
       'extended-properties">{inner}</Properties>')
CUSTOM = ('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
          'custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/'
          'docPropsVTypes">{inner}</Properties>')
PROPERTY = ('<property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="2" name="{name}">'
            '<vt:lpwstr>{value}</vt:lpwstr></property>')
LABEL_GUID = "0e5f1a2b-3c4d-4e5f-8a9b-0c1d2e3f4a5b"
# A name that starts with a letter outside ASCII, as a Windows profile folder can.
ACCENTED = "\u00c1very"
# Names in Chinese characters, one outside the Basic Multilingual Plane (a
# surrogate pair in UTF-16).
CJK = "\u5f20\u4f1f"
CJK_SUPPLEMENTARY = "\U00020000\u4f1f"
# A name that both cp936 and cp932 can store.
KANJI = "\u5c71\u7530"
SECRETS = ("avery", "contoso", LABEL_GUID, ACCENTED.lower(),
           "\u044f\u0440\u043e\u0441\u043b\u0430\u0432", CJK, CJK_SUPPLEMENTARY, KANJI)
TYPES = "application/vnd.openxmlformats-officedocument.spreadsheetml."


def package(parts: dict[str, str | bytes] | None = None, extra_types: str = "") -> bytes:
    """A minimal workbook package with the given parts added or replaced."""
    contents: dict[str, str | bytes] = {
        "[Content_Types].xml": CONTENT_TYPES.format(extra=extra_types),
        "_rels/.rels": RELATIONSHIPS.format(
            inner='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                  'officeDocument/2006/relationships/officeDocument" '
                  'Target="xl/workbook.xml"/>'),
        "xl/workbook.xml": WORKBOOK.format(inner=""),
        "docProps/core.xml": CORE.format(inner=""),
        **(parts or {})}
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in contents.items():
            archive.writestr(name, content)
    return buffer.getvalue()


CLEAN = package({"docProps/app.xml": APP.format(inner="<Company></Company><Manager/>"),
                 "xl/sharedStrings.xml": '<sst><si><t>See https://www.example.com/docs and '
                                         'C:\\Program Files\\Example</t></si></sst>'})


def vba_with(source_line: str) -> bytes:
    """A VBA project, written as the build writes one, with the line in a module's source."""
    module = vba_project.Module("MPaths", 'Attribute VB_Name = "MPaths"\r\nOption Explicit\r\n'
                                          f"{source_line}\r\n")
    return vba_project.build_vba_project([vba_project.document_module("ThisWorkbook", True),
                                          module])


def vba_in_code_page(code_page: int, source_line: str) -> bytes:
    """A VBA project like vba_with()'s, but declaring another code page in its dir
    stream and storing its module source in that code page, written with the
    parts of tools/devbuild/vba_project.py."""
    module = vba_project.Module("MPaths", 'Attribute VB_Name = "MPaths"\r\nOption Explicit\r\n'
                                          f"{source_line}\r\n")
    modules = [vba_project.document_module("ThisWorkbook", True), module]
    record = b"\x03\x00\x02\x00\x00\x00"  # PROJECTCODEPAGE, 2 bytes
    directory = vba_project.build_dir("VBAProject", modules)
    assert directory.count(record + (1252).to_bytes(2, "little")) == 1
    directory = directory.replace(record + (1252).to_bytes(2, "little"),
                                  record + code_page.to_bytes(2, "little"))
    encoding = codecs.lookup(f"cp{code_page}").name
    vba = {"_VBA_PROJECT": bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]),
           "dir": vba_project.compress(directory)}
    for item in modules:
        vba[item.name] = vba_project.compress(item.source.encode(encoding))
    project_id = "{00000000-0000-0000-0000-000000000005}"
    return vba_project.write_cfb({"PROJECT": vba_project.build_project_text(project_id, modules),
                                  "PROJECTwm": vba_project.build_projectwm(modules),
                                  "VBA": vba})


class ScanTest(unittest.TestCase):
    def findings(self, parts: dict[str, str | bytes], extra_types: str = "") -> list[str]:
        found = [str(finding) for finding in privacy.scan(package(parts, extra_types))]
        for finding in found:
            with self.subTest(finding=finding):
                for secret in SECRETS:
                    self.assertNotIn(secret, finding.lower())
        return found

    def assertFinds(self, parts: dict[str, str | bytes], *expected: str,
                    extra_types: str = "") -> None:
        found = self.findings(parts, extra_types)
        for text in expected:
            with self.subTest(expected=text):
                self.assertIn(text, found)

    def test_a_clean_package_passes(self) -> None:
        self.assertEqual(privacy.scan(CLEAN), [])

    def test_the_layout_template_fixture_passes(self) -> None:
        self.assertEqual(privacy.scan(workbook_template.template()), [])

    @unittest.skipUnless((ROOT / "src/workbook/Template.xlsx").is_file(),
                         "src/workbook/Template.xlsx is not committed")
    def test_the_committed_layout_template_passes(self) -> None:
        self.assertEqual(privacy.scan_file(ROOT / "src/workbook/Template.xlsx"), [])

    @unittest.skipIf(shutil.which("git") is None, "git is not installed")
    def test_a_production_build_passes(self) -> None:
        with tempfile.TemporaryDirectory(prefix="privacy-") as directory:
            target = Path(directory) / "Workbook-production.xlsm"
            build_workbook.build(ROOT, "production", target)
            self.assertEqual(privacy.scan_file(target), [])

    def test_an_abs_path_to_a_onedrive_folder_fails(self) -> None:
        for url, description in (
                ("C:\\Users\\avery\\OneDrive - Contoso\\Finance\\",
                 "a OneDrive for Business folder named after its organization"),
                ("https://contoso-my.sharepoint.com/personal/avery_contoso_com/Documents/",
                 "a OneDrive for Business URL"),
                ("https://d.docs.live.net/0123abcd/Documents/", "a OneDrive URL"),
                ("C:\\Users\\avery\\Documents\\", "an absolute path under a user profile"),
                ("/Users/avery/Library/CloudStorage/", "an absolute path under a user profile"),
                ("D:\\Projects\\Risk\\", "an absolute path")):
            with self.subTest(url=url):
                self.assertEqual(self.findings({"xl/workbook.xml": WORKBOOK.format(
                    inner=ABS_PATH.format(url=url))}),
                    [f"xl/workbook.xml: path or URL (x15ac:absPath holds {description})"])

    def test_a_onedrive_url_in_a_relationship_target_fails(self) -> None:
        target = ("https://contoso-my.sharepoint.com/personal/avery_contoso_com/Documents/"
                  "Inputs.xlsx")
        self.assertEqual(self.findings({"xl/externalLinks/_rels/externalLink1.xml.rels":
                                        RELATIONSHIPS.format(inner=EXTERNAL.format(
                                            target=target))}),
                         ["xl/externalLinks/_rels/externalLink1.xml.rels: path or URL "
                          "(Relationship@Target holds a OneDrive for Business URL)"])

    def test_paths_and_urls_fail_in_any_xml_part(self) -> None:
        rels = "xl/worksheets/_rels/sheet1.xml.rels"
        cases = {
            "https://contoso.sharepoint.com/sites/Risk/Shared Documents/Inputs.xlsx":
                "a SharePoint URL",
            "file:///C:/Data/Inputs.xlsx": "a file: URL",
            "\\\\?\\D:\\Data\\Inputs.xlsx": "an extended-length path that starts with \\\\?\\",
            "\\\\fileserver.contoso.local\\risk\\Inputs.xlsx": "a path on a network share",
            "/home/avery/data/Inputs.xlsx": "an absolute path under a user profile",
            "C:/Users/avery/Desktop/Inputs.xlsx": "an absolute path under a user profile",
            "D:\\Sync\\OneDrive\\Inputs.xlsx": "a OneDrive folder",
        }
        for target, description in cases.items():
            with self.subTest(target=target):
                self.assertEqual(self.findings({rels: RELATIONSHIPS.format(
                    inner=EXTERNAL.format(target=target))}),
                    [f"{rels}: path or URL (Relationship@Target holds {description})"])
        connection = ('<connections><connection id="1" name="Positions"><dbPr connection='
                      '"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\avery\\'
                      'positions.accdb"/></connection></connections>')
        self.assertEqual(self.findings({"xl/connections.xml": connection}),
                         ["xl/connections.xml: path or URL (dbPr@connection holds an absolute "
                          "path under a user profile)"])
        strings = ("<sst><si><t>Saved from https://contoso.sharepoint.com/sites/Risk</t></si>"
                   "</sst>")
        self.assertEqual(self.findings({"xl/sharedStrings.xml": strings}),
                         ["xl/sharedStrings.xml: path or URL (t holds a SharePoint URL)"])

    def test_public_profiles_system_folders_and_web_urls_are_not_paths(self) -> None:
        for text in ("C:\\Users\\Public\\Documents\\Inputs.xlsx", "C:\\Users\\<name>\\",
                     "C:\\Windows\\System32\\stdole2.tlb",
                     "C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL",
                     "https://www.example.com/home/avery/", "../Users/avery/Inputs.xlsx",
                     "/Users/Shared/Inputs.xlsx", "OneDrive-Personal is a folder name",
                     "a OneDrive - Contoso folder in prose",
                     "\\\\\\\\ is two escaped backslashes"):
            with self.subTest(text=text):
                self.assertIsNone(privacy.describe(text))

    def test_a_user_profile_path_in_the_vba_project_fails(self) -> None:
        user_path = "C:\\Users\\avery\\AppData\\Roaming\\Microsoft\\AddIns\\Tools.xlam"
        expected = ["xl/vbaProject.bin: path or URL (VBA project text holds an absolute path "
                    "under a user profile)"]
        # Module source is stored compressed: the path is found only once decompressed.
        compressed = vba_with(f'Const ADDIN = "{user_path}"')
        self.assertNotIn(b"C:\\Users", compressed)
        self.assertEqual(self.findings({"xl/vbaProject.bin": compressed}), expected)
        # Text stored as is, such as a reference in Excel's cache, as ASCII and UTF-16LE.
        for stored in (user_path.encode("ascii"), user_path.encode("utf-16-le")):
            with self.subTest(stored=stored[:8]):
                blob = vba_with("Option Base 0") + b"\x00\x13" + stored + b"\x00\x00"
                self.assertEqual(self.findings({"xl/vbaProject.bin": blob}), expected)
        url = "https://contoso.sharepoint.com/sites/Risk/Tools.xlam".encode("utf-16-le")
        self.assertEqual(self.findings({"xl/vbaProject.bin": b"\xd0\xcf\x11\xe0" + url}),
                         ["xl/vbaProject.bin: path or URL (VBA project text holds a SharePoint "
                          "URL)"])

    def test_a_profile_name_outside_ascii_in_the_vba_project_fails(self) -> None:
        expected = ["xl/vbaProject.bin: path or URL (VBA project text holds an absolute path "
                    "under a user profile)"]
        # Module source, compressed and stored in the project's code page, cp1252.
        compressed = vba_with(f'Const INPUTS = "C:\\Users\\{ACCENTED}\\Documents\\Inputs.xlsx"')
        self.assertEqual(self.findings({"xl/vbaProject.bin": compressed}), expected)
        # Text stored as is: UTF-16LE with a Latin or a Cyrillic first letter, and
        # cp1252 right after a byte that reads as a letter.
        path = "C:\\Users\\{}\\AppData\\Roaming\\Tools.xlam"
        for name, stored in (
                (ACCENTED, path.format(ACCENTED).encode("utf-16-le")),
                ("cyrillic", path.format("\u042f\u0440\u043e\u0441\u043b\u0430\u0432")
                 .encode("utf-16-le")),
                ("cp1252", b"\xc9" + path.format(ACCENTED).encode("cp1252"))):
            with self.subTest(name=name):
                blob = vba_with("Option Base 0") + b"\x00\x13" + stored + b"\x00\x00"
                self.assertEqual(self.findings({"xl/vbaProject.bin": blob}), expected)
        # The exemptions still hold, whatever follows them.
        for public in ("Public", "Default", "All Users"):
            with self.subTest(public=public):
                project = vba_with(f'Const SHARED = "C:\\Users\\{public}\\{ACCENTED}.xlsx"')
                self.assertEqual(self.findings({"xl/vbaProject.bin": project}), [])

    def test_the_dir_stream_is_read_in_the_declared_code_page(self) -> None:
        def record(number: int, payload: bytes) -> bytes:
            return number.to_bytes(2, "little") + len(payload).to_bytes(4, "little") + payload

        cyrillic = "\u042f\u0440\u043e\u0441\u043b\u0430\u0432"
        reference = f"*\\G{{00000000-0000-0000-0000-000000000000}}#1.0#0#C:\\Users\\{cyrillic}\\Tools.xlam#Tools"
        directory = (record(0x0003, (1251).to_bytes(2, "little"))
                     + record(0x000D, reference.encode("cp1251"))
                     + record(0x003E, "Tools".encode("utf-16-le"))
                     + record(0x001A, f"M{cyrillic}".encode("cp1251"))
                     + record(0x0031, (0).to_bytes(4, "little")))
        found = privacy.read_directory(directory)
        self.assertEqual(found.code_page, "cp1251")
        self.assertIn(reference, found.texts)
        self.assertIn("Tools", found.texts)
        self.assertEqual(found.offsets, {f"m{cyrillic.lower()}": 0})
        unknown = privacy.read_directory(record(0x0003, (9999).to_bytes(2, "little")))
        self.assertEqual(unknown.code_page, "cp1252")

    def test_a_cjk_profile_name_in_the_vba_project_fails(self) -> None:
        expected = ["xl/vbaProject.bin: path or URL (VBA project text holds an absolute path "
                    "under a user profile)"]
        base = vba_with("Option Base 0")
        # Stored as UTF-16LE, at an even and an odd offset, with a name in the Basic
        # Multilingual Plane and one that needs a surrogate pair.
        for label, name in (("BMP", CJK), ("surrogate pair", CJK_SUPPLEMENTARY)):
            stored = f"C:\\Users\\{name}\\Documents\\Inputs.xlsx".encode("utf-16-le")
            for padding in (b"\x00\x13", b"\x07"):
                with self.subTest(name=label, padding=padding):
                    blob = base + padding + stored + b"\x00\x00"
                    self.assertEqual(self.findings({"xl/vbaProject.bin": blob}), expected)
        # Module source in a multi-byte code page that the project declares.
        for code_page in (936, 932):
            with self.subTest(code_page=code_page):
                project = vba_in_code_page(
                    code_page, f'Const INPUTS = "C:\\Users\\{KANJI}\\Documents\\Inputs.xlsx"')
                self.assertEqual(self.findings({"xl/vbaProject.bin": project}), expected)
        # CJK text without the anchors of a path is not a finding.
        text = f"{CJK}\\{CJK}: Users {CJK}".encode("utf-16-le")
        self.assertEqual(self.findings({"xl/vbaProject.bin": base + text}), [])

    def test_system_library_references_in_the_vba_project_pass(self) -> None:
        # The build references stdole under C:\\Windows and the Office library under
        # C:\\Program Files, as every VBA project does.
        project = vba_with('Const NOTE = "See C:\\Program Files\\Example"')
        self.assertEqual(self.findings({"xl/vbaProject.bin": project}), [])

    def test_a_file_sharing_user_name_fails(self) -> None:
        book = WORKBOOK.format(inner='<fileSharing readOnlyRecommended="1" userName="Avery"/>')
        self.assertEqual(self.findings({"xl/workbook.xml": book}),
                         ["xl/workbook.xml: person or organization "
                          "(fileSharing@userName is not blank)"])
        blank = WORKBOOK.format(inner='<fileSharing readOnlyRecommended="1" userName=" "/>')
        self.assertEqual(self.findings({"xl/workbook.xml": blank}), [])

    def test_a_person_of_threaded_comments_fails(self) -> None:
        people = ('<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/'
                  'threadedcomments"><person displayName="Avery Contoso" '
                  'id="{00000000-0000-0000-0000-000000000001}" userId="avery@contoso.example" '
                  'providerId="AD"/></personList>')
        self.assertEqual(self.findings({"xl/persons/person.xml": people}),
                         ["xl/persons/person.xml: person or organization "
                          "(person@displayName is not blank)",
                          "xl/persons/person.xml: person or organization "
                          "(person@userId is not blank)"])

    def test_parts_are_found_by_content_type_wherever_they_are(self) -> None:
        comments = ('<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/'
                    'main"><authors><author>Avery</author></authors></comments>')
        hint = "person or organization (author is not blank)"
        # The layout that openpyxl writes, found by its path even without an override.
        self.assertEqual(self.findings({"xl/comments/comment1.xml": comments}),
                         [f"xl/comments/comment1.xml: {hint}"])
        cases = {
            "xl/notes/note1.xml": (f"{TYPES}comments+xml", comments, hint),
            "xl/people/people.xml": (
                "application/vnd.ms-excel.person+xml",
                '<personList><person displayName="Avery"/></personList>',
                "person or organization (person@displayName is not blank)"),
            "xl/changes/users.xml": (
                f"{TYPES}userNames+xml",
                '<users><userInfo guid="{00000000-0000-0000-0000-000000000002}" '
                'name="Avery"/></users>',
                "person or organization (userInfo@name is not blank)"),
            "props/core.xml": (
                "application/vnd.openxmlformats-package.core-properties+xml",
                CORE.format(inner="<dc:creator>Avery</dc:creator>"),
                "person or organization (creator is not blank)"),
            "props/extended.xml": (
                "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                APP.format(inner="<Company>Contoso</Company>"),
                "person or organization (Company is not blank)"),
            "props/custom.xml": (
                "application/vnd.openxmlformats-officedocument.custom-properties+xml",
                CUSTOM.format(inner=PROPERTY.format(name="ContentTypeId", value="0x0101")),
                "SharePoint metadata (property ContentTypeId is a SharePoint library "
                "property)"),
            "labels/info.xml": (
                "application/vnd.ms-office.classificationlabels+xml", "<labelList/>",
                "sensitivity label (a Purview sensitivity label part)"),
        }
        for part, (content_type, text, expected) in cases.items():
            override = f'<Override PartName="/{part.upper()}" ContentType="{content_type}"/>'
            with self.subTest(part=part):
                found = self.findings({part: text}, extra_types=override)
                self.assertIn(f"{part}: {expected}", found)
        # Without its content type, a part at another path is not one of these.
        self.assertEqual(self.findings({"xl/notes/note1.xml": comments}), [])

    def test_xml_parts_without_an_xml_suffix_are_read(self) -> None:
        core = CORE.format(inner="<dc:creator>Avery</dc:creator>")
        cases = {
            "props/core.dat": ("application/vnd.openxmlformats-package.core-properties+xml",
                               core, "person or organization (creator is not blank)"),
            "data/settings.bin": ("application/xml",
                                  '<settings path="C:\\Users\\avery\\Documents\\"/>',
                                  "path or URL (settings@path holds an absolute path under a "
                                  "user profile)"),
            "xl/extra/item.dat": ("application/vnd.example.settings+xml", "<settings>",
                                  "unreadable (is not well-formed XML, so it cannot be "
                                  "checked)"),
        }
        for part, (content_type, text, expected) in cases.items():
            override = f'<Override PartName="/{part}" ContentType="{content_type}"/>'
            with self.subTest(part=part):
                self.assertEqual(self.findings({part: text}, extra_types=override),
                                 [f"{part}: {expected}"])
        # A part with a role by its path is read whatever its suffix.
        self.assertEqual(self.findings({"docMetadata/LabelInfo.dat": "<labelList>"}),
                         ["docMetadata/LabelInfo.dat: sensitivity label (a Purview "
                          "sensitivity label part)",
                          "docMetadata/LabelInfo.dat: unreadable (is not well-formed XML, so "
                          "it cannot be checked)"])
        # VML and binary parts are not parsed unless declared XML.
        vml = ('<Default Extension="vml" ContentType="application/vnd.openxmlformats-'
               'officedocument.vmlDrawing"/><Default Extension="bin" ContentType="application/'
               'vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>')
        self.assertEqual(self.findings({"xl/drawings/vmlDrawing1.vml": "<xml><v:shape><br>",
                                        "xl/printerSettings/printerSettings1.bin": b"\x00\x01"},
                                       extra_types=vml), [])

    def test_comment_authors_and_revision_users_fail(self) -> None:
        comments = ('<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/'
                    'main"><authors><author>Avery</author><author>'
                    'tc={00000000-0000-0000-0000-000000000001}</author></authors></comments>')
        self.assertEqual(self.findings({"xl/comments1.xml": comments}),
                         ["xl/comments1.xml: person or organization (author is not blank)"])
        threaded_only = comments.replace("<author>Avery</author>", "")
        self.assertEqual(self.findings({"xl/comments1.xml": threaded_only}), [])
        users = ('<users xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                 'count="1"><userInfo guid="{00000000-0000-0000-0000-000000000002}" '
                 'name="Avery" id="1" dateTime="2026-11-02T10:00:00"/></users>')
        headers = ('<headers xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                   '<header guid="{00000000-0000-0000-0000-000000000003}" '
                   'dateTime="2026-11-02T10:00:00" userName="Avery"/></headers>')
        self.assertEqual(self.findings({"xl/revisions/userNames.xml": users,
                                        "xl/revisions/revisionHeaders.xml": headers}),
                         ["xl/revisions/userNames.xml: person or organization "
                          "(userInfo@name is not blank)",
                          "xl/revisions/revisionHeaders.xml: person or organization "
                          "(header@userName is not blank)"])

    def test_document_properties_that_name_a_person_or_organization_fail(self) -> None:
        self.assertEqual(self.findings({
            "docProps/core.xml": CORE.format(inner="<dc:creator>Avery</dc:creator>"
                                                   "<cp:lastModifiedBy>Avery</cp:lastModifiedBy>"),
            "docProps/app.xml": APP.format(inner="<Company>Contoso</Company>"
                                                 "<Manager>Avery</Manager>")}),
            [f"{part}: person or organization ({field} is not blank)"
             for part, fields in privacy.PERSONAL_FIELDS.items() for field in fields])

    def test_a_purview_label_property_fails(self) -> None:
        properties = "".join(PROPERTY.format(name=f"MSIP_Label_{LABEL_GUID}_{field}",
                                             value=value)
                             for field, value in (("Enabled", "true"), ("Name", "Contoso"),
                                                  ("SiteId", LABEL_GUID)))
        self.assertEqual(self.findings({"docProps/custom.xml": CUSTOM.format(inner=properties)}),
                         ["docProps/custom.xml: sensitivity label (property MSIP_Label_… names "
                          "a Purview sensitivity label)"])

    def test_a_label_info_part_fails(self) -> None:
        label = (f'<clbl:labelList xmlns:clbl="http://schemas.microsoft.com/office/2020/mipLabel'
                 f'Metadata"><clbl:label id="{{{LABEL_GUID}}}" enabled="1" method="Standard" '
                 f'siteId="{{{LABEL_GUID}}}" removed="0"/></clbl:labelList>')
        self.assertIn("docMetadata/LabelInfo.xml: sensitivity label (a Purview sensitivity "
                      "label part)", self.findings({"docMetadata/LabelInfo.xml": label}))

    def test_references_to_a_label_part_fail(self) -> None:
        override = ('<Override PartName="/docMetadata/LabelInfo.xml" ContentType="application/'
                    'vnd.ms-office.classificationlabels+xml"/>')
        self.assertEqual(self.findings({}, extra_types=override),
                         ["[Content_Types].xml: sensitivity label (Override@PartName refers to "
                          "a Purview sensitivity label part)",
                          "[Content_Types].xml: sensitivity label (Override@ContentType refers "
                          "to a Purview sensitivity label part)"])
        relationship = ('<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/'
                        '2020/02/relationships/classificationlabels" '
                        'Target="docMetadata/LabelInfo.xml"/>')
        self.assertEqual(self.findings({"_rels/.rels": RELATIONSHIPS.format(
            inner=relationship)}),
            ["_rels/.rels: sensitivity label (Relationship@Type refers to a Purview "
             "sensitivity label part)",
             "_rels/.rels: sensitivity label (Relationship@Target refers to a Purview "
             "sensitivity label part)"])

    def test_sharepoint_library_properties_fail(self) -> None:
        cases = {"ContentTypeId": "property ContentTypeId",
                 "_dlc_DocIdItemGuid": "property _dlc_…",
                 "MediaServiceImageTags": "property MediaService…",
                 "TaxCatchAll": "property TaxCatchAll",
                 "lcf76f155ced4ddcb4097134ff3c332f":
                     "property lcf76f155ced4ddcb4097134ff3c332f",
                 "_ip_UnifiedCompliancePolicyProperties": "property _ip_UnifiedCompliancePolicy…"}
        for name, hint in cases.items():
            with self.subTest(name=name):
                custom = CUSTOM.format(inner=PROPERTY.format(name=name, value="0x0101"))
                self.assertEqual(self.findings({"docProps/custom.xml": custom}),
                                 [f"docProps/custom.xml: SharePoint metadata ({hint} is a "
                                  "SharePoint library property)"])
        own = CUSTOM.format(inner=PROPERTY.format(name="Regime", value="Synthetic"))
        self.assertEqual(self.findings({"docProps/custom.xml": own}), [])

    def test_a_sharepoint_schema_in_custom_xml_fails(self) -> None:
        schema = ('<p:properties xmlns:p="http://schemas.microsoft.com/office/2006/metadata/'
                  'properties" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                  '<documentManagement><TaxCatchAll xmlns="0c1d2e3f-0000-0000-0000-0000000000'
                  '01"/></documentManagement></p:properties>')
        self.assertEqual(self.findings({"customXml/item1.xml": schema}),
                         ["customXml/item1.xml: SharePoint metadata (holds elements of a "
                          "SharePoint metadata namespace)"])
        content_type = ('<ct:contentTypeSchema xmlns:ct="http://schemas.microsoft.com/office/'
                        '2006/metadata/contentType" ct:_="" ma:_="" xmlns:ma="http://'
                        'schemas.microsoft.com/office/2006/metadata/properties/metaAttributes"'
                        ' ct:contentTypeName="Document"/>')
        self.assertEqual(self.findings({"customXml/item2.xml": content_type}),
                         ["customXml/item2.xml: SharePoint metadata (holds elements of a "
                          "SharePoint metadata namespace)"])
        item_props = ('<ds:datastoreItem ds:itemID="{00000000-0000-0000-0000-000000000004}" '
                      'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/'
                      'customXml"><ds:schemaRefs><ds:schemaRef ds:uri="http://schemas.'
                      'microsoft.com/office/2006/metadata/properties"/></ds:schemaRefs>'
                      '</ds:datastoreItem>')
        self.assertEqual(self.findings({"customXml/itemProps1.xml": item_props}),
                         ["customXml/itemProps1.xml: SharePoint metadata (schemaRef@uri "
                          "refers to a SharePoint metadata schema)"])
        bibliography = ('<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/'
                        '2006/bibliography"/>')
        self.assertEqual(self.findings({"customXml/item3.xml": bibliography}), [])

    def test_encoding_and_character_references_hide_nothing(self) -> None:
        utf16 = ('<?xml version="1.0" encoding="UTF-16"?>' + WORKBOOK.format(
            inner='<fileSharing userName="Avery"/>')).encode("utf-16")
        self.assertEqual(self.findings({"xl/workbook.xml": utf16}),
                         ["xl/workbook.xml: person or organization "
                          "(fileSharing@userName is not blank)"])
        escaped = RELATIONSHIPS.format(inner=EXTERNAL.format(
            target="https://contoso-my&#46;sharepoint&#x2E;com/personal/avery/"))
        self.assertEqual(self.findings({"xl/_rels/workbook.xml.rels": escaped}),
                         ["xl/_rels/workbook.xml.rels: path or URL (Relationship@Target holds "
                          "a OneDrive for Business URL)"])
        label = CUSTOM.format(inner=PROPERTY.format(name="MSIP&#95;Label_x_Enabled",
                                                    value="true"))
        self.assertEqual(self.findings({"docProps/custom.xml": label}),
                         ["docProps/custom.xml: sensitivity label (property MSIP_Label_… "
                          "names a Purview sensitivity label)"])

    def test_a_part_that_cannot_be_read_fails(self) -> None:
        self.assertEqual(self.findings({"docProps/app.xml": "<Properties><Company>"}),
                         ["docProps/app.xml: unreadable (is not well-formed XML, so it cannot "
                          "be checked)"])
        entities = ('<?xml version="1.0"?><!DOCTYPE p [<!ENTITY a "Avery">]>'
                    '<Properties><Company>&a;</Company></Properties>')
        self.assertEqual(self.findings({"docProps/app.xml": entities}),
                         ["docProps/app.xml: unreadable (declares a DTD, so it cannot be "
                          "checked)"])
        # A part that is not XML is not parsed: an image or the VBA project.
        self.assertEqual(self.findings({"xl/media/image1.png": b"\x89PNG\r\n\x1a\n\0"}), [])

    def test_something_other_than_a_package_is_refused(self) -> None:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("readme.txt", "text")
        for data in (b"not a zip", buffer.getvalue(), b""):
            with self.subTest(data=data[:8]):
                self.assertFalse(privacy.is_package(data))
                with self.assertRaises(privacy.NotAPackage):
                    privacy.scan(data)
        self.assertTrue(privacy.is_package(CLEAN))

    def test_a_damaged_vba_project_is_still_read_as_stored(self) -> None:
        for blob in (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\xff" * 600, b"\x01\x02",
                     vba_with("Option Base 0")[:1200]):
            with self.subTest(blob=blob[:4]):
                self.assertEqual(self.findings({"xl/vbaProject.bin": blob}), [])
        with self.assertRaises(ValueError):
            privacy.decompress(b"\x02")


class MainTest(unittest.TestCase):
    def run_main(self, files: dict[str, bytes]) -> tuple[int, list[str]]:
        with tempfile.TemporaryDirectory(prefix="privacy-") as directory:
            paths = []
            for name, data in files.items():
                path = Path(directory) / name
                if data != b"<missing>":
                    path.write_bytes(data)
                paths.append(str(path))
            with contextlib.redirect_stdout(io.StringIO()) as output:
                status = privacy.main(paths)
            return status, [line.replace(f"{directory}/", "").replace(f"{directory}\\", "")
                            for line in output.getvalue().splitlines()]

    def test_clean_files_pass(self) -> None:
        self.assertEqual(self.run_main({"Clean.xlsx": CLEAN}),
                         (0, ["workbook-privacy: PASS (0 finding(s) in 1 file(s), "
                              "0 unreadable)"]))

    def test_findings_fail_with_the_file_part_kind_and_hint(self) -> None:
        private = package({"xl/workbook.xml": WORKBOOK.format(inner=ABS_PATH.format(
            url="C:\\Users\\avery\\OneDrive - Contoso\\"))})
        self.assertEqual(self.run_main({"Clean.xlsx": CLEAN, "Book.xlsx": private}),
                         (1, ["Book.xlsx: xl/workbook.xml: path or URL (x15ac:absPath holds a "
                              "OneDrive for Business folder named after its organization)",
                              "workbook-privacy: FAIL (1 finding(s) in 2 file(s), "
                              "0 unreadable)"]))

    def test_unreadable_input_exits_with_2(self) -> None:
        status, lines = self.run_main({"Book.xlsx": b"not a zip", "Gone.xlsx": b"<missing>",
                                       "Clean.xlsx": CLEAN})
        self.assertEqual(status, 2)
        self.assertTrue(lines[0].startswith("Book.xlsx: cannot be checked: not a ZIP"), lines)
        self.assertTrue(lines[1].startswith("Gone.xlsx: cannot be checked: "), lines)
        self.assertEqual(lines[2], "workbook-privacy: ERROR (0 finding(s) in 3 file(s), "
                                   "2 unreadable)")

    def test_the_command_line_runs(self) -> None:
        with tempfile.TemporaryDirectory(prefix="privacy-") as directory:
            path = Path(directory) / "Clean.xlsx"
            path.write_bytes(CLEAN)
            result = subprocess.run([sys.executable, str(ROOT / "tools/check_workbook_privacy.py"),
                                     str(path)], capture_output=True, text=True, check=False)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("workbook-privacy: PASS", result.stdout)


if __name__ == "__main__":
    unittest.main()
