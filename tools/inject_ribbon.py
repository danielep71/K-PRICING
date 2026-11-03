#!/usr/bin/env python3
"""Inject RibbonX into one explicit Excel package, or check RibbonX against its VBA callbacks.

--target rewrites one named, untracked .xlsm or .xlam: it writes
customUI/customUI14.xml and its single package relationship, keeps every other
ZIP entry byte-for-byte, writes a backup and replaces the file atomically. It
never searches for workbooks and never runs Excel.

--check is offline: it validates the customUI XML (stable unique KPR control
ids, built-in imageMso icons only) and that every callback it names exists in
the VBA callback module with the Office 2010 signature, and reports Public
ribbon-shaped procedures the XML does not use. No package is involved.

Exit status: 0 success, 1 validation failure (reported on stderr), 2 usage error.
"""

from __future__ import annotations

import argparse
import contextlib
import io
import os
import posixpath
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
import zlib
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import quoteattr

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CUSTOMUI = ROOT / "src" / "ribbon" / "customUI14.xml"
DEFAULT_CALLBACKS = ROOT / "src" / "modules" / "KPR_UI_Ribbon.bas"

CUSTOMUI_NS = "http://schemas.microsoft.com/office/2009/07/customui"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
# customUI14 (Office 2010) uses the 2007 extensibility type; the 2006 type is
# the Office 2007 customUI.xml part, which this tool leaves alone.
RIBBON_REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
OFFICE_DOCUMENT_TYPES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
)
CUSTOMUI_PART = "customUI/customUI14.xml"
RIBBON_REL_ID = "rIdKprRibbon"
# Excel refuses a package whose main part type does not match its extension.
WORKBOOK_CONTENT_TYPES = {
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    ".xlam": "application/vnd.ms-excel.addin.macroEnabled.main+xml",
}
XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
ATTRIBUTE_ESCAPES = {"\t": "&#9;", "\n": "&#10;", "\r": "&#13;"}

MAX_XML_BYTES = 8 * 1024 * 1024
MAX_PACKAGE_BYTES = 256 * 1024 * 1024
MAX_ENTRY_BYTES = 1024 * 1024 * 1024
UNSAFE_XML_DECLARATION = re.compile(r"<!\s*(?:DOCTYPE|ENTITY)\b", re.IGNORECASE)

# ZIP records written by hand so that preserved entries are copied verbatim:
# zipfile can only write an entry by compressing it again.
LOCAL_HEADER = struct.Struct("<4s2B4HL2L2H")
CENTRAL_HEADER = struct.Struct("<4s4B4HL2L5H2L")
END_RECORD = struct.Struct("<4s4H2LH")
LOCAL_SIGNATURE = b"PK\x03\x04"
CENTRAL_SIGNATURE = b"PK\x01\x02"
END_SIGNATURE = b"PK\x05\x06"
DESCRIPTOR_SIGNATURE = b"PK\x07\x08"
ZIP64_LIMIT = 0xFFFFFFFF
ZIP64_EXTRA_ID = 0x0001
FLAG_ENCRYPTED = 0x1
FLAG_DESCRIPTOR = 0x8
FLAG_UTF8 = 0x800
# Fixed stamp for every entry this tool writes, so equal inputs give equal bytes.
FIXED_DATE_TIME = (1980, 1, 1, 0, 0, 0)

# Git variables that would point the tracked-file check at another repository.
GIT_REDIRECTS = ("GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_OBJECT_DIRECTORY",
                 "GIT_ALTERNATE_OBJECT_DIRECTORIES", "GIT_COMMON_DIR", "GIT_NAMESPACE")

CONTROL_ID = re.compile(r"^KPR_[A-Za-z0-9_]+$")
VBA_IDENTIFIER = re.compile(r"[A-Za-z][A-Za-z0-9_]*")
CUSTOM_IMAGE_ATTRIBUTES = ("image", "loadImage")
GETTER_ATTRIBUTE = re.compile(r"get[A-Z][A-Za-z]*")
# Item getters receive the item index between the control and the returned value.
ITEM_GETTERS = frozenset({"getItemLabel", "getItemID", "getItemImage",
                          "getItemScreentip", "getItemSupertip"})
TOGGLE_CONTROLS = frozenset({"toggleButton", "checkBox"})
SELECTION_CONTROLS = frozenset({"dropDown", "gallery"})

CONTINUATION = re.compile(r"(?:^|[ \t])_[ \t]*$")
PROCEDURE = re.compile(
    r"^\s*(?:(?P<scope>Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"(?P<name>[A-Za-z][A-Za-z0-9_]*)[%&!#@$]?\s*(?P<rest>.*)$", re.IGNORECASE)
PARAMETER = re.compile(
    r"^(?:(?P<optional>Optional)\s+)?(?:(?P<mode>ByVal|ByRef)\s+)?"
    r"(?:(?P<paramarray>ParamArray)\s+)?(?P<name>[A-Za-z][A-Za-z0-9_]*)(?P<suffix>[%&!#@$])?"
    r"\s*(?P<array>\(\s*\))?\s*(?:As\s+(?:New\s+)?(?P<type>[A-Za-z][A-Za-z0-9_.]*))?"
    r"\s*(?:=.*)?$", re.IGNORECASE)
TYPE_SUFFIXES = {"%": "Integer", "&": "Long", "!": "Single", "#": "Double",
                 "@": "Currency", "$": "String"}


class RibbonError(Exception):
    """A validation failure: nothing was changed."""


# --- Safe XML ---------------------------------------------------------------


def qualified(namespace: str, name: str) -> str:
    """ElementTree's {namespace}name form of a tag."""
    return "{" + namespace + "}" + name


def local_name(tag: str) -> str:
    return tag.rpartition("}")[2]


def parse_xml(label: str, data: bytes) -> ET.Element:
    """Parse UTF-8 XML after bounding its size and rejecting DTD/entity declarations."""
    if len(data) > MAX_XML_BYTES:
        raise RibbonError(f"{label}: XML exceeds the {MAX_XML_BYTES}-byte limit")
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError as error:
        raise RibbonError(f"{label}: XML is not UTF-8: {error}") from None
    if UNSAFE_XML_DECLARATION.search(text):
        raise RibbonError(f"{label}: XML DTD and entity declarations are prohibited")
    try:
        return ET.fromstring(text)  # noqa: S314 -- bounded input with DTD/entity declarations rejected above.
    except ET.ParseError as error:
        raise RibbonError(f"{label}: invalid XML: {error}") from None


def check_flat(label: str, root: ET.Element, namespace: str, children: Iterable[str]) -> None:
    """Accept only the flat shape Office writes, so that serialize() loses nothing."""
    allowed = {qualified(namespace, name) for name in children}
    for element in root:
        if (element.tag not in allowed or len(element) or (element.text or "").strip()
                or (element.tail or "").strip() or any(k.startswith("{") for k in element.attrib)):
            raise RibbonError(f"{label}: unexpected {element.tag} content; refusing to rewrite it")
    if root.attrib or (root.text or "").strip():
        raise RibbonError(f"{label}: unexpected root attributes or text; refusing to rewrite it")


def serialize(root: ET.Element, namespace: str) -> bytes:
    """Write a flat OPC document (see check_flat) as Office does, attributes in order."""
    parts = [XML_DECLARATION, f'<{local_name(root.tag)} xmlns="{namespace}">']
    for element in root:
        attributes = "".join(f" {key}={quoteattr(value, ATTRIBUTE_ESCAPES)}"
                             for key, value in element.attrib.items())
        parts.append(f"<{local_name(element.tag)}{attributes}/>")
    parts.append(f"</{local_name(root.tag)}>")
    return "".join(parts).encode("utf-8")


def read_bounded(path: Path, limit: int) -> bytes:
    with path.open("rb") as handle:
        data = handle.read(limit + 1)
    if len(data) > limit:
        raise RibbonError(f"{path}: larger than the {limit}-byte limit")
    return data


def load_customui(path: Path) -> tuple[bytes, ET.Element]:
    """Return the customUI file's bytes and root after checking the Office 2010 root."""
    if not path.is_file():
        raise RibbonError(f"customUI file not found: {path} (pass --customui PATH)")
    data = read_bounded(path, MAX_XML_BYTES)
    root = parse_xml(str(path), data)
    if root.tag != qualified(CUSTOMUI_NS, "customUI"):
        raise RibbonError(f"{path}: the root element must be customUI in namespace "
                          f"{CUSTOMUI_NS} (found {root.tag})")
    return data, root


# --- Verbatim ZIP rewriting -------------------------------------------------


@dataclass(frozen=True)
class Member:
    info: zipfile.ZipInfo
    raw: bytes  # verbatim local header, compressed data and any data descriptor

    @property
    def name(self) -> str:
        return self.info.filename


@dataclass(frozen=True)
class Package:
    members: tuple[Member, ...]
    comment: bytes
    archive: zipfile.ZipFile

    def find(self, name: str) -> Member | None:
        key = name.casefold()
        return next((m for m in self.members if m.name.casefold() == key), None)

    def read(self, member: Member) -> bytes:
        if member.info.file_size > MAX_XML_BYTES:
            raise RibbonError(f"{member.name}: larger than the {MAX_XML_BYTES}-byte XML limit")
        return self.archive.read(member.info)


def encoded_name(info: zipfile.ZipInfo) -> bytes:
    return info.orig_filename.encode("utf-8" if info.flag_bits & FLAG_UTF8 else "cp437")


def extra_ids(extra: bytes) -> set[int]:
    ids: set[int] = set()
    offset = 0
    while offset + 4 <= len(extra):
        header_id, size = struct.unpack_from("<2H", extra, offset)
        ids.add(header_id)
        offset += 4 + size
    return ids


def dos_stamp(date_time: tuple[int, int, int, int, int, int]) -> tuple[int, int]:
    year, month, day, hour, minute, second = date_time
    return hour << 11 | minute << 5 | second // 2, (year - 1980) << 9 | month << 5 | day


def check_entry(info: zipfile.ZipInfo) -> None:
    name = info.filename
    if info.flag_bits & FLAG_ENCRYPTED:
        raise RibbonError(f"ZIP entry {name!r} is encrypted")
    if info.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
        raise RibbonError(f"ZIP entry {name!r} uses unsupported compression {info.compress_type}")
    if (max(info.compress_size, info.file_size, info.header_offset) >= ZIP64_LIMIT
            or ZIP64_EXTRA_ID in extra_ids(info.extra)):
        raise RibbonError(f"ZIP entry {name!r} needs ZIP64, which Office packages do not use")
    if info.file_size > MAX_ENTRY_BYTES:
        raise RibbonError(f"ZIP entry {name!r} expands beyond {MAX_ENTRY_BYTES} bytes")
    if "\0" in info.orig_filename:
        raise RibbonError(f"ZIP entry name {info.orig_filename!r} contains a NUL byte")


def raw_member(data: bytes, info: zipfile.ZipInfo) -> bytes:
    """Return the entry's bytes exactly as stored, from local header to data descriptor."""
    check_entry(info)
    start = info.header_offset
    header = data[start:start + LOCAL_HEADER.size]
    if len(header) != LOCAL_HEADER.size or header[:4] != LOCAL_SIGNATURE:
        raise RibbonError(f"ZIP entry {info.filename!r} has no valid local header")
    name_length, extra_length = LOCAL_HEADER.unpack(header)[-2:]
    name_start = start + LOCAL_HEADER.size
    if data[name_start:name_start + name_length] != encoded_name(info):
        raise RibbonError(f"ZIP entry {info.filename!r}: local and central names differ")
    end = name_start + name_length + extra_length + info.compress_size
    if info.flag_bits & FLAG_DESCRIPTOR:
        end += 16 if data[end:end + 4] == DESCRIPTOR_SIGNATURE else 12
    if end > len(data):
        raise RibbonError(f"ZIP entry {info.filename!r} is truncated")
    return data[start:end]


def load_package(data: bytes) -> Package:
    """Open and integrity-check a ZIP package without changing anything."""
    try:
        archive = zipfile.ZipFile(io.BytesIO(data))
    except (zipfile.BadZipFile, ValueError) as error:
        raise RibbonError(f"not a ZIP package: {error}") from None
    infos = archive.infolist()
    if len(infos) >= 0xFFFF:
        raise RibbonError("too many ZIP entries for a package without ZIP64")
    members: list[Member] = []
    seen: set[str] = set()
    for info in infos:
        key = info.filename.casefold()
        if key in seen:  # OPC part names compare case-insensitively
            raise RibbonError(f"duplicate ZIP entry {info.filename!r}")
        seen.add(key)
        members.append(Member(info, raw_member(data, info)))
    try:
        damaged = archive.testzip()
    except (zipfile.BadZipFile, zlib.error, EOFError) as error:
        raise RibbonError(f"damaged ZIP package: {error}") from None
    if damaged is not None:
        raise RibbonError(f"ZIP entry {damaged!r} fails its CRC check")
    return Package(tuple(members), archive.comment, archive)


def new_member(name: str, payload: bytes) -> Member:
    """A deflated entry with a fixed timestamp and MS-DOS host, independent of the machine."""
    compressor = zlib.compressobj(9, zlib.DEFLATED, -15)
    compressed = compressor.compress(payload) + compressor.flush()
    info = zipfile.ZipInfo(name, FIXED_DATE_TIME)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.create_system = 0
    info.create_version = info.extract_version = 20
    info.CRC = zlib.crc32(payload)
    info.compress_size = len(compressed)
    info.file_size = len(payload)
    name_bytes = name.encode("ascii")
    dos_time, dos_date = dos_stamp(FIXED_DATE_TIME)
    header = LOCAL_HEADER.pack(LOCAL_SIGNATURE, 20, 0, 0, zipfile.ZIP_DEFLATED, dos_time,
                               dos_date, info.CRC, len(compressed), len(payload),
                               len(name_bytes), 0)
    return Member(info, header + name_bytes + compressed)


def central_record(info: zipfile.ZipInfo, offset: int) -> bytes:
    name = encoded_name(info)
    dos_time, dos_date = dos_stamp(info.date_time)
    header = CENTRAL_HEADER.pack(
        CENTRAL_SIGNATURE, info.create_version, info.create_system, info.extract_version,
        info.reserved, info.flag_bits, info.compress_type, dos_time, dos_date, info.CRC,
        info.compress_size, info.file_size, len(name), len(info.extra), len(info.comment),
        info.volume, info.internal_attr, info.external_attr, offset)
    return header + name + info.extra + info.comment


def assemble(members: Sequence[Member], comment: bytes) -> bytes:
    body = bytearray()
    directory = bytearray()
    for member in members:
        directory += central_record(member.info, len(body))
        body += member.raw
    if len(body) + len(directory) >= ZIP64_LIMIT:
        raise RibbonError("rebuilt package would need ZIP64")
    end = END_RECORD.pack(END_SIGNATURE, 0, 0, len(members), len(members), len(directory),
                          len(body), len(comment))
    return bytes(body + directory) + end + comment


# --- Package validation and injection plan ----------------------------------


@dataclass(frozen=True)
class Parts:
    content_types: Member
    content_types_root: ET.Element
    relationships: Member
    relationships_root: ET.Element
    workbook: str


@dataclass(frozen=True)
class Plan:
    members: tuple[Member, ...]
    changed: frozenset[str]  # case-folded names of written entries
    notes: tuple[str, ...]


def require_member(package: Package, name: str) -> Member:
    member = package.find(name)
    if member is None:
        raise RibbonError(f"package has no {name} part")
    return member


def part_name(target: str) -> str:
    """Resolve a package-relative relationship target to a ZIP entry name."""
    name = posixpath.normpath(target.lstrip("/"))
    if name.startswith("..") or name in ("", "."):
        raise RibbonError(f"relationship target {target!r} leaves the package")
    return name


def office_document(root: ET.Element) -> str:
    targets = [rel for rel in root if rel.get("Type") in OFFICE_DOCUMENT_TYPES]
    if len(targets) != 1:
        raise RibbonError(f"_rels/.rels must have exactly one officeDocument relationship "
                          f"(found {len(targets)})")
    rel = targets[0]
    if rel.get("TargetMode", "Internal") != "Internal" or not rel.get("Target"):
        raise RibbonError("the officeDocument relationship must target a part in the package")
    return part_name(rel.get("Target", ""))


def check_workbook_type(root: ET.Element, workbook: str, extension: str) -> None:
    expected = WORKBOOK_CONTENT_TYPES[extension]
    wanted = "/" + workbook.casefold()
    found = [o.get("ContentType") for o in root.iter(qualified(CONTENT_TYPES_NS, "Override"))
             if (o.get("PartName") or "").casefold() == wanted]
    if found != [expected]:
        raise RibbonError(f"{extension} package must declare /{workbook} as {expected} "
                          f"(found {', '.join(str(f) for f in found) or 'no override'})")


def validate_package(package: Package, extension: str) -> Parts:
    content_types = require_member(package, "[Content_Types].xml")
    relationships = require_member(package, "_rels/.rels")
    types_root = parse_xml(content_types.name, package.read(content_types))
    if types_root.tag != qualified(CONTENT_TYPES_NS, "Types"):
        raise RibbonError("[Content_Types].xml has no OPC Types root")
    check_flat(content_types.name, types_root, CONTENT_TYPES_NS, ("Default", "Override"))
    rels_root = parse_xml(relationships.name, package.read(relationships))
    if rels_root.tag != qualified(RELATIONSHIPS_NS, "Relationships"):
        raise RibbonError("_rels/.rels has no OPC Relationships root")
    check_flat(relationships.name, rels_root, RELATIONSHIPS_NS, ("Relationship",))
    workbook = require_member(package, office_document(rels_root))
    if local_name(parse_xml(workbook.name, package.read(workbook)).tag) != "workbook":
        raise RibbonError(f"{workbook.name} is not a SpreadsheetML workbook part")
    check_workbook_type(types_root, workbook.name, extension)
    return Parts(content_types, types_root, relationships, rels_root, workbook.name)


def is_ribbon_relationship(rel: ET.Element) -> bool:
    return (rel.get("Type") == RIBBON_REL_TYPE
            and rel.get("TargetMode", "Internal") == "Internal"
            and posixpath.normpath((rel.get("Target") or "").lstrip("/")).casefold()
            == CUSTOMUI_PART.casefold())


def update_relationships(root: ET.Element) -> tuple[bytes | None, list[str]]:
    """Keep exactly one customUI14 relationship, reusing its Id; leave the rest intact."""
    rels = [rel for rel in root if rel.tag == qualified(RELATIONSHIPS_NS, "Relationship")]
    matching = [rel for rel in rels if is_ribbon_relationship(rel)]
    if matching:
        rel_id = matching[0].get("Id", "")
        for duplicate in matching[1:]:
            root.remove(duplicate)
        notes = [f"_rels/.rels: kept relationship {rel_id}"]
        if len(matching) > 1:
            notes.append(f"_rels/.rels: removed {len(matching) - 1} duplicate relationship(s)")
    else:
        taken = {rel.get("Id") for rel in rels}
        rel_id, number = RIBBON_REL_ID, 2
        while rel_id in taken:
            rel_id, number = f"{RIBBON_REL_ID}{number}", number + 1
        ET.SubElement(root, qualified(RELATIONSHIPS_NS, "Relationship"),
                      {"Id": rel_id, "Type": RIBBON_REL_TYPE, "Target": CUSTOMUI_PART})
        notes = [f"_rels/.rels: added relationship {rel_id}"]
    ids = [rel.get("Id") for rel in root if rel.tag == qualified(RELATIONSHIPS_NS, "Relationship")]
    if not all(ids) or len(set(ids)) != len(ids):
        raise RibbonError("_rels/.rels has missing or duplicate relationship Ids")
    others = sum(1 for rel in root if rel.get("Type") == RIBBON_REL_TYPE) - 1
    if others:
        notes.append(f"_rels/.rels: left {others} other ui/extensibility relationship(s)")
    changed = len(matching) != 1
    return (serialize(root, RELATIONSHIPS_NS) if changed else None), notes


def update_content_types(root: ET.Element) -> bytes | None:
    """Add the xml Default the customUI part relies on, only when it is absent."""
    defaults = [el for el in root if el.tag == qualified(CONTENT_TYPES_NS, "Default")]
    if any((el.get("Extension") or "").casefold() == "xml" for el in defaults):
        return None
    element = ET.Element(qualified(CONTENT_TYPES_NS, "Default"),
                         {"Extension": "xml", "ContentType": "application/xml"})
    root.insert(list(root).index(defaults[-1]) + 1 if defaults else 0, element)
    return serialize(root, CONTENT_TYPES_NS)


def plan_injection(package: Package, parts: Parts, customui: bytes) -> Plan:
    """Decide the entries to write; every other entry is reused verbatim and in order."""
    replacements: dict[str, Member] = {}
    rels_payload, notes = update_relationships(parts.relationships_root)
    if rels_payload is not None:
        replacements[parts.relationships.name.casefold()] = new_member(
            parts.relationships.name, rels_payload)
    types_payload = update_content_types(parts.content_types_root)
    if types_payload is not None:
        replacements[parts.content_types.name.casefold()] = new_member(
            parts.content_types.name, types_payload)
        notes.append("[Content_Types].xml: added Default xml -> application/xml")
    existing = package.find(CUSTOMUI_PART)
    current = (package.archive.read(existing.info)
               if existing is not None and existing.info.file_size == len(customui) else None)
    if current != customui:
        replacements[CUSTOMUI_PART.casefold()] = new_member(CUSTOMUI_PART, customui)
        action = "added" if existing is None else "replaced"
        notes.insert(0, f"{CUSTOMUI_PART}: {action} ({len(customui)} bytes)")
    else:
        notes.insert(0, f"{CUSTOMUI_PART}: unchanged")
    changed = frozenset(replacements)
    members = [replacements.pop(m.name.casefold(), m) for m in package.members]
    members += replacements.values()  # only the customUI part can be new
    return Plan(tuple(members), changed, tuple(notes))


# --- Safe file replacement --------------------------------------------------


def refuse_tracked(target: Path) -> None:
    """Refuse a target that git tracks; a path outside any repository is untracked."""
    folder = target.parent.resolve()
    if not any((parent / ".git").exists() for parent in (folder, *folder.parents)):
        return
    env = {k: v for k, v in os.environ.items() if k not in GIT_REDIRECTS}
    env["LC_ALL"] = "C"  # the untracked case is recognised by git's message
    command = ["git", "-C", str(folder), "ls-files", "--error-unmatch", "--",
               f":(literal,icase){target.name}"]
    try:
        result = subprocess.run(command, capture_output=True, text=True, env=env,
                                check=False, timeout=60)
    except (OSError, subprocess.TimeoutExpired) as error:
        raise RibbonError(f"cannot confirm that {target} is untracked by git: {error}") from None
    if result.returncode == 0:
        raise RibbonError(f"{target} is tracked by git; refusing to rewrite a tracked Office "
                          "file. Copy it to an untracked path and inject there.")
    if result.returncode != 1 or "did not match" not in result.stderr:
        raise RibbonError(f"cannot confirm that {target} is untracked by git: "
                          f"{result.stderr.strip() or result.returncode}")


def check_target(target: Path, backup: Path) -> str:
    extension = target.suffix.casefold()
    if extension not in WORKBOOK_CONTENT_TYPES:
        raise RibbonError(f"{target}: only .xlsm and .xlam packages can be injected")
    if target.is_symlink() or not target.is_file():
        raise RibbonError(f"{target}: not a regular file")
    if backup.resolve() == target.resolve():
        raise RibbonError("the backup path must differ from the target")
    refuse_tracked(target)
    return extension


def write_backup(backup: Path, original: bytes) -> str:
    """Create the backup exclusively; an existing backup is reused only if identical."""
    try:
        handle = backup.open("xb")
    except FileExistsError:
        if backup.is_file() and not backup.is_symlink() and backup.read_bytes() == original:
            return f"backup: {backup} (already present, identical)"
        raise RibbonError(f"backup {backup} already exists and differs from the target; "
                          "move it away or pass --backup PATH") from None
    try:
        with handle:
            handle.write(original)
            handle.flush()
            os.fsync(handle.fileno())
        if backup.read_bytes() != original:
            raise RibbonError(f"backup {backup} did not read back identically")
    except BaseException:
        backup.unlink()
        raise
    return f"backup: {backup} (written)"


def verify_rebuilt(data: bytes, original: Package, plan: Plan, customui: bytes,
                   extension: str) -> None:
    """Re-open the written package: valid, converged, and every other entry verbatim."""
    package = load_package(data)
    again = plan_injection(package, validate_package(package, extension), customui)
    if again.changed:
        raise RibbonError("rebuilt package failed re-validation: injection did not converge")
    kept = [m.raw for m in original.members if m.name.casefold() not in plan.changed]
    after = [m.raw for m in package.members if m.name.casefold() not in plan.changed]
    if kept != after:
        raise RibbonError("rebuilt package failed re-validation: preserved entries differ")


def replace_target(target: Path, original: bytes, rebuilt: bytes, backup: Path,
                   check: tuple[Package, Plan, bytes, str]) -> str:
    """Write a temp file beside the target, verify it, back up, then os.replace."""
    handle, temp_name = tempfile.mkstemp(prefix=f".{target.name}.", suffix=".tmp",
                                         dir=target.parent)
    temp = Path(temp_name)
    try:
        with os.fdopen(handle, "wb") as stream:
            stream.write(rebuilt)
            stream.flush()
            os.fsync(stream.fileno())
        written = temp.read_bytes()
        if written != rebuilt:
            raise RibbonError("temporary package did not read back identically")
        verify_rebuilt(written, *check)
        shutil.copymode(target, temp)
        note = write_backup(backup, original)
        if target.read_bytes() != original:
            raise RibbonError(f"{target} changed while it was being rebuilt; nothing replaced")
        os.replace(temp, target)
    finally:
        with contextlib.suppress(FileNotFoundError):
            temp.unlink()
    return note


def inject(target: Path, customui_path: Path, backup: Path | None = None) -> list[str]:
    """Inject customUI14 into an explicit package; return the summary lines."""
    backup = backup or target.with_name(target.name + ".bak")
    extension = check_target(target, backup)
    customui, _ = load_customui(customui_path)
    original = read_bounded(target, MAX_PACKAGE_BYTES)
    package = load_package(original)
    plan = plan_injection(package, validate_package(package, extension), customui)
    if not plan.changed:
        return [f"inject-ribbon: {target} already carries this customUI; nothing written"]
    rebuilt = assemble(plan.members, package.comment)
    note = replace_target(target, original, rebuilt, backup,
                          (package, plan, customui, extension))
    return [f"inject-ribbon: updated {target}", *(f"- {n}" for n in (*plan.notes, note))]


# --- RibbonX and VBA callback consistency -----------------------------------


@dataclass(frozen=True)
class Finding:
    rule: str
    message: str


@dataclass(frozen=True)
class CallbackUse:
    name: str
    attribute: str
    tag: str
    control: str | None

    def where(self) -> str:
        target = f"<{self.tag} id={self.control!r}>" if self.control else f"<{self.tag}>"
        return f"{self.attribute} on {target}"


@dataclass(frozen=True)
class Slot:
    text: str
    types: frozenset[str | None]
    returned: bool = False


@dataclass(frozen=True)
class Parameter:
    name: str
    mode: str  # "byval", "byref" or "" (ByRef by default)
    type: str | None
    special: bool  # Optional, ParamArray or an array parameter

    def type_key(self) -> str | None:
        return self.type.rpartition(".")[2].casefold() if self.type else None


@dataclass(frozen=True)
class Procedure:
    name: str
    kind: str
    scope: str
    params: tuple[Parameter, ...] | None  # None when the list cannot be parsed
    line: int


CONTROL_SLOT = Slot("control As IRibbonControl", frozenset({"iribboncontrol"}))
RETURNED_SLOT = Slot("ByRef returnedVal [As Variant]", frozenset({"variant", None}), True)
INDEX_SLOT = Slot("index As Integer", frozenset({"integer", "long"}))
ON_LOAD = (Slot("ribbon As IRibbonUI", frozenset({"iribbonui"})),)
BUTTON_ACTION = (CONTROL_SLOT,)
TOGGLE_ACTION = (CONTROL_SLOT, Slot("pressed As Boolean", frozenset({"boolean"})))
SELECTION_ACTION = (CONTROL_SLOT, Slot("selectedId As String", frozenset({"string"})),
                    Slot("selectedIndex As Integer", frozenset({"integer", "long"})))
CHANGE = (CONTROL_SLOT, Slot("text As String", frozenset({"string"})))
GETTER = (CONTROL_SLOT, RETURNED_SLOT)
ITEM_GETTER = (CONTROL_SLOT, INDEX_SLOT, RETURNED_SLOT)


def is_callback_attribute(attribute: str) -> bool:
    return attribute in ("onLoad", "onAction", "onChange") or bool(
        GETTER_ATTRIBUTE.fullmatch(attribute))


def expected_shape(use: CallbackUse) -> tuple[Slot, ...]:
    """The Office 2010 callback signature for an attribute on a given control."""
    if use.attribute == "onLoad":
        return ON_LOAD
    if use.attribute == "onAction":
        if use.tag in TOGGLE_CONTROLS:
            return TOGGLE_ACTION
        return SELECTION_ACTION if use.tag in SELECTION_CONTROLS else BUTTON_ACTION
    if use.attribute == "onChange":
        return CHANGE
    return ITEM_GETTER if use.attribute in ITEM_GETTERS else GETTER


def ribbon_findings(root: ET.Element) -> tuple[list[Finding], list[CallbackUse]]:
    """Check control ids and images; collect every callback the XML names."""
    findings: list[Finding] = []
    uses: list[CallbackUse] = []
    seen: dict[str, int] = {}
    for element in root.iter():
        tag = local_name(element.tag)
        control = element.get("id")
        if control is not None:
            seen[control] = seen.get(control, 0) + 1
            if not CONTROL_ID.fullmatch(control):
                findings.append(Finding("control-id", f"<{tag}> id {control!r} must match "
                                        f"{CONTROL_ID.pattern} (stable KPR ids)"))
        for attribute in CUSTOM_IMAGE_ATTRIBUTES:
            if attribute in element.attrib:
                findings.append(Finding("custom-image", f"<{tag} id={control!r}> uses "
                                        f"{attribute}; use a built-in imageMso icon"))
        uses += [CallbackUse(value, attribute, tag, control)
                 for attribute, value in element.attrib.items()
                 if is_callback_attribute(attribute)]
    findings += [Finding("duplicate-id", f"control id {control!r} is used {count} times")
                 for control, count in seen.items() if count > 1]
    return findings, uses


def logical_lines(text: str) -> list[tuple[int, str]]:
    """Join VBA ` _` continuations; return (first physical line, logical line)."""
    lines: list[tuple[int, str]] = []
    pending: list[str] = []
    start = 1
    for number, line in enumerate(re.split(r"\r\n|\r|\n", text), 1):
        if not pending:
            start = number
        match = CONTINUATION.search(line)
        if match:
            pending.append(line[:match.start()] + " ")
            continue
        lines.append((start, "".join(pending) + line))
        pending = []
    if pending:
        lines.append((start, "".join(pending)))
    return lines


def code_only(line: str) -> str:
    """Drop the comment and empty every string literal of one logical VBA line."""
    kept: list[str] = []
    in_string = False
    for char in line:
        if char == '"':
            in_string = not in_string
            kept.append(char)
        elif in_string:
            continue
        elif char == "'":
            break
        else:
            kept.append(char)
    code = "".join(kept)
    return "" if re.match(r"\s*Rem\b", code, re.IGNORECASE) else code


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


def parse_parameter(text: str) -> Parameter | None:
    match = PARAMETER.match(text)
    if match is None:
        return None
    declared = match.group("type") or TYPE_SUFFIXES.get(match.group("suffix") or "")
    special = bool(match.group("optional") or match.group("paramarray") or match.group("array"))
    return Parameter(match.group("name"), (match.group("mode") or "").casefold(), declared,
                     special)


def parameter_list(rest: str) -> tuple[Parameter, ...] | None:
    """Parse the parenthesised list after a procedure name; None if it cannot be read."""
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
    params = [parse_parameter(piece) for piece in split_top_level(inner)]
    return None if None in params else tuple(p for p in params if p is not None)


def vba_procedures(text: str) -> list[Procedure]:
    procedures: list[Procedure] = []
    for number, line in logical_lines(text):
        match = PROCEDURE.match(code_only(line))
        if match is None:
            continue
        kind = " ".join(match.group("kind").split()).casefold()
        scope = (match.group("scope") or "public").casefold()  # VBA defaults to Public
        procedures.append(Procedure(match.group("name"), kind, "public" if scope == "global"
                                    else scope, parameter_list(match.group("rest")), number))
    return procedures


def slot_problem(position: int, slot: Slot, param: Parameter) -> str | None:
    label = f"parameter {position} ({param.name})"
    if param.special:
        return f"{label} must not be Optional, ParamArray or an array"
    if slot.returned:
        if param.mode == "byval":
            return f"{label} must be ByRef: Office returns the value through it"
        if param.type_key() not in slot.types:
            return f"{label} must be Variant or untyped, not {param.type}"
        return None
    if param.type_key() not in slot.types:
        return f"{label} must be declared {slot.text}, found {param.type or 'untyped'}"
    return None


def signature_findings(use: CallbackUse, procedure: Procedure) -> list[Finding]:
    title = f"{procedure.name} (line {procedure.line}, {use.where()})"
    if procedure.kind != "sub":
        return [Finding("callback-signature", f"{title} must be a Sub, not a {procedure.kind}")]
    if procedure.scope != "public":
        return [Finding("callback-signature", f"{title} must be Public, not {procedure.scope}")]
    if procedure.params is None:
        return [Finding("callback-signature", f"{title}: cannot read the parameter list")]
    shape = expected_shape(use)
    expected = f"Public Sub {use.name}({', '.join(slot.text for slot in shape)})"
    if len(procedure.params) != len(shape):
        return [Finding("callback-signature", f"{title} takes {len(procedure.params)} "
                        f"parameter(s); expected {expected}")]
    problems = [slot_problem(n, slot, param)
                for n, (slot, param) in enumerate(zip(shape, procedure.params), 1)]
    return [Finding("callback-signature", f"{title}: {problem}; expected {expected}")
            for problem in problems if problem]


def is_ribbon_shaped(procedure: Procedure) -> bool:
    return any(p.type_key() in ("iribboncontrol", "iribbonui") for p in procedure.params or ())


def callback_findings(uses: Iterable[CallbackUse],
                      procedures: Sequence[Procedure]) -> list[Finding]:
    """Every referenced callback must be a matching Public Sub; no unused ribbon Subs."""
    findings: list[Finding] = []
    by_name: dict[str, Procedure] = {}
    for procedure in procedures:
        key = procedure.name.casefold()
        if key in by_name:
            findings.append(Finding("duplicate-procedure", f"{procedure.name} is declared "
                                    f"more than once (line {procedure.line})"))
        by_name.setdefault(key, procedure)
    referenced: set[str] = set()
    for use in uses:
        if not VBA_IDENTIFIER.fullmatch(use.name):
            findings.append(Finding("callback-name", f"{use.where()} names {use.name!r}; "
                                    "use a plain VBA procedure name"))
            continue
        referenced.add(use.name.casefold())
        declared = by_name.get(use.name.casefold())
        if declared is None:
            findings.append(Finding("missing-callback", f"{use.name} ({use.where()}) has no "
                                    "Public Sub in the callback module"))
        else:
            findings += signature_findings(use, declared)
    findings += [Finding("dead-callback", f"{p.name} (line {p.line}) takes a ribbon "
                         "parameter but no RibbonX attribute names it")
                 for p in procedures if p.scope == "public" and p.kind == "sub"
                 and is_ribbon_shaped(p) and p.name.casefold() not in referenced]
    return list(dict.fromkeys(findings))


def read_vba(path: Path) -> str:
    if not path.is_file():
        raise RibbonError(f"VBA callback module not found: {path} (pass --callbacks PATH)")
    try:
        return read_bounded(path, MAX_XML_BYTES).decode("cp1252")
    except UnicodeDecodeError as error:
        raise RibbonError(f"{path}: not cp1252 text: {error}") from None


def check_ribbon(customui_path: Path, callbacks_path: Path) -> list[Finding]:
    _, root = load_customui(customui_path)
    findings, uses = ribbon_findings(root)
    procedures = vba_procedures(read_vba(callbacks_path))
    return findings + callback_findings(uses, procedures)


# --- Command line -----------------------------------------------------------


def parse_args(argv: Sequence[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--target", type=Path,
                      help="explicit .xlsm or .xlam package to rewrite (never discovered)")
    mode.add_argument("--check", action="store_true",
                      help="offline: validate the customUI XML against the VBA callbacks")
    parser.add_argument("--customui", type=Path, default=DEFAULT_CUSTOMUI,
                        help="customUI14 XML (default: src/ribbon/customUI14.xml)")
    parser.add_argument("--callbacks", type=Path,
                        help="VBA callback module for --check "
                             "(default: src/modules/KPR_UI_Ribbon.bas)")
    parser.add_argument("--backup", type=Path,
                        help="backup path for --target (default: <target>.bak)")
    args = parser.parse_args(argv)
    if args.check and args.backup:
        parser.error("--backup applies only to --target")
    if args.target and args.callbacks:
        parser.error("--callbacks applies only to --check")
    if args.check and args.callbacks is None:
        args.callbacks = DEFAULT_CALLBACKS
    return args


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        if args.target:
            print("\n".join(inject(args.target, args.customui, args.backup)))
            return 0
        findings = check_ribbon(args.customui, args.callbacks)
    except (RibbonError, OSError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1
    for finding in findings:
        print(f"FAIL: [{finding.rule}] {finding.message}", file=sys.stderr)
    if findings:
        print(f"ribbon-check: FAIL ({len(findings)} finding(s))", file=sys.stderr)
        return 1
    print("ribbon-check: PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
