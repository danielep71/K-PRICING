"""Minimal writer for an Office VBA project storage (vbaProject.bin).

Implements, from the [MS-OVBA] and [MS-CFB] specifications:
  - the VBA compression algorithm (2.4.1),
  - the dir stream (2.3.4.2), PROJECT / PROJECTwm streams (2.3.1),
  - the Data Encryption of CMG / DPB / GC (2.4.3),
  - a version-3 compound file writer (512-byte sectors, mini stream).
_VBA_PROJECT is written with Version 0xFFFF and no performance cache, so the
project holds no compiled code: the host reads it from source and compiles
each module when it is first used, and needs no Trust access to the VBA
project object model (docs/methodology/VALIDATION_PLAN.md, HB-05). Standard
modules, class modules and document modules are supported; forms, which need
a designer storage, are not. tests/test_devbuild.py reads the output back with
an independent decoder.
"""
from __future__ import annotations

import random
import struct
import uuid
from dataclasses import dataclass
from typing import Any

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF

WORKBOOK_BASE = "0{00020819-0000-0000-C000-000000000046}"
WORKSHEET_BASE = "0{00020820-0000-0000-C000-000000000046}"


# ============================================================================ compression
def _copy_token_params(diff: int) -> tuple[int, int]:
    bit_count = 4
    while (1 << bit_count) < diff:
        bit_count += 1
    max_len = (0xFFFF >> bit_count) + 3
    return bit_count, max_len


def compress(data: bytes) -> bytes:
    out = bytearray([0x01])
    pos = 0
    while pos < len(data) or (pos == 0 and not data):
        chunk = data[pos:pos + 4096]
        out += _compress_chunk(chunk)
        pos += 4096
        if not data:
            break
    return bytes(out)


# The MS-OVBA chunk compression algorithm as one loop, kept in the
# specification's shape so it can be checked against it step by step.
def _compress_chunk(chunk: bytes) -> bytes:  # noqa: C901
    body = bytearray()
    cur = 0
    n = len(chunk)
    index: dict[bytes, list[int]] = {}

    def add(i: int) -> None:
        if i + 3 <= n:
            index.setdefault(chunk[i:i + 3], []).append(i)

    while cur < n:
        flag_pos = len(body)
        body.append(0)
        flag = 0
        for bit in range(8):
            if cur >= n:
                break
            best_len, best_off = 0, 0
            if cur > 0:
                bit_count, max_len = _copy_token_params(cur)
                limit = min(max_len, n - cur)
                cands = index.get(chunk[cur:cur + 3], [])
                for cand in reversed(cands[-256:]):
                    off = cur - cand
                    if off > (1 << bit_count):
                        break
                    ln = 0
                    while ln < limit and chunk[cand + ln] == chunk[cur + ln]:
                        ln += 1
                    if ln > best_len:
                        best_len, best_off = ln, off
                        if ln == limit:
                            break
            if best_len >= 3:
                bit_count, _ = _copy_token_params(cur)
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                body += struct.pack("<H", token)
                flag |= 1 << bit
                for i in range(cur, cur + best_len):
                    add(i)
                cur += best_len
            else:
                body.append(chunk[cur])
                add(cur)
                cur += 1
        body[flag_pos] = flag
    if len(body) + 2 > 4098:
        if len(chunk) != 4096:
            # [MS-OVBA] allows a raw chunk only as 4096 bytes and a short chunk only
            # last; a short last chunk that will not compress cannot be represented
            # without padding the module source, so stop instead of writing it.
            raise ValueError(f"the last {len(chunk)} bytes do not compress into one chunk")
        return struct.pack("<H", (4095 & 0x0FFF) | (0b011 << 12)) + chunk
    header = ((len(body) + 2 - 3) & 0x0FFF) | (0b011 << 12) | (1 << 15)
    return struct.pack("<H", header) + bytes(body)


# ============================================================================ data encryption
def _encrypt(data: bytes, project_id: str, rng: random.Random) -> str:
    seed = rng.randrange(256)
    version = 2
    proj_key = sum(project_id.encode("cp1252")) & 0xFF
    out = bytearray([seed, seed ^ version, seed ^ proj_key])
    unenc1, enc1, enc2 = proj_key, seed ^ proj_key, seed ^ version

    def emit(byte: int) -> None:
        nonlocal unenc1, enc1, enc2
        b = byte ^ ((enc2 + unenc1) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, byte

    for _ in range((seed & 6) // 2):
        emit(rng.randrange(256))
    for b in struct.pack("<I", len(data)):
        emit(b)
    for b in data:
        emit(b)
    return out.hex().upper()


def decrypt(hex_text: str) -> bytes:
    raw = bytes.fromhex(hex_text)
    seed, version_enc, proj_key_enc = raw[0], raw[1], raw[2]
    assert seed ^ version_enc == 2, "bad version"
    proj_key = seed ^ proj_key_enc
    unenc1, enc1, enc2 = proj_key, proj_key_enc, version_enc
    pos = 3
    plain = bytearray()
    for _ in range((seed & 6) // 2 + 4):
        b = raw[pos] ^ ((enc2 + unenc1) & 0xFF)
        enc2, enc1, unenc1 = enc1, raw[pos], b
        plain.append(b)
        pos += 1
    length = struct.unpack("<I", bytes(plain[-4:]))[0]
    data = bytearray()
    for _ in range(length):
        b = raw[pos] ^ ((enc2 + unenc1) & 0xFF)
        enc2, enc1, unenc1 = enc1, raw[pos], b
        data.append(b)
        pos += 1
    return bytes(data)


# ============================================================================ project streams
STANDARD, CLASS, DOCUMENT = "standard", "class", "document"


@dataclass
class Module:
    name: str
    source: str            # full module text including Attribute lines
    document: bool = False
    kind: str = ""         # STANDARD, CLASS or DOCUMENT; derived from document when empty

    def __post_init__(self) -> None:
        if not self.kind:
            self.kind = DOCUMENT if self.document else STANDARD
        if self.kind not in (STANDARD, CLASS, DOCUMENT):
            raise ValueError(f"module {self.name}: unknown kind {self.kind!r}")
        self.document = self.kind == DOCUMENT


def document_module(name: str, workbook: bool) -> Module:
    base = WORKBOOK_BASE if workbook else WORKSHEET_BASE
    text = "\r\n".join([
        f'Attribute VB_Name = "{name}"',
        f'Attribute VB_Base = "{base}"',
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = False",
        "Attribute VB_PredeclaredId = True",
        "Attribute VB_Exposed = True",
        "Attribute VB_TemplateDerived = False",
        "Attribute VB_Customizable = True",
        "Option Explicit",
        "",
    ])
    return Module(name, text, document=True)


def _rec(rid: int, payload: bytes) -> bytes:
    return struct.pack("<HI", rid, len(payload)) + payload


def _u16(text: str) -> bytes:
    return text.encode("utf-16-le")


def _mb(text: str) -> bytes:
    return text.encode("cp1252")


REFERENCES = [
    ("stdole", r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation"),
    ("Office", r"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library"),
]


def build_dir(project_name: str, modules: list[Module]) -> bytes:
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 1))          # SYSKIND: Win32
    d += _rec(0x0002, struct.pack("<I", 0x0409))     # LCID
    d += _rec(0x0014, struct.pack("<I", 0x0409))     # LCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", 1252))       # CODEPAGE
    d += _rec(0x0004, _mb(project_name))             # NAME
    d += _rec(0x0005, b"") + _rec(0x0040, b"")       # DOCSTRING (+unicode)
    d += _rec(0x0006, b"") + _rec(0x003D, b"")       # HELPFILEPATH 1 / 2
    d += _rec(0x0007, struct.pack("<I", 0))          # HELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))          # LIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)       # VERSION: reserved=4, major, minor
    d += _rec(0x000C, b"") + _rec(0x003C, b"")       # CONSTANTS (+unicode)
    for ref_name, libid in REFERENCES:
        d += _rec(0x0016, _mb(ref_name)) + _rec(0x003E, _u16(ref_name))
        lib = _mb(libid)
        d += _rec(0x000D, struct.pack("<I", len(lib)) + lib + struct.pack("<IH", 0, 0))
    d += _rec(0x000F, struct.pack("<H", len(modules)))
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))     # PROJECTCOOKIE
    for m in modules:
        d += _rec(0x0019, _mb(m.name)) + _rec(0x0047, _u16(m.name))
        d += _rec(0x001A, _mb(m.name)) + _rec(0x0032, _u16(m.name))
        d += _rec(0x001C, b"") + _rec(0x0048, b"")
        d += _rec(0x0031, struct.pack("<I", 0))      # MODULEOFFSET: source at 0
        d += _rec(0x001E, struct.pack("<I", 0))      # HELPCONTEXT
        d += _rec(0x002C, struct.pack("<H", 0xFFFF)) # MODULECOOKIE
        d += _rec(0x0021 if m.kind == STANDARD else 0x0022, b"")  # MODULETYPE
        d += _rec(0x002B, b"")                       # module terminator
    d += _rec(0x0010, b"")                           # dir terminator
    return bytes(d)


def build_project_text(project_id: str, modules: list[Module]) -> bytes:
    rng = random.Random(project_id)  # deterministic output for a given project ID
    lines = [f'ID="{project_id}"']
    for m in modules:
        if m.kind == DOCUMENT:
            lines.append(f"Document={m.name}/&H00000000")
        elif m.kind == CLASS:
            lines.append(f"Class={m.name}")
        else:
            lines.append(f"Module={m.name}")
    lines += [
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        f'CMG="{_encrypt(struct.pack("<I", 0), project_id, rng)}"',
        f'DPB="{_encrypt(bytes([0]), project_id, rng)}"',
        f'GC="{_encrypt(bytes([0xFF]), project_id, rng)}"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    for m in modules:
        lines.append(f"{m.name}=0, 0, 0, 0, C")
    return ("\r\n".join(lines) + "\r\n").encode("cp1252")


def build_projectwm(modules: list[Module]) -> bytes:
    out = bytearray()
    for m in modules:
        out += _mb(m.name) + b"\x00" + _u16(m.name) + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)


# ============================================================================ compound file
@dataclass
class _Entry:
    name: str
    kind: int                    # 1 storage, 2 stream, 5 root
    data: bytes = b""
    children: list["_Entry"] | None = None
    sid: int = 0
    left: int = NOSTREAM
    right: int = NOSTREAM
    child: int = NOSTREAM
    start: int = ENDOFCHAIN
    size: int = 0


def _cfb_key(e: _Entry) -> tuple[int, str]:
    return (len(e.name), e.name.upper())


# Writes the MS-CFB layout in one pass (directory, FAT, mini stream); its
# branches are the format's cases, which a split would only scatter.
def write_cfb(tree: dict[str, Any]) -> bytes:  # noqa: C901
    """tree: {name: bytes | dict}; returns compound file bytes."""
    root = _Entry("Root Entry", 5, children=[])
    flat: list[_Entry] = [root]

    def build(parent: _Entry, node: dict[str, Any]) -> None:
        assert parent.children is not None  # build() only receives storage entries
        for name, value in node.items():
            if isinstance(value, dict):
                e = _Entry(name, 1, children=[])
                parent.children.append(e)
                flat.append(e)
                build(e, value)
            else:
                e = _Entry(name, 2, data=value, size=len(value))
                parent.children.append(e)
                flat.append(e)

    build(root, tree)
    for i, e in enumerate(flat):
        e.sid = i

    def link(entries: list[_Entry]) -> int:
        if not entries:
            return NOSTREAM
        entries = sorted(entries, key=_cfb_key)
        mid = len(entries) // 2
        node = entries[mid]
        node.left = link(entries[:mid])
        node.right = link(entries[mid + 1:])
        return node.sid

    for e in flat:
        if e.children is not None:
            e.child = link(e.children)

    SECTOR = 512
    fat: list[int] = []
    sectors: list[bytes] = []

    def alloc(data: bytes) -> int:
        count = max(1, -(-len(data) // SECTOR))
        first = len(sectors)
        for i in range(count):
            sectors.append(data[i * SECTOR:(i + 1) * SECTOR].ljust(SECTOR, b"\x00"))
            fat.append(first + i + 1 if i < count - 1 else ENDOFCHAIN)
        return first

    # Mini stream for small streams.
    mini = bytearray()
    minifat: list[int] = []
    for e in flat:
        if e.kind == 2 and e.size < 4096:
            if e.size == 0:
                e.start = ENDOFCHAIN
                continue
            count = -(-e.size // 64)
            first = len(minifat)
            for i in range(count):
                minifat.append(first + i + 1 if i < count - 1 else ENDOFCHAIN)
            e.start = first
            mini += e.data.ljust(count * 64, b"\x00")
    for e in flat:
        if e.kind == 2 and e.size >= 4096:
            e.start = alloc(e.data)
    if mini:
        root.start = alloc(bytes(mini))
        root.size = len(mini)
    minifat_start, n_minifat = ENDOFCHAIN, 0
    if minifat:
        mf = b"".join(struct.pack("<I", v) for v in minifat)
        minifat_start = alloc(mf.ljust(-(-len(mf) // SECTOR) * SECTOR, b"\xff"))
        n_minifat = -(-len(mf) // SECTOR)

    # Directory.
    dir_bytes = bytearray()
    for e in flat:
        name = e.name.encode("utf-16-le") + b"\x00\x00"
        entry = name.ljust(64, b"\x00")
        entry += struct.pack("<HBB", len(name), e.kind, 1)
        entry += struct.pack("<III", e.left, e.right, e.child)
        entry += b"\x00" * 16 + struct.pack("<I", 0) + b"\x00" * 16
        entry += struct.pack("<IQ", e.start if e.kind != 1 else 0, e.size if e.kind != 1 else 0)
        dir_bytes += entry
    while len(dir_bytes) % SECTOR:
        dir_bytes += b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM) \
            + b"\x00" * 36 + struct.pack("<IQ", 0, 0)
    dir_start = alloc(bytes(dir_bytes))

    # FAT sectors (they describe themselves too).
    n_fat = 1
    while True:
        total = len(sectors) + n_fat
        if n_fat * (SECTOR // 4) >= total:
            break
        n_fat += 1
    if n_fat > 109:
        raise ValueError("DIFAT not implemented")
    fat_start = len(sectors)
    fat += [FATSECT] * n_fat
    fat += [FREESECT] * (n_fat * (SECTOR // 4) - len(fat))
    fat_bytes = b"".join(struct.pack("<I", v) for v in fat)
    for i in range(n_fat):
        sectors.append(fat_bytes[i * SECTOR:(i + 1) * SECTOR])

    header = bytearray()
    header += bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 16
    header += struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6) + b"\x00" * 6
    header += struct.pack("<IIIII", 0, n_fat, dir_start, 0, 4096)
    header += struct.pack("<IIII", minifat_start, n_minifat, ENDOFCHAIN, 0)
    difat = [fat_start + i for i in range(n_fat)] + [FREESECT] * (109 - n_fat)
    header += b"".join(struct.pack("<I", v) for v in difat)
    assert len(header) == 512
    return bytes(header) + b"".join(sectors)


# ============================================================================ top level
def build_vba_project(modules: list[Module], project_id: str | None = None) -> bytes:
    project_id = project_id or "{" + str(uuid.uuid4()).upper() + "}"
    vba = {"_VBA_PROJECT": bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]),
           "dir": compress(build_dir("VBAProject", modules))}
    for m in modules:
        try:
            vba[m.name] = compress(m.source.encode("cp1252"))
        except ValueError as error:
            raise ValueError(f"module {m.name}: {error}; import it by hand as described in "
                             "INSTALLATION.md (Importing VBA into Excel)") from error
    return write_cfb({"PROJECT": build_project_text(project_id, modules),
                      "PROJECTwm": build_projectwm(modules),
                      "VBA": vba})
