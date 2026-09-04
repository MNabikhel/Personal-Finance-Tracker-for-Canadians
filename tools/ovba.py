"""Build a ``vbaProject.bin`` (MS-OVBA) from plain-text VBA sources.

Implements the pieces of [MS-OVBA] needed to author a project from scratch:
the compressed container format (section 2.4.1), the ``dir`` stream (2.3.4.2),
the ``PROJECT`` stream text (2.3.1) including the obfuscated protection-state
properties (2.4.3), ``PROJECTwm`` and ``_VBA_PROJECT``.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from . import cfb

CODE_PAGE = 1252
MBCS = "cp1252"

SIGNATURE = 0x01
MAX_CHUNK = 4096

STANDARD = "standard"
CLASS = "class"
DOCUMENT = "document"

# Attribute VB_Base values Excel writes for host document modules.
WORKBOOK_VB_BASE = "0{00020819-0000-0000-C000-000000000046}"
WORKSHEET_VB_BASE = "0{00020820-0000-0000-C000-000000000046}"

# The references Excel stores for a new workbook project - and only these.
# The VBA runtime and the Excel object library are host references that the
# project gets implicitly; a project that also lists them explicitly is not
# what Excel writes, and Excel treats a second copy of a library it already
# supplies as a name conflict.  Office resolves the two below by CLSID, so the
# embedded paths are only a hint.
DEFAULT_REFERENCES: List[Tuple[str, str]] = [
    (
        "stdole",
        "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
        "C:\\Windows\\System32\\stdole2.tlb#OLE Automation",
    ),
    (
        "Office",
        "*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.8#0#"
        "C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL#"
        "Microsoft Office 16.0 Object Library",
    ),
]


class OvbaError(Exception):
    pass


# ---------------------------------------------------------------------------
# Compression (MS-OVBA 2.4.1)
# ---------------------------------------------------------------------------


def _copy_token_help(difference: int) -> Tuple[int, int, int, int]:
    bit_count = max(_ceil_log2(difference), 4)
    length_mask = 0xFFFF >> bit_count
    offset_mask = (~length_mask) & 0xFFFF
    return length_mask, offset_mask, bit_count, length_mask + 3


def _ceil_log2(value: int) -> int:
    result = 0
    while (1 << result) < value:
        result += 1
    return result


def _compress_chunk_tokens(chunk: bytes) -> bytes:
    """Token-encode one decompressed chunk (literals plus copy tokens)."""
    out = bytearray()
    index: Dict[bytes, List[int]] = {}
    position = 0
    length = len(chunk)
    while position < length:
        flag_position = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if position >= length:
                break
            length_mask, _, bit_count, max_length = _copy_token_help(position)
            max_offset = 1 << bit_count
            best_length = 0
            best_offset = 0
            if position >= 1 and position + 2 < length:
                key = chunk[position : position + 3]
                lowest = position - max_offset
                for candidate in reversed(index.get(key, ())):
                    if candidate < lowest:
                        break
                    match = 0
                    while (
                        match < max_length
                        and position + match < length
                        and chunk[candidate + match] == chunk[position + match]
                    ):
                        match += 1
                    if match > best_length:
                        best_length = match
                        best_offset = position - candidate
                        if best_length == max_length:
                            break
            if best_length >= 3:
                token = ((best_offset - 1) << (16 - bit_count)) | (
                    (best_length - 3) & length_mask
                )
                out.extend(struct.pack("<H", token))
                flags |= 1 << bit
                for step in range(best_length):
                    spot = position + step
                    if spot + 2 < length:
                        index.setdefault(chunk[spot : spot + 3], []).append(spot)
                position += best_length
            else:
                out.append(chunk[position])
                if position + 2 < length:
                    index.setdefault(chunk[position : position + 3], []).append(position)
                position += 1
        out[flag_position] = flags
    return bytes(out)


def compress(data: bytes) -> bytes:
    """Compress a byte string into an MS-OVBA CompressedContainer."""
    out = bytearray([SIGNATURE])
    for start in range(0, len(data), MAX_CHUNK):
        chunk = data[start : start + MAX_CHUNK]
        tokens = _compress_chunk_tokens(chunk)
        # A raw chunk always carries exactly 4096 bytes, so it may only be used
        # for a full chunk; a short final chunk must be token-encoded.
        if len(chunk) == MAX_CHUNK and len(tokens) >= MAX_CHUNK:
            header = ((MAX_CHUNK + 2 - 3) & 0x0FFF) | 0x3000
            out.extend(struct.pack("<H", header))
            out.extend(chunk)
        else:
            header = ((len(tokens) + 2 - 3) & 0x0FFF) | 0x3000 | 0x8000
            out.extend(struct.pack("<H", header))
            out.extend(tokens)
    return bytes(out)


def decompress(data: bytes) -> bytes:
    """Inverse of :func:`compress`; used by the test-suite as a cross-check."""
    if not data or data[0] != SIGNATURE:
        raise OvbaError("bad compressed container signature")
    out = bytearray()
    position = 1
    while position < len(data):
        chunk_start = position
        (header,) = struct.unpack_from("<H", data, position)
        size = (header & 0x0FFF) + 3
        if (header >> 12) & 0x07 != 0b011:
            raise OvbaError("bad chunk signature")
        compressed = bool(header & 0x8000)
        chunk_end = min(len(data), chunk_start + size)
        position = chunk_start + 2
        chunk_output_start = len(out)
        if not compressed:
            out.extend(data[position : position + MAX_CHUNK])
            position += MAX_CHUNK
            continue
        while position < chunk_end:
            flags = data[position]
            position += 1
            for bit in range(8):
                if position >= chunk_end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[position])
                    position += 1
                    continue
                (token,) = struct.unpack_from("<H", data, position)
                position += 2
                length_mask, offset_mask, bit_count, _ = _copy_token_help(
                    len(out) - chunk_output_start
                )
                match_length = (token & length_mask) + 3
                match_offset = ((token & offset_mask) >> (16 - bit_count)) + 1
                source = len(out) - match_offset
                if source < chunk_output_start:
                    raise OvbaError("copy token points before chunk start")
                for step in range(match_length):
                    out.append(out[source + step])
    return bytes(out)


# ---------------------------------------------------------------------------
# PROJECT stream protection properties (MS-OVBA 2.4.3)
# ---------------------------------------------------------------------------


def project_key(project_id: str) -> int:
    """Checksum of the project CLSID string, braces included."""
    return sum(project_id.encode(MBCS)) & 0xFF


def encrypt(data: bytes, key: int, seed: int, ignored_byte: int = 0x07) -> str:
    """Obfuscate ``data`` into the hex form used by CMG/DPB/GC."""
    version_enc = seed ^ 0x02
    key_enc = seed ^ key
    out = bytearray([seed, version_enc, key_enc])

    unencrypted_1 = key
    encrypted_1 = key_enc
    encrypted_2 = version_enc

    def push(plain: int) -> None:
        nonlocal unencrypted_1, encrypted_1, encrypted_2
        cipher = plain ^ ((encrypted_2 + unencrypted_1) & 0xFF)
        out.append(cipher)
        encrypted_2 = encrypted_1
        encrypted_1 = cipher
        unencrypted_1 = plain

    for _ in range((seed & 0x06) // 2):
        push(ignored_byte)
    for byte in struct.pack("<I", len(data)):
        push(byte)
    for byte in data:
        push(byte)
    return "".join(f"{value:02X}" for value in out)


# ---------------------------------------------------------------------------
# Project model
# ---------------------------------------------------------------------------


@dataclass
class Module:
    name: str
    source: str
    kind: str = STANDARD

    @property
    def is_document(self) -> bool:
        return self.kind == DOCUMENT

    def normalised_source(self) -> str:
        text = self.source.replace("\r\n", "\n").replace("\r", "\n")
        if not text.endswith("\n"):
            text += "\n"
        return text.replace("\n", "\r\n")


@dataclass
class Project:
    name: str = "VBAProject"
    project_id: str = "{00000000-0000-0000-0000-000000000000}"
    modules: List[Module] = field(default_factory=list)
    references: List[Tuple[str, str]] = field(default_factory=list)
    sys_kind: int = 1  # Win32
    lcid: int = 0x0409
    description: str = ""

    def add(self, module: Module) -> None:
        self.modules.append(module)


def _record(record_id: int, payload: bytes) -> bytes:
    return struct.pack("<HI", record_id, len(payload)) + payload


def _sized_string(record_id: int, text: str, reserved: Optional[int] = None) -> bytes:
    encoded = text.encode(MBCS)
    out = struct.pack("<HI", record_id, len(encoded)) + encoded
    if reserved is not None:
        wide = text.encode("utf-16-le")
        out += struct.pack("<HI", reserved, len(wide)) + wide
    return out


def build_dir(project: Project) -> bytes:
    """Serialize the (uncompressed) ``dir`` stream."""
    out = bytearray()

    out += _record(0x0001, struct.pack("<I", project.sys_kind))
    out += _record(0x0002, struct.pack("<I", project.lcid))
    out += _record(0x0014, struct.pack("<I", project.lcid))
    out += _record(0x0003, struct.pack("<H", CODE_PAGE))
    out += _sized_string(0x0004, project.name)
    out += _sized_string(0x0005, project.description, reserved=0x0040)
    out += _sized_string(0x0006, "", reserved=0x003D)
    out += _record(0x0007, struct.pack("<I", 0))
    out += _record(0x0008, struct.pack("<I", 0))
    out += struct.pack("<HIIH", 0x0009, 0x00000004, 1, 0)
    out += _sized_string(0x000C, "", reserved=0x003C)

    for name, libid in project.references:
        out += _sized_string(0x0016, name, reserved=0x003E)
        encoded = libid.encode(MBCS)
        payload = struct.pack("<I", len(encoded)) + encoded + struct.pack("<IH", 0, 0)
        out += _record(0x000D, payload)

    out += _record(0x000F, struct.pack("<H", len(project.modules)))
    out += _record(0x0013, struct.pack("<H", 0xFFFF))

    for module in project.modules:
        out += _sized_string(0x0019, module.name)
        wide = module.name.encode("utf-16-le")
        out += struct.pack("<HI", 0x0047, len(wide)) + wide
        out += _sized_string(0x001A, module.name, reserved=0x0032)
        out += _sized_string(0x001C, "", reserved=0x0048)
        out += _record(0x0031, struct.pack("<I", 0))  # source starts at byte 0
        out += _record(0x001E, struct.pack("<I", 0))
        out += _record(0x002C, struct.pack("<H", 0xFFFF))
        # MODULETYPE: 0x0021 is a procedural module; 0x0022 is any module
        # with a class behind it - document, class or designer (2.3.4.2.3.2.8).
        out += _record(0x0021 if module.kind == STANDARD else 0x0022, b"")
        out += _record(0x002B, b"")  # MODULE terminator: Id plus reserved u32

    out += _record(0x0010, b"")  # dir terminator: Id plus reserved u32
    return bytes(out)


def build_project_stream(project: Project) -> bytes:
    """Serialize the root ``PROJECT`` stream text."""
    lines = [f'ID="{project.project_id}"']
    for module in project.modules:
        if module.kind == DOCUMENT:
            lines.append(f"Document={module.name}/&H00000000")
        elif module.kind == CLASS:
            lines.append(f"Class={module.name}")
        else:
            lines.append(f"Module={module.name}")
    lines.append(f'Name="{project.name}"')
    lines.append('HelpContextID="0"')
    lines.append('VersionCompatible32="393222000"')

    key = project_key(project.project_id)
    lines.append(f'CMG="{encrypt(struct.pack("<I", 0), key, 0x07)}"')
    lines.append(f'DPB="{encrypt(bytes([0x00]), key, 0x0E)}"')
    lines.append(f'GC="{encrypt(bytes([0xFF]), key, 0x15)}"')
    lines.append("")
    lines.append("[Host Extender Info]")
    lines.append("&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000")
    lines.append("")
    lines.append("[Workspace]")
    for module in project.modules:
        lines.append(f"{module.name}=0, 0, 0, 0, C")
    return ("\r\n".join(lines) + "\r\n").encode(MBCS)


def build_projectwm(project: Project) -> bytes:
    out = bytearray()
    for module in project.modules:
        out += module.name.encode(MBCS) + b"\x00"
        out += module.name.encode("utf-16-le") + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)


def build_vba_project_stream() -> bytes:
    """``_VBA_PROJECT`` header only.

    Version 0xFFFF is the interoperable value: it tells Office that no usable
    performance cache is present, so the project is recompiled from source.
    """
    return struct.pack("<HHBH", 0x61CC, 0xFFFF, 0x00, 0x0001)


def build(project: Project) -> bytes:
    """Return the bytes of a complete ``vbaProject.bin``."""
    if not project.modules:
        raise OvbaError("a VBA project needs at least one module")

    root = cfb.Storage("Root Entry")
    root.add_stream("PROJECT", build_project_stream(project))
    root.add_stream("PROJECTwm", build_projectwm(project))

    vba = root.add_storage("VBA")
    vba.add_stream("_VBA_PROJECT", build_vba_project_stream())
    vba.add_stream("dir", compress(build_dir(project)))
    for module in project.modules:
        source = module.normalised_source().encode(MBCS)
        vba.add_stream(module.name, compress(source))

    return cfb.serialize(root)


def module_header(name: str, kind: str = STANDARD, vb_base: str = "") -> str:
    """The ``Attribute`` preamble Excel writes at the top of a module."""
    if kind == STANDARD:
        return f'Attribute VB_Name = "{name}"\n'
    if kind == DOCUMENT:
        lines = [
            f'Attribute VB_Name = "{name}"',
            f'Attribute VB_Base = "{vb_base or WORKSHEET_VB_BASE}"',
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = True",
            "Attribute VB_Exposed = True",
            "Attribute VB_TemplateDerived = False",
            "Attribute VB_Customizable = True",
        ]
    else:
        lines = [
            "VERSION 1.0 CLASS",
            "BEGIN",
            "  MultiUse = -1  'True",
            "END",
            f'Attribute VB_Name = "{name}"',
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = False",
            "Attribute VB_Exposed = False",
        ]
    return "\n".join(lines) + "\n"
