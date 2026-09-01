"""Minimal writer for the Compound File Binary format (MS-CFB, version 3).

Only the subset needed to emit a ``vbaProject.bin`` is implemented: storages,
streams, the FAT/mini-FAT allocation chains and a directory tree.  The output is
deterministic (all timestamps are zero) so builds are reproducible.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field
from typing import Iterable, List, Union

SECTOR_SIZE = 512
MINI_SECTOR_SIZE = 64
MINI_STREAM_CUTOFF = 4096

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
NOSTREAM = 0xFFFFFFFF

_TYPE_STORAGE = 1
_TYPE_STREAM = 2
_TYPE_ROOT = 5

_COLOR_BLACK = 1

MAX_ENTRY_NAME = 31


@dataclass
class Stream:
    name: str
    data: bytes


@dataclass
class Storage:
    name: str
    entries: List[Union["Storage", Stream]] = field(default_factory=list)

    def add_stream(self, name: str, data: bytes) -> Stream:
        stream = Stream(name, data)
        self.entries.append(stream)
        return stream

    def add_storage(self, name: str) -> "Storage":
        storage = Storage(name)
        self.entries.append(storage)
        return storage


class CfbError(Exception):
    pass


def _entry_sort_key(name: str) -> tuple:
    """CFB directory ordering: shorter names first, then case-insensitive."""
    return (len(name), name.upper())


@dataclass
class _DirEntry:
    name: str
    entry_type: int
    left: int = NOSTREAM
    right: int = NOSTREAM
    child: int = NOSTREAM
    start_sector: int = ENDOFCHAIN
    size: int = 0
    clsid: bytes = b"\x00" * 16

    def pack(self) -> bytes:
        encoded = self.name.encode("utf-16-le")
        if len(encoded) > 62:
            raise CfbError(f"directory entry name too long: {self.name!r}")
        name_field = encoded + b"\x00\x00"
        name_field += b"\x00" * (64 - len(name_field))
        return struct.pack(
            "<64sHBBIII16sIQQIQ",
            name_field,
            len(encoded) + 2,
            self.entry_type,
            _COLOR_BLACK,
            self.left,
            self.right,
            self.child,
            self.clsid,
            0,  # state bits
            0,  # creation time
            0,  # modified time
            self.start_sector,
            self.size,
        )


def _build_tree(indexes: List[int], entries: List[_DirEntry]) -> int:
    """Link a sorted list of sibling entries into a balanced binary tree."""
    if not indexes:
        return NOSTREAM
    mid = len(indexes) // 2
    root = indexes[mid]
    entries[root].left = _build_tree(indexes[:mid], entries)
    entries[root].right = _build_tree(indexes[mid + 1 :], entries)
    return root


def _pad_to(data: bytes, block: int) -> bytes:
    remainder = len(data) % block
    if remainder:
        data += b"\x00" * (block - remainder)
    return data


def _chain(fat: List[int], sectors: Iterable[int]) -> None:
    sectors = list(sectors)
    for current, following in zip(sectors, sectors[1:]):
        fat[current] = following
    if sectors:
        fat[sectors[-1]] = ENDOFCHAIN


def serialize(root: Storage) -> bytes:
    """Serialize a storage tree into a CFB (version 3) byte string."""
    entries: List[_DirEntry] = [
        _DirEntry(name="Root Entry", entry_type=_TYPE_ROOT, size=0)
    ]
    # (directory index, payload) for every stream, in directory order.
    stream_payloads: List[tuple] = []

    def walk(storage: Storage, dir_index: int) -> None:
        child_indexes = []
        sub_storages = []
        for item in sorted(storage.entries, key=lambda e: _entry_sort_key(e.name)):
            if len(item.name) > MAX_ENTRY_NAME:
                raise CfbError(f"name exceeds {MAX_ENTRY_NAME} chars: {item.name!r}")
            index = len(entries)
            if isinstance(item, Stream):
                entries.append(
                    _DirEntry(
                        name=item.name,
                        entry_type=_TYPE_STREAM,
                        size=len(item.data),
                    )
                )
                stream_payloads.append((index, item.data))
            else:
                entries.append(_DirEntry(name=item.name, entry_type=_TYPE_STORAGE))
                sub_storages.append((item, index))
            child_indexes.append(index)

        entries[dir_index].child = _build_tree(child_indexes, entries)
        # Recurse only after every sibling exists, keeping indexes grouped.
        for sub_storage, sub_index in sub_storages:
            walk(sub_storage, sub_index)

    walk(root, 0)

    big_streams = [(i, d) for i, d in stream_payloads if len(d) >= MINI_STREAM_CUTOFF]
    small_streams = [
        (i, d) for i, d in stream_payloads if 0 < len(d) < MINI_STREAM_CUTOFF
    ]

    # Mini stream: small stream payloads concatenated on 64-byte boundaries.
    mini_stream = bytearray()
    mini_fat: List[int] = []
    for index, data in small_streams:
        first_mini_sector = len(mini_stream) // MINI_SECTOR_SIZE
        padded = _pad_to(data, MINI_SECTOR_SIZE)
        count = len(padded) // MINI_SECTOR_SIZE
        mini_stream.extend(padded)
        for offset in range(count):
            mini_fat.append(first_mini_sector + offset + 1)
        mini_fat[-1] = ENDOFCHAIN
        entries[index].start_sector = first_mini_sector

    if mini_fat:
        slots_per_sector = SECTOR_SIZE // 4
        padded_slots = -(-len(mini_fat) // slots_per_sector) * slots_per_sector
        mini_fat_bytes = b"".join(
            struct.pack("<I", value)
            for value in mini_fat + [FREESECT] * (padded_slots - len(mini_fat))
        )
    else:
        mini_fat_bytes = b""

    directory_bytes_len = len(entries) * 128
    directory_sector_count = -(-directory_bytes_len // SECTOR_SIZE)

    def sector_count(byte_length: int) -> int:
        return -(-byte_length // SECTOR_SIZE)

    big_sector_counts = [sector_count(len(d)) for _, d in big_streams]
    mini_stream_sectors = sector_count(len(mini_stream))
    mini_fat_sectors = sector_count(len(mini_fat_bytes))

    non_fat_sectors = (
        sum(big_sector_counts)
        + mini_stream_sectors
        + mini_fat_sectors
        + directory_sector_count
    )

    fat_sectors = 1
    while True:
        total = non_fat_sectors + fat_sectors
        needed = max(1, sector_count(total * 4))
        if needed == fat_sectors:
            break
        fat_sectors = needed
    if fat_sectors > 109:
        raise CfbError("file too large for this writer (DIFAT sectors required)")

    total_sectors = non_fat_sectors + fat_sectors
    fat = [FREESECT] * (fat_sectors * SECTOR_SIZE // 4)

    next_sector = 0

    def allocate(count: int) -> List[int]:
        nonlocal next_sector
        allocated = list(range(next_sector, next_sector + count))
        next_sector += count
        return allocated

    payload = bytearray()

    for (index, data), count in zip(big_streams, big_sector_counts):
        sectors = allocate(count)
        entries[index].start_sector = sectors[0]
        _chain(fat, sectors)
        payload.extend(_pad_to(data, SECTOR_SIZE))

    if mini_stream:
        sectors = allocate(mini_stream_sectors)
        entries[0].start_sector = sectors[0]
        entries[0].size = len(mini_stream)
        _chain(fat, sectors)
        payload.extend(_pad_to(bytes(mini_stream), SECTOR_SIZE))
    else:
        entries[0].start_sector = ENDOFCHAIN

    if mini_fat_bytes:
        mini_fat_sector_ids = allocate(mini_fat_sectors)
        _chain(fat, mini_fat_sector_ids)
        payload.extend(mini_fat_bytes)
        first_mini_fat_sector = mini_fat_sector_ids[0]
    else:
        first_mini_fat_sector = ENDOFCHAIN

    directory_sector_ids = allocate(directory_sector_count)
    _chain(fat, directory_sector_ids)
    directory_bytes = b"".join(entry.pack() for entry in entries)
    # Unused directory slots must be zero-filled empty entries.
    payload.extend(_pad_to(directory_bytes, SECTOR_SIZE))

    fat_sector_ids = allocate(fat_sectors)
    for sector in fat_sector_ids:
        fat[sector] = FATSECT

    if next_sector != total_sectors:
        raise CfbError("sector accounting mismatch")

    fat_bytes = b"".join(struct.pack("<I", value) for value in fat)
    payload.extend(fat_bytes)

    difat = list(fat_sector_ids) + [FREESECT] * (109 - len(fat_sector_ids))

    header = struct.pack(
        "<8s16sHHHHH6sIIIIIIIII",
        b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1",
        b"\x00" * 16,
        0x003E,  # minor version
        0x0003,  # major version (512-byte sectors)
        0xFFFE,  # little-endian byte order
        9,  # sector shift (2**9 == 512)
        6,  # mini sector shift (2**6 == 64)
        b"\x00" * 6,  # reserved
        0,  # number of directory sectors (unused in v3)
        fat_sectors,
        directory_sector_ids[0],
        0,  # transaction signature
        MINI_STREAM_CUTOFF,
        first_mini_fat_sector,
        mini_fat_sectors,
        ENDOFCHAIN,  # first DIFAT sector
        0,  # number of DIFAT sectors
    )
    header += b"".join(struct.pack("<I", value) for value in difat)
    if len(header) != SECTOR_SIZE:
        raise CfbError(f"bad header length {len(header)}")

    return header + bytes(payload)
