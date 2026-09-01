"""Conformance tests for the MS-CFB writer and the MS-OVBA project builder.

The compression and encryption vectors are the worked examples published in
[MS-OVBA] sections 3.2 and 2.3.1.15-2.3.1.17, so a pass here means the output
matches Microsoft's own reference data.
"""

from __future__ import annotations

import io
import os
import struct
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tools import cfb, ovba  # noqa: E402


def unhex(text: str) -> bytes:
    return bytes.fromhex(text.replace(" ", ""))


NO_COMPRESSION = (
    "abcdefghijklmnopqrstuv.",
    "01 19 B0 00 61 62 63 64 65 66 67 68 00 69 6A 6B 6C 6D 6E 6F 70 00 71 72 "
    "73 74 75 76 2E",
)

NORMAL_COMPRESSION = (
    "#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa",
    "01 2F B0 00 23 61 61 61 62 63 64 65 82 66 00 70 61 67 68 69 6A 01 38 08 "
    "61 6B 6C 00 20 6D 6E 6F 70 06 71 02 70 04 00 72 73 74 75 76 10 77 78 79 "
    "7A 00 2C",
)

MAXIMUM_COMPRESSION = ("a" * 73, "01 03 B0 02 61 45 00")


class CompressionTests(unittest.TestCase):
    def test_decompresses_spec_vectors(self):
        for text, compressed in (
            NO_COMPRESSION,
            NORMAL_COMPRESSION,
            MAXIMUM_COMPRESSION,
        ):
            self.assertEqual(
                ovba.decompress(unhex(compressed)), text.encode("ascii"), text[:20]
            )

    def test_matches_spec_compressed_output(self):
        for text, compressed in (
            NO_COMPRESSION,
            NORMAL_COMPRESSION,
            MAXIMUM_COMPRESSION,
        ):
            self.assertEqual(
                ovba.compress(text.encode("ascii")).hex(),
                unhex(compressed).hex(),
                text[:20],
            )

    def test_round_trips_awkward_lengths(self):
        # A short final chunk must be token-encoded; a raw chunk would pad the
        # data out to 4096 bytes with nulls.
        payload = bytes((index * 7 + (index >> 3)) % 251 for index in range(20000))
        for length in (1, 8, 9, 63, 4093, 4094, 4095, 4096, 4097, 8192, 8193, 20000):
            data = payload[:length]
            self.assertEqual(ovba.decompress(ovba.compress(data)), data, length)

    def test_round_trips_repetitive_and_random_data(self):
        import random

        random.seed(20260901)
        cases = [
            b"",
            b"a",
            b"Option Explicit\r\n" * 900,
            bytes(random.randrange(256) for _ in range(9000)),
            (b"Sub Foo()\r\n    Debug.Print 1\r\nEnd Sub\r\n" * 400),
        ]
        for data in cases:
            self.assertEqual(ovba.decompress(ovba.compress(data)), data)

    def test_cross_checks_against_oletools(self):
        try:
            from oletools.olevba import decompress_stream
        except ImportError:  # pragma: no cover - optional dependency
            self.skipTest("oletools not installed")
        data = ("Attribute VB_Name = \"modX\"\r\n" + "Debug.Print 1\r\n" * 500).encode()
        self.assertEqual(bytes(decompress_stream(bytearray(ovba.compress(data)))), data)


class EncryptionTests(unittest.TestCase):
    """[MS-OVBA] 2.3.1.15-2.3.1.17 worked example, ProjKey 0xDF."""

    PROJECT_ID = "{917DED54-440B-4FD1-A5C1-74ACF261E600}"

    def test_project_key_checksum(self):
        self.assertEqual(ovba.project_key(self.PROJECT_ID), 0xDF)

    def test_protection_state_vectors(self):
        key = ovba.project_key(self.PROJECT_ID)
        self.assertEqual(
            ovba.encrypt(b"\x00\x00\x00\x00", key, 0x07),
            "0705D8E3D8EDDBF1DBF1DBF1DBF1",
        )
        self.assertEqual(ovba.encrypt(b"\x00", key, 0x0E), "0E0CD1ECDFF4E7F5E7F5E7")
        self.assertEqual(ovba.encrypt(b"\xff", key, 0x15), "1517CAF1D6F9D7F9D706")


class CfbTests(unittest.TestCase):
    def _sample(self):
        root = cfb.Storage("Root Entry")
        root.add_stream("PROJECT", b"ID=\"{}\"\r\n" * 4)
        storage = root.add_storage("VBA")
        storage.add_stream("dir", b"\x01" + b"payload" * 100)
        storage.add_stream("big", bytes(range(256)) * 40)  # forces FAT sectors
        storage.add_stream("small", b"tiny")
        return root

    def test_readable_by_olefile(self):
        import olefile

        blob = cfb.serialize(self._sample())
        self.assertEqual(len(blob) % cfb.SECTOR_SIZE, 0)
        with olefile.OleFileIO(io.BytesIO(blob)) as ole:
            paths = {"/".join(entry) for entry in ole.listdir()}
            self.assertEqual(
                paths, {"PROJECT", "VBA/dir", "VBA/big", "VBA/small"}
            )
            self.assertEqual(ole.openstream("VBA/small").read(), b"tiny")
            self.assertEqual(
                ole.openstream("VBA/big").read(), bytes(range(256)) * 40
            )

    def test_deterministic(self):
        self.assertEqual(
            cfb.serialize(self._sample()), cfb.serialize(self._sample())
        )


class ProjectTests(unittest.TestCase):
    def _project(self):
        project = ovba.Project(
            name="TestProject",
            project_id="{917DED54-440B-4FD1-A5C1-74ACF261E600}",
            references=list(ovba.DEFAULT_REFERENCES),
        )
        project.add(
            ovba.Module(
                "ThisWorkbook",
                ovba.module_header(
                    "ThisWorkbook", ovba.DOCUMENT, ovba.WORKBOOK_VB_BASE
                )
                + "Option Explicit\n\nPrivate Sub Workbook_Open()\n"
                "    MsgBox \"hello\"\nEnd Sub\n",
                ovba.DOCUMENT,
            )
        )
        project.add(
            ovba.Module(
                "modTest",
                ovba.module_header("modTest")
                + "Option Explicit\n\n" + "'padding\n" * 900 + "Sub Go()\nEnd Sub\n",
            )
        )
        project.add(
            ovba.Module(
                "clsThing",
                ovba.module_header("clsThing", ovba.CLASS)
                + "Option Explicit\n\nPublic Value As Long\n",
                ovba.CLASS,
            )
        )
        return project

    def test_olevba_extracts_modules(self):
        try:
            from oletools.olevba import VBA_Parser
        except ImportError:  # pragma: no cover - optional dependency
            self.skipTest("oletools not installed")

        blob = ovba.build(self._project())
        parser = VBA_Parser("vbaProject.bin", data=blob)
        try:
            self.assertTrue(parser.detect_vba_macros())
            found = {
                name: code for _, _, name, code in parser.extract_macros()
            }
        finally:
            parser.close()
        # olevba appends .cls to document/class modules and .bas to procedural
        # ones, so the extensions also confirm the MODULETYPE records: a class
        # module is 0x0022 like a document module, not 0x0021.
        self.assertEqual(set(found), {"ThisWorkbook.cls", "modTest.bas", "clsThing.cls"})
        found = {name.rsplit(".", 1)[0]: code for name, code in found.items()}
        self.assertIn("Private Sub Workbook_Open()", found["ThisWorkbook"])
        self.assertIn("Sub Go()", found["modTest"])
        self.assertIn("Public Value As Long", found["clsThing"])

    def test_dir_stream_round_trips(self):
        project = self._project()
        raw = ovba.build_dir(project)
        self.assertEqual(ovba.decompress(ovba.compress(raw)), raw)
        self.assertTrue(raw.startswith(b"\x01\x00\x04\x00\x00\x00"))

    def test_module_types_follow_the_kind_of_module(self):
        # MS-OVBA 2.3.4.2.3.2.8: 0x0021 is a procedural module, 0x0022 any
        # module with a class behind it - documents and class modules alike.
        # Excel refuses a project whose PROJECT stream says Class= while the
        # dir stream says procedural.
        raw = ovba.build_dir(self._project())
        types = {}
        position = 0
        current = None
        while position + 6 <= len(raw):
            record_id, size = struct.unpack_from("<HI", raw, position)
            if record_id == 0x0009:          # PROJECTVERSION: fixed 12 bytes
                position += 12
                continue
            payload = raw[position + 6:position + 6 + size]
            if record_id == 0x0019:          # MODULENAME
                current = payload.decode("cp1252")
            elif record_id in (0x0021, 0x0022):
                types[current] = record_id
            position += 6 + size
        self.assertEqual(types, {"ThisWorkbook": 0x0022, "modTest": 0x0021,
                                 "clsThing": 0x0022})

    def test_only_the_references_excel_stores_are_written(self):
        # The VBA runtime and the Excel library are implicit host references.
        # Excel's own files list stdole and Office only; listing the host's
        # libraries again reads to Excel as a name conflict.
        names = [name for name, _ in ovba.DEFAULT_REFERENCES]
        self.assertEqual(names, ["stdole", "Office"])
        raw = ovba.build_dir(self._project())
        self.assertNotIn(b"Visual Basic For Applications", raw)
        self.assertNotIn(b"Microsoft Excel", raw)
        self.assertIn(b"OLE Automation", raw)

    def test_project_stream_declares_modules(self):
        text = ovba.build_project_stream(self._project()).decode("cp1252")
        self.assertIn("Document=ThisWorkbook/&H00000000", text)
        self.assertIn("Module=modTest", text)
        self.assertIn('Name="TestProject"', text)
        self.assertIn("[Workspace]", text)
        self.assertRegex(text, r'CMG="[0-9A-F]+"')
        self.assertRegex(text, r'DPB="[0-9A-F]+"')
        self.assertRegex(text, r'GC="[0-9A-F]+"')


if __name__ == "__main__":
    unittest.main()
