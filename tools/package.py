"""Serialises an openpyxl workbook as a reproducible package - plain or macro-enabled.

openpyxl can only preserve a ``vbaProject.bin`` that it read from an existing
file, so the workbook is saved as a normal package and the three things that
make a package macro-enabled are patched in afterwards:

1. ``xl/vbaProject.bin`` itself;
2. a content type for it, plus the macro-enabled content type on the workbook
   part - this is what makes Excel offer to enable macros;
3. a relationship from the workbook part to the binary.

The plain ``.xlsx`` edition goes through the same re-packing without those
patches, so that both editions are written with fixed timestamps and two
builds of the same source produce identical bytes.
"""

from __future__ import annotations

import io
import re
import zipfile
from typing import Dict, List, Optional, Tuple

from openpyxl.workbook import Workbook

CONTENT_TYPES = "[Content_Types].xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
CORE_PROPERTIES = "docProps/core.xml"
VBA_PART = "xl/vbaProject.bin"

CREATED = re.compile(r"(<dcterms:created\b[^>]*>)([^<]*)(</dcterms:created>)")
MODIFIED = re.compile(r"(<dcterms:modified\b[^>]*>)([^<]*)(</dcterms:modified>)")

SHEET_MAIN = ("application/vnd.openxmlformats-officedocument."
              "spreadsheetml.sheet.main+xml")
MACRO_MAIN = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
VBA_CONTENT_TYPE = "application/vnd.ms-office.vbaProject"
VBA_REL_TYPE = ("http://schemas.microsoft.com/office/2006/relationships/"
                "vbaProject")

FIXED_DATE = (2026, 1, 1, 0, 0, 0)


class PackageError(Exception):
    pass


def _patch_content_types(xml: str) -> str:
    if SHEET_MAIN not in xml:
        raise PackageError("workbook content type override not found")
    xml = xml.replace(SHEET_MAIN, MACRO_MAIN)
    if 'Extension="bin"' not in xml:
        default = f'<Default Extension="bin" ContentType="{VBA_CONTENT_TYPE}"/>'
        xml = re.sub(r"(<Types\b[^>]*>)", r"\1" + default, xml, count=1)
    return xml


def _patch_workbook_rels(xml: str) -> str:
    if VBA_REL_TYPE in xml:
        return xml
    used = {int(match) for match in re.findall(r'Id="rId(\d+)"', xml)}
    next_id = max(used, default=0) + 1
    relationship = (f'<Relationship Id="rId{next_id}" Type="{VBA_REL_TYPE}" '
                    f'Target="vbaProject.bin"/>')
    if "</Relationships>" not in xml:
        raise PackageError("workbook relationships part is malformed")
    return xml.replace("</Relationships>", relationship + "</Relationships>")


def _patch_core_properties(xml: str) -> str:
    """Date the file from its own creation rather than from the clock.

    openpyxl stamps dcterms:modified with the current time as it saves, which
    would be the one thing in the package that changed between two builds of
    identical source.  Nothing has happened to the file since it was written,
    so the two timestamps should agree anyway.
    """
    created = CREATED.search(xml)
    if created is None:
        return xml
    return MODIFIED.sub(lambda match: match.group(1) + created.group(2)
                        + match.group(3), xml)


def _repack(workbook: Workbook, vba_project: Optional[bytes]) -> bytes:
    plain = io.BytesIO()
    workbook.save(plain)
    plain.seek(0)

    parts: List[Tuple[str, bytes]] = []
    with zipfile.ZipFile(plain) as source:
        for name in source.namelist():
            data = source.read(name)
            if name == CONTENT_TYPES and vba_project is not None:
                data = _patch_content_types(data.decode("utf-8")).encode("utf-8")
            elif name == WORKBOOK_RELS and vba_project is not None:
                data = _patch_workbook_rels(data.decode("utf-8")).encode("utf-8")
            elif name == CORE_PROPERTIES:
                data = _patch_core_properties(data.decode("utf-8")).encode("utf-8")
            parts.append((name, data))

    if not any(name == WORKBOOK_RELS for name, _ in parts):
        raise PackageError(f"{WORKBOOK_RELS} is missing from the package")
    if vba_project is not None:
        parts.append((VBA_PART, vba_project))

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in parts:
            info = zipfile.ZipInfo(name, date_time=FIXED_DATE)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            target.writestr(info, data)
    return out.getvalue()


def to_xlsm(workbook: Workbook, vba_project: bytes) -> bytes:
    """Serialise ``workbook`` as a macro-enabled package."""
    return _repack(workbook, vba_project)


def to_xlsx(workbook: Workbook) -> bytes:
    """Serialise ``workbook`` as a plain package, dated the same way."""
    return _repack(workbook, None)


def describe(package: bytes) -> Dict[str, object]:
    """Summary of the macro-related parts, used by the tests."""
    with zipfile.ZipFile(io.BytesIO(package)) as archive:
        names = archive.namelist()
        content_types = archive.read(CONTENT_TYPES).decode("utf-8")
        rels = archive.read(WORKBOOK_RELS).decode("utf-8")
        vba = archive.read(VBA_PART) if VBA_PART in names else b""
    return {
        "names": names,
        "macro_content_type": MACRO_MAIN in content_types,
        "sheet_content_type": SHEET_MAIN in content_types,
        "bin_default": VBA_CONTENT_TYPE in content_types,
        "vba_relationship": VBA_REL_TYPE in rels,
        "vba_project": vba,
    }
