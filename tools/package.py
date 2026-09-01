"""Turns an openpyxl workbook into a macro-enabled .xlsm with a VBA project.

openpyxl can only preserve a ``vbaProject.bin`` that it read from an existing
file, so the workbook is saved as a normal package and the three things that
make a package macro-enabled are patched in afterwards:

1. ``xl/vbaProject.bin`` itself;
2. a content type for it, plus the macro-enabled content type on the workbook
   part - this is what makes Excel offer to enable macros;
3. a relationship from the workbook part to the binary.

Everything is written with fixed timestamps so two builds of the same source
produce identical bytes.
"""

from __future__ import annotations

import io
import re
import zipfile
from typing import Dict, List, Tuple

from openpyxl.workbook import Workbook

CONTENT_TYPES = "[Content_Types].xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
VBA_PART = "xl/vbaProject.bin"

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


def to_xlsm(workbook: Workbook, vba_project: bytes) -> bytes:
    """Serialise ``workbook`` as a macro-enabled package."""
    plain = io.BytesIO()
    workbook.save(plain)
    plain.seek(0)

    parts: List[Tuple[str, bytes]] = []
    with zipfile.ZipFile(plain) as source:
        for name in source.namelist():
            data = source.read(name)
            if name == CONTENT_TYPES:
                data = _patch_content_types(data.decode("utf-8")).encode("utf-8")
            elif name == WORKBOOK_RELS:
                data = _patch_workbook_rels(data.decode("utf-8")).encode("utf-8")
            parts.append((name, data))

    if not any(name == WORKBOOK_RELS for name, _ in parts):
        raise PackageError(f"{WORKBOOK_RELS} is missing from the package")
    parts.append((VBA_PART, vba_project))

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in parts:
            info = zipfile.ZipInfo(name, date_time=FIXED_DATE)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            target.writestr(info, data)
    return out.getvalue()


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
