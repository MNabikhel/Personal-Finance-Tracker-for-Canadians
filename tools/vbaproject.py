"""Assembles ``vba/*.bas`` / ``vba/*.cls`` into a ``vbaProject.bin``.

Excel expects a document module for every worksheet, keyed by the sheet's
``codeName``.  Those modules hold no code here, but leaving them out makes
Excel treat the project as damaged, so they are generated to match whatever
code names the workbook builder assigned.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Iterable, List

from . import ovba

VBA_DIR = Path(__file__).resolve().parent.parent / "vba"

PROJECT_NAME = "CanadianFinanceTracker"

# A fixed project id keeps builds byte-for-byte reproducible.  Office only uses
# it to derive the obfuscation key for the protection-state properties.
PROJECT_ID = "{4D1BE0C7-6F4A-4B27-9A65-0C1E2D3F5A80}"

NAME_ATTRIBUTE = re.compile(r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"', re.MULTILINE)

# Load order matters only for readability in the VBA editor, but a stable order
# keeps the generated binary reproducible.
MODULE_ORDER = [
    "modConst",
    "modUtil",
    "modParse",
    "modProfiles",
    "modAccounts",
    "modLedger",
    "modRules",
    "modTransfers",
    "modImport",
    "modHousehold",
    "modReport",
    "modSetup",
    "modUI",
]

SHEET_MODULE = """Attribute VB_Name = "{name}"
Attribute VB_Base = "{vb_base}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = True
Attribute VB_Customizable = True
Option Explicit

'== {sheet} ==
' Sheet module kept empty on purpose: all behaviour lives in the standard
' modules so it can be read and reviewed as plain text in the repository.
"""


def read_sources(directory: Path = VBA_DIR) -> Dict[str, str]:
    """Map module name -> source text for every file in ``vba/``."""
    sources: Dict[str, str] = {}
    for path in sorted(directory.iterdir()):
        if path.suffix not in (".bas", ".cls"):
            continue
        text = path.read_text(encoding="utf-8")
        match = NAME_ATTRIBUTE.search(text)
        if not match:
            raise ValueError(f"{path.name} has no Attribute VB_Name line")
        name = match.group(1)
        if name in sources:
            raise ValueError(f"duplicate module name {name!r}")
        sources[name] = text
    return sources


def build(sheet_code_names: Iterable[str],
          directory: Path = VBA_DIR) -> bytes:
    """Return the bytes of a ``vbaProject.bin`` for this workbook."""
    sources = read_sources(directory)

    if "ThisWorkbook" not in sources:
        raise ValueError("vba/ThisWorkbook.cls is required")

    project = ovba.Project(
        name=PROJECT_NAME,
        project_id=PROJECT_ID,
        references=[ovba.VBA_REFERENCE, ovba.EXCEL_REFERENCE]
        + ovba.DEFAULT_REFERENCES,
        description="Personal finance tracker for Canadian households.",
    )

    project.add(ovba.Module("ThisWorkbook", sources.pop("ThisWorkbook"),
                            ovba.DOCUMENT))

    for code_name in sheet_code_names:
        project.add(
            ovba.Module(
                code_name,
                SHEET_MODULE.format(name=code_name,
                                    vb_base=ovba.WORKSHEET_VB_BASE,
                                    sheet=code_name),
                ovba.DOCUMENT,
            )
        )

    ordered: List[str] = [n for n in MODULE_ORDER if n in sources]
    ordered += sorted(name for name in sources if name not in MODULE_ORDER)
    missing = [name for name in MODULE_ORDER if name not in sources]
    if missing:
        raise ValueError(f"expected modules are missing from vba/: {missing}")

    for name in ordered:
        kind = ovba.CLASS if sources[name].startswith("VERSION 1.0 CLASS") \
            else ovba.STANDARD
        project.add(ovba.Module(name, sources[name], kind))

    return ovba.build(project)
