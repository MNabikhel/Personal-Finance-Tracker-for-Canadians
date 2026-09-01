"""Runs the repository's VBA source inside LibreOffice Basic.

The point is to execute the shipped module text rather than a Python
re-implementation of it, so the date/amount/CSV parsing, the merchant clean-up
and the duplicate keys are all exercised for real.

Two accommodations are needed, neither of which changes any statement:

* ``Option VBASupport 1`` is prepended, which is how LibreOffice is told to
  provide the VBA runtime (``InStrRev``, ``Like``, the ``vb*`` constants);
* the ``Attribute VB_Name`` lines are dropped, since they are a container-level
  detail of the .bas format rather than code.

Only the modules that do not need Excel's object model can be driven this way.
"""

from __future__ import annotations

import os
import re
import tempfile
from pathlib import Path
from typing import Iterable, List, Sequence

from . import libreoffice

VBA_DIR = Path(__file__).resolve().parent.parent / "vba"

ATTRIBUTE_LINE = re.compile(r"^Attribute .*$", re.MULTILINE)

# Everything the parsing, clean-up and rule matching code needs, in dependency
# order.  None of these modules touches Excel's object model.
PARSING_MODULES = ["modConst", "modUtil", "modParse", "modRules", "modImport"]

# The fixed part of the probe: it owns the output file so that a failure inside
# the test's own code is still reported instead of vanishing.
HARNESS = """Option VBASupport 1

Private mOut As Integer

Sub Probe()
    On Error GoTo Fail
    mOut = FreeFile
    Open "{output}" For Output As #mOut
    Body.Run
    Close #mOut
    Exit Sub
Fail:
    Print #mOut, "{marker}" & Err & ": " & Error$
    Close #mOut
End Sub

' Writes one tab separated line; every test result goes through here.
Sub Emit(ParamArray values() As Variant)
    Dim i As Long, line As String
    For i = LBound(values) To UBound(values)
        If i > LBound(values) Then line = line & Chr$(9)
        line = line & CStr(values(i))
    Next i
    Print #mOut, line
End Sub

' Booleans and numbers print differently across locales and dialects, so the
' comparisons the tests make are normalised here.
Function Flag(ByVal value As Boolean) As String
    If value Then Flag = "yes" Else Flag = "no"
End Function

Function Money(ByVal value As Double) As String
    Money = Format$(value, "0.00")
End Function

Function Stamp(ByVal value As Date) As String
    Stamp = Format$(value, "yyyy-mm-dd")
End Function
"""

ERROR_MARKER = "!basic-error "


class BasicError(AssertionError):
    """A compile or runtime failure inside the Basic code under test."""


def module_source(name: str) -> str:
    if not name.endswith((".bas", ".cls")):
        name += ".bas"
    text = (VBA_DIR / name).read_text(encoding="utf-8")
    return "Option VBASupport 1\n" + ATTRIBUTE_LINE.sub("", text)


def run(body: str, modules: Iterable[str] = PARSING_MODULES) -> str:
    """Load ``modules``, run ``Sub Run`` from ``body``, return what it emitted.

    ``body`` must define ``Sub Run`` and report results by calling ``Emit``.
    """
    handle, output = tempfile.mkstemp(prefix="cft-basic-", suffix=".txt")
    os.close(handle)
    os.unlink(output)

    document = libreoffice.blank_spreadsheet()
    try:
        libraries = document.BasicLibraries
        if not libraries.hasByName("Standard"):
            libraries.createLibrary("Standard")
        standard = libraries.getByName("Standard")

        for name in modules:
            _put(standard, name, module_source(name))
        _put(standard, "Body", "Option VBASupport 1\n\n" + body)
        _put(standard, "Probe",
             HARNESS.format(output=output, marker=ERROR_MARKER))

        provider = libreoffice.script_provider(document)
        script = provider.getScript(
            "vnd.sun.star.script:Standard.Probe.Probe"
            "?language=Basic&location=document")
        script.invoke((), (), ())
    finally:
        document.close(True)

    if not os.path.exists(output):
        raise BasicError(
            "the probe wrote nothing at all, which means the harness itself "
            "failed to compile")
    try:
        with open(output, encoding="utf-8", errors="replace") as stream:
            text = stream.read()
    finally:
        os.unlink(output)

    if ERROR_MARKER in text:
        raise BasicError(text[text.index(ERROR_MARKER):].strip())
    return text


def rows(text: str) -> List[List[str]]:
    return [line.split("\t") for line in text.splitlines() if line]


def basic_string(text: str) -> str:
    """A Basic string literal for ``text``, including any awkward characters."""
    out: List[str] = []
    literal = ""
    for char in text:
        if char in ('"', "\r", "\n", "\t") or ord(char) > 126 or ord(char) < 32:
            if literal:
                out.append('"' + literal + '"')
                literal = ""
            out.append(f"ChrW$({ord(char)})")
        else:
            literal += char
    if literal or not out:
        out.append('"' + literal + '"')
    return " & ".join(out)


def basic_array(values: Sequence[str]) -> str:
    return "Array(" + ", ".join(basic_string(value) for value in values) + ")"


def _put(library, name: str, code: str) -> None:
    if library.hasByName(name):
        library.replaceByName(name, code)
    else:
        library.insertByName(name, code)
