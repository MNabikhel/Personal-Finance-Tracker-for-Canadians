"""A headless LibreOffice, used as an independent reader of what we build.

LibreOffice is a completely separate implementation of both OOXML and the
Basic dialect VBA is based on, so getting the same numbers out of it is real
evidence that the workbook is well formed - far better than asserting that our
own writer wrote what our own writer wrote.

One instance is started for the whole test session and shut down at exit.
"""

from __future__ import annotations

import atexit
import os
import shutil
import socket
import subprocess
import tempfile
import time
from typing import Optional

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")

_process: Optional[subprocess.Popen] = None
_context = None
_profile: Optional[str] = None


class Unavailable(Exception):
    """Raised when LibreOffice or its Python-UNO bridge is not installed."""


def _free_port() -> int:
    with socket.socket() as probe:
        probe.bind(("127.0.0.1", 0))
        return probe.getsockname()[1]


def _shutdown():
    global _process
    if _process and _process.poll() is None:
        _process.terminate()
        try:
            _process.wait(timeout=20)
        except subprocess.TimeoutExpired:
            _process.kill()
    if _profile:
        shutil.rmtree(_profile, ignore_errors=True)


def context():
    """Connect to (and if necessary start) the shared soffice instance."""
    global _process, _context, _profile
    if _context is not None:
        return _context

    try:
        import uno  # noqa: F401
    except ImportError as exc:
        raise Unavailable("the Python-UNO bridge is not installed") from exc
    if SOFFICE is None:
        raise Unavailable("soffice is not on PATH")

    import uno

    port = _free_port()
    _profile = tempfile.mkdtemp(prefix="cft-loffice-")
    _process = subprocess.Popen(
        [
            SOFFICE, "--headless", "--norestore", "--nologo", "--nodefault",
            "--nofirststartwizard", "--nolockcheck",
            f"-env:UserInstallation=file://{_profile}",
            f"--accept=socket,host=127.0.0.1,port={port};urp;",
        ],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    atexit.register(_shutdown)

    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    url = (f"uno:socket,host=127.0.0.1,port={port};urp;"
           "StarOffice.ComponentContext")

    deadline = time.monotonic() + 90
    last: Optional[Exception] = None
    while time.monotonic() < deadline:
        if _process.poll() is not None:
            raise Unavailable("soffice exited during start-up")
        try:
            _context = resolver.resolve(url)
            return _context
        except Exception as exc:  # the listener is not up yet
            last = exc
            time.sleep(0.4)
    raise Unavailable(f"soffice did not accept a connection: {last}")


def desktop():
    ctx = context()
    return ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx)


def _properties(**values):
    import uno
    from com.sun.star.beans import PropertyValue

    out = []
    for key, value in values.items():
        item = PropertyValue()
        item.Name = key
        item.Value = value
        out.append(item)
    return tuple(out)


def open_document(path: str):
    """Load a file, allowing its macros to be inspected (never run on load)."""
    import uno

    url = uno.systemPathToFileUrl(os.path.abspath(path))
    # MacroExecutionMode 4 (ALWAYS_EXECUTE_NO_WARN) makes the Basic libraries
    # visible; nothing is auto-run because the document has no Auto_Open.
    return desktop().loadComponentFromURL(
        url, "_blank", 0, _properties(Hidden=True, MacroExecutionMode=4))


def blank_spreadsheet():
    return desktop().loadComponentFromURL(
        "private:factory/scalc", "_blank", 0, _properties(Hidden=True))


def script_provider(document):
    ctx = context()
    factory = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.script.provider.MasterScriptProviderFactory", ctx)
    return factory.createScriptProvider(document)
