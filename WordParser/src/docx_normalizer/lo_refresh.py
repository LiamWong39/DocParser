"""Update all indexes (TOC) in a .docx with headless LibreOffice.

Runs under a Python that can `import uno` (usually the system python3), NOT the
project venv, so it must not import anything from docx_normalizer.

    python3 lo_refresh.py in.docx out.docx [soffice]
"""

import os
import shutil
import socket
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def free_port():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def main(src, dst, soffice="soffice"):
    profile = tempfile.mkdtemp(prefix="docxnorm-lo-")
    port = free_port()
    conn = f"socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
    proc = subprocess.Popen(
        [soffice, f"-env:UserInstallation={uno.systemPathToFileUrl(profile)}", "--headless", "--invisible",
         "--nologo", "--norestore", "--nodefault", f"--accept={conn}"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    desktop = None
    try:
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_ctx)
        ctx = None
        for _ in range(120):
            try:
                ctx = resolver.resolve(f"uno:{conn}")
                break
            except Exception:
                if proc.poll() is not None:
                    raise RuntimeError("soffice exited before accepting connections")
                time.sleep(0.5)
        if ctx is None:
            raise RuntimeError("could not connect to soffice")
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(os.path.abspath(src)), "_blank", 0, (prop("Hidden", True),)
        )
        if doc is None:
            raise RuntimeError("LibreOffice could not open the document")
        # Fetch the index objects once: the collection goes stale after the first update().
        indexes = doc.getDocumentIndexes()
        items = [indexes.getByName(name) for name in indexes.getElementNames()]
        # Twice: filling the TOC can push headings to later pages.
        for _ in range(2):
            for item in items:
                item.update()
        doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)), (prop("FilterName", "MS Word 2007 XML"),))
        doc.close(True)
        print(f"{len(items)} index(es) updated")
    finally:
        if desktop is not None:
            try:
                desktop.terminate()
            except Exception:
                pass  # the bridge drops when soffice exits
        try:
            proc.wait(timeout=20)
        except subprocess.TimeoutExpired:
            proc.kill()
        shutil.rmtree(profile, ignore_errors=True)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    try:
        main(*sys.argv[1:4])
    except Exception as exc:  # report cleanly to the parent process
        print(f"error: {exc}", file=sys.stderr)
        sys.exit(1)
