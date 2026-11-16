"""Convert legacy .doc to .docx: Word COM first, LibreOffice headless as fallback."""
import gc
import glob
import os
import shutil
import subprocess

import addonHandler
from logHandler import log

addonHandler.initTranslation()

_NO_WINDOW = 0x08000000


def _via_word(src, dst):
    import comtypes
    import comtypes.client
    comtypes.CoInitialize()
    app = doc = None
    try:
        app = comtypes.client.CreateObject("Word.Application", dynamic=True)
        app.Visible = False
        app.DisplayAlerts = 0
        app.AutomationSecurity = 3
        doc = app.Documents.Open(src, False, True, False, "x")
        doc.SaveAs2(dst, 12)  # wdFormatXMLDocument
        return os.path.exists(dst)
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit(False)
        except Exception:
            pass
        doc = app = None
        gc.collect()
        comtypes.CoUninitialize()


def _find_soffice():
    p = shutil.which("soffice")
    if p:
        return p
    for pat in (r"C:\Program Files\LibreOffice*\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice*\program\soffice.exe"):
        hits = glob.glob(pat)
        if hits:
            return hits[0]
    return None


def _via_libreoffice(src, out_dir):
    exe = _find_soffice()
    if not exe:
        return None
    subprocess.run([exe, "--headless", "--convert-to", "docx", "--outdir", out_dir, src],
                   capture_output=True, timeout=180, creationflags=_NO_WINDOW)
    out = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + ".docx")
    return out if os.path.exists(out) else None


def convert_to_docx(src, out_dir):
    """Run in a worker thread. Returns path to .docx or raises RuntimeError."""
    os.makedirs(out_dir, exist_ok=True)
    dst = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + ".docx")
    try:
        if _via_word(src, dst):
            return dst
    except Exception:
        log.warning("docImages: Word conversion failed", exc_info=True)
    out = _via_libreoffice(src, out_dir)
    if out:
        return out
    raise RuntimeError(_("Could not convert this .doc file. It needs Microsoft Word or LibreOffice, "
                         "and may be password protected."))
