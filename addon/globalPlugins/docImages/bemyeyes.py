import ctypes
import json
import os
import re
import subprocess
import threading
import time

import globalVars
from logHandler import log

# docImages.json, in the NVDA configuration folder, uses these keys:
#   "aumid"          written by hand, for a Be My Eyes app whose name does not contain "Be My Eyes"
#   "detectedAumid"  written by this add-on when it finds the app (a remembered result)
#   "skipBeMyEyesCheck"  written by the installation dialog ("Do not remind me again")
_CFG = os.path.join(globalVars.appArgs.configPath, "docImages.json")
_AUMID_RE = re.compile(r"[\w.\-]+![\w.\-]+")
_NO_WINDOW = 0x08000000  # CREATE_NO_WINDOW

# Be My Eyes for Windows in the Microsoft Store (installTasks.py uses the same product).
STORE_ID = "9MSW46LTDWGF"
STORE_URI = f"ms-windows-store://pdp/?ProductId={STORE_ID}"
STORE_WEB = f"https://apps.microsoft.com/detail/{STORE_ID.lower()}"

_FOUND_FOR = 600   # seconds a positive search result is trusted
_MISSING_FOR = 30  # seconds a negative one is trusted: the app may be installing right now
_memo = {}


def _load():
    try:
        with open(_CFG, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save(d):
    try:
        with open(_CFG, "w", encoding="utf-8") as f:
            json.dump(d, f)
    except Exception:
        log.debugWarning("docImages: cannot save config", exc_info=True)


def _valid(aumid):
    return aumid if isinstance(aumid, str) and _AUMID_RE.fullmatch(aumid) else None


def _search():
    """("found", id), ("missing", None), or ("unknown", None) when the search itself could not run."""
    cmd = ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
           "Get-StartApps | Where-Object { $_.Name -like '*Be My Eyes*' } "
           "| Select-Object -First 1 -ExpandProperty AppID"]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=20, creationflags=_NO_WINDOW)
    except Exception:
        log.debugWarning("docImages: the Be My Eyes search could not run", exc_info=True)
        return "unknown", None
    if result.returncode != 0:
        return "unknown", None
    found = _valid(result.stdout.strip())
    return ("found", found) if found else ("missing", None)


def availability(force=False):
    """("found", id), ("missing", None) or ("unknown", None). An id written by hand in docImages.json
    counts as found. May take seconds: call it off the main thread."""
    manual = _valid(_load().get("aumid"))
    if manual:
        return "found", manual
    now = time.monotonic()
    memo = _memo.get("result")
    if memo and not force:
        status, aumid, when = memo
        if now - when < {"found": _FOUND_FOR, "missing": _MISSING_FOR}.get(status, 0):
            return status, aumid
    status, aumid = _search()
    if status == "found":
        config = _load()
        config["detectedAumid"] = aumid
        _save(config)
    _memo["result"] = (status, aumid, now)
    return status, aumid


def is_available():
    """False only when the app is known to be missing. When the search could not run, assume it is there."""
    return availability()[0] != "missing"


def send_to_bemyeyes(file_path: str, aumid: str) -> bool:
    """Launch Be My Eyes via shell:appsFolder\\<AUMID> with file_path as parameter.
    ShellExecute returns > 32 on success."""
    fn = ctypes.windll.shell32.ShellExecuteW
    fn.restype = ctypes.c_void_p
    fn.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_wchar_p,
                   ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_int]
    param = f'"{file_path}"' if " " in file_path else file_path
    result = fn(None, "open", f"shell:appsFolder\\{aumid}", param, None, 1)
    return (result or 0) > 32


def send_async(file_path, on_done):
    """on_done(status) with status in 'ok', 'notfound', 'failed', 'nopillow'; called from the worker thread."""
    def run():
        try:
            from .engine import docimages
            from .engine.model import DocImageError
            status, aumid = availability()
            if status == "missing":
                return on_done("notfound")
            aumid = aumid or _valid(_load().get("detectedAumid"))  # the search could not run: use what was found before
            if not aumid:
                return on_done("failed")
            try:
                prepared = docimages.prepare_for_sharing(file_path)
            except DocImageError as e:
                return on_done("nopillow" if e.code == "nopillow" else "failed")
            if send_to_bemyeyes(prepared, aumid):
                return on_done("ok")
            # The launch failed: the app may have been uninstalled since it was found.
            on_done("notfound" if availability(force=True)[0] == "missing" else "failed")
        except Exception:
            log.error("docImages: send failed", exc_info=True)
            on_done("failed")
    threading.Thread(target=run, daemon=True).start()


def open_store():
    """Open the Be My Eyes page in the Microsoft Store. Returns False when it cannot be opened."""
    for target in (STORE_URI, STORE_WEB):
        try:
            os.startfile(target)
            return True
        except Exception:
            log.debugWarning("docImages: could not open %s", target, exc_info=True)
    return False
