"""Reach an Office application's object model from a window handle, the same way NVDA's own
Office app modules do (AccessibleObjectFromWindow with OBJID_NATIVEOM). Windows only; no NVDA imports."""
import ctypes
from ctypes import wintypes

OBJID_NATIVEOM = -16
# Window classes that expose the object model: Word document pane, Excel sheet window, PowerPoint panes.
_CLASSES = {"winword": ("_WwG",), "excel": ("EXCEL7",), "powerpnt": ("paneClassDC", "screenClass")}


def _child_windows(top_hwnd, wanted):
    user32 = ctypes.windll.user32
    proc_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user32.EnumChildWindows.argtypes = [wintypes.HWND, proc_type, wintypes.LPARAM]
    found = []

    def callback(hwnd, _lparam):
        buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, buf, 256)
        if buf.value in wanted:
            found.append(hwnd)
        return True
    user32.EnumChildWindows(top_hwnd, proc_type(callback), 0)
    return found


def _native_object(hwnd):
    import comtypes.automation
    import comtypes.client
    ptr = ctypes.POINTER(comtypes.automation.IDispatch)()
    ctypes.oledll.oleacc.AccessibleObjectFromWindow(
        hwnd, OBJID_NATIVEOM, ctypes.byref(ptr._iid_), ctypes.byref(ptr))
    return comtypes.client.dynamic.Dispatch(ptr)


def document_from_window(app, top_hwnd):
    """The Document / Workbook / Presentation shown in the given top-level window, or None."""
    for hwnd in _child_windows(top_hwnd, _CLASSES[app]):
        try:
            win = _native_object(hwnd)
            if app == "excel":
                doc = win.Application.ActiveWorkbook
            elif app == "winword":
                doc = win.Document
            else:
                doc = win.Presentation
            if doc is not None and str(doc.FullName):
                return doc
        except Exception:
            continue
    return None
