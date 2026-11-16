"""Install-time checks for Document Image Access.

NVDA calls onInstall() when the add-on is installed or upgraded. This file looks for the Be My Eyes
app for Windows and, when it is missing, offers to open its Microsoft Store page. The add-on works
without Be My Eyes; only "Send to Be My Eyes" needs it.

onInstall() is not always called on the main thread, so every dialog goes through _run_on_main().
"""
import json
import os
import re
import subprocess
import threading

import addonHandler
import globalVars
import gui
import ui
import wx
from logHandler import log

try:
    addonHandler.initTranslation()
except Exception:  # translations are not always available while an add-on is still pending install
    def _(text):
        return text

# Be My Eyes for Windows in the Microsoft Store.
_STORE_ID = "9MSW46LTDWGF"
_STORE_URI = f"ms-windows-store://pdp/?ProductId={_STORE_ID}"
_STORE_WEB = f"https://apps.microsoft.com/detail/{_STORE_ID.lower()}"

# Same file that globalPlugins/docImages/bemyeyes.py reads ("aumid" is for ids written by hand).
_CONFIG = os.path.join(globalVars.appArgs.configPath, "docImages.json")
_AUMID_RE = re.compile(r"[\w.\-]+![\w.\-]+")
_NO_WINDOW = 0x08000000  # CREATE_NO_WINDOW


# ---------------------------------------------------------------- config file
def _load_config():
    try:
        with open(_CONFIG, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_config(**updates):
    config = _load_config()
    config.update(updates)
    try:
        with open(_CONFIG, "w", encoding="utf-8") as f:
            json.dump(config, f)
    except Exception:
        log.debugWarning("docImages: could not save the configuration", exc_info=True)


# ---------------------------------------------------------------- finding the app
def _find_bemyeyes():
    """Returns ("found", app id), ("missing", None) or ("unknown", None) when the check itself failed."""
    command = ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
               "Get-StartApps | Where-Object { $_.Name -like '*Be My Eyes*' } "
               "| Select-Object -First 1 -ExpandProperty AppID"]
    try:
        result = subprocess.run(command, capture_output=True, text=True, timeout=20, creationflags=_NO_WINDOW)
    except Exception:
        log.debugWarning("docImages: the Be My Eyes check could not run", exc_info=True)
        return "unknown", None
    if result.returncode != 0:
        return "unknown", None
    app_id = result.stdout.strip()
    if _AUMID_RE.fullmatch(app_id):
        return "found", app_id
    return "missing", None


def _run_on_main(fn):
    """Run fn on the main thread and wait for its result. Never blocks forever, even if fn fails."""
    if threading.current_thread() is threading.main_thread():
        return fn()
    done = threading.Event()
    box = {}

    def wrapper():
        try:
            box["result"] = fn()
        except Exception:
            log.error("docImages: install dialog failed", exc_info=True)
        finally:
            done.set()
    wx.CallAfter(wrapper)
    done.wait()
    return box.get("result")


def _open_store():
    try:
        os.startfile(_STORE_URI)
    except Exception:
        try:
            os.startfile(_STORE_WEB)
        except Exception:
            log.warning("docImages: could not open the Microsoft Store", exc_info=True)
            return False
    return True


# ---------------------------------------------------------------- dialog
class _GetBeMyEyesDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title=_("Document Image Access - Be My Eyes not found"))
        sizer = wx.BoxSizer(wx.VERTICAL)
        desc = wx.StaticText(
            self, wx.ID_ANY,
            _("Be My Eyes for Windows is not installed on this computer.\n\n"
              "Document Image Access can still list, open, copy, save and export the images in your "
              "documents. Sending an image to Be My Eyes for a description needs the Windows app, which "
              "is free from the Microsoft Store.\n\n"
              "Install it now and then choose Check again, or continue and install it later."))
        desc.Wrap(440)
        sizer.Add(desc, 0, wx.ALL, 12)

        buttons = wx.BoxSizer(wx.VERTICAL)
        self.btnStore = wx.Button(self, wx.ID_ANY, _("Open the Microsoft Store to install Be My Eyes"))
        self.btnCopy = wx.Button(self, wx.ID_ANY, _("Copy the Microsoft Store link"))
        self.btnCheck = wx.Button(self, wx.ID_ANY, _("Check again"))
        self.chkSkip = wx.CheckBox(self, wx.ID_ANY, _("Do not remind me again"))
        self.btnContinue = wx.Button(self, wx.ID_CANCEL, _("Continue without Be My Eyes"))
        for control in (self.btnStore, self.btnCopy, self.btnCheck):
            buttons.Add(control, 0, wx.EXPAND | wx.BOTTOM, 6)
        buttons.Add(self.chkSkip, 0, wx.TOP | wx.BOTTOM, 6)
        buttons.Add(self.btnContinue, 0, wx.EXPAND)
        sizer.Add(buttons, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 12)
        self.SetSizerAndFit(sizer)
        self.CentreOnScreen()

        self.btnStore.Bind(wx.EVT_BUTTON, self._on_store)
        self.btnCopy.Bind(wx.EVT_BUTTON, self._on_copy)
        self.btnCheck.Bind(wx.EVT_BUTTON, self._on_check)
        self.btnContinue.Bind(wx.EVT_BUTTON, self._on_continue)
        self.Bind(wx.EVT_CLOSE, self._on_continue)
        self.SetEscapeId(wx.ID_CANCEL)
        self.btnStore.SetFocus()

    def _on_store(self, event):
        if _open_store():
            ui.message(_("Opening the Microsoft Store"))
        else:
            wx.MessageBox(_("The Microsoft Store could not be opened. Use the Copy button to get the link."),
                          _("Be My Eyes"), wx.OK | wx.ICON_WARNING, self)

    def _on_copy(self, event):
        if wx.TheClipboard.Open():
            try:
                wx.TheClipboard.SetData(wx.TextDataObject(_STORE_WEB))
            finally:
                wx.TheClipboard.Close()
        wx.MessageBox(_("The Microsoft Store link was copied. Paste it into the address bar of your browser."),
                      _("Link copied"), wx.OK | wx.ICON_INFORMATION, self)

    def _on_check(self, event):
        with wx.BusyCursor():
            status, app_id = _find_bemyeyes()
        if status == "found":
            _save_config(detectedAumid=app_id)
            wx.MessageBox(_("Be My Eyes was found. You can now send images to it."),
                          _("Be My Eyes"), wx.OK | wx.ICON_INFORMATION, self)
            self.EndModal(wx.ID_OK)
        elif status == "missing":
            wx.MessageBox(_("Be My Eyes is still not installed. Finish the installation in the Microsoft Store, "
                            "then choose Check again."), _("Be My Eyes"), wx.OK | wx.ICON_INFORMATION, self)
        else:
            wx.MessageBox(_("The check could not be done. You can continue and install Be My Eyes later."),
                          _("Be My Eyes"), wx.OK | wx.ICON_WARNING, self)

    def _on_continue(self, event):
        if self.chkSkip.GetValue():
            _save_config(skipBeMyEyesCheck=True)
        self.EndModal(wx.ID_CANCEL)


def _show_dialog():
    dlg = _GetBeMyEyesDialog(gui.mainFrame)
    try:
        dlg.ShowModal()
    finally:
        dlg.Destroy()


# ---------------------------------------------------------------- NVDA entry point
def onInstall() -> None:
    config = _load_config()
    manual = config.get("aumid")  # written by hand for an app that is not named "Be My Eyes"
    if isinstance(manual, str) and _AUMID_RE.fullmatch(manual):
        return
    status, app_id = _find_bemyeyes()
    if status == "found":
        _save_config(detectedAumid=app_id)  # remembered, so that the first "Send to Be My Eyes" needs no search
        return
    if status == "unknown" or config.get("skipBeMyEyesCheck"):
        return
    _run_on_main(_show_dialog)
