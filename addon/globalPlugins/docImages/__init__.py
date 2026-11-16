import functools
import os
import shutil
import tempfile
import threading
import uuid

import addonHandler
import api
import core
import globalPluginHandler
import gui
import scriptHandler
import tones
import ui
import winUser
import wx
from gui import guiHelper
from logHandler import log
from scriptHandler import script

from . import bemyeyes, winobj, wordconv
from .engine import docimages, wordlayout
from .engine.model import DocImageError
from .engine.liveoffice import LiveExcel, LivePowerPoint
from .engine.wordlive import LiveWord, LiveWordError

addonHandler.initTranslation()

try:
    ngettext
except NameError:  # older NVDA versions do not provide it
    def ngettext(singular, plural, n):
        return singular if n == 1 else plural

TEMP_ROOT = os.path.join(tempfile.gettempdir(), "nvda_docImages")


# ---------------------------------------------------------------- text helpers
def _location(img):
    """Human-readable location of an image. Every string is complete so it can be translated."""
    if img.fmt == "pptx" or (img.fmt == "odf" and img.slide):
        # Translators: Location of an image in a presentation.
        return _("Slide {n}").format(n=img.slide)
    if img.sheet:
        if img.chartsheet:
            return _("Chart sheet {sheet}").format(sheet=img.sheet)
        if img.kind == "incell":
            # Translators: An image placed inside a spreadsheet cell.
            return _("Sheet {sheet}, cell {cell}, in cell").format(sheet=img.sheet, cell=img.cell)
        if img.kind == "background":
            return _("Sheet {sheet}, background").format(sheet=img.sheet)
        if img.kind == "headerfooter":
            return {"LH": _("Sheet {sheet}, left header"), "CH": _("Sheet {sheet}, center header"),
                    "RH": _("Sheet {sheet}, right header"), "LF": _("Sheet {sheet}, left footer"),
                    "CF": _("Sheet {sheet}, center footer"), "RF": _("Sheet {sheet}, right footer"),
                    }.get(img.hf, _("Sheet {sheet}, header or footer")).format(sheet=img.sheet)
        if img.cell:
            # Translators: Location of an image in a spreadsheet.
            return _("Sheet {sheet}, cell {cell}").format(sheet=img.sheet, cell=img.cell)
        return _("Sheet {sheet}").format(sheet=img.sheet)
    if img.fmt == "docx" and img.part != "body":
        return {"header": _("In header"), "footer": _("In footer"),
                "footnotes": _("In footnotes"), "endnotes": _("In endnotes")}.get(img.part, _("Page unknown"))
    if img.page:
        if img.line:
            return _("Page {page}, line {line}").format(page=img.page, line=img.line)
        return _("Page {page}, line unknown").format(page=img.page)
    if img.page_est:
        # Translators: The page number is estimated from the page breaks stored in the file.
        return _("Page about {page}").format(page=img.page_est)
    return _("Page unknown")


def _cell(img, col):
    """NVDA reads a list row as its cells joined by ';' without column headers,
    so every cell describes itself."""
    if col == 0:
        if img.kind == "chart":
            return _("Chart {n}").format(n=img.index)
        return _("Image {n}").format(n=img.index)
    if col == 1:
        return _location(img)
    if col == 2:
        return _("Alt text: {text}").format(text=img.alt) if img.alt else _("No alt text")
    parts = []
    if img.in_textbox:
        parts.append(_("In a text box."))
    if img.kind == "chart" and img.title:
        parts.append(_("Title: {text}").format(text=img.title))
    if img.context and img.context != img.title:
        parts.append(_("Near: {text}").format(text=img.context))
    return " ".join(parts)


def _unlisted_note(n):
    return ngettext("{n} picture file in the document could not be located.",
                    "{n} picture files in the document could not be located.", n).format(n=n)


def _chart_note(n):
    return ngettext("{n} chart, which is not an image.", "{n} charts, which are not images.", n).format(n=n)


def _join(first, second):
    return _("{first}. {second}").format(first=first, second=second)


def _extra_notes(charts, unlisted):
    notes = []
    if charts:
        notes.append(_chart_note(charts))
    if unlisted:
        notes.append(_unlisted_note(unlisted))
    return notes


def _announcement(name, images, charts=0, unlisted=0):
    n = len(images)
    if not n:
        text = _("No images found in {name}").format(name=name)
    else:
        missing = sum(1 for i in images if not i.alt)
        if not missing:
            text = ngettext("{n} image, all with alt text", "{n} images, all with alt text", n).format(n=n)
        else:
            text = ngettext("{n} image, {m} without alt text", "{n} images, {m} without alt text",
                            n).format(n=n, m=missing)
    for note in _extra_notes(charts, unlisted):
        text = _join(text, note)
    return text


def _dialog_title(name, pictures, charts):
    if pictures:
        return ngettext("{n} image in {name}", "{n} images in {name}", pictures).format(n=pictures, name=name)
    if charts:
        return ngettext("{n} chart in {name}", "{n} charts in {name}", charts).format(n=charts, name=name)
    return name


def _initial_index(entry, items, live):
    """Row to start on. Stay on the row the user left when nothing moved since; otherwise start at the
    image closest to where the user is in the open document; otherwise the first row."""
    last = entry.get("last_index")
    if last is not None and not 0 <= last < len(items):
        last = None
    key = nearest = None
    if live is not None:
        try:
            key, nearest = live.position_key(), live.nearest_index(items)
        except Exception:
            log.debugWarning("docImages: could not read the position in the document", exc_info=True)
    if last is not None and (live is None or key is None or nearest is None or key == entry.get("last_key")):
        return last
    return nearest if nearest is not None else 0


def _live_error_text(code):
    return {
        "unsaved": _("Save the document first, then open this list again."),
        "changed": _("The document has changed since this list was made. Open the list again."),
        "notfound": _("Could not find this item in the open document."),
        "readonly": _("The document is read-only. Reopen it with editing enabled, then open this list again."),
        "protected": _("The document or sheet has editing restrictions, so this cannot be changed."),
        "final": _("The document is marked as final. Choose Edit Anyway, then open this list again."),
        "protectedview": _("The document is open in Protected View. Choose Enable Editing, "
                           "then open this list again."),
        "denied": _("The application refused the change. See the NVDA log for details."),
        "busy": _("The application is busy, for example while a cell is being edited. "
                  "Finish that and try again."),
        "unsupported": _("This kind of image cannot be reached or changed from here."),
        "exportfailed": _("Could not export it. See the NVDA log for details."),
        "nopage": _("The page of this image is not known, so the page cannot be exported."),
        "nosheet": _("Exporting a whole sheet is not available. Only Word pages and PowerPoint slides can be exported."),
    }.get(code, _("The operation failed."))


def _error_text(e):
    return {"unsupported": _("This file type is not supported."),
            "corrupt": _("Could not open this file. It may be password protected or corrupt."),
            "nocontent": _("This file has no readable document content."),
            "nopillow": _("This kind of image needs the Pillow library, which this version of NVDA does not "
                          "include.")}.get(e.code, _("Could not read this file."))


class _Progress:
    """Looping beep while something slow runs in a worker thread."""

    def __init__(self):
        self._stop = None

    def start(self):
        self.stop()
        ev = threading.Event()
        self._stop = ev

        def run():
            if ev.wait(0.3):  # quick jobs stay silent
                return
            while not ev.is_set():
                try:
                    tones.beep(600, 40)
                except Exception:
                    return
                ev.wait(1.0)
        threading.Thread(target=run, daemon=True).start()

    def stop(self):
        if self._stop:
            self._stop.set()
            self._stop = None


def _offer_store():
    """Be My Eyes is not installed: offer to open its Microsoft Store page."""
    answer = wx.MessageBox(
        _("Be My Eyes for Windows was not found on this computer. It is free from the Microsoft Store. "
          "Open its Store page now?\n\nIf it is installed under another name, set \"aumid\" in docImages.json."),
        _("Be My Eyes"), wx.YES_NO | wx.ICON_QUESTION)
    if answer == wx.YES and not bemyeyes.open_store():
        ui.message(_("The Microsoft Store could not be opened."))


def _copy_to_clipboard(path):
    img = wx.Image(path)
    if not img.IsOk() or not wx.TheClipboard.Open():
        return False
    try:
        wx.TheClipboard.SetData(wx.BitmapDataObject(wx.Bitmap(img)))
        wx.TheClipboard.Flush()
    finally:
        wx.TheClipboard.Close()
    return True


# ---------------------------------------------------------------- dialog
class ImagesListCtrl(wx.ListCtrl):
    """Virtual report list: rows are produced on demand, so thousands of images stay fast."""
    images = ()

    def OnGetItemText(self, item, col):
        try:
            return _cell(self.images[item], col)
        except IndexError:
            return ""


def _focus_list_after(method):
    """After a button action has spoken its result, hand the focus back to the list."""
    @functools.wraps(method)
    def wrapper(self, *args, **kwargs):
        try:
            return method(self, *args, **kwargs)
        finally:
            try:
                self.listCtrl.SetFocus()
            except RuntimeError:
                pass  # the action closed the dialog
    return wrapper


class DocImagesDialog(wx.Dialog):
    _instance = None

    def __init__(self, parent, title, images, live=None, chart_dir="", initial=0, state=None, can_send=True):
        super().__init__(parent, title=title)
        self.canSend = can_send
        self.state = state if state is not None else {}
        self.images = images      # pictures first, then charts
        self.live = live
        self.chartDir = chart_dir
        self.progress = _Progress()
        outer = wx.BoxSizer(wx.VERTICAL)
        h = guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
        self.listCtrl = h.addLabeledControl(
            _("&Images:"), ImagesListCtrl,
            style=wx.LC_REPORT | wx.LC_VIRTUAL | wx.LC_SINGLE_SEL, size=(760, 300))
        self.listCtrl.images = images
        for i, (label, width) in enumerate((
                (_("Image"), 90), (_("Location"), 200), (_("Alt text"), 260), (_("Nearby text"), 300))):
            self.listCtrl.InsertColumn(i, label, width=width)
        self.listCtrl.SetItemCount(len(images))
        row1 = guiHelper.ButtonHelper(wx.HORIZONTAL)
        self.openBtn = row1.addButton(self, label=_("&Open"))
        self.saveBtn = row1.addButton(self, label=_("&Save as..."))
        self.copyBtn = row1.addButton(self, label=_("Cop&y image"))
        self.sendBtn = row1.addButton(self, label=_("Send to &Be My Eyes"))
        if not can_send:  # Be My Eyes is not installed: nothing to send to
            self.sendBtn.Enable(False)
            row1.sizer.Hide(self.sendBtn)
        row2 = guiHelper.ButtonHelper(wx.HORIZONTAL)
        self.gotoBtn = row2.addButton(self, label=_("&Go to in document"))
        self.editBtn = row2.addButton(self, label=_("&Edit alt text..."))
        self.pageBtn = row2.addButton(self, label=_("&Whole page or slide..."))
        self.exportBtn = row2.addButton(self, label=_("E&xport all..."))
        self.closeBtn = row2.addButton(self, id=wx.ID_CLOSE, label=_("&Close"))
        h.addItem(row1.sizer)
        h.addItem(row2.sizer)
        outer.Add(h.sizer, border=guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
        self.SetSizer(outer)
        outer.Fit(self)
        self.listCtrl.Select(initial)
        self.listCtrl.Focus(initial)
        self.listCtrl.EnsureVisible(initial)
        self.state["last_index"] = initial
        self._updateButtons()
        self.openBtn.SetDefault()
        self.SetEscapeId(wx.ID_CLOSE)
        self.listCtrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.onOpen)
        self.listCtrl.Bind(wx.EVT_LIST_ITEM_SELECTED, self.onSelected)
        self.openBtn.Bind(wx.EVT_BUTTON, self.onOpen)
        self.saveBtn.Bind(wx.EVT_BUTTON, self.onSave)
        self.copyBtn.Bind(wx.EVT_BUTTON, self.onCopy)
        self.sendBtn.Bind(wx.EVT_BUTTON, self.onSend)
        self.gotoBtn.Bind(wx.EVT_BUTTON, self.onGoto)
        self.editBtn.Bind(wx.EVT_BUTTON, self.onEditAlt)
        self.pageBtn.Bind(wx.EVT_BUTTON, self.onPage)
        self.exportBtn.Bind(wx.EVT_BUTTON, self.onExport)
        self.closeBtn.Bind(wx.EVT_BUTTON, lambda e: self.Close())
        self.Bind(wx.EVT_CLOSE, self.onClose)
        self.CentreOnScreen()

    def _say(self, text):
        """Speak at once. The handlers speak first and the focus goes back to the list afterwards
        (see _focus_list_after), so the user hears the result before the list is announced."""
        ui.message(text)

    def onSelected(self, evt):
        self.state["last_index"] = evt.GetIndex()
        self._updateButtons()

    def _current(self):
        i = self.listCtrl.GetFirstSelected()
        return self.images[i] if i >= 0 else None

    def _updateButtons(self):
        img = self._current()
        applies = bool(self.live and img and self.live.applies(img))
        self.gotoBtn.Enable(applies)
        self.editBtn.Enable(applies)
        self.pageBtn.Enable(applies)

    def _ensureFile(self, img, quiet=False):
        """Path of the image file. A chart has none until the live application exports it."""
        if img.path and os.path.exists(img.path):
            return img.path
        if img.kind != "chart":
            return None
        if not self.live or not self.live.can_export(img):
            if not quiet:
                self._say(_("Charts can only be exported while the document is open in Excel or PowerPoint."))
            return None
        dest = os.path.join(self.chartDir, f"chart_{img.index:03d}.png")
        try:
            os.makedirs(self.chartDir, exist_ok=True)
            self.live.export_chart(img, dest)
        except LiveWordError as e:
            log.warning("docImages: chart export failed: %s %s", e.code, e.details)
            if not quiet:
                self._say(_live_error_text(e.code))
            return None
        except Exception:
            log.warning("docImages: chart export failed", exc_info=True)
            if not quiet:
                self._say(_("Could not export the chart. See the NVDA log for details."))
            return None
        img.path = dest
        return dest

    # -- actions on the selected image
    # --- the actions, on any image file (a picture, an exported chart, an exported page)
    def _open(self, path):
        try:
            os.startfile(path)
        except Exception:
            log.error("docImages: open failed", exc_info=True)
            self._say(_("Could not open the image"))

    def _save(self, path, default_name):
        with wx.FileDialog(self, _("Save image"), defaultFile=default_name,
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dlg:
            if dlg.ShowModal() != wx.ID_OK:
                return
            try:
                shutil.copyfile(path, dlg.GetPath())
                self._say(_("Image saved"))
            except Exception:
                log.error("docImages: save failed", exc_info=True)
                self._say(_("Could not save the image"))

    def _copy(self, source):
        def work():
            try:
                path = docimages.prepare_for_sharing(source)
            except DocImageError as e:
                wx.CallAfter(self._say, _error_text(e))
                return
            except Exception:
                log.error("docImages: image conversion failed", exc_info=True)
                path = None
            wx.CallAfter(self._finishCopy, path)
        threading.Thread(target=work, daemon=True).start()

    def _finishCopy(self, path):
        ok = False
        try:
            ok = bool(path) and _copy_to_clipboard(path)
        except Exception:
            log.error("docImages: clipboard failed", exc_info=True)
        self._say(_("Image copied to the clipboard") if ok else _("Could not copy the image"))

    def _send(self, path):
        self._say(_("Sending to Be My Eyes"))
        self.progress.start()

        def done(status):
            self.progress.stop()
            if status == "notfound":
                wx.CallAfter(_offer_store)
                return
            msg = {"ok": _("Be My Eyes launched with the image"),
                   "failed": _("Could not send the image to Be My Eyes"),
                   "nopillow": _error_text(DocImageError("nopillow"))}[status]
            wx.CallAfter(self._say, msg)
        bemyeyes.send_async(path, done)

    @_focus_list_after
    def onOpen(self, evt):
        img = self._current()
        path = self._ensureFile(img) if img else None
        if path:
            self._open(path)

    @_focus_list_after
    def onSave(self, evt):
        img = self._current()
        path = self._ensureFile(img) if img else None
        if path:
            self._save(path, docimages.export_name(img))

    @_focus_list_after
    def onCopy(self, evt):
        img = self._current()
        source = self._ensureFile(img) if img else None
        if source:
            self._copy(source)

    @_focus_list_after
    def onSend(self, evt):
        if not self.canSend:
            return
        img = self._current()
        path = self._ensureFile(img) if img else None
        if path:
            self._send(path)

    @_focus_list_after
    def onPage(self, evt):
        """Picture the whole page (Word) or slide (PowerPoint) that holds the selected image, then offer
        the usual actions on it. Useful to have Be My Eyes describe the overall layout."""
        img = self._current()
        if not img or not self.live:
            return
        base = f"slide_{img.slide:02d}" if img.fmt == "pptx" else f"page_{int(img.page or 0):03d}"
        try:
            os.makedirs(self.chartDir, exist_ok=True)
            path = self.live.export_page(img, os.path.join(self.chartDir, base))
            path = docimages.prepare_for_sharing(path)  # Word hands over a metafile
        except LiveWordError as e:
            log.warning("docImages: page export failed: %s %s", e.code, e.details)
            self._say(_live_error_text(e.code))
            return
        except DocImageError as e:
            self._say(_error_text(e))
            return
        except Exception:
            log.warning("docImages: page export failed", exc_info=True)
            self._say(_("Could not export the page. See the NVDA log for details."))
            return
        menu = wx.Menu()
        choices = [(_("&Open"), lambda e: self._open(path)),
                   (_("&Save as..."), lambda e: self._save(path, os.path.basename(path))),
                   (_("&Copy to the clipboard"), lambda e: self._copy(path))]
        if self.canSend:
            choices.append((_("Send to &Be My Eyes"), lambda e: self._send(path)))
        for label, action in choices:
            self.Bind(wx.EVT_MENU, action, menu.Append(wx.ID_ANY, label))
        pos = self.pageBtn.GetPosition()
        self.PopupMenu(menu, wx.Point(pos.x, pos.y + self.pageBtn.GetSize().height))
        menu.Destroy()

    @_focus_list_after
    def onGoto(self, evt):
        img = self._current()
        if not img or not self.live:
            return
        try:
            status = self.live.goto(img)
        except LiveWordError as e:
            log.warning("docImages: go to failed: %s %s", e.code, e.details)
            self._say(_live_error_text(e.code))
            return
        except Exception:
            log.warning("docImages: go to failed", exc_info=True)
            self._say(_("Could not move to the image. The document may have been closed."))
            return
        hwnd = self.live.hwnd
        self.Close()
        try:
            if hwnd:
                winUser.setForegroundWindow(hwnd)
        except Exception:
            log.debugWarning("docImages: could not refocus the application", exc_info=True)
        if status != "ok":
            core.callLater(600, ui.message, _(
                "The document has changed since this list was made, so the position may be approximate."))

    @_focus_list_after
    def onEditAlt(self, evt):
        i = self.listCtrl.GetFirstSelected()
        img = self._current()
        if not img or not self.live:
            return
        current = img.alt
        try:  # the alt text as it is in the open document right now
            live_alt = self.live.get_alt(img)
            if live_alt is not None:
                current = live_alt
        except Exception:
            log.debugWarning("docImages: could not read the current alt text", exc_info=True)
        prompt = (_("Alt text for chart {n}:") if img.kind == "chart" else _("Alt text for image {n}:")).format(
            n=img.index)
        with wx.TextEntryDialog(self, prompt, _("Edit alt text"), value=current,
                                style=wx.OK | wx.CANCEL | wx.TE_MULTILINE) as dlg:
            if dlg.ShowModal() != wx.ID_OK:
                return
            text = " ".join(dlg.GetValue().split())
        try:
            self.live.set_alt(img, text)
        except LiveWordError as e:
            if e.details:
                log.warning("docImages: the application refused the alt text change: %s", e.details)
            self._say(_live_error_text(e.code))
            return
        except Exception:
            log.warning("docImages: set alt text failed", exc_info=True)
            self._say(_("Could not change the alt text. The document may have been closed."))
            return
        img.alt = text
        self.listCtrl.RefreshItem(i)
        self._say(_("Alt text updated. Save the document to keep the change."))

    @_focus_list_after
    def onExport(self, evt):
        with wx.DirDialog(self, _("Choose a folder for the images")) as dlg:
            if dlg.ShowModal() != wx.ID_OK:
                return
            folder = dlg.GetPath()
        self._say(_("Exporting images"))
        for img in self.images:  # charts are exported through the application, which needs this thread
            if img.kind == "chart":
                self._ensureFile(img, quiet=True)
        self.progress.start()

        def work():
            try:
                n = docimages.export_all(self.images, folder)
                msg = ngettext("Exported {n} file", "Exported {n} files", n).format(n=n)
            except Exception:
                log.error("docImages: export failed", exc_info=True)
                msg = _("Could not export the images")
            self.progress.stop()
            wx.CallAfter(self._say, msg)
        threading.Thread(target=work, daemon=True).start()

    def onClose(self, evt):
        self.progress.stop()
        if self.live:
            try:  # remember where the user was, to restore the same row if nothing moves meanwhile
                self.state["last_key"] = self.live.position_key()
            except Exception:
                self.state.pop("last_key", None)
            self.live.release()  # do not keep Word's COM objects alive after the dialog is gone
            self.live = None
        DocImagesDialog._instance = None
        self.Destroy()


# ---------------------------------------------------------------- plugin
class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = _("Document Image Access")

    def __init__(self):
        super().__init__()
        self._progress = _Progress()
        self._lock = threading.Lock()
        self._busy = False
        self._wantFull = False
        self._path = None        # file of the last first press, reused by the second press
        self._live = None        # LiveWord for that file when it came from Word
        self._cache = None       # scan result of the last file
        self._toldNoBeMyEyes = False  # said once per NVDA session that sending is turned off
        _clean_temp()

    def terminate(self):
        self._progress.stop()
        self._releaseLive()
        self._cache = None
        _clean_temp()
        super().terminate()

    def _releaseLiveIfIdle(self):
        """After a quick summary nobody asked for the list, so stop holding Word's COM objects."""
        if not DocImagesDialog._instance and not self._busy:
            self._releaseLive()

    def _releaseLive(self):
        if self._live:
            self._live.release()
            self._live = None

    # ---- finding the source file (main thread; uses COM) ----
    def _from_nvda_objects(self, app, focus):
        if app == "winword":
            getters = (lambda: focus.WinwordDocumentObject,
                       lambda: focus.WinwordApplicationObject.ActiveDocument)
        elif app == "excel":
            getters = (lambda: focus.excelWorksheetObject.Parent,
                       lambda: focus.excelCellObject.Worksheet.Parent,
                       lambda: focus.excelWindowObject.Application.ActiveWorkbook)
        else:  # powerpnt
            getters = (lambda: focus.ppObjectModel.Presentation,
                       lambda: focus.ppObjectModel.Application.ActivePresentation)
        for getter in getters:
            try:
                doc = getter()
                if doc is not None and str(doc.FullName):
                    return doc
            except Exception:
                continue
        return None

    def _from_running_object_table(self, app):
        import comtypes.client
        prog, attr = {"winword": ("Word.Application", "ActiveDocument"),
                      "excel": ("Excel.Application", "ActiveWorkbook"),
                      "powerpnt": ("PowerPoint.Application", "ActivePresentation")}[app]
        return getattr(comtypes.client.GetActiveObject(prog, dynamic=True), attr)

    def _office_document(self, app):
        """The open Word document / Excel workbook / PowerPoint presentation. Tries, in order: the COM
        objects NVDA's app modules hold for the focused control (not present when NVDA uses UI Automation
        for that control), the object model of the window itself (what NVDA's modules use underneath),
        and finally the running object table."""
        focus = api.getFocusObject()
        attempts = (
            ("NVDA objects", lambda: self._from_nvda_objects(app, focus)),
            ("native object model", lambda: winobj.document_from_window(app, api.getForegroundObject().windowHandle)),
            ("running object table", lambda: self._from_running_object_table(app)),
        )
        errors = []
        for label, attempt in attempts:
            try:
                doc = attempt()
                if doc is not None and str(doc.FullName):
                    log.debug("docImages: %s document found via %s", app, label)
                    return doc
                errors.append(f"{label}: no document")
            except Exception as e:
                errors.append(f"{label}: {e!r}")
        log.warning("docImages: could not reach the document in %s. Focus object: %s. Attempts: %s",
                    app, [c.__name__ for c in type(focus).__mro__], "; ".join(errors))
        return None

    def _office_path(self, app):
        """Returns (path, message, allow_file_dialog_fallback)."""
        doc = self._office_document(app)
        if doc is None:
            log.warning("docImages: no document object reachable from the focus in %s", app)
            return None, _("Could not reach the open document. Choose the file manually."), True
        try:
            path = str(doc.FullName)
            saved = bool(doc.Saved)
        except Exception:
            log.warning("docImages: could not read the document properties", exc_info=True)
            return None, _("Could not reach the open document. Choose the file manually."), True
        if path.lower().startswith(("http://", "https://")):
            return None, _("This document is stored online. Choose a local copy."), True
        if not os.path.isabs(path) or not os.path.exists(path):
            return None, _("Please save the document first."), False
        try:  # Go to, alt text editing and chart export work on the open document
            live_class = {"winword": LiveWord, "excel": LiveExcel, "powerpnt": LivePowerPoint}[app]
            self._live = live_class(doc, api.getForegroundObject().windowHandle)
        except Exception:
            log.debugWarning("docImages: live access to %s unavailable", app, exc_info=True)
        return path, (None if saved else _(
            "The document has unsaved changes; using the last saved version.")), True

    def _explorer_path(self, hwnd):
        import comtypes.client
        shell = comtypes.client.CreateObject("Shell.Application", dynamic=True)
        wins = shell.Windows()
        for i in range(wins.Count):
            try:
                w = wins.Item(i)
                if int(w.HWND) != hwnd:
                    continue
                sel = w.Document.SelectedItems()
                if sel.Count > 0:
                    return str(sel.Item(0).Path)
            except Exception:
                continue
        return None

    def _locate(self):
        fg = api.getForegroundObject()
        app = (getattr(fg.appModule, "appName", "") or "").lower()
        try:
            if app in ("winword", "excel", "powerpnt"):
                return self._office_path(app)
            if app == "explorer":
                p = self._explorer_path(fg.windowHandle)
                if p and os.path.splitext(p)[1].lower() in docimages.SUPPORTED_EXTS:
                    return p, None, True
        except Exception:
            log.warning("docImages: locate failed", exc_info=True)
        return None, None, True

    def _ask_file(self):
        exts = ";".join("*" + e for e in sorted(docimages.SUPPORTED_EXTS))
        gui.mainFrame.prePopup()
        try:
            with wx.FileDialog(gui.mainFrame, _("Choose a document"),
                               wildcard=_("Supported documents") + f"|{exts}",
                               style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dlg:
                return dlg.GetPath() if dlg.ShowModal() == wx.ID_OK else None
        finally:
            gui.mainFrame.postPopup()

    # ---- the script ----
    @script(description=_("Announces how many images the current document has. "
                          "Press twice quickly to list them."),
            gesture="kb:NVDA+shift+i")
    def script_documentImages(self, gesture):
        if DocImagesDialog._instance:
            DocImagesDialog._instance.Raise()
            return
        full = scriptHandler.getLastScriptRepeatCount() >= 1
        if full and self._path:
            self._begin(self._path, True)
            return
        self._releaseLive()
        self._path = None
        path, msg, fallback = self._locate()
        if msg:
            ui.message(msg)
        if path:
            self._path = path
            self._begin(path, full)
        elif fallback:
            wx.CallAfter(self._chooseAndRun)

    def _chooseAndRun(self):
        p = self._ask_file()
        if p:
            self._path = p
            self._begin(p, True)

    def _begin(self, path, full):
        with self._lock:
            if self._busy:
                if full:
                    self._wantFull = True  # second press while the first job is still running
                return
            self._busy = True
            self._wantFull = full
        self._progress.start()
        threading.Thread(target=self._worker, args=(path,), daemon=True).start()

    # ---- background work ----
    def _prepare(self, path):
        st = os.stat(path)
        key = (path, st.st_mtime_ns, st.st_size)
        if self._cache and self._cache["key"] == key:
            return self._cache
        if self._cache:
            shutil.rmtree(self._cache["work"], ignore_errors=True)
        work = os.path.join(TEMP_ROOT, uuid.uuid4().hex)
        src = path
        if os.path.splitext(path)[1].lower() in docimages.LEGACY_WORD_EXTS:
            wx.CallAfter(ui.message, _("Converting the .doc file with Word. This can take a while."))
            src = wordconv.convert_to_docx(path, work)
        images_found = docimages.scan_images(src)
        entry = {"key": key, "src": src, "work": work, "images": images_found,
                 "extracted": False, "annotated": False, "exact": False, "unlisted": 0, "can_send": True,
                 "charts": getattr(images_found, "charts", 0),
                 "chart_items": list(getattr(images_found, "chart_items", ()))}
        try:
            entry["unlisted"] = docimages.unlisted_count(src, entry["images"])
        except Exception:
            log.debugWarning("docImages: could not count unlisted pictures", exc_info=True)
        self._cache = entry
        return entry

    def _worker(self, path):
        name = os.path.basename(path)
        quick = False
        try:
            entry = self._prepare(path)
            images = entry["images"]
            with self._lock:
                if not (self._wantFull and (images or entry["chart_items"])):
                    self._busy = False
                    quick = True
            if quick:
                self._progress.stop()
                wx.CallAfter(ui.message, _announcement(name, images, entry["charts"], entry["unlisted"]))
                wx.CallAfter(core.callLater, 10000, self._releaseLiveIfIdle)
                return
            if not entry["extracted"]:
                docimages.extract_files(entry["src"], images, os.path.join(entry["work"], "images"))
                entry["extracted"] = True
            if not entry["annotated"]:
                entry["annotated"] = True
                if docimages.family(entry["src"]) == "docx":
                    try:
                        entry["exact"] = wordlayout.annotate(entry["src"], list(images) + list(entry["chart_items"]))
                    except Exception:
                        log.warning("docImages: Word layout lookup failed", exc_info=True)
            try:  # a search for the Be My Eyes app: remembered for a while, so usually instant
                entry["can_send"] = bemyeyes.is_available()
            except Exception:
                log.debugWarning("docImages: could not check for Be My Eyes", exc_info=True)
                entry["can_send"] = True
            self._progress.stop()
            wx.CallAfter(self._show, name, entry)
        except Exception as e:
            self._progress.stop()
            if isinstance(e, DocImageError):
                msg = _error_text(e)
            elif isinstance(e, RuntimeError):
                msg = str(e)
            else:
                log.error("docImages: worker failed", exc_info=True)
                msg = _("Could not read images from this file")
            wx.CallAfter(ui.message, msg)
        finally:
            with self._lock:
                self._busy = False

    def _show(self, name, entry):
        images, charts = entry["images"], entry["chart_items"]
        gui.mainFrame.prePopup()
        items = list(images) + list(charts)
        dlg = DocImagesDialog(gui.mainFrame, _dialog_title(name, len(images), len(charts)), items, self._live,
                              os.path.join(entry["work"], "charts"),
                              initial=_initial_index(entry, items, self._live), state=entry,
                              can_send=entry.get("can_send", True))
        self._live = None  # the dialog owns it now
        DocImagesDialog._instance = dlg
        dlg.Show()
        dlg.Raise()
        dlg.listCtrl.SetFocus()
        gui.mainFrame.postPopup()
        title = dlg.GetTitle()
        if docimages.family(entry["src"]) == "docx" and not entry["exact"]:
            title = _("{title}. Page and line numbers could not be read from Word; pages are estimated.").format(
                title=title)
        for note in _extra_notes(entry["charts"], entry["unlisted"]):
            title = _join(title, note)
        if not entry.get("can_send", True) and not self._toldNoBeMyEyes:
            self._toldNoBeMyEyes = True
            title = _join(title, _("Be My Eyes was not found, so sending to it is turned off."))
        ui.message(title)


def _clean_temp():
    shutil.rmtree(TEMP_ROOT, ignore_errors=True)
