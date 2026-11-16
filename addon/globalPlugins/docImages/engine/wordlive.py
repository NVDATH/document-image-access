"""Operations on the document the user has open in Word (main thread only). No NVDA imports."""
import os


def in_protected_view(doc):
    """True when this document is open in Protected View (Word, Excel and PowerPoint share the API)."""
    try:
        mine = os.path.normcase(str(doc.FullName))
        pv = doc.Application.ProtectedViewWindows
        for i in range(1, int(pv.Count) + 1):
            w = pv.Item(i)
            if os.path.normcase(os.path.join(str(w.SourcePath), str(w.SourceName))) == mine:
                return True
    except Exception:
        pass
    return False


class LiveWordError(Exception):
    """code: unsaved, changed, notfound, readonly, protected, final, protectedview, denied, busy,
    unsupported, exportfailed.
    details carries diagnostics for the log."""

    def __init__(self, code, details=""):
        super().__init__(code)
        self.code = code
        self.details = details


class LiveWord:
    def __init__(self, doc, hwnd):
        self.doc = doc
        self.hwnd = hwnd
        self.saved = bool(doc.Saved)            # the list was built from the saved file
        self.end = int(doc.Content.End)         # detects edits made after the list was built

    def release(self):
        self.doc = None

    # what the dialog asks before enabling buttons
    def applies(self, img):
        return img.fmt == "docx"

    def can_export(self, img):
        return img.kind == "chart" and img.fmt == "docx" and img.part == "body"

    def can_page(self, img):
        return img.fmt == "docx" and img.part == "body"

    def position_key(self):
        """Where the user is in the document, as a comparable value."""
        return int(self.doc.ActiveWindow.Selection.Start)

    def nearest_index(self, items):
        """Index of the listed image closest to the caret, or None."""
        pos = self.position_key()
        best = None
        for i, img in enumerate(items):
            if img.fmt == "docx" and img.part == "body" and img.char_start is not None:
                d = abs(img.char_start - pos)
                if best is None or d < best[0]:
                    best = (d, i)
        return best[1] if best else None

    def get_alt(self, img):
        """The alt text as it is in the open document right now, or None."""
        shape = self._shape(img) if img.part == "body" else None
        return str(shape.AlternativeText) if shape is not None else None

    def export_chart(self, img, dest):
        if not self.unchanged():
            raise LiveWordError("changed")
        shape = self._shape(img)
        if shape is None:
            raise LiveWordError("notfound")
        try:
            shape.Chart.Export(dest, "PNG")
        except Exception as e:
            raise LiveWordError("protectedview" if in_protected_view(self.doc) else "exportfailed", repr(e))
        if not os.path.exists(dest):
            raise LiveWordError("exportfailed")

    def export_page(self, img, dest_base):
        """The page holding the image, as an enhanced metafile (the caller converts it to PNG)."""
        if not img.page:
            raise LiveWordError("nopage")
        if not self.unchanged():
            raise LiveWordError("changed")
        n = int(img.page)
        try:
            start = int(self.doc.GoTo(1, 1, n).Start)  # wdGoToPage, wdGoToAbsolute
            nxt = int(self.doc.GoTo(1, 1, n + 1).Start)
            end = nxt if nxt > start else int(self.doc.Content.End)
            data = self.doc.Range(start, end).EnhMetaFileBits
            dest = dest_base + ".emf"
            with open(dest, "wb") as f:
                f.write(bytes(data))
        except Exception as e:
            raise LiveWordError("exportfailed", repr(e))
        return dest

    def unchanged(self):
        return int(self.doc.Content.End) == self.end

    def _start(self, img):
        if img.char_start is not None:
            return img.char_start
        return int(self.doc.Paragraphs.Item(img.paragraph).Range.Start)

    def goto(self, img):
        """Move the caret in front of the image. Returns 'ok' or 'approximate'."""
        if img.part != "body":
            raise LiveWordError("unsupported")
        status = "ok" if self.unchanged() else "approximate"
        start = self._start(img)
        self.doc.Activate()
        self.doc.Range(start, start).Select()
        return status

    def _shape(self, img):
        doc = self.doc
        start = self._start(img)
        try:
            rng = doc.Range(start, start + 1)
            if rng.InlineShapes.Count > 0:
                return rng.InlineShapes.Item(1)
        except Exception:
            pass
        try:
            sr = doc.Range(start, start).Paragraphs.Item(1).Range.ShapeRange
            c = sr.Count
            if c > 0:
                return sr.Item(min(img.slot, c - 1) + 1)
        except Exception:
            pass
        return None

    def _blocked_by(self):
        """Why Word will refuse edits to this document, or None."""
        doc = self.doc
        for code, check in (  # Protected View first: it also reads as read-only, but has a clearer cure
                ("protectedview", lambda: in_protected_view(doc)),
                ("readonly", lambda: bool(doc.ReadOnly)),
                ("protected", lambda: int(doc.ProtectionType) != -1),  # wdNoProtection = -1
                ("final", lambda: bool(doc.Final))):
            try:
                if check():
                    return code
            except Exception:
                continue
        return None

    def _diagnostics(self, shape):
        out = []
        for label, getter in (
                ("ReadOnly", lambda: self.doc.ReadOnly),
                ("ProtectionType", lambda: self.doc.ProtectionType),
                ("Final", lambda: self.doc.Final),
                ("ReadOnlyRecommended", lambda: self.doc.ReadOnlyRecommended),
                ("ProtectedViewWindows", lambda: self.doc.Application.ProtectedViewWindows.Count),
                ("shape.Type", lambda: shape.Type)):
            try:
                out.append(f"{label}={getter()}")
            except Exception as e:
                out.append(f"{label}=<{e.__class__.__name__}>")
        return ", ".join(out)

    def set_alt(self, img, text):
        if img.part != "body":
            raise LiveWordError("unsupported")
        if not self.saved:
            raise LiveWordError("unsaved")
        if not self.unchanged():
            raise LiveWordError("changed")
        reason = self._blocked_by()
        if reason:
            raise LiveWordError(reason)
        shape = self._shape(img)
        if shape is None:
            raise LiveWordError("notfound")
        try:
            shape.AlternativeText = " ".join(text.split())
        except Exception as e:
            raise LiveWordError("denied", f"{e!r}; {self._diagnostics(shape)}")
