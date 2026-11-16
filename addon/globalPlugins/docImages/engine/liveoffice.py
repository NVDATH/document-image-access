"""Operations on the workbook / presentation the user has open in Excel or PowerPoint (main thread only).
Shapes are found by the name stored in the file, so no position guessing is needed. No NVDA imports."""
import os
import re

from .wordlive import LiveWordError, in_protected_view

# COM errors that mean "the application is busy right now" (editing a cell, a dialog is open, ...)
_BUSY = {-2147418111, -2147417846, -2146777998}  # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER, VBA_E_IGNORE


def _hresult(e):
    code = getattr(e, "hresult", None)
    if code is None and getattr(e, "args", None) and isinstance(e.args[0], int):
        code = e.args[0]
    return code


def _guard(fn):
    """Run a COM call; turn 'application busy' into a LiveWordError, let other errors through."""
    try:
        return fn()
    except LiveWordError:
        raise
    except Exception as e:
        if _hresult(e) in _BUSY:
            raise LiveWordError("busy", repr(e))
        raise


def _cell_rc(address):
    m = re.fullmatch(r"([A-Za-z]+)(\d+)", address or "")
    if not m:
        return None
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)), col


def _find_shape(shapes, name):
    """Shape by name, also searching inside groups. None when absent."""
    try:
        return shapes.Item(name)
    except Exception:
        pass
    try:
        for i in range(1, int(shapes.Count) + 1):
            shape = shapes.Item(i)
            try:
                if int(shape.Type) == 6:  # msoGroup
                    found = _find_shape(shape.GroupItems, name)
                    if found is not None:
                        return found
            except Exception:
                continue
    except Exception:
        pass
    return None


class _LiveBase:
    def __init__(self, doc, hwnd):
        self.doc = doc
        self.hwnd = hwnd

    def release(self):
        self.doc = None

    def can_export(self, img):
        return img.kind == "chart"

    def can_page(self, img):
        return False

    def export_page(self, img, dest_base):
        raise LiveWordError("nosheet")  # no safe way to picture a whole sheet without altering the workbook

    def _diagnostics(self, *getters):
        out = []
        for label, getter in getters:
            try:
                out.append(f"{label}={getter()}")
            except Exception as e:
                out.append(f"{label}=<{e.__class__.__name__}>")
        return ", ".join(out)


class LiveExcel(_LiveBase):
    def applies(self, img):
        return img.fmt == "xlsx"

    def _sheet(self, img):
        return self.doc.Worksheets.Item(img.sheet)

    def goto(self, img):
        if img.kind in ("background", "headerfooter"):
            raise LiveWordError("unsupported")

        def run():
            self.doc.Activate()
            if img.chartsheet:
                self.doc.Charts.Item(img.sheet).Activate()
            else:
                ws = self._sheet(img)
                ws.Activate()
                if img.cell:
                    ws.Range(img.cell).Select()
        _guard(run)
        return "ok"

    def _blocked_by(self, ws):
        doc = self.doc
        for code, check in (
                ("protectedview", lambda: in_protected_view(doc)),
                ("readonly", lambda: bool(doc.ReadOnly)),
                ("protected", lambda: bool(ws.ProtectDrawingObjects))):
            try:
                if check():
                    return code
            except Exception:
                continue
        return None

    def _alt_targets(self, ws, img):
        """Objects that carry the alt text of the shape, most direct first."""
        shape = _find_shape(ws.Shapes, img.shape_name)
        if shape is not None:
            yield shape
        if img.kind == "chart":
            try:
                yield ws.ChartObjects().Item(img.shape_name).ShapeRange.Item(1)
            except Exception:
                pass

    def get_alt(self, img):
        if img.kind in ("incell", "background", "headerfooter") or img.chartsheet or not img.shape_name:
            return None
        for target in self._alt_targets(self._sheet(img), img):
            try:
                return str(target.AlternativeText)
            except Exception:
                continue
        return None

    def set_alt(self, img, text):
        # In-cell pictures, backgrounds, header pictures and chart sheets have no shape to change.
        if img.kind in ("incell", "background", "headerfooter") or img.chartsheet or not img.shape_name:
            raise LiveWordError("unsupported")
        cleaned = " ".join(text.split())

        def run():
            ws = self._sheet(img)
            reason = self._blocked_by(ws)
            if reason:
                raise LiveWordError(reason)
            errors, found = [], False
            for target in self._alt_targets(ws, img):
                found = True
                try:
                    target.AlternativeText = cleaned
                    return
                except Exception as e:
                    if _hresult(e) in _BUSY:
                        raise
                    errors.append(repr(e))
            if not found:
                raise LiveWordError("notfound")
            raise LiveWordError("denied", "; ".join(errors) + "; " + self._diagnostics(
                ("ReadOnly", lambda: self.doc.ReadOnly),
                ("ProtectDrawingObjects", lambda: ws.ProtectDrawingObjects)))
        _guard(run)

    def export_chart(self, img, dest):
        def run():
            if in_protected_view(self.doc):  # Excel then fails with a misleading "dimension" error
                raise LiveWordError("protectedview")
            if img.chartsheet:
                chart = self.doc.Charts.Item(img.sheet)
            else:
                chart = self._sheet(img).ChartObjects().Item(img.shape_name).Chart
            if not chart.Export(dest, "PNG") or not os.path.exists(dest):
                raise LiveWordError("exportfailed")
        try:
            _guard(run)
        except LiveWordError:
            raise
        except Exception as e:
            raise LiveWordError("protectedview" if in_protected_view(self.doc) else "exportfailed", repr(e))

    # --- where the user is, so the list can start near it
    def _position(self):
        name = str(self.doc.ActiveSheet.Name)
        try:
            cell = self.doc.Application.ActiveCell
            return name, (int(cell.Row), int(cell.Column))
        except Exception:
            return name, None  # a chart sheet has no active cell

    def position_key(self):
        return self._position()

    def nearest_index(self, items):
        name, rc = self._position()
        order = {}
        for i in range(1, int(self.doc.Sheets.Count) + 1):
            order[str(self.doc.Sheets.Item(i).Name)] = i
        best = None
        for i, img in enumerate(items):
            if img.fmt != "xlsx":
                continue
            gap = abs(order.get(img.sheet, 0) - order.get(name, 0))
            cell = _cell_rc(img.cell)
            if gap:
                d = 10 ** 7 * gap
            elif rc and cell:
                d = abs(cell[0] - rc[0]) * 1000 + abs(cell[1] - rc[1])
            else:
                d = 10 ** 6
            if best is None or d < best[0]:
                best = (d, i)
        return best[1] if best else None


class LivePowerPoint(_LiveBase):
    def applies(self, img):
        return img.fmt == "pptx"

    def _slide(self, img):
        return self.doc.Slides.Item(int(img.slide))

    def can_page(self, img):
        return img.fmt == "pptx" and bool(img.slide)

    def export_page(self, img, dest_base):
        """The whole slide as a PNG."""
        dest = dest_base + ".png"

        def run():
            slide = self._slide(img)
            try:
                w = float(self.doc.PageSetup.SlideWidth)
                h = float(self.doc.PageSetup.SlideHeight)
                slide.Export(dest, "PNG", 1600, max(1, round(1600 * h / w)))
            except Exception as e:
                if _hresult(e) in _BUSY:
                    raise
                slide.Export(dest, "PNG")  # default size
            if not os.path.exists(dest):
                raise LiveWordError("exportfailed")
            return dest
        try:
            return _guard(run)
        except LiveWordError:
            raise
        except Exception as e:
            raise LiveWordError("exportfailed", repr(e))

    def goto(self, img):
        def run():
            window = self.doc.Windows.Item(1)
            window.Activate()
            window.View.GotoSlide(int(img.slide))
            if img.shape_name:
                shape = _find_shape(self._slide(img).Shapes, img.shape_name)
                if shape is not None:
                    try:
                        shape.Select()
                    except Exception:
                        pass  # the slide is shown; selecting the shape is a bonus
        _guard(run)
        return "ok"

    def _blocked_by(self):
        doc = self.doc
        for code, check in (("protectedview", lambda: in_protected_view(doc)),
                            ("readonly", lambda: bool(doc.ReadOnly)), ("final", lambda: bool(doc.Final))):
            try:
                if check():
                    return code
            except Exception:
                continue
        return None

    def set_alt(self, img, text):
        if not img.shape_name:
            raise LiveWordError("unsupported")

        def run():
            reason = self._blocked_by()
            if reason:
                raise LiveWordError(reason)
            shape = _find_shape(self._slide(img).Shapes, img.shape_name)
            if shape is None:
                raise LiveWordError("notfound")
            try:
                shape.AlternativeText = " ".join(text.split())
            except Exception as e:
                if _hresult(e) in _BUSY:
                    raise
                raise LiveWordError("denied", f"{e!r}; " + self._diagnostics(
                    ("ReadOnly", lambda: self.doc.ReadOnly),
                    ("Final", lambda: self.doc.Final),
                    ("shape.Type", lambda: shape.Type)))
        _guard(run)

    def get_alt(self, img):
        shape = _find_shape(self._slide(img).Shapes, img.shape_name) if img.shape_name else None
        return str(shape.AlternativeText) if shape is not None else None

    def _position(self):
        win = self.doc.Windows.Item(1)
        slide = int(win.View.Slide.SlideIndex)
        shape = ""
        try:
            sel = win.Selection
            if int(sel.Type) == 2:  # ppSelectionShapes
                shape = str(sel.ShapeRange.Item(1).Name)
        except Exception:
            pass
        return slide, shape

    def position_key(self):
        return self._position()

    def nearest_index(self, items):
        slide, shape = self._position()
        best = None
        for i, img in enumerate(items):
            if img.fmt != "pptx" or not img.slide:
                continue
            exact = img.slide == slide and shape and img.shape_name == shape
            d = abs(img.slide - slide) * 10 + (0 if exact else 1)
            if best is None or d < best[0]:
                best = (d, i)
        return best[1] if best else None

    def export_chart(self, img, dest):
        def run():
            if in_protected_view(self.doc):
                raise LiveWordError("protectedview")
            shape = _find_shape(self._slide(img).Shapes, img.shape_name)
            if shape is None:
                raise LiveWordError("notfound")
            try:
                shape.Export(dest, 2)  # ppShapeFormatPNG (an undocumented method)
            except Exception as e:
                if _hresult(e) in _BUSY:
                    raise
                shape.Chart.Export(dest, "PNG")  # the chart's own exporter
            if not os.path.exists(dest):
                raise LiveWordError("exportfailed")
        try:
            _guard(run)
        except LiveWordError:
            raise
        except Exception as e:
            raise LiveWordError("exportfailed", repr(e))
