"""Exact page/line numbers for docx images, using Word's own layout in a hidden Word instance.
No NVDA imports. Call annotate() from a worker thread."""
import gc
import re

WD_ACTIVE_END_PAGE = 3        # wdActiveEndPageNumber
WD_FIRST_CHAR_LINE = 10       # wdFirstCharacterLineNumber
NEIGHBOURS = (0, -1, 1, -2, 2, -3, 3)


def _norm(s):
    return re.sub(r"[\s\x01\x07\x0b\x0c]+", " ", s or "").strip().lower()


def _object_at(rng, k):
    """(anchor range, shape) of the k-th picture-like object in rng: inline ones first, then floating ones."""
    try:
        n = rng.InlineShapes.Count
        if n > 0:
            shape = rng.InlineShapes.Item(min(k, n - 1) + 1)
            return shape.Range, shape
    except Exception:
        pass
    try:
        sr = rng.ShapeRange
        c = sr.Count
        if c > 0:
            shape = sr.Item(min(k, c - 1) + 1)
            return shape.Anchor, shape
    except Exception:
        pass
    return None, None


def _is_chart(shape):
    """True or False, or None when Word does not say."""
    try:
        return bool(shape.HasChart)
    except Exception:
        return None


def map_images(doc, images):
    """Fill page/line/char_start of body pictures and charts. A paragraph guess is only accepted if Word's
    paragraph holds an object of the same kind (picture or chart) and, for body text, starts with the same
    text. Returns how many items were located."""
    paras = doc.Paragraphs
    total = paras.Count
    located = 0
    for img in images:
        if img.fmt != "docx" or img.part != "body":
            continue
        k = img.slot
        want = _norm(img.context)[:15]
        want_chart = img.kind == "chart"
        anchor = None
        for strict in (True, False):
            for d in (NEIGHBOURS if strict else (0,)):
                j = img.paragraph + d
                if not 1 <= j <= total:
                    continue
                try:
                    rng = paras.Item(j).Range
                    cand, shape = _object_at(rng, k)
                    if cand is None:
                        continue
                    is_chart = _is_chart(shape)
                    if (want_chart and is_chart is False) or (not want_chart and is_chart is True):
                        continue
                    if strict and want and not img.in_textbox and want not in _norm(rng.Text):
                        continue
                    anchor = cand
                    break
                except Exception:
                    continue
            if anchor is not None:
                break
        if anchor is None:
            continue
        try:
            page = int(anchor.Information(WD_ACTIVE_END_PAGE))
            line = int(anchor.Information(WD_FIRST_CHAR_LINE))
        except Exception:
            continue
        try:
            img.char_start = int(anchor.Start)
        except Exception:
            pass
        if page > 0:
            img.page = page
            located += 1
        if line > 0:
            img.line = line
    return located


def annotate(docx_path, images):
    if not any(i.fmt == "docx" and i.part == "body" for i in images):
        return False
    import comtypes
    import comtypes.client
    comtypes.CoInitialize()
    app = doc = None
    try:
        app = comtypes.client.CreateObject("Word.Application", dynamic=True)
        app.Visible = False
        app.DisplayAlerts = 0
        app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable: never run macros
        doc = app.Documents.Open(docx_path, False, True, False, "x")  # no convert prompt, read-only, not in MRU
        try:
            doc.ActiveWindow.View.Type = 3  # print layout: needed for line numbers
        except Exception:
            pass
        return map_images(doc, images) > 0
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
