import os
import types

import pytest

from engine import docimages, liveoffice, wordlayout, wordlive
from engine.wordlive import LiveWordError


# ---------------------------------------------------------------- fake Word
class Chart:
    def Export(self, path, kind):
        open(path, "wb").write(b"PNG")


class Shape:
    def __init__(self, rng, is_chart):
        self.Range, self.HasChart, self.Chart, self.AlternativeText = rng, is_chart, Chart(), "old alt"


class Rng:
    def __init__(self, text, kind, page, line, start):
        self.Text, self.page, self.line, self.Start = text, page, line, start
        self.shape = Shape(self, kind == "chart") if kind else None

    @property
    def InlineShapes(self):
        return types.SimpleNamespace(Count=1 if self.shape else 0, Item=lambda k: self.shape)

    def Information(self, n):
        return self.page if n == 3 else self.line


def word(protected_view=False, caret=95):
    paras = [Rng("Intro.\r", None, 1, 1, 0), Rng("Chart follows this line:\r", None, 1, 3, 10),
             Rng("\x01\r", "chart", 1, 5, 40), Rng("Picture after the chart: \x01\r", "pic", 2, 2, 100)]
    doc = types.SimpleNamespace(
        Paragraphs=types.SimpleNamespace(Count=4, Item=lambda j: types.SimpleNamespace(Range=paras[j - 1])),
        Saved=True, ReadOnly=False, ProtectionType=-1, Final=False, FullName="/t/a.docx",
        Content=types.SimpleNamespace(End=200),
        ActiveWindow=types.SimpleNamespace(Selection=types.SimpleNamespace(Start=caret)),
        Application=types.SimpleNamespace(ProtectedViewWindows=types.SimpleNamespace(
            Count=1 if protected_view else 0,
            Item=lambda i: types.SimpleNamespace(SourcePath="/t", SourceName="a.docx"))))
    doc.Range = lambda a, b: next((p for p in paras if p.Start == a), paras[0])
    doc.GoTo = lambda what, which, n: types.SimpleNamespace(Start={1: 0, 2: 100, 3: 100}[n])
    return doc, paras


@pytest.fixture
def items(samples):
    result = docimages.scan_images(samples["docx_chart"])
    return list(result), list(result.chart_items)


def test_pictures_and_charts_are_located_by_kind(items):
    (picture,), (chart,) = items
    doc, _ = word()
    assert wordlayout.map_images(doc, [picture, chart]) == 2
    assert (chart.page, chart.line, chart.char_start) == (1, 5, 40)
    assert (picture.page, picture.line, picture.char_start) == (2, 2, 100)


def test_a_picture_guessed_on_a_chart_paragraph_is_rejected(items):
    (picture,), _ = items
    doc, _ = word()
    picture.paragraph = 3  # the chart's paragraph
    wordlayout.map_images(doc, [picture])
    assert picture.char_start == 100


def test_word_chart_export_and_page_export(items, tmp_path):
    (picture,), (chart,) = items
    doc, _ = word()
    wordlayout.map_images(doc, [picture, chart])
    live = wordlive.LiveWord(doc, 1)
    assert live.can_export(chart) and not live.can_export(picture) and live.can_page(picture)
    dest = str(tmp_path / "c.png")
    live.export_chart(chart, dest)
    assert open(dest, "rb").read() == b"PNG"
    doc.Range = lambda a, b: types.SimpleNamespace(EnhMetaFileBits=tuple(b"EMF"), Start=a)
    emf = live.export_page(picture, str(tmp_path / "page"))
    assert emf.endswith(".emf") and open(emf, "rb").read() == b"EMF"
    with pytest.raises(LiveWordError) as e:
        live.export_page(docimages.DocImage(9, "m"), str(tmp_path / "x"))
    assert e.value.code == "nopage"


def test_word_refuses_when_the_document_changed(items, tmp_path):
    (picture,), (chart,) = items
    doc, _ = word()
    wordlayout.map_images(doc, [picture, chart])
    live = wordlive.LiveWord(doc, 1)
    doc.Content = types.SimpleNamespace(End=201)
    for action in (lambda: live.export_chart(chart, str(tmp_path / "c.png")),
                   lambda: live.export_page(picture, str(tmp_path / "p")), lambda: live.set_alt(picture, "x")):
        with pytest.raises(LiveWordError) as e:
            action()
        assert e.value.code in ("changed",)


def test_word_protected_view_wins_over_read_only(items):
    (picture,), _ = items
    doc, _ = word(protected_view=True)
    doc.ReadOnly = True
    live = wordlive.LiveWord(doc, 1)
    picture.char_start = 100
    with pytest.raises(LiveWordError) as e:
        live.set_alt(picture, "x")
    assert e.value.code == "protectedview"


def test_word_alt_text_is_read_and_written(items):
    (picture,), _ = items
    doc, _ = word()
    picture.char_start = 100
    live = wordlive.LiveWord(doc, 1)
    assert live.get_alt(picture) == "old alt"
    live.set_alt(picture, "A  new\n text")
    assert live.get_alt(picture) == "A new text"


def test_word_nearest_image_and_position(items):
    (picture,), (chart,) = items
    picture.char_start, chart.char_start = 100, 40
    live = wordlive.LiveWord(word(caret=45)[0], 1)
    assert live.position_key() == 45
    assert live.nearest_index([picture, chart]) == 1


# ---------------------------------------------------------------- fake Excel / PowerPoint
class PShape:
    def __init__(self, name, kind=13):
        self.Name, self.Type, self.AlternativeText = name, kind, "old alt"
        self.Chart = types.SimpleNamespace(Export=lambda path, kind: open(path, "wb").write(b"C"))

    def Export(self, path, kind):
        raise Exception("Shape.Export unavailable")

    def Select(self):
        pass


class Shapes:
    def __init__(self, shapes):
        self.shapes, self.Count = shapes, len(shapes)

    def Item(self, key):
        if isinstance(key, int):
            return self.shapes[key - 1]
        for shape in self.shapes:
            if shape.Name == key:
                return shape
        raise Exception("not found")


class Collection:
    def __init__(self, items):
        self.items, self.Count = items, len(items)

    def Item(self, key):
        if key in self.items:  # looked up by name, or by a key such as a slide number
            return self.items[key]
        return list(self.items.values())[key - 1]


def protected_view(active):
    return types.SimpleNamespace(
        Count=1 if active else 0, Item=lambda i: types.SimpleNamespace(SourcePath="/t", SourceName="a.xlsx"))


def workbook(pv=False, read_only=False, cell=(3, 2), sheet="Pictures"):
    pic, chart = PShape("Picture 1"), PShape("Chart 3", 3)
    ws = types.SimpleNamespace(Shapes=Shapes([pic, chart]), ProtectDrawingObjects=False, Activate=lambda: None,
                               Range=lambda a: types.SimpleNamespace(Select=lambda: None),
                               ChartObjects=lambda: types.SimpleNamespace(
                                   Item=lambda n: types.SimpleNamespace(Chart=chart.Chart, ShapeRange=types.SimpleNamespace(
                                       Item=lambda k: chart))))
    return types.SimpleNamespace(
        FullName="/t/a.xlsx", ReadOnly=read_only, Activate=lambda: None, Worksheets=Collection({"Pictures": ws}),
        Sheets=Collection({"Pictures": types.SimpleNamespace(Name="Pictures"),
                           "Other": types.SimpleNamespace(Name="Other")}),
        Charts=Collection({}), ActiveSheet=types.SimpleNamespace(Name=sheet),
        Application=types.SimpleNamespace(ProtectedViewWindows=protected_view(pv),
                                          ActiveCell=types.SimpleNamespace(Row=cell[0], Column=cell[1]))), ws


@pytest.fixture
def excel_items(samples):
    result = docimages.scan_images(samples["xlsx_floating"])
    return list(result) + list(result.chart_items)


def test_excel_alt_text_chart_export_and_refusals(excel_items, tmp_path):
    picture, _second, chart = excel_items
    book, ws = workbook()
    live = liveoffice.LiveExcel(book, 1)
    live.set_alt(picture, "New  alt")
    assert ws.Shapes.shapes[0].AlternativeText == "New alt"
    live.set_alt(chart, "Chart alt")
    assert live.get_alt(chart) == "Chart alt"
    live.export_chart(chart, str(tmp_path / "c.png"))
    assert (tmp_path / "c.png").exists()
    incell = docimages.DocImage(1, "m", fmt="xlsx", kind="incell", sheet="S", cell="A1")
    with pytest.raises(LiveWordError) as e:
        live.set_alt(incell, "x")
    assert e.value.code == "unsupported"
    with pytest.raises(LiveWordError) as e:
        live.export_page(picture, str(tmp_path / "p"))
    assert e.value.code == "nosheet"


def test_excel_protected_view_is_named_before_read_only(excel_items, tmp_path):
    picture, _second, chart = excel_items
    live = liveoffice.LiveExcel(workbook(pv=True, read_only=True)[0], 1)
    for action in (lambda: live.set_alt(picture, "x"), lambda: live.export_chart(chart, str(tmp_path / "c.png"))):
        with pytest.raises(LiveWordError) as e:
            action()
        assert e.value.code == "protectedview"


def test_excel_busy_application_is_reported(excel_items):
    picture = excel_items[0]
    book, ws = workbook()

    class Busy(Exception):
        hresult = -2146777998  # VBA_E_IGNORE: Excel is editing a cell

    def refuse(address):
        raise Busy()
    ws.Range = refuse
    with pytest.raises(LiveWordError) as e:
        liveoffice.LiveExcel(book, 1).goto(picture)
    assert e.value.code == "busy"


def test_excel_nearest_image_follows_the_active_cell(excel_items):
    shape_for = lambda live: excel_items[live.nearest_index(excel_items)].shape_name
    first = liveoffice.LiveExcel(workbook(cell=(3, 2))[0], 1)
    assert first.position_key() == ("Pictures", (3, 2)) and shape_for(first) == excel_items[0].shape_name
    near_chart = liveoffice.LiveExcel(workbook(cell=(14, 8))[0], 1)
    assert excel_items[near_chart.nearest_index(excel_items)].kind == "chart"


def test_powerpoint_goto_alt_slide_export_and_chart_fallback(samples, tmp_path):
    result = docimages.scan_images(samples["pptx"])
    picture, chart = result[1], result.chart_items[0]
    slide_shapes = Shapes([PShape("Group 1", 6), PShape(picture.shape_name), PShape(chart.shape_name, 3)])
    slide_shapes.shapes[0].GroupItems = Shapes([])
    exported = []
    slide = types.SimpleNamespace(Shapes=slide_shapes,
                                  Export=lambda path, kind, w=None, h=None: (exported.append((w, h)),
                                                                              open(path, "wb").write(b"S")))
    view = types.SimpleNamespace(shown=None, GotoSlide=lambda n: setattr(view, "shown", n), Slide=types.SimpleNamespace(SlideIndex=2))
    win = types.SimpleNamespace(Activate=lambda: None, View=view, Selection=types.SimpleNamespace(Type=0))
    presentation = types.SimpleNamespace(
        Slides=Collection({2: slide}), Windows=Collection({1: win}), ReadOnly=False, Final=False, FullName="/t/a.pptx",
        PageSetup=types.SimpleNamespace(SlideWidth=720, SlideHeight=540),
        Application=types.SimpleNamespace(ProtectedViewWindows=protected_view(False)))
    live = liveoffice.LivePowerPoint(presentation, 1)
    assert live.goto(picture) == "ok" and view.shown == 2
    live.set_alt(picture, "Alt")
    assert slide_shapes.shapes[1].AlternativeText == "Alt"
    live.export_chart(chart, str(tmp_path / "c.png"))  # Shape.Export fails, Chart.Export takes over
    assert (tmp_path / "c.png").read_bytes() == b"C"
    assert live.export_page(picture, str(tmp_path / "slide")).endswith(".png") and exported == [(1600, 1200)]
    presentation.Final = True
    with pytest.raises(LiveWordError) as e:
        live.set_alt(picture, "x")
    assert e.value.code == "final"
    assert live.nearest_index(list(result) + result.chart_items) in (1, 2)
