import types

import pytest

from engine import docimages
from engine.model import DocImage


def test_announcements(nvda, samples):
    package, _ = nvda
    result = docimages.scan_images(samples["xlsx_floating"])
    text = package._announcement("a.xlsx", result, result.charts, 0)
    assert text == "2 images, all with alt text. 1 chart, which is not an image."
    assert package._announcement("a.xlsx", [], 0, 0) == "No images found in a.xlsx"
    assert "could not be located" in package._announcement("a.xlsx", [], 0, 2)
    one = [DocImage(1, "m"), DocImage(2, "m", alt="x")]
    assert package._announcement("a", one) == "2 images, 1 without alt text"
    assert package._announcement("a", one[:1]) == "1 image, 1 without alt text"  # singular form


def test_every_cell_describes_itself(nvda, samples):
    package, _ = nvda
    result = docimages.scan_images(samples["xlsx_incell"])
    row = "; ".join(package._cell(result[0], c) for c in range(4))
    assert row == "Image 1; Sheet InCell, cell A1, in cell; Alt text: Green triangle in a cell; "
    chart = docimages.scan_images(samples["xlsx_floating"]).chart_items[0]
    assert package._cell(chart, 0) == "Chart 1" and package._cell(chart, 3) == "Title: Quarterly sales"


def test_word_locations(nvda):
    package, _ = nvda
    image = DocImage(1, "m", page_est=3)
    assert package._location(image) == "Page about 3"
    image.page, image.line = 4, 12
    assert package._location(image) == "Page 4, line 12"
    assert package._location(DocImage(2, "m", part="header")) == "In header"


@pytest.mark.parametrize("entry, key, nearest, expected", [
    ({}, 100, 4, 4),                                       # first time: nearest to the position
    ({"last_index": 2, "last_key": 100}, 100, 4, 2),       # nothing moved: the row left before
    ({"last_index": 2, "last_key": 100}, 500, 5, 5),       # the user moved: nearest again
    ({"last_index": 3}, None, None, 3),                    # opened from Explorer: the row left before
    ({}, None, None, 0),
    ({"last_index": 99}, None, None, 0),                   # stale index
])
def test_starting_row(nvda, entry, key, nearest, expected):
    package, _ = nvda
    live = None if key is None and nearest is None else types.SimpleNamespace(
        position_key=lambda: key, nearest_index=lambda items: nearest)
    assert package._initial_index(entry, list(range(7)), live) == expected


def test_the_focus_goes_back_to_the_list_after_the_result_is_spoken(nvda):
    package, spoken = nvda
    events = []

    class Dialog:
        listCtrl = types.SimpleNamespace(SetFocus=lambda: events.append("focus"))
        _say = package.DocImagesDialog._say

        @package._focus_list_after
        def action(self, evt):
            self._say("Alt text updated.")

        @package._focus_list_after
        def fails(self, evt):
            raise ValueError("boom")
    import ui
    ui.message = lambda text: events.append(text)
    Dialog().action(None)
    assert events == ["Alt text updated.", "focus"]
    events.clear()
    with pytest.raises(ValueError):
        Dialog().fails(None)
    assert events == ["focus"]  # the list still gets the focus back


def test_the_store_is_offered_when_be_my_eyes_is_missing(nvda, monkeypatch):
    package, spoken = nvda
    opened = []
    monkeypatch.setattr(package.bemyeyes, "open_store", lambda: opened.append(True) or True)
    package._offer_store()  # the stand-in message box answers Yes
    assert opened == [True]
