import json
import queue
import types

import pytest


@pytest.fixture
def bme(nvda, monkeypatch):
    package, spoken = nvda
    module = package.bemyeyes
    clock = types.SimpleNamespace(now=1000.0)
    monkeypatch.setattr(module.time, "monotonic", lambda: clock.now)
    searches = []

    def script(*results):
        outcomes = list(results)

        def search():
            searches.append(1)
            return outcomes.pop(0) if len(outcomes) > 1 else outcomes[0]
        monkeypatch.setattr(module, "_search", search)
    module.script, module.clock, module.searches = script, clock, searches
    return module


def test_an_id_written_by_hand_counts_as_installed_without_any_search(bme):
    bme.script(("missing", None))
    bme._save({"aumid": "My.App_abc!App"})
    assert bme.availability() == ("found", "My.App_abc!App") and not bme.searches


def test_the_result_is_remembered_for_a_while_then_searched_again(bme):
    bme.script(("found", "A.BeMyEyes_x!App"))
    assert bme.availability() == ("found", "A.BeMyEyes_x!App")
    bme.clock.now += 500
    bme.availability()
    assert len(bme.searches) == 1
    bme.clock.now += 200  # past ten minutes
    bme.availability()
    assert len(bme.searches) == 2
    assert bme._load()["detectedAumid"] == "A.BeMyEyes_x!App" and "aumid" not in bme._load()


def test_a_missing_app_is_searched_again_soon_so_installing_it_is_noticed(bme):
    bme.script(("missing", None), ("found", "A.BeMyEyes_x!App"))
    assert bme.is_available() is False
    bme.clock.now += 10
    assert bme.is_available() is False and len(bme.searches) == 1
    bme.clock.now += 30
    assert bme.is_available() is True and len(bme.searches) == 2


def test_when_the_search_cannot_run_the_button_stays(bme):
    bme.script(("unknown", None))
    assert bme.is_available() is True
    bme.is_available()
    assert len(bme.searches) == 2  # a failed search is never remembered


def sent(bme, monkeypatch, *, status, launched=True, pillow=True):
    """Run send_async and return what it reported."""
    answers = queue.Queue()
    monkeypatch.setattr(bme, "send_to_bemyeyes", lambda path, aumid: launched)
    # The add-on imports its engine as globalPlugins.docImages.engine, so the error class must come from there.
    docimages = __import__("globalPlugins.docImages.engine.docimages", fromlist=["docimages"])
    add_on_error = __import__("globalPlugins.docImages.engine.model", fromlist=["model"]).DocImageError

    def prepare(path):
        if not pillow:
            raise add_on_error("nopillow")
        return path
    monkeypatch.setattr(docimages, "prepare_for_sharing", prepare)
    bme.send_async("/t/a.png", answers.put)
    return answers.get(timeout=5)


def test_sending_reports_what_happened(bme, monkeypatch):
    bme.script(("found", "A.BeMyEyes_x!App"))
    assert sent(bme, monkeypatch, status="found") == "ok"
    assert sent(bme, monkeypatch, status="found", pillow=False) == "nopillow"
    bme.script(("missing", None))
    bme._memo.clear()
    assert sent(bme, monkeypatch, status="missing") == "notfound"


def test_a_failed_launch_is_told_apart_from_an_uninstalled_app(bme, monkeypatch):
    bme.script(("found", "A.BeMyEyes_x!App"), ("found", "A.BeMyEyes_x!App"))
    assert sent(bme, monkeypatch, status="found", launched=False) == "failed"
    bme.script(("found", "A.BeMyEyes_x!App"), ("missing", None))
    bme._memo.clear()
    assert sent(bme, monkeypatch, status="found", launched=False) == "notfound"


def test_the_send_item_is_left_out_of_the_page_menu_when_be_my_eyes_is_missing(nvda, samples):
    package, spoken = nvda
    from engine.model import DocImage

    class Live:
        def export_page(self, img, base):
            open(base + ".png", "wb").write(b"x")
            return base + ".png"

    class Button:
        def GetPosition(self):
            return types.SimpleNamespace(x=0, y=0)

        def GetSize(self):
            return types.SimpleNamespace(height=10)

    def labels(can_send, tmp):
        menus = []

        class Dialog:
            live, pageBtn, chartDir, canSend = Live(), Button(), str(tmp), can_send
            listCtrl = types.SimpleNamespace(SetFocus=lambda: None)
            _say = lambda self, text: spoken.append(text)
            _current = lambda self: DocImage(1, "m", fmt="pptx", slide=2)
            Bind = lambda self, ev, handler, item: None
            PopupMenu = lambda self, menu, pos: menus.append(list(menu.labels))
            onPage = package.DocImagesDialog.onPage
        Dialog().onPage(None)
        return menus[0]
    import tempfile
    with tempfile.TemporaryDirectory() as folder:
        assert labels(True, folder)[-1] == "Send to &Be My Eyes"
        assert "Send to &Be My Eyes" not in labels(False, folder) and len(labels(False, folder)) == 3
