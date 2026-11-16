import inspect
import os
import sys
import types

import pytest

sys.dont_write_bytecode = True

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PACKAGE_DIR = os.path.join(ROOT, "addon", "globalPlugins", "docImages")
sys.path.insert(0, PACKAGE_DIR)  # the engine has no NVDA imports and is tested as a plain package
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


@pytest.fixture(scope="session")
def samples(tmp_path_factory):
    import samples as builders
    return builders.build_all(tmp_path_factory.mktemp("samples"))


class _Anything:
    def __getattr__(self, name):
        return lambda *a, **k: None


@pytest.fixture
def nvda(monkeypatch, tmp_path):
    """Import the add-on package against stand-ins for the NVDA modules. Returns (package, spoken)."""
    spoken = []

    def add(name, **attrs):
        module = types.ModuleType(name)
        module.__dict__.update(attrs)
        monkeypatch.setitem(sys.modules, name, module)
        return module

    def init_translation():
        inspect.currentframe().f_back.f_globals["_"] = lambda text: text

    class Plain:
        def __init__(self, *a, **k):
            pass

    add("addonHandler", initTranslation=init_translation)
    add("api")
    add("globalPluginHandler", GlobalPlugin=Plain)
    helper = add("gui.guiHelper", BoxSizerHelper=object, ButtonHelper=object, BORDER_FOR_DIALOGS=5)
    add("gui", guiHelper=helper, mainFrame=None)
    add("scriptHandler", script=lambda **k: (lambda f: f), getLastScriptRepeatCount=lambda: 0)
    add("tones", beep=lambda *a: None)
    add("ui", message=spoken.append)
    add("winUser")
    class Menu:
        def __init__(self):
            self.labels = []

        def Append(self, wx_id, label):
            self.labels.append(label)
            return types.SimpleNamespace(label=label)

        def Destroy(self):
            pass

    add("wx", Dialog=Plain, ListCtrl=Plain, CallAfter=lambda f, *a: f(*a), ID_CLOSE=1, ID_ANY=-1, LC_REPORT=1,
        LC_VIRTUAL=2, LC_SINGLE_SEL=4, YES_NO=1, ICON_QUESTION=2, YES=3, MessageBox=lambda *a, **k: 3,
        Menu=Menu, Point=lambda x, y: (x, y), EVT_MENU=object())
    add("logHandler", log=_Anything())
    add("globalVars", appArgs=types.SimpleNamespace(configPath=str(tmp_path)))
    add("core", callLater=lambda ms, f, *a: f(*a))
    monkeypatch.syspath_prepend(os.path.join(ROOT, "addon"))
    for name in [n for n in sys.modules if n.startswith("globalPlugins")]:
        monkeypatch.delitem(sys.modules, name)
    import importlib
    package = importlib.import_module("globalPlugins.docImages")
    yield package, spoken
    for name in [n for n in sys.modules if n.startswith("globalPlugins")]:
        sys.modules.pop(name, None)
