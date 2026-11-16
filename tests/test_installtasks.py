import importlib.util
import json
import os
import queue
import subprocess
import sys
import threading
import types

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@pytest.fixture
def install(monkeypatch, tmp_path):
    def add(name, **attrs):
        module = types.ModuleType(name)
        module.__dict__.update(attrs)
        monkeypatch.setitem(sys.modules, name, module)

    def no_translations():
        raise RuntimeError("translations are not available while the add-on is pending install")

    pending = queue.Queue()
    add("addonHandler", initTranslation=no_translations)
    add("globalVars", appArgs=types.SimpleNamespace(configPath=str(tmp_path)))
    add("gui", mainFrame=object())
    add("ui", message=lambda text: None)
    add("wx", Dialog=object, CallAfter=pending.put)
    add("logHandler", log=types.SimpleNamespace(debugWarning=lambda *a, **k: None, error=lambda *a, **k: None,
                                                warning=lambda *a, **k: None))
    spec = importlib.util.spec_from_file_location("installTasks", os.path.join(ROOT, "addon", "installTasks.py"))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.pending = pending
    return module


def fake_powershell(monkeypatch, stdout="", returncode=0, missing=False):
    def run(*args, **kwargs):
        if missing:
            raise FileNotFoundError("powershell")
        return types.SimpleNamespace(stdout=stdout, returncode=returncode)
    monkeypatch.setattr(subprocess, "run", run)


def test_translations_fall_back_to_the_english_text(install):
    assert install._("Check again") == "Check again"


@pytest.mark.parametrize("stdout, returncode, missing, expected", [
    ("A25A.BeMyEyes_8wekyb3d8bbwe!App\n", 0, False, ("found", "A25A.BeMyEyes_8wekyb3d8bbwe!App")),
    ("", 0, False, ("missing", None)),
    ("", 1, False, ("unknown", None)),
    ("", 0, True, ("unknown", None)),
])
def test_finding_be_my_eyes(install, monkeypatch, stdout, returncode, missing, expected):
    fake_powershell(monkeypatch, stdout, returncode, missing)
    assert install._find_bemyeyes() == expected


def decide(install, monkeypatch, status, app_id=None, config=None):
    shown = []
    if config is not None:
        install._save_config(**config)
    monkeypatch.setattr(install, "_find_bemyeyes", lambda: (status, app_id))
    monkeypatch.setattr(install, "_run_on_main", lambda fn: fn())
    monkeypatch.setattr(install, "_show_dialog", lambda: shown.append(True))
    install.onInstall()
    return bool(shown), install._load_config()


def test_nothing_is_asked_when_the_app_is_found_and_its_id_is_kept(install, monkeypatch):
    shown, config = decide(install, monkeypatch, "found", "A!App", config={"other": 1})
    assert not shown and config == {"other": 1, "detectedAumid": "A!App"}


def test_the_user_is_asked_when_the_app_is_missing(install, monkeypatch):
    assert decide(install, monkeypatch, "missing")[0]


def test_an_id_written_by_hand_means_no_question_and_is_left_alone(install, monkeypatch):
    shown, config = decide(install, monkeypatch, "missing", config={"aumid": "My.App_abc!App"})
    assert not shown and config == {"aumid": "My.App_abc!App"}


def test_nobody_is_asked_when_reminders_are_off_or_the_check_failed(install, monkeypatch):
    assert not decide(install, monkeypatch, "missing", config={"skipBeMyEyesCheck": True})[0]
    assert not decide(install, monkeypatch, "unknown")[0]


def test_dialogs_run_on_the_main_thread_without_hanging(install):
    results = {}

    def worker():
        results["ok"] = install._run_on_main(lambda: "answer")
        results["failed"] = install._run_on_main(lambda: 1 / 0)  # must not block the installation forever

    thread = threading.Thread(target=worker)
    thread.start()
    for _ in range(2):
        install.pending.get(timeout=5)()  # the main thread plays the user interface loop
    thread.join(5)
    assert not thread.is_alive() and results == {"ok": "answer", "failed": None}
    assert install._run_on_main(lambda: "direct") == "direct"  # already on the main thread


def test_the_store_link_falls_back_to_the_web_page(install, monkeypatch):
    opened = []

    def start(target):
        opened.append(target)
        if target.startswith("ms-windows-store"):
            raise OSError("no handler")
    monkeypatch.setattr(os, "startfile", start, raising=False)
    assert install._open_store() and opened == [install._STORE_URI, install._STORE_WEB]
