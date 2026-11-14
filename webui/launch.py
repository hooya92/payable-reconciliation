"""Open the web screen in a native window (pywebview + the Windows WebView2 engine).

`run()` returns False when the window cannot be created (no WebView2, pywebview missing, ...),
so the caller can fall back to the classic screen instead of leaving the user with nothing.
"""
from __future__ import annotations

import os
import sys
import threading
import time
from pathlib import Path

from application.workflow import open_saved_folder
from webui.api import Api
from webui.session import Session


def static_dir() -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
    return base / "webui" / "static"


class WebviewDialogs:
    """Native file dialogs through the pywebview window."""

    def __init__(self, webview):
        self.webview = webview
        self.window = None
        kind = getattr(webview, "FileDialog", None)
        self._open = kind.OPEN if kind else webview.OPEN_DIALOG
        self._save = kind.SAVE if kind else webview.SAVE_DIALOG

    def open_files(self):
        picked = self.window.create_file_dialog(
            self._open, allow_multiple=True, file_types=("Excel (*.xlsx;*.xlsm;*.xls)",))
        return [str(p) for p in (picked or [])]

    def save_file(self, name):
        picked = self.window.create_file_dialog(
            self._save, save_filename=name, file_types=("Excel (*.xlsx)",))
        if not picked:
            return None
        path = str(picked if isinstance(picked, str) else picked[0])
        return path if path.lower().endswith(".xlsx") else path + ".xlsx"

    def open_folder(self, path):
        open_saved_folder(path)


def _selftest(window, out_path):
    """Used by CI: prove the page loaded and can talk to Python, then close."""
    deadline = time.time() + 90
    verdict = "FAIL: page did not load"
    while time.time() < deadline:
        try:
            ok = window.evaluate_js(
                "!!document.querySelector('#app h1') && typeof window.pywebview.api.get_state === 'function'")
            if ok:
                verdict = "WEBUI_OK"
                break
        except Exception as e:  # the engine may not be ready yet
            verdict = f"FAIL: {e}"
        time.sleep(0.5)
    if out_path:
        Path(out_path).write_text(verdict, encoding="utf-8")
    window.destroy()


def run(selftest: bool = False) -> bool:
    try:
        import webview
    except Exception:
        return False
    try:
        dialogs = WebviewDialogs(webview)
        holder = {}
        session = Session(dialogs, quit_callback=lambda: holder["window"].destroy(),
                          minimize_callback=lambda: holder["window"].minimize(),
                          drag_callback=lambda mode: holder["mover"].begin(mode))
        options = dict(
            url=str(static_dir() / "index.html"), js_api=Api(session), width=470, height=700,
            min_size=(440, 560), background_color="#F5F5F7",
            # No Windows title bar: the page draws its own header, drag area and window buttons.
            frameless=True, easy_drag=False, resizable=True,
        )
        import inspect
        if "shadow" in inspect.signature(webview.create_window).parameters:
            options["shadow"] = True
        window = webview.create_window("명세서 대사", **options)
        holder["window"] = dialogs.window = window
        from webui.winmove import WindowMover
        holder["mover"] = WindowMover(window)
        from adapters import updater
        if updater.can_self_update():
            threading.Thread(target=updater.cleanup_stale, args=(Path(sys.executable),), daemon=True).start()
        if selftest:
            webview.start(_selftest, (window, os.environ.get("PAYABLE_SELFTEST_OUT", "")), gui="edgechromium")
        else:
            webview.start(gui="edgechromium" if sys.platform == "win32" else None)
        return True
    except Exception:
        return False
