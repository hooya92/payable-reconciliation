"""The Api wrapper and the development HTTP bridge (what the page really calls)."""
import json
import threading
import unittest
import urllib.request
from pathlib import Path

from webui.api import Api, _PUBLIC
from webui.server import make_server
from webui.session import Session
from webui import launch


class FakeDialogs:
    def open_files(self): return []
    def save_file(self, name): return None
    def open_folder(self, path): pass


class ApiTests(unittest.TestCase):
    def test_only_the_listed_methods_are_public(self):
        api = Api(Session(FakeDialogs()))
        public = {n for n in dir(api) if not n.startswith("_")}
        self.assertEqual(public, set(_PUBLIC))      # no session / lock attribute for pywebview to walk

    def test_unknown_file_kinds_are_rejected(self):
        api = Api(Session(FakeDialogs()))
        with self.assertRaises(ValueError):
            api.add_files("../../etc")
        with self.assertRaises(ValueError):
            api.remove_file("x", 0)

    def test_window_buttons_reach_the_window(self):
        calls = []
        api = Api(Session(FakeDialogs(), quit_callback=lambda: calls.append("quit"),
                          minimize_callback=lambda: calls.append("min")))
        api.minimize_window(); api.close_window()
        self.assertEqual(calls, ["min", "quit"])

    def test_drag_requests_reach_the_window_and_bad_modes_are_refused(self):
        seen = []
        api = Api(Session(FakeDialogs(), drag_callback=lambda mode: seen.append(mode) or True))
        self.assertTrue(api.begin_drag("move")); self.assertTrue(api.begin_drag("se"))
        self.assertEqual(seen, ["move", "se"])
        with self.assertRaises(ValueError):
            api.begin_drag("../x")

    def test_state_is_plain_json(self):
        json.dumps(Api(Session(FakeDialogs())).get_state())


class BridgeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = make_server(0)
        cls.port = cls.server.server_address[1]
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown(); cls.server.server_close()

    def post(self, path, *args):
        req = urllib.request.Request(f"http://127.0.0.1:{self.port}{path}", data=json.dumps(list(args)).encode(),
                                     headers={"Content-Type": "application/json"}, method="POST")
        return json.load(urllib.request.urlopen(req))

    def test_page_assets_are_served(self):
        for name, marker in (("", b'id="app"'), ("app.js", b"pywebviewready"), ("app.css", b"--accent")):
            body = urllib.request.urlopen(f"http://127.0.0.1:{self.port}/{name}").read()
            self.assertIn(marker, body)

    def test_api_calls_round_trip(self):
        state = self.post("/api/get_state")
        self.assertEqual(state["period"]["detect"], "none")
        self.assertFalse(self.post("/api/toggle_picker")["period"]["picker_open"] is False)

    def test_private_and_unknown_methods_and_path_escapes_are_refused(self):
        for path in ("/api/_session", "/api/nope", "/api/state"):
            with self.assertRaises(urllib.error.HTTPError) as cm:
                self.post(path)
            self.assertEqual(cm.exception.code, 404)
        with self.assertRaises(urllib.error.HTTPError):
            urllib.request.urlopen(f"http://127.0.0.1:{self.port}/../session.py")


class PageSafetyTests(unittest.TestCase):
    def test_page_escapes_every_value_it_prints(self):
        js = (Path(launch.static_dir()) / "app.js").read_text(encoding="utf-8")
        self.assertIn("function esc(", js)
        # file names and vendor names come from Excel: they must never be inserted unescaped
        for risky in ("${n}", "${x.vendor}", "${x.reason}", "${s.files"):
            self.assertNotIn(risky, js)

    def test_static_dir_has_the_page(self):
        for name in ("index.html", "app.js", "app.css", "rabbit.png", "fonts/PretendardVariable.woff2",
                     "fonts/PRETENDARD-LICENSE.txt"):
            self.assertTrue((launch.static_dir() / name).is_file(), name)

    def test_the_title_row_has_the_rabbit_and_the_window_has_no_native_title_bar(self):
        js = (launch.static_dir() / "app.js").read_text(encoding="utf-8")
        self.assertIn('src="rabbit.png"', js)
        self.assertIn('id="dragbar"', js)                # the page provides the drag area
        self.assertIn("begin_drag", js)
        css = (launch.static_dir() / "app.css").read_text(encoding="utf-8")
        import re
        height = int(re.search(r"\.dragbar\{[^}]*height:(\d+)px", css).group(1))
        self.assertGreaterEqual(height, 90)               # the whole title area, not a thin strip
        src = (Path(launch.__file__)).read_text(encoding="utf-8")
        self.assertIn("frameless=True", src)
        self.assertTrue((launch.static_dir() / "rabbit.png").read_bytes().startswith(b"\x89PNG"))

    def test_the_font_is_served_with_its_license_beside_it(self):
        from webui.server import make_server
        server = make_server(0)
        threading.Thread(target=server.serve_forever, daemon=True).start()
        try:
            port = server.server_address[1]
            resp = urllib.request.urlopen(f"http://127.0.0.1:{port}/fonts/PretendardVariable.woff2")
            self.assertEqual(resp.headers["Content-Type"], "font/woff2")
            self.assertGreater(len(resp.read()), 1_000_000)
        finally:
            server.shutdown(); server.server_close()


if __name__ == "__main__":
    unittest.main()
