"""Development bridge: serve the page over HTTP and route /api/<name> to the same Api.

    python -m webui.server            # then open http://127.0.0.1:8765/?dev=1

File dialogs are scripted (POST /__dialogs__ {"open": [...], "save": "..."}) so the page can be
driven end to end in a normal browser, e.g. by Playwright.
"""
from __future__ import annotations

import json
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from webui.api import Api
from webui.launch import static_dir
from webui.session import Session

TYPES = {".html": "text/html; charset=utf-8", ".css": "text/css; charset=utf-8", ".js": "text/javascript; charset=utf-8"}


class ScriptedDialogs:
    def __init__(self):
        self.open, self.save, self.folders = [], None, []

    def open_files(self):
        return list(self.open)

    def save_file(self, name):
        return self.save

    def open_folder(self, path):
        self.folders.append(path)


def make_server(port=8765, session_factory=None):
    dialogs = ScriptedDialogs()
    session = (session_factory or Session)(dialogs)
    api = Api(session)
    root = static_dir()

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass

        def _send(self, code, body, ctype="application/json"):
            data = body if isinstance(body, bytes) else json.dumps(body, ensure_ascii=False).encode()
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(data)

        def do_GET(self):
            name = self.path.split("?")[0].lstrip("/") or "index.html"
            target = (root / name).resolve()
            if root.resolve() not in target.parents or not target.is_file():
                return self._send(404, {"error": "not found"})
            self._send(200, target.read_bytes(), TYPES.get(target.suffix, "application/octet-stream"))

        def do_POST(self):
            body = self.rfile.read(int(self.headers.get("Content-Length") or 0)) or b"[]"
            try:
                args = json.loads(body)
                if self.path == "/__dialogs__":
                    dialogs.open = args.get("open", dialogs.open)
                    dialogs.save = args.get("save", dialogs.save)
                    return self._send(200, {"ok": True})
                name = self.path.removeprefix("/api/")
                if name.startswith("_") or not hasattr(api, name):
                    return self._send(404, {"error": "unknown method"})
                self._send(200, getattr(api, name)(*args))
            except Exception as e:
                self._send(500, {"error": str(e)})

    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    server.session, server.dialogs = session, dialogs
    return server


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
    srv = make_server(port)
    print(f"http://127.0.0.1:{port}/?dev=1", flush=True)
    srv.serve_forever()
