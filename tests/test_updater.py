import io
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from adapters import updater

SHA = "a" * 40
PAYLOAD = {
    "body": f"PayableRecon 최신 정식 배포입니다. Build: {SHA}.",
    "assets": [{"name": "PayableRecon.exe", "browser_download_url": "https://x/PayableRecon.exe", "size": 6, "digest": ""}],
}


class FakeResp(io.BytesIO):
    headers = {"Content-Length": "6"}
    def __enter__(self): return self
    def __exit__(self, *a): return False


class UpdaterTests(unittest.TestCase):
    def test_parse_release(self):
        r = updater.parse_release(PAYLOAD)
        self.assertEqual((r.sha, r.size), (SHA, 6))
        self.assertIsNone(updater.parse_release({"body": "no build info", "assets": PAYLOAD["assets"]}))
        self.assertIsNone(updater.parse_release({"body": PAYLOAD["body"], "assets": []}))

    def test_release_date_is_read_for_people(self):
        body=f"PayableRecon 최신 정식 배포입니다. Build: {SHA}. Date: 2026-11-13 14:05 KST."
        r=updater.parse_release({**PAYLOAD,"body":body})
        self.assertEqual(r.date,"2026-11-13 14:05")
        self.assertEqual(updater.parse_release(PAYLOAD).date,"")   # old notes without a date still work
        self.assertEqual(updater.version_label(),"개발 버전")      # source run

    def test_is_newer(self):
        self.assertFalse(updater.is_newer("dev", SHA))
        self.assertFalse(updater.is_newer(SHA, SHA))
        self.assertFalse(updater.is_newer(SHA[:7], SHA))
        self.assertTrue(updater.is_newer("b" * 40, SHA))

    def test_check_latest_never_raises_and_only_reports_newer(self):
        with mock.patch.object(updater.urllib.request, "urlopen", side_effect=OSError("offline")):
            self.assertIsNone(updater.check_latest())
        body = io.BytesIO(__import__("json").dumps(PAYLOAD).encode())
        with mock.patch.object(updater.urllib.request, "urlopen", return_value=FakeResp(body.getvalue())), \
             mock.patch.object(updater, "current_sha", return_value="b" * 40):
            self.assertEqual(updater.check_latest().sha, SHA)
        with mock.patch.object(updater.urllib.request, "urlopen", return_value=FakeResp(body.getvalue())), \
             mock.patch.object(updater, "current_sha", return_value=SHA):
            self.assertIsNone(updater.check_latest())

    def _download(self, data, release):
        with tempfile.TemporaryDirectory() as d, \
             mock.patch.object(updater.urllib.request, "urlopen", return_value=FakeResp(data)):
            dest = Path(d) / "u.exe"
            updater.download(release, dest)
            return dest.read_bytes()

    def test_download_verifies_size_digest_and_exe_header(self):
        import hashlib
        data = b"MZ1234"
        ok = updater.Release(SHA, "https://x/u", 6, "sha256:" + hashlib.sha256(data).hexdigest())
        self.assertEqual(self._download(data, ok), data)
        with self.assertRaises(ValueError):
            self._download(data, updater.Release(SHA, "https://x/u", 7, ""))
        with self.assertRaises(ValueError):
            self._download(data, updater.Release(SHA, "https://x/u", 6, "sha256:" + "0" * 64))
        with self.assertRaises(ValueError):
            self._download(b"<html>", updater.Release(SHA, "https://x/u", 6, ""))

    def test_source_run_never_self_updates(self):
        self.assertFalse(updater.can_self_update())

    def test_swap_script_replaces_then_relaunches(self):
        s = updater.swap_script(Path(r"C:\한글 폴더\PayableRecon.exe"), Path(r"C:\한글 폴더\PayableRecon.update.exe"))
        self.assertLess(s.index("move /y"), s.index('start ""'))
        self.assertIn("PYINSTALLER_RESET_ENVIRONMENT=1", s)
        self.assertIn("한글 폴더", s)
        # giving up must not leave the staged exe or the script behind
        tail = s[s.index(":giveup"):]
        self.assertIn("PayableRecon.update.exe", tail)
        self.assertIn('del "%~f0"', tail)


if __name__ == "__main__":
    unittest.main()
