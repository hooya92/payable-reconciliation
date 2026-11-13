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

    def test_unchanged_release_answers_304_and_keeps_the_known_one(self):
        import urllib.error
        updater._cache.update(etag="",release=None)
        first=FakeResp(__import__("json").dumps(PAYLOAD).encode()); first.headers={"ETag":'"abc"'}
        sent=[]
        def fake(req,timeout=0):
            sent.append(req.get_header("If-none-match"))
            if len(sent)==1: return first
            raise urllib.error.HTTPError(req.full_url,304,"Not Modified",{},None)
        with mock.patch.object(updater.urllib.request,"urlopen",side_effect=fake):
            a=updater.fetch_release(); b=updater.fetch_release()
        self.assertEqual(a,b); self.assertEqual(sent,[None,'"abc"'])
        updater._cache.update(etag="",release=None)

    def test_source_run_never_self_updates(self):
        self.assertFalse(updater.can_self_update())

    def test_finish_update_replaces_when_old_exe_is_free_and_relaunches(self):
        with tempfile.TemporaryDirectory() as d:
            exe=Path(d)/"PayableRecon.exe"; staged=Path(d)/"PayableRecon.update.exe"
            exe.write_bytes(b"old"); staged.write_bytes(b"MZnew")
            launched=[]
            ok,msg=updater.finish_update(exe,staged,sleep=lambda s:None,launch=lambda:launched.append(1))
            self.assertTrue(ok,msg)
            self.assertEqual(exe.read_bytes(),b"MZnew")
            self.assertEqual(launched,[1])
            self.assertFalse((Path(d)/"PayableRecon.exe.new").exists())

    def test_finish_update_waits_for_the_old_exe_to_be_released(self):
        with tempfile.TemporaryDirectory() as d:
            exe=Path(d)/"PayableRecon.exe"; staged=Path(d)/"PayableRecon.update.exe"
            exe.write_bytes(b"old"); staged.write_bytes(b"MZnew")
            real=updater.os.replace; calls=[]
            def locked_twice(src,dst):
                calls.append(1)
                if len(calls)<=2: raise PermissionError("in use")
                return real(src,dst)
            with mock.patch.object(updater.os,"replace",side_effect=locked_twice):
                ok,msg=updater.finish_update(exe,staged,sleep=lambda s:None,launch=lambda:None)
            self.assertTrue(ok,msg); self.assertEqual(len(calls),3)

    def test_finish_update_gives_up_cleanly_and_keeps_the_old_exe(self):
        with tempfile.TemporaryDirectory() as d:
            exe=Path(d)/"PayableRecon.exe"; staged=Path(d)/"PayableRecon.update.exe"
            exe.write_bytes(b"old"); staged.write_bytes(b"MZnew")
            with mock.patch.object(updater.os,"replace",side_effect=PermissionError("in use")):
                ok,msg=updater.finish_update(exe,staged,wait=1,sleep=lambda s:None,launch=lambda:None)
            self.assertFalse(ok); self.assertIn("교체하지 못했습니다",msg)
            self.assertEqual(exe.read_bytes(),b"old")
            self.assertFalse((Path(d)/"PayableRecon.exe.new").exists())

    def test_cleanup_removes_the_leftover_download(self):
        with tempfile.TemporaryDirectory() as d:
            exe=Path(d)/"PayableRecon.exe"; staged=updater.stage_path(exe)
            staged.write_bytes(b"MZ")
            updater.cleanup_stale(exe,sleep=lambda s:None)
            self.assertFalse(staged.exists())

    def test_cleanup_removes_old_swap_scripts_and_partial_files(self):
        with tempfile.TemporaryDirectory() as d, tempfile.TemporaryDirectory() as tmp:
            exe=Path(d)/"PayableRecon.exe"
            (Path(d)/"PayableRecon.exe.new").write_bytes(b"x")
            old=Path(tmp)/"payablerecon_update_abc.bat"; old.write_text("x")
            other=Path(tmp)/"somebody_else.bat"; other.write_text("x")
            with mock.patch.object(updater.tempfile,"gettempdir",return_value=tmp):
                updater.cleanup_stale(exe,sleep=lambda s:None)
            self.assertFalse(old.exists())
            self.assertFalse((Path(d)/"PayableRecon.exe.new").exists())
            self.assertTrue(other.exists())     # never touches files that are not ours

    def test_no_batch_file_or_console_is_involved(self):
        src=Path(updater.__file__).read_text(encoding="utf-8")
        # the only mention of .bat is the cleanup of the old updater's leftovers
        self.assertNotIn("cmd /c",src)
        self.assertNotIn("mkstemp",src)

if __name__ == "__main__":
    unittest.main()
