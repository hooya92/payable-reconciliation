"""Self-update from the repository's stable GitHub release.

The release notes carry "Build: <commit sha>"; the running exe carries the sha it was
built from (adapters/build_info.py). A different sha means a newer release exists.
The update is downloaded next to the exe, checked, and swapped in by a tiny batch file
once this process has exited (Windows will not overwrite a running exe).
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import urllib.request
from dataclasses import dataclass
from pathlib import Path

REPO = "hooya92/payable-reconciliation"
RELEASE_API = f"https://api.github.com/repos/{REPO}/releases/tags/download-latest"
ASSET_NAME = "PayableRecon.exe"
TIMEOUT = 8
_SHA = re.compile(r"Build:\s*([0-9a-f]{7,40})", re.I)
_DATE = re.compile(r"Date:\s*(\d{4}-\d{2}-\d{2}(?: \d{2}:\d{2})?)")


@dataclass(frozen=True)
class Release:
    sha: str
    url: str
    size: int
    digest: str  # "sha256:<hex>" from GitHub, or ""
    date: str = ""   # human-readable build time, e.g. "2026-11-13 14:05"


def current_sha() -> str:
    from adapters.build_info import BUILD_SHA
    return BUILD_SHA


def version_label() -> str:
    """What a person reads as 'the version': the build time, not a commit id."""
    from adapters.build_info import BUILD_DATE
    return BUILD_DATE or ("개발 버전" if current_sha() == "dev" else "")


def can_self_update() -> bool:
    """Only a packaged Windows exe with a known build can replace itself."""
    return sys.platform == "win32" and bool(getattr(sys, "frozen", False)) and current_sha() != "dev"


def parse_release(payload: dict) -> Release | None:
    match = _SHA.search(payload.get("body") or "")
    if not match:
        return None
    for asset in payload.get("assets") or []:
        if asset.get("name") == ASSET_NAME and asset.get("browser_download_url"):
            return Release(
                sha=match.group(1).lower(),
                url=asset["browser_download_url"],
                size=int(asset.get("size") or 0),
                digest=str(asset.get("digest") or ""),
                date=(_DATE.search(payload.get("body") or "") or [None, ""])[1],
            )
    return None


def is_newer(current: str, latest: str) -> bool:
    current, latest = current.lower(), latest.lower()
    if current in ("", "dev") or not latest:
        return False
    return not (current.startswith(latest) or latest.startswith(current))


def check_latest() -> Release | None:
    """The newer release, or None when up to date / unreachable. Never raises."""
    try:
        req = urllib.request.Request(RELEASE_API, headers={"Accept": "application/vnd.github+json", "User-Agent": "PayableRecon"})
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            release = parse_release(json.load(resp))
    except Exception:
        return None
    if release and is_newer(current_sha(), release.sha):
        return release
    return None


def download(release: Release, dest: Path, progress=None) -> None:
    """Download and verify. Raises ValueError/OSError with a Korean message on failure."""
    req = urllib.request.Request(release.url, headers={"User-Agent": "PayableRecon"})
    sha = hashlib.sha256()
    done = 0
    with urllib.request.urlopen(req, timeout=30) as resp, open(dest, "wb") as out:
        total = int(resp.headers.get("Content-Length") or release.size or 0)
        while True:
            chunk = resp.read(1 << 16)
            if not chunk:
                break
            out.write(chunk)
            sha.update(chunk)
            done += len(chunk)
            if progress:
                progress(done, total)
    if release.size and done != release.size:
        raise ValueError("내려받은 파일 크기가 맞지 않습니다. 다시 시도해 주세요.")
    if release.digest.startswith("sha256:") and sha.hexdigest() != release.digest[7:].lower():
        raise ValueError("내려받은 파일의 검증(해시)이 맞지 않아 중단했습니다.")
    with open(dest, "rb") as fh:
        if fh.read(2) != b"MZ":
            raise ValueError("내려받은 파일이 실행 파일 형식이 아닙니다.")


def stage_path(exe: Path) -> Path:
    return exe.with_name(exe.stem + ".update.exe")


def _log(message: str) -> None:
    """Best-effort trace in %TEMP% so a failed swap can be diagnosed afterwards."""
    try:
        with open(Path(tempfile.gettempdir()) / "payablerecon_update.log", "a", encoding="utf-8") as fh:
            fh.write(message + "\n")
    except OSError:
        pass


def launch_finisher(exe: Path, staged: Path) -> None:
    """Start the downloaded exe in 'finish the update' mode, then let this process exit.

    No batch file and no console window: the new exe waits until the old one is released,
    copies itself over it and starts it.
    """
    env = dict(os.environ, PYINSTALLER_RESET_ENVIRONMENT="1")
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen([str(staged), FINISH_FLAG, str(exe), str(staged)],
                     creationflags=flags, close_fds=True, env=env)


FINISH_FLAG = "--finish-update"


def finish_update(exe: Path, staged: Path, wait: float = 60.0, sleep=None, launch=None):
    """Runs inside the NEW exe. Replace `exe` with this file and start it. Returns (ok, message)."""
    import shutil
    import time
    sleep = sleep or time.sleep
    tmp = exe.with_name(exe.name + ".new")
    try:
        shutil.copyfile(staged, tmp)
        if tmp.stat().st_size != staged.stat().st_size:
            raise OSError("복사본 크기가 맞지 않습니다")
    except OSError as e:
        _log(f"copy failed: {e}")
        return False, f"새 파일을 준비하지 못했습니다: {e}"
    waited = 0.0
    last = None
    while True:
        try:
            os.replace(tmp, exe)   # fails while the old program still holds the file
            break
        except OSError as e:
            last = e
            if waited >= wait:
                try:
                    tmp.unlink()
                except OSError:
                    pass
                _log(f"replace gave up after {waited:.0f}s: {last}")
                return False, f"기존 프로그램이 닫히지 않아 교체하지 못했습니다 ({last})"
            sleep(0.5)
            waited += 0.5
    _log(f"replaced after {waited:.1f}s")
    try:
        (launch or (lambda: subprocess.Popen(
            [str(exe)], close_fds=True,
            creationflags=getattr(subprocess, "DETACHED_PROCESS", 0),
            env=dict(os.environ, PYINSTALLER_RESET_ENVIRONMENT="1"))))()
    except OSError as e:
        return False, f"교체는 끝났지만 새 버전을 자동으로 열지 못했습니다: {e}"
    return True, ""


def stop_old_swap_scripts() -> None:
    """Stop and delete swap scripts left over from the earlier batch-file updater.

    Only processes whose command line names our own script (payablerecon_update_*.bat) are touched.
    """
    for bat in Path(tempfile.gettempdir()).glob("payablerecon_update_*.bat"):
        try:
            bat.unlink()    # a looping script stops when its file is gone
        except OSError:
            pass
    if sys.platform != "win32":
        return
    command = ("Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*payablerecon_update_*' } "
               "| ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }")
    try:
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), timeout=20,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except (OSError, subprocess.SubprocessError):
        pass


def cleanup_stale(exe: Path, tries: int = 10, sleep=None) -> None:
    """Delete the leftover download of a finished update (the finisher may still be exiting)."""
    import time
    sleep = sleep or time.sleep
    stop_old_swap_scripts()
    for leftover in (exe.with_name(exe.name + ".new"),):
        try:
            leftover.unlink()
        except OSError:
            pass
    staged = stage_path(exe)
    for _ in range(tries):
        if not staged.exists():
            return
        try:
            staged.unlink()
            return
        except OSError:
            sleep(1.0)
