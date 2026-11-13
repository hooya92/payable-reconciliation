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


def swap_script(exe: Path, staged: Path) -> str:
    """Batch text: wait until the old exe is free, replace it, relaunch, delete itself."""
    return (
        "@echo off\r\n"
        "chcp 65001 >nul\r\n"
        "set PYINSTALLER_RESET_ENVIRONMENT=1\r\n"
        "set tries=0\r\n"
        ":retry\r\n"
        "set /a tries+=1\r\n"
        f'move /y "{staged}" "{exe}" >nul 2>&1\r\n'
        "if errorlevel 1 (\r\n"
        "  if %tries% geq 60 goto giveup\r\n"
        "  ping -n 2 127.0.0.1 >nul\r\n"
        "  goto retry\r\n"
        ")\r\n"
        f'start "" "{exe}"\r\n'
        '(goto) 2>nul & del "%~f0"\r\n'
        ":giveup\r\n"
        f'del "{staged}" >nul 2>&1\r\n'
        '(goto) 2>nul & del "%~f0"\r\n'
    )


def launch_swap(exe: Path, staged: Path) -> None:
    fd, bat = tempfile.mkstemp(suffix=".bat", prefix="payablerecon_update_")
    with os.fdopen(fd, "w", encoding="utf-8", newline="") as fh:
        fh.write(swap_script(exe, staged))
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | getattr(subprocess, "DETACHED_PROCESS", 0)
    subprocess.Popen(["cmd", "/c", bat], creationflags=flags, close_fds=True)
