"""Move and resize the borderless window with the Windows API.

The window has no title bar, so the page asks for a "drag" on mouse-down and we follow the real
cursor until the button is released. Cursor and window are both read in raw screen pixels, so
the result is right at any display scaling. The geometry itself is plain arithmetic (tested).
"""
from __future__ import annotations

import sys
import threading
import time

MODES = ("move", "se", "s", "e")


def geometry(mode, rect, start, cursor, min_size):
    """New (left, top, width, height) after the cursor moved from `start` to `cursor`."""
    left, top, width, height = rect
    dx, dy = cursor[0] - start[0], cursor[1] - start[1]
    min_w, min_h = min_size
    if mode == "move":
        return left + dx, top + dy, width, height
    if mode in ("se", "e"):
        width = max(min_w, width + dx)
    if mode in ("se", "s"):
        height = max(min_h, height + dy)
    return left, top, width, height


class WindowMover:
    def __init__(self, window, logical_min=(460, 640)):
        self.window = window
        self.logical_min = logical_min
        self.busy = False

    def _hwnd(self):
        return int(self.window.native.Handle.ToInt64())

    def begin(self, mode):
        """Start following the cursor. Returns False when it cannot (not Windows, button up, ...)."""
        if sys.platform != "win32" or mode not in MODES or self.busy:
            return False
        try:
            import ctypes
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            hwnd = self._hwnd()
            if not (user32.GetAsyncKeyState(0x01) & 0x8000):
                return False
            rect, cursor = wintypes.RECT(), wintypes.POINT()
            user32.GetWindowRect(hwnd, ctypes.byref(rect))
            user32.GetCursorPos(ctypes.byref(cursor))
            try:
                scale = user32.GetDpiForWindow(hwnd) / 96
            except Exception:
                scale = 1.0
            start_rect = (rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)
            start = (cursor.x, cursor.y)
            min_size = (int(self.logical_min[0] * scale), int(self.logical_min[1] * scale))
        except Exception:
            return False
        self.busy = True
        threading.Thread(target=self._follow, args=(mode, hwnd, start_rect, start, min_size), daemon=True).start()
        return True

    def _follow(self, mode, hwnd, rect, start, min_size):
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        SWP_NOZORDER, SWP_NOACTIVATE = 0x0004, 0x0010
        point = wintypes.POINT()
        last = None
        try:
            while user32.GetAsyncKeyState(0x01) & 0x8000:
                user32.GetCursorPos(ctypes.byref(point))
                target = geometry(mode, rect, start, (point.x, point.y), min_size)
                if target != last:
                    user32.SetWindowPos(hwnd, 0, *target, SWP_NOZORDER | SWP_NOACTIVATE)
                    last = target
                time.sleep(0.008)
        finally:
            self.busy = False
