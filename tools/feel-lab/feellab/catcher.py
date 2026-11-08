"""Plain windows of our own beside and below the window under test (D-127): the out-of-window drags let the button go
there. They accept nothing dropped on them, so an Office drag that turns into a system drag and drop when it leaves its
window (cells, shapes) ends without effect instead of reaching another program's window or the desktop, and they never
take the focus.

    with catcher((0, 0, 2000, 1700)):
        ...
"""
from __future__ import annotations

import contextlib
import ctypes
import threading
from ctypes import wintypes

from . import winapi

user32 = ctypes.WinDLL("user32", use_last_error=True)
user32.CreateWindowExW.restype = wintypes.HWND
user32.CreateWindowExW.argtypes = [wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD, ctypes.c_int, ctypes.c_int,
                                   ctypes.c_int, ctypes.c_int, wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, wintypes.LPVOID]
WS_POPUP, WS_VISIBLE = 0x80000000, 0x10000000
WS_EX_TOPMOST, WS_EX_TOOLWINDOW, WS_EX_NOACTIVATE = 0x8, 0x80, 0x08000000
WM_CLOSE, WM_QUIT = 0x0010, 0x0012


@contextlib.contextmanager
def catcher(rect: tuple[int, int, int, int], margin: int = 600):
    """Catch windows right of and below `rect` (left, top, right, bottom in screen pixels)."""
    l, t, r, b = rect
    sw, sh = winapi.screen_size()
    areas = [(r, t, min(margin, sw - r), b - t + min(margin, sh - b)), (l, b, r - l, min(margin, sh - b))]
    ready = threading.Event()
    state: dict = {}

    def run() -> None:
        state["thread"] = ctypes.windll.kernel32.GetCurrentThreadId()
        hwnds = []
        for x, y, w, h in areas:
            if w > 0 and h > 0:
                hwnds.append(user32.CreateWindowExW(WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE, "STATIC", "raabe drag catcher",
                                                    WS_POPUP | WS_VISIBLE, x, y, w, h, None, None, None, None))
        state["hwnds"] = hwnds
        ready.set()
        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
        for hw in hwnds:
            user32.DestroyWindow(hw)

    th = threading.Thread(target=run, daemon=True)
    th.start()
    ready.wait(5)
    try:
        yield state.get("hwnds", [])
    finally:
        user32.PostThreadMessageW(state.get("thread", 0), WM_QUIT, 0, 0)
        th.join(3)
