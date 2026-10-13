"""Thin ctypes layer over the Win32 calls the harness needs: real input via SendInput,
window lookup and placement, and fast screen-region capture. All coordinates are
physical screen pixels (the process is per-monitor DPI aware)."""

from __future__ import annotations

import ctypes
import time
import zlib
from ctypes import wintypes

user32 = ctypes.WinDLL("user32", use_last_error=True)
gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

# Per-monitor DPI awareness v2, so every coordinate below is in physical pixels.
user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))

# ----- SendInput ------------------------------------------------------------------------

INPUT_MOUSE, INPUT_KEYBOARD = 0, 1
KEYEVENTF_EXTENDEDKEY, KEYEVENTF_KEYUP, KEYEVENTF_UNICODE, KEYEVENTF_SCANCODE = 0x1, 0x2, 0x4, 0x8
MOUSEEVENTF_MOVE, MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP = 0x1, 0x2, 0x4
MOUSEEVENTF_ABSOLUTE, MOUSEEVENTF_VIRTUALDESK = 0x8000, 0x4000

ULONG_PTR = ctypes.c_size_t


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG), ("mouseData", wintypes.DWORD),
                ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", ULONG_PTR)]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD), ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("dwExtraInfo", ULONG_PTR)]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [("uMsg", wintypes.DWORD), ("wParamL", wintypes.WORD), ("wParamH", wintypes.WORD)]


class _INPUTUNION(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT), ("hi", HARDWAREINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("u", _INPUTUNION)]


user32.SendInput.argtypes = (wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int)
user32.MapVirtualKeyW.argtypes = (wintypes.UINT, wintypes.UINT)

VK = {
    "backspace": 0x08, "tab": 0x09, "enter": 0x0D, "shift": 0x10, "ctrl": 0x11, "alt": 0x12,
    "esc": 0x1B, "space": 0x20, "pageup": 0x21, "pagedown": 0x22, "end": 0x23, "home": 0x24,
    "left": 0x25, "up": 0x26, "right": 0x27, "down": 0x28, "delete": 0x2E,
}
EXTENDED = {0x21, 0x22, 0x23, 0x24, 0x25, 0x26, 0x27, 0x28, 0x2E}


def _send(*inputs: INPUT) -> None:
    arr = (INPUT * len(inputs))(*inputs)
    if user32.SendInput(len(inputs), arr, ctypes.sizeof(INPUT)) != len(inputs):
        raise ctypes.WinError(ctypes.get_last_error())


def _key(vk: int, up: bool) -> INPUT:
    flags = (KEYEVENTF_KEYUP if up else 0) | (KEYEVENTF_EXTENDEDKEY if vk in EXTENDED else 0)
    scan = user32.MapVirtualKeyW(vk, 0)
    return INPUT(type=INPUT_KEYBOARD, u=_INPUTUNION(ki=KEYBDINPUT(wVk=vk, wScan=scan, dwFlags=flags)))


def vk_for(name: str) -> int:
    name = name.lower()
    if name in VK:
        return VK[name]
    if len(name) == 1 and name.isalnum():
        return ord(name.upper())
    raise ValueError(f"unknown key {name!r}")


def key_down(vk: int) -> None:
    _send(_key(vk, False))


def key_up(vk: int) -> None:
    _send(_key(vk, True))


def press(combo: str, hold: float = 0.02) -> None:
    """Presses a combination like "ctrl+shift+right" with real virtual-key events."""
    vks = [vk_for(k) for k in combo.split("+")]
    for vk in vks:
        _send(_key(vk, False))
        time.sleep(0.005)
    time.sleep(hold)
    for vk in reversed(vks):
        _send(_key(vk, True))
        time.sleep(0.005)


def type_text(text: str, gap: float = 0.03) -> None:
    """Types characters as Unicode key events (layout-independent)."""
    for ch in text:
        code = ord(ch)
        down = INPUT(type=INPUT_KEYBOARD, u=_INPUTUNION(ki=KEYBDINPUT(wScan=code, dwFlags=KEYEVENTF_UNICODE)))
        up = INPUT(type=INPUT_KEYBOARD, u=_INPUTUNION(ki=KEYBDINPUT(wScan=code, dwFlags=KEYEVENTF_UNICODE | KEYEVENTF_KEYUP)))
        _send(down, up)
        time.sleep(gap)


def screen_size() -> tuple[int, int]:
    return user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)


def _mouse(x: float, y: float, flags: int) -> INPUT:
    w, h = screen_size()
    ax = int(round(x * 65535 / (w - 1)))
    ay = int(round(y * 65535 / (h - 1)))
    return INPUT(type=INPUT_MOUSE, u=_INPUTUNION(mi=MOUSEINPUT(dx=ax, dy=ay, dwFlags=flags | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE)))


def move(x: float, y: float) -> None:
    _send(_mouse(x, y, 0))


def mouse_down(x: float, y: float) -> None:
    _send(_mouse(x, y, MOUSEEVENTF_LEFTDOWN))


def mouse_up(x: float, y: float) -> None:
    _send(_mouse(x, y, MOUSEEVENTF_LEFTUP))


def click(x: float, y: float, count: int = 1, modifiers: str = "") -> None:
    mods = [vk_for(m) for m in modifiers.split("+") if m]
    move(x, y)
    time.sleep(0.05)
    for vk in mods:
        key_down(vk)
    for i in range(count):
        mouse_down(x, y)
        time.sleep(0.02)
        mouse_up(x, y)
        if i < count - 1:
            time.sleep(0.06)
    for vk in reversed(mods):
        key_up(vk)


def drag(x1: float, y1: float, x2: float, y2: float, steps: int = 20, step_delay: float = 0.012,
         start_clicks: int = 1) -> None:
    """Presses at (x1, y1) (optionally after extra clicks for word/paragraph drags) and moves to (x2, y2)."""
    move(x1, y1)
    time.sleep(0.05)
    for _ in range(start_clicks - 1):
        mouse_down(x1, y1); time.sleep(0.02); mouse_up(x1, y1); time.sleep(0.06)
    mouse_down(x1, y1)
    time.sleep(0.08)
    for i in range(1, steps + 1):
        move(x1 + (x2 - x1) * i / steps, y1 + (y2 - y1) * i / steps)
        time.sleep(step_delay)
    time.sleep(0.08)
    mouse_up(x2, y2)


# ----- windows --------------------------------------------------------------------------

WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


def window_title(hwnd: int) -> str:
    n = user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, buf, n + 1)
    return buf.value


def window_class(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def find_windows(predicate) -> list[int]:
    found: list[int] = []

    def cb(hwnd, _):
        if user32.IsWindowVisible(hwnd) and predicate(hwnd):
            found.append(hwnd)
        return True

    user32.EnumWindows(WNDENUMPROC(cb), 0)
    return found


def child_windows(hwnd: int) -> list[int]:
    found: list[int] = []

    def cb(h, _):
        found.append(h)
        return True

    user32.EnumChildWindows(hwnd, WNDENUMPROC(cb), 0)
    return found


def wait_for_window(predicate, timeout: float = 20.0) -> int:
    end = time.perf_counter() + timeout
    while time.perf_counter() < end:
        hits = find_windows(predicate)
        if hits:
            return hits[0]
        time.sleep(0.005)
    raise TimeoutError("window did not appear")


def place(hwnd: int, x: int, y: int, w: int, h: int) -> None:
    user32.ShowWindow(hwnd, 9)  # SW_RESTORE
    user32.SetWindowPos(hwnd, 0, x, y, w, h, 0x0004 | 0x0010)  # NOZORDER | NOACTIVATE


def activate(hwnd: int) -> None:
    """Brings a window to the foreground despite Windows' foreground-lock rules."""
    if user32.GetForegroundWindow() == hwnd:
        return
    fg = user32.GetForegroundWindow()
    fg_thread = user32.GetWindowThreadProcessId(fg, None)
    me = kernel32.GetCurrentThreadId()
    user32.AttachThreadInput(me, fg_thread, True)
    try:
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)
    finally:
        user32.AttachThreadInput(me, fg_thread, False)
    time.sleep(0.15)
    if user32.GetForegroundWindow() != hwnd:
        raise RuntimeError("could not activate the target window")


class POINT(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


class RECT(ctypes.Structure):
    _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG), ("right", wintypes.LONG), ("bottom", wintypes.LONG)]


def client_origin(hwnd: int) -> tuple[int, int]:
    pt = POINT(0, 0)
    user32.ClientToScreen(hwnd, ctypes.byref(pt))
    return pt.x, pt.y


def window_rect(hwnd: int) -> tuple[int, int, int, int]:
    r = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(r))
    return r.left, r.top, r.right, r.bottom


# ----- screen capture -------------------------------------------------------------------

class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG), ("biYPelsPerMeter", wintypes.LONG),
                ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD)]


gdi32.CreateCompatibleBitmap.restype = wintypes.HBITMAP
gdi32.CreateCompatibleDC.restype = wintypes.HDC
user32.GetDC.restype = wintypes.HDC
gdi32.SelectObject.argtypes = (wintypes.HDC, wintypes.HGDIOBJ)
gdi32.BitBlt.argtypes = (wintypes.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                         wintypes.HDC, ctypes.c_int, ctypes.c_int, wintypes.DWORD)
gdi32.GetDIBits.argtypes = (wintypes.HDC, wintypes.HBITMAP, wintypes.UINT, wintypes.UINT, ctypes.c_void_p,
                            ctypes.POINTER(BITMAPINFOHEADER), wintypes.UINT)


class Grabber:
    """Captures one fixed screen region repeatedly with minimal overhead."""

    def __init__(self, x: int, y: int, w: int, h: int):
        self.x, self.y, self.w, self.h = x, y, w, h
        self.screen_dc = user32.GetDC(0)
        self.mem_dc = gdi32.CreateCompatibleDC(self.screen_dc)
        self.bmp = gdi32.CreateCompatibleBitmap(self.screen_dc, w, h)
        gdi32.SelectObject(self.mem_dc, self.bmp)
        self.info = BITMAPINFOHEADER(biSize=ctypes.sizeof(BITMAPINFOHEADER), biWidth=w, biHeight=-h,
                                     biPlanes=1, biBitCount=32, biCompression=0)
        self.buf = ctypes.create_string_buffer(w * h * 4)

    def grab(self) -> bytes:
        gdi32.BitBlt(self.mem_dc, 0, 0, self.w, self.h, self.screen_dc, self.x, self.y, 0x00CC0020)
        gdi32.GetDIBits(self.mem_dc, self.bmp, 0, self.h, self.buf, ctypes.byref(self.info), 0)
        return self.buf.raw

    def signature(self) -> int:
        return zlib.crc32(self.grab())

    def close(self) -> None:
        gdi32.DeleteObject(self.bmp)
        gdi32.DeleteDC(self.mem_dc)
        user32.ReleaseDC(0, self.screen_dc)


def save_region_png(path: str, x: int, y: int, w: int, h: int) -> None:
    from PIL import Image
    g = Grabber(x, y, w, h)
    try:
        raw = g.grab()
    finally:
        g.close()
    Image.frombuffer("RGBA", (w, h), raw, "raw", "BGRA", 0, 1).convert("RGB").save(path)
