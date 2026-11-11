"""Step 11: how Office's menus close. Helpers shared by the probes probe_*_menuclose.py.

Office draws context menus, ribbon dropdowns, galleries and palettes in top-level "Net UI Tool Window" windows. A probe
reads only the windows of its own Office process (CLAUDE.md: never other windows on the screen)."""
from __future__ import annotations

import subprocess
import sys
import time

from . import uia, winapi

POPUP_CLASSES = ("Net UI Tool Window", "Net UI Tool Window Layered")


def own_popups(pid: int) -> list[dict]:
    """Visible popups of process `pid`, larger than a tooltip-sized sliver."""
    out = []
    for h in winapi.find_windows(lambda h: winapi.window_pid(h) == pid and winapi.window_class(h) in POPUP_CLASSES):
        l, t, r, b = winapi.window_rect(h)
        if r - l > 40 and b - t > 40:
            out.append({"w": r - l, "h": b - t, "rect": [l, t, r, b]})
    return out


def outside(pt: tuple[float, float], pops: list[dict], margin: int = 12) -> bool:
    x, y = pt
    return all(not (p["rect"][0] - margin <= x <= p["rect"][2] + margin and p["rect"][1] - margin <= y <= p["rect"][3] + margin)
               for p in pops)


def pick(candidates: dict, pops: list[dict]):
    """The first named point that lies outside every open popup: a click meant for "outside" must never land on a
    menu item."""
    for name, pt in candidates.items():
        if outside(pt, pops):
            return name, pt
    return None, None


def wait_popup(pid: int, timeout: float = 3.0) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        if own_popups(pid):
            time.sleep(0.3)
            return True
        time.sleep(0.05)
    return False


def rect_center(el) -> tuple[float, float]:
    r = uia.describe(el)["rect"]
    return r[0] + r[2] / 2, r[1] + r[3] / 2


def find_named(root, typ: str, name: str, depth: int = 14):
    return uia.find(root, lambda d: d["type"] == typ and d["name"] == name, max_depth=depth)


def split_arrow(root, name: str):
    """The arrow part ("Weitere Optionen") of a ribbon split button such as Füllfarbe."""
    sb = uia.find(root, lambda d: d["type"] == "SplitButton" and d["name"] == name, max_depth=14)
    if not sb:
        return None
    for c in uia.children(sb):
        if uia.describe(c)["type"] == "MenuItem":
            return c
    return None


def selected_ribbon_tab(root) -> str:
    for_tab = []

    def pred(d):
        return d["type"] == "TabItem" and d["class"] == "NetUIRibbonTab"
    stack = [(root, 0)]
    while stack:
        e, d = stack.pop()
        try:
            dc = uia.describe(e)
            if pred(dc):
                p = e.GetCurrentPattern(uia.U.UIA_SelectionItemPatternId)
                if p and p.QueryInterface(uia.U.IUIAutomationSelectionItemPattern).CurrentIsSelected:
                    return dc["name"]
                for_tab.append(dc["name"])
                continue
            if d < 10:
                stack.extend((c, d + 1) for c in uia.children(e))
        except Exception:  # noqa: BLE001
            continue
    return "?"


class ForeignWindow:
    """A small window of our own on the right of the screen (outside the Office window), to click into: the Office
    window then loses the focus as when a user clicks another program."""

    def __init__(self, x: int = 2300, y: int = 300):
        code = ("import tkinter as t; r=t.Tk(); r.title('raabe-focus-probe'); "
                f"r.geometry('400x300+{x}+{y}'); r.after(3600000, r.destroy); r.mainloop()")
        self.proc = subprocess.Popen([sys.executable, "-c", code])
        # A venv's python.exe starts the real interpreter as a child: find the window by its title.
        self.hwnd = winapi.wait_for_window(lambda h: winapi.window_title(h) == "raabe-focus-probe")
        l, t, r, b = winapi.window_rect(self.hwnd)
        self.point = ((l + r) / 2, (t + b) / 2 + 20)

    def close(self) -> None:
        winapi.user32.PostMessageW(self.hwnd, 0x0010, 0, 0)  # WM_CLOSE: the venv launcher's child closes too
        self.proc.terminate()
        self.proc.wait(5)
