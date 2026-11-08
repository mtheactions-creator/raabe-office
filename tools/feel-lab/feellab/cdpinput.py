"""Trusted input into RaabeCalc's WebView2 through DevTools (Input.dispatchKeyEvent / dispatchMouseEvent): the steps of
feellab/calc_catalogue.py done "by hand" in the real exe without the OS mouse and keyboard (no real-input lock; the
window can stay off screen). Keys as on a German keyboard: "+" and "-" sit where the US "]" and "/" are."""
from __future__ import annotations

import json
import time

MODS = {"alt": 1, "ctrl": 2, "meta": 4, "shift": 8}
NAMED = {
    "enter": ("Enter", "Enter", 13, "\r"), "tab": ("Tab", "Tab", 9, None), "esc": ("Escape", "Escape", 27, None),
    "backspace": ("Backspace", "Backspace", 8, None), "delete": ("Delete", "Delete", 46, None), "space": (" ", "Space", 32, " "),
    "left": ("ArrowLeft", "ArrowLeft", 37, None), "up": ("ArrowUp", "ArrowUp", 38, None), "right": ("ArrowRight", "ArrowRight", 39, None),
    "down": ("ArrowDown", "ArrowDown", 40, None), "home": ("Home", "Home", 36, None), "end": ("End", "End", 35, None),
    "pageup": ("PageUp", "PageUp", 33, None), "pagedown": ("PageDown", "PageDown", 34, None),
    "plus": ("+", "BracketRight", 0xBB, "+"), "minus": ("-", "Slash", 0xBD, "-"), "alt": ("Alt", "AltLeft", 18, None),
    **{f"f{i}": (f"F{i}", f"F{i}", 0x6F + i, None) for i in range(1, 13)},
}
# German keyboard: what Umschalt gives on the digit keys.
SHIFT_DIGITS = {"1": "!", "2": '"', "3": "§", "4": "$", "5": "%", "6": "&", "7": "/", "8": "(", "9": ")", "0": "="}


class Cdp:
    """Input for a CalcSide (its DevTools websocket)."""

    def __init__(self, side):
        self.side = side

    def call(self, method: str, params: dict):
        s = self.side
        s._id += 1
        s.ws.send(json.dumps({"id": s._id, "method": method, "params": params}))
        while True:
            msg = json.loads(s.ws.recv())
            if msg.get("id") == s._id:
                return msg.get("result")

    # ----- keys ----------------------------------------------------------------------------------

    def key(self, combo: str, settle: float = 0.06) -> None:
        parts = combo.lower().split("+") if combo != "+" else ["plus"]
        # "ctrl+plus" / "ctrl++"
        if combo.endswith("++"):
            parts = combo.lower()[:-2].split("+") + ["plus"]
        mods = 0
        for m in parts[:-1]:
            mods |= MODS[m]
        k = parts[-1]
        if k in NAMED:
            key, code, vk, text = NAMED[k]
        elif len(k) == 1 and k.isalpha():
            key, code, vk, text = (k.upper() if mods & 8 else k), f"Key{k.upper()}", ord(k.upper()), (k.upper() if mods & 8 else k)
        elif len(k) == 1 and k.isdigit():
            key, code, vk, text = (SHIFT_DIGITS[k] if mods & 8 else k), f"Digit{k}", ord(k), (SHIFT_DIGITS[k] if mods & 8 else k)
        else:
            raise ValueError(f"unknown key {combo!r}")
        down = {"type": "keyDown" if text and not (mods & 3) else "rawKeyDown", "key": key, "code": code, "windowsVirtualKeyCode": vk, "modifiers": mods}
        if text and not (mods & 3):
            down["text"] = text
        if mods == 2 and k == "a":
            down["commands"] = ["selectAll"]
        self.call("Input.dispatchKeyEvent", down)
        self.call("Input.dispatchKeyEvent", {"type": "keyUp", "key": key, "code": code, "windowsVirtualKeyCode": vk, "modifiers": mods})
        time.sleep(settle)

    def type(self, text: str) -> None:
        for ch in text:
            if ch.isalpha() and ch.isascii():
                code, vk = f"Key{ch.upper()}", ord(ch.upper())
            elif ch.isdigit():
                code, vk = f"Digit{ch}", ord(ch)
            else:
                code, vk = "", 0
            shift = 8 if ch.isupper() or ch in '!"§$%&/()=?*:;>_' else 0
            ev = {"key": ch, "code": code, "windowsVirtualKeyCode": vk, "modifiers": shift}
            self.call("Input.dispatchKeyEvent", {"type": "keyDown", "text": ch, **ev})
            self.call("Input.dispatchKeyEvent", {"type": "keyUp", **ev})
            time.sleep(0.015)

    def keytips(self, letters: str) -> None:
        """Alt pressed and released alone, then each keytip's letters ("R H C")."""
        self.key("alt", 0.15)
        for tip in letters.split():
            for ch in tip:
                self.key(ch.lower(), 0.15)
        time.sleep(0.2)

    # ----- mouse ---------------------------------------------------------------------------------

    def click(self, x: float, y: float, count: int = 1, mods: int = 0) -> None:
        self.call("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x, "y": y, "modifiers": mods})
        for n in range(1, count + 1):
            self.call("Input.dispatchMouseEvent", {"type": "mousePressed", "x": x, "y": y, "button": "left", "buttons": 1, "clickCount": n, "modifiers": mods})
            self.call("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": x, "y": y, "button": "left", "buttons": 0, "clickCount": n, "modifiers": mods})
        time.sleep(0.1)

    def drag(self, x1: float, y1: float, x2: float, y2: float, steps: int = 12, mods: int = 0) -> None:
        self.call("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x1, "y": y1, "modifiers": mods})
        self.call("Input.dispatchMouseEvent", {"type": "mousePressed", "x": x1, "y": y1, "button": "left", "buttons": 1, "clickCount": 1, "modifiers": mods})
        for i in range(1, steps + 1):
            t = i / steps
            self.call("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x1 + (x2 - x1) * t, "y": y1 + (y2 - y1) * t, "button": "left", "buttons": 1, "modifiers": mods})
            time.sleep(0.01)
        self.call("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": x2, "y": y2, "button": "left", "buttons": 0, "clickCount": 1, "modifiers": mods})
        time.sleep(0.15)

    def point(self, addr: str, where: str = "center") -> tuple[float, float]:
        p = self.side.js(f"__calc.point({json.dumps(addr)}, {json.dumps(where)})")
        return p["x"], p["y"]

    def rect(self, js_element: str) -> tuple[float, float, float, float] | None:
        return self.side.js(f"(() => {{ const e = {js_element}; if (!e) return null; const r = e.getBoundingClientRect(); return [r.left, r.top, r.width, r.height]; }})()")

    def click_el(self, js_element: str, fx: float = 0.5, fy: float = 0.5, count: int = 1) -> bool:
        r = self.rect(js_element)
        if not r:
            return False
        self.click(r[0] + r[2] * fx, r[1] + r[3] * fy, count)
        return True

    # ----- the catalogue's steps -------------------------------------------------------------------

    def step(self, st: tuple) -> None:
        kind = st[0]
        if kind == "type":
            self.type(st[1])
        elif kind == "keys":
            for _ in range(st[2] if len(st) > 2 else 1):
                self.key(st[1])
        elif kind == "keytips":
            self.keytips(st[1])
        elif kind in ("click", "dblclick"):
            x, y = self.point(st[1])
            mods = sum(MODS[m] for m in (st[2].split("+") if len(st) > 2 and st[2] else []))
            self.click(x, y, 2 if kind == "dblclick" else 1, mods)
        elif kind == "drag":
            x1, y1 = self.point(st[1])
            x2, y2 = self.point(st[2])
            self.drag(x1, y1, x2, y2)
        elif kind == "fill":
            x1, y1 = self.point(st[1].split(":")[-1], "br")
            x2, y2 = self.point(st[2])
            self.drag(x1 - 1, y1 - 1, x2, y2)
        elif kind == "colborder":
            p = self.side.js(f"__calc.header('colborder', {json.dumps(st[1])})")
            self.click(p["x"], p["y"], 2)
        elif kind in ("tabdbl", "tabclick"):
            self.click_el(f"[...document.querySelectorAll('.sheettab')].find(t => t.textContent === {json.dumps(st[1])})", count=2 if kind == "tabdbl" else 1)
        elif kind == "ribbon":
            self.side.js("document.querySelector('.ribbon [role=tab][data-keytip=R]')?.click()")
            time.sleep(0.2)
            sel = {"font.fill": "[...document.querySelectorAll('.ribbon button.color-main')].find(b => (b.getAttribute('aria-label') || '').startsWith('Füllfarbe'))"}[st[1]]
            self.click_el(sel, 0.3, 0.5)
        elif kind == "wait":
            time.sleep(st[1])
        else:
            raise ValueError(kind)
        time.sleep(0.08)
