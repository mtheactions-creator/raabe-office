"""A CDP client for RaabeWrite's page (the real exe's WebView2 or the dev build in a headless Edge) that keeps the
page's errors: uncaught exceptions, console.error and failed promises arrive as protocol events and are collected
while commands run (the by-hand audit, D-133). Input is trusted DevTools input (Input.dispatchMouseEvent/KeyEvent):
it goes into the page only, never to the real mouse or keyboard."""
from __future__ import annotations

import json
import time

import websocket

# Virtual-key codes for the keys the sweep presses.
VK = {"Enter": 13, "Backspace": 8, "Tab": 9, "Escape": 27, "Delete": 46, "Home": 36, "End": 35, "ArrowLeft": 37,
      "ArrowUp": 38, "ArrowRight": 39, "ArrowDown": 40, "F1": 112, "F7": 118, "F9": 120, "F10": 121, "F12": 123,
      "Alt": 18, "PageDown": 34, "PageUp": 33, " ": 32}


class WritePage:
    def __init__(self, ws: websocket.WebSocket):
        self.ws = ws
        self.n = 0
        self.events: list[dict] = []
        self.cdp("Runtime.enable")
        self.cdp("Log.enable")

    # ----- protocol -------------------------------------------------------------------------
    def _keep(self, m: dict) -> None:
        meth = m.get("method")
        if meth == "Runtime.exceptionThrown":
            d = m["params"]["exceptionDetails"]
            text = (d.get("exception") or {}).get("description") or d.get("text", "")
            self.events.append({"kind": "exception", "text": text[:600]})
        elif meth == "Runtime.consoleAPICalled" and m["params"].get("type") in ("error", "assert"):
            args = m["params"].get("args", [])
            text = " ".join(str(a.get("value", a.get("description", ""))) for a in args)
            self.events.append({"kind": "console.error", "text": text[:600]})
        elif meth == "Log.entryAdded" and m["params"]["entry"].get("level") == "error":
            e = m["params"]["entry"]
            self.events.append({"kind": "log.error", "text": (e.get("text", "") + " " + e.get("url", ""))[:600]})

    def cdp(self, method: str, params: dict | None = None):
        self.n += 1
        my = self.n
        self.ws.send(json.dumps({"id": my, "method": method, "params": params or {}}))
        while True:
            m = json.loads(self.ws.recv())
            if m.get("id") == my:
                if "error" in m:
                    raise RuntimeError(m["error"])
                return m.get("result")
            if "method" in m:
                self._keep(m)

    def js(self, expr: str, timeout_note: str = ""):
        r = self.cdp("Runtime.evaluate", {"expression": expr, "returnByValue": True, "awaitPromise": True})
        if "exceptionDetails" in r:
            d = r["exceptionDetails"]
            raise RuntimeError(((d.get("exception") or {}).get("description") or d.get("text", ""))[:1500] + timeout_note)
        return r["result"].get("value")

    def pump(self, seconds: float) -> None:
        """Waits, reading events that arrive meanwhile (timers, promises)."""
        end = time.perf_counter() + seconds
        while time.perf_counter() < end:
            self.js("0")
            time.sleep(0.03)

    def take_errors(self) -> list[dict]:
        out, self.events = self.events, []
        return out

    # ----- input ----------------------------------------------------------------------------
    def move(self, x: float, y: float) -> None:
        self.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x, "y": y, "button": "none"})

    def click(self, x: float, y: float, button: str = "left", count: int = 1, modifiers: int = 0) -> None:
        self.move(x, y)
        for c in range(1, count + 1):
            base = {"x": x, "y": y, "button": button, "clickCount": c, "modifiers": modifiers,
                    "buttons": 1 if button == "left" else 2}
            self.cdp("Input.dispatchMouseEvent", {"type": "mousePressed", **base})
            self.cdp("Input.dispatchMouseEvent", {"type": "mouseReleased", **base, "buttons": 0})

    def key(self, key: str, ctrl: bool = False, shift: bool = False, alt: bool = False, vk: int | None = None, code: str | None = None) -> None:
        mods = (1 if alt else 0) | (2 if ctrl else 0) | (8 if shift else 0)
        if vk is None:
            vk = VK.get(key, ord(key.upper()) if len(key) == 1 else 0)
        if code is None:
            code = key if len(key) > 1 else f"Key{key.upper()}" if key.isalpha() and key.isascii() else f"Digit{key}" if key.isdigit() else ""
        base = {"key": key, "code": code, "windowsVirtualKeyCode": vk, "nativeVirtualKeyCode": vk, "modifiers": mods}
        text = key if len(key) == 1 and not (ctrl or alt) else ""
        self.cdp("Input.dispatchKeyEvent", {"type": "keyDown" if text else "rawKeyDown", **base, **({"text": text} if text else {})})
        self.cdp("Input.dispatchKeyEvent", {"type": "keyUp", **base})

    def type(self, text: str) -> None:
        for ch in text:
            self.cdp("Input.dispatchKeyEvent", {"type": "keyDown", "text": ch, "key": ch, "unmodifiedText": ch})
            self.cdp("Input.dispatchKeyEvent", {"type": "keyUp", "key": ch})
