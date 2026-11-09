"""A Raabe app's release exe (or a dev page) driven through the DevTools protocol: trusted mouse and
key events (Input.dispatch*), JavaScript, console errors and exceptions collected. Nothing touches
the real keyboard, mouse or screen: with `background` the window stays off screen and never takes
the focus, so these runs need no real-input lock (sweep_calc.py, D-141)."""
from __future__ import annotations

import json
import os
import subprocess
import tempfile
import time
import urllib.request
from pathlib import Path

import websocket

from . import winapi

ROOT = Path(__file__).resolve().parents[3]

KEYS = {"Enter": 13, "Backspace": 8, "Tab": 9, "Escape": 27, "Delete": 46, "Home": 36, "End": 35, "PageUp": 33,
        "PageDown": 34, "ArrowLeft": 37, "ArrowUp": 38, "ArrowRight": 39, "ArrowDown": 40, " ": 32, "Alt": 18,
        "F1": 112, "F2": 113, "F3": 114, "F4": 115, "F5": 116, "F7": 118, "F9": 120, "F11": 122, "F12": 123,
        "+": 187, "-": 189, "#": 191}
MODS = {"alt": 1, "ctrl": 2, "meta": 4, "shift": 8}


class CdpPage:
    """One page's DevTools connection; events (console errors, exceptions) are kept in `errors`."""

    def __init__(self, ws_url: str):
        self.ws = websocket.create_connection(ws_url, suppress_origin=True, timeout=120)
        self.n = 0
        self.errors: list[str] = []
        self.cdp("Runtime.enable")
        self.cdp("Log.enable")

    def _event(self, m: dict) -> None:
        meth = m.get("method")
        p = m.get("params", {})
        if meth == "Runtime.exceptionThrown":
            d = p.get("exceptionDetails", {})
            self.errors.append("exception: " + (d.get("exception", {}).get("description") or d.get("text") or "?")[:400])
        elif meth == "Runtime.consoleAPICalled" and p.get("type") in ("error", "assert"):
            self.errors.append("console: " + " ".join(str(a.get("value", a.get("description", ""))) for a in p.get("args", []))[:400])
        elif meth == "Log.entryAdded" and p.get("entry", {}).get("level") == "error":
            e = p["entry"]
            # Resource loads the page does not need (favicon) are not the app's errors.
            if "favicon" not in (e.get("url") or ""):
                self.errors.append("log: " + (e.get("text") or "")[:400])

    def cdp(self, method: str, params: dict | None = None):
        self.n += 1
        self.ws.send(json.dumps({"id": self.n, "method": method, "params": params or {}}))
        while True:
            m = json.loads(self.ws.recv())
            if m.get("id") == self.n:
                if "error" in m:
                    raise RuntimeError(m["error"])
                return m.get("result")
            self._event(m)

    def pump(self, seconds: float = 0.0) -> None:
        """Lets the page run and collects the events it sends meanwhile."""
        end = time.perf_counter() + seconds
        while True:
            self.cdp("Runtime.evaluate", {"expression": "0"})
            if time.perf_counter() >= end:
                return
            time.sleep(min(0.05, max(0.0, end - time.perf_counter())))

    def js(self, expr: str):
        r = self.cdp("Runtime.evaluate", {"expression": expr, "returnByValue": True, "awaitPromise": True})
        if "exceptionDetails" in r:
            d = r["exceptionDetails"]
            raise RuntimeError((d.get("exception", {}).get("description") or d.get("text") or json.dumps(d))[:1500])
        return r["result"].get("value")

    # ----- input -------------------------------------------------------------------------------
    def mouse(self, kind: str, x: float, y: float, button: str = "left", count: int = 1, mods: int = 0, buttons: int | None = None) -> None:
        if buttons is None:
            buttons = {"left": 1, "right": 2, "middle": 4}.get(button, 0) if kind == "mousePressed" else 0
        self.cdp("Input.dispatchMouseEvent", {"type": kind, "x": x, "y": y, "button": button if kind != "mouseMoved" else "none",
                                              "clickCount": count if kind != "mouseMoved" else 0, "buttons": buttons, "modifiers": mods})

    def move(self, x: float, y: float) -> None:
        self.mouse("mouseMoved", x, y)

    def click(self, x: float, y: float, count: int = 1, mods: str = "", button: str = "left") -> None:
        m = sum(MODS[p] for p in mods.split(",") if p)
        self.mouse("mouseMoved", x, y, mods=m)
        for c in range(1, count + 1):
            self.mouse("mousePressed", x, y, button=button, count=c, mods=m)
            self.mouse("mouseReleased", x, y, button=button, count=c, mods=m)

    def drag(self, x1: float, y1: float, x2: float, y2: float, steps: int = 10) -> None:
        self.mouse("mouseMoved", x1, y1)
        self.mouse("mousePressed", x1, y1)
        for i in range(1, steps + 1):
            self.mouse("mouseMoved", x1 + (x2 - x1) * i / steps, y1 + (y2 - y1) * i / steps, buttons=1)
        self.mouse("mouseReleased", x2, y2)

    def key(self, key: str, mods: str = "") -> None:
        """A key as Windows sends it: `mods` "ctrl,shift". Single letters/digits as themselves."""
        m = sum(MODS[p] for p in mods.split(",") if p)
        vk = KEYS.get(key, ord(key.upper()) if len(key) == 1 else 0)
        if len(key) == 1 and key.isalpha():
            code = f"Key{key.upper()}"
        elif len(key) == 1 and key.isdigit():
            code = f"Digit{key}"
        else:
            code = {" ": "Space", "+": "BracketRight", "-": "Slash", "#": "Backslash"}.get(key, key)
        base = {"key": key, "code": code, "windowsVirtualKeyCode": vk, "nativeVirtualKeyCode": vk, "modifiers": m}
        self.cdp("Input.dispatchKeyEvent", {"type": "rawKeyDown", **base})
        if len(key) == 1 and not (m & 3):
            self.cdp("Input.dispatchKeyEvent", {"type": "char", "text": key, "unmodifiedText": key, **base})
        self.cdp("Input.dispatchKeyEvent", {"type": "keyUp", **base})

    def type(self, text: str) -> None:
        for ch in text:
            if ch == "\n":
                self.key("Enter")
                continue
            self.cdp("Input.dispatchKeyEvent", {"type": "keyDown", "text": ch, "key": ch, "unmodifiedText": ch})
            self.cdp("Input.dispatchKeyEvent", {"type": "keyUp", "key": ch})

    def shot(self, path: Path, clip: tuple[float, float, float, float] | None = None) -> None:
        import base64
        params: dict = {"format": "png"}
        if clip:
            params["clip"] = {"x": clip[0], "y": clip[1], "width": clip[2], "height": clip[3], "scale": 1}
        path.write_bytes(base64.b64decode(self.cdp("Page.captureScreenshot", params)["data"]))


def _pid(hwnd: int) -> int:
    return winapi.window_pid(hwnd)


class AppExe:
    """The release exe of one app as a process of our own, off screen (RAABE_FEEL_BACKGROUND)."""

    def __init__(self, app: str = "calc", port: int = 9441, exe: Path | None = None, size: tuple[int, int] = (2000, 2050), background: bool = True):
        name = {"calc": "raabe-calc", "write": "raabe-write", "present": "raabe-present"}[app]
        env_name = {"calc": "RAABE_CALC_EXE", "write": "RAABE_WRITE_EXE", "present": "RAABE_PRESENT_EXE"}[app]
        exe = Path(exe or os.environ.get(env_name) or ROOT / f"apps/{name}/src-tauri/target/release/{name}.exe")
        rec = Path(tempfile.gettempdir()) / f"{name}-feel-recovery"
        for f in rec.glob("*") if rec.exists() else []:
            f.unlink(missing_ok=True)
        env = dict(os.environ, RAABE_FEEL_DEBUG_PORT=str(port))
        if background:
            env["RAABE_FEEL_BACKGROUND"] = "1"
        self.proc = subprocess.Popen([str(exe), "--feel", app], env=env)
        suffix = {"calc": " - RaabeCalc", "write": " - RaabeWrite", "present": " - RaabePresent"}[app]
        self.hwnd = winapi.wait_for_window(lambda h: winapi.window_title(h).endswith(suffix) and _pid(h) == self.proc.pid)
        if background:
            winapi.place_quietly(self.hwnd, -4000, 0, *size)
        else:
            winapi.place(self.hwnd, 0, 0, *size)
        deadline = time.perf_counter() + 30
        while True:
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=1) as r:
                    pages = [p for p in json.load(r) if p.get("type") == "page"]
                if pages:
                    self.page = CdpPage(pages[0]["webSocketDebuggerUrl"])
                    break
            except OSError:
                pass
            if time.perf_counter() > deadline:
                raise TimeoutError("DevTools endpoint not reachable")
            time.sleep(0.1)

    def close(self) -> None:
        try:
            self.page.ws.close()
        except Exception:  # noqa: BLE001
            pass
        self.proc.terminate()
        try:
            self.proc.wait(5)
        except subprocess.TimeoutExpired:
            self.proc.kill()


class DevPage:
    """A page already open in a headless browser with DevTools on `port` (devpage.py-like), for
    iterating on the dev server without rebuilding the exe."""

    def __init__(self, port: int, url_prefix: str):
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=2) as r:
            pages = [p for p in json.load(r) if p.get("type") == "page" and p.get("url", "").startswith(url_prefix)]
        self.page = CdpPage(pages[0]["webSocketDebuggerUrl"])
        self.proc = None

    def close(self) -> None:
        self.page.ws.close()
