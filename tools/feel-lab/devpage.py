"""Drives the development build (npm run dev:write, http://localhost:5174) in a headless Edge
over the DevTools protocol: full-resolution screenshots, scripted mouse and keys, JavaScript.
For checking UI work between builds; nothing touches the real keyboard, mouse or screen.

    python devpage.py start [width height scale]   # headless Edge on the dev server (port 9400)
    python devpage.py stop
    python devpage.py open <file.docx>             # opens a document through the test API
    python devpage.py js "<expression>"            # prints the value (awaits promises)
    python devpage.py shot <out.png> [x y w h]     # CSS pixels of the page
    python devpage.py click x y [count] [mods]     # mods: ctrl,shift,alt
    python devpage.py dblclick x y
    python devpage.py drag x1 y1 x2 y2 [steps]
    python devpage.py move x y
    python devpage.py key <Key> [mods]             # Enter, ArrowLeft, a, Delete …
    python devpage.py type "<text>"
    python devpage.py reload
"""
from __future__ import annotations

import base64
import json
import os
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path

import websocket

PORT = 9400
URL = "http://localhost:5174/"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
STATE = Path(tempfile.gettempdir()) / "raabe-devpage.json"


def page_ws():
    with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/json/list", timeout=2) as r:
        pages = [p for p in json.load(r) if p.get("type") == "page" and p.get("url", "").startswith(URL)]
    return websocket.create_connection(pages[0]["webSocketDebuggerUrl"], suppress_origin=True, timeout=60)


class Page:
    def __init__(self):
        self.ws = page_ws()
        self.n = 0

    def cdp(self, method, params=None):
        self.n += 1
        self.ws.send(json.dumps({"id": self.n, "method": method, "params": params or {}}))
        while True:
            m = json.loads(self.ws.recv())
            if m.get("id") == self.n:
                if "error" in m:
                    raise RuntimeError(m["error"])
                return m.get("result")

    def js(self, expr):
        r = self.cdp("Runtime.evaluate", {"expression": expr, "returnByValue": True, "awaitPromise": True})
        if "exceptionDetails" in r:
            raise RuntimeError(json.dumps(r["exceptionDetails"])[:2000])
        return r["result"].get("value")

    def mouse(self, kind, x, y, button="left", count=1, mods=0):
        self.cdp("Input.dispatchMouseEvent", {"type": kind, "x": x, "y": y, "button": button, "clickCount": count,
                                              "buttons": 1 if kind == "mousePressed" or (kind == "mouseMoved" and button == "left-held") else 0,
                                              "modifiers": mods})


def mods_of(s):
    m = 0
    for p in (s or "").split(","):
        m |= {"alt": 1, "ctrl": 2, "meta": 4, "shift": 8}.get(p, 0)
    return m


KEYS = {"Enter": 13, "Backspace": 8, "Tab": 9, "Escape": 27, "Delete": 46, "Home": 36, "End": 35, "PageUp": 33, "PageDown": 34,
        "ArrowLeft": 37, "ArrowUp": 38, "ArrowRight": 39, "ArrowDown": 40, " ": 32, "F7": 118}


def main(a):
    cmd = a[1]
    if cmd == "start":
        w, h, s = (a[2:5] + ["2000", "1300", "1.5"][len(a[2:5]):]) if len(a) > 2 else ("2000", "1300", "1.5")
        prof = Path(tempfile.gettempdir()) / "raabe-devpage-profile"
        # InPrivate in a profile of its own: no sign-in with the Windows account, no sync.
        p = subprocess.Popen([EDGE, "--headless=new", "--inprivate", f"--remote-debugging-port={PORT}", f"--user-data-dir={prof}",
                              f"--window-size={w},{h}", f"--force-device-scale-factor={s}", "--no-first-run", "--disable-sync",
                              "--disable-features=msImplicitSignin,msEdgeSidebar,msEdgeAutoSignIn", URL],
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        STATE.write_text(json.dumps({"pid": p.pid}))
        for _ in range(100):
            try:
                pg = Page()
                if pg.js("typeof window.__app === 'object'"):
                    print("ready")
                    return
            except Exception:
                pass
            time.sleep(0.2)
        print("not ready")
        return
    if cmd == "stop":
        # Only the Edge processes of our own profile (the owner may have Edge open).
        subprocess.run(["powershell", "-NoProfile", "-Command",
                        "Get-CimInstance Win32_Process -Filter \"Name='msedge.exe'\" | Where-Object { $_.CommandLine -like '*raabe-devpage-profile*' }"
                        " | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"], capture_output=True)
        STATE.unlink(missing_ok=True)
        print("stopped")
        return
    pg = Page()
    if cmd == "js":
        print(json.dumps(pg.js(a[2]), ensure_ascii=False, indent=1))
    elif cmd == "open":
        path = Path(a[2]).resolve().as_posix()
        print(pg.js(f"(async () => {{ const r = await fetch('/@fs/{path}'); const b = new Uint8Array(await r.arrayBuffer());"
                    f" window.__app.openBytes({json.dumps(Path(path).name)}, b); await new Promise(r => setTimeout(r, 400)); return r.status; }})()"))
    elif cmd == "shot":
        params = {"format": "png", "captureBeyondViewport": False}
        if len(a) >= 7:
            params["clip"] = {"x": float(a[3]), "y": float(a[4]), "width": float(a[5]), "height": float(a[6]), "scale": 1}
        data = pg.cdp("Page.captureScreenshot", params)["data"]
        Path(a[2]).write_bytes(base64.b64decode(data))
        print(a[2])
    elif cmd in ("click", "dblclick"):
        x, y = float(a[2]), float(a[3])
        count = 2 if cmd == "dblclick" else int(a[4]) if len(a) > 4 else 1
        m = mods_of(a[5] if len(a) > 5 else "")
        pg.mouse("mouseMoved", x, y, mods=m)
        for c in range(1, count + 1):
            pg.mouse("mousePressed", x, y, count=c, mods=m)
            pg.mouse("mouseReleased", x, y, count=c, mods=m)
        time.sleep(0.15)
        print("ok")
    elif cmd == "move":
        pg.mouse("mouseMoved", float(a[2]), float(a[3]))
        print("ok")
    elif cmd == "down":  # press and hold (then "hold x y" moves, "up x y" releases)
        pg.mouse("mouseMoved", float(a[2]), float(a[3]))
        pg.mouse("mousePressed", float(a[2]), float(a[3]), mods=mods_of(a[4] if len(a) > 4 else ""))
        print("ok")
    elif cmd == "hold":
        x, y = float(a[2]), float(a[3])
        pg.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x, "y": y, "button": "left", "buttons": 1, "modifiers": mods_of(a[4] if len(a) > 4 else "")})
        time.sleep(0.05)
        print("ok")
    elif cmd == "up":
        pg.mouse("mouseReleased", float(a[2]), float(a[3]))
        print("ok")
    elif cmd == "drag":
        x1, y1, x2, y2 = map(float, a[2:6])
        steps = int(a[6]) if len(a) > 6 else 12
        pg.mouse("mouseMoved", x1, y1)
        pg.mouse("mousePressed", x1, y1)
        for i in range(1, steps + 1):
            x, y = x1 + (x2 - x1) * i / steps, y1 + (y2 - y1) * i / steps
            pg.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x, "y": y, "button": "left", "buttons": 1})
            time.sleep(0.01)
        pg.mouse("mouseReleased", x2, y2)
        time.sleep(0.15)
        print("ok")
    elif cmd == "key":
        key = a[2]
        m = mods_of(a[3] if len(a) > 3 else "")
        vk = KEYS.get(key, ord(key.upper()) if len(key) == 1 else 0)
        code = key if len(key) > 1 else f"Key{key.upper()}" if key.isalpha() else f"Digit{key}" if key.isdigit() else ""
        base = {"key": key, "code": code, "windowsVirtualKeyCode": vk, "nativeVirtualKeyCode": vk, "modifiers": m}
        pg.cdp("Input.dispatchKeyEvent", {"type": "rawKeyDown", **base})
        if len(key) == 1 and not (m & 3):
            pg.cdp("Input.dispatchKeyEvent", {"type": "char", "text": key, **base})
        pg.cdp("Input.dispatchKeyEvent", {"type": "keyUp", **base})
        time.sleep(0.05)
        print("ok")
    elif cmd == "type":
        for ch in a[2]:
            pg.cdp("Input.dispatchKeyEvent", {"type": "keyDown", "text": ch, "key": ch, "unmodifiedText": ch})
            pg.cdp("Input.dispatchKeyEvent", {"type": "keyUp", "key": ch})
        print("ok")
    elif cmd == "reload":
        pg.cdp("Page.reload")
        time.sleep(1.5)
        print("ok")
    else:
        raise SystemExit(__doc__)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv)
