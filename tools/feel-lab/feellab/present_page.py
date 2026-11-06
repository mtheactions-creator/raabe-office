"""RaabePresent's development build (npm run dev:present, http://localhost:5176) in a headless Edge of
our own over the DevTools protocol (D-070): invisible, no real input. Used by the render comparison and
the invisible catalogue checks. Only Edge processes started with our profile are ever ended."""
from __future__ import annotations

import base64
import json
import subprocess
import tempfile
import time
import urllib.request
from pathlib import Path

import websocket

PORT = 9411
URL = "http://localhost:5176/"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
PROFILE = Path(tempfile.gettempdir()) / "raabe-present-page-profile"


class PresentPage:
    def __init__(self, width=1600, height=1000, scale=1.0):
        self.proc = subprocess.Popen([EDGE, "--headless=new", "--inprivate", f"--remote-debugging-port={PORT}", f"--user-data-dir={PROFILE}",
                                      f"--window-size={width},{height}", f"--force-device-scale-factor={scale}", "--no-first-run", "--disable-sync",
                                      "--disable-features=msImplicitSignin,msEdgeSidebar,msEdgeAutoSignIn", URL],
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        self.n = 0
        self.ws = None
        for _ in range(150):
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/json/list", timeout=2) as r:
                    pages = [p for p in json.load(r) if p.get("type") == "page" and p.get("url", "").startswith(URL)]
                if pages:
                    self.ws = websocket.create_connection(pages[0]["webSocketDebuggerUrl"], suppress_origin=True, timeout=120)
                    if self.js("typeof window.__present === 'object'"):
                        return
            except Exception:  # noqa: BLE001
                pass
            time.sleep(0.2)
        raise RuntimeError("RaabePresent dev page not ready (is npm run dev:present running?)")

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

    def open_file(self, path: Path) -> int:
        p = Path(path).resolve().as_posix()
        return self.js(f"(async () => {{ const r = await fetch('/@fs/{p}'); const b = new Uint8Array(await r.arrayBuffer());"
                       f" const n = window.__present.openBytes({json.dumps(Path(p).name)}, b); await new Promise(r => setTimeout(r, 300)); return n; }})()")

    def shot(self, x, y, w, h) -> bytes:
        data = self.cdp("Page.captureScreenshot", {"format": "png", "clip": {"x": x, "y": y, "width": w, "height": h, "scale": 1}})["data"]
        return base64.b64decode(data)

    def close(self):
        try:
            self.ws.close()
        except Exception:  # noqa: BLE001
            pass
        subprocess.run(["powershell", "-NoProfile", "-Command",
                        "Get-CimInstance Win32_Process -Filter \"Name='msedge.exe'\" | Where-Object { $_.CommandLine -like '*raabe-present-page-profile*' }"
                        " | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"], capture_output=True)
