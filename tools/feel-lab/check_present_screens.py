"""Referentenansicht on a second screen (Step 8b, D-074), checked in the real app without touching the owner's
screen: the virtual display driver (D-073) is switched on ("Erweitern") for the run and off again at the end
("Nur PC-Bildschirm"); RaabePresent runs off screen (RAABE_FEEL_BACKGROUND) and keeps its main window where it
is (__presentKeepWindow), so only the show window appears, on the virtual screen. Keys go through the
DevTools protocol (no real input).

    .venv/Scripts/python.exe check_present_screens.py

Checks: the show opens a full-screen window on the other monitor; the audience sees the slide and the main
window the Referentenansicht (current slide, next slide, notes, "Folie n von m"); keys pressed in the main
window move both; ink drawn on the Referentenansicht appears on the audience's slide; Esc closes the show
window. Result: results/<date>-present-screens.json; screenshots in captures/ (not committed).
"""
from __future__ import annotations

import base64
import ctypes
import ctypes.wintypes as wt
import datetime
import json
import os
import subprocess
import tempfile
import time
import urllib.request
from pathlib import Path

import websocket

from feellab import winapi

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
FIXTURE = ROOT / "packages/ooxml/test/fixtures/pptx-office-show.pptx"
PORT = 9371
SWITCH = Path(os.environ.get("WINDIR", r"C:\Windows")) / "System32" / "DisplaySwitch.exe"


def monitors() -> list[tuple[int, int, int, int]]:
    out: list[tuple[int, int, int, int]] = []
    proc = ctypes.WINFUNCTYPE(ctypes.c_int, wt.HMONITOR, wt.HDC, ctypes.POINTER(wt.RECT), wt.LPARAM)

    def cb(_h, _dc, r, _l):
        out.append((r.contents.left, r.contents.top, r.contents.right - r.contents.left, r.contents.bottom - r.contents.top))
        return 1
    ctypes.windll.user32.EnumDisplayMonitors(None, None, proc(cb), 0)
    return out


def display(mode: str, want: int) -> list[tuple[int, int, int, int]]:
    subprocess.run([str(SWITCH), f"/{mode}"], check=False)
    deadline = time.time() + 40
    while time.time() < deadline:
        m = monitors()
        if len(m) == want:
            time.sleep(1.5)
            return monitors()
        time.sleep(0.3)
    return monitors()


class Cdp:
    def __init__(self, port: int):
        deadline = time.time() + 30
        while True:
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=1) as r:
                    pages = [p for p in json.load(r) if p.get("type") == "page" and not p.get("url", "").startswith("about:blank")]
                if pages:
                    self.ws = websocket.create_connection(pages[0]["webSocketDebuggerUrl"], suppress_origin=True)
                    break
            except OSError:
                pass
            if time.time() > deadline:
                raise TimeoutError("DevTools endpoint not reachable")
            time.sleep(0.2)
        self.n = 0

    def call(self, method: str, params: dict | None = None):
        self.n += 1
        self.ws.send(json.dumps({"id": self.n, "method": method, "params": params or {}}))
        while True:
            msg = json.loads(self.ws.recv())
            if msg.get("id") == self.n:
                return msg.get("result", {})

    def js(self, expr: str):
        r = self.call("Runtime.evaluate", {"expression": expr, "returnByValue": True, "awaitPromise": True})
        if "exceptionDetails" in r:
            raise RuntimeError(json.dumps(r["exceptionDetails"])[:400])
        return r.get("result", {}).get("value")

    def key(self, key: str, code: int):
        for t in ("rawKeyDown", "keyUp"):
            self.call("Input.dispatchKeyEvent", {"type": t, "key": key, "code": key, "windowsVirtualKeyCode": code})
        time.sleep(0.5)


def main() -> None:
    out: dict = {"date": datetime.date.today().isoformat(), "checks": {}}
    exe = Path(os.environ.get("RAABE_PRESENT_EXE") or ROOT / "apps/raabe-present/src-tauri/target/release/raabe-present.exe")
    before = monitors()
    out["monitors_before"] = len(before)
    proc = None
    try:
        mons = display("extend", 2)
        out["monitors_during"] = len(mons)
        if len(mons) < 2:
            raise SystemExit("second screen did not appear")
        env = dict(os.environ, RAABE_FEEL_DEBUG_PORT=str(PORT), RAABE_FEEL_BACKGROUND="1")
        proc = subprocess.Popen([str(exe), "--feel", "present"], env=env)
        cdp = Cdp(PORT)
        while not cdp.js("typeof window.__present === 'object'"):
            time.sleep(0.1)
        cdp.js("window.__presentKeepWindow = true")
        b64 = base64.b64encode(FIXTURE.read_bytes()).decode()
        cdp.js(f"window.__present.loadBytes({json.dumps(b64)})")
        n = cdp.js("window.__present.app.deck.slides.length")
        time.sleep(0.5)
        cdp.key("F5", 116)
        time.sleep(2.5)
        st = cdp.js("window.__present.app.show?.state()")
        out["state_start"] = st
        out["checks"]["dual"] = bool(st and st.get("dual"))
        # The show window: a window of our process titled for the show, on the second monitor, full screen.
        wins = winapi.find_windows(lambda h: winapi.window_pid(h) == proc.pid and winapi.window_title(h) == "Bildschirmpräsentation")
        r4 = winapi.window_rect(wins[0]) if wins else None
        rect = (r4[0], r4[1], r4[2] - r4[0], r4[3] - r4[1]) if r4 else None
        other = [m for m in mons if m not in before] or mons[1:]
        out["show_window"] = rect
        out["other_monitor"] = other[0] if other else None
        out["checks"]["show_window_on_other_monitor"] = bool(rect and other and abs(rect[0] - other[0][0]) <= 2 and abs(rect[2] - other[0][2]) <= 2)
        audience = "window.__present.app.show.doc"
        out["audience_slides"] = cdp.js(f"{audience}.querySelectorAll('.show-stage .slide').length")
        out["presenter_text"] = cdp.js("document.querySelector('.presenter-ui')?.innerText.replace(/\\s+/g,' ').slice(0,200)")
        out["checks"]["presenter_view_in_main"] = bool(out["presenter_text"] and f"Folie 1 von {n}" in out["presenter_text"])
        out["checks"]["audience_has_slide"] = out["audience_slides"] == 1
        if wins:
            (HERE / "captures").mkdir(exist_ok=True)
            winapi.save_region_png(str(HERE / "captures" / "screens-audience.png"), *rect)
        # Keys in the main window move the show on both screens.
        steps = st.get("steps", 0) if st else 0
        for _ in range(steps + 1):
            cdp.key("ArrowRight", 39)
        st2 = cdp.js("window.__present.app.show?.state()")
        out["state_after_next"] = st2
        txt = cdp.js("document.querySelector('.presenter-ui')?.innerText.replace(/\\s+/g,' ')") or ""
        out["checks"]["next_moves_both"] = bool(st2 and st2["slide"] == 2 and "Folie 2 von" in txt)
        same = cdp.js(f"(() => {{ const a = {audience}.querySelector('.show-stage .slide'); const b = document.querySelector('.show.presenter.dual .show-stage .slide'); return !!a && !!b && a.querySelectorAll('.shp').length === b.querySelectorAll('.shp').length; }})()")
        out["checks"]["mirror_matches_audience"] = bool(same)
        # Ink on the Referentenansicht's slide shows on the audience's.
        # Strg+P: the pen.
        cdp.call("Input.dispatchKeyEvent", {"type": "rawKeyDown", "key": "Control", "code": "ControlLeft", "windowsVirtualKeyCode": 17, "modifiers": 2})
        cdp.call("Input.dispatchKeyEvent", {"type": "rawKeyDown", "key": "p", "code": "KeyP", "windowsVirtualKeyCode": 80, "modifiers": 2})
        cdp.call("Input.dispatchKeyEvent", {"type": "keyUp", "key": "p", "code": "KeyP", "windowsVirtualKeyCode": 80, "modifiers": 2})
        cdp.call("Input.dispatchKeyEvent", {"type": "keyUp", "key": "Control", "code": "ControlLeft", "windowsVirtualKeyCode": 17})
        box = cdp.js("(() => { const r = document.querySelector('.show.presenter.dual .show-ink').getBoundingClientRect(); return [r.left, r.top, r.width, r.height]; })()")
        x, y = box[0] + box[2] * 0.3, box[1] + box[3] * 0.5
        cdp.call("Input.dispatchMouseEvent", {"type": "mousePressed", "x": x, "y": y, "button": "left", "clickCount": 1})
        for k in range(1, 8):
            cdp.call("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x + k * 15, "y": y + k * 4, "button": "left", "buttons": 1})
        cdp.call("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": x + 105, "y": y + 28, "button": "left", "clickCount": 1})
        time.sleep(0.3)
        out["ink_paths"] = cdp.js(f"[document.querySelectorAll('.show.presenter.dual .show-ink path').length, {audience}.querySelectorAll('.show-ink path').length]")
        out["checks"]["ink_on_both"] = out["ink_paths"] == [1, 1]
        if wins:
            winapi.save_region_png(str(HERE / "captures" / "screens-audience-ink.png"), *rect)
        cdp.key("Escape", 27)
        time.sleep(1.0)
        out["checks"]["esc_ends"] = cdp.js("window.__present.app.show === null")
        left = winapi.find_windows(lambda h: winapi.window_pid(h) == proc.pid and winapi.window_title(h) == "Bildschirmpräsentation")
        out["checks"]["show_window_closed"] = not left
    finally:
        if proc:
            proc.kill()
        after = display("internal", 1)
        out["monitors_after"] = len(after)
    ok = sum(1 for v in out["checks"].values() if v)
    out["summary"] = f"{ok}/{len(out['checks'])}"
    p = HERE / "results" / f"{out['date']}-present-screens.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps(out)[:2500])


if __name__ == "__main__":
    main()
