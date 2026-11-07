"""Formkontur > Skizziert as PowerPoint writes it (Step 8, D-072). COM has no property for sketched lines, so
this probe uses real input in a PowerPoint of our own (under the lock, only when no POWERPNT.EXE runs):
a rectangle per style, selected through COM, then the keytips Alt, J D, V, S and the arrow keys to the
gallery entry (Keine, Gekrümmt, Freihand, Skizze). The presentation is saved in TEMP and each shape's a:ln
is recorded (XML of our own test file, not Microsoft prose).

    .venv/Scripts/python.exe probe_powerpoint_sketch.py

Result: results/<date>-powerpoint-sketch.json.
"""
from __future__ import annotations

import datetime
import json
import re
import time
import zipfile
from pathlib import Path

from catalogue_present import LOCK, take_lock
from feellab import winapi
from feellab.present_catalogue import Case, PowerPointSide

HERE = Path(__file__).parent


def main() -> None:
    if not take_lock("step8-sketch"):
        raise SystemExit("real-input lock held by another session: try later")
    side = None
    out: dict = {"date": datetime.date.today().isoformat(), "styles": {}}
    try:
        side = PowerPointSide()
        side.reset(Case("sketch", "sketch", [], []))
        sl = side.app.ActiveWindow.View.Slide
        names = ["Gekrümmt", "Freihand", "Skizze"]
        shapes = []
        for i, n in enumerate(names):
            s = sl.Shapes.AddShape(1, 60 + i * 290, 150, 240, 160)
            s.Name = n
            shapes.append(s)
        for i, s in enumerate(shapes):
            winapi.activate(side.t.hwnd)
            s.Select()
            time.sleep(0.4)
            # Stärke and Skizziert share the keytip S: the second S moves to Skizziert, Right opens it.
            for k in ("alt", "j", "d", "v", "s", "s", "right"):
                winapi.press(k)
                time.sleep(0.35)
            if i == 0:
                # What the screen shows at this point (captures/ is not committed).
                (HERE / "captures").mkdir(exist_ok=True)
                x0, y0, x1, y1 = winapi.window_rect(side.t.hwnd)
                winapi.save_region_png(str(HERE / "captures" / "sketch-menu.png"), x0, y0, x1 - x0, y1 - y0)
            # The list: Keine is first; one Down per entry after it.
            for _ in range(i + 1):
                winapi.press("down")
                time.sleep(0.2)
            winapi.press("enter")
            time.sleep(0.6)
            LOCK.touch()
        pres = side.pres
        path = side.t.tmp / "sketch-probe.pptx"
        side.t.save_copy(pres, path)
        with zipfile.ZipFile(path) as z:
            xml = z.read("ppt/slides/slide1.xml").decode("utf-8")
        for n in names:
            m = re.search(r'name="' + n + r'".*?</p:sp>', xml, re.S)
            body = m.group(0) if m else ""
            ln = re.search(r"<a:ln[ >].*?</a:ln>", body, re.S)
            # The sketch's path is long; keep its start and the element names around it.
            text = ln.group(0) if ln else ""
            if not text:
                out.setdefault("debug", {})[n] = body[:700]
            out["styles"][n] = {"ln_start": text[:900], "ln_len": len(text), "tags": sorted(set(re.findall(r"<(\w+:\w+)", text)))}
    finally:
        if side:
            side.close()
        LOCK.unlink(missing_ok=True)
    p = HERE / "results" / f"{out['date']}-powerpoint-sketch.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print(json.dumps(out, ensure_ascii=False)[:3000])


if __name__ == "__main__":
    main()
