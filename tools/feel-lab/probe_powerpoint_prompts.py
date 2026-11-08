"""The prompts PowerPoint shows in empty placeholders (D-080, D-102): a PowerPoint of our own (only when no
POWERPNT.EXE runs), its window on the virtual display (D-073), adds a slide with each layout of its default
design and reads, through UI Automation (no input sent), the names and values of the elements on the slide that
carry a prompt. Only strings with "hinzufügen" or "Klicken" are kept (labels, never the window's title bar).

    .venv/Scripts/python.exe probe_powerpoint_prompts.py

Result: results/<date>-powerpoint-prompts.json (per layout: the prompts with the placeholder's type).
"""
from __future__ import annotations

import datetime
import json
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.powerpoint import PowerPointTarget
from check_present_screens import display, monitors

HERE = Path(__file__).parent
PH = {1: "title", 2: "body", 3: "ctrTitle", 4: "subTitle", 7: "obj", 8: "chart", 12: "tbl", 13: "clipArt", 14: "dgm", 16: "media", 18: "pic", 5: "dt", 6: "ftr", 13 + 100: "sldNum"}


def walk(el, depth: int = 0, max_depth: int = 30):
    yield el
    if depth < max_depth:
        for c in uia.children(el):
            yield from walk(c, depth + 1, max_depth)


def prompts(root) -> list[str]:
    out = []
    for el in walk(root):
        d = uia.describe(el)
        for s in (d.get("name") or "", d.get("value") or ""):
            if ("hinzufügen" in s or "Klicken" in s) and s not in out and len(s) < 120:
                out.append(s)
    return out


def main() -> None:
    out: dict = {"date": datetime.date.today().isoformat(), "layouts": []}
    before = monitors()
    mons = display("extend", 2)
    virtual = [m for m in mons if m not in before] or mons[1:]
    t = PowerPointTarget(visible=True, place=False)
    try:
        pres = t.new_presentation(with_window=True, name="prompts")
        if virtual and t.hwnd:
            vx, vy, vw, vh = virtual[0]
            t.app.WindowState = 1
            winapi.place(t.hwnd, vx, vy, vw, vh)
        lays = pres.SlideMaster.CustomLayouts
        for i in range(1, int(lays.Count) + 1):
            lay = lays(i)
            sl = pres.Slides.AddSlide(pres.Slides.Count + 1, lay)
            t.app.ActiveWindow.View.GotoSlide(sl.SlideIndex)
            time.sleep(1.2)
            phs = []
            for k in range(1, int(sl.Shapes.Placeholders.Count) + 1):
                p = sl.Shapes.Placeholders(k)
                phs.append({"name": str(p.Name), "type": int(p.PlaceholderFormat.Type)})
            root = uia.element_for(t.hwnd)
            out["layouts"].append({"layout": str(lay.Name), "placeholders": phs, "prompts": prompts(root)})
            print(lay.Name, out["layouts"][-1]["prompts"], flush=True)
        pres.Saved = -1
    finally:
        out["close"] = t.close(kill_if_unreachable=True)
        display("internal", 1)
    p = HERE / "results" / f"{out['date']}-powerpoint-prompts.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))


if __name__ == "__main__":
    main()
