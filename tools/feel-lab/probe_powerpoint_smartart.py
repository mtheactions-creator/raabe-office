"""SmartArt as PowerPoint lays it out (Step 8b, D-077). An invisible PowerPoint of our own (only when no
POWERPNT.EXE runs) lists its SmartArt layouts (name, category, the layout's id) and puts each on a slide of its
own (a 640 x 360 pt frame), with three top-level items "Eins", "Zwei", "Drei" (and each item's own sub-items
removed), then saves the presentation and reads PowerPoint's drawing of each (the shapes it made: geometry,
position, size relative to the frame, and which item's text each carries).

    .venv/Scripts/python.exe probe_powerpoint_smartart.py [--limit N]

Result: results/<date>-powerpoint-smartart.json (names, ids and numbers only: the layout definitions in the saved
file are PowerPoint's content and stay in captures/, which is not committed).
"""
from __future__ import annotations

import datetime
import json
import re
import shutil
import sys
import time
import zipfile
from pathlib import Path

from feellab import winapi
from feellab.powerpoint import PowerPointTarget
from check_present_screens import display, monitors

HERE = Path(__file__).parent
CAP = HERE / "captures" / "smartart"
FRAME = (40.0, 60.0, 640.0, 360.0)
TEXTS = ["Eins", "Zwei", "Drei", "Vier", "Fünf", "Sechs", "Sieben", "Acht"]
COUNT = 3  # items per layout (--items 5 for the second pass)
EMU = 12700


def drawing_shapes(xml: str) -> list[dict]:
    """The shapes of a drawing part: preset geometry, box (EMU) and text."""
    out = []
    for sp in re.findall(r"<dsp:sp\b.*?</dsp:sp>", xml, re.S):
        off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"', sp)
        ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', sp)
        prst = re.search(r'prstGeom prst="(\w+)"', sp)
        rot = re.search(r'<a:xfrm[^>]*rot="(-?\d+)"', sp)
        text = "".join(re.findall(r"<a:t>([^<]*)</a:t>", sp))
        if off and ext:
            out.append({"prst": prst.group(1) if prst else "cust", "x": int(off.group(1)), "y": int(off.group(2)), "w": int(ext.group(1)), "h": int(ext.group(2)),
                        "rot": int(rot.group(1)) / 60000 if rot else 0, "text": text})
    return out


def main() -> None:
    global COUNT
    limit = int(sys.argv[sys.argv.index("--limit") + 1]) if "--limit" in sys.argv else 10_000
    COUNT = int(sys.argv[sys.argv.index("--items") + 1]) if "--items" in sys.argv else 3
    suffix = "" if COUNT == 3 else f"-{COUNT}"
    CAP.mkdir(parents=True, exist_ok=True)
    out: dict = {"date": datetime.date.today().isoformat(), "frame_pt": FRAME, "layouts": []}
    # SmartArtLayouts needs a document window: PowerPoint's window goes to the virtual display (D-073), so it never
    # covers the owner's screen; the display is switched off again at the end.
    before = monitors()
    mons = display("extend", 2)
    virtual = [m for m in mons if m not in before] or mons[1:]
    t = PowerPointTarget(visible=True, place=False)
    res = HERE / "results" / f"{out['date']}-powerpoint-smartart{suffix}.json"
    try:
        first = t.new_presentation(with_window=True, name="smartart")
        if virtual and t.hwnd:
            vx, vy, vw, vh = virtual[0]
            t.app.WindowState = 1
            winapi.place(t.hwnd, vx, vy, vw, vh)
        lays = t.app.SmartArtLayouts
        n = min(int(lays.Count), limit)
        start = int(sys.argv[sys.argv.index("--start") + 1]) if "--start" in sys.argv else 1
        if start > 1 and res.exists():
            out["layouts"] = [x for x in json.loads(res.read_text(encoding="utf-8"))["layouts"] if x["index"] < start]
        # Batches of 30 layouts per presentation: saved, read and closed before the next (a large presentation
        # slows PowerPoint down), the results written after each batch.
        for b0 in range(start, n + 1, 30):
            pres = t.new_presentation(with_window=True, name=f"smartart{b0}")
            batch = []
            for i in range(b0, min(n, b0 + 29) + 1):
                lay = lays(i)
                info = {"index": i, "name": str(lay.Name), "id": str(lay.Id), "category": str(lay.Category)}
                try:
                    sl = pres.Slides.Add(pres.Slides.Count + 1, 12)  # ppLayoutBlank
                    sh = sl.Shapes.AddSmartArt(lay, *FRAME)
                    sa = sh.SmartArt
                    top = [nd for nd in (sa.AllNodes(k) for k in range(1, sa.AllNodes.Count + 1)) if int(nd.Level) == 1]
                    info["default_items"] = len(top)
                    while sa.AllNodes.Count and len(top) > COUNT:
                        top[-1].Delete(); top = top[:-1]
                    tries = 0
                    while len(top) < COUNT and tries < COUNT + 1:
                        # Some layouts add the new item below another (Nodes.Add): a few tries, then as it is.
                        tries += 1
                        sa.Nodes.Add(); top = [nd for nd in (sa.AllNodes(k) for k in range(1, sa.AllNodes.Count + 1)) if int(nd.Level) == 1]
                    for k in range(sa.AllNodes.Count, 0, -1):
                        if int(sa.AllNodes(k).Level) > 1:
                            sa.AllNodes(k).Delete()
                    for nd, txt in zip(top, TEXTS[:COUNT]):
                        nd.TextFrame2.TextRange.Text = txt
                    info["slide"] = int(sl.SlideIndex)
                except Exception as e:  # noqa: BLE001
                    info["error"] = str(e)[:160]
                batch.append(info)
            path = CAP / f"smartart-probe{suffix}-{b0}.pptx"
            t.save_copy(pres, path)
            with zipfile.ZipFile(path) as z:
                names = set(z.namelist())
                for info in batch:
                    if "slide" not in info:
                        continue
                    rels = z.read(f"ppt/slides/_rels/slide{info['slide']}.xml.rels").decode()
                    d = re.search(r'Target="\.\./diagrams/(drawing\d+\.xml)"', rels)
                    if d and f"ppt/diagrams/{d.group(1)}" in names:
                        info["shapes"] = drawing_shapes(z.read(f"ppt/diagrams/{d.group(1)}").decode("utf-8"))
            for i in range(1, pres.Slides.Count + 1):
                t.export_png(pres, i, CAP / f"smartart{suffix}-{b0 + i - 1}.png", 960)
            pres.Saved = -1
            pres.Close()
            t.own.remove(pres)
            out["layouts"] += batch
            res.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
            print("batch", b0, "done", flush=True)
        del first
    finally:
        out["close"] = t.close(kill_if_unreachable=True)
        display("internal", 1)
    res.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    ok = sum(1 for x in out["layouts"] if x.get("shapes"))
    print(f"{len(out['layouts'])} layouts, {ok} with a drawing")


if __name__ == "__main__":
    main()
