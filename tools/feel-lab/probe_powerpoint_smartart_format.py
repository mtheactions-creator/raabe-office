"""Probe: what the SmartArt Format tab's colour commands do with the whole SmartArt selected (D-154), through COM on
the SmartArt's ShapeRange as the ribbon's commands act on the selection: Fülleffekt, Formkontur, Textfüllung, and
the same on one node's shape. SmartArtLayouts needs a document window, so our PowerPoint shows one: the run holds the
real-input lock (no input is sent, but the window must not come up under another session's real-input run); it
refuses to run while a POWERPNT.EXE runs (the owner's).

    .venv/Scripts/python.exe probe_powerpoint_smartart_format.py

Result: results/<date>-powerpoint-smartart-format.json (what changed in the diagram's parts, as XML excerpts).
"""
from __future__ import annotations

import datetime
import json
import re
import time
import zipfile
from pathlib import Path

from feellab.inputlock import real_input
from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget

HERE = Path(__file__).parent


def parts(path: Path) -> dict:
    with zipfile.ZipFile(path) as z:
        out = {}
        for n in z.namelist():
            if re.match(r"ppt/(diagrams/(data|drawing)\d+\.xml|slides/slide\d+\.xml)$", n):
                out[n] = z.read(n).decode("utf-8")
        return out


def excerpt(xml: str) -> dict:
    return {
        "bg": re.findall(r"<dgm:bg>.*?</dgm:bg>|<dgm:bg/>", xml, re.S)[:1],
        "whole": re.findall(r"<dgm:whole>.*?</dgm:whole>|<dgm:whole/>", xml, re.S)[:1],
        "pt_spPr": [x for x in re.findall(r"<dgm:spPr>.*?</dgm:spPr>|<dgm:spPr/>", xml, re.S) if x != "<dgm:spPr/>"][:6],
        "dsp_spPr_fill": re.findall(r"<dsp:spPr>.*?</dsp:spPr>", xml, re.S)[:2],
        "rPr": re.findall(r"<a:rPr[^>]*>(?:(?!</a:rPr>).)*</a:rPr>", xml, re.S)[:3],
        "p_spPr": re.findall(r"<p:graphicFrame>.*?</p:nvGraphicFramePr>", xml, re.S)[:1],
    }


def main() -> None:
    out: dict = {"date": datetime.date.today().isoformat(), "built_by": "PowerPoint (Microsoft 365, German UI), COM"}
    with real_input("present-followup smartart-format", wait=1200):
        run(out)
    p = HERE / "results" / f"{out['date']}-powerpoint-smartart-format.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print("->", p.name)


def run(out: dict) -> None:
    try:
        t = PowerPointTarget(visible=True)
    except OwnerPowerPointRunning as e:
        print("skipped:", e)
        return
    try:
        cases = {
            "frame_fill_red": lambda sh: setattr(sh.Fill.ForeColor, "RGB", 0x0000FF),
            "frame_line_green": lambda sh: setattr(sh.Line.ForeColor, "RGB", 0x00FF00),
            "frame_textfill_red": lambda sh: setattr(sh.GroupItems(1).TextFrame2.TextRange.Font.Fill.ForeColor, "RGB", 0x0000FF),
            "frame_fill_shaperange": lambda sh: setattr(sh.Parent.Shapes.Range([sh.Name]).Fill.ForeColor, "RGB", 0x0000FF),
            "node1_fill_red": lambda sh: setattr(sh.SmartArt.AllNodes(1).Shapes(1).Fill.ForeColor, "RGB", 0x0000FF),
            "node1_larger": lambda sh: sh.SmartArt.AllNodes(1).Larger(),
        }
        for name, fn in cases.items():
            pres = t.new_presentation(with_window=True, name=name)
            s = pres.Slides.AddSlide(1, t.layout(pres, 7))
            sh = s.Shapes.AddSmartArt(t.app.SmartArtLayouts(1), 120, 100, 500, 300)
            base = t.tmp / f"{name}-0.pptx"
            t.save_copy(pres, base)
            try:
                fn(sh)
                time.sleep(1.0)
                after = t.tmp / f"{name}-1.pptx"
                t.save_copy(pres, after)
                a, b = parts(base), parts(after)
                out[name] = {n: excerpt(b[n]) for n in b if a.get(n) != b[n]}
            except Exception as e:  # noqa: BLE001
                # (No local paths in the result: the repository is public.)
                out[name] = "error: " + re.sub(r"[A-Za-z]:\\[^'\"]*", "<path>", str(e))[:200]
            print(name, "done", flush=True)
    finally:
        out["close"] = t.close(kill_if_unreachable=True)


if __name__ == "__main__":
    main()
