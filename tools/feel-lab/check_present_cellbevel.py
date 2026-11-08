"""Zellenabschrägung compared with PowerPoint (D-076, D-103). An invisible PowerPoint of our own (only when no
POWERPNT.EXE runs) makes a table of 3 x 4 cells, each with another of the twelve bevels (Tabellenentwurf >
Effekte > Zellenabschrägung, through COM: the cell's ThreeD.BevelTopType), saves it and exports its picture;
then RaabePresent (headless Edge on the dev server) draws the saved file and the two pictures are compared pixel
by pixel, the whole slide and each cell.

    .venv/Scripts/python.exe check_present_cellbevel.py      (npm run dev:present must be running)

Result: results/<date>-present-cellbevel.json; pictures in captures/cellbevel (not committed).
"""
from __future__ import annotations

import datetime
import io
import json
from pathlib import Path

from PIL import Image

from check_present_render import compare
from feellab.powerpoint import PowerPointTarget
from feellab.present_page import PresentPage

HERE = Path(__file__).parent
CAP = HERE / "captures" / "cellbevel"
# msoBevelType: 1 none, 2 relaxedInset, 3 circle, 4 slope, 5 cross, 6 angle, 7 softRound, 8 convex, 9 coolSlant,
# 10 divot, 11 riblet, 12 hardEdge, 13 artDeco.
BEVELS = [3, 2, 5, 9, 6, 7, 8, 4, 10, 11, 12, 13]
BOX = (60, 60, 840, 420)  # points: left, top, width, height


PRST = {2: "relaxedInset", 3: "circle", 4: "slope", 5: "cross", 6: "angle", 7: "softRound", 8: "convex", 9: "coolSlant", 10: "divot", 11: "riblet", 12: "hardEdge", 13: "artDeco"}


def add_bevels(src: Path, dst: Path) -> None:
    """The table's cells each with its bevel (a:cell3D in the cell's tcPr, before its fill), as we write it."""
    import re
    import zipfile
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            if item.filename == "ppt/slides/slide1.xml":
                xml = data.decode("utf-8")
                cells = iter(BEVELS)

                def tc(m):
                    cell = m.group(0)
                    b = f'<a:cell3D><a:bevel prst="{PRST[next(cells)]}"/></a:cell3D>'
                    if "<a:tcPr/>" in cell:
                        return cell.replace("<a:tcPr/>", f"<a:tcPr>{b}</a:tcPr>")
                    return re.sub(r"(<a:tcPr[^>]*>)((?:<a:ln\w+\b.*?</a:ln\w+>)*)", lambda q: q.group(1) + q.group(2) + b, cell, count=1, flags=re.S)
                xml = re.sub(r"<a:tc>.*?</a:tc>", tc, xml, flags=re.S)
                data = xml.encode("utf-8")
            zo.writestr(item, data)


def powerpoint() -> None:
    # PowerPointTarget refuses when another POWERPNT.EXE runs (the owner's).
    CAP.mkdir(parents=True, exist_ok=True)
    t = PowerPointTarget(visible=False)
    try:
        pres = t.new_presentation(with_window=False, name="cellbevel")
        sl = pres.Slides.Add(1, 12)
        tb = sl.Shapes.AddTable(3, 4, *BOX).Table
        for r in range(1, 4):
            for c in range(1, 5):
                tb.Cell(r, c).Shape.TextFrame.TextRange.Text = f"Z{r}{c}"
        t.save_copy(pres, CAP / "plain.pptx")
        pres.Saved = -1
        add_bevels(CAP / "plain.pptx", CAP / "cellbevel.pptx")
        bev = t.open_copy(CAP / "cellbevel.pptx")
        t.export_png(bev, 1, CAP / "powerpoint.png", 1280)
        bev.Saved = -1
    finally:
        t.close(kill_if_unreachable=True)


def main() -> None:
    if not (CAP / "cellbevel.pptx").exists() or "--record" in __import__("sys").argv:
        powerpoint()
    theirs = Image.open(CAP / "powerpoint.png").convert("RGB")
    page = PresentPage(1600, 1000, 1.0)
    try:
        page.open_file(CAP / "cellbevel.pptx")
        r = page.js(f"window.__present.renderAlone(1, {theirs.width})")
        ours = Image.open(io.BytesIO(page.shot(0, 0, r["w"], r["h"]))).convert("RGB")
    finally:
        page.close()
    out: dict = {"date": datetime.date.today().isoformat(), "slide": compare(theirs, ours), "cells": []}
    s = theirs.width / 960
    cw, ch = BOX[2] / 4, BOX[3] / 3
    for k in range(12):
        r, c = divmod(k, 4)
        box = tuple(int(v * s) for v in (BOX[0] + c * cw, BOX[1] + r * ch, BOX[0] + (c + 1) * cw, BOX[1] + (r + 1) * ch))
        out["cells"].append({"bevel": BEVELS[k], **compare(theirs.crop(box), ours.crop(box))})
    side = Image.new("RGB", (theirs.width * 2 + 10, theirs.height), "white")
    side.paste(theirs, (0, 0)); side.paste(ours, (theirs.width + 10, 0))
    side.save(CAP / "compare.png")
    p = HERE / "results" / f"{out['date']}-present-cellbevel.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print("slide", out["slide"], "cells", [c["mean"] for c in out["cells"]])


if __name__ == "__main__":
    main()
