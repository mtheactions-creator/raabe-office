"""RaabeWrite's drawing of Word's shapes against Word's (D-038, package C): the fixture
word-shapes.docx exported as PDF by both, the painted outlines compared shape by shape
(page-relative shapes, so both put them at the same place). Uses the release exe; its
window is moved but nothing is typed.

    .venv/Scripts/python.exe check_shapes_render.py
"""
from __future__ import annotations

import json
import tempfile
import time
from pathlib import Path

import check_shapes as cs
from feellab import pdftext, winapi
from feellab.targets import WINDOW_RECT, raabe_write
from probe_word_shapes import outline

HERE = Path(__file__).parent
FIXTURE = HERE.parents[1] / "packages" / "ooxml" / "test" / "fixtures" / "word-shapes.docx"
TMP = Path(tempfile.gettempdir()) / "raabe-shapes"


def ours(pdf_path: Path) -> list[dict]:
    t = raabe_write("short", port=9380)
    try:
        winapi.place(t.hwnd, *WINDOW_RECT)
        t.js(f"__app.openPath({str(FIXTURE)!r})")
        time.sleep(1.5)
        pdf_path.unlink(missing_ok=True)
        t.js(f"__app.exportPdfTo({str(pdf_path)!r})")
        for _ in range(40):
            if pdf_path.exists() and pdf_path.stat().st_size > 1000:
                break
            time.sleep(0.25)
        time.sleep(0.5)
    finally:
        t.close()
    p = pdftext.Pdf(pdf_path.read_bytes())
    num = p.pages()[0]
    return pdftext.paths(p.contents(num), p.page_box(num)[1])


def main() -> None:
    TMP.mkdir(exist_ok=True)
    word_json = TMP / "fixture-word.json"
    outline(FIXTURE, word_json)
    word = json.loads(word_json.read_text(encoding="utf-8"))[0]["paths"]
    mine = ours(TMP / "fixture-ours.pdf")
    # Shapes are where Word drew filled or stroked paths away from the text (big enough).
    def boxes(paths):
        out = []
        for p in paths:
            pts = cs.flatten(p["d"])
            if not pts:
                continue
            xs = [q[0] for q in pts]
            ys = [q[1] for q in pts]
            out.append({"box": (min(xs), min(ys), max(xs), max(ys)), "pts": pts, "fill": p.get("fill"), "stroke": p.get("stroke"), "paint": p["paint"]})
        return out
    wb = [b for b in boxes(word) if (b["box"][2] - b["box"][0]) > 8 or (b["box"][3] - b["box"][1]) > 8]
    ob = boxes(mine)
    result = []
    for w in wb:
        cx, cy = (w["box"][0] + w["box"][2]) / 2, (w["box"][1] + w["box"][3]) / 2
        near = [o for o in ob if abs((o["box"][0] + o["box"][2]) / 2 - cx) < 12 and abs((o["box"][1] + o["box"][3]) / 2 - cy) < 12]
        pts = [q for o in near for q in o["pts"]]
        dev = round(cs.deviation(w["pts"], pts), 2) if pts else None
        colors = sorted({o["fill"] or o["stroke"] or "" for o in near})
        result.append({"at": [round(cx), round(cy)], "paint": w["paint"], "word": w["fill"] or w["stroke"], "ours": colors, "deviation": dev})
    for r in result:
        print(r)
    (HERE / "results" / "2026-10-14-shapes-render.json").write_text(json.dumps(result, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
