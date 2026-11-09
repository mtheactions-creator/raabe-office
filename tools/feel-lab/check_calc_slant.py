"""Turned text slants its cell in the real app (D-125), without real input: RaabeCalc (release exe, RAABE_CALC_EXE for
another build) runs off screen; B2 gets text turned by 45 degrees with thin borders all round and a fill, D2 the same
turn without borders, F2 -45 degrees with borders. The grid is captured through DevTools and read: along the line from
the cell's bottom corners to the top moved by the row's height (Excel's PDF, probe_excel_rotated_borders.py) there
must be border pixels, where the upright border would be there must be none; D2 must stay a rectangle (its fill only
inside the cell).

    .venv/Scripts/python.exe check_calc_slant.py

Result: results/<date>-calc-slant.json.
"""
from __future__ import annotations

import base64
import datetime
import io
import json
import os
from pathlib import Path

from PIL import Image

from feellab import calc_catalogue, winapi
from feellab.calc_catalogue import CalcSide

HERE = Path(__file__).parent
os.environ["RAABE_FEEL_BACKGROUND"] = "1"
calc_catalogue.winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4000, 0, w, h)


def shot(c: CalcSide, a: str, b: str) -> tuple[Image.Image, dict, dict]:
    r0 = c.js(f"__calc.point({json.dumps(a)}, 'tl')")
    r1 = c.js(f"__calc.point({json.dumps(b)}, 'br')")
    c.ws.send(json.dumps({"id": 99101, "method": "Page.captureScreenshot",
                          "params": {"format": "png", "clip": {"x": r0["x"], "y": r0["y"], "width": r1["x"] - r0["x"], "height": r1["y"] - r0["y"], "scale": 1}}}))
    while True:
        msg = json.loads(c.ws.recv())
        if msg.get("id") == 99101:
            break
    return Image.open(io.BytesIO(base64.b64decode(msg["result"]["data"]))).convert("RGB"), r0, r1


def dark(p) -> bool:
    # Darker than the fill and the gridlines (a thin slanted line is drawn antialiased).
    return sum(p) < 450


def main() -> None:
    c = CalcSide(port=9479)
    report: dict = {"checks": []}

    def add(name, ok, detail=None):
        report["checks"].append({"check": name, "ok": bool(ok), "detail": detail})

    try:
        c.js("__calc.reset({}, 'A1', 1)")
        c.js("""(() => { const app = __calc.app, wb = app.wb, sh = app.sheet;
          const side = { style: 'thin', color: '#000000' };
          const mk = (rot, borders) => wb.styles.derive(undefined, (s) => { s.rotation = rot; s.fill = '#FFE5CC';
            if (borders) s.borders = { left: side, right: side, top: side, bottom: side }; });
          sh.set(1, 1, { v: 'Text', s: mk(45, true) });
          sh.set(1, 3, { v: 'Text', s: mk(45, false) });
          sh.set(1, 5, { v: 'Text', s: mk(135, true) });
          app.grid.contentChanged(); })()""")
        import time
        time.sleep(0.8)
        img, r0, r1 = shot(c, "A1", "H4")
        img.save(Path(os.environ.get("TEMP", ".")) / "raabe-calc-slant.png")
        scale = img.width / (r1["x"] - r0["x"])
        box = lambda addr: {k: (c.js(f"__calc.point({json.dumps(addr)}, {json.dumps(k)})")) for k in ("tl", "br")}

        def cell(addr):
            b = box(addr)
            return ((b["tl"]["x"] - r0["x"]) * scale, (b["tl"]["y"] - r0["y"]) * scale, (b["br"]["x"] - r0["x"]) * scale, (b["br"]["y"] - r0["y"]) * scale)

        def near_dark(x, y, r=2):
            return any(dark(img.getpixel((int(min(img.width - 1, max(0, x + dx))), int(min(img.height - 1, max(0, y + dy)))))) for dx in range(-r, r + 1) for dy in range(-r, r + 1))

        for addr, sign, label in (("B2", 1, "45 degrees"), ("F2", -1, "-45 degrees")):
            x0, y0, x1, y1 = cell(addr)
            h = y1 - y0
            hits = sum(near_dark(x0 + sign * h * t, y1 - h * t) for t in (0.25, 0.5, 0.75))
            upright = sum(near_dark(x0, y1 - h * t, 1) for t in (0.4, 0.6)) if sign > 0 else sum(near_dark(x0, y1 - h * t, 1) for t in (0.4, 0.6))
            add(f"{addr} ({label}): the left border runs along the text (3/3 points)", hits == 3,
                {"hits": hits, "box": [round(v, 1) for v in (x0, y0, x1, y1)], "size": [img.width, img.height]})
            add(f"{addr} ({label}): no upright left border in the middle", upright == 0, upright)
        x0, y0, x1, y1 = cell("D2")
        h = y1 - y0
        fill_out = img.getpixel((int(x1 + h * 0.5), int(y0 + h * 0.3)))
        add("D2 (no border): stays a rectangle, no fill beyond its right edge", fill_out != (255, 229, 204), list(fill_out))
    finally:
        c.close()
    ok = sum(1 for r in report["checks"] if r["ok"])
    print("checks", ok, "/", len(report["checks"]))
    for r in report["checks"]:
        if not r["ok"]:
            print("   FAIL", json.dumps(r, ensure_ascii=False))
    (HERE / "results" / f"{datetime.date.today().isoformat()}-calc-slant.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
