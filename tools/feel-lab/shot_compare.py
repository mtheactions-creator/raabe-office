"""Screenshots of the same workbook in Excel and in RaabeCalc, side by side (D-066, method part 3).

    .venv/Scripts/python.exe shot_compare.py <file.xlsx> [sheet] [range] [--out name]

Excel (our own invisible instance) copies the range as it is shown on screen (CopyPicture), which
includes its shapes, charts and pictures; RaabeCalc (off screen) opens the file at 100 % and its
grid area of the same range is captured through DevTools. Both are written to TEMP and put next to
each other in <name>.png (default: shot-compare.png in TEMP) for a look by eye; the script also
prints a coarse difference score (mean absolute difference of the greyscale images scaled to the
same size). Uses the clipboard while Excel copies. Nothing is committed from TEMP.
"""
from __future__ import annotations

import base64
import io
import json
import sys
import tempfile
import time
from pathlib import Path

from PIL import Image, ImageChops, ImageGrab, ImageOps

from feellab.calc_catalogue import CalcSide
from feellab.excel import ExcelTarget

TMP = Path(tempfile.gettempdir())


def excel_shot(path: Path, sheet: str | None, rng: str) -> Image.Image:
    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(path))
        ws = wb.Worksheets(sheet) if sheet else wb.Worksheets(1)
        img = None
        for _ in range(10):
            try:
                ws.Range(rng).CopyPicture(1, 2)  # xlScreen, xlBitmap
                time.sleep(0.3)
                img = ImageGrab.grabclipboard()
                if isinstance(img, Image.Image):
                    break
            except Exception:  # noqa: BLE001
                time.sleep(0.5)
        wb.Close(False)
        if not isinstance(img, Image.Image):
            raise RuntimeError("Excel's picture did not reach the clipboard")
        return img.convert("RGB")
    finally:
        x.close()


def calc_shot(path: Path, sheet: str | None, rng: str) -> Image.Image:
    c = CalcSide()
    try:
        c.js(f"__calc.open({json.dumps(str(path))})")
        if sheet:
            c.js(f"(() => {{ const app = __calc.app; const i = app.wb.sheets.findIndex((s) => s.name === {json.dumps(sheet)}); if (i >= 0) app.showSheet ? app.showSheet(i) : (app.wb.active = i); }})()")
        c.js("(() => { const v = __calc.app.sheet.view; v.zoom = 100; v.top = 0; v.left = 0; __calc.app.grid.contentChanged(); })()")
        a, b = (rng.split(":") + [rng])[:2]
        time.sleep(0.6)
        r0 = c.js(f"__calc.point({json.dumps(a)}, 'tl')")
        r1 = c.js(f"__calc.point({json.dumps(b)}, 'br')")
        x0, y0 = r0["x"], r0["y"]
        c.ws.send(json.dumps({"id": 99001, "method": "Page.captureScreenshot",
                              "params": {"format": "png", "clip": {"x": x0, "y": y0, "width": r1["x"] - x0, "height": r1["y"] - y0, "scale": 1}}}))
        while True:
            msg = json.loads(c.ws.recv())
            if msg.get("id") == 99001:
                break
        return Image.open(io.BytesIO(base64.b64decode(msg["result"]["data"]))).convert("RGB")
    finally:
        c.close()


def main() -> None:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    out = TMP / "shot-compare.png"
    if "--out" in sys.argv:
        out = TMP / (sys.argv[sys.argv.index("--out") + 1] + ".png")
        args = [a for a in args if a != sys.argv[sys.argv.index("--out") + 1]]
    path = Path(args[0]).resolve()
    sheet = args[1] if len(args) > 1 and args[1] != "-" else None
    rng = args[2] if len(args) > 2 else "A1:L24"
    ex = excel_shot(path, sheet, rng)
    ca = calc_shot(path, sheet, rng)
    ex.save(TMP / "shot-excel.png")
    ca.save(TMP / "shot-calc.png")
    side = Image.new("RGB", (ex.width + ca.width + 12, max(ex.height, ca.height)), "white")
    side.paste(ex, (0, 0))
    side.paste(ca, (ex.width + 12, 0))
    side.save(out)
    g1 = ImageOps.grayscale(ex)
    g2 = ImageOps.grayscale(ca.resize(ex.size))
    diff = ImageChops.difference(g1, g2)
    score = sum(i * n for i, n in enumerate(diff.histogram())) / (ex.width * ex.height)
    print(json.dumps({"excel": list(ex.size), "calc": list(ca.size), "meanDiff": round(score, 2), "out": out.name}))


if __name__ == "__main__":
    main()
