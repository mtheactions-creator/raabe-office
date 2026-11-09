"""Compare: every chart type as RaabeCalc draws it against Excel's own picture (D-086).

Excel's side comes from probe_excel_charttypes.py (Chart.Export PNGs in TEMP/raabe-charttypes; the
probe is run when they are missing). Our side, two ways:
- `dev` (default): the chart engine draws each chart of Excel's workbook (packages/ooxml/test/
  fixtures/excel-charttypes.xlsx) as SVG (packages/charts/test/charttypes.test.ts with
  RAABE_CHARTTYPES_SVG), and a headless Edge with a profile of its own in TEMP turns it into a PNG;
- `calc`: the release exe (RAABE_CALC_EXE for another build) opens the same workbook and each chart's
  element is captured through DevTools.
Per type: the mean absolute difference of the greyscale pictures scaled to the same size (0 = equal,
255 = opposite) and side-by-side pictures in TEMP/raabe-charttypes-compare. No input, nothing visible
of Excel; `calc` opens RaabeCalc's window (off to the side, no input).
Result: results/2026-11-08-calc-charttypes.json (scores per type and mode).

    .venv/Scripts/python.exe compare_calc_charttypes.py [dev|calc] [type ...]
"""
from __future__ import annotations

import base64
import io
import json
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from PIL import Image, ImageChops, ImageOps

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
TMP = Path(tempfile.gettempdir())
EXCEL = TMP / "raabe-charttypes"
# RAABE_CHARTTYPES_TAG keeps parallel runs (several sessions on one PC) in folders of their own.
_TAG = os.environ.get("RAABE_CHARTTYPES_TAG", "")
OURS = TMP / f"raabe-charttypes-ours{_TAG}"
OUT = TMP / f"raabe-charttypes-compare{_TAG}"
RESULT = HERE / "results" / "2026-11-08-calc-charttypes.json"
FIXTURE = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "excel-charttypes.xlsx"
EDGE = [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"]


def flat(im: Image.Image) -> Image.Image:
    im = im.convert("RGBA")
    bg = Image.new("RGBA", im.size, "white")
    return Image.alpha_composite(bg, im).convert("RGB")


def score(a: Image.Image, b: Image.Image) -> float:
    g1 = ImageOps.grayscale(a)
    g2 = ImageOps.grayscale(b.resize(a.size))
    diff = ImageChops.difference(g1, g2)
    return round(sum(i * n for i, n in enumerate(diff.histogram())) / (a.width * a.height), 2)


def dev_pictures(names: list[str]) -> dict[str, Image.Image]:
    OURS.mkdir(exist_ok=True)
    env = dict(os.environ, RAABE_CHARTTYPES_SVG=str(OURS))
    r = subprocess.run("npx vitest run packages/charts/test/charttypes.test.ts", cwd=ROOT, env=env, shell=True, capture_output=True, text=True, errors="replace")
    if r.returncode:
        print(r.stdout[-3000:])
        raise SystemExit("the chart engine's test failed")
    edge = next(p for p in EDGE if Path(p).exists())
    profile = TMP / f"raabe-edge-headless{_TAG}"
    out = {}
    for n in names:
        png = OURS / f"{n}.png"
        if png.exists():
            png.unlink()
        subprocess.run([edge, "--headless=new", "--disable-gpu", "--hide-scrollbars", f"--user-data-dir={profile}", "--force-device-scale-factor=1",
                        "--window-size=722,434", f"--screenshot={png}", (OURS / f"{n}.html").as_uri()], capture_output=True, timeout=180)
        if png.exists():
            out[n] = flat(Image.open(png))
    return out


def calc_pictures(names: list[str]) -> dict[str, Image.Image]:
    from feellab.calc_catalogue import CalcSide
    c = CalcSide(port=9467)
    out = {}
    try:
        c.js(f"__calc.open({json.dumps(str(FIXTURE))})")
        time.sleep(1.5)
        for n in names:
            idx = c.js(f"__calc.app.wb.sheets.findIndex((s) => s.name === {json.dumps(n)})")
            if idx is None or idx < 0:
                continue
            c.js(f"(() => {{ const app = __calc.app; app.showSheet ? app.showSheet({idx}) : (app.wb.active = {idx}); const v = app.sheet.view; v.zoom = 100; v.top = 0; v.left = 0; app.grid.contentChanged(); }})()")
            time.sleep(0.8)
            box = c.js("(() => { const el = document.querySelector('.chart-obj svg'); if (!el) return null; const r = el.getBoundingClientRect(); return { x: r.x, y: r.y, w: r.width, h: r.height }; })()")
            if not box:
                continue
            c._id += 1
            c.ws.send(json.dumps({"id": c._id, "method": "Page.captureScreenshot",
                                  "params": {"format": "png", "clip": {"x": box["x"], "y": box["y"], "width": box["w"], "height": box["h"], "scale": 2}}}))
            while True:
                msg = json.loads(c.ws.recv())
                if msg.get("id") == c._id:
                    break
            out[n] = flat(Image.open(io.BytesIO(base64.b64decode(msg["result"]["data"]))))
    finally:
        c.close()
    return out


def main() -> None:
    args = sys.argv[1:]
    mode = args.pop(0) if args and args[0] in ("dev", "calc") else "dev"
    if not EXCEL.exists() or not any(EXCEL.glob("*.png")):
        subprocess.run([sys.executable, str(HERE / "probe_excel_charttypes.py")], cwd=HERE, check=True)
    names = args or sorted(p.stem for p in EXCEL.glob("*.png"))
    ours = dev_pictures(names) if mode == "dev" else calc_pictures(names)
    OUT.mkdir(exist_ok=True)
    scores = {}
    for n in names:
        if n not in ours:
            print(f"{n:28s} missing")
            continue
        ex = flat(Image.open(EXCEL / f"{n}.png"))
        ca = ours[n]
        scores[n] = score(ex, ca)
        side = Image.new("RGB", (ex.width * 2 + 12, ex.height), "white")
        side.paste(ex, (0, 0))
        side.paste(ca.resize(ex.size), (ex.width + 12, 0))
        side.save(OUT / f"{n}-{mode}.png")
        print(f"{n:28s} {scores[n]:6.2f}")
    data = json.loads(RESULT.read_text(encoding="utf-8")) if RESULT.exists() else {}
    data.setdefault("about", "Mean absolute greyscale difference (0-255) between Excel's Chart.Export PNG and RaabeCalc's drawing of the same chart, scaled to the same size; dev = chart engine via headless Edge, calc = the RaabeCalc exe.")
    data.setdefault(mode, {}).update(scores)
    RESULT.write_text(json.dumps(data, indent=1, sort_keys=True), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
