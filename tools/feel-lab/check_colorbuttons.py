"""Check (Step 11, part 4, D-157): the colour split buttons look as Office's. Crops Schriftfarbe and Füllfarbe in Excel
and RaabeCalc, Schriftfarbe, Texthervorhebungsfarbe in Word and RaabeWrite, at the screen's scale (150 %), and measures
in each crop the colour bar (box and colour) and the glyph above it (box of the dark pixels). Fails when a bar differs
by more than 1 device pixel or its colour differs, when the glyph's bottom reaches into the bar, or when there is a
second bar. Real input (window in front for the screenshot) under the lock.

    .venv/Scripts/python.exe check_colorbuttons.py [calc] [write]
"""
from __future__ import annotations

import datetime
import json
import os
import sys
import tempfile
import time
from pathlib import Path

from PIL import Image

from feellab import uia, winapi
from feellab.inputlock import real_input

HERE = Path(__file__).parent
TMP = Path(tempfile.mkdtemp(prefix="raabe-colorbtn-"))


def measure(png: Path, main_w: int) -> dict:
    """Bar: the widest run of saturated colour rows; glyph: dark pixels above it. Only the main part (left `main_w`)."""
    im = Image.open(png).convert("RGB")
    w, h = im.size
    sat = lambda p: max(p) - min(p) > 120  # noqa: E731
    dark = lambda p: sum(p) < 330 and max(p) - min(p) < 60  # noqa: E731
    bar = [(x, y) for y in range(h) for x in range(min(w, main_w)) if sat(im.getpixel((x, y)))]
    rows = sorted({y for _, y in bar})
    # Separate bands of rows: a second band would be a second bar.
    bands, cur = [], []
    for y in rows:
        if cur and y != cur[-1] + 1:
            bands.append(cur)
            cur = []
        cur.append(y)
    if cur:
        bands.append(cur)
    # A bar is a band at least 2 rows high and 15 device pixels wide (a glyph's coloured detail, the blue drop of the
    # paint bucket, is narrower).
    width = lambda band: (lambda xs: max(xs) - min(xs) + 1)([x for x, y in bar if band[0] <= y <= band[-1]])  # noqa: E731
    bands = [b for b in bands if len(b) >= 2 and width(b) >= 15]
    out: dict = {"bands": len(bands)}
    if bands:
        b = max(bands, key=width)
        xs = [x for x, y in bar if b[0] <= y <= b[-1]]
        out["bar"] = [min(xs), b[0], max(xs), b[-1]]
        out["bar_color"] = list(im.getpixel(((min(xs) + max(xs)) // 2, (b[0] + b[-1]) // 2)))
    g = [(x, y) for y in range(h) for x in range(min(w, main_w)) if dark(im.getpixel((x, y)))]
    # Outline boxes (an empty frame under the glyph) are dark too: count rows that are a long horizontal line.
    lines = [y for y in range(h) if sum(1 for x in range(min(w, main_w)) if dark(im.getpixel((x, y)))) > 14]
    out["dark_lines"] = lines
    if g:
        out["glyph"] = [min(x for x, _ in g), min(y for _, y in g), max(x for x, _ in g), max(y for _, y in g)]
    return out


def office_side(app: str, names: list[str]) -> dict:
    res = {}
    if app == "excel":
        from feellab.excel import ExcelTarget
        t = ExcelTarget(visible=True)
        top, pid, close = t.hwnd, t.pid, t.close
    else:
        from feellab.targets import WordTarget
        t = WordTarget()
        top = winapi.user32.GetAncestor(int(t.app.ActiveWindow.Hwnd), 2)
        pid, close = winapi.window_pid(top), t.close
    try:
        time.sleep(1.5)
        winapi.activate(top)
        time.sleep(0.5)
        root = uia.element_for(top)
        for n in names:
            e = uia.find(root, lambda d: d["type"] == "SplitButton" and d["name"] == n, max_depth=14)
            if not e or winapi.window_pid(winapi.foreground()) != pid:
                res[n] = {"error": "not found or window not in front"}
                continue
            l, t_, w, h = uia.describe(e)["rect"]
            main = uia.find(e, lambda d: d["type"] == "Button", max_depth=2)
            mw = uia.describe(main)["rect"][2] if main else w - 22
            p = TMP / f"{app}-{n}.png"
            winapi.save_region_png(str(p), l, t_, w, h)
            res[n] = measure(p, mw)
    finally:
        close()
    return res


def raabe_side(app: str, names: list[str]) -> dict:
    res = {}
    if app == "calc":
        from feellab.calc_catalogue import CalcSide
        side = CalcSide()
        js, hwnd, close = side.js, side.hwnd, side.close
        ox, oy, dpr = side._origin()
    else:
        from feellab.targets import raabe_write
        side = raabe_write()
        js, hwnd, close = side.js, side.hwnd, side.close
        ox, oy = side._viewport_origin()
        dpr = side.js("devicePixelRatio")
    try:
        time.sleep(1.0)
        winapi.activate(hwnd)
        time.sleep(0.5)
        for n in names:
            r = js(f"""(() => {{ const s = [...document.querySelectorAll('.rsplit')].find((x) => x.querySelector('.color-main')?.getAttribute('aria-label') === {json.dumps(n)});
              if (!s) return null; const a = s.getBoundingClientRect(), m = s.querySelector('.color-main').getBoundingClientRect();
              return [a.left, a.top, a.width, a.height, m.width]; }})()""")
            if not r or winapi.window_pid(winapi.foreground()) != winapi.window_pid(hwnd):
                res[n] = {"error": "not found or window not in front"}
                continue
            p = TMP / f"{app}-{n}.png"
            winapi.save_region_png(str(p), round(ox + r[0] * dpr), round(oy + r[1] * dpr), round(r[2] * dpr), round(r[3] * dpr))
            res[n] = measure(p, round(r[4] * dpr))
    finally:
        close()
    return res


def compare(office: dict, ours: dict) -> list[str]:
    out = []
    if "error" in office or "error" in ours:
        return [f"error: {office.get('error') or ours.get('error')}"]
    if ours["bands"] != 1:
        out.append(f"{ours['bands']} colour bars")
    if "bar" in office and "bar" in ours:
        ob, rb = office["bar"], ours["bar"]
        dw, dh = (rb[2] - rb[0]) - (ob[2] - ob[0]), (rb[3] - rb[1]) - (ob[3] - ob[1])
        if abs(dw) > 1 or abs(dh) > 1:
            out.append(f"bar size {rb[2] - rb[0] + 1}x{rb[3] - rb[1] + 1} vs Office {ob[2] - ob[0] + 1}x{ob[3] - ob[1] + 1}")
        if max(abs(a - b) for a, b in zip(office["bar_color"], ours["bar_color"])) > 12:
            out.append(f"bar colour {ours['bar_color']} vs Office {office['bar_color']}")
        if "glyph" in ours and ours["glyph"][3] >= rb[1]:
            out.append("glyph reaches into the bar")
        # The glyph's height above the bar, within 2 device pixels.
        if "glyph" in ours and "glyph" in office:
            og, rg = office["glyph"], ours["glyph"]
            if abs((rb[1] - rg[1]) - (ob[1] - og[1])) > 2:
                out.append(f"glyph top {rb[1] - rg[1]} px above the bar vs Office {ob[1] - og[1]}")
    if [y for y in ours.get("dark_lines", []) if "bar" in ours and y < ours["bar"][1]]:
        out.append("a dark line above the bar (an outline box)")
    return out


def main() -> None:
    which = [a for a in sys.argv[1:] if a in ("calc", "write")] or ["calc", "write"]
    pairs = {"calc": ("excel", ["Schriftfarbe", "Füllfarbe"]), "write": ("word", ["Schriftfarbe", "Texthervorhebungsfarbe"])}
    out: dict = {"date": datetime.date.today().isoformat(), "apps": {}}
    total = ok = 0
    with real_input("check_colorbuttons"):
        for app in which:
            office_app, names = pairs[app]
            off, ours = office_side(office_app, names), raabe_side(app, names)
            out["apps"][app] = {}
            for n in names:
                d = compare(off[n], ours[n])
                out["apps"][app][n] = {"office": off[n], "ours": ours[n], "diff": d}
                total += 1
                ok += not d
                print(app, n, "ok" if not d else d)
    out["passed"], out["total"] = ok, total
    path = HERE / "results" / f"{out['date']}-colorbuttons.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(f"{ok}/{total} as in Office; crops in {TMP if os.environ.get('KEEP') else 'TEMP'}; wrote {path.name}")


if __name__ == "__main__":
    main()
