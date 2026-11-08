"""Pictures with Bildformat against Excel (D-087), both ways, without real input:

1. Excel -> RaabeCalc: RaabeCalc (off screen) opens the workbook Excel saved in probe_excel_pictures.py
   (packages/ooxml/test/fixtures/xlsx-pictures.xlsx, one picture per setting) and must read every setting
   (excel-pictures.json); each picture's block of cells is captured through DevTools and compared with
   Excel's own rendering (TEMP/raabe-pictures/excel-<case>.png, Range.CopyPicture): mean absolute
   difference of the greyscale images, next to the difference Excel's picture has from the plain one
   (what RaabeCalc showed before D-087, when it drew every picture plainly).
2. RaabeCalc -> Excel: in a new workbook RaabeCalc inserts the same test picture once per setting and gives
   it the setting through its picture commands (Bildformat's API, each change one undo step), saves; our own
   invisible Excel opens the file and reads every picture back through COM (the same properties the probe
   recorded for Excel's own pictures) and draws it (CopyPicture), compared with Excel's own picture.

    .venv/Scripts/python.exe check_calc_pictures.py

Result: results/<date>-calc-pictures.json. RaabeCalc's window is placed off screen; the focus goes back to
the window that had it."""
from __future__ import annotations

import base64
import datetime
import io
import json
import sys
import tempfile
import time
from pathlib import Path

from PIL import Image, ImageChops, ImageGrab, ImageOps

from feellab import winapi
from feellab.excel import ExcelTarget
from probe_excel_pictures import BH, BW, PNG, make_png, readback

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
TMP = Path(tempfile.gettempdir())
FIX = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "xlsx-pictures.xlsx"
REC = json.loads((ROOT / "packages" / "ooxml" / "test" / "fixtures" / "excel-pictures.json").read_text(encoding="utf-8"))
SHOTS = TMP / "raabe-pictures"
OURS = TMP / "raabe-pictures-ours.xlsx"
SKIP = {"left", "top", "name", "zorder"}
# Readback values that depend on how the picture was made, not on its setting (COM's AddPicture leaves
# Seitenverhältnis sperren off; RaabeCalc's inserted pictures have it on, as Excel's Einfügen > Bilder).
LOOSE = {"lockAspect"}


def calc_side():
    """RaabeCalc off screen (CalcSide places its window; here it goes left of the screen)."""
    from feellab import calc_catalogue
    fg = winapi.foreground()
    real = winapi.place
    winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4200, 0, 2000, 1400)
    try:
        c = calc_catalogue.CalcSide()
    finally:
        winapi.place = real
    if fg and winapi.foreground() != fg:
        try:
            winapi.activate(fg, 0)
        except Exception:  # noqa: BLE001
            pass
    return c


def capture(c, a: str, b: str) -> Image.Image:
    r0 = c.js(f"__calc.point({json.dumps(a)}, 'tl')")
    r1 = c.js(f"__calc.point({json.dumps(b)}, 'br')")
    c.ws.send(json.dumps({"id": 99001, "method": "Page.captureScreenshot",
                          "params": {"format": "png", "clip": {"x": r0["x"], "y": r0["y"], "width": r1["x"] - r0["x"], "height": r1["y"] - r0["y"], "scale": 1}}}))
    while True:
        msg = json.loads(c.ws.recv())
        if msg.get("id") == 99001:
            break
    return Image.open(io.BytesIO(base64.b64decode(msg["result"]["data"]))).convert("RGB")


def col_name(c: int) -> str:
    s = ""
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def block_cells(i: int) -> tuple[str, str, int]:
    per_row = 5
    r = 1 + (i // per_row) * BH
    c = 1 + (i % per_row) * BW
    return f"{col_name(c)}{r}", f"{col_name(c + BW - 1)}{r + BH - 1}", r - 1


def diff(a: Image.Image, b: Image.Image) -> float:
    g1 = ImageOps.grayscale(a)
    g2 = ImageOps.grayscale(b.resize(a.size))
    h = ImageChops.difference(g1, g2).histogram()
    return round(sum(i * n for i, n in enumerate(h)) / (a.width * a.height), 2)


def excel_shots(path: Path, names: list[str], prefix: str) -> tuple[dict, dict]:
    """Our Excel opens a workbook: every picture read back through COM and every block drawn."""
    x = ExcelTarget(visible=False)
    back, shots = {}, {}
    try:
        wb = x.app.Workbooks.Open(str(path))
        ws = wb.Worksheets(1)
        for s in ws.Shapes:
            back[s.Name] = readback(s)
        for i, n in enumerate(names):
            a, b, _ = block_cells(i)
            img = None
            for _ in range(8):
                try:
                    ws.Range(f"{a}:{b}").CopyPicture(1, 2)
                    time.sleep(0.3)
                    img = ImageGrab.grabclipboard()
                    if isinstance(img, Image.Image):
                        break
                except Exception:  # noqa: BLE001
                    time.sleep(0.4)
            if isinstance(img, Image.Image):
                img = img.convert("RGB")
                img.save(SHOTS / f"{prefix}-{n}.png")
                shots[n] = img
        wb.Close(False)
    finally:
        x.close()
    return back, shots


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    make_png(PNG)
    SHOTS.mkdir(exist_ok=True)
    cases = [c for c in REC["cases"] if c["case"] != "group"]
    names = [c["case"] for c in REC["cases"]]
    report: dict = {"excel->raabe": {}, "raabe->excel": {}}
    c = calc_side()
    try:
        # ----- 1. Excel's file in RaabeCalc ----------------------------------------------------
        c.js(f"__calc.open({json.dumps(str(FIX))})")
        time.sleep(1.0)
        read = c.js("""(() => (__calc.app.sheet.pictures ?? []).map((p) => ({ name: p.name, from: p.from, to: p.to ?? null, size: p.size ?? null,
            format: p.format ?? null, alt: p.alt ?? null, title: p.title ?? null, hidden: !!p.hidden, baked: p.baked ?? null, group: !!p.group })))()""")
        by = {p["name"]: p for p in read}
        expect = {
            "crop": lambda f: f and f.get("crop", {}).get("l") == 0.05556, "lum": lambda f: f and f.get("lum") == {"bright": 0.4, "contrast": -0.2},
            "grayscale": lambda f: f and f.get("recolor") == {"kind": "grayscale"}, "bw": lambda f: f and f.get("recolor", {}).get("kind") == "bw",
            "washout": lambda f: f and f.get("recolor") == {"kind": "washout"}, "transparentColor": lambda f: f and f.get("transparentColor") == "FFFFFF",
            "brightnessContrast": lambda f: f and f.get("brightness") == 0.2 and f.get("contrast") == -0.2, "saturation": lambda f: f and f.get("saturation") == 2,
            "temperature": lambda f: f and f.get("temperature") == 8800, "sharpen": lambda f: f and f.get("sharpen") == 0.5, "soften": lambda f: f and f.get("sharpen") == -0.5,
            "rotate30": lambda f: f and f.get("rotation") == 30, "flipH": lambda f: f and f.get("flipH") is True, "flipV": lambda f: f and f.get("flipV") is True,
            "line": lambda f: f and f.get("line", {}).get("color") == "C00000" and f["line"].get("width") == 4.5, "shadowOuter": lambda f: f and f.get("shadow", {}).get("kind") == "outer",
            "shadowInner": lambda f: f and f.get("shadow", {}).get("kind") == "inner", "glow": lambda f: f and f.get("glow", {}).get("radius") == 8,
            "softEdge": lambda f: f and f.get("softEdge") == 6, "reflection": lambda f: f and f.get("reflection") is not None, "bevel": lambda f: f and f.get("bevel") is not None,
            "rotation3d": lambda f: f and "lat=" in (f.get("scene3dXml") or ""), "oval": lambda f: f and f.get("geometry") == "ellipse",
            "roundRect": lambda f: f and f.get("geometry") == "roundRect",
        }
        e2r = report["excel->raabe"]
        e2r["read"] = {}
        for cs in cases:
            p = by.get(f"Bild {cs['case']}")
            if cs["case"] in expect:
                ok = bool(p) and bool(expect[cs["case"]](p["format"]))
            elif cs["case"] == "altText":
                ok = bool(p) and p["alt"] == "Ein Testbild mit vier Farben" and p["title"] == "Testbild"
            elif cs["case"] == "hidden":
                ok = bool(p) and p["hidden"]
            else:
                ok = bool(p)
            e2r["read"][cs["case"]] = ok
        e2r["read"]["group"] = bool(by.get("Bild group", {}).get("group"))
        # Rendering of each block, against Excel's own.
        c.js("(() => { const v = __calc.app.sheet.view; v.zoom = 100; v.left = 0; v.top = 0; __calc.app.grid.contentChanged(); })()")
        e2r["look"] = {}
        plain_excel = Image.open(SHOTS / "excel-plain.png").convert("RGB") if (SHOTS / "excel-plain.png").exists() else None
        for i, n in enumerate(names):
            a, b, top = block_cells(i)
            left = (i % 5) * BW
            c.js(f"(() => {{ const v = __calc.app.sheet.view; v.top = {top}; v.left = {left}; __calc.app.grid.contentChanged(); }})()")
            time.sleep(0.5)
            img = capture(c, a, b)
            img.save(SHOTS / f"calc-{n}.png")
            ex = SHOTS / f"excel-{n}.png"
            if ex.exists():
                e = Image.open(ex).convert("RGB")
                e2r["look"][n] = {"ours": diff(e, img), "plainDrawing": diff(e, plain_excel) if plain_excel else None}
        print("excel->raabe read", sum(e2r["read"].values()), "/", len(e2r["read"]), flush=True)

        # ----- 2. RaabeCalc's pictures in Excel ----------------------------------------------------
        png = "data:image/png;base64," + base64.b64encode(PNG.read_bytes()).decode()
        c.js("__calc.reset({}, 'A1', 1)")
        for cs, p in [(cs, by.get(f"Bild {cs['case']}")) for cs in cases]:
            if not p:
                continue
            spec = {"from": p["from"], "to": p["to"], "size": p["size"], "format": p["format"], "alt": p["alt"], "title": p["title"], "hidden": p["hidden"], "name": p["name"]}
            c.js(f"""(async () => {{ const app = __calc.app, pics = __calc.pictures, s = {json.dumps(spec)};
              const img = await new Promise((res) => {{ const i = new Image(); i.onload = () => res({{ src: i.src, w: i.naturalWidth, h: i.naturalHeight }}); i.src = {json.dumps(png)}; }});
              pics.insertImage(img);
              // The setting through the picture commands: as the tab and the pane give it (one undo step each).
              pics.edit("Test", (q) => ({{ ...q, name: s.name, from: s.from, to: s.to ?? undefined, size: s.size ?? q.size, lockAspect: undefined }}));
              if (s.format) pics.setFormat("Test", s.format);
              if (s.alt || s.title) pics.edit("Alternativtext", (q) => ({{ ...q, alt: s.alt ?? undefined, title: s.title ?? undefined }}));
              if (s.hidden) pics.edit("Ausblenden", (q) => ({{ ...q, hidden: true }}));
              pics.deselect(); }})()""")
        # Commands of the tab on one more picture: Rechtsdrehung 90 Grad and Horizontal spiegeln.
        c.js(f"""(async () => {{ const pics = __calc.pictures;
          const img = await new Promise((res) => {{ const i = new Image(); i.onload = () => res({{ src: i.src, w: i.naturalWidth, h: i.naturalHeight }}); i.src = {json.dumps(png)}; }});
          __calc.app.grid.selectCell(80, 2); pics.insertImage(img); pics.edit("Test", (q) => ({{ ...q, name: "Bild commands" }}));
          await __calc.app.commands.run("pic.rotateRight"); await __calc.app.commands.run("pic.flipH"); }})()""")
        steps = c.js("__calc.app.history.undoStack?.length ?? null")
        # Rückgängig takes back the last change (Horizontal spiegeln), Wiederholen does it again.
        flip = "(() => { const p = (__calc.app.sheet.pictures ?? []).find((q) => q.name === 'Bild commands'); return !!(p && p.format && p.format.flipH); })()"
        c.js("__calc.app.commands.run('edit.undo')")
        undone = c.js(flip)
        c.js("__calc.app.commands.run('edit.redo')")
        redone = c.js(flip)
        c.js(f"__calc.save({json.dumps(str(OURS))})")
    finally:
        c.close()
    back, shots = excel_shots(OURS, names[:-1], "ours")
    r2e = report["raabe->excel"]
    r2e["undo"] = {"steps": steps, "undoneFlip": not undone, "redoneFlip": redone}
    r2e["readback"] = {}
    for cs in cases:
        mine = back.get(f"Bild {cs['case']}")
        want = cs["com"]
        if not mine:
            r2e["readback"][cs["case"]] = {"ok": False, "missing": True}
            continue
        bad = {k: [want.get(k), mine.get(k)] for k in want if k not in SKIP | LOOSE and _differs(want.get(k), mine.get(k))}
        r2e["readback"][cs["case"]] = {"ok": not bad, **({"differ": bad} if bad else {})}
    cmd = back.get("Bild commands")
    r2e["commands"] = {"rotation": cmd and cmd.get("rotation"), "flipH": cmd and cmd.get("flipH")}
    r2e["look"] = {}
    for n, img in shots.items():
        ex = SHOTS / f"excel-{n}.png"
        if ex.exists():
            r2e["look"][n] = diff(Image.open(ex).convert("RGB"), img)
    ok1 = sum(report["excel->raabe"]["read"].values())
    ok2 = sum(1 for v in r2e["readback"].values() if v.get("ok"))
    report["summary"] = {"excel->raabe read": f"{ok1}/{len(report['excel->raabe']['read'])}", "raabe->excel readback": f"{ok2}/{len(r2e['readback'])}",
                         "commands": r2e["commands"], "undo": r2e["undo"]}
    out = HERE / "results" / f"{datetime.date.today():%Y-%m-%d}-calc-pictures.json"
    out.write_text(json.dumps(report, ensure_ascii=False, indent=1) + "\n", encoding="utf-8", newline="\n")
    print(json.dumps(report["summary"], ensure_ascii=False))
    print("written", out.name)


def _differs(a, b) -> bool:
    if isinstance(a, float) and isinstance(b, (int, float)):
        return abs(a - b) > 0.6
    return a != b


if __name__ == "__main__":
    main()
