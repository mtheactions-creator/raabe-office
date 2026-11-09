"""Does a Format pane setting change RaabeCalc's chart where it changes Excel's? (D-113)

For a list of settings (one model change each) RaabeCalc (off screen, DevTools) draws a column chart with and
without the setting and saves the file; our own invisible Excel opens that file and exports the same charts
as pictures. For each setting the script finds the region of the chart that changed (pixels that differ from
the chart without the setting) on both sides, in coordinates relative to the chart's size, and compares the
two regions (overlap of the boxes, IoU). Excel draws from our file, so this also shows Excel reads the
setting.

    .venv/Scripts/python.exe compare_calc_chartpane_pictures.py

Result: results/2026-11-09-calc-chartpane-pictures.json.
"""
from __future__ import annotations

import base64
import datetime
import io
import json
import tempfile
import time
from pathlib import Path

from PIL import Image, ImageChops

from check_calc_chartpane import DATA, calc_side
from feellab.calc_catalogue import Case
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir()) / "raabe-chartpane-pics"
GRAD = {"kind": "linear", "angle": 90, "stops": [{"pos": 0, "color": "#FFFFFF"}, {"pos": 100, "color": "#4472C4"}]}
SOLID = {"type": "solid", "color": "#C00000", "width": 2}
SETTINGS: list[tuple[str, dict]] = [
    ("Wertachse: Einfarbige Linie", {"format": {"valAxis": {"line": SOLID}}}),
    ("Wertachse: Haupttyp Außen (mit Linie)", {"format": {"valAxis": {"line": SOLID, "majorTick": "out"}}}),
    ("Wertachse: Maximum 100, Hauptstriche 25", {"format": {"valAxis": {"max": 100, "majorUnit": 25}}}),
    ("Wertachse: Logarithmische Skalierung", {"format": {"valAxis": {"logBase": 10}}}),
    ("Wertachse: Werte in umgekehrter Reihenfolge", {"format": {"valAxis": {"reverse": True}}}),
    ("Wertachse: Horizontale Achse schneidet bei 30", {"format": {"valAxis": {"crosses": 30}}}),
    ("Wertachse: Anzeigeeinheiten Hunderte", {"format": {"valAxis": {"dispUnits": "hundreds"}}}),
    ("Wertachse: Beschriftungsposition Hoch", {"format": {"valAxis": {"labelPos": "high"}}}),
    ("Kategorieachse: Beschriftungsposition Niedrig", {"format": {"catAxis": {"labelPos": "low"}}}),
    ("Kategorieachse: Bei größter Rubrik", {"format": {"catAxis": {"crosses": "max"}}}),
    ("Kategorieachse: Abstand zur Achse 500", {"format": {"catAxis": {"lblOffset": 500}}}),
    ("Kategorieachse: Kategorien umgekehrt", {"format": {"catAxis": {"reverse": True}}}),
    ("Zeichnungsfläche: Farbverlauf", {"format": {"plot": {"fill": {"type": "gradient", "gradient": GRAD}}}}),
    ("Diagrammbereich: Musterfüllung", {"format": {"chart": {"fill": {"type": "pattern", "pattern": {"preset": "dkUpDiag", "fg": "#A5A5A5", "bg": "#FFFFFF"}}}}}),
    ("Diagrammbereich: Rahmen doppelt 4 Pt.", {"format": {"chart": {"line": {"type": "solid", "color": "#000000", "width": 4, "compound": "dbl"}}}}),
    ("Hauptgitternetz: Strich", {"format": {"majorGrid": {"line": {"type": "solid", "color": "#7F7F7F", "width": 1, "dash": "dash"}}}}),
    ("Legende: Oben rechts", {"legend": "tr"}),
    ("Titel: 20 Pt. fett", {"format": {"title": {"text": {"size": 20, "bold": True}}}}),
    ("Reihe A: Schatten", {"series0": {"effects": {"shadow": {"color": "#000000", "transparency": 40, "blur": 4, "angle": 45, "distance": 6}}}}),
    ("Reihe A: Leuchten 8 Pt.", {"series0": {"effects": {"glow": {"color": "#FFC000", "size": 8, "transparency": 40}}}}),
    ("Reihe B: Musterfüllung", {"series1": {"fill": {"type": "pattern", "pattern": {"preset": "pct50", "fg": "#ED7D31", "bg": "#FFFFFF"}}}}),
    ("Reihe A: Punkt 3 eigene Farbe", {"series0": {"points": [{"idx": 2, "fill": {"type": "solid", "color": "#00B050"}}]}}),
    ("Reihen: Abstandsbreite 50", {"gapWidth": 50}),
    ("Reihe B auf der Sekundärachse", {"series1": {"secondary": True}}),
    ("Datenbeschriftungen", {"dataLabels": True}),
]


def apply_js(patch: dict) -> str:
    return f"""(m) => {{ const p = {json.dumps(patch)};
      for (const [k, v] of Object.entries(p)) {{
        if (k === 'format') m.format = v; else if (k.startsWith('series')) Object.assign(m.series[Number(k.slice(6))], v); else m[k] = v;
      }} delete m.raw; }}"""


def changed_box(a: Image.Image, b: Image.Image) -> tuple[float, float, float, float] | None:
    # A quarter of the size: anti-aliasing shifted by a fraction of a pixel is not a change.
    size = (max(1, a.width // 4), max(1, a.height // 4))
    a, b = a.convert("RGB").resize(size, Image.BOX), b.convert("RGB").resize(size, Image.BOX)
    diff = ImageChops.difference(a, b).convert("L").point(lambda v: 255 if v > 40 else 0)
    box = diff.getbbox()
    if not box:
        return None
    w, h = a.size
    return box[0] / w, box[1] / h, box[2] / w, box[3] / h


def iou(p, q) -> float:
    if not p or not q:
        return 0.0
    x0, y0, x1, y1 = max(p[0], q[0]), max(p[1], q[1]), min(p[2], q[2]), min(p[3], q[3])
    inter = max(0, x1 - x0) * max(0, y1 - y0)
    union = (p[2] - p[0]) * (p[3] - p[1]) + (q[2] - q[0]) * (q[3] - q[1]) - inter
    return inter / union if union > 0 else 0.0


def shot(c, k: int) -> Image.Image:
    r = c.js(f"(() => {{ const e = document.querySelectorAll('.chart-obj svg')[{k}]; e.scrollIntoView({{block: 'center'}}); const b = e.getBoundingClientRect(); return {{x: b.left, y: b.top, w: b.width, h: b.height}}; }})()")
    c._id += 1
    c.ws.send(json.dumps({"id": c._id, "method": "Page.captureScreenshot", "params": {"format": "png", "clip": {"x": r["x"], "y": r["y"], "width": r["w"], "height": r["h"], "scale": 2}}}))
    while True:
        msg = json.loads(c.ws.recv())
        if msg.get("id") == c._id:
            return Image.open(io.BytesIO(base64.b64decode(msg["result"]["data"])))


def main() -> None:
    TMP.mkdir(parents=True, exist_ok=True)
    path = TMP / "pictures.xlsx"
    c = calc_side(9431)
    ours: list[Image.Image] = []
    try:
        # Another test instance's AutoRecover file may be offered: discarded (Verwerfen).
        c.js("[...document.querySelectorAll('button')].find((b) => b.textContent.trim() === 'Verwerfen')?.click()")
        c.reset(Case("chartpane-pics", DATA, [], "A1"))
        for k in range(len(SETTINGS) + 1):
            c.js(f"__calc.app.grid.select({{r0:0,c0:0,r1:4,c1:2}}); __calc.app.charts.insert('column')")
            time.sleep(0.1)
            col, row = (k % 3) * 9 + 4, (k // 3) * 16
            c.js(f"(() => {{ const m = __calc.app.sheet.charts[{k}]; m.from = {{col:{col},row:{row},x:0,y:0}}; m.to = {{col:{col + 8},row:{row + 15},x:0,y:0}}; m.name = 'Diagramm {k + 1}'; }})()")
            if k:
                c.js(f"__calc.app.charts.changeSelected({apply_js(SETTINGS[k - 1][1])})")
        c.js("__calc.app.charts.refresh()")
        time.sleep(0.5)
        c.js("[...document.querySelectorAll('button')].find((b) => b.textContent.trim() === 'Verwerfen')?.click(); __calc.app.charts.select?.(null)")
        time.sleep(0.3)
        for k in range(len(SETTINGS) + 1):
            ours.append(shot(c, k))
            ours[-1].save(TMP / f"o{k:02d}.png")
        c.js(f"__calc.save({json.dumps(str(path))})")
    finally:
        c.close()
    x = ExcelTarget(visible=False)
    excel: list[Image.Image] = []
    try:
        wb = x.app.Workbooks.Open(str(path))
        ws = wb.Worksheets(1)
        for k in range(1, len(SETTINGS) + 2):
            f = TMP / f"x{k:02d}.png"
            img = None
            for _ in range(4):
                # Excel's export now and then leaves an empty file: once more after a moment.
                # Excel exports only charts in its window's view: each one goes to the top left first.
                co = ws.ChartObjects(k)
                co.Top, co.Left = 0, 0
                co.Chart.Export(str(f), "PNG")
                try:
                    img = Image.open(f).copy()
                    break
                except Exception:  # noqa: BLE001
                    time.sleep(0.5)
            excel.append(img)
        wb.Close(False)
    finally:
        x.close()
    rows = []
    for k, (name, patch) in enumerate(SETTINGS, start=1):
        bo = changed_box(ours[0], ours[k])
        bx = changed_box(excel[0], excel[k]) if excel[0] is not None and excel[k] is not None else None
        r = {"setting": name, "patch": patch, "ours_box": bo, "excel_box": bx, "iou": round(iou(bo, bx), 2)}
        r["ok"] = bool(bo) and bool(bx) and r["iou"] >= 0.3
        rows.append(r)
        print("ok " if r["ok"] else "BAD", f"{r['iou']:.2f}", name, flush=True)
    out = HERE / "results" / "2026-11-09-calc-chartpane-pictures.json"
    out.write_text(json.dumps({"date": datetime.date.today().isoformat(), "passed": sum(1 for r in rows if r["ok"]), "total": len(rows), "rows": rows},
                              ensure_ascii=False, indent=1, default=lambda v: [round(t, 3) for t in v]), encoding="utf-8", newline="\n")
    print("passed", sum(1 for r in rows if r["ok"]), "/", len(rows))


if __name__ == "__main__":
    main()
