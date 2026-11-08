"""Windows, views and ink against Excel (D-093), both ways, without real input:
1. RaabeCalc (off screen, never focused, its own DevTools port) splits Tabelle1 at C5 (Teilen) and Tabelle2 at A1
   (the middle of the window), opens a second window (Neues Fenster) showing Tabelle2 at C3 with 150 %, switches
   Tabelle3 to the Seitenlayout view at 85 % without Lineal, adds a Benutzerdefinierte Ansicht with rows 3:4
   hidden, and draws ink on Tabelle1: one stroke through synthesized pointer events (DevTools
   Input.dispatchMouseEvent, the real mouse is not touched) with the red pen, one with the Textmarker. It saves;
   our own invisible Excel opens the file and must report the same through COM: windows and their captions, each
   window's sheet, cell and zoom, the split rows/columns, the view type and ruler, the custom view and the ink
   (msoInk shapes with their cells).
2. RaabeCalc opens Excel's own files (packages/ooxml/test/fixtures/excel-*.xlsx from probe_excel_windows.py and
   probe_excel_ink.py) and must show what Excel recorded: the split panes, both windows as document tabs with
   Excel's captions, the Seitenlayout view, the custom views and the ink drawings.
Result: results/<date>-calc-windows.json.
"""
from __future__ import annotations

import datetime
import json
import os
import tempfile
from pathlib import Path

from feellab import calc_catalogue, winapi
from feellab.calc_catalogue import Case, CalcSide
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
TMP = Path(tempfile.gettempdir())
FIX = ROOT / "packages" / "ooxml" / "test" / "fixtures"
REC = json.loads((FIX / "excel-windows.json").read_text(encoding="utf-8"))
INK = json.loads((FIX / "excel-ink.json").read_text(encoding="utf-8"))

os.environ["RAABE_FEEL_BACKGROUND"] = "1"
calc_catalogue.winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4000, 0, w, h)
PORT = 9396

DATA = {f"{c}{r}": str(r * 10 + i) for r in range(1, 21) for i, c in enumerate("ABCDEF")}

BUILD = r"""(async () => {
  const app = __calc.app, P = __calc.parse, cmd = (id) => app.commands.run(id);
  const sel = (a) => { const r = P(a); app.grid.selectCell(r.r0, r.c0); };
  const sheet = (i) => { app.activateSheet(i); };
  // Teilen at C5 on Tabelle1, at A1 on Tabelle2.
  sheet(0); sel("C5"); await cmd("view.split");
  sheet(1); sel("A1"); await cmd("view.split");
  // Seitenlayout on Tabelle3 at 85 % without Lineal.
  sheet(2); await cmd("view.pageLayout"); app.setZoom(85); await cmd("view.ruler");
  // Benutzerdefinierte Ansicht with rows 3:4 hidden on Tabelle1 (Ausgeblendete Zeilen, Druckeinstellungen).
  sheet(0);
  for (const r of [2, 3]) app.sheet.rows.set(r, { ...(app.sheet.rows.get(r) ?? {}), hidden: true });
  app.grid.refreshLayout();
  app.windows.addView("Ohne Zeilen", true, true);
  for (const r of [2, 3]) delete app.sheet.rows.get(r).hidden;
  app.grid.refreshLayout();
  return { split1: app.wb.sheets[0].view.split, split2: app.wb.sheets[1].view.split, top1: [app.wb.sheets[0].view.top, app.wb.sheets[0].view.left] };
})()"""

WINDOW2 = r"""(async () => {
  const app = __calc.app, P = __calc.parse, cmd = (id) => app.commands.run(id);
  await cmd("view.newWindow");
  app.activateSheet(1);
  const r = P("C3"); app.grid.selectCell(r.r0, r.c0);
  app.setZoom(150);
  const tabs = [...document.querySelectorAll(".doc-tab, [role=tab].doctab, .titlebar .tab")].map((t) => t.textContent);
  return { tabs: app.docList.flatMap((d) => app.windows.tabs(d)).map((t) => t.name), shown: app.wb.shownWindow ?? 0, domTabs: tabs };
})()"""

INKJS = r"""(async () => {
  const app = __calc.app;
  app.windows.select(app.windows.activeId(app.activeDoc).replace(/#\d+$/, "#0"));
  app.activateSheet(0);
  const a = app.grid.rectOnCanvas(__calc.parse("H3")), b = app.grid.rectOnCanvas(__calc.parse("J6"));
  const r = app.grid.el.getBoundingClientRect();
  app.ink.setTool({ t: "pen", i: 1 });
  return { x0: r.left + a.x + 5, y0: r.top + a.y + 5, x1: r.left + b.x + 20, y1: r.top + b.y + 10 };
})()"""


def mouse(c: CalcSide, kind: str, x: float, y: float, buttons: int) -> None:
    c.js_cdp("Input.dispatchMouseEvent", {"type": kind, "x": x, "y": y, "button": "left", "buttons": buttons, "clickCount": 1 if kind != "mouseMoved" else 0})


def cdp(self: CalcSide, method: str, params: dict):
    self._id += 1
    self.ws.send(json.dumps({"id": self._id, "method": method, "params": params}))
    while True:
        msg = json.loads(self.ws.recv())
        if msg.get("id") == self._id:
            return msg.get("result")


CalcSide.js_cdp = cdp  # type: ignore[attr-defined]


def main() -> None:
    ours_file = TMP / "raabe-windows-ours.xlsx"
    report: dict = {"raabe->excel": [], "excel->raabe": []}
    c = CalcSide(PORT)
    try:
        c.reset(Case("windows", DATA, [], "A1", sheets=3))
        built = c.js(BUILD)
        report["built"] = built
        report["window2"] = c.js(WINDOW2)
        pos = c.js(INKJS)
        # A stroke drawn with synthesized pointer events (DevTools): a zigzag from H3 towards J6.
        mouse(c, "mouseMoved", pos["x0"], pos["y0"], 0)
        mouse(c, "mousePressed", pos["x0"], pos["y0"], 1)
        for i in range(1, 25):
            t = i / 24
            mouse(c, "mouseMoved", pos["x0"] + (pos["x1"] - pos["x0"]) * t, pos["y0"] + (pos["y1"] - pos["y0"]) * t + (6 if i % 2 else -6), 1)
        mouse(c, "mouseReleased", pos["x1"], pos["y1"], 0)
        # The Textmarker through the test hook (points in CSS px of the grid).
        ink = c.js(r"""(() => { const app = __calc.app; const a = app.grid.rectOnCanvas(__calc.parse("B12"));
          app.ink.testStroke(4, Array.from({ length: 20 }, (_, i) => ({ x: a.x + i * 8, y: a.y + 8, t: i * 12 })));
          app.ink.setTool({ t: "select" });
          return app.sheet.inks.map((k) => ({ from: [k.from.col, k.from.row], strokes: k.strokes.length, pts: k.strokes[0].pts.length, color: k.strokes[0].brush.color, kind: k.strokes[0].brush.kind, w: k.size.w, h: k.size.h })); })()""")
        report["ink"] = ink
        c.js(f"__calc.save({json.dumps(str(ours_file))})")
    finally:
        c.close()

    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(ours_file))
        wins = [wb.Windows(i) for i in range(1, wb.Windows.Count + 1)]
        ex: dict = {"captions": [w.Caption for w in wins]}
        name = wb.Name
        w1 = x.app.Windows(f"{name}  -  1") if len(wins) > 1 else wins[0]
        w2 = x.app.Windows(f"{name}  -  2") if len(wins) > 1 else None
        ex["w1"] = {"sheet": w1.ActiveSheet.Name, "cell": w1.ActiveCell.Address.replace("$", ""), "zoom": w1.Zoom}
        if w2:
            ex["w2"] = {"sheet": w2.ActiveSheet.Name, "cell": w2.ActiveCell.Address.replace("$", ""), "zoom": w2.Zoom}
        w1.Activate()
        out = {}
        names = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
        ex["names"] = names
        for i, sname in zip(range(1, 4), ("Tabelle1", "Tabelle2", "Tabelle3")):
            wb.Worksheets(i).Activate()
            w = x.app.ActiveWindow
            out[sname] = {"split": bool(w.Split), "freeze": bool(w.FreezePanes), "splitRow": w.SplitRow, "splitColumn": w.SplitColumn, "panes": w.Panes.Count,
                          "paneScroll": [[w.Panes(k).ScrollRow, w.Panes(k).ScrollColumn] for k in range(1, w.Panes.Count + 1)], "view": w.View, "ruler": bool(w.DisplayRuler), "zoom": w.Zoom}
        ex["sheets"] = out
        ex["custom"] = [wb.CustomViews(i).Name for i in range(1, wb.CustomViews.Count + 1)]
        ws = wb.Worksheets(1)
        ex["shapes"] = [{"name": ws.Shapes(i).Name, "type": ws.Shapes(i).Type, "tl": ws.Shapes(i).TopLeftCell.Address.replace("$", ""), "w": ws.Shapes(i).Width, "h": ws.Shapes(i).Height}
                        for i in range(1, ws.Shapes.Count + 1)]
        if "Ohne Zeilen" in ex["custom"]:
            wb.CustomViews("Ohne Zeilen").Show()
            ex["custom_shows_hidden"] = [bool(ws.Rows(3).Hidden), bool(ws.Rows(4).Hidden), bool(ws.Rows(5).Hidden)]
        wb.Close(False)
    finally:
        x.close()
    report["excel"] = ex
    ink = report["ink"]
    checks = [
        ("two windows", len(ex["captions"]) == 2),
        ("captions as Excel's (Name  -  1/2)", sorted(ex["captions"]) == [f"{ours_file.name}  -  1", f"{ours_file.name}  -  2"]),
        ("window 2 shows the second sheet, C3 at 150 %", ex.get("w2") == {"sheet": ex["names"][1], "cell": "C3", "zoom": 150}),
        ("window 1 shows the first sheet", ex["w1"]["sheet"] == ex["names"][0]),
        # Excel at this PC's 150 % draws Calibri rows 14,5 pt high where the file says 15 pt: a quarter row is allowed.
        ("Tabelle1 split at C5: 4 rows, 2 columns, 4 panes", abs(ex["sheets"]["Tabelle1"]["splitRow"] - 4) < 0.25 and ex["sheets"]["Tabelle1"]["splitColumn"] == 2 and ex["sheets"]["Tabelle1"]["panes"] == 4),
        ("Tabelle1 split not frozen", ex["sheets"]["Tabelle1"]["split"] and not ex["sheets"]["Tabelle1"]["freeze"]),
        ("Tabelle1 bottom-right pane from C5", ex["sheets"]["Tabelle1"]["paneScroll"][-1] == [5, 3]),
        ("Tabelle2 split at A1: 4 panes", ex["sheets"]["Tabelle2"]["panes"] == 4 and ex["sheets"]["Tabelle2"]["split"]),
        ("Tabelle3 Seitenlayout (View 3)", ex["sheets"]["Tabelle3"]["view"] == 3),
        ("Tabelle3 without Lineal", ex["sheets"]["Tabelle3"]["ruler"] is False),
        ("Tabelle3 zoom 85", ex["sheets"]["Tabelle3"]["zoom"] == 85),
        ("custom view", ex["custom"] == ["Ohne Zeilen"]),
        ("custom view hides rows 3:4", ex.get("custom_shows_hidden") == [True, True, False]),
        ("two ink drawings", len(ex["shapes"]) == 2),
        ("both are msoInk (23)", all(s["type"] == 23 for s in ex["shapes"])),
        ("pen stroke at H3", ex["shapes"][0]["tl"] == "H3" if ex["shapes"] else False),
        ("Textmarker at B12", ex["shapes"][1]["tl"] == "B12" if len(ex["shapes"]) > 1 else False),
        ("our pen stroke has the synthesized points", ink[0]["pts"] >= 20 and ink[0]["color"] == "#E71224"),
    ]
    # Ink size: Excel's points against ours (px at 100 % = 0.75 pt).
    for i, s in enumerate(ex["shapes"][:2]):
        checks.append((f"ink {i + 1} width within 1.5 pt", abs(s["w"] - ink[i]["w"] * 0.75) <= 1.5))
    for name, ok in checks:
        report["raabe->excel"].append({"check": name, "ok": bool(ok)})
        print("R->E", name, ok, flush=True)

    # 2. Excel's files in RaabeCalc.
    c = CalcSide(PORT)
    try:
        c.js(f"__calc.open({json.dumps(str(FIX / 'excel-split.xlsx'))})")
        got = c.js("""(() => __calc.app.wb.sheets.map((s) => ({ split: s.view.split ? [s.view.split.x, s.view.split.y, s.view.split.top, s.view.split.left] : null, br: [s.view.top, s.view.left] })))()""")
        for i, case in enumerate(REC["splits"]):
            p = case["com"]["panes"]
            want_tl = [p[0]["scrollRow"] - 1, p[0]["scrollColumn"] - 1]
            want_br = [p[-1]["scrollRow"] - 1, p[-1]["scrollColumn"] - 1]
            g = got[i]
            ok = bool(g["split"]) and g["split"][2:] == want_tl and g["br"] == want_br
            report["excel->raabe"].append({"check": f"split {case['name']}", "ok": ok, **({} if ok else {"got": g, "want": [want_tl, want_br]})})
        # The panes drawn: the split sheet shows four panes with the bars where Excel has them.
        bars = c.js("""(async () => { const app = __calc.app; app.activateSheet(2); app.grid.refreshLayout(); await new Promise((r) => setTimeout(r, 400)); const g = app.grid.geometry();
          const s = app.sheet.view.split, rows = (px) => app.grid.hitCell(app.grid.geometry().headerW + 5, px - 2).row;
          return { above: s ? rows(Number(app.grid.el.querySelector('.split-bar.h')?.style.top.replace('px', '') ?? 0)) : null, panes: app.grid.el.querySelectorAll('.split-bar:not([hidden])').length, x: app.sheet.view.split.x / 15, y: app.sheet.view.split.y / 15, hw: g.headerW, hh: g.headerH }; })()""")
        report["excel->raabe"].append({"check": "both split bars shown (both3x5)", "ok": bars["panes"] == 2, "bars": bars})
        report["excel->raabe"].append({"check": "the horizontal bar below row 5 as in Excel (both3x5)", "ok": bars.get("above") == 4, "bars": bars})
        c.js(f"__calc.open({json.dumps(str(FIX / 'excel-windows.xlsx'))})")
        w = c.js("""(() => { const app = __calc.app; return { tabs: app.docList.flatMap((d) => app.windows.tabs(d)).map((t) => t.name).filter((n) => n.startsWith('excel-windows')),
          w2: app.windows.viewOf(app.wb, 1, app.wb.sheets[1]).zoom, a2: app.windows.activeOf(app.wb, 1) }; })()""")
        report["excel->raabe"].append({"check": "two windows as document tabs with Excel's captions", "ok": w["tabs"] == ["excel-windows  -  1", "excel-windows  -  2"], "got": w["tabs"]})
        report["excel->raabe"].append({"check": "window 2 shows Tabelle2 at 150 %", "ok": w["w2"] == 150 and w["a2"] == 1})
        sw = c.js("""(() => { const app = __calc.app; const id = app.windows.activeId(app.activeDoc).replace(/#\\d+$/, '#1'); app.windows.select(id);
          return { sheet: app.sheet.name, zoom: app.sheet.view.zoom, cell: [app.sheet.view.active.row, app.sheet.view.active.col], title: document.title }; })()""")
        report["excel->raabe"].append({"check": "switching to window 2 shows its sheet, cell and zoom", "ok": sw["sheet"] == "Tabelle2" and sw["zoom"] == 150 and sw["cell"] == [2, 2], "got": sw})
        c.js(f"__calc.open({json.dumps(str(FIX / 'excel-pagelayout.xlsx'))})")
        pl = c.js("""(async () => { const app = __calc.app; app.grid.refreshLayout(); await new Promise((r) => setTimeout(r, 400)); const v = app.sheet.view; const p = app.pageLayout.pages; return { layout: !!v.layout, zoom: v.zoom, normal: v.normalZoom, pages: p ? [p.paperY.length, p.paperX.length] : null,
          overlay: !!app.grid.el.querySelector('.pl-overlay'), checked: app.commands.isChecked('view.pageLayout') }; })()""")
        report["excel->raabe"].append({"check": "Seitenlayout view at 85 % (Normal 100 %)", "ok": pl["layout"] and pl["zoom"] == 85 and pl["normal"] == 100, "got": pl})
        report["excel->raabe"].append({"check": "pages drawn", "ok": pl["overlay"] and pl["checked"]})
        c.js(f"__calc.open({json.dumps(str(FIX / 'excel-customviews.xlsx'))})")
        cv = c.js("""(() => { const app = __calc.app; const names = app.wb.customViews.map((c) => c.name);
          const g = app.wb.customViews.find((c) => c.name === 'Ohne Zeilen').guid; app.windows.showView(g);
          const s = app.sheet; const r = [!!s.rows.get(2)?.hidden, !!s.rows.get(3)?.hidden, s.view.zoom, s.view.active.row, s.view.active.col, s.page?.orientation ?? null];
          const g2 = app.wb.customViews.find((c) => c.name === 'Alles').guid; app.windows.showView(g2);
          return { names, ohne: r, alles: [!!s.rows.get(2)?.hidden, s.view.zoom, s.view.active.row, s.view.active.col] }; })()""")
        want_ohne = REC["custom"]["after_show_ohne"]
        report["excel->raabe"].append({"check": "Excel's three custom views listed", "ok": sorted(cv["names"]) == sorted(REC["custom"]["names"])})
        report["excel->raabe"].append({"check": "Anzeigen 'Ohne Zeilen' as Excel (rows 3:4 hidden, 120 %, D5, landscape)", "ok": cv["ohne"] == [True, True, want_ohne["zoom"], 4, 3, "landscape"], "got": cv["ohne"]})
        want_alles = REC["custom"]["after_show_alles"]
        report["excel->raabe"].append({"check": "Anzeigen 'Alles' as Excel (rows shown, 100 %, B2)", "ok": cv["alles"] == [False, want_alles["zoom"], 1, 1], "got": cv["alles"]})
        c.js(f"__calc.open({json.dumps(str(FIX / 'excel-ink.xlsx'))})")
        ik = c.js("""(() => { const app = __calc.app; app.grid.refreshLayout(); return { n: app.sheet.inks.length, svgs: app.grid.el.querySelectorAll('.ink-layer svg').length,
          cells: app.sheet.inks.map((k) => [k.from.col, k.from.row]), pics: (app.sheet.pictures ?? []).length }; })()""")
        com = INK["com"]
        want_cells = [[ord(s["tl"][0]) - 65, int(s["tl"][1:]) - 1] for s in com]
        report["excel->raabe"].append({"check": "Excel's four ink drawings read as ink (not pictures)", "ok": ik["n"] == 4 and ik["pics"] == 0})
        report["excel->raabe"].append({"check": "ink at Excel's cells (COM TopLeftCell)", "ok": ik["cells"] == want_cells, "got": ik["cells"], "want": want_cells})
        # Saved again by RaabeCalc unchanged, Excel reads the ink back.
        back = TMP / "raabe-ink-roundtrip.xlsx"
        c.js(f"__calc.save({json.dumps(str(back))})")
    finally:
        c.close()
    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(back))
        ws = wb.Worksheets(1)
        types = [ws.Shapes(i).Type for i in range(1, ws.Shapes.Count + 1)]
        wb.Close(False)
    finally:
        x.close()
    report["excel->raabe"].append({"check": "RaabeCalc's save of Excel's ink opens in Excel as four msoInk", "ok": types == [23, 23, 23, 23], "got": types})
    for r in report["excel->raabe"]:
        print("E->R", r["check"], r["ok"], flush=True)
    (HERE / "results" / f"{datetime.date.today().isoformat()}-calc-windows.json").write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    for key in ("raabe->excel", "excel->raabe"):
        v = report[key]
        print(key, sum(1 for r in v if r["ok"]), "/", len(v))


if __name__ == "__main__":
    main()
