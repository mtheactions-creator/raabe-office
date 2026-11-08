"""Datenschnitte against Excel (D-092), both ways, without real input; RaabeCalc runs off screen and never takes the focus.

1. RaabeCalc -> Excel: RaabeCalc makes a table with two slicers (Fach: Mathe and Deutsch chosen with a click and a
   Strg+click; Klasse: two columns, style Dunkel 2, Elemente ohne Daten kennzeichnen in place, absteigend) and a
   PivotTable with a slicer (only Englisch), saves; our own invisible Excel opens the file: the slicers, their
   settings and chosen items, the table's hidden rows and the PivotTable's report must be RaabeCalc's.
2. Excel -> RaabeCalc: RaabeCalc opens Excel's file (fixture excel-slicers.xlsx): every slicer's items in Excel's
   order with Excel's selected / no-data state, and the hidden rows; a click in RaabeCalc (Standard: Mathe), saved,
   opened in Excel: Excel shows that filter.
3. The look: the item buttons of Excel's rendering (TEMP/raabe-slicers/excel-cleared-*.png, the sheet Zwei's last state; from
   probe_excel_slicers.py) against RaabeCalc's buttons: place, size and colour of each state.
Result: results/<date>-calc-slicers.json.
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
SHOTS = TMP / "raabe-slicers"
PORT = 9474
os.environ["RAABE_FEEL_BACKGROUND"] = "1"
calc_catalogue.winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4000, 0, w, h)

DATA = {"A1": "Name", "B1": "Klasse", "C1": "Fach", "D1": "Note",
        "A2": "Anna", "B2": "8a", "C2": "Mathe", "D2": "2", "A3": "Ben", "B3": "8a", "C3": "Mathe", "D3": "3",
        "A4": "Cem", "B4": "8b", "C4": "Mathe", "D4": "1", "A5": "Dora", "B5": "8b", "C5": "Deutsch", "D5": "4",
        "A6": "Anna", "B6": "8a", "C6": "Deutsch", "D6": "1", "A7": "Ben", "B7": "8a", "C7": "Englisch", "D7": "2",
        "A8": "Cem", "B8": "8b", "C8": "Deutsch", "D8": "3", "A9": "Emil", "B9": "8c", "C9": "Englisch", "D9": "5"}

BUILD = r"""(async () => {
  const app = __calc.app, sl = app.slicers, P = __calc.parse;
  const t = app.tables.create(P("A1:D9"), true); t.name = "Noten";
  sl.add([{ sourceName: "Fach", table: { name: "Noten", column: 2 } }, { sourceName: "Klasse", table: { name: "Noten", column: 1 } }]);
  const fach = () => sl.byName("Fach"), klasse = () => sl.byName("Klasse");
  sl.click(fach(), "Mathe");
  sl.click(fach(), "Deutsch", true);
  sl.edit("Spalten", klasse(), (s) => ({ ...s, columns: 2, style: "SlicerStyleDark2" }));
  const kc = sl.cacheOf(klasse());
  sl.editCache("Datenschnitteinstellungen", kc, (c) => { c.crossFilter = "showItemsWithNoData"; c.sort = "descending"; });
  // A PivotTable on a new sheet with a slicer on Fach.
  app.addSheet(true);
  app.pivots.add(app.sheet, { name: "PivotNoten", source: { sheet: "Tabelle1", r0: 0, c0: 0, r1: 8, c1: 3 }, at: { row: 2, col: 0 }, rows: [1], cols: [], values: [{ field: 3, func: "sum" }], filters: [] });
  app.grid.selectCell(3, 0);
  sl.add([{ sourceName: "Fach", pivot: { sheet: app.sheet.name, name: "PivotNoten" } }]);
  const pf = sl.byName("Fach 1");
  sl.click(pf, "Englisch");
  const out = {};
  for (const s of app.wb.sheets) {
    out[s.name] = { hidden: [...s.rows].filter(([, i]) => i.hidden).map(([r]) => r + 1).sort((a, b) => a - b),
      slicers: (s.slicers ?? []).map((x) => ({ name: x.name, caption: x.caption, columns: x.columns ?? 1, style: x.style ?? "SlicerStyleLight1",
        items: sl.items(sl.cacheOf(x)).map((i) => [i.text, i.selected, i.hasData]) })),
      cells: Array.from({ length: 8 }, (_, r) => Array.from({ length: 2 }, (_, c) => { const v = s.get(r, c)?.v; return v === undefined ? null : v; })) };
  }
  return out;
})()"""

ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826273: "#VALUE!"}
CROSS = {1: "none", 2: "showItemsWithDataAtTop", 3: "showItemsWithNoData", 4: "hide"}


def excel_read(path: Path) -> dict:
    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(path))
        x.app.CalculateFull()
        try:
            wb.RefreshAll()
            x.app.CalculateUntilAsyncQueriesDone()
        except Exception:  # noqa: BLE001
            pass
        # Excel may still be busy refreshing the PivotTable ("call rejected"): wait until it answers.
        import time
        for _ in range(40):
            try:
                _ = wb.SlicerCaches.Count
                break
            except Exception:  # noqa: BLE001
                time.sleep(0.5)
        out: dict = {"caches": {}, "sheets": {}}
        for i in range(1, wb.SlicerCaches.Count + 1):
            sc = wb.SlicerCaches(i)
            items = []
            for k in range(1, sc.SlicerItems.Count + 1):
                it = sc.SlicerItems(k)
                try:
                    has = bool(it.HasData)
                except Exception:  # noqa: BLE001
                    has = None
                items.append([it.Name, bool(it.Selected), has])
            out["caches"][sc.Name] = {"source": sc.SourceName, "sort": sc.SortItems, "cross": CROSS.get(sc.CrossFilterType), "items": items,
                                      "slicers": [{"name": s.Name, "caption": s.Caption, "columns": s.NumberOfColumns, "style": s.Style.Name if not isinstance(s.Style, str) else s.Style}
                                                  for s in sc.Slicers]}
        for ws in wb.Worksheets:
            used = ws.UsedRange
            hidden = [r for r in range(1, used.Row + used.Rows.Count) if ws.Rows(r).Hidden]
            rep = None
            if ws.PivotTables().Count:
                pt = ws.PivotTables(1)
                pt.RefreshTable()
                rep = [list(r) for r in pt.TableRange1.Value]
            out["sheets"][ws.Name] = {"hidden": hidden, "pivot": rep}
        wb.Close(False)
        return out
    finally:
        x.close()


def main() -> None:
    report: dict = {"raabe->excel": [], "excel->raabe": [], "look": []}
    ours_file = TMP / "raabe-slicers-ours.xlsx"
    clicked = TMP / "raabe-slicers-clicked.xlsx"
    rec = json.loads((FIX / "excel-slicers.json").read_text(encoding="utf-8"))
    c = CalcSide(PORT)
    try:
        c.reset(Case("slicers", DATA, [], "A1", sheets=1))
        ours = c.js(BUILD)
        c.js(f"__calc.save({json.dumps(str(ours_file))})")
        # 2. Excel's file.
        c.js(f"__calc.open({json.dumps(str(FIX / 'excel-slicers.xlsx'))})")
        got = c.js("""(() => { const app = __calc.app, sl = app.slicers; const out = {};
          app.wb.sheets.forEach((s, i) => { out[s.name] = { hidden: [...s.rows].filter(([, x]) => x.hidden).map(([r]) => r + 1).sort((a, b) => a - b),
            slicers: Object.fromEntries((s.slicers ?? []).map((x) => [x.name, sl.cacheOf(x) ? sl.items(sl.cacheOf(x)).map((it) => [it.text, it.selected, it.hasData]) : null])) }; });
          return out; })()""")
        # A click in RaabeCalc on Excel's file: Standard > Fach > Mathe only.
        c.js("""(() => { const app = __calc.app; const i = app.wb.sheets.findIndex((s) => s.name === "Standard"); app.wb.active = i; app.grid.refreshLayout();
          const s = app.slicers.byName("Fach"); app.slicers.click(s, "Mathe"); })()""")
        c.js(f"__calc.save({json.dumps(str(clicked))})")
        # 3. The look: Excel's two-filters case as RaabeCalc draws it.
        look = c.js("""(async () => { const app = __calc.app; const i = app.wb.sheets.findIndex((s) => s.name === "Zwei"); app.wb.active = i; app.grid.refreshLayout(); app.grid.contentChanged();
          await new Promise((r) => setTimeout(r, 300)); app.slicers.place();
          const out = {};
          for (const e of document.querySelectorAll(".slicer-layer .slicer")) {
            const sb = e.getBoundingClientRect(); const z = app.sheet.view.zoom / 100;
            const btns = [...e.querySelectorAll(".sl-btn")].map((b) => { const r = b.getBoundingClientRect(); const cs = getComputedStyle(b);
              return { text: b.textContent, x: (r.left - sb.left) / z, y: (r.top - sb.top) / z, w: r.width / z, h: r.height / z, fill: cs.backgroundColor, border: cs.borderTopColor, color: cs.color }; });
            const head = e.querySelector(".sl-head"); const cap = head ? head.querySelector(".sl-cap").textContent : null;
            out[cap] = { w: sb.width / z, h: sb.height / z, border: getComputedStyle(e).borderTopColor, btns };
          }
          return out; })()""")
    finally:
        c.close()
    report["ours"] = ours

    # 1. RaabeCalc's file in Excel.
    ex = excel_read(ours_file)
    report["excel"] = ex
    checks = report["raabe->excel"]

    def add(name, ok, detail=None):
        checks.append({"check": name, "ok": bool(ok), **({"detail": detail} if not ok and detail is not None else {})})
    caches = ex["caches"]
    add("Excel finds three slicer caches", len(caches) == 3, list(caches))
    by_slicer = {s["name"]: (k, v) for k, v in caches.items() for s in v["slicers"]}
    for sheet, st in ours.items():
        for s in st["slicers"]:
            hit = by_slicer.get(s["name"])
            add(f"{s['name']}: Excel has the slicer", hit is not None)
            if not hit:
                continue
            name, cache = hit
            xs = next(x for x in cache["slicers"] if x["name"] == s["name"])
            add(f"{s['name']}: caption, columns, style", (xs["caption"], xs["columns"], xs["style"]) == (s["caption"], s["columns"], s["style"]), [xs, s])
            # Excel lists the items in its own order (data first unless set otherwise), as RaabeCalc shows them.
            add(f"{s['name']}: items with Selected and HasData as RaabeCalc shows them", cache["items"] == s["items"], [cache["items"], s["items"]])
    k = by_slicer.get("Klasse")
    if k:
        add("Klasse: absteigend and Elemente ohne Daten in place", (k[1]["sort"], k[1]["cross"]) == (3, "showItemsWithNoData"), k[1])
    add("the table's hidden rows are RaabeCalc's", ex["sheets"]["Tabelle1"]["hidden"] == ours["Tabelle1"]["hidden"], [ex["sheets"]["Tabelle1"]["hidden"], ours["Tabelle1"]["hidden"]])
    piv = [n for n in ex["sheets"] if ex["sheets"][n]["pivot"]]
    if piv:
        rep = ex["sheets"][piv[0]]["pivot"]
        mine = [r for r in ours[piv[0]]["cells"][2:] if r[0] is not None]
        add("the PivotTable's report after Excel's refresh is RaabeCalc's", [[r[0], r[1]] for r in rep] == [[a, b] for a, b in mine][:len(rep)], [rep, mine])
    else:
        add("the PivotTable exists in Excel", False)

    # 2. Excel's file in RaabeCalc.
    for case, cr in rec["cases"].items():
        # The file holds each sheet's last state: "default2" and "cleared" stand for "default" and "two-filters".
        if case in ("default", "two-filters"):
            continue
        sheet = cr["sheet"]
        mine = got.get(sheet, {})
        for s in cr["slicers"]:
            want = [[i["name"], i["selected"], i["hasData"]] for i in s["items"]]
            if s["crossFilter"] == 4:
                want = [w for w in want if w[2]]
            have = mine.get("slicers", {}).get(s["name"])
            report["excel->raabe"].append({"check": f"{case} {s['name']}: items", "ok": have == want, **({"have": have, "want": want} if have != want else {})})
        if cr.get("hiddenRows") is not None:
            report["excel->raabe"].append({"check": f"{case}: hidden rows", "ok": mine.get("hidden") == cr["hiddenRows"], **({"have": mine.get("hidden"), "want": cr["hiddenRows"]} if mine.get("hidden") != cr["hiddenRows"] else {})})
    ex2 = excel_read(clicked)
    std = ex2["caches"].get("Datenschnitt_Fach", {})
    report["excel->raabe"].append({"check": "clicked in RaabeCalc: Excel's Fach slicer shows only Mathe", "ok": [i[:2] for i in std.get("items", [])] == [["Deutsch", False], ["Englisch", False], ["Mathe", True]], "items": std.get("items")})
    report["excel->raabe"].append({"check": "clicked in RaabeCalc: Excel hides the other rows", "ok": ex2["sheets"]["Standard"]["hidden"] == [5, 6, 7, 8, 9], "hidden": ex2["sheets"]["Standard"]["hidden"]})
    report["excel->raabe"].append({"check": "clicked in RaabeCalc: the other sheets' slicers unchanged", "ok": len(ex2["caches"]) == 18})

    # 3. The look.
    try:
        from PIL import Image
        for cap, png in (("Fach", "excel-cleared-Fach 2.png"), ("Klasse", "excel-cleared-Klasse 1.png")):
            im = Image.open(SHOTS / png).convert("RGB")
            ours_l = look.get(cap)
            if not ours_l:
                report["look"].append({"check": f"{cap}: drawn", "ok": False})
                continue
            for b in ours_l["btns"]:
                cx, cy = int(b["x"] + b["w"] - 12), int(b["y"] + b["h"] / 2)
                ex_fill = im.getpixel((cx, cy))
                mine = tuple(int(v) for v in b["fill"][b["fill"].index("(") + 1:-1].split(",")[:3])
                diff = max(abs(a - c) for a, c in zip(ex_fill, mine))
                # Excel's button edges: scan the column at the button's middle for its border rows.
                col = [im.getpixel((int(b["x"] + 30), y)) for y in range(im.height)]
                report["look"].append({"check": f"{cap} {b['text']}: fill colour within 3 of Excel's", "ok": diff <= 3, "excel": ex_fill, "ours": mine})
                inside = [y for y in range(int(b["y"]) - 3, int(b["y"] + b["h"]) + 3) if 0 <= y < im.height and col[y] == ex_fill]
                if inside:
                    top, bottom = inside[0] - 1, inside[-1] + 1
                    report["look"].append({"check": f"{cap} {b['text']}: top and height within 1 px of Excel's", "ok": abs(top - b["y"]) <= 1 and abs((bottom - top + 1) - b["h"]) <= 1.5,
                                           "excel": [top, bottom - top + 1], "ours": [round(b["y"], 1), round(b["h"], 1)]})
    except FileNotFoundError:
        report["look"].append({"check": "Excel's rendering found (run probe_excel_slicers.py first)", "ok": False})

    for key in ("raabe->excel", "excel->raabe", "look"):
        v = report[key]
        print(key, sum(1 for r in v if r["ok"]), "/", len(v))
        for r in v:
            if not r["ok"]:
                print("   FAIL", json.dumps(r, ensure_ascii=False, default=str)[:400])
    (HERE / "results" / f"{datetime.date.today().isoformat()}-calc-slicers.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=1, default=str).replace(str(TMP), "<TEMP>").replace(str(TMP).replace("\\", "\\\\"), "<TEMP>"), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
