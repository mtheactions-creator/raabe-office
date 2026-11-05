"""PivotTables against Excel (D-066), both ways, without real input:
1. RaabeCalc (off screen) makes the PivotTables of probe_excel_pivot.py's configurations and saves the
   workbook; our own invisible Excel opens it, must see each as a PivotTable with the same fields, and
   after its refresh on opening its report must equal RaabeCalc's cells.
2. Excel makes them and saves; RaabeCalc opens the file and must read the same PivotTables.
Result: results/<date>-pivots.json.
"""
from __future__ import annotations

import datetime
import json
import tempfile
import time
from pathlib import Path

from feellab.calc_catalogue import Case, CalcSide
from feellab.excel import ExcelTarget
from probe_excel_pivot import CONFIGS, DATA, FUNC

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir())
NAMES = DATA[0]


def col(c: int) -> str:
    s = ""
    c += 1
    while c:
        c, m = divmod(c - 1, 26)
        s = chr(65 + m) + s
    return s


def model(name: str, cfg: dict, sheet: str) -> dict:
    idx = NAMES.index
    return {
        "name": "P" + name.replace("-", ""), "source": {"sheet": "Daten", "r0": 0, "c0": 0, "r1": len(DATA) - 1, "c1": len(NAMES) - 1},
        "at": {"row": 2, "col": 0}, "rows": [idx(f) for f in cfg.get("rows", [])], "cols": [idx(f) for f in cfg.get("cols", [])],
        "values": [{"field": idx(f), "func": fn} for f, fn in cfg["values"]],
        "filters": [{"field": idx(cfg["filter"][0]), "items": [cfg["filter"][1]]}] if "filter" in cfg else [],
        "layout": {1: "tabular", 2: "outline"}.get(cfg.get("layout"), "compact"),
    }


def busy(fn, tries: int = 60):
    """Excel refuses calls while it refreshes the PivotTables it just opened: try again."""
    for k in range(tries):
        try:
            return fn()
        except Exception as e:  # noqa: BLE001
            if "-2147418111" not in str(e) or k == tries - 1:
                raise
            time.sleep(0.5)


def norm(v):
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, float):
        return round(v, 9)
    return v


def main() -> None:
    raabe_file, excel_file = TMP / "raabe-pivots.xlsx", TMP / "excel-pivots.xlsx"
    report: dict = {"raabe->excel": [], "excel->raabe": []}
    setup = {f"{col(c)}{r + 1}": str(v) for r, row in enumerate(DATA) for c, v in enumerate(row)}
    c = CalcSide()
    ours: dict = {}
    try:
        c.reset(Case("pivots", setup, [], "x", sheets=1))
        c.js("__calc.app.sheet.name = 'Daten'")
        for i, (name, cfg) in enumerate(CONFIGS.items()):
            c.js(f"""(() => {{ const app = __calc.app; app.addSheet(true); app.sheet.name = {json.dumps('P' + str(i + 1))};
              app.pivots.add(app.sheet, {json.dumps(model(name, cfg, 'P' + str(i + 1)))}); }})()""")
            ours[name] = c.js("""(() => { const s = __calc.app.sheet; const p = s.pivots[0]; const a = p.area; const out = [];
              for (let r = a.r0; r <= a.r1; r++) { const row = []; for (let k = a.c0; k <= a.c1; k++) { const v = s.value(r, k); row.push(v === undefined ? null : v); } out.push(row); }
              return { area: a, cells: out }; })()""")
        c.js(f"__calc.save({json.dumps(str(raabe_file))})")
    finally:
        c.close()
    x = ExcelTarget(visible=False)
    try:
        wb = busy(lambda: x.app.Workbooks.Open(str(raabe_file)))
        for _ in range(60):  # the refresh on opening runs on after Open returns
            if busy(lambda: x.app.Ready):
                break
            time.sleep(0.5)
        time.sleep(1)
        for i, (name, cfg) in enumerate(CONFIGS.items()):
            ws = busy(lambda: wb.Worksheets(f"P{i + 1}"))
            rec = {"config": name}
            try:
                pt = busy(lambda: ws.PivotTables(1))
                busy(pt.RefreshTable)
                rec["rows"] = busy(lambda: [pt.RowFields.Item(k).Name for k in range(1, pt.RowFields.Count + 1)]) if cfg.get("rows") else []
                rec["data"] = busy(lambda: [pt.DataFields.Item(k).Name for k in range(1, pt.DataFields.Count + 1)])
                vals = busy(lambda: pt.TableRange2.Value)
                got = [[norm(v) for v in row] for row in (vals if isinstance(vals, tuple) else ((vals,),))]
                want = [[norm(v) for v in row] for row in ours[name]["cells"]]
                rec["ok"] = got == want and rec["rows"] == cfg.get("rows", [])
                if not rec["ok"]:
                    rec["excel"] = got
                    rec["raabe"] = want
            except Exception as e:  # noqa: BLE001
                rec["ok"] = False
                rec["error"] = str(e)[:150]
            report["raabe->excel"].append(rec)
            print("R->E", name, rec["ok"], rec.get("error", ""), flush=True)
        wb.Close(False)
        # 2. Excel makes them.
        wb = x.app.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "Daten"
        for r, row in enumerate(DATA, start=1):
            for k, v in enumerate(row, start=1):
                ws.Cells(r, k).Value = v
        src = ws.Range(ws.Cells(1, 1), ws.Cells(len(DATA), len(NAMES)))
        for i, (name, cfg) in enumerate(CONFIGS.items()):
            dest = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            dest.Name = f"P{i + 1}"
            pt = wb.PivotCaches().Create(1, src).CreatePivotTable(dest.Range("A3"), "P" + name.replace("-", ""))
            for f in cfg.get("rows", []):
                pt.PivotFields(f).Orientation = 1
            for f in cfg.get("cols", []):
                pt.PivotFields(f).Orientation = 2
            if "filter" in cfg:
                fld = pt.PivotFields(cfg["filter"][0])
                fld.Orientation = 3
                fld.CurrentPage = cfg["filter"][1]
            for f, fn in cfg["values"]:
                pt.AddDataField(pt.PivotFields(f), None, FUNC[fn])
            if "layout" in cfg:
                pt.RowAxisLayout(cfg["layout"])
        wb.SaveAs(str(excel_file), 51)
        wb.Close(False)
    finally:
        x.close()
    c = CalcSide()
    try:
        c.js(f"__calc.open({json.dumps(str(excel_file))})")
        got = c.js("__calc.app.wb.sheets.filter((s) => s.pivots).map((s) => ({ sheet: s.name, p: (({ raw, rawKey, ...r }) => r)(s.pivots[0]) }))")
    finally:
        c.close()
    for i, (name, cfg) in enumerate(CONFIGS.items()):
        m = next((g["p"] for g in got if g["sheet"] == f"P{i + 1}"), None)
        want = model(name, cfg, f"P{i + 1}")
        ok = m is not None and all(m.get(k) == want[k] for k in ("rows", "cols", "values", "filters")) and (m.get("layout") or "compact") == want["layout"] \
            and m["source"] == want["source"]
        report["excel->raabe"].append({"config": name, "ok": ok, **({} if ok else {"got": m, "want": want})})
        print("E->R", name, ok, flush=True)
    (HERE / "results" / f"{datetime.date.today().isoformat()}-pivots.json").write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    for k, v in report.items():
        print(k, sum(1 for r in v if r["ok"]), "/", len(v))


if __name__ == "__main__":
    main()
