"""Excel tables against Excel (D-066), both ways, without real input:
1. RaabeCalc (off screen) makes tables with different options and saves; our own invisible Excel opens
   the file and must see each ListObject with the same name, range, style, options and columns.
2. Excel makes the same tables and saves; RaabeCalc opens the file and must read the same models.
Result: results/<date>-calc-tables.json. Also leaves TEMP/excel-tables.xlsx for shot_compare.py.
"""
from __future__ import annotations

import datetime
import json
import tempfile
from pathlib import Path

from feellab.calc_catalogue import Case, CalcSide
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir())
DATA = [["Name", "Fach", "Note"], ["Anna", "Mathe", "2"], ["Ben", "Deutsch", "3"], ["Cem", "Mathe", "1"], ["Dora", "Englisch", "2"]]
CASES = [
    {"name": "Noten", "at": (0, 0), "style": "TableStyleMedium2", "opts": {}},
    {"name": "Erste", "at": (0, 5), "style": "TableStyleMedium9", "opts": {"firstColumn": True, "rowStripes": False, "columnStripes": True}},
    {"name": "Summe", "at": (8, 0), "style": "TableStyleLight1", "opts": {"lastColumn": True}, "totals": True},
    {"name": "Dunkel", "at": (8, 5), "style": "TableStyleDark1", "opts": {"filterButton": False}},
]
OPTS = {"firstColumn": "ShowTableStyleFirstColumn", "lastColumn": "ShowTableStyleLastColumn", "rowStripes": "ShowTableStyleRowStripes",
        "columnStripes": "ShowTableStyleColumnStripes", "filterButton": "ShowAutoFilterDropDown"}
DEFAULTS = {"firstColumn": False, "lastColumn": False, "rowStripes": True, "columnStripes": False, "filterButton": True}


def col(c: int) -> str:
    s = ""
    c += 1
    while c:
        c, m = divmod(c - 1, 26)
        s = chr(65 + m) + s
    return s


def addr(r: int, c: int) -> str:
    return f"{col(c)}{r + 1}"


def setup() -> dict:
    out = {}
    for k in CASES:
        r0, c0 = k["at"]
        for i, row in enumerate(DATA):
            for j, v in enumerate(row):
                out[addr(r0 + i, c0 + j)] = v
    return out


def rng(k: dict) -> str:
    r0, c0 = k["at"]
    return f"{addr(r0, c0)}:{addr(r0 + len(DATA) - 1, c0 + len(DATA[0]) - 1)}"


def main() -> None:
    raabe_file, excel_file = TMP / "raabe-tables.xlsx", TMP / "excel-tables.xlsx"
    report: dict = {"raabe->excel": [], "excel->raabe": []}
    c = CalcSide()
    try:
        c.reset(Case("tables", setup(), [], "A1", sheets=1))
        for k in CASES:
            opts = {**k["opts"]}
            c.js(f"""(() => {{ const app = __calc.app; const t = app.tables.create(__calc.parse({json.dumps(rng(k))}), true, {json.dumps(k['style'])});
              const sheet = app.sheet; const i = sheet.tables.indexOf(t); const x = sheet.tables[i];
              x.name = {json.dumps(k['name'])}; Object.assign(x, {json.dumps(opts)});
              if ({json.dumps(bool(k.get('totals')))}) {{ app.grid.select({{ r0: x.ref.r0, c0: x.ref.c0, r1: x.ref.r0, c1: x.ref.c0 }}); app.tables.setTotals(true); }} }})()""")
        c.js(f"__calc.save({json.dumps(str(raabe_file))})")
    finally:
        c.close()
    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(raabe_file))
        ws = wb.Worksheets(1)
        seen = {ws.ListObjects(i).Name: ws.ListObjects(i) for i in range(1, ws.ListObjects.Count + 1)}
        for k in CASES:
            lo = seen.get(k["name"])
            bad = []
            got = None
            if lo is None:
                bad = ["missing"]
            else:
                got = {"range": lo.Range.Address.replace("$", ""), "style": lo.TableStyle.Name if lo.TableStyle else "", "totals": bool(lo.ShowTotals),
                       "cols": [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)], **{o: bool(getattr(lo, p)) for o, p in OPTS.items()}}
                want_range = rng(k)
                if k.get("totals"):
                    r0, c0 = k["at"]
                    want_range = f"{addr(r0, c0)}:{addr(r0 + len(DATA), c0 + len(DATA[0]) - 1)}"
                if got["range"] != want_range:
                    bad.append("range")
                if got["style"] != k["style"]:
                    bad.append("style")
                if got["totals"] != bool(k.get("totals")):
                    bad.append("totals")
                if got["cols"] != DATA[0]:
                    bad.append("cols")
                for o in OPTS:
                    if got[o] != k["opts"].get(o, DEFAULTS[o]):
                        bad.append(o)
            report["raabe->excel"].append({"table": k["name"], "ok": not bad, **({"bad": bad, "excel": got} if bad else {})})
            print("R->E", k["name"], not bad, bad, flush=True)
        wb.Close(False)
        wb = x.app.Workbooks.Add()
        ws = wb.Worksheets(1)
        for a, v in setup().items():
            ws.Range(a).Value = float(v) if v.isdigit() else v
        for k in CASES:
            lo = ws.ListObjects.Add(1, ws.Range(rng(k)), None, 1)
            lo.Name = k["name"]
            lo.TableStyle = k["style"]
            for o, p in OPTS.items():
                if o in k["opts"]:
                    setattr(lo, p, k["opts"][o])
            if k.get("totals"):
                lo.ShowTotals = True
        wb.SaveAs(str(excel_file), 51)
        wb.Close(False)
    finally:
        x.close()
    c = CalcSide()
    try:
        c.js(f"__calc.open({json.dumps(str(excel_file))})")
        got = c.js("__calc.app.sheet.tables?.map((t) => (({ raw, rawKey, ...r }) => r)(t)) ?? []")
    finally:
        c.close()
    by = {t["name"]: t for t in got}
    for k in CASES:
        t = by.get(k["name"])
        bad = []
        if t is None:
            bad = ["missing"]
        else:
            if t["style"] != k["style"]:
                bad.append("style")
            if bool(t["totals"]) != bool(k.get("totals")):
                bad.append("totals")
            if [cc["name"] for cc in t["columns"]] != DATA[0]:
                bad.append("cols")
            for o in OPTS:
                if bool(t[o]) != k["opts"].get(o, DEFAULTS[o]):
                    bad.append(o)
        report["excel->raabe"].append({"table": k["name"], "ok": not bad, **({"bad": bad, "got": t} if bad else {})})
        print("E->R", k["name"], not bad, bad, flush=True)
    (HERE / "results" / f"{datetime.date.today().isoformat()}-calc-tables.json").write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    for key, v in report.items():
        print(key, sum(1 for r in v if r["ok"]), "/", len(v))


if __name__ == "__main__":
    main()
