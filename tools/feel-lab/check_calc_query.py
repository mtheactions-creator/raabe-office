"""Workbooks with Power Query queries and a data model in RaabeCalc (D-080, D-092), without real input:
1. Our own invisible Excel makes a workbook whose query reads a local CSV (written to TEMP by this script),
   loads it as a table on a sheet and also into the data model (Connections.Add2 ... CreateModelConnection),
   refreshes and saves it (TEMP/raabe-query-excel.xlsx; never committed: the query holds the CSV's path).
2. RaabeCalc (off screen, never focused) opens it: the table's cells and the sheets must be Excel's; it
   saves it unchanged and again after typing into another sheet.
3. Excel opens both of RaabeCalc's files: the queries (Queries.Count, formulas), the connections, the data
   model's tables, the loaded table and its cells must be as before, and a refresh must still work.
Result: results/<date>-calc-query.json (no paths).
"""
from __future__ import annotations

import datetime
import json
import os
import tempfile
import time
import zipfile
from pathlib import Path

from feellab import calc_catalogue, winapi
from feellab.calc_catalogue import CalcSide
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir())
PORT = 9473
os.environ["RAABE_FEEL_BACKGROUND"] = "1"
calc_catalogue.winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4000, 0, w, h)

CSV = "Name;Fach;Note\nAnna;Mathe;2\nBen;Deutsch;3\nCem;Mathe;1\nDora;Englisch;4\n"


def excel_state(wb) -> dict:
    out = {"queries": [], "connections": [], "model": [], "tables": [], "sheets": [ws.Name for ws in wb.Worksheets]}
    for i in range(1, wb.Queries.Count + 1):
        q = wb.Queries(i)
        out["queries"].append({"name": q.Name, "formula": q.Formula.replace(str(TMP), "<TEMP>")})
    for i in range(1, wb.Connections.Count + 1):
        c = wb.Connections(i)
        out["connections"].append({"name": c.Name, "type": c.Type, "inModel": bool(getattr(c, "InModel", False))})
    try:
        m = wb.Model
        out["model"] = [m.ModelTables(i).Name for i in range(1, m.ModelTables.Count + 1)]
    except Exception:  # noqa: BLE001
        out["model"] = None
    for ws in wb.Worksheets:
        for lo in ws.ListObjects:
            out["tables"].append({"sheet": ws.Name, "name": lo.Name, "range": lo.Range.Address.replace("$", ""), "source": lo.SourceType,
                                  "values": [list(r) for r in lo.Range.Value]})
    return out


def make_excel(path: Path, csv: Path) -> dict:
    x = ExcelTarget(visible=False)
    try:
        wb = x.wb
        m = (f'let\n    Quelle = Csv.Document(File.Contents("{csv}"),[Delimiter=";", Columns=3, Encoding=65001, QuoteStyle=QuoteStyle.None]),\n'
             '    Header = Table.PromoteHeaders(Quelle, [PromoteAllScalars=true]),\n'
             '    Typen = Table.TransformColumnTypes(Header,{{"Name", type text}, {"Fach", type text}, {"Note", Int64.Type}})\nin\n    Typen')
        wb.Queries.Add("Noten", m)
        ws = wb.Worksheets(1)
        ws.Name = "Noten"
        lo = ws.ListObjects.Add(0, 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Noten;Extended Properties=""', None, 1, ws.Range("$A$1"))
        qt = lo.QueryTable
        qt.CommandType = 2
        qt.CommandText = ["SELECT * FROM [Noten]"]
        qt.ListObject.DisplayName = "Noten"
        qt.Refresh(False)
        # A second query only in the data model.
        wb.Queries.Add("NotenModell", "let\n    Quelle = Noten\nin\n    Quelle")
        wb.Connections.Add2("Abfrage - NotenModell", "", 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=NotenModell;Extended Properties=""',
                            '"NotenModell"', 6, True, False)
        ws2 = wb.Worksheets.Add(After=ws)
        ws2.Name = "Notizen"
        ws2.Range("A1").Value = "frei"
        for prop in ("Author", "Last author"):
            try:
                wb.BuiltinDocumentProperties.Item(prop).Value = "Raabe Office Test"
            except Exception:  # noqa: BLE001
                pass
        if path.exists():
            path.unlink()
        wb.SaveAs(str(path), 51)
        return excel_state(wb)
    finally:
        x.close()


def read_excel(path: Path) -> dict:
    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(path))
        st = excel_state(wb)
        try:
            wb.RefreshAll()
            x.app.CalculateUntilAsyncQueriesDone()
            st["refreshed"] = True
        except Exception as e:  # noqa: BLE001
            st["refreshed"] = f"error {type(e).__name__}"
        st["afterRefresh"] = [t["values"] for t in excel_state(wb)["tables"]]
        wb.Close(False)
        return st
    finally:
        x.close()


def parts(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as z:
        return sorted(n for n in z.namelist() if not n.startswith("docProps"))


def main() -> None:
    csv = TMP / "raabe-query-noten.csv"
    csv.write_text(CSV, encoding="utf-8")
    src = TMP / "raabe-query-excel.xlsx"
    report: dict = {"checks": []}
    before = make_excel(src, csv)
    report["excel"] = before
    c = CalcSide(PORT)
    try:
        c.js(f"__calc.open({json.dumps(str(src))})")
        time.sleep(1.0)
        ours = c.js("""(() => { const app = __calc.app; const wb = app.wb; return { sheets: wb.sheets.map((s) => s.name),
          tables: wb.sheets.flatMap((s) => (s.tables ?? []).map((t) => ({ sheet: s.name, name: t.name, ref: t.ref }))),
          a2: wb.sheets[0].get(1, 0)?.v ?? null, c3: wb.sheets[0].get(2, 2)?.v ?? null,
          connections: app.getData ? app.getData.connections().map((x) => x.name) : null }; })()""")
        report["calc"] = ours
        same = TMP / "raabe-query-same.xlsx"
        edited = TMP / "raabe-query-edited.xlsx"
        c.js(f"__calc.save({json.dumps(str(same))})")
        c.js("""(() => { const app = __calc.app; app.activateSheet(1); const r = __calc.parse("B2"); app.grid.selectCell(r.r0, r.c0); app.gridHost().commit(r, "geändert", {}); })()""")
        c.js(f"__calc.save({json.dumps(str(edited))})")
        # Daten abrufen > Aus Arbeitsmappe (the Navigator's choice) and Aus Text/CSV into this workbook.
        loaded = c.js(f"""(async () => {{ const app = __calc.app; await app.getData.loadItem({json.dumps(str(src))}, "Noten");
          const s = app.sheet; const t = s.tables?.[0];
          const a = {{ sheet: s.name, table: t?.name, ref: t?.ref, style: t?.style, rows: Array.from({{ length: 5 }}, (_, r) => [0, 1, 2].map((c) => s.value(r, c))) }};
          await app.getData.csv({json.dumps(str(csv))});
          const s2 = app.sheet, t2 = s2.tables?.[0];
          a.csv = {{ sheet: s2.name, table: t2?.name, rows: Array.from({{ length: 5 }}, (_, r) => [0, 1, 2].map((c) => s2.value(r, c))) }};
          return a; }})()""")
        report["loaded"] = loaded
    finally:
        c.close()
    checks = report["checks"]
    checks.append({"check": "RaabeCalc shows Excel's sheets", "ok": ours["sheets"] == before["sheets"]})
    checks.append({"check": "RaabeCalc reads the loaded table", "ok": any(t["name"] == "Noten" for t in ours["tables"]) and ours["a2"] == "Anna" and ours["c3"] == 3})
    if ours.get("connections") is not None:
        checks.append({"check": "Vorhandene Verbindungen lists the file's connections", "ok": sorted(ours["connections"]) == sorted(x["name"] for x in before["connections"])})
    want_rows = before["tables"][0]["values"]
    norm = lambda rows: [[int(v) if isinstance(v, float) and v.is_integer() else v for v in r] for r in rows]
    checks.append({"check": "Aus Arbeitsmappe: the table on a new sheet named after it (Noten (2), table Noten_2, TableStyleMedium2)",
                   "ok": loaded["sheet"] == "Noten (2)" and loaded["table"] == "Noten_2" and loaded["style"] == "TableStyleMedium2" and norm(loaded["rows"]) == norm(want_rows), "got": loaded})
    checks.append({"check": "Aus Text/CSV: the file's data as a table on a new sheet of this workbook",
                   "ok": loaded["csv"]["sheet"] == "raabe-query-noten" and norm(loaded["csv"]["rows"]) == norm(want_rows), "got": loaded["csv"]})
    orig_parts = parts(src)
    for label, f in (("unchanged", same), ("edited", edited)):
        p = parts(f)
        keep = [n for n in orig_parts if n.startswith(("customXml", "xl/connections", "xl/model", "xl/queryTables"))]
        checks.append({"check": f"{label}: query, connection and model parts kept", "ok": all(n in p for n in keep), "missing": [n for n in keep if n not in p]})
        st = read_excel(f)
        report[f"excel-{label}"] = st
        checks.append({"check": f"{label}: Excel finds the queries", "ok": [q["name"] for q in st["queries"]] == [q["name"] for q in before["queries"]]
                       and [q["formula"] for q in st["queries"]] == [q["formula"] for q in before["queries"]]})
        checks.append({"check": f"{label}: Excel finds the connections", "ok": [(x["name"], x["inModel"]) for x in st["connections"]] == [(x["name"], x["inModel"]) for x in before["connections"]]})
        checks.append({"check": f"{label}: Excel finds the data model's tables", "ok": st["model"] == before["model"]})
        checks.append({"check": f"{label}: the loaded table is the same", "ok": [(t["name"], t["range"], t["values"]) for t in st["tables"]] == [(t["name"], t["range"], t["values"]) for t in before["tables"]]})
        checks.append({"check": f"{label}: Excel refreshes the query", "ok": st["refreshed"] is True and st["afterRefresh"] == [t["values"] for t in before["tables"]]})
    for ch in checks:
        print(ch["ok"], ch["check"], ch.get("missing", ""), flush=True)
    report["passed"] = f"{sum(1 for x in checks if x['ok'])}/{len(checks)}"
    print(report["passed"])
    text = json.dumps(report, ensure_ascii=False, indent=1, default=str).replace(str(TMP), "<TEMP>").replace(str(TMP).replace("\\", "\\\\"), "<TEMP>")
    (HERE / "results" / f"{datetime.date.today().isoformat()}-calc-query.json").write_text(text, encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
