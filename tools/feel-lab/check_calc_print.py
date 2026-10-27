"""Printing against Excel (package F): the same sheet exported as PDF by Excel (its own invisible
instance, ExportAsFixedFormat) and by RaabeCalc (Als PDF exportieren), then compared page by
page: the number of pages, the page size, and on each page the number of text lines and where
the first and last line sit. No real input. Result: results/<date>-calc-print.json.

    .venv/Scripts/python.exe check_calc_print.py
"""
from __future__ import annotations

import datetime
import json
import tempfile
import time
from pathlib import Path

from feellab import pdftext
from feellab.calc_catalogue import CalcSide
from feellab.excel import ExcelTarget

TMP = Path(tempfile.gettempdir())
CASES = {
    # name: (rows, cols, page setup as JS for RaabeCalc, the same for Excel)
    "120x12 portrait": (120, 12, {}, {}),
    "60x5 landscape": (60, 5, {"orientation": "landscape"}, {"Orientation": 2}),
    "200x3 fit width": (200, 3, {"fitWidth": 1}, {"Zoom": False, "FitToPagesWide": 1, "FitToPagesTall": False}),
}


def summary(path: Path) -> dict:
    pages = pdftext.read(path)
    out = []
    for p in pages:
        lines = pdftext.baselines(p)
        out.append({"w": round(p.width), "h": round(p.height), "lines": len(lines),
                    "first": lines[0]["y"] if lines else None, "last": lines[-1]["y"] if lines else None,
                    "left": min((l["x"] for l in lines), default=None)})
    return {"pages": len(pages), "detail": out}


def main() -> None:
    results: dict = {}
    x = ExcelTarget(visible=False)
    try:
        st = x.wb.Styles("Standard")
        st.Font.Name = "Calibri"
        st.Font.Size = 11
        for name, (rows, cols, _js, ps) in CASES.items():
            ws = x.wb.Worksheets.Add()
            for r in range(1, rows + 1):
                ws.Range(ws.Cells(r, 1), ws.Cells(r, cols)).Value = tuple(r * 100 + c for c in range(1, cols + 1))
            for k, v in ps.items():
                setattr(ws.PageSetup, k, v)
            out = TMP / f"raabe-print-excel-{len(results)}.pdf"
            ws.ExportAsFixedFormat(0, str(out))
            results[name] = {"excel": summary(out)}
            out.unlink(missing_ok=True)
    finally:
        x.close()
    c = CalcSide()
    try:
        for i, (name, (rows, cols, js, _ps)) in enumerate(CASES.items()):
            setup = {}
            from feellab.calc_catalogue import addr_rc  # noqa: F401
            cells = "".join(f"s.set({r}, {col}, {{ v: {(r + 1) * 100 + col + 1} }});" for r in range(rows) for col in range(cols))
            c.reset(type("C", (), {"setup": setup, "start": "A1", "sheets": 1})())
            c.js(f"(() => {{ const app = __calc.app; const s = app.wb.sheet(); app.change('t', () => {{ app.history.whole(s); {cells} s.page = {json.dumps(js)}; }}); }})()")
            out = TMP / f"raabe-print-calc-{i}.pdf"
            c.js(f"__calc.pdf({json.dumps(str(out))})")
            time.sleep(0.5)
            results[name]["raabecalc"] = summary(out)
            out.unlink(missing_ok=True)
    finally:
        c.close()
    for name, r in results.items():
        e, o = r["excel"], r["raabecalc"]
        r["samePages"] = e["pages"] == o["pages"]
        print(f"{name:18} pages Excel {e['pages']} RaabeCalc {o['pages']}; lines per page Excel {[p['lines'] for p in e['detail']]} RaabeCalc {[p['lines'] for p in o['detail']]}")
    Path(__file__).parent.joinpath("results", f"{datetime.date.today().isoformat()}-calc-print.json").write_text(
        json.dumps(results, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
