"""RaabeCalc's templates opened in Excel (its own invisible instance): each is written by the app,
opened by Excel without a repair, and Excel's results of the formulas and the number of charts are
read. Result: results/<date>-calc-templates.json.
"""
from __future__ import annotations

import datetime
import json
import tempfile
from pathlib import Path

from feellab.calc_catalogue import CalcSide
from feellab.excel import ExcelTarget

TMP = Path(tempfile.gettempdir())


def main() -> None:
    out: dict = {}
    c = CalcSide()
    try:
        for t in ("notenliste", "notenspiegel", "messwerte"):
            c.js(f"__calc.template('{t}', {json.dumps(str(TMP / f'raabe-template-{t}.xlsx'))})")
    finally:
        c.close()
    x = ExcelTarget(visible=False)
    try:
        for t in ("notenliste", "notenspiegel", "messwerte"):
            p = TMP / f"raabe-template-{t}.xlsx"
            wb = x.app.Workbooks.Open(str(p), CorruptLoad=0)
            ws = wb.Worksheets(1)
            used = ws.UsedRange
            errors = [c.GetAddress(False, False) for c in used.Cells if isinstance(c.Value2, int) and c.Value2 < -2146826000]
            out[t] = {"sheet": ws.Name, "charts": ws.ChartObjects().Count, "used": used.GetAddress(False, False), "errors": errors,
                      "sample": {"E4": ws.Range("E4").Text, "B12": ws.Range("B12").Text, "G31": ws.Range("G31").Text}}
            wb.Close(False)
            p.unlink(missing_ok=True)
    finally:
        x.close()
    for k, v in out.items():
        print(k, v)
    Path(__file__).parent.joinpath("results", f"{datetime.date.today().isoformat()}-calc-templates.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
