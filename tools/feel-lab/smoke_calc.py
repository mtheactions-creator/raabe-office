"""Smoke test of RaabeCalc's later features together (D-065), checked by Excel: in the real app a
class list gets a Datenüberprüfung, Teilergebnisse (with their outline), a Symbolsatz, a picture,
Kopf- und Fußzeile, a manual page break, Blatt schützen and Arbeitsmappe schützen; the workbook
is saved to TEMP and opened in an Excel of our own (invisible), which must open it without repair
and see each feature. Result: results/<date>-calc-smoke.json.

    .venv/Scripts/python.exe smoke_calc.py
"""
from __future__ import annotations

import datetime
import json
import tempfile
import time
from pathlib import Path

from feellab.calc_catalogue import Case, CalcSide
from feellab.excel import ExcelTarget

PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
DATA = {"A1": "Klasse", "B1": "Name", "C1": "Note", "A2": "7a", "B2": "Anna", "C2": "2", "A3": "7a", "B3": "Ben", "C3": "3",
        "A4": "7b", "B4": "Cem", "C4": "1", "A5": "7b", "B5": "Dora", "C5": "4", "A6": "7b", "B6": "Emil", "C6": "2"}


def main() -> None:
    out = Path(tempfile.gettempdir()) / "raabe-calc-smoke.xlsx"
    out.unlink(missing_ok=True)
    res: dict = {"app": {}, "excel": {}}
    c = CalcSide()
    try:
        c.reset(Case("smoke", DATA, [], "x"))
        # Teilergebnis through its dialog (defaults: by Klasse, Summe of the last column).
        c.js("__calc.app.grid.selectCell(1, 0)")
        c.js("void __calc.app.commands.run('data.subtotal')")
        time.sleep(0.6)
        c.js("[...document.querySelectorAll('.dialog button')].find(b => b.textContent === 'OK').click()")
        time.sleep(0.6)
        res["app"]["rows"] = c.js("__calc.app.sheet.usedRange().r1 + 1")
        # Datenüberprüfung, Symbolsatz, picture, header/footer, page break, protection.
        c.js("""(() => {
          const app = __calc.app, s = app.sheet;
          app.change('smoke', () => {
            app.history.whole(s);
            s.validations = [{ ref: [{ r0: 1, c0: 2, r1: 20, c1: 2 }], type: 'whole', operator: 'between', formula1: '1', formula2: '6', allowBlank: true, showErrorMessage: true }];
            s.conditional = [...(s.conditional ?? []), { ref: [{ r0: 1, c0: 2, r1: 9, c1: 2 }], rules: [{ type: 'iconSet', priority: 1, iconSet: { set: '3Arrows', cfvo: [{ type: 'percent', val: '0' }, { type: 'percent', val: '33' }, { type: 'percent', val: '67' }] } }] }];
            s.page = { ...(s.page ?? {}), header: '&CNotenliste', footer: '&RSeite &P', rowBreaks: [4] };
            s.protection = { sheet: '1', objects: '1', scenarios: '1' };
          });
        })()""")
        c.js("__calc.app.pictures.insertPicture({ from: { col: 5, row: 1, x: 0, y: 0 }, to: { col: 7, row: 6, x: 0, y: 0 }, size: { w: 128, h: 100 }, src: 'data:image/png;base64,%s', name: 'Logo' })" % PNG)
        c.js("(() => { const app = __calc.app; app.change('book', () => { app.history.sheets(); app.wb.protection = { lockStructure: '1' }; }); })()")
        res["app"]["sheetGreyed"] = {k: not c.js(f"__calc.app.commands.isEnabled('{k}')") for k in ("data.subtotal", "data.group", "font.bold", "cells.insertSheet")}
        c.js(f"__calc.save({json.dumps(str(out))})")
        time.sleep(1.0)
    finally:
        c.close()
    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(out))
        ws = wb.Worksheets(1)
        e = res["excel"]
        e["opened"] = True
        e["labels"] = [ws.Cells(r, 1).Text for r in range(1, 10)]
        e["totals"] = [ws.Cells(r, 3).Formula for r in range(1, 10) if str(ws.Cells(r, 3).Formula).startswith("=")]
        e["values"] = [ws.Cells(r, 3).Value for r in range(1, 10) if str(ws.Cells(r, 3).Formula).startswith("=")]
        e["outline"] = [ws.Rows(r).OutlineLevel for r in range(1, 10)]
        e["validation"] = ws.Range("C2").Validation.Type
        e["iconSet"] = any(ws.Cells.FormatConditions(i).Type == 6 for i in range(1, ws.Cells.FormatConditions.Count + 1))
        e["pictures"] = [s.Name for s in ws.Shapes if s.Type == 13]
        e["header"] = ws.PageSetup.CenterHeader
        e["break"] = [r for r in range(2, 12) if ws.Rows(r).PageBreak == -4135]
        e["sheetProtected"] = bool(ws.ProtectContents)
        e["bookProtected"] = bool(wb.ProtectStructure)
        wb.Close(False)
    except Exception as err:  # noqa: BLE001
        res["excel"]["error"] = f"{err.__class__.__name__}: {err}"
    finally:
        x.close()
        out.unlink(missing_ok=True)
    e = res["excel"]
    res["pass"] = bool(e.get("opened") and e.get("validation") == 1 and e.get("iconSet") and e.get("pictures") == ["Logo"]
                       and e.get("sheetProtected") and e.get("bookProtected") and e.get("break") and len(e.get("totals", [])) == 3
                       and all(res["app"]["sheetGreyed"].values()))
    print(json.dumps(res, indent=1, ensure_ascii=False))
    Path(__file__).parent.joinpath("results", f"{datetime.date.today().isoformat()}-calc-smoke.json").write_text(json.dumps(res, indent=1, ensure_ascii=False), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
