"""How Excel lays a sheet out on paper (D-094), recorded in our own invisible Excel (DispatchEx); no real input.

Part "metrics": for several standard fonts (the Standard cell style), columns of known widths and rows of known
heights framed with thin borders are exported as PDF (ExportAsFixedFormat); the border lines' positions in the PDF
give each column's printed width and each row's printed height in points, next to the widths Excel shows on screen.

Part "pages": for each workbook of the look corpus (TEMP/raabe-corpus/excel), Excel's automatic and manual page
breaks of the first sheet, its used range and the cells under its drawing objects, so RaabeCalc's pages can be
compared row by row.

Part "rows": the printed height of a standard row for more fonts and sizes. Part "fontrows": rows Excel made
taller for a larger font, on screen and on paper.

    .venv/Scripts/python.exe probe_excel_print.py [metrics] [rows] [fontrows] [pages]

Writes results/<date>-excel-print.json (no paths, no names)."""
from __future__ import annotations

import datetime
import json
import sys
import tempfile
from pathlib import Path

from feellab import pdftext
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir())
FONTS = [("Aptos Narrow", 11), ("Calibri", 11), ("Arial", 10), ("Times New Roman", 12), ("Calibri", 10), ("Verdana", 10), ("Aptos Narrow", 12)]
WIDTHS = [0.5, 1, 2, 5, 8.43, 10, 10.71, 12, 15, 20, 25]
HEIGHTS = [None, 8, 10, 12, 14.5, 15, 15.75, 18, 20, 24, 30, 45, 60]


def lines(pdf: Path) -> tuple[list[float], list[float], list[dict]]:
    """x of vertical and y of horizontal line pieces (centre lines, points from the top left)."""
    data = pdftext.Pdf(pdf.read_bytes())
    num = data.pages()[0]
    _w, h = data.page_box(num)
    ps = pdftext.paths(data.contents(num), h)
    xs, ys = [], []
    for p in ps:
        if p["paint"] != "stroke":
            continue
        pts = [s[1:3] for s in p["d"] if s[0] in ("M", "L")]
        if len(pts) < 2:
            continue
        x0, x1 = min(q[0] for q in pts), max(q[0] for q in pts)
        y0, y1 = min(q[1] for q in pts), max(q[1] for q in pts)
        if x1 - x0 < 1.5 and y1 - y0 > 3:
            xs.append(round((x0 + x1) / 2, 2))
        elif y1 - y0 < 1.5 and x1 - x0 > 3:
            ys.append(round((y0 + y1) / 2, 2))
    return sorted(set(xs)), sorted(set(ys)), ps[:6]


def metrics(xl) -> list[dict]:
    out = []
    for name, size in FONTS:
        wb = xl.app.Workbooks.Add()
        try:
            st = wb.Styles("Standard")
            st.Font.Name = name
            st.Font.Size = size
            ws = wb.Worksheets(1)
            for i, w in enumerate(WIDTHS):
                ws.Columns(i + 1).ColumnWidth = w
            for i, hh in enumerate(HEIGHTS):
                if hh is not None:
                    ws.Rows(i + 1).RowHeight = hh
            rng = ws.Range(ws.Cells(1, 1), ws.Cells(len(HEIGHTS), len(WIDTHS)))
            for e in (7, 8, 9, 10, 11, 12):
                rng.Borders(e).LineStyle = 1
                rng.Borders(e).Weight = 2
            ps = ws.PageSetup
            ps.PaperSize = 9  # A4 (the printer may not know A3: Excel then scales the page)
            ps.Orientation = 2
            ps.LeftMargin = ps.TopMargin = 36
            ps.RightMargin = ps.BottomMargin = 36
            pdf = TMP / "raabe-probe-print.pdf"
            ws.ExportAsFixedFormat(0, str(pdf))
            xs, ys, sample = lines(pdf)
            row = {"font": name, "size": size, "standardWidth": ws.StandardWidth,
                   "screen": {"cols": [ws.Columns(i + 1).Width for i in range(len(WIDTHS))], "rows": [ws.Rows(i + 1).Height for i in range(len(HEIGHTS))]},
                   "widths": WIDTHS, "heights": HEIGHTS, "pdfX": xs, "pdfY": ys,
                   "printedCols": [round(b - a, 2) for a, b in zip(xs, xs[1:])], "printedRows": [round(b - a, 2) for a, b in zip(ys, ys[1:])],
                   "sample": sample}
            pdf.unlink(missing_ok=True)
            out.append(row)
            print(name, size, "cols", row["printedCols"], "rows", row["printedRows"], flush=True)
        finally:
            wb.Close(False)
    return out


def default_rows(xl) -> list[dict]:
    """The printed height of a standard row (and of a 30-pt row) for more standard fonts."""
    out = []
    for name in ["Aptos Narrow", "Aptos", "Calibri", "Arial", "Times New Roman", "Cambria", "Segoe UI", "Verdana", "Tahoma", "Consolas", "Georgia", "Calibri Light"]:
        for size in [9, 10, 11, 12, 14]:
            wb = xl.app.Workbooks.Add()
            try:
                st = wb.Styles("Standard")
                st.Font.Name = name
                st.Font.Size = size
                ws = wb.Worksheets(1)
                ws.Rows(3).RowHeight = 30
                rng = ws.Range("A1:B4")
                for e in (7, 8, 9, 10, 11, 12):
                    rng.Borders(e).LineStyle = 1
                pdf = TMP / "raabe-probe-print.pdf"
                ws.ExportAsFixedFormat(0, str(pdf))
                xs, ys, _ = lines(pdf)
                pdf.unlink(missing_ok=True)
                row = {"font": name, "size": size, "screenDefault": ws.Rows(1).Height, "screenCol": ws.Columns(1).Width,
                       "printedRows": [round(b - a, 3) for a, b in zip(ys, ys[1:])], "printedCols": [round(b - a, 3) for a, b in zip(xs, xs[1:])]}
                out.append(row)
                print(row, flush=True)
            finally:
                wb.Close(False)
    return out


def font_rows(xl) -> list[dict]:
    """Rows Excel sizes for a larger font (standard font Aptos Narrow 11): their screen height and printed height."""
    cases = [("Aptos Narrow", 16), ("Aptos Narrow", 18), ("Aptos Narrow", 20), ("Verdana", 14), ("Verdana", 18), ("Segoe UI", 18),
             ("Times New Roman", 18), ("Calibri", 18), ("Georgia", 14), ("Consolas", 18)]
    wb = xl.app.Workbooks.Add()
    try:
        ws = wb.Worksheets(1)
        for i, (f, s) in enumerate(cases):
            c = ws.Cells(i + 1, 1)
            c.Value = f"{f} {s}"
            c.Font.Name = f
            c.Font.Size = s
        rng = ws.Range(ws.Cells(1, 1), ws.Cells(len(cases), 2))
        for e in (7, 8, 9, 10, 11, 12):
            rng.Borders(e).LineStyle = 1
        pdf = TMP / "raabe-probe-print.pdf"
        ws.ExportAsFixedFormat(0, str(pdf))
        _xs, ys, _ = lines(pdf)
        pdf.unlink(missing_ok=True)
        pr = [round(b - a, 2) for a, b in zip(ys, ys[1:])]
        out = [{"font": f, "size": s, "screen": ws.Rows(i + 1).Height, "printed": pr[i] if i < len(pr) else None} for i, (f, s) in enumerate(cases)]
        for r in out:
            print(r, flush=True)
        return out
    finally:
        wb.Close(False)


def pages(xl) -> dict:
    folder = TMP / "raabe-corpus" / "excel"
    out = {}
    for f in sorted(folder.glob("*.xlsx")):
        wb = xl.app.Workbooks.Open(str(f), ReadOnly=True)
        try:
            ws = wb.Worksheets(1)
            ur = ws.UsedRange
            shapes = []
            for i in range(1, ws.Shapes.Count + 1):
                s = ws.Shapes(i)
                shapes.append({"type": s.Type, "tl": s.TopLeftCell.GetAddress(False, False), "br": s.BottomRightCell.GetAddress(False, False),
                               "left": s.Left, "top": s.Top, "w": s.Width, "h": s.Height})
            ps = ws.PageSetup
            info = {"used": ur.GetAddress(False, False), "shapes": shapes, "printArea": ps.PrintArea,
                    "zoom": ps.Zoom, "pages": ps.Pages.Count,
                    "hBreaks": [(ws.HPageBreaks(i).Location.Row, ws.HPageBreaks(i).Type) for i in range(1, ws.HPageBreaks.Count + 1)],
                    "vBreaks": [(ws.VPageBreaks(i).Location.Column, ws.VPageBreaks(i).Type) for i in range(1, ws.VPageBreaks.Count + 1)],
                    "margins": [ps.LeftMargin, ps.RightMargin, ps.TopMargin, ps.BottomMargin, ps.HeaderMargin, ps.FooterMargin],
                    "center": [ps.CenterHorizontally, ps.CenterVertically], "gridlines": ps.PrintGridlines}
            out[f.stem] = info
            print(f.stem, info["pages"], info["used"], info["hBreaks"], info["vBreaks"], [(s["tl"], s["br"]) for s in shapes], flush=True)
        finally:
            wb.Close(False)
    return out


def main() -> None:
    parts = sys.argv[1:] or ["metrics", "rows", "fontrows", "pages"]
    path = HERE / "results" / f"{datetime.date.today().isoformat()}-excel-print.json"
    res = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    xl = ExcelTarget(visible=False)
    try:
        if "metrics" in parts:
            res["metrics"] = metrics(xl)
        if "rows" in parts:
            res["defaultRows"] = default_rows(xl)
        if "fontrows" in parts:
            res["fontRows"] = font_rows(xl)
        if "pages" in parts:
            res["pages"] = pages(xl)
    finally:
        xl.close()
    path.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
