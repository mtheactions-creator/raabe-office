"""Records how Excel draws the borders of cells with turned text (D-125): our own invisible Excel, COM only.

Cells with a thin border on all sides, a fill and text turned by -90 … 90 degrees (Ausrichtung), each in a row of
its own and in one row side by side (an angled header row); also a cell with only a left border, one with a thick
border, and the same with Zellen verbinden. The sheet goes to PDF (ExportAsFixedFormat) and every painted path is
read from the PDF (feellab/pdftext.paths): where the vertical borders run, whether the fill turns with them, and
whether the text stays inside. Written to results/2026-11-09-excel-rotated-borders.json.

    .venv/Scripts/python.exe probe_excel_rotated_borders.py
"""
from __future__ import annotations

import json
import tempfile
from pathlib import Path

from feellab import pdftext
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir())
OUT = HERE / "results" / "2026-11-09-excel-rotated-borders.json"
ANGLES = [45, 30, 60, 15, -45, -30, 90, -90, 0]


def main() -> None:
    x = ExcelTarget(visible=False)
    out: dict = {"cases": []}
    try:
        ws = x.sheet
        ws.Columns("A:L").ColumnWidth = 10
        layout = []
        # One cell per row (column B), row height 60 pt.
        for i, a in enumerate(ANGLES):
            r = 2 + 2 * i
            ws.Rows(r).RowHeight = 60
            layout.append({"case": f"single {a}", "cell": f"B{r}", "angle": a})
        # An angled header row: C..G with 45 degrees side by side, row 40 at 60 pt.
        ws.Rows(40).RowHeight = 60
        layout.append({"case": "row 45", "cell": "C40:G40", "angle": 45})
        ws.Rows(42).RowHeight = 60
        layout.append({"case": "row -45", "cell": "C42:G42", "angle": -45})
        ws.Rows(44).RowHeight = 60
        layout.append({"case": "left only 45", "cell": "B44", "angle": 45, "edges": [7]})
        ws.Rows(46).RowHeight = 60
        layout.append({"case": "thick 45", "cell": "B46", "angle": 45, "weight": -4138})
        ws.Rows(48).RowHeight = 60
        layout.append({"case": "nofill 45", "cell": "B48", "angle": 45, "nofill": True})
        ws.Rows(50).RowHeight = 30
        layout.append({"case": "low row 45", "cell": "B50", "angle": 45})
        layout.append({"case": "noborder 45", "cell": "B52", "angle": 45, "edges": []})
        for c in layout:
            rng = ws.Range(c["cell"])
            rng.Value2 = "Überschrift"
            rng.Orientation = c["angle"]
            if not c.get("nofill"):
                rng.Interior.Color = 0xCCE5FF
            for e in c.get("edges", [7, 8, 9, 10, 11] if ":" in c["cell"] else [7, 8, 9, 10]):
                rng.Borders(e).LineStyle = 1
                rng.Borders(e).Weight = c.get("weight", 2)
            cell = ws.Range(c["cell"].split(":")[0])
            c["box"] = {"left": cell.Left, "top": cell.Top, "width": rng.Width, "height": rng.Height}
        ps = ws.PageSetup
        ps.PaperSize = 9
        ps.LeftMargin = ps.TopMargin = ps.RightMargin = ps.BottomMargin = 36
        ps.CenterHorizontally = False
        pdf = TMP / "raabe-probe-rotated-borders.pdf"
        ps.PrintArea = "$A$1:$H$53"
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = 1
        ws.ExportAsFixedFormat(0, str(pdf))
        doc = pdftext.Pdf(pdf.read_bytes())
        pages = []
        for num in doc.pages():
            w, h = doc.page_box(num)
            content = doc.contents(num)
            pages.append({"w": w, "h": h, "paths": pdftext.paths(content, h),
                          "text": [{"x": round(r.x, 2), "y": round(r.y, 2), "size": r.size} for r in pdftext.text_runs(content, h)]})
        pdf.unlink(missing_ok=True)
        out["pages"] = pages
        out["layout"] = layout
        out["zoom"] = ps.Zoom
    finally:
        x.close()
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=0), encoding="utf-8", newline="\n")
    print("written", OUT, len(out.get("pages", [])), "pages", sum(len(p["paths"]) for p in out.get("pages", [])), "paths")


if __name__ == "__main__":
    main()
