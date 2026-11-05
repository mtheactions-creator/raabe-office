"""Shapes drawn in Excel the way a user draws them (D-066): our own visible Excel goes into draw mode
for a shape (CommandBars.ExecuteMso, as the Formen gallery does) and a real mouse drag, or a click,
draws it. Recorded for the current Office theme and for the Office 2013-2022 theme: the shape's
XML as Excel saves it (style references, text body), its size after a click, its line and fill.
Takes over mouse and keyboard: run only when the owner is not using the PC.
Result: results/<date>-excel-drawshape.json.
"""
from __future__ import annotations

import datetime
import json
import re
import tempfile
import time
import zipfile
from pathlib import Path

from feellab import winapi
from feellab.calc_catalogue import _pt_to_px
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir())
THEMES = {"office": None, "office2013": r"C:\Program Files\Microsoft Office\root\Document Themes 16\Office 2013 - 2022 Theme.thmx"}
# (mso id, drag from cell, to cell or None for a click)
SHAPES = [
    ("ShapeRectangle", "B2", "D6"),
    ("ShapeOval", "F2", None),
    ("ShapeStraightConnector", "B9", "E12"),
    ("ShapeStraightConnectorArrow", "F9", "H12"),
    ("TextBoxInsertExcel", "B15", "D18"),
    ("ShapeRoundedRectangle", "F15", None),
]


def screen(x: ExcelTarget, addr: str) -> tuple[float, float]:
    win = x.app.ActiveWindow
    r = x.sheet.Range(addr)
    return win.PointsToScreenPixelsX(0) + _pt_to_px(r.Left + 5), win.PointsToScreenPixelsY(0) + _pt_to_px(r.Top + 4)


def main() -> None:
    out: dict = {}
    for theme, path in THEMES.items():
        x = ExcelTarget(visible=True)
        try:
            if path:
                x.wb.ApplyTheme(path)
            x.app.ActiveWindow.Zoom = 100
            rec = []
            for mso, a, b in SHAPES:
                winapi.activate(x.hwnd)
                x.sheet.Range("A1").Select()
                time.sleep(0.2)
                try:
                    x.app.CommandBars.ExecuteMso(mso)
                except Exception as e:  # noqa: BLE001
                    rec.append({"mso": mso, "error": str(e)[:120]})
                    continue
                time.sleep(0.4)
                p = screen(x, a)
                if b:
                    q = screen(x, b)
                    winapi.drag(p[0], p[1], q[0], q[1])
                else:
                    winapi.click(p[0], p[1])
                time.sleep(0.4)
                winapi.press("esc")
                time.sleep(0.2)
                n = x.sheet.Shapes.Count
                sh = x.sheet.Shapes(n) if n else None
                one = {"mso": mso, "drag": bool(b)}
                if sh is not None:
                    one.update({"name": sh.Name, "w": round(sh.Width, 2), "h": round(sh.Height, 2), "lineW": sh.Line.Weight,
                                "lineRGB": f"{sh.Line.ForeColor.RGB:06X}", "fillVisible": bool(sh.Fill.Visible), "fillRGB": f"{sh.Fill.ForeColor.RGB:06X}"})
                rec.append(one)
            f = TMP / f"drawshape-{theme}.xlsx"
            x.wb.SaveAs(str(f), 51)
            with zipfile.ZipFile(f) as z:
                xml = z.read("xl/drawings/drawing1.xml").decode("utf-8")
            anchors = re.findall(r"<xdr:twoCellAnchor\b[\s\S]*?</xdr:twoCellAnchor>", xml)
            for one, ax in zip([r for r in rec if "name" in r], anchors):
                one["style"] = (re.search(r"<xdr:style>[\s\S]*?</xdr:style>", ax) or [""])[0]
                one["bodyPr"] = (re.search(r"<a:bodyPr\b[^>]*>", ax) or [""])[0]
                one["pPr"] = (re.search(r"<a:pPr\b[^>]*>", ax) or [""])[0]
                one["spPr"] = (re.search(r"<xdr:spPr>[\s\S]*?</xdr:spPr>", ax) or [""])[0]
                one["ext"] = (re.search(r"<a:ext cx=\"\d+\" cy=\"\d+\"/>", ax) or [""])[0]
            out[theme] = rec
        finally:
            x.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-drawshape.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    for t, rec in out.items():
        for r in rec:
            print(t, r.get("mso"), r.get("name"), r.get("w"), r.get("h"), r.get("lineW"), r.get("lineRGB"), r.get("error", ""))


if __name__ == "__main__":
    main()
