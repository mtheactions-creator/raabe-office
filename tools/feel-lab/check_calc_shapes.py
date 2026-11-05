"""Shapes and text boxes against Excel (D-066), both ways, without real input:
1. RaabeCalc (off screen) puts shapes on a sheet and saves; our own invisible Excel opens the file and
   must see each shape with the same kind, place, size, colours, line, text and rotation.
2. Excel draws the same shapes and saves; RaabeCalc opens the file and must read the same models.
Result: results/<date>-calc-shapes.json.
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
COL_PT, ROW_PT = 60.0, 14.5  # Excel 365 defaults (Aptos Narrow 11): 80 px columns, 14.5 pt rows

# Each case: our model, and how Excel makes it (msoAutoShapeType or a line).
CASES = [
    {"name": "rect-red", "geom": "rect", "excel": ("shape", 1), "box": (1, 2, 3, 4), "fill": "#FF0000", "line": {"color": "#0000FF", "width": 2}},
    {"name": "ellipse-text", "geom": "ellipse", "excel": ("shape", 9), "box": (4, 1, 2, 3), "fill": "#00B050", "text": "Hallo", "bold": True, "size": 14},
    {"name": "arrow-rot", "geom": "rightArrow", "excel": ("shape", 33), "box": (0, 7, 3, 2), "rot": 30},
    {"name": "roundrect-dash", "geom": "roundRect", "excel": ("shape", 5), "box": (5, 6, 2, 2), "line": {"color": "#FFC000", "width": 3, "dash": "dash"}},
    {"name": "triangle-flip", "geom": "triangle", "excel": ("shape", 7), "box": (8, 1, 2, 3), "flipV": True},
    {"name": "star-nofill", "geom": "star5", "excel": ("shape", 92), "box": (8, 6, 2, 3), "nofill": True},
    {"name": "line-arrow", "geom": "straightConnector1", "excel": ("line",), "box": (1, 11, 4, 2), "line": {"color": "#FF0000", "width": 1.5, "tail": "triangle"}},
    {"name": "textbox", "geom": "rect", "excel": ("textbox",), "box": (6, 10, 3, 2), "text": "Zwei\nZeilen", "textBox": True},
]
DASH = {"dash": 4, "sysDot": 3, "sysDash": 10}
ARROW = {"triangle": 2, "stealth": 4, "oval": 6, "arrow": 3, "diamond": 5}
GEOM = {1: "rect", 9: "ellipse", 33: "rightArrow", 5: "roundRect", 7: "triangle", 92: "star5"}


def bgr(hex_: str) -> int:
    r, g, b = int(hex_[1:3], 16), int(hex_[3:5], 16), int(hex_[5:7], 16)
    return r | (g << 8) | (b << 16)


def model(c: dict) -> dict:
    col, row, w, h = c["box"]
    m: dict = {"geom": c["geom"], "from": {"col": col, "row": row, "x": 0, "y": 0}, "to": {"col": col + w, "row": row + h, "x": 0, "y": 0},
               "size": {"w": round(w * COL_PT / 0.75), "h": round(h * ROW_PT / 0.75)}, "name": c["name"]}
    if c["excel"][0] == "line":
        m["connector"] = True
    if "fill" in c:
        m["fill"] = {"type": "solid", "color": c["fill"]}
    if c.get("nofill"):
        m["fill"] = {"type": "none"}
    if "line" in c:
        m["line"] = {"type": "solid", **c["line"]}
    if "text" in c:
        m["text"] = c["text"]
    if c.get("bold") or c.get("size"):
        m["font"] = {k: v for k, v in (("bold", c.get("bold")), ("size", c.get("size"))) if v}
    for k in ("rot", "flipV", "textBox"):
        if k in c:
            m[k] = c[k]
    return m


def excel_view(sh) -> dict:
    """What Excel shows of a shape, in our terms."""
    out: dict = {"left": round(sh.Left, 1), "top": round(sh.Top, 1), "width": round(sh.Width, 1), "height": round(sh.Height, 1),
                 "rot": round(sh.Rotation, 1)}
    if sh.Connector or sh.Type == 9:  # msoLine
        out["geom"] = "line"
    elif sh.Type == 17:  # msoTextBox
        out["geom"] = "textbox"
    else:
        out["geom"] = GEOM.get(sh.AutoShapeType, sh.AutoShapeType)
    out["fill"] = None if not sh.Fill.Visible else sh.Fill.ForeColor.RGB
    out["line"] = None if not sh.Line.Visible else {"rgb": sh.Line.ForeColor.RGB, "w": round(sh.Line.Weight, 2), "dash": sh.Line.DashStyle,
                                                     "tail": sh.Line.EndArrowheadStyle}
    out["flipV"] = bool(sh.VerticalFlip)
    try:
        tr = sh.TextFrame2.TextRange
        out["text"] = tr.Text.replace("\r", "\n")
        out["bold"] = bool(tr.Font.Bold) if tr.Text else None
        out["size"] = tr.Font.Size if tr.Text else None
    except Exception:  # noqa: BLE001
        out["text"] = ""
    return out


def expect(c: dict) -> dict:
    """What Excel must show for a case."""
    col, row, w, h = c["box"]
    e: dict = {"left": col * COL_PT, "top": row * ROW_PT, "width": w * COL_PT, "height": h * ROW_PT, "rot": float(c.get("rot", 0))}
    e["geom"] = {"line": "line", "textbox": "textbox"}.get(c["excel"][0], c["geom"])
    if "fill" in c:
        e["fill"] = bgr(c["fill"])
    if c.get("nofill"):
        e["fill"] = None
    if "line" in c:
        ln = c["line"]
        e["line"] = {"rgb": bgr(ln["color"]), "w": float(ln["width"]), "dash": DASH.get(ln.get("dash", ""), 1), "tail": ARROW.get(ln.get("tail", ""), 1)}
    e["flipV"] = bool(c.get("flipV"))
    e["text"] = c.get("text", "")
    if c.get("bold"):
        e["bold"] = True
    if c.get("size"):
        e["size"] = float(c["size"])
    return e


def same(e: dict, got: dict) -> list[str]:
    bad = []
    for k, v in e.items():
        g = got.get(k)
        if isinstance(v, float) and isinstance(g, (int, float)):
            if abs(v - g) > 1.0:
                bad.append(k)
        elif v != g:
            bad.append(k)
    return bad


def main() -> None:
    raabe_file, excel_file = TMP / "raabe-shapes.xlsx", TMP / "excel-shapes.xlsx"
    report: dict = {"raabe->excel": [], "excel->raabe": []}
    c = CalcSide()
    try:
        c.reset(Case("shapes", {"A1": "x"}, [], "x", sheets=1))
        c.js(f"(() => {{ __calc.app.sheet.shapes = {json.dumps([model(k) for k in CASES])}; }})()")
        c.js(f"__calc.save({json.dumps(str(raabe_file))})")
    finally:
        c.close()
    x = ExcelTarget(visible=False)
    try:
        wb = x.app.Workbooks.Open(str(raabe_file))
        ws = wb.Worksheets(1)
        seen = {ws.Shapes(i).Name: ws.Shapes(i) for i in range(1, ws.Shapes.Count + 1)}
        for k in CASES:
            sh = seen.get(k["name"])
            got = excel_view(sh) if sh is not None else None
            e = expect(k)
            bad = ["missing"] if got is None else same(e, got)
            report["raabe->excel"].append({"case": k["name"], "ok": not bad, **({"bad": bad, "excel": got, "want": e} if bad else {})})
            print("R->E", k["name"], not bad, bad, flush=True)
        wb.Close(False)
        # 2. Excel draws them.
        wb = x.app.Workbooks.Add()
        ws = wb.Worksheets(1)
        for k in CASES:
            col, row, w, h = k["box"]
            left, top, wid, hei = col * COL_PT, row * ROW_PT, w * COL_PT, h * ROW_PT
            kind = k["excel"][0]
            if kind == "shape":
                sh = ws.Shapes.AddShape(k["excel"][1], left, top, wid, hei)
            elif kind == "line":
                sh = ws.Shapes.AddConnector(1, left, top, left + wid, top + hei)
            else:
                sh = ws.Shapes.AddTextbox(1, left, top, wid, hei)
            sh.Name = k["name"]
            if "fill" in k:
                sh.Fill.Visible = True
                sh.Fill.Solid()
                sh.Fill.ForeColor.RGB = bgr(k["fill"])
            if k.get("nofill"):
                sh.Fill.Visible = False
            if "line" in k:
                ln = k["line"]
                sh.Line.Visible = True
                sh.Line.ForeColor.RGB = bgr(ln["color"])
                sh.Line.Weight = ln["width"]
                if "dash" in ln:
                    sh.Line.DashStyle = DASH[ln["dash"]]
                if "tail" in ln:
                    sh.Line.EndArrowheadStyle = ARROW[ln["tail"]]
            if "text" in k:
                sh.TextFrame2.TextRange.Text = k["text"].replace("\n", "\r")
                if k.get("bold"):
                    sh.TextFrame2.TextRange.Font.Bold = True
                if k.get("size"):
                    sh.TextFrame2.TextRange.Font.Size = k["size"]
            if "rot" in k:
                sh.Rotation = k["rot"]
            if k.get("flipV"):
                sh.Flip(1)  # msoFlipVertical
        wb.SaveAs(str(excel_file), 51)
        wb.Close(False)
    finally:
        x.close()
    c = CalcSide()
    try:
        c.js(f"__calc.open({json.dumps(str(excel_file))})")
        got = c.js("__calc.app.wb.sheets[0].shapes?.map((s) => (({ raw, rawKey, ...r }) => r)(s)) ?? []")
    finally:
        c.close()
    by_name = {g.get("name"): g for g in got}
    for k in CASES:
        g = by_name.get(k["name"])
        want = model(k)
        bad = []
        if g is None:
            bad = ["missing"]
        else:
            if g["geom"] != want["geom"] and not (k["excel"][0] == "line" and g["geom"] in ("straightConnector1", "line")):
                bad.append("geom")
            for key in ("from", "to"):
                if any(abs(g[key][q] - want[key][q]) > 1 for q in ("col", "row")) or any(abs(g[key][q] - want[key][q]) > 2 for q in ("x", "y")):
                    bad.append(key)
            if "fill" in want and (g.get("fill") or {}).get("type") != want["fill"]["type"] or (want.get("fill", {}).get("color") and (g.get("fill") or {}).get("color") != want["fill"]["color"]):
                bad.append("fill")
            if "line" in want:
                gl = g.get("line") or {}
                for q, v in want["line"].items():
                    if (abs(gl.get(q, -1) - v) > 0.05) if isinstance(v, (int, float)) else gl.get(q) != v:
                        bad.append("line." + q)
            if (g.get("text") or "") != want.get("text", ""):
                bad.append("text")
            for q in ("rot", "flipV", "textBox"):
                if q in want and g.get(q) != want[q]:
                    bad.append(q)
            if "font" in want and any((g.get("font") or {}).get(q) != v for q, v in want["font"].items()):
                bad.append("font")
        report["excel->raabe"].append({"case": k["name"], "ok": not bad, **({"bad": bad, "got": g, "want": want} if bad else {})})
        print("E->R", k["name"], not bad, bad, flush=True)
    (HERE / "results" / f"{datetime.date.today().isoformat()}-calc-shapes.json").write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    for key, v in report.items():
        print(key, sum(1 for r in v if r["ok"]), "/", len(v))


if __name__ == "__main__":
    main()
