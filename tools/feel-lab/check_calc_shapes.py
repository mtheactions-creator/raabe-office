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
    {"name": "roundrect-adj", "geom": "roundRect", "excel": ("shape", 5), "box": (10, 10, 3, 3), "adj": 30000},
    {"name": "arrow-both", "geom": "straightConnector1", "excel": ("line",), "box": (10, 1, 2, 4), "line": {"head": "oval", "tail": "stealth"}},
    {"name": "group", "geom": "group", "excel": ("group",), "box": (12, 13, 4, 3)},
    {"name": "freeform", "geom": "rect", "excel": ("freeform",), "box": (0, 14, 2, 3)},
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
    if "adj" in c:
        m["adjust"] = {"adj": c["adj"]}
    wpx, hpx = m["size"]["w"], m["size"]["h"]
    if c["excel"][0] == "group":
        kid = lambda geom, x: {"geom": geom, "from": m["from"], "to": m["to"], "rel": {"x": x, "y": 0, "w": wpx / 2, "h": hpx}}
        m["children"] = [kid("rect", 0), kid("ellipse", wpx / 2)]
        m["frame"] = {"w": wpx, "h": hpx}
    if c["excel"][0] == "freeform":
        W, H = round(wpx * 9525), round(hpx * 9525)
        m["custom"] = f"rect l t r b{chr(10)}path w={W} h={H}: M 0 {H} L {W // 2} 0 L {W} {H} Z"
    return m


def excel_view(sh) -> dict:
    """What Excel shows of a shape, in our terms."""
    out: dict = {"left": round(sh.Left, 1), "top": round(sh.Top, 1), "width": round(sh.Width, 1), "height": round(sh.Height, 1),
                 "rot": round(sh.Rotation, 1)}
    if sh.Type == 6:  # msoGroup
        out["geom"] = "group"
        out["items"] = sh.GroupItems.Count
        return out
    if sh.Type == 5:  # msoFreeform
        out["geom"] = "freeform"
        out["nodes"] = sh.Nodes.Count
        return out
    if sh.Connector or sh.Type == 9:  # msoLine
        out["geom"] = "line"
    elif sh.Type == 17:  # msoTextBox
        out["geom"] = "textbox"
    else:
        out["geom"] = GEOM.get(sh.AutoShapeType, sh.AutoShapeType)
    out["fill"] = None if not sh.Fill.Visible else sh.Fill.ForeColor.RGB
    out["line"] = None if not sh.Line.Visible else {"rgb": sh.Line.ForeColor.RGB, "w": round(sh.Line.Weight, 2), "dash": sh.Line.DashStyle,
                                                     "tail": sh.Line.EndArrowheadStyle, "head": sh.Line.BeginArrowheadStyle}
    out["flipV"] = bool(sh.VerticalFlip)
    try:
        out["adj"] = round(sh.Adjustments.Item(1), 3) if sh.Adjustments.Count else None
    except Exception:  # noqa: BLE001
        out["adj"] = None
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
    e["geom"] = {"line": "line", "textbox": "textbox", "group": "group", "freeform": "freeform"}.get(c["excel"][0], c["geom"])
    if c["excel"][0] == "group":
        e["items"] = 2
    if c["excel"][0] in ("group", "freeform"):
        return e
    if "fill" in c:
        e["fill"] = bgr(c["fill"])
    if c.get("nofill"):
        e["fill"] = None
    if "line" in c:
        ln = c["line"]
        e["line"] = {"rgb": bgr(ln["color"]) if "color" in ln else None, "w": float(ln["width"]) if "width" in ln else None,
                     "dash": DASH.get(ln.get("dash", ""), 1), "tail": ARROW.get(ln.get("tail", ""), 1), "head": ARROW.get(ln.get("head", ""), 1)}
    e["flipV"] = bool(c.get("flipV"))
    if "adj" in c:
        e["adj"] = round(c["adj"] / 100000, 3)
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
        if isinstance(v, dict) and isinstance(g, dict):
            bad += [f"{k}.{q}" for q, x in v.items() if x is not None and (abs(x - g.get(q, -99)) > 0.05 if isinstance(x, float) else x != g.get(q))]
        elif isinstance(v, float) and isinstance(g, (int, float)):
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
            elif kind == "group":
                a1 = ws.Shapes.AddShape(1, left, top, wid / 2, hei)
                a2 = ws.Shapes.AddShape(9, left + wid / 2, top, wid / 2, hei)
                sh = ws.Shapes.Range([a1.Name, a2.Name]).Group()
            elif kind == "freeform":
                fb = ws.Shapes.BuildFreeform(0, left, top + hei)
                fb.AddNodes(0, 0, left + wid / 2, top)
                fb.AddNodes(0, 0, left + wid, top + hei)
                fb.AddNodes(0, 0, left, top + hei)
                sh = fb.ConvertToShape()
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
                if "color" in ln:
                    sh.Line.ForeColor.RGB = bgr(ln["color"])
                if "width" in ln:
                    sh.Line.Weight = ln["width"]
                if "dash" in ln:
                    sh.Line.DashStyle = DASH[ln["dash"]]
                if "tail" in ln:
                    sh.Line.EndArrowheadStyle = ARROW[ln["tail"]]
                if "head" in ln:
                    sh.Line.BeginArrowheadStyle = ARROW[ln["head"]]
            if "text" in k:
                sh.TextFrame2.TextRange.Text = k["text"].replace("\n", "\r")
                if k.get("bold"):
                    sh.TextFrame2.TextRange.Font.Bold = True
                if k.get("size"):
                    sh.TextFrame2.TextRange.Font.Size = k["size"]
            if "rot" in k:
                sh.Rotation = k["rot"]
            if "adj" in k:
                sh.Adjustments.SetItem(1, k["adj"] / 100000)
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
            if k["excel"][0] == "group" and [c.get("geom") for c in g.get("children") or []] != ["rect", "ellipse"]:
                bad.append("children")
            if k["excel"][0] == "freeform" and not g.get("custom"):
                bad.append("custom")
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
            if "adjust" in want and any(abs((g.get("adjust") or {}).get(q, -1) - v) > 50 for q, v in want["adjust"].items()):
                bad.append("adjust")
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
