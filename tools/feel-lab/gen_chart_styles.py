"""Turns the charts Excel saved in each Diagrammformatvorlage (probe_excel_chart_filter.py: the gallery's entry n applied
through the ribbon to a column, line, bar and pie chart; fixtures in packages/ooxml/test/fixtures/excel-chart-filter/
style-<kind>-<n>.xml) into apps/raabe-calc/src/charts/chart-styles.gen.ts (D-118): per style what it sets on each
element (fills, lines, text, which elements show, data labels, gap width), colours as theme references.

    python gen_chart_styles.py
"""
from __future__ import annotations

import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
FIX = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "excel-chart-filter"
OUT = ROOT / "apps" / "raabe-calc" / "src" / "charts" / "chart-styles.gen.ts"
RES = HERE / "results" / "2026-11-09-excel-chart-filter.json"
NS = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart", "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def q(tag: str) -> str:
    p, t = tag.split(":")
    return f"{{{NS[p]}}}{t}"


def color(el) -> dict | None:
    """A colour element's reference: {"s": scheme} or {"rgb"}, with lumMod/lumOff/alpha/shade/tint."""
    if el is None:
        return None
    for c in el:
        name = c.tag.split("}")[1]
        if name not in ("schemeClr", "srgbClr", "prstClr", "sysClr"):
            continue
        out: dict = {}
        v = c.get("val")
        if name == "schemeClr":
            out["s"] = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}.get(v, v)
        elif name == "srgbClr":
            out["rgb"] = "#" + v.upper()
        elif name == "sysClr":
            out["rgb"] = "#" + (c.get("lastClr") or "000000").upper()
        else:
            out["rgb"] = {"white": "#FFFFFF", "black": "#000000"}.get(v, "#000000")
        for m in c:
            mn = m.tag.split("}")[1]
            if mn in ("lumMod", "lumOff", "alpha", "shade", "tint"):
                out[mn] = int(m.get("val")) / 100000
        return out
    return None


def fill(sp) -> dict | None:
    if sp is None:
        return None
    for f in sp:
        n = f.tag.split("}")[1]
        if n == "noFill":
            return {"none": True}
        if n == "solidFill":
            return {"c": color(f)}
        if n == "gradFill":
            gs = f.find("a:gsLst", NS)
            stops = list(gs) if gs is not None else []
            return {"c": color(stops[0]) if stops else None, "approx": "grad"}
        if n == "pattFill":
            fg = f.find("a:fgClr", NS)
            return {"c": color(fg), "approx": "patt"}
    return None


def line(sp) -> dict | None:
    if sp is None:
        return None
    ln = sp.find("a:ln", NS)
    if ln is None:
        return None
    out: dict = {}
    if ln.get("w"):
        out["w"] = round(int(ln.get("w")) / 12700, 2)
    if ln.find("a:noFill", NS) is not None:
        out["none"] = True
    sf = ln.find("a:solidFill", NS)
    if sf is not None:
        out["c"] = color(sf)
    dash = ln.find("a:prstDash", NS)
    if dash is not None and dash.get("val") != "solid":
        out["dash"] = dash.get("val")
    return out or None


def text(tx) -> dict | None:
    if tx is None:
        return None
    d = tx.find(".//a:defRPr", NS)
    if d is None:
        return None
    out: dict = {}
    if d.get("sz"):
        out["size"] = int(d.get("sz")) / 100
    if d.get("b") == "1":
        out["bold"] = True
    sf = d.find("a:solidFill", NS)
    if sf is not None:
        out["c"] = color(sf)
    return out or None


def elem(el) -> dict:
    out = {}
    if el is None:
        return out
    sp = el.find("c:spPr", NS)
    f, ln = fill(sp), line(sp)
    t = text(el.find("c:txPr", NS))
    if f:
        out["fill"] = f
    if ln:
        out["line"] = ln
    if t:
        out["text"] = t
    return out


def axis(ax) -> dict:
    if ax is None:
        return {}
    out = elem(ax)
    if ax.find("c:delete", NS) is not None and ax.find("c:delete", NS).get("val") == "1":
        out["hidden"] = True
    mg = ax.find("c:majorGridlines", NS)
    if mg is not None:
        out["major"] = elem(mg) or {"on": True}
        out["major"]["on"] = True
    ng = ax.find("c:minorGridlines", NS)
    if ng is not None:
        out["minor"] = elem(ng) or {"on": True}
        out["minor"]["on"] = True
    tick = ax.find("c:majorTickMark", NS)
    if tick is not None and tick.get("val") != "none":
        out["tick"] = tick.get("val")
    return out


def style_of(path: Path) -> dict:
    root = ET.parse(path).getroot()
    chart = root.find("c:chart", NS)
    plot = chart.find("c:plotArea", NS)
    out: dict = {"chart": elem(root), "plot": elem(plot)}
    title = chart.find("c:title", NS)
    out["title"] = elem(title) if title is not None else {"off": True}
    legend = chart.find("c:legend", NS)
    if legend is None:
        out["legend"] = {"pos": "none"}
    else:
        out["legend"] = {"pos": legend.find("c:legendPos", NS).get("val"), **elem(legend)}
    group = next(g for g in plot if g.tag.endswith("Chart"))
    gap = group.find("c:gapWidth", NS)
    if gap is not None:
        out["gap"] = int(gap.get("val"))
    ov = group.find("c:overlap", NS)
    if ov is not None:
        out["overlap"] = int(ov.get("val"))
    out["cat"] = axis(plot.find("c:catAx", NS)) if plot.find("c:catAx", NS) is not None else {}
    out["val"] = axis(plot.find("c:valAx", NS)) if plot.find("c:valAx", NS) is not None else {}
    sers = group.findall("c:ser", NS)
    series = []
    for i, s in enumerate(sers):
        e = elem(s)
        mk = s.find("c:marker", NS)
        if mk is not None:
            sym = mk.find("c:symbol", NS)
            size = mk.find("c:size", NS)
            m = {"symbol": sym.get("val") if sym is not None else None}
            if size is not None:
                m["size"] = int(size.get("val"))
            m.update(elem(mk))
            e["marker"] = m
        series.append(e)
    out["series"] = series
    dl = sers[0].find("c:dLbls", NS) if sers else None
    if dl is not None and dl.find("c:showVal", NS) is not None and dl.find("c:showVal", NS).get("val") == "1":
        lab = elem(dl)
        pos = dl.find("c:dLblPos", NS)
        if pos is not None:
            lab["pos"] = pos.get("val")
        for k in ("showCatName", "showPercent", "showSerName"):
            el = dl.find(f"c:{k}", NS)
            if el is not None and el.get("val") == "1":
                lab[k] = True
        out["labels"] = lab
    pt = sers[0].findall("c:dPt", NS) if sers else []
    if pt:
        out["points"] = [elem(p) for p in pt]
    return out


def main() -> None:
    res = json.loads(RES.read_text(encoding="utf-8"))
    styles: dict = {}
    for kind in ("column", "bar", "line", "pie"):
        rows = []
        for n in range(1, 15):
            case = res["cases"].get(f"style-{kind}-{n}")
            if not case or case.get("error"):
                continue
            st = style_of(FIX / f"style-{kind}-{n}.xml")
            st["id"] = case["style"]
            rows.append(st)
        styles[kind] = rows
    OUT.write_text(
        "// Generated by tools/feel-lab/gen_chart_styles.py from the charts German Excel 365 saved in each entry of the\n"
        "// Diagrammformatvorlagen gallery (probe_excel_chart_filter.py, D-118). Do not edit.\n\n"
        "/* eslint-disable */\n"
        "export const CHART_STYLES: Record<string, any[]> = " + json.dumps(styles, ensure_ascii=False, separators=(",", ":")) + ";\n",
        encoding="utf-8", newline="\n")
    print("written", OUT.name, {k: len(v) for k, v in styles.items()})


if __name__ == "__main__":
    main()
