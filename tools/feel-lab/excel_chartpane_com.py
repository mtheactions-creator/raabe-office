"""What Excel's COM reports for a chart setting of RaabeCalc's Format pane (D-111): for the parts of
a chart model that differ from the unchanged chart, the Excel object model's properties and the
values they must have when Excel opened our file. Used by check_calc_chartpane.py."""
from __future__ import annotations

import json

DASH = {"solid": 1, "sysDash": 10, "sysDot": 11, "dash": 4, "dashDot": 5, "lgDash": 7, "lgDashDot": 8, "lgDashDotDot": 9}
TICK = {"none": -4142, "in": 2, "out": 3, "cross": 4}
LBLPOS = {"nextTo": 4, "high": -4127, "low": -4134, "none": -4142}
DLPOS = {"ctr": -4108, "inEnd": 3, "inBase": 4, "outEnd": 2, "bestFit": 5, "l": -4131, "r": -4152, "t": 0, "b": 1}
MARKER = {"square": 1, "diamond": 2, "triangle": 3, "star": 5, "circle": 8, "plus": 9, "x": -4168, "dot": -4118, "dash": -4115, "none": -4142}
COMPOUND = {"sng": 1, "dbl": 2, "thinThick": 3, "thickThin": 4, "tri": 5}
ARROW = {"triangle": 2, "arrow": 3, "stealth": 4, "diamond": 5, "oval": 6}
UNITS = {"hundreds": -2, "thousands": -3, "tenThousands": -4, "hundredThousands": -5, "millions": -6, "tenMillions": -7, "hundredMillions": -8, "billions": -9, "trillions": -10}
LEGEND = {"t": -4160, "b": -4107, "l": -4131, "r": -4152, "tr": 2}
FILL_TYPE = {"gradient": 3, "pattern": 2}


def rgb(hex_: str) -> int:
    r, g, b = int(hex_[1:3], 16), int(hex_[3:5], 16), int(hex_[5:7], 16)
    return r + (g << 8) + (b << 16)


def differs(a, b) -> bool:
    return json.dumps(a, sort_keys=True) != json.dumps(b, sort_keys=True)


def shape_checks(obj: str, new: dict | None, old: dict | None, line_obj: str | None = None) -> list[tuple[str, object]]:
    """Füllung, Linie and Effekte of an element (Format.Fill, Format.Line, Format.Shadow …)."""
    new, old = new or {}, old or {}
    out: list[tuple[str, object]] = []
    f = new.get("fill")
    if differs(f, old.get("fill")) and f:
        t = f.get("type")
        if t == "none":
            out.append((f"{obj}.Format.Fill.Visible", 0))
        elif t == "solid":
            out.append((f"{obj}.Format.Fill.ForeColor.RGB", rgb(f["color"])))
            if f.get("transparency"):
                out.append((f"{obj}.Format.Fill.Transparency", round(f["transparency"] / 100, 2)))
        elif t in FILL_TYPE:
            out.append((f"{obj}.Format.Fill.Type", FILL_TYPE[t]))
            if t == "gradient":
                out.append((f"{obj}.Format.Fill.GradientStops.Count", len(f["gradient"]["stops"])))
            if t == "pattern":
                out.append((f"{obj}.Format.Fill.ForeColor.RGB", rgb(f["pattern"]["fg"])))
        elif t == "picture" and f.get("picture"):
            out.append((f"{obj}.Format.Fill.Type in (4, 6)", True))
    lo = line_obj or f"{obj}.Format.Line"
    ln = new.get("line")
    if differs(ln, old.get("line")) and ln:
        t = ln.get("type")
        if t == "none":
            out.append((f"{lo}.Visible", 0))
        else:
            out.append((f"{lo}.Visible", -1))
            if t == "solid" and ln.get("color"):
                out.append((f"{lo}.ForeColor.RGB", rgb(ln["color"])))
            if ln.get("width") is not None:
                out.append((f"{lo}.Weight", ln["width"]))
            if ln.get("dash"):
                out.append((f"{lo}.DashStyle", DASH.get(ln["dash"], 1)))
            if ln.get("compound"):
                out.append((f"{lo}.Style", COMPOUND[ln["compound"]]))
            if ln.get("head"):
                out.append((f"{lo}.BeginArrowheadStyle", ARROW[ln["head"]["type"]]))
            if ln.get("tail"):
                out.append((f"{lo}.EndArrowheadStyle", ARROW[ln["tail"]["type"]]))
    fx, ofx = new.get("effects") or {}, old.get("effects") or {}
    if differs(fx.get("shadow"), ofx.get("shadow")):
        out.append((f"{obj}.Format.Shadow.Visible", -1 if fx.get("shadow") else 0))
    if differs(fx.get("glow"), ofx.get("glow")) and fx.get("glow"):
        out.append((f"{obj}.Format.Glow.Radius", fx["glow"]["size"]))
    if differs(fx.get("softEdges"), ofx.get("softEdges")) and fx.get("softEdges"):
        out.append((f"{obj}.Format.SoftEdge.Radius", fx["softEdges"]))
    if differs(fx.get("format3d"), ofx.get("format3d")) and (fx.get("format3d") or {}).get("top"):
        out.append((f"{obj}.Format.ThreeD.BevelTopType != 1", True))
    return out


def text_checks(obj: str, new: dict | None, old: dict | None) -> list[tuple[str, object]]:
    new, old = new or {}, old or {}
    font = f"{obj}.Format.TextFrame2.TextRange.Font"
    out: list[tuple[str, object]] = []
    for k, prop, conv in [("size", "Size", lambda v: v), ("bold", "Bold", lambda v: -1 if v else 0), ("italic", "Italic", lambda v: -1 if v else 0)]:
        if differs(new.get(k), old.get(k)) and new.get(k) is not None:
            out.append((f"{font}.{prop}", conv(new[k])))
    if differs(new.get("color"), old.get("color")) and new.get("color"):
        out.append((f"{font}.Fill.ForeColor.RGB", rgb(new["color"])))
    if differs(new.get("underline"), old.get("underline")) and new.get("underline"):
        out.append((f"{font}.UnderlineStyle != 0", True))
    if differs(new.get("strike"), old.get("strike")) and new.get("strike"):
        out.append((f"{font}.Strike != 0", True))
    if differs(new.get("outline"), old.get("outline")) and new.get("outline"):
        out.append((f"{font}.Line.Visible", 0 if new["outline"]["type"] == "none" else -1))
    if differs(new.get("fill"), old.get("fill")) and (new.get("fill") or {}).get("type") in FILL_TYPE:
        out.append((f"{font}.Fill.Type", FILL_TYPE[new["fill"]["type"]]))
    fx, ofx = new.get("effects") or {}, old.get("effects") or {}
    if differs(fx.get("glow"), ofx.get("glow")) and fx.get("glow"):
        out.append((f"{font}.Glow.Radius", fx["glow"]["size"]))
    if differs(fx.get("shadow"), ofx.get("shadow")) and fx.get("shadow"):
        out.append((f"{font}.Shadow.Visible", -1))
    return out


def axis_checks(ax: str, new: dict | None, old: dict | None, other: str) -> list[tuple[str, object]]:
    new, old = new or {}, old or {}
    out: list[tuple[str, object]] = []
    ch = lambda k: differs(new.get(k), old.get(k))  # noqa: E731
    for k, prop in [("min", "MinimumScale"), ("max", "MaximumScale"), ("majorUnit", "MajorUnit"), ("minorUnit", "MinorUnit")]:
        if ch(k) and new.get(k) is not None:
            out.append((f"{ax}.{prop}", new[k]))
    if ch("logBase"):
        out.append((f"{ax}.ScaleType", -4133 if new.get("logBase") else -4132))
        if new.get("logBase"):
            out.append((f"{ax}.LogBase", new["logBase"]))
    if ch("reverse"):
        out.append((f"{ax}.ReversePlotOrder", bool(new.get("reverse"))))
    if ch("majorTick"):
        out.append((f"{ax}.MajorTickMark", TICK[new.get("majorTick") or "none"]))
    if ch("minorTick"):
        out.append((f"{ax}.MinorTickMark", TICK[new.get("minorTick") or "none"]))
    if ch("labelPos"):
        out.append((f"{ax}.TickLabelPosition", LBLPOS[new.get("labelPos") or "nextTo"]))
    if ch("numFmt") and new.get("numFmt"):
        out.append((f"{ax}.TickLabels.NumberFormat", new["numFmt"]))
    if ch("crosses"):
        v = new.get("crosses")
        if isinstance(v, (int, float)):
            out.append((f"{ax}.CrossesAt", v))
        elif v == "max":
            out.append((f"{ax}.Crosses", 2))
        else:
            out.append((f"{ax}.Crosses", -4105))
    if ch("dispUnits"):
        v = new.get("dispUnits")
        out.append((f"{ax}.DisplayUnit", UNITS.get(v, -4114) if v is not None else -4142))
    if ch("dispUnitsLabel") and new.get("dispUnits") is not None:
        out.append((f"{ax}.HasDisplayUnitLabel", new.get("dispUnitsLabel") is not False))
    if ch("tickLblSkip") and new.get("tickLblSkip"):
        out.append((f"{ax}.TickLabelSpacing", new["tickLblSkip"]))
    if ch("tickMarkSkip") and new.get("tickMarkSkip"):
        out.append((f"{ax}.TickMarkSpacing", new["tickMarkSkip"]))
    if ch("lblOffset"):
        out.append((f"{ax}.TickLabels.Offset", new.get("lblOffset", 100)))
    if ch("onTicks"):
        out.append((f"{ax}.AxisBetweenCategories", not new.get("onTicks")))
    if ch("axisType"):
        out.append((f"{ax}.CategoryType", {"text": 2, "date": 3}.get(new.get("axisType"), -4105)))
    if ch("visible"):
        out.append((f"{other}.HasAxis({ax.split('(')[1].rstrip(')')})", new.get("visible") is not False))
    out += shape_checks(ax, new, old)
    t, ot = new.get("text") or {}, old.get("text") or {}
    out += text_checks(f"{ax}.TickLabels" if False else ax, t, ot)
    if differs((t.get("box") or {}).get("rotation"), (ot.get("box") or {}).get("rotation")) and (t.get("box") or {}).get("rotation") is not None:
        out.append((f"{ax}.TickLabels.Orientation", -t["box"]["rotation"]))
    return out


def model_checks(base: dict, new: dict) -> list[tuple[str, object]]:
    out: list[tuple[str, object]] = []
    f, of = new.get("format") or {}, base.get("format") or {}
    for k, obj in [("chart", "ch.ChartArea"), ("plot", "ch.PlotArea"), ("title", "ch.ChartTitle"), ("legend", "ch.Legend")]:
        out += shape_checks(obj, f.get(k), of.get(k))
        out += text_checks(obj, (f.get(k) or {}).get("text"), (of.get(k) or {}).get("text"))
    if differs((f.get("chart") or {}).get("roundedCorners"), (of.get("chart") or {}).get("roundedCorners")):
        out.append(("ch.ChartArea.RoundedCorners", bool((f.get("chart") or {}).get("roundedCorners"))))
    if differs((f.get("legend") or {}).get("overlay"), (of.get("legend") or {}).get("overlay")):
        out.append(("ch.Legend.IncludeInLayout", not (f.get("legend") or {}).get("overlay")))
    if differs((f.get("title") or {}).get("overlay"), (of.get("title") or {}).get("overlay")):
        out.append(("ch.ChartTitle.IncludeInLayout", not (f.get("title") or {}).get("overlay")))
    out += axis_checks("ch.Axes(2)", f.get("valAxis"), of.get("valAxis"), "ch")
    out += axis_checks("ch.Axes(1)", f.get("catAxis"), of.get("catAxis"), "ch")
    out += axis_checks("ch.Axes(2, 2)", f.get("valAxis2"), of.get("valAxis2"), "ch")
    out += shape_checks("ch.Axes(2).MajorGridlines", f.get("majorGrid"), of.get("majorGrid"))
    if (f.get("minorGrid") or {}).get("show"):
        out += shape_checks("ch.Axes(2).MinorGridlines", f.get("minorGrid"), of.get("minorGrid"))
    if (f.get("catGrid") or {}).get("show"):
        out += shape_checks("ch.Axes(1).MajorGridlines", f.get("catGrid"), of.get("catGrid"))
    for k, n in [("cat", 1), ("val", 2)]:
        at, oat = (f.get("axisTitles") or {}).get(k), (of.get("axisTitles") or {}).get(k)
        out += shape_checks(f"ch.Axes({n}).AxisTitle", at, oat)
        out += text_checks(f"ch.Axes({n}).AxisTitle", (at or {}).get("text"), (oat or {}).get("text"))
    dl, odl = f.get("dataLabels") or {}, of.get("dataLabels") or {}
    if differs(dl, odl) and new.get("dataLabels"):
        d = "ch.SeriesCollection(1).DataLabels()"
        for k, prop in [("showVal", "ShowValue"), ("showCat", "ShowCategoryName"), ("showSer", "ShowSeriesName"), ("showPercent", "ShowPercentage"), ("showLegendKey", "ShowLegendKey"), ("showBubble", "ShowBubbleSize")]:
            if differs(dl.get(k), odl.get(k)):
                out.append((f"{d}.{prop}", dl.get(k) if k != "showVal" else dl.get(k) is not False))
        if differs(dl.get("pos"), odl.get("pos")) and dl.get("pos"):
            out.append((f"{d}.Position", DLPOS[dl["pos"]]))
        if differs(dl.get("separator"), odl.get("separator")) and dl.get("separator"):
            out.append((f"{d}.Separator", dl["separator"]))
        if differs(dl.get("numFmt"), odl.get("numFmt")) and dl.get("numFmt"):
            out.append((f"{d}.NumberFormat", dl["numFmt"]))
        if differs(dl.get("fromCells"), odl.get("fromCells")):
            out.append((f"{d}.ShowRange", bool(dl.get("fromCells"))))
        out += shape_checks(d, dl, odl)
        out += text_checks(d, dl.get("text"), odl.get("text"))
    for k, prop in [("gapWidth", "GapWidth"), ("overlap", "Overlap"), ("firstSliceAngle", "FirstSliceAngle"), ("holeSize", "DoughnutHoleSize")]:
        if differs(new.get(k), base.get(k)) and new.get(k) is not None:
            out.append((f"ch.ChartGroups(1).{prop}", new[k]))
    if differs(new.get("varyColors"), base.get("varyColors")):
        out.append(("ch.ChartGroups(1).VaryByCategories", bool(new.get("varyColors"))))
    if differs(new.get("legend"), base.get("legend")) and new.get("legend") in LEGEND:
        out.append(("ch.Legend.Position", LEGEND[new["legend"]]))
    for i, (s, os_) in enumerate(zip(new.get("series", []), base.get("series", [])), start=1):
        sc = f"ch.SeriesCollection({i})"
        out += shape_checks(sc, {k: s.get(k) for k in ("fill", "line", "effects")}, {k: os_.get(k) for k in ("fill", "line", "effects")})
        for k, prop, conv in [("smooth", "Smooth", bool), ("invertIfNegative", "InvertIfNegative", bool), ("explosion", "Explosion", lambda v: v or 0)]:
            if differs(s.get(k), os_.get(k)):
                out.append((f"{sc}.{prop}", conv(s.get(k))))
        if differs(s.get("secondary"), os_.get("secondary")):
            out.append((f"{sc}.AxisGroup", 2 if s.get("secondary") else 1))
        mk, omk = s.get("marker") or {}, os_.get("marker") or {}
        if differs(mk.get("symbol"), omk.get("symbol")) and mk.get("symbol") and mk["symbol"] != "auto":
            out.append((f"{sc}.MarkerStyle", MARKER[mk["symbol"]]))
        if differs(mk.get("size"), omk.get("size")) and mk.get("size"):
            out.append((f"{sc}.MarkerSize", mk["size"]))
        if differs(mk.get("fill"), omk.get("fill")) and (mk.get("fill") or {}).get("type") == "solid":
            out.append((f"{sc}.MarkerBackgroundColor", rgb(mk["fill"]["color"])))
        if differs(mk.get("line"), omk.get("line")) and (mk.get("line") or {}).get("type") == "solid":
            out.append((f"{sc}.MarkerForegroundColor", rgb(mk["line"]["color"])))
        opts = {p["idx"]: p for p in os_.get("points") or []}
        for p in s.get("points") or []:
            q = opts.get(p["idx"], {})
            pt = f"{sc}.Points({p['idx'] + 1})"
            out += shape_checks(pt, p, q)
            if differs(p.get("explosion"), q.get("explosion")):
                out.append((f"{pt}.Explosion", p.get("explosion") or 0))
        tl, otl = s.get("trendline") or {}, os_.get("trendline") or {}
        if differs(tl, otl) and tl:
            t = f"{sc}.Trendlines(1)"
            for k, prop in [("order", "Order"), ("period", "Period"), ("forward", "Forward"), ("backward", "Backward"), ("intercept", "Intercept"), ("name", "Name")]:
                if differs(tl.get(k), otl.get(k)) and tl.get(k) is not None:
                    out.append((f"{t}.{prop}", tl[k]))
            for k, prop in [("showEq", "DisplayEquation"), ("showR2", "DisplayRSquared")]:
                if differs(tl.get(k), otl.get(k)):
                    out.append((f"{t}.{prop}", bool(tl.get(k))))
            out += shape_checks(t, tl, otl, f"{t}.Format.Line")
        eb, oeb = s.get("errorBars") or {}, os_.get("errorBars") or {}
        if differs(eb, oeb) and eb:
            e = f"{sc}.ErrorBars"
            if differs(eb.get("noEndCap"), oeb.get("noEndCap")):
                out.append((f"{e}.EndStyle", 2 if eb.get("noEndCap") else 1))
            out += shape_checks(e, eb, oeb)
    pr, opr = new.get("props") or {}, base.get("props") or {}
    if differs(pr.get("placement"), opr.get("placement")):
        out.append(("co.Placement", {"twoCell": 1, "oneCell": 2, "absolute": 3}[pr.get("placement") or "twoCell"]))
    if differs(pr.get("print"), opr.get("print")):
        out.append(("co.PrintObject", pr.get("print") is not False))
    if differs(pr.get("locked"), opr.get("locked")):
        out.append(("co.Locked", pr.get("locked") is not False))
    if differs(pr.get("altText"), opr.get("altText")):
        out.append(("co.ShapeRange.AlternativeText", pr.get("altText") or ""))
    if differs(pr.get("altTitle"), opr.get("altTitle")):
        out.append(("co.ShapeRange.Title", pr.get("altTitle") or ""))
    if differs(pr.get("lockAspect"), opr.get("lockAspect")):
        out.append(("co.ShapeRange.LockAspectRatio", -1 if pr.get("lockAspect") else 0))
    return out


def close(a, b) -> bool:
    if isinstance(a, bool) or isinstance(b, bool):
        return bool(a) == bool(b) if isinstance(a, (bool, int)) and isinstance(b, (bool, int)) else a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) < 0.02
    return a == b


def com_checks(ch, base: dict, new: dict) -> list[dict]:
    co = ch.Parent  # noqa: F841 (used by eval)
    out = []
    for expr, want in model_checks(base, new):
        try:
            got = eval(expr)  # noqa: S307 (our own expressions)
        except Exception as e:  # noqa: BLE001
            got = f"error {str(e)[:60]}"
        if isinstance(want, bool) and isinstance(got, int) and not isinstance(got, bool):
            got = bool(got)
        out.append({"what": expr, "want": want, "got": got if isinstance(got, (int, float, str, bool)) or got is None else str(got), "ok": close(got, want)})
    return out
