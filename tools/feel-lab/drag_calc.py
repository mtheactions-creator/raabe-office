"""The drag test in RaabeCalc (D-080 C, D-096): every anchor dragged along feellab/dragtest.py's paths with real input,
checked for jumps, plus the Esc and Strg+Z drags as Excel does them (probe_excel_drag_esc.py: Esc ends a drag of
cells, borders or the fill handle normally, and puts a shape or chart back). Takes mouse and keyboard (the real-input lock).

    .venv/Scripts/python.exe drag_calc.py [anchor ...]     (names, or parts of them)
    .venv/Scripts/python.exe drag_calc.py --paths out,mods,zoom [anchor ...]

--paths runs D-127's paths instead (out of the window and back, the button let go outside, Umschalt/Strg/Alt mid-drag
as Excel does them (probe_excel_drag_mods.py), other zoom levels) into results/2026-11-09-drag-calc-paths.json.
RAABE_CALC_EXE points at another build. Result: results/<date>-drag-calc.json.
"""
from __future__ import annotations

import contextlib
import datetime
import json
import math
import sys
import time
from pathlib import Path

from feellab import dragtest, winapi
from feellab.catcher import catcher
from feellab.calc_catalogue import Case, CalcSide
from feellab.dragtest import Anchor, analyse, check_paths, merge_report, run_drag, run_families, select_paths, summary
from feellab.inputlock import real_input, touch

HERE = Path(__file__).parent
PATHS = [p.name for p in __import__("feellab.dragtest", fromlist=["paths"]).paths()]
CASE = Case("drag", {"A1": "Name", "B1": "Wert", "A2": "Anna", "B2": "12", "A3": "Ben", "B3": "7", "A4": "Cem", "B4": "9"}, [], "A1:D5", start="B3", sheets=4)

# Shared JS: the grid, its canvas rectangle, a cell rectangle in client px.
PRE = "const g = __calc.grid, c = g.canvas.getBoundingClientRect(); const R = (r) => { const b = g.rectOnCanvas(r); return { l: c.left + b.x, t: c.top + b.y, r: c.left + b.x + b.w, b: c.top + b.y + b.h, w: b.w, h: b.h }; };"


def js(body: str) -> str:
    return f"(() => {{ {PRE} {body} }})()"


# D-127: the view's scroll position (auto-scroll), the zoom, and scrolling a start point into the window.
SCROLL = "(__calc.grid.hscroll.scrollLeft + __calc.grid.vScrollEl.scrollTop)"
ZOOM = lambda z: f"__calc.app.setZoom({z})"
# The cell under the start point scrolled into view with some room around it (the scroll bars reach only the used range).
REVEAL = ("(() => { const g = __calc.grid, c = g.canvas.getBoundingClientRect(), hit = g.hitCell({x} - c.left, {y} - c.top);"
          " g.reveal(hit.row + 3, hit.col + 2); g.updateScrollbars(); g.invalidate(); __calc.app.grid.contentChanged(); })()")
WINDOW = (0, 0, 2000, 1700)  # room below and right of the window for the out-of-window paths


def constrain(dx: float, dy: float, v0: dict) -> dict:
    """Umschalt while moving: along the axis the pointer moved further on (Excel)."""
    return {"x": v0["x"] + dx, "y": v0["y"]} if abs(dx) >= abs(dy) else {"x": v0["x"], "y": v0["y"] + dy}


def corner(dx: float, dy: float, v0: dict) -> dict:
    """A corner handle: the corner follows the pointer, the opposite corner stays."""
    return {"x": v0["x"] + dx, "y": v0["y"] + dy, "l": v0["l"], "t": v0["t"]}


def corner_aspect(dx: float, dy: float, v0: dict) -> dict:
    """Umschalt at a corner: the proportions stay, the side that grows more decides (Excel)."""
    w, h = v0["x"] - v0["l"], v0["y"] - v0["t"]
    k = max((w + dx) / w, (h + dy) / h)
    return {"x": v0["l"] + w * k, "y": v0["t"] + h * k, "l": v0["l"], "t": v0["t"]}


def corner_centre(dx: float, dy: float, v0: dict) -> dict:
    """Strg at a corner: the size changes around the middle (Excel)."""
    return {"x": v0["x"] + dx, "y": v0["y"] + dy, "l": v0["l"] - dx, "t": v0["t"] - dy}



def col_border(col: int) -> Anchor:
    return Anchor(
        f"column border {chr(65 + col)}",
        start=js(f"const r = R({{ r0: 0, c0: {col}, r1: 0, c1: {col} }}); return {{ x: r.r, y: r.t - 8 }};"),
        probe=f"__calc.app.grid.rectOnCanvas({{ r0: 0, c0: {col}, r1: 0, c1: {col} }}).w",
        expect=lambda dx, dy, v0: max(0.0, v0 + dx), tol=1.5, esc="keeps", quantum=1.0, out_follow=True, scroll=SCROLL, autoscroll=False, out_cancel=True)


def row_border(row: int) -> Anchor:
    return Anchor(
        f"row border {row + 1}",
        start=js(f"const r = R({{ r0: {row}, c0: 0, r1: {row}, c1: 0 }}); return {{ x: r.l - 14, y: r.b }};"),
        probe=f"__calc.app.grid.rectOnCanvas({{ r0: {row}, c0: 0, r1: {row}, c1: 0 }}).h",
        expect=lambda dx, dy, v0: max(0.0, v0 + dy), tol=1.5, esc="keeps", check="slow-down", quantum=1.0, out_follow=True, scroll=SCROLL, autoscroll=False, out_cancel=True)


def fill_expect(dx: float, dy: float, v0: dict, ctx: dict) -> dict:
    # Excel fills along the axis the pointer moved further on (in cells).
    if abs(dy / ctx["rh"]) >= abs(dx / ctx["cw"]):
        return {"x": v0["x"], "y": v0["y"] + dy}
    return {"x": v0["x"] + dx, "y": v0["y"]}


def anchors(cw: float, rh: float) -> list[Anchor]:
    cells = {"x": cw, "y": rh}
    # The cells' contents as one number (a sum of per-cell hashes, so the order does not matter).
    content = ("(() => { let sum = 0; for (const [key, cell] of __calc.app.sheet.cells) { if (cell.v === null) continue; let h = 7;"
               " for (const ch of key + '=' + cell.v) h = (h * 31 + ch.charCodeAt(0)) % 1000003; sum = (sum + h) % 1000000007; } return sum; })()")
    # The far corner of the fill range (the ghost while dragging): the edge on the side it grows to.
    fill_probe = js("const s = g.selection[0], q = g.ghost ?? s, a = R(q); return { x: q.c0 < s.c0 ? a.l : a.r, y: q.r0 < s.r0 ? a.t : a.b };")
    chart_setup = ("__calc.grid.select(__calc.parse('A1:B4'), { row: 0, col: 0 }); __calc.app.charts.insert('column');")
    chart_el = "document.querySelector('.chart-obj')"
    shape_setup = ("__calc.app.sheet.shapes = [{ from: { col: 3, row: 4, x: 0, y: 0 }, to: { col: 5, row: 12, x: 0, y: 0 }, geom: 'rect', name: 'Rechteck 1' }];"
                   " __calc.app.shapes.deselect(); __calc.grid.refreshLayout && __calc.grid.contentChanged();")
    shape_el = "document.querySelector('.shp')"
    shape_center = f"(() => {{ const r = {shape_el}.getBoundingClientRect(); return {{ x: r.left + r.width / 2, y: r.top + r.height / 2 }}; }})()"
    slider = "document.querySelector('.statusbar .zoom input[type=range]')"
    # Chromium's range input: the thumb (16 px) runs over the width minus its own width.
    slider_px = f"(() => {{ const s = {slider}, r = s.getBoundingClientRect(); return r.left + 8 + Number(s.value) / 100 * (r.width - 16); }})()"
    vs = "document.querySelector('.vscroll')"
    # The native scroll bar's thumb: arrows 17 px, the thumb proportional to what is shown.
    thumb = (f"const v = {vs}, r = v.getBoundingClientRect(), track = v.clientHeight - 34, len = Math.max(17, track * v.clientHeight / v.scrollHeight);"
             " const top = r.top + 17 + (v.scrollHeight > v.clientHeight ? v.scrollTop / (v.scrollHeight - v.clientHeight) * (track - len) : 0);")
    CHARTS = "document.querySelectorAll('.chart-obj').length"
    SHAPES = "document.querySelectorAll('.shp').length"
    tab_probe = ("(() => { const t = document.querySelector('.sheettab.drop') ?? document.querySelector('.sheettab.on'); const r = t.getBoundingClientRect(); return r.left + r.width / 2; })()")
    return [
        col_border(1),
        row_border(2),
        Anchor("fill handle", start=js("const a = R(g.selection[0]); return { x: a.r, y: a.b, cw: a.w, rh: a.h };"), probe=fill_probe,
               expect=fill_expect, tol=0.01, quantum=cells, monotone=False, undo_probe=content, check="slow-down", esc="keeps",
               scroll=SCROLL, autoscroll=True,
               note="value: the far corner of the fill range (client px); one cell of tolerance per axis"),
        Anchor("move selection by its border", start=js("const a = R(g.selection[0]); return { x: a.l + 2, y: (a.t + a.b) / 2 };"),
               probe=js("const a = R(g.ghost ?? g.selection[0]); return { x: a.l, y: a.t };"),
               expect=lambda dx, dy, v0: {"x": v0["x"] + dx, "y": v0["y"] + dy}, tol=0.01, quantum=cells, undo_probe=content, check="diagonal", esc="keeps",
               scroll=SCROLL, mods={"shift": "skip"}, out_cancel=True,
               note="value: the target range's top left corner"),
        Anchor("select a range", start=js("const a = R(g.selection[0]); return { x: (a.l + a.r) / 2, y: (a.t + a.b) / 2 };"),
               probe=js("const a = R({ r0: g.head.row, c0: g.head.col, r1: g.head.row, c1: g.head.col }); return { x: (a.l + a.r) / 2, y: (a.t + a.b) / 2 };"),
               expect=lambda dx, dy, v0: {"x": v0["x"] + dx, "y": v0["y"] + dy}, tol=0.01, quantum=cells, esc=False, undo=False,
               scroll=SCROLL, autoscroll=True, out_dir="down",
               note="value: the middle of the active corner cell"),
        Anchor("chart move", setup=chart_setup,
               start=f"(() => {{ const r = {chart_el}.getBoundingClientRect(); return {{ x: r.left + 6, y: r.top + 6 }}; }})()",
               probe=f"(() => {{ const r = {chart_el}.getBoundingClientRect(); return {{ x: r.left, y: r.top }}; }})()",
               expect=lambda dx, dy, v0: {"x": v0["x"] + dx, "y": v0["y"] + dy}, threshold=3,
               scroll=SCROLL, count=CHARTS, copies={"ctrl": 1}, mods={"shift": constrain, "alt": ("snap", cw)}),
        Anchor("chart corner handle", setup=chart_setup,
               start=f"(() => {{ const r = {chart_el}.querySelector('.chart-handle[data-h=se]').getBoundingClientRect(); return {{ x: r.left + r.width / 2, y: r.top + r.height / 2 }}; }})()",
               probe=f"(() => {{ const r = {chart_el}.getBoundingClientRect(); return {{ x: r.right, y: r.bottom, l: r.left, t: r.top }}; }})()",
               expect=corner, threshold=3, out_follow=True, scroll=SCROLL,
               mods={"shift": corner_aspect, "ctrl": corner_centre, "alt": ("snap", cw)}),
        Anchor("shape move (not selected before)", setup=shape_setup, start=shape_center,
               probe=f"(() => {{ const r = {shape_el}.getBoundingClientRect(); return {{ x: r.left, y: r.top }}; }})()",
               expect=lambda dx, dy, v0: {"x": v0["x"] + dx, "y": v0["y"] + dy},
               scroll=SCROLL, count=SHAPES, copies={"ctrl": 1}, mods={"shift": constrain, "alt": ("snap", cw)}),
        Anchor("shape corner handle", setup=shape_setup, click=shape_center,
               start=f"(() => {{ const r = {shape_el}.querySelector('.shp-h.se').getBoundingClientRect(); return {{ x: r.left + r.width / 2, y: r.top + r.height / 2 }}; }})()",
               probe=f"(() => {{ const r = {shape_el}.getBoundingClientRect(); return {{ x: r.right, y: r.bottom, l: r.left, t: r.top }}; }})()",
               expect=corner, out_follow=True, scroll=SCROLL,
               mods={"shift": corner_aspect, "ctrl": corner_centre, "alt": ("snap", cw)}),
        Anchor("shape rotation handle", setup=shape_setup, click=shape_center,
               start=(f"(() => {{ const e = {shape_el}, r = e.getBoundingClientRect(), h = e.querySelector('.shp-h.rot').getBoundingClientRect();"
                      " return { x: h.left + h.width / 2, y: h.top + h.height / 2, cx: r.left + r.width / 2, cy: r.top + r.height / 2 }; })()"),
               probe=f"(() => {{ const m = /rotate\\(([-\\d.e]+)deg\\)/.exec({shape_el}.style.transform); return m ? Number(m[1]) : 0; }})()",
               expect=lambda dx, dy, v0, ctx: math.degrees(math.atan2(ctx["y"] + dy - ctx["cy"], ctx["x"] + dx - ctx["cx"])) + 90,
               tol=1.5, wrap=360, per_px=1.5, out_follow=True, mods={"shift": ("snap", 15.0)},
               note="value: degrees; expected: the angle of the pointer around the middle"),
        Anchor("zoom slider", start=f"(() => {{ const r = {slider}.getBoundingClientRect(); return {{ x: {slider_px}, y: r.top + r.height / 2, lo: r.left + 8, hi: r.right - 8 }}; }})()",
               probe=slider_px, expect=lambda dx, dy, v0, ctx: min(ctx["hi"], max(ctx["lo"], v0 + dx)), tol=1.5, esc=False, undo=False,
               out_follow=True, zoomed=False, note="value: the thumb's middle (client px); Excel cannot undo zooming"),
        Anchor("vertical scroll bar thumb", start=js(f"{thumb} return {{ x: r.left + 8, y: top + len / 2, lo: r.top + 17 }};"),
               probe=js(f"{thumb} return top;"), expect=lambda dx, dy, v0, ctx: max(ctx["lo"], v0 + dy), tol=2, esc=False, undo=False, planned=True, lag=4,
               zoomed=False, out_dir="down",
               paths=[p for p in PATHS if p != "long"],
               note="value: the thumb's top (from scrollTop and the scroll bar's geometry); not the long path: Windows' scroll bars"
                    " put the thumb back when the pointer goes far to the side. Flaky: in 2 of 4 runs the zigzag's last move (back up past"
                    " the start, beside the window's edge) was not applied by the native scroll bar (the thumb stays 8 px down; no jump;"
                    " RaabeCalc's own scroll code is not called during the drag)"),
        Anchor("sheet tab reorder", start="(() => { const r = document.querySelector('.sheettab.on').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()",
               probe=tab_probe, expect=lambda dx, dy, v0: v0 + dx, tol=0.01, quantum=90, undo=False, threshold=6, lag=3,
               zoomed=False, count="document.querySelectorAll('.sheettab').length", copies={"ctrl": 1},
               note="value: the middle of the tab the sheet would go to; Excel cannot undo moving a sheet"),
        Anchor("sheet bar splitter", start="(() => { const r = document.querySelector('.sheetsplit').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2, hi: document.querySelector('.sheetbar').clientWidth - 80 }; })()",
               probe="document.querySelector('.sheetleft').getBoundingClientRect().width", expect=lambda dx, dy, v0, ctx: min(ctx["hi"], max(120.0, v0 + dx)),
               esc=False, undo=False, out_follow=True, zoomed=False,
               note="between the sheet tabs and the horizontal scroll bar (Excel has no Namenfeld splitter drag in RaabeCalc yet)"),
    ]


def main_paths(families: list[str], want: list[str]) -> None:
    out = HERE / "results" / "2026-11-09-drag-calc-paths.json"
    report: dict = {"app": "raabe-calc", "families": families, "anchors": []}
    with (contextlib.nullcontext() if dragtest.INPUT == "cdp" else real_input("drag_calc.py --paths (drag-paths agent)")):
        side = CalcSide(port=int(__import__("os").environ.get("RAABE_DRAG_PORT", "9481")))  # a port of our own: other sessions' runs use the default
        guard = catcher(WINDOW) if dragtest.INPUT != "cdp" else contextlib.nullcontext()
        guard.__enter__()
        try:
            winapi.place(side.hwnd, *WINDOW)
            side.reset(CASE)
            size = side.js("(() => { const b = __calc.grid.rectOnCanvas({ r0: 2, c0: 1, r1: 2, c1: 1 }); return { w: b.w, h: b.h }; })()")
            sel = [a for a in anchors(size["w"], size["h"]) if not want or any(w.lower() in a.name.lower() for w in want)]

            def prepare(anchor: Anchor) -> None:
                side.reset(CASE)

            report["anchors"] = run_families(side, sel, side.js, prepare, families, zoom_js=ZOOM, reveal=REVEAL,
                                             log=lambda *a: print(*a, flush=True), touch=touch)
        finally:
            guard.__exit__(None, None, None)
            side.close()
    report = merge_report(out, report)
    out.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    for a in report["anchors"]:
        print(f"{a['anchor']}: {a['ok']}/{a['of']}")


def main() -> None:
    want = sys.argv[1:]
    if "--cdp" in want:
        want = [w for w in want if w != "--cdp"]
        dragtest.cdp_mode()
    if want and want[0] == "--paths":
        main_paths(want[1].split(","), want[2:])
        return
    report: dict = {"app": "raabe-calc", "anchors": []}
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-drag-calc.json"
    with real_input("drag_calc.py"):
        side = CalcSide()
        try:
            side.reset(CASE)
            size = side.js("(() => { const b = __calc.grid.rectOnCanvas({ r0: 2, c0: 1, r1: 2, c1: 1 }); return { w: b.w, h: b.h }; })()")
            for anchor in anchors(size["w"], size["h"]):
                if want and not any(w.lower() in anchor.name.lower() for w in want):
                    continue
                res = []
                for p in select_paths(anchor) + check_paths(anchor):
                    side.reset(CASE)
                    winapi.activate(side.hwnd)
                    time.sleep(0.2)
                    rec = run_drag(side, anchor, p, side.js)
                    winapi.press("esc") if p.name != "escape" else None
                    a = analyse(anchor, rec)
                    res.append(a)
                    print(anchor.name, p.name, "ok" if a["ok"] else "FAIL", {k: v for k, v in a.items() if k not in ("path", "ok", "v0", "after")}, flush=True)
                    touch()
                report["anchors"].append(summary(anchor, res))
        finally:
            side.close()
    if want and out.exists():
        # A partial run replaces only its anchors.
        old = json.loads(out.read_text(encoding="utf-8"))
        names = {a["anchor"] for a in report["anchors"]}
        report["anchors"] = [a for a in old.get("anchors", []) if a["anchor"] not in names] + report["anchors"]
    out.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    for a in report["anchors"]:
        print(f"{a['anchor']}: {a['ok']}/{a['of']}")


if __name__ == "__main__":
    main()
