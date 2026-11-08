"""The drag test in RaabeWrite (D-080 C, D-096): the ruler's markers, a picture's handles, a table's borders, selecting
text, the zoom slider and the scroll bar, each dragged along feellab/dragtest.py's paths with real input and checked
for jumps, plus the Esc and Strg+Z drags. Takes mouse and keyboard (the real-input lock).

    .venv/Scripts/python.exe drag_write.py [anchor ...]     (names, or parts of them)

Documents: the feel fixture "long" (built in the app), Word's own word-images.docx and word-rich.docx from
packages/ooxml/test/fixtures, copied to TEMP. RAABE_WRITE_EXE points at another build.
Result: results/<date>-drag-write.json.
"""
from __future__ import annotations

import datetime
import json
import math
import shutil
import sys
import tempfile
import time
from pathlib import Path

from feellab import winapi
from feellab.catalogue import RaabeSide
from feellab.dragtest import Anchor, analyse, check_paths, paths, run_drag, select_paths, summary
from feellab.inputlock import real_input, touch

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
FIX = ROOT / "packages/ooxml/test/fixtures"
TMP = Path(tempfile.gettempdir()) / "raabe-drag"
CM = 28.3465
PATHS = [p.name for p in paths()]

# ----- documents ------------------------------------------------------------------------------

# The body paragraph the ruler and selection anchors use: indents 2 cm / first line +1 cm / right 1 cm, a tab stop at 8 cm.
LONG = ("(() => { __probe.load('long'); const e = __app.editor(); const p = e.doc.main.findIndex((q, i) => i > 0 && q.text.length > 200);"
        f" e.setSel({{ p, o: 3 }}, {{ p, o: 3 }}); e.setParaProps({{ indLeft: {2 * CM}, indFirst: {1 * CM}, indRight: {1 * CM}, tabs: [{{ pos: {8 * CM}, align: 'left' }}] }});"
        " e.update(); e.view.scrollTo ? 0 : 0; return p; })()")
RULER_ON = "(() => { const r = document.querySelector('.hruler'); if (r && r.hidden) __app.run('view.ruler'); return !!r; })()"
OBJ = "String.fromCharCode(0xfffc)"
SEL_INLINE = (f"(() => {{ const e = __app.editor(), O = {OBJ}; const p = e.doc.main.findIndex((q) => q.text.includes(O));"
              " const o = e.doc.main[p].text.indexOf(O); e.setSel({ p, o }, { p, o: o + 1 }); e.update(); return !!e.selectedImage(); })()")
SEL_FLOAT = ("(() => { const e = __app.editor(); const p = e.doc.main.findIndex((q) => (q.floats || []).some((f) => f.image));"
             " const i = e.doc.main[p].floats.findIndex((f) => f.image); e.selectFloat(p, i); e.update(); return !!e.selectedImage(); })()")


def open_doc(side: RaabeSide, name: str) -> None:
    """A fresh copy of a fixture (the documents a drag changed are discarded)."""
    TMP.mkdir(exist_ok=True)
    t = side.t
    t.js("__app.newDoc()")
    t.js("__app.closeOthers()")
    src = FIX / name
    dst = TMP / name
    shutil.copyfile(src, dst)
    t.js(f"__app.openPath({str(dst)!r})")
    deadline = time.perf_counter() + 15
    while time.perf_counter() < deadline:
        docs = t.js("__app.docs()") or []
        if any(d["active"] and d["path"] and Path(d["path"]) == dst for d in docs) and t.js("__probe.settled()"):
            break
        time.sleep(0.1)
    t.js("__app.closeOthers()")
    time.sleep(0.4)


# ----- anchors --------------------------------------------------------------------------------

def guide_or(marker_js: str, axis: str = "x") -> str:
    """The ruler's value: the guide line while dragging, else where the marker is (client px)."""
    if axis == "x":
        return ("(() => { const g = document.querySelector('.rul-guide'); const a = document.querySelector('.docarea').getBoundingClientRect();"
                " if (g && !g.hidden && g.style.left) return a.left + parseFloat(g.style.left);"
                f" const m = {marker_js}; return m ? m.parentElement.getBoundingClientRect().left + parseFloat(m.style.left) : null; }})()")
    return ("(() => { const g = document.querySelector('.rul-guide'); const a = document.querySelector('.docarea').getBoundingClientRect();"
            " if (g && !g.hidden && g.style.top) return a.top + parseFloat(g.style.top);"
            f" const m = {marker_js}; return m ? m.parentElement.getBoundingClientRect().top + parseFloat(m.style.top) : null; }})()")


def ruler(name: str, marker_js: str, axis: str = "x", nudge: float = 0.0, y_at: float | None = None, hi: bool = False) -> Anchor:
    px = "(__app.editor().view.zoom * 96 / 72)"
    if axis == "x":
        start = (f"(() => {{ const m = {marker_js}; const r = m.getBoundingClientRect(), t = m.parentElement.getBoundingClientRect();"
                 f" return {{ x: t.left + parseFloat(m.style.left) + {nudge}, y: {'r.top + r.height / 2' if y_at is None else f'r.top + {y_at}'}, step: 0.25 * {CM} * {px},"
                 " hi: (() => { const e = __app.editor(); return e.view.pageElement(e.stats().caretPage).getBoundingClientRect().right; })() }; })()")
    else:
        start = (f"(() => {{ const m = {marker_js}; const r = m.getBoundingClientRect(), t = m.parentElement.getBoundingClientRect();"
                 f" return {{ x: r.left + r.width / 2, y: t.top + parseFloat(m.style.top), step: 0.25 * {CM} * {px} }}; }})()")
    # 0.25 cm at 100 % is 9.45 px: the snapping may hold the value back or pull it ahead by up to one step.
    return Anchor(f"ruler: {name}", setup=f"{RULER_ON}; {LONG}", start=start, probe=guide_or(marker_js, axis),
                  expect=(lambda dx, dy, v0, ctx: min(ctx["hi"], v0 + dx) if hi else v0 + dx) if axis == "x" else (lambda dx, dy, v0: v0 + dy), esc=False,
                  tol=1.5, snap=9.5, check="slow-right" if axis == "x" else "slow-down",
                  note="value: the guide line (client px); snaps to 0.25 cm (9.45 px at 100 %); Esc not checked (Word's not recorded)")


def rect_js(body: str) -> str:
    return f"(() => {{ const e = __app.editor(), r = e.view.selectedObjectRect(); if (!r) return null; {body} }})()"


def cell_js(body: str) -> str:
    """The first table's first cell on the page (layout frame) as client points a (top right) and b (bottom right)."""
    return ("(() => { const e = __app.editor(), L = e.layout; let page = -1, fr = null;"
            " L.pages.forEach((pg, i) => { for (const id of pg.frames) { const f = L.frames[id]; if (!fr && f.kind === 'cell' && f.story === 'main' && f.cell && f.cell.endsWith(':0:0')) { fr = f; page = i; } } });"
            " if (!fr) return null; const a = e.view.pageToClient(page, fr.x + fr.w, fr.y), b = e.view.pageToClient(page, fr.x + fr.w, fr.y + fr.h),"
            f" l = e.view.pageToClient(page, fr.x, fr.y);"
            " const nx = L.frames.find((f) => f.kind === 'cell' && f.story === 'main' && f.cell === fr.cell.replace(/:0$/, ':1'));"
            f" const hi = nx ? e.view.pageToClient(page, nx.x + nx.w - 12 * 96 / 72, nx.y).x : 1e9; {body} }})()")


def corner_expect(dx: float, dy: float, v0: dict) -> dict:
    # A picture's corner keeps its proportions: the side that grows more decides (Word).
    k = max((v0["w"] + dx) / v0["w"], (v0["h"] + dy) / v0["h"])
    k = max(k, 8 / min(v0["w"], v0["h"]))
    return {"w": v0["w"] * k, "h": v0["h"] * k}


def anchors() -> list[Anchor]:
    first = "document.querySelector('.hruler .rul-first')"
    hanging = "document.querySelector('.hruler .rul-hanging')"
    leftbox = "document.querySelector('.hruler .rul-leftbox')"
    right = "document.querySelector('.hruler .rul-right')"
    tab = "document.querySelector('.hruler .rul-tab')"
    lrand = "document.querySelector('.hruler .rul-edge.h[data-tip=\"Linker Rand\"]')"
    orand = "document.querySelector('.vruler .rul-edge.v[data-tip=\"Oberer Rand\"]')"
    ws = "document.querySelector('.docarea .workspace:not([hidden])')"
    thumb = (f"const v = {ws}, r = v.getBoundingClientRect(), track = v.clientHeight - 34, len = Math.max(17, track * v.clientHeight / v.scrollHeight);"
             " const top = r.top + 17 + (v.scrollHeight > v.clientHeight ? v.scrollTop / (v.scrollHeight - v.clientHeight) * (track - len) : 0);")
    slider = "document.querySelector('.statusbar .zoom input[type=range]')"
    slider_px = f"(() => {{ const s = {slider}, r = s.getBoundingClientRect(); return r.left + 8 + Number(s.value) / 100 * (r.width - 16); }})()"
    caret = "(() => { const s = __probe.selection(), p = __probe.pointAt(s.head); return { x: p.x, y: p.y + p.h / 2 }; })()"
    rot = "(() => { const i = __app.editor().selectedImage(); return i ? (i.obj.pic && i.obj.pic.rotation) || 0 : null; })()"
    return [
        ruler("Erstzeileneinzug", first),
        ruler("Hängender Einzug", hanging),
        ruler("Linker Einzug", leftbox),
        ruler("Rechter Einzug", right, hi=True),
        ruler("Tabstopp", tab, nudge=1.5),
        ruler("Linker Rand", lrand, y_at=4),
        ruler("Oberer Rand", orand, axis="y"),
        Anchor("picture corner handle (inline)", setup=f"__app.editor() && 0", start=rect_js("return { x: r.right, y: r.bottom };"),
               probe=rect_js("return { w: r.width, h: r.height };"), expect=corner_expect, tol=1.5, per_px=1.6,
               note="word-images.docx, the inline picture (3:2); value: its size, proportional, so the width may grow 1.5 px per px down"),
        Anchor("picture side handle (floating)", start=rect_js("return { x: r.right, y: r.top + r.height / 2 };"),
               probe=rect_js("return r.width;"), expect=lambda dx, dy, v0: max(8.0, v0 + dx), tol=1.5, snap=6,
               note="word-images.docx, the square-wrapped picture; snaps to the margins within 6 px"),
        Anchor("picture move (floating)", start=rect_js("return { x: r.left + r.width / 2, y: r.top + r.height / 2 };"),
               probe=rect_js("return { x: r.left, y: r.top };"), expect=lambda dx, dy, v0: {"x": v0["x"] + dx, "y": v0["y"] + dy}, tol=1.5, snap=6),
        Anchor("picture rotation handle", start=rect_js("const z = e.view.zoom; return { x: r.left + r.width / 2, y: r.top - 17 * z, cx: r.left + r.width / 2, cy: r.top + r.height / 2 };"),
               probe=rot, expect=lambda dx, dy, v0, ctx: math.degrees(math.atan2(ctx["y"] + dy - ctx["cy"], ctx["x"] + dx - ctx["cx"])) + 90,
               tol=1.5, wrap=360, per_px=1.2, note="value: degrees; expected: the angle of the pointer around the middle"),
        Anchor("table column border", start=cell_js("return { x: a.x, y: (a.y + b.y) / 2, hi: hi };"), probe=cell_js("return a.x;"),
               expect=lambda dx, dy, v0, ctx: min(ctx["hi"], v0 + dx), tol=1.5,
               note="word-rich.docx, the border right of the first cell; the next column keeps at least 12 pt"),
        Anchor("table row border", start=cell_js("return { x: (l.x + a.x) / 2, y: b.y };"), probe=cell_js("return b.y;"),
               expect=lambda dx, dy, v0: v0 + max(0.0, dy), tol=1.5, check="slow-down",
               note="word-rich.docx, the first row's bottom border; the row never gets lower than its text"),
        Anchor("select text", setup=LONG, start=caret, probe=caret, expect=lambda dx, dy, v0: {"x": v0["x"] + dx, "y": v0["y"] + dy},
               tol=0.01, quantum={"x": 110, "y": 32}, esc=False, undo=False, paths=[n for n in PATHS if n != "long"],
               note="value: the caret (selection's head); Word's whole-word selection lets it run ahead by a word, a line; not the"
                    " long path: it reaches the paragraph's short last line, where the caret rightly stops at the line's end"),
        Anchor("zoom slider", start=f"(() => {{ const r = {slider}.getBoundingClientRect(); return {{ x: {slider_px}, y: r.top + r.height / 2, lo: r.left + 8, hi: r.right - 8 }}; }})()",
               probe=slider_px, expect=lambda dx, dy, v0, ctx: min(ctx["hi"], max(ctx["lo"], v0 + dx)), tol=1.5, esc=False, undo=False,
               note="value: the thumb's middle (client px); Word cannot undo zooming"),
        Anchor("vertical scroll bar thumb", setup=LONG, start=f"(() => {{ {thumb} return {{ x: r.right - 8, y: top + len / 2, lo: r.top + 17 }}; }})()",
               probe=f"(() => {{ {thumb} return top; }})()", expect=lambda dx, dy, v0, ctx: max(ctx["lo"], v0 + dy), tol=2, esc=False, undo=False,
               planned=True, lag=4, paths=[n for n in PATHS if n != "long"],
               note="value: the thumb's top (from scrollTop and the scroll bar's geometry); not the long path: Windows' scroll bars"
                    " put the thumb back when the pointer goes far to the side"),
    ]


DOCS = {"picture corner handle (inline)": ("word-images.docx", SEL_INLINE), "picture side handle (floating)": ("word-images.docx", SEL_FLOAT),
        "picture move (floating)": ("word-images.docx", SEL_FLOAT), "picture rotation handle": ("word-images.docx", SEL_FLOAT),
        "table column border": ("word-rich.docx", None), "table row border": ("word-rich.docx", None)}


def prepare(side: RaabeSide, anchor: Anchor) -> None:
    doc = DOCS.get(anchor.name)
    if doc:
        open_doc(side, doc[0])
        if doc[1] and not side.t.js(doc[1]):
            raise RuntimeError(f"{anchor.name}: nothing selected")
    else:
        side.t.js("__app.newDoc()")
        side.t.js("__app.closeOthers()")
    time.sleep(0.3)


def main() -> None:
    want = sys.argv[1:]
    report: dict = {"app": "raabe-write", "anchors": []}
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-drag-write.json"
    with real_input("drag_write.py"):
        side = RaabeSide()
        js = side.t.js
        try:
            for anchor in anchors():
                if want and not any(w.lower() in anchor.name.lower() for w in want):
                    continue
                res = []
                for p in select_paths(anchor) + check_paths(anchor):
                    try:
                        prepare(side, anchor)
                    except Exception as e:  # noqa: BLE001
                        res.append({"path": p.name, "ok": False, "error": str(e)[:200]})
                        continue
                    winapi.activate(side.hwnd)
                    time.sleep(0.2)
                    rec = run_drag(side, anchor, p, js)
                    if p.name != "escape":
                        winapi.press("esc")
                    a = analyse(anchor, rec)
                    res.append(a)
                    print(anchor.name, p.name, "ok" if a["ok"] else "FAIL", {k: v for k, v in a.items() if k not in ("path", "ok", "v0", "after")}, flush=True)
                    touch()
                report["anchors"].append(summary(anchor, res))
        finally:
            side.close()
    if want and out.exists():
        old = json.loads(out.read_text(encoding="utf-8"))
        names = {a["anchor"] for a in report["anchors"]}
        report["anchors"] = [a for a in old.get("anchors", []) if a["anchor"] not in names] + report["anchors"]
    out.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    for a in report["anchors"]:
        print(f"{a['anchor']}: {a['ok']}/{a['of']}")


if __name__ == "__main__":
    main()
