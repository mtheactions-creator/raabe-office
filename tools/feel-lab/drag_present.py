"""The drag test in RaabePresent (D-080 C, D-110): every anchor dragged along feellab/dragtest.py's paths with real input
and checked for jumps (follow, jump, direction, release); then per anchor Esc during a drag and Strg+Z after one, which
must bring back the value from before. Values are read from the app's own elements in client CSS px (a frame's left
edge, a size, a column's width, a pane's size, the thumbnails' drop line) or degrees (rotation). Takes mouse and keyboard
(the real-input lock, feellab/inputlock.py).

    .venv/Scripts/python.exe drag_present.py [anchor ...]
    .venv/Scripts/python.exe drag_present.py --paths out,mods,zoom [anchor ...]

--paths runs D-127's paths instead (out of the window and back, the button let go outside, Umschalt/Strg/Alt mid-drag
as PowerPoint does them (probe_powerpoint_drag_mods.py), other zoom levels) into results/2026-11-09-drag-present-paths.json.

Result: results/<date>-drag-present.json.
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
from feellab.dragtest import Anchor, Path as DragPath, analyse, merge_report, paths, run_drag, run_families, screen_point
from feellab.inputlock import real_input, touch
from feellab.present_catalogue import Case, PresentSide

HERE = Path(__file__).parent
SID = "__present.state().shapes.filter((s) => !s.ph)[0].id"
RECT = {"kind": "rect", "box": [300, 150, 200, 120]}
G: dict = {}  # measured once per run: the shape's height in px, the table's next column, the zoom slider's track …
# D-127: the slide view's scroll position, revealing a start point, the window with room below and right of it.
SCROLL = "(document.querySelector('.editor').scrollTop + document.querySelector('.editor').scrollLeft)"
REVEAL = "(() => { const e = document.querySelector('.editor'); e.scrollLeft += {dx}; e.scrollTop += {dy}; })()"
WINDOW = (0, 0, 2000, 1700)
SHAPES = "__present.state().shapes.filter((s) => !s.ph).length"
ASPECT = RECT["box"][3] / RECT["box"][2]  # the test rectangle's height / width



def shape(name: str, kind: dict, prop: str, expect, start: str, setup: str = "", tol: float = 2.0, snap: float = 5.0, **kw) -> tuple[Anchor, dict]:
    probe = f"(() => {{ const r = __present.shapeRect({SID}); return r ? r.{prop} : null; }})()"
    return Anchor(name, start=start, probe=probe, expect=expect, tol=tol, snap=snap, setup=setup, **kw), {"shapes": [kind]}


MOVE = {"scroll": SCROLL, "count": SHAPES, "copies": {"ctrl": 1}, "mods": {"alt": "free"}}


SELECT = f"__present.app.editor.select([{SID}])"


def handle(k: str, prop: str, sign: int, axis: str) -> tuple[Anchor, dict]:
    # Past the opposite side the shape flips (as in PowerPoint): the size grows again.
    exp = (lambda dx, dy, v0: abs(v0 + sign * (dx if axis == "x" else dy)))
    # Strg: the size changes around the middle (twice the motion); Umschalt at a corner: the proportions stay.
    centre = (lambda dx, dy, v0: abs(v0 + 2 * sign * (dx if axis == "x" else dy)))

    def aspect(dx, dy, v0):
        h0 = v0 * ASPECT
        kk = max((v0 + sign * dx) / v0, (h0 + sign * dy) / h0)
        return abs(v0 * kk)
    mods = {"ctrl": centre, "alt": "free"}
    if len(k) == 2:
        mods["shift"] = aspect
    return shape(f"handle {k}", RECT, prop, exp, start=f"__present.handlePoint({SID}, '{k}')", setup=SELECT, snap=0.0,
                 out_follow=True, mods=mods)


def rotation() -> tuple[Anchor, dict]:
    probe = (f"(() => {{ const e = document.querySelector('.stage .shp[data-id=\"' + {SID} + '\"]'); const m = /rotate\\(([-\\d.e]+)deg/.exec(e?.style.transform ?? '');"
             " const a = m ? Number(m[1]) : 0; return a > 180 ? a - 360 : a; })()")

    def expect(dx, dy, v0, ctx):
        # The handle relative to the shape's centre where the press was (px; it changes with the zoom, D-127).
        rx, ry = ctx["x"] - ctx["cx"], ctx["y"] - ctx["cy"]
        a = math.degrees(math.atan2(ry + dy, rx + dx)) + 90
        return a - 360 if a > 180 else a
    start = (f"(() => {{ const p = __present.handlePoint({SID}, 'rot'), r = __present.shapeRect({SID});"
             " return { x: p.x, y: p.y, cx: r.x + r.w / 2, cy: r.y + r.h / 2 }; })()")
    # The angle may stay behind by the drag threshold (4 px of the first movement, about 2.7 degrees at this handle's
    # radius): the threshold PowerPoint's catalogue cases showed, not a jump (D-110).
    return Anchor("rotation handle", start=start, probe=probe, expect=expect, tol=1.5, snap=2.8, setup=SELECT,
                  wrap=360, out_follow=True, mods={"shift": ("snap", 15.0)}), {"shapes": [RECT]}


def table_column() -> tuple[Anchor, dict]:
    probe = "(() => document.querySelector('.stage td[data-r=\"0\"][data-c=\"0\"]')?.getBoundingClientRect().width ?? null)()"
    start = "(() => { const r = document.querySelector('.stage td[data-r=\"0\"][data-c=\"0\"]').getBoundingClientRect(); return { x: r.right, y: r.top + r.height / 2 }; })()"
    return Anchor("table column border", start=start, probe=probe, expect=lambda dx, dy, v0: min(max(G["col_min"], v0 + dx), v0 + G["col_room"]), tol=1.5,
                  out_follow=True), \
        {"shapes": [{"kind": "table", "box": [100, 120, 600, 120]}]}


def thumbnails() -> tuple[Anchor, dict]:
    # The drop line's place in the list; before the drag starts, where it would be for the slide's own place.
    probe = ("(() => { const list = document.querySelector('.thumbs'); const l = list.querySelector('.drop-line');"
             " const lr = list.getBoundingClientRect(); if (l) return parseFloat(l.style.top);"
             " const i = __present.app.deck.slides.findIndex((x) => x.path === window.__dragSlide);"
             " const it = document.querySelector('.thumb-item[data-i=\"' + i + '\"]').getBoundingClientRect(); return it.top - lr.top + list.scrollTop - 4; })()")
    start = "(() => { const r = __present.thumbRect(2); return { x: r.x + r.w / 2, y: r.y + r.h / 2 - 6 }; })()"
    # The probe follows the dragged slide after the drop (its place, not the place it had).
    return Anchor("slide thumbnail", start=start, probe=probe, expect=lambda dx, dy, v0: v0 + dy, tol=2.0, quantum=G.get("thumb_pitch", 100), monotone=True,
                  setup="window.__dragSlide = __present.app.deck.slides[1].path", zoomed=False, out_dir="down"), \
        {"shapes": [], "slides": 4}


def zoom_slider() -> tuple[Anchor, dict]:
    # RaabePresent's own slider (D-110): the thumb's centre, read from the page.
    thumb = "(() => { const r = document.querySelector('.zoom-thumb').getBoundingClientRect(); return r.left + r.width / 2; })()"
    start = "(() => { const r = document.querySelector('.zoom-thumb').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()"
    return Anchor("zoom slider", start=start, probe=thumb, expect=lambda dx, dy, v0: min(max(G["zoom_lo"], v0 + dx), G["zoom_hi"]), tol=1.5,
                  out_follow=True, zoomed=False), {"shapes": [RECT]}


def pane_splitter() -> tuple[Anchor, dict]:
    probe = "document.querySelector('.present-main').firstElementChild.getBoundingClientRect().width"
    start = "(() => { const r = document.querySelector('.present-main > .splitter').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()"
    return Anchor("thumbnail pane splitter", start=start, probe=probe, expect=lambda dx, dy, v0: min(max(90.0, v0 + dx), 420.0), tol=1.5,
                  out_follow=True, zoomed=False), {"shapes": [RECT]}


def notes_splitter() -> tuple[Anchor, dict]:
    probe = "document.querySelector('.notes-wrap').getBoundingClientRect().height"
    start = "(() => { const r = document.querySelector('.notes-grip').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()"
    return Anchor("notes splitter", start=start, probe=probe, expect=lambda dx, dy, v0: min(max(24.0, v0 - dy), G["notes_hi"]), tol=1.5,
                  out_follow=True, zoomed=False, out_dir="down"), {"shapes": [RECT]}


def anchors() -> dict[str, tuple[Anchor, dict]]:
    centre = f"(() => {{ const r = __present.shapeRect({SID}); return {{ x: r.x + r.w * 0.55, y: r.y + r.h * 0.5 }}; }})()"
    top_edge = f"(() => {{ const r = __present.shapeRect({SID}); return {{ x: r.x + r.w * 0.3, y: r.y }}; }})()"
    pic = ("(async () => { const c = document.createElement('canvas'); c.width = 240; c.height = 150; const g = c.getContext('2d'); g.fillStyle = '#3a7'; g.fillRect(0, 0, 240, 150);"
           " const b = await new Promise((r) => c.toBlob(r)); __present.app.insertPictureBytes(new Uint8Array(await b.arrayBuffer()), 'image/png'); __present.app.editor.select([]); })()")
    out = [
        shape("move shape", RECT, "x", lambda dx, dy, v0: v0 + dx, centre, **MOVE),
        shape("move text box", {"kind": "textbox", "box": [300, 150, 240, 40], "text": "Ein Textfeld"}, "x", lambda dx, dy, v0: v0 + dx, top_edge, setup=SELECT, **MOVE),
        shape("move picture", {"kind": "rect", "box": [20, 20, 10, 10]}, "x", lambda dx, dy, v0: v0 + dx,
              f"(() => {{ const v = __present.state().shapes.filter((s) => !s.ph).pop(); const r = __present.shapeRect(v.id); return {{ x: r.x + r.w * 0.5, y: r.y + r.h * 0.5 }}; }})()", setup=pic, **MOVE),
        handle("se", "w", 1, "x"), handle("e", "w", 1, "x"), handle("s", "h", 1, "y"),
        handle("nw", "w", -1, "x"), handle("n", "h", -1, "y"), handle("w", "w", -1, "x"),
        rotation(), table_column(), thumbnails(), zoom_slider(), pane_splitter(), notes_splitter(),
    ]
    # The picture's probe and start take the last shape (the picture), not the small marker.
    a, spec = out[2]
    a.probe = "(() => { const v = __present.state().shapes.filter((s) => !s.ph).pop(); const r = __present.shapeRect(v.id); return r ? r.x : null; })()"
    return {a.name: (a, spec) for a, spec in out}


# Esc during a drag and Strg+Z after one: what to compare (document state for document drags, the value for panes).
UNDOABLE = {"move shape", "move text box", "move picture", "handle se", "handle e", "handle s", "handle nw", "handle n", "handle w", "rotation handle",
            "table column border", "slide thumbnail"}
ORDER = "__present.app.deck.slides.map((s) => s.path).join(',')"


def esc_and_undo(side, anchor: Anchor, spec: dict) -> dict:
    """Esc while dragging (the value from before must come back), then a real drag and Strg+Z (the same)."""
    state = ORDER if anchor.name == "slide thumbnail" else anchor.probe
    out: dict = {}
    for kind in ("esc", "undo"):
        if kind == "undo" and anchor.name not in UNDOABLE:
            out["undo"] = "n/a"
            continue
        reset(side, spec)
        if anchor.setup:
            side.js(anchor.setup)
            time.sleep(0.3)
        st = side.js(anchor.start)
        before = side.js(state)
        sx, sy = screen_point(side, st["x"], st["y"])
        winapi.move(sx, sy)
        time.sleep(0.08)
        winapi.mouse_down(sx, sy)
        time.sleep(0.05)
        for i in range(1, 31):
            px, py = screen_point(side, st["x"] + i * 1.5, st["y"] + i)
            winapi.move(px, py)
            time.sleep(0.016)
        time.sleep(0.15)
        if kind == "esc":
            winapi.press("esc")
            time.sleep(0.1)
        winapi.mouse_up(px, py)
        time.sleep(0.3)
        if kind == "undo":
            moved = side.js(state)
            winapi.press("ctrl+z")
            time.sleep(0.4)
            out["undo_moved"] = moved != before
        after = side.js(state)
        same = after == before if not isinstance(before, (int, float)) else isinstance(after, (int, float)) and abs(after - before) <= anchor.tol
        out[kind] = "ok" if same else f"{before} -> {after}"
    return out


def reset(side, spec: dict) -> None:
    side.reset(Case("drag", "drag", spec.get("shapes", []), [], slides=spec.get("slides", 1)))
    winapi.activate(side.hwnd)
    if dragtest.INPUT == "cdp":
        # The slide area gets the focus (PresentSide.reset's real click beside the slide).
        from feellab.present_catalogue import SLIDE_H, SLIDE_W
        c = side.js(f"__present.slidePointToClient({SLIDE_W + 20}, {SLIDE_H / 2})")
        dragtest.Input(side).click(c["x"], c["y"])
        time.sleep(0.2)
    time.sleep(0.2)


def measure(side) -> None:
    """Values the expected functions need, from the app as it is drawn."""
    reset(side, {"shapes": [RECT]})
    side.js(SELECT)
    time.sleep(0.2)
    r = side.js(f"__present.shapeRect({SID})")
    h = side.js(f"__present.handlePoint({SID}, 'rot')")
    G["rot_rel"] = (h["x"] - (r["x"] + r["w"] / 2), h["y"] - (r["y"] + r["h"] / 2))
    z = side.js("(() => { const s = document.querySelector('.zoom-track').getBoundingClientRect(); return [s.left + 4, s.left + 96]; })()")
    G["zoom_lo"], G["zoom_hi"] = z[0], z[1]
    G["notes_hi"] = side.js("document.querySelector('.center').getBoundingClientRect().height * 0.8")
    reset(side, {"shapes": [{"kind": "table", "box": [100, 120, 600, 120]}]})
    w = side.js("[...document.querySelectorAll('.stage td[data-r=\"0\"]')].map((t) => t.getBoundingClientRect().width)")
    per = side.js("(() => { const a = __present.slidePointToClient(0, 0), b = __present.slidePointToClient(100, 0); return (b.x - a.x) / 100; })()")
    G["col_min"] = 54864 / 12700 * per
    G["col_room"] = w[1] - G["col_min"]
    reset(side, {"shapes": [], "slides": 4})
    a, b = side.js("__present.thumbRect(1)"), side.js("__present.thumbRect(2)")
    G["thumb_pitch"] = b["y"] - a["y"]


def main_paths(families: list[str], names: list[str]) -> None:
    # RAABE_DRAG_RESULTS: another file (runs on an older build, for the numbers before a fix).
    out = Path(__import__("os").environ.get("RAABE_DRAG_RESULTS") or HERE / "results" / "2026-11-09-drag-present-paths.json")
    report: dict = {"app": "raabe-present", "families": families, "anchors": []}
    with (contextlib.nullcontext() if dragtest.INPUT == "cdp" else real_input("drag_present.py --paths (drag-paths agent)")):
        side = PresentSide(port=int(__import__("os").environ.get("RAABE_DRAG_PORT") or (9492 if dragtest.INPUT == "cdp" else 9491)))  # ports of our own (real input, cdp)
        guard = catcher(WINDOW) if dragtest.INPUT != "cdp" else contextlib.nullcontext()
        guard.__enter__()
        try:
            winapi.place(side.hwnd, *WINDOW)
            time.sleep(0.5)
            measure(side)
            z0 = side.js("__present.app.editor.zoom")
            report["measured"] = {k: (list(v) if isinstance(v, tuple) else v) for k, v in G.items()}
            table = anchors()
            sel = [table[n] for n in table if not names or n in names]
            specs = {a.name: spec for a, spec in sel}
            report["anchors"] = run_families(side, [a for a, _ in sel], side.js, lambda a: reset(side, specs[a.name]), families,
                                             zoom_js=lambda z: f"__present.app.setZoom({z0 if z == 100 else z})", reveal=REVEAL,
                                             log=lambda *a: print(*a, flush=True), touch=touch)
            report["zoom_default"] = z0
        finally:
            guard.__exit__(None, None, None)
            side.close()
    report = merge_report(out, report)
    out.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    for a in report["anchors"]:
        print(f"{a['anchor']}: {a['ok']}/{a['of']}")


def main() -> None:
    names = sys.argv[1:]
    if "--cdp" in names:
        names = [w for w in names if w != "--cdp"]
        dragtest.cdp_mode()
    if names and names[0] == "--paths":
        main_paths(names[1].split(","), names[2:])
        return
    report: dict = {"date": datetime.date.today().isoformat(), "anchors": []}
    with real_input("drag_present.py (RaabePresent session)"):
        side = PresentSide()
        try:
            measure(side)
            report["measured"] = {k: (list(v) if isinstance(v, tuple) else v) for k, v in G.items()}
            table = anchors()
            for name in names or list(table):
                anchor, spec = table[name]
                res = []
                for p in paths():
                    reset(side, spec)
                    rec = run_drag(side, anchor, p, side.js)
                    a = analyse(anchor, rec)
                    res.append(a)
                    print(name, p.name, "ok" if a["ok"] else "FAIL", {k: v for k, v in a.items() if k not in ("path", "ok", "v0", "after")}, flush=True)
                    touch()
                eu = esc_and_undo(side, anchor, spec)
                print(name, "esc/undo", eu, flush=True)
                report["anchors"].append({"anchor": name, "ok": sum(1 for r in res if r["ok"]), "of": len(res), **eu, "drags": res})
                touch()
        finally:
            side.close()
    (HERE / "results" / f"{report['date']}-drag-present.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    for a in report["anchors"]:
        print(f"{a['anchor']}: {a['ok']}/{a['of']} esc {a['esc']} undo {a['undo']}")


if __name__ == "__main__":
    main()
