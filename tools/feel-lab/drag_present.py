"""The drag test in RaabePresent (D-080 C, D-110): every anchor dragged along feellab/dragtest.py's paths with real input
and checked for jumps (follow, jump, direction, release); then per anchor Esc during a drag and Strg+Z after one, which
must bring back the value from before. Values are read from the app's own elements in client CSS px (a frame's left
edge, a size, a column's width, a pane's size, the thumbnails' drop line) or degrees (rotation). Takes mouse and keyboard
(the real-input lock, feellab/inputlock.py).

    .venv/Scripts/python.exe drag_present.py [anchor ...]

Result: results/<date>-drag-present.json.
"""
from __future__ import annotations

import datetime
import json
import math
import sys
import time
from pathlib import Path

from feellab import winapi
from feellab.dragtest import Anchor, Path as DragPath, analyse, paths, run_drag, screen_point
from feellab.inputlock import real_input, touch
from feellab.present_catalogue import Case, PresentSide

HERE = Path(__file__).parent
SID = "__present.state().shapes.filter((s) => !s.ph)[0].id"
RECT = {"kind": "rect", "box": [300, 150, 200, 120]}
G: dict = {}  # measured once per run: the shape's height in px, the table's next column, the zoom slider's track …


def shape(name: str, kind: dict, prop: str, expect, start: str, setup: str = "", tol: float = 2.0, snap: float = 5.0) -> tuple[Anchor, dict]:
    probe = f"(() => {{ const r = __present.shapeRect({SID}); return r ? r.{prop} : null; }})()"
    return Anchor(name, start=start, probe=probe, expect=expect, tol=tol, snap=snap, setup=setup), {"shapes": [kind]}


SELECT = f"__present.app.editor.select([{SID}])"


def handle(k: str, prop: str, sign: int, axis: str) -> tuple[Anchor, dict]:
    # Past the opposite side the shape flips (as in PowerPoint): the size grows again.
    exp = (lambda dx, dy, v0: abs(v0 + sign * (dx if axis == "x" else dy)))
    return shape(f"handle {k}", RECT, prop, exp, start=f"__present.handlePoint({SID}, '{k}')", setup=SELECT, snap=0.0)


def rotation() -> tuple[Anchor, dict]:
    probe = (f"(() => {{ const e = document.querySelector('.stage .shp[data-id=\"' + {SID} + '\"]'); const m = /rotate\\(([-\\d.e]+)deg/.exec(e?.style.transform ?? '');"
             " const a = m ? Number(m[1]) : 0; return a > 180 ? a - 360 : a; })()")

    def expect(dx, dy, v0):
        rx, ry = G["rot_rel"]  # the handle relative to the shape's centre, px
        a = math.degrees(math.atan2(ry + dy, rx + dx)) + 90
        return a - 360 if a > 180 else a
    # The angle may stay behind by the drag threshold (4 px of the first movement, about 2.7 degrees at this handle's
    # radius): the threshold PowerPoint's catalogue cases showed, not a jump (D-110).
    return Anchor("rotation handle", start=f"__present.handlePoint({SID}, 'rot')", probe=probe, expect=expect, tol=1.5, snap=2.8, setup=SELECT), {"shapes": [RECT]}


def table_column() -> tuple[Anchor, dict]:
    probe = "(() => document.querySelector('.stage td[data-r=\"0\"][data-c=\"0\"]')?.getBoundingClientRect().width ?? null)()"
    start = "(() => { const r = document.querySelector('.stage td[data-r=\"0\"][data-c=\"0\"]').getBoundingClientRect(); return { x: r.right, y: r.top + r.height / 2 }; })()"
    return Anchor("table column border", start=start, probe=probe, expect=lambda dx, dy, v0: min(max(G["col_min"], v0 + dx), v0 + G["col_room"]), tol=1.5), \
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
                  setup="window.__dragSlide = __present.app.deck.slides[1].path"), \
        {"shapes": [], "slides": 4}


def zoom_slider() -> tuple[Anchor, dict]:
    # RaabePresent's own slider (D-110): the thumb's centre, read from the page.
    thumb = "(() => { const r = document.querySelector('.zoom-thumb').getBoundingClientRect(); return r.left + r.width / 2; })()"
    start = "(() => { const r = document.querySelector('.zoom-thumb').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()"
    return Anchor("zoom slider", start=start, probe=thumb, expect=lambda dx, dy, v0: min(max(G["zoom_lo"], v0 + dx), G["zoom_hi"]), tol=1.5), {"shapes": [RECT]}


def pane_splitter() -> tuple[Anchor, dict]:
    probe = "document.querySelector('.present-main').firstElementChild.getBoundingClientRect().width"
    start = "(() => { const r = document.querySelector('.present-main > .splitter').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()"
    return Anchor("thumbnail pane splitter", start=start, probe=probe, expect=lambda dx, dy, v0: min(max(90.0, v0 + dx), 420.0), tol=1.5), {"shapes": [RECT]}


def notes_splitter() -> tuple[Anchor, dict]:
    probe = "document.querySelector('.notes-wrap').getBoundingClientRect().height"
    start = "(() => { const r = document.querySelector('.notes-grip').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()"
    return Anchor("notes splitter", start=start, probe=probe, expect=lambda dx, dy, v0: min(max(24.0, v0 - dy), G["notes_hi"]), tol=1.5), {"shapes": [RECT]}


def anchors() -> dict[str, tuple[Anchor, dict]]:
    centre = f"(() => {{ const r = __present.shapeRect({SID}); return {{ x: r.x + r.w * 0.55, y: r.y + r.h * 0.5 }}; }})()"
    top_edge = f"(() => {{ const r = __present.shapeRect({SID}); return {{ x: r.x + r.w * 0.3, y: r.y }}; }})()"
    pic = ("(async () => { const c = document.createElement('canvas'); c.width = 240; c.height = 150; const g = c.getContext('2d'); g.fillStyle = '#3a7'; g.fillRect(0, 0, 240, 150);"
           " const b = await new Promise((r) => c.toBlob(r)); __present.app.insertPictureBytes(new Uint8Array(await b.arrayBuffer()), 'image/png'); __present.app.editor.select([]); })()")
    out = [
        shape("move shape", RECT, "x", lambda dx, dy, v0: v0 + dx, centre),
        shape("move text box", {"kind": "textbox", "box": [300, 150, 240, 40], "text": "Ein Textfeld"}, "x", lambda dx, dy, v0: v0 + dx, top_edge, setup=SELECT),
        shape("move picture", {"kind": "rect", "box": [20, 20, 10, 10]}, "x", lambda dx, dy, v0: v0 + dx,
              f"(() => {{ const v = __present.state().shapes.filter((s) => !s.ph).pop(); const r = __present.shapeRect(v.id); return {{ x: r.x + r.w * 0.5, y: r.y + r.h * 0.5 }}; }})()", setup=pic),
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


def main() -> None:
    names = sys.argv[1:]
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
