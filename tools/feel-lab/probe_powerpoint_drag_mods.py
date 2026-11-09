"""What PowerPoint does with modifier keys pressed mid-drag and with the pointer leaving the window (D-127, the drag
test's reference for RaabePresent): our own PowerPoint (only when no POWERPNT.EXE runs, D-070), real drags on a
rectangle (body, the handles se, e, s, nw, n, w, the rotation handle), a text box's edge and a table's column border,
each along feellab/officedrag.py's variants (plain; Umschalt, Strg, Alt let go mid-drag or held over the release; out of
the window and back; the button let go outside). Read back through COM: the shape's place, size and rotation, the
column widths and the number of shapes (copies).
Takes over mouse and keyboard (the real-input lock). Result: results/2026-11-09-powerpoint-drag-mods.json.
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

from feellab import winapi
from feellab.inputlock import real_input, touch
from feellab import officedrag
from feellab.catcher import catcher
from feellab.officedrag import VARIANTS, drag
from feellab.powerpoint import owner_powerpoint_running
from feellab.present_catalogue import Case, PowerPointSide

HERE = Path(__file__).parent
RECT = {"kind": "rect", "box": [300, 150, 200, 120]}
TEXTBOX = {"kind": "textbox", "box": [300, 150, 240, 40], "text": "Ein Textfeld"}
TABLE = {"kind": "table", "box": [100, 120, 600, 120]}
WIN = (0, 0, 2000, 1700)


def main() -> None:
    want = sys.argv[1:]
    if owner_powerpoint_running():
        raise SystemExit("POWERPNT.EXE is running (not ours): skipped")
    res: list[dict] = []
    with real_input("probe_powerpoint_drag_mods.py (drag-paths agent)"):
        side = PowerPointSide()
        side.reset(Case("mods", "mods", [RECT], []))
        winapi.place(side.t.hwnd, *WIN)  # room below the window; placed once, before the cases (PowerPoint fits the slide again)
        time.sleep(1.0)
        guard = catcher(WIN)
        guard.__enter__()
        try:
            def ours() -> bool:
                fg = winapi.foreground()
                return winapi.window_pid(fg) == winapi.window_pid(side.t.hwnd)

            def setup(shape: dict, select: bool):
                def f():
                    side.reset(Case("mods", "mods", [shape], []))
                    time.sleep(0.3)
                    if select:
                        side._shape(0).Select()
                        time.sleep(0.3)
                return f

            def geo():
                s = side._shape(0)
                return [round(float(s.Left), 1), round(float(s.Top), 1), round(float(s.Width), 1), round(float(s.Height), 1), round(float(s.Rotation), 1)]

            def cols():
                t = side._shape(0).Table
                return [round(float(t.Columns(1).Width), 1), round(float(t.Columns(2).Width), 1)]

            def count():
                return int(side.app.ActiveWindow.View.Slide.Shapes.Count)

            def centre():
                x, y, w, h = side.shape_box(0)
                return side.slide_point(x + w * 0.55, y + h * 0.5)

            def top_edge():
                x, y, w, h = side.shape_box(0)
                return side.slide_point(x + w * 0.3, y)

            def col_border():
                s = side._shape(0)
                t = s.Table
                return side.slide_point(float(s.Left) + float(t.Columns(1).Width), float(s.Top) + float(t.Rows(1).Height) / 2)

            cases = [
                ("move shape", setup(RECT, False), centre, "right", geo),
                ("move text box", setup(TEXTBOX, True), top_edge, "right", geo),
                *[(f"handle {k}", setup(RECT, True), (lambda k=k: side.handle_point(0, k)), "right" if k in ("se", "e", "nw", "w") else "down", geo)
                  for k in ("se", "e", "s", "nw", "n", "w")],
                ("rotation handle", setup(RECT, True), lambda: side.handle_point(0, "rot"), "right", geo),
                ("table column border", setup(TABLE, False), col_border, "right", cols),
            ]
            for name, prep, point, direction, read in cases:
                if want and not any(w == name for w in want):
                    continue
                out: dict = {"case": name, "out": direction}
                try:
                    for v in VARIANTS:
                        def attempt(v=v) -> dict:
                            prep()
                            p = point()
                            officedrag.front(side.t.hwnd, ours)
                            before, n0 = read(), count()
                            drag(p, v, ours, window=winapi.window_rect(side.t.hwnd), direction=direction)
                            after, n1 = read(), count()
                            return {"before": before, "after": after, "copies": n1 - n0}
                        out[v] = officedrag.run_variant(attempt, ours)
                        officedrag.settle(ours)
                        time.sleep(0.2)
                        touch()
                except Exception as e:  # noqa: BLE001
                    out["error"] = str(e)[:200]
                res.append(out)
                print(json.dumps(out, ensure_ascii=False), flush=True)
                touch()
        finally:
            guard.__exit__(None, None, None)
            side.close()
    path = HERE / "results" / "2026-11-09-powerpoint-drag-mods.json"
    if want and path.exists():
        old = json.loads(path.read_text(encoding="utf-8")).get("powerpoint", [])
        names = {r["case"] for r in res}
        res = [r for r in old if r["case"] not in names] + res
    path.write_text(json.dumps({"powerpoint": res}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
