"""Where PowerPoint's rotation handle is (Step 8, D-072): the behaviour catalogue's rotate cases missed it
with an estimated 27 device pixels above the frame. This probe selects a rectangle in a PowerPoint of our
own (real input, under the lock, only when no POWERPNT.EXE is running) and drags from several distances
above the top edge's middle; the first that rotates the shape is the handle.

    .venv/Scripts/python.exe probe_powerpoint_rothandle.py

Result: results/<date>-powerpoint-rothandle.json (distances tried and the rotation each gave).
"""
from __future__ import annotations

import datetime
import json
import time
from pathlib import Path

from catalogue_present import LOCK, RECT, take_lock
from feellab import winapi
from feellab.present_catalogue import Case, PowerPointSide

HERE = Path(__file__).parent


def main() -> None:
    if not take_lock("step8-rothandle"):
        raise SystemExit("real-input lock held by another session: try later")
    side = None
    tried: list[dict] = []
    try:
        side = PowerPointSide()
        for off in (14, 24, 36, 38, 40, 42, 44, 46, 48, 50, 53, 56):
            case = Case(f"rot {off}", "rotate", [RECT], [("click", (0, .5, 0))])
            side.reset(case)
            side.run(case)
            x, y, w, _ = side.shape_box(0)
            px, py = side.slide_point(x + w / 2, y)
            sx, sy = px, py - off
            ex, ey = sx + 80, sy + 40
            winapi.move(sx, sy)
            time.sleep(0.05)
            winapi.mouse_down(sx, sy)
            time.sleep(0.1)
            for i in range(1, 21):
                winapi.move(sx + (ex - sx) * i / 20, sy + (ey - sy) * i / 20)
                time.sleep(0.015)
            time.sleep(0.15)
            winapi.mouse_up(ex, ey)
            time.sleep(0.4)
            rot = float(side._shape(0).Rotation)
            tried.append({"offset_px": off, "rotation": round(rot, 1)})
            print(off, rot, flush=True)
            LOCK.touch()
    finally:
        if side:
            side.close()
        LOCK.unlink(missing_ok=True)
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-powerpoint-rothandle.json"
    out.write_text(json.dumps({"date": datetime.date.today().isoformat(), "zoom": 66, "scaling": "150 %", "tried": tried}, indent=1), encoding="utf-8", newline=chr(10))
    print("->", out.name)


if __name__ == "__main__":
    main()
