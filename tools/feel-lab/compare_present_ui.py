"""UI screenshots of PowerPoint and RaabePresent side by side (D-070): the main tabs of the ribbon over a
new presentation, the same window size, one PNG per tab in captures/ui-compare/ (not committed: it shows
Office's look). For looking at, not for numbers.

    .venv/Scripts/python.exe compare_present_ui.py

Real input (keytips in PowerPoint): only with the real-input lock, and only when no POWERPNT.EXE runs.
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

from PIL import Image

from catalogue_present import LOCK, take_lock
from feellab import winapi
from feellab.powerpoint import PowerPointTarget, owner_powerpoint_running
from feellab.present_catalogue import PresentSide

HERE = Path(__file__).parent
OUT = HERE / "captures" / "ui-compare"
# Tab, PowerPoint's keytip, our ribbon's tab id.
TABS = [("Start", "r", "start"), ("Einfügen", "i", "insert"), ("Entwurf", "u+2", "design"), ("Animationen", "v", "animations"),
        ("Bildschirmpräsentation", "l", "slideshow"), ("Überprüfen", "p+2", "review"), ("Ansicht", "f", "view")]


def grab(hwnd: int, path: Path) -> None:
    x, y, r, b = winapi.window_rect(hwnd)
    winapi.save_region_png(str(path), x, y, r - x, b - y)


def main() -> None:
    if owner_powerpoint_running():
        raise SystemExit("POWERPNT.EXE is running (the owner's): not now")
    if not take_lock("step8-ui-compare"):
        raise SystemExit("real-input lock held")
    OUT.mkdir(parents=True, exist_ok=True)
    try:
        ppt = PowerPointTarget(visible=True)
        try:
            ppt.new_presentation(with_window=True, name="ui")
            time.sleep(2)
            for name, keys, _ in TABS:
                winapi.activate(ppt.hwnd)
                winapi.press("alt"); time.sleep(0.4)
                for k in keys.split("+"):
                    winapi.press(k); time.sleep(0.3)
                winapi.press("esc"); time.sleep(0.2); winapi.press("esc"); time.sleep(0.6)
                grab(ppt.hwnd, OUT / f"ppt-{name}.png")
                LOCK.touch()
        finally:
            print("PowerPoint:", ppt.close(kill_if_unreachable=True))
        side = PresentSide()
        try:
            side.js("__present.newDoc()")
            time.sleep(1)
            for name, _, tab in TABS:
                side.js(f"__present.app.ribbon.select({tab!r})")
                time.sleep(0.6)
                grab(side.hwnd, OUT / f"present-{name}.png")
        finally:
            side.close()
        for name, _, _ in TABS:
            a, b = Image.open(OUT / f"ppt-{name}.png"), Image.open(OUT / f"present-{name}.png")
            h = 520  # the ribbon and the top of the slide
            out = Image.new("RGB", (max(a.width, b.width), 2 * h + 10), "white")
            out.paste(a.crop((0, 0, a.width, h)), (0, 0))
            out.paste(b.crop((0, 0, b.width, h)), (0, h + 10))
            out.save(OUT / f"compare-{name}.png")
        print("written", OUT)
    finally:
        LOCK.unlink(missing_ok=True)


if __name__ == "__main__":
    sys.exit(main())
