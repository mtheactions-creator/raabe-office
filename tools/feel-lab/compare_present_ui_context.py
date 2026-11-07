"""The contextual tabs of PowerPoint and RaabePresent side by side (D-072): Formformat with a shape selected,
Bildformat with a picture, Tabellenentwurf and Tabellenlayout with a table. Same rules and output folder as
compare_present_ui.py (captures/ui-compare, not committed).

    .venv/Scripts/python.exe compare_present_ui_context.py
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

from PIL import Image

from catalogue_present import LOCK, take_lock
from compare_present_ui import OUT, grab
from feellab import winapi
from feellab.powerpoint import PowerPointTarget, owner_powerpoint_running
from feellab.present_catalogue import PresentSide

HERE = Path(__file__).parent
PNG = HERE.parent.parent / "packages" / "ooxml" / "test" / "fixtures" / "pptx-png" / "pptx-shapes-1.png"
# Name, what to put on the slide, PowerPoint's keytip, our tab id.
CASES = [("Formformat", "shape", "j+d", "shapeFormat"), ("Bildformat", "picture", "j+p", "pictureFormat"),
         ("Tabellenentwurf", "table", "j+t", "tableDesign"), ("Tabellenlayout", "table", "j+l", "tableLayout")]


def main() -> None:
    if owner_powerpoint_running():
        raise SystemExit("POWERPNT.EXE is running (the owner's): not now")
    if not take_lock("step8-ui-compare-context"):
        raise SystemExit("real-input lock held")
    OUT.mkdir(parents=True, exist_ok=True)
    try:
        ppt = PowerPointTarget(visible=True)
        try:
            pres = ppt.new_presentation(with_window=True, name="ui")
            slide = pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts(7)) if pres.Slides.Count == 0 else pres.Slides(1)
            ppt.app.ActiveWindow.View.GotoSlide(1)
            for s in list(slide.Shapes):
                s.Delete()
            time.sleep(1)
            for name, what, keys, _ in CASES:
                for s in list(slide.Shapes):
                    s.Delete()
                shp = (slide.Shapes.AddShape(1, 200, 120, 240, 140) if what == "shape" else
                       slide.Shapes.AddPicture(str(PNG), 0, -1, 200, 120, 320, 180) if what == "picture" else
                       slide.Shapes.AddTable(3, 3, 120, 120, 600, 120))
                winapi.activate(ppt.hwnd)
                shp.Select()
                time.sleep(0.6)
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
            for name, what, _, tab in CASES:
                spec = {"layout": "Leer", "shapes": [{"kind": "rect", "box": [200, 120, 240, 140]}] if what == "shape" else
                        [{"kind": "table", "box": [120, 120, 600, 120]}] if what == "table" else []}
                side.js(f"__present.setupCase({__import__('json').dumps(spec)})")
                if what == "picture":
                    data = list(PNG.read_bytes())
                    side.js(f"(async () => {{ __present.app.insertPictureBytes(new Uint8Array({data}), 'image/png'); await new Promise(r => setTimeout(r, 400)); return 1; }})()")
                side.js("(() => { const a = __present.app; const v = a.deck.view(a.slidePath()).shapes; a.editor.select([v[v.length - 1].id]); return 1; })()")
                time.sleep(0.4)
                side.js(f"__present.app.ribbon.select({tab!r})")
                time.sleep(0.6)
                grab(side.hwnd, OUT / f"present-{name}.png")
        finally:
            side.close()
        for name, *_ in CASES:
            a, b = Image.open(OUT / f"ppt-{name}.png"), Image.open(OUT / f"present-{name}.png")
            h = 300
            out = Image.new("RGB", (max(a.width, b.width), 2 * h + 10), "white")
            out.paste(a.crop((0, 0, a.width, h)), (0, 0))
            out.paste(b.crop((0, 0, b.width, h)), (0, h + 10))
            out.save(OUT / f"compare-{name}.png")
        print("written", OUT)
    finally:
        LOCK.unlink(missing_ok=True)


if __name__ == "__main__":
    sys.exit(main())
