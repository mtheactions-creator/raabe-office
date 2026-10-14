"""Screenshots of RaabeWrite's Bildformat tab and its galleries, for comparison with Word's
(probe_word_picture.py shots). Inserts a test picture through the test API, opens each
gallery with a real click. Output: captures/raabe-picture/ (local only).

    .venv/Scripts/python.exe capture_picture_tab.py
"""
from __future__ import annotations

import time
from pathlib import Path

from feellab import winapi
from feellab.targets import WINDOW_RECT, raabe_write
from probe_word_picture import PNG, make_png

HERE = Path(__file__).parent
CAP = HERE / "captures" / "raabe-picture"
CAP.mkdir(parents=True, exist_ok=True)

FIND = """(label) => { const b = [...document.querySelectorAll('.ribbon-card button')].find((x) => (x.getAttribute('aria-label') || x.textContent || '').replace(/\\s+/g, ' ').trim().startsWith(label));
  if (!b) return null; const r = b.getBoundingClientRect(); return [r.left, r.top, r.width, r.height]; }"""


def main() -> None:
    make_png(PNG)
    t = raabe_write("short", port=9374)
    try:
        winapi.place(t.hwnd, *WINDOW_RECT)
        winapi.activate(t.hwnd)
        t.js("__app.newDoc()")
        time.sleep(0.6)
        t.js("__app.closeOthers && __app.closeOthers()")
        t.js("__probe.focus()")
        t.js(f"__app.insertPicture({str(PNG)!r})")
        time.sleep(1.0)
        scale = winapi.dpi_scale(t.hwnd)
        ox, oy = winapi.client_origin(t.hwnd)
        winapi.save_region_png(str(CAP / "tab.png"), ox, oy, int(2000 * scale), int(200 * scale))

        def click(label: str, bottom: bool = False) -> bool:
            r = t.js(f"({FIND})({label!r})")
            if not r:
                print("not found", label)
                return False
            x, y, w, h = r
            winapi.click(ox + (x + w / 2) * scale, oy + (y + (h - 6 if bottom else h / 2)) * scale)
            time.sleep(1.0)
            return True

        for label, name in (("Korrekturen", "corrections"), ("Farbe", "color"), ("Transparenz", "transparency"),
                            ("Schnellformatvorlagen", "styles"), ("Bildeffekte", "effects"), ("Position", "position"),
                            ("Zuschneiden", "crop"), ("Drehen", "rotate"), ("Bildrahmen", "border")):
            if click(label, bottom=label in ("Zuschneiden", "Bildrahmen")):
                winapi.save_region_png(str(CAP / f"{name}.png"), ox, oy, int(1700 * scale), int(1150 * scale))
                winapi.press("esc")
                time.sleep(0.4)
                t.js("__probe.focus()")
        # A style applied, then the document with the picture.
        t.js("""(() => { const e = __app.editor(); e.setPictureFormat({ line: { color: 'FFFFFF', width: 7 }, shadow: { kind: 'outer', color: '000000', transparency: 0.6, blur: 4.3, dist: 1.4, dir: 90 } }); })()""")
        time.sleep(0.6)
        winapi.save_region_png(str(CAP / "styled.png"), ox, oy, int(1400 * scale), int(800 * scale))
    finally:
        t.close()
    print("saved to", CAP)


if __name__ == "__main__":
    main()
