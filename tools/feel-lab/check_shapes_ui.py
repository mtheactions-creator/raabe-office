"""Shapes in RaabeWrite with real input (D-038, package C): Einfügen > Formen, drawing a
shape with the mouse, typing into it, Esc, the yellow handle, a Formenart from the ribbon, a
text box (Textfeld erstellen), an arrow line, saving — and Word reading the result. Takes over
mouse and keyboard. Screenshots go to captures/shapes-ui/ (local only).

    .venv/Scripts/python.exe check_shapes_ui.py
"""
from __future__ import annotations

import json
import re
import tempfile
import time
from pathlib import Path

from feellab import winapi
from feellab.targets import WINDOW_RECT, raabe_write

HERE = Path(__file__).parent
CAP = HERE / "captures" / "shapes-ui"
CAP.mkdir(parents=True, exist_ok=True)
OUT = Path(tempfile.gettempdir()) / "raabe-shapes-ui.docx"

STATE = """(() => { const e = __app.editor(); const s = e.selectedShape();
  const floats = e.doc.main.flatMap((p) => p.floats ?? []);
  return JSON.stringify({ sel: s && { editing: s.editing, w: Math.round(s.obj.width), h: Math.round(s.obj.height), geom: s.obj.pic?.geometry ?? 'rect', adjust: s.obj.shape?.adjust ?? null, style: (s.obj.shape?.styleXml ?? '').slice(0, 60) },
    shapes: floats.filter((f) => f.shape).map((f) => ({ kind: f.shape.kind, geom: f.pic?.geometry ?? 'rect', text: f.shape.story ? (e.doc.stories.get(f.shape.story) ?? []).map((p) => p.text).join('|') : '' })),
    story: e.storyKey, tab: __app.activeRibbonTab() }); })()"""

RECT = """(sel) => { const b = document.querySelector(sel); if (!b) return null; const r = b.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; }"""


def main() -> None:
    out: dict = {}
    t = raabe_write("short", port=9377)
    try:
        winapi.place(t.hwnd, *WINDOW_RECT)
        winapi.activate(t.hwnd)
        t.js("__app.newDoc()")
        time.sleep(0.6)
        t.js("__app.closeOthers && __app.closeOthers()")
        t.js("__probe.focus()")
        scale = winapi.dpi_scale(t.hwnd)
        ox, oy = winapi.client_origin(t.hwnd)
        to_screen = lambda x, y: (ox + x * scale, oy + y * scale)  # noqa: E731
        page = t.js("(() => { const r = document.querySelector('.page').getBoundingClientRect(); return [r.left, r.top]; })()")
        shot = lambda name: winapi.save_region_png(str(CAP / f"{name}.png"), *[int(v) for v in to_screen(page[0] - 40, page[1] - 20)], int(820 * scale), int(560 * scale))  # noqa: E731

        def click_css(sel: str) -> bool:
            p = t.js(f"({RECT})({json.dumps(sel)})")
            if not p:
                print("not found", sel)
                return False
            winapi.click(*to_screen(*p))
            time.sleep(0.6)
            return True

        # 1. Einfügen > Formen > Rechteck: abgerundete Ecken, drawn by dragging.
        click_css("[role=tab][data-id=insert], [role=tab]:nth-child(2)")
        t.js("[...document.querySelectorAll('[role=tab]')].find((x) => x.textContent.trim() === 'Einfügen')?.click()")
        time.sleep(0.5)
        click_css("button[aria-label='Formen']")
        click_css("button[aria-label='Rechteck: abgerundete Ecken']")
        sx, sy = to_screen(page[0] + 150, page[1] + 160)
        winapi.drag(sx, sy, sx + 220 * scale, sy + 110 * scale, steps=14)
        time.sleep(0.8)
        out["drawn"] = json.loads(t.js(STATE))
        print("drawn", out["drawn"], flush=True)
        shot("1-drawn")
        # 2. Typing goes into the shape; Esc selects the shape again.
        winapi.type_text("Hallo Welt")
        time.sleep(0.5)
        out["typed"] = json.loads(t.js(STATE))
        shot("2-typed")
        winapi.press("esc")
        time.sleep(0.4)
        out["esc"] = json.loads(t.js(STATE))
        # 3. The yellow handle (corner radius) dragged to the right.
        h = t.js("(() => { const els = [...document.querySelectorAll('.page div')].filter((d) => d.style.background.includes('255, 210, 0') || d.style.background === 'rgb(255, 210, 0)'); if (!els.length) return null; const r = els[0].getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; })()")
        if h:
            hx, hy = to_screen(*h)
            winapi.drag(hx, hy, hx + 30 * scale, hy, steps=10)
            time.sleep(0.5)
        out["adjusted"] = json.loads(t.js(STATE))
        shot("3-adjusted")
        # 4. A Formenart from the ribbon (Formformat): the fourth tile.
        # The tiles shown at this window width (the others are in the full gallery).
        tiles = t.js("(() => [...document.querySelectorAll('.rgallery.tiles:not(.shape-tiles-ribbon) .item')].filter((b) => b.getBoundingClientRect().width > 0).map((b) => { const r = b.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2, b.getAttribute('data-tip')]; }))()")
        if tiles:
            pick = tiles[min(2, len(tiles) - 1)]
            winapi.click(*to_screen(pick[0], pick[1]))
            time.sleep(0.6)
            out["style_tip"] = pick[2]
        out["styled"] = json.loads(t.js(STATE))
        shot("4-styled")
        # 5. Textfeld erstellen (Formformat), drawn, with text.
        click_css("button[aria-label='Textfeld erstellen'], [data-cmd='insert.textBoxDraw']")
        sx, sy = to_screen(page[0] + 420, page[1] + 160)
        winapi.drag(sx, sy, sx + 200 * scale, sy + 80 * scale, steps=12)
        time.sleep(0.6)
        winapi.type_text("Ein Textfeld")
        time.sleep(0.4)
        winapi.press("esc")
        winapi.press("esc")
        out["textbox"] = json.loads(t.js(STATE))
        # 6. Einfügen > Formen > Linienpfeil, drawn.
        t.js("[...document.querySelectorAll('[role=tab]')].find((x) => x.textContent.trim() === 'Einfügen')?.click()")
        time.sleep(0.4)
        click_css("button[aria-label='Formen']")
        click_css("button[aria-label='Linienpfeil']")
        sx, sy = to_screen(page[0] + 150, page[1] + 330)
        winapi.drag(sx, sy, sx + 260 * scale, sy + 40 * scale, steps=12)
        time.sleep(0.6)
        out["arrow"] = json.loads(t.js(STATE))
        shot("5-all")
        OUT.unlink(missing_ok=True)
        t.js(f"__app.saveTo({str(OUT)!r})")
        time.sleep(1.5)
    finally:
        t.close()
    # 7. Word reads it.
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(str(OUT), ReadOnly=True)
        x = d.Content.WordOpenXML
        shapes = []
        for i in range(1, d.Shapes.Count + 1):
            s = d.Shapes(i)
            try:
                text = s.TextFrame.TextRange.Text.strip() if s.TextFrame.HasText else ""
            except Exception:  # noqa: BLE001
                text = ""
            shapes.append({"name": s.Name, "type": s.Type, "auto": s.AutoShapeType, "w": round(s.Width, 1), "h": round(s.Height, 1), "text": text})
        out["word"] = {"shapes": shapes, "prst": re.findall(r'<a:prstGeom prst="([^"]+)"', x), "tailEnd": re.findall(r"<a:tailEnd[^>]*/>", x),
                       "theme": "theme1" in str([p.Name for p in []]), "adj": re.findall(r'<a:gd name="adj" fmla="val (\d+)"/>', x)}
        d.Close(0)
    finally:
        app.Quit(0)
    print(json.dumps(out, indent=1, ensure_ascii=False))


if __name__ == "__main__":
    main()
