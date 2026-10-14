"""Pictures in RaabeWrite with real input (D-038, package C): pasting a picture from the
clipboard (Strg+V), the reported PNG (if present), cropping with the crop marks, the rotation
handle, a picture style from the ribbon gallery, and the result read by Word. Uses the
clipboard; takes over mouse and keyboard. Screenshots go to captures/pictures/ (local only).

    .venv/Scripts/python.exe check_pictures.py
"""
from __future__ import annotations

import json
import tempfile
import time
from pathlib import Path

import win32clipboard

from feellab import winapi
from feellab.targets import WINDOW_RECT, raabe_write
from probe_word_picture import PNG, make_png

HERE = Path(__file__).parent
CAP = HERE / "captures" / "pictures"
CAP.mkdir(parents=True, exist_ok=True)
OUT = Path(tempfile.gettempdir()) / "raabe-pictures.docx"
REPORTED = Path.home() / "Downloads" / "logo.5f159bae.png"

STATE = """(() => { const e = __app.editor(); const s = e.selectedImage();
  return JSON.stringify({ sel: s && { kind: s.kind, w: Math.round(s.obj.width * 10) / 10, h: Math.round(s.obj.height * 10) / 10, pic: s.obj.pic },
    images: e.doc.paras.reduce((n, p) => n + (p.objects ?? []).filter((o) => o.kind === 'image').length, 0), box: e.imageBox() }); })()"""


def put_png(path: Path) -> None:
    data = path.read_bytes()
    fmt = win32clipboard.RegisterClipboardFormat("PNG")
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(fmt, data)
    finally:
        win32clipboard.CloseClipboard()


def main() -> None:
    make_png(PNG)
    out: dict = {}
    t = raabe_write("short", port=9376)
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

        # 1. Strg+V with a PNG on the clipboard.
        put_png(PNG)
        winapi.press("ctrl+v")
        time.sleep(1.2)
        out["paste"] = json.loads(t.js(STATE))
        print("paste", out["paste"], flush=True)
        # 2. The reported PNG (500 dpi): Word inserts it at 3.26 x 3.18 cm.
        if REPORTED.exists():
            t.js("(() => { const e = __app.editor(); e.select(e.doc.end(), e.doc.end()); })()")
            winapi.press("enter")
            t.js(f"__app.insertPicture({str(REPORTED)!r})")
            time.sleep(1.2)
            out["reported"] = json.loads(t.js(STATE))
            out["reported_word_size"] = [92.3, 90.1]
        # Select the first picture again.
        t.js("(() => { const e = __app.editor(); e.select({ p: 0, o: 0 }, { p: 0, o: 1 }); })()")
        time.sleep(0.5)
        # 3. Zuschneiden: drag the right crop mark 40 px to the left.
        t.js("__app.editor().startCrop()")
        time.sleep(0.5)
        b = json.loads(t.js(STATE))["box"]
        page = t.js("(() => { const r = document.querySelector('.page').getBoundingClientRect(); return [r.left, r.top]; })()")
        z = t.js("__app.editor().zoom")
        rx, ry = page[0] + (b["x"] + b["w"]) * z, page[1] + (b["y"] + b["h"] / 2) * z
        winapi.save_region_png(str(CAP / "crop-mode.png"), *[int(v) for v in to_screen(page[0], page[1])], int(900 * scale), int(500 * scale))
        sx, sy = to_screen(rx, ry)
        winapi.drag(sx, sy, sx - 40 * scale, sy, steps=12)
        time.sleep(0.4)
        winapi.press("esc")
        time.sleep(0.4)
        out["after_crop"] = json.loads(t.js(STATE))
        # 4. The rotation handle: drag from above the centre to the right (about 45°).
        b = out["after_crop"]["box"]
        cx, cy = page[0] + (b["x"] + b["w"] / 2) * z, page[1] + (b["y"] - 17) * z
        sx, sy = to_screen(cx, cy)
        winapi.drag(sx, sy, sx + (b["h"] / 2 + 17) * z * scale, sy + (b["h"] / 2 + 17) * z * scale, steps=15)
        time.sleep(0.4)
        out["after_rotate"] = json.loads(t.js(STATE))
        winapi.save_region_png(str(CAP / "rotated.png"), *[int(v) for v in to_screen(page[0], page[1])], int(900 * scale), int(500 * scale))
        t.js("(() => { const e = __app.editor(); e.setPictureFormat({ rotation: undefined }); })()")
        # 5. A picture style from the ribbon gallery (the first tile).
        tile = t.js("(() => { const b = document.querySelector('.rgallery.tiles .item'); if (!b) return null; const r = b.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; })()")
        if tile:
            winapi.click(*to_screen(*tile))
            time.sleep(0.6)
        out["after_style"] = json.loads(t.js(STATE))
        winapi.save_region_png(str(CAP / "styled.png"), *[int(v) for v in to_screen(page[0], page[1])], int(900 * scale), int(500 * scale))
        OUT.unlink(missing_ok=True)
        t.js(f"__app.saveTo({str(OUT)!r})")
        time.sleep(1.5)
    finally:
        t.close()
    # 6. Word reads the result.
    import win32com.client
    import re
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(str(OUT), ReadOnly=True)
        x = d.Content.WordOpenXML
        out["word"] = {
            "pictures": d.InlineShapes.Count,
            "first": {"w": round(d.InlineShapes(1).Width, 1), "h": round(d.InlineShapes(1).Height, 1)},
            "srcRect": (re.search(r"<a:srcRect[^>]*/>", x) or [None])[0],
            "ln": (re.search(r'<a:ln w="\d+"', x) or [None])[0],
        }
        d.Close(0)
    finally:
        app.Quit(0)
    print(json.dumps(out, indent=1, ensure_ascii=False))


if __name__ == "__main__":
    main()
