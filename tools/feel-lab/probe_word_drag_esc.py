"""What Esc does during a mouse drag in Word, and Strg+Z after one (D-096, the drag test's reference): our own visible
Word, real drags on a table's column border and row border, an inline picture's corner handle, a floating picture
(body, side handle, rotation handle), each once with Esc pressed before the button goes up and once without (the
control), then Strg+Z after the control drag. Read back through COM. The ruler is not probed (its markers' places are
not in the object model).
Takes over mouse and keyboard (the real-input lock). Result: results/<date>-word-drag-esc.json.
"""
from __future__ import annotations

import datetime
import json
import sys
import tempfile
import time
from pathlib import Path

from PIL import Image

from feellab import winapi
from feellab.inputlock import real_input, touch
from feellab.targets import WINDOW_RECT

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir())
PNG = TMP / "raabe-drag-esc.png"
TEXT = "Ein Absatz mit etwas Text, damit das Bild umflossen werden kann. " * 4


def main() -> None:
    import win32com.client
    sys.stdout.reconfigure(encoding="utf-8")
    Image.new("RGB", (240, 160), (70, 120, 200)).save(PNG)
    res: list[dict] = []
    with real_input("probe_word_drag_esc.py"):
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = True
        pid = None
        try:
            d = app.Documents.Add()
            win = d.Windows(1)
            win.View.Type = 3
            win.View.Zoom.Percentage = 100
            winapi.place(win.Hwnd, *WINDOW_RECT)
            pid = winapi.window_pid(win.Hwnd)

            def ours() -> bool:
                fg = winapi.foreground()
                return winapi.window_pid(fg) == pid

            def drag(p, delta, esc: bool) -> None:
                winapi.activate(win.Hwnd)
                if not ours():
                    raise RuntimeError("Word is not in front")
                winapi.move(*p)
                time.sleep(0.3)
                winapi.mouse_down(*p)
                time.sleep(0.08)
                for i in range(1, 31):
                    winapi.move(p[0] + delta[0] * i / 30, p[1] + delta[1] * i / 30)
                    time.sleep(0.015)
                time.sleep(0.2)
                if esc:
                    winapi.press("esc")
                    time.sleep(0.2)
                winapi.mouse_up(p[0] + delta[0], p[1] + delta[1])
                time.sleep(0.6)

            def setup() -> None:
                d.Content.Delete()
                d.Content.Text = "\r".join([TEXT, "", TEXT, TEXT, ""])
                t = d.Tables.Add(d.Paragraphs(2).Range, 2, 2, 1, 0)
                for i in (1, 2):
                    for j in (1, 2):
                        t.Cell(i, j).Range.Text = f"Z{i}S{j}"
                d.InlineShapes.AddPicture(str(PNG), False, True, d.Paragraphs(d.Paragraphs.Count).Range)
                s = d.Shapes.AddPicture(str(PNG), False, True, 300, 20, 120, 80, d.Paragraphs(1).Range)
                s.Name = "Probe"
                s.WrapFormat.Type = 0
                d.UndoClear()
                time.sleep(0.4)

            def rect(obj) -> tuple[float, float, float, float]:
                left, top, w, h = win.GetPoint(0, 0, 0, 0, obj)
                return float(left), float(top), float(w), float(h)

            tbl = lambda: d.Tables(1)
            fl = lambda: d.Shapes("Probe")
            il = lambda: d.InlineShapes(1)

            def col_border():
                c = tbl().Cell(1, 2).Range
                l, t, w, h = rect(c)
                return l - 5.4 * 2, t + h / 2  # the cell's left edge: its text minus the 5,4 pt padding (2 px per pt)

            def row_border():
                a = rect(tbl().Cell(1, 1).Range)
                b = rect(tbl().Cell(2, 1).Range)
                return a[0] + 20, (a[1] + a[3] + b[1]) / 2

            def sel_inline():
                il().Select()
                time.sleep(0.3)
                l, t, w, h = rect(il().Range)
                return l + w, t + h

            def sel_float(where: str):
                def f():
                    fl().Select()
                    time.sleep(0.3)
                    l, t, w, h = rect(fl())
                    return {"body": (l + w / 2, t + h / 2), "side": (l + w, t + h / 2)}[where]
                return f

            def sel_rot(k: float):
                def f():
                    fl().Select()
                    time.sleep(0.3)
                    l, t, w, h = rect(fl())
                    return l + w / 2, t - k
                return f

            geo = lambda: [round(fl().Left, 1), round(fl().Top, 1), round(fl().Width, 1), round(fl().Height, 1), round(fl().Rotation, 1)]
            cases = [
                ("table column border", [col_border], (60, 0), lambda: round(tbl().Cell(1, 1).Width, 1)),
                ("table row border", [row_border], (0, 40), lambda: round(tbl().Rows(1).Height, 1)),
                ("picture corner handle (inline)", [sel_inline], (60, 40), lambda: [round(il().Width, 1), round(il().Height, 1)]),
                ("picture move (floating)", [sel_float("body")], (60, 30), geo),
                ("picture side handle (floating)", [sel_float("side")], (60, 0), geo),
                ("picture rotation handle", [sel_rot(k) for k in (28, 34, 40, 22, 46)], (60, 0), geo),
            ]
            for name, points, delta, read in cases:
                out: dict = {"case": name}
                try:
                    point = None
                    for i, pt in enumerate(points):
                        setup()
                        before = read()
                        drag(pt(), delta, False)
                        winapi.press("esc")
                        after = read()
                        if after != before:
                            point = pt
                            out["plain"] = {"before": before, "after": after}
                            if ours():
                                winapi.press("ctrl+z")
                                time.sleep(0.5)
                            out["plain"]["undone"] = read()
                            out["plain"]["undo_restores"] = out["plain"]["undone"] == before
                            break
                    if point is None:
                        out["error"] = "the plain drag changed nothing"
                    else:
                        setup()
                        before = read()
                        drag(point(), delta, True)
                        after = read()
                        out["esc"] = {"before": before, "after": after, "changed": after != before}
                        out["esc_cancels"] = after == before
                        winapi.press("esc")
                except Exception as e:  # noqa: BLE001
                    out["error"] = str(e)[:200]
                res.append(out)
                print(json.dumps(out, ensure_ascii=False), flush=True)
                touch()
            d.Close(0)
        finally:
            try:
                app.Quit(0)
            except Exception:  # noqa: BLE001
                pass
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-word-drag-esc.json"
    out.write_text(json.dumps({"word": res}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
