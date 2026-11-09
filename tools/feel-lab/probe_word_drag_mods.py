"""What Word does with modifier keys pressed mid-drag and with the pointer leaving the window (D-127, the drag test's
reference for RaabeWrite): our own visible Word, real drags on a table's column and row border, an inline picture's
corner handle, a floating picture (body, side handle, rotation handle) and a text selection, each along
feellab/officedrag.py's variants (plain; Umschalt, Strg, Alt let go mid-drag or held over the release; out of the
window and back; the button let go outside). Read back through COM: the anchor's value, the number of pictures
(copies) and the scroll position. The ruler is not probed (its markers' places are not in the object model).
Takes over mouse and keyboard (the real-input lock). Result: results/2026-11-09-word-drag-mods.json.
"""
from __future__ import annotations

import json
import sys
import tempfile
import time
from pathlib import Path

from PIL import Image

from feellab import winapi
from feellab.inputlock import real_input, touch
from feellab import officedrag
from feellab.catcher import catcher
from feellab.officedrag import VARIANTS, drag

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir())
PNG = TMP / "raabe-drag-mods.png"
TEXT = "Ein Absatz mit etwas Text, damit das Bild umflossen werden kann. " * 4
RECT = (0, 0, 2000, 1700)


def main() -> None:
    import win32com.client
    sys.stdout.reconfigure(encoding="utf-8")
    want = sys.argv[1:]
    Image.new("RGB", (240, 160), (70, 120, 200)).save(PNG)
    res: list[dict] = []
    with real_input("probe_word_drag_mods.py (drag-paths agent)"):
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = True
        guard = catcher(RECT)
        guard.__enter__()
        try:
            d = app.Documents.Add()
            win = d.Windows(1)
            win.View.Type = 3
            win.View.Zoom.Percentage = 100
            winapi.place(win.Hwnd, *RECT)
            pid = winapi.window_pid(win.Hwnd)

            def ours() -> bool:
                return winapi.window_pid(winapi.foreground()) == pid

            def front() -> None:
                officedrag.front(win.Hwnd, ours)

            def setup() -> None:
                d.Content.Delete()
                d.Content.Text = "\r".join([TEXT, "", TEXT, TEXT, TEXT, TEXT, ""])
                t = d.Tables.Add(d.Paragraphs(2).Range, 2, 2, 1, 0)
                for i in (1, 2):
                    for j in (1, 2):
                        t.Cell(i, j).Range.Text = f"Z{i}S{j}"
                d.InlineShapes.AddPicture(str(PNG), False, True, d.Paragraphs(d.Paragraphs.Count).Range)
                s = d.Shapes.AddPicture(str(PNG), False, True, 300, 20, 120, 80, d.Paragraphs(1).Range)
                s.Name = "Probe"
                s.WrapFormat.Type = 0
                d.UndoClear()
                win.ActivePane.VerticalPercentScrolled = 0
                d.Range(0, 0).Select()
                time.sleep(0.4)

            def rect(obj) -> tuple[float, float, float, float]:
                left, top, w, h = win.GetPoint(0, 0, 0, 0, obj)
                return float(left), float(top), float(w), float(h)

            tbl = lambda: d.Tables(1)
            fl = lambda: d.Shapes("Probe")
            il = lambda: d.InlineShapes(1)

            def col_border():
                l, t, w, h = rect(tbl().Cell(1, 2).Range)
                return l - 5.4 * 2, t + h / 2

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

            def text_point():
                r = d.Paragraphs(4).Range.Characters(30)
                l, t, w, h = rect(r)
                return l + 1, t + h / 2

            geo = lambda: [round(fl().Left, 1), round(fl().Top, 1), round(fl().Width, 1), round(fl().Height, 1), round(fl().Rotation, 1)]
            cases = [
                ("table column border", [col_border], "right", lambda: round(tbl().Cell(1, 1).Width, 1)),
                ("table row border", [row_border], "down", lambda: round(tbl().Rows(1).Height, 1)),
                ("picture corner handle (inline)", [sel_inline], "right", lambda: [round(il().Width, 1), round(il().Height, 1)]),
                ("picture move (floating)", [sel_float("body")], "right", geo),
                ("picture side handle (floating)", [sel_float("side")], "right", geo),
                ("picture rotation handle", [sel_rot(k) for k in (28, 34, 40, 22, 46)], "right", geo),
                ("select text", [text_point], "down", lambda: [app.Selection.Start, app.Selection.End]),
            ]
            for name, points, direction, read in cases:
                if want and not any(w in name for w in want):
                    continue
                out: dict = {"case": name, "out": direction}
                try:
                    point = None
                    for i, pt in enumerate(points):
                        setup()
                        front()
                        before = read()
                        front()
                        officedrag.run_variant(lambda pt=pt: (setup(), front(), drag(pt(), "plain", ours), {})[-1], ours)
                        after = read()
                        winapi.press("esc")
                        if after != before:
                            point = pt
                            out["point"] = i
                            break
                    if point is None:
                        out["error"] = "the plain drag changed nothing"
                    else:
                        for v in VARIANTS:
                            def attempt(v=v) -> dict:
                                setup()
                                front()
                                p = point()
                                front()
                                before = read()
                                n0 = d.Shapes.Count + d.InlineShapes.Count
                                s0 = win.ActivePane.VerticalPercentScrolled
                                drag(p, v, ours, window=winapi.window_rect(win.Hwnd), direction=direction)
                                after = read()
                                n1 = d.Shapes.Count + d.InlineShapes.Count
                                s1 = win.ActivePane.VerticalPercentScrolled
                                return {"before": before, "after": after, "copies": n1 - n0, "scrolled": s1 - s0, "table_width": round(sum(tbl().Cell(1, j).Width for j in (1, 2)), 1)}
                            out[v] = officedrag.run_variant(attempt, ours)
                            officedrag.settle(ours)
                            time.sleep(0.2)
                            touch()
                except Exception as e:  # noqa: BLE001
                    out["error"] = str(e)[:200]
                res.append(out)
                print(json.dumps(out, ensure_ascii=False), flush=True)
                touch()
            d.Close(0)
        finally:
            guard.__exit__(None, None, None)
            try:
                app.Quit(0)
            except Exception:  # noqa: BLE001
                pass
    path = HERE / "results" / "2026-11-09-word-drag-mods.json"
    if want and path.exists():
        old = json.loads(path.read_text(encoding="utf-8")).get("word", [])
        names = {r["case"] for r in res}
        res = [r for r in old if r["case"] not in names] + res
    path.write_text(json.dumps({"word": res}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
