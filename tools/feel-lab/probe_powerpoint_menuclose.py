"""Probe (Step 11, part 1): how PowerPoint's menus close. Slide 1 holds two rectangles, the first one selected. Opens a
popup (context menu of the shape, of the first thumbnail, of the empty slide; the Schriftfarbe palette; the Layout
gallery), then does one thing (click on the empty slide, on the other shape, on thumbnail 3, right click on the empty
slide, a ribbon tab, Fett, Esc, a click into another program's window) and records: is a popup still open, which slide
is shown, what is selected, did Fett act. Own PowerPoint (only when the owner's is not running), real input under the
lock.

    .venv/Scripts/python.exe probe_powerpoint_menuclose.py
"""
from __future__ import annotations

import datetime
import json
import os
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.inputlock import real_input
from feellab.popups import ForeignWindow, find_named, own_popups, pick, rect_center, selected_ribbon_tab, split_arrow, wait_popup
from feellab.powerpoint import PowerPointTarget

HERE = Path(__file__).parent

OPENERS = ["ctx_shape", "ctx_thumb", "ctx_slide", "dd_fontcolor", "gallery_layout"]
ACTIONS = ["click_empty", "click_shape2", "click_thumb3", "rclick_empty", "click_ribbontab", "click_bold", "esc",
           "click_other_app"]


class Probe:
    def __init__(self):
        self.p = PowerPointTarget(visible=True)
        self.pres = self.p.new_presentation(with_window=True)
        for _ in range(3):
            self.p.add_slide(self.pres, 7)
        s1 = self.pres.Slides(1)
        a = s1.Shapes.AddShape(1, 100, 100, 200, 120)
        a.TextFrame.TextRange.Text = "A"
        s1.Shapes.AddShape(1, 600, 300, 200, 120).TextFrame.TextRange.Text = "B"
        self.pid = self.p.pid
        self.hwnd = self.p.hwnd
        self.root = uia.element_for(self.hwnd)
        time.sleep(1.5)
        # Slide points to screen pixels, once (COM refuses it while the thumbnail pane has the focus).
        w = self.p.app.ActiveWindow
        x0, y0 = w.PointsToScreenPixelsX(0), w.PointsToScreenPixelsY(0)
        x1, y1 = w.PointsToScreenPixelsX(960), w.PointsToScreenPixelsY(540)
        self.map = (x0, y0, (x1 - x0) / 960, (y1 - y0) / 540)

    def ours(self) -> bool:
        return winapi.window_pid(winapi.foreground()) == self.pid

    def px(self, x: float, y: float) -> tuple[float, float]:
        x0, y0, sx, sy = self.map
        return x0 + x * sx, y0 + y * sy

    def thumb(self, i: int) -> tuple[float, float]:
        # Thumbnails are not in the automation tree: their places in our 2000 px window (recorded from a screenshot).
        l, t, _, _ = winapi.window_rect(self.hwnd)
        return l + 170, t + {1: 362, 2: 522, 3: 682}[i]

    def reset(self) -> None:
        for _ in range(3):
            if own_popups(self.pid):
                winapi.activate(self.hwnd)
                winapi.press("esc")
                time.sleep(0.3)
        winapi.activate(self.hwnd)
        time.sleep(0.3)
        s1 = self.pres.Slides(1)
        s1.Shapes(1).TextFrame.TextRange.Font.Bold = False
        self.p.app.ActiveWindow.View.GotoSlide(1)
        # One real click on the slide gives the window the focus (after the other program had it), then shape 1.
        winapi.click(*self.px(450, 480))
        time.sleep(0.3)
        s1.Shapes(1).Select()
        uia.select(find_named(self.root, "TabItem", "Start"))
        time.sleep(0.4)
        self.pts = {
            "shape1": self.px(200, 160), "thumb1": self.thumb(1), "slide": self.px(450, 470),
            "fontcolor": rect_center(split_arrow(self.root, "Schriftfarbe")),
            "layout": rect_center(find_named(self.root, "MenuItem", "Layout")),
        }
        self.targets = {
            "click_empty": {"empty_right_top": self.px(880, 50), "empty_bottom": self.px(450, 500), "empty_right": self.px(900, 250)},
            "rclick_empty": {"empty_right_top": self.px(880, 50), "empty_bottom": self.px(450, 500), "empty_right": self.px(900, 250)},
            "click_shape2": {"shape2": self.px(700, 360), "shape2_corner": self.px(780, 405)},
            "click_thumb3": {"thumb3": self.thumb(3), "thumb2": self.thumb(2)},
            "click_ribbontab": {n: rect_center(find_named(self.root, "TabItem", n)) for n in ("Einfügen", "Entwurf", "Ansicht")},
            "click_bold": {"Fett": rect_center(find_named(self.root, "Button", "Fett"))},
        }

    def open(self, opener: str) -> bool:
        for _ in range(2):
            if not self.ours():
                winapi.activate(self.hwnd)
            if opener == "ctx_shape":
                winapi.right_click(*self.pts["shape1"])
            elif opener == "ctx_thumb":
                winapi.right_click(*self.pts["thumb1"])
            elif opener == "ctx_slide":
                winapi.right_click(*self.pts["slide"])
            elif opener == "dd_fontcolor":
                winapi.click(*self.pts["fontcolor"])
            else:
                winapi.click(*self.pts["layout"])
            if wait_popup(self.pid):
                return True
        return False

    def act(self, action: str, other: ForeignWindow) -> str | None:
        if not self.ours():
            raise RuntimeError("our PowerPoint is not in front")
        if action == "esc":
            winapi.press("esc")
            return "esc"
        if action == "click_other_app":
            winapi.click(*other.point)
            return "other"
        name, pt = pick(self.targets[action], own_popups(self.pid))
        if name is None:
            return None
        (winapi.right_click if action == "rclick_empty" else winapi.click)(*pt)
        return name

    def state(self) -> dict:
        time.sleep(0.7)
        pops = own_popups(self.pid)
        st = {"popup_open": bool(pops), "popups": pops, "foreground_ours": self.ours(),
              "ribbon_tab": selected_ribbon_tab(self.root)}
        for _ in range(3):
            if own_popups(self.pid):
                winapi.activate(self.hwnd)
                winapi.press("esc")
                time.sleep(0.3)
        w = self.p.app.ActiveWindow
        st["slide"] = w.View.Slide.SlideIndex
        sel = w.Selection
        st["selection_type"] = sel.Type  # 0 none, 1 slides, 2 shapes, 3 text
        st["shapes"] = [sel.ShapeRange(i).TextFrame.TextRange.Text for i in range(1, sel.ShapeRange.Count + 1)] if sel.Type in (2, 3) else []
        st["shape1_bold"] = bool(self.pres.Slides(1).Shapes(1).TextFrame.TextRange.Font.Bold)
        return st

    def close(self) -> None:
        for pr in self.p.own:
            try:
                pr.Close()
            except Exception:  # noqa: BLE001
                pass
        if self.p.app.Presentations.Count == 0:
            self.p.app.Quit()


def main() -> None:
    out: dict = {"date": datetime.date.today().isoformat(), "app": "PowerPoint", "cases": {}}
    only = set(filter(None, os.environ.get("ONLY", "").split(",")))
    with real_input("probe_powerpoint_menuclose"):
        pr = Probe()
        try:
            other = ForeignWindow()
        except Exception:
            pr.close()
            raise
        try:
            for opener in OPENERS:
                for action in ACTIONS:
                    key = f"{opener}+{action}"
                    if only and key not in only:
                        continue
                    pr.reset()
                    if not pr.open(opener):
                        out["cases"][key] = {"error": "popup did not open"}
                        print(key, "NOT OPENED")
                        continue
                    target = pr.act(action, other)
                    if target is None:
                        out["cases"][key] = {"error": "every target under the popup"}
                        print(key, "NO TARGET")
                        continue
                    st = pr.state()
                    st["target"] = target
                    out["cases"][key] = st
                    print(key, {k: st[k] for k in ("target", "popup_open", "slide", "selection_type", "shapes", "ribbon_tab", "shape1_bold")})
                    winapi.move(1900, 1950)
        finally:
            other.close()
            pr.close()
    path = HERE / "results" / f"{out['date']}-powerpoint-menuclose.json"
    if only and path.exists():
        old = json.loads(path.read_text(encoding="utf-8"))
        old["cases"].update(out["cases"])
        out = old
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("wrote", path.name)


if __name__ == "__main__":
    main()
