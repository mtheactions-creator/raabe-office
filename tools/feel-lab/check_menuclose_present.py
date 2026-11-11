"""Check (Step 11, part 1, D-156): RaabePresent's menus close as PowerPoint's. Runs the cases of
probe_powerpoint_menuclose.py in RaabePresent with trusted DevTools input (no real input) and compares what changed with
PowerPoint's recording (results/<date>-powerpoint-menuclose.json): is a popup still open, did the slide change, did the
selection change (its kind and shapes), did the ribbon tab change.

PowerPoint's "before" for an opener is its state after the same popup lost the window's focus (that changes nothing).

    .venv/Scripts/python.exe check_menuclose_present.py [powerpoint-results.json]
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab import menuclose as M
from feellab.present_catalogue import PresentSide

HERE = Path(__file__).parent
SHAPES = [{"kind": "rect", "box": [100, 100, 200, 120], "text": "A"}, {"kind": "rect", "box": [600, 300, 200, 120], "text": "B"}]
COM_TYPES = {0: "None", 1: "Slides", 2: "Shapes", 3: "Text"}


class Present:
    def __init__(self):
        self.side = PresentSide()
        self.js = self.side.js

    @property
    def ws(self):
        return self.side.ws

    @property
    def _id(self):
        return self.side._id

    @_id.setter
    def _id(self, v):
        self.side._id = v

    def sp(self, x: float, y: float) -> tuple[float, float]:
        p = self.js(f"__present.slidePointToClient({x}, {y})")
        return p["x"], p["y"]

    def thumb(self, n: int) -> tuple[float, float]:
        r = self.js(f"__present.thumbRect({n})")
        return r["x"] + r["w"] / 2, r["y"] + r["h"] / 2

    def reset(self) -> None:
        M.close_popups(self)
        self.js(f"__present.setupCase({json.dumps({'layout': 'Leer', 'slides': 4, 'shapes': SHAPES})})")
        time.sleep(0.3)
        M.mouse(self, *M.rect(self, ".tabrow [role=tab]", "Start"))
        # Shape 1 selected by a click, as in PowerPoint's probe.
        M.mouse(self, *self.sp(450, 480))
        M.mouse(self, *self.sp(110, 110))
        time.sleep(0.15)
        self.ids = [s["id"] for s in self.js("__present.state()")["shapes"] if not s.get("ph")]
        self.pts = {"shape1": self.sp(200, 160), "thumb1": self.thumb(1), "slide": self.sp(450, 470),
                    "fontcolor": M.color_arrow(self, "Schriftfarbe"), "layout": M.rect(self, "button.rb", "Layout")}
        empty = {"empty_right_top": self.sp(880, 50), "empty_bottom": self.sp(450, 500), "empty_right": self.sp(900, 250)}
        self.targets = {
            "click_empty": empty, "rclick_empty": empty,
            "click_shape2": {"shape2": self.sp(700, 360), "shape2_corner": self.sp(780, 405)},
            "click_thumb3": {"thumb3": self.thumb(3), "thumb2": self.thumb(2)},
            "click_ribbontab": {n: M.rect(self, ".tabrow [role=tab]", n) for n in ("Einfügen", "Entwurf", "Ansicht")},
            "click_bold": {"Fett": M.rect(self, "button.rb", "Fett")},
        }

    def open(self, opener: str) -> bool:
        p = self.pts
        if opener == "ctx_shape":
            M.mouse(self, *p["shape1"], "right")
        elif opener == "ctx_thumb":
            M.mouse(self, *p["thumb1"], "right")
        elif opener == "ctx_slide":
            M.mouse(self, *p["slide"], "right")
        else:
            M.mouse(self, *p["fontcolor" if opener == "dd_fontcolor" else "layout"])
        return M.wait_popup(self)

    def act(self, action: str) -> str | None:
        if action == "esc":
            M.key(self)
            return "esc"
        if action == "click_other_app":
            self.js("window.dispatchEvent(new Event('blur'))")
            return "other"
        name, pt = M.pick_outside(self.targets[action], M.popups(self))
        if name is None:
            return None
        M.mouse(self, *pt, "right" if action == "rclick_empty" else "left")
        return name

    def state(self) -> dict:
        time.sleep(0.3)
        st = {"popup_open": bool(M.popups(self)), "ribbon_tab": self.js(M.JS_RIBBON_TAB)}
        M.close_popups(self)
        s = self.js("__present.state()")
        sel = self.js("__present.comSelection()")
        st["slide"] = s["slide"]
        st["selection_type"] = sel["type"]
        st["shapes"] = ["AB"[self.ids.index(i)] if i in self.ids else "?" for i in sel.get("ids", [])]
        return st

    def close(self) -> None:
        self.side.close()


def kind(t: str) -> str:
    # PowerPoint reports a shape with text under a right click as Text; both mean "this shape is selected".
    return "Shapes" if t in ("Shapes", "Text") else t


def main() -> None:
    res = sorted((HERE / "results").glob("*-powerpoint-menuclose.json"))
    ref = json.loads(Path(sys.argv[1] if len(sys.argv) > 1 else res[-1]).read_text(encoding="utf-8"))
    app = Present()
    out = {"date": datetime.date.today().isoformat(), "reference": ref["date"], "cases": {}}
    ok = 0
    try:
        for key, case in ref["cases"].items():
            if "error" in case:
                continue
            opener, action = key.split("+")
            before_ref = ref["cases"].get(f"{opener}+click_other_app", {})
            app.reset()
            if not app.open(opener):
                out["cases"][key] = {"error": "popup did not open"}
                print("FAIL", key, "popup did not open")
                continue
            b = app.js("__present.comSelection()")
            before = {"slide": app.js("__present.state()")["slide"], "type": kind(b["type"]),
                      "shapes": ["AB"[app.ids.index(i)] if i in app.ids else "?" for i in b.get("ids", [])]}
            target = app.act(action)
            if target is None:
                out["cases"][key] = {"error": "every target under the popup"}
                print("SKIP", key)
                continue
            st = app.state()
            ours = {"popup_open": st["popup_open"], "slide_changed": st["slide"] != before["slide"],
                    "selection_changed": (kind(st["selection_type"]), st["shapes"]) != (before["type"], before["shapes"]),
                    "ribbon_changed": st["ribbon_tab"] != "Start"}
            ref_sel = lambda c: (kind(COM_TYPES.get(c["selection_type"], "?")), c["shapes"])  # noqa: E731
            theirs = {"popup_open": case["popup_open"], "slide_changed": case["slide"] != before_ref.get("slide", 1),
                      "selection_changed": ref_sel(case) != ref_sel(before_ref) if before_ref else False,
                      "ribbon_changed": case["ribbon_tab"] != "Start"}
            diff = {k: [theirs[k], ours[k]] for k in ours if ours[k] != theirs[k]}
            out["cases"][key] = {"target": target, "ours": st, "before": before, "diff": diff}
            ok += not diff
            print("ok  " if not diff else "FAIL", key, diff or "")
    finally:
        app.close()
    out["passed"], out["total"] = ok, len(out["cases"])
    path = HERE / "results" / f"{out['date']}-present-menuclose.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(f"{ok}/{len(out['cases'])} as in PowerPoint; wrote {path.name}")


if __name__ == "__main__":
    main()
