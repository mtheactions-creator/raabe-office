"""Check (Step 11, part 1, D-156): RaabeWrite's menus close as Word's. Runs the cases of probe_word_menuclose.py in
RaabeWrite with trusted DevTools input (no real input) and compares what changed with Word's recording
(results/<date>-word-menuclose.json): is a popup still open, did the selection move, did the ribbon tab change, did Fett
act on the selected first word.

    .venv/Scripts/python.exe check_menuclose_write.py [word-results.json]
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab import menuclose as M
from feellab.targets import raabe_write

HERE = Path(__file__).parent


class Write:
    def __init__(self):
        self.t = raabe_write("long")
        self.js = self.t.js
        self.ws, self._id = self.t.ws, 0

    # menuclose's helpers send through `ws` and count `_id`: share the target's counter.
    @property
    def _id(self):  # noqa: D401
        return self.t._id

    @_id.setter
    def _id(self, v):
        self.t._id = v

    def point(self, o: int) -> tuple[float, float]:
        a, b = self.js(f"__probe.pointAt({o})"), self.js(f"__probe.pointAt({o + 1})")
        x1 = b["x"] if abs(b["y"] - a["y"]) < 1 else a["x"] + 6
        return (a["x"] + x1) / 2, a["y"] + a.get("h", 14) / 2

    def reset(self) -> None:
        M.close_popups(self)
        self.js("__probe.load('long')")
        time.sleep(0.2)
        M.mouse(self, *M.rect(self, ".tabrow [role=tab]", "Start"))
        self.js("__probe.setSelection(0, 5)")
        time.sleep(0.15)
        n = len(self.js("__probe.text()"))
        # Text points well below and right of where popups open; only points inside the window.
        h, w = self.js("innerHeight"), self.js("innerWidth")
        cands = {}
        for o in range(300, min(n - 2, 6000), 97):
            x, y = self.point(o)
            if 0.45 * h < y < h - 80 and 0.35 * w < x < w - 40 and self.js(f"__probe.text()[{o}]").strip():
                cands[f"@{o}"] = (x, y)
            if len(cands) >= 6:
                break
        self.pts = {"sel": self.point(2), "fontcolor": M.color_arrow(self, "Schriftfarbe"),
                    "bullets": self.js("""(() => { const s = [...document.querySelectorAll('.rsplit')].find((x) => x.querySelector('[aria-label="Aufzählungszeichen"]') && x.offsetWidth > 0);
                      const r = s.querySelector('.arrow').getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; })()""")}
        self.targets = {
            "click_text": cands, "rclick_text": cands,
            "click_ribbontab": {n: M.rect(self, ".tabrow [role=tab]", n) for n in ("Einfügen", "Layout", "Ansicht")},
            "click_bold": {"Fett": M.rect(self, "button.rb", "Fett")},
        }

    def open(self, opener: str) -> bool:
        if opener == "ctx_text":
            M.mouse(self, *self.pts["sel"], "right")
        else:
            M.mouse(self, *self.pts["fontcolor" if opener == "dd_fontcolor" else "bullets"])
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
        M.mouse(self, *pt, "right" if action == "rclick_text" else "left")
        return name

    def state(self) -> dict:
        time.sleep(0.3)
        st = {"popup_open": bool(M.popups(self)), "ribbon_tab": self.js(M.JS_RIBBON_TAB)}
        M.close_popups(self)
        st["selection"] = self.js("__probe.selection()")
        f = self.js("__probe.formatAt(2)") or {}
        st["first_word_bold"] = bool(f.get("bold"))
        return st

    def close(self) -> None:
        self.t.close()


def main() -> None:
    res = sorted((HERE / "results").glob("*-word-menuclose.json"))
    ref = json.loads(Path(sys.argv[1] if len(sys.argv) > 1 else res[-1]).read_text(encoding="utf-8"))
    app = Write()
    out = {"date": datetime.date.today().isoformat(), "reference": ref["date"], "cases": {}}
    ok = 0
    try:
        for key, case in ref["cases"].items():
            if "error" in case:
                continue
            opener, action = key.split("+")
            app.reset()
            if not app.open(opener):
                out["cases"][key] = {"error": "popup did not open"}
                print("FAIL", key, "popup did not open")
                continue
            target = app.act(action)
            if target is None:
                out["cases"][key] = {"error": "every target under the popup"}
                print("SKIP", key)
                continue
            st = app.state()
            sel = st["selection"]
            sel = [sel["anchor"], sel["head"]] if isinstance(sel, dict) else list(sel)
            ours = {"popup_open": st["popup_open"], "selection_moved": sorted(sel) != [0, 5],
                    "ribbon_changed": st["ribbon_tab"] != "Start", "bold": st["first_word_bold"]}
            theirs = {"popup_open": case["popup_open"], "selection_moved": sorted(case["selection"]) != [0, 5],
                      "ribbon_changed": case["ribbon_tab"] != "Start", "bold": case["first_word_bold"]}
            diff = {k: [theirs[k], ours[k]] for k in ours if ours[k] != theirs[k]}
            out["cases"][key] = {"target": target, "ours": st, "diff": diff}
            ok += not diff
            print("ok  " if not diff else "FAIL", key, diff or "")
    finally:
        app.close()
    out["passed"], out["total"] = ok, len(out["cases"])
    path = HERE / "results" / f"{out['date']}-write-menuclose.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(f"{ok}/{len(out['cases'])} as in Word; wrote {path.name}")


if __name__ == "__main__":
    main()
