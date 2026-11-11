"""Check (Step 11, part 1, D-156): RaabeCalc's menus close as Excel's. Runs the cases of probe_excel_menuclose.py in
RaabeCalc with trusted DevTools mouse input (Input.dispatchMouseEvent; no real input, the screen stays free) and
compares what changed with Excel's recording (results/<date>-excel-menuclose.json): is a popup still open, did the
selection move, did the sheet, the ribbon tab or the zoom change, did Fett act, is a cell being edited.

    .venv/Scripts/python.exe check_menuclose_calc.py [excel-results.json]
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab.calc_catalogue import CalcSide

HERE = Path(__file__).parent

# The DOM side: is a popup open (anything floating above the app: menus, palettes, galleries, panels).
JS_POPUP = """(() => {
  const vis = (e) => e && e.isConnected && e.offsetWidth > 40 && e.offsetHeight > 30 && getComputedStyle(e).display !== 'none' && getComputedStyle(e).visibility !== 'hidden';
  const fixed = [...document.body.children].filter((e) => !e.matches('.tooltip, script, style') && getComputedStyle(e).position === 'fixed' && vis(e));
  const own = [...document.querySelectorAll('.quick-panel, .filter-panel, .pick-list, .dv-list, .chart-flyout')].filter(vis);
  const shell = document.documentElement.hasAttribute('data-popup-open') ? [...document.body.children].filter((e) => e.matches('.menu-root, .tile-panel, .menu, .palette, .rpanel, .group-popup') || (getComputedStyle(e).position === 'fixed' && vis(e))) : [];
  return [...new Set([...fixed, ...own, ...shell])].filter(vis).map((e) => { const r = e.getBoundingClientRect(); return [r.left, r.top, r.right, r.bottom]; });
})()"""

JS_RECT = """((sel, text) => {
  const els = [...document.querySelectorAll(sel)].filter((e) => !text || e.textContent.trim() === text || e.getAttribute('aria-label') === text);
  const e = els.find((x) => x.offsetWidth > 0);
  if (!e) return null;
  const r = e.getBoundingClientRect();
  return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
})"""


class Calc:
    def __init__(self):
        self.side = CalcSide()
        self.js = self.side.js

    def mouse(self, x: float, y: float, button: str = "left") -> None:
        cdp = self.side
        for t in ("mouseMoved", "mousePressed", "mouseReleased"):
            params = {"type": t, "x": x, "y": y, "button": "none" if t == "mouseMoved" else button,
                      "buttons": 0 if t != "mousePressed" else (1 if button == "left" else 2),
                      "clickCount": 0 if t == "mouseMoved" else 1}
            cdp._id += 1
            cdp.ws.send(json.dumps({"id": cdp._id, "method": "Input.dispatchMouseEvent", "params": params}))
            while json.loads(cdp.ws.recv()).get("id") != cdp._id:
                pass
            time.sleep(0.03)

    def key(self, name: str, code: int) -> None:
        cdp = self.side
        for t in ("rawKeyDown", "keyUp"):
            cdp._id += 1
            cdp.ws.send(json.dumps({"id": cdp._id, "method": "Input.dispatchKeyEvent",
                                    "params": {"type": t, "key": name, "code": name, "windowsVirtualKeyCode": code}}))
            while json.loads(cdp.ws.recv()).get("id") != cdp._id:
                pass

    def rect(self, sel: str, text: str = "") -> tuple[float, float]:
        p = self.js(f"{JS_RECT}({json.dumps(sel)}, {json.dumps(text)})")
        if not p:
            raise RuntimeError(f"not found: {sel} {text}")
        return p["x"], p["y"]

    def popups(self) -> list:
        return self.js(JS_POPUP)

    def reset(self) -> None:
        for _ in range(3):
            if self.popups():
                self.key("Escape", 27)
                time.sleep(0.15)
        self.js('__calc.reset({"B2": "x", "D5": "y"}, "B2", 4, [])')
        tab = self.rect(".tabrow [role=tab]", "Start")
        self.mouse(*tab)
        time.sleep(0.2)
        P = lambda a: (lambda p: (p["x"], p["y"]))(self.js(f"__calc.point({json.dumps(a)}, 'center')"))  # noqa: E731
        H = lambda k, w: (lambda p: (p["x"], p["y"]))(self.js(f"__calc.header({json.dumps(k)}, {json.dumps(str(w))})"))  # noqa: E731
        tabs = lambda n: self.rect(".sheettab", n)  # noqa: E731
        fill = self.js("""(() => { const s = [...document.querySelectorAll('.rsplit')].find((x) => x.querySelector('.color-main')?.getAttribute('aria-label') === 'Füllfarbe');
          const r = s.querySelector('.arrow').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()""")
        self.pts = {"B2": P("B2"), "colC": H("col", "C"), "row4": H("row", 4), "tab1": tabs("Tabelle1"),
                    "fill": (fill["x"], fill["y"]), "styles": self.rect("button.rb", "Zellenformatvorlagen")}
        self.targets = {
            "click_cell": {"D5": P("D5"), "J5": P("J5"), "J20": P("J20"), "D30": P("D30")},
            "rclick_cell": {"F7": P("F7"), "L8": P("L8"), "L22": P("L22")},
            "click_colhead": {"E": H("col", "E"), "M": H("col", "M"), "P": H("col", "P")},
            "click_sheettab": {"Tabelle3": tabs("Tabelle3"), "Tabelle5": tabs("Tabelle5")},
            "click_ribbontab": {n: self.rect(".tabrow [role=tab]", n) for n in ("Einfügen", "Formeln", "Ansicht")},
            "click_bold": {"Fett": self.rect("button.rb", "Fett")},
            "click_fxbar": {"fx": self.rect(".fx-input")},
            "click_zoomin": {"Vergrößern": self.rect(".statusbar button", "Vergrößern")},
        }

    def open(self, opener: str) -> bool:
        p = self.pts
        if opener.startswith("ctx_"):
            self.mouse(*p[{"ctx_cell": "B2", "ctx_colhead": "colC", "ctx_rowhead": "row4", "ctx_sheettab": "tab1"}[opener]], "right")
        else:
            self.mouse(*p["fill" if opener == "dd_fill" else "styles"])
        for _ in range(10):
            time.sleep(0.1)
            if self.popups():
                return True
        return False

    def act(self, action: str) -> str | None:
        if action == "esc":
            self.key("Escape", 27)
            return "esc"
        if action == "click_other_app":
            # Without real input: the window's blur, as when another program takes the focus.
            self.js("window.dispatchEvent(new Event('blur'))")
            return "other"
        pops = self.popups()
        for name, (x, y) in self.targets[action].items():
            if all(not (r[0] - 8 <= x <= r[2] + 8 and r[1] - 8 <= y <= r[3] + 8) for r in pops):
                self.mouse(x, y, "right" if action == "rclick_cell" else "left")
                return name
        return None

    def state(self) -> dict:
        time.sleep(0.3)
        st = self.js('__calc.state(["B2"])')
        out = {"popup_open": bool(self.popups()), "selection": st.get("selection"), "sheet": st.get("sheet"),
               "mode": st.get("mode"), "zoom": self.js("__calc.app.sheet.view.zoom"),
               "ribbon_tab": self.js("document.querySelector('.tabrow [role=tab][aria-selected=true]')?.textContent?.trim()"),
               "b2_bold": bool((st.get("cells", {}).get("B2") or {}).get("bold"))}
        for _ in range(3):
            if self.popups():
                self.key("Escape", 27)
                time.sleep(0.15)
        return out

    def close(self) -> None:
        self.side.close()


def deltas(st: dict, before: dict) -> dict:
    return {
        "popup_open": st["popup_open"],
        "selection_moved": st["selection"] != before["selection"],
        "sheet_changed": st["sheet"] != "Tabelle1",
        "ribbon_changed": st["ribbon_tab"] != "Start",
        "bold": bool(st["b2_bold"]),
        "zoom_changed": abs((st["zoom"] or 0) - before["zoom"]) > 1e-6,
        "editing": st["mode"] not in ("Bereit", "?", None),
    }


def excel_deltas(case: dict, opener: str) -> dict:
    before_sel = {"ctx_colhead": "C:C", "ctx_rowhead": "4:4"}.get(opener, "B2")
    return {
        "popup_open": case["popup_open"],
        "selection_moved": case["selection"] != before_sel,
        "sheet_changed": case["sheet"] != "Tabelle1",
        "ribbon_changed": case["ribbon_tab"] != "Start",
        "bold": case["b2_bold"],
        "zoom_changed": case.get("zoom", 100) != 100,
        "editing": case["mode"] not in ("Bereit", "?"),
    }


def main() -> None:
    res = sorted((HERE / "results").glob("*-excel-menuclose.json"))
    ref = json.loads(Path(sys.argv[1] if len(sys.argv) > 1 else res[-1]).read_text(encoding="utf-8"))
    app = Calc()
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
            before = app.js('__calc.state([])')
            before = {"selection": before.get("selection"), "zoom": app.js("__calc.app.sheet.view.zoom")}
            target = app.act(action)
            if target is None:
                out["cases"][key] = {"error": "every target under the popup"}
                print("SKIP", key)
                continue
            st = app.state()
            ours, theirs = deltas(st, before), excel_deltas(case, opener)
            diff = {k: [theirs[k], ours[k]] for k in ours if ours[k] != theirs[k]}
            out["cases"][key] = {"target": target, "ours": st, "diff": diff}
            ok += not diff
            print("ok  " if not diff else "FAIL", key, diff or "")
    finally:
        app.close()
    out["passed"] = ok
    out["total"] = len(out["cases"])
    path = HERE / "results" / f"{out['date']}-calc-menuclose.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(f"{ok}/{len(out['cases'])} as in Excel; wrote {path.name}")


if __name__ == "__main__":
    main()
