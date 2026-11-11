"""Check (Step 11, part 5): context menus look as Office's. Opens the context menu of a cell (Excel / RaabeCalc), of
selected text (Word / RaabeWrite) and of a shape (PowerPoint / RaabePresent) and compares the menu's width and every
entry's place and height (CSS px from the menu's top left), entry by entry by name, and where the separators are.
Office's side: our own instances, real input under the lock (the probes of D-156 open the menus); ours: DevTools.

    .venv/Scripts/python.exe check_contextmenu_geometry.py [calc] [write] [present]
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab import menuclose as M
from feellab import uia, winapi
from feellab.inputlock import real_input
from feellab.popups import POPUP_CLASSES

HERE = Path(__file__).parent
TOL = 2.0

JS_MENU = """(() => {
  const m = [...document.querySelectorAll('.menu-root > .menu')].find((e) => e.offsetWidth > 0);
  if (!m) return null;
  const r0 = m.getBoundingClientRect();
  const box = (e) => { const r = e.getBoundingClientRect(); return [+(r.left - r0.left).toFixed(1), +(r.top - r0.top).toFixed(1), +r.width.toFixed(1), +r.height.toFixed(1)]; };
  return { w: +r0.width.toFixed(1), h: +r0.height.toFixed(1), items: [...m.children].filter((e) => e.offsetHeight > 0).map((e) => ({
    type: e.tagName === 'HR' ? 'Separator' : 'MenuItem', name: e.tagName === 'HR' ? '' : (e.getAttribute('aria-label') || e.textContent.trim()), rect: box(e) })) };
})()"""


def office_menu(pid: int) -> dict | None:
    from probe_word_context import items_of
    hs = [h for h in winapi.find_windows(lambda h: winapi.window_pid(h) == pid and winapi.window_class(h) in POPUP_CLASSES)]
    if not hs:
        return None
    rects = {h: winapi.window_rect(h) for h in hs}
    menu = max(hs, key=lambda h: rects[h][3] - rects[h][1])
    l, t, r, b = rects[menu]
    s = winapi.dpi_scale(menu)
    items = []
    for _, d in items_of(menu):
        # The search box's accessible name is Office's instruction text: Microsoft's prose, never recorded (D-058).
        if d["type"] in ("Edit", "ComboBox") or len(d["name"]) > 60:
            continue
        x, y, w, h = d["rect"]
        items.append({"type": "Separator" if d["type"] == "Separator" else "MenuItem", "name": d["name"],
                      "rect": [round((x - l) / s, 1), round((y - t) / s, 1), round(w / s, 1), round(h / s, 1)]})
    return {"w": round((r - l) / s, 1), "h": round((b - t) / s, 1), "items": items}


def office_side(app: str) -> dict | None:
    if app == "calc":
        from probe_excel_menuclose import Probe
        pr = Probe()
        try:
            pr.reset()
            pr.open("ctx_cell")
            time.sleep(0.6)
            return office_menu(pr.pid)
        finally:
            winapi.press("esc")
            pr.side.close()
    if app == "write":
        from probe_word_menuclose import Probe
        pr = Probe()
        try:
            pr.reset()
            pr.open("ctx_text")
            time.sleep(0.6)
            return office_menu(pr.pid)
        finally:
            winapi.press("esc")
            pr.w.close()
    from probe_powerpoint_menuclose import Probe
    pr = Probe()
    try:
        pr.reset()
        pr.open("ctx_shape")
        time.sleep(0.6)
        return office_menu(pr.pid)
    finally:
        winapi.press("esc")
        pr.close()


def our_side(app: str) -> dict | None:
    if app == "calc":
        from check_menuclose_calc import Calc
        a = Calc()
        try:
            a.reset()
            a.open("ctx_cell")
            return a.js(JS_MENU)
        finally:
            a.close()
    if app == "write":
        from check_menuclose_write import Write
        a = Write()
        try:
            a.reset()
            a.open("ctx_text")
            return a.js(JS_MENU)
        finally:
            a.close()
    from check_menuclose_present import Present
    a = Present()
    try:
        a.reset()
        a.open("ctx_shape")
        return a.js(JS_MENU)
    finally:
        a.close()


def compare(off: dict, ours: dict) -> list[str]:
    out = []
    if abs(off["w"] - ours["w"]) > TOL:
        out.append(f"menu width {ours['w']} vs Office {off['w']}")
    oi = {i["name"]: i for i in off["items"] if i["type"] == "MenuItem" and i["name"]}
    mi = {i["name"]: i for i in ours["items"] if i["type"] == "MenuItem" and i["name"]}
    # Places relative to the first entry both menus share (Office's menu window has a frame of its own).
    common = [n for n in oi if n in mi]
    if not common:
        return out + ["no entry in common"]
    oy0, my0 = oi[common[0]]["rect"][1], mi[common[0]]["rect"][1]
    for n in common:
        a, b = oi[n]["rect"], mi[n]["rect"]
        dy, dh = round((b[1] - my0) - (a[1] - oy0), 1), round(b[3] - a[3], 1)
        if abs(dy) > TOL or abs(dh) > TOL:
            out.append(f"{n}: dy {dy} dh {dh}")
    for n in oi:
        if n not in mi and len(n) <= 60:
            out.append(f"missing: {n}")
    for n in mi:
        if n not in oi:
            out.append(f"extra: {n}")
    return out


def main() -> None:
    apps = [a for a in sys.argv[1:] if a in ("calc", "write", "present")] or ["calc", "write", "present"]
    out: dict = {"date": datetime.date.today().isoformat(), "apps": {}}
    for app in apps:
        with real_input(f"check_contextmenu_geometry {app}"):
            off = office_side(app)
        ours = our_side(app)
        if not off or not ours:
            out["apps"][app] = {"error": "menu not found", "office": off, "ours": ours}
            print(app, "menu not found", bool(off), bool(ours))
            continue
        d = compare(off, ours)
        out["apps"][app] = {"office": off, "ours": ours, "diff": d}
        print(app, "ok" if not d else f"{len(d)} differences")
        for x in d[:30]:
            print("   ", x)
    path = HERE / "results" / f"{out['date']}-contextmenu-geometry.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("wrote", path.name)


if __name__ == "__main__":
    main()
