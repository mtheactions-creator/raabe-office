"""RaabeCalc against Excel's recordings for D-120 and D-122, "by hand" in the real exe without real input (off screen,
trusted DevTools input):
- keytips: Ansicht > Makros as recorded in Excel (Alt, F, M opens the menu; K, R, V; A while recording);
- Namenfeld / Bearbeitungsleiste (results/2026-11-09-excel-formulabar.json): the handle's pointer, dragging it (the
  Namenfeld's width follows, within Excel's smallest and largest), the expand chevron and Strg+Umschalt+U (Excel's
  FormulaBarHeight after each), dragging the lower edge;
- Datenschnitt with several items (results/2026-11-09-excel-slicer-multi.json): every recorded case replayed with
  clicks, Strg/Umschalt+clicks, drags, Alt+S and the Mehrfachauswahl button; the items chosen after each step compared.

    .venv/Scripts/python.exe check_calc_misc.py

Result: results/2026-11-09-calc-misc.json.
"""
from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path

from feellab import calc_catalogue, winapi
from feellab.calc_catalogue import CalcSide
from feellab.cdpinput import MODS, Cdp

HERE = Path(__file__).parent
RESULT = HERE / "results" / "2026-11-09-calc-misc.json"
os.environ["RAABE_FEEL_BACKGROUND"] = "1"
calc_catalogue.winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4000, 0, w, h)


def load(name: str) -> dict:
    p = HERE / "results" / name
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}


def keytips(c: CalcSide, inp: Cdp) -> list[dict]:
    out = []

    def check(name: str, ok: bool, detail=None) -> None:
        out.append({"check": name, "ok": bool(ok), "detail": detail})

    c.js("__calc.reset({}, 'A1', 1)")
    c.js("__calc.grid.focus()")
    inp.keytips("F M")
    time.sleep(0.3)
    menu = c.js("[...document.querySelectorAll('.menu [role=menuitem]')].map(b => [b.textContent, b.dataset.keytip || ''])")
    check("Alt, F, M opens Makros' menu with K, R, V", menu == [["Makros anzeigen", "K"], ["Makro aufzeichnen...", "R"], ["Relative Verweise verwenden", "V"]], menu)
    badges = c.js("[...document.querySelectorAll('.keytip')].map(b => b.textContent)")
    check("the menu shows its keytips", sorted(badges) == ["K", "R", "V"], badges)
    inp.key("k")
    time.sleep(0.4)
    check("K: Makros anzeigen (the Makro dialog)", c.js("[...document.querySelectorAll('.dialog')].some(d => d.textContent.includes('Makroname') && d.textContent.includes('Ausführen'))"))
    inp.key("esc")
    time.sleep(0.2)
    c.js("__calc.grid.focus()")
    inp.keytips("F M V")
    check("V: Relative Verweise verwenden on", c.js("__calc.app.recorder.relative") is True)
    inp.keytips("F M")
    time.sleep(0.2)
    check("the menu shows Relative Verweise verwenden checked", c.js("document.querySelector('.menu [aria-checked]')?.getAttribute('aria-checked')") == "true")
    inp.key("v")
    check("V again: off", c.js("__calc.app.recorder.relative") is False)
    c.js("__calc.grid.focus()")
    inp.keytips("F M R")
    time.sleep(0.3)
    inp.click_el("[...document.querySelectorAll('.dialog button')].find(b => b.textContent === 'OK')")
    time.sleep(0.3)
    check("R: Makro aufzeichnen starts a recording", c.js("__calc.app.recorder.recording") is True)
    check("the status bar says it is recording", c.js("document.querySelector('.statusbar .recbtn').getAttribute('aria-label')") == "Makroaufzeichnung Wird aufgezeichnet")
    c.js("__calc.grid.focus()")
    inp.keytips("F M")
    time.sleep(0.2)
    menu = c.js("[...document.querySelectorAll('.menu [role=menuitem]')].map(b => [b.textContent, b.dataset.keytip || ''])")
    check("while recording: Aufzeichnung beenden with A", ["Aufzeichnung beenden", "A"] in menu, menu)
    inp.key("a")
    time.sleep(0.2)
    check("A ends the recording", c.js("__calc.app.recorder.recording") is False)
    return out


def formulabar(c: CalcSide, inp: Cdp, ex: dict) -> list[dict]:
    out = []

    def check(name: str, ok: bool, detail=None) -> None:
        out.append({"check": name, "ok": bool(ok), "detail": detail})

    c.js("__calc.reset({'A1': 'x'}, 'A1', 1)")
    c.js("__calc.grid.focus()")
    scale = ex.get("scale", 1.5)
    nb0 = c.js("document.querySelector('.namebox-wrap').getBoundingClientRect().width")
    check("Namenfeld starts as wide as Excel's", abs(nb0 - ex["nameboxWidth"]["start"] / scale) <= 2, [nb0, ex["nameboxWidth"]["start"] / scale])
    check("the handle has the left-right pointer", c.js("getComputedStyle(document.querySelector('.fbar-sep')).cursor") == "ew-resize")
    r = inp.rect("document.querySelector('.fbar-sep')")
    cx, cy = r[0] + r[2] / 2, r[1] + r[3] / 2
    inp.drag(cx, cy, cx - 400, cy)
    w_min = c.js("document.querySelector('.namebox-wrap').getBoundingClientRect().width")
    check("dragged far left: Excel's smallest Namenfeld", abs(w_min - ex["nameboxWidth"]["min"] / scale) <= 2, [w_min, ex["nameboxWidth"]["min"] / scale])
    r = inp.rect("document.querySelector('.fbar-sep')")
    cx = r[0] + r[2] / 2
    inp.drag(cx, cy, cx + 160, cy)
    w_mid = c.js("document.querySelector('.namebox-wrap').getBoundingClientRect().width")
    check("dragged 160 px right: 160 px wider", abs(w_mid - (w_min + 160)) <= 2, [w_min, w_mid])
    r = inp.rect("document.querySelector('.fbar-sep')")
    cx = r[0] + r[2] / 2
    inp.drag(cx, cy, c.js("document.querySelector('.fbar').getBoundingClientRect().right") - 2, cy)
    w_max = c.js("document.querySelector('.namebox-wrap').getBoundingClientRect().width")
    room = c.js("document.querySelector('.fbar').getBoundingClientRect().right - document.querySelector('.namebox-wrap').getBoundingClientRect().right")
    want_room = (2000 - ex["namebox"][0] - ex["nameboxWidth"]["max"]) / scale  # what Excel leaves right of it in its 2000 px window
    check("dragged far right: the room Excel leaves the buttons and the bar", abs(room - want_room) <= 2, [w_max, room, want_room])
    c.js("document.querySelector('.namebox-wrap').style.width = ''")
    # Expanding.
    lines = lambda: c.js("__calc.grid.fbarLines")  # noqa: E731
    inp.click_el("document.querySelector('.fx-expand')")
    check("the chevron expands to Excel's height", lines() == ex.get("heightAfterExpandClick"), [lines(), ex.get("heightAfterExpandClick")])
    inp.click_el("document.querySelector('.fx-expand')")
    check("the chevron again collapses", lines() == 1, lines())
    c.js("__calc.grid.focus()")
    inp.key("ctrl+shift+u")
    # Recorded (debug run of probe_excel_formulabar's steps with screenshots): in German Excel Strg+Umschalt+U underlines
    # (FormulaBarHeight stays 1, Font.Underline 2); the formula bar has no shortcut there.
    check("Strg+Umschalt+U underlines, as in German Excel (the bar stays)", lines() == 1 and c.js("__calc.app.wb.styles.get(__calc.app.sheet.get(0, 0)?.s).font.underline") == "single",
          [lines()])
    inp.click_el("document.querySelector('.fx-expand')")
    r = inp.rect("document.querySelector('.fbar-edge')")
    x, y = r[0] + 300, r[1] + r[3] / 2
    inp.drag(x, y, x, y + 18 * 4)
    check("the lower edge dragged four lines down: four lines more", lines() == ex.get("heightAfterExpandClick", 3) + 4, lines())
    r = inp.rect("document.querySelector('.fbar-edge')")
    inp.drag(x, r[1] + r[3] / 2, x, r[1] - 18 * 30)
    check("dragged up: back to one line", lines() == 1, lines())
    return out


SLICER_SETUP = r"""(() => {
  const app = __calc.app, sl = app.slicers, P = __calc.parse;
  const t = app.tables.create(P("A1:B7"), true); t.name = "Faecher";
  sl.add([{ sourceName: "Fach", table: { name: "Faecher", column: 1 } }]);
  sl.setMulti(false);
  return sl.items(sl.cacheOf(sl.byName("Fach"))).map((i) => i.text);
})()"""
CHOSEN = "(() => { const sl = __calc.app.slicers; return sl.items(sl.cacheOf(sl.byName('Fach'))).filter((i) => i.selected).map((i) => i.text); })()"


def slicers(c: CalcSide, inp: Cdp, ex: dict) -> list[dict]:
    out = []
    data = {"A1": "Name", "B1": "Fach", "A2": "Anna", "B2": "Mathe", "A3": "Ben", "B3": "Deutsch", "A4": "Cem", "B4": "Englisch",
            "A5": "Dora", "B5": "Bio", "A6": "Emil", "B6": "Kunst", "A7": "Finn", "B7": "Mathe"}
    items = ex.get("items", [])

    def btn(i: int) -> str:
        return f"[...document.querySelectorAll('.slicer-layer .sl-btn')].find(b => b.title === {json.dumps(items[i])})"

    # One slicer for all cases, its filter cleared between them, as the Excel probe did (Umschalt+click's start, the item
    # clicked last, carries over from one case to the next there too).
    c.js(f"__calc.reset({json.dumps(data)}, 'D2', 1)")
    shown = c.js(SLICER_SETUP)
    for name, case in ex.get("cases", {}).items():
        c.js("(() => { const sl = __calc.app.slicers; sl.select(sl.cacheOf(sl.byName('Fach')), null); __calc.grid.selectCell(1, 3); })()")
        time.sleep(0.3)
        trail = []
        for st in case["steps"]:
            kind = st[0]
            if kind == "click":
                r = inp.rect(btn(st[1]))
                inp.click(r[0] + r[2] / 2, r[1] + r[3] / 2, 1, MODS.get(st[2], 0) if st[2] else 0)
            elif kind == "drag":
                a, b = inp.rect(btn(st[1])), inp.rect(btn(st[2]))
                inp.drag(a[0] + a[2] / 2, a[1] + a[3] / 2, b[0] + b[2] / 2, b[1] + b[3] / 2, mods=MODS.get(st[3], 0) if st[3] else 0)
            elif kind == "keys":
                inp.key(st[1])
            elif kind == "multi":
                inp.click_el("document.querySelector('.slicer-layer [data-act=multi]')")
            time.sleep(0.2)
            trail.append({"selected": c.js(CHOSEN), "multi": c.js("document.querySelector('.slicer-layer [data-act=multi]')?.classList.contains('pressed') ?? null")})
        want = case["final"]
        out.append({"check": f"Datenschnitt {name}", "ok": trail[-1]["selected"] == want and shown == items, "detail": {"excel": want, "raabecalc": trail}})
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    report: dict = {}
    c = CalcSide()
    inp = Cdp(c)
    try:
        report["keytips"] = keytips(c, inp)
        fb = load("2026-11-09-excel-formulabar.json")
        if fb:
            report["formulabar"] = formulabar(c, inp, fb)
        sm = load("2026-11-09-excel-slicer-multi.json")
        if sm:
            report["slicers"] = slicers(c, inp, sm)
    finally:
        c.close()
    counts = {k: f"{sum(x['ok'] for x in v)}/{len(v)}" for k, v in report.items()}
    for k, v in report.items():
        for x in v:
            print(("OK  " if x["ok"] else "DIFF") + f" {k}: {x['check']}" + ("" if x["ok"] else f"  {json.dumps(x['detail'], ensure_ascii=False)[:400]}"))
    print(counts)
    RESULT.write_text(json.dumps({"counts": counts, "report": report}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
