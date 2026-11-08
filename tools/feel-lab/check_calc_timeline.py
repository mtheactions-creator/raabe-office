"""Zeitachsen against Excel (D-121), both ways, without real input:
1. Excel -> RaabeCalc: RaabeCalc (off screen, its own DevTools port) opens Excel's own file
   (packages/ooxml/test/fixtures/excel-timeline.xlsx from probe_excel_timeline.py) and must show what Excel recorded
   through COM: each timeline's selection (start and end day), level, Anzeigen options and caption, the selection label,
   and each PivotTable's report with Excel's values.
2. RaabeCalc by hand -> Excel: a table of dated marks and a PivotTable; then, like a student, with trusted DevTools input
   into the real exe (Input.dispatchKeyEvent / dispatchMouseEvent, the OS mouse is not touched): the keytips Alt, I, Ö
   (Einfügen > Zeitachse), Datum ticked in Zeitachsen einfügen, OK; a click on a month; Umschalt+click on a later month;
   dragging the selection's right edge; the Zeitebene menu to Quartale and dragging across two quarters; the mouse
   wheel; Filter löschen in the header; Kopfzeile off through the keytips Alt, J, L, K. After each step RaabeCalc saves;
   our own invisible Excel opens the file and must report the same through COM: the timeline's level and options, its
   cache's filter (TimelineState start and end) and the PivotTable's report.
Result: results/<date>-calc-timeline.json.

    .venv/Scripts/python.exe check_calc_timeline.py
"""
from __future__ import annotations

import datetime
import json
import os
import tempfile
import time
from pathlib import Path

from feellab import calc_catalogue, winapi
from feellab.calc_catalogue import Case, CalcSide
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
TMP = Path(tempfile.gettempdir())
FIX = ROOT / "packages" / "ooxml" / "test" / "fixtures"
REC = json.loads((FIX / "excel-timeline.json").read_text(encoding="utf-8"))

os.environ["RAABE_FEEL_BACKGROUND"] = "1"
calc_catalogue.winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4000, 0, w, h)
PORT = 9398
EPOCH = datetime.date(1899, 12, 30)


def serial(iso: str | None) -> int | None:
    return None if not iso else (datetime.date.fromisoformat(iso[:10]) - EPOCH).days


def cdp(self: CalcSide, method: str, params: dict):
    self._id += 1
    self.ws.send(json.dumps({"id": self._id, "method": method, "params": params}))
    while True:
        msg = json.loads(self.ws.recv())
        if msg.get("id") == self._id:
            return msg.get("result")


CalcSide.cdp = cdp  # type: ignore[attr-defined]

KEYS = {"Alt": ("AltLeft", 18), "Shift": ("ShiftLeft", 16), "ö": ("Semicolon", 192)}


def key(c: CalcSide, k: str, mods: int = 0) -> None:
    code, vk = KEYS.get(k, (f"Key{k.upper()}", ord(k.upper())))
    text = k if len(k) == 1 else ""
    c.cdp("Input.dispatchKeyEvent", {"type": "keyDown" if text else "rawKeyDown", "key": k, "code": code, "windowsVirtualKeyCode": vk, "modifiers": mods, **({"text": text} if text else {})})
    c.cdp("Input.dispatchKeyEvent", {"type": "keyUp", "key": k, "code": code, "windowsVirtualKeyCode": vk, "modifiers": mods})
    time.sleep(0.12)


def keytips(c: CalcSide, *letters: str) -> None:
    key(c, "Alt")
    time.sleep(0.25)
    for ch in letters:
        key(c, ch)
        time.sleep(0.2)
    time.sleep(0.4)


def mouse(c: CalcSide, kind: str, x: float, y: float, buttons: int = 0, mods: int = 0) -> None:
    c.cdp("Input.dispatchMouseEvent", {"type": kind, "x": x, "y": y, "button": "left" if kind != "mouseMoved" or buttons else "none", "buttons": buttons,
                                       "clickCount": 1 if kind != "mouseMoved" else 0, "modifiers": mods})


def click(c: CalcSide, x: float, y: float, mods: int = 0) -> None:
    mouse(c, "mouseMoved", x, y, 0, mods)
    mouse(c, "mousePressed", x, y, 1, mods)
    mouse(c, "mouseReleased", x, y, 0, mods)
    time.sleep(0.3)


def drag(c: CalcSide, x0: float, y0: float, x1: float, y1: float, steps: int = 12) -> None:
    mouse(c, "mouseMoved", x0, y0)
    mouse(c, "mousePressed", x0, y0, 1)
    for i in range(1, steps + 1):
        mouse(c, "mouseMoved", x0 + (x1 - x0) * i / steps, y0 + (y1 - y0) * i / steps, 1)
    mouse(c, "mouseReleased", x1, y1)
    time.sleep(0.3)


FACH = ["Mathe", "Deutsch", "Englisch", "Biologie"]
DATA = {"A1": "Datum", "B1": "Fach", "C1": "Note"}
for i in range(24):
    d = datetime.date(2025, 8, 18) + datetime.timedelta(days=17 * i)
    DATA[f"A{i + 2}"] = d.strftime("%d.%m.%Y")
    DATA[f"B{i + 2}"] = FACH[i % 4]
    DATA[f"C{i + 2}"] = str(1 + (i * 7) % 6)

PIVOT = r"""(() => { const app = __calc.app, s = app.sheet;
  app.pivots.add(s, { name: "PivotNoten", source: { sheet: s.name, r0: 0, c0: 0, r1: 24, c1: 2 }, at: { row: 2, col: 5 }, rows: [1], cols: [], values: [{ field: 2, func: "count" }], filters: [] });
  app.grid.selectCell(3, 5); app.grid.focus(); return app.pivots.current()?.name; })()"""

STATE = r"""(() => { const app = __calc.app, t = app.timelines.byName("Datum"); if (!t) return null;
  const p = app.sheet.pivots[0]; const rows = [];
  for (let r = p.area.r0; r <= p.area.r1; r++) { const row = []; for (let c = p.area.c0; c <= p.area.c1; c++) row.push(app.sheet.value(r, c)); rows.push(row); }
  return { sel: app.timelines.selection(t), level: t.level, header: t.showHeader !== false, label: app.timelines.label("Datum"), report: rows }; })()"""


def block(c: CalcSide, y: int, m: int, d: int = 1) -> dict:
    s = (datetime.date(y, m, d) - EPOCH).days
    p = c.js(f"__calc.app.timelines.blockPoint('Datum', {s})")
    if not p:
        raise RuntimeError(f"block {y}-{m} not shown")
    return p


def excel_state(x: ExcelTarget, path: Path) -> dict:
    wb = x.app.Workbooks.Open(str(path))
    try:
        out: dict = {"timelines": []}
        for i in range(1, wb.SlicerCaches.Count + 1):
            sc = wb.SlicerCaches(i)
            if sc.SlicerCacheType != 2:
                continue
            ts = sc.TimelineState
            try:
                st, en = (ts.StartDate, ts.EndDate) if ts.FilterType else (None, None)
            except Exception:  # noqa: BLE001
                st = en = None
            for k in range(1, sc.Slicers.Count + 1):
                sl = sc.Slicers(k)
                vs = sl.TimelineViewState
                out["timelines"].append({"name": sl.Name, "level": vs.Level, "header": bool(vs.ShowHeader),
                                         "sel": None if st is None else [serial(str(st)), serial(str(en))]})
        ws = wb.Worksheets(1)
        pt = ws.PivotTables(1)
        out["report"] = [[None if v is None else v for v in r] for r in pt.TableRange1.Value]
        return out
    finally:
        wb.Close(False)


def same_report(a, b) -> bool:
    norm = lambda rows: [[(float(v) if isinstance(v, (int, float)) else v) for v in r] for r in rows]  # noqa: E731
    return norm(a) == norm(b)


def main() -> None:
    report: dict = {"excel->raabe": [], "raabe->excel": []}
    states: list[tuple[str, dict, dict]] = []
    c = CalcSide(PORT)
    try:
        # 1. Excel's file in RaabeCalc.
        c.js(f"__calc.open({json.dumps(str(FIX / 'excel-timeline.xlsx'))})")
        time.sleep(0.8)
        for name, case in REC["cases"].items():
            if name == "beforeClear":
                continue
            want = case["timelines"][0]
            got = c.js(f"""(() => {{ const app = __calc.app; const i = app.wb.sheets.findIndex((s) => s.name === {json.dumps(case['sheet'])}); app.activateSheet(i);
              const t = app.timelines.byName({json.dumps(want['name'])}); if (!t) return null; app.grid.contentChanged();
              const ps = app.sheet.pivots.map((p) => {{ const rows = []; for (let r = p.area.r0; r <= p.area.r1; r++) {{ const row = []; for (let c = p.area.c0; c <= p.area.c1; c++) row.push(app.sheet.value(r, c)); rows.push(row); }} return rows; }});
              return {{ sel: app.timelines.selection(t), level: t.level, caption: t.caption, style: t.style ?? "TimeSlicerStyleLight1",
                opts: [t.showHeader !== false, t.showSelectionLabel !== false, t.showTimeLevel !== false, t.showHorizontalScrollbar !== false], reports: ps }}; }})()""")
            time.sleep(0.2)
            label = c.js(f"__calc.app.timelines.label({json.dumps(want['name'])})")
            checks = []
            if got is None:
                checks.append(("timeline", False, "missing"))
            else:
                ws = want["state"]
                exp_sel = None if not ws["startDate"] else {"from": serial(ws["startDate"]), "to": serial(ws["endDate"])}
                checks.append(("selection", got["sel"] == exp_sel, f"{got['sel']} / {exp_sel}"))
                checks.append(("level", got["level"] == want["view"]["level"], f"{got['level']} / {want['view']['level']}"))
                checks.append(("caption", got["caption"] == want["caption"], got["caption"]))
                checks.append(("style", got["style"] == want["style"], got["style"]))
                v = want["view"]
                checks.append(("options", got["opts"] == [v["showHeader"], v["showSelectionLabel"], v["showTimeLevel"], v["showScrollbar"]], str(got["opts"])))
                for k, r in enumerate(case["reports"]):
                    checks.append((f"report{k}", k < len(got["reports"]) and same_report(got["reports"][k], r["values"]), ""))
                report.setdefault("labels", {})[name] = label
            for what, ok, info in checks:
                report["excel->raabe"].append({"case": name, "check": what, "ok": ok, "info": info})

        # 2. By hand in RaabeCalc.
        c.reset(Case("timeline", DATA, [], "A1", sheets=1))
        c.js(PIVOT)
        time.sleep(0.3)
        c.cdp("Page.bringToFront", {})
        keytips(c, "i", "ö")
        dlg = c.js("""(() => { const b = [...document.querySelectorAll('[role=dialog] input[type=checkbox]')].find((x) => x.getAttribute('aria-label') === 'Datum');
          if (!b) return null; const r = b.getBoundingClientRect(); const ok = [...document.querySelectorAll('[role=dialog] button')].find((x) => x.textContent === 'OK').getBoundingClientRect();
          return { title: document.querySelector('[role=dialog]')?.textContent?.slice(0, 40), box: [r.left + r.width / 2, r.top + r.height / 2], ok: [ok.left + ok.width / 2, ok.top + ok.height / 2] }; })()""")
        report["dialog"] = dlg
        if not dlg:
            raise RuntimeError("Zeitachsen einfügen did not open")
        click(c, *dlg["box"])
        click(c, *dlg["ok"])
        time.sleep(0.5)

        def snap(name: str, expect: dict) -> None:
            path = TMP / f"raabe-timeline-{name}.xlsx"
            if path.exists():
                path.unlink()
            st = c.js(STATE)
            c.js(f"__calc.save({json.dumps(str(path))})")
            states.append((name, {**st, "path": str(path)}, expect))

        snap("inserted", {"sel": None, "level": 2})
        p = block(c, 2025, 10)
        click(c, p["x"], p["y"])
        snap("click", {"sel": [(datetime.date(2025, 10, 1) - EPOCH).days, (datetime.date(2025, 10, 31) - EPOCH).days], "level": 2})
        p = block(c, 2026, 1)
        click(c, p["x"], p["y"], 8)  # Umschalt
        snap("shiftclick", {"sel": [(datetime.date(2025, 10, 1) - EPOCH).days, (datetime.date(2026, 1, 31) - EPOCH).days], "level": 2})
        # The selection's right edge from January to March.
        p, q = block(c, 2026, 1), block(c, 2026, 3)
        drag(c, p["right"], p["y"], q["x"] + 2, q["y"])
        snap("edge", {"sel": [(datetime.date(2025, 10, 1) - EPOCH).days, (datetime.date(2026, 3, 31) - EPOCH).days], "level": 2})
        # Zeitebene: Quartale through its menu, then a drag across Q1 and Q2 2026.
        lv = c.js("(() => { const e = document.querySelector('.timeline .tl-lvl'); const r = e.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; })()")
        click(c, *lv)
        item = c.js("(() => { const e = [...document.querySelectorAll('.menu button, [role=menu] button, [role=menuitem]')].find((x) => x.textContent.trim().startsWith('Quartale')); if (!e) return null; const r = e.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; })()")
        if item:
            click(c, *item)
        time.sleep(0.4)
        a, b = block(c, 2026, 1), block(c, 2026, 4)
        drag(c, a["x"], a["y"], b["x"], b["y"])
        snap("quarters", {"sel": [(datetime.date(2026, 1, 1) - EPOCH).days, (datetime.date(2026, 6, 30) - EPOCH).days], "level": 1})
        # Filter löschen in the header.
        cl = c.js("(() => { const e = document.querySelector('.timeline [data-act=clear]'); const r = e.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; })()")
        click(c, *cl)
        snap("cleared", {"sel": None, "level": 1})
        # Kopfzeile off through the contextual tab's keytips.
        keytips(c, "j", "l", "k")
        snap("noheader", {"sel": None, "level": 1, "header": False})
    finally:
        c.close()

    x = ExcelTarget(visible=False)
    try:
        for name, ours, expect in states:
            ex = excel_state(x, Path(ours["path"]))
            tl = ex["timelines"][0] if ex["timelines"] else None
            checks = [
                ("ours as intended", ours["sel"] == (None if expect["sel"] is None else {"from": expect["sel"][0], "to": expect["sel"][1]}) and ours["level"] == expect["level"], f"{ours['sel']} {ours['level']} {ours['label']}"),
                ("Excel finds the timeline", tl is not None, ""),
            ]
            if tl:
                checks.append(("Excel's filter", tl["sel"] == (None if ours["sel"] is None else [ours["sel"]["from"], ours["sel"]["to"]]), f"{tl['sel']}"))
                checks.append(("Excel's level", tl["level"] == ours["level"], f"{tl['level']}"))
                checks.append(("Excel's header", tl["header"] == ours["header"], f"{tl['header']}"))
            checks.append(("Excel's report", same_report(ex["report"], ours["report"]), f"{ex['report']} / {ours['report']}"[:300]))
            for what, ok, info in checks:
                report["raabe->excel"].append({"step": name, "check": what, "ok": ok, "info": info})
    finally:
        x.close()

    for k in ("excel->raabe", "raabe->excel"):
        ok = sum(r["ok"] for r in report[k])
        print(f"{k}: {ok}/{len(report[k])}")
        for r in report[k]:
            if not r["ok"]:
                print("   DIFF", r)
    report["summary"] = {k: f"{sum(r['ok'] for r in report[k])}/{len(report[k])}" for k in ("excel->raabe", "raabe->excel")}
    path = HERE / "results" / "2026-11-09-calc-timeline.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
