"""The chart Format pane by hand (D-111): every control of every element's pane on the main chart
types, driven through the real RaabeCalc release exe with trusted DevTools input (mouse clicks,
keys and typed text into its WebView2; the OS mouse and keyboard are not touched; the window stays
off screen and is never focused).

For each chart type (column, bar, line with markers, pie, scatter, area, Verbund) the script
inserts a chart, selects it with a click, opens the pane with Strg+1 and walks through the
elements with the pane's own element list (Diagrammoptionen ▾). In every element's pane it opens
each option set, tab and section and uses every control once as a student would: clicks radio
buttons, check boxes, colour buttons (then a swatch), pattern tiles and stop buttons; picks the
next entry of a dropdown with the arrow key; types a value into a number box (and clicks its
spin button); drags a slider; types into a text box. After each use it records
  (a) whether the chart model changed and whether the drawing (the SVG) changed,
  (b) whether Strg+Z brings back the model and the drawing as they were.
The models after each change are then saved in .xlsx files (one per chart type, one chart per
change), opened again in RaabeCalc (the setting must come back) and in our own invisible Excel
(DispatchEx), which reports the setting through COM where a COM property exists for it and
exports each chart as a picture: the picture must differ from the chart without the change
wherever ours does.

    .venv/Scripts/python.exe check_calc_chartpane.py [--kinds column,pie] [--no-excel]

Result: results/2026-11-09-calc-chartpane.json; prints a summary and every failure.
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import os
import tempfile
import time
from pathlib import Path

from feellab import calc_catalogue, winapi
from feellab.calc_catalogue import Case, CalcSide

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir()) / "raabe-chartpane"
PORT = 9417
DATA = {"A1": "Monat", "B1": "A", "C1": "B", "A2": "Jan", "B2": "3", "C2": "40", "A3": "Feb", "B3": "-1", "C3": "20", "A4": "Mrz", "B4": "5", "C4": "60",
        "A5": "Apr", "B5": "2", "C5": "30", "D1": "Text", "D2": "eins", "D3": "zwei", "D4": "drei", "D5": "vier"}
KINDS = ["column", "bar", "lineMarkers", "pie", "scatter", "area", "combo"]
# Buttons that open a native dialog or read the clipboard are not pressed by the sweep.
SKIP_BUTTONS = {"Einfügen...", "Zwischenablage"}
# Controls whose effect is not drawn (stored only): their unchanged picture is correct.
INVISIBLE = {"Gesperrt", "Objekt drucken", "Titel", "Beschreibung", "Seitenverhältnis sperren", "Nur von Zellposition abhängig",
             "Von Zellposition und -größe unabhängig", "Von Zellposition und -größe abhängig", "Mit Form drehen", "Bereich der Datenbeschriftung",
             # Vertical alignment inside a box that fits its text, a unit of tick marks or gridlines not shown: Excel shows no change either.
             "Vertikale Ausrichtung", "Teilstriche"}


def calc_side(port: int = PORT) -> CalcSide:
    """RaabeCalc off screen and never focused (its window goes left of the screen)."""
    fg = winapi.foreground()
    real = calc_catalogue.winapi.place
    calc_catalogue.winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4200, 0, 1900, 1300)
    try:
        c = CalcSide(port)
    finally:
        calc_catalogue.winapi.place = real
    if fg and winapi.foreground() != fg:
        try:
            winapi.activate(fg, 0)
        except Exception:  # noqa: BLE001
            pass
    return c


class Driver:
    """Trusted input through the DevTools protocol (Input.dispatch*) into RaabeCalc's page."""

    def __init__(self, c: CalcSide):
        self.c = c

    def cdp(self, method: str, params: dict):
        c = self.c
        c._id += 1
        c.ws.send(json.dumps({"id": c._id, "method": method, "params": params}))
        while True:
            msg = json.loads(c.ws.recv())
            if msg.get("id") == c._id:
                return msg.get("result")

    def js(self, expr: str):
        return self.c.js(expr)

    def click(self, x: float, y: float, count: int = 1) -> None:
        self.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x, "y": y, "button": "none", "buttons": 0})
        for n in range(1, count + 1):
            self.cdp("Input.dispatchMouseEvent", {"type": "mousePressed", "x": x, "y": y, "button": "left", "buttons": 1, "clickCount": n})
            self.cdp("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": x, "y": y, "button": "left", "buttons": 0, "clickCount": n})
        time.sleep(0.05)

    def drag(self, x0: float, y0: float, x1: float, y1: float) -> None:
        self.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x0, "y": y0, "button": "none", "buttons": 0})
        self.cdp("Input.dispatchMouseEvent", {"type": "mousePressed", "x": x0, "y": y0, "button": "left", "buttons": 1, "clickCount": 1})
        for k in range(1, 6):
            self.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x0 + (x1 - x0) * k / 5, "y": y0 + (y1 - y0) * k / 5, "button": "left", "buttons": 1})
        self.cdp("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": x1, "y": y1, "button": "left", "buttons": 0, "clickCount": 1})
        time.sleep(0.05)

    KEYS = {"Enter": 13, "ArrowDown": 40, "ArrowUp": 38, "ArrowLeft": 37, "ArrowRight": 39, "Escape": 27, "Tab": 9, "a": 65, "z": 90, "1": 49, "Home": 36, "End": 35}

    def key(self, key: str, ctrl: bool = False) -> None:
        code = self.KEYS[key]
        mods = 2 if ctrl else 0
        dom = {"a": "KeyA", "z": "KeyZ", "1": "Digit1"}.get(key, key)
        text = "\r" if key == "Enter" else ""
        self.cdp("Input.dispatchKeyEvent", {"type": "rawKeyDown" if ctrl or not text else "keyDown", "key": key, "code": dom, "windowsVirtualKeyCode": code, "modifiers": mods, **({"text": text} if text else {})})
        self.cdp("Input.dispatchKeyEvent", {"type": "keyUp", "key": key, "code": dom, "windowsVirtualKeyCode": code, "modifiers": mods})
        time.sleep(0.04)

    def type(self, text: str) -> None:
        self.cdp("Input.insertText", {"text": text})


# The page's view of the pane: every visible control with its kind, label, centre and state.
PANE = r"""(() => {
  const p = document.querySelector('.taskpane');
  if (!p || p.hidden) return null;
  const r = (el) => { const b = el.getBoundingClientRect(); return [b.left + b.width / 2, b.top + b.height / 2, b.width, b.height]; };
  const sets = [...p.querySelectorAll('.tp-set')].map((b) => ({ label: b.textContent.replace('˅', '').trim(), on: b.classList.contains('on'), at: r(b) }));
  const tabs = [...p.querySelectorAll('.tp-tab')].map((b) => ({ label: b.getAttribute('aria-label'), on: b.classList.contains('on'), at: r(b) }));
  const secs = [...p.querySelectorAll('.tp-sec')].map((s, si) => {
    const head = s.querySelector('.tp-sec-head');
    const ctrls = [];
    for (const [ci, el] of [...s.querySelectorAll('.tp-sec-body > *')].entries()) {
      el.dataset.sweep = si + ':' + ci;
      const label = (el.querySelector('.tp-label')?.textContent ?? el.textContent ?? '').trim();
      const inp = el.querySelector('input,select,button');
      let kind = 'other';
      if (el.classList.contains('tp-heading')) kind = 'heading';
      else if (el.querySelector('input[type=radio]')) kind = 'radio';
      else if (el.querySelector('input[type=checkbox]')) kind = 'check';
      else if (el.querySelector('select')) kind = 'select';
      else if (el.querySelector('.picdrop')) kind = 'picture';
      else if (el.querySelector('input.tp-num')) kind = 'number';
      else if (el.querySelector('.tp-color')) kind = 'color';
      else if (el.querySelector('input.tp-text')) kind = 'text';
      else if (el.querySelector('.tp-btn')) kind = 'button';
      else if (el.classList.contains('tp-grad')) kind = 'gradient';
      else if (el.classList.contains('tp-pattern-box')) kind = 'pattern';
      const disabled = !!(inp && inp.disabled);
      const checked = !!el.querySelector('input:checked');
      const c = { kind, label, disabled, checked, id: si + ':' + ci };
      if (kind === 'radio' || kind === 'check') c.at = r(el.querySelector('input'));
      if (kind === 'select') { const s = el.querySelector('select'); c.at = r(s); c.value = s.value; c.n = s.options.length; c.index = s.selectedIndex; }
      if (kind === 'picture') { const b = el.querySelector('.picdrop'); c.at = r(b); c.value = b.dataset.value; c.disabled = b.disabled; }
      if (kind === 'number') { c.at = r(el.querySelector('input.tp-num')); c.value = el.querySelector('input.tp-num').value; const up = el.querySelector('.tp-spin button'); if (up) c.up = r(up); const sl = el.querySelector('.tp-slider'); if (sl) { c.slider = r(sl); } const rs = el.querySelector('.tp-reset'); if (rs && !rs.disabled) c.reset = r(rs); }
      if (kind === 'color') c.at = r(el.querySelector('.tp-color'));
      if (kind === 'text') { c.at = r(el.querySelector('input.tp-text')); c.value = el.querySelector('input.tp-text').value; }
      if (kind === 'button') c.at = r(el.querySelector('.tp-btn'));
      if (kind === 'gradient') { const bs = el.querySelectorAll('.tp-icon-btn'); c.add = r(bs[0]); c.del = bs[1] && !bs[1].disabled ? r(bs[1]) : null; const st = el.querySelectorAll('.tp-grad-stop'); c.stop = st.length > 1 ? r(st[1]) : null; c.bar = r(el.querySelector('.tp-grad-bar')); }
      if (kind === 'pattern') { const t = [...el.querySelectorAll('.tp-pattern')]; const on = t.findIndex((x) => x.classList.contains('on')); c.at = r(t[(on + 7) % t.length]); }
      ctrls.push(c);
    }
    return { title: head.textContent.replace(/^[◢▷]/, '').trim(), open: head.getAttribute('aria-expanded') === 'true', at: r(head), controls: ctrls };
  });
  const title = p.querySelector('.tp-title')?.textContent;
  return { title, sets, tabs, secs };
})()"""

STATE = r"""(() => {
  const app = __calc.app;
  const k = app.sheet.charts.length - 1;
  const m = app.sheet.charts[k];
  const { raw, from, to, ...rest } = m;
  const svg = document.querySelectorAll('.chart-obj > svg')[k];
  // The drawing without its random definition ids.
  const pic = svg ? svg.innerHTML.replace(/\b(cd|rc|cm)[a-z0-9]{5,}\d*/g, 'id') : '';
  return { model: JSON.stringify(rest), from: JSON.stringify([from, to]), pic, el: app.charts.elementName() };
})()"""


def h(s: str) -> str:
    return hashlib.sha1(s.encode("utf-8")).hexdigest()[:12]


class Sweep:
    def __init__(self, d: Driver, kind: str):
        self.d = d
        self.kind = kind
        self.results: list[dict] = []
        self.models: list[dict] = []  # models after each change (for the files)

    def pane(self):
        return self.d.js(PANE)

    def state(self):
        return self.d.js(STATE)

    def setup(self) -> None:
        d = self.d
        d.c.reset(Case("chartpane", DATA, [], "A1"))
        kind = "column" if self.kind == "combo" else self.kind
        rng = "{r0:0,c0:0,r1:4,c1:2}"
        d.js(f"__calc.app.grid.select({rng}); __calc.app.charts.insert({json.dumps(kind)})")
        if self.kind == "combo":
            # Verbund: series B as a line on the Sekundärachse (Diagrammtyp ändern does the same).
            d.js("__calc.app.charts.changeSelected((m) => { m.series[1].kind = 'lineMarkers'; m.series[1].secondary = true; })")
        # Datenbeschriftungen, a Trendlinie and Fehlerindikatoren, so that their panes are in the element list (Diagrammelemente).
        d.js("__calc.app.charts.changeSelected((m) => { m.dataLabels = true; if (!['pie','doughnut'].includes(m.kind)) { m.series[0].trendline = { type: 'linear', showEq: false, showR2: false }; m.series[0].errorBars = { dir: 'both', type: 'fixedVal', value: 1 }; } })")
        time.sleep(0.3)
        # Select the chart with a click on its area and open the pane with Strg+1.
        box = d.js("(() => { const b = document.querySelectorAll('.chart-obj')[document.querySelectorAll('.chart-obj').length - 1].getBoundingClientRect(); return [b.left + 12, b.top + 12]; })()")
        d.click(*box)
        d.key("1", ctrl=True)
        time.sleep(0.2)

    def elements(self) -> list[str]:
        return self.d.js("__calc.app.charts.elementNames()")

    def choose(self, name: str) -> bool:
        """Picks an element in the pane's own list (the first option set's dropdown)."""
        p = self.pane()
        if not p:
            return False
        first = p["sets"][0]
        if not first["on"]:
            self.d.click(*first["at"][:2])
            time.sleep(0.05)
        self.d.click(*self.pane()["sets"][0]["at"][:2])
        time.sleep(0.1)
        at = self.d.js(f"(() => {{ const it = [...document.querySelectorAll('.popup .menu .item, .menu .item')].find((b) => b.textContent === {json.dumps(name)}); if (!it) return null; const r = it.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; }})()")
        if not at:
            self.d.key("Escape")
            return False
        self.d.click(*at)
        time.sleep(0.15)
        return self.d.js("__calc.app.charts.elementName()") == name

    def undo(self) -> None:
        # The focus leaves the pane's box (Esc as a student would press it), then Strg+Z.
        self.d.js("document.activeElement && document.activeElement.blur && document.activeElement.blur()")
        self.d.key("z", ctrl=True)
        time.sleep(0.15)

    def use(self, c: dict) -> str | None:
        """Uses one control once; returns what was done or None when it is skipped."""
        d = self.d
        k = c["kind"]
        if k in ("heading", "other") or c.get("disabled"):
            return None
        if k == "radio":
            if c["checked"]:
                return None
            d.click(*c["at"][:2])
            return "click"
        if k == "check":
            d.click(*c["at"][:2])
            return "click"
        if k == "select":
            d.js(f"(() => {{ const s = [...document.querySelectorAll('.taskpane select')].find((x) => {{ const b = x.getBoundingClientRect(); return Math.abs(b.left + b.width / 2 - {c['at'][0]}) < 2 && Math.abs(b.top + b.height / 2 - {c['at'][1]}) < 2; }}); s && s.focus(); }})()")
            d.key("ArrowUp" if c["index"] == c["n"] - 1 else "ArrowDown")
            return "arrow key"
        if k == "picture":
            # A picture dropdown (D-161): opened with a click, the next picture by arrow key, picked with Enter.
            d.click(*c["at"][:2])
            time.sleep(0.1)
            last = d.js("(() => { const l = document.querySelector('.picdrop-list'); if (!l) return null; const i = [...l.querySelectorAll('.pd-item')]; return i.findIndex((x) => x === document.activeElement) === i.length - 1; })()")
            if last is None:
                return None
            d.key("ArrowLeft" if last else "ArrowRight")
            d.key("Enter")
            return "picture"
        if k == "number":
            if c.get("slider"):
                x, y, w, _ = c["slider"]
                d.click(x + w * 0.3, y)
                return "slider"
            d.click(*c["at"][:2], count=3)
            d.key("a", ctrl=True)
            d.type(number_for(c))
            d.key("Enter")
            return "typed " + number_for(c)
        if k == "color":
            d.click(*c["at"][:2])
            time.sleep(0.1)
            sw = d.js("(() => { const s = [...document.querySelectorAll('.menu.palette .swatches button, .menu.palette .swatch')]; const b = s[25] ?? s[0]; if (!b) return null; const r = b.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; })()")
            if not sw:
                d.key("Escape")
                return None
            d.click(*sw)
            return "swatch"
        if k == "text":
            d.click(*c["at"][:2], count=3)
            d.key("a", ctrl=True)
            d.type("Probe 1")
            d.key("Enter")
            return "typed text"
        if k == "button":
            if c["label"] in SKIP_BUTTONS:
                return None
            d.click(*c["at"][:2])
            return "click"
        if k == "gradient":
            d.click(*c["add"][:2])
            return "add stop"
        if k == "pattern":
            d.click(*c["at"][:2])
            return "pattern"
        return None

    def run(self, only: str | None = None) -> None:
        import re
        self.setup()
        for el in self.elements():
            if only and not re.search(only, el):
                continue
            if not self.choose(el):
                self.results.append({"kind": self.kind, "element": el, "control": "(element list)", "ok": False, "why": "not selectable in the pane's list"})
                continue
            self.sweep_element(el)

    def sweep_element(self, el: str) -> None:
        d = self.d
        p = self.pane()
        for si in range(len(p["sets"])):
            p = self.pane()
            if not p["sets"][si]["on"]:
                d.click(*p["sets"][si]["at"][:2])
                time.sleep(0.05)
                p = self.pane()
                if not p["sets"][si]["on"]:
                    d.key("Escape")
                    continue
            for ti in range(len(p["tabs"])):
                p = self.pane()
                if ti >= len(p["tabs"]):
                    break
                d.click(*p["tabs"][ti]["at"][:2])
                time.sleep(0.05)
                p = self.pane()
                tab = p["tabs"][ti]["label"]
                for ki in range(len(p["secs"])):
                    p = self.pane()
                    sec = p["secs"][ki]
                    if not sec["open"]:
                        d.click(*sec["at"][:2])
                        time.sleep(0.05)
                    self.sweep_section(el, p["sets"][si]["label"], tab, ki)

    def sweep_section(self, el: str, set_label: str, tab: str, ki: int) -> None:
        p = self.pane()
        if not p or ki >= len(p["secs"]):
            return
        for i in range(len(p["secs"][ki]["controls"])):
            p = self.pane()
            if not p or ki >= len(p["secs"]):
                break
            ctrls = p["secs"][ki]["controls"]
            if i >= len(ctrls):
                break
            self.use_all(el, set_label, tab, ki, ctrls[i], 0)

    def use_all(self, el: str, set_label: str, tab: str, ki: int, c: dict, depth: int) -> None:
        sec = self.pane()["secs"][ki]["title"]
        self.one(el, set_label, tab, sec, ki, c, depth)
        if c["kind"] == "number" and c.get("up") and not c.get("disabled"):
            self.one(el, set_label, tab, sec, ki, {**c, "kind": "spin"}, depth)

    def one(self, el: str, set_label: str, tab: str, sec: str, ki: int, c: dict, depth: int) -> None:
        d = self.d
        # The control into view (the pane's body scrolls), then its place again.
        d.js(f"(() => {{ const e = document.querySelector('.taskpane [data-sweep=\"{c['id']}\"]'); e && e.scrollIntoView({{block: 'center'}}); }})()")
        time.sleep(0.03)
        p = self.pane()
        fresh = next((x for x in (p["secs"][ki]["controls"] if p and ki < len(p["secs"]) else []) if x["id"] == c["id"] and x["label"] == c["label"]), None)
        if fresh:
            c = {**fresh, "kind": c["kind"]} if c["kind"] == "spin" else fresh
        before = self.state()
        shown = {(x["kind"], x["label"]) for x in self.pane()["secs"][ki]["controls"]}
        if c["kind"] == "spin":
            d.click(*c["up"][:2])
            how = "spin up"
        else:
            how = self.use(c)
        if how is None:
            if c["kind"] not in ("heading", "other") and not c.get("checked"):
                self.results.append({"kind": self.kind, "element": el, "set": set_label, "tab": tab, "section": sec, "control": c["label"], "type": c["kind"], "skipped": True,
                                     "why": "disabled" if c.get("disabled") else "not pressed", "depth": depth})
            return
        time.sleep(0.12)
        after = self.state()
        changed = after["model"] != before["model"] or after["from"] != before["from"]
        drawn = after["pic"] != before["pic"]
        rec = {"kind": self.kind, "element": el, "set": set_label, "tab": tab, "section": sec, "control": c["label"], "type": c["kind"], "how": how,
               "changed": changed, "drawn": drawn, "depth": depth}
        self.results.append(rec)
        if changed:
            self.models.append({"label": f"{el} / {sec} / {c['label']} ({how})", "model": json.loads(after["model"]), "base": json.loads(before["model"]), "drawn": drawn})
            rec["file_index"] = len(self.models) - 1
            # Controls this choice revealed (Farbverlauf's stops, Musterfüllung's tiles …) are used before it is undone.
            if depth < 2 and c["kind"] in ("radio", "select", "picture", "check"):
                p = self.pane()
                if p and ki < len(p["secs"]):
                    fresh = [(k, x) for k, x in enumerate(p["secs"][ki]["controls"]) if (x["kind"], x["label"]) not in shown]
                    for k, _x in fresh:
                        q = self.pane()
                        if not q or ki >= len(q["secs"]) or k >= len(q["secs"][ki]["controls"]):
                            break
                        self.use_all(el, set_label, tab, ki, q["secs"][ki]["controls"][k], depth + 1)
            self.undo()
            back = self.state()
            rec["undo_model"] = back["model"] == before["model"] and back["from"] == before["from"]
            rec["undo_drawn"] = back["pic"] == before["pic"]
            rec["pane_after_undo"] = self.pane() is not None and back["el"] == before["el"]
            if self.pane() is None:
                # The pane closed: open it again on the same element.
                d.click(*d.js("(() => { const b = document.querySelectorAll('.chart-obj')[document.querySelectorAll('.chart-obj').length - 1].getBoundingClientRect(); return [b.left + 12, b.top + 12]; })()"))
                d.key("1", ctrl=True)
                self.choose(el)
        rec["ok"] = bool(changed and (drawn or c["label"] in INVISIBLE) and rec.get("undo_model") and rec.get("undo_drawn"))
        if not changed:
            rec["why"] = "no change"
        elif not drawn and c["label"] not in INVISIBLE:
            rec["why"] = "not drawn"
        elif not (rec.get("undo_model") and rec.get("undo_drawn")):
            rec["why"] = "undo"


def visible_expected(rec: dict) -> bool:
    return True


def number_for(c: dict) -> str:
    label = c["label"]
    fixed = {"Minimum": "-3", "Maximum": "12", "Haupt": "2", "Hilfs": "0,5", "Basis": "5", "Achsenwert": "2", "Kategorienummer": "2", "Intervalleinheit": "2",
             "Intervall zwischen Teilstrichen": "2", "Abstand von Achse": "400", "Breite": "3", "Größe": "9", "Höhe": "8", "Winkel": "30", "Schriftgrad": "16",
             "Benutzerdefinierter Winkel": "-30", "Wert": "2", "Vorwärts": "1", "Rückwärts": "1", "Schnittpunkt": "2", "Grad": "3", "Zeitraum": "3",
             "Dezimalstellen": "3", "Abstand": "6", "Weichzeichnen": "8", "Tiefe": "10", "Kontur": "2", "Position": "40", "Datum": "2"}
    v = fixed.get(label)
    if v:
        return v
    try:
        cur = float(str(c.get("value", "0")).split(" ")[0].replace(",", ".").rstrip("%°") or 0)
    except ValueError:
        cur = 0
    return str(round(cur + 5, 2)).replace(".", ",")


def to_files(c: CalcSide, kind: str, models: list[dict]) -> Path | None:
    """One chart per change in a workbook RaabeCalc saves (as a student's file would be)."""
    if not models:
        return None
    path = TMP / f"chartpane-{kind}.xlsx"
    c.reset(Case("chartpane", DATA, [], "A1"))
    js = []
    for k, m in enumerate(models):
        col, row = (k % 4) * 9, (k // 4) * 16 + 7
        js.append({"model": m["model"], "from": {"col": col, "row": row, "x": 0, "y": 0}, "to": {"col": col + 8, "row": row + 15, "x": 0, "y": 0}})
    c.js(f"""(() => {{ const list = {json.dumps(js)}; const s = __calc.app.sheet;
      s.charts = list.map((x, i) => ({{ ...x.model, from: x.from, to: x.to, name: 'Diagramm ' + (i + 1) }}));
      __calc.app.charts.refresh(); }})()""")
    c.js(f"__calc.save({json.dumps(str(path))})")
    return path


def reopen(c: CalcSide, path: Path, models: list[dict]) -> list[tuple[bool, str]]:
    c.js(f"__calc.open({json.dumps(str(path))})")
    time.sleep(0.5)
    got = c.js("JSON.parse(JSON.stringify(__calc.app.sheet.charts.map((m) => { const { raw, from, to, name, ...rest } = m; return rest; })))")
    out = []
    for m, g in zip(models, got or []):
        want = dict(m["model"])
        want.pop("name", None)
        a, b = canon(want), canon(g)
        out.append((a == b, "; ".join(diff(a, b))[:400]))
    return out + [(False, "missing")] * (len(models) - len(got or []))


def diff(a, b, path: str = "") -> list[str]:
    if isinstance(a, dict) and isinstance(b, dict):
        out = []
        for k in sorted(set(a) | set(b)):
            out += diff(a.get(k), b.get(k), f"{path}.{k}")
        return out
    if isinstance(a, list) and isinstance(b, list) and len(a) == len(b):
        out = []
        for i, (x, y) in enumerate(zip(a, b)):
            out += diff(x, y, f"{path}[{i}]")
        return out
    return [] if a == b else [f"{path}: {json.dumps(a, ensure_ascii=False)[:80]} -> {json.dumps(b, ensure_ascii=False)[:80]}"]


def canon(x):
    """Model JSON without undefined-equivalents (empty objects, absent vs false defaults stay distinct)."""
    if isinstance(x, dict):
        out = {k: canon(v) for k, v in sorted(x.items()) if v is not None and k != "raw"}
        return {k: v for k, v in out.items() if v != {}}
    if isinstance(x, list):
        return [canon(v) for v in x]
    if isinstance(x, float) and x.is_integer():
        return int(x)
    return x


def excel_check(path: Path, models: list[dict], kind: str) -> list[dict]:
    """Our own invisible Excel opens the file: COM reports the settings, and each chart's picture is compared with the unchanged chart's."""
    from feellab.excel import ExcelTarget
    from excel_chartpane_com import com_checks
    x = ExcelTarget(visible=False)
    out: list[dict] = []
    try:
        try:
            wb = x.app.Workbooks.Open(str(path))
        except Exception as e:  # noqa: BLE001
            # Excel refused the file: one record that says so (the sweep goes on with the next chart type).
            return [{"index": -1, "label": "(file)", "error": f"Excel did not open the file: {str(e)[:120]}"}]
        ws = wb.Worksheets(1)
        pics = TMP / f"pics-{kind}"
        pics.mkdir(parents=True, exist_ok=True)
        # The unchanged chart: Excel draws it from the first model's base.
        for k, m in enumerate(models, start=1):
            rec: dict = {"index": k - 1, "label": m["label"]}
            try:
                ch = ws.ChartObjects(k).Chart
                checks = com_checks(ch, m["base"], m["model"])
                rec["com"] = checks
                f = pics / f"{k:03d}.png"
                ch.Export(str(f))
                rec["pic"] = h(f.read_bytes().hex())
            except Exception as e:  # noqa: BLE001
                rec["error"] = str(e)[:120]
            out.append(rec)
        wb.Close(False)
    finally:
        x.close()
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--kinds", default=",".join(KINDS))
    ap.add_argument("--element", default=None)
    ap.add_argument("--no-excel", action="store_true")
    ap.add_argument("--port", type=int, default=PORT)
    ap.add_argument("--out", default="2026-11-09-calc-chartpane.json")
    a = ap.parse_args()
    TMP.mkdir(parents=True, exist_ok=True)
    report: dict = {"meta": {"date": datetime.date.today().isoformat(), "exe": "RAABE_CALC_EXE" if os.environ.get("RAABE_CALC_EXE") else "release", "kinds": a.kinds}, "kinds": {}}
    c = calc_side(a.port)
    d = Driver(c)
    out = HERE / "results" / a.out
    try:
        for kind in a.kinds.split(","):
            t0 = time.time()
            s = Sweep(d, kind)
            try:
                s.run(a.element)
            except Exception as e:  # noqa: BLE001
                s.results.append({"kind": kind, "control": "(sweep)", "ok": False, "why": f"error {str(e)[:200]}"})
            path = to_files(c, kind, s.models)
            back = reopen(c, path, s.models) if path else []
            for r in s.results:
                if "file_index" in r:
                    ok, why = back[r["file_index"]] if r["file_index"] < len(back) else (False, "missing")
                    r["reopened"] = ok
                    if not ok:
                        r["reopen_diff"] = why
            report["kinds"][kind] = {"results": s.results, "file": path.name if path else None, "seconds": round(time.time() - t0)}
            if path and not a.no_excel:
                report["kinds"][kind]["excel"] = excel_check(path, s.models, kind)
            print(kind, summary(s.results, report["kinds"][kind].get("excel")), flush=True)
            # Written after every chart type, so a stopped run keeps what it measured.
            out.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    finally:
        c.close()
    out.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("written", out)


def summary(results: list[dict], excel: list[dict] | None) -> str:
    used = [r for r in results if not r.get("skipped") and r.get("control") not in ("(sweep)", "(element list)")]
    changed = sum(1 for r in used if r.get("changed"))
    drawn = sum(1 for r in used if r.get("drawn"))
    undo = sum(1 for r in used if r.get("undo_model") and r.get("undo_drawn"))
    reopened = sum(1 for r in used if r.get("reopened"))
    skipped = sum(1 for r in results if r.get("skipped"))
    s = f"controls used {len(used)}, changed {changed}, drawn {drawn}, undone {undo}, reopened {reopened}, skipped {skipped}"
    if excel:
        com = [c for e in excel for c in e.get("com", [])]
        s += f"; Excel COM {sum(1 for c in com if c['ok'])}/{len(com)}, errors {sum(1 for e in excel if 'error' in e)}"
    return s


if __name__ == "__main__":
    main()
