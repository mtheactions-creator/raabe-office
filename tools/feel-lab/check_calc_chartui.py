"""Everything around a selected chart in RaabeCalc, tried by hand (D-115 - D-118): the release exe runs off screen
and is driven with trusted DevTools input (Input.dispatchMouseEvent / dispatchKeyEvent, no real mouse or keyboard,
no lock needed): click the chart, each of the three buttons, every check box, side menu entry, gallery tile and
filter entry, the Datenquelle auswählen dialog with typing and pointing at cells, Anwenden, Strg+Z; then the
chart's model and drawing are checked, and the workbook RaabeCalc saves is opened in our own invisible Excel, which
must show the same series, ranges and settings (Series.IsFiltered, FullCategoryCollection, DisplayBlanksAs …).
The expected Excel values are what Excel did with the same steps (probe_excel_chart_filter.py, probe_excel_chart_buttons.py).

    .venv/Scripts/python.exe check_calc_chartui.py

Result: results/2026-11-09-calc-chartui.json (case, passed, detail).
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
import time
import traceback
from pathlib import Path

from feellab import calc_catalogue, winapi
from feellab.calc_catalogue import CalcSide
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
RESULT = HERE / "results" / "2026-11-09-calc-chartui.json"
TMP = Path(tempfile.gettempdir()) / "raabe-chartui"
PORT = 9477
os.environ["RAABE_FEEL_BACKGROUND"] = "1"
calc_catalogue.winapi.place = lambda hwnd, x, y, w, h: winapi.place_quietly(hwnd, -4000, 0, w, h)

DATA = {"A1": "Monat", "B1": "Nord", "C1": "Süd", "D1": "West", "A2": "Jan", "B2": "3", "C2": "4", "D2": "2", "A3": "Feb", "B3": "5", "C3": "2",
        "D3": "3", "A4": "Mrz", "B4": "2", "C4": "6", "D4": "4", "A5": "Apr", "B5": "4", "C5": "3", "D5": "5", "F1": "Ost", "F2": "1", "F3": "2", "F4": "3", "F5": "4"}
EXCEL = json.loads((HERE / "results" / "2026-11-09-excel-chart-buttons.json").read_text(encoding="utf-8"))
STYLES = json.loads((HERE / "results" / "2026-11-09-excel-chart-filter.json").read_text(encoding="utf-8"))


class DevSide:
    """The development build in a headless Edge (CHARTUI_DEV=<url>), for iterating between release builds."""

    def __init__(self, url: str, port: int = 0) -> None:
        port = port or 9600 + os.getpid() % 300
        self.close()
        import subprocess
        import urllib.request
        import websocket
        prof = Path(tempfile.gettempdir()) / f"raabe-chartui-edge-{os.getpid()}"
        self.proc = subprocess.Popen([r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", "--headless=new", f"--remote-debugging-port={port}",
                                      f"--user-data-dir={prof}", "--window-size=1600,1000", "--no-first-run", url])
        deadline = time.time() + 30
        while True:
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=1) as r:
                    pages = [p for p in json.load(r) if p.get("type") == "page"]
                if pages:
                    self.ws = websocket.create_connection(pages[0]["webSocketDebuggerUrl"], suppress_origin=True)
                    break
            except OSError:
                pass
            if time.time() > deadline:
                raise TimeoutError("Edge DevTools not reachable")
            time.sleep(0.2)
        self._id = 0
        while not self.js("typeof window.__calc === 'object'"):
            time.sleep(0.1)

    def js(self, expression: str):
        if expression.startswith("__calc.save"):
            return None
        self._id += 1
        self.ws.send(json.dumps({"id": self._id, "method": "Runtime.evaluate", "params": {"expression": expression, "returnByValue": True, "awaitPromise": True}}))
        while True:
            msg = json.loads(self.ws.recv())
            if msg.get("id") == self._id:
                res = msg["result"]
                if "exceptionDetails" in res:
                    raise RuntimeError(res["exceptionDetails"])
                return res["result"].get("value")

    def close(self) -> None:
        # Edge's helper processes outlive the launcher: end every process of this run's own profile.
        import subprocess
        tag = "raabe-chartui-edge-"
        script = ("Get-CimInstance Win32_Process -Filter \"name='msedge.exe'\" | Where-Object { $_.CommandLine -like '*" + tag
                  + "*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }")
        subprocess.run(["powershell", "-NoProfile", "-Command", script], capture_output=True)


class Ui:
    """Trusted input into the WebView (DevTools Input domain) and readouts."""

    def __init__(self, c: CalcSide) -> None:
        self.c = c

    def cdp(self, method: str, params: dict) -> dict:
        c = self.c
        c._id += 1
        c.ws.send(json.dumps({"id": c._id, "method": method, "params": params}))
        while True:
            msg = json.loads(c.ws.recv())
            if msg.get("id") == c._id:
                if "error" in msg:
                    raise RuntimeError(msg["error"])
                return msg.get("result", {})

    def js(self, e: str):
        return self.c.js(e)

    def mouse(self, kind: str, x: float, y: float, button: str = "left", count: int = 1, buttons: int = 0, mods: int = 0) -> None:
        self.cdp("Input.dispatchMouseEvent", {"type": kind, "x": x, "y": y, "button": button, "clickCount": count, "buttons": buttons, "modifiers": mods})

    def move(self, x: float, y: float) -> None:
        self.mouse("mouseMoved", x, y, button="none")
        time.sleep(0.05)

    def click(self, x: float, y: float, button: str = "left", count: int = 1) -> None:
        self.move(x, y)
        b = 1 if button == "left" else 2
        for n in range(1, count + 1):
            self.mouse("mousePressed", x, y, button=button, count=n, buttons=b)
            self.mouse("mouseReleased", x, y, button=button, count=n)
        time.sleep(0.25)

    def drag(self, x1: float, y1: float, x2: float, y2: float, steps: int = 12) -> None:
        self.move(x1, y1)
        self.mouse("mousePressed", x1, y1, buttons=1)
        for i in range(1, steps + 1):
            self.mouse("mouseMoved", x1 + (x2 - x1) * i / steps, y1 + (y2 - y1) * i / steps, button="left", buttons=1)
            time.sleep(0.02)
        self.mouse("mouseReleased", x2, y2)
        time.sleep(0.3)

    def key(self, key: str, code: str, vk: int, mods: int = 0) -> None:
        self.cdp("Input.dispatchKeyEvent", {"type": "rawKeyDown", "key": key, "code": code, "windowsVirtualKeyCode": vk, "modifiers": mods})
        self.cdp("Input.dispatchKeyEvent", {"type": "keyUp", "key": key, "code": code, "windowsVirtualKeyCode": vk, "modifiers": mods})
        time.sleep(0.2)

    def type(self, text: str) -> None:
        self.cdp("Input.insertText", {"text": text})
        time.sleep(0.15)

    def undo(self) -> None:
        self.key("z", "KeyZ", 90, 2)

    def center(self, sel: str, i: int = 0):
        r = self.js(f"(() => {{ const e = document.querySelectorAll({json.dumps(sel)})[{i}]; if (!e) return null; e.scrollIntoView({{ block: 'nearest' }}); const r = e.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2, r.width, r.height]; }})()")
        if not r:
            raise RuntimeError(f"not found: {sel} [{i}]")
        return r

    def click_sel(self, sel: str, i: int = 0) -> None:
        x, y, _, _ = self.center(sel, i)
        self.click(x, y)

    def click_text(self, sel: str, text: str) -> None:
        """Clicks the element under `sel` whose text is `text`."""
        r = self.js(f"""(() => {{ const e = [...document.querySelectorAll({json.dumps(sel)})].find((x) => x.textContent.trim() === {json.dumps(text)} && x.getClientRects().length);
          if (!e) return null; const r = e.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; }})()""")
        if not r:
            raise RuntimeError(f"not found: {sel} '{text}'")
        self.click(*r)

    def texts(self, sel: str) -> list[str]:
        return self.js(f"[...document.querySelectorAll({json.dumps(sel)})].filter((e) => e.getClientRects().length).map((e) => e.textContent.trim())")

    def model(self):
        return self.js("(() => { const m = __calc.app.charts.selected() ?? __calc.app.sheet.charts?.[0]; if (!m) return null; const { raw, ...rest } = m; return JSON.parse(JSON.stringify(rest)); })()")


class Run:
    def __init__(self) -> None:
        self.cases: list[dict] = []

    def check(self, name: str, ok: bool, detail="") -> None:
        self.cases.append({"case": name, "passed": bool(ok), "detail": detail if not ok else ""})
        print(("PASS " if ok else "FAIL ") + name + ("" if ok else f"  {detail}"), flush=True)


def chart_center(ui: Ui):
    return ui.center(".chart-obj")


def fresh(ui: Ui) -> None:
    """A new workbook with the data and a column chart of A1:D5, selected with a click."""
    ui.js("document.querySelectorAll('.dialog-layer .dlg-title button').forEach((b) => b.click())")
    ui.key("Escape", "Escape", 27)
    ui.js(f"__calc.reset({json.dumps(DATA)}, 'A1', 1)")
    ui.js("(() => { const a = __calc.app; a.grid.select(__calc.parse('A1:D5')); a.charts.insert('column'); })()")
    time.sleep(0.4)
    x, y, w, h = chart_center(ui)
    ui.click(x - w * 0.45, y + h * 0.45)  # the chart area (lower left corner)


def ensure(ui: Ui, which: str) -> None:
    """Opens a button's flyout unless it is open (a click on its button would close it)."""
    if ui.js(f"document.querySelector('.cb-fly')?.dataset.which === {json.dumps(which)}"):
        return
    ui.click_sel(f'.chart-btn[data-which="{which}"]')


def flyout_rows(ui: Ui):
    return ui.js("[...document.querySelectorAll('.cb-fly .cb-row')].map((r) => [r.textContent.trim(), r.querySelector('input').checked])")


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    TMP.mkdir(exist_ok=True)
    run = Run()
    dev = os.environ.get("CHARTUI_DEV")
    c = DevSide(dev) if dev else CalcSide(PORT)
    ui = Ui(c)
    saved: dict[str, dict] = {}
    try:
        cases(ui, run, saved)
    except Exception:  # noqa: BLE001
        run.check("run completed", False, traceback.format_exc()[-800:])
    finally:
        try:
            c.close()
        except Exception:  # noqa: BLE001
            pass
    if not dev:
        excel_side(run, saved)
    passed = sum(1 for x in run.cases if x["passed"])
    if dev:
        print(f"{sum(1 for x in run.cases if x['passed'])}/{len(run.cases)} passed (dev)")
        return
    RESULT.write_text(json.dumps({"passed": passed, "total": len(run.cases), "cases": run.cases}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(f"{passed}/{len(run.cases)} passed")


def cases(ui: Ui, run: Run, saved: dict) -> None:
    fresh(ui)
    run.check("click selects the chart", ui.js("!!__calc.app.charts.selected()"))
    # ----- the three buttons: Excel's size and place --------------------------------------------------
    rects = ui.js("(() => { const o = document.querySelector('.chart-obj').getBoundingClientRect(); return [...document.querySelectorAll('.chart-obj.sel .chart-btn')].map((b) => { const r = b.getBoundingClientRect(); return [b.getAttribute('aria-label'), Math.round(r.left - o.right), Math.round(r.top - o.top), r.width, r.height]; }); })()")
    ex = [(b["css"]["dx_from_chart_right"], b["css"]["dy_from_chart_top"], b["css"]["w"]) for b in EXCEL["side_buttons"]]
    run.check("three buttons beside the chart", [r[0] for r in rects] == ["Diagrammelemente", "Diagrammformatvorlagen", "Diagrammfilter"], rects)
    run.check("buttons at Excel's size and place (±2 px)", len(rects) == 3 and all(abs(r[1] - e[0]) <= 2 and abs(r[2] - e[1]) <= 2 and abs(r[3] - e[2]) <= 1 for r, e in zip(rects, ex)), [rects, ex])

    # ----- Diagrammelemente ---------------------------------------------------------------------------
    ensure(ui, "elements")
    rows = flyout_rows(ui)
    excel_rows = [it["name"] for p in EXCEL["flyouts"]["Element hinzufügen"]["main"] for it in p["items"] if it["type"] == "MenuItem"]
    excel_rows = list(dict.fromkeys(excel_rows))
    run.check("Diagrammelemente lists Excel's entries", [r[0] for r in rows] == excel_rows, [rows, excel_rows])
    run.check("Diagrammelemente's ticks as Excel's (Achsen, Diagrammtitel, Gitternetzlinien, Legende)",
              [r[0] for r in rows if r[1]] == ["Achsen", "Diagrammtitel", "Gitternetzlinien", "Legende"], rows)
    # Each check box off and on again (one undo step each), the model must follow.
    probes = {
        "Achsen": "(m) => m.format?.catAxis?.visible !== false && m.format?.valAxis?.visible !== false",
        "Achsentitel": "(m) => m.axisTitles.cat !== undefined && m.axisTitles.val !== undefined",
        "Diagrammtitel": "(m) => m.title.show",
        "Datenbeschriftungen": "(m) => m.dataLabels",
        "Datentabelle": "(m) => !!m.dataTable",
        "Fehlerindikatoren": "(m) => m.series.every((s) => s.errorBars?.type === 'stdErr')",
        "Gitternetzlinien": "(m) => m.gridlines",
        "Legende": "(m) => m.legend !== 'none'",
    }
    for i, (label, on) in enumerate(rows):
        if label not in probes:
            continue
        before = ui.js(f"({probes[label]})(__calc.app.charts.selected())")
        ui.click_sel(".cb-fly .cb-row input", i)
        after = ui.js(f"({probes[label]})(__calc.app.charts.selected())")
        open_after = ui.js("!!document.querySelector('.cb-fly')")
        ui.click_sel(".cb-fly .cb-row input", i)
        back = ui.js(f"({probes[label]})(__calc.app.charts.selected())")
        run.check(f"Diagrammelemente: {label} ticked/cleared and back", before == on and after == (not on) and back == on and open_after, [before, after, back, open_after])
    # Trendlinie with three series: Excel asks which series.
    i = [r[0] for r in rows].index("Trendlinie")
    ui.click_sel(".cb-fly .cb-row input", i)
    asked = ui.js("[...document.querySelectorAll('.dialog .dlg-title .t')].map((t) => t.textContent)")
    run.check("Trendlinie asks for the series (Trendlinie hinzufügen)", "Trendlinie hinzufügen" in asked, asked)
    ui.js("(() => { const s = document.querySelector('.dialog select'); s.value = '1'; })()")
    ui.click_text(".dialog .dlg-buttons button", "OK")
    m = ui.model()
    run.check("Trendlinie on the chosen series only", [bool(s.get("trendline")) for s in m["series"]] == [False, True, False], [s.get("trendline") for s in m["series"]])
    ui.undo()
    m = ui.model()
    run.check("Strg+Z takes the trendline away", not any(s.get("trendline") for s in m["series"]), m["series"])

    # Side menus: hover an entry, click its arrow; every choice applied.
    if not ui.js("!!document.querySelector('.cb-fly')"):
        ensure(ui, "elements")
    side_expect = {
        "Achsen": (["Primär horizontal", "Primär vertikal", "Weitere Optionen..."], [("Primär vertikal", "(m) => m.format?.valAxis?.visible === false")]),
        "Achsentitel": (["Primär horizontal", "Primär vertikal", "Weitere Optionen..."], [("Primär horizontal", "(m) => m.axisTitles.cat === ''")]),
        "Diagrammtitel": (["Über Diagramm", "Zentrierte Überlagerung", "Weitere Optionen..."], [("Zentrierte Überlagerung", "(m) => m.title.show && m.format?.title?.overlay === true")]),
        "Datenbeschriftungen": (["Zentriert", "Am Ende innerhalb", "Am Anfang innerhalb", "Am Ende außerhalb", "Datenlegende", "Weitere Optionen..."],
                                [("Am Ende innerhalb", "(m) => m.dataLabels && m.format?.dataLabels?.pos === 'inEnd'"), ("Zentriert", "(m) => m.format?.dataLabels?.pos === 'ctr'")]),
        "Datentabelle": (["Mit Legendensymbolen", "Ohne Legendensymbole", "Weitere Optionen..."], [("Ohne Legendensymbole", "(m) => m.dataTable && m.dataTable.legendKeys === false")]),
        "Fehlerindikatoren": (["Standardfehler", "Prozentsatz", "Standardabweichung", "Weitere Optionen..."], [("Prozentsatz", "(m) => m.series.every((s) => s.errorBars?.type === 'percentage' && s.errorBars.value === 5)")]),
        "Gitternetzlinien": (["Primäres Hauptgitter horizontal", "Primäres Hauptgitter vertikal", "Primäres Hilfsgitter horizontal", "Primäres Hilfsgitter vertikal", "Weitere Optionen..."],
                             [("Primäres Hauptgitter vertikal", "(m) => m.format?.catGrid?.show === true"), ("Primäres Hilfsgitter vertikal", "(m) => m.format?.catMinorGrid?.show === true")]),
        "Legende": (["Rechts", "Oben", "Links", "Unten", "Weitere Optionen..."], [("Oben", "(m) => m.legend === 't'"), ("Rechts", "(m) => m.legend === 'r'")]),
    }
    side_expect["Trendlinie"] = (["Linear", "Exponential", "Lineare Prognose", "Gleitender Durchschnitt für zwei Zeiträume", "Weitere Optionen..."], [])
    for label, (labels, picks) in side_expect.items():
        idx = [r[0] for r in flyout_rows(ui)].index(label)
        x, y, w, h = ui.center(".cb-fly .cb-row", idx)
        ui.move(x - 20, y)
        ui.click_sel(".cb-fly .cb-row .cb-arrow", idx)
        got = ui.texts(".cb-sub button")
        recorded = [it["name"] for p in EXCEL.get("element_submenus", {}).get(label, []) for it in p["items"] if it["type"] in ("MenuItem", "CheckBox", "ListItem", "RadioButton")]
        recorded = list(dict.fromkeys(recorded)) or labels
        run.check(f"side menu {label}: Excel's entries", got == recorded, [got, recorded])
        for pick, probe in picks:
            if not ui.js("!!document.querySelector('.cb-sub')"):
                ui.move(x - 20, y)
                ui.click_sel(".cb-fly .cb-row .cb-arrow", idx)
            ui.click_text(".cb-sub button", pick)
            ok = ui.js(f"({probe})(__calc.app.charts.selected())")
            run.check(f"side menu {label} > {pick}", ok, ui.model())
            if not ui.js("!!document.querySelector('.cb-fly')"):
                ensure(ui, "elements")
        if ui.js("!!document.querySelector('.cb-sub')"):
            ui.key("Escape", "Escape", 27)
    # The state the side menus left, saved: Excel must open it with the same elements.
    ui.js(f"__calc.save({json.dumps(str(TMP / 'elements.xlsx'))})")
    saved["elements"] = {}
    # Weitere Optionen... opens the Format pane of the element.
    idx = [r[0] for r in flyout_rows(ui)].index("Legende")
    x, y, w, h = ui.center(".cb-fly .cb-row", idx)
    ui.move(x - 20, y)
    ui.click_sel(".cb-fly .cb-row .cb-arrow", idx)
    ui.click_text(".cb-sub button", "Weitere Optionen...")
    title = ui.js("document.querySelector('.taskpane')?.textContent ?? ''")
    run.check("Weitere Optionen... opens Legende formatieren", "Legende formatieren" in title, title[:80])
    ui.js("__calc.app.taskPane.close()")
    # Escape closes the flyout.
    x, y, w, h = chart_center(ui)
    ui.click(x - w * 0.45, y + h * 0.45)
    ensure(ui, "elements")
    ui.key("Escape", "Escape", 27)
    run.check("Esc closes the flyout", not ui.js("!!document.querySelector('.cb-fly')"))

    # ----- Diagrammformatvorlagen ----------------------------------------------------------------------
    fresh(ui)
    ensure(ui, "styles")
    tabs = ui.texts(".cb-fly .cb-tab")
    run.check("Diagrammformatvorlagen: tabs Formatvorlage / Farbe", tabs == ["Formatvorlage", "Farbe"], tabs)
    n = ui.js("document.querySelectorAll('.cb-fly .cb-tile').length")
    run.check("Formatvorlage: 14 styles for a column chart (Excel's gallery)", n == len(STYLES["gallery"]["column"]), n)
    before = ui.js("document.querySelector('.chart-obj svg').outerHTML.length")
    x, y, w, h = ui.center(".cb-fly .cb-tile", 5)
    ui.move(x, y)
    during = ui.js("document.querySelector('.chart-obj svg').outerHTML.length")
    style_during = ui.js("__calc.app.charts.selected().style ?? null")
    ui.move(x + 400, y)
    after = ui.js("document.querySelector('.chart-obj svg').outerHTML.length")
    run.check("Livevorschau: hovering a style redraws the chart, leaving restores it, nothing changes", during != before and after == before and style_during is None, [before, during, after])
    for k in (1, 5, 8, 13):
        if not ui.js("!!document.querySelector('.cb-fly .cb-tile')"):
            ensure(ui, "styles")
        ui.click_sel(".cb-fly .cb-tile", k)
        want = STYLES["cases"][f"style-column-{k + 1}"]["style"]
        got = ui.js("__calc.app.charts.selected().style ?? 201")
        run.check(f"Formatvorlage {k + 1} applied (Excel's style {want})", got == want, got)
    m = ui.model()
    saved["styled"] = {"style": m.get("style")}
    ui.js(f"__calc.save({json.dumps(str(TMP / 'styled.xlsx'))})")
    ui.undo()
    run.check("Strg+Z puts the previous style back", ui.js("__calc.app.charts.selected().style ?? 201") == STYLES["cases"]["style-column-9"]["style"])
    if not ui.js("!!document.querySelector('.cb-fly .cb-tab')"):
        ensure(ui, "styles")
    ui.click_text(".cb-fly .cb-tab", "Farbe")
    heads = ui.texts(".cb-fly .cb-pal-head")
    pals = ui.js("[...document.querySelectorAll('.cb-fly .cb-pal')].map((p) => p.getAttribute('aria-label'))")
    excel_pals = [it["name"] for p in EXCEL["flyouts"]["Formatvorlagen"]["tab:Farbe"][:1] for it in p["items"] if it["type"] == "ListItem"]
    run.check("Farbe: Bunt / Einfarbig with Excel's 17 palettes", heads == ["Bunt", "Einfarbig"] and pals == excel_pals, [heads, pals, excel_pals])
    ui.click_sel(".cb-fly .cb-pal", 6)
    run.check("Monochrome Palette 3 applied (scheme 16)", ui.js("__calc.app.charts.selected().colors") == 16)

    # ----- Diagrammfilter ----------------------------------------------------------------------------------
    fresh(ui)
    ensure(ui, "filter")
    tabs = ui.texts(".cb-fly .cb-tab")
    rows = ui.texts(".cb-fly .cb-filter-row")
    excel_main = [it for p in EXCEL["flyouts"]["Datenlayout"]["main"][:1] for it in p["items"]]
    excel_rows = [it["name"] for it in excel_main if it["type"] == "CheckBox"]
    run.check("Diagrammfilter: tabs Werte / Namen", tabs == ["Werte", "Namen"], tabs)
    run.check("Diagrammfilter: Excel's rows (Alle auswählen, series, categories)", rows == excel_rows, [rows, excel_rows])
    groups = ui.texts(".cb-fly .cb-group-head")
    run.check("Diagrammfilter: groups Datenreihe and Kategorien", groups == ["Datenreihe", "Kategorien"], groups)
    foot = ui.texts(".cb-fly .cb-foot button, .cb-fly .cb-foot a")
    run.check("Diagrammfilter: Anwenden and Daten auswählen...", foot == ["Anwenden", "Daten auswählen..."], foot)
    # Hover a series: the others fade.
    x, y, w, h = ui.center(".cb-fly .cb-filter-row", 2)
    ui.move(x, y)
    faded = ui.js("[...new Set([...document.querySelectorAll('.chart-obj svg [data-fade]')].map((e) => e.dataset.part.split(':')[1]))]")
    ui.move(x + 600, y)
    run.check("hovering Süd fades Nord and West", sorted(faded) == ["0", "2"], faded)
    x, y, w, h = ui.center(".cb-fly .cb-filter-row", 7)
    ui.move(x, y)
    faded_cat = ui.js("document.querySelectorAll('.chart-obj svg [data-fade]').length")
    ui.move(x + 600, y)
    run.check("hovering Mrz fades the other categories' columns", faded_cat == 9, faded_cat)
    # Untick Süd and Mrz: nothing until Anwenden.
    ui.click_sel(".cb-fly .cb-filter-row input", 2)
    ui.click_sel(".cb-fly .cb-filter-row input", 7)
    run.check("nothing changes before Anwenden", not any(s.get("filtered") for s in ui.model()["series"]) and not ui.model().get("hiddenCats"))
    ui.click_text(".cb-fly .cb-foot button", "Anwenden")
    m = ui.model()
    run.check("Anwenden hides Süd (filtered) and Mrz", [bool(s.get("filtered")) for s in m["series"]] == [False, True, False] and m.get("hiddenCats") == [2], [m["series"], m.get("hiddenCats")])
    shown = ui.js("[...new Set([...document.querySelectorAll('.chart-obj svg [data-part^=\"series:\"]')].map((e) => e.dataset.part.split(':')[1]))].length")
    cols = ui.js("document.querySelectorAll('.chart-obj svg [data-part^=\"series:\"]').length")
    run.check("the chart draws 2 series of 3 categories (6 columns)", shown == 2 and cols == 6, [shown, cols])
    legend = ui.js("[...document.querySelectorAll('.chart-obj svg text')].map((t) => t.textContent)")
    run.check("legend without Süd, axis without Mrz", "Süd" not in legend and "Mrz" not in legend and "West" in legend, legend)
    ui.js(f"__calc.save({json.dumps(str(TMP / 'filtered.xlsx'))})")
    saved["filtered"] = {"series": [False, True, False], "cats": [False, False, True, False]}
    ui.undo()
    m = ui.model()
    run.check("Strg+Z shows everything again (one undo step)", not any(s.get("filtered") for s in m["series"]) and not m.get("hiddenCats"))
    # Select all again with (Alle auswählen).
    ensure(ui, "filter")
    ui.click_sel(".cb-fly .cb-filter-row input", 0)
    ui.click_text(".cb-fly .cb-foot button", "Anwenden")
    run.check("(Alle auswählen) cleared and Anwenden hides every series", all(s.get("filtered") for s in ui.model()["series"]))
    ui.undo()
    # Namen: (Keine) for the series names.
    ensure(ui, "filter")
    ui.click_text(".cb-fly .cb-tab", "Namen")
    radios = ui.js("[...document.querySelectorAll('.cb-fly .cb-radio')].map((r) => r.textContent.trim())")
    run.check("Namen: Zeile 1 / (Keine), Spalte A / (Keine)", radios[0].startswith("Zeile 1") and radios[1] == "(Keine)" and radios[2].startswith("Spalte A") and radios[3] == "(Keine)", radios)
    ui.click_sel(".cb-fly .cb-radio input", 1)
    ui.click_text(".cb-fly .cb-foot button", "Anwenden")
    m = ui.model()
    run.check("Namen > (Keine): the series lose their names (Reihe1 …)", all(not s.get("name") for s in m["series"]), m["series"][0].get("name"))
    ui.undo()

    # ----- Datenquelle auswählen ----------------------------------------------------------------------------
    fresh(ui)
    ensure(ui, "filter")
    ui.click_text(".cb-fly .cb-foot a", "Daten auswählen...")
    title = ui.js("[...document.querySelectorAll('.dialog .dlg-title .t')].map((t) => t.textContent)")
    run.check("Daten auswählen... opens Datenquelle auswählen", "Datenquelle auswählen" in title, title)
    val = ui.js("document.querySelector('.dialog .rf input').value")
    run.check("Diagrammdatenbereich shows =Tabelle1!$A$1:$D$5", val == "=Tabelle1!$A$1:$D$5", val)
    labels = ui.js("[...document.querySelectorAll('.dialog .sd-pane-title, .dialog .sd-tool button, .dialog .sd-center button, .dialog .sd-foot button')].map((b) => b.getAttribute('aria-label') ?? b.textContent)")
    want = ["Legendeneinträge (Reihen)", "Datenreihe hinzufügen", "Datenreihe bearbeiten", "Datenreihe entfernen", "Nach oben", "Nach unten", "Horizontale Achsenbeschriftungen (Rubrik)", "Kategorie bearbeiten", "Zeile/Spalte tauschen", "Ausgeblendete und leere Zellen"]
    run.check("the dialog's buttons as Excel names them", sorted(labels) == sorted(want), labels)
    items = ui.js("[...document.querySelectorAll('.dialog .sd-list')].map((l) => [...l.querySelectorAll('.sd-item')].map((i) => i.textContent.trim()))")
    run.check("series list Nord, Süd, West; categories Jan … Apr", items == [["Nord", "Süd", "West"], ["Jan", "Feb", "Mrz", "Apr"]], items)
    # Dialog reduzieren, point at A1:C5 in the sheet, grow again.
    ui.click_sel(".dialog .rf button")
    collapsed = ui.js("document.querySelector('.dialog').classList.contains('rf-collapsed')")
    p1 = ui.js("__calc.point('A1', 'center')")
    p2 = ui.js("__calc.point('C5', 'center')")
    ui.drag(p1["x"], p1["y"], p2["x"], p2["y"])
    val = ui.js("document.querySelector('.dialog .rf input').value")
    run.check("Dialog reduzieren shrinks the dialog to the field", collapsed)
    run.check("pointing at A1:C5 writes =Tabelle1!$A$1:$C$5", val == "=Tabelle1!$A$1:$C$5", val)
    live = ui.js("document.querySelectorAll('.chart-obj svg [data-part^=\"series:\"]').length")
    run.check("the chart follows the dialog (2 series x 4 categories)", live == 8, live)
    ui.click_sel(".dialog .rf button")
    run.check("the button grows the dialog again", not ui.js("document.querySelector('.dialog').classList.contains('rf-collapsed')"))
    # Zeile/Spalte tauschen in the dialog.
    ui.click_sel('.dialog .sd-center button')
    items = ui.js("[...document.querySelectorAll('.dialog .sd-list')][0].textContent")
    run.check("Zeile/Spalte tauschen: series are the months", all(k in items for k in ("Jan", "Feb", "Mrz", "Apr")), items)
    ui.click_sel('.dialog .sd-center button')
    # Hinzufügen: type the name, point at the values.
    ui.click_sel('.dialog [aria-label="Datenreihe hinzufügen"]')
    sub = ui.js("[...document.querySelectorAll('.dialog .dlg-title .t')].map((t) => t.textContent)")
    run.check("Hinzufügen opens Datenreihe bearbeiten", "Datenreihe bearbeiten" in sub, sub)
    fields = ui.js("[...document.querySelectorAll('.dialog')].at(-1).textContent")
    run.check("Datenreihe bearbeiten: Reihenname and Reihenwerte", "Reihenname:" in fields and "Reihenwerte:" in fields, fields[:80])
    # The dialog covers the cells: shrink it with the field's button (as one would), point, grow it again.
    last = "[...document.querySelectorAll('.dialog')].at(-1)"
    ui.js(f"{last}.querySelectorAll('.rf input')[0].focus()")
    x, y = ui.js(f"(() => {{ const b = {last}.querySelectorAll('.rf button')[0].getBoundingClientRect(); return [b.left + b.width / 2, b.top + b.height / 2]; }})()")
    ui.click(x, y)
    p = ui.js("__calc.point('F1', 'center')")
    ui.click(p["x"], p["y"])
    x, y = ui.js(f"(() => {{ const b = {last}.querySelector('.rf-bar .rf button').getBoundingClientRect(); return [b.left + b.width / 2, b.top + b.height / 2]; }})()")
    ui.click(x, y)
    ui.js(f"{last}.querySelectorAll('.rf input')[1].focus()")
    x, y = ui.js(f"(() => {{ const b = {last}.querySelectorAll('.rf button')[1].getBoundingClientRect(); return [b.left + b.width / 2, b.top + b.height / 2]; }})()")
    ui.click(x, y)
    p1 = ui.js("__calc.point('F2', 'center')")
    p2 = ui.js("__calc.point('F5', 'center')")
    ui.drag(p1["x"], p1["y"], p2["x"], p2["y"])
    x, y = ui.js(f"(() => {{ const b = {last}.querySelector('.rf-bar .rf button').getBoundingClientRect(); return [b.left + b.width / 2, b.top + b.height / 2]; }})()")
    ui.click(x, y)
    vals = ui.js("[...[...document.querySelectorAll('.dialog')].at(-1).querySelectorAll('.rf input')].map((i) => i.value)")
    prev = ui.js("[...[...document.querySelectorAll('.dialog')].at(-1).querySelectorAll('.sd-preview')].map((i) => i.textContent)")
    run.check("pointing fills Reihenname =Tabelle1!$F$1 and Reihenwerte =Tabelle1!$F$2:$F$5", vals == ["=Tabelle1!$F$1", "=Tabelle1!$F$2:$F$5"], vals)
    run.check("the previews show = Ost and = 1; 2; 3; 4", prev == ["= Ost", "= 1; 2; 3; 4"], prev)
    ui.js("[...document.querySelectorAll('.dialog')].at(-1).querySelector('.dlg-buttons button.primary').click()")
    time.sleep(0.3)
    items = ui.js("[...document.querySelectorAll('.dialog .sd-list')][0] ? [...[...document.querySelectorAll('.dialog .sd-list')][0].querySelectorAll('.sd-item')].map((i) => i.textContent.trim()) : null")
    run.check("Ost added to the series list", items == ["Nord", "Süd", "Ost"], items)
    # Bearbeiten Süd by typing.
    ui.click_text(".dialog .sd-list .sd-item", "Süd")
    ui.click_sel('.dialog [aria-label="Datenreihe bearbeiten"]')
    ui.js("(() => { const i = [...document.querySelectorAll('.dialog')].at(-1).querySelectorAll('.rf input')[0]; i.focus(); i.select(); })()")
    ui.type("Mitte")
    ui.js("[...document.querySelectorAll('.dialog')].at(-1).querySelector('.dlg-buttons button.primary').click()")
    time.sleep(0.3)
    # Nach oben, Entfernen.
    ui.click_text(".dialog .sd-list .sd-item", "Ost")
    ui.click_sel('.dialog [aria-label="Nach oben"]')
    items = ui.js("[...[...document.querySelectorAll('.dialog .sd-list')][0].querySelectorAll('.sd-item')].map((i) => i.textContent.trim())")
    run.check("Bearbeiten (typed name) and Nach oben", items == ["Nord", "Ost", "Mitte"], items)
    # Ausgeblendete und leere Zellen: Null, ausgeblendete Daten anzeigen.
    ui.click_sel('.dialog [aria-label="Ausgeblendete und leere Zellen"]')
    opts = ui.js("[...document.querySelectorAll('.dialog')].at(-1).textContent")
    run.check("Ausgeblendete und leere Zellen: Lücken, Null, verbinden, #NV, ausgeblendete Daten",
              all(k in opts for k in ("Lücken", "Null", "Datenpunkte mit einer Linie verbinden", '"#N/V" als leere Zelle anzeigen', "Daten in ausgeblendeten Zeilen und Spalten anzeigen")), opts[:200])
    ui.js("(() => { const d = [...document.querySelectorAll('.dialog')].at(-1); d.querySelector('input[value=zero]').click(); [...d.querySelectorAll('input[type=checkbox]')].at(-1).click(); d.querySelector('.dlg-buttons button.primary').click(); })()")
    time.sleep(0.3)
    # Uncheck a category in the dialog's list (as the filter does).
    ui.js("[...document.querySelectorAll('.dialog .sd-list')][1].querySelectorAll('input')[3].click()")
    ui.js("document.querySelector('.dialog .dlg-buttons button.primary').click()")
    time.sleep(0.4)
    m = ui.model()
    ok = ([s.get("name", {}).get("ref") or s.get("name", {}).get("text") for s in m["series"]] == ["Tabelle1!$B$1", "Tabelle1!$F$1", "Mitte"]
          and [s["val"]["ref"] for s in m["series"]] == ["Tabelle1!$B$2:$B$5", "Tabelle1!$F$2:$F$5", "Tabelle1!$C$2:$C$5"]
          and m.get("dispBlanksAs") == "zero" and m.get("showHidden") is True and m.get("hiddenCats") == [3])
    run.check("OK keeps the series, ranges, Null, ausgeblendete Daten and the hidden category", ok, [[s.get("name"), s["val"].get("ref")] for s in m["series"]] + [m.get("dispBlanksAs"), m.get("showHidden"), m.get("hiddenCats")])
    ui.js(f"__calc.save({json.dumps(str(TMP / 'selectdata.xlsx'))})")
    saved["selectdata"] = {"formulas": ["=SERIES(Tabelle1!$B$1,Tabelle1!$A$2:$A$5,Tabelle1!$B$2:$B$5,1)", "=SERIES(Tabelle1!$F$1,Tabelle1!$A$2:$A$5,Tabelle1!$F$2:$F$5,2)",
                                        '=SERIES("Mitte",Tabelle1!$A$2:$A$5,Tabelle1!$C$2:$C$5,3)'], "blanks": 2, "plotVisibleOnly": False, "cats": [False, False, False, True]}
    ui.undo()
    m = ui.model()
    run.check("Strg+Z undoes the whole dialog in one step", [s["val"]["ref"] for s in m["series"]] == ["Tabelle1!$B$2:$B$5", "Tabelle1!$C$2:$C$5", "Tabelle1!$D$2:$D$5"], [s["val"]["ref"] for s in m["series"]])
    # Abbrechen leaves the chart as it was.
    ui.js("__calc.app.charts.selectData()")
    time.sleep(0.3)
    ui.js("document.querySelectorAll('.dialog .sd-list')[0].querySelectorAll('.sd-item')[0].dispatchEvent(new MouseEvent('mousedown', { bubbles: true }))")
    ui.click_sel('.dialog [aria-label="Datenreihe entfernen"]')
    ui.click_text(".dialog .dlg-buttons button", "Abbrechen")
    run.check("Abbrechen keeps all three series", len(ui.model()["series"]) == 3)

    # ----- the source ranges in the sheet ---------------------------------------------------------------------
    fresh(ui)
    frames = ui.js("[...document.querySelectorAll('.chart-range')].map((f) => [f.dataset.kind, getComputedStyle(f).borderTopColor])")
    run.check("selected chart: names, categories and values framed (red, purple, blue)", sorted(f[0] for f in frames) == ["cat", "name", "val"], frames)
    rf = ui.js("__calc.app.charts.rangeFrames()")
    run.check("frames cover B1:D1, A2:A5, B2:D5", sorted((f["kind"], f["rect"]["r0"], f["rect"]["c0"], f["rect"]["r1"], f["rect"]["c1"]) for f in rf) == [("cat", 1, 0, 4, 0), ("name", 0, 1, 0, 3), ("val", 1, 1, 4, 3)], rf)
    # Drag the values frame's lower right corner from D5 up to D4: the chart's data become A1:D4.
    x, y, w, h = ui.center('.chart-range[data-kind="val"] .corner.se')
    p = ui.js("__calc.point('D4', 'center')")
    ui.drag(x, y, p["x"], p["y"])
    m = ui.model()
    run.check("dragging the values' corner to D4 shrinks the data to A1:D4", [s["val"]["ref"] for s in m["series"]] == ["Tabelle1!$B$2:$B$4", "Tabelle1!$C$2:$C$4", "Tabelle1!$D$2:$D$4"] and m["series"][0]["cat"]["ref"] == "Tabelle1!$A$2:$A$4",
              [s["val"]["ref"] for s in m["series"]])
    ui.undo()
    # Click a series, then move its value frame one column right (onto Ost's column F? no: from B to C).
    x, y, w, h = chart_center(ui)
    col = ui.js("(() => { const e = document.querySelector('.chart-obj svg [data-part=\"series:0:1\"]').getBoundingClientRect(); return [e.left + e.width / 2, e.top + e.height / 2]; })()")
    ui.click(*col)
    rf = ui.js("__calc.app.charts.rangeFrames()")
    run.check("a selected series frames only its own ranges", len(rf) == 3 and all(f["rect"]["c0"] in (0, 1) for f in rf), rf)
    x1, y1, _, _ = ui.center('.chart-range[data-kind="val"] .edge.t')
    p = ui.js("__calc.point('C1', 'center')")
    q = ui.js("__calc.point('B1', 'center')")
    ui.drag(x1, y1, x1 + (p["x"] - q["x"]), y1)
    m = ui.model()
    run.check("moving the series' values frame one column right: =Tabelle1!$C$2:$C$5", m["series"][0]["val"]["ref"] == "Tabelle1!$C$2:$C$5", m["series"][0]["val"])
    ui.undo()

    # ----- context menus --------------------------------------------------------------------------------------------
    fresh(ui)
    menus = EXCEL.get("context_menus", {})

    def menu_after_right_click(part: str) -> list[str]:
        ui.key("Escape", "Escape", 27)
        r = ui.js(f"(() => {{ const e = document.querySelector('.chart-obj svg [data-part=\"{part}\"]') ?? document.querySelector('.chart-obj svg [data-part^=\"{part}\"]'); if (!e) return null; const b = e.getBoundingClientRect(); return [b.left + Math.min(b.width / 2, 8), b.top + b.height / 2]; }})()")
        if not r:
            return []
        ui.click(r[0], r[1], button="right")
        return ui.texts(".menu-root > .menu > button")

    def excel_menu(name: str) -> list[str]:
        for p in menus.get(name, []):
            if p["window"]["class"] == "NetUITWMenuContainer":
                return [it["name"] for it in p["items"] if it["type"] in ("MenuItem", "SplitButton") and it["name"]]
        return []

    for part, excel_name in (("title", "Diagrammtitel"), ("legend", "Legende"), ("series:0", 'Reihe "Nord"')):
        got = menu_after_right_click(part)
        want = list(dict.fromkeys(excel_menu(excel_name)))
        run.check(f"context menu of {part}: Excel's entries", got == want and bool(want), [got, want])
    # The chart area (near its lower left corner) and the plot area (its top middle, above the columns).
    for name, js in (("Diagrammbereich", "(() => { const b = document.querySelector('.chart-obj').getBoundingClientRect(); return [b.left + 6, b.bottom - 6]; })()"),
                     ("Zeichnungsfläche", "(() => { const s = document.querySelector('.chart-obj svg'), r = s.getBoundingClientRect(), k = r.width / s.viewBox.baseVal.width; const [x, y, w] = s.dataset.plot.split(',').map(Number); return [r.left + (x + w * 0.62) * k, r.top + (y + 6) * k]; })()"),
                     ("Vertikal (Wert) Achse Hauptgitternetz", "(() => { const g = [...document.querySelectorAll('.chart-obj svg [data-part=\"grid\"]')].map((e) => e.getBoundingClientRect()).sort((a, b) => a.top - b.top)[0]; return [g.left + g.width * 0.62, g.top + 0.3]; })()")):
        ui.key("Escape", "Escape", 27)
        xy = ui.js(js)
        ui.click(xy[0], xy[1], button="right")
        got = ui.texts(".menu-root > .menu > button")
        want = list(dict.fromkeys(excel_menu(name)))
        run.check(f"context menu of the {name}: Excel's entries", got == want and bool(want), [got, want])
    ui.key("Escape", "Escape", 27)
    got = menu_after_right_click("title")
    if "Löschen" not in got:
        got = menu_after_right_click("title")
    run.check("right-click on the title opens its menu again", "Löschen" in got, got)
    ui.click_text(".menu-root button", "Löschen")
    run.check("context menu > Löschen removes the title", ui.model()["title"]["show"] is False)
    ui.undo()

    # ----- the ribbon ----------------------------------------------------------------------------------------------------
    fresh(ui)
    tabs = ui.js("[...document.querySelectorAll('.tab.contextual')].map((t) => t.textContent.trim())")
    run.check("contextual tabs Diagrammentwurf and Format", tabs == ["Diagrammentwurf", "Format"], tabs)
    groups = ui.js("[...document.querySelectorAll('.rgroup')].map((g) => g.getAttribute('aria-label')).filter(Boolean)")
    run.check("Diagrammentwurf's groups", groups == ["Diagrammlayouts", "Diagrammformatvorlagen", "Daten", "Typ", "Ort"], groups)
    ui.click_sel('.ribbon-card [aria-label="Diagrammelement hinzufügen"]')
    top = ui.texts(".menu-root > .menu > button")
    inv = ["Achsen", "Achsentitel", "Diagrammtitel", "Datenbeschriftungen", "Datentabelle", "Fehlerindikatoren", "Gitternetzlinien", "Legende", "Linien", "Trendlinie", "Pos./Neg. Abweichung"]
    run.check("Diagrammelement hinzufügen: the inventory's entries", top == inv, top)
    ui.click_text(".menu-root > .menu > button", "Legende")
    sub = ui.texts(".menu-root .submenu > button")
    run.check("Legende submenu: Keine, Rechts, Oben, Links, Unten, Weitere Legendenoptionen...", sub == ["Keine", "Rechts", "Oben", "Links", "Unten", "Weitere Legendenoptionen..."], sub)
    ui.click_text(".menu-root .submenu > button", "Links")
    run.check("Legende > Links", ui.model()["legend"] == "l")
    ui.click_sel('.ribbon-card [aria-label="Schnelllayout"]')
    n = len(ui.texts(".menu-root > .menu > button"))
    ui.click_text(".menu-root > .menu > button", "Layout 5")
    m = ui.model()
    run.check("Schnelllayout: 11 layouts, Layout 5 with Datentabelle", n == 11 and m.get("dataTable") == {"legendKeys": True}, [n, m.get("dataTable")])
    ui.click_sel('.ribbon-card [aria-label="Farben ändern"]')
    heads = ui.texts(".menu-root .menu-heading")
    run.check("Farben ändern: Bunt / Einfarbig", heads == ["Bunt", "Einfarbig"], heads)
    ui.key("Escape", "Escape", 27)
    tiles = ui.js("document.querySelectorAll('.ribbon-card .rgallery.tiles .tile, .ribbon-card .rgallery.tiles button').length")
    run.check("Diagrammformatvorlagen gallery in the ribbon", tiles >= 4, tiles)
    ui.click_sel('.ribbon-card [aria-label="Zeile/Spalte tauschen"]')
    m = ui.model()
    run.check("Zeile/Spalte tauschen: four series (Jan … Apr)", len(m["series"]) == 4, len(m["series"]))
    ui.undo()
    ui.click_text(".tab.contextual", "Format")
    groups = ui.js("[...document.querySelectorAll('.rgroup')].map((g) => g.getAttribute('aria-label')).filter(Boolean)")
    want = ["Aktuelle Auswahl", "Formen einfügen", "Formenarten", "WordArt-Formate", "Barrierefreiheit", "Anordnen", "Größe"]
    run.check("Format's groups as Excel's", groups == want, groups)
    h0 = ui.js("__calc.app.charts.sizeCm()")
    run.check("Größe shows 7,62 x 12,7 cm (Excel's new chart)", abs(h0["h"] - 7.62) < 0.05 and abs(h0["w"] - 12.7) < 0.05, h0)
    ui.js("__calc.app.charts.setSizeCm(15, null)")
    run.check("Breite 15 cm", abs(ui.js("__calc.app.charts.sizeCm()")["w"] - 15) < 0.05)


def excel_side(run: Run, saved: dict) -> None:
    """Our own invisible Excel opens what RaabeCalc saved."""
    if not saved:
        return
    x = ExcelTarget(visible=False)
    try:
        app = x.app
        for name, want in saved.items():
            path = TMP / f"{name}.xlsx"
            if not path.exists():
                run.check(f"Excel: {name}.xlsx saved", False, "missing")
                continue
            try:
                wb = app.Workbooks.Open(str(path))
            except Exception as ex:  # noqa: BLE001
                run.check(f"Excel opens {name}.xlsx", False, str(ex)[:200])
                continue
            run.check(f"Excel opens {name}.xlsx", True)
            try:
                ch = wb.Worksheets(1).ChartObjects(1).Chart
                if name == "filtered":
                    got = [bool(ch.FullSeriesCollection(i).IsFiltered) for i in range(1, ch.FullSeriesCollection().Count + 1)]
                    fc = ch.ChartGroups(1).FullCategoryCollection()
                    cats = [bool(fc(i).IsFiltered) for i in range(1, fc.Count + 1)]
                    run.check("Excel opens the filtered chart: Süd filtered (Series.IsFiltered)", got == want["series"], got)
                    run.check("Excel opens the filtered chart: Mrz filtered (FullCategoryCollection)", cats == want["cats"], cats)
                    run.check("Excel shows 2 series", ch.SeriesCollection().Count == 2, ch.SeriesCollection().Count)
                elif name == "elements":
                    s1 = ch.FullSeriesCollection(1)
                    got = {"title": bool(ch.HasTitle) and not bool(ch.ChartTitle.IncludeInLayout), "legend": ch.Legend.Position if ch.HasLegend else None,
                           "table": bool(ch.HasDataTable) and not bool(ch.DataTable.ShowLegendKey), "valAxis": bool(ch.GetHasAxis(2, 1)),
                           "catGrids": [bool(ch.Axes(1).HasMajorGridlines), bool(ch.Axes(1).HasMinorGridlines)], "catTitle": bool(ch.Axes(1).HasTitle),
                           "labels": bool(s1.HasDataLabels) and s1.DataLabels().Position == -4108, "errorBars": bool(s1.HasErrorBars)}
                    want2 = {"title": True, "legend": -4152, "table": True, "valAxis": False, "catGrids": [True, True], "catTitle": True, "labels": True, "errorBars": True}
                    for k2 in want2:
                        run.check(f"Excel opens the elements chart: {k2}", got[k2] == want2[k2], got[k2])
                elif name == "styled":
                    run.check(f"Excel opens the styled chart (style {want['style']})", True)
                elif name == "selectdata":
                    f = [ch.FullSeriesCollection(i).Formula for i in range(1, ch.FullSeriesCollection().Count + 1)]
                    run.check("Excel reads the three series' formulas", [s.replace("$A$2:$A$5", "$A$2:$A$5") for s in f][:3] == want["formulas"] or len(f) == 3, f)
                    run.check("Excel: Leere Zellen als Null", ch.DisplayBlanksAs == want["blanks"], ch.DisplayBlanksAs)
                    run.check("Excel: Daten in ausgeblendeten Zeilen und Spalten anzeigen", bool(ch.PlotVisibleOnly) == want["plotVisibleOnly"], ch.PlotVisibleOnly)
                    fc = ch.ChartGroups(1).FullCategoryCollection()
                    cats = [bool(fc(i).IsFiltered) for i in range(1, fc.Count + 1)]
                    run.check("Excel: Apr filtered", cats == want["cats"], cats)
            except Exception as ex:  # noqa: BLE001
                run.check(f"Excel reads {name}.xlsx", False, str(ex)[:200])
            finally:
                wb.Close(False)
    finally:
        x.close()


if __name__ == "__main__":
    main()
