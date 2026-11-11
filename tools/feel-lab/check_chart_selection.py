"""Chart selection marks of RaabeCalc against Excel's, crop by crop (Step 11 part 2, D-170).

The release exe runs off screen and is driven through DevTools (no real input, no lock): the same column, line and pie
chart as probe_excel_chart_selection.py (same data, title, axis title, data labels, trendline), every part selected
in turn (Format > Aktuelle Auswahl, a second click for the single point), and the chart's region captured at the
screen's device pixels. The marks are measured with selmarks.py exactly as Excel's crops were and compared: number of
handles, their size, kind (grey sizing ring or blue dot), ring and fill colour, the frame's colour and thickness.
Positions are reported but not judged (the charts' own layouts differ by a few pixels).

    .venv/Scripts/python.exe check_chart_selection.py            (needs results/2026-11-11-excel-chart-selection.json)
    .venv/Scripts/python.exe check_chart_selection.py --app write|present   (the selected chart against Word's /
                                                                             PowerPoint's, D-171)
    ... --dev http://localhost:<port>/   (the dev server in a headless Edge at 150 %; results stay in TEMP)

Result: results/2026-11-11-calc-chart-selection.json (-write-, -present-). Crops: %TEMP%/raabe-chart-sel/<app>
(local only). RAABE_CALC_EXE / RAABE_WRITE_EXE / RAABE_PRESENT_EXE point at other builds.
"""
from __future__ import annotations

import json
import sys
import tempfile
import time
from pathlib import Path

from feellab import selmarks
from feellab.cdp_app import AppExe

HERE = Path(__file__).parent
EXCEL = HERE / "results" / "2026-11-11-excel-chart-selection.json"
RESULT = HERE / "results" / "2026-11-11-calc-chart-selection.json"
CROPS = Path(tempfile.gettempdir()) / "raabe-chart-sel" / "calc"
SETUP = {"A1": "Monat", "B1": "Nord", "C1": "Süd", "D1": "West", "A2": "Jan", "B2": "3", "C2": "4", "D2": "2", "A3": "Feb", "B3": "5",
         "C3": "2", "D3": "3", "A4": "Mrz", "B4": "2", "C4": "6", "D4": "4", "A5": "Apr", "B5": "4", "C5": "3", "D5": "5"}
KINDS = {"column": "column", "line": "lineMarkers", "pie": "pie"}
MARGIN = 16  # device pixels, as in the Excel probe

# Element names (Format > Aktuelle Auswahl) of each part; "point" is a second click on the series' second point.
NAMES = {"chart": "Diagrammbereich", "plot": "Zeichnungsfläche", "title": "Diagrammtitel", "legend": "Legende",
         "axisTitle": "Vertikal (Wert) Achse Titel", "valAxis": "Vertikal (Wert) Achse", "catAxis": "Horizontal (Kategorie) Achse",
         "majorGrid": "Vertikal (Wert) Achse Hauptgitternetz", "series": 'Reihe "Nord"', "labels": "Datenbeschriftungen",
         "trend": 'Reihe "Nord" Trendlinie 1'}
# Parts Excel selects that RaabeCalc's charts have no element for (yet).
MISSING = {"legendEntry": "a single legend entry is not an element of its own", "label": "a single data label is not an element of its own"}


def setup_chart(page, kind: str) -> None:
    page.js(f"__calc.reset({json.dumps(SETUP)}, 'F20', 1, [])")
    page.js("""(() => { const a = __calc.app; a.grid.select(__calc.parse('A1:D5')); a.charts.insert('%s');
      a.charts.changeSelected((m) => { m.title = { show: true, text: 'Diagrammtitel' }; m.legend = 'r';
        if (m.kind !== 'pie') m.axisTitles = { val: 'Achsentitel' }; m.dataLabels = false;
        m.series[0].trendline = m.kind === 'pie' ? undefined : { type: 'linear' };
        // The probe's Excel chart was made with ChartObjects.Add and ChartType, which keep the classic gap of 150.
        if (m.kind === 'column') { m.gapWidth = 150; m.overlap = 0; } }); })()""" % KINDS[kind])


def box(page) -> dict:
    return page.js("(() => { const r = document.querySelector('.chart-obj.sel').getBoundingClientRect(); return {x: r.left, y: r.top, w: r.width, h: r.height, dpr: devicePixelRatio}; })()")


def shot(page, path: Path, b: dict) -> None:
    m = MARGIN / b["dpr"]
    page.move(3, 3)
    page.pump(0.4)
    page.shot(path, (b["x"] - m, b["y"] - m, b["w"] + 2 * m, b["h"] + 2 * m))


class EdgeDev:
    """The dev server's page in a headless Edge at the screen's 150 % (--dev URL), for iterating between builds."""

    def __init__(self, url: str, port: int = 9487) -> None:
        import os
        import subprocess
        import urllib.request
        from feellab.cdp_app import CdpPage
        self.tag = f"raabe-chartsel-edge-{os.getpid()}"
        prof = Path(tempfile.gettempdir()) / self.tag
        self.proc = subprocess.Popen([r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", "--headless=new",
                                      f"--remote-debugging-port={port}", f"--user-data-dir={prof}", "--window-size=1600,1000",
                                      "--force-device-scale-factor=1.5", "--no-first-run", url])
        deadline = time.time() + 30
        while True:
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=1) as r:
                    pages = [p for p in json.load(r) if p.get("type") == "page"]
                if pages:
                    self.page = CdpPage(pages[0]["webSocketDebuggerUrl"])
                    return
            except OSError:
                pass
            if time.time() > deadline:
                raise TimeoutError("Edge DevTools not reachable")
            time.sleep(0.2)

    def close(self) -> None:
        import subprocess
        self.page.ws.close()
        script = ("Get-CimInstance Win32_Process -Filter \"name='msedge.exe'\" | Where-Object { $_.CommandLine -like '*" + self.tag
                  + "*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }")
        subprocess.run(["powershell", "-NoProfile", "-Command", script], capture_output=True)


def main_other(which: str) -> None:
    """RaabeWrite / RaabePresent (--app write|present): their charts have no selectable inner parts yet, so the
    selected chart's own frame and handles are compared with Word's / PowerPoint's (probe_word_chart_selection.py,
    probe_powerpoint_chart_selection.py: the chart selected as a shape)."""
    ref_name = "word" if which == "write" else "powerpoint"
    ref = json.loads((HERE / "results" / f"2026-11-11-{ref_name}-chart-selection.json").read_text(encoding="utf-8"))
    ref_part = ref["parts"]["chart" if which == "write" else "shape"]
    crops = CROPS.parent / which
    crops.mkdir(parents=True, exist_ok=True)
    port = 9489 if which == "write" else 9493
    app = EdgeDev(sys.argv[sys.argv.index("--dev") + 1], port=port) if "--dev" in sys.argv else AppExe(which, port=port - 4)
    page = app.page
    result: dict = {"what": f"Raabe{which.capitalize()}'s marks of a selected chart against {ref_name}'s (selmarks.compare)", "parts": {}}
    rect_js = "(() => {{ const r = {}; return r ? {{x: r.left, y: r.top, w: r.width, h: r.height, dpr: devicePixelRatio}} : null; }})()"
    try:
        if which == "write":
            while not page.js("typeof window.__app === 'object'"):
                time.sleep(0.1)
            page.pump(1.0)
            page.js("__app.insertChart('column')")
            page.pump(1.0)
            r = page.js(rect_js.format("__app.editor().view.selectedObjectRect()"))
            page.key("ArrowRight")
        else:
            while not page.js("typeof window.__present === 'object' && !!window.__present.app"):
                time.sleep(0.1)
            page.pump(1.5)
            for sel, text in ((".tab", "Einfügen"), (".ribbon-card [aria-label]", "Diagramm"), (".menu .menu-item, .menu button", "Gruppierte Säulen")):
                xy = page.js(f"""(() => {{ const e = [...document.querySelectorAll({json.dumps(sel)})].filter((x) => x.getClientRects().length)
                  .find((x) => (x.getAttribute('aria-label') ?? '') === {json.dumps(text)} || x.textContent.trim() === {json.dumps(text)});
                  if (!e) return null; const r = e.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; }})()""")
                page.click(*xy)
                page.pump(0.6)
            page.pump(1.0)
            cid = page.js("(() => { const a = __present.app; return a.deck.view(a.slidePath()).shapes.find((s) => s.frame === 'chart')?.id ?? null; })()")
            page.js("__present.app.editor.select([])")
            r = page.js(rect_js.format(f"document.querySelector('.stage [data-id=\"{cid}\"]')?.getBoundingClientRect()"))
        page.pump(0.5)
        result["dpr"] = page.js("devicePixelRatio")
        shot(page, crops / "column-none.png", r)
        if which == "write":
            page.click(r["x"] + r["w"] / 2, r["y"] + r["h"] / 2)
        else:
            # Selected as a click on its edge does (the slide's empty placeholder can lie over the chart's middle).
            page.js(f"__present.app.editor.select([{cid}])")
        page.pump(0.6)
        shot(page, crops / "column-chart.png", r)
        ours = selmarks.measure(crops / "column-chart.png", crops / "column-none.png", [MARGIN, MARGIN])
        cmp = selmarks.compare(ours, ref_part)
        cmp["ours"], cmp["office"] = ours["summary"], ref_part.get("summary")
        result["parts"]["column-chart"] = cmp
        result["errors"] = page.errors[:10]
    finally:
        app.close()
    result["passed"] = sum(1 for v in result["parts"].values() if v["passed"])
    result["judged"] = len(result["parts"])
    out = crops / "dev-result.json" if "--dev" in sys.argv else HERE / "results" / f"2026-11-11-{which}-chart-selection.json"
    out.write_text(json.dumps(result, indent=1, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")
    for k, v in result["parts"].items():
        print(k, "PASS" if v["passed"] else "FAIL", v["fails"], "|", v["ours"], "|", v["office"])
    print("errors:", result["errors"])
    sys.exit(0 if result["passed"] == result["judged"] else 1)


def main() -> None:
    if "--app" in sys.argv and sys.argv[sys.argv.index("--app") + 1] in ("write", "present"):
        return main_other(sys.argv[sys.argv.index("--app") + 1])
    CROPS.mkdir(parents=True, exist_ok=True)
    excel = json.loads(EXCEL.read_text(encoding="utf-8"))
    app = EdgeDev(sys.argv[sys.argv.index("--dev") + 1]) if "--dev" in sys.argv else AppExe("calc", port=9483)
    page = app.page
    result: dict = {"what": "RaabeCalc's chart selection marks against Excel's (selmarks.compare), device pixels at the "
                            "screen's scale", "parts": {}}
    try:
        while not page.js("typeof window.__calc === 'object'"):
            time.sleep(0.1)
        page.pump(1.0)
        result["dpr"] = page.js("devicePixelRatio")
        for kind in KINDS:
            setup_chart(page, kind)
            page.pump(0.6)
            b = box(page)
            # Nothing selected: the base for the chart area; the chart area: the base for the inner parts.
            page.js("__calc.app.charts.deselect()")
            page.pump(0.3)
            shot(page, CROPS / f"{kind}-none.png", b)
            page.click(b["x"] + 4, b["y"] + 4)
            page.pump(0.3)
            for part in [p.split("-", 1)[1] for p in excel["parts"] if p.startswith(kind + "-")]:
                key = f"{kind}-{part}"
                if part in MISSING:
                    result["parts"][key] = {"passed": None, "skipped": MISSING[part]}
                    continue
                if part == "labels":
                    # RaabeCalc's Datenbeschriftungen element covers every series' labels, Excel's one series': the other
                    # series' labels are hidden for the measurement (hidden labels get no marks).
                    page.js("__calc.app.charts.changeSelected((m) => { m.dataLabels = true; })")
                    page.js("""(() => { const s = document.createElement('style'); s.id = 'sel-only-first';
                      s.textContent = '.chart-obj [data-part^="labels:"]:not([data-part="labels:0"]) { display: none; }';
                      document.head.append(s); })()""")
                if part == "point":
                    page.js(f"__calc.app.charts.chooseElement({json.dumps(NAMES['series'])})")
                    page.pump(0.2)
                    r = page.js("(() => { const r = document.querySelector('.chart-obj.sel [data-part=\"series:0:1\"]').getBoundingClientRect(); return {x: r.left + r.width / 2, y: r.top + r.height / 2}; })()")
                    page.click(r["x"], r["y"])
                else:
                    page.js(f"__calc.app.charts.chooseElement({json.dumps(NAMES[part])})")
                page.pump(0.3)
                got = page.js("__calc.app.charts.elementName()")
                shot(page, CROPS / f"{key}.png", b)
                if part == "labels":
                    # The base without labels does not fit: compare against the chart area with the labels shown.
                    page.js(f"__calc.app.charts.chooseElement({json.dumps(NAMES['chart'])})")
                    page.pump(0.3)
                    shot(page, CROPS / f"{kind}-chart-labels.png", b)
                    page.js("__calc.app.charts.changeSelected((m) => { m.dataLabels = false; })")
                    page.js("document.getElementById('sel-only-first')?.remove()")
                base = CROPS / (f"{kind}-none.png" if part == "chart" else f"{kind}-chart-labels.png" if part == "labels" else f"{kind}-chart.png")
                if part != "chart" and not (CROPS / f"{kind}-chart.png").exists():
                    raise RuntimeError("the chart area crop must come first")
                ours = selmarks.measure(CROPS / f"{key}.png", base, [MARGIN, MARGIN])
                office = excel["parts"][key]
                cmp = selmarks.compare(ours, office, over_content=part in ("series", "point"))
                cmp["selected"] = got
                cmp["ours"] = ours["summary"]
                cmp["excel"] = office.get("summary")
                result["parts"][key] = cmp
            page.js("__calc.app.charts.deleteSelected()")
        result["errors"] = page.errors[:10]
    finally:
        app.close()
    judged = [v for v in result["parts"].values() if v.get("passed") is not None]
    result["passed"] = sum(1 for v in judged if v["passed"])
    result["judged"] = len(judged)
    out = CROPS / "dev-result.json" if "--dev" in sys.argv else RESULT
    out.write_text(json.dumps(result, indent=1, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")
    for k, v in result["parts"].items():
        print(k, "PASS" if v.get("passed") else "skip" if v.get("passed") is None else "FAIL", v.get("fails") or v.get("skipped") or "",
              "|", v.get("ours", ""), "|", v.get("excel", ""))
    print(f"{result['passed']} / {result['judged']} parts pass; errors: {result['errors']}")
    sys.exit(0 if result["passed"] == result["judged"] else 1)


if __name__ == "__main__":
    main()
