"""Check (Step 11 part 5, D-172): the chart Format pane (Diagrammbereich formatieren, Füllung open) of RaabeCalc,
RaabeWrite and RaabePresent against Office's (probe_office_pane_geometry.py: Excel, Word and PowerPoint draw the same
pane). Each part's box in CSS px from the pane's content edge, matched by name: the close button, the chevron before
the option sets, each option set, each icon tab, each section header, each radio row, the Farbe label and its colour
button, and the pane's width. Positions and heights must agree within 2 px; widths too, except for text whose width
follows the label (radio rows: Office's box is the text, ours the row). Our apps run off screen through DevTools
(the release exe, or --dev URL for a dev server in a headless Edge at 150 %); no real input.

    .venv/Scripts/python.exe check_pane_geometry.py [calc] [write] [present] [--dev URL]

Result: results/2026-11-11-pane-geometry.json.
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

from feellab.cdp_app import AppExe

HERE = Path(__file__).parent
OFFICE = HERE / "results" / "2026-11-11-office-pane-geometry.json"
RESULT = HERE / "results" / "2026-11-11-pane-geometry.json"
TOL = 2.0
REF = {"calc": "excel", "write": "excel", "present": "powerpoint"}  # Word's own pane is narrower: its width is the user's

JS_OURS = r"""(() => {
  const pane = [...document.querySelectorAll('.taskpane')].find((e) => e.offsetWidth > 0);
  if (!pane) return null;
  const r0 = pane.getBoundingClientRect();
  const ox = r0.left + pane.clientLeft, oy = r0.top + pane.clientTop;
  const box = (e) => { const r = e.getBoundingClientRect(); return [+(r.left - ox).toFixed(1), +(r.top - oy).toFixed(1), +r.width.toFixed(1), +r.height.toFixed(1)]; };
  const out = { width: +pane.clientWidth.toFixed(1), parts: {} };
  const put = (k, e) => { if (e && e.offsetWidth) out.parts[k] = box(e); };
  put('close', pane.querySelector('.tp-close'));
  put('anchor', pane.querySelector('.tp-anchor'));
  pane.querySelectorAll('.tp-set').forEach((e) => put('set:' + e.textContent.trim(), e));
  pane.querySelectorAll('.tp-tab').forEach((e) => put('tab:' + e.getAttribute('aria-label'), e));
  pane.querySelectorAll('.tp-sec-head').forEach((e) => put('section:' + e.textContent.trim(), e));
  pane.querySelectorAll('label.tp-choice').forEach((e) => put('radio:' + e.textContent.trim(), e));
  pane.querySelectorAll('.tp-line').forEach((e) => { const l = e.querySelector(':scope > .tp-label'); const c = e.querySelector(':scope > .tp-color');
    if (l && c) { put('label:' + l.textContent.trim(), l); put('color:' + l.textContent.trim(), c); } });
  return out;
})()"""


def office_parts(rec: dict) -> dict:
    els = rec["elements"]
    head = next(e for e in els if e["class"] == "NetUINativeHWNDHost")
    ox, oy = head["rect"][0], head["rect"][1]
    rel = lambda r: [round(r[0] - ox, 1), round(r[1] - oy, 1), r[2], r[3]]  # noqa: E731
    out = {"width": head["rect"][2], "parts": {}}
    p = out["parts"]
    for e in els:
        t, n, c = e["type"], e["name"], e["class"]
        if t == "Button" and n == "Bereich schließen":
            p["close"] = rel(e["rect"])
        elif t == "MenuItem" and c == "NetUIAnchor":
            p["anchor"] = rel(e["rect"])
        elif t == "TabItem":
            p["set:" + n] = rel(e["rect"])
        elif t == "ListItem" and c == "NetUIGalleryButton":
            p["tab:" + n] = rel(e["rect"])
        elif t == "Button" and c == "NetUIRibbonButton" and e["depth"] >= 7 and n != "Diagrammelemente":
            p["section:" + n] = rel(e["rect"])
        elif t == "RadioButton":
            p["radio:" + n] = rel(e["rect"])
        elif t == "Text" and n == "Farbe":
            p["label:Farbe"] = rel(e["rect"])
        elif t == "MenuItem" and n.startswith("Füllfarbe"):
            p["color:Farbe"] = rel(e["rect"])
    return out


def compare(off: dict, ours: dict) -> list[str]:
    fails = []
    if abs(off["width"] - ours["width"]) > TOL:
        fails.append(f"pane width {ours['width']} vs Office {off['width']}")
    for k, a in off["parts"].items():
        b = ours["parts"].get(k)
        if b is None:
            fails.append(f"missing: {k}")
            continue
        # Text decides these widths (and the next option set's place): Office's semibold labels come out about 6 %
        # narrower than the browser's (D-172), so only their row and height are judged.
        keys = ("y", "h") if k.startswith("set:") else ("x", "y", "h") if k.startswith(("radio:", "label:")) else ("x", "y", "w", "h")
        for i, name in enumerate(("x", "y", "w", "h")):
            if name in keys and abs(a[i] - b[i]) > TOL:
                fails.append(f"{k} {name} {b[i]} vs Office {a[i]}")
    return fails


class EdgeDev:
    def __init__(self, url: str, port: int):
        from check_chart_selection import EdgeDev as E
        self._e = E(url, port=port)
        self.page = self._e.page

    def close(self) -> None:
        self._e.close()


def open_pane(app: str, page) -> None:
    if app == "calc":
        while not page.js("typeof window.__calc === 'object'"):
            time.sleep(0.1)
        page.pump(1.0)
        page.js("__calc.reset({'B1':'A','C1':'B','A2':'x','B2':'1','C2':'2','A3':'y','B3':'3','C3':'4'}, 'A1', 1, [])")
        page.js("(() => { const a = __calc.app; a.grid.select(__calc.parse('A1:C3')); a.charts.insert('column'); a.charts.formatSelection(); })()")
    elif app == "write":
        while not page.js("typeof window.__app === 'object'"):
            time.sleep(0.1)
        page.pump(1.0)
        page.js("__app.insertChart('column')")
        page.pump(0.5)
        page.js("__app.run('wchart.formatObject')")
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
        page.js("__present.app.chartUi.openPane()")
    page.pump(0.8)
    # Füllung open, as in Office's recording.
    page.js("""(() => { const h = [...document.querySelectorAll('.taskpane .tp-sec-head')].find((e) => e.textContent.trim() === 'Füllung');
      if (h && h.getAttribute('aria-expanded') !== 'true') h.click(); })()""")
    page.pump(0.5)


def main() -> None:
    apps = [a for a in sys.argv[1:] if a in REF] or list(REF)
    dev = sys.argv[sys.argv.index("--dev") + 1] if "--dev" in sys.argv else None
    office = json.loads(OFFICE.read_text(encoding="utf-8"))
    result = json.loads(RESULT.read_text(encoding="utf-8")) if RESULT.exists() and not dev else {}
    result["what"] = "Our chart Format pane against Office's, CSS px from the pane's content edge (tolerance 2 px)"
    for app in apps:
        off = office_parts(office[REF[app]])
        port = {"calc": 9501, "write": 9503, "present": 9505}[app]
        a = EdgeDev(dev, port) if dev else AppExe(app, port=port)
        try:
            open_pane(app, a.page)
            ours = a.page.js(JS_OURS)
            errors = a.page.errors[:5]
        finally:
            a.close()
        if not ours:
            result[app] = {"error": "pane not found"}
            print(app, "pane not found")
            continue
        fails = compare(off, ours)
        result[app] = {"passed": not fails, "fails": fails, "office": off, "ours": ours, "errors": errors,
                       "parts": len(off["parts"]) + 1}
        print(app, "PASS" if not fails else f"{len(fails)} differences", f"({len(off['parts']) + 1} parts)")
        for f in fails:
            print("   ", f)
    out = Path(__import__("tempfile").gettempdir()) / "raabe-pane-geo" / "dev-result.json" if dev else RESULT
    out.write_text(json.dumps(result, indent=1, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
