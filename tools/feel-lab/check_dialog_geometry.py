"""Check (Step 11 part 5, D-173): the dialog frame of RaabeCalc's Zellen formatieren and RaabeWrite's Absatz and
Schriftart against Office's (probe_office_dialog_geometry.py). In CSS px, relative to the client area (below the title
bar): the title bar's height, the Schließen button, OK and Abbrechen (size, and their distance from the client's right
and bottom edge), the first tab (place and height) and the client's size. Tolerance 2 px. Our apps run off screen
through DevTools (the release exe, or --dev URL ... for dev servers in a headless Edge at 150 %); no real input.

    .venv/Scripts/python.exe check_dialog_geometry.py [calc] [write] [--dev CALC_URL WRITE_URL]

Result: results/2026-11-11-dialog-geometry.json.
"""
from __future__ import annotations

import json
import sys
import tempfile
import time
from pathlib import Path

from feellab.cdp_app import AppExe

HERE = Path(__file__).parent
OFFICE = HERE / "results" / "2026-11-11-office-dialog-geometry.json"
RESULT = HERE / "results" / "2026-11-11-dialog-geometry.json"
TOL = 2.0
CASES = {"calc": [("Zellen formatieren", "__calc.app.commands.run('cells.formatDialog')")],
         "write": [("Absatz", "__app.run('para.dialog')"), ("Schriftart", "__app.run('font.dialog')")]}

JS_OURS = r"""(() => {
  const d = [...document.querySelectorAll('.dialog')].pop();
  if (!d) return null;
  const r0 = d.getBoundingClientRect();
  const title = d.querySelector('.dlg-title').getBoundingClientRect();
  const cx = r0.left + d.clientLeft, cy = title.bottom, cw = d.clientWidth, ch = r0.bottom - d.clientTop - title.bottom;
  const box = (e) => { const r = e.getBoundingClientRect(); return [+(r.left - cx).toFixed(1), +(r.top - cy).toFixed(1), +r.width.toFixed(1), +r.height.toFixed(1)]; };
  const out = { client: [+cw.toFixed(1), +ch.toFixed(1)], titlebar: +(title.bottom - r0.top).toFixed(1), parts: {} };
  out.parts['Schließen'] = box(d.querySelector('.dlg-title button'));
  for (const b of d.querySelectorAll('.dlg-buttons button')) out.parts[b.textContent.trim()] = box(b);
  const tab = d.querySelector('[role=tab], .tabs button, .fc-tabs button');
  if (tab) out.parts['tab'] = box(tab);
  return out;
})()"""


def office_parts(rec: dict) -> dict:
    out = {"client": rec["client"], "titlebar": rec["titlebar"], "parts": {}}
    for e in rec["elements"]:
        if e["type"] == "Button" and e["name"] in ("Schließen", "OK", "Abbrechen"):
            out["parts"][e["name"]] = e["rect"]
        elif e["type"] == "TabItem" and "tab" not in out["parts"]:
            out["parts"]["tab"] = e["rect"]
    return out


def compare(off: dict, ours: dict) -> list[str]:
    fails = []
    if abs(off["titlebar"] - ours["titlebar"]) > TOL:
        fails.append(f"title bar {ours['titlebar']} vs Office {off['titlebar']}")
    for i, n in enumerate(("client width", "client height")):
        if abs(off["client"][i] - ours["client"][i]) > TOL:
            fails.append(f"{n} {ours['client'][i]} vs Office {off['client'][i]}")
    ow, oh = off["client"]
    mw, mh = ours["client"]
    for k, a in off["parts"].items():
        b = ours["parts"].get(k)
        if b is None:
            fails.append(f"missing: {k}")
            continue
        if k in ("OK", "Abbrechen", "Schließen"):
            # From the right edge (and the bottom edge for the button row; the title bar's top for Schließen).
            pa = [ow - a[0] - a[2], (oh - a[1] - a[3]) if k != "Schließen" else a[1], a[2], a[3]]
            pb = [mw - b[0] - b[2], (mh - b[1] - b[3]) if k != "Schließen" else b[1], b[2], b[3]]
            names = ("right", "bottom" if k != "Schließen" else "y", "w", "h")
        else:
            pa, pb, names = [a[0], a[1], a[3]], [b[0], b[1], b[3]], ("x", "y", "h")
        for n, x, y in zip(names, pa, pb):
            if abs(x - y) > TOL:
                fails.append(f"{k} {n} {round(y, 1)} vs Office {round(x, 1)}")
    return fails


def main() -> None:
    apps = [a for a in sys.argv[1:] if a in CASES] or list(CASES)
    dev = sys.argv[sys.argv.index("--dev") + 1:sys.argv.index("--dev") + 3] if "--dev" in sys.argv else None
    office = json.loads(OFFICE.read_text(encoding="utf-8"))
    result = json.loads(RESULT.read_text(encoding="utf-8")) if RESULT.exists() and not dev else {}
    result["what"] = "Our dialogs' frame against Office's, CSS px from the client area (tolerance 2 px)"
    for app in apps:
        port = {"calc": 9511, "write": 9513}[app]
        if dev:
            from check_chart_selection import EdgeDev
            a = EdgeDev(dev[0] if app == "calc" else dev[1], port=port)
        else:
            a = AppExe(app, port=port)
        p = a.page
        try:
            while not p.js("typeof window.__calc === 'object' || typeof window.__app === 'object'"):
                time.sleep(0.1)
            p.pump(1.0)
            for title, js in CASES[app]:
                if title not in office or "error" in office[title]:
                    result[title] = {"error": "Office not recorded"}
                    continue
                p.js(js)
                p.pump(0.8)
                ours = p.js(JS_OURS)
                p.key("Escape")
                p.pump(0.4)
                off = office_parts(office[title])
                fails = compare(off, ours) if ours else ["dialog not found"]
                result[title] = {"passed": not fails, "fails": fails, "office": off, "ours": ours}
                print(title, "PASS" if not fails else f"{len(fails)} differences")
                for f in fails:
                    print("   ", f)
            result.setdefault("errors", {})[app] = p.errors[:5]
        finally:
            a.close()
    out = Path(tempfile.gettempdir()) / "raabe-dialog-geo" / "dev-result.json" if dev else RESULT
    out.write_text(json.dumps(result, indent=1, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
