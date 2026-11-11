"""Check (Step 11, part 5): our Start tab against Office's, control by control. Office's side comes from
probe_office_ribbon_geometry.py (results/<date>-office-ribbon-geometry.json, UI Automation, CSS px relative to the
ribbon's top left); ours from the app through DevTools (no input), in the same window size. For every group with the
same name: its left edge and width; for every control with the same name inside it: its box relative to the group.
A difference over TOL px is reported. Office's add-in groups (Adobe Acrobat, Add-Ins, Copilot, Sprache, Editor) have
no counterpart here; groups that Office shows collapsed (one button) are compared by their box only.

    .venv/Scripts/python.exe check_ribbon_geometry.py [calc] [write] [present]
"""
from __future__ import annotations

import datetime
import json
import time
import sys
from pathlib import Path

HERE = Path(__file__).parent
TOL = 2.0
ADDINS = {"Adobe Acrobat", "Add-Ins", "Copilot", "Sprache", "Editor"}
OFFICE = {"calc": "excel", "write": "word", "present": "powerpoint"}

JS_OURS = """(() => {
  const card = document.querySelector('.ribbon-card');
  const groups = [...card.querySelectorAll('.rgroup')].filter((g) => g.offsetWidth > 0);
  const c0 = groups.length ? groups[0].getBoundingClientRect() : card.getBoundingClientRect();
  const top = Math.min(...groups.map((g) => g.getBoundingClientRect().top));
  const box = (e) => { const r = e.getBoundingClientRect(); return [+(r.left - c0.left).toFixed(1), +(r.top - top).toFixed(1), +r.width.toFixed(1), +r.height.toFixed(1)]; };
  return groups.map((g) => ({
    name: g.getAttribute('aria-label') || g.querySelector('.label')?.textContent?.trim() || '',
    rect: box(g),
    controls: [...g.querySelectorAll('[aria-label], .rsplit, .rcombo')].filter((e) => e.offsetWidth > 0 && !e.classList.contains('arrow') && e.getAttribute('aria-label') !== 'Weitere Optionen' && e.getAttribute('aria-label') !== 'Öffnen')
      .map((e) => {
        // A combo's label sits on its input: its box is the whole combo (input and arrow), as Office's ComboBox.
        const target = e.classList.contains('rsplit') ? e : e.closest('.rsplit') || e.closest('.rcombo') || e;
        const name = e.classList.contains('rsplit') ? (e.querySelector('[aria-label]')?.getAttribute('aria-label') || '') : e.getAttribute('aria-label');
        return { name, rect: box(target) };
      }).filter((c, i, a) => c.name && a.findIndex((d) => d.name === c.name) === i),
  }));
})()"""


def ours(app: str, tabs: list[str] | None = None) -> dict[str, list[dict]]:
    from feellab.calc_catalogue import CalcSide
    from feellab.present_catalogue import PresentSide
    from feellab.targets import raabe_write
    side = CalcSide() if app == "calc" else PresentSide() if app == "present" else raabe_write()
    try:
        from feellab import winapi
        from probe_office_ribbon_geometry import WIDE
        winapi.place(side.hwnd, *WIDE)
        import time
        time.sleep(1.0)
        res = {}
        for tab in tabs or ["Start"]:
            found = side.js(f"""(() => {{ const t = [...document.querySelectorAll('.tabrow [role=tab]')].find((b) => b.textContent.trim() === {json.dumps(tab)}); t && t.click(); return !!t; }})()""")
            time.sleep(0.6)
            res[tab] = side.js(JS_OURS) if found else None
        return res
    finally:
        side.close()


def compare(off: dict, mine: list[dict]) -> dict:
    res: dict = {"groups": {}, "missing_groups": [], "extra_groups": []}
    mine_by = {g["name"]: g for g in mine}
    for g in off["groups"]:
        if g["name"] in ADDINS:
            continue
        m = mine_by.get(g["name"])
        if not m:
            res["missing_groups"].append(g["name"])
            continue
        d: dict = {"office": g["rect"], "ours": m["rect"], "diffs": []}
        dl, dw = m["rect"][0] - g["rect"][0], m["rect"][2] - g["rect"][2]
        if abs(dw) > TOL:
            d["diffs"].append(f"width {m['rect'][2]} vs {g['rect'][2]}")
        if abs(dl) > TOL:
            d["diffs"].append(f"left {m['rect'][0]} vs {g['rect'][0]}")
        mc = {c["name"]: c for c in m["controls"]}
        ctrl_ok = ctrl_n = 0
        for c in g["controls"]:
            o = mc.get(c["name"])
            if not o:
                d["diffs"].append(f"control missing: {c['name']}")
                continue
            ctrl_n += 1
            # Relative to the group's left edge (the groups' own offsets are reported above).
            orx, ory = c["rect"][0] - g["rect"][0], c["rect"][1]
            mrx, mry = o["rect"][0] - m["rect"][0], o["rect"][1]
            delta = [round(mrx - orx, 1), round(mry - ory, 1), round(o["rect"][2] - c["rect"][2], 1), round(o["rect"][3] - c["rect"][3], 1)]
            if max(abs(x) for x in delta) > TOL:
                d["diffs"].append(f"{c['name']}: dx {delta[0]} dy {delta[1]} dw {delta[2]} dh {delta[3]}")
            else:
                ctrl_ok += 1
        d["controls_ok"], d["controls_matched"], d["controls_office"] = ctrl_ok, ctrl_n, len(g["controls"])
        res["groups"][g["name"]] = d
    office_names = {g["name"] for g in off["groups"]}
    res["extra_groups"] = [g["name"] for g in mine if g["name"] not in office_names]
    return res


# ----- contextual tabs (D-185) ------------------------------------------------------------------------------------------
# Office's UI Automation name of a contextual tab → our labels for it (Word's and PowerPoint's table "Layout" is
# "Tabellenlayout" in UI Automation).
OUR_LABELS = {"Tabellenlayout": ["Tabellenlayout", "Layout"]}
CALC_DATA = {"A1": "Monat", "B1": "Nord", "C1": "Süd", "A2": "Jan", "B2": "3", "C2": "4", "A3": "Feb", "B3": "5", "C3": "2",
             "A4": "Mrz", "B4": "2", "C4": "6"}


def png_b64() -> tuple[str, str]:
    """A small picture (the probe's sample) as a file and as base64."""
    import base64
    from probe_office_ribbon_geometry import sample_png
    p = sample_png()
    return p, base64.b64encode(Path(p).read_bytes()).decode()


def select_ours(app: str, side, kind: str) -> None:
    """Makes the same kind of selection as the Office probe, through the app's test API (no real input)."""
    path, b64 = png_b64()
    if app == "present":
        side.js(f"__present.sweepContext({json.dumps(kind)})")
        return
    if app == "write":
        side.js("(() => { __app.newDoc(); __app.closeOthers && __app.closeOthers(); })()")
        time.sleep(0.5)
        if kind == "picture":
            side.js(f"__app.insertPicture({json.dumps(path)})")
        elif kind == "shape":
            side.js("__app.editor().insertShape('rect', 'shape', 0, 200, 200, 133, 80)")
        elif kind == "table":
            side.js("(() => { const e = __app.editor(); e.insertTable(3, 3); })()")
        elif kind == "chart":
            side.js("__app.insertChart('column')")
        return
    side.js(f"__calc.reset({json.dumps(CALC_DATA)}, 'H20', 1, [])")
    time.sleep(0.2)
    if kind == "picture":
        side.js("__calc.app.pictures.insertPicture({ from: { col: 5, row: 1, x: 0, y: 0 }, to: { col: 7, row: 6, x: 0, y: 0 }, "
                f"size: {{ w: 160, h: 107 }}, src: 'data:image/png;base64,{b64}', name: 'Bild 1' }})")
    elif kind == "shape":
        from feellab.cdpinput import Cdp
        side.js("__calc.app.shapes.startDraw('rect')")
        time.sleep(0.2)
        r = side.js("(() => { const r = __calc.app.grid.el.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; })()")
        Cdp(side).click(r[0], r[1])
    elif kind == "table":
        side.js("(() => { const a = __calc.app; a.tables.create(__calc.parse('A1:C4'), true, 'TableStyleMedium2'); a.grid.select(__calc.parse('B2')); })()")
    elif kind == "chart":
        side.js("(() => { const a = __calc.app; a.grid.select(__calc.parse('A1:C4')); a.charts.insert('column'); })()")


def ours_context(app: str, ctx: dict[str, list[str]]) -> dict[str, list[dict] | None]:
    """Our contextual tabs: for each kind of selection, its tabs as JS_OURS measures them."""
    from feellab import winapi
    from feellab.calc_catalogue import CalcSide
    from feellab.present_catalogue import PresentSide
    from feellab.targets import raabe_write
    from probe_office_ribbon_geometry import WIDE
    side = CalcSide() if app == "calc" else PresentSide() if app == "present" else raabe_write()
    res: dict = {}
    try:
        winapi.place(side.hwnd, *WIDE)
        time.sleep(1.0)
        for kind, names in ctx.items():
            try:
                select_ours(app, side, kind)
            except Exception as e:  # noqa: BLE001
                print(app, kind, "selection failed:", str(e)[:200])
                continue
            time.sleep(0.8)
            for n in names:
                labels = OUR_LABELS.get(n, [n])
                found = side.js(f"""(() => {{ const t = [...document.querySelectorAll('.tabrow [role=tab].contextual')].find((b) => {json.dumps(labels)}.includes(b.textContent.trim())); t && t.click(); return !!t; }})()""")
                time.sleep(0.6)
                res[n] = side.js(JS_OURS) if found else None
        return res
    finally:
        side.close()


def summary(r: dict) -> tuple[int, int, int, int]:
    """(controls within TOL, Office's controls, groups whose width is within 10 px, groups compared)."""
    g = r["groups"].values()
    return (sum(x["controls_ok"] for x in g), sum(x["controls_office"] for x in g),
            sum(1 for x in g if abs(x["ours"][2] - x["office"][2]) <= 10), len(r["groups"]))


def main_context(apps: list[str], ref: dict, ref_name: str) -> None:
    out: dict = {"date": datetime.date.today().isoformat(), "reference": ref_name, "tolerance_px": TOL, "apps": {}}
    for app in apps:
        off = (ref.get(OFFICE[app]) or {}).get("context")
        if not off:
            print(app, "no Office recording of the contextual tabs")
            continue
        ctx: dict[str, list[str]] = {}
        for name, rec in off.items():
            if "groups" in rec:
                ctx.setdefault(rec["kind"], []).append(name)
        mine = ours_context(app, ctx)
        out["apps"][app] = {}
        for name, ot in off.items():
            if "groups" not in ot:
                continue
            if not mine.get(name):
                out["apps"][app][name] = {"error": "tab not in our app"}
                print(app, name, "not in our app")
                continue
            r = compare(ot, mine[name])
            out["apps"][app][name] = r
            ok, n, g10, ng = summary(r)
            print(f"{app} {name}: groups within 10 px {g10}/{ng}; controls within {TOL} px {ok}/{n}; "
                  f"missing {r['missing_groups']}; extra {r['extra_groups']}")
            if "-v" in sys.argv:
                for gname, g in r["groups"].items():
                    for x in g["diffs"][:12]:
                        print("   ", gname, "-", x)
    path = HERE / "results" / f"{out['date']}-ribbon-geometry-context.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("wrote", path.name)


def main() -> None:
    apps = [a for a in sys.argv[1:] if a in OFFICE] or list(OFFICE)
    if "--context" in sys.argv:
        ref_path = sorted((HERE / "results").glob("*-office-ribbon-geometry.json"))[-1]
        main_context(apps, json.loads(ref_path.read_text(encoding="utf-8")), ref_path.name)
        return
    ref_path = sorted((HERE / "results").glob("*-office-ribbon-geometry.json"))[-1]
    ref = json.loads(ref_path.read_text(encoding="utf-8"))
    out: dict = {"date": datetime.date.today().isoformat(), "reference": ref_path.name, "tolerance_px": TOL, "apps": {}}
    for app in apps:
        off = ref.get(OFFICE[app])
        if not off or "groups" not in off:
            print(app, "no Office recording")
            continue
        office_tabs = off.get("tabs") or {"Start": off}
        mine = ours(app, list(office_tabs))
        out["apps"][app] = {}
        for tab, ot in office_tabs.items():
            if "groups" not in ot:
                continue
            if not mine.get(tab):
                out["apps"][app][tab] = {"error": "tab not in our app"}
                print(app, tab, "not in our app")
                continue
            r = compare(ot, mine[tab])
            out["apps"][app][tab] = r
            n_ok = sum(g["controls_ok"] for g in r["groups"].values())
            n = sum(g["controls_office"] for g in r["groups"].values())
            gw = sum(1 for g in r["groups"].values() if not any(d.startswith("width") for d in g["diffs"]))
            g10 = summary(r)[2]
            print(f"{app} {tab}: {n_ok}/{n} controls within {TOL} px; group widths within {TOL} px {gw}/{len(r['groups'])} (10 px {g10}); "
                  f"missing {r['missing_groups']}; extra {r['extra_groups']}")
            if "-v" in sys.argv:
                for name, g in r["groups"].items():
                    for x in g["diffs"][:8]:
                        print("   ", name, "-", x)
    path = HERE / "results" / f"{out['date']}-ribbon-geometry.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("wrote", path.name)


if __name__ == "__main__":
    main()
