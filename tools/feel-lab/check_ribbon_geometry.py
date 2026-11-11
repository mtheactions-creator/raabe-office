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


def main() -> None:
    apps = [a for a in sys.argv[1:] if a in OFFICE] or list(OFFICE)
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
            g10 = sum(1 for g in r["groups"].values() if abs(g["ours"][2] - g["office"][2]) <= 10)
            r["groups_within_10px"] = g10
            print(f"{app} {tab}: {n_ok}/{n} controls within {TOL} px; group widths within {TOL} px {gw}/{len(r['groups'])}, "
                  f"within 10 px {g10}/{len(r['groups'])}; missing {r['missing_groups']}; extra {r['extra_groups']}")
            if "-v" in sys.argv:
                for name, g in r["groups"].items():
                    for x in g["diffs"][:8]:
                        print("   ", name, "-", x)
    for app, tabs in out["apps"].items():
        ok = sum(t.get("groups_within_10px", 0) for t in tabs.values())
        n = sum(len(t.get("groups", {})) for t in tabs.values())
        print(f"{app}: groups within 10 px of Office's width {ok}/{n}")
    path = HERE / "results" / f"{out['date']}-ribbon-geometry.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("wrote", path.name)


if __name__ == "__main__":
    main()
