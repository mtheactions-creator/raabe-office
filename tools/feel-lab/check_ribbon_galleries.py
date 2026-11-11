"""Check (Step 11, inline galleries, D-180): our inline galleries against Office's as the window narrows. Office's side
comes from probe_office_ribbon_galleries.py (results/<date>-office-ribbon-galleries.json: our own Office windows at a
series of widths, read through UI Automation); ours from the apps through DevTools (no input), the window placed at the
same widths. For every gallery group: the group's width, and how many tiles the gallery shows (0 once it is a button).
Tabs where Office shows add-in groups (Start in Word and PowerPoint) give Office less room than ours at the same width;
there the comparison is reported but not counted.

    .venv/Scripts/python.exe check_ribbon_galleries.py [write] [present]
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

HERE = Path(__file__).parent
OFFICE = {"write": "word", "present": "powerpoint"}
# Office's tabs with add-in groups (Adobe Acrobat, Copilot, Add-Ins, Sprache): their galleries have less room there.
ADDIN_TABS = {("word", "Start"), ("powerpoint", "Start")}
TOL_W = 10.0

JS_GROUPS = """(() => {
  const groups = [...document.querySelectorAll('.ribbon-card > .rgroup')].filter((g) => g.offsetWidth > 0);
  return groups.map((g) => {
    const gal = g.querySelector('.rgallery');
    const tiles = gal ? [...gal.querySelectorAll('.items > .item')].filter((t) => t.style.display !== 'none' && t.offsetWidth > 0) : [];
    // Tiles of the first row only (a grid gallery shows several rows).
    const top = tiles.length ? Math.min(...tiles.map((t) => t.getBoundingClientRect().top)) : 0;
    const row = tiles.filter((t) => Math.abs(t.getBoundingClientRect().top - top) < 2);
    return { name: g.getAttribute('aria-label') || g.querySelector('.label')?.textContent?.trim() || g.querySelector('[data-tip]')?.dataset.tip?.split(String.fromCharCode(10))[0] || '',
             width: +g.getBoundingClientRect().width.toFixed(1), tiles: row.length,
             tile: row.length ? [+row[0].getBoundingClientRect().width.toFixed(1), +row[0].getBoundingClientRect().height.toFixed(1)] : null,
             box: gal ? +gal.getBoundingClientRect().width.toFixed(1) : null };
  });
})()"""


def office_tiles(g: dict) -> int:
    gal = g.get("gallery")
    if not gal or not gal.get("grid"):
        return 0
    grid = gal["grid"]
    # A collapsed group still lists its gallery (in the popup it would open): no tiles in the ribbon then.
    if grid[0] + grid[2] > g["rect"][0] + g["rect"][2] + 1:
        return 0
    vis = [t for t in gal["tiles"] if grid[0] - 1 <= t["rect"][0] and t["rect"][0] + t["rect"][2] <= grid[0] + grid[2] + 1
           and abs(t["rect"][1] - grid[1]) < 2]
    return len(vis)


def ours(app: str, tabs: list[str], widths: list[int], height: int) -> dict:
    from feellab import winapi
    from feellab.present_catalogue import PresentSide
    from feellab.targets import raabe_write
    side = PresentSide() if app == "present" else raabe_write()
    res: dict = {}
    try:
        for tab in tabs:
            res[tab] = {}
            side.js(f"""(() => {{ const t = [...document.querySelectorAll('.tabrow [role=tab]')].find((b) => b.textContent.trim() === {json.dumps(tab)}); t && t.click(); }})()""")
            for w in widths:
                winapi.place(side.hwnd, 0, 0, w, height)
                time.sleep(0.8)
                res[tab][w] = side.js(JS_GROUPS)
        return res
    finally:
        side.close()


def main() -> None:
    apps = [a for a in sys.argv[1:] if a in OFFICE] or list(OFFICE)
    ref_path = sorted((HERE / "results").glob("*-office-ribbon-galleries.json"))[-1]
    ref = json.loads(ref_path.read_text(encoding="utf-8"))
    out: dict = {"date": datetime.date.today().isoformat(), "reference": ref_path.name, "tolerance_px": TOL_W, "apps": {}}
    for app in apps:
        off = ref.get(OFFICE[app])
        if not off:
            continue
        tabs = {t: v for t, v in off.items() if any(g.get("gallery") for e in v for g in e.get("groups", []))}
        widths = sorted({e["window"] for v in tabs.values() for e in v}, reverse=True)
        mine = ours(app, list(tabs), widths, 2050)
        out["apps"][app] = {}
        for tab, entries in tabs.items():
            rows = []
            for e in entries:
                mg = {g["name"]: g for g in mine[tab].get(e["window"], [])}
                for g in e.get("groups", []):
                    if "gallery" not in g:
                        continue
                    m = mg.get(g["name"])
                    rows.append({"window": e["window"], "group": g["name"], "office_width": g["rect"][2], "office_tiles": office_tiles(g),
                                 "ours_width": m["width"] if m else None, "ours_tiles": m["tiles"] if m else None,
                                 "counted": (OFFICE[app], tab) not in ADDIN_TABS})
            out["apps"][app][tab] = rows
            for r in rows:
                ok_w = r["ours_width"] is not None and abs(r["ours_width"] - r["office_width"]) <= TOL_W
                print(f"{app} {tab} {r['window']} {r['group']}: Office {r['office_width']} px, {r['office_tiles']} tiles; "
                      f"ours {r['ours_width']} px, {r['ours_tiles']} tiles {'ok' if ok_w and r['office_tiles'] == r['ours_tiles'] else 'DIFF'}"
                      f"{'' if r['counted'] else ' (add-ins in Office)'}")
        counted = [r for t in out["apps"][app].values() for r in t if r["counted"]]
        same_t = sum(1 for r in counted if r["ours_tiles"] == r["office_tiles"])
        same_w = sum(1 for r in counted if r["ours_width"] is not None and abs(r["ours_width"] - r["office_width"]) <= TOL_W)
        out["apps"][app]["summary"] = {"cases": len(counted), "same_tiles": same_t, "width_within_10px": same_w}
        print(f"{app}: {len(counted)} cases (tabs without add-ins): tiles as Office's {same_t}, group width within {TOL_W} px {same_w}")
    path = HERE / "results" / f"{out['date']}-ribbon-galleries.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print("wrote", path.name)


if __name__ == "__main__":
    main()
