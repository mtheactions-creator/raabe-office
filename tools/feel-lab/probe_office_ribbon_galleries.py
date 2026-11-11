"""Probe (Step 11, inline galleries, D-180): how Office lays out the galleries it shows inside the ribbon and how they
shrink as the window narrows. Our own Excel, Word and PowerPoint (PowerPoint only when the owner's POWERPNT.EXE is not
running), each window placed at a series of widths; at every width the tab's groups and, for the gallery groups, the
gallery's box, its visible tiles (count, size, rows) and the scroll / more buttons at its right are read through UI
Automation. No input: tabs are selected through UI Automation, the window is only moved and sized.
Writes results/<date>-office-ribbon-galleries.json (Microsoft's labels only, CSS px relative to the ribbon).

    .venv/Scripts/python.exe probe_office_ribbon_galleries.py [excel] [word] [powerpoint] [--dump]
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.inputlock import real_input
from probe_office_ribbon_geometry import ribbon_of

HERE = Path(__file__).parent

# Tab -> the groups whose galleries are read (the other groups only by their box).
TARGETS = {
    "powerpoint": {"Start": ["Zeichnen"], "Entwurf": ["Designs"], "Übergänge": ["Übergang zu dieser Folie"], "Animationen": ["Animation"],
                   "Einfügen": []},
    "word": {"Start": ["Formatvorlagen"], "Entwurf": ["Dokumentformatierung"], "Einfügen": []},
    "excel": {"Start": [], "Einfügen": []},
}
# Window widths in screen pixels (the display runs at 150 %: 3000 px are 2000 CSS px).
WIDTHS = [3000, 2850, 2700, 2550, 2400, 2250, 2100, 1950, 1800, 1650, 1500, 1350, 1200]
HEIGHT = 2050


def walk(el, depth: int = 0, max_depth: int = 4) -> list[dict]:
    out = []
    for c in uia.children(el):
        try:
            d = uia.describe(c)
        except Exception:  # noqa: BLE001
            continue
        node = {"type": d["type"], "name": d["name"], "class": d["class"], "rect": d["rect"]}
        if depth < max_depth:
            kids = walk(c, depth + 1, max_depth)
            if kids:
                node["children"] = kids
        out.append(node)
    return out


def gallery_of(group_el, css) -> dict:
    """The gallery's outer box, its visible tiles and its buttons (scroll up / down, Weitere)."""
    tree = walk(group_el, 0, 5)
    res: dict = {"controls": [], "tiles": [], "grid": None}

    def visit(nodes, inside_grid=False):
        for n in nodes:
            if n["type"] in ("DataGrid", "List") and not inside_grid:
                res["grid"] = css(n["rect"])
                visit(n.get("children", []), True)
                continue
            if inside_grid and n["type"] in ("DataItem", "ListItem", "Button", "MenuItem") and n["rect"][2] > 0:
                res["tiles"].append({"name": n["name"], "rect": css(n["rect"])})
                continue
            if not inside_grid and n["type"] in ("Button", "MenuItem", "SplitButton", "Group", "ComboBox") and n["rect"][2] > 0:
                res["controls"].append({"type": n["type"], "name": n["name"], "rect": css(n["rect"])})
            visit(n.get("children", []), inside_grid)
    visit(tree)
    return res


def record_app(root, hwnd: int, app: str, dump: bool) -> dict:
    out: dict = {}
    for tab, gal_groups in TARGETS[app].items():
        per_width = []
        for w in WIDTHS:
            winapi.place(hwnd, 0, 0, w, HEIGHT)
            time.sleep(1.2)
            groups = ribbon_of(root, tab)
            if not groups:
                per_width.append({"window": w, "error": "no groups"})
                continue
            scale = winapi.dpi_scale(hwnd)
            left = min(g[1]["rect"][0] for g in groups)
            top = min(g[1]["rect"][1] for g in groups)
            css = lambda r: [round((r[0] - left) / scale, 1), round((r[1] - top) / scale, 1), round(r[2] / scale, 1), round(r[3] / scale, 1)]  # noqa: E731
            entry: dict = {"window": w, "window_css": round(w / scale, 1), "groups": []}
            for g, dg in sorted(groups, key=lambda x: x[1]["rect"][0]):
                ge: dict = {"name": dg["name"], "rect": css(dg["rect"])}
                if dg["name"] in gal_groups:
                    ge["gallery"] = gallery_of(g, css)
                    if dump and w == WIDTHS[0]:
                        ge["tree"] = walk(g, 0, 5)
                entry["groups"].append(ge)
            per_width.append(entry)
            print(app, tab, w, [(g["name"], g["rect"][2], len(g.get("gallery", {}).get("tiles", []))) for g in entry["groups"]])
        out[tab] = per_width
    ribbon_of(root, "Start")
    return out


def main() -> None:
    which = [a for a in sys.argv[1:] if not a.startswith("-")] or ["excel", "word", "powerpoint"]
    dump = "--dump" in sys.argv
    path = HERE / "results" / f"{datetime.date.today().isoformat()}-office-ribbon-galleries.json"
    out = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    # No input is sent, but the window covers the screen while it is measured: hold the real-input lock meanwhile, so no
    # other session's real input lands on it.
    with real_input("probe_office_ribbon_galleries"):
        run(which, dump, out)
    out["date"] = datetime.date.today().isoformat()
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print("wrote", path.name)


def run(which: list[str], dump: bool, out: dict) -> None:
    if "excel" in which:
        from feellab.excel import ExcelTarget
        x = ExcelTarget(visible=True)
        try:
            time.sleep(1.5)
            out["excel"] = record_app(uia.element_for(x.hwnd), x.hwnd, "excel", dump)
        finally:
            x.close()
    if "word" in which:
        from feellab.targets import WordTarget
        wd = WordTarget()
        try:
            top = winapi.user32.GetAncestor(int(wd.app.ActiveWindow.Hwnd), 2)
            time.sleep(1.5)
            out["word"] = record_app(uia.element_for(top), top, "word", dump)
        finally:
            wd.close()
    if "powerpoint" in which:
        from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget
        try:
            p = PowerPointTarget(visible=True)
        except OwnerPowerPointRunning as e:
            print("PowerPoint skipped:", e)
        else:
            try:
                p.new_presentation(with_window=True)
                time.sleep(1.5)
                out["powerpoint"] = record_app(uia.element_for(p.hwnd), p.hwnd, "powerpoint", dump)
            finally:
                for pr in p.own:
                    pr.Close()
                if p.app.Presentations.Count == 0:
                    p.app.Quit()


if __name__ == "__main__":
    main()
