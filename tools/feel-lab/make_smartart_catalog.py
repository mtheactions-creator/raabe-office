"""SmartArt layouts for RaabePresent's engine (Step 8b, D-077), from PowerPoint's own drawing of each layout with
three items (probe_powerpoint_smartart.py): each layout's name, category and id (labels and ids, as the
inventory records them), and the arrangement family and measures our engine uses for it, read from where
PowerPoint put the three items (a row, a column, a grid, a circle, a pyramid, overlapping circles ...), the shape
it gave them and the connectors between them. Writes apps/raabe-present/src/smartart-catalog.gen.ts.

    .venv/Scripts/python.exe make_smartart_catalog.py
"""
from __future__ import annotations

import json
import math
from collections import Counter
from pathlib import Path

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
SRC = sorted((HERE / "results").glob("*-powerpoint-smartart.json"))[-1]
OUT = ROOT / "apps/raabe-present/src/smartart-catalog.gen.ts"
W, H = 640 * 12700, 360 * 12700
KNOWN = {"rect", "roundRect", "ellipse", "triangle", "trapezoid", "hexagon", "chevron", "homePlate", "rightArrow", "circularArrow", "blockArc",
         "flowChartAlternateProcess", "snip1Rect", "round2SameRect", "pentagon", "octagon", "parallelogram", "upArrow", "downArrow", "leftRightArrow", "notchedRightArrow",
         "line", "diamond", "donut", "pie", "plaque", "wedgeRectCallout", "flowChartProcess", "teardrop", "gear6", "gear9", "funnel", "arc", "chord", "frame", "can", "cube"}


def spec_for(layout: dict) -> dict:
    shapes = layout.get("shapes") or []
    texts = [s for s in shapes if s["text"] in ("Eins", "Zwei", "Drei")]
    others = [s for s in shapes if s not in texts and s["w"] * s["h"] > 0]
    cat = layout["category"]
    node = Counter(s["prst"] for s in texts).most_common(1)[0][0] if texts else "rect"
    node = node if node in KNOWN else "roundRect"
    spec: dict = {"family": "grid", "node": node, "aspect": 0.6, "gap": 0.1}
    if len(texts) < 3:
        spec["family"] = {"hierarchy": "hierarchy", "cycle": "cycle", "pyramid": "pyramid", "matrix": "matrix", "relationship": "venn"}.get(cat, "grid")
        return spec
    texts.sort(key=lambda s: ["Eins", "Zwei", "Drei"].index(s["text"]))
    # Shapes that belong to the first item (nearer to it than to the others, not between two items): measured
    # relative to the item's box, drawn with every item by our engine (a picture area above, a band behind ...).
    def centre(s):
        return (s["x"] + s["w"] / 2, s["y"] + s["h"] / 2)
    def dist(a, b):
        return math.hypot(centre(a)[0] - centre(b)[0], centre(a)[1] - centre(b)[1])
    attach = []
    t0 = texts[0]
    for s in others:
        d = [dist(s, t) for t in texts]
        if d[0] > min(d[1:]) or d[0] > max(t0["w"], t0["h"]) * 1.2 or not t0["w"] or not t0["h"]:
            continue
        before = shapes.index(s) < shapes.index(t0)
        look = "line" if s["prst"] == "line" or min(s["w"], s["h"]) == 0 else "bg" if s["w"] * s["h"] >= t0["w"] * t0["h"] * 0.9 else "conn"
        attach.append({"prst": s["prst"] if s["prst"] in KNOWN else "rect", "x": round((s["x"] - t0["x"]) / t0["w"], 3), "y": round((s["y"] - t0["y"]) / t0["h"], 3),
                       "w": round(s["w"] / t0["w"], 3), "h": round(s["h"] / t0["h"], 3), "before": before, "look": look})
    if attach:
        spec["attach"] = attach[:4]
    # The measured drawing for three items, normalised to the frame (our engine draws it as it is for three
    # items and carries it over to other numbers by the family): the items, and the shapes that are neither
    # an item nor attached to one (connectors, rings, backgrounds), with a guessed look.
    nb = lambda s: [round(s["x"] / W, 4), round(s["y"] / H, 4), round(s["w"] / W, 4), round(s["h"] / H, 4)]
    spec["items3"] = [nb(s) for s in texts]
    attached_ids = set()
    for s in others:
        d = [dist(s, t) for t in texts]
        for t in texts:
            if dist(s, t) == min(d) and d[texts.index(t)] <= max(t["w"], t["h"]) * 1.2 and min(d) == d[texts.index(t)] and texts.index(t) > 0:
                attached_ids.add(id(s))
    extras = []
    for s in others:
        d = [dist(s, t) for t in texts]
        near0 = d[0] <= min(d[1:]) and d[0] <= max(t0["w"], t0["h"]) * 1.2
        if near0 or id(s) in attached_ids:
            continue
        look = "line" if s["prst"] == "line" or min(s["w"], s["h"]) == 0 else "bg" if s["w"] * s["h"] >= t0["w"] * t0["h"] * 1.5 else "conn"
        extras.append({"prst": s["prst"] if s["prst"] in KNOWN else "rect", "box": nb(s), "rot": s["rot"], "look": look})
    if extras:
        spec["extras3"] = extras[:12]
    cx = [s["x"] + s["w"] / 2 for s in texts]
    cy = [s["y"] + s["h"] / 2 for s in texts]
    w0, h0 = texts[0]["w"], texts[0]["h"]
    spec["aspect"] = round(h0 / w0, 3) if w0 else 0.6
    same_y = max(cy) - min(cy) < H * 0.05
    same_x = max(cx) - min(cx) < W * 0.05
    # Distances from the frame's centre: on a circle when about equal and the items far apart in angle.
    r = [math.hypot(x - W / 2, y - H / 2) for x, y in zip(cx, cy)]
    on_circle = min(r) > min(W, H) * 0.15 and (max(r) - min(r)) < min(W, H) * 0.08 and not same_y and not same_x
    widths = [s["w"] for s in texts]
    if node in ("triangle", "trapezoid") or (cat == "pyramid"):
        spec["family"] = "pyramid"
    elif node == "ellipse" and cat == "relationship" and not same_y and not same_x:
        spec["family"] = "venn"
    elif on_circle:
        spec["family"] = "cycle"
        conn = Counter(s["prst"] for s in others if s["w"] < w0).most_common(1)
        if conn:
            spec["conn"] = conn[0][0] if conn[0][0] in KNOWN else "rightArrow"
    elif same_y:
        gap = (texts[1]["x"] - (texts[0]["x"] + w0)) / w0 if w0 else 0.1
        spec["gap"] = round(max(0.0, gap), 3)
        between = [s for s in others if texts[0]["x"] + w0 <= s["x"] + s["w"] / 2 <= texts[1]["x"] and s["h"] < h0 * 1.01]
        if between and gap > 0.05:
            c = between[0]
            spec.update({"family": "process", "conn": c["prst"] if c["prst"] in KNOWN else "rightArrow", "connSize": round(c["h"] / max(1, c["w"]), 3)})
        else:
            spec["family"] = "row"
    elif same_x:
        spec["family"] = "column"
        gap = (texts[1]["y"] - (texts[0]["y"] + h0)) / h0 if h0 else 0.1
        spec["gap"] = round(max(0.0, gap), 3)
        if max(widths) - min(widths) > W * 0.05:
            spec["family"] = "pyramid"
    else:
        rows = len({round(y / (H * 0.1)) for y in cy})
        gap = abs(texts[1]["x"] - (texts[0]["x"] + w0)) / w0 if w0 else 0.1
        spec["gap"] = round(min(0.5, gap), 3)
        spec["family"] = "hierarchy" if cat == "hierarchy" else "matrix" if cat == "matrix" else "grid" if rows <= 3 else "grid"
    if spec["family"] == "grid" and not same_y and not same_x:
        # Items on a circle around their common centre (cycles drawn as rings): a cycle.
        mx, my = sum(cx) / 3, sum(cy) / 3
        rr = [math.hypot(x - mx, y - my) for x, y in zip(cx, cy)]
        ang = sorted(math.degrees(math.atan2(y - my, x - mx)) for x, y in zip(cx, cy))
        gaps = [(ang[(i + 1) % 3] - ang[i]) % 360 for i in range(3)]
        if min(rr) > 0 and (max(rr) - min(rr)) < min(rr) * 0.08 and max(gaps) - min(gaps) < 30:
            spec["family"] = "cycle"
    return spec


NAMES = ["Eins", "Zwei", "Drei", "Vier", "Fünf", "Sechs", "Sieben", "Acht"]


def template(layout: dict, n: int) -> dict | None:
    """PowerPoint's whole drawing of a layout with n items, in shares of the frame: the items' boxes in order and
    every other shape (connectors, rings, picture areas, backgrounds) with a guessed look."""
    shapes = layout.get("shapes") or []
    names = NAMES[:n]
    items = [s for s in shapes if s["text"] in names]
    if len({s["text"] for s in items}) < n:
        return None
    items = sorted({s["text"]: s for s in items}.values(), key=lambda s: names.index(s["text"]))
    nb = lambda s: [round(s["x"] / W, 4), round(s["y"] / H, 4), round(s["w"] / W, 4), round(s["h"] / H, 4)]
    a0 = items[0]["w"] * items[0]["h"]
    extras = []
    for s in shapes:
        if s in items or s["w"] * s["h"] == 0 and s["prst"] != "line":
            continue
        look = "line" if s["prst"] == "line" or min(s["w"], s["h"]) == 0 else "bg" if s["w"] * s["h"] >= a0 * 0.9 else "conn"
        extras.append({"prst": s["prst"] if s["prst"] in KNOWN else "rect", "box": nb(s), "rot": s["rot"], "look": look, "before": shapes.index(s) < shapes.index(items[0])})
    return {"items": [nb(s) for s in items], "extras": extras[:3 * n + 6]}


def main() -> None:
    data = json.loads(SRC.read_text(encoding="utf-8"))
    # The other counts measured (probe_powerpoint_smartart.py --items n).
    counts = {}
    for f in sorted((HERE / "results").glob("*-powerpoint-smartart-*.json")):
        n = int(f.stem.rsplit("-", 1)[1])
        counts[n] = {l["id"]: l for l in json.loads(f.read_text(encoding="utf-8"))["layouts"]}
    counts[3] = {l["id"]: l for l in data["layouts"]}
    entries = []
    seen = set()
    for l in data["layouts"]:
        if l.get("error") or l["id"] in seen:
            continue
        seen.add(l["id"])
        spec = spec_for(l)
        tpl = {}
        for n in sorted(counts):
            t = template(counts[n].get(l["id"], {}), n)
            if t:
                tpl[str(n)] = t
        if tpl:
            spec["tpl"] = tpl
        entries.append({"name": l["name"], "category": l["category"], "id": l["id"], "spec": spec})
    fams = Counter(e["spec"]["family"] for e in entries)
    body = ["// Generated by tools/feel-lab/make_smartart_catalog.py from PowerPoint's drawing of each SmartArt layout with three",
            "// items (probe_powerpoint_smartart.py): the layout's name, category and id, and the family and measures our engine",
            "// (smartart-engine.ts) uses for it (D-077). Do not edit.", "",
            'import type { LayoutSpec } from "./smartart-engine";', "",
            "export interface CatalogEntry { name: string; category: string; id: string; spec: LayoutSpec }", "",
            "export const SMARTART_LAYOUTS: CatalogEntry[] = ["]
    for e in entries:
        body.append("  " + json.dumps(e, ensure_ascii=True) + ",")
    body += ["];", ""]
    OUT.write_text("\n".join(body), encoding="utf-8", newline="\n")
    print(len(entries), "layouts", dict(fams))


if __name__ == "__main__":
    main()
