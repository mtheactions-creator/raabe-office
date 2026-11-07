"""SmartArt layouts compared with PowerPoint (Step 8b, D-077): for every layout PowerPoint drew with three items
(probe_powerpoint_smartart.py), our engine lays out the same three items in the same frame (in a headless Edge
on the dev server) and the three item shapes are compared with PowerPoint's: the overlap of each pair of boxes
(intersection over union, 1 = the same box). A layout counts as matched when its mean overlap is 0.7 or more.
Also our diagram is written into a presentation and PowerPoint opens it (invisible, when no POWERPNT.EXE runs):
it must find the diagram with its three items' texts.

    .venv/Scripts/python.exe check_present_smartart.py      (npm run dev:present must be running)

Result: results/<date>-present-smartart.json.
"""
from __future__ import annotations

import base64
import datetime
import json
import tempfile
from pathlib import Path

from feellab.present_page import PresentPage

HERE = Path(__file__).parent
SRC = sorted((HERE / "results").glob("*-powerpoint-smartart.json"))[-1]
SRC5 = sorted((HERE / "results").glob("*-powerpoint-smartart-5.json"))
W, H = 640 * 12700, 360 * 12700


def iou(a: dict, b: dict) -> float:
    x0, y0 = max(a["x"], b["x"]), max(a["y"], b["y"])
    x1, y1 = min(a["x"] + a["w"], b["x"] + b["w"]), min(a["y"] + a["h"], b["y"] + b["h"])
    inter = max(0, x1 - x0) * max(0, y1 - y0)
    union = a["w"] * a["h"] + b["w"] * b["h"] - inter
    return inter / union if union > 0 else 0.0


def main() -> None:
    ppt = json.loads(SRC.read_text(encoding="utf-8"))
    out: dict = {"date": datetime.date.today().isoformat(), "layouts": []}
    page = PresentPage()
    try:
        ours = page.js(f"""(async () => {{ const e = await import('/src/smartart-engine.ts'); const c = await import('/src/smartart-catalog.gen.ts');
            const items = [{{text:'Eins',level:0}},{{text:'Zwei',level:0}},{{text:'Drei',level:0}}];
            return c.SMARTART_LAYOUTS.map((l) => ({{ id: l.id, shapes: e.layout(l.spec, items, {W}, {H}).filter((s) => s.item !== undefined).map((s) => ({{ item: s.item, x: s.x, y: s.y, w: s.w, h: s.h }})) }})); }})()""")
        by = {o["id"]: o for o in ours}
        out["layouts"] = compare_all(ppt, by, ["Eins", "Zwei", "Drei"])
        if SRC5:
            # Five items: what our families make of PowerPoint's three, against PowerPoint's own five.
            ours5 = page.js(f"""(async () => {{ const e = await import('/src/smartart-engine.ts'); const c = await import('/src/smartart-catalog.gen.ts');
                const items = ['Eins','Zwei','Drei','Vier','Fünf'].map((t) => ({{text:t,level:0}}));
                return c.SMARTART_LAYOUTS.map((l) => ({{ id: l.id, shapes: e.layout(l.spec, items, {W}, {H}).filter((s) => s.item !== undefined).map((s) => ({{ item: s.item, x: s.x, y: s.y, w: s.w, h: s.h }})) }})); }})()""")
            ppt5 = json.loads(SRC5[-1].read_text(encoding="utf-8"))
            out["layouts5"] = compare_all(ppt5, {o["id"]: o for o in ours5}, ["Eins", "Zwei", "Drei", "Vier", "Fünf"])
            # The same without the measured five (the family alone, as for nine items and more).
            hold = page.js(f"""(async () => {{ const e = await import('/src/smartart-engine.ts'); const c = await import('/src/smartart-catalog.gen.ts');
                const items = ['Eins','Zwei','Drei','Vier','Fünf'].map((t) => ({{text:t,level:0}}));
                return c.SMARTART_LAYOUTS.map((l) => {{ const spec = {{ ...l.spec, tpl: undefined }}; return {{ id: l.id, shapes: e.layout(spec, items, {W}, {H}).filter((s) => s.item !== undefined).map((s) => ({{ item: s.item, x: s.x, y: s.y, w: s.w, h: s.h }})) }}; }}); }})()""")
            out["layouts5_family"] = compare_all(ppt5, {o["id"]: o for o in hold}, ["Eins", "Zwei", "Drei", "Vier", "Fünf"])
        # Our diagram opened by PowerPoint: written by the app and read back through COM.
        b64 = page.js("""(async () => { const P = window.__present; P.newDoc(); const a = P.app; a.deck.setLayout([0], a.deck.layouts()[6].path); a.setCurrent(0);
            const m = await import('/src/smartart-create.ts'); const c = await import('/src/smartart-catalog.gen.ts');
            m.createDiagram(a, c.SMARTART_LAYOUTS[0], [{text:'Eins',level:0},{text:'Zwei',level:0},{text:'Unterpunkt',level:1},{text:'Drei',level:0}], { x: 838200, y: 1825625, w: 10515600, h: 4351338 });
            return P.bytes(); })()""")
        path = Path(tempfile.gettempdir()) / "raabe-smartart-check.pptx"
        path.write_bytes(base64.b64decode(b64))
        out["powerpoint_opens"] = open_in_powerpoint(path)
    finally:
        page.close()
    out["summary"] = {}
    for key in ("layouts", "layouts5", "layouts5_family"):
        ls = out.get(key) or []
        ok = sum(1 for x in ls if x["iou"] >= 0.7)
        out["summary"][key] = {"layouts": len(ls), "matched": ok, "mean_iou": round(sum(x["iou"] for x in ls) / max(1, len(ls)), 3)}
    p = HERE / "results" / f"{out['date']}-present-smartart.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps({"summary": out["summary"], "powerpoint": out["powerpoint_opens"]}, ensure_ascii=False))


def compare_all(ppt: dict, by: dict, names: list[str]) -> list[dict]:
    res = []
    for l in ppt["layouts"]:
        mine = by.get(l["id"])
        theirs = [s for s in l.get("shapes", []) if s["text"] in names]
        if not mine or len(theirs) < len(names):
            continue
        theirs.sort(key=lambda s: names.index(s["text"]))
        scores = []
        for k, t in enumerate(theirs[: len(names)]):
            m = next((s for s in mine["shapes"] if s["item"] == k), None)
            scores.append(iou(m, t) if m else 0.0)
        res.append({"name": l["name"], "category": l["category"], "iou": round(sum(scores) / len(names), 3)})
    return res


def open_in_powerpoint(path: Path) -> dict:
    from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget
    try:
        t = PowerPointTarget(visible=False)
    except OwnerPowerPointRunning as e:
        return {"skipped": str(e)}
    res: dict = {}
    try:
        pres = t.open_copy(path, with_window=False)
        sl = pres.Slides(1)
        for i in range(1, sl.Shapes.Count + 1):
            sh = sl.Shapes(i)
            if int(sh.HasSmartArt):
                sa = sh.SmartArt
                res = {"smartart": True, "nodes": int(sa.AllNodes.Count), "texts": [str(sa.AllNodes(k).TextFrame2.TextRange.Text) for k in range(1, sa.AllNodes.Count + 1)],
                       "layout": str(sa.Layout.Name)}
    except Exception as e:  # noqa: BLE001
        res = {"error": str(e)[:200]}
    finally:
        res["close"] = t.close(kill_if_unreachable=True)
    return res


if __name__ == "__main__":
    main()
