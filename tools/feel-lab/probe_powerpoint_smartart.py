"""Probe: SmartArt in PowerPoint's files (Step 8, D-072), in a PowerPoint of our own without a window.

1. Record: a SmartArt per layout (Shapes.AddSmartArt with PowerPoint's own layouts) with three nodes; the
   parts PowerPoint writes (data model, layout, style, colours, drawing), their sizes and the data model.
2. Test: the same file with the layout, style and colour parts cut down to their root element with only the
   uniqueId (no definitions of Microsoft's copied). Does PowerPoint open it, still see a SmartArt with its
   layout and nodes, and can it draw it (PNG export)?

    python probe_powerpoint_smartart.py

Result: results/<date>-powerpoint-smartart.json (part names, sizes, the data model of one layout, the test
outcome; no layout definitions). The files stay in captures/smartart/.
"""
from __future__ import annotations

import datetime
import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget

HERE = Path(__file__).parent
CAPT = HERE / "captures" / "smartart"
LAYOUTS = ["default", "process1", "cycle2", "hierarchy1", "vList2", "radial1"]


def main() -> int:
    # Application.SmartArtLayouts needs a visible PowerPoint ("Object does not exist" without one): the probe
    # shows its window (no input is sent) and therefore takes the real-input lock like the input probes.
    import tempfile, time
    lock = Path(tempfile.gettempdir()) / "raabe-real-input.lock"
    if lock.exists() and time.time() - lock.stat().st_mtime < 1800:
        print("real-input lock held: try later")
        return 3
    lock.write_text(f"step8-probe-smartart {datetime.datetime.now().isoformat(timespec='seconds')}", encoding="utf-8")
    try:
        return run()
    finally:
        lock.unlink(missing_ok=True)


def run() -> int:
    try:
        ppt = PowerPointTarget(visible=True)
    except OwnerPowerPointRunning as e:
        print("skipped:", e)
        return 2
    out: dict = {"date": datetime.date.today().isoformat(), "what": "SmartArt parts as PowerPoint writes them, and whether minimal definition parts open", "layouts": []}
    CAPT.mkdir(parents=True, exist_ok=True)
    try:
        app = ppt.app
        pres = ppt.new_presentation(name="smartart", with_window=True)
        lay_by_id = {}
        for i in range(1, app.SmartArtLayouts.Count + 1):
            lay = app.SmartArtLayouts(i)
            lay_by_id[str(lay.Id).rsplit("/", 1)[-1]] = (i, str(lay.Name), str(lay.Category))
        out["layout_count"] = len(lay_by_id)
        out["layout_names"] = {k: [v[1], v[2]] for k, v in lay_by_id.items()}
        for n, key in enumerate(LAYOUTS, start=1):
            slide = pres.Slides.Add(n, 12)
            if key not in lay_by_id:
                out["layouts"].append({"key": key, "error": "not found"})
                continue
            sh = slide.Shapes.AddSmartArt(app.SmartArtLayouts(lay_by_id[key][0]), 100, 80, 600, 380)
            nodes = sh.SmartArt.AllNodes
            texts = []
            for k in range(1, nodes.Count + 1):
                nodes(k).TextFrame2.TextRange.Text = f"Punkt {k}"
                texts.append(f"Punkt {k}")
            out["layouts"].append({"key": key, "name": lay_by_id[key][1], "category": lay_by_id[key][2], "nodes": nodes.Count, "shape": str(sh.Name)})
        path = ppt.tmp / "smartart.pptx"
        ppt.save_copy(pres, path)
        shutil.copy(path, CAPT / "smartart.pptx")
        z = zipfile.ZipFile(path)
        out["parts"] = {n: len(z.read(n)) for n in z.namelist() if "diagram" in n}
        out["slide1_frame"] = re.search(r"<p:graphicFrame>.*?</p:graphicFrame>", z.read("ppt/slides/slide1.xml").decode("utf-8"), re.S).group(0)
        out["slide1_rels"] = z.read("ppt/slides/_rels/slide1.xml.rels").decode("utf-8")
        data1 = next(n for n in z.namelist() if re.match(r"ppt/diagrams/data1\.xml$", n))
        out["data1"] = z.read(data1).decode("utf-8")
        out["content_types"] = [m for m in re.findall(r"<Override [^>]*>", z.read("[Content_Types].xml").decode("utf-8")) if "diagram" in m.lower()]
        # 2. The same file with the definition parts cut down to their uniqueId.
        cut = ppt.tmp / "smartart-min.pptx"
        with zipfile.ZipFile(path) as src, zipfile.ZipFile(cut, "w", zipfile.ZIP_DEFLATED) as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                m = re.match(r"ppt/diagrams/(layout|quickStyle|colors)\d+\.xml$", item.filename)
                if m:
                    xml = data.decode("utf-8")
                    root = re.search(r"<(dgm:(?:layoutDef|styleDef|colorsDef))\b([^>]*)>", xml)
                    attrs = root.group(2)
                    uid = re.search(r'uniqueId="[^"]*"', attrs)
                    ns = " ".join(re.findall(r'xmlns(?::\w+)?="[^"]*"', attrs))
                    data = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<{root.group(1)} {ns} {uid.group(0) if uid else ""}/>'.encode("utf-8")
                dst.writestr(item, data)
        shutil.copy(cut, CAPT / "smartart-min.pptx")
        test: dict = {}
        try:
            p2 = ppt.open_copy(cut, with_window=False)
            res = []
            for n in range(1, p2.Slides.Count + 1):
                s = p2.Slides(n)
                for k in range(1, s.Shapes.Count + 1):
                    sh = s.Shapes(k)
                    if int(sh.HasSmartArt):
                        res.append({"slide": n, "layout": str(sh.SmartArt.Layout.Id).rsplit("/", 1)[-1], "nodes": int(sh.SmartArt.AllNodes.Count),
                                    "texts": [str(sh.SmartArt.AllNodes(i).TextFrame2.TextRange.Text) for i in range(1, sh.SmartArt.AllNodes.Count + 1)]})
            test["opened"] = True
            test["smartart"] = res
            for n in range(1, p2.Slides.Count + 1):
                ppt.export_png(p2, n, CAPT / f"min-slide{n}.png", width=960)
            # Saved again by PowerPoint: does it write the definitions back in full?
            again = ppt.tmp / "smartart-min-saved.pptx"
            ppt.save_copy(p2, again)
            z2 = zipfile.ZipFile(again)
            test["parts_after_save"] = {n: len(z2.read(n)) for n in z2.namelist() if "diagram" in n}
        except Exception as e:  # noqa: BLE001
            test["opened"] = False
            test["error"] = repr(e)[:300]
        out["minimal_test"] = test
    finally:
        print(ppt.close(kill_if_unreachable=True))
    res = HERE / "results" / f"{out['date']}-powerpoint-smartart.json"
    res.write_text(json.dumps(out, indent=1, ensure_ascii=False), encoding="utf-8", newline="\n")
    print("wrote", res.name, json.dumps(out.get("minimal_test", {}))[:600])
    return 0


if __name__ == "__main__":
    sys.exit(main())
