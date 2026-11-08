"""SmartArt-Design > Farben ändern and SmartArt-Formatvorlagen as PowerPoint applies them (D-080, D-101). An
invisible PowerPoint of our own (only when no POWERPNT.EXE runs; its window on the virtual display, D-073) puts
an "Einfache Blockliste" with three items on a slide for every colour set (SmartArtColors) and every quick style
(SmartArtQuickStyles), saves the presentation and reads, for each, the colours part (colorsDef) or the quick
style part (styleDef) and PowerPoint's drawing of it.

    .venv/Scripts/python.exe probe_powerpoint_smartart_styles.py

Result: results/<date>-powerpoint-smartart-styles.json with names, ids and the parameters our drawing uses
(per style label: the colour lists and their method; the style references, 3-D scene and shape); the full parts
and drawings stay in captures/smartart-styles/ (not committed), as for the layouts (D-077).
"""
from __future__ import annotations

import datetime
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from feellab import winapi
from feellab.powerpoint import PowerPointTarget
from check_present_screens import display, monitors

HERE = Path(__file__).parent
CAP = HERE / "captures" / "smartart-styles"
FRAME = (40.0, 60.0, 640.0, 360.0)
BASIC = "urn:microsoft.com/office/officeart/2005/8/layout/default"


A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
DGM = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}"


def colour(el) -> str:
    """A colour element as a short spec: 'accent1', 'accent1 tint=40000 alpha=90000', 'FFFFFF'."""
    mods = " ".join(f"{m.tag.replace(A, '')}={m.get('val')}" for m in el)
    return (el.get("val", "") + (" " + mods if mods else "")).strip()


def colors_def(xml: str) -> dict:
    root = ET.fromstring(xml.encode("utf-8"))
    labels = {}
    for lbl in root.iter(DGM + "styleLbl"):
        e = {}
        for k, t in (("fill", "fillClrLst"), ("line", "linClrLst"), ("effect", "effectClrLst"), ("txLine", "txLinClrLst"), ("txFill", "txFillClrLst"), ("txEffect", "txEffectClrLst")):
            lst = lbl.find(DGM + t)
            if lst is not None and len(lst):
                e[k] = {"meth": lst.get("meth", "span"), "clr": [colour(c) for c in lst]}
                if lst.get("hueDir"):
                    e[k]["hueDir"] = lst.get("hueDir")
        labels[lbl.get("name")] = e
    return labels


def style_def(xml: str) -> dict:
    out: dict = {}
    for name, body in re.findall(r"<dgm:styleLbl name=\"([^\"]+)\">(.*?)</dgm:styleLbl>", xml, re.S):
        refs = {k: int(v) for k, v in re.findall(r"<a:(lnRef|fillRef|effectRef) idx=\"(\d+)\"", body)}
        font = re.search(r"<a:fontRef idx=\"(\w+)\"", body)
        cam = re.search(r"<a:camera prst=\"(\w+)\"", body)
        rig = re.search(r"<a:lightRig rig=\"(\w+)\" dir=\"(\w+)\"", body)
        sp3d = re.search(r"<dgm:sp3d\b([^>]*?)(/>|>(.*?)</dgm:sp3d>)", body, re.S)
        e: dict = {"refs": refs, "font": font.group(1) if font else None}
        if cam and cam.group(1) != "orthographicFront":
            e["camera"] = cam.group(1)
        crot = re.search(r"<a:camera[^>]*>\s*<a:rot lat=\"(\d+)\" lon=\"(\d+)\" rev=\"(\d+)\"", body)
        if crot:
            e["cameraRot"] = [int(x) for x in crot.groups()]
        if rig and rig.group(1) != "threePt":
            e["light"] = [rig.group(1), rig.group(2)]
        if sp3d:
            attrs = dict(re.findall(r"(\w+)=\"([^\"]+)\"", sp3d.group(1)))
            bevel = re.findall(r"<a:(bevelT|bevelB)\b([^>]*)/>", sp3d.group(3) or "")
            if attrs or bevel:
                e["sp3d"] = {**attrs, **{b: dict(re.findall(r"(\w+)=\"([^\"]+)\"", a)) for b, a in bevel}}
        out[name] = e
    return out


def drawing_looks(xml: str) -> list[dict]:
    """Per drawn shape: its preset, its fill and line as written, effects and 3-D present (for comparison)."""
    out = []
    for sp in re.findall(r"<dsp:sp\b.*?</dsp:sp>", xml, re.S):
        pr = re.search(r"<dsp:spPr>(.*?)</dsp:spPr>", sp, re.S)
        body = pr.group(1) if pr else ""
        prst = re.search(r"prst=\"(\w+)\"", body)
        fill = re.search(r"<a:(solidFill|gradFill|noFill|blipFill)\b", body)
        out.append({"prst": prst.group(1) if prst else None, "fill": fill.group(1) if fill else None,
                    "effects": re.findall(r"<a:(outerShdw|innerShdw|glow|reflection|softEdge)\b", body),
                    "scene3d": "<a:scene3d" in body, "sp3d": "<a:sp3d" in body})
    return out


def reparse() -> None:
    res = sorted((HERE / "results").glob("*-powerpoint-smartart-styles.json"))[-1]
    out = json.loads(res.read_text(encoding="utf-8"))
    with zipfile.ZipFile(CAP / "smartart-styles.pptx") as z:
        for kind in ("colors", "styles"):
            for info in out[kind]:
                if "slide" not in info:
                    continue
                rels = z.read(f"ppt/slides/_rels/slide{info['slide']}.xml.rels").decode()
                part = re.search(r'Target="\.\./diagrams/(%s\d+\.xml)"' % ("colors" if kind == "colors" else "quickStyle"), rels)
                if part:
                    xml = z.read(f"ppt/diagrams/{part.group(1)}").decode("utf-8")
                    info["labels"] = colors_def(xml) if kind == "colors" else style_def(xml)
    res.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print("reparsed", res.name)


def main() -> None:
    if "--reparse" in sys.argv:
        reparse()
        return
    CAP.mkdir(parents=True, exist_ok=True)
    out: dict = {"date": datetime.date.today().isoformat(), "colors": [], "styles": []}
    before = monitors()
    mons = display("extend", 2)
    virtual = [m for m in mons if m not in before] or mons[1:]
    t = PowerPointTarget(visible=True, place=False)
    res = HERE / "results" / f"{out['date']}-powerpoint-smartart-styles.json"
    try:
        pres = t.new_presentation(with_window=True, name="smartart-styles")
        if virtual and t.hwnd:
            vx, vy, vw, vh = virtual[0]
            t.app.WindowState = 1
            winapi.place(t.hwnd, vx, vy, vw, vh)
        lays = t.app.SmartArtLayouts
        lay = next(lays(i) for i in range(1, int(lays.Count) + 1) if str(lays(i).Id) == BASIC)
        todo = []
        for kind, coll in (("colors", t.app.SmartArtColors), ("styles", t.app.SmartArtQuickStyles)):
            for i in range(1, int(coll.Count) + 1):
                it = coll(i)
                info = {"index": i, "name": str(it.Name), "id": str(it.Id), "category": str(it.Category)}
                try:
                    sl = pres.Slides.Add(pres.Slides.Count + 1, 12)
                    sh = sl.Shapes.AddSmartArt(lay, *FRAME)
                    sa = sh.SmartArt
                    if kind == "colors":
                        sa.Color = it
                    else:
                        sa.QuickStyle = it
                    info["slide"] = int(sl.SlideIndex)
                except Exception as e:  # noqa: BLE001
                    info["error"] = str(e)[:160]
                out[kind].append(info)
                todo.append((kind, info))
        path = CAP / "smartart-styles.pptx"
        t.save_copy(pres, path)
        with zipfile.ZipFile(path) as z:
            for kind, info in todo:
                if "slide" not in info:
                    continue
                rels = z.read(f"ppt/slides/_rels/slide{info['slide']}.xml.rels").decode()
                part = re.search(r'Target="\.\./diagrams/(%s\d+\.xml)"' % ("colors" if kind == "colors" else "quickStyle"), rels)
                drw = re.search(r'Target="\.\./diagrams/(drawing\d+\.xml)"', rels)
                if part:
                    xml = z.read(f"ppt/diagrams/{part.group(1)}").decode("utf-8")
                    info["labels"] = colors_def(xml) if kind == "colors" else style_def(xml)
                    if kind == "styles":
                        scene = re.search(r"<dgm:styleDef\b.*?<dgm:scene3d>(.*?)</dgm:scene3d>", xml, re.S)
                        info["scene"] = {"camera": re.search(r"prst=\"(\w+)\"", scene.group(1)).group(1)} if scene else None
                if drw:
                    info["drawing"] = drawing_looks(z.read(f"ppt/diagrams/{drw.group(1)}").decode("utf-8"))
        for i in range(1, pres.Slides.Count + 1):
            t.export_png(pres, i, CAP / f"slide-{i}.png", 960)
        pres.Saved = -1
    finally:
        out["close"] = t.close(kill_if_unreachable=True)
        display("internal", 1)
    res.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(len(out["colors"]), "colour sets,", len(out["styles"]), "quick styles;", sum(1 for k in ("colors", "styles") for x in out[k] if x.get("labels")), "read")


if __name__ == "__main__":
    main()
