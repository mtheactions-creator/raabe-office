"""Is the effect the right one? (D-137) The options a student uses most, triggered by hand in RaabePresent (CDP mouse
clicks on the ribbon, as sweep_present.py does) and the same option in a PowerPoint of our own on the same starting
file (the ribbon command by its idMso through ExecuteMso where PowerPoint has one, else the COM call the command
makes). Both results are saved; what the option changed is read from each file (shape geometry, fill and outline, text
body, paragraph and run properties, slide layout, transition, animations, table grid, slide size) and the two changes
are compared.

    .venv/Scripts/python.exe compare_present_office.py [--dev] [--only NAME]

PowerPoint runs only when no POWERPNT.EXE is running (it would be the owner's); it shows a window of its own that gets
no input. Result: results/<date>-present-office-compare.json.
"""
from __future__ import annotations

import argparse
import base64
import datetime
import io
import json
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from sweep_present import TMP, Sweeper, click_at, start

HERE = Path(__file__).parent
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main", "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
A, P = "{%s}" % NS["a"], "{%s}" % NS["p"]
CM = 72 / 2.54

# (name, starting point, our steps, PowerPoint's action). Steps: tab, control name, then menu entries (names); a
# ("type", text) step types into the field; ("key", combo) presses keys. PowerPoint: ("mso", idMso) or ("com", fn).
CASES: list[tuple] = [
    ("Fett", "shape-text", ["Start", "Fett"], ("mso", "Bold")),
    ("Kursiv", "shape-text", ["Start", "Kursiv"], ("mso", "Italic")),
    ("Unterstreichen", "shape-text", ["Start", "Unterstreichen"], ("mso", "Underline")),
    ("Durchstreichen", "shape-text", ["Start", "Durchstreichen"], ("mso", "Strikethrough")),
    ("Schatten", "shape-text", ["Start", "Schatten"], ("mso", "Shadow")),
    ("Schriftgrad vergrößern", "shape-text", ["Start", "Schriftgrad vergrößern"], ("mso", "FontSizeIncrease")),
    ("Schriftgrad verkleinern", "shape-text", ["Start", "Schriftgrad verkleinern"], ("mso", "FontSizeDecrease")),
    ("Schriftart Arial", "shape-text", ["Start", "Schriftart", ("type", "Arial")], ("com", lambda c: setattr(c.tr.Font, "Name", "Arial"))),
    ("Schriftgrad 24", "shape-text", ["Start", "Schriftgrad", ("type", "24")], ("com", lambda c: setattr(c.tr.Font, "Size", 24))),
    ("Schriftfarbe Rot", "shape-text", ["Start", "Schriftfarbe (Pfeil)", "Rot"], ("com", lambda c: setattr(c.tr.Font.Color, "RGB", 0x0000FF))),
    ("Linksbündig", "shape-text", ["Start", "Linksbündig"], ("mso", "AlignLeft")),
    ("Rechtsbündig", "shape-text", ["Start", "Rechtsbündig"], ("mso", "AlignRight")),
    ("Blocksatz", "shape-text", ["Start", "Blocksatz"], ("mso", "AlignJustify")),
    ("Aufzählungszeichen", "shape-text", ["Start", "Aufzählungszeichen"], ("com", lambda c: setattr(c.tr.ParagraphFormat.Bullet, "Visible", -1))),
    ("Nummerierung", "shape-text", ["Start", "Nummerierung"], ("com", lambda c: setattr(c.tr.ParagraphFormat.Bullet, "Type", 2))),
    ("Einzug vergrößern", "placeholder", ["Start", "Einzug vergrößern"], ("mso", "IndentIncrease")),
    ("Zeilenabstand 1,5", "shape-text", ["Start", "Zeilenabstand", "1,5"], ("com", lambda c: (setattr(c.tr.ParagraphFormat, "LineRuleWithin", -1), setattr(c.tr.ParagraphFormat, "SpaceWithin", 1.5)))),
    ("Zwei Spalten", "shape-text", ["Start", "Spalten", "Zwei Spalten"], ("com", lambda c: setattr(c.sh.TextFrame2.Column, "Number", 2))),
    ("Text oben ausrichten", "shape-text", ["Start", "Text ausrichten", "Oben"], ("com", lambda c: setattr(c.sh.TextFrame, "VerticalAnchor", 1))),
    ("GROSSBUCHSTABEN", "shape-text", ["Start", "Groß-/Kleinschreibung", "GROSSBUCHSTABEN"], ("com", lambda c: c.tr.ChangeCase(3))),
    ("Rechtsdrehung 90°", "shape", ["Formformat", "Drehen", "Rechtsdrehung 90 Grad"], ("mso", "ObjectRotateRight90")),
    ("Horizontal kippen", "shape", ["Formformat", "Drehen", "Horizontal spiegeln"], ("mso", "ObjectFlipHorizontal")),
    ("Keine Füllung", "shape", ["Formformat", "Fülleffekt (Pfeil)", "Keine Füllung"], ("com", lambda c: setattr(c.sh.Fill, "Visible", 0))),
    ("Füllung Rot", "shape", ["Formformat", "Fülleffekt (Pfeil)", "Rot"], ("com", lambda c: setattr(c.sh.Fill.ForeColor, "RGB", 0x0000FF))),
    ("Keine Kontur", "shape", ["Formformat", "Formkontur (Pfeil)", "Keine Kontur"], ("com", lambda c: setattr(c.sh.Line, "Visible", 0))),
    ("Kontur 3 Pt", "shape", ["Formformat", "Formkontur (Pfeil)", "Stärke", "3 Pt."], ("com", lambda c: setattr(c.sh.Line, "Weight", 3))),
    ("Formhöhe 3 cm", "shape", ["Formformat", "Formenhöhe", ("type", "3 cm")], ("com", lambda c: setattr(c.sh, "Height", 3 * CM))),
    ("Formbreite 5 cm", "shape", ["Formformat", "Formenbreite", ("type", "5 cm")], ("com", lambda c: setattr(c.sh, "Width", 5 * CM))),
    ("In den Hintergrund", "two-shapes-second", ["Start", "Anordnen", "In den Hintergrund"], ("mso", "ObjectSendToBack")),
    ("In den Vordergrund", "two-shapes-first", ["Start", "Anordnen", "In den Vordergrund"], ("mso", "ObjectBringToFront")),
    ("Gruppieren", "two-shapes", ["Formformat", "Gruppieren", "Gruppieren"], ("mso", "ObjectsGroup")),
    ("Linksbündig ausrichten", "two-shapes", ["Formformat", "Ausrichten", "Linksbündig"], ("mso", "ObjectsAlignLeftSmart")),
    ("Oben ausrichten", "two-shapes", ["Formformat", "Ausrichten", "Oben ausrichten"], ("mso", "ObjectsAlignTopSmart")),
    ("Neue Folie", "slides", ["Start", "Neue Folie"], ("mso", "SlideNew")),
    ("Layout Titel und Inhalt", "slides", ["Start", "Layout", "Titel und Inhalt"], ("com", lambda c: setattr(c.slide, "CustomLayout", c.layout("Titel und Inhalt")))),
    ("Folie duplizieren", "slides", ["Start", "Neue Folie (Pfeil)", "Ausgewählte Folien duplizieren"], ("mso", "DuplicateSelectedSlides")),
    # COM's EntryEffect does not set the duration the ribbon's gallery sets: only the kind of transition is compared.
    ("Übergang Verblassen", "slides", ["Übergänge", "Verblassen"], ("com", lambda c: setattr(c.slide.SlideShowTransition, "EntryEffect", 3849)), "kind"),
    ("Übergang Wischen", "slides", ["Übergänge", "Wischen"], ("com", lambda c: setattr(c.slide.SlideShowTransition, "EntryEffect", 3852)), "kind"),
    ("Übergangsdauer 1,5", "slides", ["Übergänge", "Dauer", ("type", "01,50")], ("com", lambda c: setattr(c.slide.SlideShowTransition, "Duration", 1.5))),
    ("Foliengröße 4:3", "slides", ["Entwurf", "Foliengröße", "Standard (4:3)"], ("com", lambda c: setattr(c.pres.PageSetup, "SlideSize", 1))),
    ("Animation Erscheinen", "shape", ["Animationen", "Erscheinen"], ("com", lambda c: c.slide.TimeLine.MainSequence.AddEffect(c.sh, 1, 0, 1))),
    ("Animation Verblassen", "shape", ["Animationen", "Verblassen"], ("com", lambda c: c.slide.TimeLine.MainSequence.AddEffect(c.sh, 10, 0, 1))),
    ("Animation Einfliegen", "shape", ["Animationen", "Einfliegen"], ("com", lambda c: c.slide.TimeLine.MainSequence.AddEffect(c.sh, 2, 0, 1))),
    ("Zeile oberhalb einfügen", "table", ["Tabellenlayout", "Zeile Oberhalb"], ("mso", "TableInsertRowsAbove")),
    ("Spalte rechts einfügen", "table", ["Tabellenlayout", "Spalte rechts einfügen"], ("mso", "TableInsertColumnsRight")),
    ("Tabelle löschen", "table", ["Tabellenlayout", "Löschen", "Tabelle löschen"], ("mso", "TableDelete")),
    ("Bild Rechtsdrehung 90°", "picture", ["Bildformat", "Drehen", "Rechtsdrehung 90 Grad"], ("mso", "ObjectRotateRight90")),
    ("Bild zurücksetzen", "picture-rotated", ["Bildformat", "Bild zurücksetzen"], ("mso", "PictureReset")),
]

IGNORE_ATTRS = {"lang", "altLang", "dirty", "err", "smtClean", "smtId", "bmk", "kumimoji", "noProof", "kern"}


# ----- facts of a presentation ----------------------------------------------------------------------------------

def _color(el) -> str | None:
    if el is None:
        return None
    for kind in ("srgbClr", "schemeClr", "sysClr", "prstClr"):
        c = el.find(f"a:{kind}", NS)
        if c is not None:
            mods = ",".join(f"{m.tag.split('}')[1]}={m.get('val')}" for m in c)
            return f"{kind}:{c.get('val') or c.get('lastClr')}{'(' + mods + ')' if mods else ''}"
    return el.tag.split("}")[1]


def _attrs(el, skip=()) -> str:
    if el is None:
        return ""
    return ",".join(f"{k}={v}" for k, v in sorted(el.attrib.items()) if k.split("}")[-1] not in IGNORE_ATTRS and k not in skip)


def _fill(sppr) -> str | None:
    if sppr is None:
        return None
    for kind in ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"):
        f = sppr.find(f"a:{kind}", NS)
        if f is not None:
            return _color(f) if kind == "solidFill" else kind
    return None


def shape_facts(sp, key: str, out: dict) -> None:
    tag = sp.tag.split("}")[1]
    out[f"{key}/tag"] = tag
    sppr = sp.find("p:spPr", NS) if tag != "grpSp" else sp.find("p:grpSpPr", NS)
    xfrm = sppr.find("a:xfrm", NS) if sppr is not None else sp.find("p:xfrm", NS)
    if xfrm is not None:
        off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
        # Positions to the nearest 0.5 mm (the two programs round half points differently).
        q = lambda v: round(int(v) / 18000) if v is not None else None  # noqa: E731
        out[f"{key}/box"] = (q(off.get("x")) if off is not None else None, q(off.get("y")) if off is not None else None,
                             q(ext.get("cx")) if ext is not None else None, q(ext.get("cy")) if ext is not None else None)
        out[f"{key}/turn"] = (xfrm.get("rot", "0"), xfrm.get("flipH", "0") in ("1", "true"), xfrm.get("flipV", "0") in ("1", "true"))
    if sppr is not None:
        g = sppr.find("a:prstGeom", NS)
        out[f"{key}/geom"] = g.get("prst") if g is not None else ("custGeom" if sppr.find("a:custGeom", NS) is not None else None)
        out[f"{key}/fill"] = _fill(sppr)
        ln = sppr.find("a:ln", NS)
        if ln is not None:
            dash = ln.find("a:prstDash", NS)
            out[f"{key}/ln"] = (ln.get("w"), _fill(ln), dash.get("val") if dash is not None else None)
    body = sp.find("p:txBody", NS)
    if body is not None:
        bp = body.find("a:bodyPr", NS)
        out[f"{key}/bodyPr"] = _attrs(bp)
        for i, para in enumerate(body.findall("a:p", NS)):
            ppr = para.find("a:pPr", NS)
            # algn="l" is the default: written or not, the paragraph is left-aligned.
            out[f"{key}/p{i}/pPr"] = ",".join(x for x in _attrs(ppr).split(",") if x and x != "algn=l")
            if ppr is not None:
                for ch in ppr:
                    n = ch.tag.split("}")[1]
                    if n in ("lnSpc", "spcBef", "spcAft"):
                        v = ch.find("a:spcPct", NS) if ch.find("a:spcPct", NS) is not None else ch.find("a:spcPts", NS)
                        out[f"{key}/p{i}/{n}"] = _attrs(v)
                    elif n.startswith("bu"):
                        out[f"{key}/p{i}/bullet"] = f"{n}:{_attrs(ch)}"
            text = "".join((t.text or "") for t in para.iter(f"{A}t"))
            out[f"{key}/p{i}/text"] = text
            for j, r in enumerate(para.findall("a:r", NS)):
                rpr = r.find("a:rPr", NS)
                out[f"{key}/p{i}/r{j}/rPr"] = _attrs(rpr)
                if rpr is not None:
                    out[f"{key}/p{i}/r{j}/color"] = _color(rpr.find("a:solidFill", NS))
                    lat = rpr.find("a:latin", NS)
                    out[f"{key}/p{i}/r{j}/font"] = lat.get("typeface") if lat is not None else None
                    hl = rpr.find("a:highlight", NS)
                    out[f"{key}/p{i}/r{j}/highlight"] = _color(hl)
                    ef = rpr.find("a:effectLst", NS)
                    out[f"{key}/p{i}/r{j}/effects"] = ",".join(c.tag.split("}")[1] for c in ef) if ef is not None else None
    tbl = sp.find(".//a:tbl", NS)
    if tbl is not None:
        out[f"{key}/table"] = (len(tbl.findall("a:tr", NS)), len(tbl.findall("a:tblGrid/a:gridCol", NS)))
    for ch in sp.findall("p:sp", NS) + sp.findall("p:pic", NS) if tag == "grpSp" else []:
        cid = ch.find(".//p:cNvPr", NS).get("id")
        shape_facts(ch, f"{key}/{cid}", out)


def facts(data: bytes) -> dict:
    z = zipfile.ZipFile(io.BytesIO(data))
    names = set(z.namelist())
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    out: dict = {}
    sz = pres.find("p:sldSz", NS)
    out["size"] = (sz.get("cx"), sz.get("cy"))
    slides = [rels[s.get(f"{{{NS['r']}}}id")] for s in pres.findall("p:sldIdLst/p:sldId", NS)]
    out["slides"] = len(slides)
    for n, target in enumerate(slides, 1):
        path = "ppt/" + target.lstrip("/").replace("ppt/", "") if not target.startswith("/") else target.lstrip("/")
        if path not in names:
            continue
        sl = ET.fromstring(z.read(path))
        srels_p = path.replace("slides/", "slides/_rels/") + ".rels"
        srels = {r.get("Id"): (r.get("Type"), r.get("Target")) for r in ET.fromstring(z.read(srels_p))} if srels_p in names else {}
        lay = next((t for ty, t in srels.values() if ty.endswith("/slideLayout")), None)
        if lay:
            lp = "ppt/slideLayouts/" + lay.split("/")[-1]
            if lp in names:
                out[f"s{n}/layout"] = ET.fromstring(z.read(lp)).find("p:cSld", NS).get("name")
        tr = sl.find("p:transition", NS)
        if tr is None:
            # PowerPoint writes newer transitions inside mc:AlternateContent.
            for alt in sl.iter():
                if alt.tag.endswith("}transition"):
                    tr = alt
                    break
        if tr is not None:
            kids = [c.tag.split("}")[1] for c in tr]
            out[f"s{n}/transition"] = (kids[0] if kids else "", tr.get("spd"), tr.get(f"{{http://schemas.microsoft.com/office/powerpoint/2010/main}}dur") or tr.get("dur"))
        effects = [(c.get("presetClass"), c.get("presetID"), c.get("presetSubtype")) for c in sl.iter(f"{P}cTn") if c.get("presetID")]
        out[f"s{n}/anims"] = effects
        tree = sl.find("p:cSld/p:spTree", NS)
        order = []
        for sp in tree:
            nv = sp.find(".//p:cNvPr", NS)
            if nv is None:
                continue
            order.append(nv.get("id"))
            ph = sp.find(".//p:nvPr/p:ph", NS)
            # Placeholders by their type and index (PowerPoint and we number shapes alike only by chance).
            key = f"s{n}/ph:{ph.get('type', 'body')}:{ph.get('idx', '0')}" if ph is not None else f"s{n}/{sp.tag.split('}')[1]}#{len([o for o in order]) }"
            shape_facts(sp, key, out)
        out[f"s{n}/count"] = len(order)
    return out


def safe(name: str) -> str:
    return "".join(ch if ch.isalnum() or ch in " -_" else "_" for ch in name)


def delta(before: dict, after: dict) -> dict:
    return {k: after.get(k) for k in set(before) | set(after) if before.get(k) != after.get(k)}


# ----- PowerPoint ----------------------------------------------------------------------------------------------

class Ctx:
    def __init__(self, t, pres):
        self.t, self.pres = t, pres
        self.slide = pres.Slides(1)
        self.sh = None
        self.tr = None

    def layout(self, name):
        return self.t.layout(self.pres, name)


def powerpoint_side(t, before: Path, ctx: str, action, out: Path) -> dict:
    pres = t.open_copy(before, with_window=True)
    c = Ctx(t, pres)
    win = pres.Windows(1)
    win.Activate()
    win.ViewType = 9
    win.View.GotoSlide(1)
    time.sleep(0.3)
    sl = pres.Slides(1)
    shapes = [sl.Shapes(i) for i in range(1, sl.Shapes.Count + 1)]
    own = [s for s in shapes if not s.Type == 14]
    if ctx in ("shape-text", "placeholder"):
        c.sh = own[0] if ctx == "shape-text" else [s for s in shapes if s.Type == 14][1]
        c.tr = c.sh.TextFrame.TextRange
        c.tr.Select()
    elif ctx in ("shape", "picture", "picture-rotated"):
        c.sh = own[0]
        c.sh.Select()
    elif ctx == "two-shapes":
        sl.Shapes.Range([s.Name for s in own]).Select()
        c.sh = own[0]
    elif ctx in ("two-shapes-first", "two-shapes-second"):
        c.sh = own[0 if ctx == "two-shapes-first" else 1]
        c.sh.Select()
    elif ctx == "table":
        c.sh = own[0]
        c.sh.Table.Cell(1, 1).Shape.TextFrame.TextRange.Select()
    elif ctx == "slides":
        win.ViewType = 9
        try:
            win.Panes(1).Activate()
            sl.Select()
        except Exception:  # noqa: BLE001
            pass
    kind, what = action
    err = None
    try:
        if kind == "mso":
            t.app.CommandBars.ExecuteMso(what)
        else:
            what(c)
    except Exception as e:  # noqa: BLE001
        err = str(e)[:300]
    time.sleep(0.6)
    t.save_copy(pres, out)
    pres.Saved = -1
    pres.Close()
    t.own.remove(pres)
    return {"error": err}


# ----- RaabePresent --------------------------------------------------------------------------------------------

def our_side(sw: Sweeper, ctx: str, steps: list) -> tuple[bytes, bytes, dict]:
    p = sw.page
    real_ctx = {"two-shapes-first": "two-shapes", "two-shapes-second": "two-shapes", "picture-rotated": "picture"}.get(ctx, ctx)
    tab = steps[0]
    sw.reset(real_ctx, tab)
    if ctx in ("two-shapes-first", "two-shapes-second"):
        k = 0 if ctx == "two-shapes-first" else 1
        p.js(f"(() => {{ const ids = window.__present.state().shapes.filter((s) => !s.ph).map((s) => s.id); window.__present.app.editor.select([ids[{k}]]); }})()")
    if ctx == "picture-rotated":
        # A picture someone turned: Bild zurücksetzen brings it back.
        p.js("(() => { const a = window.__present.app; a.commands.run('arrange.rotateRight'); a.deck.clearHistory(); })()")
        time.sleep(0.2)
    before = base64.b64decode(p.js("window.__present.bytes()"))
    controls = p.js("window.__sw.controls()")
    name = steps[1]
    c = next((x for x in controls if x["name"] == name and x["kind"] not in ("split-arrow", "color-arrow")), None) \
        or next((x for x in controls if x["name"] == name), None) \
        or next((x for x in controls if x["name"].startswith(name)), None)
    info: dict = {}
    if not c:
        return before, before, {"error": f"control {name} not found on {tab}"}
    r = p.js(f"window.__sw.findControl({json.dumps(c['key'])})")
    rest = steps[2:]
    if rest and isinstance(rest[0], tuple) and rest[0][0] == "type":
        p.mouse(r["x"], r["y"])
        p.key("Strg+A")
        p.type(rest[0][1])
        p.key("Enter")
        rest = rest[1:]
    else:
        click_at(p, r)
    time.sleep(0.35)
    for s in rest:
        items = p.js("window.__sw.items()")
        # The last-opened popup's entries come last: look from the end.
        m = next((x for x in reversed(items) if x["name"] == s), None) or next((x for x in reversed(items) if x["name"].startswith(s)), None)
        if not m:
            info["error"] = f"entry {s} not found ({[x['name'] for x in items][:30]})"
            break
        ir = p.js(f"window.__sw.itemRect({json.dumps(m['path'])})")
        click_at(p, ir)
        time.sleep(0.4)
    errs = p.js("window.__sw.take()")
    if errs:
        info["errs"] = errs
    for _ in range(2):
        if p.js("window.__sw.overlays().length + window.__sw.dialogs().length"):
            p.key("Esc")
            time.sleep(0.1)
    after = base64.b64decode(p.js("window.__present.bytes()"))
    return before, after, info


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dev", action="store_true")
    ap.add_argument("--url", default="http://localhost:5186/")
    ap.add_argument("--port", type=int, default=9482)
    ap.add_argument("--only")
    args = ap.parse_args()
    from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget
    work = TMP / "office"
    work.mkdir(parents=True, exist_ok=True)
    cases = [c for c in CASES if not args.only or args.only.lower() in c[0].lower()]
    page = start(args.dev, args.url, args.port)
    sw = Sweeper(page, work)
    ours: list = []
    try:
        for name, ctx, steps, *_ in cases:
            b, a, info = our_side(sw, ctx, steps)
            fb, fa = work / f"{safe(name)}-before.pptx", work / f"{safe(name)}-ours.pptx"
            fb.write_bytes(b)
            fa.write_bytes(a)
            ours.append(info)
            print("ours", name, info or "", flush=True)
    finally:
        page.close()
    try:
        t = PowerPointTarget(visible=True)
    except OwnerPowerPointRunning as e:
        print("PowerPoint:", e)
        return
    results = []
    try:
        for (name, ctx, steps, action, *opt), info in zip(cases, ours):
            fb, fa, fp = work / f"{safe(name)}-before.pptx", work / f"{safe(name)}-ours.pptx", work / f"{safe(name)}-ppt.pptx"
            pinfo = powerpoint_side(t, fb, ctx, action, fp)
            before = facts(fb.read_bytes())
            fo, fpp = facts(fa.read_bytes()), facts(fp.read_bytes())
            if "kind" in opt:
                for f in (before, fo, fpp):
                    for k in [k for k in f if k.endswith("/transition") and f[k]]:
                        f[k] = f[k][0]
            d_ours, d_ppt = delta(before, fo), delta(before, fpp)
            diffs = {k: {"ours": d_ours.get(k, before.get(k)), "powerpoint": d_ppt.get(k, before.get(k))} for k in sorted(set(d_ours) | set(d_ppt))
                     if d_ours.get(k, before.get(k)) != d_ppt.get(k, before.get(k))}
            status = "error" if info.get("error") or info.get("errs") or pinfo["error"] else ("same" if not diffs else "different")
            if not d_ppt and not d_ours:
                status = "no change on either side" if status == "same" else status
            res = {"name": name, "ctx": ctx, "steps": [s if isinstance(s, str) else list(s) for s in steps], "action": [action[0], action[1] if isinstance(action[1], str) else "COM"],
                   "status": status, "ours": info, "powerpoint": pinfo, "changed_ours": sorted(d_ours), "changed_powerpoint": sorted(d_ppt), "differences": diffs}
            results.append(res)
            print(f"{status:10} {name}  {json.dumps(diffs, ensure_ascii=False, default=str)[:400]}", flush=True)
    finally:
        print(t.close(kill_if_unreachable=True))
    counts: dict = {}
    for r in results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    out = {"date": datetime.date.today().isoformat(), "target": "dev" if args.dev else "exe", "cases": len(results), "counts": counts, "results": results}
    p = HERE / "results" / f"{out['date']}-present-office-compare.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1, default=str), encoding="utf-8", newline="\n")
    print(json.dumps(counts), "->", p.name)


if __name__ == "__main__":
    main()
