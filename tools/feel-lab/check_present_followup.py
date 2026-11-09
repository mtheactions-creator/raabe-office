"""By-hand check of what the follow-up of the RaabePresent audit built (D-152 to D-154), against PowerPoint's recording
(results/2026-11-09-powerpoint-followup.json, -powerpoint-smartart-format.json). Everything goes through the real
controls of the release exe as a student uses them: trusted CDP clicks at the control's place, typing into fields,
Enter (sweep_present.py's page and helpers; invisible, no real-input lock needed).

- Übergänge > Dauer on a slide without a transition: shows 02,00; 3 typed + Enter writes the duration alone.
- Tabellenlayout > Seitenverhältnis sperren: ticked, then Höhe 8 cm typed: Breite follows (as PowerPoint, 32 cm).
- Folienmaster: Masterlayout... on the master (Datum off, then on again at Office's place), Platzhalter einfügen's
  entries on a layout (each picked and a click on the layout: the placeholder type and its 2,54 cm square; Inhalt drawn
  by dragging), Folienmaster einfügen (a second master with eleven layouts). Each with Strg+Z afterwards.
- SmartArt > Format: the tab's groups, which controls are greyed, Fülleffekt > a colour writes the background.

    .venv/Scripts/python.exe check_present_followup.py

Result: results/<date>-present-followup-check.json (passed/total).
"""
from __future__ import annotations

import datetime
import json
import re
import time
from pathlib import Path

from sweep_present import start

HERE = Path(__file__).parent
REC = json.loads((HERE / "results" / "2026-11-09-powerpoint-followup.json").read_text(encoding="utf-8"))


class Check:
    def __init__(self):
        self.page = start(False, "", 9541)
        self.out: list[dict] = []

    def js(self, e):
        return self.page.js(e)

    def click(self, r: dict, count: int = 1) -> None:
        self.page.mouse(r["x"], r["y"], count=count)
        time.sleep(0.3)

    def key(self, k: str) -> None:
        self.page.key(k)
        time.sleep(0.15)

    def type(self, text: str) -> None:
        self.page.type(text)
        time.sleep(0.15)

    def setup(self, ctx: str, tab: str | None) -> None:
        for _ in range(2):
            self.key("Esc")
        self.js(f"window.__present.sweepContext({json.dumps(ctx)})")
        time.sleep(0.4)
        if tab:
            r = self.js(f"window.__sw.tabRect({json.dumps(tab)})")
            if not r:
                raise RuntimeError(f"no tab {tab}: {self.js('window.__sw.tabs()')}")
            self.click(r)

    def control(self, name: str, kind: str | None = None) -> dict:
        cs = self.js("window.__sw.controls()")
        c = next((x for x in cs if x["name"] == name and (kind is None or x["kind"] == kind)), None)
        if not c:
            raise RuntimeError(f"no control {name} ({kind}); have {[x['name'] for x in cs][:80]}")
        r = self.js(f"window.__sw.findControl({json.dumps(c['key'])})")
        r["disabled"] = c["disabled"]
        return r

    def item(self, name: str) -> dict:
        it = next((x for x in reversed(self.js("window.__sw.items()")) if x["name"] == name), None)
        if not it:
            raise RuntimeError(f"no entry {name}; have {[x['name'] for x in self.js('window.__sw.items()')][:60]}")
        return self.js(f"window.__sw.itemRect({json.dumps(it['path'])})")

    def part(self, path_js: str) -> str:
        return self.js(f"(() => {{ const a = window.__present.app; const p = {path_js}; return p && a.deck.pptx.has(p) ? new XMLSerializer().serializeToString(a.deck.pptx.doc(p)) : ''; }})()")

    def slide_xml(self) -> str:
        return self.part("a.slidePath()")

    def case(self, name: str, fn) -> None:
        try:
            ok, got = fn()
        except Exception as e:  # noqa: BLE001
            ok, got = False, f"error: {str(e)[:300]}"
        errs = self.js("window.__sw.take()")
        self.out.append({"name": name, "ok": bool(ok) and not errs, "got": got, "errs": errs})
        print("ok  " if ok and not errs else "FAIL", name, json.dumps(got, ensure_ascii=False)[:300], errs, flush=True)


def spin_input(c: Check, name: str, group: str = "") -> dict:
    # (Tabellenlayout has two "Höhe:" boxes: the row's under Zellengröße, the table's under Tabellengröße.)
    scope = f".ribbon-card .rgroup[aria-label='{group}']" if group else ".ribbon-card"
    return c.js(f"(() => {{ const i = Array.from(document.querySelectorAll({json.dumps(scope + ' .rspinner input')})).find((x) => (x.getAttribute('aria-label') || '').replace(/:$/, '') === {json.dumps(name)}); return i ? Object.assign(window.__sw.rect(i), {{ value: i.value }}) : null; }})()")


def thumb(c: Check, i: int) -> dict:
    return c.js(f"(() => {{ const e = document.querySelectorAll('.thumb-item')[{i}]; e.scrollIntoView({{ block: 'nearest' }}); return window.__sw.rect(e); }})()")


def cases(c: Check) -> None:
    rec = REC["dauer"]

    def dauer():
        c.setup("slides", "Übergänge")
        c.click(thumb(c, 2))
        before = spin_input(c, "Dauer")
        if not before:
            raise RuntimeError("no Dauer box")
        c.click(before, 2)
        c.key("Strg+a")
        c.type("3")
        c.key("Enter")
        time.sleep(0.5)
        xml = c.slide_xml()
        after = spin_input(c, "Dauer")
        tr = re.findall(r"<p:transition[^>]*/>|<p:transition[^>]*>.*?</p:transition>", xml, re.S)
        ok = before["value"] == rec["before"]["shown"] and after["value"] == rec["typed_3"]["shown"] and any('p14:dur="3000"' in t and t.endswith("/>") for t in tr)
        c.key("Esc")
        c.js("window.__present.app.focusMain()")
        c.key("Strg+z")
        undone = "transition" not in c.slide_xml()
        return ok and undone, {"shown_before": before["value"], "shown_after": after["value"], "transition": tr, "undone": undone, "powerpoint": rec["typed_3"]["xml"]}
    c.case("Übergänge > Dauer ohne Übergang: 3 getippt", dauer)

    def dauer_arrow():
        c.setup("slides", "Übergänge")
        c.click(thumb(c, 1))
        c.click(c.control("Dauer: (Mehr)", "spinner-up"))
        time.sleep(0.4)
        v = spin_input(c, "Dauer")["value"]
        return v == rec["slide1_up"]["shown"] and 'p14:dur="2250"' in c.slide_xml(), {"shown": v, "powerpoint": rec["slide1_up"]["shown"]}
    c.case("Übergänge > Dauer ohne Übergang: Pfeil nach oben", dauer_arrow)

    ra = REC["aspect"]

    def aspect():
        c.setup("table", "Tabellenlayout")
        c.click(c.control("Seitenverhältnis sperren"))
        st0 = c.js("(() => { const v = window.__present.app.editor.selectedViews.find((x) => x.frame === 'table'); return { w: v.w, h: v.h }; })()")
        hb = spin_input(c, "Höhe", "Tabellengröße")
        c.click(hb, 2)
        c.key("Strg+a")
        c.type("8")
        c.key("Enter")
        time.sleep(0.5)
        st1 = c.js("(() => { const v = window.__present.app.editor.selectedViews.find((x) => x.frame === 'table'); return { w: v.w, h: v.h }; })()")
        ratio0, ratio1 = st0["w"] / st0["h"], st1["w"] / st1["h"]
        pp = ra["height_8cm"]["w"] / ra["height_8cm"]["h"]
        return abs(ratio1 - ratio0) < 0.01 and abs(st1["h"] - 8 * 360000) < 2000, {"before": st0, "after": st1, "ratio": round(ratio1, 3), "powerpoint_ratio": round(pp, 3)}
    c.case("Tabellenlayout > Seitenverhältnis sperren, Höhe 8 cm", aspect)

    rm = REC["master"]

    def master_layout():
        c.setup("master", "Folienmaster")
        c.click(thumb(c, 0))
        ml = c.control("Masterlayout...")
        pe = c.control("Platzhalter einfügen")
        states = {"Masterlayout...": not ml["disabled"], "Platzhalter einfügen": not pe["disabled"]}
        c.click(ml)
        info = c.js("window.__sw.dialogInfo()")
        boxes = [x["name"] for x in info[-1]["controls"] if x["type"] == "checkbox"] if info else []
        # Datum off, OK.
        r = c.js("(() => { const d = window.__sw.dialogs().pop(); const l = Array.from(d.querySelectorAll('label')).find((x) => x.textContent.trim() === 'Datum'); return window.__sw.rect(l.querySelector('input')); })()")
        c.click(r)
        c.click(c.js("window.__sw.dialogButton('ok')"))
        time.sleep(0.4)
        gone = 'type="dt"' not in c.part("a.deck.pptx.masters()[0]")
        # Again: Datum on, OK: back at Office's place.
        c.click(c.control("Masterlayout..."))
        r = c.js("(() => { const d = window.__sw.dialogs().pop(); const l = Array.from(d.querySelectorAll('label')).find((x) => x.textContent.trim() === 'Datum'); return Object.assign(window.__sw.rect(l.querySelector('input')), { checked: l.querySelector('input').checked }); })()")
        unchecked = not r["checked"]
        c.click(r)
        c.click(c.js("window.__sw.dialogButton('ok')"))
        time.sleep(0.4)
        mx = c.part("a.deck.pptx.masters()[0]")
        dt = [x for x in re.findall(r"<p:sp>.*?</p:sp>", mx, re.S) if 'type="dt"' in x]
        back = bool(dt) and '<a:off x="838200" y="6356350"/><a:ext cx="2743200" cy="365125"/>' in dt[0]
        want = {x["name"]: x["enabled"] for x in rm["master_selected_controls"]}
        same = all(states[k] == want[k] for k in states)
        return same and boxes == [b["name"] for b in rm["dialog"]["boxes"]] and gone and unchecked and back, {"enabled": states, "powerpoint": {k: want[k] for k in states}, "boxes": boxes, "removed": gone, "unchecked_after": unchecked, "restored_at_office_place": back}
    c.case("Folienmaster > Masterlayout... (Datum aus, wieder ein)", master_layout)

    def placeholders():
        c.setup("master", "Folienmaster")
        # The layout Leer (the master is entry 0; Leer is the seventh layout).
        c.click(thumb(c, 7))
        pe = c.control("Platzhalter einfügen (Pfeil)", "split-arrow")
        if pe["disabled"]:
            return False, "Platzhalter einfügen greyed on a layout"
        c.click(pe)
        names = [x["name"] for x in c.js("window.__sw.items()")]
        c.key("Esc")
        got = {}
        want = rm["placeholder_menu"]
        for k, name in enumerate(want):
            c.click(c.control("Platzhalter einfügen (Pfeil)", "split-arrow"))
            c.click(c.item(name))
            a = c.js("window.__present.slidePointToClient(200, 150)")
            if k == 0:
                b = c.js("window.__present.slidePointToClient(400, 300)")
                c.page.cdp("Input.dispatchMouseEvent", {"type": "mousePressed", "x": a["x"], "y": a["y"], "button": "left", "buttons": 1, "clickCount": 1})
                for i in range(1, 11):
                    c.page.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": a["x"] + (b["x"] - a["x"]) * i / 10, "y": a["y"] + (b["y"] - a["y"]) * i / 10, "button": "left", "buttons": 1})
                c.page.cdp("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": b["x"], "y": b["y"], "button": "left", "buttons": 0, "clickCount": 1})
            else:
                c.click(a)
            time.sleep(0.3)
            xml = c.slide_xml()
            sps = re.findall(r"<p:sp>.*?</p:sp>", xml, re.S)
            last = sps[-1] if sps else ""
            ph = re.search(r"<p:ph([^>]*)", last)
            ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', last)
            pp = rm["placeholder_picks"][name]
            pxml = pp.get("xml", "")
            pph = re.search(r"<p:ph([^>]*)", pxml)
            same_ph = bool(ph and pph) and re.sub(r'idx="\d+"', "", ph.group(1)).strip(" /") == re.sub(r'idx="\d+"', "", pph.group(1)).strip(" /")
            size_ok = k == 0 or (ext is not None and ext.group(1) == "914400" and ext.group(2) == "914400")
            cameo = name != "Cameo" or "p232:cameo" in last
            got[name] = {"ph": ph.group(1).strip() if ph else None, "powerpoint": pph.group(1).strip() if pph else None, "size": ext.groups() if ext else None, "ok": same_ph and size_ok and cameo}
            c.js("window.__present.app.focusMain()")
            c.key("Strg+z")
        return names[:len(want)] == want and all(v["ok"] for v in got.values()) and "Platzhalter" not in "".join(re.findall(r'name="([^"]*)"', c.slide_xml())), {"menu": names, "powerpoint_menu": want, "picks": got}
    c.case("Folienmaster > Platzhalter einfügen: jeder Eintrag", placeholders)

    def insert_master():
        c.setup("master", "Folienmaster")
        c.click(thumb(c, 0))
        n0 = c.js("window.__present.app.deck.pptx.masters().length")
        c.click(c.control("Folienmaster einfügen"))
        time.sleep(0.5)
        st = c.js("(() => { const a = window.__present.app, p = a.deck.pptx, ms = p.masters(); const m = ms[ms.length - 1]; const th = p.themeOf(m); return { masters: ms.length, layouts: p.layouts(m).length, theme: th ? p.doc(th).documentElement.getAttribute('name') : null, preserve: p.doc(m).documentElement.getAttribute('preserve'), thumbs: document.querySelectorAll('.thumb-item').length, masterThumbs: document.querySelectorAll('.thumb-item.master').length }; })()")
        saved = c.js("(() => { const a = window.__present.app; try { const b = a.deck.write(); return b.length > 0; } catch (e) { return String(e); } })()")
        c.js("window.__present.app.focusMain()")
        c.key("Strg+z")
        back = c.js("window.__present.app.deck.pptx.masters().length")
        r = rm["insert_master"]
        return (st["masters"] == n0 + 1 and st["layouts"] == len(r["layouts"]) and st["theme"] == r["design_name"] and st["preserve"] == "1" and st["masterThumbs"] == 2 and saved is True and back == n0,
                {**st, "undone_to": back, "powerpoint": {"designs": r["designs_after"], "layouts": len(r["layouts"]), "name": r["design_name"]}})
    c.case("Folienmaster > Folienmaster einfügen", insert_master)

    def smartart_format():
        c.setup("smartart", "Format")
        groups = c.js("Array.from(document.querySelectorAll('.ribbon-card .rgroup')).map((g) => g.getAttribute('aria-label'))")
        cs = c.js("window.__sw.controls()")
        greyed = {x["name"]: x["disabled"] for x in cs}
        want_grey = ["In 2D bearbeiten", "Form ändern", "Größer", "Kleiner", "Schnellformatvorlagen"]
        grey_ok = all(greyed.get(n, False) for n in want_grey[:4])
        c.click(c.control("Fülleffekt (Pfeil)", "color-arrow"))
        c.click(c.item("Rot"))
        time.sleep(0.4)
        data = c.js("(() => { const a = window.__present.app; const p = Object.keys(a.deck.pptx.files).concat([...a.deck.pptx.docs.keys()]).find((k) => /diagrams\\/data\\d+\\.xml$/.test(k)); return p ? new XMLSerializer().serializeToString(a.deck.pptx.doc(p)) : ''; })()")
        bg = re.findall(r"<dgm:bg>.*?</dgm:bg>", data, re.S)
        drawn = c.js("!!document.querySelector('.diagram-bg')")
        want_groups = ["Formen", "Formenarten", "WordArt-Formate", "Barrierefreiheit", "Anordnen", "Größe"]
        return groups == want_groups and grey_ok and bool(bg) and 'val="FF0000"' in bg[0] and drawn, {"groups": groups, "greyed": {n: greyed.get(n) for n in want_grey}, "bg": bg, "drawn": drawn}
    c.case("SmartArt > Format: Gruppen, grau, Fülleffekt > Rot", smartart_format)

    def diagram_part(kind: str) -> str:
        return c.js(f"(() => {{ const a = window.__present.app; const p = Object.keys(a.deck.pptx.files).concat([...a.deck.pptx.docs.keys()]).find((k) => /diagrams\\/{kind}\\d+\\.xml$/.test(k)); return p ? new XMLSerializer().serializeToString(a.deck.pptx.doc(p)) : ''; }})()")

    def smartart_text():
        c.setup("smartart", "Format")
        c.click(c.control("Textfüllung (Pfeil)", "color-arrow"))
        c.click(c.item("Rot"))
        time.sleep(0.4)
        runs = re.findall(r"<a:r>.*?</a:r>", diagram_part("drawing"), re.S)
        red = [r for r in runs if 'val="FF0000"' in r]
        data_red = len(re.findall(r'<a:rPr[^>]*>(?:(?!</a:rPr>).)*FF0000', diagram_part("data"), re.S))
        c.js("window.__present.app.focusMain()")
        c.key("Strg+z")
        back = not re.findall(r'FF0000', diagram_part("drawing"))
        return len(runs) == 3 and len(red) == 3 and data_red == 3 and back, {"runs": len(runs), "red_in_drawing": len(red), "red_in_data": data_red, "undone": back}
    c.case("SmartArt > Format > Textfüllung > Rot (ganze Grafik)", smartart_text)

    def smartart_outline():
        c.setup("smartart", "Format")
        c.click(c.control("Formkontur (Pfeil)", "color-arrow"))
        c.click(c.item("Grün"))
        time.sleep(0.4)
        whole = re.findall(r"<dgm:whole>.*?</dgm:whole>", diagram_part("data"), re.S)
        drawn = c.js("(() => { const e = document.querySelector('.diagram-bg'); return e ? e.style.boxShadow : ''; })()")
        return bool(whole) and "<a:ln" in whole[0] and "solidFill" in whole[0] and bool(drawn), {"whole": whole, "drawn": drawn}
    c.case("SmartArt > Format > Formkontur > Grün (ganze Grafik)", smartart_outline)


def main() -> None:
    c = Check()
    try:
        cases(c)
    finally:
        c.page.close()
    out = {"date": datetime.date.today().isoformat(), "passed": sum(1 for x in c.out if x["ok"]), "total": len(c.out), "cases": c.out}
    p = HERE / "results" / f"{out['date']}-present-followup-check.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print(out["passed"], "/", out["total"], "->", p.name)


if __name__ == "__main__":
    main()
