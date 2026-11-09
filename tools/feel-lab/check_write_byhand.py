"""By hand in the real RaabeWrite exe (D-151): the options the by-hand audit (D-133) could not finish with one click.

- Bildformat > Zuschneiden: crop marks appear, a mark dragged inwards crops the picture, Enter ends; the file has the
  crop (a:srcRect); Strg+Z takes it back. Also the picture's context menu Zuschneiden.
- Formformat (and Einfügen) > Formen > Freihandform: Form, Kurve, Freihand: Skizze: points clicked on the page and a
  double click (a click on the first point closes the form), or one drag for the Skizze; the file has the custom
  geometry (a:custGeom: a:lnTo for the form, a:cubicBezTo for the curve); Strg+Z takes it back.
- Tabellenentwurf > Rahmen > Rahmenpipette: the pointer becomes the pipette, a click on a cell's border picks its line
  (the Stift-group shows it), then Rahmen übertragen paints it on another cell edge; Esc ends the painter.

Input is trusted DevTools mouse and key input into our own exe, off screen (no real mouse or keyboard, no lock).
Only test APIs that read are used to check (document fingerprint, the editor's mode); every action is a click, drag
or key. Each save goes to TEMP.

    .venv/Scripts/python.exe check_write_byhand.py      # results/<date>-write-byhand.json
"""
from __future__ import annotations

import json
import os
import re
import zipfile

os.environ.setdefault("SWEEP_PORT", "9486")
os.environ.setdefault("SWEEP_DIR", "raabe-sweep-byhand")

import sweep_write as SW  # noqa: E402

HERE = SW.HERE
CURSOR = "(() => { const w = document.querySelector('.workspace'); return w ? w.style.cursor : ''; })()"


def drag(p, x0, y0, x1, y1, steps=10) -> None:
    p.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x0, "y": y0, "button": "none"})
    p.cdp("Input.dispatchMouseEvent", {"type": "mousePressed", "x": x0, "y": y0, "button": "left", "clickCount": 1, "buttons": 1})
    for k in range(1, steps + 1):
        p.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x0 + (x1 - x0) * k / steps, "y": y0 + (y1 - y0) * k / steps, "button": "left", "buttons": 1})
    p.cdp("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": x1, "y": y1, "button": "left", "clickCount": 1, "buttons": 0})


def xml_of(path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


class Check:
    def __init__(self) -> None:
        self.a = SW.App()
        self.p = self.a.p
        self.rows: list[dict] = []

    def close(self) -> None:
        self.a.close()

    def record(self, row: dict) -> None:
        row["passed"] = all(v for k, v in row.items() if k.startswith("ok_"))
        self.rows.append(row)
        print(("PASS " if row["passed"] else "FAIL ") + json.dumps(row, ensure_ascii=False)[:400], flush=True)

    def undo_ok(self, base: str) -> bool:
        p = self.p
        p.js("__probe.focus()")
        p.key("z", ctrl=True)
        p.pump(0.3)
        return p.js("__sw.docFp()") == base

    def save(self, name: str) -> str:
        path = SW.TMP / f"byhand-{name}.docx"
        self.p.js(f"__app.saveTo({json.dumps(str(path))})")
        return str(path)

    # ----- Zuschneiden ---------------------------------------------------------------------
    def crop(self, via_context: bool) -> None:
        a, p = self.a, self.p
        a.fresh("picture")
        base = p.js("__sw.docFp()")
        r = p.js("__sw.pictureRect()")
        if via_context:
            p.click(r["x"], r["y"], button="right")
            p.pump(0.35)
            item = next((i for i in p.js("__sw.popupItems()") if i["label"] == "Zuschneiden"), None)
            if item:
                p.click(item["x"], item["y"])
        else:
            a.click_tab("Bildformat")
            c = next((c for c in p.js("__sw.controls()") if c["label"] == "Zuschneiden" and c["kind"] == "split-main"), None)
            if c:
                p.click(c["x"], c["y"])
        p.pump(0.3)
        cropping = p.js("!!__app.editor().isCropping")
        # The right crop mark (middle of the right edge) dragged a quarter of the width inwards, then Enter.
        r = p.js("__sw.pictureRect()")
        drag(p, r["x"] + r["w"] / 2, r["y"], r["x"] + r["w"] / 4, r["y"])
        p.pump(0.3)
        p.key("Enter")
        p.pump(0.3)
        ended = not p.js("!!__app.editor().isCropping")
        r2 = p.js("__sw.pictureRect()")
        changed = p.js("__sw.docFp()") != base
        path = self.save("crop-context" if via_context else "crop")
        m = re.search(r"<a:srcRect [^>]*/>", xml_of(path))
        self.record({"case": "Kontextmenü > Zuschneiden" if via_context else "Bildformat > Zuschneiden", "ok_marks": cropping, "ok_doc": changed,
                     "ok_narrower": bool(r2) and r2["w"] < r["w"] - 5, "ok_ended_by_enter": ended, "ok_file": bool(m and " r=" in m.group(0)),
                     "srcRect": m.group(0) if m else None, "width": [r["w"], r2 and r2["w"]], "ok_undo": self.undo_ok(base), "errors": p.take_errors()})

    # ----- Freihandform, Kurve, Skizze -------------------------------------------------------
    def freeform(self, label: str, closed: bool = False, tab: str = "Einfügen") -> None:
        a, p = self.a, self.p
        a.fresh("caret")
        p.js("__sw.selectText('Schluss.', true)")
        base = p.js("__sw.docFp()")
        n0 = xml_of(self.save("ff-base")).count("<a:custGeom")
        a.click_tab(tab)
        c = a.find_control("Formen")
        p.click(c["x"], c["y"])
        p.pump(0.3)
        item = next((i for i in p.js("__sw.popupItems()") if i["label"].startswith(label)), None)
        if not item:
            self.record({"case": f"Formen > {label}", "ok_found": False})
            return
        p.click(item["x"], item["y"])
        p.pump(0.2)
        armed = p.js(CURSOR)
        pg = p.js("(() => { const b = document.querySelector('.page').getBoundingClientRect(); return {x: b.left + 250, y: b.top + 560}; })()")
        x, y = pg["x"], pg["y"]
        pts = [(x, y), (x + 120, y + 10), (x + 140, y + 90), (x + 30, y + 110)]
        if label.startswith("Freihand: Skizze"):
            drag(p, x, y, x + 140, y + 90, steps=20)
        else:
            for i, (px, py) in enumerate(pts):
                p.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": px, "y": py, "button": "none"})
                last = i == len(pts) - 1 and not closed
                p.click(px, py, count=2 if last else 1)
                p.pump(0.08)
            if closed:
                p.click(x + 1, y + 1)  # a click on the first point closes the form
        p.pump(0.4)
        changed = p.js("__sw.docFp()") != base
        ctx = p.js("__sw.ui().ctx")
        xml = xml_of(self.save("ff-" + re.sub(r"\W+", "-", label)))
        geoms = re.findall(r"<a:custGeom>.*?</a:custGeom>", xml, re.S)
        g = geoms[-1] if len(geoms) > n0 else ""
        want = "a:cubicBezTo" if label.startswith("Kurve") else "a:lnTo"
        self.record({"case": f"{tab} > Formen > {label}" + (" (geschlossen)" if closed else ""), "ok_armed": armed not in ("", "auto"), "cursor": armed,
                     "ok_doc": changed, "ok_shape_tab": "Formformat" in (ctx or []), "ok_file": want in g,
                     "ok_closed": ("<a:close/>" in g) == closed if not label.startswith("Freihand: Skizze") else True,
                     "segments": g.count("<a:lnTo>") + g.count("<a:cubicBezTo>"), "ok_undo": self.undo_ok(base), "errors": p.take_errors()})

    # ----- Rahmenpipette ---------------------------------------------------------------------
    def pipette(self) -> None:
        a, p = self.a, self.p
        a.fresh("table")
        base = p.js("__sw.docFp()")
        a.click_tab("Tabellenentwurf")
        pen0 = p.js("__sw.ui().combos")
        c = next((c for c in p.js("__sw.controls()") if c["group"] == "Rahmen" and c["kind"] == "split-arrow"), None)
        p.click(c["x"], c["y"])
        p.pump(0.3)
        item = next((i for i in p.js("__sw.popupItems()") if i["label"] == "Rahmenpipette"), None)
        if not item:
            self.record({"case": "Rahmenpipette", "ok_found": False})
            return
        p.click(item["x"], item["y"])
        p.pump(0.2)
        cursor = p.js(CURSOR)
        sampling = p.js("!!__app.editor().borderSampling")
        # The table's outer top border: the top edge of the first cell, read from the cell's text box.
        e = p.js("(() => { const t = __probe.text(); const i = t.indexOf('Pflanze'); const a = __probe.pointAt(i); return {x: a.x + 20, y: a.y}; })()")
        top = p.js(f"(() => {{ let y = {e['y']}; for (let k = 0; k < 40; k++, y -= 0.5) {{ const ed = __app.editor().view.cellEdgeAt({e['x']}, y); if (ed && ed.side === 'top') return y; }} return null; }})()")
        p.click(e["x"], top if top is not None else e["y"] - 4)
        p.pump(0.3)
        picked = not p.js("!!__app.editor().borderSampling")
        painting = p.js("!!__app.editor().borderPainting")
        pen1 = p.js("__sw.ui().combos")
        # Rahmen übertragen is on after picking (Word): a click on an inner edge paints the picked line there.
        cell = p.js("(() => { const t = __probe.text(); const i = t.indexOf('Tomate'); const a = __probe.pointAt(i); return {x: a.x + 10, y: a.y}; })()")
        y2 = p.js(f"(() => {{ let y = {cell['y']}; for (let k = 0; k < 40; k++, y -= 0.5) {{ const ed = __app.editor().view.cellEdgeAt({cell['x']}, y); if (ed && ed.side === 'top') return y; }} return null; }})()")
        p.click(cell["x"], y2 if y2 is not None else cell["y"] - 4)
        p.pump(0.3)
        painted = p.js("__sw.docFp()") != base
        p.key("Escape")
        p.pump(0.2)
        ended = not p.js("!!__app.editor().borderPainting")
        self.record({"case": "Tabellenentwurf > Rahmen > Rahmenpipette", "ok_pipette": sampling, "cursor": cursor[:60],
                     "ok_picked": picked, "pen": [pen0, pen1], "ok_painter_on": painting, "ok_painted": painted, "ok_esc_ends": ended,
                     "ok_undo": self.undo_ok(base) if painted else False, "errors": p.take_errors()})


def main() -> None:
    c = Check()
    try:
        c.crop(False)
        c.crop(True)
        for label in ("Freihandform: Form", "Kurve", "Freihand: Skizze"):
            c.freeform(label)
        c.freeform("Freihandform: Form", closed=True)
        c.pipette()
    finally:
        c.close()
    out = HERE / "results" / f"{SW.DATE}-write-byhand.json"
    summary = {"cases": len(c.rows), "passed": sum(1 for r in c.rows if r["passed"])}
    out.write_text(json.dumps({"date": SW.DATE, "summary": summary, "rows": c.rows}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps(summary), "->", out.name)


if __name__ == "__main__":
    main()
