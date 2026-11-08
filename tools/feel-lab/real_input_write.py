"""By-hand audit of RaabeWrite, the real-input part (D-133): the controls where only the real mouse and keyboard show
whether they work (lists that open and close, galleries, palettes, spinners, the zoom slider, dialogs moved by their
title, focus moving with Tab, keytips with the real Alt key, the context menu at the mouse, the system's colour
dialog behind "Weitere Farben...").

Holds %TEMP%/raabe-real-input.lock for the whole run; every click or key is sent only after checking that the window
in front belongs to our own RaabeWrite process. Starts its own exe; the base document is sweep_write.py's.

    .venv/Scripts/python.exe real_input_write.py
"""
from __future__ import annotations

import datetime
import json
import time
import traceback
from pathlib import Path

from feellab import inputlock, targets, winapi
from feellab.sweep_js import HELPERS
from feellab.write_page import WritePage
import sweep_write as SW

HERE = Path(__file__).parent
RESULTS: list[dict] = []


class NotOurs(Exception):
    pass


class Real:
    def __init__(self) -> None:
        self.t = targets.raabe_write("short", port=9473, background=False)
        self.pid = self.t.proc.pid
        self.p = WritePage(self.t.ws)
        self.p.js(HELPERS)
        self.dpr = self.p.js("devicePixelRatio")

    def ours(self) -> None:
        fg = winapi.foreground()
        if winapi.window_pid(fg) != self.pid:
            winapi.activate(self.t.hwnd)
            time.sleep(0.2)
            fg = winapi.foreground()
            if winapi.window_pid(fg) != self.pid:
                raise NotOurs(f"window in front is not our RaabeWrite ({winapi.window_title(fg)[:40]!r})")

    def screen(self, x: float, y: float) -> tuple[float, float]:
        ox, oy = self.t._viewport_origin()
        return ox + x * self.dpr, oy + y * self.dpr

    def click(self, r: dict, count: int = 1, right: bool = False) -> None:
        self.ours()
        x, y = self.screen(r["x"], r["y"])
        if right:
            winapi.right_click(x, y)
        else:
            winapi.click(x, y, count)
        time.sleep(0.35)

    def move(self, r: dict) -> None:
        self.ours()
        winapi.move(*self.screen(r["x"], r["y"]))
        time.sleep(0.25)

    def press(self, combo: str) -> None:
        self.ours()
        winapi.press(combo)
        time.sleep(0.3)

    def type(self, text: str) -> None:
        self.ours()
        winapi.type_text(text)
        time.sleep(0.2)

    def fresh(self, setup: str = "text") -> None:
        p = self.p
        p.js("document.querySelectorAll('.dialog-layer, body > .dialog').forEach(l => l.remove())")
        p.js(f"__app.openPath({json.dumps(str(SW.BASE))})")
        for _ in range(100):
            if p.js("!!__app.editor() && __probe.text().includes('Schluss.')"):
                break
            time.sleep(0.03)
        p.js("__app.closeOthers()")
        p.js("__probe.focus()")
        if setup == "text":
            p.js("__sw.selectText('weitere Beobachtungen')")
        elif setup == "caret":
            p.js("__sw.selectText('Schluss.', true)")
        p.js("__sw.markBase()")
        p.take_errors()

    def tab(self, name: str) -> None:
        self.click(self.p.js(f"__sw.tabRect({json.dumps(name)})"))

    def control(self, label: str, kind: str | None = None, group: str | None = None) -> dict:
        for c in self.p.js("__sw.controls()"):
            if c["label"] == label and (kind is None or c["kind"] == kind) and (group is None or c["group"] == group):
                return c
        raise KeyError(label)

    def item(self, label: str, starts: bool = False) -> dict:
        for i in self.p.js("__sw.popupItems()"):
            if i["label"] == label or (starts and i["label"].startswith(label)):
                return i
        raise KeyError(label)

    def fmt(self, needle: str = "weitere Beobachtungen") -> dict:
        return self.p.js(f"__probe.formatAt(__probe.text().indexOf({json.dumps(needle)}) + 2)")

    def close(self) -> None:
        self.t.close()


def check(name: str, ok: bool, detail: str = "", errors: list | None = None) -> None:
    ok = bool(ok) and not errors
    RESULTS.append({"name": name, "ok": ok, "detail": detail, **({"errors": errors} if errors else {})})
    print(("PASS " if ok else "FAIL ") + name + (f"  [{detail}]" if detail and not ok else ""), flush=True)


def cases(r: Real) -> None:
    p = r.p
    # 1. Schriftgrad list: open with the arrow, pick 14, list closes.
    r.fresh()
    r.tab("Start")
    r.click(r.control("Öffnen", "combo-arrow", "Schriftart") if False else [c for c in p.js("__sw.controls()") if c["kind"] == "combo-arrow"][1])
    opened = p.js("__sw.ui().popups")
    r.click(r.item("14"))
    check("Schriftgrad list opens and 14 applies", opened == 1 and r.fmt()["size"] == 14 and p.js("__sw.ui().popups") == 0,
          f"opened {opened} size {r.fmt()['size']}", p.take_errors())

    # 2. The same list closes when its arrow is clicked again (Office).
    r.fresh()
    arrow = [c for c in p.js("__sw.controls()") if c["kind"] == "combo-arrow"][1]
    r.click(arrow)
    r.click(arrow)
    check("a list's arrow clicked again closes it", p.js("__sw.ui().popups") == 0, str(p.js("__sw.ui().popups")), p.take_errors())

    # 3. Formatvorlagen gallery in the ribbon: Überschrift 1.
    r.fresh()
    r.click(r.control("Überschrift 1", "gallery-item"))
    st = p.js("__app.docInfo().paras[2].style")
    check("gallery entry Überschrift 1 applies to the paragraph", st in ("heading 1", "Überschrift 1"), str(st), p.take_errors())

    # 4. Schriftfarbe palette: the arrow, then Rot (Standardfarben).
    r.fresh()
    arrows = [c for c in p.js("__sw.controls()") if c["kind"] == "color-arrow"]
    r.click(arrows[1])
    r.move(r.item("Rot"))
    r.click(r.item("Rot"))
    check("Schriftfarbe palette: Rot", r.fmt()["color"] == "FF0000", str(r.fmt()["color"]), p.take_errors())

    # 5. Weitere Farben...: the system's colour dialog must open (and close with Esc).
    r.fresh()
    arrows = [c for c in p.js("__sw.controls()") if c["kind"] == "color-arrow"]
    r.click(arrows[1])
    before = set(winapi.find_windows(lambda h: winapi.window_class(h) == "#32770"))
    r.click(r.item("Weitere Farben..."))
    time.sleep(0.8)
    new = [h for h in winapi.find_windows(lambda h: winapi.window_class(h) == "#32770") if h not in before]
    fg = winapi.foreground()
    title = winapi.window_title(new[0]) if new else ""
    check("Weitere Farben... opens a colour dialog", bool(new), f"dialogs {[winapi.window_title(h) for h in new]}")
    if new and fg in new:
        winapi.press("esc")
        time.sleep(0.4)
    RESULTS[-1]["detail"] = f"title {title!r}"

    # 6. Layout > Einzug Links: the spinner's up arrow twice.
    r.fresh()
    r.tab("Layout")
    ups = [c for c in p.js("__sw.controls()") if c["kind"] == "spinner-step" and c["group"] == "Absatz"]
    if ups:
        r.click(ups[0])
        r.click(ups[0])
    ind = r.fmt()["indLeft"]
    check("Einzug Links spinner: two steps up", ind > 0, f"indLeft {ind} steps {[u['label'] for u in ups[:4]]}", p.take_errors())

    # 7. Spinner field: typed value and Enter.
    r.fresh()
    r.tab("Layout")
    field = next(c for c in p.js("__sw.controls()") if c["kind"] == "spinner-input" and c["group"] == "Absatz")
    r.click(field, count=3)
    r.type("2 cm")
    r.press("enter")
    ind = r.fmt()["indLeft"]
    check("Einzug Links: typing 2 cm and Enter", abs(ind - 2 * 28.3465 * 4 / 3) < 1.5 or abs(ind - 56.69) < 1, f"indLeft {ind}", p.take_errors())

    # 8. Absatz dialog: Tab moves through the fields in order, a typed value applies with Enter.
    r.fresh()
    r.tab("Start")
    r.click(r.control("Absatz...", "launcher"))
    seq = []
    for _ in range(4):
        seq.append(p.js("(() => { const a = document.activeElement; return a ? (a.tagName + ':' + (a.getAttribute('aria-label') || a.previousElementSibling?.textContent || a.textContent || '').trim().slice(0, 30)) : ''; })()"))
        r.press("tab")
    inside = p.js("!!document.activeElement.closest('.dialog')")
    check("Absatz dialog: Tab keeps the focus inside and moves on", inside and len(set(seq)) == 4, json.dumps(seq, ensure_ascii=False), p.take_errors())
    r.press("esc")

    # 9. A dialog is moved by its title bar.
    r.fresh()
    r.click(r.control("Schriftart...", "launcher"))
    t0 = p.js("(() => { const d = document.querySelector('.dialog-layer .dialog'); const b = d.getBoundingClientRect(); const t = d.querySelector('.dlg-title').getBoundingClientRect(); return {x: t.left + 60, y: t.top + t.height / 2, left: b.left, top: b.top}; })()")
    r.ours()
    x1, y1 = r.screen(t0["x"], t0["y"])
    winapi.drag(x1, y1, x1 + 150, y1 + 90, steps=15)
    time.sleep(0.3)
    t1 = p.js("(() => { const b = document.querySelector('.dialog-layer .dialog').getBoundingClientRect(); return {left: b.left, top: b.top}; })()")
    moved = abs(t1["left"] - t0["left"] - 100) < 15 and abs(t1["top"] - t0["top"] - 60) < 15
    check("dialog follows a drag of its title bar", moved, f"{t0['left']:.0f},{t0['top']:.0f} -> {t1['left']:.0f},{t1['top']:.0f}", p.take_errors())
    r.press("esc")

    # 10. Tabelle einfügen: the grid under the real mouse, 3 x 2.
    r.fresh("caret")
    r.tab("Einfügen")
    r.click(r.control("Tabelle", "dropdown"))
    cells = [i for i in p.js("__sw.popupItems()") if i["w"] < 30 and i["h"] < 30]
    if len(cells) >= 12:
        # Grid rows of 10: the third cell of the second row.
        target = cells[10 + 2]
        r.move(cells[0])
        r.move(target)
        r.click(target)
    tables = p.js("__app.docInfo().paras.filter(p => p.cell).map(p => p.cell.join(','))")
    check("Tabelle einfügen grid: 3 x 2 by hovering and clicking", len(cells) >= 12 and "1,2" in tables, f"cells {len(cells)} table cells {tables[-8:]}", p.take_errors())

    # 11. Context menu at the mouse (real right click), Schriftart... from it.
    r.fresh()
    pt = p.js("(() => { const t = __probe.text(); const i = t.indexOf('Beobachtungen'); const a = __probe.pointAt(i + 3); return {x: a.x, y: a.y + a.h / 2}; })()")
    r.click(pt, right=True)
    menu = p.js("(() => { const m = [...document.body.children].filter(c => c.classList.contains('menu-root')).pop(); if (!m) return null; const b = m.querySelector('.menu').getBoundingClientRect(); return {left: b.left, top: b.top}; })()")
    near = menu and abs(menu["left"] - pt["x"]) < 40
    check("context menu opens at the mouse", bool(near), str(menu), p.take_errors())
    r.press("esc")

    # 12. Keytips with the real Alt key: Alt, R, 1 = Fett.
    r.fresh()
    r.press("alt")
    r.press("r")
    r.press("1")
    check("keytips Alt R 1 (Fett)", r.fmt()["bold"] is True, str(r.fmt()["bold"]), p.take_errors())
    r.press("esc")

    # 13. Zoom slider in the status bar: drag the thumb to the right.
    r.fresh()
    z0 = p.js("__app.editor().zoom")
    th = p.js("(() => { const t = document.querySelector('.statusbar .zoom input, .statusbar .zoom .thumb, .statusbar .zoom [role=slider]'); if (!t) return null; const b = t.getBoundingClientRect(); return {x: b.left + b.width / 2, y: b.top + b.height / 2, w: b.width}; })()")
    if th:
        r.ours()
        x1, y1 = r.screen(th["x"], th["y"])
        winapi.drag(x1, y1, x1 + 60, y1, steps=12)
        time.sleep(0.4)
    z1 = p.js("__app.editor().zoom")
    check("zoom slider follows a drag", th is not None and z1 > z0, f"{z0} -> {z1}", p.take_errors())

    # 14. Mouse wheel scrolls the document.
    r.fresh()
    s0 = p.js("document.querySelector('.docarea').scrollTop")
    r.ours()
    da = p.js("(() => { const b = document.querySelector('.docarea').getBoundingClientRect(); return {x: b.left + b.width / 2, y: b.top + 200}; })()")
    x, y = r.screen(da["x"], da["y"])
    winapi.move(x, y)
    winapi._send(winapi.INPUT(type=winapi.INPUT_MOUSE, u=winapi._INPUTUNION(mi=winapi.MOUSEINPUT(dx=0, dy=0, mouseData=(-360) & 0xFFFFFFFF, dwFlags=0x0800))))
    time.sleep(0.5)
    s1 = p.js("document.querySelector('.docarea').scrollTop")
    check("mouse wheel scrolls the document", s1 > s0, f"{s0} -> {s1}", p.take_errors())

    # 15. Schriftart dialog: pick a size from its list with the mouse, OK.
    r.fresh()
    r.tab("Start")
    r.click(r.control("Schriftart...", "launcher"))
    opt = p.js("(() => { const d = document.querySelector('.dialog-layer .dialog'); const el = [...d.querySelectorAll('option, li, .list button, [role=option]')].find(e => e.textContent.trim() === '16'); if (!el) return null; el.scrollIntoView({block: 'nearest'}); const b = el.getBoundingClientRect(); return b.width ? {x: b.left + b.width / 2, y: b.top + b.height / 2} : null; })()")
    if opt:
        r.click(opt)
        r.press("enter")
    check("Schriftart dialog: size 16 from its list, Enter", opt is not None and r.fmt()["size"] == 16, f"opt {opt} size {r.fmt()['size']}", p.take_errors())
    r.press("esc")


def main() -> None:
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-write-real-input.json"
    with inputlock.real_input("audit-write real input", wait=1800):
        r = Real()
        try:
            cases(r)
        except NotOurs as e:
            check("run stopped", False, str(e))
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            check("run crashed", False, str(e)[:300])
        finally:
            r.close()
    passed = sum(1 for x in RESULTS if x["ok"])
    out.write_text(json.dumps({"passed": passed, "total": len(RESULTS), "cases": RESULTS}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(f"{passed}/{len(RESULTS)} ->", out.name)


if __name__ == "__main__":
    main()
