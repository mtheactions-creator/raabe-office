"""By-hand audit of RaabeWrite (D-133): every option of the real app triggered as a student does, checked and compared
with Word.

    .venv/Scripts/python.exe sweep_write.py inventory          # results/<date>-write-options.json
    .venv/Scripts/python.exe sweep_write.py sweep [--tabs start,insert] [--gallery N] [--out NAME]
    .venv/Scripts/python.exe sweep_write.py word [--in NAME]     # Word opens every saved file, compares with Word's own option

`inventory` lists every option: the ribbon definition (all tabs, contextual ones too, with menus) from the app's
command list, and what the running exe really shows (each tab's visible controls, each menu, gallery and dialog opened
by a click), so options defined but unreachable and reachable but undefined both show.

`sweep` starts the real release exe (our own instance, off screen, never focused) and for every visible control of
every tab (each contextual tab with a document set up for it: caret in a table, picture selected, header open,
shape inserted, equation open) clicks it through trusted DevTools input at the control's position, then every entry
of the menu, gallery or palette it opens, and each dialog's OK and Abbrechen. Per option it checks: no exception or
console error; a visible effect (document, selection, pane, dialog, view, ribbon state); Strg+Z restores the document
where the option changed it; the document saves. `word` then has Word open every saved file (it must open) and, for
the options with a Word counterpart (WORD_OPS), applies Word's own option to the same base document and compares
what changed (feellab/word_summary.py).

Input goes into the page only (no real mouse or keyboard, no lock needed); the real-input pass is real_input_write.py.
"""
from __future__ import annotations

import argparse
import datetime
import json
import os
import sys
import tempfile
import time
import traceback
from pathlib import Path

from feellab import targets, winapi
from feellab import word_summary as ws
from feellab.sweep_js import HELPERS
from feellab.write_page import WritePage

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir()) / os.environ.get("SWEEP_DIR", "raabe-sweep-write")
BASE = TMP / "base.docx"
DATE = datetime.date.today().isoformat()
PORT = int(os.environ.get("SWEEP_PORT", "9471"))

# Which document set-up each tab needs.
TAB_SETUP = {"Tabellenentwurf": "table", "Layout (Tabelle)": "table", "Kopf- und Fußzeile": "headerFooter",
             "Bildformat": "picture", "Formformat": "shape", "Formel": "math"}
CONTEXT_TABS = {"table": ["Tabellenentwurf", "Layout"], "headerFooter": ["Kopf- und Fußzeile"], "picture": ["Bildformat"],
                "shape": ["Formformat"], "math": ["Formel"]}
# The parts of the UI a sweep step may leave changed (view, rulers, reduced ribbon ...).
VIEW_JS = "(() => { const u = __sw.ui(); return JSON.stringify([u.view, u.shell, u.minimized, u.bodyCls]); })()"
# Closes the open panes (Navigation, Formatvorlagen, Kommentare ...) by their own close buttons.
CLOSE_PANES = """(() => { let n = 0; for (const p of document.querySelectorAll('.navpane, aside, .taskpane, [role=complementary], .tpane, .clippane, .fxpane')) {
  if (p.hidden || !p.getBoundingClientRect().width) continue;
  const b = p.querySelector('.tp-close, .nav-close, button[aria-label="Schließen"], button[data-tip="Schließen"]'); if (b) { b.click(); n++; } } return n; })()"""
# Commands deliberately left out (docs/SCOPE.md: later / out of scope) are greyed with "Zurückgestellt" in the tip.
DEFERRED_MARK = "Zurückgestellt"


def log(*a) -> None:
    print(*a, flush=True)


class App:
    """Our own RaabeWrite instance, off screen and never focused."""

    def __init__(self) -> None:
        TMP.mkdir(exist_ok=True)
        (TMP / "saved").mkdir(exist_ok=True)
        if not BASE.exists():
            make_base()
        self.t = targets.raabe_write("short", port=PORT, background=True)
        winapi.place_quietly(self.t.hwnd, -4000, 0, 2600, 1700)
        self.p = WritePage(self.t.ws)
        self.p.js(HELPERS)
        self.n_saved = 0
        self.setup_name = ""
        self.base_fp = ""
        self.base_view = ""

    def close(self) -> None:
        self.t.close()

    # ----- document set-ups ---------------------------------------------------------------
    def fresh(self, setup: str = "text") -> None:
        p = self.p
        p.key("Escape")
        p.js("document.querySelectorAll('.dialog-layer, body > .dialog').forEach(l => l.remove())")
        # The same document again, unchanged, with nothing open: no need to load it anew.
        quick = p.js("(() => { const u = __sw.ui(); return !u.popups && !u.dialogs.length && !u.panes.length && !u.backstage && u.docs === 1; })()")
        if not (quick and self.base_fp and p.js("__sw.docFp()") == self.base_fp and p.js(VIEW_JS) == self.base_view):
            # A new empty document first, so the base document is read from its file again (openPath only
            # switches to a document that is already open).
            p.js(CLOSE_PANES)
            p.js("__app.newDoc()")
            p.js("__app.closeOthers()")
            p.js(f"__app.openPath({json.dumps(str(BASE))})")
            for _ in range(150):
                if p.js("__app.docs().length === 1 && !!__app.editor() && __probe.text().includes('Schluss.')"):
                    break
                time.sleep(0.03)
            p.js("__app.closeOthers()")
            p.pump(0.15)
            if not self.base_fp:
                self.base_fp = p.js("__sw.docFp()")
                self.base_view = p.js(VIEW_JS)
        self.setup(setup)
        p.js("__sw.markBase()")
        p.take_errors()

    def setup(self, setup: str) -> None:
        p = self.p
        self.setup_name = setup
        p.js("__probe.focus()")
        if setup == "text":
            p.js("__sw.selectText('weitere Beobachtungen')")
        elif setup == "caret":
            p.js("__sw.selectText('weitere Beobachtungen', true)")
        elif setup == "table":
            p.js("__sw.selectText('Sonnenblume\\n', true)")
            p.pump(0.1)
        elif setup == "picture":
            r = p.js("__sw.pictureRect()")
            if r:
                p.click(r["x"], r["y"])
        elif setup == "headerFooter":
            # Double click into the top margin of page 1 (Word opens the header).
            r = p.js("(() => { const pg = document.querySelector('.page'); const b = pg.getBoundingClientRect(); return {x: b.left + b.width / 2, y: b.top + 30}; })()")
            p.click(r["x"], r["y"], count=2)
        elif setup == "shape":
            p.js("__sw.selectText('Schluss.', true)")
            self.click_tab("Einfügen")
            c = self.find_control("Formen")
            if c:
                p.click(c["x"], c["y"])
                p.pump(0.2)
                items = p.js("__sw.popupItems()")
                rect = next((i for i in items if i["label"].startswith("Rechteck")), None)
                if rect:
                    p.click(rect["x"], rect["y"])
                    pg = p.js("(() => { const b = document.querySelector('.page').getBoundingClientRect(); return {x: b.left + 200, y: b.top + 400}; })()")
                    p.cdp("Input.dispatchMouseEvent", {"type": "mousePressed", "x": pg["x"], "y": pg["y"], "button": "left", "clickCount": 1, "buttons": 1})
                    for k in range(1, 11):
                        p.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": pg["x"] + 12 * k, "y": pg["y"] + 8 * k, "button": "left", "buttons": 1})
                    p.cdp("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": pg["x"] + 120, "y": pg["y"] + 80, "button": "left", "clickCount": 1, "buttons": 0})
        elif setup == "math":
            p.js("__sw.selectText('Schluss.', true)")
            p.key(")", ctrl=False, shift=True, alt=True, vk=48, code="Digit0")  # Alt+Umschalt+0
        p.pump(0.25)

    # ----- ribbon -------------------------------------------------------------------------
    def click_tab(self, name: str) -> bool:
        r = self.p.js(f"__sw.tabRect({json.dumps(name)})")
        if not r:
            return False
        if self.p.js("document.querySelector('.tabrow .tab.active')?.textContent.trim()") != name:
            self.p.click(r["x"], r["y"])
            self.p.pump(0.1)
        return True

    def find_control(self, label: str, cid: str | None = None):
        for c in self.p.js("__sw.controls()"):
            if (cid and c["id"] == cid) or (not cid and c["label"] == label):
                return c
        return None

    def state(self) -> dict:
        return {"doc": self.p.js("__sw.docFp()"), "doc2": self.p.js("__sw.docFp(true)"), "ui": self.p.js("__sw.ui()")}

    def save(self, name: str) -> str | None:
        self.n_saved += 1
        path = TMP / "saved" / f"{self.n_saved:04d}.docx"
        path.parent.mkdir(exist_ok=True)
        try:
            ok = self.p.js(f"__app.saveTo({json.dumps(str(path))})")
        except RuntimeError as e:
            return f"ERROR {e}"[:200]
        return path.name if ok and path.exists() else None


def make_base() -> None:
    TMP.mkdir(exist_ok=True)
    pic = ws.png(TMP / "bild.png")
    app = ws.word_app()
    try:
        ws.build_base(app, BASE, pic)
    finally:
        app.Quit(0)


# The buttons, fields and lists inside the backstage page shown.
BS_BUTTONS = """(() => { const seen = new Map(); return [...document.querySelectorAll('.backstage .page-body button, .backstage .page-body select, .backstage .page-body input')]
  .filter(b => { const r = b.getBoundingClientRect(); return r.width > 0 && r.height > 0 && r.top < innerHeight; })
  .map(b => { const lab = ((b.getAttribute('aria-label') || b.textContent || b.title || b.tagName).trim().replace(/\\s+/g, ' ')).slice(0, 60);
    const n = (seen.get(lab) || 0) + 1; seen.set(lab, n); const r = b.getBoundingClientRect();
    return { id: lab + '#' + n, label: lab + (n > 1 ? ' #' + n : ''), x: r.left + r.width / 2, y: r.top + r.height / 2 }; }); })()"""


def effects(before: dict, after: dict) -> list[str]:
    out = []
    if after["doc"] != before["doc"]:
        out.append("doc")
    a, b = before["ui"], after["ui"]
    for k in ("tab", "ctx", "checked", "combos", "backstage", "status", "view", "shell", "panes", "dialogs", "popups", "sel", "docs", "bodyCls", "minimized", "native"):
        if a.get(k) != b.get(k):
            out.append(k)
    return out


class Sweep:
    def __init__(self, app: App, gallery: int | None, saved_index: dict) -> None:
        self.a = app
        self.gallery = gallery
        self.rows: list[dict] = []
        self.saved = saved_index

    def record(self, row: dict) -> None:
        self.rows.append(row)
        st = row["status"]
        log(f"{st:9} {row['tab']} > {row['path']}  {row.get('effects', '')} {row.get('note', '')[:140]}")

    def open_tab(self, tab: str, setup: str) -> bool:
        self.a.fresh(setup)
        return self.a.click_tab(tab)

    def judge(self, tab: str, path: str, setup: str, before: dict, after: dict, errors: list, extra: dict | None = None,
              undo: bool = True) -> dict:
        eff = effects(before, after)
        row = {"tab": tab, "path": path, "setup": setup, "effects": eff, **(extra or {})}
        if errors:
            row["errors"] = errors
        status = "ok"
        if errors:
            status = "error"
        elif not eff:
            status = "noeffect"
        if "doc" in eff and undo:
            # Strg+Z as a key, then the document must be as before.
            p = self.a.p
            row["saved"] = self.a.save(path)
            if after["ui"]["dialogs"] or after["ui"]["popups"]:
                p.key("Escape")
                p.pump(0.1)
            p.js("__probe.focus()")
            p.key("z", ctrl=True)
            p.pump(0.2)
            fp = p.js("__sw.docFp()")
            row["undo"] = fp == before["doc"]
            if not row["undo"] and p.js("__sw.docFp(true)") == before.get("doc2"):
                # Only the style sheet differs: a style used once stays defined (unused and invisible; Word leaves it out of the file).
                row["undo"] = True
                row["note"] = "undo: text restored; the unused style definition stays in the file (Word: left out)"

            if not row["undo"]:
                # Some options take more than one step in Word too; a second Strg+Z must get there.
                p.key("z", ctrl=True)
                p.pump(0.2)
                row["undo2"] = p.js("__sw.docFp()") == before["doc"]
                if not row["undo2"] and status == "ok":
                    status = "undo"
            errs = p.take_errors()
            if errs:
                row.setdefault("errors", []).extend(errs)
                status = "error"
        row["status"] = status
        return row

    # ----- one control ---------------------------------------------------------------------
    def control(self, tab: str, setup: str, c: dict) -> None:
        a, p = self.a, self.a.p
        path = f"{c['group']} / {c['label']} [{c['kind']}]"
        if c["disabled"]:
            st = "deferred" if DEFERRED_MARK in c.get("tip", "") else "disabled"
            self.record({"tab": tab, "path": path, "setup": setup, "status": st, "tip": c.get("tip", "")[:120]})
            return
        if not self.open_tab(tab, setup):
            self.record({"tab": tab, "path": path, "setup": setup, "status": "error", "note": "tab not shown"})
            return
        cc = a.find_control(c["label"], c["id"])
        if not cc:
            self.record({"tab": tab, "path": path, "setup": setup, "status": "error", "note": "control not found after reset"})
            return
        before = a.state()
        if c["kind"] in ("combo-input", "spinner-input"):
            self.combo_input(tab, setup, path, cc, before)
            return
        p.click(cc["x"], cc["y"])
        p.pump(0.35)
        after = a.state()
        errors = p.take_errors()
        ui = after["ui"]
        if ui["popups"] > before["ui"]["popups"]:
            items = p.js("__sw.popupItems()")
            row = self.judge(tab, path, setup, before, after, errors, {"kind": "popup", "items": [i["label"] for i in items][:400], "n": len(items)}, undo=False)
            # Opening a menu changes nothing in the document.
            if "doc" in row["effects"]:
                row["status"] = "docchange-on-open"
            self.record(row)
            p.key("Escape")
            self.popup_entries(tab, setup, c, path, items)
        elif ui["dialogs"] and ui["dialogs"] != before["ui"]["dialogs"]:
            info = p.js("__sw.dialogInfo()")
            self.record(self.judge(tab, path, setup, before, after, errors, {"kind": "dialog", "dialog": info}, undo=False))
            self.dialog_buttons(tab, setup, path, [("control", c)], info)
        else:
            row = self.judge(tab, path, setup, before, after, errors)
            if row["status"] == "noeffect" and c["label"] == "Kopieren":
                info = p.js("__app.clipboardInfo()")
                row["status"], row["note"] = ("ok", f"clipboard {info}") if info and info.get("text") else ("noeffect", f"clipboard {info}")
            self.record(row)

    def combo_input(self, tab: str, setup: str, path: str, cc: dict, before: dict) -> None:
        """Types a value into a ribbon field (Schriftgrad, Einzug ...) and presses Enter."""
        p = self.a.p
        value = {"Schriftart": "Arial", "Schriftgrad": "14"}.get(cc["label"], "1")
        if "Formatvorlage" in cc["label"]:
            value = "APA"
        p.click(cc["x"], cc["y"])
        p.key("a", ctrl=True)
        p.type(value)
        p.key("Enter")
        p.pump(0.3)
        after = self.a.state()
        self.record(self.judge(tab, path + f" = {value}", setup, before, after, p.take_errors()))

    def reopen(self, tab: str, setup: str, chain: list) -> bool:
        """Opens the menu chain again from a fresh document: [("control", c), ("item", i), ...]."""
        a, p = self.a, self.a.p
        if not self.open_tab(tab, setup):
            return False
        for kind, x in chain:
            if kind == "control":
                cc = a.find_control(x["label"], x["id"])
                if not cc:
                    return False
                p.click(cc["x"], cc["y"])
            else:
                items = p.js("__sw.popupItems()")
                it = next((i for i in items if i["id"] == x["id"]), None)
                if not it:
                    return False
                p.click(it["x"], it["y"])
            p.pump(0.3)
        return True

    def pick(self, items: list[dict]) -> list[dict]:
        if self.gallery is None or len(items) <= self.gallery:
            return items
        # Large galleries: the first entries, the last ones and the commands below the tiles.
        k = self.gallery
        tiles = [i for i in items if i["tag"] == "BUTTON" and not i["sub"] and i["label"] and len(i["label"]) < 3]
        return items[:k] + items[max(k, len(items) - 3):]

    def popup_entries(self, tab: str, setup: str, c: dict, path: str, items: list[dict], chain: list | None = None, depth: int = 0) -> None:
        a, p = self.a, self.a.p
        chain = chain or [("control", c)]
        for it in self.pick(items):
            ipath = f"{path} > {it['label'] or '(ohne Namen)'}"
            if it["disabled"]:
                self.record({"tab": tab, "path": ipath, "setup": setup, "status": "deferred" if DEFERRED_MARK in it.get("tip", "") else "disabled"})
                continue
            if it["tag"] in ("INPUT", "SELECT"):
                self.record({"tab": tab, "path": ipath, "setup": setup, "status": "field", "note": "field inside a panel (typed in the real-input pass)"})
                continue
            if not self.reopen(tab, setup, chain):
                self.record({"tab": tab, "path": ipath, "setup": setup, "status": "error", "note": "menu did not open again"})
                continue
            cur = p.js("__sw.popupItems()")
            hit = next((i for i in cur if i["id"] == it["id"]), None)
            if not hit and p.js(f"__sw.reveal({json.dumps(it['label'])})"):
                p.pump(0.1)
                cur = p.js("__sw.popupItems()")
                hit = next((i for i in cur if i["label"] == it["label"]), None)
            if not hit:
                self.record({"tab": tab, "path": ipath, "setup": setup, "status": "error", "note": "entry missing on reopen"})
                p.key("Escape")
                continue
            before = a.state()
            n_pop = before["ui"]["popups"]
            p.move(hit["x"], hit["y"])
            p.pump(0.15)
            p.click(hit["x"], hit["y"])
            p.pump(0.4)
            after = a.state()
            errors = p.take_errors()
            if it["sub"] and depth < 2:
                sub = [i for i in p.js("__sw.popupItems()") if i["level"] >= 1 or i["id"] not in {x["id"] for x in cur}]
                if sub:
                    self.record({"tab": tab, "path": ipath, "setup": setup, "status": "ok" if not errors else "error", "kind": "submenu",
                                 "items": [s["label"] for s in sub][:200], "errors": errors})
                    p.key("Escape")
                    self.popup_entries(tab, setup, c, ipath, sub, chain + [("item", it)], depth + 1)
                    continue
            ui = after["ui"]
            if ui["dialogs"] and ui["dialogs"] != before["ui"]["dialogs"]:
                info = p.js("__sw.dialogInfo()")
                self.record(self.judge(tab, ipath, setup, before, after, errors, {"kind": "dialog", "dialog": info}, undo=False))
                self.dialog_buttons(tab, setup, ipath, chain + [("item", it)], info)
                continue
            if ("Formen" in path or "Textfeld erstellen" in ipath) and not ("doc" in effects(before, after)) and not after["ui"]["dialogs"]:
                pg = p.js("(() => { const b = document.querySelector('.page').getBoundingClientRect(); return {x: b.left + 300, y: b.top + 520}; })()")
                p.cdp("Input.dispatchMouseEvent", {"type": "mousePressed", "x": pg["x"], "y": pg["y"], "button": "left", "clickCount": 1, "buttons": 1})
                for k in range(1, 9):
                    p.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": pg["x"] + 15 * k, "y": pg["y"] + 10 * k, "button": "left", "buttons": 1})
                p.cdp("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": pg["x"] + 120, "y": pg["y"] + 80, "button": "left", "clickCount": 1, "buttons": 0})
                p.pump(0.4)
                after = a.state()
                errors += p.take_errors()
            row = self.judge(tab, ipath, setup, before, after, errors)
            # A click that only leaves the same menu open did nothing a student sees.
            if row["effects"] == ["popups"] and ui["popups"] < n_pop:
                row["status"] = "noeffect"
            self.record(row)
            p.key("Escape")

    def dialog_buttons(self, tab: str, setup: str, path: str, chain: list, info: dict | None) -> None:
        """Abbrechen (Esc) must leave the document as it was; OK with the dialog's own values must not fail."""
        a, p = self.a, self.a.p
        if not info:
            return
        # Escape.
        before_doc = None
        if self.reopen(tab, setup, chain):
            before_doc = p.js("__sw.docFp()")
            p.key("Escape")
            p.pump(0.25)
            after = a.state()
            errs = p.take_errors()
            ok = not after["ui"]["dialogs"] and after["doc"] == before_doc and not errs
            self.record({"tab": tab, "path": path + " > Esc", "setup": setup, "status": "ok" if ok else ("error" if errs else "esc"),
                         "errors": errs, "note": "" if ok else f"dialog left: {after['ui']['dialogs']} doc same: {after['doc'] == before_doc}"})
        prim = next((x for x in info["controls"] if x["tag"] == "button" and x["primary"] and not x["disabled"]), None)
        if not prim:
            prim = next((x for x in info["controls"] if x["tag"] == "button" and x["label"] in ("OK", "Einfügen", "Schließen")), None)
        if not prim:
            return
        if not self.reopen(tab, setup, chain):
            return
        before = a.state()
        cur = p.js("__sw.dialogInfo()")
        b = next((x for x in (cur or {}).get("controls", []) if x["tag"] == "button" and x["label"] == prim["label"]), None)
        if not b:
            return
        p.click(b["x"], b["y"])
        p.pump(0.5)
        after = a.state()
        errs = p.take_errors()
        row = self.judge(tab, path + f" > {prim['label']}", setup, before, after, errs)
        if row["status"] == "noeffect" and not after["ui"]["dialogs"]:
            row["status"] = "ok"  # OK with unchanged values closes the dialog (its own effect was seen above)
            row["note"] = "closed, nothing changed"
        if after["ui"]["dialogs"] and row["status"] == "ok" and "doc" not in row["effects"]:
            row["note"] = f"dialog still open: {after['ui']['dialogs']}"
        self.record(row)
        p.key("Escape")
        p.key("Escape")

    # ----- all tabs ----------------------------------------------------------------------
    def tab(self, tab: str, setup: str, only_groups: set[str] | None = None) -> None:
        if not self.open_tab(tab, setup):
            self.record({"tab": tab, "path": "(tab)", "setup": setup, "status": "error", "note": "tab not shown"})
            return
        controls = self.a.p.js("__sw.controls()")
        log(f"--- {tab}: {len(controls)} controls")
        for c in controls:
            if only_groups and c["groupId"] not in only_groups:
                continue
            try:
                self.control(tab, setup, c)
            except Exception as e:  # noqa: BLE001
                self.record({"tab": tab, "path": f"{c['group']} / {c['label']}", "setup": setup, "status": "crash", "note": f"{e}"[:300]})
                traceback.print_exc()
                try:
                    self.a.p.key("Escape")
                except Exception:  # noqa: BLE001
                    pass


    # ----- context menus, status bar, backstage, shortcuts -------------------------------
    def context_menus(self) -> None:
        """Right click on text, a list item, a table cell and the picture; every entry clicked."""
        a, p = self.a, self.a.p
        spots = {"text": ("text", "weitere Beobachtungen"), "list": ("caret", "Tomaten"), "table": ("table", "Sonnenblume"), "picture": ("picture", None)}
        for name, (setup, needle) in spots.items():
            def point():
                if needle is None:
                    return p.js("__sw.pictureRect()")
                return p.js(f"(() => {{ const t = __probe.text(); const i = t.indexOf({json.dumps(needle)}); const a = __probe.pointAt(i + 2); return {{x: a.x, y: a.y + a.h / 2}}; }})()")

            def opener(setup=setup, point=point):
                a.fresh(setup if setup != "caret" else "text")
                if needle and setup == "caret":
                    p.js(f"__sw.selectText({json.dumps(needle)}, true)")
                pt = point()
                p.click(pt["x"], pt["y"], button="right")
                p.pump(0.35)
            opener()
            items = p.js("__sw.popupItems()")
            self.record({"tab": "Kontextmenü", "path": name, "setup": setup, "status": "ok" if items else "noeffect", "kind": "popup",
                         "items": [i["label"] for i in items], "errors": p.take_errors()})
            p.key("Escape")
            self.menu_by_opener("Kontextmenü", name, setup, opener, items)

    def menu_by_opener(self, tab: str, path: str, setup: str, opener, items: list[dict], depth: int = 0, via: list | None = None) -> None:
        a, p = self.a, self.a.p
        via = via or []
        for it in self.pick(items):
            ipath = f"{path} > {it['label'] or '(ohne Namen)'}"
            if it["disabled"]:
                self.record({"tab": tab, "path": ipath, "setup": setup, "status": "deferred" if DEFERRED_MARK in it.get("tip", "") else "disabled"})
                continue
            if it["tag"] in ("INPUT", "SELECT"):
                self.record({"tab": tab, "path": ipath, "setup": setup, "status": "field"})
                continue
            opener()
            ok = True
            for v in via:
                cur = p.js("__sw.popupItems()")
                h = next((i for i in cur if i["id"] == v["id"]), None)
                if not h:
                    ok = False
                    break
                p.click(h["x"], h["y"])
                p.pump(0.3)
            cur = p.js("__sw.popupItems()") if ok else []
            hit = next((i for i in cur if i["id"] == it["id"]), None)
            if not hit:
                self.record({"tab": tab, "path": ipath, "setup": setup, "status": "error", "note": "entry missing on reopen"})
                p.key("Escape")
                continue
            before = a.state()
            p.move(hit["x"], hit["y"])
            p.pump(0.15)
            p.click(hit["x"], hit["y"])
            p.pump(0.4)
            after = a.state()
            errors = p.take_errors()
            if it["sub"] and depth < 2:
                sub = [i for i in p.js("__sw.popupItems()") if i["level"] >= 1]
                if sub:
                    self.record({"tab": tab, "path": ipath, "setup": setup, "status": "error" if errors else "ok", "kind": "submenu", "items": [s["label"] for s in sub], "errors": errors})
                    p.key("Escape")
                    self.menu_by_opener(tab, ipath, setup, opener, sub, depth + 1, via + [it])
                    continue
            ui = after["ui"]
            if ui["dialogs"] and ui["dialogs"] != before["ui"]["dialogs"]:
                info = p.js("__sw.dialogInfo()")
                row = self.judge(tab, ipath, setup, before, after, errors, {"kind": "dialog", "dialog": info}, undo=False)
                self.record(row)
                p.key("Escape")
                p.pump(0.2)
                left = p.js("__sw.dialogs()")
                if left:
                    self.record({"tab": tab, "path": ipath + " > Esc", "setup": setup, "status": "esc", "note": f"dialog left: {left}"})
                continue
            self.record(self.judge(tab, ipath, setup, before, after, errors))
            p.key("Escape")

    def statusbar(self) -> None:
        a, p = self.a, self.a.p
        a.fresh("text")
        cells = p.js("[...document.querySelectorAll('.statusbar button, .statusbar input')].filter(b => b.getBoundingClientRect().width > 0).map((b, i) => ({i, label: (b.getAttribute('aria-label') || (b.dataset.tip || '').split('\\n')[0] || b.textContent).trim(), r: (() => { const r = b.getBoundingClientRect(); return {x: r.left + r.width / 2, y: r.top + r.height / 2}; })()}))")
        for c in cells:
            a.fresh("text")
            before = a.state()
            p.click(c["r"]["x"], c["r"]["y"])
            p.pump(0.35)
            after = a.state()
            errors = p.take_errors()
            extra = {}
            if after["ui"]["dialogs"]:
                extra = {"kind": "dialog", "dialog": p.js("__sw.dialogInfo()")}
            self.record(self.judge("Statusleiste", f"{c['label']} #{c['i']}", "text", before, after, errors, extra, undo=False))
            p.key("Escape")

    def backstage(self) -> None:
        """Datei: every entry of the backstage's list. Speichern unter, Öffnen > Durchsuchen, Drucken and Exportieren
        open the system's own dialogs or the printer and are left to the real-input pass."""
        a, p = self.a, self.a.p
        native: set[str] = set()  # the system's dialogs are stubbed (__app.stubNative): asking for one is the effect

        def open_bs():
            a.fresh("text")
            # A copy of the document, so Speichern writes there and not to the base document.
            p.js(f"__app.saveTo({json.dumps(str(TMP / 'backstage.docx'))})")
            r = p.js("__sw.tabRect('Datei')")
            p.click(r["x"], r["y"])
            p.pump(0.4)
        open_bs()
        nav = p.js("[...document.querySelectorAll('.backstage nav button, .backstage .bs-nav button, .backstage [role=tab]')].filter(b => b.getBoundingClientRect().width > 0).map(b => b.textContent.trim())")
        self.record({"tab": "Datei", "path": "(Backstage)", "setup": "text", "status": "ok" if nav else "noeffect", "items": nav, "errors": p.take_errors()})
        for name in nav:
            if name in native:
                self.record({"tab": "Datei", "path": name, "setup": "text", "status": "native", "note": "system dialog or printer: real-input pass"})
                continue
            open_bs()
            before = a.state()
            r = p.js(f"(() => {{ const b = [...document.querySelectorAll('.backstage nav button, .backstage .bs-nav button, .backstage [role=tab]')].find(b => b.textContent.trim() === {json.dumps(name)}); if (!b) return null; const r = b.getBoundingClientRect(); return {{x: r.left + r.width / 2, y: r.top + r.height / 2}}; }})()")
            if not r:
                continue
            p.click(r["x"], r["y"])
            p.pump(0.5)
            after = a.state()
            errors = p.take_errors()
            page = p.js("(() => { const m = document.querySelector('.backstage main, .backstage .bs-page, .backstage .page-area'); return m ? m.textContent.trim().slice(0, 160) : ''; })()")
            extra = {"page": page}
            if after["ui"]["dialogs"]:
                extra["dialog"] = p.js("__sw.dialogInfo()")
            row = self.judge("Datei", name, "text", before, after, errors, extra, undo=False)
            self.record(row)
            # The page's own buttons (templates, Drucken, PDF, the Informationen page ...), each from a fresh page.
            btns = p.js(BS_BUTTONS) if after["ui"]["backstage"] else []
            p.key("Escape")
            p.key("Escape")
            for b in btns[:40]:
                open_bs()
                p.js(f"(() => {{ const b = [...document.querySelectorAll('.backstage nav button')].find(b => b.textContent.trim() === {json.dumps(name)}); b && b.click(); }})()")
                p.pump(0.3)
                cur = next((x for x in p.js(BS_BUTTONS) if x["id"] == b["id"]), None)
                if not cur:
                    self.record({"tab": "Datei", "path": f"{name} > {b['label']}", "setup": "text", "status": "error", "note": "button missing on reopen"})
                    continue
                before = a.state()
                p.click(cur["x"], cur["y"])
                p.pump(0.5)
                after = a.state()
                errors = p.take_errors()
                extra = {}
                if after["ui"]["dialogs"] and after["ui"]["dialogs"] != before["ui"]["dialogs"]:
                    extra["dialog"] = p.js("__sw.dialogInfo()")
                self.record(self.judge("Datei", f"{name} > {b['label']}", "text", before, after, errors, extra, undo=False))
                p.key("Escape")
                p.key("Escape")

    def shortcuts(self, commands: list[dict]) -> None:
        """Every keyboard shortcut of the command list, pressed as keys with the target selected."""
        a, p = self.a, self.a.p
        skip = {"Strg+F1"}  # reduces the ribbon (checked with the ribbon's own button)
        names = {"Umschalt": "shift", "Strg": "ctrl", "Alt": "alt"}
        vks = {">": (226, "IntlBackslash", True), "<": (226, "IntlBackslash", False), "#": (191, "Backslash", False), "+": (187, "BracketRight", False),
               "-": (189, "Slash", False), "Leertaste": (32, "Space", False), "Eingabe": (13, "Enter", False), "Rücktaste": (8, "Backspace", False),
               "Entf": (46, "Delete", False), "Tab": (9, "Tab", False), "Pos1": (36, "Home", False), "Ende": (35, "End", False)}
        for c in commands:
            sc = c.get("shortcut")
            if not sc or c.get("deferred"):
                continue
            for combo in [s.strip() for s in sc.split(",")][:1]:
                if combo in skip:
                    self.record({"tab": "Tastenkombination", "path": f"{combo} ({c['label']})", "setup": "text", "status": "native"})
                    continue
                parts = combo.split("+") if combo != "Strg++" else ["Strg", "+"]
                if combo.endswith("++"):
                    parts = combo[:-2].split("+") + ["+"]
                mods = {names[x] for x in parts[:-1] if x in names}
                k = parts[-1]
                a.fresh("table" if c["id"].startswith("table.") else "text")
                before = a.state()
                if k in vks:
                    vk, code, shift = vks[k]
                    p.key(k if len(k) == 1 else code, ctrl="ctrl" in mods, shift="shift" in mods or shift, alt="alt" in mods, vk=vk, code=code)
                elif len(k) == 1:
                    p.key(k.lower() if k.isalpha() else k, ctrl="ctrl" in mods, shift="shift" in mods, alt="alt" in mods)
                else:
                    p.key(k, ctrl="ctrl" in mods, shift="shift" in mods, alt="alt" in mods)
                p.pump(0.35)
                after = a.state()
                errors = p.take_errors()
                extra = {"cmd": c["id"]}
                if after["ui"]["dialogs"]:
                    extra["dialog"] = (p.js("__sw.dialogInfo()") or {}).get("title")
                self.record(self.judge("Tastenkombination", f"{combo} ({c['label']})", "text", before, after, errors, extra))
                p.key("Escape")


def run_sweep(args) -> None:
    out = HERE / "results" / f"{DATE}-{args.out}.json"
    app = App()
    rows: list[dict] = []
    try:
        sw = Sweep(app, args.gallery, {})
        # Our save of the untouched base document: the reference for what each option changed.
        app.fresh("text")
        app.p.js(f"__app.saveTo({json.dumps(str(TMP / 'saved' / 'base-ours.docx'))})")
        tabs = [t["label"] for t in app.p.js("__sw.tabs()") if not t["contextual"] and t["label"] != "Datei"]
        plan: list[tuple[str, str]] = [(t, "text") for t in tabs]
        for setup, names in CONTEXT_TABS.items():
            plan += [(n, setup) for n in names]
        want = set(args.tabs.split(",")) if args.tabs else None
        for tab, setup in plan:
            if want and tab not in want and f"{tab}@{setup}" not in want:
                continue
            sw.tab(tab, setup)
        extras = set(args.extra.split(",")) if args.extra else ({"context", "status", "backstage", "keys"} if not want else set())
        if "context" in extras:
            sw.context_menus()
        if "status" in extras:
            sw.statusbar()
        if "backstage" in extras:
            sw.backstage()
        if "keys" in extras:
            sw.shortcuts(app.p.js("__app.inventory().commands"))
        rows = sw.rows
    finally:
        app.close()
        summary = {}
        for r in rows:
            summary[r["status"]] = summary.get(r["status"], 0) + 1
        out.write_text(json.dumps({"date": DATE, "summary": summary, "rows": rows}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        log(json.dumps(summary), "->", out.name)


def run_inventory(args) -> None:
    """Every option: the app's ribbon definition and command list, joined with what the sweep found in the running
    exe (controls per tab, menu and gallery entries, dialogs and their controls, context menus, status bar, backstage,
    shortcuts). Defined but not reachable, and reachable but not defined, are listed."""
    import re
    src = HERE / "results" / f"{DATE}-{args.out}.json"
    rows = json.loads(src.read_text(encoding="utf-8"))["rows"]
    app = App()
    try:
        inv = app.p.js("__app.inventory()")
    finally:
        app.close()
    norm = lambda s: re.sub(r"\s+", " ", re.sub(r"\(Strg[^)]*\)|\.\.\.|…", "", s or "")).strip().lower()
    # What the sweep reached: tab -> group -> set of control labels.
    reached: dict[tuple[str, str], set[str]] = {}
    tab_names = {"Layout": "Layout"}
    for r in rows:
        if " / " not in r["path"] or r["tab"] in ("Kontextmenü", "Statusleiste", "Datei", "Tastenkombination"):
            continue
        group, rest = r["path"].split(" / ", 1)
        label = rest.split(" [")[0]
        reached.setdefault((r["tab"], group), set()).add(norm(label))
    tabs_out = []
    defined_total = unreachable = 0
    for t in inv["tabs"]:
        tname = t["name"]
        groups = []
        for g in t["groups"]:
            ctrls = []
            got = reached.get((tname, g["name"]), set())
            for c in g["controls"] + ([dict(g["launcher"], type="launcher")] if g.get("launcher") else []):
                defined_total += 1
                ok = norm(c["name"]) in got or any(norm(c["name"]) in x or x in norm(c["name"]) for x in got if len(x) > 3)
                if not ok:
                    unreachable += 1
                ctrls.append({**c, "reached": ok})
            groups.append({"name": g["name"], "key": g.get("key"), "controls": ctrls})
        tabs_out.append({"name": tname, "key": t.get("key"), "contextual": t["contextual"], "groups": groups})
    # Reachable but not defined in the ribbon definition (own panels' controls are entries, counted below).
    defined = {(t["name"], g["name"], norm(c["name"])) for t in tabs_out for g in t["groups"] for c in g["controls"]}
    undefined = sorted({f"{r['tab']} > {r['path']}" for r in rows if " / " in r["path"] and " > " not in r["path"] and "Weitere Optionen" not in r["path"]
                        and r["tab"] not in ("Kontextmenü", "Statusleiste", "Datei", "Tastenkombination")
                        and not any(k[0] == r["tab"] and k[1] == r["path"].split(" / ")[0] and (k[2] in norm(r["path"].split(" / ", 1)[1].split(" [")[0]) or norm(r["path"].split(" / ", 1)[1].split(" [")[0]) in k[2]) for k in defined)})
    entries = [r for r in rows if " > " in r["path"] or r["tab"] in ("Kontextmenü", "Statusleiste", "Datei", "Tastenkombination")]
    dialogs = {}
    for r in rows:
        dlg = r.get("dialog")
        if isinstance(dlg, dict) and dlg.get("title"):
            dialogs.setdefault(dlg["title"], {"opened_by": f"{r['tab']} > {r['path']}", "controls": [
                {k: v for k, v in c.items() if k in ("tag", "type", "label", "options")} for c in dlg.get("controls", [])]})
    out = HERE / "results" / f"{DATE}-write-options.json"
    data = {
        "date": DATE, "source": src.name,
        "counts": {"ribbon_controls_defined": defined_total, "ribbon_controls_not_reached": unreachable, "options_swept": len(rows),
                   "menu_gallery_dialog_entries": len(entries), "dialogs": len(dialogs), "commands": len(inv["commands"]),
                   "commands_deferred": sum(1 for c in inv["commands"] if c["deferred"]),
                   "commands_not_implemented": sum(1 for c in inv["commands"] if not c["implemented"] and not c["deferred"])},
        "not_built": ["Diagrammentwurf (chart tab)", "Format (chart format tab)"],
        "tabs": tabs_out, "reachable_not_defined": undefined, "dialogs": dialogs,
        "context_menus": {r["path"]: r.get("items", []) for r in rows if r["tab"] == "Kontextmenü" and r.get("kind") == "popup"},
        "statusbar": [r["path"] for r in rows if r["tab"] == "Statusleiste"],
        "backstage": next((r.get("items") for r in rows if r["tab"] == "Datei" and r["path"] == "(Backstage)"), []),
        "shortcuts": [{"id": c["id"], "label": c["label"], "shortcut": c["shortcut"]} for c in inv["commands"] if c.get("shortcut")],
        "commands": inv["commands"],
    }
    out.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    log(json.dumps(data["counts"]), "->", out.name)


def word_select(app, d, setup: str):
    """The same selection as the sweep's set-up, in Word."""
    r = d.Content
    needle = {"table": "Sonnenblume", "picture": None}.get(setup, ws.TARGET)
    if needle is None:
        d.InlineShapes(1).Range.Select()
        return app.Selection
    r.Find.Execute(FindText=needle, MatchCase=True)
    if setup in ("caret", "table"):
        r.SetRange(r.Start + 2, r.Start + 2)
    r.Select()
    return app.Selection


def run_word(args) -> None:
    """Word opens every file the sweep saved; where Word has the same option, its change is compared with ours."""
    import shutil
    import re
    from feellab.word_ops import OPS, OPS_NO_GROUP
    src = HERE / "results" / f"{DATE}-{args.out}.json"
    data = json.loads(src.read_text(encoding="utf-8"))
    rows = data["rows"]
    wdir = TMP / "word"
    wdir.mkdir(exist_ok=True)
    app = ws.word_app()
    results = []
    try:
        base_w = ws.summarize(app, BASE)
        ours_base_file = TMP / "saved" / "base-ours.docx"
        base_o = ws.summarize(app, ours_base_file) if ours_base_file.exists() else base_w
        n = 0
        for r in rows:
            f = r.get("saved")
            if not f or f.startswith("ERROR"):
                continue
            key = f"{r['tab']} > {r['path']}"
            item = {"path": key, "file": f}
            try:
                after_o = ws.summarize(app, TMP / "saved" / f)
                item["opens"] = True
            except Exception as e:  # noqa: BLE001
                item.update(opens=False, error=str(e)[:300])
                results.append(item)
                log("NOOPEN", key)
                continue
            dO = ws.delta(base_o, after_o)
            item["ours"] = dO
            op = OPS.get(key) or OPS_NO_GROUP.get(re.sub(r"> [^>]*? / ", "> ", key, count=1))
            if op:
                n += 1
                wf = wdir / f"w{n:04d}.docx"
                shutil.copy(BASE, wf)
                d = app.Documents.Open(str(wf), AddToRecentFiles=False, Visible=False)
                try:
                    d.Windows(1).View.Type = 3
                    s = word_select(app, d, r.get("setup", "text"))
                    op(app, d, s)
                    d.Save()
                finally:
                    d.Close(0)
                after_w = ws.summarize(app, wf)
                dW = ws.delta(base_w, after_w)
                item["word"] = dW
                keys = set(dO) | set(dW)
                # Page count and word count follow from the rest; text is compared without layout.
                diff = {k: {"ours": dO.get(k, "(unchanged)"), "word": dW.get(k, "(unchanged)")} for k in keys
                        if k not in ("pages",) and dO.get(k, "(unchanged)") != dW.get(k, "(unchanged)")}
                item["same"] = not diff
                item["diff"] = diff
                log("SAME " if not diff else "DIFF ", key, "" if not diff else json.dumps(diff, ensure_ascii=False)[:400])
            results.append(item)
    finally:
        app.Quit(0)
    out = HERE / "results" / f"{DATE}-{args.out}-word.json"
    summary = {"files": len(results), "opened": sum(1 for x in results if x.get("opens")),
               "compared": sum(1 for x in results if "same" in x), "same": sum(1 for x in results if x.get("same"))}
    out.write_text(json.dumps({"date": DATE, "summary": summary, "items": results}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    log(json.dumps(summary), "->", out.name)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("what", choices=["sweep", "inventory", "word", "base"])
    ap.add_argument("--tabs", default="")
    ap.add_argument("--gallery", type=int, default=None)
    ap.add_argument("--extra", default="", help="context,status,backstage,keys (all when no --tabs)")
    ap.add_argument("--out", default="write-sweep")
    args = ap.parse_args()
    if args.what == "base":
        make_base()
    elif args.what == "sweep":
        run_sweep(args)
    elif args.what == "word":
        run_word(args)
    elif args.what == "inventory":
        run_inventory(args)


if __name__ == "__main__":
    main()
