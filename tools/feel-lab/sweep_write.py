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
import re
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
DATE = os.environ.get("SWEEP_DATE") or datetime.date.today().isoformat()
PORT = int(os.environ.get("SWEEP_PORT", "9471"))

# Which document set-up each tab needs.
TAB_SETUP = {"Tabellenentwurf": "table", "Layout (Tabelle)": "table", "Kopf- und Fußzeile": "headerFooter",
             "Bildformat": "picture", "Formformat": "shape", "Formel": "math"}
CONTEXT_TABS = {"table": ["Tabellenentwurf", "Tabellenlayout"], "headerFooter": ["Kopf- und Fußzeile"], "picture": ["Bildformat"],
                "shape": ["Formformat"], "math": ["Formel"]}
# The parts of the UI a sweep step may leave changed (view, rulers, reduced ribbon ...).
VIEW_JS = "(() => { const u = __sw.ui(); return JSON.stringify([u.view, u.shell, u.minimized, u.bodyCls]); })()"
# Closes the open panes (Navigation, Formatvorlagen, Kommentare ...) by their own close buttons.
CLOSE_PANES = """(() => { let n = 0; for (const p of document.querySelectorAll('.navpane, aside, .taskpane, [role=complementary], .tpane, .clippane, .fxpane')) {
  if (p.hidden || !p.getBoundingClientRect().width) continue;
  const b = p.querySelector('.tp-close, .nav-close, button[aria-label="Schließen"], button[data-tip="Schließen"], button[aria-label$="schließen"]'); if (b) { b.click(); n++; } } return n; })()"""
# Back into the main text after typing into a note or a comment (the set-up's own step, not an option under test).
BACK_TO_MAIN = "(() => { const e = __app.editor(); e.doc.activeStory = 'main'; e.select(e.doc.fromOffset(0), e.doc.fromOffset(0)); e.focus(); return true; })()"
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
                    p.pump(0.2)
                    p.type("Formtext")
                    p.pump(0.1)
                    # Select the shape itself again (its frame), as after drawing.
                    p.key("Escape")
        elif setup in ("notes", "revisions"):
            # Something to go to, accept or reject: two footnotes and an endnote, or a comment and tracked changes.
            if setup == "notes":
                for needle, cmd in (("Thema", "ref.footnote"), ("Schulgarten", "ref.footnote"), ("Messwerte", "ref.endnote")):
                    p.js(f"__sw.selectText({json.dumps(needle)}, true)")
                    p.js(f"__app.run({json.dumps(cmd)})")
                    p.pump(0.15)
                    p.type("Anmerkung")
                    p.js(BACK_TO_MAIN)
            else:
                p.js("__sw.selectText('Thema', false)")
                p.js("__app.run('review.newComment')")
                p.pump(0.2)
                p.type("Bitte prüfen")
                p.js(BACK_TO_MAIN)
                p.js("__probe.focus()")
                p.js("__app.run('review.trackChanges')")
                p.js("__sw.selectText('Schluss.', true)")
                p.type("neu ")
                p.js("__sw.selectText('Ende', false)")
                p.key("Delete")
                p.js("__app.run('review.trackChanges')")
            p.js("__sw.selectText('weitere Beobachtungen')")
        elif setup == "math":
            p.js("__sw.selectText('Schluss.', true)")
            p.key(")", ctrl=False, shift=True, alt=True, vk=48, code="Digit0")  # Alt+Umschalt+0
        p.pump(0.25)

    # ----- ribbon -------------------------------------------------------------------------
    def click_tab(self, name: str) -> bool:
        r = self.p.js(f"__sw.tabRect({json.dumps(name)})")
        if not r:
            return False
        active = self.p.js("(() => { const t = document.querySelector('.tabrow .tab.active'); return t ? (t.classList.contains('contextual') && t.textContent.trim() === 'Layout' ? 'Tabellenlayout' : t.textContent.trim()) : ''; })()")
        if active != name:
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
    for k in ("tab", "ctx", "checked", "combos", "backstage", "status", "view", "shell", "panes", "dialogs", "popups", "sel", "docs", "bodyCls", "minimized", "native", "popSig", "bsPage", "overlay", "focus", "comments"):
        if a.get(k) != b.get(k):
            out.append(k)
    return out


# A German keyboard (D-150): what each key gives with Umschalt and with AltGr (= Strg+Alt), its code and virtual key.
# Word binds keys, not characters, so the sweep must send what the keyboard really sends: Strg+Alt+3 arrives as "³",
# Strg+( as Strg+Umschalt+8, Alt+= as Alt+Umschalt+0.
_DIGIT_SHIFT = dict(zip("1234567890", "!\"§$%&/()="))
_DIGIT_ALTGR = {"2": "²", "3": "³", "7": "{", "8": "[", "9": "]", "0": "}"}
_LETTER_ALTGR = {"q": "@", "e": "€", "m": "µ"}
# Characters on a key of their own or with Umschalt: (code, vk, shifted?, unshifted character).
_PUNCT = {"+": ("BracketRight", 187, False), "*": ("BracketRight", 187, True), "#": ("Backslash", 191, False), "'": ("Backslash", 191, True),
          "<": ("IntlBackslash", 226, False), ">": ("IntlBackslash", 226, True), "-": ("Slash", 189, False), "_": ("Slash", 189, True),
          ",": ("Comma", 188, False), ";": ("Comma", 188, True), ".": ("Period", 190, False), ":": ("Period", 190, True)}
_PUNCT_BASE = {"BracketRight": "+", "Backslash": "#", "IntlBackslash": "<", "Slash": "-", "Comma": ",", "Period": "."}
_NAMED = {"Eingabe": ("Enter", "Enter", 13), "Return": ("Enter", "Enter", 13), "Leertaste": (" ", "Space", 32), "Rücktaste": ("Backspace", "Backspace", 8),
          "Entf": ("Delete", "Delete", 46), "Tab": ("Tab", "Tab", 9), "Pos1": ("Home", "Home", 36), "Ende": ("End", "End", 35), "Esc": ("Escape", "Escape", 27)}


def german_key(combo: str) -> dict | None:
    """The key event a German keyboard sends for a shortcut as Word writes it ("Alt+Strg+3", "Strg+(", "Strg+Eingabe")."""
    parts = combo.split("+")
    if combo.endswith("++"):
        parts = combo[:-2].split("+") + ["+"]
    mods = {m for m in parts[:-1]}
    name = parts[-1]
    ctrl, alt, shift = "Strg" in mods, "Alt" in mods, "Umschalt" in mods
    if name in _NAMED:
        key, code, vk = _NAMED[name]
    elif len(name) > 1 and name[0] == "F" and name[1:].isdigit():
        key, code, vk = name, name, 111 + int(name[1:])
    elif len(name) == 1 and name.isalpha() and name.isascii():
        low = name.lower()
        key, code, vk = (name.upper() if shift else low), f"Key{low.upper()}", ord(low.upper())
    elif name.isdigit() or name in _DIGIT_SHIFT.values():
        digit = name if name.isdigit() else next(d for d, ch in _DIGIT_SHIFT.items() if ch == name)
        if not name.isdigit():
            shift = True
        key, code, vk = (_DIGIT_SHIFT[digit] if shift else digit), f"Digit{digit}", 48 + int(digit)
    elif name in _PUNCT:
        code, vk, shifted = _PUNCT[name]
        shift = shift or shifted
        key = name if shifted or not shift else next((ch for ch, v in _PUNCT.items() if v[0] == code and v[2]), name)
    else:
        return None
    # Strg+Alt is AltGr: the key gives its AltGr character where it has one (and types it unless the app stops it).
    altgr = False
    if ctrl and alt and not shift:
        if code.startswith("Digit") and code[5:] in _DIGIT_ALTGR:
            key, altgr = _DIGIT_ALTGR[code[5:]], True
        elif code.startswith("Key") and code[3:].lower() in _LETTER_ALTGR:
            key, altgr = _LETTER_ALTGR[code[3:].lower()], True
    mods_bits = (1 if alt else 0) | (2 if ctrl else 0) | (8 if shift else 0)
    # A character is typed only when the keys give one: alone, with Umschalt, or with AltGr.
    typed = len(key) == 1 and (altgr or not ctrl and not alt)
    return {"key": key, "code": code, "vk": vk, "mods": mods_bits, "text": key if typed else ""}


def press(p, ev: dict) -> None:
    """Sends the key event (keyDown with its character, or rawKeyDown) and its keyUp."""
    base = {"key": ev["key"], "code": ev["code"], "windowsVirtualKeyCode": ev["vk"], "nativeVirtualKeyCode": ev["vk"], "modifiers": ev["mods"]}
    if ev["text"]:
        p.cdp("Input.dispatchKeyEvent", {"type": "keyDown", **base, "text": ev["text"], "unmodifiedText": ev["text"]})
    else:
        p.cdp("Input.dispatchKeyEvent", {"type": "rawKeyDown", **base})
    p.cdp("Input.dispatchKeyEvent", {"type": "keyUp", **base})


# Shortcuts whose set-up is part of the test (something typed to take back, the clipboard, the format to copy, a
# document saved, made or closed): compared from a fresh set-up, not in the same document.
SPECIAL_KEYS = {"edit.undo", "edit.redo", "edit.paste", "edit.pasteText", "edit.formatPainter", "file.new", "file.close", "file.save", "file.open",
                "insert.equation", "view.zoom100"}  # an equation opens its own edit session; the zoom is set first


class Sweep:
    def __init__(self, app: App, gallery: int | None, saved_index: dict) -> None:
        self.a = app
        self.gallery = gallery
        self.rows: list[dict] = []
        self.saved = saved_index
        self.last = time.perf_counter()

    def record(self, row: dict) -> None:
        now = time.perf_counter()
        row["secs"] = round(now - self.last, 1)
        self.last = now
        self.rows.append(row)
        st = row["status"]
        log(f"{st:9} {row['tab']} > {row['path']}  {row.get('effects', '')} {row.get('note', '')[:140]} ({row['secs']} s)")

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
        if c["groupId"] == "clipboard" and c["label"] in ("Einfügen", "Weitere Optionen"):
            self.prime_clipboard()
            before = a.state()
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
            if c["groupId"] in ("insertShapes", "shapes") or c["group"] == "Formen einfügen":
                if "doc" not in effects(before, after) and not ui["dialogs"]:
                    self.draw_on_page()
                    after = a.state()
                    errors += p.take_errors()
            row = self.judge(tab, path, setup, before, after, errors)
            if row["status"] == "noeffect" and c["label"] == "Kopieren":
                info = p.js("__app.clipboardInfo()")
                row["status"], row["note"] = ("ok", f"clipboard {info}") if info and info.get("text") else ("noeffect", f"clipboard {info}")
            self.record(row)

    def draw_on_page(self) -> None:
        """A drag on the page (drawing the shape or text box picked before)."""
        p = self.a.p
        pg = p.js("(() => { const b = document.querySelector('.page').getBoundingClientRect(); return {x: b.left + 300, y: b.top + 520}; })()")
        p.cdp("Input.dispatchMouseEvent", {"type": "mousePressed", "x": pg["x"], "y": pg["y"], "button": "left", "clickCount": 1, "buttons": 1})
        for k in range(1, 9):
            p.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": pg["x"] + 15 * k, "y": pg["y"] + 10 * k, "button": "left", "buttons": 1})
        p.cdp("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": pg["x"] + 120, "y": pg["y"] + 80, "button": "left", "clickCount": 1, "buttons": 0})
        p.pump(0.4)

    def prime_clipboard(self) -> None:
        """Einfügen needs something on the clipboard: the target text is copied first (Kopieren), the caret put after it."""
        p = self.a.p
        p.js("__app.run('edit.copy')")
        p.pump(0.3)
        p.js("__sw.selectText('Schluss.', true)")

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
        if chain and chain[0][0] == "control" and chain[0][1].get("groupId") == "clipboard":
            self.prime_clipboard()
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
            p.pump(0.22)
        return True

    def pick(self, items: list[dict]) -> list[dict]:
        # Menus are taken whole; only long galleries (more than 14 entries) are sampled.
        if self.gallery is None or len(items) <= max(self.gallery, 14):
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
            pr = p.js(f"__sw.prep({hit['idx']})")
            if pr:
                hit = {**hit, "x": pr["x"], "y": pr["y"]}
                if not pr["hits"]:
                    self.record({"tab": tab, "path": ipath, "setup": setup, "status": "error", "note": "entry covered: a click at its place does not reach it"})
                    p.key("Escape")
                    continue
            before = a.state()
            n_pop = before["ui"]["popups"]
            p.move(hit["x"], hit["y"])
            p.pump(0.08)
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
            if ("Transparente Farbe bestimmen" in ipath or "Rahmenpipette" in ipath) and "doc" not in effects(before, after):
                # Modes that wait for a click: on the picture (its colour) or on the table's first border.
                tgt = p.js("__sw.pictureRect()") if "Transparente" in ipath else p.js(
                    "(() => { const t = __probe.text(); const i = t.indexOf('Pflanze'); const a = __probe.pointAt(i); return {x: a.x - 3, y: a.y + a.h / 2}; })()")
                if tgt:
                    p.click(tgt["x"] + (25 if "Transparente" in ipath else 0), tgt["y"] + (-35 if "Transparente" in ipath else 0))
                    p.pump(0.4)
                    after = a.state()
                    errors += p.take_errors()
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
        for it in (items if depth == 0 else self.pick(items)):  # a context menu entry by entry
            ipath = f"{path} > {it['label'] or '(ohne Namen)'}"
            if it["disabled"]:
                self.record({"tab": tab, "path": ipath, "setup": setup, "status": "deferred" if DEFERRED_MARK in it.get("tip", "") else "disabled"})
                continue
            if it["tag"] in ("INPUT", "SELECT"):
                self.record({"tab": tab, "path": ipath, "setup": setup, "status": "field"})
                continue
            # The paste entries need something on the clipboard (the clipboard outlives the set-up's reset).
            if tab == "Kontextmenü" and re.search(r"Formatierung beibehalten|Nur den Text übernehmen|Formatierung zusammenführen|^Einfügen", it["label"] or ""):
                self.prime_clipboard()
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
            pr = p.js(f"__sw.prep({hit['idx']})")
            if pr:
                hit = {**hit, "x": pr["x"], "y": pr["y"]}
            before = a.state()
            p.move(hit["x"], hit["y"])
            p.pump(0.08)
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
            row = self.judge(tab, ipath, setup, before, after, errors)
            if row["status"] == "noeffect" and re.match(r"(Kopieren|Ausschneiden)", it["label"] or ""):
                # Kopieren changes nothing on the page: what it put on the clipboard is its effect.
                info = p.js("__app.clipboardInfo()")
                got = bool(info and (info.get("text") or info.get("html") or info.get("image")))
                row["status"], row["note"] = ("ok" if got else "noeffect"), f"clipboard {info}"
            self.record(row)
            p.key("Escape")

    def statusbar(self) -> None:
        a, p = self.a, self.a.p
        a.fresh("text")
        cells = p.js("[...document.querySelectorAll('.statusbar button, .statusbar input, .statusbar [role=button]')].filter(b => b.getBoundingClientRect().width > 0).map((b, i) => ({i, label: (b.getAttribute('aria-label') || (b.dataset.tip || '').split('\\n')[0] || b.textContent).trim(), r: (() => { const r = b.getBoundingClientRect(); return {x: r.left + r.width / 2, y: r.top + r.height / 2}; })()}))")
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
            if name == "Schließen":
                p.key("Escape")
                p.js("__app.newDoc()")
                r0 = p.js("__sw.tabRect('Datei')")
                p.click(r0["x"], r0["y"])
                p.pump(0.4)
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
        """Every keyboard shortcut of the command list, pressed as a German keyboard sends it (german_key: the key's
        character with Umschalt or AltGr, its code and virtual key), with the target selected. Each is compared with
        the same command run directly from the same set-up (D-150): the shortcut must change the document the same
        way and open the same dialog or pane, so a shortcut that types a character instead (AltGr) fails."""
        a, p = self.a, self.a.p
        skip = {"Strg+F1"}  # reduces the ribbon (checked with the ribbon's own button)
        for c in commands:
            sc = c.get("shortcut")
            if not sc or c.get("deferred"):
                continue
            for combo in [s.strip() for s in sc.split(",")][:1]:
                path = f"{combo} ({c['label']})"
                if os.environ.get("SWEEP_KEYS") and not re.search(os.environ["SWEEP_KEYS"], path):
                    continue
                if combo in skip:
                    self.record({"tab": "Tastenkombination", "path": path, "setup": "text", "status": "native"})
                    continue
                setup = "table" if c["id"].startswith("table.") else "text"
                ev = german_key(combo)
                if ev is None:
                    self.record({"tab": "Tastenkombination", "path": path, "setup": setup, "status": "error", "note": "key not known to the sweep"})
                    continue
                cid = c["id"]

                def prepare() -> None:
                    a.fresh(setup)
                    if cid == "file.close":
                        p.js("__app.newDoc()")  # the last document would close the window
                    elif cid == "file.save":
                        # Speichern writes to a copy, not the base document; the file's time shows the save.
                        p.js(f"__app.saveTo({json.dumps(str(TMP / 'keys-save.docx'))})")
                    elif cid in ("edit.undo", "edit.redo"):
                        p.type("x")  # something to take back
                        p.pump(0.1)
                        if cid == "edit.redo":
                            p.key("z", ctrl=True)
                            p.pump(0.1)
                    elif cid in ("edit.paste", "edit.pasteText"):
                        self.prime_clipboard()
                    elif cid == "edit.formatPainter":
                        p.js("__sw.selectText('Ein Bericht', false)")  # the heading's format is copied ...
                    elif cid == "view.zoom100":
                        press(p, german_key("Strg++"))  # another zoom first, so 100% has something to do
                        p.pump(0.2)
                    p.js("__sw.markBase()")
                    p.take_errors()

                def finish() -> None:
                    if cid == "edit.formatPainter":
                        # ... and put on the target with Alt+Strg+V, as Word's CopyFormat / PasteFormat.
                        p.js("__sw.selectText('weitere Beobachtungen')")
                        press(p, german_key("Strg+Alt+V"))
                    p.pump(0.35)

                # The command itself, run from the same set-up: the reference.
                prepare()
                mt0 = (TMP / "keys-save.docx").stat().st_mtime if cid == "file.save" else 0
                ref0 = a.state()
                if cid == "edit.formatPainter":
                    press(p, ev)  # no command of its own (the ribbon's brush differs): the shortcut is its own reference
                else:
                    p.js(f"(void __app.run({json.dumps(cid)}), 0)")
                finish()
                ref = a.state()
                p.take_errors()
                p.key("Escape")
                # The shortcut, in the same document where the command changed it: back with Rückgängig first, so the
                # two results can be compared exactly (a document read anew has other internal ids).
                restored = False
                if ref["doc"] != ref0["doc"] and cid not in SPECIAL_KEYS and not cid.startswith("math."):
                    p.pump(0.1)
                    for _ in range(3):
                        if p.js("__sw.docFp()") == ref0["doc"]:
                            break
                        p.js("(void __app.run('edit.undo'), 0)")
                        p.pump(0.15)
                    restored = p.js("__sw.docFp()") == ref0["doc"] and not p.js("__sw.dialogs()").__len__()
                if restored:
                    a.setup(setup)
                    p.js("__sw.markBase()")
                    p.take_errors()
                else:
                    prepare()
                before = a.state()
                press(p, ev)
                finish()
                after = a.state()
                errors = p.take_errors()
                extra = {"cmd": cid, "keys": {k: ev[k] for k in ("key", "code", "vk", "mods")}}
                if after["ui"]["dialogs"]:
                    extra["dialog"] = (p.js("__sw.dialogInfo()") or {}).get("title")
                if cid == "edit.copy":
                    info = p.js("__app.clipboardInfo()")
                    extra["note"] = f"clipboard {info}"
                    if info and info.get("text"):
                        extra["clipboard"] = True
                if cid == "file.save":
                    extra["saved_file"] = (TMP / "keys-save.docx").stat().st_mtime > mt0
                # Not edits (Rückgängig and Wiederholen themselves; documents made or closed; a setting): no Strg+Z check.
                undo = cid not in ("file.new", "file.close", "file.open", "review.trackChanges", "edit.undo", "edit.redo")
                row = self.judge("Tastenkombination", path, setup, before, after, errors, extra, undo=undo)
                if row["status"] == "noeffect" and (extra.get("clipboard") or extra.get("saved_file")):
                    row["status"] = "ok"
                # The same as the command: what changed (document, dialogs, panes, view ...), and in the same document
                # the same result.
                keys = ("doc", "dialogs", "panes", "view", "backstage", "docs", "comments", "focus")
                e_ref, e_key = set(effects(ref0, ref)) & set(keys), set(effects(before, after)) & set(keys)
                diff = sorted(e_ref ^ e_key)
                # A comment carries the minute it was written: compared by what changed, not exactly.
                if restored and "doc" in e_ref and after["doc"] != ref["doc"] and cid not in ("review.newComment", "insert.comment"):
                    diff.append("doc result")
                for k in ("dialogs", "panes"):
                    if after["ui"].get(k) != ref["ui"].get(k) and k not in diff:
                        diff.append(k)
                row["same_as_command"] = not diff
                if row["status"] == "ok" and diff:
                    row["status"] = "differs"
                    row["note"] = (row.get("note", "") + " differs from the command run directly: " + ",".join(diff)).strip()
                self.record(row)
                p.key("Escape")


def run_sweep(args) -> None:
    out = HERE / "results" / f"{DATE}-{args.out}.json"
    app = App()
    rows: list[dict] = []
    sw = None
    try:
        sw = Sweep(app, args.gallery, {})
        # Our save of the untouched base document: the reference for what each option changed.
        app.fresh("text")
        app.p.js(f"__app.saveTo({json.dumps(str(TMP / 'saved' / 'base-ours.docx'))})")
        tabs = [t["label"] for t in app.p.js("__sw.tabs()") if not t["contextual"] and t["label"] != "Datei"]
        plan: list[tuple[str, str]] = [(t, "text") for t in tabs]
        for setup, names in CONTEXT_TABS.items():
            plan += [(n, setup) for n in names]
        # Referenzen and Überprüfen once more with notes, a comment and tracked changes to go to, accept and reject.
        plan += [("Referenzen", "notes"), ("Überprüfen", "revisions")]
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
        if sw is not None:
            rows = sw.rows
        try:
            app.close()
        except Exception:  # noqa: BLE001
            pass
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
    rows = []
    for name in args.out.split(","):
        rows += json.loads((HERE / "results" / f"{DATE}-{name}.json").read_text(encoding="utf-8"))["rows"]

    class _Src:
        name = args.out
    src = _Src()
    app = App()
    try:
        inv = app.p.js("__app.inventory()")
    finally:
        app.close()
    norm = lambda s: re.sub(r"\s+", " ", re.sub(r"\(Strg[^)]*\)|\.\.\.|…|-\n", "", s or "")).strip().lower()
    # What the sweep reached: tab -> group -> set of control labels.
    reached: dict[tuple[str, str], set[str]] = {}
    tab_names = {"Layout": "Layout"}
    for r in rows:
        if " / " not in r["path"] or r["tab"] in ("Kontextmenü", "Statusleiste", "Datei", "Tastenkombination"):
            continue
        group, rest = r["path"].split(" / ", 1)
        label = rest.split(" [")[0]
        got = reached.setdefault((r["tab"], group), set())
        got.add(norm(label))
        # Gallery entries reached through Weitere and entries of the control's own menu count for the group too.
        if " > " in rest:
            got.add(norm(rest.rsplit(" > ", 1)[1]))
        got.add("*")
    tabs_out = []
    defined_total = unreachable = 0
    for t in inv["tabs"]:
        # The table's Layout tab is "Tabellenlayout" in the sweep (the main tab has the same label).
        tname = "Tabellenlayout" if t["contextual"] and t["name"] == "Layout" else t["name"]
        groups = []
        for g in t["groups"]:
            ctrls = []
            got = reached.get((tname, g["name"]), set())
            for c in g["controls"] + ([dict(g["launcher"], type="launcher")] if g.get("launcher") else []):
                defined_total += 1
                ok = norm(c["name"]) in got or any(norm(c["name"]) in x or x in norm(c["name"]) for x in got if len(x) > 3) \
                    or (c.get("type") in ("tiles", "combo", "spinner") and "*" in got)
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


# Options that change nothing on the base document because it already is so (Word changes nothing either): the
# entry for the current state, removing what is not there, Weitersuchen with an empty field ...
ALREADY = [
    r"> (Ohne|Keine Farbe|Automatisch|Keine|Kein Rahmen|Standard|Dokument|Office|Hochformat|Eins|Normal)$",
    r"Zahlenformatvorlagen > Standard$", r"Stil-Sets > Standard$", r"Schriftarten \[dropdown\] > Office,", r"Format \[dropdown\] > A4",
    r"entfernen$", r"Linksbündig \[button\]$", r"Einzug verkleinern \[button\]$", r"Formatvorlagen / Standard \[gallery-item\]$",
    r"> (Weitersuchen|Suchen|Gehe zu|Ersetzen|Alle ersetzen)$", r"Gehe zu > Gehe zu$", r"Markieren \[dropdown\] > Objekte auswählen$",
    r"Schriftgrad \[combo-input\] = 11$", r"Öffnen \[combo-arrow\] > 11$",
    # Going to the next note, comment or change when the base document has none (Word: nothing happens).
    r"(Nächste|Vorherige) (Fußnote|Endnote)", r"Nächste Fußnote \[split-main\]$", r"(Vorheriger|Nächster) Kommentar", r"(Vorherige|Nächste) Änderung",
    # Showing or filtering comments when there are none; the markup mode already chosen.
    r"Kommentare / Weitere Optionen \[split-arrow\] > (Kontextbezogen|Liste)$", r"Markup \[dropdown\] > Kommentare$", r"Markup filtern \[dropdown\] > Kommentare$",
    r"> Markup: alle$",
    # The picture as inserted (no correction, full saturation, no transparency, nothing to reset), no picture on the
    # clipboard; the table already in Tabellenraster; the header already open.
    r"Schärfen und Weichzeichnen: 0 %$", r"Sättigung: 100 %$", r"Transparenz: 0 %$", r"Bild zurücksetzen \[split-main\]$", r"> Bild (und Größe )?zurücksetzen$",
    r"Bild ändern \[dropdown\] > Aus Zwischenablage$", r"Tabellenraster( \[gallery-item\])?$", r"Zur Kopfzeile wechseln \[button\]$",
    r"Kopfzeile \[dropdown\] > Kopfzeile bearbeiten$",
    # View gridlines of a table that has borders all round: nothing to see (Word neither).
    r"> Gitternetzlinien anzeigen$",
    # In a new, empty equation: nothing to convert; Unicode is the linear format already. Katalogfilter opens its
    # list inside the open panel (which the sweep does not count as a new popup).
    r"Konvertierungen / Unicode \[button\]$", r"Konvertierungen / Konvertieren \[split-main\]$", r"Konvertierungen / Weitere Optionen \[split-arrow\] > ",
    r"> Katalogfilter$",
    # The only shape on the page: nothing to move behind (Word neither).
    r"Formformat > Anordnen / Ebene nach hinten \[split-main\]$",
    # Drucklayout is the view already; Datei opens on Startseite.
    r"^Statusleiste > Drucklayout #\d+$", r"^Datei > Startseite$",
    # The target starts mid-sentence and has no capitals inside: Word's Satzanfang groß changes nothing (recorded).
    r"Groß-/Kleinschreibung \[dropdown\] > Ersten Buchstaben im Satz großschreiben\.$", r"Ansichten / Drucklayout \[button\]$", r"Zoom / 100% \[button\]$",
]
# The same, for options found in run B (D-151): each changes nothing visible on the audit's document, as in Word.
ALREADY += [
    # No change and no tracking in the plain text set-up: nothing to accept or reject.
    r"Alle Änderungen (annehmen|ablehnen) und Nachverfolgung beenden$",
    # The picture is inline already, uncropped (it fills its frame: Ausfüllen and Einpassen change nothing), and a
    # rectangle; Als Standardlayout festlegen is a setting for pictures inserted later.
    r"> Mit Text in Zeile$", r"> Als Standardlayout festlegen$", r"Auf Form zuschneiden > Rechteck$",
    r"Größe / Weitere Optionen \[split-arrow\] > (Ausfüllen|Einpassen)$",
    # What the next Ausrichten aligns to (a setting of the menu; nothing moves until Ausrichten is used).
    r"Ausrichten \[dropdown\] > (An Seite ausrichten|Am Seitenrand ausrichten)$",
    # The new rectangle: no gradient, its text centred already; the only shape, in front already.
    r"Farbverlauf > Keine Farbverläufe$", r"Text ausrichten \[dropdown\] > Zentriert$", r"Ebene nach vorne \[split-main\]$",
    # No picture on the clipboard.
    r"Bild ändern > Aus Zwischenablage(\.\.\.)?$",
    # The print page of a one-page document: no page before or after.
    r"^Datei > Drucken > (Vorherige|Nächste) Seite$",
]
# Options that do nothing on the audit's document in Word either, recorded (results/2026-11-09-word-followup.json),
# also in the notes and revisions set-ups.
RECORDED_NOTHING = [
    # Diese Änderung annehmen / ablehnen with no change in the selection: Word refuses (ExecuteMso fails, nothing changes).
    r"Weitere Optionen \[split-arrow\] > Diese Änderung (annehmen|ablehnen)$",
    # Kontextbezogen is how comments are shown already.
    r"Kommentare / Weitere Optionen \[split-arrow\] > Kontextbezogen$",
]
# The print page's own fields (copies, pages, the page range list): typed into, not clicked.
FIELDS = [r"^Datei > Drucken > (INPUT( #\d+)?|Alle Seiten drucken.*)$"]
# Options Word does not take back with Strg+Z either (a setting, not an edit).
NO_UNDO_IN_WORD = [r"Änderungen nachverfolgen \[split-main\]$", r"Änderungen nachverfolgen$",
                   # Closing a header opened empty drops the empty header again (no edit to take back; Word keeps an empty part).
                   r"Kopf- und Fußzeile schließen \[button\]$"]
# Options the sweep cannot finish the way a student does; counted as not tested (not passed), each with its reason.
UNTESTED = [
    # Freihandform and Kurve are drawn point by point and finished with a double click; the sweep only drags.
    r"(Freihandform: Form|Kurve|Freihand: Skizze)( \[gallery-item\])?$",
    # Rahmenpipette waits for a click on a border; the sweep's click does not land on the table's border line.
    r"> Rahmenpipette$",
    # The mini toolbar's Format übertragen arms the painter; it acts on the next selection, which the sweep does not make.
    r"^Kontextmenü > \w+ > Format übertragen$",
    # Zuschneiden switches the picture into cropping; the sweep sees no change it can tell from the page (not verified).
    r"Zuschneiden( \[split-main\])?$",
    # A click in the middle of the zoom slider leaves the zoom (the slider is dragged in the real-input pass).
    r"^Statusleiste > Zoom #\d+$",
]
FAIL = {"error", "undo", "crash", "esc", "docchange-on-open", "noeffect"}


def classify(rows: list[dict]) -> dict:
    """passed / failed / greyed per the sweep's rows (ALREADY counts as passed)."""
    import re
    out = {"options": len(rows), "passed": 0, "failed": 0, "untested": 0, "greyed": 0, "deferred": 0, "fields": 0, "failures": []}
    for r in rows:
        st = r["status"]
        key = f"{r['tab']} > {r['path']}"
        # An entry whose only effect is that its menu closed did nothing a student sees.
        eff = set(r.get("effects") or [])
        if st == "ok" and "popups" in eff and eff <= {"popups", "popSig"} and r.get("kind") not in ("popup", "submenu") and " > " in r["path"] \
                and not r["path"].endswith(" > Esc"):
            st = "noeffect"
        if st == "noeffect" and not (r.get("setup") in ("notes", "revisions") and re.search(r"Fußnote|Endnote|Kommentar|Änderung", key)) \
                and any(re.search(p, key) for p in ALREADY):
            st = "ok"
        if st == "noeffect" and any(re.search(p, key) for p in RECORDED_NOTHING):
            st = "ok"
        if st == "noeffect" and any(re.search(p, key) for p in FIELDS):
            st = "field"
        if st == "undo" and any(re.search(p, key) for p in NO_UNDO_IN_WORD):
            st = "ok"
        if st == "noeffect" and any(re.search(p, key) for p in UNTESTED):
            st = "untested"
        if st in ("ok", "native"):
            out["passed"] += 1
        elif st == "untested":
            out["untested"] += 1
        elif st == "deferred":
            out["deferred"] += 1
        elif st == "disabled":
            out["greyed"] += 1
        elif st == "field":
            out["fields"] += 1
        else:
            out["failed"] += 1
            out["failures"].append({"path": key, "status": st, "note": r.get("note", ""), "errors": r.get("errors", [])[:2]})
    return out


def run_report(args) -> None:
    rows = []
    for name in args.out.split(","):
        rows += json.loads((HERE / "results" / f"{DATE}-{name}.json").read_text(encoding="utf-8"))["rows"]
    c = classify(rows)
    print(json.dumps({k: v for k, v in c.items() if k != "failures"}, ensure_ascii=False))
    for f in c["failures"]:
        print(f"{f['status']:9} {f['path']}  {f['note'][:120]}  {json.dumps(f['errors'], ensure_ascii=False)[:200]}")


def word_select(app, d, setup: str):
    """The same selection as the sweep's set-up, in Word."""
    # In the table the cell "Sonnenblume" (the list above has "Sonnenblumen").
    r = d.Tables(1).Range if setup == "table" else d.Content
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
            # The notes and revisions set-ups add notes, a comment and changes first, which Word's side does not.
            if r.get("setup") in ("notes", "revisions", "shape", "math", "headerFooter"):
                op = None
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
                except Exception as e:  # noqa: BLE001
                    item["word_error"] = str(e)[:300]
                finally:
                    d.Close(0)
                if "word_error" in item:
                    log("WORDERR", key, item["word_error"][:200])
                    results.append(item)
                    continue
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
    ap.add_argument("what", choices=["sweep", "inventory", "word", "base", "report"])
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
    elif args.what == "report":
        run_report(args)


if __name__ == "__main__":
    main()
