"""Probe: PowerPoint's complete inventory (Step 8, made like probe_excel_inventory.py, D-066).

In a PowerPoint of our own (visible, real clicks) every ribbon tab is walked, including the tabs
that appear only for a selected object or in a master view (shape, text box, text being edited,
picture, table, chart, SmartArt, group, equation, Folienmaster, Handzettelmaster, Notizenmaster);
every menu, gallery and split button is opened (submenus one level deep), every dialog launcher and
every entry ending in "..." is opened and its controls listed (each page of a dialog with tabs),
then closed with Esc or Abbrechen, never with OK. Then the context menus (right-click) of the
slide panel, placeholders, text, objects, the empty slide and the notes pane; the File backstage
pages and the PowerPoint-Optionen; the keytips. Labels and control types only: no descriptions or
tooltips (D-058). Buttons that act at once are listed, not pressed; cloud features are listed and
marked, never opened.

PowerPoint runs as one process per user (feellab/powerpoint.py): the probe refuses to run while a
POWERPNT.EXE runs (the owner's). Real input: only one session on this PC may use it at a time, so
the probe takes %TEMP%/raabe-real-input.lock exclusively (refreshed while it runs, deleted at the
end, also on failure) and exits with code 75 if another session holds it. Every click checks that
the point lies on a window of our own PowerPoint. If a sign-in, consent, licence or what's-new
dialog appears, nothing in it is clicked: the probe ends its own PowerPoint and stops (code 77).

    .venv/Scripts/python.exe probe_powerpoint_inventory.py [part …]
    parts: the contexts below, "Kontextmenüs", "Backstage", "Zugriffstasten" (default: all)

Results: captures/<date>-powerpoint-inventory-full.json (everything, not committed) and
results/<date>-powerpoint-inventory.json (labels only, feellab.sanitize.clean), written after each
tab and part, so a crash keeps what was done; a rerun continues where it stopped.
"""
from __future__ import annotations

import ctypes
import datetime
import json
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget, powerpoint_pids

HERE = Path(__file__).parent
DATE = "2026-11-06"
OUT = HERE / "captures" / f"{DATE}-powerpoint-inventory-full.json"
RESULT = HERE / "results" / f"{DATE}-powerpoint-inventory.json"
LOCK = Path(tempfile.gettempdir()) / "raabe-real-input.lock"
LOCK_TAG = "step8-present-inventory"

# Cloud and online features: listed, never opened.
CLOUD = ("Freigeben", "Teilen", "Copilot", "Designer", "Design-Ideen", "Add-Ins", "Office-Add-Ins", "Meine Add-Ins", "Store",
         "Online", "Stockbilder", "Stock", "Piktogramme", "3D-Modelle", "Übersetzen", "Intelligente Suche", "Live-Untertitel",
         "Untertitel", "Cameo", "Kamera", "Diktieren", "Vertraulichkeit", "Microsoft Forms", "Formulare", "Teams",
         "Präsentieren in Teams", "Ideen", "Referentencoach", "Mit Coach", "Coach", "Bing", "OneDrive", "SharePoint", "Anmelden",
         "Aktivität", "Loop", "Clipchamp", "Zusammenfassung", "Erstellen mit Copilot")
# Never pressed, not even to read them: they act at once, record, start the show, leave PowerPoint,
# or show the owner's files, folders or other windows (file dialogs are recognised and not read anyway).
NEVER_START = ("Speichern", "Drucken", "Schließen", "Beenden", "Konto", "Feedback", "Hilfe", "Kontakt", "Schulung", "Neuigkeiten",
               "Makro", "Visual Basic", "Ausführen", "Von Anfang an", "Ab aktueller Folie", "Von aktueller Folie", "Probepräsentation",
               "Neue Zeiten testen", "Bildschirmausschnitt", "Screenshot", "Link", "Hyperlink", "Vorschau", "Wiedergabe",
               "Abspielen", "Vorlesen", "Bildschirmpräsentation starten", "Präsentieren", "Lesemodus", "Leseansicht",
               # File dialogs show the owner's folders: known ones are not even opened.
               "Folien aus Gliederung", "Dieses Gerät", "Audio auf meinem PC", "Aus Datei", "Bild aus Datei", "Durchsuchen",
               # A chart's data opens in Excel (another program).
               "Daten auswählen", "Daten bearbeiten", "Daten in Excel bearbeiten", "Daten aktualisieren",
               # Modes that wait for a click on the slide.
               "Pipette", "Format übertragen",
               # Searches the owner's OneDrive/SharePoint; after it PowerPoint's ribbon could not be read for a while.
               "Folien wiederverwenden")
NEVER_IN = ("aufzeichn", "Aufzeichn", "aufnehm", "Aufnehm", "Bildschirmaufzeichnung")
# Reading these dialogs through UI Automation made PowerPoint crash (filled in when it happens).
NEVER_READ: tuple[str, ...] = ()
# Dialogs that must not be clicked at all: the probe stops (title words).
BLOCKING = ("Anmelden", "Anmeldung", "Datenschutz", "Lizenz", "Neuerungen", "Neuigkeiten", "Willkommen", "Aktivieren",
            "Synchronis", "Konto", "Zustimm", "Microsoft-Konto", "Sign in", "Privacy", "License", "What's New")
IGNORE_CLASSES = ("PPTFrameClass", "MSO_BORDEREFFECT_WINDOW_CLASS")
BACKSTAGE_PAGES = ("Neu", "Informationen", "Exportieren", "Optionen")


class Blocked(RuntimeError):
    """A sign-in, consent, licence or what's-new dialog appeared."""


def cloud(name: str) -> bool:
    return any(name.startswith(c) for c in CLOUD)


def never(name: str) -> bool:
    return any(name.startswith(n) for n in NEVER_START) or any(n in name for n in NEVER_IN) or any(name.startswith(n) for n in NEVER_READ)


# ----- the real-input lock ------------------------------------------------------------------------

def acquire_lock() -> bool:
    try:
        with open(LOCK, "x", encoding="utf-8") as f:
            f.write(f"{LOCK_TAG} {datetime.datetime.now().isoformat(timespec='seconds')}\n")
        return True
    except FileExistsError:
        try:
            age = (time.time() - LOCK.stat().st_mtime) / 60
            who = LOCK.read_text(encoding="utf-8", errors="replace").split(" ")[0][:40]
        except OSError:
            age, who = 0, "?"
        print(f"real-input lock held by {who!r} ({age:.0f} min old)", flush=True)
        return False


def own_lock() -> bool:
    try:
        return LOCK.read_text(encoding="utf-8", errors="replace").startswith(LOCK_TAG)
    except OSError:
        return False


def refresh_lock() -> None:
    if own_lock():
        os.utime(LOCK)


def release_lock() -> None:
    if own_lock():
        LOCK.unlink()


# ----- UI Automation helpers ------------------------------------------------------------------------

def has_expand(el) -> bool:
    try:
        return bool(el.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId))
    except Exception:  # noqa: BLE001
        return False


def expand_state(el) -> int | None:
    try:
        return el.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId).QueryInterface(
            uia.U.IUIAutomationExpandCollapsePattern).CurrentExpandCollapseState
    except Exception:  # noqa: BLE001
        return None


def all_desc(el) -> list[tuple]:
    """Every descendant as (element, description); one retry when the search itself fails (PowerPoint's
    tree changes while menus close)."""
    for attempt in range(2):
        out = []
        try:
            f = el.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
        except Exception:  # noqa: BLE001
            time.sleep(0.5)
            continue
        for i in range(f.Length):
            try:
                e = f.GetElement(i)
                d = uia.describe(e)
            except Exception:  # noqa: BLE001
                continue
            out.append((e, d))
        return out
    return []


KINDS = ("Button", "SplitButton", "MenuItem", "CheckBox", "RadioButton", "ComboBox", "Edit", "ListItem", "TabItem", "Spinner",
         "Slider", "Hyperlink", "TreeItem", "DataItem", "Text", "Group", "List", "Tree", "Menu", "Document")


def listed(pairs) -> list[dict]:
    seen = set()
    out = []
    for _, d in pairs:
        if not d["name"].strip() or d["type"] not in KINDS or len(d["name"]) > 90:
            continue
        key = (d["type"], d["name"], d["rect"][0] // 4, d["rect"][1] // 4)
        if key in seen:
            continue
        seen.add(key)
        item = {"name": d["name"], "type": d["type"], "enabled": d["enabled"]}
        if d["key"]:
            item["key"] = d["key"]
        out.append(item)
    return out


user32 = winapi.user32
user32.WindowFromPoint.argtypes = (winapi.POINT,)
user32.WindowFromPoint.restype = ctypes.c_void_p
user32.GetAncestor.argtypes = (ctypes.c_void_p, ctypes.c_uint)
user32.GetAncestor.restype = ctypes.c_void_p


def window_at(x: float, y: float) -> int:
    h = user32.WindowFromPoint(winapi.POINT(int(x), int(y)))
    return int(user32.GetAncestor(h, 2) or 0) if h else 0  # GA_ROOT


class Inventory:
    def __init__(self):
        self.t = PowerPointTarget(visible=True)
        self.t.remember_ribbon()
        self.pres = self.t.new_presentation(with_window=True, name="inventory")
        self.hwnd = self.t.hwnd or self.t.main_window()
        self.pid = self.t.pid
        self.app = self.t.app
        self.dialogs: dict[str, dict] = {}
        self.notes_toggled = False
        self.reselect = None  # set by Scene.build: puts the context's selection back
        self.rebuild = None   # set by Scene.build: builds the context again
        time.sleep(2.0)
        # A dialog of its own at start (sign-in, what's new, licence): touch nothing, end our PowerPoint, stop.
        for h in self.dialog_windows():
            title = winapi.window_title(h)[:60]
            self.t.close(kill_if_unreachable=True)
            raise Blocked(f"dialog at start: {title!r}")

    # ----- windows and input --------------------------------------------------------------------

    @property
    def root(self):
        """The main window's automation element, fetched anew each time: a kept one went dead during the
        first full run (after the Ansicht tab every search came back empty)."""
        if not winapi.user32.IsWindow(self.hwnd):
            self.hwnd = self.t.main_window()
        return uia.element_for(self.hwnd)

    def ours(self, h: int) -> bool:
        return bool(h) and winapi.window_pid(h) == self.pid

    def dialog_windows(self) -> list[int]:
        def big(h: int) -> bool:
            l, t, r, b = winapi.window_rect(h)
            return r - l > 120 and b - t > 80
        return [h for h in winapi.find_windows(lambda h: winapi.window_pid(h) == self.pid and h != self.hwnd
                                               and winapi.window_class(h) not in IGNORE_CLASSES
                                               and not winapi.window_class(h).startswith("Net UI Tool Window") and big(h))]

    def popups(self) -> list[int]:
        return [h for h in winapi.find_windows(lambda h: winapi.window_pid(h) == self.pid
                                               and winapi.window_class(h).startswith("Net UI Tool Window"))]

    def check_blocking(self) -> None:
        for h in self.dialog_windows():
            title = winapi.window_title(h)
            if any(w.lower() in title.lower() for w in BLOCKING):
                raise Blocked(f"dialog {title[:60]!r}")

    def front(self) -> None:
        if not self.ours(winapi.foreground()):
            try:
                winapi.activate(self.hwnd, settle=0.5)
            except RuntimeError:
                pass
        if not self.ours(winapi.foreground()):
            raise RuntimeError("our PowerPoint is not in front")

    def click(self, x: float, y: float, right: bool = False) -> None:
        """A real click, only on a window of our own PowerPoint."""
        self.front()
        if not self.ours(window_at(x, y)):
            raise RuntimeError(f"point {int(x)},{int(y)} is not on our PowerPoint")
        if right:
            winapi.right_click(x, y)
        else:
            winapi.click(x, y)

    def press(self, combo: str) -> None:
        self.front()
        winapi.press(combo)

    def close_all(self) -> None:
        """Closes menus and dialogs of our PowerPoint: Esc, then Abbrechen / Schließen if a dialog stays."""
        self.check_blocking()
        for _ in range(6):
            dlg = self.dialog_windows()
            if not dlg and not self.popups():
                return
            if dlg:
                winapi.activate(dlg[0])
            winapi.press("esc")
            time.sleep(0.35)
            for h in self.dialog_windows():
                el = uia.find(uia.element_for(h), lambda d: d["type"] == "Button" and d["name"] in ("Abbrechen", "Schließen"), max_depth=6)
                if el is not None:
                    try:
                        uia.invoke(el)
                    except Exception:  # noqa: BLE001
                        pass
                    time.sleep(0.3)
        winapi.activate(self.hwnd)

    def read_dialog(self, h: int) -> dict:
        """A dialog's controls, page by page when it has tabs; of a file dialog only its title."""
        title = winapi.window_title(h)
        self.check_blocking()
        root = uia.element_for(h)
        pairs = all_desc(root)
        if any(d["name"] in ("Dateiname:", "Dateiname") or d["name"].startswith("Adresse:") for _, d in pairs):
            return {"title": title, "file_dialog": True}
        pages = {"": listed(pairs)}
        tabs = [(e, d) for e, d in pairs if d["type"] == "TabItem" and d["name"]]
        for e, d in tabs:
            try:
                uia.select(e)
                time.sleep(0.35)
                pages[d["name"]] = listed(all_desc(root))
            except Exception:  # noqa: BLE001
                pass
        return {"title": title, "pages": pages}

    def panes(self) -> dict[str, object]:
        """Task panes docked left or right, by name (searched in the docks only, see ribbon())."""
        out = {}
        for dock in ("MsoDockRight", "MsoDockLeft"):
            d_el = uia.find(self.root, lambda d, n=dock: d["name"] == n, max_depth=4)
            if d_el is None:
                continue
            for e, d in all_desc(d_el):
                if d["type"] in ("Pane", "Custom", "Window", "Group") and d["name"] and d["rect"][2] > 0 \
                        and not d["name"].startswith("Mso"):
                    out.setdefault(d["name"], e)
                    break
        return out

    def read_pane(self, el) -> list[dict]:
        for _ in range(3):
            for e, d in all_desc(el):
                if d["type"] in ("Group", "Button", "TreeItem") and has_expand(e) and expand_state(e) == 0:
                    try:
                        e.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId).QueryInterface(uia.U.IUIAutomationExpandCollapsePattern).Expand()
                    except Exception:  # noqa: BLE001
                        pass
            time.sleep(0.3)
        return listed(all_desc(el))

    # ----- the ribbon ---------------------------------------------------------------------------

    def ribbon(self):
        """The ribbon's own element, found by walking down from the window (not by a search of the whole
        window: while a task pane such as Form formatieren is open, that search fails in PowerPoint)."""
        el = uia.find(self.root, lambda d: d["name"] == "Ribbon" and d["class"] == "NetUInetpane", max_depth=6)
        return el if el is not None else self.root

    def tabs(self) -> list[dict]:
        return [{"name": d["name"], "key": d["key"]} for _, d in all_desc(self.ribbon())
                if d["type"] == "TabItem" and d["class"] == "NetUIRibbonTab" and d["name"] and d["rect"][2] > 0]

    def lower(self):
        return uia.find(self.ribbon(), lambda d: d["name"] == "Unteres Menüband", max_depth=12)

    def close_panes(self) -> None:
        """Closes task panes (Form formatieren, Auswahl …) by their Schließen button, found by walking."""
        for dock in ("MsoDockRight", "MsoDockLeft"):
            d_el = uia.find(self.root, lambda d, n=dock: d["name"] == n, max_depth=4)
            if d_el is None:
                continue
            for _ in range(3):
                btn = uia.find(d_el, lambda d: d["type"] == "Button" and d["name"] in ("Schließen", "Bereich schließen") and d["rect"][2] > 0, max_depth=10)
                if btn is None:
                    break
                try:
                    uia.invoke(btn)
                except Exception:  # noqa: BLE001
                    break
                time.sleep(0.4)

    def controls(self) -> list[dict]:
        low = self.lower()
        out = []

        def walk(el, group):
            for c in uia.children(el):
                try:
                    d = uia.describe(c)
                except Exception:  # noqa: BLE001
                    continue
                # Ribbon groups are NetUIChunk; an in-ribbon gallery's own Group (NetUIElement) is not one.
                g = d["name"] if d["type"] in ("Group", "ToolBar") and d["name"] and d["class"] != "NetUIElement" else group
                if d["type"] in ("Button", "SplitButton", "MenuItem", "CheckBox", "ComboBox", "Edit", "List", "RadioButton", "Spinner") and d["name"]:
                    item = {"el": c, "group": group, "name": d["name"], "type": d["type"], "enabled": d["enabled"], "rect": d["rect"], "key": d["key"]}
                    if d["type"] == "List":
                        item["gallery"] = [k["name"] for k in (uia.describe(x) for x in uia.children(c)) if k["type"] == "ListItem" and k["name"]]
                    out.append(item)
                elif d["type"] == "DataGrid" and d["name"]:
                    # An in-ribbon gallery (Designs, Übergangseffekte, Schnellformatvorlagen): its visible entries;
                    # the button of the same name next to it opens the whole gallery.
                    item = {"el": c, "group": group, "name": d["name"], "type": "List", "enabled": d["enabled"], "rect": d["rect"], "key": d["key"]}
                    item["gallery"] = [k["name"] for _, k in all_desc(c) if k["type"] in ("ListItem", "DataItem") and k["name"]]
                    out.append(item)
                    continue
                # An unnamed MenuItem is a gallery's frame: its gallery and expand button are inside.
                if d["type"] not in ("MenuItem", "SplitButton", "ComboBox") or (d["type"] == "MenuItem" and not d["name"]):
                    walk(c, g)
        walk(low, "")
        return out

    def open_menu(self, ctl: dict) -> list[dict]:
        c = ctl["el"]
        x, y, w, h = ctl["rect"]
        if ctl["type"] == "SplitButton":
            arrow = [uia.describe(k) for k in uia.children(c) if uia.describe(k)["name"] in ("Weitere Optionen", "Mehr Optionen")]
            if arrow:
                x, y, w, h = arrow[0]["rect"]
            else:
                x, w = x + w - 8, 8
        before = set(self.popups())
        self.click(x + w / 2, y + h / 2)
        time.sleep(1.0)
        self.check_blocking()
        pairs = [p for p in all_desc(c) if p[0] is not None]
        for hwnd in self.popups():
            if hwnd not in before:
                pairs += all_desc(uia.element_for(hwnd))
        items = listed(pairs)
        for e, d in pairs:
            if d["type"] == "MenuItem" and d["name"] and has_expand(e) and not cloud(d["name"]) and not never(d["name"])                     and d["name"] not in ("Weitere Optionen", "Mehr Optionen", ctl["name"]):
                # (The split button's own arrow is not a submenu: a click on it closes the menu.)
                st = expand_state(e)
                if st is None or st == 3:
                    continue
                bx, by, bw, bh = d["rect"]
                if bw <= 0:
                    continue
                pop0 = set(self.popups())
                try:
                    self.click(bx + bw / 2, by + bh / 2)
                except RuntimeError:
                    continue
                time.sleep(0.8)
                sub = all_desc(e)
                for hwnd in self.popups():
                    if hwnd not in pop0:
                        sub += all_desc(uia.element_for(hwnd))
                for it in items:
                    if it["name"] == d["name"]:
                        it["submenu"] = [s for s in listed(sub) if s["name"] != d["name"]]
                        break
                winapi.press("esc")
                time.sleep(0.3)
        self.close_all()
        return items

    def open_dialog(self, ctl: dict, key: str) -> dict | None:
        x, y, w, h = ctl["rect"]
        before_panes = set(self.panes())
        self.click(x + w / 2, y + h / 2)
        end = time.time() + 2.5
        while time.time() < end and not self.dialog_windows():
            time.sleep(0.2)
        got = None
        dlg = self.dialog_windows()
        if dlg:
            time.sleep(0.4)
            got = self.read_dialog(dlg[0])
            self.dialogs[key] = got
        else:
            new = {k: v for k, v in self.panes().items() if k not in before_panes}
            if new:
                name, el = next(iter(new.items()))
                got = {"pane": name, "controls": self.read_pane(el)}
                self.dialogs[key] = got
        self.close_panes()
        self.close_all()
        return got

    def select_tab(self, tab: str):
        # Contextual tabs sit deep in the tree (below their tab group): search all of it.
        hits = [e for e, d in all_desc(self.ribbon()) if d["type"] == "TabItem" and d["name"] == tab and d["rect"][2] > 0]
        if not hits:
            raise RuntimeError(f"tab {tab} not found")
        el = hits[0]
        uia.select(el)
        time.sleep(0.8)
        return el

    def ensure_tab(self, tab: str) -> tuple:
        """Selects the tab; if it is gone (the object lost its selection, PowerPoint in a mode), closes
        what is open, puts the selection back, then rebuilds the situation. Returns (tab element, rebuilt)."""
        for attempt in range(4):
            try:
                return self.select_tab(tab), attempt >= 2
            except RuntimeError:
                self.close_all()
                for _ in range(2):
                    winapi.press("esc")
                    time.sleep(0.2)
                if attempt == 0 and self.reselect:
                    self.reselect()
                elif attempt >= 1 and self.rebuild:
                    self.rebuild()
                time.sleep(0.8 + 2 * attempt)
        raise RuntimeError(f"tab {tab} not found")

    def walk_tab(self, tab: str, list_only: bool = False) -> dict:
        el, _ = self.ensure_tab(tab)
        ctls = self.controls()
        out = []
        for ctl in ctls:
            if self.reselect:
                self.reselect()
            try:
                uia.select(el)
            except Exception:  # noqa: BLE001
                el, rebuilt = self.ensure_tab(tab)
                if rebuilt:
                    # The ribbon was made anew: the control's element and place come from it again.
                    fresh = {(c["group"], c["name"]): c for c in self.controls()}
                    ctl = fresh.get((ctl["group"], ctl["name"]), ctl)
            try:
                out.append(self.walk_control(tab, ctl, list_only))
            except Blocked:
                raise
            except Exception as e:  # noqa: BLE001
                out.append({**{k: ctl[k] for k in ("group", "name", "type", "enabled")}, "error": str(e)[:120]})
                self.close_all()
            time.sleep(0.1)
        return {"controls": out}

    def walk_control(self, tab: str, ctl: dict, list_only: bool) -> dict:
        entry = {k: ctl[k] for k in ("group", "name", "type", "enabled")}
        if ctl.get("key"):
            entry["key"] = ctl["key"]
        if ctl.get("gallery"):
            entry["gallery"] = ctl["gallery"]
        name = ctl["name"]
        menu_like = ctl["type"] in ("SplitButton", "MenuItem", "ComboBox") or (ctl["type"] == "Button" and has_expand(ctl["el"]))
        if cloud(name):
            entry["cloud"] = True
        elif list_only or never(name):
            pass
        elif ctl["enabled"] and menu_like:
            try:
                entry["items"] = self.open_menu(ctl)
            except Blocked:
                raise
            except Exception as e:  # noqa: BLE001
                entry["error"] = str(e)[:120]
                self.close_all()
        elif ctl["enabled"] and ctl["type"] == "Button" and (name.endswith("...") or ctl["rect"][2] <= 24 and ctl["rect"][3] <= 24):
            try:
                if self.open_dialog(ctl, f"{tab}/{ctl['group']}/{name}"):
                    entry["dialog"] = True
            except Blocked:
                raise
            except Exception as e:  # noqa: BLE001
                entry["error"] = str(e)[:120]
                self.close_all()
        print(f"  {ctl['group'][:24]:24} {name[:50]:50} {ctl['type']:12} {len(entry.get('items', []))}", flush=True)
        return entry

    def second_pass(self, tab: str, controls: list[dict]) -> None:
        """Menu entries ending in "...": each opens a dialog; read it (never confirmed)."""
        for c in controls:
            for it in c.get("items", []) + [s for i in c.get("items", []) for s in i.get("submenu", [])]:
                n = it["name"]
                key = f"{tab}/{c['name']}/{n}"
                if not n.endswith(("...", "…")) or cloud(n) or never(n) or key in self.dialogs or not it.get("enabled", True):
                    continue
                if self.reselect:
                    self.reselect()
                try:
                    el, _ = self.ensure_tab(tab)
                except RuntimeError as e:
                    print("  ...", key, "error", str(e)[:80], flush=True)
                    continue
                ctl = next((k for k in self.controls() if k["name"] == c["name"]), None)
                if not ctl:
                    continue
                try:
                    self.open_menu_and_pick(ctl, n, it in c.get("items", []))
                    got = self.after_pick(key)
                    print(f"  ... {key[:70]:70} {'ok' if got else '-'}", flush=True)
                except Blocked:
                    raise
                except Exception as e:  # noqa: BLE001
                    print("  ...", key, "error", str(e)[:80], flush=True)
                self.close_all()
                del el

    def open_menu_and_pick(self, ctl: dict, name: str, top: bool) -> None:
        x, y, w, h = ctl["rect"]
        if ctl["type"] == "SplitButton":
            arrow = [uia.describe(k) for k in uia.children(ctl["el"]) if uia.describe(k)["name"] in ("Weitere Optionen", "Mehr Optionen")]
            if arrow:
                x, y, w, h = arrow[0]["rect"]
        self.click(x + w / 2, y + h / 2)
        time.sleep(0.9)
        pairs = all_desc(ctl["el"])
        for hwnd in self.popups():
            pairs += all_desc(uia.element_for(hwnd))
        if not top:
            for e, d in pairs:
                if d["type"] == "MenuItem" and has_expand(e) and d["rect"][2] > 0 and not never(d["name"]) and not cloud(d["name"]):
                    bx, by, bw, bh = d["rect"]
                    self.click(bx + bw / 2, by + bh / 2)
                    time.sleep(0.7)
                    sub = all_desc(e)
                    for hwnd in self.popups():
                        sub += all_desc(uia.element_for(hwnd))
                    hit = [p for p in sub if p[1]["name"] == name and p[1]["rect"][2] > 0]
                    if hit:
                        pairs = hit
                        break
                    winapi.press("esc")
                    time.sleep(0.2)
        hit = [p for p in pairs if p[1]["name"] == name and p[1]["rect"][2] > 0]
        if hit:
            bx, by, bw, bh = hit[0][1]["rect"]
            self.click(bx + bw / 2, by + bh / 2)

    def after_pick(self, key: str) -> dict | None:
        end = time.time() + 2.5
        while time.time() < end and not self.dialog_windows():
            time.sleep(0.2)
        dlg = self.dialog_windows()
        if dlg:
            time.sleep(0.4)
            self.dialogs[key] = self.read_dialog(dlg[0])
            return self.dialogs[key]
        time.sleep(0.6)
        panes = self.panes()
        if panes:
            name, el = next(iter(panes.items()))
            self.dialogs[key] = {"pane": name, "controls": self.read_pane(el)}
            self.close_panes()
            return self.dialogs[key]
        return None

    # ----- context menus, backstage, keytips ----------------------------------------------------------

    def context_menu(self, x: float, y: float) -> dict:
        before = set(self.popups())
        self.click(x, y, right=True)
        time.sleep(1.0)
        self.check_blocking()
        menus, mini = [], []
        for h in self.popups():
            if h in before:
                continue
            items = listed(all_desc(uia.element_for(h)))
            # The mini toolbar comes with the menu in a window of its own: kept apart.
            (menus if any(i["type"] == "MenuItem" for i in items) else mini).append(items)
        out = {"items": [i for m in menus for i in m]}
        if mini:
            out["mini_toolbar"] = [i for m in mini for i in m]
        self.close_all()
        return out

    def px(self, x_pt: float, y_pt: float) -> tuple[int, int]:
        w = self.pres.Windows(1)
        # Points convert to pixels only while the slide pane is the active one (not after a click into the slide panel).
        try:
            w.Panes(2).Activate()
        except Exception:  # noqa: BLE001
            pass
        return int(w.PointsToScreenPixelsX(x_pt)), int(w.PointsToScreenPixelsY(y_pt))

    def backstage(self) -> dict:
        """Datei: every page (except the ones that show the owner's files or act), and PowerPoint-Optionen."""
        out: dict = {}
        self.close_panes()
        datei = None
        for scope in (self.ribbon(), self.root):
            hits = [e for e, d in all_desc(scope) if (d["name"] == "Datei" or d["name"].startswith("Registerkarte") and "Datei" in d["name"]) and d["type"] in ("Button", "TabItem", "MenuItem")
                    and d["rect"][2] > 0 and d["rect"][1] < 300]
            if hits:
                datei = hits[0]
                break
        if datei is None:
            return {"error": "Datei not found"}
        try:
            uia.invoke(datei)
        except Exception:  # noqa: BLE001
            uia.select(datei)
        time.sleep(2.0)
        self.check_blocking()
        nav = [(e, d) for e, d in all_desc(self.root) if d["type"] in ("ListItem", "TabItem", "Button", "MenuItem") and d["name"]
               and d["rect"][0] < 400 and d["rect"][1] > 100 and d["rect"][2] > 0]
        out["navigation"] = [{"name": d["name"], "type": d["type"]} for _, d in nav]
        names = [d["name"] for _, d in nav]
        out["pages"] = {}
        for name in names:
            # Only these pages are opened: the others show the owner's files and folders (Start, Öffnen,
            # Speichern unter, Kopie speichern), act (Speichern, Drucken, Schließen) or are the cloud's.
            if name not in BACKSTAGE_PAGES:
                out["pages"][name] = {"cloud": True} if cloud(name) else {"skipped": "not opened"}
                continue
            el = uia.find(self.root, lambda d, n=name: d["name"] == n and d["type"] in ("ListItem", "TabItem", "Button", "MenuItem") and d["rect"][0] < 400)
            if el is None:
                continue
            try:
                try:
                    uia.select(el)
                except Exception:  # noqa: BLE001
                    uia.invoke(el)
            except Exception as e:  # noqa: BLE001
                out["pages"][name] = {"error": str(e)[:80]}
                continue
            time.sleep(1.5)
            self.check_blocking()
            dlg = self.dialog_windows()
            if dlg:
                out["pages"][name] = self.options_dialog(dlg[0]) if name == "Optionen" else self.read_dialog(dlg[0])
                self.close_all()
                continue
            out["pages"][name] = [i for i in listed(all_desc(self.root)) if i["name"] not in names]
            print("  backstage", name, len(out["pages"][name]), flush=True)
        self.press("esc")
        time.sleep(0.8)
        self.close_all()
        return out

    def options_dialog(self, h: int) -> dict:
        root = uia.element_for(h)
        cats = [(e, d) for e, d in all_desc(root) if d["type"] == "ListItem" and d["name"] and d["rect"][0] < winapi.window_rect(h)[0] + 400]
        out = {"title": winapi.window_title(h), "pages": {}}
        for e, d in cats:
            try:
                uia.select(e)
                time.sleep(0.8)
                out["pages"][d["name"]] = [i for i in listed(all_desc(root)) if i["name"] not in {c[1]["name"] for c in cats}]
                print("  Optionen", d["name"], len(out["pages"][d["name"]]), flush=True)
            except Exception as ex:  # noqa: BLE001
                out["pages"][d["name"]] = {"error": str(ex)[:80]}
        return out

    def keytips(self, tabs: list[str]) -> dict:
        """Alt, then each main tab's letters: what the keytip windows show."""
        out = {}

        def shown() -> list[str]:
            got = []
            for h in winapi.find_windows(lambda h: winapi.window_pid(h) == self.pid and h != self.hwnd):
                for _, d in all_desc(uia.element_for(h)):
                    if d["name"] and len(d["name"]) <= 4:
                        got.append(d["name"])
                n = winapi.window_title(h)
                if n and len(n) <= 4:
                    got.append(n)
            return got

        self.close_all()
        self.press("alt")
        time.sleep(1.0)
        out["(Alt)"] = shown()
        self.press("esc")
        self.press("esc")
        time.sleep(0.4)
        for t in tabs:
            key = next((x["key"] for x in self.tabs() if x["name"] == t), "")
            letters = key.replace("Alt, ", "").replace(" ", "") if key.startswith("Alt, ") else ""
            if not letters or not all(ch.isalnum() and ch.isascii() for ch in letters):
                out[t] = {"key": key, "skipped": "letters not typeable as virtual keys"}
                continue
            self.press("alt")
            time.sleep(0.6)
            for ch in letters:
                self.press(ch.lower())
                time.sleep(0.3)
            time.sleep(0.8)
            out[t] = {"key": key, "shown": shown()}
            for _ in range(3):
                winapi.press("esc")
                time.sleep(0.2)
            self.close_all()
        return out


# ----- the situations ---------------------------------------------------------------------------------

def _png() -> str:
    from PIL import Image
    p = Path(tempfile.gettempdir()) / "raabe-inventory-ppt.png"
    Image.new("RGB", (240, 160), (70, 130, 180)).save(p)
    return str(p)


def excel_pids() -> set[int]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"], capture_output=True, text=True, errors="replace").stdout
    return {int(p[1]) for p in (line.strip('"').split('","') for line in out.splitlines()) if len(p) > 1 and p[1].isdigit()}


# Excel processes our PowerPoint started for a chart's data (EXCEL.EXE /Automation -Embedding): ours to end.
OWN_EXCEL: set[int] = set()


def end_own_excel() -> None:
    for pid in sorted(OWN_EXCEL & excel_pids()):
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
    OWN_EXCEL.clear()


class Scene:
    """The presentation in the state a context needs, built through COM."""

    def __init__(self, inv: Inventory):
        self.inv = inv

    def reset(self) -> tuple:
        inv = self.inv
        inv.close_all()
        inv.close_panes()
        for _ in range(3):
            inv.press("esc")
            time.sleep(0.15)
        pres = inv.pres
        w = pres.Windows(1)
        if int(w.ViewType) != 9:
            w.ViewType = 9  # ppViewNormal
            time.sleep(0.6)
        while pres.Slides.Count:
            pres.Slides(1).Delete()
        s1 = inv.t.add_slide(pres, 2)
        s1.Shapes(1).TextFrame.TextRange.Text = "Inventur"
        s1.Shapes(2).TextFrame.TextRange.Text = "Erster Punkt" + chr(13) + "Zweiter Punkt"
        s2 = inv.t.add_slide(pres, 7)
        w.View.GotoSlide(2)
        time.sleep(0.3)
        return pres, w, s1, s2

    def build(self, ctx: str):
        """Builds the context; returns the shape it is about (or None). Afterwards the inventory can put
        the selection back before each control (Esc after a menu ends text editing or deselects)."""
        sh = self._build(ctx)
        self.inv.reselect = lambda c=ctx, s=sh: self.select(c, s)
        self.inv.rebuild = lambda c=ctx: self.build(c)
        return sh

    def select(self, ctx: str, sh) -> None:
        try:
            w = self.inv.pres.Windows(1)
            try:
                w.Panes(2).Activate()
            except Exception:  # noqa: BLE001
                pass
            if ctx == "Text":
                w.View.GotoSlide(1)
                sh.TextFrame.TextRange.Characters(1, 5).Select()
            elif ctx == "Tabelle":
                w.View.GotoSlide(2)
                sh.Table.Cell(1, 1).Shape.TextFrame.TextRange.Select()
            elif ctx == "Formel":
                w.View.GotoSlide(2)
                sh.TextFrame.TextRange.Characters(2, 1).Select()
            elif sh is not None:
                w.View.GotoSlide(2)
                sh.Select()
            time.sleep(0.2)
        except Exception:  # noqa: BLE001
            pass

    def _build(self, ctx: str):
        pres, w, s1, s2 = self.reset()
        app = self.inv.app
        sh = None
        if ctx in ("Folie", "Kontext"):
            w.View.GotoSlide(1)
            w.Selection.Unselect()
        elif ctx == "Form":
            sh = s2.Shapes.AddShape(1, 200, 120, 240, 140)
        elif ctx == "Textfeld":
            sh = s2.Shapes.AddTextbox(1, 200, 120, 300, 50)
            sh.TextFrame.TextRange.Text = "Ein Textfeld"
        elif ctx == "Text":
            w.View.GotoSlide(1)
            s1.Shapes(2).TextFrame.TextRange.Characters(1, 5).Select()
            return s1.Shapes(2)
        elif ctx == "Bild":
            sh = s2.Shapes.AddPicture(_png(), 0, -1, 200, 120)
        elif ctx == "Tabelle":
            sh = s2.Shapes.AddTable(3, 3, 120, 120, 600, 150)
            w.View.GotoSlide(2)
            sh.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Zelle"
            sh.Table.Cell(1, 1).Shape.TextFrame.TextRange.Select()
            return sh
        elif ctx == "Diagramm":
            before = excel_pids()
            if before - OWN_EXCEL:
                raise RuntimeError("EXCEL.EXE runs (maybe the owner's): no chart, its data would open there")
            sh = s2.Shapes.AddChart2(-1, 51, 120, 100, 600, 340)
            time.sleep(2.5)
            OWN_EXCEL.update(excel_pids() - before)
            try:
                sh.Chart.ChartData.Workbook.Close()
            except Exception:  # noqa: BLE001
                pass
            time.sleep(1.0)
            try:
                winapi.activate(self.inv.hwnd, settle=0.5)
            except RuntimeError:
                pass
        elif ctx == "SmartArt":
            sh = s2.Shapes.AddSmartArt(app.SmartArtLayouts(1), 120, 100, 500, 300)
        elif ctx == "Gruppe":
            a = s2.Shapes.AddShape(1, 200, 120, 160, 100)
            b = s2.Shapes.AddShape(9, 400, 120, 120, 120)
            sh = s2.Shapes.Range([a.Name, b.Name]).Group()
        elif ctx == "Formel":
            tb = s2.Shapes.AddTextbox(1, 200, 160, 400, 60)
            tb.TextFrame.TextRange.Text = "a^2+b^2=c^2"
            tb.TextFrame.TextRange.Select()
            app.CommandBars.ExecuteMso("EquationInsertNew")
            time.sleep(0.6)
            app.CommandBars.ExecuteMso("EquationProfessional")
            time.sleep(0.6)
            tb.TextFrame.TextRange.Characters(2, 1).Select()
            return tb
        elif ctx == "Folienmaster":
            w.ViewType = 2  # ppViewSlideMaster
        elif ctx == "Handzettelmaster":
            w.ViewType = 4  # ppViewHandoutMaster
        elif ctx == "Notizenmaster":
            w.ViewType = 5  # ppViewNotesMaster
        if sh is not None:
            w.View.GotoSlide(2)
            sh.Select()
        time.sleep(0.6)
        return sh


CONTEXTS = ["Folie", "Form", "Textfeld", "Text", "Bild", "Tabelle", "Diagramm", "SmartArt", "Gruppe", "Formel",
            "Folienmaster", "Handzettelmaster", "Notizenmaster"]
PARTS = CONTEXTS + ["Kontextmenüs", "Backstage", "Zugriffstasten"]
SKIP_TABS = ("Datei", "Hilfe", "Acrobat")
LIST_ONLY_TABS = ("Aufzeichnen",)  # recording: the controls are listed, nothing is opened
# With nothing selected most of Start (and all of Animationen) is grey and its menus do not open: walked again
# while text is edited (Start) or a shape is selected (Animationen).
WALK_AGAIN = {"Text": ["Start"], "Form": ["Animationen"]}


def thumb_rect(inv: Inventory, slide_left: int) -> tuple[int, int, int, int]:
    l, t, r, b = winapi.window_rect(inv.hwnd)
    x0, y0, w, h = l + 20, t + 270, max(60, slide_left - l - 40), 900
    g = winapi.Grabber(x0, y0, w, h)
    try:
        raw = g.grab()
    finally:
        g.close()
    xs, ys = [], []
    for y in range(0, h, 2):
        row = y * w * 4
        for x in range(0, w, 2):
            b_, g_, r_ = raw[row + x * 4], raw[row + x * 4 + 1], raw[row + x * 4 + 2]
            if r_ > 150 and g_ < 110 and b_ < 90:
                xs.append(x)
                ys.append(y)
    if len(xs) < 20:
        raise RuntimeError("no slide thumbnail found")
    return x0 + min(xs), y0 + min(ys), max(xs) - min(xs), max(ys) - min(ys)


def save(data: dict) -> None:
    from feellab.sanitize import clean
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    RESULT.write_text(json.dumps(clean(data), ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    refresh_lock()


def context_menus(inv: Inventory, scene: Scene, data: dict) -> None:
    cm = data.setdefault("context_menus", {})
    targets = ["Folienminiatur", "Leerer Bereich im Folienbereich", "Platzhalterrahmen", "Text in Bearbeitung", "Form", "Bild",
               "Tabelle", "Diagramm", "Gruppe", "Leerer Folienbereich", "Notizenbereich", "SmartArt", "Textfeld"]
    for name in targets:
        if name in cm and not cm[name].get("error"):
            continue
        try:
            if name in ("Folienminiatur", "Leerer Bereich im Folienbereich", "Leerer Folienbereich", "Notizenbereich", "Platzhalterrahmen",
                        "Text in Bearbeitung"):
                pres, w, s1, s2 = scene.reset()
                if name == "Folienminiatur" or name == "Leerer Bereich im Folienbereich":
                    # PowerPoint's slide panel is not in UI Automation: the current slide's thumbnail is
                    # found by its highlight frame (Office's accent red) left of the slide.
                    w.View.GotoSlide(1)
                    time.sleep(0.6)
                    x, y, ww, hh = thumb_rect(inv, int(w.PointsToScreenPixelsX(0)))
                    pt = (x + ww / 2, y + hh / 2) if name == "Folienminiatur" else (x + ww / 2, y + 4 * hh + 200)
                elif name == "Leerer Folienbereich":
                    w.View.GotoSlide(2)
                    w.Selection.Unselect()
                    pt = inv.px(800, 470)
                elif name == "Notizenbereich":
                    if not inv.app.CommandBars.GetPressedMso("ShowNotes"):
                        inv.app.CommandBars.ExecuteMso("ShowNotes")
                        inv.notes_toggled = True
                        time.sleep(0.6)
                    # The notes pane is not in UI Automation either: it lies between the slide and the status bar.
                    bar = [d for _, d in all_desc(inv.root) if d["type"] == "StatusBar" and d["rect"][2] > 0]
                    if not bar:
                        raise RuntimeError("status bar not found")
                    pt = ((inv.px(0, 0)[0] + inv.px(960, 0)[0]) / 2, bar[0]["rect"][1] - 40)
                elif name == "Platzhalterrahmen":
                    w.View.GotoSlide(1)
                    ph = s1.Shapes(2)
                    ph.Select()
                    time.sleep(0.4)
                    pt = inv.px(ph.Left + 0.5, ph.Top + ph.Height / 2)
                else:  # Text in Bearbeitung
                    w.View.GotoSlide(1)
                    tr = s1.Shapes(2).TextFrame.TextRange
                    tr.Characters(1, 5).Select()
                    ln = tr.Lines(1, 1)
                    pt = inv.px(ln.BoundLeft + 30, ln.BoundTop + ln.BoundHeight / 2)
            else:
                sh = scene.build(name if name != "Textfeld" else "Textfeld")
                if name == "Tabelle":
                    tb = sh.Table
                    pt = inv.px(sh.Left + tb.Columns(1).Width + tb.Columns(2).Width / 2, sh.Top + tb.Rows(1).Height + tb.Rows(2).Height / 2)
                elif name == "Diagramm":
                    pt = inv.px(sh.Left + 8, sh.Top + 8)
                elif name == "Gruppe":
                    it = sh.GroupItems(1)
                    pt = inv.px(it.Left + it.Width / 2, it.Top + it.Height / 2)
                else:
                    pt = inv.px(sh.Left + sh.Width / 2, sh.Top + sh.Height / 2)
            try_front(inv)
            cm[name] = inv.context_menu(*pt)
            print("context menu", name, len(cm[name]["items"]), flush=True)
        except Blocked:
            raise
        except Exception as e:  # noqa: BLE001
            cm[name] = {"items": [], "error": str(e)[:120]}
            print("context menu", name, "failed:", str(e)[:100], flush=True)
            inv.close_all()
        save(data)


def try_front(inv: Inventory) -> None:
    try:
        inv.front()
    except RuntimeError:
        pass


def end(inv: Inventory | None, kill: bool = False) -> dict:
    if inv is None:
        return {}
    try:
        inv.close_all()
    except Exception:  # noqa: BLE001
        pass
    if inv.notes_toggled:
        try:
            inv.app.CommandBars.ExecuteMso("ShowNotes")
        except Exception:  # noqa: BLE001
            pass
    done = inv.t.close(kill_if_unreachable=kill)
    end_own_excel()
    return done


def run(pick: list[str], data: dict) -> None:
    inv = Inventory()
    scene = Scene(inv)
    # Tabs recorded without an error are not walked again (a rerun continues where it stopped).
    done = {k for k, v in data["tabs"].items() if not v.get("error")}
    walked = {k.split("/", 1)[1] for k in done}
    blocked = False
    try:
        # The tabs of a plain slide: the other contexts walk only the tabs they add.
        scene.build("Folie")
        try_front(inv)
        time.sleep(1.0)
        base: list[str] | None = [t["name"] for t in inv.tabs()]
        for ctx in [p for p in pick if p in CONTEXTS]:
            try:
                scene.build(ctx)
            except Exception as e:  # noqa: BLE001
                data["contexts"][ctx] = {"error": str(e)[:120]}
                print(ctx, "setup failed:", str(e)[:100], flush=True)
                save(data)
                continue
            try_front(inv)
            time.sleep(1.0)
            tabs = inv.tabs()
            if not tabs:
                # The automation tree of our PowerPoint went dead: a new PowerPoint, the same situation.
                print(ctx, "no tabs readable: new PowerPoint", flush=True)
                data["dialogs"].update(inv.dialogs)
                end(inv)
                inv = Inventory()
                scene = Scene(inv)
                scene.build(ctx)
                try_front(inv)
                time.sleep(1.0)
                tabs = inv.tabs()
            names = [t["name"] for t in tabs]
            if ctx == "Folie":
                base = names
            data["contexts"][ctx] = {"tabs": tabs, "own": [t for t in names if t not in (base or [])] if ctx != "Folie" else []}
            walk = [t for t in names if (ctx == "Folie" or t not in (base or [])) and t not in SKIP_TABS and t not in walked]
            walk += [t for t in WALK_AGAIN.get(ctx, []) if t in names and t not in walk]
            print(ctx, "tabs:", names, "walk:", walk, flush=True)
            for t in walk:
                if f"{ctx}/{t}" in done:
                    continue
                print("TAB", t, flush=True)
                try:
                    got = inv.walk_tab(t, list_only=t in LIST_ONLY_TABS)
                    data["tabs"][f"{ctx}/{t}"] = got
                    walked.add(t)
                    if t not in LIST_ONLY_TABS:
                        try:
                            inv.second_pass(t, got["controls"])
                        except Blocked:
                            raise
                        except Exception as e:  # noqa: BLE001
                            # The tab's walk is kept; only its "..." entries are incomplete.
                            got["second_pass_error"] = str(e)[:100]
                except Blocked:
                    raise
                except Exception as e:  # noqa: BLE001
                    print("  tab failed:", str(e)[:100], flush=True)
                    data["tabs"][f"{ctx}/{t}"] = {"controls": [], "error": str(e)[:100]}
                    # A dead PowerPoint, or one whose ribbon can no longer be read: a new one, and the tab once more.
                    data["dialogs"].update(inv.dialogs)
                    save(data)
                    end(inv)
                    inv = Inventory()
                    scene = Scene(inv)
                    try:
                        scene.build(ctx)
                        try_front(inv)
                        time.sleep(1.0)
                        got = inv.walk_tab(t, list_only=t in LIST_ONLY_TABS)
                        data["tabs"][f"{ctx}/{t}"] = got
                        walked.add(t)
                        print("  tab done in a new PowerPoint", flush=True)
                    except Blocked:
                        raise
                    except Exception as e2:  # noqa: BLE001
                        print("  tab failed again:", str(e2)[:100], flush=True)
                data["dialogs"].update(inv.dialogs)
                save(data)
                try:
                    scene.build(ctx)
                except Exception as e:  # noqa: BLE001
                    print("  setup failed:", str(e)[:100], flush=True)
                    end(inv)
                    inv = Inventory()
                    scene = Scene(inv)
                    scene.build(ctx)
                time.sleep(0.5)
            save(data)
        if "Kontextmenüs" in pick:
            context_menus(inv, scene, data)
        if "Backstage" in pick and not data.get("backstage", {}).get("pages"):
            scene.build("Folie")
            data["backstage"] = inv.backstage()
            save(data)
        if "Zugriffstasten" in pick:
            scene.build("Folie")
            main_tabs = [t for t in (base or []) if t not in ("Acrobat",)]
            data["keytips"] = inv.keytips(main_tabs)
            save(data)
    except Blocked:
        blocked = True
        raise
    finally:
        data["dialogs"].update(inv.dialogs)
        save(data)
        print("PowerPoint:", end(inv, kill=blocked), flush=True)


def alive(inv: Inventory) -> bool:
    try:
        _ = inv.pres.Slides.Count
        return True
    except Exception:  # noqa: BLE001
        return False


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    pick = [a for a in sys.argv[1:] if not a.startswith("-")] or PARTS
    if powerpoint_pids():
        print("POWERPNT.EXE is running (the owner's): not touched, no PowerPoint probe now", flush=True)
        return 76
    if not acquire_lock():
        return 75
    data = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    for k in ("tabs", "dialogs", "contexts"):
        data.setdefault(k, {})
    data["built_by"] = "PowerPoint (Microsoft 365, German UI)"
    data["date"] = DATE
    try:
        run(pick, data)
    except OwnerPowerPointRunning as e:
        print("skipped:", e, flush=True)
        return 76
    except Blocked as e:
        data["blocked"] = str(e)[:80]
        save(data)
        print("BLOCKED:", e, flush=True)
        return 77
    finally:
        release_lock()
    return 0


if __name__ == "__main__":
    sys.exit(main())
