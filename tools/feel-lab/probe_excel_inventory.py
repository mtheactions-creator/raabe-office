"""Probe: Excel's complete inventory (D-066). In an Excel of our own (visible, real clicks) every ribbon
tab is walked, including the tabs that appear only for a selected object (chart, table, picture,
shape, text box, PivotTable, sparkline, slicer, SmartArt); every menu, gallery and split button is
opened (submenus one level deep), every dialog launcher and every entry ending in "..." is opened and
its controls listed (each page of a dialog with tabs), then closed with Esc or Abbrechen, never with
OK. Labels and control types only: no descriptions or tooltips (D-058). Buttons that act at once are
listed, not pressed; cloud features are listed and marked, not opened.

    .venv/Scripts/python.exe probe_excel_inventory.py [context …]

Result: results/<date>-excel-inventory.json (labels only), written after each context; the full
recording with Office's descriptions in captures/ (not committed).
"""
from __future__ import annotations

import datetime
import json
import sys
import tempfile
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.calc_catalogue import Case, ExcelSide

HERE = Path(__file__).parent
# The full recording (with Office's descriptions) stays in captures/; results/ gets labels only.
OUT = HERE / "captures" / "2026-11-05-excel-inventory-full.json"
RESULT = HERE / "results" / "2026-11-05-excel-inventory.json"


def save(data: dict) -> None:
    from feellab.sanitize import clean
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    RESULT.write_text(json.dumps(clean(data), ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))

# Cloud and online features (D-066): listed, never opened.
CLOUD = ("Freigeben", "Teilen", "Copilot", "Daten analysieren", "Add-Ins abrufen", "Add-Ins", "Online-Bilder", "Piktogramme",
         "Symbole", "3D-Modelle", "Übersetzen", "Intelligente Suche", "Aktien", "Geografie", "Formulare", "Datentypen", "Vorlesen",
         "Diktieren", "Vertraulichkeit", "Ideen", "Python", "Power BI", "Neue Abfrage", "Daten abrufen", "Visio-Datenschnellansicht",
         "Bing", "Personen", "Microsoft Forms", "Office-Add-Ins", "Meine Add-Ins", "Store")
# Never pressed even in the second pass (they act at once or leave Excel).
NEVER = ("Speichern", "Drucken", "Schließen", "Beenden", "Konto", "Feedback", "Hilfe", "Kontakt", "Schulung", "Neuigkeiten",
         "Makro aufzeichnen", "Makros", "Visual Basic", "Ausführen",
         # Reading these dialogs through UI Automation made our Excel crash (recorded 2026-11-05).
         "Suchen...", "Ersetzen...", "Gehe zu...")
DIALOG_CLASSES = ("#32770", "bosa_sdm_XL9", "NUIDialog", "bosa_sdm_msword", "WebView2Dialog_v2")


def cloud(name: str) -> bool:
    return any(name.startswith(c) for c in CLOUD)


def has_expand(el) -> bool:
    try:
        p = el.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId)
        return bool(p)
    except Exception:  # noqa: BLE001
        return False


def expand(el) -> None:
    p = el.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId)
    p.QueryInterface(uia.U.IUIAutomationExpandCollapsePattern).Expand()


def all_desc(el) -> list[tuple]:
    """Every descendant with a name, as (element, description)."""
    out = []
    try:
        f = el.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
        for i in range(f.Length):
            e = f.GetElement(i)
            try:
                d = uia.describe(e)
            except Exception:  # noqa: BLE001
                continue
            out.append((e, d))
    except Exception:  # noqa: BLE001
        pass
    return out


KINDS = ("Button", "SplitButton", "MenuItem", "CheckBox", "RadioButton", "ComboBox", "Edit", "ListItem", "TabItem", "Spinner",
         "Slider", "Hyperlink", "TreeItem", "DataItem", "Text", "Group", "List", "Tree", "Menu", "Document")


def listed(pairs) -> list[dict]:
    seen = set()
    out = []
    for _, d in pairs:
        if not d["name"] or d["type"] not in KINDS or len(d["name"]) > 90:
            continue
        key = (d["type"], d["name"], d["rect"][0] // 4, d["rect"][1] // 4)
        if key in seen:
            continue
        seen.add(key)
        out.append({"name": d["name"], "type": d["type"], "enabled": d["enabled"]})
    return out


class Inventory:
    def __init__(self):
        self.side = ExcelSide()
        self.x = self.side.x
        self.root = self.side.root
        self.pid = self.x.pid
        self.dialogs: dict[str, list] = {}

    # ----- windows ------------------------------------------------------------------------------

    def dialog_windows(self) -> list[int]:
        def big(h: int) -> bool:
            l, t, r, b = winapi.window_rect(h)
            return r - l > 120 and b - t > 80
        return [h for h in winapi.find_windows(lambda h: winapi.window_pid(h) == self.pid and h != self.x.hwnd
                                               and winapi.window_class(h) in DIALOG_CLASSES and big(h))]

    def popups(self) -> list[int]:
        return [h for h in winapi.find_windows(lambda h: winapi.window_pid(h) == self.pid and winapi.window_class(h) == "Net UI Tool Window")]

    def close_all(self) -> None:
        """Closes menus and dialogs of our Excel: Esc, then Abbrechen / Schließen if a dialog stays."""
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
        winapi.activate(self.x.hwnd)

    def read_dialog(self, h: int) -> dict:
        """A dialog's controls, page by page when it has tabs."""
        root = uia.element_for(h)
        title = winapi.window_title(h)
        pages = {"": listed(all_desc(root))}
        tabs = [(e, d) for e, d in all_desc(root) if d["type"] == "TabItem" and d["name"]]
        for e, d in tabs:
            try:
                uia.select(e)
                time.sleep(0.35)
                pages[d["name"]] = listed(all_desc(root))
            except Exception:  # noqa: BLE001
                pass
        return {"title": title, "pages": pages}

    def panes(self) -> dict[str, object]:
        """Task panes docked in the main window (Format..., Auswahl, Zwischenablage ...), by name."""
        out = {}
        for e, d in all_desc(self.root):
            if d["type"] in ("Pane", "Custom", "Window") and d["name"] and d["name"].endswith(("formatieren", "bereich", "Zwischenablage", "Felder", "Feldliste")) and d["rect"][2] > 0:
                out[d["name"]] = e
        return out

    def read_pane(self, el) -> list[dict]:
        """A task pane with every section expanded (through the automation pattern, not clicks)."""
        for _ in range(3):
            for e, d in all_desc(el):
                if d["type"] in ("Group", "Button", "TreeItem") and has_expand(e):
                    try:
                        p = e.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId).QueryInterface(uia.U.IUIAutomationExpandCollapsePattern)
                        if p.CurrentExpandCollapseState == 0:
                            p.Expand()
                    except Exception:  # noqa: BLE001
                        pass
            time.sleep(0.3)
        return listed(all_desc(el))

    # ----- the ribbon ---------------------------------------------------------------------------

    def tabs(self) -> list[str]:
        return [d["name"] for _, d in all_desc(self.root) if d["type"] == "TabItem" and d["name"] and d["rect"][1] < 300]

    def lower(self):
        return uia.find(self.root, lambda d: d["name"] == "Unteres Menüband", max_depth=12)

    def controls(self) -> list[dict]:
        """The selected tab's groups and controls with their elements."""
        low = self.lower()
        out = []

        def walk(el, group):
            for c in uia.children(el):
                try:
                    d = uia.describe(c)
                except Exception:  # noqa: BLE001
                    continue
                g = d["name"] if d["type"] in ("Group", "ToolBar") and d["name"] else group
                if d["type"] in ("Button", "SplitButton", "MenuItem", "CheckBox", "ComboBox", "Edit", "List", "RadioButton", "Spinner") and d["name"]:
                    out.append({"el": c, "group": group, "name": d["name"], "type": d["type"], "enabled": d["enabled"], "rect": d["rect"]})
                if d["type"] not in ("MenuItem", "SplitButton", "ComboBox"):
                    walk(c, g)
        walk(low, "")
        return out

    def open_menu(self, ctl: dict) -> list[dict]:
        """Clicks a menu, split arrow or gallery; lists its entries and one level of submenus."""
        c = ctl["el"]
        x, y, w, h = ctl["rect"]
        if ctl["type"] == "SplitButton":
            arrow = [uia.describe(k) for k in uia.children(c) if uia.describe(k)["name"] in ("Weitere Optionen", "Mehr Optionen")]
            if arrow:
                x, y, w, h = arrow[0]["rect"]
            else:
                x, w = x + w - 8, 8
        before = set(self.popups())
        winapi.click(x + w / 2, y + h / 2)
        time.sleep(1.0)
        pairs = [p for p in all_desc(c) if p[0] is not None]
        for hwnd in self.popups():
            if hwnd not in before:
                pairs += all_desc(uia.element_for(hwnd))
        items = listed(pairs)
        # Submenus: entries with their own expansion, opened by a click (they open on click, not hover).
        for e, d in pairs:
            if d["type"] == "MenuItem" and d["name"] and has_expand(e) and not cloud(d["name"]):
                try:
                    st = e.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId).QueryInterface(uia.U.IUIAutomationExpandCollapsePattern).CurrentExpandCollapseState
                except Exception:  # noqa: BLE001
                    continue
                if st == 3:  # leaf
                    continue
                bx, by, bw, bh = d["rect"]
                pop0 = set(self.popups())
                winapi.click(bx + bw / 2, by + bh / 2)
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
        """Clicks a dialog launcher or "..." control; reads the dialog or pane it opens, then closes it."""
        x, y, w, h = ctl["rect"]
        before_panes = set(self.panes())
        winapi.click(x + w / 2, y + h / 2)
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
                close = uia.find(el, lambda d: d["type"] == "Button" and d["name"] == "Schließen", max_depth=6)
                if close is not None:
                    try:
                        uia.invoke(close)
                    except Exception:  # noqa: BLE001
                        pass
        self.close_all()
        return got

    def walk_tab(self, tab: str) -> dict:
        el = uia.find(self.root, lambda d, t=tab: d["type"] == "TabItem" and d["name"] == t)
        uia.select(el)
        time.sleep(0.8)
        ctls = self.controls()
        out = []
        for ctl in ctls:
            entry = {k: ctl[k] for k in ("group", "name", "type", "enabled")}
            name = ctl["name"]
            if cloud(name):
                entry["cloud"] = True
            elif ctl["enabled"] and ctl["type"] in ("SplitButton", "MenuItem", "ComboBox") and not any(name.startswith(n) for n in NEVER):
                try:
                    entry["items"] = self.open_menu(ctl)
                except Exception as e:  # noqa: BLE001
                    entry["error"] = str(e)[:120]
                    self.close_all()
            elif ctl["enabled"] and ctl["type"] == "Button" and (name.endswith("...") or ctl["rect"][2] <= 24 and ctl["rect"][3] <= 24) \
                    and not any(name.startswith(n) for n in NEVER):
                try:
                    got = self.open_dialog(ctl, f"{tab}/{ctl['group']}/{name}")
                    if got:
                        entry["dialog"] = True
                except Exception as e:  # noqa: BLE001
                    entry["error"] = str(e)[:120]
                    self.close_all()
            out.append(entry)
            print(f"  {ctl['group'][:24]:24} {name[:50]:50} {ctl['type']:12} {len(entry.get('items', []))}", flush=True)
            # The tab may have changed (some menus switch it back): select it again.
            uia.select(el)
            time.sleep(0.1)
        return {"controls": out}

    def second_pass(self, tab: str, controls: list[dict]) -> None:
        """Menu entries ending in "...": each opens a dialog; read it (never confirmed)."""
        for c in controls:
            for it in c.get("items", []) + [s for i in c.get("items", []) for s in i.get("submenu", [])]:
                n = it["name"]
                key = f"{tab}/{c['name']}/{n}"
                if not n.endswith("...") or cloud(n) or key in self.dialogs or any(n.startswith(x) for x in NEVER) or not it.get("enabled", True):
                    continue
                el = uia.find(self.root, lambda d, t=tab: d["type"] == "TabItem" and d["name"] == t)
                uia.select(el)
                time.sleep(0.5)
                ctl = next((k for k in self.controls() if k["name"] == c["name"]), None)
                if not ctl:
                    continue
                try:
                    self.open_menu_and_pick(ctl, n, it in c.get("items", []))
                    got = self.after_pick(key)
                    print(f"  ... {key[:70]:70} {'ok' if got else '-'}", flush=True)
                except Exception as e:  # noqa: BLE001
                    print("  ...", key, "error", e, flush=True)
                self.close_all()

    def open_menu_and_pick(self, ctl: dict, name: str, top: bool) -> None:
        x, y, w, h = ctl["rect"]
        if ctl["type"] == "SplitButton":
            arrow = [uia.describe(k) for k in uia.children(ctl["el"]) if uia.describe(k)["name"] in ("Weitere Optionen", "Mehr Optionen")]
            if arrow:
                x, y, w, h = arrow[0]["rect"]
        winapi.click(x + w / 2, y + h / 2)
        time.sleep(0.9)
        pairs = all_desc(ctl["el"])
        for hwnd in self.popups():
            pairs += all_desc(uia.element_for(hwnd))
        if not top:
            # Open the submenu that holds it first.
            for e, d in pairs:
                if d["type"] == "MenuItem" and has_expand(e):
                    bx, by, bw, bh = d["rect"]
                    winapi.click(bx + bw / 2, by + bh / 2)
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
            winapi.click(bx + bw / 2, by + bh / 2)

    def after_pick(self, key: str) -> dict | None:
        end = time.time() + 2.5
        while time.time() < end and not self.dialog_windows():
            time.sleep(0.2)
        dlg = self.dialog_windows()
        if dlg:
            time.sleep(0.4)
            self.dialogs[key] = self.read_dialog(dlg[0])
            return self.dialogs[key]
        return None


# ----- the situations that bring contextual tabs ---------------------------------------------------

def _png() -> str:
    from PIL import Image
    p = Path(tempfile.gettempdir()) / "raabe-inventory.png"
    Image.new("RGB", (120, 80), (70, 130, 180)).save(p)
    return str(p)


def setup(inv: Inventory, ctx: str) -> None:
    x = inv.x
    # Whatever a tab left open (a dialog, edit mode) goes first.
    inv.close_all()
    for _ in range(3):
        winapi.activate(x.hwnd)
        winapi.press("esc")
        time.sleep(0.2)
    ws = x.wb.Worksheets(1)
    for s in list(ws.Shapes):
        s.Delete()
    for lo in list(ws.ListObjects):
        lo.Unlist()
    ws.Cells.Clear()
    data = [("Name", "Wert"), ("Anna", 3), ("Ben", 5), ("Cem", 2), ("Dora", 4)]
    for r, row in enumerate(data, start=1):
        for c, v in enumerate(row, start=1):
            ws.Cells(r, c).Value = v
    ws.Range("A1").Select()
    if ctx == "Diagramm":
        ws.Shapes.AddChart2(-1, 51, 200, 20, 360, 220).Select()
    elif ctx == "Tabelle":
        ws.ListObjects.Add(1, ws.Range("A1:B5"), None, 1)
        ws.Range("A2").Select()
    elif ctx == "Bild":
        ws.Shapes.AddPicture(_png(), False, True, 200, 20, 120, 80).Select()
    elif ctx == "Form":
        ws.Shapes.AddShape(1, 200, 20, 120, 80).Select()
    elif ctx == "Textfeld":
        ws.Shapes.AddTextbox(1, 200, 20, 120, 60).Select()
    elif ctx == "Sparkline":
        ws.Range("D2").SparklineGroups.Add(1, "B2:B5")
        ws.Range("D2").Select()
    elif ctx == "PivotTable":
        pc = x.wb.PivotCaches().Create(1, ws.Range("A1:B5"))
        pt = pc.CreatePivotTable(ws.Range("E3"), "Pivot1")
        pt.AddDataField(pt.PivotFields("Wert"))
        pt.PivotFields("Name").Orientation = 1
        ws.Range("E4").Select()
    elif ctx == "Datenschnitt":
        lo = ws.ListObjects.Add(1, ws.Range("A1:B5"), None, 1)
        sc = x.wb.SlicerCaches.Add2(lo, "Name")
        sc.Slicers.Add(ws, None, "Name1", "Name", 30, 300, 140, 160).Shape.Select()
    elif ctx == "SmartArt":
        ws.Shapes.AddSmartArt(x.app.SmartArtLayouts(1), 200, 20, 300, 200).Select()


CONTEXTS = ["Blatt", "Diagramm", "Tabelle", "Bild", "Form", "Textfeld", "Sparkline", "PivotTable", "Datenschnitt", "SmartArt"]


def main() -> None:
    pick = sys.argv[1:] or CONTEXTS
    data = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {"tabs": {}, "dialogs": {}, "contexts": {}}
    inv = Inventory()
    done = set(data["tabs"])
    try:
        inv.side.reset(Case("inventory", {}, [], "x"))
        base = None
        for ctx in pick:
            setup(inv, ctx)
            winapi.activate(inv.x.hwnd)
            time.sleep(1.0)
            tabs = inv.tabs()
            if ctx == "Blatt":
                base = tabs
            elif base is None:
                base = data["contexts"].get("Blatt", {}).get("tabs", [])
            walk = [t for t in tabs if t not in base or ctx == "Blatt"]
            # Add-in tabs (Acrobat …) are not Excel's; Hilfe is online.
            walk = [t for t in walk if t not in ("Datei", "Hilfe", "Acrobat")]
            data["contexts"][ctx] = {"tabs": tabs, "own": [t for t in tabs if t not in base] if ctx != "Blatt" else []}
            print(ctx, "tabs:", tabs, flush=True)
            for t in walk:
                if f"{ctx}/{t}" in done:
                    continue
                print("TAB", t, flush=True)
                try:
                    got = inv.walk_tab(t)
                    data["tabs"][f"{ctx}/{t}"] = got
                    inv.second_pass(t, got["controls"])
                except Exception as e:  # noqa: BLE001
                    # Our Excel may die reading a dialog: note it, start a new one, go on with the next tab.
                    print("  tab failed:", str(e)[:100], flush=True)
                    data["tabs"].setdefault(f"{ctx}/{t}", {"controls": [], "error": str(e)[:100]})
                    try:
                        inv.side.close()
                    except Exception:  # noqa: BLE001
                        pass
                    inv = Inventory()
                    inv.side.reset(Case("inventory", {}, [], "x"))
                    setup(inv, ctx)
                data["dialogs"].update(inv.dialogs)
                save(data)
                # Each tab starts from the same situation again (a new Excel when this one is stuck).
                try:
                    setup(inv, ctx)
                except Exception as e:  # noqa: BLE001
                    print("  setup failed:", str(e)[:100], flush=True)
                    try:
                        inv.side.close()
                    except Exception:  # noqa: BLE001
                        pass
                    inv = Inventory()
                    inv.side.reset(Case("inventory", {}, [], "x"))
                    setup(inv, ctx)
                time.sleep(0.6)
    finally:
        inv.close_all()
        data["dialogs"].update(inv.dialogs)
        save(data)
        inv.side.close()


if __name__ == "__main__":
    main()
