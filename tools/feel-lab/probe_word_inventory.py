"""Probe: Word's ribbon as an inventory of its dropdowns (Step 11, D-165), for check_picture_controls.py. In a Word of
our own (DispatchEx; the window placed off-screen, never activated, no real input) every ribbon tab is selected
through UI Automation, every control that drops down is opened through ExpandCollapse (never invoked), the entries
of its open list are read with their UI Automation types (ListItem: a gallery entry; MenuItem: a menu entry) and
submenus are opened one level deep. The contextual tabs come from a shape, a picture and a table made by COM in the
probe's own document. Labels and control types only, no descriptions (D-058). Word's open dropdowns take the
foreground even off-screen: each tab is read under the real-input lock (released between tabs), every dropdown is
collapsed until no popup of ours is left, and the result is saved after each tab (a stopped run goes on from there).

    .venv/Scripts/python.exe probe_word_inventory.py

Result: results/2026-11-12-word-inventory.json, in the form of the Excel and PowerPoint inventories.
"""
from __future__ import annotations

import json
import tempfile
import time
from pathlib import Path

import probe_excel_chartpane as P
from feellab import targets, uia, winapi
from feellab.inputlock import real_input

HERE = Path(__file__).parent
OUT = HERE / "results" / "2026-11-12-word-inventory.json"
U = uia.U
UIA = uia._uia
user32 = __import__("ctypes").windll.user32
TABS = ["Start", "Einfügen", "Zeichnen", "Entwurf", "Layout", "Referenzen", "Sendungen", "Überprüfen", "Ansicht"]
CONTROL = ("Button", "SplitButton", "MenuItem", "ComboBox")
ENTRY = ("ListItem", "MenuItem", "Button", "RadioButton", "CheckBox", "SplitButton", "DataItem")
SKIP_NAMES = ("Zeile nach oben", "Zeile nach unten", "Seite nach oben", "Seite nach unten", "Öffnen", "Schließen", "Weitere Optionen")


def popup_entries(root, pid: int, skip_rect) -> list[dict]:
    """Entries of the open dropdown: NetUI shows them in a NetUIToolWindow (main tree and a window of its own)."""
    out: list[dict] = []
    seen: set = set()
    holders = []
    cond = UIA.CreatePropertyCondition(U.UIA_ClassNamePropertyId, "NetUIToolWindow")
    try:
        f = root.FindAll(U.TreeScope_Descendants, cond)
        holders += [f.GetElement(i) for i in range(f.Length)]
    except Exception:  # noqa: BLE001
        pass
    for h in winapi.find_windows(lambda h: winapi.window_pid(h) == pid):
        if user32.IsWindowVisible(h) and winapi.window_class(h) == "Net UI Tool Window":
            holders.append(uia.element_for(h))
    for hd in holders:
        try:
            f = hd.FindAll(U.TreeScope_Descendants, UIA.CreateTrueCondition())
        except Exception:  # noqa: BLE001
            continue
        for i in range(f.Length):
            e = f.GetElement(i)
            d = P.d_of(e)
            if not d or d["type"] not in ENTRY or not d["name"].strip() or d["name"] in SKIP_NAMES:
                continue
            key = (d["name"], tuple(d["rect"]))
            if key in seen or d["rect"][2] <= 0:
                continue
            seen.add(key)
            out.append({"name": d["name"], "type": d["type"], "_el": e})
    return out


def open_and_read(root, pid: int, el, depth: int) -> list[dict]:
    p = P.expander(el)
    if p is None:
        return []
    items: list[dict] = []
    for _ in range(2):
        try:
            p.Expand()
        except Exception:  # noqa: BLE001
            return []
        time.sleep(0.8)
        items = popup_entries(root, pid, None)
        if items:
            break
        try:
            p.Collapse()
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.3)
    out = []
    for it in items[:220]:
        o = {"name": it["name"], "type": it["type"]}
        # One level of submenus (Formkontur > Striche …), read while the parent list is open.
        if depth == 0 and it["type"] in ("MenuItem", "SplitButton") and P.expander(it["_el"]) is not None:
            sub = []
            q = P.expander(it["_el"])
            try:
                q.Expand()
                time.sleep(0.7)
                names = {x["name"] for x in items}
                sub = [{"name": x["name"], "type": x["type"]} for x in popup_entries(root, pid, None) if x["name"] not in names]
                q.Collapse()
                time.sleep(0.3)
            except Exception:  # noqa: BLE001
                pass
            o["submenu"] = sub
        out.append(o)
    close_popups(p, pid)
    return out


def close_popups(p, pid: int) -> None:
    """Collapses the dropdown until no popup window of our Word is visible (an open popup holds the foreground)."""
    for _ in range(4):
        try:
            p.Collapse()
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.3)
        left = [h for h in winapi.find_windows(lambda h: winapi.window_pid(h) == pid)
                if user32.IsWindowVisible(h) and winapi.window_class(h) == "Net UI Tool Window"]
        if not left:
            return


def read_tab(root, pid: int, name: str) -> list[dict]:
    tab = uia.find(root, lambda d: d["type"] == "TabItem" and d["name"] == name)
    if tab is None:
        return []
    try:
        uia.select(tab)
    except Exception:  # noqa: BLE001
        return []
    time.sleep(1.0)
    ribbon = uia.find(root, lambda d: d["name"] == "Unteres Menüband")
    if ribbon is None:
        return []
    controls = []
    for g in P.kids(ribbon):
        gd = P.d_of(g)
        if not gd:
            continue
        stack = [(g, gd["name"])]
        found = []
        while stack:
            e, grp = stack.pop()
            for k in reversed(P.kids(e)):
                kd = P.d_of(k)
                if not kd:
                    continue
                if kd["type"] in CONTROL and P.expander(k) is not None and kd["enabled"]:
                    found.append((k, kd, grp))
                elif kd["type"] in ("Group", "Pane", "Custom", "ToolBar", "List"):
                    stack.append((k, kd["name"] if kd["type"] == "Group" and kd["name"] else grp))
        for k, kd, grp in found:
            items = open_and_read(root, pid, k, 0)
            controls.append({"group": grp, "name": kd["name"], "type": kd["type"], "items": items})
            print(f"  {name} / {grp} / {kd['name']}: {len(items)}", flush=True)
    return controls


def make_png() -> str:
    from PIL import Image
    p = Path(tempfile.gettempdir()) / "raabe-word-inventory.png"
    Image.new("RGB", (120, 80), (91, 155, 213)).save(p)
    return str(p)


def main() -> None:
    w = targets.WordTarget("short", visible=False)
    w.app.Visible = True
    winapi.place_quietly(w.hwnd, -4200, 0, 1900, 1200)
    pid = winapi.window_pid(w.hwnd)
    print("word pid", pid, flush=True)
    data: dict = {"built_by": "Word (Microsoft 365, German UI), own instance off-screen, UI Automation only", "date": time.strftime("%Y-%m-%d"), "tabs": {}}
    if OUT.exists():
        data = json.loads(OUT.read_text(encoding="utf-8"))

    def save() -> None:
        # Entries that name the owner (the signed-in account in a menu) are left out, as in the Excel inventory.
        from feellab.sanitize import clean
        text = json.dumps(clean(data), ensure_ascii=False, indent=1)
        from feellab.sanitize import private_names
        low = text.lower()
        if any(n in low for n in private_names()) or "c:\\\\users" in low:
            print("result holds a private name or path; not saved")
        else:
            OUT.write_text(text + "\n", encoding="utf-8", newline="\n")

    def tab(key: str, name: str, setup=None) -> None:
        """One tab under the real-input lock (Word's open dropdowns take the foreground), saved at once."""
        if key in data["tabs"]:
            return
        with real_input(f"word inventory {name}"):
            obj = setup() if setup else None
            time.sleep(1.0)
            data["tabs"][key] = {"controls": read_tab(root, pid, name)}
            if obj is not None:
                try:
                    obj.Delete()
                except Exception:  # noqa: BLE001
                    pass
        save()

    try:
        root = uia.element_for(w.hwnd)
        for t in TABS:
            tab(f"Dokument/{t}", t)
        doc = w.doc

        # Contextual tabs: a shape, a picture and a table of the probe's own document.
        def shape():
            s = doc.Shapes.AddShape(1, 100, 100, 120, 80)
            s.Select()
            return s

        def picture():
            s = doc.InlineShapes.AddPicture(make_png(), False, True, doc.Range(0, 0))
            s.Select()
            return s

        tab("Form/Formformat", "Formformat", shape)
        tab("Bild/Bildformat", "Bildformat", picture)
        tbl = doc.Tables.Add(doc.Range(0, 0), 3, 3)
        tbl.Cell(1, 1).Select()
        tab("Tabelle/Tabellenentwurf", "Tabellenentwurf")
        tab("Tabelle/Layout", "Layout")
    finally:
        w.close()


if __name__ == "__main__":
    main()
