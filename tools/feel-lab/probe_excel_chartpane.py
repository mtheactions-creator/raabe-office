"""Probe: Excel's chart Format pane in detail (for RaabeCalc's pane). In an Excel of our own (DispatchEx, window
placed off-screen, no real input) charts of seven types are made; each chart element is selected through COM and
the pane is opened with `CommandBars.ExecuteMso("ChartFormatSelection")`. The pane is then read through UI
Automation only: the option sets ("Achsenoptionen" / "Textoptionen"), the icon tabs, every section (expanded through
the automation pattern) with its controls (type, label, value, enabled), the entries of every dropdown (opened and
closed through ExpandCollapse) and, for every radio group, which controls each choice reveals (selected through
SelectionItem; the first choice is put back afterwards). Labels only, no tooltips or descriptions (D-058).

A second part measures Excel's chart drawing through its own PDF export (`ExportAsFixedFormat`, page zoom 100 %,
so PDF points are chart points): tick mark lengths, the label offset, display unit labels and shadow presets.

    .venv/Scripts/python.exe probe_excel_chartpane.py [panes [chart ...]] [measure]

Result: results/2026-11-09-excel-chart-pane.json (merged: a run with some charts keeps the others).
"""
from __future__ import annotations

import json
import os
import re
import sys
import tempfile
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
OUT = HERE / "results" / "2026-11-09-excel-chart-pane.json"
U = uia.U
UIA = uia._uia

# ----- UI Automation helpers -------------------------------------------------------------------------------------


def pattern(el, pid, iface):
    try:
        p = el.GetCurrentPattern(pid)
        if not p:
            return None
        return p.QueryInterface(iface)
    except Exception:  # noqa: BLE001
        return None


def expander(el):
    return pattern(el, U.UIA_ExpandCollapsePatternId, U.IUIAutomationExpandCollapsePattern)


def by_class(root, cls: str):
    cond = UIA.CreatePropertyCondition(U.UIA_ClassNamePropertyId, cls)
    try:
        return root.FindFirst(U.TreeScope_Descendants, cond)
    except Exception:  # noqa: BLE001
        return None


def kids(el) -> list:
    try:
        return uia.children(el)
    except Exception:  # noqa: BLE001
        return []


def d_of(el) -> dict | None:
    try:
        return uia.describe(el)
    except Exception:  # noqa: BLE001
        return None


def value_of(el, d: dict):
    t = d["type"]
    if t in ("Edit", "ComboBox"):
        p = pattern(el, U.UIA_ValuePatternId, U.IUIAutomationValuePattern)
        if p:
            try:
                return p.CurrentValue
            except Exception:  # noqa: BLE001
                return None
    if t == "CheckBox" or (t == "Button" and d["class"] == "NetUIStickyButton"):
        p = pattern(el, U.UIA_TogglePatternId, U.IUIAutomationTogglePattern)
        if p:
            return {0: False, 1: True, 2: "gemischt"}.get(p.CurrentToggleState)
    if t in ("RadioButton", "ListItem"):
        p = pattern(el, U.UIA_SelectionItemPatternId, U.IUIAutomationSelectionItemPattern)
        if p:
            return bool(p.CurrentIsSelected)
    if t in ("Slider", "Spinner"):
        p = pattern(el, U.UIA_RangeValuePatternId, U.IUIAutomationRangeValuePattern)
        if p:
            return p.CurrentValue
    return None


def range_of(el):
    p = pattern(el, U.UIA_RangeValuePatternId, U.IUIAutomationRangeValuePattern)
    if p:
        try:
            return [p.CurrentMinimum, p.CurrentMaximum]
        except Exception:  # noqa: BLE001
            return None
    return None


# ----- the pane ----------------------------------------------------------------------------------------------------


class Pane:
    def __init__(self, x: ExcelTarget):
        self.x = x
        self.root = uia.element_for(x.hwnd)

    def body(self):
        """The pane's content (class NetUInetpane), named like the pane ("Achse formatieren")."""
        return by_class(self.root, "NetUInetpane")

    def title(self) -> str:
        b = self.body()
        return (d_of(b) or {}).get("name", "") if b else ""

    def open(self, title_before: str | None = None) -> bool:
        try:
            self.x.app.CommandBars.ExecuteMso("ChartFormatSelection")
        except Exception:  # noqa: BLE001
            return False
        for _ in range(20):
            time.sleep(0.25)
            t = self.title()
            if t and t != title_before:
                return True
        return bool(self.title())

    def sets(self) -> list[tuple[str, object]]:
        """The option sets: TabItems in the header (Achsenoptionen, Textoptionen)."""
        b = self.body()
        out = []
        head = by_class(b, "NetUIPanViewer") if b else None
        if head is None:
            return out
        cond = UIA.CreatePropertyCondition(U.UIA_ControlTypePropertyId, U.UIA_TabItemControlTypeId)
        f = head.FindAll(U.TreeScope_Descendants, cond)
        for i in range(f.Length):
            e = f.GetElement(i)
            out.append((e.CurrentName, e))
        return out

    def tabs(self) -> list[tuple[str, object]]:
        b = self.body()
        g = by_class(b, "NetUIGalleryContainer") if b else None
        if g is None:
            return []
        return [(e.CurrentName, e) for e in kids(g) if (d_of(e) or {}).get("type") == "ListItem"]

    def scroller(self):
        b = self.body()
        return by_class(b, "NetUIScrollViewer") if b else None

    def sections(self) -> list[tuple[str, object]]:
        """Top-level groups of the scroll viewer that start with a header button of their own name."""
        s = self.scroller()
        if s is None:
            return []
        out = []
        for e in kids(s):
            d = d_of(e)
            if not d or d["type"] != "Group":
                continue
            ks = kids(e)
            h = d_of(ks[0]) if ks else None
            if h and h["type"] == "Button" and h["class"] == "NetUIRibbonButton":
                out.append((h["name"] or d["name"], e))
            else:
                out.append(("", e))
        return out

    def section(self, title: str):
        for t, e in self.sections():
            if t == title:
                return e
        return None


def expand_section(pane: Pane, title: str) -> None:
    e = pane.section(title)
    if e is None or len(kids(e)) > 1:
        return
    ks = kids(e)
    if not ks:
        return
    # The header button is a toggle (NetUIRibbonButton with TogglePattern).
    tp = pattern(ks[0], U.UIA_TogglePatternId, U.IUIAutomationTogglePattern)
    try:
        if tp is not None:
            if tp.CurrentToggleState == 0:
                tp.Toggle()
        elif expander(e) is not None:
            expander(e).Expand()
        else:
            uia.invoke(ks[0])
    except Exception:  # noqa: BLE001
        pass
    for _ in range(10):
        time.sleep(0.15)
        e = pane.section(title)
        if e is not None and len(kids(e)) > 1:
            return


CONTROLS = ("Edit", "CheckBox", "RadioButton", "ComboBox", "Button", "SplitButton", "MenuItem", "Spinner", "Slider",
            "ListItem", "Hyperlink")


def walk(el, group: str, out: list, skip_first_header: bool) -> None:
    ks = kids(el)
    for i, c in enumerate(ks):
        d = d_of(c)
        if not d:
            continue
        t = d["type"]
        if skip_first_header and i == 0 and t == "Button" and d["class"] == "NetUIRibbonButton":
            continue
        if t in ("Image", "ScrollBar", "Thumb"):
            continue
        if t == "Group":
            walk(c, d["name"] or group, out, False)
            continue
        if t == "Text":
            out.append({"t": "Text", "label": d["name"], "group": group, "_el": c, "_d": d})
            continue
        if t in CONTROLS:
            item = {"t": t, "label": d["name"], "group": group, "_el": c, "_d": d}
            if t == "ComboBox":
                v = value_of(c, d)
                if v is None:
                    for k in kids(c):
                        kd = d_of(k)
                        if kd and kd["type"] == "Edit":
                            v = value_of(k, kd)
                item["value"] = v
            out.append(item)
            if t in ("ComboBox", "Spinner"):
                continue
            # Some buttons hold a label and an image only; menu items can hold sub-buttons.
            if t in ("MenuItem", "SplitButton"):
                walk(c, group, out, False)
            continue
        walk(c, group, out, False)


COLOR = re.compile(r"^(\S*farbe|Farbe|Vordergrund|Hintergrund)\s+(.+)$")
STOPS = re.compile(r"^Schieberegler mit mehreren Stopps: (\d+) \D+ (\d+)")


def normalize(items: list) -> None:
    """Office names some dropdown buttons with their current value ("Strichtyp Uni", "Füllfarbe RGB(255, 255, 255)"):
    split into label and value, so that panes compare by their controls."""
    prev_text = None
    for i in items:
        if i["t"] == "Text":
            prev_text = i["label"]
            continue
        lab = i["label"]
        if i["t"] in ("MenuItem", "SplitButton", "Button"):
            m = COLOR.match(lab)
            if m:
                i["label"], i["value"], i["picker"] = m.group(1), m.group(2), "Farbe"
            elif prev_text and lab.startswith(prev_text + " "):
                i["label"], i["value"] = prev_text, lab[len(prev_text) + 1:]
            elif lab in SKIP_ENTRIES:
                i["picker"] = "Farbe"
        if i["t"] == "Slider":
            m = STOPS.match(lab)
            if m:
                i["label"], i["value"] = "Farbverlaufstopps", f"{m.group(2)} Stopps"
        prev_text = None


def finish(items: list, section_title: str) -> list[dict]:
    """Drops texts that only label a control or a group; adds values, ranges and enabled state."""
    normalize(items)
    names = {i["label"] for i in items if i["t"] != "Text"}
    groups = {i["group"] for i in items}
    out = []
    for i in items:
        d = i["_d"]
        if i["t"] == "Text" and (i["label"] in names or i["label"] in groups or i["label"] == section_title or not i["label"]):
            continue
        if i["t"] == "Button" and d["class"] == "NetUIRibbonButton" and i["label"] == section_title:
            continue
        o = {"t": i["t"], "label": i["label"]}
        if i["group"] and i["group"] != section_title:
            o["group"] = i["group"]
        v = i.get("value", value_of(i["_el"], d)) if i["t"] != "Text" else None
        if v is not None:
            o["value"] = v
        if i["t"] in ("Edit", "Spinner", "Slider"):
            r = range_of(i["_el"])
            if r:
                o["range"] = r
        if not d["enabled"]:
            o["disabled"] = True
        if i["t"] in ("Button", "SplitButton", "MenuItem") and expander(i["_el"]) is not None:
            o["menu"] = True
        if i.get("picker"):
            o["picker"] = i["picker"]
        out.append(o)
    return out


def read_section(pane: Pane, title: str) -> list[dict]:
    e = pane.section(title) if title else pane.scroller()
    if e is None:
        return []
    items: list = []
    walk(e, title, items, bool(title))
    return finish(items, title)


def raw_section(pane: Pane, title: str) -> list:
    e = pane.section(title) if title else pane.scroller()
    if e is None:
        return []
    items: list = []
    walk(e, title, items, bool(title))
    return items


ENTRY_TYPES = ("ListItem", "MenuItem", "Button", "RadioButton", "CheckBox", "DataItem", "SplitButton")
NOISE = ("Zeile nach oben", "Zeile nach unten", "Seite nach oben", "Seite nach unten", "Position", "Öffnen", "Schließen")


def names_below(el) -> list[str]:
    """Entry names below an element in document order; named groups (gallery categories) as "[name]"."""
    out: list[str] = []
    try:
        f = el.FindAll(U.TreeScope_Descendants, UIA.CreateTrueCondition())
        n = f.Length if f else 0
    except Exception:  # noqa: BLE001
        return out
    for i in range(n):
        d = d_of(f.GetElement(i))
        if not d or not d["name"].strip():
            continue
        if d["type"] == "Group" and d["class"] == "NetUIGalleryCategoryContainer":
            out.append(f"[{d['name'].strip()}]")
        elif d["type"] in ENTRY_TYPES and d["name"] not in NOISE and d["name"] not in out:
            out.append(d["name"])
    return out


def popup_entries(x: ExcelTarget, root, before: set[int]) -> list[str]:
    """Entries of an open dropdown: NetUI shows them in the main window's tree (a NetUIToolWindow menu)."""
    cond = UIA.CreatePropertyCondition(U.UIA_ClassNamePropertyId, "NetUIToolWindow")
    names: list[str] = []
    try:
        f = root.FindAll(U.TreeScope_Descendants, cond)
        for i in range(f.Length):
            names += [n for n in names_below(f.GetElement(i)) if n not in names]
    except Exception:  # noqa: BLE001
        pass
    if names:
        return names
    for h in [h for h in winapi.find_windows(lambda h: winapi.window_pid(h) == x.pid) if h not in before]:
        if winapi.window_class(h) == "Net UI Tool Window":
            names += [n for n in names_below(uia.element_for(h)) if n not in names]
    return names


SKIP_ENTRIES = ("Farbe", "Füllfarbe", "Rahmenfarbe", "Linienfarbe", "Umrissfarbe", "Textfarbe")


def dropdown_entries(pane: Pane, el, d: dict) -> list[str] | None:
    """Opens a dropdown through ExpandCollapse, reads its entries, closes it again."""
    p = expander(el)
    if p is None:
        return None
    x = pane.x
    before = set(winapi.find_windows(lambda h: winapi.window_pid(h) == x.pid))
    names: list[str] = []
    # A gallery opened for the first time shows its entries to UI Automation only from the second opening on.
    for attempt in range(3):
        try:
            p.Expand()
        except Exception:  # noqa: BLE001
            return None
        for wait in (0.6, 0.8):
            time.sleep(wait)
            # Entries listed below the control itself (combo boxes), else in the dropdown's menu.
            names = names_below(el)
            if len(names) < 2:
                names = popup_entries(x, pane.root, before)
            if names:
                break
        if names:
            break
        try:
            p.Collapse()
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.5)
    try:
        p.Collapse()
    except Exception:  # noqa: BLE001
        pass
    time.sleep(0.3)
    # A popup still open: collapse through the control once more.
    left = [h for h in winapi.find_windows(lambda h: winapi.window_pid(h) == x.pid) if h not in before
            and winapi.user32.IsWindowVisible(h)]
    if left:
        try:
            p.Collapse()
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.3)
    return names


ENTRY_CACHE: dict[str, list[str]] = {}  # gallery label -> entries (galleries are the same in every pane)


def wants_entries(c: dict) -> bool:
    if c.get("disabled") or c.get("picker"):
        return False
    return c["t"] == "ComboBox" or bool(c.get("menu"))


def entries_of(pane: Pane, c: dict, el, d: dict) -> list[str] | None:
    galleryish = c["t"] != "ComboBox"
    key = f'{CUR[0]}|{c.get("group", "")}|{c["label"]}'
    if galleryish and key in ENTRY_CACHE:
        return ENTRY_CACHE[key]
    ents = dropdown_entries(pane, el, d)
    if not ents and c["t"] == "ComboBox":
        btn = next((k for k in kids(el) if (d_of(k) or {}).get("type") == "Button"), None)
        if btn is not None:
            ents = dropdown_entries(pane, btn, d_of(btn) or {})
    if ents and galleryish:
        ENTRY_CACHE[key] = ents
    return ents


CUR = [""]


def deep_section(pane: Pane, title: str) -> dict:
    """Controls with values, dropdown entries, and what each radio choice reveals."""
    CUR[0] = title
    items = raw_section(pane, title)
    controls = finish(items, title)
    # Dropdown entries (colour pickers excluded: their swatches are the theme colours).
    for c, i in zip([c for c in controls], [i for i in items if not _dropped(i, items, title)]):
        if wants_entries(c):
            ents = entries_of(pane, c, i["_el"], i["_d"])
            if ents:
                c["entries"] = ents
    out = {"title": title, "controls": controls}
    # Radio groups: select each choice and record what changes.
    radios: dict[str, list[str]] = {}
    for c in controls:
        if c["t"] == "RadioButton" and not c.get("disabled"):
            radios.setdefault(c.get("group", title), []).append(c["label"])
    reveal = {}
    base = [(c["t"], c["label"], bool(c.get("disabled"))) for c in controls]
    for g, labels in radios.items():
        if len(labels) < 2:
            continue
        first = next((c["label"] for c in controls if c["t"] == "RadioButton" and c.get("group", title) == g and c.get("value")), None)
        for lab in labels:
            if lab == first:
                continue
            if not select_radio(pane, title, g, lab):
                continue
            time.sleep(0.5)
            now = read_section(pane, title)
            have = [(c["t"], c["label"], bool(c.get("disabled"))) for c in now]
            added = [c for c in now if (c["t"], c["label"], bool(c.get("disabled"))) not in base]
            removed = [{"t": t, "label": l} for (t, l, dis) in base if (t, l, dis) not in have]
            # Dropdowns that appear only with this choice.
            if added:
                raw = raw_section(pane, title)
                finish(raw, title)  # normalizes the labels
                for a in added:
                    if wants_entries(a):
                        hit = next((i for i in raw if i["t"] == a["t"] and i["label"] == a["label"]), None)
                        if hit is not None:
                            ents = entries_of(pane, a, hit["_el"], hit["_d"])
                            if ents:
                                a["entries"] = ents
            r = {}
            if added:
                r["adds"] = added
            if removed:
                r["removes"] = removed
            reveal[f"{g} > {lab}" if g and g != title else lab] = r or "keine Änderung"
        if first:
            select_radio(pane, title, g, first)
            time.sleep(0.4)
    if reveal:
        out["choices"] = reveal
    return out


def _dropped(i, items, title) -> bool:
    names = {j["label"] for j in items if j["t"] != "Text"}
    groups = {j["group"] for j in items}
    d = i["_d"]
    if i["t"] == "Text" and (i["label"] in names or i["label"] in groups or i["label"] == title or not i["label"]):
        return True
    if i["t"] == "Button" and d["class"] == "NetUIRibbonButton" and i["label"] == title:
        return True
    return False


def select_radio(pane: Pane, title: str, group: str, label: str) -> bool:
    for i in raw_section(pane, title):
        if i["t"] == "RadioButton" and i["label"] == label and (i["group"] or title) == (group or title):
            e = i["_el"]
            try:
                uia.select(e)
                return True
            except Exception:  # noqa: BLE001
                try:
                    uia.invoke(e)
                    return True
                except Exception:  # noqa: BLE001
                    return False
    return False


def select_item(e) -> None:
    try:
        uia.select(e)
    except Exception:  # noqa: BLE001
        uia.invoke(e)


# ----- reading a whole pane, with sharing of identical sections ---------------------------------------------------

SEEN: dict[tuple, str] = {}  # section signature -> pane key where it was recorded in full


def signature(set_label: str, tab: str, title: str, controls: list[dict]) -> tuple:
    return (set_label, tab, title, tuple((c["t"], c["label"], c.get("group", "")) for c in controls))


def read_pane(pane: Pane, key: str) -> dict:
    title = pane.title()
    result = {"title": title, "sets": []}
    sets = pane.sets() or [("", None)]
    for set_label, _ in sets:
        if set_label:
            e = next((e for n, e in pane.sets() if n == set_label), None)
            if e is not None:
                select_item(e)
                time.sleep(0.6)
        s_out = {"label": set_label, "tabs": []}
        tab_names = [n for n, _ in pane.tabs()] or [""]
        for tab in tab_names:
            if tab:
                e = next((e for n, e in pane.tabs() if n == tab), None)
                if e is None:
                    continue
                select_item(e)
                time.sleep(0.6)
            t_out = {"label": tab, "sections": []}
            titles = [t for t, _ in pane.sections()] or [""]
            for st in titles:
                if st:
                    expand_section(pane, st)
                quick = read_section(pane, st)
                sig = signature(set_label, tab, st, quick)
                if sig in SEEN:
                    t_out["sections"].append({"title": st, "same_as": SEEN[sig]})
                    continue
                full = deep_section(pane, st)
                SEEN[sig] = key
                t_out["sections"].append(full)
            s_out["tabs"].append(t_out)
        result["sets"].append(s_out)
    return result


# ----- charts and elements -----------------------------------------------------------------------------------------

CHARTS = {"column": 51, "bar": 57, "line": 65, "pie": 5, "scatter": -4169, "area": 1, "combo": 51}


def fill_data(ws) -> None:
    rows = [("Monat", "A", "B"), ("Jan", 4, 6), ("Feb", -2, 5), ("Mär", 7, 3), ("Apr", 5, 8)]
    for r, row in enumerate(rows, 1):
        for c, v in enumerate(row, 1):
            ws.Cells(r, c).Value = v
    xy = [("X", "Y1", "Y2"), (1, 4, 6), (2.5, -2, 5), (4, 7, 3), (6, 5, 8)]
    for r, row in enumerate(xy, 1):
        for c, v in enumerate(row, 5):
            ws.Cells(r, c).Value = v


def make_chart(ws, kind: str):
    sh = ws.Shapes.AddChart2(-1, CHARTS[kind], 20, 120, 480, 300)
    ch = sh.Chart
    ch.SetSourceData(ws.Range("E1:G5" if kind == "scatter" else "A1:C5"))
    if kind == "combo":
        s2 = ch.SeriesCollection(2)
        s2.ChartType = 65
        s2.AxisGroup = 2
    ch.HasTitle = True
    ch.HasLegend = True
    s1 = ch.SeriesCollection(1)
    s1.HasDataLabels = True
    if kind != "pie":
        s1.Trendlines().Add()
        s1.ErrorBar(1, 1, 1, 1.0)  # xlY, xlBoth, xlFixedValue
        ch.Axes(2).HasTitle = True
        ch.Axes(2).HasMajorGridlines = True
    return sh, ch


def elements(kind: str, ch) -> list[tuple[str, object]]:
    out = [("Diagrammbereich", lambda: ch.ChartArea.Select()),
           ("Zeichnungsfläche", lambda: ch.PlotArea.Select()),
           ("Diagrammtitel", lambda: ch.ChartTitle.Select()),
           ("Legende", lambda: ch.Legend.Select())]
    if kind != "pie":
        out += [("Achse 2 (Wert)", lambda: ch.Axes(2).Select()),
                ("Achse 1 (Kategorie bzw. X)", lambda: ch.Axes(1).Select()),
                ("Hauptgitternetzlinien", lambda: ch.Axes(2).MajorGridlines.Select()),
                ("Achsentitel", lambda: ch.Axes(2).AxisTitle.Select())]
    out += [("Datenreihe", lambda: ch.SeriesCollection(1).Select()),
            ("Datenpunkt", lambda: ch.SeriesCollection(1).Points(2).Select()),
            ("Datenbeschriftungen", lambda: ch.SeriesCollection(1).DataLabels().Select()),
            ("Datenbeschriftung (eine)", lambda: ch.SeriesCollection(1).Points(2).DataLabel.Select())]
    if kind != "pie":
        out += [("Trendlinie", lambda: ch.SeriesCollection(1).Trendlines(1).Select()),
                ("Fehlerindikatoren", lambda: ch.SeriesCollection(1).ErrorBars.Select())]
    if kind == "combo":
        out += [("Sekundärachse (Wert)", lambda: ch.Axes(2, 2).Select()),
                ("Datenreihe 2 (Linie, Sekundärachse)", lambda: ch.SeriesCollection(2).Select())]
    if kind == "line":
        out += [("Datenreihe 2", lambda: ch.SeriesCollection(2).Select())]
    return out


EXPECT = (("Diagrammbereich", "diagrammbereich"), ("Zeichnungsfläche", "zeichnungsfläche"), ("Diagrammtitel", "diagrammtitel"),
          ("Legende", "legende"), ("Gitternetz", "gitternetz"), ("Achsentitel", "achsentitel"), ("Datenreihe", "datenreihe"),
          ("Datenpunkt", "datenpunkt"), ("Datenbeschriftungen", "datenbeschriftungen"),
          ("Datenbeschriftung", "datenbeschriftung formatieren"), ("Trendlinie", "trendlinie"), ("Fehler", "fehler"),
          ("Achse", "achse formatieren"))


def expected(name: str) -> str:
    """A word the pane's title holds for this element (to know the pane has followed the selection)."""
    return next((w for k, w in EXPECT if k.lower() in name.lower()), "formatieren")


def load() -> dict:
    if OUT.exists():
        return json.loads(OUT.read_text(encoding="utf-8"))
    return {"meta": {}, "panes": {}, "measures": {}}


def save(data: dict, part: str) -> None:
    """Writes one part ("panes" or "measures") into the file as it is now (the two parts can run at once)."""
    from feellab.sanitize import private_names
    cur = load()
    cur.setdefault(part, {}).update(data[part])
    cur["meta"] = meta(cur)
    text = json.dumps(cur, ensure_ascii=False, indent=1)
    low = text.lower()
    bad = [n for n in private_names() if n in low] + (["c:\\\\users"] if "c:\\\\users" in low else [])
    if bad:
        raise SystemExit("result holds a private name or path; not saved")
    OUT.write_text(text + "\n", encoding="utf-8", newline="\n")


def start_excel() -> ExcelTarget:
    x = ExcelTarget(visible=False, place=False)
    # Off-screen and not activated: the probe sends no input, and other sessions keep the screen.
    winapi.place_quietly(x.hwnd, -4200, 0, 1800, 2000)
    x.app.Visible = True
    winapi.place_quietly(x.hwnd, -4200, 0, 1800, 2000)
    print("excel pid", x.pid, flush=True)
    return x


def run_panes(kinds: list[str]) -> None:
    data = load()
    # Sections recorded by an earlier run count as seen.
    for key, pane in data["panes"].items():
        for s in pane.get("sets", []):
            for t in s["tabs"]:
                for sec in t["sections"]:
                    if "controls" in sec:
                        SEEN[signature(s["label"], t["label"], sec["title"], sec["controls"])] = key
    x = start_excel()
    pane = Pane(x)
    try:
        ws = x.wb.Worksheets(1)
        fill_data(ws)
        for kind in kinds:
            sh, ch = make_chart(ws, kind)
            for name, select in elements(kind, ch):
                key = f"{kind}/{name}"
                old = data["panes"].get(key)
                if old and "error" not in old and "note" not in old:
                    continue  # recorded by an earlier run
                t0 = time.time()
                want = expected(name)
                ok = False
                err = None
                for attempt in range(3):
                    try:
                        if attempt:
                            # The chart object first, then the element (a title after the plot area needed it).
                            sh.Select()
                            time.sleep(0.4)
                        select()
                    except Exception as e:  # noqa: BLE001
                        err = "COM: " + str(e)[:100]
                        break
                    time.sleep(0.8)
                    if not pane.open(None):
                        continue
                    # The pane follows the selection with a delay: wait for the title of this element.
                    for _ in range(12):
                        if want in pane.title().lower():
                            ok = True
                            break
                        time.sleep(0.25)
                    if ok:
                        break
                if err or not pane.title():
                    data["panes"][key] = {"error": err or "pane did not open"}
                    continue
                time.sleep(0.5)
                res = read_pane(pane, key)
                if not ok:
                    res["note"] = f"pane title does not name the element (wanted '{want}')"
                # A pane identical to one recorded before (all its sections shared): a reference only.
                same = None
                for k, v in data["panes"].items():
                    if k != key and "sets" in v and v.get("title") == res["title"] and _refs(v, k) == _refs(res, key):
                        same = k
                        break
                if same and all("same_as" in sec for s in res["sets"] for t in s["tabs"] for sec in t["sections"]):
                    res = {"title": res["title"], "same_as": same}
                data["panes"][key] = res
                print(f"{key}: {res.get('title')} {time.time() - t0:.0f}s", flush=True)
                save(data, "panes")
            sh.Delete()
    finally:
        save(data, "panes")
        x.close()


def _refs(p: dict, key: str) -> list:
    """Each section of a pane as the key of the pane that holds it in full."""
    return [(s["label"], t["label"], sec["title"], sec.get("same_as", key)) for s in p["sets"] for t in s["tabs"] for sec in t["sections"]]


def run_markers() -> None:
    """The series pane of line and scatter charts has a second page in "Füllung und Linie" ("Markierung")."""
    data = load()
    x = start_excel()
    pane = Pane(x)
    try:
        ws = x.wb.Worksheets(1)
        fill_data(ws)
        for kind in ("line", "scatter"):
            sh, ch = make_chart(ws, kind)
            sh.Select()
            time.sleep(0.4)
            ch.SeriesCollection(1).Select()
            time.sleep(0.8)
            pane.open(None)
            for _ in range(12):
                if "datenreihe" in pane.title().lower():
                    break
                time.sleep(0.25)
            tab = next((e for n, e in pane.tabs() if n == "Füllung und Linie"), None)
            if tab is not None:
                select_item(tab)
                time.sleep(0.6)
            b = pane.body()
            f = b.FindAll(U.TreeScope_Descendants, UIA.CreateTrueCondition())
            pages = []
            for i in range(f.Length):
                d = d_of(f.GetElement(i))
                if d and d["name"] in ("Linie", "Markierung") and d["type"] in ("TabItem", "Button", "RadioButton", "ListItem"):
                    pages.append((d["name"], d["type"]))
            hit = None
            f = b.FindAll(U.TreeScope_Descendants, UIA.CreateTrueCondition())
            for i in range(f.Length):
                e = f.GetElement(i)
                d = d_of(e)
                if d and d["name"] == "Markierung" and d["type"] in ("TabItem", "Button", "RadioButton", "ListItem"):
                    hit = e
                    break
            out = {"page switch": pages}
            if hit is None:
                out["error"] = "no Markierung page"
            else:
                select_item(hit)
                time.sleep(0.8)
                CUR[0] = "Markierung"
                secs = []
                for st, _ in pane.sections():
                    if st:
                        expand_section(pane, st)
                    secs.append(deep_section(pane, st))
                out["sections"] = secs
            data.setdefault("markers", {})[f"{kind}/Datenreihe: Füllung und Linie > Markierung"] = out
            print(kind, "markers", [s["title"] for s in out.get("sections", [])], flush=True)
            save(data, "markers")
            sh.Delete()
    finally:
        try:
            data.setdefault("markers", {})
            save(data, "markers")
        finally:
            x.close()


def _shape(p: dict):
    return [(s["label"], [(t["label"], [sec["title"] for sec in t["sections"]]) for t in s["tabs"]]) for s in p.get("sets", [])]


def meta(data: dict) -> dict:
    m = data.get("meta") or {}
    m.update({
        "recorded": "2026-11-08/09, Excel 365 (German), own instance (DispatchEx) off-screen, no real input",
        "how": "chart element selected through COM; pane opened with CommandBars.ExecuteMso('ChartFormatSelection'); read "
               "through UI Automation; sections expanded and radios selected through automation patterns; dropdowns "
               "opened and closed through ExpandCollapse; labels only (no tooltips)",
        "data": "A1:C5 Monat/A/B (Jan 4 6, Feb -2 5, Mär 7 3, Apr 5 8); scatter E1:G5 X/Y1/Y2 with X 1, 2.5, 4, 6",
        "charts": {k: v for k, v in CHARTS.items()},
        "chart setup": "title, legend, data labels on series 1; except pie: linear trendline and fixed-value error bars "
                       "(1) on series 1, title on Axes(2), major gridlines on Axes(2); combo: series 2 as line with "
                       "markers on the secondary axis",
        "fields": "t control type; label exact German; value current value (checkbox/toggle true/false, radio selected); "
                  "range [min, max] when exposed; disabled when greyed; menu: opens a dropdown; entries: its entries; "
                  "picker 'Farbe': colour picker (entries not recorded); group: sub-heading inside the section; "
                  "choices: per radio choice (other than the selected one) the controls it adds or removes; "
                  "same_as: identical to the section (or pane) recorded under that key",
    })
    return m


# ----- measurements ------------------------------------------------------------------------------------------------


def pdf_of_sheet(x: ExcelTarget, ws, path: Path):
    from feellab import pdftext
    ps = ws.PageSetup
    ps.Zoom = 100
    ps.Orientation = 2
    ws.ExportAsFixedFormat(0, str(path))
    pdf = pdftext.Pdf(path.read_bytes())
    num = next(iter(pdf.pages()))
    w, h = pdf.page_box(num)
    content = pdf.contents(num)
    return pdftext.paths(content, h), pdftext.text_runs(content, h)


def segments(paths: list[dict]) -> list[tuple[float, float, float, float, float, str]]:
    """Straight stroked segments: (x1, y1, x2, y2, width, colour)."""
    out = []
    for p in paths:
        if p["paint"] not in ("stroke", "both"):
            continue
        last = None
        for s in p["d"]:
            if s[0] == "M":
                last = s[1:3]
            elif s[0] == "L" and last:
                out.append((last[0], last[1], s[1], s[2], p["width"], p["stroke"]))
                last = s[1:3]
    return out


def measure_ticks(x: ExcelTarget, ws, tmp: Path) -> dict:
    """Tick mark lengths: a column chart with black axis lines and no labels, gridlines or fills."""
    out = {}
    for mark_name, mark in (("Außen", 3), ("Innen", 2), ("Kreuz", 4)):
        for which in ("major", "minor"):
            for s in list(ws.ChartObjects()):
                s.Delete()
            sh = ws.Shapes.AddChart2(-1, 51, 40, 40, 400, 300)
            ch = sh.Chart
            ch.SetSourceData(ws.Range("A1:C5"))
            ch.HasTitle = False
            ch.HasLegend = False
            # The chart area keeps its frame (grey): its printed rectangle gives the PDF's scale.
            ch.PlotArea.Format.Fill.Visible = False
            for i in (1, 2):
                ch.SeriesCollection(i).Format.Fill.Visible = False
                ch.SeriesCollection(i).Format.Line.Visible = False
            for ax_i, ax_name in ((2, "Wertachse"), (1, "Kategorieachse")):
                ax = ch.Axes(ax_i)
                ax.HasMajorGridlines = False
                ax.HasMinorGridlines = False
                ax.TickLabelPosition = -4142
                ax.Format.Line.Visible = True
                ax.Format.Line.ForeColor.RGB = 0
                ax.Format.Line.Weight = 0.75
                if which == "major":
                    ax.MajorTickMark = mark
                    ax.MinorTickMark = -4142
                else:
                    ax.MajorTickMark = -4142
                    ax.MinorTickMark = mark
            pdf = tmp / f"raabe-chartpane-ticks-{mark}-{which}.pdf"
            paths, _ = pdf_of_sheet(x, ws, pdf)
            scale = chart_scale(paths, sh.Width)
            row = {"pdf_scale": _r(scale)}
            # Only our 0.75 pt black axis lines (the printout adds a hairline frame round the chart).
            segs = [s for s in segments(paths) if s[5] == "000000" and s[4] > 0.5]
            vert = [s for s in segs if abs(s[0] - s[2]) < 0.01]
            hori = [s for s in segs if abs(s[1] - s[3]) < 0.01]
            # The axis lines are the longest; ticks are the short segments crossing them.
            if vert and hori:
                vaxis = max(vert, key=lambda s: abs(s[3] - s[1]))
                haxis = max(hori, key=lambda s: abs(s[2] - s[0]))
                ax_x, ax_y = vaxis[0], haxis[1]
                vt = [s for s in hori if s is not haxis and min(s[0], s[2]) <= ax_x + 0.01 and max(s[0], s[2]) >= ax_x - 0.01]
                ht = [s for s in vert if s is not vaxis and min(s[1], s[3]) <= ax_y + 0.01 and max(s[1], s[3]) >= ax_y - 0.01]
                row["Wertachse"] = {"count": len(vt), "left_of_axis": _r(max((ax_x - min(s[0], s[2]) for s in vt), default=None)),
                                    "right_of_axis": _r(max((max(s[0], s[2]) - ax_x for s in vt), default=None)),
                                    "width": vt[0][4] if vt else None}
                row["Kategorieachse"] = {"count": len(ht), "below_axis": _r(max((max(s[1], s[3]) - ax_y for s in ht), default=None)),
                                         "above_axis": _r(max((ax_y - min(s[1], s[3]) for s in ht), default=None)),
                                         "width": ht[0][4] if ht else None}
                row["axis_width"] = vaxis[4]
                # In chart points (the printout scales the chart slightly).
                for k in ("Wertachse", "Kategorieachse"):
                    row[k + " (chart pt)"] = {kk: _r(v / scale) for kk, v in row[k].items()
                                              if kk.endswith("axis") and v is not None and scale}
            else:
                row["error"] = f"segments: {len(segs)}"
            out[f"{which} {mark_name}"] = row
            pdf.unlink(missing_ok=True)
    return out


def chart_scale(paths: list[dict], width_pt: float) -> float | None:
    """PDF points per chart point: the widest closed rectangle (the chart's frame) against the chart's width."""
    best = 0.0
    for p in paths:
        xs = [s[1] for s in p["d"] if s[0] in ("M", "L")]
        if xs and any(s[0] == "Z" for s in p["d"]):
            best = max(best, max(xs) - min(xs))
    return best / width_pt if best and width_pt else None


def _r(v):
    return None if v is None else round(v, 3)


def measure_offset(x: ExcelTarget, ws, tmp: Path) -> dict:
    """Distance of the category labels from the category axis line for TickLabels.Offset 0 / 100 / 500 / 1000,
    and of the value labels from the value axis."""
    out = {}
    for s in list(ws.ChartObjects()):
        s.Delete()
    sh = ws.Shapes.AddChart2(-1, 51, 40, 40, 400, 300)
    ch = sh.Chart
    ch.SetSourceData(ws.Range("A1:C5"))
    out["default Axes(1).TickLabels.Offset"] = ch.Axes(1).TickLabels.Offset
    out["default Axes(2).TickLabels.Offset"] = ch.Axes(2).TickLabels.Offset
    ch.HasTitle = False
    ch.HasLegend = False
    for i in (1, 2):
        ch.SeriesCollection(i).Format.Fill.Visible = False
    for ax_i in (1, 2):
        ax = ch.Axes(ax_i)
        ax.HasMajorGridlines = False
        ax.Format.Line.Visible = True
        ax.Format.Line.ForeColor.RGB = 0
        ax.MajorTickMark = -4142
    rows = {}
    for off in (0, 100, 500, 1000):
        ch.Axes(1).TickLabels.Offset = off
        ch.Axes(2).TickLabels.Offset = off
        pdf = tmp / f"raabe-chartpane-offset-{off}.pdf"
        paths, runs = pdf_of_sheet(x, ws, pdf)
        scale = chart_scale(paths, sh.Width) or 1.0
        segs = [s for s in segments(paths) if s[5] == "000000" and s[4] > 0.5]
        vert = [s for s in segs if abs(s[0] - s[2]) < 0.01]
        hori = [s for s in segs if abs(s[1] - s[3]) < 0.01]
        r = {"Offset": ch.Axes(1).TickLabels.Offset, "pdf_scale": _r(scale)}
        if vert and hori:
            ax_x = max(vert, key=lambda s: abs(s[3] - s[1]))[0]
            ax_y = max(hori, key=lambda s: abs(s[2] - s[0]))[1]
            below = [t for t in runs if t.y > ax_y and t.x > ax_x - 2]
            if below:
                b = min(below, key=lambda t: t.y)
                r["Kategorie: baseline below axis (chart pt)"] = _r((b.y - ax_y) / scale)
                r["Kategorie: font size (chart pt)"] = _r(b.size / scale)
        # The value labels' right end is not in the PDF runs: the plot's inside left moves by the change in the gap.
        r["PlotArea.InsideLeft/InsideTop/InsideWidth/InsideHeight"] = [_r(ch.PlotArea.InsideLeft), _r(ch.PlotArea.InsideTop),
                                                                       _r(ch.PlotArea.InsideWidth), _r(ch.PlotArea.InsideHeight)]
        rows[str(off)] = r
        pdf.unlink(missing_ok=True)
    out["by offset"] = rows
    return out


def measure_display_units(x: ExcelTarget, ws) -> dict:
    for s in list(ws.ChartObjects()):
        s.Delete()
    sh = ws.Shapes.AddChart2(-1, 51, 40, 40, 400, 300)
    ch = sh.Chart
    ch.SetSourceData(ws.Range("A1:C5"))
    ax = ch.Axes(2)
    units = {"xlHundreds": -2, "xlThousands": -3, "xlTenThousands": -4, "xlHundredThousands": -5, "xlMillions": -6,
             "xlTenMillions": -7, "xlHundredMillions": -8, "xlThousandMillions": -9, "xlMillionMillions": -10}
    out = {"default DisplayUnit": ax.DisplayUnit}
    for name, v in units.items():
        try:
            ax.DisplayUnit = v
            ax.HasDisplayUnitLabel = True
            lab = ax.DisplayUnitLabel
            out[name] = {"value": v, "text": lab.Text, "orientation": lab.Orientation,
                         "font size": lab.Format.TextFrame2.TextRange.Font.Size}
        except Exception as e:  # noqa: BLE001
            out[name] = {"value": v, "error": str(e)[:80]}
    ax.DisplayUnit = -4142
    try:
        ax.DisplayUnit = -4114  # xlCustom
        ax.DisplayUnitCustom = 25
        ax.HasDisplayUnitLabel = True
        out["xlCustom 25"] = {"text": ax.DisplayUnitLabel.Text}
    except Exception as e:  # noqa: BLE001
        out["xlCustom 25"] = {"error": str(e)[:80]}
    sh.Delete()
    return out


def measure_shadows(x: ExcelTarget, ws) -> dict:
    """Format.Shadow after setting each msoShadowType on a column series (Office's preset list)."""
    for s in list(ws.ChartObjects()):
        s.Delete()
    sh = ws.Shapes.AddChart2(-1, 51, 40, 40, 400, 300)
    ch = sh.Chart
    ch.SetSourceData(ws.Range("A1:C5"))
    sh_ = ch.SeriesCollection(1).Format.Shadow
    out = {"default": _shadow(sh_)}
    for t in range(1, 63):
        try:
            sh_.Type = t
            out[str(t)] = _shadow(sh_)
        except Exception as e:  # noqa: BLE001
            out[str(t)] = {"error": str(e)[:60]}
    sh.Delete()
    return out


def rgb(v: int) -> str:
    return f"#{v & 255:02X}{(v >> 8) & 255:02X}{(v >> 16) & 255:02X}" if isinstance(v, int) and v >= 0 else str(v)


def _glow(f) -> dict:
    g = f.Glow
    o = {"Radius": _r(g.Radius), "Transparency": _r(g.Transparency)}
    try:
        o["Color"] = rgb(g.Color.RGB)
        o["ObjectThemeColor"] = g.Color.ObjectThemeColor
    except Exception:  # noqa: BLE001
        pass
    return o


def _soft(f) -> dict:
    s = f.SoftEdge
    return {"Type": s.Type, "Radius": _r(s.Radius)}


def measure_presets(x: ExcelTarget, ws) -> dict:
    """The pane's preset galleries (Schatten, Leuchten, Weiche Kanten) on a column series: each entry is picked
    through UI Automation, then Format.Shadow / Glow / SoftEdge are read through COM."""
    for s in list(ws.ChartObjects()):
        s.Delete()
    sh = ws.Shapes.AddChart2(-1, 51, 40, 40, 400, 300)
    ch = sh.Chart
    ch.SetSourceData(ws.Range("A1:C5"))
    sh.Select()
    ch.SeriesCollection(1).Select()
    pane = Pane(x)
    if not pane.open(None):
        return {"error": "pane did not open"}
    time.sleep(0.5)
    tab = next((e for n, e in pane.tabs() if n == "Effekte"), None)
    if tab is None:
        return {"error": "no Effekte tab"}
    select_item(tab)
    time.sleep(0.6)
    fmt = lambda: ch.SeriesCollection(1).Format  # noqa: E731
    readers = {"Schatten": lambda: _shadow(fmt().Shadow), "Leuchten": lambda: _glow(fmt()), "Weiche Kanten": lambda: _soft(fmt())}
    out: dict = {}
    for sec, reader in readers.items():
        expand_section(pane, sec)
        item = lambda: next((i for i in raw_section(pane, sec) if i["t"] == "MenuItem" and i["label"].startswith("Voreinstellungen")), None)  # noqa: E731
        it = item()
        if it is None:
            out[sec] = {"error": "no Voreinstellungen"}
            continue
        names = [n for n in (dropdown_entries(pane, it["_el"], it["_d"]) or []) if not n.startswith("[")]
        res = {}
        for nm in names:
            it = item()
            p = expander(it["_el"]) if it else None
            if p is None:
                res[nm] = {"error": "gallery gone"}
                continue
            p.Expand()
            hit = None
            for _ in range(10):
                time.sleep(0.3)
                cond = UIA.CreatePropertyCondition(U.UIA_ClassNamePropertyId, "NetUIToolWindow")
                f = pane.root.FindAll(U.TreeScope_Descendants, cond)
                for k in range(f.Length):
                    g = f.GetElement(k).FindAll(U.TreeScope_Descendants, UIA.CreateTrueCondition())
                    for j in range(g.Length):
                        e = g.GetElement(j)
                        d = d_of(e)
                        if d and d["type"] == "ListItem" and d["name"] == nm:
                            hit = e
                            break
                    if hit is not None:
                        break
                if hit is not None:
                    break
            if hit is None:
                res[nm] = {"error": "entry not found"}
                try:
                    p.Collapse()
                except Exception:  # noqa: BLE001
                    pass
                continue
            try:
                uia.invoke(hit)
            except Exception:  # noqa: BLE001
                try:
                    uia.select(hit)
                except Exception as e:  # noqa: BLE001
                    res[nm] = {"error": "pick: " + str(e)[:60]}
            time.sleep(0.6)
            try:
                if p.CurrentExpandCollapseState == 1:
                    p.Collapse()
            except Exception:  # noqa: BLE001
                pass
            try:
                res[nm] = reader()
            except Exception as e:  # noqa: BLE001
                res[nm] = {"error": str(e)[:80]}
            print(sec, nm, res[nm], flush=True)
        out[sec] = res
    sh.Delete()
    return out


def _shadow(s) -> dict:
    o = {}
    for p in ("Visible", "Type", "Style", "Blur", "OffsetX", "OffsetY", "Transparency", "Size", "Obscured", "RotateWithShape"):
        try:
            v = getattr(s, p)
            o[p] = round(v, 3) if isinstance(v, float) else v
        except Exception:  # noqa: BLE001
            pass
    try:
        o["ForeColor"] = rgb(s.ForeColor.RGB)
    except Exception:  # noqa: BLE001
        pass
    return o


def run_measure() -> None:
    data = load()
    x = start_excel()
    tmp = Path(tempfile.gettempdir())
    try:
        ws = x.wb.Worksheets(1)
        fill_data(ws)
        m = data.setdefault("measures", {})
        for name, f in (("tick marks (pt, PDF export at 100 %)", lambda: measure_ticks(x, ws, tmp)),
                        ("tick label offset (pt, PDF export at 100 %)", lambda: measure_offset(x, ws, tmp)),
                        ("display units", lambda: measure_display_units(x, ws)),
                        ("shadow types (Format.Shadow.Type = n)", lambda: measure_shadows(x, ws)),
                        ("effect presets picked in the pane (series Format via COM)", lambda: measure_presets(x, ws))):
            try:
                m[name] = f()
            except Exception as e:  # noqa: BLE001
                m[name] = {"error": str(e)[:200]}
            print(name, json.dumps(m[name], ensure_ascii=False)[:400], flush=True)
            save(data, "measures")
    finally:
        save(data, "measures")
        x.close()


def compact() -> None:
    """Sections whose controls (type, label, group, enabled, entries) and radio choices equal a section recorded in an
    earlier pane become references; the values that differ (default colours ...) are kept next to the reference."""
    data = load()
    first: dict[str, tuple[str, dict]] = {}
    for key, p in data["panes"].items():
        for s in p.get("sets", []):
            for t in s["tabs"]:
                for i, sec in enumerate(t["sections"]):
                    if "controls" not in sec:
                        continue
                    shape = json.dumps([sec["title"], [(c["t"], c["label"], c.get("group"), c.get("disabled"), c.get("entries"))
                                                       for c in sec["controls"]], sorted(sec.get("choices", {}))], ensure_ascii=False)
                    if shape not in first:
                        first[shape] = (key, sec)
                        continue
                    ref, base = first[shape]
                    if ref == key:
                        continue
                    diff = {c["label"]: c.get("value") for c, b in zip(sec["controls"], base["controls"]) if c.get("value") != b.get("value")}
                    new = {"title": sec["title"], "same_as": ref}
                    if diff:
                        new["values"] = diff
                    t["sections"][i] = new
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1) + "\n", encoding="utf-8", newline="\n")


def main() -> None:
    args = sys.argv[1:] or ["panes", "measure"]
    if args == ["compact"]:
        compact()
        return
    if args == ["markers"]:
        run_markers()
        return
    if "measure" in args:
        run_measure()
        args = [a for a in args if a != "measure"]
    if args and args[0] == "panes":
        kinds = args[1:] or list(CHARTS)
        run_panes(kinds)


if __name__ == "__main__":
    os.chdir(HERE)
    main()
