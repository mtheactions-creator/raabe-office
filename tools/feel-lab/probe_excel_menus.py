"""Probe: the entries of Excel's ribbon menus (Rahmen, Einfügen, Löschen, Format, AutoSumme,
Ausfüllen, Löschen, Sortieren und Filtern, Suchen und Auswählen, Bedingte Formatierung,
Verbinden und zentrieren, Ausrichtung, Fenster fixieren, Druckbereich, Ausrichtung of the page,
Berechnungsoptionen …) for the audit of RaabeCalc's menus (package F). Own visible Excel; real
clicks on the ribbon. Labels only (Microsoft's descriptions are not kept).
Result: results/<date>-excel-menus.json.

    .venv/Scripts/python.exe probe_excel_menus.py
"""
from __future__ import annotations

import datetime
import json
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.calc_catalogue import Case, ExcelSide
from probe_word_context import items_of, popups

HERE = Path(__file__).parent
MENUS = [
    ("Start", "Einfügen", "arrow"), ("Start", "Rahmen", "arrow"), ("Start", "Unterstreichen", "arrow"),
    ("Start", "Ausrichtung", "item"), ("Start", "Verbinden und zentrieren", "arrow"), ("Start", "Buchhaltungszahlenformat", "arrow"),
    ("Start", "Bedingte Formatierung", "item"), ("Start", "Zellen einfügen", "arrow"), ("Start", "Zellen löschen...", "arrow"),
    ("Start", "Format", "item"), ("Start", "AutoSumme", "arrow"), ("Start", "Ausfüllen", "item"), ("Start", "Löschen", "item"),
    ("Start", "Sortieren und Filtern", "item"), ("Start", "Suchen und Auswählen", "item"),
    ("Seitenlayout", "Seitenränder", "item"), ("Seitenlayout", "Ausrichtung", "item"), ("Seitenlayout", "Druckbereich", "item"),
    ("Formeln", "Logisch", "item"), ("Formeln", "Text", "item"), ("Formeln", "Berechnungsoptionen", "item"),
    ("Ansicht", "Fenster fixieren", "item"),
]


def open_menu(side: ExcelSide, tab: str, name: str, how: str) -> list[dict]:
    el = uia.find(side.root, lambda d, t=tab: d["type"] == "TabItem" and d["name"] == t)
    uia.select(el)
    time.sleep(0.6)
    lower = uia.find(side.root, lambda d: d["name"] == "Unteres Menüband", max_depth=12)
    target = None
    stack = [lower]
    while stack and target is None:
        e = stack.pop()
        for c in uia.children(e):
            d = uia.describe(c)
            if d["name"] == name and d["type"] in ("SplitButton", "MenuItem", "Button"):
                target = (c, d)
                break
            stack.append(c)
    if target is None:
        return [{"error": "not found"}]
    c, d = target
    x, y, w, h = d["rect"]
    if how == "arrow" and d["type"] == "SplitButton":
        # The arrow part: the "Weitere Optionen" child.
        arrow = [uia.describe(k) for k in uia.children(c) if uia.describe(k)["name"] == "Weitere Optionen"]
        if arrow:
            x, y, w, h = arrow[0]["rect"]
    winapi.click(x + w / 2, y + h / 2)
    time.sleep(0.9)
    # Office shows the open menu's entries below the expanded control in the window's tree.
    items: list[dict] = []
    try:
        f = c.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
        for i in range(f.Length):
            dd = uia.describe(f.GetElement(i))
            if dd["type"] in ("MenuItem", "Button", "SplitButton", "CheckBox", "RadioButton", "ListItem", "Group", "Text") and dd["name"] and len(dd["name"]) < 80:
                items.append({"name": dd["name"], "type": dd["type"], "enabled": dd["enabled"]})
    except Exception as e:  # noqa: BLE001
        items.append({"error": str(e)})
    winapi.press("esc")
    time.sleep(0.3)
    winapi.press("esc")
    time.sleep(0.3)
    return items


def main() -> None:
    side = ExcelSide()
    out: dict = {}
    try:
        side.reset(Case("menus", {"A1": "1"}, [], "x"))
        winapi.activate(side.x.hwnd)
        time.sleep(0.5)
        for tab, name, how in MENUS:
            out[f"{tab}/{name}"] = open_menu(side, tab, name, how)
            print(tab, name, [i.get("name") for i in out[f"{tab}/{name}"]], flush=True)
    finally:
        side.close()
    (HERE / "results" / f"{datetime.date.today().isoformat()}-excel-menus.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
