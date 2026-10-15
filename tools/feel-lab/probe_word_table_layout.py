"""Probe: Word's Tabellenlayout tab beyond what Step 3 built (Step 4c F). In our own visible
Word instance, with the caret in a 3 x 3 table: the dialogs Zellen einfügen (launcher of Zeilen
und Spalten), Tabellenoptionen (Zellenbegrenzungen), Tabelle in Text umwandeln (Daten > In Text
konvertieren) and the AutoAnpassen menu (entries as UI Automation sees them); then, through
COM, what Tabelle teilen, In Text konvertieren, the four Zellen einfügen choices and the three
AutoAnpassen choices do to the document (text, rows, the table's XML properties). Takes over
mouse and screen. Result: results/<date>-word-table-layout.json.

    .venv/Scripts/python.exe probe_word_table_layout.py [com | tab]   # com: only the behaviour part, tab: the ribbon tab
"""
from __future__ import annotations

import datetime
import json
import re
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from probe_word_dialogs import Word, new_windows, snapshot

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-word-table-layout.json"


def table_state(d) -> dict:
    out = {"text": d.Content.Text.replace("\r", "|").replace("\x07", "¤"), "tables": d.Tables.Count}
    if d.Tables.Count:
        t = d.Tables(1)
        out["rows"] = t.Rows.Count
        out["cells"] = [f"{c.RowIndex},{c.ColumnIndex}:{c.Range.Text.rstrip(chr(13) + chr(7))}" for c in t.Range.Cells]
    return out


def table_props(d) -> dict:
    """tblW, tblLayout, the grid and the cell widths of the first table, from its WordprocessingML."""
    xml = d.Tables(1).Range.WordOpenXML
    tbl = xml[xml.index("<w:tbl>"):]
    pr = re.search(r"<w:tblPr>(.*?)</w:tblPr>", tbl).group(1)
    return {
        "tblW": re.findall(r"<w:tblW [^>]*/>", pr),
        "tblLayout": re.findall(r"<w:tblLayout [^>]*/>", pr),
        "grid": re.findall(r'<w:gridCol w:w="(\d+)"', tbl),
        "tcW": re.findall(r"<w:tcW [^>]*/>", tbl)[:6],
    }


def record_dialog(w: Word, el, menu_item: str | None = None) -> dict:
    x, y, cw, ch = uia.describe(el)["rect"]
    before = set(winapi.find_windows(lambda h: True))
    winapi.click(x + cw / 2, y + ch / 2)
    time.sleep(1.2)
    if menu_item:
        hit = uia.find(w.root, lambda d: d["name"] == menu_item and d["type"] in ("MenuItem", "Button"), max_depth=20)
        if hit is None:
            for h in new_windows(before):
                hit = uia.find(uia.element_for(h), lambda d: d["name"] == menu_item, max_depth=12) or hit
        if hit is None:
            winapi.press("esc")
            return {"error": f"no menu item {menu_item}"}
        r = uia.describe(hit)["rect"]
        before = set(winapi.find_windows(lambda h: True))
        winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
        time.sleep(1.2)
    wins = new_windows(before)
    dialog = next((h for h in wins if winapi.window_class(h) in ("bosa_sdm_msword", "NUIDialog", "#32770")), None)
    if dialog is None:
        # A menu: its entries show below the expanded control.
        items = [i for i in snapshot(el, (0, 0)) if i["type"] in ("MenuItem", "Button")]
        winapi.press("esc")
        time.sleep(0.4)
        return {"menu": [i["name"] for i in items]}
    l, t, r, b = winapi.window_rect(dialog)
    el = uia.element_for(dialog)
    items = []
    f = el.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    for i in range(f.Length):
        e = f.GetElement(i)
        dd = uia.describe(e)
        if not dd["name"] and dd["type"] not in ("Edit", "ComboBox"):
            continue
        state = ""
        try:
            if dd["type"] == "RadioButton":
                pat = e.GetCurrentPattern(uia.U.UIA_SelectionItemPatternId).QueryInterface(uia.U.IUIAutomationSelectionItemPattern)
                state = "*" if pat.CurrentIsSelected else ""
            elif dd["type"] == "CheckBox":
                pat = e.GetCurrentPattern(uia.U.UIA_TogglePatternId).QueryInterface(uia.U.IUIAutomationTogglePattern)
                state = "[x]" if pat.CurrentToggleState == 1 else "[ ]"
            elif dd["type"] in ("Edit", "ComboBox"):
                pat = e.GetCurrentPattern(uia.U.UIA_ValuePatternId).QueryInterface(uia.U.IUIAutomationValuePattern)
                state = f"={pat.CurrentValue}"
        except Exception:  # noqa: BLE001
            pass
        x, y, ww, hh = dd["rect"]
        items.append({"type": dd["type"], "name": dd["name"] + state, "rect": [x - l, y - t, ww, hh]})
    res = {"title": winapi.window_title(dialog), "size": [r - l, b - t], "items": items}
    winapi.press("esc")
    time.sleep(0.6)
    return res


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    only_com = "com" in sys.argv[1:]
    out: dict = json.loads(OUT.read_text(encoding="utf-8")) if (only_com or "tab" in sys.argv[1:]) and OUT.exists() else {}
    w = Word()
    d, app = w.d, w.app
    try:
        def fresh(rows: int = 3, cols: int = 3):
            d.Content.Delete()
            d.Content.Text = "Vor\rNach"
            t = d.Tables.Add(d.Range(d.Paragraphs(1).Range.End, d.Paragraphs(1).Range.End), rows, cols, 1, 0)
            for r in range(1, rows + 1):
                for c in range(1, cols + 1):
                    t.Cell(r, c).Range.Text = f"Z{r}S{c}"
            return t

        # ----- dialogs and menus
        t = fresh()
        t.Cell(2, 2).Range.Select()
        app.Selection.Collapse(1)
        time.sleep(0.4)
        w.tab("Tabellenlayout")
        if "tab" in sys.argv[1:]:
            # The tab's controls with Word's keytips (UI Automation's AccessKey), in a window wide
            # enough for every group to show its controls.
            winapi.place(w.hwnd, 0, 0, 3700, 1300)
            time.sleep(1.0)
            w.tab("Tabellenlayout")
            l, t0, _, _ = winapi.window_rect(w.hwnd)
            (HERE / "captures").mkdir(exist_ok=True)
            winapi.save_region_png(str(HERE / "captures" / "word-Tabellenlayout.png"), l, t0 + 110, 2700, 170)
            lower = uia.find(w.root, lambda x: x["type"] == "Pane" and x["name"].startswith("Unter"), max_depth=12)
            out["tab"] = [i for i in snapshot(lower, (l, t0)) if i["key"] or i["type"] == "Group"] if lower is not None else []
            print("tab", len(out["tab"]), flush=True)
            return
        for name, find, item in () if only_com else (
            ("Zellen einfügen", lambda: w.group_launcher("Zeilen und Spalten"), None),
            ("Tabellenoptionen", lambda: w.control("Zellenbegrenzungen..."), None),
            ("AutoAnpassen", lambda: w.control("AutoAnpassen"), None),
            ("Tabelle in Text", lambda: w.control("In Text konvertieren...") or w.control("Daten"), "In Text konvertieren..."),
            ("Zellengröße launcher", lambda: w.group_launcher("Zellengröße"), None),
        ):
            el = find()
            if el is None:
                out[name] = {"error": "not found"}
                print(name, "not found", flush=True)
                continue
            direct = uia.describe(el)["name"] != "Daten"
            out[name] = record_dialog(w, el, None if direct else item)
            print(name, json.dumps(out[name], ensure_ascii=False)[:400], flush=True)
            w.tab("Tabellenlayout")

        # ----- behaviour through COM
        for row in (2, 1):
            t = fresh()
            t.Rows(row).Range.Select()
            app.Selection.SplitTable()
            out[f"Tabelle teilen vor Zeile {row}"] = table_state(d) | {"styles": [d.Paragraphs(i).Style.NameLocal for i in range(1, d.Paragraphs.Count + 1)][:14]}
        for sep, label in ((1, "Tabstopps"), (0, "Absatzmarken"), (2, "Semikola")):
            t = fresh(2, 2)
            t.Cell(1, 2).Range.Text = "a\rb"
            t.ConvertToText(Separator=sep)
            out[f"In Text: {label}"] = table_state(d)
        for k, label in enumerate(("Zellen nach rechts verschieben", "Zellen nach unten verschieben", "Ganze Zeile einfügen", "Ganze Spalte einfügen")):
            t = fresh()
            t.Cell(2, 2).Select()
            app.Selection.InsertCells(k)
            out[f"Zellen einfügen: {label}"] = table_state(d)
        for k, label in ((1, "Inhalt automatisch anpassen"), (2, "Fenster automatisch anpassen"), (0, "Feste Spaltenbreite")):
            t = fresh(2, 3)
            t.Cell(1, 1).Range.Text = "kurz"
            t.Cell(1, 2).Range.Text = "ein etwas längerer Text in der Zelle"
            t.AutoFitBehavior(k)
            out[f"AutoAnpassen: {label}"] = table_props(d)
        out["AutoAnpassen: new table"] = (fresh(2, 3), table_props(d))[1]
        for k, v in out.items():
            if k.startswith(("Tabelle teilen", "In Text", "Zellen einfügen:", "AutoAnpassen:")):
                print(k, json.dumps(v, ensure_ascii=False)[:400], flush=True)
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
