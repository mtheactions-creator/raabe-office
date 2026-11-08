"""Records the Sparkline tab of German Excel 365 (D-085) through UI Automation, without real input:
the keytips of its controls, which controls are enabled for each sparkline type and selection, and
the 36 styles of the Sparklineformatvorlage gallery with the colours
each one gives (the gallery item is invoked through UI Automation, the colours are read from the
file Excel saves).

Our own Excel (`ExcelTarget`), made visible off screen. Selecting the tab or invoking a gallery item
makes Excel the foreground window for a moment (recorded: under 0.1 s); the probe gives the focus
back to the window that had it after every such step. Nothing is typed or clicked.

    .venv/Scripts/python.exe probe_excel_sparkline_ui.py

Writes results/2026-11-08-excel-sparkline-ui.json (labels, keytips, states), the styles into
packages/ooxml/test/fixtures/excel-sparklines.json ("styles") and
apps/raabe-calc/src/sparkline-styles.gen.ts."""

from __future__ import annotations

import json
import re
import sys
import tempfile
import time
import zipfile
from pathlib import Path

from feellab import uia, winapi
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
RESULT = HERE / "results" / "2026-11-08-excel-sparkline-ui.json"
FIXTURE = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "excel-sparklines.json"
GEN = ROOT / "apps" / "raabe-calc" / "src" / "sparkline-styles.gen.ts"
TMP = Path(tempfile.gettempdir())
COLOURS = ("colorSeries", "colorNegative", "colorAxis", "colorMarkers", "colorFirst", "colorLast", "colorHigh", "colorLow")

fg0 = winapi.foreground()


def give_back() -> None:
    if fg0 and winapi.foreground() != fg0:
        try:
            winapi.activate(fg0, 0)
        except Exception:  # noqa: BLE001
            pass


def descendants(root) -> list[tuple]:
    out = []
    f = root.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    for i in range(f.Length):
        e = f.GetElement(i)
        try:
            out.append((e, uia.describe(e)))
        except Exception:  # noqa: BLE001
            continue
    return out


def ribbon_controls(root) -> list[dict]:
    """The lower ribbon's named controls (the selected tab)."""
    return [d for _, d in descendants(root) if 74 < d["rect"][1] < 268 and d["name"] and d["type"] != "ListItem"]


def states(root, names) -> dict:
    found = {d["name"]: d["enabled"] for d in ribbon_controls(root)}
    return {n: found.get(n) for n in names}


STATE_NAMES = ["Daten bearbeiten", "Linie", "Säule", "Gewinn/Verlust", "Höchstpunkt", "Tiefpunkt", "Negative Punkte", "Erster Punkt",
               "Letzter Punkt", "Datenpunkte", "Sparklinefarbe", "Datenpunktfarbe", "Achse", "Gruppieren", "Gruppierung aufheben", "Löschen"]


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    result: dict = {}
    x = ExcelTarget(visible=False)
    path = TMP / "raabe-sparkline-styles.xlsx"
    try:
        wb = x.wb
        ws = wb.Worksheets(1)
        for r in range(1, 41):
            for c, v in enumerate((3, 5, -2, 8, 4), 1):
                ws.Cells(r, c).Value2 = v
        # F1:F3 one group of three, F5 one of one (column), F6 one of one (Gewinn/Verlust), H1:H36 one each for the styles.
        ws.Range("F1:F3").SparklineGroups.Add(1, "A1:E3")
        ws.Range("F5").SparklineGroups.Add(2, "A5:E5")
        ws.Range("F6").SparklineGroups.Add(3, "A6:E6")
        ws.Range("F7").SparklineGroups.Add(1, "A7:E7")
        for i in range(36):
            ws.Range(f"H{i + 1}").SparklineGroups.Add(1, f"A{i + 1}:E{i + 1}")
        ws.Range("F1").Select()
        winapi.place_quietly(x.hwnd, -4000, 0, 2400, 1300)
        x.app.Visible = True
        time.sleep(1.0)
        give_back()
        root = uia.element_for(x.hwnd)
        tab = uia.find(root, lambda d: d["type"] == "TabItem" and d["name"] == "Sparkline", 25)
        result["tab"] = {"name": "Sparkline", "key": uia.describe(tab)["key"]}
        uia.select(tab)
        give_back()
        time.sleep(1.0)
        result["controls"] = [{k: d[k] for k in ("type", "name", "key", "enabled")} for d in ribbon_controls(root)]

        # Enabled states by selection (selection set through COM).
        result["states"] = {}
        for label, sel in (("line group of 3, F1", "F1"), ("column group of 1", "F5"), ("winloss group of 1", "F6"), ("line group of 1", "F7"),
                           ("two groups F3:F5", "F3:F5"), ("whole group F1:F3", "F1:F3"), ("no sparkline A1", "A1")):
            ws.Range(sel).Select()
            time.sleep(0.6)
            result["states"][label] = states(root, STATE_NAMES)
            print(label, result["states"][label], flush=True)
        ws.Range("F1").Select()
        time.sleep(0.6)
        # Leaving the sparklines (A1) took the tab away: select it again (the gallery needs it).
        tab = uia.find(root, lambda d: d["type"] == "TabItem" and d["name"] == "Sparkline", 25)
        uia.select(tab)
        give_back()
        time.sleep(1.0)

        # The style gallery: expand it, list the styles, then apply style i to the sparkline in H(i+1).
        names: list[str] = []
        styles = []
        for i in range(36):
            ws.Range(f"H{i + 1}").Select()
            time.sleep(0.4)
            more = next(e for e, d in descendants(root) if d["name"] == "Sparklineformatvorlage" and d["type"] == "Button")
            uia.invoke(more)
            give_back()
            time.sleep(0.8)
            items = [(e, d) for e, d in _popups(x.pid) + descendants(root) if d["type"] == "ListItem" and "Sparklineformat" in d["name"]]
            seen = []
            for e, d in items:
                if d["name"] not in [s[1]["name"] for s in seen]:
                    seen.append((e, d))
            if not names:
                names = [d["name"] for _, d in seen]
                print(len(names), "styles", flush=True)
            target = next((e for e, d in seen if d["name"] == names[i]), None)
            if target is None:
                styles.append({"name": names[i], "error": "not found"})
                continue
            try:
                uia.invoke(target)
            except Exception:  # noqa: BLE001
                uia.select(target)
            give_back()
            time.sleep(0.5)
            styles.append({"name": names[i], "cell": f"H{i + 1}"})
        result["styles"] = [s["name"] for s in styles]
        x.app.Visible = False
        wb.SaveAs(str(path), 51)
    finally:
        give_back()
        x.close()
        give_back()

    with zipfile.ZipFile(path) as z:
        xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    groups = {}
    for g in re.findall(r"<x14:sparklineGroup\b[\s\S]*?</x14:sparklineGroup>", xml):
        cell = re.search(r"<xm:sqref>([^<]+)</xm:sqref>", g).group(1)
        groups[cell] = {c: re.search(rf"<x14:{c}\b([^/]*)/>", g).group(1).strip() for c in COLOURS}
    fixture = json.loads(FIXTURE.read_text(encoding="utf-8"))
    fixture["styles"] = [{**s, "colors": groups.get(s.get("cell", ""), {})} for s in styles]
    FIXTURE.write_text(json.dumps(fixture, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    RESULT.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    write_gen(fixture["styles"])
    print("written", RESULT.name, len(fixture["styles"]), "styles")


def _popups(pid: int) -> list[tuple]:
    """Every element of our Excel's top-level windows other than the main one (menus, gallery popups)."""
    out = []
    root = uia._uia.GetRootElement()
    cond = uia._uia.CreatePropertyCondition(uia.U.UIA_ProcessIdPropertyId, pid)
    f = root.FindAll(uia.U.TreeScope_Children, cond)
    for i in range(f.Length):
        out.extend(descendants(f.GetElement(i)))
    return out


def write_gen(styles: list[dict]) -> None:
    def colour(attrs: str) -> dict:
        a = dict(re.findall(r'(\w+)="([^"]*)"', attrs))
        if "rgb" in a:
            return {"rgb": "#" + a["rgb"][-6:].upper()}
        ref = {"t": int(a["theme"])}
        if "tint" in a:
            ref["tint"] = float(a["tint"])
        return {"theme": ref}
    keys = {"colorSeries": "series", "colorNegative": "negative", "colorAxis": "axis", "colorMarkers": "markers",
            "colorFirst": "first", "colorLast": "last", "colorHigh": "high", "colorLow": "low"}
    rows = []
    for s in styles:
        cols = {keys[k]: colour(v) for k, v in s.get("colors", {}).items()}
        rows.append(f"  {{ name: {json.dumps(s['name'], ensure_ascii=False)}, colors: {json.dumps(cols)} }},")
    GEN.write_text(
        "// Generated by tools/feel-lab/probe_excel_sparkline_ui.py from the files German Excel 365 saved after each\n"
        "// style of the Sparklineformatvorlage gallery was applied (D-085). Do not edit.\n\n"
        "import type { SparkColorKey, ThemeRef } from \"@raabe/sheet-engine\";\n\n"
        "/** A style's colour: RGB, or a design colour (resolved in the workbook's theme when applied). */\n"
        "export type SparkStyleColor = { rgb: string } | { theme: ThemeRef };\n\n"
        "export interface SparklineStyle { name: string; colors: Record<SparkColorKey, SparkStyleColor> }\n\n"
        "export const SPARKLINE_STYLES: readonly SparklineStyle[] = [\n" + "\n".join(rows) + "\n];\n",
        encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
