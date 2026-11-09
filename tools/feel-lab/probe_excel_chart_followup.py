"""Probe: what Excel 365 writes for the chart settings RaabeCalc did not build yet (D-155). In an Excel of our own
(DispatchEx, window off screen, no real input), each setting is made on a column chart, the workbook is saved as a
copy in TEMP and the chart part (and its style part) is read back as XML:

- com: an axis' Keine Füllung against its automatic fill (and what Excel reports after reopening), a gradient stop
  with Helligkeit (GradientStops.Insert2), Benutzerdefiniert error amounts from ranges, ChartStyle 205 (the style
  part's id), a horizontal bar chart with a Datentabelle.
- pane: the chart area's Format pane through UI Automation (as probe_excel_chartpane.py): every entry of
  Voreingestellte Farbverläufe, the Bild- oder Texturfüllung choice without a picture, every Textur (the tiling
  only, not Office's bitmaps) and the Perspektive shadow presets.

    .venv/Scripts/python.exe probe_excel_chart_followup.py [com] [pane]

Result: results/2026-11-09-excel-chart-followup.json (XML snippets of Excel's chart parts; no texts of Office's).
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
import time
import zipfile
from pathlib import Path

from feellab import uia
from feellab.excel import ExcelTarget
from probe_excel_chartpane import (UIA, U, Pane, d_of, dropdown_entries, expand_section, expander, fill_data, raw_section,
                                   select_item, select_radio, start_excel)

HERE = Path(__file__).parent
OUT = HERE / "results" / "2026-11-09-excel-chart-followup.json"
TMP = Path(tempfile.gettempdir()) / "raabe-chart-followup"
TMP.mkdir(exist_ok=True)
_n = [0]


def parts(wb) -> dict[str, str]:
    """Saves a copy of the workbook and returns its chart and style parts as text."""
    _n[0] += 1
    p = TMP / f"copy{_n[0]}.xlsx"
    wb.SaveCopyAs(str(p))
    out = {}
    with zipfile.ZipFile(p) as z:
        for n in z.namelist():
            if n.startswith("xl/charts/") and n.endswith(".xml"):
                out[n.split("/")[-1]] = z.read(n).decode("utf-8")
    return out


def chart_xml(wb) -> str:
    ps = parts(wb)
    return next((v for k, v in sorted(ps.items()) if k.startswith("chart")), "")


def el(xml: str, tag: str) -> str:
    m = re.search(rf"<{tag}\b[\s\S]*?</{tag}>", xml)
    return m.group(0) if m else ""


def sppr_of(block: str) -> str:
    m = re.search(r"<c:spPr>[\s\S]*?</c:spPr>", block)
    return m.group(0) if m else ""


def chart_area_sppr(xml: str) -> str:
    """The chart space's own c:spPr (after c:chart)."""
    i = xml.find("</c:chart>")
    m = re.search(r"<c:spPr>[\s\S]*?</c:spPr>", xml[i:]) if i >= 0 else None
    return m.group(0) if m else ""


def column_chart(ws, kind: int = 51):
    for s in list(ws.ChartObjects()):
        s.Delete()
    sh = ws.Shapes.AddChart2(-1, kind, 20, 120, 480, 300)
    ch = sh.Chart
    ch.SetSourceData(ws.Range("A1:C5"))
    return sh, ch


def run_com(data: dict) -> None:
    x = ExcelTarget(visible=False, place=False)
    try:
        wb = x.wb
        ws = wb.Worksheets(1)
        fill_data(ws)
        out: dict = {}
        sh, ch = column_chart(ws)
        base = chart_xml(wb)
        out["axis default"] = {"valAx": sppr_of(el(base, "c:valAx")), "catAx": sppr_of(el(base, "c:catAx"))}
        ch.Axes(2).Format.Fill.Visible = 0
        ch.Axes(1).Format.Fill.Visible = 0
        xml = chart_xml(wb)
        out["axis Keine Füllung"] = {"valAx": sppr_of(el(xml, "c:valAx")), "catAx": sppr_of(el(xml, "c:catAx"))}
        p = TMP / "axisnofill.xlsx"
        wb.SaveCopyAs(str(p))
        wb2 = x.app.Workbooks.Open(str(p))
        try:
            a = wb2.Worksheets(1).ChartObjects(1).Chart.Axes(2).Format.Fill
            out["axis Keine Füllung reopened"] = {"Visible": a.Visible, "Type": a.Type}
        finally:
            wb2.Close(False)
        # A gradient stop with Helligkeit.
        sh, ch = column_chart(ws)
        f = ch.ChartArea.Format.Fill
        f.TwoColorGradient(1, 1)
        f.ForeColor.RGB = 0x0000FF
        f.BackColor.RGB = 0xFF0000
        f.GradientStops.Insert2(0x00FF00, 0.5, 0.0, 2, 0.4)
        f.GradientStops.Insert2(0x00FF00, 0.7, 0.0, 3, -0.3)
        stops = []
        for i in range(1, f.GradientStops.Count + 1):
            g = f.GradientStops(i)
            stops.append({"pos": round(g.Position, 3), "rgb": g.Color.RGB, "brightness": round(g.Color.Brightness, 3)})
        out["gradient stop Helligkeit"] = {"com": stops, "xml": chart_area_sppr(chart_xml(wb))}
        # Benutzerdefiniert error amounts (Wert angeben with ranges).
        sh, ch = column_chart(ws)
        s1 = ch.SeriesCollection(1)
        s1.ErrorBar(1, 1, -4114, "=Tabelle1!$B$2:$B$5", "=Tabelle1!$C$2:$C$5")  # xlY, xlBoth, xlErrorBarTypeCustom
        xml = chart_xml(wb)
        out["error bars custom"] = el(xml, "c:errBars")
        s1.ErrorBar(1, 2, -4114, "={1;2;3;4}")  # xlPlusValues, an array
        out["error bars custom array"] = el(chart_xml(wb), "c:errBars")
        # ChartStyle: the style part's id.
        sh, ch = column_chart(ws)
        ch.ChartStyle = 205
        ps = parts(wb)
        style = next((v for k, v in ps.items() if k.startswith("style")), "")
        out["ChartStyle 205"] = {"cs:chartStyle": (re.search(r"<cs:chartStyle\b[^>]*>", style) or [""])[0],
                                 "c:style": (re.search(r"<c:style\b[^>]*/>", next((v for k, v in sorted(ps.items()) if k.startswith("chart")), "")) or [""])[0],
                                 "mc:style": (re.search(r"<mc:AlternateContent[\s\S]*?</mc:AlternateContent>", next((v for k, v in sorted(ps.items()) if k.startswith("chart")), "")) or [""])[0]}
        # A bar chart with a Datentabelle.
        sh, ch = column_chart(ws, 57)
        ch.HasDataTable = True
        xml = chart_xml(wb)
        out["bar Datentabelle"] = {"dTable": el(xml, "c:dTable"), "plotArea layout": (re.search(r"<c:plotArea><c:layout[\s\S]*?</c:layout>|<c:plotArea><c:layout/>", xml) or [""])[0]}
        pdf = TMP / "bar-table.pdf"
        try:
            ws.ChartObjects(1).Chart.ExportAsFixedFormat(0, str(pdf))
            out["bar Datentabelle pdf"] = str(pdf.name)
        except Exception as e:  # noqa: BLE001
            out["bar Datentabelle pdf"] = "error " + str(e)[:60]
        for s in list(ws.ChartObjects()):
            s.Delete()
        data["com"] = out
    finally:
        x.close()


def gallery_pick(pane: Pane, sec: str, label: str, name: str) -> str | None:
    """Opens the dropdown `label` in section `sec` and picks the entry `name`; None when done."""
    it = next((i for i in raw_section(pane, sec) if i["t"] == "MenuItem" and i["label"].startswith(label)), None)
    p = expander(it["_el"]) if it else None
    if p is None:
        return "no dropdown"
    p.Expand()
    hit = None
    for _ in range(12):
        time.sleep(0.3)
        cond = UIA.CreatePropertyCondition(U.UIA_ClassNamePropertyId, "NetUIToolWindow")
        f = pane.root.FindAll(U.TreeScope_Descendants, cond)
        for k in range(f.Length):
            g = f.GetElement(k).FindAll(U.TreeScope_Descendants, UIA.CreateTrueCondition())
            for j in range(g.Length):
                e = g.GetElement(j)
                d = d_of(e)
                if d and d["type"] in ("ListItem", "MenuItem", "Button") and d["name"] == name:
                    hit = e
                    break
            if hit is not None:
                break
        if hit is not None:
            break
    if hit is None:
        try:
            p.Collapse()
        except Exception:  # noqa: BLE001
            pass
        return "entry not found"
    try:
        uia.invoke(hit)
    except Exception:  # noqa: BLE001
        try:
            uia.select(hit)
        except Exception as e:  # noqa: BLE001
            return "pick: " + str(e)[:60]
    time.sleep(0.7)
    try:
        if p.CurrentExpandCollapseState == 1:
            p.Collapse()
    except Exception:  # noqa: BLE001
        pass
    return None


def tab(pane: Pane, name: str) -> bool:
    t = next((e for n, e in pane.tabs() if n == name), None)
    if t is None:
        return False
    select_item(t)
    time.sleep(0.6)
    return True


def run_pane(data: dict) -> None:
    rec = json.loads((HERE / "results" / "2026-11-09-excel-chart-pane.json").read_text(encoding="utf-8"))
    area = rec["panes"]["column/Diagrammbereich"]["sets"][0]["tabs"]
    fill = area[0]["sections"][0]
    presets = next(c["entries"] for c in fill["choices"]["Farbverlauf"]["adds"] if c["label"] == "Voreingestellte Farbverläufe")
    textures = next(c["entries"] for c in fill["choices"]["Bild- oder Texturfüllung"]["adds"] if c["label"] == "Textur")
    shadows = [e for e in area[1]["sections"][0]["controls"][0]["entries"] if e.startswith("Perspektive")]
    x = start_excel()
    out: dict = {"Voreingestellte Farbverläufe": {}, "Textur": {}, "Perspektive": {}}
    try:
        wb = x.wb
        ws = wb.Worksheets(1)
        fill_data(ws)
        sh, ch = column_chart(ws)
        sh.Select()
        ch.ChartArea.Select()
        pane = Pane(x)
        if not pane.open(None):
            data["pane"] = {"error": "pane did not open"}
            return
        time.sleep(0.6)
        tab(pane, "Füllung und Linie")
        expand_section(pane, "Füllung")
        ok = select_radio(pane, "Füllung", "", "Farbverlauf")
        time.sleep(0.6)
        out["Farbverlauf chosen"] = chart_area_sppr(chart_xml(wb)) if ok else "radio not found"
        for nm in presets:
            err = gallery_pick(pane, "Füllung", "Voreingestellte Farbverläufe", nm)
            out["Voreingestellte Farbverläufe"][nm] = {"error": err} if err else chart_area_sppr(chart_xml(wb))
            print(nm, str(out["Voreingestellte Farbverläufe"][nm])[:120], flush=True)
        ok = select_radio(pane, "Füllung", "", "Bild- oder Texturfüllung")
        time.sleep(0.8)
        out["Bild- oder Texturfüllung chosen"] = chart_area_sppr(chart_xml(wb)) if ok else "radio not found"
        for nm in textures:
            err = gallery_pick(pane, "Füllung", "Textur", nm)
            if err:
                out["Textur"][nm] = {"error": err}
                continue
            sp = chart_area_sppr(chart_xml(wb))
            # Office's bitmap stays out: only the fill's tiling and stretch settings.
            out["Textur"][nm] = re.sub(r'r:embed="[^"]*"', 'r:embed="…"', re.sub(r"<a:blip\b[\s\S]*?(</a:blip>|/>)", "<a:blip/>", sp))
            print(nm, out["Textur"][nm][:120], flush=True)
        select_radio(pane, "Füllung", "", "Automatisch")
        tab(pane, "Effekte")
        expand_section(pane, "Schatten")
        for nm in shadows:
            err = gallery_pick(pane, "Schatten", "Voreinstellungen", nm)
            out["Perspektive"][nm] = {"error": err} if err else chart_area_sppr(chart_xml(wb))
            print(nm, str(out["Perspektive"][nm])[:160], flush=True)
        sh.Delete()
    finally:
        data["pane"] = out
        x.close()


def run_dialog(data: dict) -> None:
    """Fehlerindikatoren > Benutzerdefiniert > Wert angeben: the dialog's title, labels and fields (UI Automation)."""
    import win32gui
    import win32process
    x = start_excel()
    out: dict = {}
    try:
        wb = x.wb
        ws = wb.Worksheets(1)
        fill_data(ws)
        sh, ch = column_chart(ws)
        s1 = ch.SeriesCollection(1)
        s1.ErrorBar(1, 1, 1, 1.0)
        sh.Select()
        s1.ErrorBars.Select()
        pane = Pane(x)
        if not pane.open(None):
            out["error"] = "pane did not open"
            return
        time.sleep(0.6)
        sec = next((t for t, _ in pane.sections() if "Fehlerindikator" in t), None)
        if sec is None:
            out["error"] = "no section"
            return
        expand_section(pane, sec)
        select_radio(pane, sec, "", "Benutzerdefiniert")
        time.sleep(0.5)
        btn = next((i for i in raw_section(pane, sec) if i["t"] == "Button" and i["label"] == "Wert angeben"), None)
        if btn is None:
            out["error"] = "no Wert angeben"
            return
        before = set()
        win32gui.EnumWindows(lambda hw, _: before.add(hw), None)
        import threading
        threading.Thread(target=lambda: uia.invoke(btn["_el"]), daemon=True).start()
        dlg = None
        for _ in range(20):
            time.sleep(0.3)
            found: list[int] = []

            def cb(hw, _):
                if hw not in before and win32gui.IsWindowVisible(hw) and win32process.GetWindowThreadProcessId(hw)[1] == x.pid:
                    found.append(hw)
            win32gui.EnumWindows(cb, None)
            if found:
                dlg = found[0]
                break
        if dlg is None:
            out["error"] = "no dialog"
            return
        out["title"] = win32gui.GetWindowText(dlg)
        root = uia.element_for(dlg)
        f = root.FindAll(U.TreeScope_Descendants, UIA.CreateTrueCondition())
        items = []
        for k in range(f.Length):
            d = d_of(f.GetElement(k))
            if d and d["type"] in ("Text", "Edit", "Button", "Group") and d["name"]:
                e = f.GetElement(k)
                item = {"t": d["type"], "name": d["name"]}
                if d["type"] == "Edit":
                    try:
                        item["value"] = e.GetCurrentPattern(U.UIA_ValuePatternId).QueryInterface(U.IUIAutomationValuePattern).CurrentValue
                    except Exception:  # noqa: BLE001
                        pass
                r = e.CurrentBoundingRectangle
                item["rect"] = [r.left, r.top, r.right - r.left, r.bottom - r.top]
                items.append(item)
        out["items"] = items
        wr = win32gui.GetWindowRect(dlg)
        out["window"] = [wr[0], wr[1], wr[2] - wr[0], wr[3] - wr[1]]
        cancel = next((f.GetElement(k) for k in range(f.Length) if (d_of(f.GetElement(k)) or {}).get("name") == "Abbrechen"), None)
        if cancel is not None:
            uia.invoke(cancel)
        time.sleep(0.5)
        sh.Delete()
    finally:
        data["dialog"] = out
        x.close()


def run_none(data: dict) -> None:
    """Keine Füllung and Keine Linie chosen on elements whose automatic look has neither (title, legend, plot area,
    data labels, the value axis' line): the element's c:spPr before and after."""
    x = ExcelTarget(visible=False, place=False)
    out: dict = {}
    try:
        wb = x.wb
        ws = wb.Worksheets(1)
        fill_data(ws)
        sh, ch = column_chart(ws)
        ch.HasTitle = True
        ch.HasLegend = True
        ch.SeriesCollection(1).HasDataLabels = True
        base = chart_xml(wb)
        els = {"title": lambda: ch.ChartTitle.Format, "legend": lambda: ch.Legend.Format, "plotArea": lambda: ch.PlotArea.Format,
               "dLbls": lambda: ch.SeriesCollection(1).DataLabels().Format, "valAx": lambda: ch.Axes(2).Format}
        tag = {"title": "c:title", "legend": "c:legend", "plotArea": "c:plotArea", "dLbls": "c:dLbls", "valAx": "c:valAx"}

        def sp(xml: str, name: str) -> str:
            block = el(xml, tag[name])
            if name == "plotArea":
                # The plot area's own c:spPr follows its axes.
                i = block.rfind("<c:spPr>")
                return block[i:block.find("</c:spPr>", i) + 9] if i >= 0 else ""
            return sppr_of(block)
        for name, f in els.items():
            rec = {"default": sp(base, name)}
            try:
                f().Fill.Visible = 0
                rec["Keine Füllung"] = sp(chart_xml(wb), name)
            except Exception as e:  # noqa: BLE001
                rec["Keine Füllung"] = "error " + str(e)[:60]
            try:
                f().Line.Visible = 0
                rec["Keine Linie"] = sp(chart_xml(wb), name)
            except Exception as e:  # noqa: BLE001
                rec["Keine Linie"] = "error " + str(e)[:60]
            out[name] = rec
        sh.Delete()
    finally:
        data["none"] = out
        x.close()


def main() -> None:
    what = sys.argv[1:] or ["com", "pane"]
    if "none" in what:
        data0 = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
        try:
            run_none(data0)
        finally:
            OUT.write_text(json.dumps(data0, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        return
    if "dialog" in what:
        data0 = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
        try:
            run_dialog(data0)
        finally:
            OUT.write_text(json.dumps(data0, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        return
    data = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    data["meta"] = {"probe": "probe_excel_chart_followup.py", "excel": "Microsoft Excel 365 (German), own instance, no real input"}
    try:
        if "com" in what:
            run_com(data)
        if "pane" in what:
            run_pane(data)
    finally:
        OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        print("wrote", OUT)


if __name__ == "__main__":
    main()
