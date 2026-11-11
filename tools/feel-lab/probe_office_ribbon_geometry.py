"""Probe (Step 11, part 5): where Office's ribbon puts every group and control. Reads the Start tab of our own Excel,
Word and PowerPoint (the window at the harness's place and size, so the ribbon has the same width as our apps') through
UI Automation, no input: each group's and control's type, name and box in CSS pixels (screen pixels / the display's
scale) relative to the ribbon's top left. PowerPoint only when the owner's POWERPNT.EXE is not running.
Writes results/<date>-office-ribbon-geometry.json (Microsoft's labels only).

    .venv/Scripts/python.exe probe_office_ribbon_geometry.py [excel] [word] [powerpoint]
"""
from __future__ import annotations

import datetime
import json
import sys
import tempfile
import time
from pathlib import Path

from feellab import uia, winapi

HERE = Path(__file__).parent
KINDS = ("Button", "SplitButton", "MenuItem", "ComboBox", "CheckBox", "Edit", "Group", "List", "DataGrid")


# The main tabs each app shares with ours (contextual tabs need a selection; Datei is the backstage).
TABS = {
    "excel": ["Start", "Einfügen", "Seitenlayout", "Formeln", "Daten", "Überprüfen", "Ansicht"],
    "word": ["Start", "Einfügen", "Entwurf", "Layout", "Referenzen", "Sendungen", "Überprüfen", "Ansicht"],
    "powerpoint": ["Start", "Einfügen", "Entwurf", "Übergänge", "Animationen", "Bildschirmpräsentation", "Überprüfen", "Ansicht"],
}


def ribbon_of(root, name: str = "Start"):
    """The groups of one ribbon tab (selected through UI Automation, no input)."""
    tab = uia.find(root, lambda d: d["type"] == "TabItem" and d["name"] == name and d["class"] == "NetUIRibbonTab", max_depth=12)
    if not tab:
        return []
    uia.select(tab)
    time.sleep(0.8)
    # The groups: Group elements of class NetUIChunk; the panel is their common box.
    groups = []
    stack = [(root, 0)]
    while stack:
        e, d = stack.pop()
        try:
            dc = uia.describe(e)
            if dc["type"] == "Group" and dc["class"] == "NetUIChunk":
                groups.append((e, dc))
                continue
            if d < 12:
                stack.extend((c, d + 1) for c in uia.children(e))
        except Exception:  # noqa: BLE001
            continue
    return groups


def controls(group) -> list[dict]:
    out = []
    stack = [(group, 0)]
    while stack:
        e, d = stack.pop()
        for c in uia.children(e):
            try:
                dc = uia.describe(c)
            except Exception:  # noqa: BLE001
                continue
            if dc["type"] in KINDS and dc["name"] and dc["rect"][2] > 0:
                out.append({"type": dc["type"], "name": dc["name"], "rect": dc["rect"]})
            # Split buttons: their own box is the control; galleries and lists hold many entries: keep the outer box.
            if dc["type"] not in ("SplitButton", "List", "DataGrid", "ComboBox") and d < 6:
                stack.append((c, d + 1))
    return out


# Wide enough that every group shows its full layout (Office shrinks groups to make room for its add-in groups,
# which our apps do not have): the comparison is then between full layouts on both sides.
WIDE = (0, 0, 3000, 2050)


def record(root, hwnd: int, app: str = "") -> dict:
    """Start at the top level (as before) and every tab of TABS under "tabs"."""
    winapi.place(hwnd, *WIDE)
    time.sleep(1.0)
    out = record_tab(root, hwnd, "Start")
    out["tabs"] = {t: record_tab(root, hwnd, t) for t in TABS.get(app, ["Start"])}
    ribbon_of(root, "Start")
    return out


def record_tab(root, hwnd: int, name: str) -> dict:
    scale = winapi.dpi_scale(hwnd)
    groups = ribbon_of(root, name)
    if not groups:
        return {"error": "no ribbon groups"}
    left = min(g[1]["rect"][0] for g in groups)
    top = min(g[1]["rect"][1] for g in groups)
    css = lambda r: [round((r[0] - left) / scale, 1), round((r[1] - top) / scale, 1), round(r[2] / scale, 1), round(r[3] / scale, 1)]  # noqa: E731
    out = {"scale": scale, "window": list(WIDE), "groups": []}
    for g, dg in sorted(groups, key=lambda x: x[1]["rect"][0]):
        out["groups"].append({"name": dg["name"], "rect": css(dg["rect"]),
                              "controls": [{**c, "rect": css(c["rect"])} for c in controls(g)]})
    return out


# ----- contextual tabs (D-185) ------------------------------------------------------------------------------------------
# Each kind of selection and the contextual tabs Office shows for it, by their UI Automation names (Word's and PowerPoint's table Layout
# tab is "Tabellenlayout" there; its visible label is "Layout"). The selection is made through COM in our own document.
CONTEXT = {
    "excel": {"picture": ["Bildformat"], "shape": ["Formformat"], "table": ["Tabellenentwurf"], "chart": ["Diagrammentwurf", "Format"]},
    "word": {"picture": ["Bildformat"], "shape": ["Formformat"], "table": ["Tabellenentwurf", "Tabellenlayout"], "chart": ["Diagrammentwurf", "Format"]},
    "powerpoint": {"picture": ["Bildformat"], "shape": ["Formformat"], "table": ["Tabellenentwurf", "Tabellenlayout"], "chart": ["Diagrammentwurf", "Format"]},
}
CROPS = Path(tempfile.gettempdir()) / "raabe-ribbon-ctx"


def sample_png() -> str:
    from PIL import Image
    CROPS.mkdir(parents=True, exist_ok=True)
    p = CROPS / "sample.png"
    im = Image.new("RGB", (240, 160), (70, 130, 180))
    for x in range(60, 180):
        for y in range(40, 120):
            im.putpixel((x, y), (240, 200, 60))
    im.save(p)
    return str(p)


def context_tab(root, name: str):
    """A contextual tab: a NetUIRibbonTab inside a NetUITabGroup (Bildtools, Tabellentools …)."""
    walker = uia._uia.ControlViewWalker
    found = []
    stack = [(root, 0)]
    while stack:
        e, d = stack.pop()
        try:
            dc = uia.describe(e)
        except Exception:  # noqa: BLE001
            continue
        if dc["type"] == "TabItem" and dc["class"] == "NetUIRibbonTab":
            if dc["name"] == name:
                par = uia.describe(walker.GetParentElement(e))
                if par["class"] == "NetUITabGroup":
                    found.append((e, par["name"]))
            continue
        if d < 12:
            stack.extend((c, d + 1) for c in uia.children(e))
    return found[0] if found else (None, None)


def crop_ribbon(hwnd: int, rec: dict, file: str) -> None:
    """A PrintWindow picture of our own Office window (no screen capture), cut to the ribbon; stays in TEMP."""
    from probe_excel_picturedrops import print_window
    from PIL import Image
    full = CROPS / "window.png"
    if not print_window(hwnd, full) or not rec.get("groups"):
        return
    l, t, _, _ = winapi.window_rect(hwnd)
    s = rec["scale"]
    x0, y0 = rec["origin"]
    w = max(g["rect"][0] + g["rect"][2] for g in rec["groups"]) * s
    h = max(g["rect"][1] + g["rect"][3] for g in rec["groups"]) * s
    Image.open(full).crop((int(x0 - l), int(y0 - t), int(x0 - l + w), int(y0 - t + h))).save(CROPS / file)
    full.unlink(missing_ok=True)


def record_context_tab(root, hwnd: int, name: str) -> dict:
    tab, group = context_tab(root, name)
    if tab is None:
        return {"error": "contextual tab not shown"}
    uia.select(tab)
    time.sleep(1.0)
    scale = winapi.dpi_scale(hwnd)
    groups = []
    stack = [(root, 0)]
    while stack:
        e, d = stack.pop()
        try:
            dc = uia.describe(e)
            if dc["type"] == "Group" and dc["class"] == "NetUIChunk":
                groups.append((e, dc))
                continue
            if d < 12:
                stack.extend((c, d + 1) for c in uia.children(e))
        except Exception:  # noqa: BLE001
            continue
    if not groups:
        return {"error": "no ribbon groups"}
    left = min(g[1]["rect"][0] for g in groups)
    top = min(g[1]["rect"][1] for g in groups)
    css = lambda r: [round((r[0] - left) / scale, 1), round((r[1] - top) / scale, 1), round(r[2] / scale, 1), round(r[3] / scale, 1)]  # noqa: E731
    out = {"scale": scale, "window": list(WIDE), "tabGroup": group, "origin": [left, top], "groups": []}
    for g, dg in sorted(groups, key=lambda x: x[1]["rect"][0]):
        out["groups"].append({"name": dg["name"], "rect": css(dg["rect"]),
                              "controls": [{**c, "rect": css(c["rect"])} for c in controls(g)]})
    return out


def record_context(app: str, root, hwnd: int, select) -> dict:
    """For each kind of selection: make it (`select(kind)`, COM) and record its contextual tabs."""
    winapi.place(hwnd, *WIDE)
    time.sleep(1.0)
    res: dict = {}
    for kind, names in CONTEXT[app].items():
        try:
            select(kind)
        except Exception as e:  # noqa: BLE001
            for n in names:
                res[n] = {"kind": kind, "error": f"selection failed: {str(e)[:120]}"}
            continue
        time.sleep(1.5)
        for n in names:
            rec = record_context_tab(root, hwnd, n)
            rec["kind"] = kind
            if "groups" in rec:
                crop_ribbon(hwnd, rec, f"{app}-{n}.png")
            rec.pop("origin", None)
            res[n] = rec
            print(app, kind, n, rec.get("error") or f"{len(rec['groups'])} groups")
    ribbon_of(root, "Start")
    return res


def excel_context(out: dict) -> None:
    from feellab.excel import ExcelTarget
    png = sample_png()
    x = ExcelTarget(visible=True)
    try:
        time.sleep(1.5)
        ws = x.wb.Worksheets(1)
        for r, row in enumerate([["", "A", "B"], ["x", 1, 2], ["y", 3, 4], ["z", 5, 6]], 1):
            for c, v in enumerate(row, 1):
                ws.Cells(r, c).Value = v
        objs: dict = {}

        def select(kind: str) -> None:
            ws.Range("H20").Select()
            if kind == "picture":
                objs.setdefault(kind, ws.Shapes.AddPicture(png, False, True, 300, 200, 120, 80)).Select()
            elif kind == "shape":
                objs.setdefault(kind, ws.Shapes.AddShape(1, 450, 200, 100, 60)).Select()
            elif kind == "table":
                if kind not in objs:
                    objs[kind] = ws.ListObjects.Add(1, ws.Range("A1:C4"), None, 1)
                ws.Range("B2").Select()
            elif kind == "chart":
                if kind not in objs:
                    objs[kind] = ws.Shapes.AddChart2(-1, 51, 300, 300, 360, 216)
                    objs[kind].Chart.SetSourceData(ws.Range("A1:C4"))
                objs[kind].Select()

        out.setdefault("excel", {})["context"] = record_context("excel", uia.element_for(x.hwnd), x.hwnd, select)
    finally:
        x.close()


def word_context(out: dict) -> None:
    import subprocess
    from feellab.targets import WordTarget
    png = sample_png()
    w = WordTarget()
    pid = 0
    try:
        doc = w.doc
        doc.Content.Text = "Absatz eins\rAbsatz zwei\rAbsatz drei\r"
        top = winapi.user32.GetAncestor(int(w.app.ActiveWindow.Hwnd), 2)
        pid = winapi.window_pid(top)
        objs: dict = {}

        def select(kind: str) -> None:
            if kind == "picture":
                objs.setdefault(kind, doc.InlineShapes.AddPicture(png, False, True, doc.Paragraphs(1).Range)).Select()
            elif kind == "shape":
                objs.setdefault(kind, doc.Shapes.AddShape(1, 300, 300, 100, 60)).Select()
            elif kind == "table":
                if kind not in objs:
                    objs[kind] = doc.Tables.Add(doc.Paragraphs(2).Range, 3, 3)
                objs[kind].Cell(1, 1).Range.Select()
                doc.Application.Selection.Collapse(1)
            elif kind == "chart":
                if kind not in objs:
                    ish = doc.InlineShapes.AddChart2(-1, 51, doc.Paragraphs(doc.Paragraphs.Count).Range)
                    try:
                        ish.Chart.ChartData.Activate()
                        ish.Chart.ChartData.Workbook.Close()
                    except Exception:  # noqa: BLE001
                        pass
                    objs[kind] = ish
                    winapi.place(top, *WIDE)
                objs[kind].Select()

        out.setdefault("word", {})["context"] = record_context("word", uia.element_for(top), top, select)
    finally:
        q = f"Get-CimInstance Win32_Process -Filter \"name='EXCEL.EXE' and ParentProcessId={pid}\" | ForEach-Object {{ $_.ProcessId }}"
        kids = subprocess.run(["powershell", "-NoProfile", "-Command", q], capture_output=True, text=True).stdout.split() if pid else []
        w.close()
        time.sleep(3)
        for k in kids:
            subprocess.run(["taskkill", "/PID", k, "/F"], capture_output=True)


def powerpoint_context(out: dict) -> None:
    import subprocess
    from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget
    png = sample_png()
    try:
        p = PowerPointTarget(visible=True)
    except OwnerPowerPointRunning as e:
        print("PowerPoint skipped:", e)
        return
    try:
        pres = p.new_presentation(with_window=True, name="ribbonctx")
        slide = pres.Slides.Add(1, 12)
        p.app.ActiveWindow.ViewType = 9  # ppViewNormal
        objs: dict = {}

        def select(kind: str) -> None:
            if kind == "picture":
                objs.setdefault(kind, slide.Shapes.AddPicture(png, 0, -1, 60, 60, 160, 107)).Select()
            elif kind == "shape":
                objs.setdefault(kind, slide.Shapes.AddShape(1, 300, 60, 120, 80)).Select()
            elif kind == "table":
                if kind not in objs:
                    objs[kind] = slide.Shapes.AddTable(3, 3, 60, 220, 400, 120)
                objs[kind].Select()
            elif kind == "chart":
                if kind not in objs:
                    sh = slide.Shapes.AddChart2(-1, 51, 480, 220, 360, 216)
                    try:
                        sh.Chart.ChartData.Activate()
                        sh.Chart.ChartData.Workbook.Close()
                    except Exception:  # noqa: BLE001
                        pass
                    objs[kind] = sh
                    winapi.place(p.hwnd, *WIDE)
                objs[kind].Select()

        out.setdefault("powerpoint", {})["context"] = record_context("powerpoint", uia.element_for(p.hwnd), p.hwnd, select)
    finally:
        q = f"Get-CimInstance Win32_Process -Filter \"name='EXCEL.EXE' and ParentProcessId={p.pid}\" | ForEach-Object {{ $_.ProcessId }}"
        kids = subprocess.run(["powershell", "-NoProfile", "-Command", q], capture_output=True, text=True).stdout.split() if p.pid else []
        p.close(kill_if_unreachable=True)
        time.sleep(3)
        for k in kids:
            subprocess.run(["taskkill", "/PID", k, "/F"], capture_output=True)


def main_context(which: list[str]) -> None:
    path = HERE / "results" / "2026-11-11-office-ribbon-geometry.json"
    out = json.loads(path.read_text(encoding="utf-8"))
    for app, fn in (("excel", excel_context), ("word", word_context), ("powerpoint", powerpoint_context)):
        if app in which:
            fn(out)
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("crops (not committed) in", CROPS)


def main() -> None:
    which = [a for a in sys.argv[1:] if not a.startswith("-")] or ["excel", "word", "powerpoint"]
    if "--context" in sys.argv:
        main_context(which)
        return
    path = HERE / "results" / f"{datetime.date.today().isoformat()}-office-ribbon-geometry.json"
    out = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    if "excel" in which:
        from feellab.excel import ExcelTarget
        x = ExcelTarget(visible=True)
        try:
            time.sleep(1.5)
            out["excel"] = record(uia.element_for(x.hwnd), x.hwnd, "excel")
        finally:
            x.close()
    if "word" in which:
        from feellab.targets import WordTarget
        w = WordTarget()
        try:
            top = winapi.user32.GetAncestor(int(w.app.ActiveWindow.Hwnd), 2)
            time.sleep(1.5)
            out["word"] = record(uia.element_for(top), top, "word")
        finally:
            w.close()
    if "powerpoint" in which:
        from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget
        try:
            p = PowerPointTarget(visible=True)
        except OwnerPowerPointRunning as e:
            print("PowerPoint skipped:", e)
        else:
            try:
                p.new_presentation(with_window=True)
                time.sleep(1.5)
                out["powerpoint"] = record(uia.element_for(p.hwnd), p.hwnd, "powerpoint")
            finally:
                for pr in p.own:
                    pr.Close()
                if p.app.Presentations.Count == 0:
                    p.app.Quit()
    path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    for k, v in out.items():
        print(k, len(v.get("groups", [])), "groups", sum(len(g["controls"]) for g in v.get("groups", [])), "controls")


if __name__ == "__main__":
    main()
