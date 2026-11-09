"""Probe: what German Word 365 shows around a selected chart (D-129). In our own Word instance (DispatchEx), on a
copy of Word's own chart document (packages/ooxml/test/fixtures/word-charts.docx) in TEMP: the chart is selected
through COM, the contextual tabs Diagrammentwurf and Format are selected through UI Automation (no simulated input)
and every group and control is read with its name, type, access key (keytip) and whether it is enabled; then the
same with the chart made floating (ConvertToShape), and the drop-down menus that open through UI Automation
(ExpandCollapse) are read. Selecting a tab may bring our Word window to the front, so the run holds the real-input
lock (a short hold). Labels and keytips only (no tooltips or other prose): results/<date>-word-chart-ui.json.

    .venv/Scripts/python.exe probe_word_chart_ui.py
"""
from __future__ import annotations

import datetime
import json
import shutil
import tempfile
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.inputlock import real_input
from feellab.targets import WINDOW_RECT

HERE = Path(__file__).parent
FIXTURE = HERE.parent.parent / "packages" / "ooxml" / "test" / "fixtures" / "word-charts.docx"
OUT = HERE / "results" / f"{datetime.date.today()}-word-chart-ui.json"

TABS = ["Diagrammentwurf", "Format"]
MENUS = ["Diagrammelement hinzufügen", "Schnelllayout", "Farben ändern", "Position", "Textumbruch", "Ebene nach vorne",
         "Ebene nach hinten", "Ausrichten", "Drehen", "Fülleffekt", "Formkontur", "Formeffekte"]


def controls_of(group) -> list[dict]:
    out = []
    f = group.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    for i in range(f.Length):
        try:
            d = uia.describe(f.GetElement(i))
        except Exception:  # noqa: BLE001
            continue
        if not d["name"] or d["type"] in ("Text", "Image", "ToolBar", "Pane", "Custom"):
            continue
        out.append({k: d[k] for k in ("type", "name", "key", "enabled")})
    return out


def read_tab(root, name: str) -> dict:
    tab = uia.find(root, lambda x: x["type"] == "TabItem" and x["name"] == name, max_depth=14)
    if tab is None:
        return {"error": "tab not shown"}
    info = uia.describe(tab)
    try:
        uia.select(tab)
    except Exception as e:  # noqa: BLE001
        return {"error": f"select failed: {e}"}
    time.sleep(1.0)
    lower = uia.find(root, lambda x: x["type"] == "Pane" and x["name"].startswith("Unter"), max_depth=12) or root
    groups = []
    f = lower.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    for i in range(f.Length):
        el = f.GetElement(i)
        try:
            d = uia.describe(el)
        except Exception:  # noqa: BLE001
            continue
        if d["type"] == "Group" and d["name"] and d["rect"][2] > 0:
            groups.append({"name": d["name"], "key": d["key"], "controls": controls_of(el)})
    return {"key": info["key"], "groups": groups}


def read_menu(root, name: str) -> list[str] | dict:
    el = uia.find(root, lambda x: x["name"] == name and x["type"] in ("MenuItem", "SplitButton", "Button") and x["rect"][2] > 0, max_depth=16)
    if el is None:
        return {"error": "control not found"}
    try:
        p = el.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId)
        pat = p.QueryInterface(uia.U.IUIAutomationExpandCollapsePattern)
        pat.Expand()
    except Exception as e:  # noqa: BLE001
        return {"error": f"no expand: {e}"}
    time.sleep(0.9)
    items = []
    f = el.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    for i in range(f.Length):
        try:
            d = uia.describe(f.GetElement(i))
        except Exception:  # noqa: BLE001
            continue
        if d["name"] and d["type"] in ("MenuItem", "Button", "ListItem", "CheckBox", "RadioButton", "Group"):
            items.append(f"{d['type']}:{d['name']}" + (f" [{d['key']}]" if d["key"] else "") + ("" if d["enabled"] else " (aus)"))
    try:
        pat.Collapse()
    except Exception:  # noqa: BLE001
        winapi.press("esc")
    time.sleep(0.4)
    return items


def main() -> None:
    import win32com.client
    tmp = Path(tempfile.gettempdir()) / "raabe-word-chart-ui"
    tmp.mkdir(exist_ok=True)
    doc = tmp / "chart-ui.docx"
    shutil.copy(FIXTURE, doc)
    result: dict = {"source": "word-charts.docx (Word's own chart document, make_fixtures.py --charts)"}
    app = win32com.client.DispatchEx("Word.Application")
    pid = None
    try:
        with real_input("write-charts word chart ui probe"):
            app.Visible = True
            d = app.Documents.Open(str(doc))
            win = d.Windows(1)
            win.View.Type = 3
            winapi.place(win.Hwnd, *WINDOW_RECT)
            pid = winapi.window_pid(win.Hwnd)
            time.sleep(1.0)
            root = uia.element_for(win.Hwnd)
            result["inline_count"] = d.InlineShapes.Count
            d.InlineShapes(1).Select()
            time.sleep(0.8)
            result["inline"] = {t: read_tab(root, t) for t in TABS}
            result["inline_menus"] = {m: read_menu(root, m) for m in MENUS}
            # Floating: the chart as a shape with Quadrat.
            sh = d.InlineShapes(1).ConvertToShape()
            sh.WrapFormat.Type = 0
            sh.Select()
            time.sleep(0.8)
            result["float"] = {t: read_tab(root, t) for t in TABS}
            result["float_menus"] = {m: read_menu(root, m) for m in ["Position", "Textumbruch", "Ebene nach vorne", "Ebene nach hinten"]}
            result["float_com"] = {"left": sh.Left, "top": sh.Top, "wrap": sh.WrapFormat.Type, "hasChart": bool(sh.HasChart)}
            d.Close(0)
    finally:
        try:
            app.Quit(0)
        except Exception:  # noqa: BLE001
            pass
    del pid
    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("wrote", OUT)


if __name__ == "__main__":
    main()
