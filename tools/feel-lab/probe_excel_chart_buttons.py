"""Records what German Excel 365 shows around a selected chart (D-115): the three buttons beside it
(Diagrammelemente, Diagrammformatvorlagen, Diagrammfilter) with their flyouts and side menus, the context
menus of the chart and its elements, the Datenquelle auswählen dialog with its sub-dialogs, the Format tab
and the colours of the source ranges Excel marks in the sheet.

Our own Excel (`ExcelTarget`, DispatchEx), a workbook made here, nothing else on screen is read. The side
buttons do not answer UI Automation's Invoke, so the flyouts are opened with real input: the probe takes the
real-input lock, puts its Excel in front, and before every click checks that the window in front belongs to
its own Excel process. Labels, kinds, positions and sizes only (no Microsoft prose); full dumps go to
captures/ (not committed).

    .venv/Scripts/python.exe probe_excel_chart_buttons.py [pass ...]   (passes: buttons context dialog ribbon ranges)

Writes results/2026-11-09-excel-chart-buttons.json (merged with what an earlier run recorded)."""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.excel import ExcelTarget
from feellab.inputlock import real_input, touch

HERE = Path(__file__).parent
RESULT = HERE / "results" / "2026-11-09-excel-chart-buttons.json"
CAPTURES = HERE / "captures" / "chart-buttons"
DATA = [("Monat", "Nord", "Süd", "West"), ("Jan", 3, 4, 2), ("Feb", 5, 2, 3), ("Mrz", 2, 6, 4), ("Apr", 4, 3, 5)]
WIN = (40, 40, 2400, 1500)


def descendants(root) -> list[tuple]:
    out = []
    try:
        f = root.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    except Exception:  # noqa: BLE001
        return out
    for i in range(f.Length):
        e = f.GetElement(i)
        try:
            out.append((e, uia.describe(e)))
        except Exception:  # noqa: BLE001
            continue
    return out


POPUP_CLASSES = ("NUIDialog", "Net UI Tool Window", "Net UI Tool Window Layered", "#32768", "bosa_sdm_XL9", "NetUIToolWindow")


def windows(pid: int) -> list[tuple]:
    """Every top-level window of our Excel and the popups Excel keeps as children of its main window."""
    out = []
    root = uia._uia.GetRootElement()
    cond = uia._uia.CreatePropertyCondition(uia.U.UIA_ProcessIdPropertyId, pid)
    f = root.FindAll(uia.U.TreeScope_Children, cond)
    for i in range(f.Length):
        e = f.GetElement(i)
        try:
            d = uia.describe(e)
        except Exception:  # noqa: BLE001
            continue
        if d["class"] == "MSO_BORDEREFFECT_WINDOW_CLASS":
            continue
        out.append((e, d))
        if d["class"] == "XLMAIN":
            for e2, d2 in descendants(e):
                if d2["class"] in POPUP_CLASSES or (d2["type"] in ("Window", "Menu") and d2["class"] not in ("XLDESK", "EXCEL7", "EXCEL;", "EXCEL2", "EXCEL<")):
                    out.append((e2, d2))
    return out


def key(d: dict) -> str:
    return f"{d['class']}|{d['rect']}"


class Probe:
    def __init__(self) -> None:
        self.x = ExcelTarget(visible=False)
        self.pid = self.x.pid
        self.result: dict = json.loads(RESULT.read_text(encoding="utf-8")) if RESULT.exists() else {}
        self.raw: dict = {}
        ws = self.ws = self.x.wb.Worksheets(1)
        for r, row in enumerate(DATA, 1):
            for c, v in enumerate(row, 1):
                ws.Cells(r, c).Value = v
        self.shape = ws.Shapes.AddChart2(201, 51, 330, 15, 360, 216)
        self.chart = self.shape.Chart
        self.chart.SetSourceData(ws.Range("A1:D5"))

    # ----- input -------------------------------------------------------------------------------
    def ours(self) -> bool:
        fg = winapi.foreground()
        return bool(fg) and winapi.window_pid(fg) == self.pid

    def front(self) -> None:
        if not self.ours():
            winapi.activate(self.x.hwnd, 0.3)
        if not self.ours():
            raise RuntimeError("our Excel is not in front")

    def click(self, x: float, y: float, right: bool = False) -> None:
        if not self.ours():
            raise RuntimeError("window in front is not ours")
        (winapi.right_click if right else winapi.click)(x, y)

    def esc(self, n: int = 1) -> None:
        for _ in range(n):
            if not self.ours():
                return
            winapi.press("esc")
            time.sleep(0.25)

    # ----- reading -----------------------------------------------------------------------------
    def popups(self, before: set[str]) -> list[dict]:
        """The elements of windows that appeared since `before` (flyouts, menus, dialogs)."""
        out = []
        for e, d in windows(self.pid):
            if d["class"] == "XLMAIN" or key(d) in before:
                continue
            items = [d] + [dd for _, dd in descendants(e)]
            out.append({"window": d, "items": items})
        return out

    def snapshot(self) -> set[str]:
        return {key(d) for _, d in windows(self.pid)}

    def side_buttons(self) -> dict[str, dict]:
        for _ in range(15):
            out = self._side_buttons()
            if len(out) >= 3:
                return out
            time.sleep(0.3)
        return out

    def _side_buttons(self) -> dict[str, dict]:
        out = {}
        for e, d in windows(self.pid):
            if d["class"] != "NUIDialog":
                continue
            for e2, d2 in descendants(e):
                if d2["type"] == "Button" and d2["class"] == "NetUIStickyButton":
                    out[d2["name"]] = d2
        return out

    def select_chart(self) -> None:
        self.ws.Range("H20").Select()
        time.sleep(0.2)
        self.shape.Select()
        time.sleep(0.8)

    def chart_screen(self) -> dict:
        root = uia.element_for(self.x.hwnd)
        el = uia.find(root, lambda d: d["type"] == "Image" and d["name"].startswith("Diagramm"), 30)
        return uia.describe(el) if el else {}

    # ----- passes ------------------------------------------------------------------------------
    def pass_buttons(self) -> None:
        self.select_chart()
        chart = self.chart_screen()
        btns = self.side_buttons()
        scale = winapi.dpi_scale(self.x.hwnd)
        self.result["scale"] = scale
        self.result["chart_rect"] = chart.get("rect")
        self.result["side_buttons"] = [{"uia_name": n, "rect": d["rect"],
                                         "css": {"dx_from_chart_right": round((d["rect"][0] - (chart["rect"][0] + chart["rect"][2])) / scale, 1),
                                                 "dy_from_chart_top": round((d["rect"][1] - chart["rect"][1]) / scale, 1),
                                                 "w": round(d["rect"][2] / scale, 1), "h": round(d["rect"][3] / scale, 1)}}
                                        for n, d in sorted(btns.items(), key=lambda kv: kv[1]["rect"][1])]
        print("side buttons", list(btns), flush=True)
        flyouts = {}
        for name, d in sorted(btns.items(), key=lambda kv: kv[1]["rect"][1]):
            touch()
            self.front()
            before = self.snapshot()
            cx, cy = d["rect"][0] + d["rect"][2] / 2, d["rect"][1] + d["rect"][3] / 2
            self.click(cx, cy)
            time.sleep(1.2)
            pops = self.popups(before)
            CAPTURES.mkdir(parents=True, exist_ok=True)
            winapi.save_region_png(str(CAPTURES / f"flyout-{d['rect'][1]}.png"), chart["rect"][0], chart["rect"][1] - 100, 1400, 1000)
            flyouts[name] = {"main": pops}
            # Tabs inside the flyout (Formatvorlage / Farbe, Werte / Namen): click each one and read again.
            tabs = [it for p in pops for it in p["items"] if it["type"] == "TabItem"]
            for t in tabs:
                before2 = self.snapshot()
                self.click(t["rect"][0] + t["rect"][2] / 2, t["rect"][1] + t["rect"][3] / 2)
                time.sleep(1.0)
                flyouts[name][f"tab:{t['name']}"] = self.popups(before)
                del before2
            self.esc(2)
            self.select_chart()
        self.raw["flyouts"] = flyouts
        self.result["flyouts"] = {n: summarize(v) for n, v in flyouts.items()}

    def pass_subs(self) -> None:
        """Diagrammelemente's side menus: for each entry the flyout is opened afresh, the entry hovered and the
        arrow at its right end clicked (the arrows sit in a column of their own); the submenu is read."""
        self.select_chart()
        btns = self.side_buttons()
        plus = next(d for n, d in btns.items() if "Element" in n)
        subs = {}
        names = None
        i = 0
        tries = 0
        while True:
            touch()
            self.select_chart()
            try:
                self.front()
            except RuntimeError:
                tries += 1
                if tries > 3:
                    raise
                time.sleep(2)
                continue
            before = self.snapshot()
            self.click(plus["rect"][0] + plus["rect"][2] / 2, plus["rect"][1] + plus["rect"][3] / 2)
            time.sleep(1.0)
            pops = self.popups(before)
            gal = [p for p in pops if p["window"]["class"] == "NetUIGalleryContainer"]
            entries = [it for p in gal[:1] for it in p["items"] if it["type"] == "MenuItem"]
            if names is None:
                names = [e["name"] for e in entries]
            if i >= len(entries):
                self.esc(2)
                break
            it = entries[i]
            arrow_col = gal[1]["window"]["rect"] if len(gal) > 1 else [it["rect"][0] + it["rect"][2], 0, 46, 0]
            r = it["rect"]
            winapi.move(r[0] + r[2] / 2, r[1] + r[3] / 2)
            time.sleep(0.6)
            hover = [{k: a[k] for k in ("type", "name", "rect")} for p in self.popups(before) for a in p["items"]
                     if a["rect"][0] >= arrow_col[0] - 2 and abs(a["rect"][1] - r[1]) < 6 and a["type"] != "Menu"]
            before3 = self.snapshot()
            self.click(arrow_col[0] + arrow_col[2] / 2, r[1] + r[3] / 2)
            time.sleep(1.0)
            menu = self.popups(before3)
            CAPTURES.mkdir(parents=True, exist_ok=True)
            winapi.save_region_png(str(CAPTURES / f"sub-{i}.png"), r[0] - 900, r[1] - 300, 1800, 900)
            subs[it["name"]] = {"hover": hover, "menu": menu}
            print("sub", it["name"], len(menu), flush=True)
            self.esc(3)
            i += 1
        self.raw["subs"] = subs
        self.result["element_submenus"] = {n: menu_items(v["menu"]) for n, v in subs.items()}

    def pass_context(self) -> None:
        """Right-click on each chart element UI Automation lists inside the chart: their context menus (labels only)."""
        self.select_chart()
        root = uia.element_for(self.x.hwnd)
        img = uia.find(root, lambda d: d["type"] == "Image" and d["name"].startswith("Diagramm"), 30)
        parts = [d for _, d in descendants(img)] if img else []
        self.result["chart_parts"] = [{k: d[k] for k in ("type", "name", "rect")} for d in parts]
        print("parts", [(d["type"], d["name"]) for d in parts], flush=True)
        menus = {}
        seen = set()
        for d in parts:
            nm = d["name"]
            if not nm or nm in seen or d["rect"][2] <= 2 or "Punkt" in nm or "Achsenbeschriftung" in nm or nm in ('Reihe "Süd"', 'Reihe "West"')                     or "Süd" in nm or "West" in nm:
                continue
            seen.add(nm)
            try:
                menus[nm] = self._context_at(d)
            except Exception as ex:  # noqa: BLE001
                menus[nm] = []
                print("context failed", nm, ex, flush=True)
                self.esc(2)
        self.raw["context"] = menus
        self.result["context_menus"] = {n: menu_items(v) for n, v in menus.items()}

    def _context_at(self, d: dict) -> list:
        nm = d["name"]
        if True:
            r = d["rect"]
            # The chart area: near its lower left corner (its centre is the plot area); others: the centre.
            px, py = (r[0] + 8, r[1] + r[3] - 8) if nm == "Diagrammbereich" else (r[0] + r[2] / 2, r[1] + r[3] / 2)
            if nm == "Zeichnungsfläche":
                px, py = r[0] + r[2] * 0.5, r[1] + r[3] * 0.08
            if nm.startswith("Hauptgitternetz") or "Gitternetz" in nm:
                px, py = r[0] + r[2] * 0.9, r[1] + 1
            touch()
            self.select_chart()
            self.front()
            before = self.snapshot()
            self.click(px, py, right=True)
            time.sleep(1.0)
            out = self.popups(before)
            self.esc(2)
            return out

    def pass_dialog(self) -> None:
        """Datenquelle auswählen and its sub-dialogs, opened from the ribbon (UI Automation) and their buttons."""
        self.select_chart()
        root = uia.element_for(self.x.hwnd)
        tab = uia.find(root, lambda d: d["type"] == "TabItem" and d["name"] == "Diagrammentwurf", 25)
        uia.select(tab)
        time.sleep(0.8)
        self.front()
        btn = uia.find(root, lambda d: d["name"].startswith("Daten auswählen") and d["type"] == "Button", 30)
        before = self.snapshot()
        r = uia.describe(btn)["rect"]
        self.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
        time.sleep(1.5)
        dlg = self.popups(before)
        out = {"main": dlg}
        top = [p for p in dlg if p["window"]["type"] == "Window"]
        self.result["select_data"] = {"window": top[0]["window"] if top else None, "controls": controls(dlg)}
        subs = {}
        for label in ("Hinzufügen", "Bearbeiten", "Ausgeblendete und leere Zellen", "Kategorie bearbeiten"):
            items = [it for p in dlg for it in p["items"] if it["type"] == "Button" and (it["name"] == label or it["name"].startswith(label))]
            if label == "Hinzufügen":
                items = [it for p in dlg for it in p["items"] if it["type"] == "Button" and it["name"] in ("Hinzufügen", "Datenreihe hinzufügen")]
            if label == "Bearbeiten":
                items = [it for p in dlg for it in p["items"] if it["type"] == "Button" and it["name"] in ("Bearbeiten", "Datenreihe bearbeiten")]
            if label == "Kategorie bearbeiten":
                items = [it for p in dlg for it in p["items"] if it["type"] == "Button" and it["name"] in ("Kategorie bearbeiten",)]
            if not items:
                subs[label] = "not found"
                continue
            touch()
            b = items[0]
            before2 = self.snapshot()
            self.click(b["rect"][0] + b["rect"][2] / 2, b["rect"][1] + b["rect"][3] / 2)
            time.sleep(1.2)
            sub = self.popups(before2)
            subs[label] = sub
            self.esc()
            time.sleep(0.6)
        # The collapse button of Diagrammdatenbereich (Dialog reduzieren): what the dialog becomes.
        col = [it for p in dlg for it in p["items"] if it["type"] == "Button" and ("reduz" in it["name"].lower() or "Dialog" in it["name"])]
        if col:
            b = col[0]
            before2 = self.snapshot()
            self.click(b["rect"][0] + b["rect"][2] / 2, b["rect"][1] + b["rect"][3] / 2)
            time.sleep(1.0)
            subs["collapse"] = [p for p in self.popups(before) if True]
            self.esc()
            time.sleep(0.5)
        self.esc()
        time.sleep(0.5)
        out["subs"] = subs
        self.raw["dialog"] = out
        self.result["select_data"]["subs"] = {k: (controls(v) if isinstance(v, list) else v) for k, v in subs.items()}

    def pass_ribbon(self) -> None:
        """The Format tab and the Diagrammentwurf tab's controls (UI Automation, no input)."""
        self.select_chart()
        root = uia.element_for(self.x.hwnd)
        out = {}
        for tabname in ("Diagrammentwurf", "Format"):
            tab = uia.find(root, lambda d, t=tabname: d["type"] == "TabItem" and d["name"] == t, 25)
            if not tab:
                out[tabname] = "not found"
                continue
            uia.select(tab)
            time.sleep(1.0)
            out[tabname] = {"key": uia.describe(tab)["key"],
                            "controls": [{k: d[k] for k in ("type", "name", "key", "enabled", "rect")} for _, d in descendants(root)
                                         if 100 < d["rect"][1] - winapi.window_rect(self.x.hwnd)[1] < 300 and d["name"]
                                         and d["type"] not in ("ListItem", "Pane", "Text")]}
        self.result["ribbon"] = out
        self.front()

    def pass_ranges(self) -> None:
        """The source ranges Excel marks while the chart (and one series) is selected: a screenshot of A1:D5."""
        CAPTURES.mkdir(parents=True, exist_ok=True)
        self.front()
        self.select_chart()
        rng = self.ws.Range("A1:E6")
        win = self.x.app.ActiveWindow
        x0, y0 = win.PointsToScreenPixelsX(rng.Left), win.PointsToScreenPixelsY(rng.Top)
        x1, y1 = win.PointsToScreenPixelsX(rng.Left + rng.Width), win.PointsToScreenPixelsY(rng.Top + rng.Height)
        winapi.save_region_png(str(CAPTURES / "ranges-chart.png"), x0, y0, x1 - x0, y1 - y0)
        self.chart.FullSeriesCollection(2).Select()
        time.sleep(0.6)
        winapi.save_region_png(str(CAPTURES / "ranges-series2.png"), x0, y0, x1 - x0, y1 - y0)
        self.result["ranges_capture"] = {"cells": "A1:E6", "px": [x0, y0, x1 - x0, y1 - y0]}
        cols = {}
        for cell in ("A2", "B1", "B2", "C3", "D5", "A5"):
            c = self.ws.Range(cell)
            cols[cell] = [win.PointsToScreenPixelsX(c.Left) - x0, win.PointsToScreenPixelsY(c.Top) - y0,
                          win.PointsToScreenPixelsX(c.Left + c.Width) - x0, win.PointsToScreenPixelsY(c.Top + c.Height) - y0]
        self.result["ranges_cells"] = cols

    def close(self) -> None:
        RESULT.parent.mkdir(parents=True, exist_ok=True)
        RESULT.write_text(json.dumps(self.result, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        CAPTURES.mkdir(parents=True, exist_ok=True)
        (CAPTURES / f"raw-{int(time.time())}.json").write_text(json.dumps(self.raw, ensure_ascii=False, indent=1), encoding="utf-8")
        try:
            self.x.app.DisplayAlerts = False
        finally:
            self.x.close()


def summarize(v: dict) -> dict:
    out = {}
    for k, pops in v.items():
        if k == "subs":
            out[k] = {n: {"hover": s["hover"], "menu": menu_items(s.get("menu", []))} for n, s in pops.items()}
        else:
            out[k] = menu_items(pops)
    return out


def menu_items(pops: list) -> list[dict]:
    out = []
    for p in pops:
        w = p["window"]
        out.append({"window": {k: w[k] for k in ("type", "class", "rect")},
                    "items": [{k: it[k] for k in ("type", "name", "enabled", "rect")} for it in p["items"][1:]
                              if it["name"].strip() and it["type"] not in ("Pane",) and not (it["type"] in ("Text", "Edit") and len(it["name"]) > 60)]})
    return out


def controls(pops: list) -> list[dict]:
    return menu_items(pops)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    passes = sys.argv[1:] or ["ribbon", "buttons", "context", "dialog", "ranges"]
    p = Probe()
    try:
        if "ribbon" in passes:
            # The ribbon is read without input, off screen.
            winapi.place_quietly(p.x.hwnd, -4000, 0, WIN[2], WIN[3])
            p.x.app.Visible = True
            time.sleep(1.0)
            p.pass_ribbon()
        real = [n for n in passes if n != "ribbon"]
        if real:
            with real_input("charts-buttons: Excel chart buttons probe", wait=1800):
                p.x.app.Visible = True
                winapi.place(p.x.hwnd, *WIN)
                time.sleep(0.8)
                p.front()
                for n in real:
                    print("pass", n, flush=True)
                    try:
                        getattr(p, f"pass_{n}")()
                    except Exception as ex:  # noqa: BLE001
                        print("pass failed", n, ex, flush=True)
                        p.result.setdefault("errors", {})[n] = str(ex)[:200]
                        p.esc(3)
                p.x.app.Visible = False
    finally:
        p.close()
    print("written", RESULT.name)


if __name__ == "__main__":
    main()
