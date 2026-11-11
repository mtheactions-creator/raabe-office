"""Probe: how Excel draws the dropdowns of its Format panes (Step 11, D-161). In an Excel of our own (DispatchEx,
window off-screen, no real input) a line chart's series, a shape and a picture are selected through COM and their
Format pane is opened (CommandBars.ExecuteMso). Each dropdown of the pane is opened through UI Automation
(ExpandCollapse) and its open list read: the entries (names, kind, rectangles), so the layout (list or grid, columns,
entry size) and whether an entry carries a text of its own next to its picture. A picture of the open list (PrintWindow
of our own Excel's popup window) goes to captures/ for a look by eye; it is not committed (D-058).

Recorded per dropdown: "closed" (the box: size, its name), "kind" (gallery: NetUI gallery items; menu: menu items;
list: a combo box list), "count", "columns", "rows", "entry" ([w, h] in px at the screen's scaling), "texts" (whether
the open list draws text: entries wider than their picture area with a name), "names" (the labels; Microsoft's
labels only, never descriptions).

    .venv/Scripts/python.exe probe_excel_picturedrops.py

Result: results/2026-11-11-excel-picturedrops.json
"""
from __future__ import annotations

import ctypes
import json
import time
from ctypes import wintypes
from pathlib import Path

import probe_excel_chartpane as P
from feellab import uia, winapi

HERE = Path(__file__).parent
OUT = HERE / "results" / "2026-11-11-excel-picturedrops.json"
CAPS = HERE / "captures" / "picturedrops"
U = uia.U
UIA = uia._uia
user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32


def print_window(hwnd: int, path: Path) -> bool:
    """A picture of one of our own Excel's windows through PrintWindow (works off-screen; no screen capture)."""
    from PIL import Image
    l, t, r, b = winapi.window_rect(hwnd)
    w, h = r - l, b - t
    if w <= 0 or h <= 0 or w > 4000 or h > 4000:
        return False
    hdc = user32.GetWindowDC(hwnd)
    mem = gdi32.CreateCompatibleDC(hdc)
    bmp = gdi32.CreateCompatibleBitmap(hdc, w, h)
    gdi32.SelectObject(mem, bmp)
    ok = user32.PrintWindow(hwnd, mem, 2)
    info = winapi.BITMAPINFOHEADER(biSize=ctypes.sizeof(winapi.BITMAPINFOHEADER), biWidth=w, biHeight=-h, biPlanes=1, biBitCount=32, biCompression=0)
    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(mem, bmp, 0, h, buf, ctypes.byref(info), 0)
    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(mem)
    user32.ReleaseDC(hwnd, hdc)
    if not ok:
        return False
    path.parent.mkdir(parents=True, exist_ok=True)
    Image.frombuffer("RGBA", (w, h), buf.raw, "raw", "BGRA", 0, 1).convert("RGB").save(path)
    return True


ENTRY = ("ListItem", "MenuItem", "Button", "RadioButton", "DataItem", "SplitButton")


def entries(root) -> list[dict]:
    out = []
    try:
        f = root.FindAll(U.TreeScope_Descendants, UIA.CreateTrueCondition())
    except Exception:  # noqa: BLE001
        return out
    for i in range(f.Length):
        e = f.GetElement(i)
        d = P.d_of(e)
        if not d or d["type"] not in ENTRY or not d["name"].strip() or d["name"] in P.NOISE:
            continue
        w, h = d["rect"][2], d["rect"][3]
        if w <= 0 or h <= 0:
            continue
        kids = [P.d_of(k) for k in P.kids(e)]
        out.append({"name": d["name"], "type": d["type"], "class": d["class"], "rect": d["rect"],
                    "parts": sorted({k["type"] for k in kids if k})})
    return out


def layout(es: list[dict]) -> dict:
    if not es:
        return {}
    xs = sorted({e["rect"][0] for e in es})
    ys = sorted({e["rect"][1] for e in es})
    ws = sorted(e["rect"][2] for e in es)
    hs = sorted(e["rect"][3] for e in es)
    return {"count": len(es), "columns": len(xs), "rows": len(ys), "entry": [ws[len(ws) // 2], hs[len(hs) // 2]],
            "types": sorted({e["type"] for e in es}), "classes": sorted({e["class"] for e in es}),
            "parts": sorted({p for e in es for p in e["parts"]})}


def open_list(pane: P.Pane, el, tag: str) -> dict:
    """Opens a dropdown through ExpandCollapse, reads the open list and takes its picture; closes it again."""
    x = pane.x
    p = P.expander(el)
    if p is None:
        return {"error": "no ExpandCollapse"}
    before = set(winapi.find_windows(lambda h: winapi.window_pid(h) == x.pid))
    res: dict = {}
    for attempt in range(3):
        try:
            p.Expand()
        except Exception as e:  # noqa: BLE001
            return {"error": str(e)[:80]}
        time.sleep(0.9)
        # NetUI shows an open gallery in the main window's tree (NetUIToolWindow) and in a window of its own.
        pops = [h for h in winapi.find_windows(lambda h: winapi.window_pid(h) == x.pid)
                if user32.IsWindowVisible(h) and winapi.window_class(h) == "Net UI Tool Window"]
        es: list[dict] = []
        cond = UIA.CreatePropertyCondition(U.UIA_ClassNamePropertyId, "NetUIToolWindow")
        try:
            f = pane.root.FindAll(U.TreeScope_Descendants, cond)
            for i in range(f.Length):
                es += entries(f.GetElement(i))
        except Exception:  # noqa: BLE001
            pass
        if not es:
            for h in pops:
                es += entries(uia.element_for(h))
        if not es:
            es = [e for e in entries(el)]
        seen = set()
        es = [e for e in es if not (tuple(e["rect"]) + (e["name"],) in seen or seen.add(tuple(e["rect"]) + (e["name"],)))]
        if es:
            res = layout(es)
            res["names"] = [e["name"] for e in es]
            # The window that holds the entries (not the pane's own tool window).
            ex, ey = es[0]["rect"][0] + 2, es[0]["rect"][1] + 2
            holder = [hw for hw in pops if (lambda r: r[0] <= ex <= r[2] and r[1] <= ey <= r[3])(winapi.window_rect(hw))]
            holder.sort(key=lambda hw: (lambda r: (r[2] - r[0]) * (r[3] - r[1]))(winapi.window_rect(hw)))
            if holder:
                l, t, r, b = winapi.window_rect(holder[0])
                res["popup"] = [r - l, b - t]
                print_window(holder[0], CAPS / f"{tag}.png")
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
    return res


def dropdowns(pane: P.Pane, title: str, want: set[str] | None = None) -> dict:
    """Every dropdown of an open section: menus (picture galleries in Excel) and combo boxes (text lists)."""
    out = {}
    P.expand_section(pane, title)
    for i in P.raw_section(pane, title):
        t, d = i["t"], i["_d"]
        if t not in ("MenuItem", "ComboBox", "SplitButton", "Button"):
            continue
        if P.expander(i["_el"]) is None and t != "ComboBox":
            continue
        name = i["label"]
        if any(name.startswith(c) for c in ("Füllfarbe", "Konturfarbe", "Schattenfarbe", "Leuchtfarbe", "Farbe", "Vordergrund", "Hintergrund", "Tiefe", "Kontur")):
            continue
        if want and not any(name.startswith(w) for w in want):
            continue
        if not d["enabled"]:
            out[name] = {"disabled": True, "closed": d["rect"][2:]}
            continue
        tag = f"{title}-{name}".replace(" ", "_").replace("/", "_")[:60]
        el = i["_el"]
        r = open_list(pane, el, tag)
        if not r and t == "ComboBox":
            btn = next((k for k in P.kids(el) if (P.d_of(k) or {}).get("type") == "Button"), None)
            if btn is not None:
                r = open_list(pane, btn, tag)
        r["control"] = t
        r["closed"] = d["rect"][2:]
        out[name] = r
        print(f"  {title} / {name}: {r.get('count')} entries, {r.get('columns')} x {r.get('rows')}, entry {r.get('entry')}", flush=True)
    return out


def tab(pane: P.Pane, name: str) -> None:
    e = next((e for n, e in pane.tabs() if n == name), None)
    if e is not None:
        P.select_item(e)
        time.sleep(0.7)


def radio(pane: P.Pane, section: str, label: str) -> None:
    P.expand_section(pane, section)
    P.select_radio(pane, section, "", label) or P.select_radio(pane, section, "Linie", label)
    time.sleep(0.8)


def record_pane(pane: P.Pane, line: bool) -> dict:
    out: dict = {}
    tab(pane, "Füllung und Linie")
    rahmen = "Linie" if line or pane.section("Linie") is not None else "Rahmen"
    radio(pane, rahmen, "Einfarbige Linie")
    out[rahmen] = dropdowns(pane, rahmen)
    if not line:
        radio(pane, "Füllung", "Farbverlauf")
        out["Füllung: Farbverlauf"] = dropdowns(pane, "Füllung")
        radio(pane, "Füllung", "Bild- oder Texturfüllung")
        out["Füllung: Bild- oder Texturfüllung"] = dropdowns(pane, "Füllung")
        radio(pane, "Füllung", "Einfarbige Füllung")
    tab(pane, "Effekte")
    for s in ("Schatten", "Spiegelung", "Leuchten", "Weiche Kanten", "3D-Format", "3D-Drehung"):
        if pane.section(s) is not None:
            out[s] = dropdowns(pane, s)
    return out


def markers(pane: P.Pane) -> dict:
    tab(pane, "Füllung und Linie")
    b = pane.body()
    f = b.FindAll(U.TreeScope_Descendants, UIA.CreateTrueCondition())
    for i in range(f.Length):
        e = f.GetElement(i)
        d = P.d_of(e)
        if d and d["name"] == "Markierung" and d["type"] in ("TabItem", "Button", "RadioButton", "ListItem"):
            P.select_item(e)
            time.sleep(0.8)
            break
    P.expand_section(pane, "Markierungsoptionen")
    P.select_radio(pane, "Markierungsoptionen", "Markertyp", "Integriert")
    time.sleep(0.8)
    return dropdowns(pane, "Markierungsoptionen")


def open_pane(pane: P.Pane, select, word: str, mso: tuple[str, ...]) -> bool:
    select()
    time.sleep(0.8)
    for m in mso:
        try:
            pane.x.app.CommandBars.ExecuteMso(m)
        except Exception:  # noqa: BLE001
            continue
        for _ in range(16):
            time.sleep(0.25)
            if word in pane.title().lower():
                return True
    return False


def main() -> None:
    data: dict = {"meta": {
        "recorded": time.strftime("%Y-%m-%d") + ", Excel 365 (German), own instance (DispatchEx) off-screen, no real input",
        "how": "Format pane opened through CommandBars.ExecuteMso; each dropdown opened through ExpandCollapse; entries "
               "read through UI Automation; pictures of the open lists by PrintWindow into captures/ (not committed)",
        "fields": "control: UIA type of the closed box; closed: its [w, h]; count/columns/rows: the open list; entry: "
                  "median entry [w, h] (px, 150 % scaling); types/classes/parts: UIA types and classes of the entries "
                  "and of their children (a Text part means the entry draws a label); names: entry labels"},
        "panes": {}}
    x = P.start_excel()
    pane = P.Pane(x)
    try:
        ws = x.wb.Worksheets(1)
        P.fill_data(ws)
        sh, ch = P.make_chart(ws, "line")
        sh.Select()
        time.sleep(0.4)
        if open_pane(pane, lambda: ch.SeriesCollection(1).Select(), "datenreihe", ("ChartFormatSelection",)):
            print("chart series pane", flush=True)
            data["panes"]["Diagramm: Datenreihe (Linie)"] = record_pane(pane, True)
            data["panes"]["Diagramm: Datenreihe (Linie)"]["Markierung"] = markers(pane)
        if open_pane(pane, lambda: (sh.Select(), ch.ChartArea.Select()), "diagrammbereich", ("ChartFormatSelection",)):
            print("chart area pane", flush=True)
            data["panes"]["Diagramm: Diagrammbereich"] = record_pane(pane, False)
        sh.Delete()
        s = ws.Shapes.AddShape(1, 40, 40, 160, 90)
        if open_pane(pane, lambda: s.Select(), "form", ("ObjectFormatDialog", "ShapeFormatDialog", "FormatObjectDialog", "ChartFormatSelection")):
            print("shape pane", flush=True)
            data["panes"]["Form formatieren"] = record_pane(pane, False)
        else:
            data["panes"]["Form formatieren"] = {"error": "pane did not open"}
        s.Delete()
    finally:
        text = json.dumps(data, ensure_ascii=False, indent=1)
        from feellab.sanitize import private_names
        low = text.lower()
        if any(n in low for n in private_names()) or "c:\\\\users" in low:
            print("result holds a private name or path; not saved")
        else:
            OUT.write_text(text + "\n", encoding="utf-8", newline="\n")
        x.close()


if __name__ == "__main__":
    main()
