"""Probe (Step 11 part 5, D-173): where Office's dialogs put their parts. Excel's Zellen formatieren and Word's Absatz
and Schriftart, each opened in our own instance (window off screen) through Application.Dialogs(...).Show() on a
worker thread (the call waits until the dialog closes); the dialog is moved off screen at once, read through UI
Automation (type, name, box in CSS px relative to the dialog's client area, plus the title bar's height), pictured
with PrintWindow into TEMP (never committed) and closed with its Abbrechen button. A modal dialog can take the
keyboard focus, so the run holds the real-input lock. Microsoft's labels only.

    .venv/Scripts/python.exe probe_office_dialog_geometry.py [excel] [word]

Result: results/2026-11-11-office-dialog-geometry.json.
"""
from __future__ import annotations

import ctypes
import json
import sys
import tempfile
import threading
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.inputlock import real_input
from probe_excel_chartpane import d_of, kids

HERE = Path(__file__).parent
OUT = HERE / "results" / "2026-11-11-office-dialog-geometry.json"
CROPS = Path(tempfile.gettempdir()) / "raabe-dialog-geo"
OFF = (-4200, 0, 1800, 2000)
KEEP = ("Button", "CheckBox", "RadioButton", "ComboBox", "Edit", "Spinner", "TabItem", "Tab", "List", "ListItem", "Text",
        "Group", "Pane", "Document", "Image")
# Application.Dialogs constants: xlDialogFormatNumber (Zellen formatieren, Zahlen), wdDialogFormatParagraph, wdDialogFormatFont.
DIALOGS = {"excel": [("Zellen formatieren", 42)], "word": [("Absatz", 175), ("Schriftart", 174)]}


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


def client_box(hwnd: int) -> tuple[int, int, int, int]:
    r = RECT()
    ctypes.windll.user32.GetClientRect(hwnd, ctypes.byref(r))
    x, y = winapi.client_origin(hwnd)
    return x, y, r.right, r.bottom


def show_on_thread(stream, number: int, done: list) -> None:
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        app = win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(stream, pythoncom.IID_IDispatch))
        done.append(app.Dialogs(number).Show())
    except Exception as e:  # noqa: BLE001
        done.append(f"error: {e}")
    finally:
        pythoncom.CoUninitialize()


def read_dialog(pid: int, title: str, app_com, number: int) -> dict:
    import pythoncom
    stream = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, app_com._oleobj_)
    done: list = []
    th = threading.Thread(target=show_on_thread, args=(stream, number, done), daemon=True)
    th.start()
    hwnd = 0
    for _ in range(80):
        time.sleep(0.1)
        hs = winapi.find_windows(lambda h: winapi.window_pid(h) == pid and winapi.window_title(h) == title)
        if hs:
            hwnd = hs[0]
            break
    if not hwnd:
        return {"error": f"dialog {title} not found", "done": done}
    l, t, r, b = winapi.window_rect(hwnd)
    winapi.place_quietly(hwnd, OFF[0] + 200, 200, r - l, b - t)
    time.sleep(0.8)
    scale = winapi.dpi_scale(hwnd)
    l, t, r, b = winapi.window_rect(hwnd)
    cx, cy, cw, ch = client_box(hwnd)
    css = lambda v: round(v / scale, 1)  # noqa: E731
    out = {"scale": scale, "window": [css(r - l), css(b - t)], "client": [css(cw), css(ch)], "titlebar": css(cy - t),
           "elements": []}
    root = uia.element_for(hwnd)
    stack = [(root, 0)]
    cancel = None
    while stack:
        e, depth = stack.pop()
        d = d_of(e)
        if not d:
            continue
        x, y, w, h = d["rect"]
        if depth and w > 0 and h > 0 and d["type"] in KEEP:
            out["elements"].append({"depth": depth, "type": d["type"], "name": d["name"], "class": d["class"],
                                    "rect": [css(x - cx), css(y - cy), css(w), css(h)]})
            if d["type"] == "Button" and d["name"] == "Abbrechen":
                cancel = e
        if depth < 12:
            stack.extend((c, depth + 1) for c in reversed(kids(e)))
    try:
        from probe_excel_picturedrops import print_window
        print_window(hwnd, CROPS / f"{title}.png")
    except Exception:  # noqa: BLE001
        pass
    if cancel is not None:
        uia.invoke(cancel)
    else:
        winapi.user32.PostMessageW(hwnd, 0x0010, 0, 0)
    th.join(10)
    return out


def record(app_name: str) -> dict:
    res = {}
    if app_name == "excel":
        from feellab.excel import ExcelTarget
        x = ExcelTarget(visible=False, place=False)
        try:
            winapi.place_quietly(x.hwnd, *OFF)
            x.app.Visible = True
            winapi.place_quietly(x.hwnd, *OFF)
            x.sheet.Range("B2").Value = 1234.5
            x.sheet.Range("B2").Select()
            for title, n in DIALOGS["excel"]:
                res[title] = read_dialog(x.pid, title, x.app, n)
        finally:
            x.close()
        return res
    from feellab.targets import WordTarget
    w = WordTarget("short", visible=False)
    try:
        w.app.Visible = True
        winapi.place_quietly(w.hwnd, *OFF)
        pid = winapi.window_pid(w.hwnd)
        for title, n in DIALOGS["word"]:
            res[title] = read_dialog(pid, title, w.app, n)
    finally:
        try:
            w.doc.Close(0)
        finally:
            w.app.Quit(0)
    return res


def main() -> None:
    apps = [a for a in sys.argv[1:] if a in DIALOGS] or list(DIALOGS)
    data = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    CROPS.mkdir(parents=True, exist_ok=True)
    for a in apps:
        with real_input(f"dialog geometry ({a})"):
            for title, rec in record(a).items():
                data[title] = rec
                print(title, rec.get("error") or f"{len(rec['elements'])} elements, client {rec['client']}, title bar {rec['titlebar']}")
    OUT.write_text(json.dumps(data, indent=1, ensure_ascii=False) + "\n", encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
