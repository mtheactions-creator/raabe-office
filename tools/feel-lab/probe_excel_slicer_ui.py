"""Records the user interface around Datenschnitt, Daten abrufen and Prognoseblatt in German Excel 365
(D-092) through UI Automation, without real input: our own Excel, made visible off screen.

- Tabellenentwurf: every control with its keytip and state (a cell of a table active).
- PivotTable-Analyse: the same (the Filtern group with Datenschnitt einfügen).
- Datenschnitt (the contextual tab while a slicer is selected through COM): controls, keytips, the
  Datenschnittformatvorlagen gallery's entries; the dialogs Datenschnitteinstellungen and
  Berichtsverbindungen (labels and control types only, closed with Abbrechen).
- Datenschnitt auswählen (Tabellenentwurf > Datenschnitt einfügen): labels.
- Daten: the controls of Daten abrufen und transformieren with keytips (the Daten abrufen menu does not
  open through UI Automation); Zuletzt verwendete Quellen and Vorhandene Verbindungen (titles,
  buttons and tabs only: their lists show files of this PC and are not recorded).
- Einfügen: the Filter group.

Invoking a tab, menu or button through UI Automation makes Excel the foreground window for a moment;
the probe gives the focus back at once. A watchdog ends our Excel (only ours) if a dialog blocks.
Labels, types and keytips only (D-058).

    .venv/Scripts/python.exe probe_excel_slicer_ui.py

Result: results/2026-11-08-excel-slicer-ui.json.
"""
from __future__ import annotations

import json
import subprocess
import sys
import threading
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
RESULT = HERE / "results" / "2026-11-08-excel-slicer-ui.json"
DATA = [["Name", "Klasse", "Fach", "Note"], ["Anna", "8a", "Mathe", 2], ["Ben", "8a", "Deutsch", 3], ["Cem", "8b", "Mathe", 1], ["Dora", "8b", "Englisch", 4]]
DIALOG_CLASSES = ("#32770", "bosa_sdm_XL9", "NUIDialog", "bosa_sdm_msword")
KINDS = ("Button", "SplitButton", "MenuItem", "CheckBox", "RadioButton", "ComboBox", "Edit", "TabItem", "Spinner", "Text", "Group", "Hyperlink")

fg0 = winapi.foreground()


def give_back() -> None:
    if fg0 and winapi.foreground() != fg0:
        try:
            winapi.activate(fg0, 0)
        except Exception:  # noqa: BLE001
            pass


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


def ribbon(root) -> list[tuple]:
    return [(e, d) for e, d in descendants(root) if 74 < d["rect"][1] < 268 and d["name"] and d["type"] != "ListItem"]


def listed(pairs, with_lists: bool = False) -> list[dict]:
    seen, out = set(), []
    for _, d in pairs:
        if not d["name"] or len(d["name"]) > 80 or (d["type"] not in KINDS and not (with_lists and d["type"] == "ListItem")):
            continue
        key = (d["type"], d["name"], d["key"])
        if key in seen:
            continue
        seen.add(key)
        out.append({k: d[k] for k in ("type", "name", "key", "enabled")})
    return out


class Watchdog:
    """Ends our own Excel when a step hangs (a dialog the probe cannot reach)."""

    def __init__(self, pid: int):
        self.pid, self.deadline, self.fired = pid, None, False
        threading.Thread(target=self.run, daemon=True).start()

    def arm(self, seconds: float) -> None:
        self.deadline = time.time() + seconds

    def disarm(self) -> None:
        self.deadline = None

    def run(self) -> None:
        while True:
            time.sleep(0.5)
            if self.deadline and time.time() > self.deadline:
                self.fired = True
                subprocess.run(["taskkill", "/PID", str(self.pid), "/F"], capture_output=True)
                return


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    out: dict = {}
    x = ExcelTarget(visible=False)
    pid = x.pid
    dog = Watchdog(pid)

    def windows(classes=DIALOG_CLASSES) -> list[int]:
        return winapi.find_windows(lambda h: winapi.window_pid(h) == pid and h != x.hwnd and winapi.window_class(h) in classes)

    def popups() -> list[tuple]:
        res = []
        for h in winapi.find_windows(lambda h: winapi.window_pid(h) == pid and h != x.hwnd):
            res.extend(descendants(uia.element_for(h)))
        return res

    def invoke(el, wait: float = 0.8) -> None:
        dog.arm(25)
        try:
            uia.invoke(el)
        finally:
            dog.disarm()
        give_back()
        time.sleep(wait)

    def expand(el, wait: float = 0.8) -> None:
        dog.arm(25)
        try:
            p = el.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId).QueryInterface(uia.U.IUIAutomationExpandCollapsePattern)
            p.Expand()
        finally:
            dog.disarm()
        give_back()
        time.sleep(wait)

    def collapse(el) -> None:
        try:
            el.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId).QueryInterface(uia.U.IUIAutomationExpandCollapsePattern).Collapse()
        except Exception:  # noqa: BLE001
            pass
        give_back()
        time.sleep(0.4)

    def read_dialog(name: str, lists: bool = True) -> dict | None:
        """The dialog our Excel opened: its title and controls; then closed with Abbrechen (or WM_CLOSE)."""
        for _ in range(20):
            hs = windows()
            if hs:
                break
            time.sleep(0.3)
        else:
            out.setdefault("missing", []).append(name)
            return None
        h = hs[0]
        time.sleep(0.6)
        el = uia.element_for(h)
        rec = {"title": winapi.window_title(h), "class": winapi.window_class(h), "controls": listed(descendants(el), with_lists=lists)}
        btn = next((e for e, d in descendants(el) if d["type"] == "Button" and d["name"] in ("Abbrechen", "Schließen")), None)
        if btn is not None:
            invoke(btn, 0.6)
        for hh in windows():
            winapi.user32.PostMessageW(hh, 0x0010, 0, 0)
        time.sleep(0.5)
        print(name, rec["title"], len(rec["controls"]), flush=True)
        return rec

    def tab(root, name: str):
        t = uia.find(root, lambda d: d["type"] == "TabItem" and d["name"] == name, 25)
        if t is None:
            return None
        uia.select(t)
        give_back()
        time.sleep(1.0)
        return t

    def control(root, name: str, types=("Button", "SplitButton", "MenuItem")):
        return next((e for e, d in ribbon(root) if d["name"] == name and d["type"] in types), None)

    try:
        wb = x.wb
        ws = wb.Worksheets(1)
        for r, row in enumerate(DATA, 1):
            for c, v in enumerate(row, 1):
                ws.Cells(r, c).Value = v
        lo = ws.ListObjects.Add(1, ws.Range("A1:D5"), None, 1)
        lo.Name = "Noten"
        ws2 = wb.Worksheets.Add(After=ws)
        ws2.Name = "Pivot"
        pc = wb.PivotCaches().Create(1, ws.Range("A1:D5"))
        pt = pc.CreatePivotTable(ws2.Range("A3"), "PivotNoten")
        pt.PivotFields("Klasse").Orientation = 1
        pt.AddDataField(pt.PivotFields("Note"), None, -4157)
        ws.Activate()
        ws.Range("A2").Select()
        winapi.place_quietly(x.hwnd, -4000, 0, 2400, 1300)
        x.app.Visible = True
        time.sleep(1.2)
        give_back()
        root = uia.element_for(x.hwnd)

        # Tabellenentwurf.
        t = tab(root, "Tabellenentwurf")
        out["tableDesign"] = {"tab": uia.describe(t)["key"] if t else None, "controls": listed(ribbon(root))}
        b = control(root, "Datenschnitt einfügen...") or control(root, "Datenschnitt einfügen")
        if b is not None:
            invoke(b)
            out["insertDialog"] = read_dialog("Datenschnitt auswählen")

        # PivotTable-Analyse.
        ws2.Activate()
        ws2.Range("A4").Select()
        time.sleep(0.8)
        controls: list = []
        t = None
        for _ in range(3):
            t = tab(root, "PivotTable-Analyse")
            time.sleep(0.8)
            controls = listed(ribbon(root))
            if len(controls) > 3:
                break
        out["pivotAnalyze"] = {"tab": uia.describe(t)["key"] if t else None, "controls": controls}

        # Einfügen > Filter.
        ws.Activate()
        ws.Range("F2").Select()
        time.sleep(0.6)
        tab(root, "Einfügen")
        out["insert"] = [c for c in listed(ribbon(root)) if c["name"] in ("Filter", "Datenschnitt", "Zeitachse")]

        # A slicer, selected through COM: the Datenschnitt tab.
        sc = wb.SlicerCaches.Add2(lo, "Fach")
        sl = sc.Slicers.Add(ws, Top=10, Left=300, Width=144, Height=198)
        sl.Shape.Select()
        time.sleep(1.2)
        t = tab(root, "Datenschnitt")
        out["slicerTab"] = {"tab": uia.describe(t)["key"] if t else None, "controls": listed(ribbon(root))}
        # The style gallery.
        more = next((e for e, d in ribbon(root) if d["type"] == "Button" and d["name"] in ("Weitere", "Datenschnittformatvorlagen")), None)
        if more is None:
            more = next((e for e, d in ribbon(root) if "Datenschnittformat" in d["name"] and d["type"] in ("Button", "MenuItem", "SplitButton")), None)
        if more is not None:
            try:
                expand(more)
            except Exception:  # noqa: BLE001
                invoke(more)
            items = [d for _, d in popups() + descendants(root) if d["type"] == "ListItem" and "Datenschnittformat" in d["name"]]
            seen = []
            for d in items:
                if d["name"] not in seen:
                    seen.append(d["name"])
            extra = [d["name"] for _, d in popups() if d["type"] in ("MenuItem", "Button") and d["name"]]
            out["slicerStyles"] = {"items": seen, "menu": sorted(set(extra))}
            collapse(more)
            for hh in winapi.find_windows(lambda h: winapi.window_pid(h) == pid and winapi.window_class(h) == "Net UI Tool Window"):
                winapi.user32.PostMessageW(hh, 0x0010, 0, 0)
            time.sleep(0.4)
        sl.Shape.Select()
        time.sleep(0.6)
        tab(root, "Datenschnitt")
        for label in ("Datenschnitteinstellungen", "Berichtsverbindungen"):
            b = next((e for e, d in ribbon(root) if d["name"].startswith(label) and d["type"] in ("Button", "SplitButton")), None)
            if b is not None:
                invoke(b)
                out[label] = read_dialog(label)
                sl.Shape.Select()
                time.sleep(0.5)
                tab(root, "Datenschnitt")
        # Context menu entries cannot be opened without input; the tab's states with two slicers selected.
        sc2 = wb.SlicerCaches.Add2(lo, "Klasse")
        sl2 = sc2.Slicers.Add(ws, Top=10, Left=460, Width=144, Height=198)
        ws.Shapes.Range([sl.Shape.Name, sl2.Shape.Name]).Select()
        time.sleep(0.8)
        tab(root, "Datenschnitt")
        out["slicerTabTwo"] = listed(ribbon(root))

        # Daten.
        ws.Range("A2").Select()
        time.sleep(0.6)
        t = tab(root, "Daten")
        out["data"] = {"tab": uia.describe(t)["key"] if t else None, "controls": [c for c in listed(ribbon(root))]}
        # The Daten abrufen menu does not open through UI Automation (tried: Expand shows no popup), so its
        # entries are not recorded here.
        get = None
        menus: dict = {}
        if get is not None:
            expand(get)
            top = [(e, d) for e, d in popups() if d["type"] == "MenuItem" and d["name"]]
            menus["Daten abrufen"] = listed(top)
            for name in [d["name"] for _, d in top]:
                sub = next((e for e, d in popups() if d["type"] == "MenuItem" and d["name"] == name), None)
                if sub is None:
                    continue
                try:
                    p = sub.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId)
                    if not p:
                        continue
                    before = {d["name"] for _, d in popups()}
                    expand(sub, 0.7)
                    after = [(e, d) for e, d in popups() if d["type"] == "MenuItem" and d["name"] and d["name"] not in before]
                    if after:
                        menus[name] = listed(after)
                    collapse(sub)
                except Exception:  # noqa: BLE001
                    continue
            collapse(get)
            for hh in winapi.find_windows(lambda h: winapi.window_pid(h) == pid and winapi.window_class(h) == "Net UI Tool Window"):
                winapi.user32.PostMessageW(hh, 0x0010, 0, 0)
            time.sleep(0.4)
        out["dataMenus"] = menus
        for label in ("Zuletzt verwendete Quellen", "Vorhandene Verbindungen"):
            tab(root, "Daten")
            b = next((e for e, d in ribbon(root) if d["name"] == label and d["type"] in ("Button",)), None)
            if b is not None:
                invoke(b, 1.5)
                rec = read_dialog(label, lists=False)
                if rec:
                    rec["controls"] = [c for c in rec["controls"] if c["type"] in ("Button", "TabItem", "CheckBox", "RadioButton", "ComboBox")]
                out[label] = rec
    finally:
        give_back()
        try:
            x.app.Visible = False
            x.close()
        except Exception:  # noqa: BLE001
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
        give_back()
    out["watchdogFired"] = dog.fired
    RESULT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("written", RESULT.name)


if __name__ == "__main__":
    main()
