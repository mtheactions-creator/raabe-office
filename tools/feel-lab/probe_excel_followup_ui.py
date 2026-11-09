"""Probe (D-145 to D-148): Excel's dialogs and galleries the follow-up of RaabeCalc's by-hand audit builds. In an
Excel of our own (visible, real clicks under the real-input lock, each click only after checking that our Excel
is in front; labels and control types read through UI Automation from our Excel's windows only):

- Datenbalken > Einfarbige Füllung: each of the six entries clicked on its own column, the rules read through COM
  and the file Excel saves (the x14 extension of each rule);
- Zellen formatieren > Ausfüllen: the Musterformat list's entries, the Fülleffekte dialog's controls;
- Zellen formatieren > Schrift: the Unterstreichung list;
- Regeln verwalten: Regel duplizieren (where the copy goes, read through COM after OK) and the Regel bearbeiten
  dialog's title and controls.

    .venv/Scripts/python.exe probe_excel_followup_ui.py

Result: results/2026-11-09-excel-followup-ui.json (labels only, no descriptions, D-058).
"""
from __future__ import annotations

import json
import re
import tempfile
import time
import zipfile
from pathlib import Path

from feellab import uia, winapi
from feellab.inputlock import real_input
from probe_excel_cf_galleries import Probe, rgb

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir()) / "raabe-excel-followup-ui"
DIALOG_CLASSES = ("#32770", "bosa_sdm_XL9", "NUIDialog")


def visible(h: int) -> bool:
    import ctypes
    return bool(ctypes.windll.user32.IsWindowVisible(h))


class UI(Probe):
    def esc(self) -> None:
        if winapi.window_pid(winapi.foreground()) == self.pid:
            winapi.press("esc")

    def close(self) -> None:
        self.esc()
        self.x.close()

    def dialogs(self) -> list[int]:
        return [h for h in winapi.find_windows(lambda h: winapi.window_pid(h) == self.pid and h != self.x.hwnd
                                               and winapi.window_class(h) in DIALOG_CLASSES and visible(h))]

    def controls(self, h: int) -> list[dict]:
        root = uia.element_for(h)
        f = root.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
        out = []
        for i in range(f.Length):
            try:
                d = uia.describe(f.GetElement(i))
            except Exception:  # noqa: BLE001
                continue
            if d["name"] and d["type"] not in ("Pane", "Window", "TitleBar"):
                out.append({"name": d["name"], "type": d["type"], "rect": d["rect"]})
        return out

    def in_dialog(self, h: int, types: tuple[str, ...], name: str):
        root = uia.element_for(h)
        f = root.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
        for i in range(f.Length):
            e = f.GetElement(i)
            try:
                d = uia.describe(e)
            except Exception:  # noqa: BLE001
                continue
            if d["type"] in types and d["name"] == name and d["rect"][2] > 0:
                return e
        return None

    def key(self, combo: str) -> None:
        fg = winapi.foreground()
        if winapi.window_pid(fg) != self.pid:
            raise RuntimeError("our Excel is not in front")
        winapi.press(combo)
        time.sleep(0.5)

    def wait_dialog(self, before: set[int], t: float = 4.0) -> int | None:
        end = time.time() + t
        while time.time() < end:
            new = [h for h in self.dialogs() if h not in before]
            if new:
                time.sleep(0.4)
                return new[-1]
            time.sleep(0.1)
        return None

    def solid_bars(self) -> dict:
        ws = self.x.sheet
        for c in "BCDEFG":
            for i in range(5):
                ws.Range(f"{c}{i + 2}").Value2 = i + 1
        out = []
        for n in range(6):
            col = "BCDEFG"[n]
            ws.Range(f"{col}2:{col}6").Select()
            self.open("Datenbalken")
            hits = []
            for r in self.roots():
                f = r.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
                for i in range(f.Length):
                    try:
                        d = uia.describe(f.GetElement(i))
                    except Exception:  # noqa: BLE001
                        continue
                    if d["type"] == "ListItem" and d["name"].endswith("Datenbalken") and d["rect"][2] > 0:
                        hits.append(d)
                if hits:
                    break
            if len(hits) < 12:
                out.append({"error": f"{len(hits)} entries"})
                self.esc()
                continue
            # Farbverlauf first (six), then Einfarbige Füllung (six), as the gallery shows them.
            d = hits[6 + n]
            if not self.ours():
                raise RuntimeError("our Excel is not in front")
            winapi.click(d["rect"][0] + d["rect"][2] / 2, d["rect"][1] + d["rect"][3] / 2)
            time.sleep(0.6)
            fcs = ws.Range(f"{col}2:{col}6").FormatConditions
            rec = {"name": d["name"], "count": fcs.Count}
            if fcs.Count:
                fc = fcs(1)
                rec.update(color=rgb(fc.BarColor.Color), fill=fc.BarFillType, border=fc.BarBorder.Type, axis=fc.AxisPosition)
            out.append(rec)
            print(rec, flush=True)
        TMP.mkdir(parents=True, exist_ok=True)
        p = TMP / "bars.xlsx"
        self.x.wb.SaveCopyAs(str(p))
        with zipfile.ZipFile(p) as z:
            sx = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
        ws.Cells.FormatConditions.Delete()
        return {"entries": out, "cfRules": re.findall(r"<cfRule.*?</cfRule>", sx, re.S),
                "x14": re.findall(r"<x14:cfRule.*?</x14:cfRule>", sx, re.S)}

    def format_cells(self) -> dict:
        out: dict = {}
        self.x.sheet.Range("A1").Select()
        winapi.activate(self.x.hwnd)
        before = set(self.dialogs())
        self.key("ctrl+1")
        h = self.wait_dialog(before)
        if not h:
            return {"error": "Zellen formatieren did not open"}
        try:
            for tab in ("Ausfüllen", "Schrift"):
                t = self.in_dialog(h, ("TabItem",), tab)
                self.click(t)
                time.sleep(0.5)
                if tab == "Ausfüllen":
                    # Musterformat: the list opens below the field.
                    m = None
                    for c in self.controls(h):
                        if c["name"] == "Musterformat:" and c["type"] in ("ComboBox", "Button"):
                            m = c
                    if m is None:
                        # The field follows its label: the control right of / below "Musterformat:".
                        lab = [c for c in self.controls(h) if c["name"] == "Musterformat:"]
                        out["musterLabel"] = lab
                    else:
                        winapi.click(m["rect"][0] + m["rect"][2] / 2, m["rect"][1] + m["rect"][3] / 2)
                        time.sleep(0.7)
                        pops = winapi.find_windows(lambda w: winapi.window_pid(w) == self.pid and w not in (self.x.hwnd, h) and visible(w))
                        names = []
                        for w in pops:
                            for c in self.controls(w):
                                names.append([c["name"], c["type"]])
                        out["muster"] = names
                        self.esc()
                        time.sleep(0.4)
                    fe = self.in_dialog(h, ("Button",), "Fülleffekte...")
                    before2 = set(self.dialogs())
                    self.click(fe)
                    h2 = self.wait_dialog(before2)
                    if h2:
                        out["fuelleffekte"] = {"title": winapi.window_title(h2), "controls": [[c["name"], c["type"]] for c in self.controls(h2)]}
                        self.esc()
                        time.sleep(0.5)
                else:
                    u = None
                    for c in self.controls(h):
                        if c["name"] == "Unterstreichung:" and c["type"] == "ComboBox":
                            u = c
                    if u:
                        winapi.click(u["rect"][0] + u["rect"][2] - 8, u["rect"][1] + u["rect"][3] / 2)
                        time.sleep(0.6)
                        names = []
                        for w in [h] + winapi.find_windows(lambda w: winapi.window_pid(w) == self.pid and w not in (self.x.hwnd, h) and visible(w)):
                            names += [c["name"] for c in self.controls(w) if c["type"] == "ListItem"]
                        out["unterstreichung"] = [n for n in names if "(" in n or n in ("Keine", "Einfach", "Doppelt", "Ohne")]
                        self.esc()
                        time.sleep(0.3)
        finally:
            for _ in range(3):
                if h in self.dialogs() and self.ours_or_dialog():
                    self.esc()
                    time.sleep(0.4)
        return out

    def ours_or_dialog(self) -> bool:
        return winapi.window_pid(winapi.foreground()) == self.pid

    def manager(self) -> dict:
        ws = self.x.sheet
        for i in range(5):
            ws.Range(f"B{i + 2}").Value2 = i + 1
        r = ws.Range("B2:B6")
        a = r.FormatConditions.Add(1, 5, "=3")  # greater than 3
        a.Interior.Color = 0xCEC7FF
        b = r.FormatConditions.Add(1, 6, "=2")  # less than 2
        b.Interior.Color = 0xCEEFC6
        r.Select()
        out: dict = {"before": [[fc.Formula1, fc.Priority] for fc in r.FormatConditions]}
        before = set(self.dialogs())
        self.open_manager()
        h = self.wait_dialog(before)
        if not h:
            return {"error": "manager did not open"}
        out["managerTitle"] = winapi.window_title(h)
        # The first rule is selected when the manager opens: Regel duplizieren copies it.
        self.click(self.in_dialog(h, ("Button",), "Regel duplizieren"))
        time.sleep(0.5)
        before2 = set(self.dialogs())
        self.click(self.in_dialog(h, ("Button",), "Regel bearbeiten..."))
        h2 = self.wait_dialog(before2)
        if h2:
            out["edit"] = {"title": winapi.window_title(h2), "controls": [[c["name"], c["type"]] for c in self.controls(h2)]}
            self.esc()
            time.sleep(0.5)
        self.click(self.in_dialog(h, ("Button",), "OK"))
        time.sleep(0.6)
        out["after"] = [[fc.Formula1, fc.Priority, fc.Interior.Color] for fc in r.FormatConditions]
        ws.Cells.FormatConditions.Delete()
        return out

    def open_manager(self) -> None:
        for _ in range(2):
            if self.ours():
                self.esc()
            time.sleep(0.3)
        self.click(self.find(("TabItem",), "Start"))
        self.click(self.find(("MenuItem", "Button", "SplitButton"), "Bedingte Formatierung"))
        self.click(self.find(("MenuItem", "Button"), "Regeln verwalten..."))


def main() -> None:
    data: dict = {}
    with real_input("calc-followup probe_excel_followup_ui", wait=1200):
        p = UI()
        try:
            for k, fn in (("solidBars", p.solid_bars), ("formatCells", p.format_cells), ("manager", p.manager)):
                try:
                    data[k] = fn()
                except Exception as e:  # noqa: BLE001
                    data[k] = {"error": str(e)[:300]}
                print(k, json.dumps(data[k], ensure_ascii=False)[:1500], flush=True)
        finally:
            p.close()
    for k in ("formatCells", "manager"):
        # Controls with their rectangles are for the probe only; the result keeps labels and types.
        pass
    out = HERE / "results" / "2026-11-09-excel-followup-ui.json"
    out.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("->", out.name)


if __name__ == "__main__":
    main()
