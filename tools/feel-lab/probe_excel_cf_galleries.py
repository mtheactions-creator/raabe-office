"""Probe (D-143): what Excel's Bedingte Formatierung galleries do. In an Excel of our own (visible, real clicks
under the real-input lock, each click only after checking that our Excel is in front) the Farbskalen and
Datenbalken entries are clicked one after the other on B2:B6 (1 to 5) and the rule each one makes is read
through COM: the colours from the lowest to the highest value, the data bar's colour and fill. Labels as
Excel names them.

    .venv/Scripts/python.exe probe_excel_cf_galleries.py

Result: results/2026-11-09-excel-cf-galleries.json.
"""
from __future__ import annotations

import json
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.calc_catalogue import ExcelSide
from feellab.inputlock import real_input

HERE = Path(__file__).parent
POPUP = "Net UI Tool Window"


def rgb(bgr: int) -> str:
    return "#%02X%02X%02X" % (bgr & 255, (bgr >> 8) & 255, (bgr >> 16) & 255)


class Probe:
    def __init__(self):
        self.side = ExcelSide()
        self.x = self.side.x
        self.pid = self.x.pid
        self.root = self.side.root

    def ours(self) -> bool:
        return winapi.window_pid(winapi.foreground()) == self.pid

    def click(self, el) -> None:
        r = uia.describe(el)["rect"]
        if not self.ours():
            winapi.activate(self.x.hwnd)
        if not self.ours():
            raise RuntimeError("our Excel is not in front")
        winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
        time.sleep(0.6)

    def roots(self) -> list:
        # Menus open in windows of their own; galleries also show below the ribbon control in the main window.
        pops = winapi.find_windows(lambda h: winapi.window_pid(h) == self.pid and h != self.x.hwnd)
        return [uia.element_for(h) for h in pops] + [self.root]

    def find(self, kind: tuple[str, ...], name: str):
        for r in self.roots():
            try:
                f = r.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
            except Exception:  # noqa: BLE001
                continue
            for i in range(f.Length):
                e = f.GetElement(i)
                try:
                    d = uia.describe(e)
                except Exception:  # noqa: BLE001
                    continue
                if d["type"] in kind and d["name"] == name and d["rect"][2] > 0:
                    return e
        print("not found:", kind, name, flush=True)
        return None

    def items(self) -> list[str]:
        out = []
        for r in self.roots():
            f = r.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
            for i in range(f.Length):
                try:
                    d = uia.describe(f.GetElement(i))
                except Exception:  # noqa: BLE001
                    continue
                if d["type"] == "ListItem" and d["name"] and d["name"] not in out:
                    out.append(d["name"])
        return out

    def open(self, sub: str) -> None:
        # A menu of our Excel may still be in front: Esc closes it (only when the window in front is ours).
        for _ in range(3):
            if not self.ours():
                try:
                    winapi.activate(self.x.hwnd)
                except RuntimeError:
                    time.sleep(0.5)
                    continue
            if self.ours():
                winapi.press("esc")
            time.sleep(0.3)
        tab = self.find(("TabItem",), "Start")
        self.click(tab)
        cf = self.find(("MenuItem", "Button", "SplitButton"), "Bedingte Formatierung")
        self.click(cf)
        s = self.find(("MenuItem", "Button", "SplitButton", "ListItem"), sub)
        self.click(s)

    def run(self) -> dict:
        ws = self.x.sheet
        for i in range(5):
            ws.Range(f"B{i + 2}").Value2 = i + 1
        out: dict = {}
        for sub in ("Farbskalen", "Datenbalken"):
            self.open(sub)
            key = "Farbskala" if sub == "Farbskalen" else "Datenbalken"
            names = [n for n in self.items() if key in n]
            print(sub, names, flush=True)
            out[sub] = []
            winapi.press("esc")
            for n, name in enumerate(names):
                ws.Range("B2:B6").Select()
                self.open(sub)
                el = None
                # The same name twice (Farbverlauf and Einfarbige Füllung): the n-th of that name.
                same = [k for k, x in enumerate(names) if x == name]
                for r in self.roots():
                    f = r.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
                    hits = []
                    for i in range(f.Length):
                        e = f.GetElement(i)
                        try:
                            d = uia.describe(e)
                        except Exception:  # noqa: BLE001
                            continue
                        if d["type"] == "ListItem" and d["name"] == name and d["rect"][2] > 0:
                            hits.append(d["rect"])
                    if hits:
                        el = hits[min(same.index(n), len(hits) - 1)]
                        break
                if el is None:
                    out[sub].append({"name": name, "error": "not found"})
                    continue
                if not self.ours():
                    raise RuntimeError("our Excel is not in front")
                winapi.click(el[0] + el[2] / 2, el[1] + el[3] / 2)
                time.sleep(0.6)
                fcs = ws.Range("B2:B6").FormatConditions
                rec: dict = {"name": name, "count": fcs.Count}
                if fcs.Count:
                    fc = fcs(1)
                    try:
                        if sub == "Farbskalen":
                            crit = fc.ColorScaleCriteria
                            rec["colors"] = [rgb(crit(i).FormatColor.Color) for i in range(1, crit.Count + 1)]
                        else:
                            rec["color"] = rgb(fc.BarColor.Color)
                            rec["fill"] = fc.BarFillType  # 0 solid, 1 gradient
                            rec["border"] = fc.BarBorder.Type
                    except Exception as e:  # noqa: BLE001
                        rec["error"] = str(e)[:200]
                    fcs.Delete()
                out[sub].append(rec)
                print(sub, rec, flush=True)
        return out

    def close(self) -> None:
        try:
            winapi.press("esc")
        except Exception:  # noqa: BLE001
            pass
        self.x.close()


def main() -> None:
    with real_input("audit-calc probe_excel_cf_galleries", wait=3600):
        p = Probe()
        try:
            data = p.run()
        finally:
            p.close()
    out = HERE / "results" / "2026-11-09-excel-cf-galleries.json"
    out.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("->", out.name)


if __name__ == "__main__":
    main()
