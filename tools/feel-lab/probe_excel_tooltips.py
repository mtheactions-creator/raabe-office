"""Probe: Excel's tooltips for every ribbon control (D-061), as probe_word_tooltips.py does for
Word: our own visible Excel, every tab opened, each control hovered and the tooltip window read
through UI Automation. Takes over mouse and screen. Result: results/<date>-excel-tooltips.json
({tab: {control: {title, described}}}); the full texts (Microsoft's descriptions, which stay out
of the repository) go to captures/<date>-excel-tooltips-full.json. RaabeCalc's tooltips take
Excel's titles and shortcuts and describe the commands in their own words.

    .venv/Scripts/python.exe probe_excel_tooltips.py [tab …]
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from feellab.calc_catalogue import Case, ExcelSide
from probe_word_picture import popups
from probe_word_tooltips import KINDS, tooltip

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-excel-tooltips.json"
FULL = HERE / "captures" / f"{datetime.date.today()}-excel-tooltips-full.json"
TABS = ["Start", "Einfügen", "Seitenlayout", "Formeln", "Daten", "Überprüfen", "Ansicht"]


def controls(side: ExcelSide) -> list[dict]:
    lower = uia.find(side.root, lambda x: x["name"] == "Unteres Menüband", max_depth=12) or side.root
    f = lower.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    out, seen = [], set()
    for i in range(f.Length):
        d = uia.describe(f.GetElement(i))
        x, y, cw, ch = d["rect"]
        if d["type"] not in KINDS or not d["name"] or cw <= 0 or ch <= 0:
            continue
        if d["name"] in ("Weitere Optionen", "Öffnen", "Mehr", "Weniger"):
            continue
        key = (d["name"], x // 8, y // 8)
        if key in seen:
            continue
        seen.add(key)
        out.append(d)
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    only = sys.argv[1:]
    out: dict = json.loads(FULL.read_text(encoding="utf-8")) if FULL.exists() else {}
    side = ExcelSide()
    try:
        side.reset(Case("tips", {"A1": "1"}, [], "x"))
        winapi.activate(side.x.hwnd)
        for tab in TABS:
            if only and tab not in only:
                continue
            el = uia.find(side.root, lambda d, t=tab: d["type"] == "TabItem" and d["name"] == t)
            uia.select(el)
            time.sleep(0.8)
            res: dict = {}
            for d in controls(side):
                x, y, cw, ch = d["rect"]
                known = set(popups())
                winapi.move(x + cw / 2, y + ch / 2 + 3)
                time.sleep(0.15)
                winapi.move(x + cw / 2, y + ch / 2)
                text = ""
                for _ in range(12):
                    time.sleep(0.2)
                    text = tooltip(known)
                    if text:
                        break
                res[d["name"]] = text
                print(tab, "|", d["name"], "|", text[:80], flush=True)
                winapi.move(x + cw / 2, y - 60)
                time.sleep(0.25)
            out[tab] = res
    finally:
        FULL.parent.mkdir(exist_ok=True)
        FULL.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        titles: dict = {}
        for t, c in out.items():
            titles[t] = {}
            for k, v in c.items():
                parts = [x for x in v.split(" | ") if x and x != "Weitere Infos"] or [""]
                titles[t][k] = {"title": parts[0], "described": len(parts) > 1}
        OUT.write_text(json.dumps(titles, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        side.close()


if __name__ == "__main__":
    main()
