"""Probe: Word's tooltips (supertips) for every ribbon control (Step 4c F). In our own visible
Word instance every tab (the contextual ones with a picture, a shape, a table and an equation
selected) is opened, each control is hovered and the tooltip window that appears is read
through UI Automation (Word does not give the texts to UI Automation otherwise). Takes over
mouse and screen. Result: results/<date>-word-tooltips.json ({tab: {control: {title, described}}});
the full texts (Microsoft's descriptions, which stay out of the repository) go to
captures/<date>-word-tooltips-full.json. RaabeWrite's tooltips (apps/raabe-write/src/tooltips.ts)
take Word's titles and shortcuts and describe the commands in their own words.

    .venv/Scripts/python.exe probe_word_tooltips.py [tab …]
"""
from __future__ import annotations

import datetime
import json
import sys
import tempfile
import time
from pathlib import Path

from feellab import uia, winapi
from probe_word_context import make_png
from probe_word_dialogs import Word
from probe_word_picture import popups

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-word-tooltips.json"
FULL = HERE / "captures" / f"{datetime.date.today()}-word-tooltips-full.json"
PNG = Path(tempfile.gettempdir()) / "raabe-tooltip-picture.png"
TABS = ["Start", "Einfügen", "Entwurf", "Layout", "Referenzen", "Überprüfen", "Ansicht"]
KINDS = ("Button", "SplitButton", "CheckBox", "ComboBox", "MenuItem")


def tooltip(known: set[int]) -> str:
    for h in popups():
        if h in known:
            continue
        el = uia.element_for(h)
        names = [uia.describe(el)["name"]]
        try:
            f = el.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
            names += [uia.describe(f.GetElement(i))["name"] for i in range(f.Length)]
        except Exception:  # noqa: BLE001
            pass
        seen: list[str] = []
        for n in names:
            if n and n not in seen:
                seen.append(n)
        if seen:
            return " | ".join(seen)
    return ""


def controls(w: Word) -> list[dict]:
    lower = uia.find(w.root, lambda x: x["type"] == "Pane" and x["name"].startswith("Unter"), max_depth=12) or w.root
    f = lower.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    out, seen = [], set()
    l, t, r, b = winapi.window_rect(w.hwnd)
    for i in range(f.Length):
        d = uia.describe(f.GetElement(i))
        x, y, cw, ch = d["rect"]
        if d["type"] not in KINDS or not d["name"] or cw <= 0 or ch <= 0 or y > t + 330 or y < t + 100:
            continue
        # The arrow part of split buttons ("Weitere Optionen") has the button's tooltip.
        if d["name"] in ("Weitere Optionen", "Öffnen", "Mehr", "Weniger"):
            continue
        key = (d["name"], x // 8, y // 8)
        if key in seen:
            continue
        seen.add(key)
        out.append(d)
    return out


def record(w: Word, tab: str, out: dict) -> None:
    w.tab(tab)
    time.sleep(0.8)
    res: dict = {}
    for d in controls(w):
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
        print(tab, "|", d["name"], "|", text[:120], flush=True)
        winapi.move(x + cw / 2, y - 60)
        time.sleep(0.25)
    out[tab] = res


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    only = sys.argv[1:]
    out: dict = json.loads(FULL.read_text(encoding="utf-8")) if FULL.exists() else {}
    make_png(PNG)
    w = Word()
    d = w.d
    try:
        for tab in TABS:
            if not only or tab in only:
                record(w, tab, out)
        # Contextual tabs.
        d.Range(d.Content.End - 1, d.Content.End - 1).InsertParagraphAfter()
        pic = d.InlineShapes.AddPicture(str(PNG), False, True, d.Range(d.Content.End - 1, d.Content.End - 1))
        pic.Select()
        if not only or "Bildformat" in only:
            record(w, "Bildformat", out)
        shp = d.Shapes.AddShape(1, 250, 150, 120, 60, d.Paragraphs(1).Range)
        shp.Select()
        if not only or "Formformat" in only:
            record(w, "Formformat", out)
        r = d.Range(d.Content.End - 1, d.Content.End - 1)
        r.InsertParagraphAfter()
        t = d.Tables.Add(d.Range(d.Content.End - 1, d.Content.End - 1), 2, 2)
        t.Cell(1, 1).Range.Select()
        if not only or "Tabellenentwurf" in only:
            record(w, "Tabellenentwurf", out)
        if not only or "Tabellenlayout" in only:
            record(w, "Tabellenlayout", out)
        r = d.Range(d.Content.End - 1, d.Content.End - 1)
        r.InsertParagraphAfter()
        r = d.Range(d.Content.End - 1, d.Content.End - 1)
        r.Text = "a/b"
        m = d.OMaths.Add(r)
        m.OMaths(1).BuildUp()
        d.OMaths(1).Range.Select()
        if not only or "Formel" in only:
            record(w, "Formel", out)
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
        w.close()


if __name__ == "__main__":
    main()
