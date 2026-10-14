"""Probe: what Word's Entwurf galleries do. Each entry of Designs, Farben, Schriftarten,
Formatvorlagensatz (Dokumentformatierung) and Absatzabstand is applied in our own visible
Word instance (UI Automation invoke, the gallery opened with a real click), then the theme
(colours, fonts) and the styles are read through COM. Takes over mouse and screen.
Result: results/<date>-word-design.json (names, colours, fonts, style formatting).

    .venv/Scripts/python.exe probe_word_design.py [gallery …]
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

from feellab import uia, winapi
from probe_word_dialogs import Word

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-word-design.json"

# Word's style constants (WdBuiltinStyle) of the styles a style set changes.
STYLES = {"Normal": -1, "Heading 1": -2, "Heading 2": -3, "Heading 3": -4, "Title": -63, "Subtitle": -75, "Quote": -181,
          "Intense Quote": -182, "Strong": -88, "Emphasis": -89, "Subtle Emphasis": -261, "Intense Emphasis": -262,
          "Subtle Reference": -263, "Intense Reference": -264, "Book Title": -265, "List Paragraph": -180, "No Spacing": -158}
# MsoThemeColorSchemeIndex 1..12.
COLORS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
GALLERIES = [("Designs", "Designs"), ("Farben", "Farben"), ("Schriftarten", "Schriftarten"), ("Absatzabstand", "Absatzabstand"), ("Effekte", "Effekte")]


def hexrgb(bgr: int) -> str:
    return f"{bgr & 255:02X}{(bgr >> 8) & 255:02X}{(bgr >> 16) & 255:02X}"


def theme(d) -> dict:
    t = d.DocumentTheme
    cs = t.ThemeColorScheme
    fs = t.ThemeFontScheme
    return {"colors": {COLORS[i]: hexrgb(cs.Colors(i + 1).RGB) for i in range(12)},
            "major": fs.MajorFont(1).Name, "minor": fs.MinorFont(1).Name}


def styles(d) -> dict:
    out = {}
    for name, sid in STYLES.items():
        try:
            s = d.Styles(sid)
        except Exception:  # noqa: BLE001
            continue
        f, p = s.Font, s.ParagraphFormat if s.Type == 1 else None
        e = {"font": f.Name, "size": f.Size, "bold": f.Bold == -1, "italic": f.Italic == -1, "caps": f.AllCaps == -1, "smallCaps": f.SmallCaps == -1,
             "color": hexrgb(f.Color) if f.Color not in (-16777216, 9999999) else "auto", "themeColor": f.TextColor.ObjectThemeColor, "spacing": f.Spacing,
             "underline": f.Underline}
        if p is not None:
            e.update({"before": p.SpaceBefore, "after": p.SpaceAfter, "line": p.LineSpacing, "lineRule": p.LineSpacingRule, "align": p.Alignment,
                      "keepNext": p.KeepWithNext == -1, "borderBottom": p.Borders(-3).LineStyle, "borderBottomColor": hexrgb(p.Borders(-3).Color) if p.Borders(-3).LineStyle else None,
                      "borderBottomWidth": p.Borders(-3).LineWidth if p.Borders(-3).LineStyle else None,
                      "shading": hexrgb(p.Shading.BackgroundPatternColor) if p.Shading.BackgroundPatternColor not in (-16777216, 9999999) else None,
                      "indLeft": p.LeftIndent, "indRight": p.RightIndent})
        out[name] = e
    return out


def items_of(w: Word, gallery_name: str) -> tuple[list, object]:
    """Opens a gallery and returns its entries (ListItems) and the control."""
    ctl = w.control(gallery_name)
    if ctl is None:
        return [], None
    r = uia.describe(ctl)["rect"]
    winapi.click(r[0] + r[2] / 2, r[1] + r[3] - 10 if r[3] > 60 else r[1] + r[3] / 2)
    time.sleep(1.4)
    root = ctl
    found = []
    f = root.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    for i in range(f.Length):
        el = f.GetElement(i)
        d = uia.describe(el)
        if d["type"] == "ListItem" and d["name"]:
            found.append(el)
    return found, ctl


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    only = sys.argv[1:]
    out: dict = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    w = Word()
    d = w.d
    d.Content.Text = "Titel\rUntertitel\rÜberschrift 1\rText des Absatzes.\rÜberschrift 2\rNoch ein Absatz."
    for i, sid in [(1, -63), (2, -75), (3, -2), (5, -3)]:
        d.Paragraphs(i).Style = d.Styles(sid)
    time.sleep(0.5)
    try:
        out["default"] = {"theme": theme(d), "styles": styles(d)}
        w.tab("Entwurf")
        # The Formatvorlagensatz gallery lives in the ribbon itself (Dokumentformatierung).
        targets = GALLERIES + [("Formatvorlagensatz", "Dokumentformatierung")]
        for key, ctl_name in targets:
            if only and key not in only:
                continue
            entries: list[dict] = []
            if key == "Formatvorlagensatz":
                g = uia.find(w.root, lambda x: x["type"] in ("List", "Group") and x["name"] in ("Formatvorlagensätze", "Dokumentformatierung", "Formatvorlagensatz"), max_depth=16)
                more = uia.find(w.root, lambda x: x["name"] == "Weitere" and x["type"] in ("Button", "MenuItem"), max_depth=16)
                if more is not None:
                    r = uia.describe(more)["rect"]
                    winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
                    time.sleep(1.4)
                lst = []
                f = w.root.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
                for i in range(f.Length):
                    el = f.GetElement(i)
                    dd = uia.describe(el)
                    if dd["type"] == "ListItem" and dd["name"] and dd["rect"][1] > 100:
                        lst.append(dd["name"])
                winapi.press("esc")
                time.sleep(0.6)
                names = list(dict.fromkeys(lst))
                print(key, len(names), names[:5], flush=True)
            else:
                items, ctl = items_of(w, ctl_name)
                names = [uia.describe(x)["name"] for x in items]
                winapi.press("esc")
                time.sleep(0.6)
                print(key, len(names), names[:5], flush=True)
            for n in names:
                # Reopen and invoke the entry by name.
                if key == "Formatvorlagensatz":
                    more = uia.find(w.root, lambda x: x["name"] == "Weitere" and x["type"] in ("Button", "MenuItem"), max_depth=16)
                    if more is not None:
                        r = uia.describe(more)["rect"]
                        winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
                        time.sleep(1.2)
                    el = uia.find(w.root, lambda x: x["type"] == "ListItem" and x["name"] == n, max_depth=20)
                else:
                    items, ctl = items_of(w, ctl_name)
                    el = next((x for x in items if uia.describe(x)["name"] == n), None)
                if el is None:
                    print("  missing", n, flush=True)
                    winapi.press("esc")
                    continue
                try:
                    uia.invoke(el)
                except Exception:  # noqa: BLE001
                    try:
                        uia.select(el)
                    except Exception:  # noqa: BLE001
                        print("  cannot apply", n, flush=True)
                        winapi.press("esc")
                        continue
                time.sleep(1.0)
                if uia.find(w.root, lambda x: x["type"] == "ListItem" and x["name"] == n, max_depth=20) is not None:
                    winapi.press("esc")
                    time.sleep(0.4)
                rec = {"name": n, "theme": theme(d)}
                if key in ("Formatvorlagensatz", "Absatzabstand", "Designs"):
                    rec["styles"] = styles(d)
                entries.append(rec)
                print("  ", n, rec["theme"]["major"], rec["theme"]["minor"], rec["theme"]["colors"]["accent1"], flush=True)
            out[key] = entries
            OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
            # Back to Word's defaults for the next gallery.
            d.Close(0)
            w.d = d = w.app.Documents.Add()
            d.Content.Text = "Titel\rUntertitel\rÜberschrift 1\rText des Absatzes.\rÜberschrift 2\rNoch ein Absatz."
            for i, sid in [(1, -63), (2, -75), (3, -2), (5, -3)]:
                d.Paragraphs(i).Style = d.Styles(sid)
            win = d.Windows(1)
            win.View.Type = 3
            from feellab.targets import WINDOW_RECT
            winapi.place(win.Hwnd, *WINDOW_RECT)
            winapi.activate(win.Hwnd)
            w.hwnd = win.Hwnd
            w.root = uia.element_for(win.Hwnd)
            time.sleep(1.0)
            w.tab("Entwurf")
    finally:
        OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        w.close()


if __name__ == "__main__":
    main()
