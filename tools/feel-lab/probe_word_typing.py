"""Probe: Word's AutoCorrect and AutoFormat while typing, with real keyboard input in an own
Word instance (takes over the keyboard). Each scenario starts in a new empty document; the
result lists the paragraphs (text with special characters as code points, style, indents,
list label) and the list definitions Word made (first three levels).

    .venv/Scripts/python.exe probe_word_typing.py [name-filter]
"""
from __future__ import annotations

import datetime
import json
import re
import sys
import time
from pathlib import Path

from feellab import winapi
from feellab.targets import WINDOW_RECT

HERE = Path(__file__).parent

PREFIXES = ["*", "-", ">", "->", "=>", "o", "1.", "1)", "(1)", "a.", "a)", "(a)", "A.", "A)", "i.", "i)", "I.", "I)",
            "1.1", "§", "+", "•", "–", "2.", "c)", "IV.", "(A)", "(i)", "(I)", "1-", "a-", "1:", "#", "=", "—", "~"]

SCENARIOS: list[tuple[str, list]] = [
    ("apostrophe typed", [("type", "geht'")]),
    ("apostrophe then letter", [("type", "geht's")]),
    ("apostrophe at start", [("type", "'s Leben")]),
    ("quote after space", [("type", "Er sagte '")]),
    ("closing single then space", [("type", "Er sagte 'ja' ")]),
    ("(c) no space", [("type", "(c)")]),
    ("(c) after word", [("type", "Text(c)")]),
    ("--> no space", [("type", "A -->")]),
    ("... no space", [("type", "Und...")]),
    ("... then letter", [("type", "Und...so")]),
    (":) no space", [("type", "Gut :)")]),
    ("ordinal date", [("type", "Am 1. januar ")]),
    ("number period mid sentence", [("type", "Es gibt 3. die ")]),
    ("a) text capital", [("type", "a) erster ")]),
    ("indent backspace", [("type", "Absatz"), ("keys", "ctrl+m"), ("keys", "ctrl+m"), ("keys", "home"), ("keys", "backspace")]),
    ("indent backspace twice", [("type", "Absatz"), ("keys", "ctrl+m"), ("keys", "ctrl+m"), ("keys", "home"), ("keys", "backspace"), ("keys", "backspace")]),
    ("first-line tab backspace", [("type", "Absatz"), ("keys", "home"), ("keys", "tab"), ("keys", "backspace")]),
    ("first-line tab twice backspace", [("type", "Absatz"), ("keys", "home"), ("keys", "tab"), ("keys", "tab"), ("keys", "backspace")]),
    ("hanging backspace", [("type", "Absatz"), ("keys", "ctrl+t"), ("keys", "home"), ("keys", "backspace")]),
    ("list item backspace x3", [("type", "* Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "home"),
                                ("keys", "backspace"), ("keys", "backspace"), ("keys", "backspace")]),
    ("list first item tab", [("type", "* Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "up"), ("keys", "home"), ("keys", "tab")]),
    ("list first item tab twice", [("type", "* Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "up"), ("keys", "home"), ("keys", "tab"), ("keys", "tab")]),
    ("list first item shift+tab", [("type", "* Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "up"), ("keys", "home"), ("keys", "shift+tab")]),
    ("numbered first item tab", [("type", "1. Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "up"), ("keys", "home"), ("keys", "tab")]),
    ("list second item tab", [("type", "1. Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "home"), ("keys", "tab")]),
    ("1. tab", [("type", "1."), ("keys", "tab"), ("type", "Eins")]),
    ("* tab", [("type", "*"), ("keys", "tab"), ("type", "Eins")]),
    ("list first item tab x3", [("type", "* Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "up"), ("keys", "home")] + [("keys", "tab")] * 3),
    ("list first item tab x4", [("type", "* Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "up"), ("keys", "home")] + [("keys", "tab")] * 4),
    ("list first item tab x5", [("type", "* Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "up"), ("keys", "home")] + [("keys", "tab")] * 5),
    ("list first item tab then shift+tab", [("type", "* Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "up"), ("keys", "home"), ("keys", "tab"), ("keys", "shift+tab")]),
    ("list item backspace x1", [("type", "* Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "home"), ("keys", "backspace")]),
    ("list item backspace x2", [("type", "* Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "home"), ("keys", "backspace"), ("keys", "backspace")]),
    ("list item backspace x4", [("type", "* Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "home")] + [("keys", "backspace")] * 4),
    ("list first item backspace x1", [("type", "Vorher"), ("keys", "enter"), ("type", "* Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "up"), ("keys", "home"), ("keys", "backspace")]),
    ("list first item backspace x2", [("type", "Vorher"), ("keys", "enter"), ("type", "* Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "up"), ("keys", "home"), ("keys", "backspace"), ("keys", "backspace")]),
    ("list first item backspace x3", [("type", "Vorher"), ("keys", "enter"), ("type", "* Eins"), ("keys", "enter"), ("type", "Zwei"), ("keys", "up"), ("keys", "home")] + [("keys", "backspace")] * 3),
    ("indent backspace then join", [("type", "Vorher"), ("keys", "enter"), ("type", "Absatz"), ("keys", "ctrl+m"), ("keys", "home"), ("keys", "backspace"), ("keys", "backspace")]),
    ("first-line tab twice backspace twice", [("type", "Absatz"), ("keys", "home"), ("keys", "tab"), ("keys", "tab"), ("keys", "backspace"), ("keys", "backspace")]),
    ("apostrophe word then space", [("type", "geht's ")]),
    ("apostrophe end then space", [("type", "Peters' ")]),
    ("md ##", [("type", "## Erster")]),
    ("md ###", [("type", "### Erster")]),
    ("md ####", [("type", "#### Erster")]),
    ("md bold", [("type", "Ein **fettes** Wort ")]),
    ("md italic star", [("type", "Ein *kursives* Wort ")]),
    ("md italic underscore", [("type", "Ein _kursives_ Wort ")]),
    ("md strike", [("type", "Ein ~~altes~~ Wort ")]),
    ("md code", [("type", "Ein `code` Wort ")]),
    ("md # mid", [("type", "Text # Erster")]),
    ("md # then enter", [("type", "# Titel"), ("keys", "enter"), ("type", "Text")]),
    ("i- prefix", [("type", "i- Erster")]),
    ("A- prefix", [("type", "A- Erster")]),
] + [(f"prefix {p}", [("type", p + " Erster")]) for p in PREFIXES]


def paras(d) -> list[dict]:
    out = []
    for i in range(1, d.Paragraphs.Count + 1):
        p = d.Paragraphs(i)
        r = p.Range
        t = r.Text.rstrip("\r")
        lf = r.ListFormat
        out.append({
            "text": "".join(ch if 32 <= ord(ch) < 127 or ch.isalpha() else f"<{ord(ch):04X}>" for ch in t),
            "style": p.Style.NameLocal, "left": round(p.LeftIndent, 1), "first": round(p.FirstLineIndent, 1),
            "list": lf.ListString if lf.ListType else "", "level": lf.ListLevelNumber if lf.ListType else 0,
            "fmt": ";".join(f"{a}:{w.Text}" for w in r.Words for a in ("Bold", "Italic", "StrikeThrough") if getattr(w.Font, a) in (True, -1)),
            "fonts": sorted({w.Font.Name for w in r.Words}),
        })
    return out


def numbering(d) -> list[str]:
    x = d.Content.WordOpenXML
    m = re.search(r"<w:numbering[^>]*>(.*?)</w:numbering>", x, re.S)
    if not m:
        return []
    out = []
    for a in re.finditer(r'<w:abstractNum [^>]*w:abstractNumId="(\d+)".*?</w:abstractNum>', m.group(0), re.S):
        levels = re.findall(r'<w:lvl w:ilvl="([0-2])".*?</w:lvl>', a.group(0), re.S)
        for lv in re.finditer(r'<w:lvl w:ilvl="([0-2])"[^>]*>(.*?)</w:lvl>', a.group(0), re.S):
            body = re.sub(r' w:tplc="[^"]*"', "", lv.group(2))
            body = re.sub(r"<w:start w:val=\"1\"/>", "", body)
            body = "".join(ch if ord(ch) < 127 else f"<{ord(ch):04X}>" for ch in body)
            out.append(f"abs{a.group(1)} lvl{lv.group(1)}: {body}")
        del levels
    return out


def main() -> None:
    import win32com.client
    only = sys.argv[1:]
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    app.Visible = True
    app.Options.CheckSpellingAsYouType = False
    app.Options.CheckGrammarAsYouType = False
    results = {}
    try:
        for name, steps in SCENARIOS:
            if only and not any(o in name for o in only):
                continue
            d = app.Documents.Add()
            win = d.Windows(1)
            win.View.Type = 3
            winapi.place(win.Hwnd, *WINDOW_RECT)
            winapi.activate(win.Hwnd)
            time.sleep(0.5)
            for kind, arg in steps:
                if kind == "type":
                    winapi.type_text(arg, gap=0.05)
                else:
                    winapi.press(arg)
                time.sleep(0.25)
            time.sleep(0.4)
            s = app.Selection
            results[name] = {"paras": paras(d), "numbering": numbering(d), "caret": s.Start - s.Paragraphs(1).Range.Start}
            print(name, json.dumps(results[name]["paras"], ensure_ascii=False))
            for n in results[name]["numbering"]:
                print("    ", n)
            d.Close(0)
    finally:
        app.Quit(0)
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-word-typing.json"
    old = json.loads(out.read_text(encoding="utf-8")) if out.exists() else {}
    old.update(results)
    out.write_text(json.dumps(old, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
