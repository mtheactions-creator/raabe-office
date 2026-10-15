"""F3b formatting behaviour: what Word's formatting commands do at the caret and on selections
(word under the caret, toggles, mixed selections, paragraph marks, Enter, typing after a
command), with Word as the reference.

No real input: Word runs invisibly and executes the very commands its shortcuts are bound to
(WordBasic, names from Application.FindKey); RaabeWrite gets the shortcut as key events
through the DevTools protocol. Commands whose keyboard behaviour differs from the WordBasic
call (Strg+E toggling back to left) belong in behaviour.py, which uses real input.

Positions are Find(needle, delta) as in behaviour.py."""

from __future__ import annotations

from .behaviour import Find
from .targets import WebViewTarget, WordTarget

# Word command → RaabeWrite shortcut (key, ctrl, shift, alt, virtual key).
CMD: dict[str, tuple] = {
    "Bold": ("F", True, True, False, None),
    "Italic": ("K", True, True, False, None),
    "Underline": ("U", True, True, False, None),
    "DoubleUnderline": ("D", True, True, False, None),
    "WordUnderline": ("W", True, True, False, None),
    "Hidden": ("H", True, True, False, None),
    "AllCaps": ("G", True, True, False, None),
    "SmallCaps": ("Q", True, True, False, None),
    "Subscript": ("#", True, False, False, 191),
    "GrowFont": ("<", True, True, False, 226),
    "ShrinkFont": ("<", True, False, False, 226),
    "GrowFontOnePoint": ("9", True, False, False, None),
    "ShrinkFontOnePoint": ("8", True, False, False, None),
    "ResetChar": (" ", True, False, False, 32),
    "ChangeCase": ("F3", False, True, False, 114),
    "NormalStyle": ("N", True, True, False, None),
    "ResetPara": ("Q", True, False, False, None),
    "CenterPara": ("E", True, False, False, None),
    "RightPara": ("R", True, False, False, None),
    "JustifyPara": ("B", True, False, False, None),
    "Indent": ("M", True, False, False, None),
    "UnIndent": ("M", True, True, False, None),
    "HangingIndent": ("T", True, False, False, None),
    "UnHang": ("T", True, True, False, None),
    "SpacePara1": ("1", True, False, False, None),
    "SpacePara2": ("2", True, False, False, None),
    "SpacePara15": ("5", True, False, False, None),
    "Heading1": ("1", True, False, True, None),
}

L = Find("Lichtenergie")
CASES: list[tuple[str, list]] = [
    ("bold, caret inside a word", [("sel", Find("Lichtenergie", 4)), ("cmd", "Bold")]),
    ("bold at a word end, then type", [("sel", Find("Lichtenergie", 12)), ("cmd", "Bold"), ("type", "X")]),
    ("bold on a selection, twice", [("sel", Find("chemische"), Find("chemische", 9)), ("cmd", "Bold"), ("cmd", "Bold")]),
    ("bold on a partly bold selection", [("sel", Find("Lichtenergie", 4)), ("cmd", "Bold"),
                                         ("sel", Find("aus Licht"), Find("chemische", 4)), ("cmd", "Bold")]),
    ("italic and underline inside a word", [("sel", Find("Bakterien", 3)), ("cmd", "Italic"), ("cmd", "Underline")]),
    ("double underline on a selection", [("sel", Find("Algen"), Find("Algen", 5)), ("cmd", "DoubleUnderline")]),
    ("word underline on a selection", [("sel", Find("einige Bakterien"), Find("einige Bakterien", 16)), ("cmd", "WordUnderline")]),
    ("hidden, caret inside a word", [("sel", Find("einige", 2)), ("cmd", "Hidden")]),
    ("grow font twice on a selection", [("sel", Find("Pflanzen"), Find("Pflanzen", 8)), ("cmd", "GrowFont"), ("cmd", "GrowFont")]),
    ("shrink font, caret inside a word", [("sel", Find("Photosynthese.", 3)), ("cmd", "ShrinkFont")]),
    ("grow one point, caret inside a word", [("sel", Find("Algen", 2)), ("cmd", "GrowFontOnePoint")]),
    ("shrink one point on a selection", [("sel", Find("Algen"), Find("einige", 6)), ("cmd", "ShrinkFontOnePoint")]),
    ("bold+italic then Strg+Leertaste", [("sel", Find("Bakterien", 2)), ("cmd", "Bold"), ("cmd", "Italic"), ("cmd", "ResetChar")]),
    ("Strg+Leertaste at the caret, then type", [("sel", Find("Lichtenergie", 12)), ("cmd", "Bold"), ("cmd", "ResetChar"), ("type", "X")]),
    ("all caps on a selection", [("sel", Find("einige"), Find("einige", 6)), ("cmd", "AllCaps")]),
    ("small caps, caret inside a word", [("sel", Find("einige", 2)), ("cmd", "SmallCaps")]),
    ("subscript on one character", [("sel", Find("CO2", 2), Find("CO2", 3)), ("cmd", "Subscript")]),
    ("change case once", [("sel", L, Find("Lichtenergie", 12)), ("cmd", "ChangeCase")]),
    ("change case twice", [("sel", L, Find("Lichtenergie", 12)), ("cmd", "ChangeCase"), ("cmd", "ChangeCase")]),
    ("change case three times", [("sel", L, Find("Lichtenergie", 12)), ("cmd", "ChangeCase"), ("cmd", "ChangeCase"), ("cmd", "ChangeCase")]),
    ("change case, caret inside a word", [("sel", Find("Lichtenergie", 3)), ("cmd", "ChangeCase")]),
    ("change case on a sentence", [("sel", Find("Diesen Vorgang"), Find("Photosynthese.", 14)), ("cmd", "ChangeCase")]),
    ("bold at a paragraph end, Enter, type", [("sel", Find("vorkommen.", 10)), ("cmd", "Bold"), ("key", "Enter"), ("type", "Neu")]),
    ("bold paragraph, Enter at its end, type", [("sel", Find("Pflanzen, Algen"), Find("vorkommen.", 10)), ("cmd", "Bold"),
                                                ("sel", Find("vorkommen.", 10)), ("key", "Enter"), ("type", "Neu")]),
    ("bold across a paragraph end", [("sel", Find("vorkommen.", 4), Find("Für die", 3)), ("cmd", "Bold")]),
    ("type before a bold word", [("sel", Find("Lichtenergie", 4)), ("cmd", "Bold"), ("sel", L), ("type", "X")]),
    ("type after a bold word", [("sel", Find("Lichtenergie", 4)), ("cmd", "Bold"), ("sel", Find("Lichtenergie", 12)), ("type", "X")]),
    ("type at a paragraph start before bold", [("sel", Find("Pflanzen,", 2)), ("cmd", "Bold"), ("sel", Find("Pflanzen,")), ("type", "X")]),
    ("heading 1, Enter at its end, type", [("sel", Find("Für die", 3)), ("cmd", "Heading1"), ("sel", Find("Sauerstoff.", 11)),
                                          ("key", "Enter"), ("type", "X")]),
    ("Standard style on the heading", [("sel", Find("Photosynthese – ein", 3)), ("cmd", "NormalStyle")]),
    ("Backspace joins a paragraph to the heading", [("sel", Find("Pflanzen, Algen")), ("key", "Backspace")]),
    ("Delete at the heading end", [("sel", Find("Überblick", 9)), ("key", "Delete")]),
    ("Backspace joins two body paragraphs", [("sel", Find("Für die")), ("key", "Backspace")]),
    ("centre", [("sel", Find("Algen")), ("cmd", "CenterPara")]),
    ("right, then Strg+Q", [("sel", Find("Algen")), ("cmd", "RightPara"), ("cmd", "ResetPara")]),
    ("justify", [("sel", Find("Algen")), ("cmd", "JustifyPara")]),
    ("indent twice", [("sel", Find("Algen")), ("cmd", "Indent"), ("cmd", "Indent")]),
    ("indent twice, unindent once", [("sel", Find("Algen")), ("cmd", "Indent"), ("cmd", "Indent"), ("cmd", "UnIndent")]),
    ("indent over two paragraphs", [("sel", Find("Pflanzen, Algen"), Find("Für die", 3)), ("cmd", "Indent")]),
    ("hanging indent twice", [("sel", Find("Algen")), ("cmd", "HangingIndent"), ("cmd", "HangingIndent")]),
    ("hanging indent, unhang", [("sel", Find("Algen")), ("cmd", "HangingIndent"), ("cmd", "UnHang")]),
    ("double line spacing", [("sel", Find("Algen")), ("cmd", "SpacePara2")]),
    ("1.5 line spacing", [("sel", Find("Algen")), ("cmd", "SpacePara15")]),
    ("single line spacing", [("sel", Find("Algen")), ("cmd", "SpacePara1")]),
]

UNDERLINE = {0: "none", 1: "single", 2: "words", 3: "double", 4: "dotted", 6: "thick", 7: "dash", 11: "wave"}
HIGHLIGHT = {0: "none", 7: "yellow", 4: "green", 3: "cyan", 5: "magenta", 2: "blue", 6: "red", 9: "darkBlue",
             10: "darkCyan", 11: "darkGreen", 12: "darkMagenta", 13: "darkRed", 14: "darkYellow", 16: "darkGray",
             15: "lightGray", 1: "black"}
ALIGN = {0: "left", 1: "center", 2: "right", 3: "both"}
CHAR_KEYS = ("font", "size", "bold", "italic", "underline", "strike", "vertAlign", "color", "highlight",
             "caps", "smallCaps", "hidden")
PARA_KEYS = ("style", "align", "indLeft", "indFirst", "before", "after", "line", "lineRule")


def _color(c: int) -> str:
    if c in (-16777216, 0xFF000000):
        return "auto"
    return f"{c & 0xFF:02X}{(c >> 8) & 0xFF:02X}{(c >> 16) & 0xFF:02X}"


MIXED = 9999999  # wdUndefined
# Character properties as read from a Word Font object; None marks "mixed in this range".
FONT_PROPS = {
    "font": lambda f: f.Name or None,
    "size": lambda f: None if f.Size == MIXED else float(f.Size),
    "bold": lambda f: None if f.Bold == MIXED else f.Bold == -1,
    "italic": lambda f: None if f.Italic == MIXED else f.Italic == -1,
    "underline": lambda f: None if f.Underline == MIXED else UNDERLINE.get(f.Underline, str(f.Underline)),
    "strike": lambda f: None if f.StrikeThrough == MIXED else f.StrikeThrough == -1,
    "vertAlign": lambda f: None if MIXED in (f.Superscript, f.Subscript) else
        "superscript" if f.Superscript == -1 else "subscript" if f.Subscript == -1 else "baseline",
    "color": lambda f: None if f.Color == MIXED else _color(f.Color),
    "caps": lambda f: None if f.AllCaps == MIXED else f.AllCaps == -1,
    "smallCaps": lambda f: None if f.SmallCaps == MIXED else f.SmallCaps == -1,
    "hidden": lambda f: None if f.Hidden == MIXED else f.Hidden == -1,
}


def word_char(word: WordTarget, font) -> dict:
    return {k: get(font) for k, get in FONT_PROPS.items()}


def _value(word: WordTarget, key: str, a: int, b: int):
    r = word.doc.Range(a, b)
    if key == "highlight":
        v = r.HighlightColorIndex
        return None if v == MIXED else HIGHLIGHT.get(v, str(v))
    return FONT_PROPS[key](r.Font)


def _fill(word: WordTarget, key: str, a: int, b: int, out: dict, text: str) -> None:
    """Value of one property for every character in [a, b): one COM call per uniform stretch.
    Word reports a range as uniform even when spaces in it differ, so spaces (and breaks)
    are always read one by one."""
    v = _value(word, key, a, b)
    if v is not None or b - a == 1:
        for i in range(a, b):
            out[i][key] = v
        if b - a > 1:
            for i in range(a, b):
                if i >= len(text) or text[i] in " \n\t\x0b\x0c\u00a0":
                    out[i][key] = _value(word, key, i, i + 1)
        return
    m = (a + b) // 2
    _fill(word, key, a, m, out, text)
    _fill(word, key, m, b, out, text)


def word_formats(word: WordTarget, start: int, end: int) -> list[dict]:
    """Formatting of each character in [start, end] (paragraph marks included)."""
    out = {i: {} for i in range(start, end + 1)}
    text = word.text()
    for key in (*FONT_PROPS, "highlight"):
        _fill(word, key, start, end + 1, out, text)
    para: dict = {}
    for i in range(start, end + 1):
        pf = word.doc.Range(i, i).Paragraphs(1)
        pstart = pf.Range.Start
        if pstart not in para:
            rule = pf.LineSpacingRule
            para[pstart] = {
                "style": pf.Style.NameLocal, "align": ALIGN.get(pf.Alignment, str(pf.Alignment)),
                "indLeft": round(float(pf.LeftIndent), 1), "indFirst": round(float(pf.FirstLineIndent), 1),
                "before": round(float(pf.SpaceBefore), 1), "after": round(float(pf.SpaceAfter), 1),
                # Word: 0 single, 1 1.5 lines, 2 double, 3 at least, 4 exactly, 5 multiple (points of 12 pt).
                "line": {0: 1.0, 1: 1.5, 2: 2.0}.get(rule, round(float(pf.LineSpacing) / 12, 2) if rule == 5 else round(float(pf.LineSpacing), 1)),
                "lineRule": "auto" if rule in (0, 1, 2, 5) else "atLeast" if rule == 3 else "exact",
            }
        out[i].update(para[pstart])
    return [out[i] for i in range(start, end + 1)]


def word_typing(word: WordTarget) -> dict:
    return word_char(word, word.app.Selection.Font)


def ours_formats(t: WebViewTarget, start: int, end: int) -> list[dict]:
    out = t.js(f"Array.from({{length: {end - start + 1}}}, (_, i) => __probe.formatAt({start} + i))")
    for f in out:
        f["size"] = float(f["size"])
        for k in ("indLeft", "indFirst", "before", "after"):
            f[k] = round(f[k], 1)
        f["line"] = round(f["line"], 2) if f["lineRule"] == "auto" else round(f["line"], 1)
    return out


def _resolve(t, f: Find) -> int:
    return t.find(f.needle) + f.delta


def perform(t, steps: list, is_word: bool) -> None:
    for step in steps:
        kind = step[0]
        if kind == "sel":
            a = _resolve(t, step[1])
            b = _resolve(t, step[2]) if len(step) > 2 else a
            t.set_selection(a, b)
        elif kind == "cmd":
            if is_word:
                if step[1] == "Heading1":
                    t.app.Selection.Style = t.doc.Styles(-2)
                else:
                    t.wb(step[1])
            elif step[1] == "Heading1":
                # With a German keyboard Strg+Alt+1 is AltGr+1 and does nothing in Word either
                # (behaviour catalogue); Word gets the style through COM, we through the command.
                t.js("__app.run('style.h1')")
            else:
                key, ctrl, shift, alt, vk = CMD[step[1]]
                t.send_key(key, ctrl=ctrl, shift=shift, alt=alt, vk=vk)
        elif kind == "type":
            if is_word:
                t.app.Selection.TypeText(step[1])
            else:
                t.send_text(step[1])
        elif kind == "key":
            if is_word:
                s = t.app.Selection
                {"Enter": s.TypeParagraph, "Backspace": s.TypeBackspace, "Delete": lambda: s.Delete(1, 1)}[step[1]]()
            else:
                t.send_key(step[1])


def snapshot(t, is_word: bool) -> dict:
    """Text, selection, typing formatting and the formatting of the paragraphs around the
    selection (each character and the paragraph mark)."""
    text = t.text()
    anchor, head = t.selection()
    lo, hi = min(anchor, head), max(anchor, head)
    start = text.rfind("\n", 0, lo) + 1
    start = text.rfind("\n", 0, max(0, start - 1)) + 1 if start > 0 else 0  # include the previous paragraph
    end = text.find("\n", hi)
    end = len(text) if end < 0 else end
    nxt = text.find("\n", end + 1)
    end = len(text) if nxt < 0 else nxt  # and the next one
    last = min(end, len(text))
    fmt = word_formats(t, start, last) if is_word else ours_formats(t, start, last)
    if not is_word:
        # Word's Range.Text shows "all caps" text in capitals; do the same for comparison.
        chars = list(text)
        for i, f in enumerate(fmt):
            if f["caps"] and start + i < len(chars):
                chars[start + i] = chars[start + i].upper()
        text = "".join(chars)
    typing = None
    if anchor == head:
        typing = word_typing(t) if is_word else {k: v for k, v in t.js("__probe.typingFormat()").items() if k in CHAR_KEYS}
        if not is_word:
            typing["size"] = float(typing["size"])
    return {"text": text, "sel": (anchor, head), "range": (start, end), "format": fmt, "typing": typing}


def diff(ref: dict, got: dict) -> list[str]:
    out = []
    if ref["text"] != got["text"]:
        i = next((k for k, (a, b) in enumerate(zip(ref["text"], got["text"])) if a != b), min(len(ref["text"]), len(got["text"])))
        out.append(f"text differs at {i}: Word {ref['text'][max(0, i - 12):i + 12]!r} / ours {got['text'][max(0, i - 12):i + 12]!r}")
        return out
    if tuple(ref["sel"]) != tuple(got["sel"]):
        out.append(f"selection Word {ref['sel']} / ours {got['sel']}")
    start = ref["range"][0]
    for k, (a, b) in enumerate(zip(ref["format"], got["format"])):
        bad = [key for key in CHAR_KEYS + PARA_KEYS if key in a and a[key] != b.get(key)]
        if bad:
            o = start + k
            ch = ref["text"][o] if o < len(ref["text"]) else "¶"
            out.append(f"at {o} {ch!r}: " + ", ".join(f"{key} Word {a[key]!r} / ours {b.get(key)!r}" for key in bad))
            if len(out) > 6:
                break
    if ref["typing"] and got["typing"]:
        bad = [key for key in CHAR_KEYS if key in ref["typing"] and ref["typing"][key] != got["typing"].get(key)]
        if bad:
            out.append("typing: " + ", ".join(f"{key} Word {ref['typing'][key]!r} / ours {got['typing'].get(key)!r}" for key in bad))
    return out


def run(word: WordTarget, ours: WebViewTarget) -> dict:
    results = {}
    for name, steps in CASES:
        word.load("short")
        perform(word, steps, True)
        ref = snapshot(word, True)
        ours.load("short")
        ours.js("__probe.focus()")
        perform(ours, steps, False)
        got = snapshot(ours, False)
        problems = diff(ref, got)
        results[name] = {"ok": not problems, "problems": problems}
        print(("PASS " if not problems else "FAIL ") + name, flush=True)
        for p in problems:
            print("      " + p, flush=True)
    passed = sum(r["ok"] for r in results.values())
    print(f"\n{passed}/{len(results)} formatting cases match Word")
    return results
