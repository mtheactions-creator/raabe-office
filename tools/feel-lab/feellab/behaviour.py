"""F3 behaviour conformance. Each case performs real mouse/keyboard input on a fresh copy
of the "short" document and records the resulting text and selection. Word's result is
the reference. Takes over keyboard, mouse and screen.

Positions are written as Find("needle", delta): the offset of the first occurrence of
`needle` in the document text plus `delta`. All targets show the identical text."""

from __future__ import annotations

import time
from dataclasses import dataclass

from . import winapi
from .targets import Target


@dataclass(frozen=True)
class Find:
    needle: str
    delta: int = 0


# Where to point the mouse: middle of a character, caret position before a character,
# far right of the line containing a character, or in the left margin of that line.
def char(f: Find): return ("char", f)
def caret(f: Find): return ("caret", f)
def eol(f: Find): return ("eol", f)
def margin(f: Find): return ("margin", f)
def below(f: Find): return ("below", f)


CASES: list[tuple[str, list]] = [
    # ----- double / triple click ----------------------------------------------------
    ("dblclick inside word", [("click", char(Find("Lichtenergie", 4)), 2)]),
    ("dblclick word before comma", [("click", char(Find("Pflanzen", 3)), 2)]),
    ("dblclick comma", [("click", char(Find(", Algen")), 2)]),
    ("dblclick space between words", [("click", char(Find(" einige")), 2)]),
    ("dblclick last word of paragraph", [("click", char(Find("vorkommen.", 3)), 2)]),
    ("dblclick word with digits", [("click", char(Find("CO2", 1)), 2)]),
    ("dblclick hyphenated word", [("click", char(Find("Calvin-Zyklus", 2)), 2)]),
    ("dblclick word in parentheses", [("click", char(Find("Glucose", 2)), 2)]),
    ("dblclick word in quotes", [("click", char(Find("Minimumfaktor", 3)), 2)]),
    ("dblclick umlaut word", [("click", char(Find("Überblick", 3)), 2)]),
    ("tripleclick paragraph", [("click", char(Find("Wasser nimmt", 2)), 3)]),
    ("ctrl+click sentence", [("click", char(Find("Diesen Vorgang", 3)), 1, "ctrl")]),
    # ----- single clicks ------------------------------------------------------------
    ("click right of wrapped line", [("click", eol(Find("Lichtenergie")), 1)]),
    ("click right of wrapped line, type", [("click", eol(Find("Lichtenergie")), 1), ("type", "X")]),
    ("click right of last line", [("click", eol(Find("vorkommen.")), 1)]),
    ("click in paragraph spacing", [("click", below(Find("vorkommen.")), 1)]),
    ("click left margin", [("click", margin(Find("Lichtenergie")), 1)]),
    ("dblclick left margin", [("click", margin(Find("Lichtenergie")), 2)]),
    ("shift+click extends", [("click", caret(Find("Algen")), 1), ("click", caret(Find("Bakterien", 4)), 1, "shift")]),
    # ----- drag selection -----------------------------------------------------------
    ("drag inside one word", [("drag", caret(Find("Lichtenergie", 2)), caret(Find("Lichtenergie", 8)), 1)]),
    ("drag forward across words", [("drag", caret(Find("chemische", 3)), caret(Find("Energie", 3)), 1)]),
    ("drag backward across words", [("drag", caret(Find("Energie", 3)), caret(Find("chemische", 3)), 1)]),
    ("drag across paragraphs", [("drag", caret(Find("Diesen", 2)), caret(Find("benötigt", 3)), 1)]),
    ("dblclick then drag", [("drag", char(Find("Algen", 2)), char(Find("einige", 2)), 2)]),
    # ----- keyboard movement --------------------------------------------------------
    ("ctrl+right from mid-word", [("sel", Find("Lichtenergie", 3)), ("keys", "ctrl+right", 1)]),
    ("ctrl+right x4 over comma", [("sel", Find("Pflanzen")), ("keys", "ctrl+right", 4)]),
    ("ctrl+left x2", [("sel", Find("Bakterien", 4)), ("keys", "ctrl+left", 2)]),
    ("ctrl+shift+right x2", [("sel", Find("Algen")), ("keys", "ctrl+shift+right", 2)]),
    ("ctrl+right across paragraph end", [("sel", Find("vorkommen.")), ("keys", "ctrl+right", 3)]),
    ("ctrl+right over parentheses", [("sel", Find("(Glucose)", -1)), ("keys", "ctrl+right", 3)]),
    ("end on wrapped line, type", [("sel", Find("Lichtenergie")), ("keys", "end", 1), ("type", "X")]),
    ("home on wrapped line", [("sel", Find("Lichtenergie", 5)), ("keys", "home", 1)]),
    ("shift+end", [("sel", Find("Algen")), ("keys", "shift+end", 1)]),
    ("shift+right x3, shift+left", [("sel", Find("Algen")), ("keys", "shift+right", 3), ("keys", "shift+left", 1)]),
    ("down x3 keeps column", [("sel", Find("chemische", 4)), ("keys", "down", 3)]),
    ("up x2", [("sel", Find("Aufgabe 2", 10)), ("keys", "up", 2)]),
    ("shift+down", [("sel", Find("Algen")), ("keys", "shift+down", 1)]),
    ("ctrl+down x2", [("sel", Find("Algen")), ("keys", "ctrl+down", 2)]),
    ("ctrl+up mid-paragraph", [("sel", Find("Algen")), ("keys", "ctrl+up", 1)]),
    ("left collapses selection", [("sel", Find("Algen"), Find("Algen", 5)), ("keys", "left", 1)]),
    ("right collapses selection", [("sel", Find("Algen"), Find("Algen", 5)), ("keys", "right", 1)]),
    ("ctrl+a", [("sel", Find("Algen")), ("keys", "ctrl+a", 1)]),
    # ----- editing ------------------------------------------------------------------
    ("ctrl+backspace mid-word", [("sel", Find("Lichtenergie", 5)), ("keys", "ctrl+backspace", 1)]),
    ("ctrl+backspace after word", [("sel", Find("Algen und", 6)), ("keys", "ctrl+backspace", 1)]),
    ("ctrl+delete mid-word", [("sel", Find("Algen", 2)), ("keys", "ctrl+delete", 1)]),
    ("backspace at paragraph start", [("sel", Find("Für die")), ("keys", "backspace", 1)]),
    ("enter splits paragraph, type", [("sel", Find("Diesen Vorgang")), ("keys", "enter", 1), ("type", "X")]),
    ("type over double-clicked word", [("click", char(Find("Algen", 2)), 2), ("type", "Moose")]),
    ("smart cut and paste", [("click", char(Find("Algen", 2)), 2), ("keys", "ctrl+x", 1),
                             ("click", caret(Find("Bakterien")), 1), ("keys", "ctrl+v", 1)]),
    ("undo after typing two words", [("sel", Find("Algen")), ("type", "rote Algen "), ("keys", "ctrl+z", 1)]),
]

PAGE_CONTENT_WIDTH_PT = 595.3 - 2 * 70.85


def _offset(text: str, f: Find) -> int:
    i = text.find(f.needle)
    if i < 0:
        raise ValueError(f.needle)
    return i + f.delta


def _line_start(lines: list[int], offset: int) -> int:
    return max(s for s in lines if s <= offset)


def _point(t: Target, text: str, lines: list[int], spec) -> tuple[float, float]:
    kind, f = spec
    o = _offset(text, f)
    if kind == "char":
        return t.char_center(o)
    if kind == "caret":
        return t.caret_point(o)
    x0, y, h = t.point_at(_line_start(lines, o))
    scale = winapi.dpi_scale(t.hwnd)
    if kind == "eol":
        return x0 + (PAGE_CONTENT_WIDTH_PT * 4 / 3 + 30) * scale, y + h / 2
    if kind == "margin":
        return x0 - 30 * scale, y + h / 2
    if kind == "below":
        # The last line of the paragraph, then 4 px into the spacing below it.
        end = text.find("\n", o)
        xs, ys, hs = t.point_at(end if end >= 0 else len(text))
        return xs - 20 * scale, ys + hs + 6 * scale
    raise ValueError(kind)


def perform(t: Target, steps: list) -> dict:
    t.load("short")
    text = t.text()
    lines = t.lines()
    t.activate()
    time.sleep(0.7)  # longer than the double-click time, so clicks never merge across cases
    for step in steps:
        op = step[0]
        if op == "sel":
            a = _offset(text, step[1])
            h = _offset(text, step[2]) if len(step) > 2 else a
            t.set_selection(min(a, h), max(a, h))
        elif op == "click":
            x, y = _point(t, text, lines, step[1])
            winapi.click(x, y, count=step[2], modifiers=step[3] if len(step) > 3 else "")
        elif op == "drag":
            x1, y1 = _point(t, text, lines, step[1])
            x2, y2 = _point(t, text, lines, step[2])
            winapi.drag(x1, y1, x2, y2, start_clicks=step[3])
        elif op == "keys":
            for _ in range(step[2]):
                winapi.press(step[1])
                time.sleep(0.04)
        elif op == "type":
            winapi.type_text(step[1])
        time.sleep(0.25)
    time.sleep(0.3)
    result_text = t.text()
    anchor, head = t.selection()
    return {"text": result_text, "anchor": anchor, "head": head}


def _diff(a: str, b: str) -> str:
    i = 0
    while i < min(len(a), len(b)) and a[i] == b[i]:
        i += 1
    return f"@{i}: word={a[max(0, i - 15):i + 25]!r} other={b[max(0, i - 15):i + 25]!r}"


def run(targets: list[Target]) -> dict:
    raw = {t.name: {name: perform(t, steps) for name, steps in CASES} for t in targets}
    ref = raw[targets[0].name]
    cases = []
    summary = {t.name: 0 for t in targets[1:]}
    for name, _ in CASES:
        w = ref[name]
        row = {"case": name, "word": {"sel": [min(w["anchor"], w["head"]), max(w["anchor"], w["head"])], "head": w["head"],
                                      "selected": w["text"][min(w["anchor"], w["head"]):max(w["anchor"], w["head"])]}}
        for t in targets[1:]:
            r = raw[t.name][name]
            same_text = r["text"] == w["text"]
            same_sel = sorted((r["anchor"], r["head"])) == sorted((w["anchor"], w["head"]))
            ok = same_text and same_sel
            summary[t.name] += ok
            entry = {"ok": ok, "sel": sorted((r["anchor"], r["head"])), "head": r["head"]}
            if not same_text:
                entry["text_diff"] = _diff(w["text"], r["text"])
            if not same_sel:
                entry["selected"] = r["text"][min(r["anchor"], r["head"]):max(r["anchor"], r["head"])]
            row[t.name] = entry
        cases.append(row)
    return {"summary": {k: f"{v}/{len(CASES)}" for k, v in summary.items()}, "cases": cases}
