"""Equation behaviour catalogue (D-039, package D): key sequences typed into a new equation
(Alt+Umschalt+0) in Word and in RaabeWrite with real keyboard input, ending with a marker
character "§" where the caret is (or replacing what is selected). The equations are compared
by their structure in the saved file (OMML without properties that do not change the look).
Word's results are the specification for our own formula editor. Takes over keyboard and
screen.

    .venv/Scripts/python.exe catalogue_math.py [--word-only | --ours-only] [name filter …]

--ours-only replays the cases in RaabeWrite against Word's recorded results.

Results: results/<date>-catalogue-math.json.
"""
from __future__ import annotations

import datetime as dt
import json
import re
import sys
import time
import zipfile
from pathlib import Path

from feellab import winapi
from feellab.catalogue import TMP, RaabeSide, WordSide, build_docs

HERE = Path(__file__).parent
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
MARK = "§"

# Each case: name, steps after the new equation (("type", text) / ("keys", combo, n)), area.
CASES: list[tuple[str, list[tuple], str]] = []


def case(name: str, *steps: tuple, area: str = "math") -> None:
    CASES.append((name, list(steps), area))


T = lambda s: ("type", s)  # noqa: E731
K = lambda k, n=1: ("keys", k, n)  # noqa: E731

# ----- typing and building up (Word's linear format and Math AutoCorrect) ----------------------
case("plain letters", T("abc"), area="type")
case("letters and digits", T("x+2y=3"), area="type")
case("space between letters", T("a b"), area="type")
case("power without space", T("x^2"), area="buildup")
case("power with space", T("x^2 "), area="buildup")
case("power then operator", T("x^2+1"), area="buildup")
case("subscript", T("x_1 "), area="buildup")
case("sub and sup", T("x_1^2 "), area="buildup")
case("fraction", T("a/b "), area="buildup")
case("fraction of numbers", T("1/2 "), area="buildup")
case("fraction with brackets", T("(a+b)/c "), area="buildup")
case("fraction then text", T("a/b+c "), area="buildup")
case("power in brackets", T("x^(2n) "), area="buildup")
case("negative exponent", T("e^-x "), area="buildup")
case("alpha", T("\\alpha "), area="autocorrect")
case("pi r squared", T("\\pi r^2 "), area="autocorrect")
case("cdot", T("a\\cdot b"), area="autocorrect")
case("cdot space", T("a\\cdot b "), area="autocorrect")
case("less or equal", T("x<=y "), area="autocorrect")
case("not equal", T("a/=b "), area="autocorrect")
case("arrow", T("a->b "), area="autocorrect")
case("infinity", T("\\infty "), area="autocorrect")
case("times", T("a\\times b "), area="autocorrect")
case("square root", T("\\sqrt(x) "), area="buildup")
case("square root of number", T("\\sqrt 2 "), area="buildup")
case("integral", T("\\int_0^1 x dx "), area="buildup")
case("sum", T("\\sum_(i=1)^n i "), area="buildup")
case("sin x", T("sin x "), area="buildup")
case("brackets", T("(a+b) "), area="buildup")
case("absolute value", T("|x| "), area="buildup")
case("matrix", T("\\matrix(a&b@c&d) "), area="buildup")
case("bar accent", T("x\\bar "), area="buildup")
case("degree", T("90\\degree "), area="autocorrect")

# ----- caret movement in built structures ------------------------------------------------------
for n in range(1, 7):
    case(f"fraction left x{n}", T("a/b "), K("left", n), area="navigate")
for n in range(1, 6):
    case(f"power left x{n}", T("x^2 "), K("left", n), area="navigate")
for n in range(1, 5):
    case(f"root left x{n}", T("\\sqrt(x) "), K("left", n), area="navigate")
for n in range(1, 8):
    case(f"sub sup left x{n}", T("x_1^2 "), K("left", n), area="navigate")
case("fraction home", T("a/b "), K("home"), area="navigate")
case("fraction end", T("a/b "), K("home"), K("end"), area="navigate")
case("fraction up from denominator", T("a/b "), K("left"), K("up"), area="navigate")
case("fraction down from numerator", T("a/b "), K("left", 3), K("down"), area="navigate")
case("right after home", T("a/b "), K("home"), K("right"), area="navigate")
case("right x2 after home", T("a/b "), K("home"), K("right", 2), area="navigate")
case("right x3 after home", T("a/b "), K("home"), K("right", 3), area="navigate")
case("leave equation right", T("x "), K("right"), area="navigate")
case("leave equation left", T("x "), K("home"), K("left"), area="navigate")
case("ctrl right", T("ab+cd "), K("home"), K("ctrl+right"), area="navigate")

# ----- Backspace, Delete, selection -------------------------------------------------------------
case("backspace after fraction", T("a/b "), K("backspace"), area="delete")
case("backspace twice after fraction", T("a/b "), K("backspace", 2), area="delete")
case("backspace after power", T("x^2 "), K("backspace"), area="delete")
case("backspace twice after power", T("x^2 "), K("backspace", 2), area="delete")
case("backspace in denominator start", T("a/b "), K("left", 2), K("backspace"), area="delete")
case("backspace in empty numerator", T("a/b "), K("left", 3), K("backspace"), K("backspace"), area="delete")
case("delete before fraction", T("a/b "), K("home"), K("delete"), area="delete")
case("delete twice before fraction", T("a/b "), K("home"), K("delete", 2), area="delete")
case("backspace letter", T("abc"), K("backspace"), area="delete")
case("shift left over fraction", T("a/b "), K("shift+left"), area="select")
case("shift left twice", T("x+a/b "), K("shift+left", 2), area="select")
case("shift home", T("x+a/b "), K("shift+home"), area="select")
case("shift right in numerator", T("ab/c "), K("left", 4), K("shift+right"), area="select")
case("shift right out of numerator", T("ab/c "), K("left", 4), K("shift+right", 3), area="select")

# ----- Enter -----------------------------------------------------------------------------------
case("enter in equation", T("x=1"), K("enter"), area="enter")
case("shift enter in equation", T("x=1"), K("shift+enter"), area="enter")

# ----- more build-up, navigation and editing (second round) ----------------------------------
case("nth root", T("\\sqrt(3&x) "), area="buildup")
case("abs then power", T("|x|^2 "), area="buildup")
case("power of fraction", T("(a/b)^2 "), area="buildup")
case("fraction of power", T("a^2/b "), area="buildup")
case("root in fraction", T("\\sqrt(x)/2 "), area="buildup")
case("product", T("\\prod_(k=1)^n k "), area="buildup")
case("limit", T("lim_(n->\\infty) a_n "), area="buildup")
case("vector accent", T("v\\vec "), area="buildup")
case("hat accent", T("x\\hat "), area="buildup")
case("double integral", T("\\iint_D f dA "), area="buildup")
case("plus minus", T("x=+-1 "), area="autocorrect")
case("capital delta", T("\\Delta x "), area="autocorrect")
case("dots", T("1,2,...,n "), area="autocorrect")
case("cos of fraction", T("cos x/2 "), area="buildup")
case("power down to base", T("x^2 "), K("left"), K("down"), area="navigate")
case("power up from base", T("x^2 "), K("left", 2), K("up"), area="navigate")
case("sub sup up", T("x_1^2 "), K("left", 3), K("up"), area="navigate")
case("sub sup down", T("x_1^2 "), K("left"), K("down"), area="navigate")
case("sum up from body", T("\\sum_(i=1)^n i "), K("up"), area="navigate")
case("sum down from body", T("\\sum_(i=1)^n i "), K("down"), area="navigate")
case("sum left x3", T("\\sum_(i=1)^n i "), K("left", 3), area="navigate")
case("matrix right x3", T("\\matrix(a&b@c&d) "), K("home"), K("right", 3), area="navigate")
case("matrix down", T("\\matrix(a&b@c&d) "), K("home"), K("right", 2), K("down"), area="navigate")

# ----- Step 4c C: what "/" takes before it, quoted text ---------------------------------------
case("fraction number before", T("12/5 "), area="fraction")
case("fraction term before", T("2x/3 "), area="fraction")
case("fraction after plus", T("1+2x/3 "), area="fraction")
case("fraction after equals", T("y=3a/4b "), area="fraction")
case("fraction decimal comma", T("0,5/2 "), area="fraction")
case("fraction decimal point", T("2.5/2 "), area="fraction")
case("two fractions", T("1/2+3/4 "), area="fraction")
case("fraction after power", T("x^2/3 "), area="fraction")
case("fraction with minus", T("-x/2 "), area="fraction")
case("fraction greek", T("\\pi/2 "), area="fraction")
case("fraction greek symbol", T("2\\pi /3 "), area="fraction")
case("fraction subscript", T("x_1/2 "), area="fraction")
case("fraction after times", T("3\\cdot a/b "), area="fraction")
case("fraction denominator brackets", T("1/(a+b) "), area="fraction")
case("fraction bracket factor", T("2(a+b)/c "), area="fraction")
case("fraction after function", T("sin x/2 "), area="fraction")
case("fraction after comma list", T("a,b/c "), area="fraction")
case("fraction no space yet", T("2x/3"), area="fraction")
case("fraction then operator", T("2x/3+1"), area="fraction")
case("quoted text", T('"Text" '), area="text")
case("quoted two words", T('"zwei Worte" '), area="text")
case("quoted after equals", T('a="wenn" b '), area="text")
case("quoted no space", T('"abc"'), area="text")
case("quote open", T('"ab'), area="text")
case("quoted then left", T('"abc" '), K("left"), area="text")
case("quoted then left x2", T('"abc" '), K("left", 2), area="text")
case("quoted then left x4", T('"abc" '), K("left", 4), area="text")
case("quoted then left x5", T('"abc" '), K("left", 5), area="text")
case("quoted home right", T('"abc" '), K("home"), K("right"), area="text")
case("quoted backspace", T('"abc" '), K("backspace"), area="text")
case("quoted backspace x2", T('"abc" '), K("backspace", 2), area="text")
case("quoted then type", T('"abc" x'), area="text")
case("brackets left x2", T("(a+b) "), K("left", 2), area="navigate")
case("ctrl left", T("ab+cd "), K("ctrl+left"), area="navigate")
case("type after fraction", T("a/b "), K("left"), T("c"), area="type")
case("type into numerator", T("a/b "), K("left", 3), T("x"), area="type")
case("type over selection", T("abc"), K("shift+left", 2), T("x"), area="select")
case("shift end", T("a+b"), K("home"), K("shift+end"), area="select")
case("shift right over fraction", T("x+a/b "), K("home"), K("shift+right", 3), area="select")
case("delete in middle", T("abc"), K("left", 2), K("delete"), area="delete")
case("backspace after sum", T("\\sum_(i=1)^n i "), K("right"), K("backspace"), area="delete")
case("backspace into root", T("\\sqrt(x) "), K("backspace"), area="delete")

# ----- the structure of equations in a saved file -----------------------------------------------

# Properties that change what an equation is (the rest are fonts and control properties).
KEEP_PR = {"type", "chr", "limLoc", "begChr", "endChr", "sepChr", "pos", "degHide", "subHide", "supHide", "sty", "scr", "nor", "count"}


def norm(el) -> str:
    tag = el.tag.replace(M, "")
    if tag == "t":
        return repr(el.text or "")
    if tag == "r":
        texts = "".join((t.text or "") for t in el.iter(M + "t"))
        rpr = el.find(M + "rPr")
        props = ""
        if rpr is not None:
            props = ",".join(f"{c.tag.replace(M, '')}={c.get(M + 'val', '1')}" for c in rpr if c.tag.replace(M, "") in KEEP_PR)
        if not texts:
            return ""
        return f"r[{props}]{texts!r}" if props else repr(texts)
    if tag.endswith("Pr"):
        vals = [f"{c.tag.replace(M, '')}={c.get(M + 'val', '1')}" for c in el if c.tag.replace(M, "") in KEEP_PR]
        return f"@{','.join(vals)}" if vals else ""
    if tag in ("ctrlPr",) or not el.tag.startswith(M):
        return ""
    inner = [norm(c) for c in el]
    inner = [x for x in inner if x]
    return f"{tag}({' '.join(inner)})"


def merge_runs(s: str) -> str:
    """Neighbouring runs with the same properties as one (Word splits a run where a character
    is typed into it; the look is the same)."""
    pat = re.compile(r"(?<![\]\w])(r\[[^\]]*\]|)'((?:[^'\\]|\\.)*)' \1'((?:[^'\\]|\\.)*)'")
    while True:
        t = pat.sub(lambda m: f"{m.group(1)}'{m.group(2)}{m.group(3)}'", s, count=1)
        if t == s:
            return s
        s = t


def equations(path: Path) -> list[str]:
    import xml.etree.ElementTree as ET
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    out = []
    for p in root.iter(W + "p"):
        # The paragraph's pieces in order: text runs and equations (display ones in oMathPara).
        parts: list[str] = []
        for c in p:
            if c.tag == M + "oMathPara":
                parts.append("display[" + " | ".join(norm(m) for m in c.iter(M + "oMath")) + "]")
            elif c.tag == M + "oMath":
                parts.append(norm(c))
            elif c.tag in (W + "r", W + "ins"):
                t = "".join(x.text or "" for x in c.iter(W + "t"))
                if t:
                    parts.append(repr(t))
        if parts:
            out.append(" ".join(parts))
    return out


def perform(side, steps: list[tuple], docs: dict, out: Path) -> None:
    side.open(docs["empty"])
    winapi.press("alt+shift+0")
    time.sleep(0.5)
    for step in steps:
        if step[0] == "type":
            winapi.type_text(step[1], gap=0.05)
        else:
            for _ in range(step[2]):
                winapi.press(step[1])
                time.sleep(0.08)
        time.sleep(0.2)
    winapi.type_text(MARK)
    time.sleep(0.4)
    side.save(out)


def dump_cases() -> None:
    """The cases as JSON, for the editor's unit test (packages/ooxml/test/math-editor.test.ts)."""
    data = [{"name": n, "steps": [list(st) for st in steps], "area": a} for n, steps, a in CASES]
    (HERE / "results" / "math-cases.json").write_text(json.dumps(data, indent=0, ensure_ascii=False), encoding="utf-8", newline="\n")


def record_word(picked: list, docs: dict, out_dir: Path, results: dict) -> None:
    """Word's results for the cases (an own visible Word instance)."""
    import win32com.client
    word_app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    word_app.Visible = True
    word_app.Options.CheckSpellingAsYouType = False
    word_app.Options.CheckGrammarAsYouType = False
    try:
        word = WordSide(word_app)
        for name, steps, area in picked:
            path = out_dir / f"{re.sub(r'[^A-Za-z0-9]+', '_', name)}-word.docx"
            try:
                perform(word, steps, docs, path)
                results[name] = {"area": area, "word": equations(path)}
            except Exception as e:  # noqa: BLE001
                results[name] = {"area": area, "error": repr(e)}
            print("word:", name, results[name].get("word"), flush=True)
        word.close()
    finally:
        word_app.Quit(0)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    if "--dump-cases" in sys.argv:
        dump_cases()
        return
    import win32com.client
    word_only = "--word-only" in sys.argv
    ours_only = "--ours-only" in sys.argv
    only = [a for a in sys.argv[1:] if not a.startswith("--")]
    picked = [c for c in CASES if not only or any(o in c[0] or o == c[2] for o in only)]
    out_dir = TMP / "math"
    out_dir.mkdir(parents=True, exist_ok=True)
    dest = HERE / "results" / f"{dt.date.today()}-catalogue-math.json"
    if ours_only:
        recorded = sorted((HERE / "results").glob("*-catalogue-math.json"))
        old_word = json.loads(recorded[-1].read_text(encoding="utf-8")) if recorded else {}
        picked = [c for c in picked if old_word.get(c[0], {}).get("word") is not None]
    reader = win32com.client.DispatchEx("Word.Application")
    reader.Visible = False
    results: dict = {}
    try:
        docs = build_docs(reader, TMP / "docs-math")
    finally:
        reader.Quit(0)
    if ours_only:
        results = {name: {"area": area, "word": old_word[name]["word"]} for name, _s, area in picked}
    else:
        record_word(picked, docs, out_dir, results)
    if not word_only:
        ours = RaabeSide()
        try:
            for name, steps, area in picked:
                path = out_dir / f"{re.sub(r'[^A-Za-z0-9]+', '_', name)}-ours.docx"
                try:
                    perform(ours, steps, docs, path)
                    results[name]["ours"] = equations(path)
                except Exception as e:  # noqa: BLE001
                    results[name]["ours_error"] = repr(e)
                r = results[name]
                r["ok"] = [merge_runs(x) for x in r.get("word") or []] == [merge_runs(x) for x in r.get("ours") or []] and r.get("word") is not None
                print("ours:", name, "ok" if r["ok"] else f"DIFF {r.get('ours')}", flush=True)
        finally:
            ours.close()
    old = json.loads(dest.read_text(encoding="utf-8")) if dest.exists() else {}
    for k, v in results.items():
        old[k] = {**old.get(k, {}), **v}
    dest.write_text(json.dumps(old, indent=1, ensure_ascii=False), encoding="utf-8", newline="\n")
    if not word_only:
        ok = sum(1 for v in results.values() if v.get("ok"))
        print(f"{ok}/{len(results)} cases as in Word")


if __name__ == "__main__":
    main()
