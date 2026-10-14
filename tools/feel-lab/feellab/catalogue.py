"""Behaviour catalogue (D-038, package A): scripted interactions with real keyboard and mouse
input, run on identical Word-made documents in Word and in RaabeWrite. After each case both
documents are saved and read back by Word, so the comparison of paragraphs, list labels,
styles, indents, breaks, tables and formatting uses one reader for both; the caret is
compared as (paragraph text, offset) and the selected text. Takes over keyboard, mouse and
screen.

Positions are Find("needle", delta): where `needle` first occurs, plus `delta`; each
application searches its own text, so table and field characters never need mapping."""

from __future__ import annotations

import json
import shutil
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path

from . import winapi
from .targets import RAABE_WRITE_EXE, WINDOW_RECT, raabe_write

TMP = Path(tempfile.gettempdir()) / "raabe-catalogue"

FILL = ("Pflanzen nehmen über ihre Blätter Kohlenstoffdioxid aus der Luft auf und über die Wurzeln Wasser "
        "aus dem Boden. Mit der Energie des Sonnenlichts stellen sie daraus Traubenzucker her.")


@dataclass(frozen=True)
class Find:
    needle: str
    delta: int = 0


@dataclass
class Case:
    name: str
    doc: str
    steps: list
    area: str = ""
    notes: str = ""
    skip_ours: bool = False


# ----- starting documents (built by Word) ------------------------------------------------

def build_docs(app, out: Path) -> dict[str, Path]:
    """Small documents for the cases, built with Word's own commands."""
    out.mkdir(parents=True, exist_ok=True)
    docs: dict[str, Path] = {}

    def new():
        d = app.Documents.Add()
        return d

    def save(d, name: str) -> None:
        p = out / f"{name}.docx"
        d.SaveAs2(str(p), 16)
        d.Close(0)
        docs[name] = p

    # plain: heading and three paragraphs.
    d = new()
    d.Content.Text = "\r".join(["Photosynthese", FILL, "Das Wasser nimmt die Pflanze über die Wurzeln auf.", "Letzter Absatz."])
    d.Paragraphs(1).Style = d.Styles(-2)
    save(d, "plain")

    # empty: one empty paragraph (typing, AutoFormat, AutoCorrect).
    d = new()
    save(d, "empty")

    # lists: a bulleted list with two levels and a numbered list, made like the ribbon makes them.
    d = new()
    d.Content.Text = "\r".join(["Material", "Becherglas", "Wasser", "Leitungswasser", "Salz", "Ablauf", "Wasser einfüllen", "Salz zugeben", "Umrühren", "Ende"])
    gal = app.ListGalleries
    bullets = d.Range(d.Paragraphs(2).Range.Start, d.Paragraphs(5).Range.End)
    bullets.ListFormat.ApplyListTemplate(gal(1).ListTemplates(1), False, 0, 2)
    d.Paragraphs(4).Range.ListFormat.ListIndent()
    numbers = d.Range(d.Paragraphs(7).Range.Start, d.Paragraphs(9).Range.End)
    numbers.ListFormat.ApplyListTemplate(gal(2).ListTemplates(1), False, 0, 2)
    save(d, "lists")

    # table: a paragraph, a 3 x 3 grid table with text, a paragraph.
    d = new()
    d.Content.Text = "Vor der Tabelle\rNach der Tabelle"
    r = d.Range(d.Paragraphs(1).Range.End, d.Paragraphs(1).Range.End)
    t = d.Tables.Add(r, 3, 3, 1, 0)
    for i in range(1, 4):
        for k in range(1, 4):
            t.Cell(i, k).Range.Text = f"Z{i}S{k}"
    save(d, "table")

    # format: a paragraph with bold, italic and underlined words.
    d = new()
    d.Content.Text = "Ein fettes Wort, ein kursives Wort und ein unterstrichenes Wort."
    for word, attr in (("fettes", "Bold"), ("kursives", "Italic"), ("unterstrichenes", "Underline")):
        rg = d.Content.Duplicate
        rg.Find.Execute(FindText=word)
        setattr(rg.Font, attr, True)
    save(d, "format")
    return docs


# ----- the two applications ---------------------------------------------------------------

class WordSide:
    name = "word"

    def __init__(self, app):
        self.app = app
        self.doc = None

    def open(self, path: Path) -> None:
        if self.doc is not None:
            self.doc.Close(0)
        self.doc = self.app.Documents.Open(str(path), AddToRecentFiles=False)
        win = self.doc.Windows(1)
        win.View.Type = 3
        win.View.Zoom.Percentage = 100
        win.View.ShowAll = False
        self.doc.ShowSpellingErrors = False
        self.doc.ShowGrammaticalErrors = False
        self.hwnd = win.Hwnd
        winapi.place(self.hwnd, *WINDOW_RECT)
        win.ScrollIntoView(self.doc.Range(0, 0), True)
        self.app.Selection.SetRange(0, 0)
        winapi.activate(self.hwnd)
        time.sleep(0.4)

    def locate(self, f: Find) -> int:
        r = self.doc.Content.Duplicate
        r.Find.ClearFormatting()
        if not r.Find.Execute(FindText=f.needle, MatchCase=True):
            raise ValueError(f.needle)
        return r.Start + f.delta

    def set_selection(self, a: int, h: int) -> None:
        s = self.app.Selection
        s.SetRange(min(a, h), max(a, h))
        s.StartIsActive = h < a

    def point_at(self, pos: int) -> tuple[float, float, float]:
        left, top, width, height = self.doc.Windows(1).GetPoint(0, 0, 0, 0, self.doc.Range(pos, pos))
        return float(left), float(top), float(height)

    def save(self, out: Path) -> None:
        self.doc.SaveAs2(str(out), 16)

    def caret(self) -> dict:
        s = self.app.Selection
        head = s.Start if s.StartIsActive else s.End
        # The end of the document (after the last paragraph mark) has no paragraph of its own.
        at = min(head, self.doc.Content.End - 1)
        hp = self.doc.Range(at, at).Paragraphs(1).Range
        return {"para": hp.Text.rstrip("\r\x07"), "offset": head - hp.Start, "selected": s.Text if s.End > s.Start else ""}

    def close(self) -> None:
        if self.doc is not None:
            self.doc.Close(0)
            self.doc = None


class RaabeSide:
    name = "raabe-write"

    def __init__(self):
        self.t = raabe_write("short", port=9372)
        self.hwnd = self.t.hwnd
        winapi.place(self.hwnd, *WINDOW_RECT)

    def open(self, path: Path) -> None:
        t = self.t
        t.js(f"__app.openPath({str(path)!r})")
        deadline = time.perf_counter() + 15
        while time.perf_counter() < deadline:
            docs = t.js("__app.docs()") or []
            if any(d["active"] and d["path"] and Path(d["path"]) == path for d in docs) and t.js("__probe.settled()"):
                break
            time.sleep(0.1)
        # Close the other documents, so tabs do not pile up.
        t.js("__app.closeOthers && __app.closeOthers()")
        winapi.activate(self.hwnd)
        t.js("__probe.focus()")
        t.set_selection(0, 0)
        time.sleep(0.4)

    def locate(self, f: Find) -> int:
        i = self.t.text().find(f.needle)
        if i < 0:
            raise ValueError(f.needle)
        return i + f.delta

    def set_selection(self, a: int, h: int) -> None:
        self.t.set_selection(a, h)

    def point_at(self, pos: int) -> tuple[float, float, float]:
        return self.t.point_at(pos)

    def save(self, out: Path) -> None:
        self.t.js(f"__app.saveTo({str(out)!r})")
        deadline = time.perf_counter() + 10
        while time.perf_counter() < deadline and not out.exists():
            time.sleep(0.05)
        time.sleep(0.2)

    def caret(self) -> dict:
        return self.t.js("""(() => { const e = __app.editor(); const h = e.sel.head, a = e.sel.anchor;
          const lo = (a.p < h.p || (a.p === h.p && a.o < h.o)) ? a : h, hi = lo === a ? h : a;
          const para = e.doc.paras[h.p].text;
          let sel = '';
          if (lo.p !== hi.p || lo.o !== hi.o) {
            for (let p = lo.p; p <= hi.p; p++) { const t = e.doc.paras[p].text; sel += t.slice(p === lo.p ? lo.o : 0, p === hi.p ? hi.o : t.length) + (p < hi.p ? '\\r' : ''); }
          }
          return { para, offset: h.o, selected: sel }; })()""")

    def close(self) -> None:
        self.t.close()


# ----- performing a case --------------------------------------------------------------------

def _point(side, spec) -> tuple[float, float]:
    kind, f = spec
    pos = side.locate(f)
    x, y, h = side.point_at(pos)
    if kind == "char":
        x1, y1, _ = side.point_at(pos + 1)
        if abs(y1 - y) > 1:
            x1 = x + 6
        return (x + x1) / 2, y + h / 2
    if kind == "caret":
        return x, y + h / 2
    if kind == "margin":
        return x - 40 * winapi.dpi_scale(side.hwnd), y + h / 2
    raise ValueError(kind)


def perform(side, case: Case, docs: dict[str, Path], out: Path) -> dict:
    side.open(docs[case.doc])
    for step in case.steps:
        op = step[0]
        if op == "caret":
            p = side.locate(step[1])
            side.set_selection(p, p)
        elif op == "sel":
            side.set_selection(side.locate(step[1]), side.locate(step[2]))
        elif op == "keys":
            for _ in range(step[2] if len(step) > 2 else 1):
                winapi.press(step[1])
                time.sleep(0.06)
        elif op == "type":
            winapi.type_text(step[1])
        elif op == "click":
            x, y = _point(side, step[1])
            winapi.click(x, y, count=step[2] if len(step) > 2 else 1, modifiers=step[3] if len(step) > 3 else "")
        elif op == "drag":
            x1, y1 = _point(side, step[1])
            x2, y2 = _point(side, step[2])
            winapi.drag(x1, y1, x2, y2, start_clicks=1)
        elif op == "wait":
            time.sleep(step[1])
        time.sleep(0.15)
    time.sleep(0.4)
    caret = side.caret()
    side.save(out)
    return caret


# ----- reading the results with Word ---------------------------------------------------------

def signature(app, path: Path) -> list[dict]:
    d = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
    try:
        rows = []
        for p in d.Paragraphs:
            r = p.Range
            text = r.Text.rstrip("\r\x07")
            if r.Information(31):  # end-of-row mark
                continue
            lf = r.ListFormat
            pf = p.Format
            fmt = []
            for attr in ("Bold", "Italic", "Underline", "Subscript", "Superscript", "StrikeThrough", "AllCaps"):
                v = getattr(r.Font, attr)
                if v not in (0, False):
                    if v == 9999999:
                        # Mixed: the formatted words.
                        words = [w.Text.strip() for w in r.Words if getattr(w.Font, attr) not in (0, False, 9999999) and w.Text.strip()]
                        fmt.append(f"{attr}:{'/'.join(words)}")
                    else:
                        fmt.append(f"{attr}:all")
            listdef = ""
            if lf.ListType:
                try:
                    lv = lf.ListTemplate.ListLevels(lf.ListLevelNumber)
                    listdef = (f"{lv.NumberFormat}|{lv.NumberStyle}|{lv.Font.Name}|{lv.Font.Size if lv.Font.Size < 1000 else '-'}|"
                               f"{round(lv.NumberPosition, 1)}|{round(lv.TextPosition, 1)}|{round(lv.TabPosition, 1) if lv.TabPosition < 9000 else '-'}|"
                               f"{lv.Alignment}|{lv.TrailingCharacter}")
                except Exception:  # noqa: BLE001
                    listdef = "?"
            rows.append({
                "listdef": listdef,
                "text": text.replace("\x0c", "<PAGE>").replace("\x0e", "<COL>").replace("\x0b", "<BR>"),
                "style": p.Style.NameLocal if p.Style else "",
                "list": lf.ListString if lf.ListType else "",
                "level": lf.ListLevelNumber if lf.ListType else 0,
                "indent": round(pf.LeftIndent, 1), "first": round(pf.FirstLineIndent, 1),
                "align": pf.Alignment, "table": bool(r.Information(12)),
                "spacing": [pf.LineSpacingRule, round(pf.LineSpacing, 1), round(pf.SpaceBefore, 1), round(pf.SpaceAfter, 1)],
                "size": r.Font.Size if r.Font.Size != 9999999 else "/".join(f"{w.Text.strip()}={w.Font.Size}" for w in r.Words if w.Text.strip()),
                "fmt": ";".join(fmt),
            })
        return rows
    finally:
        d.Close(0)


def _same_text(word: str, ours: str, doc_ok: bool) -> bool:
    """Paragraph texts as both report them: an object (U+FFFC) in ours stands for whatever
    Word shows for it (a symbol's "("); with equal documents, case differences come from
    "all caps", which Word shows in capitals."""
    # Characters beyond U+FFFF are two UTF-16 units in both programs' offsets.
    ours = "".join(ch if ord(ch) < 0x10000 else "￾￾" for ch in ours)
    if len(word) != len(ours):
        return False
    for a, b in zip(word, ours):
        if a == b or b == "\ufffc" or (doc_ok and a.casefold() == b.casefold()):
            continue
        # Emoji: Word's text shows placeholder characters for them (one per UTF-16 unit).
        if doc_ok and (b == "￾" or ord(b) >= 0x2600):
            continue
        return False
    return True


def _sel_text(t: str) -> str:
    """Selected text without paragraph and cell marks (Word's "\r\x07" ends a cell, a second
    one the row)."""
    while "\r\x07\r\x07" in t:
        t = t.replace("\r\x07\r\x07", "\r\x07")
    return t.replace("\r\x07", "\n").replace("\x07", "").replace("\r", "\n").rstrip("\n")


def _word_caret(wc: dict) -> dict:
    """Word can select the last paragraph mark (offset one past the text); RaabeWrite's
    selection ends before it, where the document ends."""
    if wc["offset"] == len(wc["para"]) + 1 and wc["selected"].endswith("\r"):
        return {**wc, "offset": len(wc["para"])}
    return wc


def compare(word_sig: list[dict], our_sig: list[dict], wc: dict, oc: dict) -> dict:
    diffs = []
    for i in range(max(len(word_sig), len(our_sig))):
        a = word_sig[i] if i < len(word_sig) else None
        b = our_sig[i] if i < len(our_sig) else None
        if a != b:
            keys = sorted({*(a or {}), *(b or {})})
            diffs.append({"para": i, "word": {k: (a or {}).get(k) for k in keys if (a or {}).get(k) != (b or {}).get(k)},
                          "ours": {k: (b or {}).get(k) for k in keys if (a or {}).get(k) != (b or {}).get(k)},
                          "text": (a or b or {}).get("text", "")[:50]})
    wc = _word_caret(wc)
    caret_ok = _same_text(wc["para"], oc["para"], not diffs) and wc["offset"] == oc["offset"]
    sel_ok = _sel_text(wc["selected"]) == _sel_text(oc["selected"]) or (
        not diffs and _sel_text(wc["selected"]).casefold() == _sel_text(oc["selected"]).casefold())
    if "\x07" in wc["selected"]:
        # Selected cells: Word's COM text of a block and its active end are unreliable, so one
        # of the two must agree.
        caret_ok = sel_ok = caret_ok or sel_ok
    return {"ok": not diffs and caret_ok and sel_ok, "doc_ok": not diffs, "caret_ok": caret_ok, "sel_ok": sel_ok,
            "diffs": diffs[:6], "caret": {"word": wc, "ours": oc}}


def run(cases: list[Case], only: list[str] | None = None) -> dict:
    import win32com.client
    if TMP.exists():
        shutil.rmtree(TMP, ignore_errors=True)
    (TMP / "results").mkdir(parents=True)
    reader = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    reader.Visible = False
    word_app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    word_app.Visible = True
    word_app.Options.CheckSpellingAsYouType = False
    word_app.Options.CheckGrammarAsYouType = False
    picked = [c for c in cases if not only or any(o in c.name or o == c.area for o in only)]
    results: dict = {}
    try:
        docs = build_docs(reader, TMP / "docs")
        word = WordSide(word_app)
        wcaret = {}
        for c in picked:
            try:
                wcaret[c.name] = perform(word, c, docs, TMP / "results" / f"{_safe(c.name)}-word.docx")
            except Exception as e:  # noqa: BLE001
                wcaret[c.name] = {"error": repr(e)}
            print("word:", c.name, flush=True)
        word.close()
        ours = RaabeSide()
        try:
            ocaret = {}
            for c in picked:
                if c.skip_ours:
                    continue
                try:
                    ocaret[c.name] = perform(ours, c, docs, TMP / "results" / f"{_safe(c.name)}-ours.docx")
                except Exception as e:  # noqa: BLE001
                    ocaret[c.name] = {"error": repr(e)}
                print("ours:", c.name, flush=True)
        finally:
            ours.close()
        for c in picked:
            wc, oc = wcaret.get(c.name, {}), ocaret.get(c.name, {})
            if "error" in wc or "error" in oc or not oc:
                results[c.name] = {"ok": False, "error": wc.get("error") or oc.get("error") or "not run", "area": c.area}
                continue
            ws = signature(reader, TMP / "results" / f"{_safe(c.name)}-word.docx")
            os_ = signature(reader, TMP / "results" / f"{_safe(c.name)}-ours.docx")
            results[c.name] = {"area": c.area, **compare(ws, os_, wc, oc)}
    finally:
        try:
            word_app.Quit(0)
        finally:
            reader.Quit(0)
    passed = sum(1 for r in results.values() if r.get("ok"))
    return {"summary": f"{passed}/{len(results)}", "by_area": _by_area(results), "cases": results}


def _by_area(results: dict) -> dict:
    out: dict = {}
    for r in results.values():
        a = out.setdefault(r.get("area", ""), [0, 0])
        a[0] += bool(r.get("ok"))
        a[1] += 1
    return {k: f"{v[0]}/{v[1]}" for k, v in out.items()}


def _safe(name: str) -> str:
    return "".join(ch if ch.isalnum() else "_" for ch in name)[:60]


def recompare(previous: dict) -> dict:
    """Compares the saved result documents of the last run again (after changing the comparison)."""
    import win32com.client
    reader = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    reader.Visible = False
    results: dict = {}
    try:
        for name, r in previous["cases"].items():
            if "caret" not in r:
                results[name] = r
                continue
            w, o = TMP / "results" / f"{_safe(name)}-word.docx", TMP / "results" / f"{_safe(name)}-ours.docx"
            if not (w.exists() and o.exists()):
                results[name] = r
                continue
            results[name] = {"area": r.get("area", ""), **compare(signature(reader, w), signature(reader, o), r["caret"]["word"], r["caret"]["ours"])}
    finally:
        reader.Quit(0)
    passed = sum(1 for r in results.values() if r.get("ok"))
    return {"summary": f"{passed}/{len(results)}", "by_area": _by_area(results), "cases": results}
