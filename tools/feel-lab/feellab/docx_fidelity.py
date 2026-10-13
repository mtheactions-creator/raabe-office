"""F4b layout fidelity on real .docx files: the same file opened in Word and in RaabeWrite;
compares where lines and pages start, the left edge of every line and its baseline (points).

Baselines come from PDF exports (Word's own and RaabeWrite's), read with feellab.pdftext:
Word's COM positions are rounded to screen pixels and its line boxes include paragraph
spacing, while a PDF holds exactly where the text is drawn. Needs no keyboard or mouse
input: Word runs invisibly, RaabeWrite off screen."""

from __future__ import annotations

import statistics
import tempfile
import time
from pathlib import Path

from . import pdftext
from .targets import WebViewTarget

# Word writes optional and non-breaking hyphens as control characters in Range.Text.
WORD_CHARS = {"\x1f": "\u00ad", "\x1e": "\u2011"}


class WordDocument:
    """A .docx opened read-only in a separate, invisible Word process."""

    def __init__(self, path: Path):
        import win32com.client
        self.app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
        self.app.Visible = False
        try:
            self.doc = self.app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
            self.win = self.doc.Windows(1)
            self.win.View.Type = 3  # print layout
            self.win.View.Zoom.Percentage = 100
            self.win.View.ShowAll = False
        except Exception:
            self.app.Quit(0)
            raise

    def text(self) -> str:
        t = self.doc.Content.Text
        if t.endswith("\r"):
            t = t[:-1]
        for a, b in WORD_CHARS.items():
            t = t.replace(a, b)
        return t.replace("\r", "\n")

    def lines(self) -> list[int]:
        s = self.app.Selection
        end = self.doc.Content.End - 1
        s.HomeKey(6)  # wdStory
        starts: list[int] = []
        while True:
            s.HomeKey(5)  # wdLine
            if starts and s.Start <= starts[-1]:
                break
            starts.append(s.Start)
            if s.MoveDown(5, 1) == 0:
                # MoveDown stops before a paragraph with "page break before"; step over the line end.
                s.EndKey(5)
                if s.Start >= end:
                    break
                if self.doc.Range(s.Start, s.Start + 1).Text in ("\r", "\x0c", "\x0b"):
                    s.MoveRight(1, 1)
        return starts

    def boxes(self, starts: list[int]) -> list[dict]:
        """Page, x, y and height of each line in points from the page's top left.

        Word reports Information(5/6) in whole screen pixels and Information(5) ignores the
        alignment, so: x comes from GetPoint at 100 %, y and height from GetPoint at 400 %
        (1/8 pt resolution); the page origins are calibrated against Information on
        left-aligned, unindented lines."""
        import ctypes
        dpi = ctypes.windll.user32.GetDpiForSystem() / 72
        raw = []
        for o in starts:
            r = self.doc.Range(o, o)
            para = r.Paragraphs(1)
            raw.append({
                "start": o, "page": int(r.Information(3)) - 1, "info_x": float(r.Information(5)), "info_y": float(r.Information(6)),
                "left_aligned": para.Alignment == 0 and para.LeftIndent == 0 and para.FirstLineIndent == 0,
            })
        for zoom in (100, 400):
            self.win.View.Zoom.Percentage = zoom
            for b in raw:
                left, top, _w, height = self.win.GetPoint(0, 0, 0, 0, self.doc.Range(b["start"], b["start"]))
                b[f"px{zoom}"], b[f"py{zoom}"], b[f"ph{zoom}"] = left, top, height
        self.win.View.Zoom.Percentage = 100

        def origin(page: int, key: str, info: str, ppp: float) -> float:
            on_page = [b for b in raw if b["page"] == page]
            refs = [b for b in on_page if b["left_aligned"]] or on_page
            return statistics.median(b[key] - b[info] * ppp for b in refs)

        out = []
        for page in sorted({b["page"] for b in raw}):
            ox = origin(page, "px100", "info_x", dpi)
            oy = origin(page, "py400", "info_y", dpi * 4)
            for b in (b for b in raw if b["page"] == page):
                out.append({"start": b["start"], "page": page, "x": round((b["px100"] - ox) / dpi, 2),
                            "y": round((b["py400"] - oy) / (dpi * 4), 2), "h": round(b["ph400"] / (dpi * 4), 2)})
        return out

    def xs(self, offsets: list[int]) -> list[float]:
        """Caret x of each offset in points, at 400 % (only differences within a line are
        meaningful: the window origin is arbitrary)."""
        import ctypes
        ppp = ctypes.windll.user32.GetDpiForSystem() / 72 * 4
        self.win.View.Zoom.Percentage = 400
        try:
            return [self.win.GetPoint(0, 0, 0, 0, self.doc.Range(o, o))[0] / ppp for o in offsets]
        finally:
            self.win.View.Zoom.Percentage = 100

    def export_pdf(self, path: Path) -> None:
        self.doc.ExportAsFixedFormat(str(path), 17)  # wdExportFormatPDF

    def close(self) -> None:
        try:
            self.doc.Close(0)
        finally:
            self.app.Quit(0)


def open_in_raabe(target: WebViewTarget, path: Path, timeout: float = 20) -> None:
    target.js(f"__app.openPath({str(path)!r})")
    deadline = time.perf_counter() + timeout
    while time.perf_counter() < deadline:
        docs = target.js("__app.docs()") or []
        if any(d["active"] and d["path"] and Path(d["path"]) == path for d in docs):
            # Font metrics arrive asynchronously (the font database loads after start-up).
            while time.perf_counter() < deadline and not target.js("__probe.settled()"):
                time.sleep(0.1)
            time.sleep(0.2)
            return
        time.sleep(0.1)
    raise TimeoutError(f"{path.name} did not open in RaabeWrite")


def pdf_baselines(path: Path, boxes: list[dict]) -> dict[int, float]:
    """Baseline of each line (by start offset) from a PDF, matching lines page by page in
    order. Pages whose line count differs are left out (reported as None)."""
    pages = pdftext.read(path)
    out: dict[int, float] = {}
    for page, pdf_page in enumerate(pages):
        lines = [b for b in boxes if b["page"] == page]
        found = pdftext.baselines(pdf_page)
        if len(found) == len(lines):
            for b, f in zip(lines, found):
                out[b["start"]] = f["y"]
    return out


def export_raabe_pdf(target: WebViewTarget, path: Path, timeout: float = 30) -> None:
    path.unlink(missing_ok=True)
    target.js(f"__app.exportPdfTo({str(path)!r})")
    deadline = time.perf_counter() + timeout
    while time.perf_counter() < deadline:
        if path.exists() and path.stat().st_size > 0:
            time.sleep(0.3)
            return
        time.sleep(0.1)
    raise TimeoutError("RaabeWrite did not export the PDF")


def word_offsets(starts: list[int], text: str) -> dict[int, list[int]]:
    """For each line: the offsets of its word starts and of its content end."""
    out: dict[int, list[int]] = {}
    ends = starts[1:] + [len(text)]
    for a, b in zip(starts, ends):
        line = text[a:b]
        offs = [a + i for i in range(1, len(line)) if line[i - 1] == " " and line[i] not in " \n"]
        end = a + len(line.rstrip(" \n\x0b\x0c"))
        if end > a:
            offs.append(end)
        out[a] = offs
    return out


def horizontal(word_x: dict[int, float], ours_x: dict[int, float], groups: dict[int, list[int]], text: str) -> dict:
    """Word positions within each line relative to its start: Word versus ours."""
    rows = []
    for start, offs in groups.items():
        if not offs or start not in word_x:
            continue
        d = [round((ours_x[o] - ours_x[start]) - (word_x[o] - word_x[start]), 2) for o in offs]
        rows.append({"start": start, "end_drift": d[-1], "max": max(d, key=abs), "words": len(offs),
                     "context": text[start:start + 40]})
    vals = [abs(r["max"]) for r in rows]
    return {
        "lines": len(rows), "max": max(vals, default=0), "mean": round(statistics.mean(vals), 3) if vals else 0,
        "within_0.25pt": sum(v <= 0.25 for v in vals),
        "worst": sorted(rows, key=lambda r: -abs(r["max"]))[:10],
        "rows": rows,
    }


def compare(word: list[dict], ours: list[dict], text: str) -> dict:
    by_start = {b["start"]: b for b in ours}
    ref = [b["start"] for b in word]
    cand = [b["start"] for b in ours]
    matched = [w for w in word if w["start"] in by_start]
    diffs = []
    for w in matched:
        o = by_start[w["start"]]
        d = {"start": w["start"], "page": (w["page"], o["page"]), "dx": round(o["x"] - w["x"], 2),
             "dy": round(o["y"] - w["y"], 2), "dh": round(o["h"] - w["h"], 2),
             "context": text[w["start"]:w["start"] + 30]}
        # Baselines: ours from the layout model and from our PDF, Word's from its PDF.
        model = round(o["y"] + o["base"], 2)
        d["base_ours"], d["base_word"], d["para"] = model, w.get("base"), text.count(chr(10), 0, w["start"])
        if w.get("base") is not None:
            d["dbase"] = round(model - w["base"], 2)
            if o.get("pdf_base") is not None:
                d["dbase_pdf"] = round(o["pdf_base"] - w["base"], 2)
        if o.get("pdf_base") is not None:
            d["render"] = round(o["pdf_base"] - model, 2)
        diffs.append(d)
    only_word = sorted(set(ref) - set(cand))
    only_ours = sorted(set(cand) - set(ref))
    ctx = lambda o: text[max(0, o - 25):o] + "|" + text[o:o + 25]
    same_page = [d for d in diffs if d["page"][0] == d["page"][1]]
    # Line pitch (distance to the next line on the same page) is free of page-origin error.
    wy = {w["start"]: w for w in word}
    for a, b in zip(diffs, diffs[1:]):
        if a["page"][0] == b["page"][0] == a["page"][1] == b["page"][1]:
            a["dpitch"] = round((by_start[b["start"]]["y"] - by_start[a["start"]]["y"]) - (wy[b["start"]]["y"] - wy[a["start"]]["y"]), 2)
    def stat(key: str, within: float = 0.5) -> dict:
        vals = [abs(d[key]) for d in same_page if key in d]
        return {"max": max(vals, default=0), "mean": round(statistics.mean(vals), 2) if vals else 0,
                f"within_{within}pt": sum(v <= within for v in vals), "of": len(vals)}
    return {
        "lines_word": len(ref), "lines_ours": len(cand), "identical_line_starts": len(matched),
        "percent": round(100 * len(matched) / max(len(ref), 1), 2),
        "pages_word": max((b["page"] for b in word), default=0) + 1, "pages_ours": max((b["page"] for b in ours), default=0) + 1,
        "page_mismatches": sum(d["page"][0] != d["page"][1] for d in diffs),
        "x": stat("dx"), "baseline": stat("dbase", 0.2), "baseline_pdf": stat("dbase_pdf", 0.2),
        "render_vs_model": stat("render", 0.2), "box_y": stat("dy"), "box_h": stat("dh"),
        "pitch": {"max": max((abs(d["dpitch"]) for d in diffs if "dpitch" in d), default=0),
                  "within_0.2pt": sum(abs(d["dpitch"]) <= 0.2 for d in diffs if "dpitch" in d),
                  "of": sum("dpitch" in d for d in diffs)},
        "only_in_word": [{"offset": o, "context": ctx(o)} for o in only_word[:10]],
        "only_in_ours": [{"offset": o, "context": ctx(o)} for o in only_ours[:10]],
        "worst": sorted(same_page, key=lambda d: -max(abs(d["dx"]), abs(d.get("dbase", 0)), abs(d.get("dbase_pdf", 0))))[:12],
        "lines": diffs,
    }


def run(raabe: WebViewTarget, files: list[Path]) -> dict:
    out: dict = {}
    tmp = Path(tempfile.gettempdir()) / "raabe-feel-pdf"
    tmp.mkdir(exist_ok=True)
    for path in files:
        word = WordDocument(path)
        word_pdf = tmp / f"{path.stem}-word.pdf"
        try:
            text = word.text()
            starts = word.lines()
            wboxes = word.boxes(starts)
            groups = word_offsets(starts, text)
            all_offsets = sorted({o for offs in groups.values() for o in offs} | set(groups))
            word_x = dict(zip(all_offsets, word.xs(all_offsets)))
            word.export_pdf(word_pdf)
        finally:
            word.close()
        for start, y in pdf_baselines(word_pdf, wboxes).items():
            next(b for b in wboxes if b["start"] == start)["base"] = y
        open_in_raabe(raabe, path)
        ours_text = raabe.text()
        if ours_text != text:
            first = next((i for i, (a, b) in enumerate(zip(text, ours_text)) if a != b), min(len(text), len(ours_text)))
            out[path.name] = {"error": "text differs", "at": first, "word": text[first - 20:first + 20], "ours": ours_text[first - 20:first + 20]}
            continue
        ours = raabe.js("__probe.lineBoxes()")
        our_pdf = tmp / f"{path.stem}-raabe.pdf"
        export_raabe_pdf(raabe, our_pdf)
        for start, y in pdf_baselines(our_pdf, ours).items():
            next(b for b in ours if b["start"] == start)["pdf_base"] = y
        out[path.name] = compare(wboxes, ours, text)
        ours_x = dict(zip(all_offsets, raabe.js(f"__probe.xs({all_offsets})")))
        out[path.name]["horizontal"] = horizontal(word_x, ours_x, groups, text)
    return out
