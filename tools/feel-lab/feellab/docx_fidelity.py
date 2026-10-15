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

# Word writes optional and non-breaking hyphens as control characters in Range.Text;
# inline pictures and note references are one character each, as objects are in RaabeWrite.
WORD_CHARS = {"\x1f": "\u00ad", "\x1e": "\u2011", "\x01": "\ufffc", "\x02": "\ufffc"}
CELL_MARK = "\r\x07"


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
        """The main text as RaabeWrite's probe shows it: paragraphs (and table cells) joined by
        "\n". Word's end-of-row marks have no counterpart and are left out; section breaks
        end a paragraph. Builds the map between Word's positions and these offsets."""
        # Positions RaabeWrite has no character for: anchors of floating pictures, field codes
        # (begin mark, code, separator, end mark; Content.Text leaves them out), and all but the
        # first position of a field result or equation that RaabeWrite holds as one object.
        hidden: set[int] = {self.doc.Shapes(k).Anchor.Start for k in range(1, self.doc.Shapes.Count + 1)}
        objects: set[int] = set()
        for k in range(1, self.doc.Fields.Count + 1):
            f = self.doc.Fields(k)
            hidden.update(range(f.Code.Start - 1, f.Result.Start))
            hidden.add(f.Result.End)
            # TOC and hyperlinks are markers around ordinary text in RaabeWrite; other fields
            # (PAGEREF, PAGE …) are objects showing their result.
            if f.Type not in (13, 88) and f.Result.End > f.Result.Start:
                objects.add(f.Result.Start)
                hidden.update(range(f.Result.Start + 1, f.Result.End))
        for k in range(1, self.doc.OMaths.Count + 1):
            r = self.doc.OMaths(k).Range
            objects.add(r.Start)
            hidden.update(range(r.Start + 1, r.End))
        hidden -= objects
        self.hidden = sorted(hidden)
        # The text between hidden ranges, position by position.
        end = self.doc.Content.End
        chars: list[tuple[int, str]] = []
        marks: list[int] = []
        pos = 0
        while pos < end:
            if pos in hidden:
                pos += 1
                continue
            if pos in objects:
                chars.append((pos, "\ufffc"))
                pos += 1
                continue
            stop = pos
            while stop < end and stop not in hidden and stop not in objects:
                stop += 1
            seg = self.doc.Range(pos, stop).Text
            i, q = 0, pos
            while i < len(seg) and q < stop:
                if seg.startswith(CELL_MARK, i):
                    marks.append(q)
                    chars.append((q, CELL_MARK))
                    i += 2
                else:
                    chars.append((q, seg[i]))
                    i += 1
                q += 1
            pos = stop
        # End-of-row marks (Information(31): wdAtEndOfRowMarker); the others end cells,
        # including the hidden ones of vertically merged cells.
        self.row_ends = [m for m in marks if self.doc.Range(m, m).Information(31)]
        pictures = {self.doc.InlineShapes(k).Range.Start for k in range(1, self.doc.InlineShapes.Count + 1)}
        sections = {self.doc.Sections(k).Range.End - 1 for k in range(1, self.doc.Sections.Count)}
        row_set = set(self.row_ends)
        out = []
        for pos, ch in chars:
            if ch == CELL_MARK:
                if pos not in row_set:
                    out.append("\n")
            elif ch == "\x0c" and pos in sections:
                out.append("\n")
            elif ch == "\r":
                out.append("\n")
            elif pos in pictures:
                out.append("\ufffc")
            else:
                out.append(WORD_CHARS.get(ch, ch))
        t = "".join(out)
        return t[:-1] if t.endswith("\n") else t

    def skipped(self) -> list[int]:
        if getattr(self, "_skipped", None) is None:
            self._skipped = sorted(set(self.row_ends) | set(self.hidden))
        return self._skipped

    def to_ours(self, pos: int) -> int:
        import bisect
        return pos - bisect.bisect_left(self.skipped(), pos)

    def to_word(self, offset: int) -> int:
        pos = offset
        for r in self.skipped():
            if r <= pos:
                pos += 1
            else:
                break
        return pos

    def needs_positions(self) -> bool:
        """Tables, several sections, columns, fields or equations: line starts found from
        positions (Word's MoveDown moves between table cells and gets stuck in fields)."""
        if self.doc.Tables.Count or self.doc.Sections.Count > 1 or self.doc.Fields.Count or self.doc.OMaths.Count:
            return True
        # Floating objects that text flows around: MoveDown gets stuck beside them.
        if any(self.doc.Shapes(k).WrapFormat.Type in (0, 1, 2) for k in range(1, self.doc.Shapes.Count + 1)):
            return True
        return any(self.doc.Sections(k).PageSetup.TextColumns.Count > 1 for k in range(1, self.doc.Sections.Count + 1))

    def lines_by_position(self) -> list[int]:
        """Line starts: within each paragraph, where the line's page or vertical position
        changes (binary search over positions)."""
        def key(pos: int):
            r = self.doc.Range(pos, pos)
            return int(r.Information(3)), round(float(r.Information(6)), 1)
        starts: list[int] = []
        row_set = set(self.row_ends)
        # Only positions RaabeWrite has a character for (not field codes, not the inside of an
        # equation, whose parts sit at different heights).
        hidden = set(getattr(self, "hidden", []))
        for para in self.doc.Paragraphs:
            a, e = para.Range.Start, para.Range.End - 1
            if a in row_set:
                continue
            try:
                self.win.GetPoint(0, 0, 0, 0, self.doc.Range(a, a))
            except Exception:
                continue  # not shown (the text of a vertically merged cell)
            vis = [q for q in range(a, e + 1) if q not in hidden] or [a]
            i = 0
            starts.append(vis[0] if a in hidden else a)
            k = key(vis[0])
            last = len(vis) - 1
            while i < last and key(vis[last]) != k:
                lo, hi = i, last  # key(vis[lo]) == k, key(vis[hi]) != k
                while hi - lo > 1:
                    mid = (lo + hi) // 2
                    if key(vis[mid]) == k:
                        lo = mid
                    else:
                        hi = mid
                i = hi
                starts.append(vis[i])
                k = key(vis[i])
        return self.split_pieces(starts) if self.doc.Shapes.Count else starts

    def split_pieces(self, starts: list[int]) -> list[int]:
        """Rows that a picture divides (text on both sides): a new line where x jumps forward
        across the picture (Information(5), screen pixels; jumps of over 40 pt)."""
        out: list[int] = []
        ends = starts[1:] + [self.doc.Content.End - 1]
        jump = 40 * self.app.PointsToPixels(1)
        for a, e in zip(starts, ends):
            out.append(a)
            if e - a < 3:
                continue
            prev = None
            for q in range(a, e):
                r = self.doc.Range(q, q)
                x = float(r.Information(5))
                if prev is not None and x - prev > jump and self.doc.Range(q - 1, q).Text != "\t" and not self.doc.Range(q - 1, q).InlineShapes.Count:
                    out.append(q)
                prev = x
        return out

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
            while time.perf_counter() < deadline and not target.js("__probe.settled() && !window.__mathPending"):
                time.sleep(0.1)
            time.sleep(0.2)
            return
        time.sleep(0.1)
    raise TimeoutError(f"{path.name} did not open in RaabeWrite")


def pdf_baselines(path: Path, boxes: list[dict]) -> dict[int, float]:
    """Baseline of each line (by start offset) from a PDF. Every text run goes to the line
    box it lies in: vertically inside the box, and the box starting nearest to its left
    (table cells and columns side by side). Text outside the boxes (headers, footers,
    footnotes, list labels left of the text) is ignored."""
    pages = pdftext.read(path)
    out: dict[int, float] = {}
    for page, pdf_page in enumerate(pages):
        lines = [b for b in boxes if b["page"] == page]
        got: dict[int, list] = {}
        for r in pdf_page.runs:
            if not r.chars or r.size <= 0:
                continue
            cands = [b for b in lines if b["y"] - 0.5 <= r.y <= b["y"] + b["h"] + 0.5 and b["x"] <= r.x + 1.5]
            if not cands:
                continue
            best = max(cands, key=lambda b: (round(b["x"], 0), -abs(b["y"] + 0.8 * b["h"] - r.y)))
            got.setdefault(best["start"], []).append(r)
        for start, runs in got.items():
            # The baseline most characters sit on (raised or lowered characters do not count).
            groups: dict[float, int] = {}
            for r in runs:
                key = next((k for k in groups if abs(k - r.y) <= 0.3), r.y)
                groups[key] = groups.get(key, 0) + max(1, int(r.chars) if not isinstance(r.chars, str) else len(r.chars))
            best = max(groups, key=lambda k: groups[k])
            main = [r for r in runs if abs(r.y - best) <= 0.3]
            out[start] = round(sum(r.y for r in main) / len(main), 2)
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
            wstarts = word.lines_by_position() if word.needs_positions() else word.lines()
            wboxes = word.boxes(wstarts)
            for b in wboxes:
                b["start"] = word.to_ours(b["start"])
            starts = [b["start"] for b in wboxes]
            groups = word_offsets(starts, text)
            all_offsets = sorted({o for offs in groups.values() for o in offs} | set(groups))
            word_x = dict(zip(all_offsets, word.xs([word.to_word(o) for o in all_offsets])))
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
