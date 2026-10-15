"""Probe: what Word puts on the clipboard for equations (Step 4c C) and which clipboard data
Word turns back into an equation when pasting. Own invisible Word instance; uses the system
clipboard, no other input. Writes every format's content to captures/word-clipboard/ (local
only) and a summary to results/<date>-word-math-clipboard.json.

    .venv/Scripts/python.exe probe_word_math_clipboard.py copy
    .venv/Scripts/python.exe probe_word_math_clipboard.py paste <file> <format>   # e.g. "HTML Format"
"""
from __future__ import annotations

import datetime
import json
import sys
import time
from pathlib import Path

import win32clipboard

HERE = Path(__file__).parent
CAP = HERE / "captures" / "word-clipboard"
CAP.mkdir(parents=True, exist_ok=True)
OUT = HERE / "results" / f"{datetime.date.today()}-word-math-clipboard.json"

STANDARD = {1: "CF_TEXT", 2: "CF_BITMAP", 3: "CF_METAFILEPICT", 7: "CF_OEMTEXT", 8: "CF_DIB", 13: "CF_UNICODETEXT", 14: "CF_ENHMETAFILE",
            16: "CF_LOCALE", 17: "CF_DIBV5"}


def open_clip() -> None:
    for _ in range(40):
        try:
            win32clipboard.OpenClipboard()
            return
        except Exception:  # noqa: BLE001
            time.sleep(0.1)
    raise RuntimeError("clipboard busy")


def formats() -> dict[str, bytes | str | None]:
    open_clip()
    out: dict[str, bytes | str | None] = {}
    try:
        f = 0
        while True:
            f = win32clipboard.EnumClipboardFormats(f)
            if not f:
                break
            name = STANDARD.get(f) or (win32clipboard.GetClipboardFormatName(f) if f >= 0xC000 else f"#{f}")
            try:
                data = win32clipboard.GetClipboardData(f) if f not in (2, 3, 8, 14, 17) else None
            except Exception:  # noqa: BLE001
                data = None
            out[name] = data
    finally:
        win32clipboard.CloseClipboard()
    return out


def word():
    import win32com.client
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    app.Visible = False
    return app


def copy() -> None:
    app = word()
    summary: dict = {}
    try:
        d = app.Documents.Add()
        d.Content.Text = "Vor der Formel  danach.\rx=(-b±√(b^2-4ac))/2a\r"
        # An inline equation inside text, a display equation on its own line.
        r = d.Range(15, 15)
        r.Text = "a/b+x^2"
        d.OMaths.Add(r).OMaths(1).BuildUp()
        p2 = d.Paragraphs(2).Range
        rr = d.Range(p2.Start, p2.End - 1)
        d.OMaths.Add(rr).OMaths(1).BuildUp()
        time.sleep(0.3)
        # Plain text and a function name in an equation of its own.
        p3 = d.Paragraphs(3).Range
        r3 = d.Range(p3.Start, p3.Start)
        r3.Text = 'sin x="wenn" a'
        d.OMaths.Add(r3).OMaths(1).BuildUp()
        cases = {
            "plain runs": d.OMaths(3).Range,
            "inline equation": d.OMaths(1).Range,
            "display equation": d.OMaths(2).Range,
            "text with equation": d.Paragraphs(1).Range,
        }
        for name, rg in cases.items():
            rg.Copy()
            time.sleep(0.5)
            fm = formats()
            info = {}
            for k, v in fm.items():
                if v is None:
                    info[k] = None
                    continue
                text = v.decode("utf-8", "replace") if isinstance(v, bytes) else str(v)
                safe = "".join(ch if ch.isalnum() else "_" for ch in f"{name}-{k}")
                (CAP / f"math-{safe}.txt").write_text(text, encoding="utf-8", newline="\n")
                info[k] = len(text)
            summary[name] = info
            print(name, json.dumps(info, ensure_ascii=False), flush=True)
        d.Close(0)
    finally:
        app.Quit(0)
    OUT.write_text(json.dumps({"copy": summary}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


def envelope(html: str) -> bytes:
    """A CF_HTML block for an HTML document (as RaabeWrite writes it)."""
    start, end = "<!--StartFragment-->", "<!--EndFragment-->"
    doc = html if start in html else f"<html><body>{start}{html}{end}</body></html>"
    raw = doc.encode("utf-8")
    nl = chr(13) + chr(10)
    keys = ["Version:0.9", "StartHTML:{:010}", "EndHTML:{:010}", "StartFragment:{:010}", "EndFragment:{:010}"]
    n = len(nl.join(k.format(0) for k in keys) + nl)
    sf = n + raw.find(start.encode()) + len(start)
    ef = n + raw.find(end.encode())
    head = nl.join(k.format(v) for k, v in zip(keys, [0, n, n + len(raw), sf, ef])) + nl
    return head.encode() + raw


def paste(file: str, fmt: str) -> None:
    data = Path(file).read_bytes()
    if fmt == "HTML Format" and not data.startswith(b"Version:"):
        data = envelope(data.decode("utf-8"))
    open_clip()
    try:
        win32clipboard.EmptyClipboard()
        f = win32clipboard.RegisterClipboardFormat(fmt) if fmt not in ("CF_UNICODETEXT",) else 13
        win32clipboard.SetClipboardData(f, data if f != 13 else data.decode("utf-8"))
    finally:
        win32clipboard.CloseClipboard()
    app = word()
    try:
        d = app.Documents.Add()
        d.Content.Text = "Start \r"
        d.Range(6, 6).Paste()
        time.sleep(0.3)
        print("equations:", d.OMaths.Count)
        for i in range(1, d.OMaths.Count + 1):
            m = d.OMaths(i)
            print("  ", i, repr(m.Range.Text), "display" if m.Type == 0 else "inline")
        print("text:", repr(d.Content.Text))
        d.Close(0)
    finally:
        app.Quit(0)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if sys.argv[1] == "copy":
        copy()
    else:
        paste(sys.argv[2], sys.argv[3])
