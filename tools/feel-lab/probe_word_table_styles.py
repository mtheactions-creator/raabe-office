"""Probe: Word's built-in table styles (Step 4c D). Every table style Word knows gets a small
table in one document of our own invisible Word instance; the saved file gives each style's
definition as Word writes it (w:style with its conditional formats), its German name, id and
English name. Result: results/<date>-word-table-styles.json (name → {local, id, xml}).

    .venv/Scripts/python.exe probe_word_table_styles.py
"""
from __future__ import annotations

import datetime
import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-word-table-styles.json"
TMP = Path(tempfile.gettempdir()) / "raabe-table-styles.docx"


def main() -> None:
    import win32com.client
    sys.stdout.reconfigure(encoding="utf-8")
    app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
    app.Visible = False
    order: list[str] = []
    try:
        d = app.Documents.Add()
        names = [s.NameLocal for s in d.Styles if s.Type == 3]
        print(len(names), "table styles", flush=True)
        for n in names:
            r = d.Range(d.Content.End - 1, d.Content.End - 1)
            r.InsertParagraphAfter()
            r = d.Range(d.Content.End - 1, d.Content.End - 1)
            t = d.Tables.Add(r, 2, 2)
            try:
                t.Style = n
            except Exception as e:  # noqa: BLE001
                print("cannot apply", n, e, flush=True)
            order.append(n)
        d.SaveAs2(str(TMP), 16)
        d.Close(0)
    finally:
        app.Quit(0)
    with zipfile.ZipFile(TMP) as z:
        styles = z.read("word/styles.xml").decode("utf-8")
        doc = z.read("word/document.xml").decode("utf-8")
    used = re.findall(r'<w:tblStyle w:val="([^"]+)"/>', doc)
    defs = {}
    for m in re.finditer(r'<w:style w:type="table"[^>]*?w:styleId="([^"]+)"[^>]*>([\s\S]*?)</w:style>', styles):
        sid = m.group(1)
        name = re.search(r'<w:name w:val="([^"]+)"/>', m.group(0))
        defs[sid] = {"name": name.group(1) if name else sid, "xml": m.group(0)}
    out = {}
    # The document's tables in order: each table's style id; the local names in the same order
    # (Normale Tabelle is the default and gets no w:tblStyle).
    k = 0
    for local in order:
        if local in ("Normale Tabelle", "Table Normal"):
            continue
        if k >= len(used):
            break
        sid = used[k]
        k += 1
        dd = defs.get(sid)
        if dd:
            out[dd["name"]] = {"local": local, "id": sid, "xml": dd["xml"]}
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(len(out), "styles recorded", flush=True)
    for k2 in list(out)[:8]:
        print(k2, "->", out[k2]["local"], out[k2]["id"])


if __name__ == "__main__":
    main()
