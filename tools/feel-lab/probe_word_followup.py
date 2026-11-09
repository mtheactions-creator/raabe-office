"""Records in Word (own invisible instance, DispatchEx) what the by-hand audit's open RaabeWrite options do (D-149):

- Kein Leerraum on the audit's target: the font the text has afterwards, and how Word writes the style.
- Spalte links / rechts einfügen with the caret in a cell: the column widths and the table's width afterwards.
- Diese Änderung annehmen / ablehnen, Zurück and Weiter when the selection holds no change: what Word selects and
  whether a change is accepted or rejected. (The split buttons' main part, Annehmen/Ablehnen und zur nächsten
  wechseln, cannot be run through ExecuteMso.)
- The command Word binds to each of RaabeWrite's shortcuts (FindKey): `--keys`.

    .venv/Scripts/python.exe probe_word_followup.py      # results/<date>-word-followup.json

The base document is the audit's own (sweep_write.py, built by Word in TEMP). Nothing else is opened or changed."""
from __future__ import annotations

import datetime
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from feellab import word_summary as ws

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir()) / os.environ.get("SWEEP_DIR", "raabe-sweep-write")
BASE = TMP / "base.docx"
OUT = TMP / "followup"
DATE = os.environ.get("SWEEP_DATE") or datetime.date.today().isoformat()


def word_pids() -> set[int]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"], capture_output=True, text=True, encoding="cp850", errors="replace").stdout
    return {int(m.group(1)) for m in re.finditer(r'"WINWORD.EXE","(\d+)"', out)}


def copy(name: str) -> Path:
    OUT.mkdir(exist_ok=True)
    p = OUT / f"{name}.docx"
    shutil.copy(BASE, p)
    return p


def sel_info(app, d) -> dict:
    s = app.Selection
    return {"start": s.Start, "end": s.End, "text": s.Text[:60], "revisions": d.Revisions.Count,
            "revs": [{"type": r.Type, "text": r.Range.Text[:30]} for r in d.Revisions]}


def find(d, text: str):
    r = d.Content
    r.Find.Execute(FindText=text, MatchCase=True)
    return r


def no_spacing(app) -> dict:
    p = copy("nospacing")
    d = app.Documents.Open(str(p), AddToRecentFiles=False, Visible=False)
    try:
        r = find(d, ws.TARGET)
        before = r.Font.Name
        r.Select()
        app.Selection.Range.Style = "Kein Leerraum"
        after = {"target_font": app.Selection.Font.Name, "style_font": d.Styles("Kein Leerraum").Font.Name,
                 "style_base": d.Styles("Kein Leerraum").BaseStyle.NameLocal if d.Styles("Kein Leerraum").BaseStyle else ""}
        d.Save()
    finally:
        d.Close(0)
    z = zipfile.ZipFile(p)
    st = z.read("word/styles.xml").decode("utf-8")
    m = re.search(r'<w:style [^>]*w:styleId="KeinLeerraum".*?</w:style>', st, re.S)
    return {"target_font_before": before, **after, "style_xml": m.group(0) if m else None}


def table_cols(d) -> dict:
    t = d.Tables(1)
    widths = [round(t.Cell(1, j).Width, 1) for j in range(1, t.Columns.Count + 1)]
    return {"cols": t.Columns.Count, "widths": widths, "total": round(sum(widths), 1), "pref": round(t.PreferredWidth, 1), "prefType": t.PreferredWidthType,
            "autofit": t.AllowAutoFit}


def insert_col(app, method: str) -> dict:
    p = copy(f"col-{method}")
    d = app.Documents.Open(str(p), AddToRecentFiles=False, Visible=False)
    try:
        before = table_cols(d)
        r = d.Tables(1).Range
        r.Find.Execute(FindText="Sonnenblume", MatchCase=True)
        r.SetRange(r.Start + 2, r.Start + 2)
        r.Select()
        getattr(app.Selection, method)()
        after = table_cols(d)
        d.Save()
    finally:
        d.Close(0)
    z = zipfile.ZipFile(p)
    doc = z.read("word/document.xml").decode("utf-8")
    grid = re.search(r"<w:tblGrid>.*?</w:tblGrid>", doc, re.S)
    tblw = re.search(r"<w:tblW [^>]*/>", doc)
    return {"before": before, "after": after, "tblGrid": grid.group(0) if grid else None, "tblW": tblw.group(0) if tblw else None}


def revisions_doc(app, name: str):
    p = copy(name)
    d = app.Documents.Open(str(p), AddToRecentFiles=False, Visible=False)
    d.Windows(1).View.Type = 3
    r = find(d, "Thema")
    d.Comments.Add(r, "Bitte prüfen")
    d.TrackRevisions = True
    r = find(d, "Schluss.")
    r.Collapse(1)
    r.InsertBefore("neu ")
    r = find(d, "Ende")
    r.Delete()
    d.TrackRevisions = False
    find(d, ws.TARGET).Select()
    return d


def review(app) -> dict:
    out = {}
    # The split buttons' own ids run their main part (Annehmen und zur nächsten wechseln / Ablehnen und zur nächsten
    # wechseln); ReviewAcceptChange / ReviewRejectChange are the menus' "Diese Änderung annehmen/ablehnen".
    names = ["ReviewAcceptChangeMenu", "ReviewRejectChangeMenu", "ReviewAcceptChange", "ReviewRejectChange", "ReviewNextChange", "ReviewPreviousChange"]
    for name in names:
        try:
            label = app.CommandBars.GetLabelMso(name)
        except Exception as e:  # noqa: BLE001
            out[name] = {"error": str(e)[:120]}
            continue
        d = revisions_doc(app, name)
        try:
            before = sel_info(app, d)
            try:
                app.CommandBars.ExecuteMso(name)
                err = None
            except Exception as e:  # noqa: BLE001
                err = str(e)[:160]
            after = sel_info(app, d)
            # A second press from there (Word's Annehmen und weiter: the change now selected is accepted).
            try:
                app.CommandBars.ExecuteMso(name)
            except Exception:  # noqa: BLE001
                pass
            again = sel_info(app, d)
        finally:
            d.Close(0)
        out[name] = {"label": label, "error": err, "before": before, "after": after, "again": again}
        print(name, json.dumps(out[name], ensure_ascii=False)[:400], flush=True)
    return out


# RaabeWrite's shortcuts (command-list.ts) as Word key codes: Word binds keys, not characters (WdKey: letters and digits
# by their code, 13 Enter, 32 space, 118 F7, 187 the "+" key, 191 "#", 226 "<").
SHORTCUTS = {
    "Strg+N": "N", "Strg+O": "O", "Strg+S": "S", "F12": 123, "Strg+P": "P", "Strg+W": "W", "Strg+Z": "Z", "Strg+Y": "Y", "Strg+V": "V",
    "Strg+Umschalt+V": "V", "Strg+X": "X", "Strg+C": "C", "Strg+Alt+C": "C", "Strg+Alt+V": "V", "Strg+A": "A", "Strg+F": "F", "Strg+H": "H",
    "Strg+G": "G", "Strg+Umschalt+A": "A", "Strg+Umschalt+P": "P", "Strg+Umschalt+<": 226, "Strg+<": 226, "Strg+Umschalt+F": "F",
    "Strg+Umschalt+K": "K", "Strg+Umschalt+U": "U", "Strg+#": 191, "Strg+D": "D", "Strg+Umschalt+M": "M", "Strg+M": "M",
    "Strg+Umschalt+8": "8", "Strg+Umschalt++": 187, "Strg+L": "L", "Strg+E": "E", "Strg+R": "R", "Strg+B": "B", "Strg+Umschalt+D": "D",
    "Strg+Umschalt+G": "G", "Strg+1": "1", "Strg+5": "5", "Strg+2": "2", "Strg+Alt+1": "1", "Strg+Alt+2": "2", "Strg+Alt+3": "3",
    "Strg+Alt+Umschalt+S": "S", "Strg+Eingabe": 13, "Umschalt+Eingabe": 13, "Strg+K": "K", "Strg+Alt+K": "K", "Alt+Umschalt+0": "0",
    "Strg+Alt+F": "F", "Strg+Alt+D": "D", "F7": 118, "Umschalt+F7": 118, "Strg+Alt+Leertaste": 32, "Strg+Umschalt+E": "E",
    "Strg+0": "0", "Strg+F1": 112,
    # Where German Word has its heading shortcuts (Strg+Alt+1 is free there: Strg+Alt is AltGr).
    "Alt+1": "1", "Alt+2": "2", "Alt+3": "3", "Alt+Umschalt+1": "1", "Strg+Umschalt+1": "1", "Strg+Umschalt+N": "N", "Alt+Umschalt+3": "3",
}


def key_bindings(app) -> dict:
    """The Word command each shortcut runs (Application.FindKey)."""
    out = {}
    for combo, key in SHORTCUTS.items():
        mods = combo.split("+")[:-1] if not combo.endswith("++") else combo[:-2].split("+")
        code = key if isinstance(key, int) else ord(key)
        args = [code] + [{"Strg": 512, "Umschalt": 256, "Alt": 1024}[m] for m in mods]
        try:
            kc = app.BuildKeyCode(*args)
            out[combo] = app.FindKey(kc).Command or ""
        except Exception as e:  # noqa: BLE001
            out[combo] = f"ERR {str(e)[:80]}"
    return out


def main() -> None:
    if "--keys" in sys.argv:
        out = HERE / "results" / f"{DATE}-word-followup.json"
        res = json.loads(out.read_text(encoding="utf-8"))
        app = ws.word_app()
        try:
            d = app.Documents.Add()
            res["key_bindings"] = key_bindings(app)
            d.Close(0)
        finally:
            app.Quit(0)
        print(json.dumps(res["key_bindings"], ensure_ascii=False, indent=0))
        out.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
        return
    if not BASE.exists():
        sys.exit(f"no base document {BASE} (run sweep_write.py sweep first)")
    before = word_pids()
    app = ws.word_app()
    mine = word_pids() - before
    print("own Word", mine, flush=True)
    res = {"date": DATE}
    try:
        res["no_spacing"] = no_spacing(app)
        print(json.dumps(res["no_spacing"], ensure_ascii=False)[:600], flush=True)
        res["insert_col_left"] = insert_col(app, "InsertColumns")
        res["insert_col_right"] = insert_col(app, "InsertColumnsRight")
        print(json.dumps(res["insert_col_left"])[:600], flush=True)
        print(json.dumps(res["insert_col_right"])[:600], flush=True)
        res["review"] = review(app)
    finally:
        app.Quit(0)
    out = HERE / "results" / f"{DATE}-word-followup.json"
    out.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("->", out.name)


if __name__ == "__main__":
    main()
