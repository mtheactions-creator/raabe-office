"""Diagramm in RaabeWrite (Step 7b, D-088), without real input: Word's own document with seven charts
(make_fixtures.py --charts) opened in RaabeWrite, which must draw every chart (no "Diagramme" in the
unsupported list) and keep them when saving (Word opens RaabeWrite's file and finds the same charts);
both apps export a PDF, rendered with Windows' PDF renderer, and the chart areas are compared.
Result: results/<date>-write-charts.json; page pictures in captures/write-charts/ (local only).
"""
from __future__ import annotations

import datetime
import json
import subprocess
import tempfile
import time
from pathlib import Path

import numpy as np
from PIL import Image

from feellab.catalogue import RaabeSide

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
DOC = ROOT / "packages/ooxml/test/fixtures/word-charts.docx"
TMP = Path(tempfile.gettempdir())
CAP = HERE / "captures" / "write-charts"


def raster(pdf: Path, out: Path, dpi: int = 96) -> list[Path]:
    out.mkdir(parents=True, exist_ok=True)
    for f in out.glob("page-*.png"):
        f.unlink()
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(HERE / "feellab" / "pdf_raster.ps1"), str(pdf), str(out), str(dpi)],
                   check=True, capture_output=True)
    return sorted(out.glob("page-*.png"), key=lambda p: int(p.stem.split("-")[1]))


def word_side() -> dict:
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    out: dict = {}
    try:
        d = app.Documents.Open(str(DOC), ReadOnly=True)
        out["charts"] = [int(s.Chart.ChartType) for s in d.InlineShapes if s.HasChart]
        pdf = TMP / "word-charts-word.pdf"
        d.ExportAsFixedFormat(str(pdf), 17)
        d.Close(0)
        out["pdf"] = str(pdf)
        ours = TMP / "word-charts-raabe.docx"
        if ours.exists():
            d = app.Documents.Open(str(ours), ReadOnly=True)
            out["roundtrip_charts"] = [int(s.Chart.ChartType) for s in d.InlineShapes if s.HasChart]
            d.Close(0)
    finally:
        app.Quit(0)
    return out


def main() -> None:
    report: dict = {}
    side = RaabeSide()
    ours_pdf = TMP / "word-charts-raabe.pdf"
    saved = TMP / "word-charts-raabe.docx"
    for f in (ours_pdf, saved):
        if f.exists():
            f.unlink()
    try:
        side.open(DOC)
        t = side.t
        docs = t.js("__app.docs()")
        report["unsupported"] = next((d["unsupported"] for d in docs if d["active"]), None)
        t.js(f"__app.exportPdf({str(ours_pdf)!r})")
        side.save(saved)
        deadline = time.perf_counter() + 30
        while time.perf_counter() < deadline and not (ours_pdf.exists() and saved.exists()):
            time.sleep(0.2)
        time.sleep(1)
    finally:
        side.t.close() if hasattr(side.t, "close") else None
    w = word_side()
    report["word_charts"] = w["charts"]
    report["roundtrip_charts"] = w.get("roundtrip_charts")
    wp = raster(Path(w["pdf"]), CAP / "word")
    rp = raster(ours_pdf, CAP / "raabe") if ours_pdf.exists() else []
    pages = []
    for a, b in zip(wp, rp):
        ia, ib = np.asarray(Image.open(a).convert("L"), dtype=float), np.asarray(Image.open(b).convert("L"), dtype=float)
        h, ww = min(ia.shape[0], ib.shape[0]), min(ia.shape[1], ib.shape[1])
        pages.append(round(float(np.abs(ia[:h, :ww] - ib[:h, :ww]).mean()), 2))
    report["pages_word"], report["pages_raabe"], report["page_diff"] = len(wp), len(rp), pages
    checks = [
        ("no chart reported unsupported", not any("Diagramm" in u for u in (report["unsupported"] or []))),
        ("Word finds the 7 charts in RaabeWrite's file", report["roundtrip_charts"] == report["word_charts"]),
        ("same page count", len(wp) == len(rp)),
    ]
    report["checks"] = [{"check": n, "ok": ok} for n, ok in checks]
    for n, ok in checks:
        print(n, ok)
    print("page differences (mean of 255):", pages)
    (HERE / "results" / f"{datetime.date.today().isoformat()}-write-charts.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


TABLE = {"categories": ["Mo", "Di", "Mi"], "series": [{"name": "Nudeln", "values": [30, 42, 25]}, {"name": "Pizza", "values": [51, 38, 47]}]}


def part2() -> None:
    """Einfügen > Diagramm (Step 7b part 2): RaabeWrite inserts a column chart and a pie whose data are then changed
    (Daten bearbeiten), saves; Word must find both charts with their types and values, and the embedded workbooks."""
    import zipfile
    import win32com.client
    out = TMP / "write-charts-new.docx"
    if out.exists():
        out.unlink()
    side = RaabeSide()
    try:
        side.open(ROOT / "packages/ooxml/test/fixtures/word-formatting.docx")
        t = side.t
        t.js("(() => { const e = __app.editor(); const n = e.doc.paras.length - 1; e.setSel({ p: n, o: e.doc.paras[n].text.length }, { p: n, o: e.doc.paras[n].text.length }); })()")
        t.js('__app.insertChart("column")')
        # The new chart is selected (Word too): typing or inserting would replace it, so the caret goes behind it.
        t.js("(() => { const e = __app.editor(); const [, to] = e.range(); e.setSel(to, to); })()")
        t.js('__app.insertChart("pie")')
        t.js(f"__app.setChartData({json.dumps(TABLE)})")
        print("objects before saving:", t.js("__app.editor().doc.paras.flatMap(p => (p.objects || []).filter(o => o.chart).map(o => !!o.chart.workbook))"))
        t.js(f"(async () => await __app.saveTo({str(out)!r}))()")
        deadline = time.perf_counter() + 20
        while time.perf_counter() < deadline and not out.exists():
            time.sleep(0.2)
        time.sleep(1)
    finally:
        side.t.close()
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
    embedded = [n for n in names if n.startswith("word/embeddings/") and n.endswith(".xlsx")]
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    got = []
    try:
        d = app.Documents.Open(str(out), ReadOnly=True)
        for s in d.InlineShapes:
            if s.HasChart:
                ch = s.Chart
                got.append({"type": int(ch.ChartType), "series": [{"name": ch.SeriesCollection(i).Name, "values": list(ch.SeriesCollection(i).Values)}
                                                                  for i in range(1, ch.SeriesCollection().Count + 1)]})
        d.Close(0)
    finally:
        app.Quit(0)
    checks = [("Word finds 2 new charts", len(got) == 2), ("first is Gruppierte Säulen (51)", bool(got) and got[0]["type"] == 51),
              ("second is Kreis (5)", len(got) > 1 and got[1]["type"] == 5),
              ("pie has the edited values", len(got) > 1 and [round(v) for v in got[1]["series"][0]["values"]] == [30, 42, 25]),
              ("column has Word's sample", bool(got) and [round(v, 1) for v in got[0]["series"][0]["values"]] == [4.3, 2.5, 3.5, 4.5]),
              ("two embedded workbooks", len(embedded) == 2)]
    for n, ok in checks:
        print("part2", n, ok)
    path = HERE / "results" / f"{datetime.date.today().isoformat()}-write-charts.json"
    rep = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    rep["insert"] = {"checks": [{"check": n, "ok": ok} for n, ok in checks], "word": got}
    path.write_text(json.dumps(rep, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))


if __name__ == "__main__":
    import sys
    if "--insert" in sys.argv:
        part2()
    else:
        main()
