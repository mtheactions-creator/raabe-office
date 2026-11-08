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


if __name__ == "__main__":
    main()
