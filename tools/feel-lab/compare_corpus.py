"""The look corpus (D-080 D): Office's own school-like documents (make_corpus_*.py, in TEMP/raabe-corpus) opened
by Office and by RaabeX, both exported as PDF, rendered by Windows' PDF renderer at 96 dpi and compared page by
page: page count and mean grey difference (0-255) per page, the worst pages listed. No real input.

    .venv/Scripts/python.exe compare_corpus.py excel [name ...]
    .venv/Scripts/python.exe compare_corpus.py word [name ...]
    .venv/Scripts/python.exe compare_corpus.py ppt [name ...]

Result: results/<date>-corpus-<app>.json; page pictures and difference pictures in captures/corpus/ (local only).
"""
from __future__ import annotations

import datetime
import json
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import numpy as np
from PIL import Image

HERE = Path(__file__).parent
TMP = Path(tempfile.gettempdir())
CORPUS = TMP / "raabe-corpus"
CAP = HERE / "captures" / "corpus"


def raster(pdf: Path, out: Path, dpi: int = 96) -> list[Path]:
    out.mkdir(parents=True, exist_ok=True)
    for f in out.glob("page-*.png"):
        f.unlink()
    if not pdf.exists():
        return []
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(HERE / "feellab" / "pdf_raster.ps1"), str(pdf), str(out), str(dpi)],
                   check=True, capture_output=True)
    return sorted(out.glob("page-*.png"), key=lambda p: int(p.stem.split("-")[1]))


def coarse(a: Path, b: Path, k: int = 6) -> float:
    """The difference after both pages are averaged over k x k pixel blocks: a line set a pixel apart counts little,
    a missing picture, another line break or another column width counts fully."""
    ia = np.asarray(Image.open(a).convert("L"), dtype=float)
    ib = np.asarray(Image.open(b).convert("L"), dtype=float)
    h, w = min(ia.shape[0], ib.shape[0]) // k * k, min(ia.shape[1], ib.shape[1]) // k * k
    ra = ia[:h, :w].reshape(h // k, k, w // k, k).mean(axis=(1, 3))
    rb = ib[:h, :w].reshape(h // k, k, w // k, k).mean(axis=(1, 3))
    return round(float(np.abs(ra - rb).mean()), 2)


def diff(a: Path, b: Path, save: Path | None = None) -> float:
    ia = np.asarray(Image.open(a).convert("L"), dtype=float)
    ib = np.asarray(Image.open(b).convert("L"), dtype=float)
    h, w = max(ia.shape[0], ib.shape[0]), max(ia.shape[1], ib.shape[1])
    pa = np.full((h, w), 255.0)
    pb = np.full((h, w), 255.0)
    pa[:ia.shape[0], :ia.shape[1]] = ia
    pb[:ib.shape[0], :ib.shape[1]] = ib
    d = np.abs(pa - pb)
    if save is not None:
        img = np.stack([pa, pb, np.maximum(pa, pb)], axis=-1)  # cyan: ink only in Office's page, magenta: ink only in ours
        Image.fromarray(img.astype(np.uint8)).save(save)
    return round(float(d.mean()), 2)


def office_excel(files: list[Path]) -> dict[str, Path]:
    from feellab.excel import ExcelTarget
    out = {}
    xl = ExcelTarget(visible=False)
    try:
        for f in files:
            wb = xl.app.Workbooks.Open(str(f), ReadOnly=True)
            try:
                pdf = TMP / "raabe-corpus" / "pdf-office" / f"{f.stem}.pdf"
                pdf.parent.mkdir(parents=True, exist_ok=True)
                wb.Worksheets(1).ExportAsFixedFormat(0, str(pdf))
                out[f.stem] = pdf
            finally:
                wb.Close(False)
    finally:
        xl.close()
    return out


def raabe_calc(files: list[Path]) -> dict[str, Path]:
    from feellab.calc_catalogue import CalcSide
    import os
    out = {}
    # Another harness may hold the default DevTools port: RAABE_CALC_PORT picks another one.
    c = CalcSide(int(os.environ.get("RAABE_CALC_PORT", "9381")))
    try:
        for f in files:
            pdf = TMP / "raabe-corpus" / "pdf-raabe" / f"{f.stem}.pdf"
            pdf.parent.mkdir(parents=True, exist_ok=True)
            if pdf.exists():
                pdf.unlink()
            c.js(f"__calc.open({json.dumps(str(f))})")
            time.sleep(0.8)
            c.js(f"(async () => {{ const app = __calc.app; app.wb.active = 0; await __calc.pdf({json.dumps(str(pdf))}); }})()")
            deadline = time.perf_counter() + 20
            while time.perf_counter() < deadline and not pdf.exists():
                time.sleep(0.2)
            time.sleep(0.3)
            out[f.stem] = pdf
    finally:
        c.close()
    return out


def office_word(files: list[Path]) -> dict[str, Path]:
    import win32com.client
    out = {}
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    try:
        for f in files:
            d = app.Documents.Open(str(f), ReadOnly=True)
            try:
                pdf = TMP / "raabe-corpus" / "pdf-office" / f"{f.stem}.pdf"
                pdf.parent.mkdir(parents=True, exist_ok=True)
                d.ExportAsFixedFormat(str(pdf), 17)
                out[f.stem] = pdf
            finally:
                d.Close(0)
    finally:
        app.Quit(0)
    return out


def raabe_write(files: list[Path]) -> dict[str, Path]:
    from feellab.catalogue import RaabeSide
    out = {}
    side = RaabeSide()
    try:
        for f in files:
            pdf = TMP / "raabe-corpus" / "pdf-raabe" / f"{f.stem}.pdf"
            pdf.parent.mkdir(parents=True, exist_ok=True)
            if pdf.exists():
                pdf.unlink()
            side.open(f)
            side.t.js(f"__app.exportPdf({json.dumps(str(pdf))})")
            deadline = time.perf_counter() + 30
            while time.perf_counter() < deadline and not pdf.exists():
                time.sleep(0.2)
            time.sleep(0.5)
            out[f.stem] = pdf
    finally:
        side.t.close()
    return out


def main() -> None:
    app = sys.argv[1] if len(sys.argv) > 1 else "excel"
    names = sys.argv[2:]
    folder = CORPUS / app
    ext = {"excel": ".xlsx", "word": ".docx", "ppt": ".pptx"}[app]
    files = sorted(p for p in folder.glob(f"*{ext}") if not p.name.startswith("_") and (not names or p.stem in names))
    if app == "excel":
        office, ours = office_excel(files), raabe_calc(files)
    elif app == "word":
        office, ours = office_word(files), raabe_write(files)
    else:
        raise SystemExit("ppt: RaabePresent's slide renders against PowerPoint's PNGs (beside each .pptx) come with Present's render check")
    report = {"app": app, "documents": []}
    for f in files:
        wp = raster(office.get(f.stem, Path("-")), CAP / app / f.stem / "office")
        rp = raster(ours.get(f.stem, Path("-")), CAP / app / f.stem / "raabe")
        pages = [diff(a, b, CAP / app / f.stem / f"diff-{i + 1}.png") for i, (a, b) in enumerate(zip(wp, rp))]
        blocks = [coarse(a, b) for a, b in zip(wp, rp)]
        report["documents"].append({"name": f.stem, "pages_office": len(wp), "pages_raabe": len(rp), "diff": pages, "coarse": blocks,
                                    "mean": round(sum(pages) / len(pages), 2) if pages else None,
                                    "mean_coarse": round(sum(blocks) / len(blocks), 2) if blocks else None})
        print(f"{f.stem:20s} pages {len(wp)}/{len(rp)} diff {pages} coarse {blocks}", flush=True)
    docs = [d for d in report["documents"] if d["mean"] is not None]
    report["summary"] = {"documents": len(files), "same_page_count": sum(1 for d in report["documents"] if d["pages_office"] == d["pages_raabe"] and d["pages_office"]),
                         "mean_diff": round(sum(d["mean"] for d in docs) / len(docs), 2) if docs else None,
                         "mean_coarse": round(sum(d["mean_coarse"] for d in docs) / len(docs), 2) if docs else None,
                         "worst": sorted(((d["mean_coarse"], d["name"]) for d in docs), reverse=True)[:8]}
    print(report["summary"])
    (HERE / "results" / f"{datetime.date.today().isoformat()}-corpus-{app}.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
