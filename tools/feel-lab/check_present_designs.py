"""RaabePresent's own designs (D-190 to D-199): every design and its four Varianten made by RaabePresent (headless
Edge on a dev server, invisible), with a title, a content, a section and a table-and-chart slide. Then:

1. A contact sheet of all designs drawn by our renderer (TEMP/raabe-designs/sheet-*.png), to judge them by eye.
2. If no POWERPNT.EXE runs (else that is the owner's: skipped and said so), each design's file (Variante 1 and 4)
   opened by an own invisible PowerPoint (feellab/powerpoint.py): does it open without repair, what PowerPoint reads
   (theme name, colour scheme, fonts, layouts, shape counts) and PowerPoint's own picture of each slide
   (TEMP/raabe-designs/ppt-sheet-*.png beside ours).

    .venv/Scripts/python.exe check_present_designs.py [--url http://localhost:5243/] [--no-powerpoint] [name ...]

Result: results/<date>-present-designs.json (names and numbers only; files and pictures stay in TEMP).
"""
from __future__ import annotations

import base64
import datetime
import io
import json
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

import feellab.present_page as pp
from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget, owner_powerpoint_running

HERE = Path(__file__).parent
OUT = Path(tempfile.gettempdir()) / "raabe-designs"
THUMB = 320

BUILD = """async (i, k) => {
  const p = window.__present, a = p.app, D = await import('/src/designs.ts');
  const d = D.DESIGNS[i], v = D.variantOf(d, k);
  a.newDoc(v);
  const ed = a.editor;
  const type = (id, text) => { ed.enterText(id); ed.typeText(text); ed.endText(); };
  let st = p.state();
  type(st.shapes[0].id, d.name === 'Office' ? 'Office-Design' : d.name);
  type(st.shapes[1].id, 'Raabeschule · Klasse 10b · ' + (k + 1) + '. Variante');
  const lay = (t) => a.deck.layouts().find((l) => l.type === t).path;
  a.newSlide(lay('obj')); st = p.state();
  type(st.shapes[0].id, 'Worum es heute geht');
  ed.enterText(st.shapes[1].id); ed.typeText('Erster Gedanke mit einem längeren Satz\\nZweiter Gedanke\\nDritter Gedanke'); ed.changeLevel(1); ed.typeText('\\nEin Beispiel dazu'); ed.endText();
  a.newSlide(lay('secHead')); st = p.state();
  type(st.shapes[0].id, 'Zweiter Teil'); type(st.shapes[1].id, 'Was wir herausgefunden haben');
  a.newSlide(lay('titleOnly')); st = p.state();
  type(st.shapes[0].id, 'Ergebnisse in Zahlen');
  const C = await import('/src/charts.ts'); C.insertChart(a, 'column');
  const S = await import('/src/shapeops.ts');
  a.editor.edit('Größe', () => { const doc = a.deck.pptx.doc(a.slidePath()); const fr = S.shapeEls(S.treeOf(doc)).find((e) => e.localName === 'graphicFrame');
    const off = fr.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'off')[0], ext = fr.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'ext')[0];
    off.setAttribute('x', '6400000'); ext.setAttribute('cx', '4953000'); });
  a.insertTable(4, 3);
  a.editor.edit('Größe', () => { const doc = a.deck.pptx.doc(a.slidePath()); const fr = S.shapeEls(S.treeOf(doc)).filter((e) => e.localName === 'graphicFrame').at(-1);
    const off = fr.getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'xfrm')[0]?.firstElementChild ?? fr.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'off')[0];
    off.setAttribute('x', '838200'); off.setAttribute('y', '2000000'); off.nextElementSibling?.setAttribute('cx', '5100000');
    for (const g of fr.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'gridCol')) g.setAttribute('w', '1700000'); });
  await new Promise((r) => setTimeout(r, 150));
  return { name: d.name, group: d.group, slides: a.deck.slides.length, bytes: p.bytes() };
}"""


def label_font(size: int):
    for f in ("segoeui.ttf", "arial.ttf"):
        try:
            return ImageFont.truetype(f, size)
        except OSError:
            pass
    return ImageFont.load_default()


def sheet(rows: list[tuple[str, list[Image.Image]]], path: Path, cols: int) -> None:
    w, h = THUMB, THUMB * 9 // 16
    pad, head = 8, 26
    img = Image.new("RGB", (pad + cols * (w + pad), len(rows) * (h + head + pad) + pad), "#dcdcdc")
    d = ImageDraw.Draw(img)
    f = label_font(15)
    for r, (name, pics) in enumerate(rows):
        y = pad + r * (h + head + pad)
        d.text((pad, y + 3), name, fill="#202020", font=f)
        for c, pic in enumerate(pics[:cols]):
            img.paste(pic.resize((w, h)), (pad + c * (w + pad), y + head))
    img.save(path)


def main() -> None:
    args = [a for a in sys.argv[1:]]
    url = "http://localhost:5243/"
    if "--url" in args:
        url = args[args.index("--url") + 1]
        del args[args.index("--url"):args.index("--url") + 2]
    use_ppt = "--no-powerpoint" not in args
    names = [a for a in args if not a.startswith("--")]
    OUT.mkdir(exist_ok=True)
    # Our own page: own port, profile and dev server, so other sessions' pages are left alone.
    pp.URL, pp.PORT, pp.PROFILE = url, 9431, Path(tempfile.gettempdir()) / "raabe-designs-page-profile"
    page = pp.PresentPage(width=1600, height=1000)
    count = page.js("import('/src/designs.ts').then((D) => D.DESIGNS.length)")
    designs = page.js("import('/src/designs.ts').then((D) => D.DESIGNS.map((d) => d.name))")
    result: dict = {"date": datetime.date.today().isoformat(), "designs": [], "powerpoint": None}
    rows: list[tuple[str, list[Image.Image]]] = []
    files: list[tuple[str, int, Path]] = []
    try:
        for i in range(count):
            if names and designs[i] not in names:
                continue
            pics: list[Image.Image] = []
            entry = {"name": designs[i], "variants": []}
            for k in range(4):
                r = page.js(f"({BUILD})({i}, {k})")
                f = OUT / f"{i:02d}-{designs[i]}-v{k + 1}.pptx"
                f.write_bytes(base64.b64decode(r["bytes"]))
                files.append((designs[i], k, f))
                entry["group"] = r["group"]
                entry["variants"].append({"variant": k + 1, "slides": r["slides"], "bytes": f.stat().st_size})
                for s in ([1, 2, 3, 4] if k == 0 else [1] if k < 3 else [1, 2]):
                    page.js(f"window.__present.renderAlone({s}, {THUMB * 2})")
                    page.js("new Promise((r) => setTimeout(r, 120))")
                    pics.append(Image.open(io.BytesIO(page.shot(0, 0, THUMB * 2, THUMB * 9 // 8))).convert("RGB"))
                page.js("document.getElementById('render-alone')?.remove(); window.__present.app.docs.length")
            rows.append((f"{designs[i]}  ({entry['group']})  -  Variante 1: Titel, Inhalt, Abschnitt, Tabelle/Diagramm  |  Varianten 2, 3, 4 (+Inhalt)", pics))
            result["designs"].append(entry)
            print(designs[i], "ok", flush=True)
    finally:
        page.ws.close()
        subprocess.run(["powershell", "-NoProfile", "-Command",
                        "Get-CimInstance Win32_Process -Filter \"Name='msedge.exe'\" | Where-Object { $_.CommandLine -like '*raabe-designs-page-profile*' }"
                        " | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"], capture_output=True)
    for n in range(0, len(rows), 8):
        sheet(rows[n:n + 8], OUT / f"sheet-{n // 8 + 1}.png", 9)
    print("contact sheets in", OUT)

    if use_ppt:
        if owner_powerpoint_running():
            result["powerpoint"] = "skipped: a POWERPNT.EXE runs (the owner's)"
            print(result["powerpoint"])
        else:
            result["powerpoint"] = check_powerpoint(files)
    out = HERE / "results" / f"{result['date']}-present-designs.json"
    out.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print("result:", out.name)


def check_powerpoint(files: list[tuple[str, int, Path]]) -> dict:
    """Each design's Variante 1 and 4 opened by our own invisible PowerPoint; what it reads, and its pictures."""
    try:
        ppt = PowerPointTarget(visible=False)
    except OwnerPowerPointRunning as e:
        return {"skipped": str(e)}
    out: dict = {"files": [], "opened": 0, "failed": 0}
    rows: list[tuple[str, list[Image.Image]]] = []
    try:
        cur: dict[str, list[Image.Image]] = {}
        for name, k, f in files:
            if k not in (0, 3):
                continue
            rec = {"design": name, "variant": k + 1}
            try:
                pres = ppt.open_copy(f, with_window=False)
                m = pres.SlideMaster
                rec.update({
                    "theme": m.Design.Name,
                    "fonts": [m.Theme.ThemeFontScheme.MajorFont(1).Name, m.Theme.ThemeFontScheme.MinorFont(1).Name],
                    "layouts": m.CustomLayouts.Count,
                    "master_shapes": m.Shapes.Count,
                    "title_layout_shapes": m.CustomLayouts(1).Shapes.Count,
                    "section_layout_shapes": m.CustomLayouts(3).Shapes.Count,
                    "slides": pres.Slides.Count,
                    "slide_shapes": [pres.Slides(i).Shapes.Count for i in range(1, pres.Slides.Count + 1)],
                    "accent1": "%06X" % _bgr(m.Theme.ThemeColorScheme(5).RGB),
                })
                pics = []
                for i in range(1, pres.Slides.Count + 1):
                    png = OUT / f"ppt-{f.stem}-{i}.png"
                    ppt.export_png(pres, i, png, THUMB * 2)
                    pics.append(Image.open(png).convert("RGB"))
                cur.setdefault(name, []).extend(pics if k == 0 else pics[:2])
                pres.Saved = -1
                pres.Close()
                ppt.own.remove(pres)
                out["opened"] += 1
            except Exception as e:  # noqa: BLE001
                rec["error"] = str(e)[:300]
                out["failed"] += 1
            out["files"].append(rec)
            print("PowerPoint", name, k + 1, "error" if "error" in rec else "ok", flush=True)
        rows = [(f"{n}  -  PowerPoint: Variante 1 (4 Folien) | Variante 4 (2 Folien)", p) for n, p in cur.items()]
    finally:
        out["close"] = ppt.close(kill_if_unreachable=True)
    for n in range(0, len(rows), 8):
        sheet(rows[n:n + 8], OUT / f"ppt-sheet-{n // 8 + 1}.png", 6)
    return out


def _bgr(v: int) -> int:
    """COM's RGB (0xBBGGRR) as 0xRRGGBB."""
    return ((v & 0xFF) << 16) | (v & 0xFF00) | ((v >> 16) & 0xFF)


if __name__ == "__main__":
    main()
