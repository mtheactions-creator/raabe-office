"""Macro tests against PowerPoint (D-080 B, D-105). PowerPoint (invisible, our own, only when no POWERPNT.EXE runs)
builds a test presentation through COM; RaabePresent (headless Edge on the dev server) adds our test modules
(macros/*.bas, written by us; openly licensed code under macros/thirdparty with its licence) with our VBA writer
and saves a .pptm. Each test then runs on a fresh copy of that .pptm twice: in PowerPoint through COM
(Application.Run; no Office security setting is changed: a .pptm opened through COM runs its macros) and in
RaabePresent; afterwards both presentations are described the same way (per slide each shape's name, type, box,
text and font; the notes) and compared, with files the macros wrote.

    .venv/Scripts/python.exe check_present_macro_suite.py powerpoint   (records PowerPoint's results)
    .venv/Scripts/python.exe check_present_macro_suite.py present      (runs ours and compares)

Results: results/<date>-present-macro-suite.json; PowerPoint's raw results in captures/macro-suite (not committed).
"""
from __future__ import annotations

import base64
import datetime
import json
import sys
import tempfile
from pathlib import Path

from PIL import Image

HERE = Path(__file__).parent
CAP = HERE / "captures" / "macro-suite"
WORK = Path(tempfile.gettempdir()) / "raabe-macro-suite"
NOTES = WORK / "notizen.txt"
TESTS = ["FormNamen", "SchriftErsetzen", "BilderGleichGross", "NotizenExportieren", "LeerePlatzhalterLoeschen", "Countdown", "Quiz", "NamenZiehen", "Fehler", "Funktionen"]


def modules() -> list[dict]:
    out = []
    # Our own modules only: openly licensed code under macros/thirdparty runs only once the owner allows running
    # downloaded code (the session's safety check refused it on 2026-11-08, D-105).
    for p in sorted((HERE / "macros").glob("*.bas")):
        code = p.read_text(encoding="utf-8").replace("%NOTES%", str(NOTES))
        lines = code.splitlines()
        name = lines[0].split('"')[1] if lines and lines[0].startswith("Attribute VB_Name") else p.stem
        body = "\r\n".join(lines[1:] if lines and lines[0].startswith("Attribute VB_Name") else lines)
        out.append({"name": name, "kind": "standard", "code": body})
    return out


def make_deck(t) -> Path:
    """The test presentation, made by PowerPoint: a title slide, a question, pictures, empty placeholders, results."""
    WORK.mkdir(parents=True, exist_ok=True)
    pics = []
    for i, (w, h, c) in enumerate([(240, 120, (200, 60, 60)), (160, 160, (60, 160, 60)), (300, 100, (60, 60, 200))]):
        p = WORK / f"bild{i}.png"
        Image.new("RGB", (w, h), c).save(p)
        pics.append(p)
    pres = t.new_presentation(with_window=False, name="macro-suite")
    lay = pres.SlideMaster.CustomLayouts
    s1 = pres.Slides.AddSlide(1, lay(1))
    s1.Shapes(1).TextFrame.TextRange.Text = "Mathe-Quiz"
    s2 = pres.Slides.AddSlide(2, lay(2))
    s2.Shapes(1).TextFrame.TextRange.Text = "Frage 1"
    s2.Shapes(2).TextFrame.TextRange.Text = "Was ist 3 · 4?\r12\r13"
    s2.Shapes(2).TextFrame.TextRange.Font.Name = "Comic Sans MS"
    s2.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = "Notiz eins"
    s3 = pres.Slides.AddSlide(3, lay(2))
    s3.Shapes(1).TextFrame.TextRange.Text = "Bilder"
    for i, p in enumerate(pics):
        s3.Shapes.AddPicture(str(p), 0, -1, 60 + i * 280, 200)
    s3.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = "Notiz zwei\rZweite Zeile"
    s4 = pres.Slides.AddSlide(4, lay(2))
    tb = s4.Shapes.AddTextbox(1, 500, 60, 300, 40)
    tb.TextFrame.TextRange.Text = "Arial-Text"
    tb.TextFrame.TextRange.Font.Name = "Arial"
    s5 = pres.Slides.AddSlide(5, lay(6))
    s5.Shapes(1).TextFrame.TextRange.Text = "Ergebnis"
    out = s5.Shapes.AddTextbox(1, 60, 150, 600, 40)
    out.Name = "Ausgabe"
    out.TextFrame.TextRange.Text = "-"
    pts = s5.Shapes.AddShape(1, 700, 150, 120, 60)
    pts.Name = "Punkte"
    pts.TextFrame.TextRange.Text = "0"
    nm = s5.Shapes.AddTextbox(1, 60, 260, 300, 120)
    nm.Name = "Namen"
    nm.TextFrame.TextRange.Text = "Anna\rBen\rCem\rDana"
    btn = s5.Shapes.AddShape(1, 700, 260, 120, 40)
    btn.Name = "Knopf Richtig"
    btn.TextFrame.TextRange.Text = "Richtig"
    CAP.mkdir(parents=True, exist_ok=True)
    deck = CAP / "deck.pptx"
    t.save_copy(pres, deck)
    pres.Saved = -1
    return deck


def describe_powerpoint(pres) -> list[dict]:
    slides = []
    for i in range(1, pres.Slides.Count + 1):
        sl = pres.Slides(i)
        shapes = []
        for k in range(1, sl.Shapes.Count + 1):
            s = sl.Shapes(k)
            d = {"name": str(s.Name), "type": int(s.Type), "box": [round(float(v), 1) for v in (s.Left, s.Top, s.Width, s.Height)]}
            if s.HasTextFrame and s.TextFrame.HasText:
                d["text"] = str(s.TextFrame.TextRange.Text)
                d["font"] = str(s.TextFrame.TextRange.Font.Name)
            try:
                if s.Fill.Visible and int(s.Type) == 1:
                    d["fill"] = int(s.Fill.ForeColor.RGB)
            except Exception:  # noqa: BLE001
                pass
            shapes.append(d)
        notes = str(sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text)
        slides.append({"shapes": shapes, "notes": notes})
    return slides


def write_pptm(deck: Path) -> Path:
    from feellab.present_page import PresentPage
    page = PresentPage()
    try:
        page.open_file(deck)
        b64 = page.js(f"""(async () => {{ const P = window.__present; const a = P.app; const m = await import('/src/macros.ts');
            m.setModules(a, {json.dumps(modules())});
            a.deck.pptx.setType(a.deck.pptx.presPath, 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml');
            return P.bytes(); }})()""")
    finally:
        page.close()
    p = CAP / "suite.pptm"
    p.write_bytes(base64.b64decode(b64))
    return p


class Watchdog:
    """A thread that ends VBA's error boxes in our PowerPoint (process `pid`) and keeps their text: a macro that
    fails must not leave PowerPoint waiting for a click. Only windows of our own process are touched."""

    def __init__(self, pid: int):
        import threading
        self.pid, self.seen, self.stop = pid, [], False
        self.thread = threading.Thread(target=self.run, daemon=True)
        self.thread.start()

    def run(self) -> None:
        import time
        import comtypes
        from feellab import uia, winapi
        comtypes.CoInitialize()
        while not self.stop:
            for h in winapi.find_windows(lambda h: winapi.window_pid(h) == self.pid and winapi.window_class(h) == "#32770"):
                if winapi.window_title(h) != "Microsoft Visual Basic":
                    continue
                root = uia.element_for(h)
                text = uia.find(root, lambda x: x["type"] == "Text", 3)
                self.seen.append(uia.describe(text)["name"] if text is not None else "?")
                end = uia.find(root, lambda x: x["type"] == "Button" and x["name"] in ("Beenden", "OK"), 3)
                if end is not None:
                    uia.invoke(end)
            time.sleep(0.3)


def run_powerpoint() -> None:
    from feellab import winapi
    from feellab.powerpoint import PowerPointTarget
    from check_present_screens import display, monitors
    CAP.mkdir(parents=True, exist_ok=True)
    out: dict = {"date": datetime.date.today().isoformat(), "tests": {}}
    # ActivePresentation needs a document window: PowerPoint's goes to the virtual display (D-073).
    before = monitors()
    mons = display("extend", 2)
    virtual = [m for m in mons if m not in before] or mons[1:]
    t = PowerPointTarget(visible=True, place=False)
    dog = Watchdog(t.pid)
    try:
        deck = make_deck(t)
        pptm = write_pptm(deck)
        for name in TESTS:
            if NOTES.exists():
                NOTES.unlink()
            pres = t.open_copy(pptm, with_window=True)
            if virtual and t.hwnd:
                winapi.place(t.hwnd, virtual[0][0], virtual[0][1], virtual[0][2], virtual[0][3])
            r: dict = {}
            n_seen = len(dog.seen)
            try:
                t.app.Run(f"{pres.Name}!Tests.{name}")
            except Exception as e:  # noqa: BLE001
                r["error"] = str(e)[:300]
            if len(dog.seen) > n_seen:
                r["vba_error"] = dog.seen[n_seen:]
            r["slides"] = describe_powerpoint(pres)
            if NOTES.exists():
                r["file"] = NOTES.read_text(encoding="cp1252", errors="replace")
            pres.Saved = -1
            pres.Close()
            t.own.remove(pres)
            out["tests"][name] = r
            print(name, r.get("error") or r["slides"][-1]["shapes"][1].get("text"), flush=True)
    finally:
        dog.stop = True
        out["close"] = t.close(kill_if_unreachable=True)
        display("internal", 1)
    (CAP / "powerpoint.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


DESCRIBE = r"""(() => { const a = window.__present.app; const E = 12700, r1 = (v) => Math.round(v * 10) / 10;
  const type = (v) => v.ph ? 14 : v.kind === 'pic' ? 13 : v.kind === 'grpSp' ? 6 : v.frame === 'table' ? 19 : v.frame === 'chart' ? 3 : (v.kind === 'cxnSp' || v.geometry === 'line') ? 9 : v.textBox ? 17 : 1;
  const bgr = (hex) => { const n = parseInt(hex, 16); return ((n >> 16) & 255) | (n & 0xff00) | ((n & 255) << 16); };
  return a.deck.slides.map((s) => { const v = a.deck.view(s.path);
    const shapes = v.shapes.map((x) => { const d = { name: x.name, type: type(x), box: [r1(x.x / E), r1(x.y / E), r1(x.w / E), r1(x.h / E)] };
      if (x.text && !x.text.empty) { d.text = x.text.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\r'); d.font = x.text.paras[0]?.runs[0]?.font ?? ''; }
      const fl = x.format?.fill ?? x.info?.styleFill; if (type(x) === 1 && fl?.kind === 'solid') d.fill = bgr(fl.color);
      return d; });
    return { shapes, notes: a.deck.notesText(s.path).replace(/\n/g, '\r') }; }); })()"""


# PowerPoint's object model names default placeholders in English ("Title 1") although the file says "Titel 1"
# (measured); the comparison takes them as the same.
EN_DE = {"Title": "Titel", "Subtitle": "Untertitel", "Content Placeholder": "Inhaltsplatzhalter", "Text Placeholder": "Textplatzhalter",
         "Picture Placeholder": "Bildplatzhalter", "Date Placeholder": "Datumsplatzhalter", "Footer Placeholder": "Fußzeilenplatzhalter",
         "Slide Number Placeholder": "Foliennummernplatzhalter", "Picture": "Grafik", "TextBox": "Textfeld", "Rectangle": "Rechteck"}


def same_name(a: str, b: str) -> bool:
    if a == b:
        return True
    for en, de in EN_DE.items():
        if a.startswith(en + " ") and b == de + a[len(en):]:
            return True
    return False


def compare_slides(ppt: list, ours: list) -> list[str]:
    """Differences: shape count, names, types, boxes (0.6 pt), texts, fonts, fills; notes."""
    out = []
    if len(ppt) != len(ours):
        return [f"slides: {len(ppt)} vs {len(ours)}"]
    for i, (a, b) in enumerate(zip(ppt, ours), 1):
        if a["notes"] != b["notes"]:
            out.append(f"slide {i} notes: {a['notes']!r} vs {b['notes']!r}")
        if len(a["shapes"]) != len(b["shapes"]):
            out.append(f"slide {i} shapes: {[x['name'] for x in a['shapes']]} vs {[x['name'] for x in b['shapes']]}")
            continue
        for x, y in zip(a["shapes"], b["shapes"]):
            if not same_name(x["name"], y["name"]):
                out.append(f"slide {i} name: {x['name']!r} vs {y['name']!r}")
            for k in ("type", "text", "font", "fill"):
                if x.get(k) != y.get(k):
                    out.append(f"slide {i} {x['name']} {k}: {x.get(k)!r} vs {y.get(k)!r}")
            if any(abs(p - q) > 0.6 for p, q in zip(x["box"], y["box"])):
                out.append(f"slide {i} {x['name']} box: {x['box']} vs {y['box']}")
    return out


def run_present() -> None:
    from feellab.present_page import PresentPage
    rec = json.loads((CAP / "powerpoint.json").read_text(encoding="utf-8"))
    out: dict = {"date": datetime.date.today().isoformat(), "tests": {}}
    page = PresentPage()
    try:
        for name in TESTS:
            page.open_file(CAP / "suite.pptm")
            r = page.js(f"""(async () => {{ const a = window.__present.app; const m = await import('/src/macros.ts');
                const mem = new Map(); window.__macroFiles = mem; const asked = [];
                const res = await m.runMacroQuiet(a, {json.dumps(name)}, {{ fileAccess: async (k, p) => {{ asked.push(k + ' ' + p); return true; }},
                  files: {{ read: async (p) => {{ if (!mem.has(p)) throw new Error('nicht da'); return mem.get(p); }}, write: async (p, d) => {{ mem.set(p, d); }} }} }});
                const f = mem.get({json.dumps(str(NOTES))});
                return {{ ...res, asked, file: f ? Array.from(f, (b) => String.fromCharCode(b)).join('') : null }}; }})()""")
            ours = {"slides": page.js(DESCRIBE)}
            if r.get("error"):
                ours["error"] = r["error"]
            if r.get("file") is not None:
                ours["file"] = r["file"].encode("latin-1").decode("cp1252", errors="replace")
            ours["asked"] = r.get("asked", [])
            theirs = rec["tests"][name]
            diff = compare_slides(theirs["slides"], ours["slides"])
            if (theirs.get("file") or "").replace(chr(13), "") != (ours.get("file") or "").replace(chr(13), ""):
                diff.append(f"file: {theirs.get('file')!r} vs {ours.get('file')!r}")
            if theirs.get("vba_error") or ours.get("error"):
                diff.append(f"errors: {theirs.get('vba_error')} vs {ours.get('error')}")
            out["tests"][name] = {"result": "same" if not diff else "different", "diff": diff, "asked": ours["asked"]}
            print(name, "same" if not diff else diff[:4], flush=True)
    finally:
        page.close()
    out["summary"] = f"{sum(1 for v in out['tests'].values() if v['result'] == 'same')}/{len(out['tests'])}"
    p = HERE / "results" / f"{out['date']}-present-macro-suite.json"
    import tempfile as _t
    text = json.dumps(out, ensure_ascii=False, indent=1)
    for tmp in {_t.gettempdir(), str(Path(_t.gettempdir()).resolve())}:
        text = text.replace(json.dumps(tmp)[1:-1], "%TEMP%").replace(tmp, "%TEMP%")
    p.write_text(text, encoding="utf-8", newline=chr(10))
    print("same as PowerPoint:", out["summary"])


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "powerpoint":
        run_powerpoint()
    else:
        run_present()
