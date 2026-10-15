"""Line starts of a Word-built document: Word (its own invisible instance) against the
development build in devpage.py's headless Edge. Quicker than `run.py docx` while working on
the layout (no exe build); the exe stays the reference for the reports.

    .venv/Scripts/python.exe devpage.py start
    .venv/Scripts/python.exe compare_dev.py ../../packages/ooxml/test/fixtures/word-wrap.docx
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from devpage import Page  # noqa: E402
from feellab.docx_fidelity import WordDocument  # noqa: E402


def main(path: Path) -> None:
    word = WordDocument(path)
    try:
        text = word.text()
        starts = word.lines_by_position()
        wstarts = [word.to_ours(s) for s in starts]
    finally:
        word.close()
    pg = Page()
    pg.js(f"(async () => {{ const r = await fetch('/@fs/{path.resolve().as_posix()}'); window.__app.openBytes({json.dumps(path.name)}, new Uint8Array(await r.arrayBuffer())); }})()")
    time.sleep(1.0)
    for _ in range(50):
        if pg.js("__probe.settled()"):
            break
        time.sleep(0.1)
    if "--recompute" in sys.argv:
        # Our own wrap polygons instead of Word's (Eng/Transparent set again on every float).
        print(pg.js("(async () => { const e = __app.editor(); let n = 0; for (const [k, p] of e.doc.main.entries()) (p.floats ?? []).forEach((f, i) => {"
                    " if (f.wrap !== 'tight' && f.wrap !== 'through') return; const w = f.wrap; e.selectFloat(k, i); e.setImageWrap('square'); e.setImageWrap(w); n++; });"
                    " await new Promise(r => setTimeout(r, 800)); return n; })()"), "polygons recomputed")
        time.sleep(0.5)
    ours = pg.js("__probe.lines()")
    otext = pg.js("__probe.text()")
    if otext != text:
        print("texts differ:", len(otext), len(text))
    ws, os_ = set(wstarts), set(ours)
    both = ws & os_
    print(f"{path.name}: Word {len(ws)} lines, ours {len(os_)}, same starts {len(both)}")
    # Per paragraph: how many lines each has and how many starts agree.
    at = 0
    for para in text.split(chr(10)):
        a, e = at, at + len(para)
        w = [o for o in ws if a <= o <= e]
        u = [o for o in os_ if a <= o <= e]
        if len(w) > 1 or len(u) > 1:
            print(f"  {para[:28]!r:32} Word {len(w):3}  ours {len(u):3}  same {len(set(w) & set(u)):3}")
        at = e + 1
    if "--para" in sys.argv:
        k = int(sys.argv[sys.argv.index("--para") + 1])
        paras = text.split(chr(10))
        a = sum(len(x) + 1 for x in paras[:k])
        e = a + len(paras[k])
        rows = sorted({o for o in ws | os_ if a <= o <= e})
        for o in rows:
            print(f"   {'W' if o in ws else ' '}{'O' if o in os_ else ' '} {o:6} {text[o:o + 40]!r}")
    ctx = lambda o: text[max(0, o - 25):o].replace("\n", "¶") + "|" + text[o:o + 25].replace("\n", "¶")
    for o in sorted(ws - os_)[:15]:
        print("  only Word:", o, ctx(o))
    for o in sorted(os_ - ws)[:15]:
        print("  only ours:", o, ctx(o))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(Path(sys.argv[1]).resolve())
