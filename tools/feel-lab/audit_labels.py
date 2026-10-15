"""Audit (Step 4c F): labels the ribbon cuts off, at window widths from 1920 down to 800 CSS
pixels, on every tab (the contextual ones with a picture, a shape, a table and an equation).
Runs in the headless development page (devpage.py, `npm run dev:write`): nothing touches the
real screen. A text is cut off when it reaches past an ancestor that hides overflow, or past
the ribbon's right edge. Result: results/<date>-audit-labels.json.

    .venv/Scripts/python.exe audit_labels.py [width …]
"""
from __future__ import annotations

import datetime
import json
import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-audit-labels.json"
PY = str(HERE / ".venv" / "Scripts" / "python.exe")
FIX = HERE.parents[1] / "packages" / "ooxml" / "test" / "fixtures"
WIDTHS = [1920, 1600, 1366, 1280, 1100, 1000, 900, 800]

CHECK = r"""(() => {
  const card = document.querySelector('.ribbon-card');
  if (!card) return null;
  const edge = card.getBoundingClientRect();
  const out = [];
  const clips = (el) => { const list = []; for (let e = el; e && e !== card.parentElement; e = e.parentElement) { const cs = getComputedStyle(e); if (cs.overflowX !== 'visible' || cs.overflow !== 'visible') list.push(e); } return list; };
  const walker = document.createTreeWalker(card, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const t = n.textContent.trim();
    const el = n.parentElement;
    // Style names in the Formatvorlagen gallery are cut at the tile's edge, as in Word.
    if (!t || !el || el.closest('svg') || el.closest('.rgallery .item') || el.closest('[hidden]') || getComputedStyle(el).visibility === 'hidden' || getComputedStyle(el).display === 'none') continue;
    const r = document.createRange(); r.selectNodeContents(n);
    const tr = r.getBoundingClientRect();
    if (!tr.width || !tr.height) continue;
    let cut = tr.right > edge.right + 1 ? 'ribbon edge' : '';
    for (const c of clips(el)) {
      const cr = c.getBoundingClientRect();
      if (tr.left < cr.left - 1 || tr.right > cr.right + 1 || tr.top < cr.top - 1 || tr.bottom > cr.bottom + 1) { cut = (typeof c.className === 'string' && c.className) || c.tagName; break; }
    }
    // Faded or ellipsised ends count too.
    const cs = getComputedStyle(el);
    if (!cut && cs.textOverflow === 'ellipsis' && el.scrollWidth > el.clientWidth + 1) cut = 'ellipsis';
    if (cut) out.push({ text: t.slice(0, 40), cut, x: Math.round(tr.left), w: Math.round(tr.width) });
  }
  return out;
})()"""


def dev(*args: str) -> str:
    return subprocess.run([PY, str(HERE / "devpage.py"), *args], capture_output=True, text=True, encoding="utf-8").stdout.strip()


def tabs() -> list[str]:
    return json.loads(dev("js", "[...document.querySelectorAll('.tabrow [role=tab]')].filter(t => t.offsetParent).map(t => t.textContent.trim())") or "[]")


def select_tab(name: str) -> None:
    dev("js", f"(() => {{ const t = [...document.querySelectorAll('.tabrow [role=tab]')].find(t => t.textContent.trim() === {json.dumps(name)}); t?.click(); return !!t; }})()")
    time.sleep(0.5)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    widths = [int(a) for a in sys.argv[1:]] or WIDTHS
    out: dict = {}
    setups = [
        ("Text", None, None),
        ("Bild", "word-images.docx", "(() => { const e = __app.editor(); const ps = e.paragraphs(); for (let p = 0; p < ps.length; p++) { const o = ps[p].text.indexOf('\\ufffc'); if (o >= 0 && ps[p].objects?.some(x => x.kind === 'image')) { e.select({p, o}, {p, o: o + 1}); return p; } } return -1; })()"),
        ("Form", "word-shapes.docx", "(() => { const e = __app.editor(); const ps = e.paragraphs(); for (let p = 0; p < ps.length; p++) { const fs = ps[p].floats ?? []; for (let i = 0; i < fs.length; i++) if (fs[i].shape) { e.selectFloat(p, i); return p; } } return -1; })()"),
        ("Tabelle", "word-tables.docx", "(() => { const e = __app.editor(); const p = e.paragraphs().findIndex(x => x.cell); e.select({p, o: 0}, {p, o: 0}); return p; })()"),
    ]
    for label, doc, select in setups:
        dev("reload")
        time.sleep(2.5)
        if doc:
            dev("open", str(FIX / doc))
            time.sleep(1.5)
            dev("js", select)
            time.sleep(0.5)
        for w in widths:
            dev("viewport", str(w), "900", "1.5")
            for tab in tabs():
                if label != "Text" and tab in ("Start", "Einfügen", "Entwurf", "Layout", "Referenzen", "Überprüfen", "Ansicht", "Datei") and not (label == "Tabelle" and tab == "Layout"):
                    continue
                select_tab(tab)
                res = json.loads(dev("js", CHECK) or "[]")
                key = f"{tab}@{w}"
                if res:
                    out[key] = res
                    print(key, [f"{r['text']} ({r['cut']})" for r in res][:8], flush=True)
    dev("viewport", "reset")
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(len(out), "tab/width combinations with cut-off labels")


if __name__ == "__main__":
    main()
