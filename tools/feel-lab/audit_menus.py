"""Audit (Step 4c F): RaabeWrite's menus and galleries entry by entry against Word's
recordings (results/*-word-dialogs.json made with probe_word_dialogs/panes, and the context
menus of *-word-context*.json). Each recorded menu is opened in the headless development page
(devpage.py) the same way (tab and control, or a right click on the same kind of object) and
its entries are compared by name: what Word has and we lack. Differences that are decisions
are listed in EXPLAINED. Nothing touches the real screen. Result: results/<date>-audit-menus.json.

    .venv/Scripts/python.exe audit_menus.py [menu name filter …]
"""
from __future__ import annotations

import datetime
import json
import re
import subprocess
import sys
import time
from pathlib import Path

from probe_word_dialogs import ITEMS as DIALOG_ITEMS
from probe_word_panes import ITEMS as PANE_ITEMS

HERE = Path(__file__).parent
OUT = HERE / "results" / f"{datetime.date.today()}-audit-menus.json"
PY = str(HERE / ".venv" / "Scripts" / "python.exe")
# Entries that are no menu entries (scroll bars, filters, window buttons).
NOISE = re.compile(r"^(Bildlauf|Seite nach|Zeile nach|Schließen|Minimieren|Maximieren|Nur hoher Kontrast|Katalogfilter|Weitere Optionen|Öffnen)\b")

# Differences on purpose (or depending on the document), with the reason.
EXPLAINED = {
    "Copy Link to Location...": "cloud documents only",
    "Laut vorlesen": "deferred (D-041)",
    "Übersetzen": "deferred (D-041)",
    "In Office Math konvertieren": "handwriting conversion, not planned",
    "Alle in Office Math konvertieren": "handwriting conversion, not planned",
    "Auswahl im Editor überprüfen": "Word's Editor pane (cloud); F7 checks spelling",
    "Bild bearbeiten": "Microsoft's online picture editor",
    "Symbol: 183": "recently used symbols differ per user",
    "White Square": "recently used symbols differ per user",
    "Latin Capital Letter O With Stroke": "recently used symbols differ per user",
    "Abstand vor Absatz entfernen": "depends on the paragraph (hinzufügen when there is none)",
}
# Word recorded these in a new document of its own, which has the 2023 Office theme.
OFFICE_2023 = "{dk1:'000000',lt1:'FFFFFF',dk2:'0E2841',lt2:'E8E8E8',accent1:'156082',accent2:'E97132',accent3:'196B24',accent4:'0F9ED5',accent5:'A02B93',accent6:'4EA72E',hlink:'467886',folHlink:'96607D'}"
SETUP = {"page color menu": "(() => { const e = __app.editor(); e.setTheme({ ...e.theme, colors: %s }); })()" % OFFICE_2023}
EXPLAINED_MENUS = {"cover page gallery": "RaabeWrite's own cover pages for school work (D-048)"}

OPEN = r"""(async (tab, control, how) => {
  const t = [...document.querySelectorAll('.tabrow [role=tab]')].find(x => x.textContent.trim() === tab);
  if (!t) return { error: 'no tab ' + tab };
  t.click();
  await new Promise(r => setTimeout(r, 300));
  const card = document.querySelector('.ribbon-card');
  const all = [...card.querySelectorAll('button, [role=button], .rcombo')];
  const labels = (b) => [b.getAttribute('aria-label'), b.querySelector('input')?.getAttribute('aria-label'), (b.dataset.tip || '').split('\n')[0], b.textContent].filter(Boolean).map((s) => s.trim());
  let b = all.find(x => labels(x).includes(control)) || all.find(x => labels(x).some((n) => n.startsWith(control)));
  if (!b) return { error: 'no control ' + control };
  const before = new Set(document.querySelectorAll('.menu, .rpanel, .popup'));
  if (how === 'arrow' || how.startsWith('arrow:')) {
    const wrap = b.closest('.rsplit') || b.parentElement;
    const arrow = [...wrap.querySelectorAll('button')].pop();
    arrow.click();
  } else if (how === 'combo') {
    const arrow = b.querySelector('button[aria-label="Öffnen"]') || b.parentElement.querySelector('button[aria-label="Öffnen"]');
    (arrow || b).click();
  } else b.click();
  await new Promise(r => setTimeout(r, 500));
  const pops = [...document.querySelectorAll('.menu, .rpanel, .popup')].filter(p => !before.has(p) && p.isConnected);
  const names = new Set();
  for (const p of pops) for (const e of p.querySelectorAll('button, [role=menuitem], .title, label')) {
    for (const v of [e.getAttribute('aria-label'), (e.dataset.tip || '').split('\n')[0], e.textContent]) if (v && v.trim()) names.add(v.trim());
  }
  document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
  return { names: [...names], popups: pops.length };
})"""


FIX = HERE.parents[1] / "packages" / "ooxml" / "test" / "fixtures"
POINT = ("(() => { const e = __app.editor(); const r = e.view.selectedObjectRect(); if (r) return [r.x + %s, r.y + %s];"
         " const b = e.view.caretBox(e.sel.anchor); const p = e.view.pageToClient(b.page, b.x, b.y + b.h / 2); return [p.x + 3, p.y]; })()")
FIND_OBJ = "(() => { const e = __app.editor(); const ps = e.paragraphs(); for (let p = 0; p < ps.length; p++) { %s } return -1; })()"
# Word's recorded context menus: kind → (document, setup script, offset into the object).
CONTEXT = {
    "text": (None, "(() => { const e = __app.editor(); e.insertText('Hallo Welt'); e.select({p: 0, o: 2}, {p: 0, o: 2}); })()", "c"),
    "list": (None, "(() => { const e = __app.editor(); e.insertText('Eins'); __app.run('para.bullets'); e.select({p: 0, o: 2}, {p: 0, o: 2}); })()", "c"),
    "table": (None, "(() => { const e = __app.editor(); e.insertTable(3, 3); e.insertText('a'); })()", "c"),
    # Word's recording: the whole 2 x 2 table selected.
    "table-cells": (None, "(() => { const e = __app.editor(); e.insertTable(2, 2); e.insertText('a'); __app.run('table.selectTable'); })()", "c"),
    "picture": ("word-images.docx", FIND_OBJ % "const o = ps[p].text.indexOf('\\ufffc'); if (o >= 0 && ps[p].objects?.some(x => x.kind === 'image')) { e.select({p, o}, {p, o: o + 1}); return p; }", "c"),
    "floating-picture": ("word-images.docx", FIND_OBJ % "const fs = ps[p].floats ?? []; for (let i = 0; i < fs.length; i++) if (!fs[i].shape) { e.selectFloat(p, i); return p; }", "c"),
    "textbox-frame": ("word-shapes.docx", FIND_OBJ % "const fs = ps[p].floats ?? []; for (let i = 0; i < fs.length; i++) if (fs[i].shape?.kind === 'textbox') { e.selectFloat(p, i); return p; }", "edge"),
}
ITEMS_JS = ("[...document.querySelectorAll('.menu')].filter(m => m.isConnected).flatMap(m => [...m.querySelectorAll('button, [role=menuitem]')]"
            ".map(b => (b.getAttribute('aria-label') || b.textContent || '').trim()))")


def context_menus(out: dict, only: list[str]) -> None:
    recorded: dict = {}
    for f in sorted((HERE / "results").glob("*-word-context*.json")):
        for m in json.loads(f.read_text(encoding="utf-8")):
            if isinstance(m, dict) and "items" in m:
                recorded[m["kind"]] = [i["name"] for i in m["items"] if i["type"] in ("MenuItem", "SplitButton") and i["name"]]
    for kind, (doc, setup, off) in CONTEXT.items():
        if kind not in recorded or (only and not any(o in kind for o in only)):
            continue
        dev("reload")
        time.sleep(2.5)
        if doc:
            dev("open", str(FIX / doc))
            time.sleep(1.5)
        else:
            dev("js", "__app.newDoc()")
            time.sleep(0.5)
        dev("js", setup)
        time.sleep(0.4)
        dx, dy = ("r.width / 2", "r.height / 2") if off == "c" else ("2", "2")
        pt = json.loads(dev("js", POINT % (dx, dy)) or "null")
        if not pt:
            out[f"context {kind}"] = {"error": "no point"}
            print("context", kind, "ERROR no point", flush=True)
            continue
        dev("rclick", str(pt[0]), str(pt[1]))
        time.sleep(0.4)
        ours = [norm(x) for x in json.loads(dev("js", ITEMS_JS) or "[]")]
        dev("key", "Escape")
        word = list(dict.fromkeys(recorded[kind]))
        missing = [w for w in word if not any(norm(w) == o or o.startswith(norm(w)) for o in ours)]
        res = {"word": len(word), "missing": [m for m in missing if m not in EXPLAINED], "explained": {m: EXPLAINED[m] for m in missing if m in EXPLAINED}}
        out[f"context {kind}"] = res
        print("context", kind, f"{len(word) - len(missing)}/{len(word)}", "missing:", res["missing"], flush=True)


def dev(*args: str) -> str:
    return subprocess.run([PY, str(HERE / "devpage.py"), *args], capture_output=True, text=True, encoding="utf-8").stdout.strip()


def norm(s: str) -> str:
    return re.sub(r"[\s.…:]+", " ", s.replace("\u00ad", "")).strip().lower()


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    only = sys.argv[1:]
    recorded: dict = {}
    for f in sorted((HERE / "results").glob("*-word-dialogs.json")):
        recorded.update(json.loads(f.read_text(encoding="utf-8")))
    ways = {name: (tab, control, how) for name, tab, control, how, *_ in [*DIALOG_ITEMS, *PANE_ITEMS]}
    dev("reload")
    time.sleep(2.5)
    dev("js", "__app.newDoc()")
    time.sleep(1)
    out: dict = {}
    for name, data in recorded.items():
        if not isinstance(data, dict) or data.get("window") != "Net UI Tool Window" or name not in ways:
            continue
        if only and not any(o in name for o in only):
            continue
        tab, control, how = ways[name]
        items = data["items"]
        if how.startswith("menu:"):
            # Word's group was collapsed: its menu is the one behind the named entry; ours is on the ribbon.
            control, how = how[5:], "click"
            items = items[max(i for i, x in enumerate(items) if x["name"] == control) + 1:]
        word = [re.sub(r" \(Keine Daten\)$", "", i["name"]) for i in items if i["type"] in ("MenuItem", "ListItem", "SplitButton", "Button", "CheckBox", "RadioButton")
                and i["name"] and not NOISE.match(i["name"]) and i["name"] != control]
        if name in SETUP:
            dev("js", "__app.newDoc()")
            dev("js", SETUP[name])
        res = json.loads(dev("js", f"({OPEN})({json.dumps(tab)}, {json.dumps(control)}, {json.dumps(how)})") or "{}")
        if "error" in res:
            out[name] = {"error": res["error"]}
            print(name, "ERROR", res["error"], flush=True)
            continue
        ours = [norm(x) for x in res["names"]]
        missing = [w for w in dict.fromkeys(word) if not any(norm(w) == o or norm(w) in o for o in ours)]
        out[name] = {"word": len(set(word)), "missing": [m for m in missing if m not in EXPLAINED], "explained": {m: EXPLAINED[m] for m in missing if m in EXPLAINED}}
        if name in EXPLAINED_MENUS and missing:
            out[name] = {"word": len(set(word)), "missing": [], "explained": {"*": EXPLAINED_MENUS[name]}}
        print(name, f"{len(set(word)) - len(missing)}/{len(set(word))}", "missing:", out[name]["missing"][:10], flush=True)
    context_menus(out, only)
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
