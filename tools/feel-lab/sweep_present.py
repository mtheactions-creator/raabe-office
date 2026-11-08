"""By-hand sweep of RaabePresent (D-137): every option of the ribbon (all tabs, contextual ones included), every
entry of its menus, galleries and palettes (one level of submenus), the dialogs they open (OK and Abbrechen), the
context menus, the Datei backstage, the status bar and the keyboard shortcuts, triggered as a student triggers them:
trusted CDP mouse clicks at the control's place and CDP key events (no test API for the option itself). The starting
point for each option (a shape with text, a picture, a table, a chart, an equation, a SmartArt, the master views) is
made by sweep.ts in the app; then per option:

    - no exception, rejected promise or console error;
    - an effect: the document's parts change, or the UI does (a menu, a dialog, a pane, the view, the selection);
    - a document change is undone by Strg+Z (all parts as before);
    - the changed presentation is written (deck.write) and parses (zip and every XML part); with --office the files
      are opened by PowerPoint (an instance of our own, only when no POWERPNT.EXE runs; office_open_present.py).

Commands deliberately greyed (deferred, SCOPE.md "later") are listed as such, not as failures. Native file dialogs and
the clipboard are replaced in the page by stand-ins (a fixture path, a file in TEMP, an in-memory clipboard), so the
sweep never shows an OS window or touches the clipboard other sessions use.

    .venv/Scripts/python.exe sweep_present.py inventory [--dev]       -> results/<date>-present-options.json
    .venv/Scripts/python.exe sweep_present.py sweep [--dev] [--tab T] [--ctx C] [--max-items N]
                                                                       -> results/<date>-present-sweep.json

--dev runs against the dev server (npm run dev:present on port 5186 here: --url) in a headless Edge of our own;
otherwise the release exe (RAABE_PRESENT_EXE) is started with its DevTools port. Nothing takes the OS mouse.
"""
from __future__ import annotations

import argparse
import datetime
import io
import json
import os
import subprocess
import sys
import tempfile
import time
import urllib.request
import zipfile
from pathlib import Path
from xml.dom import minidom

import websocket

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
TMP = Path(tempfile.gettempdir()) / "raabe-present-sweep"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

# Which starting point each tab is swept in (contextual tabs appear by themselves there).
CONTEXTS: list[tuple[str, list[str]]] = [
    ("shape-text", ["Start", "Überprüfen"]),
    ("shape", ["Einfügen", "Zeichnen", "Animationen", "Formformat"]),
    ("slides", ["Entwurf", "Übergänge", "Bildschirmpräsentation", "Aufzeichnen", "Ansicht"]),
    ("picture", ["Bildformat"]),
    ("table", ["Tabellenentwurf", "Tabellenlayout"]),
    ("chart", ["Diagrammentwurf", "Format"]),
    ("equation", ["Formel"]),
    ("smartart", ["SmartArt-Design", "Format"]),
    ("master", ["Folienmaster"]),
    ("notesmaster", ["Notizenmaster"]),
    ("handoutmaster", ["Handzettelmaster"]),
    ("freistellen", ["Freistellen"]),
]

# Context menus: where the right click goes in which starting point.
CONTEXT_MENUS = [
    ("Form", "shape", "shape"), ("Textfeld", "textbox", "shape"), ("Text in Bearbeitung", "shape-text", "text"),
    ("Platzhaltertext", "placeholder", "text"), ("Tabelle", "table", "text"), ("Bild", "picture", "shape"),
    ("Diagramm", "chart", "shape"), ("SmartArt", "smartart", "shape"), ("Gruppe", "two-shapes", "shape"),
    ("Leerer Folienbereich", "nothing", "slide"), ("Folienminiatur", "slides", "thumb"), ("Notizenbereich", "slides", "notes"),
]

# Values typed into the ribbon's fields (combo boxes and spinners), by the field's accessible name.
TYPED = {"Schriftart": "Arial", "Schriftgrad": "24", "Höhe": "3 cm", "Breite": "5 cm", "Formhöhe": "3 cm", "Formbreite": "5 cm",
         "Dauer": "01,50", "Verzögerung": "00,50", "Nach": "00:03,00", "Zeilenhöhe": "1,5 cm", "Spaltenbreite": "4 cm"}

# ----- the page's helpers (installed once per page) ------------------------------------------------------------

HELPERS = r"""
(() => {
  if (window.__sw) return 'again';
  const sw = window.__sw = { errs: [], ext: [], clip: { text: '', html: '' } };
  addEventListener('error', (e) => sw.errs.push('error: ' + (e.message || '') + ' @' + (e.filename || '').split('/').pop() + ':' + (e.lineno || '')));
  addEventListener('unhandledrejection', (e) => { const r = e.reason; sw.errs.push('rejection: ' + (r && (r.stack || r.message) || String(r)).slice(0, 600)); });
  const ce = console.error.bind(console);
  console.error = (...a) => { sw.errs.push('console: ' + a.map((x) => (x && x.stack) || String(x)).join(' ').slice(0, 600)); ce(...a); };
  for (const c of Array.from(document.body.children)) c.setAttribute('data-sw-base', '');
  const app = window.__present.app, P = app.platform;
  // Stand-ins for the OS: file dialogs answer with prepared files, the clipboard stays in the page.
  P.openImageDialog = async () => sw.files.png;
  P.openImagesDialog = async () => [sw.files.png, sw.files.png2];
  P.openFileDialog = async () => sw.files.pptx;
  P.openAnyDialog = async (t, f) => { const ex = (f || []).flatMap((x) => x.extensions || []); return ex.includes('thmx') || ex.includes('pptx') ? sw.files.pptx : ex.includes('txt') || ex.includes('rtf') ? sw.files.txt : ex.includes('mp4') ? sw.files.mp4 : ex.includes('wav') || ex.includes('mp3') ? sw.files.wav : sw.files.png; };
  P.saveFileDialog = async (n, k) => sw.files.dir + '/sweep-save.' + k;
  P.saveAnyDialog = async (t, n, f) => sw.files.dir + '/sweep-save-any.' + (((f && f[0] && f[0].extensions) || ['bin'])[0]);
  P.openExternal = async (t) => { sw.ext.push(t); };
  P.clipboardRead = async () => sw.clip.text;
  P.clipboardWrite = async (t) => { sw.clip = { text: t, html: '' }; };
  P.clipboardReadHtml = async () => sw.clip.html;
  P.clipboardWriteHtml = async (h, t) => { sw.clip = { text: t, html: h }; };
  P.clipboardReadImage = async () => null;
  const label = (e) => (e.getAttribute('aria-label') || (e.dataset && e.dataset.tip || '').split('\n')[0] || e.textContent || e.getAttribute('title') || '').replace(/\s+/g, ' ').trim();
  sw.label = label;
  const vis = (e) => { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0 && getComputedStyle(e).visibility !== 'hidden'; };
  sw.rect = (e) => { const r = e.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2, w: r.width, h: r.height }; };
  sw.kind = (e) => {
    const c = e.classList;
    if (c.contains('launcher')) return 'launcher';
    if (e.tagName === 'INPUT') return e.closest('.rspinner') ? 'spinner-input' : 'combo-input';
    if (c.contains('rcombo-face')) return 'combo-face';
    if (e.closest('.rcombo')) return 'combo-arrow';
    if (e.closest('.spin')) return e.getAttribute('aria-label') === 'Mehr' ? 'spinner-up' : 'spinner-down';
    if (e.closest('.rgallery')) return c.contains('item') ? 'gallery-item' : 'gallery-more';
    if (c.contains('rcheck')) return 'check';
    if (c.contains('color-main')) return 'color-main';
    if (c.contains('arrow') && e.closest('.rsplit')) return e.parentElement.querySelector('.color-main') ? 'color-arrow' : 'split-arrow';
    if (e.closest('.rsplit')) return 'split-main';
    if (c.contains('has-chev')) return 'dropdown';
    if (e.closest('.rgroup.collapsed')) return 'collapsed-group';
    return 'button';
  };
  // The active tab's controls, each with a key that finds it again after a fresh start.
  sw.controls = () => {
    const card = document.querySelector('.ribbon-card');
    const out = [];
    if (!card) return out;
    for (const g of Array.from(card.querySelectorAll('.rgroup'))) {
      const gname = g.getAttribute('aria-label') || g.dataset.group || '?';
      const els = Array.from(g.querySelectorAll('button, input, .rcombo-face')).filter((e) => vis(e) && !(e.closest('.scroll') && e.getAttribute('aria-label') !== 'Weitere'));
      els.forEach((e, i) => {
        let name = label(e);
        const kind = sw.kind(e);
        if (kind === 'split-arrow' || kind === 'color-arrow') name = label(e.parentElement.firstElementChild) + ' (Pfeil)';
        if (kind === 'gallery-more') name = 'Weitere';
        if (kind === 'spinner-up' || kind === 'spinner-down') name = label(e.closest('.rspinner').querySelector('input')) + (kind === 'spinner-up' ? ' (Mehr)' : ' (Weniger)');
        if (kind === 'combo-arrow') name = label(e.closest('.rcombo').querySelector('input')) + ' (Öffnen)';
        if (kind === 'combo-face') name = label(e.closest('.rcombo').querySelector('input') || e) + ' (Feld)';
        out.push({ key: (g.dataset.group || gname) + '#' + i, group: gname, name, kind, disabled: !!e.disabled || e.classList.contains('disabled') || !!e.closest('.disabled'),
                   tip: (e.dataset && e.dataset.tip || (e.closest('[data-tip]') && e.closest('[data-tip]').dataset.tip) || '').slice(0, 300) });
      });
    }
    return out;
  };
  sw.findControl = (key) => {
    const [gid, n] = key.split('#');
    const card = document.querySelector('.ribbon-card');
    const g = Array.from(card.querySelectorAll('.rgroup')).find((x) => (x.dataset.group || x.getAttribute('aria-label')) === gid);
    if (!g) return null;
    const els = Array.from(g.querySelectorAll('button, input, .rcombo-face')).filter((e) => vis(e) && !(e.closest('.scroll') && e.getAttribute('aria-label') !== 'Weitere'));
    const e = els[Number(n)];
    if (!e) return null;
    e.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    return Object.assign(sw.rect(e), { disabled: !!e.disabled || e.classList.contains('disabled') });
  };
  sw.tabs = () => Array.from(document.querySelectorAll('.tabrow .tab')).filter(vis).map((t) => ({ name: t.textContent.trim(), contextual: t.classList.contains('contextual'), on: t.getAttribute('aria-selected') === 'true' || t.classList.contains('active') }));
  sw.tabRect = (name) => { const t = Array.from(document.querySelectorAll('.tabrow .tab')).filter(vis).find((x) => x.textContent.trim() === name); return t ? sw.rect(t) : null; };
  // What lies above the app: popups, menus, dialogs, the backstage, panes.
  sw.overlays = () => Array.from(document.body.children).filter((c) => !c.hasAttribute('data-sw-base') && !c.classList.contains('tooltip') && vis(c)).map((c) => c);
  sw.dialogs = () => Array.from(document.querySelectorAll('[role=dialog]')).filter(vis);
  sw.items = () => {
    // Entries of the open popup(s) (the topmost menu last): menu entries, palette swatches, gallery tiles, buttons.
    const ov = sw.overlays().filter((c) => !c.querySelector('[role=dialog]') && c.getAttribute('role') !== 'dialog');
    const out = [];
    ov.forEach((root, ri) => {
      const els = Array.from(root.querySelectorAll('button, [role=menuitem], [role=option], input, select, .item, .tile')).filter((e) => vis(e) && !(e.tagName !== 'BUTTON' && e.querySelector('button')));
      const uniq = els.filter((e, i) => els.indexOf(e) === i && !els.some((o) => o !== e && o.contains(e) && o.tagName === 'BUTTON'));
      uniq.forEach((e, i) => out.push({ path: ri + ':' + i, name: label(e), sub: e.classList.contains('has-sub') || e.getAttribute('aria-haspopup') === 'true', disabled: !!e.disabled || e.getAttribute('aria-disabled') === 'true',
        tag: e.tagName.toLowerCase(), type: e.type || '' }));
    });
    return out;
  };
  sw.itemEl = (path) => {
    const [ri, i] = path.split(':').map(Number);
    const root = sw.overlays().filter((c) => !c.querySelector('[role=dialog]') && c.getAttribute('role') !== 'dialog')[ri];
    if (!root) return null;
    const els = Array.from(root.querySelectorAll('button, [role=menuitem], [role=option], input, select, .item, .tile')).filter((e) => vis(e) && !(e.tagName !== 'BUTTON' && e.querySelector('button')));
    const uniq = els.filter((e, k) => els.indexOf(e) === k && !els.some((o) => o !== e && o.contains(e) && o.tagName === 'BUTTON'));
    return uniq[i] || null;
  };
  sw.itemRect = (path) => { const e = sw.itemEl(path); if (!e) return null; e.scrollIntoView({ block: 'nearest' }); return Object.assign(sw.rect(e), { name: label(e) }); };
  sw.dialogInfo = () => sw.dialogs().map((d) => ({ title: d.getAttribute('aria-label') || (d.querySelector('.dlg-title .t') || {}).textContent || '',
    controls: Array.from(d.querySelectorAll('input, select, textarea, button')).filter(vis).map((e) => ({ tag: e.tagName.toLowerCase(), type: e.type || '', name: label(e) || (e.closest('label') ? e.closest('label').textContent.trim().slice(0, 60) : '') })) }));
  sw.dialogButton = (which) => {
    const d = sw.dialogs().pop();
    if (!d) return null;
    const bs = Array.from(d.querySelectorAll('.dlg-buttons button, button')).filter(vis);
    const b = which === 'ok' ? (d.querySelector('.dlg-buttons button.primary') || bs.find((x) => /^(OK|Einfügen|Schließen|Übernehmen|Erstellen|Speichern|Ja)$/.test(x.textContent.trim())))
      : bs.find((x) => /^(Abbrechen|Schließen|Nein)$/.test(x.textContent.trim()) || x.getAttribute('aria-label') === 'Schließen');
    return b ? Object.assign(sw.rect(b), { name: b.textContent.trim() }) : null;
  };
  const hash = (s) => { let h = 2166136261; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return (h >>> 0).toString(16) + ':' + s.length; };
  // The document: every XML part (parsed once, so a part brought back by Strg+Z is compared as written).
  sw.fp = () => {
    const a = window.__present.app, d = a.deck, p = d.pptx;
    // Only the parts the presentation reaches (a slide taken back by Strg+Z may stay in the package unwritten).
    // docProps: written by every save (the time of the last change), not by the option.
    const keys = Array.from(p.reachable()).filter((k) => /\.(xml|rels)$/.test(k) && !k.startsWith('docProps/')).sort();
    const out = { _slides: d.slides.map((s) => s.path).join(','), _docs: a.docs.length };
    for (const k of keys) {
      if (k.endsWith('.rels')) continue;
      try { p.doc(k); out[k] = hash(p.text(k)); } catch (e) { out[k] = 'unreadable'; }
      try { const r = p.rels(k); out[k + '.rels'] = hash(r.xml()); } catch (e) { /* none */ }
    }
    for (const k of p.reachable()) if (!/\.(xml|rels)$/.test(k) && p.files[k]) out[k] = 'bin:' + p.files[k].length;
    return out;
  };
  // The UI: what a student sees change apart from the document.
  sw.ui = () => {
    const a = window.__present.app, st = window.__present.state();
    const pane = Array.from(document.querySelectorAll('aside')).filter((x) => !x.hidden && vis(x)).map((x) => (x.querySelector('.tp-title, h2, .title') || x).textContent.trim().slice(0, 40));
    return { view: a.view, master: a.masterMode ? a.masterKind : '', show: !!a.show, zoom: Math.round(a.editor.zoom * 100) / 100, slide: st.slide, slides: st.slides,
      sel: JSON.stringify(st.selection), tab: (document.querySelector('.tabrow .tab[aria-selected=true]') || {}).textContent || '', pane: pane.join('|'),
      body: document.body.className + '|' + (document.body.dataset.view || ''), notes: !!(document.querySelector('.notes-wrap') && vis(document.querySelector('.notes-wrap'))),
      backstage: !!document.querySelector('.backstage:not([hidden])') && vis(document.querySelector('.backstage')), docs: a.docs.length, title: document.title,
      ovl: sw.overlays().map((c) => c.className || c.tagName).join('|'), dlg: sw.dialogs().map((d) => d.getAttribute('aria-label') || '').join('|'),
      n: document.querySelectorAll('.present-main *').length, grid: !!document.querySelector('.grid, .gridlines, .ruler'), tool: a.editor.tool ? a.editor.tool.kind : '',
      clip: sw.clip.text.length + ':' + (a.clipboard ? a.clipboard.text.length + '/' + a.clipboard.pastes : '') + ':' + (a.slideClipboard ? a.slideClipboard.length : ''),
      painter: !!a.painter, ext: sw.ext.length };
  };
  // A ResizeObserver that settles in a later frame is reported by the browser as an error event; it is not one of ours.
  sw.pane = () => Array.from(document.querySelectorAll('aside.taskpane')).find((x) => !x.hidden && vis(x)) || null;
  sw.paneInfo = () => { const a = sw.pane(); if (!a) return null; return { title: (a.querySelector('.tp-title') || {}).textContent || '',
    sets: Array.from(a.querySelectorAll('.tp-set')).map((b) => b.textContent.replace('˅', '').trim()), tabs: Array.from(a.querySelectorAll('.tp-tab')).map((b) => b.getAttribute('aria-label')) }; };
  sw.paneRect = (sel, i) => { const a = sw.pane(); const e = a && Array.from(a.querySelectorAll(sel)).filter(vis)[i || 0]; if (!e) return null; e.scrollIntoView({ block: 'nearest' }); return sw.rect(e); };
  const paneEls = () => { const a = sw.pane(); return a ? Array.from(a.querySelectorAll('.tp-body input, .tp-body select, .tp-body button:not(.tp-sec-head)')).filter(vis) : []; };
  sw.paneControls = () => paneEls().map((e, i) => {
    const line = e.closest('.tp-line, .tp-row, label');
    const lab = e.closest('label') ? e.closest('label').textContent.trim() : ((line && line.querySelector('.tp-label')) || {}).textContent || label(e);
    const sec = (e.closest('.tp-sec') && e.closest('.tp-sec').querySelector('.tp-sec-head') || {}).textContent || '';
    return { i, section: sec.replace(/^[◢▷]/, '').trim(), name: (lab || '').trim().slice(0, 60) + (e.getAttribute('aria-label') && e.tagName === 'BUTTON' && !e.classList.contains('tp-color') ? ' [' + e.getAttribute('aria-label') + ']' : ''),
      tag: e.tagName.toLowerCase(), type: e.type || '', cls: e.className, disabled: !!e.disabled, value: e.value || '' };
  });
  sw.paneControlRect = (i) => { const e = paneEls()[i]; if (!e) return null; e.scrollIntoView({ block: 'center' }); const r = sw.rect(e); r.w = e.getBoundingClientRect().width; r.left = e.getBoundingClientRect().left; return r; };
  sw.ver = () => { const a = window.__present.app; return a.deck.version + ':' + a.docs.length + ':' + (a.active && a.active.id); };
  sw.take = () => { const e = sw.errs.filter((x) => !x.includes('ResizeObserver loop')); sw.errs = []; return e; };
  return 'ok';
})()
"""

# ----- CDP ------------------------------------------------------------------------------------------------------

KEYS = {"Esc": ("Escape", "Escape", 27), "Enter": ("Enter", "Enter", 13), "Tab": ("Tab", "Tab", 9), "Entf": ("Delete", "Delete", 46),
        "Rücktaste": ("Backspace", "Backspace", 8), "Leertaste": (" ", "Space", 32), "Pos1": ("Home", "Home", 36), "Ende": ("End", "End", 35),
        "Bild auf": ("PageUp", "PageUp", 33), "Bild ab": ("PageDown", "PageDown", 34), "Nach-oben": ("ArrowUp", "ArrowUp", 38),
        "Nach-unten": ("ArrowDown", "ArrowDown", 40), "Nach-links": ("ArrowLeft", "ArrowLeft", 37), "Nach-rechts": ("ArrowRight", "ArrowRight", 39),
        ">": (">", "IntlBackslash", 226), "<": ("<", "IntlBackslash", 226), "-": ("-", "Minus", 189), "#": ("#", "Backslash", 191),
        "+": ("+", "BracketRight", 187), "0": ("0", "Digit0", 48)}
MODS = {"Alt": 1, "Strg": 2, "Umschalt": 8}


class Page:
    def __init__(self, ws_url: str, proc: subprocess.Popen, kind: str):
        self.ws = websocket.create_connection(ws_url, suppress_origin=True, timeout=180)
        self.proc, self.kind, self.n = proc, kind, 0
        self.events: list[dict] = []
        self.cdp("Runtime.enable")
        self.cdp("Emulation.setFocusEmulationEnabled", {"enabled": True})

    def cdp(self, method: str, params: dict | None = None):
        self.n += 1
        self.ws.send(json.dumps({"id": self.n, "method": method, "params": params or {}}))
        while True:
            m = json.loads(self.ws.recv())
            if m.get("id") == self.n:
                if "error" in m:
                    raise RuntimeError(f"{method}: {m['error']}")
                return m.get("result")
            if m.get("method") == "Runtime.exceptionThrown":
                d = m["params"]["exceptionDetails"]
                self.events.append({"exception": (d.get("exception", {}).get("description") or d.get("text", ""))[:600]})

    def js(self, expr: str):
        r = self.cdp("Runtime.evaluate", {"expression": expr, "returnByValue": True, "awaitPromise": True})
        if "exceptionDetails" in r:
            raise RuntimeError(json.dumps(r["exceptionDetails"])[:1500])
        return r["result"].get("value")

    def mouse(self, x: float, y: float, button: str = "left", count: int = 1) -> None:
        self.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x, "y": y})
        for t in ("mousePressed", "mouseReleased"):
            self.cdp("Input.dispatchMouseEvent", {"type": t, "x": x, "y": y, "button": button, "buttons": 1 if button == "left" else 2, "clickCount": count})

    def hover(self, x: float, y: float) -> None:
        self.cdp("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": x, "y": y})

    def key(self, combo: str) -> None:
        parts = combo.split("+") if combo not in ("+",) else ["+"]
        if combo.endswith("++"):
            parts = combo[:-2].split("+") + ["+"]
        mods = sum(MODS.get(p, 0) for p in parts[:-1])
        k = parts[-1]
        if k in KEYS:
            key, code, vk = KEYS[k]
        elif len(k) == 1 and k.isalpha():
            key, code, vk = (k.upper() if mods & 8 else k.lower()), f"Key{k.upper()}", ord(k.upper())
        elif len(k) == 1 and k.isdigit():
            key, code, vk = k, f"Digit{k}", ord(k)
        elif k.startswith("F") and k[1:].isdigit():
            key, code, vk = k, k, 111 + int(k[1:])
        else:
            raise ValueError(f"key {combo}")
        if k == ">":
            mods |= 8
        down = {"type": "rawKeyDown", "key": key, "code": code, "windowsVirtualKeyCode": vk, "modifiers": mods}
        self.cdp("Input.dispatchKeyEvent", down)
        if len(key) == 1 and not mods & 3:
            self.cdp("Input.dispatchKeyEvent", {"type": "char", "text": key, "key": key, "code": code, "modifiers": mods})
        self.cdp("Input.dispatchKeyEvent", {"type": "keyUp", "key": key, "code": code, "windowsVirtualKeyCode": vk, "modifiers": mods})

    def type(self, text: str) -> None:
        self.cdp("Input.insertText", {"text": text})

    def close(self) -> None:
        try:
            self.ws.close()
        except Exception:  # noqa: BLE001
            pass
        if self.kind == "exe":
            self.proc.kill()
        else:
            subprocess.run(["powershell", "-NoProfile", "-Command",
                            "Get-CimInstance Win32_Process -Filter \"Name='msedge.exe'\" | Where-Object { $_.CommandLine -like '*" + self.profile + "*' }"
                            " | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"], capture_output=True)


def _ws(port: int, pred, deadline: float) -> str:
    while True:
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=2) as r:
                pages = [p for p in json.load(r) if p.get("type") == "page" and pred(p)]
            if pages:
                return pages[0]["webSocketDebuggerUrl"]
        except OSError:
            pass
        if time.perf_counter() > deadline:
            raise TimeoutError("DevTools endpoint not reachable")
        time.sleep(0.2)


def start(dev: bool, url: str, port: int) -> Page:
    if dev:
        prof = Path(tempfile.gettempdir()) / f"raabe-present-sweep-profile-{port}"
        proc = subprocess.Popen([EDGE, "--headless=new", "--inprivate", f"--remote-debugging-port={port}", f"--user-data-dir={prof}", "--window-size=2560,1300",
                                 "--force-device-scale-factor=1", "--no-first-run", "--disable-sync", "about:blank"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        page = Page(_ws(port, lambda p: True, time.perf_counter() + 40), proc, "dev")
        page.profile = prof.name
        # A built page (vite preview) exposes the test API only in feel runs, as the exe does.
        page.cdp("Page.enable")
        page.cdp("Page.addScriptToEvaluateOnNewDocument", {"source": "window.__FEEL = 'present';"})
        page.cdp("Page.navigate", {"url": url})
    else:
        sys.path.insert(0, str(HERE))
        from feellab import winapi
        from feellab.targets import _pid
        exe = Path(os.environ.get("RAABE_PRESENT_EXE") or ROOT / "apps/raabe-present/src-tauri/target/release/raabe-present.exe")
        rec = Path(tempfile.gettempdir()) / "raabe-present-feel-recovery"
        for f in rec.glob("*") if rec.exists() else []:
            f.unlink(missing_ok=True)
        proc = subprocess.Popen([str(exe), "--feel", "present"], env=dict(os.environ, RAABE_FEEL_DEBUG_PORT=str(port)))
        hwnd = winapi.wait_for_window(lambda h: winapi.window_title(h).endswith(" - RaabePresent") and _pid(h) == proc.pid, 40)
        # Wide enough that no group of the ribbon collapses (2560 CSS px at 150 %); never activated.
        winapi.place_quietly(hwnd, 0, 0, 3840, 2000)
        page = Page(_ws(port, lambda p: True, time.perf_counter() + 40), proc, "exe")
    for _ in range(300):
        if page.js("typeof window.__present === 'object' && !!window.__present.app"):
            break
        time.sleep(0.1)
    install(page)
    return page


def install(page: Page) -> None:
    """The helpers, the stand-in dialogs and (in a browser) the prepared files, into the page."""
    dev = page.kind == "dev"
    files = fixtures()
    page.js(HELPERS)
    page.js(f"window.__sw.files = {json.dumps(files)}")
    if dev:
        # The browser's stand-in platform reads only files it was given: the prepared files are handed to it.
        import base64
        blobs = {v: base64.b64encode(Path(v).read_bytes()).decode() for k, v in files.items() if k != "dir" and Path(v).exists()}
        page.js("((blobs) => { const P = window.__present.app.platform, orig = P.readFile.bind(P), got = {};"
                " for (const [k, v] of Object.entries(blobs)) { const s = atob(v), b = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i); got[k] = b; }"
                " P.readFile = async (p) => got[p] || orig(p); })(" + json.dumps(blobs) + ")")


def fixtures() -> dict:
    """Files the stand-in dialogs answer with: two pictures, a presentation, a sound, the save folder (TEMP)."""
    TMP.mkdir(parents=True, exist_ok=True)
    import struct
    import zlib

    def png(w: int, h: int, rgb: tuple[int, int, int]) -> bytes:
        raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
        def chunk(t: bytes, d: bytes) -> bytes:
            return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
        return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")
    (TMP / "bild1.png").write_bytes(png(240, 150, (50, 120, 200)))
    (TMP / "bild2.png").write_bytes(png(160, 160, (200, 80, 40)))
    # A sound: one second of a 440 Hz tone (WAV).
    import math
    n = 8000
    data = bytes(int(128 + 60 * math.sin(2 * math.pi * 440 * i / 8000)) for i in range(n))
    (TMP / "gliederung.txt").write_text("Erste Folie" + chr(10) + chr(9) + "Ein Punkt" + chr(10) + chr(9) + "Noch ein Punkt" + chr(10) + "Zweite Folie" + chr(10), encoding="utf-8", newline=chr(10))
    (TMP / "ton.wav").write_bytes(b"RIFF" + struct.pack("<I", 36 + n) + b"WAVEfmt " + struct.pack("<IHHIIHH", 16, 1, 1, 8000, 8000, 1, 8) + b"data" + struct.pack("<I", n) + data)
    pptx = ROOT / "packages/ooxml/test/fixtures"
    cand = sorted(pptx.glob("*.pptx"))
    src = next((p for p in cand if "basic" in p.name or "layout" in p.name), cand[0] if cand else None)
    if src:
        (TMP / "vorlage.pptx").write_bytes(src.read_bytes())
    return {"png": str(TMP / "bild1.png").replace("\\", "/"), "png2": str(TMP / "bild2.png").replace("\\", "/"), "wav": str(TMP / "ton.wav").replace("\\", "/"),
            "mp4": str(TMP / "bild1.png").replace("\\", "/"), "txt": str(TMP / "gliederung.txt").replace("\\", "/"), "pptx": str(TMP / "vorlage.pptx").replace("\\", "/"), "dir": str(TMP).replace("\\", "/")}


# ----- one option -----------------------------------------------------------------------------------------------

def settle(page: Page, t: float = 0.25) -> None:
    time.sleep(t)


class Sweeper:
    def __init__(self, page: Page, out_dir: Path):
        self.page = page
        self.out_dir = out_dir
        self.saved = 0
        # The document right after a fresh start, per starting point (made alike each time): its parts' hashes.
        self.base: dict[str, tuple[str, dict]] = {}
        self.ctx = ""

    def recover(self) -> None:
        """After a harness error: the page may have been loaded anew (helpers gone) or wait in a dialog."""
        p = self.page
        try:
            for _ in range(300):
                if p.js("typeof window.__present === 'object' && !!window.__present.app"):
                    break
                time.sleep(0.1)
            if not p.js("!!window.__sw"):
                install(p)
                self.base.clear()
            for _ in range(3):
                p.key("Esc")
                time.sleep(0.1)
        except Exception as e:  # noqa: BLE001
            print("recover failed", str(e)[:200], flush=True)

    def reset(self, ctx: str, tab: str | None) -> bool:
        p = self.page
        if not p.js("!!window.__sw"):
            self.recover()
        for _ in range(2):
            if p.js("window.__sw.overlays().length + window.__sw.dialogs().length"):
                p.key("Esc")
                time.sleep(0.12)
        p.js("(() => { for (const d of window.__sw.overlays()) d.remove(); })()")
        p.js(f"window.__present.sweepContext({json.dumps(ctx)})")
        self.ctx = ctx
        p.hover(5, 5)
        time.sleep(0.12)
        p.js("window.__sw.take()")
        p.events.clear()
        if tab:
            r = p.js(f"window.__sw.tabRect({json.dumps(tab)})")
            if not r:
                return False
            p.mouse(r["x"], r["y"])
            time.sleep(0.15)
        return True

    def observe(self, before_fp: dict, before_ui: dict, v0: str = "") -> dict:
        p = self.page
        # The parts are read only when the document says it changed (its version or the open documents).
        fp = p.js("window.__sw.fp()") if not v0 or p.js("window.__sw.ver()") != v0 else before_fp
        ui = p.js("window.__sw.ui()")
        changed = sorted(k for k in set(fp) | set(before_fp) if fp.get(k) != before_fp.get(k))
        uichg = sorted(k for k in ui if ui[k] != before_ui.get(k))
        errs = p.js("window.__sw.take()") + [e["exception"] for e in p.events]
        p.events.clear()
        return {"doc": changed, "ui": {k: [before_ui.get(k), ui[k]] for k in uichg if k != "n"} | ({"n": True} if "n" in uichg else {}), "errs": errs, "fp": fp}

    def save_check(self, name: str) -> dict:
        """The changed presentation as written: a zip whose XML parts all parse."""
        import base64
        p = self.page
        try:
            b64 = p.js("window.__present.bytes()")
        except RuntimeError as e:
            return {"ok": False, "why": f"write failed: {str(e)[:300]}"}
        data = base64.b64decode(b64)
        self.saved += 1
        f = self.out_dir / f"{self.saved:04d}.pptx"
        f.write_bytes(data)
        try:
            z = zipfile.ZipFile(io.BytesIO(data))
            for n in z.namelist():
                if n.endswith(".xml") or n.endswith(".rels"):
                    minidom.parseString(z.read(n))
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "why": f"{type(e).__name__}: {str(e)[:200]}", "file": f.name}
        return {"ok": True, "file": f.name}

    def undo_check(self, before_fp: dict) -> dict:
        p = self.page
        # Out of any dialog, menu or mode first (Esc), then Strg+Z as a student presses it.
        for _ in range(2):
            if p.js("window.__sw.overlays().length + window.__sw.dialogs().length"):
                p.key("Esc")
                time.sleep(0.1)
        p.key("Strg+Z")
        settle(p, 0.3)
        fp = p.js("window.__sw.fp()")
        left = sorted(k for k in set(fp) | set(before_fp) if fp.get(k) != before_fp.get(k))
        return {"ok": not left, "left": left[:8]}

    def baseline(self) -> tuple[str, dict]:
        """The document's state now; taken from the cache when the document is the fresh start of this starting point."""
        p = self.page
        v = p.js("window.__sw.ver()")
        hit = self.base.get(self.ctx)
        if hit and hit[0].split(":")[0] == v.split(":")[0] and hit[0].split(":")[1] == v.split(":")[1]:
            return v, hit[1]
        fp = p.js("window.__sw.fp()")
        if self.ctx not in self.base:
            self.base[self.ctx] = (v, fp)
        return v, fp

    def act(self, opt: dict, trigger) -> dict:
        """Runs `trigger` (the user's action) and checks what follows: errors, effect, dialogs (Abbrechen and OK), undo, saving."""
        p = self.page
        v0, before_fp = self.baseline()
        before_ui = p.js("window.__sw.ui()")
        res: dict = {}
        try:
            trigger()
        except Exception as e:  # noqa: BLE001
            return {"status": "fail", "why": f"trigger: {str(e)[:300]}"}
        settle(p)
        ob = self.observe(before_fp, before_ui, v0)
        res["errs"] = ob["errs"]
        res["doc"], res["ui"] = ob["doc"][:12], ob["ui"]
        dlg = p.js("window.__sw.dialogInfo()")
        if dlg:
            res["dialog"] = dlg[-1]
            # OK with the dialog's own values (a student who only looks and confirms).
            b = p.js("window.__sw.dialogButton('ok')")
            if b:
                p.mouse(b["x"], b["y"])
                settle(p, 0.5)
                ob2 = self.observe(before_fp, before_ui, v0)
                res["errs"] += ob2["errs"]
                res["after_ok"] = {"doc": ob2["doc"][:12], "ui": ob2["ui"], "dialog_left": bool(p.js("window.__sw.dialogs().length"))}
                ob = ob2
        doc_changed = bool(ob["doc"])
        new_doc = p.js("window.__sw.ver()").split(":")[1:] != v0.split(":")[1:]
        if new_doc:
            # Another presentation is open now (Neu, Öffnen, a template): nothing to undo in this one.
            res["new_document"] = True
            res["ui"]["document"] = True
            doc_changed = False
        if doc_changed:
            res["save"] = self.save_check(opt.get("name", ""))
            res["undo"] = self.undo_check(before_fp)
        effect = doc_changed or bool(res["ui"]) or bool(dlg) or bool(res.get("after_ok", {}).get("ui"))
        fails = []
        if res["errs"]:
            fails.append("errors")
        if doc_changed and not res["undo"]["ok"]:
            fails.append("undo")
        if doc_changed and not res["save"]["ok"]:
            fails.append("save")
        if res.get("after_ok", {}).get("dialog_left"):
            fails.append("dialog stays open after OK")
        res["status"] = "fail" if fails else ("pass" if effect else "no-effect")
        if fails:
            res["why"] = ", ".join(fails)
        return res


def click_at(page: Page, r: dict) -> None:
    page.mouse(r["x"], r["y"])


# ----- the walk ------------------------------------------------------------------------------------------------

def walk_ribbon(sw: Sweeper, ctx: str, tab: str, args, results: list, inventory: list, deferred: set[str]) -> None:
    p = sw.page
    if not sw.reset(ctx, tab):
        results.append({"ctx": ctx, "tab": tab, "status": "fail", "why": "tab not shown"})
        return
    controls = p.js("window.__sw.controls()")
    for c in controls:
        c.update({"ctx": ctx, "tab": tab})
        inventory.append(dict(c))
    for c in controls:
        try:
            if args.only and args.only.lower() not in c["name"].lower():
                continue
            oid = f"{tab} / {c['group']} / {c['name']}"
            base = {"id": oid, "ctx": ctx, "tab": tab, "group": c["group"], "name": c["name"], "kind": c["kind"]}
            if c["disabled"]:
                base["status"] = "deferred" if "Zurückgestellt" in c["tip"] else ("unbuilt" if "Noch nicht verfügbar" in c["tip"] else "disabled")
                results.append(base)
                continue
            sw.reset(ctx, tab)
            r = p.js(f"window.__sw.findControl({json.dumps(c['key'])})")
            if not r:
                results.append(base | {"status": "fail", "why": "control not found after reset"})
                continue
            if c["kind"] in ("combo-input", "spinner-input"):
                name = c["name"]
                val = next((v for k, v in TYPED.items() if k.lower() in name.lower()), None)
                if val is None:
                    results.append(base | {"status": "skipped", "why": "no value for this field"})
                    continue

                def trig(r=r, val=val):
                    p.mouse(r["x"], r["y"])
                    p.key("Strg+A")
                    p.type(val)
                    p.key("Enter")
                res = sw.act(base, trig)
                results.append(base | res | {"typed": val})
                print(f"{res['status']:9} {oid} = {val}  {res.get('why', '')}", flush=True)
                continue
            res = sw.act(base, lambda r=r: click_at(p, r))
            items = p.js("window.__sw.items()") if not res.get("dialog") else []
            opens_menu = c["kind"] in ("dropdown", "split-arrow", "color-arrow", "gallery-more", "combo-arrow", "combo-face", "collapsed-group") and items
            if opens_menu:
                res["status"] = "pass" if not res["errs"] else "fail"
                res["items"] = [i["name"] for i in items]
            results.append(base | res)
            print(f"{res['status']:9} {oid}  {res.get('why', '')}{'  [' + str(len(items)) + ' items]' if opens_menu else ''}", flush=True)
            if opens_menu and not args.no_items:
                walk_items(sw, ctx, tab, c, base, items, args, results, inventory, depth=1)
        except Exception as e:  # noqa: BLE001
            results.append({"id": str(locals().get("oid", "?")), "status": "harness-error", "why": str(e)[:400]})
            print("harness-error", str(e)[:200], flush=True)
            sw.recover()


SAMPLE = 12


def pick(items: list, limit: int) -> list:
    """All entries, or for a long gallery or palette (shapes, colours) the first, the last and every n-th one (about
    SAMPLE of them: every entry of such a list goes through the same code)."""
    if len(items) <= limit:
        return items
    step = max(1, len(items) // SAMPLE)
    keep = {0, 1, 2, len(items) - 1, len(items) - 2}
    keep |= set(range(0, len(items), step))
    return [it for i, it in enumerate(items) if i in keep]


def open_menu(sw: Sweeper, ctx: str, tab: str, c: dict, chain: list[str]) -> bool:
    p = sw.page
    sw.reset(ctx, tab)
    r = p.js(f"window.__sw.findControl({json.dumps(c['key'])})")
    if not r:
        return False
    p.mouse(r["x"], r["y"])
    time.sleep(0.25)
    for path in chain:
        ir = p.js(f"window.__sw.itemRect({json.dumps(path)})")
        if not ir:
            return False
        p.mouse(ir["x"], ir["y"])
        time.sleep(0.3)
    return True


def walk_items(sw: Sweeper, ctx: str, tab: str, c: dict, base: dict, items: list, args, results: list, inventory: list, depth: int, chain: list[str] | None = None) -> None:
    p = sw.page
    chain = chain or []
    for it in pick(items, args.max_items):
        try:
            oid = f"{base['id']} > {' > '.join([*base.get('chain_names', []), it['name']])}"
            entry = {"id": oid, "ctx": ctx, "tab": tab, "group": base["group"], "name": it["name"], "kind": "menu-item" if not it["sub"] else "submenu", "parent": base["id"]}
            inventory.append({"ctx": ctx, "tab": tab, "group": base["group"], "name": it["name"], "kind": entry["kind"], "parent": base["id"], "disabled": it["disabled"]})
            if it["disabled"]:
                results.append(entry | {"status": "disabled"})
                continue
            if it["tag"] in ("input", "select"):
                results.append(entry | {"status": "skipped", "why": f"{it['tag']} in a popup (pane sweep)"})
                continue
            if not open_menu(sw, ctx, tab, c, chain):
                results.append(entry | {"status": "fail", "why": "menu did not open again"})
                continue
            # The entry is found by its place among the open popup's entries, checked by its name.
            ir = p.js(f"window.__sw.itemRect({json.dumps(it['path'])})")
            if not ir or ir["name"] != it["name"]:
                # Menus whose entries change with the document (the list of open windows): by name.
                now = p.js("window.__sw.items()")
                m = next((x for x in now if x["name"] == it["name"]), None)
                ir = m and p.js(f"window.__sw.itemRect({json.dumps(m['path'])})")
                if not ir:
                    results.append(entry | {"status": "fail", "why": "entry not found when the menu opened again"})
                    continue
            res = sw.act(entry, lambda ir=ir: click_at(p, ir))
            sub = p.js("window.__sw.items()") if it["sub"] and not res.get("dialog") else []
            if it["sub"]:
                res["status"] = "pass" if (sub and not res["errs"]) else ("fail" if res["errs"] else res["status"])
                res["items"] = [x["name"] for x in sub]
            results.append(entry | res)
            print(f"{res['status']:9} {oid}  {res.get('why', '')}", flush=True)
            if it["sub"] and sub and depth < 2:
                # Only the submenu's own entries (the parent menu's are still open beside it).
                parent_roots = {x["path"].split(":")[0] for x in items}
                own = [x for x in sub if x["path"].split(":")[0] not in parent_roots] or sub[len(items):]
                walk_items(sw, ctx, tab, c, base | {"chain_names": [*base.get("chain_names", []), it["name"]]}, own, args, results, inventory, depth + 1, [*chain, it["path"]])
        except Exception as e:  # noqa: BLE001
            results.append({"id": str(locals().get("oid", "?")), "status": "harness-error", "why": str(e)[:400]})
            print("harness-error", str(e)[:200], flush=True)
            sw.recover()


def walk_context_menus(sw: Sweeper, args, results: list, inventory: list) -> None:
    p = sw.page
    for name, ctx, where in CONTEXT_MENUS:
        def rclick(ctx=ctx, where=where):
            sw.reset(ctx, None)
            pt = p.js("""((where) => { const a = window.__present.app, st = window.__present.state();
              const sel = st.selection.ids && st.selection.ids[0];
              if (where === 'thumb') { const r = window.__present.thumbRect(1); return { x: r.x + r.w / 2, y: r.y + r.h / 2 }; }
              if (where === 'notes') { const r = document.querySelector('.notes').getBoundingClientRect(); return { x: r.left + 40, y: r.top + 10 }; }
              if (where === 'slide') { const r = window.__present.slideRect(); return { x: r.x + r.w * 0.9, y: r.y + r.h * 0.9 }; }
              const r = window.__present.shapeRect(sel); return { x: r.x + r.w / 2, y: r.y + r.h / 2 }; })(""" + json.dumps(where) + ")")
            p.mouse(pt["x"], pt["y"], "right")
            time.sleep(0.3)
        rclick()
        items = p.js("window.__sw.items()")
        errs = p.js("window.__sw.take()")
        base = {"id": f"Kontextmenü {name}", "ctx": ctx, "tab": "Kontextmenü", "group": name, "name": name, "kind": "context-menu"}
        results.append(base | {"status": "pass" if items and not errs else "fail", "items": [i["name"] for i in items], "errs": errs})
        print(f"{'pass' if items else 'fail':9} Kontextmenü {name} [{len(items)} items]", flush=True)
        for it in pick(items, args.max_items):
            try:
                oid = f"Kontextmenü {name} > {it['name']}"
                entry = {"id": oid, "ctx": ctx, "tab": "Kontextmenü", "group": name, "name": it["name"], "kind": "menu-item" if not it["sub"] else "submenu"}
                inventory.append({"ctx": ctx, "tab": "Kontextmenü", "group": name, "name": it["name"], "kind": entry["kind"], "disabled": it["disabled"]})
                if it["disabled"]:
                    results.append(entry | {"status": "disabled"})
                    continue
                rclick()
                ir = p.js(f"window.__sw.itemRect({json.dumps(it['path'])})")
                if not ir or ir["name"] != it["name"]:
                    results.append(entry | {"status": "fail", "why": "entry not found when the menu opened again"})
                    continue
                res = sw.act(entry, lambda ir=ir: click_at(p, ir))
                if it["sub"]:
                    sub = p.js("window.__sw.items()")
                    res["items"] = [x["name"] for x in sub]
                    res["status"] = "pass" if sub and not res["errs"] else res["status"]
                results.append(entry | res)
                print(f"{res['status']:9} {oid}  {res.get('why', '')}", flush=True)
            except Exception as e:  # noqa: BLE001
                results.append({"id": str(locals().get("oid", "?")), "status": "harness-error", "why": str(e)[:400]})
                print("harness-error", str(e)[:200], flush=True)
                sw.recover()


def walk_shortcuts(sw: Sweeper, args, results: list, inventory: list) -> None:
    p = sw.page
    cmds = p.js("window.__present.app.commands.all().filter((c) => c.shortcut).map((c) => ({ id: c.id, label: c.label, shortcut: c.shortcut, on: !!c.run }))")
    for c in cmds:
        try:
            sc = c["shortcut"]
            if sc in ("Strg+W", "Strg+N", "Strg+O") or sc.startswith("Tab") or sc.startswith("Umschalt+Tab"):
                ctx = "slides"
            else:
                ctx = "shape-text" if c["id"].startswith(("font.", "para.", "edit.")) else "shape"
            oid = f"Tastenkombination {sc} ({c['label']})"
            entry = {"id": oid, "ctx": ctx, "tab": "Tastatur", "group": "Tastenkombinationen", "name": c["label"], "kind": "shortcut", "shortcut": sc}
            inventory.append({"ctx": ctx, "tab": "Tastatur", "group": "Tastenkombinationen", "name": f"{c['label']} ({sc})", "kind": "shortcut", "disabled": not c["on"]})
            if args.only and args.only.lower() not in c["label"].lower():
                continue
            sw.reset(ctx, None)
            if sc in ("Strg+V", "Strg+Alt+V"):
                # Something to paste: the selection copied first (Strg+C).
                p.key("Strg+C")
                time.sleep(0.2)
            if sc in ("Strg+Z", "Strg+Y"):
                # Something to undo: the text made bold first.
                p.key("Strg+Umschalt+F")
                time.sleep(0.2)
                if sc == "Strg+Y":
                    p.key("Strg+Z")
                    time.sleep(0.2)
            try:
                res = sw.act(entry, lambda sc=sc: p.key(sc))
            except ValueError as e:
                res = {"status": "skipped", "why": str(e)}
            if sc in ("Strg+Z", "Strg+Y") and res.get("undo"):
                res.pop("undo")
                res["status"] = "pass" if not res["errs"] else "fail"
            results.append(entry | res)
            print(f"{res['status']:9} {oid}  {res.get('why', '')}", flush=True)
        except Exception as e:  # noqa: BLE001
            results.append({"id": str(locals().get("oid", "?")), "status": "harness-error", "why": str(e)[:400]})
            print("harness-error", str(e)[:200], flush=True)
            sw.recover()


def walk_backstage_status(sw: Sweeper, args, results: list, inventory: list) -> None:
    """Datei (each page and its buttons) and the status bar's controls."""
    p = sw.page
    sw.reset("slides", None)
    r = p.js("(() => { const t = document.querySelector('.tabrow .tab'); return window.__sw.rect(t); })()")
    p.mouse(r["x"], r["y"])
    time.sleep(0.4)
    pages = p.js("Array.from(document.querySelectorAll('.backstage nav button, .backstage .bs-nav button')).filter((b) => b.getBoundingClientRect().width > 0).map((b, i) => ({ i, name: window.__sw.label(b) }))")
    for pg in pages:
        oid = f"Datei / {pg['name']}"
        entry = {"id": oid, "ctx": "slides", "tab": "Datei", "group": "Backstage", "name": pg["name"], "kind": "backstage-page"}
        inventory.append({"ctx": "slides", "tab": "Datei", "group": "Backstage", "name": pg["name"], "kind": "backstage-page", "disabled": False})

        def go(pg=pg):
            sw.reset("slides", None)
            r = p.js("(() => { const t = document.querySelector('.tabrow .tab'); return window.__sw.rect(t); })()")
            p.mouse(r["x"], r["y"])
            time.sleep(0.4)
            b = p.js(f"(() => {{ const b = Array.from(document.querySelectorAll('.backstage nav button, .backstage .bs-nav button')).filter((b) => b.getBoundingClientRect().width > 0)[{pg['i']}]; return b ? window.__sw.rect(b) : null; }})()")
            if b:
                p.mouse(b["x"], b["y"])
                time.sleep(0.3)
        go()
        btns = p.js("Array.from(document.querySelectorAll('.backstage main button, .backstage .bs-page button, .backstage .page button')).filter((b) => b.getBoundingClientRect().width > 0 && !b.disabled).map((b, i) => ({ i, name: window.__sw.label(b) }))")
        res = {"status": "pass", "buttons": [b["name"] for b in btns], "errs": p.js("window.__sw.take()")}
        if res["errs"]:
            res["status"] = "fail"
        results.append(entry | res)
        print(f"{res['status']:9} {oid} [{len(btns)} buttons]", flush=True)
        for b in pick(btns, args.max_items):
            bid = f"{oid} / {b['name']}"
            e2 = {"id": bid, "ctx": "slides", "tab": "Datei", "group": pg["name"], "name": b["name"], "kind": "backstage-button"}
            inventory.append({"ctx": "slides", "tab": "Datei", "group": pg["name"], "name": b["name"], "kind": "backstage-button", "disabled": False})
            if b["name"] in ("Schließen",) and pg["name"] in ("Schließen",):
                pass
            go()

            def trig(b=b):
                r = p.js(f"(() => {{ const e = Array.from(document.querySelectorAll('.backstage main button, .backstage .bs-page button, .backstage .page button')).filter((b) => b.getBoundingClientRect().width > 0 && !b.disabled)[{b['i']}]; return e ? window.__sw.rect(e) : null; }})()")
                if r:
                    p.mouse(r["x"], r["y"])
            res = sw.act(e2, trig)
            results.append(e2 | res)
            print(f"{res['status']:9} {bid}  {res.get('why', '')}", flush=True)
    # Status bar.
    sw.reset("slides", None)
    sb = p.js("Array.from(document.querySelectorAll('.statusbar button, .statusbar [role=slider], .status button')).filter((b) => b.getBoundingClientRect().width > 0).map((b, i) => ({ i, name: window.__sw.label(b) }))")
    for b in sb:
        bid = f"Statusleiste / {b['name']}"
        e2 = {"id": bid, "ctx": "slides", "tab": "Statusleiste", "group": "Statusleiste", "name": b["name"], "kind": "status"}
        inventory.append({"ctx": "slides", "tab": "Statusleiste", "group": "Statusleiste", "name": b["name"], "kind": "status", "disabled": False})
        sw.reset("slides", None)

        def trig(b=b):
            r = p.js(f"(() => {{ const e = Array.from(document.querySelectorAll('.statusbar button, .statusbar [role=slider], .status button')).filter((b) => b.getBoundingClientRect().width > 0)[{b['i']}]; return e ? window.__sw.rect(e) : null; }})()")
            if r:
                p.mouse(r["x"], r["y"])
        res = sw.act(e2, trig)
        results.append(e2 | res)
        print(f"{res['status']:9} {bid}  {res.get('why', '')}", flush=True)


# Task panes: opened through their ribbon control; every set, tab and section, every control in them.
PANES = [
    ("Form formatieren", "shape", "Formformat", "Objekt formatieren..."),
    ("Form formatieren (Text)", "shape-text", "Formformat", "Objekt formatieren..."),
    ("Bild formatieren", "picture", "Bildformat", None),
    ("Hintergrund formatieren", "slides", "Entwurf", "Hintergrund formatieren"),
    ("Diagramm formatieren", "chart", "Format", None),
]


def open_pane(sw: "Sweeper", ctx: str, tab: str, control: str | None, si: int = 0, ti: int = 0) -> bool:
    p = sw.page
    sw.reset(ctx, tab)
    cs = p.js("window.__sw.controls()")
    c = next((x for x in cs if (control and x["name"] == control) or (not control and x["kind"] == "launcher" and not x["disabled"])), None)
    if not c:
        return False
    r = p.js(f"window.__sw.findControl({json.dumps(c['key'])})")
    p.mouse(r["x"], r["y"])
    time.sleep(0.35)
    if not p.js("!!window.__sw.pane()"):
        return False
    for sel, i in ((".tp-set", si), (".tp-tab", ti)):
        if i:
            r = p.js(f"window.__sw.paneRect({json.dumps(sel)}, {i})")
            if r:
                p.mouse(r["x"], r["y"])
                time.sleep(0.2)
    # Every section open (a click on each closed heading, as a student opens them).
    for _ in range(30):
        r = p.js("window.__sw.paneRect('.tp-sec-head[aria-expanded=false]', 0)")
        if not r:
            break
        p.mouse(r["x"], r["y"])
        time.sleep(0.12)
    return True


def pane_action(p: "Page", c: dict, r: dict):
    """What a student does with the control: tick, pick a colour, type a number, choose the next entry, slide."""
    def go():
        if c["tag"] == "select":
            p.mouse(r["x"], r["y"])
            time.sleep(0.15)
            p.key("Nach-unten")
            p.key("Enter")
        elif c["tag"] == "input" and c["type"] == "range":
            p.mouse(r["left"] + r["w"] * 0.7, r["y"])
        elif c["tag"] == "input" and c["type"] in ("text", ""):
            v = c["value"]
            num = "".join(ch for ch in v if ch.isdigit() or ch in ",-")
            try:
                n = float(num.replace(",", ".")) if num not in ("", "-", ",") else None
            except ValueError:
                n = None
            new = ("5" if n is None else str(round(n + (10 if abs(n) >= 20 else 1), 2)).replace(".", ","))
            if "tp-text" in c["cls"]:
                new = "Sweep"
            p.mouse(r["x"], r["y"])
            p.key("Strg+A")
            p.type(new)
            p.key("Enter")
            p.key("Tab")
        elif "tp-color" in c["cls"]:
            p.mouse(r["x"], r["y"])
            time.sleep(0.25)
            items = p.js("window.__sw.items()")
            sw_items = [x for x in items if x["name"] and x["name"] != "Automatisch"]
            if sw_items:
                ir = p.js(f"window.__sw.itemRect({json.dumps(sw_items[min(14, len(sw_items) - 1)]['path'])})")
                p.mouse(ir["x"], ir["y"])
        else:
            p.mouse(r["x"], r["y"])
    return go


def walk_panes(sw: "Sweeper", args, results: list, inventory: list) -> None:
    p = sw.page
    for title, ctx, tab, control in PANES:
        if args.only and args.only.lower() not in title.lower():
            continue
        if not open_pane(sw, ctx, tab, control):
            results.append({"id": f"Bereich {title}", "ctx": ctx, "tab": "Aufgabenbereich", "group": title, "name": title, "kind": "pane", "status": "fail", "why": "pane did not open"})
            print(f"fail      Bereich {title} did not open", flush=True)
            continue
        info = p.js("window.__sw.paneInfo()")
        results.append({"id": f"Bereich {title}", "ctx": ctx, "tab": "Aufgabenbereich", "group": title, "name": info["title"], "kind": "pane", "status": "pass", "sets": info["sets"], "tabs": info["tabs"]})
        for si in range(max(1, len(info["sets"]))):
            open_pane(sw, ctx, tab, control, si, 0)
            tabs = p.js("window.__sw.paneInfo()")["tabs"]
            for ti in range(max(1, len(tabs))):
                open_pane(sw, ctx, tab, control, si, ti)
                ctrls = p.js("window.__sw.paneControls()")
                where = f"{(info['sets'] or [''])[si] if info['sets'] else ''}/{tabs[ti] if tabs else ''}"
                for c in ctrls:
                    oid = f"Bereich {title} / {where} / {c['section']} / {c['name']} ({c['tag']}{':' + c['type'] if c['type'] else ''} {c['i']})"
                    entry = {"id": oid, "ctx": ctx, "tab": "Aufgabenbereich", "group": f"{title} {where} {c['section']}", "name": c["name"], "kind": f"pane-{c['tag']}"}
                    inventory.append({"ctx": ctx, "tab": "Aufgabenbereich", "group": f"{title} {where} {c['section']}", "name": c["name"], "kind": entry["kind"], "disabled": c["disabled"]})
                    if args.only_pane and args.only_pane.lower() not in oid.lower():
                        continue
                    if c["disabled"]:
                        results.append(entry | {"status": "disabled"})
                        continue
                    try:
                        open_pane(sw, ctx, tab, control, si, ti)
                        now = p.js("window.__sw.paneControls()")
                        if c["i"] >= len(now) or now[c["i"]]["name"] != c["name"]:
                            results.append(entry | {"status": "fail", "why": "control not found when the pane opened again"})
                            continue
                        r = p.js(f"window.__sw.paneControlRect({c['i']})")
                        res = sw.act(entry, pane_action(p, c, r))
                        # A pane's controls change the document; a press that only redraws the pane (a heading) is no failure.
                        results.append(entry | res)
                        print(f"{res['status']:9} {oid}  {res.get('why', '')}", flush=True)
                    except Exception as e:  # noqa: BLE001
                        results.append(entry | {"status": "harness-error", "why": str(e)[:300]})
                        print("harness-error", oid, str(e)[:200], flush=True)
                        sw.recover()


# ----- main -----------------------------------------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("what", choices=["inventory", "sweep"])
    ap.add_argument("--dev", action="store_true")
    ap.add_argument("--url", default="http://localhost:5186/")
    ap.add_argument("--port", type=int, default=9481)
    ap.add_argument("--tab", action="append")
    ap.add_argument("--ctx", action="append")
    ap.add_argument("--only")
    ap.add_argument("--max-items", type=int, default=40)
    ap.add_argument("--no-items", action="store_true")
    ap.add_argument("--parts", default="ribbon,context,shortcuts,backstage,panes")
    ap.add_argument("--only-pane")
    ap.add_argument("--out")
    args = ap.parse_args()
    name = args.out or ("present-sweep" if args.what == "sweep" else "present-options")
    out_dir = TMP / f"saved-{name}"
    out_dir.mkdir(parents=True, exist_ok=True)
    for f in out_dir.glob("*.pptx"):
        f.unlink()
    page = start(args.dev, args.url, args.port)
    sw = Sweeper(page, out_dir)
    results: list = []
    inventory: list = []
    t0 = time.perf_counter()
    parts = args.parts.split(",")
    try:
        if args.what == "inventory":
            args.no_items = False
        if "ribbon" in parts:
            for ctx, tabs in CONTEXTS:
                if args.ctx and ctx not in args.ctx:
                    continue
                for tab in tabs:
                    if args.tab and tab not in args.tab:
                        continue
                    walk_ribbon(sw, ctx, tab, args, results, inventory, set())
        if "context" in parts and not args.tab:
            walk_context_menus(sw, args, results, inventory)
        if "shortcuts" in parts and not args.tab:
            walk_shortcuts(sw, args, results, inventory)
        if "backstage" in parts and not args.tab:
            walk_backstage_status(sw, args, results, inventory)
        if "panes" in parts and not args.tab:
            walk_panes(sw, args, results, inventory)
        defined = page.js("window.__present.inventory()")
    except Exception as e:  # noqa: BLE001
        # What was walked is written all the same.
        print("stopped:", str(e)[:400], flush=True)
        results.append({"id": "(the walk)", "status": "harness-error", "why": str(e)[:400]})
        defined = None
    finally:
        page.close()
    date = datetime.date.today().isoformat()
    counts: dict = {}
    for r in results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    summary = {"date": date, "target": "dev" if args.dev else "exe", "seconds": round(time.perf_counter() - t0), "options": len(results), "counts": counts,
               "saved_files": sw.saved}
    res_path = HERE / "results" / f"{date}-{name}.json"
    if args.what == "inventory":
        doc = {"summary": summary, "reachable": inventory, "defined": defined}
    else:
        doc = {"summary": summary, "results": [{k: v for k, v in r.items() if k != "fp"} for r in results]}
    res_path.write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps(summary, ensure_ascii=False), "->", res_path.name)


if __name__ == "__main__":
    main()
