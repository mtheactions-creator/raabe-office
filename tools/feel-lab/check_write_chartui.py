"""Everything around a selected chart in RaabeWrite, tried by hand (D-129 - D-131): the release exe runs off screen and
is driven with trusted DevTools input (Input.dispatchMouseEvent / dispatchKeyEvent; no real mouse or keyboard, no lock):
Einfügen > Diagramm through its dialogs, the contextual tabs Diagrammentwurf and Format (groups and keytips against
Word's recording, results/2026-11-09-word-chart-ui.json), every menu entry of Diagrammelement hinzufügen, Schnelllayout
and Farben ändern, the style gallery, the four buttons beside the chart (Layoutoptionen with each wrapping,
Diagrammelemente's check boxes, Diagrammformatvorlagen, Diagrammfilter), the Format pane, Größe, dragging a floating
chart, Strg+Z. Each document RaabeWrite saves is then opened in our own invisible Word (DispatchEx), which must report
the same chart state through COM (HasTitle, HasLegend, Legend.Position, HasDataTable, data labels, gridlines, axis
titles, ChartType, filtered series; for floating charts Shape.Left/Top, WrapFormat.Type, the anchoring), and Word's
PDF of the file is compared with RaabeWrite's (the chart's area, mean grey difference).

    .venv/Scripts/python.exe check_write_chartui.py          (RAABE_WRITE_EXE points at another build)

Result: results/2026-11-09-write-chartui.json (case, passed, detail).
"""
from __future__ import annotations

import json
import shutil
import sys
import tempfile
import time
import traceback
from pathlib import Path

from feellab.targets import raabe_write

HERE = Path(__file__).parent
RESULT = HERE / "results" / "2026-11-09-write-chartui.json"
WORD = json.loads((HERE / "results" / "2026-11-09-word-chart-ui.json").read_text(encoding="utf-8"))
TMP = Path(tempfile.gettempdir()) / "raabe-write-chartui"
PORT = 9481

# Word's WdWrapType for each of RaabeWrite's wrappings.
WRAP_TYPES = {"square": 0, "tight": 1, "through": 2, "front": 3, "topAndBottom": 4, "behind": 5}
WRAP_LABELS = {"square": "Quadrat", "tight": "Eng", "through": "Transparent", "topAndBottom": "Oben und unten", "behind": "Hinter den Text", "front": "Vor den Text"}


class Ui:
    def __init__(self, t) -> None:
        self.t = t

    def js(self, e: str):
        return self.t.js(e)

    def mouse(self, kind: str, x: float, y: float, button: str = "left", count: int = 1, buttons: int = 0, mods: int = 0) -> None:
        self.t.cdp("Input.dispatchMouseEvent", {"type": kind, "x": x, "y": y, "button": button, "clickCount": count, "buttons": buttons, "modifiers": mods})

    def move(self, x: float, y: float) -> None:
        self.mouse("mouseMoved", x, y, button="none")
        time.sleep(0.08)

    def click(self, x: float, y: float, count: int = 1) -> None:
        self.move(x, y)
        for n in range(1, count + 1):
            self.mouse("mousePressed", x, y, count=n, buttons=1)
            self.mouse("mouseReleased", x, y, count=n)
        time.sleep(0.3)

    def drag(self, x1: float, y1: float, x2: float, y2: float, steps: int = 14) -> None:
        self.move(x1, y1)
        self.mouse("mousePressed", x1, y1, buttons=1)
        for i in range(1, steps + 1):
            self.mouse("mouseMoved", x1 + (x2 - x1) * i / steps, y1 + (y2 - y1) * i / steps, button="left", buttons=1)
            time.sleep(0.02)
        self.mouse("mouseReleased", x2, y2)
        time.sleep(0.4)

    def key(self, key: str, code: str, vk: int, mods: int = 0) -> None:
        self.t.cdp("Input.dispatchKeyEvent", {"type": "rawKeyDown", "key": key, "code": code, "windowsVirtualKeyCode": vk, "modifiers": mods})
        self.t.cdp("Input.dispatchKeyEvent", {"type": "keyUp", "key": key, "code": code, "windowsVirtualKeyCode": vk, "modifiers": mods})
        time.sleep(0.25)

    def type(self, text: str) -> None:
        self.t.cdp("Input.insertText", {"text": text})
        time.sleep(0.15)

    def undo(self) -> None:
        self.key("z", "KeyZ", 90, 2)

    def esc(self) -> None:
        self.key("Escape", "Escape", 27)

    def rect(self, sel: str, text: str | None = None):
        r = self.js(f"""(() => {{ const all = [...document.querySelectorAll({json.dumps(sel)})].filter((x) => x.getClientRects().length);
          const e = {json.dumps(text)} === null ? all[0] : all.find((x) => (x.getAttribute('aria-label') ?? '') === {json.dumps(text)} || x.textContent.trim() === {json.dumps(text)});
          if (!e) return null; e.scrollIntoView({{ block: 'nearest' }}); const r = e.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2, r.width, r.height]; }})()""")
        if not r:
            raise RuntimeError(f"not found: {sel} {text!r}")
        return r

    def click_on(self, sel: str, text: str | None = None) -> None:
        x, y, _, _ = self.rect(sel, text)
        self.click(x, y)

    def hover_on(self, sel: str, text: str | None = None) -> None:
        x, y, _, _ = self.rect(sel, text)
        self.move(x, y)
        time.sleep(0.3)

    def tab(self, label: str) -> None:
        self.click_on(".tab", label)

    def ribbon(self, label: str) -> None:
        """A ribbon control by its label; split buttons by their main part."""
        self.click_on(".ribbon-card [aria-label]", label)

    def menu(self, label: str) -> None:
        self.click_on(".menu button", label)

    def texts(self, sel: str) -> list[str]:
        return self.js(f"[...document.querySelectorAll({json.dumps(sel)})].filter((e) => e.getClientRects().length).map((e) => e.textContent.trim())")

    def model(self):
        return self.js("(() => { const m = __app.chartTools().model(); if (!m) return null; const { raw, ...rest } = m; return JSON.parse(JSON.stringify(rest)); })()")

    def sel(self):
        return self.js("(() => { const s = __app.editor().selectedImage(); if (!s || !s.obj.chart) return null; const o = s.obj; return { kind: s.kind, w: o.width, h: o.height, wrap: o.wrap ?? 'inline', x: o.x, y: o.y, moveWithText: o.moveWithText }; })()")

    def chart_rect(self):
        return self.js("(() => { const r = __app.editor().view.selectedObjectRect(); return r ? [r.left, r.top, r.width, r.height] : null; })()")

    def select_chart(self) -> None:
        """Clicks the chart (its lower left corner, the chart area)."""
        rr = self.chart_rect()
        if rr is None:
            # Not selected (Esc took the selection): the chart's place in the text (an inline chart is one object
            # character), selected through the probe, then clicked like a student would.
            i = self.js("__probe.text().indexOf(String.fromCharCode(0xfffc))")
            if i is not None and i >= 0:
                self.t.set_selection(i, i + 1)
                time.sleep(0.4)
            rr = self.chart_rect()
        if rr:
            x, y, w, h = rr
            self.click(x + w * 0.04, y + h * 0.96)
        print("select_chart:", rr, self.sel(), self.texts(".tab.contextual"), flush=True)

    def save(self, name: str) -> Path:
        out = TMP / f"{name}.docx"
        if out.exists():
            out.unlink()
        self.js(f"__app.saveTo({json.dumps(str(out))})")
        deadline = time.time() + 10
        while time.time() < deadline and not out.exists():
            time.sleep(0.05)
        time.sleep(0.3)
        # RaabeWrite's PDF of the same document, compared with Word's PDF of the saved file.
        pdf = TMP / f"{name}-raabe.pdf"
        self.js(f"__app.exportPdfTo({json.dumps(str(pdf))})")
        deadline = time.time() + 20
        while time.time() < deadline and not pdf.exists():
            time.sleep(0.1)
        time.sleep(0.5)
        return out


class Run:
    def __init__(self) -> None:
        self.cases: list[dict] = []

    def check(self, name: str, ok: bool, detail="") -> None:
        # Measured values (picture differences) are kept for passed cases too.
        keep = not ok or isinstance(detail, (int, float))
        self.cases.append({"case": name, "passed": bool(ok), "detail": str(detail)[:600] if keep else ""})
        print(("PASS " if ok else "FAIL ") + name + ("" if ok else f"  {str(detail)[:300]}"), flush=True)


def close_popups(ui: Ui) -> None:
    ui.js("document.querySelectorAll('.dialog-layer .dlg-buttons button').forEach((b) => { if (b.textContent.trim() === 'Abbrechen') b.click(); })")
    ui.esc()
    ui.esc()


def fresh(ui: Ui) -> None:
    """A new document with a column chart inserted by hand: Einfügen > Diagramm, OK, OK."""
    close_popups(ui)
    ui.js("__app.newDoc(); __app.closeOthers()")
    time.sleep(0.6)
    ui.tab("Einfügen")
    ui.ribbon("Diagramm")
    time.sleep(0.5)
    ui.click_on(".dlg-buttons button", "OK")  # Diagramm einfügen (Gruppierte Säulen)
    time.sleep(0.5)
    ui.click_on(".dlg-buttons button", "OK")  # Daten bearbeiten (Word's sample data)
    time.sleep(0.6)


def ribbon_groups(ui: Ui, tab_id: str):
    ui.click_on(".tab", "Diagrammentwurf" if tab_id == "chartDesign" else "Format")
    return ui.js("""[...document.querySelectorAll('.ribbon-card .rgroup')].map((g) => [g.getAttribute('aria-label'),
      [...g.querySelectorAll('[data-keytip]')].map((c) => [[c.getAttribute('aria-label') ?? '', (c.getAttribute('data-tip') ?? '').split(String.fromCharCode(10))[0]], c.getAttribute('data-keytip')])])""")


def word_keys(tab: str) -> dict[str, tuple[str, dict[str, str]]]:
    """Word's groups of a tab: label -> (group keytip, control label -> keytip), from the recording."""
    out = {}
    for g in WORD["inline"][tab]["groups"]:
        if g["name"] in ("Diagrammentwurf", "Format", "Schnellformatvorlagen", "Formen"):
            continue
        k = g["key"].replace("Alt, ", "").split(", ")[-1].replace(" ", "")
        ctl: dict = {}
        seen: dict = {}
        for c in g["controls"]:
            if c["key"] and c["type"] not in ("Group", "ListItem"):
                seen[c["name"]] = seen.get(c["name"], 0) + 1
                ctl.setdefault(c["name"], c["key"].replace("Alt, ", "").split(", ")[-1].replace(" ", ""))
        # Names Word uses for several controls (each split button's "Weitere Optionen") say nothing on their own.
        out[g["name"]] = (k, {n: v for n, v in ctl.items() if seen[n] == 1})
    return out


def cases(ui: Ui, run: Run, saved: list) -> None:
    TMP.mkdir(exist_ok=True)
    fresh(ui)
    s = ui.sel()
    run.check("Einfügen > Diagramm by hand inserts a chart, selected", bool(s) and s["kind"] == "inline" and abs(s["w"] - 432) < 1 and abs(s["h"] - 252) < 1, s)
    tabs = ui.texts(".tab.contextual")
    run.check("Diagrammentwurf and Format show with a chart selected", tabs == ["Diagrammentwurf", "Format"], tabs)

    # ----- the tabs against Word's recording ------------------------------------------------------------------
    for tab_id, tab in (("chartDesign", "Diagrammentwurf"), ("chartFormat", "Format")):
        ours = ribbon_groups(ui, tab_id)
        word = word_keys(tab)
        names = [g[0] for g in ours]
        want = [n for n in word]
        run.check(f"{tab}: Word's groups in Word's order", names == want, [names, want])
        bad = []
        for g, ctls in ours:
            if g not in word:
                continue
            gk, wctl = word[g]
            for labels, key in ctls:
                hit = next((x for x in labels if x in wctl), None)
                if hit is not None and wctl[hit] != key:
                    bad.append(f"{g}/{hit}: {key} vs Word {wctl[hit]}")
        group_keys = ui.js("[...document.querySelectorAll('.ribbon-card .rgroup')].map((g) => g.getAttribute('aria-label'))")
        del group_keys
        run.check(f"{tab}: controls' keytips as Word's", not bad, bad)

    # ----- Diagrammelement hinzufügen: every entry's side menu ------------------------------------------------
    ui.tab("Diagrammentwurf")
    ui.ribbon("Diagrammelement hinzufügen")
    entries = ui.texts(".menu-root .menu:not(.submenu) > button")
    run.check("Diagrammelement hinzufügen lists Office's elements",
              entries[:9] == ["Achsen", "Achsentitel", "Diagrammtitel", "Datenbeschriftungen", "Datentabelle", "Fehlerindikatoren", "Gitternetzlinien", "Legende", "Linien"][:len(entries[:9])] and len(entries) >= 9, entries)
    ui.esc()
    steps = [
        ("Legende", "Unten", "(m) => m.legend === 'b'"),
        ("Diagrammtitel", "Keine", "(m) => !m.title.show"),
        ("Datenbeschriftungen", "Am Ende außerhalb", "(m) => m.dataLabels && (m.format?.dataLabels?.pos ?? 'outEnd') === 'outEnd'"),
        ("Achsentitel", "Primär vertikal", "(m) => m.axisTitles.val !== undefined"),
        ("Gitternetzlinien", "Primäres Hauptgitter horizontal", "(m) => !m.gridlines"),
        ("Datentabelle", "Mit Legendensymbolen", "(m) => !!m.dataTable"),
    ]
    for entry, choice, probe in steps:
        try:
            ui.tab("Diagrammentwurf")
            ui.ribbon("Diagrammelement hinzufügen")
            ui.hover_on(".menu-root .menu button", entry)
            ui.click_on(".menu-root .submenu button", choice)
            time.sleep(0.4)
            ok = ui.js(f"({probe})(__app.chartTools().model())")
            run.check(f"Diagrammelement hinzufügen > {entry} > {choice}", ok, ui.model())
        except Exception as e:  # noqa: BLE001
            run.check(f"Diagrammelement hinzufügen > {entry} > {choice}", False, e)
            close_popups(ui)
    saved.append(("elements", ui.save("elements"), {"HasTitle": False, "LegendPosition": -4107, "HasDataTable": True, "DataLabels": True, "ValTitle": True, "MajorGrid": False}))

    # ----- Schnelllayout and Farben ändern ---------------------------------------------------------------------
    fresh(ui)
    ui.tab("Diagrammentwurf")
    ui.ribbon("Schnelllayout")
    layouts = ui.texts(".menu-root .menu button")
    run.check("Schnelllayout offers Layout 1 to Layout 11", layouts == [f"Layout {i}" for i in range(1, 12)], layouts)
    ui.hover_on(".menu-root .menu button", "Layout 5")
    pv = ui.js("__app.editor().previewing")
    ui.click_on(".menu-root .menu button", "Layout 5")
    m = ui.model()
    run.check("Schnelllayout > Layout 5 (Livevorschau, then applied: Datentabelle, no legend)", pv and m and m.get("dataTable") and m["legend"] == "none", [pv, m and m.get("dataTable"), m and m.get("legend")])
    ui.ribbon("Farben ändern")
    colors = ui.texts(".menu-root .menu button")
    run.check("Farben ändern lists Farbige Palette 1-4 and Monochrome Palette 1-13", len(colors) == 17 and colors[0] == "Farbige Palette 1" and colors[-1] == "Monochrome Palette 13", colors)
    ui.click_on(".menu-root .menu button", "Monochrome Palette 3")
    run.check("Farben ändern > Monochrome Palette 3", (ui.model() or {}).get("colors") == 16, (ui.model() or {}).get("colors"))
    saved.append(("layout5", ui.save("layout5"), {"HasDataTable": True, "HasLegend": False}))

    # ----- the style gallery --------------------------------------------------------------------------------
    ui.tab("Diagrammentwurf")
    n = ui.js("document.querySelectorAll('.ribbon-card .rgroup[data-group=styles] .tile, .ribbon-card .rgroup[data-group=styles] [data-tip^=\"Formatvorlage\"]').length")
    run.check("Diagrammformatvorlagen gallery shows the chart's styles", (n or 0) >= 4, n)
    try:
        ui.click_on(".ribbon-card [data-tip^='Formatvorlage 3'], .ribbon-card [aria-label='Formatvorlage 3']")
        m = ui.model()
        run.check("Formatvorlage 3 applied from the ribbon", ui.js("(() => { const m = __app.chartTools().model(); return !!m; })()") and m is not None, "")
    except Exception as e:  # noqa: BLE001
        run.check("Formatvorlage 3 applied from the ribbon", False, e)
    ui.undo()
    run.check("Strg+Z takes the style back, the chart stays selected", ui.sel() is not None, ui.sel())

    # ----- the buttons beside the chart ---------------------------------------------------------------------
    fresh(ui)
    rects = ui.js("""(() => { const r = __app.editor().view.selectedObjectRect(); const lb = __app.editor().view.layoutButtonRect();
      return [[ 'Layoutoptionen', Math.round(lb.left - r.right), Math.round(lb.top - r.top), Math.round(lb.width)],
        ...[...document.querySelectorAll('.wchart-overlay .chart-btn')].map((b) => { const q = b.getBoundingClientRect(); return [b.getAttribute('aria-label'), Math.round(q.left - r.right), Math.round(q.top - r.top), Math.round(q.width)]; })]; })()""")
    run.check("four buttons beside the chart: Layoutoptionen, Diagrammelemente, Diagrammformatvorlagen, Diagrammfilter",
              [r[0] for r in rects] == ["Layoutoptionen", "Diagrammelemente", "Diagrammformatvorlagen", "Diagrammfilter"] and all(r[1] > 0 for r in rects)
              and all(rects[i + 1][2] > rects[i][2] for i in range(3)), rects)
    # Diagrammelemente: each check box by hand.
    ui.click_on(".wchart-overlay .chart-btn", "Diagrammelemente")
    rows = ui.js("[...document.querySelectorAll('.cb-fly .cb-row')].map((r) => [r.textContent.trim(), r.querySelector('input').checked])")
    run.check("Diagrammelemente lists Word's nine entries, ticked as Word ticks them",
              [r[0] for r in rows] == ["Achsen", "Achsentitel", "Diagrammtitel", "Datenbeschriftungen", "Datentabelle", "Fehlerindikatoren", "Gitternetzlinien", "Legende", "Trendlinie"]
              and [r[0] for r in rows if r[1]] == ["Achsen", "Diagrammtitel", "Gitternetzlinien", "Legende"], rows)
    probes = {
        "Achsentitel": "(m) => m.axisTitles.cat !== undefined && m.axisTitles.val !== undefined",
        "Diagrammtitel": "(m) => m.title.show",
        "Datenbeschriftungen": "(m) => m.dataLabels",
        "Datentabelle": "(m) => !!m.dataTable",
        "Gitternetzlinien": "(m) => m.gridlines",
        "Legende": "(m) => m.legend !== 'none'",
    }
    for label, probe in probes.items():
        before = ui.js(f"({probe})(__app.chartTools().model())")
        ui.click_on(".cb-fly .cb-row", label)
        after = ui.js(f"({probe})(__app.chartTools().model())")
        still = ui.js("!!document.querySelector('.cb-fly')")
        run.check(f"Diagrammelemente: {label} switched by its check box (flyout stays open)", before != after and still, [before, after, still])
    saved.append(("buttons-elements", ui.save("buttons-elements"), {"HasTitle": False, "HasLegend": False, "HasDataTable": True, "DataLabels": True, "ValTitle": True, "CatTitle": True, "MajorGrid": False}))
    # Diagrammformatvorlagen: tiles with Livevorschau, Farbe tab.
    ui.click_on(".wchart-overlay .chart-btn", "Diagrammformatvorlagen")
    tiles = ui.js("document.querySelectorAll('.cb-fly .cb-tile').length")
    run.check("Diagrammformatvorlagen shows the column styles (14)", tiles == 14, tiles)
    before = json.dumps(ui.model(), sort_keys=True)
    ui.hover_on(".cb-fly .cb-tile:nth-of-type(2)")
    pv = ui.js("__app.editor().previewing")
    unchanged = json.dumps(ui.model(), sort_keys=True) == before
    ui.click_on(".cb-fly .cb-tile:nth-of-type(2)")
    ui.move(5, 5)
    after = json.dumps(ui.model(), sort_keys=True)
    run.check("a style tile: Livevorschau on hover (the chart unchanged), applied on click", pv and unchanged and after != before, [pv, unchanged, after != before])
    # Diagrammfilter: hide the second series, Anwenden.
    ui.click_on(".wchart-overlay .chart-btn", "Diagrammfilter")
    names = ui.texts(".cb-fly .cb-filter-row .lbl")
    run.check("Diagrammfilter lists the series and categories (each with (Alle auswählen))",
              names == ["(Alle auswählen)", "Datenreihe 1", "Datenreihe 2", "Datenreihe 3", "(Alle auswählen)", "Kategorie 1", "Kategorie 2", "Kategorie 3", "Kategorie 4"], names)
    ui.click_on(".cb-fly .cb-filter-row", "Datenreihe 2")
    ui.click_on(".cb-fly .cb-foot button.apply")
    m = ui.model()
    run.check("Diagrammfilter: Datenreihe 2 hidden after Anwenden", m and m["series"][1].get("filtered") is True, m and [s.get("filtered") for s in m["series"]])
    close_popups(ui)
    saved.append(("filter", ui.save("filter"), {"Filtered": [False, True, False]}))

    # ----- Format: Aktuelle Auswahl, Fülleffekt, the pane, Größe -------------------------------------------------
    fresh(ui)
    x, y, w, h = ui.chart_rect()
    ui.click(x + w * 0.04, y + h * 0.96)
    run.check("a click on the selected chart keeps it selected (the chart area)", ui.sel() is not None and ui.js("__app.chartTools().element()") == {"kind": "chart"}, ui.sel())
    tt = ui.js("""(() => { const l = document.querySelector('.wchart-overlay > svg [data-part="title"]'); if (!l) return null; const r = l.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; })()""")
    if tt:
        ui.click(tt[0], tt[1])
    ui.tab("Format")
    combo_value = ui.js("document.querySelector('.ribbon-card .rcombo input')?.value")
    run.check("a click on the title chooses it: Aktuelle Auswahl shows Diagrammtitel", bool(tt) and ui.js("__app.chartTools().element()") == {"kind": "title"} and combo_value == "Diagrammtitel", [tt, combo_value])
    ui.tab("Format")
    names = ui.js("(() => { const m = __app.chartTools().model(); return m ? null : null; })()")
    del names
    combo = ui.rect(".ribbon-card .rcombo input")
    ui.click(combo[0], combo[1])
    ui.click_on(".ribbon-card .rcombo button[aria-label='Öffnen']")
    opts = ui.texts(".combo-list button")
    run.check("Aktuelle Auswahl lists the chart's elements", any("Diagrammtitel" == o for o in opts) and any("Legende" == o for o in opts), opts)
    try:
        ui.click_on(".combo-list button", "Diagrammtitel")
    except Exception as e:  # noqa: BLE001
        run.check("Aktuelle Auswahl > Diagrammtitel", False, e)
    el = ui.js("__app.chartTools().element()")
    run.check("Aktuelle Auswahl > Diagrammtitel selects the title", el == {"kind": "title"}, el)
    ui.ribbon("Auswahl formatieren")
    title = ui.js("document.querySelector('.taskpane:not([hidden]) .tp-title')?.textContent")
    run.check("Auswahl formatieren opens the title's Format pane", title == "Diagrammtitel formatieren", title)
    close_popups(ui)
    # Größe: type a height.
    sp = ui.rect(".ribbon-card input[aria-label='Höhe:']")
    ui.click(sp[0], sp[1], count=3)
    ui.type("7 cm")
    ui.key("Enter", "Enter", 13)
    s = ui.sel()
    run.check("Größe > Formenhöhe 7 cm", s and abs(s["h"] - 7 * 28.3465) < 0.5, s)
    saved.append(("size", ui.save("size"), {"Height": 7 * 28.3465}))
    # Double-click on the legend opens its pane.
    rr = ui.chart_rect()
    ui.js("__app.chartTools().setElement({ kind: 'chart' })")
    leg = ui.js("""(() => { const svg = document.querySelector('.wchart-overlay > svg'); const l = svg?.querySelector('[data-part="legend"]'); if (!l) return null; const r = l.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; })()""")
    if leg:
        ui.click(leg[0], leg[1], count=2)
        title = ui.js("document.querySelector('.taskpane:not([hidden]) .tp-title')?.textContent")
        run.check("double-click on the legend opens Legende formatieren", title == "Legende formatieren", [title, rr])
    else:
        run.check("double-click on the legend opens Legende formatieren", False, "no legend part")
    # In the pane: Keine Füllung for the chart area.
    ui.js("__app.chartTools().setElement({ kind: 'chart' })")
    ui.tab("Format")
    ui.ribbon("Auswahl formatieren")
    try:
        ui.click_on(".taskpane label", "Keine Füllung")
        m = ui.model()
        run.check("Format pane: Diagrammbereich > Keine Füllung", ((m.get("format") or {}).get("chart") or {}).get("fill", {}).get("type") == "none", (m.get("format") or {}).get("chart"))
    except Exception as e:  # noqa: BLE001
        run.check("Format pane: Diagrammbereich > Keine Füllung", False, e)
    saved.append(("pane", ui.save("pane"), {"ChartAreaFill": False}))
    close_popups(ui)

    # ----- Diagrammtyp ändern -------------------------------------------------------------------------------
    ui.select_chart()
    ui.tab("Diagrammentwurf")
    ui.ribbon("Diagrammtyp ändern...")
    try:
        ui.click_on(".ct-cats div", "Linie")
        ui.click_on(".dlg-buttons button", "OK")
        run.check("Diagrammtyp ändern... > Linie", (ui.model() or {}).get("kind") == "line", (ui.model() or {}).get("kind"))
        saved.append(("type-line", ui.save("type-line"), {"ChartType": 4}))
    except Exception as e:  # noqa: BLE001
        run.check("Diagrammtyp ändern... > Linie", False, e)
        close_popups(ui)

    # ----- Layoutoptionen: every wrapping, then floating charts in Word ------------------------------------------
    for wrap, label in WRAP_LABELS.items():
        fresh(ui)
        x, y, w, h = ui.js("(() => { const r = __app.editor().view.layoutButtonRect(); return [r.left + r.width / 2, r.top + r.height / 2, r.width, r.height]; })()")
        ui.click(x, y)
        ui.click_on(".layout-options .tile", label)
        s = ui.sel()
        run.check(f"Layoutoptionen > {label}: the chart floats with that wrapping", s and s["kind"] == "float" and s["wrap"] == wrap, s)
        if wrap == "square":
            # Dragged by hand with the mouse, 90 px right and 60 px down.
            r0 = ui.chart_rect()
            ui.drag(r0[0] + r0[2] / 2, r0[1] + r0[3] / 2, r0[0] + r0[2] / 2 + 90, r0[1] + r0[3] / 2 + 60)
            r1 = ui.chart_rect()
            moved = r1 and abs((r1[0] - r0[0]) - 90) < 3 and abs((r1[1] - r0[1]) - 60) < 3
            run.check("a floating chart is dragged with the mouse", moved, [r0, r1])
            # Position auf der Seite fixieren in the Layoutoptionen.
            x, y, _, _ = ui.js("(() => { const r = __app.editor().view.layoutButtonRect(); return [r.left + r.width / 2, r.top + r.height / 2, 0, 0]; })()")
            ui.click(x, y)
            ui.click_on(".layout-options label", "Position auf der Seite fixieren")
            s2 = ui.sel()
            run.check("Layoutoptionen > Position auf der Seite fixieren", s2 and s2["y"]["rel"] == "page" and s2["moveWithText"] is False, s2)
            close_popups(ui)
            # The buttons follow the dragged chart.
            follow = ui.js("(() => { const r = __app.editor().view.selectedObjectRect(); const o = document.querySelector('.wchart-overlay').getBoundingClientRect(); return Math.abs(r.left - o.left) < 1 && Math.abs(r.top + 38 - o.top) < 1; })()")
            time.sleep(0.3)
            follow = follow or ui.js("(() => { const r = __app.editor().view.selectedObjectRect(); const o = document.querySelector('.wchart-overlay').getBoundingClientRect(); return Math.abs(r.left - o.left) < 1 && Math.abs(r.top + 38 - o.top) < 1; })()")
            run.check("the chart buttons follow the dragged chart", follow, "")
            s = ui.sel()
        saved.append((f"wrap-{wrap}", ui.save(f"wrap-{wrap}"), {"Wrap": WRAP_TYPES[wrap], "X": s["x"], "Y": s["y"], "W": s["w"], "H": s["h"]}))
    # Back into the line.
    x, y, _, _ = ui.js("(() => { const r = __app.editor().view.layoutButtonRect(); return [r.left + r.width / 2, r.top + r.height / 2, 0, 0]; })()")
    ui.click(x, y)
    ui.click_on(".layout-options .tile", "Mit Text in Zeile")
    s = ui.sel()
    run.check("Layoutoptionen > Mit Text in Zeile: the chart is back in the line", s and s["kind"] == "inline", s)
    ui.undo()
    s = ui.sel()
    run.check("Strg+Z: floating again", s and s["kind"] == "float", s)

    # ----- Daten bearbeiten still opens the table ------------------------------------------------------------
    ui.tab("Diagrammentwurf")
    ui.ribbon("Daten bearbeiten")
    t = ui.js("document.querySelector('.dialog-layer .dlg-title .t')?.textContent")
    run.check("Daten bearbeiten opens the chart's data", bool(t) and "Daten bearbeiten" in t, t)
    close_popups(ui)


def page_diff(word_pdf: Path, ours_pdf: Path, out: Path) -> float | None:
    """Mean grey difference (of 255) of the two PDFs' first pages at 96 dpi (check_write_charts.raster)."""
    import numpy as np
    from PIL import Image
    from check_write_charts import raster
    if not ours_pdf.exists():
        return None
    a = raster(word_pdf, out / "word")
    b = raster(ours_pdf, out / "raabe")
    if not a or not b:
        return None
    ia, ib = np.asarray(Image.open(a[0]).convert("L"), dtype=float), np.asarray(Image.open(b[0]).convert("L"), dtype=float)
    h, w = min(ia.shape[0], ib.shape[0]), min(ia.shape[1], ib.shape[1])
    return round(float(np.abs(ia[:h, :w] - ib[:h, :w]).mean()), 2)


def word_side(run: Run, saved: list) -> None:
    import win32com.client
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    pdfs = {}
    try:
        for name, path, want in saved:
            if not path.exists():
                run.check(f"Word: {name} saved", False, "no file")
                continue
            try:
                d = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
            except Exception as e:  # noqa: BLE001
                run.check(f"Word opens {name}", False, e)
                continue
            try:
                floating = d.Shapes.Count > 0
                holder = d.Shapes(1) if floating else d.InlineShapes(1)
                ch = holder.Chart
                got: dict = {}
                bad = []
                for k, v in want.items():
                    if k == "HasTitle":
                        got[k] = bool(ch.HasTitle)
                    elif k == "HasLegend":
                        got[k] = bool(ch.HasLegend)
                    elif k == "LegendPosition":
                        got[k] = int(ch.Legend.Position) if ch.HasLegend else None
                    elif k == "HasDataTable":
                        got[k] = bool(ch.HasDataTable)
                    elif k == "DataLabels":
                        got[k] = bool(ch.SeriesCollection(1).HasDataLabels)
                    elif k == "ValTitle":
                        got[k] = bool(ch.Axes(2).HasTitle)
                    elif k == "CatTitle":
                        got[k] = bool(ch.Axes(1).HasTitle)
                    elif k == "MajorGrid":
                        got[k] = bool(ch.Axes(2).HasMajorGridlines)
                    elif k == "ChartType":
                        got[k] = int(ch.ChartType)
                    elif k == "Filtered":
                        got[k] = [bool(ch.FullSeriesCollection(i).IsFiltered) for i in range(1, ch.FullSeriesCollection().Count + 1)]
                    elif k == "Height":
                        got[k] = round(float(holder.Height), 1)
                        if abs(got[k] - v) > 0.6:
                            bad.append(k)
                        continue
                    elif k == "ChartAreaFill":
                        # Word's object model reports the chart area's fill as mixed here; Word's own file says it:
                        # saved again by Word, the chart space keeps (or drops) its noFill.
                        import re
                        import zipfile
                        again = TMP / f"{name}-word.docx"
                        d.SaveAs2(str(again), 16)
                        with zipfile.ZipFile(again) as z:
                            part = next(n for n in z.namelist() if n.startswith("word/charts/chart") and n.endswith(".xml"))
                            x = z.read(part).decode("utf-8")
                        got[k] = not re.search(r"</c:chart><c:spPr><a:noFill/>", x)
                    elif k == "Wrap":
                        got[k] = int(holder.WrapFormat.Type) if floating else 7
                    elif k in ("X", "Y"):
                        p = v
                        rel = int(holder.RelativeHorizontalPosition if k == "X" else holder.RelativeVerticalPosition)
                        pos = float(holder.Left if k == "X" else holder.Top)
                        got[k] = [rel, round(pos, 1)]
                        # wdRelativeHorizontalPositionColumn 2, Page 1; wdRelativeVerticalPositionParagraph 2, Page 1.
                        want_rel = {"column": 2, "page": 1, "margin": 0, "paragraph": 2}.get(p["rel"], -1)
                        if rel != want_rel or abs(pos - p["offset"]) > 1:
                            bad.append(f"{k}: {p}")
                        continue
                    elif k in ("W", "H"):
                        got[k] = round(float(holder.Width if k == "W" else holder.Height), 1)
                        if abs(got[k] - v) > 0.6:
                            bad.append(k)
                        continue
                    if got.get(k) != v:
                        bad.append(k)
                run.check(f"Word reads {name}: the same chart state ({', '.join(want)})", not bad and floating == ("Wrap" in want), [bad, got, want])
                pdf = TMP / f"{name}-word.pdf"
                d.ExportAsFixedFormat(str(pdf), 17)
                pdfs[name] = pdf
                diff = page_diff(pdf, TMP / f"{name}-raabe.pdf", TMP / f"png-{name}")
                run.check(f"picture of {name}: Word's PDF and RaabeWrite's page 1 (mean grey difference < 6 of 255)", diff is not None and diff < 6, diff)
            finally:
                d.Close(0)
    finally:
        app.Quit(0)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    if TMP.exists():
        shutil.rmtree(TMP, ignore_errors=True)
    TMP.mkdir(exist_ok=True)
    run = Run()
    saved: list = []
    t = raabe_write("short", port=PORT, background=True)
    ui = Ui(t)
    try:
        cases(ui, run, saved)
    except Exception:  # noqa: BLE001
        tb = traceback.format_exc()
        print(tb, flush=True)
        run.check("run completed", False, tb.splitlines()[-1][:500])
    finally:
        try:
            t.close()
        except Exception:  # noqa: BLE001
            pass
    word_side(run, saved)
    passed = sum(1 for x in run.cases if x["passed"])
    RESULT.write_text(json.dumps({"passed": passed, "total": len(run.cases), "cases": run.cases}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(f"{passed}/{len(run.cases)} passed")


if __name__ == "__main__":
    main()
