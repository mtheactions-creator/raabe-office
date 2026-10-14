// Start > Formatvorlagen (dialog launcher, Strg+Alt+Umschalt+S): Word's pane Formatvorlagen
// (recorded 2026-10-14) with Alle löschen, the styles and their type marks, Vorschau anzeigen,
// Verknüpfte Formatvorlagen deaktivieren, Neue Formatvorlage, Formatvorlageninspektor and
// Optionen...; each entry's menu (aktualisieren, um der Auswahl zu entsprechen, Ändern...,
// löschen, Formatvorlagenkatalog) and the dialog Formatvorlage ändern / Neue Formatvorlage von
// Formatierung erstellen with its Format menu (Zeichen..., Absatz..., Sprache...).

import { h, icon, messageBox, showDialog, showFontColors, showMenu, type MenuEntry } from "@raabe/shell";
import type { CharProps, CharState, Editor, ParaProps, ResolvedChar, ResolvedPara, StyleDef } from "@raabe/text-engine";
import type { AppApi } from "./api";
import { openModal } from "./modal";

type Show = "recommended" | "inUse" | "inDoc" | "all";
type Sort = "alpha" | "recommended" | "font" | "basedOn" | "type";
type Entry = ReturnType<Editor["styleCatalog"]>[number];

const PARA_MARK = "¶";
const THEME_COLOR_NAMES: Record<string, string> = {
  dark1: "Text 1", text1: "Text 1", light1: "Hintergrund 1", background1: "Hintergrund 1", dark2: "Text 2", text2: "Text 2", light2: "Hintergrund 2", background2: "Hintergrund 2",
  accent1: "Akzent 1", accent2: "Akzent 2", accent3: "Akzent 3", accent4: "Akzent 4", accent5: "Akzent 5", accent6: "Akzent 6", hyperlink: "Link", followedHyperlink: "Besuchter Link",
};
const LANGS: [string, string][] = [["de-DE", "Deutsch (Deutschland)"], ["de-AT", "Deutsch (Österreich)"], ["de-CH", "Deutsch (Schweiz)"], ["en-US", "Englisch (USA)"],
  ["en-GB", "Englisch (Vereinigtes Königreich)"], ["fr-FR", "Französisch (Frankreich)"], ["es-ES", "Spanisch (Spanien)"], ["it-IT", "Italienisch (Italien)"], ["la", "Latein"]];

const num = (v: number, digits = 2) => v.toLocaleString("de-DE", { maximumFractionDigits: digits });
const cm = (pt: number) => `${num(pt / 28.3465)} cm`;

/** Word's description of formatting (Formatvorlage ändern, Formatvorlageninspektor). */
export function describeFormat(p: ParaProps, r: CharProps): string[] {
  const out: string[] = [];
  const font: string[] = [];
  if (r.font) font.push(r.font);
  else if (r.fontTheme) font.push(r.fontTheme.startsWith("major") ? "+Überschriften" : "+Textkörper");
  if (r.size) font.push(`${num(r.size, 1)} Pt.`);
  if (r.bold) font.push("Fett");
  if (r.bold === false) font.push("Nicht Fett");
  if (r.italic) font.push("Kursiv");
  if (r.italic === false) font.push("Nicht Kursiv");
  if (r.underline && r.underline !== "none") font.push(r.underline === "double" ? "Doppelt unterstrichen" : "Unterstrichen");
  if (r.strike) font.push("Durchgestrichen");
  if (r.dstrike) font.push("Doppelt durchgestrichen");
  if (r.caps) font.push("Großbuchstaben");
  if (r.smallCaps) font.push("Kapitälchen");
  if (r.hidden) font.push("Ausgeblendet");
  if (r.vertAlign === "superscript") font.push("Hochgestellt");
  if (r.vertAlign === "subscript") font.push("Tiefgestellt");
  if (r.colorTheme && THEME_COLOR_NAMES[r.colorTheme]) font.push(`Schriftfarbe: ${THEME_COLOR_NAMES[r.colorTheme]}`);
  else if (r.color && r.color !== "auto") {
    const n = parseInt(r.color, 16);
    font.push(`Schriftfarbe: Benutzerdefinierte Farbe(RGB(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255}))`);
  }
  if (r.spacing) font.push(r.spacing > 0 ? `Erweitert durch ${num(r.spacing)} Pt.` : `Schmal durch ${num(-r.spacing)} Pt.`);
  if (font.length) out.push(`Schriftart: ${font.join(", ")}`);
  if (p.align && p.align !== "left") out.push(p.align === "center" ? "Zentriert" : p.align === "right" ? "Rechts" : "Blocksatz");
  const ind: string[] = [];
  if (p.indLeft) ind.push(`Links: ${cm(p.indLeft)}`);
  if (p.indRight) ind.push(`Rechts: ${cm(p.indRight)}`);
  if (p.indFirst) ind.push(p.indFirst > 0 ? `Erste Zeile: ${cm(p.indFirst)}` : `Hängend: ${cm(-p.indFirst)}`);
  if (ind.length) out.push(`Einzug: ${ind.join(", ")}`);
  if (p.line !== undefined) {
    const rule = p.lineRule ?? "auto";
    out.push(`Zeilenabstand: ${rule === "exact" ? `Genau ${num(p.line, 1)} Pt.` : rule === "atLeast" ? `Mindestens ${num(p.line, 1)} Pt.` :
      Math.abs(p.line - 1) < 0.01 ? "einfach" : Math.abs(p.line - 1.5) < 0.01 ? "1,5 Zeilen" : Math.abs(p.line - 2) < 0.01 ? "Doppelt" : `Mehrfach ${num(p.line)} ze`}`);
  }
  const sp: string[] = [];
  if (p.before !== undefined) sp.push(`Vor: ${num(p.before, 1)} Pt.`);
  if (p.after !== undefined) sp.push(`Nach: ${num(p.after, 1)} Pt.`);
  if (sp.length) out.push(`Abstand ${sp.join(" ")}`);
  if (p.contextual) out.push("Keinen Abstand zwischen Absätzen gleicher Formatvorlage einfügen");
  if (p.keepNext) out.push("Nicht vom nächsten Absatz trennen");
  if (p.keepLines) out.push("Diesen Absatz zusammenhalten");
  if (p.pageBreakBefore) out.push("Seitenumbruch oberhalb");
  if (p.outline !== undefined && p.outline < 9) out.push(`Ebene ${p.outline + 1}`);
  return out;
}

export function stylesPaneFeature(api: AppApi, docArea: HTMLElement) {
  const ed = () => api.editor();
  const c = api.commands;
  let open = false;
  const opts = { show: "recommended" as Show, sort: "recommended" as Sort, preview: false, noLinked: false, nextHeading: true };
  const pane = h("div", { class: "stylepane", hidden: true, role: "complementary", "aria-label": "Formatvorlagen" });
  docArea.append(pane);
  let inspector: HTMLElement | null = null;

  const setOpen = (on: boolean) => {
    open = on;
    pane.hidden = !on;
    lastKey = "";
    render();
    api.refresh();
    if (!on) api.focus();
  };

  const headingLevel = (e: Editor, id: string) => {
    const m = /^heading (\d)$/i.exec(e.styleDef(id)?.name ?? "");
    return m ? Number(m[1]) : 0;
  };

  /** The entries the options choose, sorted as chosen. */
  const entries = (e: Editor): Entry[] => {
    const all = e.styleCatalog().filter((x) => x.type === "paragraph" || x.type === "character" || (opts.show === "all" && x.type === "table"));
    const usedHeadings = new Set(all.filter((x) => x.inUse).map((x) => headingLevel(e, x.id)).filter(Boolean));
    const shown = all.filter((x) => {
      if (opts.show === "all") return true;
      if (opts.show === "inUse") return x.inUse;
      if (opts.show === "inDoc") return !!e.doc.styles.get(x.id);
      // Empfohlen: Word's recommended styles; Überschrift n once n − 1 is used.
      const lvl = headingLevel(e, x.id);
      if (lvl > 1) return x.inUse || (opts.nextHeading && usedHeadings.has(lvl - 1));
      return x.recommended && x.priority < 99 || x.inUse;
    });
    const font = (x: Entry) => e.styleLook(x.id).char.font;
    const typeRank = (x: Entry) => (x.type === "paragraph" ? (x.linked && !opts.noLinked ? 1 : 0) : x.type === "character" ? 2 : 3);
    const byName = (a: Entry, b: Entry) => a.name.localeCompare(b.name, "de");
    return shown.sort(opts.sort === "alpha" ? byName
      : opts.sort === "font" ? (a, b) => font(a).localeCompare(font(b)) || byName(a, b)
      : opts.sort === "basedOn" ? (a, b) => (e.styleDef(a.id)?.basedOn ?? "").localeCompare(e.styleDef(b.id)?.basedOn ?? "") || byName(a, b)
      : opts.sort === "type" ? (a, b) => typeRank(a) - typeRank(b) || byName(a, b)
      : (a, b) => a.priority - b.priority || byName(a, b));
  };

  const mark = (x: { type: string; linked: boolean }) => {
    if (x.type === "character") return h("span", { class: "smark" }, "a");
    if (x.type === "table") return h("span", { class: "smark" }, icon("table"));
    if (x.linked && !opts.noLinked) return h("span", { class: "smark" }, PARA_MARK, h("u", null, "a"));
    return h("span", { class: "smark" }, PARA_MARK);
  };

  const apply = (e: Editor, x: Entry) => {
    if (x.type === "character") e.applyCharStyle(x.id);
    else if (x.type === "paragraph") e.applyParaStyle(x.id);
    api.refresh();
    api.focus();
  };

  const entryMenu = (e: Editor, x: Entry): MenuEntry[] => {
    const def = e.styleDef(x.id);
    const inDoc = !!e.doc.styles.get(x.id);
    return [
      { label: `${x.name} aktualisieren, um der Auswahl zu entsprechen`, run: () => { e.matchStyleToSelection(x.id); api.refresh(); } },
      { label: "Ändern...", icon: "style-modify", run: () => styleDialog(e, x.id) },
      "-",
      { label: "Alles markieren: (Keine Daten)", run: () => {}, disabled: true },
      { label: "Alles entfernen: (Keine Daten)", run: () => {}, disabled: true },
      { label: `${x.name} löschen...`, disabled: x.builtin || !inDoc || !!def?.isDefault, run: () => void confirmDelete(e, x) },
      "-",
      x.qFormat
        ? { label: "Aus dem Formatvorlagenkatalog entfernen", run: () => { if (def) e.defineStyle({ ...def, qFormat: false }); api.refresh(); } }
        : { label: "Zum Formatvorlagenkatalog hinzufügen", run: () => { if (def) e.defineStyle({ ...def, qFormat: true, semiHidden: false }); api.refresh(); } },
    ];
  };

  const confirmDelete = async (e: Editor, x: Entry) => {
    const ok = await showDialog({ title: "RaabeWrite", body: `Möchten Sie die Formatvorlage „${x.name}“ aus dem Dokument löschen?`,
      buttons: [{ label: "Ja", value: true, primary: true }, { label: "Nein", value: false }], cancel: false });
    if (ok) { e.removeStyle(x.id); lastKey = ""; render(); api.refresh(); }
  };

  let lastKey = "";
  const render = () => {
    if (!open) return;
    const e = ed();
    if (!e) { pane.replaceChildren(); return; }
    const list = entries(e);
    const cur = e.charState().charStyle && e.charState().charStyle !== "Absatz-Standardschriftart" ? e.charState().charStyle! : e.currentStyle();
    const key = JSON.stringify([list.map((x) => [x.id, x.name, x.inUse, x.qFormat]), cur, opts, opts.preview ? e.doc.styles : 0]);
    if (key === lastKey) return;
    lastKey = key;
    const head = h("div", { class: "sty-head" }, h("span", null, "Formatvorlagen"),
      h("div", null,
        h("button", { "aria-label": "Optionen für den Aufgabenbereich", "data-tip": "Optionen für den Aufgabenbereich",
          onclick: (ev: MouseEvent) => showMenu(c, [{ label: "Schließen", run: () => setOpen(false) }], (ev.currentTarget as HTMLElement).getBoundingClientRect()) }, icon("chevron-down")),
        h("button", { "aria-label": "Bereich schließen", "data-tip": "Bereich schließen", onclick: () => setOpen(false) }, icon("close"))));
    const rows = h("div", { class: "sty-list" });
    const clear = h("div", { class: "sty-row", role: "button", tabIndex: 0, onclick: () => { e.clearAllFormatting(); api.refresh(); api.focus(); } }, h("span", { class: "sty-name" }, "Alle löschen"));
    rows.append(clear);
    for (const x of list) {
      const name = h("span", { class: "sty-name" }, x.name);
      if (opts.preview) {
        const look = e.styleLook(x.id).char;
        name.style.cssText = `font-family:"${look.font}";font-size:${Math.min(16, look.size) * 4 / 3}px;font-weight:${look.bold ? 700 : 400};font-style:${look.italic ? "italic" : "normal"};` +
          `color:${look.color === "auto" ? "#000" : `#${look.color}`};${look.smallCaps ? "font-variant:small-caps;" : ""}${look.caps ? "text-transform:uppercase;" : ""}`;
      }
      const arrow = h("button", { class: "sty-arrow", "aria-label": `${x.name}: Optionen`, tabIndex: -1 }, icon("chevron-down"));
      const row = h("div", { class: `sty-row${x.id === cur ? " active" : ""}`, role: "button", tabIndex: 0, "data-tip": tipFor(e, x) }, name, mark(x), arrow);
      row.onclick = (ev) => { if ((ev.target as Element).closest(".sty-arrow")) return; apply(e, x); };
      const menu = (rect: DOMRect) => showMenu(c, entryMenu(e, x), rect);
      arrow.onclick = (ev) => { ev.stopPropagation(); menu(row.getBoundingClientRect()); };
      row.oncontextmenu = (ev) => { ev.preventDefault(); menu(new DOMRect(ev.clientX, ev.clientY, 0, 0)); };
      rows.append(row);
    }
    const check = (label: string, on: boolean, set: (v: boolean) => void) => {
      const box = h("input", { type: "checkbox", checked: on }) as HTMLInputElement;
      box.onchange = () => { set(box.checked); lastKey = ""; render(); };
      return h("label", { class: "check" }, box, label);
    };
    const foot = h("div", { class: "sty-foot" },
      check("Vorschau anzeigen", opts.preview, (v) => { opts.preview = v; }),
      check("Verknüpfte Formatvorlagen deaktivieren", opts.noLinked, (v) => { opts.noLinked = v; }),
      h("div", { class: "sty-buttons" },
        h("button", { class: "icon", "aria-label": "Neue Formatvorlage", "data-tip": "Neue Formatvorlage", onclick: () => styleDialog(e, null) }, icon("style-new")),
        h("button", { class: "icon", "aria-label": "Formatvorlageninspektor", "data-tip": "Formatvorlageninspektor", onclick: () => toggleInspector() }, icon("style-inspect")),
        h("button", { class: "icon", "aria-label": "Formatvorlagen verwalten", disabled: true, "data-tip": "Formatvorlagen verwalten\nÄndern Sie Formatvorlagen mit „Ändern...“ im Menü eines Eintrags." }, icon("style-manage")),
        h("button", { class: "sty-options", onclick: () => optionsDialog() }, "Optionen...")));
    pane.replaceChildren(head, rows, foot);
  };

  const tipFor = (e: Editor, x: Entry) => {
    const def = e.styleDef(x.id);
    if (!def) return x.name;
    const lines = describeFormat(def.pPr, def.rPr);
    if (def.basedOn) lines.push(`Basiert auf: ${e.doc.styles.displayName(def.basedOn)}`);
    return `${x.name}\n${lines.join(", ") || "Keine eigene Formatierung"}`;
  };

  // ----- Optionen für Formatvorlagenbereich ---------------------------------------------------
  const optionsDialog = () => {
    const show = h("select", {}, ...([["recommended", "Empfohlen"], ["inUse", "In Verwendung"], ["inDoc", "Im aktuellen Dokument"], ["all", "Alle Formatvorlagen"]] as const)
      .map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
    show.value = opts.show;
    const sort = h("select", {}, ...([["alpha", "Alphabetisch"], ["recommended", "Wie empfohlen"], ["font", "Schriftart"], ["basedOn", "Basierend auf"], ["type", "Nach Typ"]] as const)
      .map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
    sort.value = opts.sort;
    const box = (label: string, on: boolean, disabled = false) => {
      const b = h("input", { type: "checkbox", checked: on, disabled }) as HTMLInputElement;
      return { b, el: h("label", { class: `check${disabled ? " disabled" : ""}` }, b, label) };
    };
    const fPara = box("Formatierung auf Absatzebene", false, true), fFont = box("Schriftartformatierung", false, true), fList = box("Formatierung für Aufzählungszeichen und Nummerierungen", false, true);
    const next = box("Bei Verwendung der vorherigen Ebene nächste Überschrift anzeigen", opts.nextHeading);
    const alias = box("Integrierten Namen ausblenden, wenn ein alternativer Name vorhanden ist.", false, true);
    const onlyDoc = h("input", { type: "radio", name: "styopt", checked: true }) as HTMLInputElement;
    const newDocs = h("input", { type: "radio", name: "styopt", disabled: true }) as HTMLInputElement;
    openModal("Optionen für Formatvorlagenbereich", h("div", { class: "styopts" },
      h("label", { class: "field" }, "Anzuzeigende Formatvorlagen auswählen:", show),
      h("label", { class: "field" }, "Sortierung für die Liste auswählen:", sort),
      h("div", { class: "sub" }, "Formatierung zum Anzeigen als Formatvorlagen auswählen:", fPara.el, fFont.el, fList.el),
      h("div", { class: "sub" }, "Darstellung von integrierten Formatvorlagennamen auswählen", next.el, alias.el),
      h("div", { class: "row" }, h("label", { class: "check" }, onlyDoc, "Nur in diesem Dokument"), h("label", { class: "check disabled" }, newDocs, "Neue auf dieser Vorlage basierende Dokumente"))), [
      { label: "OK", primary: true, run: () => { opts.show = show.value as Show; opts.sort = sort.value as Sort; opts.nextHeading = next.b.checked; lastKey = ""; render(); return true; } },
      { label: "Abbrechen", run: () => true },
    ], { width: 520 });
  };

  // ----- Formatvorlageninspektor ----------------------------------------------------------------
  const toggleInspector = () => {
    if (inspector) { inspector.remove(); inspector = null; return; }
    inspector = h("div", { class: "dialog floating styinspect", role: "dialog", "aria-label": "Formatvorlageninspektor" });
    document.body.append(inspector);
    renderInspector();
  };
  const renderInspector = () => {
    const e = ed();
    if (!inspector || !e) return;
    const d = e.directFormatting();
    const sheet = e.doc.styles;
    const plus = (lines: string[]) => h("div", { class: "insp-plus" }, lines.length ? `Plus: ${lines.join(", ")}` : "Plus: <keine>");
    const close = () => { inspector?.remove(); inspector = null; };
    inspector.replaceChildren(
      h("div", { class: "dlg-title" }, h("span", { class: "t" }, "Formatvorlageninspektor"), h("button", { "aria-label": "Schließen", onclick: close }, icon("close"))),
      h("div", { class: "dlg-body" },
        h("div", { class: "insp-h" }, "Absatzformatierung"),
        h("div", { class: "insp-row" }, h("span", { class: "insp-box" }, sheet.displayName(d.paraStyle)),
          h("button", { class: "icon", "data-tip": "Absatzformatvorlage auf Standard zurücksetzen", "aria-label": "Absatzformatvorlage auf Standard zurücksetzen", onclick: () => { e.applyParaStyle(sheet.defaultParagraphStyle); renderInspector(); } }, icon("eraser"))),
        h("div", { class: "insp-row" }, plus(describeFormat(d.para, {})),
          h("button", { class: "icon", "data-tip": "Absatzformatierung löschen", "aria-label": "Absatzformatierung löschen", onclick: () => { e.clearDirect("para"); renderInspector(); } }, icon("eraser"))),
        h("div", { class: "insp-h" }, "Textebenenformatierung"),
        h("div", { class: "insp-row" }, h("span", { class: "insp-box" }, d.charStyle ? sheet.displayName(d.charStyle) : "Absatz-Standardschriftart"),
          h("button", { class: "icon", "data-tip": "Zeichenformatvorlage löschen", "aria-label": "Zeichenformatvorlage löschen", onclick: () => { e.applyCharStyle("Absatz-Standardschriftart"); renderInspector(); } }, icon("eraser"))),
        h("div", { class: "insp-row" }, plus(describeFormat({}, d.char)),
          h("button", { class: "icon", "data-tip": "Zeichenformatierung löschen", "aria-label": "Zeichenformatierung löschen", onclick: () => { e.clearDirect("char"); renderInspector(); } }, icon("eraser"))),
        h("div", { class: "row" }, h("button", { onclick: () => { e.clearAllFormatting(); renderInspector(); } }, "Alle löschen"))));
  };

  // ----- Formatvorlage ändern / Neue Formatvorlage von Formatierung erstellen -------------------
  const styleDialog = (e: Editor, id: string | null) => {
    const sheet = e.doc.styles;
    const cs = e.charState();
    const base = id ? e.styleDef(id) : undefined;
    if (id && !base) return;
    const curPara = e.currentStyle();
    // A new style starts from the selection's formatting, based on its paragraph style.
    const draft: { def: StyleDef } = {
      def: base ? { ...base, pPr: { ...base.pPr }, rPr: { ...base.rPr } } : {
        id: "", name: newName(e), type: "paragraph", basedOn: curPara, next: undefined, qFormat: true, uiPriority: undefined,
        pPr: {}, rPr: newRPr(e, cs),
      },
    };
    const isNew = !base;
    const nameIn = h("input", { type: "text", value: base ? sheet.displayName(base.id) : draft.def.name, style: "width:100%" }) as HTMLInputElement;
    const typeSel = h("select", { disabled: !isNew }, h("option", { value: "paragraph" }, "Absatz"), h("option", { value: "character" }, "Zeichen"),
      h("option", { value: "linked" }, "Verknüpft (Absatz und Zeichen)"), ...(base?.type === "table" ? [h("option", { value: "table" }, "Tabelle")] : [])) as HTMLSelectElement;
    typeSel.value = base ? (base.type === "paragraph" && base.link ? "linked" : base.type) : "paragraph";
    const styleOptions = (kind: "paragraph" | "character", none: string) => {
      const list = e.styleCatalog().filter((x) => x.type === kind && x.id !== id).sort((a, b) => a.name.localeCompare(b.name, "de"));
      return [h("option", { value: "" }, none), ...list.map((x) => h("option", { value: x.id }, `${kind === "paragraph" ? PARA_MARK : "a"} ${x.name}`))];
    };
    const basedSel = h("select", {}) as HTMLSelectElement;
    const nextSel = h("select", {}) as HTMLSelectElement;
    const fillSelects = () => {
      const kind = typeSel.value === "character" ? "character" : "paragraph";
      basedSel.replaceChildren(...styleOptions(kind, "(Keine Formatvorlage)"));
      basedSel.value = draft.def.basedOn ?? "";
      nextSel.replaceChildren(...styleOptions("paragraph", ""));
      nextSel.value = draft.def.next ?? (draft.def.type === "paragraph" ? (id ?? "") : "");
      nextSel.disabled = kind === "character";
      if (!nextSel.value && kind === "paragraph" && id) nextSel.value = id;
    };
    fillSelects();
    // Formatierung: font, size, F K U, colour; alignment, line spacing, spacing, indent.
    const fonts = api.fontList();
    const fontIn = h("input", { type: "text", list: "sty-fonts", style: "width:180px" }) as HTMLInputElement;
    const fontList = h("datalist", { id: "sty-fonts" }, ...fonts.map((f) => h("option", { value: f })));
    const sizeSel = h("input", { type: "text", list: "sty-sizes", style: "width:56px" }) as HTMLInputElement;
    const sizeList = h("datalist", { id: "sty-sizes" }, ...[8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72].map((s) => h("option", { value: num(s, 1) })));
    const toggle = (label: string, tip: string, key: "bold" | "italic") => {
      const b = h("button", { class: "fmt-toggle", "data-tip": tip, "aria-label": tip, onclick: () => { draft.def = { ...draft.def, rPr: { ...draft.def.rPr, [key]: !look().char[key] } }; refresh(); } }, label);
      return b;
    };
    const boldB = toggle("F", "Fett", "bold"), italB = toggle("K", "Kursiv", "italic");
    const underB = h("button", { class: "fmt-toggle", "data-tip": "Unterstrichen", "aria-label": "Unterstrichen", style: "text-decoration:underline",
      onclick: () => { const u = look().char.underline; draft.def = { ...draft.def, rPr: { ...draft.def.rPr, underline: u && u !== "none" ? "none" : "single" } }; refresh(); } }, "U");
    const colorB = h("button", { class: "fmt-color", "aria-label": "Schriftfarbe", "data-tip": "Schriftfarbe" });
    colorB.onclick = () => showFontColors(colorB.getBoundingClientRect(), (hex) => { draft.def = { ...draft.def, rPr: { ...draft.def.rPr, color: hex, colorTheme: undefined } }; refresh(); });
    const pbtn = (ic: string, tip: string, change: (p: ResolvedPara) => Partial<ParaProps>) =>
      h("button", { class: "icon fmt-p", "data-tip": tip, "aria-label": tip, onclick: () => { draft.def = { ...draft.def, pPr: { ...draft.def.pPr, ...change(look().para) } }; refresh(); } }, icon(ic));
    const alignBtns = ([["left", "align-left", "Linksbündig"], ["center", "align-center", "Zentriert"], ["right", "align-right", "Rechtsbündig"], ["both", "align-justify", "Blocksatz"]] as const)
      .map(([v, ic, tip]) => ({ v, el: pbtn(ic, tip, () => ({ align: v })) }));
    const lineBtns = ([[1, "Einfacher Zeilenabstand"], [1.5, "1,5-facher Zeilenabstand"], [2, "Doppelter Zeilenabstand"]] as const)
      .map(([v, tip]) => ({ v, el: pbtn(v === 1 ? "line-single" : v === 1.5 ? "line-onehalf" : "line-double", tip, () => ({ line: v, lineRule: "auto" })) }));
    const spaceUp = pbtn("space-increase", "Absatzabstand vergrößern", (p) => ({ before: p.before + 6, after: p.after + 6 }));
    const spaceDown = pbtn("space-decrease", "Absatzabstand verkleinern", (p) => ({ before: Math.max(0, p.before - 6), after: Math.max(0, p.after - 6) }));
    const indDown = pbtn("indent-decrease", "Einzug verkleinern", (p) => ({ indLeft: Math.max(0, p.indLeft - 35.4) }));
    const indUp = pbtn("indent-increase", "Einzug vergrößern", (p) => ({ indLeft: p.indLeft + 35.4 }));
    const preview = h("div", { class: "sty-preview" });
    const desc = h("div", { class: "sty-desc" });
    const gallery = h("input", { type: "checkbox", checked: base ? !!base.qFormat : true }) as HTMLInputElement;
    const autoUpd = h("input", { type: "checkbox", disabled: true }) as HTMLInputElement;
    const onlyDoc = h("input", { type: "radio", name: "styscope", checked: true }) as HTMLInputElement;
    const newDocs = h("input", { type: "radio", name: "styscope", disabled: true }) as HTMLInputElement;

    const look = (): { para: ResolvedPara; char: ResolvedChar } => e.styleLookOf({ ...draft.def, id: draft.def.id || "__neu__" });
    const sample = e.doc.paras[e.sel.head.p]?.text.replace(/[\x00-\x1f\ufffc]/g, " ").trim().slice(0, 140) || "Beispieltext Beispieltext Beispieltext Beispieltext Beispieltext";
    const refresh = () => {
      const { para, char } = look();
      fontIn.value = char.font;
      sizeSel.value = num(char.size, 1);
      boldB.classList.toggle("on", char.bold);
      italB.classList.toggle("on", char.italic);
      underB.classList.toggle("on", char.underline !== "none" && !!char.underline);
      colorB.style.background = char.color === "auto" ? "#000" : `#${char.color}`;
      for (const b of alignBtns) b.el.classList.toggle("on", para.align === b.v);
      for (const b of lineBtns) b.el.classList.toggle("on", para.lineRule === "auto" && Math.abs(para.line - b.v) < 0.01);
      const px = (pt: number) => `${pt * 4 / 3 * 0.5}px`;
      const grey = (t: string) => h("p", { class: "grey" }, `${t} `.repeat(t === "Vorhergehender Absatz" ? 6 : 10).trim());
      const main = h("p", { class: "main" }, sample);
      main.style.cssText = `font-family:"${char.font}";font-size:${Math.min(28, char.size) * 4 / 3}px;font-weight:${char.bold ? 700 : 400};font-style:${char.italic ? "italic" : "normal"};` +
        `color:${char.color === "auto" ? "#000" : `#${char.color}`};text-align:${para.align === "both" ? "justify" : para.align};margin:${px(para.before)} ${px(para.indRight)} ${px(para.after)} ${px(para.indLeft)};` +
        `text-indent:${px(para.indFirst)};${char.underline && char.underline !== "none" ? "text-decoration:underline;" : ""}${char.smallCaps ? "font-variant:small-caps;" : ""}${char.caps ? "text-transform:uppercase;" : ""}`;
      preview.replaceChildren(grey("Vorhergehender Absatz"), main, grey("Nächster Absatz"));
      const lines = describeFormat(draft.def.pPr, draft.def.rPr);
      const kind = typeSel.value === "linked" ? "Verknüpft" : typeSel.value === "character" ? "Zeichen" : "Absatz";
      const meta = [`Formatvorlage: ${kind}`, ...(gallery.checked ? ["Im Formatvorlagenkatalog anzeigen"] : []), ...(draft.def.uiPriority !== undefined ? [`Priorität: ${draft.def.uiPriority}`] : [])];
      if (draft.def.basedOn) meta.push(`Basiert auf: ${sheet.displayName(draft.def.basedOn) || draft.def.basedOn}`);
      if (draft.def.type === "paragraph" && nextSel.value) meta.push(`Folgende Formatvorlage: ${sheet.displayName(nextSel.value)}`);
      desc.textContent = [...lines, ...meta].join(", ");
    };
    fontIn.addEventListener("change", () => { if (fontIn.value.trim()) { draft.def = { ...draft.def, rPr: { ...draft.def.rPr, font: fontIn.value.trim(), fontTheme: undefined } }; refresh(); } });
    sizeSel.addEventListener("change", () => { const v = Number(sizeSel.value.replace(",", ".")); if (v > 0 && v <= 1638) { draft.def = { ...draft.def, rPr: { ...draft.def.rPr, size: v } }; refresh(); } });
    basedSel.onchange = () => { draft.def = { ...draft.def, basedOn: basedSel.value || undefined }; refresh(); };
    nextSel.onchange = () => { draft.def = { ...draft.def, next: nextSel.value || undefined }; refresh(); };
    typeSel.onchange = () => {
      const t = typeSel.value === "character" ? "character" : "paragraph";
      draft.def = { ...draft.def, type: t, basedOn: t === "character" ? "Absatz-Standardschriftart" : curPara, pPr: t === "character" ? {} : draft.def.pPr };
      fillSelects();
      refresh();
    };
    gallery.onchange = refresh;

    const charState = (): CharState => {
      const ch = look().char;
      return { font: ch.font, size: ch.size, bold: ch.bold, italic: ch.italic, underline: ch.underline, strike: ch.strike, dstrike: ch.dstrike,
        superscript: ch.vertAlign === "superscript", subscript: ch.vertAlign === "subscript", caps: ch.caps, smallCaps: ch.smallCaps, hidden: ch.hidden, color: ch.color };
    };
    const formatMenu = (ev: MouseEvent) => showMenu(c, [
      { label: "Zeichen...", run: () => void api.fontDialog({ state: charState(), apply: (r) => { draft.def = { ...draft.def, rPr: { ...draft.def.rPr, ...r, ...(r.font ? { fontTheme: undefined } : {}), ...(r.color ? { colorTheme: undefined } : {}) } }; refresh(); } }) },
      { label: "Absatz...", disabled: draft.def.type === "character", run: () => void api.paragraphDialog({ state: look().para, apply: (p) => { draft.def = { ...draft.def, pPr: { ...draft.def.pPr, ...p } }; refresh(); } }) },
      { label: "Tabstopps...", disabled: true, run: () => {} },
      { label: "Rahmen...", disabled: true, run: () => {} },
      { label: "Sprache...", run: () => languageDialog(draft.def.rPr.lang ?? "de-DE", (lang) => { draft.def = { ...draft.def, rPr: { ...draft.def.rPr, lang } }; refresh(); }) },
      { label: "Positionsrahmen...", disabled: true, run: () => {} },
      { label: "Nummerierung...", disabled: true, run: () => {} },
      { label: "Tastenkombination...", disabled: true, run: () => {} },
      { label: "Texteffekte...", disabled: true, run: () => {} },
    ], (ev.currentTarget as HTMLElement).getBoundingClientRect());

    const body = h("div", { class: "stydlg" },
      h("fieldset", {}, h("legend", {}, "Eigenschaften"),
        h("div", { class: "sty-grid" }, "Name:", nameIn, h("span", { class: isNew ? "" : "dim" }, "Formatvorlagentyp:"), typeSel,
          "Formatvorlage basiert auf:", basedSel, "Formatvorlage für folgenden Absatz:", nextSel)),
      h("fieldset", {}, h("legend", {}, "Formatierung"),
        h("div", { class: "fmt-row" }, fontIn, fontList, sizeSel, sizeList, h("span", { class: "sep" }), boldB, italB, underB, colorB),
        h("div", { class: "fmt-row" }, ...alignBtns.map((b) => b.el), h("span", { class: "sep" }), ...lineBtns.map((b) => b.el), h("span", { class: "sep" }), spaceUp, spaceDown, h("span", { class: "sep wide" }), indDown, indUp),
        preview, desc),
      h("div", { class: "row" }, h("label", { class: "check" }, gallery, "Zum Formatvorlagenkatalog hinzufügen"), h("label", { class: "check disabled", "data-tip": "Automatisch aktualisieren\nRaabeWrite ändert Formatvorlagen nur über „Ändern...“ oder „aktualisieren, um der Auswahl zu entsprechen“." }, autoUpd, "Automatisch aktualisieren")),
      h("div", { class: "row" }, h("label", { class: "check" }, onlyDoc, "Nur in diesem Dokument"), h("label", { class: "check disabled" }, newDocs, "Neue auf dieser Vorlage basierende Dokumente")));
    openModal(isNew ? "Neue Formatvorlage von Formatierung erstellen" : "Formatvorlage ändern", body, [
      { label: "Format", left: true, run: () => {} },
      { label: "OK", primary: true, run: () => {
        const name = nameIn.value.trim();
        if (!name) { void messageBox("RaabeWrite", "Geben Sie einen Namen für die Formatvorlage ein."); return false; }
        const clash = e.styleCatalog().find((x) => x.name.toLowerCase() === name.toLowerCase() && x.id !== id);
        if (clash) { void messageBox("RaabeWrite", "Es gibt bereits eine Formatvorlage mit diesem Namen. Geben Sie einen anderen Namen ein."); return false; }
        let def: StyleDef = { ...draft.def, qFormat: gallery.checked || undefined, semiHidden: gallery.checked ? undefined : draft.def.semiHidden };
        // Built-in styles keep their English name in the file; a new name becomes the style's name.
        if (base) { if (name !== sheet.displayName(base.id)) def = { ...def, name }; }
        else def = { ...def, id: e.newStyleId(name), name, next: def.type === "paragraph" ? nextSel.value || undefined : undefined };
        if (def.type === "paragraph" && !def.next && isNew) def = { ...def, next: def.id };
        e.defineStyle(def, isNew);
        lastKey = "";
        render();
        api.refresh();
        return true;
      } },
      { label: "Abbrechen", run: () => true },
    ], { width: 640, cls: "stydialog" });
    // The Format button opens its menu (Word's Format ▾).
    const fmtBtn = [...document.querySelectorAll<HTMLButtonElement>(".stydialog .dlg-buttons button")].find((b) => b.textContent === "Format");
    if (fmtBtn) { fmtBtn.append(icon("chevron-down")); fmtBtn.onclick = (ev) => formatMenu(ev); }
    refresh();
    nameIn.select();
  };

  const languageDialog = (cur: string, done: (lang: string) => void) => {
    const sel = h("select", { size: 8, style: "width:100%" }, ...LANGS.map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
    sel.value = LANGS.some(([v]) => v === cur) ? cur : "de-DE";
    openModal("Sprache", h("div", {}, h("label", { class: "field" }, "Ausgewählten Text markieren als:", sel)), [
      { label: "OK", primary: true, run: () => { done(sel.value); return true; } },
      { label: "Abbrechen", run: () => true },
    ], { width: 380 });
  };

  c.implement("style.pane", { run: () => setOpen(!open), enabled: () => !!ed(), checked: () => open });
  return { render: () => { render(); renderInspector(); }, isOpen: () => open, modify: (id: string) => { const e = ed(); if (e) styleDialog(e, id); }, create: () => { const e = ed(); if (e) styleDialog(e, null); } };
}

/** Word's name for a new style: Formatvorlage1, Formatvorlage2 … */
function newName(e: Editor): string {
  const names = new Set(e.styleCatalog().map((x) => x.name.toLowerCase()));
  let n = 1;
  while (names.has(`formatvorlage${n}`)) n++;
  return `Formatvorlage${n}`;
}

/** A new style takes the selection's character formatting that differs from its paragraph style. */
function newRPr(e: Editor, cs: CharState): CharProps {
  const base = e.styleLook(e.currentStyle()).char;
  const out: Record<string, unknown> = {};
  if (cs.font && cs.font !== base.font) out.font = cs.font;
  if (cs.size && cs.size !== base.size) out.size = cs.size;
  if (cs.bold !== undefined && cs.bold !== base.bold) out.bold = cs.bold;
  if (cs.italic !== undefined && cs.italic !== base.italic) out.italic = cs.italic;
  if (cs.color && cs.color !== base.color) out.color = cs.color;
  if (cs.underline && cs.underline !== base.underline) out.underline = cs.underline;
  return out as CharProps;
}
