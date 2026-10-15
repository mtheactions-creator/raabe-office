// Aufzählungszeichen, Nummerierung and Liste mit mehreren Ebenen: Word's galleries (the
// library entries Word 365 shows) and the buttons, which apply the last chosen entry.

import { bulletLevels, displayBullet, formatNumber, multilevelLevels, numberLevels, type CharProps, type NumLevel } from "@raabe/text-engine";
import { h, hoverPreview, showDialog, showMenu, type Command } from "@raabe/shell";
import type { AppApi } from "./api";
import type { Recent } from "./symbols";

interface Entry {
  label: string;
  levels: () => NumLevel[];
  /** Label shown on each of the preview's three lines. */
  preview: string[];
  font?: string;
  /** Heading numbering: linked to Überschrift 1-3. */
  headings?: boolean;
}

// Word 365's libraries (ListGalleries, read 2026-10-14). Word's fourth bullet is a picture of
// its own; RaabeWrite has no picture bullets, so that entry is left out.
const BULLETS: [string, string][] = [
  ["\uf0b7", "Symbol"], ["o", "Courier New"], ["\uf0a7", "Wingdings"], ["\uf076", "Wingdings"],
  ["\uf0d8", "Wingdings"], ["\uf0fc", "Wingdings"],
];

const bulletEntry = ([ch, font]: [string, string]): Entry => ({
  label: "Aufzählungszeichen",
  levels: () => { const l = bulletLevels(); l[0] = { ...l[0], text: ch, rPr: { font } }; return l; },
  preview: [ch, ch, ch], font,
});

/** A numbering library entry; `pos` is where the number stands (right-aligned ones end there). */
const numberEntry = (format: string, text: string, labels: string[], align: "left" | "right" = "left", pos = 18): Entry => ({
  label: "Nummerierung",
  levels: () => { const l = numberLevels(); l[0] = { ...l[0], format, text, align, pPr: { indLeft: 36, indFirst: pos - 36 } }; return l; },
  preview: labels,
});

const NUMBERS: Entry[] = [
  numberEntry("decimal", "%1.", ["1.", "2.", "3."]),
  numberEntry("decimal", "%1)", ["1)", "2)", "3)"]),
  numberEntry("upperRoman", "%1.", ["I.", "II.", "III."], "right", 27),
  numberEntry("decimal", "(%1)", ["(1)", "(2)", "(3)"]),
  numberEntry("lowerLetter", "%1)", ["a)", "b)", "c)"]),
  numberEntry("lowerLetter", "%1.", ["a.", "b.", "c."]),
  numberEntry("lowerRoman", "%1.", ["i.", "ii.", "iii."], "right"),
];

type L = [format: string, text: string, pos: number, textPos: number, align?: "left" | "right", font?: string, suffix?: "tab" | "space" | "nothing"];
const levelsOf = (rows: L[], legal?: number): NumLevel[] => rows.map(([format, text, pos, textPos, align = "left", font, suffix = "tab"], i) => ({
  start: 1, format, text, align, suffix, pPr: { indLeft: textPos, indFirst: pos - textPos }, rPr: font ? { font } : {},
  legal: legal === i ? true : undefined,
}));

const OUTLINE_1 = (): NumLevel[] => levelsOf([
  ["decimal", "%1)", 0, 18], ["lowerLetter", "%2)", 18, 36], ["lowerRoman", "%3)", 36, 54],
  ["decimal", "(%4)", 54, 72], ["lowerLetter", "(%5)", 72, 90], ["lowerRoman", "(%6)", 90, 108],
  ["decimal", "%7.", 108, 126], ["lowerLetter", "%8.", 126, 144], ["lowerRoman", "%9.", 144, 162],
]);
const OUTLINE_BULLETS = (): NumLevel[] => levelsOf([
  ["bullet", "\uf076", 0, 18, "left", "Wingdings"], ["bullet", "\uf0d8", 18, 36, "left", "Wingdings"], ["bullet", "\uf0a7", 36, 54, "left", "Wingdings"],
  ["bullet", "\uf0b7", 54, 72, "left", "Symbol"], ["bullet", "\uf0a8", 72, 90, "left", "Symbol"], ["bullet", "\uf0d8", 90, 108, "left", "Wingdings"],
  ["bullet", "\uf0a7", 108, 126, "left", "Wingdings"], ["bullet", "\uf0b7", 126, 144, "left", "Symbol"], ["bullet", "\uf0a8", 144, 162, "left", "Symbol"],
]);
/** Heading lists: "Artikel I." / "Abschnitt 1.01" / "(a)" … */
const HEADINGS_ARTICLE = (): NumLevel[] => levelsOf([
  ["upperRoman", "Artikel %1.", 0, 0], ["decimalZero", "Abschnitt %1.%2", 0, 0], ["lowerLetter", "(%3)", 14.4, 36],
  ["lowerRoman", "(%4)", 36, 43.2, "right"], ["decimal", "%5)", 28.8, 50.4], ["lowerLetter", "%6)", 36, 57.6],
  ["lowerRoman", "%7)", 50.4, 64.8, "right"], ["lowerLetter", "%8.", 50.4, 72], ["lowerRoman", "%9.", 72, 79.2, "right"],
], 1);
const HEADINGS_ROMAN = (): NumLevel[] => levelsOf([
  ["upperRoman", "%1.", 0, 0], ["upperLetter", "%2.", 36, 36], ["decimal", "%3.", 72, 72],
  ["lowerLetter", "%4)", 108, 108], ["decimal", "(%5)", 144, 144], ["lowerLetter", "(%6)", 180, 180],
  ["lowerRoman", "(%7)", 216, 216], ["lowerLetter", "(%8)", 252, 252], ["lowerRoman", "(%9)", 288, 288],
]);
const HEADINGS_CHAPTER = (): NumLevel[] => levelsOf([
  ["decimal", "Kapitel %1", 0, 0, "left", undefined, "space"],
  ...Array.from({ length: 8 }, (): L => ["none", "", 0, 0, "left", undefined, "nothing"]),
]);

export const headingLevels = (): NumLevel[] => Array.from({ length: 9 }, (_, i) => ({
  start: 1, format: "decimal", text: Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join("."), align: "left", suffix: "tab",
  pPr: { indLeft: [21.6, 28.8, 36, 43.2, 50.4, 57.6, 64.8, 72, 79.2][i], indFirst: -[21.6, 28.8, 36, 43.2, 50.4, 57.6, 64.8, 72, 79.2][i] }, rPr: {},
}));

const MULTI: Entry[] = [
  { label: "1) a) i)", levels: OUTLINE_1, preview: ["1)", "a)", "i)"] },
  { label: "1. 1.1. 1.1.1.", levels: multilevelLevels, preview: ["1.", "1.1.", "1.1.1."] },
  { label: "Aufzählung mit mehreren Ebenen", levels: OUTLINE_BULLETS, preview: ["\uf076", "\uf0d8", "\uf0a7"], font: "Wingdings" },
  { label: "Artikel / Abschnitt", levels: HEADINGS_ARTICLE, preview: ["Artikel I. Überschrift 1", "Abschnitt 1.01 Überschrift 2", "(a) Überschrift 3"], headings: true },
  { label: "1 Überschrift 1, 1.1 Überschrift 2", levels: headingLevels, preview: ["1 Überschrift 1", "1.1 Überschrift 2", "1.1.1 Überschrift 3"], headings: true },
  { label: "I. Überschrift 1, A. Überschrift 2", levels: HEADINGS_ROMAN, preview: ["I. Überschrift 1", "A. Überschrift 2", "1. Überschrift 3"], headings: true },
  { label: "Kapitel 1 Überschrift 1", levels: HEADINGS_CHAPTER, preview: ["Kapitel 1 Überschrift 1", "Überschrift 2", "Überschrift 3"], headings: true },
];

/** A gallery tile: three lines, each a label and a grey bar (Word's list previews). */
function tile(e: Entry, deep = false): HTMLElement {
  const lines = e.preview.map((label, i) => {
    const shown = displayBullet(label, e.font);
    return h("div", { style: `display:flex;align-items:center;gap:4px;margin:5px 0 5px ${deep ? i * 8 : 0}px;font-size:11px;white-space:nowrap` },
      h("span", { style: `font-family:"${shown.font ?? "Calibri"}";min-width:12px` }, shown.text),
      e.headings ? null : h("span", { style: "flex:1;height:2px;background:#9e9e9e;min-width:30px" }));
  });
  return h("div", { style: "width:64px;overflow:hidden" }, ...lines);
}

export const LIST_COMMANDS: Command[] = [
  { id: "list.defineBullet", label: "Neues Aufzählungszeichen definieren..." },
  { id: "list.defineNumber", label: "Neues Zahlenformat definieren..." },
  // Word's dialogs for these set nine levels with every option; not built yet.
  { id: "list.defineMultilevel", label: "Neue Liste mit mehreren Ebenen definieren...", deferred: true },
  { id: "list.defineStyle", label: "Neuen Listentyp definieren...", deferred: true },
];

/** Zahlenformatvorlage in Word's order (recorded 2026-10-15, probe_word_list_define.py). */
const NUMBER_STYLES: [string, string][] = [
  ["none", "(ohne)"], ["decimal", "1, 2, 3, …"], ["upperRoman", "I, II, III, …"], ["lowerRoman", "i, ii, iii, …"],
  ["upperLetter", "A, B, C, …"], ["lowerLetter", "a, b, c, …"], ["ordinal", "1. 2. 3. …"], ["cardinalText", "Eins Zwei Drei …"],
  ["ordinalText", "Erste, Zweite, Dritte …"], ["decimalZero", "01, 02, 03, …"], ["custom:001, 002, 003, ...", "001, 002, 003, …"],
  ["custom:0001, 0002, 0003, ...", "0001, 0002, 0003, …"], ["custom:00001, 00002, 00003, ...", "00001, 00002, 00003, …"],
];
const ALIGNS: ["left" | "center" | "right", string][] = [["left", "Links"], ["center", "Zentriert"], ["right", "Rechts"]];

/** The dialog Schriftart's view of a level's character formatting. */
const charState = (r: CharProps) => ({ font: r.font, size: r.size, bold: r.bold, italic: r.italic, color: r.color, underline: r.underline as string | undefined });

export function listsFeature(api: AppApi, chooseSymbol: (start: Recent, pick: (r: Recent) => void) => void) {
  const c = api.commands;
  let lastBullet = bulletEntry(BULLETS[0]);
  let lastNumber = NUMBERS[0];
  let lastMulti = MULTI[0];
  const ed = () => api.editor();

  const change = (kind: "bullet" | "number" | "multilevel", e: Entry | null) => {
    const editor = ed();
    if (!editor) return;
    if (!e) { if (editor.listKind()) editor.toggleList(editor.listKind()!); }
    else if (e.headings) editor.numberHeadings(e.levels());
    else editor.toggleList(kind, e.levels());
  };
  const apply = (kind: "bullet" | "number" | "multilevel", e: Entry | null) => { change(kind, e); api.focus(); };
  /** A library tile with Livevorschau on the paragraphs (Word's galleries preview the list). */
  const live = (b: HTMLElement, kind: "bullet" | "number" | "multilevel", e: Entry | null) => { hoverPreview(c, b, () => change(kind, e)); return b; };

  const panel = (title: string, entries: Entry[], kind: "bullet" | "number" | "multilevel", remember: (e: Entry) => void, close: () => void, deep = false) => {
    const tiles = h("div", { class: "tiles" },
      live(h("button", { class: "tile", "data-tip": "Ohne", onclick: () => { close(); apply(kind, null); } }, h("span", { style: "font-size:12px" }, "Ohne")), kind, null),
      ...entries.map((e) => live(h("button", {
        class: "tile", "data-tip": e.label,
        onclick: () => { close(); remember(e); apply(kind, e); },
      }, tile(e, deep)), kind, e)));
    const inList = !!ed()?.listKind();
    const link = (label: string, cmd: string | null, run?: () => void, disabled = false) =>
      h("button", { class: "link", disabled: disabled || (cmd ? !c.isEnabled(cmd) : false), onclick: () => { close(); if (run) run(); else if (cmd) void c.run(cmd); } }, label);
    const levels = h("button", {
      class: "link has-sub", disabled: !inList,
      onclick: (ev: MouseEvent) => {
        const r = (ev.currentTarget as HTMLElement).getBoundingClientRect();
        close();
        showMenu(c, Array.from({ length: 9 }, (_, i) => ({ label: `Ebene ${i + 1}`, run: () => { ed()?.setListLevel(i); api.focus(); } })), new DOMRect(r.right, r.top, 0, 0));
      },
    }, "Listenebene ändern");
    const extra: HTMLElement[] = [h("hr"), levels];
    if (kind === "bullet") extra.push(link("Neues Aufzählungszeichen definieren...", "list.defineBullet"));
    else if (kind === "number") extra.push(link("Neues Zahlenformat definieren...", "list.defineNumber"), link("Nummerierungswert festlegen...", "list.setValue"));
    else extra.push(link("Neue Liste mit mehreren Ebenen definieren...", "list.defineMultilevel"), link("Neuen Listentyp definieren...", "list.defineStyle"));
    return h("div", { class: "rpanel" }, h("div", { class: "title" }, title), tiles, ...extra);
  };

  /** Vorschau: three lines, each the label and grey text bars, as in Word's dialogs. */
  const preview = (label: (i: number) => { text: string; style: string }, align: string) => {
    const box = h("div", { style: "border:1px solid var(--line, #c8c8c8);background:#fff;padding:14px 18px;height:150px;display:flex;flex-direction:column;justify-content:center;gap:16px" });
    for (let i = 0; i < 3; i++) {
      const l = label(i);
      box.append(h("div", { style: "display:flex;align-items:center;gap:8px" },
        h("span", { style: `width:34px;text-align:${align};${l.style}` }, l.text),
        h("span", { style: "flex:1;display:flex;flex-direction:column;gap:4px" }, h("i", { style: "height:3px;background:#c8c8c8" }), h("i", { style: "height:3px;background:#c8c8c8;width:70%" }))));
    }
    return box;
  };
  const css = (r: CharProps, font?: string) =>
    `font-family:"${font ?? r.font ?? "Calibri"}";${r.bold ? "font-weight:bold;" : ""}${r.italic ? "font-style:italic;" : ""}${r.color && r.color !== "auto" ? `color:#${r.color};` : ""}`;
  const select = (options: [string, string][], value: string) => {
    const el = h("select", { style: "width:210px" }, ...options.map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
    el.value = value;
    return el;
  };

  // Neues Aufzählungszeichen definieren (Word's dialog, recorded): Symbol..., Bild..., Schriftart...,
  // Ausrichtung, Vorschau. In a bullet list it changes the caret's level, elsewhere it starts a list.
  c.implement("list.defineBullet", { run: async () => {
    const e = ed();
    if (!e) return;
    const cur = e.listKind() === "bullet" ? e.listLevelDef() : null;
    let ch = cur?.level.text ?? "\uf0b7";
    let rPr: CharProps = cur?.level.rPr ?? { font: "Symbol" };
    const align = select(ALIGNS, cur?.level.align ?? "left");
    const view = h("div");
    const show = () => {
      const shown = displayBullet(ch, rPr.font);
      view.replaceChildren(preview(() => ({ text: shown.text, style: css(rPr, shown.font) }), align.value));
    };
    align.onchange = show;
    const symbol = h("button", { onclick: () => chooseSymbol({ c: ch, f: rPr.font, s: ch.charCodeAt(0) >= 0xf000 || undefined }, (r) => { ch = r.c; rPr = { ...rPr, font: r.f ?? rPr.font }; show(); }) }, "Symbol...");
    const picture = h("button", { disabled: true, "data-tip": "Bild...\nRaabeWrite hat keine Bilder als Aufzählungszeichen." }, "Bild...");
    const font = h("button", { onclick: () => void api.fontDialog({ state: charState(rPr), apply: (r) => { rPr = { ...rPr, ...r }; show(); } }) }, "Schriftart...");
    show();
    const ok = await showDialog({
      title: "Neues Aufzählungszeichen definieren",
      body: h("div", { style: "width:400px" },
        h("fieldset", null, h("legend", null, "Aufzählungszeichen"), h("div", { style: "display:flex;gap:12px" }, symbol, picture, font)),
        h("label", { class: "field" }, "Ausrichtung:", align),
        h("fieldset", null, h("legend", null, "Vorschau"), view)),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    const a = align.value as "left" | "center" | "right";
    if (cur) e.updateListLevel((l) => ({ ...l, format: "bullet", text: ch, rPr, align: a }));
    else {
      const entry: Entry = { label: "Aufzählungszeichen", preview: [ch, ch, ch], font: rPr.font,
        levels: () => { const l = bulletLevels(); l[0] = { ...l[0], text: ch, rPr, align: a }; return l; } };
      lastBullet = entry;
      apply("bullet", entry);
    }
    api.focus();
  }, enabled: () => !!ed() });

  // Neues Zahlenformat definieren (Word's dialog, recorded): Zahlenformatvorlage, Schriftart...,
  // Zahlenformat (the number shaded inside the text around it), Ausrichtung, Vorschau.
  c.implement("list.defineNumber", { run: async () => {
    const e = ed();
    if (!e) return;
    const cur = e.listKind() === "number" ? e.listLevelDef() : null;
    const ilvl = cur?.ilvl ?? 0;
    const ph = `%${ilvl + 1}`;
    let rPr: CharProps = cur?.level.rPr ?? {};
    const startFmt = cur && NUMBER_STYLES.some(([v]) => v === cur.level.format) ? cur.level.format : "decimal";
    const style = select(NUMBER_STYLES, startFmt);
    // The text around the number; the number itself is the shaded part (Word's field).
    const parts = (cur?.level.text ?? `${ph}.`).split(ph);
    const before = h("input", { type: "text", value: parts[0] ?? "", style: "width:70px" }) as HTMLInputElement;
    const after = h("input", { type: "text", value: parts.slice(1).join(ph), style: "width:70px" }) as HTMLInputElement;
    const sample = h("span", { style: "background:#d9d9d9;padding:2px 4px;border-radius:2px" });
    const align = select(ALIGNS, cur?.level.align ?? "left");
    const view = h("div");
    const show = () => {
      sample.textContent = formatNumber(1, style.value) || "\u2003";
      view.replaceChildren(preview((i) => ({ text: `${before.value}${formatNumber(i + 1, style.value)}${after.value}`, style: css(rPr) }), align.value));
    };
    for (const el of [style, align]) el.onchange = show;
    for (const el of [before, after]) el.oninput = show;
    const font = h("button", { onclick: () => void api.fontDialog({ state: charState(rPr), apply: (r) => { rPr = { ...rPr, ...r }; show(); } }) }, "Schriftart...");
    show();
    const ok = await showDialog({
      title: "Neues Zahlenformat definieren",
      body: h("div", { style: "width:380px" },
        h("fieldset", null, h("legend", null, "Zahlenformat"),
          h("label", { class: "field" }, "Zahlenformatvorlage:", h("div", { style: "display:flex;gap:12px" }, style, font)),
          h("label", { class: "field" }, "Zahlenformat:", h("div", { style: "display:flex;align-items:center;gap:2px" }, before, sample, after)),
          h("label", { class: "field" }, "Ausrichtung:", align)),
        h("fieldset", null, h("legend", null, "Vorschau"), view)),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    const format = style.value, text = format === "none" ? `${before.value}${after.value}` : `${before.value}${ph}${after.value}`;
    const a = align.value as "left" | "center" | "right";
    if (cur) e.updateListLevel((l) => ({ ...l, format, text, rPr, align: a }));
    else {
      const entry: Entry = { label: "Nummerierung", preview: [1, 2, 3].map((n) => `${before.value}${formatNumber(n, format)}${after.value}`),
        levels: () => { const l = numberLevels(); l[0] = { ...l[0], format, text: text.replace(ph, "%1"), rPr, align: a }; return l; } };
      lastNumber = entry;
      apply("number", entry);
    }
    api.focus();
  }, enabled: () => !!ed() });

  const has = () => !!ed();
  c.implement("para.bullets", { run: () => apply("bullet", lastBullet), enabled: has, checked: () => ed()?.listKind() === "bullet" });
  c.implement("para.numbering", { run: () => apply("number", lastNumber), enabled: has, checked: () => ed()?.listKind() === "number" });
  c.implement("para.multilevel", { run: () => apply("multilevel", lastMulti), enabled: has, checked: () => ed()?.listKind() === "multilevel" });

  return {
    bullets: (close: () => void) => panel("Aufzählungszeichenbibliothek", BULLETS.map(bulletEntry), "bullet", (e) => { lastBullet = e; }, close),
    numbers: (close: () => void) => panel("Nummerierungsbibliothek", NUMBERS, "number", (e) => { lastNumber = e; }, close),
    multilevel: (close: () => void) => panel("Listenbibliothek", MULTI, "multilevel", (e) => { lastMulti = e; }, close, true),
  };
}
