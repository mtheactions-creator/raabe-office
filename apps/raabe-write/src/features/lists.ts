// Aufzählungszeichen, Nummerierung and Liste mit mehreren Ebenen: Word's galleries (the
// library entries Word 365 shows) and the buttons, which apply the last chosen entry.

import { bulletLevels, displayBullet, multilevelLevels, numberLevels, type NumLevel } from "@raabe/text-engine";
import { h } from "@raabe/shell";
import type { AppApi } from "./api";

interface Entry {
  label: string;
  levels: () => NumLevel[];
  /** Label shown on each of the preview's three lines. */
  preview: string[];
  font?: string;
  /** Heading numbering: linked to Überschrift 1-3. */
  headings?: boolean;
}

const BULLETS: [string, string][] = [
  ["\uf0b7", "Symbol"], ["o", "Courier New"], ["\uf0a7", "Wingdings"], ["\uf076", "Wingdings"],
  ["\uf0d8", "Wingdings"], ["\uf0fc", "Wingdings"], ["–", "Calibri"],
];

const bulletEntry = ([ch, font]: [string, string]): Entry => ({
  label: "Aufzählungszeichen",
  levels: () => { const l = bulletLevels(); l[0] = { ...l[0], text: ch, rPr: { font } }; return l; },
  preview: [ch, ch, ch], font,
});

const numberEntry = (format: string, text: string, labels: string[], align: "left" | "right" = "left"): Entry => ({
  label: "Nummerierung",
  levels: () => { const l = numberLevels(); l[0] = { ...l[0], format, text, align }; return l; },
  preview: labels,
});

const NUMBERS: Entry[] = [
  numberEntry("decimal", "%1.", ["1.", "2.", "3."]),
  numberEntry("decimal", "%1)", ["1)", "2)", "3)"]),
  numberEntry("upperRoman", "%1.", ["I.", "II.", "III."], "right"),
  numberEntry("upperLetter", "%1.", ["A.", "B.", "C."]),
  numberEntry("lowerLetter", "%1)", ["a)", "b)", "c)"]),
  numberEntry("lowerLetter", "%1.", ["a.", "b.", "c."]),
  numberEntry("lowerRoman", "%1.", ["i.", "ii.", "iii."], "right"),
];

const outline = (formats: string[], suffix: string): NumLevel[] => Array.from({ length: 9 }, (_, i) => ({
  start: 1, format: formats[i % formats.length], text: `%${i + 1}${suffix}`, align: "left", suffix: "tab",
  pPr: { indLeft: 18 * (i + 1), indFirst: -18 }, rPr: {},
}));

const headingLevels = (): NumLevel[] => Array.from({ length: 9 }, (_, i) => ({
  start: 1, format: "decimal", text: Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join("."), align: "left", suffix: "tab",
  pPr: { indLeft: [21.6, 28.8, 36, 43.2, 50.4, 57.6, 64.8, 72, 79.2][i], indFirst: -[21.6, 28.8, 36, 43.2, 50.4, 57.6, 64.8, 72, 79.2][i] }, rPr: {},
}));

const MULTI: Entry[] = [
  { label: "1. 1.1. 1.1.1.", levels: multilevelLevels, preview: ["1.", "1.1.", "1.1.1."] },
  { label: "1) a) i)", levels: () => outline(["decimal", "lowerLetter", "lowerRoman"], ")"), preview: ["1)", "a)", "i)"] },
  { label: "Aufzählung", levels: bulletLevels, preview: ["•", "o", "▪"] },
  { label: "Überschriften nummerieren", levels: headingLevels, preview: ["1 Überschrift 1", "1.1 Überschrift 2", "1.1.1 Überschrift 3"], headings: true },
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

export function listsFeature(api: AppApi) {
  let lastBullet = bulletEntry(BULLETS[0]);
  let lastNumber = NUMBERS[0];
  let lastMulti = MULTI[0];
  const ed = () => api.editor();

  const apply = (kind: "bullet" | "number" | "multilevel", e: Entry | null) => {
    const editor = ed();
    if (!editor) return;
    if (!e) { if (editor.listKind()) editor.toggleList(editor.listKind()!); }
    else if (e.headings) editor.numberHeadings(e.levels());
    else editor.toggleList(kind, e.levels());
    api.focus();
  };

  const panel = (title: string, entries: Entry[], kind: "bullet" | "number" | "multilevel", remember: (e: Entry) => void, close: () => void, deep = false) => {
    const tiles = h("div", { class: "tiles" },
      h("button", { class: "tile", "data-tip": "Ohne", onclick: () => { close(); apply(kind, null); } }, h("span", { style: "font-size:12px" }, "Ohne")),
      ...entries.map((e) => h("button", {
        class: "tile", "data-tip": e.label,
        onclick: () => { close(); remember(e); apply(kind, e); },
      }, tile(e, deep))));
    const extra: HTMLElement[] = [];
    if (kind !== "bullet") {
      extra.push(h("hr"), h("button", { class: "link", onclick: () => { close(); ed()?.restartNumbering(); api.focus(); } }, "Neu beginnen mit 1"));
    }
    extra.push(h("button", { class: "link", onclick: () => { close(); ed()?.changeListLevel(1); api.focus(); } }, "Eine Ebene tiefer (Tab)"),
      h("button", { class: "link", onclick: () => { close(); ed()?.changeListLevel(-1); api.focus(); } }, "Eine Ebene höher (Umschalt+Tab)"));
    return h("div", { class: "rpanel" }, h("div", { class: "title" }, title), tiles, ...extra);
  };

  const c = api.commands;
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
