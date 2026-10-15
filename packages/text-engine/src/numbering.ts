// Lists: Word's numbering definitions (numbering.xml) turned into the label each paragraph
// shows ("1.", "a)", "•"). Counters run per abstract definition in document order; a level
// restarts after a higher level, and a w:num with a start override restarts the list.

import type { AbstractNum, CharProps, NumLevel, Numbering, Para, ParaProps } from "./model";
import type { StyleSheet } from "./styles";

export interface ListLabel {
  readonly text: string;
  readonly level: NumLevel;
  /** Formatting of the label: the paragraph mark's plus the level's run properties. */
  readonly props: CharProps;
  readonly numId: number;
  readonly ilvl: number;
}

/** Bullets of symbol fonts as Unicode (Symbol and Wingdings use private-use code points). */
const SYMBOL_BULLETS: Record<string, string> = {
  "\uf0b7": "•", "\uf0a7": "▪", "\uf0d8": "➢", "\uf0fc": "✓", "\uf076": "❖", "\uf0a8": "○",
  "\uf06e": "■", "\uf071": "❑", "\uf0e0": "➢", "\uf0de": "➤", "\uf0d7": "▸",
};

/** The level text with symbol-font bullets mapped to Unicode, and its font. */
export function displayBullet(text: string, font: string | undefined): { text: string; font: string | undefined } {
  const mapped = [...text].map((ch) => SYMBOL_BULLETS[ch] ?? ch).join("");
  const symbolFont = font && /^(symbol|wingdings|webdings)/i.test(font);
  return { text: mapped, font: symbolFont && mapped !== text ? "Calibri" : font };
}

const ROMAN: [number, string][] = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];

function roman(n: number): string {
  let out = "";
  for (const [v, s] of ROMAN) while (n >= v) { out += s; n -= v; }
  return out;
}

function letters(n: number): string {
  // Word: a … z, aa … zz, aaa … (the letter repeats).
  if (n <= 0) return "";
  const i = (n - 1) % 26, times = Math.floor((n - 1) / 26) + 1;
  return String.fromCharCode(97 + i).repeat(times);
}

const ORDINAL_DE = (n: number) => `${n}.`;

const ONES = ["", "eins", "zwei", "drei", "vier", "fünf", "sechs", "sieben", "acht", "neun", "zehn", "elf", "zwölf", "dreizehn", "vierzehn", "fünfzehn", "sechzehn", "siebzehn", "achtzehn", "neunzehn"];
const TENS = ["", "", "zwanzig", "dreißig", "vierzig", "fünfzig", "sechzig", "siebzig", "achtzig", "neunzig"];

/** German number words as Word writes them (cardinalText, lower case; "ein" inside compounds). */
function wordsDe(n: number): string {
  if (n <= 0) return "null";
  if (n >= 1000) return String(n);
  const h = Math.floor(n / 100), r = n % 100;
  const head = h ? `${h === 1 ? "ein" : ONES[h]}hundert` : "";
  if (!r) return head;
  if (r < 20) return head + (r === 1 && h ? "eins" : ONES[r]);
  const t = Math.floor(r / 10), o = r % 10;
  return head + (o ? `${o === 1 ? "ein" : ONES[o]}und` : "") + TENS[t];
}

const cap = (s: string) => s.charAt(0).toUpperCase() + s.slice(1);

/** German ordinal words (ordinalText): Erste, Zweite, Dritte … Zwanzigste. */
function ordinalWordsDe(n: number): string {
  const r = n % 100;
  if (r > 0 && r < 20) {
    const stem = r === 1 ? "ers" : r === 3 ? "drit" : r === 7 ? "sieb" : r === 8 ? "ach" : ONES[r];
    const head = n - r ? wordsDe(n - r) : "";
    return `${head}${stem}te`;
  }
  return `${wordsDe(n)}ste`;
}

export function formatNumber(n: number, format: string): string {
  switch (format) {
    case "decimal": return String(n);
    case "decimalZero": return n < 10 ? `0${n}` : String(n);
    case "lowerLetter": return letters(n);
    case "upperLetter": return letters(n).toUpperCase();
    case "lowerRoman": return roman(n);
    case "upperRoman": return roman(n).toUpperCase();
    case "ordinal": return ORDINAL_DE(n);
    case "cardinalText": return cap(wordsDe(n));
    case "ordinalText": return cap(ordinalWordsDe(n));
    case "none": return "";
    case "bullet": return "";
    default: {
      // Word 2010's custom formats ("001, 002, 003, ..."): zero-padded numbers.
      const m = /^custom:(0+)1\b/.exec(format);
      return m ? String(n).padStart(m[1].length + 1, "0") : String(n);
    }
  }
}

/** The effective level definition of a numbering instance. */
export function levelOf(numbering: Numbering, numId: number, ilvl: number): { level: NumLevel; abstract: AbstractNum } | null {
  const num = numbering.nums.get(numId);
  if (!num) return null;
  const abs = numbering.abstracts.get(num.abstractId);
  if (!abs) return null;
  const level = num.overrides?.get(ilvl)?.level ?? abs.levels[ilvl];
  return level ? { level, abstract: abs } : null;
}

/** numId and level of a paragraph: direct numbering, else from its style. */
export function numberingOf(p: ParaProps, styles: StyleSheet): { numId: number; ilvl: number } | null {
  if (p.numId !== undefined) return p.numId > 0 ? { numId: p.numId, ilvl: p.ilvl ?? 0 } : null;
  const s = styles.numberingOfStyle(p.style ?? styles.defaultParagraphStyle);
  return s && s.numId > 0 ? { numId: s.numId, ilvl: p.ilvl ?? s.ilvl } : null;
}

/** List labels for every paragraph of a story (undefined where there is none). */
export function listLabels(paras: readonly Para[], numbering: Numbering, styles: StyleSheet): (ListLabel | undefined)[] {
  const out: (ListLabel | undefined)[] = new Array(paras.length);
  if (!numbering.nums.size) return out;
  // Counters per abstract definition, and which num instance last used each.
  const counters = new Map<number, number[]>();
  const lastNum = new Map<number, number>();
  for (let i = 0; i < paras.length; i++) {
    const p = paras[i];
    if (p.block) continue;
    const n = numberingOf(p.props, styles);
    if (!n) continue;
    const found = levelOf(numbering, n.numId, n.ilvl);
    if (!found) continue;
    const { level, abstract } = found;
    let c = counters.get(abstract.id);
    if (!c) { c = new Array(9).fill(0); counters.set(abstract.id, c); }
    const num = numbering.nums.get(n.numId)!;
    // A new num instance with a start override restarts its levels.
    if (lastNum.get(abstract.id) !== n.numId) {
      if (lastNum.has(abstract.id) && num.overrides) {
        for (const [lv, ov] of num.overrides) if (ov.start !== undefined) c[lv] = 0;
      }
      lastNum.set(abstract.id, n.numId);
    }
    const startOf = (lv: number) => num.overrides?.get(lv)?.start ?? (num.overrides?.get(lv)?.level ?? abstract.levels[lv])?.start ?? 1;
    c[n.ilvl] = c[n.ilvl] === 0 ? startOf(n.ilvl) : c[n.ilvl] + 1;
    // Deeper levels restart (unless their definition says otherwise).
    for (let lv = n.ilvl + 1; lv < 9; lv++) {
      const def = num.overrides?.get(lv)?.level ?? abstract.levels[lv];
      const restart = def?.restart;
      if (restart === 0) continue;
      if (restart === undefined || restart - 1 <= n.ilvl) c[lv] = 0;
    }
    let text: string;
    if (level.format === "bullet") text = level.text;
    else {
      text = level.text.replace(/%([1-9])/g, (_m, d: string) => {
        const lv = Number(d) - 1;
        const def = num.overrides?.get(lv)?.level ?? abstract.levels[lv];
        const value = c[lv] === 0 ? startOf(lv) : c[lv];
        const fmt = level.legal ? "decimal" : def?.format ?? "decimal";
        return formatNumber(value, fmt);
      });
    }
    out[i] = { text, level, props: { ...p.mark, ...level.rPr, extra: undefined, rev: undefined }, numId: n.numId, ilvl: n.ilvl };
  }
  return out;
}

// ----- built-in lists (Aufzählungszeichen, Nummerierung, Liste mit mehreren Ebenen) -----------

const lvl = (i: number, format: string, text: string, font?: string): NumLevel => ({
  start: 1, format, text, align: "left", suffix: "tab",
  pPr: { indLeft: 36 * (i + 1), indFirst: -18 },
  rPr: font ? { font } : {},
});

/** Word's default bullet list (•, o, ▪ repeating), as created by "Aufzählungszeichen". */
export function bulletLevels(): NumLevel[] {
  const symbols: [string, string][] = [["\uf0b7", "Symbol"], ["o", "Courier New"], ["\uf0a7", "Wingdings"]];
  return Array.from({ length: 9 }, (_, i) => lvl(i, "bullet", symbols[i % 3][0], symbols[i % 3][1]));
}

/** Word's default numbered list (1. a. i. repeating), as created by "Nummerierung". */
export function numberLevels(): NumLevel[] {
  const formats = ["decimal", "lowerLetter", "lowerRoman"];
  return Array.from({ length: 9 }, (_, i) => ({
    ...lvl(i, formats[i % 3], `%${i + 1}.`),
    align: formats[i % 3] === "lowerRoman" ? "right" as const : "left" as const,
    pPr: { indLeft: 36 * (i + 1), indFirst: formats[i % 3] === "lowerRoman" ? -9 : -18 },
  }));
}

/** Word's "1. 1.1. 1.1.1." multilevel list. */
export function multilevelLevels(): NumLevel[] {
  return Array.from({ length: 9 }, (_, i) => ({
    ...lvl(i, "decimal", Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join(".") + "."),
    pPr: { indLeft: [18, 39.6, 61.2, 86.4, 111.6, 136.8, 162, 187.2, 216][i], indFirst: -[18, 21.6, 25.2, 32.4, 39.6, 46.8, 54, 61.2, 72][i] },
  }));
}

/** A copy of `n` with a new list (abstract definition and instance); returns the new numId. */
export function addList(n: Numbering, levels: NumLevel[]): { numbering: Numbering; numId: number } {
  const absId = Math.max(-1, ...n.abstracts.keys()) + 1;
  const numId = Math.max(0, ...n.nums.keys()) + 1;
  const abstracts = new Map(n.abstracts);
  abstracts.set(absId, { id: absId, levels });
  const nums = new Map(n.nums);
  nums.set(numId, { id: numId, abstractId: absId });
  return { numbering: { ...n, abstracts, nums }, numId };
}

/** A new instance of an existing list that restarts at `start` on `ilvl` ("Neu beginnen mit"). */
export function restartList(n: Numbering, numId: number, ilvl: number, start = 1): { numbering: Numbering; numId: number } {
  const base = n.nums.get(numId);
  if (!base) return { numbering: n, numId };
  const id = Math.max(0, ...n.nums.keys()) + 1;
  const nums = new Map(n.nums);
  nums.set(id, { id, abstractId: base.abstractId, overrides: new Map([[ilvl, { start }]]) });
  return { numbering: { ...n, nums }, numId: id };
}
