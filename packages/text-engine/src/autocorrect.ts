// AutoKorrektur and AutoFormat während der Eingabe, as Word 365 does them with its default
// settings (measured on 2026-10-14, D-042): typographic quotes for the text's language,
// capital letters at the start of sentences and table cells, ZWei GRoßbuchstaben, weekdays,
// accidental Caps Lock, replacements such as (c) → ©, ... → …, --> → Wingdings arrow,
// :) → emoji, fractions ½ ¼ ¾, en and em dashes. The abbreviation and typo lists are our own.
// Entries of Word's list that end in a symbol ((c), -->, ..., :)) are replaced as soon as
// their last character is typed; words are corrected when they end.
//
// Everything works on one paragraph's text and returns the edits to make; the editor makes
// them as their own undo step, so Strg+Z right after a correction brings the typed text back.

export interface Fix {
  /** Range in the paragraph to replace. */
  a: number;
  b: number;
  text: string;
  /** A symbol from a symbol font instead of text (Word's w:sym). */
  symbol?: { font: string; char: number };
  /** Font for the replacement text (emoji in Segoe UI Emoji). */
  font?: string;
}

/** Characters that end a word and trigger corrections. */
export const TRIGGERS = new Set([" ", "\t", "\n", ".", ",", ";", ":", "!", "?", ")", "]", "}", "\"", "'", "»", "“", "”", "‘", "’"]);

const LETTER = /\p{L}/u;
const UPPER = /\p{Lu}/u;
const LOWER = /\p{Ll}/u;

/** Abbreviations after which a period does not end a sentence (lower case, with the dot). */
const ABBREVIATIONS = new Set([
  // German
  "abb.", "abk.", "abs.", "abt.", "allg.", "anm.", "aufl.", "ausg.", "bd.", "bes.", "betr.", "bspw.", "bsp.", "bzgl.", "bzw.",
  "ca.", "chr.", "d.h.", "dgl.", "dr.", "dt.", "ebd.", "einschl.", "etc.", "evtl.", "f.", "ff.", "fa.", "fr.", "frz.", "geb.", "gegr.",
  "ggf.", "gr.", "hr.", "hrn.", "hrsg.", "i.a.", "inkl.", "jh.", "jhd.", "kap.", "kath.", "lat.", "max.", "min.", "mio.", "mrd.",
  "nr.", "o.ä.", "o.g.", "pkt.", "prof.", "s.", "sog.", "st.", "std.", "str.", "tel.", "u.a.", "u.ä.", "usw.", "v.a.", "vgl.",
  "z.b.", "z.t.", "zz.", "zzt.", "zit.", "jan.", "feb.", "apr.", "aug.", "sept.", "okt.", "nov.", "dez.",
  // English
  "e.g.", "i.e.", "mr.", "mrs.", "ms.", "vs.", "approx.", "dept.", "no.", "p.", "pp.",
  // French and Spanish
  "p.ex.", "env.", "mme.", "sr.", "sra.", "pág.",
]);

/** Common typing slips (our own list): wrong → right. Only whole words, case kept. */
const TYPOS: Record<string, string> = {
  dei: "die", udn: "und", nciht: "nicht", ncht: "nicht", nict: "nicht", ihc: "ich", isch: "sich", dsa: "das", dss: "dass",
  asl: "als", acuh: "auch", ahcu: "auch", wnen: "wenn", wrid: "wird", eien: "eine", eienn: "einen", sien: "sein",
  snid: "sind", sidn: "sind", ahben: "haben", hbaen: "haben", werdne: "werden", wedren: "werden", mti: "mit", mmit: "mit",
  vno: "von", fpr: "für", odre: "oder", abre: "aber", desen: "dessen", nru: "nur", shcon: "schon", scohn: "schon",
  nohc: "noch", imemr: "immer", immmer: "immer", jezt: "jetzt",
  teh: "the", adn: "and", taht: "that", wiht: "with", hte: "the", recieve: "receive", becuase: "because",
};

const WEEKDAYS: Record<string, string[]> = {
  de: ["montag", "dienstag", "mittwoch", "donnerstag", "freitag", "samstag", "sonnabend", "sonntag"],
  en: ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"],
};

/** Plain replacements of Word's AutoCorrect list (symbols), and its formatted arrows. */
const REPLACE: Record<string, string> = { "(c)": "©", "(r)": "®", "(tm)": "™", "(e)": "€", "...": "…" };
const SYMBOLS: Record<string, number> = { "-->": 0xf0e0, "<--": 0xf0df, "==>": 0xf0e8, "<==": 0xf0e7, "<=>": 0xf0f3 };
/** Word's smileys, now emoji in Segoe UI Emoji. */
const EMOJI: Record<string, string> = {
  ":)": "\u{1F60A}", ":-)": "\u{1F60A}", ";)": "\u{1F609}", ";-)": "\u{1F609}",
  ":(": "\u2639", ":-(": "\u2639", ":|": "\u{1F610}", ":-|": "\u{1F610}",
};
const INSTANT_KEYS = [...Object.keys(REPLACE), ...Object.keys(SYMBOLS), ...Object.keys(EMOJI)].sort((x, y) => y.length - x.length);
/** Last characters of the entries that are replaced at once. */
export const INSTANT_TRIGGERS = new Set(INSTANT_KEYS.map((k) => k[k.length - 1]));

/**
 * An entry of the AutoCorrect list completed by the character just typed (before `at`):
 * "(c)" → "©" even right after a word, "-->" → arrow, "..." → "…", ":)" → emoji (smileys
 * only after a space or at the start).
 */
export function instantFix(text: string, at: number): Fix | null {
  for (const key of INSTANT_KEYS) {
    const a = at - key.length;
    if (a < 0 || text.slice(a, at).toLowerCase() !== key) continue;
    if (EMOJI[key]) {
      if (a > 0 && !/\s/.test(text[a - 1])) continue;
      return { a, b: at, text: EMOJI[key], font: "Segoe UI Emoji" };
    }
    if (SYMBOLS[key] !== undefined) return { a, b: at, text: "", symbol: { font: "Wingdings", char: SYMBOLS[key] } };
    if (key === "..." && text[a - 1] === ".") continue;
    return { a, b: at, text: REPLACE[key] };
  }
  return null;
}
const FRACTIONS: Record<string, string> = { "1/2": "½", "1/4": "¼", "3/4": "¾" };

/** Opening and closing quotes by language: [double open, double close, single open, single close]. */
function quotes(lang: string): [string, string, string, string] {
  const l = lang.slice(0, 2).toLowerCase();
  if (l === "de") return ["„", "“", "‚", "‘"];
  if (l === "fr") return ["«\u00a0", "\u00a0»", "‘", "’"];
  if (l === "es") return ["“", "”", "‘", "’"];
  return ["“", "”", "‘", "’"];
}

/**
 * A typed straight quote (the character before `at`): opening after a space, a bracket or at
 * the start, otherwise closing; an apostrophe inside a word is the closing single quote.
 */
export function smartQuote(text: string, at: number, lang: string): Fix | null {
  const q = text[at - 1];
  if (q !== "\"" && q !== "'") return null;
  const before = text[at - 2];
  const opening = at < 2 || /[\s([{–—\u00a0„‚«]/.test(before ?? "");
  const [dO, dC, sO, sC] = quotes(lang);
  const rep = q === "\"" ? (opening ? dO : dC) : (opening ? sO : sC);
  return { a: at - 1, b: at, text: rep };
}

function sentenceStart(text: string, wordStart: number): boolean {
  let i = wordStart - 1;
  // Opening quotes and brackets before the word belong to it.
  while (i >= 0 && /["'(\[„‚“‘«]/.test(text[i])) i--;
  if (i < 0) return true;
  if (!/\s/.test(text[i])) return false;
  while (i >= 0 && /\s/.test(text[i])) i--;
  if (i < 0) return true;
  // Closing quotes and brackets after the sentence end.
  while (i >= 0 && /["'),\]“‘”’»]/.test(text[i]) && text[i] !== ",") i--;
  if (i < 0 || !/[.!?]/.test(text[i])) return false;
  if (text[i] === "." ) {
    if (text[i - 1] === ".") return false; // "..." continues the sentence
    let j = i - 1;
    while (j >= 0 && !/\s/.test(text[j])) j--;
    const token = text.slice(j + 1, i + 1).toLowerCase().replace(/^["'(\[]+/, "");
    if (ABBREVIATIONS.has(token)) return false;
    // Word ends the sentence after "1." too: "Am 1. januar" → "Am 1. Januar".
    if (/^\p{L}\.$/u.test(token)) return false; // single initials ("J. Smith")
  }
  return true;
}

/**
 * Corrections when a word has just been finished: `end` is where the trigger character
 * was typed (the word ends there). `cellStart`: the paragraph starts a table cell.
 */
export function wordFixes(text: string, end: number, lang: string, opts: { cellStart?: boolean; capsLock?: boolean } = {}): Fix[] {
  const fixes: Fix[] = [];
  let start = end;
  while (start > 0 && !/\s/.test(text[start - 1])) start--;
  const token = text.slice(start, end);
  if (!token) return fixes;
  const l = lang.slice(0, 2).toLowerCase();

  // Replacements from the AutoCorrect list (the symbol entries went when typed).
  if (FRACTIONS[token]) { fixes.push({ a: start, b: end, text: FRACTIONS[token] }); return fixes; }
  // An apostrophe inside a word: the German closing quote ‘ becomes ’ ("geht‘s" → "geht’s").
  for (let i = start + 1; i < end - 1; i++) {
    if (text[i] === "\u2018" && LETTER.test(text[i - 1]) && LETTER.test(text[i + 1])) fixes.push({ a: i, b: i + 1, text: "\u2019" });
  }

  // The word without surrounding punctuation.
  let ws = start, we = end;
  while (ws < we && !LETTER.test(text[ws]) && !/\d/.test(text[ws])) ws++;
  while (we > ws && !LETTER.test(text[we - 1]) && !/\d/.test(text[we - 1])) we--;
  let word = text.slice(ws, we);
  if (!word) return fixes;

  // Dashes: "Haus - Garten" → en dash when "Garten" is done, "Haus--Garten" → em dash.
  const em = /^(\p{L}+)--(\p{L}+)$/u.exec(token);
  if (em) { fixes.push({ a: start + em[1].length, b: start + em[1].length + 2, text: "—" }); }
  else if (text.slice(start - 3, start) === " - " && start >= 4 && !/\s/.test(text[start - 4])) fixes.push({ a: start - 2, b: start - 1, text: "–" });
  else if (text.slice(start - 4, start) === " -- " && start >= 5 && !/\s/.test(text[start - 5])) fixes.push({ a: start - 3, b: start - 1, text: "–" });

  if (!/^\p{L}+$/u.test(word)) return fixes;
  // List labels ("a)", "b.") keep their letter; AutoFormat turns them into a list.
  const listLabel = /^\(?\p{L}[.)\-]?$/u.test(token) && (/[.)\-]$/.test(token) || /[.)\-]/.test(text[end] ?? ""));
  const lower = word.toLowerCase();
  let fixed = word;
  // Typing slips.
  const typo = TYPOS[lower];
  if (typo && typo !== lower) fixed = UPPER.test(word[0]) ? typo[0].toUpperCase() + typo.slice(1) : typo;
  // ZWei GRoßbuchstaben am Wortanfang korrigieren.
  if (fixed.length >= 3 && UPPER.test(fixed[0]) && UPPER.test(fixed[1]) && LOWER.test(fixed[2]) && !/\p{Lu}/u.test(fixed.slice(2))) {
    fixed = fixed[0] + fixed[1].toLowerCase() + fixed.slice(2);
  }
  // Unbeabsichtigtes Verwenden der FESTSTELLTASTE korrigieren: "hALLO" → "Hallo".
  // Word only does this while Caps Lock is on.
  if (opts.capsLock && fixed.length >= 3 && LOWER.test(fixed[0]) && /^\p{Lu}+$/u.test(fixed.slice(1))) fixed = fixed[0].toUpperCase() + fixed.slice(1).toLowerCase();
  // Wochentage immer großschreiben.
  if ((WEEKDAYS[l] ?? []).includes(fixed)) fixed = fixed[0].toUpperCase() + fixed.slice(1);
  // Jeden Satz (und jede Tabellenzelle) mit einem Großbuchstaben beginnen.
  if (!listLabel && LOWER.test(fixed[0]) && (sentenceStart(text, start) || (opts.cellStart && text.slice(0, start).trim() === ""))) {
    fixed = fixed[0].toUpperCase() + fixed.slice(1);
  }
  if (fixed !== word) fixes.push({ a: ws, b: we, text: fixed });
  return fixes;
}

/** Border lines from "---", "___", "===", "***", "~~~", "###" typed alone and Enter (AutoFormat). */
export function borderLine(text: string): { style: string; width: number } | null {
  const m = /^(-{3,}|_{3,}|={3,}|\*{3,}|~{3,}|#{3,})$/.exec(text);
  if (!m) return null;
  const c = text[0];
  return c === "-" ? { style: "single", width: 0.75 } : c === "_" ? { style: "single", width: 1.5 } : c === "=" ? { style: "double", width: 0.75 }
    : c === "*" ? { style: "dotted", width: 1.5 } : c === "~" ? { style: "wave", width: 0.75 } : { style: "thinThickThinSmallGap", width: 2.25 };
}
