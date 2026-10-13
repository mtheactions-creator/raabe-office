// Word's notion of "words", used by double-click, Strg+arrow and Strg+Backspace.
// A token is a run of word characters or a run of punctuation, followed by any spaces.
// Spaces at the very start of a paragraph form a token of their own.

const SPACE = 0, WORD = 1, PUNCT = 2;
type CharClass = typeof SPACE | typeof WORD | typeof PUNCT;

const WORD_CHAR = /[\p{L}\p{N}\p{M}_]/u;

export function isWordChar(ch: string | undefined): boolean {
  return ch !== undefined && WORD_CHAR.test(ch);
}

function classify(text: string, i: number): CharClass {
  const ch = text[i];
  if (ch === " " || ch === " " || ch === "\t") return SPACE;
  if (WORD_CHAR.test(ch)) return WORD;
  // An apostrophe between two word characters belongs to the word ("geht's").
  if ((ch === "'" || ch === "’") && i > 0 && i < text.length - 1 &&
      WORD_CHAR.test(text[i - 1]) && WORD_CHAR.test(text[i + 1])) return WORD;
  return PUNCT;
}

function isTokenStart(text: string, i: number): boolean {
  if (i === 0) return true;
  const c = classify(text, i);
  if (c === SPACE) return false;
  return classify(text, i - 1) !== c;
}

/** Token containing the character at `i`, as [start, end) including trailing spaces. */
export function tokenAt(text: string, i: number): [number, number] {
  if (text.length === 0) return [0, 0];
  i = Math.min(i, text.length - 1);
  let start = i;
  while (start > 0 && !isTokenStart(text, start)) start--;
  let end = i + 1;
  while (end < text.length && !isTokenStart(text, end)) end++;
  return [start, end];
}

/** Strg+Rechts: start of the next token, or the paragraph end. */
export function nextTokenStart(text: string, i: number): number {
  for (let j = i + 1; j < text.length; j++) if (isTokenStart(text, j)) return j;
  return text.length;
}

/** Strg+Links: start of the current token if inside it, otherwise of the previous one. */
export function prevTokenStart(text: string, i: number): number {
  for (let j = i - 1; j > 0; j--) if (isTokenStart(text, j)) return j;
  return 0;
}

/** Strg+Klick: the sentence around `i`, including its trailing spaces. */
export function sentenceAt(text: string, i: number): [number, number] {
  const isEnd = (j: number) => /[.!?]/.test(text[j]) && (j + 1 >= text.length || text[j + 1] === " ");
  let start = Math.min(i, text.length);
  while (start > 0) {
    let k = start - 1;
    while (k >= 0 && text[k] === " ") k--;
    if (k >= 0 && k < start - 1 && isEnd(k)) break;
    start--;
  }
  let end = Math.min(i, text.length);
  while (end < text.length && !isEnd(end)) end++;
  if (end < text.length) end++;
  while (end < text.length && text[end] === " ") end++;
  return [start, end];
}

export function countWords(texts: readonly string[]): number {
  let n = 0;
  for (const t of texts) {
    let inWord = false;
    for (let i = 0; i < t.length; i++) {
      const space = t[i] === " " || t[i] === "\t" || t[i] === " ";
      if (!space && !inWord) n++;
      inWord = !space;
    }
  }
  return n;
}
