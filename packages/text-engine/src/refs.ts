// Textmarken, Querverweise, Beschriftungen and the fields they use, as Word writes them:
// bookmarks are w:bookmarkStart/End pairs (hidden ones start with "_": _Ref for cross
// references, _Toc for the table of contents); a cross reference is a REF, PAGEREF or NOTEREF
// field to a bookmark; a caption numbers itself with a SEQ field. Field results are computed
// here (Felder aktualisieren, F9).

import { MAIN, type Doc, type Pos } from "./doc";
import { OBJ, type Anchor, type InlineObject, type Para } from "./model";
import { formatNumber } from "./numbering";

export interface Bookmark {
  readonly name: string;
  readonly start: Pos;
  readonly end: Pos;
}

const idOf = (a: Anchor): string | null => {
  const raw = (a.data as { raw?: string } | undefined)?.raw ?? "";
  return /<w:bookmark(?:Start|End)\b[^>]*\bw:id="(\d+)"/.exec(raw)?.[1] ?? null;
};
const isEnd = (a: Anchor) => /<w:bookmarkEnd\b/.test((a.data as { raw?: string } | undefined)?.raw ?? "");

/** Every bookmark of the main text, in document order. */
export function bookmarks(doc: Doc): Bookmark[] {
  const starts = new Map<string, { name: string; start: Pos }>();
  const out: Bookmark[] = [];
  doc.main.forEach((p, i) => {
    for (const a of p.anchors ?? []) {
      const id = idOf(a);
      if (a.bookmark && id) starts.set(id, { name: a.bookmark, start: { p: i, o: a.at } });
      else if (id && isEnd(a)) {
        const s = starts.get(id);
        if (s) { out.push({ name: s.name, start: s.start, end: { p: i, o: a.at } }); starts.delete(id); }
      }
    }
  });
  for (const s of starts.values()) out.push({ name: s.name, start: s.start, end: s.start });
  return out.sort((a, b) => a.start.p - b.start.p || a.start.o - b.start.o);
}

function nextId(doc: Doc): number {
  let n = 0;
  for (const paras of [doc.main, ...doc.stories.values()]) for (const p of paras) for (const a of p.anchors ?? []) {
    const id = idOf(a);
    if (id) n = Math.max(n, Number(id) + 1);
  }
  return n;
}

const withAnchors = (p: Para, more: Anchor[]): Para => ({ ...p, anchors: [...(p.anchors ?? []), ...more].sort((x, y) => x.at - y.at) });

/** Textmarke > Hinzufügen: a bookmark over [from, to) (one of the same name is moved). */
export function addBookmark(doc: Doc, name: string, from: Pos, to: Pos): void {
  deleteBookmark(doc, name);
  const id = nextId(doc);
  const main = [...doc.main];
  main[from.p] = withAnchors(main[from.p], [{ at: from.o, data: { raw: `<w:bookmarkStart w:id="${id}" w:name="${escAttr(name)}"/>` }, bookmark: name }]);
  main[to.p] = withAnchors(main[to.p], [{ at: to.o, data: { raw: `<w:bookmarkEnd w:id="${id}"/>` } }]);
  doc.setStory(MAIN, main);
}

/** Textmarke > Löschen. */
export function deleteBookmark(doc: Doc, name: string): void {
  let id: string | null = null;
  for (const p of doc.main) for (const a of p.anchors ?? []) if (a.bookmark === name) id = idOf(a);
  if (id === null) return;
  doc.setStory(MAIN, doc.main.map((p) => (p.anchors?.some((a) => idOf(a) === id)
    ? { ...p, anchors: p.anchors.filter((a) => idOf(a) !== id) } : p)));
}

const escAttr = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");

/** Word's bookmark names: a letter first, then letters, digits and _ (at most 40). */
export function validBookmarkName(name: string): boolean {
  return /^[\p{L}][\p{L}\p{N}_]{0,39}$/u.test(name);
}

let refSeq = 0;

/**
 * A hidden _Ref bookmark over [from, to) for a cross reference (or the one already there);
 * returns its name.
 */
export function refBookmark(doc: Doc, from: Pos, to: Pos): string {
  const existing = bookmarks(doc).find((b) => b.name.startsWith("_Ref") && b.start.p === from.p && b.start.o === from.o && b.end.p === to.p && b.end.o === to.o);
  if (existing) return existing.name;
  const name = `_Ref${String(Date.now() % 1e9).padStart(9, "0")}${String(refSeq++ % 1000).padStart(3, "0")}`;
  addBookmark(doc, name, from, to);
  return name;
}

// ----- captions ---------------------------------------------------------------------------------

/** The SEQ fields of a label (Abbildung, Tabelle, Formel) in document order. */
export function seqFields(doc: Doc, label: string): { p: number; o: number }[] {
  const re = new RegExp(`^\\s*SEQ\\s+${label.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`, "i");
  const out: { p: number; o: number }[] = [];
  doc.main.forEach((p, i) => {
    let k = 0;
    for (let o = 0; o < p.text.length; o++) {
      if (p.text[o] !== OBJ) continue;
      const obj = p.objects?.[k++];
      if (obj?.kind === "field" && re.test(obj.instr ?? "")) out.push({ p: i, o });
    }
  });
  return out;
}

/** The captions of a label: paragraph, number, the whole text and the text after "Abbildung 1: ". */
export function captions(doc: Doc, label: string): { p: number; number: string; text: string; rest: string }[] {
  return seqFields(doc, label).map(({ p, o }) => {
    const para = doc.main[p];
    const obj = objectAt(para, o);
    const after = para.text.slice(o + 1).replace(OBJ_RE, "");
    return { p, number: obj?.display ?? "", text: para.text.slice(0, o).replace(OBJ_RE, "") + (obj?.display ?? "") + after, rest: after.replace(/^[\s:.\-–—]+/, "") };
  });
}

const OBJ_RE = new RegExp(OBJ, "g");

function objectAt(p: Para, o: number): InlineObject | undefined {
  let k = 0;
  for (let i = 0; i < o; i++) if (p.text[i] === OBJ) k++;
  return p.objects?.[k];
}

// ----- field results (Felder aktualisieren) ---------------------------------------------------

/**
 * Recomputes the results of SEQ, REF, PAGEREF and NOTEREF fields in the main text: SEQ counts
 * per label, REF shows the bookmark's text (or "oben"/"unten" with \p), PAGEREF its page,
 * NOTEREF the note's number. Returns whether anything changed.
 */
export function updateFields(doc: Doc, pageOf: (p: number) => string, noteNumber: (story: string) => string | null, numberOf: (p: number) => string = () => ""): boolean {
  // Numbers first: a reference shows the caption's new number.
  const seq = updatePass(doc, pageOf, noteNumber, numberOf, true);
  return updatePass(doc, pageOf, noteNumber, numberOf, false) || seq;
}

function updatePass(doc: Doc, pageOf: (p: number) => string, noteNumber: (story: string) => string | null, numberOf: (p: number) => string, seqOnly: boolean): boolean {
  const marks = new Map(bookmarks(doc).map((b) => [b.name, b]));
  const counters = new Map<string, number>();
  let changed = false;
  const textOf = (b: Bookmark) => {
    if (b.start.p !== b.end.p) return doc.main[b.start.p].text.slice(b.start.o).replace(OBJ_RE, "");
    const para = doc.main[b.start.p];
    let out = "";
    let k = 0;
    for (let i = 0; i < b.end.o; i++) {
      const ch = para.text[i];
      const obj = ch === OBJ ? para.objects?.[k++] : undefined;
      if (i < b.start.o) continue;
      out += ch === OBJ ? obj?.display ?? "" : ch;
    }
    return out;
  };
  const main = doc.main.map((para, pi) => {
    if (!para.objects?.length) return para;
    let k = 0, here = 0;
    const objects = [...para.objects];
    for (let o = 0; o < para.text.length; o++) {
      if (para.text[o] !== OBJ) continue;
      const idx = k++;
      const obj = objects[idx];
      if (obj?.kind !== "field") continue;
      const instr = obj.instr ?? "";
      let display: string | null = null;
      let m: RegExpExecArray | null;
      if ((m = /^\s*SEQ\s+(\S+)/i.exec(instr))) {
        if (!seqOnly) continue;
        const n = (counters.get(m[1].toLowerCase()) ?? 0) + 1;
        counters.set(m[1].toLowerCase(), n);
        // Nummerierung: \* ARABIC, alphabetic, ALPHABETIC, roman, ROMAN.
        const fmt = /\\\*\s*(\w+)/.exec(instr)?.[1] ?? "ARABIC";
        display = fmt === "alphabetic" ? formatNumber(n, "lowerLetter") : fmt === "ALPHABETIC" ? formatNumber(n, "upperLetter")
          : fmt === "roman" ? formatNumber(n, "lowerRoman") : fmt === "ROMAN" ? formatNumber(n, "upperRoman") : String(n);
      } else if (!seqOnly && (m = /^\s*(PAGEREF|REF|NOTEREF)\s+(\S+)/i.exec(instr))) {
        const b = marks.get(m[2]);
        if (!b) continue;
        const kind = m[1].toUpperCase();
        const above = b.start.p < pi || (b.start.p === pi && b.start.o < o) ? "oben" : "unten";
        const p = /\\p\b/.test(instr);
        // \p: "oben"/"unten"; a page reference to another page says "auf Seite 3".
        if (p && kind === "PAGEREF") display = pageOf(b.start.p) === pageOf(pi) ? above : `auf Seite ${pageOf(b.start.p)}`;
        else if (p) display = above;
        else if (kind === "PAGEREF") display = pageOf(b.start.p);
        else if (kind === "NOTEREF") {
          const para2 = doc.main[b.start.p];
          const ref = objectAt(para2, b.start.o);
          display = ref?.note ? noteNumber(ref.note) : null;
        } else if (/\\r\b/.test(instr)) display = numberOf(b.start.p);
        else display = textOf(b);
      }
      if (display !== null && display !== obj.display) { objects[idx] = { ...obj, display, data: undefined }; changed = true; here++; }
    }
    return here ? { ...para, objects } : para;
  });
  if (changed) doc.setStory(MAIN, main);
  return changed;
}
