// Step 4 document operations: foot- and endnotes, the table of contents (Inhaltsverzeichnis)
// and comments. Like edits.ts they replace paragraph objects and work through Doc.

import { Doc, MAIN, type Pos } from "./doc";
import { EMPTY, OBJ, intern, makePara, type Anchor, type CharProps, type CommentMeta, type InlineObject, type Para } from "./model";

// ----- notes ---------------------------------------------------------------------------------

/** A new foot- or endnote at the caret: the reference in the text, the note with its mark. */
export function insertNote(doc: Doc, at: Pos, kind: "fn" | "en", refStyle: string, textStyle: string): { id: string; ref: Pos } {
  let n = 0;
  for (const k of doc.stories.keys()) if (k.startsWith(`${kind}:`)) n = Math.max(n, Number(k.slice(3)) || 0);
  const id = `${kind}:${n + 1}`;
  const refProps = intern({ rStyle: refStyle });
  const mark: InlineObject = { kind: "note", note: id, display: "", data: { inner: kind === "fn" ? "<w:footnoteRef/>" : "<w:endnoteRef/>" } };
  const note: Para = {
    text: `${OBJ} `, runs: [{ len: 1, props: refProps }, { len: 1, props: EMPTY }], props: { style: textStyle }, mark: EMPTY, objects: [mark],
  };
  doc.setStory(id, [note]);
  const ref: InlineObject = { kind: "note", note: id, display: "", data: undefined };
  const piece: Para = { text: OBJ, runs: [{ len: 1, props: refProps }], props: {}, mark: EMPTY, objects: [ref] };
  const end = doc.replace(at, at, [piece]);
  return { id, ref: end };
}

/** Note ids referenced in the main text, in order. */
export function referencedNotes(doc: Doc): string[] {
  const out: string[] = [];
  for (const p of doc.main) for (const o of p.objects ?? []) if (o.note && !out.includes(o.note)) out.push(o.note);
  return out;
}

/** Position of the reference of note `id` in the main text. */
export function noteReference(doc: Doc, id: string): Pos | null {
  const main = doc.main;
  for (let p = 0; p < main.length; p++) {
    let k = 0;
    const para = main[p];
    for (let i = 0; i < para.text.length; i++) {
      if (para.text[i] !== OBJ) continue;
      if (para.objects?.[k++]?.note === id) return { p, o: i };
    }
  }
  return null;
}

// ----- table of contents -------------------------------------------------------------------

export interface TocEntry {
  p: number;
  level: number;
  text: string;
  /** The heading's number ("2.1") when headings are numbered. */
  label?: string;
}

let tocSeq = 1;
let bookmarkSeq = 900000;

/** Adds a _Toc bookmark around a heading (or finds its existing one); returns its name. */
export function tocBookmark(doc: Doc, p: number): string {
  const para = doc.main[p];
  const existing = para.anchors?.find((a) => a.bookmark?.startsWith("_Toc"));
  if (existing?.bookmark) return existing.bookmark;
  const id = bookmarkSeq++;
  const name = `_Toc${String(Date.now() % 1e9).padStart(9, "0")}${String(id % 1000).padStart(3, "0")}`;
  const anchors: Anchor[] = [
    ...(para.anchors ?? []),
    { at: 0, data: { raw: `<w:bookmarkStart w:id="${id}" w:name="${name}"/>` }, bookmark: name },
    { at: para.text.length, data: { raw: `<w:bookmarkEnd w:id="${id}"/>` } },
  ];
  const next = [...doc.main];
  next[p] = { ...para, anchors };
  doc.setStory(MAIN, next);
  return name;
}

/**
 * The entries of a table of contents as paragraphs: one per heading, linked to it (Strg+Klick
 * in Word), with a dotted tab to the page number. The paragraphs carry the field's start and
 * end markers (TOC \o "1-3" \h \z \u), so Word sees an ordinary table of contents.
 */
/** How entries look (Abbildungsverzeichnis dialog): page numbers, right-aligned with a leader, links. */
export interface TocLook {
  pageNumbers?: boolean;
  rightAlign?: boolean;
  leader?: "dot" | "hyphen" | "underscore" | "none";
  links?: boolean;
}

export function tocParagraphs(doc: Doc, entries: TocEntry[], pageOf: (p: number) => string, styles: string[], width: number, field?: string,
  instr = 'TOC \\o "1-3" \\h \\z \\u', look: TocLook = {}): Para[] {
  const numbers = look.pageNumbers !== false, right = look.rightAlign !== false, links = look.links !== false;
  const leader = look.leader ?? "dot";
  const id = field ?? `toc${tocSeq++}`;
  const out: Para[] = [];
  const r = (inner: string) => `<w:r>${inner}</w:r>`;
  const start = r('<w:fldChar w:fldCharType="begin"/>') + r(`<w:instrText xml:space="preserve"> ${instr} </w:instrText>`) + r('<w:fldChar w:fldCharType="separate"/>');
  if (!entries.length) {
    // Word's message when there are no headings.
    const msg = /\\c\b/.test(instr) ? "Es wurden keine Einträge für ein Abbildungsverzeichnis gefunden." : "Keine Einträge für das Inhaltsverzeichnis gefunden.";
    const p: Para = { ...makePara(msg), field: id,
      anchors: [{ at: 0, data: { raw: start }, field: { id, edge: "start", instr } }, { at: msg.length, data: { raw: r('<w:fldChar w:fldCharType="end"/>') }, field: { id, edge: "end" } }] };
    return [p];
  }
  entries.forEach((e, i) => {
    const bm = tocBookmark(doc, e.p);
    // Numbered headings: the number, a tab, the text (Word: left tab at 22 pt per level).
    const head = e.label ? `${e.label}\t${e.text}` : e.text;
    const text = numbers ? `${head}${right ? "\t" : " "}${OBJ}` : head;
    const numTab = e.label ? [{ pos: 22 * Math.min(3, e.level), align: "left" as const }] : [];
    const page: InlineObject = { kind: "field", data: undefined, display: pageOf(e.p), instr: `PAGEREF ${bm} \\h` };
    const linkId = `${id}-${i}`;
    // Strg+Klick on an entry goes to the heading (the link carries its bookmark).
    const anchors: Anchor[] = links ? [
      { at: 0, data: { raw: `<w:hyperlink w:anchor="${bm}" w:history="1">` }, link: { id: linkId, edge: "start", anchor: bm } },
      { at: text.length, data: { raw: "</w:hyperlink>" }, link: { id: linkId, edge: "end" } },
    ] : [];
    if (i === 0) anchors.unshift({ at: 0, data: { raw: start }, field: { id, edge: "start", instr } });
    if (i === entries.length - 1) anchors.push({ at: text.length, data: { raw: r('<w:fldChar w:fldCharType="end"/>') }, field: { id, edge: "end" } });
    out.push({
      text, runs: [{ len: text.length, props: EMPTY }], mark: EMPTY, objects: numbers ? [page] : [], anchors, field: id,
      props: { style: styles[Math.min(styles.length, e.level) - 1], tabs: [...numTab, ...(numbers && right ? [{ pos: width, align: "right" as const, leader }] : [])] },
    });
  });
  return out;
}

/**
 * The table of contents fields of the main text: id and the range of entry paragraphs.
 * Word often ends the field at the start of the paragraph after the entries (the first
 * heading); that paragraph is `tail` and not part of the table.
 */
export function tocFields(doc: Doc, kind: RegExp = /^\s*TOC\b/i): { id: string; a: number; b: number; tail?: number; instr: string }[] {
  const out: { id: string; a: number; b: number; tail?: number; instr: string }[] = [];
  doc.main.forEach((p, i) => {
    if (!p.field) return;
    let instr: string | undefined;
    for (const q of doc.main) for (const a of q.anchors ?? []) if (a.field?.id === p.field && a.field.edge === "start" && kind.test(a.field.instr ?? "")) instr = a.field.instr;
    if (instr === undefined) return;
    const cur = out.find((x) => x.id === p.field);
    if (cur) cur.b = i; else out.push({ id: p.field, a: i, b: i, instr });
  });
  for (const f of out) {
    const last = doc.main[f.b];
    const end = last.anchors?.find((x) => x.field?.id === f.id && x.field.edge === "end");
    if (f.b > f.a && end && end.at === 0 && last.text.length > 0) { f.tail = f.b; f.b--; }
  }
  return out;
}

/**
 * The paragraphs of a Literaturverzeichnis (BIBLIOGRAPHY field): one per source, formatted,
 * with the hanging indent Word uses; "Keine Quellen" when the document cites none.
 */
export function bibliographyParagraphs(entries: { text: string; italic?: boolean }[][], style: string, field?: string): Para[] {
  const id = field ?? `bib${tocSeq++}`;
  const instr = " BIBLIOGRAPHY ";
  const r = (inner: string) => `<w:r>${inner}</w:r>`;
  const start = r('<w:fldChar w:fldCharType="begin"/>') + r(`<w:instrText xml:space="preserve">${instr}</w:instrText>`) + r('<w:fldChar w:fldCharType="separate"/>');
  const list = entries.length ? entries : [[{ text: "Im aktuellen Dokument sind keine Quellen vorhanden." }]];
  return list.map((pieces, i) => {
    const text = pieces.map((x) => x.text).join("");
    const runs = pieces.filter((x) => x.text).map((x) => ({ len: x.text.length, props: x.italic ? intern({ italic: true }) : EMPTY }));
    const anchors: Anchor[] = [];
    if (i === 0) anchors.push({ at: 0, data: { raw: start }, field: { id, edge: "start", instr } });
    if (i === list.length - 1) anchors.push({ at: text.length, data: { raw: r('<w:fldChar w:fldCharType="end"/>') }, field: { id, edge: "end" } });
    return { text, runs: runs.length ? runs : [{ len: 0, props: EMPTY }], mark: EMPTY, anchors, field: id,
      props: { style, ...(entries.length ? { indLeft: 36, indFirst: -36 } : {}) } };
  });
}

/** Removes a field's markers from the paragraph after a table of contents (see tocFields). */
export function stripTocTail(doc: Doc, f: { id: string; tail?: number }): void {
  if (f.tail === undefined) return;
  const p = doc.main[f.tail];
  const main = [...doc.main];
  main[f.tail] = { ...p, field: p.field === f.id ? undefined : p.field, anchors: p.anchors?.filter((x) => x.field?.id !== f.id) };
  doc.setStory(MAIN, main);
}

// ----- comments -----------------------------------------------------------------------------

export function withAnchor(p: Para, a: Anchor): Para {
  return { ...p, anchors: [...(p.anchors ?? []), a].sort((x, y) => x.at - y.at) };
}

/** A new comment on [from, to): range markers in the text, the comment story and its data. */
export function addComment(doc: Doc, from: Pos, to: Pos, meta: Omit<CommentMeta, "id">, textStyle: string, refStyle: string, text = ""): string {
  let n = -1;
  for (const k of doc.stories.keys()) if (k.startsWith("cmt:")) n = Math.max(n, Number(k.slice(4)) || 0);
  const id = `cmt:${n + 1}`;
  const main = [...doc.main];
  main[from.p] = withAnchor(main[from.p], { at: from.o, data: undefined, comment: { id, edge: "start" } });
  main[to.p] = withAnchor(main[to.p], { at: to.o, data: undefined, comment: { id, edge: "end" } });
  main[to.p] = withAnchor(main[to.p], { at: to.o, data: undefined, comment: { id, edge: "ref" } });
  doc.setStory(MAIN, main);
  doc.setStory(id, commentParas(text, textStyle, refStyle));
  doc.setComment({ ...meta, id }, id);
  return id;
}

function commentParas(text: string, textStyle: string, refStyle: string): Para[] {
  const lines = text.split("\n");
  return lines.map((line, i) => {
    const p = makePara(line, { style: textStyle });
    // The first paragraph starts with the comment mark Word expects.
    return i === 0 ? { ...p, anchors: [{ at: 0, data: { raw: `<w:r><w:rPr><w:rStyle w:val="${refStyle}"/></w:rPr><w:annotationRef/></w:r>` } }] } : p;
  });
}

/** Replaces a comment's text (plain text; lines become paragraphs). */
export function setCommentText(doc: Doc, id: string, text: string, textStyle: string, refStyle: string): void {
  const old = doc.story(id);
  const keepFirst = old?.[0]?.anchors;
  const paras = commentParas(text, textStyle, refStyle);
  if (keepFirst?.length) paras[0] = { ...paras[0], anchors: keepFirst.map((a) => ({ ...a, at: 0 })) };
  doc.setStory(id, paras);
}

/** The plain text of a comment. */
export function commentText(doc: Doc, id: string): string {
  return (doc.story(id) ?? []).map((p) => p.text.replace(/\ufffc/g, "")).join("\n");
}

/** Deletes a comment and its replies. */
export function deleteComment(doc: Doc, id: string): void {
  const ids = new Set([id]);
  for (const [k, m] of doc.comments) if (m.parent && ids.has(m.parent)) ids.add(k);
  const main = doc.main.map((p) => (p.anchors?.some((a) => a.comment && ids.has(a.comment.id))
    ? { ...p, anchors: p.anchors.filter((a) => !(a.comment && ids.has(a.comment.id))) } : p));
  doc.setStory(MAIN, main);
  for (const k of ids) { doc.deleteStory(k); doc.setComment(null, k); }
}

/** Where comments sit in the main text: id, start and end (main-text positions), in order. */
export function commentRanges(doc: Doc): { id: string; start: Pos; end: Pos }[] {
  const starts = new Map<string, Pos>(), ends = new Map<string, Pos>();
  doc.main.forEach((p, i) => {
    for (const a of p.anchors ?? []) {
      if (!a.comment) continue;
      if (a.comment.edge === "start") starts.set(a.comment.id, { p: i, o: a.at });
      else if (a.comment.edge === "end") ends.set(a.comment.id, { p: i, o: a.at });
      else if (!ends.has(a.comment.id)) ends.set(a.comment.id, { p: i, o: a.at });
    }
  });
  const out: { id: string; start: Pos; end: Pos }[] = [];
  for (const [id, end] of ends) out.push({ id, start: starts.get(id) ?? end, end });
  return out.sort((a, b) => a.start.p - b.start.p || a.start.o - b.start.o);
}

/** A reply to comment `parent`. */
export function addReply(doc: Doc, parent: string, meta: Omit<CommentMeta, "id" | "parent">, text: string, textStyle: string, refStyle: string): string {
  const r = commentRanges(doc).find((x) => x.id === parent);
  const at = r?.end ?? { p: 0, o: 0 };
  const id = addComment(doc, r?.start ?? at, at, meta, textStyle, refStyle, text);
  const m = doc.comments.get(id)!;
  doc.setComment({ ...m, parent }, id);
  return id;
}

export type { CharProps };
