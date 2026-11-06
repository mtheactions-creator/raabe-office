// Editing a text body (a:txBody) in place (D-070): positions are (paragraph, offset) with a line
// break counting one character and a field its text; inserting continues the run before the caret
// (PowerPoint takes the formatting of the character before), formatting splits runs at the range's
// ends and keeps every other attribute of rPr and pPr as it was.

import { PptxDom } from "@raabe/ooxml";

const { kid, kids, make, NS_A } = PptxDom;

export interface Pos { p: number; o: number }

export const cmpPos = (a: Pos, b: Pos) => (a.p !== b.p ? a.p - b.p : a.o - b.o);
export const minPos = (a: Pos, b: Pos) => (cmpPos(a, b) <= 0 ? a : b);
export const maxPos = (a: Pos, b: Pos) => (cmpPos(a, b) <= 0 ? b : a);

interface Seg { el: Element; start: number; len: number; kind: "r" | "br" | "fld" }

export function paras(body: Element): Element[] {
  return kids(body, "p");
}

export function segments(p: Element): Seg[] {
  const out: Seg[] = [];
  let at = 0;
  for (const c of kids(p)) {
    if (c.localName === "r" || c.localName === "fld") {
      const len = (kid(c, "t")?.textContent ?? "").length;
      out.push({ el: c, start: at, len, kind: c.localName });
      at += len;
    } else if (c.localName === "br") {
      out.push({ el: c, start: at, len: 1, kind: "br" });
      at += 1;
    }
  }
  return out;
}

export function paraText(p: Element): string {
  return kids(p).map((c) => (c.localName === "br" ? "\n" : c.localName === "r" || c.localName === "fld" ? kid(c, "t")?.textContent ?? "" : "")).join("");
}

export function bodyText(body: Element): string {
  return paras(body).map(paraText).join("\r");
}

export function paraLength(p: Element): number {
  return paraText(p).length;
}

/** Makes sure the body has at least one paragraph. */
export function ensurePara(body: Element): Element {
  let p = paras(body)[0];
  if (!p) {
    p = make(body.ownerDocument, "a:p", {}, make(body.ownerDocument, "a:endParaRPr", { lang: "de-DE", dirty: "0" }));
    body.append(p);
  }
  return p;
}

function newRun(doc: Document, text: string, rPr?: Element | null): Element {
  const r = make(doc, "a:r");
  const pr = rPr ? (rPr.cloneNode(true) as Element) : make(doc, "a:rPr", { lang: "de-DE" });
  if (pr.localName !== "rPr") {
    // From an endParaRPr: the same attributes and children as a run's rPr.
    const x = make(doc, "a:rPr");
    for (const a of Array.from(pr.attributes)) x.setAttribute(a.name, a.value);
    for (const c of Array.from(pr.childNodes)) x.append(c.cloneNode(true));
    r.append(x);
  } else r.append(pr);
  pr.removeAttribute?.("smtClean");
  r.append(make(doc, "a:t", {}, text));
  return r;
}

/** Splits the run containing offset `o` so that a run boundary lies at `o`; returns the element after it (or null at the end). */
function splitAt(p: Element, o: number): Element | null {
  for (const s of segments(p)) {
    if (o <= s.start) return s.el;
    if (o < s.start + s.len) {
      if (s.kind !== "r") return s.el;
      const t = kid(s.el, "t")!;
      const text = t.textContent ?? "";
      const k = o - s.start;
      const right = s.el.cloneNode(true) as Element;
      t.textContent = text.slice(0, k);
      kid(right, "t")!.textContent = text.slice(k);
      s.el.after(right);
      return right;
    }
  }
  return kid(p, "endParaRPr") ?? null;
}

/** The run (or end mark) whose formatting new text at `o` takes: the character before, else after. */
export function formatAt(p: Element, o: number): Element | null {
  const segs = segments(p);
  let prev: Seg | undefined;
  for (const s of segs) {
    if (s.start < o && (s.kind === "r" || s.kind === "fld")) prev = s;
    if (s.start >= o) break;
  }
  if (prev) return kid(prev.el, "rPr") ?? null;
  const next = segs.find((s) => s.kind === "r");
  if (next) return kid(next.el, "rPr") ?? null;
  return kid(p, "endParaRPr") ?? null;
}

/** Inserts text (no line breaks) at `pos`; `rPr` overrides the formatting taken from the text before. */
export function insertText(body: Element, pos: Pos, text: string, rPr?: Element | null): Pos {
  const p = paras(body)[pos.p] ?? ensurePara(body);
  const segs = segments(p);
  const fmt = rPr ?? formatAt(p, pos.o);
  // Continue the run before the caret when it has the formatting wanted.
  const before = segs.find((s) => s.kind === "r" && s.start < pos.o && pos.o <= s.start + s.len);
  if (before && !rPr) {
    const t = kid(before.el, "t")!;
    const v = t.textContent ?? "";
    const k = pos.o - before.start;
    t.textContent = v.slice(0, k) + text + v.slice(k);
    clearDirty(before.el);
    return { p: pos.p, o: pos.o + text.length };
  }
  const after = splitAt(p, pos.o);
  const run = newRun(body.ownerDocument, text, fmt);
  p.insertBefore(run, after && after.parentElement === p ? after : null);
  if (!kid(p, "endParaRPr")) p.append(make(body.ownerDocument, "a:endParaRPr", { lang: "de-DE", dirty: "0" }));
  else p.append(kid(p, "endParaRPr")!);
  return { p: pos.p, o: pos.o + text.length };
}

function clearDirty(r: Element): void {
  kid(r, "rPr")?.removeAttribute("smtClean");
}

/** Umschalt+Eingabe: a line break inside the paragraph. */
export function insertBreak(body: Element, pos: Pos): Pos {
  const p = paras(body)[pos.p];
  const fmt = formatAt(p, pos.o);
  const after = splitAt(p, pos.o);
  const br = make(body.ownerDocument, "a:br");
  if (fmt) br.append(rprCopy(body.ownerDocument, fmt));
  p.insertBefore(br, after && after.parentElement === p ? after : kid(p, "endParaRPr") ?? null);
  return { p: pos.p, o: pos.o + 1 };
}

function rprCopy(doc: Document, from: Element): Element {
  const x = make(doc, "a:rPr");
  for (const a of Array.from(from.attributes)) if (a.name !== "smtClean") x.setAttribute(a.name, a.value);
  for (const c of Array.from(from.childNodes)) x.append(c.cloneNode(true));
  return x;
}

/** Eingabe: splits the paragraph; the new one takes the paragraph's properties (level, alignment). */
export function splitPara(body: Element, pos: Pos): Pos {
  const ps = paras(body);
  const p = ps[pos.p];
  const after = splitAt(p, pos.o);
  const np = make(body.ownerDocument, "a:p");
  const pPr = kid(p, "pPr");
  if (pPr) np.append(pPr.cloneNode(true));
  // Runs from the split point move; the old paragraph keeps a mark formatted like its last text.
  const moving: Element[] = [];
  let cur: Element | null = after && after.parentElement === p ? after : null;
  while (cur && cur.localName !== "endParaRPr") { moving.push(cur); cur = cur.nextElementSibling; }
  for (const m of moving) np.append(m);
  const end = kid(p, "endParaRPr");
  const fmt = formatAt(p, pos.o) ?? end;
  if (end) np.append(end.cloneNode(true));
  else np.append(make(body.ownerDocument, "a:endParaRPr", { lang: "de-DE", dirty: "0" }));
  if (fmt && fmt !== end) {
    const e = make(body.ownerDocument, "a:endParaRPr");
    for (const a of Array.from(fmt.attributes)) if (a.name !== "smtClean") e.setAttribute(a.name, a.value);
    for (const c of Array.from(fmt.childNodes)) e.append(c.cloneNode(true));
    end?.remove();
    p.append(e);
  }
  p.after(np);
  return { p: pos.p + 1, o: 0 };
}

/** Deletes the text from `a` to `b`; paragraphs in between are joined into the first. */
export function deleteRange(body: Element, a: Pos, b: Pos): Pos {
  const from = minPos(a, b), to = maxPos(a, b);
  if (cmpPos(from, to) === 0) return from;
  const ps = paras(body);
  if (from.p === to.p) {
    cutPara(ps[from.p], from.o, to.o);
    return from;
  }
  const first = ps[from.p], last = ps[to.p];
  cutPara(first, from.o, paraLength(first));
  cutPara(last, 0, to.o);
  const end = kid(first, "endParaRPr");
  for (const c of kids(last)) if (c.localName !== "pPr" && c.localName !== "endParaRPr") first.insertBefore(c, end ?? null);
  for (let i = from.p + 1; i <= to.p; i++) ps[i].remove();
  return from;
}

function cutPara(p: Element, a: number, b: number): void {
  if (a >= b) return;
  const startEl = splitAt(p, a);
  splitAt(p, b);
  // Elements whose start lies in [a, b) go.
  for (const s of segments(p)) if (s.start >= a && s.start + s.len <= b && s.len > 0) {
    // Keep the formatting of an emptied paragraph in its mark (PowerPoint keeps typing in it).
    if (segments(p).length === 1 && s.kind === "r") {
      const end = kid(p, "endParaRPr");
      const rPr = kid(s.el, "rPr");
      if (end && rPr) for (const at2 of Array.from(rPr.attributes)) if (at2.name !== "smtClean" && !end.hasAttribute(at2.name)) end.setAttribute(at2.name, at2.value);
    }
    s.el.remove();
  }
  void startEl;
}

/** The text between two positions (paragraphs joined by "\r", PowerPoint's clipboard text uses CRLF). */
export function rangeText(body: Element, a: Pos, b: Pos): string {
  const from = minPos(a, b), to = maxPos(a, b);
  const ps = paras(body).map(paraText);
  if (from.p === to.p) return ps[from.p]?.slice(from.o, to.o) ?? "";
  const parts = [ps[from.p].slice(from.o), ...ps.slice(from.p + 1, to.p), ps[to.p].slice(0, to.o)];
  return parts.join("\r\n");
}

/** Applies `fn` to the rPr of every run in the range (split at its ends). An empty range changes nothing. */
export function formatRange(body: Element, a: Pos, b: Pos, fn: (rPr: Element) => void, withEnds = false): void {
  const from = minPos(a, b), to = maxPos(a, b);
  const ps = paras(body);
  for (let i = from.p; i <= to.p; i++) {
    const p = ps[i];
    if (!p) continue;
    const s = i === from.p ? from.o : 0, e = i === to.p ? to.o : paraLength(p);
    if (s < e) {
      splitAt(p, s);
      splitAt(p, e);
      for (const seg of segments(p)) if (seg.start >= s && seg.start + seg.len <= e && seg.kind !== "br") fn(ensureRPr(seg.el));
    }
    // The paragraph mark goes with the text when the range reaches the paragraph's end.
    if (withEnds || (e === paraLength(p) && (s < e || i < to.p))) fn(ensureEnd(p));
  }
}

/** Formatting of a whole text body (a selected frame): every run and paragraph mark. */
export function formatAll(body: Element, fn: (rPr: Element) => void): void {
  for (const p of paras(body)) {
    for (const c of kids(p)) if (c.localName === "r" || c.localName === "fld" || c.localName === "br") fn(ensureRPr(c));
    fn(ensureEnd(p));
  }
}

function ensureRPr(r: Element): Element {
  let pr = kid(r, "rPr");
  if (!pr) { pr = make(r.ownerDocument, "a:rPr", { lang: "de-DE" }); r.insertBefore(pr, r.firstChild); }
  pr.removeAttribute("smtClean");
  return pr;
}

function ensureEnd(p: Element): Element {
  let e = kid(p, "endParaRPr");
  if (!e) { e = make(p.ownerDocument, "a:endParaRPr", { lang: "de-DE", dirty: "0" }); p.append(e); }
  return e;
}

/** The pPr of paragraphs from..to, created when missing (placed first in the paragraph). */
export function paraProps(body: Element, fromP: number, toP: number, fn: (pPr: Element, p: Element, i: number) => void): void {
  const ps = paras(body);
  for (let i = Math.min(fromP, toP); i <= Math.max(fromP, toP); i++) {
    const p = ps[i];
    if (!p) continue;
    let pr = kid(p, "pPr");
    if (!pr) { pr = make(p.ownerDocument, "a:pPr"); p.insertBefore(pr, p.firstChild); }
    fn(pr, p, i);
    if (!pr.attributes.length && !pr.children.length) pr.remove();
  }
}

/** Sets or removes (null) an attribute. */
export function setAttr(el: Element, name: string, v: string | null): void {
  if (v === null) el.removeAttribute(name);
  else el.setAttribute(name, v);
}

/** Replaces the fill (a:solidFill …) of an rPr with a colour (hex or scheme name), or removes it. */
export function setRunColor(rPr: Element, color: string | null): void {
  for (const c of kids(rPr)) if (["solidFill", "noFill", "gradFill", "pattFill", "blipFill", "grpFill"].includes(c.localName)) c.remove();
  if (color === null) return;
  const doc = rPr.ownerDocument;
  const clr = /^[0-9A-F]{6}$/i.test(color) ? make(doc, "a:srgbClr", { val: color.toUpperCase() }) : make(doc, "a:schemeClr", { val: color });
  const fill = make(doc, "a:solidFill", {}, clr);
  // Schema order: ln, fills, effects, highlight, uLnTx…, latin…
  const before = kids(rPr).find((c) => !["ln"].includes(c.localName));
  rPr.insertBefore(fill, before ?? null);
}

/** Sets the latin typeface of an rPr (with ea and cs left alone). */
export function setRunFont(rPr: Element, face: string): void {
  let latin = kid(rPr, "latin");
  if (!latin) {
    latin = make(rPr.ownerDocument, "a:latin");
    const order = ["ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"];
    const before = kids(rPr).find((c) => order.includes(c.localName));
    rPr.insertBefore(latin, before ?? null);
  }
  latin.setAttribute("typeface", face);
  for (const n of ["panose", "pitchFamily", "charset"]) latin.removeAttribute(n);
}

export function setHighlight(rPr: Element, hex: string | null): void {
  kid(rPr, "highlight")?.remove();
  if (!hex) return;
  const doc = rPr.ownerDocument;
  const h = make(doc, "a:highlight", {}, make(doc, "a:srgbClr", { val: hex }));
  const order = ["uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"];
  rPr.insertBefore(h, kids(rPr).find((c) => order.includes(c.localName)) ?? null);
}

const WORD = /[\p{L}\p{N}_]/u;

/** The word around `o` (Doppelklick; formatting at a caret applies to it in PowerPoint). */
export function wordAt(p: Element, o: number): { from: number; to: number } {
  const t = paraText(p);
  let a = o, b = o;
  while (a > 0 && WORD.test(t[a - 1])) a--;
  while (b < t.length && WORD.test(t[b])) b++;
  return { from: a, to: b };
}

/** The word to select on a double-click: the word with its trailing spaces (PowerPoint), or the punctuation. */
export function wordSelection(p: Element, o: number): { from: number; to: number } {
  const t = paraText(p);
  if (o < t.length && !WORD.test(t[o]) && !(o > 0 && WORD.test(t[o - 1]))) {
    if (t[o] === " ") { let a = o, b = o; while (a > 0 && t[a - 1] === " ") a--; while (b < t.length && t[b] === " ") b++; return { from: a, to: b }; }
    return { from: o, to: o + 1 };
  }
  const w = wordAt(p, o);
  let b = w.to;
  while (b < t.length && t[b] === " ") b++;
  return { from: w.from, to: b };
}

/** Strg+Links / Strg+Rechts: the start of the previous / next word. */
export function wordJump(body: Element, pos: Pos, dir: -1 | 1): Pos {
  const ps = paras(body);
  const t = paraText(ps[pos.p]);
  let o = pos.o;
  if (dir > 0) {
    if (o >= t.length) return pos.p + 1 < ps.length ? { p: pos.p + 1, o: 0 } : pos;
    while (o < t.length && WORD.test(t[o])) o++;
    if (o === pos.o) o++;
    while (o < t.length && t[o] === " ") o++;
    return { p: pos.p, o };
  }
  if (o === 0) return pos.p > 0 ? { p: pos.p - 1, o: paraText(ps[pos.p - 1]).length } : pos;
  while (o > 0 && t[o - 1] === " ") o--;
  if (o > 0 && WORD.test(t[o - 1])) while (o > 0 && WORD.test(t[o - 1])) o--;
  else o--;
  return { p: pos.p, o };
}

export function endPos(body: Element): Pos {
  const ps = paras(body);
  return { p: Math.max(0, ps.length - 1), o: ps.length ? paraLength(ps[ps.length - 1]) : 0 };
}

/** Replaces the whole text by `text` (typing on a selected frame), keeping the first run's formatting. */
export function replaceAll(body: Element, text: string): Pos {
  const ps = paras(body);
  const first = ps[0] ?? ensurePara(body);
  const fmt = formatAt(first, 1) ?? kid(first, "endParaRPr");
  const keep = fmt ? fmt.cloneNode(true) as Element : null;
  for (let i = 1; i < ps.length; i++) ps[i].remove();
  for (const c of kids(first)) if (c.localName !== "pPr" && c.localName !== "endParaRPr") c.remove();
  if (!text) return { p: 0, o: 0 };
  return insertText(body, { p: 0, o: 0 }, text, keep);
}

export { NS_A };
