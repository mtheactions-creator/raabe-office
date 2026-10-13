// Änderungen nachverfolgen: with tracking on, deleted text stays (struck through) and new
// text is marked as inserted, both with author and time; Annehmen and Ablehnen resolve them.
// Deleting text one has inserted oneself removes it for real, as in Word.

import { Doc, cmp, type Pos } from "./doc";
import { intern, joinParas, normalizeRuns, slicePara, type CharProps, type Para, type Revision } from "./model";

let revSeq = 1000;

/** A new revision (one per typing burst or deletion). */
export function newRevision(kind: "ins" | "del", author: string): Revision {
  return { kind, author, date: new Date().toISOString().replace(/\.\d+Z$/, "Z"), id: revSeq++ };
}

/** Rebuilds a paragraph keeping only the character ranges `keep` (sorted, disjoint). */
function keepRanges(p: Para, keep: [number, number][]): Para {
  if (!keep.length) return { ...slicePara(p, 0, 0), sect: p.sect, mark: p.mark, props: p.props };
  let out = slicePara(p, keep[0][0], keep[0][1], true);
  for (let i = 1; i < keep.length; i++) out = joinParas(out, slicePara(p, keep[i][0], keep[i][1], true));
  return { ...out, props: p.props, mark: p.mark, sect: p.sect, cell: p.cell, group: p.group, field: p.field };
}

/**
 * Marks [from, to) as deleted by `rev`; text that `rev.author` inserted (still tracked) goes
 * for real. Paragraph marks inside the range are marked deleted. Returns the paragraphs'
 * new state through `doc`.
 */
export function markDeleted(doc: Doc, from: Pos, to: Pos, rev: Revision): void {
  if (cmp(from, to) >= 0) return;
  const paras = [...doc.paras];
  for (let i = from.p; i <= to.p; i++) {
    const p = paras[i];
    if (p.block) continue;
    const a = i === from.p ? from.o : 0, b = i === to.p ? to.o : p.text.length;
    // Own insertions disappear; everything else in range gets the deletion mark.
    const drop: [number, number][] = [];
    let pos = 0;
    for (const r of p.runs) {
      const ra = pos, rb = pos + r.len;
      pos = rb;
      const lo = Math.max(ra, a), hi = Math.min(rb, b);
      if (hi > lo && r.props.rev?.kind === "ins" && r.props.rev.author === rev.author) drop.push([lo, hi]);
    }
    let next = mapRange(p, a, b, (props) => (props.rev?.kind === "del" ? props : intern({ ...props, rev })));
    if (drop.length) {
      const keep: [number, number][] = [];
      let x = 0;
      for (const [lo, hi] of drop) { if (lo > x) keep.push([x, lo]); x = hi; }
      if (x < p.text.length) keep.push([x, p.text.length]);
      next = keepRanges(next, keep);
    }
    if (i < to.p && !(p.mark.rev?.kind === "ins" && p.mark.rev.author === rev.author)) next = { ...next, mark: intern({ ...next.mark, rev }) };
    paras[i] = next;
  }
  doc.paras = paras;
  // Own inserted paragraph marks inside the range join their paragraphs.
  for (let i = to.p - 1; i >= from.p; i--) {
    const p = doc.paras[i];
    if (p.mark.rev?.kind === "ins" && p.mark.rev.author === rev.author && i + 1 < doc.paras.length) {
      doc.replace({ p: i, o: p.text.length }, { p: i + 1, o: 0 }, []);
    }
  }
}

function mergeRanges(r: [number, number][]): [number, number][] {
  const out: [number, number][] = [];
  for (const x of r.filter(([a, b]) => b > a).sort((p, q) => p[0] - q[0])) {
    const last = out[out.length - 1];
    if (last && x[0] <= last[1]) last[1] = Math.max(last[1], x[1]); else out.push([...x] as [number, number]);
  }
  return out;
}

/** Applies `change` to the run properties of [a, b) of a paragraph. */
function mapRange(p: Para, a: number, b: number, change: (props: CharProps) => CharProps): Para {
  if (b <= a) return p;
  const runs: { len: number; props: CharProps }[] = [];
  let pos = 0;
  for (const r of p.runs) {
    const ra = pos, rb = pos + r.len;
    pos = rb;
    const lo = Math.max(ra, a), hi = Math.min(rb, b);
    if (hi <= lo) { runs.push(r); continue; }
    if (ra < lo) runs.push({ len: lo - ra, props: r.props });
    runs.push({ len: hi - lo, props: change(r.props) });
    if (hi < rb) runs.push({ len: rb - hi, props: r.props });
  }
  return { ...p, runs: normalizeRuns(runs) };
}

/** Whether text or paragraph marks in [from, to) carry revisions. */
export function hasRevisions(paras: readonly Para[], from: Pos, to: Pos): boolean {
  for (let i = from.p; i <= to.p && i < paras.length; i++) {
    const p = paras[i];
    if (p.mark.rev && (i < to.p || from.p === to.p)) return true;
    let pos = 0;
    for (const r of p.runs) {
      const ra = pos, rb = pos + r.len;
      pos = rb;
      const a = i === from.p ? from.o : 0, b = i === to.p ? to.o : p.text.length;
      if (r.props.rev && Math.min(rb, b) > Math.max(ra, a)) return true;
      if (r.props.rev && a === b && ra <= a && a <= rb) return true;
    }
  }
  return false;
}

/**
 * Annehmen (accept) or Ablehnen (reject) the revisions in [from, to): insertions are kept
 * or removed, deletions removed or kept; marked paragraph marks join or stay.
 */
export function resolveRevisions(doc: Doc, from: Pos, to: Pos, accept: boolean): void {
  const paras = [...doc.paras];
  const joins: number[] = [];
  for (let i = from.p; i <= to.p && i < paras.length; i++) {
    const p = paras[i];
    if (p.block) continue;
    const a = i === from.p ? from.o : 0, b = i === to.p ? to.o : p.text.length;
    const whole = a === b && from.p === to.p;
    const keep: [number, number][] = [];
    const clear: [number, number][] = [];
    let pos = 0;
    for (const r of p.runs) {
      const ra = pos, rb = pos + r.len;
      pos = rb;
      const rev = r.props.rev;
      const lo = whole && ra <= a && a <= rb ? ra : Math.max(ra, a), hi = whole && ra <= a && a <= rb ? rb : Math.min(rb, b);
      if (!rev || hi <= lo) { keep.push([ra, rb]); continue; }
      const drop = (rev.kind === "ins") !== accept;
      if (ra < lo) keep.push([ra, lo]);
      if (!drop) { keep.push([lo, hi]); clear.push([lo, hi]); }
      if (hi < rb) keep.push([hi, rb]);
    }
    // Remove the revision from kept text first (offsets before dropping), then drop.
    let next = p;
    for (const [x, y] of clear) next = mapRange(next, x, y, (props) => intern({ ...props, rev: undefined }));
    next = keepRanges(next, mergeRanges(keep));
    const mrev = p.mark.rev;
    if (mrev && (i < to.p || from.p === to.p)) {
      const joinIt = (mrev.kind === "del") === accept;
      next = { ...next, mark: intern({ ...next.mark, rev: undefined }) };
      if (joinIt && i + 1 < paras.length) joins.push(i);
    }
    paras[i] = next;
  }
  doc.paras = paras;
  for (const i of joins.reverse()) {
    const p = doc.paras[i];
    if (i + 1 < doc.paras.length) doc.replace({ p: i, o: p.text.length }, { p: i + 1, o: 0 }, []);
  }
}

/** The next (or previous) revision after (before) a position: its range. */
export function findRevision(paras: readonly Para[], at: Pos, dir: 1 | -1): { from: Pos; to: Pos } | null {
  const spans: { from: Pos; to: Pos }[] = [];
  paras.forEach((p, i) => {
    let pos = 0;
    let cur: { from: Pos; to: Pos; rev: Revision } | null = null;
    for (const r of p.runs) {
      const ra = pos, rb = pos + r.len;
      pos = rb;
      if (r.props.rev) {
        if (cur && cur.rev === r.props.rev && cur.to.o === ra) cur.to = { p: i, o: rb };
        else { if (cur) spans.push(cur); cur = { from: { p: i, o: ra }, to: { p: i, o: rb }, rev: r.props.rev }; }
      } else if (cur) { spans.push(cur); cur = null; }
    }
    if (cur) spans.push(cur);
    if (p.mark.rev) spans.push({ from: { p: i, o: p.text.length }, to: { p: i, o: p.text.length } });
  });
  if (dir > 0) return spans.find((s) => cmp(s.from, at) > 0) ?? null;
  return [...spans].reverse().find((s) => cmp(s.from, at) < 0) ?? null;
}
