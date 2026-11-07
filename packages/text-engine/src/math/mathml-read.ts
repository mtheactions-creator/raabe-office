// MathML (presentation markup) read into the equation model (D-079): what Windows' Math Input Panel puts on the
// clipboard for a handwritten equation (Freihandgleichung), and MathML pasted from elsewhere. Elements:
// mrow, mi, mn, mo, mtext, ms, mspace, mfrac, msqrt, mroot, msup, msub, msubsup, munder, mover, munderover,
// mfenced, mtable, mtr, mtd, mstyle, mpadded, mphantom, semantics, menclose; n-ary operators (∑ ∫ ∏ …) take the
// element after them as their operand, as the linear format does.

import type { MNode, Seq } from "./ast";

const NARY = new Set(["\u2211", "\u220f", "\u2210", "\u222b", "\u222c", "\u222d", "\u222e", "\u222f", "\u2230", "\u22c3", "\u22c2", "\u22c1", "\u22c0", "\u2a01", "\u2a02", "\u2a00"]);
const ACCENT_CHARS = new Set(["^", "\u02c6", "\u0302", "~", "\u02dc", "\u0303", "\u2192", "\u20d7", "\u02d9", "\u0307", "\u00a8", "\u0308", "\u02c7", "\u030c", "\u02d8", "\u0306", "\u00b4", "\u0301", "`", "\u0300"]);
const ACCENT_MAP: Record<string, string> = { "^": "\u0302", "\u02c6": "\u0302", "~": "\u0303", "\u02dc": "\u0303", "\u2192": "\u20d7", "\u02d9": "\u0307", "\u00a8": "\u0308", "\u02c7": "\u030c", "\u02d8": "\u0306", "\u00b4": "\u0301", "`": "\u0300" };

const kids = (el: Element) => Array.from(el.children);
const text = (el: Element) => (el.textContent ?? "").replace(/\s+/g, " ").trim();

/** MathML source (a <math> element, with or without a namespace prefix) as the equation model, or null. */
export function mathmlToSeq(src: string): Seq | null {
  const start = src.search(/<(\w+:)?math[\s>]/);
  if (start < 0) return null;
  const end = src.search(/<\/(\w+:)?math>/);
  const xml = src.slice(start, end < 0 ? undefined : end + src.slice(end).indexOf(">") + 1);
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) return null;
  const root = doc.documentElement;
  const seq = row(kids(root));
  return seq.length ? seq : null;
}

/** A sequence of elements as one Seq (n-ary operators take the next element as their operand). */
function row(els: Element[]): Seq {
  const out: Seq = [];
  for (let i = 0; i < els.length; i++) {
    const el = els[i];
    const node = one(el);
    // An n-ary operator (alone or with limits) takes the next element as its operand.
    const nary = naryOf(el);
    if (nary) {
      const e = i + 1 < els.length ? one(els[++i]) : [];
      out.push({ ...nary, e });
      continue;
    }
    out.push(...node);
  }
  return out;
}

function naryOf(el: Element): Extract<MNode, { k: "nary" }> | null {
  const name = el.localName;
  const opOf = (x: Element | undefined) => (x && x.localName === "mo" && NARY.has(text(x)) ? text(x) : null);
  if (name === "mo" && opOf(el)) return { k: "nary", chr: opOf(el)!, e: [], und: false };
  const c = kids(el);
  if (["munderover", "msubsup"].includes(name) && opOf(c[0])) return { k: "nary", chr: opOf(c[0])!, sub: arg(c[1]), sup: arg(c[2]), e: [], und: name === "munderover" };
  if (["munder", "msub"].includes(name) && opOf(c[0])) return { k: "nary", chr: opOf(c[0])!, sub: arg(c[1]), e: [], und: name === "munder" };
  if (["mover", "msup"].includes(name) && opOf(c[0])) return { k: "nary", chr: opOf(c[0])!, sup: arg(c[1]), e: [], und: name === "mover" };
  return null;
}

const arg = (el: Element | undefined): Seq => (el ? one(el) : []);

/** One element as a Seq (most are one node; mrow and the like are several). */
function one(el: Element): Seq {
  const c = kids(el);
  switch (el.localName) {
    case "math": case "mrow": case "mstyle": case "mpadded": case "merror": return row(c);
    case "mphantom": return [];
    case "semantics": return c[0] ? one(c[0]) : [];
    case "annotation": case "annotation-xml": return [];
    case "mi": {
      const s = text(el);
      const variant = el.getAttribute("mathvariant");
      // A single-letter identifier is italic by default; a longer one (a function name) upright.
      return s ? [{ k: "t", s, sty: variant === "normal" || s.length > 1 ? "p" : variant === "bold" ? "b" : variant === "bold-italic" ? "bi" : undefined }] : [];
    }
    case "mn": case "mo": { const s = text(el); return s ? [{ k: "t", s }] : []; }
    case "mtext": case "ms": { const s = el.textContent ?? ""; return s ? [{ k: "t", s, nor: true, sty: "p" }] : []; }
    case "mspace": return [{ k: "t", s: " " }];
    case "mfrac": return [{ k: "frac", num: arg(c[0]), den: arg(c[1]), type: el.getAttribute("linethickness") === "0" ? "noBar" : el.getAttribute("bevelled") === "true" ? "skw" : undefined }];
    case "msqrt": return [{ k: "rad", e: row(c) }];
    case "mroot": return [{ k: "rad", deg: arg(c[1]), e: arg(c[0]) }];
    case "msup": return [{ k: "scr", base: arg(c[0]), sup: arg(c[1]) }];
    case "msub": return [{ k: "scr", base: arg(c[0]), sub: arg(c[1]) }];
    case "msubsup": return [{ k: "scr", base: arg(c[0]), sub: arg(c[1]), sup: arg(c[2]) }];
    case "mover": case "munder": case "munderover": {
      const base = arg(c[0]);
      const over = el.localName === "munder" ? undefined : c[el.localName === "mover" ? 1 : 2];
      const under = el.localName === "mover" ? undefined : c[1];
      const ch = over && over.localName === "mo" ? text(over) : "";
      if (over && !under && (el.getAttribute("accent") === "true" || ACCENT_CHARS.has(ch))) return [{ k: "acc", chr: ACCENT_MAP[ch] ?? ch, e: base }];
      if (over && !under && (ch === "¯" || ch === "‾" || ch === "_")) return [{ k: "bar", top: true, e: base }];
      if (under && !over && under.localName === "mo" && ["_", "¯", "‾"].includes(text(under))) return [{ k: "bar", top: false, e: base }];
      let node: MNode = { k: "lim", upper: false, e: base, lim: arg(under) };
      if (!under) node = { k: "lim", upper: true, e: base, lim: arg(over) };
      else if (over) node = { k: "lim", upper: true, e: [{ k: "lim", upper: false, e: base, lim: arg(under) }], lim: arg(over) };
      return [node];
    }
    case "mfenced": {
      const open = el.getAttribute("open") ?? "(", close = el.getAttribute("close") ?? ")";
      return [{ k: "d", beg: open, end: close, items: c.length ? c.map((x) => one(x)) : [[]] }];
    }
    case "mtable": return [{ k: "m", rows: c.filter((r) => r.localName === "mtr" || r.localName === "mlabeledtr").map((r) => kids(r).filter((d) => d.localName === "mtd").map((d) => row(kids(d)))) }];
    case "menclose": return [{ k: "box", e: row(c) }];
    default: return row(c);
  }
}
