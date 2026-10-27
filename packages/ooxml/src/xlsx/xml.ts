// String-level XML for SpreadsheetML. Worksheets can be large, and whatever RaabeCalc does
// not model is written back exactly as Excel wrote it (D-024), so parts are cut into their
// top-level elements as text instead of being parsed into a DOM.

/** An element found in a string: [start, end) is the whole element. */
export interface El {
  /** Qualified name as written ("c", "x14ac:dyDescent", "mc:AlternateContent"). */
  name: string;
  start: number;
  end: number;
  /** End of the start tag (index after its ">"). */
  open: number;
  /** Start of the end tag ("</"), or `end` for an empty element. */
  close: number;
}

/** Index after the markup that starts at `i` ("<" of a tag, comment, PI or CDATA). */
function skipTag(xml: string, i: number): number {
  if (xml.startsWith("<!--", i)) return end(xml.indexOf("-->", i + 4), 3, xml);
  if (xml.startsWith("<![CDATA[", i)) return end(xml.indexOf("]]>", i + 9), 3, xml);
  if (xml.startsWith("<?", i)) return end(xml.indexOf("?>", i + 2), 2, xml);
  if (xml.startsWith("<!", i)) return end(xml.indexOf(">", i + 2), 1, xml);
  // A start or end tag: ">" outside quotes.
  let q = "";
  for (let j = i + 1; j < xml.length; j++) {
    const ch = xml[j];
    if (q) { if (ch === q) q = ""; }
    else if (ch === '"' || ch === "'") q = ch;
    else if (ch === ">") return j + 1;
  }
  return xml.length;
}

function end(at: number, len: number, xml: string): number {
  return at < 0 ? xml.length : at + len;
}

function tagName(xml: string, i: number): string {
  let j = i;
  while (j < xml.length) {
    const c = xml.charCodeAt(j);
    // whitespace, "/", ">"
    if (c === 32 || c === 9 || c === 10 || c === 13 || c === 47 || c === 62) break;
    j++;
  }
  return xml.slice(i, j);
}

/** The direct child elements of the range [from, to) (the content of one element). */
export function children(xml: string, from: number, to: number): El[] {
  const out: El[] = [];
  let i = xml.indexOf("<", from);
  while (i >= 0 && i < to) {
    const c = xml.charCodeAt(i + 1);
    if (c === 33 || c === 63) { i = xml.indexOf("<", skipTag(xml, i)); continue; } // "!" "?"
    if (c === 47) { i = xml.indexOf("<", skipTag(xml, i)); continue; } // stray end tag
    const el = element(xml, i);
    out.push(el);
    i = xml.indexOf("<", el.end);
  }
  return out;
}

/** The element whose start tag begins at `i`. */
export function element(xml: string, i: number): El {
  const name = tagName(xml, i + 1);
  const open = skipTag(xml, i);
  if (xml.charCodeAt(open - 2) === 47) return { name, start: i, end: open, open, close: open };
  // Find the matching end tag, counting nested elements.
  let depth = 1;
  let j = open;
  for (;;) {
    j = xml.indexOf("<", j);
    if (j < 0) return { name, start: i, end: xml.length, open, close: xml.length };
    const c = xml.charCodeAt(j + 1);
    const after = skipTag(xml, j);
    if (c === 47) {
      if (--depth === 0) return { name, start: i, end: after, open, close: j };
    } else if (c !== 33 && c !== 63 && xml.charCodeAt(after - 2) !== 47) depth++;
    j = after;
  }
}

/** The document's root element (after the XML declaration). */
export function root(xml: string): El {
  let i = xml.indexOf("<");
  while (i >= 0 && (xml.charCodeAt(i + 1) === 63 || xml.charCodeAt(i + 1) === 33)) i = xml.indexOf("<", skipTag(xml, i));
  return element(xml, i);
}

/** Local name without prefix. */
export function local(name: string): string {
  const k = name.indexOf(":");
  return k < 0 ? name : name.slice(k + 1);
}

/** The first direct child with this (qualified) name. */
export function child(xml: string, el: El, name: string): El | undefined {
  return children(xml, el.open, el.close).find((c) => c.name === name);
}

export function text(xml: string, el: El): string {
  return xml.slice(el.open, el.close);
}

export function outer(xml: string, el: El): string {
  return xml.slice(el.start, el.end);
}

/** Attributes of an element's start tag, values unescaped. */
export function attrs(xml: string, el: El): Record<string, string> {
  return parseAttrs(xml.slice(el.start, el.open));
}

const ATTR = /([^\s=<>/"']+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

export function parseAttrs(tag: string): Record<string, string> {
  const out: Record<string, string> = {};
  const sp = tag.search(/\s/);
  if (sp < 0) return out;
  ATTR.lastIndex = sp;
  for (let m; (m = ATTR.exec(tag)); ) out[m[1]] = unesc(m[2] ?? m[3] ?? "");
  return out;
}

/** The start tag of an element with attributes set (string) or removed (null); others kept in place. */
export function setAttrs(tag: string, change: Record<string, string | null | undefined>): string {
  let out = tag;
  for (const [name, value] of Object.entries(change)) {
    if (value === undefined) continue;
    const re = new RegExp(`(\\s)${name.replace(/[.:]/g, "\\$&")}\\s*=\\s*("[^"]*"|'[^']*')`);
    if (re.test(out)) out = out.replace(re, value === null ? "" : `$1${name}="${esc(value)}"`);
    else if (value !== null) {
      const at = out.endsWith("/>") ? out.length - 2 : out.length - 1;
      out = `${out.slice(0, at).replace(/\s+$/, "")} ${name}="${esc(value)}"${out.slice(at)}`;
    }
  }
  return out;
}

/** The XML declaration as Excel writes it (CR LF after it). */
export const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** Text content as Excel writes it (quotes stay as they are). */
export function escText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

const ENT: Record<string, string> = { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" };

export function unesc(s: string): string {
  // Line ends are normalised by XML (written ones; &#13; stays a CR).
  if (s.indexOf("\r") >= 0) s = s.replace(/\r\n?/g, "\n");
  if (s.indexOf("&") < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, (_, e: string) =>
    e[0] === "#" ? String.fromCodePoint(e[1] === "x" ? parseInt(e.slice(2), 16) : Number(e.slice(1))) : ENT[e]);
}

/**
 * Text of an ST_Xstring: Excel writes characters XML cannot hold as _xHHHH_ (and a literal
 * "_x" sequence as _x005F_xHHHH_).
 */
export function decodeX(s: string): string {
  if (s.indexOf("_x") < 0) return s;
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h: string) => String.fromCharCode(parseInt(h, 16)));
}

export function encodeX(s: string): string {
  // eslint-disable-next-line no-control-regex
  return s.replace(/_x(?=[0-9A-Fa-f]{4}_)/g, "_x005F_x").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g,
    (c) => `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")}_`);
}

/** <t> element text with xml:space="preserve" where Excel writes it. */
export function tXml(s: string, tag = "t"): string {
  const t = esc(encodeX(s));
  return /^[\s]|[\s]$/.test(s) ? `<${tag} xml:space="preserve">${t}</${tag}>` : `<${tag}>${t}</${tag}>`;
}

/** A part cut into its root element's children, so each can be kept or replaced. */
export interface Split {
  /** Everything before the first child (declaration and root start tag). */
  head: string;
  parts: { name: string; xml: string }[];
  /** The root end tag and anything after it. */
  tail: string;
  /** The root element's start tag alone. */
  rootTag: string;
}

export function split(xml: string): Split {
  const r = root(xml);
  if (r.open === r.end) {
    // An empty root (<sheetData/> style): give it an end tag so children can be added.
    const tag = xml.slice(r.start, r.end);
    const open = tag.slice(0, -2).replace(/\s+$/, "") + ">";
    return { head: xml.slice(0, r.start) + open, parts: [], tail: `</${r.name}>` + xml.slice(r.end), rootTag: open };
  }
  const kids = children(xml, r.open, r.close);
  const parts = kids.map((k) => ({ name: k.name, xml: outer(xml, k) }));
  return { head: xml.slice(0, kids.length ? kids[0].start : r.close), parts, tail: xml.slice(kids.length ? kids[kids.length - 1].end : r.close), rootTag: xml.slice(r.start, r.open) };
}

export function joinSplit(s: Split): string {
  return s.head + s.parts.map((p) => p.xml).join("") + s.tail;
}

/**
 * Puts a part in schema order: replaces the existing one of that name, removes it (xml null)
 * or inserts it before the first part that comes later in `order`. Unknown parts keep the
 * rank of the known part before them.
 */
export function place(s: Split, name: string, xml: string | null, order: readonly string[]): void {
  const at = s.parts.findIndex((p) => p.name === name);
  if (at >= 0) {
    if (xml === null) s.parts.splice(at, 1);
    else s.parts[at] = { name, xml };
    return;
  }
  if (xml === null) return;
  const rank = order.indexOf(name);
  let last = -1;
  let i = 0;
  for (; i < s.parts.length; i++) {
    const r = order.indexOf(s.parts[i].name);
    if (r >= 0) last = r;
    if (last > rank) break;
  }
  s.parts.splice(i, 0, { name, xml });
}

/** Replaces all parts of a name (conditionalFormatting …) by a list, in schema order. */
export function placeAll(s: Split, name: string, xmls: string[], order: readonly string[]): void {
  const first = s.parts.findIndex((p) => p.name === name);
  s.parts = s.parts.filter((p) => p.name !== name);
  if (!xmls.length) return;
  let at = first;
  if (at < 0) {
    place(s, name, "", order);
    at = s.parts.findIndex((p) => p.name === name);
    s.parts.splice(at, 1);
  }
  s.parts.splice(at, 0, ...xmls.map((xml) => ({ name, xml })));
}
