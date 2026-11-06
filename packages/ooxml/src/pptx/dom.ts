// DOM helpers for PresentationML. A presentation's parts stay XML documents in memory (D-070):
// edits change the elements the model understands and leave every other element and attribute
// as PowerPoint wrote it, so whatever RaabePresent does not know is written back verbatim.

export const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
export const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
export const NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
export const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
export const NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/** EMU per point and per CSS pixel at 96 dpi. */
export const EMU_PT = 12700;
export const EMU_PX = 9525;

const PREFIX_NS: Record<string, string> = { a: NS_A, p: NS_P, r: NS_R, c: NS_C, p14: NS_P14 };

export function parseXml(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const err = doc.getElementsByTagName("parsererror")[0];
  if (err) throw new Error(`XML: ${err.textContent?.slice(0, 200)}`);
  return doc;
}

export function serialize(doc: Document | Element): string {
  const root = "documentElement" in doc ? doc.documentElement : doc;
  return XML_DECL + new XMLSerializer().serializeToString(root);
}

export const kids = (el: Element | undefined | null, name?: string): Element[] =>
  el ? Array.from(el.children).filter((c) => !name || c.localName === name) : [];

export const kid = (el: Element | undefined | null, name: string): Element | undefined =>
  el ? Array.from(el.children).find((c) => c.localName === name) : undefined;

/** A descendant along local names ("nvSpPr/nvPr/ph"). */
export function at(el: Element | undefined | null, path: string): Element | undefined {
  let cur: Element | undefined = el ?? undefined;
  for (const name of path.split("/")) {
    cur = kid(cur, name);
    if (!cur) return undefined;
  }
  return cur;
}

export function num(el: Element | undefined | null, name: string): number | undefined {
  const v = el?.getAttribute(name);
  if (v === null || v === undefined || v === "") return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

export function bool(el: Element | undefined | null, name: string): boolean | undefined {
  const v = el?.getAttribute(name);
  if (v === null || v === undefined) return undefined;
  return v === "1" || v === "true" || v === "on";
}

/** A new element "a:t" in its namespace, owned by `doc`. */
export function make(doc: Document, qname: string, attrs: Record<string, string | number | undefined> = {}, ...children: (Element | string)[]): Element {
  const [prefix, local] = qname.includes(":") ? qname.split(":") : ["", qname];
  const ns = prefix ? PREFIX_NS[prefix] : doc.documentElement.namespaceURI;
  const el = doc.createElementNS(ns ?? null, prefix ? `${prefix}:${local}` : local);
  for (const [k, v] of Object.entries(attrs)) if (v !== undefined) el.setAttribute(k, String(v));
  for (const c of children) el.append(typeof c === "string" ? doc.createTextNode(c) : c);
  return el;
}

/** Parses a fragment (with a:/p:/r: prefixes) into elements owned by `doc`. */
export function fragment(doc: Document, xml: string): Element[] {
  const wrap = parseXml(`<w xmlns:a="${NS_A}" xmlns:p="${NS_P}" xmlns:r="${NS_R}" xmlns:c="${NS_C}" xmlns:p14="${NS_P14}">${xml}</w>`);
  return Array.from(wrap.documentElement.children).map((c) => doc.importNode(c, true) as Element);
}

/** The first child of `parent` named `name`, created (and placed before the first of `before`) when missing. */
export function ensure(parent: Element, qname: string, before: readonly string[] = []): Element {
  const local = qname.split(":").pop()!;
  const hit = kid(parent, local);
  if (hit) return hit;
  const el = make(parent.ownerDocument, qname);
  const next = Array.from(parent.children).find((c) => before.includes(c.localName));
  parent.insertBefore(el, next ?? null);
  return el;
}

export function removeKids(parent: Element | undefined, ...names: string[]): void {
  for (const c of kids(parent)) if (names.includes(c.localName)) c.remove();
}

export const esc = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
