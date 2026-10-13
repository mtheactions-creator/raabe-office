// Small XML helpers for WordprocessingML.

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

export function parse(xml: string): Document {
  return new DOMParser().parseFromString(xml, "application/xml");
}

const serializer = () => new XMLSerializer();

/** Serialises an element; namespace declarations needed by it are added by the serializer. */
export function outer(el: Element): string {
  return serializer().serializeToString(el);
}

export function kids(el: Element): Element[] {
  return Array.from(el.children);
}

export function child(el: Element, name: string): Element | null {
  for (const c of Array.from(el.children)) if (c.namespaceURI === W && c.localName === name) return c;
  return null;
}

export function attr(el: Element | null, name: string): string | null {
  if (!el) return null;
  return el.getAttributeNS(W, name) ?? el.getAttribute(`w:${name}`);
}

/** OOXML on/off value: present without w:val means on. */
export function onOff(el: Element | null): boolean | undefined {
  if (!el) return undefined;
  const v = attr(el, "val");
  return v === null || !(v === "0" || v === "false" || v === "off");
}

export function num(v: string | null): number | undefined {
  if (v === null || v === "") return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/** All w:t text below an element (for display text of preserved content). */
export function textOf(el: Element, includeDeleted = false): string {
  let out = "";
  const walk = (e: Element) => {
    for (const c of Array.from(e.children)) {
      if (c.localName === "t" || (includeDeleted && c.localName === "delText")) out += c.textContent ?? "";
      else if (c.localName === "tab") out += "\t";
      else if (c.localName === "br" || c.localName === "cr") out += " ";
      else walk(c);
    }
  };
  walk(el);
  return out;
}
