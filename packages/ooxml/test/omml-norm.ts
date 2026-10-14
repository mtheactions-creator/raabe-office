// The structure of an equation as catalogue_math.py normalises it: only what changes the look.

export const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
export const KEEP_PR = new Set(["type", "chr", "limLoc", "begChr", "endChr", "sepChr", "pos", "degHide", "subHide", "supHide", "sty", "scr", "nor", "count"]);

/** Python's repr of a string (single quotes unless the text has one). */
export const repr = (s: string) => (s.includes("'") && !s.includes('"') ? `"${s}"` : `'${s.replace(/\\/g, "\\\\").replace(/'/g, "\\'")}'`);

export function norm(el: Element, keep: ReadonlySet<string> = KEEP_PR): string {
  const tag = el.localName;
  if (el.namespaceURI !== M_NS) return "";
  if (tag === "t") return repr(el.textContent ?? "");
  if (tag === "r") {
    const texts = Array.from(el.getElementsByTagNameNS(M_NS, "t")).map((t) => t.textContent ?? "").join("");
    if (!texts) return "";
    const rpr = Array.from(el.children).find((c) => c.localName === "rPr" && c.namespaceURI === M_NS);
    const props = rpr ? Array.from(rpr.children).filter((c) => keep.has(c.localName)).map((c) => `${c.localName}=${c.getAttributeNS(M_NS, "val") ?? "1"}`).join(",") : "";
    return props ? `r[${props}]${repr(texts)}` : repr(texts);
  }
  if (tag.endsWith("Pr")) {
    const vals = Array.from(el.children).filter((c) => keep.has(c.localName)).map((c) => `${c.localName}=${c.getAttributeNS(M_NS, "val") || "1"}`);
    return vals.length ? `@${vals.join(",")}` : "";
  }
  if (tag === "ctrlPr") return "";
  const inner = Array.from(el.children).map((c) => norm(c, keep)).filter(Boolean);
  return `${tag}(${inner.join(" ")})`;
}


/** Stricter: also the properties a round trip must keep (growing, shapes, matrix layout). */
export const KEEP_ALL = new Set([...KEEP_PR, "grow", "shp", "plcHide", "baseJc", "mcJc", "argSz", "brk", "aln", "lit", "opEmu", "noBreak", "jc", "alnScr"]);
