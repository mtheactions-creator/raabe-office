// Run and paragraph properties: reading w:rPr / w:pPr into the model and writing them back.
// Elements the model does not interpret are kept as XML strings and written back in the
// position the schema requires, so files stay valid and nothing is lost.

import type { CharProps, ParaProps, TabStop } from "@raabe/text-engine";
import { attr, child, esc, kids, num, onOff, outer, W } from "./xml";

// Child order required by the schema (ECMA-376 CT_RPr / CT_PPr).
const RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
  "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs",
  "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"];
const PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers",
  "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
  "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
  "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"];

const RPR_KNOWN = new Set(["rStyle", "rFonts", "b", "i", "caps", "smallCaps", "strike", "dstrike", "vanish", "color", "spacing", "kern", "sz", "highlight", "u", "vertAlign", "lang"]);
const PPR_KNOWN = new Set(["pStyle", "keepNext", "keepLines", "pageBreakBefore", "widowControl", "tabs", "spacing", "ind", "contextualSpacing", "jc", "outlineLvl", "rPr"]);

/** A preserved element: its local name (for ordering) and XML. */
export interface Extra {
  name: string;
  xml: string;
}

export function isExtra(x: unknown): x is Extra {
  return typeof x === "object" && x !== null && "name" in x && "xml" in x;
}

const twipsToPt = (v: string | null) => { const n = num(v); return n === undefined ? undefined : n / 20; };
const ptToTwips = (pt: number) => String(Math.round(pt * 20));

// ----- reading --------------------------------------------------------------------------

export function readRPr(el: Element | null): CharProps {
  if (!el) return {};
  const p: Record<string, unknown> = {};
  const extra: Extra[] = [];
  for (const c of kids(el)) {
    if (c.namespaceURI !== W || !RPR_KNOWN.has(c.localName)) { extra.push({ name: c.localName, xml: outer(c) }); continue; }
    switch (c.localName) {
      case "rStyle": p.rStyle = attr(c, "val") ?? undefined; break;
      case "rFonts": {
        p.font = attr(c, "ascii") ?? attr(c, "hAnsi") ?? undefined;
        p.fontTheme = attr(c, "asciiTheme") ?? attr(c, "hAnsiTheme") ?? undefined;
        // Other script slots (east Asian, complex script) are kept as they are.
        const others = ["eastAsia", "cs", "eastAsiaTheme", "cstheme", "hint"].filter((a) => attr(c, a) !== null);
        if (others.length) extra.push({ name: "rFontsOther", xml: others.map((a) => `w:${a}="${esc(attr(c, a)!)}"`).join(" ") });
        break;
      }
      case "b": p.bold = onOff(c); break;
      case "i": p.italic = onOff(c); break;
      case "caps": p.caps = onOff(c); break;
      case "smallCaps": p.smallCaps = onOff(c); break;
      case "strike": p.strike = onOff(c); break;
      case "dstrike": p.dstrike = onOff(c); break;
      case "vanish": p.hidden = onOff(c); break;
      case "color": {
        p.color = attr(c, "val") ?? undefined;
        const theme = ["themeColor", "themeShade", "themeTint"].filter((a) => attr(c, a) !== null);
        if (theme.length) p.colorTheme = theme.map((a) => `w:${a}="${esc(attr(c, a)!)}"`).join(" ");
        break;
      }
      case "spacing": p.spacing = twipsToPt(attr(c, "val")); break;
      case "kern": { const v = num(attr(c, "val")); p.kern = v === undefined ? undefined : v / 2; break; }
      case "sz": { const v = num(attr(c, "val")); p.size = v === undefined ? undefined : v / 2; break; }
      case "highlight": p.highlight = attr(c, "val") ?? undefined; break;
      case "u": p.underline = attr(c, "val") ?? "single"; break;
      case "vertAlign": p.vertAlign = (attr(c, "val") ?? "baseline") as CharProps["vertAlign"]; break;
      case "lang": {
        p.lang = attr(c, "val") ?? undefined;
        const others = ["eastAsia", "bidi"].filter((a) => attr(c, a) !== null);
        if (others.length) extra.push({ name: "langOther", xml: others.map((a) => `w:${a}="${esc(attr(c, a)!)}"`).join(" ") });
        break;
      }
    }
  }
  if (extra.length) p.extra = extra.map((e) => `${e.name}\u0000${e.xml}`);
  return p as CharProps;
}

function tabs(el: Element): TabStop[] {
  return kids(el).filter((t) => t.localName === "tab").map((t) => ({
    pos: (num(attr(t, "pos")) ?? 0) / 20,
    align: (({ start: "left", end: "right", num: "left" } as Record<string, string>)[attr(t, "val") ?? "left"] ?? attr(t, "val") ?? "left") as TabStop["align"],
    leader: (attr(t, "leader") ?? undefined) as TabStop["leader"],
  }));
}

export function readPPr(el: Element | null): { props: ParaProps; mark: CharProps } {
  if (!el) return { props: {}, mark: {} };
  const p: Record<string, unknown> = {};
  const extra: string[] = [];
  let mark: CharProps = {};
  for (const c of kids(el)) {
    if (c.namespaceURI !== W || !PPR_KNOWN.has(c.localName)) { extra.push(`${c.localName}\u0000${outer(c)}`); continue; }
    switch (c.localName) {
      case "pStyle": p.style = attr(c, "val") ?? undefined; break;
      case "keepNext": p.keepNext = onOff(c); break;
      case "keepLines": p.keepLines = onOff(c); break;
      case "pageBreakBefore": p.pageBreakBefore = onOff(c); break;
      case "widowControl": p.widow = onOff(c); break;
      case "contextualSpacing": p.contextual = onOff(c); break;
      case "outlineLvl": p.outline = num(attr(c, "val")); break;
      case "tabs": p.tabs = tabs(c); break;
      case "jc": {
        const v = attr(c, "val");
        p.align = v === "center" ? "center" : v === "right" || v === "end" ? "right" : v === "both" || v === "distribute" ? "both" : "left";
        break;
      }
      case "spacing": {
        p.before = twipsToPt(attr(c, "before"));
        p.after = twipsToPt(attr(c, "after"));
        if (attr(c, "beforeAutospacing") !== null) p.beforeAuto = attr(c, "beforeAutospacing") !== "0";
        if (attr(c, "afterAutospacing") !== null) p.afterAuto = attr(c, "afterAutospacing") !== "0";
        const line = num(attr(c, "line"));
        const rule = attr(c, "lineRule") ?? "auto";
        if (line !== undefined) {
          p.lineRule = rule === "exact" ? "exact" : rule === "atLeast" ? "atLeast" : "auto";
          p.line = p.lineRule === "auto" ? line / 240 : line / 20;
        }
        const keep = ["beforeLines", "afterLines"].filter((a) => attr(c, a) !== null);
        if (keep.length) extra.push(`spacingOther\u0000${keep.map((a) => `w:${a}="${esc(attr(c, a)!)}"`).join(" ")}`);
        break;
      }
      case "ind": {
        const left = attr(c, "left") ?? attr(c, "start");
        const right = attr(c, "right") ?? attr(c, "end");
        p.indLeft = twipsToPt(left);
        p.indRight = twipsToPt(right);
        const hanging = twipsToPt(attr(c, "hanging"));
        const first = twipsToPt(attr(c, "firstLine"));
        p.indFirst = hanging !== undefined ? -hanging : first;
        const keep = ["leftChars", "rightChars", "startChars", "endChars", "hangingChars", "firstLineChars"].filter((a) => attr(c, a) !== null);
        if (keep.length) extra.push(`indOther\u0000${keep.map((a) => `w:${a}="${esc(attr(c, a)!)}"`).join(" ")}`);
        break;
      }
      case "rPr": mark = readRPr(c); break;
    }
  }
  if (extra.length) p.extra = extra;
  return { props: p as ParaProps, mark };
}

// ----- writing --------------------------------------------------------------------------

function splitExtra(list: readonly unknown[] | undefined): { name: string; xml: string }[] {
  return (list ?? []).map((x) => {
    const s = String(x);
    const i = s.indexOf("\u0000");
    return i < 0 ? { name: "zzz", xml: s } : { name: s.slice(0, i), xml: s.slice(i + 1) };
  });
}

function ordered(parts: { name: string; xml: string }[], order: string[]): string {
  const idx = (n: string) => { const i = order.indexOf(n); return i < 0 ? order.length : i; };
  return parts.map((p, i) => ({ ...p, i })).sort((a, b) => idx(a.name) - idx(b.name) || a.i - b.i).map((p) => p.xml).join("");
}

const flag = (name: string, v: boolean | undefined) => (v === undefined ? null : `<w:${name}${v ? "" : ' w:val="0"'}/>`);

export interface WriteOptions {
  /** Theme font name lookup; when set, theme references are written as explicit font names too. */
  themeFonts?: { major: string; minor: string };
  /** The package has a theme part, so theme references may be written. */
  hasTheme: boolean;
}

export function writeRPr(p: CharProps, opts: WriteOptions, tag = "rPr"): string {
  const parts: { name: string; xml: string }[] = [];
  const extra = splitExtra(p.extra);
  const add = (name: string, xml: string | null) => { if (xml) parts.push({ name, xml }); };
  if (p.rStyle) add("rStyle", `<w:rStyle w:val="${esc(p.rStyle)}"/>`);
  const fontsOther = extra.find((e) => e.name === "rFontsOther")?.xml;
  if (p.font || p.fontTheme || fontsOther) {
    const a: string[] = [];
    let name = p.font;
    if (p.fontTheme && opts.themeFonts && !name) name = p.fontTheme.startsWith("major") ? opts.themeFonts.major : opts.themeFonts.minor;
    if (name) a.push(`w:ascii="${esc(name)}" w:hAnsi="${esc(name)}"`);
    if (p.fontTheme && opts.hasTheme) a.push(`w:asciiTheme="${esc(p.fontTheme)}" w:hAnsiTheme="${esc(p.fontTheme)}"`);
    if (fontsOther) a.push(fontsOther);
    add("rFonts", `<w:rFonts ${a.join(" ")}/>`);
  }
  add("b", flag("b", p.bold));
  add("i", flag("i", p.italic));
  add("caps", flag("caps", p.caps));
  add("smallCaps", flag("smallCaps", p.smallCaps));
  add("strike", flag("strike", p.strike));
  add("dstrike", flag("dstrike", p.dstrike));
  add("vanish", flag("vanish", p.hidden));
  if (p.color) add("color", `<w:color w:val="${esc(p.color)}"${p.colorTheme && opts.hasTheme ? ` ${p.colorTheme}` : ""}/>`);
  if (p.spacing !== undefined) add("spacing", `<w:spacing w:val="${ptToTwips(p.spacing)}"/>`);
  if (p.kern !== undefined) add("kern", `<w:kern w:val="${Math.round(p.kern * 2)}"/>`);
  if (p.size !== undefined) add("sz", `<w:sz w:val="${Math.round(p.size * 2)}"/>`);
  if (p.highlight) add("highlight", `<w:highlight w:val="${esc(p.highlight)}"/>`);
  if (p.underline) add("u", `<w:u w:val="${esc(p.underline)}"/>`);
  if (p.vertAlign) add("vertAlign", `<w:vertAlign w:val="${esc(p.vertAlign)}"/>`);
  const langOther = extra.find((e) => e.name === "langOther")?.xml;
  if (p.lang || langOther) add("lang", `<w:lang${p.lang ? ` w:val="${esc(p.lang)}"` : ""}${langOther ? ` ${langOther}` : ""}/>`);
  for (const e of extra) if (e.name !== "rFontsOther" && e.name !== "langOther") parts.push(e);
  const body = ordered(parts, RPR_ORDER);
  return body ? `<w:${tag}>${body}</w:${tag}>` : "";
}

export function writePPr(p: ParaProps, mark: CharProps, opts: WriteOptions): string {
  const parts: { name: string; xml: string }[] = [];
  const extra = splitExtra(p.extra);
  const add = (name: string, xml: string | null) => { if (xml) parts.push({ name, xml }); };
  if (p.style) add("pStyle", `<w:pStyle w:val="${esc(p.style)}"/>`);
  add("keepNext", flag("keepNext", p.keepNext));
  add("keepLines", flag("keepLines", p.keepLines));
  add("pageBreakBefore", flag("pageBreakBefore", p.pageBreakBefore));
  add("widowControl", flag("widowControl", p.widow));
  if (p.tabs?.length) {
    add("tabs", `<w:tabs>${p.tabs.map((t) => `<w:tab w:val="${esc(t.align)}"${t.leader && t.leader !== "none" ? ` w:leader="${esc(t.leader)}"` : ""} w:pos="${ptToTwips(t.pos)}"/>`).join("")}</w:tabs>`);
  }
  const spacingOther = extra.find((e) => e.name === "spacingOther")?.xml;
  if (p.before !== undefined || p.after !== undefined || p.line !== undefined || p.beforeAuto || p.afterAuto || spacingOther) {
    const a: string[] = [];
    if (p.before !== undefined) a.push(`w:before="${ptToTwips(p.before)}"`);
    if (p.beforeAuto) a.push('w:beforeAutospacing="1"');
    if (p.after !== undefined) a.push(`w:after="${ptToTwips(p.after)}"`);
    if (p.afterAuto) a.push('w:afterAutospacing="1"');
    if (p.line !== undefined) {
      const rule = p.lineRule ?? "auto";
      a.push(`w:line="${rule === "auto" ? Math.round(p.line * 240) : ptToTwips(p.line)}" w:lineRule="${rule}"`);
    }
    if (spacingOther) a.push(spacingOther);
    add("spacing", `<w:spacing ${a.join(" ")}/>`);
  }
  const indOther = extra.find((e) => e.name === "indOther")?.xml;
  if (p.indLeft !== undefined || p.indRight !== undefined || p.indFirst !== undefined || indOther) {
    const a: string[] = [];
    if (p.indLeft !== undefined) a.push(`w:left="${ptToTwips(p.indLeft)}"`);
    if (p.indRight !== undefined) a.push(`w:right="${ptToTwips(p.indRight)}"`);
    if (p.indFirst !== undefined) a.push(p.indFirst < 0 ? `w:hanging="${ptToTwips(-p.indFirst)}"` : `w:firstLine="${ptToTwips(p.indFirst)}"`);
    if (indOther) a.push(indOther);
    add("ind", `<w:ind ${a.join(" ")}/>`);
  }
  add("contextualSpacing", flag("contextualSpacing", p.contextual));
  if (p.align) add("jc", `<w:jc w:val="${p.align}"/>`);
  if (p.outline !== undefined) add("outlineLvl", `<w:outlineLvl w:val="${p.outline}"/>`);
  add("rPr", writeRPr(mark, opts) || null);
  for (const e of extra) if (e.name !== "spacingOther" && e.name !== "indOther") parts.push(e);
  const body = ordered(parts, PPR_ORDER);
  return body ? `<w:pPr>${body}</w:pPr>` : "";
}

export { child };
