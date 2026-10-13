// Styles: definitions with inheritance (basedOn), document defaults and theme fonts, and the
// resolution of a run's or paragraph's effective formatting the way Word does it:
// defaults → paragraph style chain → character style chain → direct formatting. Toggle
// properties (bold, italic, caps …) set by both the paragraph and the character style
// cancel out (ECMA-376 §17.7.3).

import type { Align, CharProps, LineRule, ParaProps, TabStop } from "./model";

export type StyleType = "paragraph" | "character" | "table" | "numbering";

export interface StyleDef {
  readonly id: string;
  /** Built-in English name as stored in files ("heading 1") or a custom name. */
  readonly name: string;
  readonly type: StyleType;
  readonly basedOn?: string;
  readonly next?: string;
  readonly link?: string;
  readonly isDefault?: boolean;
  readonly uiPriority?: number;
  readonly qFormat?: boolean;
  readonly semiHidden?: boolean;
  readonly pPr: ParaProps;
  readonly rPr: CharProps;
  /** Original definition from the file, written back unchanged. */
  readonly data?: unknown;
}

export interface Theme {
  readonly majorFont: string;
  readonly minorFont: string;
}

export interface ResolvedChar {
  readonly font: string;
  readonly size: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: string;
  readonly strike: boolean;
  readonly dstrike: boolean;
  readonly vertAlign: "superscript" | "subscript" | "baseline";
  readonly color: string;
  readonly highlight: string;
  readonly caps: boolean;
  readonly smallCaps: boolean;
  readonly hidden: boolean;
  readonly spacing: number;
  readonly kern: number;
}

export interface ResolvedPara {
  readonly style: string;
  readonly align: Align;
  readonly indLeft: number;
  readonly indRight: number;
  readonly indFirst: number;
  readonly before: number;
  readonly after: number;
  readonly line: number;
  readonly lineRule: LineRule;
  readonly contextual: boolean;
  readonly keepNext: boolean;
  readonly keepLines: boolean;
  readonly widow: boolean;
  readonly pageBreakBefore: boolean;
  readonly tabs: readonly TabStop[];
}

const TOGGLES = ["bold", "italic", "caps", "smallCaps", "strike", "dstrike", "hidden"] as const;

/** German names Word shows for built-in styles. */
export const GERMAN_NAMES: Record<string, string> = {
  "normal": "Standard", "no spacing": "Kein Leerraum", "title": "Titel", "subtitle": "Untertitel",
  "quote": "Zitat", "intense quote": "Intensives Zitat", "list paragraph": "Listenabsatz",
  "subtle emphasis": "Schwache Hervorhebung", "emphasis": "Hervorhebung", "intense emphasis": "Intensive Hervorhebung",
  "strong": "Fett", "subtle reference": "Schwacher Verweis", "intense reference": "Intensiver Verweis",
  "book title": "Buchtitel", "caption": "Beschriftung", "hyperlink": "Hyperlink",
  "default paragraph font": "Absatz-Standardschriftart", "footnote text": "Fußnotentext", "footnote reference": "Fußnotenzeichen",
  "header": "Kopfzeile", "footer": "Fußzeile", "toc heading": "Inhaltsverzeichnisüberschrift",
  ...Object.fromEntries(Array.from({ length: 9 }, (_, i) => [`heading ${i + 1}`, `Überschrift ${i + 1}`])),
  ...Object.fromEntries(Array.from({ length: 9 }, (_, i) => [`toc ${i + 1}`, `Verzeichnis ${i + 1}`])),
};

function merge<T extends object>(base: T, over: object | undefined): T {
  if (!over) return base;
  const out: Record<string, unknown> = { ...(base as Record<string, unknown>) };
  for (const [k, v] of Object.entries(over)) if (v !== undefined && k !== "extra") out[k] = v;
  return out as T;
}

/** Character merge: a font name given at a level replaces an inherited theme font (as in Word). */
function mergeR(base: CharProps, over: CharProps | undefined): CharProps {
  const out = merge(base, over);
  if (over?.font !== undefined && over.fontTheme === undefined) delete (out as { fontTheme?: string }).fontTheme;
  return out;
}

function mergeTabs(base: readonly TabStop[], over: readonly TabStop[] | undefined): TabStop[] {
  if (!over?.length) return [...base];
  let out = [...base];
  for (const t of over) {
    out = out.filter((x) => Math.abs(x.pos - t.pos) > 0.05);
    if (t.align !== "clear") out.push(t);
  }
  return out.sort((a, b) => a.pos - b.pos);
}

export class StyleSheet {
  private paraCache = new WeakMap<ParaProps, ResolvedPara>();
  private charCache = new Map<string, Map<CharProps, ResolvedChar>>();
  readonly defaultParagraphStyle: string;

  constructor(
    readonly styles: ReadonlyMap<string, StyleDef>,
    readonly defaultsR: CharProps,
    readonly defaultsP: ParaProps,
    readonly theme: Theme,
  ) {
    const def = [...styles.values()].find((s) => s.type === "paragraph" && s.isDefault);
    this.defaultParagraphStyle = def?.id ?? "Standard";
  }

  get(id: string | undefined): StyleDef | undefined {
    return id ? this.styles.get(id) : undefined;
  }

  /** Style id for a built-in name ("heading 1"), if the document defines it. */
  idForName(name: string): string | undefined {
    const n = name.toLowerCase();
    for (const s of this.styles.values()) if (s.name.toLowerCase() === n) return s.id;
    return undefined;
  }

  displayName(id: string): string {
    const s = this.styles.get(id);
    if (!s) return id;
    return GERMAN_NAMES[s.name.toLowerCase()] ?? s.name;
  }

  /** Returns a sheet with an extra style (from the built-in set when a document lacks it). */
  withStyle(def: StyleDef): StyleSheet {
    const map = new Map(this.styles);
    map.set(def.id, def);
    return new StyleSheet(map, this.defaultsR, this.defaultsP, this.theme);
  }

  private chain(id: string | undefined): StyleDef[] {
    const out: StyleDef[] = [];
    const seen = new Set<string>();
    for (let s = this.get(id); s && !seen.has(s.id); s = this.get(s.basedOn)) {
      seen.add(s.id);
      out.unshift(s);
    }
    return out;
  }

  private styleRPr(id: string | undefined): CharProps {
    let r: CharProps = {};
    for (const s of this.chain(id)) r = mergeR(r, s.rPr);
    return r;
  }

  resolvePara(p: ParaProps): ResolvedPara {
    let hit = this.paraCache.get(p);
    if (hit) return hit;
    const styleId = p.style ?? this.defaultParagraphStyle;
    let acc: ParaProps = { ...this.defaultsP };
    let tabs: TabStop[] = [...(this.defaultsP.tabs ?? [])];
    for (const s of this.chain(styleId)) {
      acc = merge(acc, s.pPr);
      tabs = mergeTabs(tabs, s.pPr.tabs);
    }
    acc = merge(acc, p);
    tabs = mergeTabs(tabs, p.tabs);
    hit = {
      style: styleId,
      align: acc.align ?? "left",
      indLeft: acc.indLeft ?? 0,
      indRight: acc.indRight ?? 0,
      indFirst: acc.indFirst ?? 0,
      before: acc.before ?? 0,
      after: acc.after ?? 0,
      line: acc.line ?? 1,
      lineRule: acc.lineRule ?? "auto",
      contextual: acc.contextual ?? false,
      keepNext: acc.keepNext ?? false,
      keepLines: acc.keepLines ?? false,
      widow: acc.widow ?? true,
      pageBreakBefore: acc.pageBreakBefore ?? false,
      tabs,
    };
    this.paraCache.set(p, hit);
    return hit;
  }

  /** Effective formatting of a run (or a paragraph mark) in a paragraph with `paraStyle`. */
  resolveChar(paraStyle: string | undefined, r: CharProps): ResolvedChar {
    const key = paraStyle ?? this.defaultParagraphStyle;
    let byStyle = this.charCache.get(key);
    if (!byStyle) { byStyle = new Map(); this.charCache.set(key, byStyle); }
    let hit = byStyle.get(r);
    if (hit) return hit;
    const fromPara = mergeR({ ...this.defaultsR }, this.styleRPr(key));
    let acc: Record<string, unknown> = { ...fromPara };
    if (r.rStyle) {
      const fromChar = this.styleRPr(r.rStyle);
      if (fromChar.font !== undefined && fromChar.fontTheme === undefined) delete acc.fontTheme;
      for (const [k, v] of Object.entries(fromChar)) {
        if (v === undefined || k === "extra") continue;
        acc[k] = (TOGGLES as readonly string[]).includes(k) && fromPara[k as keyof CharProps] === true && v === true ? false : v;
      }
    }
    const c = mergeR(acc as CharProps, r);
    const font = c.fontTheme ? (c.fontTheme.startsWith("major") ? this.theme.majorFont : this.theme.minorFont) : (c.font ?? this.theme.minorFont);
    hit = {
      font,
      size: c.size ?? 10,
      bold: !!c.bold,
      italic: !!c.italic,
      underline: c.underline ?? "none",
      strike: !!c.strike,
      dstrike: !!c.dstrike,
      vertAlign: c.vertAlign ?? "baseline",
      color: c.color ?? "auto",
      highlight: c.highlight ?? "none",
      caps: !!c.caps,
      smallCaps: !!c.smallCaps,
      hidden: !!c.hidden,
      spacing: c.spacing ?? 0,
      kern: c.kern ?? 0,
    };
    byStyle.set(r, hit);
    return hit;
  }
}

// ----- the built-in style set for new documents ---------------------------------------------

const BLUE = "2F5496", DARK_BLUE = "1F3763", ACCENT = "4472C4", GREY_TEXT = "404040", SUB = "5A5A5A";

function p(id: string, name: string, pPr: ParaProps, rPr: CharProps, extra: Partial<StyleDef> = {}): StyleDef {
  return { id, name, type: "paragraph", basedOn: "Standard", next: "Standard", qFormat: true, pPr, rPr, ...extra };
}

function c(id: string, name: string, rPr: CharProps, uiPriority: number): StyleDef {
  return { id, name, type: "character", basedOn: "Absatz-Standardschriftart", uiPriority, qFormat: true, pPr: {}, rPr };
}

const heading = (n: number, before: number, size: number, color: string): StyleDef =>
  p(`berschrift${n}`, `heading ${n}`, { keepNext: true, keepLines: true, before, after: 0, outline: n - 1 },
    { fontTheme: "majorHAnsi", size, color }, { uiPriority: 9, link: `berschrift${n}Zchn` });

/** Word 2016–2019 defaults (Calibri 11, D-011) with the style ids German Word writes. */
export function defaultStyleSheet(): StyleSheet {
  const defs: StyleDef[] = [
    { id: "Standard", name: "Normal", type: "paragraph", isDefault: true, qFormat: true, pPr: {}, rPr: {} },
    heading(1, 12, 16, BLUE),
    heading(2, 2, 13, BLUE),
    heading(3, 2, 12, DARK_BLUE),
    p("Titel", "Title", { after: 0, line: 1, lineRule: "auto", contextual: true }, { fontTheme: "majorHAnsi", size: 28, spacing: -0.5, kern: 14 }, { uiPriority: 10 }),
    p("Untertitel", "Subtitle", {}, { color: SUB, spacing: 0.75 }, { uiPriority: 11 }),
    p("Zitat", "Quote", { before: 10, align: "center", indLeft: 43.2, indRight: 43.2 }, { italic: true, color: GREY_TEXT }, { uiPriority: 29 }),
    p("IntensivesZitat", "Intense Quote", { before: 18, after: 18, align: "center", indLeft: 43.2, indRight: 43.2 }, { italic: true, color: ACCENT }, { uiPriority: 30 }),
    p("KeinLeerraum", "No Spacing", { after: 0, line: 1, lineRule: "auto" }, {}, { uiPriority: 1, basedOn: undefined }),
    p("Listenabsatz", "List Paragraph", { indLeft: 36, contextual: true }, {}, { uiPriority: 34 }),
    { id: "Absatz-Standardschriftart", name: "Default Paragraph Font", type: "character", isDefault: true, uiPriority: 1, semiHidden: true, pPr: {}, rPr: {} },
    c("SchwacheHervorhebung", "Subtle Emphasis", { italic: true, color: GREY_TEXT }, 19),
    c("Hervorhebung", "Emphasis", { italic: true }, 20),
    c("IntensiveHervorhebung", "Intense Emphasis", { italic: true, color: ACCENT }, 21),
    c("Fett", "Strong", { bold: true }, 22),
    c("SchwacherVerweis", "Subtle Reference", { smallCaps: true, color: SUB }, 31),
    c("IntensiverVerweis", "Intense Reference", { bold: true, smallCaps: true, color: ACCENT, spacing: 0.25 }, 32),
    c("Buchtitel", "Book Title", { bold: true, italic: true, spacing: 0.25 }, 33),
  ];
  return new StyleSheet(
    new Map(defs.map((d) => [d.id, d])),
    { fontTheme: "minorHAnsi", size: 11, lang: "de-DE" },
    { after: 8, line: 1.08, lineRule: "auto", widow: true },
    { majorFont: "Calibri Light", minorFont: "Calibri" },
  );
}

/** The built-in definition for a name (to add to documents that lack it). */
export function builtinStyle(name: string): StyleDef | undefined {
  const sheet = defaultStyleSheet();
  const id = sheet.idForName(name);
  return id ? sheet.get(id) : undefined;
}
