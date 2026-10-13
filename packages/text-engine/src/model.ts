// The document model. A paragraph is its text plus a partition of that text into runs of
// identical character formatting, direct paragraph properties and the properties of the
// paragraph mark. Content the engine does not understand (fields, hyperlinks, images,
// bookmarks, tables …) travels along as opaque payloads so that files can be written back
// without loss (D-004). All values are immutable: edits create new paragraph objects, which
// keeps undo cheap and lets the layout cache per paragraph object.

export type Underline = "single" | "double" | "thick" | "dotted" | "dash" | "wave" | "words" | "none" | string;

/** Character formatting. `undefined` means "inherit from styles". Sizes in points. */
export interface CharProps {
  readonly font?: string;
  /** Theme font reference ("minorHAnsi", "majorHAnsi") when the font comes from the theme. */
  readonly fontTheme?: string;
  readonly size?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: Underline;
  readonly strike?: boolean;
  readonly dstrike?: boolean;
  readonly vertAlign?: "superscript" | "subscript" | "baseline";
  /** "RRGGBB" or "auto". */
  readonly color?: string;
  /** Theme colour attributes as read from the file, kept while the colour is unchanged. */
  readonly colorTheme?: string;
  /** OOXML highlight name ("yellow", "green" …) or "none". */
  readonly highlight?: string;
  readonly caps?: boolean;
  readonly smallCaps?: boolean;
  readonly hidden?: boolean;
  /** Extra character spacing in points (Laufweite). */
  readonly spacing?: number;
  /** Kerning applies from this font size (points) upwards; 0 = off. */
  readonly kern?: number;
  readonly lang?: string;
  /** Character style id. */
  readonly rStyle?: string;
  /** Run properties the engine does not interpret, kept for writing back. */
  readonly extra?: readonly unknown[];
}

export type Align = "left" | "center" | "right" | "both";
export type LineRule = "auto" | "exact" | "atLeast";

export interface TabStop {
  /** Position in points from the left margin. */
  readonly pos: number;
  readonly align: "left" | "center" | "right" | "decimal" | "bar" | "clear";
  readonly leader?: "none" | "dot" | "hyphen" | "underscore" | "heavy" | "middleDot";
}

/** Paragraph formatting. Lengths in points. */
export interface ParaProps {
  readonly style?: string;
  readonly align?: Align;
  readonly indLeft?: number;
  readonly indRight?: number;
  /** First-line indent; negative = hanging indent. */
  readonly indFirst?: number;
  readonly before?: number;
  readonly after?: number;
  readonly beforeAuto?: boolean;
  readonly afterAuto?: boolean;
  /** With lineRule "auto": multiple of single spacing (1 = single). Otherwise points. */
  readonly line?: number;
  readonly lineRule?: LineRule;
  readonly contextual?: boolean;
  readonly keepNext?: boolean;
  readonly keepLines?: boolean;
  readonly widow?: boolean;
  readonly pageBreakBefore?: boolean;
  readonly tabs?: readonly TabStop[];
  readonly outline?: number;
  /** Paragraph properties the engine does not interpret (numbering, borders …). */
  readonly extra?: readonly unknown[];
}

export interface Run {
  readonly len: number;
  readonly props: CharProps;
}

/** Something inline the engine shows but does not edit inside: a field, a link, an image … */
export interface InlineObject {
  readonly kind: "field" | "link" | "image" | "note" | "revision" | "other";
  /** Opaque payload for the file format layer. */
  readonly data: unknown;
  /** Text shown in place of the object (field result, link text …). */
  readonly display: string;
  /** Size of a placeholder box (images), in points. */
  readonly width?: number;
  readonly height?: number;
}

/** A zero-width marker kept at an offset (bookmark, comment range …). */
export interface Anchor {
  readonly at: number;
  readonly data: unknown;
}

/** A block the engine does not edit (table, multi-paragraph field …), shown as a placeholder. */
export interface BlockObject {
  readonly kind: "table" | "field" | "other";
  readonly data: unknown;
  readonly label: string;
  /** Text lines shown read-only inside the placeholder. */
  readonly lines: readonly string[];
  /** Estimated height in points (used for pagination until the block is supported). */
  readonly height: number;
}

export interface Para {
  /** Text; "\t" tab, "\u000b" line break, "\u000c" page break, "\ufffc" object. */
  readonly text: string;
  /** Formatting runs covering `text` exactly (empty for an empty paragraph). */
  readonly runs: readonly Run[];
  readonly props: ParaProps;
  /** Formatting of the paragraph mark (height of empty paragraphs, typing at the start). */
  readonly mark: CharProps;
  /** One entry per "\ufffc" in `text`, in order. */
  readonly objects?: readonly InlineObject[];
  readonly anchors?: readonly Anchor[];
  /** When set, the whole paragraph is this read-only block. */
  readonly block?: BlockObject;
  /** Opaque content that preceded this paragraph in the file (body-level bookmarks …). */
  readonly leading?: readonly unknown[];
}

export const OBJ = "\ufffc";
export const LINE_BREAK = "\u000b";
export const PAGE_BREAK = "\u000c";
export const EMPTY: CharProps = Object.freeze({});

// ----- interning --------------------------------------------------------------------------

const interned = new Map<string, CharProps>();

function propsKey(p: CharProps): string {
  const keys = Object.keys(p).filter((k) => (p as Record<string, unknown>)[k] !== undefined).sort();
  return keys.map((k) => {
    const v = (p as Record<string, unknown>)[k];
    return `${k}=${k === "extra" ? (v as unknown[]).map(String).join("\u0001") : String(v)}`;
  }).join("|");
}

/** Returns the canonical (shared, frozen) object for these properties, so runs compare by identity. */
export function intern(p: CharProps): CharProps {
  const clean: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(p)) if (v !== undefined) clean[k] = v;
  const key = propsKey(clean as CharProps);
  let hit = interned.get(key);
  if (!hit) {
    hit = Object.freeze(clean) as CharProps;
    interned.set(key, hit);
  }
  return hit;
}

// ----- paragraph helpers ------------------------------------------------------------------

export function makePara(text: string, props: ParaProps = {}, runProps: CharProps = EMPTY, mark: CharProps = EMPTY): Para {
  return { text, runs: text ? [{ len: text.length, props: intern(runProps) }] : [], props, mark: intern(mark) };
}

/** Merges neighbouring runs with identical properties and drops empty ones. */
export function normalizeRuns(runs: readonly Run[]): Run[] {
  const out: Run[] = [];
  for (const r of runs) {
    if (r.len <= 0) continue;
    const last = out[out.length - 1];
    if (last && last.props === r.props) out[out.length - 1] = { len: last.len + r.len, props: r.props };
    else out.push(r);
  }
  return out;
}

/** Character properties of the character at `offset` (the paragraph mark's at the end). */
export function propsAt(p: Para, offset: number): CharProps {
  let pos = 0;
  for (const r of p.runs) {
    if (offset < pos + r.len) return r.props;
    pos += r.len;
  }
  return p.mark;
}

/** Formatting new text takes at `offset`: that of the character before, as in Word. */
export function insertionProps(p: Para, offset: number): CharProps {
  if (!p.text.length) return p.mark;
  if (offset > 0) {
    // After a line break or object, continue with the following text's formatting.
    const before = p.text[offset - 1];
    if (before !== OBJ) return propsAt(p, offset - 1);
  }
  return offset < p.text.length ? propsAt(p, offset) : propsAt(p, offset - 1);
}

function sliceRuns(runs: readonly Run[], from: number, to: number): Run[] {
  const out: Run[] = [];
  let pos = 0;
  for (const r of runs) {
    const a = Math.max(from, pos), b = Math.min(to, pos + r.len);
    if (b > a) out.push({ len: b - a, props: r.props });
    pos += r.len;
    if (pos >= to) break;
  }
  return out;
}

function countObj(text: string, from: number, to: number): number {
  let n = 0;
  for (let i = from; i < to; i++) if (text.charCodeAt(i) === 0xfffc) n++;
  return n;
}

/** The part [from, to) of a paragraph, keeping runs, objects and anchors aligned. */
export function slicePara(p: Para, from: number, to: number): Para {
  const before = countObj(p.text, 0, from);
  const inside = countObj(p.text, from, to);
  const objects = p.objects?.slice(before, before + inside);
  const anchors = p.anchors?.filter((a) => a.at >= from && a.at <= to).map((a) => ({ at: a.at - from, data: a.data }));
  return {
    ...p,
    text: p.text.slice(from, to),
    runs: sliceRuns(p.runs, from, to),
    objects: objects?.length ? objects : undefined,
    anchors: anchors?.length ? anchors : undefined,
  };
}

/** Joins two paragraph pieces; the result keeps `a`'s paragraph properties and `b`'s mark. */
export function joinParas(a: Para, b: Para): Para {
  const shift = a.text.length;
  const objects = [...(a.objects ?? []), ...(b.objects ?? [])];
  const anchors = [...(a.anchors ?? []), ...(b.anchors ?? []).map((x) => ({ at: x.at + shift, data: x.data }))];
  return {
    ...a,
    text: a.text + b.text,
    runs: normalizeRuns([...a.runs, ...b.runs]),
    mark: b.mark,
    objects: objects.length ? objects : undefined,
    anchors: anchors.length ? anchors : undefined,
  };
}

/** Text with one formatting, as a paragraph piece (for inserting). */
export function textPiece(text: string, props: CharProps): Para {
  return { text, runs: text ? [{ len: text.length, props: intern(props) }] : [], props: {}, mark: EMPTY };
}

/** Changes character properties on [from, to) with `change(old) => new`. */
export function mapRuns(p: Para, from: number, to: number, change: (old: CharProps) => CharProps): Para {
  const out: Run[] = [];
  let pos = 0;
  for (const r of p.runs) {
    const start = pos, end = pos + r.len;
    pos = end;
    const a = Math.max(from, start), b = Math.min(to, end);
    if (b <= a) { out.push(r); continue; }
    if (a > start) out.push({ len: a - start, props: r.props });
    out.push({ len: b - a, props: intern(change(r.props)) });
    if (end > b) out.push({ len: end - b, props: r.props });
  }
  return { ...p, runs: normalizeRuns(out) };
}

/** Replaces the text of [from, to) with same-length text (change case), keeping formatting. */
export function replaceSameLength(p: Para, from: number, text: string): Para {
  return { ...p, text: p.text.slice(0, from) + text + p.text.slice(from + text.length) };
}

/** Checks the invariants; used by tests. */
export function validPara(p: Para): boolean {
  const total = p.runs.reduce((n, r) => n + r.len, 0);
  return total === p.text.length && countObj(p.text, 0, p.text.length) === (p.objects?.length ?? 0) &&
    p.runs.every((r) => r.len > 0);
}
