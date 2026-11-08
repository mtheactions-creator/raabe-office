// UserForms as Office stores them in a VBA project (D-080): the designer storage of a form module, read into our
// form description (forms.ts) and written from it. Written from the published specifications [MS-OFORMS] (the
// controls' binary formats), [MS-OVBA] 2.2.10/2.3.5 (designer storage, VBFrame) and [MS-OLEDS] 2.3.8 (CompObj);
// the comments describe the format in our own words.
//
// The storage of a form (named like the module's stream) holds:
// - "\x01CompObj": the OLE class information (user type "Microsoft Forms 2.0 Form", ProgID Forms.Form.1);
// - "\x03VBFrame": text with the designer's CLSID, the form's Caption, ClientWidth/ClientHeight in twips and its
//   StartUpPosition;
// - "f": the form control (FormControl): its properties, then the list of its controls' sites (name, ID, tab
//   index, position, which control type, how many bytes its data takes in "o");
// - "o": the data of every simple control, one after the other in site order;
// - "iNN" (NN the control's ID, at least two digits): a storage of its own for each container control (Frame),
//   again with "f", "o" and "\x01CompObj".
//
// Every control record has the same shape: a version, the byte count of what follows, a bit mask of the
// properties that differ from the format's defaults, then those properties: the small ones (4 bytes or less, each
// aligned to its own size relative to the start of the record, the block padded to 4), then the larger ones
// (sizes, positions, strings padded to 4), then pictures and fonts. Strings whose characters all fit in one byte
// are stored as one byte per character ("compressed"), others as UTF-16. Lengths are HIMETRIC (1/100 mm).
//
// Reading keeps what we do not interpret: unknown properties of a known control stay in its raw values and are
// written back; controls we do not model (MultiPage, TabStrip, other ActiveX controls, nested sites) are kept as
// their bytes; and a form that was not edited is written back byte for byte as it was read.
//
// Writer's choices (for D-100+): DrawBuffer 32000 (Office's value), LogicalSize equal to DisplayedSize, the form's
// font as a StdFont, the controls' fonts as TextProps, sites with depth 0 and SITE_TYPE ST_Ole, no class table
// (only built-in controls are written), no design extender, Frames with their own storage and fPromoteControls.

import type { VbaControlDesign, VbaControlType, VbaFont, VbaFormDesign } from "./forms";
import { newControl, newForm } from "./forms";

export const CLSID_FORM = "{C62A69F0-16DC-11CE-9E98-00AA00574A4F}";
export const CLSID_FRAME = "{6E182020-F460-11CE-9BCD-00AA00608E01}";
const GUID_STDFONT = "{0BE35203-8F91-11CE-9DE3-00AA004BB851}";
const GUID_TEXTPROPS = "{AFC20920-DA4E-11CE-B943-00AA006887B4}";
const HIMETRIC_PER_PT = 2540 / 72;

/** Points to HIMETRIC, and back rounded to 0.05 pt as Office shows sizes. */
const toHm = (pt: number) => Math.round(pt * HIMETRIC_PER_PT);
const toPt = (hm: number) => Math.round((hm / HIMETRIC_PER_PT) * 20) / 20;

// ----- bytes ---------------------------------------------------------------------------------------------------

export function guidBytes(g: string): number[] {
  const h = g.replace(/[{}-]/g, "");
  const b = (i: number) => parseInt(h.slice(i * 2, i * 2 + 2), 16);
  return [b(3), b(2), b(1), b(0), b(5), b(4), b(7), b(6), ...Array.from({ length: 8 }, (_, i) => b(8 + i))];
}
export function guidString(b: Uint8Array, at = 0): string {
  const x = (i: number) => b[at + i].toString(16).padStart(2, "0").toUpperCase();
  return `{${x(3)}${x(2)}${x(1)}${x(0)}-${x(5)}${x(4)}-${x(7)}${x(6)}-${x(8)}${x(9)}-${[10, 11, 12, 13, 14, 15].map(x).join("")}}`;
}
/** A fresh random GUID (VB_Base of a new form). */
export function newGuid(): string {
  const r = crypto.getRandomValues(new Uint8Array(16));
  r[6] = (r[6] & 0x0f) | 0x40; r[8] = (r[8] & 0x3f) | 0x80;
  const h = Array.from(r, (x) => x.toString(16).padStart(2, "0").toUpperCase()).join("");
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}}`;
}

class Out {
  b: number[] = [];
  get len() { return this.b.length; }
  u8(n: number) { this.b.push(n & 255); }
  u16(n: number) { this.b.push(n & 255, (n >>> 8) & 255); }
  u32(n: number) { this.b.push(n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255); }
  bytes(x: ArrayLike<number>) { for (let i = 0; i < x.length; i++) this.b.push(x[i]); }
  align(n: number, from = 0) { while ((this.b.length - from) % n) this.b.push(0); }
  setU16(at: number, n: number) { this.b[at] = n & 255; this.b[at + 1] = (n >>> 8) & 255; }
}
const rd = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength);

/** fmString: one byte per character when every character fits in one byte, otherwise UTF-16. */
function encodeStr(s: string): { bytes: number[]; compressed: boolean } {
  if (s && Array.from(s).every((c) => c.charCodeAt(0) < 256)) return { bytes: Array.from(s, (c) => c.charCodeAt(0)), compressed: true };
  const bytes: number[] = [];
  for (let i = 0; i < s.length; i++) bytes.push(s.charCodeAt(i) & 255, s.charCodeAt(i) >>> 8);
  return { bytes, compressed: false };
}
function decodeStr(b: Uint8Array, at: number, cb: number, compressed: boolean): string {
  let s = "";
  if (compressed) for (let i = 0; i < cb; i++) s += String.fromCharCode(b[at + i]);
  else for (let i = 0; i + 1 < cb; i += 2) s += String.fromCharCode(b[at + i] | (b[at + i + 1] << 8));
  return s;
}

// ----- the generic control record ----------------------------------------------------------------------------

type Kind = "u8" | "u16" | "i16" | "u32" | "i32" | "str" | "size" | "pos" | "flag" | "pic" | "font";
interface Field { bit: number; name: string; kind: Kind }
type Values = Map<string, unknown>;
const F = (bit: number, name: string, kind: Kind): Field => ({ bit, name, kind });

// The property masks of [MS-OFORMS] 2.2 in bit order (the order of the data in the record).
const FORM_FIELDS = [F(1, "BackColor", "u32"), F(2, "ForeColor", "u32"), F(3, "NextAvailableID", "u32"), F(6, "BooleanProperties", "u32"), F(7, "BorderStyle", "u8"), F(8, "MousePointer", "u8"), F(9, "ScrollBars", "u8"),
  F(10, "DisplayedSize", "size"), F(11, "LogicalSize", "size"), F(12, "ScrollPosition", "pos"), F(13, "GroupCnt", "u32"), F(15, "MouseIcon", "pic"), F(16, "Cycle", "u8"), F(17, "SpecialEffect", "u8"), F(18, "BorderColor", "u32"),
  F(19, "Caption", "str"), F(20, "Font", "font"), F(21, "Picture", "pic"), F(22, "Zoom", "u32"), F(23, "PictureAlignment", "u8"), F(24, "PictureTiling", "flag"), F(25, "PictureSizeMode", "u8"), F(26, "ShapeCookie", "u32"), F(27, "DrawBuffer", "u32")];
const SITE_FIELDS = [F(0, "Name", "str"), F(1, "Tag", "str"), F(2, "ID", "i32"), F(3, "HelpContextID", "i32"), F(4, "BitFlags", "u32"), F(5, "ObjectStreamSize", "u32"), F(6, "TabIndex", "i16"), F(7, "ClsidCacheIndex", "u16"),
  F(8, "Position", "pos"), F(9, "GroupID", "u16"), F(11, "ControlTipText", "str"), F(12, "RuntimeLicKey", "str"), F(13, "ControlSource", "str"), F(14, "RowSource", "str")];
const TEXTPROPS_FIELDS = [F(0, "FontName", "str"), F(1, "FontEffects", "u32"), F(2, "FontHeight", "u32"), F(4, "FontCharSet", "u8"), F(5, "FontPitchAndFamily", "u8"), F(6, "ParagraphAlign", "u8"), F(7, "FontWeight", "u16")];
const BUTTON_FIELDS = [F(0, "ForeColor", "u32"), F(1, "BackColor", "u32"), F(2, "VariousPropertyBits", "u32"), F(3, "Caption", "str"), F(4, "PicturePosition", "u32"), F(5, "Size", "size"), F(6, "MousePointer", "u8"), F(7, "Picture", "pic"),
  F(8, "Accelerator", "u16"), F(9, "TakeFocusOnClick", "flag"), F(10, "MouseIcon", "pic")];
const LABEL_FIELDS = [F(0, "ForeColor", "u32"), F(1, "BackColor", "u32"), F(2, "VariousPropertyBits", "u32"), F(3, "Caption", "str"), F(4, "PicturePosition", "u32"), F(5, "Size", "size"), F(6, "MousePointer", "u8"), F(7, "BorderColor", "u32"),
  F(8, "BorderStyle", "u16"), F(9, "SpecialEffect", "u16"), F(10, "Picture", "pic"), F(11, "Accelerator", "u16"), F(12, "MouseIcon", "pic")];
const MORPH_FIELDS = [F(0, "VariousPropertyBits", "u32"), F(1, "BackColor", "u32"), F(2, "ForeColor", "u32"), F(3, "MaxLength", "u32"), F(4, "BorderStyle", "u8"), F(5, "ScrollBars", "u8"), F(6, "DisplayStyle", "u8"), F(7, "MousePointer", "u8"),
  F(8, "Size", "size"), F(9, "PasswordChar", "u16"), F(10, "ListWidth", "u32"), F(11, "BoundColumn", "u16"), F(12, "TextColumn", "i16"), F(13, "ColumnCount", "i16"), F(14, "ListRows", "u16"), F(15, "cColumnInfo", "u16"),
  F(16, "MatchEntry", "u8"), F(17, "ListStyle", "u8"), F(18, "ShowDropButtonWhen", "u8"), F(20, "DropButtonStyle", "u8"), F(21, "MultiSelect", "u8"), F(22, "Value", "str"), F(23, "Caption", "str"), F(24, "PicturePosition", "u32"),
  F(25, "BorderColor", "u32"), F(26, "SpecialEffect", "u32"), F(27, "MouseIcon", "pic"), F(28, "Picture", "pic"), F(29, "Accelerator", "u16"), F(31, "Reserved", "flag"), F(32, "GroupName", "str")];
const SCROLL_FIELDS = [F(0, "ForeColor", "u32"), F(1, "BackColor", "u32"), F(2, "VariousPropertyBits", "u32"), F(3, "Size", "size"), F(4, "MousePointer", "u8"), F(5, "Min", "i32"), F(6, "Max", "i32"), F(7, "Position", "i32"),
  F(9, "PrevEnabled", "i32"), F(10, "NextEnabled", "i32"), F(11, "SmallChange", "i32"), F(12, "LargeChange", "i32"), F(13, "Orientation", "i32"), F(14, "ProportionalThumb", "i16"), F(15, "Delay", "u32"), F(16, "MouseIcon", "pic")];
const SPIN_FIELDS = [F(0, "ForeColor", "u32"), F(1, "BackColor", "u32"), F(2, "VariousPropertyBits", "u32"), F(3, "Size", "size"), F(5, "Min", "i32"), F(6, "Max", "i32"), F(7, "Position", "i32"), F(8, "PrevEnabled", "i32"),
  F(9, "NextEnabled", "i32"), F(10, "SmallChange", "i32"), F(11, "Orientation", "i32"), F(12, "Delay", "u32"), F(13, "MouseIcon", "pic"), F(14, "MousePointer", "u8")];
const IMAGE_FIELDS = [F(2, "AutoSize", "flag"), F(3, "BorderColor", "u32"), F(4, "BackColor", "u32"), F(5, "BorderStyle", "u8"), F(6, "MousePointer", "u8"), F(7, "PictureSizeMode", "u8"), F(8, "SpecialEffect", "u8"), F(9, "Size", "size"),
  F(10, "Picture", "pic"), F(11, "PictureAlignment", "u8"), F(12, "PictureTiling", "flag"), F(13, "VariousPropertyBits", "u32"), F(14, "MouseIcon", "pic")];

/** Reads one record (header, mask, data, extra data, stream data) at `at`. */
function readRecord(b: Uint8Array, at: number, fields: Field[], maskBytes: 4 | 8): { values: Values; next: number } {
  const v = rd(b);
  const cb = v.getUint16(at + 2, true);
  const lo = v.getUint32(at + 4, true), hi = maskBytes === 8 ? v.getUint32(at + 8, true) : 0;
  const has = (bit: number) => (bit < 32 ? (lo >>> bit) & 1 : (hi >>> (bit - 32)) & 1) === 1;
  const known = new Set(fields.map((f) => f.bit));
  for (let bit = 0; bit < maskBytes * 8; bit++) if (has(bit) && !known.has(bit)) throw new Error(`unknown property bit ${bit}`);
  const values: Values = new Map();
  let p = at + 4 + maskBytes;
  const align = (n: number) => { while ((p - at) % n) p++; };
  const strs = new Map<string, [number, boolean]>();
  for (const f of fields) {
    if (!has(f.bit)) continue;
    switch (f.kind) {
      case "u8": values.set(f.name, b[p]); p += 1; break;
      case "u16": align(2); values.set(f.name, v.getUint16(p, true)); p += 2; break;
      case "i16": align(2); values.set(f.name, v.getInt16(p, true)); p += 2; break;
      case "u32": align(4); values.set(f.name, v.getUint32(p, true)); p += 4; break;
      case "i32": align(4); values.set(f.name, v.getInt32(p, true)); p += 4; break;
      case "str": { align(4); const n = v.getUint32(p, true); strs.set(f.name, [n & 0x7fffffff, (n & 0x80000000) !== 0]); p += 4; break; }
      case "pic": case "font": align(2); p += 2; break;
      case "flag": values.set(f.name, true); break;
    }
  }
  align(4);
  for (const f of fields) {
    if (!has(f.bit)) continue;
    if (f.kind === "size" || f.kind === "pos") { values.set(f.name, [v.getInt32(p, true), v.getInt32(p + 4, true)]); p += 8; }
    else if (f.kind === "str") { const [n, c] = strs.get(f.name)!; values.set(f.name, decodeStr(b, p, n, c)); p += n; align(4); }
  }
  p = at + 4 + cb;
  for (const f of fields) {
    if (!has(f.bit)) continue;
    if (f.kind === "pic") {
      // GuidAndPicture: CLSID_StdPicture, then the preamble 0x0000746C, the size, the picture.
      const size = v.getUint32(p + 20, true);
      values.set(f.name, b.slice(p, p + 24 + size));
      p += 24 + size;
    } else if (f.kind === "font") {
      const g = guidString(b, p);
      if (g === GUID_STDFONT) {
        const q = p + 16;
        const len = b[q + 10];
        values.set(f.name, { std: { charset: v.getInt16(q + 1, true), flags: b[q + 3], weight: v.getInt16(q + 4, true), height: v.getUint32(q + 6, true), face: decodeStr(b, q + 11, len, true) } });
        p = q + 11 + len;
      } else if (g === GUID_TEXTPROPS) {
        const r = readRecord(b, p + 16, TEXTPROPS_FIELDS, 4);
        values.set(f.name, { textProps: r.values });
        p = r.next;
      } else throw new Error(`unknown font ${g}`);
    }
  }
  return { values, next: p };
}

/** Writes one record from its values (a property is written when it has a value). */
function writeRecord(values: Values, fields: Field[], maskBytes: 4 | 8, minor: number, major: number): number[] {
  const o = new Out();
  let lo = 0, hi = 0;
  for (const f of fields) if (values.has(f.name)) { if (f.bit < 32) lo |= 1 << f.bit; else hi |= 1 << (f.bit - 32); }
  o.u8(minor); o.u8(major); o.u16(0); o.u32(lo >>> 0); if (maskBytes === 8) o.u32(hi >>> 0);
  const strs = new Map<string, { bytes: number[]; compressed: boolean }>();
  for (const f of fields) {
    if (!values.has(f.name)) continue;
    const x = values.get(f.name) as number;
    switch (f.kind) {
      case "u8": o.u8(x); break;
      case "u16": case "i16": o.align(2); o.u16(x); break;
      case "u32": case "i32": o.align(4); o.u32(x); break;
      case "str": { o.align(4); const s = encodeStr(String(values.get(f.name))); strs.set(f.name, s); o.u32((s.bytes.length | (s.compressed ? 0x80000000 : 0)) >>> 0); break; }
      case "pic": case "font": o.align(2); o.u16(0xffff); break;
    }
  }
  o.align(4);
  for (const f of fields) {
    if (!values.has(f.name)) continue;
    if (f.kind === "size" || f.kind === "pos") { const [a, b] = values.get(f.name) as [number, number]; o.u32(a); o.u32(b); }
    else if (f.kind === "str") { o.bytes(strs.get(f.name)!.bytes); o.align(4); }
  }
  o.setU16(2, o.len - 4);
  for (const f of fields) {
    if (!values.has(f.name)) continue;
    if (f.kind === "pic") o.bytes(values.get(f.name) as Uint8Array);
    else if (f.kind === "font") {
      const fv = values.get(f.name) as { std?: StdFont; textProps?: Values };
      if (fv.std) {
        o.bytes(guidBytes(GUID_STDFONT));
        const face = encodeStr(fv.std.face).bytes.slice(0, 31);
        o.u8(1); o.u16(fv.std.charset); o.u8(fv.std.flags); o.u16(fv.std.weight); o.u32(fv.std.height); o.u8(face.length); o.bytes(face);
      } else { o.bytes(guidBytes(GUID_TEXTPROPS)); o.bytes(writeRecord(fv.textProps!, TEXTPROPS_FIELDS, 4, 0, 2)); }
    }
  }
  return o.b;
}
interface StdFont { charset: number; flags: number; weight: number; height: number; face: string }

// ----- control types --------------------------------------------------------------------------------------------

/** ClsidCacheIndex values of the built-in controls ([MS-OFORMS] 2.4.5). */
const CACHE: Partial<Record<VbaControlType, number>> = { Image: 12, Frame: 14, SpinButton: 16, CommandButton: 17, TabStrip: 18, Label: 21, TextBox: 23, ListBox: 24, ComboBox: 25, CheckBox: 26, OptionButton: 27, ToggleButton: 28, ScrollBar: 47, MultiPage: 57 };
const BY_CACHE = new Map<number, VbaControlType>(Object.entries(CACHE).map(([k, v]) => [v as number, k as VbaControlType]));
const MORPH_STYLE: Partial<Record<VbaControlType, number>> = { TextBox: 1, ListBox: 2, ComboBox: 3, CheckBox: 4, OptionButton: 5, ToggleButton: 6 };
const STYLE_TYPE: Record<number, VbaControlType> = { 1: "TextBox", 2: "ListBox", 3: "ComboBox", 4: "CheckBox", 5: "OptionButton", 6: "ToggleButton", 7: "ComboBox" };
const isMorph = (t: VbaControlType) => t in MORPH_STYLE;
const fieldsOf = (t: VbaControlType): { fields: Field[]; mask: 4 | 8; text: boolean } | null => {
  if (t === "CommandButton") return { fields: BUTTON_FIELDS, mask: 4, text: true };
  if (t === "Label") return { fields: LABEL_FIELDS, mask: 4, text: true };
  if (isMorph(t)) return { fields: MORPH_FIELDS, mask: 8, text: true };
  if (t === "ScrollBar") return { fields: SCROLL_FIELDS, mask: 4, text: false };
  if (t === "SpinButton") return { fields: SPIN_FIELDS, mask: 4, text: false };
  if (t === "Image") return { fields: IMAGE_FIELDS, mask: 4, text: false };
  return null;
};

/** The format's defaults ([MS-OFORMS] 2.5) of the properties we map. */
const DEF_BACK: Partial<Record<VbaControlType, number>> = { CheckBox: 0x80000005, ComboBox: 0x80000005, ListBox: 0x80000005, OptionButton: 0x80000005, TextBox: 0x80000005, ToggleButton: 0x80000005, CommandButton: 0x8000000f, Image: 0x8000000f, Label: 0x8000000f, ScrollBar: 0x8000000f, SpinButton: 0x8000000f };
const DEF_FORE: Partial<Record<VbaControlType, number>> = { CheckBox: 0x80000008, ComboBox: 0x80000008, ListBox: 0x80000008, OptionButton: 0x80000008, TextBox: 0x80000008, ToggleButton: 0x80000008, CommandButton: 0x80000012, Label: 0x80000012, ScrollBar: 0x80000012, SpinButton: 0x80000012 };
const DEF_VARIOUS = (t: VbaControlType) => (isMorph(t) ? 0x2c80081b : t === "Label" ? 0x0080001b : 0x0000001b);
const SITE_FLAGS_DEFAULT = 0x33;
// VariousPropertiesBitfield bits.
const VB_ENABLED = 1 << 1, VB_LOCKED = 1 << 2, VB_BACKSTYLE = 1 << 3, VB_EDITABLE = 1 << 14, VB_WORDWRAP = 1 << 23, VB_AUTOSIZE = 1 << 28, VB_MULTILINE = 1 << 31;

function various(c: VbaControlDesign, start: number): number {
  let v = start;
  const set = (bit: number, on: boolean | undefined) => { if (on === undefined) return; v = on ? v | bit : v & ~bit; };
  set(VB_ENABLED, c.enabled ?? true);
  if (["CheckBox", "ComboBox", "CommandButton", "ListBox", "OptionButton", "TextBox", "ToggleButton"].includes(c.type)) set(VB_LOCKED, c.locked ?? false);
  if (!["ListBox", "ScrollBar", "SpinButton", "TabStrip"].includes(c.type)) set(VB_BACKSTYLE, c.backStyle === undefined ? undefined : c.backStyle !== 0);
  if (["CheckBox", "ComboBox", "CommandButton", "Label", "OptionButton", "TextBox", "ToggleButton"].includes(c.type)) set(VB_WORDWRAP, c.wordWrap);
  if (["CheckBox", "ComboBox", "CommandButton", "Label", "OptionButton", "TextBox", "ToggleButton"].includes(c.type)) set(VB_AUTOSIZE, c.autoSize);
  if (c.type === "TextBox") { set(VB_MULTILINE, c.multiLine); v |= VB_EDITABLE; }
  if (c.type === "ComboBox") set(VB_EDITABLE, c.style !== 2);
  return v >>> 0;
}

function fontToTextProps(f: VbaFont | undefined, align: number | undefined, raw: Values | undefined): Values {
  const t: Values = new Map(raw ?? []);
  const font = f ?? { name: "Tahoma", size: 8, bold: false, italic: false, underline: false, strikethrough: false };
  t.set("FontName", font.name);
  const effects = ((raw?.get("FontEffects") as number) ?? 0) & ~0xf | (font.bold ? 1 : 0) | (font.italic ? 2 : 0) | (font.underline ? 4 : 0) | (font.strikethrough ? 8 : 0);
  if (effects) t.set("FontEffects", effects >>> 0); else t.delete("FontEffects");
  const h = Math.round(font.size * 20);
  if (h !== 160) t.set("FontHeight", h); else t.delete("FontHeight");
  if (font.bold) t.set("FontWeight", 700); else t.delete("FontWeight");
  const pa = align === 2 ? 3 : align === 3 ? 2 : align === 1 ? 1 : (raw?.get("ParagraphAlign") as number | undefined);
  if (pa && pa !== 1) t.set("ParagraphAlign", pa); else t.delete("ParagraphAlign");
  return t;
}
function textPropsToFont(t: Values): { font: VbaFont; align?: number } {
  const eff = (t.get("FontEffects") as number) ?? 0;
  const weight = (t.get("FontWeight") as number) ?? 400;
  const pa = t.get("ParagraphAlign") as number | undefined;
  return {
    font: { name: (t.get("FontName") as string) ?? "MS Sans Serif", size: ((t.get("FontHeight") as number) ?? 160) / 20, bold: (eff & 1) !== 0 || weight >= 600, italic: (eff & 2) !== 0, underline: (eff & 4) !== 0, strikethrough: (eff & 8) !== 0 },
    align: pa === 3 ? 2 : pa === 2 ? 3 : pa === 1 ? 1 : undefined,
  };
}

interface ControlRaw { id: number; values: Values; site: Values; text?: Values; tail?: Uint8Array; storage?: Map<string, Uint8Array>; clsids?: Map<string, string> }

/** The values of a simple control's record, from its design (unknown properties kept from what was read). */
function controlValues(c: VbaControlDesign, raw: ControlRaw | undefined): Values {
  const t = c.type;
  const v: Values = new Map(raw?.values ?? []);
  const put = (k: string, x: unknown, def: unknown) => { if (x === undefined || x === def) v.delete(k); else v.set(k, x); };
  put("ForeColor", c.foreColor === undefined ? undefined : c.foreColor >>> 0, DEF_FORE[t]);
  put("BackColor", c.backColor === undefined ? undefined : c.backColor >>> 0, DEF_BACK[t]);
  put("VariousPropertyBits", various(c, (raw?.values.get("VariousPropertyBits") as number) ?? DEF_VARIOUS(t)), DEF_VARIOUS(t));
  v.set("Size", [toHm(c.width), toHm(c.height)]);
  if (t === "CommandButton" || t === "Label" || t === "CheckBox" || t === "OptionButton" || t === "ToggleButton") put("Caption", c.caption ?? "", "");
  if (t === "CommandButton" || t === "Label" || isMorph(t)) put("Accelerator", c.accelerator ? c.accelerator.charCodeAt(0) : undefined, 0);
  if (t === "Label") { put("BorderStyle", c.borderStyle, 0); put("SpecialEffect", c.specialEffect, 0); put("BorderColor", c.borderColor === undefined ? undefined : c.borderColor >>> 0, 0x80000006); }
  if (t === "Image") { put("BorderStyle", c.borderStyle, 1); put("SpecialEffect", c.specialEffect, 0); put("BorderColor", c.borderColor === undefined ? undefined : c.borderColor >>> 0, 0x80000006); put("AutoSize", c.autoSize ? true : undefined, undefined); }
  if (isMorph(t)) {
    put("DisplayStyle", t === "ComboBox" && c.style === 2 ? 7 : MORPH_STYLE[t], 1);
    put("MaxLength", c.maxLength, 0);
    put("PasswordChar", c.passwordChar ? c.passwordChar.charCodeAt(0) : undefined, 0);
    put("BorderStyle", c.borderStyle, 0);
    put("SpecialEffect", c.specialEffect, 2);
    put("BorderColor", c.borderColor === undefined ? undefined : c.borderColor >>> 0, 0x80000006);
    if (t === "ListBox" || t === "ComboBox") put("ColumnCount", c.columnCount, 1);
    if (t === "ListBox") put("MultiSelect", c.multiSelect, 0);
    if (t === "CheckBox" || t === "OptionButton") put("GroupName", c.groupName ?? "", "");
    let value: string | undefined;
    if (t === "TextBox" || t === "ComboBox") value = c.value === null || c.value === undefined ? "" : String(c.value);
    else if (t === "CheckBox" || t === "OptionButton" || t === "ToggleButton") value = c.value === true ? "1" : c.value === false || c.value === undefined ? "0" : undefined;
    else value = "";
    if (value === undefined) v.delete("Value"); else if (value === "" && !(t === "CheckBox" || t === "OptionButton" || t === "ToggleButton")) v.delete("Value"); else v.set("Value", value);
    v.set("Reserved", true);
  }
  if (t === "ScrollBar" || t === "SpinButton") {
    put("Min", c.min, 0);
    put("Max", c.max, t === "ScrollBar" ? 32767 : 100);
    put("Position", typeof c.value === "number" ? c.value : undefined, 0);
    put("SmallChange", c.smallChange, 1);
    if (t === "ScrollBar") put("LargeChange", c.largeChange, 1);
    put("Orientation", c.orientation, -1);
  }
  return v;
}

/** A control's design from its record values. */
function controlFromValues(type: VbaControlType, name: string, v: Values, text: Values | undefined): VbaControlDesign {
  const c = newControl(type, name);
  const g = <T>(k: string) => v.get(k) as T | undefined;
  c.foreColor = g<number>("ForeColor") ?? DEF_FORE[type];
  c.backColor = g<number>("BackColor") ?? DEF_BACK[type];
  const size = g<[number, number]>("Size");
  if (size) { c.width = toPt(size[0]); c.height = toPt(size[1]); }
  const vb = g<number>("VariousPropertyBits") ?? DEF_VARIOUS(type);
  c.enabled = (vb & VB_ENABLED) !== 0;
  if (["CheckBox", "ComboBox", "CommandButton", "ListBox", "OptionButton", "TextBox", "ToggleButton"].includes(type)) c.locked = (vb & VB_LOCKED) !== 0;
  if (!["ListBox", "ScrollBar", "SpinButton", "TabStrip"].includes(type)) c.backStyle = vb & VB_BACKSTYLE ? 1 : 0;
  if (["CheckBox", "ComboBox", "CommandButton", "Label", "OptionButton", "TextBox", "ToggleButton"].includes(type)) { c.wordWrap = (vb & VB_WORDWRAP) !== 0; c.autoSize = (vb & VB_AUTOSIZE) !== 0; }
  if (type === "TextBox") c.multiLine = (vb & VB_MULTILINE) !== 0;
  if (v.has("Caption")) c.caption = g<string>("Caption");
  else if ("caption" in c) c.caption = "";
  const acc = g<number>("Accelerator");
  if (acc) c.accelerator = String.fromCharCode(acc);
  if (type === "Label" || type === "Image" || isMorph(type)) {
    c.borderStyle = g<number>("BorderStyle") ?? (type === "Image" ? 1 : 0);
    c.specialEffect = g<number>("SpecialEffect") ?? (isMorph(type) ? 2 : 0);
    if (v.has("BorderColor")) c.borderColor = g<number>("BorderColor");
  }
  if (type === "Image" && v.has("AutoSize")) c.autoSize = true;
  if (isMorph(type)) {
    if (type === "ComboBox") c.style = g<number>("DisplayStyle") === 7 ? 2 : 0;
    if (v.has("MaxLength")) c.maxLength = g<number>("MaxLength");
    const pc = g<number>("PasswordChar");
    if (pc) c.passwordChar = String.fromCharCode(pc);
    if (type === "ListBox" || type === "ComboBox") c.columnCount = g<number>("ColumnCount") ?? 1;
    if (type === "ListBox") c.multiSelect = g<number>("MultiSelect") ?? 0;
    if (type === "CheckBox" || type === "OptionButton") c.groupName = g<string>("GroupName") ?? "";
    const val = g<string>("Value") ?? "";
    if (type === "TextBox" || type === "ComboBox") c.value = val;
    else if (type === "CheckBox" || type === "OptionButton" || type === "ToggleButton") c.value = val === "1" ? true : val === "0" || val === "" ? false : null;
  }
  if (type === "ScrollBar" || type === "SpinButton") {
    c.min = g<number>("Min") ?? 0;
    c.max = g<number>("Max") ?? (type === "ScrollBar" ? 32767 : 100);
    c.value = g<number>("Position") ?? 0;
    c.smallChange = g<number>("SmallChange") ?? 1;
    if (type === "ScrollBar") c.largeChange = g<number>("LargeChange") ?? 1;
    c.orientation = g<number>("Orientation") ?? -1;
  }
  if (text) { const tp = textPropsToFont(text); c.font = tp.font; if (tp.align) c.textAlign = tp.align; }
  return c;
}

// ----- the form control (f stream) ------------------------------------------------------------------------------

interface SiteEntry { depth: number; type: number; values: Values; bytes: Uint8Array }
interface FormRecord { values: Values; classTable: Uint8Array | null; classCount: number; sites: SiteEntry[]; tail: Uint8Array }

function readFormControl(b: Uint8Array): FormRecord {
  if (b[0] !== 0 || b[1] !== 4) throw new Error("not a form control");
  const v = rd(b);
  const r = readRecord(b, 0, FORM_FIELDS, 4);
  let p = r.next;
  const flags = (r.values.get("BooleanProperties") as number) ?? 4;
  let classTable: Uint8Array | null = null, classCount = 0;
  if (!(flags & 0x8000)) {
    classCount = v.getUint16(p, true);
    const start = p + 2;
    p = start;
    for (let i = 0; i < classCount; i++) p += 4 + v.getUint16(p + 2, true);
    classTable = b.slice(start, p);
    p = start + classTable.length;
  }
  const count = v.getUint32(p, true), cbytes = v.getUint32(p + 4, true);
  p += 8;
  const start = p;
  const types: { depth: number; type: number }[] = [];
  while (types.length < count) {
    const depth = b[p], tc = b[p + 1];
    p += 2;
    if (tc & 0x80) { const n = tc & 0x7f, type = b[p++]; for (let i = 0; i < n; i++) types.push({ depth, type }); }
    else types.push({ depth, type: tc });
  }
  while ((p - start) % 4) p++;
  const sites: SiteEntry[] = [];
  for (const t of types) {
    const s = readRecord(b, p, SITE_FIELDS, 4);
    sites.push({ ...t, values: s.values, bytes: b.slice(p, s.next) });
    p = s.next;
  }
  p = start + cbytes;
  return { values: r.values, classTable, classCount, sites, tail: b.slice(p) };
}

function writeFormControl(values: Values, sites: { depth: number; bytes: number[] | Uint8Array }[], classTable: Uint8Array | null, classCount: number, tail: Uint8Array | null): Uint8Array {
  const o = new Out();
  o.bytes(writeRecord(values, FORM_FIELDS, 4, 0, 4));
  const flags = (values.get("BooleanProperties") as number) ?? 4;
  if (!(flags & 0x8000)) { o.u16(classTable ? classCount : 0); if (classTable) o.bytes(classTable); }
  const body = new Out();
  for (const s of sites) { body.u8(s.depth); body.u8(1); }
  body.align(4);
  for (const s of sites) body.bytes(s.bytes);
  o.u32(sites.length);
  o.u32(body.len);
  o.bytes(body.b);
  if (tail && flags & 0x4000) o.bytes(tail);
  return Uint8Array.from(o.b);
}

/** CompObj ([MS-OLEDS] 2.3.8): the header with the class, the user type, the clipboard format and the ProgID. */
function compObj(clsid: string, userType: string, progId: string): Uint8Array {
  const o = new Out();
  o.u32(0xfffe0001); o.u32(0x00000a03); o.u32(0xffffffff); o.bytes(guidBytes(clsid));
  const ansi = (s: string) => { o.u32(s.length + 1); o.bytes(Array.from(s, (c) => c.charCodeAt(0))); o.u8(0); };
  ansi(userType); ansi("Embedded Object"); ansi(progId);
  o.u32(0x71b239f4); o.u32(0); o.u32(0); o.u32(0);
  return Uint8Array.from(o.b);
}

// ----- reading a form storage ------------------------------------------------------------------------------------

/** What we keep of a form to write it back (VbaFormDesign.raw). */
interface FormRaw { values: Values; classTable: Uint8Array | null; classCount: number; tail: Uint8Array; opaque: { index: number; site: SiteEntry; data: Uint8Array | null; storage: Map<string, Uint8Array> | null; clsids: Map<string, string> | null }[]; vbframe: string[]; snapshot: string; streams: Map<string, Uint8Array>; clsids: Map<string, string> }

/** The design as a comparable string (without what the reader keeps). */
export function designSnapshot(d: VbaFormDesign | VbaControlDesign): string {
  return JSON.stringify(d, (k, v) => (k === "raw" || k === "list" || k === "listIndex" || k === "selected" ? undefined : v));
}

const sub = (streams: Map<string, Uint8Array>, prefix: string) => {
  const out = new Map<string, Uint8Array>();
  for (const [k, v] of streams) if (k.startsWith(prefix + "/")) out.set(k.slice(prefix.length + 1), v);
  return out;
};
const subIds = (ids: Map<string, string>, prefix: string) => {
  const out = new Map<string, string>();
  for (const [k, v] of ids) if (k === prefix) out.set("", v); else if (k.startsWith(prefix + "/")) out.set(k.slice(prefix.length + 1), v);
  return out;
};

/** The controls of a container (a form's or a Frame's "f" and "o"). */
function readContainer(streams: Map<string, Uint8Array>, clsids: Map<string, string>): { rec: FormRecord; controls: VbaControlDesign[]; opaque: FormRaw["opaque"] } {
  const rec = readFormControl(streams.get("f")!);
  const o = streams.get("o") ?? new Uint8Array();
  let op = 0;
  const controls: VbaControlDesign[] = [];
  const opaque: FormRaw["opaque"] = [];
  rec.sites.forEach((s, index) => {
    const sv = s.values;
    const id = (sv.get("ID") as number) ?? 0;
    const flags = (sv.get("BitFlags") as number) ?? SITE_FLAGS_DEFAULT;
    const streamed = (flags & 0x10) !== 0;
    const size = (sv.get("ObjectStreamSize") as number) ?? 0;
    const data = streamed ? o.slice(op, op + size) : null;
    if (streamed) op += size;
    const cache = (sv.get("ClsidCacheIndex") as number) ?? 0x7fff;
    const store = `i${String(id).padStart(2, "0")}`;
    const name = (sv.get("Name") as string) ?? "";
    let type = BY_CACHE.get(cache);
    let c: VbaControlDesign | null = null;
    try {
      if (s.depth === 0 && type === "Frame" && !streamed) {
        const inner = sub(streams, store);
        const r = readContainer(inner, subIds(clsids, store));
        c = newControl("Frame", name);
        frameFromValues(c, r.rec.values);
        c.controls = r.controls;
        c.raw = { id, values: r.rec.values, site: sv, storage: inner, clsids: subIds(clsids, store), frame: { classTable: r.rec.classTable, classCount: r.rec.classCount, tail: r.rec.tail, opaque: r.opaque } } as unknown as ControlRaw as unknown as VbaControlDesign["raw"];
      } else if (s.depth === 0 && type && data && fieldsOf(type)) {
        const def = fieldsOf(type)!;
        const r = readRecord(data, 0, def.fields, def.mask);
        if (cache === 15) type = STYLE_TYPE[(r.values.get("DisplayStyle") as number) ?? 1] ?? "TextBox";
        if (isMorph(type)) type = STYLE_TYPE[(r.values.get("DisplayStyle") as number) ?? 1] ?? type;
        let text: Values | undefined, next = r.next;
        if (def.text && next < data.length) { const t = readRecord(data, next, TEXTPROPS_FIELDS, 4); text = t.values; next = t.next; }
        c = controlFromValues(type, name, r.values, text);
        c.raw = { id, values: r.values, site: sv, text, tail: next < data.length ? data.slice(next) : undefined } as unknown as VbaControlDesign["raw"];
      }
    } catch { c = null; }
    if (!c) {
      opaque.push({ index, site: s, data, storage: streamed ? null : sub(streams, store), clsids: streamed ? null : subIds(clsids, store) });
      return;
    }
    const pos = sv.get("Position") as [number, number] | undefined;
    c.left = pos ? toPt(pos[0]) : 0;
    c.top = pos ? toPt(pos[1]) : 0;
    c.visible = (flags & 2) !== 0;
    if (c.type !== "Label") c.tabStop = (flags & 1) !== 0;
    c.tabIndex = (sv.get("TabIndex") as number) ?? controls.length;
    if (sv.has("Tag")) c.tag = sv.get("Tag") as string;
    if (sv.has("ControlTipText")) c.controlTipText = sv.get("ControlTipText") as string;
    controls.push(c);
  });
  return { rec, controls, opaque };
}

function frameFromValues(c: VbaControlDesign, v: Values): void {
  const size = v.get("DisplayedSize") as [number, number] | undefined;
  if (size) { c.width = toPt(size[0]); c.height = toPt(size[1]); }
  c.caption = (v.get("Caption") as string) ?? "";
  c.backColor = (v.get("BackColor") as number) ?? 0x8000000f;
  c.foreColor = (v.get("ForeColor") as number) ?? 0x80000012;
  c.specialEffect = (v.get("SpecialEffect") as number) ?? 0;
  c.borderStyle = (v.get("BorderStyle") as number) ?? 0;
  c.enabled = (((v.get("BooleanProperties") as number) ?? 4) & 4) !== 0;
  const f = v.get("Font") as { std?: StdFont; textProps?: Values } | undefined;
  if (f) c.font = fontOf(f);
}
function fontOf(f: { std?: StdFont; textProps?: Values }): VbaFont {
  if (f.std) return { name: f.std.face, size: f.std.height / 10000, bold: f.std.weight >= 600, italic: (f.std.flags & 2) !== 0, underline: (f.std.flags & 4) !== 0, strikethrough: (f.std.flags & 8) !== 0 };
  return textPropsToFont(f.textProps!).font;
}
function stdFont(f: VbaFont | undefined, raw: { std?: StdFont } | undefined): { std: StdFont } {
  const font = f ?? { name: "Tahoma", size: 8, bold: false, italic: false, underline: false, strikethrough: false };
  return { std: { charset: raw?.std?.charset ?? 0, flags: (font.italic ? 2 : 0) | (font.underline ? 4 : 0) | (font.strikethrough ? 8 : 0), weight: font.bold ? 700 : 400, height: Math.round(font.size * 10000), face: font.name } };
}

/** A form from the streams of its designer storage (paths relative to it: "f", "o", "i01/f" …). */
export function readFormStorage(name: string, streams: Map<string, Uint8Array>, clsids: Map<string, string> = new Map()): VbaFormDesign {
  const d = newForm(name);
  const r = readContainer(streams, clsids);
  const v = r.rec.values;
  const size = v.get("DisplayedSize") as [number, number] | undefined;
  d.width = size ? toPt(size[0]) : toPt(4000);
  d.height = size ? toPt(size[1]) : toPt(3000);
  d.caption = (v.get("Caption") as string) ?? "";
  d.backColor = (v.get("BackColor") as number) ?? 0x8000000f;
  d.foreColor = (v.get("ForeColor") as number) ?? 0x80000012;
  const f = v.get("Font") as { std?: StdFont; textProps?: Values } | undefined;
  if (f) d.font = fontOf(f);
  d.controls = r.controls;
  // VBFrame: Caption, ClientWidth/Height (twips), StartUpPosition; other lines kept.
  const vbf = streams.get("\u0003VBFrame");
  const lines = vbf ? Array.from(vbf, (c) => String.fromCharCode(c)).join("").split(/\r?\n/) : [];
  for (const l of lines) {
    const m = /^\s*(\w+)\s*=\s*(.*?)\s*$/.exec(l);
    if (!m) continue;
    const val = m[2].replace(/\s*'.*$/, "");
    if (m[1] === "Caption") d.caption = val.replace(/^"|"$/g, "").replace(/""/g, '"');
    else if (m[1] === "StartUpPosition") d.startUpPosition = Number(val);
    else if (m[1] === "ClientWidth" && !size) d.width = Number(val) / 20;
    else if (m[1] === "ClientHeight" && !size) d.height = Number(val) / 20;
  }
  const raw: FormRaw = { values: v, classTable: r.rec.classTable, classCount: r.rec.classCount, tail: r.rec.tail, opaque: r.opaque, vbframe: lines, snapshot: "", streams, clsids };
  d.raw = raw as unknown as VbaFormDesign["raw"];
  raw.snapshot = designSnapshot(d);
  return d;
}

// ----- writing a form storage ------------------------------------------------------------------------------------

/** Ids: those read stay, new controls get the next free ones. */
function assignIds(d: VbaFormDesign): Map<VbaControlDesign, number> {
  const ids = new Map<VbaControlDesign, number>();
  const used = new Set<number>();
  const all: VbaControlDesign[] = [];
  const walk = (l: VbaControlDesign[]) => { for (const c of l) { all.push(c); if (c.controls) walk(c.controls); } };
  walk(d.controls);
  for (const c of all) { const id = (c.raw as unknown as ControlRaw | undefined)?.id; if (id && !used.has(id)) { ids.set(c, id); used.add(id); } }
  const fr = d.raw as unknown as FormRaw | undefined;
  for (const o of fr?.opaque ?? []) used.add((o.site.values.get("ID") as number) ?? 0);
  let next = 1;
  for (const c of all) if (!ids.has(c)) { while (used.has(next)) next++; ids.set(c, next); used.add(next); }
  return ids;
}

function siteValues(c: VbaControlDesign, id: number, tabIndex: number, objectSize: number, streamed: boolean, raw: Values | undefined): Values {
  const v: Values = new Map(raw ?? []);
  v.set("Name", c.name);
  v.set("ID", id);
  if (c.tag) v.set("Tag", c.tag); else v.delete("Tag");
  if (c.controlTipText) v.set("ControlTipText", c.controlTipText); else v.delete("ControlTipText");
  let flags = (raw?.get("BitFlags") as number | undefined) ?? SITE_FLAGS_DEFAULT;
  flags = c.visible === false ? flags & ~2 : flags | 2;
  if (c.type !== "Label") flags = c.tabStop === false ? flags & ~1 : flags | 1;
  flags = streamed ? flags | 0x10 : flags & ~0x10;
  if (c.type === "Frame") flags |= 0x40000;
  if (flags === SITE_FLAGS_DEFAULT) v.delete("BitFlags"); else v.set("BitFlags", flags >>> 0);
  if (streamed) v.set("ObjectStreamSize", objectSize); else v.delete("ObjectStreamSize");
  v.set("TabIndex", tabIndex);
  v.set("ClsidCacheIndex", CACHE[c.type]!);
  v.set("Position", [toHm(c.left), toHm(c.top)]);
  return v;
}

/** A container's "f", "o" and child storages (the form itself, or a Frame). */
function writeContainer(values: Values, controls: VbaControlDesign[], ids: Map<VbaControlDesign, number>, keep: { classTable: Uint8Array | null; classCount: number; tail: Uint8Array | null; opaque: FormRaw["opaque"] }, out: Map<string, Uint8Array>, clsids: Map<string, string>, prefix: string): void {
  const o = new Out();
  const sites: { depth: number; bytes: number[] | Uint8Array }[] = [];
  const opaque = [...keep.opaque];
  const order = [...controls].sort((a, b) => (a.tabIndex ?? 0) - (b.tabIndex ?? 0));
  const tabOf = new Map(order.map((c, i) => [c, i]));
  let index = 0;
  const flushOpaque = () => {
    while (opaque.length && opaque[0].index <= index) {
      const x = opaque.shift()!;
      sites.push({ depth: x.site.depth, bytes: x.site.bytes });
      if (x.data) o.bytes(x.data);
      if (x.storage) {
        const store = `i${String((x.site.values.get("ID") as number) ?? 0).padStart(2, "0")}`;
        for (const [k, v] of x.storage) out.set(`${prefix}${store}/${k}`, v);
        for (const [k, v] of x.clsids ?? []) clsids.set(`${prefix}${store}${k ? "/" + k : ""}`, v);
      }
      index++;
    }
  };
  for (const c of controls) {
    flushOpaque();
    const raw = c.raw as unknown as (ControlRaw & { frame?: { classTable: Uint8Array | null; classCount: number; tail: Uint8Array; opaque: FormRaw["opaque"] } }) | undefined;
    const id = ids.get(c)!;
    if (c.type === "Frame") {
      const store = `i${String(id).padStart(2, "0")}`;
      writeContainer(frameValues(c, raw?.values), c.controls ?? [], ids, { classTable: raw?.frame?.classTable ?? null, classCount: raw?.frame?.classCount ?? 0, tail: raw?.frame?.tail ?? null, opaque: raw?.frame?.opaque ?? [] }, out, clsids, `${prefix}${store}/`);
      out.set(`${prefix}${store}/\u0001CompObj`, compObj(CLSID_FRAME, "Microsoft Forms 2.0 Frame", "Forms.Frame.1"));
      clsids.set(`${prefix}${store}`, CLSID_FRAME);
      sites.push({ depth: 0, bytes: writeRecord(siteValues(c, id, tabOf.get(c)!, 0, false, raw?.site), SITE_FIELDS, 4, 0, 0) });
    } else {
      const def = fieldsOf(c.type);
      if (!def) throw new Error(`cannot write a ${c.type} yet`);
      const data = writeRecord(controlValues(c, raw), def.fields, def.mask, 0, 2);
      if (def.text) data.push(...writeRecord(fontToTextProps(c.font, c.textAlign, raw?.text), TEXTPROPS_FIELDS, 4, 0, 2));
      if (raw?.tail && isMorph(c.type) && raw.values.has("cColumnInfo")) data.push(...raw.tail);
      o.bytes(data);
      sites.push({ depth: 0, bytes: writeRecord(siteValues(c, id, tabOf.get(c)!, data.length, true, raw?.site), SITE_FIELDS, 4, 0, 0) });
    }
    index++;
  }
  index = Number.MAX_SAFE_INTEGER;
  flushOpaque();
  const v: Values = new Map(values);
  const maxId = Math.max(0, ...[...ids.values()], ...keep.opaque.map((x) => (x.site.values.get("ID") as number) ?? 0));
  if (maxId) v.set("NextAvailableID", maxId); else v.delete("NextAvailableID");
  out.set(`${prefix}f`, writeFormControl(v, sites, keep.classTable, keep.classCount, keep.tail));
  out.set(`${prefix}o`, Uint8Array.from(o.b));
}

function frameValues(c: VbaControlDesign, raw: Values | undefined): Values {
  const v: Values = new Map(raw ?? []);
  const put = (k: string, x: unknown, def: unknown) => { if (x === undefined || x === def) v.delete(k); else v.set(k, x); };
  put("BackColor", c.backColor === undefined ? undefined : c.backColor >>> 0, 0x8000000f);
  put("ForeColor", c.foreColor === undefined ? undefined : c.foreColor >>> 0, 0x80000012);
  put("Caption", c.caption ?? "", "");
  // A new Frame has Office's etched border.
  put("SpecialEffect", c.specialEffect ?? (raw ? undefined : 3), 0);
  put("BorderStyle", c.borderStyle, 0);
  const flags = (((raw?.get("BooleanProperties") as number) ?? 4) & ~4) | (c.enabled === false ? 0 : 4);
  put("BooleanProperties", flags, 4);
  v.set("DisplayedSize", [toHm(c.width), toHm(c.height)]);
  v.set("LogicalSize", [toHm(c.width), toHm(c.height)]);
  v.set("Font", stdFont(c.font, v.get("Font") as { std?: StdFont } | undefined));
  v.set("DrawBuffer", (raw?.get("DrawBuffer") as number) ?? 32000);
  return v;
}

/** The designer storage of a form: streams by path relative to it, and the CLSIDs of its storages ("" itself). */
export function writeFormStorage(d: VbaFormDesign): { streams: Map<string, Uint8Array>; clsids: Map<string, string> } {
  const raw = d.raw as unknown as FormRaw | undefined;
  if (raw && raw.snapshot === designSnapshot(d)) return { streams: new Map(raw.streams), clsids: new Map(raw.clsids.size ? raw.clsids : [["", CLSID_FORM]]) };
  const out = new Map<string, Uint8Array>();
  const clsids = new Map<string, string>([["", CLSID_FORM]]);
  const v: Values = new Map(raw?.values ?? []);
  const put = (k: string, x: unknown, def: unknown) => { if (x === undefined || x === def) v.delete(k); else v.set(k, x); };
  put("BackColor", d.backColor === undefined ? undefined : d.backColor >>> 0, 0x8000000f);
  put("ForeColor", d.foreColor === undefined ? undefined : d.foreColor >>> 0, 0x80000012);
  put("Caption", d.caption, "");
  v.set("DisplayedSize", [toHm(d.width), toHm(d.height)]);
  if (!raw) v.set("LogicalSize", [toHm(d.width), toHm(d.height)]);
  v.set("Font", stdFont(d.font, v.get("Font") as { std?: StdFont } | undefined));
  v.set("DrawBuffer", (raw?.values.get("DrawBuffer") as number) ?? 32000);
  writeContainer(v, d.controls, assignIds(d), { classTable: raw?.classTable ?? null, classCount: raw?.classCount ?? 0, tail: raw?.tail ?? null, opaque: raw?.opaque ?? [] }, out, clsids, "");
  out.set("\u0001CompObj", compObj(CLSID_FORM, "Microsoft Forms 2.0 Form", "Forms.Form.1"));
  out.set("\u0003VBFrame", vbFrame(d, raw?.vbframe));
  return { streams: out, clsids };
}

/** The VBFrame text ([MS-OVBA] 2.3.5): our values for Caption, size and StartUpPosition, other lines kept. */
function vbFrame(d: VbaFormDesign, old: string[] | undefined): Uint8Array {
  const pad = (k: string) => `   ${k.padEnd(16)}=   `;
  const ours: [string, string][] = [
    ["Caption", `"${d.caption.slice(0, 130).replace(/"/g, '""')}"`],
    ["ClientHeight", String(Math.round(d.height * 20))],
    ["ClientLeft", "120"],
    ["ClientTop", "465"],
    ["ClientWidth", String(Math.round(d.width * 20))],
    ["StartUpPosition", `${d.startUpPosition ?? 1}  '${["Manual", "CenterOwner", "CenterScreen", "WindowsDefault"][d.startUpPosition ?? 1] ?? ""}`],
  ];
  const kept = new Map<string, string>();
  for (const l of old ?? []) { const m = /^\s*(\w+)\s*=\s*(.*)$/.exec(l); if (m && !/^(Caption|ClientHeight|ClientWidth|StartUpPosition)$/.test(m[1])) kept.set(m[1], m[2]); }
  const lines = ["VERSION 5.00", `Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ${d.name} `];
  for (const [k, v] of ours) lines.push(pad(k) + (kept.has(k) ? kept.get(k)! : v));
  for (const [k, v] of kept) if (!ours.some(([o]) => o === k)) lines.push(pad(k) + v);
  lines.push("End", "");
  const s = lines.join("\r\n");
  return Uint8Array.from(Array.from(s, (c) => (c.charCodeAt(0) < 256 ? c.charCodeAt(0) : 0x3f)));
}

/** A form whose storage we could not read: an empty design that writes the storage back as it was read. */
export function opaqueForm(name: string, streams: Map<string, Uint8Array>, clsids: Map<string, string>): VbaFormDesign {
  const d = newForm(name);
  const raw: FormRaw = { values: new Map(), classTable: null, classCount: 0, tail: new Uint8Array(), opaque: [], vbframe: [], snapshot: "", streams, clsids };
  d.raw = raw as unknown as VbaFormDesign["raw"];
  raw.snapshot = designSnapshot(d);
  return d;
}
