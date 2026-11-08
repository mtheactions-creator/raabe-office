// Makros in RaabeCalc (D-080, D-091): the workbook's VBA project (xl/vbaProject.bin of an .xlsm) read and written with
// @raabe/vba and run by its interpreter against an Excel-like object model: Application, Workbook, Worksheets and Range
// with the properties and methods school macros use, named arguments (Key1:=…) as Excel names them. As in Excel,
// macros of an opened file are off until "Inhalt aktivieren"; Ansicht > Makros (Alt+F8) lists them. Macros reach files
// only after asking (D-080). One run is one undo step; formulas a macro writes are worked out at once.

import { h, showDialog } from "@raabe/shell";
import { CellError, colName, isDateFormat, parseRange, toFileFormula, toLocalFormula, type Cell, type CellStyle, type Rect, type Sheet, type Value } from "@raabe/sheet-engine";
import { NOT_A_METHOD, VbaError, VbaRuntime, readProject, vbaDateToSerial as serialOfDate, vbaSerialToDate as dateOfSerial, vbaToNumber, vbaToString, writeProject, type VbaArray, type VbaHost, type VbaModule, type VbaObject, type VbaValue } from "@raabe/vba";
import * as A from "./actions";
import type { CalcApp } from "./app";
import { macroApp, type MacroApp } from "./macro-app";

/** Excel's constants that school macros use (xlUp …), as the object model's globals. */
export const XL: Record<string, number> = {
  xlup: -4162, xldown: -4121, xltoleft: -4159, xltoright: -4161, xlcelltypeblanks: 4, xlcelltypeconstants: 2, xlcelltypeformulas: -4123,
  xlcelltypelastcell: 11, xlcelltypevisible: 12, xlascending: 1, xldescending: 2, xlyes: 1, xlno: 2, xlguess: 0, xlwhole: 1, xlpart: 2,
  xlvalues: -4163, xlformulas: -4123, xlcenter: -4108, xlleft: -4131, xlright: -4152, xltop: -4160, xlbottom: -4107, xlgeneral: 1, xljustify: -4130,
  xlshiftup: -4162, xlshiftdown: -4121, xlshifttoleft: -4159, xlshifttoright: -4161, xlsheetvisible: -1, xlsheethidden: 0, xlsheetveryhidden: 2,
  xlnone: -4142, xlcontinuous: 1, xldash: -4115, xldot: -4118, xldouble: -4119, xlhairline: 1, xlthin: 2, xlmedium: -4138, xlthick: 4,
  xledgeleft: 7, xledgetop: 8, xledgebottom: 9, xledgeright: 10, xlinsidevertical: 11, xlinsidehorizontal: 12, xlcalculationmanual: -4135,
  xlcalculationautomatic: -4105, xlpastevalues: -4163, xlpasteformats: -4122, xlpasteall: -4104, xlpasteformulas: -4123,
  xlopenxmlworkbook: 51, xlopenxmlworkbookmacroenabled: 52, xlbyrows: 1, xlbycolumns: 2, xlnext: 1, xlprevious: 2, xlcolorindexnone: -4142,
  xlsolid: 1, xlunderlinestylesingle: 2, xlunderlinestylenone: -4142, xlautomatic: -4105, xlfillcopy: 1, xlfillseries: 2, xlfilldefault: 0,
};

/** Excel's 56-colour palette for ColorIndex (1-based). */
const PALETTE = ["000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "800000", "008000", "000080", "808000", "800080",
  "008080", "C0C0C0", "808080", "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", "000080", "FF00FF", "FFFF00",
  "00FFFF", "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99", "3366FF",
  "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333"];

/** Excel's parameter names of the methods (so that Key1:=…, After:=… reach their place). */
const ARGS: Record<string, string[]> = {
  add: ["Before", "After", "Count", "Type"], sort: ["Key1", "Order1", "Key2", "Type", "Order2", "Key3", "Order3", "Header", "OrderCustom", "MatchCase", "Orientation"],
  copy: ["Destination"], cut: ["Destination"], find: ["What", "After", "LookIn", "LookAt", "SearchOrder", "SearchDirection", "MatchCase"],
  pastespecial: ["Paste", "Operation", "SkipBlanks", "Transpose"], delete: ["Shift"], insert: ["Shift", "CopyOrigin"], specialcells: ["Type", "Value"],
  autofill: ["Destination", "Type"], borderaround: ["LineStyle", "Weight", "ColorIndex", "Color"], protect: ["Password"], unprotect: ["Password"],
  inputbox: ["Prompt", "Title", "Default", "Left", "Top", "HelpFile", "HelpContextID", "Type"], address: ["RowAbsolute", "ColumnAbsolute", "ReferenceStyle", "External", "RelativeTo"],
  offset: ["RowOffset", "ColumnOffset"], resize: ["RowSize", "ColumnSize"], move: ["Before", "After"], autofilter: ["Field", "Criteria1", "Operator", "Criteria2"],
  removeduplicates: ["Columns", "Header"], saveas: ["Filename", "FileFormat"], cells: ["RowIndex", "ColumnIndex"], replace: ["What", "Replacement", "LookAt", "SearchOrder", "MatchCase"],
};
/** Sheets' Copy and Move take Before/After. */
const SHEET_ARGS: Record<string, string[]> = { copy: ["Before", "After"], move: ["Before", "After"] };

const enabled = new WeakSet<object>();

type Get = (args: VbaValue[]) => VbaValue;
type Call = (args: VbaValue[]) => VbaValue | Promise<VbaValue>;

function obj(props: Record<string, Get>, setters: Record<string, (v: VbaValue, args: VbaValue[]) => void> = {}, methods: Record<string, Call> = {}, extra: Partial<VbaObject> = {}, names: Record<string, string[]> = ARGS): VbaObject {
  const low = <T>(r: Record<string, T>) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k.toLowerCase(), v]));
  const P = low(props), SET = low(setters), M = low(methods);
  return {
    kind: "object",
    get: (name, args) => {
      const g = P[name.toLowerCase()];
      if (g) return g(args);
      const m = M[name.toLowerCase()];
      if (m) return m(args) as VbaValue;
      throw new VbaError(`Objekt unterstützt diese Eigenschaft oder Methode nicht: ${name}`, 438);
    },
    set: (name, v, args) => {
      const s = SET[name.toLowerCase()];
      if (!s) throw new VbaError(`Objekt unterstützt diese Eigenschaft oder Methode nicht: ${name}`, 438);
      s(v, args ?? []);
    },
    // A property with arguments (End(xlUp), Cells(2, 3), Offset(1)) is Excel's own parameterised property, not an
    // indexed collection: answered here, before the interpreter would index the property's value.
    call: (name, args) => {
      const m = M[name.toLowerCase()];
      if (m) return m(args) as VbaValue;
      const k = name.toLowerCase();
      return k in P && args.length ? P[k](args) : NOT_A_METHOD;
    },
    argNames: (name) => names[name.toLowerCase()] ?? ARGS[name.toLowerCase()],
    ...extra,
  };
}

const isObj = (v: VbaValue): v is VbaObject => !!v && typeof v === "object" && (v as VbaObject).kind === "object";
const num = (v: VbaValue, d = 0) => (v === undefined || v === null ? d : vbaToNumber(v));
const bgrOf = (hex: string) => { const n = parseInt(hex.replace("#", ""), 16); return ((n >> 16) & 255) | (n & 0xff00) | ((n & 255) << 16); };
const hexOf = (bgr: VbaValue) => { const n = num(bgr); return "#" + [n & 255, (n >> 8) & 255, (n >> 16) & 255].map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase(); };

function vbaValueOf(cell: Cell | undefined, fmt: string, value2: boolean): VbaValue {
  const v = cell?.v ?? null;
  if (v === null) return undefined;
  if (v instanceof CellError) return v.code;
  if (typeof v === "number" && !value2 && isDateFormat(fmt)) return dateOfSerial(v);
  return v;
}

type RangeRef = { sheet: Sheet; rect: Rect; areas: Rect[] };
const R = Symbol("range");
const S = Symbol("sheet");

export function model(app: MacroApp, runtime: () => VbaRuntime | null): VbaHost["global"] {
  const wb = () => app.wb;
  const si = (s: Sheet) => wb().sheets.indexOf(s);
  const ctx = () => ({ wb: wb(), history: app.history });
  const recalc = (sheet: Sheet, r: Rect) => {
    const cells: { row: number; col: number }[] = [];
    for (let row = r.r0; row <= Math.min(r.r1, r.r0 + 2000); row++) for (let col = r.c0; col <= Math.min(r.c1, r.c0 + 200); col++) cells.push({ row, col });
    app.calc.changed(si(sheet), cells);
  };
  const colArg = (v: VbaValue) => (typeof v === "string" && /^[A-Za-z]+$/.test(v) ? v.toUpperCase().split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) : num(v));
  const refOf = (o: VbaValue): RangeRef => { const r = isObj(o) ? (o as unknown as Record<symbol, RangeRef>)[R] : undefined; if (!r) throw new VbaError("Typen unverträglich", 13); return r; };
  const sheetOf = (o: VbaValue): Sheet => { const s = isObj(o) ? (o as unknown as Record<symbol, Sheet>)[S] : undefined; if (!s) throw new VbaError("Typen unverträglich", 13); return s; };
  const bounds = (sheet: Sheet) => sheet.usedRange();

  function parseRef(sheet: Sheet, text: string, origin = { r0: 0, c0: 0 }): VbaObject {
    let s = sheet, t = text.replace(/\$/g, "").trim();
    const bang = t.lastIndexOf("!");
    if (bang > 0) { const found = wb().sheetByName(t.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'")); if (!found) throw new VbaError("Index außerhalb des gültigen Bereichs", 9); s = found; t = t.slice(bang + 1); }
    const dn = wb().names.find((n) => n.name.toLowerCase() === t.toLowerCase());
    if (dn) return parseRef(s, dn.ref.replace(/^=/, ""));
    const areas = t.split(/[;,]/).map((p) => p.trim()).filter(Boolean).map((p) => {
      let m = /^([A-Za-z]+):([A-Za-z]+)$/.exec(p);
      if (m) return { r0: 0, r1: 1048575, c0: colArg(m[1]) - 1, c1: colArg(m[2]) - 1 };
      m = /^(\d+):(\d+)$/.exec(p);
      if (m) return { r0: Number(m[1]) - 1, r1: Number(m[2]) - 1, c0: 0, c1: 16383 };
      const r = parseRange(p.toUpperCase());
      if (!r) throw new VbaError("Die Methode 'Range' für das Objekt '_Global' ist fehlgeschlagen.", 1004);
      return { r0: r.r0 + origin.r0, r1: r.r1 + origin.r0, c0: r.c0 + origin.c0, c1: r.c1 + origin.c0 };
    });
    return rangeObj(s, areas[0], areas);
  }

  function rangeObj(sheet: Sheet, rect: Rect, areas: Rect[] = [rect]): VbaObject {
    const rows = rect.r1 - rect.r0 + 1, cols = rect.c1 - rect.c0 + 1;
    const styleOf = (r: number, c: number) => wb().styles.get(sheet.get(r, c)?.s ?? sheet.rows.get(r)?.s ?? sheet.cols.get(c)?.s);
    const each = (fn: (r: number, c: number) => void) => { for (const a of areas) for (let r = a.r0; r <= a.r1; r++) for (let c = a.c0; c <= a.c1; c++) fn(r, c); };
    const sub = (r0: number, c0: number, r1: number, c1: number) => rangeObj(sheet, { r0, c0, r1, c1 });
    const fmt = (change: (s: CellStyle) => void) => A.format(ctx(), sheet, areas, change);
    const write = (r: number, c: number, v: VbaValue, kind: "value" | "formula" | "local") => {
      // A locked cell of a protected sheet is not written (Excel: Laufzeitfehler 1004).
      if (sheet.protection && styleOf(r, c).locked !== false) throw new VbaError("Die Zelle befindet sich auf einem geschützten Blatt und kann nicht geändert werden.", 1004);
      app.history.cell(sheet, r, c);
      const old = sheet.get(r, c);
      if (v === undefined || v === null || v === "") { sheet.set(r, c, old?.s !== undefined || old?.note ? { v: null, s: old?.s, note: old?.note } : undefined); return; }
      if (typeof v === "string" && v.startsWith("=") && v.length > 1) { sheet.set(r, c, { v: null, f: kind === "local" ? toFileFormula(v, wb()) : v.slice(1), s: old?.s, note: old?.note }); return; }
      if (v instanceof Date) { sheet.set(r, c, { v: serialOfDate(v), s: isDateFormat(styleOf(r, c).numFmt) ? old?.s : wb().styles.derive(old?.s, (s) => { s.numFmt = "m/d/yyyy"; }), note: old?.note }); return; }
      if (typeof v === "string" && kind === "local") { A.enter(ctx(), sheet, { r0: r, c0: c, r1: r, c1: c }, v); return; }
      if (typeof v === "string" && /^-?\d+(\.\d+)?(E[+-]?\d+)?$/i.test(v.trim())) { sheet.set(r, c, { v: Number(v.trim()), s: old?.s, note: old?.note }); return; }
      sheet.set(r, c, { v: (typeof v === "object" ? vbaToString(v) : v) as Value, s: old?.s, note: old?.note });
    };
    const setAll = (v: VbaValue, kind: "value" | "formula" | "local") => {
      if (v && typeof v === "object" && (v as VbaArray).kind === "array") {
        const outer = (v as VbaArray).items;
        const nested = outer.length > 0 && !!outer[0] && typeof outer[0] === "object" && (outer[0] as VbaArray).kind === "array";
        for (let i = 0; i < rows; i++) for (let k = 0; k < cols; k++) {
          const item = nested ? (outer[i] as VbaArray | undefined)?.items[k] : rows === 1 ? outer[k] : outer[i];
          write(rect.r0 + i, rect.c0 + k, item === undefined ? "#N/A" : item, kind);
        }
      } else each((r, c) => write(r, c, v, kind));
      recalc(sheet, rect);
    };
    const getAll = (value2: boolean): VbaValue => {
      if (rows === 1 && cols === 1) return vbaValueOf(sheet.get(rect.r0, rect.c0), styleOf(rect.r0, rect.c0).numFmt, value2);
      const items: VbaValue[] = [];
      for (let r = rect.r0; r <= rect.r1; r++) {
        const row: VbaValue[] = [];
        for (let c = rect.c0; c <= rect.c1; c++) row.push(vbaValueOf(sheet.get(r, c), styleOf(r, c).numFmt, value2));
        items.push({ kind: "array", lb: 1, items: row });
      }
      return { kind: "array", lb: 1, items };
    };
    const has = (r: number, c: number) => { const v = sheet.value(r, c); return v !== null && v !== undefined && v !== ""; };
    const end = (dir: number) => {
      const [dr, dc] = dir === XL.xlup ? [-1, 0] : dir === XL.xldown ? [1, 0] : dir === XL.xltoleft ? [0, -1] : [0, 1];
      const inside = (rr: number, cc: number) => rr >= 0 && cc >= 0 && rr < 1048576 && cc < 16384;
      let r = rect.r0, c = rect.c0;
      if (has(r, c) && inside(r + dr, c + dc) && has(r + dr, c + dc)) {
        while (inside(r + dr, c + dc) && has(r + dr, c + dc)) { r += dr; c += dc; }
      } else {
        const b = bounds(sheet);
        r += dr; c += dc;
        while (inside(r, c) && !has(r, c)) {
          if (!b || (dr > 0 && r > b.r1) || (dc > 0 && c > b.c1)) { if (dr > 0) r = 1048575; if (dc > 0) c = 16383; break; }
          r += dr; c += dc;
        }
        r = Math.max(0, Math.min(1048575, r)); c = Math.max(0, Math.min(16383, c));
      }
      return sub(r, c, r, c);
    };
    const currentRegion = () => {
      let { r0, c0, r1, c1 } = rect;
      for (let grew = true; grew;) {
        grew = false;
        const rowHas = (r: number) => { for (let c = Math.max(0, c0 - 1); c <= c1 + 1; c++) if (has(r, c)) return true; return false; };
        const colHas = (c: number) => { for (let r = Math.max(0, r0 - 1); r <= r1 + 1; r++) if (has(r, c)) return true; return false; };
        if (r0 > 0 && rowHas(r0 - 1)) { r0--; grew = true; }
        if (rowHas(r1 + 1)) { r1++; grew = true; }
        if (c0 > 0 && colHas(c0 - 1)) { c0--; grew = true; }
        if (colHas(c1 + 1)) { c1++; grew = true; }
      }
      return sub(r0, c0, r1, c1);
    };
    const s0 = () => styleOf(rect.r0, rect.c0);
    const font = () => obj({
      Bold: () => !!s0().font.bold, Italic: () => !!s0().font.italic, Underline: () => (s0().font.underline ? XL.xlunderlinestylesingle : XL.xlunderlinestylenone),
      Size: () => s0().font.size, Name: () => s0().font.name, Color: () => bgrOf(s0().font.color ?? "#000000"), Strikethrough: () => !!s0().font.strike,
      ColorIndex: () => { const i = PALETTE.indexOf((s0().font.color ?? "#000000").replace("#", "").toUpperCase()); return i < 0 ? XL.xlautomatic : i + 1; },
    }, {
      Bold: (v) => fmt((s) => { s.font = { ...s.font, bold: !!v || undefined }; }),
      Italic: (v) => fmt((s) => { s.font = { ...s.font, italic: !!v || undefined }; }),
      Underline: (v) => fmt((s) => { s.font = { ...s.font, underline: v === false || num(v) === XL.xlunderlinestylenone ? undefined : "single" }; }),
      Size: (v) => fmt((s) => { s.font = { ...s.font, size: num(v) }; }),
      Name: (v) => fmt((s) => { s.font = { ...s.font, name: vbaToString(v), scheme: undefined }; }),
      Color: (v) => fmt((s) => { s.font = { ...s.font, color: hexOf(v), theme: undefined }; }),
      ColorIndex: (v) => fmt((s) => { const i = num(v); s.font = { ...s.font, color: i > 0 && i <= 56 ? "#" + PALETTE[i - 1] : undefined, theme: undefined }; }),
      Strikethrough: (v) => fmt((s) => { s.font = { ...s.font, strike: !!v || undefined }; }),
    });
    const interior = () => obj({
      Color: () => bgrOf(s0().fill ?? "#FFFFFF"),
      ColorIndex: () => { const f = s0().fill; if (!f) return XL.xlnone; const i = PALETTE.indexOf(f.replace("#", "").toUpperCase()); return i < 0 ? XL.xlnone : i + 1; },
      Pattern: () => (s0().fill ? XL.xlsolid : XL.xlnone),
    }, {
      Color: (v) => fmt((s) => { s.fill = hexOf(v); s.fillTheme = undefined; s.fillXml = undefined; }),
      ColorIndex: (v) => fmt((s) => { const i = num(v); s.fill = i > 0 && i <= 56 ? "#" + PALETTE[i - 1] : undefined; s.fillTheme = undefined; s.fillXml = undefined; }),
      Pattern: (v) => { if (num(v) === XL.xlnone) fmt((s) => { s.fill = undefined; s.fillTheme = undefined; }); },
    });
    const lineStyle = (ls: number, w: number): import("@raabe/sheet-engine").BorderSide["style"] | undefined =>
      ls === XL.xlnone ? undefined : ls === XL.xldouble ? "double" : ls === XL.xldash ? (w === XL.xlmedium ? "mediumDashed" : "dashed") : ls === XL.xldot ? "dotted"
        : w === XL.xlthick ? "thick" : w === XL.xlmedium ? "medium" : w === XL.xlhairline ? "hair" : "thin";
    const border = (which: number | null) => {
      const st = { ls: XL.xlcontinuous, w: XL.xlthin, color: undefined as string | undefined };
      const apply = () => {
        const style = lineStyle(st.ls, st.w);
        for (const a of areas) for (let r = a.r0; r <= a.r1; r++) for (let c = a.c0; c <= a.c1; c++) {
          const sides: ("left" | "top" | "bottom" | "right")[] = [];
          if (which === null) sides.push("left", "top", "bottom", "right");
          if (which === XL.xledgetop && r === a.r0) sides.push("top");
          if (which === XL.xledgebottom && r === a.r1) sides.push("bottom");
          if (which === XL.xledgeleft && c === a.c0) sides.push("left");
          if (which === XL.xledgeright && c === a.c1) sides.push("right");
          if (which === XL.xlinsidehorizontal && r < a.r1) sides.push("bottom");
          if (which === XL.xlinsidevertical && c < a.c1) sides.push("right");
          if (!sides.length) continue;
          A.format(ctx(), sheet, [{ r0: r, c0: c, r1: r, c1: c }], (s) => {
            s.borders = { ...s.borders };
            for (const sd of sides) s.borders[sd] = style ? { style, ...(st.color ? { color: st.color } : {}) } : undefined;
          });
        }
      };
      return obj({ LineStyle: () => st.ls, Weight: () => st.w, Color: () => bgrOf(st.color ?? "#000000") },
        { LineStyle: (v) => { st.ls = num(v); apply(); }, Weight: (v) => { st.w = num(v); apply(); }, Color: (v) => { st.color = hexOf(v); apply(); } });
    };
    const special = (kind: number): VbaValue => {
      const b = bounds(sheet);
      if (kind === XL.xlcelltypelastcell) return b ? sub(b.r1, b.c1, b.r1, b.c1) : sub(0, 0, 0, 0);
      const r1 = b ? Math.min(rect.r1, b.r1) : rect.r0, c1 = b ? Math.min(rect.c1, b.c1) : rect.c0;
      const hits: Rect[] = [];
      for (let r = rect.r0; r <= r1; r++) for (let c = rect.c0; c <= c1; c++) {
        const cell = sheet.get(r, c);
        const ok = kind === XL.xlcelltypeblanks ? !has(r, c) && cell?.f === undefined
          : kind === XL.xlcelltypeconstants ? cell?.f === undefined && has(r, c)
          : kind === XL.xlcelltypeformulas ? cell?.f !== undefined
          : kind === XL.xlcelltypevisible ? !sheet.rows.get(r)?.hidden && !sheet.cols.get(c)?.hidden : false;
        if (ok) hits.push({ r0: r, c0: c, r1: r, c1: c });
      }
      if (!hits.length) throw new VbaError("Keine Zellen gefunden.", 1004);
      return rangeObj(sheet, hits[0], hits);
    };
    const find = (what: string, whole: boolean, after: VbaValue): VbaValue => {
      const b = bounds(sheet);
      if (!b) return null;
      const r1 = Math.min(rect.r1, b.r1), c1 = Math.min(rect.c1, b.c1);
      const cells: [number, number][] = [];
      for (let r = rect.r0; r <= r1; r++) for (let c = rect.c0; c <= c1; c++) cells.push([r, c]);
      const start = isObj(after) ? refOf(after).rect : null;
      const from = start ? cells.findIndex(([r, c]) => r === start.r0 && c === start.c0) + 1 : 0;
      const want = what.toLowerCase();
      for (let k = 0; k < cells.length; k++) {
        const [r, c] = cells[(from + k) % cells.length];
        const t = app.shownText(sheet, r, c).toLowerCase();
        if (whole ? t === want : t.includes(want)) return sub(r, c, r, c);
      }
      return null;
    };
    const self: VbaObject = obj({
      Value: () => getAll(false), Value2: () => getAll(true), Text: () => app.shownText(sheet, rect.r0, rect.c0),
      Formula: () => { const c = sheet.get(rect.r0, rect.c0); return c?.f !== undefined ? "=" + c.f : vbaValueOf(c, "General", true) ?? ""; },
      FormulaLocal: () => { const c = sheet.get(rect.r0, rect.c0); return c?.f !== undefined ? "=" + toLocalFormula(c.f, wb()) : vbaValueOf(c, "General", true) ?? ""; },
      HasFormula: () => sheet.get(rect.r0, rect.c0)?.f !== undefined,
      Address: (a) => {
        const ra = a[0] === undefined ? true : !!a[0], ca = a[1] === undefined ? true : !!a[1];
        const one = (r: number, c: number) => `${ca ? "$" : ""}${colName(c)}${ra ? "$" : ""}${r + 1}`;
        return rows === 1 && cols === 1 ? one(rect.r0, rect.c0) : `${one(rect.r0, rect.c0)}:${one(rect.r1, rect.c1)}`;
      },
      Row: () => rect.r0 + 1, Column: () => rect.c0 + 1, Count: () => areas.reduce((n, a) => n + (a.r1 - a.r0 + 1) * (a.c1 - a.c0 + 1), 0),
      Areas: () => obj({ Count: () => areas.length, Item: (a) => rangeObj(sheet, areas[num(a[0]) - 1]) }),
      Rows: (a) => (a.length ? sub(rect.r0 + num(a[0]) - 1, rect.c0, rect.r0 + num(a[0]) - 1, rect.c1)
        : obj({ Count: () => rows, Item: (b) => sub(rect.r0 + num(b[0]) - 1, rect.c0, rect.r0 + num(b[0]) - 1, rect.c1) }, {}, {}, { items: () => Array.from({ length: rows }, (_, i) => sub(rect.r0 + i, rect.c0, rect.r0 + i, rect.c1)) })),
      Columns: (a) => (a.length ? sub(rect.r0, rect.c0 + colArg(a[0]) - 1, rect.r1, rect.c0 + colArg(a[0]) - 1)
        : obj({ Count: () => cols, Item: (b) => sub(rect.r0, rect.c0 + num(b[0]) - 1, rect.r1, rect.c0 + num(b[0]) - 1) }, {}, {}, { items: () => Array.from({ length: cols }, (_, i) => sub(rect.r0, rect.c0 + i, rect.r1, rect.c0 + i)) })),
      Cells: (a) => {
        if (!a.length) return self;
        if (a.length === 1) { const n = num(a[0]) - 1; return sub(rect.r0 + Math.floor(n / cols), rect.c0 + (n % cols), rect.r0 + Math.floor(n / cols), rect.c0 + (n % cols)); }
        return sub(rect.r0 + num(a[0]) - 1, rect.c0 + colArg(a[1]) - 1, rect.r0 + num(a[0]) - 1, rect.c0 + colArg(a[1]) - 1);
      },
      Item: (a) => sub(rect.r0 + num(a[0]) - 1, rect.c0 + (a.length > 1 ? colArg(a[1]) - 1 : 0), rect.r0 + num(a[0]) - 1, rect.c0 + (a.length > 1 ? colArg(a[1]) - 1 : 0)),
      Range: (a) => (isObj(a[0]) && isObj(a[1]) ? span(refOf(a[0]), refOf(a[1])) : parseRef(sheet, vbaToString(a[0]), { r0: rect.r0, c0: rect.c0 })),
      Offset: (a) => sub(rect.r0 + num(a[0]), rect.c0 + num(a[1]), rect.r1 + num(a[0]), rect.c1 + num(a[1])),
      Resize: (a) => sub(rect.r0, rect.c0, rect.r0 + (a[0] === undefined ? rows : num(a[0])) - 1, rect.c0 + (a[1] === undefined ? cols : num(a[1])) - 1),
      End: (a) => end(num(a[0])), CurrentRegion: () => currentRegion(),
      EntireRow: () => sub(rect.r0, 0, rect.r1, 16383), EntireColumn: () => sub(0, rect.c0, 1048575, rect.c1),
      Font: () => font(), Interior: () => interior(), Borders: (a) => border(a.length ? num(a[0]) : null),
      NumberFormat: () => s0().numFmt, NumberFormatLocal: () => app.localFormat(s0().numFmt),
      HorizontalAlignment: () => ({ left: XL.xlleft, center: XL.xlcenter, right: XL.xlright, justify: XL.xljustify } as Record<string, number>)[s0().h ?? ""] ?? XL.xlgeneral,
      VerticalAlignment: () => ({ top: XL.xltop, center: XL.xlcenter } as Record<string, number>)[s0().v ?? ""] ?? XL.xlbottom,
      WrapText: () => !!s0().wrap, ColumnWidth: () => app.columnWidth(sheet, rect.c0), RowHeight: () => app.rowHeight(sheet, rect.r0),
      Hidden: () => (rect.c0 === 0 && rect.c1 >= 16383 ? !!sheet.rows.get(rect.r0)?.hidden : !!sheet.cols.get(rect.c0)?.hidden),
      MergeCells: () => sheet.merges.some((m) => m.r0 <= rect.r0 && m.r1 >= rect.r0 && m.c0 <= rect.c0 && m.c1 >= rect.c0),
      Locked: () => s0().locked !== false, Worksheet: () => sheetObj(sheet), Parent: () => sheetObj(sheet),
      Comment: () => { const n = sheet.get(rect.r0, rect.c0)?.note; return n === undefined ? null : obj({ Text: () => n }, {}, { Delete: () => { app.setNote(sheet, rect.r0, rect.c0, null); return undefined; } }); },
    }, {
      Value: (v) => setAll(v, "value"), Value2: (v) => setAll(v, "value"), Formula: (v) => setAll(v, "formula"), FormulaLocal: (v) => setAll(v, "local"),
      NumberFormat: (v) => fmt((s) => { s.numFmt = vbaToString(v); }), NumberFormatLocal: (v) => fmt((s) => { s.numFmt = app.fileFormat(vbaToString(v)); }),
      HorizontalAlignment: (v) => fmt((s) => { const n = num(v); s.h = n === XL.xlleft ? "left" : n === XL.xlcenter ? "center" : n === XL.xlright ? "right" : n === XL.xljustify ? "justify" : undefined; }),
      VerticalAlignment: (v) => fmt((s) => { const n = num(v); s.v = n === XL.xltop ? "top" : n === XL.xlcenter ? "center" : undefined; }),
      WrapText: (v) => fmt((s) => { s.wrap = !!v || undefined; }),
      ColumnWidth: (v) => { for (let c = rect.c0; c <= Math.min(rect.c1, rect.c0 + 500); c++) app.setColumnWidth(sheet, c, num(v)); },
      RowHeight: (v) => { for (let r = rect.r0; r <= Math.min(rect.r1, rect.r0 + 5000); r++) app.setRowHeight(sheet, r, num(v)); },
      Hidden: (v) => {
        if (rect.c0 === 0 && rect.c1 >= 16383) for (let r = rect.r0; r <= rect.r1; r++) { app.history.row(sheet, r); const i = sheet.rows.get(r); sheet.rows.set(r, { ...(i ?? {}), hidden: !!v || undefined }); }
        else for (let c = rect.c0; c <= rect.c1; c++) { app.history.col(sheet, c); const i = sheet.cols.get(c); sheet.cols.set(c, { ...(i ?? {}), hidden: !!v || undefined }); }
      },
      MergeCells: (v) => app.merge(sheet, rect, !!v),
      Locked: (v) => fmt((s) => { s.locked = v ? undefined : false; }),
    }, {
      Select: () => { app.selectRange(sheet, rect); return undefined; }, Activate: () => { app.selectRange(sheet, rect); return undefined; },
      ClearContents: () => { A.clearContents(ctx(), sheet, areas); recalc(sheet, rect); return undefined; },
      Clear: () => { A.clearAll(ctx(), sheet, areas, "all"); recalc(sheet, rect); return undefined; },
      ClearFormats: () => { A.clearAll(ctx(), sheet, areas, "formats"); return undefined; },
      ClearComments: () => { A.clearAll(ctx(), sheet, areas, "notes"); return undefined; },
      Merge: () => { app.merge(sheet, rect, true); return undefined; }, UnMerge: () => { app.merge(sheet, rect, false); return undefined; },
      Delete: (a) => { app.deleteCells(sheet, rect, a[0] === undefined ? null : num(a[0]) === XL.xltoleft ? "left" : "up"); app.recalc(); return undefined; },
      Insert: (a) => { app.insertCells(sheet, rect, a[0] === undefined ? null : num(a[0]) === XL.xltoright ? "right" : "down"); app.recalc(); return undefined; },
      Copy: (a) => {
        if (!isObj(a[0])) throw new VbaError("Kopieren in die Zwischenablage aus Makros ist nicht möglich; geben Sie Destination an.", 1004);
        const d = refOf(a[0]); A.copyCells(ctx(), { sheet, rect }, d.sheet, { row: d.rect.r0, col: d.rect.c0 }); app.recalc(); return undefined;
      },
      Cut: (a) => {
        if (!isObj(a[0])) throw new VbaError("Ausschneiden in die Zwischenablage aus Makros ist nicht möglich; geben Sie Destination an.", 1004);
        const d = refOf(a[0]); A.copyCells(ctx(), { sheet, rect }, d.sheet, { row: d.rect.r0, col: d.rect.c0 }); A.clearAll(ctx(), sheet, [rect], "all"); app.recalc(); return undefined;
      },
      AutoFit: () => { if (rect.r0 === 0 && rect.r1 >= 1048575) app.autoFitColumns(sheet, rect.c0, rect.c1); return undefined; },
      Sort: (a) => {
        const keys: { col: number; desc: boolean }[] = [];
        const add = (k: VbaValue, o: VbaValue) => { if (isObj(k)) keys.push({ col: refOf(k).rect.c0, desc: num(o, XL.xlascending) === XL.xldescending }); };
        add(a[0], a[1]); add(a[2], a[4]); add(a[5], a[6]);
        const header = num(a[7], XL.xlno) === XL.xlyes;
        app.sortRows(sheet, header ? { ...rect, r0: rect.r0 + 1 } : rect, keys.length ? keys : [{ col: rect.c0, desc: false }]);
        app.recalc();
        return undefined;
      },
      Find: (a) => find(vbaToString(a[0]), num(a[3], XL.xlpart) === XL.xlwhole, a[1]),
      Replace: (a) => {
        const what = vbaToString(a[0]), by = vbaToString(a[1] ?? ""), whole = num(a[2], XL.xlpart) === XL.xlwhole;
        each((r, c) => {
          const cell = sheet.get(r, c);
          if (!cell || cell.f !== undefined || typeof cell.v !== "string") return;
          const next = whole ? (cell.v.toLowerCase() === what.toLowerCase() ? by : cell.v) : cell.v.split(new RegExp(what.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi")).join(by);
          if (next !== cell.v) { app.history.cell(sheet, r, c); sheet.set(r, c, { ...cell, v: next }); }
        });
        recalc(sheet, rect);
        return true;
      },
      SpecialCells: (a) => special(num(a[0])),
      BorderAround: (a) => { const b = border(null); if (a[0] !== undefined) b.set!("LineStyle", a[0], []); if (a[1] !== undefined) b.set!("Weight", a[1], []); if (a[3] !== undefined) b.set!("Color", a[3], []); return true; },
      AddComment: (a) => { app.setNote(sheet, rect.r0, rect.c0, vbaToString(a[0] ?? "")); return obj({ Text: () => vbaToString(a[0] ?? "") }); },
      Calculate: () => { recalc(sheet, rect); return undefined; },
      FillDown: () => { A.autoFill(ctx(), sheet, { ...rect, r1: rect.r0 }, rect, { copy: true }); app.recalc(); return undefined; },
      FillRight: () => { A.autoFill(ctx(), sheet, { ...rect, c1: rect.c0 }, rect, { copy: true }); app.recalc(); return undefined; },
      AutoFill: (a) => { if (isObj(a[0])) A.autoFill(ctx(), sheet, rect, refOf(a[0]).rect, { series: num(a[1]) === XL.xlfillseries }); app.recalc(); return true; },
    }, {
      // Range("A1:B5")(2, 1) is its Item; without an index the value.
      default: (a) => (a.length ? self.get("Item", a) : getAll(false)),
      // For Each over a whole sheet, row or column visits the cells in use (Excel would visit billions).
      items: () => {
        const out: VbaValue[] = [];
        const b = bounds(sheet);
        for (const x of areas) {
          const big = (x.r1 - x.r0 + 1) * (x.c1 - x.c0 + 1) > 1_000_000;
          const a = big ? (b ? { r0: Math.max(x.r0, b.r0), c0: Math.max(x.c0, b.c0), r1: Math.min(x.r1, b.r1), c1: Math.min(x.c1, b.c1) } : null) : x;
          if (a) for (let r = a.r0; r <= a.r1; r++) for (let c = a.c0; c <= a.c1; c++) out.push(sub(r, c, r, c));
        }
        return out;
      },
      typeName: () => "Range",
    });
    (self as unknown as Record<symbol, RangeRef>)[R] = { sheet, rect, areas };
    return self;
  }

  const span = (a: RangeRef, b: RangeRef) => rangeObj(a.sheet, { r0: Math.min(a.rect.r0, b.rect.r0), c0: Math.min(a.rect.c0, b.rect.c0), r1: Math.max(a.rect.r1, b.rect.r1), c1: Math.max(a.rect.c1, b.rect.c1) });

  function sheetObj(sheet: Sheet): VbaObject {
    const self = obj({
      Name: () => sheet.name, Index: () => si(sheet) + 1, CodeName: () => sheet.name,
      Range: (a) => (isObj(a[0]) && isObj(a[1]) ? span(refOf(a[0]), refOf(a[1])) : parseRef(sheet, vbaToString(a[0]))),
      Cells: (a) => (a.length ? rangeObj(sheet, { r0: num(a[0]) - 1, c0: colArg(a[1]) - 1, r1: num(a[0]) - 1, c1: colArg(a[1]) - 1 }) : rangeObj(sheet, { r0: 0, c0: 0, r1: 1048575, c1: 16383 })),
      Rows: (a) => (a.length ? (typeof a[0] === "string" ? parseRef(sheet, String(a[0])) : rangeObj(sheet, { r0: num(a[0]) - 1, c0: 0, r1: num(a[0]) - 1, c1: 16383 })) : rangeObj(sheet, { r0: 0, c0: 0, r1: 1048575, c1: 16383 }).get("Rows", [])),
      Columns: (a) => (a.length ? (typeof a[0] === "string" ? parseRef(sheet, String(a[0]).includes(":") ? String(a[0]) : `${a[0]}:${a[0]}`) : rangeObj(sheet, { r0: 0, c0: num(a[0]) - 1, r1: 1048575, c1: num(a[0]) - 1 })) : rangeObj(sheet, { r0: 0, c0: 0, r1: 1048575, c1: 16383 }).get("Columns", [])),
      UsedRange: () => rangeObj(sheet, bounds(sheet) ?? { r0: 0, c0: 0, r1: 0, c1: 0 }),
      Visible: () => (sheet.hidden === "very" ? XL.xlsheetveryhidden : sheet.hidden ? XL.xlsheethidden : XL.xlsheetvisible),
      ProtectContents: () => !!sheet.protection, Parent: () => workbookObj(),
      Tab: () => obj({ Color: () => (sheet.tabColor ? bgrOf(sheet.tabColor) : false) }, { Color: (v) => app.setTabColor(sheet, v === false ? null : hexOf(v)) }),
      AutoFilterMode: () => !!sheet.autoFilter, FilterMode: () => !!sheet.filterMode,
    }, {
      Name: (v) => { const err = app.renameSheet(si(sheet), vbaToString(v)); if (err) throw new VbaError(err, 1004); },
      Visible: (v) => app.setSheetVisible(sheet, num(v) === XL.xlsheetveryhidden ? "veryHidden" : v === false || num(v) === XL.xlsheethidden ? "hidden" : "visible"),
    }, {
      Activate: () => { app.activateSheet(si(sheet)); return undefined; }, Select: () => { app.activateSheet(si(sheet)); return undefined; },
      Delete: () => { app.deleteSheetAt(si(sheet)); return true; },
      Copy: (a) => { app.copySheet(si(sheet), isObj(a[1]) ? si(sheetOf(a[1])) + 1 : isObj(a[0]) ? si(sheetOf(a[0])) : wb().sheets.length); return undefined; },
      Move: (a) => { app.moveSheetTo(si(sheet), isObj(a[1]) ? si(sheetOf(a[1])) + 1 : isObj(a[0]) ? si(sheetOf(a[0])) : wb().sheets.length); return undefined; },
      Protect: (a) => { app.protectSheet(sheet, a[0] === undefined ? "" : vbaToString(a[0])); return undefined; },
      Unprotect: (a) => { if (!app.unprotectSheet(sheet, a[0] === undefined ? "" : vbaToString(a[0]))) throw new VbaError("Das angegebene Kennwort ist falsch.", 1004); return undefined; },
      Calculate: () => { app.recalc(); return undefined; },
    }, { typeName: () => "Worksheet" }, SHEET_ARGS);
    (self as unknown as Record<symbol, Sheet>)[S] = sheet;
    return self;
  }

  function sheetsColl(): VbaObject {
    const item = (k: VbaValue) => {
      const s = typeof k === "number" ? wb().sheets[k - 1] : wb().sheetByName(vbaToString(k));
      if (!s) throw new VbaError("Index außerhalb des gültigen Bereichs", 9);
      return sheetObj(s);
    };
    return obj({ Count: () => wb().sheets.length, Item: (a) => item(a[0]) }, {}, {
      Add: (a) => {
        const at = isObj(a[1]) ? si(sheetOf(a[1])) + 1 : isObj(a[0]) ? si(sheetOf(a[0])) : wb().active;
        let last: Sheet | undefined;
        for (let i = 0; i < Math.max(1, num(a[2], 1)); i++) last = app.insertSheetAt(at + i);
        return sheetObj(last!);
      },
    }, { items: () => wb().sheets.map((s) => sheetObj(s)), default: (a) => item(a[0]), typeName: () => "Sheets" });
  }

  function workbookObj(): VbaObject {
    return obj({
      Name: () => app.docTitle(), FullName: () => app.docPath() ?? app.docTitle(), Path: () => (app.docPath() ?? "").replace(/[\\/][^\\/]*$/, ""),
      Worksheets: (a) => (a.length ? sheetsColl().get("Item", a) : sheetsColl()), Sheets: (a) => (a.length ? sheetsColl().get("Item", a) : sheetsColl()),
      ActiveSheet: () => sheetObj(app.sheet),
      Names: () => obj({ Count: () => wb().names.length }, {}, { Add: (a) => { wb().names.push({ name: vbaToString(a[0]), ref: vbaToString(a[1]).replace(/^=/, "") }); return undefined; } }),
    }, {}, {
      Save: async () => { await app.saveTo(null); return undefined; },
      Activate: () => undefined, Close: () => undefined,
    }, { typeName: () => "Workbook" });
  }

  const wsfNames = ["Sum", "Average", "Max", "Min", "Count", "CountA", "CountIf", "SumIf", "AverageIf", "VLookup", "HLookup", "Match", "Index", "Round", "RoundUp",
    "RoundDown", "Median", "StDev", "Large", "Small", "Rank", "CountBlank", "SumProduct", "Trim", "Proper", "Power", "Mode", "Var", "Text", "Weekday", "Pi", "Sqrt",
    "Abs", "Int", "Mod", "Ln", "Exp", "Fact", "Combin", "Concat", "TextJoin"];
  /** WorksheetFunction: the engine works the function out on the arguments (ranges as references). */
  const worksheetFunction = (fn: string, args: VbaValue[]): VbaValue => {
    const lit = (v: VbaValue): string => {
      if (isObj(v)) { const r = refOf(v); const q = r.sheet.name.replace(/'/g, "''"); return `'${q}'!${colName(r.rect.c0)}${r.rect.r0 + 1}:${colName(r.rect.c1)}${r.rect.r1 + 1}`; }
      if (typeof v === "string") return `"${v.replace(/"/g, '""')}"`;
      if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
      if (v instanceof Date) return String(serialOfDate(v));
      return v === undefined || v === null ? "" : String(v);
    };
    const r = app.calc.evaluate(wb().active, 0, 0, `${fn.toUpperCase()}(${args.map(lit).join(",")})`);
    const v = Array.isArray(r) ? r[0]?.[0] : r;
    if (v instanceof CellError) throw new VbaError(`Die ${fn}-Eigenschaft des WorksheetFunction-Objektes kann nicht zugeordnet werden.`, 1004);
    return v === null ? 0 : (v as VbaValue);
  };
  const wsf = obj({}, {}, Object.fromEntries(wsfNames.map((n) => [n, (a: VbaValue[]) => worksheetFunction(n, a)])), { typeName: () => "WorksheetFunction" });

  const application: VbaObject = obj({
    ActiveSheet: () => sheetObj(app.sheet), ActiveWorkbook: () => workbookObj(), ThisWorkbook: () => workbookObj(),
    ActiveCell: () => { const a = app.sheet.view.active; return rangeObj(app.sheet, { r0: a.row, c0: a.col, r1: a.row, c1: a.col }); },
    Selection: () => { const sel = app.sheet.view.selection; return rangeObj(app.sheet, sel[0], sel); },
    Worksheets: (a) => (a.length ? sheetsColl().get("Item", a) : sheetsColl()), Sheets: (a) => (a.length ? sheetsColl().get("Item", a) : sheetsColl()),
    Range: (a) => sheetObj(app.sheet).get("Range", a), Cells: (a) => sheetObj(app.sheet).get("Cells", a),
    Rows: (a) => sheetObj(app.sheet).get("Rows", a), Columns: (a) => sheetObj(app.sheet).get("Columns", a),
    WorksheetFunction: () => wsf, ScreenUpdating: () => true, DisplayAlerts: () => app.alerts, EnableEvents: () => true,
    Calculation: () => (wb().calcMode === "manual" ? XL.xlcalculationmanual : XL.xlcalculationautomatic),
    StatusBar: () => false, UserName: () => app.userName, Version: () => "16.0", Name: () => "Microsoft Excel", CutCopyMode: () => false,
  }, {
    ScreenUpdating: () => undefined, DisplayAlerts: (v) => { app.alerts = !!v; }, EnableEvents: () => undefined, CutCopyMode: () => undefined,
    Calculation: (v) => { wb().calcMode = num(v) === XL.xlcalculationmanual ? "manual" : undefined; },
    StatusBar: (v) => app.statusText(v === false ? null : vbaToString(v)),
  }, {
    Calculate: () => { app.recalc(); return undefined; }, Wait: () => true, Volatile: () => undefined,
    Union: (a) => { const rs = a.filter(isObj).map(refOf); return rangeObj(rs[0].sheet, rs[0].rect, rs.flatMap((r) => r.areas)); },
    Intersect: (a) => {
      const rs = a.filter(isObj).map(refOf);
      const x = rs.map((r) => r.rect).reduce((p, q) => ({ r0: Math.max(p.r0, q.r0), c0: Math.max(p.c0, q.c0), r1: Math.min(p.r1, q.r1), c1: Math.min(p.c1, q.c1) }));
      return x.r0 > x.r1 || x.c0 > x.c1 ? null : rangeObj(rs[0].sheet, x);
    },
    InputBox: async (a) => {
      const t = await app.inputBox(vbaToString(a[0] ?? ""), a[1] === undefined ? "Eingabe" : vbaToString(a[1]), a[2] === undefined ? "" : vbaToString(a[2]));
      if (t === null) return false;
      const type = num(a[7], 2);
      if (type === 1) { const n = Number(t.replace(",", ".")); if (Number.isNaN(n)) throw new VbaError("Ungültige Zahl.", 13); return n; }
      if (type === 8) return parseRef(app.sheet, t);
      return t;
    },
    Goto: (a) => { if (isObj(a[0])) { const r = refOf(a[0]); app.selectRange(r.sheet, r.rect); } return undefined; },
    Run: async (a) => runtime()?.run(vbaToString(a[0]), a.slice(1)),
  }, { typeName: () => "Application" });

  const lazy = (name: string): VbaObject => {
    const target = () => application.get(name, []) as VbaObject;
    const own: Partial<VbaObject> = { default: (a: VbaValue[]) => (a.length ? application.get(name, a) : (target().default?.([]) ?? target())) };
    return new Proxy({ kind: "object" } as VbaObject, {
      get: (_t, prop) => (prop === "kind" ? "object" : prop === "default" ? own.default : Reflect.get(target(), prop)),
      has: (_t, prop) => prop in target(),
    });
  };

  return (name: string): VbaValue => {
    const k = name.toLowerCase();
    if (k in XL) return XL[k];
    switch (k) {
      case "application": return application;
      case "activesheet": case "activeworkbook": case "thisworkbook": case "activecell": case "selection": case "worksheets": case "sheets":
      case "worksheetfunction":
        return application.get(name, []);
      // Range("A1"), Cells(3, 2), Rows(5), Worksheets("x"): the interpreter takes the global first and indexes it
      // then; these stand for Excel's method, with the index or without (Rows.Count, Worksheets.Add).
      case "range": case "cells": case "rows": case "columns": case "worksheets": case "sheets":
        return lazy(name);
    }
    const s = wb().sheetByName(name);
    return s ? sheetObj(s) : undefined;
  };
}

/** Runs `name` from `code` against a MacroApp (tests): the error, if any, as VBA numbers it. */
export async function runOn(m: MacroApp, code: string, name: string): Promise<{ number: number; message: string; line: number } | null> {
  let rt: VbaRuntime | null = null;
  rt = new VbaRuntime({ global: model(m, () => rt), msgBox: async () => 1, inputBox: async () => "" }, [code]);
  let err: ReturnType<typeof asError> | null = null;
  await m.changeAsync(`Makro ${name}`, async () => { try { await rt!.run(name); } catch (e) { err = asError(e); } });
  return err;
}

// ----- the project, running, the dialog ---------------------------------------------------------------------

export function modulesOf(app: CalcApp): VbaModule[] {
  const bytes = app.wb.vba;
  if (!bytes) return [];
  try { return readProject(bytes); } catch { return []; }
}

export function setModules(app: CalcApp, modules: VbaModule[]): void {
  app.wb.vba = writeProject(modules);
  enabled.add(app.wb);
}

export const macrosEnabled = (app: CalcApp) => enabled.has(app.wb);
export const enableMacros = (app: CalcApp) => enabled.add(app.wb);

const BUTTONS = (buttons: number): [string, number][] => {
  switch (buttons & 7) {
    case 1: return [["OK", 1], ["Abbrechen", 2]];
    case 2: return [["Abbrechen", 3], ["Wiederholen", 4], ["Ignorieren", 5]];
    case 3: return [["Ja", 6], ["Nein", 7], ["Abbrechen", 2]];
    case 4: return [["Ja", 6], ["Nein", 7]];
    case 5: return [["Wiederholen", 4], ["Abbrechen", 2]];
    default: return [["OK", 1]];
  }
};

/** The result of a run (for tests and the caller): errors as VBA numbers them, the messages shown, Debug.Print lines. */
export interface MacroRun { ok: boolean; error?: { number: number; message: string; line: number }; shown: string[]; printed: string[] }

/** Runs a macro by name as one undo step; a runtime error is reported as Excel reports it. */
export async function runMacro(app: CalcApp, name: string, test?: { answers: (string | number)[] }): Promise<MacroRun> {
  const m = macroApp(app);
  if (test) { m.testMode = true; m.answers = [...test.answers]; enabled.add(app.wb); }
  if (!macrosEnabled(app)) { await securityNotice(); if (!macrosEnabled(app)) return { ok: false, shown: [], printed: [] }; }
  let rt: VbaRuntime | null = null;
  const global = model(m, () => rt);
  const hst: VbaHost = {
    global, appName: "Microsoft Excel",
    msgBox: (prompt, buttons, title) => m.msgBox(prompt, BUTTONS(buttons), title || "Microsoft Excel"),
    inputBox: async (prompt, title, def) => (await m.inputBox(prompt, title || "Microsoft Excel", def)) ?? "",
  };
  try { rt = new VbaRuntime(hst, modulesOf(app)); } catch (e) { const err = asError(e); if (!test) await report(err, name); return { ok: false, error: err, shown: m.shown, printed: [] }; }
  let failure: ReturnType<typeof asError> | null = null;
  await m.changeAsync(`Makro ${name}`, async () => { try { await rt!.run(name); } catch (e) { failure = asError(e); } });
  app.calc.recalc();
  // Sheets added, renamed, hidden or activated: the sheet tabs, the grid and the ribbon show the workbook anew.
  if (!app.wb.sheet() || app.wb.sheet().hidden) app.wb.active = Math.max(0, app.wb.sheets.findIndex((s) => !s.hidden));
  app.activate((app as unknown as { active: Parameters<CalcApp["activate"]>[0] }).active);
  app.grid.contentChanged();
  if (failure && !test) await report(failure, name);
  return failure ? { ok: false, error: failure, shown: m.shown, printed: rt.printed } : { ok: true, shown: m.shown, printed: rt.printed };
}

function asError(e: unknown): { number: number; message: string; line: number } {
  const err = e instanceof VbaError ? e : new VbaError(String((e as Error)?.message ?? e));
  return { number: err.number, message: err.message, line: err.line };
}

async function report(err: { number: number; message: string; line: number }, name: string): Promise<void> {
  await showDialog({ title: "Microsoft Visual Basic", body: h("div", { style: "max-width:440px" }, h("b", null, `Laufzeitfehler '${err.number}':`), h("div", null, err.message), err.line ? h("div", { class: "muted" }, `Makro ${name}, Zeile ${err.line}`) : ""),
    buttons: [{ label: "Beenden", value: true, primary: true }], cancel: true });
}

let pendingEnable: (() => void) | null = null;
async function securityNotice(): Promise<void> {
  const yes = await showDialog({ title: "RaabeCalc", body: h("div", { style: "max-width:420px" }, "Die Makros dieser Arbeitsmappe sind deaktiviert. Aktivieren Sie sie nur, wenn Sie der Datei vertrauen."),
    buttons: [{ label: "Inhalt aktivieren", value: true }, { label: "Abbrechen", value: false, primary: true }], cancel: false });
  if (yes) pendingEnable?.();
}

/** The message bar of a workbook with macros (SICHERHEITSWARNUNG, Inhalt aktivieren) in `slot`, below the ribbon. */
export function messageBar(app: CalcApp, slot: HTMLElement): void {
  slot.replaceChildren();
  pendingEnable = () => enabled.add(app.wb);
  if (!app.wb.vba || macrosEnabled(app) || dismissed.has(app.wb)) return;
  const wb = app.wb;
  slot.append(h("div", { class: "infobar macro-bar", role: "status" },
    h("b", null, "SICHERHEITSWARNUNG"), h("span", { class: "msg" }, "Makros wurden deaktiviert."),
    h("button", { onclick: () => { enabled.add(wb); slot.replaceChildren(); } }, "Inhalt aktivieren"),
    h("button", { "aria-label": "Schließen", onclick: () => { dismissed.add(wb); slot.replaceChildren(); } }, "×")));
}
const dismissed = new WeakSet<object>();

/** Ansicht > Makros > Makros anzeigen (Alt+F8): Makroname, the list, Ausführen, Löschen. */
export async function macroDialog(app: CalcApp): Promise<void> {
  pendingEnable = () => enabled.add(app.wb);
  const mods = modulesOf(app);
  let names: string[] = [];
  try { names = new VbaRuntime({ global: () => undefined, msgBox: async () => 1, inputBox: async () => "" }, mods).macros(); } catch { /* a module that does not parse */ }
  const nameInput = h("input", { type: "text", value: names[0] ?? "", style: "width:100%", "aria-label": "Makroname" }) as HTMLInputElement;
  const list = h("select", { size: 10, style: "width:100%", "aria-label": "Makros" }, ...names.map((n) => h("option", { value: n }, n))) as HTMLSelectElement;
  list.addEventListener("change", () => { nameInput.value = list.value; });
  const v = await showDialog<"run" | "delete" | "">({ title: "Makro", body: h("div", { class: "form", style: "width:380px" }, h("label", null, "Makroname:"), nameInput, list, h("div", null, "Makros in: ", h("b", null, "Alle offenen Arbeitsmappen"))),
    buttons: [{ label: "Ausführen", value: "run", primary: true }, { label: "Löschen", value: "delete" }, { label: "Abbrechen", value: "" }], cancel: "" });
  const name = nameInput.value.trim();
  if (v === "run" && name) await runMacro(app, name);
  if (v === "delete" && name) {
    for (const m of mods) m.code = m.code.replace(new RegExp(`(Public |Private )?Sub ${name}\\b[\\s\\S]*?End Sub\\r?\\n?`, "i"), "");
    setModules(app, mods);
  }
}
