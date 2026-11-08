// ZELLE (CELL), INFO, BLATT (SHEET), BLÄTTER (SHEETS) and HYPERLINK (D-080), as recorded in German
// Excel (tools/feel-lab/probe_excel_info.py): the German and English info types, German Excel's own
// format codes for ZELLE("Format") (F2, .2 for #.##0,00, W0 for currency, U4 for hh:mm, D1 to D3).

import { colName } from "../address";
import { Ctx, def, deref } from "./eval";
import { Arr, E, Ref, isErr, toStr, type Any, type V } from "./values";

/** German Excel's code for a number format (ZELLE("Format")), from the format in the file's (English) form. */
export function cellFormatCode(code: string): string {
  const sections: string[] = [];
  let cur = "", q = false;
  for (let i = 0; i < code.length; i++) {
    const c = code[i];
    if (c === '"') q = !q;
    if (c === "\\" && !q) { cur += c + (code[i + 1] ?? ""); i++; continue; }
    if (c === ";" && !q) { sections.push(cur); cur = ""; continue; }
    cur += c;
  }
  sections.push(cur);
  const plain = (s: string) => s.replace(/"[^"]*"/g, "").replace(/\\./g, "").replace(/_./g, "");
  const pos = plain(sections[0]);
  // A currency symbol counts also in quotes or as [$€-407] (W0, W2: recorded).
  const currency = /[€$£¥]|\[\$[^\]-]/.test(sections[0]);
  const negColor = sections.length > 1 && /\[(red|black|blue|cyan|green|magenta|white|yellow|color\s*\d+)\]/i.test(sections[1]);
  const minus = negColor ? "-" : "";
  const body = pos.replace(/\[[^\]]*\]/g, (m) => (/^\[\$/.test(m) ? "¤" : /^\[h+\]|^\[m+\]|^\[s+\]/i.test(m) ? "°" : ""));
  if (/general/i.test(body) || body.trim() === "" || body.trim() === "@") return "S" + minus;
  const lower = body.toLowerCase();
  if (/[dyhs]/.test(lower) || (/m/.test(lower) && !/[0#?]/.test(lower))) {
    if (lower.includes("°") || /dddd|mmmm/.test(lower)) return "S" + minus;
    const ampm = /am\/pm|a\/p/.test(lower);
    const hasDate = /d|y/.test(lower) || (/m/.test(lower) && !/h|s/.test(lower));
    if (hasDate) {
      const d = /d/.test(lower), y = /y/.test(lower), m = /m/.test(lower);
      if (d && m && y) return "D1" + minus;
      if (d && m) return "D2" + minus;
      if (m && y) return "D3" + minus;
      return "S" + minus;
    }
    const secs = /s/.test(lower);
    return (ampm ? (secs ? "U1" : "U2") : secs ? "U3" : "U4") + minus;
  }
  const dot = body.indexOf(".");
  const decimals = dot < 0 ? 0 : (body.slice(dot + 1).match(/^[0#?]*/)?.[0].length ?? 0);
  if (/\?\s*\/\s*\?|\/\d/.test(body)) return "S" + minus;
  if (body.includes("%")) return "P" + decimals + minus;
  if (/e[+-]/i.test(body)) return "E" + decimals + minus;
  if (currency || /[€$£¥¤]/.test(body)) return "W" + decimals + minus;
  if (/#,##|0,0/.test(body)) return "." + decimals + minus;
  if (/[0#]/.test(body)) return "F" + decimals + minus;
  return "S" + minus;
}

const KINDS: Record<string, string> = {
  adresse: "address", address: "address", spalte: "col", col: "col", farbe: "color", color: "color", inhalt: "contents", contents: "contents",
  dateiname: "filename", filename: "filename", format: "format", klammern: "parentheses", parentheses: "parentheses",
  "präfix": "prefix", prefix: "prefix", schutz: "protect", protect: "protect", zeile: "row", row: "row", typ: "type", type: "type",
  breite: "width", width: "width",
};

/** The first cell of a reference, or of the formula's own cell when there is none. */
function target(ctx: Ctx, a: Any | undefined): { sheet: number; row: number; col: number } | null {
  if (a === undefined) return { sheet: ctx.sheet, row: ctx.row, col: ctx.col };
  if (!(a instanceof Ref)) return null;
  const r = a.areas[0];
  return { sheet: a.sheet, row: r.r0, col: r.c0 };
}

def("CELL", 1, 2, (args, ctx) => {
  const k = deref(args[0], ctx);
  if (isErr(k)) return k;
  const kind = KINDS[toStr(k as V).toString().trim().toLowerCase()];
  if (!kind) return E.value;
  const at = target(ctx, args[1] === null || args[1] === undefined ? undefined : args[1]);
  if (!at) return E.value;
  const wb = ctx.host.wb;
  const sheet = wb.sheets[at.sheet];
  const cell = sheet?.get(at.row, at.col);
  const style = wb.styles.get(cell?.s ?? sheet?.rows.get(at.row)?.s ?? sheet?.cols.get(at.col)?.s);
  const v = ctx.host.value(at.sheet, at.row, at.col);
  const fmt = style.numFmt ?? "General";
  switch (kind) {
    case "address": {
      const a = `$${colName(at.col)}$${at.row + 1}`;
      if (at.sheet === ctx.sheet) return a;
      return `[${wb.docName ?? "Mappe1"}]${sheet.name}!${a}`;
    }
    case "col": return at.col + 1;
    case "row": return at.row + 1;
    case "color": return cellFormatCode(fmt).endsWith("-") ? 1 : 0;
    case "contents": return v ?? 0;
    case "filename": return wb.docPath ? `${wb.docPath.replace(/[^\\/]*$/, "")}[${wb.docPath.replace(/^.*[\\/]/, "")}]${sheet.name}` : "";
    case "format": return cellFormatCode(fmt);
    case "parentheses": return /\(/.test(fmt.split(";")[0].replace(/"[^"]*"/g, "")) ? 1 : 0;
    case "prefix": {
      if (typeof v !== "string") return "";
      const h = style.h;
      return h === "right" ? '"' : h === "center" ? "^" : h === "fill" ? "\\" : "'";
    }
    case "protect": return style.locked === false ? 0 : 1;
    case "type": return v === null ? "b" : typeof v === "string" ? "l" : "w";
    case "width": {
      const w = sheet?.cols.get(at.col)?.width;
      const chars = w ?? sheet?.defaultColWidth ?? 8.43;
      return new Arr(1, 2, [Math.trunc(chars), w === undefined]);
    }
  }
  return E.value;
});

def("INFO", 1, 1, (args, ctx) => {
  const k = deref(args[0], ctx);
  if (isErr(k)) return k;
  const wb = ctx.host.wb;
  switch (toStr(k as V).toString().trim().toLowerCase()) {
    case "dateienzahl": case "numfile": return wb.sheets.length;
    case "system": return "pcdos";
    case "ursprung": case "origin": {
      const v = wb.sheets[ctx.sheet]?.view;
      return `$A:$${colName(v?.left ?? 0)}$${(v?.top ?? 0) + 1}`;
    }
    case "version": case "release": return "16.0";
    case "osversion": return "Windows (64-bit) NT 10.00";
    case "recalc": return wb.calcMode === "manual" ? "Manuell" : "Automatisch";
    case "verzeichnis": case "directory": return wb.docPath ? wb.docPath.replace(/[^\\/]*$/, "") : "";
  }
  return E.value;
});

/** The sheet a value names: a reference's (first) sheet, or a sheet's name. */
def("SHEET", 0, 1, (args, ctx) => {
  if (!args.length || args[0] === null) return ctx.sheet + 1;
  const a = args[0];
  if (a instanceof Ref) return a.sheet + 1;
  const d = deref(a, ctx);
  if (isErr(d)) return d;
  if (typeof d !== "string") return E.na;
  const i = ctx.sheetIndex(d);
  return i < 0 ? E.na : i + 1;
});

def("SHEETS", 0, 1, (args, ctx) => {
  if (!args.length || args[0] === null) return ctx.host.wb.sheets.length;
  const a = args[0];
  if (a instanceof Ref) return (a.sheetEnd ?? a.sheet) - a.sheet + 1;
  const d = deref(a, ctx);
  return isErr(d) ? d : E.value;
});

/** The cell shows the friendly name (or the address); clicking it follows the link (the app's part). */
def("HYPERLINK", 1, 2, (args, ctx) => {
  const link = deref(args[0], ctx);
  if (isErr(link)) return link;
  if (args.length > 1 && args[1] !== null) {
    const f = deref(args[1], ctx);
    return f instanceof Arr ? E.value : f ?? 0;
  }
  return link instanceof Arr ? E.value : link ?? 0;
});

