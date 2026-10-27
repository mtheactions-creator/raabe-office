// Cells through the clipboard. Inside RaabeCalc a copy keeps everything (formulas, formats,
// merges) and the source keeps its marching ants until Esc or the next entry, as in Excel.
// Other programs get tab-separated text and an HTML table (Word and Excel read both); from
// them RaabeCalc reads HTML tables (Excel's, Word's, web pages) or tab-separated text.

import { type Rect, type Sheet, type Workbook } from "@raabe/sheet-engine";
import { shown } from "./engine";

export interface Copied {
  sheet: Sheet;
  rect: Rect;
  cut: boolean;
  /** The text put on the system clipboard: when it is still there, a paste uses the cells. */
  text: string;
}

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** Tab-separated text of a range as Excel puts it on the clipboard (what the cells show). */
export function rangeText(wb: Workbook, sheet: Sheet, r: Rect): string {
  const lines: string[] = [];
  for (let row = r.r0; row <= r.r1; row++) {
    const cells: string[] = [];
    for (let col = r.c0; col <= r.c1; col++) {
      const c = sheet.get(row, col);
      const t = shown(c, wb.styles.get(c?.s)).text;
      cells.push(/[\t\n"]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t);
    }
    lines.push(cells.join("\t"));
  }
  return lines.join("\r\n") + "\r\n";
}

/** An HTML table with the cells' look, for Word and other programs. */
export function rangeHtml(wb: Workbook, sheet: Sheet, r: Rect): string {
  const rows: string[] = [];
  for (let row = r.r0; row <= r.r1; row++) {
    const tds: string[] = [];
    for (let col = r.c0; col <= r.c1; col++) {
      const merge = sheet.mergeAt(row, col);
      if (merge && (merge.r0 !== row || merge.c0 !== col)) continue;
      const c = sheet.get(row, col);
      const st = wb.styles.get(c?.s);
      const s = shown(c, st);
      const css = [
        `font-family:${st.font.name}`, `font-size:${st.font.size}pt`, st.font.bold ? "font-weight:700" : "", st.font.italic ? "font-style:italic" : "",
        st.font.underline ? "text-decoration:underline" : "", st.font.color ? `color:${st.font.color}` : "", st.fill ? `background:${st.fill}` : "",
        `text-align:${st.h && st.h !== "general" ? st.h : s.align}`, "white-space:nowrap",
        ...(["top", "right", "bottom", "left"] as const).map((k) => (st.borders[k] ? `border-${k}:${st.borders[k]!.style === "thick" ? 2 : 1}pt solid ${st.borders[k]!.color ?? "#000"}` : "")),
      ].filter(Boolean).join(";");
      const span = merge ? ` rowspan="${merge.r1 - merge.r0 + 1}" colspan="${merge.c1 - merge.c0 + 1}"` : "";
      const num = typeof c?.v === "number" ? ` x:num="${c.v}"` : "";
      // Excel reads the number format from mso-number-format, its characters escaped as it writes them.
      const fmt = st.numFmt !== "General" ? `;mso-number-format:"${msoEscape(st.numFmt)}"` : "";
      tds.push(`<td${span}${num} style="${esc(css + fmt)}">${esc(s.text)}</td>`);
    }
    rows.push(`<tr>${tds.join("")}</tr>`);
  }
  return `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><body><table style="border-collapse:collapse">${rows.join("")}</table></body></html>`;
}

/** How a pasted cell looks (from the HTML's classes and styles): applied after its text is entered. */
export interface PastedLook { bold?: boolean; italic?: boolean; underline?: boolean; color?: string; fill?: string; numFmt?: string }

/** A pasted HTML table or text as rows of cell texts (what would be typed into each cell), with their look. */
export function parseClipboard(html: string, text: string): { text: string; look?: PastedLook }[][] {
  if (html && /<table/i.test(html)) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const table = doc.querySelector("table");
    if (table) {
      const classes = cssClasses(html);
      const out: { text: string; look?: PastedLook }[][] = [];
      const taken = new Map<string, true>();
      table.querySelectorAll("tr").forEach((tr, ri) => {
        const row = out[ri] ?? (out[ri] = []);
        let ci = 0;
        tr.querySelectorAll("td,th").forEach((td) => {
          while (taken.has(`${ri},${ci}`)) ci++;
          const cell = td as HTMLTableCellElement;
          const num = cell.getAttribute("x:num");
          const t = num !== null && num !== "" ? num.replace(".", ",") : (cell.textContent ?? "").replace(/\u00a0/g, " ").trim();
          const css = [...(cell.getAttribute("class") ?? "").split(/\s+/).map((c) => classes.get(c) ?? ""), cell.getAttribute("style") ?? ""].join(";");
          row[ci] = { text: t, look: lookOf(css, cell.tagName === "TH") };
          for (let dr = 0; dr < (cell.rowSpan || 1); dr++) for (let dc = 0; dc < (cell.colSpan || 1); dc++) if (dr || dc) taken.set(`${ri + dr},${ci + dc}`, true);
          ci += cell.colSpan || 1;
        });
      });
      return out.map((r) => Array.from(r, (v) => v ?? { text: "" }));
    }
  }
  return parseTsv(text).map((r) => r.map((t) => ({ text: t })));
}

/** Excel's <style> rules (".xl65 {mso-number-format:0%;}") by class name. */
function cssClasses(html: string): Map<string, string> {
  const out = new Map<string, string>();
  const style = /<style[^>]*>([\s\S]*?)<\/style>/i.exec(html)?.[1] ?? "";
  for (const m of style.matchAll(/\.([\w-]+)\s*\{([^}]*)\}/g)) out.set(m[1], m[2]);
  return out;
}

const NAMED_FORMATS: Record<string, string> = {
  "Short Date": "m/d/yyyy", "Medium Date": "d-mmm-yy", "Long Date": "[$-F800]dddd, mmmm dd, yyyy", "General Date": "m/d/yyyy h:mm",
  "Short Time": "h:mm", "Medium Time": "h:mm AM/PM", "Long Time": "h:mm:ss", Percent: "0.00%", Fixed: "0.00", Standard: "#,##0.00", Scientific: "0.00E+00",
};

function lookOf(css: string, th: boolean): PastedLook | undefined {
  const get = (prop: string) => { const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*([^;]+)`, "gi"); let last: string | undefined; for (const x of css.matchAll(m)) last = x[1].trim(); return last; };
  const look: PastedLook = {};
  const weight = get("font-weight");
  if ((weight && (weight === "bold" || Number(weight) >= 600)) || (th && !weight)) look.bold = true;
  if (get("font-style") === "italic") look.italic = true;
  if (/underline/.test(get("text-decoration") ?? "")) look.underline = true;
  const color = get("color");
  if (color && !/^(black|windowtext|#000000|auto)$/i.test(color)) look.color = cssHex(color);
  const bg = get("background") ?? get("background-color");
  if (bg && !/^(white|#ffffff|transparent|none|auto)$/i.test(bg.split(" ")[0])) look.fill = cssHex(bg.split(" ")[0]);
  const fmt = get("mso-number-format");
  if (fmt && fmt !== "General" && fmt !== '"General"') {
    const raw = fmt.replace(/^"|"$/g, "").replace(/\\(.)/g, "$1");
    look.numFmt = NAMED_FORMATS[raw] ?? raw;
  }
  return Object.keys(look).length ? look : undefined;
}

/** A CSS colour (name, #rgb, rgb()) as "#RRGGBB". */
function cssHex(c: string): string | undefined {
  const ctx = document.createElement("canvas").getContext("2d");
  if (!ctx) return undefined;
  ctx.fillStyle = "#000000";
  ctx.fillStyle = c;
  const v = String(ctx.fillStyle);
  return v.startsWith("#") ? v.toUpperCase() : undefined;
}

/** A number format in mso-number-format's form: its punctuation escaped as Excel writes it. */
function msoEscape(code: string): string {
  // Excel names its locale-dependent built-in formats ("Short Date" is TT.MM.JJJJ in German Excel).
  const named = Object.entries(NAMED_FORMATS).find(([, c]) => c === code);
  if (named && /Date|Time/.test(named[0])) return named[0];
  return code.replace(/"/g, '\\"').replace(/([#,.])/g, "\\$1");
}

export function parseTsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cur = "";
  let quoted = false;
  const t = text.replace(/\r\n?/g, "\n").replace(/\n$/, "");
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (quoted) {
      if (ch === '"' && t[i + 1] === '"') { cur += '"'; i++; }
      else if (ch === '"') quoted = false;
      else cur += ch;
    } else if (ch === '"' && cur === "") quoted = true;
    else if (ch === "\t") { row.push(cur); cur = ""; }
    else if (ch === "\n") { row.push(cur); rows.push(row); row = []; cur = ""; }
    else cur += ch;
  }
  row.push(cur);
  rows.push(row);
  return rows;
}
