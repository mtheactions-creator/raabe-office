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
      tds.push(`<td${span}${num} style="${esc(css)}">${esc(s.text)}</td>`);
    }
    rows.push(`<tr>${tds.join("")}</tr>`);
  }
  return `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><body><table style="border-collapse:collapse">${rows.join("")}</table></body></html>`;
}

/** A pasted HTML table or text as rows of cell texts (what would be typed into each cell). */
export function parseClipboard(html: string, text: string): string[][] {
  if (html && /<table/i.test(html)) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const table = doc.querySelector("table");
    if (table) {
      const out: string[][] = [];
      const taken = new Map<string, true>();
      table.querySelectorAll("tr").forEach((tr, ri) => {
        const row: string[] = out[ri] ?? (out[ri] = []);
        let ci = 0;
        tr.querySelectorAll("td,th").forEach((td) => {
          while (taken.has(`${ri},${ci}`)) ci++;
          const cell = td as HTMLTableCellElement;
          // Excel writes numbers as shown and the stored value in x:num.
          const num = cell.getAttribute("x:num");
          const t = num !== null && num !== "" ? num.replace(".", ",") : (cell.textContent ?? "").replace(/\u00a0/g, " ").trim();
          row[ci] = t;
          for (let dr = 0; dr < (cell.rowSpan || 1); dr++) for (let dc = 0; dc < (cell.colSpan || 1); dc++) if (dr || dc) taken.set(`${ri + dr},${ci + dc}`, true);
          ci += cell.colSpan || 1;
        });
      });
      return out.map((r) => Array.from(r, (v) => v ?? ""));
    }
  }
  return parseTsv(text);
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
