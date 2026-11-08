// R1C1 references as Excel's macro recorder writes formulas (ActiveCell.FormulaR1C1 = "=SUM(R[-2]C:R[-1]C)", D-119):
// relative parts in brackets from the cell that holds the formula, absolute parts as numbers.

import { colName } from "../address";
import { lex, type RefEnd } from "./lexer";
import { sheetPrefix } from "./convert";

const rowPart = (e: RefEnd, at: number) => (e.rowAbs ? `R${e.row! + 1}` : e.row === at ? "R" : `R[${e.row! - at}]`);
const colPart = (e: RefEnd, at: number) => (e.colAbs ? `C${e.col! + 1}` : e.col === at ? "C" : `C[${e.col! - at}]`);

function endR1C1(e: RefEnd, row: number, col: number): string {
  return (e.row !== undefined ? rowPart(e, row) : "") + (e.col !== undefined ? colPart(e, col) : "");
}

/** "SUM(A2:A3)" (file form, English) in cell (row, col) → "SUM(R[-2]C:R[-1]C)". */
export function toR1C1(file: string, row: number, col: number): string {
  const toks = lex(file.startsWith("=") ? file.slice(1) : file, "file");
  let out = "";
  for (const t of toks) {
    if (t.k !== "ref" || !t.ref) { out += t.text; continue; }
    const r = t.ref;
    const pre = r.sheet === undefined ? "" : r.sheet2 !== undefined ? sheetPrefix(r.sheet + ":" + r.sheet2, [r.sheet, r.sheet2]) : sheetPrefix(r.sheet);
    const a = endR1C1(r.a, row, col);
    // A whole row or column that is one row or column: Excel writes it once (R[1], C2).
    const b = r.b ? endR1C1(r.b, row, col) : null;
    out += pre + (b === null || (b === a && (r.a.row === undefined || r.a.col === undefined)) ? a : `${a}:${b}`);
  }
  return out;
}

const REF = /R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?|R(\[-?\d+\]|\d+)|C(\[-?\d+\]|\d+)/y;

/** Excel's R1C1 formula text for the cell (row, col) → A1 form ("=SUM(R[-2]C:R[-1]C)" in A4 → "=SUM(A2:A3)"). */
export function fromR1C1(text: string, row: number, col: number): string {
  let out = "";
  let i = 0;
  const part = (p: string | undefined, at: number) => (p === undefined ? at : p.startsWith("[") ? at + Number(p.slice(1, -1)) : Number(p) - 1);
  const abs = (p: string | undefined) => p !== undefined && !p.startsWith("[");
  while (i < text.length) {
    const ch = text[i];
    if (ch === '"') {
      const end = text.indexOf('"', i + 1);
      let j = end < 0 ? text.length : end + 1;
      while (text[j] === '"') { const k = text.indexOf('"', j + 1); j = k < 0 ? text.length : k + 1; }
      out += text.slice(i, j);
      i = j;
      continue;
    }
    if ((ch === "R" || ch === "C" || ch === "r" || ch === "c") && !/[\p{L}\p{N}_.$]/u.test(text[i - 1] ?? "")) {
      REF.lastIndex = 0;
      const sub = text.slice(i).toUpperCase();
      const m = REF.exec(sub);
      const after = m ? sub[m[0].length] ?? "" : "";
      if (m && !/[\p{L}\p{N}_.(]/u.test(after)) {
        const prevColon = out.endsWith(":"), nextColon = after === ":";
        if (m[3] !== undefined) {
          // A whole row: "R[1]" alone is the row on its own (5:5).
          const r = part(m[3], row) + 1, t = `${abs(m[3]) ? "$" : ""}${r}`;
          out += prevColon || nextColon ? t : `${t}:${t}`;
        } else if (m[4] !== undefined) {
          const c = colName(part(m[4], col)), t = `${abs(m[4]) ? "$" : ""}${c}`;
          out += prevColon || nextColon ? t : `${t}:${t}`;
        } else {
          out += `${abs(m[2]) ? "$" : ""}${colName(part(m[2], col))}${abs(m[1]) ? "$" : ""}${part(m[1], row) + 1}`;
        }
        i += m[0].length;
        continue;
      }
    }
    out += ch;
    i++;
  }
  return out;
}
