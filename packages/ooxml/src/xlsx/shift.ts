// Moves the relative references of a formula in file form, as Excel does for the cells of a
// shared formula (<f t="shared" si="…"/>): every reference without "$" moves by the cell's
// distance from the formula's first cell. Strings, sheet names in quotes, structured
// references and function names are left alone. (Kept local to the .xlsx layer; the formula
// engine has its own reference handling.)

import { MAX_COLS, MAX_ROWS, colIndex, colName } from "@raabe/sheet-engine";

const IDENT = /[A-Za-z0-9_.\\]/;
const CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/;
const COLS = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(.!])/;
const ROWS = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![A-Za-z0-9_(.!])/;

/** The formula `f` (file form, no "=") as it reads `dr` rows and `dc` columns further on. */
export function shiftFormula(f: string, dr: number, dc: number): string {
  if (!dr && !dc) return f;
  let out = "";
  let i = 0;
  const n = f.length;
  while (i < n) {
    const ch = f[i];
    if (ch === '"') {
      // String literal ("" inside is a quote).
      let j = i + 1;
      while (j < n && !(f[j] === '"' && f[j + 1] !== '"')) j += f[j] === '"' ? 2 : 1;
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === "'") {
      // Quoted sheet name ('' inside is a quote); the reference after "!" is handled next.
      let j = i + 1;
      while (j < n && !(f[j] === "'" && f[j + 1] !== "'")) j += f[j] === "'" ? 2 : 1;
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === "[") {
      // Structured reference or external workbook index: up to the matching "]".
      let depth = 0, j = i;
      for (; j < n; j++) {
        if (f[j] === "[") depth++;
        else if (f[j] === "]" && --depth === 0) break;
      }
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === "#") {
      // Error literal (#REF!, #N/A …).
      const m = /^#[A-Za-z0-9/]+[!?]?/.exec(f.slice(i));
      const len = m ? m[0].length : 1;
      out += f.slice(i, i + len);
      i += len;
      continue;
    }
    const prev = i > 0 ? f[i - 1] : "";
    if ((ch === "$" || /[A-Za-z0-9]/.test(ch)) && !(prev && IDENT.test(prev))) {
      const rest = f.slice(i);
      let m = COLS.exec(rest);
      if (m && colIndex(m[2]) >= 0 && colIndex(m[2]) < MAX_COLS && colIndex(m[4]) < MAX_COLS) {
        const a = col(m[1], m[2], dc), b = col(m[3], m[4], dc);
        out += a === null || b === null ? "#REF!" : `${a}:${b}`;
        i += m[0].length;
        continue;
      }
      m = ROWS.exec(rest);
      if (m) {
        const a = row(m[1], m[2], dr), b = row(m[3], m[4], dr);
        out += a === null || b === null ? "#REF!" : `${a}:${b}`;
        i += m[0].length;
        continue;
      }
      m = CELL.exec(rest);
      if (m && !(rest.length > m[0].length && /[A-Za-z0-9_(.!]/.test(rest[m[0].length])) && Number(m[4]) >= 1 && colIndex(m[2]) < MAX_COLS) {
        const c = col(m[1], m[2], dc), r = row(m[3], m[4], dr);
        out += c === null || r === null ? "#REF!" : c + r;
        i += m[0].length;
        continue;
      }
      // Any other name or number: copy it whole so its digits are not read as references.
      const w = /^[A-Za-z0-9_.\\$]+/.exec(rest);
      const len = w ? w[0].length : 1;
      out += f.slice(i, i + len);
      i += len;
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}

function col(abs: string, name: string, dc: number): string | null {
  if (abs) return abs + name.toUpperCase();
  const c = colIndex(name) + dc;
  return c < 0 || c >= MAX_COLS ? null : colName(c);
}

function row(abs: string, digits: string, dr: number): string | null {
  if (abs) return abs + digits;
  const r = Number(digits) + dr;
  return r < 1 || r > MAX_ROWS ? null : String(r);
}
