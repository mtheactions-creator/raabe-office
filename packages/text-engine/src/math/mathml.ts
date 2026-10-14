// Equations drawn as MathML (Cambria Math, laid out by the browser with the font's MATH
// table), the way Word lays out the same OMML: letters italic (capital Greek upright),
// numbers and function names upright, sums with limits above and below in display equations,
// integrals with limits at the side, lim/max/min with the limit below.
//
// For the formula editor every drawn atom says where it is: `data-a="path|start|count"` (the
// characters start…start+count of the sequence at `path`, or the structure at `start`), every
// slot is an mrow with `data-s="path"`, and an empty slot shows Word's dotted placeholder. A
// path is the steps "i.slot" joined by "/" ("" for the equation itself).

import { mergeText, parseLatex, type MNode, type Seq } from "./ast";

const UPRIGHT_GREEK = /[Α-Ω]/;
const ACCENT_GLYPH: Record<string, string> = {
  "̂": "^", "̃": "~", "̇": "˙", "̈": "¨", "⃗": "→", "̅": "¯",
  "̌": "ˇ", "̆": "˘", "́": "´", "̀": "`", "̊": "˚", "⃖": "←",
};
const SPACE_EM: Record<string, string> = { " ": "0.1667em", " ": "0.2222em", " ": "0.2778em", " ": "1em", " ": "0.5em", " ": "0.0833em" };
const PUNCT = new Set([",", ";", ":", ".", "!", "?"]);
/** Word's placeholder for an empty slot (a dotted square). */
export const PLACEHOLDER = "⬚";
/** Word's text in an empty equation. */
export const EMPTY_EQUATION = "Geben Sie hier eine Formel ein.";

const mesc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const at = (path: string, start: number, count: number) => ` data-a="${path}|${start}|${count}"`;

interface Ctx { display: boolean; placeholders: boolean; smallFrac: boolean }

function mathmlRun(n: Extract<MNode, { k: "t" }>, path: string, start: number): string[] {
  const chars = [...n.s];
  if (n.nor) return [`<mtext${n.sty === "b" ? ' mathvariant="bold"' : ""}${at(path, start, chars.length)}>${mesc(n.s)}</mtext>`];
  const out: string[] = [];
  const variant = n.scr === "double-struck" ? "double-struck" : n.scr === "script" ? "script" : n.scr === "fraktur" ? "fraktur"
    : n.scr === "sans-serif" ? "sans-serif" : n.scr === "monospace" ? "monospace" : n.sty === "b" ? "bold" : n.sty === "bi" ? "bold-italic"
      : n.sty === "p" ? "normal" : "";
  // Upright letters (sty p, e.g. a function name or \mathrm) stay one identifier.
  if (variant === "normal" && /^[A-Za-z]+$/.test(n.s)) return [`<mi${n.s.length === 1 ? ' mathvariant="normal"' : ""}${at(path, start, chars.length)}>${mesc(n.s)}</mi>`];
  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i];
    const a = (count = 1) => at(path, start + i, count);
    if (SPACE_EM[ch]) { out.push(`<mspace width="${SPACE_EM[ch]}"${a()}/>`); continue; }
    // A typed space takes no room (math spacing is automatic) but keeps its place.
    if (ch === " ") { out.push(`<mspace width="0"${a()}/>`); continue; }
    if (/[0-9]/.test(ch)) {
      let num = ch;
      while (i + 1 < chars.length && /[0-9.,]/.test(chars[i + 1]) && /[0-9]/.test(chars[i + 2] ?? "")) num += chars[++i];
      out.push(`<mn${at(path, start + i - [...num].length + 1, [...num].length)}>${num}</mn>`);
      continue;
    }
    if (/\p{L}/u.test(ch) || ch === "∞" || ch === "∂" || ch === "∇" || ch === "∅") {
      const v = variant && variant !== "normal" ? variant : variant === "normal" || UPRIGHT_GREEK.test(ch) || /[∞∂∇∅]/.test(ch) ? "normal" : "";
      out.push(`<mi${v ? ` mathvariant="${v}"` : ""}${a()}>${mesc(ch)}</mi>`);
      continue;
    }
    if (ch === "′") { out.push(`<mo${a()}>′</mo>`); continue; }
    // A typed hyphen is drawn as a minus sign, as Word does.
    const shown = ch === "-" ? "−" : ch;
    out.push(PUNCT.has(ch) ? `<mo separator="true"${a()}>${mesc(shown)}</mo>` : `<mo${a()}>${mesc(shown)}</mo>`);
  }
  return out;
}

/** A slot (or the whole equation): its atoms in an mrow that says which sequence it is. */
function mml(s: Seq | undefined, path: string, c: Ctx, base = 0): string {
  const parts: string[] = [];
  let idx = base;
  for (const n of mergeText(s ?? [])) {
    if (n.k === "t") { parts.push(...mathmlRun(n, path, idx)); idx += [...n.s].length; continue; }
    parts.push(mathmlNode(n, path, idx, c));
    idx++;
  }
  if (!parts.length && c.placeholders) parts.push(`<mi mathvariant="normal" class="mph" data-ph="1"${at(path, base, 0)}>${PLACEHOLDER}</mi>`);
  return `<mrow data-s="${path}">${parts.join("")}</mrow>`;
}

const sub = (path: string, i: number, slot: string) => `${path ? `${path}/` : ""}${i}.${slot}`;

function mathmlNode(n: MNode, path: string, i: number, c: Ctx): string {
  const S = (slot: string, seq: Seq | undefined) => mml(seq, sub(path, i, slot), c);
  const A = at(path, i, 1);
  const display = c.display;
  switch (n.k) {
    case "t": return "";
    case "frac": {
      if (n.type === "lin" || n.type === "skw") return `<mrow${A}>${S("num", n.num)}<mo>/</mo>${S("den", n.den)}</mrow>`;
      // Word keeps fractions in display equations at full size, nested ones too (its default:
      // "Größe von verschachtelten Brüchen für Anzeigeformeln reduzieren" off).
      const full = (x: string) => (display && !c.smallFrac ? x.replace("<mrow ", '<mrow displaystyle="true" ') : x);
      return `<mfrac${n.type === "noBar" ? ' linethickness="0"' : ""}${A}>${full(S("num", n.num))}${full(S("den", n.den))}</mfrac>`;
    }
    case "scr": {
      const b = S("base", n.base);
      if (n.pre) return `<mmultiscripts${A}>${b}<mprescripts/>${n.sub ? S("sub", n.sub) : "<none/>"}${n.sup ? S("sup", n.sup) : "<none/>"}</mmultiscripts>`;
      if (n.sub && n.sup) return `<msubsup${A}>${b}${S("sub", n.sub)}${S("sup", n.sup)}</msubsup>`;
      if (n.sup) return `<msup${A}>${b}${S("sup", n.sup)}</msup>`;
      return `<msub${A}>${b}${S("sub", n.sub)}</msub>`;
    }
    case "rad": return n.deg ? `<mroot${A}>${S("e", n.e)}${S("deg", n.deg)}</mroot>` : `<msqrt${A}>${S("e", n.e)}</msqrt>`;
    case "d": {
      const parts: string[] = [];
      const st = n.noGrow ? 'stretchy="false"' : 'stretchy="true"';
      if (n.beg) parts.push(`<mo fence="true" ${st}>${mesc(n.beg)}</mo>`);
      n.items.forEach((x, k) => { if (k) parts.push(`<mo ${st}>|</mo>`); parts.push(S(`item${k}`, x)); });
      if (n.end) parts.push(`<mo fence="true" ${st}>${mesc(n.end)}</mo>`);
      return `<mrow${A}>${parts.join("")}</mrow>`;
    }
    case "nary": {
      const op = `<mo largeop="true" movablelimits="false">${mesc(n.chr)}</mo>`;
      const under = n.und && display;
      let head = op;
      if (n.sub && n.sup) head = under ? `<munderover>${op}${S("sub", n.sub)}${S("sup", n.sup)}</munderover>` : `<msubsup>${op}${S("sub", n.sub)}${S("sup", n.sup)}</msubsup>`;
      else if (n.sub) head = under ? `<munder>${op}${S("sub", n.sub)}</munder>` : `<msub>${op}${S("sub", n.sub)}</msub>`;
      else if (n.sup) head = under ? `<mover>${op}${S("sup", n.sup)}</mover>` : `<msup>${op}${S("sup", n.sup)}</msup>`;
      // Word's display sums and integrals take 0.08 em more above and below (word-step4.docx).
      if (display) head = `<mrow style="padding-block:0.08em">${head}</mrow>`;
      return `<mrow${A}>${head}${S("e", n.e)}</mrow>`;
    }
    case "func": return `<mrow${A}>${S("name", n.name)}<mo>⁡</mo>${S("e", n.e)}</mrow>`;
    case "acc": return `<mover accent="true"${A}>${S("e", n.e)}<mo stretchy="true">${mesc(ACCENT_GLYPH[n.chr] ?? n.chr)}</mo></mover>`;
    case "bar": return n.top ? `<mover accent="true"${A}>${S("e", n.e)}<mo stretchy="true">‾</mo></mover>` : `<munder accentunder="true"${A}>${S("e", n.e)}<mo stretchy="true">_</mo></munder>`;
    case "lim": return n.upper ? `<mover${A}>${S("e", n.e)}${S("lim", n.lim)}</mover>` : `<munder${A}>${S("e", n.e)}${S("lim", n.lim)}</munder>`;
    case "group": return n.top ? `<mover${A}>${S("e", n.e)}<mo stretchy="true">${n.chr}</mo></mover>` : `<munder${A}>${S("e", n.e)}<mo stretchy="true">${n.chr}</mo></munder>`;
    case "m": {
      const cols = Math.max(1, ...n.rows.map((r) => r.length));
      // Platzhalter ausblenden: empty cells show nothing.
      const cell = (seq: Seq | undefined, key: string) => (n.plcHide ? mml(seq, sub(path, i, key), { ...c, placeholders: false }) : S(key, seq));
      const align = n.baseJc === "top" ? ' align="top"' : n.baseJc === "bot" ? ' align="bottom"' : "";
      const colAlign = n.mcJc?.length ? ` columnalign="${Array.from({ length: cols }, (_x, ci) => n.mcJc![ci] ?? "center").join(" ")}"` : "";
      return `<mtable${align}${colAlign}${A}>${n.rows.map((r, ri) => `<mtr>${Array.from({ length: cols }, (_x, ci) => `<mtd>${cell(r[ci], `cell${ri}_${ci}`)}</mtd>`).join("")}</mtr>`).join("")}</mtable>`;
    }
    // Alignment marks (&) are not shown.
    case "eqArr": return `<mtable${A}>${n.rows.map((r, ri) => `<mtr><mtd>${S(`row${ri}`, r.map((x) => (x.k === "t" ? { ...x, s: x.s.replace(/&/g, "") } : x)))}</mtd></mtr>`).join("")}</mtable>`;
    case "box": return `<mrow style="border:1px solid currentColor;padding:1px 2px"${A}>${S("e", n.e)}</mrow>`;
  }
}

/** MathML for showing an equation (display: alone in its paragraph, in display style). */
export function latexToMathml(latex: string, display: boolean): string {
  return seqToMathml(parseLatex(latex), display);
}

/**
 * MathML for an equation tree. Lines of a display equation (line breaks) are rows of a table;
 * an empty equation shows Word's prompt.
 */
export function seqToMathml(seq: Seq, display: boolean, placeholders = true, smallFrac = false): string {
  const c: Ctx = { display, placeholders, smallFrac };
  const merged = mergeText(seq);
  if (!merged.length && placeholders) return `<math${display ? ' displaystyle="true"' : ""}><mrow data-s=""><mtext class="mph-text"${at("", 0, 0)}>${EMPTY_EQUATION}</mtext></mrow></math>`;
  const rows = lines(merged);
  let body: string;
  if (rows.length > 1) {
    // Lines keep their positions in the one sequence (the line breaks count).
    let offset = 0;
    const trs = rows.map((r) => {
      const inner = mml(r, "", c, offset);
      offset += r.reduce((n, x) => n + (x.k === "t" ? [...x.s].length : 1), 0) + 1;
      return `<mtr><mtd>${inner}</mtd></mtr>`;
    });
    body = `<mtable columnalign="center" data-s="">${trs.join("")}</mtable>`;
  } else body = mml(merged, "", c);
  return `<math${display ? ' displaystyle="true"' : ""}>${body}</math>`;
}

/** Splits a top-level sequence at line breaks (several lines → an equation array). */
export function lines(s: Seq): Seq[] {
  const out: Seq[] = [[]];
  for (const n of s) {
    if (n.k === "t" && n.s === "\n") out.push([]);
    else out[out.length - 1].push(n);
  }
  return out;
}
