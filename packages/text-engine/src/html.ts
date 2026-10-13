// Formatted copy and paste with other programs (Word, browsers): the selection as HTML on
// the clipboard, and HTML from the clipboard as paragraphs with character formatting,
// headings, lists and tables. Word's own HTML ("mso-" styles, list markers) is understood.

import { nextTableId } from "./doc";
import { EMPTY, OBJ, intern, makePara, type CellDef, type CharProps, type Para, type ParaProps, type RowDef, type TableDef } from "./model";
import type { StyleSheet } from "./styles";
import { HIGHLIGHT } from "./view";

const escHtml = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** Paragraphs as HTML with inline styles (what Word reads when pasting). */
export function toHtml(paras: readonly Para[], styles: StyleSheet): string {
  const out: string[] = [];
  const para = (p: Para) => {
    const rp = styles.resolvePara(p.props);
    const lvl = styles.outlineLevel(rp.style);
    const tag = lvl !== undefined && lvl < 6 ? `h${lvl + 1}` : "p";
    const css = [`margin:${rp.before}pt 0 ${rp.after}pt ${rp.indLeft}pt`];
    if (rp.indFirst) css.push(`text-indent:${rp.indFirst}pt`);
    if (rp.align !== "left") css.push(`text-align:${rp.align === "both" ? "justify" : rp.align}`);
    if (rp.lineRule === "auto" && rp.line !== 1) css.push(`line-height:${Math.round(rp.line * 100)}%`);
    let inner = "";
    let pos = 0, k = 0;
    for (const r of p.runs) {
      const c = styles.resolveChar(rp.style, r.props);
      const s: string[] = [`font-family:'${c.font}'`, `font-size:${c.size}pt`];
      if (c.bold) s.push("font-weight:bold");
      if (c.italic) s.push("font-style:italic");
      const deco = [c.underline !== "none" ? "underline" : "", c.strike || c.dstrike ? "line-through" : ""].filter(Boolean);
      if (deco.length) s.push(`text-decoration:${deco.join(" ")}`);
      if (c.color !== "auto") s.push(`color:#${c.color}`);
      if (c.highlight !== "none" && HIGHLIGHT[c.highlight]) s.push(`background:${HIGHLIGHT[c.highlight]}`);
      if (c.vertAlign !== "baseline") s.push(`vertical-align:${c.vertAlign === "superscript" ? "super" : "sub"}`);
      let text = "";
      for (let i = pos; i < pos + r.len; i++) {
        const ch = p.text[i];
        if (ch === OBJ) { text += escHtml(p.objects?.[k++]?.display ?? ""); continue; }
        text += ch === "\t" ? "&emsp;" : ch === "\u000b" ? "<br>" : ch === "\u000c" || ch === "\u000e" ? "" : escHtml(ch);
      }
      pos += r.len;
      if (text) inner += `<span style="${s.join(";")}">${text}</span>`;
    }
    return `<${tag} style="${css.join(";")}">${inner || "&nbsp;"}</${tag}>`;
  };
  for (let i = 0; i < paras.length; i++) {
    const p = paras[i];
    if (p.block) continue;
    if (p.cell) {
      let j = i;
      while (j + 1 < paras.length && paras[j + 1].cell?.table.id === p.cell.table.id) j++;
      const def = p.cell.table;
      let html = '<table style="border-collapse:collapse">';
      def.rows.forEach((row, r) => {
        html += "<tr>";
        row.cells.forEach((cd, c) => {
          if (cd.vMerge === "continue") return;
          const cellParas = paras.slice(i, j + 1).filter((x) => x.cell!.row === r && x.cell!.col === c);
          const span = cd.span > 1 ? ` colspan="${cd.span}"` : "";
          html += `<td${span} style="border:solid windowtext 1pt;padding:0 5.4pt;vertical-align:top">${cellParas.map(para).join("")}</td>`;
        });
        html += "</tr>";
      });
      out.push(`${html}</table>`);
      i = j;
      continue;
    }
    out.push(para(p));
  }
  return `<html><body><!--StartFragment-->${out.join("")}<!--EndFragment--></body></html>`;
}

// ----- reading HTML ---------------------------------------------------------------------------

const BLOCK = new Set(["P", "DIV", "H1", "H2", "H3", "H4", "H5", "H6", "LI", "BLOCKQUOTE", "PRE", "TABLE", "UL", "OL", "TR", "TD", "TH", "BODY", "SECTION", "ARTICLE"]);

const pt = (v: string): number | undefined => {
  const m = /^(-?[\d.]+)(pt|px|em|%)?$/.exec(v.trim());
  if (!m) return undefined;
  const n = Number(m[1]);
  return m[2] === "px" ? n * 0.75 : m[2] === "em" ? n * 11 : m[2] === "%" ? undefined : n;
};

const NAMED: Record<string, string> = { black: "000000", red: "FF0000", blue: "0000FF", green: "008000", yellow: "FFFF00", white: "FFFFFF", gray: "808080", grey: "808080" };
function hexColor(v: string): string | undefined {
  v = v.trim().toLowerCase();
  if (NAMED[v]) return NAMED[v];
  let m = /^#([0-9a-f]{6})$/.exec(v);
  if (m) return m[1].toUpperCase();
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(v);
  if (m) return (m[1] + m[1] + m[2] + m[2] + m[3] + m[3]).toUpperCase();
  m = /^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(v);
  if (m) return [m[1], m[2], m[3]].map((x) => Number(x).toString(16).padStart(2, "0")).join("").toUpperCase();
  return undefined;
}

const HIGHLIGHT_BY_HEX = Object.fromEntries(Object.entries(HIGHLIGHT).map(([k, v]) => [v.slice(1).toUpperCase(), k]));

/** Character formatting from an element's tag and inline style. */
function charFrom(el: Element, base: CharProps): CharProps {
  const c: Record<string, unknown> = { ...base };
  const tag = el.tagName;
  if (tag === "B" || tag === "STRONG") c.bold = true;
  if (tag === "I" || tag === "EM") c.italic = true;
  if (tag === "U") c.underline = "single";
  if (tag === "S" || tag === "STRIKE" || tag === "DEL") c.strike = true;
  if (tag === "SUP") c.vertAlign = "superscript";
  if (tag === "SUB") c.vertAlign = "subscript";
  const st = (el as HTMLElement).style;
  if (!st) return c as CharProps;
  if (st.fontWeight) c.bold = st.fontWeight === "bold" || Number(st.fontWeight) >= 600;
  if (st.fontStyle) c.italic = st.fontStyle === "italic";
  if (st.textDecoration || st.textDecorationLine) {
    const d = st.textDecorationLine || st.textDecoration;
    if (d.includes("underline")) c.underline = "single";
    if (d.includes("line-through")) c.strike = true;
  }
  if (st.fontFamily) {
    const f = st.fontFamily.split(",")[0].replace(/["']/g, "").trim();
    if (f && !/^(serif|sans-serif|monospace|inherit|initial)$/i.test(f)) { c.font = f; c.fontTheme = undefined; }
  }
  if (st.fontSize) { const s = pt(st.fontSize); if (s && s > 1 && s < 400) c.size = Math.round(s * 2) / 2; }
  if (st.color) { const x = hexColor(st.color); if (x) c.color = x === "000000" ? undefined : x; }
  const bg = st.backgroundColor || st.background;
  if (bg) { const x = hexColor(bg); if (x && HIGHLIGHT_BY_HEX[x]) c.highlight = HIGHLIGHT_BY_HEX[x]; }
  if (st.verticalAlign === "super") c.vertAlign = "superscript";
  if (st.verticalAlign === "sub") c.vertAlign = "subscript";
  return c as CharProps;
}

interface Builder {
  text: string;
  runs: { len: number; props: CharProps }[];
  props: ParaProps;
  list?: "bullet" | "number";
  level?: number;
}

/**
 * Paragraphs from HTML. `styleFor` maps built-in names ("heading 1") to this document's style
 * ids; lists come back as `list` markers the caller turns into numbering.
 */
export function fromHtml(html: string, styleFor: (name: string) => string | undefined):
  { paras: (Para & { list?: "bullet" | "number"; level?: number })[] } {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const out: (Para & { list?: "bullet" | "number"; level?: number })[] = [];
  let cur: Builder | null = null;
  const start = (props: ParaProps = {}, list?: "bullet" | "number", level?: number) => { finish(); cur = { text: "", runs: [], props, list, level }; };
  const finish = () => {
    const b = cur;
    cur = null;
    if (!b) return;
    const text = b.text.replace(/\s+$/, "");
    let len = 0;
    const runs = [];
    for (const r of b.runs) { const l = Math.min(r.len, text.length - len); if (l > 0) { runs.push({ len: l, props: intern(r.props) }); len += l; } }
    out.push({ ...makePara(""), text, runs, props: b.props, mark: EMPTY, list: b.list, level: b.level });
  };
  const add = (s: string, props: CharProps) => {
    if (!s) return;
    if (!cur) start();
    const b = cur!;
    // Whitespace collapses as in HTML; text at a paragraph start loses its leading spaces.
    let t = s.replace(/[\r\n\t ]+/g, " ");
    if (!b.text || b.text.endsWith(" ")) t = t.replace(/^ /, "");
    if (!t) return;
    b.text += t;
    const last = b.runs[b.runs.length - 1];
    if (last && last.props === props) last.len += t.length; else b.runs.push({ len: t.length, props });
  };
  const listDepth = (el: Element) => { let d = 0; for (let e = el.parentElement; e; e = e.parentElement) if (e.tagName === "UL" || e.tagName === "OL") d++; return Math.max(0, d - 1); };

  const walk = (node: Node, props: CharProps) => {
    if (node.nodeType === Node.TEXT_NODE) { add(node.textContent ?? "", props); return; }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const el = node as HTMLElement;
    const tag = el.tagName;
    if (tag === "STYLE" || tag === "SCRIPT" || tag === "HEAD" || tag === "TITLE" || tag === "META") return;
    // Word's list markers ("1.", "·") are spans with mso-list:Ignore.
    if ((el.getAttribute("style") ?? "").includes("mso-list:Ignore")) return;
    if (tag === "BR") { if (cur) { cur.text += "\u000b"; const l = cur.runs[cur.runs.length - 1]; if (l) l.len += 1; else cur.runs.push({ len: 1, props: intern(props) }); } return; }
    if (tag === "TABLE") { finish(); table(el); return; }
    if (BLOCK.has(tag) && tag !== "UL" && tag !== "OL" && tag !== "TR" && tag !== "TD" && tag !== "TH" && tag !== "BODY") {
      const pp: Record<string, unknown> = {};
      const m = /^H([1-6])$/.exec(tag);
      if (m) pp.style = styleFor(`heading ${m[1]}`);
      const st = el.style;
      if (st.textAlign) pp.align = st.textAlign === "center" ? "center" : st.textAlign === "right" ? "right" : st.textAlign === "justify" ? "both" : undefined;
      const cls = el.getAttribute("class") ?? "";
      const msoList = /mso-list:\s*l\d+\s+level(\d+)/.exec(el.getAttribute("style") ?? "");
      let list: "bullet" | "number" | undefined, level: number | undefined;
      if (tag === "LI") { list = el.parentElement?.tagName === "OL" ? "number" : "bullet"; level = listDepth(el); }
      else if (msoList || /MsoListParagraph/.test(cls)) {
        const marker = el.querySelector("[style*='mso-list:Ignore']")?.textContent?.trim() ?? "";
        list = /^\d+[.)]|^[a-z][.)]|^[ivx]+[.)]/i.test(marker) ? "number" : "bullet";
        level = msoList ? Number(msoList[1]) - 1 : 0;
      }
      start(pp as ParaProps, list, level);
      const cp = charFrom(el, props);
      for (const ch of Array.from(el.childNodes)) walk(ch, intern(cp));
      finish();
      return;
    }
    const cp = intern(charFrom(el, props));
    for (const ch of Array.from(el.childNodes)) walk(ch, cp);
  };

  const table = (t: HTMLElement) => {
    const trs = Array.from(t.querySelectorAll(":scope > tr, :scope > tbody > tr, :scope > thead > tr"));
    if (!trs.length) return;
    const width = 453.6;
    const cols = Math.max(...trs.map((tr) => Array.from(tr.children).reduce((s, td) => s + Math.max(1, Number(td.getAttribute("colspan") ?? 1)), 0)));
    const grid = Array.from({ length: cols }, () => width / cols);
    const rows: RowDef[] = trs.map((tr) => ({ cells: Array.from(tr.children).map((td): CellDef => ({ span: Math.max(1, Number(td.getAttribute("colspan") ?? 1)) })) }));
    const def: TableDef = { id: nextTableId(), grid, rows, style: styleFor("Table Grid"), look: { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true } };
    trs.forEach((tr, r) => Array.from(tr.children).forEach((td, c) => {
      const before = out.length;
      for (const ch of Array.from(td.childNodes)) walk(ch, EMPTY);
      finish();
      if (out.length === before) out.push({ ...makePara("") });
      for (let i = before; i < out.length; i++) out[i] = { ...out[i], cell: { table: def, row: r, col: c } };
    }));
  };

  const body = doc.body;
  // Only the copied part (Word and browsers mark it with StartFragment comments).
  walk(body, EMPTY);
  finish();
  // Empty paragraphs from the HTML's layout whitespace at the ends go.
  while (out.length > 1 && !out[out.length - 1].text && !out[out.length - 1].cell) out.pop();
  while (out.length > 1 && !out[0].text && !out[0].cell) out.shift();
  return { paras: out };
}
