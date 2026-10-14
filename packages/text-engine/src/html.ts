// Formatted copy and paste with other programs (Word, browsers): the selection as HTML on
// the clipboard, and HTML from the clipboard as paragraphs with character formatting,
// headings, lists and tables. Lists and tables are written the way Word writes them
// ("mso-list" paragraphs with @list definitions, tables with column widths), so Word keeps
// their indents, levels and widths; Word's own HTML is read back the same way.

import { nextTableId } from "./doc";
import { EMPTY, OBJ, intern, makePara, type CellDef, type CharProps, type NumLevel, type Numbering, type Para, type ParaProps, type RowDef, type TableDef } from "./model";
import { levelOf, numberingOf, type ListLabel } from "./numbering";
import type { StyleSheet } from "./styles";
import { HIGHLIGHT } from "./view";

const escHtml = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const n1 = (x: number) => (Math.round(x * 10) / 10).toFixed(1);

const WORD_FORMAT: Record<string, string> = {
  bullet: "bullet", lowerLetter: "alpha-lower", upperLetter: "alpha-upper", lowerRoman: "roman-lower", upperRoman: "roman-upper",
  none: "none", decimalZero: "arabic-leading-zero",
};
const OUR_FORMAT: Record<string, string> = Object.fromEntries(Object.entries(WORD_FORMAT).map(([k, v]) => [v, k]));

/** A level text for Word's CSS: "%1." → "\"%1\\.\"" (bullets: the character itself). */
function cssLevelText(l: NumLevel): string {
  if (l.format === "bullet") return l.text;
  return `"${l.text.replace(/[^%\w]/g, (c) => `\\${c}`)}"`;
}

/** The symbol font characters U+F0xx as Word's HTML shows them (the low byte). */
const htmlBullet = (text: string) => text.replace(/[\uf020-\uf0ff]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xf000));

/**
 * Paragraphs as HTML with inline styles (what Word reads when pasting). `labels` are the list
 * labels of these paragraphs as the document shows them.
 */
export function toHtml(paras: readonly Para[], styles: StyleSheet, numbering?: Numbering, labels?: readonly (ListLabel | undefined)[]): string {
  const out: string[] = [];
  const lists = new Map<number, number>(); // numId → l index
  const listCss: string[] = [];
  const listId = (numId: number): number => {
    let id = lists.get(numId);
    if (id !== undefined) return id;
    id = lists.size;
    lists.set(numId, id);
    const levels = numbering ? Array.from({ length: 9 }, (_, k) => levelOf(numbering, numId, k)?.level) : [];
    levels.forEach((l, k) => {
      if (!l) return;
      const css = [
        l.format !== "decimal" ? `mso-level-number-format:${WORD_FORMAT[l.format] ?? "decimal"}` : "",
        `mso-level-text:${cssLevelText(l)}`,
        l.start !== 1 ? `mso-level-start-at:${l.start}` : "",
        "mso-level-tab-stop:none",
        `mso-level-number-position:${l.align}`,
        l.suffix !== "tab" ? `mso-level-suffix:${l.suffix === "space" ? "space" : "none"}` : "",
        `margin-left:${n1(l.pPr.indLeft ?? 0)}pt`,
        `text-indent:${n1(l.pPr.indFirst ?? 0)}pt`,
        l.rPr.font ? `font-family:"${l.rPr.font}"` : "",
      ].filter(Boolean);
      listCss.push(`@list l${id}:level${k + 1}\n\t{${css.join(";\n\t")};}`);
    });
    return id;
  };

  const runs = (p: Para, rpStyle: string) => {
    let inner = "";
    let pos = 0, k = 0;
    for (const r of p.runs) {
      const c = styles.resolveChar(rpStyle, r.props);
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
        text += ch === "\t" ? "<span style='mso-tab-count:1'>&emsp;</span>" : ch === "\u000b" ? "<br>" : ch === "\u000c" || ch === "\u000e" ? "" : escHtml(ch);
      }
      pos += r.len;
      if (text) inner += `<span style="${s.join(";")}">${text}</span>`;
    }
    return inner;
  };

  const para = (p: Para, i: number, inCell: boolean) => {
    const n = numberingOf(p.props, styles);
    const label = labels?.[i];
    const lvl = n && numbering ? levelOf(numbering, n.numId, n.ilvl)?.level : undefined;
    const rp = styles.resolvePara(p.props, lvl ? { level: lvl.pPr } : undefined);
    const outline = styles.outlineLevel(rp.style);
    const tag = outline !== undefined && outline < 6 ? `h${outline + 1}` : "p";
    const css = [`margin:${n1(rp.before)}pt 0pt ${n1(rp.after)}pt ${n1(rp.indLeft)}pt`];
    if (rp.indFirst) css.push(`text-indent:${n1(rp.indFirst)}pt`);
    if (rp.align !== "left") css.push(`text-align:${rp.align === "both" ? "justify" : rp.align}`);
    if (rp.lineRule === "auto" && rp.line !== 1) css.push(`line-height:${Math.round(rp.line * 100)}%`);
    let marker = "";
    let cls = inCell ? "MsoNormal" : "MsoNormal";
    if (n && lvl) {
      css.push(`mso-list:l${listId(n.numId)} level${n.ilvl + 1} lfo${n.numId}`);
      cls = "MsoListParagraph";
      const text = htmlBullet(label?.text ?? (lvl.format === "bullet" ? lvl.text : ""));
      const font = lvl.rPr.font ? `font-family:"${lvl.rPr.font}"` : "";
      marker = `<![if !supportLists]><span style='${font}'><span style='mso-list:Ignore'>${escHtml(text)}` +
        `<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><![endif]>`;
    }
    return `<${tag} class=${cls} style="${css.join(";")}">${marker}${runs(p, rp.style) || "&nbsp;"}</${tag}>`;
  };

  for (let i = 0; i < paras.length; i++) {
    const p = paras[i];
    if (p.block) continue;
    if (p.cell) {
      let j = i;
      while (j + 1 < paras.length && paras[j + 1].cell?.table.id === p.cell.table.id) j++;
      const def = p.cell.table;
      const grid = def.grid;
      let html = "<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none;mso-yfti-tbllook:1184'>";
      def.rows.forEach((row, r) => {
        const hgt = row.height ? ` style='height:${n1(row.height)}pt${row.heightRule === "exact" ? ";mso-height-rule:exactly" : ""}'` : "";
        html += `<tr${hgt}>`;
        let gx = 0;
        row.cells.forEach((cd, c) => {
          const span = Math.max(1, cd.span);
          const w = cd.width ?? grid.slice(gx, gx + span).reduce((a, b) => a + b, 0);
          gx += span;
          if (cd.vMerge === "continue") return;
          let rows = 1;
          if (cd.vMerge === "restart") while (def.rows[r + rows]?.cells[c]?.vMerge === "continue") rows++;
          const cellParas: [Para, number][] = [];
          for (let k = i; k <= j; k++) if (paras[k].cell!.row === r && paras[k].cell!.col === c) cellParas.push([paras[k], k]);
          const attrs = `${span > 1 ? ` colspan=${span}` : ""}${rows > 1 ? ` rowspan=${rows}` : ""} width=${Math.round(w / 0.75)} valign=${cd.vAlign === "center" ? "middle" : cd.vAlign ?? "top"}`;
          const bg = cd.shading?.fill && cd.shading.fill !== "auto" ? `;background:#${cd.shading.fill}` : "";
          html += `<td${attrs} style='width:${n1(w)}pt;border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt${bg}'>${cellParas.map(([x, k]) => para(x, k, true)).join("")}</td>`;
        });
        html += "</tr>";
      });
      out.push(`${html}</table>`);
      i = j;
      continue;
    }
    out.push(para(p, i, false));
  }
  const head = listCss.length ? `<head><meta charset="utf-8"><style>\n<!--\n${listCss.join("\n")}\n-->\n</style></head>` : `<head><meta charset="utf-8"></head>`;
  return `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">${head}` +
    `<body><!--StartFragment-->${out.join("")}<!--EndFragment--></body></html>`;
}

// ----- reading HTML ---------------------------------------------------------------------------

const BLOCK = new Set(["P", "DIV", "H1", "H2", "H3", "H4", "H5", "H6", "LI", "BLOCKQUOTE", "PRE", "TABLE", "UL", "OL", "TR", "TD", "TH", "BODY", "SECTION", "ARTICLE"]);

/** A CSS length in points (pt, px, cm, mm, in, em). */
const pt = (v: string): number | undefined => {
  const m = /^(-?[\d.]+)\s*(pt|px|em|%|cm|mm|in)?$/.exec(v.trim());
  if (!m) return undefined;
  const n = Number(m[1]);
  switch (m[2]) {
    case "px": return n * 0.75;
    case "em": return n * 11;
    case "cm": return n * 28.3465;
    case "mm": return n * 2.83465;
    case "in": return n * 72;
    case "%": return undefined;
    default: return n;
  }
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
  listKey?: string;
}

/** A list level as Word's HTML defines it (@list l0:level1 {…}). */
export interface HtmlLevel {
  format?: string;
  text?: string;
  start?: number;
  align?: "left" | "center" | "right";
  suffix?: "tab" | "space" | "nothing";
  font?: string;
  indLeft?: number;
  indFirst?: number;
}

export type HtmlPara = Para & { list?: "bullet" | "number"; level?: number; listKey?: string };

/** Declarations of a CSS block: "a:b; c:d" → { a: "b", c: "d" }. */
function decls(body: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const part of body.split(";")) {
    const i = part.indexOf(":");
    if (i > 0) out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim();
  }
  return out;
}

/** Word's style sheet in the clipboard HTML: list levels and paragraph classes. */
function wordStyles(doc: Document): { levels: Map<string, HtmlLevel[]>; classes: Map<string, Record<string, string>> } {
  const levels = new Map<string, HtmlLevel[]>();
  const classes = new Map<string, Record<string, string>>();
  const css = Array.from(doc.querySelectorAll("style")).map((s) => s.textContent ?? "").join("\n").replace(/<!--|-->/g, "").replace(/\/\*[\s\S]*?\*\//g, "");
  for (const m of css.matchAll(/@list\s+(l\d+):level(\d+)\s*\{([^}]*)\}/g)) {
    const d = decls(m[3]);
    const lv: HtmlLevel = {};
    const fmt = d["mso-level-number-format"];
    lv.format = fmt ? OUR_FORMAT[fmt] ?? (fmt === "image" ? "bullet" : "decimal") : "decimal";
    const text = d["mso-level-text"];
    if (text !== undefined) lv.text = text.replace(/^"(.*)"$/, "$1").replace(/\\(.)/g, "$1");
    if (d["mso-level-start-at"]) lv.start = Number(d["mso-level-start-at"]);
    const pos = d["mso-level-number-position"];
    if (pos === "right" || pos === "center" || pos === "left") lv.align = pos;
    const suf = d["mso-level-suffix"];
    lv.suffix = suf === "space" ? "space" : suf === "none" ? "nothing" : "tab";
    if (d["font-family"]) lv.font = d["font-family"].split(",")[0].replace(/["']/g, "").trim();
    if (d["margin-left"]) lv.indLeft = pt(d["margin-left"]);
    if (d["text-indent"]) lv.indFirst = pt(d["text-indent"]);
    const list = levels.get(m[1]) ?? [];
    list[Number(m[2]) - 1] = lv;
    levels.set(m[1], list);
  }
  for (const m of css.matchAll(/([^{}@]+)\{([^}]*)\}/g)) {
    for (const sel of m[1].split(",")) {
      const k = /^\s*(?:p|li|div)\.([\w-]+)\s*$/.exec(sel);
      if (k) classes.set(k[1], { ...classes.get(k[1]), ...decls(m[2]) });
    }
  }
  return { levels, classes };
}

/**
 * Paragraphs from HTML. `styleFor` maps built-in names ("heading 1") to this document's style
 * ids; lists come back as `list`/`level` markers with `listKey` naming a list of `lists`
 * (Word's own list definitions: symbols, fonts and indents), which the caller turns into
 * numbering.
 */
export function fromHtml(html: string, styleFor: (name: string) => string | undefined):
  { paras: HtmlPara[]; lists: Map<string, HtmlLevel[]> } {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const word = wordStyles(doc);
  const lists = new Map<string, HtmlLevel[]>();
  const out: HtmlPara[] = [];
  let cur: Builder | null = null;
  const start = (props: ParaProps = {}, list?: "bullet" | "number", level?: number, listKey?: string) => { finish(); cur = { text: "", runs: [], props, list, level, listKey }; };
  const finish = () => {
    const b = cur;
    cur = null;
    if (!b) return;
    const text = b.text.replace(/\s+$/, "");
    let len = 0;
    const runs = [];
    for (const r of b.runs) { const l = Math.min(r.len, text.length - len); if (l > 0) { runs.push({ len: l, props: intern(r.props) }); len += l; } }
    out.push({ ...makePara(""), text, runs, props: b.props, mark: EMPTY, list: b.list, level: b.level, listKey: b.listKey });
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
  const addRaw = (s: string, props: CharProps) => {
    if (!cur) start();
    const b = cur!;
    b.text += s;
    const last = b.runs[b.runs.length - 1];
    if (last && last.props === props) last.len += s.length; else b.runs.push({ len: s.length, props });
  };
  const listDepth = (el: Element) => { let d = 0; for (let e = el.parentElement; e; e = e.parentElement) if (e.tagName === "UL" || e.tagName === "OL") d++; return Math.max(0, d - 1); };
  /** A paragraph's own CSS, its class's first (Word's MsoListParagraphCxSpMiddle …). */
  const paraCss = (el: HTMLElement): Record<string, string> => {
    const cls = (el.getAttribute("class") ?? "").split(/\s+/).filter(Boolean);
    const merged: Record<string, string> = {};
    for (const c of cls) Object.assign(merged, word.classes.get(c));
    return { ...merged, ...decls(el.getAttribute("style") ?? "") };
  };

  const walk = (node: Node, props: CharProps) => {
    if (node.nodeType === Node.TEXT_NODE) { add(node.textContent ?? "", props); return; }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const el = node as HTMLElement;
    const tag = el.tagName;
    if (tag === "STYLE" || tag === "SCRIPT" || tag === "HEAD" || tag === "TITLE" || tag === "META") return;
    const style = el.getAttribute("style") ?? "";
    // Word's list markers ("1.", "·") are spans with mso-list:Ignore.
    if (style.includes("mso-list:Ignore")) return;
    // Word's tabs.
    if (/mso-tab-count/.test(style)) { addRaw("\t", intern(props)); return; }
    if (tag === "BR") { if (cur) { cur.text += "\u000b"; const l = cur.runs[cur.runs.length - 1]; if (l) l.len += 1; else cur.runs.push({ len: 1, props: intern(props) }); } return; }
    if (tag === "TABLE") { finish(); table(el); return; }
    if (BLOCK.has(tag) && tag !== "UL" && tag !== "OL" && tag !== "TR" && tag !== "TD" && tag !== "TH" && tag !== "BODY") {
      const pp: Record<string, unknown> = {};
      const m = /^H([1-6])$/.exec(tag);
      if (m) pp.style = styleFor(`heading ${m[1]}`);
      const st = el.style;
      if (st.textAlign) pp.align = st.textAlign === "center" ? "center" : st.textAlign === "right" ? "right" : st.textAlign === "justify" ? "both" : undefined;
      const cls = el.getAttribute("class") ?? "";
      const css = paraCss(el);
      const msoList = /l(\d+)\s+level(\d+)(?:\s+lfo(\d+))?/.exec(css["mso-list"] ?? "");
      let list: "bullet" | "number" | undefined, level: number | undefined, listKey: string | undefined;
      if (tag === "LI") { list = el.parentElement?.tagName === "OL" ? "number" : "bullet"; level = listDepth(el); }
      else if (msoList || /MsoListParagraph/.test(cls)) {
        level = msoList ? Number(msoList[2]) - 1 : 0;
        const defs = msoList ? word.levels.get(`l${msoList[1]}`) : undefined;
        const marker = el.querySelector("[style*='mso-list:Ignore']")?.textContent?.trim() ?? "";
        const fmt = defs?.[level]?.format;
        list = fmt ? (fmt === "bullet" ? "bullet" : "number") : /^\(?\d+[.)]|^\(?[a-z][.)]|^\(?[ivx]+[.)]/i.test(marker) ? "number" : "bullet";
        if (msoList && defs) {
          listKey = `l${msoList[1]}|${msoList[3] ?? ""}`;
          const levels = lists.get(listKey) ?? defs.map((d) => (d ? { ...d } : d));
          const lv = levels[level] ?? (levels[level] = {});
          // The paragraph's indent is the level's when the definition has none.
          const left = css["margin-left"] !== undefined ? pt(css["margin-left"]) : undefined;
          const first = css["text-indent"] !== undefined ? pt(css["text-indent"]) : undefined;
          if (lv.indLeft === undefined && left !== undefined) lv.indLeft = left;
          if (lv.indFirst === undefined && first !== undefined) lv.indFirst = first;
          if (lv.text === undefined && list === "bullet" && marker) lv.text = marker;
          lists.set(listKey, levels);
        }
      } else {
        // Indents of ordinary paragraphs.
        const left = css["margin-left"] !== undefined ? pt(css["margin-left"]) : undefined;
        const first = css["text-indent"] !== undefined ? pt(css["text-indent"]) : undefined;
        if (left && Math.abs(left) > 0.05) pp.indLeft = left;
        if (first && Math.abs(first) > 0.05) pp.indFirst = first;
      }
      start(pp as ParaProps, list, level, listKey);
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
    // Grid columns: spans and row spans placed like a browser does.
    const occupied: boolean[][] = trs.map(() => []);
    const placed: { tr: number; td: HTMLElement; col: number; span: number; rows: number }[] = [];
    trs.forEach((tr, r) => {
      let col = 0;
      for (const td of Array.from(tr.children) as HTMLElement[]) {
        if (td.tagName !== "TD" && td.tagName !== "TH") continue;
        while (occupied[r][col]) col++;
        const span = Math.max(1, Number(td.getAttribute("colspan") ?? 1) || 1);
        const rows = Math.max(1, Number(td.getAttribute("rowspan") ?? 1) || 1);
        for (let rr = r; rr < Math.min(trs.length, r + rows); rr++) for (let cc = col; cc < col + span; cc++) occupied[rr][cc] = true;
        placed.push({ tr: r, td, col, span, rows });
        col += span;
      }
    });
    const cols = Math.max(1, ...occupied.map((row) => row.length));
    // Column widths from the cells (Word writes width:…pt and width=… pixels).
    const grid: (number | undefined)[] = new Array(cols).fill(undefined);
    const widthOf = (td: HTMLElement) => {
      const w = decls(td.getAttribute("style") ?? "").width;
      const fromCss = w ? pt(w) : undefined;
      const attr = Number(td.getAttribute("width"));
      return fromCss ?? (attr > 0 ? attr * 0.75 : undefined);
    };
    for (const c of placed) {
      const w = widthOf(c.td);
      if (w === undefined) continue;
      if (c.span === 1) grid[c.col] ??= w;
    }
    for (const c of placed) {
      const w = widthOf(c.td);
      if (w === undefined || c.span === 1) continue;
      const missing = Array.from({ length: c.span }, (_, k) => c.col + k).filter((k) => grid[k] === undefined);
      const known = Array.from({ length: c.span }, (_, k) => grid[c.col + k] ?? 0).reduce((a, b) => a + b, 0);
      for (const k of missing) grid[k] = Math.max(10, (w - known) / missing.length);
    }
    const knownSum = grid.reduce<number>((a, b) => a + (b ?? 0), 0);
    const unknown = grid.filter((x) => x === undefined).length;
    const fill = unknown ? Math.max(20, (453.6 - knownSum) / unknown) : 0;
    const widths = grid.map((x) => x ?? fill);
    const cls = t.getAttribute("class") ?? "";
    const bordered = /MsoTableGrid/.test(cls) || Number(t.getAttribute("border")) > 0 || placed.some((c) => /border[^;]*solid/.test(c.td.getAttribute("style") ?? ""));
    const rows: RowDef[] = trs.map((tr, r) => {
      const st = decls(tr.getAttribute("style") ?? "");
      const hgt = st.height ? pt(st.height) : undefined;
      const cells: CellDef[] = [];
      const inRow = placed.filter((c) => c.tr === r || (c.tr < r && c.tr + c.rows > r));
      inRow.sort((a, b) => a.col - b.col);
      for (const c of inRow) {
        const css = decls(c.td.getAttribute("style") ?? "");
        const bg = css.background ?? css["background-color"];
        const fillHex = bg ? hexColor(bg) : undefined;
        const va = c.td.getAttribute("valign") ?? css["vertical-align"];
        cells.push({
          span: c.span,
          vMerge: c.rows > 1 ? (c.tr === r ? "restart" : "continue") : undefined,
          shading: fillHex && c.tr === r ? { fill: fillHex } : undefined,
          vAlign: va === "middle" || va === "center" ? "center" : va === "bottom" ? "bottom" : undefined,
        });
      }
      return { cells, height: hgt && hgt > 1 ? hgt : undefined, heightRule: hgt && hgt > 1 ? (/exactly/.test(st["mso-height-rule"] ?? "") ? "exact" : "atLeast") : undefined };
    });
    const def: TableDef = {
      id: nextTableId(), grid: widths, rows, style: bordered ? styleFor("Table Grid") : undefined,
      look: { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true },
    };
    rows.forEach((_row, r) => {
      const inRow = placed.filter((c) => c.tr === r || (c.tr < r && c.tr + c.rows > r)).sort((a, b) => a.col - b.col);
      inRow.forEach((c, k) => {
        const before = out.length;
        if (c.tr === r) {
          for (const ch of Array.from(c.td.childNodes)) walk(ch, EMPTY);
          finish();
        }
        if (out.length === before) out.push({ ...makePara("") });
        for (let i = before; i < out.length; i++) out[i] = { ...out[i], cell: { table: def, row: r, col: k } };
      });
    });
  };

  const body = doc.body;
  walk(body, EMPTY);
  finish();
  // Empty paragraphs from the HTML's layout whitespace at the ends go.
  while (out.length > 1 && !out[out.length - 1].text && !out[out.length - 1].cell) out.pop();
  while (out.length > 1 && !out[0].text && !out[0].cell) out.shift();
  return { paras: out, lists };
}
