// Placing a document on pages: sections (page size, margins, columns, headers and footers,
// page numbers), tables, footnotes and floating pictures. Rules measured against Word 365
// (decisions D-025, D-030):
// - the space between two paragraphs is the larger of "after" and "before"; "before" counts
//   at the top of the document and is dropped at the top of later pages and columns;
// - widow/orphan control, keep with next (chains), keep lines together, page break before;
// - a header taller than the top margin pushes the body down (and a footer up);
// - footnotes sit at the bottom of the page under a separator; a line whose footnote does
//   not fit moves to the next page;
// - columns before a continuous section break are balanced.
// Every line gets a position in page coordinates (px from the page's top left) and the frame
// it belongs to (text column, table cell, header, footer, note), which the view uses for
// drawing and hit testing.

import { PT } from "./fonts";
import { fieldType, type Layouter, type Line, type ParaEnv, type ParaLayout } from "./layout";
import {
  DEFAULT_SECTION, EMPTY, OBJ, cellKey, type BorderSpec, type Borders, type CellMargins, type CharProps, type FloatObject,
  type ImageData, type Numbering, type PageSetup, type Para, type ParaProps, type Section, type Shading, type TableDef,
} from "./model";
import { formatNumber, listLabels, type ListLabel } from "./numbering";
import type { StyleSheet, TableStyleDef, TableStylePart } from "./styles";
import type { DocSettings } from "./doc";

// ----- results -------------------------------------------------------------------------------

export interface Frame {
  readonly id: number;
  readonly page: number;
  /** Layout key of the story: "main", a note id, or a header/footer instance ("hdr:2@3"). */
  readonly story: string;
  readonly kind: "body" | "cell" | "header" | "footer" | "note";
  readonly x: number;
  readonly y: number;
  readonly w: number;
  h: number;
  /** For cells: which cell (cellKey). */
  readonly cell?: string;
}

export interface PlacedLine {
  readonly para: number;
  readonly line: number;
  readonly page: number;
  /** Left edge of the line's frame and top of the line box, px from the page's top left. */
  readonly x: number;
  readonly y: number;
  readonly frame: number;
  /** Drawn but not editable here (repeated table header rows). */
  readonly ghost?: boolean;
  /** Neither drawn nor editable (text of vertically merged cells). */
  readonly hide?: boolean;
}

export interface StoryLayout {
  readonly paras: ParaLayout[];
  readonly placed: PlacedLine[];
  /** Index into `placed` of each paragraph's first editable line. */
  readonly firstPlaced: number[];
  /** Index into `placed` of every editable line, per paragraph and line. */
  readonly index: number[][];
}

export type Deco =
  | { readonly kind: "rect"; readonly story: string; readonly x: number; readonly y: number; readonly w: number; readonly h: number; readonly fill: string }
  | { readonly kind: "line"; readonly story: string; readonly x1: number; readonly y1: number; readonly x2: number; readonly y2: number; readonly width: number; readonly color: string; readonly style: string }
  | { readonly kind: "image"; readonly story: string; readonly x: number; readonly y: number; readonly w: number; readonly h: number; readonly image?: ImageData; readonly behind: boolean; readonly para: number; readonly index: number };

export interface PageLayout {
  readonly index: number;
  readonly setup: PageSetup;
  readonly section: number;
  /** Page number shown (restarts per section where the section says so) and its text. */
  readonly number: number;
  readonly numberText: string;
  /** Header/footer story ids shown on this page and their layout keys. */
  header?: string;
  footer?: string;
  headerKey?: string;
  footerKey?: string;
  readonly deco: Deco[];
  readonly frames: number[];
  bodyTop: number;
  bodyBottom: number;
  /** Note story ids shown at the bottom of this page. */
  readonly notes: string[];
  notesHeight: number;
}

export interface Layout {
  readonly main: StoryLayout;
  /** Other stories: notes by id, header/footer instances by "id@page". */
  readonly stories: ReadonlyMap<string, StoryLayout>;
  readonly pages: readonly PageLayout[];
  readonly frames: readonly Frame[];
  readonly pageCount: number;
  /** The main story's lines (shortcuts). */
  readonly paras: readonly ParaLayout[];
  readonly placed: readonly PlacedLine[];
  readonly firstPlaced: readonly number[];
  /** Number shown for each note reference, by note story id. */
  readonly noteNumbers: ReadonlyMap<string, string>;
}

export interface FlowInput {
  readonly main: readonly Para[];
  readonly stories: ReadonlyMap<string, readonly Para[]>;
  readonly settings: DocSettings;
  readonly numbering: Numbering;
}

// ----- helpers -------------------------------------------------------------------------------

export const visibleBorder = (b: BorderSpec | undefined): b is BorderSpec => !!b && b.style !== "nil" && b.style !== "none" && b.size > 0;
const colorOf = (c: string) => (c === "auto" || !c ? "#000000" : `#${c}`);
const fillOf = (s: Shading | undefined) => (s && s.fill && s.fill !== "auto" ? `#${s.fill}` : null);

/** Word's default cell margins (0.19 cm left and right). */
export const DEFAULT_CELL_MARGINS: CellMargins = { left: 5.4, right: 5.4, top: 0, bottom: 0 };

const newStory = (n: number): StoryLayout => ({ paras: new Array(n), placed: [], firstPlaced: new Array(n).fill(-1), index: [] });

interface Exclusion {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly wrap: string;
  readonly side: string;
}

interface CellGeom {
  readonly x: number;
  readonly w: number;
  readonly contentX: number;
  readonly contentW: number;
  readonly margins: CellMargins;
  readonly paras: number[];
  readonly layouts: ParaLayout[];
  /** Content height including cell margins (px). */
  readonly need: number;
  readonly vMerge?: "restart" | "continue";
  readonly shading: string | null;
  readonly borders: Borders;
  readonly vAlign: "top" | "center" | "bottom";
  /** Rows a restart cell spans (vertical merge). */
  span: number;
}

interface RowGeom {
  readonly cells: CellGeom[];
  height: number;
  readonly header: boolean;
  readonly cantSplit: boolean;
}

interface TableGeom {
  readonly def: TableDef;
  readonly x: number;
  readonly w: number;
  readonly rows: RowGeom[];
}

interface SectionRange {
  readonly sect: Section;
  readonly a: number;
  readonly b: number;
  readonly headers: Section["headers"];
  readonly footers: Section["footers"];
}

// ----- the flow ------------------------------------------------------------------------------

class Flow {
  readonly frames: Frame[] = [];
  readonly pages: PageLayout[] = [];
  readonly stories = new Map<string, StoryLayout>();
  readonly main: StoryLayout;
  readonly noteNumbers = new Map<string, string>();
  private readonly styles: StyleSheet;
  private readonly labels: (ListLabel | undefined)[];
  private readonly sections: SectionRange[] = [];
  private readonly excl = new Map<number, Exclusion[]>();
  private readonly noteLayouts = new Map<string, { paras: ParaLayout[]; height: number; width: number }>();
  private totalPages = "1";

  // Current position in the main story.
  private page!: PageLayout;
  private cols: Frame[] = [];
  private col = 0;
  private y = 0;
  /** At the top of a page or column (where "before" is dropped). */
  private fresh = true;
  private prev: ParaLayout | null = null;
  private docTop = true;
  /** First main line of the current section on the current page (for column balancing). */
  private sectionStart = 0;

  constructor(private lt: Layouter, private input: FlowInput) {
    this.styles = lt.context.styles;
    this.main = newStory(input.main.length);
    this.labels = listLabels(input.main, input.numbering, this.styles);
    this.numberNotes();
    this.splitSections();
  }

  private numberNotes(): void {
    let fn = 0, en = 0;
    for (const p of this.input.main) {
      for (const o of p.objects ?? []) {
        if (!o.note || this.noteNumbers.has(o.note)) continue;
        if (o.note.startsWith("en:")) this.noteNumbers.set(o.note, formatNumber(++en, "lowerRoman"));
        else this.noteNumbers.set(o.note, String(++fn));
      }
    }
  }

  private splitSections(): void {
    const paras = this.input.main;
    let a = 0;
    let headers: Section["headers"] = {}, footers: Section["footers"] = {};
    const push = (sect: Section, b: number) => {
      // Missing header/footer references are inherited from the previous section (Word).
      headers = { ...headers, ...defined(sect.headers) };
      footers = { ...footers, ...defined(sect.footers) };
      this.sections.push({ sect, a, b, headers, footers });
      a = b + 1;
    };
    for (let i = 0; i < paras.length - 1; i++) if (paras[i].sect) push(paras[i].sect!, i);
    push(this.input.settings.section ?? DEFAULT_SECTION, paras.length - 1);
  }

  run(): void {
    // Two passes when a header or footer shows NUMPAGES: the first finds the page count.
    this.flowMain();
    const total = String(this.pages.length);
    if (total !== this.totalPages && this.usesPageCount()) {
      this.totalPages = total;
      this.reset();
      this.flowMain();
    }
    for (const page of this.pages) this.finishPage(page);
  }

  private usesPageCount(): boolean {
    for (const paras of this.input.stories.values()) {
      for (const p of paras) for (const o of p.objects ?? []) if (fieldType(o.instr) === "NUMPAGES" || fieldType(o.instr) === "SECTIONPAGES") return true;
    }
    return false;
  }

  private reset(): void {
    this.frames.length = 0;
    this.pages.length = 0;
    this.stories.clear();
    this.excl.clear();
    this.main.placed.length = 0;
    this.main.firstPlaced.fill(-1);
    this.cols = [];
    this.col = 0;
    this.y = 0;
    this.fresh = true;
    this.prev = null;
    this.docTop = true;
    this.sectionStart = 0;
  }

  // ----- pages ---------------------------------------------------------------------------

  private frame(f: Omit<Frame, "id">): Frame {
    const fr: Frame = { ...f, id: this.frames.length };
    this.frames.push(fr);
    this.pages[f.page]?.frames.push(fr.id);
    return fr;
  }

  private newPage(si: number, firstOfSection: boolean): void {
    const sr = this.sections[si];
    const sect = sr.sect;
    const setup = sect.page;
    const index = this.pages.length;
    const prevPage = this.pages[index - 1];
    const number = firstOfSection && sect.pageStart !== undefined ? sect.pageStart : (prevPage ? prevPage.number + 1 : 1);
    const page: PageLayout = {
      index, setup, section: si, number, numberText: formatNumber(number, sect.pageFormat ?? "decimal"),
      deco: [], frames: [], bodyTop: setup.marginTop * PT, bodyBottom: (setup.height - setup.marginBottom) * PT, notes: [], notesHeight: 0,
    };
    this.pages.push(page);
    this.page = page;
    // Header and footer: first page of the section, even pages, or the default.
    const pick = (refs: Section["headers"]) => {
      if (firstOfSection && sect.titlePage) return refs.first;
      if (this.input.settings.evenOdd && number % 2 === 0) return refs.even;
      return refs.default;
    };
    page.header = pick(sr.headers);
    page.footer = pick(sr.footers);
    const hdrH = page.header ? this.headerFooter(page, page.header, "header") : 0;
    const ftrH = page.footer ? this.headerFooter(page, page.footer, "footer") : 0;
    page.bodyTop = Math.max(setup.marginTop * PT, setup.header * PT + hdrH);
    page.bodyBottom = Math.min((setup.height - setup.marginBottom) * PT, (setup.height - setup.footer) * PT - ftrH);
    this.sectionStart = this.main.placed.length;
    this.startColumns(si, page.bodyTop);
    this.fresh = true;
    this.prev = null;
  }

  /** Column frames of section `si` from `top` (px) to the bottom of the body. */
  private startColumns(si: number, top: number): void {
    const sect = this.sections[si].sect;
    const page = this.page;
    const s = page.setup;
    const left = (s.marginLeft + s.gutter) * PT;
    const width = (s.width - s.marginLeft - s.marginRight - s.gutter) * PT;
    const c = sect.columns;
    const n = Math.max(1, c.count);
    const widths: { x: number; w: number }[] = [];
    if (n === 1) widths.push({ x: left, w: width });
    else if (c.cols && c.cols.length === n) {
      let x = left;
      for (const col of c.cols) { widths.push({ x, w: col.width * PT }); x += (col.width + col.space) * PT; }
    } else {
      const w = (width - (n - 1) * c.space * PT) / n;
      for (let k = 0; k < n; k++) widths.push({ x: left + k * (w + c.space * PT), w });
    }
    this.cols = widths.map((wc) => this.frame({ page: page.index, story: "main", kind: "body", x: wc.x, y: top, w: wc.w, h: page.bodyBottom - top }));
    if (n > 1 && c.sep) {
      for (let k = 1; k < n; k++) {
        const x = (widths[k - 1].x + widths[k - 1].w + widths[k].x) / 2;
        page.deco.push({ kind: "line", story: "main", x1: x, y1: top, x2: x, y2: page.bodyBottom, width: 0.5 * PT, color: "#000000", style: "single" });
      }
    }
    this.col = 0;
    this.y = 0;
  }

  private get cur(): Frame {
    return this.cols[this.col];
  }

  /** Space left in the current column, above the page's footnotes (px). */
  private room(): number {
    return this.page.bodyBottom - this.page.notesHeight - (this.cur.y + this.y);
  }

  private nextColumn(si: number): void {
    if (this.col < this.cols.length - 1) { this.col++; this.y = 0; this.fresh = true; this.prev = null; return; }
    this.newPage(si, false);
  }

  /** Content bottom of the current page's columns (px). */
  private columnsBottom(): number {
    let bottom = this.cur.y + this.y;
    for (let i = this.sectionStart; i < this.main.placed.length; i++) {
      const pl = this.main.placed[i];
      if (pl.page !== this.page.index) continue;
      bottom = Math.max(bottom, pl.y + this.main.paras[pl.para].lines[pl.line].height);
    }
    return bottom;
  }

  /**
   * Before a continuous section break: spreads the section's lines on this page evenly over
   * its columns (Word balances them, the first column taking the extra line).
   */
  private balance(): void {
    if (this.cols.length < 2) return;
    const lines = this.main.placed.slice(this.sectionStart).map((pl, k) => ({ pl, k: this.sectionStart + k }))
      .filter(({ pl }) => pl.page === this.page.index);
    if (lines.length < 2 || lines.some(({ pl }) => this.frames[pl.frame].kind !== "body")) return;
    const height = (pl: PlacedLine) => this.main.paras[pl.para].lines[pl.line].height;
    // Lines as a stream: gaps between lines of one column kept, column tops dropped.
    const items = lines.map(({ pl }, j) => {
      const prev = lines[j - 1]?.pl;
      const gap = prev && prev.frame === pl.frame ? pl.y - (prev.y + height(prev)) : 0;
      return { gap, h: height(pl) };
    });
    const total = items.reduce((s, it) => s + it.gap + it.h, 0);
    const top = this.cols[0].y;
    const n = this.cols.length;
    for (let target = total / n; target <= total + 1; target += 1) {
      const assign: { col: number; y: number }[] = [];
      let col = 0, y = 0;
      for (const it of items) {
        let gap = y === 0 ? 0 : it.gap;
        if (y > 0 && y + gap + it.h > target + 0.01 && col < n - 1) { col++; y = 0; gap = 0; }
        assign.push({ col, y: y + gap });
        y += gap + it.h;
      }
      if (y > target + 0.01 && col === n - 1 && target < total) continue;
      lines.forEach(({ pl, k }, j) => {
        const fr = this.cols[assign[j].col];
        this.main.placed[k] = { ...pl, x: fr.x, y: top + assign[j].y, frame: fr.id };
      });
      return;
    }
  }

  // ----- headers and footers ----------------------------------------------------------------

  /** Lays out a header or footer instance for `page`; returns its height (px). */
  private headerFooter(page: PageLayout, id: string, kind: "header" | "footer"): number {
    const paras = this.input.stories.get(id);
    if (!paras) return 0;
    const s = page.setup;
    const x = s.marginLeft * PT, w = (s.width - s.marginLeft - s.marginRight) * PT;
    const fields = { PAGE: page.numberText, NUMPAGES: this.totalPages, SECTIONPAGES: this.totalPages };
    const key = `${id}@${page.index}`;
    let top = s.header * PT;
    if (kind === "footer") {
      // The footer's bottom sits at its distance from the page edge: measure, then place.
      const probe = this.simpleFlow(key, paras, { x, y: 0, w, page: -1 }, fields);
      top = (s.height - s.footer) * PT - probe.height;
    }
    const fr = this.frame({ page: page.index, story: key, kind, x, y: top, w, h: 0 });
    const story = this.simpleFlow(key, paras, { x, y: top, w, page: page.index, frame: fr.id }, fields);
    fr.h = story.height;
    this.stories.set(key, story.layout);
    if (kind === "header") page.headerKey = key; else page.footerKey = key;
    return story.height;
  }

  /**
   * Lays out a story without pages (header, footer) into the box at x, y with width w.
   * With page −1 it only measures.
   */
  private simpleFlow(key: string, paras: readonly Para[], box: { x: number; y: number; w: number; page: number; frame?: number },
    fields?: Record<string, string>): { layout: StoryLayout; height: number } {
    const layout = newStory(paras.length);
    const labels = listLabels(paras, this.input.numbering, this.styles);
    const page = box.page >= 0 ? this.pages[box.page] : null;
    let y = 0;
    let prev: ParaLayout | null = null;
    for (let i = 0; i < paras.length; i++) {
      const para = paras[i];
      if (para.cell) {
        let j = i;
        while (j + 1 < paras.length && paras[j + 1].cell?.table.id === para.cell.table.id) j++;
        y += prev ? prev.after : 0;
        const g = this.tableGeom(paras, i, j, box.x, box.w, labels, fields);
        for (const row of g.rows) {
          if (page) this.placeRow(g, row, page, box.y + y, row.height, key, layout, paras, { ghost: false });
          else for (const c of row.cells) c.paras.forEach((pi, k) => { layout.paras[pi] = c.layouts[k]; });
          y += row.height;
        }
        prev = null;
        i = j;
        continue;
      }
      const pl = this.lt.paragraph(para, { width: box.w, label: labels[i], fields, notes: this.noteNumbers });
      layout.paras[i] = pl;
      y += prev ? this.gapBetween(prev, pl) : pl.before;
      const boxTop = y;
      y += pl.borderTop;
      if (page) {
        layout.firstPlaced[i] = layout.placed.length;
        for (let k = 0; k < pl.lines.length; k++) {
          layout.placed.push({ para: i, line: k, page: page.index, x: box.x, y: box.y + y, frame: box.frame ?? -1 });
          y += pl.lines[k].height;
        }
      } else for (const l of pl.lines) y += l.height;
      y += pl.borderBottom;
      if (page) {
        this.paraDeco(page, key, pl, box.x, box.y + boxTop, y - boxTop, paras[i + 1]);
        if (para.floats?.length) this.placeFloats(page, key, i, para.floats, box.y + boxTop, box, false);
      }
      prev = pl;
    }
    y += prev ? prev.after : 0;
    return { layout, height: y };
  }

  // ----- the main story ------------------------------------------------------------------

  private flowMain(): void {
    const paras = this.input.main;
    this.sections.forEach((sr, si) => {
      if (si === 0) this.newPage(0, true);
      else {
        const start = sr.sect.start;
        if (start === "continuous") {
          // Same page: the previous section's columns are balanced, the new ones start below.
          this.balance();
          const top = this.columnsBottom();
          if (top >= this.page.bodyBottom - this.page.notesHeight - 1) this.newPage(si, false);
          else {
            this.sectionStart = this.main.placed.length;
            this.startColumns(si, top);
            this.fresh = false;
            this.prev = null;
          }
        } else if (start === "nextColumn" && this.col < this.cols.length - 1) {
          this.col++; this.y = 0; this.fresh = true; this.prev = null;
        } else {
          this.newPage(si, true);
          const want = start === "evenPage" ? 0 : start === "oddPage" ? 1 : -1;
          if (want >= 0 && this.page.number % 2 !== want) this.newPage(si, false);
        }
      }
      for (let i = sr.a; i <= sr.b; i++) {
        const para = paras[i];
        if (para.cell) {
          let j = i;
          while (j + 1 <= sr.b && paras[j + 1].cell?.table.id === para.cell.table.id) j++;
          this.placeTable(si, i, j);
          i = j;
          continue;
        }
        this.placePara(si, i);
      }
    });
  }

  private env(i: number, width: number): ParaEnv {
    return { width, label: this.labels[i], notes: this.noteNumbers };
  }

  /** Lines of `pl` (from line `from`) that fit in `room` px, honouring widow/orphan control. */
  private fitting(pl: ParaLayout, from: number, room: number): number {
    const n = pl.lines.length - from;
    let k = 0, acc = 0;
    while (k < n && acc + pl.lines[from + k].height <= room + 0.01) { acc += pl.lines[from + k].height; k++; }
    if (pl.rp.keepLines && from === 0 && k < n) return 0;
    if (pl.rp.widow) {
      if (n > 1 && k === 1 && from === 0) k = 0;
      if (n >= 3 && n - k === 1 && k > 1) k -= 1;
      if (n === 2 && k === 1 && from === 0) k = 0;
    }
    return k;
  }

  private gapBetween(a: ParaLayout | null, b: ParaLayout): number {
    if (!a) return b.before;
    const same = a.rp.style === b.rp.style;
    const after = a.rp.contextual && same ? 0 : a.after;
    const before = b.rp.contextual && same ? 0 : b.before;
    return Math.max(after, before);
  }

  /** Height a keep-with-next chain starting at paragraph i needs (px). */
  private chainHeight(i: number, width: number): number {
    const paras = this.input.main;
    let h = 0;
    let prev: ParaLayout | null = null;
    for (let j = i; j < paras.length && !paras[j].cell; j++) {
      const pl = this.lt.paragraph(paras[j], this.env(j, width));
      if (prev) h += this.gapBetween(prev, pl);
      const lines = pl.lines.map((l) => l.height);
      if (pl.rp.keepNext && j + 1 < paras.length) { h += lines.reduce((s, x) => s + x, 0); prev = pl; continue; }
      const need = pl.rp.keepLines ? lines.length : Math.min(lines.length, pl.rp.widow && lines.length > 1 ? 2 : 1);
      for (let k = 0; k < need; k++) h += lines[k];
      break;
    }
    return h;
  }

  /** Footnotes referenced in a line that are not yet on the current page, and their height. */
  private notesIn(para: Para, line: Line): { ids: string[]; height: number } {
    const ids: string[] = [];
    if (!para.objects?.some((o) => o.note)) return { ids, height: 0 };
    let objIdx = 0;
    for (let i = 0; i < line.end && i < para.text.length; i++) {
      if (para.text[i] !== OBJ) continue;
      const o = para.objects[objIdx++];
      if (i >= line.start && o?.note && o.note.startsWith("fn:") && !this.page.notes.includes(o.note) && this.input.stories.has(o.note)) ids.push(o.note);
    }
    if (!ids.length) return { ids, height: 0 };
    let h = this.page.notes.length ? 0 : this.separatorHeight();
    for (const id of ids) h += this.noteLayout(id).height;
    return { ids, height: h };
  }

  private separatorHeight(): number {
    // The separator is one line in the Normal style's font, without spacing.
    const c = this.styles.resolveChar(this.styles.defaultParagraphStyle, EMPTY);
    return this.lt.context.fonts.lineHeight(c.font, c.size) * PT;
  }

  private noteLayout(id: string): { paras: ParaLayout[]; height: number; width: number } {
    const s = this.page.setup;
    const width = (s.width - s.marginLeft - s.marginRight) * PT;
    const hit = this.noteLayouts.get(id);
    if (hit && hit.width === width) return hit;
    const paras = this.input.stories.get(id) ?? [];
    let h = 0;
    let prev: ParaLayout | null = null;
    const out: ParaLayout[] = [];
    for (const p of paras) {
      const pl = this.lt.paragraph(p, { width, notes: this.noteNumbers });
      h += prev ? this.gapBetween(prev, pl) : pl.before;
      for (const l of pl.lines) h += l.height;
      prev = pl;
      out.push(pl);
    }
    if (prev) h += prev.after;
    const res = { paras: out, height: h, width };
    this.noteLayouts.set(id, res);
    return res;
  }

  private placePara(si: number, i: number): void {
    const paras = this.input.main;
    const para = paras[i];
    let width = this.cur.w;
    let pl = this.lt.paragraph(para, this.env(i, width));
    if (pl.rp.pageBreakBefore && !(this.fresh && this.y === 0 && this.col === 0)) {
      this.newPage(si, false);
      width = this.cur.w;
      pl = this.lt.paragraph(para, this.env(i, width));
    }
    // "Before" counts at the top of the document but is dropped at the top of every later page
    // and column, whatever started it (measured: natural, manual and "page break before").
    const gapFor = (p: ParaLayout) => (this.y === 0 && this.fresh ? (this.docTop ? p.before : 0) : this.gapBetween(this.prev, p)) + p.borderTop;
    let gap = gapFor(pl);
    const n0 = pl.lines.length;
    let k = this.fitting(pl, 0, this.room() - gap);
    if (k === n0 && pl.rp.keepNext && i + 1 < paras.length && this.y > 0 && !paras[i + 1].cell) {
      // Keep with next: this paragraph and the start of the following chain share a page.
      const nextPl = this.lt.paragraph(paras[i + 1], this.env(i + 1, width));
      const need = pl.lines.reduce((s, l) => s + l.height, 0) + this.gapBetween(pl, nextPl) + this.chainHeight(i + 1, width);
      if (gap + need > this.room() + 0.01 && this.chainHeight(i, width) <= this.cur.h) k = 0;
    }
    if (k === 0 && this.y > 0) {
      this.nextColumn(si);
      if (this.cur.w !== width) { width = this.cur.w; pl = this.lt.paragraph(para, this.env(i, width)); }
      gap = gapFor(pl);
    }
    // Pictures anchored here, and pictures already on the page, wrap the text.
    const textTop = this.cur.y + this.y + gap;
    if (para.floats?.length) this.placeFloats(this.page, "main", i, para.floats, textTop - pl.borderTop, this.cur, true);
    if (this.excl.get(this.page.index)?.length) pl = this.withExclusions(para, i, width, textTop) ?? pl;
    const n = pl.lines.length;
    k = Math.max(1, this.fitting(pl, 0, this.room() - gap)); // taller than a page: split anyway
    this.y += gap;
    this.fresh = false;
    this.docTop = false;
    this.main.paras[i] = pl;
    this.main.firstPlaced[i] = this.main.placed.length;
    let pieceTop = this.cur.y + this.y - pl.borderTop;
    let pieceFrame = this.cur;
    let pieceFirst = true;
    const closePiece = (last: boolean) => {
      const bottom = this.cur.y + this.y;
      this.paraDeco(this.page, "main", pl, pieceFrame.x, pieceTop, bottom - pieceTop, last ? paras[i + 1] : null, !pieceFirst, !last);
    };
    let j = 0;
    let take = k;
    while (j < n) {
      let broke = false;
      for (let t = 0; t < take && j < n; t++) {
        const line = pl.lines[j];
        const notes = this.notesIn(para, line);
        if (notes.ids.length && this.y > 0 && line.height + notes.height > this.room() + 0.01) break; // its footnote does not fit
        if (notes.ids.length) { this.page.notes.push(...notes.ids); this.page.notesHeight += notes.height; }
        this.main.placed.push({ para: i, line: j, page: this.page.index, x: this.cur.x, y: this.cur.y + this.y, frame: this.cur.id });
        this.y += line.height;
        j++;
        if (line.pageBreakAfter || line.columnBreakAfter) {
          if (j >= n) {
            // The break ends the paragraph: the next paragraph starts on the new page.
            this.y += pl.borderBottom;
            closePiece(true);
            if (line.pageBreakAfter) this.newPage(si, false); else this.nextColumn(si);
            return;
          }
          broke = true;
          closePiece(false);
          if (line.pageBreakAfter) this.newPage(si, false); else this.nextColumn(si);
          break;
        }
      }
      if (j >= n) break;
      if (!broke) { closePiece(false); this.nextColumn(si); }
      this.fresh = false;
      pieceTop = this.cur.y + this.y;
      pieceFrame = this.cur;
      pieceFirst = false;
      take = Math.max(1, this.fitting(pl, j, this.room()));
    }
    this.y += pl.borderBottom;
    closePiece(true);
    this.prev = pl;
  }

  // ----- paragraph borders and shading --------------------------------------------------------

  /**
   * Shading and borders of a paragraph (or the piece of it on one page) at x, y with height h.
   * Paragraphs with equal borders form one box; `contTop`/`contBottom` mark pieces that
   * continue from or on another page (no top/bottom border there).
   */
  private paraDeco(page: PageLayout, story: string, pl: ParaLayout, x: number, y: number, h: number, next: Para | undefined | null,
    contTop = false, contBottom = false): void {
    const rp = pl.rp;
    const b = rp.borders;
    const fill = fillOf(rp.shading);
    if (!b && !fill) return;
    const lb = b && visibleBorder(b.left) ? b.left : undefined;
    const rb = b && visibleBorder(b.right) ? b.right : undefined;
    const left = x + rp.indLeft * PT - (lb ? (lb.space + lb.size) * PT : 0);
    const right = x + pl.width - rp.indRight * PT + (rb ? (rb.space + rb.size) * PT : 0);
    if (fill) page.deco.push({ kind: "rect", story, x: left, y, w: right - left, h, fill });
    if (!b) return;
    const sameNext = !!next && !next.cell && sameBorders(this.styles.resolvePara(next.props).borders, b);
    const line = (spec: BorderSpec | undefined, x1: number, y1: number, x2: number, y2: number) => {
      if (visibleBorder(spec)) page.deco.push({ kind: "line", story, x1, y1, x2, y2, width: spec.size * PT, color: colorOf(spec.color), style: spec.style });
    };
    if (!contTop && visibleBorder(b.top) && pl.borderTop > 0) line(b.top, left, y + (b.top.size * PT) / 2, right, y + (b.top.size * PT) / 2);
    if (!contBottom) {
      if (sameNext) line(b.between, left, y + h, right, y + h);
      else if (visibleBorder(b.bottom) && pl.borderBottom > 0) line(b.bottom, left, y + h - (b.bottom.size * PT) / 2, right, y + h - (b.bottom.size * PT) / 2);
    }
    line(b.left, left, y, left, y + h);
    line(b.right, right, y, right, y + h);
  }

  // ----- floating pictures --------------------------------------------------------------------

  private placeFloats(page: PageLayout, story: string, i: number, floats: readonly FloatObject[], paraTop: number,
    column: { x: number; w: number }, wrap: boolean): void {
    const s = page.setup;
    floats.forEach((f, index) => {
      const w = f.width * PT, h = f.height * PT;
      const relX = (rel: string): [number, number] => {
        if (rel === "page") return [0, s.width * PT];
        if (rel === "column" || rel === "character") return [column.x, column.w];
        if (rel === "leftMargin" || rel === "insideMargin") return [0, s.marginLeft * PT];
        if (rel === "rightMargin" || rel === "outsideMargin") return [(s.width - s.marginRight) * PT, s.marginRight * PT];
        return [s.marginLeft * PT, (s.width - s.marginLeft - s.marginRight) * PT];
      };
      const relY = (rel: string): [number, number] => {
        if (rel === "page") return [0, s.height * PT];
        if (rel === "margin") return [s.marginTop * PT, (s.height - s.marginTop - s.marginBottom) * PT];
        if (rel === "topMargin") return [0, s.marginTop * PT];
        if (rel === "bottomMargin") return [(s.height - s.marginBottom) * PT, s.marginBottom * PT];
        return [paraTop, 0];
      };
      const [x0, xw] = relX(f.x.rel);
      const [y0, yh] = relY(f.y.rel);
      const x = f.x.align
        ? (f.x.align === "center" ? x0 + (xw - w) / 2 : f.x.align === "right" || f.x.align === "outside" ? x0 + xw - w : x0)
        : x0 + f.x.offset * PT;
      const y = f.y.align && yh > 0
        ? (f.y.align === "center" ? y0 + (yh - h) / 2 : f.y.align === "bottom" || f.y.align === "outside" ? y0 + yh - h : y0)
        : y0 + f.y.offset * PT;
      page.deco.push({ kind: "image", story, x, y, w, h, image: f.image, behind: f.wrap === "behind", para: i, index });
      if (wrap && f.wrap !== "front" && f.wrap !== "behind" && f.wrap !== "none") {
        const [dt, db, dl, dr] = f.dist.map((d) => d * PT);
        const list = this.excl.get(page.index) ?? [];
        list.push({ x: x - dl, y: y - dt, w: w + dl + dr, h: h + dt + db, wrap: f.wrap, side: f.side });
        this.excl.set(page.index, list);
      }
    });
  }

  /** Lays out a paragraph against the current page's pictures, or null when none is near. */
  private withExclusions(para: Para, i: number, width: number, paraTop: number): ParaLayout | null {
    const list = this.excl.get(this.page.index);
    if (!list?.length) return null;
    const fr = this.cur;
    const minRoom = 36 * PT;
    let any = false;
    const box = (top: number, h: number) => {
      const y1 = paraTop + top, y2 = y1 + h;
      let left = 0, right = width, push = 0;
      let hit = false;
      for (const e of list) {
        if (e.y + e.h <= y1 || e.y >= y2) continue;
        const ex1 = e.x - fr.x, ex2 = e.x + e.w - fr.x;
        if (ex2 <= 0 || ex1 >= width) continue;
        hit = true;
        if (e.wrap === "topAndBottom") { push = Math.max(push, e.y + e.h - y1); continue; }
        const leftRoom = ex1 - left, rightRoom = right - ex2;
        // Text flows on the side with room (Word's "both sides" uses the larger one when
        // only one side has at least about an inch).
        const side = e.side === "left" ? "left" : e.side === "right" ? "right" : leftRoom >= rightRoom ? "left" : "right";
        if (side === "left") right = Math.min(right, ex1);
        else left = Math.max(left, ex2);
        if (right - left < minRoom) push = Math.max(push, e.y + e.h - y1);
      }
      if (hit) any = true;
      return hit ? (push > 0 ? { left: 0, right: width, push } : { left, right }) : null;
    };
    const pl = this.lt.paragraph(para, { ...this.env(i, width), box });
    return any ? pl : null;
  }

  // ----- tables ---------------------------------------------------------------------------

  /** Geometry of the table formed by paragraphs a..b (cells laid out, rows measured). */
  private tableGeom(paras: readonly Para[], a: number, b: number, frameX: number, frameW: number,
    labels: (ListLabel | undefined)[], fields?: Record<string, string>): TableGeom {
    const def = paras[a].cell!.table;
    const style = this.styles.tableStyle(def.style);
    const gridW = def.grid.map((g) => g * PT);
    const total = gridW.reduce((s, x) => s + x, 0);
    let x = frameX + (def.indent ?? 0) * PT;
    if (def.align === "center") x = frameX + (frameW - total) / 2;
    else if (def.align === "right") x = frameX + frameW - total;
    const margins: CellMargins = { ...DEFAULT_CELL_MARGINS, ...(style?.cellMargins ?? {}), ...(def.cellMargins ?? {}) };
    const byCell = new Map<string, number[]>();
    for (let i = a; i <= b; i++) {
      const k = cellKey(paras[i].cell!);
      const list = byCell.get(k);
      if (list) list.push(i); else byCell.set(k, [i]);
    }
    const look = def.look ?? { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true };
    const nRows = def.rows.length;
    const rows: RowGeom[] = def.rows.map((row, r) => {
      let gx = 0;
      const cells = row.cells.map((cd, c): CellGeom => {
        const span = Math.max(1, cd.span);
        const cx = x + sum(gridW, 0, gx);
        const cw = sum(gridW, gx, gx + span) || (cd.width ?? 50) * PT;
        const m: CellMargins = { ...margins, ...(cd.margins ?? {}) };
        const conds = this.conditions(style, look, r, nRows, gx, span, gridW.length);
        const pPr = mergeParts(style?.pPr, conds.map((p) => p.pPr));
        const rPr = mergeChars(style?.rPr, conds.map((p) => p.rPr));
        const shading = fillOf(cd.shading) ?? [...conds].reverse().map((p) => fillOf(p.shading)).find((v) => v) ?? fillOf(def.shading) ?? fillOf(style?.shading) ?? null;
        const borders: Borders = Object.assign({}, style?.borders ?? {}, ...conds.map((p) => p.borders ?? {}), def.borders ?? {}, cd.borders ?? {});
        const contentW = Math.max(4, cw - (m.left + m.right) * PT);
        const ids = byCell.get(`${def.id}:${r}:${c}`) ?? [];
        const layouts: ParaLayout[] = [];
        let h = 0;
        let prev: ParaLayout | null = null;
        for (const i of ids) {
          const pl = this.lt.paragraph(paras[i], { width: contentW, label: labels[i], tablePPr: pPr, tableRPr: rPr, fields, notes: this.noteNumbers });
          layouts.push(pl);
          h += prev ? this.gapBetween(prev, pl) : pl.before;
          h += pl.borderTop + pl.lines.reduce((s, l) => s + l.height, 0) + pl.borderBottom;
          prev = pl;
        }
        if (prev) h += prev.after;
        gx += span;
        return {
          x: cx, w: cw, contentX: cx + m.left * PT, contentW, margins: m, paras: ids, layouts, need: h + (m.top + m.bottom) * PT,
          vMerge: cd.vMerge, shading, borders, vAlign: cd.vAlign ?? "top", span: 1,
        };
      });
      return { cells, height: 0, header: !!row.header, cantSplit: !!row.cantSplit };
    });
    // Vertically merged cells: a restart cell spans the following "continue" cells below it.
    rows.forEach((row, r) => {
      for (const cell of row.cells) {
        if (cell.vMerge !== "restart") continue;
        let span = 1;
        while (r + span < rows.length && rows[r + span].cells.some((c) => Math.abs(c.x - cell.x) < 0.5 && c.vMerge === "continue")) span++;
        cell.span = span;
      }
    });
    rows.forEach((row, r) => {
      let h = 0;
      for (const c of row.cells) if (c.vMerge !== "continue" && c.span === 1) h = Math.max(h, c.need);
      const rd = def.rows[r];
      if (rd.height !== undefined) h = rd.heightRule === "exact" ? rd.height * PT : Math.max(h, rd.height * PT);
      row.height = h;
    });
    // A merged cell taller than the rows it spans makes the last of them taller.
    rows.forEach((row, r) => {
      for (const cell of row.cells) {
        if (cell.span <= 1) continue;
        const have = rows.slice(r, r + cell.span).reduce((s, x) => s + x.height, 0);
        if (cell.need > have) rows[r + cell.span - 1].height += cell.need - have;
      }
    });
    return { def, x, w: total, rows };
  }

  /** Conditional formats of a table style that apply to a cell, in Word's order. */
  private conditions(style: TableStyleDef | undefined, look: NonNullable<TableDef["look"]>, r: number, nRows: number,
    gx: number, span: number, nCols: number): TableStylePart[] {
    const out: TableStylePart[] = [];
    const cond = style?.cond;
    if (!style || !cond) return out;
    const band = (n: number, size: number) => (Math.floor(n / Math.max(1, size)) % 2 === 0 ? 1 : 2);
    const bodyCol = gx - (look.firstColumn ? 1 : 0);
    if (!look.noVBand && bodyCol >= 0) { const c = cond[`band${band(bodyCol, style.colBand ?? 1)}Vert`]; if (c) out.push(c); }
    const bodyRow = r - (look.firstRow ? 1 : 0);
    if (!look.noHBand && bodyRow >= 0) { const c = cond[`band${band(bodyRow, style.rowBand ?? 1)}Horz`]; if (c) out.push(c); }
    if (look.firstColumn && gx === 0 && cond.firstCol) out.push(cond.firstCol);
    if (look.lastColumn && gx + span === nCols && cond.lastCol) out.push(cond.lastCol);
    if (look.firstRow && r === 0 && cond.firstRow) out.push(cond.firstRow);
    if (look.lastRow && r === nRows - 1 && cond.lastRow) out.push(cond.lastRow);
    return out;
  }

  /**
   * Places one row at `top` (page px) with height `h`. `from`/`to` limit the lines placed
   * per paragraph (a row split across pages).
   */
  private placeRow(g: TableGeom, row: RowGeom, page: PageLayout, top: number, h: number, story: string, layout: StoryLayout,
    paras: readonly Para[], opt: { ghost: boolean; from?: Map<number, number>; to?: Map<number, number> }): void {
    const r = g.rows.indexOf(row);
    const split = !!opt.from || !!opt.to;
    for (const cell of row.cells) {
      const merged = cell.vMerge === "continue";
      const cellH = cell.span > 1 && !split ? h + sum(g.rows.map((x) => x.height), r + 1, r + cell.span) : h;
      const fr = this.frame({ page: page.index, story, kind: "cell", x: cell.x, y: top, w: cell.w, h: cellH, cell: cell.paras.length ? cellKey(paras[cell.paras[0]].cell!) : undefined });
      if (!merged && cell.shading) page.deco.push({ kind: "rect", story, x: cell.x, y: top, w: cell.w, h: cellH, fill: cell.shading });
      // Content, vertically aligned in the cell.
      let y = top + cell.margins.top * PT;
      if (cell.vAlign !== "top" && !split) {
        const free = cellH - cell.need;
        if (free > 0) y += cell.vAlign === "center" ? free / 2 : free;
      }
      let prev: ParaLayout | null = null;
      cell.paras.forEach((pi, k) => {
        const pl = cell.layouts[k];
        layout.paras[pi] = pl;
        if (merged) {
          // Text of a merged-away cell: kept in the model, not shown (as in Word).
          if (!opt.ghost) for (let ln = 0; ln < pl.lines.length; ln++) layout.placed.push({ para: pi, line: ln, page: page.index, x: cell.contentX, y: top, frame: fr.id, hide: true });
          return;
        }
        const start = opt.from?.get(pi) ?? 0;
        const end = opt.to?.get(pi) ?? pl.lines.length;
        if (start === 0) y += prev ? this.gapBetween(prev, pl) : pl.before;
        const boxTop = y;
        if (start === 0) y += pl.borderTop;
        if (!opt.ghost && layout.firstPlaced[pi] === -1 && start < end) layout.firstPlaced[pi] = layout.placed.length;
        for (let ln = start; ln < end; ln++) {
          layout.placed.push({ para: pi, line: ln, page: page.index, x: cell.contentX, y, frame: fr.id, ghost: opt.ghost || undefined });
          y += pl.lines[ln].height;
        }
        if (end === pl.lines.length) y += pl.borderBottom;
        if (end > start) this.paraDeco(page, story, pl, cell.contentX, boxTop, y - boxTop, end === pl.lines.length ? paras[cell.paras[k + 1]] ?? null : null, start > 0, end < pl.lines.length);
        prev = pl;
      });
      if (!merged) this.cellBorders(page, story, g, r, cell, top, cellH, !!opt.from);
    }
  }

  private cellBorders(pg: PageLayout, story: string, g: TableGeom, r: number, cell: CellGeom, top: number, h: number, continued: boolean): void {
    const b = cell.borders;
    const firstRow = r === 0 || continued, lastRow = r + cell.span >= g.rows.length;
    const firstCol = Math.abs(cell.x - g.x) < 0.5, lastCol = Math.abs(cell.x + cell.w - (g.x + g.w)) < 0.5;
    const pick = (outer: BorderSpec | undefined, inner: BorderSpec | undefined, isOuter: boolean) => (isOuter ? outer : inner ?? outer);
    const line = (spec: BorderSpec | undefined, x1: number, y1: number, x2: number, y2: number) => {
      if (visibleBorder(spec)) pg.deco.push({ kind: "line", story, x1, y1, x2, y2, width: spec.size * PT, color: colorOf(spec.color), style: spec.style });
    };
    // Each edge is drawn once: tops by every row, bottoms by the last row (and at page
    // breaks), lefts by every cell, rights by the last cell.
    line(pick(b.top, b.insideH, firstRow), cell.x, top, cell.x + cell.w, top);
    if (lastRow) line(b.bottom, cell.x, top + h, cell.x + cell.w, top + h);
    line(pick(b.left, b.insideV, firstCol), cell.x, top, cell.x, top + h);
    if (lastCol) line(b.right, cell.x + cell.w, top, cell.x + cell.w, top + h);
  }

  private placeTable(si: number, a: number, b: number): void {
    const paras = this.input.main;
    if (!(this.y === 0 && this.fresh)) this.y += this.prev ? this.prev.after : 0;
    this.fresh = false;
    this.docTop = false;
    const g = this.tableGeom(paras, a, b, this.cur.x, this.cur.w, this.labels);
    let nHeader = 0;
    while (nHeader < g.rows.length && g.rows[nHeader].header) nHeader++;
    const repeatHeaders = () => {
      for (let r = 0; r < nHeader; r++) {
        const hr = g.rows[r];
        if (hr.height > this.room()) break;
        this.placeRow(g, hr, this.page, this.cur.y + this.y, hr.height, "main", this.main, paras, { ghost: true });
        this.y += hr.height;
      }
    };
    for (let r = 0; r < g.rows.length; r++) {
      const row = g.rows[r];
      if (row.height > this.room() + 0.01 && this.y > 0) {
        // Rows break across pages at line boundaries unless they may not; a row of which no
        // line fits moves completely.
        const bottomBorder = () => this.closeTableBottom(g, r);
        if (!row.cantSplit && this.splitRow(si, g, row, r >= nHeader ? repeatHeaders : null)) continue;
        bottomBorder();
        this.nextColumn(si);
        if (r >= nHeader) repeatHeaders();
      }
      this.placeRow(g, row, this.page, this.cur.y + this.y, row.height, "main", this.main, paras, { ghost: false });
      this.y += row.height;
    }
    this.prev = null;
  }

  /** A table broken before row r gets its bottom border at the end of the page. */
  private closeTableBottom(g: TableGeom, r: number): void {
    if (r === 0) return;
    const y = this.cur.y + this.y;
    for (const cell of g.rows[r - 1].cells) {
      const spec = cell.borders.bottom;
      if (visibleBorder(spec)) this.page.deco.push({ kind: "line", story: "main", x1: cell.x, y1: y, x2: cell.x + cell.w, y2: y, width: spec.size * PT, color: colorOf(spec.color), style: spec.style });
    }
  }

  /** Breaks a row at line boundaries: what fits stays here. Returns false if no line fits. */
  private splitRow(si: number, g: TableGeom, row: RowGeom, repeatHeaders: (() => void) | null): boolean {
    const room = this.room() - 0.01;
    const cut = new Map<number, number>();
    let any = false;
    for (const cell of row.cells) {
      if (cell.vMerge === "continue") continue;
      let y = cell.margins.top * PT;
      let prev: ParaLayout | null = null;
      let full = false;
      cell.paras.forEach((pi, k) => {
        const pl = cell.layouts[k];
        let n = 0;
        if (!full) {
          y += (prev ? this.gapBetween(prev, pl) : pl.before) + pl.borderTop;
          for (const l of pl.lines) { if (y + l.height > room) break; y += l.height; n++; }
          if (n < pl.lines.length) full = true;
          y += pl.borderBottom;
        }
        cut.set(pi, n);
        if (n > 0) any = true;
        prev = pl;
      });
    }
    if (!any) return false;
    const top = this.cur.y + this.y;
    this.placeRow(g, row, this.page, top, this.room(), "main", this.main, this.input.main, { ghost: false, to: cut });
    this.y = this.page.bodyBottom - this.page.notesHeight - this.cur.y;
    this.closeTableBottom(g, g.rows.indexOf(row) + 1);
    this.nextColumn(si);
    repeatHeaders?.();
    // The rest of the row on the next page.
    let restH = 0;
    for (const cell of row.cells) {
      let h = cell.margins.top * PT;
      cell.paras.forEach((pi, k) => {
        const pl = cell.layouts[k];
        const from = cut.get(pi) ?? 0;
        if (from === 0 && k > 0) h += pl.before;
        for (let ln = from; ln < pl.lines.length; ln++) h += pl.lines[ln].height;
      });
      h += cell.margins.bottom * PT + (cell.layouts[cell.layouts.length - 1]?.after ?? 0);
      restH = Math.max(restH, h);
    }
    this.placeRow(g, row, this.page, this.cur.y + this.y, restH, "main", this.main, this.input.main, { ghost: false, from: cut });
    this.y += restH;
    return true;
  }

  // ----- footnotes ------------------------------------------------------------------------

  private finishPage(page: PageLayout): void {
    if (!page.notes.length) return;
    const s = page.setup;
    const x = s.marginLeft * PT, w = (s.width - s.marginLeft - s.marginRight) * PT;
    let y = page.bodyBottom - page.notesHeight;
    const sep = this.separatorHeight();
    // Word's separator: a 2-inch line in the middle of a Normal line.
    page.deco.push({ kind: "line", story: "notes", x1: x, y1: y + sep / 2, x2: x + 144 * PT, y2: y + sep / 2, width: 0.5 * PT, color: "#000000", style: "single" });
    y += sep;
    for (const id of page.notes) {
      const nl = this.noteLayouts.get(id);
      if (!nl) continue;
      const fr = this.frame({ page: page.index, story: id, kind: "note", x, y, w, h: nl.height });
      const layout = newStory(nl.paras.length);
      let yy = 0;
      let prev: ParaLayout | null = null;
      nl.paras.forEach((pl, i) => {
        layout.paras[i] = pl;
        yy += prev ? this.gapBetween(prev, pl) : pl.before;
        layout.firstPlaced[i] = layout.placed.length;
        pl.lines.forEach((ln, k) => {
          layout.placed.push({ para: i, line: k, page: page.index, x, y: y + yy, frame: fr.id });
          yy += ln.height;
        });
        prev = pl;
      });
      this.stories.set(id, layout);
      y += nl.height;
    }
  }
}

function sum(a: readonly number[], from: number, to: number): number {
  let s = 0;
  for (let i = from; i < to && i < a.length; i++) s += a[i];
  return s;
}

function defined<T extends object>(o: T): Partial<T> {
  return Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined)) as Partial<T>;
}

function sameBorders(a: Borders | undefined, b: Borders | undefined): boolean {
  if (!a || !b) return false;
  const k = (x: BorderSpec | undefined) => (x ? `${x.style}:${x.size}:${x.color}:${x.space}` : "");
  return k(a.top) === k(b.top) && k(a.bottom) === k(b.bottom) && k(a.left) === k(b.left) && k(a.right) === k(b.right);
}

// Merged table-style formatting must keep its identity between layouts (paragraph caches
// are keyed by object identity), so merges are cached per combination.
const EMPTY_HOLDER = {};
let keySeq = 0;
const keys = new WeakMap<object, number>();
function objKey(o: object): number {
  let k = keys.get(o);
  if (k === undefined) { k = ++keySeq; keys.set(o, k); }
  return k;
}

function cachedMerge<T extends object>(cache: WeakMap<object, Map<string, T>>, base: T | undefined, more: (T | undefined)[]): T | undefined {
  const list = more.filter((x): x is T => !!x);
  if (!list.length) return base;
  const key = list.map(objKey).join(",");
  const holder = base ?? EMPTY_HOLDER;
  let m = cache.get(holder);
  if (!m) { m = new Map(); cache.set(holder, m); }
  let hit = m.get(key);
  if (!hit) { hit = Object.assign({}, base ?? {}, ...list) as T; m.set(key, hit); }
  return hit;
}

const partCache = new WeakMap<object, Map<string, ParaProps>>();
const charCache = new WeakMap<object, Map<string, CharProps>>();
const mergeParts = (base: ParaProps | undefined, more: (ParaProps | undefined)[]) => cachedMerge(partCache, base, more);
const mergeChars = (base: CharProps | undefined, more: (CharProps | undefined)[]) => cachedMerge(charCache, base, more);

function buildIndex(story: StoryLayout): void {
  story.index.length = 0;
  story.placed.forEach((pl, i) => {
    if (pl.ghost) return;
    const row = (story.index[pl.para] ??= []);
    if (row[pl.line] === undefined) row[pl.line] = i;
  });
  for (let p = 0; p < story.firstPlaced.length; p++) {
    if (story.index[p]) story.firstPlaced[p] = story.index[p][0];
  }
}

/** Lays out a whole document. */
export function layoutDocument(lt: Layouter, input: FlowInput): Layout {
  const flow = new Flow(lt, input);
  flow.run();
  buildIndex(flow.main);
  for (const s of flow.stories.values()) buildIndex(s);
  return {
    main: flow.main, stories: flow.stories, pages: flow.pages, frames: flow.frames, pageCount: flow.pages.length,
    paras: flow.main.paras, placed: flow.main.placed, firstPlaced: flow.main.firstPlaced, noteNumbers: flow.noteNumbers,
  };
}
