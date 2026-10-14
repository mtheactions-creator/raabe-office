// Public API of the text engine (decision D-017: controlled editing).

export { Editor, FONT_STEPS, type ChangeEvent, type EditorStats, type CharState, type CaseMode, type CopiedFormat, type DrawTool } from "./editor";
export { DEFAULT_MATH, Doc, MAIN, cmp, nextTableId, type Pos, type DocState, type DocSettings, type DocSnapshot, type MathSettings, type NoteProps, type Watermark } from "./doc";
export { Fonts, PT, type FontMetrics, type MetricsProvider } from "./fonts";
export { Layouter, fieldType, type LayoutContext } from "./layout";
export { DEFAULT_CELL_MARGINS, layoutDocument, type Deco, type Frame, type Layout, type PageLayout, type PlacedLine, type StoryLayout } from "./flow";
export {
  A4, COLUMN_BREAK, DEFAULT_SECTION, DEFAULT_SHAPE_BODY, EMPTY, EMPTY_NUMBERING, LINE_BREAK, OBJ, ONE_COLUMN, PAGE_BREAK, cellKey, intern, makePara, validPara,
  type AltText, type ShapeBody, type ShapeInfo,
  type AbstractNum, type Align, type Anchor, type BlockObject, type BorderSpec, type Borders, type CellDef, type CellMargins,
  type CellRef, type CharProps, type Columns, type CommentMeta, type GroupInfo, type GroupMember, type LineNumbers, type PageBorders, type FloatObject, type HeaderRefs, type ImageData,
  type InlineObject, type LineRule, type NumInstance, type NumLevel, type Numbering, type PageSetup, type Para,
  type ParaGroup, type ParaProps, type Revision, type RowDef, type Run, type Section, type Shading, type TabStop,
  type TableDef, type Underline,
} from "./model";
export {
  addList, bulletLevels, displayBullet, formatNumber, levelOf, listLabels, multilevelLevels, numberLevels, numberingOf,
  restartList, type ListLabel,
} from "./numbering";
export {
  themeColor,
  GERMAN_NAMES, OFFICE_2013_COLORS, StyleSheet, builtinStyle, defaultStyleSheet,
  type ResolvedChar, type ResolvedPara, type StyleDef, type StyleType, type TableStyleDef, type TableStylePart, type Theme,
} from "./styles";
export { AUTHOR_COLORS, HIGHLIGHT, MIN_ZOOM, MAX_ZOOM, imageUrl, storyIdOf, type TextMark } from "./view";
export { countWords } from "./words";
export { imageInfo, naturalSizePt, type ImageInfo } from "./images";
export { effectExtent, pictureElement, plainPicture, type Fill, type LineEnd, type PictureFormat, type PictureShadow } from "./picture";
export { customGeometry, flattenPath, hasPreset, presetAdjust, presetGeometry, presetNames, type GeomHandle, type GeomPath, type Geometry } from "./geometry";
export { SHAPE_NAMES, shadeColor, shadeLinear, shapeElement, shapeGeometry, shapeLook, shapeTextRect } from "./shape";
export { fromHtml, toHtml } from "./html";
export * as Revisions from "./revisions";
export * as Edits from "./edits";
export * as School from "./school";
export { installProbe, type EngineProbe, type LineBox, type ProbeFormat } from "./probe";

import type { Editor } from "./editor";
import type { EngineProbe, ProbeFormat } from "./probe";
import type { DocState } from "./doc";
import { propsAt, type CharProps } from "./model";
import { PT } from "./fonts";

/** Adapts an editor to the feel-lab measurement interface (see docs/FEEL.md). */
export function probeFor(getEditor: () => Editor, load: (name: string) => Partial<DocState>): EngineProbe {
  const lineStart = (e: Editor, para: number, line: number) =>
    e.doc.toOffset({ p: para, o: e.layout.paras[para].lines[line].start });
  const format = (e: Editor, para: number, props: CharProps): ProbeFormat => {
    const s = e.doc.styles;
    const rp = s.resolvePara(e.doc.paras[para].props);
    const c = s.resolveChar(rp.style, props);
    return {
      font: c.font, size: c.size, bold: c.bold, italic: c.italic, underline: c.underline, strike: c.strike,
      vertAlign: c.vertAlign, color: c.color, highlight: c.highlight, caps: c.caps, smallCaps: c.smallCaps, hidden: c.hidden,
      dstrike: c.dstrike, style: s.displayName(rp.style), align: rp.align, indLeft: rp.indLeft, indFirst: rp.indFirst,
      before: rp.before, after: rp.after, line: rp.line, lineRule: rp.lineRule,
    };
  };
  return {
    engine: "b",
    text: () => getEditor().doc.text(),
    selection: () => {
      const e = getEditor();
      return { anchor: e.doc.toOffset(e.sel.anchor), head: e.doc.toOffset(e.sel.head) };
    },
    setSelection(anchor, head) {
      const e = getEditor();
      e.select(e.doc.fromOffset(anchor), e.doc.fromOffset(head));
    },
    pointAt: (offset) => getEditor().view.clientPoint(getEditor().doc.fromOffset(offset)),
    lines() {
      const e = getEditor();
      return e.layout.placed.filter((pl) => !pl.ghost && !pl.hide).map((pl) => lineStart(e, pl.para, pl.line));
    },
    pages() {
      const e = getEditor();
      const out: number[] = [];
      let page = -1;
      for (const pl of e.layout.placed) {
        if (pl.ghost || pl.hide) continue;
        if (pl.page !== page) {
          page = pl.page;
          out.push(lineStart(e, pl.para, pl.line));
        }
      }
      return out;
    },
    lineBoxes() {
      const e = getEditor();
      const r = (v: number) => Math.round(v * 100) / 100;
      return e.layout.placed.filter((pl) => !pl.ghost && !pl.hide).map((pl) => {
        const l = e.layout.paras[pl.para].lines[pl.line];
        return {
          start: lineStart(e, pl.para, pl.line), page: pl.page,
          x: r((pl.x + l.x0) / PT), y: r(pl.y / PT), h: r(l.height / PT), base: r(l.baseline / PT),
        };
      });
    },
    xs(offsets) {
      const e = getEditor();
      return offsets.map((offset) => {
        const pos = e.doc.fromOffset(offset);
        const pl = e.layout.paras[pos.p];
        const j = Math.max(0, pl.lines.findIndex((l, i) => pos.o >= l.start && (pos.o < l.end || i === pl.lines.length - 1)));
        const line = pl.lines[j];
        const placed = e.layout.placed[e.layout.main.index[pos.p]?.[j] ?? e.layout.firstPlaced[pos.p]];
        const k = Math.min(pos.o - line.start, line.charX.length - 1);
        return Math.round(((placed?.x ?? 0) + line.x0 + line.charX[Math.max(0, k)]) / PT * 100) / 100;
      });
    },
    formatAt(offset) {
      const e = getEditor();
      const pos = e.doc.fromOffset(offset);
      const p = e.doc.paras[pos.p];
      return format(e, pos.p, pos.o < p.text.length ? propsAt(p, pos.o) : p.mark);
    },
    typingFormat() {
      const e = getEditor();
      return format(e, e.sel.head.p, e.typingProps());
    },
    settled: () => getEditor().fonts.pendingCount === 0,
    load: (name) => getEditor().load(load(name)),
    focus: () => getEditor().focus(),
    reveal(offset) {
      const e = getEditor();
      e.view.scrollIntoView(e.doc.fromOffset(offset), false);
      e.update();
    },
  };
}
export * from "./math/ast";
export { latexToMathml as mathmlFromLatex, lines, seqToMathml } from "./math/mathml";
export { mathElements } from "./view";
export { TEXTFX_ORDER, W14_NS, fxColor, fxColorXml, fxJoin, fxNumbers, fxParts, fxSet, outlineXml, textFxCss, type TextFxPart } from "./textfx";
export { sortDate, sortNumber, sortOrder, type SortKey, type SortSpec, type SortType } from "./sort";
export { MathEditor, buildLinear, explode, getSlot, implode, mathAutoCorrect, slotsOf, type MathPos, type MathRange, type MathSignal, type MathStep } from "./math/edit";
export { seqToLinear } from "./math/linear";
export * from "./bibliography";
