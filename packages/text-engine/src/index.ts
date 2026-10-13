// Public API of the text engine (decision D-017: controlled editing).

export { Editor, type ChangeEvent, type EditorStats } from "./editor";
export { Doc, type Para, type Pos } from "./doc";
export { PAGE, PT, STYLES, type StyleId, type ParaStyle } from "./geometry";
export { MIN_ZOOM, MAX_ZOOM } from "./view";
export { countWords } from "./words";
export { installProbe, type EngineProbe } from "./probe";

import type { Editor } from "./editor";
import type { EngineProbe } from "./probe";
import type { Para } from "./doc";

/** Adapts an editor to the feel-lab measurement interface (see docs/FEEL.md). */
export function probeFor(getEditor: () => Editor, load: (name: string) => Para[]): EngineProbe {
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
      return e.layout.placed.map((pl) => e.doc.toOffset({ p: pl.para, o: e.layout.lines[pl.para][pl.line].start }));
    },
    pages() {
      const e = getEditor();
      const out: number[] = [];
      let page = -1;
      for (const pl of e.layout.placed) {
        if (pl.page !== page) {
          page = pl.page;
          out.push(e.doc.toOffset({ p: pl.para, o: e.layout.lines[pl.para][pl.line].start }));
        }
      }
      return out;
    },
    load: (name) => getEditor().load(load(name)),
    focus: () => getEditor().focus(),
    reveal(offset) {
      const e = getEditor();
      e.view.scrollIntoView(e.doc.fromOffset(offset), false);
      e.update();
    },
  };
}
