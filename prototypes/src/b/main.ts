// Prototype B: "controlled editing". Own model, line breaking, pagination, caret and
// selection; the browser only renders text and delivers raw input.

import { fixture, fixtureName } from "../shared/fixtures";
import { installProbe } from "../shared/probe";
import { Editor } from "./editor";

const editor = new Editor(document.getElementById("workspace")!, fixture(fixtureName()));

installProbe({
  engine: "b",
  text: () => editor.doc.text(),
  selection: () => ({
    anchor: editor.doc.toOffset(editor.sel.anchor),
    head: editor.doc.toOffset(editor.sel.head),
  }),
  setSelection(anchor, head) {
    editor.sel = { anchor: editor.doc.fromOffset(anchor), head: editor.doc.fromOffset(head), headUpstream: false };
    editor.update({ scroll: true });
  },
  pointAt(offset) {
    return editor.view.clientPoint(editor.doc.fromOffset(offset));
  },
  lines() {
    return editor.layout.placed.map((pl) =>
      editor.doc.toOffset({ p: pl.para, o: editor.layout.lines[pl.para][pl.line].start }));
  },
  pages() {
    const out: number[] = [];
    let page = -1;
    for (const pl of editor.layout.placed) {
      if (pl.page !== page) {
        page = pl.page;
        out.push(editor.doc.toOffset({ p: pl.para, o: editor.layout.lines[pl.para][pl.line].start }));
      }
    }
    return out;
  },
  load(name) {
    editor.load(fixture(name));
  },
  focus: () => editor.focus(),
  reveal(offset) {
    editor.view.scrollIntoView(editor.doc.fromOffset(offset), false);
    editor.update();
  },
});

editor.focus();
