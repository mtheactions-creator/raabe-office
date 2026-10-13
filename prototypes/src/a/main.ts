// Prototype A: "native editing". The browser's contenteditable does input, caret,
// selection and text layout; ProseMirror keeps a document model in sync. No pagination:
// the document is one long page, which is one of the known weaknesses of this approach.

import { Schema, Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { keymap } from "prosemirror-keymap";
import { baseKeymap } from "prosemirror-commands";
import { history, undo, redo } from "prosemirror-history";
import { PAGE, STYLES, fontCss, lineHeightPx, PT, type StyleId } from "../shared/geometry";
import { fixture, fixtureName } from "../shared/fixtures";
import { installProbe, markInputHandled } from "../shared/probe";

const schema = new Schema({
  nodes: {
    doc: { content: "para+" },
    para: {
      content: "text*",
      attrs: { style: { default: "normal" } },
      toDOM: (n) => ["p", { class: `ps-${n.attrs.style}` }, 0],
      parseDOM: [{ tag: "p" }],
    },
    text: {},
  },
});

function styleSheet(): string {
  return (Object.keys(STYLES) as StyleId[]).map((id) => {
    const s = STYLES[id];
    return `.ps-${id}{font:${fontCss(s)};color:${s.color};line-height:${lineHeightPx(s)}px;` +
      `margin:${s.spaceBeforePt * PT}px 0 ${s.spaceAfterPt * PT}px;white-space:pre-wrap}`;
  }).join("\n") + `
.page-a{width:${PAGE.width}px;min-height:${PAGE.height}px;box-sizing:border-box;
padding:${PAGE.marginTop}px ${PAGE.marginRight}px ${PAGE.marginBottom}px ${PAGE.marginLeft}px}
.page-a .ProseMirror{outline:none;caret-color:#000}
.page-a .ProseMirror ::selection{background:#c6c6c6}`;
}

function buildDoc(name: string): PMNode {
  return schema.node("doc", null, fixture(name).map((p) =>
    schema.node("para", { style: p.style }, p.text ? schema.text(p.text) : [])));
}

const style = document.createElement("style");
style.textContent = styleSheet();
document.head.append(style);

const workspace = document.getElementById("workspace")!;
const page = document.createElement("div");
page.className = "page page-a";
workspace.append(page);

let lastInputTs: number | null = null;
const plugins = [history(), keymap({ "Mod-z": undo, "Mod-y": redo, "Mod-Shift-z": redo }), keymap(baseKeymap)];

const view = new EditorView(page, {
  state: EditorState.create({ doc: buildDoc(fixtureName()), plugins }),
  dispatchTransaction(tr) {
    view.updateState(view.state.apply(tr));
    if (tr.docChanged) {
      updateStatus();
      if (lastInputTs !== null) markInputHandled(lastInputTs);
    }
    lastInputTs = null;
  },
});
for (const type of ["keydown", "beforeinput"]) {
  view.dom.addEventListener(type, (e) => { lastInputTs ??= e.timeStamp; }, { capture: true });
}

function updateStatus(): void {
  const words = view.state.doc.textBetween(0, view.state.doc.content.size, " ").split(/\s+/).filter(Boolean).length;
  document.getElementById("status-page")!.textContent = "Seite 1 von 1";
  document.getElementById("status-words")!.textContent = `${words.toLocaleString("de-DE")} Wörter`;
}
updateStatus();

// ----- offset <-> ProseMirror position -------------------------------------------------

function offsetToPos(offset: number): number {
  let remaining = offset;
  let result = view.state.doc.content.size - 1;
  let found = false;
  view.state.doc.forEach((node, pos) => {
    if (found) return;
    const len = node.textContent.length;
    if (remaining <= len) { result = pos + 1 + remaining; found = true; return; }
    remaining -= len + 1;
  });
  return result;
}

function posToOffset(pos: number): number {
  let offset = 0;
  let result = 0;
  let found = false;
  view.state.doc.forEach((node, start) => {
    if (found) return;
    const len = node.textContent.length;
    if (pos <= start + 1 + len) { result = offset + Math.max(0, pos - start - 1); found = true; return; }
    offset += len + 1;
  });
  return found ? result : offset - 1;
}

installProbe({
  engine: "a",
  text: () => view.state.doc.textBetween(0, view.state.doc.content.size, "\n"),
  selection: () => ({
    anchor: posToOffset(view.state.selection.anchor),
    head: posToOffset(view.state.selection.head),
  }),
  setSelection(anchor, head) {
    const sel = TextSelection.create(view.state.doc, offsetToPos(anchor), offsetToPos(head));
    view.dispatch(view.state.tr.setSelection(sel).scrollIntoView());
  },
  pointAt(offset) {
    const c = view.coordsAtPos(offsetToPos(offset));
    return { x: c.left, y: c.top, h: c.bottom - c.top };
  },
  lines() {
    // Walk every character and record where its box starts a new row.
    const starts: number[] = [];
    let offset = 0;
    const range = document.createRange();
    for (const p of Array.from(view.dom.children)) {
      let lastTop = -Infinity;
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      let first = true;
      for (let n = walker.nextNode() as Text | null; n; n = walker.nextNode() as Text | null) {
        for (let i = 0; i < n.length; i++) {
          range.setStart(n, i); range.setEnd(n, i + 1);
          const r = range.getClientRects()[0];
          if (!r) continue;
          if (first || r.top > lastTop + 2) { starts.push(offset + i); first = false; }
          lastTop = r.top;
        }
        offset += n.length;
      }
      if (first) starts.push(offset);
      offset += 1;
    }
    return starts;
  },
  pages: () => [0],
  load(name) {
    view.updateState(EditorState.create({ doc: buildDoc(name), plugins }));
    updateStatus();
  },
  focus: () => view.focus(),
  reveal(offset) {
    const c = view.coordsAtPos(offsetToPos(offset));
    const r = workspace.getBoundingClientRect();
    if (c.top < r.top + 40 || c.bottom > r.bottom - 40) workspace.scrollTop += c.top - r.top - r.height / 3;
  },
});

view.focus();
