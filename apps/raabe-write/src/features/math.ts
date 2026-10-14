// Formeln (D-039): equations are drawn as MathML in Cambria Math (laid out by the browser with
// the font's MATH table) and edited in place by our own formula editor, the way Word does it
// (packages/text-engine/src/math/edit.ts). The caret, the selection and Word's equation box are
// drawn from the rendered MathML (every atom says where it is in the tree); every change goes
// into the document's equation tree at once. Einfügen > Formel (Alt+Umschalt+0) and the
// contextual tab "Formel" with Word's groups, galleries and symbols, recorded from Word
// (tools/feel-lab/probe_word_equation_ui.py, probe_word_equation_gallery.py).
// In the file equations are Office Math (OMML), see packages/ooxml/src/math.ts.

import {
  MathEditor, PT, buildLinear, explode, implode, parseLatex, seqText, seqToLatex, seqToLinear, seqToMathml,
  type Editor, type InlineObject, type MathPos, type MathStep, type MNode, type Pos, type Seq,
} from "@raabe/text-engine";
import { h, icon, showDialog, showMenu, type Command, type Control, type MenuEntry, type Tab } from "@raabe/shell";
import type { AppApi } from "./api";
import { BUILT_IN, MATH_GALLERIES, MATH_GALLERY_WIDTHS, SYMBOL_CATEGORIES, type GalleryItem } from "./math-gallery";
import { mathOptionsDialog, userDefaults, userOptions } from "./math-options";

export { BUILT_IN };

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const MATH_COMMANDS: Command[] = [
  cmd("math.new", "Neue Formel einfügen", "equation", { shortcut: "Alt+Umschalt+0" }),
  cmd("math.gallery", "Formel", "equation", { tip: "Formel\nFügen Sie allgemeine mathematische Formeln zu Ihrem Dokument hinzu, z. B. die Fläche eines Kreises oder die binomische Formel." }),
  cmd("math.ink", "Freihandgleichung", "ink-equation", { deferred: true }),
  cmd("math.unicode", "Unicode", "math-unicode", { tip: "Unicode\nVerwenden Sie die lineare Unicode-Formatierung beim Eingeben einer Formel." }),
  cmd("math.latex", "LaTeX", "math-latex", { tip: "LaTeX\nVerwenden Sie die lineare LaTeX-Formatierung beim Eingeben einer Formel." }),
  cmd("math.normalText", "Text", "math-normal-text", { tip: "Text\nGeben Sie in Ihrer Formel Text ohne mathematische Formatierung ein." }),
  cmd("math.convert", "Konvertieren", "math-convert", { tip: "Konvertieren\nKonvertieren Sie die ausgewählte Formel in das professionelle oder lineare Format." }),
  cmd("math.toProfessional", "Aktuell – Professionell", "math-convert"),
  cmd("math.toLinear", "Aktuell – Linear", "math-linear"),
  cmd("math.allProfessional", "Alle – Professionell", "math-convert"),
  cmd("math.allLinear", "Alle – Linear", "math-linear"),
  cmd("math.options", "Formeloptionen..."),
  cmd("math.symbols", "Formelsymbole", "math-symbols"),
  cmd("math.fraction", "Bruch", "math-fraction", { tip: "Bruch\nFügen Sie einen Bruch mit Zähler und Nenner ein." }),
  cmd("math.script", "Hoch/Tief", "math-script", { tip: "Hoch-/Tiefstellung\nFügen Sie eine Hoch- oder Tiefstellung ein." }),
  cmd("math.radical", "Wurzel", "math-radical", { tip: "Wurzel\nFügen Sie eine Wurzel ein." }),
  cmd("math.integral", "Integral", "math-integral", { tip: "Integral\nFügen Sie ein Integral ein." }),
  cmd("math.nary", "Großer Operator", "math-nary", { tip: "Großer Operator\nFügen Sie einen großen Operator ein, z. B. eine Summe oder ein Produkt." }),
  cmd("math.bracket", "Eckige Klammer", "math-bracket", { tip: "Eckige Klammer\nFügen Sie Klammern ein." }),
  cmd("math.function", "Funktion", "math-function", { tip: "Funktion\nFügen Sie eine trigonometrische Funktion ein." }),
  cmd("math.accent", "Akzent", "math-accent", { tip: "Akzent\nFügen Sie einen Akzent ein." }),
  cmd("math.limit", "Grenzwert und Logarithmus", "math-limit", { tip: "Grenzwert und Logarithmus\nFügen Sie einen Grenzwert oder Logarithmus ein." }),
  cmd("math.operator", "Operator", "math-operator", { tip: "Operator\nFügen Sie einen Operator ein." }),
  cmd("math.matrix", "Matrix", "math-matrix", { tip: "Matrix\nFügen Sie eine Matrix ein." }),
];

// ----- MathML rendering ------------------------------------------------------------------------

const templates = new Map<string, Element>();
const sizes = new Map<string, { width: number; height: number; depth: number }>();
let measurer: HTMLElement | null = null;

const treeOf = (o: InlineObject): Seq => o.math ?? (o.latex ? parseLatex(o.latex) : []);
/** Formeloptionen of the document being drawn: nested fractions in display equations smaller. */
let smallFrac = () => false;
/** The drawing's cache key: the tree itself (LaTeX does not carry every property). */
const keyOf = (o: InlineObject) => `${smallFrac() ? "s" : ""}${o.math ? JSON.stringify(o.math) : o.latex ?? ""}`;

/** The MathML of an equation (from the same tree that is saved as OMML). */
function mathTemplate(o: InlineObject, display: boolean): Element {
  const key = `${display ? "D" : "I"}${keyOf(o)}`;
  let t = templates.get(key);
  if (!t) {
    let xml = "";
    try { xml = seqToMathml(treeOf(o), display, true, smallFrac()); } catch { xml = ""; }
    const tpl = document.createElement("template");
    tpl.innerHTML = xml || '<math><mspace width="0.6em"></mspace></math>';
    t = tpl.content.firstElementChild ?? document.createElement("span");
    // Word sets matrix rows closer than the browser's default (measured: word-step4.docx).
    for (const td of Array.from(t.querySelectorAll("mtd"))) (td as HTMLElement).style.paddingBlock = "0.3ex";
    if (templates.size > 2000) templates.clear();
    templates.set(key, t);
  }
  return t;
}

function styled(el: Element, sizePx: number): HTMLElement {
  const e = el as HTMLElement;
  e.style.fontFamily = "'Cambria Math', 'STIX Two Math', 'Latin Modern Math', serif";
  e.style.fontSize = `${sizePx}px`;
  e.style.whiteSpace = "nowrap";
  return e;
}

/** Equation size in points: width, height and depth below the baseline. */
export function measureMath(o: InlineObject, size: number, display: boolean): { width: number; height: number; depth: number } {
  const key = `${display ? "D" : "I"}${size}|${keyOf(o)}`;
  let s = sizes.get(key);
  if (!s) {
    if (!measurer) {
      measurer = h("div", { "aria-hidden": "true", style: "position:absolute;left:-10000px;top:0;visibility:hidden;white-space:nowrap;line-height:normal" });
      document.body.append(measurer);
    }
    const m = styled(mathTemplate(o, display).cloneNode(true) as Element, size * PT);
    const mark = h("span", { style: "display:inline-block;width:0;height:0;vertical-align:baseline" });
    measurer.replaceChildren(m, mark);
    const r = m.getBoundingClientRect(), b = mark.getBoundingClientRect().bottom;
    s = { width: r.width / PT, height: r.height / PT, depth: Math.max(0, r.bottom - b) / PT };
    if (!(s.width > 0)) s.width = size * 0.6;
    if (sizes.size > 4000) sizes.clear();
    sizes.set(key, s);
  }
  return s;
}

export function renderMath(o: InlineObject, sizePx: number, display: boolean): HTMLElement {
  return styled(mathTemplate(o, display).cloneNode(true) as Element, sizePx);
}

/** A rendering for galleries (empty slots as Word's dotted placeholders). */
function preview(seq: Seq, px = 16, display = false): HTMLElement {
  return renderMath({ kind: "math", data: undefined, display: "", latex: `${seqToLatex(seq)}|preview`, math: seq }, px, display);
}

// ----- where things are in a drawn equation ----------------------------------------------------

const pathKey = (path: readonly MathStep[]) => path.map((s) => `${s.i}.${s.slot}`).join("/");
const parsePath = (key: string): MathStep[] => (key ? key.split("/").map((s) => { const [i, slot] = s.split("."); return { i: Number(i), slot }; }) : []);

interface Atom { start: number; count: number; el: Element }

/** Where the atoms and slots of a drawn equation are (from its data-a and data-s marks). */
class Geometry {
  readonly atoms = new Map<string, Atom[]>();
  readonly slots = new Map<string, Element[]>();

  constructor(readonly root: Element) {
    for (const el of Array.from(root.querySelectorAll("[data-a]"))) {
      const [path, start, count] = (el.getAttribute("data-a") ?? "||").split("|");
      const list = this.atoms.get(path) ?? [];
      list.push({ start: Number(start), count: Number(count), el });
      this.atoms.set(path, list);
    }
    for (const list of this.atoms.values()) list.sort((a, b) => a.start - b.start);
    for (const el of Array.from(root.querySelectorAll("[data-s]"))) {
      const k = el.getAttribute("data-s") ?? "";
      const list = this.slots.get(k) ?? [];
      list.push(el);
      this.slots.set(k, list);
    }
  }

  /** The x of the gap before character `k` of an atom (client px). */
  private xIn(a: Atom, k: number): number {
    const r = a.el.getBoundingClientRect();
    if (k <= 0) return r.left;
    if (k >= a.count) return r.right;
    const text = a.el.firstChild;
    if (text && text.nodeType === Node.TEXT_NODE) {
      const chars = [...(text.textContent ?? "")];
      if (chars.length === a.count) {
        const range = document.createRange();
        range.setStart(text, chars.slice(0, k).join("").length);
        range.collapse(true);
        const rr = range.getBoundingClientRect();
        if (rr.left || rr.width) return rr.left;
      }
    }
    return r.left + (r.width * k) / a.count;
  }

  private visible(path: string): Atom[] {
    return (this.atoms.get(path) ?? []).filter((a) => a.count > 0);
  }

  /** The vertical extent of a sequence: its slot, or the line of the equation it is on. */
  private extent(path: string, near?: Atom): { top: number; bottom: number } {
    const els = this.slots.get(path) ?? [this.root];
    // A multi-line equation has one mrow per line, all with the path "".
    const el = near ? els.find((e) => e.contains(near.el)) ?? els[0] : els[0];
    const r = el.getBoundingClientRect();
    return { top: r.top, bottom: r.bottom };
  }

  /** The caret for a position: x and its height (client px). */
  caret(p: MathPos): { x: number; top: number; bottom: number } {
    const key = pathKey(p.path);
    const list = this.visible(key);
    const hit = list.find((a) => p.o >= a.start && p.o < a.start + a.count);
    if (hit) return { x: this.xIn(hit, p.o - hit.start), ...this.extent(key, hit) };
    const before = [...list].reverse().find((a) => a.start + a.count <= p.o);
    if (before) return { x: before.el.getBoundingClientRect().right, ...this.extent(key, before) };
    if (list.length) return { x: list[0].el.getBoundingClientRect().left, ...this.extent(key, list[0]) };
    // An empty slot: at its placeholder (or the prompt of an empty equation).
    const ph = (this.atoms.get(key) ?? [])[0];
    const r = (ph?.el ?? this.slots.get(key)?.[0] ?? this.root).getBoundingClientRect();
    return { x: r.left + 0.5, top: r.top, bottom: r.bottom };
  }

  /** The selection's rectangle: the atoms from…to of a sequence (client px). */
  rect(path: readonly MathStep[], from: number, to: number): DOMRect | null {
    let x1 = Infinity, x2 = -Infinity, y1 = Infinity, y2 = -Infinity;
    for (const a of this.visible(pathKey(path))) {
      const lo = Math.max(from, a.start), hi = Math.min(to, a.start + a.count);
      if (lo >= hi) continue;
      const r = a.el.getBoundingClientRect();
      x1 = Math.min(x1, this.xIn(a, lo - a.start));
      x2 = Math.max(x2, this.xIn(a, hi - a.start));
      y1 = Math.min(y1, r.top);
      y2 = Math.max(y2, r.bottom);
    }
    if (x1 === Infinity) return null;
    return new DOMRect(x1, y1, x2 - x1, y2 - y1);
  }

  /** The gap of a sequence nearest to an x (Up/Down keep the column). */
  nearest(path: readonly MathStep[], x: number, length: number): number {
    const list = this.visible(pathKey(path));
    let o = 0, dist = Infinity;
    for (const a of list) {
      for (let k = 0; k <= a.count; k++) {
        const d = Math.abs(this.xIn(a, k) - x);
        if (d < dist) { dist = d; o = a.start + k; }
      }
    }
    return Math.min(o, length);
  }

  /** The position under a client point: the innermost slot there, the nearest gap in it. */
  pick(x: number, y: number): MathPos {
    let best: { key: string; area: number; depth: number } | null = null;
    for (const [key, els] of this.slots) {
      for (const el of els) {
        const r = el.getBoundingClientRect();
        if (x < r.left - 2 || x > r.right + 2 || y < r.top - 2 || y > r.bottom + 2) continue;
        const depth = key ? key.split("/").length : 0;
        const area = r.width * r.height;
        if (!best || depth > best.depth || (depth === best.depth && area < best.area)) best = { key, area, depth };
      }
    }
    const key = best?.key ?? "";
    const list = this.visible(key);
    let o = 0, dist = Infinity;
    for (const a of list) {
      const r = a.el.getBoundingClientRect();
      // On a multi-line equation only the gaps of the line under the point count.
      const dy = y < r.top ? r.top - y : y > r.bottom ? y - r.bottom : 0;
      for (let k = 0; k <= a.count; k++) {
        const d = Math.abs(this.xIn(a, k) - x) + dy * 4;
        if (d < dist) { dist = d; o = a.start + k; }
      }
    }
    return { path: parsePath(key), o };
  }
}

// ----- the formula editor in place --------------------------------------------------------------

interface Session {
  editor: Editor;
  pos: Pos;
  me: MathEditor;
  /** An undo step exists for this edit (the insertion, or the first change). */
  recorded: boolean;
  /** Inserted just now, nothing typed yet. */
  isNew: boolean;
  /** Word's equation box and the selection (behind the text), handle and arrow (in front). */
  marks: HTMLElement[];
  /** The x Up/Down keep (client px). */
  column: number | null;
}

type LinearFormat = "unicode" | "latex";

function storedFormat(): LinearFormat {
  try { return localStorage.getItem("raabe.math.linear") === "latex" ? "latex" : "unicode"; } catch { return "unicode"; }
}

/** The first empty slot inside the atoms from…to of a sequence (where a template wants input). */
function firstEmpty(me: MathEditor, path: readonly MathStep[], from: number, to: number): MathPos | null {
  const seq = me.seqAt(path);
  for (let i = from; i < to; i++) {
    const n = seq[i];
    if (!n || n.k === "t") continue;
    for (const slot of slotKeys(n)) {
      const p = [...path, { i, slot }];
      const inner = me.seqAt(p);
      if (!inner.length) return { path: p, o: 0 };
      const deeper = firstEmpty(me, p, 0, inner.length);
      if (deeper) return deeper;
    }
  }
  return null;
}

function slotKeys(n: MNode): string[] {
  switch (n.k) {
    case "t": return [];
    case "frac": return ["num", "den"];
    case "scr": return n.pre ? [...(n.sub ? ["sub"] : []), ...(n.sup ? ["sup"] : []), "base"] : ["base", ...(n.sub ? ["sub"] : []), ...(n.sup ? ["sup"] : [])];
    case "rad": return n.deg ? ["deg", "e"] : ["e"];
    case "d": return n.items.map((_, i) => `item${i}`);
    case "nary": return [...(n.sub ? ["sub"] : []), ...(n.sup ? ["sup"] : []), "e"];
    case "func": return ["name", "e"];
    case "lim": return ["e", "lim"];
    case "m": return n.rows.flatMap((r, ri) => r.map((_, ci) => `cell${ri}_${ci}`));
    case "eqArr": return n.rows.map((_, i) => `row${i}`);
    default: return ["e"];
  }
}

const hasStructure = (seq: Seq) => seq.some((n) => n.k !== "t");

/** Equations saved with "Auswahl im Formelkatalog speichern" (this user, this PC). */
function savedEquations(): { name: string; seq: Seq }[] {
  try {
    const v = JSON.parse(localStorage.getItem("raabe.math.saved") ?? "[]") as { name: string; seq: Seq }[];
    return Array.isArray(v) ? v.filter((x) => typeof x?.name === "string" && Array.isArray(x.seq)) : [];
  } catch { return []; }
}

export function mathFeature(api: AppApi) {
  const c = api.commands;
  const ed = () => api.editor();
  let session: Session | null = null;
  let format: LinearFormat = storedFormat();
  /** What Kopieren took from inside an equation (pasted back as structures). */
  let clip: { seq: Seq; text: string } | null = null;

  const view = (s: Session) => s.editor.view;

  /** The session, if it belongs to the active document (another document ends it). */
  const live = (): Session | null => {
    if (session && session.editor !== ed()) end(session);
    return session;
  };

  /** The editing options: linear format, automatic build-up, where limits go. */
  const configure = (e: Editor, me: MathEditor) => {
    const m = e.mathSettings();
    me.latex = format === "latex";
    me.autoBuild = userOptions().autoBuild;
    me.limits = { int: m.intLim === "undOvr", nary: m.naryLim === "undOvr" };
  };

  // ----- drawing ---------------------------------------------------------------------------

  const clearMarks = (s: Session) => { for (const m of s.marks) m.remove(); s.marks = []; };

  /** Draws Word's equation box, the selection and the caret from the rendered equation. */
  const draw = (s: Session) => {
    if (session !== s) return;
    const v = view(s);
    clearMarks(s);
    if (!s.editor.mathAt(s.pos)) { end(s); return; }
    const found = s.editor.mathElement(s.pos);
    if (!found) { v.caretOverride = null; v.refreshCaret(); return; }
    const pageEl = v.pageElement(found.page);
    if (!pageEl) return;
    const g = new Geometry(found.el);
    const toPage = (x: number, y: number) => v.pointAt(x, y);
    const box = found.el.getBoundingClientRect();
    const tl = toPage(box.left, box.top), br = toPage(box.right, box.bottom);
    const layer = h("div", { class: "math-session back" });
    const front = h("div", { class: "math-session front" });
    // Word's box (measured at 150 %): a grey frame over the whole line around the equation,
    // 3.5 px beside it; the Formeloptionen column (14 px, arrow at the bottom) on its right; the
    // move handle (7 x 15 px, a dotted grip) at its top left.
    const ob = s.editor.objectBox(s.pos);
    const top = ob ? ob.y - v.pageTop(found.page) : tl.y - 1;
    const height = ob ? ob.h : br.y - tl.y + 2;
    const left = tl.x - 4, width = br.x - tl.x + 8;
    layer.append(h("div", { class: "math-box", style: `left:${left}px;top:${top}px;width:${width}px;height:${height}px` }));
    const handle = h("div", { class: "math-handle", style: `left:${left - 7}px;top:${top}px;height:15px`, "aria-label": "Feld zum Verschieben" });
    const arrow = h("button", { class: "math-options", style: `left:${left + width - 1}px;top:${top}px;height:${height}px`, "aria-label": "Formeloptionen", "data-tip": "Formeloptionen" },
      h("span", { class: "tri" }));
    arrow.addEventListener("mousedown", (e) => { e.stopPropagation(); e.preventDefault(); optionsMenu(s, arrow); });
    // Dragging the handle moves the equation to where it is dropped.
    handle.style.pointerEvents = "auto";
    handle.addEventListener("mousedown", (e) => {
      e.stopPropagation();
      e.preventDefault();
      const up = (u: MouseEvent) => {
        window.removeEventListener("mouseup", up);
        document.body.style.cursor = "";
        if (session !== s) return;
        const at = s.editor.moveObject(s.pos, s.editor.posAtPoint(u.clientX, u.clientY));
        if (at && (at.p !== s.pos.p || at.o !== s.pos.o)) { s.pos = at; s.recorded = true; }
        draw(s);
      };
      document.body.style.cursor = "move";
      window.addEventListener("mouseup", up);
    });
    front.append(handle, arrow);
    // The selection (structures whole), or the caret.
    if (s.me.hasSelection) {
      const r = s.me.range();
      const rect = g.rect(r.path, r.from, r.to);
      if (rect) {
        const a = toPage(rect.left, rect.top), b = toPage(rect.right, rect.bottom);
        layer.append(h("div", { class: "math-sel", style: `left:${a.x}px;top:${a.y}px;width:${b.x - a.x}px;height:${b.y - a.y}px` }));
      }
      v.caretOverride = { page: found.page, x: 0, y: 0, h: 0 };
    } else {
      const cr = g.caret(s.me.head);
      const a = toPage(cr.x, cr.top), b = toPage(cr.x, cr.bottom);
      v.caretOverride = { page: found.page, x: a.x, y: a.y, h: Math.max(4, b.y - a.y) };
    }
    pageEl.append(layer, front);
    s.marks = [layer, front];
    v.refreshCaret();
  };

  /** The tree into the document (one undo step for the whole edit), then redrawn. */
  const sync = (s: Session) => {
    s.editor.setMath(s.pos, s.me.value(), !s.recorded, true);
    s.recorded = true;
    s.isNew = false;
    draw(s);
    api.refresh();
  };

  // ----- opening and closing ---------------------------------------------------------------

  smallFrac = () => !!ed()?.mathSettings().smallFrac;

  /** A new document starts with the Formeloptionen saved as Standardwerte. */
  const attach = (e: Editor, isNew = false) => {
    const defaults = isNew ? userDefaults() : null;
    if (defaults) e.setMathSettings(defaults, false);
    e.useMath(measureMath, renderMath);
    e.onEditMath = (pos, enter, point) => open(e, pos, enter, point);
  };

  const open = (e: Editor, pos: Pos, enter: "start" | "end" | "click", point?: { x: number; y: number }, inserted = false) => {
    if (session) close(null);
    const o = e.mathAt(pos);
    if (!o) return;
    const me = new MathEditor(treeOf(o), enter === "start" ? "start" : "end");
    configure(e, me);
    const s: Session = { editor: e, pos, me, recorded: inserted, isNew: inserted, marks: [], column: null };
    session = s;
    // The document's own caret waits after the equation; ours is drawn instead.
    e.caretAt({ p: pos.p, o: pos.o + 1 });
    e.interceptKey = (ev) => onKey(s, ev);
    e.interceptText = (text) => onText(s, text);
    e.interceptMouse = (ev) => onMouse(s, ev);
    if (point) {
      const found = e.mathElement(pos);
      if (found) { const p = new Geometry(found.el).pick(point.x, point.y); me.anchor = me.head = p; }
    }
    draw(s);
    api.refresh();
    if (inserted) api.selectTab("equation");
  };

  /** Stops editing (the equation stays as it is). */
  const end = (s: Session) => {
    if (session !== s) return;
    session = null;
    clearMarks(s);
    const e = s.editor;
    e.interceptKey = e.interceptText = e.interceptMouse = null;
    e.view.caretOverride = null;
    e.view.refreshCaret();
  };

  /** Ends editing; the caret goes before or after the equation, or stays where it is (a click elsewhere). */
  const close = (caret: "before" | "after" | null) => {
    const s = session;
    if (!s) return;
    end(s);
    const e = s.editor;
    if (caret && e.mathAt(s.pos)) e.caretAt(caret === "after" ? { p: s.pos.p, o: s.pos.o + 1 } : s.pos);
    api.refresh();
  };

  /** The equation emptied by deleting goes, as in Word. */
  const remove = (s: Session) => {
    end(s);
    s.editor.setMath(s.pos, [], !s.recorded);
    s.editor.caretAt(s.pos);
    api.refresh();
  };

  // ----- keyboard ----------------------------------------------------------------------------

  const geometry = (s: Session) => {
    const found = s.editor.mathElement(s.pos);
    return found ? new Geometry(found.el) : null;
  };

  /** Keys while an equation is edited; false lets the document handle them (after leaving). */
  const onKey = (s: Session, ev: KeyboardEvent): boolean => {
    const me = s.me;
    const ctrl = ev.ctrlKey || ev.metaKey, shift = ev.shiftKey;
    const done = (changed: boolean) => { if (changed) sync(s); else draw(s); return true; };
    const leaveOr = (sig: { leave?: "before" | "after" }) => {
      if (sig.leave) { close(sig.leave); return true; }
      return done(false);
    };
    if (ev.key !== "ArrowUp" && ev.key !== "ArrowDown") s.column = null;
    if (ev.altKey && !ctrl) return false;
    switch (ev.key) {
      case "ArrowRight": return leaveOr(ctrl ? me.word(1, shift) : me.right(shift));
      case "ArrowLeft": return leaveOr(ctrl ? me.word(-1, shift) : me.left(shift));
      case "ArrowUp": case "ArrowDown": {
        const dir = ev.key === "ArrowUp" ? -1 : 1;
        const g = geometry(s);
        if (g && s.column === null) s.column = g.caret(me.head).x;
        const before = JSON.stringify(me.head);
        me.vertical(dir, g && s.column !== null ? (path, len) => g.nearest(path, s.column!, len) : undefined);
        if (JSON.stringify(me.head) === before) {
          // Nothing above or below in the equation: the caret goes on in the text, if there is a
          // line to go to (Word stays in the equation on the first and last line).
          const edge = s.editor.lineEdges(s.pos);
          if ((dir < 0 && edge.first) || (dir > 0 && edge.last)) return done(false);
          close(dir < 0 ? "before" : "after");
          return false;
        }
        return done(false);
      }
      case "Home": me.home(shift); return done(false);
      case "End": me.end(shift); return done(false);
      case "Backspace": {
        const sig = me.backspace();
        if (sig.remove) { remove(s); return true; }
        if (sig.backspaceBefore) { close("before"); return false; }
        return done(true);
      }
      case "Delete": {
        const sig = me.del();
        if (sig.remove) { remove(s); return true; }
        if (sig.deleteAfter) { close("after"); return false; }
        return done(true);
      }
      case "Enter":
        if (shift) { me.newLine(); return done(true); }
        // Eingabe ends the equation: a new paragraph after it.
        close("after");
        return false;
      case "Tab": case "Escape": return true;
      case "PageUp": case "PageDown": close(null); return false;
    }
    if (ctrl) {
      const k = ev.key.toLowerCase();
      if (k === "c" || k === "x") { cutCopy(s, k === "x"); return true; }
      if (k === "v") { void paste(s); return true; }
      if (k === "a") { me.selectAll(); return done(false); }
      // Undo, redo and everything else act on the document: the edit ends first for undo.
      if (k === "z" || k === "y") { close(null); return false; }
      return false;
    }
    return false;
  };

  const paste = async (s: Session) => {
    const { text } = await api.clipboard().catch(() => ({ text: "" }));
    if (session !== s) return;
    if (clip && (text === clip.text || !text)) s.me.insertAtoms(explode(structuredClone(clip.seq) as Seq));
    else if (text) s.me.insertAtoms([...text.replace(/\r?\n/g, " ")].map((ch) => ({ k: "t", s: ch })));
    else return;
    sync(s);
  };

  const onText = (s: Session, text: string): boolean => {
    for (const ch of text) {
      if (ch === "\r" || ch === "\n") continue;
      s.me.type(ch);
    }
    sync(s);
    return true;
  };

  // ----- mouse -------------------------------------------------------------------------------

  const onMouse = (s: Session, ev: MouseEvent): boolean => {
    const found = s.editor.mathElement(s.pos);
    if (!found) { end(s); return false; }
    const r = found.el.getBoundingClientRect();
    const inside = ev.clientX >= r.left - 4 && ev.clientX <= r.right + 4 && ev.clientY >= r.top - 3 && ev.clientY <= r.bottom + 3;
    if (!inside) { close(null); return false; }
    const g = new Geometry(found.el);
    const p = g.pick(ev.clientX, ev.clientY);
    s.column = null;
    if (ev.detail >= 3) s.me.selectAll();
    else if (ev.detail === 2) {
      // A double click selects the letters and digits around the point, else the atom there.
      const seq = s.me.seqAt(p.path);
      const word = (n: MNode | undefined) => n?.k === "t" && /[\p{L}\p{N}]/u.test(n.s);
      let a = p.o, b = p.o;
      while (a > 0 && word(seq[a - 1])) a--;
      while (b < seq.length && word(seq[b])) b++;
      if (a === b) { if (b < seq.length) b++; else if (a > 0) a--; }
      s.me.anchor = { path: p.path, o: a };
      s.me.head = { path: p.path, o: b };
    } else if (ev.shiftKey) s.me.head = p;
    else s.me.anchor = s.me.head = p;
    draw(s);
    const move = (m: MouseEvent) => {
      if (session !== s) return;
      const g2 = geometry(s);
      if (!g2) return;
      s.me.head = g2.pick(m.clientX, m.clientY);
      draw(s);
    };
    const up = () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up); };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
    return true;
  };

  // ----- inserting -------------------------------------------------------------------------------

  /** A gallery entry or symbol: into the equation being edited, else as a new equation. */
  const insert = (seq: Seq) => {
    const e = ed();
    if (!e) return;
    const s = live();
    if (s) {
      const me = s.me;
      const nodes = structuredClone(seq) as Seq;
      if (nodes.length === 1 && nodes[0].k !== "t") me.insertStructure(nodes[0]);
      else {
        const at = me.head.path, from = me.hasSelection ? me.range().from : me.head.o;
        me.insertAtoms(explode(nodes));
        const empty = firstEmpty(me, at, from, me.head.o);
        if (empty) me.anchor = me.head = empty;
      }
      sync(s);
      e.focus();
      return;
    }
    const at = e.insertMath("", seq);
    if (!at) return;
    open(e, at, "end", undefined, true);
    const opened = session as Session | null;
    if (opened) {
      const empty = firstEmpty(opened.me, [], 0, opened.me.root.length);
      if (empty) { opened.me.anchor = opened.me.head = empty; draw(opened); }
    }
  };

  const insertSymbol = (ch: string) => {
    if (!session) { insert([{ k: "t", s: ch }]); return; }
    for (const c2 of ch) session.me.type(c2);
    sync(session);
    session.editor.focus();
  };

  const newEquation = () => {
    const e = ed();
    if (!e || live()) return;
    const at = e.insertMath("");
    if (at) open(e, at, "end", undefined, true);
  };

  // ----- Konvertierungen ---------------------------------------------------------------------

  const linear = (seq: Seq): Seq => {
    const text = format === "latex" ? seqToLatex(seq) : seqToLinear(seq);
    return text ? [{ k: "t", s: text }] : [];
  };
  const professional = (seq: Seq): Seq => {
    if (hasStructure(seq)) return seq;
    const text = seqText(seq);
    if (!text.trim()) return seq;
    return format === "latex" ? parseLatex(text) : buildLinear(text);
  };
  const convertCurrent = (to: "linear" | "professional" | "toggle") => {
    const s = session;
    if (!s) return;
    const cur = s.me.value();
    const target = to === "toggle" ? (hasStructure(cur) ? "linear" : "professional") : to;
    const next = target === "linear" ? linear(cur) : professional(cur);
    s.me = new MathEditor(next, "end");
    configure(s.editor, s.me);
    sync(s);
    s.editor.focus();
  };
  const convertAll = (to: "linear" | "professional") => {
    const e = ed();
    if (!e) return;
    if (session) close(null);
    e.mapAllMath(to === "linear" ? linear : professional);
    api.refresh();
  };

  // ----- Formeloptionen (the arrow at the box) -----------------------------------------------------

  const optionsMenu = (s: Session, anchor: HTMLElement) => {
    const found = s.editor.mathElement(s.pos);
    const display = found?.el.getAttribute("displaystyle") === "true";
    const o = s.editor.mathAt(s.pos);
    const jc = o?.mathJc ?? s.editor.mathSettings().defJc;
    const moved = (at: Pos | null) => { if (at) { s.pos = at; s.recorded = true; } draw(s); s.editor.focus(); api.refresh(); };
    const entries: MenuEntry[] = [
      { label: "Als neue Formel speichern...", icon: "save", run: saveSelection },
      { label: "Professionell", icon: "math-convert", run: () => convertCurrent("professional") },
      { label: "Linear", icon: "math-linear", run: () => convertCurrent("linear") },
      { label: "Alle – Professionell", icon: "math-convert", run: () => convertAll("professional") },
      { label: "Alle – Linear", icon: "math-linear", run: () => convertAll("linear") },
      { label: display ? "In 'Inline' ändern" : "In 'Anzeige' ändern", run: () => moved(s.editor.setMathDisplay(s.pos, !display)) },
    ];
    if (display) {
      entries.push({ label: "Ausrichtung", submenu: ([["left", "Links"], ["right", "Rechts"], ["center", "Zentriert"], ["centerGroup", "Zentriert als Gruppe"]] as const).map(([v, label]) => ({
        label, checked: jc === v, run: () => { s.editor.setMathJc(s.pos, v); moved(null); },
      })) });
    }
    showMenu(c, entries, anchor.getBoundingClientRect());
  };

  const toggleNormal = () => {
    const s = session;
    if (!s) return;
    if (s.me.hasSelection) {
      // Text in the selection becomes plain text (or math again).
      const r = s.me.range();
      const seq = s.me.seqAt(r.path);
      const atoms = seq.slice(r.from, r.to).filter((n) => n.k === "t") as Extract<MNode, { k: "t" }>[];
      const to = !atoms.every((n) => n.nor);
      for (let i = r.from; i < r.to; i++) {
        const n = seq[i];
        if (n.k === "t") seq[i] = to ? { ...n, nor: true } : (({ nor: _n, ...rest }) => rest)(n);
      }
      sync(s);
    } else {
      s.me.normal = !s.me.normal;
      draw(s);
      api.refresh();
    }
    s.editor.focus();
  };


  // ----- the context menu inside an equation (Word's, recorded with probe_word_equation_context.py) --

  /** The structure the menu is about: the selected one, else the one whose slot holds the caret. */
  const contextNode = (me: MathEditor): { node: MNode; parent: MathStep[]; i: number; slot?: string } | null => {
    if (me.hasSelection) {
      const r = me.range();
      const seq = me.seqAt(r.path);
      if (r.to - r.from === 1 && seq[r.from].k !== "t") return { node: seq[r.from], parent: [...r.path], i: r.from };
    }
    const p = me.head.path;
    if (!p.length) return null;
    const last = p[p.length - 1];
    const parent = p.slice(0, -1);
    return { node: me.seqAt(parent)[last.i], parent: [...parent], i: last.i, slot: last.slot };
  };

  /** Replaces a structure by what is left of it (Wurzel entfernen …); the caret goes after. */
  const unwrap = (s: Session, parent: MathStep[], i: number, keep: Seq) => {
    const seq = s.me.seqAt(parent);
    seq.splice(i, 1, ...keep);
    s.me.anchor = s.me.head = { path: parent, o: i + keep.length };
    sync(s);
  };

  const OPERATOR_CHARS = /^[=+−\-±∓×÷⋅<>≤≥≠≈≡→⇒⇔∈∉⊂⊆∪∩]$/;

  /** Word's entries for the structure at the caret (or the operator next to it). */
  const structureEntries = (s: Session): MenuEntry[] => {
    const me = s.me;
    const out: MenuEntry[] = [];
    const change = (fn: () => void) => () => { fn(); sync(s); s.editor.focus(); };
    const ctx = contextNode(me);
    // An operator at the caret: a line break or an alignment point before it.
    const here = me.seqAt(me.head.path);
    const op = [here[me.head.o], here[me.head.o - 1]].find((n) => n?.k === "t" && OPERATOR_CHARS.test(n.s));
    if (op && !me.head.path.length) {
      const at = here.indexOf(op);
      out.push(
        { label: "Manuellen Umbruch einfügen", run: change(() => { here.splice(at, 0, { k: "t", s: "\n" }); me.anchor = me.head = { path: [], o: at + 1 }; }) },
        { label: "An diesem Zeichen ausrichten", run: change(() => { here.splice(at, 0, { k: "t", s: "&" }); me.anchor = me.head = { path: [], o: at + 1 }; }) },
      );
    }
    if (!ctx) return out;
    const { node: n, parent, i } = ctx;
    switch (n.k) {
      case "frac": {
        const kinds: [("skw" | "lin" | undefined), string][] = [[undefined, "Bruch mit waagerechtem Bruchstrich"], ["skw", "Versetzter Bruch mit schrägem Bruchstrich"], ["lin", "Bruch mit schrägem Bruchstrich"]];
        for (const [type, label] of kinds) {
          if (type === n.type || (type === undefined && n.type === undefined)) continue;
          out.push({ label: `In '${label}' ändern`, run: change(() => { if (type) (n as { type?: string }).type = type; else delete (n as { type?: string }).type; }) });
        }
        if (n.type !== "noBar") out.push({ label: "Bruchstrich entfernen", run: change(() => { (n as { type?: string }).type = "noBar"; }) });
        break;
      }
      case "scr":
        if (n.sub && n.sup && !n.pre) out.push({ label: "Skripts vor Basis platzieren", run: change(() => { (n as { pre?: boolean }).pre = true; }) });
        if (n.sub && n.sup && n.pre) out.push({ label: "Skripts nach Basis platzieren", run: change(() => { delete (n as { pre?: boolean }).pre; }) });
        if (n.sub) out.push({ label: "Tiefstellung entfernen", run: () => (n.sup ? change(() => { delete (n as { sub?: Seq }).sub; me.anchor = me.head = { path: parent, o: i + 1 }; })() : unwrap(s, parent, i, n.base)) });
        if (n.sup) out.push({ label: "Hochstellung entfernen", run: () => (n.sub ? change(() => { delete (n as { sup?: Seq }).sup; me.anchor = me.head = { path: parent, o: i + 1 }; })() : unwrap(s, parent, i, n.base)) });
        if (n.sub && n.sup) out.push({ label: "Skripts ausrichten", checked: !!n.aln, run: change(() => { (n as { aln?: boolean }).aln = !n.aln || undefined; }) });
        break;
      case "rad":
        out.push({ label: "Wurzel entfernen", run: () => unwrap(s, parent, i, n.e) });
        if (n.deg) out.push({ label: "Grad ausblenden", run: change(() => { delete (n as { deg?: Seq }).deg; me.anchor = me.head = { path: parent, o: i + 1 }; }) });
        else out.push({ label: "Grad anzeigen", run: change(() => { (n as { deg?: Seq }).deg = []; me.anchor = me.head = { path: [...parent, { i, slot: "deg" }], o: 0 }; }) });
        break;
      case "nary":
        out.push(
          { label: "Positionen von Grenzwert ändern", run: change(() => { (n as { und: boolean }).und = !n.und; }) },
          { label: "N-ären Operator dehnen", checked: !!n.grow, run: change(() => { (n as { grow?: boolean }).grow = !n.grow || undefined; }) },
        );
        break;
      case "d": {
        const k = ctx.slot?.startsWith("item") ? Number(ctx.slot.slice(4)) : n.items.length - 1;
        out.push(
          { label: "Argument davor einfügen", run: change(() => { n.items.splice(k, 0, []); me.anchor = me.head = { path: [...parent, { i, slot: `item${k}` }], o: 0 }; }) },
          { label: "Argument dahinter einfügen", run: change(() => { n.items.splice(k + 1, 0, []); me.anchor = me.head = { path: [...parent, { i, slot: `item${k + 1}` }], o: 0 }; }) },
          { label: "Einschlusszeichen entfernen", run: () => unwrap(s, parent, i, n.items.flatMap((x, j) => (j ? [{ k: "t", s: "|" } as MNode, ...x] : x))) },
        );
        if (n.beg) out.push({ label: "Öffnende eckige Klammer entfernen", run: change(() => { (n as { beg: string }).beg = ""; }) });
        if (n.end) out.push({ label: "Schließende eckige Klammer entfernen", run: change(() => { (n as { end: string }).end = ""; }) });
        out.push(
          { label: "Eckige Klammern dehnen", checked: !n.noGrow, run: change(() => { (n as { noGrow?: boolean }).noGrow = !n.noGrow || undefined; }) },
          { label: "Eckige Klammern an Argumenthöhe anpassen", checked: n.shp === "match", run: change(() => { (n as { shp?: string }).shp = n.shp ? undefined : "match"; }) },
        );
        break;
      }
      case "m": {
        const [r, col] = ctx.slot?.startsWith("cell") ? ctx.slot.slice(4).split("_").map(Number) : [0, 0];
        const cols = Math.max(1, ...n.rows.map((x) => x.length));
        const cellPath = (ri: number, ci: number) => ({ path: [...parent, { i, slot: `cell${ri}_${ci}` }], o: 0 });
        out.push(
          { label: "Platzhalter ausblenden", checked: !!n.plcHide, run: change(() => { (n as { plcHide?: boolean }).plcHide = !n.plcHide || undefined; }) },
          { label: "Einfügen", submenu: [
            { label: "Zeile vorher einfügen", run: change(() => { n.rows.splice(r, 0, Array.from({ length: cols }, () => [])); me.anchor = me.head = cellPath(r, col); }) },
            { label: "Zeile danach einfügen", run: change(() => { n.rows.splice(r + 1, 0, Array.from({ length: cols }, () => [])); me.anchor = me.head = cellPath(r + 1, col); }) },
            { label: "Spalte vorher einfügen", run: change(() => { for (const row of n.rows) row.splice(col, 0, []); if (n.mcJc) n.mcJc.splice(col, 0, "center"); me.anchor = me.head = cellPath(r, col); }) },
            { label: "Spalte nachher einfügen", run: change(() => { for (const row of n.rows) row.splice(col + 1, 0, []); if (n.mcJc) n.mcJc.splice(col + 1, 0, "center"); me.anchor = me.head = cellPath(r, col + 1); }) },
          ] },
          { label: "Löschen", submenu: [
            { label: "Zeile löschen", disabled: n.rows.length < 2, run: change(() => { n.rows.splice(r, 1); me.anchor = me.head = cellPath(Math.min(r, n.rows.length - 1), col); }) },
            { label: "Spalte löschen", disabled: cols < 2, run: change(() => { for (const row of n.rows) row.splice(col, 1); n.mcJc?.splice(col, 1); me.anchor = me.head = cellPath(r, Math.min(col, cols - 2)); }) },
          ] },
          { label: "Matrixausrichtung", submenu: ([["top", "Oben"], [undefined, "Zentriert"], ["bot", "Unten"]] as const).map(([v, label]) => ({
            label, checked: n.baseJc === v, run: change(() => { (n as { baseJc?: string }).baseJc = v; }),
          })) },
          { label: "Spaltenausrichtung", submenu: ([["left", "Links"], ["center", "Zentriert"], ["right", "Rechts"]] as const).map(([v, label]) => ({
            label, checked: (n.mcJc?.[col] ?? "center") === v,
            run: change(() => { const jc = n.mcJc ?? Array.from({ length: cols }, () => "center" as const); jc[col] = v; (n as { mcJc?: typeof jc }).mcJc = jc.some((x) => x !== "center") ? jc : undefined; }),
          })) },
        );
        break;
      }
      case "lim":
        out.push(
          { label: n.upper ? "Grenzwert unter Basis platzieren" : "Grenzwert über Basis platzieren", run: change(() => { (n as { upper: boolean }).upper = !n.upper; }) },
          { label: "Grenzwert entfernen", run: () => unwrap(s, parent, i, n.e) },
        );
        break;
      default: break;
    }
    return out;
  };

  /** Ausschneiden / Kopieren of the selection inside the equation. */
  const cutCopy = (s: Session, cut: boolean) => {
    const me = s.me;
    if (!me.hasSelection) return;
    const r = me.range();
    const seq = structuredClone(me.seqAt(r.path).slice(r.from, r.to)) as Seq;
    const text = userOptions().copyMathml ? seqToMathml(seq, false, false) : seqToLinear(seq);
    clip = { seq, text };
    void api.platform.clipboardWrite(text).catch(() => {});
    if (cut) { const sig = me.del(); if (sig.remove) remove(s); else sync(s); }
  };

  /**
   * Word's context menu while an equation is edited: a right-click moves the caret there
   * (unless it is in the selection); null when no equation is edited.
   */
  const contextMenu = (x: number, y: number): MenuEntry[] | null => {
    const s = live();
    if (!s) return null;
    const found = s.editor.mathElement(s.pos);
    if (found) {
      const r = found.el.getBoundingClientRect();
      if (x < r.left - 4 || x > r.right + 4 || y < r.top - 3 || y > r.bottom + 3) { close(null); return null; }
      if (!s.me.hasSelection) { const p = new Geometry(found.el).pick(x, y); s.me.anchor = s.me.head = p; draw(s); }
    }
    const structure = structureEntries(s);
    return [
      { label: "Ausschneiden", icon: "cut", disabled: !s.me.hasSelection, run: () => cutCopy(s, true) },
      { label: "Kopieren", icon: "copy", disabled: !s.me.hasSelection, run: () => cutCopy(s, false) },
      { label: "Einfügeoptionen:", icons: [{ icon: "paste", tip: "Einfügen", run: () => void paste(s) }] },
      "-",
      { label: "Professionell", icon: "math-convert", run: () => convertCurrent("professional") },
      { label: "Linear", icon: "math-linear", run: () => convertCurrent("linear") },
      { label: "Als neue Formel speichern...", icon: "save", run: saveSelection },
      ...(structure.length ? ["-" as const, ...structure] : []),
      "-",
      { cmd: "font.dialog", label: "Schriftart..." },
      { cmd: "para.dialog", label: "Absatz..." },
    ];
  };

  // ----- commands ----------------------------------------------------------------------------

  const editing = () => !!live();
  c.implement("insert.equation", { run: newEquation, enabled: () => !!ed() });
  c.implement("math.new", { run: newEquation, enabled: () => !!ed() });
  c.implement("math.gallery", { run: () => {}, enabled: () => !!ed() });
  c.implement("math.unicode", {
    run: () => { format = "unicode"; if (session) session.me.latex = false; try { localStorage.setItem("raabe.math.linear", format); } catch { /* private mode */ } api.refresh(); session?.editor.focus(); },
    enabled: editing, checked: () => format === "unicode",
  });
  c.implement("math.latex", {
    run: () => { format = "latex"; if (session) session.me.latex = true; try { localStorage.setItem("raabe.math.linear", format); } catch { /* private mode */ } api.refresh(); session?.editor.focus(); },
    enabled: editing, checked: () => format === "latex",
  });
  c.implement("math.normalText", { run: toggleNormal, enabled: editing, checked: () => !!session?.me.normal });
  c.implement("math.convert", { run: () => convertCurrent("toggle"), enabled: editing });
  c.implement("math.toProfessional", { run: () => convertCurrent("professional"), enabled: editing });
  c.implement("math.toLinear", { run: () => convertCurrent("linear"), enabled: editing });
  c.implement("math.allProfessional", { run: () => convertAll("professional"), enabled: () => !!ed() });
  c.implement("math.allLinear", { run: () => convertAll("linear"), enabled: () => !!ed() });
  c.implement("math.options", {
    run: () => {
      const e = ed();
      if (!e) return;
      void mathOptionsDialog(e).then(() => {
        if (session) { configure(session.editor, session.me); draw(session); session.editor.focus(); }
        api.refresh();
      });
    },
    enabled: () => !!ed(),
  });
  for (const id of ["math.symbols", ...Object.keys(MATH_GALLERIES)]) c.implement(id, { run: () => {}, enabled: () => !!ed() });

  // ----- galleries ----------------------------------------------------------------------------

  /** A gallery entry drawn as Word draws it (display style, Word's tile size with its 4 px margin). */
  const tile = (g: GalleryItem, size: [number, number], onPick: () => void) =>
    h("button", { class: "mtile", style: `width:${size[0] - 8}px;height:${size[1] - 8}px`, "data-tip": g.label, "aria-label": g.label, onclick: () => onPick() },
      preview(g.seq, 15, true));

  const structurePanel = (id: string) => (close2: () => void) => {
    const panel = h("div", { class: "rpanel mathpanel", style: `width:${MATH_GALLERY_WIDTHS[id] ?? 340}px` });
    for (const grp of MATH_GALLERIES[id] ?? []) {
      panel.append(h("div", { class: "title" }, grp.heading));
      panel.append(h("div", { class: "mgrid" }, ...grp.items.map((x) => tile(x, x.size ?? grp.tile, () => { close2(); insert(x.seq); }))));
    }
    return panel;
  };

  let category = 0;
  const symbolGrid = (cat: number, close2: () => void) => {
    const box = h("div", { class: "msymgroups" });
    for (const grp of SYMBOL_CATEGORIES[cat]?.groups ?? []) {
      // Word heads the first group with the category's name.
      box.append(h("div", { class: "title" }, grp.heading || SYMBOL_CATEGORIES[cat].name));
      box.append(h("div", { class: "msyms" }, ...grp.items.map(([ch, label]) =>
        h("button", { class: "msym", "data-tip": label, "aria-label": label, onclick: () => { close2(); insertSymbol(ch); } }, ch))));
    }
    return box;
  };

  const symbolPanel = (close2: () => void) => {
    const panel = h("div", { class: "rpanel mathpanel msympanel" });
    const name = h("span", {}, SYMBOL_CATEGORIES[category]?.name ?? "");
    const filter = h("button", { class: "mfilter", "aria-label": "Katalogfilter" }, name, icon("chevron-down"));
    let grid = symbolGrid(category, close2);
    const list = h("div", { class: "menu mcatlist", role: "menu", hidden: true });
    for (const [i, cat] of SYMBOL_CATEGORIES.entries()) {
      list.append(h("button", { role: "menuitemradio", class: i === category ? "on" : "", onclick: () => {
        category = i;
        name.textContent = cat.name;
        for (const b of Array.from(list.children)) b.classList.toggle("on", b === list.children[i]);
        const next = symbolGrid(category, close2);
        grid.replaceWith(next);
        grid = next;
        list.hidden = true;
        api.refresh();
      } }, cat.name));
    }
    filter.onclick = () => { list.hidden = !list.hidden; };
    panel.append(h("div", { class: "mfilterrow" }, filter, list), grid);
    return panel;
  };

  /** The ribbon's symbol gallery: the current category, two rows as in Word. */
  const ribbonSymbols = (): { ch: string; label: string }[] =>
    (SYMBOL_CATEGORIES[category]?.groups ?? []).flatMap((g) => g.items.map(([ch, label]) => ({ ch, label })));

  /** Einfügen > Formel and Formel > Formel: Word's built-in equations and the saved ones. */
  const galleryPanel = (close2: () => void, fromInsert = false) => {
    const panel = h("div", { class: "rpanel mathpanel mathgallery" });
    const entry = (name: string, seq: Seq) => h("button", { class: "mbuilt", "aria-label": name, onclick: () => { close2(); insert(seq); } },
      h("div", { class: "mname" }, name), h("div", { class: "mprev" }, preview(seq, 16, true)));
    const saved = savedEquations();
    if (saved.length) panel.append(h("div", { class: "title" }, "Allgemein"), ...saved.map((x) => entry(x.name, x.seq)));
    panel.append(h("div", { class: "title" }, "Integriert"), ...BUILT_IN.map((b) => entry(b.name, b.seq)));
    const link = (label: string, run: (() => void) | null, tip?: string) =>
      h("button", { class: "link", disabled: !run, ...(tip ? { "data-tip": tip } : {}), onclick: () => { close2(); run?.(); } }, label);
    panel.append(h("hr"), link("Weitere Formeln von Office.com", null, "Weitere Formeln von Office.com\nNicht verfügbar: RaabeWrite arbeitet offline."));
    if (fromInsert) panel.append(link("Neue Formel einfügen", newEquation), link("Freihandgleichung einfügen", null, "Freihandgleichung einfügen\nZurückgestellt"));
    panel.append(link("Auswahl im Formelkatalog speichern...", session ? saveSelection : null));
    return panel;
  };

  /** Auswahl im Formelkatalog speichern: the equation (or its selection) under a name, for this user. */
  const saveSelection = () => {
    const s = session;
    if (!s) return;
    const r = s.me.range();
    const seq = s.me.hasSelection ? (structuredClone(s.me.seqAt(r.path).slice(r.from, r.to)) as Seq) : s.me.value();
    const input = h("input", { type: "text", value: "", style: "width:260px" }) as HTMLInputElement;
    void showDialog({
      title: "Neuen Baustein erstellen",
      body: h("div", { class: "form" }, h("label", {}, "Name:", input), h("div", {}, "Katalog: Formeln"), h("div", {}, "Kategorie: Allgemein")),
      buttons: [{ label: "OK", value: "ok", primary: true }, { label: "Abbrechen", value: "cancel" }],
      cancel: "cancel",
    }).then((v) => {
      const name = input.value.trim();
      if (v !== "ok" || !name) return;
      const list = savedEquations().filter((x) => x.name !== name);
      list.push({ name, seq: implode(seq) });
      try { localStorage.setItem("raabe.math.saved", JSON.stringify(list)); } catch { /* private mode */ }
    });
  };

  const dd = (id: string, key: string, popup: (close: () => void) => HTMLElement, size: "large" | "medium" = "large"): Control =>
    ({ type: "dropdown", cmd: id, key, size, menu: [], popup });
  const symbolTiles: Control = {
    type: "tiles", cmd: "math.symbols", key: "O", count: 40, visible: 20, size: [20, 20], cls: "msymtiles",
    render: (i) => { const x = ribbonSymbols()[i]; return x ? h("span", { class: "msymtile" }, x.ch) : null; },
    tip: (i) => ribbonSymbols()[i]?.label ?? "",
    pick: (i) => { const x = ribbonSymbols()[i]; if (x) insertSymbol(x.ch); },
    redraw: () => String(category),
    panel: symbolPanel,
  };
  const tab: Tab = {
    id: "equation", label: "Formel", key: "JE", contextual: true, shrink: ["structures", "symbols", "conversions", "tools"],
    groups: [
      { id: "tools", label: "Tools", key: "ZT", icon: "equation",
        controls: [{ type: "stack", items: [dd("math.gallery", "A", galleryPanel, "medium"), { type: "button", cmd: "math.ink", key: "H", size: "medium" }] }] },
      { id: "conversions", label: "Konvertierungen", key: "Z1", icon: "math-convert", launcher: { cmd: "math.options", key: "F" },
        controls: [
          { type: "stack", items: [
            { type: "button", cmd: "math.unicode", key: "UF", size: "medium" },
            { type: "button", cmd: "math.latex", key: "TX", size: "medium" },
            { type: "button", cmd: "math.normalText", key: "N", size: "medium" },
          ] },
          { type: "split", cmd: "math.convert", key: "K1", size: "large", menu: [{ cmd: "math.toProfessional" }, { cmd: "math.toLinear" }, { cmd: "math.allProfessional" }, { cmd: "math.allLinear" }] },
        ] },
      { id: "symbols", label: "Symbole", key: "ZS", icon: "math-symbols", controls: [symbolTiles] },
      { id: "structures", label: "Strukturen", key: "ZU", icon: "math-fraction",
        controls: [["math.fraction", "B"], ["math.script", "T1"], ["math.radical", "W"], ["math.integral", "I"], ["math.nary", "R"], ["math.bracket", "C"],
          ["math.function", "U1"], ["math.accent", "K2"], ["math.limit", "G"], ["math.operator", "E"], ["math.matrix", "M"]].map(([id, k]) => dd(id, k, structurePanel(id))) },
    ],
  };
  const insertMenu: MenuEntry[] = [
    ...BUILT_IN.map((b) => ({ label: b.name, run: () => insert(b.seq), preview: h("div", { class: "mprev" }, preview(b.seq, 15, true)) })),
    "-", { cmd: "math.new" },
  ];
  return {
    attach, tab, insertMenu, insertPanel: (close2: () => void) => galleryPanel(close2, true), editing, contextMenu,
    focus: () => session?.editor.focus(), close: () => close(null),
  };
}
