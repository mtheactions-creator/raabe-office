// Formeln: equations are shown as MathML in Cambria Math (Word's equation font, laid out by
// the browser with the font's MATH table) and edited in place with MathLive (MIT, offline).
// Einfügen > Formel (Alt+Umschalt+0) and the contextual tab "Formel" with Word's structures.
// In the file they are Office Math (OMML), see packages/ooxml/src/math.ts.

import type { Editor, Pos } from "@raabe/text-engine";
import { PT } from "@raabe/text-engine";
import { latexToMathml } from "@raabe/ooxml";
import { h, type Command, type Control, type MenuEntry, type Tab } from "@raabe/shell";
import type { AppApi } from "./api";

type MathLive = typeof import("mathlive");
type Field = InstanceType<MathLive["MathfieldElement"]>;

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const MATH_COMMANDS: Command[] = [
  cmd("math.new", "Neue Formel einfügen", "equation", { shortcut: "Alt+Umschalt+0" }),
  cmd("math.gallery", "Formel", "equation"),
  cmd("math.normalText", "Normaler Text", "math-normal-text", { tip: "Normaler Text\nText in der Formel ohne mathematische Formatierung schreiben." }),
  cmd("math.symbols", "Symbole", "math-symbols"),
  cmd("math.fraction", "Bruch", "math-fraction"),
  cmd("math.script", "Skript", "math-script"),
  cmd("math.radical", "Radikal", "math-radical"),
  cmd("math.integral", "Integral", "math-integral"),
  cmd("math.nary", "Großer Operator", "math-nary"),
  cmd("math.bracket", "Klammer", "math-bracket"),
  cmd("math.function", "Funktion", "math-function"),
  cmd("math.accent", "Akzent", "math-accent"),
  cmd("math.limit", "Grenzwert und Logarithmus", "math-limit"),
  cmd("math.operator", "Operator", "math-operator"),
  cmd("math.matrix", "Matrix", "math-matrix"),
];

/** Word's built-in equations (Einfügen > Formel). */
export const BUILT_IN: { name: string; latex: string }[] = [
  { name: "Binomischer Lehrsatz", latex: "(x+a)^n=\\sum_{k=0}^{n}\\binom{n}{k}x^ka^{n-k}" },
  { name: "Erweiterung einer Summe", latex: "(1+x)^n=1+\\frac{nx}{1!}+\\frac{n(n-1)x^2}{2!}+\\cdots" },
  { name: "Fläche eines Kreises", latex: "A=\\pi r^2" },
  { name: "Fourier-Reihe", latex: "f(x)=a_0+\\sum_{n=1}^{\\infty}\\left(a_n\\cos\\frac{n\\pi x}{L}+b_n\\sin\\frac{n\\pi x}{L}\\right)" },
  { name: "Pythagoreischer Lehrsatz", latex: "a^2+b^2=c^2" },
  { name: "Quadratische Formel", latex: "x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}" },
  { name: "Taylor-Entwicklung", latex: "e^x=1+\\frac{x}{1!}+\\frac{x^2}{2!}+\\frac{x^3}{3!}+\\cdots,\\quad-\\infty<x<\\infty" },
  { name: "Trigonometrische Identität 1", latex: "\\sin\\alpha\\pm\\sin\\beta=2\\sin\\frac{1}{2}(\\alpha\\pm\\beta)\\cos\\frac{1}{2}(\\alpha\\mp\\beta)" },
  { name: "Trigonometrische Identität 2", latex: "\\cos\\alpha+\\cos\\beta=2\\cos\\frac{1}{2}(\\alpha+\\beta)\\cos\\frac{1}{2}(\\alpha-\\beta)" },
];

const P = "#?";
/** Strukturen: what each button of the Formel tab offers (MathLive placeholders #?). */
const STRUCTURES: Record<string, { heading?: string; items: string[] }[]> = {
  "math.fraction": [{ heading: "Bruch", items: [`\\frac{${P}}{${P}}`, `{${P}}/{${P}}`, `\\tfrac{${P}}{${P}}`] },
    { heading: "Häufig verwendete Brüche", items: ["\\frac{dy}{dx}", "\\frac{\\Delta y}{\\Delta x}", "\\frac{\\partial y}{\\partial x}", "\\frac{\\pi}{2}", "\\frac{1}{2}"] }],
  "math.script": [{ heading: "Hoch- und Tiefstellungen", items: [`${P}^{${P}}`, `${P}_{${P}}`, `${P}_{${P}}^{${P}}`, `{}_{${P}}^{${P}}${P}`] },
    { heading: "Häufig verwendete Hoch- und Tiefstellungen", items: ["x^{2}", "e^{-i\\omega t}", "x_{n}", "x_{1}^{2}"] }],
  "math.radical": [{ heading: "Wurzeln", items: [`\\sqrt{${P}}`, `\\sqrt[${P}]{${P}}`, `\\sqrt[3]{${P}}`] },
    { heading: "Häufig verwendete Wurzeln", items: ["\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}", "\\sqrt{a^2+b^2}"] }],
  "math.integral": [{ heading: "Integrale", items: [`\\int ${P}`, `\\int_{${P}}^{${P}}${P}`, `\\iint ${P}`, `\\iint_{${P}}^{${P}}${P}`, `\\iiint ${P}`, `\\oint ${P}`] },
    { heading: "Differenziale", items: ["dx", "dy", "d\\theta"] }],
  "math.nary": [{ heading: "Summen", items: [`\\sum ${P}`, `\\sum_{${P}}^{${P}}${P}`, `\\sum_{${P}}${P}`] },
    { heading: "Produkte und Koprodukte", items: [`\\prod_{${P}}^{${P}}${P}`, `\\coprod_{${P}}^{${P}}${P}`] },
    { heading: "Vereinigungen und Schnittmengen", items: [`\\bigcup_{${P}}^{${P}}${P}`, `\\bigcap_{${P}}^{${P}}${P}`] },
    { heading: "Häufig verwendete große Operatoren", items: ["\\sum_{k=0}^{n}\\binom{n}{k}", "\\sum_{i=1}^{n}x_i"] }],
  "math.bracket": [{ heading: "Klammern", items: [`\\left(${P}\\right)`, `\\left[${P}\\right]`, `\\left\\{${P}\\right\\}`, `\\left|${P}\\right|`, `\\left\\Vert${P}\\right\\Vert`,
    `\\left\\langle${P}\\right\\rangle`, `\\left\\lfloor${P}\\right\\rfloor`, `\\left\\lceil${P}\\right\\rceil`, `\\left[${P}\\right[`, `\\left]${P}\\right]`] },
    { heading: "Fallunterscheidungen und Stapel", items: [`\\begin{cases}${P}\\\\${P}\\end{cases}`, `\\binom{${P}}{${P}}`] }],
  "math.function": [{ heading: "Trigonometrische Funktionen", items: [`\\sin ${P}`, `\\cos ${P}`, `\\tan ${P}`, `\\cot ${P}`, `\\sin^{-1}${P}`, `\\cos^{-1}${P}`, `\\tan^{-1}${P}`] },
    { heading: "Häufig verwendete Funktionen", items: ["\\sin\\theta", "\\cos 2x", "\\tan\\theta=\\frac{\\sin\\theta}{\\cos\\theta}"] }],
  "math.accent": [{ heading: "Akzente", items: [`\\dot{${P}}`, `\\ddot{${P}}`, `\\hat{${P}}`, `\\check{${P}}`, `\\tilde{${P}}`, `\\bar{${P}}`, `\\vec{${P}}`,
    `\\overrightarrow{${P}}`, `\\overline{${P}}`, `\\underline{${P}}`, `\\overbrace{${P}}`, `\\underbrace{${P}}`, `\\boxed{${P}}`] }],
  "math.limit": [{ heading: "Funktionen", items: [`\\log_{${P}}${P}`, `\\log ${P}`, `\\lim_{${P}}${P}`, `\\min_{${P}}${P}`, `\\max_{${P}}${P}`, `\\ln ${P}`] },
    { heading: "Häufig verwendete Funktionen", items: ["\\lim_{n\\to\\infty}\\left(1+\\frac{1}{n}\\right)^n", "\\max_{0\\le x\\le 1}xe^{-x^2}"] }],
  "math.operator": [{ heading: "Grundlegende Operatoren", items: ["\\coloneqq", "==", "+=", "-=", "\\overset{\\text{def}}{=}", "\\overset{\\Delta}{=}", "\\overset{!}{=}"] },
    { heading: "Operatorstrukturen", items: [`\\overset{${P}}{\\rightarrow}`, `\\underset{${P}}{\\rightarrow}`, `\\overset{${P}}{\\leftarrow}`, `\\overset{${P}}{\\Rightarrow}`] }],
  "math.matrix": [{ heading: "Leere Matrizen", items: [`\\begin{matrix}${P}&${P}\\end{matrix}`, `\\begin{matrix}${P}\\\\${P}\\end{matrix}`, `\\begin{matrix}${P}&${P}\\\\${P}&${P}\\end{matrix}`,
    `\\begin{matrix}${P}&${P}&${P}\\\\${P}&${P}&${P}\\\\${P}&${P}&${P}\\end{matrix}`] },
    { heading: "Matrizen mit Klammern", items: [`\\begin{pmatrix}${P}&${P}\\\\${P}&${P}\\end{pmatrix}`, `\\begin{bmatrix}${P}&${P}\\\\${P}&${P}\\end{bmatrix}`,
      `\\begin{vmatrix}${P}&${P}\\\\${P}&${P}\\end{vmatrix}`, `\\begin{pmatrix}${P}&${P}&${P}\\\\${P}&${P}&${P}\\\\${P}&${P}&${P}\\end{pmatrix}`] },
    { heading: "Einheitsmatrizen", items: ["\\begin{pmatrix}1&0\\\\0&1\\end{pmatrix}", "\\begin{pmatrix}1&0&0\\\\0&1&0\\\\0&0&1\\end{pmatrix}"] }],
};

/** Symbole (Grundlegende Mathematik, griechische Buchstaben): character → what MathLive inserts. */
const SYMBOLS: { heading: string; items: [string, string][] }[] = [
  { heading: "Grundlegende Mathematik", items: [["±", "\\pm"], ["∞", "\\infty"], ["=", "="], ["≠", "\\ne"], ["~", "\\sim"], ["×", "\\times"], ["÷", "\\div"], ["!", "!"],
    ["∝", "\\propto"], ["<", "<"], ["≪", "\\ll"], [">", ">"], ["≫", "\\gg"], ["≤", "\\le"], ["≥", "\\ge"], ["∓", "\\mp"], ["≅", "\\cong"], ["≈", "\\approx"],
    ["≡", "\\equiv"], ["∀", "\\forall"], ["∂", "\\partial"], ["√", "\\sqrt{#?}"], ["∪", "\\cup"], ["∩", "\\cap"], ["∅", "\\emptyset"], ["%", "\\%"], ["°", "\\degree"],
    ["∆", "\\Delta"], ["∇", "\\nabla"], ["∃", "\\exists"], ["∄", "\\nexists"], ["∈", "\\in"], ["∉", "\\notin"], ["⊂", "\\subset"], ["⊆", "\\subseteq"],
    ["←", "\\leftarrow"], ["↑", "\\uparrow"], ["→", "\\rightarrow"], ["↓", "\\downarrow"], ["↔", "\\leftrightarrow"], ["⇒", "\\Rightarrow"], ["⇔", "\\Leftrightarrow"],
    ["∴", "\\therefore"], ["+", "+"], ["−", "-"], ["¬", "\\neg"], ["∧", "\\land"], ["∨", "\\lor"], ["⋅", "\\cdot"], ["∘", "\\circ"], ["⋯", "\\cdots"], ["…", "\\ldots"],
    ["⋮", "\\vdots"], ["⋱", "\\ddots"], ["ℕ", "\\mathbb{N}"], ["ℤ", "\\mathbb{Z}"], ["ℚ", "\\mathbb{Q}"], ["ℝ", "\\mathbb{R}"], ["ℂ", "\\mathbb{C}"]] },
  { heading: "Griechische Buchstaben", items: [["α", "\\alpha"], ["β", "\\beta"], ["γ", "\\gamma"], ["δ", "\\delta"], ["ε", "\\varepsilon"], ["ζ", "\\zeta"], ["η", "\\eta"],
    ["θ", "\\theta"], ["ι", "\\iota"], ["κ", "\\kappa"], ["λ", "\\lambda"], ["μ", "\\mu"], ["ν", "\\nu"], ["ξ", "\\xi"], ["π", "\\pi"], ["ρ", "\\rho"], ["σ", "\\sigma"],
    ["τ", "\\tau"], ["φ", "\\varphi"], ["χ", "\\chi"], ["ψ", "\\psi"], ["ω", "\\omega"], ["Γ", "\\Gamma"], ["Δ", "\\Delta"], ["Θ", "\\Theta"], ["Λ", "\\Lambda"],
    ["Π", "\\Pi"], ["Σ", "\\Sigma"], ["Φ", "\\Phi"], ["Ψ", "\\Psi"], ["Ω", "\\Omega"]] },
];

// ----- MathLive (loaded when an equation is first edited) ------------------------------------

let ml: MathLive | null = null;
let loading: Promise<MathLive> | null = null;

function loadMathLive(): Promise<MathLive> {
  if (ml) return Promise.resolve(ml);
  loading ??= import("mathlive").then((m) => {
    m.MathfieldElement.fontsDirectory = new URL("mathlive-fonts/", location.href).href;
    m.MathfieldElement.soundsDirectory = null;
    m.MathfieldElement.keypressSound = null;
    m.MathfieldElement.plonkSound = null;
    m.MathfieldElement.locale = "de";
    ml = m;
    return m;
  });
  return loading;
}

// ----- MathML rendering ------------------------------------------------------------------------

const templates = new Map<string, Element>();
const sizes = new Map<string, { width: number; height: number; depth: number }>();
let measurer: HTMLElement | null = null;

/** The MathML of an equation (from the same parse as the OMML that is saved). */
function mathTemplate(latex: string, display: boolean): Element {
  const key = `${display ? "D" : "I"}${latex}`;
  let t = templates.get(key);
  if (!t) {
    let xml = "";
    try { xml = latex.trim() ? latexToMathml(latex, display) : ""; } catch { xml = ""; }
    const tpl = document.createElement("template");
    tpl.innerHTML = xml || '<math><mspace width="0.6em"></mspace></math>';
    t = tpl.content.firstElementChild ?? document.createElement("span");
    // Word sets matrix rows closer than the browser's default (measured: word-step4.docx).
    for (const td of Array.from(t.querySelectorAll("mtd"))) (td as HTMLElement).style.paddingBlock = "0.3ex";
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
export function measureMath(latex: string, size: number, display: boolean): { width: number; height: number; depth: number } {
  const key = `${display ? "D" : "I"}${size}|${latex}`;
  let s = sizes.get(key);
  if (!s) {
    if (!measurer) {
      measurer = h("div", { "aria-hidden": "true", style: "position:absolute;left:-10000px;top:0;visibility:hidden;white-space:nowrap;line-height:normal" });
      document.body.append(measurer);
    }
    const m = styled(mathTemplate(latex, display).cloneNode(true) as Element, size * PT);
    const mark = h("span", { style: "display:inline-block;width:0;height:0;vertical-align:baseline" });
    measurer.replaceChildren(m, mark);
    const r = m.getBoundingClientRect(), b = mark.getBoundingClientRect().bottom;
    s = { width: r.width / PT, height: r.height / PT, depth: Math.max(0, r.bottom - b) / PT };
    if (!(s.width > 0)) s.width = size * 0.6;
    sizes.set(key, s);
  }
  return s;
}

export function renderMath(latex: string, sizePx: number, display: boolean): HTMLElement {
  return styled(mathTemplate(latex, display).cloneNode(true) as Element, sizePx);
}

/** A small rendering for galleries (placeholders as dotted boxes). */
function preview(latex: string, px = 16, display = false): HTMLElement {
  return renderMath(latex.replace(/#\?/g, "⬚"), px, display);
}

// ----- editing in place -------------------------------------------------------------------------

interface Session {
  editor: Editor;
  pos: Pos;
  mf: Field;
  original: string;
  /** An undo step exists for this edit (the insertion, or the first change). */
  recorded: boolean;
  /** Inserted just now, and nothing typed yet: leaving it empty undoes the insertion. */
  isNew: boolean;
  typed: boolean;
}

export function mathFeature(api: AppApi) {
  const c = api.commands;
  const ed = () => api.editor();
  let session: Session | null = null;
  let closing = false;


  const place = (s: Session) => {
    const box = s.editor.objectBox(s.pos);
    if (!box) return;
    Object.assign(s.mf.style, {
      left: `${box.x - 3}px`, top: `${box.y}px`, minWidth: `${Math.max(box.w + 6, 40)}px`, minHeight: `${box.h}px`, fontSize: `${box.size * PT}px`,
    });
  };

  const attach = (e: Editor) => {
    e.useMath(measureMath, renderMath);
    e.onEditMath = (pos, enter) => void open(e, pos, enter);
  };

  const open = async (e: Editor, pos: Pos, enter: "start" | "end" | "click", inserted = false, isNew = false) => {
    if (session) close(null);
    const m = await loadMathLive();
    const o = e.mathAt(pos);
    if (!o) return;
    const mf = new m.MathfieldElement();
    mf.classList.add("math-edit");
    const s: Session = { editor: e, pos, mf, original: o.latex ?? "", recorded: inserted, isNew, typed: false };
    session = s;
    // MathLive takes some settings only once the field is in the page.
    e.view.container.append(mf);
    mf.mathVirtualKeyboardPolicy = "manual";
    mf.smartFence = true;
    mf.menuItems = [];
    mf.placeholder = "\\text{Geben Sie hier eine Formel ein.}";
    mf.value = o.latex ?? "";
    place(s);
    mf.addEventListener("input", () => {
      s.typed = true;
      if (session !== s || !mf.value.trim()) return;
      s.editor.setMathLatex(s.pos, mf.value, !s.recorded);
      s.recorded = true;
      place(s);
    });
    mf.addEventListener("move-out", (ev) => {
      ev.preventDefault();
      const dir = (ev as CustomEvent<{ direction: string }>).detail.direction;
      close(dir === "backward" || dir === "upward" ? "before" : "after");
    });
    mf.addEventListener("keydown", (ev) => {
      ev.stopPropagation();
      if (ev.key === "Escape" || (ev.key === "Enter" && !ev.shiftKey)) { ev.preventDefault(); close("after"); }
      else if (ev.key === "s" && ev.ctrlKey) { ev.preventDefault(); close(null); void c.run("file.save"); }
    });
    mf.addEventListener("focusout", () => setTimeout(() => {
      if (session === s && !closing && document.activeElement !== mf) close(null);
    }, 0));
    mf.focus();
    if (enter === "start") mf.executeCommand("moveToMathfieldStart");
    else mf.executeCommand("moveToMathfieldEnd");
    api.refresh();
    api.selectTab("equation");
  };

  /** Ends editing; the caret goes before or after the equation, or stays where the user clicked. */
  const close = (caret: "before" | "after" | null) => {
    const s = session;
    if (!s) return;
    closing = true;
    session = null;
    const latex = s.mf.value.trim();
    const e = s.editor;
    if (!latex) {
      if (s.isNew && !s.typed) e.undo();
      else e.setMathLatex(s.pos, "", !s.recorded);
      if (caret) e.caretAt(s.pos);
    } else {
      if (latex !== s.original) e.setMathLatex(s.pos, latex, !s.recorded);
      if (caret) e.caretAt(caret === "after" ? { p: s.pos.p, o: s.pos.o + 1 } : s.pos);
    }
    s.mf.remove();
    closing = false;
    api.refresh();
    if (caret) api.selectTab("insert");
  };

  const insert = async (latex = "") => {
    const e = ed();
    if (!e) return;
    if (session) {
      session.mf.insert(latex, { selectionMode: "placeholder", focus: true });
      return;
    }
    const at = e.insertMath(latex);
    if (at) await open(e, at, "end", true, !latex);
  };

  const editing = () => !!session;
  const focus = () => session?.mf.focus();
  const inField = (fn: (mf: Field) => void) => () => { if (session) { fn(session.mf); session.mf.focus(); } };

  c.implement("insert.equation", { run: () => void insert(), enabled: () => !!ed() });
  c.implement("math.new", { run: () => void insert(), enabled: () => !!ed() });
  c.implement("math.gallery", { run: () => {}, enabled: () => !!ed() });
  c.implement("math.normalText", {
    run: inField((mf) => mf.executeCommand(["switchMode", mf.mode === "text" ? "math" : "text"])),
    enabled: editing, checked: () => session?.mf.mode === "text",
  });
  for (const id of ["math.symbols", ...Object.keys(STRUCTURES)]) c.implement(id, { run: () => {}, enabled: editing });

  // ----- galleries ----------------------------------------------------------------------------
  const tile = (latex: string, onPick: () => void, label?: string) =>
    h("button", { class: "mtile", "data-tip": label ?? "", onclick: () => onPick() }, preview(latex, 15));

  const structurePanel = (id: string) => (close2: () => void) => {
    const panel = h("div", { class: "rpanel mathpanel" });
    for (const g of STRUCTURES[id]) {
      if (g.heading) panel.append(h("div", { class: "title" }, g.heading));
      panel.append(h("div", { class: "mgrid" }, ...g.items.map((x) => tile(x, () => { close2(); void insert(x); }))));
    }
    return panel;
  };

  const symbolPanel = (close2: () => void) => {
    const panel = h("div", { class: "rpanel mathpanel" });
    for (const g of SYMBOLS) {
      panel.append(h("div", { class: "title" }, g.heading));
      panel.append(h("div", { class: "msyms" }, ...g.items.map(([ch, tex]) =>
        h("button", { class: "msym", "data-tip": tex, onclick: () => { close2(); void insert(tex); } }, ch))));
    }
    return panel;
  };

  const galleryPanel = (close2: () => void) => {
    const panel = h("div", { class: "rpanel mathpanel mathgallery" }, h("div", { class: "title" }, "Integriert"));
    for (const b of BUILT_IN) {
      panel.append(h("button", { class: "mbuilt", onclick: () => { close2(); void insert(b.latex); } },
        h("div", { class: "mname" }, b.name), h("div", { class: "mprev" }, preview(b.latex, 16, true))));
    }
    panel.append(h("hr"), h("button", { class: "link", onclick: () => { close2(); void insert(); } }, "Neue Formel einfügen"));
    return panel;
  };

  const dd = (id: string, key: string, popup: (close: () => void) => HTMLElement, size: "large" | "medium" = "large"): Control =>
    ({ type: "dropdown", cmd: id, key, size, menu: [], popup });
  const tab: Tab = {
    id: "equation", label: "Formel", key: "JI", contextual: true, shrink: ["structures", "symbols", "tools"],
    groups: [
      { id: "tools", label: "Tools", key: "ZT", icon: "equation",
        controls: [dd("math.gallery", "E", galleryPanel), { type: "button", cmd: "math.normalText", key: "X", size: "large" }] },
      { id: "symbols", label: "Symbole", key: "ZS", icon: "math-symbols", controls: [dd("math.symbols", "S", symbolPanel)] },
      { id: "structures", label: "Strukturen", key: "ZU", icon: "math-fraction",
        controls: [["math.fraction", "F"], ["math.script", "S1"], ["math.radical", "R"], ["math.integral", "I"], ["math.nary", "L"], ["math.bracket", "K"],
          ["math.function", "U"], ["math.accent", "A"], ["math.limit", "G"], ["math.operator", "O"], ["math.matrix", "M"]].map(([id, k]) => dd(id, k, structurePanel(id))) },
    ],
  };
  const insertMenu: MenuEntry[] = [
    ...BUILT_IN.map((b) => ({ label: b.name, run: () => void insert(b.latex), preview: h("div", { class: "mprev" }, preview(b.latex, 15, true)) })),
    "-", { cmd: "math.new" },
  ];
  return { attach, tab, insertMenu, editing, focus, close: () => close(null) };
}
