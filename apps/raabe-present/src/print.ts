// Printing and PDF (D-070): Datei > Drucken with PowerPoint's layouts: Ganzseitige Folien (a page of
// the slide's size, as the PDF export), Notizenseiten, Gliederung and Handzettel with 1, 2, 3, 4, 6 or
// 9 slides per A4 page (3 with lines for notes beside them); which slides (Alle Folien, Aktuelle Folie,
// Benutzerdefinierter Bereich "1-3,5"); a preview of each page. Hidden slides are left out, as in
// PowerPoint's default. Drucken uses Windows' print dialog, PDF the platform's PDF printer.

import { h, icon } from "@raabe/shell";
import type { PresentApp } from "./app";
import { renderSlide } from "./render/slide";
import { EMU_PX } from "./render/text";

export type PrintLayout = "full" | "notes" | "outline" | "h1" | "h2" | "h3" | "h4" | "h6" | "h9";

export const LAYOUTS: [PrintLayout, string][] = [
  ["full", "Ganzseitige Folien"], ["notes", "Notizenseiten"], ["outline", "Gliederung"],
  ["h1", "1 Folie"], ["h2", "2 Folien"], ["h3", "3 Folien"], ["h4", "4 Folien horizontal"], ["h6", "6 Folien horizontal"], ["h9", "9 Folien horizontal"],
];

export interface PrintSettings { layout: PrintLayout; which: "all" | "current" | "custom"; range: string; frame: boolean }

const A4 = { w: 8.27, h: 11.69 };

function chosen(app: PresentApp, s: PrintSettings): number[] {
  const deck = app.deck;
  const shown = deck.slides.map((_, i) => i).filter((i) => deck.pptx.doc(deck.slides[i].path).documentElement.getAttribute("show") !== "0");
  if (s.which === "current") return [app.current()];
  if (s.which === "custom") {
    const out: number[] = [];
    for (const part of s.range.split(/[,;]/)) {
      const m = /^\s*(\d+)\s*(?:-\s*(\d+))?\s*$/.exec(part);
      if (!m) continue;
      const a = Number(m[1]), b = Number(m[2] ?? m[1]);
      for (let k = Math.min(a, b); k <= Math.max(a, b); k++) if (k >= 1 && k <= deck.slides.length) out.push(k - 1);
    }
    return out;
  }
  return shown;
}

function slideBox(app: PresentApp, i: number, widthIn: number, frame: boolean): HTMLElement {
  const deck = app.deck;
  const size = deck.pptx.slideSize;
  const w = widthIn * 96, s = w / (size.cx / EMU_PX);
  const box = h("div", { style: `position:relative;overflow:hidden;width:${w}px;height:${(size.cy / EMU_PX) * s}px;${frame ? "outline:1px solid #000" : ""}` });
  const r = renderSlide(deck.view(deck.slides[i].path), size, { editing: false, slideNumber: i + 1, image: (p) => app.image(p), pptx: deck.pptx });
  r.el.style.transform = `scale(${s})`; r.el.style.transformOrigin = "0 0";
  box.append(r.el);
  return box;
}

/** The printed pages of the settings, each a div of its paper size in CSS inches. */
export function printPages(app: PresentApp, s: PrintSettings): { pages: HTMLElement[]; paper: { w: number; h: number } } {
  const deck = app.deck;
  const size = deck.pptx.slideSize;
  const list = chosen(app, s);
  if (s.layout === "full") {
    const paper = { w: size.cx / 914400, h: size.cy / 914400 };
    return { paper, pages: list.map((i) => h("div", { class: "print-page", style: `width:${paper.w}in;height:${paper.h}in;position:relative;overflow:hidden` }, slideBox(app, i, paper.w, s.frame))) };
  }
  const page = () => h("div", { class: "print-page", style: `width:${A4.w}in;height:${A4.h}in;position:relative;overflow:hidden;box-sizing:border-box;padding:0.5in 0.6in;font:11pt Calibri, Carlito, sans-serif;background:#fff` });
  if (s.layout === "notes") {
    return { paper: A4, pages: list.map((i) => {
      const pg = page();
      pg.append(h("div", { style: "display:flex;justify-content:center;margin-top:0.6in" }, slideBox(app, i, 6, true)),
        h("div", { style: "margin:0.4in 0.3in 0;white-space:pre-wrap;font-size:12pt" }, deck.notesText(deck.slides[i].path)),
        h("div", { style: "position:absolute;right:0.6in;bottom:0.4in;font-size:9pt" }, String(i + 1)));
      return pg;
    }) };
  }
  if (s.layout === "outline") {
    const pg = page();
    for (const i of list) {
      const v = deck.view(deck.slides[i].path);
      pg.append(h("div", { style: "margin-top:8pt;font-weight:bold" }, `${i + 1}  ${v.shapes.filter((x) => x.ph && (x.ph.type === "title" || x.ph.type === "ctrTitle")).map((x) => x.text?.paras.map((p) => p.runs.map((r) => r.text).join("")).join(" ") ?? "").join(" ")}`));
      for (const sh of v.shapes) if (sh.ph && !["title", "ctrTitle", "dt", "ftr", "sldNum"].includes(sh.ph.type) && sh.text && !sh.text.empty)
        for (const p of sh.text.paras) pg.append(h("div", { style: `margin-left:${0.3 + p.lvl * 0.3}in` }, `• ${p.runs.map((r) => r.text).join("")}`));
    }
    return { paper: A4, pages: [pg] };
  }
  const per = Number(s.layout.slice(1));
  const cols = per >= 4 ? (per === 4 ? 2 : per === 6 ? 2 : 3) : 1;
  const rows = Math.ceil(per / cols);
  const cellW = per === 3 ? 3.2 : (A4.w - 1.2 - (cols - 1) * 0.3) / cols;
  const cellH = (A4.h - 1.6) / rows;
  const ratio = size.cy / size.cx;
  const sw = Math.min(cellW, (cellH - 0.2) / ratio);
  const pages: HTMLElement[] = [];
  for (let k = 0; k < list.length; k += per) {
    const pg = page();
    const grid = h("div", { style: `display:grid;grid-template-columns:repeat(${cols}, ${per === 3 ? "auto 1fr" : "1fr"});gap:0.25in 0.3in;align-items:center;justify-items:center;height:100%` });
    for (const i of list.slice(k, k + per)) {
      grid.append(slideBox(app, i, sw, true));
      // Handzettel mit 3 Folien: lines for notes beside each slide (PowerPoint).
      if (per === 3) grid.append(h("div", { style: "width:100%;align-self:stretch;display:flex;flex-direction:column;justify-content:space-evenly" }, ...Array.from({ length: 6 }, () => h("div", { style: "border-bottom:1px solid #777;height:1px" }))));
    }
    pg.append(grid, h("div", { style: "position:absolute;right:0.6in;bottom:0.35in;font-size:9pt" }, String(pages.length + 1)));
    pages.push(pg);
  }
  return { paper: A4, pages };
}

async function withPrintRoot<T>(pages: HTMLElement[], paper: { w: number; h: number }, run: () => Promise<T>): Promise<T> {
  const root = document.getElementById("print-root") ?? document.body.appendChild(h("div", { id: "print-root" }));
  root.replaceChildren(...pages);
  const style = h("style", null, `@page { size: ${paper.w}in ${paper.h}in; margin: 0 }`);
  document.head.append(style);
  document.body.classList.add("printing");
  await Promise.all(Array.from(root.querySelectorAll("img")).map((img) => (img as HTMLImageElement).decode().catch(() => {})));
  try {
    return await run();
  } finally {
    document.body.classList.remove("printing");
    root.replaceChildren();
    style.remove();
  }
}

export async function printSlides(app: PresentApp, path: string, s: PrintSettings = { layout: "full", which: "all", range: "", frame: false }): Promise<void> {
  const { pages, paper } = printPages(app, s);
  await withPrintRoot(pages, paper, () => app.platform.printToPdf(path, { width: paper.w, height: paper.h }));
}

/** Datei > Drucken: settings at the left, the preview of each page at the right. */
export function printPage(app: PresentApp, s: PrintSettings): HTMLElement {
  let i = 0;
  const preview = h("div", { class: "print-preview" });
  const label = h("span");
  const show = () => {
    const { pages, paper } = printPages(app, s);
    i = Math.min(i, Math.max(0, pages.length - 1));
    preview.replaceChildren();
    const pg = pages[i];
    if (!pg) { preview.append(h("p", { class: "muted" }, "Es gibt nichts zu drucken.")); label.textContent = "0 von 0"; return; }
    const scale = Math.min(0.75, (innerHeight - 220) / (paper.h * 96), 620 / (paper.w * 96));
    pg.style.transform = `scale(${scale})`; pg.style.transformOrigin = "top left";
    preview.append(h("div", { style: `width:${paper.w * 96 * scale}px;height:${paper.h * 96 * scale}px;overflow:hidden;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.25)` }, pg));
    label.textContent = `${i + 1} von ${pages.length}`;
  };
  const layout = h("select", { onchange: (e: Event) => { s.layout = (e.target as HTMLSelectElement).value as PrintLayout; show(); } }, ...LAYOUTS.map(([v, l]) => h("option", { value: v, selected: v === s.layout || undefined }, l))) as HTMLSelectElement;
  const which = h("select", { onchange: (e: Event) => { s.which = (e.target as HTMLSelectElement).value as PrintSettings["which"]; show(); } },
    h("option", { value: "all" }, "Alle Folien drucken"), h("option", { value: "current" }, "Aktuelle Folie drucken"), h("option", { value: "custom" }, "Benutzerdefinierter Bereich")) as HTMLSelectElement;
  which.value = s.which;
  const range = h("input", { type: "text", value: s.range, placeholder: "Folien:", style: "width:160px", oninput: (e: Event) => { s.range = (e.target as HTMLInputElement).value; if (s.which !== "custom") { s.which = "custom"; which.value = "custom"; } show(); } }) as HTMLInputElement;
  const frame = h("input", { type: "checkbox", onchange: (e: Event) => { s.frame = (e.target as HTMLInputElement).checked; show(); } }) as HTMLInputElement;
  frame.checked = s.frame;
  const doPrint = async () => {
    const { pages, paper } = printPages(app, s);
    await withPrintRoot(pages, paper, () => new Promise<void>((resolve) => { addEventListener("afterprint", () => resolve(), { once: true }); print(); }));
  };
  const toPdf = async () => {
    const path = await app.platform.saveFileDialog(`${document.title.replace(/ - RaabePresent$/, "")}.pdf`, "pdf");
    if (path) await printSlides(app, path, s);
  };
  show();
  return h("div", { class: "print-layout" },
    h("div", { class: "print-settings" }, h("h1", null, "Drucken"),
      h("button", { class: "bigbtn", onclick: () => void doPrint() }, icon("print", "l"), "Drucken"),
      h("button", { class: "bigbtn", onclick: () => void toPdf() }, icon("export", "l"), "Als PDF speichern"),
      h("h2", null, "Einstellungen"), which, range, layout, h("label", { class: "dlg-check" }, frame, " Folien mit Rahmen")),
    h("div", { class: "print-right" }, preview, h("div", { class: "print-nav" },
      h("button", { onclick: () => { i = Math.max(0, i - 1); show(); } }, "◀"), label, h("button", { onclick: () => { i++; show(); } }, "▶"))));
}
