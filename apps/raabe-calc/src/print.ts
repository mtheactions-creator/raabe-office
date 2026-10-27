// Printing and PDF as Excel lays sheets out on paper: the used range (or the Druckbereich),
// cut into pages first down, then across, at the page's printable size with Excel's default
// margins ("Normal": 1,9 cm top and bottom, 1,8 cm left and right), Skalierung or An Format
// anpassen, optional gridlines and row/column headings.

import { MAX_COLS, cellName, colName, type PageSetup, type Rect, type Sheet } from "@raabe/sheet-engine";
import { h, icon, showDialog, type MenuEntry } from "@raabe/shell";
import type { CalcApp } from "./app";
import { shown } from "./engine";
import { DEFAULT_COL_CHARS, charsToPx96, ptToPx } from "./grid/metrics";

/** A column's printed width in px at 96 dpi (the standard width rounded up to 8 px). */
function colPx96(sheet: Sheet, c: number): number {
  const w = sheet.cols.get(c)?.width;
  if (sheet.cols.get(c)?.hidden) return 0;
  if (w !== undefined) return charsToPx96(w);
  return Math.ceil((Math.floor((sheet.defaultColWidth ?? DEFAULT_COL_CHARS) + 1e-6) * 7 + 5) / 8) * 8;
}

const PAPERS: { code: number; name: string; w: number; h: number }[] = [
  { code: 9, name: "A4", w: 8.27, h: 11.69 },
  { code: 8, name: "A3", w: 11.69, h: 16.54 },
  { code: 11, name: "A5", w: 5.83, h: 8.27 },
  { code: 1, name: "Letter", w: 8.5, h: 11 },
  { code: 5, name: "Legal", w: 8.5, h: 14 },
];

const MARGINS: Record<string, NonNullable<PageSetup["margins"]>> = {
  "Normale Seitenränder": { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  "Breite Seitenränder": { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
  "Schmale Seitenränder": { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
};

const cm = (inch: number) => `${(inch * 2.54).toFixed(1).replace(".", ",")} cm`;

export function printFeature(app: CalcApp) {
  const page = (): PageSetup => app.sheet.page ?? {};
  const setPage = (label: string, fn: (p: PageSetup) => void) => app.change(label, () => {
    app.history.whole(app.sheet);
    const p = { ...(app.sheet.page ?? {}) };
    fn(p);
    app.sheet.page = p;
  });

  function paper(p: PageSetup) {
    const sz = PAPERS.find((x) => x.code === (p.paperSize ?? 9)) ?? PAPERS[0];
    return p.orientation === "landscape" ? { ...sz, w: sz.h, h: sz.w } : sz;
  }

  /** The pages of a sheet: row and column ranges, Excel's order (down, then across). */
  function paginate(sheet: Sheet): { rows: [number, number]; cols: [number, number] }[] {
    const p = sheet.page ?? {};
    const area: Rect | null = p.printArea ?? sheet.usedRange();
    if (!area) return [];
    const sz = paper(p);
    const m = p.margins ?? MARGINS["Normale Seitenränder"];
    const pw = (sz.w - m.left - m.right) * 96, ph = (sz.h - m.top - m.bottom) * 96;
    const colW = (c: number) => colPx96(sheet, c);
    const rowH = (r: number) => (sheet.rows.get(r)?.hidden ? 0 : ptToPx(sheet.rows.get(r)?.height ?? sheet.defaultRowHeight));
    let scale = (p.scale ?? 100) / 100;
    if (p.fitWidth || p.fitHeight) {
      let tw = 0, th = 0;
      for (let c = area.c0; c <= area.c1; c++) tw += colW(c);
      for (let r = area.r0; r <= area.r1; r++) th += rowH(r);
      const sw = p.fitWidth ? (pw * p.fitWidth) / tw : Infinity, sh = p.fitHeight ? (ph * p.fitHeight) / th : Infinity;
      scale = Math.max(0.1, Math.min(1, sw, sh));
    }
    const cut = (from: number, to: number, size: (i: number) => number, room: number) => {
      const out: [number, number][] = [];
      let start = from, used = 0;
      for (let i = from; i <= to; i++) {
        const s = size(i) * scale;
        if (used + s > room && i > start) { out.push([start, i - 1]); start = i; used = 0; }
        used += s;
      }
      out.push([start, to]);
      return out;
    };
    const colPages = cut(area.c0, area.c1, colW, pw);
    const rowPages = cut(area.r0, area.r1, rowH, ph);
    const pages: { rows: [number, number]; cols: [number, number] }[] = [];
    for (const cols of colPages) for (const rows of rowPages) pages.push({ rows, cols });
    (paginate as unknown as { scale: number }).scale = scale;
    return pages;
  }

  /** One page as HTML at its printed size. */
  function renderPage(sheet: Sheet, pg: { rows: [number, number]; cols: [number, number] }): HTMLElement {
    const p = sheet.page ?? {};
    const sz = paper(p);
    const m = p.margins ?? MARGINS["Normale Seitenränder"];
    const scale = (paginate as unknown as { scale: number }).scale ?? 1;
    const table = h("table", { class: "print-table", style: `transform:scale(${scale});transform-origin:top left` });
    const colgroup = h("colgroup");
    if (p.headings) colgroup.append(h("col", { style: "width:30px" }));
    for (let c = pg.cols[0]; c <= pg.cols[1]; c++) {
      const w = colPx96(sheet, c);
      colgroup.append(h("col", { style: `width:${w}px` }));
    }
    table.append(colgroup);
    if (p.headings) table.append(h("tr", null, h("th"), ...range(pg.cols[0], pg.cols[1]).map((c) => h("th", null, colName(c)))));
    for (let r = pg.rows[0]; r <= pg.rows[1]; r++) {
      if (sheet.rows.get(r)?.hidden) continue;
      const tr = h("tr", { style: `height:${ptToPx(sheet.rows.get(r)?.height ?? sheet.defaultRowHeight)}px` });
      if (p.headings) tr.append(h("th", null, String(r + 1)));
      for (let c = pg.cols[0]; c <= pg.cols[1]; c++) {
        const merge = sheet.mergeAt(r, c);
        if (merge && (merge.r0 !== r || merge.c0 !== c)) continue;
        const cell = sheet.get(r, c);
        const st = app.wb.styles.get(cell?.s ?? sheet.rows.get(r)?.s ?? sheet.cols.get(c)?.s);
        const s = shown(cell, st);
        const css = [
          `font:${st.font.italic ? "italic " : ""}${st.font.bold ? "bold " : ""}${st.font.size}pt "${st.font.name}",Carlito,sans-serif`,
          st.font.color || s.color ? `color:${s.color ?? st.font.color}` : "", st.fill ? `background:${st.fill}` : "",
          `text-align:${st.h && st.h !== "general" ? st.h : s.align}`, `vertical-align:${st.v === "top" ? "top" : st.v === "center" ? "middle" : "bottom"}`,
          st.wrap ? "white-space:pre-wrap" : "white-space:pre", st.font.underline ? "text-decoration:underline" : "",
          ...(["top", "right", "bottom", "left"] as const).map((k) => (st.borders[k] ? `border-${k}:${st.borders[k]!.style === "medium" ? 2 : st.borders[k]!.style === "thick" ? 3 : 1}px ${st.borders[k]!.style === "double" ? "double" : st.borders[k]!.style.includes("dash") ? "dashed" : st.borders[k]!.style.includes("dot") ? "dotted" : "solid"} ${st.borders[k]!.color ?? "#000"}` : "")),
        ].filter(Boolean).join(";");
        const attrs: Record<string, string> = { style: css };
        if (merge) { attrs.rowspan = String(merge.r1 - merge.r0 + 1); attrs.colspan = String(merge.c1 - merge.c0 + 1); }
        tr.append(h("td", attrs, s.text));
      }
      table.append(tr);
    }
    if (p.gridLines) table.classList.add("grid");
    return h("div", { class: "print-page", style: `width:${sz.w}in;height:${sz.h}in;padding:${m.top}in ${m.right}in ${m.bottom}in ${m.left}in` }, table);
  }

  function pages(): HTMLElement[] {
    return paginate(app.sheet).map((pg) => renderPage(app.sheet, pg));
  }

  async function exportPdf(): Promise<void> {
    const path = await app.platform.saveFileDialog(`${document.title.replace(/ - RaabeCalc$/, "")}.pdf`, "pdf");
    if (!path) return;
    const root = document.getElementById("print-root") ?? document.body.appendChild(h("div", { id: "print-root" }));
    root.replaceChildren(...pages());
    const sz = paper(page());
    document.body.classList.add("printing");
    try {
      await app.platform.printToPdf(path, { width: sz.w, height: sz.h });
    } finally {
      document.body.classList.remove("printing");
      root.replaceChildren();
    }
  }

  function printNow(): void {
    const root = document.getElementById("print-root") ?? document.body.appendChild(h("div", { id: "print-root" }));
    root.replaceChildren(...pages());
    const sz = paper(page());
    const style = h("style", null, `@page { size: ${sz.w}in ${sz.h}in; margin: 0 }`);
    document.head.append(style);
    document.body.classList.add("printing");
    addEventListener("afterprint", () => { document.body.classList.remove("printing"); root.replaceChildren(); style.remove(); }, { once: true });
    print();
  }

  /** Datei > Drucken: settings left, preview right. */
  function printPage(): HTMLElement {
    const list = pages();
    let i = 0;
    const preview = h("div", { class: "print-preview" });
    const label = h("span");
    const show = () => {
      preview.replaceChildren();
      const pg = list[i];
      if (!pg) { preview.append(h("p", { class: "muted" }, "Es gibt nichts zu drucken.")); label.textContent = "0 von 0"; return; }
      const sz = paper(page());
      const scale = Math.min(0.7, (innerHeight - 260) / (sz.h * 96), 560 / (sz.w * 96));
      pg.style.transform = `scale(${scale})`;
      pg.style.transformOrigin = "top left";
      preview.append(h("div", { class: "sheet-holder", style: `width:${sz.w * 96 * scale}px;height:${sz.h * 96 * scale}px;overflow:hidden;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.25)` }, pg));
      label.textContent = `${i + 1} von ${list.length}`;
    };
    show();
    const p = page();
    const orient = h("select", { onchange: (e: Event) => { setOrientation((e.target as HTMLSelectElement).value as "portrait" | "landscape"); app.commands.run("file.print"); } },
      h("option", { value: "portrait", selected: p.orientation !== "landscape" || undefined }, "Hochformat"), h("option", { value: "landscape", selected: p.orientation === "landscape" || undefined }, "Querformat"));
    const paperSel = h("select", { onchange: (e: Event) => { setPage("Format", (x) => { x.paperSize = Number((e.target as HTMLSelectElement).value); }); app.commands.run("file.print"); } },
      ...PAPERS.map((x) => h("option", { value: String(x.code), selected: (p.paperSize ?? 9) === x.code || undefined }, `${x.name} (${(x.w * 2.54).toFixed(1).replace(".", ",")} x ${(x.h * 2.54).toFixed(1).replace(".", ",")} cm)`)));
    return h("div", { class: "print-layout" },
      h("div", { class: "print-settings" }, h("h1", null, "Drucken"),
        h("button", { class: "bigbtn", onclick: printNow }, icon("print", "l"), "Drucken"),
        h("h2", null, "Einstellungen"), orient, paperSel,
        h("button", { onclick: () => void pageSetupDialog() }, "Seite einrichten")),
      h("div", { class: "print-right" }, preview,
        h("div", { class: "print-nav" }, h("button", { onclick: () => { if (i > 0) { i--; show(); } } }, "◀"), label, h("button", { onclick: () => { if (i < list.length - 1) { i++; show(); } } }, "▶"))));
  }

  function setOrientation(o: "portrait" | "landscape"): void {
    setPage("Ausrichtung", (p) => { p.orientation = o; });
  }

  async function pageSetupDialog(): Promise<void> {
    const p = page();
    const m = { ...(p.margins ?? MARGINS["Normale Seitenränder"]) };
    const inputs = (["top", "bottom", "left", "right"] as const).map((k) => {
      const input = h("input", { type: "text", value: (m[k] * 2.54).toFixed(1).replace(".", ","), style: "width:60px" }) as HTMLInputElement;
      return { k, input, el: h("label", { class: "field" }, { top: "Oben:", bottom: "Unten:", left: "Links:", right: "Rechts:" }[k], input) };
    });
    const scale = h("input", { type: "text", value: String(p.scale ?? 100), style: "width:60px" }) as HTMLInputElement;
    const grid = h("input", { type: "checkbox", checked: p.gridLines || undefined }) as HTMLInputElement;
    const heads = h("input", { type: "checkbox", checked: p.headings || undefined }) as HTMLInputElement;
    const body = h("div", null, h("div", { class: "fc-label" }, "Seitenränder (cm)"), ...inputs.map((x) => x.el),
      h("label", { class: "field" }, "Verkleinern/Vergrößern (%):", scale),
      h("label", { class: "check" }, grid, "Gitternetzlinien drucken"), h("label", { class: "check" }, heads, "Zeilen- und Spaltenüberschriften drucken"));
    const ok = await showDialog({ title: "Seite einrichten", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    setPage("Seite einrichten", (x) => {
      const nm = { ...m };
      for (const i of inputs) { const v = Number(i.input.value.replace(",", ".")); if (v >= 0) nm[i.k] = v / 2.54; }
      x.margins = nm;
      const s = Number(scale.value);
      if (s >= 10 && s <= 400) { x.scale = s; x.fitWidth = x.fitHeight = undefined; }
      x.gridLines = grid.checked || undefined;
      x.headings = heads.checked || undefined;
    });
  }

  return {
    exportPdf,
    page: printPage,
    orientation: () => page().orientation ?? "portrait",
    setOrientation,
    setPrintArea: (r: Rect | null) => setPage("Druckbereich", (p) => { p.printArea = r ?? undefined; }),
    pageSetupDialog,
    printOption: (k: "gridLines" | "headings") => !!page()[k],
    togglePrintOption: (k: "gridLines" | "headings") => setPage("Blattoptionen", (p) => { p[k] = !p[k] || undefined; }),
    marginsMenu: (): MenuEntry[] => [
      ...Object.entries(MARGINS).map(([name, mm]): MenuEntry => ({
        label: `${name}  (Oben: ${cm(mm.top)}  Unten: ${cm(mm.bottom)}  Links: ${cm(mm.left)}  Rechts: ${cm(mm.right)})`,
        run: () => setPage("Seitenränder", (p) => { p.margins = { ...mm }; }),
      })),
      { label: "Benutzerdefinierte Seitenränder...", run: () => void pageSetupDialog() },
    ],
    sizeMenu: (): MenuEntry[] => PAPERS.map((x) => ({ label: `${x.name}  ${(x.w * 2.54).toFixed(1).replace(".", ",")} cm x ${(x.h * 2.54).toFixed(1).replace(".", ",")} cm`, run: () => setPage("Format", (p) => { p.paperSize = x.code; }) })),
    setFit: (k: "width" | "height", n: number) => setPage("An Format anpassen", (p) => { if (k === "width") p.fitWidth = n || undefined; else p.fitHeight = n || undefined; }),
    scale: () => page().scale ?? 100,
    setScale: (n: number) => setPage("Skalierung", (p) => { p.scale = Math.max(10, Math.min(400, n)); p.fitWidth = p.fitHeight = undefined; }),
  };
}

function range(a: number, b: number): number[] {
  const out: number[] = [];
  for (let i = a; i <= b; i++) out.push(i);
  return out;
}

void MAX_COLS; void cellName;
