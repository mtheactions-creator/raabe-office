// Start > Rahmenlinie > Rahmenlinien zeichnen (D-066): Rahmenlinie zeichnen (a click on a cell's
// edge draws that line, a drag the outline of the dragged cells), Rahmenraster zeichnen (a drag
// draws every line of the cells), Rahmenlinie löschen (the eraser), Linienfarbe and Linienart
// (choosing one switches drawing on, as in Excel). Esc or the same command again ends it; the
// pointer is a pencil or an eraser of our own.

import { colorGrid, h, type MenuEntry } from "@raabe/shell";
import type { BorderSide, Borders, Rect } from "@raabe/sheet-engine";
import type { CalcApp } from "./app";

export type DrawMode = "draw" | "grid" | "erase";

/** Excel's line styles in the order of its Linienart menu, drawn as our own samples. */
export const LINE_STYLES: { style: string; dash: string; width: number; double?: boolean }[] = [
  { style: "thin", dash: "", width: 1 }, { style: "hair", dash: "1 1", width: 1 }, { style: "dotted", dash: "1 2", width: 1 },
  { style: "dashDotDot", dash: "6 2 2 2 2 2", width: 1 }, { style: "dashDot", dash: "6 2 2 2", width: 1 }, { style: "dashed", dash: "3 1", width: 1 },
  { style: "mediumDashDotDot", dash: "6 2 2 2 2 2", width: 2 }, { style: "slantDashDot", dash: "8 2 3 2", width: 2 }, { style: "mediumDashDot", dash: "8 2 3 2", width: 2 },
  { style: "mediumDashed", dash: "6 2", width: 2 }, { style: "medium", dash: "", width: 2 }, { style: "thick", dash: "", width: 3 },
  { style: "double", dash: "", width: 1, double: true },
];

const PENCIL = `url("data:image/svg+xml,${encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20"><path d="M2 18l2-6 10-10 4 4-10 10z" fill="#fff" stroke="#000"/><path d="M2 18l2-6 4 4z" fill="#000"/></svg>')}") 1 19, crosshair`;
const ERASER = `url("data:image/svg+xml,${encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20"><path d="M3 15l8-8 6 6-5 5H6z" fill="#fff" stroke="#000"/><path d="M3 15l4-4 6 6-1 1H6z" fill="#e88"/></svg>')}") 3 17, crosshair`;

export function borderDrawFeature(app: CalcApp) {
  let mode: DrawMode | null = null;
  let color: string | undefined;
  let style = "thin";
  let preview: HTMLElement | null = null;

  const side = (): BorderSide => ({ style, ...(color ? { color } : {}) });

  function start(m: DrawMode): void {
    // The grid exists once the app is built: listen from the first use on.
    if (!attached) { attached = true; app.grid.el.addEventListener("pointerdown", onDown, true); }
    if (app.grid.isEditing) app.grid.commitEdit("none");
    mode = mode === m ? null : m;
    app.grid.el.style.cursor = mode ? (mode === "erase" ? ERASER : PENCIL) : "";
    app.commands.refresh();
  }

  function stop(): void {
    if (!mode) return;
    mode = null;
    app.grid.el.style.cursor = "";
    preview?.remove();
    preview = null;
    app.commands.refresh();
  }

  /** The cell edge nearest to the point (grid coordinates), as Excel's pencil snaps to it. */
  function edgeAt(x: number, y: number): { row: number; col: number; edge: keyof Borders } | null {
    const hit = app.grid.hitCell(x, y);
    const r = app.grid.rectOnCanvas({ r0: hit.row, c0: hit.col, r1: hit.row, c1: hit.col });
    if (!r) return null;
    const d = { top: y - r.y, bottom: r.y + r.h - y, left: x - r.x, right: r.x + r.w - x };
    const edge = (Object.keys(d) as (keyof typeof d)[]).reduce((a, b) => (d[a] <= d[b] ? a : b));
    return { row: hit.row, col: hit.col, edge };
  }

  /** Sets (or with null clears) the given sides of the cells; a cleared line goes on both cells. */
  function apply(label: string, edits: { row: number; col: number; edge: "top" | "bottom" | "left" | "right"; to: BorderSide | null }[]): void {
    if (!edits.length) return;
    const sheet = app.sheet;
    const opposite = { top: "bottom", bottom: "top", left: "right", right: "left" } as const;
    const all = [...edits];
    for (const e of edits) if (!e.to) {
      const r = e.edge === "top" ? e.row - 1 : e.edge === "bottom" ? e.row + 1 : e.row, c = e.edge === "left" ? e.col - 1 : e.edge === "right" ? e.col + 1 : e.col;
      if (r >= 0 && c >= 0) all.push({ row: r, col: c, edge: opposite[e.edge], to: null });
    }
    app.change(label, () => {
      for (const e of all) {
        app.history.cells(sheet, { r0: e.row, c0: e.col, r1: e.row, c1: e.col });
        const c = sheet.get(e.row, e.col) ?? { v: null };
        const s = app.wb.styles.derive(c.s, (st) => {
          if (e.to) st.borders[e.edge] = { ...e.to };
          else delete st.borders[e.edge];
        });
        sheet.set(e.row, e.col, { ...c, s });
      }
    });
  }

  function rangeEdits(r: Rect, m: DrawMode): Parameters<typeof apply>[1] {
    const out: Parameters<typeof apply>[1] = [];
    const to = m === "erase" ? null : side();
    for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) {
      const inner = m !== "draw";
      if (row === r.r0 || inner) out.push({ row, col, edge: "top", to });
      if (row === r.r1 || inner) out.push({ row, col, edge: "bottom", to });
      if (col === r.c0 || inner) out.push({ row, col, edge: "left", to });
      if (col === r.c1 || inner) out.push({ row, col, edge: "right", to });
    }
    return out;
  }

  const onDown = (e: PointerEvent) => {
    if (!mode || e.button !== 0 || !(e.target as HTMLElement).closest(".grid-canvas")) return;
    e.preventDefault();
    e.stopPropagation();
    const box = app.grid.el.getBoundingClientRect();
    const x0 = e.clientX - box.left, y0 = e.clientY - box.top;
    const a = app.grid.hitCell(x0, y0);
    let b = a;
    let dragged = false;
    const move = (ev: PointerEvent) => {
      const x = ev.clientX - box.left, y = ev.clientY - box.top;
      if (Math.hypot(x - x0, y - y0) > 4) dragged = true;
      b = app.grid.hitCell(x, y);
      if (!dragged) return;
      const r = app.grid.rectOnCanvas({ r0: Math.min(a.row, b.row), c0: Math.min(a.col, b.col), r1: Math.max(a.row, b.row), c1: Math.max(a.col, b.col) });
      if (!r) return;
      preview ??= app.grid.el.appendChild(h("div", { style: "position:absolute;pointer-events:none;border:1px dashed #555;z-index:5" }));
      Object.assign(preview.style, { left: `${r.x}px`, top: `${r.y}px`, width: `${r.w}px`, height: `${r.h}px` });
    };
    const up = (ev: PointerEvent) => {
      removeEventListener("pointermove", move, true);
      removeEventListener("pointerup", up, true);
      preview?.remove();
      preview = null;
      if (!mode) return;
      const label = mode === "erase" ? "Rahmenlinie löschen" : "Rahmenlinie zeichnen";
      if (!dragged) {
        const edge = edgeAt(ev.clientX - box.left, ev.clientY - box.top);
        if (!edge) return;
        if (mode === "erase") {
          // The eraser on a cell clears all its lines; on an edge just that one.
          apply(label, rangeEdits({ r0: edge.row, c0: edge.col, r1: edge.row, c1: edge.col }, "erase"));
        } else apply(label, [{ row: edge.row, col: edge.col, edge: edge.edge as "top", to: side() }]);
        return;
      }
      apply(label, rangeEdits({ r0: Math.min(a.row, b.row), c0: Math.min(a.col, b.col), r1: Math.max(a.row, b.row), c1: Math.max(a.col, b.col) }, mode));
    };
    addEventListener("pointermove", move, true);
    addEventListener("pointerup", up, true);
  };
  let attached = false;
  addEventListener("keydown", (e) => { if (mode && e.key === "Escape") { e.preventDefault(); stop(); } }, true);

  /** Linienfarbe: Excel's colour palette; picking one switches drawing on. */
  const colorFlyout = (): HTMLElement => {
    const box = h("div", { class: "palette" });
    box.append(h("button", { type: "button", class: "item", onclick: () => { color = undefined; if (!mode) start("draw"); } }, "Automatisch"));
    box.append(...colorGrid((hex) => { color = hex; if (!mode || mode === "erase") start("draw"); document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" })); }, app.wb.theme));
    return box;
  };

  /** Linienart: samples of the thirteen line styles (and "Ohne Linie" is the eraser in Excel's menu). */
  const styleFlyout = (): HTMLElement => {
    const box = h("div", { class: "menu line-styles" });
    for (const ls of LINE_STYLES) {
      const svg = ls.double
        ? `<svg width="90" height="10"><line x1="2" y1="3.5" x2="88" y2="3.5" stroke="#000"/><line x1="2" y1="6.5" x2="88" y2="6.5" stroke="#000"/></svg>`
        : `<svg width="90" height="10"><line x1="2" y1="5" x2="88" y2="5" stroke="#000" stroke-width="${ls.width}"${ls.dash ? ` stroke-dasharray="${ls.dash}"` : ""}/></svg>`;
      const b = h("button", { type: "button", class: `item${ls.style === style ? " on" : ""}`, "aria-label": ls.style, onclick: () => { style = ls.style; if (!mode || mode === "erase") start("draw"); } });
      b.innerHTML = svg;
      box.append(b);
    }
    return box;
  };

  const menuEntries = (): MenuEntry[] => [
    { label: "Linienfarbe", icon: "font-color", flyout: colorFlyout },
    { label: "Linienart", icon: "line-style", flyout: styleFlyout },
  ];

  return { start, stop, mode: () => mode, menuEntries };
}
