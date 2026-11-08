// Windows and views of a workbook (D-093), as recorded in Excel 365 (probe_excel_windows.py):
// - Teilen: <pane xSplit ySplit topLeftCell activePane/> without state (split is the default); xSplit/ySplit
//   are 1/20 pt from the window's left/top edge with the headers, at the window's zoom; the sheetView's
//   topLeftCell is the top-left pane's first cell, the pane's topLeftCell the bottom-right pane's (A1 left
//   out); one <selection pane=…> per pane other than the top-left, the inactive ones on their pane's first cell.
// - Seitenlayout: view="pageLayout" with zoomScale = zoomScalePageLayoutView and zoomScaleNormal; showRuler="0".
// - Neues Fenster: one workbookView per window (visibility="hidden" for Ausblenden, activeTab), and every
//   sheet one sheetView per window (workbookViewId), tabSelected on each window's active sheet.
// - Benutzerdefinierte Ansichten: customWorkbookViews (name, guid, …) with one customSheetView per sheet.

import { cellName, parseRange, type Rect, type View, type Workbook } from "@raabe/sheet-engine";
import { attrs, children, esc, type El } from "./xml";

const parseCell = (a: string) => {
  const r = parseRange(a);
  return r ? { row: r.r0, col: r.c0 } : undefined;
};

type Pane = NonNullable<View["split"]>["pane"];

/** A sheetView's split and Seitenlayout settings into the view (after the frozen panes were read). */
export function readViewExtras(a: Record<string, string>, pa: Record<string, string> | undefined, v: View, sels: Record<string, string>[] = []): void {
  if (pa && (!pa.state || pa.state === "split") && (Number(pa.xSplit ?? 0) > 0 || Number(pa.ySplit ?? 0) > 0)) {
    const br = parseCell(pa.topLeftCell ?? "A1") ?? { row: 0, col: 0 };
    const x = Number(pa.xSplit ?? 0), y = Number(pa.ySplit ?? 0);
    v.split = { x, y, top: v.top, left: v.left, pane: (pa.activePane ?? "topLeft") as Pane };
    const others = sels.filter((q) => (q.pane ?? "topLeft") !== v.split!.pane && q.activeCell);
    if (others.length) v.split.sels = Object.fromEntries(others.map((q) => [q.pane ?? "topLeft", { ac: q.activeCell, sqref: q.sqref ?? q.activeCell }]));
    // Without a horizontal bar the panes share their rows; without a vertical one their columns.
    v.top = y ? br.row : v.top;
    v.left = x ? br.col : v.left;
  }
  if (a.view === "pageLayout") {
    v.layout = true;
    v.zoom = Number(a.zoomScalePageLayoutView ?? a.zoomScale ?? 100) || 100;
    v.normalZoom = Number(a.zoomScaleNormal ?? 100) || 100;
  }
  if (a.showRuler === "0" || a.showRuler === "false") v.noRuler = true;
}

/** The panes a split has, in the file's order. */
export function splitPanes(s: NonNullable<View["split"]>): NonNullable<Pane>[] {
  if (s.x && s.y) return ["topLeft", "topRight", "bottomLeft", "bottomRight"];
  return s.x ? ["topLeft", "topRight"] : ["topLeft", "bottomLeft"];
}

/** A pane's first cell. */
export function paneCell(v: View, p: NonNullable<Pane>): { row: number; col: number } {
  const s = v.split!;
  const top = p.startsWith("top") && s.y ? s.top : v.top;
  const left = p.endsWith("Left") && s.x ? s.left : v.left;
  return { row: top, col: left };
}

/** The <pane> and <selection> elements of a split view as Excel writes them. */
export function splitXml(v: View): { pane: string; sels: string; active: string } {
  const s = v.split!;
  const panes = splitPanes(s);
  const active = s.pane && panes.includes(s.pane) ? s.pane : panes[panes.length - 1];
  const br = cellName(v.top, v.left);
  const pane = `<pane${s.x ? ` xSplit="${Math.round(s.x)}"` : ""}${s.y ? ` ySplit="${Math.round(s.y)}"` : ""}${br !== "A1" ? ` topLeftCell="${br}"` : ""}${active !== "topLeft" ? ` activePane="${active}"` : ""}/>`;
  const sqref = v.selection.map((r: Rect) => (r.r0 === r.r1 && r.c0 === r.c1 ? cellName(r.r0, r.c0) : `${cellName(r.r0, r.c0)}:${cellName(r.r1, r.c1)}`)).join(" ") || cellName(v.active.row, v.active.col);
  const sels = panes.map((p) => {
    if (p === active) {
      const ac = cellName(v.active.row, v.active.col);
      if (p === "topLeft" && ac === "A1" && sqref === "A1") return "";
      return `<selection${p !== "topLeft" ? ` pane="${p}"` : ""} activeCell="${ac}" sqref="${sqref}"/>`;
    }
    if (p === "topLeft") return "";
    const kept = s.sels?.[p];
    if (kept) return `<selection pane="${p}" activeCell="${kept.ac}" sqref="${kept.sqref}"/>`;
    const c = paneCell(v, p);
    const a = cellName(c.row, c.col);
    return `<selection pane="${p}" activeCell="${a}" sqref="${a}"/>`;
  }).join("");
  return { pane, sels, active };
}

/** The sheetView attributes of the Seitenlayout view (null removes one). */
export function layoutAttrs(v: View, oldTag: string): Record<string, string | null> {
  const out: Record<string, string | null> = { showRuler: v.noRuler ? "0" : null };
  if (v.layout) {
    out.view = "pageLayout";
    out.zoomScale = v.zoom !== 100 ? String(v.zoom) : null;
    out.zoomScaleNormal = String(v.normalZoom ?? 100);
    out.zoomScalePageLayoutView = String(v.zoom);
  } else if (/\bview="pageLayout"/.test(oldTag)) out.view = null;
  return out;
}

/** Windows (bookViews) and Benutzerdefinierte Ansichten (customWorkbookViews) of the workbook. */
export function readBookViews(xml: string, book: El | undefined, custom: El | undefined, wb: Workbook, sheetOf: (fileIndex: number) => number): void {
  const views = book ? children(xml, book.open, book.close).filter((k) => k.name === "workbookView") : [];
  if (views.length > 1 || views.some((k) => attrs(xml, k).visibility === "hidden")) {
    wb.windows = views.map((k) => {
      const a = attrs(xml, k);
      return { active: sheetOf(Number(a.activeTab ?? 0)), ...(a.visibility === "hidden" ? { hidden: true } : {}), attrs: a };
    });
  }
  if (custom) {
    wb.customViews = children(xml, custom.open, custom.close).map((k) => {
      const a = attrs(xml, k);
      return { name: a.name ?? "", guid: a.guid ?? "", print: a.includePrintSettings !== "0", rowCol: a.includeHiddenRowCol !== "0", attrs: a };
    });
  }
}

/** <bookViews> for the model's windows; `active` maps a sheet to its file index. */
export function bookViewsXml(wb: Workbook, fileIndex: (sheet: number) => number, first: string | undefined): string | null {
  const wins = wb.windows;
  if (!wins?.length) return null;
  const firstAttrs = first ? Object.fromEntries([...first.matchAll(/([\w:]+)="([^"]*)"/g)].map((m) => [m[1], m[2]])) : {};
  return `<bookViews>${wins.map((w, i) => {
    const base: Record<string, string> = { ...(w.attrs ?? (i ? { ...firstAttrs } : firstAttrs)) };
    delete base.visibility; delete base.activeTab; delete base["xr2:uid"];
    const tab = fileIndex(i === 0 ? wb.active : w.active);
    const own = w.attrs?.["xr2:uid"];
    const parts = [...(w.hidden ? [["visibility", "hidden"]] : []), ...Object.entries(base).filter(([k]) => k !== "firstSheet" || Number(base.firstSheet) <= tab),
      ...(tab ? [["activeTab", String(tab)]] : []), ...(own ? [["xr2:uid", own]] : [])];
    return `<workbookView${parts.map(([k, v]) => ` ${k}="${esc(v)}"`).join("")}/>`;
  }).join("")}</bookViews>`;
}

/** <customWorkbookViews> for the model's Benutzerdefinierte Ansichten (null: none). */
export function customBookViewsXml(wb: Workbook): string | null {
  if (!wb.customViews?.length) return null;
  return `<customWorkbookViews>${wb.customViews.map((c) => {
    const a: Record<string, string> = { ...c.attrs, name: c.name, guid: c.guid };
    if (c.print) delete a.includePrintSettings; else a.includePrintSettings = "0";
    if (c.rowCol) delete a.includeHiddenRowCol; else a.includeHiddenRowCol = "0";
    const order = ["name", "guid", "includePrintSettings", "includeHiddenRowCol"];
    const keys = [...order.filter((k) => k in a), ...Object.keys(a).filter((k) => !order.includes(k))];
    return `<customWorkbookView${keys.map((k) => ` ${k}="${esc(a[k])}"`).join("")}/>`;
  }).join("")}</customWorkbookViews>`;
}
