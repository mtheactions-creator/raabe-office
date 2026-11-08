// Sparklines in RaabeCalc (D-085): Einfügen > Sparklines and the contextual tab Sparkline of German Excel 365,
// which appears while the active cell holds a sparkline. Every setting belongs to a group (x14:sparklineGroup)
// and is changed for all groups the selection touches, as Excel does; each change is one undo step. Labels and
// keytips as recorded (probe_excel_sparkline_ui.py), the 36 styles of the gallery from Excel's files
// (sparkline-styles.gen.ts), the behaviour of Gruppieren, Gruppierung aufheben and Löschen from Excel's files
// (probe_excel_sparklines.py): new groups come first in the file, Gruppieren takes the settings of the active
// cell's group, Gruppierung aufheben makes each selected sparkline a group of its own.

import { colorGrid, h, messageBox, showDialog, type Control, type MenuEntry, type Tab } from "@raabe/shell";
import { parseRange, rangeName, type Rect, type SparkColor, type SparkColorKey, type SparklineGroup } from "@raabe/sheet-engine";
import { defaultSparkColors } from "@raabe/ooxml";
import type { CalcApp } from "./app";
import { drawSparkline, sparkRanges, sparkRgb, sparkSeries, splitRef } from "./grid/sparkline-draw";
import { SPARKLINE_STYLES } from "./sparkline-styles.gen";

type Kind = SparklineGroup["type"];
type Run = (id: string, fn: () => void | Promise<void>, enabled?: () => boolean, checked?: () => boolean) => void;

/** Stärke: the line weights Office offers (labels as PowerPoint's recorded Stärke menu, D-070). */
export const SPARK_WEIGHTS: [string, number][] = [["1/4 Pt.", 0.25], ["½ Pt.", 0.5], ["3/4 Pt.", 0.75], ["1 Pt.", 1], ["1½ Pt.", 1.5], ["2¼ Pt.", 2.25], ["3 Pt.", 3], ["4½ Pt.", 4.5], ["6 Pt.", 6]];

const inR = (r: Rect, row: number, col: number) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;

/** A sheet name as a reference prefix (quoted when it needs it, as Excel writes it). */
export function sheetRef(name: string): string {
  return /^[A-Za-z_À-ɏ][A-Za-z0-9_.À-ɏ]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

/**
 * One sparkline per row (or column) of `data`, placed in the cells of `place` (a column of as many cells as
 * `data` has rows, or a row of as many cells as it has columns); null when they do not fit.
 */
export function sparkItems(sheetName: string, data: Rect, place: Rect, dataSheet = sheetName): SparklineGroup["items"] | null {
  const cells: { row: number; col: number }[] = [];
  for (let r = place.r0; r <= place.r1; r++) for (let c = place.c0; c <= place.c1; c++) cells.push({ row: r, col: c });
  const byRow = place.c0 === place.c1 && cells.length === data.r1 - data.r0 + 1;
  const byCol = !byRow && place.r0 === place.r1 && cells.length === data.c1 - data.c0 + 1;
  if (!byRow && !byCol) return null;
  const q = sheetRef(dataSheet);
  return cells.map((c, i) => ({ ...c, f: `${q}!${rangeName(byRow ? { ...data, r0: data.r0 + i, r1: data.r0 + i } : { ...data, c0: data.c0 + i, c1: data.c0 + i })}` }));
}

/** A new group as Excel makes it (Einfügen > Sparklines, recorded): Gewinn/Verlust shows negative points. */
export function newGroup(type: Kind, items: SparklineGroup["items"]): SparklineGroup {
  return { type, colors: defaultSparkColors(), ...(type === "stacked" ? { negative: true } : {}), displayEmptyCellsAs: "gap", items };
}

/** Gruppieren: the selected sparklines become one group with the settings of `lead`'s group (first in the file). */
export function groupSparklines(groups: SparklineGroup[], sel: Rect[], lead: { row: number; col: number }): SparklineGroup[] {
  const hit = (i: { row: number; col: number }) => sel.some((r) => inR(r, i.row, i.col));
  const leader = groups.find((g) => g.items.some((i) => i.row === lead.row && i.col === lead.col)) ?? groups.find((g) => g.items.some(hit));
  if (!leader) return groups;
  const items = groups.flatMap((g) => g.items.filter(hit)).sort((a, b) => a.row - b.row || a.col - b.col);
  const rest = groups.map((g) => ({ ...g, items: g.items.filter((i) => !hit(i)) })).filter((g) => g.items.length);
  return [{ ...structuredClone(leader), items }, ...rest];
}

/** Gruppierung aufheben: each selected sparkline of a group of several becomes a group of its own (first in the file). */
export function ungroupSparklines(groups: SparklineGroup[], sel: Rect[]): SparklineGroup[] {
  const hit = (i: { row: number; col: number }) => sel.some((r) => inR(r, i.row, i.col));
  const single: SparklineGroup[] = [];
  const rest = groups.map((g) => {
    if (g.items.length < 2 || !g.items.some(hit)) return g;
    for (const i of g.items.filter(hit)) single.unshift({ ...structuredClone(g), items: [i] });
    return { ...g, items: g.items.filter((i) => !hit(i)) };
  }).filter((g) => g.items.length);
  return [...single, ...rest];
}

/** Löschen: the selected sparklines, or ("groups") every group the selection touches. */
export function clearSparklines(groups: SparklineGroup[], sel: Rect[], whole: boolean): SparklineGroup[] {
  const hit = (i: { row: number; col: number }) => sel.some((r) => inR(r, i.row, i.col));
  return groups.flatMap((g) => (whole && g.items.some(hit) ? [] : [{ ...g, items: g.items.filter((i) => !hit(i)) }])).filter((g) => g.items.length);
}

/** A palette pick of colorGrid (hex, Word reference) in the form designs.pick takes ("accent1,20000,80000"). */
function gridPick(hex: string, ref?: string): string {
  const key = ref ? /w:themeColor="(\w+)"/.exec(ref)?.[1] : undefined;
  if (!key) return hex;
  const k = ({ background1: "bg1", text1: "tx1", background2: "bg2", text2: "tx2" } as Record<string, string>)[key] ?? key;
  const tint = /themeTint="(\w+)"/.exec(ref!)?.[1], shade = /themeShade="(\w+)"/.exec(ref!)?.[1];
  if (tint) { const p = 100 - Math.round((parseInt(tint, 16) / 255) * 100); return `${k},${(100 - p) * 1000},${p * 1000}`; }
  if (shade) return `${k},${Math.round((parseInt(shade, 16) / 255) * 100) * 1000}`;
  return k;
}

export function sparklinesFeature(app: CalcApp) {
  const sheet = () => app.sheet;
  const sel = () => app.grid.selection;
  const groups = () => sheet().sparklines ?? [];
  const at = (row: number, col: number) => groups().find((g) => g.items.some((i) => i.row === row && i.col === col));

  /** The group of the active cell's sparkline (the tab shows while there is one). */
  function current(): SparklineGroup | undefined {
    const a = sheet().view.active;
    return at(a.row, a.col);
  }
  /** Every group the selection touches (the active cell's first). */
  function touched(): SparklineGroup[] {
    const out = groups().filter((g) => g.items.some((i) => sel().some((r) => inR(r, i.row, i.col))));
    const c = current();
    return c ? [c, ...out.filter((g) => g !== c)] : out;
  }
  const has = () => !!current();

  /** One undo step changing every touched group. */
  function edit(label: string, fn: (g: SparklineGroup) => void): void {
    const s = sheet();
    if (!touched().length) return;
    app.change(label, () => {
      app.history.whole(s);
      const hit = new Set(touched());
      s.sparklines = groups().map((g) => { if (!hit.has(g)) return g; const n = structuredClone(g); fn(n); return n; });
    });
  }
  function replace(label: string, fn: (gs: SparklineGroup[]) => SparklineGroup[]): void {
    const s = sheet();
    app.change(label, () => {
      app.history.whole(s);
      const next = fn(groups());
      s.sparklines = next.length ? next : undefined;
    });
  }

  /** Einfügen > Sparklines: one sparkline per row (or column) of `data` in the cells of `place`; false when they do not fit. */
  function add(type: Kind, data: Rect, place: Rect, dataSheet?: string): boolean {
    const s = sheet();
    const items = sparkItems(s.name, data, place, dataSheet);
    if (!items) return false;
    replace("Sparklines", (gs) => {
      // A cell holds one sparkline: new ones replace those in their cells. Excel puts a new group first.
      const taken = new Set(items.map((i) => `${i.row},${i.col}`));
      const kept = gs.map((g) => ({ ...g, items: g.items.filter((i) => !taken.has(`${i.row},${i.col}`)) })).filter((g) => g.items.length);
      return [newGroup(type, items), ...kept];
    });
    return true;
  }

  const colorOf = (value: string): SparkColor | null => {
    const p = app.designs.pick(value);
    return p ? { rgb: p.hex, ...(p.theme ? { theme: p.theme } : {}) } : null;
  };
  function setColor(key: SparkColorKey, value: string, label: string): void {
    const c = colorOf(value);
    if (c) edit(label, (g) => { g.colors[key] = c; });
  }
  /** "Weitere Farben...": the system's colour chooser. */
  function moreColor(key: SparkColorKey, label: string): void {
    const input = h("input", { type: "color", value: sparkRgb(current()?.colors[key] ?? { rgb: "#000000" }, app.wb.theme).toLowerCase(), style: "position:fixed;opacity:0;width:0;height:0" }) as HTMLInputElement;
    document.body.append(input);
    input.addEventListener("change", () => { setColor(key, input.value.slice(1).toUpperCase(), label); input.remove(); });
    input.addEventListener("blur", () => setTimeout(() => input.remove(), 1000));
    input.click();
  }
  const paletteFlyout = (key: SparkColorKey, label: string) => (): HTMLElement =>
    h("div", { class: "menu palette" }, ...colorGrid((hex, ref) => setColor(key, gridPick(hex, ref), label), app.designs.paletteTheme().colors),
      h("hr"), h("button", { onclick: () => moreColor(key, label) }, "Weitere Farben..."));

  // ----- dialogs (labels as Excel's, explanations in our own words) ---------------------------------
  const field = (label: string, value: string, width = 220) => {
    const input = h("input", { type: "text", value, style: `width:${width}px` }) as HTMLInputElement;
    return { input, el: h("label", { class: "field", style: "display:flex;flex-direction:column;gap:4px;margin:6px 0" }, label, input) };
  };
  const ok = (title: string, body: HTMLElement) => showDialog({ title, body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  const refText = (f: string) => { const r = splitRef(f, sheet().name); return r.sheet === sheet().name ? f.slice(f.lastIndexOf("!") + 1) : f; };
  /** A typed range: "A1:E3" (this sheet) or "Blatt!A1:E3". */
  const readRef = (text: string): { sheet: string; rect: Rect } | null => {
    const t = text.trim().replace(/^=/, "");
    const r = splitRef(t.includes("!") ? t : `${sheetRef(sheet().name)}!${t}`, sheet().name);
    return r.rect && (r.sheet === sheet().name || app.wb.sheetByName(r.sheet)) ? { sheet: app.wb.sheetByName(r.sheet)?.name ?? r.sheet, rect: r.rect } : null;
  };
  const bounds = (rs: Rect[]): Rect => ({ r0: Math.min(...rs.map((r) => r.r0)), c0: Math.min(...rs.map((r) => r.c0)), r1: Math.max(...rs.map((r) => r.r1)), c1: Math.max(...rs.map((r) => r.c1)) });

  /** Gruppenspeicherort und -daten bearbeiten...: the group's data and where its sparklines are. */
  async function editGroup(): Promise<void> {
    const g = current();
    if (!g) return;
    const srcs = g.items.map((i) => splitRef(i.f, sheet().name));
    const sameSheet = srcs.every((s) => s.sheet === srcs[0].sheet && s.rect);
    const dataText = sameSheet ? refText(`${sheetRef(srcs[0].sheet)}!${rangeName(bounds(srcs.map((s) => s.rect!)))}`) : g.items[0].f;
    const d = field("Datenbereich:", dataText), p = field("Positionsbereich:", rangeName(bounds(g.items.map((i) => ({ r0: i.row, c0: i.col, r1: i.row, c1: i.col })))));
    const body = h("div", null, h("div", { class: "dlg-text" }, "Die Daten der Sparklines:"), d.el, h("div", { class: "dlg-text" }, "Die Zellen, in denen sie stehen:"), p.el);
    if (!(await ok("Sparklines bearbeiten", body))) return;
    const data = readRef(d.input.value), place = parseRange(p.input.value.trim().replace(/\$/g, ""));
    if (!data || !place) { await messageBox("RaabeCalc", "Der Daten- oder Positionsbereich ist ungültig."); return; }
    const items = sparkItems(sheet().name, data.rect, place, data.sheet);
    if (!items) { await messageBox("RaabeCalc", "Der Positionsbereich muss eine Zeile oder Spalte sein, die zu den Daten passt."); return; }
    replace("Sparklines bearbeiten", (gs) => {
      const taken = new Set(items.map((i) => `${i.row},${i.col}`));
      return gs.map((x) => (x === g ? { ...structuredClone(x), items } : { ...x, items: x.items.filter((i) => !taken.has(`${i.row},${i.col}`)) })).filter((x) => x.items.length);
    });
  }

  /** Daten einer einzelnen Sparkline bearbeiten...: the active cell's sparkline alone. */
  async function editSingle(): Promise<void> {
    const g = current();
    const a = sheet().view.active;
    const item = g?.items.find((i) => i.row === a.row && i.col === a.col);
    if (!g || !item) return;
    const d = field("Datenbereich der Sparkline:", refText(item.f));
    if (!(await ok("Sparklinedaten bearbeiten", h("div", null, d.el)))) return;
    const data = readRef(d.input.value);
    if (!data) { await messageBox("RaabeCalc", "Der Datenbereich ist ungültig."); return; }
    const f = `${sheetRef(data.sheet)}!${rangeName(data.rect)}`;
    edit("Sparklinedaten bearbeiten", (x) => { if (x.items.some((i) => i.row === a.row && i.col === a.col)) x.items = x.items.map((i) => (i.row === a.row && i.col === a.col ? { ...i, f } : i)); });
  }

  /** Ausgeblendete und leere Zellen...: how empty cells show, whether hidden rows and columns count. */
  async function hiddenCells(): Promise<void> {
    const g = current();
    if (!g) return;
    const radio = (value: SparklineGroup["displayEmptyCellsAs"], label: string) => {
      const r = h("input", { type: "radio", name: "spark-empty", value }) as HTMLInputElement;
      r.checked = g.displayEmptyCellsAs === value;
      return { r, el: h("label", { style: "display:block;margin:3px 0 3px 12px" }, r, " ", label) };
    };
    const opts = [radio("gap", "Lücken"), radio("zero", "Null"), radio("span", "Datenpunkte mit einer Linie verbinden")];
    if (g.type !== "line") opts[2].r.disabled = true;
    const hid = h("input", { type: "checkbox" }) as HTMLInputElement;
    hid.checked = !!g.displayHidden;
    const body = h("div", null, h("fieldset", { style: "border:1px solid #d0d0d0;padding:4px 8px" }, h("legend", null, "Leere Zellen anzeigen als"), ...opts.map((o) => o.el)),
      h("label", { style: "display:block;margin-top:8px" }, hid, " Daten in ausgeblendeten Zeilen und Spalten anzeigen"));
    if (!(await ok("Ausgeblendete und leere Zelleneinstellungen", body))) return;
    const empty = opts.find((o) => o.r.checked)?.r.value as SparklineGroup["displayEmptyCellsAs"] | undefined;
    edit("Ausgeblendete und leere Zellen", (x) => { if (empty) x.displayEmptyCellsAs = empty; if (hid.checked) x.displayHidden = true; else delete x.displayHidden; });
  }

  /** Datumsachsentyp...: the range of dates, one per value of each sparkline. */
  async function dateAxis(): Promise<void> {
    const g = current();
    if (!g) return;
    const d = field("Datumsbereich:", g.dateRange ? refText(g.dateRange) : "");
    if (!(await ok("Sparkline-Datumsbereich", h("div", null, h("div", { class: "dlg-text" }, "Die Zellen mit den Datumswerten zu den Sparklinedaten:"), d.el)))) return;
    const r = readRef(d.input.value);
    const n = (x: Rect) => (x.r1 - x.r0 + 1) * (x.c1 - x.c0 + 1);
    const src = splitRef(g.items[0].f, sheet().name).rect;
    if (!r || (src && n(r.rect) !== n(src)) || (r.rect.r0 !== r.rect.r1 && r.rect.c0 !== r.rect.c1)) {
      await messageBox("RaabeCalc", "Der Datumsbereich muss eine Zeile oder Spalte mit so vielen Zellen sein, wie jede Sparkline Werte hat.");
      return;
    }
    const f = `${sheetRef(r.sheet)}!${rangeName(r.rect)}`;
    edit("Datumsachsentyp", (x) => { x.dateRange = f; });
  }

  /** Benutzerdefinierter Wert... for the vertical axis' minimum or maximum. */
  async function customValue(which: "min" | "max"): Promise<void> {
    const g = current();
    if (!g) return;
    const now = which === "min" ? g.manualMin : g.manualMax;
    const d = field(which === "min" ? "Minimalwert der vertikalen Achse:" : "Höchstwert der vertikalen Achse:", String(now ?? 0).replace(".", ","), 120);
    if (!(await ok("Einstellungen für die vertikale Sparklineachse", h("div", null, d.el)))) return;
    const v = Number(d.input.value.trim().replace(/\./g, "").replace(",", "."));
    if (!d.input.value.trim() || !Number.isFinite(v)) { await messageBox("RaabeCalc", "Bitte eine Zahl eingeben."); return; }
    edit("Achse", (x) => {
      if (which === "min") { x.minAxisType = "custom"; x.manualMin = v; } else { x.maxAxisType = "custom"; x.manualMax = v; }
    });
  }
  function axisType(which: "min" | "max", t: "individual" | "group"): void {
    edit("Achse", (x) => {
      if (which === "min") { if (t === "individual") delete x.minAxisType; else x.minAxisType = t; delete x.manualMin; }
      else { if (t === "individual") delete x.maxAxisType; else x.maxAxisType = t; delete x.manualMax; }
    });
  }

  /** Zeile/Spalte wechseln: possible when the group's data is a square block (Excel greys it out otherwise). */
  function swapBlock(): { sheet: string; rect: Rect } | null {
    const g = current();
    if (!g) return null;
    const srcs = g.items.map((i) => splitRef(i.f, sheet().name));
    if (!srcs.every((s) => s.rect && s.sheet === srcs[0].sheet)) return null;
    const b = bounds(srcs.map((s) => s.rect!));
    return b.r1 - b.r0 === b.c1 - b.c0 && b.r1 > b.r0 && g.items.length === b.r1 - b.r0 + 1 ? { sheet: srcs[0].sheet, rect: b } : null;
  }
  function switchRowCol(): void {
    const blk = swapBlock();
    const g = current();
    if (!blk || !g) return;
    const byRow = splitRef(g.items[0].f, sheet().name).rect!.r0 === splitRef(g.items[0].f, sheet().name).rect!.r1;
    const q = sheetRef(blk.sheet);
    edit("Zeile/Spalte wechseln", (x) => {
      x.items = x.items.map((it, i) => ({ ...it, f: `${q}!${rangeName(byRow ? { ...blk.rect, c0: blk.rect.c0 + i, c1: blk.rect.c0 + i } : { ...blk.rect, r0: blk.rect.r0 + i, r1: blk.rect.r0 + i })}` }));
    });
  }

  // ----- the style gallery ------------------------------------------------------------------------
  const styleColor = (c: (typeof SPARKLINE_STYLES)[number]["colors"][SparkColorKey]): SparkColor =>
    "theme" in c ? { rgb: app.designs.resolve(c.theme), theme: { ...c.theme } } : { rgb: c.rgb };
  function applyStyle(i: number): void {
    const st = SPARKLINE_STYLES[i];
    edit("Sparklineformatvorlage", (g) => { for (const k of Object.keys(st.colors) as SparkColorKey[]) g.colors[k] = styleColor(st.colors[k]); });
  }
  function currentStyle(): number | undefined {
    const g = current();
    if (!g) return undefined;
    const same = (a: SparkColor, b: (typeof SPARKLINE_STYLES)[number]["colors"][SparkColorKey]) =>
      "theme" in b ? a.theme?.t === b.theme.t && Math.abs((a.theme.tint ?? 0) - (b.theme.tint ?? 0)) < 1e-6 : !a.theme && a.rgb.toUpperCase() === b.rgb.toUpperCase();
    const i = SPARKLINE_STYLES.findIndex((s) => (Object.keys(s.colors) as SparkColorKey[]).every((k) => same(g.colors[k], s.colors[k])));
    return i < 0 ? undefined : i;
  }
  const SAMPLE = [3, 5, 2, 6, 4, 7, 1, 5];
  const SAMPLE_WL = [3, -2, 4, 5, -3, 2, -1, 4];
  function tile(i: number, w = 46, hgt = 36): HTMLElement {
    const g = current() ?? newGroup("line", []);
    const st = SPARKLINE_STYLES[i];
    const look: SparklineGroup = { ...g, colors: Object.fromEntries((Object.keys(st.colors) as SparkColorKey[]).map((k) => [k, styleColor(st.colors[k])])) as SparklineGroup["colors"],
      displayXAxis: false, dateRange: undefined, minAxisType: undefined, maxAxisType: undefined, rightToLeft: false };
    const dpr = window.devicePixelRatio || 1;
    const cv = h("canvas", { width: String(Math.round(w * dpr)), height: String(Math.round(hgt * dpr)), style: `width:${w}px;height:${hgt}px;display:block` }) as HTMLCanvasElement;
    const ctx = cv.getContext("2d");
    if (ctx) {
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      const values = look.type === "line" ? SAMPLE : SAMPLE_WL;
      drawSparkline(ctx, look, { values }, { lo: Math.min(...values), hi: Math.max(...values) }, { x: 2, y: 2, w: w - 4, h: hgt - 4 }, 1.6, dpr);
    }
    return cv;
  }
  const tip = (i: number) => SPARKLINE_STYLES[i].name.replace(/^[^,]*, /, "");
  function stylesPanel(close: () => void): HTMLElement {
    const now = currentStyle();
    const grid = h("div", { class: "spark-tiles", style: "display:grid;grid-template-columns:repeat(6, 52px);gap:4px;padding:6px" });
    SPARKLINE_STYLES.forEach((_, i) => grid.append(h("button", {
      class: `spark-tile${i === now ? " on" : ""}`, "data-tip": tip(i), "aria-label": tip(i),
      style: `width:52px;height:42px;padding:2px;border:1px solid ${i === now ? "#c43e1c" : "#d0d0d0"};background:#fff`,
      onclick: () => { close(); applyStyle(i); },
    }, tile(i))));
    return h("div", { class: "menu" }, grid);
  }

  // ----- the tab ----------------------------------------------------------------------------------
  const isLine = () => current()?.type === "line";
  const multi = () => touched().length > 1;
  let tabDef: Tab | null = null;
  let lastSeries: string | null = null;
  function tab(): Tab {
    const b = (cmd: string, key: string, size: "large" | "small" = "large"): Control => ({ type: "button", cmd, key, size });
    const ck = (cmd: string, key: string): Control => ({ type: "check", cmd, key });
    const pointMenu = (): MenuEntry[] => {
      const e = (k: SparkColorKey, label: string, disabled = false): MenuEntry => ({ label, flyout: paletteFlyout(k, "Datenpunktfarbe"), disabled });
      return [e("negative", "Negative Punkte"), e("markers", "Datenpunkte", !isLine()), "-", e("high", "Höchstpunkt"), e("low", "Tiefpunkt"), e("first", "Erster Punkt"), e("last", "Letzter Punkt")];
    };
    const axisMenu = (): MenuEntry[] => {
      const g = current();
      const item = (label: string, run: () => void, checked: boolean, disabled = false): MenuEntry => ({ label, run, checked, disabled });
      const vert = (which: "min" | "max"): MenuEntry[] => {
        const t = (which === "min" ? g?.minAxisType : g?.maxAxisType) ?? "individual";
        return [item("Automatisch für jede Sparkline", () => axisType(which, "individual"), t === "individual"),
          item("Identisch für alle Sparklines", () => axisType(which, "group"), t === "group"),
          item("Benutzerdefinierter Wert...", () => void customValue(which), t === "custom")];
      };
      return [
        { heading: "Horizontale Achsenoptionen" },
        item("Allgemeiner Achsentyp", () => edit("Achse", (x) => { delete x.dateRange; }), !g?.dateRange),
        item("Datumsachsentyp...", () => void dateAxis(), !!g?.dateRange),
        "-",
        item("Achse anzeigen", () => edit("Achse", (x) => { if (x.displayXAxis) delete x.displayXAxis; else x.displayXAxis = true; }), !!g?.displayXAxis),
        item("Daten von rechts nach links darstellen", () => edit("Achse", (x) => { if (x.rightToLeft) delete x.rightToLeft; else x.rightToLeft = true; }), !!g?.rightToLeft),
        { heading: "Optionen für den Minimalwert der vertikalen Achse" }, ...vert("min"),
        { heading: "Optionen für den Höchstwert der vertikalen Achse" }, ...vert("max"),
      ];
    };
    const weights = (): MenuEntry[] => SPARK_WEIGHTS.map(([label, w]) => ({ label, checked: (current()?.lineWeight ?? 0.75) === w, disabled: !isLine(),
      run: () => edit("Stärke", (x) => { if (w === 0.75) delete x.lineWeight; else x.lineWeight = w; }) }));
    return tabDef ??= {
      id: "sparkline", label: "Sparkline", key: "JD", contextual: true, shrink: [],
      groups: [
        { id: "spark", label: "Sparkline", key: "ZS", icon: "edit", controls: [
          { type: "split", cmd: "spark.editData", key: "B", size: "large", menu: [{ cmd: "spark.editGroup" }, { cmd: "spark.editSingle" }, "-", { cmd: "spark.hidden" }, { cmd: "spark.switchRowCol" }] }] },
        { id: "type", label: "Typ", key: "ZT", icon: "chart-line", controls: [b("spark.typeLine", "I"), b("spark.typeColumn", "Ä"), b("spark.typeWinLoss", "G")] },
        { id: "show", label: "Anzeigen", key: "ZA", icon: "chart-line", controls: [
          { type: "stack", items: [ck("spark.high", "H"), ck("spark.low", "T"), ck("spark.negative", "N")] },
          { type: "stack", items: [ck("spark.first", "E"), ck("spark.last", "L"), ck("spark.markers", "M")] }] },
        { id: "style", label: "Formatvorlage", key: "ZF", icon: "styles", controls: [
          { type: "tiles", cmd: "spark.styles", key: "S", count: SPARKLINE_STYLES.length, visible: 6, size: [46, 36], cls: "spark-tiles-ribbon",
            render: (i) => tile(i), tip, pick: (i) => applyStyle(i), current: currentStyle,
            redraw: () => { const g = current(); return g ? `${g.type}|${JSON.stringify(g.colors)}|${g.markers}|${g.high}|${g.low}|${g.first}|${g.last}|${g.negative}|${g.lineWeight}` : ""; },
            panel: (close) => stylesPanel(close) },
          { type: "stack", items: [
            { type: "color", cmd: "spark.color", key: "P", palette: "font", showLabel: true, current: () => sparkRgb(current()?.colors.series ?? { rgb: "#376092" }, app.wb.theme).slice(1),
              pick: (c) => { if (c !== "auto") { lastSeries = c; setColor("series", c, "Sparklinefarbe"); } }, theme: () => app.designs.paletteTheme(),
              extra: () => [{ label: "Stärke", submenu: weights }] },
            { type: "dropdown", cmd: "spark.pointColor", key: "D", size: "small", get menu() { return pointMenu(); } }] }] },
        { id: "group", label: "Gruppieren", key: "ZG", icon: "group-objects", controls: [
          { type: "dropdown", cmd: "spark.axis", key: "A", size: "large", get menu() { return axisMenu(); } },
          { type: "stack", items: [b("spark.group", "U", "small"), b("spark.ungroup", "R", "small"),
            { type: "split", cmd: "spark.clear", key: "Ö", size: "small", menu: [{ cmd: "spark.clearSelected" }, { cmd: "spark.clearGroups" }] }] }] },
      ],
    };
  }

  /** The tab's commands (app.ts hands its `run`, which also asks Blatt schützen). */
  function implement(run: Run): void {
    const flag = (id: string, k: "high" | "low" | "negative" | "first" | "last" | "markers", label: string, enabled = has) =>
      run(id, () => edit(label, (g) => { const on = !current()?.[k]; if (on) g[k] = true; else delete g[k]; }), enabled, () => !!current()?.[k]);
    const type = (id: string, t: Kind) => run(id, () => edit("Sparklinetyp", (g) => { g.type = t; }), has, () => current()?.type === t);
    run("spark.editData", () => void editGroup(), () => has() && !multi());
    run("spark.editGroup", () => void editGroup(), () => has() && !multi());
    run("spark.editSingle", () => void editSingle(), () => has() && !multi());
    run("spark.hidden", () => void hiddenCells(), has);
    run("spark.switchRowCol", switchRowCol, () => !!swapBlock());
    type("spark.typeLine", "line");
    type("spark.typeColumn", "column");
    type("spark.typeWinLoss", "stacked");
    flag("spark.high", "high", "Höchstpunkt");
    flag("spark.low", "low", "Tiefpunkt");
    flag("spark.negative", "negative", "Negative Punkte");
    flag("spark.first", "first", "Erster Punkt");
    flag("spark.last", "last", "Letzter Punkt");
    flag("spark.markers", "markers", "Datenpunkte", () => has() && isLine());
    run("spark.styles", () => {}, has);
    // The button's face: the colour picked last (as Excel's split colour buttons).
    run("spark.color", () => { if (lastSeries) setColor("series", lastSeries, "Sparklinefarbe"); }, has);
    run("spark.pointColor", () => {}, has);
    run("spark.axis", () => {}, has);
    run("spark.group", () => { const a = sheet().view.active; replace("Gruppieren", (gs) => groupSparklines(gs, sel(), a)); }, () => has() && multi());
    run("spark.ungroup", () => replace("Gruppierung aufheben", (gs) => ungroupSparklines(gs, sel())), () => touched().some((g) => g.items.length > 1));
    run("spark.clear", () => replace("Sparklines löschen", (gs) => clearSparklines(gs, sel(), false)), has);
    run("spark.clearSelected", () => replace("Sparklines löschen", (gs) => clearSparklines(gs, sel(), false)), has);
    run("spark.clearGroups", () => replace("Sparklinegruppen löschen", (gs) => clearSparklines(gs, sel(), true)), has);
  }

  /** For the check script: the series values and vertical ranges a group's sparklines are drawn from. */
  function drawn(g: SparklineGroup) {
    const series = g.items.map((it) => sparkSeries(app.wb, sheet(), g, it.f));
    return { series, ranges: sparkRanges(g, series) };
  }

  return { current, touched, add, edit, setColor, axisType, tab, implement, applyStyle, currentStyle, stylesPanel, drawn, editGroup, editSingle, hiddenCells, dateAxis, customValue };
}
