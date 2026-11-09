// Prognoseblatt (Daten > Prognose, D-092), as German Excel 365 makes it (recorded with probe_excel_forecast.py: the
// dialog through UI Automation, the sheet Excel created through COM and from its file): the dialog Prognosearbeitsblatt
// erstellen (Liniendiagramm / Säulendiagramm, Prognoseende, Optionen: Prognosestart, Konfidenzintervall, Saisonalität
// automatisch / manuell, Prognosestatistiken aufnehmen, Zeitachsenbereich, Wertebereich, Fehlende Punkte füllen mit,
// Duplikate aggregieren mit) and the new sheet before the data's: a table (TableStyleMedium2) with the timeline, the
// values, "Schätzer(…)" with PROGNOSE.ETS and the confidence bounds with PROGNOSE.ETS.KONFINT ("Untere/Obere
// Konfidenzgrenze(…)" for lines, one "Konfidenzintervall(…)" column for columns), the last value repeated in the
// forecast columns so the lines join, future dates by the timeline's step (calendar months for monthly data), the
// statistics table (Alpha … RMSE with PROGNOSE.ETS.STAT) three columns to the right, and the chart at Excel's place.

import { h, messageBox, showDialog } from "@raabe/shell";
import { newChart, type ChartModel } from "@raabe/charts";
import {
  Sheet, colName, etsAddMonths, fileFormat, fitEts, nextTableName, parseRange, rangeName, regularSeries, toFileFormula,
  type Rect, type SheetTable,
} from "@raabe/sheet-engine";
import type { CalcApp } from "./app";
import { sheetAxes } from "./grid/metrics";

type Run = (id: string, fn: () => void | Promise<void>, enabled?: () => boolean, checked?: () => boolean) => void;

export interface ForecastOptions {
  /** The timeline and values (with the header row when `header`), on `sheet`. */
  sheet: Sheet;
  timeline: Rect;
  values: Rect;
  header: boolean;
  chart: "line" | "column";
  /** Prognoseende (a date serial or a number on the timeline). */
  end: number;
  /** Prognosestart: the forecast starts after it; default the timeline's last point. */
  start?: number;
  /** Konfidenzintervall (0.95), null when unticked. */
  confidence: number | null;
  /** Saisonalität: 1 = automatisch erkennen, else the length set by hand. */
  season: number;
  stats: boolean;
  /** Fehlende Punkte füllen mit: Interpolation (1) or Nullen (0). */
  fill: 0 | 1;
  /** Duplikate aggregieren mit: 1 Mittelwert … 7 Summe (PROGNOSE.ETS's codes). */
  agg: number;
}

/** Duplikate aggregieren mit: the entries (in PROGNOSE.ETS's order; labels as Excel's function names). */
const AGG = ["Mittelwert", "Anzahl", "Anzahl2", "Max", "Median", "Min", "Summe"];

const deFmt = (x: number) => String(x).replace(".", ",");

export function forecastFeature(app: CalcApp) {
  /** The data the dialog starts from: the selection (or the region around the active cell), two columns. */
  function detect(): { timeline: Rect; values: Rect; header: boolean } | null {
    const a = app.sheet.view.active;
    const sel = app.grid.selection.at(-1)!;
    const r = sel.r0 === sel.r1 && sel.c0 === sel.c1 ? app.grid.currentRegion(a.row, a.col) : sel;
    if (r.c1 - r.c0 < 1 || r.r1 - r.r0 < 2) return null;
    const header = typeof app.sheet.value(r.r0, r.c0 + 1) === "string" || typeof app.sheet.value(r.r0, r.c0) === "string";
    return { timeline: { r0: r.r0, c0: r.c0, r1: r.r1, c1: r.c0 }, values: { r0: r.r0, c0: r.c0 + 1, r1: r.r1, c1: r.c0 + 1 }, header };
  }

  function numbers(sheet: Sheet, r: Rect, header: boolean): number[] {
    const out: number[] = [];
    for (let row = r.r0 + (header ? 1 : 0); row <= r.r1; row++) { const v = sheet.value(row, r.c0); out.push(typeof v === "number" ? v : NaN); }
    return out;
  }

  /** The timeline's points after the last one up to `end` (calendar months for monthly data). */
  function future(ts: number[], ys: number[], end: number, fill: 0 | 1, agg: number): number[] | null {
    const s = regularSeries(ts.filter((_, i) => !Number.isNaN(ys[i])), ys.filter((v) => !Number.isNaN(v)), fill === 0, agg as 1);
    if (!s) return null;
    const last = Math.max(...ts);
    const out: number[] = [];
    for (let k = 1; k < 100000; k++) {
      const t = s.months ? etsAddMonths(last, k) : last + k * s.step;
      if (t > end + 1e-9) break;
      out.push(t);
    }
    return out;
  }

  /** The default Prognoseende: as many steps after the data as half its points (recorded: 24 months give 12). */
  function defaultEnd(ts: number[], ys: number[]): number {
    const s = regularSeries(ts, ys);
    const last = Math.max(...ts);
    const k = Math.max(1, Math.ceil(ts.length / 2));
    if (!s) return last;
    return s.months ? etsAddMonths(last, k) : last + k * s.step;
  }

  /** Builds the sheet (one undo step after the new sheet) and returns its name. */
  function build(o: ForecastOptions): string | null {
    const src = o.sheet;
    const tl = numbers(src, o.timeline, o.header), vs = numbers(src, o.values, o.header);
    if (tl.some((t) => Number.isNaN(t)) || tl.length < 3) { void messageBox("RaabeCalc", "Die Zeitachse braucht Zahlen oder Datumswerte."); return null; }
    let n = tl.length;
    // Prognosestart before the last point: only the data up to it is used (Excel's backcast).
    if (o.start !== undefined && o.start < tl[n - 1]) n = Math.max(3, tl.filter((t) => t <= o.start!).length);
    const ts = tl.slice(0, n), ys = vs.slice(0, n);
    const fut = future(ts, ys, o.end, o.fill, o.agg);
    if (!fut) { void messageBox("RaabeCalc", "Die Zeitachse braucht einen gleichmäßigen Abstand zwischen ihren Werten."); return null; }
    if (!fut.length) { void messageBox("RaabeCalc", "Das Prognoseende muss nach dem letzten Wert der Zeitachse liegen."); return null; }
    const headT = o.header ? String(src.value(o.timeline.r0, o.timeline.c0) ?? "Zeitachse") : "Zeitachse";
    const headV = o.header ? String(src.value(o.values.r0, o.values.c0) ?? "Werte") : "Werte";
    const fmtT = src.get(o.timeline.r0 + (o.header ? 1 : 0), o.timeline.c0)?.s;
    const fmtV = src.get(o.values.r0 + (o.header ? 1 : 0), o.values.c0)?.s;
    const line = o.chart === "line";
    const conf = o.confidence !== null;
    const cols = conf ? (line ? 5 : 4) : 3;
    const last = n + 1; // 1-based row of the last data point
    const endRow = n + 1 + fut.length;
    const V = `$B$2:$B$${last}`, T = `$A$2:$A$${last}`;
    const tail = `${o.season};${o.fill}${o.agg !== 1 ? `;${o.agg}` : ""}`;
    const level = deFmt(o.confidence ?? 0.95);
    app.addSheet(false);
    const sheet = app.sheet;
    const wb = app.wb;
    const two = wb.styles.derive(undefined, (s) => { s.numFmt = fileFormat("0,00"); });
    const statFmt = wb.styles.derive(undefined, (s) => { s.numFmt = fileFormat("#.##0,00"); });
    const tables = wb.sheets.flatMap((s) => (s.tables ?? []).map((t) => t.name));
    app.change("Prognoseblatt", () => {
      app.history.whole(sheet);
      const put = (r: number, c: number, v: string | number, s?: number) => sheet.set(r, c, { v, s });
      const f = (r: number, c: number, text: string, s?: number) => sheet.set(r, c, { v: null, f: toFileFormula(text), s });
      const heads = [headT, headV, `Schätzer(${headV})`, ...(conf ? (line ? [`Untere Konfidenzgrenze(${headV})`, `Obere Konfidenzgrenze(${headV})`] : [`Konfidenzintervall(${headV})`]) : [])];
      heads.forEach((x, c) => put(0, c, x));
      for (let i = 0; i < n; i++) { put(i + 1, 0, ts[i], fmtT); if (!Number.isNaN(ys[i])) put(i + 1, 1, ys[i], fmtV); }
      // The lines start at the last value (Excel writes it as a number into the forecast columns).
      if (line) { put(n, 2, ys[n - 1]); if (conf) { put(n, 3, ys[n - 1], two); put(n, 4, ys[n - 1], two); } }
      fut.forEach((t, k) => {
        const r = n + 1 + k, a = `A${r + 1}`;
        put(r, 0, t, fmtT);
        f(r, 2, `=PROGNOSE.ETS(${a};${V};${T};${tail})`);
        if (conf && line) {
          f(r, 3, `=C${r + 1}-PROGNOSE.ETS.KONFINT(${a};${V};${T};${level};${tail})`, two);
          f(r, 4, `=C${r + 1}+PROGNOSE.ETS.KONFINT(${a};${V};${T};${level};${tail})`, two);
        } else if (conf) f(r, 3, `=PROGNOSE.ETS.KONFINT(${a};${V};${T};${level};${tail})`);
      });
      const name = nextTableName(tables);
      const main: SheetTable = { name, ref: { r0: 0, c0: 0, r1: endRow - 1, c1: cols - 1 }, header: true, totals: false, columns: heads.map((x) => ({ name: x })),
        style: "TableStyleMedium2", firstColumn: false, lastColumn: false, rowStripes: true, columnStripes: false, filterButton: true };
      const list = [main];
      if (o.stats) {
        const c0 = cols + 2;
        const labels = ["Alpha", "Beta", "Gamma", "MASE", "SMAPE", "MAE", "RMSE"];
        put(0, c0, "Statistik"); put(0, c0 + 1, "Wert");
        // The statistics name the season found when it was set by hand (recorded: Manuell 6 gives ";6;1").
        labels.forEach((l, k) => { put(k + 1, c0, l); f(k + 1, c0 + 1, `=PROGNOSE.ETS.STAT(${V};${T};${k + 1};${tail})`, statFmt); });
        list.push({ name: nextTableName([...tables, name]), ref: { r0: 0, c0, r1: labels.length, c1: c0 + 1 }, header: true, totals: false, columns: [{ name: "Statistik" }, { name: "Wert" }],
          style: "TableStyleMedium2", firstColumn: false, lastColumn: false, rowStripes: true, columnStripes: false, filterButton: true });
      }
      sheet.tables = list;
    });
    app.autofitCols([...Array(cols + (o.stats ? 4 : 0)).keys()]);
    // The chart where Excel puts it: 401.25 pt from the left, 107.5 pt from the top, 724.5 x 346.5 pt.
    const ax = sheetAxes(sheet);
    const pt = (x: number) => (x * 96) / 72;
    const at = (x: number, y: number) => { const c = ax.cols.at(x), r = ax.rows.at(y); return { col: c, row: r, x: x - ax.cols.pos(c), y: y - ax.rows.pos(r) }; };
    const x0 = pt(401.25), y0 = pt(107.5);
    const q = (s: string) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(s) ? s : `'${s.replace(/'/g, "''")}'`);
    const ref = (c: number, r0: number, r1: number) => `${q(sheet.name)}!$${colName(c)}$${r0}:$${colName(c)}$${r1}`;
    const cat = { ref: ref(0, 2, endRow), cache: [] as (string | number | null)[] };
    const accent2 = wb.theme.accent2 ?? "#ED7D31";
    const series: ChartModel["series"] = [
      { name: { ref: `${q(sheet.name)}!$B$1` }, val: { ref: ref(1, 2, endRow), cache: [] }, ...(line ? { line: { type: "solid", width: 2.25 } } : {}) },
      { name: { ref: `${q(sheet.name)}!$C$1` }, cat, val: { ref: ref(2, 2, endRow), cache: [] }, ...(line ? { line: { type: "solid", width: 2 } } : {}) },
    ];
    // Columns: the interval as Fehlerindikatoren on the forecast, plus and minus from the Konfidenzintervall column, grey at
    // 40 % opacity (recorded from Excel's file, D-125).
    if (conf && !line) {
      const amount = { ref: ref(3, 2, endRow), cache: [] as (number | null)[] };
      series[1].errorBars = { dir: "both", type: "cust", plus: amount, minus: { ...amount, cache: [] }, line: { type: "solid", width: 0.75, color: "#595959", transparency: 59.608 } };
    }
    if (conf && line) for (const c of [3, 4]) series.push({ name: { ref: `${q(sheet.name)}!$${colName(c)}$1` }, cat, val: { ref: ref(c, 2, endRow), cache: [] }, line: { type: "solid", width: 1, color: accent2 }, color: accent2 });
    const m = newChart(line ? "line" : "column", series, at(x0, y0), at(x0 + pt(724.5), y0 + pt(346.5)));
    m.title = { show: false };
    m.legend = "b" as ChartModel["legend"];
    if (!line) { m.gapWidth = 0; m.overlap = 100; }
    const others = wb.sheets.flatMap((s) => s.charts ?? []).length;
    m.name = `Diagramm ${others + 1}`;
    app.change("Prognoseblatt", () => { app.history.charts(sheet); sheet.charts = [...(sheet.charts ?? []), m]; });
    app.charts.refresh();
    app.grid.selectCell(0, 0);
    return sheet.name;
  }

  /** Prognosearbeitsblatt erstellen. */
  async function dialog(): Promise<void> {
    if (app.grid.isEditing) app.grid.commitEdit("none");
    const d = detect();
    if (!d) { await messageBox("RaabeCalc", "Wählen Sie zwei Spalten aus: eine Zeitachse (Datumswerte oder Zahlen) und die Werte dazu."); return; }
    const src = app.sheet;
    const tl = numbers(src, d.timeline, d.header), vs = numbers(src, d.values, d.header);
    const fmtOf = (r: Rect) => app.wb.styles.get(src.get(r.r1, r.c0)?.s).numFmt;
    const isDate = /[dmy]/i.test(fmtOf(d.timeline).replace(/\[[^\]]*\]|"[^"]*"/g, ""));
    const showT = (x: number) => (isDate ? new Date(Date.UTC(1899, 11, 30) + x * 86400000).toLocaleDateString("de-DE", { day: "2-digit", month: "2-digit", year: "numeric", timeZone: "UTC" }) : deFmt(x));
    const readT = (s: string): number | null => {
      const m = /^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/.exec(s.trim());
      if (m) { const y = Number(m[3]) < 100 ? 2000 + Number(m[3]) : Number(m[3]); return Math.round((Date.UTC(y, Number(m[2]) - 1, Number(m[1])) - Date.UTC(1899, 11, 30)) / 86400000); }
      const x = Number(s.trim().replace(",", "."));
      return Number.isFinite(x) ? x : null;
    };
    const lastT = Math.max(...tl.filter((x) => !Number.isNaN(x)));
    let kind: "line" | "column" = "line";
    const lineBtn = h("button", { type: "button", "aria-pressed": "true", style: "padding:6px 10px" }, "Liniendiagramm erstellen") as HTMLButtonElement;
    const colBtn = h("button", { type: "button", "aria-pressed": "false", style: "padding:6px 10px" }, "Säulendiagramm erstellen") as HTMLButtonElement;
    const setKind = (k: "line" | "column") => { kind = k; lineBtn.setAttribute("aria-pressed", String(k === "line")); colBtn.setAttribute("aria-pressed", String(k === "column")); lineBtn.style.fontWeight = k === "line" ? "700" : "400"; colBtn.style.fontWeight = k === "column" ? "700" : "400"; };
    lineBtn.addEventListener("click", () => setKind("line"));
    colBtn.addEventListener("click", () => setKind("column"));
    setKind("line");
    const input = (v: string, w = 120) => h("input", { type: "text", value: v, style: `width:${w}px` }) as HTMLInputElement;
    const end = input(showT(defaultEnd(tl.filter((x) => !Number.isNaN(x)), vs.filter((x) => !Number.isNaN(x)))));
    const start = input(showT(lastT));
    const confOn = h("input", { type: "checkbox", checked: true }) as HTMLInputElement;
    const confVal = input("95 %", 60);
    const autoS = h("input", { type: "radio", name: "fcseason", checked: true }) as HTMLInputElement;
    const manS = h("input", { type: "radio", name: "fcseason" }) as HTMLInputElement;
    const ser = regularSeries(tl, vs);
    const found = ser ? fitEts(ser.y, 1).m : 0;
    const seasonVal = input(String(found || 1), 50);
    const stats = h("input", { type: "checkbox" }) as HTMLInputElement;
    const q = (s: string) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(s) ? s : `'${s.replace(/'/g, "''")}'`);
    const tlRef = input(`${q(src.name)}!${rangeName(d.timeline).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2")}`, 200);
    const vRef = input(`${q(src.name)}!${rangeName(d.values).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2")}`, 200);
    const fill = h("select", null, h("option", { value: "1" }, "Interpolation"), h("option", { value: "0" }, "Nullen")) as HTMLSelectElement;
    const agg = h("select", null, ...AGG.map((a, i) => h("option", { value: String(i + 1) }, a))) as HTMLSelectElement;
    const opts = h("div", { style: "display:none;flex-direction:column;gap:6px;border-top:1px solid #ddd;padding-top:8px" },
      h("label", { class: "field" }, "Prognosestart ", start),
      h("label", null, confOn, " Konfidenzintervall ", confVal),
      h("div", null, "Saisonalität"),
      h("label", null, autoS, " Automatisch erkennen"), h("label", null, manS, " Manuell festlegen ", seasonVal),
      h("label", null, stats, " Prognosestatistiken aufnehmen"),
      h("label", { class: "field" }, "Zeitachsenbereich ", tlRef),
      h("label", { class: "field" }, "Wertebereich ", vRef),
      h("label", { class: "field" }, "Fehlende Punkte füllen mit ", fill),
      h("label", { class: "field" }, "Duplikate aggregieren mit ", agg));
    const optBtn = h("button", { type: "button", style: "align-self:flex-start" }, "Optionen ▸") as HTMLButtonElement;
    optBtn.addEventListener("click", () => { const open = opts.style.display === "none"; opts.style.display = open ? "flex" : "none"; optBtn.textContent = open ? "Optionen ▾" : "Optionen ▸"; });
    const body = h("div", { style: "display:flex;flex-direction:column;gap:10px;min-width:440px" },
      h("div", null, "Aus den vorhandenen Daten wird ein neues Blatt mit der Prognose, ihrem Diagramm und den Grenzen der Unsicherheit erstellt."),
      h("div", { style: "display:flex;gap:8px" }, lineBtn, colBtn),
      h("label", { class: "field" }, "Prognoseende ", end), optBtn, opts);
    const ok = await showDialog({ title: "Prognosearbeitsblatt erstellen", body, buttons: [{ label: "Erstellen", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const e = readT(end.value), s = readT(start.value);
    if (e === null) { await messageBox("RaabeCalc", "Das Prognoseende ist kein gültiges Datum."); return; }
    const pr = (v: string) => { const t = v.replace(/^=/, ""); const bang = t.lastIndexOf("!"); return parseRange((bang >= 0 ? t.slice(bang + 1) : t).replace(/\$/g, "")); };
    const tr = pr(tlRef.value) ?? d.timeline, vr = pr(vRef.value) ?? d.values;
    const level = Number(confVal.value.replace("%", "").trim().replace(",", ".")) / 100;
    build({
      sheet: src, timeline: tr, values: vr, header: d.header, chart: kind, end: e, start: s ?? undefined,
      confidence: confOn.checked && level > 0 && level < 1 ? level : null,
      season: manS.checked ? Math.max(0, Math.round(Number(seasonVal.value) || 0)) : 1, stats: stats.checked,
      fill: fill.value === "0" ? 0 : 1, agg: Number(agg.value) || 1,
    });
  }

  function implement(run: Run): void {
    app.commands.register({ id: "data.forecast", label: "Prognoseblatt", icon: "trend" });
    run("data.forecast", () => void dialog(), () => !app.grid.isEditing);
  }

  return { implement, build, detect, dialog };
}
