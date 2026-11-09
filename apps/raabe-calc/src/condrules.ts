// Bedingte Formatierung > Neue Regel... (also "Weitere Regeln..." of the galleries) and Regeln verwalten...
// (Excel's "Neue Formatierungsregel" and "Manager für Regeln zur bedingten Formatierung", D-143): the six rule
// types of Excel's dialog with their descriptions, a format chosen with "Formatieren..." (Schriftschnitt,
// Schriftfarbe, Füllfarbe), and the list of rules with Neue Regel, Regel löschen, Nach oben, Nach unten.
// Labels as Excel's (recorded in results/2026-11-05-excel-inventory.json), explanations in our own words.

import { colorGrid, h, messageBox, showDialog } from "@raabe/shell";
import { cellName, rangeName, toFileFormula, type CondFormat, type CondRule, type Rect } from "@raabe/sheet-engine";
import type { CalcApp } from "./app";
import * as A from "./actions";

type Format = NonNullable<CondRule["format"]>;
type NewRule = Omit<CondRule, "priority">;

const TYPES = [
  "Alle Zellen basierend auf ihren Werten formatieren",
  "Nur Zellen formatieren, die enthalten",
  "Nur obere oder untere Werte formatieren",
  "Nur Werte über oder unter dem Durchschnitt formatieren",
  "Nur eindeutige oder doppelte Werte formatieren",
  "Formel zur Ermittlung der zu formatierenden Zellen verwenden",
] as const;

const OPERATORS: [string, string][] = [
  ["between", "zwischen"], ["notBetween", "nicht zwischen"], ["equal", "gleich"], ["notEqual", "ungleich"],
  ["greaterThan", "größer als"], ["lessThan", "kleiner als"], ["greaterThanOrEqual", "größer oder gleich"], ["lessThanOrEqual", "kleiner oder gleich"],
];

const sel = (label: string, options: [string, string][], value?: string) =>
  h("select", { "aria-label": label }, ...options.map(([v, l]) => h("option", { value: v, selected: v === value || undefined }, l))) as HTMLSelectElement;
const input = (label: string, value = "", width = 120) => h("input", { type: "text", value, "aria-label": label, style: `width:${width}px` }) as HTMLInputElement;
const row = (...kids: (HTMLElement | string)[]) => h("div", { style: "display:flex;gap:6px;align-items:center;margin:4px 0" }, ...kids);

/** A value typed in the dialog as the file keeps it: numbers with a point, =formulas without "=", text quoted. */
export function toFileValue(v: string): string {
  const t = v.trim();
  if (/^-?[\d.]*,?\d+$/.test(t)) return t.replace(/\./g, "").replace(",", ".");
  if (t.startsWith("=")) return t.slice(1);
  return `"${t.replace(/"/g, '""')}"`;
}

/** How a rule reads in the manager's list (our own short wording). */
export function describeRule(r: CondRule): string {
  const op = OPERATORS.find(([v]) => v === r.operator)?.[1] ?? r.operator ?? "";
  const val = (i: number) => (r.formulas?.[i] ?? "").replace(/^"(.*)"$/, "$1");
  switch (r.type) {
    case "cellIs": return `Zellwert ${op} ${val(0)}${r.formulas && r.formulas.length > 1 ? ` und ${val(1)}` : ""}`;
    case "containsText": return `Zelle enthält "${r.text ?? ""}"`;
    case "notContainsText": return `Zelle enthält nicht "${r.text ?? ""}"`;
    case "beginsWith": return `Zelle beginnt mit "${r.text ?? ""}"`;
    case "endsWith": return `Zelle endet mit "${r.text ?? ""}"`;
    case "containsBlanks": return "Leere Zellen";
    case "duplicateValues": return "Doppelte Werte";
    case "uniqueValues": return "Eindeutige Werte";
    case "top10": return `${r.bottom ? "Untere" : "Obere"} ${r.rank ?? 10}${r.percent ? " %" : ""}`;
    case "aboveAverage": return r.above === false ? "Unter dem Durchschnitt" : "Über dem Durchschnitt";
    case "expression": return `Formel: =${r.formulas?.[0] ?? ""}`;
    case "timePeriod": return `Datum: ${r.timePeriod ?? ""}`;
    case "dataBar": return "Datenbalken";
    case "colorScale": return `${r.colors?.length ?? 2}-Farben-Skala`;
    case "iconSet": return "Symbolsatz";
    default: return "Regel";
  }
}

export function condRules(app: CalcApp, addRule: (rule: NewRule) => Promise<void>) {
  /** Formatieren...: Schriftschnitt, Schriftfarbe and Füllfarbe of the rule (a part of Excel's Zellen formatieren). */
  async function formatDialog(start: Format): Promise<Format | null> {
    const f: Format = { ...start };
    const style = sel("Schriftschnitt:", [["", "Standard"], ["i", "Kursiv"], ["b", "Fett"], ["bi", "Fett Kursiv"]], (f.bold ? "b" : "") + (f.italic ? "i" : "") || "");
    const sample = h("div", { class: "fc-sample", style: "padding:6px 10px" }, "AaBbCcYyZz");
    const paint = () => {
      sample.style.background = f.fill ?? "#fff";
      sample.style.color = f.color ?? "#000";
      sample.style.fontWeight = style.value.includes("b") ? "700" : "400";
      sample.style.fontStyle = style.value.includes("i") ? "italic" : "normal";
    };
    style.onchange = paint;
    const fontGrid = h("div", { class: "fc-palette" }, ...colorGrid((c) => { f.color = hexColor(c) || undefined; paint(); }));
    const fillGrid = h("div", { class: "fc-palette" }, ...colorGrid((c) => { f.fill = hexColor(c) || undefined; paint(); }));
    paint();
    const body = h("div", null, row(h("span", null, "Schriftschnitt:"), style),
      h("div", { class: "fc-label" }, "Farbe"), fontGrid, h("button", { type: "button", onclick: () => { f.color = undefined; paint(); } }, "Automatisch"),
      h("div", { class: "fc-label" }, "Hintergrundfarbe"), fillGrid, h("button", { type: "button", onclick: () => { f.fill = undefined; paint(); } }, "Keine Farbe"),
      h("div", { class: "fc-label" }, "Vorschau"), sample);
    const ok = await showDialog({ title: "Zellen formatieren", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return null;
    f.bold = style.value.includes("b") || undefined;
    f.italic = style.value.includes("i") || undefined;
    return f;
  }

  /** The rule a "Neue Formatierungsregel" dialog describes; null when it is not complete. */
  async function ruleDialog(title: string, startType = 1, start?: CondRule): Promise<NewRule | null> {
    let type = startType;
    let format: Format = start?.format ? { ...start.format } : {};
    const list = h("div", { class: "cf-types", role: "listbox", "aria-label": "Regeltyp auswählen", style: "border:1px solid #c8c8c8;padding:2px;margin-bottom:8px" });
    const desc = h("div", { class: "cf-desc" });
    const preview = h("div", { class: "fc-sample", style: "padding:6px 10px;min-width:160px" }, "AaBbCcYyZz");
    const paintPreview = () => {
      preview.style.background = format.fill ?? "#fff";
      preview.style.color = format.color ?? "#000";
      preview.style.fontWeight = format.bold ? "700" : "400";
      preview.style.fontStyle = format.italic ? "italic" : "normal";
      preview.textContent = format.fill || format.color || format.bold || format.italic ? "AaBbCcYyZz" : "Kein Format festgelegt";
    };
    const formatRow = () => row(h("span", null, "Vorschau:"), preview,
      h("button", { type: "button", onclick: async () => { const f = await formatDialog(format); if (f) { format = f; paintPreview(); } } }, "Formatieren..."));
    let read: () => NewRule | null = () => null;
    const show = () => {
      for (const [i, el] of [...list.children].entries()) el.classList.toggle("on", i === type);
      desc.replaceChildren();
      paintPreview();
      if (type === 0) {
        const styleSel = sel("Formatstil", [["2", "2-Farben-Skala"], ["3", "3-Farben-Skala"], ["bar", "Datenbalken"]]);
        const minC = h("input", { type: "color", value: "#F8696B", "aria-label": "Minimumfarbe" }) as HTMLInputElement;
        const midC = h("input", { type: "color", value: "#FFEB84", "aria-label": "Mittelpunktfarbe" }) as HTMLInputElement;
        const maxC = h("input", { type: "color", value: "#63BE7B", "aria-label": "Maximumfarbe" }) as HTMLInputElement;
        const mid = row(h("span", null, "Mittelpunkt"), midC);
        const colors = h("div", null, row(h("span", null, "Minimum"), minC), mid, row(h("span", null, "Maximum"), maxC));
        styleSel.onchange = () => { mid.hidden = styleSel.value !== "3"; minC.disabled = styleSel.value === "bar"; };
        styleSel.onchange(new Event("change"));
        desc.append(h("div", { class: "fc-label" }, "Alle Zellen basierend auf ihren Werten formatieren:"), row(h("span", null, "Formatstil"), styleSel), colors);
        read = () => {
          const up = (i: HTMLInputElement) => i.value.toUpperCase();
          if (styleSel.value === "bar") return { type: "dataBar", colors: [up(maxC)] };
          return { type: "colorScale", colors: styleSel.value === "3" ? [up(minC), up(midC), up(maxC)] : [up(minC), up(maxC)] };
        };
        return;
      }
      if (type === 1) {
        const what = sel("Nur Zellen formatieren mit", [["value", "Zellwert"], ["text", "Bestimmter Text"]]);
        const op = sel("Operator", OPERATORS, "between");
        const textOp = sel("Operator", [["containsText", "enthält"], ["notContainsText", "enthält nicht"], ["beginsWith", "beginnend mit"], ["endsWith", "endend mit"]]);
        const v1 = input("Erster Wert"), v2 = input("Zweiter Wert"), und = h("span", null, "und");
        const line = row(what, op, textOp, v1, und, v2);
        const sync = () => {
          const isValue = what.value === "value", isText = what.value === "text";
          op.hidden = !isValue; textOp.hidden = !isText; v1.hidden = !(isValue || isText);
          const two = isValue && (op.value === "between" || op.value === "notBetween");
          und.hidden = !two; v2.hidden = !two;
        };
        what.onchange = sync; op.onchange = sync;
        sync();
        desc.append(h("div", { class: "fc-label" }, "Nur Zellen formatieren mit:"), line, formatRow());
        read = () => {
          if (what.value === "text") return v1.value ? textRule(textOp.value as TextType, v1.value, app.grid.selection[0], format) : null;
          const two = op.value === "between" || op.value === "notBetween";
          if (!v1.value.trim() || (two && !v2.value.trim())) return null;
          return { type: "cellIs", operator: op.value, formulas: two ? [toFileValue(v1.value), toFileValue(v2.value)] : [toFileValue(v1.value)], format };
        };
        return;
      }
      if (type === 2) {
        const which = sel("Oben oder unten", [["top", "Oben"], ["bottom", "Unten"]]);
        const n = h("input", { type: "number", min: "1", max: "1000", value: "10", style: "width:60px", "aria-label": "Rang" }) as HTMLInputElement;
        const pct = h("input", { type: "checkbox", "aria-label": "% des ausgewählten Bereichs" }) as HTMLInputElement;
        desc.append(h("div", { class: "fc-label" }, "Werte formatieren, die in folgenden Rangfolgen sind:"), row(which, n, h("label", { class: "check" }, pct, "% des ausgewählten Bereichs")), formatRow());
        read = () => {
          const v = Math.round(Number(n.value));
          if (!(v >= 1 && v <= (pct.checked ? 100 : 1000))) return null;
          return { type: "top10", rank: v, ...(pct.checked ? { percent: true } : {}), ...(which.value === "bottom" ? { bottom: true } : {}), format };
        };
        return;
      }
      if (type === 3) {
        const which = sel("Über oder unter", [["above", "über"], ["below", "unter"]]);
        desc.append(h("div", { class: "fc-label" }, "Werte formatieren, die folgende Bedingung erfüllen:"), row(which, h("span", null, "dem Durchschnitt des ausgewählten Bereichs")), formatRow());
        read = () => ({ type: "aboveAverage", above: which.value === "above", format });
        return;
      }
      if (type === 4) {
        const which = sel("Alle formatieren", [["dup", "doppelte"], ["uni", "eindeutige"]]);
        desc.append(h("div", { class: "fc-label" }, "Alle formatieren:"), row(which, h("span", null, "Werte im ausgewählten Bereich")), formatRow());
        read = () => ({ type: which.value === "uni" ? "uniqueValues" : "duplicateValues", format });
        return;
      }
      const f = input("Werte formatieren, für die diese Formel wahr ist", start?.type === "expression" ? `=${start.formulas?.[0] ?? ""}` : "=", 300);
      desc.append(h("div", { class: "fc-label" }, "Werte formatieren, für die diese Formel wahr ist:"), f, formatRow());
      read = () => {
        const t = f.value.trim().replace(/^=/, "");
        if (!t) return null;
        let file: string;
        try { file = toFileFormula(t, app.wb).replace(/^=/, ""); } catch { return null; }
        return { type: "expression", formulas: [file], format };
      };
    };
    TYPES.forEach((t, i) => list.append(h("button", { type: "button", role: "option", class: "cf-type", style: "display:block;width:100%;text-align:left", onclick: () => { type = i; show(); } }, `► ${t}`)));
    show();
    const body = h("div", { style: "min-width:520px" }, h("div", { class: "fc-label" }, "Regeltyp auswählen:"), list, h("div", { class: "fc-label" }, "Regelbeschreibung bearbeiten:"), desc);
    for (;;) {
      const ok = await showDialog({ title, body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
      if (!ok) return null;
      const r = read();
      if (r) return r;
      await messageBox("RaabeCalc", "Die Regel ist noch nicht vollständig: Geben Sie einen gültigen Wert oder eine Formel ein.");
    }
  }

  async function newRule(startType = 1): Promise<void> {
    const r = await ruleDialog("Neue Formatierungsregel", startType);
    if (r) await addRule(r);
  }

  /** Regeln verwalten...: the rules of the selection or of the sheet; changes apply with OK as one step. */
  async function manager(): Promise<void> {
    const sheet = app.sheet;
    let work: CondFormat[] = structuredClone(sheet.conditional ?? []);
    const scope = sel("Formatierungsregeln anzeigen für", [["sel", "Aktuelle Auswahl"], ["sheet", "Dieses Arbeitsblatt"]]);
    const table = h("div", { class: "cf-rules", role: "listbox", style: "border:1px solid #c8c8c8;min-height:140px;max-height:260px;overflow:auto;margin:6px 0" });
    let picked: { f: CondFormat; r: CondRule } | null = null;
    const flat = () => work.flatMap((f) => f.rules.map((r) => ({ f, r }))).sort((a, b) => a.r.priority - b.r.priority)
      .filter((x) => scope.value === "sheet" || x.f.ref.some((r) => app.grid.selection.some((s) => A.overlaps(r, s))));
    const render = () => {
      table.replaceChildren(h("div", { style: "display:grid;grid-template-columns:1fr 120px 140px;font-weight:600;padding:2px 4px" }, h("span", null, "Regel (in angezeigter Reihenfolge angewendet)"), h("span", null, "Format"), h("span", null, "Wird angewendet auf")));
      for (const x of flat()) {
        const f = x.r.format ?? {};
        const line = h("button", { type: "button", role: "option", style: `display:grid;grid-template-columns:1fr 120px 140px;width:100%;text-align:left;padding:2px 4px;${picked?.r === x.r ? "background:#cde8ff" : ""}`,
          onclick: () => { picked = x; render(); } },
        h("span", null, describeRule(x.r)),
        h("span", { style: `background:${f.fill ?? "transparent"};color:${f.color ?? "inherit"};font-weight:${f.bold ? 700 : 400}` }, x.r.type === "dataBar" || x.r.type === "colorScale" || x.r.type === "iconSet" ? "" : "AaBbCcYyZz"),
        h("span", null, "=" + x.f.ref.map(rangeName).join(";")));
        table.append(line);
      }
    };
    scope.onchange = render;
    const renumber = (order: CondRule[]) => order.forEach((r, i) => { r.priority = i + 1; });
    const move = (d: -1 | 1) => {
      if (!picked) return;
      const all = work.flatMap((f) => f.rules).sort((a, b) => a.priority - b.priority);
      const i = all.indexOf(picked.r), j = i + d;
      if (j < 0 || j >= all.length) return;
      [all[i], all[j]] = [all[j], all[i]];
      renumber(all);
      render();
    };
    const buttons = row(
      h("button", { type: "button", onclick: async () => {
        const r = await ruleDialog("Neue Formatierungsregel");
        if (!r) return;
        const prio = Math.max(0, ...work.flatMap((f) => f.rules.map((x) => x.priority))) + 1;
        work.push({ ref: app.grid.selection.map((x) => ({ ...x })), rules: [{ ...r, priority: prio }] });
        render();
      } }, "Neue Regel..."),
      h("button", { type: "button", onclick: () => {
        if (!picked) return;
        const p = picked;
        p.f.rules = p.f.rules.filter((x) => x !== p.r);
        work = work.filter((f) => f.rules.length);
        renumber(work.flatMap((f) => f.rules).sort((a, b) => a.priority - b.priority));
        picked = null;
        render();
      } }, "Regel löschen"),
      h("button", { type: "button", "aria-label": "Nach oben", onclick: () => move(-1) }, "▲"),
      h("button", { type: "button", "aria-label": "Nach unten", onclick: () => move(1) }, "▼"));
    render();
    const body = h("div", { style: "min-width:560px" }, row(h("span", null, "Formatierungsregeln anzeigen für:"), scope), buttons, table);
    const ok = await showDialog({ title: "Manager für Regeln zur bedingten Formatierung", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Schließen", value: false }], cancel: false });
    if (!ok || JSON.stringify(work) === JSON.stringify(sheet.conditional ?? [])) return;
    app.change("Bedingte Formatierung", () => {
      app.history.whole(sheet);
      sheet.conditional = work.length ? work : undefined;
    });
  }

  return { newRule, manager, ruleDialog };
}

/**
 * A colour of the shell's colour grid ("C00000") as the model keeps colours ("#C00000"). The grid's bare hex
 * made a cell's fill from Zellen formatieren > Ausfüllen invisible and lost on saving (D-142).
 */
export function hexColor(c: string): string {
  if (!c) return "";
  return /^#?[0-9A-Fa-f]{6}$/.test(c) ? "#" + c.replace("#", "").toUpperCase() : c;
}

export type TextType ="containsText" | "notContainsText" | "beginsWith" | "endsWith";

/**
 * A text rule with the formula Excel writes for it (for the top-left cell): Excel tests the cells with that
 * formula, so a rule without one highlights nothing there (D-143).
 */
export function textRule(type: TextType, text: string, at: Rect, format?: Format): NewRule {
  const c = cellName(at.r0, at.c0), q = `"${text.replace(/"/g, '""')}"`;
  const f = { containsText: `NOT(ISERROR(SEARCH(${q},${c})))`, notContainsText: `ISERROR(SEARCH(${q},${c}))`,
    beginsWith: `LEFT(${c},LEN(${q}))=${q}`, endsWith: `RIGHT(${c},LEN(${q}))=${q}` }[type];
  return { type, operator: type, text, formulas: [f], ...(format ? { format } : {}) };
}
