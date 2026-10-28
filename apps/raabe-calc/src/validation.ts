// Datenüberprüfung as German Excel 365 does it (recorded with probe_excel_validation.py):
// the dialog with Einstellungen, Eingabemeldung and Fehlermeldung; the list arrow beside the
// active cell and its list below the cell; the input message beside the cell; and the check of
// a typed entry with Excel's three kinds of error message (Stopp, Warnung, Information).
// Nothing is drawn on the canvas: arrow, list and message are DOM inside the grid.

import {
  CellError, cellKey, criterionAt, evaluate, formatValue, formulaRefs, listItems, literalList, numberEditText, parseInput, toFileFormula,
  toLocalFormula, validationAt, validationRanges, checkValue, listValues,
  type Rect, type Sheet, type Validation, type ValidationOperator, type ValidationType, type Value, type Workbook,
} from "@raabe/sheet-engine";
import { h, messageBox, showDialog } from "@raabe/shell";
import type { CalcApp } from "./app";
import * as A from "./actions";
import { readInput, shown } from "./engine";
import "./validation.css";

/** Zulassen, with Excel's German entries. */
export const TYPES_DE: [ValidationType, string][] = [
  ["any", "Jeden Wert"], ["whole", "Ganze Zahl"], ["decimal", "Dezimal"], ["list", "Liste"], ["date", "Datum"], ["time", "Zeit"],
  ["textLength", "Textlänge"], ["custom", "Benutzerdefiniert"],
];

/** Daten, with Excel's German entries. */
export const OPERATORS_DE: [ValidationOperator, string][] = [
  ["between", "zwischen"], ["notBetween", "nicht zwischen"], ["equal", "gleich"], ["notEqual", "ungleich"], ["greaterThan", "größer als"],
  ["lessThan", "kleiner als"], ["greaterThanOrEqual", "größer als oder gleich"], ["lessThanOrEqual", "kleiner als oder gleich"],
];

export const STYLES_DE: [NonNullable<Validation["errorStyle"]>, string][] = [["stop", "Stopp"], ["warning", "Warnung"], ["information", "Information"]];

/** Our own wording for the message when the validation has none (Excel's text is not copied, D-058). */
export const DEFAULT_ERROR = "Der eingegebene Wert ist in dieser Zelle nicht zulässig (Datenüberprüfung).";
export const DEFAULT_TITLE = "RaabeCalc";

/** The labels of the criteria fields for a type and operator, as Excel's dialog names them. */
export function criteriaLabels(type: ValidationType, op: ValidationOperator = "between"): string[] {
  if (type === "any") return [];
  if (type === "list") return ["Quelle:"];
  if (type === "custom") return ["Formel:"];
  const [lo, hi, one] = type === "date" ? ["Anfangsdatum:", "Enddatum:", "Datum:"] : type === "time" ? ["Anfangszeit:", "Endzeit:", "Zeit:"]
    : ["Minimum:", "Maximum:", type === "textLength" ? "Länge:" : "Wert:"];
  if (op === "between" || op === "notBetween") return [lo, hi];
  if (op === "equal" || op === "notEqual") return [one];
  return [op === "greaterThan" || op === "greaterThanOrEqual" ? lo : hi];
}

const NUMBER = /^-?\d+(\.\d+)?(E[+-]?\d+)?$/i;

/**
 * A criterion as the dialog shows it (German): lists as "Ja;Nein", dates as TT.MM.JJJJ, times
 * as hh:mm:ss, numbers with a decimal comma, formulas and references with "=" (recorded:
 * Excel's Formula1 reads back this way).
 */
export function toLocalCriterion(v: Validation, which: 1 | 2, wb?: Workbook): string {
  const f = which === 1 ? v.formula1 : v.formula2;
  if (f === undefined || f === "") return "";
  if (which === 1 && v.type === "list") {
    const items = listItems(v);
    if (items) return items.join(";");
  }
  if (NUMBER.test(f) && v.type !== "list") {
    const n = Number(f);
    if (v.type === "date") return formatValue(n, "dd.mm.yyyy").text;
    if (v.type === "time") return formatValue(n, "hh:mm:ss").text;
    return numberEditText(n);
  }
  // A constant text shows without "=" and quotes (Excel: ="x" reads back as x).
  if (/^"(?:[^"]|"")*"$/.test(f) && v.type !== "list") return f.slice(1, -1).replace(/""/g, '"');
  return "=" + toLocalFormula(f, wb);
}

/**
 * A criterion typed in the dialog in the file's form; null when it is not valid for the type.
 * Lists split at ";" (German Excel's list separator) unless they start with "=".
 */
export function toFileCriterion(type: ValidationType, text: string, wb?: Workbook): Pick<Validation, "formula1" | "items"> | string | null {
  const t = text.trim();
  if (t === "") return null;
  if (t.startsWith("=")) {
    try {
      return toFileFormula(t, wb);
    } catch {
      return null;
    }
  }
  if (type === "list") return literalList(text.split(";"));
  if (type === "custom") {
    const p = parseInput(t);
    if (typeof p.value === "number") return String(p.value);
    if (typeof p.value === "boolean") return p.value ? "TRUE" : "FALSE";
    return `"${t.replace(/"/g, '""')}"`;
  }
  const p = parseInput(t);
  if (typeof p.value !== "number" || p.formula !== undefined) return null;
  if (type === "whole" && !Number.isInteger(p.value)) return null;
  return String(p.value);
}

/** What a typed entry becomes in a cell: its value, or a formula's result there. */
export function entryValue(wb: Workbook, sheetIndex: number, row: number, col: number, text: string): Value {
  const sheet = wb.sheets[sheetIndex];
  const cell = sheet.get(row, col);
  const style = wb.styles.get(cell?.s ?? sheet.rows.get(row)?.s ?? sheet.cols.get(col)?.s);
  const r = readInput(text, style);
  if (r.f === undefined) return r.v;
  try {
    const x = evaluate(wb, sheetIndex, row, col, r.f);
    return Array.isArray(x) ? (x[0]?.[0] ?? null) : x;
  } catch {
    return new CellError("#VALUE!");
  }
}

/** Whether a typed entry passes the cell's validation (true when the cell has none). */
export function isValidEntry(wb: Workbook, sheetIndex: number, row: number, col: number, text: string): boolean {
  const v = validationAt(wb.sheets[sheetIndex], row, col);
  if (!v || text === "") return true;
  return checkValue(wb, sheetIndex, v, row, col, entryValue(wb, sheetIndex, row, col, text));
}

/** Rectangles of `r` outside `cut`. */
function subtract(r: Rect, cut: Rect): Rect[] {
  if (cut.r1 < r.r0 || cut.r0 > r.r1 || cut.c1 < r.c0 || cut.c0 > r.c1) return [r];
  const out: Rect[] = [];
  if (cut.r0 > r.r0) out.push({ r0: r.r0, c0: r.c0, r1: cut.r0 - 1, c1: r.c1 });
  if (cut.r1 < r.r1) out.push({ r0: cut.r1 + 1, c0: r.c0, r1: r.r1, c1: r.c1 });
  const r0 = Math.max(r.r0, cut.r0), r1 = Math.min(r.r1, cut.r1);
  if (cut.c0 > r.c0) out.push({ r0, c0: r.c0, r1, c1: cut.c0 - 1 });
  if (cut.c1 < r.c1) out.push({ r0, c0: cut.c1 + 1, r1, c1: r.c1 });
  return out;
}

const overlaps = (a: Rect, b: Rect) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;

/** The sheet's validations with `ranges` set to `v` (or to none): the old ones give those cells up. */
export function applyValidation(sheet: Sheet, ranges: Rect[], v: Omit<Validation, "ref"> | null): void {
  const list: Validation[] = [];
  for (const old of sheet.validations ?? []) {
    let ref = old.ref;
    for (const cut of ranges) ref = ref.flatMap((r) => subtract(r, cut));
    if (ref.length) list.push(ref === old.ref ? old : { ...old, ref });
  }
  if (v) list.push({ ...v, ref: ranges.map((r) => ({ ...r })) });
  sheet.validations = list.length ? list : undefined;
}

const sameSettings = (a: Validation, b: Validation) => JSON.stringify({ ...a, ref: 0 }) === JSON.stringify({ ...b, ref: 0 });

/** Whether a validation does anything (Jeden Wert without an input message does not). */
const meaningful = (v: Omit<Validation, "ref">) => v.type !== "any" || (!!v.showInputMessage && !!(v.promptTitle || v.prompt));

type Answer = "ok" | "retry" | "cancel";

export function validationFeature(app: CalcApp, enter?: (target: Rect, text: string) => void) {
  // Without the app's own path (gridHost.commit), entries go in as typed input does.
  const commit = enter ?? ((target: Rect, text: string) => app.change("Eingabe", () => A.enter(app.ctx, app.sheet, target, text)));
  let arrow: HTMLElement | null = null;
  let prompt: HTMLElement | null = null;
  let list: { el: HTMLElement; items: string[]; index: number; row: number; col: number } | null = null;

  const active = () => app.sheet.view.active;
  const cellRect = (row: number, col: number): Rect => app.sheet.mergeAt(row, col) ?? { r0: row, c0: col, r1: row, c1: col };
  const sheetIndex = () => app.wb.sheets.indexOf(app.sheet);

  // ----- arrow, list and input message --------------------------------------------------------

  /** The selection moved (or the sheet, the workbook, the validations changed). */
  function afterSelect(): void {
    closeList();
    const { row, col } = active();
    const v = validationAt(app.sheet, row, col);
    const wantsArrow = v?.type === "list" && !v.showDropDown;
    if (wantsArrow && !arrow) {
      arrow = h("button", { class: "dv-arrow", tabindex: "-1", "aria-label": "Dropdownliste" });
      arrow.innerHTML = '<svg viewBox="0 0 9 5"><path d="M0 0h9L4.5 5z" fill="currentColor"/></svg>';
      arrow.addEventListener("mousedown", (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (list) closeList();
        else openList();
      });
      app.grid.el.append(arrow);
    } else if (!wantsArrow && arrow) {
      arrow.remove();
      arrow = null;
    }
    const wantsPrompt = !!v?.showInputMessage && !!(v.promptTitle || v.prompt);
    if (wantsPrompt) {
      prompt ??= h("div", { class: "dv-prompt", role: "tooltip" });
      prompt.replaceChildren(...(v!.promptTitle ? [h("b", null, v!.promptTitle)] : []), v!.prompt ?? "");
      if (!prompt.isConnected) app.grid.el.append(prompt);
    } else if (prompt) {
      prompt.remove();
      prompt = null;
    }
    place();
  }

  /** Puts arrow, list and message beside the active cell (after scrolling, zooming, resizing). */
  function place(): void {
    if (!arrow && !prompt && !list) return;
    const { row, col } = active();
    const b = app.grid.rectOnCanvas(cellRect(row, col));
    const W = app.grid.el.clientWidth, H = app.grid.el.clientHeight;
    const visible = !!b && b.x + b.w > 0 && b.y + b.h > 0 && b.x < W && b.y < H;
    if (arrow) {
      arrow.hidden = !visible;
      if (b) {
        const hgt = Math.max(14, Math.min(b.h, 21));
        arrow.style.left = `${b.x + b.w + 1}px`;
        arrow.style.top = `${b.y + b.h - hgt}px`;
        arrow.style.height = `${hgt}px`;
      }
    }
    if (prompt) {
      prompt.hidden = !visible;
      if (b) {
        prompt.style.left = `${b.x + Math.min(b.w, 64) / 2}px`;
        prompt.style.top = `${b.y + b.h + 8}px`;
      }
    }
    if (list && b) {
      list.el.style.left = `${b.x}px`;
      list.el.style.top = `${b.y + b.h + 1}px`;
      list.el.style.minWidth = `${b.w + (arrow ? 18 : 0)}px`;
    }
  }

  /** The entries the list offers, as the cells show them (Excel lists the source's texts). */
  function entries(v: Validation, row: number, col: number): string[] {
    const items = listItems(v);
    if (items) return items;
    const src = sourceRange(v, row, col);
    if (src) {
      const out: string[] = [];
      const { sheet, rect } = src;
      for (let r = rect.r0; r <= Math.min(rect.r1, rect.r0 + 10000); r++) {
        for (let c = rect.c0; c <= Math.min(rect.c1, rect.c0 + 200); c++) {
          const cell = sheet.cells.get(cellKey(r, c));
          if (!cell || cell.v === null || cell.v === "") continue;
          out.push(shown(cell, app.wb.styles.get(cell.s ?? sheet.rows.get(r)?.s ?? sheet.cols.get(c)?.s)).text);
        }
      }
      return out;
    }
    return listValues(app.wb, sheetIndex(), v, row, col).map((x) => (typeof x === "number" ? numberEditText(x) : typeof x === "boolean" ? (x ? "WAHR" : "FALSCH") : String(x)));
  }

  /** The cells a list's source names (a reference or a name), to show their formatted texts. */
  function sourceRange(v: Validation, row: number, col: number): { sheet: Sheet; rect: Rect } | null {
    if (!v.formula1) return null;
    let f = criterionAt(v, v.formula1, row, col);
    const here = sheetIndex();
    const name = app.wb.names.find((n) => n.name.toLowerCase() === f.toLowerCase() && n.sheet === here)
      ?? app.wb.names.find((n) => n.name.toLowerCase() === f.toLowerCase() && n.sheet === undefined);
    if (name) f = name.ref;
    try {
      const local = toLocalFormula(f, app.wb);
      const refs = formulaRefs(local);
      if (refs.length !== 1 || refs[0].start !== 0 || refs[0].end !== local.length) return null;
      const sheet = refs[0].sheet !== undefined ? app.wb.sheetByName(refs[0].sheet) : app.sheet;
      return sheet ? { sheet, rect: refs[0].ref } : null;
    } catch {
      return null;
    }
  }

  /** Opens the list below the active cell (the arrow, Alt+Unten); false when the cell has none. */
  function openList(): boolean {
    const { row, col } = active();
    const v = validationAt(app.sheet, row, col);
    if (v?.type !== "list" || v.showDropDown) return false;
    if (app.grid.isEditing) app.grid.commitEdit("none");
    closeList();
    const items = entries(v, row, col);
    const el = h("div", { class: "dv-list", role: "listbox" });
    const current = shown(app.sheet.get(row, col), app.wb.styles.get(app.sheet.get(row, col)?.s)).text;
    items.forEach((t, i) => {
      const item = h("div", { role: "option" }, t);
      item.addEventListener("mousemove", () => highlight(i));
      item.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); choose(i); });
      el.append(item);
    });
    // Eight entries show at once, the rest scroll (Excel).
    el.style.maxHeight = `${8 * 17 + 2}px`;
    list = { el, items, index: items.indexOf(current), row, col };
    app.grid.el.append(el);
    place();
    highlight(list.index);
    addEventListener("keydown", listKey, true);
    addEventListener("pointerdown", outside, true);
    return true;
  }

  function highlight(i: number): void {
    if (!list) return;
    list.index = i;
    [...list.el.children].forEach((c, k) => c.classList.toggle("on", k === i));
    const it = list.el.children[i] as HTMLElement | undefined;
    if (it) {
      if (it.offsetTop < list.el.scrollTop) list.el.scrollTop = it.offsetTop;
      else if (it.offsetTop + it.offsetHeight > list.el.scrollTop + list.el.clientHeight) list.el.scrollTop = it.offsetTop + it.offsetHeight - list.el.clientHeight;
    }
  }

  function closeList(): void {
    if (!list) return;
    list.el.remove();
    list = null;
    removeEventListener("keydown", listKey, true);
    removeEventListener("pointerdown", outside, true);
  }

  function outside(e: Event): void {
    if (list && !list.el.contains(e.target as Node) && !(arrow && arrow.contains(e.target as Node))) closeList();
  }

  function listKey(e: KeyboardEvent): void {
    if (!list) return;
    const n = list.items.length;
    const step = (d: number) => { e.preventDefault(); e.stopPropagation(); highlight(Math.max(0, Math.min(n - 1, (list!.index < 0 && d > 0 ? -1 : list!.index) + d))); };
    if (e.key === "ArrowDown" && !e.altKey) step(1);
    else if (e.key === "ArrowUp" && !e.altKey) step(-1);
    else if (e.key === "PageDown") step(7);
    else if (e.key === "PageUp") step(-7);
    else if (e.key === "Home") step(-n);
    else if (e.key === "End") step(n);
    else if (e.key === "Enter" || e.key === "Tab") {
      e.preventDefault();
      e.stopPropagation();
      if (list.index >= 0) choose(list.index);
      else closeList();
    } else if (e.key === "Escape" || (e.altKey && (e.key === "ArrowUp" || e.key === "ArrowDown"))) {
      e.preventDefault();
      e.stopPropagation();
      closeList();
    } else if (!["Shift", "Control", "Alt"].includes(e.key)) closeList();
  }

  /** An entry chosen: it goes into the cell as if typed. */
  function choose(i: number): void {
    if (!list) return;
    const { row, col, items } = list;
    closeList();
    commit(cellRect(row, col), items[i]);
    app.grid.focus();
  }

  // ----- checking an entry -----------------------------------------------------------------------

  /** True when an entry needs no message (no validation, no error message, or a valid entry). */
  function quickCheck(row: number, col: number, text: string): boolean {
    const v = validationAt(app.sheet, row, col);
    return !v || !v.showErrorMessage || isValidEntry(app.wb, sheetIndex(), row, col, text);
  }

  /**
   * Checks a typed entry before it is committed. "retry": edit the text again (Wiederholen,
   * Nein); "cancel": drop the entry (Abbrechen); "ok": commit it (valid, Ja, OK).
   */
  async function check(row: number, col: number, text: string): Promise<Answer> {
    if (quickCheck(row, col, text)) return "ok";
    return errorBox(validationAt(app.sheet, row, col)!);
  }

  /** Excel's message box for an invalid entry, with the validation's own title and text. */
  function errorBox(v: Validation): Promise<Answer> {
    const style = v.errorStyle ?? "stop";
    const message = v.error || DEFAULT_ERROR;
    const icon = h("span", { class: `dv-icon ${style}`, "aria-hidden": "true" }, style === "stop" ? "✕" : style === "warning" ? "!" : "i");
    const body = h("div", { class: "dv-msg" }, icon, h("div", null, style === "warning" ? `${message}\n\nMöchten Sie fortfahren?` : message));
    const title = v.errorTitle || DEFAULT_TITLE;
    if (style === "warning") {
      return showDialog<Answer>({ title, body, buttons: [{ label: "Ja", value: "ok" }, { label: "Nein", value: "retry", primary: true }, { label: "Abbrechen", value: "cancel" }], cancel: "cancel" });
    }
    if (style === "information") {
      return showDialog<Answer>({ title, body, buttons: [{ label: "OK", value: "ok", primary: true }, { label: "Abbrechen", value: "cancel" }], cancel: "cancel" });
    }
    return showDialog<Answer>({ title, body, buttons: [{ label: "Wiederholen", value: "retry", primary: true }, { label: "Abbrechen", value: "cancel" }], cancel: "cancel" });
  }

  // ----- the dialog ---------------------------------------------------------------------------

  interface Form {
    type: ValidationType; operator: ValidationOperator; c1: string; c2: string; allowBlank: boolean; dropDown: boolean; all: boolean;
    showInput: boolean; promptTitle: string; prompt: string; showError: boolean; errorStyle: NonNullable<Validation["errorStyle"]>; errorTitle: string; error: string;
  }

  const blankForm = (): Form => ({
    type: "any", operator: "between", c1: "", c2: "", allowBlank: true, dropDown: true, all: false,
    showInput: true, promptTitle: "", prompt: "", showError: true, errorStyle: "stop", errorTitle: "", error: "",
  });

  function formOf(v: Validation): Form {
    return {
      type: v.type, operator: v.operator ?? "between", c1: toLocalCriterion(v, 1, app.wb), c2: toLocalCriterion(v, 2, app.wb),
      allowBlank: !!v.allowBlank, dropDown: !v.showDropDown, all: false, showInput: !!v.showInputMessage, promptTitle: v.promptTitle ?? "", prompt: v.prompt ?? "",
      showError: !!v.showErrorMessage, errorStyle: v.errorStyle ?? "stop", errorTitle: v.errorTitle ?? "", error: v.error ?? "",
    };
  }

  /** The dialog's body for a form; `read` writes the fields back into it. */
  function dialogBody(f: Form, tab: number): { body: HTMLElement; read: () => void; tab: () => number } {
    const sel = (pairs: [string, string][], value: string) => {
      const s = h("select", null, ...pairs.map(([k, l]) => h("option", { value: k }, l))) as HTMLSelectElement;
      s.value = value;
      return s;
    };
    const check = (label: string, on: boolean) => {
      const input = h("input", { type: "checkbox", checked: on || undefined }) as HTMLInputElement;
      return { el: h("label", { class: "check" }, input, label), input };
    };
    const text = (value: string) => h("input", { type: "text", value }) as HTMLInputElement;
    const row = (label: string, input: HTMLElement) => h("label", { class: "dv-row" }, h("span", null, label), input);

    // Einstellungen.
    const typeSel = sel(TYPES_DE, f.type);
    const opSel = sel(OPERATORS_DE, f.operator);
    const blank = check("Leere Zellen ignorieren", f.allowBlank);
    const drop = check("Zellendropdown", f.dropDown);
    const c1 = text(f.c1), c2 = text(f.c2);
    const all = check("Diese Änderungen auf alle Zellen mit den gleichen Einstellungen anwenden", f.all);
    const criteria = h("div");
    const opRow = row("Daten:", opSel);
    const layout = () => {
      const t = typeSel.value as ValidationType;
      const labels = criteriaLabels(t, opSel.value as ValidationOperator);
      opRow.hidden = t === "any" || t === "list" || t === "custom";
      drop.el.hidden = t !== "list";
      blank.input.disabled = t === "any";
      blank.el.classList.toggle("off", t === "any");
      criteria.replaceChildren(...labels.map((l, i) => row(l, i ? c2 : c1)));
    };
    typeSel.addEventListener("change", () => {
      // Another kind of data starts with empty criteria (Excel).
      c1.value = "";
      c2.value = "";
      layout();
    });
    opSel.addEventListener("change", layout);
    const clearAll = h("button", { type: "button", class: "dv-clear" }, "Alle löschen");
    const settings = h("div", { class: "dv-panel" },
      h("div", { class: "dv-group" }, "Gültigkeitskriterien"), row("Zulassen:", typeSel), h("div", { class: "dv-checks" }, blank.el, drop.el),
      opRow, criteria, h("hr"), all.el, clearAll);

    // Eingabemeldung.
    const showInput = check("Eingabemeldung anzeigen, wenn Zelle ausgewählt wird", f.showInput);
    const pTitle = text(f.promptTitle);
    const pText = h("textarea", { maxlength: "255" }, f.prompt) as HTMLTextAreaElement;
    pTitle.maxLength = 32;
    const input = h("div", { class: "dv-panel" }, showInput.el, h("div", { class: "dv-group" }, "Bei Auswahl dieser Zelle folgende Eingabemeldung anzeigen:"),
      row("Titel:", pTitle), h("div", null, "Eingabemeldung:"), pText);

    // Fehlermeldung.
    const showError = check("Fehlermeldung anzeigen, wenn ungültige Daten eingegeben wurden", f.showError);
    const styleSel = sel(STYLES_DE, f.errorStyle);
    const eTitle = text(f.errorTitle);
    const eText = h("textarea", { maxlength: "225" }, f.error) as HTMLTextAreaElement;
    eTitle.maxLength = 32;
    const error = h("div", { class: "dv-panel" }, showError.el, h("div", { class: "dv-group" }, "Bei Eingabe ungültiger Daten folgende Fehlermeldung anzeigen:"),
      row("Typ:", styleSel), row("Titel:", eTitle), h("div", null, "Fehlermeldung:"), eText);

    clearAll.addEventListener("click", () => {
      const b = blankForm();
      typeSel.value = b.type; opSel.value = b.operator; c1.value = ""; c2.value = "";
      blank.input.checked = true; drop.input.checked = true; all.input.checked = false;
      showInput.input.checked = true; pTitle.value = ""; pText.value = "";
      showError.input.checked = true; styleSel.value = "stop"; eTitle.value = ""; eText.value = "";
      layout();
    });

    const panels = [settings, input, error];
    const names = ["Einstellungen", "Eingabemeldung", "Fehlermeldung"];
    let current = tab;
    const nav = h("div", { class: "fc-tabs", role: "tablist" });
    const show = (i: number) => {
      current = i;
      panels.forEach((p, k) => { p.hidden = k !== i; });
      [...nav.children].forEach((b, k) => b.classList.toggle("on", k === i));
    };
    names.forEach((n, i) => nav.append(h("button", { type: "button", role: "tab", onclick: () => show(i) }, n)));
    const body = h("div", { class: "dv" }, nav, ...panels);
    layout();
    show(tab);
    const read = () => {
      Object.assign(f, {
        type: typeSel.value as ValidationType, operator: opSel.value as ValidationOperator, c1: c1.value, c2: c2.value,
        allowBlank: blank.input.checked, dropDown: drop.input.checked, all: all.input.checked,
        showInput: showInput.input.checked, promptTitle: pTitle.value, prompt: pText.value,
        showError: showError.input.checked, errorStyle: styleSel.value as Form["errorStyle"], errorTitle: eTitle.value, error: eText.value,
      });
    };
    return { body, read, tab: () => current };
  }

  /** The form as a validation; a message (our own words) when a criterion is missing or wrong. */
  function validationOf(f: Form): Omit<Validation, "ref"> | string {
    const v: Omit<Validation, "ref"> = { type: f.type };
    const labels = criteriaLabels(f.type, f.operator);
    if (f.type !== "any" && f.type !== "list" && f.type !== "custom" && f.operator !== "between") v.operator = f.operator;
    const texts = [f.c1, f.c2];
    for (let i = 0; i < labels.length; i++) {
      if (!texts[i].trim()) return `Bitte einen Wert für „${labels[i].replace(/:$/, "")}“ eingeben.`;
      const c = toFileCriterion(f.type, texts[i], app.wb);
      if (c === null) return `Der Wert für „${labels[i].replace(/:$/, "")}“ passt nicht zu „${TYPES_DE.find(([k]) => k === f.type)![1]}“.`;
      if (typeof c === "string") v[i ? "formula2" : "formula1"] = c;
      else Object.assign(v, c);
    }
    if (labels.length === 2 && v.formula1 && v.formula2 && NUMBER.test(v.formula1) && NUMBER.test(v.formula2) && Number(v.formula1) > Number(v.formula2)) {
      return `„${labels[1].replace(/:$/, "")}“ darf nicht kleiner sein als „${labels[0].replace(/:$/, "")}“.`;
    }
    if (f.allowBlank && f.type !== "any") v.allowBlank = true;
    if (f.type === "list" && !f.dropDown) v.showDropDown = true;
    if (f.showInput) v.showInputMessage = true;
    if (f.promptTitle) v.promptTitle = f.promptTitle;
    if (f.prompt) v.prompt = f.prompt;
    if (f.showError) v.showErrorMessage = true;
    if (f.errorStyle !== "stop") v.errorStyle = f.errorStyle;
    if (f.errorTitle) v.errorTitle = f.errorTitle;
    if (f.error) v.error = f.error;
    return v;
  }

  /** Datenüberprüfung… (Daten > Datentools). */
  async function openDialog(): Promise<void> {
    if (app.grid.isEditing) app.grid.commitEdit("none");
    closeList();
    const sheet = app.sheet;
    let ranges = app.grid.selection.map((r) => ({ ...r }));
    const touched = (sheet.validations ?? []).filter((v) => v.ref.some((r) => ranges.some((s) => overlaps(r, s))));
    const kinds = touched.filter((v, i) => touched.findIndex((w) => sameSettings(v, w)) === i);
    let base: Validation | undefined;
    if (kinds.length > 1) {
      const go = await yesNo("Die Auswahl enthält verschiedene Einstellungen der Datenüberprüfung. Sollen sie gelöscht und neue festgelegt werden?");
      if (!go) return;
    } else if (kinds.length === 1) {
      base = kinds[0];
      const covered = ranges.every((s) => coveredBy(s, base!.ref));
      if (!covered) {
        const go = await yesNo("Ein Teil der Auswahl hat keine Datenüberprüfung. Sollen deren Einstellungen auf die ganze Auswahl übertragen werden?");
        if (!go) return;
      }
    }
    const f = base ? formOf(base) : blankForm();
    let tab = 0;
    for (;;) {
      const d = dialogBody(f, tab);
      const ok = await showDialog({ title: "Datenüberprüfung", body: d.body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
      if (!ok) return;
      d.read();
      tab = d.tab();
      const v = validationOf(f);
      if (typeof v === "string") {
        await messageBox("Datenüberprüfung", v);
        continue;
      }
      if (f.all && base) ranges = [...base.ref.map((r) => ({ ...r })), ...ranges.filter((s) => !coveredBy(s, base!.ref))];
      app.change("Datenüberprüfung", () => {
        app.history.validations(sheet);
        applyValidation(sheet, ranges, meaningful(v) ? v : null);
      });
      afterSelect();
      return;
    }
  }

  const coveredBy = (s: Rect, refs: Rect[]) => refs.reduce<Rect[]>((rest, r) => rest.flatMap((x) => subtract(x, r)), [s]).length === 0;

  function yesNo(text: string): Promise<boolean> {
    return showDialog({ title: "RaabeCalc", body: text, buttons: [{ label: "Ja", value: true, primary: true }, { label: "Nein", value: false }], cancel: false });
  }

  /** Inhalte auswählen > Datenüberprüfung: every cell that has one. */
  function select(): void {
    const ranges = validationRanges(app.sheet);
    if (!ranges.length) return void messageBox("RaabeCalc", "Es wurden keine Zellen gefunden.");
    app.grid.selectMany(ranges);
  }

  return { openDialog, afterSelect, place, check, quickCheck, openList, closeList, select, errorBox };
}
