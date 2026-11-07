// Formel > Konvertierungen > Formeloptionen (D-075): PowerPoint's dialog, smaller than Word's (recorded by the
// inventory): Allgemein with what Kopieren puts on the clipboard as text (MathML or the linear format), where
// lines break at binary and relational operators (Vor, Nach, Doppelt) and how subtraction is repeated, and
// Math. AutoKorrektur. The settings are this user's (RaabeWrite's keys and meanings, D-047); the copy setting
// is used when part of an equation is copied, the break settings when an equation wraps.

import { h, showDialog } from "@raabe/shell";
import { mathAutoCorrect } from "@raabe/text-engine";

export interface MathUserOptions { autoBuild: boolean; copyMathml: boolean; brkBin: "before" | "after" | "repeat"; brkBinSub: "--" | "-+" | "+-" }

const KEY = "raabe.math.options";
const DEFAULTS: MathUserOptions = { autoBuild: true, copyMathml: false, brkBin: "before", brkBinSub: "--" };

export function mathUserOptions(): MathUserOptions {
  try { const v = localStorage.getItem(KEY); return v ? { ...DEFAULTS, ...JSON.parse(v) } : { ...DEFAULTS }; } catch { return { ...DEFAULTS }; }
}

function save(o: MathUserOptions): void { try { localStorage.setItem(KEY, JSON.stringify(o)); } catch { /* private mode */ } }

export async function mathOptionsDialog(): Promise<void> {
  const o = mathUserOptions();
  const radio = (label: string, on: boolean) => { const input = h("input", { type: "radio", name: "mcopy", checked: on }) as HTMLInputElement; return { input, el: h("label", { class: "check indent" }, input, label) }; };
  const ml = radio("MathML als Nur-Text in die Zwischenablage kopieren", o.copyMathml);
  const lin = radio("Lineares Format als Nur-Text in die Zwischenablage kopieren", !o.copyMathml);
  const brk = h("select", { "aria-label": "Zeilen mit binären und relationalen Operatoren umbrechen" },
    ...([["before", "Vor"], ["after", "Nach"], ["repeat", "Doppelt"]] as const).map(([v, l]) => h("option", { value: v, selected: o.brkBin === v }, l))) as HTMLSelectElement;
  const sub = h("select", { "aria-label": "Operatoren für die Subtraktion duplizieren als" },
    ...(["--", "-+", "+-"] as const).map((v) => h("option", { value: v, selected: o.brkBinSub === v }, v))) as HTMLSelectElement;
  const subRow = h("label", { class: "field row indent" }, "Operatoren für die Subtraktion duplizieren als:", sub);
  const sync = () => { sub.disabled = brk.value !== "repeat"; subRow.classList.toggle("disabled", sub.disabled); };
  brk.onchange = sync;
  sync();
  const ac = h("button", { type: "button", onclick: () => void autoCorrectList() }, "Math. AutoKorrektur...");
  const body = h("div", { class: "mathoptions", style: "width:480px" },
    h("fieldset", null, h("legend", null, "Allgemein"),
      h("div", null, "Beim Kopieren einer Formel:"), ml.el, lin.el,
      h("label", { class: "field row" }, "Zeilen mit binären und relationalen Operatoren umbrechen:", brk), subRow,
      h("div", { class: "row end" }, ac)));
  const v = await showDialog<"ok" | "cancel">({ title: "Formeloptionen", body, buttons: [{ label: "OK", value: "ok", primary: true }, { label: "Abbrechen", value: "cancel" }], cancel: "cancel" });
  if (v !== "ok") return;
  save({ ...o, copyMathml: ml.input.checked, brkBin: brk.value as MathUserOptions["brkBin"], brkBinSub: sub.value as MathUserOptions["brkBinSub"] });
}

/** Math. AutoKorrektur: the replacements the equation editor makes (RaabeWrite's list). */
function autoCorrectList() {
  const rows = mathAutoCorrect();
  const table = h("table", { class: "mlist" }, h("thead", null, h("tr", null, h("th", null, "Ersetzen:"), h("th", null, "Durch:"))),
    h("tbody", null, ...rows.map((r) => h("tr", null, ...r.map((x) => h("td", null, x))))));
  return showDialog({ title: "AutoKorrektur von Mathematik", body: h("div", { class: "mlistbox" }, table), buttons: [{ label: "Schließen", value: undefined, primary: true }], cancel: undefined });
}
