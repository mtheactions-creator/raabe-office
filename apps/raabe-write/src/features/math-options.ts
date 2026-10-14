// Formel > Konvertierungen > Formeloptionen, as Word's dialog: the document's settings (m:mathPr:
// limits, alignment, margins, line breaks) and this user's own (automatic build-up, what
// Kopieren puts on the clipboard as text, the defaults for new documents).

import { DEFAULT_MATH, FUNCTIONS, LIMIT_FUNCTIONS, mathAutoCorrect, type Editor, type MathSettings } from "@raabe/text-engine";
import { h, showDialog } from "@raabe/shell";
import { CM, parseNum } from "./api";

export interface MathUserOptions {
  /** Ausdrücke automatisch in professionelles Format konvertieren. */
  autoBuild: boolean;
  /** Beim Kopieren einer Formel: MathML (else the linear format) as plain text. */
  copyMathml: boolean;
}

const read = <T>(key: string, fallback: T): T => {
  try { const v = localStorage.getItem(key); return v ? { ...fallback, ...JSON.parse(v) } : fallback; } catch { return fallback; }
};
const write = (key: string, v: unknown) => { try { localStorage.setItem(key, JSON.stringify(v)); } catch { /* private mode */ } };

export const userOptions = (): MathUserOptions => read("raabe.math.options", { autoBuild: true, copyMathml: false });

/** "Standardwerte ...": the settings new documents start with (Word: the Normal template). */
export const userDefaults = (): MathSettings | null => {
  try { const v = localStorage.getItem("raabe.math.defaults"); return v ? { ...DEFAULT_MATH, ...JSON.parse(v) } : null; } catch { return null; }
};

const cm = (pt: number) => `${(pt / CM).toLocaleString("de-DE", { maximumFractionDigits: 2 })} cm`;
const pts = (v: string, fallback: number) => { const n = parseNum(v); return Number.isFinite(n) ? Math.max(0, n * CM) : fallback; };

function select<T extends string>(opts: [T, string][], value: T): HTMLSelectElement {
  const sel = h("select", null, ...opts.map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
  sel.value = value;
  return sel;
}

function radio(name: string, label: string, on: boolean): { input: HTMLInputElement; el: HTMLElement } {
  const input = h("input", { type: "radio", name, checked: on }) as HTMLInputElement;
  return { input, el: h("label", { class: "check indent" }, input, label) };
}

function check(label: string, on: boolean, cls = ""): { input: HTMLInputElement; el: HTMLElement } {
  const input = h("input", { type: "checkbox", checked: on }) as HTMLInputElement;
  return { input, el: h("label", { class: `check ${cls}` }, input, label) };
}

/** Math. AutoKorrektur and Erkannte Funktionen: the lists the formula editor uses. */
function listDialog(title: string, head: [string, string] | [string], rows: string[][]) {
  const table = h("table", { class: "mlist" },
    h("thead", null, h("tr", null, ...head.map((x) => h("th", null, x)))),
    h("tbody", null, ...rows.map((r) => h("tr", null, ...r.map((x) => h("td", null, x))))));
  return showDialog({ title, body: h("div", { class: "mlistbox" }, table), buttons: [{ label: "Schließen", value: undefined, primary: true }], cancel: undefined });
}

export async function mathOptionsDialog(e: Editor): Promise<void> {
  const m = e.mathSettings();
  const u = userOptions();
  const font = select([["Cambria Math", "Cambria Math"]], "Cambria Math");
  const auto = check("Ausdrücke automatisch in professionelles Format konvertieren", u.autoBuild);
  const small = check("Größe von verschachtelten Brüchen für Anzeigeformeln reduzieren", m.smallFrac);
  const copyMl = radio("mcopy", "MathML als Nur-Text in die Zwischenablage kopieren", u.copyMathml);
  const copyLin = radio("mcopy", "Lineares Format als Nur-Text in die Zwischenablage kopieren", !u.copyMathml);
  const brk = select([["before", "Vor"], ["after", "Nach"], ["repeat", "Doppelt"]], m.brkBin);
  const sub = select([["--", "--"], ["-+", "-+"], ["+-", "+-"]], m.brkBinSub);
  const subRow = h("label", { class: "field row indent" }, "Operatoren für die Subtraktion duplizieren als:", sub);
  const syncSub = () => { sub.disabled = brk.value !== "repeat"; subRow.classList.toggle("disabled", sub.disabled); };
  brk.onchange = syncSub;
  syncSub();
  const acButton = h("button", { type: "button", onclick: () => void listDialog("AutoKorrektur von Mathematik", ["Ersetzen:", "Durch:"], mathAutoCorrect()) }, "Math. Autokorrektur...");
  const fnButton = h("button", { type: "button", onclick: () => void listDialog("Erkannte mathematische Funktionen", ["Funktion"], [...FUNCTIONS, ...LIMIT_FUNCTIONS].sort().map((f) => [f])) }, "Erkannte Funktionen...");

  const intSide = radio("mint", "Integralgrenzwerte seitlich platzieren", m.intLim === "subSup");
  const intOver = radio("mint", "Integralgrenzwerte zentriert oberhalb und unterhalb platzieren", m.intLim === "undOvr");
  const narySide = radio("mnary", "N-äre Grenzwerte seitlich platzieren", m.naryLim === "subSup");
  const naryOver = radio("mnary", "N-äre Grenzwerte direkt oberhalb und unterhalb platzieren", m.naryLim === "undOvr");
  const disp = check("Die folgenden Einstellungen für Mathematik in einer eigenen Zeile verwenden", m.dispDef);
  const lm = h("input", { type: "text", value: cm(m.lMargin), style: "width:80px" }) as HTMLInputElement;
  const rm = h("input", { type: "text", value: cm(m.rMargin), style: "width:80px" }) as HTMLInputElement;
  const jc = select([["left", "Links"], ["right", "Rechts"], ["center", "Zentriert"], ["centerGroup", "Zentriert als Gruppe"]], m.defJc);
  const indent = radio("mwrap", "Einzug für umbrochene Zeilen:", !m.wrapRight);
  const right = radio("mwrap", "Umbrochene Zeilen rechtsbündig ausrichten", m.wrapRight);
  const wi = h("input", { type: "text", value: cm(m.wrapIndent), style: "width:80px" }) as HTMLInputElement;
  const own = h("div", { class: "indent" },
    h("div", { class: "grid2" }, "Linker Rand:", lm, "Rechter Rand:", rm),
    h("div", { class: "note" }, "Mathematische Ränder werden Absatzrändern hinzugefügt."),
    h("label", { class: "field row" }, "Ausrichtung:", jc),
    h("div", null, "Für Mathematik mit Umbruch in eine neue Zeile:"),
    h("div", { class: "row" }, indent.el, wi), right.el);
  const syncOwn = () => { for (const x of Array.from(own.querySelectorAll("input, select"))) (x as HTMLInputElement).disabled = !disp.input.checked; };
  disp.input.onchange = syncOwn;
  syncOwn();

  const collect = (): MathSettings => ({
    font: font.value, smallFrac: small.input.checked, brkBin: brk.value as MathSettings["brkBin"], brkBinSub: sub.value as MathSettings["brkBinSub"],
    dispDef: disp.input.checked, lMargin: pts(lm.value, m.lMargin), rMargin: pts(rm.value, m.rMargin), defJc: jc.value as MathSettings["defJc"],
    wrapIndent: pts(wi.value, m.wrapIndent), wrapRight: right.input.checked,
    intLim: intOver.input.checked ? "undOvr" : "subSup", naryLim: narySide.input.checked ? "subSup" : "undOvr",
  });

  const body = h("div", { class: "mathoptions", style: "width:560px" },
    h("fieldset", null, h("legend", null, "Allgemein"),
      h("label", { class: "field row" }, "Standardschriftart für mathematische Bereiche:", font),
      auto.el, small.el,
      h("div", null, "Beim Kopieren einer Formel:"), copyMl.el, copyLin.el,
      h("label", { class: "field row" }, "Zeilen mit binären und relationalen Operatoren umbrechen:", brk), subRow,
      h("div", { class: "row end" }, acButton, fnButton)),
    h("fieldset", null, h("legend", null, "Mathematik anzeigen"),
      h("div", null, "Beim Erstellen von Integralen"), intSide.el, intOver.el,
      h("div", null, "Beim Erstellen von anderen n-ären Operatoren"), narySide.el, naryOver.el,
      disp.el, own));
  for (;;) {
    const v = await showDialog<"ok" | "cancel" | "defaults">({
      title: "Formeloptionen", body,
      buttons: [{ label: "Standardwerte ...", value: "defaults" }, { label: "OK", value: "ok", primary: true }, { label: "Abbrechen", value: "cancel" }],
      cancel: "cancel",
    });
    if (v === "cancel") return;
    if (v === "defaults") {
      const yes = await showDialog({
        title: "RaabeWrite",
        body: "Möchten Sie die Standardformeleinstellungen ändern? Diese Änderung wirkt sich auf alle neuen Dokumente aus.",
        buttons: [{ label: "Ja", value: true, primary: true }, { label: "Nein", value: false }], cancel: false,
      });
      if (yes) write("raabe.math.defaults", collect());
      continue;
    }
    write("raabe.math.options", { autoBuild: auto.input.checked, copyMathml: copyMl.input.checked });
    const next = collect();
    if (JSON.stringify(next) !== JSON.stringify(m)) e.setMathSettings(next);
    return;
  }
}
