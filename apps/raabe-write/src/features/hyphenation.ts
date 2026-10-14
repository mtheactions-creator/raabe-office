// Layout > Silbentrennung as in Word (recorded 2026-10-14): Keine, Automatisch, Manuell and
// Silbentrennungsoptionen... (Automatische Silbentrennung, Wörter in Großbuchstaben trennen,
// Silbentrennzone greyed as Word 2013 and later have it, Aufeinanderfolgende Trennstriche,
// Manuell...). Manuelle Silbentrennung goes through the words automatic hyphenation would
// divide at a line end and proposes the place ("mit-hil-fe"); Ja puts an optional hyphen there.

import { h, messageBox, showDialog, type Command, type MenuEntry } from "@raabe/shell";
import { loadHyphenation, type Editor } from "@raabe/text-engine";
import type { AppApi } from "./api";
import { languageName } from "./languages";

export const HYPHENATION_COMMANDS: Command[] = [
  { id: "layout.hyphenationNone", label: "Keine" },
  { id: "layout.hyphenationAuto", label: "Automatisch" },
  { id: "layout.hyphenationManual", label: "Manuell" },
  { id: "layout.hyphenationOptions", label: "Silbentrennungsoptionen..." },
];

export const HYPHENATION_MENU: MenuEntry[] = [
  { cmd: "layout.hyphenationNone" }, { cmd: "layout.hyphenationAuto" }, { cmd: "layout.hyphenationManual" }, "-", { cmd: "layout.hyphenationOptions" },
];

const SOFT_HYPHEN = String.fromCharCode(0xad);

export function hyphenationFeature(api: AppApi) {
  const ed = () => api.editor();
  const c = api.commands;
  const set = (auto: boolean) => { const e = ed(); if (e) { e.setHyphenation({ ...e.hyphenation, auto }); api.refresh(); api.focus(); } };

  // ----- Manuelle Silbentrennung ------------------------------------------------------------------
  const manual = async () => {
    const e = ed();
    if (!e) return;
    // The patterns of the languages in the text (they load on first use).
    const langs = new Set([e.proofingAt().lang, e.doc.styles.defaultsR.lang ?? "de-DE"]);
    for (const p of e.doc.main) for (const r of p.runs) if (r.props.lang) langs.add(r.props.lang);
    await Promise.all([...langs].map((l) => loadHyphenation(l).catch(() => null)));
    let from = { p: 0, o: 0 };
    let asked = 0;
    for (;;) {
      const prop = e.hyphenationProposal(from);
      if (!prop) break;
      asked++;
      e.selectRange(prop.p, prop.a, prop.b);
      const choice = await proposalDialog(e, prop);
      if (choice === null) { api.focus(); return; }
      if (choice >= 0) {
        e.replaceRange(prop.p, prop.a + choice, prop.a + choice, SOFT_HYPHEN);
        from = { p: prop.p, o: prop.b + 1 };
      } else from = { p: prop.p, o: prop.b };
    }
    await messageBox("RaabeWrite", asked ? "Die Silbentrennung ist abgeschlossen." : "Es wurden keine Wörter gefunden, die getrennt werden können.");
    api.focus();
  };

  /** The dialog for one word: its hyphenation points, the proposal marked; Ja, Nein, Abbrechen. */
  const proposalDialog = (e: Editor, prop: NonNullable<ReturnType<Editor["hyphenationProposal"]>>): Promise<number | null> => {
    let at = prop.points.includes(prop.at) ? prop.at : prop.points[prop.points.length - 1] ?? prop.at;
    const field = h("div", { class: "hy-word", tabIndex: 0, role: "textbox", "aria-label": "Trennvorschlag:" });
    const paint = () => {
      field.replaceChildren();
      let last = 0;
      for (const p of [...prop.points, prop.word.length]) {
        field.append(h("span", null, prop.word.slice(last, p)));
        if (p < prop.word.length) {
          const mark = h("span", { class: `hy-point${p === at ? " on" : ""}${p > prop.at ? " off" : ""}`, onclick: () => { if (p <= prop.at) { at = p; paint(); } } }, "-");
          field.append(mark);
        }
        last = p;
      }
    };
    field.addEventListener("keydown", (ev) => {
      const i = prop.points.indexOf(at);
      if (ev.key === "ArrowLeft" && i > 0) { at = prop.points[i - 1]; paint(); ev.preventDefault(); }
      // Only places up to the proposal fit on the line.
      if (ev.key === "ArrowRight" && i < prop.points.length - 1 && prop.points[i + 1] <= prop.at) { at = prop.points[i + 1]; paint(); ev.preventDefault(); }
    });
    paint();
    const lang = e.proofingAt().lang;
    return showDialog<"yes" | "no" | null>({
      title: `Manuelle Silbentrennung: ${languageName(lang)}`,
      body: h("div", { class: "hydlg" }, h("label", { class: "row" }, h("span", null, "Trennvorschlag:"), field)),
      buttons: [{ label: "Ja", value: "yes", primary: true }, { label: "Nein", value: "no" }, { label: "Abbrechen", value: null }],
      cancel: null,
    }).then((v) => (v === "yes" ? at : v === "no" ? -1 : null));
  };

  // ----- Silbentrennungsoptionen ---------------------------------------------------------------
  const optionsDialog = async () => {
    const e = ed();
    if (!e) return;
    const cur = e.hyphenation;
    const auto = h("input", { type: "checkbox", checked: cur.auto }) as HTMLInputElement;
    const caps = h("input", { type: "checkbox", checked: cur.caps }) as HTMLInputElement;
    const zone = h("input", { type: "text", value: `${((cur.zone ?? 21.25) / 28.3465).toLocaleString("de-DE", { maximumFractionDigits: 2 })} cm`, disabled: true, style: "width:90px" }) as HTMLInputElement;
    const limit = h("input", { type: "text", value: cur.limit ? String(cur.limit) : "Unbegrenzt", style: "width:90px", list: "hy-limits" }) as HTMLInputElement;
    const limits = h("datalist", { id: "hy-limits" }, h("option", { value: "Unbegrenzt" }), ...[1, 2, 3, 4, 5].map((n) => h("option", { value: String(n) })));
    let goManual = false;
    const manualBtn = h("button", { onclick: () => { goManual = true; (document.querySelector(".dialog-layer .dlg-buttons button.primary") as HTMLButtonElement | null)?.click(); } }, "Manuell...");
    const ok = await showDialog({
      title: "Silbentrennung",
      body: h("div", { class: "hyopts" },
        h("label", { class: "check" }, auto, "Automatische Silbentrennung"),
        h("label", { class: "check" }, caps, "Wörter in Großbuchstaben trennen"),
        h("div", { class: "grid2" }, h("span", { class: "muted" }, "Silbentrennzone:"), zone, h("span", null, "Aufeinanderfolgende Trennstriche:"), limit, limits),
        h("div", null, manualBtn)),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    const n = Number.parseInt(limit.value, 10);
    e.setHyphenation({ ...cur, auto: auto.checked, caps: caps.checked, limit: Number.isFinite(n) && n > 0 ? n : 0 });
    api.refresh();
    if (goManual) await manual();
    else api.focus();
  };

  c.implement("layout.hyphenation", { run: () => {}, enabled: () => !!ed() });
  c.implement("layout.hyphenationNone", { run: () => set(false), enabled: () => !!ed(), checked: () => !ed()?.hyphenation.auto });
  c.implement("layout.hyphenationAuto", { run: () => set(true), enabled: () => !!ed(), checked: () => !!ed()?.hyphenation.auto });
  c.implement("layout.hyphenationManual", { run: () => void manual(), enabled: () => !!ed() });
  c.implement("layout.hyphenationOptions", { run: () => void optionsDialog(), enabled: () => !!ed() });
}
