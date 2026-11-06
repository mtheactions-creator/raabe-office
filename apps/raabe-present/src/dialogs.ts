// PowerPoint's Schriftart and Absatz dialogs (D-070): the labels as PowerPoint's (recorded in the
// inventory), applied to the selection like the ribbon's buttons. Units in German notation (cm, pt).

import { h, showDialog } from "@raabe/shell";
import type { RunStyle, TextPara } from "@raabe/ooxml";
import * as T from "./textops";

export interface DialogHost {
  runStyle(): RunStyle | null;
  paraStyle(): TextPara | null;
  fonts(): string[];
  formatRuns(label: string, fn: (rPr: Element) => void): void;
  formatParas(label: string, fn: (pPr: Element) => void): void;
}

const num = (s: string) => Number(s.replace(",", ".").replace(/[^\d.\-]/g, ""));
const de = (n: number, d = 2) => (Math.round(n * 10 ** d) / 10 ** d).toString().replace(".", ",");

function select(options: string[], value: string): HTMLSelectElement {
  const s = h("select", null, ...options.map((o) => h("option", { value: o }, o))) as HTMLSelectElement;
  s.value = value;
  return s;
}

function row(label: string, ...els: HTMLElement[]): HTMLElement {
  return h("label", { class: "dlg-row" }, h("span", { class: "dlg-label" }, label), ...els);
}

export async function fontDialog(host: DialogHost): Promise<void> {
  const st = host.runStyle();
  const face = h("input", { type: "text", value: st?.font ?? "", list: "dlg-fonts", style: "width:200px" }) as HTMLInputElement;
  const list = h("datalist", { id: "dlg-fonts" }, ...host.fonts().map((f) => h("option", { value: f })));
  const style = select(["Standard", "Kursiv", "Fett", "Fett Kursiv"], st ? (st.bold && st.italic ? "Fett Kursiv" : st.bold ? "Fett" : st.italic ? "Kursiv" : "Standard") : "Standard");
  const size = h("input", { type: "text", value: st ? de(st.size, 1) : "", style: "width:60px" }) as HTMLInputElement;
  const color = h("input", { type: "color", value: `#${st?.color ?? "000000"}` }) as HTMLInputElement;
  const under = select(["(ohne)", "Einfach", "Doppelt", "Dick", "Gepunktet", "Gestrichelt", "Wellenlinie"], st?.underline === "sng" ? "Einfach" : st?.underline === "dbl" ? "Doppelt" : st?.underline === "heavy" ? "Dick" : st?.underline === "dotted" ? "Gepunktet" : st?.underline === "dash" ? "Gestrichelt" : st?.underline === "wavy" ? "Wellenlinie" : "(ohne)");
  const cb = (label: string, on: boolean) => { const c = h("input", { type: "checkbox" }) as HTMLInputElement; c.checked = on; return { c, el: h("label", { class: "dlg-check" }, c, ` ${label}`) }; };
  const strike = cb("Durchgestrichen", st?.strike === "sngStrike"), dstrike = cb("Doppelt durchgestrichen", st?.strike === "dblStrike");
  const sup = cb("Hochgestellt", (st?.baseline ?? 0) > 0), sub = cb("Tiefgestellt", (st?.baseline ?? 0) < 0);
  const offset = h("input", { type: "text", value: `${de(Math.abs((st?.baseline ?? 0) * 100), 0)} %`, style: "width:60px" }) as HTMLInputElement;
  const small = cb("Kapitälchen", st?.caps === "small"), all = cb("Großbuchstaben", st?.caps === "all"), eq = cb("Zeichenhöhe ausgleichen", false);
  const spacing = select(["Normal", "Erweitert", "Schmal"], (st?.spacing ?? 0) > 0 ? "Erweitert" : (st?.spacing ?? 0) < 0 ? "Schmal" : "Normal");
  const by = h("input", { type: "text", value: `${de(Math.abs(st?.spacing ?? 0), 1)} Pt.`, style: "width:70px" }) as HTMLInputElement;
  const pageA = h("div", { class: "dlg-page" },
    row("Schriftart für lateinischen Text:", face, list), row("Schriftschnitt:", style), row("Größe:", size),
    h("fieldset", null, h("legend", null, "Alle Textzeichen"), row("Schriftfarbe", color), row("Unterstreichungsart", under)),
    h("fieldset", null, h("legend", null, "Effekte"), strike.el, dstrike.el, sup.el, row("Offset:", offset), sub.el, small.el, all.el, eq.el));
  const pageB = h("div", { class: "dlg-page", hidden: true }, row("Abstand:", spacing), row("Um:", by));
  const tabs = h("div", { class: "dlg-tabs" });
  const tabA = h("button", { class: "on", onclick: () => { pageA.hidden = false; pageB.hidden = true; tabA.classList.add("on"); tabB.classList.remove("on"); } }, "Schriftart");
  const tabB = h("button", { onclick: () => { pageA.hidden = true; pageB.hidden = false; tabB.classList.add("on"); tabA.classList.remove("on"); } }, "Zeichenabstand");
  tabs.append(tabA, tabB);
  const ok = await showDialog({ title: "Schriftart", body: h("div", { class: "font-dialog" }, tabs, pageA, pageB), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  const uMap: Record<string, string> = { "(ohne)": "none", Einfach: "sng", Doppelt: "dbl", Dick: "heavy", Gepunktet: "dotted", Gestrichelt: "dash", Wellenlinie: "wavy" };
  const sz = num(size.value), off = num(offset.value), sp = num(by.value);
  host.formatRuns("Schriftart", (r) => {
    if (face.value) T.setRunFont(r, face.value);
    r.setAttribute("b", style.value.includes("Fett") ? "1" : "0");
    r.setAttribute("i", style.value.includes("Kursiv") ? "1" : "0");
    if (sz > 0) r.setAttribute("sz", String(Math.round(sz * 100)));
    T.setRunColor(r, color.value.replace("#", "").toUpperCase());
    r.setAttribute("u", uMap[under.value]);
    r.setAttribute("strike", dstrike.c.checked ? "dblStrike" : strike.c.checked ? "sngStrike" : "noStrike");
    r.setAttribute("baseline", String(sup.c.checked ? Math.round((off || 30) * 1000) : sub.c.checked ? -Math.round((off || 25) * 1000) : 0));
    r.setAttribute("cap", all.c.checked ? "all" : small.c.checked ? "small" : "none");
    if (spacing.value === "Normal") r.removeAttribute("spc"); else r.setAttribute("spc", String(Math.round((spacing.value === "Schmal" ? -1 : 1) * (sp || 1) * 100)));
  });
}

export async function paragraphDialog(host: DialogHost): Promise<void> {
  const p = host.paraStyle();
  const cm = (emu: number) => `${de(emu / 360000)} cm`;
  const align = select(["Links", "Zentriert", "Rechts", "Blocksatz", "Verteilt"], p ? { l: "Links", ctr: "Zentriert", r: "Rechts", just: "Blocksatz", dist: "Verteilt" }[p.algn] : "Links");
  const before = h("input", { type: "text", value: cm(p ? p.marL + Math.min(0, p.indent) : 0), style: "width:80px" }) as HTMLInputElement;
  const special = select(["(ohne)", "Erste Zeile", "Hängend"], !p || !p.indent ? "(ohne)" : p.indent > 0 ? "Erste Zeile" : "Hängend");
  const by = h("input", { type: "text", value: cm(Math.abs(p?.indent ?? 0)), style: "width:80px" }) as HTMLInputElement;
  const spB = h("input", { type: "text", value: `${de(p?.spcBefPts ?? 0, 1)} Pt.`, style: "width:80px" }) as HTMLInputElement;
  const spA = h("input", { type: "text", value: `${de(p?.spcAftPts ?? 0, 1)} Pt.`, style: "width:80px" }) as HTMLInputElement;
  const pct = p?.lnPct ?? 1;
  const kind = select(["Einfach", "1,5 Zeilen", "Doppelt", "Genau", "Mehrfach"], p?.lnPts !== undefined ? "Genau" : pct === 1 ? "Einfach" : pct === 1.5 ? "1,5 Zeilen" : pct === 2 ? "Doppelt" : "Mehrfach");
  const at = h("input", { type: "text", value: p?.lnPts !== undefined ? `${de(p.lnPts, 1)} Pt.` : de(pct), style: "width:80px" }) as HTMLInputElement;
  const body = h("div", { class: "para-dialog" },
    h("div", { class: "dlg-tabs" }, h("button", { class: "on" }, "Einzüge und Abstände"), h("button", { disabled: true }, "Asiatische Typografie")),
    h("fieldset", null, h("legend", null, "Allgemein"), row("Ausrichtung:", align)),
    h("fieldset", null, h("legend", null, "Einzug"), row("Vor Text:", before), row("Sondereinzug:", special), row("Um:", by)),
    h("fieldset", null, h("legend", null, "Abstand"), row("Vor:", spB), row("Nach:", spA), row("Zeilenabstand:", kind), row("Maß:", at)));
  const ok = await showDialog({ title: "Absatz", body, buttons: [{ label: "Tabstopps...", value: "tabs", left: true }, { label: "OK", value: "ok", primary: true }, { label: "Abbrechen", value: "cancel" }], cancel: "cancel" });
  if (ok !== "ok") return;
  const emu = (s: string) => Math.round(num(s) * 360000);
  host.formatParas("Absatz", (pPr) => {
    pPr.setAttribute("algn", { Links: "l", Zentriert: "ctr", Rechts: "r", Blocksatz: "just", Verteilt: "dist" }[align.value] ?? "l");
    const ind = special.value === "(ohne)" ? 0 : special.value === "Erste Zeile" ? emu(by.value) : -emu(by.value);
    pPr.setAttribute("marL", String(emu(before.value) - Math.min(0, ind)));
    pPr.setAttribute("indent", String(ind));
    const doc = pPr.ownerDocument;
    // pPr's children in schema order.
    const ORDER = ["lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont", "buNone", "buAutoNum", "buChar", "buBlip", "tabLst", "defRPr", "extLst"];
    const set = (name: string, inner: Element) => {
      for (const c of Array.from(pPr.children)) if (c.localName === name) c.remove();
      const el = doc.createElementNS(T.NS_A, `a:${name}`);
      el.append(inner);
      const after = Array.from(pPr.children).find((c) => ORDER.indexOf(c.localName) > ORDER.indexOf(name));
      pPr.insertBefore(el, after ?? null);
    };
    const pts = (v: number) => { const e = doc.createElementNS(T.NS_A, "a:spcPts"); e.setAttribute("val", String(Math.round(v * 100))); return e; };
    const pctEl = (v: number) => { const e = doc.createElementNS(T.NS_A, "a:spcPct"); e.setAttribute("val", String(Math.round(v * 100000))); return e; };
    const lk = kind.value;
    set("lnSpc", lk === "Genau" ? pts(num(at.value)) : pctEl(lk === "Einfach" ? 1 : lk === "1,5 Zeilen" ? 1.5 : lk === "Doppelt" ? 2 : num(at.value) || 1));
    set("spcBef", pts(num(spB.value)));
    set("spcAft", pts(num(spA.value)));
  });
}
