// Layout > Seite einrichten: Seitenränder, Ausrichtung, Format, Spalten, Umbrüche and the
// "Seite einrichten" dialog. Changes apply to the section with the caret (with one section:
// the whole document), as in Word.

import type { Columns, LineNumbers, PageSetup, Section } from "@raabe/text-engine";
import { h, showDialog, type Command, type MenuEntry } from "@raabe/shell";
import { CM, fmtCm, parseNum, type AppApi } from "./api";
import { openModal } from "./modal";

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const PAGE_COMMANDS: Command[] = [
  cmd("insert.columnBreak", "Spalte", "breaks", { shortcut: "Strg+Umschalt+Eingabe", keywords: ["Spaltenumbruch"] }),
  cmd("section.nextPage", "Nächste Seite", "section-break", { keywords: ["Abschnittsumbruch", "Abschnittswechsel"] }),
  cmd("section.continuous", "Fortlaufend", "section-break", { keywords: ["Abschnittsumbruch", "Abschnittswechsel"] }),
  cmd("section.evenPage", "Gerade Seite", "section-break", { keywords: ["Abschnittsumbruch"] }),
  cmd("section.oddPage", "Ungerade Seite", "section-break", { keywords: ["Abschnittsumbruch"] }),
  cmd("orientation.portrait", "Hochformat", "portrait"),
  cmd("orientation.landscape", "Querformat", "landscape"),
  cmd("columns.more", "Weitere Spalten...", "columns"),
  cmd("margins.custom", "Benutzerdefinierte Seitenränder...", "margins"),
  cmd("size.more", "Weitere Papierformate...", "page-size"),
];

const cm = (v: number) => v * CM;

export const MARGINS: { name: string; top: number; bottom: number; left: number; right: number }[] = [
  { name: "Normal", top: cm(2.5), bottom: cm(2), left: cm(2.5), right: cm(2.5) },
  { name: "Schmal", top: cm(1.27), bottom: cm(1.27), left: cm(1.27), right: cm(1.27) },
  { name: "Mittel", top: cm(2.54), bottom: cm(2.54), left: cm(1.91), right: cm(1.91) },
  { name: "Breit", top: cm(2.54), bottom: cm(2.54), left: cm(5.08), right: cm(5.08) },
];

export const SIZES: { name: string; w: number; h: number }[] = [
  { name: "A3", w: 841.9, h: 1190.55 },
  { name: "A4", w: 595.3, h: 841.9 },
  { name: "A5", w: 419.55, h: 595.3 },
  { name: "B5 (JIS)", w: 515.95, h: 728.5 },
  { name: "Letter", w: 612, h: 792 },
  { name: "Legal", w: 612, h: 1008 },
];

/** Word rotates the margins with the page (measured: landscape top = portrait right …). */
export function rotate(p: PageSetup, landscape: boolean): PageSetup {
  const isLandscape = p.width > p.height;
  if (isLandscape === landscape) return p;
  return landscape
    ? { ...p, width: p.height, height: p.width, marginTop: p.marginRight, marginBottom: p.marginLeft, marginLeft: p.marginTop, marginRight: p.marginBottom }
    : { ...p, width: p.height, height: p.width, marginTop: p.marginLeft, marginBottom: p.marginRight, marginLeft: p.marginBottom, marginRight: p.marginTop };
}

const near = (a: number, b: number) => Math.abs(a - b) < 0.6;

export function pageSetupFeature(api: AppApi) {
  const ed = () => api.editor();
  const c = api.commands;
  const has = () => !!ed();
  const page = () => ed()?.section().page;
  const change = (fn: (s: Section) => Section) => { ed()?.changeSections(fn, false); api.focus(); };

  c.implement("layout.margins", { run: () => {}, enabled: has });
  c.implement("layout.orientation", { run: () => {}, enabled: has });
  c.implement("layout.size", { run: () => {}, enabled: has });
  c.implement("layout.columns", { run: () => {}, enabled: has });
  c.implement("layout.breaks", { run: () => {}, enabled: has });
  c.implement("orientation.portrait", { run: () => change((s) => ({ ...s, page: rotate(s.page, false) })), enabled: has, checked: () => { const p = page(); return !!p && p.width <= p.height; } });
  c.implement("orientation.landscape", { run: () => change((s) => ({ ...s, page: rotate(s.page, true) })), enabled: has, checked: () => { const p = page(); return !!p && p.width > p.height; } });
  const mainOnly = () => !!ed() && ed()!.doc.activeStory === "main";
  c.implement("insert.columnBreak", { run: () => { ed()?.insertBreak("column"); api.focus(); }, enabled: mainOnly });
  for (const [id, start] of [["section.nextPage", "nextPage"], ["section.continuous", "continuous"], ["section.evenPage", "evenPage"], ["section.oddPage", "oddPage"]] as const) {
    c.implement(id, { run: () => { ed()?.insertSectionBreak(start); api.focus(); }, enabled: mainOnly });
  }
  c.implement("columns.more", { run: () => void columnsDialog(), enabled: has });
  c.implement("margins.custom", { run: () => void pageDialog("margins"), enabled: has });
  c.implement("size.more", { run: () => void pageDialog("paper"), enabled: has });
  c.implement("layout.pageSetup", { run: () => void pageDialog("margins"), enabled: has });

  const setColumns = (count: number, cols?: Columns["cols"]) => change((s) => ({ ...s, columns: { ...s.columns, count, cols } }));

  // ----- galleries ------------------------------------------------------------------------
  const marginsPanel = (close: () => void) => {
    const p = page();
    const list = MARGINS.map((m) => h("button", {
      class: p && near(p.marginTop, m.top) && near(p.marginBottom, m.bottom) && near(p.marginLeft, m.left) && near(p.marginRight, m.right) ? "checked" : "",
      onclick: () => { close(); change((s) => ({ ...s, page: { ...s.page, marginTop: m.top, marginBottom: m.bottom, marginLeft: m.left, marginRight: m.right } })); },
    }, h("span", { class: "ico-l", style: "width:32px;height:40px;border:1px solid #9e9e9e;background:#fff;position:relative;flex:none" },
      h("span", { style: `position:absolute;left:${m.left / 20}px;right:${m.right / 20}px;top:${m.top / 20}px;bottom:${m.bottom / 20}px;border:1px dashed #9e9e9e` })),
    h("span", null, h("b", null, m.name), h("div", { class: "desc" },
      `Oben: ${fmtCm(m.top)}\tUnten: ${fmtCm(m.bottom)}\nLinks: ${fmtCm(m.left)}\tRechts: ${fmtCm(m.right)}`))));
    return h("div", { class: "rpanel" }, h("div", { class: "wide" }, ...list), h("hr"),
      h("button", { class: "link", onclick: () => { close(); void pageDialog("margins"); } }, "Benutzerdefinierte Seitenränder..."));
  };

  const sizePanel = (close: () => void) => {
    const p = page();
    const list = SIZES.map((sz) => h("button", {
      class: p && ((near(p.width, sz.w) && near(p.height, sz.h)) || (near(p.width, sz.h) && near(p.height, sz.w))) ? "checked" : "",
      onclick: () => {
        close();
        change((s) => {
          const land = s.page.width > s.page.height;
          return { ...s, page: { ...s.page, width: land ? sz.h : sz.w, height: land ? sz.w : sz.h } };
        });
      },
    }, h("span", null, h("b", null, sz.name), h("div", { class: "desc" }, `${fmtCm(sz.w)} x ${fmtCm(sz.h)}`))));
    return h("div", { class: "rpanel" }, h("div", { class: "wide" }, ...list), h("hr"),
      h("button", { class: "link", onclick: () => { close(); void pageDialog("paper"); } }, "Weitere Papierformate..."));
  };

  const columnsMenu: MenuEntry[] = [
    { label: "Eins", icon: "col-one", run: () => setColumns(1) },
    { label: "Zwei", icon: "col-two", run: () => setColumns(2) },
    { label: "Drei", icon: "col-three", run: () => setColumns(3) },
    { label: "Links", icon: "col-left", run: () => { const w = textWidth(); setColumns(2, [{ width: (w - 35.4) / 3, space: 35.4 }, { width: (w - 35.4) * 2 / 3, space: 0 }]); } },
    { label: "Rechts", icon: "col-right", run: () => { const w = textWidth(); setColumns(2, [{ width: (w - 35.4) * 2 / 3, space: 35.4 }, { width: (w - 35.4) / 3, space: 0 }]); } },
    "-", { cmd: "columns.more" },
  ];
  const textWidth = () => { const p = page(); return p ? p.width - p.marginLeft - p.marginRight - p.gutter : 453.6; };

  const breaksMenu: MenuEntry[] = [
    { heading: "Seitenumbrüche" }, { cmd: "insert.pageBreak", label: "Seite" }, { cmd: "insert.columnBreak" }, { cmd: "insert.lineBreak" },
    { heading: "Abschnittsumbrüche" }, { cmd: "section.nextPage" }, { cmd: "section.continuous" }, { cmd: "section.evenPage" }, { cmd: "section.oddPage" },
  ];

  // ----- dialogs --------------------------------------------------------------------------
  const columnsDialog = async () => {
    const s = ed()?.section();
    if (!s) return;
    const count = h("input", { type: "number", min: "1", max: "6", value: String(s.columns.count), style: "width:70px" });
    const space = h("input", { type: "text", value: fmtCm(s.columns.space), style: "width:90px" });
    const sep = h("input", { type: "checkbox", checked: s.columns.sep });
    const ok = await showDialog({
      title: "Spalten", body: h("div", { style: "width:300px" }, h("div", { class: "grid2" }, "Anzahl der Spalten:", count, "Abstand:", space),
        h("label", { class: "check" }, sep, "Zwischenlinie")),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    const n = Math.max(1, Math.min(6, Number(count.value) || 1));
    const sp = parseNum(space.value) * CM;
    change((x) => ({ ...x, columns: { count: n, space: Number.isFinite(sp) && sp >= 0 ? sp : 35.4, sep: sep.checked } }));
  };

  const pageDialog = async (first: "margins" | "paper" | "layout") => {
    const e = ed();
    if (!e) return;
    const s = e.section();
    const p = s.page;
    const input = (v: number) => h("input", { type: "text", value: fmtCm(v), style: "width:90px" });
    const top = input(p.marginTop), bottom = input(p.marginBottom), left = input(p.marginLeft), right = input(p.marginRight), gutter = input(p.gutter);
    const portrait = h("input", { type: "radio", name: "orient", checked: p.width <= p.height });
    const landscape = h("input", { type: "radio", name: "orient", checked: p.width > p.height });
    const width = input(p.width), height = input(p.height);
    const sizeSel = h("select", null, ...SIZES.map((x) => h("option", { value: x.name }, x.name)), h("option", { value: "custom" }, "Benutzerdefiniertes Format"));
    const match = SIZES.find((x) => (near(p.width, x.w) && near(p.height, x.h)) || (near(p.width, x.h) && near(p.height, x.w)));
    sizeSel.value = match?.name ?? "custom";
    sizeSel.onchange = () => {
      const x = SIZES.find((y) => y.name === sizeSel.value);
      if (!x) return;
      const land = landscape.checked;
      width.value = fmtCm(land ? x.h : x.w); height.value = fmtCm(land ? x.w : x.h);
    };
    const hdr = input(p.header), ftr = input(p.footer);
    const titlePage = h("input", { type: "checkbox", checked: s.titlePage });
    const evenOdd = h("input", { type: "checkbox", checked: !!e.doc.settings.evenOdd });
    const applyTo = h("select", null, h("option", { value: "section" }, e.sectionCount() > 1 ? "Aktuellen Abschnitt" : "Gesamtes Dokument"), h("option", { value: "all" }, "Gesamtes Dokument"));
    if (e.sectionCount() <= 1) applyTo.disabled = true;
    const pages: Record<string, HTMLElement> = {
      margins: h("div", null,
        h("fieldset", null, h("legend", null, "Seitenränder"), h("div", { class: "grid2" }, "Oben:", top, "Unten:", bottom, "Links:", left, "Rechts:", right, "Bundsteg:", gutter)),
        h("fieldset", null, h("legend", null, "Ausrichtung"), h("label", { class: "check" }, portrait, "Hochformat"), h("label", { class: "check" }, landscape, "Querformat"))),
      paper: h("div", null, h("fieldset", null, h("legend", null, "Papierformat"), sizeSel, h("div", { class: "grid2", style: "margin-top:8px" }, "Breite:", width, "Höhe:", height))),
      layout: h("div", null,
        h("fieldset", null, h("legend", null, "Kopf- und Fußzeilen"), h("label", { class: "check" }, evenOdd, "Gerade/ungerade anders"), h("label", { class: "check" }, titlePage, "Erste Seite anders"),
          h("div", { class: "grid2", style: "margin-top:8px" }, "Kopfzeile vom Seitenrand:", hdr, "Fußzeile vom Seitenrand:", ftr)),
        h("fieldset", null, h("legend", null, "Vorschau"),
          h("div", { class: "row", style: "gap:8px" },
            h("button", { onclick: () => lineNumbersDialog() }, "Zeilennummern..."),
            h("button", { onclick: () => void api.commands.run("design.pageBorders") }, "Rahmen...")))),
    };
    const tabs = h("div", { class: "tabs" });
    const show = (k: string) => {
      for (const [name, el] of Object.entries(pages)) el.hidden = name !== k;
      tabs.querySelectorAll("button").forEach((b) => b.classList.toggle("active", b.dataset.k === k));
    };
    for (const [k, label] of [["margins", "Seitenränder"], ["paper", "Papier"], ["layout", "Layout"]]) {
      tabs.append(h("button", { "data-k": k, onclick: () => show(k) }, label));
    }
    show(first);
    // Changing the orientation in the dialog swaps the paper size fields.
    const swap = () => { const w = width.value; width.value = height.value; height.value = w; };
    portrait.onchange = () => { if (parseNum(width.value) > parseNum(height.value)) swap(); };
    landscape.onchange = () => { if (parseNum(width.value) < parseNum(height.value)) swap(); };
    const body = h("div", { style: "width:420px" }, tabs, ...Object.values(pages), h("div", { class: "grid2", style: "margin-top:8px" }, "Übernehmen für:", applyTo));
    const ok = await showDialog({ title: "Seite einrichten", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const v = (el: HTMLInputElement, fallback: number) => { const n = parseNum(el.value) * CM; return Number.isFinite(n) && n >= 0 ? n : fallback; };
    let next: PageSetup = {
      ...p, marginTop: v(top, p.marginTop), marginBottom: v(bottom, p.marginBottom), marginLeft: v(left, p.marginLeft), marginRight: v(right, p.marginRight),
      gutter: v(gutter, p.gutter), width: v(width, p.width), height: v(height, p.height), header: v(hdr, p.header), footer: v(ftr, p.footer),
    };
    if (landscape.checked !== next.width > next.height) next = { ...next, width: next.height, height: next.width };
    e.changeSections((x) => ({ ...x, page: next, titlePage: titlePage.checked }), applyTo.value === "all");
    if (evenOdd.checked !== !!e.doc.settings.evenOdd) e.setEvenOdd(evenOdd.checked);
    api.focus();
  };

  // ----- Zeilennummern (Word's menu, recorded 2026-10-14, and dialog) -----------------------------
  const setLineNumbers = (ln: LineNumbers | undefined) => {
    const e = ed();
    if (!e) return;
    e.changeSections((s) => ({ ...s, lineNumbers: ln }), false);
    api.focus();
  };
  const current = () => ed()?.section().lineNumbers;
  const withRestart = (restart: LineNumbers["restart"]): LineNumbers => ({ countBy: 1, start: 1, ...(current() ?? {}), restart });
  const lineNumbersMenu: MenuEntry[] = [
    { label: "Keine", get checked() { return !current(); }, run: () => setLineNumbers(undefined) },
    { label: "Fortlaufend", get checked() { return current()?.restart === "continuous"; }, run: () => setLineNumbers(withRestart("continuous")) },
    { label: "Jede Seite neu beginnen", get checked() { return current()?.restart === "newPage"; }, run: () => setLineNumbers(withRestart("newPage")) },
    { label: "Jeden Abschnitt neu beginnen", get checked() { return current()?.restart === "newSection"; }, run: () => setLineNumbers(withRestart("newSection")) },
    { label: "Für aktuellen Absatz unterdrücken", get checked() { return !!ed()?.paraState().suppressLineNumbers; },
      run: () => { const e = ed(); if (e) { e.setParaProps({ suppressLineNumbers: !e.paraState().suppressLineNumbers || undefined }); api.focus(); } } },
    "-",
    { label: "Zeilennummerierungsoptionen...", run: () => void pageDialog("layout") },
  ];
  c.implement("layout.lineNumbers", { run: () => {}, enabled: has });

  const lineNumbersDialog = () => {
    const ln = current();
    const add = h("input", { type: "checkbox", checked: !!ln }) as HTMLInputElement;
    const start = h("input", { type: "number", min: "1", max: "32767", value: String(ln?.start ?? 1), style: "width:80px" }) as HTMLInputElement;
    const dist = h("input", { type: "text", value: ln?.distance !== undefined ? fmtCm(ln.distance) : "Auto", style: "width:80px" }) as HTMLInputElement;
    const by = h("input", { type: "number", min: "1", max: "100", value: String(ln?.countBy ?? 1), style: "width:80px" }) as HTMLInputElement;
    const radio = (label: string, v: LineNumbers["restart"]) => {
      const r = h("input", { type: "radio", name: "lnrestart", checked: (ln?.restart ?? "newPage") === v, "data-v": v }) as HTMLInputElement;
      return { r, el: h("label", { class: "check" }, r, label) };
    };
    const modes = [radio("Jede Seite neu beginnen", "newPage"), radio("Jeden Abschnitt neu beginnen", "newSection"), radio("Fortlaufend", "continuous")];
    const fields = [start, dist, by, ...modes.map((m) => m.r)];
    const sync = () => { for (const f of fields) f.disabled = !add.checked; };
    add.onchange = sync;
    sync();
    openModal("Zeilennummern", h("div", { class: "lndlg" },
      h("label", { class: "check" }, add, "Zeilennummern hinzufügen"),
      h("div", { class: "grid2" }, "Beginnen mit:", start, "Abstand vom Text:", dist, "Zählintervall:", by),
      h("fieldset", null, h("legend", null, "Nummerierung"), ...modes.map((m) => m.el))), [
      { label: "OK", primary: true, run: () => {
        if (!add.checked) { setLineNumbers(undefined); return true; }
        const d = dist.value.trim().toLowerCase() === "auto" || !dist.value.trim() ? undefined : parseNum(dist.value) * CM;
        const restart = (modes.find((m) => m.r.checked)?.r.dataset.v ?? "newPage") as LineNumbers["restart"];
        setLineNumbers({ countBy: Math.max(1, Math.round(Number(by.value) || 1)), start: Math.max(1, Math.round(Number(start.value) || 1)),
          distance: d !== undefined && Number.isFinite(d) ? d : undefined, restart });
        return true;
      } },
      { label: "Abbrechen", run: () => true },
    ], { width: 380 });
  };

  return { marginsPanel, sizePanel, columnsMenu, breaksMenu, lineNumbersMenu };
}
