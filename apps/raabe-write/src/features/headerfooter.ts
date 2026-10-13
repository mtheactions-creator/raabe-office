// Kopf- und Fußzeilen and Seitenzahlen: the galleries on Einfügen and the contextual tab
// "Kopf- und Fußzeile" shown while a header or footer is edited.

import { h, showDialog, type Command, type Control, type MenuEntry, type Tab } from "@raabe/shell";
import { CM, fmtCm, parseNum, type AppApi } from "./api";

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const HEADER_COMMANDS: Command[] = [
  cmd("header.blank", "Leer", "header"),
  cmd("header.blank3", "Leer (drei Spalten)", "header"),
  cmd("header.edit", "Kopfzeile bearbeiten", "header"),
  cmd("header.remove", "Kopfzeile entfernen", "header-remove"),
  cmd("footer.blank", "Leer", "footer"),
  cmd("footer.blank3", "Leer (drei Spalten)", "footer"),
  cmd("footer.edit", "Fußzeile bearbeiten", "footer"),
  cmd("footer.remove", "Fußzeile entfernen", "footer-remove"),
  cmd("pageNumber.topLeft", "Einfache Zahl 1", "page-top-left"),
  cmd("pageNumber.topCenter", "Einfache Zahl 2", "page-top-center"),
  cmd("pageNumber.topRight", "Einfache Zahl 3", "page-top-right"),
  cmd("pageNumber.bottomLeft", "Einfache Zahl 1", "page-bottom-left"),
  cmd("pageNumber.bottomCenter", "Einfache Zahl 2", "page-bottom-center"),
  cmd("pageNumber.bottomRight", "Einfache Zahl 3", "page-bottom-right"),
  cmd("pageNumber.ofPages", "Seite X von Y", "page-bottom-center", { keywords: ["Seitenzahl", "Seitenanzahl"] }),
  cmd("pageNumber.here", "Aktuelle Position", "page-number"),
  cmd("pageNumber.format", "Seitenzahlen formatieren...", "page-number"),
  cmd("pageNumber.remove", "Seitenzahlen entfernen", "remove"),
  cmd("hf.close", "Kopf- und Fußzeile schließen", "close-x", { tip: "Kopf- und Fußzeile schließen\nZurück zum Text des Dokuments (oder Doppelklick in den Text)." }),
  cmd("hf.gotoHeader", "Zur Kopfzeile wechseln", "go-header"),
  cmd("hf.gotoFooter", "Zur Fußzeile wechseln", "go-footer"),
  cmd("hf.firstPage", "Erste Seite anders", undefined),
  cmd("hf.evenOdd", "Untersch. gerade ungerade Seiten", undefined, { tip: "Unterschiedliche gerade und ungerade Seiten\nAndere Kopf- und Fußzeilen für gerade und ungerade Seiten." }),
  cmd("hf.headerDistance", "Abstand der Kopfzeile von oben", undefined),
  cmd("hf.footerDistance", "Abstand der Fußzeile von unten", undefined),
];

const FIELD = (instr: string, display: string) => ({ field: instr, display });

export function headerFooterFeature(api: AppApi) {
  const ed = () => api.editor();
  const c = api.commands;
  const has = () => !!ed();
  const inHf = () => !!ed()?.isHeaderFooter();
  const run = (id: string, fn: () => void, enabled = has, checked?: () => boolean) => c.implement(id, { run: () => { fn(); api.focus(); }, enabled, checked });

  const fill = (kind: "header" | "footer", align: "left" | "center" | "right", parts: (string | { field: string; display: string })[]) => {
    ed()?.setHeaderFooterContent(kind, align, parts);
    api.selectTab("headerFooter");
  };
  c.implement("insert.header", { run: () => {}, enabled: has });
  c.implement("insert.footer", { run: () => {}, enabled: has });
  c.implement("insert.pageNumber", { run: () => {}, enabled: has });
  run("header.blank", () => fill("header", "left", [""]));
  run("header.blank3", () => fill("header", "left", ["[Hier eingeben]\t[Hier eingeben]\t[Hier eingeben]"]));
  run("footer.blank", () => fill("footer", "left", [""]));
  run("footer.blank3", () => fill("footer", "left", ["[Hier eingeben]\t[Hier eingeben]\t[Hier eingeben]"]));
  run("header.edit", () => { ed()?.editHeaderFooter("header"); api.selectTab("headerFooter"); });
  run("footer.edit", () => { ed()?.editHeaderFooter("footer"); api.selectTab("headerFooter"); });
  run("header.remove", () => ed()?.removeHeaderFooter("header"));
  run("footer.remove", () => ed()?.removeHeaderFooter("footer"));
  for (const [id, kind, align] of [["pageNumber.topLeft", "header", "left"], ["pageNumber.topCenter", "header", "center"], ["pageNumber.topRight", "header", "right"],
    ["pageNumber.bottomLeft", "footer", "left"], ["pageNumber.bottomCenter", "footer", "center"], ["pageNumber.bottomRight", "footer", "right"]] as const) {
    run(id, () => fill(kind, align, [FIELD("PAGE", "1")]));
  }
  run("pageNumber.ofPages", () => fill("footer", "center", ["Seite ", FIELD("PAGE", "1"), " von ", FIELD("NUMPAGES", "1")]));
  run("pageNumber.here", () => ed()?.insertField("PAGE", "1"));
  run("pageNumber.remove", () => ed()?.removePageNumbers());
  run("pageNumber.format", () => void formatDialog());
  run("hf.close", () => { ed()?.closeHeaderFooter(); api.selectTab("start"); }, inHf);
  run("hf.gotoHeader", () => ed()?.editHeaderFooter("header"), inHf);
  run("hf.gotoFooter", () => ed()?.editHeaderFooter("footer"), inHf);
  run("hf.firstPage", () => { const e = ed(); if (e) e.setTitlePage(!e.section().titlePage); }, has, () => !!ed()?.section().titlePage);
  run("hf.evenOdd", () => { const e = ed(); if (e) e.setEvenOdd(!e.doc.settings.evenOdd); }, has, () => !!ed()?.doc.settings.evenOdd);
  c.implement("hf.headerDistance", { run: () => {}, enabled: inHf });
  c.implement("hf.footerDistance", { run: () => {}, enabled: inHf });

  const formatDialog = async () => {
    const e = ed();
    if (!e) return;
    const s = e.section();
    const fmt = h("select", null, ...[["decimal", "1, 2, 3, ..."], ["lowerLetter", "a, b, c, ..."], ["upperLetter", "A, B, C, ..."],
      ["lowerRoman", "i, ii, iii, ..."], ["upperRoman", "I, II, III, ..."]].map(([v, l]) => h("option", { value: v }, l)));
    fmt.value = s.pageFormat ?? "decimal";
    const cont = h("input", { type: "radio", name: "pn", checked: s.pageStart === undefined });
    const startAt = h("input", { type: "radio", name: "pn", checked: s.pageStart !== undefined });
    const start = h("input", { type: "number", min: "0", value: String(s.pageStart ?? 1), style: "width:70px" });
    const ok = await showDialog({
      title: "Seitenzahlenformat",
      body: h("div", { style: "width:320px" }, h("div", { class: "grid2" }, "Zahlenformat:", fmt),
        h("fieldset", null, h("legend", null, "Seitennummerierung"),
          h("label", { class: "check" }, cont, "Fortsetzen vom vorherigen Abschnitt"),
          h("label", { class: "check" }, startAt, "Beginnen bei:", start))),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    e.changeSections((x) => ({ ...x, pageFormat: fmt.value === "decimal" ? undefined : fmt.value, pageStart: startAt.checked ? Math.max(0, Number(start.value) || 1) : undefined }), false);
  };

  const headerMenu: MenuEntry[] = [{ heading: "Integriert" }, { cmd: "header.blank" }, { cmd: "header.blank3" }, "-", { cmd: "header.edit" }, { cmd: "header.remove" }];
  const footerMenu: MenuEntry[] = [{ heading: "Integriert" }, { cmd: "footer.blank" }, { cmd: "footer.blank3" }, "-", { cmd: "footer.edit" }, { cmd: "footer.remove" }];
  const pageNumberMenu: MenuEntry[] = [
    { heading: "Seitenanfang" }, { cmd: "pageNumber.topLeft" }, { cmd: "pageNumber.topCenter" }, { cmd: "pageNumber.topRight" },
    { heading: "Seitenende" }, { cmd: "pageNumber.bottomLeft" }, { cmd: "pageNumber.bottomCenter" }, { cmd: "pageNumber.bottomRight" }, { cmd: "pageNumber.ofPages" },
    "-", { cmd: "pageNumber.here" }, { cmd: "pageNumber.format" }, { cmd: "pageNumber.remove" },
  ];

  const page = () => ed()?.section().page;
  const distance = (which: "header" | "footer") => ({
    value: () => { const p = page(); return p ? fmtCm(which === "header" ? p.header : p.footer) : ""; },
    commit: (v: string) => { const n = parseNum(v) * CM; if (Number.isFinite(n) && n >= 0) ed()?.changeSections((s) => ({ ...s, page: { ...s.page, [which]: n } }), false); },
    step: (d: 1 | -1) => { const p = page(); if (!p) return; const cur = (which === "header" ? p.header : p.footer) / CM; const n = Math.max(0, Math.round((cur + d * 0.1) * 100) / 100) * CM; ed()?.changeSections((s) => ({ ...s, page: { ...s.page, [which]: n } }), false); },
  });
  const dd = (id: string, key: string, size: "large" | "medium", menu: MenuEntry[]): Control => ({ type: "dropdown", cmd: id, key, size, menu });
  const b = (id: string, key: string, size: "large" | "medium" | "small" = "small"): Control => ({ type: "button", cmd: id, key, size });
  const hd = distance("header"), fd = distance("footer");

  const tab: Tab = {
    id: "headerFooter", label: "Kopf- und Fußzeile", key: "JH", contextual: true, shrink: ["position", "options", "navigation", "insert", "hf"],
    groups: [
      { id: "hf", label: "Kopf- und Fußzeile", key: "ZK", icon: "header",
        controls: [dd("insert.header", "H", "large", headerMenu), dd("insert.footer", "O", "large", footerMenu), dd("insert.pageNumber", "NU", "large", pageNumberMenu)] },
      { id: "insert", label: "Einfügen", key: "ZE", icon: "date-time", controls: [b("insert.dateTime", "D", "large")] },
      { id: "navigation", label: "Navigation", key: "ZN", icon: "go-header",
        controls: [b("hf.gotoHeader", "H1", "large"), b("hf.gotoFooter", "F1", "large")] },
      { id: "options", label: "Optionen", key: "ZO", icon: "first-page",
        controls: [{ type: "stack", items: [{ type: "check", cmd: "hf.firstPage", key: "A" }, { type: "check", cmd: "hf.evenOdd", key: "V" }] }] },
      { id: "position", label: "Position", key: "ZP", icon: "margins",
        controls: [{ type: "stack", items: [
          { type: "spinner", cmd: "hf.headerDistance", label: "Kopfzeile von oben:", icon: "go-header", key: "P", value: hd.value, commit: hd.commit, step: hd.step },
          { type: "spinner", cmd: "hf.footerDistance", label: "Fußzeile von unten:", icon: "go-footer", key: "B", value: fd.value, commit: fd.commit, step: fd.step },
        ] }] },
      { id: "close", label: "Schließen", key: "ZS", icon: "close-x", controls: [b("hf.close", "C", "large")] },
    ],
  };
  return { tab, headerMenu, footerMenu, pageNumberMenu };
}
