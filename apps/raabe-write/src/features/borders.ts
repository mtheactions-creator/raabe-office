// Rahmen und Schattierung (Start > Rahmen > Rahmen und Schattierung..., Entwurf >
// Seitenränder): Word's dialog with Rahmen (the paragraph's borders), Seitenrand (the section's
// page borders, w:pgBorders, recorded 2026-10-14) and Schattierung (the paragraph's fill).
// Einstellung Ohne / Kontur / Benutzerdefiniert, Formatvorlage, Farbe, Breite, the preview's
// side buttons, Übernehmen für and Optionen (distances, Gemessen von).

import { colorGrid, h, icon, showPopup, closePopup } from "@raabe/shell";
import type { BorderSpec, Borders, Editor, PageBorders } from "@raabe/text-engine";
import type { AppApi } from "./api";
import { openModal } from "./modal";

const STYLES: [string, string][] = [["single", "solid"], ["dotted", "dotted"], ["dashSmallGap", "dashed"], ["dashed", "dashed"], ["dotDash", "dashed"],
  ["dotDotDash", "dotted"], ["double", "double"], ["triple", "double"], ["thinThickSmallGap", "double"], ["thickThinSmallGap", "double"], ["wave", "solid"]];
const WIDTHS: [number, string][] = [[0.25, "¼ Pt."], [0.5, "½ Pt."], [0.75, "¾ Pt."], [1, "1 Pt."], [1.5, "1 ½ Pt."], [2.25, "2 ¼ Pt."], [3, "3 Pt."], [4.5, "4 ½ Pt."], [6, "6 Pt."]];
type Side = "top" | "bottom" | "left" | "right";
const SIDES: Side[] = ["top", "bottom", "left", "right"];
const SIDE_TIP: Record<Side, string> = { top: "Rahmenlinie oben", bottom: "Rahmenlinie unten", left: "Rahmenlinie links", right: "Rahmenlinie rechts" };

export function bordersFeature(api: AppApi) {
  const ed = () => api.editor();

  const dialog = (first: "borders" | "page" | "shading" = "borders") => {
    const e = ed();
    if (!e) return;
    const rp = e.paraState();
    const para: Partial<Record<Side, BorderSpec>> = { top: rp.borders?.top, bottom: rp.borders?.bottom, left: rp.borders?.left, right: rp.borders?.right };
    const pb = e.section().pageBorders;
    const page: Partial<Record<Side, BorderSpec>> = { top: pb?.top, bottom: pb?.bottom, left: pb?.left, right: pb?.right };
    const pageOpts = { from: pb?.offsetFrom ?? "text", space: { top: pb?.top?.space ?? 24, bottom: pb?.bottom?.space ?? 24, left: pb?.left?.space ?? 24, right: pb?.right?.space ?? 24 } };
    const paraSpace = { top: rp.borders?.top?.space ?? 1, bottom: rp.borders?.bottom?.space ?? 1, left: rp.borders?.left?.space ?? 4, right: rp.borders?.right?.space ?? 4 };
    let fill = rp.shading?.fill && rp.shading.fill !== "auto" ? rp.shading.fill : "";

    /** One tab's controls for a border set (paragraph or page). */
    const borderTab = (set: Partial<Record<Side, BorderSpec>>, kind: "para" | "page") => {
      const any = SIDES.map((s) => set[s]).find(Boolean);
      const cur = { style: any?.style ?? "single", color: any?.color ?? "auto", size: any?.size ?? 0.5 };
      const styleList = h("div", { class: "bd-styles", role: "listbox" });
      const colorBtn = h("button", { class: "bd-color" }) as HTMLButtonElement;
      const widthSel = h("select", {}, ...WIDTHS.map(([v, l]) => h("option", { value: String(v) }, l))) as HTMLSelectElement;
      widthSel.value = String(WIDTHS.reduce((a, b) => (Math.abs(b[0] - cur.size) < Math.abs(a[0] - cur.size) ? b : a))[0]);
      const preview = h("div", { class: "bd-preview" }, ...Array.from({ length: 9 }, (_, i) => h("i", { style: i === 0 ? "width:60%" : i === 8 ? "width:70%" : "" })));
      const sideBtns = Object.fromEntries(SIDES.map((s) => [s, h("button", { class: `bd-side ${s}`, "data-tip": SIDE_TIP[s], "aria-label": SIDE_TIP[s] }, icon(`border-${s}`))])) as Record<Side, HTMLButtonElement>;
      const settings = h("div", { class: "bd-settings" });
      const spaceOf = (s: Side) => (kind === "page" ? pageOpts.space[s] : paraSpace[s]);
      const spec = (s: Side): BorderSpec => ({ style: cur.style, color: cur.color, size: cur.size, space: spaceOf(s) });
      const paint = () => {
        const css = (b: BorderSpec | undefined) => (b ? `${Math.max(1, b.size * 1.33)}px ${STYLES.find(([k]) => k === b.style)?.[1] ?? "solid"} ${b.color === "auto" ? "#000" : `#${b.color}`}` : "1px dashed transparent");
        preview.style.borderTop = css(set.top);
        preview.style.borderBottom = css(set.bottom);
        preview.style.borderLeft = css(set.left);
        preview.style.borderRight = css(set.right);
        for (const s of SIDES) sideBtns[s].classList.toggle("on", !!set[s]);
        colorBtn.replaceChildren(h("span", { style: `background:${cur.color === "auto" ? "#000" : `#${cur.color}`}` }), cur.color === "auto" ? "Automatisch" : "", icon("chevron-down"));
        styleList.querySelectorAll("button").forEach((b) => b.classList.toggle("active", (b as HTMLElement).dataset.s === cur.style));
        const box = SIDES.every((s) => set[s]), noneSet = SIDES.every((s) => !set[s]);
        settings.querySelectorAll("button").forEach((b) => b.classList.toggle("active", (b as HTMLElement).dataset.k === (noneSet ? "none" : box ? "box" : "custom")));
      };
      for (const [k, css] of STYLES) {
        const b = h("button", { "data-s": k, "aria-label": k, onclick: () => { cur.style = k; for (const s of SIDES) if (set[s]) set[s] = spec(s); paint(); } },
          h("span", { style: `border-top:${k === "double" || k === "triple" || k.startsWith("thin") || k.startsWith("thick") ? "3px" : "1px"} ${css} #000` }));
        styleList.append(b);
      }
      colorBtn.onclick = () => {
        const pick = (hex: string) => { closePopup(); cur.color = hex; for (const s of SIDES) if (set[s]) set[s] = spec(s); paint(); };
        showPopup(h("div", { class: "menu palette" }, h("button", { onclick: () => pick("auto") }, h("span", { class: "auto-swatch" }), "Automatisch"), ...colorGrid((hex) => pick(hex), e.theme.colors)),
          colorBtn.getBoundingClientRect());
      };
      widthSel.onchange = () => { cur.size = Number(widthSel.value); for (const s of SIDES) if (set[s]) set[s] = spec(s); paint(); };
      for (const s of SIDES) sideBtns[s].onclick = () => { set[s] = set[s] ? undefined : spec(s); paint(); };
      const setting = (k: string, label: string, run: (() => void) | null, tip?: string) => h("button", {
        class: "bd-setting", "data-k": k, disabled: !run, ...(tip ? { "data-tip": tip } : {}), onclick: () => { run?.(); paint(); },
      }, h("span", { class: `bd-icon ${k}` }), label);
      settings.append(
        setting("none", "Ohne", () => { for (const s of SIDES) set[s] = undefined; }),
        setting("box", "Kontur", () => { for (const s of SIDES) set[s] = spec(s); }),
        setting("shadow", "Schatten", null, "Schatten\nRaabeWrite zeichnet Rahmen ohne Schatten."),
        setting("three-d", "3D", null, "3D\nRaabeWrite zeichnet Rahmen ohne 3D-Effekt."),
        setting("custom", kind === "page" ? "Anpassen" : "Benutzerdefiniert", () => {}));
      const applyTo = kind === "page"
        ? h("select", {}, h("option", { value: "all" }, "Gesamtes Dokument"), h("option", { value: "section" }, "Diesen Abschnitt"),
          h("option", { value: "first" }, "Diesen Abschnitt - nur erste Seite"), h("option", { value: "notFirst" }, "Diesen Abschnitt - alle außer erster Seite")) as HTMLSelectElement
        : h("select", {}, h("option", { value: "para" }, "Absatz")) as HTMLSelectElement;
      if (kind === "page") applyTo.value = pb?.display === "firstPage" ? "first" : pb?.display === "notFirstPage" ? "notFirst" : e.sectionCount() > 1 ? "section" : "all";
      const options = h("button", { onclick: () => optionsDialog(kind, paint) }, "Optionen...");
      const effects = h("select", { disabled: true, "data-tip": "Effekte\nRaabeWrite hat keine Bildrahmen." }, h("option", null, "(ohne)"));
      const el = h("div", { class: "bd-tab" },
        h("div", { class: "bd-col" }, h("div", { class: "bd-label" }, "Einstellung:"), settings),
        h("div", { class: "bd-col mid" }, h("div", { class: "bd-label" }, "Formatvorlage:"), styleList, h("div", { class: "bd-label" }, "Farbe:"), colorBtn,
          h("div", { class: "bd-label" }, "Breite:"), widthSel, ...(kind === "page" ? [h("div", { class: "bd-label" }, "Effekte:"), effects] : [])),
        h("div", { class: "bd-col right" }, h("div", { class: "bd-label" }, "Vorschau"),
          h("div", { class: "bd-hint" }, "Diagramm oder Schaltflächen klicken, um Rahmen hinzuzufügen"),
          h("div", { class: "bd-prevwrap" }, h("div", { class: "bd-vbtns" }, sideBtns.top, sideBtns.bottom), preview, h("div", { class: "bd-hbtns" }, sideBtns.left, sideBtns.right)),
          h("div", { class: "bd-label" }, "Übernehmen für:"), applyTo, h("div", { class: "row end" }, options)));
      // A click on the preview's edges toggles that side.
      preview.addEventListener("click", (ev) => {
        const r = preview.getBoundingClientRect();
        const dx = Math.min(ev.clientX - r.left, r.right - ev.clientX), dy = Math.min(ev.clientY - r.top, r.bottom - ev.clientY);
        const s: Side = dx < dy ? (ev.clientX - r.left < r.right - ev.clientX ? "left" : "right") : (ev.clientY - r.top < r.bottom - ev.clientY ? "top" : "bottom");
        set[s] = set[s] ? undefined : spec(s);
        paint();
      });
      paint();
      return { el, applyTo };
    };

    const optionsDialog = (kind: "para" | "page", after: () => void) => {
      const sp = kind === "page" ? pageOpts.space : paraSpace;
      const num = (v: number) => h("input", { type: "number", min: "0", max: "31", value: String(v), style: "width:70px" }) as HTMLInputElement;
      const t = num(sp.top), b = num(sp.bottom), l = num(sp.left), r = num(sp.right);
      const from = h("select", {}, h("option", { value: "text" }, "Text"), h("option", { value: "page" }, "Seitenrand")) as HTMLSelectElement;
      from.value = pageOpts.from;
      openModal("Rahmen- und Schattierungsoptionen", h("div", { class: "bd-opts" },
        h("fieldset", null, h("legend", null, "Abstand vom Text"), h("div", { class: "grid2" }, "Oben:", t, "Unten:", b, "Links:", l, "Rechts:", r)),
        ...(kind === "page" ? [h("label", { class: "field row" }, "Gemessen von:", from)] : [])), [
        { label: "OK", primary: true, run: () => {
          Object.assign(sp, { top: Number(t.value) || 0, bottom: Number(b.value) || 0, left: Number(l.value) || 0, right: Number(r.value) || 0 });
          if (kind === "page") pageOpts.from = from.value as "page" | "text";
          const set = kind === "page" ? page : para;
          for (const s of SIDES) if (set[s]) set[s] = { ...set[s]!, space: sp[s] };
          after();
          return true;
        } },
        { label: "Abbrechen", run: () => true },
      ], { width: 360 });
    };

    const tParas = borderTab(para, "para"), tPage = borderTab(page, "page");
    // Schattierung: Füllung (Muster as Word's, not drawn by RaabeWrite).
    const fillBtn = h("button", { class: "bd-color wide" }) as HTMLButtonElement;
    const paintFill = () => fillBtn.replaceChildren(h("span", { style: fill ? `background:#${fill}` : "background:#fff;border:1px solid #ccc" }), fill ? "" : "Keine Farbe", icon("chevron-down"));
    fillBtn.onclick = () => {
      const pick = (hex: string) => { closePopup(); fill = hex; paintFill(); };
      showPopup(h("div", { class: "menu palette" }, ...colorGrid((hex) => pick(hex), e.theme.colors), h("hr"), h("button", { onclick: () => pick("") }, "Keine Farbe")), fillBtn.getBoundingClientRect());
    };
    paintFill();
    const shadingTab = h("div", { class: "bd-tab shading" },
      h("div", { class: "bd-col" }, h("div", { class: "bd-label" }, "Füllung"), fillBtn,
        h("div", { class: "bd-label" }, "Muster"), h("label", { class: "field row" }, "Lage:", h("select", { disabled: true, "data-tip": "Muster\nRaabeWrite schattiert mit einer Farbe." }, h("option", null, "Transparent"))),
        h("label", { class: "field row" }, "Farbe:", h("select", { disabled: true }, h("option", null, "Automatisch")))),
      h("div", { class: "bd-col right" }, h("div", { class: "bd-label" }, "Übernehmen für:"), h("select", {}, h("option", null, "Absatz"))));

    const tabs = h("div", { class: "tabs" });
    const pages: Record<string, HTMLElement> = { borders: tParas.el, page: tPage.el, shading: shadingTab };
    const show = (k: string) => {
      for (const [n, el] of Object.entries(pages)) el.hidden = n !== k;
      tabs.querySelectorAll("button").forEach((b) => b.classList.toggle("active", (b as HTMLElement).dataset.k === k));
    };
    for (const [k, label] of [["borders", "Rahmen"], ["page", "Seitenrand"], ["shading", "Schattierung"]]) tabs.append(h("button", { "data-k": k, onclick: () => show(k) }, label));
    show(first);

    const save = (x: Editor) => {
      // Paragraph borders and shading (kept as they were when unchanged).
      const before = JSON.stringify([rp.borders?.top, rp.borders?.bottom, rp.borders?.left, rp.borders?.right]);
      const borders: Borders | undefined = SIDES.some((s) => para[s]) ? { ...(rp.borders ?? {}), top: para.top, bottom: para.bottom, left: para.left, right: para.right } : undefined;
      if (JSON.stringify([para.top, para.bottom, para.left, para.right]) !== before) x.setParaProps({ borders });
      const oldFill = rp.shading?.fill && rp.shading.fill !== "auto" ? rp.shading.fill : "";
      if (fill !== oldFill) x.setParaProps({ shading: fill ? { fill, val: "clear", color: "auto" } : undefined });
      // Page borders.
      const nextPb: PageBorders | undefined = SIDES.some((s) => page[s]) ? {
        top: page.top, bottom: page.bottom, left: page.left, right: page.right, offsetFrom: pageOpts.from,
        ...(tPage.applyTo.value === "first" ? { display: "firstPage" as const } : tPage.applyTo.value === "notFirst" ? { display: "notFirstPage" as const } : {}),
      } : undefined;
      if (JSON.stringify(nextPb ?? null) !== JSON.stringify(pb ?? null)) x.changeSections((s) => ({ ...s, pageBorders: nextPb }), tPage.applyTo.value === "all");
    };
    openModal("Rahmen und Schattierung", h("div", { class: "bddlg" }, tabs, ...Object.values(pages)), [
      { label: "OK", primary: true, run: () => { const x = ed(); if (x) { save(x); api.refresh(); api.focus(); } return true; } },
      { label: "Abbrechen", run: () => true },
    ], { width: 700, cls: "bordersdialog" });
  };

  api.commands.implement("design.pageBorders", { run: () => dialog("page"), enabled: () => !!ed() });
  api.commands.implement("border.dialog", { run: () => dialog("borders"), enabled: () => !!ed() });
  return { dialog };
}
