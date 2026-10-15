// Alternativtext as Word 365 shows it: a task pane on the right (not a dialog) for the selected
// picture or shape, which follows the selection and keeps what is typed when the field is left.

import { h, icon } from "@raabe/shell";
import type { AppApi } from "./api";
import type { Dock } from "./dock";

export function altTextPane(api: AppApi, docArea: HTMLElement, dock: Dock) {
  const ed = () => api.editor();
  const pane = h("div", { class: "tpane altpane", hidden: true, role: "complementary", "aria-label": "Alternativtext" });
  docArea.append(pane);
  dock.add(pane, "right", 280);
  let open = false;
  /** The object the fields show (its text is saved when it changes or the pane closes). */
  let shown: string | null = null;

  const text = h("textarea", { rows: 6, class: "alt-text", "aria-label": "Alternativtext" }) as HTMLTextAreaElement;
  const deco = h("input", { type: "checkbox" }) as HTMLInputElement;
  const hint = h("div", { class: "alt-hint" });

  const save = () => {
    const e = ed();
    const cur = e?.pictureAlt();
    if (!e || !cur) return;
    const descr = text.value.trim() || undefined;
    if ((cur.descr ?? undefined) === descr && !!cur.decorative === deco.checked) return;
    e.setPictureAlt({ descr, decorative: deco.checked || undefined, title: cur.title });
  };
  text.addEventListener("change", save);
  deco.addEventListener("change", () => { text.disabled = deco.checked; save(); });

  const close = () => { save(); open = false; pane.hidden = true; api.refresh(); api.focus(); };

  const refresh = () => {
    if (!open) return;
    const e = ed();
    const cur = e?.pictureAlt() ?? null;
    const sel = e?.selectedImage() ?? null;
    const key = sel ? `${sel.kind}:${sel.p}:${sel.kind === "float" ? sel.index : sel.o}` : null;
    if (key === shown && document.activeElement === text) return;
    shown = key;
    const has = !!cur;
    text.disabled = !has || !!cur?.decorative;
    deco.disabled = !has;
    text.value = cur?.descr ?? "";
    deco.checked = !!cur?.decorative;
    hint.textContent = has ? "Wie würden Sie dieses Objekt und seinen Kontext für eine Person beschreiben, die blind ist oder eine Sehbehinderung hat?"
      : "Wählen Sie ein Bild oder eine Form aus, um Alternativtext hinzuzufügen.";
  };

  pane.append(
    h("div", { class: "tp-head" }, h("span", null, "Alternativtext"),
      h("div", null, h("button", { "aria-label": "Bereich schließen", "data-tip": "Bereich schließen", onclick: close }, icon("close")))),
    h("div", { class: "tp-body alt-body" }, hint, text,
      h("label", { class: "check" }, deco, "Als dekorativ markieren"),
      h("div", { class: "alt-note" }, "Dekorative Objekte (etwa Rahmen) machen eine Seite schöner, sagen aber nichts aus; Sprachausgaben übergehen sie.")));

  return {
    open: () => { open = true; shown = null; pane.hidden = false; refresh(); text.focus(); },
    refresh,
    get isOpen() { return open; },
  };
}
