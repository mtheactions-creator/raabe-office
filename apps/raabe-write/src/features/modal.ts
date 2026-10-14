// A dialog that stays open across actions (Textmarke > Löschen, Symbol > Einfügen …): Word's
// look (title bar, body, buttons at the bottom right), Escape and the close box cancel. A
// button's `run` returns true to close the dialog.

import { h, icon } from "@raabe/shell";

export interface ModalButton {
  label: string;
  primary?: boolean;
  disabled?: boolean;
  run: () => boolean | void | Promise<boolean | void>;
}

export interface Modal {
  close(): void;
  button(label: string): HTMLButtonElement | undefined;
  readonly el: HTMLElement;
}

export function openModal(title: string, body: HTMLElement, buttons: ModalButton[], opts: { width?: number; onClose?: () => void; cls?: string } = {}): Modal {
  const btns = buttons.map((b) => {
    const el = h("button", { class: b.primary ? "primary" : "", disabled: !!b.disabled }, b.label) as HTMLButtonElement;
    el.onclick = async () => { if ((await b.run()) === true) close(); };
    return el;
  });
  const box = h("div", { class: `dialog ${opts.cls ?? ""}`, role: "dialog", "aria-modal": "true", "aria-label": title, style: opts.width ? `width:${opts.width}px` : "" },
    h("div", { class: "dlg-title" }, h("span", { class: "t" }, title), h("button", { "aria-label": "Schließen", onclick: () => close() }, icon("close"))),
    h("div", { class: "dlg-body" }, body),
    h("div", { class: "dlg-buttons" }, ...btns));
  const layer = h("div", { class: "dialog-layer" }, box);
  const previous = document.activeElement as HTMLElement | null;
  const key = (e: KeyboardEvent) => {
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); close(); }
    else if (e.key === "Enter" && !(e.target instanceof HTMLTextAreaElement) && !(e.target instanceof HTMLButtonElement)) {
      const i = buttons.findIndex((b) => b.primary);
      if (i >= 0 && !btns[i].disabled) { e.preventDefault(); btns[i].click(); }
    }
  };
  let closed = false;
  const close = () => {
    if (closed) return;
    closed = true;
    layer.remove();
    removeEventListener("keydown", key, true);
    opts.onClose?.();
    previous?.focus?.();
  };
  addEventListener("keydown", key, true);
  document.body.append(layer);
  (body.querySelector<HTMLElement>("input, select, textarea, [tabindex]") ?? btns[0])?.focus();
  return { close, button: (label) => btns[buttons.findIndex((b) => b.label === label)], el: box };
}
