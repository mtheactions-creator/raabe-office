// Modal dialogs and message boxes inside the window (Enter = default button, Escape =
// cancel), e.g. Word's "Möchten Sie die Änderungen an „Dokument1“ speichern?".

import { h, icon } from "./dom";

export interface DialogButton<T> {
  label: string;
  value: T;
  primary?: boolean;
}

export function showDialog<T>(opts: { title: string; body: HTMLElement | string; buttons: DialogButton<T>[]; cancel: T }): Promise<T> {
  return new Promise((resolve) => {
    const previous = document.activeElement as HTMLElement | null;
    const done = (v: T) => {
      layer.remove();
      removeEventListener("keydown", key, true);
      previous?.focus?.();
      resolve(v);
    };
    const buttons = opts.buttons.map((b) => h("button", { class: b.primary ? "primary" : "", onclick: () => done(b.value) }, b.label));
    const dialog = h("div", { class: "dialog", role: "dialog", "aria-modal": "true", "aria-label": opts.title },
      h("div", { class: "dlg-title" }, h("span", { class: "t" }, opts.title),
        h("button", { "aria-label": "Schließen", onclick: () => done(opts.cancel) }, icon("close"))),
      h("div", { class: "dlg-body" }, opts.body),
      h("div", { class: "dlg-buttons" }, ...buttons));
    const layer = h("div", { class: "dialog-layer" }, dialog);
    const key = (e: KeyboardEvent) => {
      if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); done(opts.cancel); }
      else if (e.key === "Enter" && !(e.target instanceof HTMLTextAreaElement)) {
        const focused = document.activeElement;
        if (focused instanceof HTMLButtonElement && dialog.contains(focused)) return;
        const primary = opts.buttons.find((b) => b.primary);
        if (primary) { e.preventDefault(); done(primary.value); }
      } else if (e.key === "Tab") {
        // Keep focus inside the dialog.
        const focusables = Array.from(dialog.querySelectorAll<HTMLElement>("button, input, select, textarea"));
        const i = focusables.indexOf(document.activeElement as HTMLElement);
        if (e.shiftKey && i <= 0) { e.preventDefault(); focusables[focusables.length - 1]?.focus(); }
        else if (!e.shiftKey && i === focusables.length - 1) { e.preventDefault(); focusables[0]?.focus(); }
      }
    };
    addEventListener("keydown", key, true);
    document.body.append(layer);
    const first = dialog.querySelector<HTMLElement>(".dlg-body input, .dlg-body select") ?? buttons[opts.buttons.findIndex((b) => b.primary)] ?? buttons[0];
    first?.focus();
  });
}

export function messageBox(title: string, text: string): Promise<void> {
  return showDialog({ title, body: text, buttons: [{ label: "OK", value: undefined, primary: true }], cancel: undefined });
}
