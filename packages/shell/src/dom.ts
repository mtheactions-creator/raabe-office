// A tiny DOM helper: no framework in the shell either, so nothing stands between an
// input event and the screen.

import { ICONS } from "./icons.gen";

type Child = Node | string | null | undefined | false;
type Props = {
  class?: string;
  style?: string;
  title?: string;
  [attr: `aria-${string}`]: string | undefined;
  [attr: `data-${string}`]: string | undefined;
  [event: `on${string}`]: ((e: never) => void) | string | undefined;
} & Record<string, unknown>;

export function h<K extends keyof HTMLElementTagNameMap>(tag: K, props: Props | null = null, ...children: Child[]): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  if (props) {
    for (const [k, v] of Object.entries(props)) {
      if (v === undefined || v === null || v === false) continue;
      if (k.startsWith("on") && typeof v === "function") el.addEventListener(k.slice(2), v as EventListener);
      else if (k === "class") el.className = String(v);
      else if (k === "style") el.style.cssText = String(v);
      else if (k in el && typeof v !== "string") (el as unknown as Record<string, unknown>)[k] = v;
      else el.setAttribute(k, v === true ? "" : String(v));
    }
  }
  for (const c of children) if (c !== null && c !== undefined && c !== false) el.append(c);
  return el;
}

/** An icon from the generated Fluent set: "s" = 16-20 px, "l" = 32 px. */
export function icon(name: string | undefined, size: "s" | "l" = "s"): HTMLElement {
  const span = document.createElement("span");
  span.className = `ico ico-${size}`;
  const svg = name ? ICONS[name]?.[size] : undefined;
  if (svg) span.innerHTML = svg;
  return span;
}

export function clear(el: Element): void {
  el.replaceChildren();
}
