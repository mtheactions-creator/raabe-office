// Task panes docked beside the document, as in Word: Navigation, Überarbeitungsbereich and
// Zwischenablage on the left, Formatvorlagen, Auswahl and Kommentare on the right, the
// horizontal Überarbeitungsbereich at the bottom, the rulers above ("top") and left of the
// document ("inner", below the horizontal ruler). Open panes stand side by side; the document
// takes the room that is left (CSS variables --dock-l, --dock-r, --dock-t, --dock-b).

export type DockSide = "left" | "right" | "bottom" | "top" | "inner";

export interface Dock {
  /** Docks a pane (shown while not `hidden`); calling it again moves it to another side. */
  add(el: HTMLElement, side: DockSide, size: number): void;
  layout(): void;
}

export function makeDock(area: HTMLElement): Dock {
  const panes: { el: HTMLElement; side: DockSide; size: number }[] = [];
  const layout = () => {
    let l = 0, r = 0, b = 0, t = 0, inner = 0;
    for (const p of panes) {
      if (p.el.hidden) continue;
      if (p.side === "left") { p.el.style.left = `${l}px`; p.el.style.right = ""; l += p.size; }
      else if (p.side === "right") { p.el.style.right = `${r}px`; p.el.style.left = ""; r += p.size; }
    }
    // The horizontal ruler spans the document area (above the vertical ruler too).
    for (const p of panes) {
      if (p.el.hidden || p.side !== "top") continue;
      p.el.style.left = `${l}px`;
      p.el.style.right = `${r}px`;
      p.el.style.top = `${t}px`;
      t += p.size;
    }
    for (const p of panes) {
      if (p.el.hidden || p.side !== "inner") continue;
      p.el.style.left = `${l + inner}px`;
      p.el.style.top = `${t}px`;
      inner += p.size;
    }
    // A bottom pane spans the document's width between the side panes.
    for (const p of panes) {
      if (p.el.hidden || p.side !== "bottom") continue;
      p.el.style.left = `${l}px`;
      p.el.style.right = `${r}px`;
      p.el.style.bottom = `${b}px`;
      b += p.size;
    }
    for (const p of panes) if (!p.el.hidden && p.side === "inner") p.el.style.bottom = `${b}px`;
    area.style.setProperty("--dock-l", `${l + inner}px`);
    area.style.setProperty("--dock-r", `${r}px`);
    area.style.setProperty("--dock-t", `${t}px`);
    area.style.setProperty("--dock-b", `${b}px`);
  };
  const watch = new MutationObserver(layout);
  return {
    add(el, side, size) {
      const at = panes.findIndex((p) => p.el === el);
      if (at >= 0) panes[at] = { el, side, size };
      else {
        panes.push({ el, side, size });
        watch.observe(el, { attributes: true, attributeFilter: ["hidden"] });
      }
      layout();
    },
    layout,
  };
}
