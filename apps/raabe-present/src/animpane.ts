// Animationsbereich (D-070): the slide's main sequence in order, each effect with its click number,
// its kind (Eingang green, Hervorheben yellow, Beenden red) and its shape's name; a click selects the
// shape, the arrows move the effect earlier or later, Entf removes it, "Alle wiedergeben" previews on
// the slide. Start (Beim Klicken, Mit Vorherigen, Nach Vorherigen) and Früher / Später in the ribbon act
// on the effects of the selected shapes. The timing tree is changed as PowerPoint writes it.

import { TaskPane } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";

const { kid, kids } = PptxDom;

interface Effect { par: Element; ctn: Element; spid: number; cls: string; click: number; node: string }

/** The effects of the main sequence with their click groups. */
export function effects(root: Element): { main: Element | null; list: Effect[] } {
  const timing = kid(root, "timing");
  const main = timing ? Array.from(timing.getElementsByTagNameNS(PptxDom.NS_P, "cTn")).find((c) => c.getAttribute("nodeType") === "mainSeq") ?? null : null;
  const list: Effect[] = [];
  let click = 0;
  for (const clickPar of kids(kid(main, "childTnLst"), "par")) {
    const withClick = Array.from(clickPar.getElementsByTagNameNS(PptxDom.NS_P, "cTn")).filter((c) => c.hasAttribute("presetClass"));
    if (withClick.some((c) => c.getAttribute("nodeType") === "clickEffect")) click++;
    for (const ctn of withClick) {
      list.push({ par: ctn.parentElement!, ctn, spid: Number(ctn.getElementsByTagNameNS(PptxDom.NS_P, "spTgt")[0]?.getAttribute("spid") ?? 0),
        cls: ctn.getAttribute("presetClass") ?? "entr", click: Math.max(click, 1), node: ctn.getAttribute("nodeType") ?? "clickEffect" });
    }
  }
  return { main, list };
}

const LABEL: Record<string, string> = { entr: "Eingang", emph: "Hervorheben", exit: "Beenden", path: "Animationspfad" };

export class AnimationPane {
  readonly pane: TaskPane;
  constructor(private app: PresentApp, host: HTMLElement) {
    this.pane = new TaskPane(host);
  }

  get isOpen(): boolean { return this.pane.isOpen; }

  toggle(): void {
    if (this.pane.isOpen) this.pane.close();
    else this.pane.show({ title: "Animationsbereich", sets: [{ label: "", tabs: [{ id: "a", label: "Animationen", icon: "animation", sections: () => [{ title: "Animationen", open: true, controls: () => [{ type: "custom", render: () => this.list() }] }] }] }] });
  }

  refresh(): void { if (this.pane.isOpen) this.pane.refresh(); }

  private list(): HTMLElement {
    const box = document.createElement("div");
    box.className = "anim-list";
    const path = this.app.slidePath();
    if (!path) return box;
    const root = this.app.deck.pptx.doc(path).documentElement;
    const names = new Map(this.app.deck.view(path).shapes.map((s) => [s.id, s.name]));
    const { list } = effects(root);
    const play = document.createElement("button");
    play.textContent = "▶ Alle wiedergeben";
    play.onclick = () => void this.app.commands.run("anim.preview");
    box.append(play);
    list.forEach((e, i) => {
      const row = document.createElement("div");
      row.className = `anim-row anim-${e.cls}${this.app.editor.sel.includes(e.spid) ? " on" : ""}`;
      const num = document.createElement("span");
      num.className = "anim-num";
      num.textContent = e.node === "clickEffect" ? String(e.click) : "";
      const name = document.createElement("span");
      name.className = "anim-name";
      name.textContent = names.get(e.spid) ?? `Form ${e.spid}`;
      name.title = LABEL[e.cls] ?? e.cls;
      const up = document.createElement("button");
      up.textContent = "▲"; up.setAttribute("aria-label", "Früher");
      up.onclick = (ev) => { ev.stopPropagation(); this.move(i, -1); };
      const down = document.createElement("button");
      down.textContent = "▼"; down.setAttribute("aria-label", "Später");
      down.onclick = (ev) => { ev.stopPropagation(); this.move(i, 1); };
      row.append(num, Object.assign(document.createElement("span"), { className: "anim-icon" }), name, up, down);
      row.onclick = () => this.app.editor.select([e.spid]);
      row.tabIndex = 0;
      row.onkeydown = (ev) => { if (ev.key === "Delete") { ev.preventDefault(); this.remove(i); } };
      box.append(row);
    });
    if (!list.length) box.append(Object.assign(document.createElement("div"), { className: "muted", textContent: "Auf dieser Folie gibt es keine Animationen." }));
    return box;
  }

  /** Moves effect i one place earlier or later (its click group moves with it when it starts one). */
  move(i: number, d: -1 | 1): void {
    const path = this.app.slidePath();
    if (!path) return;
    this.app.editor.edit(d < 0 ? "Früher" : "Später", () => {
      const root = this.app.deck.pptx.doc(path).documentElement;
      const { list } = effects(root);
      const e = list[i], o = list[i + d];
      if (!e || !o) return;
      // The outermost click group of each effect.
      const group = (x: Effect) => x.par.parentElement!.parentElement!.parentElement!.parentElement!.parentElement!;
      const ge = group(e), go = group(o);
      if (ge === go) { if (d < 0) o.par.before(e.par); else o.par.after(e.par); }
      else if (d < 0) go.before(ge); else go.after(ge);
    });
    this.refresh();
  }

  remove(i: number): void {
    const path = this.app.slidePath();
    if (!path) return;
    this.app.editor.edit("Animation entfernen", () => {
      const root = this.app.deck.pptx.doc(path).documentElement;
      const e = effects(root).list[i];
      if (!e) return;
      let p: Element | null = e.par;
      // The effect, then every container left empty.
      while (p && p.localName === "par") {
        const parent: Element | null = p.parentElement;
        p.remove();
        if (parent && parent.localName === "childTnLst" && !parent.children.length && parent.parentElement?.getAttribute("nodeType") !== "mainSeq") p = parent.parentElement?.parentElement ?? null;
        else p = null;
      }
    });
    this.refresh();
  }

  /** Start: Beim Klicken, Mit Vorherigen, Nach Vorherigen for the selected shapes' effects. */
  setStart(kind: "clickEffect" | "withEffect" | "afterEffect"): void {
    const path = this.app.slidePath();
    if (!path) return;
    const ids = new Set(this.app.editor.sel);
    this.app.editor.edit("Start", () => {
      const root = this.app.deck.pptx.doc(path).documentElement;
      for (const e of effects(root).list) if (ids.has(e.spid)) e.ctn.setAttribute("nodeType", kind);
    });
    this.refresh();
  }
}
