// The macro editor (D-079, D-106): a window over the app like Office's VBA editor. Left the project (Module,
// Klassenmodule, Formulare); in the middle a module's code with line numbers and breakpoints, or a UserForm's
// designer (Werkzeugsammlung, Eigenschaften); below the Direktbereich (Debug.Print, statements typed while
// paused) and the Lokal-Fenster. Ausführen (F5), Einzelschritt (F8), Prozedurschritt (Umschalt+F8), Prozedur
// abschließen (Strg+Umschalt+F8), Haltepunkt ein/aus (F9) and Zurücksetzen drive @raabe/vba's debugger. Nothing
// here knows PowerPoint: the app gives its modules, how to save them and its VbaHost (object model, dialogs, file
// access), so RaabeCalc can open the same editor. Shown forms (UserForm.Show) are drawn by showVbaForm, which the
// app's own macro runs use too.

import { allControls, newControl, newForm, VbaRuntime, visibleCode, type VbaBreak, type VbaControlDesign, type VbaControlType, type VbaForm, type VbaHost, type VbaModule, type VbaStepMode, type VbaValue } from "@raabe/vba";
import { h } from "@raabe/shell";

export interface IdeHost {
  /** The document's name for the title bar. */
  title: string;
  modules(): VbaModule[];
  save(modules: VbaModule[]): void;
  /** The app's VbaHost with the editor's additions (debugBreak, debugPrint, showForm …) merged in. */
  vbaHost(extra: Partial<VbaHost>): VbaHost;
}

const PT = 4 / 3; // CSS pixels per point
const LINE = 18; // code line height in pixels

let styled = false;
function style(): void {
  if (styled) return;
  styled = true;
  document.head.append(h("style", null, `
.vba-ide{position:fixed;inset:0;z-index:5000;background:#f3f3f3;display:grid;grid-template-rows:auto auto 1fr 170px;font:12px "Segoe UI",sans-serif;color:#222}
.vba-ide .vi-title{background:#2b579a;color:#fff;padding:6px 10px;display:flex;justify-content:space-between;align-items:center}
.vba-ide .vi-title button{background:none;border:0;color:#fff;font-size:15px;cursor:default;padding:0 6px}
.vba-ide .vi-bar{display:flex;gap:4px;padding:4px 6px;border-bottom:1px solid #ccc;background:#fafafa;flex-wrap:wrap;align-items:center}
.vba-ide .vi-bar button{border:1px solid transparent;background:none;padding:3px 8px;border-radius:2px;font:inherit}
.vba-ide .vi-bar button:hover{background:#e5eefc;border-color:#9bbbe8}
.vba-ide .vi-bar .sep{width:1px;height:20px;background:#ccc;margin:0 4px}
.vba-ide .vi-main{display:grid;grid-template-columns:200px 1fr;min-height:0}
.vba-ide .vi-tree{border-right:1px solid #ccc;overflow:auto;background:#fff;padding:4px}
.vba-ide .vi-tree .grp{font-weight:600;margin-top:6px}
.vba-ide .vi-tree .it{padding:2px 4px 2px 16px;cursor:default;white-space:nowrap}
.vba-ide .vi-tree .it.on{background:#cce4f7}
.vba-ide .vi-code{display:grid;grid-template-columns:52px 1fr;min-height:0;background:#fff;position:relative}
.vba-ide .vi-gutter{overflow:hidden;background:#f0f0f0;border-right:1px solid #ddd;font:13px Consolas,monospace;color:#888;text-align:right;user-select:none}
.vba-ide .vi-gutter div{height:${LINE}px;line-height:${LINE}px;padding-right:6px;position:relative;cursor:default}
.vba-ide .vi-gutter div.bp::before{content:"";position:absolute;left:5px;top:4px;width:10px;height:10px;border-radius:50%;background:#a4262c}
.vba-ide .vi-gutter div.cur{background:#fff100;color:#000}
.vba-ide textarea.vi-text{border:0;outline:none;resize:none;font:13px Consolas,monospace;line-height:${LINE}px;padding:0 6px;white-space:pre;overflow:auto;tab-size:4}
.vba-ide .vi-cur{position:absolute;left:52px;right:0;height:${LINE}px;background:rgba(255,241,0,.45);pointer-events:none}
.vba-ide .vi-bottom{display:grid;grid-template-columns:1fr 1fr;border-top:1px solid #ccc;min-height:0}
.vba-ide .vi-pane{display:flex;flex-direction:column;min-height:0;border-right:1px solid #ccc;background:#fff}
.vba-ide .vi-pane h4{margin:0;padding:3px 6px;background:#e8e8e8;font-weight:600;font-size:12px}
.vba-ide .vi-log{flex:1;overflow:auto;font:12px Consolas,monospace;white-space:pre;padding:2px 6px}
.vba-ide .vi-imm{border:0;border-top:1px solid #ddd;font:12px Consolas,monospace;padding:3px 6px;outline:none}
.vba-ide .vi-locals{flex:1;overflow:auto}
.vba-ide .vi-locals table{border-collapse:collapse;width:100%;font:12px Consolas,monospace}
.vba-ide .vi-locals td,.vba-ide .vi-locals th{border:1px solid #e3e3e3;padding:1px 6px;text-align:left}
.vba-ide .vi-status{margin-left:auto;color:#555}
.vba-ide .vi-design{display:grid;grid-template-columns:120px 1fr 230px;min-height:0;background:#e9e9e9}
.vba-ide .vi-tools{display:flex;flex-direction:column;gap:2px;padding:6px;background:#fafafa;border-right:1px solid #ccc}
.vba-ide .vi-tools button{text-align:left;border:1px solid #ddd;background:#fff;padding:3px 6px;font:inherit}
.vba-ide .vi-tools button.on{background:#cce4f7;border-color:#7aaee6}
.vba-ide .vi-canvas{overflow:auto;padding:20px}
.vba-ide .vi-props{background:#fff;border-left:1px solid #ccc;overflow:auto}
.vba-ide .vi-props table{border-collapse:collapse;width:100%}
.vba-ide .vi-props td{border:1px solid #e3e3e3;padding:1px 4px}
.vba-ide .vi-props input{width:100%;border:0;font:inherit;box-sizing:border-box}
.vba-form{position:fixed;z-index:6000;background:#f0f0f0;border:1px solid #888;box-shadow:0 6px 24px rgba(0,0,0,.3);font:8pt Tahoma,sans-serif;color:#000}
.vba-form .vf-title{background:#fff;padding:5px 8px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #ddd;font:9pt "Segoe UI",sans-serif;cursor:move;user-select:none}
.vba-form .vf-title button{border:0;background:none;font-size:14px;cursor:default}
.vba-form .vf-body,.vba-ide .vf-body{position:relative;overflow:hidden}
.vf-body .vf-c{position:absolute;box-sizing:border-box;font:inherit}
.vf-body .vf-sel{outline:1px dashed #0063b1;outline-offset:1px}
.vf-body .vf-grip{position:absolute;right:-4px;bottom:-4px;width:7px;height:7px;background:#fff;border:1px solid #0063b1;cursor:nwse-resize}
.vba-form-shade{position:fixed;inset:0;z-index:5990;background:transparent}
`));
}

/** A colour of a form (VBA's OLE colour: system colours by index, else BGR) as CSS. */
function cssColor(c: number | undefined, fallback: string): string {
  if (c === undefined) return fallback;
  if (c >= 0x80000000) {
    const sys: Record<number, string> = { 0x0f: "#f0f0f0", 0x12: "#000000", 0x05: "#ffffff", 0x08: "#000000", 0x10: "#a0a0a0", 0x0d: "#0078d7", 0x0e: "#ffffff" };
    return sys[c & 0xff] ?? fallback;
  }
  return `rgb(${c & 255},${(c >> 8) & 255},${(c >> 16) & 255})`;
}

const text = (v: VbaValue): string => (v === null || v === undefined ? "" : typeof v === "boolean" ? (v ? "Wahr" : "Falsch") : String(v));

/** The HTML of one control: live (a shown form, events through the runtime) or for the designer. */
function controlEl(c: VbaControlDesign, live: ((prop: string, value: VbaValue) => void) | null, click: (() => void) | null): HTMLElement {
  const box = `left:${c.left * PT}px;top:${c.top * PT}px;width:${c.width * PT}px;height:${c.height * PT}px;` +
    `color:${cssColor(c.foreColor, "#000")};${c.font ? `font:${c.font.italic ? "italic " : ""}${c.font.bold ? "bold " : ""}${c.font.size}pt ${c.font.name},sans-serif;` : ""}` +
    (c.visible === false && live ? "display:none;" : "");
  const off = c.enabled === false;
  let el: HTMLElement;
  switch (c.type as VbaControlType) {
    case "CommandButton": case "ToggleButton":
      el = h("button", { type: "button", disabled: off, style: box + `background:${cssColor(c.backColor, "#e1e1e1")};border:1px solid #adadad` }, c.caption ?? "");
      if (c.type === "ToggleButton" && c.value === true) el.style.borderStyle = "inset";
      if (live) el.addEventListener("click", () => { if (c.type === "ToggleButton") live("value", !(c.value === true)); else click?.(); });
      break;
    case "Label":
      el = h("div", { style: box + `background:${c.backStyle === 0 ? "transparent" : cssColor(c.backColor, "transparent")};${c.borderStyle ? "border:1px solid #000;" : ""}white-space:pre-wrap;overflow:hidden` }, c.caption ?? "");
      if (live) el.addEventListener("click", () => click?.());
      break;
    case "TextBox": {
      const multi = !!c.multiLine;
      const inp = (multi ? h("textarea", { style: box + "resize:none" }) : h("input", { type: c.passwordChar ? "password" : "text", style: box })) as HTMLInputElement;
      inp.value = text(c.value);
      inp.disabled = off; inp.readOnly = !!c.locked || !live;
      if (live) inp.addEventListener("input", () => live("value", inp.value));
      el = inp;
      break;
    }
    case "CheckBox": case "OptionButton": {
      const inp = h("input", { type: c.type === "CheckBox" ? "checkbox" : "radio", disabled: off || !live }) as HTMLInputElement;
      inp.checked = c.value === true;
      if (live) inp.addEventListener("change", () => live("value", inp.checked));
      el = h("label", { style: box + "display:flex;align-items:center;gap:4px;white-space:nowrap;overflow:hidden" }, inp, c.caption ?? "");
      break;
    }
    case "ListBox": case "ComboBox": {
      const sel = h("select", { style: box, disabled: off || !live }) as HTMLSelectElement;
      if (c.type === "ListBox") sel.size = Math.max(2, Math.floor((c.height * PT) / 16));
      const rows = (c.list ?? []) as VbaValue[][];
      if (c.type === "ComboBox") sel.append(h("option", { value: "-1" }, ""));
      rows.forEach((r, i) => sel.append(h("option", { value: String(i) }, text(Array.isArray(r) ? r[0] : r))));
      sel.value = String(c.listIndex ?? -1);
      if (live) sel.addEventListener("change", () => live("listIndex", Number(sel.value)));
      el = sel;
      break;
    }
    case "Frame": {
      el = h("fieldset", { style: box + "margin:0;padding:0;border:1px solid #bbb" }, h("legend", { style: "font:inherit;padding:0 2px" }, c.caption ?? ""));
      const inner = h("div", { class: "vf-body", style: "position:relative;width:100%;height:100%" });
      for (const k of c.controls ?? []) inner.append(controlEl(k, live && ((p, v) => live(`${k.name}\u0000${p}`, v)), null));
      el.append(inner);
      break;
    }
    case "SpinButton": case "ScrollBar": {
      const inp = h("input", { type: c.type === "SpinButton" ? "number" : "range", style: box, disabled: off || !live, min: String(c.min ?? 0), max: String(c.max ?? 100), step: String(c.smallChange ?? 1) }) as HTMLInputElement;
      inp.value = text(c.value);
      if (live) inp.addEventListener("input", () => live("value", Number(inp.value)));
      el = inp;
      break;
    }
    default:
      el = h("div", { style: box + "border:1px solid #888;background:#fff" }, c.type === "Image" ? "" : c.type);
  }
  el.classList.add("vf-c");
  el.dataset.name = c.name;
  return el;
}

/**
 * A shown UserForm (UserForm.Show): a window over the app, its controls live. Typing, ticking and choosing go
 * to the runtime (setFormValue: Change, Click), buttons fire their _Click, the X asks the form (QueryClose). A
 * modal form shades the app. Returns the function that removes it (hideForm).
 */
export function showVbaForm(rt: () => VbaRuntime | null, form: VbaForm, modal: boolean): () => void {
  style();
  const d = form.design;
  const shade = modal ? h("div", { class: "vba-form-shade" }) : null;
  const body = h("div", { class: "vf-body", style: `width:${d.width * PT}px;height:${d.height * PT}px;background:${cssColor(d.backColor, "#f0f0f0")}` });
  const close = h("button", { type: "button", title: "Schließen" }, "✕");
  const win = h("div", { class: "vba-form", role: "dialog", "aria-label": d.caption }, h("div", { class: "vf-title" }, h("span", null, d.caption), close), body);
  const draw = () => {
    body.replaceChildren();
    for (const c of d.controls) {
      body.append(controlEl(c, (prop, value) => {
        const [name, p] = prop.includes("\u0000") ? prop.split("\u0000") : [c.name, prop];
        void rt()?.setFormValue(form, name, p === "listIndex" ? "ListIndex" : p === "value" ? "Value" : p, value);
      }, () => void rt()?.fire(form, `${c.name}_Click`)));
    }
    (win.firstElementChild!.firstElementChild as HTMLElement).textContent = d.caption;
  };
  form.onChange = () => draw();
  close.addEventListener("click", () => void rt()?.closeForm(form, 0));
  draw();
  if (shade) document.body.append(shade);
  document.body.append(win);
  const r = win.getBoundingClientRect();
  win.style.left = `${Math.max(0, (innerWidth - r.width) / 2)}px`;
  win.style.top = `${Math.max(0, (innerHeight - r.height) / 2)}px`;
  // The title bar moves the window.
  win.firstElementChild!.addEventListener("mousedown", (e) => {
    const m = e as MouseEvent, x0 = m.clientX - win.offsetLeft, y0 = m.clientY - win.offsetTop;
    const mv = (ev: MouseEvent) => { win.style.left = `${ev.clientX - x0}px`; win.style.top = `${ev.clientY - y0}px`; };
    const up = () => { removeEventListener("mousemove", mv); removeEventListener("mouseup", up); };
    addEventListener("mousemove", mv); addEventListener("mouseup", up);
  });
  return () => { win.remove(); shade?.remove(); form.onChange = undefined; };
}

/** The editor. Opens over the app; resolves when it is closed. */
export function openIde(ide: IdeHost, focus?: string): Promise<void> {
  style();
  const mods: VbaModule[] = ide.modules().map((m) => ({ ...m, code: visibleCode(m.code) }));
  if (!mods.length) mods.push({ name: "Modul1", kind: "standard", code: "" });
  const bps = new Map<string, Set<number>>();
  let cur = Math.max(0, focus ? mods.findIndex((m) => new RegExp(`Sub ${focus}\\b`, "i").test(m.code)) : 0);
  let showDesign = !!mods[cur].form;
  let rt: VbaRuntime | null = null;
  let paused: VbaBreak | null = null;
  let dirty = false;
  const shown = new Map<VbaForm, () => void>();

  const root = h("div", { class: "vba-ide", role: "dialog", "aria-label": "Makro-Editor" });
  const titleText = h("span");
  const status = h("span", { class: "vi-status" });
  const tree = h("div", { class: "vi-tree", "aria-label": "Projekt" });
  const center = h("div", { style: "min-height:0;display:grid" });
  const log = h("div", { class: "vi-log", "aria-label": "Direktbereich" });
  const imm = h("input", { class: "vi-imm", placeholder: "Anweisung (Eingabe führt sie aus)", "aria-label": "Direktbereich Eingabe" }) as HTMLInputElement;
  const locals = h("div", { class: "vi-locals", "aria-label": "Lokal" });
  const gutter = h("div", { class: "vi-gutter" });
  const area = h("textarea", { class: "vi-text", spellcheck: "false", "aria-label": "Code" }) as HTMLTextAreaElement;
  const curLine = h("div", { class: "vi-cur", style: "display:none" });
  const codeView = h("div", { class: "vi-code" }, gutter, area, curLine);

  const print = (s: string) => { log.append(document.createTextNode(s.endsWith("\n") ? s : s + "\n")); log.scrollTop = log.scrollHeight; };
  const syncCode = () => { if (!mods[cur].form || !showDesign) { if (mods[cur].code !== area.value) dirty = true; mods[cur].code = area.value; } };
  const procAtCaret = (): string | null => {
    const before = area.value.slice(0, area.selectionStart).split(/\r?\n/);
    for (let i = before.length - 1; i >= 0; i--) { const m = /^\s*(?:Public\s+|Private\s+)?Sub\s+(\w+)\s*\(\s*\)/i.exec(before[i]); if (m) return m[1]; if (/^\s*End\s+Sub/i.test(before[i]) && i < before.length - 1) break; }
    const first = /^\s*(?:Public\s+)?Sub\s+(\w+)\s*\(\s*\)/im.exec(area.value);
    return first ? first[1] : null;
  };

  const drawGutter = () => {
    const n = area.value.split(/\r?\n/).length;
    const set = bps.get(mods[cur].name) ?? new Set<number>();
    gutter.replaceChildren(...Array.from({ length: n }, (_, i) => {
      const d = h("div", { class: `${set.has(i + 1) ? "bp" : ""}${paused && paused.module === mods[cur].name && paused.line === i + 1 ? " cur" : ""}` }, String(i + 1));
      d.addEventListener("click", () => toggleBp(i + 1));
      return d;
    }));
    gutter.scrollTop = area.scrollTop;
    if (paused && paused.module.toLowerCase() === mods[cur].name.toLowerCase()) { curLine.style.display = ""; curLine.style.top = `${(paused.line - 1) * LINE - area.scrollTop}px`; } else curLine.style.display = "none";
  };
  area.addEventListener("scroll", () => { gutter.scrollTop = area.scrollTop; drawGutter(); });
  area.addEventListener("input", () => { dirty = true; drawGutter(); });
  area.addEventListener("keydown", (e) => {
    if (e.key === "Tab") { e.preventDefault(); const s = area.selectionStart; area.value = area.value.slice(0, s) + "    " + area.value.slice(area.selectionEnd); area.selectionStart = area.selectionEnd = s + 4; drawGutter(); }
  });
  const caretLine = () => area.value.slice(0, area.selectionStart).split(/\r?\n/).length;
  const toggleBp = (line = caretLine()) => {
    const name = mods[cur].name;
    const set = bps.get(name) ?? new Set<number>();
    if (set.has(line)) set.delete(line); else set.add(line);
    bps.set(name, set);
    rt?.toggleBreakpoint(name, line);
    drawGutter();
  };

  const drawTree = () => {
    const group = (label: string, pick: (m: VbaModule) => boolean) => {
      const items = mods.map((m, i) => [m, i] as const).filter(([m]) => pick(m));
      if (!items.length) return [];
      return [h("div", { class: "grp" }, label), ...items.map(([m, i]) => {
        const it = h("div", { class: `it${i === cur ? " on" : ""}`, tabindex: "0" }, m.name);
        it.addEventListener("click", () => select(i, !!m.form));
        it.addEventListener("dblclick", () => select(i, false));
        return it;
      })];
    };
    tree.replaceChildren(...group("Formulare", (m) => !!m.form), ...group("Module", (m) => m.kind === "standard"), ...group("Klassenmodule", (m) => m.kind === "class" && !m.form));
  };
  const select = (i: number, design: boolean) => { syncCode(); cur = i; showDesign = design && !!mods[i].form; draw(); };

  const draw = () => {
    const m = mods[cur];
    titleText.textContent = `${ide.title} - Makro-Editor - [${m.name} (${showDesign ? "UserForm" : "Code"})]`;
    drawTree();
    if (m.form && showDesign) center.replaceChildren(designer(m));
    else { area.value = m.code; center.replaceChildren(codeView); drawGutter(); }
  };

  // ----- the UserForm designer -------------------------------------------------------------------------------
  let tool: VbaControlType | null = null;
  let selected: string | null = null;
  const TOOLS: [VbaControlType, string][] = [["Label", "Bezeichnungsfeld"], ["TextBox", "Textfeld"], ["ComboBox", "Kombinationsfeld"], ["ListBox", "Listenfeld"], ["CheckBox", "Kontrollkästchen"],
    ["OptionButton", "Optionsfeld"], ["ToggleButton", "Umschaltfeld"], ["Frame", "Rahmen"], ["CommandButton", "Befehlsschaltfläche"], ["ScrollBar", "Bildlaufleiste"], ["SpinButton", "Drehfeld"], ["Image", "Anzeige"]];
  const designer = (m: VbaModule): HTMLElement => {
    const f = m.form!;
    const tools = h("div", { class: "vi-tools", "aria-label": "Werkzeugsammlung" }, h("b", null, "Werkzeugsammlung"),
      ...TOOLS.map(([t, label]) => { const b = h("button", { type: "button", class: tool === t ? "on" : "" }, label); b.addEventListener("click", () => { tool = tool === t ? null : t; draw(); }); return b; }));
    const body = h("div", { class: "vf-body", style: `width:${f.width * PT}px;height:${f.height * PT}px;background:${cssColor(f.backColor, "#f0f0f0")};background-image:radial-gradient(#999 0.6px,transparent 0.6px);background-size:${6 * PT}px ${6 * PT}px` });
    const win = h("div", { class: "vba-form", style: "position:relative;z-index:auto;display:inline-block" }, h("div", { class: "vf-title" }, h("span", null, f.caption), h("button", { type: "button", tabindex: "-1" }, "✕")), body);
    const snap = (v: number) => Math.round(v / 6) * 6;
    const uniq = (t: VbaControlType) => { let k = 1; while (allControls(f.controls).some((c) => c.name.toLowerCase() === `${t}${k}`.toLowerCase())) k++; return `${t}${k}`; };
    for (const c of f.controls) {
      const el = controlEl(c, null, null);
      if (c.name === selected) { el.classList.add("vf-sel"); const g = h("div", { class: "vf-grip" }); el.append(g); el.style.overflow = "visible";
        g.addEventListener("mousedown", (e) => { e.stopPropagation(); const x0 = (e as MouseEvent).clientX, y0 = (e as MouseEvent).clientY, w0 = c.width, h0 = c.height;
          const mv = (ev: MouseEvent) => { c.width = Math.max(6, snap(w0 + (ev.clientX - x0) / PT)); c.height = Math.max(6, snap(h0 + (ev.clientY - y0) / PT)); el.style.width = `${c.width * PT}px`; el.style.height = `${c.height * PT}px`; };
          const up = () => { removeEventListener("mousemove", mv); removeEventListener("mouseup", up); dirty = true; draw(); };
          addEventListener("mousemove", mv); addEventListener("mouseup", up); }); }
      el.addEventListener("mousedown", (e) => {
        e.stopPropagation(); e.preventDefault();
        const was = selected;
        selected = c.name;
        const x0 = (e as MouseEvent).clientX, y0 = (e as MouseEvent).clientY, l0 = c.left, t0 = c.top;
        let moved = false;
        const mv = (ev: MouseEvent) => { moved = true; c.left = Math.max(0, snap(l0 + (ev.clientX - x0) / PT)); c.top = Math.max(0, snap(t0 + (ev.clientY - y0) / PT)); el.style.left = `${c.left * PT}px`; el.style.top = `${c.top * PT}px`; };
        // Drawn again only when something changed: a double-click needs the same element twice.
        const up = () => { removeEventListener("mousemove", mv); removeEventListener("mouseup", up); if (moved) dirty = true; if (moved || was !== c.name) draw(); };
        addEventListener("mousemove", mv); addEventListener("mouseup", up);
      });
      // A double-click opens the code at the control's Click (or Change) procedure, made when missing.
      el.addEventListener("dblclick", () => { const ev = ["TextBox", "ComboBox", "SpinButton", "ScrollBar"].includes(c.type) ? "Change" : "Click"; openProc(m, `${c.name}_${ev}`); });
      body.append(el);
    }
    body.addEventListener("mousedown", (e) => {
      const r = body.getBoundingClientRect();
      if (tool) {
        const c = newControl(tool, uniq(tool), snap(((e as MouseEvent).clientX - r.left) / PT), snap(((e as MouseEvent).clientY - r.top) / PT));
        f.controls.push(c); selected = c.name; tool = null; dirty = true;
      } else selected = null;
      draw();
    });
    body.addEventListener("dblclick", (e) => { if (e.target === body) openProc(m, "UserForm_Click"); });
    const target = f.controls.find((c) => c.name === selected);
    const rows: [string, () => string, (v: string) => void][] = target ? [
      ["(Name)", () => target.name, (v) => { if (/^[A-Za-z]\w*$/.test(v)) { selected = target.name = v; } }],
      ...(target.caption !== undefined ? [["Caption", () => target.caption ?? "", (v: string) => { target.caption = v; }] as [string, () => string, (v: string) => void]] : []),
      ["Left", () => String(target.left), (v) => { target.left = Number(v) || 0; }], ["Top", () => String(target.top), (v) => { target.top = Number(v) || 0; }],
      ["Width", () => String(target.width), (v) => { target.width = Number(v) || target.width; }], ["Height", () => String(target.height), (v) => { target.height = Number(v) || target.height; }],
      ...(target.value !== undefined && target.type !== "ListBox" ? [["Value", () => text(target.value), (v: string) => { target.value = target.type === "CheckBox" || target.type === "OptionButton" || target.type === "ToggleButton" ? /^(wahr|true|-1|1)$/i.test(v) : target.type === "SpinButton" || target.type === "ScrollBar" ? Number(v) || 0 : v; }] as [string, () => string, (v: string) => void]] : []),
      ["Enabled", () => (target.enabled === false ? "Falsch" : "Wahr"), (v) => { target.enabled = !/^(falsch|false|0)$/i.test(v); }],
      ["Visible", () => (target.visible === false ? "Falsch" : "Wahr"), (v) => { target.visible = !/^(falsch|false|0)$/i.test(v); }],
    ] : [
      ["(Name)", () => m.name, (v) => { if (/^[A-Za-z]\w*$/.test(v)) { m.name = v; f.name = v; } }], ["Caption", () => f.caption, (v) => { f.caption = v; }],
      ["Width", () => String(f.width), (v) => { f.width = Number(v) || f.width; }], ["Height", () => String(f.height), (v) => { f.height = Number(v) || f.height; }],
    ];
    const props = h("div", { class: "vi-props", "aria-label": "Eigenschaften" }, h("h4", { style: "margin:0;padding:3px 6px;background:#e8e8e8" }, `Eigenschaften - ${target ? target.name : m.name}`),
      h("table", null, ...rows.map(([k, get, set]) => {
        const inp = h("input", { value: get(), "aria-label": k }) as HTMLInputElement;
        inp.addEventListener("change", () => { set(inp.value); dirty = true; draw(); });
        return h("tr", null, h("td", null, k), h("td", null, inp));
      })));
    return h("div", { class: "vi-design" }, tools, h("div", { class: "vi-canvas" }, win), props);
  };
  const openProc = (m: VbaModule, proc: string) => {
    if (!new RegExp(`Sub\\s+${proc}\\b`, "i").test(m.code)) m.code = `${m.code.replace(/\s*$/, "")}${m.code.trim() ? "\r\n\r\n" : ""}Private Sub ${proc}()\r\n\r\nEnd Sub\r\n`;
    dirty = true;
    showDesign = false;
    draw();
    const i = m.code.search(new RegExp(`Sub\\s+${proc}\\b`, "i"));
    const lineEnd = m.code.indexOf("\n", i);
    area.focus();
    area.selectionStart = area.selectionEnd = lineEnd + 1;
  };

  // ----- running and debugging -------------------------------------------------------------------------------
  const drawLocals = () => {
    if (!rt || !paused) { locals.replaceChildren(); return; }
    const vars = rt.locals();
    locals.replaceChildren(h("table", null, h("tr", null, h("th", null, "Ausdruck"), h("th", null, "Wert"), h("th", null, "Typ")),
      ...vars.map((v) => h("tr", null, h("td", null, v.name), h("td", null, typeof v.value === "object" && v.value !== null ? "" : text(v.value)), h("td", null, v.type)))));
  };
  const onBreak = (info: VbaBreak) => {
    paused = info;
    const i = mods.findIndex((m) => m.name.toLowerCase() === info.module.toLowerCase());
    if (i >= 0 && (i !== cur || showDesign)) { syncCode(); cur = i; showDesign = false; draw(); }
    status.textContent = `[Unterbrechen] ${info.proc}, Zeile ${info.line}`;
    drawGutter();
    drawLocals();
  };
  const step = (mode: VbaStepMode) => {
    if (!rt || !paused) return;
    paused = null; status.textContent = "[Ausführen]"; drawGutter(); drawLocals();
    rt.resume(mode);
  };
  const run = async () => {
    if (paused) { step("continue"); return; }
    if (rt) return;
    syncCode();
    const name = procAtCaret();
    if (!name) { print("Keine Prozedur ohne Argumente am Cursor."); return; }
    try {
      rt = new VbaRuntime(ide.vbaHost({
        debugPrint: (s) => print(s),
        debugBreak: (info) => { onBreak(info); },
        showForm: (form, modal) => { shown.set(form, showVbaForm(() => rt, form, modal)); },
        hideForm: (form) => { shown.get(form)?.(); shown.delete(form); },
      }), mods);
      for (const [mod, set] of bps) for (const l of set) rt.setBreakpoint(mod, l);
    } catch (e) { print(`Fehler beim Kompilieren: ${(e as Error).message}`); rt = null; return; }
    status.textContent = "[Ausführen]";
    try { await rt.run(name); }
    catch (e) { const err = e as { number?: number; message?: string; line?: number }; print(`Laufzeitfehler '${err.number ?? 5}': ${err.message ?? e}${err.line ? ` (Zeile ${err.line})` : ""}`); }
    finally {
      // A shown modeless form keeps the runtime until it closes.
      if (!shown.size) { rt = null; status.textContent = ""; }
      paused = null; drawGutter(); drawLocals();
    }
  };
  const reset = async () => { if (!rt) return; if (paused) step("stop"); await rt.reset(); for (const close of shown.values()) close(); shown.clear(); rt = null; paused = null; status.textContent = ""; drawGutter(); drawLocals(); };
  imm.addEventListener("keydown", (e) => {
    if (e.key !== "Enter" || !imm.value.trim()) return;
    const line = imm.value;
    imm.value = "";
    print(line);
    if (rt && paused) void rt.immediate(line).catch((er) => print(`Fehler: ${(er as Error).message}`)).then(() => drawLocals());
    else {
      // Outside a run: the statement runs in a module of its own (as Office's Direktbereich does).
      try {
        const once = new VbaRuntime(ide.vbaHost({ debugPrint: (s) => print(s) }), [...mods, { name: "Direktbereich", kind: "standard", code: `Sub DirektbereichAusfuehren()\r\n${line.replace(/^\?/, "Debug.Print ")}\r\nEnd Sub` }]);
        void once.run("DirektbereichAusfuehren").catch((er) => print(`Laufzeitfehler: ${(er as Error).message}`));
      } catch (er) { print(`Fehler: ${(er as Error).message}`); }
    }
  });

  // ----- inserting, saving, closing --------------------------------------------------------------------------
  const free = (base: string) => { let k = 1; while (mods.some((m) => m.name.toLowerCase() === `${base}${k}`.toLowerCase())) k++; return `${base}${k}`; };
  const insert = (kind: "standard" | "class" | "form") => {
    syncCode();
    const name = free(kind === "standard" ? "Modul" : kind === "class" ? "Klasse" : "UserForm");
    mods.push(kind === "form" ? { name, kind: "class", code: "", form: newForm(name) } : { name, kind, code: "Option Explicit\r\n" });
    cur = mods.length - 1; showDesign = kind === "form"; dirty = true;
    draw();
  };
  const save = () => { syncCode(); ide.save(mods.map((m) => ({ ...m }))); dirty = false; status.textContent = "Gespeichert."; };
  let done: () => void = () => {};
  const close = async () => { await reset(); if (dirty) save(); root.remove(); removeEventListener("keydown", keys, true); done(); };
  const keys = (e: KeyboardEvent) => {
    if (!root.isConnected) return;
    const k = e.key;
    const handled = k === "F5" ? (void run(), true) : k === "F8" && e.ctrlKey && e.shiftKey ? (step("out"), true) : k === "F8" && e.shiftKey ? (step("over"), true) : k === "F8" ? (paused ? step("into") : (rt?.pause(), void run()), true)
      : k === "F9" ? (toggleBp(), true) : k === "s" && e.ctrlKey ? (save(), true) : k === "Pause" && e.ctrlKey ? (rt?.pause(), true) : k === "F7" ? (select(cur, false), true) : false;
    if (handled) { e.preventDefault(); e.stopPropagation(); }
  };
  addEventListener("keydown", keys, true);

  const btn = (label: string, title: string, f: () => void) => { const b = h("button", { type: "button", title }, label); b.addEventListener("click", f); return b; };
  const insertMenu = h("select", { "aria-label": "Einfügen" }, h("option", { value: "" }, "Einfügen"), h("option", { value: "form" }, "UserForm"), h("option", { value: "standard" }, "Modul"), h("option", { value: "class" }, "Klassenmodul")) as HTMLSelectElement;
  insertMenu.addEventListener("change", () => { if (insertMenu.value) insert(insertMenu.value as "form" | "standard" | "class"); insertMenu.value = ""; });
  root.append(
    h("div", { class: "vi-title" }, titleText, h("span", null, btn("✕", "Schließen und zurück zu RaabePresent", () => void close()))),
    h("div", { class: "vi-bar" }, btn("Speichern", "Speichern (Strg+S)", save), h("span", { class: "sep" }), insertMenu, h("span", { class: "sep" }),
      btn("▶ Ausführen", "Sub/UserForm ausführen (F5)", () => void run()), btn("⏸ Unterbrechen", "Unterbrechen (Strg+Pause)", () => rt?.pause()), btn("■ Zurücksetzen", "Zurücksetzen", () => void reset()),
      h("span", { class: "sep" }), btn("Haltepunkt", "Haltepunkt ein/aus (F9)", () => toggleBp()), btn("Einzelschritt", "Einzelschritt (F8)", () => step("into")),
      btn("Prozedurschritt", "Prozedurschritt (Umschalt+F8)", () => step("over")), btn("Prozedur abschließen", "Prozedur abschließen (Strg+Umschalt+F8)", () => step("out")),
      h("span", { class: "sep" }), btn("Code anzeigen", "Code anzeigen (F7)", () => select(cur, false)), btn("Objekt anzeigen", "Objekt anzeigen", () => select(cur, true)), status),
    h("div", { class: "vi-main" }, tree, center),
    h("div", { class: "vi-bottom" }, h("div", { class: "vi-pane" }, h("h4", null, "Direktbereich"), log, imm), h("div", { class: "vi-pane" }, h("h4", null, "Lokal"), locals)),
  );
  document.body.append(root);
  draw();
  return new Promise((res) => { done = res; });
}
