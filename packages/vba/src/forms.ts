// UserForms (D-080): the plain description of a form and its controls that the app draws and its designer edits
// (VbaFormDesign, VbaControlDesign), and the objects VBA code sees while a form is loaded (the form's members,
// Controls, each control's properties and methods). The interpreter (interp.ts) makes a form module's instance,
// the app shows it through VbaHost.showForm and reports the user's input with VbaRuntime.setFormValue and
// VbaRuntime.fire. Changes made by VBA reach the app through VbaForm.onChange. Sizes are in points, colours are
// OLE_COLORs as VBA uses them (&H8000000F is the system colour "button face", otherwise &H00BBGGRR).

import { NOT_A_METHOD, VbaError, isArr, toBool, toNum, toStr, type VbaArray, type VbaObject, type VbaValue } from "./values";

export type VbaControlType = "CommandButton" | "Label" | "TextBox" | "CheckBox" | "OptionButton" | "ToggleButton" | "ListBox" | "ComboBox" | "Frame" | "Image" | "SpinButton" | "ScrollBar" | "MultiPage" | "Page" | "TabStrip";
export interface VbaFont { name: string; size: number; bold: boolean; italic: boolean; underline: boolean; strikethrough: boolean }

/** A control of a form, as the designer edits it and the app draws it; the runtime keeps its current values here. */
export interface VbaControlDesign {
  type: VbaControlType;
  name: string;
  left: number; top: number; width: number; height: number;
  caption?: string;
  /** TextBox and ComboBox: the text; CheckBox, OptionButton, ToggleButton: True, False or Null; SpinButton and
   *  ScrollBar: the number; ListBox: the bound column of the selected row (Null when none); MultiPage: the page index. */
  value?: VbaValue;
  visible?: boolean; enabled?: boolean; locked?: boolean;
  tabIndex?: number; tabStop?: boolean;
  backColor?: number; foreColor?: number; borderColor?: number;
  /** 0 transparent, 1 opaque. */
  backStyle?: number;
  /** 0 none, 1 single. */
  borderStyle?: number;
  /** 0 flat, 1 raised, 2 sunken, 3 etched, 6 bump. */
  specialEffect?: number;
  font?: VbaFont;
  controlTipText?: string; tag?: string; groupName?: string; accelerator?: string;
  multiLine?: boolean; wordWrap?: boolean; autoSize?: boolean; passwordChar?: string; maxLength?: number;
  /** 1 left, 2 centre, 3 right. */
  textAlign?: number;
  min?: number; max?: number; smallChange?: number; largeChange?: number;
  /** -1 automatic, 0 vertical, 1 horizontal. */
  orientation?: number;
  columnCount?: number;
  /** 0 single, 1 multi, 2 extended. */
  multiSelect?: number;
  /** ComboBox: 0 drop-down combo, 2 drop-down list. */
  style?: number;
  /** Runtime only (not stored in the file, as in Office): the rows of a ListBox or ComboBox (one string per column). */
  list?: string[][];
  listIndex?: number;
  selected?: boolean[];
  /** Frame and Page: their controls. */
  controls?: VbaControlDesign[];
  /** MultiPage: its pages (type Page). */
  pages?: VbaControlDesign[];
  /** Bytes of the file we do not interpret (a picture, unknown properties), written back unchanged. */
  raw?: VbaRawControl;
}
/** What the file reader keeps of a control to write it back as it was. */
export interface VbaRawControl { [key: string]: unknown }

export interface VbaFormDesign {
  name: string;
  caption: string;
  /** The inside (client) size in points; VBA's Width and Height add the window frame. */
  width: number; height: number;
  backColor?: number; foreColor?: number;
  font?: VbaFont;
  /** 0 manual, 1 centred on the owner (the default), 2 centred on the screen, 3 Windows' default. */
  startUpPosition?: number;
  left?: number; top?: number;
  tag?: string;
  controls: VbaControlDesign[];
  /** What the file reader keeps to write the form back unchanged when it was not edited. */
  raw?: VbaRawControl;
}

/** A loaded form as the app sees it: the live design (current values) and whether it is shown. */
export interface VbaForm {
  readonly name: string;
  readonly design: VbaFormDesign;
  readonly visible: boolean;
  readonly modal: boolean;
  /** Set by the app: called after VBA changed the form (control null) or a control (prop: "caption", "value", "list" …). */
  onChange?: (control: VbaControlDesign | null, prop: string) => void;
}

// ----- defaults --------------------------------------------------------------------------------------------

export const SYS_BUTTONFACE = 0x8000000f, SYS_BUTTONTEXT = 0x80000012, SYS_WINDOW = 0x80000005, SYS_WINDOWTEXT = 0x80000008;
export const defaultFont = (): VbaFont => ({ name: "Tahoma", size: 8, bold: false, italic: false, underline: false, strikethrough: false });
/** The window frame VBA's Width and Height add to the inside size: a new form is 240 x 180 points outside and
 *  235.5 x 160.5 inside (ClientWidth 4710, ClientHeight 3210 twips, as in [MS-OVBA] 3.1.7). */
export const FORM_FRAME = { width: 4.5, height: 19.5 };

const SIZES: Record<VbaControlType, [number, number]> = {
  CommandButton: [72, 24], Label: [72, 18], TextBox: [72, 18], CheckBox: [108, 18], OptionButton: [108, 18], ToggleButton: [72, 24],
  ListBox: [72, 72], ComboBox: [72, 18], Frame: [144, 72], Image: [72, 72], SpinButton: [12.75, 25.5], ScrollBar: [12.75, 63.75],
  MultiPage: [150, 110], Page: [150, 110], TabStrip: [150, 110],
};
const BUTTONISH = new Set<VbaControlType>(["CommandButton", "Label", "CheckBox", "OptionButton", "ToggleButton", "Frame", "Image", "SpinButton", "ScrollBar", "MultiPage", "Page", "TabStrip"]);

/** A control with Office's defaults for its type (the designer's new control). */
export function newControl(type: VbaControlType, name: string, left = 0, top = 0): VbaControlDesign {
  const [width, height] = SIZES[type];
  const c: VbaControlDesign = { type, name, left, top, width, height, visible: true, enabled: true, tabStop: true,
    backColor: BUTTONISH.has(type) ? SYS_BUTTONFACE : SYS_WINDOW, foreColor: BUTTONISH.has(type) ? SYS_BUTTONTEXT : SYS_WINDOWTEXT, font: defaultFont() };
  switch (type) {
    case "CommandButton": case "Label": case "Frame": case "Page": c.caption = name; break;
    case "CheckBox": case "OptionButton": case "ToggleButton": c.caption = name; c.value = false; break;
    case "TextBox": c.value = ""; c.specialEffect = 2; break;
    case "ComboBox": c.value = ""; c.specialEffect = 2; c.list = []; c.listIndex = -1; c.style = 0; c.columnCount = 1; break;
    case "ListBox": c.value = null; c.specialEffect = 2; c.list = []; c.listIndex = -1; c.multiSelect = 0; c.selected = []; c.columnCount = 1; break;
    case "SpinButton": c.value = 0; c.min = 0; c.max = 100; c.smallChange = 1; c.orientation = -1; break;
    case "ScrollBar": c.value = 0; c.min = 0; c.max = 32767; c.smallChange = 1; c.largeChange = 1; c.orientation = -1; break;
    case "MultiPage": c.value = 0; c.pages = [newControl("Page", "Page1"), newControl("Page", "Page2")]; break;
    case "Image": c.borderStyle = 1; break;
  }
  if (type === "Label") { c.wordWrap = true; c.backStyle = 1; c.tabStop = false; }
  // An Image has no text colour; Image, SpinButton and ScrollBar show no text.
  if (type === "Image") delete c.foreColor;
  if (type === "Image" || type === "SpinButton" || type === "ScrollBar") delete c.font;
  if (type === "Frame" || type === "Page") c.controls = [];
  return c;
}
/** A new, empty form with Office's defaults. */
export function newForm(name: string): VbaFormDesign {
  return { name, caption: name, width: 240 - FORM_FRAME.width, height: 180 - FORM_FRAME.height, backColor: SYS_BUTTONFACE, foreColor: SYS_BUTTONTEXT, font: defaultFont(), startUpPosition: 1, controls: [] };
}

/** A deep copy of a design (the runtime's live copy; raw data is shared, it is never changed). */
export function cloneDesign<T extends VbaFormDesign | VbaControlDesign>(d: T): T {
  const copy = (v: unknown): unknown => {
    if (Array.isArray(v)) return v.map(copy);
    if (v instanceof Uint8Array || v instanceof Map) return v;
    if (v && typeof v === "object") { const o: Record<string, unknown> = {}; for (const [k, x] of Object.entries(v)) o[k] = k === "raw" ? x : copy(x); return o; }
    return v;
  };
  return copy(d) as T;
}

/** Every control of a form or container, depth first (the order of UserForm.Controls). */
export function allControls(list: VbaControlDesign[]): VbaControlDesign[] {
  const out: VbaControlDesign[] = [];
  const walk = (l: VbaControlDesign[]) => { for (const c of l) { out.push(c); if (c.pages) walk(c.pages); if (c.controls) walk(c.controls); } };
  walk(list);
  return out;
}

// ----- runtime ----------------------------------------------------------------------------------------------

/** How a loaded form reaches its module's event procedures (CommandButton1_Click …). */
export interface FormEvents { event(proc: string, args: VbaValue[]): Promise<void> }

const low = (s: string) => s.toLowerCase();
const PROGIDS: Record<string, VbaControlType> = {
  "forms.commandbutton.1": "CommandButton", "forms.label.1": "Label", "forms.textbox.1": "TextBox", "forms.checkbox.1": "CheckBox",
  "forms.optionbutton.1": "OptionButton", "forms.togglebutton.1": "ToggleButton", "forms.listbox.1": "ListBox", "forms.combobox.1": "ComboBox",
  "forms.frame.1": "Frame", "forms.image.1": "Image", "forms.spinbutton.1": "SpinButton", "forms.scrollbar.1": "ScrollBar", "forms.multipage.1": "MultiPage",
};

/** A loaded form: its live design, its controls as VBA objects and its shown state. */
export class FormState implements VbaForm {
  readonly design: VbaFormDesign;
  visible = false;
  modal = false;
  onChange?: (control: VbaControlDesign | null, prop: string) => void;
  /** Controls by lower-case name. */
  readonly byName = new Map<string, ControlObj>();
  private closeWaiters: (() => void)[] = [];
  constructor(design: VbaFormDesign, readonly events: FormEvents) {
    this.design = cloneDesign(design);
    this.index();
  }
  get name(): string { return this.design.name; }
  index(): void {
    this.byName.clear();
    const walk = (l: VbaControlDesign[], parent: VbaControlDesign[]) => { for (const c of l) { this.byName.set(low(c.name), new ControlObj(this, c, parent)); if (c.pages) walk(c.pages, c.pages); if (c.controls) walk(c.controls, c.controls); } };
    walk(this.design.controls, this.design.controls);
  }
  control(name: string): ControlObj | undefined { return this.byName.get(low(name)); }
  changed(c: VbaControlDesign | null, prop: string): void { try { this.onChange?.(c, prop); } catch { /* the app's drawing must not stop the macro */ } }
  /** Resolves when the form is hidden or unloaded (a modal Show waits for it). */
  waitHidden(): Promise<void> { return this.visible ? new Promise((r) => this.closeWaiters.push(r)) : Promise.resolve(); }
  hidden(): void { this.visible = false; const w = this.closeWaiters; this.closeWaiters = []; for (const r of w) r(); }

  /** UserForm.Controls: every control, 0-based, by index or name; Add and Remove for controls made by code. */
  controlsObject(list: () => VbaControlDesign[], container: VbaControlDesign[]): VbaObject {
    const objs = () => allControlsOf(list()).map((c) => this.control(c.name)!).filter(Boolean);
    const item = (k: VbaValue) => {
      const l = objs();
      const v = typeof k === "number" ? l[k] : l.find((o) => low(o.d.name) === low(toStr(k)));
      if (!v) throw new VbaError("Das angegebene Objekt wurde nicht gefunden", -2147024809);
      return v;
    };
    return {
      kind: "object",
      typeName: () => "Controls",
      items: () => objs(),
      default: (a) => item(a[0]),
      get: (n, a) => { const k = low(n); if (k === "count") return objs().length; if (k === "item") return item(a[0]); throw new VbaError(`Objekt unterstützt diese Eigenschaft oder Methode nicht: ${n}`, 438); },
      argNames: (n) => (low(n) === "add" ? ["bstrProgID", "Name", "Visible"] : undefined),
      call: (n, a) => {
        const k = low(n);
        if (k === "add") {
          const type = PROGIDS[low(toStr(a[0]))];
          if (!type) throw new VbaError(`Ungültige Klassenzeichenfolge: ${toStr(a[0])}`, -2147221005);
          let name = a[1] === undefined ? "" : toStr(a[1]);
          if (!name) { let i = 1; while (this.control(type + i)) i++; name = type + i; }
          const c = newControl(type, name);
          if (a[2] !== undefined) c.visible = toBool(a[2]);
          container.push(c);
          this.index();
          this.changed(null, "controls");
          return this.control(name)!;
        }
        if (k === "remove") {
          const o = item(a[0]);
          const at = o.parent.indexOf(o.d);
          if (at >= 0) o.parent.splice(at, 1);
          this.index();
          this.changed(null, "controls");
          return undefined;
        }
        if (k === "clear") { container.splice(0); this.index(); this.changed(null, "controls"); return undefined; }
        if (k === "count" || k === "item") return NOT_A_METHOD;
        return NOT_A_METHOD;
      },
    };
  }
}
function allControlsOf(list: VbaControlDesign[]): VbaControlDesign[] { return allControls(list); }

/** The font of a form or control (Font.Size, Font.Bold …). */
export function fontObject(get: () => VbaFont, changed: () => void): VbaObject {
  const key = (n: string): keyof VbaFont | null => { const k = low(n); return k === "name" ? "name" : k === "size" ? "size" : k === "bold" ? "bold" : k === "italic" ? "italic" : k === "underline" ? "underline" : k === "strikethrough" ? "strikethrough" : null; };
  return {
    kind: "object",
    typeName: () => "Font",
    default: () => get().name,
    get: (n) => { const k = key(n); if (k) return get()[k]; if (low(n) === "weight") return get().bold ? 700 : 400; throw new VbaError(`Objekt unterstützt diese Eigenschaft oder Methode nicht: ${n}`, 438); },
    set: (n, v) => {
      const f = get(), k = key(n);
      if (k === "name") f.name = toStr(v); else if (k === "size") f.size = toNum(v); else if (k) (f as unknown as Record<string, boolean>)[k] = toBool(v);
      else if (low(n) === "weight") f.bold = toNum(v) >= 600;
      else throw new VbaError(`Eigenschaft kann nicht gesetzt werden: ${n}`, 438);
      changed();
    },
  };
}

/** Design properties VBA reads and writes by name (lower case), beyond the special ones handled in ControlObj. */
const PLAIN = ["name", "caption", "left", "top", "width", "height", "visible", "enabled", "locked", "tabIndex", "tabStop", "backColor", "foreColor", "borderColor", "backStyle", "borderStyle", "specialEffect",
  "controlTipText", "tag", "groupName", "accelerator", "multiLine", "wordWrap", "autoSize", "passwordChar", "maxLength", "textAlign", "min", "max", "smallChange", "largeChange", "orientation", "columnCount", "multiSelect", "style"] as const;
const PLAIN_KEY = new Map<string, keyof VbaControlDesign>(PLAIN.map((k) => [low(k), k]));
const BOOL_PROPS = new Set(["visible", "enabled", "locked", "tabstop", "multiline", "wordwrap", "autosize"]);
const STR_PROPS = new Set(["name", "caption", "controltiptext", "tag", "groupname", "accelerator", "passwordchar"]);

/** A control as VBA sees it. Setting Value, Text or ListIndex fires the control's Change (and Click) events. */
export class ControlObj implements VbaObject {
  readonly kind = "object";
  constructor(readonly form: FormState, readonly d: VbaControlDesign, readonly parent: VbaControlDesign[]) {}
  typeName(): string { return this.d.type; }
  default(args: VbaValue[]): VbaValue { return args.length && (this.d.type === "ListBox" || this.d.type === "ComboBox") ? this.listGet(args) : this.value(); }
  private changed(prop: string) { this.form.changed(this.d, prop); }
  private isList() { return this.d.type === "ListBox" || this.d.type === "ComboBox"; }

  value(): VbaValue {
    const d = this.d;
    if (d.type === "ListBox") { const i = d.listIndex ?? -1; return i >= 0 && d.list?.[i] ? d.list[i][0] ?? null : null; }
    if (d.type === "CommandButton") return false;
    if (d.type === "Label" || d.type === "Frame" || d.type === "Image" || d.type === "Page") throw new VbaError("Objekt unterstützt diese Eigenschaft oder Methode nicht: Value", 438);
    return d.value;
  }
  text(): string {
    const d = this.d;
    if (d.type === "ListBox") { const v = this.value(); return v === null ? "" : toStr(v); }
    if (d.type === "TextBox" || d.type === "ComboBox") return toStr(d.value);
    throw new VbaError("Objekt unterstützt diese Eigenschaft oder Methode nicht: Text", 438);
  }
  private listGet(args: VbaValue[]): VbaValue {
    const d = this.d;
    if (!args.length) {
      // List without index: the rows as a 2-D array.
      const rows = d.list ?? [];
      const cols = Math.max(1, d.columnCount ?? 1);
      return { kind: "array", lb: 0, items: rows.map((r) => ({ kind: "array", lb: 0, items: Array.from({ length: cols }, (_, j) => r[j] ?? null) })) } as VbaArray;
    }
    const row = d.list?.[toNum(args[0])];
    if (!row) throw new VbaError("Eigenschaft List konnte nicht abgerufen werden. Ungültiges Argument.", 381);
    const v = row[args[1] === undefined ? 0 : toNum(args[1])];
    return v === undefined ? null : v;
  }

  get(name: string, args: VbaValue[]): VbaValue {
    const k = low(name), d = this.d;
    switch (k) {
      case "value": return this.value();
      case "text": return this.text();
      case "font": return fontObject(() => (d.font ??= { name: "Tahoma", size: 8, bold: false, italic: false, underline: false, strikethrough: false }), () => this.changed("font"));
      case "controls": if (d.controls) return this.form.controlsObject(() => d.controls!, d.controls); break;
      case "pages": if (d.pages) return this.pagesObject(); break;
      case "listcount": if (this.isList()) return d.list?.length ?? 0; break;
      case "listindex": if (this.isList()) return d.listIndex ?? -1; break;
      case "list": if (this.isList()) return this.listGet(args); break;
      case "column": if (this.isList()) return args.length >= 2 ? this.listGet([args[1], args[0]]) : this.listGet([d.listIndex ?? 0, args[0]]); break;
      case "selected": if (this.isList()) return d.multiSelect ? !!d.selected?.[toNum(args[0])] : (d.listIndex ?? -1) === toNum(args[0]); break;
      case "index": if (d.type === "Page") { const i = this.parent.indexOf(d); return i; } break;
      case "parent": return null;
      case "object": return this;
      case "selecteditem": if (d.type === "MultiPage") return this.pageObject(toNum(d.value ?? 0)); break;
    }
    const pk = PLAIN_KEY.get(k);
    if (pk) { const v = d[pk]; return v === undefined ? (BOOL_PROPS.has(k) ? k === "visible" || k === "enabled" || k === "tabstop" : STR_PROPS.has(k) ? "" : 0) : (v as VbaValue); }
    const r = this.call(name, args);
    if (r !== NOT_A_METHOD) return r;
    throw new VbaError(`Objekt unterstützt diese Eigenschaft oder Methode nicht: ${name}`, 438);
  }

  /** A property set without events (the app's own changes, and what the runtime sets before it fires events).
   *  Returns the names of the events the change calls for. */
  setQuiet(name: string, v: VbaValue, args: VbaValue[] = []): string[] {
    const k = low(name), d = this.d;
    const events: string[] = [];
    switch (k) {
      case "value": case "text": {
        if (d.type === "ListBox") { const i = (d.list ?? []).findIndex((r) => r[0] === toStr(v)); return this.setQuiet("listindex", v === null ? -1 : i); }
        if (k === "text" && d.type !== "TextBox" && d.type !== "ComboBox") throw new VbaError("Objekt unterstützt diese Eigenschaft oder Methode nicht: Text", 438);
        let nv: VbaValue = v;
        if (d.type === "TextBox" || d.type === "ComboBox") nv = toStr(v);
        else if (d.type === "CheckBox" || d.type === "OptionButton" || d.type === "ToggleButton") nv = v === null ? null : toBool(v);
        else if (d.type === "SpinButton" || d.type === "ScrollBar") { const lo = Math.min(d.min ?? 0, d.max ?? 100), hi = Math.max(d.min ?? 0, d.max ?? 100); nv = Math.max(lo, Math.min(hi, Math.round(toNum(v)))); }
        else if (d.type === "MultiPage") { nv = Math.round(toNum(v)); if (nv < 0 || nv >= (d.pages?.length ?? 0)) throw new VbaError("Ungültiger Eigenschaftswert", 380); }
        else throw new VbaError("Objekt unterstützt diese Eigenschaft oder Methode nicht: Value", 438);
        if (d.value === nv) return events;
        d.value = nv;
        if (d.type === "ComboBox") d.listIndex = (d.list ?? []).findIndex((r) => r[0] === nv);
        this.changed("value");
        events.push("Change");
        if ((d.type === "CheckBox" || d.type === "ToggleButton") || (d.type === "OptionButton" && nv === true)) events.push("Click");
        if (d.type === "OptionButton" && nv === true) {
          // The other option buttons of the group (same container and GroupName) are switched off.
          for (const o of this.parent) if (o !== d && o.type === "OptionButton" && (o.groupName ?? "") === (d.groupName ?? "") && o.value === true) { o.value = false; this.form.changed(o, "value"); }
        }
        return events;
      }
      case "listindex": {
        if (!this.isList()) break;
        const i = Math.round(toNum(v));
        if (i < -1 || i >= (d.list?.length ?? 0)) throw new VbaError("Eigenschaft ListIndex konnte nicht festgelegt werden. Ungültiger Eigenschaftswert.", 380);
        if ((d.listIndex ?? -1) === i) return events;
        d.listIndex = i;
        if (d.type === "ComboBox") d.value = i >= 0 ? d.list![i][0] ?? "" : d.value;
        this.changed("listIndex");
        return ["Change", "Click"];
      }
      case "list": {
        if (!this.isList()) break;
        if (!args.length) {
          // List = array: rows from a 1-D or 2-D array.
          if (!isArr(v)) throw new VbaError("Typen unverträglich", 13);
          d.list = v.items.map((r) => (isArr(r) ? r.items.map((x) => toStr(x)) : [toStr(r)]));
          d.listIndex = -1; d.selected = [];
          this.changed("list");
          return events;
        }
        const r = toNum(args[0]), c = args[1] === undefined ? 0 : toNum(args[1]);
        const row = d.list?.[r];
        if (!row) throw new VbaError("Eigenschaft List konnte nicht festgelegt werden. Ungültiges Argument.", 381);
        row[c] = toStr(v);
        this.changed("list");
        return events;
      }
      case "selected": {
        if (!this.isList()) break;
        const i = toNum(args[0]);
        if (!d.multiSelect) return this.setQuiet("listindex", toBool(v) ? i : -1);
        (d.selected ??= [])[i] = toBool(v);
        this.changed("selected");
        return ["Change"];
      }
      case "font": throw new VbaError("Eigenschaft kann nicht gesetzt werden: Font", 438);
    }
    const pk = PLAIN_KEY.get(k);
    if (!pk) throw new VbaError(`Objekt unterstützt diese Eigenschaft oder Methode nicht: ${name}`, 438);
    const nv = BOOL_PROPS.has(k) ? toBool(v) : STR_PROPS.has(k) ? toStr(v) : toNum(v);
    if (k === "name") { (d as unknown as Record<string, VbaValue>)[pk] = nv; this.form.index(); }
    else (d as unknown as Record<string, VbaValue>)[pk] = nv;
    this.changed(pk);
    return events;
  }
  set(name: string, value: VbaValue, args: VbaValue[]): void { this.setQuiet(name, value, args); }

  /** Methods: AddItem, RemoveItem, Clear, SetFocus, Move, DropDown. */
  call(name: string, args: VbaValue[]): VbaValue {
    const k = low(name), d = this.d;
    switch (k) {
      case "additem": {
        if (!this.isList()) break;
        const list = (d.list ??= []);
        const at = args[1] === undefined ? list.length : toNum(args[1]);
        if (at < 0 || at > list.length) throw new VbaError("Ungültiges Argument", 5);
        list.splice(at, 0, [args[0] === undefined ? "" : toStr(args[0])]);
        if (d.multiSelect) (d.selected ??= []).splice(at, 0, false);
        if ((d.listIndex ?? -1) >= at) d.listIndex = (d.listIndex ?? -1) + 1;
        this.changed("list");
        return undefined;
      }
      case "removeitem": {
        if (!this.isList()) break;
        const i = toNum(args[0]);
        if (!d.list || i < 0 || i >= d.list.length) throw new VbaError("Ungültiges Argument", -2147024809);
        d.list.splice(i, 1);
        d.selected?.splice(i, 1);
        if ((d.listIndex ?? -1) === i) d.listIndex = -1; else if ((d.listIndex ?? -1) > i) d.listIndex = (d.listIndex ?? 0) - 1;
        this.changed("list");
        return undefined;
      }
      case "clear": if (this.isList()) { d.list = []; d.listIndex = -1; d.selected = []; if (d.type === "ComboBox") d.value = ""; this.changed("list"); return undefined; } break;
      case "setfocus": this.changed("focus"); return undefined;
      case "dropdown": this.changed("dropdown"); return undefined;
      case "move": {
        const [l, t, w, h] = args;
        if (l !== undefined) d.left = toNum(l);
        if (t !== undefined) d.top = toNum(t);
        if (w !== undefined) d.width = toNum(w);
        if (h !== undefined) d.height = toNum(h);
        this.changed("position");
        return undefined;
      }
      case "zorder": this.changed("zorder"); return undefined;
    }
    return NOT_A_METHOD;
  }
  argNames(name: string): string[] | undefined { return low(name) === "additem" ? ["pvargItem", "pvargIndex"] : low(name) === "move" ? ["Left", "Top", "Width", "Height"] : undefined; }

  private pageObject(i: number): ControlObj {
    const p = this.d.pages?.[i];
    if (!p) throw new VbaError("Ungültiges Argument", -2147024809);
    return this.form.control(p.name)!;
  }
  private pagesObject(): VbaObject {
    const d = this.d;
    const item = (k: VbaValue) => (typeof k === "number" ? this.pageObject(k) : this.pageObject((d.pages ?? []).findIndex((p) => low(p.name) === low(toStr(k)))));
    return {
      kind: "object",
      typeName: () => "Pages",
      items: () => (d.pages ?? []).map((p) => this.form.control(p.name)!),
      default: (a) => item(a[0]),
      get: (n, a) => { const k = low(n); if (k === "count") return d.pages?.length ?? 0; if (k === "item") return item(a[0]); throw new VbaError(`Objekt unterstützt diese Eigenschaft oder Methode nicht: ${n}`, 438); },
      call: (n, a) => {
        const k = low(n);
        if (k === "add") {
          const pages = (d.pages ??= []);
          let name = a[0] === undefined ? "" : toStr(a[0]);
          if (!name) { let i = pages.length + 1; while (this.form.control("Page" + i)) i++; name = "Page" + i; }
          const p = newControl("Page", name);
          if (a[1] !== undefined) p.caption = toStr(a[1]);
          pages.splice(a[2] === undefined ? pages.length : toNum(a[2]), 0, p);
          this.form.index();
          this.changed("pages");
          return this.form.control(name)!;
        }
        if (k === "remove") { const i = typeof a[0] === "number" ? a[0] : (d.pages ?? []).findIndex((p) => low(p.name) === low(toStr(a[0]))); d.pages?.splice(i, 1); this.form.index(); this.changed("pages"); return undefined; }
        return NOT_A_METHOD;
      },
    };
  }
}
