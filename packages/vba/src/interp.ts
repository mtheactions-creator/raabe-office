// Our VBA interpreter (D-079): the language as Office macros in school use it, run against an object model the
// app gives (RaabePresent's Application/ActivePresentation, RaabeCalc's Excel model). The syntax is read by
// parser.ts; this file runs it:
// - standard, class, document and form modules (VbaRuntime takes the code of standard modules, or VbaModule
//   records with their kind); objects of class modules with New, member variables, Property Get/Let/Set, Me,
//   Class_Initialize and Class_Terminate; Collection and an in-memory Scripting.Dictionary; TypeName, TypeOf … Is;
// - events: Public Event, RaiseEvent, WithEvents variables with their handlers (x_Ereignis), and procedures the app
//   fires by name (fire: CommandButton1_Click, UserForm_Initialize, Workbook_Open …) or for its own objects (raise);
// - errors as VBA handles them: On Error Resume Next / GoTo Marke / GoTo 0, Resume, Resume Next, the Err object;
// - a debugger: breakpoints, single steps (into, over, out), Stop, Debug.Print and Debug.Assert, the call stack
//   and local variables while paused (VbaHost.debugBreak);
// - files only through the app and only after the app asked the user (VbaHost.fileAccess, VbaHost.files, D-080):
//   Open … For Input/Output/Append, Print #, Write #, Input #, Line Input #, Close, FreeFile, EOF, LOF, Dir, Kill,
//   MkDir, RmDir, FileCopy, Name; GetSetting/SaveSetting in a store of the app (or in memory), never the registry;
// - UserForms (forms.ts): a form module's default instance, Show (modal or not), Hide, Load, Unload, Controls.
// A macro still cannot reach outside the app: Shell, CreateObject (except Scripting.Dictionary), GetObject,
// SendKeys, Environ and Declare stop with a runtime error (D-066).
//
// Class_Terminate (decided for D-100+): the runtime counts the references to an object held in variables, array
// elements, member variables, Collections and Dictionaries. When the count reaches zero, Terminate runs at the end
// of the statement that dropped the last reference (Set x = Nothing, a local variable at End Sub). An object that
// was handed to the app's object model is kept alive (the app may hold it), so its Terminate does not run.

import { BUILTIN_PARAMS, NOT_BUILTIN, pureBuiltin } from "./builtins";
import { ControlObj, FORM_FRAME, FormState, defaultFont, fontObject, type VbaForm, type VbaFormDesign } from "./forms";
import { VbaCollection, VbaDictionary, VbaErr, VbaRecord, type Holder } from "./objects";
import { parseExpression, parseModule, parseStatements, type Expr, type ParsedModule, type PrintItem, type Proc, type Ref, type Stmt, type VarDecl } from "./parser";
import type { VbaModule } from "./project";
import {
  CONSTANTS, NOT_A_METHOD, SCALAR_TYPES, VbaError, coerce, copyArray, decodeCp1252, defaultOf, encodeCp1252, flat, getIn, isArr, isObj, like, makeArray, setIn,
  toBool, toDate, toNum, toStr, type VbaArray, type VbaObject, type VbaValue,
} from "./values";

export { NOT_A_METHOD, VbaError, type VbaArray, type VbaObject, type VbaValue } from "./values";
export { parseModule } from "./parser";

// ----- what the app gives -------------------------------------------------------------------------------------

/** Why a macro asks to reach a file: the app asks the user before it allows it (D-080). */
export type VbaFileAccessKind = "open" | "save" | "append" | "export" | "delete" | "copy" | "rename" | "mkdir" | "rmdir" | "list";
/** The app's file access for macros (paths as the macro gives them; the app resolves them). */
export interface VbaFileSystem {
  read(path: string): Promise<Uint8Array>;
  write(path: string, data: Uint8Array): Promise<void>;
  /** The entries of a folder (Dir). */
  list?(folder: string): Promise<{ name: string; dir: boolean }[]>;
  /** A file's size and kind, or null when it does not exist (FileLen, FileDateTime, Dir of one file). */
  stat?(path: string): Promise<{ size: number; dir: boolean; modified?: Date } | null>;
  remove?(path: string): Promise<void>;
  rename?(from: string, to: string): Promise<void>;
  mkdir?(path: string): Promise<void>;
  rmdir?(path: string): Promise<void>;
}
/** Where GetSetting and SaveSetting keep their values (the app's own settings, never the Windows registry). */
export interface VbaSettingsStore {
  get(app: string, section: string, key: string): string | undefined;
  set(app: string, section: string, key: string, value: string): void;
  delete(app: string, section: string, key?: string): void;
  all?(app: string, section: string): [string, string][];
}
export type VbaStepMode = "continue" | "into" | "over" | "out" | "stop";
/** Where and why the run paused (breakpoint, single step, Stop, Unterbrechen, Debug.Assert False). */
export interface VbaBreak { reason: "breakpoint" | "step" | "stop" | "pause" | "assert"; module: string; proc: string; line: number; depth: number }
export interface VbaStackFrame { module: string; proc: string; line: number; depth: number }
export interface VbaVariable { name: string; value: VbaValue; type: string; members?: VbaVariable[] }

export interface VbaHost {
  /** Global names of the app (Application, ActivePresentation, ActiveSheet …); undefined when unknown. */
  global(name: string): VbaValue;
  msgBox(prompt: string, buttons: number, title: string): Promise<number>;
  inputBox(prompt: string, title: string, def: string): Promise<string>;
  /** The app's name in messages and as the default title of MsgBox and InputBox (default "RaabePresent"). */
  appName?: string;
  /** A line of Debug.Print (the Direktbereich). */
  debugPrint?(text: string): void;
  /** The run paused: the app shows where; the returned mode (or a later VbaRuntime.resume) continues it. Without
   *  this member breakpoints, Stop and single steps are ignored. */
  debugBreak?(info: VbaBreak): void | Promise<VbaStepMode | void>;
  /** Asks the user whether the macro may do this with the file (each time); false or no answer refuses it. */
  fileAccess?(kind: VbaFileAccessKind, path: string, target?: string): Promise<boolean>;
  files?: VbaFileSystem;
  settings?: VbaSettingsStore;
  /** Show a UserForm (the app draws form.design and keeps it current through form.onChange). The runtime does not
   *  wait for a returned promise: a modal Show returns when the form is hidden or unloaded. */
  showForm?(form: VbaForm, modal: boolean): void | Promise<void>;
  hideForm?(form: VbaForm): void;
  /** The object a document module stands for (Me in DieseArbeitsmappe, Tabelle1 …). */
  documentObject?(module: string): VbaObject | undefined;
}

// ----- the runtime's structures ---------------------------------------------------------------------------------

class Exit { constructor(readonly what: string) {} }
class EndSig {}
class Jump { constructor(readonly label: string, readonly handler: boolean) {} }
class ResumeSig { constructor(readonly mode: "retry" | "next" | "label", readonly label?: string) {} }
const isSignal = (e: unknown): boolean => e instanceof Exit || e instanceof EndSig || e instanceof Jump || e instanceof ResumeSig;
const MISSING = Symbol("missing");
type Maybe = VbaValue | typeof MISSING;
const low = (s: string) => s.toLowerCase();
const lit = (v: VbaValue): Expr => ({ e: "lit", v });

interface ModuleInfo {
  name: string;
  kind: "standard" | "class" | "document";
  p: ParsedModule;
  /** Module-level variables (standard and document modules). */
  scope: Scope;
  form?: VbaFormDesign;
  predeclared: boolean;
  defaultInstance: Instance | null;
  /** Public member variables (class modules). */
  publicVars: Set<string>;
}
interface Owner { module: ModuleInfo; inst: Instance | null }
class Scope {
  vars = new Map<string, VbaValue>();
  refs = new Map<string, Ref>();
  types = new Map<string, string>();
  names = new Map<string, string>();
  autoNew = new Map<string, string>();
  withEvents?: Set<string>;
  constructor(readonly parent: Scope | null, readonly owner: Owner | null) {}
}
interface Frame { module: ModuleInfo; proc: Proc; scope: Scope; me: Instance | null; line: number; depth: number; parent: Frame | null; last?: Stmt }
/** gen: End raised in one run (an event during a modal form) also ends the runs waiting below it. */
interface Ctx { withObj: VbaValue[]; resumeNext: boolean; handler: string | null; inHandler: boolean; frame: Frame; top: number; errTop: number; gen: number }
interface Sink { scope: Scope; key: string; owner: Owner }
interface OpenFile { path: string; mode: "input" | "output" | "append"; data: string; pos: number; out: string; dirty: boolean; col: number }

/** An object of a class or form module. */
class Instance implements VbaObject {
  readonly kind = "object";
  refs = 0;
  dead = false;
  form: FormState | null = null;
  formLoaded = false;
  readonly scope: Scope;
  constructor(readonly cls: ModuleInfo, global: Scope) { this.scope = new Scope(global, { module: cls, inst: this }); }
  typeName(): string { return this.cls.name; }
  /** The app reads a public member variable (procedures need the runtime: VbaRuntime.callMethod). */
  get(name: string): VbaValue {
    const k = low(name);
    if (this.cls.publicVars.has(k)) return this.scope.vars.get(k);
    throw new VbaError(`Objekt unterstützt diese Eigenschaft oder Methode nicht: ${name}`, 438);
  }
  set(name: string, value: VbaValue): void {
    const k = low(name);
    if (!this.cls.publicVars.has(k)) throw new VbaError(`Objekt unterstützt diese Eigenschaft oder Methode nicht: ${name}`, 438);
    this.scope.vars.set(k, coerce(this.scope.types.get(k), value));
  }
}

const DECL_PROC: Proc = { name: "(Deklarationen)", kind: "sub", params: [], body: [], isPrivate: true, isStatic: false, line: 0, labels: new Map(), attrs: {} };
const ROOT_PROC: Proc = { ...DECL_PROC, name: "(Aufruf)" };
const bpKey = (module: string, line: number) => `${low(module)}\u0000${line}`;

/** Our VBA interpreter. */
export class VbaRuntime {
  /** Debug.Print lines. */
  readonly printed: string[] = [];
  private modules: ModuleInfo[] = [];
  private byName = new Map<string, ModuleInfo>();
  private global = new Scope(null, null);
  private steps = 0;
  private initialized = false;
  private err: VbaErr;
  private pending = new Map<Instance, number>();
  private instances = 0;
  private curDepth = 0;
  private sinks = new Map<object, Sink[]>();
  private pinned = new Set<Instance>();
  private files = new Map<number, OpenFile>();
  private dirList: string[] | null = null;
  private memSettings = new Map<string, string>();
  private printBuf = "";
  private bps = new Set<string>();
  private stepMode: "none" | "into" | "over" | "out" = "none";
  private stepDepth = 0;
  private pauseRequested = false;
  private pausedCtx: Ctx | null = null;
  private pausedInfo: VbaBreak | null = null;
  private resumeFn: ((m: VbaStepMode) => void) | null = null;
  private debugSuppressed = 0;
  private statics = new Map<Proc, Map<object, Scope>>();
  private gen = 0;
  private elTypes = new WeakMap<VbaArray, string>();
  private formInst = new WeakMap<VbaForm, Instance>();
  private readonly holder: Holder = { retain: (v) => this.retain(v), release: (v) => this.release(v, this.curDepth) };
  private readonly appName: string;

  /** The code of standard modules (as before), or modules with their kind and name (class, document, form). */
  constructor(private host: VbaHost, sources: (string | VbaModule)[], private limit = 5_000_000) {
    this.appName = host.appName ?? "RaabePresent";
    this.err = new VbaErr();
    sources.forEach((src, i) => {
      const m = typeof src === "string" ? { name: "", kind: "standard" as const, code: src } : src;
      const p = parseModule(m.code);
      const attrName = p.attributes.get("name")?.replace(/^"|"$/g, "");
      const name = m.name || attrName || `Modul${i + 1}`;
      const predeclared = !!m.form || /^true$/i.test(p.attributes.get("predeclaredid") ?? "");
      const info: ModuleInfo = { name, kind: m.form ? "class" : m.kind, p, scope: null as unknown as Scope, form: m.form, predeclared, defaultInstance: null, publicVars: new Set() };
      info.scope = new Scope(this.global, { module: info, inst: null });
      for (const d of p.decls) if (d.s === "dim" && d.isPublic) for (const v of d.vars) info.publicVars.add(low(v.name));
      this.modules.push(info);
      if (!this.byName.has(low(name))) this.byName.set(low(name), info);
    });
  }

  // ----- the app's entry points ------------------------------------------------------------------------------

  /** The macros a user can run (public Subs without parameters), as PowerPoint's Makros dialog lists them; those of
   *  document modules with the module's name (Tabelle1.Makro). */
  macros(): string[] {
    const out: string[] = [];
    for (const m of this.modules) {
      if (m.kind === "class") continue;
      for (const p of m.p.procs.values()) if (p.kind === "sub" && !p.isPrivate && p.params.every((x) => x.optional)) out.push(m.kind === "document" ? `${m.name}.${p.name}` : p.name);
    }
    return out;
  }

  /** Runs a macro (Modul1.Makro, Makro or Datei!Modul1.Makro). */
  async run(name: string, args: VbaValue[] = []): Promise<VbaValue> {
    const found = this.findMacro(name);
    if (!found) throw new VbaError(`Das Makro "${name}" wurde nicht gefunden.`, 35);
    return this.entry(async (ctx) => {
      const me = found.mod.kind === "class" ? await this.defaultInstance(found.mod, ctx) : null;
      return this.invoke(found.proc, found.mod, me, args.map(lit), undefined, ctx.frame.scope, ctx);
    });
  }

  /** Fires an event procedure of a module or a loaded form (UserForm_Initialize, CommandButton1_Click,
   *  Workbook_Open …). Arguments are passed ByRef: the procedure's changes are written back into `args` (Cancel).
   *  Returns false (and does nothing) when the module has no such procedure. */
  async fire(target: string | VbaForm, proc: string, args: VbaValue[] = []): Promise<boolean> {
    let mod: ModuleInfo | undefined, inst: Instance | null = null;
    if (typeof target === "string") mod = this.byName.get(low(target));
    else { inst = this.formInst.get(target) ?? null; mod = inst?.cls; }
    const p = mod?.p.procs.get(low(proc));
    if (!mod || !p) return false;
    await this.entry(async (ctx) => {
      if (!inst && mod!.kind === "class") inst = await this.defaultInstance(mod!, ctx);
      await this.invoke(p, mod!, inst, this.slots(args), undefined, ctx.frame.scope, ctx);
    });
    return true;
  }

  /** An event of one of the app's objects (Application, Presentation …): runs the handlers of every WithEvents
   *  variable that holds the object. Returns how many handlers ran. */
  async raise(source: VbaObject, event: string, args: VbaValue[] = []): Promise<number> {
    let n = 0;
    await this.entry(async (ctx) => { n = await this.raiseOn(source, event, this.slots(args), undefined, ctx.frame.scope, ctx); });
    return n;
  }

  /** The user closed a form with its X: UserForm_QueryClose (it may cancel), then Unload. */
  async closeForm(form: VbaForm, closeMode = 0): Promise<boolean> {
    const inst = this.formInst.get(form);
    if (!inst) return true;
    let closed = true;
    await this.entry(async (ctx) => { closed = await this.unload(inst, ctx, closeMode); });
    return closed;
  }

  /** The user changed a control (typed, clicked, chose an entry): the value is set and the control's events run
   *  (Change, Click) as Office runs them. */
  async setFormValue(form: VbaForm, control: string, prop: string, value: VbaValue, args: VbaValue[] = []): Promise<void> {
    const inst = this.formInst.get(form);
    const c = (form as FormState).control?.(control);
    if (!inst || !c) throw new VbaError(`Steuerelement nicht gefunden: ${control}`, 438);
    const events = c.setQuiet(prop, value, args);
    if (!events.length) return;
    await this.entry(async (ctx) => { for (const ev of events) await this.callEvent(inst, `${c.d.name}_${ev}`, [], ctx); });
  }

  /** Calls a method or property of a class instance or form from the app (outside a running macro). */
  async callMethod(obj: VbaObject, name: string, args: VbaValue[] = []): Promise<VbaValue> {
    return this.entry(async (ctx) => this.memberCall(obj, name, args.map(lit), undefined, ctx.frame.scope, ctx));
  }

  /** Ends everything as VBA's End and the editor's Zurücksetzen do: module variables are reset, forms closed, files
   *  closed (their output written). Class_Terminate does not run, as in VBA. */
  async reset(): Promise<void> {
    await this.closeAllFiles();
    for (const m of this.modules) {
      const inst = m.defaultInstance;
      if (inst?.form?.visible) { inst.form.hidden(); this.host.hideForm?.(inst.form); }
      m.defaultInstance = null;
      m.scope = new Scope(this.global, { module: m, inst: null });
    }
    this.global = new Scope(null, null);
    for (const m of this.modules) m.scope = new Scope(this.global, { module: m, inst: null });
    this.pending.clear();
    this.pinned.clear();
    this.sinks.clear();
    this.statics.clear();
    this.initialized = false;
    this.err.clear();
  }

  // ----- debugger ----------------------------------------------------------------------------------------------

  /** The lines of a module that hold a statement (where a breakpoint can be set). */
  statementLines(module: string): number[] {
    const m = this.byName.get(low(module));
    if (!m) return [];
    const out = new Set<number>();
    const walk = (l: Stmt[]) => {
      for (const s of l) {
        if (s.s !== "label" && s.s !== "dim" && s.s !== "const") out.add(s.line);
        if (s.s === "if") { s.branches.forEach((b) => walk(b.body)); walk(s.other); }
        else if (s.s === "for" || s.s === "foreach" || s.s === "do" || s.s === "with") walk(s.body);
        else if (s.s === "select") s.cases.forEach((c) => walk(c.body));
      }
    };
    for (const p of [...m.p.procs.values(), ...[...m.p.props.values()].flatMap((x) => [x.get, x.let, x.set].filter((y): y is Proc => !!y))]) walk(p.body);
    return [...out].sort((a, b) => a - b);
  }
  setBreakpoint(module: string, line: number, on = true): void { if (on) this.bps.add(bpKey(module, line)); else this.bps.delete(bpKey(module, line)); }
  toggleBreakpoint(module: string, line: number): boolean { const k = bpKey(module, line); if (this.bps.has(k)) { this.bps.delete(k); return false; } this.bps.add(k); return true; }
  clearBreakpoints(module?: string): void { if (!module) this.bps.clear(); else for (const k of [...this.bps]) if (k.startsWith(low(module) + "\u0000")) this.bps.delete(k); }
  breakpoints(): { module: string; line: number }[] {
    return [...this.bps].map((k) => { const [m, l] = k.split("\u0000"); return { module: this.byName.get(m)?.name ?? m, line: Number(l) }; });
  }
  /** Unterbrechen: pause at the next statement. */
  pause(): void { this.pauseRequested = true; }
  /** Continues a paused run: Fortsetzen (continue), Einzelschritt (into), Prozedurschritt (over), Prozedur
   *  abschließen (out), Zurücksetzen (stop: the run ends as with End). */
  resume(mode: VbaStepMode = "continue"): void { this.resumeFn?.(mode); }
  /** Where the run is paused, or null. */
  pausedAt(): VbaBreak | null { return this.pausedInfo; }
  /** Aufrufliste: the procedures of the paused run, innermost first. */
  callStack(): VbaStackFrame[] {
    const out: VbaStackFrame[] = [];
    for (let f: Frame | null = this.pausedCtx?.frame ?? null; f && f.proc !== ROOT_PROC; f = f.parent) out.push({ module: f.module.name, proc: f.proc.name, line: f.line, depth: f.depth });
    return out;
  }
  /** Lokal-Fenster: the variables of a paused procedure (0 = innermost), Me first. */
  locals(frame = 0): VbaVariable[] {
    const f = this.frameAt(frame);
    if (!f) return [];
    const out: VbaVariable[] = [];
    if (f.me) out.push({ name: "Me", value: f.me, type: f.me.cls.name, members: this.inspect(f.me) });
    const s = f.scope;
    const keys = new Set([...f.proc.params.map((p) => low(p.name)).filter((k) => s.vars.has(k) || s.refs.has(k)), ...s.vars.keys(), ...s.refs.keys()]);
    for (const k of keys) {
      const v = s.refs.get(k)?.get() ?? s.vars.get(k);
      out.push({ name: s.names.get(k) ?? k, value: v, type: this.typeName(v, s.types.get(k)), members: isObj(v) || isArr(v) ? this.inspect(v) : undefined });
    }
    return out;
  }
  /** The members of an object or array, one level (for the Lokal-Fenster's tree). */
  inspect(v: VbaValue): VbaVariable[] {
    if (v instanceof Instance) return [...v.scope.vars.entries()].map(([k, x]) => ({ name: v.scope.names.get(k) ?? k, value: x, type: this.typeName(x, v.scope.types.get(k)) }));
    if (v instanceof VbaRecord) return v.fieldNames().map((n) => { const x = v.get(n, []); return { name: n, value: x, type: this.typeName(x) }; });
    if (isArr(v)) return v.items.slice(0, 1000).map((x, i) => ({ name: `(${v.lb + i})`, value: x, type: this.typeName(x) }));
    if (v instanceof VbaCollection) return v.items().slice(0, 1000).map((x, i) => ({ name: `Item ${i + 1}`, value: x, type: this.typeName(x) }));
    return [];
  }
  /** Evaluates an expression in a paused procedure (or at module level): the Direktbereich's "? x". */
  async evaluate(expression: string, frame = 0): Promise<VbaValue> {
    const x = parseExpression(expression);
    const ctx = this.debugCtx(frame);
    this.debugSuppressed++;
    try { return await this.eval(x, ctx.frame.scope, ctx); } finally { this.debugSuppressed--; }
  }
  /** Direktbereich: "? Ausdruck" (or Print) prints the value through debugPrint, anything else runs as statements. */
  async immediate(text: string, frame = 0): Promise<void> {
    const t = text.trim();
    if (!t) return;
    const m = /^(\?|print\s)(.*)$/i.exec(t);
    if (m) { const v = await this.evaluate(m[2], frame); this.printLine(v === null ? "Null" : this.printValue(v)); return; }
    const stmts = parseStatements(t);
    const ctx = this.debugCtx(frame);
    this.debugSuppressed++;
    try { await this.exec(stmts, ctx.frame.scope, ctx); } finally { this.debugSuppressed--; this.flushPrint(); }
  }
  private frameAt(i: number): Frame | null {
    let f: Frame | null = this.pausedCtx?.frame ?? null;
    for (let k = 0; f && k < i; k++) f = f.parent;
    return f && f.proc !== ROOT_PROC ? f : null;
  }
  private debugCtx(frame: number): Ctx {
    const f = this.frameAt(frame);
    if (f) return { withObj: [...(this.pausedCtx && frame === 0 ? this.pausedCtx.withObj : [])], resumeNext: false, handler: null, inHandler: false, frame: f, top: 0, errTop: 0, gen: this.gen };
    return this.rootCtx();
  }
  private get debugOn(): boolean { return !!this.host.debugBreak && this.debugSuppressed === 0; }
  private async debugCheck(st: Stmt, ctx: Ctx, sameLine: boolean): Promise<void> {
    const f = ctx.frame;
    let reason: VbaBreak["reason"] | null = null;
    if (this.pauseRequested) reason = "pause";
    else if (this.stepMode === "into" || (this.stepMode === "over" && f.depth <= this.stepDepth) || (this.stepMode === "out" && f.depth < this.stepDepth)) reason = "step";
    else if (this.bps.size && this.bps.has(bpKey(f.module.name, st.line))) reason = "breakpoint";
    // Several statements on one line (If … Then …, a: b) pause once.
    if (!reason || (sameLine && reason !== "pause")) return;
    await this.pauseAt(reason, ctx);
  }
  private async pauseAt(reason: VbaBreak["reason"], ctx: Ctx): Promise<void> {
    if (!this.host.debugBreak) return;
    this.pauseRequested = false;
    this.stepMode = "none";
    const f = ctx.frame;
    const info: VbaBreak = { reason, module: f.module.name, proc: f.proc.name, line: f.line, depth: f.depth };
    this.pausedCtx = ctx;
    this.pausedInfo = info;
    this.flushPrint();
    let mode: VbaStepMode;
    try {
      mode = await new Promise<VbaStepMode>((res) => {
        this.resumeFn = res;
        try {
          const r = this.host.debugBreak!(info);
          if (r && typeof (r as Promise<unknown>).then === "function") (r as Promise<VbaStepMode | void>).then((m) => { if (m) res(m); }, () => res("stop"));
        } catch { res("stop"); }
      });
    } finally { this.pausedCtx = null; this.pausedInfo = null; this.resumeFn = null; }
    if (mode === "stop") throw new EndSig();
    if (mode === "into") this.stepMode = "into";
    else if (mode === "over" || mode === "out") { this.stepMode = mode; this.stepDepth = f.depth; }
  }

  // ----- running -----------------------------------------------------------------------------------------------

  private rootCtx(): Ctx {
    const mod = this.modules.find((m) => m.kind === "standard") ?? this.modules[0] ?? this.dummyModule();
    const frame: Frame = { module: mod, proc: ROOT_PROC, scope: mod.scope, me: null, line: 0, depth: 0, parent: null };
    return { withObj: [], resumeNext: false, handler: null, inHandler: false, frame, top: 0, errTop: 0, gen: this.gen };
  }
  private dummyModule(): ModuleInfo {
    const p = parseModule("");
    const m: ModuleInfo = { name: "Modul1", kind: "standard", p, scope: null as unknown as Scope, predeclared: false, defaultInstance: null, publicVars: new Set() };
    m.scope = new Scope(this.global, { module: m, inst: null });
    this.modules.push(m);
    this.byName.set("modul1", m);
    return m;
  }

  /** One call from the app: module initialisation, the step limit, End, and what remains after the run. */
  private async entry<T>(fn: (ctx: Ctx) => Promise<T>): Promise<T | undefined> {
    const outer = this.steps;
    this.steps = 0;
    const ctx = this.rootCtx();
    try {
      await this.ensureInit(ctx);
      return await fn(ctx);
    } catch (e) {
      if (e instanceof EndSig) { if (ctx.gen === this.gen) { this.gen++; await this.reset(); } return undefined; }
      if (e instanceof Exit) return undefined;
      if (e instanceof Jump || e instanceof ResumeSig) throw new VbaError(e instanceof Jump ? `Sprungmarke nicht definiert: ${e.label}` : "Resume ohne Fehler", e instanceof Jump ? 35 : 20);
      throw e instanceof VbaError ? e : new VbaError(String((e as Error)?.message ?? e), 5);
    } finally {
      if (this.pending.size) await this.settle(0, ctx).catch(() => undefined);
      this.flushPrint();
      await this.flushFiles().catch(() => undefined);
      this.steps = Math.max(this.steps, outer);
    }
  }

  private async ensureInit(ctx: Ctx): Promise<void> {
    if (this.initialized) return;
    this.initialized = true;
    for (const m of this.modules) {
      if (m.kind === "class") continue;
      const frame: Frame = { module: m, proc: DECL_PROC, scope: m.scope, me: null, line: 0, depth: 1, parent: ctx.frame };
      await this.runDecls(m, m.scope, { ...ctx, frame });
    }
  }

  /** Module-level declarations into a module scope (Public ones of standard modules into the global scope). */
  private async runDecls(m: ModuleInfo, scope: Scope, ctx: Ctx): Promise<void> {
    for (const d of m.p.decls) {
      ctx.frame.line = d.line;
      const target = (d.s === "dim" || d.s === "const") && d.isPublic && m.kind === "standard" ? this.global : scope;
      try {
        if (d.s === "dim") for (const v of d.vars) await this.declare(target, v, ctx);
        else if (d.s === "const") { const k = low(d.name); target.names.set(k, d.name); target.vars.set(k, await this.eval(d.value, scope, ctx)); }
      } catch (e) { throw this.located(e, d.line, ctx); }
    }
  }

  private findMacro(name: string): { proc: Proc; mod: ModuleInfo } | null {
    const bare = name.replace(/^.*!/, "");
    const dot = bare.lastIndexOf(".");
    if (dot > 0) {
      const mod = this.byName.get(low(bare.slice(0, dot)));
      const p = mod?.p.procs.get(low(bare.slice(dot + 1)));
      if (mod && p) return { proc: p, mod };
    }
    const k = low(bare.slice(dot + 1));
    for (const kind of ["standard", "document", "class"] as const) for (const m of this.modules) if (m.kind === kind) { const p = m.p.procs.get(k); if (p) return { proc: p, mod: m }; }
    return null;
  }

  /** A procedure by name as code in `from` sees it: its own module first, then public procedures of standard modules. */
  private findProc(k: string, from: ModuleInfo): { proc: Proc; mod: ModuleInfo } | null {
    const own = from.p.procs.get(k);
    if (own) return { proc: own, mod: from };
    for (const m of this.modules) if (m.kind === "standard" && m !== from) { const p = m.p.procs.get(k); if (p && !p.isPrivate) return { proc: p, mod: m }; }
    return null;
  }

  private tick(line: number): void {
    if (++this.steps > this.limit) throw new VbaError("Das Makro wurde angehalten: es lief zu lange (Endlosschleife?).", 6, line);
  }

  private located(e: unknown, line: number, ctx: Ctx): unknown {
    if (isSignal(e)) return e;
    let err = e instanceof VbaError ? e : new VbaError(String((e as Error)?.message ?? e), 5, line);
    if (!err.line && line) { const n = new VbaError(err.message, err.number, line); n.module = err.module; n.proc = err.proc; err = n; }
    if (!err.module) { err.module = ctx.frame.module.name; err.proc = ctx.frame.proc.name; }
    return err;
  }

  private async exec(stmts: Stmt[], scope: Scope, ctx: Ctx): Promise<void> {
    for (const st of stmts) {
      this.tick(st.line);
      const f = ctx.frame;
      const prev = f.line;
      f.line = st.line;
      this.curDepth = f.depth;
      if (ctx.gen !== this.gen) throw new EndSig();
      const prevStmt = f.last;
      f.last = st;
      if (st.s !== "label" && st.s !== "dim" && st.s !== "const" && (this.pauseRequested || this.debugOn)) await this.debugCheck(st, ctx, prevStmt !== st && prev === st.line);
      try {
        await this.stmt(st, scope, ctx);
      } catch (e) {
        if (isSignal(e)) throw e;
        const err = this.located(e, st.line, ctx) as VbaError;
        if (ctx.resumeNext) { this.err.take(err); continue; }
        if (ctx.handler && !ctx.inHandler) { this.err.take(err); throw new Jump(ctx.handler, true); }
        throw err;
      }
      if (this.pending.size) await this.settle(f.depth, ctx);
    }
  }

  /** A procedure's body: its labels, GoTo, the error handler and Resume. */
  private async execBody(p: Proc, scope: Scope, ctx: Ctx): Promise<void> {
    const body = p.body;
    let i = 0;
    while (i < body.length) {
      try {
        for (; i < body.length; i++) { ctx.top = i; await this.exec([body[i]], scope, ctx); }
      } catch (e) {
        if (e instanceof Jump) {
          const to = p.labels.get(low(e.label));
          if (to === undefined) throw this.located(new VbaError(`Sprungmarke nicht definiert: ${e.label}`, 35), ctx.frame.line, ctx);
          if (e.handler) { ctx.errTop = ctx.top; ctx.inHandler = true; }
          i = to;
          continue;
        }
        if (e instanceof ResumeSig) {
          if (!ctx.inHandler) throw this.located(new VbaError("Resume ohne Fehler", 20), ctx.frame.line, ctx);
          ctx.inHandler = false;
          this.err.clear();
          if (e.mode === "next") i = ctx.errTop + 1;
          else if (e.mode === "retry") i = ctx.errTop;
          else { const to = p.labels.get(low(e.label ?? "")); if (to === undefined) throw this.located(new VbaError(`Sprungmarke nicht definiert: ${e.label}`, 35), ctx.frame.line, ctx); i = to; }
          continue;
        }
        throw e;
      }
    }
  }

  private orderArgs(p: Proc, argExprs: (Expr | null)[], names: (string | null)[] | undefined): (Expr | null | undefined)[] {
    const out: (Expr | null | undefined)[] = [];
    const last = p.params[p.params.length - 1];
    argExprs.forEach((a, i) => {
      const nm = names?.[i];
      if (nm) {
        const j = p.params.findIndex((q) => low(q.name) === low(nm));
        if (j < 0) throw new VbaError(`Benanntes Argument nicht gefunden: ${nm}`, 448);
        out[j] = a;
      } else out[i] = a;
    });
    if (out.length > p.params.length && !last?.paramArray) throw new VbaError("Falsche Anzahl an Argumenten oder ungültige Eigenschaftszuweisung", 450);
    return out;
  }

  /** Calls a procedure: binds its parameters (ByRef to the caller's variables), runs its body, returns its value. */
  private async invoke(p: Proc, mod: ModuleInfo, me: Instance | null, argExprs: (Expr | null)[], names: (string | null)[] | undefined, caller: Scope, cctx: Ctx, propValue: Maybe = MISSING): Promise<VbaValue> {
    const scope = new Scope(me ? me.scope : mod.scope, { module: mod, inst: me });
    const frame: Frame = { module: mod, proc: p, scope, me, line: p.line, depth: cctx.frame.depth + 1, parent: cctx.frame };
    if (frame.depth > 3000) throw new VbaError("Nicht genügend Stapelspeicher", 28);
    const ctx: Ctx = { withObj: [], resumeNext: false, handler: null, inHandler: false, frame, top: 0, errTop: 0, gen: cctx.gen };
    const params = propValue === MISSING ? p.params : p.params.slice(0, -1);
    const ordered = this.orderArgs({ ...p, params }, argExprs, names);
    for (let i = 0; i < params.length; i++) {
      const prm = params[i];
      const k = low(prm.name);
      scope.names.set(k, prm.name);
      if (prm.paramArray) {
        const rest: VbaValue[] = [];
        for (let j = i; j < ordered.length; j++) { const a = ordered[j]; rest.push(a ? await this.eval(a, caller, cctx) : undefined); }
        scope.vars.set(k, { kind: "array", lb: 0, items: rest });
        break;
      }
      const a = ordered[i];
      if (prm.type) scope.types.set(k, prm.type);
      if (!a) {
        if (!prm.optional) throw new VbaError(`Argument ist nicht optional: ${prm.name}`, 449);
        this.store(scope, k, prm.def ? coerce(prm.type, await this.eval(prm.def, scope, ctx)) : prm.type && SCALAR_TYPES.has(prm.type) ? defaultOf(prm.type) : undefined);
        continue;
      }
      if (prm.byRef) { const ref = await this.refOf(a, caller, cctx); if (ref) { scope.refs.set(k, ref); continue; } }
      let v = await this.eval(a, caller, cctx);
      if (isArr(v)) v = copyArray(v);
      else if (v instanceof VbaRecord) v = v.copy();
      this.store(scope, k, coerce(prm.type, v));
    }
    if (propValue !== MISSING) {
      const last = p.params[p.params.length - 1];
      if (last) { const k = low(last.name); scope.names.set(k, last.name); if (last.type) scope.types.set(k, last.type); this.store(scope, k, coerce(last.type, propValue)); }
    }
    const resultKey = low(p.name);
    if (p.kind === "function" || p.kind === "get") { scope.names.set(resultKey, p.name); if (p.type) scope.types.set(resultKey, p.type); scope.vars.set(resultKey, this.initial(p.type)); }
    if (me) this.retain(me);
    let result: VbaValue;
    try {
      try {
        await this.execBody(p, scope, ctx);
      } catch (e) {
        if (!(e instanceof Exit && (e.what === "sub" || e.what === "function" || e.what === "property"))) throw e;
      }
      if (ctx.inHandler) this.err.clear();
      result = p.kind === "function" || p.kind === "get" ? scope.vars.get(resultKey) : undefined;
    } finally {
      this.releaseScope(scope, cctx.frame.depth);
      if (me) this.release(me, cctx.frame.depth);
    }
    return result;
  }

  /** A ByRef argument: the caller's variable or array element; null when the argument is a value. */
  private async refOf(a: Expr, scope: Scope, ctx: Ctx): Promise<Ref | null> {
    if (a.e === "slot") return a.ref;
    if (a.e === "id" && !a.br) {
      const k = low(a.name);
      if (k === "me") return null;
      const owner = this.lookupScope(k, scope);
      if (owner) {
        const r = owner.refs.get(k);
        if (r) return r;
        return { get: () => owner.vars.get(k), set: (v) => this.store(owner, k, coerce(owner.types.get(k), v)) };
      }
      // Not a variable: a procedure, module, object name or constant goes by value; anything else is a new variable.
      if (ctx.frame.me || this.findProc(k, ctx.frame.module) || this.byName.has(k) || CONSTANTS[k] !== undefined || this.host.global(a.name) !== undefined) return null;
      const target = ctx.frame.scope;
      target.vars.set(k, undefined);
      target.names.set(k, a.name);
      return { get: () => target.vars.get(k), set: (v) => this.store(target, k, v) };
    }
    if (a.e === "call" && a.fn.e === "id" && a.args.length) {
      const owner = this.lookupScope(low(a.fn.name), scope);
      const arr = owner ? owner.refs.get(low(a.fn.name))?.get() ?? owner.vars.get(low(a.fn.name)) : undefined;
      if (isArr(arr)) {
        const idx = (await this.args(a.args, scope, ctx)).map(toNum);
        return { get: () => getIn(arr, idx), set: (v) => this.setElement(arr, idx, v) };
      }
    }
    return null;
  }

  private lookupScope(k: string, scope: Scope | null): Scope | null {
    for (let s = scope; s; s = s.parent) if (s.vars.has(k) || s.refs.has(k)) return s;
    return null;
  }

  // ----- references and Terminate ------------------------------------------------------------------------------

  private retain(v: VbaValue): void {
    if (!this.instances) return;
    if (v instanceof Instance) v.refs++;
    else if (isArr(v)) for (const x of v.items) this.retain(x);
    else if (v instanceof VbaRecord) for (const x of v.fields.values()) this.retain(x);
  }
  private release(v: VbaValue, depth: number): void {
    if (!this.instances) return;
    if (v instanceof Instance) {
      v.refs--;
      if (v.refs <= 0 && !v.dead) { const d = this.pending.get(v); this.pending.set(v, d === undefined ? depth : Math.min(d, depth)); }
    } else if (isArr(v)) for (const x of v.items) this.release(x, depth);
    else if (v instanceof VbaRecord) for (const x of v.fields.values()) this.release(x, depth);
  }
  /** An object handed to the app's model stays alive (the app may keep it). */
  private pin(v: VbaValue): void { if (v instanceof Instance && !this.pinned.has(v)) { this.pinned.add(v); v.refs++; } }
  private releaseScope(scope: Scope, depth: number): void {
    if (scope.withEvents) for (const k of scope.withEvents) this.bindEvents(scope, k, scope.vars.get(k), null);
    if (!this.instances) return;
    for (const v of scope.vars.values()) this.release(v, depth);
  }
  /** Objects whose last reference went away in a statement at this depth (or deeper) are terminated. */
  private async settle(depth: number, ctx: Ctx): Promise<void> {
    for (let guard = 0; this.pending.size && guard < 100000; guard++) {
      let next: Instance | null = null;
      for (const [inst, d] of this.pending) {
        if (inst.refs > 0 || inst.dead) { this.pending.delete(inst); continue; }
        if (d >= depth) { next = inst; break; }
      }
      if (!next) return;
      this.pending.delete(next);
      await this.terminate(next, ctx);
    }
  }
  private async terminate(inst: Instance, ctx: Ctx): Promise<void> {
    inst.dead = true;
    if (inst.form) { if (inst.formLoaded) { this.hideForm(inst); inst.formLoaded = false; await this.callEvent(inst, "UserForm_Terminate", [], ctx); } }
    else await this.callEvent(inst, "Class_Terminate", [], ctx);
    this.releaseScope(inst.scope, ctx.frame.depth);
  }

  /** Writes a variable (reference counts and WithEvents bindings kept). */
  private store(scope: Scope, k: string, v: VbaValue): void {
    const old = scope.vars.get(k);
    if (old === v && scope.vars.has(k)) return;
    this.retain(v);
    scope.vars.set(k, v);
    if (scope.withEvents?.has(k)) this.bindEvents(scope, k, old, v);
    this.release(old, this.curDepth);
  }
  private setElement(arr: VbaArray, idx: number[], v: VbaValue): void {
    const t = this.elTypes.get(arr);
    this.retain(v);
    const old = setIn(arr, idx, t ? coerce(t, v) : v);
    this.release(old, this.curDepth);
  }

  // ----- events ------------------------------------------------------------------------------------------------

  private bindEvents(scope: Scope, k: string, old: VbaValue, v: VbaValue): void {
    if (!scope.owner) return;
    if (isObj(old)) { const l = this.sinks.get(old); if (l) { const i = l.findIndex((s) => s.scope === scope && s.key === k); if (i >= 0) l.splice(i, 1); } }
    if (isObj(v)) { const l = this.sinks.get(v) ?? []; l.push({ scope, key: k, owner: scope.owner }); this.sinks.set(v, l); }
  }
  private async raiseOn(source: object, event: string, args: (Expr | null)[], names: (string | null)[] | undefined, scope: Scope, ctx: Ctx): Promise<number> {
    const list = this.sinks.get(source)?.slice() ?? [];
    let n = 0;
    for (const s of list) {
      if (s.scope.vars.get(s.key) !== source) continue;
      const p = s.owner.module.p.procs.get(`${s.key}_${low(event)}`);
      if (!p) continue;
      await this.invoke(p, s.owner.module, s.owner.inst, args, names, scope, ctx);
      n++;
    }
    return n;
  }
  /** Calls an event procedure of an instance if its module has it (Class_Initialize, UserForm_Activate …). */
  private async callEvent(inst: Instance, proc: string, args: VbaValue[], ctx: Ctx): Promise<void> {
    const p = inst.cls.p.procs.get(low(proc));
    if (p) await this.invoke(p, inst.cls, inst, this.slots(args), undefined, ctx.frame.scope, ctx);
  }
  /** Values as ByRef slots: a procedure's changes are written back into the array. */
  private slots(args: VbaValue[]): Expr[] { return args.map((_, i) => ({ e: "slot", ref: { get: () => args[i], set: (v) => { args[i] = v; } } }) as Expr); }

  // ----- objects -----------------------------------------------------------------------------------------------

  private async create(cls: string, ctx: Ctx): Promise<VbaValue> {
    const k = low(cls).replace(/^(vba|scripting|msforms)\./, "");
    if (k === "collection") return new VbaCollection(this.holder);
    if (k === "dictionary") return new VbaDictionary(this.holder);
    const m = this.byName.get(k);
    if (m && m.kind === "class") return this.newInstance(m, ctx);
    throw new VbaError(`Objekterstellung durch ActiveX-Komponente nicht möglich: ${cls}`, 429);
  }
  private async newInstance(m: ModuleInfo, ctx: Ctx): Promise<Instance> {
    const inst = new Instance(m, this.global);
    this.instances++;
    this.pending.set(inst, ctx.frame.depth);
    const frame: Frame = { module: m, proc: DECL_PROC, scope: inst.scope, me: inst, line: 0, depth: ctx.frame.depth + 1, parent: ctx.frame };
    await this.runDecls(m, inst.scope, { ...ctx, withObj: [], frame });
    if (m.form) this.loadForm(inst);
    await this.callEvent(inst, m.form ? "UserForm_Initialize" : "Class_Initialize", [], ctx);
    return inst;
  }
  private async defaultInstance(m: ModuleInfo, ctx: Ctx): Promise<Instance> {
    if (!m.predeclared) throw new VbaError(`Ungültige Verwendung des Objekts: ${m.name}`, 91);
    if (!m.defaultInstance || m.defaultInstance.dead) {
      const inst = await this.newInstance(m, ctx);
      m.defaultInstance = inst;
      this.retain(inst);
    } else if (m.form && !m.defaultInstance.formLoaded) await this.reload(m.defaultInstance, ctx);
    return m.defaultInstance;
  }
  private loadForm(inst: Instance): void {
    const st = new FormState(inst.cls.form!, { event: async (proc, args) => { await this.entry(async (ctx) => this.callEvent(inst, proc, args, ctx)); } });
    inst.form = st;
    inst.formLoaded = true;
    this.formInst.set(st, inst);
  }
  /** A form that was unloaded and is used again is loaded again (new controls, UserForm_Initialize). */
  private async reload(inst: Instance, ctx: Ctx): Promise<void> {
    if (!inst.cls.form || inst.formLoaded) return;
    this.loadForm(inst);
    await this.callEvent(inst, "UserForm_Initialize", [], ctx);
  }

  private typeOf(v: VbaValue, cls: string): boolean {
    const k = low(cls);
    if (!isObj(v)) return false;
    if (k === "object") return true;
    if (v instanceof Instance) return low(v.cls.name) === k || v.cls.p.implements.some((i) => low(i) === k) || (!!v.cls.form && k === "userform");
    if (v instanceof ControlObj) return low(v.d.type) === k || k === "control";
    return low(v.typeName?.() ?? "") === k;
  }
  private typeName(v: VbaValue, declared?: string): string {
    if (declared && SCALAR_TYPES.has(declared) && !isArr(v)) return declared === "variant" ? `Variant/${this.typeName(v)}` : declared[0].toUpperCase() + declared.slice(1).replace("ptr", "Ptr");
    if (v === undefined) return "Empty";
    if (v === null) return declared && !SCALAR_TYPES.has(declared) ? "Nothing" : "Null";
    if (typeof v === "string") return "String";
    if (typeof v === "boolean") return "Boolean";
    if (typeof v === "number") return Number.isInteger(v) ? (Math.abs(v) <= 32767 ? "Integer" : Math.abs(v) <= 2147483647 ? "Long" : "Double") : "Double";
    if (v instanceof Date) return "Date";
    if (isArr(v)) { const t = this.elTypes.get(v); return t ? t[0].toUpperCase() + t.slice(1) + "()" : "Variant()"; }
    if (isObj(v)) return v.typeName?.() ?? "Object";
    return "Variant";
  }
  /** The value a control stands for in expressions (TextBox1 = "x" compares its Value). */
  private prim(v: VbaValue): VbaValue { return v instanceof ControlObj ? v.value() : v; }
  private initial(type: string | undefined): VbaValue {
    if (type) for (const m of this.modules) { const t = m.p.types.get(type); if (t) return this.record(t); }
    return defaultOf(type);
  }
  private record(t: { name: string; fields: VarDecl[] }): VbaRecord {
    const fields = new Map<string, VbaValue>(), names = new Map<string, string>();
    for (const f of t.fields) {
      names.set(low(f.name), f.name);
      fields.set(low(f.name), f.dims && f.dims.length ? makeArray(f.dims.map((d) => (d.e === "bin" && d.op === "range" ? [toNum((d.a as { v: VbaValue }).v), toNum((d.b as { v: VbaValue }).v)] : [0, toNum((d as { v: VbaValue }).v)]) as [number, number]), () => this.initial(f.type)) : this.initial(f.type));
    }
    return new VbaRecord(t.name, fields, names);
  }

  private async declare(scope: Scope, v: VarDecl, ctx: Ctx): Promise<void> {
    const k = low(v.name);
    // Dim is not an executable reset: a Dim reached again (in a loop) keeps the variable.
    if (scope.vars.has(k) || scope.refs.has(k)) return;
    scope.names.set(k, v.name);
    if (v.type) scope.types.set(k, v.type);
    if (v.withEvents) (scope.withEvents ??= new Set()).add(k);
    let val: VbaValue;
    if (v.dims) {
      if (v.dims.length) { const arr = makeArray(await this.bounds(v.dims, scope, ctx), () => this.initial(v.type)); if (v.type) this.elTypes.set(arr, v.type); val = arr; }
      else val = { kind: "array", lb: 0, items: [] };
    } else if (v.isNew) { val = null; scope.autoNew.set(k, v.type ?? ""); }
    else val = this.initial(v.type);
    this.store(scope, k, val);
  }

  private staticScope(p: Proc, me: Instance | null): Scope {
    let m = this.statics.get(p);
    if (!m) { m = new Map(); this.statics.set(p, m); }
    const key = me ?? p;
    let s = m.get(key);
    if (!s) { s = new Scope(null, null); m.set(key, s); }
    return s;
  }

  // ----- statements --------------------------------------------------------------------------------------------

  private async stmt(st: Stmt, scope: Scope, ctx: Ctx): Promise<void> {
    switch (st.s) {
      case "dim":
        for (const v of st.vars) {
          if (st.isStatic || ctx.frame.proc.isStatic) {
            const ss = this.staticScope(ctx.frame.proc, ctx.frame.me);
            await this.declare(ss, v, ctx);
            const k = low(v.name);
            scope.names.set(k, v.name);
            if (v.type) scope.types.set(k, v.type);
            scope.refs.set(k, { get: () => ss.vars.get(k), set: (x) => this.store(ss, k, coerce(ss.types.get(k), x)) });
          } else await this.declare(scope, v, ctx);
        }
        return;
      case "const": { const k = low(st.name); scope.names.set(k, st.name); scope.vars.set(k, await this.eval(st.value, scope, ctx)); return; }
      case "redim":
        for (const v of st.vars) {
          if (v.target.e !== "id") throw new VbaError("ReDim nur für Variablen", 5);
          const k = low(v.target.name);
          const dims = await this.bounds(v.dims, scope, ctx);
          const owner = this.lookupScope(k, scope) ?? scope;
          const cur = owner.refs.get(k)?.get() ?? owner.vars.get(k);
          const type = (isArr(cur) ? this.elTypes.get(cur) : undefined) ?? owner.types.get(k);
          const fresh = makeArray(dims, () => this.initial(type));
          if (type) this.elTypes.set(fresh, type);
          if (st.preserve && isArr(cur)) {
            const copy = (from: VbaArray, to: VbaArray) => { for (let i = 0; i < to.items.length; i++) { const j = i + to.lb - from.lb; if (j < 0 || j >= from.items.length) continue; const x = from.items[j], y = to.items[i]; if (isArr(x) && isArr(y)) copy(x, y); else to.items[i] = x; } };
            copy(cur, fresh);
          }
          if (owner.refs.has(k)) owner.refs.get(k)!.set(fresh); else { owner.names.set(k, owner.names.get(k) ?? v.target.name); this.store(owner, k, fresh); }
        }
        return;
      case "assign": await this.assign(st.target, await this.eval(st.value, scope, ctx), scope, ctx, st.set); return;
      case "expr": {
        const x = st.value;
        // A bare name is a call of a procedure without arguments.
        await this.eval(x.e === "id" || x.e === "member" ? { e: "call", fn: x, args: [] } : x, scope, ctx, true);
        return;
      }
      case "if":
        for (const b of st.branches) if (toBool(this.prim(await this.eval(b.cond, scope, ctx)))) { await this.exec(b.body, scope, ctx); return; }
        await this.exec(st.other, scope, ctx);
        return;
      case "for": {
        const from = toNum(this.prim(await this.eval(st.from, scope, ctx))), to = toNum(this.prim(await this.eval(st.to, scope, ctx)));
        const step = st.step ? toNum(await this.eval(st.step, scope, ctx)) : 1;
        await this.assign(st.v, from, scope, ctx, false);
        try {
          for (let i = from; step >= 0 ? i <= to : i >= to; ) {
            await this.exec(st.body, scope, ctx);
            i = toNum(await this.eval(st.v, scope, ctx)) + step;
            await this.assign(st.v, i, scope, ctx, false);
            this.tick(st.line);
          }
        } catch (e) { if (!(e instanceof Exit && e.what === "for")) throw e; }
        return;
      }
      case "foreach": {
        let coll = await this.eval(st.coll, scope, ctx);
        if (coll instanceof Instance) {
          // A class with a NewEnum member (VB_UserMemId = -4) enumerates what that member returns.
          const p = [...coll.cls.p.procs.values(), ...[...coll.cls.p.props.values()].map((x) => x.get).filter((y): y is Proc => !!y)].find((q) => q.attrs.usermemid === "-4");
          if (!p) throw new VbaError("Objekt unterstützt diese Eigenschaft oder Methode nicht", 438);
          coll = await this.invoke(p, coll.cls, coll, [], undefined, scope, ctx);
        }
        if (!isArr(coll) && !(isObj(coll) && coll.items)) throw new VbaError(coll === null || coll === undefined ? "Objektvariable oder With-Blockvariable nicht festgelegt" : "Objekt unterstützt diese Eigenschaft oder Methode nicht", coll === null || coll === undefined ? 91 : 438);
        const list = isArr(coll) ? flat(coll) : (coll as VbaObject).items!();
        try {
          for (const item of list) { await this.assign(st.v, item, scope, ctx, true); await this.exec(st.body, scope, ctx); this.tick(st.line); }
        } catch (e) { if (!(e instanceof Exit && e.what === "for")) throw e; }
        return;
      }
      case "do":
        try {
          for (;;) {
            if (st.cond && !st.post) { const c = toBool(this.prim(await this.eval(st.cond, scope, ctx))); if (st.until ? c : !c) break; }
            await this.exec(st.body, scope, ctx);
            this.tick(st.line);
            if (st.cond && st.post) { const c = toBool(this.prim(await this.eval(st.cond, scope, ctx))); if (st.until ? c : !c) break; }
          }
        } catch (e) { if (!(e instanceof Exit && e.what === "do")) throw e; }
        return;
      case "select": {
        const v = this.prim(await this.eval(st.value, scope, ctx));
        const text = ctx.frame.module.p.options.compareText;
        for (const c of st.cases) {
          let hit = c.tests === null;
          for (const t of c.tests ?? []) {
            if (t.k === "eq") hit = this.cmp(v, this.prim(await this.eval(t.v, scope, ctx)), text) === 0;
            else if (t.k === "range") hit = this.cmp(v, this.prim(await this.eval(t.a, scope, ctx)), text) >= 0 && this.cmp(v, this.prim(await this.eval(t.b, scope, ctx)), text) <= 0;
            else hit = this.compareOp(t.op, v, this.prim(await this.eval(t.v, scope, ctx)), text);
            if (hit) break;
          }
          if (hit) { await this.exec(c.body, scope, ctx); return; }
        }
        return;
      }
      case "with": {
        const obj = await this.eval(st.obj, scope, ctx);
        this.retain(obj);
        ctx.withObj.push(obj);
        try { await this.exec(st.body, scope, ctx); } finally { ctx.withObj.pop(); this.release(obj, ctx.frame.depth); }
        return;
      }
      case "exit": throw new Exit(st.what);
      case "end": throw new EndSig();
      case "stop": if (this.debugOn) await this.pauseAt("stop", ctx); return;
      case "assert": if (this.debugOn && !toBool(await this.eval(st.cond, scope, ctx))) await this.pauseAt("assert", ctx); return;
      case "onerror":
        this.err.clear();
        if (st.mode === "next") { ctx.resumeNext = true; ctx.handler = null; }
        else if (st.mode === "zero") { ctx.resumeNext = false; ctx.handler = null; }
        else if (st.mode === "reset") ctx.inHandler = false;
        else { ctx.resumeNext = false; ctx.handler = st.label ?? null; }
        return;
      case "resume": throw new ResumeSig(st.mode, st.label);
      case "goto": throw new Jump(st.label, false);
      case "label": return;
      case "raise": {
        const me = ctx.frame.me;
        if (!me) throw new VbaError("RaiseEvent ist nur in Klassenmodulen möglich", 5);
        await this.raiseOn(me, st.name, st.args, st.names, scope, ctx);
        return;
      }
      case "print": {
        const text = await this.printText(st.items, scope, ctx, st.file ? (this.fileOf(toNum(await this.eval(st.file, scope, ctx)))?.col ?? 0) : this.printBuf.length);
        if (!st.file) {
          this.printBuf += text.text;
          if (text.newline) { this.printLine(this.printBuf); this.printBuf = ""; }
          return;
        }
        const f = this.openFile(toNum(await this.eval(st.file, scope, ctx)), "output");
        this.writeOut(f, text.text + (text.newline ? "\r\n" : ""));
        return;
      }
      case "write": {
        const f = this.openFile(toNum(await this.eval(st.file, scope, ctx)), "output");
        const parts: string[] = [];
        for (const it of st.items) parts.push(it ? writeValue(await this.eval(it, scope, ctx)) : "");
        this.writeOut(f, parts.join(",") + "\r\n");
        return;
      }
      case "open": await this.openStmt(toStr(await this.eval(st.path, scope, ctx)), st.mode, toNum(await this.eval(st.file, scope, ctx))); return;
      case "close":
        if (!st.files.length) { await this.closeAllFiles(); return; }
        for (const fe of st.files) await this.closeFile(toNum(await this.eval(fe, scope, ctx)));
        return;
      case "lineinput": {
        const f = this.openFile(toNum(await this.eval(st.file, scope, ctx)), "input");
        if (f.pos >= f.data.length) throw new VbaError("Einlesen hinter Dateiende", 62);
        const m = /\r\n|\n|\r/g;
        m.lastIndex = f.pos;
        const r = m.exec(f.data);
        const line = r ? f.data.slice(f.pos, r.index) : f.data.slice(f.pos);
        f.pos = r ? r.index + r[0].length : f.data.length;
        await this.assign(st.target, line, scope, ctx, false);
        return;
      }
      case "input": {
        const f = this.openFile(toNum(await this.eval(st.file, scope, ctx)), "input");
        for (const t of st.targets) {
          if (f.pos >= f.data.length) throw new VbaError("Einlesen hinter Dateiende", 62);
          await this.assign(t, readField(f), scope, ctx, false);
        }
        return;
      }
      case "name": {
        const from = toStr(await this.eval(st.from, scope, ctx)), to = toStr(await this.eval(st.to, scope, ctx));
        const fs = await this.askFile("rename", from, to);
        if (fs.rename) await fs.rename(from, to).catch(() => { throw new VbaError(`Datei nicht gefunden: ${from}`, 53); });
        else { const data = await this.readFile(fs, from); await fs.write(to, data); if (!fs.remove) throw new VbaError("Zugriff verweigert", 70); await fs.remove(from); }
        return;
      }
      case "erase":
        for (const t of st.targets) {
          if (t.e !== "id") continue;
          const k = low(t.name);
          const owner = this.lookupScope(k, scope);
          const cur = owner ? owner.refs.get(k)?.get() ?? owner.vars.get(k) : undefined;
          if (!owner || !isArr(cur)) continue;
          const type = this.elTypes.get(cur);
          const reset = (a: VbaArray) => { a.items = a.items.map((x) => { this.release(x, ctx.frame.depth); return isArr(x) ? (reset(x), x) : this.initial(type); }); };
          if (type || cur.items.length) { if (type) reset(cur); else { const empty: VbaArray = { kind: "array", lb: 0, items: [] }; if (owner.refs.has(k)) owner.refs.get(k)!.set(empty); else this.store(owner, k, empty); } }
        }
        return;
    }
  }

  private async bounds(dims: Expr[], scope: Scope, ctx: Ctx): Promise<[number, number][]> {
    const out: [number, number][] = [];
    const base = ctx.frame.module.p.options.base;
    for (const d of dims) {
      if (d.e === "bin" && d.op === "range") out.push([toNum(await this.eval(d.a, scope, ctx)), toNum(await this.eval(d.b, scope, ctx))]);
      else out.push([base, toNum(await this.eval(d, scope, ctx))]);
    }
    return out;
  }

  private async assign(target: Expr, value: VbaValue, scope: Scope, ctx: Ctx, isSet: boolean): Promise<void> {
    if (!isSet) { if (isArr(value)) value = copyArray(value); else if (value instanceof VbaRecord) value = value.copy(); else value = this.prim(value); }
    if (target.e === "id") {
      const k = low(target.name);
      if (k === "me") throw new VbaError("Ungültige Verwendung von Me", 5);
      for (let s: Scope | null = scope; s; s = s.parent) {
        const r = s.refs.get(k);
        if (r) { r.set(value); return; }
        if (s.vars.has(k)) { this.store(s, k, coerce(s.types.get(k), value)); return; }
      }
      // Inside a class or form: its own Property Let/Set or the form's properties (Caption = …).
      const me = ctx.frame.me;
      if (me && (await this.instSet(me, target.name, value, [], scope, ctx, isSet, false)) !== MISSING) return;
      // An undeclared variable (Option Explicit is not enforced): a new local variable.
      ctx.frame.scope.names.set(k, target.name);
      this.store(ctx.frame.scope, k, value);
      return;
    }
    if (target.e === "call") {
      // Array element or a parameterised property (Cells(1, 2) = 5, a(3) = x, Item(i) = v in a class).
      const base = target.fn;
      if (base.e === "id") {
        const k = low(base.name);
        const owner = this.lookupScope(k, scope);
        if (owner) {
          const cur = owner.refs.get(k)?.get() ?? owner.vars.get(k);
          if (isArr(cur)) { this.setElement(cur, (await this.args(target.args, scope, ctx)).map(toNum), value); return; }
          if (isObj(cur)) { await this.defaultSet(cur, value, target.args, scope, ctx, isSet); return; }
          throw new VbaError(`"${base.name}" ist kein Datenfeld`, 13);
        }
        const me = ctx.frame.me;
        if (me && (await this.instSet(me, base.name, value, target.args, scope, ctx, isSet, false)) !== MISSING) return;
        const obj = await this.eval(base, scope, ctx);
        if (isObj(obj)) { await this.defaultSet(obj, value, target.args, scope, ctx, isSet); return; }
      }
      if (base.e === "member") {
        const obj = await this.objOf(base.obj, scope, ctx);
        await this.memberSet(obj, base.name, value, target.args, scope, ctx, isSet);
        return;
      }
      throw new VbaError("Zuweisung nicht möglich", 424);
    }
    if (target.e === "member") {
      if (target.obj?.e === "id" && !this.lookupScope(low(target.obj.name), scope)) {
        // Modul1.Variable = 5
        const m = this.byName.get(low(target.obj.name));
        if (m && m.kind !== "class") { const s = this.lookupScope(low(target.name), m.scope); if (s) { this.store(s, low(target.name), coerce(s.types.get(low(target.name)), value)); return; } }
      }
      const obj = await this.objOf(target.obj, scope, ctx);
      await this.memberSet(obj, target.name, value, [], scope, ctx, isSet);
      return;
    }
    if (target.e === "bang") {
      const obj = await this.objOf(target.obj, scope, ctx);
      await this.defaultSet(obj, value, [lit(target.name)], scope, ctx, isSet);
      return;
    }
    throw new VbaError("Zuweisung nicht möglich", 424);
  }

  private async defaultSet(obj: VbaValue, value: VbaValue, argExprs: (Expr | null)[], scope: Scope, ctx: Ctx, isSet: boolean): Promise<void> {
    if (obj instanceof Instance) {
      const name = this.defaultMemberName(obj);
      if (!name) throw new VbaError("Objekt unterstützt diese Eigenschaft oder Methode nicht", 438);
      await this.memberSet(obj, name, value, argExprs, scope, ctx, isSet);
      return;
    }
    if (!isObj(obj)) throw new VbaError("Objekt erforderlich", 424);
    const args = await this.args(argExprs, scope, ctx);
    if (obj instanceof ControlObj) { await this.controlSet(obj, args.length ? "List" : "Value", value, args, ctx); return; }
    this.pin(value);
    obj.set?.("", value, args);
  }

  private async objOf(e: Expr | null, scope: Scope, ctx: Ctx): Promise<VbaValue> {
    if (!e) { const w = ctx.withObj[ctx.withObj.length - 1]; if (w === undefined) throw new VbaError("Ungültiger Verweis: . ohne With", 91); return w; }
    return this.eval(e, scope, ctx);
  }

  private async args(list: (Expr | null)[], scope: Scope, ctx: Ctx): Promise<VbaValue[]> {
    const out: VbaValue[] = [];
    for (const a of list) out.push(a ? await this.eval(a, scope, ctx) : undefined);
    return out;
  }
  /** Arguments put into the order a method's parameter names give (named arguments). */
  private ordered(values: VbaValue[], names: (string | null)[] | undefined, params: string[] | undefined): VbaValue[] {
    if (!names || !params) return values;
    const out: VbaValue[] = [];
    values.forEach((v, i) => {
      const nm = names[i];
      if (!nm) { out[i] = v; return; }
      const j = params.findIndex((p) => low(p) === low(nm));
      if (j < 0) throw new VbaError(`Benanntes Argument nicht gefunden: ${nm}`, 448);
      out[j] = v;
    });
    for (let i = 0; i < out.length; i++) if (!(i in out)) out[i] = undefined;
    return out;
  }

  // ----- members of objects ------------------------------------------------------------------------------------

  private defaultMemberName(inst: Instance): string | null {
    for (const p of inst.cls.p.procs.values()) if (p.attrs.usermemid === "0") return p.name;
    for (const pp of inst.cls.p.props.values()) for (const p of [pp.get, pp.let, pp.set]) if (p?.attrs.usermemid === "0") return p.name;
    return null;
  }

  /** obj.Name(args): a method or property of any object. */
  private async memberCall(obj: VbaValue, name: string, argExprs: (Expr | null)[], names: (string | null)[] | undefined, scope: Scope, ctx: Ctx, statement = false): Promise<VbaValue> {
    if (obj instanceof Instance) {
      const r = await this.instCall(obj, name, argExprs, names, scope, ctx, ctx.frame.me !== obj, statement);
      if (r === MISSING) throw new VbaError(`Objekt unterstützt diese Eigenschaft oder Methode nicht: ${name}`, 438);
      return r;
    }
    if (obj === null || obj === undefined) throw new VbaError(`Objektvariable oder With-Blockvariable nicht festgelegt (.${name})`, 91);
    if (!isObj(obj)) throw new VbaError(`Objekt erforderlich (.${name})`, 424);
    const values = await this.args(argExprs, scope, ctx);
    const params = obj.argNames?.(name);
    // Named arguments (Key:=…): by the object's parameter names into their places, and always also by name.
    let named: Record<string, VbaValue> | undefined;
    if (names?.some((n) => n)) { named = {}; names.forEach((n, i) => { if (n) named![low(n)] = values[i]; }); }
    const args = named && !params ? values.filter((_, i) => !names![i]) : this.ordered(values, names, params);
    for (const a of args) this.pin(a);
    if (named) for (const v of Object.values(named)) this.pin(v);
    if (obj.call) { const r = obj.call(name, args, named); if (r !== NOT_A_METHOD) return r instanceof Promise ? await r : r; }
    // A collection property with an index (Slides(1), Shapes("Titel")): the collection's item.
    if (args.length) { const coll = obj.get(name, []); if (isObj(coll) && (coll.default || coll.items)) return coll.default ? coll.default(args) : coll.get("Item", args, named); }
    return obj.get(name, args, named);
  }

  /** A member of an instance: member variable, Property Get, Function or Sub, the form's own members and controls,
   *  and procedures of an implemented interface (IFoo_Bar). MISSING when the instance has no such member. */
  private async instCall(inst: Instance, name: string, argExprs: (Expr | null)[], names: (string | null)[] | undefined, scope: Scope, ctx: Ctx, outside: boolean, statement: boolean): Promise<Maybe> {
    const k = low(name);
    const cls = inst.cls;
    if ((inst.scope.vars.has(k) || inst.scope.refs.has(k)) && (!outside || cls.publicVars.has(k))) {
      let v = inst.scope.refs.get(k)?.get() ?? inst.scope.vars.get(k);
      if (v === null && inst.scope.autoNew.has(k)) { v = await this.create(inst.scope.autoNew.get(k)!, ctx); this.store(inst.scope, k, v); }
      if (!argExprs.length) return v;
      const args = await this.args(argExprs, scope, ctx);
      if (isArr(v)) return getIn(v, args.map(toNum));
      return this.callDefault(v, args.map(lit), undefined, scope, ctx);
    }
    const pp = cls.p.props.get(k);
    if (pp?.get && (!outside || !pp.get.isPrivate)) return this.invoke(pp.get, cls, inst, argExprs, names, scope, ctx);
    const p = cls.p.procs.get(k);
    if (p && (!outside || !p.isPrivate)) return this.invoke(p, cls, inst, argExprs, names, scope, ctx);
    if (cls.form) {
      if (!inst.formLoaded && !inst.dead) await this.reload(inst, ctx);
      const r = await this.formCall(inst, k, argExprs, scope, ctx, statement);
      if (r !== MISSING) return r;
    }
    for (const iface of cls.p.implements) {
      const ik = `${low(iface)}_${k}`;
      const ip = cls.p.props.get(ik)?.get ?? cls.p.procs.get(ik);
      if (ip) return this.invoke(ip, cls, inst, argExprs, names, scope, ctx);
    }
    return MISSING;
  }

  /** obj.Name(args) = value. */
  private async memberSet(obj: VbaValue, name: string, value: VbaValue, argExprs: (Expr | null)[], scope: Scope, ctx: Ctx, isSet: boolean): Promise<void> {
    if (obj instanceof Instance) {
      const r = await this.instSet(obj, name, value, argExprs, scope, ctx, isSet, ctx.frame.me !== obj);
      if (r === MISSING) throw new VbaError(`Objekt unterstützt diese Eigenschaft oder Methode nicht: ${name}`, 438);
      return;
    }
    if (obj === null || obj === undefined) throw new VbaError(`Objektvariable oder With-Blockvariable nicht festgelegt (.${name})`, 91);
    if (!isObj(obj) || !obj.set) throw new VbaError(`Objekt erforderlich (.${name})`, 424);
    const args = await this.args(argExprs, scope, ctx);
    if (obj instanceof ControlObj) { await this.controlSet(obj, name, value, args, ctx); return; }
    this.pin(value);
    obj.set(name, value, args);
  }

  private async instSet(inst: Instance, name: string, value: VbaValue, argExprs: (Expr | null)[], scope: Scope, ctx: Ctx, isSet: boolean, outside: boolean): Promise<Maybe> {
    const k = low(name);
    const cls = inst.cls;
    if ((inst.scope.vars.has(k) || inst.scope.refs.has(k)) && (!outside || cls.publicVars.has(k))) {
      if (argExprs.length) {
        const cur = inst.scope.vars.get(k);
        if (isArr(cur)) { this.setElement(cur, (await this.args(argExprs, scope, ctx)).map(toNum), value); return undefined; }
        await this.defaultSet(cur, value, argExprs, scope, ctx, isSet);
        return undefined;
      }
      const r = inst.scope.refs.get(k);
      if (r) r.set(value); else this.store(inst.scope, k, coerce(inst.scope.types.get(k), value));
      return undefined;
    }
    const pp = cls.p.props.get(k);
    const setter = isSet ? pp?.set ?? pp?.let : pp?.let ?? pp?.set;
    if (setter && (!outside || !setter.isPrivate)) { await this.invoke(setter, cls, inst, argExprs, undefined, scope, ctx, value); return undefined; }
    if (cls.form) {
      if (!inst.formLoaded && !inst.dead) await this.reload(inst, ctx);
      const r = await this.formSet(inst, k, value, argExprs, scope, ctx);
      if (r !== MISSING) return r;
    }
    for (const iface of cls.p.implements) {
      const ipp = cls.p.props.get(`${low(iface)}_${k}`);
      const s = isSet ? ipp?.set ?? ipp?.let : ipp?.let ?? ipp?.set;
      if (s) { await this.invoke(s, cls, inst, argExprs, undefined, scope, ctx, value); return undefined; }
    }
    return MISSING;
  }

  /** obj(args): the default member (an array element is handled by the caller). */
  private async callDefault(v: VbaValue, argExprs: (Expr | null)[], names: (string | null)[] | undefined, scope: Scope, ctx: Ctx): Promise<VbaValue> {
    if (v instanceof Instance) {
      const name = this.defaultMemberName(v);
      if (!name) throw new VbaError("Objekt unterstützt diese Eigenschaft oder Methode nicht", 438);
      return this.memberCall(v, name, argExprs, names, scope, ctx);
    }
    if (v === null || v === undefined) throw new VbaError("Objektvariable oder With-Blockvariable nicht festgelegt", 91);
    if (!isObj(v)) throw new VbaError("Typen unverträglich", 13);
    const args = await this.args(argExprs, scope, ctx);
    return v.default ? v.default(args) : v.get("Item", args);
  }

  // ----- forms -------------------------------------------------------------------------------------------------

  private async controlSet(c: ControlObj, name: string, value: VbaValue, args: VbaValue[], ctx: Ctx): Promise<void> {
    const events = c.setQuiet(name, value, args);
    const inst = this.formInst.get(c.form);
    if (inst) for (const ev of events) await this.callEvent(inst, `${c.d.name}_${ev}`, [], ctx);
  }

  private async formCall(inst: Instance, k: string, argExprs: (Expr | null)[], scope: Scope, ctx: Ctx, statement: boolean): Promise<Maybe> {
    const st = inst.form!;
    const d = st.design;
    const args = () => this.args(argExprs, scope, ctx);
    switch (k) {
      case "show": { const a = await args(); await this.showForm(inst, a[0] === undefined ? true : toNum(a[0]) !== 0, ctx); return undefined; }
      case "hide": this.hideForm(inst); return undefined;
      case "caption": return d.caption;
      case "name": return inst.cls.name;
      case "width": return d.width + FORM_FRAME.width;
      case "height": return d.height + FORM_FRAME.height;
      case "insidewidth": return d.width;
      case "insideheight": return d.height;
      case "left": return d.left ?? 0;
      case "top": return d.top ?? 0;
      case "backcolor": return d.backColor ?? 0x8000000f;
      case "forecolor": return d.foreColor ?? 0x80000012;
      case "tag": return d.tag ?? "";
      case "startupposition": return d.startUpPosition ?? 1;
      case "visible": return st.visible;
      case "enabled": return true;
      case "font": return fontObject(() => (d.font ??= defaultFont()), () => st.changed(null, "font"));
      case "activecontrol": return null;
      case "controls": { const c = st.controlsObject(() => d.controls, d.controls); const a = await args(); return a.length ? c.default!(a) : c; }
      case "repaint": st.changed(null, "repaint"); return undefined;
      case "printform": case "setfocus": case "whatsthismode": return undefined;
      case "move": { const [l, t, w, h] = await args(); if (l !== undefined) d.left = toNum(l); if (t !== undefined) d.top = toNum(t); if (w !== undefined) d.width = toNum(w) - FORM_FRAME.width; if (h !== undefined) d.height = toNum(h) - FORM_FRAME.height; st.changed(null, "position"); return undefined; }
    }
    const c = st.control(k);
    if (c) {
      if (!argExprs.length) return c;
      return this.memberCall(c, c.d.type === "ListBox" || c.d.type === "ComboBox" ? "List" : "Value", argExprs, undefined, scope, ctx, statement);
    }
    return MISSING;
  }

  private async formSet(inst: Instance, k: string, value: VbaValue, argExprs: (Expr | null)[], scope: Scope, ctx: Ctx): Promise<Maybe> {
    const st = inst.form!;
    const d = st.design;
    const set = (f: () => void, prop: string) => { f(); st.changed(null, prop); return undefined; };
    switch (k) {
      case "caption": return set(() => { d.caption = toStr(value); }, "caption");
      case "width": return set(() => { d.width = toNum(value) - FORM_FRAME.width; }, "size");
      case "height": return set(() => { d.height = toNum(value) - FORM_FRAME.height; }, "size");
      case "insidewidth": case "insideheight": throw new VbaError("Eigenschaft ist schreibgeschützt", 382);
      case "left": return set(() => { d.left = toNum(value); }, "position");
      case "top": return set(() => { d.top = toNum(value); }, "position");
      case "backcolor": return set(() => { d.backColor = toNum(value); }, "backColor");
      case "forecolor": return set(() => { d.foreColor = toNum(value); }, "foreColor");
      case "tag": return set(() => { d.tag = toStr(value); }, "tag");
      case "startupposition": return set(() => { d.startUpPosition = toNum(value); }, "startUpPosition");
    }
    const c = st.control(k);
    if (c) { await this.controlSet(c, argExprs.length ? "List" : "Value", value, await this.args(argExprs, scope, ctx), ctx); return undefined; }
    return MISSING;
  }

  private async showForm(inst: Instance, modal: boolean, ctx: Ctx): Promise<void> {
    if (!this.host.showForm) throw new VbaError(`UserForms können in ${this.appName} hier nicht angezeigt werden.`, 5);
    const st = inst.form!;
    if (st.visible) { if (modal && !st.modal) throw new VbaError("Formular bereits angezeigt; kann nicht modal angezeigt werden", 400); return; }
    st.visible = true;
    st.modal = modal;
    this.retain(inst); // a shown form stays alive
    try { const r = this.host.showForm(st, modal); if (r && typeof (r as Promise<void>).catch === "function") (r as Promise<void>).catch(() => undefined); } catch (e) { st.hidden(); this.release(inst, ctx.frame.depth); throw e; }
    await this.callEvent(inst, "UserForm_Activate", [], ctx);
    if (modal) await st.waitHidden();
  }
  private hideForm(inst: Instance): void {
    const st = inst.form;
    if (!st?.visible) return;
    st.hidden();
    try { this.host.hideForm?.(st); } catch { /* the app's window */ }
    this.release(inst, this.curDepth);
  }
  /** Unload: UserForm_QueryClose (Cancel stops it), the form is hidden, UserForm_Terminate; a default instance is
   *  made again the next time the form's name is used. */
  private async unload(inst: Instance, ctx: Ctx, closeMode: number): Promise<boolean> {
    if (!inst.form || !inst.formLoaded) return true;
    const qa: VbaValue[] = [0, closeMode];
    await this.callEvent(inst, "UserForm_QueryClose", qa, ctx);
    if (toNum(qa[0]) !== 0) return false;
    this.hideForm(inst);
    inst.formLoaded = false;
    await this.callEvent(inst, "UserForm_Terminate", [], ctx);
    const m = inst.cls;
    if (m.defaultInstance === inst) { m.defaultInstance = null; inst.dead = true; this.release(inst, ctx.frame.depth); }
    return true;
  }

  // ----- expressions -------------------------------------------------------------------------------------------

  private cmp(a: VbaValue, b: VbaValue, text = false): number {
    if (a === null || b === null) return a === b ? 0 : a === null ? -1 : 1;
    if (typeof a === "string" && typeof b === "string") { const x = text ? a.toLowerCase() : a, y = text ? b.toLowerCase() : b; return x < y ? -1 : x > y ? 1 : 0; }
    if (a === undefined && typeof b === "string") return b === "" ? 0 : -1;
    if (b === undefined && typeof a === "string") return a === "" ? 0 : 1;
    if ((typeof a === "string" && Number.isNaN(Number(a.replace(",", ".")))) || (typeof b === "string" && Number.isNaN(Number(b.replace(",", "."))))) { let x = toStr(a), y = toStr(b); if (text) { x = x.toLowerCase(); y = y.toLowerCase(); } return x < y ? -1 : x > y ? 1 : 0; }
    const x = toNum(a), y = toNum(b);
    return x < y ? -1 : x > y ? 1 : 0;
  }
  private compareOp(op: string, a: VbaValue, b: VbaValue, text = false): boolean {
    const c = this.cmp(a, b, text);
    return op === "=" ? c === 0 : op === "<>" ? c !== 0 : op === "<" ? c < 0 : op === ">" ? c > 0 : op === "<=" ? c <= 0 : c >= 0;
  }

  private async eval(x: Expr, scope: Scope, ctx: Ctx, statement = false): Promise<VbaValue> {
    switch (x.e) {
      case "lit": return x.v;
      case "paren": return this.eval(x.a, scope, ctx);
      case "slot": return x.ref.get();
      case "id": return this.evalId(x.name, scope, ctx, statement, !!x.br);
      case "member": {
        if (x.obj?.e === "id" && !this.lookupScope(low(x.obj.name), scope)) {
          const q = await this.qualified(x.obj.name, x.name, [], undefined, scope, ctx, statement);
          if (q !== MISSING) return q;
        }
        const obj = await this.objOf(x.obj, scope, ctx);
        return this.memberCall(obj, x.name, [], undefined, scope, ctx, statement);
      }
      case "bang": { const obj = await this.objOf(x.obj, scope, ctx); return this.callDefault(obj, [lit(x.name)], undefined, scope, ctx); }
      case "new": return this.create(x.cls, ctx);
      case "typeof": return this.typeOf(await this.eval(x.a, scope, ctx), x.cls);
      case "un": {
        const a = this.prim(await this.eval(x.a, scope, ctx));
        if (x.op === "neg") return a === null ? null : -toNum(a);
        return typeof a === "boolean" ? !a : a === null ? null : ~toNum(a);
      }
      case "bin": {
        if (x.op === "range") throw new VbaError("Ungültiger Bereich", 5);
        const ra = await this.eval(x.a, scope, ctx), rb = await this.eval(x.b, scope, ctx);
        if (x.op === "is") return (ra ?? null) === (rb ?? null);
        const a = this.prim(ra), b = this.prim(rb);
        const text = ctx.frame.module.p.options.compareText;
        if ((a === null || b === null) && !["&", "and", "or"].includes(x.op)) return null;
        switch (x.op) {
          case "+": return typeof a === "string" && typeof b === "string" ? a + b : a instanceof Date ? new Date(a.getTime() + toNum(b) * 86400000) : toNum(a) + toNum(b);
          case "-": return a instanceof Date && b instanceof Date ? toNum(a) - toNum(b) : a instanceof Date ? new Date(a.getTime() - toNum(b) * 86400000) : toNum(a) - toNum(b);
          case "*": return toNum(a) * toNum(b);
          case "/": { const d = toNum(b); if (d === 0) throw new VbaError("Division durch Null", 11); return toNum(a) / d; }
          case "\\": { const d = Math.round(toNum(b)); if (d === 0) throw new VbaError("Division durch Null", 11); return Math.trunc(Math.round(toNum(a)) / d); }
          case "mod": { const d = Math.round(toNum(b)); if (d === 0) throw new VbaError("Division durch Null", 11); return Math.round(toNum(a)) % d; }
          case "^": return Math.pow(toNum(a), toNum(b));
          case "&": return toStr(a) + toStr(b);
          case "=": case "<>": case "<": case ">": case "<=": case ">=": return this.compareOp(x.op, a, b, text);
          case "like": return like(toStr(a), toStr(b), text);
          case "and": return a === null || b === null ? (a === false || b === false ? false : null) : typeof a === "boolean" && typeof b === "boolean" ? a && b : toNum(a) & toNum(b);
          case "or": return a === null || b === null ? (a === true || b === true ? true : null) : typeof a === "boolean" && typeof b === "boolean" ? a || b : toNum(a) | toNum(b);
          case "xor": return typeof a === "boolean" && typeof b === "boolean" ? a !== b : toNum(a) ^ toNum(b);
          case "eqv": return toBool(a) === toBool(b);
          case "imp": return !toBool(a) || toBool(b);
        }
        throw new VbaError(`Operator ${x.op}`, 5);
      }
      case "call": {
        const f = x.fn;
        if (f.e === "id" && !f.br) return this.callId(f.name, x.args, x.names, scope, ctx, statement);
        if (f.e === "member") {
          if (f.obj?.e === "id" && !this.lookupScope(low(f.obj.name), scope)) {
            const q = await this.qualified(f.obj.name, f.name, x.args, x.names, scope, ctx, statement);
            if (q !== MISSING) return q;
          }
          const obj = await this.objOf(f.obj, scope, ctx);
          return this.memberCall(obj, f.name, x.args, x.names, scope, ctx, statement);
        }
        const base = await this.eval(f, scope, ctx);
        if (isArr(base)) return getIn(base, (await this.args(x.args, scope, ctx)).map(toNum));
        return this.callDefault(base, x.args, x.names, scope, ctx);
      }
    }
  }

  /** Modul1.Proc, Modul1.Variable, Farbe.Rot (an Enum member), UserForm1.TextBox1. MISSING when the name is not a
   *  module or enum. */
  private async qualified(qual: string, name: string, argExprs: (Expr | null)[], names: (string | null)[] | undefined, scope: Scope, ctx: Ctx, statement: boolean): Promise<Maybe> {
    const q = low(qual), k = low(name);
    if (q === "me") return MISSING;
    for (const m of this.modules) { const e = m.p.enums.get(q); if (e && e.some((n) => low(n) === k)) return this.evalId(name, m.scope, ctx, false, false); }
    const m = this.byName.get(q);
    if (!m || m.kind === "class") return MISSING;
    const p = m.p.procs.get(k);
    if (p) return this.invoke(p, m, null, argExprs, names, scope, ctx);
    const s = this.lookupScope(k, m.scope);
    if (s) { const v = s.refs.get(k)?.get() ?? s.vars.get(k); if (!argExprs.length) return v; if (isArr(v)) return getIn(v, (await this.args(argExprs, scope, ctx)).map(toNum)); return this.callDefault(v, argExprs, names, scope, ctx); }
    if (m.kind === "document") return MISSING;
    void statement;
    throw new VbaError(`Methode oder Datenobjekt nicht gefunden: ${qual}.${name}`, 461);
  }

  private async evalId(name: string, scope: Scope, ctx: Ctx, statement: boolean, br: boolean): Promise<VbaValue> {
    const k = low(name);
    if (k === "me" && !br) {
      if (ctx.frame.me) return ctx.frame.me;
      const d = this.host.documentObject?.(ctx.frame.module.name);
      if (d) return d;
      throw new VbaError("Ungültige Verwendung des Schlüsselworts Me", 5);
    }
    for (let s: Scope | null = scope; s; s = s.parent) {
      const r = s.refs.get(k);
      if (r) return r.get();
      if (s.vars.has(k)) {
        const v = s.vars.get(k);
        if (v === null && s.autoNew.has(k)) { const o = await this.create(s.autoNew.get(k)!, ctx); this.store(s, k, o); return o; }
        return v;
      }
    }
    const me = ctx.frame.me;
    if (me) { const r = await this.instCall(me, name, [], undefined, scope, ctx, false, statement); if (r !== MISSING) return r; }
    const p = this.findProc(k, ctx.frame.module);
    if (p) return this.invoke(p.proc, p.mod, null, [], undefined, scope, ctx);
    const m = this.byName.get(k);
    if (m && m.predeclared) return this.defaultInstance(m, ctx);
    if (m && m.kind === "document") { const d = this.host.documentObject?.(m.name) ?? this.host.global(name); if (d !== undefined) return d; }
    const b = await this.builtin(k, [], undefined, statement, scope, ctx);
    if (b !== NOT_BUILTIN) return b as VbaValue;
    const g = this.host.global(br ? `[${name}]` : name);
    if (g !== undefined) return g;
    const c = CONSTANTS[k];
    if (c !== undefined) return c;
    return undefined; // an undeclared variable is Empty (Option Explicit is not enforced)
  }

  private async callId(name: string, argExprs: (Expr | null)[], names: (string | null)[] | undefined, scope: Scope, ctx: Ctx, statement: boolean): Promise<VbaValue> {
    const k = low(name);
    const owner = this.lookupScope(k, scope);
    if (owner) {
      let v = owner.refs.get(k)?.get() ?? owner.vars.get(k);
      // Inside a Function, its own name with arguments is a recursive call.
      const fp = ctx.frame.proc;
      if (owner === ctx.frame.scope && low(fp.name) === k && (fp.kind === "function" || fp.kind === "get") && !isArr(v)) return this.invoke(fp, ctx.frame.module, ctx.frame.me, argExprs, names, scope, ctx);
      if (v === null && owner.autoNew.has(k)) { v = await this.create(owner.autoNew.get(k)!, ctx); this.store(owner, k, v); }
      if (isArr(v)) return getIn(v, (await this.args(argExprs, scope, ctx)).map(toNum));
      if (isObj(v) || v === null) return this.callDefault(v, argExprs, names, scope, ctx);
      if (!argExprs.length) return v;
      throw new VbaError(`"${name}" ist kein Datenfeld`, 13);
    }
    const me = ctx.frame.me;
    if (me) { const r = await this.instCall(me, name, argExprs, names, scope, ctx, false, statement); if (r !== MISSING) return r; }
    const p = this.findProc(k, ctx.frame.module);
    if (p) return this.invoke(p.proc, p.mod, null, argExprs, names, scope, ctx);
    const args = await this.args(argExprs, scope, ctx);
    const b = await this.builtin(k, args, names, statement, scope, ctx, argExprs);
    if (b !== NOT_BUILTIN) return b as VbaValue;
    const m = this.byName.get(k);
    if (m && m.predeclared) return this.callDefault(await this.defaultInstance(m, ctx), args.map(lit), undefined, scope, ctx);
    const g = this.host.global(name);
    if (isObj(g)) return g.default ? g.default(args) : g.get("Item", args);
    throw new VbaError(`Sub oder Function nicht definiert: ${name}`, 35);
  }

  // ----- built-ins that need the runtime -----------------------------------------------------------------------

  private blocked(what: string): never { throw new VbaError(`${what} ist in ${this.appName} nicht erlaubt: Makros bleiben in der App.`, 70); }

  private async builtin(k: string, raw: VbaValue[], names: (string | null)[] | undefined, statement: boolean, scope: Scope, ctx: Ctx, argExprs?: (Expr | null)[]): Promise<VbaValue | typeof NOT_BUILTIN> {
    const a = names && BUILTIN_PARAMS[k] ? this.ordered(raw, names, BUILTIN_PARAMS[k]) : raw;
    const s = (i: number) => toStr(a[i]), n = (i: number) => toNum(a[i]);
    switch (k) {
      case "msgbox": { this.flushPrint(); const r = await this.host.msgBox(s(0), a[1] === undefined ? 0 : n(1), a[2] === undefined ? this.appName : s(2)); return statement ? undefined : r; }
      case "inputbox": return this.host.inputBox(s(0), a[1] === undefined ? this.appName : s(1), a[2] === undefined ? "" : s(2));
      case "err": return this.err;
      case "erl": return this.err.line;
      case "debug": return { kind: "object", typeName: () => "Debug", get: () => undefined, call: (nm, args) => { if (low(nm) === "print") { this.printLine(args.map((x) => this.printValue(x)).join("")); return undefined; } if (low(nm) === "assert") return undefined; return NOT_A_METHOD; } } as VbaObject;
      case "typename": return this.typeName(a[0], argExprs?.[0]?.e === "id" ? this.lookupScope(low((argExprs[0] as { name: string }).name), scope)?.types.get(low((argExprs[0] as { name: string }).name)) : undefined).replace(/^Variant\//, "");
      case "vartype": { const v = a[0]; if (isArr(v)) return 8204; if (v === undefined) return 0; if (v === null) return 1; if (typeof v === "string") return 8; if (typeof v === "boolean") return 11; if (v instanceof Date) return 7; if (typeof v === "number") return Number.isInteger(v) ? (Math.abs(v) <= 32767 ? 2 : 3) : 5; return 9; }
      case "createobject": {
        const id = low(s(0));
        if (id === "scripting.dictionary") return new VbaDictionary(this.holder);
        if (id === "vba.collection") return new VbaCollection(this.holder);
        this.blocked(`CreateObject("${s(0)}")`);
        break;
      }
      case "getobject": this.blocked("GetObject"); break;
      case "shell": this.blocked("Shell"); break;
      case "sendkeys": this.blocked("SendKeys"); break;
      case "environ": case "environ$": this.blocked("Environ"); break;
      case "appactivate": this.blocked("AppActivate"); break;
      case "chdir": case "chdrive": this.blocked(k === "chdir" ? "ChDir" : "ChDrive"); break;
      case "curdir": case "curdir$": return "";
      case "callbyname": {
        const obj = a[0], member = s(1), type = n(2);
        const rest = a.slice(3).map(lit);
        if (type === 4 || type === 8) { await this.memberSet(obj, member, a[a.length - 1], rest.slice(0, -1), scope, ctx, type === 8); return undefined; }
        return this.memberCall(obj, member, rest, undefined, scope, ctx);
      }
      case "unload": {
        const v = a[0];
        if (v instanceof Instance && v.cls.form) { await this.unload(v, ctx, 1); return undefined; }
        if (v instanceof ControlObj) { const at = v.parent.indexOf(v.d); if (at >= 0) v.parent.splice(at, 1); v.form.index(); v.form.changed(null, "controls"); return undefined; }
        throw new VbaError("Objekt kann nicht geladen oder entladen werden", 361);
      }
      case "load": { const v = a[0]; if (v instanceof Instance && v.cls.form) { await this.reload(v, ctx); return undefined; } throw new VbaError("Objekt kann nicht geladen oder entladen werden", 361); }
      // Settings: the app's store (or memory for this run), never the Windows registry.
      case "getsetting": { const v = this.host.settings ? this.host.settings.get(s(0), s(1), s(2)) : this.memSettings.get(`${low(s(0))}\u0000${low(s(1))}\u0000${low(s(2))}`); return v ?? (a[3] === undefined ? "" : s(3)); }
      case "savesetting": { if (this.host.settings) this.host.settings.set(s(0), s(1), s(2), s(3)); else this.memSettings.set(`${low(s(0))}\u0000${low(s(1))}\u0000${low(s(2))}`, s(3)); return undefined; }
      case "deletesetting": {
        if (this.host.settings) this.host.settings.delete(s(0), s(1), a[2] === undefined ? undefined : s(2));
        else { const pre = `${low(s(0))}\u0000${low(s(1))}\u0000`; for (const key of [...this.memSettings.keys()]) if (a[2] === undefined ? key.startsWith(pre) : key === pre + low(s(2))) this.memSettings.delete(key); }
        return undefined;
      }
      case "getallsettings": {
        let pairs: [string, string][] = [];
        if (this.host.settings) pairs = this.host.settings.all?.(s(0), s(1)) ?? [];
        else { const pre = `${low(s(0))}\u0000${low(s(1))}\u0000`; pairs = [...this.memSettings.entries()].filter(([key]) => key.startsWith(pre)).map(([key, v]) => [key.slice(pre.length), v]); }
        if (!pairs.length) return undefined;
        return { kind: "array", lb: 0, items: pairs.map(([key, v]) => ({ kind: "array", lb: 0, items: [key, v] }) as VbaArray) };
      }
      // Files, through the app and after asking (D-080).
      case "freefile": { for (let i = 1; i < 256; i++) if (!this.files.has(i)) return i; throw new VbaError("Zu viele Dateien", 67); }
      case "eof": { const f = this.fileOf(n(0)); if (!f) throw new VbaError("Dateiname oder -nummer falsch", 52); return f.mode === "input" ? f.pos >= f.data.length : true; }
      case "lof": { const f = this.fileOf(n(0)); if (!f) throw new VbaError("Dateiname oder -nummer falsch", 52); return f.mode === "input" ? encodeCp1252(f.data).length : encodeCp1252(f.out).length; }
      case "loc": { const f = this.fileOf(n(0)); if (!f) throw new VbaError("Dateiname oder -nummer falsch", 52); return Math.ceil(f.pos / 128); }
      case "reset": await this.closeAllFiles(); return undefined;
      case "dir": case "dir$": return this.dir(a[0] === undefined ? undefined : s(0), a[1] === undefined ? 0 : n(1));
      case "filelen": { const fs = await this.askFile("list", s(0)); if (!fs.stat) throw new VbaError("Zugriff verweigert", 70); const st = await fs.stat(s(0)); if (!st) throw new VbaError(`Datei nicht gefunden: ${s(0)}`, 53); return st.size; }
      case "filedatetime": { const fs = await this.askFile("list", s(0)); if (!fs.stat) throw new VbaError("Zugriff verweigert", 70); const st = await fs.stat(s(0)); if (!st) throw new VbaError(`Datei nicht gefunden: ${s(0)}`, 53); return st.modified ?? new Date(0); }
      case "getattr": { const fs = await this.askFile("list", s(0)); if (!fs.stat) throw new VbaError("Zugriff verweigert", 70); const st = await fs.stat(s(0)); if (!st) throw new VbaError(`Datei nicht gefunden: ${s(0)}`, 53); return st.dir ? 16 : 32; }
      case "kill": {
        const path = s(0);
        const fs = await this.askFile("delete", path);
        if (!fs.remove) throw new VbaError("Zugriff verweigert", 70);
        if (/[*?]/.test(path)) {
          const { folder, mask } = splitPath(path);
          if (!fs.list) throw new VbaError("Zugriff verweigert", 70);
          const hits = (await fs.list(folder)).filter((e) => !e.dir && like(e.name, mask, true));
          if (!hits.length) throw new VbaError(`Datei nicht gefunden: ${path}`, 53);
          for (const e of hits) await fs.remove(joinPath(folder, e.name));
        } else await fs.remove(path).catch(() => { throw new VbaError(`Datei nicht gefunden: ${path}`, 53); });
        return undefined;
      }
      case "mkdir": { const fs = await this.askFile("mkdir", s(0)); if (!fs.mkdir) throw new VbaError("Zugriff verweigert", 70); await fs.mkdir(s(0)).catch(() => { throw new VbaError(`Fehler beim Zugriff auf Pfad/Datei: ${s(0)}`, 75); }); return undefined; }
      case "rmdir": { const fs = await this.askFile("rmdir", s(0)); if (!fs.rmdir) throw new VbaError("Zugriff verweigert", 70); await fs.rmdir(s(0)).catch(() => { throw new VbaError(`Fehler beim Zugriff auf Pfad/Datei: ${s(0)}`, 75); }); return undefined; }
      case "filecopy": { const fs = await this.askFile("copy", s(0), s(1)); const data = await this.readFile(fs, s(0)); await fs.write(s(1), data); return undefined; }
    }
    const r = pureBuiltin(k, a, ctx.frame.module.p.options.compareText, ctx.frame.module.p.options.base);
    return r;
  }

  // ----- printing ----------------------------------------------------------------------------------------------

  /** A value as Print shows it: numbers with a leading space (the sign's place) and a trailing space. */
  private printValue(v: VbaValue): string {
    if (typeof v === "number") return (v >= 0 ? " " : "") + toStr(v) + " ";
    if (v === null) return "Null";
    if (typeof v === "boolean") return v ? "Wahr" : "Falsch";
    return toStr(this.prim(v));
  }
  private async printText(items: PrintItem[], scope: Scope, ctx: Ctx, startCol: number): Promise<{ text: string; newline: boolean }> {
    let out = "";
    let col = startCol;
    const add = (t: string) => { out += t; const nl = t.lastIndexOf("\n"); col = nl >= 0 ? t.length - nl - 1 : col + t.length; };
    for (const it of items) {
      if (it.x) {
        const x = it.x;
        if (x.e === "call" && x.fn.e === "id" && (low(x.fn.name) === "spc" || low(x.fn.name) === "tab")) {
          const nArg = x.args[0] ? toNum(await this.eval(x.args[0], scope, ctx)) : 0;
          if (low(x.fn.name) === "spc") add(" ".repeat(Math.max(0, nArg)));
          else { const target = Math.max(1, nArg) - 1; add(col > target ? "\r\n" + " ".repeat(target) : " ".repeat(target - col)); }
        } else add(this.printValue(await this.eval(x, scope, ctx)));
      }
      if (it.sep === ",") add(" ".repeat(14 - (col % 14)));
    }
    const last = items[items.length - 1];
    return { text: out, newline: !last || last.sep === "" };
  }
  private printLine(line: string): void {
    this.printed.push(line);
    try { this.host.debugPrint?.(line); } catch { /* the app's window */ }
  }
  private flushPrint(): void { if (this.printBuf) { const b = this.printBuf; this.printBuf = ""; this.printLine(b); } }

  // ----- files -------------------------------------------------------------------------------------------------

  /** Asks the app (and through it the user) for this access; refused or not offered is Err 70. */
  private async askFile(kind: VbaFileAccessKind, path: string, target?: string): Promise<VbaFileSystem> {
    const fs = this.host.files;
    if (!fs || !this.host.fileAccess) throw new VbaError(`Zugriff verweigert: ${path} (Makros dürfen in ${this.appName} hier keine Dateien öffnen)`, 70);
    let ok = false;
    try { ok = await this.host.fileAccess(kind, path, target); } catch { ok = false; }
    if (!ok) throw new VbaError(`Zugriff verweigert: ${path}`, 70);
    return fs;
  }
  private async readFile(fs: VbaFileSystem, path: string): Promise<Uint8Array> {
    try { return await fs.read(path); } catch { throw new VbaError(`Datei nicht gefunden: ${path}`, 53); }
  }
  private fileOf(n: number): OpenFile | undefined { return this.files.get(Math.round(n)); }
  private openFile(n: number, want: "input" | "output"): OpenFile {
    const f = this.fileOf(n);
    if (!f) throw new VbaError("Dateiname oder -nummer falsch", 52);
    if ((want === "input") !== (f.mode === "input")) throw new VbaError("Dateimodus falsch", 54);
    return f;
  }
  private writeOut(f: OpenFile, text: string): void {
    f.out += text;
    f.dirty = true;
    const nl = text.lastIndexOf("\n");
    f.col = nl >= 0 ? text.length - nl - 1 : f.col + text.length;
  }
  private async openStmt(path: string, mode: "input" | "output" | "append" | "binary" | "random", n: number): Promise<void> {
    n = Math.round(n);
    if (n < 1 || n > 511) throw new VbaError("Dateiname oder -nummer falsch", 52);
    if (this.files.has(n)) throw new VbaError("Datei bereits geöffnet", 55);
    if (mode === "binary" || mode === "random") throw new VbaError(`Open … For ${mode === "binary" ? "Binary" : "Random"} wird in ${this.appName} nicht unterstützt`, 54);
    for (const f of this.files.values()) if (f.path === path && (f.mode !== "input" || mode !== "input")) throw new VbaError("Datei bereits geöffnet", 55);
    const fs = await this.askFile(mode === "input" ? "open" : mode === "append" ? "append" : "save", path);
    if (mode === "input") { const data = decodeCp1252(await this.readFile(fs, path)); this.files.set(n, { path, mode, data, pos: 0, out: "", dirty: false, col: 0 }); return; }
    let out = "";
    if (mode === "append") { try { out = decodeCp1252(await fs.read(path)); } catch { out = ""; } }
    // Output creates (or empties) the file at once, as VBA does.
    this.files.set(n, { path, mode, data: "", pos: 0, out, dirty: true, col: 0 });
    await this.flushFile(this.files.get(n)!);
  }
  private async flushFile(f: OpenFile): Promise<void> {
    if (f.mode === "input" || !f.dirty) return;
    f.dirty = false;
    await this.host.files!.write(f.path, encodeCp1252(f.out));
  }
  private async flushFiles(): Promise<void> { for (const f of this.files.values()) await this.flushFile(f); }
  private async closeFile(n: number): Promise<void> {
    const f = this.fileOf(n);
    if (!f) return;
    this.files.delete(Math.round(n));
    await this.flushFile(f);
  }
  private async closeAllFiles(): Promise<void> { for (const n of [...this.files.keys()]) await this.closeFile(n); }
  private async dir(pattern: string | undefined, attrs: number): Promise<string> {
    if (pattern === undefined) {
      if (!this.dirList) throw new VbaError("Ungültiger Prozeduraufruf oder ungültiges Argument", 5);
      return this.dirList.shift() ?? "";
    }
    const { folder, mask } = splitPath(pattern);
    const fs = await this.askFile("list", pattern);
    if (!/[*?]/.test(mask) && mask) {
      // One file: its name when it exists.
      const st = fs.stat ? await fs.stat(pattern) : fs.list ? (await fs.list(folder)).find((e) => low(e.name) === low(mask)) ?? null : null;
      this.dirList = [];
      return st && (!st.dir || attrs & 16) ? mask : "";
    }
    if (!fs.list) throw new VbaError("Zugriff verweigert", 70);
    const entries = await fs.list(folder).catch(() => [] as { name: string; dir: boolean }[]);
    this.dirList = entries.filter((e) => (!e.dir || attrs & 16) && like(e.name, mask || "*", true)).map((e) => e.name);
    return this.dirList.shift() ?? "";
  }
}

function splitPath(p: string): { folder: string; mask: string } {
  const i = Math.max(p.lastIndexOf("\\"), p.lastIndexOf("/"));
  return i < 0 ? { folder: "", mask: p } : { folder: p.slice(0, i), mask: p.slice(i + 1) };
}
function joinPath(folder: string, name: string): string { return folder ? `${folder}${folder.includes("/") && !folder.includes("\\") ? "/" : "\\"}${name}` : name; }

/** A value as Write # writes it (strings quoted, "." as the decimal point, #TRUE#, #2024-01-31#). */
function writeValue(v: VbaValue): string {
  if (typeof v === "string") return `"${v}"`;
  if (typeof v === "number") return String(v);
  if (typeof v === "boolean") return v ? "#TRUE#" : "#FALSE#";
  if (v === null) return "#NULL#";
  if (v === undefined) return "";
  if (v instanceof Date) { const p = (x: number) => String(x).padStart(2, "0"); return `#${v.getFullYear()}-${p(v.getMonth() + 1)}-${p(v.getDate())}${v.getHours() || v.getMinutes() || v.getSeconds() ? ` ${p(v.getHours())}:${p(v.getMinutes())}:${p(v.getSeconds())}` : ""}#`; }
  return `"${toStr(v)}"`;
}
/** One field for Input #: a quoted string, #…# literals, or text up to the next comma or line end. */
function readField(f: OpenFile): VbaValue {
  const d = f.data;
  while (f.pos < d.length && (d[f.pos] === " " || d[f.pos] === "\t")) f.pos++;
  let v: VbaValue;
  if (d[f.pos] === '"') {
    const end = d.indexOf('"', f.pos + 1);
    v = d.slice(f.pos + 1, end < 0 ? d.length : end);
    f.pos = end < 0 ? d.length : end + 1;
    while (f.pos < d.length && d[f.pos] !== "," && d[f.pos] !== "\r" && d[f.pos] !== "\n") f.pos++;
  } else {
    let end = f.pos;
    while (end < d.length && d[end] !== "," && d[end] !== "\r" && d[end] !== "\n") end++;
    const t = d.slice(f.pos, end).trim();
    f.pos = end;
    const lt = t.toLowerCase();
    v = lt === "#true#" ? true : lt === "#false#" ? false : lt === "#null#" ? null : /^#.+#$/.test(t) ? toDate(t.slice(1, -1)) : t !== "" && !Number.isNaN(Number(t)) ? Number(t) : t;
  }
  if (d[f.pos] === ",") f.pos++;
  else { if (d[f.pos] === "\r") f.pos++; if (d[f.pos] === "\n") f.pos++; }
  return v;
}

export { toStr as vbaToString, toNum as vbaToNumber };
