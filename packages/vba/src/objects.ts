// VBA's own objects in our interpreter: Collection, Scripting.Dictionary (kept in memory, no file system access),
// the Err object and the values of user-defined types (Type … End Type). They are ordinary VbaObjects, so an
// app's object model can take them as arguments.

import { NOT_A_METHOD, VbaError, isArr, isObj, toNum, toStr, type VbaObject, type VbaValue } from "./values";

/** Keeps the reference counts of class instances (Class_Terminate) while objects hold them. */
export interface Holder { retain(v: VbaValue): void; release(v: VbaValue): void }
const NO_HOLDER: Holder = { retain: () => undefined, release: () => undefined };

const low = (s: string) => s.toLowerCase();

/** VBA.Collection: items with optional unique keys (compared without case), 1-based. */
export class VbaCollection implements VbaObject {
  readonly kind = "object";
  readonly entries: { key?: string; value: VbaValue }[] = [];
  constructor(private holder: Holder = NO_HOLDER) {}
  typeName(): string { return "Collection"; }
  items(): VbaValue[] { return this.entries.map((e) => e.value); }
  argNames(name: string): string[] | undefined {
    switch (low(name)) { case "add": return ["Item", "Key", "Before", "After"]; case "item": case "remove": return ["Index"]; }
    return undefined;
  }
  private index(k: VbaValue): number {
    if (typeof k === "string") {
      const i = this.entries.findIndex((e) => e.key !== undefined && low(e.key) === low(k));
      if (i < 0) throw new VbaError("Ungültiger Prozeduraufruf oder ungültiges Argument", 5);
      return i;
    }
    const n = Math.round(toNum(k));
    if (n < 1 || n > this.entries.length) throw new VbaError("Index außerhalb des gültigen Bereichs", 9);
    return n - 1;
  }
  default(args: VbaValue[]): VbaValue { return this.entries[this.index(args[0])].value; }
  get(name: string, args: VbaValue[]): VbaValue {
    switch (low(name)) {
      case "count": return this.entries.length;
      case "item": return this.default(args);
      case "_newenum": return this;
    }
    const r = this.call(name, args);
    if (r !== NOT_A_METHOD) return r;
    throw new VbaError(`Objekt unterstützt diese Eigenschaft oder Methode nicht: ${name}`, 438);
  }
  call(name: string, args: VbaValue[]): VbaValue {
    switch (low(name)) {
      case "add": {
        const [value, key, before, after] = args;
        const k = key === undefined || key === null ? undefined : toStr(key);
        if (k !== undefined && this.entries.some((e) => e.key !== undefined && low(e.key) === low(k))) throw new VbaError("Dieser Schlüssel ist bereits einem Element der Auflistung zugeordnet", 457);
        let at = this.entries.length;
        if (before !== undefined && before !== null) at = this.index(before);
        else if (after !== undefined && after !== null) at = this.index(after) + 1;
        this.holder.retain(value);
        this.entries.splice(at, 0, { key: k, value });
        return undefined;
      }
      case "remove": { const i = this.index(args[0]); const [e] = this.entries.splice(i, 1); this.holder.release(e.value); return undefined; }
      case "item": return this.default(args);
      case "count": return this.entries.length;
    }
    return NOT_A_METHOD;
  }
}

/** Scripting.Dictionary in memory: keys of any type (strings compared with or without case, CompareMode). */
export class VbaDictionary implements VbaObject {
  readonly kind = "object";
  private map = new Map<unknown, { key: VbaValue; value: VbaValue }>();
  compareMode = 0;
  constructor(private holder: Holder = NO_HOLDER) {}
  typeName(): string { return "Dictionary"; }
  private norm(k: VbaValue): unknown {
    if (typeof k === "string") return this.compareMode === 1 ? "s:" + low(k) : "s:" + k;
    if (typeof k === "number") return "n:" + k;
    if (typeof k === "boolean") return "n:" + (k ? -1 : 0);
    if (k instanceof Date) return "d:" + k.getTime();
    if (k === undefined) return "e:";
    if (k === null) return "null:";
    return k; // objects by identity
  }
  items(): VbaValue[] { return [...this.map.values()].map((e) => e.key); }
  argNames(name: string): string[] | undefined {
    switch (low(name)) { case "add": return ["Key", "Item"]; case "exists": case "remove": case "item": return ["Key"]; }
    return undefined;
  }
  default(args: VbaValue[]): VbaValue { return this.item(args[0]); }
  private item(k: VbaValue): VbaValue {
    const e = this.map.get(this.norm(k));
    if (e) return e.value;
    // Reading a missing key adds it with an empty item, as Scripting.Dictionary does.
    this.map.set(this.norm(k), { key: k, value: undefined });
    return undefined;
  }
  get(name: string, args: VbaValue[]): VbaValue {
    switch (low(name)) {
      case "count": return this.map.size;
      case "item": return this.item(args[0]);
      case "comparemode": return this.compareMode;
      case "_newenum": return this;
    }
    const r = this.call(name, args);
    if (r !== NOT_A_METHOD) return r;
    throw new VbaError(`Objekt unterstützt diese Eigenschaft oder Methode nicht: ${name}`, 438);
  }
  set(name: string, value: VbaValue, args: VbaValue[]): void {
    switch (low(name)) {
      case "": case "item": {
        const n = this.norm(args[0]);
        const e = this.map.get(n);
        this.holder.retain(value);
        if (e) { const old = e.value; e.value = value; this.holder.release(old); } else this.map.set(n, { key: args[0], value });
        return;
      }
      case "key": {
        const e = this.map.get(this.norm(args[0]));
        if (!e) throw new VbaError("Element nicht gefunden", 32811);
        if (this.map.has(this.norm(value))) throw new VbaError("Dieser Schlüssel ist bereits einem Element dieser Auflistung zugeordnet", 457);
        this.map.delete(this.norm(args[0]));
        this.map.set(this.norm(value), { key: value, value: e.value });
        return;
      }
      case "comparemode":
        if (this.map.size) throw new VbaError("Ungültiger Prozeduraufruf oder ungültiges Argument", 5);
        this.compareMode = toNum(value);
        return;
    }
    throw new VbaError(`Eigenschaft kann nicht gesetzt werden: ${name}`, 438);
  }
  call(name: string, args: VbaValue[]): VbaValue {
    switch (low(name)) {
      case "add": {
        const n = this.norm(args[0]);
        if (this.map.has(n)) throw new VbaError("Dieser Schlüssel ist bereits einem Element dieser Auflistung zugeordnet", 457);
        this.holder.retain(args[1]);
        this.map.set(n, { key: args[0], value: args[1] });
        return undefined;
      }
      case "exists": return this.map.has(this.norm(args[0]));
      case "keys": return { kind: "array", lb: 0, items: [...this.map.values()].map((e) => e.key) };
      case "items": return { kind: "array", lb: 0, items: [...this.map.values()].map((e) => e.value) };
      case "remove": {
        const n = this.norm(args[0]);
        const e = this.map.get(n);
        if (!e) throw new VbaError("Element nicht gefunden", 32811);
        this.map.delete(n);
        this.holder.release(e.value);
        return undefined;
      }
      case "removeall": { for (const e of this.map.values()) this.holder.release(e.value); this.map.clear(); return undefined; }
    }
    return NOT_A_METHOD;
  }
}

/** The Err object: the last error (Number, Description, Source), Raise and Clear. */
export class VbaErr implements VbaObject {
  readonly kind = "object";
  number = 0;
  description = "";
  source = "";
  /** The line the error was raised in (Erl). */
  line = 0;
  constructor(private project = "VBAProject") {}
  typeName(): string { return "ErrObject"; }
  default(): VbaValue { return this.number; }
  clear(): void { this.number = 0; this.description = ""; this.source = ""; this.line = 0; }
  take(e: VbaError): void { this.number = e.number; this.description = e.message; this.source = this.project; this.line = e.line; }
  get(name: string, args: VbaValue[]): VbaValue {
    switch (low(name)) {
      case "number": return this.number;
      case "description": return this.description;
      case "source": return this.source;
      case "helpfile": return "";
      case "helpcontext": return 0;
      case "lastdllerror": return 0;
    }
    const r = this.call(name, args);
    if (r !== NOT_A_METHOD) return r;
    throw new VbaError(`Objekt unterstützt diese Eigenschaft oder Methode nicht: ${name}`, 438);
  }
  set(name: string, value: VbaValue): void {
    switch (low(name)) {
      case "number": this.number = toNum(value); return;
      case "description": this.description = toStr(value); return;
      case "source": this.source = toStr(value); return;
      case "helpfile": case "helpcontext": return;
    }
    throw new VbaError(`Eigenschaft kann nicht gesetzt werden: ${name}`, 438);
  }
  argNames(name: string): string[] | undefined { return low(name) === "raise" ? ["Number", "Source", "Description", "HelpFile", "HelpContext"] : undefined; }
  call(name: string, args: VbaValue[]): VbaValue {
    switch (low(name)) {
      case "clear": this.clear(); return undefined;
      case "raise": {
        const n = toNum(args[0]);
        const desc = args[2] === undefined ? (n === this.number && this.description ? this.description : ERROR_TEXT[n] ?? "Anwendungs- oder objektdefinierter Fehler") : toStr(args[2]);
        const e = new VbaError(desc, n);
        if (args[1] !== undefined) (e as VbaError & { source?: string }).source = toStr(args[1]);
        throw e;
      }
    }
    return NOT_A_METHOD;
  }
}

/** VBA's messages for its common error numbers (Error(n), Err.Raise without a description). */
export const ERROR_TEXT: Record<number, string> = {
  3: "Return ohne GoSub", 5: "Ungültiger Prozeduraufruf oder ungültiges Argument", 6: "Überlauf", 7: "Nicht genügend Speicher",
  9: "Index außerhalb des gültigen Bereichs", 10: "Datenfeld ist unveränderlich oder vorübergehend gesperrt", 11: "Division durch Null",
  13: "Typen unverträglich", 20: "Resume ohne Fehler", 28: "Nicht genügend Stapelspeicher", 35: "Sub oder Function nicht definiert",
  52: "Dateiname oder -nummer falsch", 53: "Datei nicht gefunden", 54: "Dateimodus falsch", 55: "Datei bereits geöffnet",
  57: "Geräte-E/A-Fehler", 58: "Datei existiert bereits", 61: "Datenträger voll", 62: "Einlesen hinter Dateiende",
  67: "Zu viele Dateien", 68: "Gerät nicht verfügbar", 70: "Zugriff verweigert", 71: "Datenträger nicht bereit",
  75: "Fehler beim Zugriff auf Pfad/Datei", 76: "Pfad nicht gefunden", 91: "Objektvariable oder With-Blockvariable nicht festgelegt",
  92: "For-Schleife nicht initialisiert", 93: "Ungültige Musterzeichenfolge", 94: "Unzulässige Verwendung von Null",
  424: "Objekt erforderlich", 429: "Objekterstellung durch ActiveX-Komponente nicht möglich", 438: "Objekt unterstützt diese Eigenschaft oder Methode nicht",
  440: "Automatisierungsfehler", 445: "Objekt unterstützt diese Aktion nicht", 448: "Benanntes Argument nicht gefunden", 449: "Argument ist nicht optional",
  450: "Falsche Anzahl an Argumenten oder ungültige Eigenschaftszuweisung", 451: "Property Let-Prozedur nicht definiert und Property Get-Prozedur hat kein Objekt zurückgegeben",
  457: "Dieser Schlüssel ist bereits einem Element dieser Auflistung zugeordnet", 459: "Objekt oder Klasse unterstützt diese Ereignismenge nicht",
};

/** A value of a user-defined type (Type Punkt … End Type): fields by name; assigning it copies it. */
export class VbaRecord implements VbaObject {
  readonly kind = "object";
  constructor(readonly type: string, readonly fields: Map<string, VbaValue>, private names: Map<string, string>) {}
  typeName(): string { return this.type; }
  copy(): VbaRecord {
    const f = new Map<string, VbaValue>();
    for (const [k, v] of this.fields) f.set(k, v instanceof VbaRecord ? v.copy() : isArr(v) ? { kind: "array", lb: v.lb, items: v.items.map((x) => (x instanceof VbaRecord ? x.copy() : x)) } : v);
    return new VbaRecord(this.type, f, this.names);
  }
  get(name: string, args: VbaValue[]): VbaValue {
    const k = low(name);
    if (!this.fields.has(k)) throw new VbaError(`Methode oder Datenobjekt nicht gefunden: ${name}`, 461);
    const v = this.fields.get(k);
    if (args.length && isArr(v)) { let cur: VbaValue = v; for (const i of args) { if (!isArr(cur)) break; cur = cur.items[toNum(i) - cur.lb]; } return cur; }
    return v;
  }
  set(name: string, value: VbaValue, args: VbaValue[]): void {
    const k = low(name);
    if (!this.fields.has(k)) throw new VbaError(`Methode oder Datenobjekt nicht gefunden: ${name}`, 461);
    const v = this.fields.get(k);
    if (args.length && isArr(v)) { let cur = v; for (let n = 0; n < args.length - 1; n++) cur = cur.items[toNum(args[n]) - cur.lb] as typeof v; cur.items[toNum(args[args.length - 1]) - cur.lb] = value; return; }
    this.fields.set(k, value instanceof VbaRecord ? value.copy() : value);
  }
  /** Field names as declared (for the Lokal-Fenster). */
  fieldNames(): string[] { return [...this.fields.keys()].map((k) => this.names.get(k) ?? k); }
}

export const isRecord = (v: VbaValue): v is VbaRecord => v instanceof VbaRecord;
export const isCollectionLike = (v: VbaValue): boolean => isObj(v) && !!(v as VbaObject).items;
