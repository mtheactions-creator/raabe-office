// Blatt schützen as Excel does it: the sheet's <sheetProtection> attributes (what users may still
// do, the password's salted SHA-512 hash), cells locked by default (Zellen formatieren > Schutz).
// On a protected sheet locked cells cannot be changed, formatting, inserting, deleting and
// sorting follow the allowed actions, Tab moves between unlocked cells, and Blatt schützen
// becomes Blattschutz aufheben.

import { h, messageBox, showDialog } from "@raabe/shell";
import type { Rect, Sheet } from "@raabe/sheet-engine";
import type { CalcApp } from "./app";

/** Blatt schützen's choices: label, the file's attribute, and whether Excel allows it by default. */
const CHOICES: [string, string, boolean][] = [
  ["Gesperrte Zellen auswählen", "selectLockedCells", true],
  ["Nicht gesperrte Zellen auswählen", "selectUnlockedCells", true],
  ["Zellen formatieren", "formatCells", false],
  ["Spalten formatieren", "formatColumns", false],
  ["Zeilen formatieren", "formatRows", false],
  ["Spalten einfügen", "insertColumns", false],
  ["Zeilen einfügen", "insertRows", false],
  ["Hyperlinks einfügen", "insertHyperlinks", false],
  ["Spalten löschen", "deleteColumns", false],
  ["Zeilen löschen", "deleteRows", false],
  ["Sortieren", "sort", false],
  ["AutoFilter verwenden", "autoFilter", false],
  ["PivotTable und PivotChart verwenden", "pivotTables", false],
  ["Objekte bearbeiten", "objects", false],
  ["Szenarios bearbeiten", "scenarios", false],
];

/** What a command needs on a protected sheet (null: always allowed). */
export function needs(id: string): string | null {
  if (/^(font|align|number|border|styles)\./.test(id) || /^(cells\.(formatDialog|lock)|edit\.(formatPainter|clearFormats|pasteFormats))$/.test(id)) return "formatCells";
  if (/^cells\.(rowHeight|autofitRows|hideRows|unhideRows)$/.test(id)) return "formatRows";
  if (/^cells\.(colWidth|autofitCols|hideCols|unhideCols|defaultWidth)$/.test(id)) return "formatColumns";
  if (id === "cells.insertRows") return "insertRows";
  if (id === "cells.insertCols") return "insertColumns";
  if (id === "cells.deleteRows") return "deleteRows";
  if (id === "cells.deleteCols") return "deleteColumns";
  if (/^cells\.(insert|insertDialog|delete|insertMenu|deleteMenu)$/.test(id) || /^align\.(merge|mergeCenter|mergeAcross|unmerge)$/.test(id)) return "never";
  if (/^(data\.(sort|sortAsc|sortDesc)|edit\.(sortAscMenu|sortDescMenu))$/.test(id)) return "sort";
  if (/^data\.(filter|filterClear|filterReapply)$/.test(id)) return "autoFilter";
  if (id === "insert.link") return "insertHyperlinks";
  if (/^insert\.(table|chart\w*|note)$/.test(id) || /^data\.(textToColumns|removeDuplicates|advancedFilter)$/.test(id) || id === "styles.table") return "never";
  // Tonight's additions (D-065): changes Excel does not allow on a protected sheet.
  if (/^data\.(group|ungroup|subtotal|goalSeek|validation|circleInvalid)$/.test(id) || /^layout\.break(Insert|Remove|Reset)$/.test(id) || id === "review.spelling") return "never";
  if (/^insert\.picture\w*$/.test(id) || /^(insert\.comment|review\.(newComment|deleteComment))$/.test(id)) return "objects";
  if (/^edit\.(paste\w*|cut|clear\w*|fill\w+|auto\w+|dateToday|timeNow|flashFill|removeLinks)$/.test(id)) return "unlocked";
  return null;
}

const utf16 = (s: string) => { const b = new Uint8Array(s.length * 2); for (let i = 0; i < s.length; i++) { b[i * 2] = s.charCodeAt(i) & 255; b[i * 2 + 1] = s.charCodeAt(i) >> 8; } return b; };
const b64 = (b: Uint8Array) => btoa(String.fromCharCode(...b));
const unb64 = (s: string) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));

/** Excel's password hash: SHA-512 over salt + password (UTF-16LE), then `spin` rounds over hash + round number. */
async function sha512(password: string, salt: Uint8Array, spin: number): Promise<string> {
  const join = (a: Uint8Array, b: Uint8Array) => { const o = new Uint8Array(a.length + b.length); o.set(a); o.set(b, a.length); return o; };
  let hash = new Uint8Array(await crypto.subtle.digest("SHA-512", join(salt, utf16(password))));
  const it = new Uint8Array(4);
  for (let i = 0; i < spin; i++) {
    new DataView(it.buffer).setUint32(0, i, true);
    hash = new Uint8Array(await crypto.subtle.digest("SHA-512", join(hash, it)));
  }
  return b64(hash);
}

/** The old 16-bit password hash of <sheetProtection password="..."> (files from older Excel versions). */
function legacyHash(password: string): string {
  let h = 0;
  for (let i = password.length - 1; i >= 0; i--) {
    h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
    h ^= password.charCodeAt(i) & 0xff;
  }
  h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
  h ^= password.length;
  h ^= 0xce4b;
  return h.toString(16).toUpperCase().padStart(4, "0");
}

/** The schema's attribute order, which Excel follows. */
const ORDER = ["password", "algorithmName", "hashValue", "saltValue", "spinCount", "sheet", "objects", "scenarios", "formatCells", "formatColumns", "formatRows",
  "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns", "deleteRows", "selectLockedCells", "sort", "autoFilter", "pivotTables", "selectUnlockedCells"];

/**
 * <sheetProtection> as Excel writes it: the password's attributes, sheet="1", objects and
 * scenarios and the select* flags where they are protected (1), the other actions where they are
 * allowed (0).
 */
export function protectionAttrs(allowed: Map<string, boolean>, password: Record<string, string> = {}): Record<string, string> {
  const a: Record<string, string> = { ...password, sheet: "1" };
  for (const [attr, on] of allowed) {
    if (attr.startsWith("select") || attr === "objects" || attr === "scenarios") { if (!on) a[attr] = "1"; }
    else if (on) a[attr] = "0";
  }
  return Object.fromEntries(ORDER.filter((k) => a[k] !== undefined).map((k) => [k, a[k]]));
}

export function protectionFeature(app: CalcApp) {
  const on = (s: Sheet = app.sheet) => s.protection?.sheet === "1" || s.protection?.sheet === "true";
  const flag = (s: Sheet, attr: string) => s.protection?.[attr];
  /** Whether the protected sheet allows an action (the attribute set to 0 allows it, select* are the other way round). */
  function allows(s: Sheet, what: string): boolean {
    if (!on(s)) return true;
    if (what === "never") return false;
    const v = flag(s, what);
    const isTrue = v === "1" || v === "true";
    // These attributes say what is protected (set) rather than what is allowed.
    if (what.startsWith("select") || what === "objects" || what === "scenarios") return !isTrue;
    return v !== undefined && !isTrue;
  }
  const lockedAt = (s: Sheet, row: number, col: number) => {
    const c = s.get(row, col);
    return app.wb.styles.get(c?.s ?? s.rows.get(row)?.s ?? s.cols.get(col)?.s).locked !== false;
  };
  function anyLocked(s: Sheet, ranges: Rect[]): boolean {
    for (const r of ranges) {
      if ((r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1) > 200000) return true;
      for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) if (lockedAt(s, row, col)) return true;
    }
    return false;
  }

  function refuse(): void {
    void messageBox("RaabeCalc", "Diese Zelle ist gesperrt, und das Blatt ist geschützt. Um sie zu ändern, heben Sie zuerst den Blattschutz auf (Überprüfen > Blattschutz aufheben); dafür kann ein Kennwort nötig sein.");
  }

  /** Before a cell is typed into: false (with the message) where it is locked on a protected sheet. */
  function canEdit(row: number, col: number): boolean {
    if (!on() || !lockedAt(app.sheet, row, col)) return true;
    refuse();
    return false;
  }

  /** Whether a command may run now; "unlocked" commands refuse with the message when they run. */
  function enabled(id: string): boolean {
    const need = needs(id);
    if (!need || need === "unlocked" || !on()) return true;
    return allows(app.sheet, need);
  }
  function guard(id: string): boolean {
    if (needs(id) !== "unlocked" || !on()) return true;
    if (!anyLocked(app.sheet, app.grid.selection)) return true;
    refuse();
    return false;
  }

  /** Before cells are written: true when none of them is locked on a protected sheet, else the message. */
  function guardRects(rects: Rect[]): boolean {
    if (!on() || !anyLocked(app.sheet, rects)) return true;
    refuse();
    return false;
  }

  /** Tab on a protected sheet: the next unlocked cell in reading order (Umschalt+Tab backwards). */
  function nextUnlocked(row: number, col: number, back: boolean): { row: number; col: number } | null {
    if (!on()) return null;
    const s = app.sheet;
    const cells: { row: number; col: number }[] = [];
    // Unlocked cells carry a style that says so; stored cells are enough (Excel's templates format them).
    for (const [key] of s.cells) {
      const r = Math.floor(key / 16384), cc = key % 16384;
      if (!lockedAt(s, r, cc)) cells.push({ row: r, col: cc });
    }
    if (!cells.length) return null;
    cells.sort((a, b) => a.row - b.row || a.col - b.col);
    const after = (p: { row: number; col: number }) => p.row > row || (p.row === row && p.col > col);
    const before = (p: { row: number; col: number }) => p.row < row || (p.row === row && p.col < col);
    return back ? [...cells].reverse().find(before) ?? cells[cells.length - 1] : cells.find(after) ?? cells[0];
  }

  async function protect(): Promise<void> {
    const s = app.sheet;
    if (on(s)) return unprotect();
    const pw = h("input", { type: "password", style: "width:100%" }) as HTMLInputElement;
    const boxes = CHOICES.map(([label, attr, def]) => {
      const input = h("input", { type: "checkbox", checked: def || undefined }) as HTMLInputElement;
      return { attr, input, el: h("label", { class: "check" }, input, label) };
    });
    const body = h("div", { class: "protect-dialog" },
      h("label", { class: "check" }, h("input", { type: "checkbox", checked: true, disabled: true }), "Arbeitsblatt und Inhalt gesperrter Zellen schützen"),
      h("div", { class: "fc-label" }, "Kennwort zum Aufheben des Blattschutzes:"), pw,
      h("div", { class: "fc-label" }, "Alle Benutzer dieses Arbeitsblatts können:"),
      h("div", { class: "protect-list" }, ...boxes.map((b) => b.el)));
    setTimeout(() => pw.focus());
    const ok = await showDialog({ title: "Blatt schützen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const attrs: Record<string, string> = {};
    if (pw.value) {
      const again = h("input", { type: "password", style: "width:100%" }) as HTMLInputElement;
      setTimeout(() => again.focus());
      const ok2 = await showDialog({ title: "Kennwort bestätigen", body: h("div", null, h("div", { class: "fc-label" }, "Kennwort erneut eingeben:"), again), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
      if (!ok2) return;
      if (again.value !== pw.value) { await messageBox("RaabeCalc", "Die beiden Kennwörter sind nicht gleich."); return; }
      const salt = crypto.getRandomValues(new Uint8Array(16));
      document.body.style.cursor = "wait";
      try {
        Object.assign(attrs, { algorithmName: "SHA-512", hashValue: await sha512(pw.value, salt, 100000), saltValue: b64(salt), spinCount: "100000" });
      } finally {
        document.body.style.cursor = "";
      }
    }
    const allowed = new Map(boxes.map((b) => [b.attr, b.input.checked]));
    app.change("Blatt schützen", () => { app.history.whole(s); s.protection = protectionAttrs(allowed, attrs); });
  }

  async function unprotect(): Promise<void> {
    const s = app.sheet;
    const p = s.protection ?? {};
    if (p.hashValue || p.password) {
      const pw = h("input", { type: "password", style: "width:100%" }) as HTMLInputElement;
      setTimeout(() => pw.focus());
      const ok = await showDialog({ title: "Blattschutz aufheben", body: h("div", null, h("div", { class: "fc-label" }, "Kennwort:"), pw), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
      if (!ok) return;
      let right = false;
      if (p.hashValue && /sha-?512/i.test(p.algorithmName ?? "")) right = (await sha512(pw.value, unb64(p.saltValue ?? ""), Number(p.spinCount ?? 0))) === p.hashValue;
      else if (p.password) right = legacyHash(pw.value) === p.password.toUpperCase();
      if (!right) { await messageBox("RaabeCalc", "Das Kennwort ist nicht richtig. Achten Sie auf Groß- und Kleinschreibung."); return; }
    }
    app.change("Blattschutz aufheben", () => { app.history.whole(s); s.protection = undefined; });
  }

  /** Arbeitsmappe schützen (Struktur): no sheets added, deleted, renamed, moved, hidden or coloured. */
  const bookLocked = () => { const v = app.wb.protection?.lockStructure; return v === "1" || v === "true"; };
  const STRUCTURE = /^(cells\.(insertSheet|deleteSheet|renameSheet|moveSheet|tabColor|hideSheet|unhideSheet)|sheet\.add)$/;

  async function askPassword(title: string, label: string): Promise<string | null> {
    const pw = h("input", { type: "password", style: "width:100%" }) as HTMLInputElement;
    setTimeout(() => pw.focus());
    const ok = await showDialog({ title, body: h("div", null, h("div", { class: "fc-label" }, label), pw), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    return ok ? pw.value : null;
  }

  async function protectBook(): Promise<void> {
    const wb = app.wb;
    if (bookLocked()) {
      const p = wb.protection ?? {};
      if (p.workbookHashValue || p.workbookPassword) {
        const pw = await askPassword("Arbeitsmappenschutz aufheben", "Kennwort:");
        if (pw === null) return;
        const right = p.workbookHashValue ? (await sha512(pw, unb64(p.workbookSaltValue ?? ""), Number(p.workbookSpinCount ?? 0))) === p.workbookHashValue : legacyHash(pw) === (p.workbookPassword ?? "").toUpperCase();
        if (!right) { await messageBox("RaabeCalc", "Das Kennwort ist nicht richtig. Achten Sie auf Groß- und Kleinschreibung."); return; }
      }
      app.change("Arbeitsmappenschutz aufheben", () => { app.history.sheets(); wb.protection = undefined; });
      return;
    }
    const pw = h("input", { type: "password", style: "width:100%" }) as HTMLInputElement;
    const structure = h("input", { type: "checkbox", checked: true }) as HTMLInputElement;
    const body = h("div", { class: "protect-dialog" }, h("div", { class: "fc-label" }, "Arbeitsmappe schützen für"),
      h("label", { class: "check" }, structure, "Struktur"), h("label", { class: "check" }, h("input", { type: "checkbox", disabled: true }), "Fenster"),
      h("div", { class: "fc-label" }, "Kennwort (optional):"), pw);
    setTimeout(() => pw.focus());
    const ok = await showDialog({ title: "Struktur und Fenster schützen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok || !structure.checked) return;
    const attrs: Record<string, string> = {};
    if (pw.value) {
      const again = await askPassword("Kennwort bestätigen", "Kennwort erneut eingeben:");
      if (again === null) return;
      if (again !== pw.value) { await messageBox("RaabeCalc", "Die beiden Kennwörter sind nicht gleich."); return; }
      const salt = crypto.getRandomValues(new Uint8Array(16));
      Object.assign(attrs, { workbookAlgorithmName: "SHA-512", workbookHashValue: await sha512(pw.value, salt, 100000), workbookSaltValue: b64(salt), workbookSpinCount: "100000" });
    }
    attrs.lockStructure = "1";
    app.change("Arbeitsmappe schützen", () => { app.history.sheets(); wb.protection = attrs; });
  }

  return { on, allows, canEdit, enabled, guard, guardRects, nextUnlocked, protect, unprotect, bookLocked, structureCommand: (id: string) => STRUCTURE.test(id), protectBook };
}

export { legacyHash, sha512 };
