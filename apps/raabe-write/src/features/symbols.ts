// Einfügen > Symbol as in Word: the menu of the 20 symbols used last ("Weitere Symbole...")
// and the Symbol dialog: Symbole (Schriftart, Subset, the grid of the font's characters,
// Zuletzt verwendete Symbole, Unicodename, Zeichencode and von:) and Sonderzeichen (Word's list
// with names and shortcuts). Recorded from Word with probe_word_dialogs.py and
// probe_word_symbol_dialog.py. The dialog stays open after Einfügen (Abbrechen becomes
// Schließen), as in Word.

import { h, icon } from "@raabe/shell";
import type { AppApi } from "./api";

/** Word's subsets (Unicode blocks) with the German names Word shows; the first 47 recorded from Word. */
const BLOCKS: [number, number, string][] = [
  [0x0020, 0x007f, "Standardlateinisch"], [0x0080, 0x00ff, "Lateinisch-1 (Ergänzung)"], [0x0100, 0x017f, "Erweitertes Lateinisch-A"],
  [0x0180, 0x024f, "Erweitertes Lateinisch-B"], [0x0250, 0x02af, "IPA-Erweiterungen"], [0x02b0, 0x02ff, "Buchstaben zur Abstanddefinition"],
  [0x0300, 0x036f, "Kombinierende diakritische Zeichen"], [0x0370, 0x03ff, "Griechisch und Koptisch"], [0x0400, 0x04ff, "Kyrillisch"],
  [0x0500, 0x052f, "Kyrillisch (Ergänzung)"], [0x0530, 0x058f, "Armenisch"], [0x0590, 0x05ff, "Hebräisch"], [0x0600, 0x06ff, "Arabisch"],
  [0x0750, 0x077f, "Arabisch (Ergänzung)"], [0x07c0, 0x07ff, "N'Ko"], [0x0900, 0x097f, "Devanagari"], [0x0e00, 0x0e7f, "Thailändisch"],
  [0x10a0, 0x10ff, "Georgisch"], [0x1a00, 0x1a1f, "Buginesisch"], [0x1d00, 0x1d7f, "Phonetische Erweiterungen"],
  [0x1d80, 0x1dbf, "Zusätzliche phonetische Erweiterungen"], [0x1dc0, 0x1dff, "Zusätzliche kombinierende diakritische Zeichen"],
  [0x1e00, 0x1eff, "Zusätzliches erweitertes Lateinisch"], [0x1f00, 0x1fff, "Griechisch (erweitert)"], [0x2000, 0x206f, "Allgemeine Interpunktion"],
  [0x2070, 0x209f, "Hoch- und tiefgestellte Zeichen"], [0x20a0, 0x20cf, "Währungssymbole"], [0x20d0, 0x20ff, "Kombinierende diakritische Zeichen für Symbole"],
  [0x2100, 0x214f, "Buchstabenähnliche Symbole"], [0x2150, 0x218f, "Zahlenzeichen"], [0x2190, 0x21ff, "Pfeile"], [0x2200, 0x22ff, "Mathematische Operatoren"],
  [0x2300, 0x23ff, "Verschiedene technische Sonderzeichen"], [0x2400, 0x243f, "Steuerzeichen-Bilder"], [0x2460, 0x24ff, "Eingeschlossene alphanumerische Zeichen"],
  [0x2500, 0x257f, "Linienzeichen"], [0x2580, 0x259f, "Blockelemente"], [0x25a0, 0x25ff, "Geometrische Formen"], [0x2600, 0x26ff, "Graphische Symbole"],
  [0x2700, 0x27bf, "Dingbats"], [0x27c0, 0x27ef, "Verschiedene mathematische Symbole-A"], [0x27f0, 0x27ff, "Zusätzliche Pfeile-A"],
  [0x2800, 0x28ff, "Braillemuster"], [0x2900, 0x297f, "Zusätzliche Pfeile-B"], [0x2980, 0x29ff, "Verschiedene mathematische Symbole-B"],
  [0x2a00, 0x2aff, "Zusätzliche mathematische Operatoren"], [0x2b00, 0x2bff, "Verschiedene Symbole und Pfeile"], [0x2c00, 0x2c5f, "Glagolitisch"],
  [0x2c60, 0x2c7f, "Erweitertes Lateinisch-C"], [0x2d00, 0x2d2f, "Georgisch (Ergänzung)"], [0x2e00, 0x2e7f, "Zusätzliche Interpunktion"],
  [0xa4d0, 0xa4ff, "Lisu"], [0xa640, 0xa69f, "Erweitertes Kyrillisch-B"], [0xa700, 0xa71f, "Modifizierende Tonzeichen"], [0xa720, 0xa7ff, "Erweitertes Lateinisch-D"],
  [0xa840, 0xa87f, "Phags-pa"], [0xfb00, 0xfb4f, "Alphabetische Präsentationszeichen"], [0xfb50, 0xfdff, "Arabische Präsentationsformen-A"],
  [0xfe00, 0xfe0f, "Variantenselektoren"], [0xfe20, 0xfe2f, "Kombinierende halbe diakritische Zeichen"], [0xfe70, 0xfeff, "Arabische Präsentationsformen-B"],
  [0xff00, 0xffef, "Halbbreite und vollbreite Formen"], [0xfff0, 0xffff, "Spezielle Zeichen"], [0x1d400, 0x1d7ff, "Mathematische alphanumerische Symbole"],
  [0x1f300, 0x1f5ff, "Verschiedene Piktogramme"], [0x1f600, 0x1f64f, "Emoticons"], [0x1f680, 0x1f6ff, "Verkehrs- und Kartensymbole"],
  [0x1f900, 0x1f9ff, "Zusätzliche Symbole und Piktogramme"],
];

/** Sonderzeichen: Word's list (character, name, shortcut), recorded from Word. */
const SPECIAL: [number, string, string][] = [
  [0x2014, "Geviertstrich", "Alt+Strg+Num -"], [0x2013, "Halbgeviertstrich", "Strg+Num -"], [0x2011, "Geschützter Trennstrich", "Alt+Umschalt+_"],
  [0x00ad, "Bedingter Trennstrich", "Alt+Strg+-"], [0x2003, "Em-Abstand", ""], [0x2002, "En-Abstand", ""], [0x2005, "1/4-Em-Abstand", ""],
  [0x00a0, "Geschütztes Leerzeichen", "Strg+Umschalt+Leertaste"], [0x00a9, "Copyright", ""], [0x00ae, "Eingetragene", "Alt+Strg+R"],
  [0x2122, "Marke", "Alt+Strg+T"], [0x00a7, "Paragrafenzeichen", ""], [0x00b6, "Absatz", ""], [0x2026, "Auslassungspunkte", "Alt+Strg+."],
  [0x2018, "Einfaches öffnendes Anführungszeichen", "Strg+','"], [0x2019, "Einfaches schließendes Anführungszeichen", "Strg+`,`"],
  [0x201c, "Öffnendes Anführungszeichen", "Strg+',\""], [0x201d, "Schließendes Anführungszeichen", "Strg+`,\""],
  [0x200c, "Bedingter Nullbreite-Wechsel", ""], [0x200d, "Null-breiter Nicht-Wechsel", ""], [0x200e, "Links-nach-Rechts-Markierung", ""],
  [0x200f, "Rechts-nach-Links-Markierung", ""], [0x202a, "Links-nach-Rechts-Einbettung", ""], [0x202b, "Rechts-nach-Links-Einbettung", ""],
  [0x202d, "Links-nach-Rechts erzwingen", ""], [0x202e, "Rechts-nach-Links erzwingen", ""], [0x202c, "Ursprüngliche Schreibrichtung wiederherstellen", ""],
];

/** What Word offers before anything was inserted. */
const DEFAULT_RECENT = ["€", "£", "¥", "©", "®", "™", "±", "≠", "≤", "≥", "÷", "×", "∞", "µ", "α", "β", "π", "Ω", "∑", "☺"];
const NORMAL = "(normaler Text)";
const COLS = 16, ROW = 31, VISIBLE_ROWS = 7;

export interface Recent { c: string; f?: string; s?: boolean }

function recentList(): Recent[] {
  try {
    const v = JSON.parse(localStorage.getItem("raabe.symbols.recent") ?? "null") as Recent[] | null;
    if (Array.isArray(v) && v.length) return v.filter((x) => typeof x?.c === "string");
  } catch { /* private mode */ }
  return DEFAULT_RECENT.map((c) => ({ c }));
}

function remember(r: Recent): void {
  const list = [r, ...recentList().filter((x) => !(x.c === r.c && x.f === r.f))].slice(0, 30);
  try { localStorage.setItem("raabe.symbols.recent", JSON.stringify(list)); } catch { /* private mode */ }
}

// ----- Unicode names (English, as Word shows them), loaded with the dialog -----------------------

let names: Map<number, string> | null = null;
let loadingNames: Promise<Map<number, string>> | null = null;

function loadNames(): Promise<Map<number, string>> {
  loadingNames ??= import("./unicode-names.gen").then(({ NAMES }) => {
    const map = new Map<number, string>();
    // Word writes the names in title case: "Less-Than Or Equal To".
    const words = NAMES.w.map((w) => w.split("-").map((p) => (p ? p[0] + p.slice(1).toLowerCase() : p)).join("-"));
    for (const [start, list] of NAMES.r as [number, number[][]][]) list.forEach((enc, k) => map.set(start + k, enc.map((i) => words[i]).join(" ")));
    names = map;
    return map;
  });
  return loadingNames;
}

const nameOf = (cp: number) => names?.get(cp) ?? "";
const chr = (cp: number) => String.fromCodePoint(cp);

export function symbolsFeature(api: AppApi) {
  const c = api.commands;
  const ed = () => api.editor();

  const insert = (r: Recent) => {
    const e = ed();
    if (!e) return;
    e.insertSymbol(r.c, r.f, !!r.s);
    remember(r);
    api.refresh();
  };

  // ----- the menu ----------------------------------------------------------------------------
  const menuPanel = (close: () => void) => {
    void loadNames().then(() => { for (const b of Array.from(grid.children) as HTMLElement[]) b.dataset.tip = b.dataset.name ? nameOf(Number(b.dataset.name)) : b.dataset.tip ?? ""; });
    const grid = h("div", { class: "symgrid" }, ...recentList().slice(0, 20).map((r) => {
      const cp = r.c.codePointAt(0) ?? 0;
      return h("button", {
        class: "symcell", style: r.f ? `font-family:"${r.f}"` : "", "data-name": r.s ? "" : String(cp), "data-tip": r.s ? `Symbol: ${cp - 0xf000}` : nameOf(cp),
        onclick: () => { close(); insert(r); api.focus(); },
      }, r.c);
    }));
    return h("div", { class: "rpanel sympanel" }, grid, h("hr"),
      h("button", { class: "menu-like", onclick: () => { close(); void dialog(); } }, icon("omega"), h("span", {}, "Weitere Symbole...")));
  };

  // ----- the dialog ----------------------------------------------------------------------------
  /**
   * The dialog Symbol. With `choose` it picks a character instead of inserting it (Neues
   * Aufzählungszeichen definieren > Symbol...): OK hands it over, Sonderzeichen is not offered.
   */
  const dialog = async (choose?: { start: Recent; pick: (r: Recent) => void }) => {
    const e = ed();
    if (!e) return;
    await loadNames();
    const fonts = [NORMAL, ...(await api.platform.listFonts().catch(() => [] as string[]))];
    let font = choose?.start.f ?? NORMAL;
    let symbolFont = false;
    let chars: number[] = [];
    let selected = 0;
    let closeLabel = "Abbrechen";

    const fontInput = h("select", { class: "symfont" }, ...fonts.map((f) => h("option", { value: f }, f))) as HTMLSelectElement;
    const subset = h("select", { class: "symsubset" }) as HTMLSelectElement;
    const subsetRow = h("label", { class: "field row" }, "Subset:", subset);
    const scroller = h("div", { class: "symscroll", tabIndex: 0 });
    const inner = h("div", { class: "syminner" });
    scroller.append(inner);
    const recentRow = h("div", { class: "symrecent" });
    const uname = h("div", { class: "symname" });
    const code = h("input", { type: "text", class: "symcode" }) as HTMLInputElement;
    const von = h("select", { class: "symvon" }) as HTMLSelectElement;
    const insertBtn = h("button", { class: "primary" }, choose ? "OK" : "Einfügen") as HTMLButtonElement;
    const cancelBtn = h("button", {}, closeLabel) as HTMLButtonElement;

    // "(normaler Text)" is the font at the caret.
    const caretFont = e.charState().font || "Calibri";
    const baseFont = () => (font === NORMAL ? caretFont : font);
    const fontCss = () => `font-family:"${font === NORMAL ? baseFont() : font}"`;

    const loadChars = async () => {
      const info = await api.platform.fontChars?.(font === NORMAL ? baseFont() : font).catch(() => null);
      symbolFont = !!info?.symbol;
      const ranges: [number, number][] = info?.found && info.ranges.length ? info.ranges
        : [[0x20, 0x7e], [0xa0, 0x17f], [0x370, 0x3ff], [0x2000, 0x206f], [0x20a0, 0x20cf], [0x2100, 0x214f], [0x2190, 0x21ff], [0x2200, 0x22ff], [0x25a0, 0x25ff], [0x2600, 0x26ff]];
      const list: number[] = [];
      for (const [a, b] of ranges) for (let cp = Math.max(a, 0x20); cp <= b; cp++) {
        if (symbolFont ? cp >= 0xf020 && cp <= 0xf0ff : (cp < 0xd800 || cp > 0xf8ff) && (cp < 0x7f || cp > 0x9f)) list.push(cp);
      }
      chars = list;
      von.replaceChildren(...(symbolFont ? ["Symbol (dezimal)"] : ["Unicode (hex)", "ASCII (dezimal)", "ASCII (hex)"]).map((v) => h("option", { value: v }, v)));
      // Subsets only for Unicode fonts: the blocks the font has.
      const blocks = symbolFont ? [] : BLOCKS.filter(([a, b]) => chars.some((cp) => cp >= a && cp <= b));
      subset.replaceChildren(...blocks.map(([a, , n]) => h("option", { value: String(a) }, n)));
      subsetRow.hidden = symbolFont;
      selected = Math.min(selected, chars.length - 1);
      inner.style.height = `${Math.ceil(chars.length / COLS) * ROW}px`;
      scroller.scrollTop = 0;
      renderGrid();
      show();
    };

    const visibleChars = () => (von.value.startsWith("ASCII") ? chars.filter((cp) => cp < 0x100) : chars);

    const renderGrid = () => {
      const list = visibleChars();
      inner.style.height = `${Math.ceil(list.length / COLS) * ROW}px`;
      const first = Math.floor(scroller.scrollTop / ROW);
      const cells: HTMLElement[] = [];
      for (let r = first; r < first + VISIBLE_ROWS + 1; r++) {
        for (let k = 0; k < COLS; k++) {
          const i = r * COLS + k;
          if (i >= list.length) break;
          const cp = list[i];
          cells.push(h("div", {
            class: `symcellg${i === selected ? " on" : ""}`, style: `${fontCss()};left:${k * (100 / COLS)}%;top:${r * ROW}px`,
            onmousedown: () => { selected = i; renderGrid(); show(); },
            ondblclick: () => { selected = i; doInsert(); },
          }, chr(cp)));
        }
      }
      // Empty cells continue the grid below the last character, as in Word.
      inner.replaceChildren(...cells);
      // The subset follows the scroll position.
      const top = list[first * COLS];
      if (top !== undefined && !symbolFont) {
        const b = BLOCKS.find(([a, z]) => top >= a && top <= z);
        if (b) subset.value = String(b[0]);
      }
    };

    const show = () => {
      const cp = visibleChars()[selected];
      if (cp === undefined) { uname.textContent = ""; return; }
      if (symbolFont) { uname.textContent = `Symbol: ${cp - 0xf000}`; code.value = String(cp - 0xf000); return; }
      uname.textContent = nameOf(cp);
      code.value = von.value === "Unicode (hex)" ? cp.toString(16).toUpperCase().padStart(4, "0") : von.value === "ASCII (hex)" ? cp.toString(16).toUpperCase() : String(cp);
    };

    const renderRecent = () => {
      recentRow.replaceChildren(...recentList().slice(0, 27).map((r) => h("div", {
        class: "symrecentcell", style: r.f ? `font-family:"${r.f}"` : "", "data-tip": r.s ? `Symbol: ${(r.c.codePointAt(0) ?? 0) - 0xf000}` : nameOf(r.c.codePointAt(0) ?? 0),
        onmousedown: () => {
          if (r.f && r.f !== font) { fontInput.value = r.f; font = r.f; void loadChars().then(() => pick(r.c.codePointAt(0) ?? 0)); }
          else pick(r.c.codePointAt(0) ?? 0);
        },
        ondblclick: () => { insert(r); renderRecent(); },
      }, r.c)));
    };

    const pick = (cp: number) => {
      const i = visibleChars().indexOf(cp);
      if (i < 0) return;
      selected = i;
      scroller.scrollTop = Math.max(0, Math.floor(i / COLS) * ROW - ROW * 2);
      renderGrid();
      show();
    };

    const doInsert = () => {
      const cp = visibleChars()[selected];
      if (cp === undefined) return;
      if (choose) { choose.pick({ c: chr(cp), f: font === NORMAL ? undefined : font, s: symbolFont || undefined }); done(); return; }
      insert({ c: chr(cp), f: font === NORMAL ? undefined : font, s: symbolFont || undefined });
      renderRecent();
      closeLabel = "Schließen";
      cancelBtn.textContent = closeLabel;
    };

    fontInput.onchange = () => { font = fontInput.value; selected = 0; void loadChars(); };
    subset.onchange = () => pick(visibleChars().find((cp) => cp >= Number(subset.value)) ?? 0);
    von.onchange = () => { selected = 0; scroller.scrollTop = 0; renderGrid(); show(); };
    scroller.addEventListener("scroll", () => renderGrid());
    scroller.addEventListener("keydown", (ev) => {
      const list = visibleChars();
      const step = ev.key === "ArrowRight" ? 1 : ev.key === "ArrowLeft" ? -1 : ev.key === "ArrowDown" ? COLS : ev.key === "ArrowUp" ? -COLS : 0;
      if (step) { ev.preventDefault(); pick(list[Math.max(0, Math.min(list.length - 1, selected + step))]); }
      else if (ev.key === "Enter") { ev.preventDefault(); ev.stopPropagation(); doInsert(); }
    });
    code.addEventListener("input", () => {
      const v = code.value.trim();
      const cp = symbolFont ? 0xf000 + Number(v) : von.value === "ASCII (dezimal)" ? Number(v) : parseInt(v, 16);
      if (Number.isFinite(cp)) pick(cp);
    });

    // Sonderzeichen.
    let special = 0;
    const specialList = h("div", { class: "symspecial", tabIndex: 0 });
    const renderSpecial = () => specialList.replaceChildren(
      h("div", { class: "symspecialhead" }, h("span", {}, "Zeichen:"), h("span", {}, ""), h("span", {}, "Tastenkombination:")),
      ...SPECIAL.map(([cp, name, keys], i) => h("div", {
        class: `symspecialrow${i === special ? " on" : ""}`, onmousedown: () => { special = i; renderSpecial(); },
        ondblclick: () => { special = i; insertSpecial(); },
      }, h("span", { class: "ch" }, cp === 0x00a0 || (cp >= 0x2000 && cp <= 0x202e) ? "" : chr(cp)), h("span", {}, name), h("span", {}, keys))));
    const insertSpecial = () => {
      const [cp] = SPECIAL[special];
      e.insertSymbol(chr(cp));
      api.refresh();
      closeLabel = "Schließen";
      cancelBtn.textContent = closeLabel;
    };

    const pageSymbols = h("div", { class: "sympage" },
      h("div", { class: "row" }, h("label", { class: "field row" }, "Schriftart:", fontInput), subsetRow),
      h("div", { class: "symgridwrap" }, scroller),
      h("div", { class: "symlabel" }, "Zuletzt verwendete Symbole:"), recentRow,
      h("div", { class: "symbottom" },
        h("div", {}, h("div", {}, "Unicodename:"), uname),
        h("label", { class: "field row" }, "Zeichencode:", code, "von:", von)),
      h("div", { class: "row" },
        h("button", { disabled: true, "data-tip": "AutoKorrektur...\nIn RaabeWrite gelten Words AutoKorrektur-Einträge; eigene lassen sich nicht hinzufügen." }, "AutoKorrektur..."),
        h("button", { disabled: true, "data-tip": "Tastenkombination...\nRaabeWrite verwendet Words Tastenkombinationen; eigene lassen sich nicht festlegen." }, "Tastenkombination..."),
        h("span", {}, "Tastenkombination:")));
    const pageSpecial = h("div", { class: "sympage", hidden: true }, specialList);
    const tabs = h("div", { class: "tabs" });
    const tab = (label: string, page: HTMLElement) => h("button", {
      class: page === pageSymbols ? "active" : "",
      onclick: (ev: MouseEvent) => {
        pageSymbols.hidden = page !== pageSymbols; pageSpecial.hidden = page !== pageSpecial;
        tabs.querySelectorAll("button").forEach((b) => b.classList.toggle("active", b === ev.currentTarget));
        if (page === pageSpecial) renderSpecial();
      },
    }, label);
    tabs.append(tab("Symbole", pageSymbols), ...(choose ? [] : [tab("Sonderzeichen", pageSpecial)]));

    // The dialog itself: modeless in Word; here a dialog that stays until Schließen.
    const box = h("div", { class: "dialog symdialog", role: "dialog", "aria-label": "Symbol" },
      h("div", { class: "dlg-title" }, h("span", { class: "t" }, "Symbol"), h("button", { "aria-label": "Schließen", onclick: () => done() }, icon("close"))),
      h("div", { class: "dlg-body" }, tabs, pageSymbols, pageSpecial),
      h("div", { class: "dlg-buttons" }, insertBtn, cancelBtn));
    const layer = h("div", { class: "dialog-layer" }, box);
    const done = () => { layer.remove(); removeEventListener("keydown", key, true); api.focus(); };
    const key = (ev: KeyboardEvent) => { if (ev.key === "Escape") { ev.preventDefault(); ev.stopPropagation(); done(); } };
    insertBtn.onclick = () => (pageSpecial.hidden ? doInsert() : insertSpecial());
    cancelBtn.onclick = done;
    addEventListener("keydown", key, true);
    document.body.append(layer);
    renderRecent();
    fontInput.value = font;
    await loadChars();
    // Word opens on the symbol used last (when choosing a bullet: on the current one).
    const last = choose?.start ?? recentList()[0];
    if (last && (choose || !last.f)) pick(last.c.codePointAt(0) ?? 0);
    scroller.focus();
  };

  c.implement("insert.symbol", { run: () => void dialog(), enabled: () => !!ed() });
  return { menuPanel, dialog };
}
