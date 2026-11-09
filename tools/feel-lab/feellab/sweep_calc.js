// Page side of sweep_calc.py (D-141): finds RaabeCalc's controls as a user sees them (visible, not
// covered), takes snapshots of the workbook and the UI so an option's effect and its undo can be
// judged. Read-only: every change to the document goes through trusted CDP input.
(() => {
  const vis = (el) => {
    if (!el || !el.isConnected) return false;
    const r = el.getBoundingClientRect();
    if (r.width < 2 || r.height < 2) return false;
    const cs = getComputedStyle(el);
    return cs.visibility !== "hidden" && cs.display !== "none" && r.right > 0 && r.bottom > 0 && r.left < innerWidth && r.top < innerHeight;
  };
  const center = (el) => { const r = el.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; };
  /** True when a click at the element's centre reaches it (not covered by something else). */
  const reachable = (el) => { const c = center(el); const at = document.elementFromPoint(c.x, c.y); return !!at && (at === el || el.contains(at)); };
  const text = (el) => (el.getAttribute("aria-label") || el.textContent || el.getAttribute("title") || "").replace(/\s+/g, " ").trim();
  const disabled = (el) => !!(el.disabled || el.classList.contains("disabled") || el.getAttribute("aria-disabled") === "true");
  const desc = (el, extra) => ({ label: text(el), ...center(el), disabled: disabled(el), reachable: reachable(el), ...extra });

  const KNOWN = new WeakSet();
  const S = {
    /** Elements on the page now: what is added after a click is a popup, a dialog or a pane. */
    mark() { for (const el of document.body.children) KNOWN.add(el); },
    added() {
      return [...document.body.children].filter((el) => !KNOWN.has(el) && !el.classList.contains("tooltip") && vis(el) || (!KNOWN.has(el) && el.classList.contains("dialog-layer")));
    },
    tabs() {
      return [...document.querySelectorAll(".tabrow .tab")].filter(vis).map((t) => desc(t, { key: t.dataset.keytip, contextual: t.classList.contains("contextual"), active: t.classList.contains("active") }));
    },
    /** The open tab's controls: the clickable parts, each with an id that stays the same between runs. */
    controls() {
      const card = document.querySelector(".ribbon-card");
      const out = [];
      const count = new Map();
      for (const g of card.querySelectorAll(".rgroup")) {
        const gname = g.getAttribute("aria-label") || g.dataset.group || "?";
        const els = g.querySelectorAll("button.rb, .rcombo, .rcheck, .rgallery .item:not(.extra), .rgallery .scroll button[aria-label=Weitere], .rspinner button, .rspinner input, button.launcher, .rgroup.collapsed > .content > button, .rcombo-face");
        for (const el of els) {
          if (!vis(el)) continue;
          if (el.closest(".rcombo") && el !== el.closest(".rcombo")) continue;
          let kind = el.classList.contains("rcombo") ? "combo" : el.classList.contains("rcheck") ? "check" : el.classList.contains("launcher") ? "launcher"
            : el.closest(".rgallery") ? (el.classList.contains("item") ? "gallery-item" : "gallery-more") : el.closest(".rspinner") ? "spinner"
            : el.classList.contains("arrow") ? "split-arrow" : el.closest(".rsplit") ? "split-main" : el.classList.contains("has-chev") ? "dropdown" : "button";
          let label = text(el);
          if (kind === "split-arrow") label = text(el.parentElement.querySelector("button:not(.arrow)")) + " ▾";
          if (kind === "combo") {
            label = el.querySelector("input")?.getAttribute("aria-label") ?? label;
            // A combo's list opens from its arrow: that is where a user clicks.
            const arrow = el.querySelector("button[aria-label='Öffnen']");
            if (arrow && vis(arrow)) { const base = `${gname}|${label}`; out.push(desc(arrow, { id: base, group: gname, kind, label, keytip: el.dataset.keytip })); continue; }
          }
          if (kind === "spinner") label = text(el.closest(".rspinner").querySelector(".lbl")) + " " + (el.tagName === "INPUT" ? "Feld" : text(el));
          const base = `${gname}|${label}`;
          const n = (count.get(base) ?? 0) + 1;
          count.set(base, n);
          out.push(desc(el, { id: n > 1 ? `${base}#${n}` : base, group: gname, kind, label, keytip: el.dataset.keytip ?? el.closest("[data-keytip]")?.dataset.keytip }));
        }
      }
      return out;
    },
    /** What a click opened: a popup's entries, a dialog, a task pane. */
    opened() {
      const added = S.added();
      const dlg = [...document.querySelectorAll(".dialog-layer .dialog")].pop();
      if (dlg) return { kind: "dialog", title: dlg.getAttribute("aria-label"), controls: S.dialogControls(dlg) };
      const pop = added.filter((el) => !el.classList.contains("dialog-layer")).pop();
      if (pop) return { kind: "popup", cls: pop.className, items: S.popupItems(pop) };
      return null;
    },
    popupItems(pop) {
      const out = [];
      const count = new Map();
      // The deepest open submenu last: its entries are the ones to pick next.
      const els = pop.querySelectorAll("button, input, select, [role=menuitem], [role=option], .item, .swatch, [data-cmd]");
      for (const el of els) {
        if (!vis(el)) continue;
        if (el.tagName !== "BUTTON" && el.closest("button")) continue;
        let label = text(el) || el.getAttribute("data-tip")?.split("\n")[0] || el.style.background || el.className;
        const menu = el.closest(".menu");
        const level = menu?.classList.contains("submenu") ? 1 : 0;
        const key = (level ? "› " : "") + label;
        const n = (count.get(key) ?? 0) + 1;
        count.set(key, n);
        out.push(desc(el, { id: n > 1 ? `${key}#${n}` : key, tag: el.tagName.toLowerCase(), level, sub: el.classList.contains("has-sub"), type: el.type || "" }));
      }
      return out;
    },
    dialogControls(dlg) {
      return [...dlg.querySelectorAll("button, input, select, textarea")].filter(vis).map((el) => {
        const lab = el.closest("label")?.textContent?.trim() || el.closest(".field, .dlg-field")?.firstChild?.textContent?.trim() || "";
        return { tag: el.tagName.toLowerCase(), type: el.type || "", label: (text(el) || lab).slice(0, 80), field: lab.slice(0, 80), value: el.type === "checkbox" || el.type === "radio" ? el.checked : (el.value ?? "").slice(0, 60), ...center(el), disabled: !!el.disabled };
      });
    },
    dialogTitle() { const d = [...document.querySelectorAll(".dialog-layer .dialog")].pop(); return d ? d.getAttribute("aria-label") : null; },
    find(kind, id) {
      if (kind === "tab") return S.tabs().find((t) => t.label === id) ?? null;
      if (kind === "ctrl") return S.controls().find((c) => c.id === id) ?? null;
      if (kind === "item") {
        const pops = [...document.body.children].filter((el) => !KNOWN.has(el) && vis(el) && !el.classList.contains("tooltip") && !el.classList.contains("dialog-layer"));
        for (const p of pops.reverse()) {
          let hit = S.popupItems(p).find((i) => i.id === id);
          // A long list (fonts): a user scrolls the entry into view first.
          if (hit && !hit.reachable) {
            const el = [...p.querySelectorAll("button, [role=menuitem], [role=option], .item")].find((e) => text(e) === hit.label);
            el?.scrollIntoView({ block: "nearest" });
            hit = S.popupItems(p).find((i) => i.id === id) ?? hit;
          }
          if (hit) return hit;
        }
        return null;
      }
      return null;
    },
    /** The workbook as data (cells with their resolved style, merges, rows, columns, rules, tables, names, page). */
    model() {
      const app = __calc.app, wb = app.wb;
      const rep = (k, v) => (v instanceof Map ? [...v.entries()] : v instanceof Set ? [...v] : (k === "raw" || k === "calc" || k === "xml" || k === "image" || k === "bytes" || k === "data") && typeof v === "object" ? "#" : v);
      const sheets = wb.sheets.map((s) => {
        const cells = [];
        for (const [k, c] of s.cells) { const { s: sid, ...rest } = c; cells.push([k, JSON.stringify(rest, rep), JSON.stringify(wb.styles.get(sid))]); }
        cells.sort((a, b) => a[0] - b[0]);
        const o = {};
        // An empty list is the same as none (charts: [] after the last chart is undone).
        for (const [k, v] of Object.entries(s)) if (!["cells", "view", "views", "raw"].includes(k) && v !== undefined && !(Array.isArray(v) && !v.length)) o[k] = v;
        let rest;
        try { rest = JSON.stringify(o, rep); } catch (e) { rest = "unserialisable:" + e; }
        return JSON.stringify(cells) + rest;
      });
      let names;
      try { names = JSON.stringify(wb.names, rep); } catch { names = "?"; }
      return JSON.stringify({ sheets, names, active: wb.active, n: wb.sheets.length });
    },
    /** What a user sees besides the cells: selection, view settings, panes, mode, status bar. */
    ui() {
      const app = __calc.app;
      const v = app.sheet.view;
      const panes = [...document.querySelectorAll(".taskpane")].filter((p) => !p.hidden && vis(p)).map((p) => p.querySelector(".tp-title")?.textContent ?? "pane");
      return JSON.stringify({
        sel: app.grid.selection, active: v.active, zoom: v.zoom, grid: v.showGrid, heads: v.showHeaders, freeze: v.freeze ?? null, split: v.split ?? null, viewMode: v.mode ?? v.viewMode ?? null,
        panes, backstage: !document.querySelector(".backstage")?.hidden, body: document.body.className, title: document.title, typing: (() => { const a = document.activeElement; return a && (a.tagName === "TEXTAREA" || a.tagName === "INPUT") ? a.className : null; })(), cursor: (() => { const c = document.querySelector(".grid canvas, canvas"); return c ? getComputedStyle(c).cursor : null; })(), tab: document.querySelector(".tabrow .tab.active")?.textContent,
        ctx: [...document.querySelectorAll(".tabrow .tab.contextual")].map((t) => t.textContent), editing: app.grid.isEditing,
        fx: !!document.querySelector(".fx-input") && vis(document.querySelector(".fx-input")),
        status: document.querySelector(".statusbar")?.textContent, sheetbar: document.querySelector(".sheetbar")?.textContent,
      });
    },
    canvas() {
      // A cheap fingerprint of what the grid draws: the canvases scaled down, hashed.
      const small = S._small ??= Object.assign(document.createElement("canvas"), { width: 240, height: 240 });
      const g = small.getContext("2d", { willReadFrequently: true });
      let h = 0;
      for (const cv of document.querySelectorAll(".grid canvas, canvas")) {
        if (cv === small || !cv.width || !cv.height) continue;
        g.clearRect(0, 0, 240, 240);
        try { g.drawImage(cv, 0, 0, 240, 240); } catch { continue; }
        const d = g.getImageData(0, 0, 240, 240).data;
        for (let i = 0; i < d.length; i += 4) h = (h * 31 + d[i] + 7 * d[i + 1] + 13 * d[i + 2]) | 0;
      }
      return String(h);
    },
    snap() { return { model: S.model(), ui: S.ui(), canvas: S.canvas() }; },
    history() { const hst = __calc.app.history; return { undo: hst.canUndo, label: hst.undoLabel ?? null }; },
  };
  window.__sweep = S;
  return true;
})();
