"""Page-side helpers of the by-hand audit (sweep_write.py, D-133), injected into RaabeWrite's page: fingerprints of
the document and the UI, and the visible controls of the ribbon, open menus, galleries and dialogs with their
positions, so the sweep can click them like a student does."""

HELPERS = r"""
(() => {
  const H = (s) => { let h = 2166136261 >>> 0; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; } return h.toString(16) + "/" + s.length; };
  const repl = (k, v) => /Cache$|^_/.test(k) ? undefined : v instanceof Map ? [...v.entries()] : v instanceof Set ? [...v] : (typeof v === "function" ? undefined : v);
  const vis = (el) => {
    if (!el || !el.isConnected) return false;
    const r = el.getBoundingClientRect();
    if (r.width <= 1 || r.height <= 1) return false;
    if (r.right < 0 || r.bottom < 0 || r.left > innerWidth || r.top > innerHeight) return false;
    const cs = getComputedStyle(el);
    if (cs.visibility === "hidden" || cs.display === "none" || +cs.opacity === 0) return false;
    return !el.closest("[hidden]");
  };
  const clean = (s) => (s || "").replace(/\s+/g, " ").trim();
  const label = (el) => clean(el.getAttribute("aria-label") || (el.dataset && el.dataset.tip || "").split("\n")[0] || el.getAttribute("title") || el.textContent).slice(0, 90);
  const rect = (el) => { const r = el.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2), w: Math.round(r.width), h: Math.round(r.height) }; };
  let baseKids = new Set(document.body.children);
  const popups = () => [...document.body.children].filter((c) => !baseKids.has(c) && vis(c) && !c.classList.contains("tooltip") && !c.classList.contains("dialog-layer") && !c.matches(".dialog, [role=dialog]"));
  const kindOf = (el) => {
    const cl = el.classList;
    if (el.tagName === "INPUT") return el.closest(".rspinner") ? "spinner-input" : "combo-input";
    if (cl.contains("launcher")) return "launcher";
    if (el.closest(".rspinner")) return "spinner-step";
    if (el.closest(".rcombo")) return "combo-arrow";
    if (cl.contains("rcheck") || el.closest(".rcheck")) return "check";
    if (el.closest(".rgallery")) return el.getAttribute("aria-label") === "Weitere" ? "gallery-more" : (el.closest(".rgallery .scroll, .rgallery-scroll") ? "gallery-scroll" : "gallery-item");
    if (cl.contains("color-main")) return "color-main";
    if (cl.contains("arrow")) return el.closest(".rcolor, .color-split") || el.previousElementSibling?.classList.contains("color-main") ? "color-arrow" : "split-arrow";
    if (el.parentElement?.classList.contains("rsplit")) return "split-main";
    if (cl.contains("has-chev")) return "dropdown";
    return "button";
  };
  const sw = {
    H,
    markBase() { baseKids = new Set(document.body.children); },
    docFp() {
      const e = window.__app.editor();
      if (!e) return "none";
      try { return H(JSON.stringify(e.doc.state(), repl)); } catch (x) { return "ERR " + x; }
    },
    sel() { try { return JSON.stringify(window.__probe.selection()) + (window.__app.editor()?.storyKey ?? ""); } catch { return ""; } },
    dialogs() { return [...document.querySelectorAll(".dialog-layer .dialog, .dialog-layer [role=dialog], body > .dialog, body > [role=dialog]")].filter(vis).map((d) => d.getAttribute("aria-label") || clean(d.querySelector(".dlg-title")?.textContent)); },
    panes() {
      return [...document.querySelectorAll("aside, .taskpane, [role=complementary], .pane, .navpane, .comments-pane")].filter(vis)
        .map((p) => clean(p.querySelector(".tp-title, .pane-title, h2, h3")?.textContent || p.getAttribute("aria-label") || p.className).slice(0, 60));
    },
    ui() {
      const e = window.__app.editor();
      const tab = clean(document.querySelector(".tabrow .tab.active")?.textContent);
      const ctx = [...document.querySelectorAll(".tabrow .tab.contextual")].map((t) => clean(t.textContent));
      const checked = [...document.querySelectorAll(".ribbon-card .checked, .ribbon-card [aria-pressed=true], .ribbon-card input[type=checkbox]:checked")].map(label).join("|");
      const combos = [...document.querySelectorAll(".ribbon-card .rcombo input, .ribbon-card .rspinner input")].map((i) => i.value).join("|");
      const bs = document.querySelector(".backstage");
      const status = clean(document.querySelector(".statusbar, .status-bar")?.textContent);
      let view = "";
      try { view = JSON.stringify({ z: e?.zoom, v: e?.viewMode, m: e?.showMarks, ro: e?.readOnly }); } catch { view = ""; }
      const shell = [...document.querySelectorAll(".ruler, .hruler, .vruler, .gridlines, .navpane")].filter(vis).length;
      return {
        tab, ctx, checked: H(checked), combos, backstage: !!bs && vis(bs), status, view, shell,
        panes: sw.panes(), dialogs: sw.dialogs(), popups: popups().length, sel: sw.sel(),
        docs: window.__app.docs().length, native: (window.__nativeCalls || []).length, bodyCls: document.body.className, minimized: !!document.querySelector(".ribbon-card.hidden, .ribbon.minimized"),
      };
    },
    /** The visible controls of the active ribbon tab, in order. */
    controls() {
      const card = document.querySelector(".ribbon-card");
      const out = [];
      if (!card) return out;
      for (const g of card.querySelectorAll(".rgroup")) {
        const group = clean(g.getAttribute("aria-label") || g.querySelector(".glabel, .rgroup-label, .label")?.textContent || g.dataset.group);
        const seen = new Map();
        for (const el of g.querySelectorAll("button, .rcombo input, .rspinner input, .rcheck")) {
          if (!vis(el)) continue;
          if (el.tagName === "BUTTON" && el.closest(".rcheck") && !el.classList.contains("rcheck")) continue;
          const kind = kindOf(el);
          if (kind === "gallery-scroll") continue;
          const lab = label(el);
          const id = `${g.dataset.group}|${kind}|${lab}`;
          const n = (seen.get(id) || 0) + 1;
          seen.set(id, n);
          out.push({ id: n > 1 ? `${id}#${n}` : id, group, groupId: g.dataset.group, kind, label: lab, key: el.dataset.keytip || el.closest("[data-keytip]")?.dataset.keytip || "",
            disabled: !!el.disabled || el.classList.contains("disabled") || !!el.closest(".disabled"), tip: (el.dataset.tip || el.closest("[data-tip]")?.dataset.tip || "").slice(0, 300), ...rect(el) });
        }
      }
      return out;
    },
    /** Clickable entries of the newest popup (menu, gallery, palette, list) or a submenu. */
    popupItems() {
      const ps = popups();
      const p = ps[ps.length - 1];
      if (!p) return [];
      const out = [];
      const seen = new Map();
      for (const el of p.querySelectorAll("button, [role=menuitem], [role=option], .tile, .item, input, select")) {
        if (!vis(el)) continue;
        if (el.tagName !== "BUTTON" && el.closest("button")) continue;
        if (el.tagName === "INPUT" && el.type === "hidden") continue;
        const txt = clean(el.textContent).slice(0, 90);
        const lab = (el.getAttribute("aria-label") ? label(el) : txt || label(el)) || (el.tagName === "INPUT" ? "[" + (el.placeholder || el.type) + "]" : "");
        const sub = el.classList.contains("has-sub") || el.getAttribute("aria-haspopup") === "true";
        const level = el.closest(".submenu") ? 1 : 0;
        const id = `${level}|${lab}`;
        const n = (seen.get(id) || 0) + 1;
        seen.set(id, n);
        out.push({ id: n > 1 ? `${id}#${n}` : id, label: lab, tag: el.tagName, sub, level, disabled: !!el.disabled || el.classList.contains("disabled"),
          tip: (el.dataset?.tip || "").slice(0, 200), ...rect(el) });
      }
      return out;
    },
    reveal(lab) {
      const ps = popups();
      const p = ps[ps.length - 1];
      if (!p) return false;
      const el = [...p.querySelectorAll("button, [role=menuitem], [role=option], .tile, .item")].find((e) => clean(e.textContent).slice(0, 90) === lab || label(e) === lab);
      if (!el) return false;
      el.scrollIntoView({ block: "center" });
      return true;
    },
    /** The open dialog's controls (its inventory) and its buttons' positions. */
    dialogInfo() {
      const d = [...document.querySelectorAll(".dialog-layer .dialog, .dialog-layer [role=dialog], body > .dialog, body > [role=dialog]")].filter(vis).pop();
      if (!d) return null;
      const controls = [];
      for (const el of d.querySelectorAll("input, select, textarea, button")) {
        if (!vis(el)) continue;
        let lab = el.getAttribute("aria-label") || "";
        if (!lab && el.id) lab = clean(d.querySelector(`label[for="${el.id}"]`)?.textContent);
        if (!lab) lab = clean(el.closest("label")?.textContent);
        if (!lab && el.tagName !== "BUTTON") { const prev = el.previousElementSibling; lab = clean(prev?.textContent).slice(0, 60); }
        if (el.tagName === "BUTTON") lab = lab || clean(el.textContent);
        controls.push({ tag: el.tagName.toLowerCase(), type: el.type || "", label: lab.slice(0, 80), value: el.tagName === "SELECT" ? el.value : (el.type === "checkbox" || el.type === "radio") ? el.checked : (el.value || "").slice(0, 40),
          options: el.tagName === "SELECT" ? el.options.length : undefined, primary: el.classList.contains("primary"), disabled: !!el.disabled, ...rect(el) });
      }
      return { title: d.getAttribute("aria-label") || clean(d.querySelector(".dlg-title")?.textContent), controls };
    },
    /** Puts the document's selection on the first occurrence of `needle` (offsets of __probe.text()). */
    selectText(needle, collapse) {
      const t = window.__probe.text();
      const i = t.indexOf(needle);
      if (i < 0) return false;
      if (collapse) window.__probe.setSelection(i + 2, i + 2); else window.__probe.setSelection(i, i + needle.length);
      return true;
    },
    pictureRect() {
      const im = [...document.querySelectorAll(".page img, .page canvas.picture, .page .picture, .page [data-picture]")].filter(vis)[0];
      return im ? rect(im) : null;
    },
    tabRect(name) {
      const t = [...document.querySelectorAll(".tabrow .tab")].find((b) => clean(b.textContent) === name && vis(b));
      return t ? rect(t) : null;
    },
    tabs() { return [...document.querySelectorAll(".tabrow .tab")].filter(vis).map((b) => ({ label: clean(b.textContent), contextual: b.classList.contains("contextual"), key: b.dataset.keytip || "", ...rect(b) })); },
  };
  window.__sw = sw;
  // The system's dialogs are not opened in the sweep; asking for one is the option's effect (D-133).
  if (window.__app.stubNative && !window.__nativeCalls) window.__app.stubNative();
  if (!window.__swColor) {
    window.__swColor = true;
    document.addEventListener("click", (e) => { if (e.target instanceof HTMLInputElement && e.target.type === "color") { (window.__nativeCalls = window.__nativeCalls || []).push("color"); } }, true);
  }
  return true;
})()
"""
