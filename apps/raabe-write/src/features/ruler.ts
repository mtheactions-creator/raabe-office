// Ansicht > Lineal: Word's horizontal ruler above the page and vertical ruler beside it
// (recorded 2026-10-14): margins grey, the text area white, centimetres counted from the
// margin, the indent markers (Erstzeileneinzug, Hängender Einzug, Linker Einzug, Rechter
// Einzug), the paragraph's tab stops and the default stops, the tab selector in the corner.
// Dragging a marker or a margin edge changes it (0.25 cm steps, Alt: free, as in Word) with a
// dotted guide line over the text; a click on the ruler sets a tab stop of the selected kind,
// dragging one off the ruler removes it.

import { h } from "@raabe/shell";
import { PT, type Editor, type TabStop } from "@raabe/text-engine";
import type { AppApi } from "./api";
import type { Dock } from "./dock";

const CM = 28.3465;
const SIZE = 22;
type Kind = "left" | "center" | "right" | "decimal" | "bar" | "first" | "hanging";
const KINDS: [Kind, string][] = [["left", "Linksbündiger Tabstopp"], ["center", "Zentrierter Tabstopp"], ["right", "Rechtsbündiger Tabstopp"],
  ["decimal", "Dezimaltabstopp"], ["bar", "Vertikale Linie"], ["first", "Erstzeileneinzug"], ["hanging", "Hängender Einzug"]];

/** The tab mark Word draws for a kind (16 × 16). */
function tabGlyph(kind: string): string {
  const p = (d: string) => `<svg viewBox="0 0 16 16" width="16" height="16"><path d="${d}" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>`;
  switch (kind) {
    case "center": return p("M8 4v8M3 12h10");
    case "right": return p("M11 4v8H4");
    case "decimal": return p("M8 4v8M3 12h10") .replace("</svg>", '<circle cx="11.5" cy="8" r="1.1" fill="currentColor"/></svg>');
    case "bar": return p("M8 3v10");
    case "first": return '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M4 4h8v3l-4 4-4-4z" fill="#fff" stroke="currentColor"/></svg>';
    case "hanging": return '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M4 12h8V9L8 5 4 9z" fill="#fff" stroke="currentColor"/></svg>';
    default: return p("M5 4v8h7");
  }
}

export function rulerFeature(api: AppApi, docArea: HTMLElement, dock: Dock) {
  const ed = () => api.editor();
  let on = false;
  let kind = 0;
  const corner = h("button", { class: "rul-corner", "aria-label": KINDS[0][1], "data-tip": KINDS[0][1] });
  const hcanvas = h("canvas", { class: "rul-canvas" }) as HTMLCanvasElement;
  const hmarks = h("div", { class: "rul-marks" });
  const htrack = h("div", { class: "rul-track" }, hcanvas, hmarks);
  const hruler = h("div", { class: "hruler", hidden: true }, corner, htrack);
  const vcanvas = h("canvas", { class: "rul-canvas" }) as HTMLCanvasElement;
  const vmarks = h("div", { class: "rul-marks" });
  const vruler = h("div", { class: "vruler", hidden: true }, vcanvas, vmarks);
  const guide = h("div", { class: "rul-guide", hidden: true });
  docArea.append(hruler, vruler, guide);
  dock.add(hruler, "top", SIZE);
  dock.add(vruler, "inner", SIZE);

  const paintCorner = () => {
    corner.innerHTML = tabGlyph(KINDS[kind][0]);
    corner.setAttribute("data-tip", KINDS[kind][1]);
    corner.setAttribute("aria-label", KINDS[kind][1]);
  };
  corner.onclick = () => { kind = (kind + 1) % KINDS.length; paintCorner(); };
  paintCorner();

  const show = (v: boolean) => {
    on = v;
    sync();
    api.refresh();
  };
  /** Which rulers the view has (Lesemodus none, Weblayout no vertical one). */
  const sync = () => {
    const mode = ed()?.viewMode ?? "print";
    hruler.hidden = !on || mode === "read";
    vruler.hidden = !on || mode !== "print";
    corner.style.visibility = vruler.hidden ? "hidden" : "";
    corner.style.width = vruler.hidden ? "0" : `${SIZE}px`;
    render();
  };

  // ----- geometry -------------------------------------------------------------------------------
  interface Geo { e: Editor; page: number; x0: number; y0: number; s: number; ml: number; mr: number; mt: number; mb: number; w: number; hgt: number; inTable: boolean }
  const geometry = (): Geo | null => {
    const e = ed();
    if (!e) return null;
    const page = e.stats().caretPage;
    const el = e.view.pageElement(page);
    const setup = e.layout.pages[page]?.setup;
    if (!el || !setup) return null;
    const pr = el.getBoundingClientRect();
    const s = pr.width / (setup.width * PT);
    return { e, page, x0: pr.left, y0: pr.top, s, ml: setup.marginLeft, mr: setup.marginRight, mt: setup.marginTop, mb: setup.marginBottom, w: setup.width, hgt: setup.height, inTable: e.inTable() };
  };

  const prep = (c: HTMLCanvasElement, w: number, hgt: number) => {
    const dpr = devicePixelRatio || 1;
    if (c.width !== Math.round(w * dpr) || c.height !== Math.round(hgt * dpr)) { c.width = Math.round(w * dpr); c.height = Math.round(hgt * dpr); c.style.width = `${w}px`; c.style.height = `${hgt}px`; }
    const g = c.getContext("2d")!;
    g.setTransform(dpr, 0, 0, dpr, 0, 0);
    g.clearRect(0, 0, w, hgt);
    return g;
  };

  /** Ticks and numbers along an axis: `from` is the margin (0 cm) in px, `px` px per pt. */
  const scale = (g: CanvasRenderingContext2D, horizontal: boolean, start: number, end: number, zero: number, px: number, len: number) => {
    g.fillStyle = "#5c5c5c";
    g.font = "9px 'Segoe UI', sans-serif";
    g.textAlign = "center";
    g.textBaseline = "middle";
    const step = 0.25 * CM * px;
    if (step < 2) return;
    const mid = len / 2;
    for (const dir of [1, -1]) {
      for (let k = dir > 0 ? 0 : 1; ; k++) {
        const at = zero + dir * k * step;
        if (at < start - 1 || at > end + 1) break;
        const q = k % 4;
        if (q === 0 && k > 0) {
          const label = String(k / 4);
          if (horizontal) g.fillText(label, at, mid + 0.5);
          else { g.save(); g.translate(mid, at); g.fillText(label, 0, 0.5); g.restore(); }
        } else if (k > 0) {
          const tick = q === 2 ? 5 : 2;
          if (horizontal) g.fillRect(Math.round(at), mid - tick / 2, 1, tick);
          else g.fillRect(mid - tick / 2, Math.round(at), tick, 1);
        }
      }
    }
  };

  let queued = false;
  const render = () => {
    if (queued) return;
    queued = true;
    requestAnimationFrame(() => { queued = false; draw(); });
  };

  const draw = () => {
    const geo = geometry();
    if (!geo || (hruler.hidden && vruler.hidden)) { hmarks.replaceChildren(); vmarks.replaceChildren(); return; }
    const { e, s } = geo;
    const px = s * PT;
    if (!hruler.hidden) {
      const tr = htrack.getBoundingClientRect();
      const g = prep(hcanvas, tr.width, SIZE);
      const left = geo.x0 - tr.left, right = left + geo.w * px;
      const tl = left + geo.ml * px, trt = right - geo.mr * px;
      // The strip: grey margins, white text area.
      g.fillStyle = "#dcdcdc";
      g.fillRect(left, 4, right - left, SIZE - 8);
      g.fillStyle = "#ffffff";
      g.fillRect(tl, 4, trt - tl, SIZE - 8);
      scale(g, true, left, right, tl, px, SIZE);
      // Default tab stops: small marks at the bottom after the last stop.
      const rp = e.paraState();
      const tabs = (rp.tabs ?? []).filter((t) => t.align !== "clear" && t.align !== "num");
      const def = e.doc.settings.defaultTab || 35.4;
      const last = Math.max(rp.indLeft, ...tabs.map((t) => t.pos));
      g.fillStyle = "#9a9a9a";
      for (let p = (Math.floor(last / def) + 1) * def; tl + p * px < trt; p += def) g.fillRect(Math.round(tl + p * px), SIZE - 5, 1, 2);
      const marks: HTMLElement[] = [];
      marks.push(edge("Linker Rand", tl, "h"), edge("Rechter Rand", trt, "h"));
      if (!geo.inTable) {
        marks.push(marker("first", "Erstzeileneinzug", tl + (rp.indLeft + rp.indFirst) * px, 0));
        marks.push(marker("hanging", "Hängender Einzug", tl + rp.indLeft * px, 9));
        marks.push(marker("leftbox", "Linker Einzug", tl + rp.indLeft * px, 16));
        marks.push(marker("right", "Rechter Einzug", trt - rp.indRight * px, 9));
        for (const t of tabs) marks.push(tabMark(t, tl + t.pos * px));
      }
      hmarks.replaceChildren(...marks);
      hmarks.dataset.tl = String(tl);
      hmarks.dataset.px = String(px);
    }
    if (!vruler.hidden) {
      const vr = vruler.getBoundingClientRect();
      const g = prep(vcanvas, SIZE, vr.height);
      const top = geo.y0 - vr.top, bottom = top + geo.hgt * px;
      const tt = top + geo.mt * px, tb = bottom - geo.mb * px;
      g.fillStyle = "#dcdcdc";
      g.fillRect(4, top, SIZE - 8, bottom - top);
      g.fillStyle = "#ffffff";
      g.fillRect(4, tt, SIZE - 8, tb - tt);
      scale(g, false, top, bottom, tt, px, SIZE);
      vmarks.replaceChildren(edge("Oberer Rand", tt, "v"), edge("Unterer Rand", tb, "v"));
    }
  };

  // ----- markers and dragging -------------------------------------------------------------------
  const marker = (k: "first" | "hanging" | "leftbox" | "right", tip: string, x: number, y: number) => {
    const el = h("div", { class: `rul-m rul-${k}`, "data-tip": tip, style: `left:${x}px;top:${y}px` });
    el.innerHTML = k === "leftbox" ? '<svg viewBox="0 0 10 6" width="10" height="6"><rect x="0.5" y="0.5" width="9" height="5" fill="#fff" stroke="#5c5c5c"/></svg>'
      : k === "first" ? '<svg viewBox="0 0 10 9" width="10" height="9"><path d="M0.5 0.5h9v4l-4.5 4-4.5-4z" fill="#fff" stroke="#5c5c5c"/></svg>'
      : '<svg viewBox="0 0 10 9" width="10" height="9"><path d="M0.5 8.5h9v-4L5 0.5 0.5 4.5z" fill="#fff" stroke="#5c5c5c"/></svg>';
    el.onmousedown = (ev) => drag(ev, k);
    return el;
  };
  const tabMark = (t: TabStop, x: number) => {
    const el = h("div", { class: "rul-tab", style: `left:${x}px`, "data-tip": KINDS.find(([k]) => k === t.align)?.[1] ?? "Tabstopp" });
    el.innerHTML = tabGlyph(t.align).replace('width="16" height="16"', 'width="9" height="9"');
    el.onmousedown = (ev) => drag(ev, "tab", t);
    return el;
  };
  const edge = (tip: string, at: number, axis: "h" | "v") => {
    const el = h("div", { class: `rul-edge ${axis}`, "data-tip": tip, style: axis === "h" ? `left:${at}px` : `top:${at}px` });
    el.onmousedown = (ev) => drag(ev, tip);
    el.ondblclick = () => api.commands.run("layout.pageSetup");
    return el;
  };

  const snap = (pt: number, free: boolean) => (free ? Math.round(pt * 10) / 10 : Math.round(pt / (0.25 * CM)) * 0.25 * CM);

  /** Drags a marker; the change is made on release (one undo step), the guide shows where. */
  const drag = (ev: MouseEvent, what: string, tab?: TabStop) => {
    ev.preventDefault();
    ev.stopPropagation();
    const geo = geometry();
    if (!geo) return;
    const { e } = geo;
    const px = geo.s * PT;
    const vertical = what === "Oberer Rand" || what === "Unterer Rand";
    const area = docArea.getBoundingClientRect();
    const tr = htrack.getBoundingClientRect();
    const tl = geo.x0 + geo.ml * px, trt = geo.x0 + (geo.w - geo.mr) * px;
    const rp = e.paraState();
    let value = 0;
    let off = false;
    const at = (cx: number, cy: number, free: boolean) => {
      if (vertical) {
        const y = what === "Oberer Rand" ? snap((cy - geo.y0) / px, free) : snap((geo.y0 + geo.hgt * px - cy) / px, free);
        value = Math.max(0, Math.min(geo.hgt / 2, y));
        guide.className = "rul-guide v";
        guide.style.cssText = `top:${(what === "Oberer Rand" ? geo.y0 + value * px : geo.y0 + (geo.hgt - value) * px) - area.top}px`;
      } else {
        let x: number;
        if (what === "Linker Rand") { value = Math.max(0, Math.min(geo.w - geo.mr - 36, snap((cx - geo.x0) / px, free))); x = geo.x0 + value * px; }
        else if (what === "Rechter Rand") { value = Math.max(0, Math.min(geo.w - geo.ml - 36, snap((geo.x0 + geo.w * px - cx) / px, free))); x = geo.x0 + (geo.w - value) * px; }
        else if (what === "right") { value = Math.max(-geo.mr, snap((trt - cx) / px, free)); x = trt - value * px; }
        else { value = snap((cx - tl) / px, free); x = tl + value * px; }
        off = what === "tab" && Math.abs(cy - (tr.top + SIZE / 2)) > 24;
        guide.className = "rul-guide h";
        guide.style.cssText = `left:${x - area.left}px;${off ? "display:none" : ""}`;
      }
      guide.hidden = false;
    };
    at(ev.clientX, ev.clientY, ev.altKey);
    const move = (m: MouseEvent) => at(m.clientX, m.clientY, m.altKey);
    const up = (m: MouseEvent) => {
      removeEventListener("mousemove", move);
      removeEventListener("mouseup", up, true);
      at(m.clientX, m.clientY, m.altKey);
      guide.hidden = true;
      apply(e, what, value, rp, tab, off);
      render();
      api.refresh();
    };
    addEventListener("mousemove", move);
    addEventListener("mouseup", up, true);
  };

  const apply = (e: Editor, what: string, value: number, rp: ReturnType<Editor["paraState"]>, tab: TabStop | undefined, off: boolean) => {
    const margin = (k: "marginLeft" | "marginRight" | "marginTop" | "marginBottom") =>
      e.changeSections((s) => ({ ...s, page: { ...s.page, [k]: value } }), false);
    switch (what) {
      case "Linker Rand": margin("marginLeft"); return;
      case "Rechter Rand": margin("marginRight"); return;
      case "Oberer Rand": margin("marginTop"); return;
      case "Unterer Rand": margin("marginBottom"); return;
      case "first": e.setParaProps({ indFirst: value - rp.indLeft }); return;
      // Hängender Einzug moves the other lines; the first line stays where it is.
      case "hanging": e.setParaProps({ indLeft: value, indFirst: rp.indLeft + rp.indFirst - value }); return;
      case "leftbox": e.setParaProps({ indLeft: value }); return;
      case "right": e.setParaProps({ indRight: value }); return;
      case "tab": {
        if (!tab) return;
        const rest = (rp.tabs ?? []).filter((t) => Math.abs(t.pos - tab.pos) > 0.05 && t.align !== "clear" && t.align !== "num");
        setTabs(e, off ? rest : [...rest, { ...tab, pos: value }]);
      }
    }
  };

  /** Gives the selected paragraphs these tab stops: direct stops plus clears for their style's. */
  const setTabs = (e: Editor, want: readonly TabStop[]) => {
    const sorted = [...want].sort((a, b) => a.pos - b.pos);
    e.setParaProps((_rp, old) => {
      const style = e.doc.styles.resolvePara({ style: old.style } as never).tabs ?? [];
      const near = (a: TabStop, b: TabStop) => Math.abs(a.pos - b.pos) < 0.05;
      const direct: TabStop[] = sorted.filter((t) => !style.some((s) => near(s, t) && s.align === t.align));
      for (const s of style) if (s.align !== "clear" && !sorted.some((t) => near(s, t))) direct.push({ pos: s.pos, align: "clear" });
      return { tabs: direct.length ? direct.sort((a, b) => a.pos - b.pos) : undefined };
    });
  };

  // A click on the white strip sets a tab stop (or an indent, with those kinds selected).
  htrack.addEventListener("mousedown", (ev) => {
    if ((ev.target as Element).closest(".rul-m, .rul-tab, .rul-edge")) return;
    const geo = geometry();
    if (!geo || geo.inTable) return;
    const px = geo.s * PT;
    const pos = snap((ev.clientX - (geo.x0 + geo.ml * px)) / px, ev.altKey);
    if (pos < 0 || pos > geo.w - geo.ml - geo.mr) return;
    ev.preventDefault();
    const e = geo.e;
    const k = KINDS[kind][0];
    const rp = e.paraState();
    if (k === "first") e.setParaProps({ indFirst: pos - rp.indLeft });
    else if (k === "hanging") e.setParaProps({ indLeft: pos, indFirst: rp.indLeft + rp.indFirst - pos });
    else {
      const tabs = (rp.tabs ?? []).filter((t) => t.align !== "clear" && t.align !== "num" && Math.abs(t.pos - pos) > 0.05);
      setTabs(e, [...tabs, { pos, align: k }]);
    }
    render();
    api.refresh();
    api.focus();
  });
  htrack.addEventListener("dblclick", () => api.commands.run("para.dialog"));

  // The rulers follow scrolling (scroll events reach the area while capturing), resizing and the zoom.
  docArea.addEventListener("scroll", () => { if (on) render(); }, true);
  addEventListener("resize", () => { if (on) render(); });

  api.commands.implement("view.ruler", { run: () => show(!on), enabled: () => !!ed(), checked: () => on });
  return { refresh: () => { if (on) sync(); } };
}

