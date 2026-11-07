// Bildschirmpräsentation > Bildschirme (D-074): which monitor shows the slide show and whether the
// Referentenansicht runs on the other one. With two or more monitors and "Referentenansicht" on, the show
// opens a borderless full-screen window on the show monitor (raabe-host's show window: `window.open` with
// `about:blank#raabe-show:<n>`) and draws the audience's slides into it, while the main window shows the
// Referentenansicht in full screen. Both are one page, so the show keeps one state (slide, animations, ink).
// The choice is this user's setting, not the file's (PowerPoint keeps it in its own settings too).

import { isTauri, type MenuEntry } from "@raabe/shell";

export interface Screen { index: number; name: string; primary: boolean; x: number; y: number; width: number; height: number }

const KEY_MONITOR = "raabe.present.monitor";
const KEY_PRESENTER = "raabe.present.presenterView";

const read = (k: string): string | null => { try { return localStorage.getItem(k); } catch { return null; } };
const write = (k: string, v: string) => { try { localStorage.setItem(k, v); } catch { /* private mode */ } };

/** "Referentenansicht verwenden" (PowerPoint's default: on). */
export const presenterViewOn = (): boolean => read(KEY_PRESENTER) !== "0";
export const setPresenterView = (on: boolean) => write(KEY_PRESENTER, on ? "1" : "0");
/** The chosen show monitor: "auto" (Automatisch), "primary" or a monitor's position as "x,y". */
export const monitorChoice = (): string => read(KEY_MONITOR) ?? "auto";
export const setMonitorChoice = (v: string) => write(KEY_MONITOR, v);

/** Test hook (dev builds and the harness): a list of screens to use in the browser, where there is no Tauri. */
declare global { interface Window { __presentScreens?: Screen[]; __presentKeepWindow?: boolean } }

let cache: Screen[] = [];

/** The monitors in Tauri's order (that order is what the show window's number refers to). */
export async function listScreens(): Promise<Screen[]> {
  if (window.__presentScreens) return (cache = window.__presentScreens);
  if (!isTauri()) return (cache = []);
  const { availableMonitors, primaryMonitor } = await import("@tauri-apps/api/window");
  const [all, primary] = await Promise.all([availableMonitors(), primaryMonitor()]);
  cache = all.map((m, index) => ({ index, name: m.name ?? `Monitor ${index + 1}`, primary: !!primary && m.position.x === primary.position.x && m.position.y === primary.position.y,
    x: m.position.x, y: m.position.y, width: m.size.width, height: m.size.height }));
  return cache;
}

/** The monitor the main window is on (most of its area), or the primary one. */
async function mainScreen(screens: Screen[]): Promise<Screen | undefined> {
  if (!isTauri()) return screens.find((s) => s.primary) ?? screens[0];
  const { getCurrentWindow } = await import("@tauri-apps/api/window");
  const w = getCurrentWindow();
  const [p, sz] = await Promise.all([w.outerPosition(), w.outerSize()]);
  const cx = p.x + sz.width / 2, cy = p.y + sz.height / 2;
  return screens.find((s) => cx >= s.x && cx < s.x + s.width && cy >= s.y && cy < s.y + s.height) ?? screens.find((s) => s.primary) ?? screens[0];
}

/**
 * Where the show goes with two screens: the chosen monitor, or (Automatisch) one the main window is not on,
 * a secondary first, as PowerPoint puts the show on the secondary monitor. Null: one screen only, or the
 * Referentenansicht is off (then PowerPoint shows the slides on the chosen monitor alone; we keep it in the
 * main window).
 */
export async function showTarget(): Promise<{ show: Screen; main: Screen } | null> {
  const screens = await listScreens();
  if (screens.length < 2 || !presenterViewOn()) return null;
  const main = await mainScreen(screens);
  if (!main) return null;
  const choice = monitorChoice();
  let show: Screen | undefined;
  if (choice === "primary") show = screens.find((s) => s.primary);
  else if (choice !== "auto") show = screens.find((s) => `${s.x},${s.y}` === choice);
  if (!show || show === main) show = screens.find((s) => s !== main && !s.primary) ?? screens.find((s) => s !== main);
  return show ? { show, main } : null;
}

/** Opens the show window on a monitor and gives it the page's look (styles, fonts, theme). */
export function openShowWindow(screen: Screen): Window | null {
  const w = isTauri() ? window.open(`about:blank#raabe-show:${screen.index}`, `raabe-show-${Date.now()}`)
    : window.open("about:blank", `raabe-show-${Date.now()}`, `popup,left=${screen.x},top=${screen.y},width=${Math.min(screen.width, 1024)},height=${Math.min(screen.height, 640)}`);
  if (!w) return null;
  const d = w.document;
  d.open(); d.write("<!doctype html><html lang=\"de\"><head><meta charset=\"utf-8\"><title>Bildschirmpräsentation</title></head><body></body></html>"); d.close();
  for (const a of Array.from(document.documentElement.attributes)) if (a.name !== "lang") d.documentElement.setAttribute(a.name, a.value);
  const base = d.createElement("base");
  base.href = document.baseURI;
  d.head.append(base);
  for (const el of Array.from(document.head.querySelectorAll("style, link[rel=stylesheet]"))) d.head.append(d.importNode(el, true));
  // Fonts the page added itself (embedded fonts of the presentation).
  try { for (const f of Array.from(document.fonts)) if ((f as FontFace & { family: string }).family) d.fonts.add(f); } catch { /* added through CSS instead */ }
  d.body.style.cssText = "margin:0;background:#000;overflow:hidden";
  return w;
}

/** Referentenansicht und Bildschirmpräsentation tauschen: the two windows change monitors (Tauri only). */
export async function swapWindows(): Promise<boolean> {
  if (!isTauri()) return false;
  const { getAllWebviewWindows, getCurrentWebviewWindow } = await import("@tauri-apps/api/webviewWindow");
  const { PhysicalPosition } = await import("@tauri-apps/api/dpi");
  const show = (await getAllWebviewWindows()).find((w) => w.label.startsWith("show-"));
  if (!show) return false;
  const main = getCurrentWebviewWindow();
  const [a, b] = await Promise.all([show.outerPosition(), main.outerPosition()]);
  if (document.fullscreenElement) await document.exitFullscreen().catch(() => {});
  await Promise.all([show.setFullscreen(false), main.setFullscreen(false)]);
  await main.setPosition(new PhysicalPosition(a.x + 40, a.y + 40));
  await show.setPosition(new PhysicalPosition(b.x + 40, b.y + 40));
  await Promise.all([show.setFullscreen(true), main.setFullscreen(true)]);
  return true;
}

/** Closes the show windows (a page's window.close does not reach a window the host made). */
export async function closeShowWindows(): Promise<void> {
  if (!isTauri()) return;
  const { getAllWebviewWindows } = await import("@tauri-apps/api/webviewWindow");
  for (const w of await getAllWebviewWindows()) if (w.label.startsWith("show-")) await w.destroy().catch(() => {});
}

/** The main window back where it was (after a swap) and out of full screen. */
export async function restoreMain(at: { x: number; y: number } | null): Promise<void> {
  if (!isTauri()) return;
  const { getCurrentWindow } = await import("@tauri-apps/api/window");
  const { PhysicalPosition } = await import("@tauri-apps/api/dpi");
  const w = getCurrentWindow();
  await w.setFullscreen(false).catch(() => {});
  if (at) await w.setPosition(new PhysicalPosition(at.x, at.y)).catch(() => {});
}

export async function mainPosition(): Promise<{ x: number; y: number } | null> {
  if (!isTauri()) return null;
  const { getCurrentWindow } = await import("@tauri-apps/api/window");
  const p = await getCurrentWindow().outerPosition();
  return { x: p.x, y: p.y };
}

export async function setMainFullscreen(on: boolean): Promise<void> {
  // Harness runs keep the main window where it is (off screen), so a test never covers the owner's screen.
  if (window.__presentKeepWindow) return;
  if (!isTauri()) { if (on) await document.documentElement.requestFullscreen?.().catch(() => {}); return; }
  const { getCurrentWindow } = await import("@tauri-apps/api/window");
  await getCurrentWindow().setFullscreen(on).catch(() => {});
}

/** Bildschirme > Bildschirm: Automatisch, Primärer Monitor and the others, as PowerPoint lists them. */
export function monitorMenu(refresh: () => void): MenuEntry[] {
  void listScreens().then(() => {});
  const choice = monitorChoice();
  const pick = (v: string) => () => { setMonitorChoice(v); refresh(); };
  const others = cache.filter((s) => !s.primary);
  return [
    { label: "Automatisch", checked: choice === "auto", run: pick("auto") },
    { label: "Primärer Monitor", checked: choice === "primary", run: pick("primary") },
    ...others.map((s, i) => ({ label: `Monitor ${i + 2}`, checked: choice === `${s.x},${s.y}`, run: pick(`${s.x},${s.y}`) })),
  ];
}
