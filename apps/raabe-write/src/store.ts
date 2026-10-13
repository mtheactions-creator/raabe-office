// Settings and the recent-file list, kept in the WebView's local storage (per user, on
// this device). Reads never throw: a broken or missing store falls back to defaults.

export interface Settings {
  userName: string;
  autoRecoverSeconds: number;
  touchMode: boolean;
}

export interface RecentFile {
  path: string;
  name: string;
  opened: number;
  pinned: boolean;
}

const DEFAULTS: Settings = { userName: "", autoRecoverSeconds: 30, touchMode: false };
const KEY_SETTINGS = "raabe.write.settings";
const KEY_RECENT = "raabe.write.recent";
const MAX_RECENT = 50;

function read<T>(key: string, fallback: T): T {
  try {
    const raw = localStorage.getItem(key);
    return raw ? { ...fallback, ...JSON.parse(raw) } : fallback;
  } catch {
    return fallback;
  }
}

function write(key: string, value: unknown): void {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Storage full or blocked: settings simply are not remembered.
  }
}

export function loadSettings(): Settings {
  return read(KEY_SETTINGS, DEFAULTS);
}

export function saveSettings(s: Settings): void {
  write(KEY_SETTINGS, s);
}

export function loadRecent(): RecentFile[] {
  try {
    const raw = localStorage.getItem(KEY_RECENT);
    const list = raw ? (JSON.parse(raw) as RecentFile[]) : [];
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
}

const same = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();

export function addRecent(path: string, name: string): RecentFile[] {
  const list = loadRecent();
  const old = list.find((r) => same(r.path, path));
  const next = [{ path, name, opened: Date.now(), pinned: old?.pinned ?? false }, ...list.filter((r) => !same(r.path, path))].slice(0, MAX_RECENT);
  write(KEY_RECENT, next);
  return next;
}

export function setPinned(path: string, pinned: boolean): void {
  write(KEY_RECENT, loadRecent().map((r) => (same(r.path, path) ? { ...r, pinned } : r)));
}

export function removeRecent(path: string): void {
  write(KEY_RECENT, loadRecent().filter((r) => !same(r.path, path)));
}
