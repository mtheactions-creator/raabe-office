// Excel's 1900 date system: serial 1 is 1 Jan 1900, and 1900 counts as a leap year (serial 60
// is the fictitious 29 Feb 1900, kept from Lotus 1-2-3), so serials from 61 on are days since
// 30 Dec 1899. Serial 0 shows as 0 Jan 1900. The last valid date is 31 Dec 9999 (2958465).

export const MAX_SERIAL = 2958465;
const DAY_MS = 86400000;
const BASE = Date.UTC(1899, 11, 31);

/** Serial of a date; month and day may overflow either way (as DATE allows). */
export function serial(y: number, m: number, d: number): number {
  // Normalise the month first; day overflow simply adds days.
  const ym = y * 12 + (m - 1);
  const yy = Math.floor(ym / 12), mm = ym - yy * 12;
  const first = realDays(yy, mm, 1);
  // `first` counts real days (1 Jan 1900 → 1); Excel adds one from 1 Mar 1900 on.
  return first + (yy > 1900 || (yy === 1900 && mm >= 2) ? 1 : 0) + d - 1;
}

function realDays(y: number, m0: number, d: number): number {
  const t = new Date(0);
  t.setUTCFullYear(y, m0, d);
  t.setUTCHours(0, 0, 0, 0);
  return Math.round((t.getTime() - BASE) / DAY_MS);
}

/** Year, month, day of a serial (its whole part); 0 → 1900-01-00, 60 → 1900-02-29. */
export function ymd(n: number): { y: number; m: number; d: number } {
  const s = Math.floor(n);
  if (s === 0) return { y: 1900, m: 1, d: 0 };
  if (s === 60) return { y: 1900, m: 2, d: 29 };
  const real = s < 60 ? s : s - 1;
  const t = new Date(BASE + real * DAY_MS);
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
}

/** Day of the week, 0 = Sunday … 6 = Saturday, as Excel counts it (serial 1 is a Sunday). */
export function weekday0(n: number): number {
  return (((Math.floor(n) - 1) % 7) + 7) % 7;
}

export function isLeap(y: number): boolean {
  return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0 || y === 1900;
}

export function daysInMonth(y: number, m: number): number {
  if (m === 2) return isLeap(y) ? 29 : 28;
  return [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];
}

/** Hours, minutes, seconds of the time part, rounded to the nearest second as Excel does. */
export function hms(n: number): { h: number; mi: number; s: number } {
  const frac = n - Math.floor(n);
  let secs = Math.round(frac * 86400 + 1e-9);
  if (secs >= 86400) secs -= 86400;
  return { h: Math.floor(secs / 3600), mi: Math.floor((secs % 3600) / 60), s: secs % 60 };
}

/** Serial of a JS time in the local time zone (for JETZT and HEUTE). */
export function serialOfDate(t: Date): number {
  const day = serial(t.getFullYear(), t.getMonth() + 1, t.getDate());
  return day + (t.getHours() * 3600 + t.getMinutes() * 60 + t.getSeconds() + t.getMilliseconds() / 1000) / 86400;
}

export const MONTHS_DE = ["januar", "februar", "märz", "april", "mai", "juni", "juli", "august", "september", "oktober", "november", "dezember"];
export const MONTHS_DE_SHORT = ["jan", "feb", "mär", "apr", "mai", "jun", "jul", "aug", "sep", "okt", "nov", "dez"];
