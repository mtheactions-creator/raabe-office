// Excel's date serials. A date is a number of days, the time of day its fraction.
// - 1900 system (Windows default): serial 1 is 1 Jan 1900, and Excel counts the 29 Feb 1900
//   that never was (serial 60, Lotus 1-2-3 compatibility), so from serial 61 (1 Mar 1900) on
//   the serial is one more than the real day count. Serial 0 shows as "00.01.1900".
// - 1904 system (old Mac workbooks): serial 0 is 1 Jan 1904.
// The last date is 31 Dec 9999 in both systems.

export interface YMD { y: number; m: number; d: number }
export interface HMS { h: number; mi: number; s: number; ms: number }

/** Serial of 31 Dec 9999 (1900 system); later values show as "####". */
export const MAX_SERIAL_1900 = 2958465;
/** Serial of 31 Dec 9999 in the 1904 system. */
export const MAX_SERIAL_1904 = 2957003;

export function maxSerial(date1904 = false): number {
  return date1904 ? MAX_SERIAL_1904 : MAX_SERIAL_1900;
}

/** Days from 1 Jan 1970 to a proleptic Gregorian date (month may be out of range). */
function civilDays(y: number, m: number, d: number): number {
  y += Math.floor((m - 1) / 12);
  m = ((m - 1) % 12 + 12) % 12 + 1;
  const yy = m <= 2 ? y - 1 : y;
  const era = Math.floor(yy / 400);
  const yoe = yy - era * 400;
  const mp = (m + 9) % 12;
  const doy = Math.floor((153 * mp + 2) / 5) + d - 1;
  const doe = yoe * 365 + Math.floor(yoe / 4) - Math.floor(yoe / 100) + doy;
  return era * 146097 + doe - 719468;
}

function fromCivil(days: number): YMD {
  const z = days + 719468;
  const era = Math.floor(z / 146097);
  const doe = z - era * 146097;
  const yoe = Math.floor((doe - Math.floor(doe / 1460) + Math.floor(doe / 36524) - Math.floor(doe / 146096)) / 365);
  const doy = doe - (365 * yoe + Math.floor(yoe / 4) - Math.floor(yoe / 100));
  const mp = Math.floor((5 * doy + 2) / 153);
  const d = doy - Math.floor((153 * mp + 2) / 5) + 1;
  const m = mp < 10 ? mp + 3 : mp - 9;
  return { y: yoe + era * 400 + (m <= 2 ? 1 : 0), m, d };
}

const DAY0_1900 = civilDays(1899, 12, 31);
const DAY0_1904 = civilDays(1904, 1, 1);

/** The calendar date of a serial's whole days (fraction ignored, negative values clamp to 0). */
export function serialToDate(serial: number, date1904 = false): YMD {
  const n = Math.max(0, Math.floor(serial));
  if (date1904) return fromCivil(DAY0_1904 + n);
  if (n === 0) return { y: 1900, m: 1, d: 0 };
  if (n === 60) return { y: 1900, m: 2, d: 29 };
  return fromCivil(DAY0_1900 + (n < 60 ? n : n - 1));
}

/** Excel's DATE(y; m; d): years 0-1899 count from 1900, months and days roll over (also
 * backwards), 29 Feb 1900 exists. NaN where Excel gives #ZAHL!: a year outside 1900 (1904)-9999
 * once the months are added (DATE(-1;13;1) is 1.1.1900), or a result outside day 0 to 31.12.9999. */
export function dateToSerial(y: number, m: number, d: number, date1904 = false): number {
  y = Math.trunc(y);
  m = Math.trunc(m);
  d = Math.trunc(d);
  if (y < 1900) y += 1900;
  y += Math.floor((m - 1) / 12);
  m = ((m - 1) % 12 + 12) % 12 + 1;
  // The year after the months are taken into it must lie in the system's range, whatever
  // the days do (DATE(1900;0;31) is #ZAHL!, DATE(10000;-1;1) is 1.11.9999).
  if (y < (date1904 ? 1904 : 1900) || y > 9999) return NaN;
  let serial: number;
  if (date1904) {
    serial = civilDays(y, m, 1) - DAY0_1904 + d - 1;
  } else {
    serial = civilDays(y, m, 1) - DAY0_1900 + d - 1;
    // Months from March 1900 on start a day later because of the fictitious 29 Feb 1900; days
    // counted on from January or February run through it (DATE(1900;2;29) = 60).
    if (y > 1900 || (y === 1900 && m >= 3)) serial += 1;
  }
  if (serial < 0 || serial > maxSerial(date1904)) return NaN;
  return serial;
}

/** Time of day of a serial, rounded to whole milliseconds; 23:59:59.9995 and later carry
 * into the next day (h is then 24 is avoided by the caller rounding the whole serial). */
export function serialToTime(serial: number): HMS {
  const frac = serial - Math.floor(serial);
  let ms = Math.round(frac * 86400000);
  if (ms >= 86400000) ms = 0;
  const h = Math.floor(ms / 3600000);
  ms -= h * 3600000;
  const mi = Math.floor(ms / 60000);
  ms -= mi * 60000;
  const s = Math.floor(ms / 1000);
  return { h, mi, s, ms: ms - s * 1000 };
}

/** TIME(h; m; s) as a fraction of a day (whole days dropped as Excel does). */
export function timeToSerial(h: number, mi: number, s: number): number {
  const secs = Math.trunc(h) * 3600 + Math.trunc(mi) * 60 + Math.trunc(s);
  if (secs < 0) return NaN;
  return (secs % 86400) / 86400;
}

/** Whole seconds of a serial as Excel's HOUR/MINUTE/SECOND see them (rounded to the second). */
function roundedSeconds(serial: number): number {
  return Math.round((serial - Math.floor(serial)) * 86400 + 1e-9);
}

/** YEAR, MONTH, DAY: NaN (#ZAHL!) outside 0 to 31.12.9999. */
export function year(serial: number, date1904 = false): number {
  return inRange(serial, date1904) ? serialToDate(serial, date1904).y : NaN;
}
export function month(serial: number, date1904 = false): number {
  return inRange(serial, date1904) ? serialToDate(serial, date1904).m : NaN;
}
export function day(serial: number, date1904 = false): number {
  return inRange(serial, date1904) ? serialToDate(serial, date1904).d : NaN;
}

function inRange(serial: number, date1904: boolean): boolean {
  return serial >= 0 && serial < maxSerial(date1904) + 1;
}
/** HOUR, MINUTE, SECOND: rounded to the second; NaN outside 0 to 31.12.9999 23:59:59. */
export function hour(serial: number, date1904 = false): number {
  return inRange(serial, date1904) ? Math.floor((roundedSeconds(serial) % 86400) / 3600) : NaN;
}
export function minute(serial: number, date1904 = false): number {
  return inRange(serial, date1904) ? Math.floor((roundedSeconds(serial) % 3600) / 60) : NaN;
}
export function second(serial: number, date1904 = false): number {
  return inRange(serial, date1904) ? roundedSeconds(serial) % 60 : NaN;
}

/** Day of the week, 0 = Sunday … 6 = Saturday, as Excel sees it (1 Jan 1900 is a Sunday). */
export function dayOfWeek(serial: number, date1904 = false): number {
  const n = Math.floor(serial);
  return (((n + (date1904 ? 5 : 6)) % 7) + 7) % 7;
}

/** Excel's WEEKDAY(serial; type). NaN for an unknown type. */
export function weekday(serial: number, type = 1, date1904 = false): number {
  if (!inRange(serial, date1904)) return NaN;
  const dow = dayOfWeek(serial, date1904); // 0 = Sunday
  switch (type) {
    case 1: case 17: return dow + 1;
    case 2: case 11: return ((dow + 6) % 7) + 1;
    case 3: return (dow + 6) % 7;
    case 12: return ((dow + 5) % 7) + 1;
    case 13: return ((dow + 4) % 7) + 1;
    case 14: return ((dow + 3) % 7) + 1;
    case 15: return ((dow + 2) % 7) + 1;
    case 16: return ((dow + 1) % 7) + 1;
    default: return NaN;
  }
}

/** Excel's leap year rule: 1900 counts as a leap year in the 1900 system. */
export function isLeapYear(y: number, date1904 = false): boolean {
  if (y === 1900 && !date1904) return true;
  return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
}

export function daysInMonth(y: number, m: number, date1904 = false): number {
  y += Math.floor((m - 1) / 12);
  m = ((m - 1) % 12 + 12) % 12 + 1;
  return [31, isLeapYear(y, date1904) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];
}

/** EDATE: the same day `months` later, clamped to the end of a shorter month. Serial 0 counts as
 * 0 Jan 1900 and 60 as 29 Feb 1900, but the month lengths are the real ones (EDATE(60;0) is 28
 * Feb). Results before 1900 are NaN; the 1904 system allows negative results. */
export function addMonths(serial: number, months: number, date1904 = false): number {
  if (!inRange(serial, date1904)) return NaN;
  const { y, m, d } = serialToDate(serial, date1904);
  const tm = m + Math.trunc(months);
  const ty = y + Math.floor((tm - 1) / 12), tmm = ((tm - 1) % 12 + 12) % 12 + 1;
  return monthSerial(ty, tmm, Math.min(d, realDaysInMonth(ty, tmm)), date1904);
}

/** EOMONTH: the last day of the month `months` after the serial's month (real calendar). */
export function endOfMonth(serial: number, months: number, date1904 = false): number {
  if (!inRange(serial, date1904)) return NaN;
  const { y, m } = serialToDate(serial, date1904);
  const tm = m + Math.trunc(months);
  const ty = y + Math.floor((tm - 1) / 12), tmm = ((tm - 1) % 12 + 12) % 12 + 1;
  return monthSerial(ty, tmm, realDaysInMonth(ty, tmm), date1904);
}

function realDaysInMonth(y: number, m: number): number {
  return [31, (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0 ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];
}

/** Serial of a real calendar date for EDATE/EOMONTH: NaN outside the system's range. */
function monthSerial(y: number, m: number, d: number, date1904: boolean): number {
  if (y > 9999) return NaN;
  if (date1904) return civilDays(y, m, d) - DAY0_1904;
  if (y < 1900) return NaN;
  const s = civilDays(y, m, d) - DAY0_1900;
  return s >= 60 ? s + 1 : s;
}

/** A JavaScript Date (local time) as a serial, for HEUTE()/JETZT(). */
export function serialFromJsDate(date: Date, date1904 = false): number {
  const days = dateToSerial(date.getFullYear(), date.getMonth() + 1, date.getDate(), date1904);
  const secs = date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds() + date.getMilliseconds() / 1000;
  return days + secs / 86400;
}

export const MONTHS_DE = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];
export const MONTHS_SHORT_DE = ["Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"];
export const DAYS_DE = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"];
export const DAYS_SHORT_DE = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];
export const MONTHS_EN = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
export const DAYS_EN = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
