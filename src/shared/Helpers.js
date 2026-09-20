/**
 * Helpers.js
 *
 * Shared pure utility helpers used across phases.
 *
 * This file is the home for small, reusable functions that don't own a
 * whole subsystem (unlike ImportIssues, SheetBlanking, or SettingsService).
 *
  * Current contents:
 *   - CUSIP normalization / detection
 *   - UI-safe alerts
 *   - Number parsing
 *   - Header / column helpers
 *   - Date / time normalization
 *   - String helper (toStr)
 *
 * Add new helper groups below with a clear section banner.
 * If this file later becomes large or mixed, we can split by topic.
 */

// =========================================================================
// CUSIP HELPERS
//   Used by Phase 1 (BuildUnifiedImportV3) and Phase 2 (mapSchwabImportByHeadersV3).
// =========================================================================

/**
 * Turn any raw value into a stable CUSIP lookup key.
 * - Numbers are truncated and left-padded to 9 digits.
 * - Strings are uppercased and stripped to [0-9A-Z] only.
 */
function normalizeCusip(v) {
  if (typeof v === "number" && isFinite(v)) {
    return String(Math.trunc(v)).padStart(9, "0");
  }
  return String(v == null ? "" : v)
    .toUpperCase()
    .replace(/[^0-9A-Z]/g, "");
}

/**
 * True when the value normalizes to a plausible CUSIP:
 * exactly 9 alphanumeric characters and at least one digit.
 */
function looksLikeCusip(v) {
  const s = normalizeCusip(v);
  return /^[0-9A-Z]{9}$/.test(s) && /[0-9]/.test(s);
}



// Debug test after moving CUSIP to use this file
function testCusipHelpers() {
  const samples = [
    "50545P309",
    " 50545p309 ",
    50545309,
    "00848K101",
    "EU",
    "LUR/CN",
    "ABCDEFGHI",   // 9 letters, no digit
    "",
    null,
  ];
  samples.forEach((v) => {
    Logger.log(
      JSON.stringify(v) +
        " → normalize=" +
        normalizeCusip(v) +
        " looksLike=" +
        looksLikeCusip(v),
    );
  });
}

// =========================================================================
// STRING HELPERS
//   Safe "always a string" conversion used by Phase 1 factories and
//   buildUnifiedImportV3.
// =========================================================================

/**
 * Convert any value to a string.
 * null / undefined → "" so callers can .trim() / .toUpperCase() safely.
 */
function toStr(v) {
  return v === null || v === undefined ? "" : String(v);
}
// =========================================================================
// UI HELPERS
//   Safe alerts that work from the spreadsheet menu and do not crash when
//   there is no UI (triggers, API, clasp run, etc.).
// =========================================================================

/**
 * Show a UI alert when possible; otherwise write to Logger.
 * Prefer this over SpreadsheetApp.getUi().alert() in shared/pipeline code.
 */
function uiAlertSafe(message) {
  try {
    SpreadsheetApp.getUi().alert(String(message || ""));
  } catch (e) {
    Logger.log("uiAlertSafe (no UI): " + message);
  }
}

/**
 * Backward-compatible alias used by the TOS import pipeline.
 * New code should call uiAlertSafe() directly.
 */
function tosUiAlertSafe(message) {
  uiAlertSafe(message);
}

// UI Test if working
function testUiAlertSafe() {
  uiAlertSafe("Helpers.js uiAlertSafe is working");
}

// =========================================================================
// NUMBER HELPERS
//   Shared parsing for quantities, prices, fees, and amounts.
//
//   toNum(v)        → Number or NaN   (use for math / isNaN checks)
//   parseNumber(v)  → Number or ""    (use when writing sheet cells)
//
//   Both understand commas, $, and accounting parentheses: (123.45) → -123.45
// =========================================================================

/**
 * Parse a value to a number for math.
 * Blank or invalid → NaN (so callers can use isNaN()).
 */
function toNum(v) {
  if (v === null || v === undefined || v === "") return NaN;
  if (typeof v === "number") return isFinite(v) ? v : NaN;

  const s = String(v).trim();
  if (!s) return NaN;

  // Accounting format: (123.45) means negative
  const neg = /^\(.*\)$/.test(s);
  const cleaned = s
    .replace(/[(),$]/g, "")
    .replace(/,/g, "")
    .trim();

  const n = Number(cleaned);
  if (isNaN(n)) return NaN;
  return neg ? -n : n;
}

/**
 * Parse a value for writing to a sheet cell.
 * Blank or invalid → "" so the cell stays empty.
 */
function parseNumber(v) {
  if (v === null || v === undefined || v === "") return "";
  const n = toNum(v);
  return isNaN(n) ? "" : n;
}

// Fast check (optional):
function testNumberHelpers() {
  const samples = [
    123.45,
    "1,234.56",
    "$99.00",
    "(50.25)",
    "  ",
    "",
    null,
    "abc",
  ];
  samples.forEach((v) => {
    Logger.log(
      JSON.stringify(v) +
        " → toNum=" +
        toNum(v) +
        " parseNumber=" +
        JSON.stringify(parseNumber(v)),
    );
  });
}

// =========================================================================
// HEADER / COLUMN HELPERS
//   Build a header→index map and look up columns by name.
//   normalizeHeader collapses whitespace so "Trade  Date" still matches.
// =========================================================================

/**
 * Normalize a header string for map keys: trim, lower-case, collapse spaces.
 */
function normalizeHeader(s) {
  return String(s == null ? "" : s)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

/**
 * Build a map: normalized header → column index (0-based).
 */
function buildHeaderIndexMap(headers) {
  const map = {};
  for (let c = 0; c < headers.length; c++) {
    const key = normalizeHeader(headers[c]);
    if (key) map[key] = c;
  }
  return map;
}

/**
 * Look up a column index by header name. Throws if not found.
 */
function col(headerMap, name) {
  const key = normalizeHeader(name);
  const idx = headerMap[key];
  if (typeof idx === "undefined") {
    throw new Error("Header not found: " + name);
  }
  return idx;
}

/**
 * Look up a column index; supports one name or an array of alternate names.
 * Returns null if none match (does not throw).
 */
function colOrNull(headerMap, nameOrNames) {
  const names = Array.isArray(nameOrNames) ? nameOrNames : [nameOrNames];
  for (let i = 0; i < names.length; i++) {
    const key = normalizeHeader(names[i]);
    if (typeof headerMap[key] !== "undefined") return headerMap[key];
  }
  return null;
}

/**
 * Throw if any required header names are missing from the map.
 */
function requireHeaders(headerMap, requiredHeaders, where) {
  const missing = [];
  for (let i = 0; i < requiredHeaders.length; i++) {
    const k = normalizeHeader(requiredHeaders[i]);
    if (typeof headerMap[k] === "undefined") missing.push(requiredHeaders[i]);
  }
  if (missing.length) {
    throw new Error(
      "Missing headers in " + where + ": " + missing.join(", "),
    );
  }
}

// Header Fast check
function testHeaderHelpers() {
  const headers = ["Account", "Trade  Date", "Quantity"];
  const map = buildHeaderIndexMap(headers);
  Logger.log(JSON.stringify(map));
  Logger.log("Account → " + col(map, "account"));
  Logger.log("Trade Date → " + col(map, "Trade Date")); // collapsed spaces
  Logger.log("Missing → " + colOrNull(map, "NoSuchColumn"));
}

// =========================================================================
// DATE / TIME HELPERS
//   One normalize path for Phase 1 Combined writes and Build Unified matching.
//
//   normalizeDate(v)          → "yyyy-MM-dd" or "" / raw fallback
//   normalizeTime(v)          → "HHmm"   (minute bucket)
//   normalizeTimeHHmmss(v)    → "HHmmss" (second precision)
//   toDateObject(date, time)  → Date from yyyy-MM-dd + HHmm or HHmmss
//
//   These do NOT apply ET→CT. Pipeline helpers tosEtToCtHHmmss /
//   tosTradesParseExecTimeRaw do that when reading raw TOS Eastern times.
//   Sheet values in TosTop / TosTrades / Combined are already Central Time.
// =========================================================================

function normalizeDate(v) {
  if (v instanceof Date)
    return Utilities.formatDate(
      v,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
  const s = v === null || v === undefined ? "" : String(v).trim();
  if (!s) return "";

  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);

  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    const mm = String(parseInt(m[1], 10)).padStart(2, "0");
    const dd = String(parseInt(m[2], 10)).padStart(2, "0");
    let yyyy = parseInt(m[3], 10);
    if (yyyy < 100) yyyy += 2000;
    return yyyy + "-" + mm + "-" + dd;
  }

  const digits = s.replace(/\D/g, "");
  if (digits.length === 7 || digits.length === 8) {
    const mm =
      digits.length === 7
        ? "0" + digits.substring(0, 1)
        : digits.substring(0, 2);
    const dd =
      digits.length === 7 ? digits.substring(1, 3) : digits.substring(2, 4);
    const yyyy =
      digits.length === 7 ? digits.substring(3, 7) : digits.substring(4, 8);
    if (/^\d{4}$/.test(yyyy)) return yyyy + "-" + mm + "-" + dd;
  }

  return s;
}

function normalizeTime(v) {
  if (v instanceof Date)
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "HHmm");

  const s = v === null || v === undefined ? "" : String(v).trim();
  if (!s) return "";

  const digits = s.replace(/\D/g, "");
  if (!digits) return "";

  if (digits.length >= 6) return digits.substring(0, 4);
  if (digits.length === 1) return ("000" + digits).slice(-4);
  if (digits.length === 2) return ("00" + digits).slice(-4);
  if (digits.length === 3) return ("0" + digits).slice(-4);
  return digits.substring(0, 4).padStart(4, "0");
}

function normalizeTimeHHmmss(v) {
  if (v instanceof Date)
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "HHmmss");

  const s = v === null || v === undefined ? "" : String(v).trim();
  if (!s) return "";

  const digits = s.replace(/\D/g, "");
  if (!digits) return "";

  if (digits.length <= 4) {
    const hhmm = digits.padStart(4, "0").slice(-4);
    return hhmm + "00";
  }

  const padded = digits.padStart(6, "0");
  return padded.substring(0, 6);
}

function toDateObject(yyyyMmDd, hhmmOrHhmmss) {
  const d = yyyyMmDd === null || yyyyMmDd === undefined ? "" : String(yyyyMmDd).trim();
  const t =
    hhmmOrHhmmss === null || hhmmOrHhmmss === undefined
      ? ""
      : String(hhmmOrHhmmss).trim();

  const m = d.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;

  const yyyy = parseInt(m[1], 10);
  const mm = parseInt(m[2], 10);
  const dd = parseInt(m[3], 10);

  const digits = t.replace(/\D/g, "");
  if (!digits) return new Date(yyyy, mm - 1, dd, 0, 0, 0, 0);

  const hhmmss =
    digits.length <= 4
      ? digits.padStart(4, "0").slice(-4) + "00"
      : digits.padStart(6, "0").substring(0, 6);

  const HH = parseInt(hhmmss.substring(0, 2), 10);
  const MIN = parseInt(hhmmss.substring(2, 4), 10);
  const SS = parseInt(hhmmss.substring(4, 6), 10);

  // Already-normalized Central Time from Combined / TosTop / TosTrades.
  // Do not add tosEtToCtOffsetMs here.
  const dt = new Date(yyyy, mm - 1, dd, HH, MIN, SS, 0);
  return isNaN(dt.getTime()) ? null : dt;
}

/**
 * formatExpForSchwabImport
 *
 * Display-only. Turns the many TOS / Sheets Exp spellings into one
 * plain-text value for the Schwab Import Exp column:
 *   29-Aug-23
 *
 * Does not rewrite Symbol. Does not change Combined / TosTrades.
 * Unparseable values are returned trimmed as-is (never blanked).
 */
function formatExpForSchwabImport(raw) {
  const parts = parseExpYearMonthDay_(raw);
  if (!parts) return toStr(raw).trim();

  const mon = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  return (
    String(parts.day) +
    "-" +
    mon[parts.month - 1] +
    "-" +
    String(parts.year).slice(-2)
  );
}

function parseExpYearMonthDay_(raw) {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return {
      year: raw.getFullYear(),
      month: raw.getMonth() + 1,
      day: raw.getDate(),
    };
  }

  const s = toStr(raw).trim();
  if (!s) return null;

  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    return {
      year: parseInt(s.substring(0, 4), 10),
      month: parseInt(s.substring(5, 7), 10),
      day: parseInt(s.substring(8, 10), 10),
    };
  }

  const monNum = {
    JAN: 1,
    JANUARY: 1,
    FEB: 2,
    FEBRUARY: 2,
    MAR: 3,
    MARCH: 3,
    APR: 4,
    APRIL: 4,
    MAY: 5,
    JUN: 6,
    JUNE: 6,
    JUL: 7,
    JULY: 7,
    AUG: 8,
    AUGUST: 8,
    SEP: 9,
    SEPTEMBER: 9,
    OCT: 10,
    OCTOBER: 10,
    NOV: 11,
    NOVEMBER: 11,
    DEC: 12,
    DECEMBER: 12,
  };

  function ymd(yearRaw, monthNum, dayRaw) {
    let year = parseInt(yearRaw, 10);
    const month = Number(monthNum);
    const day = parseInt(dayRaw, 10);
    if (year < 100) year += 2000;
    if (!year || month < 1 || month > 12 || day < 1 || day > 31) return null;
    return { year: year, month: month, day: day };
  }

  // 8/29/2023  or  8/29/23  (expired rows)
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) return ymd(m[3], parseInt(m[1], 10), m[2]);

  const u = s.toUpperCase();

  // 29 Aug 23 / 29-Aug-23 / 29 August 23 / 29-August-23
  m = u.match(/^(\d{1,2})[-\s\/]+([A-Z]{3,9})[-\s\/]+(\d{2,4})$/);
  if (m && monNum[m[2]]) return ymd(m[3], monNum[m[2]], m[1]);

  // 29Aug23 / 8September23 (no separator)
  m = u.match(/^(\d{1,2})([A-Z]{3,9})(\d{2,4})$/);
  if (m && monNum[m[2]]) return ymd(m[3], monNum[m[2]], m[1]);

  const d = new Date(s);
  if (d instanceof Date && !isNaN(d.getTime())) {
    return {
      year: d.getFullYear(),
      month: d.getMonth() + 1,
      day: d.getDate(),
    };
  }
  return null;
}