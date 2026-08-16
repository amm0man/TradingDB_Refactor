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
// (Future groups will go here, e.g. NUMBER HELPERS, UI HELPERS, etc.)
// =========================================================================

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