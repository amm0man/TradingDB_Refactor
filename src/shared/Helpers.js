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