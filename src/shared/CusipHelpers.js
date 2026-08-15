/**
 * CusipHelpers.js
 *
 * Shared CUSIP normalization and detection helpers used by Phase 1 and Phase 2.
 *
 * Why shared:
 *   Both BuildUnifiedImportV3 and mapSchwabImportByHeadersV3 need the same
 *   stable rules for turning raw sheet values into CUSIP lookup keys and for
 *   deciding whether a token "looks like" a CUSIP.
 *
 * Design notes:
 *   - normalizeCusip handles the common Sheets problem where a CUSIP arrives
 *     as a number (leading zeros and letters would otherwise be lost).
 *   - looksLikeCusip requires exactly 9 [0-9A-Z] characters AND at least one
 *     digit (pure letter strings are not treated as CUSIPs).
 *
 * Related:
 *   - BuildUnifiedImportV3.js
 *   - mapSchwabImportByHeadersV3.js (also has normalizeRenameIdentifierV3 for
 *     general ticker-like IDs that must preserve characters such as "/")
 */

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