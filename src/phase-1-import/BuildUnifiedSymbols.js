/**
 * BuildUnifiedSymbols.js
 *
 * Phase 1 – Group E: symbol and option parsing
 *
 * Sole responsibility:
 *   Clean a raw TosTrades / TosTop symbol, keep CUSIPs intact, pull the
 *   underlying out of a dotted option root, and parse OCC-like symbols
 *   such as .QQQ230309C305 into underlying + expiration + CALL/PUT + strike.
 *
 * Why a factory (createSymbolHelpers):
 *   These helpers still use the nested toStr() from buildUnifiedImportV3().
 *   Passing it in keeps that behavior identical. They also use the shared
 *   CUSIP helpers already in Helpers.js (normalizeCusip, looksLikeCusip).
 *
 * Called by:
 *   buildUnifiedImportV3() as soon as toStr exists.
 *
 * Shared helpers used from Helpers.js (already global):
 *   normalizeCusip, looksLikeCusip
 *
 * Still passed in from buildUnifiedImportV3 (not moved yet):
 *   toStr
 *
 * Related files:
 *   - BuildUnifiedImportV3.js      (orchestrator)
 *   - BuildUnifiedSheetFields.js   (Group C)
 *   - Helpers.js                   (CUSIP helpers)
 */

/**
 * createSymbolHelpers(opts)
 *
 * opts.toStr – nested string helper from buildUnifiedImportV3
 *
 * Returns the Group E function names that used to be nested in
 * buildUnifiedImportV3. Call-site argument lists are unchanged.
 */
function createSymbolHelpers(opts) {
  const toStr = opts.toStr;

  function normalizeSymbol(s) {
    const raw = toStr(s).trim().toUpperCase();
    return raw.replace(/\s+/g, "");
  }

  // Match-symbol rule:
  // - For dotted option symbols like .QQQ230309C306, use the underlying (QQQ) for matching/enrichment.
  // - If the value is CUSIP-like (9 alnum chars, must include a digit), preserve it as a CUSIP key (do not truncate).
  // AFTER — force to string BEFORE normalizeSymbol so Sheets numeric
  // coercion on CUSIP cells like 00848K101 cannot strip leading zeros or letters.
  function normalizeUnderlyingFromTradeSymbol(symRaw) {
    const s = normalizeSymbol(toStr(symRaw));
    if (!s) return "";

    //  Preserve CUSIPs as-is (normalized), so CusipMap can map them later.
    const cusipCandidate = normalizeCusip(s);
    if (looksLikeCusip(cusipCandidate) && /\d/.test(cusipCandidate)) {
      return cusipCandidate;
    }

    // Common Thinkorswim-ish format: .QQQ230309C306 or .SPY240119P450
    // Underlying = leading letters before the first digit.
    const m = s.match(/^\.?([A-Z]{1,10})/);
    if (m && m[1]) return m[1];

    // If it's just a dotted ticker like .SPX, keep the ticker without dot.
    if (s[0] === ".") return s.substring(1);

    return s;
  }

  // If TosTrades has an OCC-like dotted option symbol such as:
  //   .QQQ230309C305
  // Parse it into underlying + expiration date + CALL/PUT + strike.
  //
  // NOTE: This does NOT match plain dotted underlyings like ".SPX" (no date/type/strike).
  function parseDottedOptionSymbol(symRaw) {
    const s = normalizeSymbol(symRaw); // removes spaces, keeps dot, uppercases
    if (!s || s[0] !== ".") return null;

    // Pattern: .UNDERLYING + YYMMDD + C/P + STRIKE
    // Example: .QQQ230309C305
    const m = s.match(
      /^\.(\w{1,10})(\d{2})(\d{2})(\d{2})([CP])(\d+(?:\.\d+)?)$/,
    );
    if (!m) return null;

    const underlying = m[1].toUpperCase();
    const yy = parseInt(m[2], 10);
    const mm = parseInt(m[3], 10);
    const dd = parseInt(m[4], 10);
    const cp = m[5].toUpperCase();
    const strikeNum = parseFloat(m[6]);

    if (
      !underlying ||
      !isFinite(yy) ||
      !isFinite(mm) ||
      !isFinite(dd) ||
      !isFinite(strikeNum)
    )
      return null;
    if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return null;

    const expDate = new Date(2000 + yy, mm - 1, dd, 0, 0, 0, 0);
    const optType = cp === "C" ? "CALL" : cp === "P" ? "PUT" : "";

    if (!optType || !(expDate instanceof Date) || isNaN(expDate.getTime()))
      return null;

    return {
      underlying: underlying,
      expDate: expDate,
      optType: optType, // "CALL" or "PUT"
      strike: strikeNum,
    };
  }

  return {
    normalizeSymbol: normalizeSymbol,
    normalizeUnderlyingFromTradeSymbol: normalizeUnderlyingFromTradeSymbol,
    parseDottedOptionSymbol: parseDottedOptionSymbol,
  };
}
