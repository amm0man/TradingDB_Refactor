/**
 * BuildUnifiedExerciseAction.js
 *
 * Phase 1 – Group F: exercise / assign + action + unified symbol
 *
 * Sole responsibility:
 *   Normalize TosTrades Spread labels (EXERCISE / ASSIGN / everything else),
 *   decide Buy/Sell vs Buy to Open / Sell to Close, compute a signed stock
 *   amount from side + qty + price, and format the Schwab-style option
 *   symbol string.
 *
 * Also owns:
 *   OPTIONS_CONTRACT_MULTIPLIER = 100
 *   (You confirmed contracts are always 100 shares. If that ever changes,
 *   this constant must become symbol-specific.)
 *
 * Why a factory (createExerciseActionHelpers):
 *   These helpers still use the nested toStr() from buildUnifiedImportV3().
 *   Passing it in keeps that behavior identical. toNum() is already shared
 *   from Helpers.js.
 *
 * Called by:
 *   buildUnifiedImportV3() as soon as toStr exists.
 *
 * Shared helpers used from Helpers.js (already global):
 *   toNum
 *
 * Still passed in from buildUnifiedImportV3 (not moved yet):
 *   toStr
 *
 * Related files:
 *   - BuildUnifiedImportV3.js      (orchestrator; uses OPTIONS_CONTRACT_MULTIPLIER)
 *   - BuildUnifiedSymbols.js       (Group E)
 */

// You confirmed contracts are always 100 shares.
// If you ever trade non-100 multipliers, this must become symbol-specific.
const OPTIONS_CONTRACT_MULTIPLIER = 100;

/**
 * createExerciseActionHelpers(opts)
 *
 * opts.toStr – nested string helper from buildUnifiedImportV3
 *
 * Returns the Group F function names that used to be nested in
 * buildUnifiedImportV3. Call-site argument lists are unchanged.
 */
function createExerciseActionHelpers(opts) {
  const toStr = opts.toStr;

  // Normalize Spread strings so we have stable tags in Schwab Import.
  // Examples:
  // - "EXERCISE STOCK" -> "EXERCISE"
  // - "ASSIGN STOCK"   -> "ASSIGN"
  // - "EXERCISE"       -> "EXERCISE"
  // - "SINGLE"         -> "SINGLE"
  function normalizeSpread(spreadRaw) {
    const s = toStr(spreadRaw).trim().toUpperCase();
    if (!s) return "";
    if (s.indexOf("EXERCISE") === 0) return "EXERCISE";
    if (s.indexOf("ASSIGN") === 0 || s.indexOf("ASSIGNMENT") === 0)
      return "ASSIGN";
    return s;
  }

  function isExerciseOrAssignSpread(spreadNorm) {
    const s = toStr(spreadNorm).trim().toUpperCase();
    return s === "EXERCISE" || s === "ASSIGN";
  }

  function computeSignedAmountFromTrade(side, qtyAbs, price) {
    const s = toStr(side).trim().toUpperCase();
    const q = toNum(qtyAbs);
    const p = toNum(price);
    if (isNaN(q) || isNaN(p)) return "";
    const gross = q * p;
    if (s === "BUY") return -gross;
    if (s === "SELL") return gross;
    return "";
  }

  function actionFromTosTrades(side, posEffect, type) {
    const s = toStr(side).trim().toUpperCase();
    const p = toStr(posEffect).trim().toUpperCase();
    const t = toStr(type).trim().toUpperCase();

    const isOption = t === "CALL" || t === "PUT";

    if (isOption) {
      if (s === "BUY" && p.indexOf("OPEN") >= 0) return "Buy to Open";
      if (s === "BUY" && p.indexOf("CLOSE") >= 0) return "Buy to Close";
      if (s === "SELL" && p.indexOf("OPEN") >= 0) return "Sell to Open";
      if (s === "SELL" && p.indexOf("CLOSE") >= 0) return "Sell to Close";
      if (s === "BUY") return "Buy";
      if (s === "SELL") return "Sell";
      return "";
    }

    if (s === "BUY") return "Buy";
    if (s === "SELL") return "Sell";
    return "";
  }

  function formatUnifiedSymbol(sym, type, exp, strike) {
    const s = toStr(sym).trim().toUpperCase();
    const t = toStr(type).trim().toUpperCase();

    if (t !== "CALL" && t !== "PUT") return s;

    let expStr = "";
    if (exp instanceof Date) {
      expStr = Utilities.formatDate(
        exp,
        Session.getScriptTimeZone(),
        "MMddyyyy",
      );
    } else {
      const e = toStr(exp).trim();
      const us = e.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
      if (us) {
        const mm = String(parseInt(us[1], 10)).padStart(2, "0");
        const dd = String(parseInt(us[2], 10)).padStart(2, "0");
        let yyyy = parseInt(us[3], 10);
        if (yyyy < 100) yyyy += 2000;
        expStr = mm + dd + yyyy;
      } else {
        expStr = e.replace(/\s+/g, "");
      }
    }

    const strikeNum = toNum(strike);
    const strikeStr = isNaN(strikeNum)
      ? toStr(strike).trim()
      : strikeNum.toFixed(2);
    const cp = t === "CALL" ? "C" : "P";

    return (s + " " + expStr + " " + strikeStr + " " + cp).trim();
  }

  return {
    normalizeSpread: normalizeSpread,
    isExerciseOrAssignSpread: isExerciseOrAssignSpread,
    computeSignedAmountFromTrade: computeSignedAmountFromTrade,
    actionFromTosTrades: actionFromTosTrades,
    formatUnifiedSymbol: formatUnifiedSymbol,
  };
}
