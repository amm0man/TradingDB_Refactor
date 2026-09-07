/**
 * BuildUnifiedOutput.js
 *
 * Phase 1 – Group I: output fee grouping + sort key
 *
 * Sole responsibility:
 *   - applyFeesToTopLegRuleB: move MiscFees / FeesComm / Amount onto one
 *     deterministic “top leg” inside a multi-leg trade group
 *   - legOrderKeyForOutput: stable sort key so multi-leg rows stay adjacent
 *
 * These do not change matching. They only tidy rows already built.
 *
 * Why a factory (createOutputHelpers):
 *   Both helpers still use the nested toStr() from buildUnifiedImportV3().
 *   Passing it in keeps that behavior identical. toNum() is already shared.
 *
 * Called by:
 *   buildUnifiedImportV3() before applyFeesToTopLegRuleB(unifiedTrades)
 *   and before the section 8 sort.
 *
 * Shared helpers used from Helpers.js (already global):
 *   toNum
 *
 * Still passed in from buildUnifiedImportV3:
 *   toStr
 *
 * Related files:
 *   - BuildUnifiedImportV3.js          (orchestrator)
 *   - BuildUnifiedSpecialParsers.js    (Group H)
 */

/**
 * createOutputHelpers(opts)
 *
 * opts.toStr – nested string helper from buildUnifiedImportV3
 *
 * Returns the Group I function names that used to be nested in
 * buildUnifiedImportV3. Call-site argument lists are unchanged.
 */
function createOutputHelpers(opts) {
  const toStr = opts.toStr;

  // Rule B: group fees to “top leg” for multi-leg spreads (your original logic)
  function applyFeesToTopLegRuleB(unifiedTradesList) {
    const groups = {};

    for (let i = 0; i < unifiedTradesList.length; i++) {
      const r = unifiedTradesList[i];
      const spread = toStr(r.Spread).trim().toUpperCase();
      if (spread === "STOCK" || spread === "SINGLE" || spread === "")
        continue;

      const k = toStr(r.tradeGroupKey);
      if (!k) continue;

      if (!groups[k]) groups[k] = [];
      groups[k].push(r);
    }

    function legRankB(r) {
      const side = toStr(r.Side).trim().toUpperCase();
      const optType = toStr(r.optType).trim().toUpperCase();
      const strikeNum = toNum(r.Strike);

      const sideRank = side === "SELL" ? 0 : side === "BUY" ? 1 : 2;
      const typeRank = optType === "CALL" ? 0 : optType === "PUT" ? 1 : 2;

      let strikeSort = 0;
      if (!isNaN(strikeNum)) {
        if (optType === "CALL") strikeSort = -strikeNum;
        else if (optType === "PUT") strikeSort = strikeNum;
        else strikeSort = strikeNum;
      }

      const desc = toStr(r.Description);
      return [sideRank, typeRank, strikeSort, desc];
    }

    function addMaybe(sum, v) {
      if (v === null || v === undefined || toStr(v).trim() === "") return sum;
      const n = toNum(v);
      if (isNaN(n)) return sum;
      return sum + n;
    }

    const keys = Object.keys(groups);
    for (let g = 0; g < keys.length; g++) {
      const k = keys[g];
      const rows = groups[k];

      let sumMiscFees = 0;
      let sumFeesComm = 0;
      let sumAmount = 0;

      let sawAnyMiscFees = false;
      let sawAnyFeesComm = false;
      let sawAnyAmount = false;

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        if (toStr(r.MiscFees).trim() !== "") sawAnyMiscFees = true;
        if (toStr(r.FeesComm).trim() !== "") sawAnyFeesComm = true;
        if (toStr(r.Amount).trim() !== "") sawAnyAmount = true;

        sumMiscFees = addMaybe(sumMiscFees, r.MiscFees);
        sumFeesComm = addMaybe(sumFeesComm, r.FeesComm);
        sumAmount = addMaybe(sumAmount, r.Amount);
      }

      if (!sawAnyMiscFees && !sawAnyFeesComm && !sawAnyAmount) continue;

      // Clear all legs first
      for (let i = 0; i < rows.length; i++) {
        rows[i].MiscFees = "";
        rows[i].FeesComm = "";
        rows[i].Amount = "";
      }

      // Deterministic "top leg"
      rows.sort((a, b) => {
        const ra = legRankB(a);
        const rb = legRankB(b);
        for (let i = 0; i < ra.length; i++) {
          if (ra[i] < rb[i]) return -1;
          if (ra[i] > rb[i]) return 1;
        }
        return 0;
      });

      // Assign sums to top leg only
      rows[0].MiscFees = sawAnyMiscFees ? sumMiscFees : "";
      rows[0].FeesComm = sawAnyFeesComm ? sumFeesComm : "";
      rows[0].Amount = sawAnyAmount ? sumAmount : "";
    }
  }

  // output ordering inside a multi-leg trade group so legs stay adjacent.
  // This does NOT change amounts/fees; it only stabilizes row order for downstream logic.
  function legOrderKeyForOutput(r) {
    const spread = toStr(r.Spread).trim().toUpperCase();
    const side = toStr(r.Side).trim().toUpperCase();
    const optType = toStr(r.optType).trim().toUpperCase(); // CALL/PUT/STOCK (we set this on trade rows)
    const strikeNum = toNum(r.Strike);

    // Put CALL legs together, then PUT legs together (helps IC readability + adjacency).
    const typeRank = optType === "CALL" ? 0 : optType === "PUT" ? 1 : 2;

    // For calls: higher strike first (418 then 417). For puts: lower strike first (404 then 405).
    let strikeSort = 0;
    if (!isNaN(strikeNum)) {
      if (optType === "CALL") strikeSort = -strikeNum;
      else if (optType === "PUT") strikeSort = strikeNum;
      else strikeSort = strikeNum;
    }

    // Tie-breaker: SELL before BUY (stable, but happens after strike ordering)
    const sideRank = side === "SELL" ? 0 : side === "BUY" ? 1 : 2;

    return [typeRank, strikeSort, sideRank, toStr(r.Description)];
  }

  return {
    applyFeesToTopLegRuleB: applyFeesToTopLegRuleB,
    legOrderKeyForOutput: legOrderKeyForOutput,
  };
}
