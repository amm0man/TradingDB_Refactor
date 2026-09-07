/**
 * BuildUnifiedIcRetag.js
 *
 * Phase 1 – Group B: Iron Condor (IC) retagging helpers
 *
 * Sole responsibility:
 *   Decide when a CLOSE-like TosTrades bundle that was labeled SINGLE / VERTICAL
 *   (or blank) should inherit the spread name "IRON CONDOR" from a prior OPEN
 *   IC that still has inventory on those legs.
 *
 * Why a factory (createIcRetagHelpers):
 *   decideIcRetag() reads and updates in-memory maps that buildUnifiedImportV3()
 *   owns in section 6A:
 *     - canonicalSpreadByLifecycleBundleKey
 *     - lifecycleBundleLegSetByKey
 *     - openIcQtyByLegKey
 *     - icRetagMetrics (two counters)
 *   Passing those objects in keeps the logic identical without making the maps
 *   global.
 *
 * Called by:
 *   buildUnifiedImportV3() after the 6A maps exist, before the TosTrades
 *   pre-pass and the 6B inference loops.
 *
 * Shared helpers used from Helpers.js (already global):
 *   normalizeDate, normalizeTimeHHmmss
 *
 * Still passed in from buildUnifiedImportV3 (not moved yet):
 *   toStr, roundTo, CANONICAL_IC
 *
 * Related files:
 *   - BuildUnifiedImportV3.js      (orchestrator; owns the maps and 6B loops)
 *   - BuildUnifiedEnrichment.js    (Group A fee matching)
 */

/**
 * createIcRetagHelpers(opts)
 *
 * opts.canonicalSpreadByLifecycleBundleKey – bundleKey -> "IRON CONDOR" when inherited
 * opts.lifecycleBundleLegSetByKey          – bundleKey -> { legId: true }
 * opts.openIcQtyByLegKey                   – Account|Underlying|ExpKey|LegId -> open qty
 * opts.icRetagMetrics                      – { spreadRetaggedRowsCount, spreadRetagSkippedNoOpenIcCount }
 * opts.CANONICAL_IC                        – the string "IRON CONDOR"
 * opts.toStr                               – nested string helper from buildUnifiedImportV3
 * opts.roundTo                             – nested numeric round helper from buildUnifiedImportV3
 *
 * Returns the Group B function names that used to be nested in
 * buildUnifiedImportV3. Call-site argument lists are unchanged.
 */
function createIcRetagHelpers(opts) {
  const canonicalSpreadByLifecycleBundleKey =
    opts.canonicalSpreadByLifecycleBundleKey;
  const lifecycleBundleLegSetByKey = opts.lifecycleBundleLegSetByKey;
  const openIcQtyByLegKey = opts.openIcQtyByLegKey;
  const icRetagMetrics = opts.icRetagMetrics;
  const CANONICAL_IC = opts.CANONICAL_IC;
  const toStr = opts.toStr;
  const roundTo = opts.roundTo;

  /**
   * Called once per TosTrades row in the main trade loop.
   * Returns either the original spread label or "IRON CONDOR".
   */
  function decideIcRetag(
    Account,
    ts,
    symForMatch,
    expKey,
    posEffect,
    spreadOriginal,
    lifeBundleKey,
  ) {
    const wantsIc =
      canonicalSpreadByLifecycleBundleKey[lifeBundleKey] === CANONICAL_IC;

    if (!wantsIc) return spreadOriginal; // no change

    if (!isRetagCandidateOriginalSpread(spreadOriginal))
      return spreadOriginal;
    if (!isCloseLikeBundle(posEffect, spreadOriginal)) return spreadOriginal;

    const legsSorted = setToSortedArray(
      lifecycleBundleLegSetByKey[lifeBundleKey] || {},
    );
    if (!legsSorted.length) return spreadOriginal;

    let hasOpenIcEvidence = false;
    for (let j = 0; j < legsSorted.length; j++) {
      const legId = legsSorted[j];
      const k = makeIcLegQtyKey(Account, symForMatch, expKey, legId);
      if (Number(openIcQtyByLegKey[k] || 0) > 0) {
        hasOpenIcEvidence = true;
        break;
      }
    }

    if (hasOpenIcEvidence) {
      icRetagMetrics.spreadRetaggedRowsCount++;
      return CANONICAL_IC;
    } else {
      icRetagMetrics.spreadRetagSkippedNoOpenIcCount++;
      return spreadOriginal; // keep original
    }
  }

  function isIronCondorSpread(spreadRaw) {
    return (
      String(spreadRaw || "")
        .trim()
        .toUpperCase() === CANONICAL_IC
    );
  }

  // Only retag these “generic” labels (prevents retagging BUTTERFLY/CALENDAR/DIAGONAL/etc).
  function isRetagCandidateOriginalSpread(spreadRaw) {
    const s = String(spreadRaw || "")
      .trim()
      .toUpperCase();
    return s === "" || s === "SINGLE" || s === "VERTICAL";
  }

  // “Close-like” = anything that reduces/removes the option position.
  function isCloseLikeBundle(posEffectRaw, spreadRaw) {
    const pe = String(posEffectRaw || "")
      .trim()
      .toUpperCase();
    const sp = String(spreadRaw || "")
      .trim()
      .toUpperCase();

    if (pe.includes("CLOSE")) return true;
    if (pe.includes("ASSIGN") || pe.includes("EXERCISE")) return true;

    // Fallback if TOS encodes this in Spread instead of Pos Effect.
    if (sp === "ASSIGN" || sp === "EXERCISE") return true;

    return false;
  }

  // Bundle key = one execution “bundle” (same timestamp/account/symbol/exp/posEffect).
  // We intentionally include Exp so we don’t accidentally tie a CLOSE to the wrong expiry.
  function makeLifecycleBundleKey(
    Account,
    ts,
    symForMatch,
    expKey,
    posEffect,
  ) {
    const acc = String(Account || "")
      .trim()
      .toUpperCase();
    const sym = String(symForMatch || "")
      .trim()
      .toUpperCase();
    const pe = String(posEffect || "")
      .trim()
      .toUpperCase();
    const d = normalizeDate(ts);
    const t = normalizeTimeHHmmss(ts);
    const e = String(expKey || "").trim();
    return [acc, d, t, sym, e, pe].join("|");
  }

  function makeOpenIndexKey(Account, symForMatch, expKey) {
    const acc = String(Account || "")
      .trim()
      .toUpperCase();
    const sym = String(symForMatch || "")
      .trim()
      .toUpperCase();
    const e = String(expKey || "").trim();
    return [acc, sym, e].join("|");
  }

  function legIdFromTypeStrike(typeKey, strikeNum) {
    const t = String(typeKey || "")
      .trim()
      .toUpperCase(); // CALL/PUT
    const k = roundTo(strikeNum, 4);
    return [t, String(k)].join(":");
  }

  function setToSortedArray(setObj) {
    return Object.keys(setObj || {}).sort();
  }

  function isSubset(sortedNeedles, haystackSetObj) {
    for (let i = 0; i < sortedNeedles.length; i++) {
      if (!haystackSetObj[sortedNeedles[i]]) return false;
    }
    return true;
  }

  // key: Account|Underlying|ExpKey|LegId
  function makeIcLegQtyKey(Account, symForMatch, expKey, legId) {
    return [
      String(Account || "")
        .trim()
        .toUpperCase(),
      String(symForMatch || "")
        .trim()
        .toUpperCase(),
      String(expKey || "").trim(),
      String(legId || "")
        .trim()
        .toUpperCase(),
    ].join("|");
  }

  // Helper: normalize exp into a stable string used in keys
  function normalizeExpKey(expRaw) {
    if (expRaw instanceof Date)
      return Utilities.formatDate(
        expRaw,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd",
      );
    return toStr(expRaw).trim();
  }

  // Some spreads (like IRON CONDOR) contain BOTH CALL and PUT legs but have ONE TosTop TRD row.
  function isMultiTypeSpread(spreadU) {
    const s = toStr(spreadU).trim().toUpperCase();
    return s === "IRON CONDOR" || s === "STRANGLE" || s === "STRADDLE";
  }

  return {
    decideIcRetag: decideIcRetag,
    isIronCondorSpread: isIronCondorSpread,
    isRetagCandidateOriginalSpread: isRetagCandidateOriginalSpread,
    isCloseLikeBundle: isCloseLikeBundle,
    makeLifecycleBundleKey: makeLifecycleBundleKey,
    makeOpenIndexKey: makeOpenIndexKey,
    legIdFromTypeStrike: legIdFromTypeStrike,
    setToSortedArray: setToSortedArray,
    isSubset: isSubset,
    makeIcLegQtyKey: makeIcLegQtyKey,
    normalizeExpKey: normalizeExpKey,
    isMultiTypeSpread: isMultiTypeSpread,
  };
}
