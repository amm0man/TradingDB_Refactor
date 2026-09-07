/**
 * BuildUnifiedSpecialParsers.js
 *
 * Phase 1 – Group H: special-case TosTop description parsers
 *
 * Sole responsibility:
 *   Parse TosTop DESCRIPTION text for ordinary TRD rows, DRIP reinvestments,
 *   TDA-era fractional sells, RAD splits, and option-expiration removals.
 *
 * Why a factory (createSpecialParsers):
 *   These helpers still use the nested toStr() from buildUnifiedImportV3().
 *   Passing it in keeps that behavior identical. toNum() and looksLikeCusip()
 *   are already shared from Helpers.js.
 *
 * Called by:
 *   buildUnifiedImportV3() as soon as toStr exists (before section 5 queues).
 *
 * Shared helpers used from Helpers.js (already global):
 *   toNum, looksLikeCusip
 *
 * Still passed in from buildUnifiedImportV3:
 *   toStr
 *
 * Related files:
 *   - BuildUnifiedImportV3.js         (orchestrator; still owns emit loops)
 *   - BuildUnifiedMappingSheets.js    (Group G)
 */

/**
 * createSpecialParsers(opts)
 *
 * opts.toStr – nested string helper from buildUnifiedImportV3
 *
 * Returns the Group H function names that used to be nested in
 * buildUnifiedImportV3. Call-site argument lists are unchanged.
 */
function createSpecialParsers(opts) {
  const toStr = opts.toStr;

  function parseTosTopTradeDescription(descRaw) {
    const d = toStr(descRaw).trim();
    if (!d) return { symbol: null, absQty: null, price: null };

    const u = d.toUpperCase();

    const STRATEGYWORDS = {
      VERTICAL: true,
      SINGLE: true,
      DIAGONAL: true,
      CALENDAR: true,
      STRANGLE: true,
      STRADDLE: true,
      BUTTERFLY: true,
      CONDOR: true,
      IRON: true,
      IC: true,
      SPREAD: true,
      EXERCISE: true,
      ASSIGN: true,
    };

    // Accept both TOS-style abbreviations and full words.
    // Examples:
    // - "BOT 100 ATUSF @ 12.165"
    // - "BOUGHT 100 ATUSF @ 12.165"
    // - "SOLD -2 DNN 100 16 JAN 26 2.5 CALL .34 PHLX"
    // const qtyMatch = u.match(/\b(BOT|BOUGHT|SOLD)\b\s+([+-]?[\d,\.]+)/);
    // Accept:
    // - "BOT 2 XYZ @1.23"
    // - "BOT +2 148929102 @40.82"   (CUSIP-like)
    // - "SOLD -2 XYZ @.52"
    // const qtyMatch = u.match(/\b(BOT|BOUGHT|SOLD)\b\s+([+-]?\d[\d,]*)\b/);
    // Accept:
    // - "BOT 2 XYZ @1.23"
    // - "BOT +2 148929102 @40.82"
    // - "BOT+2 148929102 @40.82"
    // - "SOLD -1 148929102 @44.90"
    // - "SOLD-1 148929102 @44.90"
    // Allow integer or decimal quantities (e.g. -100, +2, -0.0104)
    const qtyMatch = u.match(
      /\b(BOT|BOUGHT|SOLD)\b\s*([+-]?(?:\d[\d,]*)(?:\.\d+)?)\b/,
    );

    if (!qtyMatch) return { symbol: null, absQty: null, price: null };

    const qtyRaw = qtyMatch[2];
    const qtyNum = parseFloat(String(qtyRaw).replace(/,/g, ""));
    const absQty = isNaN(qtyNum) ? null : Math.abs(qtyNum);

    // Grab text after the action + qty, then scan tokens until we hit a real ticker
    const after = u.substring(qtyMatch.index + qtyMatch[0].length).trim();

    // Pull “tokens” that resemble TOS words/symbols (keeps dots/dashes/colon)
    const tokens = after.match(/[0-9A-Z.\-:]{1,25}/g) || [];

    let symbol = null;

    for (let i = 0; i < tokens.length; i++) {
      const tok = String(tokens[i] || "").trim();
      if (!tok) continue;

      // Skip strategy words like IRON, CONDOR, VERTICAL, SPREAD, etc.
      if (STRATEGYWORDS[tok]) continue;

      // Skip pure numbers (like 100, 17, 355, etc.) BUT allow 9-char CUSIP-like tokens.
      // This lets descriptions like "SOLD -100 090628207 @1.665" parse successfully,
      // and later logic will map the CUSIP to a ticker via CusipMap.
      const isPureNumber = /^\d+(\.\d+)?$/.test(tok);
      const isCusipCandidate = looksLikeCusip(tok);

      if (isPureNumber && !isCusipCandidate) continue;

      symbol = tok;
      break;
    }

    if (!symbol) return { symbol: null, absQty, price: null };

    // Price extraction:
    // Prefer "@ 12.34" or "AT 12.34" including leading-decimal forms like "@.29" or "@-.34".
    function extractPrice(strUpper) {
      // 1) Strong signal: "@.29" or "AT .29" (supports +/-, leading decimal, and normal decimals)
      let m = strUpper.match(/(?:@|AT)\s*\$?([+-]?(?:\d+(?:\.\d+)?|\.\d+))/);
      if (m) return Math.abs(parseFloat(m[1]));

      // 2) Fallback: last numeric token, but ALSO support ".29" tokens
      const all = strUpper.match(/[+-]?(?:\d+(?:\.\d+)?|\.\d+)/g);
      if (all && all.length) return Math.abs(parseFloat(all[all.length - 1]));

      return null;
    }

    const price = extractPrice(u);

    return {
      symbol: symbol || null,
      absQty,
      price,
    };
  }

  // --- UPDATED HELPER: detect DRIP TRD reinvestment rows (Schwab-era AND TDA-era) ---
  // Schwab-era DRIP: "BOT 0.0175 XOM UPON EXXON MOBIL CORP"    — detected by UPON keyword
  // TDA-era DRIP:    "BOT +0.13 XOM @101.49692"                 — detected by fractional qty
  //   (TDA did not emit a DOI row; just the reinvestment TRD itself)
  //
  // WHY the UPON exclusion list matters:
  //   TDA used the word UPON in trade-correction rows (not DRIPs), e.g.:
  //     "SOLD -2.0 AMD ... UPON Trade Correction"
  //     "BOT 2.0 AMD ... UPON Buy Trade"
  //     "SOLD -1.0 AMD ... UPON Sell Trade"
  //   These are broker adjustments, not DRIP reinvestments. They have matching
  //   TosTrades rows and must flow through normal enrichment — never the DRIP emitter.
  //
  // WHY the isBuySide guard on Signal 2 is critical:
  //   A SOLD fractional row (e.g. "SOLD -0.13 XOM @98.67") is a LEGITIMATE STC of
  //   DRIP-accumulated shares. It has a matching TosTrades SELL row and must flow
  //   through the normal enrichment path. Without the guard, Signal 2 fires on any
  //   fractional qty (including sells), causing:
  //     (a) the STC to be dropped from the enrichment queue
  //     (b) a spurious DRIP BTO row to be emitted in its place.
  //
  // @param {string} descRaw  The raw DESCRIPTION cell value.
  // @param {number} absQty   The absolute quantity parsed from the description (may be null).
  // @returns {boolean}
  function isDripTrdDescription(descRaw, absQty) {
    const descUpper = String(descRaw).trim().toUpperCase();

    // Signal 1 — Schwab-era UPON keyword.
    // Must contain UPON but NOT be a TDA broker-correction phrase.
    // TDA correction rows always end with one of these fixed suffixes after UPON.
    if (/upon/i.test(descUpper)) {
      const TDA_CORRECTION_SUFFIXES = [
        "UPON TRADE CORRECTION",
        "UPON BUY TRADE",
        "UPON SELL TRADE",
      ];
      const isTdaCorrection = TDA_CORRECTION_SUFFIXES.some((suffix) =>
        descUpper.includes(suffix),
      );
      if (!isTdaCorrection) return true; // real Schwab DRIP reinvestment
      // Falls through to Signal 2 check if it somehow matches (it won't, but safe)
    }

    // Signal 2 — TDA-era fractional qty. MUST be a buy-side row (BOT / BOUGHT).
    // SOLD fractional rows are STC events, not DRIPs.
    const isBuySide =
      descUpper.startsWith("BOT") || descUpper.startsWith("BOUGHT");
    if (
      isBuySide &&
      absQty != null &&
      absQty !== undefined &&
      !isNaN(absQty) &&
      absQty > 0
    ) {
      if (absQty % 1 !== 0) return true; // non-integer → fractional share buy → DRIP
    }

    return false;
  }
  // --- HELPER detect TDA-era fractional SELL TRD rows with no TosTrades partner ---
  // TDA never emitted a TosTrades row for fractional share sells.
  // Example: SOLD -0.13 XOM @98.67
  // These are STC events for DRIP-accumulated partial shares that must be emitted
  // directly from TosTop — they have no TosTrades enrichment partner.
  //
  // IMPORTANT MIGRATION CUTOFF: After Schwab migration (DT: 2024-05-12, LT: 2024-05-12),
  // Schwab DOES log a matching TosTrades row for fractional sells. So we must NOT fire
  // this emitter for post-migration rows — they will be enriched normally via TosTrades.
  //
  // Detection rule — ALL FOUR must be true:
  //   1. Description starts with "SOLD" or "SOLD-"  (sell-side, TDA format)
  //   2. absQty is fractional  (absQty % 1 !== 0)
  //   3. The amount is POSITIVE (credit / proceeds from a sell)
  //   4. The row date is BEFORE the Schwab migration cutoff for that account
  //
  // @param {string}  descRaw  The raw DESCRIPTION cell value.
  // @param {number}  absQty   The absolute quantity parsed from the description.
  // @param {number}  amount   The raw AMOUNT field from TosTop (positive = credit).
  // @param {string}  dateIso  The normalized ISO date of the row (yyyy-MM-dd).
  // @param {string}  account  The account code ('DT' or 'LT').
  // @returns {boolean}
  function isTdaFractionalSellTrd(descRaw, absQty, amount, dateIso, account) {
    const descUpper = String(descRaw).trim().toUpperCase();
    const isSellSide = descUpper.startsWith("SOLD");
    if (!isSellSide) return false;
    if (
      absQty == null ||
      absQty === undefined ||
      isNaN(absQty) ||
      absQty <= 0
    )
      return false;
    if (absQty % 1 === 0) return false; // whole-share sells go through the normal path
    const amt = toNum(amount);
    if (isNaN(amt) || amt <= 0) return false; // must be a credit (positive amount)

    // Migration cutoff: both DT and LT migrated to Schwab on 2024-05-12.
    // Any fractional SELL on or after this date will have a real TosTrades partner.
    const SCHWAB_MIGRATION_CUTOFF = "2024-05-12";
    if (dateIso >= SCHWAB_MIGRATION_CUTOFF) return false; // post-migration: has TosTrades partner

    return true;
  }

  // --- HELPER detect RAD mandatory split rows ---
  // Handles all known TosTop split description formats.
  // Returns:
  //   isSplit
  //   parsedQty        -> absolute numeric qty found in the RAD line
  //   rawQty           -> signed qty as written in the RAD line
  //   isReverse
  //   looksLikePreLeg  -> removal / old-share leg
  //   looksLikePostLeg -> resulting / new-share leg
  function parseRadSplitDescription(descRaw) {
    const u = String(descRaw ?? "")
      .trim()
      .toUpperCase();

    if (!u.includes("SPLIT")) return { isSplit: false };

    const isReverse =
      u.includes("REVERSE") ||
      (!u.includes("FORWARD") && !u.startsWith("STOCK SPLIT"));

    const m = u.match(/([-+]?\d+(?:\.\d+)?)/);
    if (!m) return { isSplit: false };

    const rawQty = Number(m[1]);
    const parsedQty = Math.abs(rawQty);

    if (!isFinite(parsedQty) || parsedQty <= 0) return { isSplit: false };

    const looksLikePreLeg =
      rawQty < 0 ||
      /\bEFF\s*-/.test(u) ||
      /\bMANDATORY\s+REVERSE\s+SPLIT\s*-/.test(u) ||
      /\bOXXXREVERSE\b/.test(u) ||
      /\bXXXREVERSE\b/.test(u);

    const looksLikePostLeg = rawQty > 0;

    return {
      isSplit: true,
      parsedQty: parsedQty,
      rawQty: rawQty,
      isReverse: isReverse,
      looksLikePreLeg: looksLikePreLeg,
      looksLikePostLeg: looksLikePostLeg,
      rawText: u,
    };
  }

  // Keep only one split row from a same-timestamp broker pair.
  // Reverse split -> prefer post leg, then smallest qty.
  // Forward split -> prefer post leg, then largest qty.
  function chooseCanonicalSplitCandidate_(candidates) {
    if (!Array.isArray(candidates) || !candidates.length) return null;
    if (candidates.length === 1) return candidates[0];

    // IMPORTANT:
    // Prefer the explicit POST-split leg first.
    // Only after that do we prefer a candidate that matched SplitAdjustments.
    const explicitPost = candidates.filter(
      (c) => c && c.splitCheck && c.splitCheck.looksLikePostLeg,
    );
    const pool0 = explicitPost.length ? explicitPost : candidates;

    const withAdj = pool0.filter((c) => c && c.adjustmentMatched);
    const pool = withAdj.length ? withAdj : pool0;

    const isReverse = !!(
      pool[0] &&
      pool[0].splitCheck &&
      pool[0].splitCheck.isReverse
    );

    pool.sort(function (a, b) {
      const qa = Number(a && a.splitCheck ? a.splitCheck.parsedQty : 0);
      const qb = Number(b && b.splitCheck ? b.splitCheck.parsedQty : 0);
      return isReverse ? qa - qb : qb - qa;
    });

    return pool[0];
  }

  function buildCanonicalSplitRowFromCandidates_(candidates) {
    if (!Array.isArray(candidates) || !candidates.length) return null;

    const chosen = chooseCanonicalSplitCandidate_(candidates);
    if (!chosen) return null;

    const isReverse = !!(chosen.splitCheck && chosen.splitCheck.isReverse);

    function pickRoleCandidate_(role) {
      const pool = candidates.filter(function (c) {
        if (!c || !c.splitCheck) return false;
        return role === "PRE"
          ? !!c.splitCheck.looksLikePreLeg
          : !!c.splitCheck.looksLikePostLeg;
      });

      if (!pool.length) return null;

      pool.sort(function (a, b) {
        const qa = Number(a && a.splitCheck ? a.splitCheck.parsedQty : 0);
        const qb = Number(b && b.splitCheck ? b.splitCheck.parsedQty : 0);

        if (role === "PRE") return qb - qa;
        return isReverse ? qa - qb : qb - qa;
      });

      return pool[0];
    }

    const preCandidate = pickRoleCandidate_("PRE");
    const postCandidate = pickRoleCandidate_("POST");

    const adjSource =
      (chosen.adjustmentMatched && chosen.adjustment) ||
      candidates.find(function (c) {
        return c && c.adjustmentMatched && c.adjustment;
      }) ||
      null;

    const splitRatio = adjSource
      ? Number(adjSource.num) / Number(adjSource.den)
      : null;

    let preQty =
      preCandidate && preCandidate.splitCheck
        ? Number(preCandidate.splitCheck.parsedQty)
        : null;
    let postQty =
      postCandidate && postCandidate.splitCheck
        ? Number(postCandidate.splitCheck.parsedQty)
        : null;

    // Fallback math when only one leg is present.
    if (
      (preQty == null || !isFinite(preQty)) &&
      isFinite(splitRatio) &&
      splitRatio > 0 &&
      postQty != null &&
      isFinite(postQty)
    ) {
      preQty = Math.round((postQty / splitRatio) * 1e8) / 1e8;
    }

    if (
      (postQty == null || !isFinite(postQty)) &&
      isFinite(splitRatio) &&
      splitRatio > 0 &&
      preQty != null &&
      isFinite(preQty)
    ) {
      postQty = Math.round(preQty * splitRatio * 1e8) / 1e8;
    }

    const qtyDelta =
      preQty != null &&
      isFinite(preQty) &&
      postQty != null &&
      isFinite(postQty)
        ? Math.round((postQty - preQty) * 1e8) / 1e8
        : null;

    const splitTypeLabel = isReverse ? "REVERSE SPLIT" : "FORWARD SPLIT";
    const ratioLabel = adjSource
      ? `${adjSource.num}:${adjSource.den}`
      : "?:?";
    const preLabel = preQty != null && isFinite(preQty) ? preQty : "?";
    const postLabel = postQty != null && isFinite(postQty) ? postQty : "?";
    const splitDescOut = `${splitTypeLabel} ${ratioLabel} PRE=${preLabel} POST=${postLabel}`;

    const splitTs = chosen.splitTs;
    const splitDateDisplay =
      splitTs instanceof Date && !isNaN(splitTs.getTime())
        ? Utilities.formatDate(
            splitTs,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd",
          )
        : (chosen.splitDateIso ?? "");
    const splitTimeDisplay =
      splitTs instanceof Date && !isNaN(splitTs.getTime())
        ? Utilities.formatDate(
            splitTs,
            Session.getScriptTimeZone(),
            "HH:mm:ss",
          )
        : (chosen.splitTimeHHmmss ?? "");

    const rowObj = {
      Account: chosen.account,
      Date: splitDateDisplay,
      Time: splitTimeDisplay,
      Timestamp: splitTs,
      Action: "Split",
      Symbol: chosen.splitTicker || "",
      Description: splitDescOut,
      Spread: "STOCK",
      Quantity: qtyDelta != null ? qtyDelta : "",
      Price: "",
      NetPrice: "",
      Side: isReverse ? "REVERSE" : "FORWARD",
      PosEffect: "ADJUSTMENT",
      Exp: "",
      Strike: "",
      OrderType: "",
      MiscFees: "",
      FeesComm: "",
      Amount: "",
    };

    return {
      chosen: chosen,
      rowObj: rowObj,
      adjSource: adjSource,
      preQty: preQty,
      postQty: postQty,
      qtyDelta: qtyDelta,
      ratioLabel: ratioLabel,
      splitTypeLabel: splitTypeLabel,
    };
  }

  // Parses TosTop RAD removal descriptions so Schwab Import gets Symbol/Qty/Exp/Strike when possible.
  // Example: "Removal of Option due to expiration 3.0 QQQ 100 (WEEKLY) 9 Mar 2023 305.0 CALL"
  function parseRadRemovalDescription(descRaw) {
    const d = toStr(descRaw).trim();
    const u = d.toUpperCase();
    if (!u.includes("REMOVAL OF OPTION DUE TO EXPIRATION")) return null;

    // Quantity (best-effort).
    const qtyMatch = u.match(/EXPIRATION\s+(-?\d+(?:\.\d+)?)/);
    const qty = qtyMatch ? Math.abs(parseFloat(qtyMatch[1])) : null;

    // Underlying ticker: first all-caps token after qty span.
    const after = qtyMatch
      ? u.substring(qtyMatch.index + qtyMatch[0].length).trim()
      : u;
    const tok = after.match(/\b[A-Z]{1,6}\b/);
    const underlying = tok ? tok[0] : null;

    // Date: "9 Mar 2023"
    const dateMatch = u.match(
      /\b(\d{1,2})\s+(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+(\d{4})\b/,
    );
    let exp = null;
    if (dateMatch) {
      const day = parseInt(dateMatch[1], 10);
      const mon = dateMatch[2];
      const year = parseInt(dateMatch[3], 10);
      const monMap = {
        JAN: 0,
        FEB: 1,
        MAR: 2,
        APR: 3,
        MAY: 4,
        JUN: 5,
        JUL: 6,
        AUG: 7,
        SEP: 8,
        OCT: 9,
        NOV: 10,
        DEC: 11,
      };
      if (monMap.hasOwnProperty(mon)) {
        exp = new Date(year, monMap[mon], day, 0, 0, 0, 0);
      }
    }

    // Strike: last number before CALL/PUT.
    const typeMatch = u.match(/\b(CALL|PUT)\b/);
    const optType = typeMatch ? typeMatch[1] : null;
    let strike = null;
    if (optType) {
      const beforeType = u.substring(0, typeMatch.index);
      const nums = beforeType.match(/-?\d+(?:\.\d+)?/g);
      if (nums && nums.length) strike = parseFloat(nums[nums.length - 1]);
    }

    return {
      qty: qty,
      underlying: underlying,
      exp: exp,
      strike: strike,
      optType: optType,
    };
  }

  return {
    parseTosTopTradeDescription: parseTosTopTradeDescription,
    isDripTrdDescription: isDripTrdDescription,
    isTdaFractionalSellTrd: isTdaFractionalSellTrd,
    parseRadSplitDescription: parseRadSplitDescription,
    chooseCanonicalSplitCandidate_: chooseCanonicalSplitCandidate_,
    buildCanonicalSplitRowFromCandidates_: buildCanonicalSplitRowFromCandidates_,
    parseRadRemovalDescription: parseRadRemovalDescription,
  };
}
