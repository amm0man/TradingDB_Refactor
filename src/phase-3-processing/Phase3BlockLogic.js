/**
 * Phase3BlockLogic.js
 *
 * Phase 3 – Step 3: Block / Position Logic
 *
 * Sole responsibility:
 *   Read the cleaned Helper sheet → build open/close blocks →
 *   assign Trade Group IDs, Position IDs, Spread Group IDs →
 *   handle symbol changes, splits, RAD/EXP/EXERCISE →
 *   write the final result to the Staging sheet.
 *
 * Called by:
 *   - refreshAllScripts() in Phase3Processing.js
 *   - Directly from the "DB Tools" menu
 *
 * Key concepts inside this file:
 *   - blocks{} map keyed by account + instrument identity
 *   - Spread Group ID assignment (3 sub-passes)
 *   - Live resolution of closing legs via resolveLiveSpreadGroupId
 *   - Special handling for SYMBOL CHANGE and SPLIT rows
 *
 * Related files:
 *   - Phase3Processing.js          (orchestrator)
 *   - Phase3Step1_CopyMapping.js   (will be created later)
 *   - Phase3Step2_ValidateClean.js (will be created later)
 */

// ── LIVE-BLOCK-AWARE SPREAD GROUP RESOLVER ─────────────────────────────────
// WHY: Sub-pass C uses range-containment to find a Spread Group ID for closing
// legs. When two sequential spreads share a boundary strike (e.g. 300/310 PDS
// followed by 310/315 PDS), .find() always returns the first range that
// contains the strike — which is wrong once the first spread has closed flat.
//
// This function is called INSIDE the main block loop (Step 4) where `blocks`
// already holds a live running unit count. It picks the spread group whose
// block currently has unit > 0 among all candidates, resolving the ambiguity.
//
// Falls back to the first candidate (original behavior) only when no candidate
// has open units — which should not happen in a clean dataset but is safe.
//
// PARAMETERS:
//   acct           — "DT" or "LT"
//   ticker         — e.g. "QQQ"
//   expStr         — yyyy-MM-dd string
//   strat          — Strategy Type uppercase, e.g. "PDS"
//   strike         — numeric option strike
//   spreadRangeMap — the range map built in Sub-pass B
//   blocks         — the live block-state object from Step 4 (passed by reference)
// ─────────────────────────────────────────────────────────────────────────────
// =========================================================================
// HELPER: resolveLiveSpreadGroupId
//   Looks up the currently open spread-group ID for a given account/ticker/
//   expiration/strategy so closing legs can inherit the correct group.
// =========================================================================
function resolveLiveSpreadGroupId(
  acct,
  ticker,
  expStr,
  strat,
  strike,
  spreadRangeMap,
  blocks,
) {
  function candidatesFor(stratKey) {
    const rangeKey = `${acct}|${ticker}|${expStr}|${stratKey}`;
    const rangeGroups = spreadRangeMap[rangeKey] || [];
    return rangeGroups.filter((g) => strike >= g.min && strike <= g.max);
  }

  function pickStartedGroup(list) {
    for (const g of list) {
      const b = blocks[`${acct}|${g.groupId}`] || {};
      if (b.unit > 0) return g.groupId;
    }
    for (const g of list) {
      const b = blocks[`${acct}|${g.groupId}`] || {};
      if (b.positionId || b.openTs) return g.groupId;
    }
    return "";
  }

  let candidates = candidatesFor(strat);

  if (candidates.length === 0) {
    const prefixAcct = String(acct || "").toUpperCase();
    const prefixTkr = String(ticker || "").toUpperCase();
    const prefixExp = String(expStr || "");
    Object.keys(spreadRangeMap || {}).forEach(function (rangeKey) {
      const parts = String(rangeKey).split("|");
      if (parts.length < 4) return;
      if (String(parts[0]).toUpperCase() !== prefixAcct) return;
      if (String(parts[1]).toUpperCase() !== prefixTkr) return;
      if (String(parts[2]) !== prefixExp) return;
      const extra = (spreadRangeMap[rangeKey] || []).filter(
        (g) => strike >= g.min && strike <= g.max,
      );
      extra.forEach(function (g) {
        candidates.push(g);
      });
    });
  }

  if (candidates.length === 0) return "";
  return pickStartedGroup(candidates);
}

function hasContainingSpreadWindow(
  acct,
  ticker,
  expStr,
  strike,
  spreadRangeMap,
) {
  const prefixAcct = String(acct || "").toUpperCase();
  const prefixTkr = String(ticker || "").toUpperCase();
  const prefixExp = String(expStr || "");
  const keys = Object.keys(spreadRangeMap || {});
  for (let i = 0; i < keys.length; i++) {
    const parts = String(keys[i]).split("|");
    if (parts.length < 4) continue;
    if (String(parts[0]).toUpperCase() !== prefixAcct) continue;
    if (String(parts[1]).toUpperCase() !== prefixTkr) continue;
    if (String(parts[2]) !== prefixExp) continue;
    const groups = spreadRangeMap[keys[i]] || [];
    for (let j = 0; j < groups.length; j++) {
      const g = groups[j];
      if (strike >= g.min && strike <= g.max) return true;
    }
  }
  return false;
}

// ==================== UPDATED BLOCK LOGIC V3 - FIXED ====================
/**
 * populateStagingWithBlockLogicV3
 *
 * FIXED: Trade Group ID now increments once per open-to-flat block.
 * FIXED: RAD "Opt Expired" rows now properly close the block.
 * FIXED: Sequential same-ticker spreads sharing a boundary strike now resolve
 *        correctly via resolveLiveSpreadGroupId() in the main block loop.
 */
// =========================================================================
// STEP 3: populateStagingWithBlockLogicV3
//   Core Phase 3 engine.
//   Reads Helper, builds open/close blocks, assigns Trade Group IDs /
//   Position IDs / Spread Group IDs, handles symbol changes, exercises,
//   assignments, and writes the final result to the "Staging" sheet.
// =========================================================================
function populateStagingWithBlockLogicV3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const helperSheet = ss.getSheetByName("Helper");
  const stagingSheet = ss.getSheetByName("Staging");
  if (!helperSheet || !stagingSheet)
    throw new Error("Helper or Staging sheet not found!");

  // ── Staging Issues CTX ─────────────────────────────────────────────────────
  const ctx = stagingIssuesStart("populateStagingWithBlockLogicV3");
  importIssuesSetMetric(ctx, "SourceSheet", "Helper");
  importIssuesSetMetric(ctx, "DestSheet", "Staging");
  // ───────────────────────────────────────────────────────────────────────────

  const STRATEGY_ABBREV = {
    "SHORT IC": "SIC",
    "LONG IC": "LIC",
    "SHORT PCS": "PCS",
    "SHORT PDS": "PDS",
    "SHORT CCS": "CCS",
    "SHORT CDS": "CDS",
    "LONG BUTTERFLY": "LBF",
    "SHORT BUTTERFLY": "SBF",
    "LONG PUT": "LP",
    "SHORT PUT": "SP",
    "LONG CALL": "LC",
    "SHORT CALL": "SC",
    "LONG STOCK": "LST",
    "SHORT STOCK": "SST",
  };
  function getStratAbbrev(strategyType) {
    return STRATEGY_ABBREV[strategyType.trim().toUpperCase()] || "OTH";
  }

  function parseSymbolChangeFromNotes(row, colMap) {
    if (colMap["notes"] === undefined) return null;

    const notes = String(row[colMap["notes"] - 1] || "").trim();
    if (!notes) return null;

    const out = {
      fromRaw: "",
      toRaw: "",
      fromResolved: "",
      toResolved: "",
      rawText: notes,
    };

    notes.split("|").forEach((part) => {
      const seg = String(part || "").trim();
      const eq = seg.indexOf("=");
      if (eq === -1) return;

      const key = seg.substring(0, eq).trim().toUpperCase();
      const val = seg
        .substring(eq + 1)
        .trim()
        .toUpperCase();

      if (key === "FROM") out.fromRaw = val;
      if (key === "TO") out.toRaw = val;
      if (key === "FROM_RESOLVED") out.fromResolved = val;
      if (key === "TO_RESOLVED") out.toResolved = val;
    });

    return out.fromRaw || out.toRaw || out.fromResolved || out.toResolved
      ? out
      : null;
  }

  function parseSplitRatioFromRow(row, colMap) {
    const notesIdx = colMap["notes"] ? colMap["notes"] - 1 : -1;
    const text =
      notesIdx > -1
        ? String(row[notesIdx] || "")
            .trim()
            .toUpperCase()
        : "";

    if (!text || text.indexOf("SPLIT") === -1) return null;

    const ratioMatch = text.match(/\b(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)/);
    const preMatch = text.match(/\bPRE\s*=\s*([-+]?\d+(?:\.\d+)?)/);
    const postMatch = text.match(/\bPOST\s*=\s*([-+]?\d+(?:\.\d+)?)/);

    const numerator = ratioMatch ? Number(ratioMatch[1]) : NaN;
    const denominator = ratioMatch ? Number(ratioMatch[2]) : NaN;
    const preQty = preMatch ? Math.abs(Number(preMatch[1])) : NaN;
    const postQty = postMatch ? Math.abs(Number(postMatch[1])) : NaN;
    const splitType = text.includes("REVERSE")
      ? "REVERSE"
      : text.includes("FORWARD")
        ? "FORWARD"
        : "";

    if (
      (!isFinite(numerator) || !isFinite(denominator) || denominator === 0) &&
      !isFinite(preQty) &&
      !isFinite(postQty)
    ) {
      return null;
    }

    return {
      numerator: numerator,
      denominator: denominator,
      preQty: preQty,
      postQty: postQty,
      splitType: splitType,
      text: text,
    };
  }

  function computeExpectedPostSplitQty_(preQty, splitInfo) {
    if (
      !splitInfo ||
      !isFinite(preQty) ||
      !isFinite(splitInfo.numerator) ||
      !isFinite(splitInfo.denominator) ||
      splitInfo.denominator === 0
    )
      return NaN;

    const rawPost =
      Number(preQty) *
      (Number(splitInfo.numerator) / Number(splitInfo.denominator));
    const isReverse =
      splitInfo.splitType === "REVERSE" ||
      splitInfo.numerator < splitInfo.denominator;

    return isReverse
      ? Math.floor(rawPost + 1e-9)
      : Math.round(rawPost * 1e8) / 1e8;
  }

  function isMatchingSplitRowForBlock_(runningBeforeSplit, splitInfo) {
    if (!splitInfo) return true;

    if (
      isFinite(splitInfo.preQty) &&
      Math.abs(splitInfo.preQty - runningBeforeSplit) < 1e-8
    ) {
      return true;
    }

    if (isFinite(splitInfo.postQty)) {
      const expectedPost = computeExpectedPostSplitQty_(
        runningBeforeSplit,
        splitInfo,
      );
      if (
        isFinite(expectedPost) &&
        Math.abs(splitInfo.postQty - expectedPost) < 1e-8
      ) {
        return true;
      }
    }

    return false;
  }

  function findRenameSourceStockBlockForSplit_(
    acct,
    targetTicker,
    splitInfo,
    blocks,
  ) {
    if (!splitInfo || !isFinite(splitInfo.postQty)) return null;

    let match = null;

    Object.keys(blocks).forEach(function (blockKey) {
      if (match) return;

      const parts = blockKey.split("|");
      if (parts.length !== 2) return; // stock block keys are acct|ticker
      if (parts[0] !== acct) return;

      const sourceTicker = parts[1];
      if (!sourceTicker || sourceTicker === targetTicker) return;

      const block = blocks[blockKey];
      if (!block || !block.positionId) return;

      const preQty = Number(block.runningQty || 0);
      if (preQty <= 0) return;

      const expectedPost = computeExpectedPostSplitQty_(preQty, splitInfo);
      if (
        isFinite(expectedPost) &&
        Math.abs(expectedPost - splitInfo.postQty) < 1e-8
      ) {
        match = {
          sourceKey: blockKey,
          sourceTicker: sourceTicker,
          block: block,
          preQty: preQty,
          expectedPost: expectedPost,
        };
      }
    });

    return match;
  }

  function isCusipLikeSymbolValue_(val) {
    const v = String(val || "")
      .trim()
      .toUpperCase();
    return !!v && /^[A-Z0-9]{9}$/.test(v) && /\d/.test(v);
  }

  function resolveSymbolChangeTargetTicker_(
    fromTicker,
    toRaw,
    toResolved,
    rowTicker,
  ) {
    const from = String(fromTicker || "")
      .trim()
      .toUpperCase();
    const rowTkr = String(rowTicker || "")
      .trim()
      .toUpperCase();

    const candidates = [
      String(toResolved || "")
        .trim()
        .toUpperCase(),
      String(toRaw || "")
        .trim()
        .toUpperCase(),
      rowTkr,
    ];

    for (let i = 0; i < candidates.length; i++) {
      const candidate = candidates[i];
      if (!candidate) continue;

      // If the row itself already carries a real new ticker, do not let a
      // fallback candidate pin us back to the old symbol.
      if (
        candidate === from &&
        rowTkr &&
        rowTkr !== from &&
        !isCusipLikeSymbolValue_(rowTkr)
      ) {
        continue;
      }

      if (!isCusipLikeSymbolValue_(candidate)) return candidate;
    }

    return from;
  }

  function parseSymbolChangeFromRow(row, colMap) {
    const notes =
      colMap["notes"] !== undefined
        ? String(row[colMap["notes"] - 1]).trim()
        : "";
    const corpActions =
      colMap["corporate actions"] !== undefined
        ? String(row[colMap["corporate actions"] - 1])
            .trim()
            .toUpperCase()
        : "";
    const action =
      colMap["action"] !== undefined
        ? String(row[colMap["action"] - 1])
            .trim()
            .toUpperCase()
        : "";

    const isRenameRow =
      action === "SYMBOL CHANGE" ||
      corpActions === "SYMBOL CHANGE" ||
      corpActions === "Symbol Change".toUpperCase();

    if (!isRenameRow) return null;
    if (!notes) return null;

    // Expected Notes format from Phase 2:
    // FROM=ENCUF | TO=EU | FROM_RESOLVED=ENCUF | TO_RESOLVED=EU | RAW=Symbol Change from ENCUF to EU
    function pull(label) {
      const m = notes.match(new RegExp(label + "=([^|]+)", "i"));
      return m ? String(m[1]).trim().toUpperCase() : "";
    }

    const fromRaw = pull("FROM");
    const toRaw = pull("TO");
    const fromResolved = pull("FROM_RESOLVED");
    const toResolved = pull("TO_RESOLVED");

    return {
      fromRaw: fromRaw,
      toRaw: toRaw,
      fromResolved: fromResolved,
      toResolved: toResolved,
    };
  }

  // ── PRIORITY 1: try/catch/finally so Issues log ALWAYS gets flushed ────────
  try {
    const helperData = helperSheet.getDataRange().getValues();
    if (helperData.length < 4) {
      uiAlertSafe("No data in Helper to process.");
      return;
    }

    const colMap = {};
    helperData[0].forEach((h, i) => {
      if (typeof h === "string" && h.trim())
        colMap[h.trim().toLowerCase()] = i + 1;
    });

    let data = helperData.slice(3).map((row) => row.slice());

    const dataStartRow = 4;
    const tz = ss.getSpreadsheetTimeZone();

    // ── CTX counters ──────────────────────────────────────────────────────────
    let ctxStrikeCollisions = 0;
    let ctxBlocksOpened = 0;
    let ctxBlocksClosed = 0;
    let ctxRADRows = 0;
    let ctxMissingSpreadGroup = 0;
    // ─────────────────────────────────────────────────────────────────────────

    // Step 2: Sort by Account then Trade Time Stamp.
    const tsIdx = colMap["trade time stamp"] - 1;
    // WHY the three-level sort:
    // Level 1 — Account (DT before LT, keeps accounts cleanly separated).
    // Level 2 — Trade Time Stamp (canonical sequencing key for block logic).
    // Level 3 — Option Strike ascending (tie-breaker for same-timestamp rows).
    //   When two legs of a spread share the exact same timestamp (e.g. the final
    //   BUY TO CLOSE 310 and SELL TO CLOSE 315 both stamped 08/15/2023 8:36),
    //   JavaScript sort is not guaranteed stable. Without a tie-breaker, the
    //   SELL TO CLOSE can land before the BUY TO CLOSE, dropping unit to 0
    //   (blkClose fires), then the BUY TO CLOSE arrives with prevUnit=0 and
    //   fires blkStart — leaving the block permanently open with unit=-1.
    //   Sorting lower strike first means the BUY TO CLOSE 310 always processes
    //   before the SELL TO CLOSE 315, the unit walks down monotonically, and
    //   blkClose fires correctly on the final leg only.
    data.sort((a, b) => {
      const acctA = (a[colMap["account"] - 1] || "").toString().toUpperCase();
      const acctB = (b[colMap["account"] - 1] || "").toString().toUpperCase();
      if (acctA !== acctB) return acctA.localeCompare(acctB);

      // Same calendar day first. Then non-RAD before RAD.
      // WHY: TOS often stamps Opt Expired at 00:xx on expiration day
      // (DT SPX 9/7/2023 RAD 00:39 vs IC opens 13:20). Live-resolve
      // only attaches if the spread block already started. Displayed
      // Trade Time Stamp is not changed.
      const tsADate = a[tsIdx] instanceof Date ? a[tsIdx] : null;
      const tsBDate = b[tsIdx] instanceof Date ? b[tsIdx] : null;
      const dayA = tsADate
        ? Utilities.formatDate(tsADate, tz, "yyyy-MM-dd")
        : "";
      const dayB = tsBDate
        ? Utilities.formatDate(tsBDate, tz, "yyyy-MM-dd")
        : "";
      if (dayA !== dayB) return dayA.localeCompare(dayB);

      const actA = (a[colMap["action"] - 1] || "")
        .toString()
        .trim()
        .toUpperCase();
      const actB = (b[colMap["action"] - 1] || "")
        .toString()
        .trim()
        .toUpperCase();
      const radRankA = actA === "RAD" ? 1 : 0;
      const radRankB = actB === "RAD" ? 1 : 0;
      if (radRankA !== radRankB) return radRankA - radRankB;

      const tsA = tsADate ? tsADate.getTime() : 0;
      const tsB = tsBDate ? tsBDate.getTime() : 0;
      if (tsA !== tsB) return tsA - tsB;

      // Tie-breaker: lower strike first within the same timestamp.
      const strikeA = Number(a[colMap["option strike"] - 1]) || 0;
      const strikeB = Number(b[colMap["option strike"] - 1]) || 0;
      return strikeA - strikeB;
    });

    // ─────────────────────────────────────────────────────────────────────────
    // Step 3: Assign Spread Group ID — THREE sub-passes.
    // ─────────────────────────────────────────────────────────────────────────

    const getTsMinute = (tsVal) => {
      if (!(tsVal instanceof Date) || isNaN(tsVal.getTime())) return "NO-TS";
      return Utilities.formatDate(tsVal, tz, "yyyy-MM-dd HH:mm");
    };

    // ----- Sub-pass A: Collect opening spread legs and their strikes -----
    let openGroupStrikesMap = {};
    let openGroupStrategyMap = {};
    let openGroupFirstRowMap = {};

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const actionRaw = (row[colMap["action"] - 1] || "")
        .toString()
        .trim()
        .toUpperCase();
      const strat = (row[colMap["strategy type"] - 1] || "")
        .toString()
        .toUpperCase();

      const isSpreadOpen =
        (actionRaw === "BUY TO OPEN" || actionRaw === "SELL TO OPEN") &&
        (strat.includes("PCS") ||
          strat.includes("PDS") ||
          strat.includes("CCS") ||
          strat.includes("CDS") ||
          strat.includes("BUTTERFLY") ||
          strat.includes("IRON CONDOR") ||
          strat.includes("IC"));
      if (!isSpreadOpen) continue;

      const acct = (row[colMap["account"] - 1] || "").toString().toUpperCase();
      const ticker = (row[colMap["ticker"] - 1] || "").toString().toUpperCase();
      const exp = row[colMap["option expiration"] - 1];
      const expStr =
        exp instanceof Date
          ? Utilities.formatDate(exp, tz, "yyyy-MM-dd")
          : (exp || "").toString();
      const strike = Number(row[colMap["option strike"] - 1]) || 0;
      const tsMin = getTsMinute(row[colMap["trade time stamp"] - 1]);

      const groupKey = `${acct}|${ticker}|${expStr}|${strat}|${tsMin}`;

      if (!openGroupStrikesMap[groupKey]) {
        openGroupStrikesMap[groupKey] = [];
        openGroupFirstRowMap[groupKey] = i;
      }
      if (!openGroupStrategyMap[groupKey])
        openGroupStrategyMap[groupKey] = strat;
      openGroupStrikesMap[groupKey].push(strike);
    }

    // ----- Sub-pass B: Build spreadKeyMap and spreadRangeMap -----
    let spreadKeyMap = {};
    let spreadRangeMap = {};

    for (const [groupKey, strikes] of Object.entries(openGroupStrikesMap)) {
      const parts = groupKey.split("|");
      const acctPart = parts[0];
      const tkrPart = parts[1];
      const expPart = parts[2];
      const stratPart = parts[3];

      const unique = [...new Set(strikes)].sort((a, b) => a - b);
      const strikeMin = unique[0];
      const strikeMax = unique[unique.length - 1];

      const spreadGroupId = `SPREAD-${tkrPart}-${stratPart}-${expPart}-${strikeMin}-${strikeMax}`;
      spreadKeyMap[groupKey] = spreadGroupId;

      const rangeKey = `${acctPart}|${tkrPart}|${expPart}|${stratPart}`;
      if (!spreadRangeMap[rangeKey]) spreadRangeMap[rangeKey] = [];
      spreadRangeMap[rangeKey].push({
        min: strikeMin,
        max: strikeMax,
        groupId: spreadGroupId,
      });
    }

    // ----- Sub-pass C: Stamp Spread Group IDs -----
    // Opening legs only — closing legs are intentionally deferred to Step 4.
    const SPREAD_STRAT_TERMS = [
      "PCS",
      "PDS",
      "CCS",
      "CDS",
      "BUTTERFLY",
      "IRON CONDOR",
      "IC",
    ];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const actionRaw = (row[colMap["action"] - 1] || "")
        .toString()
        .trim()
        .toUpperCase();
      const strat = (row[colMap["strategy type"] - 1] || "")
        .toString()
        .toUpperCase();

      const isSpreadStrategy = SPREAD_STRAT_TERMS.some((t) =>
        strat.includes(t),
      );
      const isRadSpreadExpiration =
        actionRaw === "RAD" &&
        isSpreadStrategy &&
        row[colMap["option strike"] - 1] &&
        row[colMap["option expiration"] - 1];

      if (!isSpreadStrategy && !isRadSpreadExpiration) continue;

      const acct = (row[colMap["account"] - 1] || "").toString().toUpperCase();
      const ticker = (row[colMap["ticker"] - 1] || "").toString().toUpperCase();
      const exp = row[colMap["option expiration"] - 1];
      const expStr =
        exp instanceof Date
          ? Utilities.formatDate(exp, tz, "yyyy-MM-dd")
          : (exp || "").toString();

      if (actionRaw === "BUY TO OPEN" || actionRaw === "SELL TO OPEN") {
        // Opening legs: stamp using the timestamp-based group key.
        const tsMin = getTsMinute(row[colMap["trade time stamp"] - 1]);
        const groupKey = `${acct}|${ticker}|${expStr}|${strat}|${tsMin}`;
        if (spreadKeyMap[groupKey]) {
          row[colMap["spread group id"] - 1] = spreadKeyMap[groupKey];
        }
      } else {
        // ── INTENTIONALLY DEFERRED to the main block loop (Step 4). ──────────
        // WHY: Sub-pass C runs before Step 4, so the live block-unit state does
        // not exist here yet. When two sequential spreads share a boundary strike
        // (e.g. 300/310 PDS followed by 310/315 PDS), a static range-containment
        // lookup always picks the first range — which is wrong once the first
        // spread has already closed flat. The main block loop uses
        // resolveLiveSpreadGroupId() which checks which spread group currently
        // has unit > 0 at the moment the closing row is processed, giving the
        // correct answer every time.
        // Leave Spread Group ID blank here — Step 4 will populate it.
      }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Step 4: Main loop — assign all block fields and calculate P&L on close.
    // ─────────────────────────────────────────────────────────────────────────
    let blocks = {};

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const acct = (row[colMap["account"] - 1] || "").toString().toUpperCase();
      const ticker = (row[colMap["ticker"] - 1] || "").toString().toUpperCase();
      const actionRaw = row[colMap["action"] - 1].toString().trim();
      const action = actionRaw.toUpperCase();

      // Skip pure ledger rows that have no Ticker.
      // IMPORTANT:
      // SYMBOL CHANGE rows are NOT skipped here even if they are JRN-origin rows,
      // because they need to transfer the live stock block from old ticker -> new ticker.
      if (
        ["EFN", "JRN", "DOI", "CRC", "CDB"].includes(action) &&
        action !== "SYMBOL CHANGE" &&
        !ticker
      ) {
        row[colMap["position id"] - 1] = "";
        row[colMap["trade group id"] - 1] = "";
        continue;
      }

      // SPECIAL HANDLING FOR SYMBOL CHANGE rows
      // WHY:
      // A stock symbol rename should keep one continuous stock block.
      // If the row itself already carries the real post-change ticker, keep that
      // ticker on the row while inheriting the existing stock block identity.
      if (action === "SYMBOL CHANGE") {
        const scInfo = parseSymbolChangeFromRow(row, colMap);

        if (!scInfo) {
          importIssuesAdd(
            ctx,
            "WARN",
            dataStartRow + i,
            "Notes",
            "",
            "SYMBOL CHANGE row has no parsable FROM/TO info in Notes. Row left unlinked.",
          );
          if (colMap["position id"] !== undefined)
            row[colMap["position id"] - 1] = "";
          if (colMap["trade group id"] !== undefined)
            row[colMap["trade group id"] - 1] = "";
          continue;
        }

        const fromTicker = String(scInfo.fromResolved || scInfo.fromRaw || "")
          .trim()
          .toUpperCase();
        const toRaw = String(scInfo.toRaw || "")
          .trim()
          .toUpperCase();
        const toResolved = String(scInfo.toResolved || "")
          .trim()
          .toUpperCase();
        const rowTicker = String(row[colMap["ticker"] - 1] || "")
          .trim()
          .toUpperCase();

        const targetTicker = resolveSymbolChangeTargetTicker_(
          fromTicker,
          toRaw,
          toResolved,
          rowTicker,
        );
        const displayTickerCandidate =
          rowTicker || toResolved || toRaw || fromTicker;
        const outputTicker = !isCusipLikeSymbolValue_(displayTickerCandidate)
          ? displayTickerCandidate
          : targetTicker || fromTicker;

        if (!fromTicker || !outputTicker) {
          importIssuesAdd(
            ctx,
            "WARN",
            dataStartRow + i,
            "Ticker",
            "FROM=" +
              fromTicker +
              " | ROW=" +
              rowTicker +
              " | TO=" +
              toResolved,
            "SYMBOL CHANGE row missing usable old/new ticker. Row left unlinked.",
          );
          if (colMap["position id"] !== undefined)
            row[colMap["position id"] - 1] = "";
          if (colMap["trade group id"] !== undefined)
            row[colMap["trade group id"] - 1] = "";
          continue;
        }

        const oldKey = `${acct}|${fromTicker}`;
        const oldBlock = blocks[oldKey];

        if (!oldBlock || !oldBlock.positionId) {
          // Source token has no live block. Common and usually not a hole:
          //   1) TOS emits a duplicate SYMBOL CHANGE after the first row
          //      already moved Account|FROM → Account|TO.
          //   2) Phase 1 already booked the shares under the destination
          //      ticker (LURAF spin-off, PAUIF emit) so FROM never opened.
          //   3) CUSIP listing change where FROM_RESOLVED === TO_RESOLVED.
          // Attach the row to the destination block when one is live.
          // WARN only when there is no source, no dest, and it is not a no-op.
          const destTickerForLookup = String(targetTicker || outputTicker || "")
            .trim()
            .toUpperCase();
          const destKey = destTickerForLookup
            ? `${acct}|${destTickerForLookup}`
            : "";
          const destBlock = destKey ? blocks[destKey] : null;
          const destLive = !!(destBlock && destBlock.positionId);
          const isNoOpRename =
            !!fromTicker &&
            !!destTickerForLookup &&
            fromTicker === destTickerForLookup;

          if (destLive) {
            if (colMap["ticker"] !== undefined)
              row[colMap["ticker"] - 1] = outputTicker || destTickerForLookup;
            if (colMap["trade group id"] !== undefined)
              row[colMap["trade group id"] - 1] = destBlock.tradeGroupId || "";
            if (colMap["position id"] !== undefined)
              row[colMap["position id"] - 1] = destBlock.positionId || "";
            if (colMap["block number"] !== undefined)
              row[colMap["block number"] - 1] = destBlock.block || "";
            if (colMap["block start flag"] !== undefined)
              row[colMap["block start flag"] - 1] = 0;
            if (colMap["block close flag/p&l"] !== undefined)
              row[colMap["block close flag/p&l"] - 1] = 0;
            if (colMap["running position quantity"] !== undefined) {
              row[colMap["running position quantity"] - 1] = Number(
                destBlock.runningQty || 0,
              );
            }
            if (colMap["trade status"] !== undefined) {
              row[colMap["trade status"] - 1] =
                Number(destBlock.runningQty || 0) === 0 ? "Closed" : "Open";
            }
            importIssuesAdd(
              ctx,
              "INFO",
              dataStartRow + i,
              "Ticker",
              fromTicker + " -> " + (outputTicker || destTickerForLookup),
              "SYMBOL CHANGE had no live source block; attached to existing destination block (duplicate rename or shares already booked under dest).",
            );
            continue;
          }

          if (isNoOpRename) {
            importIssuesAdd(
              ctx,
              "INFO",
              dataStartRow + i,
              "Ticker",
              fromTicker + " -> " + destTickerForLookup,
              "SYMBOL CHANGE is a same-ticker listing/CUSIP no-op with no stock block under that token. Left unlinked on purpose.",
            );
            if (colMap["position id"] !== undefined)
              row[colMap["position id"] - 1] = "";
            if (colMap["trade group id"] !== undefined)
              row[colMap["trade group id"] - 1] = "";
            continue;
          }

          importIssuesAdd(
            ctx,
            "WARN",
            dataStartRow + i,
            "Ticker",
            fromTicker + " -> " + outputTicker,
            "SYMBOL CHANGE row found but no active source stock block exists for old ticker. Row left unlinked.",
          );
          if (colMap["position id"] !== undefined)
            row[colMap["position id"] - 1] = "";
          if (colMap["trade group id"] !== undefined)
            row[colMap["trade group id"] - 1] = "";
          continue;
        }

        // No usable destination yet: keep the existing source block identity,
        // but do NOT force the row ticker back to the old symbol if the row already
        // has a real post-change ticker.
        if (!targetTicker || targetTicker === fromTicker) {
          if (colMap["ticker"] !== undefined)
            row[colMap["ticker"] - 1] = outputTicker;
          if (colMap["trade group id"] !== undefined)
            row[colMap["trade group id"] - 1] = oldBlock.tradeGroupId || "";
          if (colMap["position id"] !== undefined)
            row[colMap["position id"] - 1] = oldBlock.positionId || "";
          if (colMap["block number"] !== undefined)
            row[colMap["block number"] - 1] = oldBlock.block || "";
          if (colMap["block start flag"] !== undefined)
            row[colMap["block start flag"] - 1] = 0;
          if (colMap["block close flag/p&l"] !== undefined)
            row[colMap["block close flag/p&l"] - 1] = 0;
          if (colMap["running position quantity"] !== undefined) {
            row[colMap["running position quantity"] - 1] = Number(
              oldBlock.runningQty || 0,
            );
          }
          if (colMap["trade status"] !== undefined) {
            row[colMap["trade status"] - 1] =
              Number(oldBlock.runningQty || 0) === 0 ? "Closed" : "Open";
          }

          importIssuesAdd(
            ctx,
            "INFO",
            dataStartRow + i,
            "Ticker",
            fromTicker + " -> " + outputTicker,
            "SYMBOL CHANGE kept the existing stock block identity without opening a new block because no separate destination ticker was yet usable.",
          );
          continue;
        }

        const newKey = `${acct}|${targetTicker}`;
        const newBlockExisting = blocks[newKey];

        if (
          newBlockExisting &&
          newBlockExisting.positionId &&
          Number(newBlockExisting.runningQty || 0) !== 0
        ) {
          importIssuesAdd(
            ctx,
            "WARN",
            dataStartRow + i,
            "Ticker",
            fromTicker + " -> " + targetTicker,
            "SYMBOL CHANGE destination ticker already has an active block. Source block was NOT merged automatically.",
          );
          if (colMap["position id"] !== undefined)
            row[colMap["position id"] - 1] = "";
          if (colMap["trade group id"] !== undefined)
            row[colMap["trade group id"] - 1] = "";
          continue;
        }

        const movedBlock = {
          unit: Number(oldBlock.unit || 0),
          block: Number(oldBlock.block || 1),
          runningQty: Number(oldBlock.runningQty || 0),
          pnl: Number(oldBlock.pnl || 0),
          entryCost: Number(oldBlock.entryCost || 0),
          openTs: oldBlock.openTs || null,
          positionId: oldBlock.positionId || "",
          tradeGroupId: oldBlock.tradeGroupId || "",
          strategyType: oldBlock.strategyType || "LONG STOCK",
        };

        blocks[newKey] = movedBlock;
        delete blocks[oldKey];

        if (colMap["ticker"] !== undefined)
          row[colMap["ticker"] - 1] = outputTicker;
        if (colMap["trade group id"] !== undefined)
          row[colMap["trade group id"] - 1] = movedBlock.tradeGroupId || "";
        if (colMap["position id"] !== undefined)
          row[colMap["position id"] - 1] = movedBlock.positionId || "";
        if (colMap["block number"] !== undefined)
          row[colMap["block number"] - 1] = movedBlock.block || "";
        if (colMap["block start flag"] !== undefined)
          row[colMap["block start flag"] - 1] = 0;
        if (colMap["block close flag/p&l"] !== undefined)
          row[colMap["block close flag/p&l"] - 1] = 0;
        if (colMap["running position quantity"] !== undefined) {
          row[colMap["running position quantity"] - 1] = Number(
            movedBlock.runningQty || 0,
          );
        }
        if (colMap["trade status"] !== undefined) {
          row[colMap["trade status"] - 1] =
            Number(movedBlock.runningQty || 0) === 0 ? "Closed" : "Open";
        }

        importIssuesAdd(
          ctx,
          "INFO",
          dataStartRow + i,
          "Ticker",
          fromTicker + " -> " + targetTicker + " | ROW=" + outputTicker,
          "SYMBOL CHANGE transferred the active stock block to the destination ticker while preserving the row ticker for display and analytics.",
        );
        continue;
      }
      // =====================================================================
      // END SYMBOL CHANGE HANDLER
      // =======================================================================

      // NOTE: `spreadId` is `let` — the live resolution block below may reassign it.
      let spreadId = (row[colMap["spread group id"] - 1] || "").toString();
      let qty = Number(row[colMap["quantity"] - 1]) || 0;
      const strategyType = (row[colMap["strategy type"] - 1] || "")
        .toString()
        .toUpperCase();

      let tradeType = (row[colMap["trade type"] - 1] || "")
        .toString()
        .toUpperCase()
        .trim();
      if (!tradeType && ticker) tradeType = "OPTION";

      // A row with expiration + strike + call/put is an option even when
      // Helper labeled Trade Type SPREAD and Spread Group ID is still blank.
      // Otherwise the grouping key falls through to Account|Ticker (stock)
      // and a put-assignment RAD decrements the new share lot (100 → 99).
      const expForType = row[colMap["option expiration"] - 1];
      const strikeForType = Number(row[colMap["option strike"] - 1]) || 0;
      const cpForType = (row[colMap["call/put"] - 1] || "")
        .toString()
        .toUpperCase()
        .trim();
      const hasOptionIdentity =
        !spreadId &&
        ticker &&
        strikeForType > 0 &&
        !!expForType &&
        (cpForType === "C" ||
          cpForType === "P" ||
          cpForType === "CALL" ||
          cpForType === "PUT");
      if (hasOptionIdentity && tradeType !== "OPTION") {
        tradeType = "OPTION";
        if (colMap["trade type"] !== undefined) {
          row[colMap["trade type"] - 1] = "OPTION";
        }
      }

      // === LIVE SPREAD GROUP ID RESOLUTION FOR CLOSING/RAD LEGS ================
      // WHY: Sub-pass C intentionally left closing leg Spread Group IDs blank.
      // We fill them here using resolveLiveSpreadGroupId(), which checks the live
      // block unit state to disambiguate boundary-strike collisions between
      // sequential spreads that share a strike (e.g. 300/310 then 310/315 PDS).
      // Opening legs already have their Spread Group ID from Sub-pass C — only
      // blank-spreadId rows with a spread strategy need resolution here.
      const isSpreadStrategy_ = SPREAD_STRAT_TERMS.some((t) =>
        strategyType.includes(t),
      );
      const isClosingOrRAD = action.includes("TO CLOSE") || action === "RAD";
      if (!spreadId && isSpreadStrategy_ && isClosingOrRAD && ticker) {
        const exp_ = row[colMap["option expiration"] - 1];
        const expStr_ =
          exp_ instanceof Date
            ? Utilities.formatDate(exp_, tz, "yyyy-MM-dd")
            : (exp_ || "").toString();
        const strike_ = Number(row[colMap["option strike"] - 1]) || 0;
        if (expStr_ && strike_) {
          const resolved = resolveLiveSpreadGroupId(
            acct,
            ticker,
            expStr_,
            strategyType,
            strike_,
            spreadRangeMap,
            blocks,
          );
          if (resolved) {
            row[colMap["spread group id"] - 1] = resolved;
            spreadId = resolved; // local var — used immediately in the grouping key below
          } else if (
            !hasContainingSpreadWindow(
              acct,
              ticker,
              expStr_,
              strike_,
              spreadRangeMap,
            )
          ) {
            ctxMissingSpreadGroup++;
            importIssuesAdd(
              ctx,
              "WARN",
              dataStartRow + i,
              "Spread Group ID",
              `${acct}|${ticker}|${expStr_}|${strike_}`,
              "No live spread group found for this closing/RAD row. " +
                "Possible cause: no matching open block at this timestamp. " +
                "Check that the opening trade exists and processed before this row.",
            );
          }
        }
      }
      // =========================================================================

      // === GROUPING KEY — Spread Group ID takes priority for ALL rows including RAD ===
      let key;
      if (spreadId) {
        key = `${acct}|${spreadId}`;
        tradeType = "SPREAD";
        row[colMap["trade type"] - 1] = "SPREAD";
      } else if (tradeType === "OPTION") {
        const exp = row[colMap["option expiration"] - 1];
        const expStr =
          exp instanceof Date
            ? Utilities.formatDate(exp, tz, "yyyy-MM-dd")
            : (exp || "").toString();
        const strike = Number(row[colMap["option strike"] - 1]) || 0;
        const cp = (row[colMap["call/put"] - 1] || "")
          .toString()
          .toUpperCase()
          .replace("CALL", "C")
          .replace("PUT", "P");
        key = `${acct}|${ticker}|${expStr}|${strike}|${cp}`;
      } else {
        key = `${acct}|${ticker}`;
      }

      if (!blocks[key])
        blocks[key] = {
          unit: 0,
          block: 1,
          runningQty: 0,
          pnl: 0,
          entryCost: 0,
          openTs: null,
          positionId: "",
          strategyType: "",
        };

      // === ROBUST DELTA ===
      let delta = 0;
      if (action.includes("SELL TO OPEN") || action.includes("BUY TO OPEN"))
        delta = 1;
      if (action.includes("BUY TO CLOSE") || action.includes("SELL TO CLOSE"))
        delta = -1;

      // === SPECIAL HANDLING FOR RAD "Opt Expired" rows ===
      if (action === "RAD" && ticker) {
        delta = -1;
        qty = Math.abs(qty);

        row[colMap["entry price"] - 1] = 0;
        row[colMap["total cost"] - 1] = 0;

        row[colMap["strategy type"] - 1] = blocks[key].strategyType || "";

        let cpRad = (row[colMap["call/put"] - 1] || "")
          .toString()
          .toUpperCase();
        if (cpRad === "CALL") row[colMap["call/put"] - 1] = "C";
        if (cpRad === "PUT") row[colMap["call/put"] - 1] = "P";

        row[colMap["closing date"] - 1] = row[colMap["trade date"] - 1];
      }

      // SPECIAL HANDLING FOR SPLIT rows
      // SPLIT should adjust the existing stock block in place.
      // It must NOT open a new block, close a block, or create a new Position ID.
      if (action === "SPLIT" && ticker) {
        let existingBlock = blocks[key];
        const splitInfo = parseSplitRatioFromRow(row, colMap);

        // Rename + split safety net:
        // If the new ticker has no live block yet, try to find exactly one open stock
        // block in the same account whose pre-split qty implies this post-split qty.
        if ((!existingBlock || !existingBlock.positionId) && splitInfo) {
          const renameSource = findRenameSourceStockBlockForSplit_(
            acct,
            ticker,
            splitInfo,
            blocks,
          );
          if (renameSource) {
            blocks[key] = renameSource.block;
            delete blocks[renameSource.sourceKey];
            existingBlock = blocks[key];

            importIssuesAdd(
              ctx,
              "INFO",
              dataStartRow + i,
              "Ticker",
              renameSource.sourceTicker + " -> " + ticker,
              "Inferred SYMBOL CHANGE from split context before applying SPLIT. " +
                "Matched preQty=" +
                renameSource.preQty +
                ", expectedPost=" +
                renameSource.expectedPost +
                ".",
            );
          }
        }

        if (!existingBlock || !existingBlock.positionId) {
          importIssuesAdd(
            ctx,
            "WARN",
            dataStartRow + i,
            "Ticker",
            ticker,
            "SPLIT row for " +
              ticker +
              " has no active stock block in memory. " +
              "Split may have occurred before any position was opened, or ticker spelling differs.",
          );
          continue;
        }

        const runningBeforeSplit = Number(existingBlock.runningQty || 0);

        // Ignore the broker's duplicate partner row if it does not match the live block.
        if (!isMatchingSplitRowForBlock_(runningBeforeSplit, splitInfo)) {
          importIssuesAdd(
            ctx,
            "INFO",
            dataStartRow + i,
            "Notes",
            colMap["notes"] ? String(row[colMap["notes"] - 1] || "") : "",
            "Duplicate/unmatched SPLIT row ignored. It does not match the live pre/post quantity for this block.",
          );
          continue;
        }

        let splitDelta = Number(qty || 0);

        if (
          splitInfo &&
          isFinite(splitInfo.numerator) &&
          isFinite(splitInfo.denominator) &&
          splitInfo.denominator > 0
        ) {
          const expectedPost = computeExpectedPostSplitQty_(
            runningBeforeSplit,
            splitInfo,
          );
          if (isFinite(expectedPost)) {
            splitDelta = expectedPost - runningBeforeSplit;
          }
        }

        if (!isFinite(splitDelta)) {
          importIssuesAdd(
            ctx,
            "WARN",
            dataStartRow + i,
            "Quantity",
            ticker,
            "SPLIT row for " +
              ticker +
              " has no usable split delta. " +
              "Check Notes for FORWARD/REVERSE SPLIT ratio text.",
          );
          continue;
        }

        existingBlock.runningQty = runningBeforeSplit + splitDelta;

        row[colMap["quantity"] - 1] = splitDelta;
        if (colMap["signed quantity"] !== undefined) {
          row[colMap["signed quantity"] - 1] = splitDelta;
        }
        row[colMap["running position quantity"] - 1] = existingBlock.runningQty;

        if (colMap["trade group id"] !== undefined) {
          row[colMap["trade group id"] - 1] =
            existingBlock.tradeGroupId ||
            row[colMap["trade group id"] - 1] ||
            "";
        }
        if (colMap["position id"] !== undefined) {
          row[colMap["position id"] - 1] = existingBlock.positionId || "";
        }
        if (colMap["block number"] !== undefined) {
          row[colMap["block number"] - 1] = existingBlock.block || "";
        }
        if (colMap["block start flag"] !== undefined) {
          row[colMap["block start flag"] - 1] = 0;
        }
        if (colMap["block close flag/p&l"] !== undefined) {
          row[colMap["block close flag/p&l"] - 1] = 0;
        }
        if (colMap["trade status"] !== undefined) {
          row[colMap["trade status"] - 1] =
            existingBlock.runningQty === 0 ? "Closed" : "Open";
        }

        continue;
      }
      // END SPLIT HANDLER

      const prevUnit = Number(blocks[key].unit || 0);
      const prevRunningQty = Number(blocks[key].runningQty || 0);

      // Leftover To Close after this option key already flattened.
      // TOS often keeps Pos Effect = TO CLOSE on the extra fill that
      // flips a long into a short (CVNA 2024-03-25). Treating that row
      // as another close drives Running Position Quantity through -1
      // to -2 on the cover and never Block-Closes TG002.
      // Broker Action is left unchanged. Orphan first-closes stay
      // negative because block is still 1.
      let leftoverCloseAsOpen = false;
      let blockStrategyType = strategyType;
      if (
        tradeType === "OPTION" &&
        !spreadId &&
        action.includes("TO CLOSE") &&
        action !== "RAD" &&
        prevUnit === 0 &&
        Number(blocks[key].block || 1) > 1
      ) {
        leftoverCloseAsOpen = true;
        delta = 1;

        const cpLeft = (row[colMap["call/put"] - 1] || "")
          .toString()
          .toUpperCase()
          .replace("CALL", "C")
          .replace("PUT", "P");
        if (action.includes("SELL TO CLOSE")) {
          blockStrategyType = cpLeft === "P" ? "SHORT PUT" : "SHORT CALL";
        } else if (action.includes("BUY TO CLOSE")) {
          blockStrategyType = cpLeft === "P" ? "LONG PUT" : "LONG CALL";
        }

        if (colMap["strategy type"] !== undefined) {
          const stratWrite =
            blockStrategyType === "SHORT PUT"
              ? "SHORT PUT"
              : blockStrategyType === "LONG PUT"
                ? "LONG PUT"
                : blockStrategyType === "LONG CALL"
                  ? "LONG CALL"
                  : "SHORT CALL";
          row[colMap["strategy type"] - 1] = stratWrite;
        }

        importIssuesAdd(
          ctx,
          "INFO",
          dataStartRow + i,
          "Action",
          actionRaw,
          "Leftover To Close after flat option block treated as opening the other side. Broker Action left unchanged.",
        );
      }

      // Family C — extra close/RAD on a Spread Group ID after that
      // group already flattened (block > 1, unit 0).
      // WHY: unsigned spread units can hit 0 after extra closes of
      // one leg (DT SOXL PDS 155-160, LT SPY CCS 392-393). The next
      // TO CLOSE was starting TG002 and driving Running Position
      // Quantity negative. That is fake inventory for Master.
      // WHAT: keep the broker Action, do not change units, keep the
      // row on the last closed TG. Do NOT treat this as opening the
      // other side (that is Family B, single-option only).
      // Orphan first-closes stay untouched because block is still 1.
      let extraCloseAfterFlat = false;
      if (
        spreadId &&
        !leftoverCloseAsOpen &&
        prevUnit === 0 &&
        Number(blocks[key].block || 1) > 1 &&
        delta < 0 &&
        (action.includes("TO CLOSE") || action === "RAD")
      ) {
        extraCloseAfterFlat = true;
        delta = 0;
        importIssuesAdd(
          ctx,
          "INFO",
          dataStartRow + i,
          "Action",
          actionRaw,
          "EXTRA_CLOSE_AFTER_FLAT — spread already flat; row kept on last closed TG; running qty not reduced below 0.",
        );
      }

      // Family D — RAD with no live units on this block key.
      // WHY: RAD always started as delta = -1. If the option/spread
      // was already flat (or the RAD landed on the wrong ticker,
      // e.g. SQQQ1), Running Position Quantity went negative.
      // Live RADs are unchanged: prevUnit > 0 still uses delta = -1.
      // WHAT: keep the RAD row, do not change units, stay on the
      // last closed TG when one exists. Identity bugs (SQQQ1) are
      // a later parse fix; this only stops fake inventory.
      let unmatchedRad = false;
      if (
        action === "RAD" &&
        !leftoverCloseAsOpen &&
        !extraCloseAfterFlat &&
        prevUnit === 0 &&
        delta < 0
      ) {
        unmatchedRad = true;
        delta = 0;
        importIssuesAdd(
          ctx,
          "INFO",
          dataStartRow + i,
          "Action",
          actionRaw,
          "RAD_NO_LIVE — no open units on this key; RAD row kept; running qty not reduced below 0.",
        );
      }

      // Family A — orphan first close on a single-option key.
      // WHY: BTC/STC with no matching open (DT SPX R3284 4505C,
      // R3285 4510C). There is no lot to attach to. Leftover-close
      // as open would invent a short. Leaving delta = -1 writes
      // fake inventory onto Master.
      // WHAT: keep the broker Action, do not change units.
      // Family B is the same shape with block > 1 (already handled).
      // Spreads are Family C. RAD is Family D.
      let orphanClose = false;
      if (
        tradeType === "OPTION" &&
        !spreadId &&
        !leftoverCloseAsOpen &&
        !extraCloseAfterFlat &&
        !unmatchedRad &&
        action.includes("TO CLOSE") &&
        action !== "RAD" &&
        prevUnit === 0 &&
        Number(blocks[key].block || 1) <= 1 &&
        delta < 0
      ) {
        orphanClose = true;
        delta = 0;
        importIssuesAdd(
          ctx,
          "INFO",
          dataStartRow + i,
          "Action",
          actionRaw,
          "ORPHAN_CLOSE — no open exists on this option key; row kept; running qty not reduced below 0.",
        );
      }

      blocks[key].unit += delta * qty;
      blocks[key].runningQty += delta * qty;

      const newUnit = Number(blocks[key].unit || 0);
      const newRunningQty = Number(blocks[key].runningQty || 0);
      const curBlock = blocks[key].block;
      // Extra closes belong on the group that just flattened, not TG00N+1.
      let tgBlock = curBlock;
      if ((extraCloseAfterFlat || unmatchedRad) && Number(curBlock) > 1) {
        tgBlock = curBlock - 1;
      }

      let blkStart = 0;
      let blkClose = 0;

      // STOCK positions must be split-aware.
      // Open/close is based ONLY on Running Position Quantity transitions.
      // OPTIONS / SPREADS preserve legacy unit-based behavior.
      if (tradeType === "STOCK") {
        blkStart = prevRunningQty === 0 && newRunningQty !== 0 ? 1 : 0;
        blkClose = prevRunningQty !== 0 && newRunningQty === 0 ? 1 : 0;
      } else {
        blkStart = prevUnit === 0 && newUnit !== 0 ? 1 : 0;
        blkClose = prevUnit !== 0 && newUnit === 0 ? 1 : 0;
      }

      // Reset P&L accumulators at the START of every new block.
      if (blkStart) {
        blocks[key].pnl = 0;
        blocks[key].entryCost = 0;
      }

      // ── CTX counters ──────────────────────────────────────────────────────
      if (blkStart) ctxBlocksOpened++;
      if (blkClose) ctxBlocksClosed++;
      if (action === "RAD" && ticker) ctxRADRows++;
      // ─────────────────────────────────────────────────────────────────────

      // Build unique Position ID when the block starts.
      let posId = blocks[key].positionId;
      if (blkStart && ticker && tradeType !== "") {
        const stratAbbrevPos = getStratAbbrev(blockStrategyType || "");
        const tgSuffixPos = `TG${String(tgBlock).padStart(3, "0")}`;

        if (spreadId) {
          posId = `${spreadId}-${tgSuffixPos}`;
        } else if (tradeType === "OPTION") {
          const expStr2 =
            colMap["option expiration"] !== undefined &&
            row[colMap["option expiration"] - 1] instanceof Date
              ? Utilities.formatDate(
                  row[colMap["option expiration"] - 1],
                  tz,
                  "yyMMdd",
                )
              : "";
          const strike2 = Number(row[colMap["option strike"] - 1]) || 0;
          const cp2 = row[colMap["call/put"] - 1]
            .toString()
            .toUpperCase()
            .replace("CALL", "C")
            .replace("PUT", "P");
          posId = `${acct}-${ticker}-${stratAbbrevPos}${expStr2 ? `-${expStr2}` : ""}-${String(Math.round(strike2)).padStart(5, "0")}${cp2}-${tgSuffixPos}`;
        } else {
          posId = `${acct}-${ticker}-${stratAbbrevPos}-${tgSuffixPos}`;
        }

        blocks[key].positionId = posId;
        blocks[key].strategyType = blockStrategyType;
      }
      row[colMap["position id"] - 1] = posId;

      // Trade Group ID
      const optExpRaw = row[colMap["option expiration"] - 1];
      const optExpStr =
        optExpRaw instanceof Date
          ? Utilities.formatDate(optExpRaw, tz, "yyMMdd")
          : "";
      const isOption = tradeType === "OPTION";
      const stratAbbrev = getStratAbbrev(
        blockStrategyType || blocks[key].strategyType,
      );
      const tgSuffix = `TG${String(tgBlock).padStart(3, "0")}`;

      const tradeGroupId = spreadId
        ? `${spreadId}-${tgSuffix}`
        : `${acct}-${ticker}-${stratAbbrev}${isOption && optExpStr ? "-" + optExpStr : ""}-${tgSuffix}`;

      row[colMap["trade group id"] - 1] = tradeGroupId;
      blocks[key].tradeGroupId = tradeGroupId;

      row[colMap["block start flag"] - 1] = blkStart;
      row[colMap["block number"] - 1] = tgBlock;
      row[colMap["block close flag/p&l"] - 1] = blkClose;
      row[colMap["running position quantity"] - 1] = blocks[key].runningQty;

      // === P&L CALCULATION ===
      const hasOptionFields = !!(
        row[colMap["option strike"] - 1] || row[colMap["option expiration"] - 1]
      );
      const multiplier = tradeType === "STOCK" || !hasOptionFields ? 1 : 100;
      let sign = 0;
      if (action.includes("SELL TO")) sign = 1;
      if (action.includes("BUY TO")) sign = -1;
      blocks[key].pnl +=
        sign * Number(row[colMap["entry price"] - 1]) * qty * multiplier;
      if (action.includes("OPEN") || leftoverCloseAsOpen) {
        blocks[key].entryCost += Math.abs(
          sign * Number(row[colMap["entry price"] - 1]) * qty * multiplier,
        );
      }

      if (blkClose) {
        row[colMap["realized p&l"] - 1] = blocks[key].pnl;
        row[colMap["percent p&l"] - 1] = blocks[key].entryCost
          ? (blocks[key].pnl / blocks[key].entryCost) * 100
          : 0;
        row[colMap["trade status"] - 1] = "Closed";
        row[colMap["closing date"] - 1] = row[colMap["trade date"] - 1];
        if (row[tsIdx] && blocks[key].openTs) {
          const days =
            (row[tsIdx] - blocks[key].openTs) / (1000 * 60 * 60 * 24);
          row[colMap["trade duration"] - 1] = Math.round(days * 100) / 100;
        }
        blocks[key].block++;
      } else if (blkStart) {
        row[colMap["trade status"] - 1] = "Open";
        blocks[key].openTs = row[tsIdx];
      } else if (extraCloseAfterFlat || unmatchedRad || orphanClose) {
        row[colMap["trade status"] - 1] = "Closed";
      }
    } // end main row loop

    // ── POST-PASS: Link stock settlement legs to their parent spread block ────
    const closingSpreadLookup = {};
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const blockCloseFlag = row[colMap["block close flag/p&l"] - 1];
      const rowTradeType = row[colMap["trade type"] - 1]
        .toString()
        .toUpperCase()
        .trim();
      if (blockCloseFlag !== 1 && blockCloseFlag !== "1") continue;
      if (rowTradeType === "STOCK") continue;
      const rowAcct = row[colMap["account"] - 1].toString().toUpperCase();
      const rowTicker = row[colMap["ticker"] - 1].toString().toUpperCase();
      const rowTgId = row[colMap["trade group id"] - 1].toString().trim();
      const rowStrat = row[colMap["strategy type"] - 1].toString().trim();
      const closingDate = row[colMap["closing date"] - 1];
      if (!rowAcct || !rowTicker || !rowTgId || !(closingDate instanceof Date))
        continue;
      const closingDateStr = Utilities.formatDate(
        closingDate,
        tz,
        "yyyy-MM-dd",
      );
      closingSpreadLookup[`${rowAcct}|${rowTicker}|${closingDateStr}`] = {
        tgId: rowTgId,
        strategy: rowStrat,
      };
    }

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowTradeType = row[colMap["trade type"] - 1]
        .toString()
        .toUpperCase()
        .trim();
      if (rowTradeType !== "STOCK") continue;
      const rowAction = row[colMap["action"] - 1]
        .toString()
        .trim()
        .toUpperCase();
      if (rowAction !== "BUY TO OPEN" && rowAction !== "SELL TO CLOSE")
        continue;
      const rowAcct = row[colMap["account"] - 1].toString().toUpperCase();
      const rowTicker = row[colMap["ticker"] - 1].toString().toUpperCase();
      const tradeDate = row[colMap["trade date"] - 1];
      if (!(tradeDate instanceof Date)) continue;
      const tradeDateStr = Utilities.formatDate(tradeDate, tz, "yyyy-MM-dd");
      const parent =
        closingSpreadLookup[`${rowAcct}|${rowTicker}|${tradeDateStr}`];
      if (!parent) continue;
      row[colMap["trade group id"] - 1] = `${parent.tgId}-ST`;
      row[colMap["position id"] - 1] = `${parent.tgId}-ST`;
      row[colMap["strategy type"] - 1] = parent.strategy;
      importIssuesAdd(
        ctx,
        "INFO",
        dataStartRow + i,
        "Trade Group ID",
        `${parent.tgId}-ST`,
        `Stock settlement leg linked to parent spread block. TG ID → ${parent.tgId}-ST`,
      );
    }
    // ── END POST-PASS ─────────────────────────────────────────────────────────

    let openAtEnd = 0;
    Object.values(blocks).forEach((b) => {
      const strategyUpper = String(b.strategyType || "")
        .trim()
        .toUpperCase();
      const looksLikeStock =
        strategyUpper.includes("STOCK") &&
        !strategyUpper.includes("CALL") &&
        !strategyUpper.includes("PUT");

      if (looksLikeStock) {
        if (Number(b.runningQty || 0) !== 0) openAtEnd++;
      } else {
        if (Number(b.unit || 0) !== 0) openAtEnd++;
      }
    });

    importIssuesSetMetric(ctx, "SourceRowsReadExclHeader", data.length);
    importIssuesSetMetric(ctx, "RowsWrittenExclHeader", data.length);
    importIssuesSetMetric(
      ctx,
      "SpreadGroupsBuilt",
      Object.keys(spreadKeyMap).length,
    );
    importIssuesSetMetric(ctx, "StrikeCollisions", ctxStrikeCollisions);
    importIssuesSetMetric(ctx, "BlocksOpened", ctxBlocksOpened);
    importIssuesSetMetric(ctx, "BlocksClosed", ctxBlocksClosed);
    importIssuesSetMetric(ctx, "RADRowsParsed", ctxRADRows);
    importIssuesSetMetric(ctx, "OpenPositionsAtEnd", openAtEnd);
    importIssuesSetMetric(ctx, "MissingSpreadGroupIds", ctxMissingSpreadGroup);
    importIssuesSetMetric(ctx, "Success", "1");

    if (openAtEnd > 0) {
      importIssuesAdd(
        ctx,
        "INFO",
        "End of data",
        "Open Positions",
        openAtEnd + " position(s)",
        "These positions have no closing event in the current dataset — " +
          "expected if you have live LEAP spreads still open.",
      );
    }

    const outputGrid = [helperData[0], helperData[1], helperData[2], ...data];
    stagingSheet.clearContents();
    stagingSheet
      .getRange(1, 1, outputGrid.length, outputGrid[0].length)
      .setValues(outputGrid);

    checkMissingDateTimeAndAlert(
      stagingSheet,
      4,
      "populateStagingWithBlockLogicV3",
    );

    uiAlertSafe(
      "✅ Block logic V3 updated! Trade Group ID increments once per open-to-flat block — " +
        "RAD expirations close perfectly.",
    );
  } catch (e) {
    importIssuesSetMetric(ctx, "Success", "0");
    importIssuesSetMetric(ctx, "ErrorMessage", e.message);
    importIssuesSetMetric(ctx, "ErrorStack", (e.stack || "").substring(0, 500));
    throw e;
  } finally {
    stagingIssuesFlush(ctx);
  }
}
