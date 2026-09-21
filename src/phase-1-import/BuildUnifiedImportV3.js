/**
 * BuildUnifiedImportV3.js
 *
 * Core Phase 1 function that builds the canonical "Schwab Import" sheet.
 *
 * High-level job:
 *   Merge data from:
 *     - TosTrades  (trade structure / legs / symbols authority)
 *     - TosTop     (fees, amounts, cash movements authority)
 *   into the single sheet that Phase 2 (Mapping) expects: "Schwab Import".
 *
 * Key responsibilities:
 *   - Combine both accounts (DT + LT) in one pass
 *   - Match trades to top-of-book rows (including CUSIP / symbol change handling)
 *   - Emit corporate-action and cash rows when needed
 *   - Write a clean, consistent table into "Schwab Import"
 *   - Log metrics and problems via ImportIssues.js
 *
 * This is currently the largest file in the project.
 * Current focus: correct the header and add clear high-level documentation
 * before any structural refactoring.
 *
 * Related files:
 *   - BuildUnifiedEnrichment.js    (Group A fee-matching factory)
 *   - BuildUnifiedIcRetag.js       (Group B IC-retag factory)
 *   - BuildUnifiedSheetFields.js   (Group C sheet/field factory)
 *   - BuildUnifiedSymbols.js       (Group E symbol/option factory)
 *   - BuildUnifiedExerciseAction.js (Group F exercise/action factory)
 *   - BuildUnifiedMappingSheets.js (Group G mapping-sheet factory)
 *   - BuildUnifiedSpecialParsers.js (Group H TosTop description parsers)
 *   - BuildUnifiedOutput.js        (Group I fee-group + sort-key factory)
 *   - TosSchwabImportPipeline.js   (produces TosTrades / TosTop)
 *   - ImportIssues.js
 *   - SettingsService.js
 *   - MapSchwabImportByHeadersV3.js (Phase 2 – runs after this)
 */

/**
 * buildUnifiedImportV3()
 *
 * Main entry point (called from the DB Tools menu).
 *
 * High-level job:
 *   Read the working sheets "TosTrades" and "TosTop",
 *   match trade legs to fee/amount rows,
 *   handle corporate actions and cash movements,
 *   and write one clean, unified table into the sheet "Schwab Import".
 *
 * This is the bridge between the raw TOS import pipeline
 * and Phase 2 (Mapping).
 *
 * Key behaviors:
 *   - Processes both DT and LT accounts in a single run
 *   - Uses a script lock to prevent overlapping runs
 *   - Logs everything under one Import Issues run context
 *   - Contains many nested helper functions (kept local for now)
 */
function buildUnifiedImportV3() {
  const tBuild = pipelineTimingNow();
  // ----------------------------
  // 0) Lock (prevents double-runs)
  // ----------------------------
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  // Everything in this run logs under ONE runId + step name.
  const ctx = importIssuesStart("buildUnifiedImportV3");

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1) Account mode (DT/LT)
    // We build BOTH accounts together in this function; this setting is recorded only for debugging.
    const accountModeSetting = String(getSetting("accountMode", "") || "")
      .trim()
      .toUpperCase();
    importIssuesSetMetric(
      ctx,
      "AccountModeSetting",
      accountModeSetting || "(blank)",
    );

    // =========================================================================
    // 2) LOCAL HELPERS
    //    Sheet lookup + TosTrades table access for this run.
    //    toStr now lives in Helpers.js (same name, now global).
    // =========================================================================

    // Loose find-by-name: trims and lowercases so "CusipMap " still works.
    function getSheetByNameLoose(name) {
      const target = String(name || "")
        .trim()
        .toLowerCase();
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        const sh = sheets[i];
        if (
          String(sh.getName() || "")
            .trim()
            .toLowerCase() === target
        )
          return sh;
      }
      return null;
    }

    function mustGetSheet(name) {
      const sh = ss.getSheetByName(name);
      if (!sh) throw new Error("Missing sheet: " + name);
      return sh;
    }

    /**
     * Reads a whole sheet as a "table":
     * - headers[] = row 1 strings
     * - idx{} = normalizedHeader -> columnIndex
     * - rows[] = all non-blank data rows (2D arrays)
     */
    function readSheetTable(sheet) {
      const vals = sheet.getDataRange().getValues();
      if (vals.length < 2) return { headers: [], idx: {}, rows: [] };

      const headers = vals[0].map((h) => String(h ?? "").trim());
      const idx = {};
      for (let c = 0; c < headers.length; c++) {
        idx[normalizeHeader(headers[c])] = c;
      }

      const rows = [];
      for (let r = 1; r < vals.length; r++) {
        const row = vals[r];
        const isBlank = row.every((v) => String(v ?? "").trim() === "");
        if (isBlank) continue;
        rows.push(row);
      }

      return { headers, idx, rows };
    }

    function cell(row, idx, headerName) {
      const c = idx[normalizeHeader(headerName)];
      return c === undefined ? "" : row[c];
    }

    function previewRow12(row) {
      if (!Array.isArray(row)) return row;
      return row.slice(0, 12);
    }

    // =========================================================================
    // NESTED HELPERS – ROADMAP
    //
    // Group A (fee enrichment) now lives in BuildUnifiedEnrichment.js:
    //   createTopTradeEnrichmentHelpers(...)
    //   Wired in after the TosTop queues exist (section 5).
    //
    // Everything below stays nested because many helpers need the in-memory
    // queues and maps that are built later in this same function.
    //
    // Remaining groups, in roughly this order:
    //
    //   B) Iron Condor (IC) retagging
    //      decideIcRetag (+ related lifecycle helpers later in the trade loop)
    //      → Inherit "IRON CONDOR" onto closing legs when inventory evidence exists.
    //
    //   C) Sheet / field utilities
    //      now in BuildUnifiedSheetFields.js
    //      (readSheetObjects, getField, roundTo, pricesClose)
    //
    //   D) Date & time normalization
    //      Shared in Helpers.js: normalizeDate, normalizeTime,
    //      normalizeTimeHHmmss, toDateObject
    //      Still nested here: makeTradeMatchKey
    //
    //   E) Symbol & option parsing
    //      now in BuildUnifiedSymbols.js
    //      (normalizeUnderlyingFromTradeSymbol, parseDottedOptionSymbol,
    //       normalizeSymbol)
    //
    //   F) Exercise / Assign + Action helpers
    //      now in BuildUnifiedExerciseAction.js
    //      (normalizeSpread, isExerciseOrAssignSpread,
    //       computeSignedAmountFromTrade, actionFromTosTrades,
    //       formatUnifiedSymbol)
    //
    // Later (after the main sheet loads) you will also see:
    //   • Mapping-sheet loaders (CusipMap, CorpActionStockMap, SplitAdjustments…)
    //   • Special-case parsers (DRIP, TDA fractional sells, RAD splits, etc.)
    //
    // =========================================================================

    // Group B decideIcRetag now lives in BuildUnifiedIcRetag.js
    // (createIcRetagHelpers). Wired in after the 6A maps exist.
    //
    // Group C sheet/field helpers live in BuildUnifiedSheetFields.js.

    const _fields = createSheetFieldHelpers({
      toStr: toStr,
    });
    const readSheetObjects = _fields.readSheetObjects;
    const getField = _fields.getField;
    const roundTo = _fields.roundTo;
    const pricesClose = _fields.pricesClose;

    // Date/time helpers now live in src/shared/Helpers.js
    // (normalizeDate, normalizeTime, normalizeTimeHHmmss, toDateObject).
    // Same names, so every call below is unchanged.

    function makeTradeMatchKey(
      Account,
      dateIso,
      timeHHmmss,
      symbol,
      absQty,
      price,
    ) {
      const al = toStr(Account).trim().toUpperCase();
      const sym = toStr(symbol).trim().toUpperCase();
      const t = toStr(timeHHmmss).trim(); // expects HHmmss
      const qty = String(absQty ?? "");
      const p = toNum(price);
      const pStr = isNaN(p) ? "" : p.toFixed(4);
      return [al, dateIso, t, sym, qty, pStr].join("|");
    }

    // Group E symbol helpers live in BuildUnifiedSymbols.js.
    const _syms = createSymbolHelpers({
      toStr: toStr,
    });
    const normalizeSymbol = _syms.normalizeSymbol;
    const normalizeUnderlyingFromTradeSymbol =
      _syms.normalizeUnderlyingFromTradeSymbol;
    const parseDottedOptionSymbol = _syms.parseDottedOptionSymbol;

    // Group F exercise/action helpers live in BuildUnifiedExerciseAction.js.
    // OPTIONS_CONTRACT_MULTIPLIER is declared in that file (global).
    const _ex = createExerciseActionHelpers({
      toStr: toStr,
    });
    const normalizeSpread = _ex.normalizeSpread;
    const isExerciseOrAssignSpread = _ex.isExerciseOrAssignSpread;
    const computeSignedAmountFromTrade = _ex.computeSignedAmountFromTrade;
    const actionFromTosTrades = _ex.actionFromTosTrades;
    const formatUnifiedSymbol = _ex.formatUnifiedSymbol;

    // =========================================================================
    // 3) READ INPUTS (ALL ACCOUNTS)
    //    Load the two working sheets that the rest of this function depends on:
    //      • TosTop    → fees, amounts, cash, corporate actions (source of truth for money)
    //      • TosTrades → trade legs / symbols / structure (source of truth for what was traded)
    //    Both accounts (DT + LT) are kept together; we no longer filter by account mode here.
    // =========================================================================
    const topSh = mustGetSheet("TosTop");
    const tradesSh = mustGetSheet("TosTrades");

    const tRead = pipelineTimingNow();

    // Read entire sheets (we now keep BOTH accounts together)
    let topRows = readSheetObjects(topSh);
    let tradesTbl = readSheetTable(tradesSh);

    // Basic sanity: make sure Account exists, because the combined workflow depends on it.
    if (!topRows.length) {
      importIssuesAdd(
        ctx,
        "WARN",
        "TosTop",
        "",
        "Rows",
        "0",
        "TosTop had no data rows.",
      );
    }
    importIssuesSetMetric(
      ctx,
      "SourceSheet",
      "TosTop + TosTrades (all accounts)",
    );
    importIssuesSetMetric(ctx, "TopRowsUsed", topRows.length);
    importIssuesSetMetric(ctx, "TradesRowsUsed", tradesTbl.rows.length);

    // =========================================================================
    // 4) LOAD SUPPORTING MAPPING SHEETS
    //    Reads optional helper sheets (CusipMap, CorpActionStockMap,
    //    SplitAdjustments, FakeDropOverride, etc.) that the matching logic uses.
    // =========================================================================
    const _maps = createMappingSheetHelpers({
      ctx: ctx,
      ss: ss,
      toStr: toStr,
      getSheetByNameLoose: getSheetByNameLoose,
      normalizeSymbol: normalizeSymbol,
    });
    const readCusipMap = _maps.readCusipMap;
    const readSplitAdjustments = _maps.readSplitAdjustments;
    const findCorpActionStockMapRow_ = _maps.findCorpActionStockMapRow_;
    const readCorpActionStockMap = _maps.readCorpActionStockMap;
    const parseCorpActionStockDescription_ =
      _maps.parseCorpActionStockDescription_;
    const readCorpActionsMap = _maps.readCorpActionsMap;
    const applyCorpActionsToTopRow = _maps.applyCorpActionsToTopRow;
    const applyCorpTokenRule = _maps.applyCorpTokenRule;
    const readFakeDropOverride = _maps.readFakeDropOverride;

    const splitAdjustments = readSplitAdjustments();
    const corpActionStockMap = readCorpActionStockMap();

    pipelineTimingLog(
      "buildUnifiedImportV3 read",
      tRead,
      "top=" + topRows.length + " trades=" + tradesTbl.rows.length,
    );
    const tQueues = pipelineTimingNow();

    // Group H special-case parsers live in BuildUnifiedSpecialParsers.js.
    const _parsers = createSpecialParsers({
      toStr: toStr,
    });
    const parseTosTopTradeDescription = _parsers.parseTosTopTradeDescription;
    const isDripTrdDescription = _parsers.isDripTrdDescription;
    const isTdaFractionalSellTrd = _parsers.isTdaFractionalSellTrd;
    const parseRadSplitDescription = _parsers.parseRadSplitDescription;
    const chooseCanonicalSplitCandidate_ =
      _parsers.chooseCanonicalSplitCandidate_;
    const buildCanonicalSplitRowFromCandidates_ =
      _parsers.buildCanonicalSplitRowFromCandidates_;
    const parseRadRemovalDescription = _parsers.parseRadRemovalDescription;

    // =========================================================================
    // 5) BUILD ENRICHMENT QUEUES FROM TosTop
    //    Takes the TRD (trade) rows from TosTop and builds in-memory queues
    //    keyed by timestamp / symbol / quantity so we can later attach the
    //    correct fees and amounts to each trade leg from TosTrades.
    // =========================================================================
    const corpMap = readCorpActionsMap();
    const cusipMap = readCusipMap();
    // ---- FakeDropOverride loader (add after readCusipMap) ----
    // Reads the "FakeDropOverride" sheet if it exists.
    // Sheet format: one column named "Ticker" (or "CUSIP").
    // Any ticker or CUSIP in this list will be DROPPED even if it exists in CusipMap.
    // Use this for known-bogus paper trades like NVDS whose CUSIP happens to be in CusipMap.

    const fakeDropOverrideSet = readFakeDropOverride();
    // ---- END FakeDropOverride loader ----

    const missingCusipSeenTosTop = {};
    const missingCusipSeenTosTrades = {};
    const missingTradeEnrichmentSeen = {};
    const missingTradeEnrichmentDataGapSeen = {};
    const icPackageEnrichCache = {};
    //  de-dupe "weak enrichment match" warnings so we don't log 4x per multi-leg trade
    const weakEnrichmentMatchSeen = {};

    // Two indexes:
    // - topTradeQueueByKey: exact match on (date|time|symbol|absQty|price) for deterministic pulls
    // - topTradeQueueByDateTime: all TRD rows in the same minute bucket for safe fallbacks
    const topTradeQueueByKey = {};
    const topTradeQueueByDateTime = {};

    // Group A: same helper names as before, bound to THESE queue objects.
    const _enrich = createTopTradeEnrichmentHelpers({
      topTradeQueueByKey: topTradeQueueByKey,
      topTradeQueueByDateTime: topTradeQueueByDateTime,
      toStr: toStr,
      makeTradeMatchKey: makeTradeMatchKey,
    });
    const previewTopCandidates = _enrich.previewTopCandidates;
    const removeFromExactIndex = _enrich.removeFromExactIndex;
    const pullPartialFillEnrichment_ = _enrich.pullPartialFillEnrichment_;
    const pullTopTradeEnrichment = _enrich.pullTopTradeEnrichment;
    const pullTopTradeEnrichmentButterfly =
      _enrich.pullTopTradeEnrichmentButterfly;
    const pullTopTradeEnrichmentIronCondor =
      _enrich.pullTopTradeEnrichmentIronCondor;

    // Diagnostics (super helpful when you think “there should be a TRD row in that minute”)
    const topTrdRowsSeenByDateTime = {}; // counts ALL TosTop TRD rows per minute (parse success or fail)
    const topTrdParseFailSeen = {}; // de-dupe parse-fail logging by description
    let topTrdRowsSeen = 0;
    let topTrdRowsQueued = 0;
    let topTrdRowsParseFailed = 0;

    for (let i = 0; i < topRows.length; i++) {
      const r = topRows[i];

      // Account is per-row because TosTop holds both DT and LT together.
      const Account = toStr(getField(r, "Account")).trim().toUpperCase();
      const type = toStr(getField(r, "TYPE")).trim().toUpperCase();
      if (type !== "TRD") continue;

      topTrdRowsSeen++;

      // Skip rows that have their own standalone emitter (no TosTrades partner exists).
      // Both DRIP BOT rows and TDA fractional SELL rows must be excluded from the
      // enrichment queue — the queue only services rows that have a TosTrades partner.
      const descRawForSkipCheck = toStr(getField(r, "DESCRIPTION")).trim();
      const skipCheckParsed = parseTosTopTradeDescription(descRawForSkipCheck);
      const skipCheckQty = skipCheckParsed ? skipCheckParsed.absQty : null;
      const skipCheckAmount = toNum(getField(r, "AMOUNT"));
      // Compute dateIso and Account here so the era-guard in isTdaFractionalSellTrd can use them.
      const skipCheckDateIso = normalizeDate(getField(r, "DATE"));
      const skipCheckAccount = toStr(getField(r, "Account"))
        .trim()
        .toUpperCase();
      if (isDripTrdDescription(descRawForSkipCheck, skipCheckQty)) continue;
      if (
        isTdaFractionalSellTrd(
          descRawForSkipCheck,
          skipCheckQty,
          skipCheckAmount,
          skipCheckDateIso,
          skipCheckAccount,
        )
      )
        continue;

      const dateIso = normalizeDate(getField(r, "DATE"));

      // TosTop TIME is HHmmss text now (ex: "230208")
      let timeHHmmss = normalizeTimeHHmmss(getField(r, "TIME"));
      if (!timeHHmmss) timeHHmmss = "000000";

      // Minute bucket key (HHmm) used for “closest in time within that minute”.
      const timeHHmm = normalizeTime(timeHHmmss);

      // Minute bucket key is computed BEFORE parsing so we can count “seen” rows even on parse-fail.
      const dtMinuteKey = [Account, dateIso, timeHHmm].join("|");
      topTrdRowsSeenByDateTime[dtMinuteKey] =
        Number(topTrdRowsSeenByDateTime[dtMinuteKey] || 0) + 1;

      // Timestamp used for closest-in-time selection inside the minute bucket
      const ts = toDateObject(dateIso, timeHHmmss);
      if (
        !(ts instanceof Date) ||
        isNaN(ts.getTime()) ||
        ts.getFullYear() < 2000
      ) {
        importIssuesAdd(
          ctx,
          "BADTOSTOPDATETIME",
          "TosTop",
          i + 2,
          "DATETIME",
          JSON.stringify({
            DATE: getField(r, "DATE"),
            TIME: getField(r, "TIME"),
          }),
          "Cannot build Timestamp from TosTop DATETIME",
        );
        continue;
      }

      // TRD parsing comes ONLY from DESCRIPTION (corp rules do NOT apply to TRD rows per your rule).
      const descRaw = toStr(getField(r, "DESCRIPTION")).trim();
      const parsed = parseTosTopTradeDescription(descRaw);

      // If we can’t parse symbol/qty/price, we cannot enrich trades reliably.
      if (
        !parsed ||
        !parsed.symbol ||
        parsed.absQty === null ||
        parsed.absQty === undefined
      ) {
        topTrdRowsParseFailed++;

        // De-dupe the issue logging so a repeated description doesn’t spam your Issues sheet.
        const dedupeKey = descRaw || "(blank desc) row " + (i + 2);
        if (!topTrdParseFailSeen[dedupeKey]) {
          topTrdParseFailSeen[dedupeKey] = true;
          importIssuesAdd(
            ctx,
            "TOPTRD_PARSE_FAIL",
            "TosTop",
            i + 2,
            "DESCRIPTION",
            descRaw,
            JSON.stringify({
              message:
                "Could not parse TosTop TRD DESCRIPTION into symbol/qty/price for enrichment matching.",
              parsed: parsed,
            }),
          );
        }
        continue;
      }

      // IMPORTANT: topSym is the match symbol derived from TosTop DESCRIPTION (parsed.symbol),
      // normalized the same way we normalize TosTrades symbols for enrichment matching.
      let topSym = normalizeUnderlyingFromTradeSymbol(parsed.symbol);

      // CUSIP mapping: if match symbol is CUSIP-like, map to ticker via CusipMap.
      // (normalizeUnderlyingFromTradeSymbol preserves CUSIPs; this is the “map to ticker” step)
      if (looksLikeCusip(topSym) && /\d/.test(normalizeCusip(topSym))) {
        const mapped = cusipMap[topSym];
        if (mapped) {
          topSym = mapped;
        } else if (!missingCusipSeenTosTop[topSym]) {
          missingCusipSeenTosTop[topSym] = true;
          importIssuesAdd(
            ctx,
            "MISSINGCUSIPMAP",
            "TosTop",
            i + 2,
            "DESCRIPTION",
            topSym,
            "CUSIP-like token/symbol derived from TosTop TRD DESCRIPTION not found in CusipMap.",
          );
        }
      }

      // Exact key: second-precision.
      const keyExact = makeTradeMatchKey(
        Account,
        dateIso,
        timeHHmmss,
        topSym,
        parsed.absQty,
        parsed.price,
      );

      const item = {
        miscFees: getField(r, ["Misc Fees", "MiscFees"]),
        feesComm: getField(r, [
          "Commissions & Fees",
          "Commissions Fees",
          "Commissions and Fees",
        ]),
        amount: getField(r, "AMOUNT"),

        // Used for closest-in-time selection
        topTs: ts,
        topTimeHHmmss: timeHHmmss,
        topMinuteKey: dtMinuteKey,

        // Debug fields
        topDesc: descRaw,
        topSym: topSym,
        topAbsQty: parsed.absQty,
        topPrice: parsed.price,
        topExactKey: keyExact,
      };

      if (!topTradeQueueByKey[keyExact]) topTradeQueueByKey[keyExact] = [];
      topTradeQueueByKey[keyExact].push(item);

      if (!topTradeQueueByDateTime[dtMinuteKey])
        topTradeQueueByDateTime[dtMinuteKey] = [];
      topTradeQueueByDateTime[dtMinuteKey].push(item);

      topTrdRowsQueued++;
    }

    // Metrics that tell you immediately if you're failing because TosTop parsing failed
    importIssuesSetMetric(ctx, "TopTrdRowsSeen", topTrdRowsSeen);
    importIssuesSetMetric(ctx, "TopTrdRowsQueued", topTrdRowsQueued);
    importIssuesSetMetric(ctx, "TopTrdRowsParseFailed", topTrdRowsParseFailed);

    // Group A enrichment helpers were bound just after the queues above
    // (createTopTradeEnrichmentHelpers in BuildUnifiedEnrichment.js).

    pipelineTimingLog(
      "buildUnifiedImportV3 queues",
      tQueues,
      "topTrdRowsSeen=" + topTrdRowsSeen,
    );
    const tPrepass = pipelineTimingNow();

    // =========================================================================
    // 6) CONVERT TosTrades ROWS → UNIFIED TRADE ROWS
    //    This is the main matching loop. Each trade leg from TosTrades is
    //    paired with the best matching fee/amount row from the queues built
    //    in step 5. Corporate-action and special cases are also handled here.
    // =========================================================================
    const unifiedTrades = [];
    const enrichmentConsumedByGroupKey = {};

    // ----------------------------
    // 6A) Lifecycle spread retagging (IC inheritance + inventory guard)
    // ----------------------------

    const CANONICAL_IC = "IRON CONDOR";

    // bundleKey -> { legId:true }
    const lifecycleBundleLegSetByKey = {};
    // bundleKey -> metadata (spreadRaw, ts, etc.)
    const lifecycleBundleMetaByKey = {};

    // bundleKey -> canonical spread (IRON CONDOR) when inherited
    const canonicalSpreadByLifecycleBundleKey = {};
    // bundleKey -> chosen OPEN info (for INFO logging)
    const canonicalSpreadChoiceByBundleKey = {};

    const openIcQtyByLegKey = {};

    // INFO logging de-dupe + metrics
    const spreadRetagInfoLoggedByBundleKey = {};
    const icRetagMetrics = {
      spreadRetaggedRowsCount: 0,
      spreadRetagSkippedNoOpenIcCount: 0,
    };

    // Group B: same helper names as before, bound to THESE maps.
    const _ic = createIcRetagHelpers({
      canonicalSpreadByLifecycleBundleKey: canonicalSpreadByLifecycleBundleKey,
      lifecycleBundleLegSetByKey: lifecycleBundleLegSetByKey,
      openIcQtyByLegKey: openIcQtyByLegKey,
      icRetagMetrics: icRetagMetrics,
      CANONICAL_IC: CANONICAL_IC,
      toStr: toStr,
      roundTo: roundTo,
    });
    const decideIcRetag = _ic.decideIcRetag;
    const isIronCondorSpread = _ic.isIronCondorSpread;
    const isRetagCandidateOriginalSpread = _ic.isRetagCandidateOriginalSpread;
    const isCloseLikeBundle = _ic.isCloseLikeBundle;
    const makeLifecycleBundleKey = _ic.makeLifecycleBundleKey;
    const makeOpenIndexKey = _ic.makeOpenIndexKey;
    const legIdFromTypeStrike = _ic.legIdFromTypeStrike;
    const setToSortedArray = _ic.setToSortedArray;
    const isSubset = _ic.isSubset;
    const makeIcLegQtyKey = _ic.makeIcLegQtyKey;
    const normalizeExpKey = _ic.normalizeExpKey;
    const isMultiTypeSpread = _ic.isMultiTypeSpread;

    // Pre-pass: compute a per-group "net premium" for multi-leg spreads from TosTrades legs.
    // This lets enrichment match TosTop TRD rows that show the *spread* price (e.g. @.53),
    // even when TosTrades Net Price is non-numeric (e.g. "DEBIT"/"CREDIT").
    const spreadNetAbsByGroupKey = {}; // groupKey -> abs(netPremiumPerFill)  (not sum)
    const spreadNetSignedByGroupKey = {}; // groupKey -> signed sum of leg prices (grouped)
    const spreadBuyRowsByGroupKey = {}; // groupKey -> count of BUY leg rows in this groupKey
    const spreadSellRowsByGroupKey = {}; // groupKey -> count of SELL leg rows in this groupKey
    const spreadFillCountByGroupKey = {}; // groupKey -> expected number of fills (min(buyRows, sellRows))

    const butterflyNetAbsByKey = {}; // bfKey -> abs(net per 1-lot strategy)
    const butterflyNetSignedByKey = {}; // bfKey -> signed net per 1-lot strategy
    const butterflyLegRowsByKey = {}; // bfKey -> count legs seen (diagnostic)

    // DIAGONAL/CALENDAR: legs have different expirations, so we aggregate by a coarse key
    // and later attach a sorted expiration signature.

    const diagCalAggByCoarseKey = {}; // coarseKey -> { expSet:{}, signedSum, buyN, sellN }
    const diagCalExpSigByCoarseKey = {}; // coarseKey -> "yyyy-mm-dd,yyyy-mm-dd" (sorted unique)

    // CUSTOM sanity check: if an IC-like CUSTOM group spans multiple expirations, warn (likely mis-entry).

    const customExpAggByCoarseKey = {}; // coarseKey -> { expSet:{}, callN, putN, legN }
    const customExpSigByCoarseKey = {}; // coarseKey -> "yyyy-mm-dd,yyyy-mm-dd" (sorted unique)
    const customMixedExpWarnSeen = {}; // coarseKey -> true (de-dupe warnings)

    for (let i = 0; i < tradesTbl.rows.length; i++) {
      const row = tradesTbl.rows[i];

      // Account is per-row because TosTop holds both DT and LT together.

      const Account = toStr(cell(row, tradesTbl.idx, "Account"))
        .trim()
        .toUpperCase();

      const execDtRaw = cell(row, tradesTbl.idx, "Exec Time");
      if (
        !(execDtRaw instanceof Date) ||
        isNaN(execDtRaw.getTime()) ||
        execDtRaw.getFullYear() < 2000
      )
        continue;

      // Exec Time is authoritative and should be a Date from TosTrades.
      // Trades are typically minute precision from the CSV, so seconds often end up 00.
      const ts = new Date(execDtRaw.getTime());

      const dateIso = normalizeDate(ts);

      // - timeHHmm stays available for minute bucketing.
      // - timeHHmmss becomes the *trade* time key (but will be HHmm00 when CSV has no seconds).
      const timeHHmmss = normalizeTimeHHmmss(ts); // HHmmss

      const symRawCell = cell(row, tradesTbl.idx, "Symbol");
      let sym = normalizeUnderlyingFromTradeSymbol(symRawCell);

      // Apply CusipMap to the match symbol (same idea as your main loop)
      const cusipKey = normalizeCusip(sym);
      if (looksLikeCusip(cusipKey) && cusipMap[cusipKey]) {
        sym = normalizeSymbol(cusipMap[cusipKey]);
      }

      const spreadRaw = toStr(cell(row, tradesTbl.idx, "Spread"))
        .trim()
        .toUpperCase();
      const spread = normalizeSpread(spreadRaw);
      const posEffect = toStr(cell(row, tradesTbl.idx, "Pos Effect"))
        .trim()
        .toUpperCase();
      const side = toStr(cell(row, tradesTbl.idx, "Side"))
        .trim()
        .toUpperCase();

      // Only compute for multi-leg spreads (skip STOCK/SINGLE/EXERCISE/ASSIGN and blank spreads)
      const isExerciseOrAssign = isExerciseOrAssignSpread(spread);
      // Treat EXERCISE/ASSIGN as single-like so they do NOT get grouped into spread nets.
      const isStockOrSingle =
        spread === "STOCK" || spread === "SINGLE" || isExerciseOrAssign;
      const isMultiLegSpread = !isStockOrSingle && spread !== "";
      if (!isMultiLegSpread) continue;

      const qtyAbs = Math.abs(toNum(cell(row, tradesTbl.idx, "Qty")));
      const price = toNum(cell(row, tradesTbl.idx, "Price"));
      if (isNaN(qtyAbs) || qtyAbs <= 0) continue;
      if (isNaN(price)) continue;

      // Allow overrides for dotted option symbols (MANUAL legs), e.g. ".QQQ230309C305"
      let typeKey = toStr(cell(row, tradesTbl.idx, "Type"))
        .trim()
        .toUpperCase();

      const expRaw = cell(row, tradesTbl.idx, "Exp");
      let expKey = normalizeExpKey(expRaw);

      const dottedOpt = parseDottedOptionSymbol(
        cell(row, tradesTbl.idx, "Symbol"),
      );
      if (dottedOpt) {
        const expKeyIsBlank = !toStr(expKey).trim();
        const typeIsManual = !typeKey || typeKey === "MANUAL";

        if (typeIsManual) typeKey = dottedOpt.optType; // CALL/PUT
        if (expKeyIsBlank) expKey = normalizeExpKey(dottedOpt.expDate);
      }

      // --- Lifecycle leg-set capture (captures SINGLE/VERTICAL too) ---
      // We only capture “option-like” legs where Strike is numeric and Type is CALL/PUT.
      const strikeRawForLife = cell(row, tradesTbl.idx, "Strike");
      const strikeNumForLife = toNum(strikeRawForLife);

      if (
        isFinite(strikeNumForLife) &&
        (typeKey === "CALL" || typeKey === "PUT") &&
        String(expKey || "").trim()
      ) {
        const bundleKey = makeLifecycleBundleKey(
          Account,
          ts,
          sym,
          expKey,
          posEffect,
        );

        if (!lifecycleBundleLegSetByKey[bundleKey])
          lifecycleBundleLegSetByKey[bundleKey] = {};
        lifecycleBundleLegSetByKey[bundleKey][
          legIdFromTypeStrike(typeKey, strikeNumForLife)
        ] = true;

        if (!lifecycleBundleMetaByKey[bundleKey]) {
          lifecycleBundleMetaByKey[bundleKey] = {
            Account: Account,
            ts: ts,
            sym: sym,
            expKey: expKey,
            posEffect: posEffect,
            spreadRaw: spread,
          };
        }
      }
      // --- End lifecycle capture ---

      //  Track expirations for CUSTOM groups so we can warn on mixed-exp "IC-like" CUSTOM trades.
      if (spread === "CUSTOM") {
        const customCoarseKey = [
          Account,
          dateIso,
          timeHHmmss,
          sym,
          spread,
          qtyAbs,
          posEffect,
        ].join("|");

        if (!customExpAggByCoarseKey[customCoarseKey]) {
          customExpAggByCoarseKey[customCoarseKey] = {
            expSet: {},
            callN: 0,
            putN: 0,
            legN: 0,
          };
        }

        const agg = customExpAggByCoarseKey[customCoarseKey];

        if (toStr(expKey).trim()) agg.expSet[expKey] = true;
        if (typeKey === "CALL") agg.callN++;
        if (typeKey === "PUT") agg.putN++;
        agg.legN++;
      }

      const sign = side === "SELL" ? 1 : side === "BUY" ? -1 : 0;
      if (sign === 0) continue;

      // --- BUTTERFLY net (per 1-lot strategy) ---
      // Keep butterfly separate; do NOT let it fall into the generic spread-net logic,
      // because butterfly legs are 1-2-1 and will distort generic spread grouping.
      if (spread === "BUTTERFLY") {
        const bfKey = [
          Account,
          dateIso,
          timeHHmmss,
          sym,
          spread,
          posEffect,
          expKey,
          typeKey,
        ].join("|");

        // qtyAbs is contracts; butterfly middle leg has qtyAbs=2 for a 1-lot strategy.
        // Normalize to "strategy units" by dividing by 2 when qtyAbs is even (middle leg),
        // otherwise treat as 1-unit legs.
        const strategyUnits = qtyAbs % 2 === 0 ? qtyAbs / 2 : qtyAbs;

        const prev = Number(butterflyNetSignedByKey[bfKey] || 0);
        butterflyNetSignedByKey[bfKey] = prev + sign * price * strategyUnits;

        butterflyLegRowsByKey[bfKey] =
          Number(butterflyLegRowsByKey[bfKey] || 0) + 1;

        continue; // IMPORTANT: skip generic spread-net grouping for BUTTERFLY
      }

      // --- DIAGONAL / CALENDAR ---
      // These spreads have legs with different expirations, so we aggregate by a coarse key
      // and later attach an expiration signature (sorted unique expirations).
      if (spread === "DIAGONAL" || spread === "CALENDAR") {
        const coarseKey = [
          Account,
          dateIso,
          timeHHmmss,
          sym,
          spread,
          qtyAbs,
          posEffect,
          typeKey,
        ].join("|");

        if (!diagCalAggByCoarseKey[coarseKey]) {
          diagCalAggByCoarseKey[coarseKey] = {
            expSet: {},
            signedSum: 0,
            buyN: 0,
            sellN: 0,
          };
        }

        const agg = diagCalAggByCoarseKey[coarseKey];
        agg.expSet[expKey] = true;
        agg.signedSum += sign * price;
        if (side === "BUY") agg.buyN++;
        if (side === "SELL") agg.sellN++;

        continue; // IMPORTANT: do not add to generic spread maps yet
      }

      // --- Generic multi-leg spreads (VERTICAL, CONDOR, IC, etc.) ---
      // IMPORTANT:
      // - Still include expKey to prevent same-minute collisions across expirations (your SPX case).
      // - BUT for multi-type spreads (IRON CONDOR / STRANGLE / STRADDLE), do NOT include typeKey,
      //   so CALL+PUT legs stay in ONE strategy group and we compute the strategy net premium correctly.
      const spreadTypeKey = isMultiTypeSpread(spread) ? "MIXED" : typeKey;
      const groupKey = [
        Account,
        dateIso,
        timeHHmmss,
        sym,
        spread,
        qtyAbs,
        posEffect,
        expKey,
        spreadTypeKey,
      ].join("|");

      spreadNetSignedByGroupKey[groupKey] =
        Number(spreadNetSignedByGroupKey[groupKey] || 0) + sign * price;

      if (side === "BUY") {
        spreadBuyRowsByGroupKey[groupKey] =
          Number(spreadBuyRowsByGroupKey[groupKey] || 0) + 1;
      } else if (side === "SELL") {
        spreadSellRowsByGroupKey[groupKey] =
          Number(spreadSellRowsByGroupKey[groupKey] || 0) + 1;
      }
    }

    //  Some spreads (like IRON CONDOR) contain BOTH CALL and PUT legs but have ONE TosTop TRD row.
    // For these, we must group across typeKey so we compute the strategy net (e.g. .32) and only
    // attempt enrichment once.

    //  Finalize a stable exp signature for each CUSTOM coarse group.
    Object.keys(customExpAggByCoarseKey).forEach((k) => {
      const agg = customExpAggByCoarseKey[k];
      const exps = Object.keys(agg.expSet || {})
        .filter((s) => s)
        .sort();
      customExpSigByCoarseKey[k] = exps.join(",");
    });

    // Finalize DIAGONAL/CALENDAR aggregates into the same spreadNet* maps using an exp signature.
    Object.keys(diagCalAggByCoarseKey).forEach((coarseKey) => {
      const agg = diagCalAggByCoarseKey[coarseKey];
      const exps = Object.keys(agg.expSet || {})
        .filter(Boolean)
        .sort();
      const expSig = exps.join(","); // e.g. "2023-10-19,2023-10-20"
      diagCalExpSigByCoarseKey[coarseKey] = expSig;

      const groupKey = coarseKey + "|" + expSig;

      spreadNetSignedByGroupKey[groupKey] =
        Number(spreadNetSignedByGroupKey[groupKey] || 0) +
        Number(agg.signedSum || 0);
      spreadBuyRowsByGroupKey[groupKey] =
        Number(spreadBuyRowsByGroupKey[groupKey] || 0) + Number(agg.buyN || 0);
      spreadSellRowsByGroupKey[groupKey] =
        Number(spreadSellRowsByGroupKey[groupKey] || 0) +
        Number(agg.sellN || 0);
    });

    Object.keys(butterflyNetSignedByKey).forEach((k) => {
      const v = Number(butterflyNetSignedByKey[k]);
      butterflyNetAbsByKey[k] = isFinite(v) ? Math.abs(v) : 0;
    });

    importIssuesSetMetric(
      ctx,
      "ButterflyNetComputed",
      Object.keys(butterflyNetAbsByKey).length,
    );

    // Finalize abs(net) map
    //
    // IMPORTANT:
    // For multi-leg strategies, buyN/sellN are LEG COUNTS, not "fill counts".
    // Example: IRON CONDOR has 2 BUY legs + 2 SELL legs, but the strategy net is still the full sumSigned (e.g. .32),
    // not half (.16). So we should NOT divide by min(buyN, sellN).
    Object.keys(spreadNetSignedByGroupKey).forEach((k) => {
      const sumSigned = Number(spreadNetSignedByGroupKey[k] || 0);
      const absSum = Math.abs(sumSigned);

      // Keep expectedFillCount = 1 for now (prevents accidental multi-pulls and incorrect fee summing).
      // If you later want true multi-fill aggregation, it needs a different approach than leg counts.
      spreadFillCountByGroupKey[k] = 1;

      spreadNetAbsByGroupKey[k] = absSum;
    });

    importIssuesSetMetric(
      ctx,
      "SpreadGroupsNetComputed",
      Object.keys(spreadNetAbsByGroupKey).length,
    );

    // ----------------------------
    // 6B) Infer IC inheritance mapping (subset of prior IC OPEN)
    // ----------------------------

    const openIcBundlesByAccSymExp = {}; // idxKey -> array of {ts, legSetObj}

    Object.keys(lifecycleBundleMetaByKey).forEach((bundleKey) => {
      const meta = lifecycleBundleMetaByKey[bundleKey];
      const pe = String(meta.posEffect || "").toUpperCase();

      if (!pe.includes("OPEN")) return;
      if (!isIronCondorSpread(meta.spreadRaw)) return; // exact match only

      const idxKey = makeOpenIndexKey(meta.Account, meta.sym, meta.expKey);
      if (!openIcBundlesByAccSymExp[idxKey])
        openIcBundlesByAccSymExp[idxKey] = [];

      openIcBundlesByAccSymExp[idxKey].push({
        ts: meta.ts,
        legSetObj: lifecycleBundleLegSetByKey[bundleKey] || {},
      });
    });

    // Sort opens oldest->newest so we can pick nearest prior open.
    Object.keys(openIcBundlesByAccSymExp).forEach((k) => {
      openIcBundlesByAccSymExp[k].sort(
        (a, b) => a.ts.getTime() - b.ts.getTime(),
      );
    });

    Object.keys(lifecycleBundleMetaByKey).forEach((bundleKey) => {
      const meta = lifecycleBundleMetaByKey[bundleKey];

      // Only consider CLOSE-like bundles.
      if (!isCloseLikeBundle(meta.posEffect, meta.spreadRaw)) return;

      // Mixed TOS labels: one same-second close leg can already say
      // IRON CONDOR while its sibling says VERTICAL. After normalizeExpKey
      // they share this bundleKey. First-wins meta.spreadRaw is then IC,
      // and skipping here left the VERTICAL row untagged
      // (DT AAPL 182.5P 2023-08-04 09:07:07, DT NVDA 430C 2023-08-15 08:33:56).
      // Still set the canonical map so decideIcRetag can retag the
      // VERTICAL sibling. Rows whose original spread is already IC are
      // not retag candidates and stay IC.

      const idxKey = makeOpenIndexKey(meta.Account, meta.sym, meta.expKey);
      const opens = openIcBundlesByAccSymExp[idxKey] || [];
      if (!opens.length) return;

      const closeLegSetSorted = setToSortedArray(
        lifecycleBundleLegSetByKey[bundleKey] || {},
      );
      if (!closeLegSetSorted.length) return;

      let chosen = null;
      for (let i = opens.length - 1; i >= 0; i--) {
        const o = opens[i];
        if (o.ts.getTime() > meta.ts.getTime()) continue;
        if (isSubset(closeLegSetSorted, o.legSetObj)) {
          chosen = o;
          break;
        }
      }

      if (chosen) {
        canonicalSpreadByLifecycleBundleKey[bundleKey] = CANONICAL_IC;
        canonicalSpreadChoiceByBundleKey[bundleKey] = { openTs: chosen.ts };
      }
    });

    pipelineTimingLog("buildUnifiedImportV3 prepass", tPrepass);
    const tMain = pipelineTimingNow();

    for (let i = 0; i < tradesTbl.rows.length; i++) {
      const row = tradesTbl.rows[i];

      // Account is per-row because TosTrades holds both DT and LT together.
      const Account = toStr(cell(row, tradesTbl.idx, "Account"))
        .trim()
        .toUpperCase();

      // Exec Time is authoritative and should be a Date (from your TosTrades push step).
      const execDtRaw = cell(row, tradesTbl.idx, "Exec Time");

      if (
        !(execDtRaw instanceof Date) ||
        isNaN(execDtRaw.getTime()) ||
        execDtRaw.getFullYear() < 2000
      ) {
        importIssuesAdd(
          ctx,
          "BAD_EXEC_TIME",
          "TosTrades",
          i + 2,
          "Exec Time",
          execDtRaw,
          JSON.stringify({
            message: "Expected a real DateTime in TosTrades Exec Time",
            rowPreview: previewRow12(row),
          }),
        );
        continue;
      }

      // Exec Time is authoritative and should be a Date from TosTrades.
      //  preserve seconds (and milliseconds if present, though TOS exports typically use seconds).
      const ts = new Date(execDtRaw.getTime());

      const dateIso = normalizeDate(ts);

      //
      // - timeHHmm stays available for TosTop minute-based enrichment fallback.
      // - timeHHmmss becomes the *trade* time key and the display "Time" value.
      const timeHHmm = normalizeTime(ts); // HHmm
      const timeHHmmss = normalizeTimeHHmmss(ts); // HHmmss

      const symRawCell = cell(row, tradesTbl.idx, "Symbol");
      const symRawNorm = normalizeSymbol(symRawCell); // for human trace
      let sym = normalizeUnderlyingFromTradeSymbol(symRawCell); // for matching + unifiedSymbol when possible

      // CUSIP mapping should apply to the normalized underlying as well
      const cusipKey = normalizeCusip(sym);

      if (looksLikeCusip(cusipKey)) {
        const mapped = cusipMap[cusipKey];
        if (mapped) {
          sym = normalizeSymbol(mapped);
        } else if (!missingCusipSeenTosTrades[cusipKey]) {
          missingCusipSeenTosTrades[cusipKey] = true;
          importIssuesAdd(
            ctx,
            "MISSING_CUSIP_MAP",
            "TosTrades",
            i + 2,
            "Symbol",
            cusipKey,
            JSON.stringify({
              message: "CUSIP not found in CusipMap after normalization",
              originalSymbolCell: symRawCell,
              rowPreview: previewRow12(row),
            }),
          );
        }
      }

      // This is the symbol used for enrichment matching against TosTop TRD rows.
      const symForMatch = sym;

      const spreadRaw = toStr(cell(row, tradesTbl.idx, "Spread"))
        .trim()
        .toUpperCase();

      const spread = normalizeSpread(spreadRaw);
      const isExerciseOrAssign = isExerciseOrAssignSpread(spread);
      const qtyAbs = Math.abs(toNum(cell(row, tradesTbl.idx, "Qty")));
      const price = toNum(cell(row, tradesTbl.idx, "Price"));

      const netPriceRaw = cell(row, tradesTbl.idx, "Net Price");
      const netPriceNum = toNum(netPriceRaw);
      const hasNumericNet = !isNaN(netPriceNum);

      //  Define these early so later code (warnings, computed net logic, enrichment gating)
      // can safely reference them without TDZ (temporal dead zone) errors.
      const isStockOrSingle =
        spread === "STOCK" || spread === "SINGLE" || isExerciseOrAssign;
      const isMultiLegSpread = !isStockOrSingle && spread !== "";
      // These MUST be defined before groupKey and before gapKey logging.
      const orderType = toStr(cell(row, tradesTbl.idx, "Order Type"))
        .trim()
        .toUpperCase();
      const posEffect = toStr(cell(row, tradesTbl.idx, "Pos Effect"))
        .trim()
        .toUpperCase();

      // NEW WARNING: IC-like CUSTOM group with mixed expirations (likely broker mis-entry / edited trade).
      // Runs here (main loop) because:
      // - symForMatch exists
      // - customExpSigByCoarseKey has already been finalized after the pre-pass
      if (spread === "CUSTOM") {
        const customCoarseKey = [
          Account,
          dateIso,
          timeHHmmss,
          symForMatch,
          spread,
          qtyAbs,
          posEffect,
        ].join("|");
        const expSig = customExpSigByCoarseKey[customCoarseKey] || "";
        const agg = customExpAggByCoarseKey[customCoarseKey];

        const isIClike = !!(
          agg &&
          agg.legN >= 4 &&
          agg.callN > 0 &&
          agg.putN > 0
        );
        const hasMultipleExp = expSig.includes(",");

        if (
          isIClike &&
          hasMultipleExp &&
          !customMixedExpWarnSeen[customCoarseKey]
        ) {
          customMixedExpWarnSeen[customCoarseKey] = true;

          importIssuesAdd(
            ctx,
            "WARN_CUSTOM_MIXED_EXPIRATIONS",
            "TosTrades",
            i + 2,
            "Exp",
            expSig,
            JSON.stringify({
              message:
                "CUSTOM spread appears IC-like but has mixed expirations across legs. Verify statement; repair before downstream.",
              customCoarseKey: customCoarseKey,
              expSig: expSig,
              legCounts: agg
                ? { legN: agg.legN, callN: agg.callN, putN: agg.putN }
                : null,
              rowPreview: previewRow12(row),
            }),
          );
        }
      }

      // Allow overrides for dotted option symbols (MANUAL legs)
      let typeKey = toStr(cell(row, tradesTbl.idx, "Type"))
        .trim()
        .toUpperCase();
      const side = toStr(cell(row, tradesTbl.idx, "Side"))
        .trim()
        .toUpperCase();

      let exp = cell(row, tradesTbl.idx, "Exp");
      let strike = cell(row, tradesTbl.idx, "Strike");

      //  If symbol looks like ".QQQ230309C305", treat it as a real option leg.
      const dottedOpt = parseDottedOptionSymbol(
        cell(row, tradesTbl.idx, "Symbol"),
      );
      if (dottedOpt) {
        // Only override when the sheet is missing structured fields (common for MANUAL legs)
        const expIsBlank = !(exp instanceof Date) && toStr(exp).trim() === "";
        const strikeIsBlank =
          toStr(strike).trim() === "" || isNaN(toNum(strike));
        const typeIsManual = !typeKey || typeKey === "MANUAL";

        if (typeIsManual) typeKey = dottedOpt.optType;
        if (expIsBlank) exp = dottedOpt.expDate;
        if (strikeIsBlank) strike = dottedOpt.strike;
      }

      // EXERCISE / ASSIGN: TosTrades often already has the share movement
      // (BUY +100 STOCK) next to the option close. We emit ONE synthetic
      // stock row from the option leg and put TosTop cash on that row.
      // Skip this companion stock line so Schwab Import does not double the shares.
      if (isExerciseOrAssign && typeKey === "STOCK") {
        continue;
      }

      // Baseline match price (used for TosTop enrichment matching).
      // STOCK/SINGLE/EXERCISE/ASSIGN: prefer numeric Net Price; otherwise use leg Price.
      // Spreads: start with leg Price; we may override later with computedSpreadNetAbs.
      let matchPriceForPull = price;
      if (spread === "STOCK" || spread === "SINGLE" || isExerciseOrAssign) {
        matchPriceForPull = hasNumericNet ? netPriceNum : price;
      } else {
        matchPriceForPull = price;
      }
      // Butterfly: override match price using strategy net (per 1-lot) when available.
      if (spread === "BUTTERFLY") {
        const expKey =
          exp instanceof Date
            ? Utilities.formatDate(
                exp,
                Session.getScriptTimeZone(),
                "yyyy-MM-dd",
              )
            : toStr(exp).trim();

        const bfKey = [
          Account,
          dateIso,
          timeHHmmss,
          sym,
          spread,
          posEffect,
          expKey,
          typeKey,
        ].join("|");
        const bfNetAbs = butterflyNetAbsByKey[bfKey];

        if (
          typeof bfNetAbs === "number" &&
          isFinite(bfNetAbs) &&
          bfNetAbs > 0
        ) {
          matchPriceForPull = bfNetAbs;
        }
      }

      // Group key - MUST match the pre-pass keys (spreadNetAbsByGroupKey, spreadFillCountByGroupKey...)
      // IMPORTANT: Trades Exec Time can be minute-only (HH:mm -> seconds = 00), so multiple different spreads
      // can land in the same minute with the same timeHHmmss (e.g. 082800).
      //
      // Therefore, for multi-leg options spreads we MUST include Exp (and Type) in the group identity,
      // otherwise different expirations collide and you get summed fees/amount (-235) on one spread.
      // Must use the same normalizeExpKey as the 6A/6B IC maps.
      // Date vs "4 Aug 23" vs "4 August 23" vs "18-Aug-23" must be one key.
      const expKey = normalizeExpKey(exp);

      // ---------------------------- IC retagging — now one clean call to the new helper ----------------------------
      const spreadOriginal = spread;
      const lifeBundleKey = makeLifecycleBundleKey(
        Account,
        ts,
        symForMatch,
        expKey,
        posEffect,
      );

      let spreadOut = decideIcRetag(
        Account,
        ts,
        symForMatch,
        expKey,
        posEffect,
        spreadOriginal,
        lifeBundleKey,
      );

      // --- UPDATE: IC inventory counts by leg (uses spreadOut) ---
      if (spreadOut === CANONICAL_IC) {
        const legsSorted = setToSortedArray(
          lifecycleBundleLegSetByKey[lifeBundleKey] || {},
        );
        const isOpen = String(posEffect || "")
          .toUpperCase()
          .includes("OPEN");
        const isCloseLike = isCloseLikeBundle(posEffect, spreadOut);

        if (legsSorted.length && (isOpen || isCloseLike)) {
          for (let j = 0; j < legsSorted.length; j++) {
            const legId = legsSorted[j];
            const k = makeIcLegQtyKey(Account, symForMatch, expKey, legId);
            const prev = Number(openIcQtyByLegKey[k] || 0);

            if (isOpen) openIcQtyByLegKey[k] = prev + Number(qtyAbs || 0);
            else if (isCloseLike)
              openIcQtyByLegKey[k] = Math.max(0, prev - Number(qtyAbs || 0));
          }
        }
      }
      let groupKey = "";

      if (spread === "STOCK" || spread === "SINGLE" || isExerciseOrAssign) {
        // Stocks/singles include price so multiple same-minute fills don't collapse.
        groupKey = [
          Account,
          dateIso,
          timeHHmmss,
          symForMatch,
          spread,
          qtyAbs,
          roundTo(matchPriceForPull, 4),
          orderType,
        ].join("|");
      } else if (spread === "DIAGONAL" || spread === "CALENDAR") {
        // Must match the pre-pass diagonal/calendar shape:
        // coarseKey = [Account, dateIso, timeHHmmss, sym, spread, qtyAbs, posEffect, typeKey]
        // groupKey  = coarseKey + "|" + expSig
        const coarseKey = [
          Account,
          dateIso,
          timeHHmmss,
          symForMatch,
          spread,
          qtyAbs,
          posEffect,
          typeKey,
        ].join("|");
        const expSig = diagCalExpSigByCoarseKey[coarseKey] || expKey; // fallback if something weird happens
        groupKey = coarseKey + "|" + expSig;
      } else {
        // Generic multi-leg spreads:
        // - include Exp to prevent collisions across expirations
        // - use MIXED for multi-type spreads so IC CALL+PUT legs share one groupKey
        const spreadTypeKey = isMultiTypeSpread(spread) ? "MIXED" : typeKey;
        groupKey = [
          Account,
          dateIso,
          timeHHmmss,
          symForMatch,
          spread,
          qtyAbs,
          posEffect,
          expKey,
          spreadTypeKey,
        ].join("|");
      }
      //  For CUSTOM spreads, legs can have blank Exp/Strike and Type=MANUAL, which can cause
      // multiple distinct groupKey values even though TosTop has ONE TRD row for the strategy.
      // So use a coarser key ONLY for enrichment-consumption gating.
      const enrichmentGroupKey =
        spread === "CUSTOM"
          ? [
              Account,
              dateIso,
              timeHHmmss,
              symForMatch,
              spread,
              qtyAbs,
              posEffect,
            ].join("|")
          : groupKey;

      // Computed spread net (from the pre-pass). Only exists for multi-leg spreads.
      let computedSpreadNetAbs = null;
      if (spread !== "STOCK" && spread !== "SINGLE" && !isExerciseOrAssign) {
        computedSpreadNetAbs = spreadNetAbsByGroupKey[groupKey];
      }
      const expectedFillCount = Number(
        spreadFillCountByGroupKey[groupKey] || 1,
      );

      // If this is a multi-leg spread and TosTrades Net Price is non-numeric,
      // and we ALSO failed to compute a strategy net from legs, enrichment matching will be weaker.
      let hasComputedSpreadNet =
        typeof computedSpreadNetAbs === "number" &&
        isFinite(computedSpreadNetAbs) &&
        computedSpreadNetAbs > 0;

      // NEW: also check butterfly net (pre-pass already computed it)
      if (!hasComputedSpreadNet && spread === "BUTTERFLY") {
        const expKey =
          exp instanceof Date
            ? Utilities.formatDate(
                exp,
                Session.getScriptTimeZone(),
                "yyyy-MM-dd",
              )
            : toStr(exp).trim();
        const bfKey = [
          Account,
          dateIso,
          timeHHmmss,
          sym,
          spread,
          posEffect,
          expKey,
          typeKey,
        ].join("|");
        const bfNet = butterflyNetAbsByKey[bfKey];
        if (typeof bfNet === "number" && isFinite(bfNet) && bfNet > 0) {
          hasComputedSpreadNet = true;
          computedSpreadNetAbs = bfNet; // reuse for matchPrice below
        }
      }

      if (isMultiLegSpread && !hasNumericNet && !hasComputedSpreadNet) {
        // De-dupe per strategy group (not per leg row)
        const warnKey = groupKey;

        if (!weakEnrichmentMatchSeen[warnKey]) {
          weakEnrichmentMatchSeen[warnKey] = true;

          importIssuesAdd(
            ctx,
            "WARN_WEAKENRICHMATCHPRICE",
            "TosTrades",
            i + 2,
            "Net Price / ComputedNet",
            String(netPriceRaw ?? ""),
            JSON.stringify({
              message:
                "Net Price is non-numeric and computed spread net is missing; enrichment matching will rely on leg Price (weaker).",
              Account: Account,
              dateIso: dateIso,
              timeHHmmss: timeHHmmss,
              symForMatch: symForMatch,
              spread: spread,
              qtyAbs: qtyAbs,
              posEffect: posEffect,
              expKey: expKey,
              typeKey: typeKey,
              groupKey: groupKey,
            }),
          );
        }
      }

      // If we have a computed spread net, use it as the enrichment match price.
      if (hasComputedSpreadNet) {
        matchPriceForPull = computedSpreadNetAbs;
      }

      let miscFees = "";
      let feesComm = "";
      let amount = "";

      // If this is an EXERCISE/ASSIGN option event, we'll create a synthetic STOCK row.
      // Any TosTop enrichment Amount belongs on that STOCK row (not on the option row).
      let exerciseStockEnrichment = null; // { miscFees, feesComm, amount }
      // Build the trade fields for the unified Schwab Import row.
      const action = actionFromTosTrades(side, posEffect, typeKey);
      const unifiedSymbol = formatUnifiedSymbol(sym, typeKey, exp, strike);

      const desc =
        "TOS Trades " +
        side +
        " " +
        toStr(cell(row, tradesTbl.idx, "Qty")).trim() +
        " " +
        symRawNorm +
        " " +
        typeKey +
        " " +
        toStr(cell(row, tradesTbl.idx, "Price")).trim() +
        " " +
        posEffect;

      // ====================================================================
      // FAKE TRADE FILTER
      // Drops paper/simulated FAKE trades but keeps real trades that TOS
      // incorrectly tagged FAKE (identified via CusipMap resolution).
      //
      // Cases handled:
      // Case 1: FAKE + dotted option symbol (.QQQ...) → always DROP (paper)
      // Case 2: FAKE + CUSIP not in CusipMap → DROP but WARN loudly in Import Issues
      //         so you can investigate and add to CusipMap if it was real.
      // Case 3: FAKE + CUSIP in CusipMap + ticker NOT in FakeDropOverride → KEEP (real trade)
      // Case 4: FAKE + CUSIP in CusipMap BUT ticker IS in FakeDropOverride → DROP (known bogus)
      // ====================================================================
      const descForFakeCheck = toStr(desc).toUpperCase();
      const hasFakeTag =
        descForFakeCheck.includes(" FAKE ") ||
        descForFakeCheck.includes(" FAKE\t") ||
        descForFakeCheck.endsWith(" FAKE");

      if (hasFakeTag) {
        // `sym` is already post-CusipMap resolved at this point in the loop.
        // `symRawNorm` is the normalized raw symbol before CusipMap lookup.
        // We check whether the original raw symbol was CUSIP-like, and whether
        // CusipMap resolved it to a real ticker.
        const rawNormForFake = normalizeSymbol(symRawCell);
        const cusipKeyForFake = normalizeCusip(rawNormForFake);
        const rawIsCusipLike =
          looksLikeCusip(cusipKeyForFake) && /[0-9]/.test(cusipKeyForFake);
        const rawIsDottedOption = toStr(symRawCell).trim().startsWith(".");
        const cusipResolvedToTicker =
          rawIsCusipLike && !!cusipMap[cusipKeyForFake];

        // Case 4 check: is the resolved ticker in the FakeDropOverride list?
        // `fakeDropOverrideSet` is loaded once before the loop (see helper below).
        const resolvedTicker = cusipResolvedToTicker
          ? cusipMap[cusipKeyForFake]
          : sym;
        const isOverriddenDrop =
          fakeDropOverrideSet[resolvedTicker.toUpperCase()] === true ||
          fakeDropOverrideSet[cusipKeyForFake] === true;

        if (rawIsDottedOption) {
          // Case 1: Dotted option symbol → confirmed paper trade, drop silently.
          importIssuesAdd(
            ctx,
            "INFO",
            "TosTrades",
            i + 2,
            "FAKE-DROPPED",
            desc.substring(0, 100),
            JSON.stringify({
              reason: "dottedOption",
              sym,
              symRawCell,
              Account,
              dateIso,
            }),
          );
          continue;
        } else if (isOverriddenDrop) {
          // Case 4: CUSIP is in CusipMap BUT ticker is on the FakeDropOverride list.
          // This handles the NVDS scenario: a paper trade whose ticker exists in CusipMap.
          importIssuesAdd(
            ctx,
            "INFO",
            "TosTrades",
            i + 2,
            "FAKE-DROPPED-OVERRIDE",
            desc.substring(0, 100),
            JSON.stringify({
              reason: "FakeDropOverride",
              sym,
              resolvedTicker,
              Account,
              dateIso,
            }),
          );
          continue;
        } else if (rawIsCusipLike && !cusipResolvedToTicker) {
          // Case 2: CUSIP-like but NOT in CusipMap.
          // First check FakeDropOverride by raw CUSIP directly (handles deleted-from-CusipMap overrides).
          if (fakeDropOverrideSet[cusipKeyForFake] === true) {
            importIssuesAdd(
              ctx,
              "INFO",
              "TosTrades",
              i + 2,
              "FAKE-DROPPED-OVERRIDE",
              desc.substring(0, 100),
              JSON.stringify({
                reason: "FakeDropOverride-byCusip",
                cusip: cusipKeyForFake,
                sym,
                Account,
                dateIso,
              }),
            );
            continue;
          }
          // Genuinely unknown CUSIP → warn loudly so you can investigate.
          importIssuesAdd(
            ctx,
            "WARN",
            "TosTrades",
            i + 2,
            "FAKE-DROPPED-CUSIP-MISSING",
            desc.substring(0, 100),
            JSON.stringify({
              message:
                "FAKE-tagged row dropped because its CUSIP was not found in CusipMap. " +
                "Verify in Schwab Transaction History: if this was a real trade, add the " +
                "CUSIP to CusipMap and re-run. If it was paper, add the CUSIP to FakeDropOverride.",
              cusip: cusipKeyForFake,
              sym,
              symRawCell,
              Account,
              dateIso,
            }),
          );
          continue;
        } else if (cusipResolvedToTicker) {
          // Case 3: CUSIP resolved to a known ticker and NOT overridden → keep the row.
          // Log at INFO level for traceability (PCSA, DBGI, BGLC, LURAF etc.).
          importIssuesAdd(
            ctx,
            "INFO",
            "TosTrades",
            i + 2,
            "FAKE-KEPT",
            desc.substring(0, 100),
            JSON.stringify({
              reason: "cusipResolvedInMap",
              sym,
              cusip: cusipKeyForFake,
              Account,
              dateIso,
            }),
          );
          // Fall through — row is processed normally below.
        } else {
          // Catch-all: FAKE tag, symbol is not a dotted option and not CUSIP-like.
          // This means the raw TosTrades Symbol column contains a plain ticker (e.g. "AMC", "AGLE").
          // Check FakeDropOverride for the plain ticker first — allows manual drop of these without code changes.
          const plainTickerUpper = toStr(sym).trim().toUpperCase();
          if (fakeDropOverrideSet[plainTickerUpper] === true) {
            importIssuesAdd(
              ctx,
              "INFO",
              "TosTrades",
              i + 2,
              "FAKE-DROPPED-OVERRIDE",
              desc.substring(0, 100),
              JSON.stringify({
                reason: "FakeDropOverride-byPlainTicker",
                sym: plainTickerUpper,
                Account,
                dateIso,
              }),
            );
            continue;
          }
          // Not overridden — warn loudly so you can decide: add to FakeDropOverride (paper) or investigate.
          importIssuesAdd(
            ctx,
            "WARN",
            "TosTrades",
            i + 2,
            "FAKE-DROPPED-UNKNOWN",
            desc.substring(0, 100),
            JSON.stringify({
              message:
                "FAKE-tagged row dropped: symbol is a plain ticker (not CUSIP, not dotted option). " +
                "If this is a real trade, verify via Schwab history. " +
                "If it is paper, add the ticker to FakeDropOverride and re-run.",
              sym,
              symRawCell,
              Account,
              dateIso,
            }),
          );
          continue;
        }
      }
      // ====================================================================
      // END FAKE TRADE FILTER
      // ====================================================================

      // Enrichment gating
      // NOTE: isStockOrSingle + isMultiLegSpread are defined earlier in the loop (TDZ-safe).
      const alreadyConsumed =
        !!enrichmentConsumedByGroupKey[enrichmentGroupKey];

      // For EXERCISE/ASSIGN we DO want enrichment (TosTop has the cash movement),
      // but we will move the pulled Amount to the synthetic STOCK leg.
      let shouldAttemptEnrichment = false;
      if (isStockOrSingle) {
        shouldAttemptEnrichment = true;
      } else if (isMultiLegSpread && !alreadyConsumed) {
        shouldAttemptEnrichment = side === "SELL"; // deterministic: attempt on SELL leg
      }

      //  For BUTTERFLY, TosTrades middle leg qty is 2x the strategy qty (1:-2:1).
      // TosTop TRD rows typically reflect the strategy qty (e.g., "BOT +1 BUTTERFLY ...").
      let qtyAbsForEnrichment = qtyAbs;
      if (
        spread === "BUTTERFLY" &&
        isFinite(qtyAbsForEnrichment) &&
        qtyAbsForEnrichment > 0 &&
        qtyAbsForEnrichment % 2 === 0
      ) {
        qtyAbsForEnrichment = qtyAbsForEnrichment / 2;
      }

      if (shouldAttemptEnrichment) {
        // Primary attempt
        let pulledResult;
        if (spread === "BUTTERFLY") {
          pulledResult = pullTopTradeEnrichment(
            Account,
            ts,
            symForMatch,
            qtyAbsForEnrichment,
            matchPriceForPull,
            1,
          );
          if (!pulledResult || !pulledResult.item) {
            pulledResult = pullTopTradeEnrichmentButterfly(
              Account,
              dateIso,
              timeHHmm,
              symForMatch,
            );
          }
        } else if (spread === "IRON CONDOR") {
          const icCacheKey = [
            Account,
            dateIso,
            timeHHmm,
            String(symForMatch || "")
              .trim()
              .toUpperCase(),
            String(qtyAbsForEnrichment),
          ].join("|");

          if (icPackageEnrichCache[icCacheKey]) {
            pulledResult = icPackageEnrichCache[icCacheKey];
          } else {
            pulledResult = pullTopTradeEnrichmentIronCondor(
              Account,
              dateIso,
              timeHHmm,
              symForMatch,
              ts,
              qtyAbsForEnrichment,
            );
            if (!pulledResult || !pulledResult.item) {
              pulledResult = pullTopTradeEnrichment(
                Account,
                ts,
                symForMatch,
                qtyAbsForEnrichment,
                matchPriceForPull,
                expectedFillCount,
              );
            }
            if (pulledResult && pulledResult.item) {
              icPackageEnrichCache[icCacheKey] = pulledResult;
            }
          }
        } else {
          pulledResult = pullTopTradeEnrichment(
            Account,
            ts,
            symForMatch,
            qtyAbsForEnrichment,
            matchPriceForPull,
            expectedFillCount,
          );
        }

        let pulled = pulledResult ? pulledResult.item : null;
        let pullDebug = pulledResult ? pulledResult.debug : null;

        // Spread fallback:
        // If computed-net match fails, try leg price as a secondary attempt.
        if (!pulled) {
          const canTryLegFallback =
            !(
              spread === "STOCK" ||
              spread === "SINGLE" ||
              isExerciseOrAssign
            ) &&
            typeof computedSpreadNetAbs === "number" &&
            isFinite(computedSpreadNetAbs) &&
            computedSpreadNetAbs > 0 &&
            !pricesClose(computedSpreadNetAbs, price);

          if (canTryLegFallback) {
            const pulledResult2 = pullTopTradeEnrichment(
              Account,
              ts,
              symForMatch,
              qtyAbsForEnrichment,
              price, // <-- use leg price on fallback
              expectedFillCount,
            );

            if (pulledResult2 && pulledResult2.item) {
              pulledResult = pulledResult2;
              pulled = pulledResult2.item;
              pullDebug = pulledResult2.debug;
              matchPriceForPull = price;
            }
          }
        }

        if (pulled) {
          const mf = toNum(pulled.miscFees);
          const fc = toNum(pulled.feesComm);
          const am = toNum(pulled.amount);

          miscFees = isNaN(mf) ? (pulled.miscFees ?? "") : mf;
          feesComm = isNaN(fc) ? (pulled.feesComm ?? "") : fc;
          amount = isNaN(am) ? (pulled.amount ?? "") : am;

          // If this is an option EXERCISE/ASSIGN event, move the cash movement to the synthetic STOCK leg.
          if (isExerciseOrAssign && (typeKey === "CALL" || typeKey === "PUT")) {
            exerciseStockEnrichment = {
              miscFees: miscFees,
              feesComm: feesComm,
              amount: amount,
            };
            miscFees = "";
            feesComm = "";
            amount = "";
          }

          enrichmentConsumedByGroupKey[enrichmentGroupKey] = true;
        } else {
          // pullTopTradeEnrichment() debug counters:
          // - counts.minuteBucket = how many TosTop TRD candidates exist in that minute bucket
          // - counts.symQtyCandidates = how many match (symbol + qty) inside that minute
          // For butterfly debug, we still use debug.dtBucketCount.
          const dtCount =
            pullDebug && pullDebug.counts
              ? Number(pullDebug.counts.minuteBucket || 0)
              : Number(
                  pullDebug && pullDebug.dtBucketCount
                    ? pullDebug.dtBucketCount
                    : 0,
                );

          const symCount =
            pullDebug && pullDebug.counts
              ? Number(pullDebug.counts.symQtyCandidates || 0)
              : 0;

          // Classify failure (data gap vs mismatch) like your original code
          if (dtCount === 0) {
            // Same rule as NO_SYMBOL_IN_MINUTE: exercise/assign cash lives on
            // the synthetic STOCK leg. Do not treat the companion stock line
            // as a missing TosTop TRD.
            if (!isExerciseOrAssignSpread(spread)) {
              const gapKey = [
                Account, // <-- add this for Mode B
                dateIso,
                timeHHmm,
                symForMatch, // <-- use the same symbol you matched with (optional but recommended)
                spread,
                toStr(exp),
                typeKey,
                toStr(strike),
                posEffect,
                side,
                orderType,
              ].join("|");

              if (!missingTradeEnrichmentDataGapSeen[gapKey]) {
                missingTradeEnrichmentDataGapSeen[gapKey] = true;
                importIssuesAdd(
                  ctx,
                  "TRADE_ENRICHMENT_DATA_GAP_NO_TRD_ROWS",
                  "TosTrades",
                  i + 2,
                  "TosTop TRD minute",
                  gapKey,
                  JSON.stringify({
                    message: "No TosTop TRD rows exist in this minute",
                    pullDebug,
                  }),
                );
              }
            }
          } else if (symCount === 0) {
            // NEW: skip gap warning for EXERCISE/ASSIGN — enrichment is moved to synthetic STOCK leg
            if (isExerciseOrAssignSpread(spread)) {
              // silent on purpose — TosTop uses different DESCRIPTION for cash movement
            } else {
              const gapKey = [
                Account,
                dateIso,
                timeHHmm,
                symForMatch,
                spread,
                toStr(exp),
                typeKey,
                toStr(strike),
                posEffect,
                side,
                orderType,
              ].join("|");

              if (!missingTradeEnrichmentDataGapSeen[gapKey]) {
                missingTradeEnrichmentDataGapSeen[gapKey] = true;
                importIssuesAdd(
                  ctx,
                  "TRADE_ENRICHMENT_DATA_GAP_NO_SYMBOL_IN_MINUTE",
                  "TosTrades",
                  i + 2,
                  "TosTop TRD minute",
                  gapKey,
                  JSON.stringify({
                    message:
                      "TosTop has TRD rows in this minute, but none match symbol",
                    pullDebug,
                  }),
                );
              }
            }
          } else {
            if (!missingTradeEnrichmentSeen[groupKey]) {
              missingTradeEnrichmentSeen[groupKey] = true;
              importIssuesAdd(
                ctx,
                "TRADE_ENRICHMENT_NOT_FOUND",
                "TosTrades",
                i + 2,
                "TosTop TRD match",
                groupKey,
                JSON.stringify({
                  message:
                    "No matching TosTop TRD row found for fees/amount enrichment",
                  pullDebug,
                }),
              );
            }
          }
        }
      }

      const rowObj = {
        Account: Account,
        Date: dateIso, // display-friendly date
        Time: timeHHmmss, // HHmmss text, matches Timestamp but is not the sort key
        Timestamp: ts, // authoritative Date object for ordering/grouping

        Action: action,
        Symbol: unifiedSymbol,
        Description: desc,

        Spread: spreadOut,
        Quantity: isNaN(qtyAbs) ? "" : qtyAbs,
        Price: isNaN(price) ? "" : price,

        NetPrice: toStr(netPriceRaw).trim(),

        Side: side,
        PosEffect: posEffect,
        Exp: exp,
        Strike: strike,
        OrderType: orderType,

        MiscFees: miscFees,
        FeesComm: feesComm,
        Amount: amount,

        // internal grouping helpers
        tradeGroupKey: groupKey,
        optType: typeKey,
      };

      unifiedTrades.push(rowObj);

      // ---------------------------- Synthetic STOCK leg for EXERCISE / ASSIGN ----------------------------
      // For option exercise/assignment events, emit an additional STOCK row so the share movement is explicit.
      // Fees/amount enrichment is pulled from TosTop and applied to the STOCK leg (not the option leg).

      if (
        isExerciseOrAssign &&
        (typeKey === "CALL" || typeKey === "PUT") &&
        isFinite(qtyAbs) &&
        qtyAbs > 0
      ) {
        const strikeNum = toNum(strike);
        const stockPrice = isFinite(strikeNum)
          ? strikeNum
          : hasNumericNet
            ? netPriceNum
            : NaN;

        // Infer stock direction from whether the option close was SELL (closing long) or BUY (closing short).
        // - If closing long CALL => BUY stock
        // - If closing long PUT  => SELL stock
        // - If closing short CALL (assigned) => SELL stock
        // - If closing short PUT  (assigned) => BUY stock
        let stockSide = "";
        if (side === "SELL") {
          stockSide = typeKey === "CALL" ? "BUY" : "SELL";
        } else if (side === "BUY") {
          stockSide = typeKey === "CALL" ? "SELL" : "BUY";
        }

        const stockQtyShares = qtyAbs * OPTIONS_CONTRACT_MULTIPLIER;

        let stockAmount = "";
        let stockMiscFees = "";
        let stockFeesComm = "";

        // Prefer TosTop enrichment for exercises/assignments (matches TosTop cash movement rows).
        if (exerciseStockEnrichment) {
          stockMiscFees = exerciseStockEnrichment.miscFees;
          stockFeesComm = exerciseStockEnrichment.feesComm;
          stockAmount = exerciseStockEnrichment.amount;
        }

        // If TosTop enrichment didn't provide Amount, compute from strike.
        if (
          toStr(stockAmount).trim() === "" &&
          isFinite(stockPrice) &&
          stockSide
        ) {
          stockAmount = computeSignedAmountFromTrade(
            stockSide,
            stockQtyShares,
            stockPrice,
          );
        }

        const stockPosEffect =
          stockSide === "BUY"
            ? "TO OPEN"
            : stockSide === "SELL"
              ? "TO CLOSE"
              : "";
        const stockAction =
          stockSide === "BUY" ? "Buy" : stockSide === "SELL" ? "Sell" : "";

        const stockRowObj = {
          Account: Account,
          Date: dateIso,
          Time: timeHHmmss,
          Timestamp: ts,

          Action: stockAction,
          Symbol: symForMatch,
          Description:
            "Synthetic STOCK leg from " +
            spread +
            " " +
            typeKey +
            " (" +
            side +
            " " +
            posEffect +
            ")",

          Spread: "STOCK",
          Quantity: stockQtyShares,
          Price: isFinite(stockPrice) ? stockPrice : "",
          NetPrice: "",

          Side: stockSide,
          PosEffect: stockPosEffect,
          Exp: "",
          Strike: "",
          OrderType: orderType,

          MiscFees: stockMiscFees,
          FeesComm: stockFeesComm,
          Amount: stockAmount,

          // Keep grouped with the option row for sorting adjacency
          tradeGroupKey: groupKey,
          optType: "STOCK",
        };

        unifiedTrades.push(stockRowObj);
      }
    }

    // Group I output helpers live in BuildUnifiedOutput.js.
    const _out = createOutputHelpers({
      toStr: toStr,
    });
    const applyFeesToTopLegRuleB = _out.applyFeesToTopLegRuleB;
    const legOrderKeyForOutput = _out.legOrderKeyForOutput;

    applyFeesToTopLegRuleB(unifiedTrades);

        pipelineTimingLog(
      "buildUnifiedImportV3 main",
      tMain,
      "unifiedTrades=" + unifiedTrades.length,
    );
    const tNontrade = pipelineTimingNow();

    // =========================================================================
    // 7) CONVERT TosTop NON-TRADE ROWS → UNIFIED NON-TRADE ROWS
    //    Handles everything in TosTop that is NOT a regular trade (TYPE != TRD):
    //    cash movements, fees, dividends, transfers, corporate actions, etc.
    //    These rows are turned into the corresponding rows in the final
    //    "Schwab Import" table.
    // =========================================================================
    const unifiedNonTrades = [];
    const radSplitCandidateBuffer = {};

    for (let i = 0; i < topRows.length; i++) {
      const r = topRows[i];

      //  carry account through to Schwab Import (canonical).
      const Account = toStr(getField(r, "Account")).trim().toUpperCase();

      const type = toStr(getField(r, "TYPE")).trim().toUpperCase();

      if (type === "TRD") {
        // Both the DRIP BTO emitter and the TDA Fractional Sell emitter handle TRD
        // rows that have NO TosTrades partner. They share this gate so only one
        // emitter fires per row, then continue to the next TosTop row.
        // The isDripTrdDescription() and isTdaFractionalSellTrd() helpers are
        // defined earlier in this function. Change 2 already skips both row types
        // from the enrichment queue, so there is no double-processing risk.

        const emitDescRaw = toStr(getField(r, "DESCRIPTION")).trim();
        const emitParsed = parseTosTopTradeDescription(emitDescRaw);
        const emitQty = emitParsed ? emitParsed.absQty : null;
        const emitAmountRaw = getField(r, "AMOUNT");
        const emitAmount = toNum(emitAmountRaw);

        // -------------------------------------------------------
        // DRIP BTO emitter: Schwab-era "UPON" rows and TDA-era
        // fractional BOT rows (e.g. "BOT +0.13 XOM @101.49692").
        // Action = Buy to Open / STOCK.
        // -------------------------------------------------------
        if (isDripTrdDescription(emitDescRaw, emitQty)) {
          if (!emitParsed || !emitParsed.symbol || emitQty == null) {
            importIssuesAdd(
              ctx,
              "WARN",
              "TosTop",
              i + 2,
              "DESCRIPTION",
              emitDescRaw,
              "DRIP TRD row: could not parse symbol/qty from description. Row skipped.",
            );
          } else {
            const dripAccount = toStr(getField(r, "Account"))
              .trim()
              .toUpperCase();
            const dripDateIso = normalizeDate(getField(r, "DATE"));
            let dripTimeHHmmss = normalizeTimeHHmmss(getField(r, "TIME"));
            if (!dripTimeHHmmss) dripTimeHHmmss = "000000";
            const dripTs = toDateObject(dripDateIso, dripTimeHHmmss);

            // Back-calculate price from |Amount| ÷ Qty.
            let dripPrice = null;
            if (!isNaN(emitAmount) && !isNaN(emitQty) && emitQty !== 0) {
              dripPrice = roundTo(Math.abs(emitAmount) / emitQty, 4);
            }

            let dripSym = normalizeUnderlyingFromTradeSymbol(emitParsed.symbol);
            if (
              looksLikeCusip(dripSym) &&
              /[0-9A-Z]{9}/.test(normalizeCusip(dripSym))
            ) {
              const mapped = cusipMap[dripSym];
              if (mapped) dripSym = normalizeSymbol(mapped);
            }

            const dripRowObj = {
              Account: dripAccount,
              Date: dripDateIso,
              Time: dripTimeHHmmss,
              Timestamp:
                dripTs instanceof Date && !isNaN(dripTs.getTime())
                  ? dripTs
                  : null,
              Action: "Buy",
              Symbol: dripSym,
              Description:
                "DRIP BUY +" + emitQty + " " + dripSym + " UPON REINVESTMENT",
              Spread: "STOCK",
              Quantity: emitQty,
              Price: dripPrice !== null ? dripPrice : "",
              NetPrice: dripPrice !== null ? String(dripPrice) : "",
              Side: "BUY",
              PosEffect: "TO OPEN",
              Exp: "",
              Strike: "",
              OrderType: "",
              MiscFees: "",
              FeesComm: "",
              Amount: !isNaN(emitAmount) ? emitAmount : "",
              tradeGroupKey: "",
              optType: "",
            };
            unifiedTrades.push(dripRowObj);
            importIssuesAdd(
              ctx,
              "INFO",
              "TosTop",
              i + 2,
              "DRIP",
              dripSym,
              "DRIP TRD row emitted as Buy row. Qty: " +
                emitQty +
                " Price: " +
                dripPrice +
                " Amount: " +
                emitAmount,
            );
          }

          // -------------------------------------------------------
          // TDA Fractional Sell emitter: TDA-era fractional SELL
          // rows (e.g. "SOLD -0.13 XOM @98.67") with no TosTrades
          // partner. TDA never logged these in the Trades blotter.
          // Action = Sell to Close / STOCK.
          // -------------------------------------------------------
        } else if (
          isTdaFractionalSellTrd(
            emitDescRaw,
            emitQty,
            emitAmount,
            normalizeDate(getField(r, "DATE")),
            toStr(getField(r, "Account")).trim().toUpperCase(),
          )
        ) {
          if (!emitParsed || !emitParsed.symbol || emitQty == null) {
            importIssuesAdd(
              ctx,
              "WARN",
              "TosTop",
              i + 2,
              "DESCRIPTION",
              emitDescRaw,
              "TDA fractional sell TRD: could not parse symbol/qty. Row skipped.",
            );
          } else {
            const sellAccount = toStr(getField(r, "Account"))
              .trim()
              .toUpperCase();
            const sellDateIso = normalizeDate(getField(r, "DATE"));
            let sellTimeHHmmss = normalizeTimeHHmmss(getField(r, "TIME"));
            if (!sellTimeHHmmss) sellTimeHHmmss = "000000";
            const sellTs = toDateObject(sellDateIso, sellTimeHHmmss);

            // Prefer parsed price from description (@98.67); fall back to |Amount| ÷ Qty.
            let sellPrice =
              emitParsed.price !== null && !isNaN(emitParsed.price)
                ? emitParsed.price
                : !isNaN(emitAmount) && emitQty
                  ? roundTo(Math.abs(emitAmount) / emitQty, 4)
                  : null;

            let sellSym = normalizeUnderlyingFromTradeSymbol(emitParsed.symbol);
            if (
              looksLikeCusip(sellSym) &&
              /[0-9A-Z]{9}/.test(normalizeCusip(sellSym))
            ) {
              const mapped = cusipMap[sellSym];
              if (mapped) sellSym = normalizeSymbol(mapped);
            }

            const sellRowObj = {
              Account: sellAccount,
              Date: sellDateIso,
              Time: sellTimeHHmmss,
              Timestamp:
                sellTs instanceof Date && !isNaN(sellTs.getTime())
                  ? sellTs
                  : null,
              Action: "Sell",
              Symbol: sellSym,
              Description:
                "SELL -" + emitQty + " " + sellSym + " @" + (sellPrice ?? ""),
              Spread: "STOCK",
              Quantity: emitQty,
              Price: sellPrice !== null ? sellPrice : "",
              NetPrice: sellPrice !== null ? String(sellPrice) : "",
              Side: "SELL",
              PosEffect: "TO CLOSE",
              Exp: "",
              Strike: "",
              OrderType: "",
              MiscFees: toNum(getField(r, "Misc Fees", "MiscFees")) || "",
              FeesComm:
                toNum(
                  getField(
                    r,
                    "Commissions & Fees",
                    "Commissions Fees",
                    "Commissions and Fees",
                  ),
                ) || "",
              Amount: !isNaN(emitAmount) ? emitAmount : "",
              tradeGroupKey: "",
              optType: "",
            };
            unifiedTrades.push(sellRowObj);
            importIssuesAdd(
              ctx,
              "INFO",
              "TosTop",
              i + 2,
              "TDA-DRIP-SELL",
              sellSym,
              "TDA fractional sell TRD emitted as Sell to Close row. Qty: " +
                emitQty +
                " Price: " +
                sellPrice +
                " Amount: " +
                emitAmount,
            );
          }
        }
        // Whether the DRIP emitter or the TDA fractional sell emitter fired (or
        // neither — meaning this TRD has a TosTrades partner already handled in
        // the enrichment queue), always skip past the non-trade processing below.
        continue;
      } // end if (type === 'TRD')

      const dateIso = normalizeDate(getField(r, "DATE"));

      // TosTop TIME is HHmmss text now (ex: "230208")
      let timeHHmmss = normalizeTimeHHmmss(getField(r, "TIME"));
      if (!timeHHmmss) timeHHmmss = "000000";

      // Keep HHmm available if you ever need minute bucketing in non-trades logic
      const timeHHmm = normalizeTime(timeHHmmss);

      const ts = toDateObject(dateIso, timeHHmmss);
      if (
        !(ts instanceof Date) ||
        isNaN(ts.getTime()) ||
        ts.getFullYear() < 2000
      ) {
        importIssuesAdd(
          ctx,
          "BADTOSTOPDATETIME",
          "TosTop",
          i + 2,
          "DATETIME",
          JSON.stringify({
            DATE: getField(r, "DATE"),
            TIME: getField(r, "TIME"),
          }),
          "Cannot build Timestamp from TosTop DATETIME",
        );
        continue;
      }

      const desc = toStr(getField(r, "DESCRIPTION")).trim();

      let qtyOut = "";
      let expOut = "";
      let strikeOut = "";

      const miscFees = getField(r, ["Misc Fees", "MiscFees"]);
      const feesComm = getField(r, [
        "Commissions & Fees",
        "Commissions Fees",
        "Commissions and Fees",
      ]);
      const amount = getField(r, "AMOUNT");

      let symbolOut = "";

      // ── Map-driven corporate-action STOCK emitter ──────────────────────────
      // For rows like:
      //   MANDATORY - EXCHANGE 375.0 ISENF
      //   NON-TAXABLE SPIN OFF/LIQUIDATION DISTRIBUTION 96.0 50545P309
      // we emit a synthetic STOCK row into unifiedTrades so downstream block
      // logic sees the delivered shares at the authoritative timestamp.
      const corpActionStockParsed = parseCorpActionStockDescription_(
        desc,
        cusipMap,
      );
      if (corpActionStockParsed) {
        const corpActionStockMapRow = findCorpActionStockMapRow_(
          corpActionStockMap,
          Account,
          dateIso,
          corpActionStockParsed.phrase,
          corpActionStockParsed.resolvedSymbol,
        );

        if (corpActionStockMapRow) {
          const qtyMultiplier = Number(
            corpActionStockMapRow.qtyMultiplier || 1,
          );
          const emitQty =
            Math.round(corpActionStockParsed.parsedQty * qtyMultiplier * 1e8) /
            1e8;
          const emitSymbol = normalizeSymbol(
            corpActionStockMapRow.emitSymbol ||
              corpActionStockParsed.resolvedSymbol,
          );
          const emitActionU = String(corpActionStockMapRow.emitAction || "BUY")
            .trim()
            .toUpperCase();

          // IGNORE / SKIP = matched on purpose, do not invent a position
          // (1:1 same-ticker recap such as Stock Merger 62.0 SA).
          if (emitActionU === "IGNORE" || emitActionU === "SKIP") {
            importIssuesAdd(
              ctx,
              "INFO",
              "TosTop",
              i + 2,
              "DESCRIPTION",
              desc,
              "CorpActionStockMap Emit Action=" +
                emitActionU +
                "; no synthetic STOCK row. " +
                "Phrase=" +
                corpActionStockParsed.phrase +
                "; MatchSymbol=" +
                corpActionStockParsed.resolvedSymbol +
                "; Qty=" +
                corpActionStockParsed.parsedQty +
                ".",
            );
          } else {
            const stockAction = emitActionU === "SELL" ? "Sell" : "Buy";
            const stockSide = emitActionU === "SELL" ? "SELL" : "BUY";
            const stockPosEffect =
              emitActionU === "SELL" ? "TO CLOSE" : "TO OPEN";

            const syntheticDesc =
              `${corpActionStockParsed.phrase} ${corpActionStockMapRow.sourceSymbol} -> ${emitSymbol} ` +
              `QTY=${emitQty} RAWTOKEN=${corpActionStockParsed.rawToken}`;

            const corpActionStockRowObj = {
              Account: Account,
              Date: dateIso,
              Time: timeHHmmss,
              Timestamp: ts,
              Action: stockAction,
              Symbol: emitSymbol,
              Description: syntheticDesc,
              Spread: "STOCK",
              Quantity: emitQty,
              Price: 0,
              NetPrice: 0,
              Side: stockSide,
              PosEffect: stockPosEffect,
              Exp: "",
              Strike: "",
              OrderType: "",
              MiscFees: "",
              FeesComm: "",
              Amount: "",
              tradeGroupKey: "",
              optType: "",
            };

            unifiedTrades.push(corpActionStockRowObj);
            symbolOut = emitSymbol;

            importIssuesAdd(
              ctx,
              "INFO",
              "TosTop",
              i + 2,
              "DESCRIPTION",
              desc,
              `Synthetic corp-action STOCK row emitted. Phrase=${corpActionStockParsed.phrase}; ` +
                `Source=${corpActionStockMapRow.sourceSymbol}; Emit=${emitSymbol}; Qty=${emitQty}; ` +
                `MatchSymbol=${corpActionStockParsed.resolvedSymbol}.`,
            );
          }
        } else {
          if (!symbolOut && corpActionStockParsed.resolvedSymbol) {
            symbolOut = corpActionStockParsed.resolvedSymbol;
          }

          importIssuesAdd(
            ctx,
            "WARN",
            "TosTop",
            i + 2,
            "DESCRIPTION",
            desc,
            `Corp-action share delivery row detected but no CorpActionStockMap match was found. ` +
              `Account=${Account}; Date=${dateIso}; Phrase=${corpActionStockParsed.phrase}; ` +
              `MatchSymbol=${corpActionStockParsed.resolvedSymbol}; Qty=${corpActionStockParsed.parsedQty}.`,
          );
        }
      }

      // ── RAD: Mandatory Split (forward or reverse) ──────────────────────────
      // Detect MANDATORY [REVERSE] SPLIT rows and BUFFER same-timestamp split candidates.
      // We emit only one canonical SPLIT row per Account + Date + Time so broker
      // partner rows do not double-apply the split downstream.
      if (type === "RAD") {
        const splitCheck = parseRadSplitDescription(getField(r, "DESCRIPTION"));
        if (splitCheck.isSplit) {
          const splitDateIso = normalizeDate(getField(r, "DATE"));
          const splitTimeHHmmss =
            normalizeTimeHHmmss(getField(r, "TIME")) || "000000";
          const splitTs = toDateObject(splitDateIso, splitTimeHHmmss);
          const splitAccount = toStr(getField(r, "Account"))
            .trim()
            .toUpperCase();

          // IMPORTANT: do NOT include ticker in this key.
          // A split + rename pair can arrive at the same timestamp under two different symbols.
          const splitBufferKey = [
            splitAccount,
            splitDateIso,
            splitTimeHHmmss,
          ].join("|");

          // Look up the SplitAdjustments entry for this ticker + date.
          const splitDescUpper = String(getField(r, "DESCRIPTION") ?? "")
            .trim()
            .toUpperCase();

          // Extract ticker from RAD description — walk tokens right-to-left,
          // skip all known split keywords, then try CusipMap if the token looks like a CUSIP.
          const tokensRad = splitDescUpper.match(/[A-Z][A-Z0-9]{0,9}/g) || [];
          const SPLIT_SKIP = new Set([
            "MANDATORY",
            "REVERSE",
            "SPLIT",
            "FORWARD",
            "STOCK",
            "WITH",
            "SHARES",
            "INC",
            "CORP",
            "LTD",
            "LLC",
            "THE",
            "EFF",
            "EFFECTIVE",
            "XXX",
            "OXXXREVERSE",
            "XXXREVERSE",
            "XASX",
            "XTSE",
            "XLON",
            "XNAS",
            "XNYS",
            "XOTC",
            "PINK",
          ]);

          let splitTicker = "";
          for (let t = tokensRad.length - 1; t >= 0; t--) {
            const tok = tokensRad[t];
            if (SPLIT_SKIP.has(tok)) continue;
            if (!isNaN(tok)) continue;
            if (tok.length < 1) continue;

            const cusipCandidate = normalizeCusip(tok);
            if (looksLikeCusip(cusipCandidate) && /\d/.test(cusipCandidate)) {
              const mapped = cusipMap[cusipCandidate];
              if (mapped) {
                splitTicker = normalizeSymbol(mapped);
              } else {
                importIssuesAdd(
                  ctx,
                  "WARN",
                  "TosTop",
                  "RAD",
                  "DESCRIPTION",
                  getField(r, "DESCRIPTION"),
                  `RAD SPLIT ticker token looks like a CUSIP (${tok}) but is not in CusipMap. ` +
                    `Add it to CusipMap to resolve the ticker.`,
                );
              }
              break;
            }

            splitTicker = tok;
            break;
          }

          const adj = splitAdjustments.find(
            (a) => a.ticker === splitTicker && a.dateIso === splitDateIso,
          );

          if (!adj) {
            importIssuesAdd(
              ctx,
              "WARN",
              "TosTop",
              "RAD",
              "DESCRIPTION",
              getField(r, "DESCRIPTION"),
              `RAD SPLIT detected for ${splitTicker} on ${splitDateIso} but no matching entry found in SplitAdjustments sheet. ` +
                `Add a row: Ticker=${splitTicker} | Split Date=${splitDateIso} | Ratio Numerator=? | Ratio Denominator=? | Split Type=${splitCheck.isReverse ? "REVERSE" : "FORWARD"}`,
            );
          }

          const splitCandidate = {
            account: splitAccount,
            splitDateIso: splitDateIso,
            splitTimeHHmmss: splitTimeHHmmss,
            splitTs: splitTs,
            splitTicker: splitTicker,
            splitCheck: splitCheck,
            adjustmentMatched: !!adj,
            adjustment: adj || null,
            rawDescription: getField(r, "DESCRIPTION"),
            rowNumber: i + 2,
          };

          if (!radSplitCandidateBuffer[splitBufferKey]) {
            radSplitCandidateBuffer[splitBufferKey] = [];
          }
          radSplitCandidateBuffer[splitBufferKey].push(splitCandidate);

          importIssuesAdd(
            ctx,
            "INFO",
            "TosTop",
            "RAD",
            "DESCRIPTION",
            getField(r, "DESCRIPTION"),
            `${splitCheck.isReverse ? "REVERSE SPLIT" : "FORWARD SPLIT"} buffered for ${splitTicker || "[blank ticker]"} on ${splitDateIso} ${splitTimeHHmmss}: ` +
              `parsedQty=${splitCheck.parsedQty}, explicitPre=${splitCheck.looksLikePreLeg ? "Y" : "N"}, ` +
              `explicitPost=${splitCheck.looksLikePostLeg ? "Y" : "N"}, adjustmentMatched=${adj ? "Y" : "N"}`,
          );

          continue; // handled — do not fall through to other RAD logic
        }
        // ... rest of your existing RAD handler continues here (removal of option, etc.)
      }

      if (type === "RAD") {
        const parsedRad = parseRadRemovalDescription(desc);
        if (parsedRad) {
          if (parsedRad.qty !== null && parsedRad.qty !== undefined)
            qtyOut = parsedRad.qty;
          if (parsedRad.exp instanceof Date && !isNaN(parsedRad.exp.getTime()))
            expOut = parsedRad.exp;
          if (parsedRad.strike !== null && parsedRad.strike !== undefined)
            strikeOut = parsedRad.strike;
          if (parsedRad.underlying) symbolOut = parsedRad.underlying;
        }
      }
      const corpPhraseHit = applyCorpActionsToTopRow(desc, corpMap);
      if (corpPhraseHit.handled) {
        if (corpPhraseHit.resultSymbol) {
          symbolOut = corpPhraseHit.resultSymbol;
        }
      } else {
        // Token scan for CUSIP-like tokens (standalone 9-char tokens only)
        // Using word boundaries avoids false positives inside OCC option symbols like ".UUUU240920P5".
        const hits =
          String(desc)
            .toUpperCase()
            .match(/\b[0-9A-Z]{9}\b/g) || [];
        for (let h = 0; h < hits.length; h++) {
          const tok = hits[h];

          const corpTokHit = applyCorpTokenRule(tok, corpMap);
          if (corpTokHit.handled) {
            if (corpTokHit.resultSymbol) symbolOut = corpTokHit.resultSymbol;
            continue;
          }

          // added to fix CUSIP leading character issue to resolve IE with a leading alpha character is being tagged as that letter in 'Symbol' eg Y1146L125 and Y3005A109 are tagged as Y, G8377A108 is tagged as just G
          const cusipKey = normalizeCusip(tok);
          if (looksLikeCusip(cusipKey) && /\d/.test(cusipKey)) {
            const mapped = cusipMap[cusipKey];
            if (mapped) {
              symbolOut = normalizeSymbol(mapped);
            } else if (!missingCusipSeenTosTop[cusipKey]) {
              missingCusipSeenTosTop[cusipKey] = true;
              importIssuesAdd(
                ctx,
                "MISSINGCUSIPMAP",
                "TosTop",
                i + 2,
                "DESCRIPTION",
                cusipKey,
                "CUSIP-like token found in TosTop DESCRIPTION but not in CusipMap.",
              );
            }
          }
        }
      }

      unifiedNonTrades.push({
        Account: Account, // <-- NEW
        Date: dateIso,
        Time: timeHHmmss,
        Timestamp: ts,

        Action: type,
        Symbol: symbolOut,
        Description: desc,

        Spread: "",
        Quantity: qtyOut,
        Price: "",

        NetPrice: "",

        Side: "",
        PosEffect: "",
        Exp: expOut,
        Strike: strikeOut,
        OrderType: "",

        MiscFees: miscFees,
        FeesComm: feesComm,
        Amount: amount,
      });
    }

    Object.keys(radSplitCandidateBuffer)
      .sort()
      .forEach(function (splitBufferKey) {
        const candidates = radSplitCandidateBuffer[splitBufferKey] || [];
        if (!candidates.length) return;

        const built = buildCanonicalSplitRowFromCandidates_(candidates);
        if (!built || !built.rowObj || !built.chosen) return;

        unifiedNonTrades.push(built.rowObj);

        importIssuesAdd(
          ctx,
          "INFO",
          "TosTop",
          "RAD",
          "DESCRIPTION",
          built.chosen.rawDescription,
          `Canonical RAD SPLIT emitted for ${built.chosen.splitTicker || "[blank ticker]"} on ${built.chosen.splitDateIso} ${built.chosen.splitTimeHHmmss}. ` +
            `Candidate count=${candidates.length}. ratio=${built.ratioLabel}. preQty=${built.preQty ?? "?"}. ` +
            `postQty=${built.postQty ?? "?"}. delta=${built.qtyDelta ?? "?"}. ` +
            `Adjustment source=${built.adjSource ? built.adjSource.ticker : "[none]"}.`,
        );

        if (candidates.length > 1) {
          candidates.forEach(function (candidate) {
            if (candidate === built.chosen) return;

            importIssuesAdd(
              ctx,
              "INFO",
              "TosTop",
              "RAD",
              "DESCRIPTION",
              candidate.rawDescription,
              `Duplicate same-timestamp RAD SPLIT ignored. ` +
                `Kept ${built.chosen.splitTicker || "[blank ticker]"} pre=${built.preQty ?? "?"} post=${built.postQty ?? "?"}, ` +
                `ignored ${candidate.splitTicker || "[blank ticker]"} parsedQty=${candidate.splitCheck && candidate.splitCheck.parsedQty != null ? candidate.splitCheck.parsedQty : "?"}.`,
            );
          });
        }
      });

    // =========================================================================
    // 8) COMBINE + SORT BY TIMESTAMP
    //    Merge the trade rows (from step 6) and the non-trade rows (from step 7),
    //    then sort everything by Timestamp so downstream Phase 2 / Phase 3 logic
    //    sees a single chronological sequence. Multi-leg groups stay adjacent.
    // =========================================================================
    pipelineTimingLog(
      "buildUnifiedImportV3 nontrade",
      tNontrade,
      "unifiedNonTrades=" + unifiedNonTrades.length,
    );
    const tSort = pipelineTimingNow();

    const all = unifiedTrades.concat(unifiedNonTrades);

    all.sort((a, b) => {
      // 1) Timestamp is authoritative
      const at = a.Timestamp instanceof Date ? a.Timestamp.getTime() : 0;
      const bt = b.Timestamp instanceof Date ? b.Timestamp.getTime() : 0;
      if (at !== bt) return at - bt;

      // 2) Keep DT/LT stable if both accounts are present in one Schwab Import
      const accA = toStr(a.Account).trim().toUpperCase();
      const accB = toStr(b.Account).trim().toUpperCase();
      if (accA !== accB) return accA.localeCompare(accB);

      // 3) If both are trade rows and have the same tradeGroupKey, keep legs adjacent + ordered
      const gA = toStr(a.tradeGroupKey).trim();
      const gB = toStr(b.tradeGroupKey).trim();

      // Prefer grouped trades before non-trades when timestamp ties
      if (gA && !gB) return -1;
      if (!gA && gB) return 1;

      // If both have group keys, group them together
      if (gA && gB) {
        if (gA !== gB) return gA.localeCompare(gB);

        const ka = legOrderKeyForOutput(a);
        const kb = legOrderKeyForOutput(b);
        for (let i = 0; i < ka.length; i++) {
          if (ka[i] < kb[i]) return -1;
          if (ka[i] > kb[i]) return 1;
        }
        return 0;
      }

      // 4) Fallback: stable string compare
      const as = String(a.Date) + String(a.Time) + String(a.Description);
      const bs = String(b.Date) + String(b.Time) + String(b.Description);
      return as.localeCompare(bs);
    });

    // =========================================================================
    // 9) ASSEMBLE 2-D OUTPUT ARRAY + WRITE TO "Schwab Import"
    //    Convert the named-property objects in all[] into a plain 2-D array that
    //    matches the canonical Schwab Import header order exactly, then write it.
    //    Timestamp is the authoritative source for Date and Time columns.
    // =========================================================================

    pipelineTimingLog("buildUnifiedImportV3 sort", tSort, "all=" + all.length);
    const tWrite = pipelineTimingNow();

    const outSh = mustGetSheet("Schwab Import");
    const prevLastRow = outSh.getLastRow();
    const prevLastCol = outSh.getLastColumn();

    // Clear the sheet completely before writing (except row 1 which we overwrite)
    if (prevLastRow > 0) outSh.clearContents();

    // Canonical Schwab Import headers — must match the project schema exactly.
    const headers = [
      "Account",
      "Date",
      "Time",
      "Time Stamp",
      "Action",
      "Symbol",
      "Description",
      "Spread",
      "Quantity",
      "Price",
      "Net Price",
      "Side",
      "Pos Effect",
      "Exp",
      "Strike",
      "Order Type",
      "Misc Fees",
      "Fees Comm",
      "Amount",
    ];

    // Map each unified object to a positional array in header order.
    // Internal-only fields (tradeGroupKey, optType) are intentionally excluded.
    const out = [headers].concat(
      all.map((r) => [
        r.Account ?? "",
        // Date derived from Timestamp (authoritative) — written as a value, no formula needed.
        r.Timestamp instanceof Date && !isNaN(r.Timestamp.getTime())
          ? Utilities.formatDate(
              r.Timestamp,
              Session.getScriptTimeZone(),
              "MM/dd/yyyy",
            )
          : (r.Date ?? ""),
        // Time derived from Timestamp — written as a value, no formula needed.
        r.Timestamp instanceof Date && !isNaN(r.Timestamp.getTime())
          ? Utilities.formatDate(
              r.Timestamp,
              Session.getScriptTimeZone(),
              "HH:mm",
            )
          : (r.Time ?? ""),
        r.Timestamp instanceof Date && !isNaN(r.Timestamp.getTime())
          ? r.Timestamp
          : "",
        r.Action ?? "",
        r.Symbol ?? "",
        r.Description ?? "",
        r.Spread ?? "",
        r.Quantity ?? "",
        r.Price ?? "",
        r.NetPrice ?? "",
        r.Side ?? "",
        r.PosEffect ?? "",
        formatExpForSchwabImport(r.Exp),
        r.Strike ?? "",
        r.OrderType ?? "",
        r.MiscFees ?? "",
        r.FeesComm ?? "",
        r.Amount ?? "",
      ]),
    );

    // IMPORTANT DOWNSTREAM GUIDE FOR BLOCK LOGIC / VALIDATIONS / P&L (copy/paste this into your helper if you want)
    // For every synthetic STOCK row created from EXERCISE/ASSIGN:
    // if (row.Spread === 'STOCK' && row.Description.includes('Synthetic STOCK leg from')) {
    //   block.shareDelta += (row.Side === 'BUY' ? +row.Quantity : -row.Quantity); // +100 or -100 shares
    //   const optionRow = previousRowWithSameGroupKey; // same Timestamp + tradeGroupKey
    //   block.optionContractDelta += (optionRow.Side === 'SELL' ? -1 : +1); // -1 contract
    //   // Fees/Amount are already on this STOCK row from TosTop — use them directly for realized P/L
    // }
    // This keeps blocks clean and inventory counts correct in runUnifiedBlockLogic, fillRealizedAndPercentPnLOnBlockClose, etc.
    // ====================== END PERFORMANCE + GUIDE ======================

    // IMPORTANT NOTE FOR PHASE 2 & 3 (newer scripter friendly):
    // After this point "Schwab Import" is CANONICAL.
    // mapSchwabImportByHeadersV3 should read ONLY from here.
    // Do NOT re-read TosTop/TosTrades unless you are running a deliberate repair step.
    // Timestamp (column D) is the authoritative sort/group key.
    // ====================== END SAFETY CHECK ======================

    // Format key columns before writing (prevents leading-zero loss + makes timestamps readable).
    const outRowsPlanned = Math.max(1, out.length);

    // NEW: one single range call for everything — faster on 15k+ rows
    const formatRange = outSh.getRange(1, 1, outRowsPlanned, headers.length);
    tosFormatHeaderColumnsAsText(
      outSh,
      headers,
      ["Account", "Time", "Symbol"],
      outRowsPlanned,
      1,
    );

    // Apply date/time formats in one shot
    formatRange.setNumberFormat("@"); // default everything to text first
    const timeStampColA1 = headers.indexOf("Time Stamp") + 1;
    const dateColA1 = headers.indexOf("Date") + 1;
    if (timeStampColA1 > 0)
      outSh
        .getRange(1, timeStampColA1, outRowsPlanned, 1)
        .setNumberFormat("mm/dd/yyyy hh:mm:ss");
    if (dateColA1 > 0)
      outSh
        .getRange(1, dateColA1, outRowsPlanned, 1)
        .setNumberFormat("mm/dd/yyyy");

    // Optional debug timing (controlled by your existing Settings > Toggle TOS Import DEBUG Alerts)
    const debugTiming =
      String(getSetting("TOS_IMPORT_DEBUG_ALERTS", "0")) === "1";
    if (debugTiming) {
      const startWrite = new Date();
      outSh.getRange(1, 1, out.length, headers.length).setValues(out);
      const writeMs = new Date() - startWrite;
      importIssuesSetMetric(ctx, "WriteTimeMs", writeMs);
      tosMaybeDebugAlert("Write to Schwab Import took " + writeMs + " ms");
    } else {
      // normal fast path
      outSh.getRange(1, 1, out.length, headers.length).setValues(out);
    }

    // Restore Date/Time formulas derived from Timestamp (Timestamp is column D).
    // outSh.getRange('B2').setFormula('=ARRAYFORMULA(IF($D2:$D="",,INT($D2:$D)))');
    // outSh.getRange('C2').setFormula('=ARRAYFORMULA(IF($D2:$D="",,TEXT($D2:$D,"HH:mm:ss")))');

    // ── Date/Time rendering check ──────────────────────────────────────────────
    // Fires a popup if any data row is missing Date or Time after the write.
    // Schwab Import row 2 is the first data row (row 1 is the header).
    checkMissingDateTimeAndAlert(
      outSh,
      2,
      "buildUnifiedImportV3 → Schwab Import",
    );

        pipelineTimingLog(
      "buildUnifiedImportV3 write",
      tWrite,
      "outRows=" + out.length,
    );
    // ─────────────────────────────────────────────────────────────────────────

    // Clear leftovers AFTER successful write.
    const outRows = out.length;
    const outCols = headers.length;

    if (prevLastRow > outRows) {
      outSh
        .getRange(outRows + 1, 1, prevLastRow - outRows, prevLastCol)
        .clearContent();
    }
    if (prevLastCol > outCols) {
      const rowsToClear = Math.max(prevLastRow, outRows);
      outSh
        .getRange(1, outCols + 1, rowsToClear, prevLastCol - outCols)
        .clearContent();
    }

    // =========================================================================
    // 10) METRICS, FLUSH ISSUES, AND SUCCESS MESSAGE
    //    Record how many rows were written, how many enrichment matches failed,
    //    queue health, and retagging stats, then flush the Import Issues log
    //    and show a short UI summary.
    // =========================================================================
    importIssuesSetMetric(ctx, "RowsWrittenExclHeader", all.length);
    importIssuesSetMetric(ctx, "UnifiedTrades", unifiedTrades.length);
    importIssuesSetMetric(ctx, "UnifiedNonTrades", unifiedNonTrades.length);
    // ---------------------------- Enrichment summary metrics ----------------------------
    // Remaining (unconsumed) TosTop TRD rows still sitting in the queues.
    let remainingTopTrdQueueRows = 0;
    Object.keys(topTradeQueueByDateTime).forEach((k) => {
      const bucket = topTradeQueueByDateTime[k];
      if (bucket && bucket.length) remainingTopTrdQueueRows += bucket.length;
    });

    // How many "not found" cases were logged (deduped)
    importIssuesSetMetric(
      ctx,
      "TradeEnrichmentNotFoundKeys",
      Object.keys(missingTradeEnrichmentSeen).length,
    );
    importIssuesSetMetric(
      ctx,
      "TradeEnrichmentDataGapKeys",
      Object.keys(missingTradeEnrichmentDataGapSeen).length,
    );

    // Queue health
    importIssuesSetMetric(
      ctx,
      "TopTrdRowsUnconsumed",
      remainingTopTrdQueueRows,
    );
    importIssuesSetMetric(
      ctx,
      "TopTrdRowsConsumedEst",
      Math.max(0, Number(topTrdRowsQueued || 0) - remainingTopTrdQueueRows),
    );

    // Spread retagging metrics
    importIssuesSetMetric(
      ctx,
      "SpreadRetaggedRowsCount",
      icRetagMetrics.spreadRetaggedRowsCount,
    );
    importIssuesSetMetric(
      ctx,
      "SpreadRetagSkippedNoOpenIcCount",
      icRetagMetrics.spreadRetagSkippedNoOpenIcCount,
    );

    importIssuesFlush(ctx);

    //  since this run can include multiple accounts, derive a display value from the output rows.
    const accountsUsed = Array.from(
      new Set(
        all
          .map((r) =>
            String(r.Account || "")
              .trim()
              .toUpperCase(),
          )
          .filter((s) => s),
      ),
    ).sort();

    uiAlertSafe(
      "Unified Import v3 built " +
        all.length +
        " rows.\n" +
        "Accounts: " +
        (accountsUsed.length ? accountsUsed.join(", ") : "(none)") +
        "\n" +
        'Issues logged to "Import Issues": ' +
        (ctx.issues ? ctx.issues.length : 0),
    );
  } catch (err) {
    importIssuesAdd(
      ctx,
      "ERROR",
      "",
      "",
      "Exception",
      "",
      String(err && err.stack ? err.stack : err),
    );
    importIssuesFlush(ctx);
    throw err;
  } finally {
    pipelineTimingLog("buildUnifiedImportV3", tBuild);
    lock.releaseLock();
  }
}
