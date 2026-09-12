/**
 * mapSchwabImportByHeadersV3.js
 *
 * Phase 2 – Mapping
 *
 * High-level job:
 *   Read the canonical "Schwab Import" sheet (produced by Phase 1)
 *   and write a fully expanded, header-driven table into "Schwab Mapping".
 *
 * Key responsibilities:
 *   - Map every import column into the full Schwab Mapping schema
 *   - Tag Account Actions and Corporate Actions (both can be present on one row)
 *   - Resolve tickers for symbol-change and TDA-era corporate-action rows
 *   - Normalize Strategy Type (baseline + multi-leg post-processors)
 *   - Propagate Net Amount across spread legs when Schwab only supplied it once
 *   - Sort by Trade Time Stamp (authoritative sequencing key)
 *   - Log all problems and metrics to "Schwab Mapping Issues"
 *
 * Important design rules still in force:
 *   - No separate accountMode processing — Account is read per-row from Schwab Import
 *   - "Trade Time Stamp" must be a real Date object (Date/Time columns are display helpers only)
 *   - MARK TO THE MARKET rows are dropped silently (configurable)
 *
 * Related files:
 *   - BuildUnifiedImportV3.js          (Phase 1 – produces Schwab Import)
 *   - ImportIssues.js                  (shared logging engine; mappingIssues* wrappers)
 *   - Phase3Processing.js              (Phase 3 – consumes Schwab Mapping)
 *   - AuditSchwabMappingV3.js          (pre-Phase-3 gate; same Issues sheet)
 *   - MapAccountActions.js             (Group A – Account Actions helpers)
 *   - MapCorpActionRules.js            (Group E – corp keyword rules + scanner)
 *   - MapTradeFields.js                (Group B – trade field builders)
 *   - MapCashFlow.js                   (Group F – Cash Map)
 *   - MapCorpTickerResolve.js          (Group D – CUSIP / rename / Corp Action Map)
 *   - MapStrategyType.js               (Group G – Strategy Type post-processors)
 *   - MapNetAmount.js                  (Group H – Net Amount post-processors)
 *   - SettingsService.js
 * Current focus: clear high-level documentation before any structural refactoring.
 */

// =========================
// Sheet names (exact tabs)
// =========================
const SHEET_SCHWAB_IMPORT = "Schwab Import";
const SHEET_SCHWAB_MAPPING = "Schwab Mapping";
const SHEET_CASH_MAP = "Cash Map";
const SHEET_MAPPING_ISSUES = "Schwab Mapping Issues";

// =========================
// Behavior toggles
// =========================

// If true, clears entire "Schwab Mapping" (row 2+) and rewrites fresh rows.
// This is the safe choice if you sort by timestamp (row order changes each run).
const MAP_V3_REBUILD_MAPPING_SHEET = true;

// If true, drops rows whose Description contains "MARK TO THE MARKET".
// Per your rule: drop entirely (no output row, and no issue row).
const MAP_V3_DROP_MARK_TO_MARKET = true;

// =========================
// Spread group types used by all three post-processors
// (add new spreads here in the future — one place only)
const SPREAD_GROUP_TYPES = ["VERTICAL", "IRON CONDOR", "BUTTERFLY"];

// =========================
// "Schwab Import" headers (source schema for Phase 2)
// =========================
const SCHWAB_IMPORT_HEADERS = [
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

// =========================
// "Schwab Mapping" headers (full schema)
// =========================
const SCHWAB_MAPPING_HEADERS = [
  "Account",
  "Trade Time Stamp",
  "Trade Date",
  "Trade Time",
  "Ticker",
  "Action",
  "Order Type",
  "Description",
  "Quantity",
  "Entry Price",
  "Net Price",
  "Net Amount",
  "Total Cost",
  "Signed Quantity",
  "Realized P&L",
  "Percent P&L",
  "Unrealized P&L",
  "Opening Date",
  "Closing Date",
  "Trade Status",
  "Trade Duration",
  "Strategy Type",
  "Option Strike",
  "Option Expiration",
  "Call/Put",
  "Option Contract",
  "Underlying Price Entry",
  "Underlying Price Exit",
  "Delta",
  "Misc Fees",
  "Fees & Comm",
  "Corporate Actions",
  "Account Actions",
  "Cash Flow Direction",
  "Transfer Type",
  "Source",
  "Notes",
  "Reason for Entry",
  "Reason for Exit",
  "Emotional State",
  "Running Position Quantity",
  "Trade Group ID",
  "Position ID",
  "Spread Group ID",
  "Block Start Flag",
  "Block Number",
  "Block Close Flag/P&L",
];

// =========================
// Phase 2 allowed output columns (everything else remains blank here)
// =========================
const SCHWAB_MAPPING_SCOPE_HEADERS = [
  "Account",
  "Trade Time Stamp",
  "Trade Date",
  "Trade Time",
  "Ticker",
  "Action",
  "Order Type",
  "Description",
  "Quantity",
  "Entry Price",
  "Net Price",
  "Net Amount",
  "Total Cost",
  "Signed Quantity",
  "Opening Date",
  "Closing Date",
  "Strategy Type",
  "Option Strike",
  "Option Expiration",
  "Call/Put",
  "Option Contract",
  "Misc Fees",
  "Fees & Comm",
  "Corporate Actions",
  "Account Actions",
  "Cash Flow Direction",
  "Transfer Type",
];

// =========================================================================
// MAIN ENTRY POINT
// =========================================================================
/**
 * mapSchwabImportByHeadersV3()
 *
 * Called from the DB Tools menu (and from the full pipeline).
 *
 * High-level flow:
 *   1. Start a Mapping Issues run context
 *   2. Read Schwab Import + supporting sheets (Cash Map, CusipMap, Corp Action Map…)
 *   3. Transform every import row into a Schwab Mapping row (main loop)
 *   4. Run post-processors (Strategy Type, Net Amount, missing-amount warnings)
 *   5. Sort by Trade Time Stamp
 *   6. Write the result to "Schwab Mapping" and flush metrics/issues
 */
function mapSchwabImportByHeadersV3() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();

  // Create a single run context for all metrics + issues.
  // This is the core of the "Import Issues" style approach.
  const ctx = mappingIssuesStart("mapSchwabImportByHeadersV3");

  // We want the issues sheet to show "what was this run doing".
  mappingIssuesSetMetric(ctx, "SourceSheet", SHEET_SCHWAB_IMPORT);
  mappingIssuesSetMetric(ctx, "DestSheet", SHEET_SCHWAB_MAPPING);
  mappingIssuesSetMetric(ctx, "OutputSheet", SHEET_MAPPING_ISSUES);

  // Metrics counters (kept as numbers so they're easy to scan/filter in Schwab Mapping Issues)
  let droppedMarkToMarket = 0;
  let rowsRead = 0;
  let rowsWritten = 0;
  let tradeRowCount = 0;
  let nonTradeRowCount = 0;
  let missingTimestampCount = 0;
  let invalidAccountCount = 0;

  // ── NEW: Corp action ticker resolution counters ───────────────────────────
  // These tell you exactly how each corp action row got (or didn't get) its Ticker.
  let corpTickerFromSymbol = 0; // Symbol column had the ticker — fast path, no fallback needed
  let corpTickerFromTilde = 0; // Extracted from ~TICKER pattern (TDA "ORDINARY DIVIDEND~JEPI")
  let corpTickerFromMap = 0; // Resolved via Corp Action Map lookup
  let corpTickerFromDesc = 0; // Extracted by tokenizer (RAD splits/mergers with embedded ticker)
  let corpTickerSuspect = 0; // Tokenizer returned a word > 5 chars — likely wrong, left blank
  let corpTickerUnresolved = 0; // All methods failed — WARN logged, map entry needed

  // These help you quickly verify that keyword tagging is actually working.
  const accountActionCounts = {}; // tag -> count
  const corpActionCounts = {}; // tag -> count

  // =========================================================================
  // HIGH-LEVEL ROADMAP OF THIS FUNCTION
  //
  //   A) Setup          – sheets, header maps, keyword rules, Cash Map
  //   B) Main loop      – one import row → one mapping row (tagging + field fill)
  //   C) Post-processors– Strategy Type (spread groups + position tracker),
  //                       Net Amount propagation, missing-amount warnings
  //   D) Sort + Write   – authoritative Trade Time Stamp order → Schwab Mapping
  //   E) Metrics/Flush  – counters + Mapping Issues sheet
  // =========================================================================

  // Wrap everything so that even if something throws, you still get an ERROR issue row + metrics flushed.
  try {
    // --- Get required sheets ---
    const importSheet = ss.getSheetByName(SHEET_SCHWAB_IMPORT);
    if (!importSheet)
      throw new Error("Missing required sheet: " + SHEET_SCHWAB_IMPORT);

    const mappingSheet = ss.getSheetByName(SHEET_SCHWAB_MAPPING);
    if (!mappingSheet)
      throw new Error("Missing required sheet: " + SHEET_SCHWAB_MAPPING);

    // --- Read Schwab Import ---
    const importLastRow = importSheet.getLastRow();
    const importLastCol = importSheet.getLastColumn();

    if (importLastRow < 2) {
      mappingIssuesSetMetric(ctx, "SourceRowsReadExclHeader", 0);
      mappingIssuesSetMetric(ctx, "RowsWrittenExclHeader", 0);
      mappingIssuesFlush(ctx);
      uiAlertSafe("Schwab Import has no rows to map.");
      return;
    }

    const importHeaders = importSheet
      .getRange(1, 1, 1, importLastCol)
      .getValues()[0];
    const importHeaderMap = buildHeaderIndexMap(importHeaders);

    // Header sanity checks (missing + duplicates)
    assertNoDuplicateHeaders(importHeaders, SHEET_SCHWAB_IMPORT, ctx);
    requireHeaders(importHeaderMap, SCHWAB_IMPORT_HEADERS, SHEET_SCHWAB_IMPORT);

    const importData = importSheet
      .getRange(2, 1, importLastRow - 1, importLastCol)
      .getValues();
    rowsRead = importData.length;

    // --- Read Schwab Mapping headers (row 1 must exist) ---
    const mappingLastCol = mappingSheet.getLastColumn();
    const mappingHeaders = mappingSheet
      .getRange(1, 1, 1, mappingLastCol)
      .getValues()[0];
    const mappingHeaderMap = buildHeaderIndexMap(mappingHeaders);

    assertNoDuplicateHeaders(mappingHeaders, SHEET_SCHWAB_MAPPING, ctx);
    requireHeaders(
      mappingHeaderMap,
      SCHWAB_MAPPING_HEADERS,
      SHEET_SCHWAB_MAPPING,
    );
    requireHeaders(
      mappingHeaderMap,
      SCHWAB_MAPPING_SCOPE_HEADERS,
      SHEET_SCHWAB_MAPPING + " (scope columns)",
    );

    // --- Load Cash Map lookup (Account Actions -> {dir,type}) ---
    const cashFlowMap = buildCashFlowMapFromSheet(ss);
    mappingIssuesSetMetric(
      ctx,
      "CashMapUniqueActions",
      Object.keys(cashFlowMap).length,
    );

    // --- Load Corp Action Map company-name description => ticker for TDA-era rows ---
    // Returns an array sorted longest-pattern-first so more-specific entries win...
    // If the sheet doesn't exist yet, returns [] and the fallback chain still works gracefully.
    const corpActionMap = buildCorpActionMapFromSheet(ss);
    mappingIssuesSetMetric(ctx, "CorpActionMapEntries", corpActionMap.length);

    // --- Load CusipMap for symbol/ticker change and CUSIP-based corp-action resolution ---
    // WHY
    // Schwab/TDA sometimes emits a CUSIP instead of a ticker on spin-offs, renames,
    // and certain corporate action rows. We resolve those to the canonical ticker here
    // so downstream Phase 3 block logic can keep one continuous stock position.
    const cusipMapV3 = buildCusipMapFromSheetV3(ss);
    mappingIssuesSetMetric(
      ctx,
      "CusipMapEntries",
      Object.keys(cusipMapV3).length,
    );

    // Corp action types that have NO underlying ticker by design.
    // "Cash Interest" = money market interest, "Cash In Lieu" = fractional share cash-out.
    // These should NEVER trigger the fallback chain below.

    // add the two interest adjustment types the I added to corp action rules:
    const CORP_ACTIONS_NO_TICKER = new Set([
      "Cash Interest", // Cash Alternatives Interest rows — no underlying security
      "Cash In Lieu", // Fractional share cash-out — amount only, no underlying
      "Interest Adjustment", // FREE BALANCE INTEREST ADJUSTMENT rows — TDA-era, no security
      "Margin Interest Adjustment", // MARGIN INTEREST ADJUSTMENT rows — TDA-era margin fee, no security
    ]);

    // --- Build tagging maps (keyword lists) ---
    const accountActionRules = getAccountActionsKeywordRulesV3(); // ordered rules
    const corpMap = getCorpActionsKeywordRulesV3(); // ordered rules

    // --- Output collection ---
    // outItems holds:
    // - row: the final Schwab Mapping row array
    // - importRowNum: original Schwab Import row number (stable tie-breaker for identical timestamps)
    // - spreadRaw: raw Spread label from Schwab Import (used later for group-based strategy normalization)
    const outItems = [];

    // =========================================================================
    // B) MAIN TRANSFORM LOOP
    //    One Schwab Import row → one Schwab Mapping row.
    //    Handles early drops (MARK TO THE MARKET), tagging, trade vs non-trade
    //    branching, symbol-change special cases, and baseline Strategy Type.
    // =========================================================================
    // Corp action types that legitimately have no underlying ticker.
    // These are pure cash/interest transactions — do NOT attempt description-based extraction for them.
    // All other corp action types (splits, mergers, dividends, reorganizations, transfers, etc.)
    // SHOULD have a ticker and will use the description fallback when Symbol is blank.

    for (let i = 0; i < importData.length; i++) {
      const importRowNum = i + 2; // actual sheet row number (header is row 1)
      const r = importData[i];

      // Read Description first so we can drop rows early
      const desc = String(r[col(importHeaderMap, "Description")] || "");
      const descLower = desc.trim().toLowerCase();

      // NEW: Detect special TOS synthetic stock rows from option exercise/assignment
      // These have Spread=STOCK but contain the word "CALL" in Description.
      // We want to treat them as pure stock trades (no option errors).
      const isSyntheticExerciseStockLeg = descLower.includes(
        "synthetic stock leg from exercise",
      );

      // Drop "MARK TO THE MARKET" silently (per your rule)
      if (
        MAP_V3_DROP_MARK_TO_MARKET &&
        descLower.includes("mark to the market")
      ) {
        droppedMarkToMarket++;
        continue;
      }

      // Create ONE mapped row per input row
      const mapped = new Array(mappingHeaders.length).fill("");

      // --- Pull core import fields ---
      const accountRaw = String(r[col(importHeaderMap, "Account")] || "")
        .trim()
        .toUpperCase();
      const importTimeStamp = r[col(importHeaderMap, "Time Stamp")];
      const importDate = r[col(importHeaderMap, "Date")];
      const importTime = r[col(importHeaderMap, "Time")];

      const importAction = String(
        r[col(importHeaderMap, "Action")] || "",
      ).trim();
      const symbolRaw = String(r[col(importHeaderMap, "Symbol")] || "").trim();
      const spreadRaw = String(r[col(importHeaderMap, "Spread")] || "").trim();

      const qtyRaw = r[col(importHeaderMap, "Quantity")];
      const priceRaw = r[col(importHeaderMap, "Price")];
      const netPriceRaw = r[col(importHeaderMap, "Net Price")];
      const miscFeesRaw = r[col(importHeaderMap, "Misc Fees")];
      const feesCommRaw = r[col(importHeaderMap, "Fees Comm")];
      const amountRaw = r[col(importHeaderMap, "Amount")];

      // Parse Schwab Import Amount (often net for multi-leg spreads; may be blank on some legs)
      const importAmount = parseNumber(amountRaw);

      const sideRaw = String(r[col(importHeaderMap, "Side")] || "").trim();
      const posEffectRaw = String(
        r[col(importHeaderMap, "Pos Effect")] || "",
      ).trim();

      const expRaw = r[col(importHeaderMap, "Exp")];
      const strikeRaw = r[col(importHeaderMap, "Strike")];
      const orderTypeRaw = String(
        r[col(importHeaderMap, "Order Type")] || "",
      ).trim();

      // Authoritative sequencing key (must remain a real DateTime object)
      const ts = normalizeImportTimeStamp(
        importTimeStamp,
        importDate,
        importTime,
      );

      // =========================
      // Tagging: Account Actions + Corporate Actions (can both be present)
      // =========================
      const accountActionTag = deriveAccountActionTagV3(
        descLower,
        accountRaw,
        importAmount,
        accountActionRules,
      );
      const corpActionTag = findFirstKeywordTag(descLower, corpMap);

      // Track tag counts (this becomes a SUPER useful metric when something "stops tagging")
      if (accountActionTag)
        accountActionCounts[accountActionTag] =
          (accountActionCounts[accountActionTag] || 0) + 1;
      if (corpActionTag)
        corpActionCounts[corpActionTag] =
          (corpActionCounts[corpActionTag] || 0) + 1;

      // Detect trade rows:
      // Side + Pos Effect is the authority (BUY/SELL and OPEN/CLOSE).
      // A Corporate Actions tag is a label only — it must not wipe Quantity /
      // Entry Price on a real trade. That bug hit the Phase 1 synthetic
      // STOCK MERGER receipt: Description contains "STOCK MERGER", so the
      // keyword rules tagged it, isTrade went false, and Mapping wrote
      // Action=Buy with blank qty/price.
      //
      // RAD merger / dividend rows stay non-trade because they have blank
      // Side and Pos Effect, so isTradeBySidePosEffect() returns false.
      // Account Actions still vetoes (fees, journals, ACH).
      // DRIP no longer needs a special case; those rows already have
      // Side=BUY and Pos Effect=TO OPEN.
      const isTrade =
        !accountActionTag && isTradeBySidePosEffect(sideRaw, posEffectRaw);
      // =========================
      // Fill "Schwab Mapping" scope columns
      // =========================

      // Account (NEW: read from row, no ScriptProperties mode)
      mapped[col(mappingHeaderMap, "Account")] = accountRaw;

      if (accountRaw !== "DT" && accountRaw !== "LT") {
        invalidAccountCount++;
        mappingIssuesAdd(
          ctx,
          "ERROR",
          importRowNum,
          "Account",
          accountRaw,
          "Account must be DT or LT.",
        );
      }

      // Trade Time Stamp (DateTime value, not string)
      mapped[col(mappingHeaderMap, "Trade Time Stamp")] = ts || "";

      // Trade Date + Trade Time are for readability/filtering.
      // (Trade Time Stamp remains authoritative for sequencing/grouping.)
      if (ts instanceof Date && !isNaN(ts)) {
        mapped[col(mappingHeaderMap, "Trade Date")] = new Date(
          ts.getFullYear(),
          ts.getMonth(),
          ts.getDate(),
        );
        // AFTER — anchored to the spreadsheet's timezone, not the script runner's account
        const tz = ss.getSpreadsheetTimeZone();
        mapped[col(mappingHeaderMap, "Trade Time")] = Utilities.formatDate(
          ts,
          tz,
          "HHmm",
        );
      } else {
        // Still write something readable if possible
        mapped[col(mappingHeaderMap, "Trade Date")] =
          normalizeImportDateOnly(importDate) || "";
        mapped[col(mappingHeaderMap, "Trade Time")] =
          normalizeImportTimeOnly(importTime) || "";

        missingTimestampCount++;
        mappingIssuesAdd(
          ctx,
          "ERROR",
          importRowNum,
          "Time Stamp",
          String(importTimeStamp || ""),
          "Missing/invalid Time Stamp. Row was still written, but Trade Time Stamp is blank (sorting/grouping may be affected).",
        );
      }

      // ── Ticker — four-step fallback chain for corp action rows ─────────────────
      // Step 1: Symbol column — fast path, reliable for trade rows and modern Schwab rows.
      // Steps 2-5 only run for corp action rows where Symbol was blank.
      // CORP_ACTIONS_NO_TICKER types skip the fallback entirely (no underlying security).
      let ticker = extractTickerFromSymbol(symbolRaw);

      // AFTER — also fires for any DOI row or untagged RAD non-trade row with a blank ticker.
      //
      // Why widen to importAction === 'DOI':
      //   TDA-era dividend rows sometimes have NO keyword prefix at all — the description
      //   IS the company name: "JPMORGAN EQUITY PREMIUM INCOME ETF 4.18 US$".
      //   There is nothing to add to getCorpActionsKeywordRulesV3_() for these because
      //   the description itself is the key. The Corp Action Map is the only solution,
      //   but the map is inside this gate. Widening to DOI opens the gate for those rows.
      //
      // Why also include RAD + !isTrade:
      //   Safety net — a RAD non-trade row that somehow got past getCorpActionsKeywordRulesV3_()
      //   without a tag still needs its ticker resolved before going downstream.
      const isCorpIncomeRow =
        (corpActionTag && !CORP_ACTIONS_NO_TICKER.has(corpActionTag)) ||
        (importAction === "DOI" &&
          !CORP_ACTIONS_NO_TICKER.has(corpActionTag)) ||
        (importAction === "RAD" && !isTrade && !accountActionTag);

      if (isCorpIncomeRow) {
        if (ticker) {
          // Symbol already gave us the ticker — count it only when a corp action tag is involved.
          if (corpActionTag && !CORP_ACTIONS_NO_TICKER.has(corpActionTag)) {
            corpTickerFromSymbol++;
          }
        } else {
          // Step 2 — Tilde-prefix pattern: "ORDINARY DIVIDEND~JEPI" or "Ordinary Dividend~JEPI 3.43 US$"
          // The regex only matches short alpha tokens (≤6 chars) so "~JPMORGAN EQUITY..." won't fire.
          const tildeTicker = extractTickerFromTildePattern(desc);
          if (tildeTicker) {
            ticker = tildeTicker;
            corpTickerFromTilde++;
          }

          // Step 3 — Corp Action Map: company-name descriptions like "PFIZER INC 3.41 US$"
          if (!ticker) {
            const mapTicker = lookupTickerFromCorpActionMap(
              desc,
              corpActionMap,
            );
            if (mapTicker) {
              ticker = mapTicker;
              corpTickerFromMap++;
            }
          }

          // Step 4 — Tokenizer: RAD splits/mergers/transfers where the ticker IS embedded
          // as a clean all-caps token (e.g. "MANDATORY REVERSE SPLIT ACMR RATIO 1:10" → ACMR).
          // If the tokenizer returns a token > 5 chars it is almost certainly a company-name
          // word (e.g. "PFIZER"), not a ticker. Log a WARN and leave blank.
          if (!ticker) {
            const descTicker = extractTickerFromCorpActionDesc(desc);
            if (descTicker) {
              if (descTicker.length > 5) {
                // Suspect: too long to be a real ticker. Flag and leave blank.
                corpTickerSuspect++;
                mappingIssuesAdd(
                  ctx,
                  "WARN",
                  importRowNum,
                  "Ticker",
                  descTicker,
                  'Corp action Ticker extraction suspect — "' +
                    descTicker +
                    '" (' +
                    descTicker.length +
                    " chars) " +
                    "looks like a company name word, not a ticker symbol (max valid ticker = 5 chars). " +
                    'Add a row to Corp Action Map: Company Name Pattern = "' +
                    desc.substring(0, 60) +
                    '" | Ticker = ??? ' +
                    "then re-run mapSchwabImportByHeadersV3.",
                );
              } else {
                ticker = descTicker;
                corpTickerFromDesc++;
              }
            } else {
              // Step 5 — All methods failed. Log actionable WARN with copy-paste map entry hint.
              corpTickerUnresolved++;
              mappingIssuesAdd(
                ctx,
                "WARN",
                importRowNum,
                "Ticker",
                "",
                "Corp action Ticker unresolved — Symbol blank, no ~Ticker pattern found, no Corp Action Map match, " +
                  "and tokenizer returned nothing. " +
                  'Add to Corp Action Map: Company Name Pattern = "' +
                  desc.substring(0, 60) +
                  '" | Ticker = ??? ' +
                  "then re-run mapSchwabImportByHeadersV3.",
              );
            }
          }
        }
      }

      // Ticker alias: if the resolved ticker is listed as a Corp Action Map
      // pattern, use that row's Ticker. Example: QUALIFIED DIVIDEND~PKI
      // plus map row PKI | RVTY → write RVTY so Phase 3 stays on one position.
      // Exact pattern match only (PKI = PKI). Long company-name patterns
      // such as "F3 URANIUM CORP" are not aliases and are left alone.
      if (ticker && corpActionMap && corpActionMap.length) {
        const tkrKey = String(ticker).trim().toUpperCase();
        for (let a = 0; a < corpActionMap.length; a++) {
          const entry = corpActionMap[a];
          if (
            entry.pattern === tkrKey &&
            entry.ticker &&
            entry.ticker !== tkrKey
          ) {
            ticker = entry.ticker;
            corpTickerFromMap++;
            break;
          }
        }
      }

      mapped[col(mappingHeaderMap, "Ticker")] = ticker;

      // Order Type + Description
      mapped[col(mappingHeaderMap, "Order Type")] = orderTypeRaw;
      mapped[col(mappingHeaderMap, "Description")] = desc;

      // Fees (these are allowed Phase 2 scope columns)
      mapped[col(mappingHeaderMap, "Misc Fees")] = parseNumber(miscFeesRaw);
      mapped[col(mappingHeaderMap, "Fees & Comm")] = parseNumber(feesCommRaw);

      // Corporate + Account Actions tags (NEW: both can be present)
      mapped[col(mappingHeaderMap, "Corporate Actions")] = corpActionTag || "";
      mapped[col(mappingHeaderMap, "Account Actions")] = accountActionTag || "";

      // =========================================================================
      // Trade vs non-trade mapping
      //    Trade rows get Quantity / Entry Price / Signed Quantity / Action, etc.
      //    Non-trade rows (cash, corp actions, journals…) get the lighter mapping.
      // =========================================================================
      if (isTrade) {
        tradeRowCount++;

        // ---- Trade row ----
        const qty = Math.abs(parseNumber(qtyRaw) || 0);
        const entryPrice = parseNumber(priceRaw);

        mapped[col(mappingHeaderMap, "Quantity")] = qty || "";
        mapped[col(mappingHeaderMap, "Entry Price")] = entryPrice;

        // Net Price (import-provided net view for spreads; may be blank on many legs)
        mapped[col(mappingHeaderMap, "Net Price")] = parseNumber(netPriceRaw);

        // Net Amount (from Schwab Import Amount; often only present on one leg of a spread)
        mapped[col(mappingHeaderMap, "Net Amount")] = importAmount;

        // Signed Quantity: BUY => +qty, SELL => -qty
        mapped[col(mappingHeaderMap, "Signed Quantity")] = buildSignedQuantity(
          sideRaw,
          qty,
        );

        // Total Cost (leg-level): signedQty * entryPrice * multiplier
        // Equities/ETFs options use 100 multiplier; stocks use 1.
        const signedQty = Number(
          mapped[col(mappingHeaderMap, "Signed Quantity")],
        );
        const multiplier =
          String(spreadRaw || "")
            .trim()
            .toUpperCase() === "STOCK"
            ? 1
            : 100;

        let legTotal = "";
        if (!isNaN(signedQty) && !isNaN(entryPrice) && entryPrice !== "") {
          legTotal = signedQty * entryPrice * multiplier;
        }
        mapped[col(mappingHeaderMap, "Total Cost")] = legTotal;

        // Action: combine Side + Pos Effect (trades only)
        mapped[col(mappingHeaderMap, "Action")] = buildTradeAction(
          sideRaw,
          posEffectRaw,
        );

        // Opening/Closing Date based on Pos Effect (date-only)
        if (ts instanceof Date && !isNaN(ts)) {
          const dOnly = new Date(ts.getFullYear(), ts.getMonth(), ts.getDate());
          const pe = posEffectRaw.trim().toUpperCase();
          if (pe.includes("OPEN"))
            mapped[col(mappingHeaderMap, "Opening Date")] = dOnly;
          if (pe.includes("CLOSE"))
            mapped[col(mappingHeaderMap, "Closing Date")] = dOnly;
        }

        // Option fields (if present)
        const expDate = normalizeExpiration(expRaw);
        const strike = parseNumber(strikeRaw);
        const callPut = extractCallPut(symbolRaw, desc);

        // NEW: For synthetic exercise stock legs we intentionally clear option fields
        // (they are not real options — just TOS bookkeeping for the +100 shares)
        if (isSyntheticExerciseStockLeg) {
          mapped[col(mappingHeaderMap, "Option Strike")] = "";
          mapped[col(mappingHeaderMap, "Option Expiration")] = "";
          mapped[col(mappingHeaderMap, "Call/Put")] = "";
          mapped[col(mappingHeaderMap, "Option Contract")] = "";
        } else {
          if (expDate)
            mapped[col(mappingHeaderMap, "Option Expiration")] = expDate;
          if (strike !== "" && strike !== null && typeof strike !== "undefined")
            mapped[col(mappingHeaderMap, "Option Strike")] = strike;
          if (callPut) mapped[col(mappingHeaderMap, "Call/Put")] = callPut;
          // Option Contract (simple readable format)
          if (
            ticker &&
            expDate &&
            callPut &&
            strike !== "" &&
            strike !== null &&
            typeof strike !== "undefined"
          ) {
            mapped[col(mappingHeaderMap, "Option Contract")] =
              ticker +
              " " +
              Utilities.formatDate(expDate, tz, "yyyy-MM-dd") +
              " " +
              callPut +
              " " +
              strike;
          }
        }

        // ---- Trade validations (log issues, but still output row) ----
        if (!ticker)
          mappingIssuesAdd(
            ctx,
            "ERROR",
            importRowNum,
            "Symbol",
            symbolRaw,
            "Missing Ticker (could not parse from Symbol).",
          );
        if (!qty || isNaN(qty))
          mappingIssuesAdd(
            ctx,
            "ERROR",
            importRowNum,
            "Quantity",
            String(qtyRaw || ""),
            "Missing/invalid Quantity.",
          );
        if (
          entryPrice === "" ||
          entryPrice === null ||
          typeof entryPrice === "undefined" ||
          isNaN(entryPrice)
        ) {
          mappingIssuesAdd(
            ctx,
            "ERROR",
            importRowNum,
            "Price",
            String(priceRaw || ""),
            "Missing/invalid Entry Price.",
          );
        }
        const act = String(mapped[col(mappingHeaderMap, "Action")] || "");
        if (!act)
          mappingIssuesAdd(
            ctx,
            "ERROR",
            importRowNum,
            "Side/Pos Effect",
            sideRaw + " / " + posEffectRaw,
            "Missing trade Action (could not build Buy/Sell to Open/Close).",
          );
        const looksOption =
          !isSyntheticExerciseStockLeg &&
          !!(
            expRaw ||
            strikeRaw ||
            callPut ||
            spreadRaw === "SINGLE" ||
            spreadRaw === "VERTICAL" ||
            spreadRaw === "IRON CONDOR" ||
            spreadRaw === "BUTTERFLY"
          );
        if (looksOption) {
          if (!expDate)
            mappingIssuesAdd(
              ctx,
              "ERROR",
              importRowNum,
              "Exp",
              String(expRaw || ""),
              "Option trade missing/invalid Exp.",
            );
          if (
            strike === "" ||
            strike === null ||
            typeof strike === "undefined" ||
            isNaN(strike)
          )
            mappingIssuesAdd(
              ctx,
              "ERROR",
              importRowNum,
              "Strike",
              String(strikeRaw || ""),
              "Option trade missing/invalid Strike.",
            );
          if (!callPut)
            mappingIssuesAdd(
              ctx,
              "ERROR",
              importRowNum,
              "Call/Put",
              symbolRaw,
              "Option trade missing Call/Put.",
            );
        }
      } else if (importAction === "Split") {
        nonTradeRowCount++;

        // ---- Split adjustment row ----
        // WHY:
        // 1) Keep the high-level tag in Corporate Actions for filtering/auditing.
        // 2) Preserve the EXACT raw split text in Notes so Phase 3 can parse
        //    ratios like FORWARD SPLIT 2:1 or REVERSE SPLIT 1:4 later.
        // 3) Do not touch normal trade logic.

        // Best-effort: use the upstream quantity delta if Schwab Import provided it.
        // If qtyRaw is blank or non-numeric, leave it blank here and let Phase 3
        // compute the delta from runningQty + ratio in Notes instead.
        const rawQty =
          qtyRaw === null || typeof qtyRaw === "undefined" ? "" : qtyRaw;
        const parsedQty = parseNumber(rawQty);
        const splitQty = parsedQty === "" || isNaN(parsedQty) ? "" : parsedQty;

        mapped[col(mappingHeaderMap, "Action")] = importAction; // Split

        mapped[col(mappingHeaderMap, "Quantity")] = splitQty; // raw upstream signed delta if present
        mapped[col(mappingHeaderMap, "Signed Quantity")] = splitQty; // same, already signed
        mapped[col(mappingHeaderMap, "Total Cost")] = ""; // no cash value for split rows

        // Preserve the exact descriptive split text for downstream parsing.
        // Description can be any type here (string, formula, blank) so coerce to
        // a trimmed string explicitly.
        const splitDesc = String(desc || "").trim();
        mapped[col(mappingHeaderMap, "Notes")] = splitDesc;

        // Strategy Type and Trade Type are already stamped upstream STOCK.
        // Corporate Actions is already stamped upstream Reverse Split / Stock Split.
        // No cash flow logic, no option fields, no PL for split rows.
      } else {
        nonTradeRowCount++;

        // ---- Non-trade row ----
        // Keep Action for visibility (EFN/CRC/JRN/DOI/etc)
        mapped[col(mappingHeaderMap, "Action")] = importAction;

        // Non-trade actions: Total Cost comes from Schwab Import Amount (cashflow sign matters)
        mapped[col(mappingHeaderMap, "Total Cost")] = importAmount;

        // ===== SYMBOL CHANGE SPECIAL HANDLING =====
        // WHY:
        // 1) Phase 2 must be the ONLY place that decides what a symbol-change row means.
        // 2) If TO is CUSIP-like, keep Ticker anchored to the FROM side so the rename
        //    does NOT happen early.
        // 3) If TO is a real ticker symbol, flip Ticker to the TO side on the true rename row.
        if (
          String(importAction || "")
            .trim()
            .toUpperCase() === "JRN"
        ) {
          const symbolChangeInfo = parseSymbolChangeDescriptionV3(
            desc,
            cusipMapV3,
          );

          if (symbolChangeInfo) {
            const fromRaw = String(symbolChangeInfo.fromRaw || "")
              .trim()
              .toUpperCase();
            const toRaw = String(symbolChangeInfo.toRaw || "")
              .trim()
              .toUpperCase();
            const fromResolved = String(symbolChangeInfo.fromResolved || "")
              .trim()
              .toUpperCase();
            const toResolved = String(symbolChangeInfo.toResolved || "")
              .trim()
              .toUpperCase();

            const toRawLooksLikeCusip =
              /^[A-Z0-9]{9}$/.test(toRaw) && /\d/.test(toRaw);

            let symbolChangeTicker = "";
            if (toRawLooksLikeCusip) {
              symbolChangeTicker = fromResolved || fromRaw || ticker || "";
            } else {
              symbolChangeTicker =
                toResolved || toRaw || fromResolved || fromRaw || ticker || "";
            }

            const parsedRenameQty = parseNumber(qtyRaw);
            const renameQty =
              parsedRenameQty === "" || isNaN(parsedRenameQty)
                ? ""
                : Math.abs(parsedRenameQty);

            mapped[col(mappingHeaderMap, "Ticker")] = symbolChangeTicker;
            mapped[col(mappingHeaderMap, "Corporate Actions")] =
              "Symbol Change";
            mapped[col(mappingHeaderMap, "Action")] = "SYMBOL CHANGE";
            mapped[col(mappingHeaderMap, "Notes")] =
              "FROM=" +
              fromRaw +
              " | TO=" +
              toRaw +
              " | FROM_RESOLVED=" +
              fromResolved +
              " | TO_RESOLVED=" +
              toResolved +
              " | RAW=" +
              desc;
            mapped[col(mappingHeaderMap, "Quantity")] = renameQty;
            mapped[col(mappingHeaderMap, "Signed Quantity")] = renameQty;
            mapped[col(mappingHeaderMap, "Total Cost")] = "";
            mapped[col(mappingHeaderMap, "Strategy Type")] = "LONG STOCK";
          }
        }
        // ===== END SYMBOL CHANGE SPECIAL HANDLING =====

        // If it's a cash-ish row but Amount is blank, that’s worth seeing quickly.

        // If it's a cash-ish row but Amount is blank, that’s worth seeing quickly.
        if (
          importAction &&
          (importAction === "EFN" || importAction === "JRN") &&
          (importAmount === "" || importAmount === null)
        ) {
          mappingIssuesAdd(
            ctx,
            "WARN",
            importRowNum,
            "Amount",
            String(amountRaw || ""),
            "Non-trade cash-like row has blank Amount (Total Cost). Cash Flow Direction may be blank.",
          );
        }

        // Cash Flow Direction + Transfer Type:
        // - Direction = Amount sign (via Total Cost)
        // - Transfer Type = Cash Map lookup (fallback to Account Actions tag)
        //
        // IMPORTANT: RAD rows are "removal/adjustment" style records (corporate/option lifecycle)
        // and very commonly have blank Amount.
        // They are NOT useful for cash in/out tracking, and they create noisy WARN spam.
        // So: skip cash flow logic entirely for RAD.
        if (importAction === "RAD") {
          // Intentionally leave Cash Flow Direction + Transfer Type blank.
        } else {
          applyCashFlowFromMapV3(
            mapped,
            mappingHeaderMap,
            cashFlowMap,
            ctx,
            importRowNum,
          );
        }
      }

      // =========================
      // Strategy Type (single-row baseline)
      // Note: multi-leg spreads get final consistent labeling in the post-processor.
      // =========================
      const callPutFinal = extractCallPut(symbolRaw, desc);
      const actionFinal = String(mapped[col(mappingHeaderMap, "Action")] || "");

      // Pass importAmount as the "net hint" (not leg Total Cost)
      mapped[col(mappingHeaderMap, "Strategy Type")] = normalizeStrategyType(
        spreadRaw,
        actionFinal,
        callPutFinal,
        importAmount,
      );

      // Save output row + metadata
      outItems.push({
        row: mapped,
        importRowNum: importRowNum,
        spreadRaw: spreadRaw,
      });
    }

    // Group-based Strategy Type normalization for multi-leg spreads
    // (VERTICAL / IRON CONDOR / BUTTERFLY) so all legs get the same label.
    postProcessStrategyTypeBySpreadGroups(outItems, mappingHeaderMap);

    // NEW: Position-tracker pass — re-labels Strategy Type on CLOSE rows (including
    // SINGLE-spread closes and mixed VERTICAL groups) to match the spread the position
    // was originally OPENED as, using FIFO quantity accounting per strike.
    // Must run AFTER postProcessStrategyTypeBySpreadGroups so OPEN rows already have
    // their best labels before the ledger is populated.
    postProcessStrategyTypeByPositionTrackerV3(outItems, mappingHeaderMap);

    // Copy Net Amount to all legs in a spread group if Schwab only provided it on one leg
    postProcessNetAmountBySpreadGroups(outItems, mappingHeaderMap);

    // Add WARN issues when a whole spread group is missing Net Amount
    postProcessWarnMissingNetAmountBySpreadGroupsV3(
      outItems,
      mappingHeaderMap,
      ctx,
    );
    // =========================================================================
    // D) SORT BY TRADE TIME STAMP
    //    Authoritative sequencing key for everything downstream.
    //    Stable tie-breaker: original Schwab Import row number when timestamps match.
    // =========================================================================
    sortMappingRowsByTradeTimeStamp(outItems, mappingHeaderMap);

    // Convert to 2D values array for writing
    const outRows = outItems.map(function (it) {
      return it.row;
    });

    // Write to "Schwab Mapping" (rebuild mode keeps formatting because writeMappingRowsV3_ uses clearContent)
    writeMappingRowsV3(mappingSheet, mappingHeaders, outRows);
    rowsWritten = outRows.length;

    // ── Date/Time rendering check ──────────────────────────────────────────────
    // Fires a popup if any data row is missing Trade Date or Trade Time after
    // the write. Schwab Mapping row 2 is the first data row (row 1 is the header).
    checkMissingDateTimeAndAlert(
      mappingSheet,
      2,
      "mapSchwabImportByHeadersV3 → Schwab Mapping",
    );
    // ─────────────────────────────────────────────────────────────────────────

    // =========================
    // Metrics (written as METRIC rows in Schwab Mapping Issues)
    // =========================
    mappingIssuesSetMetric(ctx, "SourceRowsReadExclHeader", rowsRead);
    mappingIssuesSetMetric(ctx, "RowsWrittenExclHeader", rowsWritten);
    mappingIssuesSetMetric(ctx, "DroppedMarkToMarket", droppedMarkToMarket);
    mappingIssuesSetMetric(ctx, "TradeRows", tradeRowCount);
    mappingIssuesSetMetric(ctx, "NonTradeRows", nonTradeRowCount);
    mappingIssuesSetMetric(
      ctx,
      "MissingOrInvalidTimeStampRows",
      missingTimestampCount,
    );
    mappingIssuesSetMetric(ctx, "InvalidAccountRows", invalidAccountCount);

    // ── NEW: Corp action ticker resolution breakdown ──────────────────────────
    mappingIssuesSetMetric(ctx, "CorpTickerFromSymbol", corpTickerFromSymbol);
    mappingIssuesSetMetric(ctx, "CorpTickerFromTilde", corpTickerFromTilde);
    mappingIssuesSetMetric(ctx, "CorpTickerFromMap", corpTickerFromMap);
    mappingIssuesSetMetric(ctx, "CorpTickerFromDesc", corpTickerFromDesc);
    mappingIssuesSetMetric(ctx, "CorpTickerSuspect", corpTickerSuspect);
    mappingIssuesSetMetric(ctx, "CorpTickerUnresolved", corpTickerUnresolved);
    // ── end new metrics ───────────────────────────────────────────────────────

    // Optional: compact breakdown strings (handy, but keep it readable)
    mappingIssuesSetMetric(
      ctx,
      "AccountActionsBreakdown",
      formatCountsForMetric(accountActionCounts),
    );
    mappingIssuesSetMetric(
      ctx,
      "CorporateActionsBreakdown",
      formatCountsForMetric(corpActionCounts),
    );

    // Flush issues + metrics at the end of the run
    mappingIssuesFlush(ctx);

    uiAlertSafe(
      "mapSchwabImportByHeadersV3 finished.\n" +
        "RunId: " +
        ctx.runId +
        "\n" +
        "Rows written: " +
        rowsWritten +
        "\n" +
        "Issues logged: " +
        (ctx.issues ? ctx.issues.length : 0),
    );
  } catch (err) {
    // If something truly unexpected happens, log it as an ERROR issue row.
    mappingIssuesAdd(
      ctx,
      "ERROR",
      "",
      "Exception",
      err && err.message ? err.message : String(err),
      err && err.stack ? err.stack : "",
    );

    // Still write whatever metrics exist so you can see "how far it got".
    mappingIssuesSetMetric(ctx, "SourceRowsReadExclHeader", rowsRead);
    mappingIssuesSetMetric(ctx, "RowsWrittenExclHeader", rowsWritten);
    mappingIssuesSetMetric(ctx, "DroppedMarkToMarket", droppedMarkToMarket);
    mappingIssuesSetMetric(ctx, "TradeRows", tradeRowCount);
    mappingIssuesSetMetric(ctx, "NonTradeRows", nonTradeRowCount);

    mappingIssuesFlush(ctx);
    throw err; // keep normal Apps Script failure behavior (so you see the red error)
  }
}

/** Writes mapping rows to "Schwab Mapping". */
function writeMappingRowsV3(mappingSheet, mappingHeaders, outRows) {
  // Ensure enough rows
  const neededRows = outRows.length + 1; // + header
  const maxRows = mappingSheet.getMaxRows();
  if (neededRows > maxRows) {
    mappingSheet.insertRowsAfter(maxRows, neededRows - maxRows);
  }

  if (MAP_V3_REBUILD_MAPPING_SHEET) {
    // Clear everything below header (safe if you sort / rebuild)
    const lastCol = mappingSheet.getLastColumn();
    const lastRow = Math.max(mappingSheet.getLastRow(), 2);
    mappingSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }

  if (outRows.length) {
    mappingSheet
      .getRange(2, 1, outRows.length, mappingHeaders.length)
      .setValues(outRows);
  }
}

// =========================================================================
// ACCOUNT ACTIONS TAGGING HELPERS
//   Moved to src/phase-2-mapping/MapAccountActions.js
//   getAccountActionsKeywordRulesV3
//   deriveAccountActionTagV3
//   deriveJournalTransferDirectionTag
//   Call sites in mapSchwabImportByHeadersV3() are unchanged.
// =========================================================================

/** Make a compact "A=3 | B=10 | C=1" metric string (sorted by count desc). */
function formatCountsForMetric(countsObj) {
  const keys = Object.keys(countsObj || {});
  if (!keys.length) return "";

  keys.sort(function (a, b) {
    return (countsObj[b] || 0) - (countsObj[a] || 0);
  });

  // Keep it readable; you can increase this if you want the full breakdown.
  const maxItems = 20;
  const parts = [];
  for (let i = 0; i < Math.min(maxItems, keys.length); i++) {
    const k = keys[i];
    parts.push(k + "=" + countsObj[k]);
  }
  return parts.join(" | ");
}

// Missing-Net-Amount WARN moved to src/phase-2-mapping/MapNetAmount.js
// postProcessWarnMissingNetAmountBySpreadGroupsV3

// =====================================================
// Core helper functions (mostly unchanged from your prior version)
// =====================================================

/** Checks for duplicate Headers */
function assertNoDuplicateHeaders(headers, where, ctx) {
  // ctx is optional — pass it from mapSchwabImportByHeadersV3 so a flush
  // happens before the throw. Called from assertNoDuplicateHeaders(headers, sheet, ctx).
  const seen = {};
  const dups = [];
  for (let i = 0; i < headers.length; i++) {
    const key = String(headers[i]).trim().toLowerCase();
    if (!key) continue;
    if (seen[key]) dups.push(String(headers[i]));
    seen[key] = true;
  }
  if (dups.length) {
    const msg = `Duplicate headers in ${where}: ${dups.join(", ")}. Run aborted — fix the header row and re-run.`;
    if (ctx) {
      mappingIssuesAdd(ctx, "ERROR", "", "Headers", dups.join(", "), msg);
      mappingIssuesFlush(ctx);
    }
    throw new Error(msg);
  }
}

/** Normalize Time Stamp into a real Date object (authoritative). */
function normalizeImportTimeStamp(timeStampValue, dateValue, timeValue) {
  // If already a Date, use it.
  if (timeStampValue instanceof Date && !isNaN(timeStampValue))
    return timeStampValue;

  // If numeric (Sheets can store date-times as numbers), convert.
  if (typeof timeStampValue === "number") {
    const d = new Date(Math.round((timeStampValue - 25569) * 86400 * 1000));
    if (d instanceof Date && !isNaN(d)) return d;
  }

  // If string, try parse.
  const tsStr = String(timeStampValue || "").trim();
  if (tsStr) {
    const d = new Date(tsStr);
    if (d instanceof Date && !isNaN(d)) return d;
  }

  // Fallback: build from Date + Time columns
  const dOnly = normalizeImportDateOnly(dateValue);
  const tOnly = normalizeImportTimeOnly(timeValue);
  if (dOnly && tOnly) {
    const hhmm = String(tOnly).padStart(4, "0");
    const hh = Number(hhmm.slice(0, 2));
    const mm = Number(hhmm.slice(2, 4));
    const d = new Date(
      dOnly.getFullYear(),
      dOnly.getMonth(),
      dOnly.getDate(),
      hh,
      mm,
      0,
      0,
    );
    return isNaN(d) ? null : d;
  }

  return null;
}

/** Normalize Date-only field to Date (midnight). */
function normalizeImportDateOnly(dateValue) {
  if (dateValue instanceof Date && !isNaN(dateValue)) {
    return new Date(
      dateValue.getFullYear(),
      dateValue.getMonth(),
      dateValue.getDate(),
    );
  }
  const s = String(dateValue || "").trim();
  if (!s) return null;
  const d = new Date(s);
  if (d instanceof Date && !isNaN(d))
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return null;
}

/** Normalize Time-only field to HHmm string (e.g., "0741"). */
function normalizeImportTimeOnly(timeValue) {
  const s = String(timeValue || "").trim();
  if (!s) return "";
  const cleaned = s.replace(/\D/g, "");
  if (!cleaned) return "";
  return cleaned.padStart(4, "0").slice(-4);
}

// =========================================================================
// TRADE FIELD BUILDERS
//   Moved to src/phase-2-mapping/MapTradeFields.js
//   isTradeBySidePosEffect
//   buildTradeAction
//   buildSignedQuantity
//   extractTickerFromSymbol
//   extractCallPut
//   Call sites in mapSchwabImportByHeadersV3() are unchanged.
// =========================================================================
// =========================================================================
// CORP ACTION MAP + SYMBOL-CHANGE HELPERS
//   Moved to src/phase-2-mapping/MapCorpTickerResolve.js
//   buildCusipMapFromSheetV3
//   normalizeRenameIdentifierV3
//   looksLikeCusipIdentifierV3
//   resolveRenameIdentifierToTickerV3
//   parseSymbolChangeDescriptionV3
//   buildCorpActionMapFromSheet
//   extractTickerFromTildePattern
//   lookupTickerFromCorpActionMap
//   extractTickerFromCorpActionDesc
//   Call sites in mapSchwabImportByHeadersV3() are unchanged.
// =========================================================================

/** Normalize Exp field to Date (midnight). */
function normalizeExpiration(expRaw) {
  if (!expRaw) return null;
  if (expRaw instanceof Date && !isNaN(expRaw))
    return new Date(expRaw.getFullYear(), expRaw.getMonth(), expRaw.getDate());

  const s = String(expRaw || "").trim();
  if (!s) return null;

  const d = new Date(s);
  if (d instanceof Date && !isNaN(d))
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return null;
}

// findFirstKeywordTag moved to src/phase-2-mapping/MapCorpActionRules.js

/** Sort rows by "Trade Time Stamp" ascending.
 *  Enhanced tie-breakers so spread legs stack nicely within the same minute bucket.
 */
function sortMappingRowsByTradeTimeStamp(items, mappingHeaderMap) {
  const tsIdx = col(mappingHeaderMap, "Trade Time Stamp");

  // Additional indexes for tie-break sorting (only used when timestamps tie)
  const acctIdx = col(mappingHeaderMap, "Account");
  const tickerIdx = col(mappingHeaderMap, "Ticker");
  const expIdx = col(mappingHeaderMap, "Option Expiration");
  const strikeIdx = col(mappingHeaderMap, "Option Strike");
  const actionIdx = col(mappingHeaderMap, "Action");

  function toValidDate(v) {
    return v instanceof Date && !isNaN(v) ? v : null;
  }

  function toMs(v) {
    const d = toValidDate(v);
    return d ? d.getTime() : null;
  }

  function normStr(v) {
    return String(v || "")
      .trim()
      .toUpperCase();
  }

  function toNumOrNull(v) {
    const n = typeof v === "number" ? v : Number(String(v || "").trim());
    return isNaN(n) ? null : n;
  }

  function actionRank(act) {
    // This is only a tie-breaker helper.
    // Goal: keep paired legs readable (Buy before Sell typically reads cleaner).
    const a = String(act || "");
    if (a.startsWith("Buy")) return 0;
    if (a.startsWith("Sell")) return 1;
    return 2;
  }

  items.sort(function (a, b) {
    const ta = a.row[tsIdx];
    const tb = b.row[tsIdx];

    const da = toValidDate(ta);
    const db = toValidDate(tb);

    // 1) Primary: timestamp (rows with timestamps come first)
    if (da && db) {
      if (da < db) return -1;
      if (da > db) return 1;

      // ===== Same exact timestamp: enhanced tie-breakers =====

      // 2) Account (keeps DT rows together, LT rows together)
      const acctA = normStr(a.row[acctIdx]);
      const acctB = normStr(b.row[acctIdx]);
      if (acctA !== acctB) return acctA < acctB ? -1 : 1;

      // 3) Ticker
      const tkrA = normStr(a.row[tickerIdx]);
      const tkrB = normStr(b.row[tickerIdx]);
      if (tkrA !== tkrB) return tkrA < tkrB ? -1 : 1;

      // 4) Expiration (THIS is the key that will make your SPX example stack as pairs)
      const expA = toMs(a.row[expIdx]);
      const expB = toMs(b.row[expIdx]);
      // Put dated expirations before blank expirations
      if (expA !== null && expB === null) return -1;
      if (expA === null && expB !== null) return 1;
      if (expA !== null && expB !== null && expA !== expB) return expA - expB;

      // 5) Strike (so within an expiration group, 4340 then 4345, etc.)
      const strikeA = toNumOrNull(a.row[strikeIdx]);
      const strikeB = toNumOrNull(b.row[strikeIdx]);
      if (strikeA !== null && strikeB === null) return -1;
      if (strikeA === null && strikeB !== null) return 1;
      if (strikeA !== null && strikeB !== null && strikeA !== strikeB)
        return strikeA - strikeB;

      // 6) Buy/Sell readability
      const actRankA = actionRank(a.row[actionIdx]);
      const actRankB = actionRank(b.row[actionIdx]);
      if (actRankA !== actRankB) return actRankA - actRankB;

      // 7) Stable final tie-breaker: original Schwab Import row order
      return (a.importRowNum || 0) - (b.importRowNum || 0);
    }

    if (da && !db) return -1;
    if (!da && db) return 1;

    // If neither has a timestamp, stabilize by import row order
    return (a.importRowNum || 0) - (b.importRowNum || 0);
  });
}

// =====================================================
// Cash Map support
//   Moved to src/phase-2-mapping/MapCashFlow.js
//   buildCashFlowMapFromSheet
//   applyCashFlowFromMapV3
//   Call sites in mapSchwabImportByHeadersV3() are unchanged.
// =====================================================

// =========================================================================
// STRATEGY TYPE POST-PROCESSORS
//   Moved to src/phase-2-mapping/MapStrategyType.js
//   normalizeStrategyType
//   postProcessStrategyTypeBySpreadGroups
//   postProcessStrategyTypeByPositionTrackerV3
//   Call sites in mapSchwabImportByHeadersV3() are unchanged.
//
// =========================================================================

// =====================================================
// Corporate Actions keyword rules
//   Moved to src/phase-2-mapping/MapCorpActionRules.js
//   findFirstKeywordTag
//   getCorpActionsKeywordRulesV3
//   Call sites in mapSchwabImportByHeadersV3() are unchanged.
// =====================================================
