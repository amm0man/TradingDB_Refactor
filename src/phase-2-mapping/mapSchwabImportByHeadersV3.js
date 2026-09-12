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
 *   - AuditSchwabMappingV3.js         (pre-Phase-3 gate; same Issues sheet)
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
//   Keyword rules + directional journal logic.
//   Used inside the main loop to fill the "Account Actions" column.
// =========================================================================

/**
 * Base keyword rules for Account Actions where the tag is NOT dependent on
 * account direction.
 *
 * Directional transfers are handled in deriveAccountActionTagV3().
 */
function getAccountActionsKeywordRulesV3() {
  return [
    { keyword: "subscription fee", tag: "Subscription Fee" },
    { keyword: "foreign security fee", tag: "Foreign Security Fee" },

    { keyword: "removal of option due to expiration", tag: "Opt Expired" },
    { keyword: "expired", tag: "Option Expired" },

    { keyword: "non-qualified div", tag: "Non-Qualified Dividend" },

    {
      keyword: "transfer of security or option in",
      tag: "Transfer of Security or Option In",
    },
    {
      keyword: "transfer of security or option out",
      tag: "Transfer of Security or Option Out",
    },

    {
      keyword: "removal of option due to exercise",
      tag: "Option Removal - Exercised",
    },
    {
      keyword: "removal of option due to assignment",
      tag: "Option Removal - Assignment",
    },

    { keyword: "reorganization fee", tag: "Reorganization Fee" },
    

    { keyword: "incoming account transfer", tag: "Incoming Account Transfer" },
    { keyword: "outgoing account transfer", tag: "Outgoing Account Transfer" },

    // Funding / ACH-ish patterns (direction comes from Amount sign)
    { keyword: "new account funding", tag: "Initial Account Funding" },
    {
      keyword: "electronic new account funding",
      tag: "Initial Account Funding",
    },
    { keyword: "client requested electronic", tag: "ACH In or Out" },
    { keyword: "malvern nation", tag: "ACH In or Out" },
    { keyword: "electronic funding", tag: "ACH In or Out" },

    { keyword: "miscellaneous journal entry", tag: "Miscellaneous Journal Entry" },
    { keyword: "cash alternatives", tag: "Cash Alternatives Interest" },
    { keyword: "schwab1 int", tag: "Credit Interest" },

    { keyword: "account migration from tda", tag: "Account Migration" },
  ];
}

/**
 * Derive Account Actions tag:
 * - Directional journal transfers based on "...750" / "...937" and frm/to phrasing
 * - Internal transfers ("internal transfer of cash", "third party") based on Account + Amount sign
 * - Otherwise fall back to simple keyword rules
 */
function deriveAccountActionTagV3(
  descLower,
  accountRaw,
  importAmount,
  accountActionRules,
) {
  const d = String(descLower || "")
    .trim()
    .toLowerCase();
  const account = String(accountRaw || "")
    .trim()
    .toUpperCase();

  // 1) Directional journal transfers (best signal if present)
  const journalTag = deriveJournalTransferDirectionTag(d);
  if (journalTag) return journalTag;

  // 2) Internal transfers that don't always include account numbers in description
  if (d.includes("internal transfer of cash") || d.includes("third party")) {
    // If we have a usable signed amount, we can infer direction based on:
    // - Positive amount = cash INTO this row’s account
    // - Negative amount = cash OUT of this row’s account
    const n =
      typeof importAmount === "number" ? importAmount : Number(importAmount);

    if ((account === "DT" || account === "LT") && !isNaN(n) && n !== 0) {
      const other = account === "DT" ? "LT" : "DT";

      // Example:
      // - Account=DT, Amount=+2000 => Transfer from LT to DT Account
      // - Account=DT, Amount=-2000 => Transfer from DT to LT Account
      if (n > 0)
        return "Transfer from " + other + " to " + account + " Account";
      if (n < 0)
        return "Transfer from " + account + " to " + other + " Account";
    }

    // If amount is missing/0, we can’t safely infer direction.
    // Return a neutral tag (you can filter these easily later).
    return "Internal Transfer (Direction Unknown)";
  }

  // 3) Fall back to ordered keyword rules
  return findFirstKeywordTag(d, accountActionRules);
}

/**
 * Infer direction from "JOURNAL FRM ..." / "JOURNAL TO ..." plus account endings:
 * - DT ends with 750
 * - LT ends with 937
 *
 * We want BOTH sides of the transfer to share the SAME label, e.g.:
 * "Transfer from DT to LT Account"
 */
function deriveJournalTransferDirectionTag(descLower) {
  const d = String(descLower || "").toLowerCase();

  // Normalize some common variants
  const hasFrm = d.includes("journal frm") || d.includes("journal from");
  const hasTo = d.includes("journal to");

  const has750 = d.includes("750");
  const has937 = d.includes("937");

  // If it's a journal and we see BOTH account endings, pick direction based on frm/to.
  // - "frm 750" means FROM DT -> TO LT
  // - "to 937" means TO LT -> FROM DT
  if (hasFrm && has750) return "Transfer from DT to LT Account";
  if (hasFrm && has937) return "Transfer from LT to DT Account";
  if (hasTo && has937) return "Transfer from DT to LT Account";
  if (hasTo && has750) return "Transfer from LT to DT Account";

  return "";
}

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

// =====================================================
// Spread group warnings -> Schwab Mapping Issues
// =====================================================

/**
 * WARN when a whole spread group has missing Net Amount on all legs.
 * Uses Schwab Mapping Issues instead of the old Error Log array.
 */
function postProcessWarnMissingNetAmountBySpreadGroupsV3(
  outItems,
  mappingHeaderMap,
  ctx,
) {
  const idxTs = col(mappingHeaderMap, "Trade Time Stamp");
  const idxTicker = col(mappingHeaderMap, "Ticker");
  const idxDesc = col(mappingHeaderMap, "Description");
  const idxExp = col(mappingHeaderMap, "Option Expiration");
  const idxNetAmount = col(mappingHeaderMap, "Net Amount");

  function toMs(d) {
    return d instanceof Date && !isNaN(d) ? d.getTime() : "";
  }

  // Group by Spread + TimeStamp + Ticker + Expiration
  const groups = {};
  for (let i = 0; i < outItems.length; i++) {
    const it = outItems[i];
    const sp = String(it.spreadRaw || "")
      .trim()
      .toUpperCase();
    if (!SPREAD_GROUP_TYPES.includes(sp)) continue;

    const row = it.row;
    const key =
      sp +
      "|" +
      toMs(row[idxTs]) +
      "|" +
      String(row[idxTicker] || "") +
      "|" +
      toMs(row[idxExp]);

    if (!groups[key]) groups[key] = [];
    groups[key].push(it);
  }

  Object.keys(groups).forEach(function (key) {
    const items = groups[key];
    if (!items.length) return;

    // If ANY leg has a Net Amount, we consider it OK.
    let hasAnyNet = false;
    for (let i = 0; i < items.length; i++) {
      const v = items[i].row[idxNetAmount];
      if (typeof v === "number" && !isNaN(v) && v !== 0) {
        hasAnyNet = true;
        break;
      }
      if (String(v || "").trim() !== "") {
        hasAnyNet = true;
        break;
      }
    }
    if (hasAnyNet) return;

    // Log one WARN using first leg as representative
    const first = items[0];
    const row = first.row;
    const spread = String(first.spreadRaw || "")
      .trim()
      .toUpperCase();

    mappingIssuesAdd(
      ctx,
      "WARN",
      first.importRowNum,
      "Net Amount",
      "",
      "Missing Net Amount on all legs for spread group (" +
        spread +
        "). Ticker=" +
        String(row[idxTicker] || "") +
        ", Exp=" +
        String(row[idxExp] || "") +
        ", Desc=" +
        String(row[idxDesc] || ""),
    );
  });
}

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

/** Identify trade rows strictly by Side + Pos Effect. */
function isTradeBySidePosEffect(sideRaw, posEffectRaw) {
  const side = String(sideRaw || "")
    .trim()
    .toUpperCase();
  const pe = String(posEffectRaw || "")
    .trim()
    .toUpperCase();
  if (side !== "BUY" && side !== "SELL") return false;
  if (!pe.includes("OPEN") && !pe.includes("CLOSE")) return false;
  return true;
}

/** Build "Buy to Open" etc. */
function buildTradeAction(sideRaw, posEffectRaw) {
  const side = String(sideRaw || "")
    .trim()
    .toUpperCase();
  const pe = String(posEffectRaw || "")
    .trim()
    .toUpperCase();

  const sideNice = side === "BUY" ? "Buy" : side === "SELL" ? "Sell" : "";
  const peNice = pe.includes("OPEN")
    ? "Open"
    : pe.includes("CLOSE")
      ? "Close"
      : "";
  if (!sideNice || !peNice) return "";
  return sideNice + " to " + peNice;
}

/** Signed quantity: BUY positive, SELL negative. */
function buildSignedQuantity(sideRaw, qtyAbs) {
  const side = String(sideRaw || "")
    .trim()
    .toUpperCase();
  const q = Number(qtyAbs || 0);
  if (!q) return "";
  return side === "SELL" ? -Math.abs(q) : Math.abs(q);
}

/** Extract ticker from Symbol. For options like "MRVL 11/19/2021 70.00 C", ticker = "MRVL". */
function extractTickerFromSymbol(symbolRaw) {
  const s = String(symbolRaw || "").trim();
  if (!s) return "";
  return s.split(/\s+/)[0].trim();
}

/** Extract Call/Put from Symbol or Description. Returns "C" or "P" or "". */
function extractCallPut(symbolRaw, desc) {
  const s = String(symbolRaw || "").trim();
  const m = s.match(/\b([CP])\b\s*$/i);
  if (m) return m[1].toUpperCase();

  const d = String(desc || "").toUpperCase();
  if (d.includes(" CALL ") || d.includes(" CALL")) return "C";
  if (d.includes(" PUT ") || d.includes(" PUT")) return "P";
  return "";
}

// =========================================================================
// CORP ACTION MAP + SYMBOL-CHANGE HELPERS
//   Resolve tickers for TDA-era corporate-action rows and for "Symbol Change
//   from X to Y" journals. Phase 2 is the single place that decides the
//   canonical Ticker for these rows so Phase 3 can keep blocks continuous.
// =========================================================================
// Symbol / ticker change helpers
// WHY
// Schwab emits symbol changes as JRN rows like:
//   Symbol Change from ENCUF to EU
//   Symbol Change from 50545P309 to LUR/CN
//   Symbol Change from LUR/CN to LURAF
// These are NOT cash journals. They are identity-transfer rows for the same stock position.
// Phase 2 must preserve the exact text in Notes and resolve both sides to canonical tickers
// where possible so Phase 3 can transfer the live stock block without breaking continuity.

// Reads CusipMap into a normalized lookup object.
// Accepts headers: CUSIP in col A style, and Symbol/Ticker/Underlying as the mapped value.
function buildCusipMapFromSheetV3(ss) {
  const sh = ss.getSheetByName("CusipMap");
  if (!sh) return {};

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return {};

  const vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = vals[0].map((h) =>
    String(h || "")
      .trim()
      .toLowerCase(),
  );

  const cusipIdx = headers.findIndex((h) => h === "cusip");
  const symIdx = headers.findIndex(
    (h) => h === "symbol" || h === "ticker" || h === "underlying",
  );
  if (cusipIdx === -1 || symIdx === -1) return {};

  const out = {};
  for (let r = 1; r < vals.length; r++) {
    // Force string first so Sheets number-coercion cannot drop leading zeros / letters.
    const rawCusip = normalizeCusip(String(vals[r][cusipIdx] ?? "").trim());
    const rawSym = String(vals[r][symIdx] || "")
      .trim()
      .toUpperCase();
    if (!rawCusip || !rawSym) continue;
    if (!looksLikeCusip(rawCusip)) continue;
    out[rawCusip] = rawSym;
  }
  return out;
}

/**
 * Normalize a rename-side identifier (ticker or CUSIP-like token).
 * Preserves characters such as "/" that appear in real tickers (e.g. LUR/CN).
 * For pure CUSIP key work, use normalizeCusip() from CusipHelpers.js instead.
 */
function normalizeRenameIdentifierV3(v) {
  return String(v || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");
}

/** Thin wrapper – prefer the shared looksLikeCusip() going forward. */
function looksLikeCusipIdentifierV3(v) {
  return looksLikeCusip(v);
}

function resolveRenameIdentifierToTickerV3(rawId, cusipMap) {
  const id = normalizeRenameIdentifierV3(rawId);
  if (!id) return "";

  // If it looks like a CUSIP, try CusipMap first (shared normalizer).
  const cusipKey = normalizeCusip(id);
  if (looksLikeCusip(cusipKey) && cusipMap && cusipMap[cusipKey]) {
    return String(cusipMap[cusipKey]).trim().toUpperCase();
  }

  // Otherwise treat it as a ticker-like identifier and normalize for storage.
  return id;
}

function parseSymbolChangeDescriptionV3(desc, cusipMap) {
  const text = String(desc || "").trim();
  if (!text) return null;

  // Examples:
  // Symbol Change from ENCUF to EU
  // Symbol Change from 50545P309 to LUR/CN
  // Symbol Change from UNG to 912318300
  const m = text.match(/SYMBOL\s+CHANGE\s+FROM\s+(.+?)\s+TO\s+(.+)$/i);
  if (!m) return null;

  const fromRaw = String(m[1] || "")
    .trim()
    .toUpperCase();
  const toRaw = String(m[2] || "")
    .trim()
    .toUpperCase();

  const fromResolved = resolveRenameIdentifierToTickerV3(fromRaw, cusipMap);
  const toResolved = resolveRenameIdentifierToTickerV3(toRaw, cusipMap);

  return {
    fromRaw: fromRaw,
    toRaw: toRaw,
    fromResolved: fromResolved,
    toResolved: toResolved,
    notesText: text,
  };
}

// Loads Corp Action Map sheet into a sorted array of { pattern, ticker } objects.
// Patterns are sorted longest-first so more-specific entries win over shorter ones.
// If the sheet doesn't exist, returns [] — the fallback chain degrades gracefully.
// Sheet columns: Company Name Pattern (col A) | Ticker (col B) | Notes (col C, ignored)
function buildCorpActionMapFromSheet(ss) {
  const sh = ss.getSheetByName("Corp Action Map");
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return []; // header row only, no data

  const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  const entries = [];
  for (const row of data) {
    const pattern = String(row[0] || "")
      .trim()
      .toUpperCase();
    const ticker = String(row[1] || "")
      .trim()
      .toUpperCase();
    if (pattern && ticker) {
      entries.push({ pattern, ticker });
    }
  }
  // Longest pattern first: "JPMORGAN EQUITY PREMIUM INCOME ETF" (34 chars) beats "JPMORGAN" (8 chars)
  // if someone ever adds a shorter prefix entry by mistake.
  entries.sort((a, b) => b.pattern.length - a.pattern.length);
  return entries;
}

// Extracts a ticker symbol from TDA-era tilde-delimited descriptions.
// Examples:
//   "ORDINARY DIVIDEND~JEPI"          → "JEPI"
//   "Ordinary Dividend~JEPI 3.43 US$" → "JEPI"
//   "Ordinary Dividend~BAC 2.10 US$"  → "BAC"
// The regex only matches ≤6 alpha chars after ~ followed by a non-alpha boundary
// (whitespace, digit, or end of string). This naturally rejects company names:
//   "~JPMORGAN EQUITY" → no match because "JPMORGAN" is followed by " " but
//   we'd match "JPMORGA" (6 chars) and then "N" is NOT a valid boundary → no match.
// Returns '' if no match.
// AFTER — reject placeholder tokens used by TDA when no description was available:
function extractTickerFromTildePattern(desc) {
  if (!desc) return "";
  const m = desc.match(/~([A-Za-z]{1,6})(?:\s|$|\d)/);
  if (!m) return "";
  const token = m[1].toUpperCase();
  // TDA used "~NO DESCRIPTION" as a placeholder meaning "not applicable."
  // "NO", "NA", "NONE", "TBD" after a tilde are always placeholders, never tickers.
  const TILDE_REJECTS = new Set(["NO", "NA", "NONE", "TBD", "NULL"]);
  if (TILDE_REJECTS.has(token)) return "";
  return token;
}

// Looks up a ticker from the Corp Action Map for company-name style descriptions.
// Uses case-insensitive substring match. Since entries are sorted longest-first,
// the most specific pattern wins automatically.
// Returns '' if no match (caller falls through to tokenizer or logs WARN).
function lookupTickerFromCorpActionMap(desc, corpActionMap) {
  if (!corpActionMap || !corpActionMap.length || !desc) return "";
  const upper = desc.toUpperCase();
  for (const entry of corpActionMap) {
    if (upper.includes(entry.pattern)) return entry.ticker;
  }
  return "";
}

// Extracts a likely ticker symbol from a Schwab corporate action Description string.
// Used as a fallback when the Symbol field is blank on RAD/DOI corporate event rows.
//
// How Schwab formats these descriptions:
//   "MANDATORY REVERSE SPLIT ACMR RATIO 1:10 (1 NEW FOR 10 OLD)"  → "ACMR"
//   "QUALIFIED DIVIDEND AAPL"                                      → "AAPL"
//   "NON-QUALIFIED DIV NVDA"                                       → "NVDA"
//   "MANDATORY MERGER INTO MMBI"                                   → "MMBI"
//   "FORWARD SPLIT WITH STOCK SPLIT 2 FOR 1 TSLA"                  → "TSLA"
//   "NON-TAXABLE SPIN OFF UHAL"                                    → "UHAL"
//   "TRANSFER OF SECURITY OR OPTION IN XYZ"                       → "XYZ"
//   "REORGANIZED ISSUE CZNC"                                       → "CZNC"
//   "MANDATORY EXCHANGE MMBI"                                      → "MMBI"
//   "PENDING RECEIPT OF NEW SHARES ACMR"                          → "ACMR"
//
// Approach: split on whitespace and delimiters, skip known action-phrase words
// and numeric tokens, return the FIRST short ALL-CAPS alpha token that looks like a ticker.
// Returns '' if no ticker-like token is found (e.g., "CASH ALTERNATIVES INTEREST").
//
// NOTE: Does NOT handle dot/dollar-prefixed index symbols like $SPX.X — those contain
// non-alpha characters and are filtered out by the /^[A-Z]{1,6}$/ check.
// For RAD rows on SPX/NDX index corporate actions, Symbol is typically populated,
// so the primary extractTickerFromSymbol() path handles them correctly.
function extractTickerFromCorpActionDesc(desc) {
  if (!desc) return "";

  // Action-phrase words that appear in Schwab corp action descriptions but are NOT tickers.
  // Keep this list conservative — only add words you are certain are never used as tickers.
  const SKIP = new Set([
    "MANDATORY",
    "REVERSE",
    "SPLIT",
    "STOCK",
    "FORWARD",
    "MERGER",
    "EXCHANGE",
    "TRANSFER",
    "SECURITY",
    "OPTION",
    "SPIN",
    "OFF",
    "LIQUIDATION",
    "PENDING",
    "RECEIPT",
    "QUALIFIED",
    "DIVIDEND",
    "NON",
    "TAXABLE",
    "DIV",
    "INTEREST",
    "CASH",
    "ALTERNATIVES",
    "REORGANIZATION",
    "REORGANIZED",
    "ISSUE",
    "LIEU",
    "FRACTIONAL",
    "SHARES",
    "IN",
    "OF",
    "FOR",
    "INTO",
    "WITH",
    "AND",
    "OR",
    "THE",
    "DUE",
    "TO",
    "RATIO",
    "NEW",
    "OLD",
    "RECORD",
    "DATE",
    "UPON",
    "CORPORATE",
    "ACTION",
    "FREE",
    "BALANCE",
    "ADJUSTMENT",
    "MARGIN",
    "DESCRIPTION",
    "INTEREST",
    "ORDINARY",
    "QUALIFIED",
    "DISTRIBUTION",
    "PARTNERSHIP",
  ]);

  // Split on whitespace and common delimiters (colon, comma, parens, slash, period, dash).
  // The slash split handles "UHAL/B" → ["UHAL", "B"] so we get "UHAL" first.
  const tokens = desc.toUpperCase().split(/[\s,:.()\[\]\/\-]+/);

  for (const token of tokens) {
    if (!token) continue;
    // Accept only 1-6 purely alpha characters — the standard ticker format.
    // This deliberately rejects numeric tokens (ratio numbers like "10", "1"),
    // and dot/dollar-prefixed index symbols like "$SPX" (handled by Symbol-primary path).
    if (!/^[A-Z]{1,6}$/.test(token)) continue;
    if (SKIP.has(token)) continue;
    return token; // first non-skip purely alpha token is the ticker
  }

  return "";
}

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
// AFTER — supports both { keyword, tag } and { re, tag } rules
// Keyword rules use .includes() against descLower (already lowercased).
// Regex rules use .test() against the original or lowercased string as needed.
function findFirstKeywordTag(descLower, rules) {
  if (!descLower) return "";
  for (let i = 0; i < (rules || []).length; i++) {
    const rule = rules[i];
    if (rule.re) {
      // Regex rule: test against descLower (regex has /i flag so casing doesn't matter)
      if (rule.re.test(descLower)) return String(rule.tag || "");
    } else {
      const kw = String(rule.keyword || "").toLowerCase();
      if (!kw) continue;
      if (descLower.includes(kw)) return String(rule.tag || "");
    }
  }
  return "";
}

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
// Cash Map support (keep, but direction is Amount sign)
// =====================================================

/** Build cashFlowMap from "Cash Map" sheet. Keyed by Account Actions. */
function buildCashFlowMapFromSheet(ss) {
  const sh = ss.getSheetByName(SHEET_CASH_MAP);
  if (!sh) throw new Error("Could not find required sheet: " + SHEET_CASH_MAP);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return {};

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const hm = buildHeaderIndexMap(headers);

  // Expect: Account Actions | CashFlowDir | Transfer Type
  requireHeaders(
    hm,
    ["Account Actions", "CashFlowDir", "Transfer Type"],
    SHEET_CASH_MAP,
  );

  const out = {};
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const action = String(r[col(hm, "Account Actions")] || "").trim();
    if (!action) continue;

    const dir = String(r[col(hm, "CashFlowDir")] || "").trim();
    const type = String(r[col(hm, "Transfer Type")] || "").trim();

    out[action] = { dir: dir, type: type };
  }

  return out;
}

/**
 * Apply Cash Flow Direction + Transfer Type:
 * - Direction: derived from Total Cost sign (your rule: direction from Amount sign)
 * - Transfer Type: from Cash Map if available, else fallback to Account Actions tag
 *
 * NOTE: called only on NON-TRADE rows in this script.
 */
function applyCashFlowFromMapV3(
  mappedRow,
  mappingHeaderMap,
  cashFlowMap,
  ctx,
  importRowNum,
) {
  const accountActionTag = String(
    mappedRow[col(mappingHeaderMap, "Account Actions")] || "",
  ).trim();
  if (!accountActionTag) return;

  // Transfer Type: prefer Cash Map
  const meta = cashFlowMap[accountActionTag] || null;
  const type = meta && meta.type ? meta.type : accountActionTag;

  // Direction: sign of Total Cost (which for non-trade rows = import Amount)
  const totalCostRaw = mappedRow[col(mappingHeaderMap, "Total Cost")];
  const n =
    typeof totalCostRaw === "number"
      ? totalCostRaw
      : Number(
          String(totalCostRaw || "")
            .replace(/[$,]/g, "")
            .trim(),
        );

  let dir = "";
  if (!isNaN(n) && n !== 0) dir = n > 0 ? "Inflow" : "Outflow";

  // Fallback: if we cannot compute sign, use Cash Map direction ONLY as a last resort.
  if (!dir && meta && meta.dir) dir = meta.dir;

  if (!dir) {
    mappingIssuesAdd(
      ctx,
      "WARN",
      importRowNum,
      "Cash Flow Direction",
      "",
      "Could not derive direction (Total Cost blank/0/non-numeric). Account Actions=" +
        accountActionTag,
    );
  }

  mappedRow[col(mappingHeaderMap, "Cash Flow Direction")] = dir;
  mappedRow[col(mappingHeaderMap, "Transfer Type")] = type;
}

// =========================================================================
// STRATEGY TYPE POST-PROCESSORS
//   1. normalizeStrategyType               – single-row baseline
//   2. postProcessStrategyTypeBySpreadGroups – force same label on all legs
//   3. postProcessStrategyTypeByPositionTrackerV3 – FIFO close-row inheritance
//   4. postProcessNetAmountBySpreadGroups  – copy Net Amount across legs
// =========================================================================

function normalizeStrategyType(
  spreadRaw,
  action,
  callPut,
  importAmountForSign,
) {
  // NOTE: we receive importAmount (Net Amount hint from Schwab Import) here.
  // For spread legs the FINAL Strategy Type is decided later in postProcessStrategyTypeBySpreadGroups_ anyway.
  // This baseline is only used for SINGLE / STOCK / simple cases.
  const sp = String(spreadRaw || "")
    .trim()
    .toUpperCase();
  const act = String(action || "").trim();
  const cp = String(callPut || "")
    .trim()
    .toUpperCase();
  const tc =
    typeof importAmountForSign === "number"
      ? importAmountForSign
      : Number(importAmountForSign || 0);
  const isOpen = act.endsWith("to Open");
  const isClose = act.endsWith("to Close");
  const isBuy = act.startsWith("Buy ");
  const isSell = act.startsWith("Sell ");

  // Infer OPENING debit/credit sign:
  // - Opening trade: sign is the row's Total Cost
  // - Closing trade: sign is inverted (because it reverses the entry cashflow)
  let entrySign = tc;
  if (isClose && !isNaN(tc)) entrySign = -tc;
  const entryIsCredit = !isNaN(entrySign) && entrySign > 0;
  const entryIsDebit = !isNaN(entrySign) && entrySign < 0;

  // STOCK
  if (sp === "STOCK") {
    if (isOpen && isBuy) return "Long Stock";
    if (isOpen && isSell) return "Short Stock";
    if (isClose && isSell) return "Long Stock";
    if (isClose && isBuy) return "Short Stock";
    return "Stock";
  }

  // SINGLE-LEG OPTION
  if (sp === "SINGLE") {
    if (cp === "C") {
      if (isOpen && isBuy) return "Long Call";
      if (isOpen && isSell) return "Short Call";
      if (isClose && isSell) return "Long Call";
      if (isClose && isBuy) return "Short Call";
      return "Call";
    }
    if (cp === "P") {
      if (isOpen && isBuy) return "Long Put";
      if (isOpen && isSell) return "Short Put";
      if (isClose && isSell) return "Long Put";
      if (isClose && isBuy) return "Short Put";
      return "Put";
    }
    return "Option";
  }

  if (sp === "VERTICAL") {
    // OPEN legs get a directional single-leg label so the position tracker
    // has meaningful context for mixed-ticket VERTICAL groups (e.g., one
    // CLOSE leg + one OPEN leg in the same TOS ticket).
    // postProcessStrategyTypeBySpreadGroups will overwrite this for
    // pure-OPEN or pure-CLOSE VERTICAL groups with the correct PDS/PCS/etc.
    if (isOpen) {
      if (cp === "P") return isBuy ? "Long Put" : "Short Put";
      if (cp === "C") return isBuy ? "Long Call" : "Short Call";
    }
    // CLOSE legs and ambiguous cases: return generic Vertical.
    // postProcessStrategyTypeBySpreadGroups handles the final label.
    return "Vertical";
  }

  // IRON CONDOR -> Long/Short IC
  if (sp === "IRON CONDOR") {
    if (!entryIsCredit && !entryIsDebit) return "IC";
    return entryIsCredit ? "Short IC" : "Long IC";
  }

  // BUTTERFLY -> Long/Short Butterfly
  if (sp === "BUTTERFLY") {
    if (!entryIsCredit && !entryIsDebit) return "Butterfly";
    return entryIsDebit ? "Long Butterfly" : "Short Butterfly";
  }

  // Default: keep original label for debugging
  return spreadRaw || "";
}

function postProcessStrategyTypeBySpreadGroups(outItems, mappingHeaderMap) {
  const idxTs = col(mappingHeaderMap, "Trade Time Stamp");
  const idxTicker = col(mappingHeaderMap, "Ticker");
  const idxAction = col(mappingHeaderMap, "Action");
  const idxExp = col(mappingHeaderMap, "Option Expiration");
  const idxStrike = col(mappingHeaderMap, "Option Strike");
  const idxCallPut = col(mappingHeaderMap, "Call/Put");
  const idxSignedQty = col(mappingHeaderMap, "Signed Quantity");
  const idxStrategy = col(mappingHeaderMap, "Strategy Type");

  function toMs(d) {
    return d instanceof Date && !isNaN(d) ? d.getTime() : "";
  }

  const groups = {};
  for (let i = 0; i < outItems.length; i++) {
    const it = outItems[i];
    const row = it.row;

    const sp = String(it.spreadRaw || "")
      .trim()
      .toUpperCase();
    if (!SPREAD_GROUP_TYPES.includes(sp)) continue;

    const ts = row[idxTs];
    const tkr = String(row[idxTicker] || "");
    const exp = row[idxExp];

    const cp = String(row[idxCallPut] || "")
      .trim()
      .toUpperCase();
    const key =
      sp +
      "|" +
      toMs(ts) +
      "|" +
      tkr +
      "|" +
      toMs(exp) +
      "|" +
      (sp === "BUTTERFLY" ? cp : "");

    if (!groups[key]) groups[key] = [];
    groups[key].push(it);
  }

  Object.keys(groups).forEach(function (key) {
    const items = groups[key];
    if (!items.length) return;

    const sp = String(items[0].spreadRaw || "")
      .trim()
      .toUpperCase();

    const strikeQty = {};
    let callPutGroup = "";

    for (let i = 0; i < items.length; i++) {
      const row = items[i].row;
      const act = String(row[idxAction] || "");
      const isClose = act.endsWith("to Close");

      const strike = row[idxStrike];
      const sq = Number(row[idxSignedQty]);

      const cp = String(row[idxCallPut] || "")
        .trim()
        .toUpperCase();
      if (!callPutGroup && cp) callPutGroup = cp;

      if (strike === "" || strike === null || typeof strike === "undefined")
        continue;
      if (isNaN(Number(strike))) continue;
      if (isNaN(sq) || sq === 0) continue;

      const entrySq = isClose ? -sq : sq;
      const k = String(Number(strike));
      strikeQty[k] = (strikeQty[k] || 0) + entrySq;
    }

    function setAll(label) {
      for (let i = 0; i < items.length; i++) items[i].row[idxStrategy] = label;
    }

    const cp = callPutGroup;

    if (sp === "VERTICAL") {
      const strikes = Object.keys(strikeQty)
        .map(Number)
        .sort((a, b) => a - b);
      if (strikes.length < 2) return;

      let shortStrike = null,
        longStrike = null;
      for (let i = 0; i < strikes.length; i++) {
        const s = strikes[i];
        const q = strikeQty[String(s)] || 0;
        if (q < 0) shortStrike = s;
        if (q > 0) longStrike = s;
      }
      if (shortStrike === null || longStrike === null) return;

      // ── NEW: detect mixed OPEN/CLOSE groups ──────────────────────────────
      // A "mixed group" is a single TOS ticket that simultaneously closes one
      // leg and opens another (e.g., 5/19 9:38: Sell-to-Close 310 + Buy-to-Open 315).
      // For mixed groups we must NOT apply the group label to the OPEN legs,
      // because the open leg belongs to a *different* spread lifecycle than the
      // close leg. Labeling both "Vertical" creates single-leg PDS fragments
      // downstream. Instead: label CLOSE legs with the resolved spread name,
      // and leave OPEN legs untouched (their label comes from the position
      // tracker or the 9:59 group pass).
      const hasMixedDirections =
        items.some((it) => it.row[idxAction].endsWith("to Close")) &&
        items.some((it) => it.row[idxAction].endsWith("to Open"));

      const label =
        cp === "C"
          ? shortStrike < longStrike
            ? "CCS"
            : "CDS"
          : shortStrike < longStrike
            ? "PDS"
            : "PCS";

      if (hasMixedDirections) {
        // Only stamp CLOSE legs; leave OPEN legs alone.
        for (let i = 0; i < items.length; i++) {
          const act = items[i].row[idxAction];
          if (act.endsWith("to Close")) {
            items[i].row[idxStrategy] = label;
          }
          // OPEN legs in a mixed group: intentionally not touched here.
          // The position tracker will inherit whatever label normalizeStrategyType
          // set (Long Put / Short Put / etc.) and the next full-group VERTICAL
          // pass (9:59 ticket) will stamp PDS on those OPEN legs correctly.
        }
      } else {
        // Pure OPEN or pure CLOSE group — original behavior, stamp all legs.
        setAll(label);
      }
      return;
    }

    if (sp === "IRON CONDOR") {
      const callStrikeQty = {};
      const putStrikeQty = {};

      for (let i = 0; i < items.length; i++) {
        const row = items[i].row;

        const act = String(row[idxAction] || "");
        const isClose = act.endsWith("to Close");

        const cp = String(row[idxCallPut] || "")
          .trim()
          .toUpperCase();
        const strike = row[idxStrike];
        const sq = Number(row[idxSignedQty]);

        if (!cp || isNaN(Number(strike)) || isNaN(sq) || sq === 0) continue;
        const entrySq = isClose ? -sq : sq;

        const map =
          cp === "C" ? callStrikeQty : cp === "P" ? putStrikeQty : null;
        if (!map) continue;

        const k = String(Number(strike));
        map[k] = (map[k] || 0) + entrySq;
      }

      function findShortLong(map) {
        const strikes = Object.keys(map)
          .map(Number)
          .sort(function (a, b) {
            return a - b;
          });
        let shortS = null,
          longS = null;
        for (let i = 0; i < strikes.length; i++) {
          const s = strikes[i];
          const q = map[String(s)] || 0;
          if (q < 0) shortS = s;
          if (q > 0) longS = s;
        }
        return { shortStrike: shortS, longStrike: longS };
      }

      const c = findShortLong(callStrikeQty);
      const p = findShortLong(putStrikeQty);
      if (c.shortStrike === null || c.longStrike === null) return;
      if (p.shortStrike === null || p.longStrike === null) return;

      const isShortIC =
        c.longStrike > c.shortStrike && p.longStrike < p.shortStrike;
      setAll(isShortIC ? "Short IC" : "Long IC");
      return;
    }

    if (sp === "BUTTERFLY") {
      const strikes = Object.keys(strikeQty)
        .map(Number)
        .sort(function (a, b) {
          return a - b;
        });
      if (strikes.length < 3) return;

      let bodyStrike = null;
      let bodyQty = 0;
      for (let i = 0; i < strikes.length; i++) {
        const s = strikes[i];
        const q = strikeQty[String(s)] || 0;
        if (Math.abs(q) > Math.abs(bodyQty)) {
          bodyQty = q;
          bodyStrike = s;
        }
      }
      if (bodyStrike === null) return;

      let wingsAllPositive = true;
      let wingsAllNegative = true;
      for (let i = 0; i < strikes.length; i++) {
        const s = strikes[i];
        if (s === bodyStrike) continue;
        const q = strikeQty[String(s)] || 0;
        if (q <= 0) wingsAllPositive = false;
        if (q >= 0) wingsAllNegative = false;
      }

      if (bodyQty < 0 && wingsAllPositive) {
        setAll("Long Butterfly");
        return;
      }
      if (bodyQty > 0 && wingsAllNegative) {
        setAll("Short Butterfly");
        return;
      }
      return;
    }
  });
}

/**
 * postProcessStrategyTypeByPositionTrackerV3
 *
 * Runs AFTER postProcessStrategyTypeBySpreadGroups.
 *
 * Simulates a running FIFO position ledger per {Account, Ticker, Expiration, Strike}
 * in chronological order to re-label Strategy Type on CLOSE rows where the current
 * label is ambiguous (Long Put, Short Put, Long Call, Short Call, Vertical, etc.)
 * but the ledger can identify which named spread the position was originally opened as.
 *
 * OPEN rows are NEVER re-labeled here — postProcessStrategyTypeBySpreadGroups already
 * set the best possible label for opens, and we must trust that to populate the ledger
 * correctly. The only exception: if an OPEN row has a label that is already a named
 * spread (PDS, PCS, etc.), that label is what gets pushed to the ledger and later
 * applied to the matching CLOSE rows.
 *
 * Key design decisions:
 * - FIFO lot consumption: partial closes reduce the front lot's qty before moving on.
 * - Majority-qty wins when a close spans multiple lots with different labels.
 * - Only re-labels CLOSE rows where the ledger gives a MORE SPECIFIC answer than
 *   the current label (i.e., current is generic and ledger has a named spread label).
 * - 'Vertical' on an OPEN row that cannot be resolved by the group post-processor
 *   (e.g., a mixed-group open that is truly a standalone leg) is left as-is and
 *   pushed to the ledger as 'Vertical'. Downstream CLOSE rows inheriting 'Vertical'
 *   from the ledger will NOT be re-labeled because 'Vertical' is also in GENERIC_LABELS,
 *   meaning the re-label guard (currentIsGeneric || bestIsNamedSpread) only fires
 *   when bestLabel is a named spread — which 'Vertical' is not.
 */
function postProcessStrategyTypeByPositionTrackerV3(
  outItems,
  mappingHeaderMap,
) {
  const idxTs = col(mappingHeaderMap, "Trade Time Stamp");
  const idxAcct = col(mappingHeaderMap, "Account");
  const idxTicker = col(mappingHeaderMap, "Ticker");
  const idxExp = col(mappingHeaderMap, "Option Expiration");
  const idxStrike = col(mappingHeaderMap, "Option Strike");
  const idxAction = col(mappingHeaderMap, "Action");
  const idxSignedQty = col(mappingHeaderMap, "Signed Quantity");
  const idxStrategy = col(mappingHeaderMap, "Strategy Type");

  function toMs(d) {
    return d instanceof Date && !isNaN(d) ? d.getTime() : null;
  }

  // ── Sort a working index chronologically for correct ledger simulation ──
  const sortedIndices = outItems
    .map((_, i) => i)
    .sort((a, b) => {
      const ta = toMs(outItems[a].row[idxTs]);
      const tb = toMs(outItems[b].row[idxTs]);
      if (ta !== null && tb !== null && ta !== tb) return ta - tb;
      if (ta !== null && tb === null) return -1;
      if (ta === null && tb !== null) return 1;
      return (outItems[a].importRowNum || 0) - (outItems[b].importRowNum || 0);
    });

  // ledger[acct][ticker][expMs][strikeKey] = [ { label, qty }, ... ]  (FIFO lots)
  const ledger = {};

  function getLedgerBucket(acct, ticker, expMs, strikeKey) {
    if (!ledger[acct]) ledger[acct] = {};
    if (!ledger[acct][ticker]) ledger[acct][ticker] = {};
    if (!ledger[acct][ticker][expMs]) ledger[acct][ticker][expMs] = {};
    if (!ledger[acct][ticker][expMs][strikeKey])
      ledger[acct][ticker][expMs][strikeKey] = [];
    return ledger[acct][ticker][expMs][strikeKey];
  }

  // Collect relabels and apply them after the full simulation to avoid
  // any ordering side-effects within the same timestamp group.
  const relabelMap = {}; // outItems index → new label string

  const LABEL_RANK = {
    PDS: 0,
    PCS: 1,
    CDS: 2,
    CCS: 3,
    "Short IC": 4,
    "Long IC": 5,
    "Long Butterfly": 6,
    "Short Butterfly": 7,
    "Long Call": 8,
    "Short Call": 9,
    "Long Put": 10,
    "Short Put": 11,
    IC: 12,
    Butterfly: 13,
    Vertical: 14,
    Call: 15,
    Put: 16,
    Option: 17,
    "": 18,
  };
  // Generic labels = ambiguous single-leg or unresolved group labels.
  // A CLOSE row with one of these labels is a candidate for re-labeling
  // IF the ledger has a named-spread answer.
  const GENERIC_LABELS = new Set([
    "Long Put",
    "Short Put",
    "Long Call",
    "Short Call",
    "Put",
    "Call",
    "Option",
    "Vertical",
    "IC",
    "Butterfly",
    "",
  ]);

  for (const idx of sortedIndices) {
    const it = outItems[idx];
    const row = it.row;

    const acct = String(row[idxAcct] || "")
      .trim()
      .toUpperCase();
    const ticker = String(row[idxTicker] || "")
      .trim()
      .toUpperCase();
    const expDate = row[idxExp];
    const strike = row[idxStrike];
    const action = String(row[idxAction] || "").trim();
    const signedQtyRaw = row[idxSignedQty];

    if (!acct || !ticker) continue;
    const expMs = toMs(expDate);
    if (expMs === null) continue; // not an option row
    if (strike === "" || strike === null || strike === undefined) continue;
    const strikeKey = String(Number(strike));
    if (isNaN(Number(strikeKey))) continue;

    const signedQty =
      typeof signedQtyRaw === "number"
        ? signedQtyRaw
        : Number(String(signedQtyRaw || "").replace(/,/g, ""));
    if (isNaN(signedQty) || signedQty === 0) continue;

    const isOpen = action.endsWith("to Open");
    const isClose = action.endsWith("to Close");
    if (!isOpen && !isClose) continue;

    const currentLabel = String(row[idxStrategy] || "").trim();
    const bucket = getLedgerBucket(acct, ticker, expMs, strikeKey);

    if (isOpen) {
      // Push a new FIFO lot with whatever label the group post-processor set.
      // 'Vertical' here means the open was an isolated/standalone leg that the
      // group post-processor couldn't resolve — we intentionally preserve that.
      bucket.push({ label: currentLabel, qty: Math.abs(signedQty) });
    } else {
      // CLOSE — consume FIFO lots and determine the best label for this row.
      let remaining = Math.abs(signedQty);
      const labelsConsumed = [];

      while (remaining > 0 && bucket.length > 0) {
        const lot = bucket[0];
        if (lot.qty <= remaining) {
          labelsConsumed.push({ label: lot.label, qty: lot.qty });
          remaining -= lot.qty;
          bucket.shift();
        } else {
          labelsConsumed.push({ label: lot.label, qty: remaining });
          lot.qty -= remaining;
          remaining = 0;
        }
      }
      // If we consumed more than the ledger knew about (missing history),
      // fall back to the current label for the unmatched portion.
      if (remaining > 0 && currentLabel) {
        labelsConsumed.push({ label: currentLabel, qty: remaining });
      }
      if (labelsConsumed.length === 0) continue;

      // Majority-qty wins; named-spread labels win ties over generic ones.
      const tally = {};
      for (const { label, qty } of labelsConsumed) {
        tally[label] = (tally[label] || 0) + qty;
      }
      const bestLabel = Object.keys(tally).sort((a, b) => {
        const qDiff = tally[b] - tally[a];
        if (qDiff !== 0) return qDiff;
        return (
          (LABEL_RANK[a] !== undefined ? LABEL_RANK[a] : 99) -
          (LABEL_RANK[b] !== undefined ? LABEL_RANK[b] : 99)
        );
      })[0];

      // Re-label guard — corrected:
      const currentIsGeneric = GENERIC_LABELS.has(currentLabel);
      const bestIsGeneric = GENERIC_LABELS.has(bestLabel) || bestLabel === "";

      // Only block the re-label in one case:
      // current is a named spread AND best is generic — ledger can't improve things.
      // In all other cases (generic→named, generic→generic, named→named-different),
      // the ledger's answer is at least as good or better than what's there.
      const shouldRelabel =
        bestLabel !== currentLabel && !(!currentIsGeneric && bestIsGeneric); // block only: named current + generic best

      if (shouldRelabel) {
        relabelMap[idx] = bestLabel;
      }
    }
  }

  // Apply all collected relabels
  for (const [idxStr, newLabel] of Object.entries(relabelMap)) {
    outItems[Number(idxStr)].row[idxStrategy] = newLabel;
  }
}

function postProcessNetAmountBySpreadGroups(outItems, mappingHeaderMap) {
  const idxTs = col(mappingHeaderMap, "Trade Time Stamp");
  const idxTicker = col(mappingHeaderMap, "Ticker");
  const idxExp = col(mappingHeaderMap, "Option Expiration");
  const idxNetAmount = col(mappingHeaderMap, "Net Amount");

  function toMs(d) {
    return d instanceof Date && !isNaN(d) ? d.getTime() : "";
  }

  const groups = {};
  for (let i = 0; i < outItems.length; i++) {
    const it = outItems[i];
    const sp = String(it.spreadRaw || "")
      .trim()
      .toUpperCase();
    if (!SPREAD_GROUP_TYPES.includes(sp)) continue;

    const row = it.row;
    const key =
      sp +
      "|" +
      toMs(row[idxTs]) +
      "|" +
      String(row[idxTicker] || "") +
      "|" +
      toMs(row[idxExp]);

    if (!groups[key]) groups[key] = [];
    groups[key].push(it);
  }

  Object.keys(groups).forEach(function (key) {
    const items = groups[key];
    let net = "";

    for (let i = 0; i < items.length; i++) {
      const v = items[i].row[idxNetAmount];
      if (typeof v === "number" && !isNaN(v) && v !== 0) {
        net = v;
        break;
      }
      if (String(v || "").trim() !== "") {
        net = v;
        break;
      }
    }

    if (net === "" || net === null || typeof net === "undefined") return;

    for (let i = 0; i < items.length; i++) items[i].row[idxNetAmount] = net;
  });
}

// =====================================================
// Corporate Actions helpers — CONSOLIDATED single source of truth (Phase 2)
// Phase 3 (copyMappingToImportByHeaders) trusts whatever this sets.
// No re-evaluation happens in Phase 3 — this is the only place that tags corp actions.
//
// ORDER MATTERS: more-specific patterns MUST come before less-specific ones.
// Rules support both { keyword, tag } (substring match) and { re, tag } (regex match).
// =====================================================
function getCorpActionsKeywordRulesV3() {
  return [
    // ── Dividends / Interest ──────────────────────────────────────────────────
    { keyword: "qualified dividend", tag: "Dividend" },
    { keyword: "ordinary dividend", tag: "Dividend" }, // TDA "ORDINARY DIVIDEND~JEPI"
    { keyword: "non-qualified dividend", tag: "Dividend" }, // long form variant
    { keyword: "special dividend", tag: "Dividend" }, // one-time special divs
    { keyword: "return of capital", tag: "Return of Capital" }, // REITs / MLPs
    { keyword: "monthly dividend", tag: "Dividend" }, // some TDA formats
    { keyword: "cash dividend", tag: "Cash Dividend" },
    { keyword: "reinvest dividend", tag: "DRIP" },
    { keyword: "reinvest shares", tag: "DRIP" },
    // +++ NEW: matches DRIP rows emitted by buildUnifiedImportV3
    // Description format: "DRIP BUY +0.0175 XOM UPON REINVESTMENT"
    { keyword: "upon reinvestment", tag: "DRIP" },
    // +++ END NEW
    { keyword: "foreign tax paid", tag: "Foreign Tax Paid" },
    { keyword: "foreign tax withheld", tag: "Foreign Tax Withheld" },
    { keyword: "bond interest", tag: "Bond Interest" },
    { keyword: "cash alternatives interest", tag: "Cash Interest" },
    { keyword: "partnership distribution", tag: "Partnership Distribution" },
    { keyword: "free balance interest", tag: "Interest Adjustment" },
    {
      keyword: "margin interest adjustment",
      tag: "Margin Interest Adjustment",
    },

    // ── Splits — more-specific FIRST ─────────────────────────────────────────
    { keyword: "mandatory reverse split", tag: "Mandatory Reverse Split" },
    { keyword: "reverse split", tag: "Reverse Split" },
    { keyword: "forward split with stock split", tag: "Forward Split" },
    { keyword: "stock split", tag: "Stock Split" },
    { keyword: "split", tag: "Stock Split" }, // catch-all — after all specific split variants

    // ── Mergers / Reorganizations ─────────────────────────────────────────────
    { keyword: "mandatory merger", tag: "Mandatory Merger" },
    { keyword: "stock merger", tag: "Stock Merger" },
    { keyword: "merger", tag: "Stock Merger" },
    { keyword: "reorganized issue", tag: "Reorganization" },
    { keyword: "mandatory exchange", tag: "Mandatory Exchange" },

    // ── Spin-offs / Liquidations ──────────────────────────────────────────────
    { keyword: "non-taxable spin off", tag: "Spin-off/Liquidation" },

    // ── Transfers ────────────────────────────────────────────────────────────
    // More-specific in/out variants before the generic one
    {
      keyword: "transfer of security or option in",
      tag: "Transfer of Security or Option In",
    },
    {
      keyword: "transfer of security or option out",
      tag: "Transfer of Security or Option Out",
    },
    { keyword: "transfer of security or option", tag: "Transfer of Security" },

    // ── Cash / Miscellaneous ─────────────────────────────────────────────────
    { keyword: "cash in lieu of fractional shares", tag: "Cash In Lieu" },

    // ── Pending / Received shares ─────────────────────────────────────────────
    { keyword: "pending receipt of new s", tag: "Pending Corp Action" },

    // ── Broad catch-all: paired credit-side rows for corporate restructurings ──
    // Schwab emits debit/credit pairs. The debit side is caught by 'pending receipt' above.
    // The credit side uses an abbreviated format: COMPANYNAME SINGLEACTIONLETTER QTY NEWTICKER
    // e.g.  "ACME F4 100 XYZ"
    // Placed LAST — only fires after every specific pattern has been tested.
    {
      re: /[A-Z]{4,}\s+[A-Z]\d+\s+[\d.]+\s+[A-Z]{1,6}/i,
      tag: "Corp Action Received Shares",
    },
  ];
}
