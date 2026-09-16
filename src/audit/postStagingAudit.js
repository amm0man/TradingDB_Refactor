// ============================================================================
// auditPipelineIntegrity
//
// PURPOSE: Read-only scan of Staging sheet. Detects data quality problems
//          that slip past the normal pipeline (orphaned closes, missing stock
//          legs, duplicate timestamps, malformed option fields, etc.).
//          Writes findings to a dedicated "Audit Results" sheet.
//          NEVER modifies Helper, Import, or Staging.
//
// CALL:    Run standalone from the Apps Script menu, OR add as the final
//          step in refreshAllScripts() after populateStagingWithBlockLogicV3().
//
// OUTPUT:  "Audit Results" sheet — color-coded by severity:
//          🔴 ERROR  = data integrity failure, will corrupt P&L or block logic
//          🟡 WARN   = likely problem, needs human review
//          🟢 INFO   = informational only, no action required
//          ⚪ SUMMARY = per-account pipeline statistics
// ============================================================================
function auditPipelineIntegrity() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stagingSheet = ss.getSheetByName("Staging");
  if (!stagingSheet) {
    SpreadsheetApp.getUi().alert(
      "❌ Staging sheet not found — run refreshAllScripts() first.",
    );
    return;
  }

  const allData = stagingSheet.getDataRange().getValues();
  if (allData.length < 4) {
    SpreadsheetApp.getUi().alert(
      "❌ Staging has no data rows (expected data from row 4 down).",
    );
    return;
  }

  const rawHeaders = allData[0];
  const col = {};
  rawHeaders.forEach((h, i) => {
    const norm = h.toString().trim().toLowerCase();
    if (norm) col[norm] = i;
  });

  const data = allData.slice(3);
  const DATA_START_ROW = 4;
  const tz = ss.getSpreadsheetTimeZone(); // FIX 2: capture once, reuse everywhere

  const OPEN_ACTIONS = ["BUY TO OPEN", "SELL TO OPEN"];
  const CLOSE_ACTIONS = ["BUY TO CLOSE", "SELL TO CLOSE"];
  const TRADE_ACTIONS = [...OPEN_ACTIONS, ...CLOSE_ACTIONS];

  const findings = [];

  function flag(
    check,
    severity,
    dataIdx,
    account,
    ticker,
    tradeDate,
    action,
    field,
    value,
    detail,
  ) {
    const tdDisplay =
      tradeDate instanceof Date
        ? Utilities.formatDate(tradeDate, tz, "M/d/yyyy") // FIX 2: uses tz
        : (tradeDate || "").toString();
    findings.push([
      check,
      severity,
      dataIdx + DATA_START_ROW,
      (account || "").toString().toUpperCase(),
      (ticker || "").toString().toUpperCase(),
      tdDisplay,
      (action || "").toString().toUpperCase(),
      field,
      value.toString(),
      detail,
    ]);
  }

  function cv(row, colName) {
    const i = col[colName];
    return i !== undefined ? row[i] : "";
  }
  function cvStr(row, colName) {
    return (cv(row, colName) || "").toString().trim();
  }
  function cvUpper(row, colName) {
    return cvStr(row, colName).toUpperCase();
  }
  function cvNum(row, colName) {
    return Number(cv(row, colName)) || 0;
  }
  function cvDate(row, colName) {
    const v = cv(row, colName);
    return v instanceof Date && !isNaN(v.getTime()) ? v : null;
  }

  // ── Sets built once — used by multiple checks ─────────────────────────────
  // (moved outside try so catch block can still reference findings safely)
  const posHasStart = new Set();
  const posHasClose = new Set();
  const posFirstRow = {};
  // Account|Ticker books that already have a STOCK (or -ST) Block Start.
  // Used by CHECK 4 so add-on lots are not treated as missing opens.
  const stockBookHasStart = new Set();

  try {
    // ════════════════════════════════════════════════════════════════════════
    // CHECK 1 — Negative Running Position Quantity
    // ════════════════════════════════════════════════════════════════════════
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const runQty = cvNum(row, "running position quantity");
      const ticker = cvStr(row, "ticker");
      if (runQty < 0 && ticker) {
        flag(
          "NEG_RUNNING_QTY",
          "ERROR",
          i,
          cv(row, "account"),
          ticker,
          cv(row, "trade date"),
          cv(row, "action"),
          "Running Position Quantity",
          runQty,
          "Quantity went negative (" +
            runQty +
            "). A closing row arrived with no " +
            "matching opener. Check for a missing BUY/SELL TO OPEN in this dataset, " +
            "or a wrong Account assignment on the opening leg.",
        );
      }
    }

    // ════════════════════════════════════════════════════════════════════════
    // CHECK 2 — Block Close Flag = 1 but Running Position Quantity ≠ 0
    // ════════════════════════════════════════════════════════════════════════
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const blkClose = cvNum(row, "block close flag/p&l");
      const runQty = cvNum(row, "running position quantity");
      const ticker = cvStr(row, "ticker");
      if (blkClose === 1 && runQty !== 0 && ticker) {
        flag(
          "CLOSE_QTY_NONZERO",
          "ERROR",
          i,
          cv(row, "account"),
          ticker,
          cv(row, "trade date"),
          cv(row, "action"),
          "Running Position Quantity",
          runQty,
          "Block Close Flag = 1 but Running Position Quantity = " +
            runQty +
            " (expected 0). One or more legs of this position may be missing. " +
            "P&L on this block is likely incorrect.",
        );
      }
    }

    // ════════════════════════════════════════════════════════════════════════
    // CHECK 3 — Duplicate trade executions
    // WHY: Account|Timestamp|Ticker|Action is too thin. TOS/Schwab often
    //      writes two valid fills in the same second (partial fills, two
    //      lots, or two spread legs). Those must not be ERROR.
    //      A real cloned row matches Quantity + Price + option identity too.
    //      Even then, two identical-size fills can share a second, so leftover
    //      matches are INFO (look-at-this), not data-integrity failures.
    //
    // 2026-09-16: WRITE_DUPLICATE_TIMESTAMP_INFO is false so Audit Results
    // is not flooded with ~1796 valid same-second fills. The scan is kept
    // so this can be turned back on without rewriting the key. Do not
    // promote these back to ERROR.
    // ════════════════════════════════════════════════════════════════════════
    const WRITE_DUPLICATE_TIMESTAMP_INFO = false;
    const tsMap = {};
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const acct = cvUpper(row, "account");
      const ticker = cvStr(row, "ticker");
      const action = cvUpper(row, "action");
      if (!ticker || !TRADE_ACTIONS.includes(action)) continue;
      const ts = cv(row, "trade time stamp");
      const tsStr = ts instanceof Date ? ts.toISOString() : ts.toString();
      if (!tsStr) continue;
      const qty = cvNum(row, "quantity");
      const price = cvNum(row, "price");
      const contract = cvUpper(row, "option contract");
      const strike = cvStr(row, "option strike");
      const cp = cvUpper(row, "call/put");
      const key = `${acct}|${tsStr}|${ticker}|${action}|${qty}|${price}|${contract}|${strike}|${cp}`;
      if (!tsMap[key]) tsMap[key] = [];
      tsMap[key].push(i);
    }
    if (WRITE_DUPLICATE_TIMESTAMP_INFO) {
      for (const [key, rows] of Object.entries(tsMap)) {
        if (rows.length > 1) {
          rows.forEach((i) => {
            const row = data[i];
            const parts = key.split("|");
            flag(
              "DUPLICATE_TIMESTAMP",
              "INFO",
              i,
              cv(row, "account"),
              cv(row, "ticker"),
              cv(row, "trade date"),
              cv(row, "action"),
              "Trade Time Stamp",
              parts[1],
              "Same-second executions with the same qty/price/contract (" +
                rows.length +
                " rows). Valid broker fills unless Schwab Mapping " +
                "is also doubled. Staging rows: " +
                rows.map((r) => r + DATA_START_ROW).join(", "),
            );
          });
        }
      }
    }

    // ════════════════════════════════════════════════════════════════════════
    // CHECK 4 — Position ID exists but no Block Start Flag in the block
    // WHY: Block Start Flag means “this ticker book went from 0 to open.”
    //      Assignment / extra stock buys get their own Position ID (lot
    //      label, often parentTgId-ST) while Running Position Quantity stays
    //      on the ticker book. Those add-on IDs correctly have Start = 0.
    //      Only flag when this Account+Ticker stock book never started.
    // ════════════════════════════════════════════════════════════════════════
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const posId = cvStr(row, "position id");
      const ticker = cvStr(row, "ticker");
      if (!posId || !ticker) continue;
      if (cvNum(row, "block start flag") === 1) {
        posHasStart.add(posId);
        const tt = cvUpper(row, "trade type");
        if (tt === "STOCK" || posId.toUpperCase().endsWith("-ST")) {
          stockBookHasStart.add(cvUpper(row, "account") + "|" + ticker);
        }
      }
      if (cvNum(row, "block close flag/p&l") === 1) posHasClose.add(posId);
      if (posFirstRow[posId] === undefined) posFirstRow[posId] = i;
    }

    const reportedNoStart = new Set();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const posId = cvStr(row, "position id");
      const ticker = cvStr(row, "ticker");
      if (!posId || !ticker || reportedNoStart.has(posId)) continue;
      if (posHasStart.has(posId)) continue;

      const first = data[posFirstRow[posId]];
      const firstType = cvUpper(first, "trade type");
      const firstAcct = cvUpper(first, "account");
      const firstTicker = cvStr(first, "ticker");
      const bookKey = firstAcct + "|" + firstTicker;
      const isStockLot =
        firstType === "STOCK" || posId.toUpperCase().endsWith("-ST");
      const runQty = cvNum(first, "running position quantity");
      const qty = cvNum(first, "quantity");
      const isAddOnLot =
        isStockLot && (stockBookHasStart.has(bookKey) || runQty > qty);

      if (isAddOnLot) {
        reportedNoStart.add(posId);
        continue;
      }

      reportedNoStart.add(posId);
      flag(
        "POSID_NO_BLOCK_START",
        "ERROR",
        posFirstRow[posId],
        cv(first, "account"),
        ticker,
        cv(first, "trade date"),
        cv(first, "action"),
        "Position ID",
        posId,
        "Position ID found on rows but no Block Start Flag = 1 exists. " +
          "The opening trade for this position is missing from the dataset. " +
          "P&L, Trade Duration, and Running Quantity for this block are unreliable.",
      );
    }

    // ════════════════════════════════════════════════════════════════════════
    // CHECK 5 — Block Start Flag = 1 but Position ID is blank
    // ════════════════════════════════════════════════════════════════════════
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const ticker = cvStr(row, "ticker");
      if (!ticker) continue;
      if (cvNum(row, "block start flag") === 1 && !cvStr(row, "position id")) {
        flag(
          "BLKSTART_NO_POSID",
          "WARN",
          i,
          cv(row, "account"),
          ticker,
          cv(row, "trade date"),
          cv(row, "action"),
          "Position ID",
          "(blank)",
          "Block Start Flag = 1 but Position ID was not generated. " +
            "Likely cause: Trade Type is blank on this row. " +
            "Check Strategy Type and Trade Type in Helper for this row.",
        );
      }
    }

    // ════════════════════════════════════════════════════════════════════════
    // CHECK 6 — Option rows missing Expiration, Call/Put, or Option Contract
    // ════════════════════════════════════════════════════════════════════════
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const ticker = cvStr(row, "ticker");
      const tradeType = cvUpper(row, "trade type");
      const action = cvUpper(row, "action");
      if (!ticker) continue;
      if (![...TRADE_ACTIONS, "RAD"].includes(action)) continue;
      if (tradeType !== "OPTION" && tradeType !== "SPREAD") continue;

      const strike = cv(row, "option strike");
      const expDate = cv(row, "option expiration");
      const cp = cvUpper(row, "call/put");
      const contract = cvStr(row, "option contract");

      const hasStrike =
        strike !== "" && !isNaN(Number(strike)) && Number(strike) > 0;
      const hasExp =
        expDate instanceof Date ||
        (expDate && expDate.toString().trim() !== "");
      const validCp = cp === "C" || cp === "P";

      if (hasStrike && !hasExp) {
        flag(
          "OPTION_MISSING_EXP",
          "WARN",
          i,
          cv(row, "account"),
          ticker,
          cv(row, "trade date"),
          action,
          "Option Expiration",
          "(blank)",
          "Strike = " +
            strike +
            " but Option Expiration is blank. " +
            "Block grouping key is wrong — this row may merge with a different expiry.",
        );
      }
      if (hasStrike && !validCp) {
        flag(
          "OPTION_MISSING_CP",
          "WARN",
          i,
          cv(row, "account"),
          ticker,
          cv(row, "trade date"),
          action,
          "Call/Put",
          cp || "(blank)",
          "Strike = " +
            strike +
            " but Call/Put is blank or not C/P. " +
            "Block grouping key is wrong.",
        );
      }
      if (hasStrike && hasExp && validCp && !contract) {
        flag(
          "OPTION_MISSING_CONTRACT",
          "WARN",
          i,
          cv(row, "account"),
          ticker,
          cv(row, "trade date"),
          action,
          "Option Contract",
          "(blank)",
          "Option has Strike + Expiration + C/P but Option Contract (OCC format) is blank. " +
            "validateAndCleanImportToHelperV3 may have failed to build the OCC string for this row.",
        );
      }
    }

    // ════════════════════════════════════════════════════════════════════════
    // CHECK 7 — Spread rows missing Spread Group ID
    // ════════════════════════════════════════════════════════════════════════
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const ticker = cvStr(row, "ticker");
      const tradeType = cvUpper(row, "trade type");
      const action = cvUpper(row, "action");
      if (!ticker || !TRADE_ACTIONS.includes(action)) continue;
      if (tradeType === "SPREAD" && !cvStr(row, "spread group id")) {
        flag(
          "SPREAD_MISSING_GROUP_ID",
          "WARN",
          i,
          cv(row, "account"),
          ticker,
          cv(row, "trade date"),
          action,
          "Spread Group ID",
          "(blank)",
          "Trade Type = Spread but Spread Group ID is blank. " +
            "Legs of this spread will be treated as individual option positions. " +
            "Check Strategy Type and Option Expiration on the opening leg.",
        );
      }
    }

    // ════════════════════════════════════════════════════════════════════════
    // CHECK 8 — Trade Status / Block Close Flag consistency
    // ════════════════════════════════════════════════════════════════════════
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const ticker = cvStr(row, "ticker");
      const action = cvUpper(row, "action");
      if (!ticker || ![...TRADE_ACTIONS, "RAD"].includes(action)) continue;

      const status = cvUpper(row, "trade status");
      const blkClose = cvNum(row, "block close flag/p&l");
      const closingDate = cv(row, "closing date");
      const hasClosingDate =
        closingDate instanceof Date ||
        (closingDate && closingDate.toString().trim() !== "");

      if (blkClose === 1 && status && status !== "CLOSED") {
        flag(
          "CLOSE_STATUS_MISMATCH",
          "WARN",
          i,
          cv(row, "account"),
          ticker,
          cv(row, "trade date"),
          action,
          "Trade Status",
          status,
          'Block Close Flag = 1 but Trade Status = "' +
            status +
            '" (expected "Closed"). ' +
            "Dashboard close filters will miss this position.",
        );
      }
      if (status === "CLOSED" && !hasClosingDate) {
        flag(
          "CLOSED_NO_DATE",
          "WARN",
          i,
          cv(row, "account"),
          ticker,
          cv(row, "trade date"),
          action,
          "Closing Date",
          "(blank)",
          "Trade Status = Closed but Closing Date is blank. " +
            "Trade Duration cannot be calculated.",
        );
      }
    }

    // ════════════════════════════════════════════════════════════════════════
    // CHECK 9 — RAD Assignment without a matching same-date stock leg
    // WHY: Put assignment delivers long shares (BUY TO OPEN).
    //      Call assignment delivers short shares (SELL TO OPEN) or sells
    //      a covered long (SELL TO CLOSE). The old check only looked for BTO,
    //      so every short-call assignment WARNed even when the stock row exists.
    //      Cash-settled index tickers never deliver shares — skip those.
    //      Keep CASH_SETTLED_INDEX_TICKERS in sync with
    //      copyMappingToImportByHeaders in Phase3Step1_CopyMapping.js.
    // ════════════════════════════════════════════════════════════════════════
    const CASH_SETTLED_INDEX_TICKERS = {
      SPX: true,
      SPXW: true,
      XSP: true,
      NDX: true,
      NDXP: true,
      RUT: true,
      RUTW: true,
      VIX: true,
    };

    const stockBtoByDateKey = new Set();
    const stockStoByDateKey = new Set();
    const stockStcByDateKey = new Set();
    const exerciseByDateKey = new Set();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const action = cvUpper(row, "action");
      const tradeType = cvUpper(row, "trade type");
      const ticker = cvStr(row, "ticker");
      if (tradeType !== "STOCK" || !ticker) continue;
      const acct = cvUpper(row, "account");
      const td = cv(row, "trade date");
      const tdStr = td instanceof Date ? td.toDateString() : td.toString();
      const key = `${acct}|${ticker}|${tdStr}`;
      if (action === "BUY TO OPEN") stockBtoByDateKey.add(key);
      if (action === "SELL TO OPEN") stockStoByDateKey.add(key);
      if (action === "SELL TO CLOSE") stockStcByDateKey.add(key);
    }

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (cvUpper(row, "action") !== "RAD") continue;
      const ticker = cvStr(row, "ticker");
      if (!ticker) continue;
      const acctAct = cvUpper(row, "account actions");
      if (acctAct.indexOf("EXERCISE") === -1) continue;
      const td = cv(row, "trade date");
      const tdStr = td instanceof Date ? td.toDateString() : td.toString();
      exerciseByDateKey.add(
        cvUpper(row, "account") + "|" + ticker + "|" + tdStr,
      );
    }

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const action = cvUpper(row, "action");
      const ticker = cvStr(row, "ticker");
      if (action !== "RAD" || !ticker) continue;
      if (CASH_SETTLED_INDEX_TICKERS[ticker.toUpperCase()]) continue;
      // Opt Expired shorts have the same C/P + signed-qty shape as
      // assignment. Only warn when Account Actions says this RAD
      // actually assigned or exercised. UNG assignment rows say
      // "Option Removal - Assignment". The 43 leftovers said "Opt Expired".
      const acctAct = cvUpper(row, "account actions");
      if (
        acctAct.indexOf("ASSIGN") === -1 &&
        acctAct.indexOf("EXERCISE") === -1
      ) {
        continue;
      }
      const cp = cvUpper(row, "call/put");
      const signedQty = cvNum(row, "signed quantity");
      const isPutAssign = cp === "P" && signedQty > 0;
      const isCallAssign = cp === "C" && signedQty < 0;
      if (!isPutAssign && !isCallAssign) continue;

      const acct = cvUpper(row, "account");
      const td = cv(row, "trade date");
      const tdStr = td instanceof Date ? td.toDateString() : td.toString();
      const key = `${acct}|${ticker}|${tdStr}`;

      const hasLeg = isPutAssign
        ? stockBtoByDateKey.has(key)
        : stockStoByDateKey.has(key) || stockStcByDateKey.has(key);

      // Spread expiration: short stamped Assignment, long stamped Exercised
      // the same day (LT SPY 402/403 PCS 6/11/2022). Broker nets the shares;
      // no STOCK row is expected.
      if (!hasLeg && exerciseByDateKey.has(key)) continue;

      if (!hasLeg) {
        flag(
          "ASSIGNMENT_NO_STOCK_LEG",
          "WARN",
          i,
          acct,
          ticker,
          td,
          action,
          "Assignment Stock Leg",
          "(missing)",
          cp +
            " assignment at $" +
            cvStr(row, "option strike") +
            " found but no same-date STOCK " +
            (isPutAssign ? "BUY TO OPEN" : "SELL TO OPEN / SELL TO CLOSE") +
            " exists for " +
            ticker +
            " on " +
            tdStr +
            ".",
        );
      }
    }

    // ════════════════════════════════════════════════════════════════════════
    // CHECK 10 — Realized P&L = 0 on a Block Close (non-RAD)
    // ════════════════════════════════════════════════════════════════════════
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const ticker = cvStr(row, "ticker");
      const action = cvUpper(row, "action");
      const blkClose = cvNum(row, "block close flag/p&l");
      const pnl = cvNum(row, "realized p&l");
      if (!ticker || blkClose !== 1 || action === "RAD") continue;
      if (pnl === 0) {
        flag(
          "ZERO_PNL_ON_CLOSE",
          "WARN",
          i,
          cv(row, "account"),
          ticker,
          cv(row, "trade date"),
          action,
          "Realized P&L",
          0,
          "Realized P&L = $0 on a non-RAD block close. " +
            "This may be correct (breakeven trade) or may indicate the entry cost " +
            "accumulator was not populated (missing opening leg or $0 entry price on opener).",
        );
      }
    }

    // ════════════════════════════════════════════════════════════════════════
    // CHECK 11 — Trade Time Stamp is blank or not a real Date  ← NEW
    // WHY: A blank Trade Time Stamp means this row will sort to the wrong
    //      position in any pipeline re-run, silently breaking block sequencing.
    //      validateAndCleanImportToHelperV3 should have caught this — if it
    //      reaches Staging blank, something bypassed validation entirely.
    // ════════════════════════════════════════════════════════════════════════
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const ticker = cvStr(row, "ticker");
      const action = cvUpper(row, "action");
      if (!ticker || !TRADE_ACTIONS.includes(action)) continue;
      const ts = cv(row, "trade time stamp");
      if (!(ts instanceof Date) || isNaN(ts.getTime())) {
        flag(
          "MISSING_TIMESTAMP",
          "ERROR",
          i,
          cv(row, "account"),
          ticker,
          cv(row, "trade date"),
          action,
          "Trade Time Stamp",
          String(ts || "(blank)"),
          "Trade row is missing a valid Trade Time Stamp. This row cannot be sequenced " +
            "reliably. Re-check the source row in Helper and trace back to Import or Schwab Mapping.",
        );
      }
    }

    // ════════════════════════════════════════════════════════════════════════
    // CHECK 12 — Account is not exactly DT or LT  ← NEW
    // WHY: Any account value other than DT or LT means SUMIFS by account
    //      in the dashboard will silently miss these rows. Should have been
    //      caught by validateAndCleanImportToHelperV3 but defensively checked here.
    // ════════════════════════════════════════════════════════════════════════
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const ticker = cvStr(row, "ticker");
      if (!ticker) continue;
      const acct = cvUpper(row, "account");
      if (acct !== "DT" && acct !== "LT") {
        flag(
          "INVALID_ACCOUNT",
          "ERROR",
          i,
          acct,
          ticker,
          cv(row, "trade date"),
          cv(row, "action"),
          "Account",
          acct || "(blank)",
          'Account must be exactly "DT" or "LT". This row will be excluded from ' +
            "all account-filtered dashboard views. Fix in Helper and re-run Stage 3.",
        );
      }
    }

    // ════════════════════════════════════════════════════════════════════════
    // CHECK 13 — Closed position where Closing Date < Trade Date  ← NEW
    // WHY: A closing date earlier than the trade date is physically impossible
    //      and indicates a timestamp parsing error — likely a date-only field
    //      being set to the epoch (Jan 1 1900) instead of a real date.
    // ════════════════════════════════════════════════════════════════════════
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const ticker = cvStr(row, "ticker");
      if (!ticker) continue;
      const tradeDate = cvDate(row, "trade date");
      const closingDate = cvDate(row, "closing date");
      if (!tradeDate || !closingDate) continue;
      if (closingDate < tradeDate) {
        flag(
          "CLOSING_DATE_BEFORE_OPEN",
          "ERROR",
          i,
          cv(row, "account"),
          ticker,
          tradeDate,
          cv(row, "action"),
          "Closing Date",
          Utilities.formatDate(closingDate, tz, "M/d/yyyy"),
          "Closing Date (" +
            Utilities.formatDate(closingDate, tz, "M/d/yyyy") +
            ") is before Trade Date (" +
            Utilities.formatDate(tradeDate, tz, "M/d/yyyy") +
            "). This is a timestamp parsing error. Trade Duration will be negative or wrong.",
        );
      }
    }

    // ════════════════════════════════════════════════════════════════════════
    // SUMMARY — Per-account pipeline statistics
    // ════════════════════════════════════════════════════════════════════════
    const acctStats = {};
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const acct = cvUpper(row, "account");
      const ticker = cvStr(row, "ticker");
      if (!acct) continue;
      if (!acctStats[acct]) {
        acctStats[acct] = {
          rows: 0,
          opens: 0,
          closes: 0,
          openPosIds: new Set(),
          radRows: 0,
        };
      }
      acctStats[acct].rows++;
      if (ticker) {
        if (cvNum(row, "block start flag") === 1) acctStats[acct].opens++;
        if (cvNum(row, "block close flag/p&l") === 1) acctStats[acct].closes++;
        if (cvUpper(row, "action") === "RAD") acctStats[acct].radRows++;
        const posId = cvStr(row, "position id");
        if (posId && !posHasClose.has(posId))
          acctStats[acct].openPosIds.add(posId);
      }
    }

    for (const [acct, s] of Object.entries(acctStats)) {
      findings.push([
        "SUMMARY",
        "INFO",
        "",
        acct,
        "",
        "",
        "",
        "Pipeline Stats",
        `Rows: ${s.rows} | BlocksOpened: ${s.opens} | BlocksClosed: ${s.closes} | ` +
          `OpenPositions: ${s.openPosIds.size} | RAD Rows: ${s.radRows}`,
        "Informational only — overall pipeline health for this account.",
      ]);
    }
  } catch (e) {
    // Append a visible exception row so you always know something went wrong
    // even if the sheet write below partially succeeds.
    findings.push([
      "AUDIT_EXCEPTION",
      "ERROR",
      "",
      "",
      "",
      "",
      "",
      "Exception",
      e.message,
      "auditPipelineIntegrity threw an unexpected error. Stack: " +
        (e.stack || "(no stack)"),
    ]);
  } finally {
    // ══════════════════════════════════════════════════════════════════════
    // WRITE TO "Audit Results" — runs even if a check threw
    // ══════════════════════════════════════════════════════════════════════
    let auditSheet = ss.getSheetByName("Audit Results");
    if (!auditSheet) {
      auditSheet = ss.insertSheet("Audit Results");
    } else {
      auditSheet.clearContents();
      auditSheet.clearFormats();
    }

    const AUDIT_HEADERS = [
      "Check",
      "Severity",
      "Staging Row",
      "Account",
      "Ticker",
      "Trade Date",
      "Action",
      "Field",
      "Value",
      "Detail",
    ];

    auditSheet
      .getRange(1, 1, 1, AUDIT_HEADERS.length)
      .setValues([AUDIT_HEADERS])
      .setFontWeight("bold")
      .setBackground("#37474f")
      .setFontColor("#ffffff");

    if (findings.length > 0) {
      auditSheet
        .getRange(2, 1, findings.length, AUDIT_HEADERS.length)
        .setValues(findings);

      const SEVERITY_COLORS = {
        ERROR: "#fce8e6",
        WARN: "#fff8e1",
        INFO: "#e8f5e9",
        SUMMARY: "#e3f2fd",
      };

      for (let i = 0; i < findings.length; i++) {
        const severity = findings[i][1];
        auditSheet
          .getRange(i + 2, 1, 1, AUDIT_HEADERS.length)
          .setBackground(SEVERITY_COLORS[severity] || "#ffffff");
      }
      auditSheet.getRange(2, 2, findings.length, 1).setFontWeight("bold");
    } else {
      auditSheet
        .getRange(2, 1)
        .setValue("✅ No issues found — pipeline is clean.");
      auditSheet.getRange(2, 1).setBackground("#e8f5e9");
    }

    auditSheet.setFrozenRows(1);
    auditSheet.autoResizeColumns(1, AUDIT_HEADERS.length);
    auditSheet.setColumnWidth(AUDIT_HEADERS.length, 480);

    const errorCount = findings.filter((f) => f[1] === "ERROR").length;
    const warnCount = findings.filter((f) => f[1] === "WARN").length;
    const infoCount = findings.filter((f) => f[1] === "INFO").length;
    const summaryCount = findings.filter((f) => f[1] === "SUMMARY").length;

    SpreadsheetApp.getUi().alert(
      '✅ Audit complete! See "Audit Results" sheet.\n\n' +
        "🔴 Errors:   " +
        errorCount +
        "  (data integrity — fix before using P&L)\n" +
        "🟡 Warnings: " +
        warnCount +
        "  (likely problems — review manually)\n" +
        "🟢 Info:     " +
        infoCount +
        "  (informational only)\n" +
        "⚪ Summary:  " +
        summaryCount +
        "  (per-account pipeline stats)",
    );
  } // end finally
}
