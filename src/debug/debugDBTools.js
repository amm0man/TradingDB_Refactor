/**
 * auditOpenPositions
 *
 * Standalone diagnostic function. Reads the Staging sheet and writes a
 * summary of every Trade Group ID / Position ID whose Running Position
 * Quantity is non-zero (i.e. the block never closed flat).
 *
 * Output sheet: "Open Positions Audit"
 * One row per open block. Columns:
 *   Account | Ticker | Trade Group ID | Position ID | Strategy Type |
 *   Trade Type | Block Number | Block Start Date | Last Activity Date |
 *   Days Open | Running Position Quantity | Legs In Block | First Action | Notes
 *
 * WHY a separate sheet instead of an alert:
 *   ~18,000 rows may produce dozens of open positions. A sheet lets you
 *   sort by Days Open, filter by Account, or CTRL+F a specific ticker
 *   without running the function again.
 *
 * HOW TO RUN:
 *   Apps Script editor → select auditOpenPositions → click Run.
 *   Or add a menu item pointing to this function.
 */
function auditOpenPositions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stagingSheet = ss.getSheetByName('Staging');
  if (!stagingSheet) {
    SpreadsheetApp.getUi().alert('Staging sheet not found.');
    return;
  }

  const tz = ss.getSpreadsheetTimeZone();

  // ── Read Staging ────────────────────────────────────────────────────────────
  const allData = stagingSheet.getDataRange().getValues();
  if (allData.length < 4) {
    SpreadsheetApp.getUi().alert('No data in Staging.');
    return;
  }

  // Build colMap from header row (row 1, index 0).
  const colMap = {};
  allData[0].forEach((h, i) => {
    if (typeof h === 'string' && h.trim()) colMap[h.trim().toLowerCase()] = i;
  });

  // Helper: safe column read — returns '' if column doesn't exist.
  const get = (row, col) => (colMap[col] !== undefined ? row[colMap[col]] : '');

  // Data starts at row 4 (index 3) — rows 1-3 are headers/metadata.
  const dataRows = allData.slice(3);

  // ── Build one summary object per Trade Group ID ─────────────────────────────
  // Key = Trade Group ID. We track:
  //   - The final Running Position Quantity (last row for this TG ID)
  //   - Block Number
  //   - Block Start Date (Trade Date of the blkStart=1 row)
  //   - Last Activity Date (Trade Date of the most recent row)
  //   - Leg count (rows in this TG ID)
  //   - First Action on the block
  //   - Account, Ticker, Strategy Type, Trade Type, Position ID
  const groups = {};

  for (const row of dataRows) {
    const tgId = get(row, 'trade group id').toString().trim();
    const posId = get(row, 'position id').toString().trim();
    const account = get(row, 'account').toString().trim();
    const ticker = get(row, 'ticker').toString().trim();
    const stratType = get(row, 'strategy type').toString().trim();
    const tradeType = get(row, 'trade type').toString().trim();
    const blockNum = get(row, 'block number');
    const blkStartFlag = get(row, 'block start flag');
    const runQty = Number(get(row, 'running position quantity')) || 0;
    const tradeDate = get(row, 'trade date');
    const action = get(row, 'action').toString().trim();
    const blockClose = get(row, 'block close flag/p&l');

    if (!tgId) continue; // skip ledger rows with no Trade Group ID

    if (!groups[tgId]) {
      groups[tgId] = {
        tgId,
        posId,
        account,
        ticker,
        stratType,
        tradeType,
        blockNum,
        blockStartDate: null,
        lastActivityDate: null,
        runningQty: 0,
        legCount: 0,
        firstAction: '',
        isClosed: false,
      };
    }

    const g = groups[tgId];

    // Always update to the latest runningQty and lastActivityDate seen.
    g.runningQty = runQty;
    g.legCount += 1;
    if (tradeDate instanceof Date) {
      if (!g.lastActivityDate || tradeDate > g.lastActivityDate) {
        g.lastActivityDate = tradeDate;
      }
    }

    // Capture block start date and first action from the blkStart=1 row.
    if (blkStartFlag === 1 || blkStartFlag === '1') {
      if (tradeDate instanceof Date) g.blockStartDate = tradeDate;
      g.firstAction = action;
    }

    // Mark closed if any row has Block Close Flag = 1.
    if (blockClose === 1 || blockClose === '1') g.isClosed = true;
  }

  // ── Filter to OPEN positions only ──────────────────────────────────────────
  // A position is "open" if:
  //   (a) Running Position Quantity !== 0   ← the primary signal
  //   OR
  //   (b) isClosed is false AND legCount > 0  ← catches blocks with no close row
  const today = new Date();
  const openGroups = Object.values(groups).filter(g => g.runningQty !== 0 || !g.isClosed);

  if (openGroups.length === 0) {
    SpreadsheetApp.getUi().alert('✅ No open positions found in Staging. All blocks closed flat!');
    return;
  }

  // Sort: Account A→Z, then Days Open descending (oldest open first).
  openGroups.sort((a, b) => {
    if (a.account !== b.account) return a.account.localeCompare(b.account);
    const daysA = a.blockStartDate ? (today - a.blockStartDate) / 86400000 : 0;
    const daysB = b.blockStartDate ? (today - b.blockStartDate) / 86400000 : 0;
    return daysB - daysA; // descending — oldest open at top
  });

  // ── Write to "Open Positions Audit" sheet ──────────────────────────────────
  let auditSheet = ss.getSheetByName('Open Positions Audit');
  if (!auditSheet) {
    auditSheet = ss.insertSheet('Open Positions Audit');
  } else {
    auditSheet.clearContents();
    auditSheet.clearFormats();
  }

  // Header row.
  const headers = [
    'Account',
    'Ticker',
    'Strategy Type',
    'Trade Type',
    'Trade Group ID',
    'Position ID',
    'Block Number',
    'Block Start Date',
    'Last Activity Date',
    'Days Open',
    'Running Qty',
    'Legs In Block',
    'First Action',
    'Notes',
  ];

  const outputRows = [headers];

  for (const g of openGroups) {
    const daysOpen = g.blockStartDate
      ? Math.round((today - g.blockStartDate) / 86400000)
      : '';

    const blockStartStr = g.blockStartDate
      ? Utilities.formatDate(g.blockStartDate, tz, 'MM/dd/yyyy')
      : 'Unknown';

    const lastActivityStr = g.lastActivityDate
      ? Utilities.formatDate(g.lastActivityDate, tz, 'MM/dd/yyyy')
      : 'Unknown';

    // Auto-diagnose common issues for the Notes column.
    let notes = '';
    if (g.runningQty < 0) notes = '⚠️ Negative running qty — possible extra closing leg';
    if (g.runningQty > 20) notes = '⚠️ Very high running qty — check for duplicate open rows';
    if (!g.isClosed && g.legCount === 1) notes = '⚠️ Single-leg block — may be missing paired leg(s)';
    if (daysOpen > 365) notes = (notes ? notes + ' | ' : '') + '⚠️ Open > 1 year — verify not a stale block';

    outputRows.push([
      g.account,
      g.ticker,
      g.stratType,
      g.tradeType,
      g.tgId,
      g.posId,
      g.blockNum,
      blockStartStr,
      lastActivityStr,
      daysOpen,
      g.runningQty,
      g.legCount,
      g.firstAction,
      notes,
    ]);
  }

  auditSheet.getRange(1, 1, outputRows.length, headers.length).setValues(outputRows);

  // ── Formatting ─────────────────────────────────────────────────────────────
  // Bold header row.
  auditSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9ead3');

  // Freeze header row.
  auditSheet.setFrozenRows(1);

  // Color rows by account: DT = light blue tint, LT = light yellow tint.
  for (let r = 2; r <= outputRows.length; r++) {
    const acct = outputRows[r - 1][0];
    const bg = acct === 'DT' ? '#dce6f1' : '#fff2cc';
    auditSheet.getRange(r, 1, 1, headers.length).setBackground(bg);
  }

  // Flag rows with Notes in red.
  for (let r = 2; r <= outputRows.length; r++) {
    const notesVal = outputRows[r - 1][13]; // Notes column
    if (notesVal && notesVal.includes('⚠️')) {
      auditSheet.getRange(r, 14).setFontColor('#cc0000').setFontWeight('bold');
    }
  }

  // Auto-resize all columns.
  for (let c = 1; c <= headers.length; c++) {
    auditSheet.autoResizeColumn(c);
  }

  // Navigate to the audit sheet.
  ss.setActiveSheet(auditSheet);

  SpreadsheetApp.getUi().alert(
    `Open Positions Audit complete.\n\n` +
    `Found ${openGroups.length} open position(s) across ${dataRows.length.toLocaleString()} Staging rows.\n\n` +
    `Results written to "Open Positions Audit" sheet.\n` +
    `Sorted by Account → Days Open (oldest first).`
  );
}

// Phase 3 debug to classify negative Running Position Quantity in Staging
function classifyNegRunningQtyFirstCross() {
  const EPS = 1e-8;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staging = ss.getSheetByName("Staging");
  if (!staging) {
    uiAlertSafe("Staging sheet not found.");
    return;
  }

  const all = staging.getDataRange().getValues();
  if (all.length < 4) {
    uiAlertSafe("Staging has no data rows.");
    return;
  }

  const headers = all[0].map(function (h) {
    return String(h || "").trim().toLowerCase();
  });
  function idx(name) {
    const i = headers.indexOf(name);
    if (i < 0) throw new Error("Missing Staging header: " + name);
    return i;
  }

  const iAcct = idx("account");
  const iTicker = idx("ticker");
  const iDate = idx("trade date");
  const iAction = idx("action");
  const iType = idx("trade type");
  const iSpread = idx("spread group id");
  const iQty = idx("quantity");
  const iRun = idx("running position quantity");
  const iPos = headers.indexOf("position id");
  const iTg = headers.indexOf("trade group id");
  const iExp = headers.indexOf("option expiration");
  const iStrike = headers.indexOf("option strike");
  const iCp = headers.indexOf("call/put");
  const iBlk = headers.indexOf("block number");

  function familyOf(g) {
    const action = String(g.firstAction || "").toUpperCase();
    const type = String(g.firstTradeType || "").toUpperCase();
    const spread = String(g.firstSpread || "").trim();
    const ticker = String(g.ticker || "");
    const block = Number(g.firstBlock || 1);

    if (type === "STOCK") return { family: "E", fix: "CRUMB_OR_FLOAT — skip if abs(run)<1e-8; else clamp stock" };
    if (action === "RAD") {
      if (/^\w+\d+$/.test(ticker)) {
        return { family: "D-ID", fix: "FIX_IDENTITY — ticker looks like OCC root+year (SQQQ1). Parse, then rebuild." };
      }
      return { family: "D", fix: "CLAMP_RAD — unmatched RAD, keep row, running stays 0" };
    }
    if (spread && (action.indexOf("TO CLOSE") !== -1)) {
      return { family: "C", fix: "CLAMP_EXTRA_SPREAD_CLOSE — stay on last closed TG, do not open TG002 short" };
    }
    if (!spread && action.indexOf("TO CLOSE") !== -1 && block <= 1) {
      return { family: "A", fix: "CLAMP_ORPHAN_CLOSE — no open exists, do not invent a short" };
    }
    if (!spread && action.indexOf("TO CLOSE") !== -1 && block > 1) {
      return { family: "B", fix: "DONE leftover-close-as-open (single option flip)" };
    }
    return { family: "?", fix: "SPOT_CHECK" };
  }

  const groups = {};
  for (let r = 3; r < all.length; r++) {
    const row = all[r];
    const ticker = String(row[iTicker] || "").trim().toUpperCase();
    const runQty = Number(row[iRun]);
    if (!ticker || !(runQty < -EPS)) continue;

    const acct = String(row[iAcct] || "").trim().toUpperCase();
    const key = acct + "|" + ticker;
    if (!groups[key]) {
      groups[key] = {
        account: acct,
        ticker: ticker,
        firstRow: r + 1,
        firstDate: row[iDate],
        firstAction: String(row[iAction] || ""),
        firstTradeType: String(row[iType] || ""),
        firstSpread: String(row[iSpread] || ""),
        firstQty: row[iQty],
        firstRunQty: runQty,
        firstPosId: iPos >= 0 ? String(row[iPos] || "") : "",
        firstTg: iTg >= 0 ? String(row[iTg] || "") : "",
        firstExp: iExp >= 0 ? row[iExp] : "",
        firstStrike: iStrike >= 0 ? row[iStrike] : "",
        firstCp: iCp >= 0 ? String(row[iCp] || "") : "",
        firstBlock: iBlk >= 0 ? row[iBlk] : "",
        minRunQty: runQty,
        lastRow: r + 1,
        lastDate: row[iDate],
        lastAction: String(row[iAction] || ""),
        lastRunQty: runQty,
        negCount: 0
      };
    }
    const g = groups[key];
    g.negCount++;
    if (runQty < g.minRunQty) g.minRunQty = runQty;
    g.lastRow = r + 1;
    g.lastDate = row[iDate];
    g.lastAction = String(row[iAction] || "");
    g.lastRunQty = runQty;
  }

  const outHeaders = [
    "Family",
    "Fix",
    "Account",
    "Ticker",
    "NegRowCount",
    "FirstStagingRow",
    "FirstDate",
    "FirstAction",
    "FirstTradeType",
    "FirstSpreadGroupId",
    "FirstTradeGroupId",
    "FirstQty",
    "FirstRunningQty",
    "FirstPositionId",
    "FirstBlock",
    "FirstExp",
    "FirstStrike",
    "FirstCP",
    "MinRunningQty",
    "LastStagingRow",
    "LastDate",
    "LastAction",
    "LastRunningQty"
  ];

  const body = Object.keys(groups)
    .sort(function (a, b) {
      return groups[b].negCount - groups[a].negCount;
    })
    .map(function (k) {
      const g = groups[k];
      const fam = familyOf(g);
      return [
        fam.family,
        fam.fix,
        g.account,
        g.ticker,
        g.negCount,
        g.firstRow,
        g.firstDate,
        g.firstAction,
        g.firstTradeType,
        g.firstSpread,
        g.firstTg,
        g.firstQty,
        g.firstRunQty,
        g.firstPosId,
        g.firstBlock,
        g.firstExp,
        g.firstStrike,
        g.firstCp,
        g.minRunQty,
        g.lastRow,
        g.lastDate,
        g.lastAction,
        g.lastRunQty
      ];
    });

  let out = ss.getSheetByName("NEG_RUNNING_QTY Classify");
  if (!out) out = ss.insertSheet("NEG_RUNNING_QTY Classify");
  out.clear();
  out.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders]);
  if (body.length) {
    out.getRange(2, 1, body.length, outHeaders.length).setValues(body);
  }
  out.setFrozenRows(1);

  uiAlertSafe(
    "NEG_RUNNING_QTY first-cross: " +
      body.length +
      " Account|Ticker families from " +
      body.reduce(function (n, r) { return n + Number(r[4] || 0); }, 0) +
      " negative rows (crumbs abs<1e-8 skipped). See sheet NEG_RUNNING_QTY Classify."
  );
}

/**
 * classifyNegSpxByPositionId
 *
 * WHY:
 *   classifyNegRunningQtyFirstCross groups by Account|Ticker, so all DT SPX
 *   negatives collapse into one family (30 rows, MinRunningQty -300).
 *   SPX is many expirations and strikes. This helper splits those Staging
 *   rows by Position ID when present, else by exp + strike + C/P.
 *
 * SCOPE:
 *   Account = DT, Ticker = SPX, Running Position Quantity < 0.
 *   Does not change Staging. Does not change leftover-close.
 *
 * OUTPUT:
 *   Sheet "NEG_SPX By Position ID" — one row per option key
 *   Sheet "NEG_SPX Rows" — every negative DT SPX Staging row
 *
 * HOW TO RUN:
 *   Apps Script editor → select classifyNegSpxByPositionId → Run.
 */
function classifyNegSpxByPositionId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staging = ss.getSheetByName("Staging");
  if (!staging) {
    uiAlertSafe("Staging sheet not found.");
    return;
  }

  const all = staging.getDataRange().getValues();
  if (all.length < 4) {
    uiAlertSafe("Staging has no data rows.");
    return;
  }

  const headers = all[0].map(function (h) {
    return String(h || "")
      .trim()
      .toLowerCase();
  });
  function idx(name) {
    const i = headers.indexOf(name);
    if (i < 0) throw new Error("Missing Staging header: " + name);
    return i;
  }

  const iAcct = idx("account");
  const iTicker = idx("ticker");
  const iDate = idx("trade date");
  const iTime = headers.indexOf("trade time");
  const iAction = idx("action");
  const iType = idx("trade type");
  const iSpread = idx("spread group id");
  const iTg = headers.indexOf("trade group id");
  const iPos = headers.indexOf("position id");
  const iStrat = headers.indexOf("strategy type");
  const iQty = idx("quantity");
  const iRun = idx("running position quantity");
  const iStart = headers.indexOf("block start flag");
  const iBlk = headers.indexOf("block number");
  const iClose = headers.indexOf("block close flag/p&l");
  const iExp = headers.indexOf("option expiration");
  const iStrike = headers.indexOf("option strike");
  const iCp = headers.indexOf("call/put");

  function kindFromFirstAction(action) {
    const a = String(action || "")
      .toUpperCase()
      .trim();
    if (a === "RAD") return "RAD_NO_LIVE";
    if (a.indexOf("TO CLOSE") !== -1) return "ORPHAN_OR_EXTRA_CLOSE";
    if (a.indexOf("TO OPEN") !== -1) return "OPEN_THEN_NEGATIVE";
    return "OTHER";
  }

  const groups = {};
  const detail = [];

  for (let r = 3; r < all.length; r++) {
    const row = all[r];
    const acct = String(row[iAcct] || "")
      .trim()
      .toUpperCase();
    const ticker = String(row[iTicker] || "")
      .trim()
      .toUpperCase();
    const runQty = Number(row[iRun]);
    if (acct !== "DT" || ticker !== "SPX" || !(runQty < 0)) continue;

    const posId = iPos >= 0 ? String(row[iPos] || "").trim() : "";
    const exp = iExp >= 0 ? row[iExp] : "";
    const strike = iStrike >= 0 ? row[iStrike] : "";
    const cp = iCp >= 0
      ? String(row[iCp] || "")
          .trim()
          .toUpperCase()
          .replace("CALL", "C")
          .replace("PUT", "P")
      : "";
    const identity =
      "DT|SPX|" +
      String(exp || "") +
      "|" +
      String(strike || "") +
      "|" +
      cp;
    const groupKey = posId || identity;

    if (!groups[groupKey]) {
      groups[groupKey] = {
        positionId: posId,
        identity: identity,
        firstRow: r + 1,
        firstDate: row[iDate],
        firstAction: String(row[iAction] || ""),
        firstTradeType: String(row[iType] || ""),
        firstSpread: String(row[iSpread] || ""),
        firstTg: iTg >= 0 ? String(row[iTg] || "") : "",
        firstStrat: iStrat >= 0 ? String(row[iStrat] || "") : "",
        firstQty: row[iQty],
        firstRunQty: runQty,
        firstExp: exp,
        firstStrike: strike,
        firstCp: cp,
        firstBlkStart: iStart >= 0 ? row[iStart] : "",
        firstBlkNum: iBlk >= 0 ? row[iBlk] : "",
        firstBlkClose: iClose >= 0 ? row[iClose] : "",
        minRunQty: runQty,
        lastRow: r + 1,
        lastDate: row[iDate],
        lastAction: String(row[iAction] || ""),
        lastRunQty: runQty,
        negCount: 0,
        kind: kindFromFirstAction(row[iAction]),
      };
    }

    const g = groups[groupKey];
    g.negCount++;
    if (runQty < g.minRunQty) g.minRunQty = runQty;
    g.lastRow = r + 1;
    g.lastDate = row[iDate];
    g.lastAction = String(row[iAction] || "");
    g.lastRunQty = runQty;

    detail.push([
      r + 1,
      row[iDate],
      iTime >= 0 ? row[iTime] : "",
      String(row[iAction] || ""),
      String(row[iType] || ""),
      String(row[iSpread] || ""),
      iTg >= 0 ? String(row[iTg] || "") : "",
      posId,
      iStrat >= 0 ? String(row[iStrat] || "") : "",
      row[iQty],
      runQty,
      iStart >= 0 ? row[iStart] : "",
      iBlk >= 0 ? row[iBlk] : "",
      iClose >= 0 ? row[iClose] : "",
      exp,
      strike,
      cp,
      identity,
      kindFromFirstAction(row[iAction]),
    ]);
  }

  const sumHeaders = [
    "Kind",
    "Position ID",
    "Option Identity",
    "NegRowCount",
    "FirstStagingRow",
    "FirstDate",
    "FirstAction",
    "FirstTradeType",
    "FirstSpreadGroupId",
    "FirstTradeGroupId",
    "FirstStrategyType",
    "FirstQty",
    "FirstRunningQty",
    "FirstExp",
    "FirstStrike",
    "FirstCP",
    "FirstBlkStart",
    "FirstBlkNum",
    "FirstBlkClose",
    "MinRunningQty",
    "LastStagingRow",
    "LastDate",
    "LastAction",
    "LastRunningQty",
  ];

  const sumBody = Object.keys(groups)
    .sort(function (a, b) {
      return groups[a].firstRow - groups[b].firstRow;
    })
    .map(function (k) {
      const g = groups[k];
      return [
        g.kind,
        g.positionId,
        g.identity,
        g.negCount,
        g.firstRow,
        g.firstDate,
        g.firstAction,
        g.firstTradeType,
        g.firstSpread,
        g.firstTg,
        g.firstStrat,
        g.firstQty,
        g.firstRunQty,
        g.firstExp,
        g.firstStrike,
        g.firstCp,
        g.firstBlkStart,
        g.firstBlkNum,
        g.firstBlkClose,
        g.minRunQty,
        g.lastRow,
        g.lastDate,
        g.lastAction,
        g.lastRunQty,
      ];
    });

  const detHeaders = [
    "StagingRow",
    "Trade Date",
    "Trade Time",
    "Action",
    "Trade Type",
    "Spread Group ID",
    "Trade Group ID",
    "Position ID",
    "Strategy Type",
    "Quantity",
    "Running Position Quantity",
    "Block Start Flag",
    "Block Number",
    "Block Close Flag/P&L",
    "Option Expiration",
    "Option Strike",
    "Call/Put",
    "Option Identity",
    "Kind",
  ];

  function writeSheet(name, headerRow, bodyRows) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.clear();
    sh.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    if (bodyRows.length) {
      sh.getRange(2, 1, bodyRows.length, headerRow.length).setValues(bodyRows);
    }
    sh.setFrozenRows(1);
  }

  writeSheet("NEG_SPX By Position ID", sumHeaders, sumBody);
  writeSheet("NEG_SPX Rows", detHeaders, detail);

  uiAlertSafe(
    "DT SPX negatives: " +
      detail.length +
      " rows across " +
      sumBody.length +
      " Position ID / option keys. See NEG_SPX By Position ID and NEG_SPX Rows.",
  );
}