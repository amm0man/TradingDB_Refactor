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