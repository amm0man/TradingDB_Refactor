/**
 * Phase3Processing.js
 *
 * Phase 3 – Core processing pipeline.
 *
 * High-level job:
 *   Take the clean "Schwab Mapping" sheet (from Phase 2) and turn it into
 *   the final Staging data that can be appended to Master.
 *
 * Pipeline steps (also run together by refreshAllScripts):
 *   1. copyMappingToImportByHeaders()     → Schwab Mapping → Import
 *   2. validateAndCleanImportToHelperV3() → Import → Helper (validation + timestamps)
 *   3. populateStagingWithBlockLogicV3()  → Helper → Staging (block / position logic)
 *
 * Also contains:
 *   - Master utilities (append, backup, clear)
 *   - Supporting helpers used by the three steps above
 *
 * The custom menu (onOpen) now lives in Menu.js.
 *
 * Related files:
 *   - Menu.js                            (custom "DB Tools" menu)
 *   - mapSchwabImportByHeadersV3.js      (Phase 2)
 *   - BuildUnifiedImportV3.js            (Phase 1)
 *   - ImportIssues.js / SheetBlanking.js (shared)
 */

/**  Helper functions
 * 
 * 
*/

// =========================================================================
// SMALL SHARED HELPERS (used by the Phase 3 steps below)
// =========================================================================
function logAction(action, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('DB_log');
  // If the log sheet doesn't exist, create it and add headers
  if (!logSheet) {
    logSheet = ss.insertSheet('DB_log');
    logSheet.appendRow(['Timestamp', 'Action', 'Details']);
  }
  // Log the action
  logSheet.appendRow([
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    action,
    details
  ]);
}

// Tiny helpers
function getValByHeader(row, headers, colName) {
  const idx = headers.indexOf(colName.toLowerCase());
  return idx > -1 ? row[idx] : "";
}

function ensureValidationErrorSheet() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validation Errors");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Validation Errors");
    sheet.appendRow(["Row", "Column", "Error", "Suggested Fix"]);
  }
  return sheet;
}

/** End Helpers
 * 
 * 
 */

// =========================================================================
// ORCHESTRATOR
//   Runs the three Phase 3 steps under one shared RunId so all Staging Issues
//   rows from a full refresh can be filtered together.
// =========================================================================
function refreshAllScripts() {
  // Set a stable RunId so all three Stage 3 steps share one RunId in Staging Issues.
  // The importIssuesStart / stagingIssuesStart functions read this key automatically.
  // const runId = new Date().toISOString(); //writes Zulu time to RunID
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const runId = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss zzz");
  setSetting('ACTIVE_IMPORT_RUN_ID', runId);

  try {
    copyMappingToImportByHeaders();         // Step 1: Schwab Mapping → Import
    validateAndCleanImportToHelperV3();     // Step 2: Validation + full timestamps
    populateStagingWithBlockLogicV3();      // Step 3: Block logic using Trade Time Stamp
    // auditPipelineIntegrity();               // Step 4 Read-only scan of Staging sheet. Detects data quality problems that slip past the normal pipeline 
    SpreadsheetApp.getUi().alert(
      "✅ Full refresh complete!\n" +
      "RunId: " + runId + "\n\n" +
      "Check the 'Staging Issues' sheet and filter by this RunId to review."
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Pipeline error: " + e.message + "\nCheck Staging Issues sheet.");
    throw e;
  } finally {
    // Always clear the active RunId — even if an error occurred
    setSetting('ACTIVE_IMPORT_RUN_ID', '');
  }
}

// =========================================================================
// STEP 1: copyMappingToImportByHeaders
//   Copies the current "Schwab Mapping" sheet into the "Import" sheet
//   using header-name matching (order-independent).
//   Also does early Strategy Type / settlement-row handling that Phase 3 needs.
// =========================================================================
function copyMappingToImportByHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Staging Issues CTX ───────────────────────────────────────────────────
  const ctx = stagingIssuesStart('copyMappingToImportByHeaders');
  importIssuesSetMetric(ctx, 'SourceSheet', 'Schwab Mapping');
  importIssuesSetMetric(ctx, 'DestSheet', 'Import');
  // ────────────────────────────────────────────────────────────────────────

  // ── PRIORITY 1: try/catch/finally so the Issues log ALWAYS gets flushed ─
  // WHY: Without this wrapper, any mid-function throw (bad sheet, empty data,
  // regex error, etc.) exits silently and leaves ZERO log output in
  // Staging Issues for this step — the worst time to have no diagnostics.
  // The finally block guarantees stagingIssuesFlush(ctx) runs even on crash.
  try {

    var mappingSheet = ss.getSheetByName('Schwab Mapping');
    var importSheet = ss.getSheetByName('Import');

    // Get headers
    var mappingHeaders = mappingSheet.getRange(1, 1, 1, mappingSheet.getLastColumn()).getValues()[0];
    var importHeaders = importSheet.getRange(1, 1, 1, importSheet.getLastColumn()).getValues()[0];

    // Build column index map (temporarily INCLUDE Description so we can parse
    // it for RAD/EXP/EXERCISE rows)
    var colMap = [];
    var descriptionColInImport = -1;
    for (var i = 0; i < importHeaders.length; i++) {
      var headerName = importHeaders[i].trim().toLowerCase();
      if (headerName === 'description') {
        colMap.push(null);
        descriptionColInImport = i;
      } else {
        var idx = mappingHeaders.findIndex(function (x) {
          return typeof x === 'string' && x.trim().toLowerCase() === headerName;
        });
        colMap.push(idx >= 0 ? idx : null);
      }
    }

    // === ROBUST TIMESTAMP COLUMN MAPPING ===
    const tsHeaderNames = ['trade time stamp', 'trade date', 'trade time'];
    for (let h of tsHeaderNames) {
      const importIdx = importHeaders.findIndex(header => header.trim().toLowerCase() === h);
      if (importIdx > -1) {
        const mappingIdx = mappingHeaders.findIndex(
          header => typeof header === 'string' && header.trim().toLowerCase() === h
        );
        if (mappingIdx >= 0) colMap[importIdx] = mappingIdx;
      }
    }

    // Get all data from Schwab Mapping (row 2 down)
    var lastMappingRow = mappingSheet.getLastRow();
    if (lastMappingRow < 2) {
      SpreadsheetApp.getUi().alert('No data found to copy from Schwab Mapping.');
      return; // finally will still flush ctx
    }
    var numRows = lastMappingRow - 1;
    var mappingData = mappingSheet.getRange(2, 1, numRows, mappingHeaders.length).getValues();

    // ── CTX: record source row count ────────────────────────────────────────
    importIssuesSetMetric(ctx, 'SourceRowsReadExclHeader', numRows);
    // ────────────────────────────────────────────────────────────────────────

    // === Hoist ALL column-index lookups ABOVE the row loop ==================
    // WHY: These never change between rows. Declaring them once is faster and
    // avoids scope issues inside nested if-blocks.
    //
    // PRIORITY 3 FIX: Removed duplicate `totalCostImportIdx` declaration.
    // It was identical to `totalCostIndex` — both pointed to the same column.
    // All references below now use `totalCostIndex` exclusively.
    const tickerIndex = importHeaders.findIndex(h => h.trim().toLowerCase() === 'ticker');
    const quantityIndex = importHeaders.findIndex(h => h.trim().toLowerCase() === 'quantity');
    const strikeIndex = importHeaders.findIndex(h => h.trim().toLowerCase() === 'option strike');
    const expIndex = importHeaders.findIndex(h => h.trim().toLowerCase() === 'option expiration');
    const cpIndex = importHeaders.findIndex(h => h.trim().toLowerCase() === 'call/put');
    const totalCostIndex = importHeaders.findIndex(h => h.trim().toLowerCase() === 'total cost');
    const signedQuantityIndex = importHeaders.findIndex(h => h.trim().toLowerCase() === 'signed quantity');
    const strategyTypeIndex = importHeaders.findIndex(h => h.trim().toLowerCase() === 'strategy type');
    const tradeTypeIndex = importHeaders.findIndex(h => h.trim().toLowerCase() === 'trade type');
    const entryPriceIndex = importHeaders.findIndex(h => h.trim().toLowerCase() === 'entry price');
    const actionIdx = importHeaders.findIndex(h => h.trim().toLowerCase() === 'action');
    const corpActionsIndex = importHeaders.findIndex(h => h.trim().toLowerCase() === 'corporate actions');
    const descriptionIdx = mappingHeaders.findIndex(
      x => typeof x === 'string' && x.trim().toLowerCase() === 'description'
    );
    // ========================================================================

    // ── CTX counters — declared before the loop ──────────────────────────────
    var outputData = [];
    let skippedCount = 0;
    let radParsed = 0;
    let radFailed = 0;
    let tickerNorms = 0;
    // ────────────────────────────────────────────────────────────────────────

    for (var r = 0; r < mappingData.length; r++) {

      // Build rowArr mapped to Import columns
      var rowArr = [];
      for (var c = 0; c < colMap.length; c++) {
        rowArr.push(colMap[c] === null ? '' : mappingData[r][colMap[c]]);
      }

      // Read Action and Description for this row
      var action = (rowArr[actionIdx] || '').toString().trim().toUpperCase();
      var descStr = ((descriptionIdx >= 0) ? mappingData[r][descriptionIdx] : '').toString().trim();
      var upperDesc = descStr.toUpperCase();

      // TotalCost for skip guards — uses totalCostIndex (duplicate removed, Priority 3)
      let totalCostVal = (totalCostIndex > -1) ? (Number(rowArr[totalCostIndex]) || 0) : 0;

      // ── SKIP 1: TOS duplicate BTO Stock row from EXERCISE ─────────────────
      // Schwab emits both a correct "Synthetic STOCK leg from EXERCISE" row
      // (TotalCost = Price × Qty) AND a spurious "TOS Trades BUY STOCK TO OPEN"
      // row where TotalCost is 100× too large due to the options contract multiplier.
      // FINGERPRINT: TotalCost ≈ EntryPrice × Quantity × 100 (0.1% float tolerance).
      if (upperDesc.includes('TOS TRADES') &&
        upperDesc.includes('BUY') &&
        upperDesc.includes('STOCK') &&
        upperDesc.includes('TO OPEN')) {

        const epForSkip = (entryPriceIndex > -1) ? (Number(rowArr[entryPriceIndex]) || 0) : 0;
        const qtyForSkip = (quantityIndex > -1) ? (Number(rowArr[quantityIndex]) || 0) : 0;
        const expectedCorrectCost = epForSkip * qtyForSkip;

        const isOptionsMultiplierError = expectedCorrectCost > 0 &&
          Math.abs(totalCostVal - expectedCorrectCost * 100) / (expectedCorrectCost * 100) < 0.001;

        if (isOptionsMultiplierError) {
          skippedCount++;
          importIssuesAdd(ctx, 'SKIP', r + 2, 'Description',
            upperDesc.substring(0, 80),
            'TOS EXERCISE duplicate skipped — TotalCost (' + totalCostVal +
            ') is 100× the correct cost (' + expectedCorrectCost +
            '). Synthetic STOCK leg row carries the correct values.');
          continue;
        }
        // If NOT a 100× error → fall through and keep the row normally.
      }

      // ── SKIP 2: Synthetic exercise stock leg duplicate ────────────────────
      if (upperDesc.includes('EXERCISE') &&
        upperDesc.includes('SYNTHETIC') &&
        upperDesc.includes('STOCK LEG') &&
        !upperDesc.includes('FROM EXERCISE')) {
        skippedCount++;
        importIssuesAdd(ctx, 'SKIP', r + 2, 'Description',
          upperDesc.substring(0, 80),
          'Synthetic exercise stock leg skipped');
        continue;
      }

      // ── Corporate Action / RAD / EXP / Exercise parsing ───────────────────
      if ((action === 'RAD' || action === 'EXP' ||
        upperDesc.includes('EXERCISE') || upperDesc.includes('ASSIGNMENT') ||
        upperDesc.includes('BOT') || upperDesc.includes('UPON') ||
        upperDesc.includes('SYNTHETIC') || upperDesc.includes('REMOVED DUE TO')) && descStr) {

        // CORPORATE ACTION PASS-THROUGH ──────────────────────────────────────
        // Phase 2 is the single authoritative source for Corporate Actions tagging.
        // colMap already populated rowArr[corpActionsIndex] from Schwab Mapping.
        // We check whether that value is present and skip option parsing if so.
        const existingCorpAction = (corpActionsIndex !== -1)
          ? String(rowArr[corpActionsIndex] || '').trim()
          : '';
        const corpUpper = existingCorpAction.toUpperCase();

        if (existingCorpAction) {
          const corpTicker = (tickerIndex > -1)
            ? String(rowArr[tickerIndex] || '').trim().toUpperCase()
            : '';

          // Make stock-lineage corporate actions explicit before they ever reach Import.
          // This avoids downstream ambiguity where the row survives validation but still
          // looks too blank to the stock block logic.
          if ((corpUpper === 'SYMBOL CHANGE' || corpUpper === 'SPLIT') && corpTicker) {
            if (actionIdx > -1 && !String(rowArr[actionIdx] || '').trim()) {
              rowArr[actionIdx] = corpUpper;
            }
            if (tradeTypeIndex > -1) rowArr[tradeTypeIndex] = 'Stock';
            if (strategyTypeIndex > -1 && !String(rowArr[strategyTypeIndex] || '').trim()) {
              rowArr[strategyTypeIndex] = 'LONG STOCK';
            }

            importIssuesAdd(
              ctx,
              'INFO',
              r + 2,
              'Corporate Actions',
              corpUpper + ' | ' + corpTicker,
              'Corporate action row passed through with explicit stock Action / Trade Type / Strategy Type for downstream block logic.'
            );
          }

          if (descriptionColInImport !== -1) rowArr[descriptionColInImport] = '';
          outputData.push(rowArr);
          continue;
        }
        // END CORPORATE ACTION PASS-THROUGH ───────────────────────────────────

        // Guard: BOT/SOLD UPON rows are stock legs from exercise/assignment.
        // They are handled by the dedicated BOT/SOLD UPON parser below and must
        // NEVER enter the option-removal parser, even if the description also
        // contains EXERCISE or ASSIGNMENT.
        const isBotSoldUpon = ((upperDesc.includes('BOT') || upperDesc.includes('SOLD'))
          && upperDesc.includes('UPON'));

        const looksLikeOptionRemoval = !isBotSoldUpon && (
          upperDesc.includes('REMOVAL OF OPTION') ||  // Formats A and D
          upperDesc.includes('REMOVED DUE TO') ||  // Format B
          /\.[A-Z]{1,6}\d{6}[CP]\d/i.test(descStr)    // Format C: bare OCC symbol
        );

        // WHY !isBotSoldUpon prefix on the EXP arm: Schwab stamps some
        // BOT/SOLD UPON stock-delivery legs with action=EXP. Without this guard
        // they fall into the option-removal parser and fail all four formats.
        if (!isBotSoldUpon && (looksLikeOptionRemoval || action === 'EXP')) {

          let parsed = false;

          // ── FORMAT A ───────────────────────────────────────────────────────
          // "Removal of Option due to expiration QTY TICKER 100 (opt) DD MON YY[YY] STRIKE CALL|PUT"
          const removalMatch = descStr.match(
            /Removal of Option due to expiration\s*([-+]?\d+\.?\d*)\s*(\S+)\s+\d+\s*(?:\([^)]+\))?\s*(\d{1,2})\s*(\w{3})\s*(\d{2,4})\s*(\d+\.?\d*)\s*(CALL|PUT)/i
          );
          if (removalMatch) {
            const qtyStr = removalMatch[1];
            const tkr = removalMatch[2];
            const day = removalMatch[3];
            const monStr = removalMatch[4].toUpperCase();
            const year = removalMatch[5];
            const strikeStr = removalMatch[6];
            const cpStr = removalMatch[7].toUpperCase();

            const yearFull = year.length === 2
              ? (Number(year) <= 29 ? '20' + year : '19' + year)
              : year;

            // PRIORITY 4: Ticker normalization ($SPX.X → SPX) — Format A already had this. ✅
            if (tickerIndex > -1 && tkr) {
              const tkrNorm = tkr.toUpperCase().replace(/^\$/, '').replace(/\.[A-Z]+$/, '');
              rowArr[tickerIndex] = tkrNorm;
              if (tkrNorm !== tkr.toUpperCase()) {
                tickerNorms++;
                importIssuesAdd(ctx, 'INFO', r + 2, 'Ticker',
                  tkr + ' → ' + tkrNorm,
                  'RAD ticker normalized from Schwab alternate index symbol (Format A)');
              }
            }

            const qtyNum = Number(qtyStr);
            if (quantityIndex > -1) rowArr[quantityIndex] = qtyNum;
            if (signedQuantityIndex > -1) rowArr[signedQuantityIndex] = qtyNum;
            if (strikeIndex > -1) rowArr[strikeIndex] = Number(strikeStr);
            if (cpIndex > -1) rowArr[cpIndex] = (cpStr === 'CALL') ? 'C' : 'P';
            if (expIndex > -1) {
              const monthMap = { JAN: 1, FEB: 2, MAR: 3, APR: 4, MAY: 5, JUN: 6, JUL: 7, AUG: 8, SEP: 9, OCT: 10, NOV: 11, DEC: 12 };
              rowArr[expIndex] = new Date(Number(yearFull), (monthMap[monStr] || 1) - 1, Number(day));
            }
            if (totalCostIndex > -1) rowArr[totalCostIndex] = 0;
            if (entryPriceIndex > -1) rowArr[entryPriceIndex] = 0;
            if (strategyTypeIndex > -1) rowArr[strategyTypeIndex] = '';
            if (tradeTypeIndex > -1) rowArr[tradeTypeIndex] = 'Option';
            radParsed++;
            parsed = true;
          }

          // ── FORMAT B ───────────────────────────────────────────────────────
          // "Removed due to Expiration/Assignment ... QTY .OCCSTRING"
          // OCC symbol: .TICKER[6-digit YYMMDD][C|P][STRIKE]
          // Examples: "... -6.0 .SPY241231P580"  |  "... 2.0 .UUUU241220P6"
          if (!parsed) {
            const occMatch = descStr.match(
              /Removed due to (?:Expiration|Assignment).*?\s+([-+]?\d+\.?\d*)\s+\.([A-Z0-9]+?)(\d{6})([CP])(\d+\.?\d*)\s*$/i
            );
            if (occMatch) {
              const qtyStr = occMatch[1];
              const tkr = occMatch[2].toUpperCase();
              const yymmdd = occMatch[3];
              const cpStr = occMatch[4].toUpperCase();
              const strikeStr = occMatch[5];

              const yy = Number(yymmdd.substring(0, 2));
              const mm = Number(yymmdd.substring(2, 4)) - 1;
              const dd = Number(yymmdd.substring(4, 6));
              const fullYear = yy <= 29 ? 2000 + yy : 1900 + yy;

              // PRIORITY 4 FIX: Format B was missing ticker normalization.
              // Added same $-strip and .X-suffix-strip as Formats A and D.
              if (tickerIndex > -1) {
                const tkrNorm = tkr.replace(/^\$/, '').replace(/\.[A-Z]+$/, '');
                rowArr[tickerIndex] = tkrNorm;
                if (tkrNorm !== tkr) {
                  tickerNorms++;
                  importIssuesAdd(ctx, 'INFO', r + 2, 'Ticker',
                    tkr + ' → ' + tkrNorm,
                    'RAD ticker normalized from Schwab alternate index symbol (Format B)');
                }
              }

              const qtyNum = Number(qtyStr);
              if (quantityIndex > -1) rowArr[quantityIndex] = qtyNum;
              if (signedQuantityIndex > -1) rowArr[signedQuantityIndex] = qtyNum;
              if (strikeIndex > -1) rowArr[strikeIndex] = Number(strikeStr);
              if (cpIndex > -1) rowArr[cpIndex] = cpStr; // already C or P
              if (expIndex > -1) rowArr[expIndex] = new Date(fullYear, mm, dd);
              if (totalCostIndex > -1) rowArr[totalCostIndex] = 0;
              if (entryPriceIndex > -1) rowArr[entryPriceIndex] = 0;
              if (strategyTypeIndex > -1) rowArr[strategyTypeIndex] = '';
              if (tradeTypeIndex > -1) rowArr[tradeTypeIndex] = 'Option';

              const isAssignment = upperDesc.includes('ASSIGNMENT');
              if (isAssignment) {
                importIssuesAdd(ctx, 'INFO', r + 2, 'Action',
                  tkr + ' ' + cpStr + ' ' + strikeStr + ' exp ' + yymmdd,
                  'Assignment row parsed via Format B OCC symbol — verify stock leg exists in dataset');
              }
              radParsed++;
              parsed = true;
            }
          }

          // ── FORMAT C ───────────────────────────────────────────────────────
          // RAD row with bare OCC symbol — no "Removed due to" prefix.
          // Example: "PUT ENERGY FUELS INC $6 EXP 08/16/24: ASG: 1.0 .UUUU240816P6"
          if (!parsed) {
            const occBareMatch = descStr.match(
              /([-+]?\d+\.?\d*)\s+\.([A-Z0-9]+?)(\d{6})([CP])(\d+\.?\d*)\s*$/i
            );
            if (occBareMatch) {
              const qtyStr = occBareMatch[1];
              const tkr = occBareMatch[2].toUpperCase();
              const yymmdd = occBareMatch[3];
              const cpStr = occBareMatch[4].toUpperCase();
              const strikeStr = occBareMatch[5];

              const yy = Number(yymmdd.substring(0, 2));
              const mm = Number(yymmdd.substring(2, 4)) - 1;
              const dd = Number(yymmdd.substring(4, 6));
              const fullYear = yy <= 29 ? 2000 + yy : 1900 + yy;

              // PRIORITY 4 FIX: Format C was missing ticker normalization.
              // Added same $-strip and .X-suffix-strip as Formats A and D.
              if (tickerIndex > -1) {
                const tkrNorm = tkr.replace(/^\$/, '').replace(/\.[A-Z]+$/, '');
                rowArr[tickerIndex] = tkrNorm;
                if (tkrNorm !== tkr) {
                  tickerNorms++;
                  importIssuesAdd(ctx, 'INFO', r + 2, 'Ticker',
                    tkr + ' → ' + tkrNorm,
                    'RAD ticker normalized from Schwab alternate index symbol (Format C)');
                }
              }

              const qtyNum = Number(qtyStr);
              if (quantityIndex > -1) rowArr[quantityIndex] = qtyNum;
              if (signedQuantityIndex > -1) rowArr[signedQuantityIndex] = qtyNum;
              if (strikeIndex > -1) rowArr[strikeIndex] = Number(strikeStr);
              if (cpIndex > -1) rowArr[cpIndex] = cpStr;
              if (expIndex > -1) rowArr[expIndex] = new Date(fullYear, mm, dd);
              if (totalCostIndex > -1) rowArr[totalCostIndex] = 0;
              if (entryPriceIndex > -1) rowArr[entryPriceIndex] = 0;
              if (strategyTypeIndex > -1) rowArr[strategyTypeIndex] = '';
              if (tradeTypeIndex > -1) rowArr[tradeTypeIndex] = 'Option';

              const isAssignment = upperDesc.includes('ASG:');
              importIssuesAdd(ctx, 'INFO', r + 2, 'Action',
                tkr + ' ' + cpStr + ' $' + strikeStr + ' exp ' + yymmdd,
                'Format C — bare OCC assignment/expiration parsed. ' +
                (isAssignment
                  ? 'ASSIGNMENT: verify BOT UPON stock leg exists on same date.'
                  : 'EXPIRATION'));
              radParsed++;
              parsed = true;
            }
          }

          // ── FORMAT D ───────────────────────────────────────────────────────
          // "Removal of Option due to exercise/assignment [-]QTY TICKER
          //  100 (PERIOD) DD Mon YY[YY] STRIKE CALL|PUT"
          // Examples:
          //   "REMOVAL OF OPTION DUE TO EXERCISE -10.0 SPY 100 (Weeklys) 10 JUN 22 402 PUT"
          //   "Removal of Option due to exercise -1.0 $SPX.X 100 (WEEKLY) 31 Aug 2023 4515.0 PUT"
          if (!parsed) {
            const removalExAsgMatch = descStr.match(
              /Removal of Option due to (?:exercise|assignment)[-\s]?([-\d.]{1,23})\s+(\$?[A-Z][A-Z0-9.]*)\s+100\s*(?:\([^)]+\))?\s*(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{2,4})\s+([\d.]+)\s+(CALL|PUT)/i
            );
            if (removalExAsgMatch) {
              const qtyStr = removalExAsgMatch[1];
              const tkr = removalExAsgMatch[2];
              const day = removalExAsgMatch[3];
              const monStr = removalExAsgMatch[4].toUpperCase();
              const year = removalExAsgMatch[5];
              const strikeStr = removalExAsgMatch[6];
              const cpStr = removalExAsgMatch[7].toUpperCase();

              const yearFull = year.length === 2
                ? (Number(year) <= 29 ? '20' + year : '19' + year)
                : year;

              // PRIORITY 4: Ticker normalization — Format D already had this. ✅
              const tkrNorm = tkr.toUpperCase().replace(/^\$/, '').replace(/\.[A-Z]+$/, '');
              if (tickerIndex > -1) rowArr[tickerIndex] = tkrNorm;
              if (tkrNorm !== tkr.toUpperCase()) tickerNorms++;

              const qtyNum = Number(qtyStr);
              if (quantityIndex > -1) rowArr[quantityIndex] = qtyNum;
              if (signedQuantityIndex > -1) rowArr[signedQuantityIndex] = qtyNum;
              if (strikeIndex > -1) rowArr[strikeIndex] = Number(strikeStr);
              if (cpIndex > -1) rowArr[cpIndex] = cpStr === 'CALL' ? 'C' : 'P';
              if (expIndex > -1) {
                const monthMap = { JAN: 1, FEB: 2, MAR: 3, APR: 4, MAY: 5, JUN: 6, JUL: 7, AUG: 8, SEP: 9, OCT: 10, NOV: 11, DEC: 12 };
                rowArr[expIndex] = new Date(Number(yearFull), (monthMap[monStr] || 1) - 1, Number(day));
              }
              if (totalCostIndex > -1) rowArr[totalCostIndex] = 0;
              if (entryPriceIndex > -1) rowArr[entryPriceIndex] = 0;
              if (strategyTypeIndex > -1) rowArr[strategyTypeIndex] = '';
              if (tradeTypeIndex > -1) rowArr[tradeTypeIndex] = 'Option';

              const isAsg = upperDesc.includes('ASSIGNMENT');
              importIssuesAdd(ctx, 'INFO', r + 2, 'Action',
                `${tkrNorm} ${cpStr === 'CALL' ? 'C' : 'P'} ${strikeStr} exp ${day}${monStr}${yearFull}`,
                `Format D: Removal of Option due to ${isAsg ? 'ASSIGNMENT' : 'EXERCISE'} parsed. ` +
                (isAsg
                  ? 'Confirm matching BUY/SELL stock leg exists on the same Trade Date.'
                  : 'Confirm matching BOT UPON row exists on the same Trade Date.'));
              radParsed++;
              parsed = true;
            }
          }
          // ── END FORMAT D ───────────────────────────────────────────────────

          // No format matched — genuine parse failure
          if (!parsed) {
            radFailed++;
            importIssuesAdd(ctx, 'WARN', r + 2, 'Description', descStr.substring(0, 120),
              'RAD/Removal row — no regex matched. New Schwab format? Add a Format E block to the RAD parser.');
          }

        } // end option-removal parser block

        // ── BOT / SOLD UPON Stock Assignment & Exercise Legs ─────────────────
        // Schwab emits a separate stock row for every option exercise or assignment.
        // The row arrives with a blank Action but Ticker is already pre-mapped from
        // the Schwab Mapping Ticker column. We set Action, Qty, Entry Price, and
        // Total Cost — and MUST NOT overwrite the pre-mapped ticker when the
        // description contains a CUSIP instead of a symbol.
        //
        // BOT  [QTY] [TICKER|CUSIP] UPON ... → action = BUY TO OPEN,    signedQty = +qty
        // SOLD [QTY] [TICKER|CUSIP] UPON ... → action = SELL TO CLOSE,  signedQty = -qty
        if ((upperDesc.includes('BOT') || upperDesc.includes('SOLD')) && upperDesc.includes('UPON')) {
          const botSoldUponMatch = descStr.match(/(BOT|SOLD)\s+([-\d.]+)\s+([A-Z0-9]+)\s+UPON/i);
          if (botSoldUponMatch) {
            const verbUpper = botSoldUponMatch[1].toUpperCase();
            const qtyNum = Math.abs(Number(botSoldUponMatch[2]));
            const descTkr = botSoldUponMatch[3].toUpperCase();
            const isSell = (verbUpper === 'SOLD');

            // CUSIP detection: all-digit "ticker" in description → use the
            // pre-mapped ticker from Schwab Mapping's Ticker column instead.
            const descTkrIsCusip = /^\d+$/.test(descTkr);
            const preMappedTkr = (tickerIndex > -1)
              ? rowArr[tickerIndex].toString().trim().toUpperCase()
              : '';
            const tkr = (descTkrIsCusip && preMappedTkr) ? preMappedTkr : descTkr;

            // Entry price: derive from TotalCost ÷ Qty (= strike on assignments).
            // Schwab stores debits negative for BOT rows; Math.abs() normalises both.
            const cost = Math.abs(totalCostVal) || 0;
            const ep = (qtyNum > 0 && cost > 0)
              ? Math.round((cost / qtyNum) * 10000) / 10000
              : 0;

            if (actionIdx > -1) rowArr[actionIdx] = isSell ? 'SELL TO CLOSE' : 'BUY TO OPEN';
            if (tickerIndex > -1) rowArr[tickerIndex] = tkr;
            if (quantityIndex > -1) rowArr[quantityIndex] = qtyNum;
            if (signedQuantityIndex > -1) rowArr[signedQuantityIndex] = isSell ? -qtyNum : qtyNum;
            if (entryPriceIndex > -1) rowArr[entryPriceIndex] = ep;
            if (totalCostIndex > -1) rowArr[totalCostIndex] = cost;
            if (strategyTypeIndex > -1) rowArr[strategyTypeIndex] = isSell ? '' : 'LONG STOCK';
            if (tradeTypeIndex > -1) rowArr[tradeTypeIndex] = 'Stock';

            let infoNote;
            if (descTkrIsCusip && preMappedTkr) {
              infoNote = `Description contained CUSIP "${descTkr}" — used pre-mapped ticker "${tkr}" from Schwab Mapping Ticker column. Verify ticker is correct.`;
            } else if (descTkrIsCusip && !preMappedTkr) {
              infoNote = `⚠ Description contained CUSIP "${descTkr}" and no pre-mapped ticker was available — CUSIP used as ticker. Fix the Ticker field manually in Import.`;
            } else {
              infoNote = isSell
                ? 'SOLD UPON stock delivery leg parsed. Confirm matching short call assignment or long put exercise RAD row exists on the same Trade Date.'
                : 'BOT UPON stock acquisition leg parsed. Confirm matching short put assignment or long call exercise RAD row exists on the same Trade Date.';
            }
            importIssuesAdd(ctx, 'INFO', r + 2, 'Action',
              `${isSell ? 'SELL TO CLOSE' : 'BUY TO OPEN'} ${qtyNum} ${tkr} @ ${ep}`,
              infoNote);
          }
        }
        // ── END BOT / SOLD UPON ───────────────────────────────────────────────

        // Non-RAD rows (EFN, JRN, EXERCISE, ASSIGNMENT, stock BTO, etc.) that
        // entered this outer block fall through here with no counter change.

        // === DNN Total Cost fix (synthetic exercise stock leg) ================
        if (totalCostIndex > -1 && entryPriceIndex > -1 && quantityIndex > -1 &&
          upperDesc.includes('EXERCISE') &&
          tickerIndex > -1 &&
          (rowArr[tickerIndex] || '').toString().trim() !== '' &&
          ((action || '').toString().toUpperCase() === 'BUY TO OPEN' ||
            (rowArr[actionIdx] || '').toString().toUpperCase().includes('BUY TO OPEN'))) {

          const entryPriceVal = Number(rowArr[entryPriceIndex]) || 0;
          const qtyVal = Number(rowArr[quantityIndex]) || 0;
          if (entryPriceVal > 0 && qtyVal > 0) {
            rowArr[totalCostIndex] = entryPriceVal * qtyVal;
            if (signedQuantityIndex > -1) rowArr[signedQuantityIndex] = Math.abs(qtyVal);
          }
        }

      } // end corporate action block

      // Always clear description column and keep the row
      if (descriptionColInImport > -1) rowArr[descriptionColInImport] = '';
      outputData.push(rowArr);

    } // end main for loop

    var finalNumRows = outputData.length;

    // ── POST-LOOP PASS: Forward-fill Strategy Type onto settlement rows ────────
    // WHY: RAD (option removal) and BOT/SOLD UPON (stock delivery) rows arrive
    // from Schwab with blank Ticker and blank Strategy Type. The parsers above
    // correctly populate Ticker, Strike, Expiration, and Call/Put from the
    // Description field, but they cannot know the original trade's Strategy Type
    // (PCS, CCS, Long Butterfly, Short IC, etc.) — that information only exists
    // on the opening legs written earlier in outputData.
    //
    // Without Strategy Type on settlement rows, populateStagingWithBlockLogicV3
    // cannot group them under the same Trade Group ID as the opening legs,
    // breaking block-level P&L and trade lifecycle tracking.
    //
    // ALGORITHM:
    // 1. Build a lookup map: for each (Account + Ticker) key, track the most
    //    recent non-blank Strategy Type seen in outputData row order.
    // 2. Walk outputData a second time. For any row whose Action is RAD or whose
    //    Description was a BOT/SOLD UPON row (now stored as BUY TO OPEN /
    //    SELL TO CLOSE with blank Strategy Type AND no Strike/Exp/CP), try to
    //    fill Strategy Type from the map IF the ticker and account match.
    // 3. Never overwrite a row that already has Strategy Type populated.
    // 4. Cluster by exact Trade Time Stamp (not just date) to avoid cross-
    //    contaminating two different same-day same-ticker trades.
    //
    // COLUMN INDICES used: actionIdx, tickerIndex, strategyTypeIndex,
    // strikeIndex, expIndex, cpIndex, accountIdx (all hoisted above the loop).
    // ──────────────────────────────────────────────────────────────────────────

    // We need the accountIdx in Import headers — hoist it if not already present.
    // (It was hoisted above the main loop as part of the standard set.)
    // Build a forward-fill strategy map keyed by "account|ticker".
    // Value = the most recent Strategy Type seen for that key.
    const strategyFillMap = {};

    // FIRST SUB-PASS: collect Strategy Types from opening/closing trade rows.
    // ---- FIXED CODE ----
    // Key includes OptionExpiration so NVDA Short IC (exp 5/26/23) never
    // collides with a NVDA Long Put at any other expiration.
    // Multi-leg spreads share the same expiration on all legs, so all legs
    // produce the same key and write the same Strategy Type — no conflict.
    const openingActions = ['BUY TO OPEN', 'SELL TO OPEN', 'BUY TO CLOSE', 'SELL TO CLOSE'];
    const acctIdxFill = importHeaders.findIndex(h => h.trim().toLowerCase() === 'account');
    const tsIdxFill = importHeaders.findIndex(h => h.trim().toLowerCase() === 'trade time stamp');
    const expIdxFill = expIndex; // already hoisted above the main loop

    function makeStrategyKey(acct, ticker, exp) {
      // Normalize expiration to a consistent string key (ms timestamp or blank)
      const expKey = (exp instanceof Date && !isNaN(exp)) ? exp.getTime().toString() : String(exp).trim();
      return acct + '|' + ticker + '|' + expKey;
    }

    for (let fi = 0; fi < outputData.length; fi++) {
      const fRow = outputData[fi];
      const fAction = (actionIdx !== -1) ? fRow[actionIdx].toString().trim().toUpperCase() : '';
      const fTicker = (tickerIndex !== -1) ? fRow[tickerIndex].toString().trim().toUpperCase() : '';
      const fStrategy = (strategyTypeIndex !== -1) ? fRow[strategyTypeIndex].toString().trim().toUpperCase() : '';
      const fAcct = (acctIdxFill !== -1) ? fRow[acctIdxFill].toString().trim().toUpperCase() : '';
      const fExp = (expIdxFill !== -1) ? fRow[expIdxFill] : '';
      if (fTicker && fStrategy && openingActions.includes(fAction)) {
        const key = makeStrategyKey(fAcct, fTicker, fExp);
        strategyFillMap[key] = fStrategy;
      }
    }

    // SECOND SUB-PASS: fill blank Strategy Type on settlement rows.
    // ---- FIXED CODE ----
    // Uses expiration-scoped key (matching the first sub-pass) so a RAD row
    // for NVDA C 330 5/26/23 maps to SHORT IC, not to a Long Put on a
    // different NVDA expiration anywhere else in the dataset.
    // For stock delivery legs (BOT/SOLD UPON, no expiration), fall back to
    // the ticker-only key so those rows still get filled.
    const settlementActions = ['RAD', 'EXP'];
    for (let si = 0; si < outputData.length; si++) {
      const sRow = outputData[si];
      const sAction = (actionIdx !== -1) ? sRow[actionIdx].toString().trim().toUpperCase() : '';
      const sTicker = (tickerIndex !== -1) ? sRow[tickerIndex].toString().trim().toUpperCase() : '';
      const sStrategy = (strategyTypeIndex !== -1) ? sRow[strategyTypeIndex].toString().trim() : '';
      const sAcct = (acctIdxFill !== -1) ? sRow[acctIdxFill].toString().trim().toUpperCase() : '';
      const sStrike = (strikeIndex !== -1) ? sRow[strikeIndex] : '';
      const sExp = (expIndex !== -1) ? sRow[expIndex] : '';
      const sCp = (cpIndex !== -1) ? sRow[cpIndex].toString().trim() : '';

      // Never overwrite a row that already has Strategy Type
      if (sStrategy) continue;
      // Never fill blank-ticker rows (JRN, EFN, DOI, etc.)
      if (!sTicker) continue;

      const isSettlementAction = settlementActions.includes(sAction);
      // Stock delivery leg: BUY TO OPEN or SELL TO CLOSE with no option fields
      const isStockDeliveryLeg = (sAction === 'BUY TO OPEN' || sAction === 'SELL TO CLOSE') &&
        !sStrike && !sExp && !sCp;
      if (!isSettlementAction && !isStockDeliveryLeg) continue;

      // Primary lookup: expiration-scoped key (handles option RAD/EXP rows)
      let fillStrategy = strategyFillMap[makeStrategyKey(sAcct, sTicker, sExp)];

      // Fallback: ticker-only key for stock delivery legs (no expiration on the row)
      // Also covers edge cases where the expiration key didn't match but ticker does.
      if (!fillStrategy) {
        // Build a ticker-only fallback from the map by scanning all keys for this acct+ticker prefix.
        // We pick the entry whose expiration is closest (or the only one if there's just one).
        const prefix = sAcct + '|' + sTicker + '|';
        const candidates = Object.keys(strategyFillMap).filter(k => k.startsWith(prefix));
        if (candidates.length === 1) {
          fillStrategy = strategyFillMap[candidates[0]];
        } else if (candidates.length > 1 && isStockDeliveryLeg) {
          // For stock delivery legs we can't match by expiration — just take the most recent
          // (last key in insertion order; JS objects preserve insertion order for string keys).
          fillStrategy = strategyFillMap[candidates[candidates.length - 1]];
        }
        // If multiple candidates and NOT a stock delivery leg, leave blank (ambiguous — safer than wrong).
      }

      if (fillStrategy) {
        outputData[si][strategyTypeIndex] = fillStrategy;
        importIssuesAdd(ctx, 'INFO', si + 4, 'Strategy Type', fillStrategy,
          `Strategy Type forward-filled (exp-scoped) → ${sAction} settlement row for ${sTicker}`);
      }
    }
    // ── END POST-LOOP PASS ────────────────────────────────────────────────────

    // === SAFE CLEAR of Import before writing ===
    var lastImportRow = importSheet.getLastRow();
    if (lastImportRow >= 4) {
      importSheet.getRange(4, 1, lastImportRow - 3, importHeaders.length).clearContent();
    }

    // Paste output
    importSheet.getRange(4, 1, finalNumRows, importHeaders.length).setValues(outputData);

    // ── CTX: finalize metrics ──────────────────────────────────────────────
    importIssuesSetMetric(ctx, 'RowsWrittenExclHeader', finalNumRows);
    importIssuesSetMetric(ctx, 'RowsSkipped', skippedCount);
    importIssuesSetMetric(ctx, 'RADRowsParsed', radParsed);
    importIssuesSetMetric(ctx, 'RADParseFailures', radFailed);
    importIssuesSetMetric(ctx, 'TickerNormalizations', tickerNorms);
    importIssuesSetMetric(ctx, 'Success', '1');
    // ────────────────────────────────────────────────────────────────────────

    // Generic post-run check for Google server-side date rendering gaps.
    checkMissingDateTimeAndAlert(importSheet, 4, 'copyMappingToImportByHeaders');

    SpreadsheetApp.getUi().alert(
      finalNumRows + ' rows copied from Schwab Mapping to Import.'
    );

  } catch (e) {
    // ── PRIORITY 1: Crash handler — log error metrics then re-throw ─────────
    // WHY: Re-throwing lets refreshAllScripts() catch it for its own alert.
    // The finally block below guarantees the flush happens either way.
    importIssuesSetMetric(ctx, 'Success', '0');
    importIssuesSetMetric(ctx, 'ErrorMessage', e.message);
    importIssuesSetMetric(ctx, 'ErrorStack', (e.stack || '').substring(0, 500));
    throw e;

  } finally {
    // ── PRIORITY 1: ALWAYS flush — even on throw ─────────────────────────
    // This was the original bug: stagingIssuesFlush(ctx) only ran at the end
    // of the try block, so any crash produced a blank Staging Issues row.
    // Moving it here guarantees you always get a log entry for this step.
    stagingIssuesFlush(ctx);
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// HELPER: Robust timestamp parser with fallback
// Placed here at the top so it's easy to find and share with other functions.
//
// PRIORITY 2 FIX (parseTradeTimeStamp): The old split(':') call silently broke
// when Trade Time was stored as "HHmm" (e.g. "0932") instead of "HH:mm"
// ("09:32").  mapSchwabImportByHeadersV3 formats Trade Time as "HHmm" with no
// colon, so any row that needed the fallback path got hours=932 and minutes=NaN,
// producing a wildly wrong or null timestamp.
// The new parser handles both "HH:mm" and "HHmm", plus an optional seconds
// component ("HH:mm:ss" or "HHmmss") so future Schwab format changes won't
// silently break anything.
// ─────────────────────────────────────────────────────────────────────────────
// =========================================================================
// HELPER: parseTradeTimeStamp
//   Builds a reliable Date object from Trade Time Stamp (preferred) or from
//   the separate Trade Date + Trade Time columns as fallback.
// =========================================================================
function parseTradeTimeStamp(tsVal, tradeDateVal, tradeTimeVal, ss) {
  let fullTimestamp = null;

  // Try Trade Time Stamp first (primary path — should always be present)
  if (tsVal) {
    fullTimestamp = (tsVal instanceof Date) ? tsVal : new Date(tsVal);
    if (isNaN(fullTimestamp.getTime())) {
      fullTimestamp = new Date(tsVal.toString().replace(/-/g, '/'));
    }
  }

  // FALLBACK: build from Trade Date + Trade Time when Trade Time Stamp is
  // missing or unparseable (fixes DT rows that arrive without a full timestamp)
  if (!fullTimestamp || isNaN(fullTimestamp.getTime())) {
    if (tradeDateVal && tradeTimeVal) {
      let baseDate = (tradeDateVal instanceof Date) ? tradeDateVal : new Date(tradeDateVal);
      if (!isNaN(baseDate.getTime())) {
        const timeStr = tradeTimeVal.toString().trim();
        let hours, minutes, seconds;

        if (timeStr.includes(':')) {
          // Format "HH:mm" or "HH:mm:ss"
          const parts = timeStr.split(':').map(Number);
          [hours, minutes, seconds = 0] = parts;
        } else if (timeStr.length >= 4) {
          // Format "HHmm" or "HHmmss" — no colon (Schwab Mapping output format)
          hours = Number(timeStr.substring(0, 2));
          minutes = Number(timeStr.substring(2, 4));
          seconds = timeStr.length >= 6 ? Number(timeStr.substring(4, 6)) : 0;
        } else {
          hours = NaN; minutes = NaN; seconds = 0;
        }

        if (!isNaN(hours) && !isNaN(minutes)) {
          fullTimestamp = new Date(
            baseDate.getFullYear(),
            baseDate.getMonth(),
            baseDate.getDate(),
            hours,
            minutes,
            seconds || 0
          );
        }
      }
    }
  }

  return fullTimestamp && !isNaN(fullTimestamp.getTime()) ? fullTimestamp : null;
}

/**
 * validateAndCleanImportToHelperV3
 * Runs right after copyMappingToImportByHeaders().
 *
 * WHY: Makes data 100% pristine AND derives clean Trade Date / Trade Time from
 * the authoritative Trade Time Stamp in BOTH Import and Helper.
 *
 * Changes in this version (5/3/2026):
 *   PRIORITY 1 — try/catch/finally wrapper so Staging Issues is ALWAYS flushed
 *                even on a mid-function crash.
 *   PRIORITY 2 — All importHeaders.indexOf() calls hoisted above the row loop.
 *                The old code re-ran ~12 header searches on every single row
 *                (45,000+ redundant ops on a full DT import).
 *   PRIORITY 2 — parseTradeTimeStamp now handles "HHmm" (no colon) as well as
 *                "HH:mm", matching the format written by mapSchwabImportByHeadersV3.
 */
// =========================================================================
// STEP 2: validateAndCleanImportToHelperV3
//   Runs immediately after copyMappingToImportByHeaders.
//   Cleans data, forces full timestamps, and writes the pristine result
//   into the "Helper" sheet that the block logic will read.
// =========================================================================
function validateAndCleanImportToHelperV3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('Import');
  const helperSheet = ss.getSheetByName('Helper');
  if (!importSheet || !helperSheet) throw new Error('Import or Helper sheet not found!');

  // ── Staging Issues CTX ─────────────────────────────────────────────────────
  const ctx = stagingIssuesStart('validateAndCleanImportToHelperV3');
  importIssuesSetMetric(ctx, 'SourceSheet', 'Import');
  importIssuesSetMetric(ctx, 'DestSheet', 'Helper');
  // ───────────────────────────────────────────────────────────────────────────

  // ── PRIORITY 1: try/catch/finally so Issues log ALWAYS gets flushed ────────
  // WHY: Without this wrapper, any mid-function throw (sheet missing, bad data,
  // regex error, Sheets API limit) exits silently and leaves ZERO log output in
  // Staging Issues for this step — the worst time to have no diagnostics.
  // The finally block guarantees stagingIssuesFlush(ctx) runs even on crash.
  try {

    // 1. Read Import data (starts at row 4)
    const importData = importSheet.getRange(
      4, 1,
      importSheet.getLastRow() - 3,
      importSheet.getLastColumn()
    ).getValues();

    importIssuesSetMetric(ctx, 'SourceRowsReadExclHeader', importData.length);
    let tickerNormCount = 0;

    // ── Raw importHeaders array (lowercase-trimmed for indexOf matching) ──────
    const importHeaders = importSheet
      .getRange(1, 1, 1, importSheet.getLastColumn())
      .getValues()[0]
      .map(h => h.trim().toLowerCase());

    // ── PRIORITY 2: Hoist ALL header index lookups ABOVE the row loop ─────────
    // WHY: importHeaders never changes between rows. Calling indexOf() inside the
    // loop ran every search on every row (12 searches × 15,000 DT rows = 180,000
    // redundant string comparisons per pipeline run). Hoisted once here = zero
    // redundancy. Also eliminates any risk of a typo on one pass vs another.
    const tsIdx = importHeaders.indexOf('trade time stamp');
    const dateIdx = importHeaders.indexOf('trade date');
    const timeIdx = importHeaders.indexOf('trade time');
    const tickerIdx = importHeaders.indexOf('ticker');
    const strategyTypeIdx = importHeaders.indexOf('strategy type');
    const tradeTypeIdx = importHeaders.indexOf('trade type');
    const signedQuantityIdx = importHeaders.indexOf('signed quantity');
    const quantityIdx = importHeaders.indexOf('quantity');
    const actionIdx = importHeaders.indexOf('action');
    const entryPriceIdx = importHeaders.indexOf('entry price');
    const optionContractIdx = importHeaders.indexOf('option contract');
    const strikeOCIdx = importHeaders.indexOf('option strike');
    const expOCIdx = importHeaders.indexOf('option expiration');
    const cpOCIdx = importHeaders.indexOf('call/put');
    const accountIdx = importHeaders.indexOf('account');
    const corpActionsIdx = importHeaders.indexOf('corporate actions');
    const accountActionsIdx = importHeaders.indexOf('account actions');
    // ─────────────────────────────────────────────────────────────────────────

    // Also need the uppercase-target column indices (used in the forEach below)
    // Hoisted here so the forEach doesn't re-search importHeaders each iteration.
    const uppercaseTextColumns = ['Account', 'Ticker', 'Action', 'Trade Type', 'Call/Put', 'Strategy Type'];
    const uppercaseColIndices = uppercaseTextColumns.map(
      colName => importHeaders.indexOf(colName.toLowerCase())
    );

    // 2. Make sure Trade Time Stamp column exists on Helper
    const helperHeaders = helperSheet.getRange(1, 1, 1, helperSheet.getLastColumn()).getValues()[0];
    let tsCol = helperHeaders.findIndex(h => h.toString().trim().toLowerCase() === 'trade time stamp') + 1;
    if (tsCol === 0) {
      helperSheet.getRange(1, helperSheet.getLastColumn() + 1).setValue('Trade Time Stamp');
      tsCol = helperSheet.getLastColumn();
    }

    let outputRows = [];
    let errors = [];
    const errorSheet = ensureValidationErrorSheet();
    const nonTradeActions = ['EFN', 'RAD', 'JRN', 'DOI', 'EXP', 'CRC', 'CDB'];

    // === CLEAR OLD ERRORS ===
    if (errorSheet.getLastRow() > 1) {
      errorSheet.getRange(2, 1, errorSheet.getLastRow() - 1, errorSheet.getLastColumn()).clearContent();
    }

    // ── TIMEZONE: resolved once here so Utilities.formatDate doesn't need to
    // call ss.getSpreadsheetTimeZone() on every row (minor but free win).
    const tz = ss.getSpreadsheetTimeZone();

    // ─────────────────────────────────────────────────────────────────────────
    // MAIN ROW LOOP
    // All header index lookups use the hoisted constants above — never indexOf
    // inside the loop body.
    // ─────────────────────────────────────────────────────────────────────────
    for (let r = 0; r < importData.length; r++) {
      const row = importData[r];
      let cleanRow = [...row];
      let errorMsg = '';

      // ── 1. TIMESTAMP: derive from authoritative Trade Time Stamp (with fallback) ──
      // Uses the hoisted tsIdx / dateIdx / timeIdx (no indexOf in the loop).
      let fullTimestamp = null;
      if (tsIdx > -1) {
        const tsVal = row[tsIdx];
        const dateVal = (dateIdx > -1) ? row[dateIdx] : null;
        const timeVal = (timeIdx > -1) ? row[timeIdx] : null;

        fullTimestamp = parseTradeTimeStamp(tsVal, dateVal, timeVal, ss);

        if (fullTimestamp) {
          cleanRow[tsIdx] = fullTimestamp;
          const derivedDate = new Date(
            fullTimestamp.getFullYear(),
            fullTimestamp.getMonth(),
            fullTimestamp.getDate()
          );
          const derivedTime = Utilities.formatDate(fullTimestamp, tz, 'HH:mm');
          if (dateIdx > -1) cleanRow[dateIdx] = derivedDate;
          if (timeIdx > -1) cleanRow[timeIdx] = derivedTime;
        } else {
          errorMsg = 'Invalid Trade Time Stamp';
        }
      }

      // ── 2. TRADE TYPE: derive here so Helper is the canonical pristine source ──
      // WHY: A blank Strategy Type must produce a blank Trade Type, not 'Option'.
      // Schwab settlement rows (stock exercise/assignment legs, EXP cash rows) arrive
      // with blank Strategy Type — defaulting them to 'Option' caused those rows to
      // fail the option field validation check (no Strike/Exp/CP on a stock row).
      // Rule: blank Ticker OR blank Strategy Type → blank Trade Type.
      // Only rows with an explicit Strategy Type keyword get a derived Trade Type.
      if (tickerIdx > -1 && strategyTypeIdx > -1 && tradeTypeIdx > -1) {
        const ticker = (row[tickerIdx] || '').toString().trim().toUpperCase();
        const strategyType = (row[strategyTypeIdx] || '').toString().trim().toUpperCase();
        if (ticker && strategyType) {
          let tradeType = 'Option';

          // A row with no Strike AND no Expiration is a stock row, full stop.
          // Catches settlement stock-delivery legs whose Strategy Type was forward-filled
          // from the parent spread (e.g. SHORT IC) which would otherwise derive 'Option'.
          const hasNoOptionFields = !cleanRow[strikeOCIdx] && !cleanRow[expOCIdx];
          if (hasNoOptionFields || strategyType.includes('STOCK')) tradeType = 'Stock';
          else if (strategyType.includes('PCS') || strategyType.includes('PDS') ||
            strategyType.includes('CCS') || strategyType.includes('CDS') ||
            strategyType.includes('BUTTERFLY') || strategyType.includes('IRON CONDOR')) tradeType = 'Spread';

          cleanRow[tradeTypeIdx] = tradeType;
        } else {
          cleanRow[tradeTypeIdx] = '';
        }
      }

      // ── 2B. CORPORATE ACTION CANONICALIZATION ──────────────────────────
      // WHY: SYMBOL CHANGE and SPLIT rows may legally survive validation even
      // with sparse upstream fields, but Helper should still classify them as
      // stock-lineage rows so Staging does not have to guess.
      const corpActionVal = (corpActionsIdx > -1 ? cleanRow[corpActionsIdx] : '')
        .toString().trim().toUpperCase();
      const cleanActionVal = (actionIdx > -1 ? cleanRow[actionIdx] : '')
        .toString().trim().toUpperCase();
      const cleanTickerVal = (tickerIdx > -1 ? cleanRow[tickerIdx] : '')
        .toString().trim().toUpperCase();

      if (
        cleanTickerVal &&
        (
          corpActionVal === 'SYMBOL CHANGE' ||
          corpActionVal === 'SPLIT' ||
          cleanActionVal === 'SYMBOL CHANGE' ||
          cleanActionVal === 'SPLIT'
        )
      ) {
        if (tradeTypeIdx > -1) cleanRow[tradeTypeIdx] = 'Stock';
        if (strategyTypeIdx > -1 && !String(cleanRow[strategyTypeIdx] || '').trim()) {
          cleanRow[strategyTypeIdx] = 'LONG STOCK';
        }
      }

      // ── 3. UPPERCASE text columns ─────────────────────────────────────────
      // Uses pre-computed uppercaseColIndices array — no inner indexOf call.
      for (let u = 0; u < uppercaseColIndices.length; u++) {
        const idx = uppercaseColIndices[u];
        if (idx > -1) {
          cleanRow[idx] = (cleanRow[idx] || '').toString().trim().toUpperCase();
        }
      }

      // ── 4. TICKER NORMALIZATION: strip Schwab alternate index prefixes ─────
      // "$SPX.X" → "SPX" | "$NDX.X" → "NDX" | "SPY" → "SPY" (no change).
      // Uses hoisted tickerIdx.
      if (tickerIdx > -1) {
        const tkrRaw = (cleanRow[tickerIdx] || '').toString().trim().toUpperCase();
        const tkrNormalized = tkrRaw.replace(/^\$/, '').replace(/\.[A-Z]+$/, '');
        if (tkrNormalized !== tkrRaw) {
          tickerNormCount++;
          cleanRow[tickerIdx] = tkrNormalized;
          // Uncomment the line below only for debugging a specific normalization issue:
          // importIssuesAdd(ctx, 'INFO', r + 4, 'Ticker', `${tkrRaw} → ${tkrNormalized}`, 'Ticker normalized from Schwab alternate index symbol');
        }
      }

      // ── 5. SIGNED QUANTITY ────────────────────────────────────────────────
      // Uses hoisted signedQuantityIdx / quantityIdx / actionIdx.
      if (signedQuantityIdx > -1 && quantityIdx > -1 && actionIdx > -1) {
        const qtyVal = Number(row[quantityIdx]);
        const actionVal = (row[actionIdx] || '').toString().toUpperCase().trim();
        cleanRow[signedQuantityIdx] = !isNaN(qtyVal)
          ? (actionVal.includes('SELL') ? -1 * qtyVal : qtyVal)
          : '';
      }

      // ── 6. QUANTITY validation ────────────────────────────────────────────
      // Uses hoisted quantityIdx.
      if (quantityIdx > -1) {
        const valNum = Number(row[quantityIdx]);
        cleanRow[quantityIdx] = !isNaN(valNum)
          ? valNum
          : (row[quantityIdx] ? '❌ INVALID QTY' : '');
      }

      // ── 7. ENTRY PRICE validation ─────────────────────────────────────────
      // Uses hoisted entryPriceIdx.
      if (entryPriceIdx > -1) {
        const valNum = Number(row[entryPriceIdx]);
        cleanRow[entryPriceIdx] = !isNaN(valNum)
          ? valNum
          : (row[entryPriceIdx] ? '❌ INVALID PRICE' : '');
      }

      // ── 8. OPTION CONTRACT: standardize to OCC format ────────────────────
      // Uses hoisted optionContractIdx / tickerIdx (reused) / strikeOCIdx /
      // expOCIdx / cpOCIdx.
      if (optionContractIdx > -1 && tickerIdx > -1 &&
        strikeOCIdx > -1 && expOCIdx > -1 && cpOCIdx > -1) {
        const ocTicker = (row[tickerIdx] || '').toString().toUpperCase().trim();
        const strike = Number(row[strikeOCIdx]);
        const exp = row[expOCIdx];
        let cp = (row[cpOCIdx] || '').toString().toUpperCase().trim();
        if (cp === 'CALL') cp = 'C';
        if (cp === 'PUT') cp = 'P';

        let expDatePart = '';
        if (exp) {
          if (exp instanceof Date) {
            const yy = String(exp.getFullYear()).slice(-2);
            const mm = String(exp.getMonth() + 1).padStart(2, '0');
            const dd = String(exp.getDate()).padStart(2, '0');
            expDatePart = yy + mm + dd;
          } else {
            const expStr = exp.toString().trim();
            const dateMatch = expStr.match(/(\d{2})[\/\-](\d{2})[\/\-](\d{4})/);
            if (dateMatch) {
              expDatePart = dateMatch[3].slice(-2) + dateMatch[1] + dateMatch[2];
            } else if (expStr.length === 8 && /^\d{8}$/.test(expStr)) {
              expDatePart = expStr.slice(2);
            }
          }
        }

        const strikePart = !isNaN(strike)
          ? String(Math.round(strike * 1000)).padStart(8, '0')
          : '';

        if (ocTicker && expDatePart && (cp === 'C' || cp === 'P') && strikePart) {
          cleanRow[optionContractIdx] =
            ocTicker.replace(/[^A-Z]/g, '') + expDatePart + cp + strikePart;
        }
      }

      // ── 9. VALIDATION RULES ───────────────────────────────────────────────
      // Uses hoisted actionIdx / accountIdx / tickerIdx / quantityIdx /
      // tradeTypeIdx / strikeOCIdx / expOCIdx / cpOCIdx /
      // corpActionsIdx / accountActionsIdx.
      const actionRaw = (row[actionIdx] || '').toString().trim().toUpperCase();
      const isTradeRow = ['BUY TO OPEN', 'SELL TO OPEN', 'BUY TO CLOSE', 'SELL TO CLOSE']
        .includes(actionRaw);

      const acct = (accountIdx > -1 ? row[accountIdx] : '')
        .toString().trim().toUpperCase();
      if (acct !== 'DT' && acct !== 'LT') errorMsg = 'Account must be DT or LT';

      const tickerVal = (tickerIdx > -1 ? row[tickerIdx] : '').toString().trim().toUpperCase();
      if (isTradeRow && !tickerVal) errorMsg = 'Ticker required for trade actions';

      const qtyCheck = Number(quantityIdx > -1 ? row[quantityIdx] : '');
      if (isTradeRow && isNaN(qtyCheck)) errorMsg = 'Quantity must be a number';

      // --- AFTER ---
      // WHY: Read tradeTypeCheck from cleanRow (not raw row) because Trade Type
      // is derived in Step 2 above and written to cleanRow — the raw row[tradeTypeIdx]
      // may be blank (first run) or a prior-run value (subsequent runs), both of
      // which cause incorrect validation behaviour. Also read strike/exp/cp from
      // cleanRow so Step 8 normalizations are visible here.
      // WHY: Wrap in !nonTradeActions guard so RAD, JRN, EFN etc. never reach the
      // option field check — they legitimately have no Strike/Exp/C-P requirement.
      const tradeTypeCheck = (tradeTypeIdx > -1 ? cleanRow[tradeTypeIdx] : '')
        .toString().trim().toUpperCase();
      if (!nonTradeActions.includes(actionRaw) &&
        (tradeTypeCheck.includes('OPTION') || tradeTypeCheck.includes('SPREAD'))) {
        if (!(strikeOCIdx > -1 && cleanRow[strikeOCIdx]) ||
          !(expOCIdx > -1 && cleanRow[expOCIdx]) ||
          !(cpOCIdx > -1 && cleanRow[cpOCIdx])) {
          errorMsg = 'Options need Strike, Expiration, Call/Put';
        }
      }

      // Corporate / ledger rows clear any validation error — they don't need
      // Ticker, Quantity, or option fields.
      const corpAction = (corpActionsIdx > -1) ? row[corpActionsIdx] : '';
      const accountAction = (accountActionsIdx > -1) ? row[accountActionsIdx] : '';
      if (corpAction || accountAction || nonTradeActions.includes(actionRaw)) {
        errorMsg = '';
      }

      if (errorMsg) {
        errors.push([r + 4, 'Action/Ticker', errorMsg, 'Fix in Import sheet and re-run']);
        importIssuesAdd(ctx, 'ERROR', r + 4, 'Action/Ticker', errorMsg,
          'Row excluded from Helper — fix in Import and re-run');
      } else {
        outputRows.push(cleanRow);
      }

    } // end main row loop

    // ── SAFE CLEAR both sheets before writing ─────────────────────────────────
    // WHY: If the previous run wrote more rows than this run, stale rows at the
    // bottom would silently flow into Staging on the next block-logic pass.
    const helperLastRow = helperSheet.getLastRow();
    if (helperLastRow >= 4) {
      helperSheet.getRange(4, 1, helperLastRow - 3, helperSheet.getLastColumn()).clearContent();
    }
    const importLastRow = importSheet.getLastRow();
    if (importLastRow >= 4) {
      importSheet.getRange(4, 1, importLastRow - 3, importSheet.getLastColumn()).clearContent();
    }

    // Write clean rows to both Helper and Import
    if (outputRows.length > 0) {
      helperSheet.getRange(4, 1, outputRows.length, importData[0].length).setValues(outputRows);
      importSheet.getRange(4, 1, outputRows.length, importData[0].length).setValues(outputRows);
    }

    // === FORCE CORRECT DISPLAY FORMATS ON BOTH SHEETS ===
    if (outputRows.length > 0) {
      const tsColNum = tsIdx + 1;
      const timeColNum = timeIdx + 1;
      importSheet.getRange(4, tsColNum, outputRows.length, 1).setNumberFormat('M/d/yyyy HH:mm');
      importSheet.getRange(4, timeColNum, outputRows.length, 1).setNumberFormat('HH:mm');
      helperSheet.getRange(4, tsCol, outputRows.length, 1).setNumberFormat('M/d/yyyy HH:mm');
      helperSheet.getRange(4, timeColNum, outputRows.length, 1).setNumberFormat('HH:mm');
    }

    // Write validation errors to the Validation Errors sheet if any
    if (errors.length > 0) {
      errorSheet.getRange(errorSheet.getLastRow() + 1, 1, errors.length, 4).setValues(errors);
    }

    // ── CTX: finalize metrics ───────────────────────────────────────────────
    importIssuesSetMetric(ctx, 'RowsWrittenExclHeader', outputRows.length);
    importIssuesSetMetric(ctx, 'ValidationErrors', errors.length);
    importIssuesSetMetric(ctx, 'TickerNormalizations', tickerNormCount);
    importIssuesSetMetric(ctx, 'Success', '1');
    // ────────────────────────────────────────────────────────────────────────

    // Generic post-run check for Google server-side date rendering gaps.
    checkMissingDateTimeAndAlert(helperSheet, 4, 'validateAndCleanImportToHelperV3');

    if (errors.length > 0) {
      SpreadsheetApp.getUi().alert(
        '⚠️ Validation found ' + errors.length + ' errors — check Validation Errors sheet!'
      );
    } else {
      SpreadsheetApp.getUi().alert(
        '✅ All data pristine — Import and Helper now have FULL derivations and correct display!'
      );
    }

  } catch (e) {
    // ── PRIORITY 1: Crash handler — log error metrics then re-throw ──────────
    importIssuesSetMetric(ctx, 'Success', '0');
    importIssuesSetMetric(ctx, 'ErrorMessage', e.message);
    importIssuesSetMetric(ctx, 'ErrorStack', (e.stack || '').substring(0, 500));
    throw e;

  } finally {
    // ── PRIORITY 1: ALWAYS flush — even on throw ──────────────────────────
    // The old placement of stagingIssuesFlush(ctx) was at the end of the normal
    // flow only. Moving it here guarantees a log entry exists for every run,
    // successful or not.
    stagingIssuesFlush(ctx);
  }
}

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
function resolveLiveSpreadGroupId(acct, ticker, expStr, strat, strike, spreadRangeMap, blocks) {
  const rangeKey = `${acct}|${ticker}|${expStr}|${strat}`;
  const rangeGroups = (spreadRangeMap[rangeKey] || []);
  const candidates = rangeGroups.filter(g => strike >= g.min && strike <= g.max);

  if (candidates.length === 0) return '';
  if (candidates.length === 1) return candidates[0].groupId;

  // Multiple candidates — prefer the one whose block currently has unit > 0.
  for (const g of candidates) {
    const blockKey = `${acct}|${g.groupId}`;
    if ((blocks[blockKey] || {}).unit > 0) return g.groupId;
  }

  // Fallback: first candidate (original behavior; safe for normal trades).
  return candidates[0].groupId;
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
  const helperSheet = ss.getSheetByName('Helper');
  const stagingSheet = ss.getSheetByName('Staging');
  if (!helperSheet || !stagingSheet) throw new Error('Helper or Staging sheet not found!');

  // ── Staging Issues CTX ─────────────────────────────────────────────────────
  const ctx = stagingIssuesStart('populateStagingWithBlockLogicV3');
  importIssuesSetMetric(ctx, 'SourceSheet', 'Helper');
  importIssuesSetMetric(ctx, 'DestSheet', 'Staging');
  // ───────────────────────────────────────────────────────────────────────────

  const STRATEGY_ABBREV = {
    'SHORT IC': 'SIC',
    'LONG IC': 'LIC',
    'SHORT PCS': 'PCS',
    'SHORT PDS': 'PDS',
    'SHORT CCS': 'CCS',
    'SHORT CDS': 'CDS',
    'LONG BUTTERFLY': 'LBF',
    'SHORT BUTTERFLY': 'SBF',
    'LONG PUT': 'LP',
    'SHORT PUT': 'SP',
    'LONG CALL': 'LC',
    'SHORT CALL': 'SC',
    'LONG STOCK': 'LST',
    'SHORT STOCK': 'SST',
  };
  function getStratAbbrev(strategyType) {
    return STRATEGY_ABBREV[strategyType.trim().toUpperCase()] || 'OTH';
  }
  

  function parseSymbolChangeFromNotes(row, colMap) {
    if (colMap['notes'] === undefined) return null;

    const notes = String(row[colMap['notes'] - 1] || '').trim();
    if (!notes) return null;

    const out = {
      fromRaw: '',
      toRaw: '',
      fromResolved: '',
      toResolved: '',
      rawText: notes
    };

    notes.split('|').forEach(part => {
      const seg = String(part || '').trim();
      const eq = seg.indexOf('=');
      if (eq === -1) return;

      const key = seg.substring(0, eq).trim().toUpperCase();
      const val = seg.substring(eq + 1).trim().toUpperCase();

      if (key === 'FROM') out.fromRaw = val;
      if (key === 'TO') out.toRaw = val;
      if (key === 'FROM_RESOLVED') out.fromResolved = val;
      if (key === 'TO_RESOLVED') out.toResolved = val;
    });

    return (out.fromRaw || out.toRaw || out.fromResolved || out.toResolved) ? out : null;
  }


  function parseSplitRatioFromRow(row, colMap) {
    const notesIdx = colMap['notes'] ? colMap['notes'] - 1 : -1;
    const text = notesIdx > -1 ? String(row[notesIdx] || '').trim().toUpperCase() : '';

    if (!text || text.indexOf('SPLIT') === -1) return null;

    const ratioMatch = text.match(/\b(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)/);
    const preMatch = text.match(/\bPRE\s*=\s*([-+]?\d+(?:\.\d+)?)/);
    const postMatch = text.match(/\bPOST\s*=\s*([-+]?\d+(?:\.\d+)?)/);

    const numerator = ratioMatch ? Number(ratioMatch[1]) : NaN;
    const denominator = ratioMatch ? Number(ratioMatch[2]) : NaN;
    const preQty = preMatch ? Math.abs(Number(preMatch[1])) : NaN;
    const postQty = postMatch ? Math.abs(Number(postMatch[1])) : NaN;
    const splitType = text.includes('REVERSE') ? 'REVERSE' : (text.includes('FORWARD') ? 'FORWARD' : '');

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
      text: text
    };
  }

  function computeExpectedPostSplitQty_(preQty, splitInfo) {
    if (
      !splitInfo ||
      !isFinite(preQty) ||
      !isFinite(splitInfo.numerator) ||
      !isFinite(splitInfo.denominator) ||
      splitInfo.denominator === 0
    ) return NaN;

    const rawPost = Number(preQty) * (Number(splitInfo.numerator) / Number(splitInfo.denominator));
    const isReverse = splitInfo.splitType === 'REVERSE' || splitInfo.numerator < splitInfo.denominator;

    return isReverse
      ? Math.floor(rawPost + 1e-9)
      : Math.round(rawPost * 1e8) / 1e8;
  }

  function isMatchingSplitRowForBlock_(runningBeforeSplit, splitInfo) {
    if (!splitInfo) return true;

    if (isFinite(splitInfo.preQty) && Math.abs(splitInfo.preQty - runningBeforeSplit) < 1e-8) {
      return true;
    }

    if (isFinite(splitInfo.postQty)) {
      const expectedPost = computeExpectedPostSplitQty_(runningBeforeSplit, splitInfo);
      if (isFinite(expectedPost) && Math.abs(splitInfo.postQty - expectedPost) < 1e-8) {
        return true;
      }
    }

    return false;
  }

  function findRenameSourceStockBlockForSplit_(acct, targetTicker, splitInfo, blocks) {
    if (!splitInfo || !isFinite(splitInfo.postQty)) return null;

    let match = null;

    Object.keys(blocks).forEach(function (blockKey) {
      if (match) return;

      const parts = blockKey.split('|');
      if (parts.length !== 2) return; // stock block keys are acct|ticker
      if (parts[0] !== acct) return;

      const sourceTicker = parts[1];
      if (!sourceTicker || sourceTicker === targetTicker) return;

      const block = blocks[blockKey];
      if (!block || !block.positionId) return;

      const preQty = Number(block.runningQty || 0);
      if (preQty <= 0) return;

      const expectedPost = computeExpectedPostSplitQty_(preQty, splitInfo);
      if (isFinite(expectedPost) && Math.abs(expectedPost - splitInfo.postQty) < 1e-8) {
        match = {
          sourceKey: blockKey,
          sourceTicker: sourceTicker,
          block: block,
          preQty: preQty,
          expectedPost: expectedPost
        };
      }
    });

    return match;
  }
  
  function isCusipLikeSymbolValue_(val) {
  const v = String(val || '').trim().toUpperCase();
  return !!v && /^[A-Z0-9]{9}$/.test(v) && /\d/.test(v);
}

function resolveSymbolChangeTargetTicker_(fromTicker, toRaw, toResolved, rowTicker) {
  const from = String(fromTicker || '').trim().toUpperCase();
  const rowTkr = String(rowTicker || '').trim().toUpperCase();

  const candidates = [
    String(toResolved || '').trim().toUpperCase(),
    String(toRaw || '').trim().toUpperCase(),
    rowTkr
  ];

  for (let i = 0; i < candidates.length; i++) {
    const candidate = candidates[i];
    if (!candidate) continue;

    // If the row itself already carries a real new ticker, do not let a
    // fallback candidate pin us back to the old symbol.
    if (candidate === from && rowTkr && rowTkr !== from && !isCusipLikeSymbolValue_(rowTkr)) {
      continue;
    }

    if (!isCusipLikeSymbolValue_(candidate)) return candidate;
  }

  return from;
}


  function parseSymbolChangeFromRow(row, colMap) {
    const notes = colMap['notes'] !== undefined ? String(row[colMap['notes'] - 1]).trim() : '';
    const corpActions = colMap['corporate actions'] !== undefined
      ? String(row[colMap['corporate actions'] - 1]).trim().toUpperCase()
      : '';
    const action = colMap['action'] !== undefined
      ? String(row[colMap['action'] - 1]).trim().toUpperCase()
      : '';

    const isRenameRow =
      action === 'SYMBOL CHANGE' ||
      corpActions === 'SYMBOL CHANGE' ||
      corpActions === 'Symbol Change'.toUpperCase();

    if (!isRenameRow) return null;
    if (!notes) return null;

    // Expected Notes format from Phase 2:
    // FROM=ENCUF | TO=EU | FROM_RESOLVED=ENCUF | TO_RESOLVED=EU | RAW=Symbol Change from ENCUF to EU
    function pull(label) {
      const m = notes.match(new RegExp(label + '=([^|]+)', 'i'));
      return m ? String(m[1]).trim().toUpperCase() : '';
    }

    const fromRaw = pull('FROM');
    const toRaw = pull('TO');
    const fromResolved = pull('FROM_RESOLVED');
    const toResolved = pull('TO_RESOLVED');

    return {
      fromRaw: fromRaw,
      toRaw: toRaw,
      fromResolved: fromResolved,
      toResolved: toResolved
    };
  }

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
  // ─────────────────────────────────────────────────────────────────────────────
  function resolveLiveSpreadGroupId(acct, ticker, expStr, strat, strike, spreadRangeMap, blocks) {
    const rangeKey = `${acct}|${ticker}|${expStr}|${strat}`;
    const rangeGroups = (spreadRangeMap[rangeKey] || []);
    const candidates = rangeGroups.filter(g => strike >= g.min && strike <= g.max);

    if (candidates.length === 0) return '';
    if (candidates.length === 1) return candidates[0].groupId;

    // Multiple candidates — prefer the one whose block currently has unit > 0.
    for (const g of candidates) {
      const blockKey = `${acct}|${g.groupId}`;
      if ((blocks[blockKey] || {}).unit > 0) return g.groupId;
    }

    // Fallback: first candidate (original behavior; safe for normal trades).
    return candidates[0].groupId;
  }

  // ── PRIORITY 1: try/catch/finally so Issues log ALWAYS gets flushed ────────
  try {

    const helperData = helperSheet.getDataRange().getValues();
    if (helperData.length < 4) {
      SpreadsheetApp.getUi().alert('No data in Helper to process.');
      return;
    }

    const colMap = {};
    helperData[0].forEach((h, i) => {
      if (typeof h === 'string' && h.trim()) colMap[h.trim().toLowerCase()] = i + 1;
    });

    let data = helperData.slice(3).map(row => row.slice());

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
    const tsIdx = colMap['trade time stamp'] - 1;
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
      const acctA = (a[colMap['account'] - 1] || '').toString().toUpperCase();
      const acctB = (b[colMap['account'] - 1] || '').toString().toUpperCase();
      if (acctA !== acctB) return acctA.localeCompare(acctB);

      const tsA = a[tsIdx] instanceof Date ? a[tsIdx].getTime() : 0;
      const tsB = b[tsIdx] instanceof Date ? b[tsIdx].getTime() : 0;
      if (tsA !== tsB) return tsA - tsB;

      // Tie-breaker: lower strike first within the same timestamp.
      const strikeA = Number(a[colMap['option strike'] - 1]) || 0;
      const strikeB = Number(b[colMap['option strike'] - 1]) || 0;
      return strikeA - strikeB;
    });
    // ─────────────────────────────────────────────────────────────────────────
    // Step 3: Assign Spread Group ID — THREE sub-passes.
    // ─────────────────────────────────────────────────────────────────────────

    const getTsMinute = (tsVal) => {
      if (!(tsVal instanceof Date) || isNaN(tsVal.getTime())) return 'NO-TS';
      return Utilities.formatDate(tsVal, tz, 'yyyy-MM-dd HH:mm');
    };

    // ----- Sub-pass A: Collect opening spread legs and their strikes -----
    let openGroupStrikesMap = {};
    let openGroupStrategyMap = {};
    let openGroupFirstRowMap = {};

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const actionRaw = (row[colMap['action'] - 1] || '').toString().trim().toUpperCase();
      const strat = (row[colMap['strategy type'] - 1] || '').toString().toUpperCase();

      const isSpreadOpen = (
        (actionRaw === 'BUY TO OPEN' || actionRaw === 'SELL TO OPEN') &&
        (strat.includes('PCS') || strat.includes('PDS') ||
          strat.includes('CCS') || strat.includes('CDS') ||
          strat.includes('BUTTERFLY') || strat.includes('IRON CONDOR'))
      );
      if (!isSpreadOpen) continue;

      const acct = (row[colMap['account'] - 1] || '').toString().toUpperCase();
      const ticker = (row[colMap['ticker'] - 1] || '').toString().toUpperCase();
      const exp = row[colMap['option expiration'] - 1];
      const expStr = exp instanceof Date
        ? Utilities.formatDate(exp, tz, 'yyyy-MM-dd')
        : (exp || '').toString();
      const strike = Number(row[colMap['option strike'] - 1]) || 0;
      const tsMin = getTsMinute(row[colMap['trade time stamp'] - 1]);

      const groupKey = `${acct}|${ticker}|${expStr}|${strat}|${tsMin}`;

      if (!openGroupStrikesMap[groupKey]) {
        openGroupStrikesMap[groupKey] = [];
        openGroupFirstRowMap[groupKey] = i;
      }
      if (!openGroupStrategyMap[groupKey]) openGroupStrategyMap[groupKey] = strat;
      openGroupStrikesMap[groupKey].push(strike);
    }

    // ----- Sub-pass B: Build spreadKeyMap and spreadRangeMap -----
    let spreadKeyMap = {};
    let spreadRangeMap = {};

    for (const [groupKey, strikes] of Object.entries(openGroupStrikesMap)) {
      const parts = groupKey.split('|');
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
      spreadRangeMap[rangeKey].push({ min: strikeMin, max: strikeMax, groupId: spreadGroupId });
    }

    // ----- Sub-pass C: Stamp Spread Group IDs -----
    // Opening legs only — closing legs are intentionally deferred to Step 4.
    const SPREAD_STRAT_TERMS = ['PCS', 'PDS', 'CCS', 'CDS', 'BUTTERFLY', 'IRON CONDOR'];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const actionRaw = (row[colMap['action'] - 1] || '').toString().trim().toUpperCase();
      const strat = (row[colMap['strategy type'] - 1] || '').toString().toUpperCase();

      const isSpreadStrategy = SPREAD_STRAT_TERMS.some(t => strat.includes(t));
      const isRadSpreadExpiration = (
        actionRaw === 'RAD' &&
        isSpreadStrategy &&
        row[colMap['option strike'] - 1] &&
        row[colMap['option expiration'] - 1]
      );

      if (!isSpreadStrategy && !isRadSpreadExpiration) continue;

      const acct = (row[colMap['account'] - 1] || '').toString().toUpperCase();
      const ticker = (row[colMap['ticker'] - 1] || '').toString().toUpperCase();
      const exp = row[colMap['option expiration'] - 1];
      const expStr = exp instanceof Date
        ? Utilities.formatDate(exp, tz, 'yyyy-MM-dd')
        : (exp || '').toString();

      if (actionRaw === 'BUY TO OPEN' || actionRaw === 'SELL TO OPEN') {
        // Opening legs: stamp using the timestamp-based group key.
        const tsMin = getTsMinute(row[colMap['trade time stamp'] - 1]);
        const groupKey = `${acct}|${ticker}|${expStr}|${strat}|${tsMin}`;
        if (spreadKeyMap[groupKey]) {
          row[colMap['spread group id'] - 1] = spreadKeyMap[groupKey];
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
      const acct = (row[colMap['account'] - 1] || '').toString().toUpperCase();
      const ticker = (row[colMap['ticker'] - 1] || '').toString().toUpperCase();
      const actionRaw = row[colMap['action'] - 1].toString().trim();
      const action = actionRaw.toUpperCase();

      // Skip pure ledger rows that have no Ticker.
      // IMPORTANT:
      // SYMBOL CHANGE rows are NOT skipped here even if they are JRN-origin rows,
      // because they need to transfer the live stock block from old ticker -> new ticker.
      if (['EFN', 'JRN', 'DOI', 'CRC', 'CDB'].includes(action) && action !== 'SYMBOL CHANGE' && !ticker) {
        row[colMap['position id'] - 1] = '';
        row[colMap['trade group id'] - 1] = '';
        continue;
      }

      // SPECIAL HANDLING FOR SYMBOL CHANGE rows
// WHY:
// A stock symbol rename should keep one continuous stock block.
// If the row itself already carries the real post-change ticker, keep that
// ticker on the row while inheriting the existing stock block identity.
if (action === 'SYMBOL CHANGE') {
  const scInfo = parseSymbolChangeFromRow(row, colMap);

  if (!scInfo) {
    importIssuesAdd(
      ctx,
      'WARN',
      dataStartRow + i,
      'Notes',
      '',
      'SYMBOL CHANGE row has no parsable FROM/TO info in Notes. Row left unlinked.'
    );
    if (colMap['position id'] !== undefined) row[colMap['position id'] - 1] = '';
    if (colMap['trade group id'] !== undefined) row[colMap['trade group id'] - 1] = '';
    continue;
  }

  const fromTicker = String(scInfo.fromResolved || scInfo.fromRaw || '').trim().toUpperCase();
  const toRaw = String(scInfo.toRaw || '').trim().toUpperCase();
  const toResolved = String(scInfo.toResolved || '').trim().toUpperCase();
  const rowTicker = String(row[colMap['ticker'] - 1] || '').trim().toUpperCase();

  const targetTicker = resolveSymbolChangeTargetTicker_(fromTicker, toRaw, toResolved, rowTicker);
  const displayTickerCandidate = rowTicker || toResolved || toRaw || fromTicker;
  const outputTicker = !isCusipLikeSymbolValue_(displayTickerCandidate)
    ? displayTickerCandidate
    : (targetTicker || fromTicker);

  if (!fromTicker || !outputTicker) {
    importIssuesAdd(
      ctx,
      'WARN',
      dataStartRow + i,
      'Ticker',
      'FROM=' + fromTicker + ' | ROW=' + rowTicker + ' | TO=' + toResolved,
      'SYMBOL CHANGE row missing usable old/new ticker. Row left unlinked.'
    );
    if (colMap['position id'] !== undefined) row[colMap['position id'] - 1] = '';
    if (colMap['trade group id'] !== undefined) row[colMap['trade group id'] - 1] = '';
    continue;
  }

  const oldKey = `${acct}|${fromTicker}`;
  const oldBlock = blocks[oldKey];

  if (!oldBlock || !oldBlock.positionId) {
    importIssuesAdd(
      ctx,
      'WARN',
      dataStartRow + i,
      'Ticker',
      fromTicker + ' -> ' + outputTicker,
      'SYMBOL CHANGE row found but no active source stock block exists for old ticker. Row left unlinked.'
    );
    if (colMap['position id'] !== undefined) row[colMap['position id'] - 1] = '';
    if (colMap['trade group id'] !== undefined) row[colMap['trade group id'] - 1] = '';
    continue;
  }

  // No usable destination yet: keep the existing source block identity,
  // but do NOT force the row ticker back to the old symbol if the row already
  // has a real post-change ticker.
  if (!targetTicker || targetTicker === fromTicker) {
    if (colMap['ticker'] !== undefined) row[colMap['ticker'] - 1] = outputTicker;
    if (colMap['trade group id'] !== undefined) row[colMap['trade group id'] - 1] = oldBlock.tradeGroupId || '';
    if (colMap['position id'] !== undefined) row[colMap['position id'] - 1] = oldBlock.positionId || '';
    if (colMap['block number'] !== undefined) row[colMap['block number'] - 1] = oldBlock.block || '';
    if (colMap['block start flag'] !== undefined) row[colMap['block start flag'] - 1] = 0;
    if (colMap['block close flag/p&l'] !== undefined) row[colMap['block close flag/p&l'] - 1] = 0;
    if (colMap['running position quantity'] !== undefined) {
      row[colMap['running position quantity'] - 1] = Number(oldBlock.runningQty || 0);
    }
    if (colMap['trade status'] !== undefined) {
      row[colMap['trade status'] - 1] = Number(oldBlock.runningQty || 0) === 0 ? 'Closed' : 'Open';
    }

    importIssuesAdd(
      ctx,
      'INFO',
      dataStartRow + i,
      'Ticker',
      fromTicker + ' -> ' + outputTicker,
      'SYMBOL CHANGE kept the existing stock block identity without opening a new block because no separate destination ticker was yet usable.'
    );
    continue;
  }

  const newKey = `${acct}|${targetTicker}`;
  const newBlockExisting = blocks[newKey];

  if (newBlockExisting && newBlockExisting.positionId && Number(newBlockExisting.runningQty || 0) !== 0) {
    importIssuesAdd(
      ctx,
      'WARN',
      dataStartRow + i,
      'Ticker',
      fromTicker + ' -> ' + targetTicker,
      'SYMBOL CHANGE destination ticker already has an active block. Source block was NOT merged automatically.'
    );
    if (colMap['position id'] !== undefined) row[colMap['position id'] - 1] = '';
    if (colMap['trade group id'] !== undefined) row[colMap['trade group id'] - 1] = '';
    continue;
  }

  const movedBlock = {
    unit: Number(oldBlock.unit || 0),
    block: Number(oldBlock.block || 1),
    runningQty: Number(oldBlock.runningQty || 0),
    pnl: Number(oldBlock.pnl || 0),
    entryCost: Number(oldBlock.entryCost || 0),
    openTs: oldBlock.openTs || null,
    positionId: oldBlock.positionId || '',
    tradeGroupId: oldBlock.tradeGroupId || '',
    strategyType: oldBlock.strategyType || 'LONG STOCK'
  };

  blocks[newKey] = movedBlock;
  delete blocks[oldKey];

  if (colMap['ticker'] !== undefined) row[colMap['ticker'] - 1] = outputTicker;
  if (colMap['trade group id'] !== undefined) row[colMap['trade group id'] - 1] = movedBlock.tradeGroupId || '';
  if (colMap['position id'] !== undefined) row[colMap['position id'] - 1] = movedBlock.positionId || '';
  if (colMap['block number'] !== undefined) row[colMap['block number'] - 1] = movedBlock.block || '';
  if (colMap['block start flag'] !== undefined) row[colMap['block start flag'] - 1] = 0;
  if (colMap['block close flag/p&l'] !== undefined) row[colMap['block close flag/p&l'] - 1] = 0;
  if (colMap['running position quantity'] !== undefined) {
    row[colMap['running position quantity'] - 1] = Number(movedBlock.runningQty || 0);
  }
  if (colMap['trade status'] !== undefined) {
    row[colMap['trade status'] - 1] = Number(movedBlock.runningQty || 0) === 0 ? 'Closed' : 'Open';
  }

  importIssuesAdd(
    ctx,
    'INFO',
    dataStartRow + i,
    'Ticker',
    fromTicker + ' -> ' + targetTicker + ' | ROW=' + outputTicker,
    'SYMBOL CHANGE transferred the active stock block to the destination ticker while preserving the row ticker for display and analytics.'
  );
  continue;
}
// =====================================================================
// END SYMBOL CHANGE HANDLER
// =======================================================================

      // NOTE: `spreadId` is `let` — the live resolution block below may reassign it.
      let spreadId = (row[colMap['spread group id'] - 1] || '').toString();
      let qty = Number(row[colMap['quantity'] - 1]) || 0;
      const strategyType = (row[colMap['strategy type'] - 1] || '').toString().toUpperCase();

      let tradeType = (row[colMap['trade type'] - 1] || '').toString().toUpperCase().trim();
      if (!tradeType && ticker) tradeType = 'OPTION';

      // === LIVE SPREAD GROUP ID RESOLUTION FOR CLOSING/RAD LEGS ================
      // WHY: Sub-pass C intentionally left closing leg Spread Group IDs blank.
      // We fill them here using resolveLiveSpreadGroupId(), which checks the live
      // block unit state to disambiguate boundary-strike collisions between
      // sequential spreads that share a strike (e.g. 300/310 then 310/315 PDS).
      // Opening legs already have their Spread Group ID from Sub-pass C — only
      // blank-spreadId rows with a spread strategy need resolution here.
      const isSpreadStrategy_ = SPREAD_STRAT_TERMS.some(t => strategyType.includes(t));
      const isClosingOrRAD = (action.includes('TO CLOSE') || action === 'RAD');
      if (!spreadId && isSpreadStrategy_ && isClosingOrRAD && ticker) {
        const exp_ = row[colMap['option expiration'] - 1];
        const expStr_ = exp_ instanceof Date
          ? Utilities.formatDate(exp_, tz, 'yyyy-MM-dd')
          : (exp_ || '').toString();
        const strike_ = Number(row[colMap['option strike'] - 1]) || 0;
        if (expStr_ && strike_) {
          const resolved = resolveLiveSpreadGroupId(
            acct, ticker, expStr_, strategyType, strike_, spreadRangeMap, blocks
          );
          if (resolved) {
            row[colMap['spread group id'] - 1] = resolved;
            spreadId = resolved; // local var — used immediately in the grouping key below
          } else {
            ctxMissingSpreadGroup++;
            importIssuesAdd(ctx, 'WARN', dataStartRow + i, 'Spread Group ID',
              `${acct}|${ticker}|${expStr_}|${strike_}`,
              'No live spread group found for this closing/RAD row. ' +
              'Possible cause: no matching open block at this timestamp. ' +
              'Check that the opening trade exists and processed before this row.');
          }
        }
      }
      // =========================================================================

      // === GROUPING KEY — Spread Group ID takes priority for ALL rows including RAD ===
      let key;
      if (spreadId) {
        key = `${acct}|${spreadId}`;
        tradeType = 'SPREAD';
        row[colMap['trade type'] - 1] = 'SPREAD';
      } else if (tradeType === 'OPTION') {
        const exp = row[colMap['option expiration'] - 1];
        const expStr = exp instanceof Date
          ? Utilities.formatDate(exp, tz, 'yyyy-MM-dd')
          : (exp || '').toString();
        const strike = Number(row[colMap['option strike'] - 1]) || 0;
        const cp = (row[colMap['call/put'] - 1] || '').toString().toUpperCase()
          .replace('CALL', 'C').replace('PUT', 'P');
        key = `${acct}|${ticker}|${expStr}|${strike}|${cp}`;
      } else {
        key = `${acct}|${ticker}`;
      }

      if (!blocks[key]) blocks[key] = {
        unit: 0,
        block: 1,
        runningQty: 0,
        pnl: 0,
        entryCost: 0,
        openTs: null,
        positionId: '',
        strategyType: ''
      };

      // === ROBUST DELTA ===
      let delta = 0;
      if (action.includes('SELL TO OPEN') || action.includes('BUY TO OPEN')) delta = 1;
      if (action.includes('BUY TO CLOSE') || action.includes('SELL TO CLOSE')) delta = -1;

      // === SPECIAL HANDLING FOR RAD "Opt Expired" rows ===
      if (action === 'RAD' && ticker) {
        delta = -1;
        qty = Math.abs(qty);

        row[colMap['entry price'] - 1] = 0;
        row[colMap['total cost'] - 1] = 0;

        row[colMap['strategy type'] - 1] = blocks[key].strategyType || '';

        let cpRad = (row[colMap['call/put'] - 1] || '').toString().toUpperCase();
        if (cpRad === 'CALL') row[colMap['call/put'] - 1] = 'C';
        if (cpRad === 'PUT') row[colMap['call/put'] - 1] = 'P';

        row[colMap['closing date'] - 1] = row[colMap['trade date'] - 1];
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
          const renameSource = findRenameSourceStockBlockForSplit_(acct, ticker, splitInfo, blocks);
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
              "Matched preQty=" + renameSource.preQty +
              ", expectedPost=" + renameSource.expectedPost + "."
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
            "SPLIT row for " + ticker + " has no active stock block in memory. " +
            "Split may have occurred before any position was opened, or ticker spelling differs."
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
            "Duplicate/unmatched SPLIT row ignored. It does not match the live pre/post quantity for this block."
          );
          continue;
        }

        let splitDelta = Number(qty || 0);

        if (splitInfo && isFinite(splitInfo.numerator) && isFinite(splitInfo.denominator) && splitInfo.denominator > 0) {
          const expectedPost = computeExpectedPostSplitQty_(runningBeforeSplit, splitInfo);
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
            "SPLIT row for " + ticker + " has no usable split delta. " +
            "Check Notes for FORWARD/REVERSE SPLIT ratio text."
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
          row[colMap["trade group id"] - 1] = existingBlock.tradeGroupId || row[colMap["trade group id"] - 1] || "";
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
          row[colMap["trade status"] - 1] = existingBlock.runningQty === 0 ? "Closed" : "Open";
        }

        continue;
      }
      // END SPLIT HANDLER

      const prevUnit = Number(blocks[key].unit || 0);
      const prevRunningQty = Number(blocks[key].runningQty || 0);

      blocks[key].unit += delta * qty;
      blocks[key].runningQty += delta * qty;

      const newUnit = Number(blocks[key].unit || 0);
      const newRunningQty = Number(blocks[key].runningQty || 0);
      const curBlock = blocks[key].block;

      let blkStart = 0;
      let blkClose = 0;

      // STOCK positions must be split-aware.
      // Open/close is based ONLY on Running Position Quantity transitions.
      // OPTIONS / SPREADS preserve legacy unit-based behavior.
      if (tradeType === 'STOCK') {
        blkStart = (prevRunningQty === 0 && newRunningQty !== 0) ? 1 : 0;
        blkClose = (prevRunningQty !== 0 && newRunningQty === 0) ? 1 : 0;
      } else {
        blkStart = (prevUnit === 0 && newUnit !== 0) ? 1 : 0;
        blkClose = (prevUnit !== 0 && newUnit === 0) ? 1 : 0;
      }

      // Reset P&L accumulators at the START of every new block.
      if (blkStart) {
        blocks[key].pnl = 0;
        blocks[key].entryCost = 0;
      }

      // ── CTX counters ──────────────────────────────────────────────────────
      if (blkStart) ctxBlocksOpened++;
      if (blkClose) ctxBlocksClosed++;
      if (action === 'RAD' && ticker) ctxRADRows++;
      // ─────────────────────────────────────────────────────────────────────

      // Build unique Position ID when the block starts.
      let posId = blocks[key].positionId;
      if (blkStart && ticker && tradeType !== '') {
        const stratAbbrevPos = getStratAbbrev(strategyType || '');
        const tgSuffixPos = `TG${String(curBlock).padStart(3, '0')}`;

        if (spreadId) {
          posId = `${spreadId}-${tgSuffixPos}`;
        } else if (tradeType === 'OPTION') {
          const expStr2 = (colMap['option expiration'] !== undefined &&
            row[colMap['option expiration'] - 1] instanceof Date)
            ? Utilities.formatDate(row[colMap['option expiration'] - 1], tz, 'yyMMdd') : '';
          const strike2 = Number(row[colMap['option strike'] - 1]) || 0;
          const cp2 = row[colMap['call/put'] - 1].toString().toUpperCase()
            .replace('CALL', 'C').replace('PUT', 'P');
          posId = `${acct}-${ticker}-${stratAbbrevPos}${expStr2 ? `-${expStr2}` : ''}-${String(Math.round(strike2)).padStart(5, '0')}${cp2}-${tgSuffixPos}`;
        } else {
          posId = `${acct}-${ticker}-${stratAbbrevPos}-${tgSuffixPos}`;
        }

        blocks[key].positionId = posId;
        blocks[key].strategyType = strategyType;
      }
      row[colMap['position id'] - 1] = posId;

      // Trade Group ID
      const optExpRaw = row[colMap['option expiration'] - 1];
      const optExpStr = optExpRaw instanceof Date ? Utilities.formatDate(optExpRaw, tz, 'yyMMdd') : '';
      const isOption = tradeType === 'OPTION';
      const stratAbbrev = getStratAbbrev(strategyType || blocks[key].strategyType);
      const tgSuffix = `TG${String(curBlock).padStart(3, '0')}`;

      const tradeGroupId = spreadId
        ? `${spreadId}-${tgSuffix}`
        : `${acct}-${ticker}-${stratAbbrev}${isOption && optExpStr ? '-' + optExpStr : ''}-${tgSuffix}`;

      row[colMap['trade group id'] - 1] = tradeGroupId;
      blocks[key].tradeGroupId = tradeGroupId;

      row[colMap['block start flag'] - 1] = blkStart;
      row[colMap['block number'] - 1] = curBlock;
      row[colMap['block close flag/p&l'] - 1] = blkClose;
      row[colMap['running position quantity'] - 1] = blocks[key].runningQty;

      // === P&L CALCULATION ===
      const hasOptionFields = !!(row[colMap['option strike'] - 1] || row[colMap['option expiration'] - 1]);
      const multiplier = (tradeType === 'STOCK' || !hasOptionFields) ? 1 : 100;
      let sign = 0;
      if (action.includes('SELL TO')) sign = 1;
      if (action.includes('BUY TO')) sign = -1;
      blocks[key].pnl += sign * Number(row[colMap['entry price'] - 1]) * qty * multiplier;
      if (action.includes('OPEN')) {
        blocks[key].entryCost += Math.abs(sign * Number(row[colMap['entry price'] - 1]) * qty * multiplier);
      }

      if (blkClose) {
        row[colMap['realized p&l'] - 1] = blocks[key].pnl;
        row[colMap['percent p&l'] - 1] = blocks[key].entryCost
          ? (blocks[key].pnl / blocks[key].entryCost * 100)
          : 0;
        row[colMap['trade status'] - 1] = 'Closed';
        row[colMap['closing date'] - 1] = row[colMap['trade date'] - 1];
        if (row[tsIdx] && blocks[key].openTs) {
          const days = (row[tsIdx] - blocks[key].openTs) / (1000 * 60 * 60 * 24);
          row[colMap['trade duration'] - 1] = Math.round(days * 100) / 100;
        }
        blocks[key].block++;

      } else if (blkStart) {
        row[colMap['trade status'] - 1] = 'Open';
        blocks[key].openTs = row[tsIdx];
      }

    } // end main row loop

    // ── POST-PASS: Link stock settlement legs to their parent spread block ────
    const closingSpreadLookup = {};
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const blockCloseFlag = row[colMap['block close flag/p&l'] - 1];
      const rowTradeType = row[colMap['trade type'] - 1].toString().toUpperCase().trim();
      if (blockCloseFlag !== 1 && blockCloseFlag !== '1') continue;
      if (rowTradeType === 'STOCK') continue;
      const rowAcct = row[colMap['account'] - 1].toString().toUpperCase();
      const rowTicker = row[colMap['ticker'] - 1].toString().toUpperCase();
      const rowTgId = row[colMap['trade group id'] - 1].toString().trim();
      const rowStrat = row[colMap['strategy type'] - 1].toString().trim();
      const closingDate = row[colMap['closing date'] - 1];
      if (!rowAcct || !rowTicker || !rowTgId || !(closingDate instanceof Date)) continue;
      const closingDateStr = Utilities.formatDate(closingDate, tz, 'yyyy-MM-dd');
      closingSpreadLookup[`${rowAcct}|${rowTicker}|${closingDateStr}`] = { tgId: rowTgId, strategy: rowStrat };
    }

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowTradeType = row[colMap['trade type'] - 1].toString().toUpperCase().trim();
      if (rowTradeType !== 'STOCK') continue;
      const rowAction = row[colMap['action'] - 1].toString().trim().toUpperCase();
      if (rowAction !== 'BUY TO OPEN' && rowAction !== 'SELL TO CLOSE') continue;
      const rowAcct = row[colMap['account'] - 1].toString().toUpperCase();
      const rowTicker = row[colMap['ticker'] - 1].toString().toUpperCase();
      const tradeDate = row[colMap['trade date'] - 1];
      if (!(tradeDate instanceof Date)) continue;
      const tradeDateStr = Utilities.formatDate(tradeDate, tz, 'yyyy-MM-dd');
      const parent = closingSpreadLookup[`${rowAcct}|${rowTicker}|${tradeDateStr}`];
      if (!parent) continue;
      row[colMap['trade group id'] - 1] = `${parent.tgId}-ST`;
      row[colMap['position id'] - 1] = `${parent.tgId}-ST`;
      row[colMap['strategy type'] - 1] = parent.strategy;
      importIssuesAdd(ctx, 'INFO', dataStartRow + i, 'Trade Group ID',
        `${parent.tgId}-ST`,
        `Stock settlement leg linked to parent spread block. TG ID → ${parent.tgId}-ST`);
    }
    // ── END POST-PASS ─────────────────────────────────────────────────────────

    let openAtEnd = 0;
    Object.values(blocks).forEach(b => {
      const strategyUpper = String(b.strategyType || '').trim().toUpperCase();
      const looksLikeStock =
        strategyUpper.includes('STOCK') &&
        !strategyUpper.includes('CALL') &&
        !strategyUpper.includes('PUT');

      if (looksLikeStock) {
        if (Number(b.runningQty || 0) !== 0) openAtEnd++;
      } else {
        if (Number(b.unit || 0) !== 0) openAtEnd++;
      }
    });

    importIssuesSetMetric(ctx, 'SourceRowsReadExclHeader', data.length);
    importIssuesSetMetric(ctx, 'RowsWrittenExclHeader', data.length);
    importIssuesSetMetric(ctx, 'SpreadGroupsBuilt', Object.keys(spreadKeyMap).length);
    importIssuesSetMetric(ctx, 'StrikeCollisions', ctxStrikeCollisions);
    importIssuesSetMetric(ctx, 'BlocksOpened', ctxBlocksOpened);
    importIssuesSetMetric(ctx, 'BlocksClosed', ctxBlocksClosed);
    importIssuesSetMetric(ctx, 'RADRowsParsed', ctxRADRows);
    importIssuesSetMetric(ctx, 'OpenPositionsAtEnd', openAtEnd);
    importIssuesSetMetric(ctx, 'MissingSpreadGroupIds', ctxMissingSpreadGroup);
    importIssuesSetMetric(ctx, 'Success', '1');

    if (openAtEnd > 0) {
      importIssuesAdd(ctx, 'INFO', 'End of data', 'Open Positions',
        openAtEnd + ' position(s)',
        'These positions have no closing event in the current dataset — ' +
        'expected if you have live LEAP spreads still open.');
    }

    const outputGrid = [helperData[0], helperData[1], helperData[2], ...data];
    stagingSheet.clearContents();
    stagingSheet.getRange(1, 1, outputGrid.length, outputGrid[0].length).setValues(outputGrid);

    checkMissingDateTimeAndAlert(stagingSheet, 4, 'populateStagingWithBlockLogicV3');

    SpreadsheetApp.getUi().alert(
      '✅ Block logic V3 updated! Trade Group ID increments once per open-to-flat block — ' +
      'RAD expirations close perfectly.'
    );

  } catch (e) {
    importIssuesSetMetric(ctx, 'Success', '0');
    importIssuesSetMetric(ctx, 'ErrorMessage', e.message);
    importIssuesSetMetric(ctx, 'ErrorStack', (e.stack || '').substring(0, 500));
    throw e;

  } finally {
    stagingIssuesFlush(ctx);
  }
}

// =========================================================================
// MASTER UTILITIES
//   Final write and safety tools that operate on the Master sheet.
// =========================================================================
function appendStagingToMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("Master");
  const staging = ss.getSheetByName("Staging");
  if (!master || !staging) throw new Error("Sheets not found.");

  const numCols = staging.getLastColumn();

  // Get Staging data (row 4 down)
  // Changed to see if all rows are filled
  const accountColIdx = staging.getRange(1, 1, 1, numCols).getValues()[0].indexOf("Trade Date") + 1;
  //const accountColIdx = staging.getRange(1, 1, 1, numCols).getValues()[0].indexOf("Trade Date") + 1;
  const stagingData = staging.getRange(4, 1, staging.getLastRow() - 3, numCols).getValues()
    .filter(row => row[accountColIdx - 1] !== "" && row[accountColIdx - 1] !== null);

  // Find first empty row in Master (after header)
  const firstEmptyMasterRow = master.getLastRow() + 1;

  if (stagingData.length) {
    master.getRange(firstEmptyMasterRow, 1, stagingData.length, numCols).setValues(stagingData);
  }
}
function backupMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), m = ss.getSheetByName('Master');
  const name = 'Master_Backup_' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd_HHmmss');
  m.copyTo(ss).setName(name);
  //logAction('BACKUP MASTER',name);
}
/**
 * Clears all data rows from the Master sheet while keeping the header row.
 * Note: Sheet name is case-sensitive — must be exactly "Master".
 */
function clearMasterExceptHeader() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Master');
  if (!sheet) {
    throw new Error('Sheet "Master" not found.');
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
}

