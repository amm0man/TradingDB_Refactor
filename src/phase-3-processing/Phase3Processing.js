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
 *    - appendStagingToMaster(), backupMasterSheet() - clearMasterExceptHeader() lives in SheetBlanking.js
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




// =========================================================================
// STEP 2: Validate & Clean has been moved
//   → see Phase3Step2_ValidateClean.js
//   (parseTradeTimeStamp + validateAndCleanImportToHelperV3)
// =========================================================================

// =========================================================================
// STEP 3: Block Logic has been moved
//   → see Phase3BlockLogic.js
//   (populateStagingWithBlockLogicV3 + resolveLiveSpreadGroupId)
// =========================================================================

// =========================================================================
// MASTER UTILITIES have been moved
//   → see Phase3MasterUtils.js
//   (appendStagingToMaster + backupMasterSheet)
// =========================================================================



