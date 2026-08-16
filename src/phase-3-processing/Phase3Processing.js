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

// =========================================================================
// STEP 3: Block Logic has been moved
//   → see Phase3BlockLogic.js
//   (populateStagingWithBlockLogicV3 + resolveLiveSpreadGroupId)
// =========================================================================

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

