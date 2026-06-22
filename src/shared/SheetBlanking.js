
/**
 * SheetBlanking.js
 *
 * Collection of helper functions to safely clear data from various sheets
 * while preserving header rows.
 *
 * Used heavily during re-imports and pipeline resets so old data doesn't
 * mix with new data.
 *
 * Main functions:
 *   - tosBlankSheetExceptHeader_()     → core reusable blanking engine (private)
 *   - tosBlankTosTop(), tosBlankTosTrades(), etc. → specific TOS sheets
 *   - tosBlankALLTOSSheets()           → one-click clear of all TOS-related sheets
 *   - blankAllSchwabSheets()           → clears Schwab Import + Schwab Mapping
 *   - blankAllPrepSheets()             → clears Import / Helper / Staging
 *
 * Note: There is a large block of commented-out legacy code below
 * (the old clearSchwabImportExceptHeader that protected formula columns).
 * That code can be safely removed once we're confident it's no longer needed.
 */
/**
 * TOSTrades and TOSTop and Combined Sheets Blanking
 * Clears a sheet's contents EXCEPT the header row (row 1).
 * - If the sheet does not exist, it is created.
 * - If row 1 is blank and defaultHeaders is provided, the default headers are written.
 */
function tosBlankSheetExceptHeader_(sheetName, defaultHeaders) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(sheetName);
    if (!sh) sh = ss.insertSheet(sheetName);

    // Determine how many columns we should treat as "in use" for clearing.
    const lastCol = Math.max(sh.getLastColumn(), (defaultHeaders && defaultHeaders.length) ? defaultHeaders.length : 0, 1);

    // Read header row (row 1).
    const headerRange = sh.getRange(1, 1, 1, lastCol);
    const header = headerRange.getValues()[0].map(v => String(v ?? '').trim());

    const headerIsBlank = header.every(h => !h);

    // If header row is blank, and we were given a default header, write it.
    if (headerIsBlank && defaultHeaders && defaultHeaders.length) {
      sh.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
    }

    // Clear everything below header row.
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    }

    // SpreadsheetApp.getUi().alert('Cleared (kept headers): ' + sheetName);

  } finally {
    lock.releaseLock();
  }
}

// ----------------------------
// Individual clear functions
// ----------------------------
function tosBlankTosTop() {
  // TosTop canonical schema (Account-first)
  const headers = ['Account', 'DATE', 'TIME', 'TYPE', 'DESCRIPTION', 'Misc Fees', 'Commissions & Fees', 'AMOUNT'];
  return tosBlankSheetExceptHeader_('TosTop', headers);
}

function tosBlankTOSTopCombined() {
  // Conservative default header; actual importer will overwrite header anyway
  const headers = ['Account', 'SourceFile', 'DATE', 'TIME', 'TYPE', 'DESCRIPTION', 'Misc Fees', 'Commissions & Fees', 'AMOUNT', 'TimeRaw'];
  return tosBlankSheetExceptHeader_('TOS Top - Combined', headers);
}

function tosBlankTosTrades() {
  // TosTrades canonical schema (Account-first)
  const headers = ['Account', 'Exec Time', 'Spread', 'Side', 'Qty', 'Pos Effect', 'Symbol', 'Exp', 'Strike', 'Type', 'Price', 'Net Price', 'Order Type'];
  return tosBlankSheetExceptHeader_('TosTrades', headers);
}

function tosBlankTOSTradesCombined() {
  // Conservative default header; actual importer will overwrite header anyway
  const headers = ['Account', 'SourceFile', 'Exec Time', 'Spread', 'Side', 'Qty', 'Pos Effect', 'Symbol', 'Exp', 'Strike', 'Type', 'Price', 'Net Price', 'Order Type'];
  return tosBlankSheetExceptHeader_('TOS Trades - Combined', headers);
}

// ----------------------------
// One-click: clear them all
// ----------------------------

function tosBlankALLTOSSheets() {
  tosBlankTosTop();
  tosBlankTOSTopCombined();
  tosBlankTosTrades();
  tosBlankTOSTradesCombined();
}

/**
 * Blank Schwab Import and Schwab Mapping Sheets
 */
function blankAllSchwabSheets() { clearSchwabImportExceptHeader(); clearSchwabMappingExceptHeader(); logAction('BLANK ALL Schwab SHEETS', ''); }


function clearSchwabImportExceptHeader() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Schwab Import');
  if (!sheet) throw new Error('Missing sheet: "Schwab Import"');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return; // header only — nothing to clear

  // Clear ALL data below row 1 — no formula columns to protect anymore.
  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
}


// Clear Schwab Mapping
function clearSchwabMappingExceptHeader() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Schwab Mapping");
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
}
function blankImport() { clearData('Import', 3); logAction('BLANK IMPORT', ''); }
function blankHelper() { clearData('Helper', 4); logAction('BLANK HELPER', ''); }
function blankStaging() { clearData('Staging', 4); logAction('BLANK STAGING', ''); }
function blankAllPrepSheets() { blankImport(); blankHelper(); blankStaging(); logAction('BLANK ALL PREP SHEETS', ''); }

function clearData(sheetName, startRow) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lr = sh.getLastRow();
  if (lr >= startRow) sh.getRange(startRow, 1, lr - startRow + 1, sh.getLastColumn()).clearContent();
}
