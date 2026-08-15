/**
 * TosSchwabImportPipeline.js
 *
 * Core Phase 1 import pipeline.
 * Brings trade data from thinkorswim (TOS) CSV exports into the Trading Database.
 *
 * High-level flow:
 *   Drive folders (TosTrades / TosTop)
 *     → "TOS Trades - Combined" + "TOS Top - Combined"
 *     → "TosTrades" + "TosTop"
 *     → later handed off to Schwab Mapping / Phase 2
 *
 * Key responsibilities:
 *   - Folder selection and CSV import (single account or both LT+DT)
 *   - Parsing and normalizing TOS trade / top-of-book data
 *   - Writing to Combined sheets while preserving the other account’s rows
 *   - Pushing Combined data into the working TosTrades / TosTop sheets
 *   - ET → CT time correction
 *   - Issue logging via ImportIssues.js
 *   - Debug alerts controlled by SettingsService.js
 *
 * This is one of the largest files in the project (~1,700 lines).
 * Current focus: improve readability with clear comments before any
 * structural refactoring.
 *
 * Related files:
 *   - SettingsService.js
 *   - ImportIssues.js
 *   - TosSheetWriteHelpers.js
 *   - BuildUnifiedImportV3.js (next stage after this pipeline)
 */

/** ---- Configuration ---- */
const tosConfig = {
  tradesCombinedSheetName: 'TOS Trades - Combined',
  topCombinedSheetName: 'TOS Top - Combined',
  tosTradesSheetName: 'TosTrades',
  tosTopSheetName: 'TosTop',

  // If true, imports for one account will REPLACE only that Account's rows in Combined sheets,
  // while preserving the other account's rows in the same Combined sheet.
  // This is the key behavior that lets you keep LT + DT together without overwriting.
  combinedReplaceOnlyThatAccount: true
};

// ─────────────────────────────────────────────────────────────────────────────
// ET → CT DST-aware offset helpers
//
// WHY THIS EXISTS:
//   TOS Account Statement exports always write timestamps in Eastern Time (ET),
//   regardless of where you live. Users in Central Time (CT) are UTC-6 in winter
//   (CST) and UTC-5 in summer (CDT). Eastern Time is UTC-5 in winter (EST) and
//   UTC-4 in summer (EDT). The difference is:
//     - Winter (CST vs EST): CT = ET − 1 hour → add +1 hour to correct
//     - Summer (CDT vs EDT): CT = ET − 0 hours → add +0 hours (no correction)
//   DST in the US begins the 2nd Sunday of March at 2:00 AM and ends the
//   1st Sunday of November at 2:00 AM. All market-hours trades occur between
//   8:30 AM and 3:00 PM CT and never straddle the 2 AM switchover, so
//   checking by date is safe and accurate.
//
// IMPORTANT: This correction applies to ALL years, past and future, because
//   the ET-export behavior has been consistent across TDA and Schwab-era TOS.
//   The 2022 vs 2025 discrepancy you observed is purely a DST phenomenon:
//   winter trades appear 1 hour early, summer trades appear correct.
// ─────────────────────────────────────────────────────────────────────────────
/**
 * Returns the milliseconds to ADD to an ET-stamped Date to convert it to CT.
 * Result is either 0 (summer/DST active) or 3,600,000 (winter/standard time).
 * @param {Date} dateInEt - A Date object whose hour values reflect Eastern Time.
 * @returns {number} Milliseconds offset (0 or 3600000).
 */
// AFTER — correct, CT is always UTC-1 behind ET, every day of the year:
function tosEtToCtOffsetMs(dateInEt) {
  // Central Time (CT) is always exactly 1 hour behind Eastern Time (ET),
  // year-round. Both zones observe US DST simultaneously, so the offset
  // never changes: CST(UTC-6) vs EST(UTC-5) in winter = -1hr,
  // CDT(UTC-5) vs EDT(UTC-4) in summer = -1hr. Always add +1 hour.
  return 60 * 60 * 1000; // 3,600,000 ms
}

/**
 * Applies the ET → CT +1 hour correction to an HHmmss text string,
 * given the trade date (needed to handle midnight rollovers correctly).
 * Returns a corrected HHmmss string (6 chars, zero-padded).
 * @param {string} hhmmss  - e.g. "073115"
 * @param {string} dateIso - e.g. "2022-06-09" (yyyy-MM-dd)
 * @returns {string} corrected HHmmss, e.g. "083115"
 */
function tosEtToCtHHmmss(hhmmss, dateIso) {
  if (!hhmmss || hhmmss.length < 4) return hhmmss;
  // Build a Date from the date + time so we can add 1 hour safely
  // (handles the rare midnight rollover case automatically)
  const HH = parseInt(hhmmss.substring(0, 2), 10);
  const MIN = parseInt(hhmmss.substring(2, 4), 10);
  const SS = hhmmss.length >= 6 ? parseInt(hhmmss.substring(4, 6), 10) : 0;
  const m = dateIso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return hhmmss; // can't parse date, return unchanged
  const dEt = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10), HH, MIN, SS, 0);
  if (isNaN(dEt.getTime())) return hhmmss;
  const corrected = new Date(dEt.getTime() + tosEtToCtOffsetMs(dEt));
  const cHH = String(corrected.getHours()).padStart(2, '0');
  const cMIN = String(corrected.getMinutes()).padStart(2, '0');
  const cSS = String(corrected.getSeconds()).padStart(2, '0');
  return cHH + cMIN + cSS;
}

/**
 * Converts a TosTop DATE cell value (e.g. "6/9/2022", "6/9/22", or a Date object)
 * to a yyyy-MM-dd ISO string for use in sorting and time correction.
 * @param {string|Date} dateVal
 * @returns {string} e.g. "2022-06-09", or "" if unparseable
 */
function tosTopNormalizeDateToIso(dateVal) {
  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
    const y = dateVal.getFullYear();
    const m = String(dateVal.getMonth() + 1).padStart(2, '0');
    const d = String(dateVal.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }

  const s = String(dateVal ?? '').trim();
  if (!s) return '';

  // m/d/yyyy OR m/d/yy
  const match = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (match) {
    let yyyy = match[3];
    if (yyyy.length === 2) {
      const yy = parseInt(yyyy, 10);
      yyyy = String(yy >= 70 ? 1900 + yy : 2000 + yy);
    }
    return yyyy + '-' + match[1].padStart(2, '0') + '-' + match[2].padStart(2, '0');
  }

  // Already yyyy-MM-dd
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    return s.substring(0, 10);
  }

  return '';
}

/**
 * ScriptProperties keys (stored server-side, persistent across runs).
 * Note: The key strings can be anything; these are stable and descriptive.
 */
const tosProps = {
  tradesFolderIdLT: 'TOS_TRADES_FOLDER_ID_LT',
  tradesFolderIdDT: 'TOS_TRADES_FOLDER_ID_DT',
  topFolderIdLT: 'TOS_TOP_FOLDER_ID_LT',
  topFolderIdDT: 'TOS_TOP_FOLDER_ID_DT',
};

/**
 * Canonical account selector for this project:
 * - Reads SettingService.gs key: accountMode (DT/LT)
 * - Falls back to DT if missing
 */
function tosGetCurrentAccount() {
  const v = String(getSetting('accountMode', 'DT') || '').trim().toUpperCase();
  return (v === 'DT' || v === 'LT') ? v : 'DT';
}

/** ---- Folder setup UI helpers ---- */

/**
 * DEBUG-only alerts for TOS import steps.
 * Controlled by SettingsService ScriptProperty: TOS_IMPORT_DEBUG_ALERTS ("0"/"1")
 */
function tosMaybeDebugAlert(message) {
  const enabled = String(getSetting('TOS_IMPORT_DEBUG_ALERTS', '0')).trim() === '1';
  if (!enabled) return;



  try {
    SpreadsheetApp.getUi().alert(String(message || ''));
  } catch (e) {
    // If no UI is available (triggers, APIs), don't fail the run—just log.
    Logger.log('tosMaybeDebugAlert (no UI): ' + message);
  }
}

function tosUiAlertSafe(message) {
  try {
    SpreadsheetApp.getUi().alert(String(message || ''));
  } catch (e) {
    Logger.log('tosUiAlertSafe (no UI): ' + message);
  }
}


function tosTradesSetFolderIdLT() {
  tosSetFolderId('TOS Trades Folder ID (LT)', tosProps.tradesFolderIdLT);
}

function tosTradesSetFolderIdDT() {
  tosSetFolderId('TOS Trades Folder ID (DT)', tosProps.tradesFolderIdDT);
}

function tosTopSetFolderIdLT() {
  tosSetFolderId('TOS Top Folder ID (LT)', tosProps.topFolderIdLT);
}

function tosTopSetFolderIdDT() {
  tosSetFolderId('TOS Top Folder ID (DT)', tosProps.topFolderIdDT);
}

function tosSetFolderId(title, propKey) {
  const ui = SpreadsheetApp.getUi();

  const current = String(getSetting(propKey, '') || '').trim();
  const msg =
    'Paste the Google Drive Folder ID.\n\n' +
    'Setting key: ' + propKey + '\n' +
    (current ? ('Current: ' + current + '\n\n') : '\n') +
    'Tip: Open the folder in Drive and copy the long ID from the URL.';

  const resp = ui.prompt(title, msg, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const folderId = String(resp.getResponseText() || '').trim();
  if (!folderId) {
    ui.alert('Folder ID was blank.');
    return;
  }

  // Validate the folder ID (catches typos)
  try {
    DriveApp.getFolderById(folderId);
  } catch (e) {
    ui.alert('That folder ID did not validate in Drive.\n\n' + e);
    return;
  }

  // Canonical settings write
  setSetting(propKey, folderId);

  ui.alert('Saved folder ID:\n' + propKey + ' = ' + folderId);
}


// -----------------------------------------------------------------------------
// Menu / wrapper functions
// -----------------------------------------------------------------------------

// ------------------------------
// BOTH accounts menu actions
// ------------------------------
function tosTradesImportFromFolderBothAccounts() {
  tosTradesImportFromFolder(tosProps.tradesFolderIdDT, 'DT');
  tosTradesImportFromFolder(tosProps.tradesFolderIdLT, 'LT');
}

function tosTopImportFromFolderBothAccounts() {
  tosTopImportFromFolder(tosProps.topFolderIdDT, 'DT');
  tosTopImportFromFolder(tosProps.topFolderIdLT, 'LT');
}

// One-button: Drive CSVs -> Combined -> TosTop/TosTrades -> Schwab Import (both accounts)
function tosRunFullTosToSchwabImportBothAccounts() {
  // One pipeline-wide RunId so ALL steps can be filtered together in "Import Issues".
  const pipelineRunId = new Date().toISOString();
  setSetting('ACTIVE_IMPORT_RUN_ID', pipelineRunId);

  try {
    tosTradesImportFromFolderBothAccounts();
    tosTopImportFromFolderBothAccounts();
    pushTosCombinedToBoth();
    buildUnifiedImportV3();
  } finally {
    // Always clear so a later manual step doesn’t accidentally reuse this RunId.
    setSetting('ACTIVE_IMPORT_RUN_ID', '');
  }
}



function tosTradesImportFromFolderCurrentAccount() {
  const Account = tosGetCurrentAccount(); // "DT" or "LT"
  const folderKey = (Account === 'LT') ? tosProps.tradesFolderIdLT : tosProps.tradesFolderIdDT;
  tosTradesImportFromFolder(folderKey, Account);
}

function tosTopImportFromFolderCurrentAccount() {
  const Account = tosGetCurrentAccount(); // "DT" or "LT"
  const folderKey = (Account === 'LT') ? tosProps.topFolderIdLT : tosProps.topFolderIdDT;
  tosTopImportFromFolder(folderKey, Account);
}


function pushTosCombinedToBoth() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    pushTosTradesCombinedToTosTrades();
    pushTosTopCombinedToTosTop();
  } finally {
    lock.releaseLock();   // ✅ always runs — even if an inner function throws
  }
}

/** ======================================================================
 *  Trades import: CSV folder -> "TOS Trades - Combined"
 *  ====================================================================== */

function tosTradesImportFromFolder(folderIdPropKey, Account) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  const ctx = importIssuesStart('tosTradesImportFromFolder');
  importIssuesSetMetric(ctx, 'Account', Account || '');
  importIssuesSetMetric(ctx, 'OutputSheet', tosConfig.tradesCombinedSheetName);

  // NEW: prove which property key the caller asked us to use
  importIssuesSetMetric(ctx, 'FolderIdPropKey', folderIdPropKey || '');

  try {
    // NEW: guard - if caller forgot to pass a prop key, fail loudly (prevents accidental fallback behavior elsewhere)
    if (!folderIdPropKey) {
      importIssuesAdd(ctx, 'ERROR', '', '', 'FolderIdPropKey', '', 'Missing folderIdPropKey argument when calling tosTradesImportFromFolder().');
      importIssuesFlush(ctx);
      tosUiAlertSafe('Folder ID prop key was not provided to tosTradesImportFromFolder().');
      return;
    }

    const folderId = String(getSetting(folderIdPropKey, '') || '').trim();


    // NEW: log the exact folder id we read from ScriptProperties
    importIssuesSetMetric(ctx, 'FolderIdUsed', folderId || '');

    if (!folderId) {
      importIssuesAdd(ctx, 'ERROR', '', '', 'FolderId', '', 'Folder ID not set for key: ' + folderIdPropKey);
      importIssuesFlush(ctx);
      tosUiAlertSafe('Folder ID not set. Run the folder setup menu item first.');
      return;
    }

    const folder = DriveApp.getFolderById(folderId);

    // NEW: log the folder name so you can instantly tell if it’s the DT or LT folder
    const folderName = folder.getName();
    importIssuesSetMetric(ctx, 'FolderNameUsed', folderName);
    // Optional immediate proof popup (controlled by Settings > Toggle TOS Import DEBUG Alerts)
    tosMaybeDebugAlert(
      'tosTradesImportFromFolder DEBUG\n\n' +
      'Account: ' + (Account || '') + '\n' +
      'FolderIdPropKey: ' + folderIdPropKey + '\n' +
      'FolderIdUsed: ' + folderId + '\n' +
      'FolderNameUsed: ' + folderName
    );



    const files = folder.getFiles();
    const csvFiles = [];

    while (files.hasNext()) {
      const f = files.next();
      const name = (f.getName() || '').toLowerCase();
      const mime = f.getMimeType();

      // TOS exports are often text/csv but sometimes show as "Microsoft Excel"
      if (name.endsWith('.csv') || mime === MimeType.CSV || mime === MimeType.MICROSOFT_EXCEL) {
        csvFiles.push(f);
      }
    }

    importIssuesSetMetric(ctx, 'FilesFound', csvFiles.length);

    if (!csvFiles.length) {
      importIssuesAdd(ctx, 'WARN', '', '', 'Files', '', 'No CSV-like files found in the folder.');
      importIssuesFlush(ctx);
      tosUiAlertSafe('No CSV-like files found in folder.');
      return;
    }

    // Sort deterministically by file name (helps debugging + repeatability)
    csvFiles.sort((a, b) => (a.getName() || '').localeCompare(b.getName() || ''));

    // NEW: record the first file name as a quick sanity check (optional but very useful)
    importIssuesSetMetric(ctx, 'FirstCsvFile', (csvFiles[0] && csvFiles[0].getName()) ? csvFiles[0].getName() : '');

    // Read + combine rows from all files
    let allRows = [];           // { Account, sourceFile, row }
    let canonicalHeader = null; // the first header we find becomes the output header

    for (let i = 0; i < csvFiles.length; i++) {
      const file = csvFiles[i];
      const parsed = tosTradesParseOneCsvFile(file, ctx);

      if (!parsed || !parsed.rows || !parsed.rows.length) continue;

      importIssuesSetMetric(ctx, 'ParsedRowsTotal', Number(ctx.metrics.ParsedRowsTotal || 0) + parsed.rows.length);

      if (!canonicalHeader) canonicalHeader = parsed.header;

      parsed.rows.forEach(r => {
        allRows.push({
          Account: Account || '',
          sourceFile: file.getName(),
          row: r
        });
      });
    }

    if (!canonicalHeader) {
      importIssuesAdd(ctx, 'ERROR', '', '', 'Header', '', 'Could not find "Exec Time" + "Spread" header row in any CSV.');
      importIssuesFlush(ctx);
      tosUiAlertSafe('Could not find the "Exec Time / Spread / ..." header row in any CSV.');
      return;
    }

    // Fill-down Exec Time + Spread within each file (never carry to next file)
    const idxExec = canonicalHeader.indexOf('Exec Time');
    const idxSpread = canonicalHeader.indexOf('Spread');
    if (idxExec < 0 || idxSpread < 0) {
      throw new Error('Trades header missing Exec Time or Spread.');
    }

    let lastExec = '';
    let lastSpread = '';
    let lastFile = '';

    for (let i = 0; i < allRows.length; i++) {
      const obj = allRows[i];
      const r = obj.row;

      if (obj.sourceFile !== lastFile) {
        lastFile = obj.sourceFile;
        lastExec = '';
        lastSpread = '';
      }

      if (tosTradesIsRowBlank(r)) {
        ctx.metrics.BlankRowsSkipped++;
        ctx.metrics.TotalRowsSkipped++;
        continue;
      }

      const execVal = String(r[idxExec] ?? '').trim();
      const spreadVal = String(r[idxSpread] ?? '').trim();

      if (execVal) lastExec = execVal;
      else r[idxExec] = lastExec;

      if (spreadVal) lastSpread = spreadVal;
      else r[idxSpread] = lastSpread;
    }

    // Dedupe: overlap-safe but preserves legitimate duplicates within a single file.
    //
    // IMPORTANT: The key includes Account so LT/DT identical trades never collapse each other.
    const bucketsByKey = {}; // key -> { rowTemplate, countsByFile: { [fileName]: n }, Account }

    for (let i = 0; i < allRows.length; i++) {
      const obj = allRows[i];
      const r = obj.row;

      if (tosTradesIsRowBlank(r)) continue;

      const accountKey = String(obj.Account || '');
      const rowKey = r.join('\u0001');
      const key = accountKey + '\u0001' + rowKey;

      if (!bucketsByKey[key]) {
        bucketsByKey[key] = {
          rowTemplate: r.slice(),
          countsByFile: {},
          Account: accountKey
        };
      }

      const f = obj.sourceFile || '';
      bucketsByKey[key].countsByFile[f] = (bucketsByKey[key].countsByFile[f] || 0) + 1;
    }

    const deduped = []; // { Account, sourceFile, row }

    const keys = Object.keys(bucketsByKey);
    ctx.metrics.CombinedRowsTotal = allRows.length;
    ctx.metrics.UniqueKeys = keys.length;

    for (let k = 0; k < keys.length; k++) {
      const key = keys[k];
      const bucket = bucketsByKey[key];

      // Pick the file that had the most occurrences of this exact row
      let bestFile = '';
      let bestCount = 0;

      const fileNames = Object.keys(bucket.countsByFile);
      for (let i = 0; i < fileNames.length; i++) {
        const f = fileNames[i];
        const c = bucket.countsByFile[f] || 0;

        if (c > bestCount) {
          bestCount = c;
          bestFile = f;
        } else if (c === bestCount && f < bestFile) {
          bestFile = f;
        }
      }

      // Emit the row bestCount times to preserve legitimate duplicates
      for (let n = 0; n < bestCount; n++) {
        deduped.push({
          Account: bucket.Account,
          sourceFile: bestFile,
          row: bucket.rowTemplate.slice()
        });
      }
    }

    ctx.metrics.DedupedRows = deduped.length;

    // Sort by Exec Time (oldest -> newest)
    deduped.sort((a, b) => {
      const da = tosTradesParseExecTime(a.row[idxExec]);
      const db = tosTradesParseExecTime(b.row[idxExec]);

      if (da && db) return da - db;
      if (da && !db) return -1;
      if (!da && db) return 1;

      // Tie-breaker: Account, filename, row content
      const al = (a.Account || '').localeCompare(b.Account || '');
      if (al !== 0) return al;

      const fa = a.sourceFile || '';
      const fb = b.sourceFile || '';
      if (fa !== fb) return fa.localeCompare(fb);

      return a.row.join('|').localeCompare(b.row.join('|'));
    });

    // -----------------------------
    // Build output rows:
    //   Account | SourceFile | (canonicalHeader...) | TimeRaw
    //
    // NEW: TimeRaw lets you see the exact Exec Time text as it appeared in the CSV row.
    // This makes it MUCH easier to debug "same-minute" collisions and ordering.
    // -----------------------------
    const newOut = [];

    const outHeader = ["Account", "SourceFile", ...canonicalHeader, "TimeRaw"];
    newOut.push(outHeader);

    // Exec Time index comes from earlier canonicalHeader lookup (idxExec defined above).


    // NEW metric: how many parsed rows have Exec Time with no seconds.
    // This will typically be "all rows" for current TOS Trades exports (they only show HH:mm),
    // but keeping the metric helps prevent future confusion.
    let execTimeMissingSecondsCount = 0;
    let execTimeWithSecondsCount = 0;

    for (let i = 0; i < deduped.length; i++) {
      const obj = deduped[i];

      // Defensive: keep Symbol as string so CUSIP-like values don't get coerced.
      const rowCopy = obj.row.slice();
      const idxSymbolInCanonical = canonicalHeader.indexOf('Symbol');
      if (idxSymbolInCanonical >= 0) rowCopy[idxSymbolInCanonical] = String(rowCopy[idxSymbolInCanonical] ?? '').trim();

      const timeRaw = idxExec >= 0 ? String(rowCopy[idxExec] ?? '').trim() : '';

      const hasSeconds = /\d{2}[AP]?M?$/i.test(timeRaw);
      if (timeRaw) {
        if (hasSeconds) execTimeWithSecondsCount++;
        else execTimeMissingSecondsCount++;
      }

      // ET→CT: parse the raw CSV string, apply +1hr, store as "yyyy-MM-dd HH:mm:ss" text.
      // Storing as text prevents Google Sheets from re-shifting the value with its own timezone.
      if (idxExec >= 0 && timeRaw) {
        const parsedEt = tosTradesParseExecTimeRaw(timeRaw); // see new helper below
        if (parsedEt) {
          rowCopy[idxExec] = parsedEt; // now a "yyyy-MM-dd HH:mm:ss" string
        }
      }

      newOut.push([obj.Account, obj.sourceFile, ...rowCopy, timeRaw]);

    }

    // Write metrics so Import Issues tells us whether we can trust seconds from Trades.
    importIssuesSetMetric(ctx, 'TradesExecTimeWithSecondsCount', execTimeWithSecondsCount);
    importIssuesSetMetric(ctx, 'TradesExecTimeMissingSecondsCount', execTimeMissingSecondsCount);

    // Optional: log one WARN if seconds are missing (helps prevent downstream confusion).
    if (execTimeMissingSecondsCount > 0 && execTimeWithSecondsCount === 0) {
      importIssuesAdd(ctx, 'WARN', 'tosTradesImportFromFolder', '', 'Exec Time',
        'Seconds missing in Trades CSV',
        'TOS Trades CSV Exec Time appears to be minute precision (HH:mm). Matching will rely on TosTop seconds + price tie-breakers.');
    }



    // Write into Combined sheet. If combinedReplaceOnlyThatAccount==true, we preserve the other accounts already in the sheet.
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = tosGetOrCreateSheet(ss, tosConfig.tradesCombinedSheetName);

    // Keep the other account's rows, replace only this account.
    const finalOut = tosMergeAccountLabeledCombined(sh, newOut, Account);

    // Clear prior contents BEFORE applying text formats + writing values.
    sh.clearContents();

    // CRITICAL preserve leading zeros CUSIP-like Symbols by forcing Symbol to text BEFORE setValues.
    tosFormatHeaderColumnAsText(sh, finalOut[0], 'Symbol', finalOut.length, 1);
    // CRITICAL store Exec Time as plain text so Sheets cannot coerce it to a Date serial.
    // Without this, getValues() on a downstream read returns a shifted Date instead of the string.
    tosFormatHeaderColumnAsText(sh, finalOut[0], 'Exec Time', finalOut.length, 1);
    // Now write the sheet.
    sh.getRange(1, 1, finalOut.length, finalOut[0].length).setValues(finalOut);



    // Safety sort by Exec Time in-sheet (works as text too, but best if consistent)
    const execCol = finalOut[0].indexOf('Exec Time') + 1;
    if (execCol > 0 && finalOut.length > 2) {
      sh.getRange(2, 1, finalOut.length - 1, finalOut[0].length).sort({ column: execCol, ascending: true });
    }

    ctx.metrics.RowsWrittenExclHeader = finalOut.length - 1;
    importIssuesFlush(ctx);

    tosUiAlertSafe(
      'Trades import complete.\n' +
      'Account: ' + (Account || '') + '\n' +
      'FolderPropKey: ' + folderIdPropKey + '\n' +
      'FolderName: ' + folderName + '\n' +
      'Files: ' + csvFiles.length + '\n' +
      'Rows written (excl header): ' + (finalOut.length - 1)
    );

  } catch (err) {
    importIssuesAdd(ctx, 'ERROR', '', '', 'Exception', '', String(err && err.stack ? err.stack : err));
    importIssuesFlush(ctx);
    throw err;
  } finally {
    lock.releaseLock();
  }
}


/**
 * Trades CSV parser:
 * - Finds header row containing "Exec Time" and "Spread"
 * - Uses stop rule: blank row then "EQUITIES" in column A stops
 * - Hard stops if "OPTIONS" or "PROFITS AND LOSSES" appear in column A
 */
function tosTradesParseOneCsvFile(file, ctx) {
  let text = file.getBlob().getDataAsString();

  // Strip BOM if present (can break exact header matches)
  text = text.replace(/^\uFEFF/, '');

  // Parse comma first; if it looks 1-column and contains tabs, re-parse as tab-delimited
  let grid = Utilities.parseCsv(text);
  const looksSingleColumn = grid && grid.length && grid[0].length === 1;
  const containsTabs = text.indexOf('\t') !== -1;

  if (looksSingleColumn && containsTabs) {
    grid = Utilities.parseCsv(text, '\t');
  }

  // Find header row containing Exec Time + Spread
  let headerRowIdx = -1;
  let startCol = -1;

  for (let r = 0; r < grid.length; r++) {
    const row = grid[r].map(v => (v == null ? '' : String(v).trim()));
    const execIdx = row.indexOf('Exec Time');
    const spreadIdx = row.indexOf('Spread');

    if (execIdx !== -1 && spreadIdx !== -1) {
      headerRowIdx = r;
      startCol = execIdx;
      break;
    }
  }

  if (headerRowIdx < 0) return null;

  const fullHeader = grid[headerRowIdx].map(v => (v == null ? '' : String(v).trim()));
  const header = fullHeader.slice(startCol);

  const idxSide = header.indexOf('Side');
  const idxSymbol = header.indexOf('Symbol');
  const idxQty = header.indexOf('Qty');
  const idxPrice = header.indexOf('Price');

  const rows = [];
  let sawBlank = false;

  function isBlankRawRow(raw) {
    return raw.every(v => String(v ?? '').trim() === '');
  }

  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const raw = grid[r].map(v => (v == null ? '' : String(v).trim()));

    // Track blank rows (before slicing)
    if (isBlankRawRow(raw)) {
      sawBlank = true;
      continue;
    }

    const colA = String(raw[0] || '').trim().toUpperCase();

    // Stop rule: blank row followed by "EQUITIES" in column A
    if (sawBlank && colA === 'EQUITIES') break;

    // Hard stops for other sections that appear after the trade table
    if (colA === 'OPTIONS' || colA === 'PROFITS AND LOSSES') break;

    sawBlank = false;

    // Slice row to trade-table columns
    let row = raw.slice(startCol);

    // Skip repeated header rows mid-file
    if (String(row[0] || '').trim().toUpperCase() === 'EXEC TIME') continue;

    if (tosTradesIsRowBlank(row)) continue;

    // Normalize length
    while (row.length < header.length) row.push('');
    if (row.length > header.length) row.length = header.length;

    // Keep only trade-ish rows (BUY/SELL + symbol present)
    const side = idxSide >= 0 ? String(row[idxSide] || '').trim().toUpperCase() : '';
    const symbol = idxSymbol >= 0 ? String(row[idxSymbol] || '').trim() : '';

    if (!/^(BUY|SELL)$/.test(side)) continue;
    if (!symbol) continue;

    // Exec Time validation only if populated (legs may be blank)
    const execStr = String(row[0] || '').trim();
    if (execStr) {
      const dt = tosTradesParseExecTime(execStr);
      if (!dt || isNaN(dt.getTime()) || dt.getFullYear() < 2000) continue;
    }

    // Qty validation only if populated
    if (idxQty >= 0) {
      const qtyStr = String(row[idxQty] || '').trim();
      if (qtyStr) {
        const q = Number(qtyStr.replace(/[$,]/g, ''));
        if (!isFinite(q) || q === 0) continue;
      }
    }

    // Price validation only if populated (optional)
    if (idxPrice >= 0) {
      // No strict skip here; some rows may have weird price text.
      // Leaving it as-is preserves traceability.
    }

    rows.push(row);
  }

  return { header: header, rows: rows };
}

function tosTradesIsRowBlank(row) {
  if (!row || !row.length) return true;
  for (let i = 0; i < row.length; i++) {
    if (String(row[i] ?? '').trim() !== '') return false;
  }
  return true;
}

/**
 * Parses a raw TOS Trades Exec Time CSV string, applies ET→CT +1hr correction,
 * and returns a "yyyy-MM-dd HH:mm:ss" STRING (not a Date object).
 * Storing as text prevents Google Sheets from re-shifting the display time.
 * Returns null if unparseable.
 */
function tosTradesParseExecTimeRaw(v) {
  const t = String(v ?? '').trim();
  if (!t) return null;
  const m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i);
  if (!m) return null;
  let mm = parseInt(m[1], 10), dd = parseInt(m[2], 10), yyyy = parseInt(m[3], 10);
  let HH = parseInt(m[4], 10), MIN = parseInt(m[5], 10), SS = m[6] ? parseInt(m[6], 10) : 0;
  const ap = (m[7] || '').toUpperCase();
  if ([mm, dd, yyyy, HH, MIN, SS].some(n => !isFinite(n))) return null;
  if (yyyy < 100) yyyy += 2000;
  if (ap === 'PM' && HH !== 12) HH += 12;
  if (ap === 'AM' && HH === 12) HH = 0;
  // ET→CT: always +1 hour
  HH += 1;
  if (HH >= 24) { HH -= 24; dd += 1; }
  // Return as formatted string — Sheets will display exactly this value
  const yStr = String(yyyy);
  const mStr = String(mm).padStart(2, '0');
  const dStr = String(dd).padStart(2, '0');
  const hStr = String(HH).padStart(2, '0');
  const minStr = String(MIN).padStart(2, '0');
  const sStr = String(SS).padStart(2, '0');
  return `${yStr}-${mStr}-${dStr} ${hStr}:${minStr}:${sStr}`;
}

/**
 * Robust Exec Time parser for common TOS exports:
 *  - "11/8/2021 7:32"
 *  - "10/27/2021 13:33"
 *  - optional seconds
 *  - optional AM/PM
 */
function tosTradesParseExecTime(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;

  const t = String(v ?? '').trim();
  if (!t) return null;

  const m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i);
  if (!m) return null;

  let mm = parseInt(m[1], 10);
  let dd = parseInt(m[2], 10);
  let yyyy = parseInt(m[3], 10);
  let HH = parseInt(m[4], 10);
  let MIN = parseInt(m[5], 10);
  let SS = m[6] ? parseInt(m[6], 10) : 0;
  const ap = (m[7] || '').toUpperCase();

  if ([mm, dd, yyyy, HH, MIN, SS].some(n => !isFinite(n))) return null;
  if (yyyy < 100) yyyy += 2000;

  if (ap === 'PM' && HH < 12) HH += 12;
  if (ap === 'AM' && HH === 12) HH = 0;

  // REPLACE WITH — apply +1hr in the HH before constructing, then use UTC to prevent Sheets re-shifting:
  HH = HH + 1; // ET→CT: always +1 hour (CT is always 1hr behind ET, year-round)
  // Handle midnight rollover
  if (HH >= 24) { HH -= 24; dd += 1; }
  return new Date(Date.UTC(yyyy, mm - 1, dd, HH, MIN, SS, 0));
}
/** ======================================================================
 *  Top import: CSV folder -> "TOS Top - Combined"
 *  ====================================================================== */
function tosTopImportFromFolder(folderIdPropKey, Account) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  const ctx = importIssuesStart('tosTopImportFromFolder');
  importIssuesSetMetric(ctx, 'Account', Account || '');
  importIssuesSetMetric(ctx, 'OutputSheet', tosConfig.topCombinedSheetName);

  // NEW: prove which ScriptProperties key the caller asked us to use
  importIssuesSetMetric(ctx, 'FolderIdPropKey', folderIdPropKey || '');

  try {
    // NEW: guard - missing prop key means we cannot reliably read a folder id
    if (!folderIdPropKey) {
      importIssuesAdd(ctx, 'ERROR', '', '', 'FolderIdPropKey', '', 'Missing folderIdPropKey argument when calling tosTopImportFromFolder().');
      importIssuesFlush(ctx);
      tosUiAlertSafe('Folder ID prop key was not provided to tosTopImportFromFolder().');
      return;
    }

    const folderId = String(getSetting(folderIdPropKey, '') || '').trim();

    // NEW: log what we actually read
    importIssuesSetMetric(ctx, 'FolderIdUsed', folderId || '');

    if (!folderId) {
      importIssuesAdd(ctx, 'ERROR', '', '', 'FolderId', '', 'Folder ID not set for key: ' + folderIdPropKey);
      importIssuesFlush(ctx);
      tosUiAlertSafe('Folder ID not set. Run the folder setup menu item first.');
      return;
    }

    const folder = DriveApp.getFolderById(folderId);

    // NEW: folder name is the best “sanity check” that we’re in DT vs LT
    const folderName = folder.getName();
    importIssuesSetMetric(ctx, 'FolderNameUsed', folderName);

    // Optional immediate proof popup (controlled by Settings > Toggle TOS Import DEBUG Alerts)
    tosMaybeDebugAlert(
      'tosTopImportFromFolder DEBUG\n\n' +
      'Account: ' + (Account || '') + '\n' +
      'FolderIdPropKey: ' + folderIdPropKey + '\n' +
      'FolderIdUsed: ' + folderId + '\n' +
      'FolderNameUsed: ' + folderName
    );


    const files = folder.getFiles();
    const csvFiles = [];

    while (files.hasNext()) {
      const f = files.next();
      const name = (f.getName() || '').toLowerCase();
      const mime = f.getMimeType();

      // TOS exports are often text/csv but sometimes show as "Microsoft Excel"
      if (name.endsWith('.csv') || mime === MimeType.CSV || mime === MimeType.MICROSOFT_EXCEL) {
        csvFiles.push(f);
      }
    }

    importIssuesSetMetric(ctx, 'FilesFound', csvFiles.length);

    if (!csvFiles.length) {
      importIssuesAdd(ctx, 'WARN', '', '', 'Files', '', 'No CSV-like files found in the folder.');
      importIssuesFlush(ctx);
      tosUiAlertSafe('No CSV-like files found in folder.');
      return;
    }

    // Sort deterministically for repeatability
    csvFiles.sort((a, b) => (a.getName() || '').localeCompare(b.getName() || ''));

    // Handy sanity metric
    importIssuesSetMetric(ctx, 'FirstCsvFile', (csvFiles[0] && csvFiles[0].getName()) ? csvFiles[0].getName() : '');

    let canonicalHeader = null;

    // rowsAll holds fully-normalized row signatures for dedupe
    const rowsAll = []; // { Account, sourceFile, row, dt, timeRaw }

    for (const file of csvFiles) {
      const parsed = tosTopParseOneCsvFile(file);
      if (!parsed || !parsed.rows || !parsed.rows.length) continue;

      importIssuesSetMetric(ctx, 'ParsedRowsTotal', Number(ctx.metrics.ParsedRowsTotal || 0) + parsed.rows.length);

      if (!canonicalHeader) canonicalHeader = parsed.header;

      const idxDate = parsed.header.indexOf('DATE');
      const idxTime = parsed.header.indexOf('TIME');
      if (idxDate < 0 || idxTime < 0) throw new Error('TosTop: DATE/TIME headers not found.');

      for (const r of parsed.rows) {
        // REPLACE WITH:
        const dateRaw = r[idxDate]; // preserve original — may be Date object OR string from CSV
        const dateStr = tosTopNormalizeDateToIso(dateRaw); // always produces yyyy-MM-dd regardless of input type
        const timeRaw = String(r[idxTime] ?? '').trim(); // may include seconds
        // NEW: TIME in the Combined sheet is stored as HHmmss text (second precision).
        // Apply ET→CT 1 hour correction using the trade date as context.
        const timeHHmmssEt = tosTopNormalizeTimeToHHmmss(timeRaw);
        const timeHHmmss = tosEtToCtHHmmss(timeHHmmssEt, dateStr ?? '');
        r[idxTime] = timeHHmmss;

        if (timeRaw && !timeHHmmss) {
          importIssuesAdd(ctx, "BADTIME", file.getName(), "", "TIME", timeRaw, "Could not normalize TIME to HHmmss");
        }

        // NEW: dt is now second-precision (used only for sorting/dedupe inside Combined building).
        const dt = tosTopParseDateTimeMinute(dateStr, timeHHmmss);

        ctx.metrics.CombinedRowsTotal++;

        rowsAll.push({
          Account: Account || '',
          sourceFile: file.getName(),
          row: r,
          dt: dt,
          timeRaw: timeRaw
        });
      }
    }

    if (!canonicalHeader) {
      importIssuesAdd(ctx, 'ERROR', '', '', 'Header', '', 'Could not find TosTop header row in any CSV.');
      importIssuesFlush(ctx);
      tosUiAlertSafe('Could not find the TosTop (Cash Balance) header row in any CSV.');
      return;
    }

    // Dedupe overlap-safe, preserving legitimate duplicates
    // IMPORTANT: key includes Account so LT/DT don’t collapse each other.
    const bucketsByKey = {}; // key -> { rowTemplate, dt, countsByFile, timeRawByFile, Account }

    for (const obj of rowsAll) {
      const r = obj.row;

      const accountKey = String(obj.Account || '');
      const rowKey = r.join('\u0001');
      const key = accountKey + '\u0001' + rowKey;

      if (!bucketsByKey[key]) {
        bucketsByKey[key] = {
          rowTemplate: r.slice(),
          dt: obj.dt || null,
          countsByFile: {},
          timeRawByFile: {},
          Account: accountKey
        };
      }

      const f = obj.sourceFile || '';
      bucketsByKey[key].countsByFile[f] = (bucketsByKey[key].countsByFile[f] || 0) + 1;

      // Keep the first raw time we saw for this key in this file (good for tracing)
      if (bucketsByKey[key].timeRawByFile[f] === undefined) {
        bucketsByKey[key].timeRawByFile[f] = (obj.timeRaw ?? '');
      }
    }

    const deduped = []; // { Account, sourceFile, row, dt, timeRaw }
    const keys = Object.keys(bucketsByKey);
    ctx.metrics.UniqueKeys = keys.length;

    for (let k = 0; k < keys.length; k++) {
      const key = keys[k];
      const bucket = bucketsByKey[key];

      // Pick the file that had the most occurrences of this exact row
      let bestFile = '';
      let bestCount = 0;

      const fileNames = Object.keys(bucket.countsByFile);
      for (let i = 0; i < fileNames.length; i++) {
        const f = fileNames[i];
        const c = bucket.countsByFile[f] || 0;

        if (c > bestCount) {
          bestCount = c;
          bestFile = f;
        } else if (c === bestCount && f < bestFile) {
          bestFile = f;
        }
      }

      const bestTimeRaw = (bucket.timeRawByFile && bucket.timeRawByFile[bestFile] !== undefined)
        ? bucket.timeRawByFile[bestFile]
        : '';

      // Emit the row bestCount times to preserve legitimate duplicates
      for (let n = 0; n < bestCount; n++) {
        deduped.push({
          Account: bucket.Account,
          sourceFile: bestFile,
          row: bucket.rowTemplate.slice(),
          dt: bucket.dt,
          timeRaw: bestTimeRaw
        });
      }
    }

    ctx.metrics.DedupedRows = deduped.length;

    // Sort oldest -> newest
    deduped.sort((a, b) => {
      const da = a.dt, db = b.dt;

      if (da && db) return da - db;
      if (da && !db) return -1;
      if (!da && db) return 1;

      // Tie-breaker: Account then row content
      const al = (a.Account || '').localeCompare(b.Account || '');
      if (al !== 0) return al;

      return a.row.join('|').localeCompare(b.row.join('|'));
    });

    // Output: Account + SourceFile + canonicalHeader + TimeRaw
    const newOut = [];
    newOut.push(['Account', 'SourceFile', ...canonicalHeader, 'TimeRaw']);

    for (const obj of deduped) {
      newOut.push([obj.Account, obj.sourceFile, ...obj.row, (obj.timeRaw ?? '')]);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = tosGetOrCreateSheet(ss, tosConfig.topCombinedSheetName);

    // REPLACE WITH:
    const finalOut = tosMergeAccountLabeledCombined(sh, newOut, Account);

    // Post-merge: re-derive TIME from TimeRaw for every row.
    // Ensures kept rows (other account) are always CT-corrected from their original ET value.
    const fiDateCol = finalOut[0].indexOf('DATE');
    const fiTimeCol = finalOut[0].indexOf('TIME');
    const fiTimeRawCol = finalOut[0].indexOf('TimeRaw');

    if (fiDateCol >= 0 && fiTimeCol >= 0 && fiTimeRawCol >= 0) {
      for (let i = 1; i < finalOut.length; i++) {
        const row = finalOut[i];
        const timeRawVal = row[fiTimeRawCol];
        const dateIso = tosTopNormalizeDateToIso(row[fiDateCol]);
        let timeRawStr;
        if (timeRawVal instanceof Date && !isNaN(timeRawVal.getTime())) {
          const h = String(timeRawVal.getHours()).padStart(2, '0');
          const m = String(timeRawVal.getMinutes()).padStart(2, '0');
          const s = String(timeRawVal.getSeconds()).padStart(2, '0');
          timeRawStr = h + m + s;
        } else {
          timeRawStr = tosTopNormalizeTimeToHHmmss(String(timeRawVal ?? '').trim());
        }
        const timeCorrected = tosEtToCtHHmmss(timeRawStr, dateIso ?? '');
        if (timeCorrected) row[fiTimeCol] = timeCorrected;
      }
    }

    // In-memory sort by DATE then TIME — must happen BEFORE setValues.
    if (finalOut.length > 2) {
      const dataRows = finalOut.slice(1);
      dataRows.sort((a, b) => {
        const dateA = tosTopNormalizeDateToIso(a[fiDateCol]) || '';
        const dateB = tosTopNormalizeDateToIso(b[fiDateCol]) || '';
        if (dateA < dateB) return -1;
        if (dateA > dateB) return 1;
        const timeA = String(a[fiTimeCol] ?? '');
        const timeB = String(b[fiTimeCol] ?? '');
        if (timeA < timeB) return -1;
        if (timeA > timeB) return 1;
        return 0;
      });
      finalOut.splice(1, finalOut.length - 1, ...dataRows);
    }

    sh.clearContents();
    tosFormatHeaderColumnsAsText(sh, finalOut[0], ['TIME'], finalOut.length, 1);
    tosFormatHeaderColumnsAsText(sh, finalOut[0], ['TimeRaw'], finalOut.length, 1);
    sh.getRange(1, 1, finalOut.length, finalOut[0].length).setValues(finalOut);
    tosFormatHeaderColumnsAsText(sh, finalOut[0], ['TIME'], finalOut.length, 1);
    tosFormatHeaderColumnsAsText(sh, finalOut[0], ['TimeRaw'], finalOut.length, 1);

    ctx.metrics.RowsWrittenExclHeader = finalOut.length - 1;
    importIssuesFlush(ctx);

    tosUiAlertSafe(
      'Top import complete.\n' +
      'Account: ' + (Account || '') + '\n' +
      'FolderPropKey: ' + folderIdPropKey + '\n' +
      'FolderName: ' + folderName + '\n' +
      'Files: ' + csvFiles.length + '\n' +
      'Rows written (excl header): ' + (finalOut.length - 1)
    );

  } catch (err) {
    importIssuesAdd(ctx, 'ERROR', '', '', 'Exception', '', String(err && err.stack ? err.stack : err));
    importIssuesFlush(ctx);
    throw err;
  } finally {
    lock.releaseLock();
  }
}


function tosTopParseOneCsvFile(file) {
  let text = file.getBlob().getDataAsString().replace(/^\uFEFF/, '');

  // Parse comma first; if it looks 1-column and contains tabs, re-parse as tab-delimited
  let grid = Utilities.parseCsv(text);
  const looksSingleColumn = grid && grid.length && grid[0].length === 1;
  if (looksSingleColumn && text.indexOf('\t') !== -1) {
    grid = Utilities.parseCsv(text, '\t');
  }

  // Find header row containing DATE + TIME + DESCRIPTION
  let headerRowIdx = -1;
  let startCol = -1;

  for (let r = 0; r < grid.length; r++) {
    const row = grid[r].map(v => (v == null ? '' : String(v).trim()));
    const dateIdx = row.indexOf('DATE');
    const timeIdx = row.indexOf('TIME');
    const descIdx = row.indexOf('DESCRIPTION');

    if (dateIdx !== -1 && timeIdx !== -1 && descIdx !== -1) {
      headerRowIdx = r;
      startCol = dateIdx;
      break;
    }
  }

  if (headerRowIdx < 0) return null;

  const fullHeader = grid[headerRowIdx].map(v => (v == null ? '' : String(v).trim()));
  const header = fullHeader.slice(startCol);

  const idxDesc = header.indexOf('DESCRIPTION');
  const rows = [];

  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    let row = grid[r].map(v => (v == null ? '' : String(v).trim())).slice(startCol);

    while (row.length < header.length) row.push('');
    if (row.length > header.length) row.length = header.length;

    if (tosTradesIsRowBlank(row)) break;

    const desc = (idxDesc >= 0 ? String(row[idxDesc] || '').trim() : '');
    if (desc.toUpperCase() === 'TOTAL') break;

    rows.push(row);
  }

  return { header: header, rows: rows };
}

function tosTopNormalizeTimeToHHmmss(timeStr) {
  // Returns HHmmss text (second precision) — authoritative for the Timestamp rule.
  // Called from tosTopImportFromFolder and pushTosTopCombinedToTosTop.
  // This matches the new raw-data reality where TOS exports now include seconds.

  const raw = String(timeStr ?? "").trim();
  if (!raw) return "";

  // Case 1: colon time like "08:28:12" or "8:28" or "08:28"
  // - We want HHmmss.
  if (raw.indexOf(":") >= 0) {
    const parts = raw.split(":");
    const hh = String(parts[0] ?? "").replace(/\D/g, "").padStart(2, "0");
    const mm = String(parts[1] ?? "").replace(/\D/g, "").padStart(2, "0");
    const ss = String(parts[2] ?? "0").replace(/\D/g, "").padStart(2, "0");
    return (hh + mm + ss).substring(0, 6);
  }

  // Case 2: digits-only forms: "82812", "082812", "0828", "7"
  const digits = raw.replace(/\D/g, "");
  if (!digits) return "";

  // Pad to 6 digits, then take HHmmss.
  // Examples:
  // - "0828"    -> "000828" -> "000828" (not what we want visually)
  // So we instead treat 4-digit as HHmm and append "00".
  if (digits.length <= 4) {
    const hhmm = digits.padStart(4, "0").slice(-4);
    return hhmm + "00";
  }

  // If it already has seconds (5-6+ digits), pad to 6 and take first 6.
  const padded = digits.padStart(6, "0");
  return padded.substring(0, 6);
}


function tosTopParseDateTimeMinute(dateStr, hhmmOrHhmmss) {
  // NOTE: Name is historical.
  // NEW: Accept HHmm OR HHmmss and build a Date with seconds if provided.

  const d = String(dateStr).trim().split("/");
  if (d.length !== 3) return null;

  const mm = parseInt(d[0], 10);
  const dd = parseInt(d[1], 10);
  let yyyy = parseInt(d[2], 10);
  if (![mm, dd, yyyy].every(n => isFinite(n))) return null;
  if (yyyy < 100) yyyy += 2000;

  const rawT = String(hhmmOrHhmmss ?? "").trim();
  const digits = rawT.replace(/\D/g, "");
  if (!digits) return new Date(yyyy, mm - 1, dd, 0, 0, 0, 0);

  // HHmmss preferred; HHmm allowed
  const padded = (digits.length <= 4)
    ? (digits.padStart(4, "0") + "00")
    : digits.padStart(6, "0").substring(0, 6);

  const HH = parseInt(padded.substring(0, 2), 10);
  const MIN = parseInt(padded.substring(2, 4), 10);
  const SS = parseInt(padded.substring(4, 6), 10);

  if (![HH, MIN, SS].every(n => isFinite(n))) return null;
  const dEt = new Date(yyyy, mm - 1, dd, HH, MIN, SS, 0);
  return isNaN(dEt.getTime()) ? dEt : new Date(dEt.getTime() + tosEtToCtOffsetMs(dEt));
}


/** ======================================================================
 *  Push Combined -> TosTrades (with Account)
 *  ====================================================================== */

function pushTosTradesCombinedToTosTrades() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  const ctx = importIssuesStart('pushTosTradesCombinedToTosTrades');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const src = ss.getSheetByName(tosConfig.tradesCombinedSheetName);
    const dst = ss.getSheetByName(tosConfig.tosTradesSheetName);

    if (!src) throw new Error('Missing sheet: ' + tosConfig.tradesCombinedSheetName);
    if (!dst) throw new Error('Missing sheet: ' + tosConfig.tosTradesSheetName);

    ctx.metrics.SourceSheet = tosConfig.tradesCombinedSheetName;
    ctx.metrics.DestSheet = tosConfig.tosTradesSheetName;

    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert(
      'Push TosTrades?',
      'This will REPLACE TosTrades with selected columns from TOS Trades - Combined.\n\nContinue?',
      ui.ButtonSet.OK_CANCEL
    );
    if (resp !== ui.Button.OK) return;

    const values = src.getDataRange().getValues();
    if (values.length < 2) throw new Error('No data found in ' + tosConfig.tradesCombinedSheetName);

    ctx.metrics.SourceRowsReadExclHeader = values.length - 1;

    const headers = values[0].map(h => String(h ?? '').trim());

    function headerIndex(name) {
      const j = headers.indexOf(name);
      if (j === -1) throw new Error('Missing header in Combined Trades: ' + name);
      return j;
    }

    // We REQUIRE Account to exist in the Combined sheet for your new “keep both accounts together” design.
    const idxAccount = headerIndex('Account');

    // Output schema now includes Account first
    const wanted = [
      'Account',
      'Exec Time', 'Spread', 'Side', 'Qty', 'Pos Effect', 'Symbol',
      'Exp', 'Strike', 'Type', 'Price', 'Net Price', 'Order Type'
    ];

    const idx = {};
    for (let i = 0; i < wanted.length; i++) {
      if (wanted[i] === 'Account') {
        idx[wanted[i]] = idxAccount;
      } else {
        idx[wanted[i]] = headerIndex(wanted[i]);
      }
    }

    const out = [wanted];

    for (let r = 1; r < values.length; r++) {
      const row = values[r];

      const hasAny = wanted.some(h => String(row[idx[h]] ?? '').trim() !== '');
      if (!hasAny) {
        ctx.metrics.BlankRowsSkipped++;
        ctx.metrics.TotalRowsSkipped++;
        continue;
      }

      ctx.metrics.SourceNonBlankRows++;

      // ── Exec Time: parse stored "yyyy-MM-dd HHmmss" or "yyyy-MM-dd HH:mm:ss" → real Date ──
      const execRaw = row[idx['Exec Time']];
      let execDate = null;

      if (typeof execRaw === 'string' && execRaw.trim()) {
        const s = execRaw.trim();
        // Accept both "2023-03-09 083044" and "2023-03-09 08:30:44"
        const m = s.match(/^(\d{4})-(\d{2})-(\d{2})[\s\t]+(\d{2}):?(\d{2}):?(\d{2})$/);
        if (m) {
          const [, yyyy, MM, dd, HH, MIN, SS] = m.map(Number);
          execDate = new Date(yyyy, MM - 1, dd, HH, MIN, SS, 0);
          if (isNaN(execDate.getTime()) || execDate.getFullYear() < 2000) execDate = null;
        }
      } else if (execRaw instanceof Date && !isNaN(execRaw.getTime())) {
        // Sheets-coerced Date: read UTC fields to recover original wall-clock time
        const yyyy = execRaw.getUTCFullYear();
        const MM = execRaw.getUTCMonth() + 1;
        const dd = execRaw.getUTCDate();
        const HH = execRaw.getUTCHours();
        const MIN = execRaw.getUTCMinutes();
        const SS = execRaw.getUTCSeconds();
        if (yyyy >= 2000) {
          execDate = new Date(yyyy, MM - 1, dd, HH, MIN, SS, 0);
          if (isNaN(execDate.getTime()) || execDate.getFullYear() < 2000) execDate = null;
        }
      }

      if (!execDate) {
        importIssuesAdd(ctx, 'BAD_EXEC_TIME_TYPE', '', r + 1, 'Exec Time',
          String(execRaw ?? ''),
          `Could not parse Exec Time to Date. type=${typeof execRaw}, value="${execRaw}"`);
        continue;
      }

      const outRow = wanted.map(h => {
        if (h === 'Exec Time') return execDate;
        return row[idx[h]];
      });

      // Defensive: ensure Symbol is always a string in-memory
      const symJ = wanted.indexOf('Symbol');
      if (symJ >= 0) outRow[symJ] = String(outRow[symJ] ?? '').trim();

      out.push(outRow);

    }

    // Transaction-safe write: write new block first, then clear leftovers
    const prevLastRow = Math.max(dst.getLastRow(), 1);
    const prevLastCol = Math.max(dst.getLastColumn(), 1);
    const outRows = out.length;
    const outCols = out[0].length;

    // CRITICAL: preserve leading zeros (CUSIP-like Symbols) by forcing Symbol to text BEFORE setValues.
    tosFormatHeaderColumnAsText(dst, out[0], 'Symbol', outRows, 1);

    // Now write the output values
    dst.getRange(1, 1, outRows, outCols).setValues(out);


    // Format Exec Time column as datetime so the Date object displays readably
    if (outRows > 1) dst.getRange(2, 2, outRows - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');


    if (prevLastRow > outRows) {
      dst.getRange(outRows + 1, 1, prevLastRow - outRows, prevLastCol).clearContent();
    }
    if (prevLastCol > outCols) {
      const rowsToClear = Math.max(prevLastRow, outRows);
      dst.getRange(1, outCols + 1, rowsToClear, prevLastCol - outCols).clearContent();
    }

    // Sort by Exec Time (col 2)
    if (outRows > 2) {
      dst.getRange(2, 1, outRows - 1, outCols).sort({ column: 2, ascending: true });
    }

    ctx.metrics.RowsWrittenExclHeader = outRows - 1;
    importIssuesFlush(ctx);

    ui.alert('Push complete.\nRows written (excluding header): ' + (outRows - 1));

  } catch (err) {
    importIssuesAdd(ctx, 'ERROR', '', '', 'Exception', '', String(err && err.stack ? err.stack : err));
    importIssuesFlush(ctx);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/** ======================================================================
 *  Push Combined -> TosTop (with Account)
 *  ====================================================================== */

function pushTosTopCombinedToTosTop() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  const ctx = importIssuesStart('pushTosTopCombinedToTosTop');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const src = ss.getSheetByName(tosConfig.topCombinedSheetName);
    const dst = ss.getSheetByName(tosConfig.tosTopSheetName);

    if (!src) throw new Error('Missing sheet: ' + tosConfig.topCombinedSheetName);
    if (!dst) throw new Error('Missing sheet: ' + tosConfig.tosTopSheetName);

    ctx.metrics.SourceSheet = tosConfig.topCombinedSheetName;
    ctx.metrics.DestSheet = tosConfig.tosTopSheetName;

    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert(
      'Push TosTop?',
      'This will REPLACE TosTop with selected columns from TOS Top - Combined.\n\nContinue?',
      ui.ButtonSet.OK_CANCEL
    );
    if (resp !== ui.Button.OK) return;

    const values = src.getDataRange().getValues();
    if (values.length < 2) throw new Error('No data found in ' + tosConfig.topCombinedSheetName);

    ctx.metrics.SourceRowsReadExclHeader = values.length - 1;

    const headers = values[0].map(h => String(h ?? '').trim());

    function headerIndex(nameOrNames) {
      const names = Array.isArray(nameOrNames) ? nameOrNames : [nameOrNames];
      for (let i = 0; i < names.length; i++) {
        const j = headers.indexOf(names[i]);
        if (j !== -1) return j;
      }
      throw new Error('Missing header in Combined Top. Tried: ' + names.join(' | '));
    }


    const idxAccount = headerIndex('Account');

    const idx = {
      Account: idxAccount,
      date: headerIndex('DATE'),
      time: headerIndex('TIME'),
      type: headerIndex('TYPE'),
      desc: headerIndex('DESCRIPTION'),
      miscFees: headerIndex('Misc Fees'),
      commFees: headerIndex(['Commissions & Fees', 'Commissions Fees', 'Commissions and Fees']),
      amount: headerIndex('AMOUNT')
    };

    // Output schema now includes Account first (and TIME is HHmmss text now).
    const outHeaders = ["Account", "DATE", "TIME", "TYPE", "DESCRIPTION", "Misc Fees", "Commissions Fees", "AMOUNT"];

    // IMPORTANT: out must be a 2D array for setValues().
    // So we start with [outHeaders], NOT outHeaders.
    const out = [outHeaders];

    // Replace with this version (adds counters + logs only the first 20 blank-date examples so Import Issues doesn’t explode): 1/24 1645
    let topCombinedBlankDateCount = 0;
    let topCombinedBlankDateLogged = 0;

    for (let r = 1; r < values.length; r++) {
      const row = values[r];

      const dateVal = row[idx.date];
      const dateIsBlank = (dateVal === '' || dateVal === null || dateVal === undefined);

      if (dateIsBlank) {
        topCombinedBlankDateCount++;

        if (topCombinedBlankDateLogged < 20) {
          topCombinedBlankDateLogged++;

          importIssuesAdd(
            ctx,
            'TOPCOMBINED_BLANKDATE',
            String(row[idx.sourceFile] ?? ''),
            (r + 1),
            'DATE',
            '',
            JSON.stringify({
              Account: String(row[idx.Account] ?? ''),
              TIME: String(row[idx.time] ?? ''),
              TYPE: String(row[idx.type] ?? ''),
              DESCRIPTION: String(row[idx.desc] ?? '').substring(0, 140)
            })
          );
        }
      }


      const isBlank = row.every(v => String(v ?? "").trim() === "");
      if (isBlank) {
        ctx.metrics.BlankRowsSkipped++;
        ctx.metrics.TotalRowsSkipped++;
        continue;
      }

      const descUpper = String(row[idx.desc] ?? "").trim().toUpperCase();
      if (descUpper === "TOTAL") continue;

      const timeRaw = String(row[idx.time] ?? "").trim();


      const timeHHmmss = tosTopNormalizeTimeToHHmmss(timeRaw);

      if (timeRaw && !timeHHmmss) {
        importIssuesAdd(ctx, "BADTIME", "", (r + 1), "TIME", timeRaw, "TIME could not be normalized to HHmmss");
      }

      const hasAny =
        String(row[idx.date] ?? "").trim() ||
        String(timeHHmmss ?? "").trim() ||
        String(row[idx.type] ?? "").trim() ||
        String(row[idx.desc] ?? "").trim() ||
        String(row[idx.miscFees] ?? "").trim() ||
        String(row[idx.commFees] ?? "").trim() ||
        String(row[idx.amount] ?? "").trim();

      if (!hasAny) continue;

      // CRITICAL: each call to out.push MUST push ONE ARRAY (one row).
      out.push([
        row[idx.Account],
        row[idx.date],
        timeHHmmss,
        row[idx.type],
        row[idx.desc],
        row[idx.miscFees],
        row[idx.commFees],
        row[idx.amount]
      ]);
    }

    // -----------------------------
    // Transaction-safe write sizing
    // -----------------------------
    const prevLastRow = Math.max(dst.getLastRow(), 1);
    const prevLastCol = Math.max(dst.getLastColumn(), 1);

    const outRows = out.length;
    const outCols = out[0].length;

    // Force text columns BEFORE setValues (TIME is HHmmss string)
    tosFormatHeaderColumnsAsText(dst, out[0], ['TIME'], outRows, 1);

    dst.getRange(1, 1, outRows, outCols).setValues(out);


    if (prevLastRow > outRows) {
      dst.getRange(outRows + 1, 1, prevLastRow - outRows, prevLastCol).clearContent();
    }
    if (prevLastCol > outCols) {
      const rowsToClear = Math.max(prevLastRow, outRows);
      dst.getRange(1, outCols + 1, rowsToClear, prevLastCol - outCols).clearContent();
    }

    importIssuesSetMetric(ctx, 'TopCombinedBlankDateCount', topCombinedBlankDateCount);
    importIssuesSetMetric(ctx, 'TopCombinedBlankDateLogged', topCombinedBlankDateLogged);


    ctx.metrics.RowsWrittenExclHeader = outRows - 1;
    importIssuesFlush(ctx);

    ui.alert('Push complete.\nRows written (excluding header): ' + (outRows - 1));

  } catch (err) {
    importIssuesAdd(ctx, 'ERROR', '', '', 'Exception', '', String(err && err.stack ? err.stack : err));
    importIssuesFlush(ctx);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/** ======================================================================
 *  Shared helper: create/get sheet
 *  ====================================================================== */

function tosGetOrCreateSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/**
 * Merge strategy for Combined sheets:
 *
 * - Each import produces a fresh "newOut" table for ONE account.
 * - If combinedReplaceOnlyThatAccount=true, we read the existing Combined sheet,
 *   keep rows whose Account != this account, and replace rows for this account.
 *
 * Why this matters:
 * - It lets you run "Import LT" then "Import DT" and keep both in the same Combined sheet.
 * - It also makes each account import idempotent (re-running LT won’t duplicate LT rows).
 */
function tosMergeAccountLabeledCombined(existingSheet, newOut, Account) {
  if (!tosConfig.combinedReplaceOnlyThatAccount) {
    return newOut;
  }

  const label = String(Account || '');
  if (!label) {
    // If Account is blank, we cannot safely replace-by-account.
    // Fall back to returning only the newOut table.
    return newOut;
  }

  const existingValues = existingSheet.getDataRange().getValues();
  if (!existingValues || existingValues.length < 2) {
    return newOut;
  }

  const existingHeaders = existingValues[0].map(h => String(h ?? '').trim());
  const idxAccount = existingHeaders.indexOf('Account');

  // If existing sheet doesn't have Account (maybe older data), safest is to overwrite.
  if (idxAccount < 0) {
    return newOut;
  }

  // Keep only rows that belong to OTHER accounts.
  const keepRows = [];
  for (let r = 1; r < existingValues.length; r++) {
    const row = existingValues[r];
    const rowLabel = String(row[idxAccount] ?? '').trim();
    if (rowLabel && rowLabel !== label) {
      keepRows.push(row);
    }
  }

  // Now combine:
  // - Use newOut's header as canonical header (fresh)
  // - Append kept rows (other accounts) under it
  // - Append new rows for this account under it
  //
  // NOTE: This assumes both tables share the same header layout.
  // If you later change columns, you'll want a header-mapping merge (not implemented here).

  const finalOut = [];
  finalOut.push(newOut[0]);

  // Existing rows might have different column count. Normalize them.
  for (let i = 0; i < keepRows.length; i++) {
    const row = keepRows[i].slice();
    while (row.length < newOut[0].length) row.push('');
    if (row.length > newOut[0].length) row.length = newOut[0].length;
    finalOut.push(row);
  }

  // New rows (skip header)
  for (let i = 1; i < newOut.length; i++) {
    finalOut.push(newOut[i]);
  }

  return finalOut;
}
