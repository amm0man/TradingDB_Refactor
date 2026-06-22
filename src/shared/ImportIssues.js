/**
 * ImportIssues.js
 *
 * Shared logging and diagnostics service used across the entire trading pipeline.
 *
 * Provides a CTX (context) based system that can write to THREE different issue sheets:
 *   - "Import Issues"
 *   - "Staging Issues"
 *   - "Schwab Mapping Issues"
 *
 * All three sheets use the same 8-column schema and share the same underlying engine.
 * This file also contains the generic date/time validation helper
 * checkMissingDateTimeAndAlert() used after import and mapping steps.
 *
 * Key concepts:
 *   - Start → Add issues/metrics during processing → Flush at the end
 *   - Supports multiple write modes (APPEND_BOTTOM, APPEND_TOP, OVERWRITE)
 *   - RunId ties all logs from one pipeline execution together
 */
/**
 * checkMissingDateTimeAndAlert
 *
 * Generic post-run check: scans any sheet's data rows (starting at rowStart)
 * and alerts if any row that has a Ticker/Symbol or Action is missing a date
 * and/or time value.
 *
 * Supports BOTH sheet schemas automatically:
 *   - "Schwab Mapping"  → looks for "Trade Date" / "Trade Time"  (Phase 2 output)
 *   - "Schwab Import"   → looks for "Date" / "Time"              (Phase 1 output)
 *
 * WHY: Google Sheets occasionally fails to render date/time cell values on the
 * server side, leaving valid rows with blank Date and/or Time even though the
 * timestamp is correct. This check fires a popup so you can re-run immediately
 * rather than carry blank dates downstream.
 *
 * @param {Sheet}  sheet      - The sheet to inspect (Schwab Import or Schwab Mapping)
 * @param {number} rowStart   - First data row (usually 2 for Schwab Import, check your sheet)
 * @param {string} stepName   - Human-readable step name for the alert message
 */
function checkMissingDateTimeAndAlert(sheet, rowStart, stepName) {
  const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headers = rawHeaders.map(h => h.toString().trim().toLowerCase());

  // ── Auto-detect schema ──────────────────────────────────────────────────────
  // Schwab Mapping uses "Trade Date" / "Trade Time"
  // Schwab Import uses "Date" / "Time"
  // We try the Mapping schema first; fall back to the Import schema.
  let dateIdx = headers.indexOf('trade date');
  let timeIdx = headers.indexOf('trade time');
  const useMappingSchema = dateIdx !== -1 || timeIdx !== -1;

  if (!useMappingSchema) {
    dateIdx = headers.indexOf('date');
    timeIdx = headers.indexOf('time');
  }

  // Identifier columns differ by schema too
  // Schwab Mapping has "Ticker"; Schwab Import has "Symbol"
  const tickerIdx  = headers.indexOf('ticker');
  const symbolIdx  = headers.indexOf('symbol');
  const actionIdx  = headers.indexOf('action');

  // If the sheet has neither a date column nor a time column, nothing to check.
  if (dateIdx === -1 && timeIdx === -1) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < rowStart) return;

  const data = sheet.getRange(rowStart, 1, lastRow - rowStart + 1, sheet.getLastColumn()).getValues();

  let missingCount   = 0;
  let firstMissingRow = -1;
  const missingDetails = [];   // collect a few examples for the popup

  for (let i = 0; i < data.length; i++) {
    const row    = data[i];

    // Use whichever identifier column this schema provides
    const ticker = tickerIdx > -1  ? row[tickerIdx].toString().trim()  : '';
    const symbol = symbolIdx > -1  ? row[symbolIdx].toString().trim()  : '';
    const action = actionIdx > -1  ? row[actionIdx].toString().trim()  : '';

    const hasDate = dateIdx > -1
      ? (row[dateIdx] instanceof Date || row[dateIdx].toString().trim() !== '')
      : true;
    const hasTime = timeIdx > -1
      ? (row[timeIdx].toString().trim() !== '')
      : true;

    // Only flag rows that have a Ticker, Symbol, or Action — skip genuinely
    // blank rows that legitimately have no date/time.
    if ((ticker || symbol || action) && (!hasDate || !hasTime)) {
      missingCount++;
      const sheetRow = rowStart + i;
      if (firstMissingRow === -1) firstMissingRow = sheetRow;
      if (missingDetails.length < 5) {
        const id = ticker || symbol || action || '(unknown)';
        missingDetails.push('  Row ' + sheetRow + ': ' + id +
          (!hasDate ? ' — missing Date' : '') +
          (!hasTime ? ' — missing Time' : ''));
      }
    }
  }

  if (missingCount > 0) {
    const schemaLabel = useMappingSchema ? 'Trade Date / Trade Time' : 'Date / Time';
    const exampleLines = missingDetails.length > 0
      ? '\n\nFirst affected rows:\n' + missingDetails.join('\n')
        + (missingCount > 5 ? '\n  ... (' + (missingCount - 5) + ' more)' : '')
      : '';

    SpreadsheetApp.getUi().alert(
      '⚠️ ' + stepName + ': ' + missingCount + ' row(s) are missing ' + schemaLabel + '.' +
      exampleLines +
      '\n\nThis is likely a Google server-side rendering delay.\n' +
      'Please re-run the pipeline immediately to correct.'
    );
  }
}

/**
 * ImportIssues.gs
 *
 * Provides a shared CTX-based logging system for the entire pipeline.
 * THREE target sheets are supported from this one file:
 *   - "Import Issues"         → importIssuesStart()   / importIssuesFlush()
 *   - "Staging Issues"        → stagingIssuesStart()  / stagingIssuesFlush()
 *   - "Schwab Mapping Issues" → mappingIssuesStart()  / mappingIssuesFlush()
 *
 * All rows follow the same 8-column schema:
 *   When | RunId | Step | Kind | SourceRow | Field | Value | Meta
 *
 * How to read it:
 *   - Filter by RunId  → see one complete pipeline run.
 *   - Filter by Step   → isolate a specific function.
 *   - Filter by Kind   → METRIC rows are summaries; WARN/ERROR rows are problems.
 */

// ─────────────────────────────────────────────────────────────────────────────
// SHARED SCHEMA
// ─────────────────────────────────────────────────────────────────────────────
// CHANGED: Removed 'SourceFile'. Renamed 'SheetRow' → 'SourceRow'. Now 8 cols.
const ISSUES_HEADERS = ['When', 'RunId', 'Step', 'Kind', 'SourceRow', 'Field', 'Value', 'Meta'];


// ─────────────────────────────────────────────────────────────────────────────
// CTX LIFECYCLE
// ─────────────────────────────────────────────────────────────────────────────

/**
 * importIssuesStart(stepName)
 * Creates a run context for the "Import Issues" sheet.
 * Call this at the TOP of any function that should log there.
 */
function importIssuesStart(stepName) {
  return _issuesStart(stepName, 'Import Issues');
}

/**
 * stagingIssuesStart(stepName)
 * Creates a run context for the "Staging Issues" sheet.
 * Call this at the TOP of copyMappingToImportByHeaders,
 * validateAndCleanImportToHelperV3, and populateStagingWithBlockLogicV3.
 */
function stagingIssuesStart(stepName) {
  return _issuesStart(stepName, 'Staging Issues');
}

/**
 * mappingIssuesStart(stepName)
 * MOVED HERE from mapSchwabImportByHeadersV3.gs.
 * Creates a run context for the "Schwab Mapping Issues" sheet.
 * Call this at the TOP of mapSchwabImportByHeadersV3() and auditSchwabMappingV3().
 */
function mappingIssuesStart(stepName) {
  return _issuesStart(stepName, 'Schwab Mapping Issues');
}

/**
 * _issuesStart — private generic factory.
 * Creates a context object. The sheetTarget property tells _genericIssuesFlush
 * which sheet to write to when the corresponding Flush function is called.
 *
 * You should NEVER call this directly. Always use the three public *Start functions above.
 */
function _issuesStart(stepName, sheetTarget) {
  const activeRunId = String(getSetting('ACTIVE_IMPORT_RUN_ID', '') || '').trim();
  const runId = activeRunId || new Date().toISOString();
  const step = String(stepName || '').trim();

  return {
    runId: runId,
    step: step,
    sheetTarget: sheetTarget,   // which sheet to write to on Flush
    startedAt: new Date(),
    issues: [],
    metrics: {
      Step: step,
      RunId: runId,
      StartedAt: new Date(),
      EndedAt: '',
      DurationMs: '',
      Success: '',
      ErrorMessage: '',
      ErrorStack: '',

      SourceSheet: '',
      DestSheet: '',
      SourceRowsReadExclHeader: 0,
      RowsWrittenExclHeader: 0,
      RowsSkipped: 0,
      ValidationErrors: 0,
      RADRowsParsed: 0,
      RADParseFailures: 0,
      TickerNormalizations: 0,
      SpreadGroupsBuilt: 0,
      StrikeCollisions: 0,
      BlocksOpened: 0,
      BlocksClosed: 0,
      OpenPositionsAtEnd: 0,
      IssuesCount: 0
    }
  };
}


// ─────────────────────────────────────────────────────────────────────────────
// CTX HELPERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * importIssuesAdd(ctx, kind, sourceRow, field, value, meta)
 *
 * Adds one issue/log row to the CTX's in-memory issue list.
 * Nothing is written to the sheet yet — that happens when you call Flush.
 *
 * Works for ALL three CTX types (Import, Staging, and Mapping).
 *
 *   ctx       — the context object returned by one of the *Start functions
 *   kind      — 'ERROR' | 'WARN' | 'SKIP' | 'INFO'
 *   sourceRow — sheet row number (or a descriptive string like "Sorted row 42")
 *   field     — the column or property name that has the problem
 *   value     — the actual bad value (truncate long strings before passing)
 *   meta      — human-readable explanation of what went wrong / what was done
 */
function importIssuesAdd(ctx, kind, sourceRow, field, value, meta) {
  if (!ctx) throw new Error('importIssuesAdd: ctx was not provided.');
  ctx.issues.push({
    when: new Date(),
    runId: ctx.runId,
    step: ctx.step,
    kind: kind || '',
    sourceRow: sourceRow == null ? '' : String(sourceRow),
    field: field || '',
    value: value == null ? '' : String(value),
    meta: meta == null ? '' : String(meta)
  });
}

/**
 * mappingIssuesAdd(ctx, kind, sourceRow, field, value, meta)
 * MOVED HERE from mapSchwabImportByHeadersV3.gs.
 *
 * Alias of importIssuesAdd(). Exists so that mapSchwabImportByHeadersV3.gs
 * does NOT need any code changes — it can keep calling mappingIssuesAdd()
 * and this function makes that name available. Both do the exact same thing.
 */
function mappingIssuesAdd(ctx, kind, sourceRow, field, value, meta) {
  importIssuesAdd(ctx, kind, sourceRow, field, value, meta);
}

/**
 * importIssuesSetMetric(ctx, key, val)
 * Sets a named metric on the CTX. Works for ALL three CTX types.
 * Call with one of the metric key names from the list in _issuesStart above.
 */
function importIssuesSetMetric(ctx, key, val) {
  if (!ctx || !ctx.metrics) return;
  ctx.metrics[String(key)] = val;
}

/**
 * mappingIssuesSetMetric(ctx, key, val)
 * MOVED HERE from mapSchwabImportByHeadersV3.gs.
 *
 * Alias of importIssuesSetMetric(). Exists so that mapSchwabImportByHeadersV3.gs
 * does NOT need any code changes. Both do the exact same thing.
 */
function mappingIssuesSetMetric(ctx, key, val) {
  importIssuesSetMetric(ctx, key, val);
}

/**
 * importIssuesFlush(ctx)
 * Writes metrics + issues to the "Import Issues" sheet.
 * Call this as the LAST thing before your function returns (including in catch blocks).
 */
function importIssuesFlush(ctx) {
  _genericIssuesFlush(ctx, 'Import Issues');
}

/**
 * stagingIssuesFlush(ctx)
 * Writes metrics + issues to the "Staging Issues" sheet.
 * Call this as the LAST thing in copyMappingToImportByHeaders,
 * validateAndCleanImportToHelperV3, and populateStagingWithBlockLogicV3.
 */
function stagingIssuesFlush(ctx) {
  _genericIssuesFlush(ctx, 'Staging Issues');
}

/**
 * mappingIssuesFlush(ctx)
 * MOVED HERE from mapSchwabImportByHeadersV3.gs.
 * Writes metrics + issues to the "Schwab Mapping Issues" sheet.
 * Call this as the LAST thing in mapSchwabImportByHeadersV3() and auditSchwabMappingV3().
 */
function mappingIssuesFlush(ctx) {
  _genericIssuesFlush(ctx, 'Schwab Mapping Issues');
}


// ─────────────────────────────────────────────────────────────────────────────
// GENERIC ENGINE (private)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * _genericIssuesFlush(ctx, sheetName)
 * Shared engine used by all three Flush functions above.
 * Uses ISSUES_HEADERS (8 cols).
 *
 * Write mode is read from a ScriptProperty key auto-derived from the sheet name
 * by uppercasing and replacing spaces with underscores, then appending _WRITE_MODE:
 *   "Import Issues"         → "IMPORT_ISSUES_WRITE_MODE"
 *   "Staging Issues"        → "STAGING_ISSUES_WRITE_MODE"
 *   "Schwab Mapping Issues" → "SCHWAB_MAPPING_ISSUES_WRITE_MODE"
 *
 * Do NOT call this directly — always call one of the three Flush wrappers above.
 */
function _genericIssuesFlush(ctx, sheetName) {
  if (!ctx) throw new Error('_genericIssuesFlush: ctx was not provided.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  // Ensure header row is correct (rebuilds automatically if schema changed)
  const firstRow = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0] || [];
  const headerCorrect = ISSUES_HEADERS.every((h, i) => String(firstRow[i] || '').trim() === h);
  if (!headerCorrect) {
    sh.clearContents();
    sh.getRange(1, 1, 1, ISSUES_HEADERS.length).setValues([ISSUES_HEADERS]);
  }

  // Determine write mode from ScriptProperties.
  const modeKey = sheetName.toUpperCase().replace(/ /g, '_') + '_WRITE_MODE';
  const modeRaw = String(getSetting(modeKey, 'APPEND_TOP') || '').trim().toUpperCase();
  const mode = ['APPEND_TOP', 'OVERWRITE', 'APPEND_BOTTOM'].includes(modeRaw)
    ? modeRaw : 'APPEND_BOTTOM';

  // Finalize metrics
  const endedAt = new Date();
  ctx.metrics.EndedAt = endedAt;
  ctx.metrics.StartedAt = ctx.startedAt;
  if (ctx.startedAt instanceof Date && !isNaN(ctx.startedAt.getTime())) {
    ctx.metrics.DurationMs = endedAt.getTime() - ctx.startedAt.getTime();
  }
  if (!ctx.metrics.Success) {
    ctx.metrics.Success = (ctx.issues || []).some(
      x => String(x.kind || '').toUpperCase() === 'ERROR'
    ) ? '0' : '1';
  }
  ctx.metrics.IssuesCount = (ctx.issues || []).length;

  // Build output rows
  const out = [];

  // 1) One METRIC row per key
  Object.keys(ctx.metrics).forEach(k => {
    out.push([
      new Date(),
      ctx.runId,
      ctx.step,
      'METRIC',
      '',               // SourceRow is blank for METRIC rows
      k,
      ctx.metrics[k] == null ? '' : String(ctx.metrics[k]),
      ''
    ]);
  });

  // Spacer row to visually separate METRIC rows from issue rows in the sheet
  out.push(['', ctx.runId, ctx.step, '---', '', '', '', '']);

  // 2) Issue rows
  (ctx.issues || []).forEach(x => {
    out.push([
      x.when,
      x.runId,
      x.step,
      x.kind,
      x.sourceRow,
      x.field,
      x.value,
      x.meta
    ]);
  });

  if (out.length) {
    if (mode === 'OVERWRITE') {
      const lastRow = sh.getLastRow();
      if (lastRow > 1) {
        sh.getRange(2, 1, lastRow - 1, Math.max(ISSUES_HEADERS.length, sh.getLastColumn())).clearContent();
      }
      sh.getRange(2, 1, out.length, ISSUES_HEADERS.length).setValues(out);

    } else if (mode === 'APPEND_TOP') {
      sh.insertRowsAfter(1, out.length);
      sh.getRange(2, 1, out.length, ISSUES_HEADERS.length).setValues(out);

    } else {
      // APPEND_BOTTOM (default)
      sh.getRange(sh.getLastRow() + 1, 1, out.length, ISSUES_HEADERS.length).setValues(out);
    }
  }

  // Format the "When" column (col A) as a readable timestamp
  const usedRows = Math.max(1, sh.getLastRow());
  sh.getRange(1, 1, usedRows, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}


// ─────────────────────────────────────────────────────────────────────────────
// WRITE MODE PROMPTS  (callable from your Settings menu)
// ─────────────────────────────────────────────────────────────────────────────

/** Controls "Import Issues" write mode */
function promptSetImportIssuesWriteMode() {
  _promptSetWriteMode('IMPORT_ISSUES_WRITE_MODE', 'Import Issues');
}

/**
 * Controls "Schwab Mapping Issues" write mode.
 * CHANGED: setting key updated to SCHWAB_MAPPING_ISSUES_WRITE_MODE to match
 * the auto-derived key in _genericIssuesFlush (sheet name → key name).
 */
function promptSetMappingIssuesWriteMode() {
  _promptSetWriteMode('SCHWAB_MAPPING_ISSUES_WRITE_MODE', 'Schwab Mapping Issues');
}

/** Controls "Staging Issues" write mode */
function promptSetStagingIssuesWriteMode() {
  _promptSetWriteMode('STAGING_ISSUES_WRITE_MODE', 'Staging Issues');
}

/** Private generic prompt — avoids copy-paste for each sheet */
function _promptSetWriteMode(settingKey, friendlyName) {
  const ui = SpreadsheetApp.getUi();
  const current = String(getSetting(settingKey, 'APPEND_BOTTOM') || '').trim().toUpperCase() || 'APPEND_BOTTOM';

  const msg =
    'Choose how "' + friendlyName + '" is written:\n\n' +
    'APPEND_BOTTOM  = keep history, newest at bottom\n' +
    'APPEND_TOP     = keep history, newest at top (recommended)\n' +
    'OVERWRITE      = show only the latest run\n\n' +
    'Current: ' + current + '\n\n' +
    'Type one of: APPEND_BOTTOM, APPEND_TOP, OVERWRITE';

  const resp = ui.prompt(friendlyName + ' Write Mode', msg, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const val = String(resp.getResponseText() || '').trim().toUpperCase();
  const allowed = { APPEND_BOTTOM: true, APPEND_TOP: true, OVERWRITE: true };

  if (!allowed[val]) {
    ui.alert('Invalid value: "' + val + '"\n\nAllowed: APPEND_BOTTOM, APPEND_TOP, OVERWRITE');
    return;
  }

  setSetting(settingKey, val);
  ui.alert('Saved.\n\n' + settingKey + ' = ' + val);
}
