/**
 * Phase3MasterUtils.js
 *
 * Phase 3 – Master sheet utilities
 *
 * Sole responsibility:
 *   - Append the finished Staging data to the Master sheet
 *   - Create a timestamped backup of the Master sheet
 *
 * Called by:
 *   - The "DB Tools" menu items
 *   - (optionally) other Phase 3 scripts
 *
 * Note:
 *   clearMasterExceptHeader() already lives in shared/SheetBlanking.js
 */
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
  const accountColIdx =
    staging.getRange(1, 1, 1, numCols).getValues()[0].indexOf("Trade Date") + 1;
  //const accountColIdx = staging.getRange(1, 1, 1, numCols).getValues()[0].indexOf("Trade Date") + 1;
  const stagingData = staging
    .getRange(4, 1, staging.getLastRow() - 3, numCols)
    .getValues()
    .filter(
      (row) => row[accountColIdx - 1] !== "" && row[accountColIdx - 1] !== null,
    );

  // Find first empty row in Master (after header)
  const firstEmptyMasterRow = master.getLastRow() + 1;

  if (stagingData.length) {
    master
      .getRange(firstEmptyMasterRow, 1, stagingData.length, numCols)
      .setValues(stagingData);

    // Staging formats are not copied by setValues.
    // Trade Time is a time-of-day serial; without HH:mm Master shows 12/30/1899.
    const headers = master
      .getRange(1, 1, 1, numCols)
      .getValues()[0]
      .map(function (h) {
        return String(h || "")
          .trim()
          .toLowerCase();
      });
    const tsCol = headers.indexOf("trade time stamp") + 1;
    const tmCol = headers.indexOf("trade time") + 1;
    if (tsCol > 0) {
      master
        .getRange(firstEmptyMasterRow, tsCol, stagingData.length, 1)
        .setNumberFormat("M/d/yyyy HH:mm");
    }
    if (tmCol > 0) {
      master
        .getRange(firstEmptyMasterRow, tmCol, stagingData.length, 1)
        .setNumberFormat("HH:mm");
    }
  }
}
function backupMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(),
    m = ss.getSheetByName("Master");
  const name =
    "Master_Backup_" +
    Utilities.formatDate(
      new Date(),
      ss.getSpreadsheetTimeZone(),
      "yyyyMMdd_HHmmss",
    );
  m.copyTo(ss).setName(name);
  //logAction('BACKUP MASTER',name);
}

/**
 * First clean Master load.
 * Snapshot → blank Master data → append current Staging.
 * Do not use this as the daily incremental path.
 */
function replaceMasterFromStaging() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "Replace Master from Staging",
    "This is the first-load path, not daily append.\n\n" +
      "1) Backup Master (new Master_Backup_… tab)\n" +
      "2) Blank Master data (keeps row 1 via clearMasterExceptHeader)\n" +
      "3) Append Staging rows from row 4 (same as appendStagingToMaster)\n\n" +
      "Cancel if Master already has the history you want to keep.",
    ui.ButtonSet.OK_CANCEL,
  );
  if (resp !== ui.Button.OK) return;

  const t0 = pipelineTimingNow();
  backupMasterSheet();
  pipelineTimingLog("backupMasterSheet", t0);

  const t1 = pipelineTimingNow();
  clearMasterExceptHeader();
  pipelineTimingLog("clearMasterExceptHeader", t1);

  const t2 = pipelineTimingNow();
  appendStagingToMaster();
  pipelineTimingLog("appendStagingToMaster after clear (replace)", t2);

  uiAlertSafe(
    "Master replaced from Staging.\n" +
      "A Master_Backup_… tab was created.\n" +
      "Timings are on DB_log (Action = TIMING).",
  );
}

/**
 * Incremental Master write.
 *
 * Uses collectIncrementalCandidates_() (same cutoff + fingerprint as
 * Inspect incremental delta). Appends Incremental Seeded Preview data
 * rows only. Does not call appendStagingToMaster. Does not write
 * Helper or Staging.
 *
 * On the 9/1/2026 16:41:18 freeze candidates = 0, so this is a no-op.
 * When candidates > 0 later, leftover replay rows on the preview
 * sheet are skipped unless their fingerprint is a live candidate
 * and is not already on Master.
 */
function appendIncrementalPreviewToMaster() {
  const t0 = pipelineTimingNow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const tz = ss.getSpreadsheetTimeZone();

  const collected = collectIncrementalCandidates_();
  if (!collected) return;

  const lastTxt = Utilities.formatDate(
    collected.lastMasterTs,
    tz,
    "M/d/yyyy HH:mm:ss",
  );

  const master = ss.getSheetByName("Master");
  const staging = ss.getSheetByName("Staging");
  const masterLastBefore = master ? master.getLastRow() : "";
  const stagingLastBefore = staging ? staging.getLastRow() : "";

  if (collected.candidates === 0) {
    pipelineTimingLog(
      "appendIncrementalPreviewToMaster",
      t0,
      "candidates=0 no-op lastTs=" + lastTxt,
    );
    uiAlertSafe(
      "Incremental append skipped — no new Helper rows.\n\n" +
        "Master last ts: " +
        lastTxt +
        "\n" +
        "Candidates: 0\n" +
        "AFTER new: " +
        collected.afterNew +
        "\n" +
        "AT_LAST_TS unmatched: " +
        collected.atUnmatched +
        "\n" +
        "Master last row before/after: " +
        masterLastBefore +
        " / " +
        masterLastBefore +
        "\n" +
        "Staging last row before/after: " +
        stagingLastBefore +
        " / " +
        stagingLastBefore +
        "\n\n" +
        "Candidates 0 means Master already has these Helper rows.\n" +
        "Did not write Helper / Staging / Master.\n" +
        "Did not call appendStagingToMaster.",
    );
    return;
  }

  const preview = ss.getSheetByName("Incremental Seeded Preview");
  if (!preview) {
    uiAlertSafe(
      "Incremental Seeded Preview not found.\n" +
        "Run Preview incremental Step 4 first.",
    );
    return;
  }

  const previewLast = preview.getLastRow();
  const previewLastCol = preview.getLastColumn();
  if (previewLast < 4) {
    uiAlertSafe(
      "Incremental Seeded Preview has no data rows.\n" +
        "Run Preview incremental Step 4 first.\n" +
        "Candidates this run: " +
        collected.candidates,
    );
    return;
  }

  function headerMap_(grid) {
    const col = {};
    grid[0].forEach(function (h, i) {
      const norm = String(h || "")
        .trim()
        .toLowerCase();
      if (norm) col[norm] = i;
    });
    return col;
  }

  const candidateFp = {};
  for (let i = 0; i < collected.candidateRows.length; i++) {
    candidateFp[collected.candidateRows[i].fingerprint] = true;
  }

  const mLast = master.getLastRow();
  const mLastCol = master.getLastColumn();
  const mGrid = master.getRange(1, 1, mLast, mLastCol).getValues();
  const mCol = headerMap_(mGrid);
  const masterFp = {};
  for (let r = 1; r < mGrid.length; r++) {
    const acct = String(mGrid[r][mCol["account"] || 0] || "").trim();
    if (!acct) continue;
    masterFp[incrementalRowFingerprint_(mGrid[r], mCol)] = r + 1;
  }

  const pGrid = preview.getRange(1, 1, previewLast, previewLastCol).getValues();
  const pCol = headerMap_(pGrid);
  const writeCols = Math.min(mLastCol, previewLastCol);
  const toAppend = [];
  let skippedAlreadyOnMaster = 0;
  let skippedNotCandidate = 0;
  let skippedNoAccount = 0;

  // Preview rows 2–3 are Helper spacers. Data starts at row 4 (index 3).
  for (let r = 3; r < pGrid.length; r++) {
    const row = pGrid[r];
    const acct = String(
      pCol["account"] !== undefined ? row[pCol["account"]] : "",
    ).trim();
    if (!acct) {
      skippedNoAccount++;
      continue;
    }
    const fp = incrementalRowFingerprint_(row, pCol);
    if (masterFp[fp]) {
      skippedAlreadyOnMaster++;
      continue;
    }
    if (!candidateFp[fp]) {
      skippedNotCandidate++;
      continue;
    }
    toAppend.push(row.slice(0, writeCols));
  }

  if (!toAppend.length) {
    pipelineTimingLog(
      "appendIncrementalPreviewToMaster",
      t0,
      "candidates=" +
        collected.candidates +
        " append=0 already=" +
        skippedAlreadyOnMaster +
        " notCand=" +
        skippedNotCandidate,
    );
    uiAlertSafe(
      "Incremental append skipped — nothing new to write.\n\n" +
        "Master last ts: " +
        lastTxt +
        "\n" +
        "Candidates: " +
        collected.candidates +
        "\n" +
        "Preview data rows skipped (already on Master): " +
        skippedAlreadyOnMaster +
        "\n" +
        "Preview data rows skipped (not a live candidate): " +
        skippedNotCandidate +
        "\n" +
        "Master last row before/after: " +
        masterLastBefore +
        " / " +
        masterLastBefore +
        "\n\n" +
        "Did not write Master.",
    );
    return;
  }

  const resp = ui.alert(
    "Append incremental preview → Master",
    "Master last ts: " +
      lastTxt +
      "\n" +
      "Live candidates: " +
      collected.candidates +
      "\n" +
      "Rows that will append: " +
      toAppend.length +
      "\n" +
      "Skipped already on Master: " +
      skippedAlreadyOnMaster +
      "\n" +
      "Skipped leftover / not a candidate: " +
      skippedNotCandidate +
      "\n\n" +
      "This backs up Master, then appends those preview rows only.\n" +
      "It does not write Staging or call appendStagingToMaster.\n\n" +
      "Cancel if those counts look wrong.",
    ui.ButtonSet.OK_CANCEL,
  );
  if (resp !== ui.Button.OK) {
    pipelineTimingLog(
      "appendIncrementalPreviewToMaster",
      t0,
      "cancelled candidates=" + collected.candidates,
    );
    return;
  }

  backupMasterSheet();

  const firstEmpty = master.getLastRow() + 1;
  master
    .getRange(firstEmpty, 1, toAppend.length, writeCols)
    .setValues(toAppend);

  const headers = master
    .getRange(1, 1, 1, writeCols)
    .getValues()[0]
    .map(function (h) {
      return String(h || "")
        .trim()
        .toLowerCase();
    });
  const tsCol = headers.indexOf("trade time stamp") + 1;
  const tmCol = headers.indexOf("trade time") + 1;
  if (tsCol > 0) {
    master
      .getRange(firstEmpty, tsCol, toAppend.length, 1)
      .setNumberFormat("M/d/yyyy HH:mm");
  }
  if (tmCol > 0) {
    master
      .getRange(firstEmpty, tmCol, toAppend.length, 1)
      .setNumberFormat("HH:mm");
  }

  const masterLastAfter = master.getLastRow();
  const stagingLastAfter = staging ? staging.getLastRow() : "";

  pipelineTimingLog(
    "appendIncrementalPreviewToMaster",
    t0,
    "appended=" +
      toAppend.length +
      " master=" +
      masterLastBefore +
      "→" +
      masterLastAfter,
  );

  uiAlertSafe(
    "Incremental append OK.\n\n" +
      "Master last ts: " +
      lastTxt +
      "\n" +
      "Rows appended: " +
      toAppend.length +
      "\n" +
      "Master last row before/after: " +
      masterLastBefore +
      " / " +
      masterLastAfter +
      "\n" +
      "Staging last row before/after: " +
      stagingLastBefore +
      " / " +
      stagingLastAfter +
      "\n" +
      "A Master_Backup_… tab was created.\n" +
      "Did not write Helper / Staging.\n" +
      "Did not call appendStagingToMaster.",
  );
}
