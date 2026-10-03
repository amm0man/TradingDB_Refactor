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
 * Weekly step 11. Candidates 0 → no-op (Master already has these rows).
 * Candidates > 0 → backup Master, then append Incremental Seeded Preview
 * data rows whose fingerprint is a live candidate and is not already
 * on Master. Does not call appendStagingToMaster.
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

  // Stop before backup/append if a date the seed needs is blank.
  // Trade Time Stamp blank breaks the next candidate fingerprint.
  // Option Expiration blank breaks option block keys.
  const stampIdx = pCol["trade time stamp"];
  const expIdx = pCol["option expiration"];
  const cpIdx = pCol["call/put"];
  const strikeIdx = pCol["option strike"];
  let blankDateRows = 0;
  const blankDateSamples = [];
  for (let i = 0; i < toAppend.length; i++) {
    const row = toAppend[i];
    const stamp = stampIdx !== undefined ? row[stampIdx] : "";
    const hasStamp = stamp instanceof Date || String(stamp || "").trim() !== "";
    const cp = cpIdx !== undefined ? String(row[cpIdx] || "").trim() : "";
    const strike =
      strikeIdx !== undefined ? String(row[strikeIdx] || "").trim() : "";
    const isOption = cp !== "" || strike !== "";
    const exp = expIdx !== undefined ? row[expIdx] : "";
    const hasExp =
      !isOption || exp instanceof Date || String(exp || "").trim() !== "";
    if (!hasStamp || !hasExp) {
      blankDateRows++;
      if (blankDateSamples.length < 5) {
        const acct = pCol["account"] !== undefined ? row[pCol["account"]] : "";
        const ticker = pCol["ticker"] !== undefined ? row[pCol["ticker"]] : "";
        blankDateSamples.push(
          String(acct) +
            " " +
            String(ticker) +
            (!hasStamp ? " — missing Trade Time Stamp" : "") +
            (!hasExp ? " — missing Option Expiration" : ""),
        );
      }
    }
  }
  if (blankDateRows > 0) {
    uiAlertSafe(
      "Incremental append stopped. Did not backup Master.\n\n" +
        blankDateRows +
        " preview row(s) are missing a date the seed needs.\n" +
        blankDateSamples.join("\n") +
        "\n\nFix Incremental Seeded Preview, then run step 11 again.",
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

/**
 * One-shot Master repair for the 9/28 incremental append.
 *
 * URA 50C and XE 30C covers were appended as Short Call
 * (DT-*-SC-270115-TG001, blank spread). Seed-aware Step 4 now
 * attaches them on Incremental Seeded Preview. This copies the
 * Preview block fields onto the matching Master rows.
 *
 * Does not append. Does not write Helper / Staging / Preview.
 * Does not call appendStagingToMaster or refreshAllScripts.
 *
 * Expected: 3 URA BTC + 2 XE BTC = 5 Master rows.
 */
function repairMasterUraXeCoversFromPreview() {
  const t0 = pipelineTimingNow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const master = ss.getSheetByName("Master");
  const preview = ss.getSheetByName("Incremental Seeded Preview");
  const staging = ss.getSheetByName("Staging");
  if (!master) {
    uiAlertSafe('Sheet "Master" not found.');
    return;
  }
  if (!preview) {
    uiAlertSafe(
      "Incremental Seeded Preview not found.\n" +
        "Run Debug: replay Helper slice (pre-increment seed) first.",
    );
    return;
  }

  const masterLastBefore = master.getLastRow();
  const stagingLastBefore = staging ? staging.getLastRow() : "";
  const previewLast = preview.getLastRow();
  const previewLastCol = preview.getLastColumn();
  if (previewLast < 4) {
    uiAlertSafe(
      "Incremental Seeded Preview has no data rows.\n" +
        "Run Debug: replay Helper slice (pre-increment seed) first.",
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

  function cell_(row, col, name) {
    return col[name] !== undefined ? row[col[name]] : "";
  }

  function isTargetTicker_(ticker) {
    const t = String(ticker || "")
      .trim()
      .toUpperCase();
    return t === "URA" || t === "XE";
  }

  function isCloseAction_(action) {
    const a = String(action || "")
      .trim()
      .toUpperCase();
    return a.indexOf("TO CLOSE") !== -1 || a === "RAD";
  }

  function previewLooksAttached_(spreadId, posId) {
    const s = String(spreadId || "")
      .trim()
      .toUpperCase();
    const p = String(posId || "")
      .trim()
      .toUpperCase();
    if (s.indexOf("SPREAD-URA-CDS-") === 0) return true;
    if (s.indexOf("SPREAD-XE-CDS-") === 0) return true;
    if (p.indexOf("SPREAD-URA-CDS-") === 0) return true;
    if (p.indexOf("SPREAD-XE-CDS-") === 0) return true;
    return false;
  }

  function masterNeedsRepair_(spreadId, posId, tradeGroupId) {
    const s = String(spreadId || "").trim();
    const ids = (
      String(posId || "") +
      " " +
      String(tradeGroupId || "")
    ).toUpperCase();
    if (!s) return true;
    if (ids.indexOf("-SC-") !== -1) return true;
    if (ids.indexOf("-CDS-") === -1) return true;
    return false;
  }

  const fields = [
    "spread group id",
    "position id",
    "trade group id",
    "strategy type",
    "trade type",
    "running position quantity",
    "block start flag",
    "block close flag/p&l",
    "block number",
  ];

  const mLast = master.getLastRow();
  const mLastCol = master.getLastColumn();
  const mGrid = master.getRange(1, 1, mLast, mLastCol).getValues();
  const mCol = headerMap_(mGrid);
  const pGrid = preview.getRange(1, 1, previewLast, previewLastCol).getValues();
  const pCol = headerMap_(pGrid);

  const need = ["account", "ticker", "action"].concat(fields);
  for (let i = 0; i < need.length; i++) {
    if (mCol[need[i]] === undefined || pCol[need[i]] === undefined) {
      uiAlertSafe(
        "Master or Incremental Seeded Preview is missing header:\n" + need[i],
      );
      return;
    }
  }

  const masterByFp = {};
  const masterFpDup = {};
  for (let r = 1; r < mGrid.length; r++) {
    const acct = String(cell_(mGrid[r], mCol, "account") || "").trim();
    if (!acct) continue;
    const fp = incrementalRowFingerprint_(mGrid[r], mCol);
    if (masterByFp[fp]) masterFpDup[fp] = true;
    masterByFp[fp] = r + 1;
  }

  const hits = [];
  let previewTargets = 0;
  let skippedAlreadyGood = 0;
  let skippedNoMaster = 0;
  let skippedDupFp = 0;

  for (let r = 3; r < pGrid.length; r++) {
    const prow = pGrid[r];
    const ticker = cell_(prow, pCol, "ticker");
    if (!isTargetTicker_(ticker)) continue;
    if (!isCloseAction_(cell_(prow, pCol, "action"))) continue;
    const pSpread = cell_(prow, pCol, "spread group id");
    const pPos = cell_(prow, pCol, "position id");
    if (!previewLooksAttached_(pSpread, pPos)) continue;
    previewTargets++;

    const fp = incrementalRowFingerprint_(prow, pCol);
    if (masterFpDup[fp]) {
      skippedDupFp++;
      continue;
    }
    const mRowNum = masterByFp[fp];
    if (!mRowNum) {
      skippedNoMaster++;
      continue;
    }
    const mrow = mGrid[mRowNum - 1];
    if (
      !masterNeedsRepair_(
        cell_(mrow, mCol, "spread group id"),
        cell_(mrow, mCol, "position id"),
        cell_(mrow, mCol, "trade group id"),
      )
    ) {
      skippedAlreadyGood++;
      continue;
    }

    const changes = [];
    for (let f = 0; f < fields.length; f++) {
      const name = fields[f];
      const fromVal = mrow[mCol[name]];
      const toVal = prow[pCol[name]];
      const fromTxt = String(fromVal == null ? "" : fromVal).trim();
      const toTxt = String(toVal == null ? "" : toVal).trim();
      if (fromTxt !== toTxt) {
        changes.push({
          field: name,
          fromVal: fromVal,
          toVal: toVal,
          fromTxt: fromTxt,
          toTxt: toTxt,
        });
      }
    }
    if (!changes.length) {
      skippedAlreadyGood++;
      continue;
    }

    hits.push({
      masterRow: mRowNum,
      previewRow: r + 1,
      ticker: String(ticker || "").toUpperCase(),
      action: String(cell_(prow, pCol, "action") || ""),
      fp: fp,
      changes: changes,
    });
  }

  if (!hits.length) {
    uiAlertSafe(
      "URA/XE Master repair — nothing to write.\n\n" +
        "Preview attached covers scanned: " +
        previewTargets +
        "\n" +
        "Already matched Preview: " +
        skippedAlreadyGood +
        "\n" +
        "Preview cover with no Master fingerprint: " +
        skippedNoMaster +
        "\n" +
        "Duplicate Master fingerprint (skipped): " +
        skippedDupFp +
        "\n" +
        "Master last row: " +
        masterLastBefore +
        "\n\n" +
        "Did not write Master.",
    );
    return;
  }

  let plan = "Repair " + hits.length + " Master row(s) from Preview.\n\n";
  if (hits.length !== 5) {
    plan += "NOTE: expected 5 covers (3 URA + 2 XE). Review the list.\n\n";
  }
  for (let i = 0; i < hits.length; i++) {
    const h = hits[i];
    plan +=
      "Master R" +
      h.masterRow +
      " " +
      h.ticker +
      " " +
      h.action +
      " (" +
      h.changes.length +
      " fields)\n";
  }
  plan +=
    "\nOK creates Master_Backup_… then overwrites those cells only.\n" +
    "Master last row stays " +
    masterLastBefore +
    ".\n" +
    "Cancel leaves Master unchanged.";

  const resp = ui.alert(
    "Repair Master URA/XE covers",
    plan,
    ui.ButtonSet.OK_CANCEL,
  );
  if (resp !== ui.Button.OK) return;

  backupMasterSheet();

  const log = [
    [
      "Master Row",
      "Preview Row",
      "Ticker",
      "Action",
      "Field",
      "Master Before",
      "Preview",
    ],
  ];
  let cells = 0;
  for (let i = 0; i < hits.length; i++) {
    const h = hits[i];
    for (let c = 0; c < h.changes.length; c++) {
      const ch = h.changes[c];
      const col1 = mCol[ch.field] + 1;
      master.getRange(h.masterRow, col1).setValue(ch.toVal);
      cells++;
      log.push([
        h.masterRow,
        h.previewRow,
        h.ticker,
        h.action,
        ch.field,
        ch.fromTxt,
        ch.toTxt,
      ]);
    }
  }

  let logSh = ss.getSheetByName("URA XE Master Repair Log");
  if (!logSh) logSh = ss.insertSheet("URA XE Master Repair Log");
  logSh.clearContents();
  logSh.getRange(1, 1, log.length, log[0].length).setValues(log);
  logSh.setFrozenRows(1);

  const masterLastAfter = master.getLastRow();
  const stagingLastAfter = staging ? staging.getLastRow() : "";

  pipelineTimingLog(
    "repairMasterUraXeCoversFromPreview",
    t0,
    "rows=" +
      hits.length +
      " cells=" +
      cells +
      " masterLast=" +
      masterLastAfter,
  );

  uiAlertSafe(
    "URA/XE Master repair wrote " +
      hits.length +
      " row(s), " +
      cells +
      " cell(s).\n\n" +
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
      "Preview last row (unchanged): " +
      previewLast +
      "\n\n" +
      "Open URA XE Master Repair Log for before/after.\n" +
      "Then run Inspect seed blocks from Master.\n" +
      "Want URA CDS live unit 3, XE CDS live unit 2, no live SC keys.\n" +
      "Did not append. Did not write Helper / Staging.",
  );
}
