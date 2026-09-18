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
