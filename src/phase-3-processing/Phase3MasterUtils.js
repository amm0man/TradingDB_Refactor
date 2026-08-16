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