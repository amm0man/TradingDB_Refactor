/**
 * Phase3Processing.js
 *
 * Phase 3 – Thin Orchestrator
 *
 * Sole responsibility:
 *   Coordinate the three Phase 3 steps and provide a couple of tiny shared helpers.
 *
 * Pipeline (run by refreshAllScripts):
 *   1. copyMappingToImportByHeaders()     → Schwab Mapping → Import
 *      (lives in Phase3Step1_CopyMapping.js)
 *   2. validateAndCleanImportToHelperV3() → Import → Helper
 *      (lives in Phase3Step2_ValidateClean.js)
 *   3. populateStagingWithBlockLogicV3()  → Helper → Staging
 *      (lives in Phase3BlockLogic.js)
 *
 * Related files in this folder:
 *   - Phase3Step1_CopyMapping.js
 *   - Phase3Step2_ValidateClean.js
 *   - Phase3BlockLogic.js
 *   - Phase3MasterUtils.js
 *
 * Other related files:
 *   - Menu.js
 *   - shared/SheetBlanking.js          (clearMasterExceptHeader)
 *   - shared/ImportIssues.js
 *   - shared/Helpers.js
 */

// =========================================================================
// SMALL SHARED HELPERS
// =========================================================================

function logAction(action, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("DB_log");
  if (!logSheet) {
    logSheet = ss.insertSheet("DB_log");
    logSheet.appendRow(["Timestamp", "Action", "Details"]);
  }
  logSheet.appendRow([
    Utilities.formatDate(
      new Date(),
      ss.getSpreadsheetTimeZone(),
      "yyyy-MM-dd HH:mm:ss",
    ),
    action,
    details,
  ]);
}

function ensureValidationErrorSheet() {
  let sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validation Errors");
  if (!sheet) {
    sheet =
      SpreadsheetApp.getActiveSpreadsheet().insertSheet("Validation Errors");
    sheet.appendRow(["Row", "Column", "Error", "Suggested Fix"]);
  }
  return sheet;
}

// =========================================================================
// ORCHESTRATOR
//   Runs the three Phase 3 steps under one shared RunId so all Staging Issues
//   rows from a full refresh can be filtered together.
// =========================================================================

function refreshAllScripts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const runId = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss zzz");
  setSetting("ACTIVE_IMPORT_RUN_ID", runId);

  try {
    copyMappingToImportByHeaders(); // Step 1 → Phase3Step1_CopyMapping.js
    validateAndCleanImportToHelperV3(); // Step 2 → Phase3Step2_ValidateClean.js
    populateStagingWithBlockLogicV3(); // Step 3 → Phase3BlockLogic.js

    uiAlertSafe(
      "✅ Full refresh complete!\n" +
        "RunId: " +
        runId +
        "\n\n" +
        "Check the 'Staging Issues' sheet and filter by this RunId to review.",
    );
  } catch (e) {
    uiAlertSafe(
      "❌ Pipeline error: " + e.message + "\nCheck Staging Issues sheet.",
    );
  } finally {
    // Always clear the active RunId — even if an error occurred
    setSetting("ACTIVE_IMPORT_RUN_ID", "");
  }
}
