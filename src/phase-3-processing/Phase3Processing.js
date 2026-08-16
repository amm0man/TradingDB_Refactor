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
// STEP 1: Copy Mapping → Import has been moved
//   → see Phase3Step1_CopyMapping.js
//   (copyMappingToImportByHeaders)
// =========================================================================

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



