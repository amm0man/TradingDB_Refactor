/**
 * Menu.js
 *
 * Builds the custom "DB Tools" menu that appears when the spreadsheet opens.
 *
 * This file contains only the menu definition (onOpen).
 * All actual processing logic lives in the phase folders and in DBTools.js.
 *
 * Menu structure:
 *   DB Tools
 *   ├── Manual Entry
 *   ├── Raw Data for Sorting   (TOS / Schwab import pipeline)
 *   ├── Schwab Actions         (Phase 1 + Phase 2 + settings)
 *   └── Data Actions           (Phase 3 + Master utilities)
 *
 * Related files:
 *   - DBTools.js                 (Phase 3 processing + refreshAllScripts)
 *   - phase-1-import/*           (buildUnifiedImportV3, TOS import helpers)
 *   - phase-2-mapping/*          (mapSchwabImportByHeadersV3)
 *   - shared/*                   (SheetBlanking, SettingsService, ImportIssues)
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Manual Entry submenu
  const manualEntryMenu = ui
    .createMenu("Manual Entry")
    .addItem("Send Entry to Import", "sendEntryToImport");

  // Raw Data for Sorting submenu (TOS ingest)
  const TosSchwabImport = ui
    .createMenu("Raw Data for Sorting")
    .addItem("Set folder (LT): TosTop + TosTrades", "tosSetCsvFolderIdLT")
    .addItem("Set folder (DT): TosTop + TosTrades", "tosSetCsvFolderIdDT")
    .addSeparator()

    // Import raw CSVs -> Combined
    .addItem(
      "Import TosTrades Current Account → TOS Trades - Combined",
      "tosTradesImportFromFolderCurrentAccount",
    )
    .addItem(
      "Import TosTop Current Account → TOS Top - Combined",
      "tosTopImportFromFolderCurrentAccount",
    )
    .addSeparator()

    // Import BOTH accounts -> Combined
    .addItem(
      "Import TosTrades BOTH Accounts → TOS Trades - Combined",
      "tosTradesImportFromFolderBothAccounts",
    )
    .addItem(
      "Import TosTop BOTH Accounts → TOS Top - Combined",
      "tosTopImportFromFolderBothAccounts",
    )
    .addSeparator()

    // Blank/reset staging
    .addItem("Blank TOS Trades - Combined", "tosBlankTOSTradesCombined")
    .addItem("Blank TOS Trades", "tosBlankTosTrades")
    .addItem("Blank TOS Top - Combined", "tosBlankTOSTopCombined")
    .addItem("Blank TOS Top", "tosBlankTosTop")
    .addItem("Blank all TOS sheets", "tosBlankALLTOSSheets")
    .addSeparator()

    // Push Combined -> TosTop / TosTrades
    .addItem(
      "Push: TOS Trades - Combined → TosTrades",
      "pushTosTradesCombinedToTosTrades",
    )
    .addItem("Push: TOS Top - Combined → TosTop", "pushTosTopCombinedToTosTop")
    .addItem("Push BOTH (Top + Trades)", "pushTosCombinedToBoth")
    .addSeparator()

    // One button end-to-end raw -> Schwab Import
    .addItem(
      "Run FULL (BOTH Accounts): CSV → Combined → TosTop/TosTrades → Schwab Import",
      "tosRunFullTosToSchwabImportBothAccounts",
    );

  const settingsMenu = ui
    .createMenu("Settings")
    .addItem("Set Account Mode DT / LT", `promptSetAccountMode`)
    .addItem(
      "Toggle TOS Import DEBUG Alerts",
      `promptToggleTosImportDebugAlerts`,
    )
    .addSeparator()
    .addItem("Set Import Issues Write Mode", `promptSetImportIssuesWriteMode`)
    .addItem("Set Mapping Issues Write Mode", `promptSetMappingIssuesWriteMode`)
    .addItem("Set Staging Issues Write Mode", `promptSetStagingIssuesWriteMode`)
    .addItem("Show Current Settings", `showCurrentSettingsDialog`);

  // Schwab Actions submenu
  const schwabActionsMenu = ui
    .createMenu("Schwab Actions")
    .addItem("Build Unified Import V3", "buildUnifiedImportV3")
    .addItem("Run Schwab Mapping Script", "mapSchwabImportByHeadersV3")
    .addItem(
      "2b. Audit Schwab Mapping (pre-Phase-3 gate)",
      "auditSchwabMappingV3",
    )
    .addItem("Move Schwab Mapping to Import", "copyMappingToImportByHeaders")
    .addItem("Blank Schwab Import", "clearSchwabImportExceptHeader")
    .addItem("Blank Schwab Mapping", "clearSchwabMappingExceptHeader")
    .addItem("Blank all Schwab Sheets", "blankAllSchwabSheets")
    .addSeparator()
    .addSubMenu(settingsMenu);

  // Data Actions submenu
  const dataActionsMenu = ui
    .createMenu("Data Actions")
    .addItem("Schwab Mapping to Import and run scripts", "refreshAllScripts")
    .addItem(
      "Run Audit on Staging - Final Check before push to Master",
      "auditPipelineIntegrity",
    )
    .addItem("Push Staging → Master (Append)", "appendStagingToMaster")
    .addItem("Backup Master (Snapshot)", "backupMasterSheet")
    .addSeparator()
    .addItem("Blank Import", "blankImport")
    .addItem("Blank Helper", "blankHelper")
    .addItem("Blank Staging", "blankStaging")
    .addItem("Blank ALL Prep Sheets", "blankAllPrepSheets")
    .addSeparator()
    .addItem("Blank Master", "clearMasterExceptHeader");

  // Top-level menu
  ui.createMenu("DB Tools")
    .addSubMenu(manualEntryMenu)
    .addSubMenu(TosSchwabImport)
    .addSubMenu(schwabActionsMenu)
    .addSubMenu(dataActionsMenu)
    .addToUi();
}
