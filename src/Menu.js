function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // --- Incremental: weekly / monthly path (use this) ---
  const incrementalMenu = ui
    .createMenu("Incremental (weekly / monthly)")
    .addItem(
      "Inspect incremental delta (Helper vs Master)",
      "inspectIncrementalDeltaFromMaster",
    )
    .addItem(
      "Preview incremental candidates (Helper vs Master)",
      "previewIncrementalCandidatesFromHelper",
    )
    .addItem(
      "Incremental dry run (0-row proof)",
      "runIncrementalFromHelperDryRun",
    )
    .addItem(
      "Preview incremental seeded output (headers only)",
      "previewIncrementalSeededOutputFromHelper",
    )
    .addItem(
      "Preview incremental Step 4 (preview sheet only)",
      "previewIncrementalStep4FromHelper",
    )
    .addItem(
      "Append incremental preview → Master (candidates only)",
      "appendIncrementalPreviewToMaster",
    )
    .addSeparator()
    .addSeparator()
    .addItem(
      "Debug: replay last print",
      "previewIncrementalStep4ReplayLastPrint",
    )
    .addItem(
      "Debug: replay last calendar day",
      "previewIncrementalStep4ReplayLastCalendarDay",
    )
    .addItem("Inspect seed blocks from Master", "inspectSeedBlocksFromMaster")
    .addItem("Backup Master (Snapshot)", "backupMasterSheet");

  // --- Full rebuild raw ingest (catastrophe / first load) ---
  const tosLegacyImportMenu = ui
    .createMenu("Legacy (one section at a time)")
    .addItem(
      "Import TosTrades Current Account → TOS Trades - Combined",
      "tosTradesImportFromFolderCurrentAccount",
    )
    .addItem(
      "Import TosTop Current Account → TOS Top - Combined",
      "tosTopImportFromFolderCurrentAccount",
    )
    .addSeparator()
    .addItem(
      "Import TosTrades BOTH Accounts → TOS Trades - Combined",
      "tosTradesImportFromFolderBothAccounts",
    )
    .addItem(
      "Import TosTop BOTH Accounts → TOS Top - Combined",
      "tosTopImportFromFolderBothAccounts",
    );

  const tosFullImportMenu = ui
    .createMenu("Raw Data for Sorting (FULL Combined)")
    .addItem("Set folder (LT): TosTop + TosTrades", "tosSetCsvFolderIdLT")
    .addItem("Set folder (DT): TosTop + TosTrades", "tosSetCsvFolderIdDT")
    .addSeparator()
    .addItem(
      "Import BOTH sections Current Account → Combined (FULL)",
      "tosImportBothSectionsFromFolderCurrentAccount",
    )
    .addItem(
      "Import BOTH sections BOTH Accounts → Combined (FULL)",
      "tosImportBothSectionsFromFolderBothAccounts",
    )
    .addSeparator()
    .addItem("Blank TOS Trades - Combined", "tosBlankTOSTradesCombined")
    .addItem("Blank TOS Trades", "tosBlankTosTrades")
    .addItem("Blank TOS Top - Combined", "tosBlankTOSTopCombined")
    .addItem("Blank TOS Top", "tosBlankTosTop")
    .addItem("Blank all TOS sheets", "tosBlankALLTOSSheets")
    .addSeparator()
    .addItem(
      "Push: TOS Trades - Combined → TosTrades",
      "pushTosTradesCombinedToTosTrades",
    )
    .addItem("Push: TOS Top - Combined → TosTop", "pushTosTopCombinedToTosTop")
    .addItem("Push BOTH (Top + Trades)", "pushTosCombinedToBoth")
    .addSeparator()
    .addItem(
      "Run FULL (BOTH Accounts): CSV → Combined → TosTop/TosTrades → Schwab Import",
      "tosRunFullTosToSchwabImportBothAccounts",
    )
    .addSeparator()
    .addSubMenu(tosLegacyImportMenu);

  const settingsMenu = ui
    .createMenu("Settings")
    .addItem("Set Account Mode DT / LT", "promptSetAccountMode")
    .addItem(
      "Toggle TOS Import DEBUG Alerts",
      "promptToggleTosImportDebugAlerts",
    )
    .addSeparator()
    .addItem("Set Import Issues Write Mode", "promptSetImportIssuesWriteMode")
    .addItem("Set Mapping Issues Write Mode", "promptSetMappingIssuesWriteMode")
    .addItem("Set Staging Issues Write Mode", "promptSetStagingIssuesWriteMode")
    .addItem("Show Current Settings", "showCurrentSettingsDialog");

  const schwabFullMenu = ui
    .createMenu("Schwab Actions (FULL)")
    .addItem("Build Unified Import V3", "buildUnifiedImportV3")
    .addItem("Run Schwab Mapping Script", "mapSchwabImportByHeadersV3")
    .addItem(
      "2b. Audit Schwab Mapping (pre-Phase-3 gate)",
      "auditSchwabMappingV3",
    )
    .addItem("Blank Schwab Import", "clearSchwabImportExceptHeader")
    .addItem("Blank Schwab Mapping", "clearSchwabMappingExceptHeader")
    .addItem("Blank all Schwab Sheets", "blankAllSchwabSheets")
    .addSeparator()
    .addSubMenu(settingsMenu);

  const phase3FullMenu = ui
    .createMenu("Phase 3 (FULL walk)")
    .addItem(
      "Schwab Mapping to Import and run thru Block Logic",
      "refreshAllScripts",
    )
    .addItem("1. Move Schwab Mapping to Import", "copyMappingToImportByHeaders")
    .addItem("2. Validate and Clean Import", "validateAndCleanImportToHelperV3")
    .addItem("3. Run Block Logic on Helper", "populateStagingWithBlockLogicV3")
    .addItem(
      "Run Audit on Staging - Final Check before push to Master",
      "auditPipelineIntegrity",
    )
    .addSeparator()
    .addItem(
      "Push Staging → Master (Append ALL Staging — not incremental)",
      "appendStagingToMaster",
    )
    .addItem(
      "Replace Master from Staging (first load)",
      "replaceMasterFromStaging",
    )
    .addSeparator()
    .addItem("Blank Import", "blankImport")
    .addItem("Blank Helper", "blankHelper")
    .addItem("Blank Staging", "blankStaging")
    .addItem("Blank ALL Prep Sheets", "blankAllPrepSheets")
    .addSeparator()
    .addItem("Blank Master", "clearMasterExceptHeader");

  const fullRebuildMenu = ui
    .createMenu("Full rebuild (all data)")
    .addSubMenu(tosFullImportMenu)
    .addSubMenu(schwabFullMenu)
    .addSubMenu(phase3FullMenu);

  ui.createMenu("DB Tools")
    .addSubMenu(incrementalMenu)
    .addSubMenu(fullRebuildMenu)
    .addToUi();
}
