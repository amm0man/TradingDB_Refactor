# TradingDB_Refactor

Google Apps Script project for a custom trading database that imports, maps, processes, and audits trade data from thinkorswim (TOS) and Schwab into a clean Master sheet.

## Project Goals

- Maintain a reliable, auditable trading database
- Support complex options spreads and multi-leg positions
- Keep data pristine before writing to the Master sheet (read-only source of truth)
- Gradually refactor for better maintainability, reduced duplication, and smaller production code size

## Architecture Overview (3-Phase System)

| Phase | Folder                  | Purpose |
|-------|-------------------------|---------|
| **Phase 1** | `src/phase-1-import/`      | Import raw data from TOS and Schwab |
| **Phase 2** | `src/phase-2-mapping/`     | Standardize and map imported data using headers |
| **Phase 3** | `src/phase-3-processing/`  | Core processing logic, block handling, and writing to Master |
| **Shared**  | `src/shared/`              | Common utilities and services used across phases |
| **Audit**   | `src/audit/`               | Post-processing validation and integrity checks |
| **Debug**   | `src/debug/`               | Standalone debugging utilities (excluded from core production code) |

## Current Folder Structure
src/
├── audit/
│   └── postStagingAudit.js
├── debug/
│   ├── README.md
│   ├── debugBuildUnifiedImportV3.js
│   ├── debugDBTools.js
│   └── debugSchwabHeaderMapping.js
├── phase-1-import/
│   ├── BuildUnifiedImportV3.js
│   ├── TosSchwabImportPipeline.js
│   └── TosSheetWriteHelpers.js
├── phase-2-mapping/
│   └── MapSchwabImportByHeadersV3.js
├── phase-3-processing/
│   └── DBTools.js
└── shared/
├── ImportIssues.js
├── SettingsService.js
└── SheetBlanking.js


## Key Notes

- `DBTools.js` contains the custom menu and the majority of Phase 3 logic.
- Debug utilities are intentionally separated in `src/debug/` so they can be easily excluded from the final production codebase.
- The `audit/` folder contains validation functions that run after processing but before writing to the Master sheet.
- This project uses `clasp` for local development with Google Apps Script.

## Current Status

This project is in active refactoring mode (started June 2026). The focus is on:
- Removing dead/commented-out code
- Consolidating duplicate logic
- Improving folder organization and naming consistency
- Reducing overall production code size in preparation for deeper analysis and optimization

## Workflow

- Work is done locally using VS Code + `clasp`
- Changes are pushed with `clasp push` and committed via Git
- New chats are started for larger files to keep context manageable

---

**Last Updated:** June 23, 2026