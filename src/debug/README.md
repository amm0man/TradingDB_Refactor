# debug/

This folder contains standalone debugging and diagnostic utilities used during development of the Trading Database project.

## Purpose

These scripts are **not** part of the core import/mapping/processing logic. They exist to help troubleshoot, inspect data, or test specific behaviors in isolation.

## Long-Term Goal

As we continue refactoring:
- Debug utilities will be moved here incrementally from the main Phase 1, 2, and 3 files.
- This keeps the core scripts smaller and cleaner.
- When the project is eventually handed off for deep analysis and optimization, the entire `debug/` folder can be excluded or archived so the delivered codebase stays focused and lean.

## Guidelines

- Files in this folder should be clearly named to indicate what they debug (e.g., `debugSchwabHeaderMapping.js`).
- They may contain temporary or experimental code.
- They should **not** be required by any production functions in `phase-1-import/`, `phase-2-mapping/`, `phase-3-processing/`, or `shared/`.

## Status

This folder was created June 23, 2026 as part of the ongoing refactor to improve maintainability and prepare for a future clean handoff.