/**
 * Phase3Step2_ValidateClean.js
 *
 * Phase 3 – Step 2: Validate & Clean
 *
 * Sole responsibility:
 *   Take the Import sheet (created by Step 1) and produce a clean, fully
 *   timestamped Helper sheet that the Block Logic can safely consume.
 *
 * What this file does:
 *   - Derives reliable full timestamps (parseTradeTimeStamp)
 *   - Canonicalises corporate actions
 *   - Upper-cases key text fields
 *   - Normalises tickers / OCC symbols
 *   - Validates quantities and prices
 *   - Writes the cleaned result to both Import and Helper
 *   - Logs validation issues
 *
 * Called by:
 *   - refreshAllScripts() in Phase3Processing.js
 *   - Directly from the menu if desired
 *
 * Related files:
 *   - Phase3Step1_CopyMapping.js   (will be created next)
 *   - Phase3BlockLogic.js          (consumes the Helper sheet)
 */
// ─────────────────────────────────────────────────────────────────────────────
// HELPER: Robust timestamp parser with fallback
// Placed here at the top so it's easy to find and share with other functions.
//
// PRIORITY 2 FIX (parseTradeTimeStamp): The old split(':') call silently broke
// when Trade Time was stored as "HHmm" (e.g. "0932") instead of "HH:mm"
// ("09:32").  mapSchwabImportByHeadersV3 formats Trade Time as "HHmm" with no
// colon, so any row that needed the fallback path got hours=932 and minutes=NaN,
// producing a wildly wrong or null timestamp.
// The new parser handles both "HH:mm" and "HHmm", plus an optional seconds
// component ("HH:mm:ss" or "HHmmss") so future Schwab format changes won't
// silently break anything.
// ─────────────────────────────────────────────────────────────────────────────
// =========================================================================
// HELPER: parseTradeTimeStamp
//   Builds a reliable Date object from Trade Time Stamp (preferred) or from
//   the separate Trade Date + Trade Time columns as fallback.
// =========================================================================
function parseTradeTimeStamp(tsVal, tradeDateVal, tradeTimeVal, ss) {
  let fullTimestamp = null;

  // Try Trade Time Stamp first (primary path — should always be present)
  if (tsVal) {
    fullTimestamp = tsVal instanceof Date ? tsVal : new Date(tsVal);
    if (isNaN(fullTimestamp.getTime())) {
      fullTimestamp = new Date(tsVal.toString().replace(/-/g, "/"));
    }
  }

  // FALLBACK: build from Trade Date + Trade Time when Trade Time Stamp is
  // missing or unparseable (fixes DT rows that arrive without a full timestamp)
  if (!fullTimestamp || isNaN(fullTimestamp.getTime())) {
    if (tradeDateVal && tradeTimeVal) {
      let baseDate =
        tradeDateVal instanceof Date ? tradeDateVal : new Date(tradeDateVal);
      if (!isNaN(baseDate.getTime())) {
        const timeStr = tradeTimeVal.toString().trim();
        let hours, minutes, seconds;

        if (timeStr.includes(":")) {
          // "16:5", "16:05", "9:5:00" — pad each piece so minutes 0-9 stay :05
          const parts = timeStr.split(":");
          hours = Number(parts[0]);
          minutes = Number(String(parts[1] || "0").padStart(2, "0"));
          seconds = Number(String(parts[2] || "0").padStart(2, "0"));
        } else {
          const digits = timeStr.replace(/\D/g, "");
          if (digits.length === 3) {
            // "165" came from unpadded 16:5 → 16:05, not 01:65
            const asHmm = Number(digits.substring(0, 2));
            const asMin1 = Number(digits.substring(2).padStart(2, "0"));
            const asHm = Number(digits.substring(0, 1));
            const asMin2 = Number(digits.substring(1, 3));
            if (asHmm >= 0 && asHmm <= 23 && asMin1 >= 0 && asMin1 <= 9) {
              hours = asHmm;
              minutes = asMin1;
              seconds = 0;
            } else if (asHm >= 0 && asHm <= 9 && asMin2 >= 0 && asMin2 <= 59) {
              hours = asHm;
              minutes = asMin2;
              seconds = 0;
            } else {
              hours = NaN;
              minutes = NaN;
              seconds = 0;
            }
          } else if (digits.length >= 4) {
            hours = Number(digits.substring(0, 2));
            minutes = Number(digits.substring(2, 4));
            seconds = digits.length >= 6 ? Number(digits.substring(4, 6)) : 0;
          } else {
            hours = NaN;
            minutes = NaN;
            seconds = 0;
          }
        }
      }
    }
  }

  return fullTimestamp && !isNaN(fullTimestamp.getTime())
    ? fullTimestamp
    : null;
}

/**
 * validateAndCleanImportToHelperV3
 * Runs right after copyMappingToImportByHeaders().
 *
 * WHY: Makes data 100% pristine AND derives clean Trade Date / Trade Time from
 * the authoritative Trade Time Stamp in BOTH Import and Helper.
 *
 * Changes in this version (5/3/2026):
 *   PRIORITY 1 — try/catch/finally wrapper so Staging Issues is ALWAYS flushed
 *                even on a mid-function crash.
 *   PRIORITY 2 — All importHeaders.indexOf() calls hoisted above the row loop.
 *                The old code re-ran ~12 header searches on every single row
 *                (45,000+ redundant ops on a full DT import).
 *   PRIORITY 2 — parseTradeTimeStamp now handles "HHmm" (no colon) as well as
 *                "HH:mm", matching the format written by mapSchwabImportByHeadersV3.
 */
// =========================================================================
// STEP 2: validateAndCleanImportToHelperV3
//   Runs immediately after copyMappingToImportByHeaders.
//   Cleans data, forces full timestamps, and writes the pristine result
//   into the "Helper" sheet that the block logic will read.
// =========================================================================
function validateAndCleanImportToHelperV3() {
  const tVal = pipelineTimingNow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName("Import");
  const helperSheet = ss.getSheetByName("Helper");
  if (!importSheet || !helperSheet)
    throw new Error("Import or Helper sheet not found!");

  // ── Staging Issues CTX ─────────────────────────────────────────────────────
  const ctx = stagingIssuesStart("validateAndCleanImportToHelperV3");
  importIssuesSetMetric(ctx, "SourceSheet", "Import");
  importIssuesSetMetric(ctx, "DestSheet", "Helper");
  // ───────────────────────────────────────────────────────────────────────────

  // ── PRIORITY 1: try/catch/finally so Issues log ALWAYS gets flushed ────────
  // WHY: Without this wrapper, any mid-function throw (sheet missing, bad data,
  // regex error, Sheets API limit) exits silently and leaves ZERO log output in
  // Staging Issues for this step — the worst time to have no diagnostics.
  // The finally block guarantees stagingIssuesFlush(ctx) runs even on crash.
  try {
    // 1. Read Import data (starts at row 4)
    const importData = importSheet
      .getRange(4, 1, importSheet.getLastRow() - 3, importSheet.getLastColumn())
      .getValues();

    importIssuesSetMetric(ctx, "SourceRowsReadExclHeader", importData.length);
    let tickerNormCount = 0;

    // ── Raw importHeaders array (lowercase-trimmed for indexOf matching) ──────
    const importHeaders = importSheet
      .getRange(1, 1, 1, importSheet.getLastColumn())
      .getValues()[0]
      .map((h) => h.trim().toLowerCase());

    // ── PRIORITY 2: Hoist ALL header index lookups ABOVE the row loop ─────────
    // WHY: importHeaders never changes between rows. Calling indexOf() inside the
    // loop ran every search on every row (12 searches × 15,000 DT rows = 180,000
    // redundant string comparisons per pipeline run). Hoisted once here = zero
    // redundancy. Also eliminates any risk of a typo on one pass vs another.
    const tsIdx = importHeaders.indexOf("trade time stamp");
    const dateIdx = importHeaders.indexOf("trade date");
    const timeIdx = importHeaders.indexOf("trade time");
    const tickerIdx = importHeaders.indexOf("ticker");
    const strategyTypeIdx = importHeaders.indexOf("strategy type");
    const tradeTypeIdx = importHeaders.indexOf("trade type");
    const signedQuantityIdx = importHeaders.indexOf("signed quantity");
    const quantityIdx = importHeaders.indexOf("quantity");
    const actionIdx = importHeaders.indexOf("action");
    const entryPriceIdx = importHeaders.indexOf("entry price");
    const optionContractIdx = importHeaders.indexOf("option contract");
    const strikeOCIdx = importHeaders.indexOf("option strike");
    const expOCIdx = importHeaders.indexOf("option expiration");
    const cpOCIdx = importHeaders.indexOf("call/put");
    const accountIdx = importHeaders.indexOf("account");
    const corpActionsIdx = importHeaders.indexOf("corporate actions");
    const accountActionsIdx = importHeaders.indexOf("account actions");
    // ─────────────────────────────────────────────────────────────────────────

    // Also need the uppercase-target column indices (used in the forEach below)
    // Hoisted here so the forEach doesn't re-search importHeaders each iteration.
    const uppercaseTextColumns = [
      "Account",
      "Ticker",
      "Action",
      "Trade Type",
      "Call/Put",
      "Strategy Type",
    ];
    const uppercaseColIndices = uppercaseTextColumns.map((colName) =>
      importHeaders.indexOf(colName.toLowerCase()),
    );

    // 2. Make sure Trade Time Stamp column exists on Helper
    const helperHeaders = helperSheet
      .getRange(1, 1, 1, helperSheet.getLastColumn())
      .getValues()[0];
    let tsCol =
      helperHeaders.findIndex(
        (h) => h.toString().trim().toLowerCase() === "trade time stamp",
      ) + 1;
    if (tsCol === 0) {
      helperSheet
        .getRange(1, helperSheet.getLastColumn() + 1)
        .setValue("Trade Time Stamp");
      tsCol = helperSheet.getLastColumn();
    }

    let outputRows = [];
    let errors = [];
    const errorSheet = ensureValidationErrorSheet();
    const nonTradeActions = ["EFN", "RAD", "JRN", "DOI", "EXP", "CRC", "CDB"];

    // === CLEAR OLD ERRORS ===
    if (errorSheet.getLastRow() > 1) {
      errorSheet
        .getRange(2, 1, errorSheet.getLastRow() - 1, errorSheet.getLastColumn())
        .clearContent();
    }

    // ── TIMEZONE: resolved once here so Utilities.formatDate doesn't need to
    // call ss.getSpreadsheetTimeZone() on every row (minor but free win).
    const tz = ss.getSpreadsheetTimeZone();

    // ─────────────────────────────────────────────────────────────────────────
    // MAIN ROW LOOP
    // All header index lookups use the hoisted constants above — never indexOf
    // inside the loop body.
    // ─────────────────────────────────────────────────────────────────────────
    for (let r = 0; r < importData.length; r++) {
      const row = importData[r];
      let cleanRow = [...row];
      let errorMsg = "";

      // ── 1. TIMESTAMP: derive from authoritative Trade Time Stamp (with fallback) ──
      // Uses the hoisted tsIdx / dateIdx / timeIdx (no indexOf in the loop).
      let fullTimestamp = null;
      if (tsIdx > -1) {
        const tsVal = row[tsIdx];
        const dateVal = dateIdx > -1 ? row[dateIdx] : null;
        const timeVal = timeIdx > -1 ? row[timeIdx] : null;

        fullTimestamp = parseTradeTimeStamp(tsVal, dateVal, timeVal, ss);

        if (fullTimestamp) {
          cleanRow[tsIdx] = fullTimestamp;
          const derivedDate = new Date(
            fullTimestamp.getFullYear(),
            fullTimestamp.getMonth(),
            fullTimestamp.getDate(),
          );
          const derivedTime = Utilities.formatDate(fullTimestamp, tz, "HH:mm");
          if (dateIdx > -1) cleanRow[dateIdx] = derivedDate;
          if (timeIdx > -1) cleanRow[timeIdx] = derivedTime;
        } else {
          errorMsg = "Invalid Trade Time Stamp";
        }
      }

      // ── 2. TRADE TYPE: derive here so Helper is the canonical pristine source ──
      // WHY: A blank Strategy Type must produce a blank Trade Type, not 'Option'.
      // Schwab settlement rows (stock exercise/assignment legs, EXP cash rows) arrive
      // with blank Strategy Type — defaulting them to 'Option' caused those rows to
      // fail the option field validation check (no Strike/Exp/CP on a stock row).
      // Rule: blank Ticker OR blank Strategy Type → blank Trade Type.
      // Only rows with an explicit Strategy Type keyword get a derived Trade Type.
      if (tickerIdx > -1 && strategyTypeIdx > -1 && tradeTypeIdx > -1) {
        const ticker = (row[tickerIdx] || "").toString().trim().toUpperCase();
        const strategyType = (row[strategyTypeIdx] || "")
          .toString()
          .trim()
          .toUpperCase();
        if (ticker && strategyType) {
          let tradeType = "Option";

          // A row with no Strike AND no Expiration is a stock row, full stop.
          // Catches settlement stock-delivery legs whose Strategy Type was forward-filled
          // from the parent spread (e.g. SHORT IC) which would otherwise derive 'Option'.
          const hasNoOptionFields =
            !cleanRow[strikeOCIdx] && !cleanRow[expOCIdx];
          if (hasNoOptionFields || strategyType.includes("STOCK"))
            tradeType = "Stock";
          else if (
            strategyType.includes("PCS") ||
            strategyType.includes("PDS") ||
            strategyType.includes("CCS") ||
            strategyType.includes("CDS") ||
            strategyType.includes("BUTTERFLY") ||
            strategyType.includes("IRON CONDOR")
          )
            tradeType = "Spread";

          cleanRow[tradeTypeIdx] = tradeType;
        } else {
          cleanRow[tradeTypeIdx] = "";
        }
      }

      // ── 2B. CORPORATE ACTION CANONICALIZATION ──────────────────────────
      // WHY: SYMBOL CHANGE and SPLIT rows may legally survive validation even
      // with sparse upstream fields, but Helper should still classify them as
      // stock-lineage rows so Staging does not have to guess.
      const corpActionVal = (
        corpActionsIdx > -1 ? cleanRow[corpActionsIdx] : ""
      )
        .toString()
        .trim()
        .toUpperCase();
      const cleanActionVal = (actionIdx > -1 ? cleanRow[actionIdx] : "")
        .toString()
        .trim()
        .toUpperCase();
      const cleanTickerVal = (tickerIdx > -1 ? cleanRow[tickerIdx] : "")
        .toString()
        .trim()
        .toUpperCase();

      if (
        cleanTickerVal &&
        (corpActionVal === "SYMBOL CHANGE" ||
          corpActionVal === "SPLIT" ||
          cleanActionVal === "SYMBOL CHANGE" ||
          cleanActionVal === "SPLIT")
      ) {
        if (tradeTypeIdx > -1) cleanRow[tradeTypeIdx] = "Stock";
        if (
          strategyTypeIdx > -1 &&
          !String(cleanRow[strategyTypeIdx] || "").trim()
        ) {
          cleanRow[strategyTypeIdx] = "LONG STOCK";
        }
      }

      // ── 3. UPPERCASE text columns ─────────────────────────────────────────
      // Uses pre-computed uppercaseColIndices array — no inner indexOf call.
      for (let u = 0; u < uppercaseColIndices.length; u++) {
        const idx = uppercaseColIndices[u];
        if (idx > -1) {
          cleanRow[idx] = (cleanRow[idx] || "").toString().trim().toUpperCase();
        }
      }

      // ── 4. TICKER NORMALIZATION: strip Schwab alternate index prefixes ─────
      // "$SPX.X" → "SPX" | "$NDX.X" → "NDX" | "SPY" → "SPY" (no change).
      // Uses hoisted tickerIdx.
      if (tickerIdx > -1) {
        const tkrRaw = (cleanRow[tickerIdx] || "")
          .toString()
          .trim()
          .toUpperCase();
        const tkrNormalized = tkrRaw
          .replace(/^\$/, "")
          .replace(/\.[A-Z]+$/, "");
        if (tkrNormalized !== tkrRaw) {
          tickerNormCount++;
          cleanRow[tickerIdx] = tkrNormalized;
          // Uncomment the line below only for debugging a specific normalization issue:
          // importIssuesAdd(ctx, 'INFO', r + 4, 'Ticker', `${tkrRaw} → ${tkrNormalized}`, 'Ticker normalized from Schwab alternate index symbol');
        }
      }

      // ── 5. SIGNED QUANTITY ────────────────────────────────────────────────
      // Uses hoisted signedQuantityIdx / quantityIdx / actionIdx.
      if (signedQuantityIdx > -1 && quantityIdx > -1 && actionIdx > -1) {
        const qtyVal = Number(row[quantityIdx]);
        const actionVal = (row[actionIdx] || "")
          .toString()
          .toUpperCase()
          .trim();
        cleanRow[signedQuantityIdx] = !isNaN(qtyVal)
          ? actionVal.includes("SELL")
            ? -1 * qtyVal
            : qtyVal
          : "";
      }

      // ── 6. QUANTITY validation ────────────────────────────────────────────
      // Uses hoisted quantityIdx.
      if (quantityIdx > -1) {
        const valNum = Number(row[quantityIdx]);
        cleanRow[quantityIdx] = !isNaN(valNum)
          ? valNum
          : row[quantityIdx]
            ? "❌ INVALID QTY"
            : "";
      }

      // ── 7. ENTRY PRICE validation ─────────────────────────────────────────
      // Uses hoisted entryPriceIdx.
      if (entryPriceIdx > -1) {
        const valNum = Number(row[entryPriceIdx]);
        cleanRow[entryPriceIdx] = !isNaN(valNum)
          ? valNum
          : row[entryPriceIdx]
            ? "❌ INVALID PRICE"
            : "";
      }

      // ── 8. OPTION CONTRACT: standardize to OCC format ────────────────────
      // Uses hoisted optionContractIdx / tickerIdx (reused) / strikeOCIdx /
      // expOCIdx / cpOCIdx.
      if (
        optionContractIdx > -1 &&
        tickerIdx > -1 &&
        strikeOCIdx > -1 &&
        expOCIdx > -1 &&
        cpOCIdx > -1
      ) {
        const ocTicker = (row[tickerIdx] || "").toString().toUpperCase().trim();
        const strike = Number(row[strikeOCIdx]);
        const exp = row[expOCIdx];
        let cp = (row[cpOCIdx] || "").toString().toUpperCase().trim();
        if (cp === "CALL") cp = "C";
        if (cp === "PUT") cp = "P";

        let expDatePart = "";
        if (exp) {
          if (exp instanceof Date) {
            const yy = String(exp.getFullYear()).slice(-2);
            const mm = String(exp.getMonth() + 1).padStart(2, "0");
            const dd = String(exp.getDate()).padStart(2, "0");
            expDatePart = yy + mm + dd;
          } else {
            const expStr = exp.toString().trim();
            const dateMatch = expStr.match(/(\d{2})[\/\-](\d{2})[\/\-](\d{4})/);
            if (dateMatch) {
              expDatePart =
                dateMatch[3].slice(-2) + dateMatch[1] + dateMatch[2];
            } else if (expStr.length === 8 && /^\d{8}$/.test(expStr)) {
              expDatePart = expStr.slice(2);
            }
          }
        }

        const strikePart = !isNaN(strike)
          ? String(Math.round(strike * 1000)).padStart(8, "0")
          : "";

        if (
          ocTicker &&
          expDatePart &&
          (cp === "C" || cp === "P") &&
          strikePart
        ) {
          cleanRow[optionContractIdx] =
            ocTicker.replace(/[^A-Z]/g, "") + expDatePart + cp + strikePart;
        }
      }

      // ── 9. VALIDATION RULES ───────────────────────────────────────────────
      // Uses hoisted actionIdx / accountIdx / tickerIdx / quantityIdx /
      // tradeTypeIdx / strikeOCIdx / expOCIdx / cpOCIdx /
      // corpActionsIdx / accountActionsIdx.
      const actionRaw = (row[actionIdx] || "").toString().trim().toUpperCase();
      const isTradeRow = [
        "BUY TO OPEN",
        "SELL TO OPEN",
        "BUY TO CLOSE",
        "SELL TO CLOSE",
      ].includes(actionRaw);

      const acct = (accountIdx > -1 ? row[accountIdx] : "")
        .toString()
        .trim()
        .toUpperCase();
      if (acct !== "DT" && acct !== "LT") errorMsg = "Account must be DT or LT";

      const tickerVal = (tickerIdx > -1 ? row[tickerIdx] : "")
        .toString()
        .trim()
        .toUpperCase();
      if (isTradeRow && !tickerVal)
        errorMsg = "Ticker required for trade actions";

      const qtyCheck = Number(quantityIdx > -1 ? row[quantityIdx] : "");
      if (isTradeRow && isNaN(qtyCheck)) errorMsg = "Quantity must be a number";

      // --- AFTER ---
      // WHY: Read tradeTypeCheck from cleanRow (not raw row) because Trade Type
      // is derived in Step 2 above and written to cleanRow — the raw row[tradeTypeIdx]
      // may be blank (first run) or a prior-run value (subsequent runs), both of
      // which cause incorrect validation behaviour. Also read strike/exp/cp from
      // cleanRow so Step 8 normalizations are visible here.
      // WHY: Wrap in !nonTradeActions guard so RAD, JRN, EFN etc. never reach the
      // option field check — they legitimately have no Strike/Exp/C-P requirement.
      const tradeTypeCheck = (tradeTypeIdx > -1 ? cleanRow[tradeTypeIdx] : "")
        .toString()
        .trim()
        .toUpperCase();
      if (
        !nonTradeActions.includes(actionRaw) &&
        (tradeTypeCheck.includes("OPTION") || tradeTypeCheck.includes("SPREAD"))
      ) {
        if (
          !(strikeOCIdx > -1 && cleanRow[strikeOCIdx]) ||
          !(expOCIdx > -1 && cleanRow[expOCIdx]) ||
          !(cpOCIdx > -1 && cleanRow[cpOCIdx])
        ) {
          errorMsg = "Options need Strike, Expiration, Call/Put";
        }
      }

      // Corporate / ledger rows clear any validation error — they don't need
      // Ticker, Quantity, or option fields.
      const corpAction = corpActionsIdx > -1 ? row[corpActionsIdx] : "";
      const accountAction =
        accountActionsIdx > -1 ? row[accountActionsIdx] : "";
      if (corpAction || accountAction || nonTradeActions.includes(actionRaw)) {
        errorMsg = "";
      }

      if (errorMsg) {
        errors.push([
          r + 4,
          "Action/Ticker",
          errorMsg,
          "Fix in Import sheet and re-run",
        ]);
        importIssuesAdd(
          ctx,
          "ERROR",
          r + 4,
          "Action/Ticker",
          errorMsg,
          "Row excluded from Helper — fix in Import and re-run",
        );
      } else {
        outputRows.push(cleanRow);
      }
    } // end main row loop

    // ── SAFE CLEAR both sheets before writing ─────────────────────────────────
    // WHY: If the previous run wrote more rows than this run, stale rows at the
    // bottom would silently flow into Staging on the next block-logic pass.
    const helperLastRow = helperSheet.getLastRow();
    if (helperLastRow >= 4) {
      helperSheet
        .getRange(4, 1, helperLastRow - 3, helperSheet.getLastColumn())
        .clearContent();
    }
    const importLastRow = importSheet.getLastRow();
    if (importLastRow >= 4) {
      importSheet
        .getRange(4, 1, importLastRow - 3, importSheet.getLastColumn())
        .clearContent();
    }

    // Write clean rows to both Helper and Import
    if (outputRows.length > 0) {
      helperSheet
        .getRange(4, 1, outputRows.length, importData[0].length)
        .setValues(outputRows);
      importSheet
        .getRange(4, 1, outputRows.length, importData[0].length)
        .setValues(outputRows);
    }

    // === FORCE CORRECT DISPLAY FORMATS ON BOTH SHEETS ===
    if (outputRows.length > 0) {
      const tsColNum = tsIdx + 1;
      const timeColNum = timeIdx + 1;
      importSheet
        .getRange(4, tsColNum, outputRows.length, 1)
        .setNumberFormat("M/d/yyyy HH:mm");
      importSheet
        .getRange(4, timeColNum, outputRows.length, 1)
        .setNumberFormat("HH:mm");
      helperSheet
        .getRange(4, tsCol, outputRows.length, 1)
        .setNumberFormat("M/d/yyyy HH:mm");
      helperSheet
        .getRange(4, timeColNum, outputRows.length, 1)
        .setNumberFormat("HH:mm");
    }

    // Write validation errors to the Validation Errors sheet if any
    if (errors.length > 0) {
      errorSheet
        .getRange(errorSheet.getLastRow() + 1, 1, errors.length, 4)
        .setValues(errors);
    }

    // ── CTX: finalize metrics ───────────────────────────────────────────────
    importIssuesSetMetric(ctx, "RowsWrittenExclHeader", outputRows.length);
    importIssuesSetMetric(ctx, "ValidationErrors", errors.length);
    importIssuesSetMetric(ctx, "TickerNormalizations", tickerNormCount);
    importIssuesSetMetric(ctx, "Success", "1");
    // ────────────────────────────────────────────────────────────────────────

    // Generic post-run check for Google server-side date rendering gaps.
    checkMissingDateTimeAndAlert(
      helperSheet,
      4,
      "validateAndCleanImportToHelperV3",
    );

    if (errors.length > 0) {
      uiAlertSafe(
        "⚠️ Validation found " +
          errors.length +
          " errors — check Validation Errors sheet!",
      );
    } else {
      uiAlertSafe(
        "✅ All data pristine — Import and Helper now have FULL derivations and correct display!",
      );
    }
  } catch (e) {
    // ── PRIORITY 1: Crash handler — log error metrics then re-throw ──────────
    importIssuesSetMetric(ctx, "Success", "0");
    importIssuesSetMetric(ctx, "ErrorMessage", e.message);
    importIssuesSetMetric(ctx, "ErrorStack", (e.stack || "").substring(0, 500));
    throw e;
  } finally {
    // ── PRIORITY 1: ALWAYS flush — even on throw ──────────────────────────
    // The old placement of stagingIssuesFlush(ctx) was at the end of the normal
    // flow only. Moving it here guarantees a log entry exists for every run,
    // successful or not.
    pipelineTimingLog("validateAndCleanImportToHelperV3", tVal);
    stagingIssuesFlush(ctx);
  }
}
