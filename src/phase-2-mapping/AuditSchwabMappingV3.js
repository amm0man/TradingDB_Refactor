/**
 * AuditSchwabMappingV3.js
 *
 * Phase 2 – Pre-Phase-3 gate
 *
 * Sole responsibility:
 *   Read the current "Schwab Mapping" sheet and log ERROR / WARN rows
 *   to "Schwab Mapping Issues" before Phase 3 runs.
 *
 * Called by:
 *   - Menu item "2b. Audit Schwab Mapping (pre-Phase-3 gate)"
 *
 * Related files:
 *   - mapSchwabImportByHeadersV3.js  (writes Schwab Mapping)
 *   - ImportIssues.js                (mappingIssues* + uiAlertSafe via Helpers.js)
 *
 * Do not change audit rules in this extract. Body is a straight move.
 */


// ============================================================================
// Pre-Phase-3 Schwab Mapping Audit
// Run this AFTER mapSchwabImportByHeadersV3 and BEFORE refreshAllScripts.
// Logs all findings to Schwab Mapping Issues (same sheet as Phase 2).
//
// Target before running Phase 3:
//   AuditErrors = 0    (hard block — do not proceed with any errors)
//   AuditWarns  = 0    (review each warn; accept or fix before proceeding)
// ============================================================================
function auditSchwabMappingV3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ctx = mappingIssuesStart("auditSchwabMappingV3");

  const sh = ss.getSheetByName(SHEET_SCHWAB_MAPPING);
  if (!sh) {
    mappingIssuesAdd(
      ctx,
      "ERROR",
      "",
      "Sheet",
      SHEET_SCHWAB_MAPPING,
      "Schwab Mapping sheet not found. Run mapSchwabImportByHeadersV3 first.",
    );
    mappingIssuesFlush(ctx);
    return;
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    mappingIssuesAdd(
      ctx,
      "WARN",
      "",
      "Sheet",
      SHEET_SCHWAB_MAPPING,
      "Schwab Mapping has no data rows. Run mapSchwabImportByHeadersV3 first.",
    );
    mappingIssuesFlush(ctx);
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const hm = buildHeaderIndexMap(headers);
  const data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();

  // Safe column reader — returns '' if header not found.
  function get(row, headerName) {
    const idx = hm[String(headerName).trim().toLowerCase()];
    return idx !== undefined ? row[idx] : "";
  }

  // ── Known phantom words: short tokens the tokenizer might extract that are
  //    definitely NOT real US equity tickers.
  //    Rule: only add a word here when you are 100% certain it cannot be a
  //    real ticker. When in doubt, leave it off and let the cross-reference
  //    check (Pass 2) surface it as a warning.
  const KNOWN_PHANTOM_TICKERS = new Set([
    "NO",
    "NA",
    "FREE",
    "NONE",
    "TBD",
    "NULL",
    "INC",
    "ETF",
    "LP",
    "LLC",
    "LTD",
    "CORP",
    "FUND",
    "TRUST",
    "BALANCE",
    "ADJUSTMENT",
    "MARGIN",
    "ORDINARY",
    "QUALIFIED",
    "DESCRIPTION",
  ]);

  // ── Corp action types that require an underlying ticker (non-cash types).
  //    If a row has one of these corp action tags AND a blank Ticker, it needs
  //    a Corp Action Map entry.
  const CORP_ACTIONS_NEED_TICKER = new Set([
    "Dividend",
    "Cash Dividend",
    "DRIP",
    "Non-Qualified Dividend",
    "Partnership Distribution",
    "Return of Capital",
    "Foreign Tax Withheld",
    "Reverse Split",
    "Mandatory Reverse Split",
    "Stock Split",
    "Forward Split",
    "Stock Merger",
    "Mandatory Merger",
    "Mandatory Exchange",
    "Reorganization",
    "Spin-off/Liquidation",
    "Transfer of Security",
    "Transfer of Security or Option In",
    "Transfer of Security or Option Out",
    "Pending Corp Action",
    "Corp Action Received Shares",
  ]);

  let totalRows = 0;
  let auditErrors = 0;
  let auditWarns = 0;

  // For Pass 2 cross-reference check.
  const tradeTickers = new Set(); // tickers seen on actual trade rows
  const corpOnlyTickers = new Map(); // ticker → first sheet row; corp action rows only

  // ──────────────────────────────────────────────────────────────────────────
  // PASS 1 — Row-level checks
  // ──────────────────────────────────────────────────────────────────────────
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const rowNum = r + 2; // actual sheet row (row 1 = header)
    totalRows++;

    const ticker = String(get(row, "Ticker") || "")
      .trim()
      .toUpperCase();
    const action = String(get(row, "Action") || "").trim();
    const corpAction = String(get(row, "Corporate Actions") || "").trim();
    const acctAction = String(get(row, "Account Actions") || "").trim();
    const qty = get(row, "Quantity");
    const entryPrice = get(row, "Entry Price");
    const signedQty = get(row, "Signed Quantity");
    const optStrike = get(row, "Option Strike");
    const optExp = get(row, "Option Expiration");
    const callPut = String(get(row, "Call/Put") || "")
      .trim()
      .toUpperCase();
    const ts = get(row, "Trade Time Stamp");
    const account = String(get(row, "Account") || "")
      .trim()
      .toUpperCase();
    const rawDesc = String(get(row, "Description") || "").trim();

    const isTrade = action.includes(" to "); // "Buy to Open", "Sell to Close", etc.
    const hasStrike =
      optStrike !== "" && optStrike !== null && optStrike !== undefined;
    const hasExp = optExp instanceof Date && !isNaN(optExp.getTime());
    const hasCp = callPut === "C" || callPut === "P";
    const isOption = isTrade && (hasStrike || hasExp || hasCp);

    // Track tickers for cross-reference check in Pass 2.
    if (isTrade && ticker) tradeTickers.add(ticker);
    if (
      !isTrade &&
      corpAction &&
      CORP_ACTIONS_NEED_TICKER.has(corpAction) &&
      ticker
    ) {
      if (!corpOnlyTickers.has(ticker)) corpOnlyTickers.set(ticker, rowNum);
    }

    // ── CHECK 1: Trade Time Stamp ──────────────────────────────────────────
    if (!(ts instanceof Date) || isNaN(ts.getTime())) {
      mappingIssuesAdd(
        ctx,
        "ERROR",
        rowNum,
        "Trade Time Stamp",
        String(ts || ""),
        "Missing or invalid Trade Time Stamp. This row will not sort or group correctly in Phase 3.",
      );
      auditErrors++;
    }

    // ── CHECK 2: Account must be DT or LT ────────────────────────────────
    if (account !== "DT" && account !== "LT") {
      mappingIssuesAdd(
        ctx,
        "ERROR",
        rowNum,
        "Account",
        account,
        'Account must be DT or LT. Found: "' + account + '".',
      );
      auditErrors++;
    }

    // ── CHECK 3: Known phantom ticker ────────────────────────────────────
    if (ticker && KNOWN_PHANTOM_TICKERS.has(ticker)) {
      mappingIssuesAdd(
        ctx,
        "WARN",
        rowNum,
        "Ticker",
        ticker,
        '"' +
          ticker +
          '" is a known phantom word extracted by the tokenizer, not a real ticker. ' +
          "Fix: add a Corp Action Map row for this Description and re-run mapSchwabImportByHeadersV3. " +
          'Description: "' +
          rawDesc.substring(0, 60) +
          '"',
      );
      auditWarns++;
    }

    // ── CHECK 4: Ticker longer than 6 chars (always a parsing artifact) ──
    if (ticker && ticker.length > 6 && !/^[.$]/.test(ticker)) {
      // Allow dot-prefixed OCC symbols (.SPY231231P400) — those are handled by Phase 3.
      mappingIssuesAdd(
        ctx,
        "WARN",
        rowNum,
        "Ticker",
        ticker,
        'Ticker "' +
          ticker +
          '" is ' +
          ticker.length +
          " chars — longer than any valid US equity ticker (max 6). " +
          "Likely a company name word extracted by the tokenizer. " +
          "Fix via Corp Action Map then re-run mapSchwabImportByHeadersV3.",
      );
      auditWarns++;
    }

    // ── CHECK 5: Trade row — Ticker required ─────────────────────────────
    if (isTrade && !ticker) {
      mappingIssuesAdd(
        ctx,
        "ERROR",
        rowNum,
        "Ticker",
        "",
        'Trade row Action = "' +
          action +
          '" has no Ticker. Cannot build blocks in Phase 3.',
      );
      auditErrors++;
    }

    // ── CHECK 6: Corp action row — Ticker required for non-cash types ────
    if (
      !isTrade &&
      corpAction &&
      CORP_ACTIONS_NEED_TICKER.has(corpAction) &&
      !ticker
    ) {
      mappingIssuesAdd(
        ctx,
        "WARN",
        rowNum,
        "Ticker",
        "",
        'Corp action "' +
          corpAction +
          '" requires a Ticker but none was resolved. ' +
          "Add to Corp Action Map and re-run mapSchwabImportByHeadersV3. " +
          'Description: "' +
          rawDesc.substring(0, 60) +
          '"',
      );
      auditWarns++;
    }

    // ── CHECK 7: Trade row — Quantity must be a positive number ──────────
    if (isTrade) {
      const qtyNum =
        typeof qty === "number"
          ? qty
          : Number(String(qty || "").replace(/,/g, ""));
      if (isNaN(qtyNum) || qtyNum <= 0) {
        mappingIssuesAdd(
          ctx,
          "ERROR",
          rowNum,
          "Quantity",
          String(qty || ""),
          'Trade row has zero or invalid Quantity. Action = "' + action + '".',
        );
        auditErrors++;
      }
    }

    // ── CHECK 8: Trade row — Entry Price must be positive ────────────────
    if (isTrade) {
      const ep =
        typeof entryPrice === "number"
          ? entryPrice
          : Number(String(entryPrice || "").replace(/,/g, ""));
      if (isNaN(ep) || ep <= 0) {
        mappingIssuesAdd(
          ctx,
          "ERROR",
          rowNum,
          "Entry Price",
          String(entryPrice || ""),
          'Trade row has zero or invalid Entry Price. Action = "' +
            action +
            '".',
        );
        auditErrors++;
      }
    }

    // ── CHECK 9: Signed Quantity sign must match Action direction ─────────
    if (isTrade) {
      const sq =
        typeof signedQty === "number"
          ? signedQty
          : Number(String(signedQty || ""));
      if (!isNaN(sq) && sq !== 0) {
        const isBuyAction = action.toUpperCase().startsWith("BUY");
        const isSellAction = action.toUpperCase().startsWith("SELL");
        if (isBuyAction && sq < 0) {
          mappingIssuesAdd(
            ctx,
            "ERROR",
            rowNum,
            "Signed Quantity",
            String(sq),
            'Buy action "' +
              action +
              '" has negative Signed Quantity (' +
              sq +
              "). Expected positive.",
          );
          auditErrors++;
        }
        if (isSellAction && sq > 0) {
          mappingIssuesAdd(
            ctx,
            "ERROR",
            rowNum,
            "Signed Quantity",
            String(sq),
            'Sell action "' +
              action +
              '" has positive Signed Quantity (' +
              sq +
              "). Expected negative.",
          );
          auditErrors++;
        }
      }
    }

    // ── CHECK 10: Option rows — Strike, Exp, Call/Put all present ─────────
    if (isOption) {
      if (!hasStrike) {
        mappingIssuesAdd(
          ctx,
          "ERROR",
          rowNum,
          "Option Strike",
          String(optStrike || ""),
          'Option trade missing Strike. Ticker = "' +
            ticker +
            '" Action = "' +
            action +
            '".',
        );
        auditErrors++;
      }
      if (!hasExp) {
        mappingIssuesAdd(
          ctx,
          "ERROR",
          rowNum,
          "Option Expiration",
          String(optExp || ""),
          "Option trade missing or invalid Expiration.",
        );
        auditErrors++;
      }
      if (!hasCp) {
        mappingIssuesAdd(
          ctx,
          "ERROR",
          rowNum,
          "Call/Put",
          callPut,
          'Option trade missing Call/Put. Expected C or P, got "' +
            callPut +
            '".',
        );
        auditErrors++;
      }
    }

    // ── CHECK 11: Partial option fields — Strike without Exp or vice versa ──
    if (isTrade && !isOption) {
      if (hasStrike && !hasExp) {
        mappingIssuesAdd(
          ctx,
          "WARN",
          rowNum,
          "Option Expiration",
          "",
          "Row has Strike (" +
            optStrike +
            ") but no Expiration. Inconsistent option fields.",
        );
        auditWarns++;
      }
      if (!hasStrike && hasExp) {
        mappingIssuesAdd(
          ctx,
          "WARN",
          rowNum,
          "Option Strike",
          "",
          "Row has Expiration but no Strike. Inconsistent option fields.",
        );
        auditWarns++;
      }
    }

    // ── CHECK 12: Ghost row — no classification at all ───────────────────
    if (!action && !corpAction && !acctAction) {
      mappingIssuesAdd(
        ctx,
        "WARN",
        rowNum,
        "Action",
        rawDesc.substring(0, 60),
        "Row has no Action, no Corporate Actions tag, and no Account Actions tag. " +
          "Completely unclassified. This may be an unrecognised action type from a TDA-era format. " +
          "Find this row in Schwab Import and identify the Action type.",
      );
      auditWarns++;
    }
  } // end Pass 1

  // ──────────────────────────────────────────────────────────────────────────
  // PASS 2 — Cross-reference: corp action tickers never seen in any trade row
  //
  // WHY THIS MATTERS ("what don't I know?"):
  // A ticker that appears ONLY in corp action rows (dividends, splits) and
  // NEVER in a trade row could be:
  //   A) Legitimate — a long-held position never traded in this dataset
  //   B) A phantom — a company-name word incorrectly extracted as a ticker
  //
  // You cannot tell the difference from the data alone. This check surfaces
  // every such ticker so you can verify each one manually, ONCE.
  // After verification, add it to the VERIFIED_CORP_ONLY_TICKERS set below
  // to silence the warn on future runs.
  // ──────────────────────────────────────────────────────────────────────────

  // Add verified long-held positions here after you've manually confirmed them.
  // Format: the actual ticker string, exact case as it appears in Schwab Mapping.
  const VERIFIED_CORP_ONLY_TICKERS = new Set([
    // 'JEPI',  // example: confirmed ETF held long-term, dividends only in this dataset
    // 'URNM',  // example: confirmed after manual check
    'PALAF' // added 9/11/26 PALAF had a reverse split so has an entry in Corp Actions. Resolves thru Phase 3 block logic correctly
  ]);

  let crossRefWarns = 0;
  for (const [tkr, firstRowNum] of corpOnlyTickers.entries()) {
    if (!tradeTickers.has(tkr) && !VERIFIED_CORP_ONLY_TICKERS.has(tkr)) {
      mappingIssuesAdd(
        ctx,
        "WARN",
        firstRowNum,
        "Ticker",
        tkr,
        '"' +
          tkr +
          '" appears in Corporate Actions rows but has NO matching trade row in this dataset. ' +
          'Either: (A) a legitimate long-held position never traded here — add "' +
          tkr +
          '" to ' +
          "VERIFIED_CORP_ONLY_TICKERS in auditSchwabMappingV3 to silence this warn, OR " +
          "(B) a phantom ticker extracted from a company name — fix via Corp Action Map.",
      );
      crossRefWarns++;
      auditWarns++;
    }
  }

  // ── Metrics ────────────────────────────────────────────────────────────────
  mappingIssuesSetMetric(ctx, "AuditRowsChecked", totalRows);
  mappingIssuesSetMetric(ctx, "AuditErrors", auditErrors);
  mappingIssuesSetMetric(ctx, "AuditWarns", auditWarns);
  mappingIssuesSetMetric(ctx, "AuditCrossRefWarns", crossRefWarns);
  mappingIssuesSetMetric(ctx, "UniqueTradeTickers", tradeTickers.size);
  mappingIssuesSetMetric(ctx, "UniqueCorpOnlyTickers", corpOnlyTickers.size);
  mappingIssuesFlush(ctx);

  const icon = auditErrors > 0 ? "🔴" : auditWarns > 0 ? "🟡" : "✅";
  uiAlertSafe(
    icon +
      " Schwab Mapping Audit Complete\n\n" +
      "Rows checked:        " +
      totalRows +
      "\n" +
      "Errors (block Phase 3): " +
      auditErrors +
      "\n" +
      "Warnings (review):   " +
      auditWarns +
      (crossRefWarns > 0
        ? "  ← includes " + crossRefWarns + " cross-reference warns"
        : "") +
      "\n\n" +
      (auditErrors > 0
        ? "🔴 Fix ALL errors before running Phase 3. Check Schwab Mapping Issues."
        : auditWarns > 0
          ? "🟡 Review warnings in Schwab Mapping Issues. Resolve or verify before Phase 3."
          : "✅ Schwab Mapping is clean. Safe to run refreshAllScripts."),
  );
}