/**
 * MapWriteSort.js
 *
 * Phase 2 – Group I: write, sort, header check, metric formatter
 *
 * Functions (same names as before — do not rename):
 *   - writeMappingRowsV3
 *   - formatCountsForMetric
 *   - assertNoDuplicateHeaders
 *   - sortMappingRowsByTradeTimeStamp
 *
 * Called by:
 *   mapSchwabImportByHeadersV3()
 *
 * Uses (already global — do not redeclare):
 *   MAP_V3_REBUILD_MAPPING_SHEET
 *   col                 (Helpers.js)
 *   mappingIssuesAdd / mappingIssuesFlush   (ImportIssues.js)
 *
 * Do not redeclare SHEET_* consts or MAP_V3_* toggles here.
 * Nested toNumOrNull stays inside the sort function (do not merge with toNum).
 */

/** Writes mapping rows to "Schwab Mapping". */
function writeMappingRowsV3(mappingSheet, mappingHeaders, outRows) {
  const tWriteFn = pipelineTimingNow();

  // Ensure enough rows
  const neededRows = outRows.length + 1; // + header
  const maxRows = mappingSheet.getMaxRows();
  if (neededRows > maxRows) {
    mappingSheet.insertRowsAfter(maxRows, neededRows - maxRows);
  }

  if (MAP_V3_REBUILD_MAPPING_SHEET) {
    // Clear everything below header (safe if you sort / rebuild)
    const lastCol = mappingSheet.getLastColumn();
    const lastRow = Math.max(mappingSheet.getLastRow(), 2);
    mappingSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }

  if (outRows.length) {
    const tSet = pipelineTimingNow();
    mappingSheet
      .getRange(2, 1, outRows.length, mappingHeaders.length)
      .setValues(outRows);
    pipelineTimingLog(
      "writeMappingRowsV3 setValues",
      tSet,
      "rows=" + outRows.length,
    );

    const tFmt = pipelineTimingNow();
    const tsCol = mappingHeaders.indexOf("Trade Time Stamp") + 1;
    const dateCol = mappingHeaders.indexOf("Trade Date") + 1;
    const n = outRows.length;
    if (tsCol > 0) {
      mappingSheet
        .getRange(2, tsCol, n, 1)
        .setNumberFormat("mm/dd/yyyy hh:mm:ss");
    }
    if (dateCol > 0) {
      mappingSheet.getRange(2, dateCol, n, 1).setNumberFormat("mm/dd/yyyy");
    }
    pipelineTimingLog("writeMappingRowsV3 numberFormat", tFmt, "rows=" + n);
  }

  pipelineTimingLog(
    "writeMappingRowsV3",
    tWriteFn,
    "rows=" + outRows.length,
  );
}

/** Make a compact "A=3 | B=10 | C=1" metric string (sorted by count desc). */
function formatCountsForMetric(countsObj) {
  const keys = Object.keys(countsObj || {});
  if (!keys.length) return "";

  keys.sort(function (a, b) {
    return (countsObj[b] || 0) - (countsObj[a] || 0);
  });

  // Keep it readable; you can increase this if you want the full breakdown.
  const maxItems = 20;
  const parts = [];
  for (let i = 0; i < Math.min(maxItems, keys.length); i++) {
    const k = keys[i];
    parts.push(k + "=" + countsObj[k]);
  }
  return parts.join(" | ");
}

/** Checks for duplicate Headers */
function assertNoDuplicateHeaders(headers, where, ctx) {
  // ctx is optional — pass it from mapSchwabImportByHeadersV3 so a flush
  // happens before the throw. Called from assertNoDuplicateHeaders(headers, sheet, ctx).
  const seen = {};
  const dups = [];
  for (let i = 0; i < headers.length; i++) {
    const key = String(headers[i]).trim().toLowerCase();
    if (!key) continue;
    if (seen[key]) dups.push(String(headers[i]));
    seen[key] = true;
  }
  if (dups.length) {
    const msg = `Duplicate headers in ${where}: ${dups.join(", ")}. Run aborted — fix the header row and re-run.`;
    if (ctx) {
      mappingIssuesAdd(ctx, "ERROR", "", "Headers", dups.join(", "), msg);
      mappingIssuesFlush(ctx);
    }
    throw new Error(msg);
  }
}

/** Sort rows by "Trade Time Stamp" ascending.
 *  Enhanced tie-breakers so spread legs stack nicely within the same minute bucket.
 */
function sortMappingRowsByTradeTimeStamp(items, mappingHeaderMap) {
  const tsIdx = col(mappingHeaderMap, "Trade Time Stamp");

  // Additional indexes for tie-break sorting (only used when timestamps tie)
  const acctIdx = col(mappingHeaderMap, "Account");
  const tickerIdx = col(mappingHeaderMap, "Ticker");
  const expIdx = col(mappingHeaderMap, "Option Expiration");
  const strikeIdx = col(mappingHeaderMap, "Option Strike");
  const actionIdx = col(mappingHeaderMap, "Action");

  function toValidDate(v) {
    return v instanceof Date && !isNaN(v) ? v : null;
  }

  function toMs(v) {
    const d = toValidDate(v);
    return d ? d.getTime() : null;
  }

  function normStr(v) {
    return String(v || "")
      .trim()
      .toUpperCase();
  }

  function toNumOrNull(v) {
    const n = typeof v === "number" ? v : Number(String(v || "").trim());
    return isNaN(n) ? null : n;
  }

  function actionRank(act) {
    // This is only a tie-breaker helper.
    // Goal: keep paired legs readable (Buy before Sell typically reads cleaner).
    const a = String(act || "");
    if (a.startsWith("Buy")) return 0;
    if (a.startsWith("Sell")) return 1;
    return 2;
  }

  items.sort(function (a, b) {
    const ta = a.row[tsIdx];
    const tb = b.row[tsIdx];

    const da = toValidDate(ta);
    const db = toValidDate(tb);

    // 1) Primary: timestamp (rows with timestamps come first)
    if (da && db) {
      if (da < db) return -1;
      if (da > db) return 1;

      // ===== Same exact timestamp: enhanced tie-breakers =====

      // 2) Account (keeps DT rows together, LT rows together)
      const acctA = normStr(a.row[acctIdx]);
      const acctB = normStr(b.row[acctIdx]);
      if (acctA !== acctB) return acctA < acctB ? -1 : 1;

      // 3) Ticker
      const tkrA = normStr(a.row[tickerIdx]);
      const tkrB = normStr(b.row[tickerIdx]);
      if (tkrA !== tkrB) return tkrA < tkrB ? -1 : 1;

      // 4) Expiration (THIS is the key that will make your SPX example stack as pairs)
      const expA = toMs(a.row[expIdx]);
      const expB = toMs(b.row[expIdx]);
      // Put dated expirations before blank expirations
      if (expA !== null && expB === null) return -1;
      if (expA === null && expB !== null) return 1;
      if (expA !== null && expB !== null && expA !== expB) return expA - expB;

      // 5) Strike (so within an expiration group, 4340 then 4345, etc.)
      const strikeA = toNumOrNull(a.row[strikeIdx]);
      const strikeB = toNumOrNull(b.row[strikeIdx]);
      if (strikeA !== null && strikeB === null) return -1;
      if (strikeA === null && strikeB !== null) return 1;
      if (strikeA !== null && strikeB !== null && strikeA !== strikeB)
        return strikeA - strikeB;

      // 6) Buy/Sell readability
      const actRankA = actionRank(a.row[actionIdx]);
      const actRankB = actionRank(b.row[actionIdx]);
      if (actRankA !== actRankB) return actRankA - actRankB;

      // 7) Stable final tie-breaker: original Schwab Import row order
      return (a.importRowNum || 0) - (b.importRowNum || 0);
    }

    if (da && !db) return -1;
    if (!da && db) return 1;

    // If neither has a timestamp, stabilize by import row order
    return (a.importRowNum || 0) - (b.importRowNum || 0);
  });
}
