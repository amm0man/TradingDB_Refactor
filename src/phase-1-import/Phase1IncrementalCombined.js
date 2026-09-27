/**
 * Phase1IncrementalCombined.js
 *
 * Read-only Phase 4 Combined preview.
 * Parses the same LT/DT folders as Combined Both, builds the same
 * Combined out-grid, then counts rows whose dedupe key is not already
 * on TOS Trades - Combined / TOS Top - Combined.
 *
 * Does not write those Combined sheets.
 * Does not call tosImportBothSectionsFromFolderBothAccounts
 * (that path still replaceEntireSheet: true).
 */

function previewIncrementalCombinedMergeBothAccounts() {
  const t0 = pipelineTimingNow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const tradesCombined = ss.getSheetByName(tosConfig.tradesCombinedSheetName);
  const topCombined = ss.getSheetByName(tosConfig.topCombinedSheetName);
  const tradesLastBefore = tradesCombined ? tradesCombined.getLastRow() : "";
  const topLastBefore = topCombined ? topCombined.getLastRow() : "";

  const dt = tosImportBothSectionsFromFolder(
    tosGetAccountFolderIdPropKey("DT"),
    "DT",
    { skipWrite: true },
  );
  const lt = tosImportBothSectionsFromFolder(
    tosGetAccountFolderIdPropKey("LT"),
    "LT",
    { skipWrite: true },
  );

  const tradesHeader =
    (dt && dt.tradesHeader) || (lt && lt.tradesHeader) || null;
  const topHeader = (dt && dt.topHeader) || (lt && lt.topHeader) || null;
  const tradesAllRows = []
    .concat((dt && dt.tradesAllRows) || [])
    .concat((lt && lt.tradesAllRows) || []);
  const topRowsAll = []
    .concat((dt && dt.topRowsAll) || [])
    .concat((lt && lt.topRowsAll) || []);
  const fileCount =
    Number((dt && dt.fileCount) || 0) + Number((lt && lt.fileCount) || 0);

  const tradesCtx = importIssuesStart("previewIncrementalCombinedMerge:Trades");
  const topCtx = importIssuesStart("previewIncrementalCombinedMerge:Top");
  importIssuesSetMetric(tradesCtx, "Account", "DT+LT");
  importIssuesSetMetric(topCtx, "Account", "DT+LT");

  const parsedTrades = tosTradesWriteCombinedFromParsed(
    tradesCtx,
    "DT+LT",
    "",
    "DT+LT",
    fileCount,
    tradesAllRows,
    tradesHeader,
    { buildOnly: true },
  );
  const parsedTop = tosTopWriteCombinedFromParsed(
    topCtx,
    "DT+LT",
    "",
    "DT+LT",
    fileCount,
    topRowsAll,
    topHeader,
    { buildOnly: true },
  );

  const tradesCmp = tosDiffParsedCombinedAgainstSheet_(
    parsedTrades,
    tradesCombined,
    tradesHeader,
    "TRADES",
  );
  const topCmp = tosDiffParsedCombinedAgainstSheet_(
    parsedTop,
    topCombined,
    topHeader,
    "TOP",
  );

  tosWriteIncrementalCombinedPreview_(
    ss,
    "Incremental Combined Trades Preview",
    parsedTrades && parsedTrades[0],
    tradesCmp.newRows,
  );
  tosWriteIncrementalCombinedPreview_(
    ss,
    "Incremental Combined Top Preview",
    parsedTop && parsedTop[0],
    topCmp.newRows,
  );

  const fileNote = tosDescribeCombinedSourceFiles_(
    dt,
    lt,
    tradesCombined,
    topCombined,
  );

  pipelineTimingLog(
    "previewIncrementalCombinedMergeBothAccounts",
    t0,
    "tradesNew=" +
      tradesCmp.newCount +
      " topNew=" +
      topCmp.newCount +
      " files=" +
      fileCount,
  );

  const tradesSample = tosSampleTradesKeyDebug_(
    parsedTrades,
    tradesCombined,
    tradesHeader,
  );

  uiAlertSafe(
    "Incremental Combined preview OK (read-only).\n\n" +
      "Drive CSV files parsed: " +
      fileCount +
      "\n" +
      fileNote +
      "\n" +
      "TOS Trades - Combined last row before: " +
      tradesLastBefore +
      "\n" +
      "  parsed unique: " +
      tradesCmp.parsedCount +
      "\n" +
      "  already on Combined: " +
      tradesCmp.alreadyCount +
      "\n" +
      "  NEW: " +
      tradesCmp.newCount +
      "\n" +
      "TOS Top - Combined last row before: " +
      topLastBefore +
      "\n" +
      "  parsed unique: " +
      topCmp.parsedCount +
      "\n" +
      "  already on Combined: " +
      topCmp.alreadyCount +
      "\n" +
      "  NEW: " +
      topCmp.newCount +
      "\n" +
      tradesSample +
      "Combined last rows after must match before.\n" +
      "Preview sheets: Incremental Combined Trades Preview / Top Preview.\n\n" +
      "On this freeze, with no new CSV, expect NEW 0 / 0.\n" +
      "Did not write TOS Trades - Combined or TOS Top - Combined.",
  );
}

function tosCombinedCellToKeyString_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(
      v,
      SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
      "yyyy-MM-dd HH:mm:ss",
    );
  }
  return String(v == null ? "" : v).trim();
}

function tosCombinedExpToKeyString_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(
      v,
      SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
      "d-MMM-yy",
    );
  }
  return String(v == null ? "" : v).trim();
}

function tosCombinedCanonicalFromOutRow_(
  outHeader,
  row,
  canonicalHeader,
  kind,
) {
  const inner = [];
  for (let i = 0; i < canonicalHeader.length; i++) {
    const name = canonicalHeader[i];
    const idx = outHeader.indexOf(name);
    if (idx < 0) {
      inner.push("");
      continue;
    }
    if (name === "Exp") inner.push(tosCombinedExpToKeyString_(row[idx]));
    else inner.push(tosCombinedCellToKeyString_(row[idx]));
  }

  // Top Combined TIME is ET→CT corrected after the dedupe key is built.
  // TimeRaw stays the original HHmmss text, so Top keys must use TimeRaw.
  // Trades Combined Exec Time is already the converted text the minute-key
  // iso branch expects. Do not use Trades TimeRaw — Sheets often turns
  // "8/31/23 7:42:21" into a Date, and String(Date) will not match.
  if (kind === "TOP") {
    const idxTimeRaw = outHeader.indexOf("TimeRaw");
    const raw =
      idxTimeRaw >= 0 ? tosCombinedCellToKeyString_(row[idxTimeRaw]) : "";
    const idxTime = canonicalHeader.indexOf("TIME");
    if (raw && idxTime >= 0) inner[idxTime] = raw;
  }
  return inner;
}

function tosSampleTradesKeyDebug_(
  parsedNewOut,
  combinedSheet,
  canonicalHeader,
) {
  if (!parsedNewOut || parsedNewOut.length < 2 || !combinedSheet) return "";
  if (!canonicalHeader) return "";

  const pHeader = parsedNewOut[0];
  const pRow = parsedNewOut[1];
  const pAcct =
    pHeader.indexOf("Account") >= 0 ? pRow[pHeader.indexOf("Account")] : "";
  const pInner = tosCombinedCanonicalFromOutRow_(
    pHeader,
    pRow,
    canonicalHeader,
    "TRADES",
  );
  const pKey = tosTradesDedupeKey_(pAcct, canonicalHeader, pInner);
  const pTime = tosTradesExecTimeMinuteKey_(
    pInner[canonicalHeader.indexOf("Exec Time")],
  );

  const cLastCol = combinedSheet.getLastColumn();
  const cGrid = combinedSheet.getRange(1, 1, 2, cLastCol).getValues();
  const cHeader = cGrid[0];
  const cRow = cGrid[1];
  const cAcct =
    cHeader.indexOf("Account") >= 0 ? cRow[cHeader.indexOf("Account")] : "";
  const cInner = tosCombinedCanonicalFromOutRow_(
    cHeader,
    cRow,
    canonicalHeader,
    "TRADES",
  );
  const cKey = tosTradesDedupeKey_(cAcct, canonicalHeader, cInner);
  const cTime = tosTradesExecTimeMinuteKey_(
    cInner[canonicalHeader.indexOf("Exec Time")],
  );
  const cExecRaw =
    cHeader.indexOf("Exec Time") >= 0 ? cRow[cHeader.indexOf("Exec Time")] : "";
  const cTimeRaw =
    cHeader.indexOf("TimeRaw") >= 0 ? cRow[cHeader.indexOf("TimeRaw")] : "";
  const cExp = cHeader.indexOf("Exp") >= 0 ? cRow[cHeader.indexOf("Exp")] : "";
  const pExp = pHeader.indexOf("Exp") >= 0 ? pRow[pHeader.indexOf("Exp")] : "";

  return (
    "Trades key sample (row 2 vs parsed row 1):\n" +
    "  Combined Exec Time type/value: " +
    typeof cExecRaw +
    " / " +
    String(cExecRaw) +
    "\n" +
    "  Combined TimeRaw type/value: " +
    typeof cTimeRaw +
    " / " +
    String(cTimeRaw) +
    "\n" +
    "  Combined minute key: " +
    cTime +
    "\n" +
    "  Parsed minute key: " +
    pTime +
    "\n" +
    "  Combined account: " +
    String(cAcct) +
    "\n" +
    "  Parsed account: " +
    String(pAcct) +
    "\n" +
    "  key strings equal: " +
    "  Combined Exp type/value: " +
    typeof cExp +
    " / " +
    String(cExp) +
    "\n" +
    "  Parsed Exp type/value: " +
    typeof pExp +
    " / " +
    String(pExp) +
    "\n" +
    "  Combined Exp key: " +
    tosNormalizeExpKey_(cInner[canonicalHeader.indexOf("Exp")]) +
    "\n" +
    "  Parsed Exp key: " +
    tosNormalizeExpKey_(pInner[canonicalHeader.indexOf("Exp")]) +
    "\n" +
    (cKey === pKey) +
    "\n\n"
  );
}

function tosDiffParsedCombinedAgainstSheet_(
  parsedNewOut,
  combinedSheet,
  canonicalHeader,
  kind,
) {
  const empty = {
    parsedCount: 0,
    alreadyCount: 0,
    newCount: 0,
    newRows: [],
  };
  if (!parsedNewOut || parsedNewOut.length < 2 || !canonicalHeader) {
    return empty;
  }

  const parsedHeader = parsedNewOut[0];
  const existingKeys = {};
  if (combinedSheet && combinedSheet.getLastRow() >= 2) {
    const cLast = combinedSheet.getLastRow();
    const cLastCol = combinedSheet.getLastColumn();
    const cGrid = combinedSheet.getRange(1, 1, cLast, cLastCol).getValues();
    const cHeader = cGrid[0];
    for (let r = 1; r < cGrid.length; r++) {
      const account =
        cHeader.indexOf("Account") >= 0
          ? cGrid[r][cHeader.indexOf("Account")]
          : "";
      if (!String(account || "").trim()) continue;
      const inner = tosCombinedCanonicalFromOutRow_(
        cHeader,
        cGrid[r],
        canonicalHeader,
        kind,
      );
      const key =
        kind === "TRADES"
          ? tosTradesDedupeKey_(account, canonicalHeader, inner)
          : tosTopDedupeKey_(account, canonicalHeader, inner);
      existingKeys[key] = true;
    }
  }

  let alreadyCount = 0;
  const newRows = [];
  for (let r = 1; r < parsedNewOut.length; r++) {
    const row = parsedNewOut[r];
    const account =
      parsedHeader.indexOf("Account") >= 0
        ? row[parsedHeader.indexOf("Account")]
        : "";
    const inner = tosCombinedCanonicalFromOutRow_(
      parsedHeader,
      row,
      canonicalHeader,
      kind,
    );
    const key =
      kind === "TRADES"
        ? tosTradesDedupeKey_(account, canonicalHeader, inner)
        : tosTopDedupeKey_(account, canonicalHeader, inner);
    if (existingKeys[key]) alreadyCount++;
    else newRows.push(row);
  }

  return {
    parsedCount: parsedNewOut.length - 1,
    alreadyCount: alreadyCount,
    newCount: newRows.length,
    newRows: newRows,
  };
}

function tosWriteIncrementalCombinedPreview_(ss, sheetName, header, newRows) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  sh.clear();
  if (!header || !header.length) return;
  const out = [header].concat(newRows || []);
  sh.getRange(1, 1, out.length, header.length).setValues(out);
  sh.setFrozenRows(1);
}

function tosDescribeCombinedSourceFiles_(dt, lt, tradesCombined, topCombined) {
  const driveNames = {};
  function addParsed(pack) {
    if (!pack) return;
    const rows = []
      .concat(pack.tradesAllRows || [])
      .concat(pack.topRowsAll || []);
    for (let i = 0; i < rows.length; i++) {
      const n = String(rows[i].sourceFile || "").trim();
      if (n) driveNames[n] = true;
    }
  }
  addParsed(dt);
  addParsed(lt);

  const combinedNames = {};
  function addSheet(sh) {
    if (!sh || sh.getLastRow() < 2) return;
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const idx = headers.indexOf("SourceFile");
    if (idx < 0) return;
    const vals = sh.getRange(2, idx + 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      const n = String(vals[i][0] || "").trim();
      if (n) combinedNames[n] = true;
    }
  }
  addSheet(tradesCombined);
  addSheet(topCombined);

  const driveList = Object.keys(driveNames).sort();
  const onlyDrive = [];
  for (let i = 0; i < driveList.length; i++) {
    if (!combinedNames[driveList[i]]) onlyDrive.push(driveList[i]);
  }
  return (
    "SourceFile names in this parse: " +
    driveList.length +
    "\n" +
    "SourceFile names not yet on Combined: " +
    (onlyDrive.length ? onlyDrive.join(", ") : "none")
  );
}

/**
 * Incremental Combined write.
 *
 * Same parse + key compare as previewIncrementalCombinedMergeBothAccounts.
 * Appends only rows whose dedupe key is not already on Combined.
 * Does not call tosImportBothSectionsFromFolderBothAccounts
 * (that path still replaceEntireSheet: true).
 *
 * On this freeze NEW is 0 / 0, so this is a no-op.
 */
function mergeIncrementalCombinedNewRowsBothAccounts() {
  const t0 = pipelineTimingNow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const tradesCombined = ss.getSheetByName(tosConfig.tradesCombinedSheetName);
  const topCombined = ss.getSheetByName(tosConfig.topCombinedSheetName);
  const tradesLastBefore = tradesCombined ? tradesCombined.getLastRow() : "";
  const topLastBefore = topCombined ? topCombined.getLastRow() : "";

  const dt = tosImportBothSectionsFromFolder(
    tosGetAccountFolderIdPropKey("DT"),
    "DT",
    { skipWrite: true },
  );
  const lt = tosImportBothSectionsFromFolder(
    tosGetAccountFolderIdPropKey("LT"),
    "LT",
    { skipWrite: true },
  );

  const tradesHeader =
    (dt && dt.tradesHeader) || (lt && lt.tradesHeader) || null;
  const topHeader = (dt && dt.topHeader) || (lt && lt.topHeader) || null;
  const tradesAllRows = []
    .concat((dt && dt.tradesAllRows) || [])
    .concat((lt && lt.tradesAllRows) || []);
  const topRowsAll = []
    .concat((dt && dt.topRowsAll) || [])
    .concat((lt && lt.topRowsAll) || []);
  const fileCount =
    Number((dt && dt.fileCount) || 0) + Number((lt && lt.fileCount) || 0);

  const tradesCtx = importIssuesStart("mergeIncrementalCombined:Trades");
  const topCtx = importIssuesStart("mergeIncrementalCombined:Top");
  importIssuesSetMetric(tradesCtx, "Account", "DT+LT");
  importIssuesSetMetric(topCtx, "Account", "DT+LT");

  const parsedTrades = tosTradesWriteCombinedFromParsed(
    tradesCtx,
    "DT+LT",
    "",
    "DT+LT",
    fileCount,
    tradesAllRows,
    tradesHeader,
    { buildOnly: true },
  );
  const parsedTop = tosTopWriteCombinedFromParsed(
    topCtx,
    "DT+LT",
    "",
    "DT+LT",
    fileCount,
    topRowsAll,
    topHeader,
    { buildOnly: true },
  );

  const tradesCmp = tosDiffParsedCombinedAgainstSheet_(
    parsedTrades,
    tradesCombined,
    tradesHeader,
    "TRADES",
  );
  const topCmp = tosDiffParsedCombinedAgainstSheet_(
    parsedTop,
    topCombined,
    topHeader,
    "TOP",
  );

  if (tradesCmp.newCount === 0 && topCmp.newCount === 0) {
    pipelineTimingLog(
      "mergeIncrementalCombinedNewRowsBothAccounts",
      t0,
      "no-op tradesNew=0 topNew=0",
    );
    uiAlertSafe(
      "Incremental Combined merge skipped — no new rows.\n\n" +
        "Drive CSV files parsed: " +
        fileCount +
        "\n" +
        "TOS Trades - Combined last row before/after: " +
        tradesLastBefore +
        " / " +
        tradesLastBefore +
        "\n" +
        "TOS Top - Combined last row before/after: " +
        topLastBefore +
        " / " +
        topLastBefore +
        "\n" +
        "Trades NEW: 0\n" +
        "Top NEW: 0\n\n" +
        "Did not write Combined.\n" +
        "Did not call tosImportBothSectionsFromFolderBothAccounts.",
    );
    return;
  }

  const resp = ui.alert(
    "Merge new Combined rows",
    "Trades NEW: " +
      tradesCmp.newCount +
      "\n" +
      "Top NEW: " +
      topCmp.newCount +
      "\n\n" +
      "This appends those rows onto Combined.\n" +
      "It does not replace the Combined sheets.\n\n" +
      "Cancel if those counts look wrong.",
    ui.ButtonSet.OK_CANCEL,
  );
  if (resp !== ui.Button.OK) {
    pipelineTimingLog(
      "mergeIncrementalCombinedNewRowsBothAccounts",
      t0,
      "cancelled tradesNew=" +
        tradesCmp.newCount +
        " topNew=" +
        topCmp.newCount,
    );
    return;
  }

  tosAppendCombinedNewRows_(
    tradesCombined,
    parsedTrades && parsedTrades[0],
    tradesCmp.newRows,
    "TRADES",
  );
  tosAppendCombinedNewRows_(
    topCombined,
    parsedTop && parsedTop[0],
    topCmp.newRows,
    "TOP",
  );

  pipelineTimingLog(
    "mergeIncrementalCombinedNewRowsBothAccounts",
    t0,
    "appended trades=" +
      tradesCmp.newCount +
      " top=" +
      topCmp.newCount +
      " tradesRow=" +
      (tradesCombined ? tradesCombined.getLastRow() : "") +
      " topRow=" +
      (topCombined ? topCombined.getLastRow() : ""),
  );

  uiAlertSafe(
    "Incremental Combined merge OK.\n\n" +
      "Trades appended: " +
      tradesCmp.newCount +
      "\n" +
      "Top appended: " +
      topCmp.newCount +
      "\n" +
      "TOS Trades - Combined last row before/after: " +
      tradesLastBefore +
      " / " +
      (tradesCombined ? tradesCombined.getLastRow() : "") +
      "\n" +
      "TOS Top - Combined last row before/after: " +
      topLastBefore +
      " / " +
      (topCombined ? topCombined.getLastRow() : "") +
      "\n\n" +
      "Did not replace Combined.\n" +
      "Did not call tosImportBothSectionsFromFolderBothAccounts.",
  );
}

function tosAppendCombinedNewRows_(sheet, header, newRows, kind) {
  if (!sheet || !header || !newRows || !newRows.length) return;

  const rows = newRows.map(function (r) {
    return r.slice();
  });

  if (kind === "TOP") {
    const fiDateCol = header.indexOf("DATE");
    const fiTimeCol = header.indexOf("TIME");
    const fiTimeRawCol = header.indexOf("TimeRaw");
    if (fiDateCol >= 0 && fiTimeCol >= 0 && fiTimeRawCol >= 0) {
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const timeRawVal = row[fiTimeRawCol];
        const dateIso = tosTopNormalizeDateToIso(row[fiDateCol]);
        let timeRawStr;
        if (timeRawVal instanceof Date && !isNaN(timeRawVal.getTime())) {
          const h = String(timeRawVal.getHours()).padStart(2, "0");
          const m = String(timeRawVal.getMinutes()).padStart(2, "0");
          const s = String(timeRawVal.getSeconds()).padStart(2, "0");
          timeRawStr = h + m + s;
        } else {
          timeRawStr = normalizeTimeHHmmss(String(timeRawVal ?? "").trim());
        }
        const timeCorrected = tosEtToCtHHmmss(timeRawStr, dateIso ?? "");
        if (timeCorrected) row[fiTimeCol] = timeCorrected;
      }
    }
  }

  const writeCols = Math.min(sheet.getLastColumn(), header.length);
  const firstEmpty = sheet.getLastRow() + 1;
  const out = rows.map(function (r) {
    return r.slice(0, writeCols);
  });
  const headerRow = sheet.getRange(1, 1, 1, writeCols).getValues()[0];

  // rowsCount = number of new rows. startRow = first empty Combined row.
  // Same helper signature Combined Both uses:
  // tosFormatHeaderColumnAsText(sheet, headerRow, name, rowsCount, startRow)
  if (kind === "TRADES") {
    tosFormatHeaderColumnAsText(
      sheet,
      headerRow,
      "Exec Time",
      out.length,
      firstEmpty,
    );
    tosFormatHeaderColumnAsText(
      sheet,
      headerRow,
      "Symbol",
      out.length,
      firstEmpty,
    );
    sheet.getRange(firstEmpty, 1, out.length, writeCols).setValues(out);
  } else {
    tosFormatHeaderColumnsAsText(
      sheet,
      headerRow,
      ["TIME", "TimeRaw"],
      out.length,
      firstEmpty,
    );
    sheet.getRange(firstEmpty, 1, out.length, writeCols).setValues(out);
  }
}
