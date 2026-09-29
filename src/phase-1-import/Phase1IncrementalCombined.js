/**
 * Phase1IncrementalCombined.js
 *
 * Incremental Combined + Incremental Push.
 *
 * Preview / merge parse the same LT/DT folders as Combined Both,
 * skip yearlies already on BOTH Combined sheets, always re-parse
 * Incremental.csv, then append only rows whose overlap key is new.
 * Does not call tosImportBothSectionsFromFolderBothAccounts
 * (that path still replaceEntireSheet: true).
 *
 * Incremental Push copies Combined rows whose SourceFile ends with
 * Incremental.csv onto working TosTrades / TosTop (slice only).
 */

function previewIncrementalCombinedMergeBothAccounts() {
  const t0 = pipelineTimingNow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const tradesCombined = ss.getSheetByName(tosConfig.tradesCombinedSheetName);
  const topCombined = ss.getSheetByName(tosConfig.topCombinedSheetName);
  const tradesLastBefore = tradesCombined ? tradesCombined.getLastRow() : "";
  const topLastBefore = topCombined ? topCombined.getLastRow() : "";

  const tradesNames = tosCollectSheetSourceFileNames_(tradesCombined);
  const topNames = tosCollectSheetSourceFileNames_(topCombined);
  const skipSourceFileNames = tosYearlySourceFilesToSkip_(
    tradesCombined,
    topCombined,
    tradesNames,
    topNames,
  );
  const dt = tosImportBothSectionsFromFolder(
    tosGetAccountFolderIdPropKey("DT"),
    "DT",
    { skipWrite: true, skipSourceFileNames: skipSourceFileNames },
  );
  const lt = tosImportBothSectionsFromFolder(
    tosGetAccountFolderIdPropKey("LT"),
    "LT",
    { skipWrite: true, skipSourceFileNames: skipSourceFileNames },
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
  const fileCounts = tosSumImportFileCounts_(dt, lt);
  const fileCount = fileCounts.parsed;

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
    tradesNames,
    topNames,
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
      "Drive CSV files found: " +
      fileCounts.found +
      "\n" +
      "Drive CSV files parsed: " +
      fileCount +
      "\n" +
      "Yearlies skipped (already on Combined): " +
      fileCounts.skipped +
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
      "If those Incremental.csv files are already on Combined, expect NEW 0 / 0.\n" +
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
    "Trades key sample (Combined row 2 vs first parsed row):\n" +
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
    "  key strings equal: " +
    (cKey === pKey) +
    "\n" +
    "  (false is OK after file-skip: Combined row 2 is oldest history,\n" +
    "   parsed row 1 is the first Incremental.csv row.)\n\n"
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

function tosDescribeCombinedSourceFiles_(
  dt,
  lt,
  tradesCombined,
  topCombined,
  tradesNamesOpt,
  topNamesOpt,
) {
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

  const tradesNames =
    tradesNamesOpt || tosCollectSheetSourceFileNames_(tradesCombined);
  const topNames = topNamesOpt || tosCollectSheetSourceFileNames_(topCombined);
  const combinedNames = {};
  const tKeys = Object.keys(tradesNames);
  for (let i = 0; i < tKeys.length; i++) combinedNames[tKeys[i]] = true;
  const pKeys = Object.keys(topNames);
  for (let i = 0; i < pKeys.length; i++) combinedNames[pKeys[i]] = true;

  const driveList = Object.keys(driveNames).sort();
  const onlyDrive = [];
  for (let i = 0; i < driveList.length; i++) {
    if (!combinedNames[driveList[i]]) onlyDrive.push(driveList[i]);
  }
  return (
    "SourceFile names in this parse: " +
    driveList.length +
    (driveList.length ? "\n  " + driveList.join("\n  ") : "") +
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
 * When NEW is 0 this is a no-op (append is skipped).
 */
function mergeIncrementalCombinedNewRowsBothAccounts() {
  const t0 = pipelineTimingNow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const tradesCombined = ss.getSheetByName(tosConfig.tradesCombinedSheetName);
  const topCombined = ss.getSheetByName(tosConfig.topCombinedSheetName);
  const tradesLastBefore = tradesCombined ? tradesCombined.getLastRow() : "";
  const topLastBefore = topCombined ? topCombined.getLastRow() : "";

  const skipSourceFileNames = tosYearlySourceFilesToSkip_(
    tradesCombined,
    topCombined,
  );
  const dt = tosImportBothSectionsFromFolder(
    tosGetAccountFolderIdPropKey("DT"),
    "DT",
    { skipWrite: true, skipSourceFileNames: skipSourceFileNames },
  );
  const lt = tosImportBothSectionsFromFolder(
    tosGetAccountFolderIdPropKey("LT"),
    "LT",
    { skipWrite: true, skipSourceFileNames: skipSourceFileNames },
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
  const fileCounts = tosSumImportFileCounts_(dt, lt);
  const fileCount = fileCounts.parsed;

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
        "Drive CSV files found: " +
        fileCounts.found +
        "\n" +
        "Drive CSV files parsed: " +
        fileCount +
        "\n" +
        "Yearlies skipped (already on Combined): " +
        fileCounts.skipped +
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

function tosIsIncrementalSourceFile_(name) {
  return /incremental\.csv$/i.test(String(name || "").trim());
}

/**
 * Distinct SourceFile names already stored on a Combined sheet.
 */
function tosCollectSheetSourceFileNames_(sh) {
  const names = {};
  if (!sh || sh.getLastRow() < 2) return names;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = headers.indexOf("SourceFile");
  if (idx < 0) return names;
  const vals = sh.getRange(2, idx + 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    const n = String(vals[i][0] || "").trim();
    if (n) names[n] = true;
  }
  return names;
}

/**
 * Yearlies already on BOTH Combined sheets. Incremental.csv is never skipped.
 */
function tosYearlySourceFilesToSkip_(
  tradesCombined,
  topCombined,
  tradesNamesOpt,
  topNamesOpt,
) {
  const tradesNames =
    tradesNamesOpt || tosCollectSheetSourceFileNames_(tradesCombined);
  const topNames = topNamesOpt || tosCollectSheetSourceFileNames_(topCombined);
  const skip = {};
  const keys = Object.keys(tradesNames);
  for (let i = 0; i < keys.length; i++) {
    const n = keys[i];
    if (topNames[n] && !tosIsIncrementalSourceFile_(n)) skip[n] = true;
  }
  return skip;
}

function tosSumImportFileCounts_(dt, lt) {
  return {
    parsed:
      Number((dt && dt.fileCount) || 0) + Number((lt && lt.fileCount) || 0),
    found:
      Number((dt && dt.filesFound) || (dt && dt.fileCount) || 0) +
      Number((lt && lt.filesFound) || (lt && lt.fileCount) || 0),
    skipped:
      Number((dt && dt.filesSkipped) || 0) +
      Number((lt && lt.filesSkipped) || 0),
  };
}

function tosCombinedHeaderIndex_(headers, name) {
  const j = headers.indexOf(name);
  if (j === -1) throw new Error("Missing Combined header: " + name);
  return j;
}

/**
 * Combined rows whose SourceFile ends with Incremental.csv,
 * mapped to the working TosTrades Push columns.
 * Uses getDisplayValues so Exec Time / Exp stay text (chimera fix).
 */
function tosBuildIncrementalTradesPushRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(tosConfig.tradesCombinedSheetName);
  if (!src)
    throw new Error("Missing sheet: " + tosConfig.tradesCombinedSheetName);

  const displays = src.getDataRange().getDisplayValues();
  if (!displays || displays.length < 2) {
    return { out: null, sourceFileNames: [], firstDate: "", lastDate: "" };
  }

  const headers = displays[0].map(function (h) {
    return String(h == null ? "" : h).trim();
  });
  const idxAccount = tosCombinedHeaderIndex_(headers, "Account");
  const idxSource = tosCombinedHeaderIndex_(headers, "SourceFile");
  const wanted = [
    "Account",
    "Exec Time",
    "Spread",
    "Side",
    "Qty",
    "Pos Effect",
    "Symbol",
    "Exp",
    "Strike",
    "Type",
    "Price",
    "Net Price",
    "Order Type",
  ];
  const idx = { Account: idxAccount };
  for (let i = 0; i < wanted.length; i++) {
    if (wanted[i] !== "Account") {
      idx[wanted[i]] = tosCombinedHeaderIndex_(headers, wanted[i]);
    }
  }

  const out = [wanted];
  const sourceFileNames = {};
  let firstDate = "";
  let lastDate = "";

  for (let r = 1; r < displays.length; r++) {
    const row = displays[r];
    const sourceFile = String(
      row[idxSource] == null ? "" : row[idxSource],
    ).trim();
    if (!tosIsIncrementalSourceFile_(sourceFile)) continue;
    sourceFileNames[sourceFile] = true;

    const hasAny = wanted.some(function (h) {
      return String(row[idx[h]] == null ? "" : row[idx[h]]).trim() !== "";
    });
    if (!hasAny) continue;

    const execDisplay = String(row[idx["Exec Time"]] || "").trim();
    const expDisplay = String(row[idx["Exp"]] || "").trim();
    let execOut = "";
    const mIso = execDisplay.match(
      /^(\d{4})-(\d{2})-(\d{2})[\s\t]+(\d{2}):(\d{2})(?::(\d{2}))?/,
    );
    if (mIso) {
      execOut =
        mIso[1] +
        "-" +
        mIso[2] +
        "-" +
        mIso[3] +
        " " +
        mIso[4] +
        ":" +
        mIso[5] +
        ":" +
        (mIso[6] || "00");
    } else if (execDisplay) {
      execOut = execDisplay;
    }
    if (!execOut) continue;

    const outRow = wanted.map(function (h) {
      if (h === "Exec Time") return execOut;
      if (h === "Exp") return expDisplay;
      return row[idx[h]];
    });
    const symJ = wanted.indexOf("Symbol");
    if (symJ >= 0)
      outRow[symJ] = String(outRow[symJ] == null ? "" : outRow[symJ]).trim();
    out.push(outRow);

    const day = execOut.substring(0, 10);
    if (!firstDate || day < firstDate) firstDate = day;
    if (!lastDate || day > lastDate) lastDate = day;
  }

  return {
    out: out,
    sourceFileNames: Object.keys(sourceFileNames),
    firstDate: firstDate,
    lastDate: lastDate,
  };
}

/**
 * Combined rows whose SourceFile ends with Incremental.csv,
 * mapped to the working TosTop Push columns.
 */
function tosBuildIncrementalTopPushRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(tosConfig.topCombinedSheetName);
  if (!src) throw new Error("Missing sheet: " + tosConfig.topCombinedSheetName);

  const values = src.getDataRange().getValues();
  if (!values || values.length < 2) {
    return { out: null, sourceFileNames: [], firstDate: "", lastDate: "" };
  }

  const headers = values[0].map(function (h) {
    return String(h == null ? "" : h).trim();
  });
  function headerIndex(nameOrNames) {
    const names = Array.isArray(nameOrNames) ? nameOrNames : [nameOrNames];
    for (let i = 0; i < names.length; i++) {
      const j = headers.indexOf(names[i]);
      if (j !== -1) return j;
    }
    throw new Error("Missing Combined Top header. Tried: " + names.join(" | "));
  }

  const idx = {
    Account: headerIndex("Account"),
    SourceFile: headerIndex("SourceFile"),
    date: headerIndex("DATE"),
    time: headerIndex("TIME"),
    type: headerIndex("TYPE"),
    desc: headerIndex("DESCRIPTION"),
    miscFees: headerIndex("Misc Fees"),
    commFees: headerIndex([
      "Commissions & Fees",
      "Commissions Fees",
      "Commissions and Fees",
    ]),
    amount: headerIndex("AMOUNT"),
  };

  const outHeaders = [
    "Account",
    "DATE",
    "TIME",
    "TYPE",
    "DESCRIPTION",
    "Misc Fees",
    "Commissions Fees",
    "AMOUNT",
  ];
  const out = [outHeaders];
  const sourceFileNames = {};
  let firstDate = "";
  let lastDate = "";

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const sourceFile = String(
      row[idx.SourceFile] == null ? "" : row[idx.SourceFile],
    ).trim();
    if (!tosIsIncrementalSourceFile_(sourceFile)) continue;
    sourceFileNames[sourceFile] = true;

    const descUpper = String(row[idx.desc] == null ? "" : row[idx.desc])
      .trim()
      .toUpperCase();
    if (descUpper === "TOTAL") continue;

    const timeRaw = String(row[idx.time] == null ? "" : row[idx.time]).trim();
    const timeHHmmss = normalizeTimeHHmmss(timeRaw);
    const dateVal = row[idx.date];
    const dateKey =
      tosTopNormalizeDateToIso(dateVal) || String(dateVal || "").trim();

    const hasAny =
      String(dateVal == null ? "" : dateVal).trim() ||
      String(timeHHmmss || "").trim() ||
      String(row[idx.type] == null ? "" : row[idx.type]).trim() ||
      String(row[idx.desc] == null ? "" : row[idx.desc]).trim();
    if (!hasAny) continue;

    out.push([
      row[idx.Account],
      dateVal,
      timeHHmmss,
      row[idx.type],
      row[idx.desc],
      row[idx.miscFees],
      row[idx.commFees],
      row[idx.amount],
    ]);
    if (dateKey) {
      if (!firstDate || dateKey < firstDate) firstDate = dateKey;
      if (!lastDate || dateKey > lastDate) lastDate = dateKey;
    }
  }

  return {
    out: out,
    sourceFileNames: Object.keys(sourceFileNames),
    firstDate: firstDate,
    lastDate: lastDate,
  };
}

function tosWriteIncrementalPushPreview_(ss, sheetName, out) {
  const sh = tosGetOrCreateSheet(ss, sheetName);
  sh.clear();
  if (!out || !out.length) return sh;
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  return sh;
}

function tosReplaceWorkingSheetFromPushGrid_(
  dst,
  out,
  textHeaders,
  execTimeFormat,
) {
  if (!dst || !out || out.length < 1) return;
  const prevLastRow = Math.max(dst.getLastRow(), 1);
  const prevLastCol = Math.max(dst.getLastColumn(), 1);
  const outRows = out.length;
  const outCols = out[0].length;

  tosFormatHeaderColumnsAsText(dst, out[0], textHeaders, outRows, 1);
  dst.getRange(1, 1, outRows, outCols).setValues(out);

  if (execTimeFormat && outRows > 1) {
    dst.getRange(2, 2, outRows - 1, 1).setNumberFormat(execTimeFormat);
  }

  if (prevLastRow > outRows) {
    dst
      .getRange(outRows + 1, 1, prevLastRow - outRows, prevLastCol)
      .clearContent();
  }
  if (prevLastCol > outCols) {
    const rowsToClear = Math.max(prevLastRow, outRows);
    dst
      .getRange(1, outCols + 1, rowsToClear, prevLastCol - outCols)
      .clearContent();
  }

  if (execTimeFormat && outRows > 2) {
    dst
      .getRange(2, 1, outRows - 1, outCols)
      .sort({ column: 2, ascending: true });
  }
}

function previewIncrementalPushFromCombined() {
  const t0 = pipelineTimingNow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tradesBefore = ss.getSheetByName(tosConfig.tosTradesSheetName);
  const topBefore = ss.getSheetByName(tosConfig.tosTopSheetName);
  const tradesLast = tradesBefore ? tradesBefore.getLastRow() : "";
  const topLast = topBefore ? topBefore.getLastRow() : "";

  const tradesPack = tosBuildIncrementalTradesPushRows_();
  const topPack = tosBuildIncrementalTopPushRows_();
  const tradesRows = tradesPack.out ? tradesPack.out.length - 1 : 0;
  const topRows = topPack.out ? topPack.out.length - 1 : 0;

  tosWriteIncrementalPushPreview_(
    ss,
    "Incremental TosTrades Preview",
    tradesPack.out,
  );
  tosWriteIncrementalPushPreview_(
    ss,
    "Incremental TosTop Preview",
    topPack.out,
  );

  pipelineTimingLog(
    "previewIncrementalPushFromCombined",
    t0,
    "trades=" + tradesRows + " top=" + topRows,
  );

  uiAlertSafe(
    "Incremental Push preview OK (working sheets not written).\n\n" +
      "SourceFile filter: name ends with Incremental.csv\n" +
      "Trades files: " +
      (tradesPack.sourceFileNames.join(", ") || "none") +
      "\n" +
      "  rows: " +
      tradesRows +
      "\n" +
      "  dates: " +
      tradesPack.firstDate +
      " → " +
      tradesPack.lastDate +
      "\n" +
      "Top files: " +
      (topPack.sourceFileNames.join(", ") || "none") +
      "\n" +
      "  rows: " +
      topRows +
      "\n" +
      "  dates: " +
      topPack.firstDate +
      " → " +
      topPack.lastDate +
      "\n\n" +
      "Working TosTrades last row still: " +
      tradesLast +
      "\n" +
      "Working TosTop last row still: " +
      topLast +
      "\n" +
      "Combined last rows unchanged (14525 / 15026 tonight).\n" +
      "Preview sheets: Incremental TosTrades Preview / Incremental TosTop Preview.\n\n" +
      "Tonight expect Trades 97 and Top 114.",
  );
}

function pushIncrementalCombinedToWorkingSheets() {
  const t0 = pipelineTimingNow();
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dstTrades = ss.getSheetByName(tosConfig.tosTradesSheetName);
  const dstTop = ss.getSheetByName(tosConfig.tosTopSheetName);
  if (!dstTrades)
    throw new Error("Missing sheet: " + tosConfig.tosTradesSheetName);
  if (!dstTop) throw new Error("Missing sheet: " + tosConfig.tosTopSheetName);

  const tradesBefore = dstTrades.getLastRow();
  const topBefore = dstTop.getLastRow();
  const tradesPack = tosBuildIncrementalTradesPushRows_();
  const topPack = tosBuildIncrementalTopPushRows_();
  const tradesRows = tradesPack.out ? tradesPack.out.length - 1 : 0;
  const topRows = topPack.out ? topPack.out.length - 1 : 0;

  const resp = ui.alert(
    "Replace working TosTrades / TosTop with Incremental slice?",
    "Trades rows: " +
      tradesRows +
      "  (" +
      tradesPack.firstDate +
      " → " +
      tradesPack.lastDate +
      ")\n" +
      "Top rows: " +
      topRows +
      "  (" +
      topPack.firstDate +
      " → " +
      topPack.lastDate +
      ")\n\n" +
      "This REPLACES TosTrades and TosTop with those rows only.\n" +
      "Combined is not written.\n" +
      "Full history is still on Combined. Restore with Full rebuild → Push BOTH.\n\n" +
      "Cancel if counts are not 97 / 114 tonight.",
    ui.ButtonSet.OK_CANCEL,
  );
  if (resp !== ui.Button.OK) {
    pipelineTimingLog(
      "pushIncrementalCombinedToWorkingSheets",
      t0,
      "cancelled trades=" + tradesRows + " top=" + topRows,
    );
    return;
  }

  tosReplaceWorkingSheetFromPushGrid_(
    dstTrades,
    tradesPack.out,
    ["Symbol", "Exec Time", "Exp"],
    "yyyy-mm-dd hh:mm:ss",
  );
  tosReplaceWorkingSheetFromPushGrid_(dstTop, topPack.out, ["TIME"], "");

  pipelineTimingLog(
    "pushIncrementalCombinedToWorkingSheets",
    t0,
    "trades=" +
      tradesRows +
      " top=" +
      topRows +
      " tradesRow=" +
      dstTrades.getLastRow() +
      " topRow=" +
      dstTop.getLastRow(),
  );

  uiAlertSafe(
    "Incremental Push OK.\n\n" +
      "TosTrades last row before/after: " +
      tradesBefore +
      " / " +
      dstTrades.getLastRow() +
      "  (wrote " +
      tradesRows +
      ")\n" +
      "TosTop last row before/after: " +
      topBefore +
      " / " +
      dstTop.getLastRow() +
      "  (wrote " +
      topRows +
      ")\n\n" +
      "Combined last rows unchanged.\n" +
      "Did not call pushTosCombinedToBoth.",
  );
}

/**
 * Read-only. Lists each Drive CSV and why Incremental Combined
 * PARSEs or SKIPs it. Does not write Combined.
 */
function debugListIncrementalFileSkip() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tradesCombined = ss.getSheetByName(tosConfig.tradesCombinedSheetName);
  const topCombined = ss.getSheetByName(tosConfig.topCombinedSheetName);
  const tradesNames = tosCollectSheetSourceFileNames_(tradesCombined);
  const topNames = tosCollectSheetSourceFileNames_(topCombined);
  const skip = tosYearlySourceFilesToSkip_(
    tradesCombined,
    topCombined,
    tradesNames,
    topNames,
  );

  const lines = [];
  function walk(account) {
    const resolved = tosResolveAccountCsvFolder(
      tosGetAccountFolderIdPropKey(account),
      account,
      importIssuesStart("debugListIncrementalFileSkip:" + account),
    );
    if (!resolved) {
      lines.push(account + ": folder not resolved");
      return;
    }
    const files = tosListCsvFilesInFolder(
      resolved.folder,
      importIssuesStart("debugListIncrementalFileSkipList:" + account),
    );
    lines.push(
      "--- " + account + " Drive files: " + (files ? files.length : 0) + " ---",
    );
    if (!files) return;
    for (let i = 0; i < files.length; i++) {
      const n = String(files[i].getName() || "").trim();
      const onTrades = tradesNames[n] ? "Y" : "N";
      const onTop = topNames[n] ? "Y" : "N";
      let why;
      if (tosIsIncrementalSourceFile_(n)) why = "PARSE Incremental.csv";
      else if (skip[n]) why = "SKIP on both Combined";
      else if (onTrades === "N" && onTop === "N")
        why = "PARSE name not on Combined";
      else if (onTrades === "N") why = "PARSE missing from Trades Combined";
      else if (onTop === "N") why = "PARSE missing from Top Combined";
      else why = "PARSE unexpected";
      lines.push(n);
      lines.push("  trades=" + onTrades + " top=" + onTop + " " + why);
    }
  }
  walk("DT");
  walk("LT");
  uiAlertSafe(lines.join("\n"));
}
