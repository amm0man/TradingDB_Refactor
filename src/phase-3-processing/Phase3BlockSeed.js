/**
 * Phase3BlockSeed.js
 *
 * Phase 4 seed helpers. Not wired into the daily full rebuild.
 *
 *   emptyBlockState()              — same object the Step 4 loop already creates
 *   blockKeyFromStagingLikeRow_()  — same key rules as populateStagingWithBlockLogicV3
 *   parseSymbolChangeFromMasterRow_ — Family S: FROM/TO on a Master SYMBOL CHANGE row
 *   seedBlocksFromMaster()         — rebuild blocks{} from Master last-row state
 *   inspectSeedBlocksFromMaster()  — menu: preview only (writes Block Seed Preview)
 *
 * populateStagingWithBlockLogicV3(seedBlocks) still starts from {} when
 * called with no argument. Passing a seed is for a later incremental function.
 *
 * Family S (2026-09-18):
 *   populateStagingWithBlockLogicV3 SYMBOL CHANGE moves blocks{} old ticker →
 *   new ticker and writes the rename row under the NEW ticker. History rows
 *   under ENCUF / ISENF / SV stay on the sheet with running qty and
 *   Block Close Flag 0. Last-row-per-key seed then marks those old names LIVE.
 *
 *   Do not invent a sale. Do not delete history. Do not set Block Close Flag 1
 *   on the last old-ticker row while running qty is still open — audit CHECK 2
 *   (close flag 1 but running qty ≠ 0) would fire.
 *
 *   Seed flattens:
 *     1) Notes FROM= / FROM_RESOLVED= when those tokens differ from dest
 *     2) Earlier stock key when the same Position ID / Trade Group ID later
 *        appears under a different ticker
 *     3) Classified aliases (ENCUF→EU, ISENF→ISOU, SV→SMR) when both keys exist
 */

/**
 * Inspect-only map of Family S retired keys. seedBlocksFromMaster fills this.
 * Must never be copied onto blocks{} — Object.keys(blocks) would treat it
 * as a live block key in a later incremental Step 4.
 */
var SEED_RETIRED_BY_RENAME_ = {};

/**
 * Classified rename-source → dest tickers for seed honesty.
 * Seed-only. Not used by the Phase 3 full rebuild.
 * Add a row when Inspect shows LIVE under a retired name and Master
 * confirms a rename (not a still-held separate lot).
 */
var FAMILY_S_TICKER_ALIASES_ = {
  ENCUF: "EU",
  ISENF: "ISOU",
  SV: "SMR",
};

/**
 * Factory for one blocks{} entry.
 * Shape matches the object created inside populateStagingWithBlockLogicV3
 * when if (!blocks[key]) fires.
 */
function emptyBlockState() {
  return {
    unit: 0,
    block: 1,
    runningQty: 0,
    pnl: 0,
    entryCost: 0,
    openTs: null,
    positionId: "",
    strategyType: "",
  };
}

/**
 * Parse the TG number from Trade Group ID.
 * Examples: DT-SPY-LP-230417-TG001  → 1
 *           SPREAD-SPY-CCS-...-TG002 → 2
 *           DT-UNG-LST-TG001-ST → 1
 */
function parseTgBlockNumber_(tradeGroupId) {
  const s = String(tradeGroupId || "");
  const m = s.match(/-TG(\d{3})(?:-ST)?$/i);
  return m ? Number(m[1]) : 0;
}

/**
 * Build the same block key Step 4 uses.
 * colMap values are 0-based indexes (unlike populateStaging colMap which is 1-based).
 */
function blockKeyFromStagingLikeRow_(row, col, tz) {
  const acct = String(row[col["account"]] || "")
    .trim()
    .toUpperCase();
  const ticker = String(row[col["ticker"]] || "")
    .trim()
    .toUpperCase();
  const spreadId = String(row[col["spread group id"]] || "").trim();
  let tradeType = String(row[col["trade type"]] || "")
    .trim()
    .toUpperCase();

  if (spreadId) return acct + "|" + spreadId;

  if (tradeType === "OPTION") {
    const exp = row[col["option expiration"]];
    const expStr =
      exp instanceof Date
        ? Utilities.formatDate(exp, tz, "yyyy-MM-dd")
        : String(exp || "");
    const strike = Number(row[col["option strike"]]) || 0;
    const cp = String(row[col["call/put"]] || "")
      .toUpperCase()
      .replace("CALL", "C")
      .replace("PUT", "P");
    return acct + "|" + ticker + "|" + expStr + "|" + strike + "|" + cp;
  }

  return acct + "|" + ticker;
}

/**
 * Stock book key only (Account|Ticker). Option and spread keys have more parts.
 */
function isStockAcctTickerKey_(key) {
  const parts = String(key || "").split("|");
  return parts.length === 2 && !!parts[0] && !!parts[1];
}

/**
 * Family S helper. Master col map is 0-based.
 *
 * Expected Notes (Phase 2):
 *   FROM=ENCUF | TO=EU | FROM_RESOLVED=ENCUF | TO_RESOLVED=EU | RAW=...
 *
 * Returns null when the row is not a rename or Notes have no FROM/TO.
 */
function parseSymbolChangeFromMasterRow_(row, col) {
  const action =
    col["action"] !== undefined
      ? String(row[col["action"]] || "")
          .trim()
          .toUpperCase()
      : "";
  const corp =
    col["corporate actions"] !== undefined
      ? String(row[col["corporate actions"]] || "")
          .trim()
          .toUpperCase()
      : "";

  const isRename =
    action === "SYMBOL CHANGE" || corp.indexOf("SYMBOL CHANGE") !== -1;
  if (!isRename) return null;

  const notes =
    col["notes"] !== undefined ? String(row[col["notes"]] || "") : "";
  if (!notes) return null;

  function pull(label) {
    const m = notes.match(
      new RegExp("(?:^|\\|\\s*)" + label + "=([^|]+)", "i"),
    );
    return m ? String(m[1]).trim().toUpperCase() : "";
  }

  let fromRaw = pull("FROM");
  let toRaw = pull("TO");
  const fromResolved = pull("FROM_RESOLVED");
  const toResolved = pull("TO_RESOLVED");
  const raw = pull("RAW");

  const phraseSrc = raw || notes;
  const phrase = String(phraseSrc).match(
    /SYMBOL CHANGE FROM\s+([A-Z0-9.\/]+)\s+TO\s+([A-Z0-9.\/]+)/i,
  );
  if (phrase) {
    if (!fromRaw) fromRaw = String(phrase[1]).trim().toUpperCase();
    if (!toRaw) toRaw = String(phrase[2]).trim().toUpperCase();
  }

  if (!fromRaw && !fromResolved && !toRaw && !toResolved) return null;

  return {
    fromRaw: fromRaw,
    fromResolved: fromResolved,
    toRaw: toRaw,
    toResolved: toResolved,
  };
}

/**
 * Flatten a retired rename-source key.
 * Same shape as a closed last row in the main seed walk.
 * Does not delete the key — incremental Step 4 may see a later lot
 * under that old name and needs next-TG ready.
 */
function flattenRetiredRenameSource_(blocks, oldKey, blockNumHint) {
  if (!oldKey || oldKey.charAt(oldKey.length - 1) === "|") return;
  if (!blocks[oldKey]) blocks[oldKey] = emptyBlockState();
  const b = blocks[oldKey];
  const blockNum = Number(blockNumHint || b.block || 1) || 1;
  b.unit = 0;
  b.runningQty = 0;
  b.block = blockNum + 1;
  b.pnl = 0;
  b.entryCost = 0;
  b.openTs = null;
  b.positionId = "";
}

/**
 * Read Master and return a blocks{} object the Step 4 loop can resume from.
 *
 * Closed last row (Block Close Flag/P&L = 1, or running qty ~ 0):
 *   unit 0, runningQty 0, block = last Block Number + 1
 *
 * Open last row:
 *   unit + runningQty from Running Position Quantity
 *   block from Block Number (fallback: parse Trade Group ID)
 *
 * Does not write Staging or Master.
 */
function seedBlocksFromMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("Master");
  if (!master) throw new Error('Sheet "Master" not found.');

  const tz = ss.getSpreadsheetTimeZone();
  const lastRow = master.getLastRow();
  const lastCol = master.getLastColumn();
  if (lastRow < 2) return {};

  const grid = master.getRange(1, 1, lastRow, lastCol).getValues();
  const col = {};
  grid[0].forEach(function (h, i) {
    const norm = String(h || "")
      .trim()
      .toLowerCase();
    if (norm) col[norm] = i;
  });

  const need = ["account", "ticker", "trade type", "running position quantity"];
  for (let i = 0; i < need.length; i++) {
    if (col[need[i]] === undefined) {
      throw new Error(
        "seedBlocksFromMaster: Master is missing header " + need[i],
      );
    }
  }

  const blocks = {};
  const retiredByRename = {};
  const posIdLastStockKey = {};

  for (let r = 1; r < grid.length; r++) {
    const row = grid[r];
    const acct = String(row[col["account"]] || "")
      .trim()
      .toUpperCase();
    if (!acct) continue;

    const key = blockKeyFromStagingLikeRow_(row, col, tz);
    if (!key || key === "|" || key.charAt(key.length - 1) === "|") continue;

    if (!blocks[key]) blocks[key] = emptyBlockState();
    const b = blocks[key];

    const runQty = Number(row[col["running position quantity"]]) || 0;
    const closeRaw =
      col["block close flag/p&l"] !== undefined
        ? row[col["block close flag/p&l"]]
        : 0;
    const startRaw =
      col["block start flag"] !== undefined ? row[col["block start flag"]] : 0;
    const isClose = closeRaw === 1 || closeRaw === "1";
    const isStart = startRaw === 1 || startRaw === "1";

    let blockNum = 0;
    if (col["block number"] !== undefined) {
      blockNum = Number(row[col["block number"]]) || 0;
    }
    if (!blockNum && col["trade group id"] !== undefined) {
      blockNum = parseTgBlockNumber_(row[col["trade group id"]]);
    }

    if (isStart) {
      b.openTs =
        col["trade time stamp"] !== undefined
          ? row[col["trade time stamp"]]
          : null;
      if (col["position id"] !== undefined) {
        b.positionId = String(row[col["position id"]] || "");
      }
      if (col["strategy type"] !== undefined) {
        b.strategyType = String(row[col["strategy type"]] || "");
      }
    }

    if (col["trade group id"] !== undefined) {
      b.tradeGroupId = String(row[col["trade group id"]] || "");
    }
    if (col["position id"] !== undefined && !isClose) {
      b.positionId = String(row[col["position id"]] || b.positionId || "");
    }
    if (col["strategy type"] !== undefined && !isClose) {
      b.strategyType = String(
        row[col["strategy type"]] || b.strategyType || "",
      );
    }

    const flat = isClose || Math.abs(runQty) < 1e-8;
    if (flat) {
      b.unit = 0;
      b.runningQty = 0;
      b.block = (blockNum || b.block || 1) + 1;
      b.pnl = 0;
      b.entryCost = 0;
      b.openTs = null;
      b.positionId = "";
    } else {
      b.unit = runQty;
      b.runningQty = runQty;
      b.block = blockNum || b.block || 1;
    }

    const sc = parseSymbolChangeFromMasterRow_(row, col);
    if (sc) {
      const rowTkr = String(row[col["ticker"]] || "")
        .trim()
        .toUpperCase();
      const destTicker = sc.toResolved || sc.toRaw || rowTkr;
      const destSet = {};
      if (sc.toResolved) destSet[sc.toResolved] = true;
      if (sc.toRaw) destSet[sc.toRaw] = true;
      if (rowTkr) destSet[rowTkr] = true;

      const sources = [sc.fromRaw, sc.fromResolved];
      for (let s = 0; s < sources.length; s++) {
        const src = sources[s];
        if (!src || destSet[src]) continue;
        retiredByRename[acct + "|" + src] = destTicker;
      }
    }

    if (isStockAcctTickerKey_(key)) {
      const posId =
        col["position id"] !== undefined
          ? String(row[col["position id"]] || "").trim()
          : "";
      const tgId =
        col["trade group id"] !== undefined
          ? String(row[col["trade group id"]] || "").trim()
          : "";
      const ident = posId || tgId;
      if (ident) {
        const pidKey = acct + "::" + ident;
        const prev = posIdLastStockKey[pidKey];
        if (prev && prev !== key && isStockAcctTickerKey_(prev)) {
          retiredByRename[prev] = key.split("|")[1];
        }
        posIdLastStockKey[pidKey] = key;
      }
    }
  }

  Object.keys(FAMILY_S_TICKER_ALIASES_).forEach(function (srcTkr) {
    const destTkr = FAMILY_S_TICKER_ALIASES_[srcTkr];
    Object.keys(blocks).forEach(function (key) {
      if (!isStockAcctTickerKey_(key)) return;
      const parts = key.split("|");
      if (parts[1] !== srcTkr) return;
      const destKey = parts[0] + "|" + destTkr;
      if (blocks[destKey]) {
        retiredByRename[key] = destTkr;
      }
    });
  });

  const retiredKeys = Object.keys(retiredByRename);
  for (let i = 0; i < retiredKeys.length; i++) {
    const oldKey = retiredKeys[i];
    const hint = blocks[oldKey] ? blocks[oldKey].block : 1;
    flattenRetiredRenameSource_(blocks, oldKey, hint);
  }

  SEED_RETIRED_BY_RENAME_ = retiredByRename;
  return blocks;
}

/**
 * Menu helper. Reads Master, builds seed, writes Block Seed Preview.
 * Does not run populateStagingWithBlockLogicV3 and does not touch Master data.
 */
function inspectSeedBlocksFromMaster() {
  const t0 = pipelineTimingNow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("Master");
  if (!master) {
    uiAlertSafe("Master sheet not found. Load Master from Staging first.");
    return;
  }
  if (master.getLastRow() < 2) {
    uiAlertSafe(
      "Master has no data rows. Run Replace Master from Staging (first load) after a clean Staging.",
    );
    return;
  }

  const blocks = seedBlocksFromMaster();
  const retiredByRename = SEED_RETIRED_BY_RENAME_ || {};

  const keys = Object.keys(blocks);
  let openCount = 0;
  let closedCount = 0;
  let retiredCount = 0;
  const preview = [
    [
      "Block Key",
      "LiveOrClosed",
      "unit",
      "runningQty",
      "block (next TG if closed)",
      "positionId",
      "strategyType",
      "tradeGroupId",
      "Family S note",
    ],
  ];

  keys.sort();
  for (let i = 0; i < keys.length; i++) {
    const b = blocks[keys[i]];
    const live = Number(b.unit || 0) !== 0 || Number(b.runningQty || 0) !== 0;
    if (live) openCount++;
    else closedCount++;
    const renamedTo = retiredByRename[keys[i]];
    let note = "";
    if (renamedTo) {
      retiredCount++;
      note = "SYMBOL CHANGE retired this key → " + renamedTo;
    }
    preview.push([
      keys[i],
      live ? "LIVE" : "CLOSED",
      b.unit,
      b.runningQty,
      b.block,
      b.positionId || "",
      b.strategyType || "",
      b.tradeGroupId || "",
      note,
    ]);
  }

  let sh = ss.getSheetByName("Block Seed Preview");
  if (!sh) sh = ss.insertSheet("Block Seed Preview");
  sh.clearContents();
  sh.getRange(1, 1, preview.length, preview[0].length).setValues(preview);
  sh.setFrozenRows(1);

  pipelineTimingLog(
    "inspectSeedBlocksFromMaster",
    t0,
    "keys=" +
      keys.length +
      " live=" +
      openCount +
      " closed=" +
      closedCount +
      " familySRetired=" +
      retiredCount,
  );

  uiAlertSafe(
    "Block seed preview written.\n\n" +
      "Keys: " +
      keys.length +
      "\nLive: " +
      openCount +
      "\nClosed (next TG ready): " +
      closedCount +
      "\nFamily S retired rename sources: " +
      retiredCount +
      "\n\nOpen sheet Block Seed Preview.\n" +
      "Look at LT|ENCUF, LT|ISENF, LT|SV — should be CLOSED.\n" +
      "Dest keys LT|EU, LT|ISOU, LT|SMR keep their own last-row state.\n\n" +
      "Full rebuild is unchanged — populateStagingWithBlockLogicV3() still starts empty.",
  );
}

/**
 * Read-only Phase 4 cutoff inspect.
 * Does not write Master, Staging, or Helper data rows.
 * Writes sheet "Incremental Delta Preview" (at-or-after cutoff rows only).
 *
 * After a first-load replace, helperAfter must be 0.
 * helperAtLastTs is the same-second bucket — timestamp alone cannot
 * tell those Helper rows from the Master rows already loaded.
 */
function inspectIncrementalDeltaFromMaster() {
  const t0 = pipelineTimingNow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const ui = SpreadsheetApp.getUi();

  const master = ss.getSheetByName("Master");
  const helper = ss.getSheetByName("Helper");
  if (!master) {
    uiAlertSafe("Master sheet not found.");
    return;
  }
  if (!helper) {
    uiAlertSafe("Helper sheet not found.");
    return;
  }

  function headerMap_(grid) {
    const col = {};
    grid[0].forEach(function (h, i) {
      const norm = String(h || "")
        .trim()
        .toLowerCase();
      if (norm) col[norm] = i;
    });
    return col;
  }

  function rowTs_(row, col) {
    const tsVal =
      col["trade time stamp"] !== undefined
        ? row[col["trade time stamp"]]
        : null;
    const dateVal =
      col["trade date"] !== undefined ? row[col["trade date"]] : null;
    const timeVal =
      col["trade time"] !== undefined ? row[col["trade time"]] : null;
    return parseTradeTimeStamp(tsVal, dateVal, timeVal, ss);
  }

  function cell_(row, col, name) {
    return col[name] !== undefined ? row[col[name]] : "";
  }

  const mLast = master.getLastRow();
  const mLastCol = master.getLastColumn();
  if (mLast < 2) {
    uiAlertSafe("Master has no data rows.");
    return;
  }
  const mGrid = master.getRange(1, 1, mLast, mLastCol).getValues();
  const mCol = headerMap_(mGrid);
  if (mCol["trade time stamp"] === undefined) {
    uiAlertSafe("Master is missing header Trade Time Stamp.");
    return;
  }

  let lastMasterTs = null;
  let masterAtLast = 0;
  let masterDataRows = 0;
  for (let r = 1; r < mGrid.length; r++) {
    const acct = String(mGrid[r][mCol["account"] || 0] || "").trim();
    if (!acct) continue;
    masterDataRows++;
    const ts = rowTs_(mGrid[r], mCol);
    if (!ts) continue;
    if (!lastMasterTs || ts.getTime() > lastMasterTs.getTime()) {
      lastMasterTs = ts;
    }
  }
  if (!lastMasterTs) {
    uiAlertSafe("Master has no parseable Trade Time Stamp.");
    return;
  }
  const lastMs = lastMasterTs.getTime();
  for (let r = 1; r < mGrid.length; r++) {
    const ts = rowTs_(mGrid[r], mCol);
    if (ts && ts.getTime() === lastMs) masterAtLast++;
  }

  const hLast = helper.getLastRow();
  const hLastCol = helper.getLastColumn();
  if (hLast < 2) {
    uiAlertSafe("Helper has no data rows.");
    return;
  }
  const hGrid = helper.getRange(1, 1, hLast, hLastCol).getValues();
  const hCol = headerMap_(hGrid);
  if (hCol["trade time stamp"] === undefined) {
    uiAlertSafe("Helper is missing header Trade Time Stamp.");
    return;
  }

  let helperDataRows = 0;
  let helperAfter = 0;
  let helperAtLast = 0;
  let lastHelperTs = null;
  const preview = [
    [
      "Which",
      "Helper Row",
      "Account",
      "Ticker",
      "Action",
      "Trade Time Stamp",
      "Quantity",
      "Option Strike",
      "Call/Put",
      "Position ID",
    ],
  ];

  for (let r = 1; r < hGrid.length; r++) {
    const row = hGrid[r];
    const acct = String(cell_(row, hCol, "account") || "").trim();
    if (!acct) continue;
    helperDataRows++;
    const ts = rowTs_(row, hCol);
    if (!ts) continue;
    if (!lastHelperTs || ts.getTime() > lastHelperTs.getTime()) {
      lastHelperTs = ts;
    }
    const ms = ts.getTime();
    let which = "";
    if (ms > lastMs) {
      helperAfter++;
      which = "AFTER";
    } else if (ms === lastMs) {
      helperAtLast++;
      which = "AT_LAST_TS";
    } else {
      continue;
    }
    preview.push([
      which,
      r + 1,
      acct,
      cell_(row, hCol, "ticker"),
      cell_(row, hCol, "action"),
      ts,
      cell_(row, hCol, "quantity"),
      cell_(row, hCol, "option strike"),
      cell_(row, hCol, "call/put"),
      cell_(row, hCol, "position id"),
    ]);
  }

  let sh = ss.getSheetByName("Incremental Delta Preview");
  if (!sh) sh = ss.insertSheet("Incremental Delta Preview");
  sh.clearContents();
  sh.getRange(1, 1, preview.length, preview[0].length).setValues(preview);
  sh.setFrozenRows(1);

  const lastMasterTxt = Utilities.formatDate(
    lastMasterTs,
    tz,
    "M/d/yyyy HH:mm:ss",
  );
  const lastHelperTxt = lastHelperTs
    ? Utilities.formatDate(lastHelperTs, tz, "M/d/yyyy HH:mm:ss")
    : "(none)";

  pipelineTimingLog(
    "inspectIncrementalDeltaFromMaster",
    t0,
    "masterData=" +
      masterDataRows +
      " helperData=" +
      helperDataRows +
      " helperAfter=" +
      helperAfter +
      " helperAtLastTs=" +
      helperAtLast +
      " masterAtLastTs=" +
      masterAtLast,
  );

  uiAlertSafe(
    "Incremental delta (read-only).\n\n" +
      "Master last ts: " +
      lastMasterTxt +
      "\n" +
      "Helper last ts: " +
      lastHelperTxt +
      "\n" +
      "Master data rows: " +
      masterDataRows +
      "\n" +
      "Helper data rows: " +
      helperDataRows +
      "\n" +
      "Helper AFTER last Master ts: " +
      helperAfter +
      "\n" +
      "Helper AT last Master ts: " +
      helperAtLast +
      "\n" +
      "Master AT last ts: " +
      masterAtLast +
      "\n\n" +
      "After a first load, AFTER must be 0.\n" +
      "AT_LAST_TS rows are the same-second bucket.\n" +
      "Open Incremental Delta Preview.",
  );
}

/**
 * Stable identity for “is this Helper row already on Master?”
 * Not a Position ID (Helper does not have Step 4 ids yet).
 * Not timestamp alone (same-second fills exist in history).
 *
 * Account|ms|Ticker|Action|qty|strike|C/P
 */
function incrementalRowFingerprint_(row, col) {
  function cell(name) {
    return col[name] !== undefined ? row[col[name]] : "";
  }
  const acct = String(cell("account") || "")
    .trim()
    .toUpperCase();
  const ticker = String(cell("ticker") || "")
    .trim()
    .toUpperCase();
  const action = String(cell("action") || "")
    .trim()
    .toUpperCase();
  const ts = (function () {
    const tsVal = cell("trade time stamp");
    if (tsVal instanceof Date && !isNaN(tsVal.getTime()))
      return tsVal.getTime();
    return "";
  })();
  const qtyN = toNum(cell("quantity"));
  const qty = isNaN(qtyN) ? "" : String(Math.round(qtyN * 1e8) / 1e8);
  const strikeN = toNum(cell("option strike"));
  const strike = isNaN(strikeN) ? "" : String(Math.round(strikeN * 1e4) / 1e4);
  let cp = String(cell("call/put") || "")
    .trim()
    .toUpperCase()
    .replace("CALL", "C")
    .replace("PUT", "P");
  if (cp.charAt(0) === "C") cp = "C";
  else if (cp.charAt(0) === "P") cp = "P";
  return [acct, ts, ticker, action, qty, strike, cp].join("|");
}

/**
 * Read-only. Builds the set of Helper rows an incremental runner
 * would send into seeded block logic.
 *
 * New = Helper ts > last Master ts
 *    OR Helper ts == last Master ts AND fingerprint not on Master
 *       at that same timestamp.
 *
 * Writes Incremental Candidates Preview (AFTER + AT_LAST_TS rows).
 * Does not write Helper, Staging, or Master data.
 * Does not call populateStagingWithBlockLogicV3.
 */
function previewIncrementalCandidatesFromHelper() {
  const t0 = pipelineTimingNow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  const master = ss.getSheetByName("Master");
  const helper = ss.getSheetByName("Helper");
  if (!master) {
    uiAlertSafe("Master sheet not found.");
    return;
  }
  if (!helper) {
    uiAlertSafe("Helper sheet not found.");
    return;
  }

  function headerMap_(grid) {
    const col = {};
    grid[0].forEach(function (h, i) {
      const norm = String(h || "")
        .trim()
        .toLowerCase();
      if (norm) col[norm] = i;
    });
    return col;
  }

  function rowTs_(row, col) {
    const tsVal =
      col["trade time stamp"] !== undefined
        ? row[col["trade time stamp"]]
        : null;
    const dateVal =
      col["trade date"] !== undefined ? row[col["trade date"]] : null;
    const timeVal =
      col["trade time"] !== undefined ? row[col["trade time"]] : null;
    return parseTradeTimeStamp(tsVal, dateVal, timeVal, ss);
  }

  function cell_(row, col, name) {
    return col[name] !== undefined ? row[col[name]] : "";
  }

  const mLast = master.getLastRow();
  const mLastCol = master.getLastColumn();
  if (mLast < 2) {
    uiAlertSafe("Master has no data rows.");
    return;
  }
  const mGrid = master.getRange(1, 1, mLast, mLastCol).getValues();
  const mCol = headerMap_(mGrid);

  let lastMasterTs = null;
  for (let r = 1; r < mGrid.length; r++) {
    const acct = String(mGrid[r][mCol["account"] || 0] || "").trim();
    if (!acct) continue;
    const ts = rowTs_(mGrid[r], mCol);
    if (!ts) continue;
    if (!lastMasterTs || ts.getTime() > lastMasterTs.getTime()) {
      lastMasterTs = ts;
    }
  }
  if (!lastMasterTs) {
    uiAlertSafe("Master has no parseable Trade Time Stamp.");
    return;
  }
  const lastMs = lastMasterTs.getTime();

  const masterFpAtLast = {};
  for (let r = 1; r < mGrid.length; r++) {
    const ts = rowTs_(mGrid[r], mCol);
    if (!ts || ts.getTime() !== lastMs) continue;
    masterFpAtLast[incrementalRowFingerprint_(mGrid[r], mCol)] = r + 1;
  }

  const hLast = helper.getLastRow();
  const hLastCol = helper.getLastColumn();
  if (hLast < 2) {
    uiAlertSafe("Helper has no data rows.");
    return;
  }
  const hGrid = helper.getRange(1, 1, hLast, hLastCol).getValues();
  const hCol = headerMap_(hGrid);

  let afterNew = 0;
  let atMatched = 0;
  let atUnmatched = 0;
  const preview = [
    [
      "Status",
      "Candidate",
      "Helper Row",
      "Fingerprint",
      "Master Row At Last Ts",
      "Account",
      "Ticker",
      "Action",
      "Trade Time Stamp",
      "Quantity",
      "Option Strike",
      "Call/Put",
    ],
  ];

  for (let r = 1; r < hGrid.length; r++) {
    const row = hGrid[r];
    const acct = String(cell_(row, hCol, "account") || "").trim();
    if (!acct) continue;
    const ts = rowTs_(row, hCol);
    if (!ts) continue;
    const ms = ts.getTime();
    if (ms < lastMs) continue;

    const fp = incrementalRowFingerprint_(row, hCol);
    let status = "";
    let candidate = "N";
    let masterRow = "";

    if (ms > lastMs) {
      status = "AFTER_NEW";
      candidate = "Y";
      afterNew++;
    } else {
      if (masterFpAtLast[fp]) {
        status = "AT_LAST_TS_MATCHED";
        candidate = "N";
        masterRow = masterFpAtLast[fp];
        atMatched++;
      } else {
        status = "AT_LAST_TS_UNMATCHED";
        candidate = "Y";
        atUnmatched++;
      }
    }

    preview.push([
      status,
      candidate,
      r + 1,
      fp,
      masterRow,
      acct,
      cell_(row, hCol, "ticker"),
      cell_(row, hCol, "action"),
      ts,
      cell_(row, hCol, "quantity"),
      cell_(row, hCol, "option strike"),
      cell_(row, hCol, "call/put"),
    ]);
  }

  const candidates = afterNew + atUnmatched;

  let sh = ss.getSheetByName("Incremental Candidates Preview");
  if (!sh) sh = ss.insertSheet("Incremental Candidates Preview");
  sh.clearContents();
  sh.getRange(1, 1, preview.length, preview[0].length).setValues(preview);
  sh.setFrozenRows(1);

  pipelineTimingLog(
    "previewIncrementalCandidatesFromHelper",
    t0,
    "afterNew=" +
      afterNew +
      " atMatched=" +
      atMatched +
      " atUnmatched=" +
      atUnmatched +
      " candidates=" +
      candidates,
  );

  uiAlertSafe(
    "Incremental candidates (read-only).\n\n" +
      "Master last ts: " +
      Utilities.formatDate(lastMasterTs, tz, "M/d/yyyy HH:mm:ss") +
      "\n" +
      "AFTER new: " +
      afterNew +
      "\n" +
      "AT_LAST_TS matched: " +
      atMatched +
      "\n" +
      "AT_LAST_TS unmatched: " +
      atUnmatched +
      "\n" +
      "Candidates (would enter incremental Step 4): " +
      candidates +
      "\n\n" +
      "On the 9/1/2026 16:41:18 freeze this must be 0 / 1 / 0 / 0.\n" +
      "Open Incremental Candidates Preview.\n" +
      "Does not write Staging or Master.",
  );
}

/**
 * Shared cutoff + fingerprint walk used by the dry-run runner.
 * Same rules as previewIncrementalCandidatesFromHelper.
 * Returns null after an alert when Master/Helper cannot be read.
 *
 * candidateRows = Helper rows that would enter incremental Step 4.
 * Preview sheet is NOT written here.
 */
function collectIncrementalCandidates_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("Master");
  const helper = ss.getSheetByName("Helper");
  if (!master) {
    uiAlertSafe("Master sheet not found.");
    return null;
  }
  if (!helper) {
    uiAlertSafe("Helper sheet not found.");
    return null;
  }

  function headerMap_(grid) {
    const col = {};
    grid[0].forEach(function (h, i) {
      const norm = String(h || "")
        .trim()
        .toLowerCase();
      if (norm) col[norm] = i;
    });
    return col;
  }

  function rowTs_(row, col) {
    const tsVal =
      col["trade time stamp"] !== undefined
        ? row[col["trade time stamp"]]
        : null;
    const dateVal =
      col["trade date"] !== undefined ? row[col["trade date"]] : null;
    const timeVal =
      col["trade time"] !== undefined ? row[col["trade time"]] : null;
    return parseTradeTimeStamp(tsVal, dateVal, timeVal, ss);
  }

  function cell_(row, col, name) {
    return col[name] !== undefined ? row[col[name]] : "";
  }

  const mLast = master.getLastRow();
  const mLastCol = master.getLastColumn();
  if (mLast < 2) {
    uiAlertSafe("Master has no data rows.");
    return null;
  }
  const mGrid = master.getRange(1, 1, mLast, mLastCol).getValues();
  const mCol = headerMap_(mGrid);

  let lastMasterTs = null;
  for (let r = 1; r < mGrid.length; r++) {
    const acct = String(mGrid[r][mCol["account"] || 0] || "").trim();
    if (!acct) continue;
    const ts = rowTs_(mGrid[r], mCol);
    if (!ts) continue;
    if (!lastMasterTs || ts.getTime() > lastMasterTs.getTime()) {
      lastMasterTs = ts;
    }
  }
  if (!lastMasterTs) {
    uiAlertSafe("Master has no parseable Trade Time Stamp.");
    return null;
  }
  const lastMs = lastMasterTs.getTime();

  const masterFpAtLast = {};
  for (let r = 1; r < mGrid.length; r++) {
    const ts = rowTs_(mGrid[r], mCol);
    if (!ts || ts.getTime() !== lastMs) continue;
    masterFpAtLast[incrementalRowFingerprint_(mGrid[r], mCol)] = r + 1;
  }

  const hLast = helper.getLastRow();
  const hLastCol = helper.getLastColumn();
  if (hLast < 2) {
    uiAlertSafe("Helper has no data rows.");
    return null;
  }
  const hGrid = helper.getRange(1, 1, hLast, hLastCol).getValues();
  const hCol = headerMap_(hGrid);

  let afterNew = 0;
  let atMatched = 0;
  let atUnmatched = 0;
  const candidateRows = [];

  for (let r = 1; r < hGrid.length; r++) {
    const row = hGrid[r];
    const acct = String(cell_(row, hCol, "account") || "").trim();
    if (!acct) continue;
    const ts = rowTs_(row, hCol);
    if (!ts) continue;
    const ms = ts.getTime();
    if (ms < lastMs) continue;

    const fp = incrementalRowFingerprint_(row, hCol);
    if (ms > lastMs) {
      afterNew++;
      candidateRows.push({
        helperRow: r + 1,
        status: "AFTER_NEW",
        fingerprint: fp,
        ts: ts,
      });
    } else if (masterFpAtLast[fp]) {
      atMatched++;
    } else {
      atUnmatched++;
      candidateRows.push({
        helperRow: r + 1,
        status: "AT_LAST_TS_UNMATCHED",
        fingerprint: fp,
        ts: ts,
      });
    }
  }

  return {
    lastMasterTs: lastMasterTs,
    afterNew: afterNew,
    atMatched: atMatched,
    atUnmatched: atUnmatched,
    candidates: afterNew + atUnmatched,
    candidateRows: candidateRows,
  };
}

/**
 * Phase 4 dry run. Does not write Helper, Staging, or Master.
 * Does not call populateStagingWithBlockLogicV3.
 *
 * On the 9/1/2026 16:41:18 freeze this must report candidates = 0
 * and then return. Seed is loaded only to prove seedBlocksFromMaster
 * still runs. Live seed count is logged; nothing is applied to Staging.
 *
 * If candidates > 0 the function STOPS. That is new work, not this freeze.
 */
function runIncrementalFromHelperDryRun() {
  const t0 = pipelineTimingNow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  const collected = collectIncrementalCandidates_();
  if (!collected) return;

  const lastTxt = Utilities.formatDate(
    collected.lastMasterTs,
    tz,
    "M/d/yyyy HH:mm:ss",
  );

  if (collected.candidates !== 0) {
    pipelineTimingLog(
      "runIncrementalFromHelperDryRun STOPPED",
      t0,
      "candidates=" + collected.candidates,
    );
    uiAlertSafe(
      "Incremental dry run STOPPED.\n\n" +
        "Master last ts: " +
        lastTxt +
        "\n" +
        "AFTER new: " +
        collected.afterNew +
        "\n" +
        "AT_LAST_TS unmatched: " +
        collected.atUnmatched +
        "\n" +
        "Candidates: " +
        collected.candidates +
        "\n\n" +
        "This freeze should be 0. Do not write Staging or Master.\n" +
        "Run Preview incremental candidates and paste that sheet.",
    );
    return;
  }

  const blocks = seedBlocksFromMaster();
  const keys = Object.keys(blocks);
  let live = 0;
  for (let i = 0; i < keys.length; i++) {
    const b = blocks[keys[i]];
    if (Number(b.unit || 0) !== 0 || Number(b.runningQty || 0) !== 0) {
      live++;
    }
  }

  pipelineTimingLog(
    "runIncrementalFromHelperDryRun",
    t0,
    "candidates=0 seedKeys=" + keys.length + " live=" + live,
  );

  uiAlertSafe(
    "Incremental dry run OK.\n\n" +
      "Master last ts: " +
      lastTxt +
      "\n" +
      "Candidates: 0\n" +
      "AFTER new: 0\n" +
      "AT_LAST_TS matched: " +
      collected.atMatched +
      "\n" +
      "Seed keys: " +
      keys.length +
      "\n" +
      "Seed live: " +
      live +
      "\n\n" +
      "No Helper / Staging / Master write.\n" +
      "populateStagingWithBlockLogicV3 was not called.\n" +
      "Expect seed keys 2514 and live 82 on this freeze.",
  );
}

/**
 * Empty-output proof for Phase 4.
 * Writes sheet "Incremental Seeded Preview" with Helper header rows 1-3
 * and zero data rows when candidates = 0.
 *
 * Does not write Helper, Staging, or Master.
 * Does not call populateStagingWithBlockLogicV3.
 *
 * Why not call Step 4: that function always reads all of Helper and
 * replace-writes all of Staging. Seed + filter inside it is a later slice.
 *
 * If candidates > 0 this STOPS after the same header-only write so
 * Staging cannot be touched by accident.
 */
function previewIncrementalSeededOutputFromHelper() {
  const t0 = pipelineTimingNow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  const collected = collectIncrementalCandidates_();
  if (!collected) return;

  const helper = ss.getSheetByName("Helper");
  const lastCol = helper.getLastColumn();
  const headerGrid = helper.getRange(1, 1, 3, lastCol).getValues();

  let sh = ss.getSheetByName("Incremental Seeded Preview");
  if (!sh) sh = ss.insertSheet("Incremental Seeded Preview");
  sh.clearContents();
  sh.getRange(1, 1, 3, lastCol).setValues(headerGrid);
  sh.setFrozenRows(3);

  const lastTxt = Utilities.formatDate(
    collected.lastMasterTs,
    tz,
    "M/d/yyyy HH:mm:ss",
  );
  const blocks = seedBlocksFromMaster();
  const keys = Object.keys(blocks);
  let live = 0;
  for (let i = 0; i < keys.length; i++) {
    const b = blocks[keys[i]];
    if (Number(b.unit || 0) !== 0 || Number(b.runningQty || 0) !== 0) {
      live++;
    }
  }

  // Rows 2-3 on Helper/Staging are blank spacers. getLastRow() skips
  // them and reports 1 even after we write 3 header rows. Frozen rows
  // and the range we set are the real proof.
  const wroteHeaderRows = 3;
  const previewLast = sh.getLastRow();
  const previewFrozen = sh.getFrozenRows();

  pipelineTimingLog(
    "previewIncrementalSeededOutputFromHelper",
    t0,
    "candidates=" +
      collected.candidates +
      " wroteHeaderRows=" +
      wroteHeaderRows +
      " previewGetLastRow=" +
      previewLast +
      " frozen=" +
      previewFrozen +
      " seedKeys=" +
      keys.length +
      " live=" +
      live,
  );

  if (collected.candidates !== 0) {
    uiAlertSafe(
      "Incremental seeded preview STOPPED.\n\n" +
        "Master last ts: " +
        lastTxt +
        "\n" +
        "Candidates: " +
        collected.candidates +
        "\n" +
        "AFTER new: " +
        collected.afterNew +
        "\n" +
        "AT_LAST_TS unmatched: " +
        collected.atUnmatched +
        "\n\n" +
        "Wrote Helper headers only to Incremental Seeded Preview.\n" +
        "Did not run block logic. Staging and Master were not written.",
    );
    return;
  }

  uiAlertSafe(
    "Incremental seeded preview OK.\n\n" +
      "Master last ts: " +
      lastTxt +
      "\n" +
      "Candidates: 0\n" +
      "Wrote header rows: " +
      wroteHeaderRows +
      "\n" +
      "Frozen rows: " +
      previewFrozen +
      " (must be 3)\n" +
      "getLastRow: " +
      previewLast +
      " (1 is OK if rows 2-3 are blank)\n" +
      "Seed keys: " +
      keys.length +
      "\n" +
      "Seed live: " +
      live +
      "\n\n" +
      "No Helper / Staging / Master write.\n" +
      "populateStagingWithBlockLogicV3 was not called.\n" +
      "Expect frozen 3, keys 2514, live 82.",
  );
}
