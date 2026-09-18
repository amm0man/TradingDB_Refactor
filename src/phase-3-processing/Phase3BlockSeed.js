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
 *   Seed instead: when a Master row is a SYMBOL CHANGE, flatten acct|FROM after
 *   applying the dest-ticker last-row state. Inspect then shows LT|ENCUF,
 *   LT|ISENF, LT|SV as CLOSED. The dest key (EU / ISOU / SMR) stays whatever
 *   its own last row says.
 */

/**
 * Inspect-only map of Family S retired keys. seedBlocksFromMaster fills this.
 * Must never be copied onto blocks{} — Object.keys(blocks) would treat it
 * as a live block key in a later incremental Step 4.
 */
var SEED_RETIRED_BY_RENAME_ = {};

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

  const isRename = action === "SYMBOL CHANGE" || corp === "SYMBOL CHANGE";
  if (!isRename) return null;

  const notes =
    col["notes"] !== undefined ? String(row[col["notes"]] || "") : "";
  if (!notes) return null;

  function pull(label) {
    const m = notes.match(new RegExp(label + "=([^|]+)", "i"));
    return m ? String(m[1]).trim().toUpperCase() : "";
  }

  const fromRaw = pull("FROM");
  const toRaw = pull("TO");
  const fromResolved = pull("FROM_RESOLVED");
  const toResolved = pull("TO_RESOLVED");

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
 *   (matches the live blocks[key].block++ after blkClose)
 *
 * Open last row:
 *   unit + runningQty from Running Position Quantity
 *   block from Block Number (fallback: parse Trade Group ID)
 *   positionId / strategyType / tradeGroupId / openTs from last Block Start
 *
 * Family S: after applying the dest-ticker row, flatten acct|FROM from Notes.
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
  // acct|OLD → NEW  (inspect preview only; never stored on blocks{})
  const retiredByRename = {};

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

    // Family S — collect rename SOURCE tokens only.
    // WHY ENCUF closed but ISENF/SV stayed LIVE on the first pass:
    // fromTicker used FROM_RESOLVED first. After Phase 2 alias,
    // FROM_RESOLVED is often already the dest (ISOU / SMR), so we
    // flattened the wrong key (or skipped because oldKey === dest key).
    // FROM=ISENF / FROM=SV is the sheet history key that must retire.
    // Flatten AFTER this walk so a later leftover old-ticker row cannot
    // reopen the source key.
    // Family S pass 2 — Position ID / Trade Group ID continuity.
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
  }

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
  const retiredByRename = {};
  const posIdLastStockKey = {};

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
