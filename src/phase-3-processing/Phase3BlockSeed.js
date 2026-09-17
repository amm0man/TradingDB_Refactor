/**
 * Phase3BlockSeed.js
 *
 * Phase 4 seed helpers. Not wired into the daily full rebuild.
 *
 *   emptyBlockState()              — same object the Step 4 loop already creates
 *   blockKeyFromStagingLikeRow_()  — same key rules as populateStagingWithBlockLogicV3
 *   seedBlocksFromMaster()         — rebuild blocks{} from Master last-row state
 *   inspectSeedBlocksFromMaster()  — menu: preview only (writes Block Seed Preview)
 *
 * populateStagingWithBlockLogicV3(seedBlocks) still starts from {} when
 * called with no argument. Passing a seed is for a later incremental function.
 */

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
 *           DT-UNG-SP-221230-TG001-ST → 1
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

  const need = [
    "account",
    "ticker",
    "trade type",
    "running position quantity",
  ];
  for (let i = 0; i < need.length; i++) {
    if (col[need[i]] === undefined) {
      throw new Error(
        "seedBlocksFromMaster: Master is missing header " + need[i],
      );
    }
  }

  const blocks = {};

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
      col["block start flag"] !== undefined
        ? row[col["block start flag"]]
        : 0;
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
      b.strategyType = String(row[col["strategy type"]] || b.strategyType || "");
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
  }

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
  const keys = Object.keys(blocks);
  let openCount = 0;
  let closedCount = 0;
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
    ],
  ];

  keys.sort();
  for (let i = 0; i < keys.length; i++) {
    const b = blocks[keys[i]];
    const live = Number(b.unit || 0) !== 0 || Number(b.runningQty || 0) !== 0;
    if (live) openCount++;
    else closedCount++;
    preview.push([
      keys[i],
      live ? "LIVE" : "CLOSED",
      b.unit,
      b.runningQty,
      b.block,
      b.positionId || "",
      b.strategyType || "",
      b.tradeGroupId || "",
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
      closedCount,
  );

  uiAlertSafe(
    "Block seed preview written.\n\n" +
      "Keys: " +
      keys.length +
      "\nLive: " +
      openCount +
      "\nClosed (next TG ready): " +
      closedCount +
      "\n\nOpen sheet Block Seed Preview.\n" +
      "Full rebuild is unchanged — populateStagingWithBlockLogicV3() still starts empty.",
  );
}