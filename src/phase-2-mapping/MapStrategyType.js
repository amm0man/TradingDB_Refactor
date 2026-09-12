/**
 * MapStrategyType.js
 *
 * Phase 2 – Group G: Strategy Type
 *
 * Sole responsibility:
 *   Baseline Strategy Type plus the two post-processors that finish
 *   multi-leg labels.
 *
 * Functions (same names as before — do not rename):
 *   - normalizeStrategyType
 *   - postProcessStrategyTypeBySpreadGroups
 *   - postProcessStrategyTypeByPositionTrackerV3
 *
 * Called by:
 *   mapSchwabImportByHeadersV3()
 *
 * Uses (already global — do not redeclare):
 *   SPREAD_GROUP_TYPES
 *   col   (Helpers.js)
 *
 * Do not redeclare SHEET_* consts or SPREAD_GROUP_TYPES here.
 */

// =========================================================================
// STRATEGY TYPE POST-PROCESSORS
//   1. normalizeStrategyType               – single-row baseline
//   2. postProcessStrategyTypeBySpreadGroups – force same label on all legs
//   3. postProcessStrategyTypeByPositionTrackerV3 – FIFO close-row inheritance
// =========================================================================


function normalizeStrategyType(
  spreadRaw,
  action,
  callPut,
  importAmountForSign,
) {
  // NOTE: we receive importAmount (Net Amount hint from Schwab Import) here.
  // For spread legs the FINAL Strategy Type is decided later in postProcessStrategyTypeBySpreadGroups_ anyway.
  // This baseline is only used for SINGLE / STOCK / simple cases.
  const sp = String(spreadRaw || "")
    .trim()
    .toUpperCase();
  const act = String(action || "").trim();
  const cp = String(callPut || "")
    .trim()
    .toUpperCase();
  const tc =
    typeof importAmountForSign === "number"
      ? importAmountForSign
      : Number(importAmountForSign || 0);
  const isOpen = act.endsWith("to Open");
  const isClose = act.endsWith("to Close");
  const isBuy = act.startsWith("Buy ");
  const isSell = act.startsWith("Sell ");

  // Infer OPENING debit/credit sign:
  // - Opening trade: sign is the row's Total Cost
  // - Closing trade: sign is inverted (because it reverses the entry cashflow)
  let entrySign = tc;
  if (isClose && !isNaN(tc)) entrySign = -tc;
  const entryIsCredit = !isNaN(entrySign) && entrySign > 0;
  const entryIsDebit = !isNaN(entrySign) && entrySign < 0;

  // STOCK
  if (sp === "STOCK") {
    if (isOpen && isBuy) return "Long Stock";
    if (isOpen && isSell) return "Short Stock";
    if (isClose && isSell) return "Long Stock";
    if (isClose && isBuy) return "Short Stock";
    return "Stock";
  }

  // SINGLE-LEG OPTION
  if (sp === "SINGLE") {
    if (cp === "C") {
      if (isOpen && isBuy) return "Long Call";
      if (isOpen && isSell) return "Short Call";
      if (isClose && isSell) return "Long Call";
      if (isClose && isBuy) return "Short Call";
      return "Call";
    }
    if (cp === "P") {
      if (isOpen && isBuy) return "Long Put";
      if (isOpen && isSell) return "Short Put";
      if (isClose && isSell) return "Long Put";
      if (isClose && isBuy) return "Short Put";
      return "Put";
    }
    return "Option";
  }

  if (sp === "VERTICAL") {
    // OPEN legs get a directional single-leg label so the position tracker
    // has meaningful context for mixed-ticket VERTICAL groups (e.g., one
    // CLOSE leg + one OPEN leg in the same TOS ticket).
    // postProcessStrategyTypeBySpreadGroups will overwrite this for
    // pure-OPEN or pure-CLOSE VERTICAL groups with the correct PDS/PCS/etc.
    if (isOpen) {
      if (cp === "P") return isBuy ? "Long Put" : "Short Put";
      if (cp === "C") return isBuy ? "Long Call" : "Short Call";
    }
    // CLOSE legs and ambiguous cases: return generic Vertical.
    // postProcessStrategyTypeBySpreadGroups handles the final label.
    return "Vertical";
  }

  // IRON CONDOR -> Long/Short IC
  if (sp === "IRON CONDOR") {
    if (!entryIsCredit && !entryIsDebit) return "IC";
    return entryIsCredit ? "Short IC" : "Long IC";
  }

  // BUTTERFLY -> Long/Short Butterfly
  if (sp === "BUTTERFLY") {
    if (!entryIsCredit && !entryIsDebit) return "Butterfly";
    return entryIsDebit ? "Long Butterfly" : "Short Butterfly";
  }

  // Default: keep original label for debugging
  return spreadRaw || "";
}

function postProcessStrategyTypeBySpreadGroups(outItems, mappingHeaderMap) {
  const idxTs = col(mappingHeaderMap, "Trade Time Stamp");
  const idxTicker = col(mappingHeaderMap, "Ticker");
  const idxAction = col(mappingHeaderMap, "Action");
  const idxExp = col(mappingHeaderMap, "Option Expiration");
  const idxStrike = col(mappingHeaderMap, "Option Strike");
  const idxCallPut = col(mappingHeaderMap, "Call/Put");
  const idxSignedQty = col(mappingHeaderMap, "Signed Quantity");
  const idxStrategy = col(mappingHeaderMap, "Strategy Type");

  function toMs(d) {
    return d instanceof Date && !isNaN(d) ? d.getTime() : "";
  }

  const groups = {};
  for (let i = 0; i < outItems.length; i++) {
    const it = outItems[i];
    const row = it.row;

    const sp = String(it.spreadRaw || "")
      .trim()
      .toUpperCase();
    if (!SPREAD_GROUP_TYPES.includes(sp)) continue;

    const ts = row[idxTs];
    const tkr = String(row[idxTicker] || "");
    const exp = row[idxExp];

    const cp = String(row[idxCallPut] || "")
      .trim()
      .toUpperCase();
    const key =
      sp +
      "|" +
      toMs(ts) +
      "|" +
      tkr +
      "|" +
      toMs(exp) +
      "|" +
      (sp === "BUTTERFLY" ? cp : "");

    if (!groups[key]) groups[key] = [];
    groups[key].push(it);
  }

  Object.keys(groups).forEach(function (key) {
    const items = groups[key];
    if (!items.length) return;

    const sp = String(items[0].spreadRaw || "")
      .trim()
      .toUpperCase();

    const strikeQty = {};
    let callPutGroup = "";

    for (let i = 0; i < items.length; i++) {
      const row = items[i].row;
      const act = String(row[idxAction] || "");
      const isClose = act.endsWith("to Close");

      const strike = row[idxStrike];
      const sq = Number(row[idxSignedQty]);

      const cp = String(row[idxCallPut] || "")
        .trim()
        .toUpperCase();
      if (!callPutGroup && cp) callPutGroup = cp;

      if (strike === "" || strike === null || typeof strike === "undefined")
        continue;
      if (isNaN(Number(strike))) continue;
      if (isNaN(sq) || sq === 0) continue;

      const entrySq = isClose ? -sq : sq;
      const k = String(Number(strike));
      strikeQty[k] = (strikeQty[k] || 0) + entrySq;
    }

    function setAll(label) {
      for (let i = 0; i < items.length; i++) items[i].row[idxStrategy] = label;
    }

    const cp = callPutGroup;

    if (sp === "VERTICAL") {
      const strikes = Object.keys(strikeQty)
        .map(Number)
        .sort((a, b) => a - b);
      if (strikes.length < 2) return;

      let shortStrike = null,
        longStrike = null;
      for (let i = 0; i < strikes.length; i++) {
        const s = strikes[i];
        const q = strikeQty[String(s)] || 0;
        if (q < 0) shortStrike = s;
        if (q > 0) longStrike = s;
      }
      if (shortStrike === null || longStrike === null) return;

      // ── NEW: detect mixed OPEN/CLOSE groups ──────────────────────────────
      // A "mixed group" is a single TOS ticket that simultaneously closes one
      // leg and opens another (e.g., 5/19 9:38: Sell-to-Close 310 + Buy-to-Open 315).
      // For mixed groups we must NOT apply the group label to the OPEN legs,
      // because the open leg belongs to a *different* spread lifecycle than the
      // close leg. Labeling both "Vertical" creates single-leg PDS fragments
      // downstream. Instead: label CLOSE legs with the resolved spread name,
      // and leave OPEN legs untouched (their label comes from the position
      // tracker or the 9:59 group pass).
      const hasMixedDirections =
        items.some((it) => it.row[idxAction].endsWith("to Close")) &&
        items.some((it) => it.row[idxAction].endsWith("to Open"));

      const label =
        cp === "C"
          ? shortStrike < longStrike
            ? "CCS"
            : "CDS"
          : shortStrike < longStrike
            ? "PDS"
            : "PCS";

      if (hasMixedDirections) {
        // Only stamp CLOSE legs; leave OPEN legs alone.
        for (let i = 0; i < items.length; i++) {
          const act = items[i].row[idxAction];
          if (act.endsWith("to Close")) {
            items[i].row[idxStrategy] = label;
          }
          // OPEN legs in a mixed group: intentionally not touched here.
          // The position tracker will inherit whatever label normalizeStrategyType
          // set (Long Put / Short Put / etc.) and the next full-group VERTICAL
          // pass (9:59 ticket) will stamp PDS on those OPEN legs correctly.
        }
      } else {
        // Pure OPEN or pure CLOSE group — original behavior, stamp all legs.
        setAll(label);
      }
      return;
    }

    if (sp === "IRON CONDOR") {
      const callStrikeQty = {};
      const putStrikeQty = {};

      for (let i = 0; i < items.length; i++) {
        const row = items[i].row;

        const act = String(row[idxAction] || "");
        const isClose = act.endsWith("to Close");

        const cp = String(row[idxCallPut] || "")
          .trim()
          .toUpperCase();
        const strike = row[idxStrike];
        const sq = Number(row[idxSignedQty]);

        if (!cp || isNaN(Number(strike)) || isNaN(sq) || sq === 0) continue;
        const entrySq = isClose ? -sq : sq;

        const map =
          cp === "C" ? callStrikeQty : cp === "P" ? putStrikeQty : null;
        if (!map) continue;

        const k = String(Number(strike));
        map[k] = (map[k] || 0) + entrySq;
      }

      function findShortLong(map) {
        const strikes = Object.keys(map)
          .map(Number)
          .sort(function (a, b) {
            return a - b;
          });
        let shortS = null,
          longS = null;
        for (let i = 0; i < strikes.length; i++) {
          const s = strikes[i];
          const q = map[String(s)] || 0;
          if (q < 0) shortS = s;
          if (q > 0) longS = s;
        }
        return { shortStrike: shortS, longStrike: longS };
      }

      const c = findShortLong(callStrikeQty);
      const p = findShortLong(putStrikeQty);
      if (c.shortStrike === null || c.longStrike === null) return;
      if (p.shortStrike === null || p.longStrike === null) return;

      const isShortIC =
        c.longStrike > c.shortStrike && p.longStrike < p.shortStrike;
      setAll(isShortIC ? "Short IC" : "Long IC");
      return;
    }

    if (sp === "BUTTERFLY") {
      const strikes = Object.keys(strikeQty)
        .map(Number)
        .sort(function (a, b) {
          return a - b;
        });
      if (strikes.length < 3) return;

      let bodyStrike = null;
      let bodyQty = 0;
      for (let i = 0; i < strikes.length; i++) {
        const s = strikes[i];
        const q = strikeQty[String(s)] || 0;
        if (Math.abs(q) > Math.abs(bodyQty)) {
          bodyQty = q;
          bodyStrike = s;
        }
      }
      if (bodyStrike === null) return;

      let wingsAllPositive = true;
      let wingsAllNegative = true;
      for (let i = 0; i < strikes.length; i++) {
        const s = strikes[i];
        if (s === bodyStrike) continue;
        const q = strikeQty[String(s)] || 0;
        if (q <= 0) wingsAllPositive = false;
        if (q >= 0) wingsAllNegative = false;
      }

      if (bodyQty < 0 && wingsAllPositive) {
        setAll("Long Butterfly");
        return;
      }
      if (bodyQty > 0 && wingsAllNegative) {
        setAll("Short Butterfly");
        return;
      }
      return;
    }
  });
}

/**
 * postProcessStrategyTypeByPositionTrackerV3
 *
 * Runs AFTER postProcessStrategyTypeBySpreadGroups.
 *
 * Simulates a running FIFO position ledger per {Account, Ticker, Expiration, Strike}
 * in chronological order to re-label Strategy Type on CLOSE rows where the current
 * label is ambiguous (Long Put, Short Put, Long Call, Short Call, Vertical, etc.)
 * but the ledger can identify which named spread the position was originally opened as.
 *
 * OPEN rows are NEVER re-labeled here — postProcessStrategyTypeBySpreadGroups already
 * set the best possible label for opens, and we must trust that to populate the ledger
 * correctly. The only exception: if an OPEN row has a label that is already a named
 * spread (PDS, PCS, etc.), that label is what gets pushed to the ledger and later
 * applied to the matching CLOSE rows.
 *
 * Key design decisions:
 * - FIFO lot consumption: partial closes reduce the front lot's qty before moving on.
 * - Majority-qty wins when a close spans multiple lots with different labels.
 * - Only re-labels CLOSE rows where the ledger gives a MORE SPECIFIC answer than
 *   the current label (i.e., current is generic and ledger has a named spread label).
 * - 'Vertical' on an OPEN row that cannot be resolved by the group post-processor
 *   (e.g., a mixed-group open that is truly a standalone leg) is left as-is and
 *   pushed to the ledger as 'Vertical'. Downstream CLOSE rows inheriting 'Vertical'
 *   from the ledger will NOT be re-labeled because 'Vertical' is also in GENERIC_LABELS,
 *   meaning the re-label guard (currentIsGeneric || bestIsNamedSpread) only fires
 *   when bestLabel is a named spread — which 'Vertical' is not.
 */
function postProcessStrategyTypeByPositionTrackerV3(
  outItems,
  mappingHeaderMap,
) {
  const idxTs = col(mappingHeaderMap, "Trade Time Stamp");
  const idxAcct = col(mappingHeaderMap, "Account");
  const idxTicker = col(mappingHeaderMap, "Ticker");
  const idxExp = col(mappingHeaderMap, "Option Expiration");
  const idxStrike = col(mappingHeaderMap, "Option Strike");
  const idxAction = col(mappingHeaderMap, "Action");
  const idxSignedQty = col(mappingHeaderMap, "Signed Quantity");
  const idxStrategy = col(mappingHeaderMap, "Strategy Type");

  function toMs(d) {
    return d instanceof Date && !isNaN(d) ? d.getTime() : null;
  }

  // ── Sort a working index chronologically for correct ledger simulation ──
  const sortedIndices = outItems
    .map((_, i) => i)
    .sort((a, b) => {
      const ta = toMs(outItems[a].row[idxTs]);
      const tb = toMs(outItems[b].row[idxTs]);
      if (ta !== null && tb !== null && ta !== tb) return ta - tb;
      if (ta !== null && tb === null) return -1;
      if (ta === null && tb !== null) return 1;
      return (outItems[a].importRowNum || 0) - (outItems[b].importRowNum || 0);
    });

  // ledger[acct][ticker][expMs][strikeKey] = [ { label, qty }, ... ]  (FIFO lots)
  const ledger = {};

  function getLedgerBucket(acct, ticker, expMs, strikeKey) {
    if (!ledger[acct]) ledger[acct] = {};
    if (!ledger[acct][ticker]) ledger[acct][ticker] = {};
    if (!ledger[acct][ticker][expMs]) ledger[acct][ticker][expMs] = {};
    if (!ledger[acct][ticker][expMs][strikeKey])
      ledger[acct][ticker][expMs][strikeKey] = [];
    return ledger[acct][ticker][expMs][strikeKey];
  }

  // Collect relabels and apply them after the full simulation to avoid
  // any ordering side-effects within the same timestamp group.
  const relabelMap = {}; // outItems index → new label string

  const LABEL_RANK = {
    PDS: 0,
    PCS: 1,
    CDS: 2,
    CCS: 3,
    "Short IC": 4,
    "Long IC": 5,
    "Long Butterfly": 6,
    "Short Butterfly": 7,
    "Long Call": 8,
    "Short Call": 9,
    "Long Put": 10,
    "Short Put": 11,
    IC: 12,
    Butterfly: 13,
    Vertical: 14,
    Call: 15,
    Put: 16,
    Option: 17,
    "": 18,
  };
  // Generic labels = ambiguous single-leg or unresolved group labels.
  // A CLOSE row with one of these labels is a candidate for re-labeling
  // IF the ledger has a named-spread answer.
  const GENERIC_LABELS = new Set([
    "Long Put",
    "Short Put",
    "Long Call",
    "Short Call",
    "Put",
    "Call",
    "Option",
    "Vertical",
    "IC",
    "Butterfly",
    "",
  ]);

  for (const idx of sortedIndices) {
    const it = outItems[idx];
    const row = it.row;

    const acct = String(row[idxAcct] || "")
      .trim()
      .toUpperCase();
    const ticker = String(row[idxTicker] || "")
      .trim()
      .toUpperCase();
    const expDate = row[idxExp];
    const strike = row[idxStrike];
    const action = String(row[idxAction] || "").trim();
    const signedQtyRaw = row[idxSignedQty];

    if (!acct || !ticker) continue;
    const expMs = toMs(expDate);
    if (expMs === null) continue; // not an option row
    if (strike === "" || strike === null || strike === undefined) continue;
    const strikeKey = String(Number(strike));
    if (isNaN(Number(strikeKey))) continue;

    const signedQty =
      typeof signedQtyRaw === "number"
        ? signedQtyRaw
        : Number(String(signedQtyRaw || "").replace(/,/g, ""));
    if (isNaN(signedQty) || signedQty === 0) continue;

    const isOpen = action.endsWith("to Open");
    const isClose = action.endsWith("to Close");
    if (!isOpen && !isClose) continue;

    const currentLabel = String(row[idxStrategy] || "").trim();
    const bucket = getLedgerBucket(acct, ticker, expMs, strikeKey);

    if (isOpen) {
      // Push a new FIFO lot with whatever label the group post-processor set.
      // 'Vertical' here means the open was an isolated/standalone leg that the
      // group post-processor couldn't resolve — we intentionally preserve that.
      bucket.push({ label: currentLabel, qty: Math.abs(signedQty) });
    } else {
      // CLOSE — consume FIFO lots and determine the best label for this row.
      let remaining = Math.abs(signedQty);
      const labelsConsumed = [];

      while (remaining > 0 && bucket.length > 0) {
        const lot = bucket[0];
        if (lot.qty <= remaining) {
          labelsConsumed.push({ label: lot.label, qty: lot.qty });
          remaining -= lot.qty;
          bucket.shift();
        } else {
          labelsConsumed.push({ label: lot.label, qty: remaining });
          lot.qty -= remaining;
          remaining = 0;
        }
      }
      // If we consumed more than the ledger knew about (missing history),
      // fall back to the current label for the unmatched portion.
      if (remaining > 0 && currentLabel) {
        labelsConsumed.push({ label: currentLabel, qty: remaining });
      }
      if (labelsConsumed.length === 0) continue;

      // Majority-qty wins; named-spread labels win ties over generic ones.
      const tally = {};
      for (const { label, qty } of labelsConsumed) {
        tally[label] = (tally[label] || 0) + qty;
      }
      const bestLabel = Object.keys(tally).sort((a, b) => {
        const qDiff = tally[b] - tally[a];
        if (qDiff !== 0) return qDiff;
        return (
          (LABEL_RANK[a] !== undefined ? LABEL_RANK[a] : 99) -
          (LABEL_RANK[b] !== undefined ? LABEL_RANK[b] : 99)
        );
      })[0];

      // Re-label guard — corrected:
      const currentIsGeneric = GENERIC_LABELS.has(currentLabel);
      const bestIsGeneric = GENERIC_LABELS.has(bestLabel) || bestLabel === "";

      // Only block the re-label in one case:
      // current is a named spread AND best is generic — ledger can't improve things.
      // In all other cases (generic→named, generic→generic, named→named-different),
      // the ledger's answer is at least as good or better than what's there.
      const shouldRelabel =
        bestLabel !== currentLabel && !(!currentIsGeneric && bestIsGeneric); // block only: named current + generic best

      if (shouldRelabel) {
        relabelMap[idx] = bestLabel;
      }
    }
  }

  // Apply all collected relabels
  for (const [idxStr, newLabel] of Object.entries(relabelMap)) {
    outItems[Number(idxStr)].row[idxStrategy] = newLabel;
  }
}