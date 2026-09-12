/**
 * MapNetAmount.js
 *
 * Phase 2 – Group H: Net Amount post-processors
 *
 * Sole responsibility:
 *   Copy a single Schwab Net Amount across all legs of the same spread
 *   group, and WARN when every leg is blank.
 *
 * Functions (same names as before — do not rename):
 *   - postProcessNetAmountBySpreadGroups
 *   - postProcessWarnMissingNetAmountBySpreadGroupsV3
 *
 * Called by:
 *   mapSchwabImportByHeadersV3()
 *
 * Uses (already global — do not redeclare):
 *   SPREAD_GROUP_TYPES
 *   col                 (Helpers.js)
 *   mappingIssuesAdd    (ImportIssues.js)
 *
 * Do not redeclare SHEET_* consts or SPREAD_GROUP_TYPES here.
 */

/**
 * WARN when a whole spread group has missing Net Amount on all legs.
 * Uses Schwab Mapping Issues instead of the old Error Log array.
 */
function postProcessWarnMissingNetAmountBySpreadGroupsV3(
  outItems,
  mappingHeaderMap,
  ctx,
) {
  const idxTs = col(mappingHeaderMap, "Trade Time Stamp");
  const idxTicker = col(mappingHeaderMap, "Ticker");
  const idxDesc = col(mappingHeaderMap, "Description");
  const idxExp = col(mappingHeaderMap, "Option Expiration");
  const idxNetAmount = col(mappingHeaderMap, "Net Amount");

  function toMs(d) {
    return d instanceof Date && !isNaN(d) ? d.getTime() : "";
  }

  // Group by Spread + TimeStamp + Ticker + Expiration
  const groups = {};
  for (let i = 0; i < outItems.length; i++) {
    const it = outItems[i];
    const sp = String(it.spreadRaw || "")
      .trim()
      .toUpperCase();
    if (!SPREAD_GROUP_TYPES.includes(sp)) continue;

    const row = it.row;
    const key =
      sp +
      "|" +
      toMs(row[idxTs]) +
      "|" +
      String(row[idxTicker] || "") +
      "|" +
      toMs(row[idxExp]);

    if (!groups[key]) groups[key] = [];
    groups[key].push(it);
  }

  Object.keys(groups).forEach(function (key) {
    const items = groups[key];
    if (!items.length) return;

    // If ANY leg has a Net Amount, we consider it OK.
    let hasAnyNet = false;
    for (let i = 0; i < items.length; i++) {
      const v = items[i].row[idxNetAmount];
      if (typeof v === "number" && !isNaN(v) && v !== 0) {
        hasAnyNet = true;
        break;
      }
      if (String(v || "").trim() !== "") {
        hasAnyNet = true;
        break;
      }
    }
    if (hasAnyNet) return;

    // Log one WARN using first leg as representative
    const first = items[0];
    const row = first.row;
    const spread = String(first.spreadRaw || "")
      .trim()
      .toUpperCase();

    mappingIssuesAdd(
      ctx,
      "WARN",
      first.importRowNum,
      "Net Amount",
      "",
      "Missing Net Amount on all legs for spread group (" +
        spread +
        "). Ticker=" +
        String(row[idxTicker] || "") +
        ", Exp=" +
        String(row[idxExp] || "") +
        ", Desc=" +
        String(row[idxDesc] || ""),
    );
  });
}

function postProcessNetAmountBySpreadGroups(outItems, mappingHeaderMap) {
  const idxTs = col(mappingHeaderMap, "Trade Time Stamp");
  const idxTicker = col(mappingHeaderMap, "Ticker");
  const idxExp = col(mappingHeaderMap, "Option Expiration");
  const idxNetAmount = col(mappingHeaderMap, "Net Amount");

  function toMs(d) {
    return d instanceof Date && !isNaN(d) ? d.getTime() : "";
  }

  const groups = {};
  for (let i = 0; i < outItems.length; i++) {
    const it = outItems[i];
    const sp = String(it.spreadRaw || "")
      .trim()
      .toUpperCase();
    if (!SPREAD_GROUP_TYPES.includes(sp)) continue;

    const row = it.row;
    const key =
      sp +
      "|" +
      toMs(row[idxTs]) +
      "|" +
      String(row[idxTicker] || "") +
      "|" +
      toMs(row[idxExp]);

    if (!groups[key]) groups[key] = [];
    groups[key].push(it);
  }

  Object.keys(groups).forEach(function (key) {
    const items = groups[key];
    let net = "";

    for (let i = 0; i < items.length; i++) {
      const v = items[i].row[idxNetAmount];
      if (typeof v === "number" && !isNaN(v) && v !== 0) {
        net = v;
        break;
      }
      if (String(v || "").trim() !== "") {
        net = v;
        break;
      }
    }

    if (net === "" || net === null || typeof net === "undefined") return;

    for (let i = 0; i < items.length; i++) items[i].row[idxNetAmount] = net;
  });
}