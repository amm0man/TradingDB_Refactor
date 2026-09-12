/**
 * MapCashFlow.js
 *
 * Phase 2 – Group F: Cash Map
 *
 * Sole responsibility:
 *   Load sheet "Cash Map" and stamp Cash Flow Direction + Transfer Type
 *   on non-trade Mapping rows.
 *
 * Functions (same names as before — do not rename):
 *   - buildCashFlowMapFromSheet
 *   - applyCashFlowFromMapV3
 *
 * Called by:
 *   mapSchwabImportByHeadersV3()
 *
 * Uses (already global — do not redeclare):
 *   SHEET_CASH_MAP
 *   buildHeaderIndexMap, requireHeaders, col   (Helpers.js)
 *   mappingIssuesAdd                           (ImportIssues.js)
 *
 * Do not redeclare SHEET_* consts here.
 */

/** Build cashFlowMap from "Cash Map" sheet. Keyed by Account Actions. */
function buildCashFlowMapFromSheet(ss) {
  const sh = ss.getSheetByName(SHEET_CASH_MAP);
  if (!sh) throw new Error("Could not find required sheet: " + SHEET_CASH_MAP);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return {};

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const hm = buildHeaderIndexMap(headers);

  // Expect: Account Actions | CashFlowDir | Transfer Type
  requireHeaders(
    hm,
    ["Account Actions", "CashFlowDir", "Transfer Type"],
    SHEET_CASH_MAP,
  );

  const out = {};
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const action = String(r[col(hm, "Account Actions")] || "").trim();
    if (!action) continue;

    const dir = String(r[col(hm, "CashFlowDir")] || "").trim();
    const type = String(r[col(hm, "Transfer Type")] || "").trim();

    out[action] = { dir: dir, type: type };
  }

  return out;
}

/**
 * Apply Cash Flow Direction + Transfer Type:
 * - Direction: derived from Total Cost sign (your rule: direction from Amount sign)
 * - Transfer Type: from Cash Map if available, else fallback to Account Actions tag
 *
 * NOTE: called only on NON-TRADE rows in this script.
 */
function applyCashFlowFromMapV3(
  mappedRow,
  mappingHeaderMap,
  cashFlowMap,
  ctx,
  importRowNum,
) {
  const accountActionTag = String(
    mappedRow[col(mappingHeaderMap, "Account Actions")] || "",
  ).trim();
  if (!accountActionTag) return;

  // Transfer Type: prefer Cash Map
  const meta = cashFlowMap[accountActionTag] || null;
  const type = meta && meta.type ? meta.type : accountActionTag;

  // Direction: sign of Total Cost (which for non-trade rows = import Amount)
  const totalCostRaw = mappedRow[col(mappingHeaderMap, "Total Cost")];
  const n =
    typeof totalCostRaw === "number"
      ? totalCostRaw
      : Number(
          String(totalCostRaw || "")
            .replace(/[$,]/g, "")
            .trim(),
        );

  let dir = "";
  if (!isNaN(n) && n !== 0) dir = n > 0 ? "Inflow" : "Outflow";

  // Fallback: if we cannot compute sign, use Cash Map direction ONLY as a last resort.
  if (!dir && meta && meta.dir) dir = meta.dir;

  if (!dir) {
    mappingIssuesAdd(
      ctx,
      "WARN",
      importRowNum,
      "Cash Flow Direction",
      "",
      "Could not derive direction (Total Cost blank/0/non-numeric). Account Actions=" +
        accountActionTag,
    );
  }

  mappedRow[col(mappingHeaderMap, "Cash Flow Direction")] = dir;
  mappedRow[col(mappingHeaderMap, "Transfer Type")] = type;
}