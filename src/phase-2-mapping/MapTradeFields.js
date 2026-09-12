/**
 * MapTradeFields.js
 *
 * Phase 2 – Group B: trade field builders
 *
 * Sole responsibility:
 *   Decide whether an import row is a trade, and build Action,
 *   Signed Quantity, Ticker-from-Symbol, and Call/Put.
 *
 * Functions (same names as before — do not rename):
 *   - isTradeBySidePosEffect
 *   - buildTradeAction
 *   - buildSignedQuantity
 *   - extractTickerFromSymbol
 *   - extractCallPut
 *
 * Called by:
 *   mapSchwabImportByHeadersV3() in the main loop
 *
 * Do not redeclare SHEET_* consts here.
 */

/** Identify trade rows strictly by Side + Pos Effect. */
function isTradeBySidePosEffect(sideRaw, posEffectRaw) {
  const side = String(sideRaw || "")
    .trim()
    .toUpperCase();
  const pe = String(posEffectRaw || "")
    .trim()
    .toUpperCase();
  if (side !== "BUY" && side !== "SELL") return false;
  if (!pe.includes("OPEN") && !pe.includes("CLOSE")) return false;
  return true;
}

/** Build "Buy to Open" etc. */
function buildTradeAction(sideRaw, posEffectRaw) {
  const side = String(sideRaw || "")
    .trim()
    .toUpperCase();
  const pe = String(posEffectRaw || "")
    .trim()
    .toUpperCase();

  const sideNice = side === "BUY" ? "Buy" : side === "SELL" ? "Sell" : "";
  const peNice = pe.includes("OPEN")
    ? "Open"
    : pe.includes("CLOSE")
      ? "Close"
      : "";
  if (!sideNice || !peNice) return "";
  return sideNice + " to " + peNice;
}

/** Signed quantity: BUY positive, SELL negative. */
function buildSignedQuantity(sideRaw, qtyAbs) {
  const side = String(sideRaw || "")
    .trim()
    .toUpperCase();
  const q = Number(qtyAbs || 0);
  if (!q) return "";
  return side === "SELL" ? -Math.abs(q) : Math.abs(q);
}

/** Extract ticker from Symbol. For options like "MRVL 11/19/2021 70.00 C", ticker = "MRVL". */
function extractTickerFromSymbol(symbolRaw) {
  const s = String(symbolRaw || "").trim();
  if (!s) return "";
  return s.split(/\s+/)[0].trim();
}

/** Extract Call/Put from Symbol or Description. Returns "C" or "P" or "". */
function extractCallPut(symbolRaw, desc) {
  const s = String(symbolRaw || "").trim();
  const m = s.match(/\b([CP])\b\s*$/i);
  if (m) return m[1].toUpperCase();

  const d = String(desc || "").toUpperCase();
  if (d.includes(" CALL ") || d.includes(" CALL")) return "C";
  if (d.includes(" PUT ") || d.includes(" PUT")) return "P";
  return "";
}