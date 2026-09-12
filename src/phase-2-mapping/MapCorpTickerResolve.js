/**
 * MapCorpTickerResolve.js
 *
 * Phase 2 – Group D: CUSIP / symbol-change / Corp Action Map ticker resolution
 *
 * Sole responsibility:
 *   Resolve the canonical Ticker for TDA-era corp-action rows and for
 *   "Symbol Change from X to Y" journals.
 *
 * Functions (same names as before — do not rename):
 *   - buildCusipMapFromSheetV3
 *   - normalizeRenameIdentifierV3
 *   - looksLikeCusipIdentifierV3
 *   - resolveRenameIdentifierToTickerV3
 *   - parseSymbolChangeDescriptionV3
 *   - buildCorpActionMapFromSheet
 *   - extractTickerFromTildePattern
 *   - lookupTickerFromCorpActionMap
 *   - extractTickerFromCorpActionDesc
 *
 * Called by:
 *   mapSchwabImportByHeadersV3()
 *
 * Shared from Helpers.js (already global):
 *   normalizeCusip, looksLikeCusip
 *
 * Sheet tabs this group reads (string names, not SHEET_* consts):
 *   "CusipMap"         — CUSIP → ticker
 *   "Corp Action Map"  — company-name pattern or exact-ticker alias → ticker
 *
 * Do not mix those with Phase 1 "CorpActionsMap" / "CorpActionStockMap".
 * Do not redeclare SHEET_* consts here.
 */

// Reads CusipMap into a normalized lookup object.
// Accepts headers: CUSIP in col A style, and Symbol/Ticker/Underlying as the mapped value.
function buildCusipMapFromSheetV3(ss) {
  const sh = ss.getSheetByName("CusipMap");
  if (!sh) return {};

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return {};

  const vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = vals[0].map((h) =>
    String(h || "")
      .trim()
      .toLowerCase(),
  );

  const cusipIdx = headers.findIndex((h) => h === "cusip");
  const symIdx = headers.findIndex(
    (h) => h === "symbol" || h === "ticker" || h === "underlying",
  );
  if (cusipIdx === -1 || symIdx === -1) return {};

  const out = {};
  for (let r = 1; r < vals.length; r++) {
    // Force string first so Sheets number-coercion cannot drop leading zeros / letters.
    const rawCusip = normalizeCusip(String(vals[r][cusipIdx] ?? "").trim());
    const rawSym = String(vals[r][symIdx] || "")
      .trim()
      .toUpperCase();
    if (!rawCusip || !rawSym) continue;
    if (!looksLikeCusip(rawCusip)) continue;
    out[rawCusip] = rawSym;
  }
  return out;
}

/**
 * Normalize a rename-side identifier (ticker or CUSIP-like token).
 * Preserves characters such as "/" that appear in real tickers (e.g. LUR/CN).
 * For pure CUSIP key work, use normalizeCusip() from CusipHelpers.js instead.
 */
function normalizeRenameIdentifierV3(v) {
  return String(v || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");
}

/** Thin wrapper – prefer the shared looksLikeCusip() going forward. */
function looksLikeCusipIdentifierV3(v) {
  return looksLikeCusip(v);
}

function resolveRenameIdentifierToTickerV3(rawId, cusipMap) {
  const id = normalizeRenameIdentifierV3(rawId);
  if (!id) return "";

  // If it looks like a CUSIP, try CusipMap first (shared normalizer).
  const cusipKey = normalizeCusip(id);
  if (looksLikeCusip(cusipKey) && cusipMap && cusipMap[cusipKey]) {
    return String(cusipMap[cusipKey]).trim().toUpperCase();
  }

  // Otherwise treat it as a ticker-like identifier and normalize for storage.
  return id;
}

function parseSymbolChangeDescriptionV3(desc, cusipMap) {
  const text = String(desc || "").trim();
  if (!text) return null;

  // Examples:
  // Symbol Change from ENCUF to EU
  // Symbol Change from 50545P309 to LUR/CN
  // Symbol Change from UNG to 912318300
  const m = text.match(/SYMBOL\s+CHANGE\s+FROM\s+(.+?)\s+TO\s+(.+)$/i);
  if (!m) return null;

  const fromRaw = String(m[1] || "")
    .trim()
    .toUpperCase();
  const toRaw = String(m[2] || "")
    .trim()
    .toUpperCase();

  const fromResolved = resolveRenameIdentifierToTickerV3(fromRaw, cusipMap);
  const toResolved = resolveRenameIdentifierToTickerV3(toRaw, cusipMap);

  return {
    fromRaw: fromRaw,
    toRaw: toRaw,
    fromResolved: fromResolved,
    toResolved: toResolved,
    notesText: text,
  };
}

// Loads Corp Action Map sheet into a sorted array of { pattern, ticker } objects.
// Patterns are sorted longest-first so more-specific entries win over shorter ones.
// If the sheet doesn't exist, returns [] — the fallback chain degrades gracefully.
// Sheet columns: Company Name Pattern (col A) | Ticker (col B) | Notes (col C, ignored)
function buildCorpActionMapFromSheet(ss) {
  const sh = ss.getSheetByName("Corp Action Map");
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return []; // header row only, no data

  const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  const entries = [];
  for (const row of data) {
    const pattern = String(row[0] || "")
      .trim()
      .toUpperCase();
    const ticker = String(row[1] || "")
      .trim()
      .toUpperCase();
    if (pattern && ticker) {
      entries.push({ pattern, ticker });
    }
  }
  // Longest pattern first: "JPMORGAN EQUITY PREMIUM INCOME ETF" (34 chars) beats "JPMORGAN" (8 chars)
  // if someone ever adds a shorter prefix entry by mistake.
  entries.sort((a, b) => b.pattern.length - a.pattern.length);
  return entries;
}

// Extracts a ticker symbol from TDA-era tilde-delimited descriptions.
// Examples:
//   "ORDINARY DIVIDEND~JEPI"          → "JEPI"
//   "Ordinary Dividend~JEPI 3.43 US$" → "JEPI"
//   "Ordinary Dividend~BAC 2.10 US$"  → "BAC"
// The regex only matches ≤6 alpha chars after ~ followed by a non-alpha boundary
// (whitespace, digit, or end of string). This naturally rejects company names:
//   "~JPMORGAN EQUITY" → no match because "JPMORGAN" is followed by " " but
//   we'd match "JPMORGA" (6 chars) and then "N" is NOT a valid boundary → no match.
// Returns '' if no match.
// AFTER — reject placeholder tokens used by TDA when no description was available:
function extractTickerFromTildePattern(desc) {
  if (!desc) return "";
  const m = desc.match(/~([A-Za-z]{1,6})(?:\s|$|\d)/);
  if (!m) return "";
  const token = m[1].toUpperCase();
  // TDA used "~NO DESCRIPTION" as a placeholder meaning "not applicable."
  // "NO", "NA", "NONE", "TBD" after a tilde are always placeholders, never tickers.
  const TILDE_REJECTS = new Set(["NO", "NA", "NONE", "TBD", "NULL"]);
  if (TILDE_REJECTS.has(token)) return "";
  return token;
}

// Looks up a ticker from the Corp Action Map for company-name style descriptions.
// Uses case-insensitive substring match. Since entries are sorted longest-first,
// the most specific pattern wins automatically.
// Returns '' if no match (caller falls through to tokenizer or logs WARN).
function lookupTickerFromCorpActionMap(desc, corpActionMap) {
  if (!corpActionMap || !corpActionMap.length || !desc) return "";
  const upper = desc.toUpperCase();
  for (const entry of corpActionMap) {
    if (upper.includes(entry.pattern)) return entry.ticker;
  }
  return "";
}

// Extracts a likely ticker symbol from a Schwab corporate action Description string.
// Used as a fallback when the Symbol field is blank on RAD/DOI corporate event rows.
function extractTickerFromCorpActionDesc(desc) {
  if (!desc) return "";

  // Action-phrase words that appear in Schwab corp action descriptions but are NOT tickers.
  // Keep this list conservative — only add words you are certain are never used as tickers.
  const SKIP = new Set([
    "MANDATORY",
    "REVERSE",
    "SPLIT",
    "STOCK",
    "FORWARD",
    "MERGER",
    "EXCHANGE",
    "TRANSFER",
    "SECURITY",
    "OPTION",
    "SPIN",
    "OFF",
    "LIQUIDATION",
    "PENDING",
    "RECEIPT",
    "QUALIFIED",
    "DIVIDEND",
    "NON",
    "TAXABLE",
    "DIV",
    "INTEREST",
    "CASH",
    "ALTERNATIVES",
    "REORGANIZATION",
    "REORGANIZED",
    "ISSUE",
    "LIEU",
    "FRACTIONAL",
    "SHARES",
    "IN",
    "OF",
    "FOR",
    "INTO",
    "WITH",
    "AND",
    "OR",
    "THE",
    "DUE",
    "TO",
    "RATIO",
    "NEW",
    "OLD",
    "RECORD",
    "DATE",
    "UPON",
    "CORPORATE",
    "ACTION",
    "FREE",
    "BALANCE",
    "ADJUSTMENT",
    "MARGIN",
    "DESCRIPTION",
    "INTEREST",
    "ORDINARY",
    "QUALIFIED",
    "DISTRIBUTION",
    "PARTNERSHIP",
  ]);

  // Split on whitespace and common delimiters (colon, comma, parens, slash, period, dash).
  // The slash split handles "UHAL/B" → ["UHAL", "B"] so we get "UHAL" first.
  const tokens = desc.toUpperCase().split(/[\s,:.()\[\]\/\-]+/);

  for (const token of tokens) {
    if (!token) continue;
    // Accept only 1-6 purely alpha characters — the standard ticker format.
    // This deliberately rejects numeric tokens (ratio numbers like "10", "1"),
    // and dot/dollar-prefixed index symbols like "$SPX" (handled by Symbol-primary path).
    if (!/^[A-Z]{1,6}$/.test(token)) continue;
    if (SKIP.has(token)) continue;
    return token; // first non-skip purely alpha token is the ticker
  }

  return "";
}