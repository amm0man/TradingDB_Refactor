/**
 * MapCorpActionRules.js
 *
 * Phase 2 – Group E: Corporate Actions keyword rules + shared keyword scanner
 *
 * Sole responsibility:
 *   - Ordered keyword/regex rules that fill the "Corporate Actions" column
 *   - Shared first-match scanner used by both Account Actions and Corp Actions
 *
 * Functions (same names as before — do not rename):
 *   - findFirstKeywordTag
 *   - getCorpActionsKeywordRulesV3
 *
 * Called by:
 *   - mapSchwabImportByHeadersV3() (corp tag in the main loop)
 *   - deriveAccountActionTagV3() in MapAccountActions.js (fallback)
 *
 * Do not redeclare SHEET_* consts here.
 *
 * ORDER MATTERS in getCorpActionsKeywordRulesV3: more-specific patterns
 * must stay before less-specific ones. Do not alphabetize the list.
 */

// AFTER — supports both { keyword, tag } and { re, tag } rules
// Keyword rules use .includes() against descLower (already lowercased).
// Regex rules use .test() against the original or lowercased string as needed.
function findFirstKeywordTag(descLower, rules) {
  if (!descLower) return "";
  for (let i = 0; i < (rules || []).length; i++) {
    const rule = rules[i];
    if (rule.re) {
      // Regex rule: test against descLower (regex has /i flag so casing doesn't matter)
      if (rule.re.test(descLower)) return String(rule.tag || "");
    } else {
      const kw = String(rule.keyword || "").toLowerCase();
      if (!kw) continue;
      if (descLower.includes(kw)) return String(rule.tag || "");
    }
  }
  return "";
}

// =====================================================
// Corporate Actions helpers — CONSOLIDATED single source of truth (Phase 2)
// Phase 3 (copyMappingToImportByHeaders) trusts whatever this sets.
// No re-evaluation happens in Phase 3 — this is the only place that tags corp actions.
//
// ORDER MATTERS: more-specific patterns MUST come before less-specific ones.
// Rules support both { keyword, tag } (substring match) and { re, tag } (regex match).
// =====================================================
function getCorpActionsKeywordRulesV3() {
  return [
    // ── Dividends / Interest ──────────────────────────────────────────────────
    { keyword: "qualified dividend", tag: "Dividend" },
    { keyword: "ordinary dividend", tag: "Dividend" }, // TDA "ORDINARY DIVIDEND~JEPI"
    { keyword: "non-qualified dividend", tag: "Dividend" }, // long form variant
    { keyword: "special dividend", tag: "Dividend" }, // one-time special divs
    { keyword: "return of capital", tag: "Return of Capital" }, // REITs / MLPs
    { keyword: "monthly dividend", tag: "Dividend" }, // some TDA formats
    { keyword: "cash dividend", tag: "Cash Dividend" },
    { keyword: "reinvest dividend", tag: "DRIP" },
    { keyword: "reinvest shares", tag: "DRIP" },
    // +++ NEW: matches DRIP rows emitted by buildUnifiedImportV3
    // Description format: "DRIP BUY +0.0175 XOM UPON REINVESTMENT"
    { keyword: "upon reinvestment", tag: "DRIP" },
    // +++ END NEW
    { keyword: "foreign tax paid", tag: "Foreign Tax Paid" },
    { keyword: "foreign tax withheld", tag: "Foreign Tax Withheld" },
    { keyword: "bond interest", tag: "Bond Interest" },
    { keyword: "cash alternatives interest", tag: "Cash Interest" },
    { keyword: "partnership distribution", tag: "Partnership Distribution" },
    { keyword: "free balance interest", tag: "Interest Adjustment" },
    {
      keyword: "margin interest adjustment",
      tag: "Margin Interest Adjustment",
    },

    // ── Splits — more-specific FIRST ─────────────────────────────────────────
    { keyword: "mandatory reverse split", tag: "Mandatory Reverse Split" },
    { keyword: "reverse split", tag: "Reverse Split" },
    { keyword: "forward split with stock split", tag: "Forward Split" },
    { keyword: "stock split", tag: "Stock Split" },
    { keyword: "split", tag: "Stock Split" }, // catch-all — after all specific split variants

    // ── Mergers / Reorganizations ─────────────────────────────────────────────
    { keyword: "mandatory merger", tag: "Mandatory Merger" },
    { keyword: "stock merger", tag: "Stock Merger" },
    { keyword: "merger", tag: "Stock Merger" },
    { keyword: "reorganized issue", tag: "Reorganization" },
    { keyword: "mandatory exchange", tag: "Mandatory Exchange" },

    // ── Spin-offs / Liquidations ──────────────────────────────────────────────
    { keyword: "non-taxable spin off", tag: "Spin-off/Liquidation" },

    // ── Transfers ────────────────────────────────────────────────────────────
    // More-specific in/out variants before the generic one
    {
      keyword: "transfer of security or option in",
      tag: "Transfer of Security or Option In",
    },
    {
      keyword: "transfer of security or option out",
      tag: "Transfer of Security or Option Out",
    },
    { keyword: "transfer of security or option", tag: "Transfer of Security" },

    // ── Cash / Miscellaneous ─────────────────────────────────────────────────
    { keyword: "cash in lieu of fractional shares", tag: "Cash In Lieu" },

    // ── Pending / Received shares ─────────────────────────────────────────────
    { keyword: "pending receipt of new s", tag: "Pending Corp Action" },

    // ── Broad catch-all: paired credit-side rows for corporate restructurings ──
    // Schwab emits debit/credit pairs. The debit side is caught by 'pending receipt' above.
    // The credit side uses an abbreviated format: COMPANYNAME SINGLEACTIONLETTER QTY NEWTICKER
    // e.g.  "ACME F4 100 XYZ"
    // Placed LAST — only fires after every specific pattern has been tested.
    {
      re: /[A-Z]{4,}\s+[A-Z]\d+\s+[\d.]+\s+[A-Z]{1,6}/i,
      tag: "Corp Action Received Shares",
    },
  ];
}