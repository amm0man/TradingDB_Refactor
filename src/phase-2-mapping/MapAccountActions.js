/**
 * MapAccountActions.js
 *
 * Phase 2 – Group A: Account Actions tagging
 *
 * Sole responsibility:
 *   Keyword rules + directional journal logic that fill the
 *   "Account Actions" column on Schwab Mapping.
 *
 * Functions (same names as before — do not rename):
 *   - getAccountActionsKeywordRulesV3
 *   - deriveAccountActionTagV3
 *   - deriveJournalTransferDirectionTag
 *
 * Called by:
 *   mapSchwabImportByHeadersV3() in the main mapping loop
 *
 * Still lives in mapSchwabImportByHeadersV3.js (Group E later):
 *   findFirstKeywordTag  — used by deriveAccountActionTagV3 fallback
 *
 * Do not redeclare SHEET_* consts here. They are already top-level in
 * mapSchwabImportByHeadersV3.js and are shared across Apps Script files.
 */

/**
 * Base keyword rules for Account Actions where the tag is NOT dependent on
 * account direction.
 *
 * Directional transfers are handled in deriveAccountActionTagV3().
 */
function getAccountActionsKeywordRulesV3() {
  return [
    { keyword: "subscription fee", tag: "Subscription Fee" },
    { keyword: "foreign security fee", tag: "Foreign Security Fee" },

    { keyword: "removal of option due to expiration", tag: "Opt Expired" },
    { keyword: "expired", tag: "Option Expired" },

    { keyword: "non-qualified div", tag: "Non-Qualified Dividend" },

    {
      keyword: "transfer of security or option in",
      tag: "Transfer of Security or Option In",
    },
    {
      keyword: "transfer of security or option out",
      tag: "Transfer of Security or Option Out",
    },

    {
      keyword: "removal of option due to exercise",
      tag: "Option Removal - Exercised",
    },
    {
      keyword: "removal of option due to assignment",
      tag: "Option Removal - Assignment",
    },

    { keyword: "reorganization fee", tag: "Reorganization Fee" },
    

    { keyword: "incoming account transfer", tag: "Incoming Account Transfer" },
    { keyword: "outgoing account transfer", tag: "Outgoing Account Transfer" },

    // Funding / ACH-ish patterns (direction comes from Amount sign)
    { keyword: "new account funding", tag: "Initial Account Funding" },
    {
      keyword: "electronic new account funding",
      tag: "Initial Account Funding",
    },
    { keyword: "client requested electronic", tag: "ACH In or Out" },
    { keyword: "malvern nation", tag: "ACH In or Out" },
    { keyword: "electronic funding", tag: "ACH In or Out" },

    { keyword: "miscellaneous journal entry", tag: "Miscellaneous Journal Entry" },
    { keyword: "cash alternatives", tag: "Cash Alternatives Interest" },
    { keyword: "schwab1 int", tag: "Credit Interest" },

    { keyword: "account migration from tda", tag: "Account Migration" },
  ];
}

/**
 * Security transfer receipts (shares or options moving in/out of the account).
 *
 * WHY: "Transfer of Security or Option In 3.0 AAPL" is tagged as an Account
 * Action. That tag used to veto isTrade, so Mapping left Quantity blank and
 * Action = RAD. Phase 3 then did RAD && ticker with qty 0 (no-op), and a
 * later sell went negative (LT AAPL 2022-12-29 sold 6, running -3).
 *
 * This helper does NOT change the Account Actions label. It only says
 * "this tagged row is also a $0 stock lot that Phase 3 must book."
 *
 * Description shape we have seen:
 *   TRANSFER OF SECURITY OR OPTION IN 3.0 AAPL
 *   transfer of security or option out 2.0 MSFT
 *
 * Returns null when the tag is not a security transfer, or qty cannot be read.
 */
function parseSecurityTransferReceiptV3(accountActionTag, descRaw) {
  const tag = String(accountActionTag || "").trim();
  const isIn = tag === "Transfer of Security or Option In";
  const isOut = tag === "Transfer of Security or Option Out";
  if (!isIn && !isOut) return null;

  const desc = String(descRaw || "");
  const m = desc.match(
    /transfer of security or option\s+(in|out)\s+(-?\d+(?:\.\d+)?)(?:\s+([A-Za-z][A-Za-z0-9./]{0,9}))?/i,
  );
  if (!m) return null;

  const qty = Math.abs(parseFloat(m[2]));
  if (!isFinite(qty) || qty <= 0) return null;

  let tickerHint = String(m[3] || "")
    .replace(/\.[A-Za-z]+$/, "")
    .toUpperCase();
  if (tickerHint === "CALL" || tickerHint === "PUT") tickerHint = "";

  return {
    direction: isIn ? "IN" : "OUT",
    side: isIn ? "BUY" : "SELL",
    posEffect: isIn ? "TO OPEN" : "TO CLOSE",
    qty: qty,
    ticker: tickerHint,
  };
}

/**
 * Derive Account Actions tag:
 * - Directional journal transfers based on "...750" / "...937" and frm/to phrasing
 * - Internal transfers ("internal transfer of cash", "third party") based on Account + Amount sign
 * - Otherwise fall back to simple keyword rules
 */
function deriveAccountActionTagV3(
  descLower,
  accountRaw,
  importAmount,
  accountActionRules,
) {
  const d = String(descLower || "")
    .trim()
    .toLowerCase();
  const account = String(accountRaw || "")
    .trim()
    .toUpperCase();

  // 1) Directional journal transfers (best signal if present)
  const journalTag = deriveJournalTransferDirectionTag(d);
  if (journalTag) return journalTag;

  // 2) Internal transfers that don't always include account numbers in description
  if (d.includes("internal transfer of cash") || d.includes("third party")) {
    // If we have a usable signed amount, we can infer direction based on:
    // - Positive amount = cash INTO this row’s account
    // - Negative amount = cash OUT of this row’s account
    const n =
      typeof importAmount === "number" ? importAmount : Number(importAmount);

    if ((account === "DT" || account === "LT") && !isNaN(n) && n !== 0) {
      const other = account === "DT" ? "LT" : "DT";

      // Example:
      // - Account=DT, Amount=+2000 => Transfer from LT to DT Account
      // - Account=DT, Amount=-2000 => Transfer from DT to LT Account
      if (n > 0)
        return "Transfer from " + other + " to " + account + " Account";
      if (n < 0)
        return "Transfer from " + account + " to " + other + " Account";
    }

    // If amount is missing/0, we can’t safely infer direction.
    // Return a neutral tag (you can filter these easily later).
    return "Internal Transfer (Direction Unknown)";
  }

  // 3) Fall back to ordered keyword rules
  return findFirstKeywordTag(d, accountActionRules);
}

/**
 * Infer direction from "JOURNAL FRM ..." / "JOURNAL TO ..." plus account endings:
 * - DT ends with 750
 * - LT ends with 937
 *
 * We want BOTH sides of the transfer to share the SAME label, e.g.:
 * "Transfer from DT to LT Account"
 */
function deriveJournalTransferDirectionTag(descLower) {
  const d = String(descLower || "").toLowerCase();

  // Normalize some common variants
  const hasFrm = d.includes("journal frm") || d.includes("journal from");
  const hasTo = d.includes("journal to");

  const has750 = d.includes("750");
  const has937 = d.includes("937");

  // If it's a journal and we see BOTH account endings, pick direction based on frm/to.
  // - "frm 750" means FROM DT -> TO LT
  // - "to 937" means TO LT -> FROM DT
  if (hasFrm && has750) return "Transfer from DT to LT Account";
  if (hasFrm && has937) return "Transfer from LT to DT Account";
  if (hasTo && has937) return "Transfer from DT to LT Account";
  if (hasTo && has750) return "Transfer from LT to DT Account";

  return "";
}