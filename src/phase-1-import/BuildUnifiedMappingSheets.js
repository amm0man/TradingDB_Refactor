/**
 * BuildUnifiedMappingSheets.js
 *
 * Phase 1 – Group G: mapping-sheet loaders
 *
 * Sole responsibility:
 *   Read the optional helper sheets that buildUnifiedImportV3() uses to
 *   resolve CUSIPs, corporate-action stock emits, token/phrase rules,
 *   split adjustments, and fake-trade overrides.
 *
 * Sheets:
 *   - CusipMap
 *   - SplitAdjustments
 *   - CorpActionStockMap
 *   - CorpActionsMap
 *   - FakeDropOverride
 *
 * Why a factory (createMappingSheetHelpers):
 *   The loaders log through the current Import Issues ctx and look up
 *   sheets with getSheetByNameLoose / ss. Passing those in keeps logging
 *   and sheet lookup identical to the nested version.
 *
 * Called by:
 *   buildUnifiedImportV3() at the start of section 4 (after toStr,
 *   getSheetByNameLoose, and normalizeSymbol exist).
 *
 * Shared helpers used from Helpers.js (already global):
 *   normalizeHeader, normalizeCusip, looksLikeCusip, normalizeDate, toNum
 *
 * Still passed in from buildUnifiedImportV3:
 *   ctx, ss, toStr, getSheetByNameLoose, normalizeSymbol
 *
 * Related files:
 *   - BuildUnifiedImportV3.js      (orchestrator; still calls the loaders)
 *   - BuildUnifiedSymbols.js       (normalizeSymbol)
 *   - ImportIssues.js
 */

/**
 * createMappingSheetHelpers(opts)
 *
 * opts.ctx                  – importIssues context for this run
 * opts.ss                   – SpreadsheetApp.getActiveSpreadsheet()
 * opts.toStr                – nested string helper
 * opts.getSheetByNameLoose  – nested loose sheet finder
 * opts.normalizeSymbol      – Group E helper already bound in the orchestrator
 *
 * Returns the Group G function names that used to be nested in
 * buildUnifiedImportV3. Call-site argument lists are unchanged.
 */
function createMappingSheetHelpers(opts) {
  const ctx = opts.ctx;
  const ss = opts.ss;
  const toStr = opts.toStr;
  const getSheetByNameLoose = opts.getSheetByNameLoose;
  const normalizeSymbol = opts.normalizeSymbol;

  function readCusipMap() {
    const desiredName = "CusipMap";
    const sh = getSheetByNameLoose(desiredName);
    if (!sh) {
      importIssuesAdd(
        ctx,
        "CUSIPMAP_MISSING",
        "CusipMap",
        1,
        "Sheet",
        desiredName,
        "Could not find sheet by name (trimmed + case-insensitive).",
      );
      return {};
    }

    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) {
      importIssuesAdd(
        ctx,
        "CUSIPMAP_EMPTY",
        "CusipMap",
        1,
        "Rows",
        vals.length,
        "CusipMap has header only or no data.",
      );
      return {};
    }

    const headers = vals[0].map((h) => toStr(h).trim());
    const idx = {};
    for (let c = 0; c < headers.length; c++)
      idx[normalizeHeader(headers[c])] = c;

    function col(nameOrNames) {
      const names = Array.isArray(nameOrNames) ? nameOrNames : [nameOrNames];
      for (let i = 0; i < names.length; i++) {
        const k = normalizeHeader(names[i]);
        if (idx[k] !== undefined) return idx[k];
      }
      return null;
    }

    const iCusip = col("CUSIP");
    const iSym = col(["Symbol", "Ticker", "Underlying"]);
    if (iCusip === null || iSym === null) {
      throw new Error(
        "CusipMap sheet must have headers CUSIP and Symbol (or Ticker/Underlying).",
      );
    }

    const map = {};
    for (let r = 1; r < vals.length; r++) {
      // Force CUSIP cell to string BEFORE normalizeCusip so Sheets number-coercion
      // (which silently drops leading zeros and non-digit characters) cannot corrupt
      // alphanumeric CUSIPs like 00848K101 → 848101.
      const cusipRaw = toStr(vals[r][iCusip]).trim();
      const cusip = normalizeCusip(cusipRaw);
      const sym = toStr(vals[r][iSym]).trim().toUpperCase();
      if (!cusip || !sym) continue;
      if (!looksLikeCusip(cusip)) continue;
      map[cusip] = sym;
    }
    return map;
  }

  function readSplitAdjustments() {
    const sh = getSheetByNameLoose("SplitAdjustments");
    if (!sh) return [];
    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return [];
    const hdr = vals[0].map((h) =>
      String(h ?? "")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, ""),
    );
    const iT = hdr.indexOf("ticker");
    const iD = hdr.indexOf("splitdate");
    const iN = hdr.indexOf("rationumerator");
    const iDn = hdr.indexOf("ratiodenominator");
    const iST = hdr.indexOf("splittype");
    if (iT < 0 || iD < 0 || iN < 0 || iDn < 0 || iST < 0) {
      importIssuesAdd(
        ctx,
        "WARN",
        "SplitAdjustments",
        1,
        "Headers",
        hdr.join(", "),
        "SplitAdjustments sheet missing required column(s). Expected: Ticker|Split Date|Ratio Numerator|Ratio Denominator|Split Type",
      );
      return [];
    }
    const out = [];
    for (let r = 1; r < vals.length; r++) {
      const ticker = toStr(vals[r][iT]).trim().toUpperCase();
      const rawDate = vals[r][iD];
      const dateIso = normalizeDate(rawDate);
      const num = toNum(vals[r][iN]);
      const den = toNum(vals[r][iDn]);
      const stype = toStr(vals[r][iST]).trim().toUpperCase(); // 'FORWARD' or 'REVERSE'
      if (!ticker || !dateIso || isNaN(num) || isNaN(den) || den === 0)
        continue;
      out.push({ ticker, dateIso, num, den, stype });
    }
    importIssuesSetMetric(ctx, "SplitAdjustmentsLoaded", out.length);
    return out;
  }

  function findCorpActionStockMapRow_(
    entries,
    account,
    dateIso,
    phrase,
    matchSymbol,
  ) {
    const acctU = toStr(account).trim().toUpperCase();
    const dateU = toStr(dateIso).trim();
    const phraseU = toStr(phrase).trim().toUpperCase();
    const matchU = normalizeSymbol(toStr(matchSymbol).trim().toUpperCase());

    for (let i = 0; i < entries.length; i++) {
      const e = entries[i];
      if (!e) continue;

      const accountOk = !e.account || e.account === acctU;
      const dateOk = !e.dateIso || e.dateIso === dateU;
      const phraseOk = e.phrase === phraseU;
      const matchOk = e.matchSymbol === matchU;

      if (accountOk && dateOk && phraseOk && matchOk) return e;
    }

    return null;
  }

  function readCorpActionStockMap() {
    const sh = getSheetByNameLoose("CorpActionStockMap");
    if (!sh) return [];

    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return [];

    const hdr = vals[0].map((h) =>
      String(h ?? "")
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, ""),
    );

    function idxOfHeader(name) {
      return hdr.indexOf(
        String(name)
          .trim()
          .toLowerCase()
          .replace(/[^a-z0-9]+/g, ""),
      );
    }

    const iAccount = idxOfHeader("Account");
    const iDate = idxOfHeader("Effective Date");
    const iPhrase = idxOfHeader("Phrase");
    const iMatchSymbol = idxOfHeader("Match Symbol");
    const iSourceSymbol = idxOfHeader("Source Symbol");
    const iEmitSymbol = idxOfHeader("Emit Symbol");
    const iEmitAction = idxOfHeader("Emit Action");
    const iQtyMultiplier = idxOfHeader("Quantity Multiplier");
    const iNotes = idxOfHeader("Notes");

    if (
      iDate < 0 ||
      iPhrase < 0 ||
      iMatchSymbol < 0 ||
      iSourceSymbol < 0 ||
      iEmitSymbol < 0 ||
      iEmitAction < 0
    ) {
      importIssuesAdd(
        ctx,
        "WARN",
        "CorpActionStockMap",
        1,
        "Headers",
        hdr.join(", "),
        "CorpActionStockMap is missing required headers. Expected: Account | Effective Date | Phrase | Match Symbol | Source Symbol | Emit Symbol | Emit Action | Quantity Multiplier | Notes",
      );
      return [];
    }

    const out = [];
    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];

      const account =
        iAccount >= 0 ? toStr(row[iAccount]).trim().toUpperCase() : "";
      const dateIso = iDate >= 0 ? normalizeDate(row[iDate]) : "";
      const phrase =
        iPhrase >= 0 ? toStr(row[iPhrase]).trim().toUpperCase() : "";
      const matchSymbol =
        iMatchSymbol >= 0
          ? normalizeSymbol(toStr(row[iMatchSymbol]).trim().toUpperCase())
          : "";
      const sourceSymbol =
        iSourceSymbol >= 0
          ? normalizeSymbol(toStr(row[iSourceSymbol]).trim().toUpperCase())
          : "";
      const emitSymbol =
        iEmitSymbol >= 0
          ? normalizeSymbol(toStr(row[iEmitSymbol]).trim().toUpperCase())
          : "";
      const emitAction =
        iEmitAction >= 0 ? toStr(row[iEmitAction]).trim().toUpperCase() : "";
      const qtyMultiplierRaw = iQtyMultiplier >= 0 ? row[iQtyMultiplier] : "";
      const qtyMultiplier = isNaN(toNum(qtyMultiplierRaw))
        ? 1
        : toNum(qtyMultiplierRaw);
      const notes = iNotes >= 0 ? toStr(row[iNotes]).trim() : "";

      if (
        !dateIso ||
        !phrase ||
        !matchSymbol ||
        !sourceSymbol ||
        !emitSymbol ||
        !emitAction
      )
        continue;

      out.push({
        account: account,
        dateIso: dateIso,
        phrase: phrase,
        matchSymbol: matchSymbol,
        sourceSymbol: sourceSymbol,
        emitSymbol: emitSymbol,
        emitAction: emitAction,
        qtyMultiplier: qtyMultiplier,
        notes: notes,
      });
    }

    importIssuesSetMetric(ctx, "CorpActionStockMapLoaded", out.length);
    return out;
  }

  function parseCorpActionStockDescription_(descRaw, cusipMap) {
    const u = String(descRaw ?? "")
      .trim()
      .toUpperCase();
    if (!u) return null;

    let phrase = "";
    if (u.includes("MANDATORY - EXCHANGE")) {
      phrase = "MANDATORY - EXCHANGE";
    } else if (u.includes("NON-TAXABLE SPIN OFF/LIQUIDATION DISTRIBUTION")) {
      phrase = "NON-TAXABLE SPIN OFF/LIQUIDATION DISTRIBUTION";
    } else {
      return null;
    }

    const afterPhrase = u.substring(u.indexOf(phrase) + phrase.length).trim();
    const qtyMatch = afterPhrase.match(/([-+]?\d+(?:\.\d+)?)/);
    if (!qtyMatch) return null;

    const rawQty = Number(qtyMatch[1]);
    const parsedQty = Math.abs(rawQty);
    if (!isFinite(parsedQty) || parsedQty <= 0) return null;

    const afterQty = afterPhrase
      .substring(qtyMatch.index + qtyMatch[0].length)
      .trim();
    const tokens = afterQty.match(/[A-Z0-9\/]{1,20}/g) || [];

    let rawToken = "";
    for (let t = 0; t < tokens.length; t++) {
      const tok = String(tokens[t] || "")
        .trim()
        .toUpperCase();
      if (!tok) continue;
      if (/^\d+(?:\.\d+)?$/.test(tok)) continue;
      rawToken = tok;
      break;
    }

    if (!rawToken) return null;

    let resolvedSymbol = normalizeSymbol(rawToken);
    const cusipCandidate = normalizeCusip(rawToken);
    if (looksLikeCusip(cusipCandidate) && /\d/.test(cusipCandidate)) {
      const mapped = cusipMap[cusipCandidate];
      if (mapped) {
        resolvedSymbol = normalizeSymbol(mapped);
      } else {
        resolvedSymbol = cusipCandidate;
      }
    }

    return {
      phrase: phrase,
      rawQty: rawQty,
      parsedQty: parsedQty,
      rawToken: rawToken,
      resolvedSymbol: resolvedSymbol,
      rawText: u,
    };
  }

  function readCorpActionsMap() {
    const sh = ss.getSheetByName("CorpActionsMap");
    if (!sh) return { tokenRules: [], phraseRules: [] };

    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return { tokenRules: [], phraseRules: [] };

    const hdr = vals[0].map((h) => String(h ?? "").trim());
    const idxToken = hdr.indexOf("Token");
    const idxTokenAction = hdr.indexOf("TokenAction");
    const idxPhrase = hdr.indexOf("Phrase");
    const idxPhraseAction = hdr.indexOf("PhraseAction");
    const idxResultSymbol = hdr.indexOf("ResultSymbol"); // optional

    if (idxToken < 0 && idxPhrase < 0)
      throw new Error(
        "CorpActionsMap must have Token and/or Phrase columns.",
      );
    if (idxToken >= 0 && idxTokenAction < 0)
      throw new Error("CorpActionsMap missing TokenAction header.");
    if (idxPhrase >= 0 && idxPhraseAction < 0)
      throw new Error("CorpActionsMap missing PhraseAction header.");

    const tokenRules = [];
    const phraseRules = [];

    for (let r = 1; r < vals.length; r++) {
      const token =
        idxToken >= 0
          ? String(vals[r][idxToken] ?? "")
              .trim()
              .toUpperCase()
          : "";
      const tokenAction =
        idxTokenAction >= 0
          ? String(vals[r][idxTokenAction] ?? "")
              .trim()
              .toUpperCase()
          : "";
      const phrase =
        idxPhrase >= 0
          ? String(vals[r][idxPhrase] ?? "")
              .trim()
              .toUpperCase()
          : "";
      const phraseAction =
        idxPhraseAction >= 0
          ? String(vals[r][idxPhraseAction] ?? "")
              .trim()
              .toUpperCase()
          : "";
      const resultSymbol =
        idxResultSymbol >= 0
          ? String(vals[r][idxResultSymbol] ?? "")
              .trim()
              .toUpperCase()
          : "";

      if (token && tokenAction)
        tokenRules.push({ token, action: tokenAction, resultSymbol });
      if (phrase && phraseAction)
        phraseRules.push({ phrase, action: phraseAction, resultSymbol });
    }

    phraseRules.sort((a, b) => b.phrase.length - a.phrase.length);

    return { tokenRules, phraseRules };
  }

  function applyCorpActionsToTopRow(descRaw, corpMap) {
    const desc = String(descRaw ?? "");
    const descU = desc.toUpperCase();

    for (let i = 0; i < corpMap.phraseRules.length; i++) {
      const r = corpMap.phraseRules[i];
      if (!r.phrase) continue;
      if (!descU.includes(r.phrase)) continue;

      if (r.action === "IGNORE") return { handled: true, resultSymbol: null };
      if (r.action === "SETSYMBOL")
        return { handled: true, resultSymbol: r.resultSymbol || null };
      return { handled: true, resultSymbol: null };
    }

    return { handled: false, resultSymbol: null };
  }

  function applyCorpTokenRule(tokenRaw, corpMap) {
    const tok = String(tokenRaw ?? "")
      .trim()
      .toUpperCase();
    if (!tok) return { handled: false, resultSymbol: null };

    for (let i = 0; i < corpMap.tokenRules.length; i++) {
      const r = corpMap.tokenRules[i];
      if (r.token !== tok) continue;

      if (r.action === "IGNORE") return { handled: true, resultSymbol: null };
      if (r.action === "SETSYMBOL")
        return { handled: true, resultSymbol: r.resultSymbol || null };
      return { handled: true, resultSymbol: null };
    }

    return { handled: false, resultSymbol: null };
  }

  function readFakeDropOverride() {
    const sh = getSheetByNameLoose("FakeDropOverride");
    if (!sh) return {}; // Sheet is optional — if absent, no overrides apply.
    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return {};
    const headers = vals[0].map((h) => toStr(h).trim().toUpperCase());
    // Accept a column named "Ticker", "CUSIP", or "Symbol"
    let col = headers.indexOf("TICKER");
    if (col < 0) col = headers.indexOf("CUSIP");
    if (col < 0) col = headers.indexOf("SYMBOL");
    if (col < 0) {
      importIssuesAdd(
        ctx,
        "WARN",
        "FakeDropOverride",
        1,
        "Header",
        headers.join(", "),
        'FakeDropOverride sheet found but missing a "Ticker", "CUSIP", or "Symbol" column. No overrides applied.',
      );
      return {};
    }
    const set = {};
    for (let r = 1; r < vals.length; r++) {
      const v = toStr(vals[r][col]).trim().toUpperCase();
      if (v) set[v] = true;
    }
    importIssuesSetMetric(
      ctx,
      "FakeDropOverrideCount",
      Object.keys(set).length,
    );
    return set;
  }

  return {
    readCusipMap: readCusipMap,
    readSplitAdjustments: readSplitAdjustments,
    findCorpActionStockMapRow_: findCorpActionStockMapRow_,
    readCorpActionStockMap: readCorpActionStockMap,
    parseCorpActionStockDescription_: parseCorpActionStockDescription_,
    readCorpActionsMap: readCorpActionsMap,
    applyCorpActionsToTopRow: applyCorpActionsToTopRow,
    applyCorpTokenRule: applyCorpTokenRule,
    readFakeDropOverride: readFakeDropOverride,
  };
}
