/**
 * @file debugSchwabHeaderMapping.js
 * @description Standalone debugger originally located in MapSchwabImportByHeadersV3.js.
 *              Used for troubleshooting corporate action / stock emitter mapping behavior
 *              during Schwab header-based imports (Phase 2).
 *
 * This file was extracted on June 23, 2026 as part of moving standalone debug utilities
 * into the src/debug/ folder. It is intentionally kept separate from core production code.
 *
 * Location of original function before extraction:
 *   src/phase-2-mapping/MapSchwabImportByHeadersV3.js → debugCorpActionStockEmitterV3()
 *
 * Usage:
 * - Run this file independently when you need to debug corporate action or emitter
 *   mapping issues in Schwab imports.
 * - It can safely be excluded from any future clean handoff of the main project.
 */

function debugCorpActionStockEmitterV3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function toStr(v) {
    return v == null ? '' : String(v);
  }

  function normalizeHeader(s) {
    return toStr(s).trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
  }

  function getSheetByNameLoose(name) {
    const target = String(name).trim().toLowerCase();
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const sh = sheets[i];
      if (String(sh.getName()).trim().toLowerCase() === target) return sh;
    }
    return null;
  }

  function mustGetSheet(name) {
    const sh = getSheetByNameLoose(name);
    if (!sh) throw new Error('Missing sheet: ' + name);
    return sh;
  }

  function readSheetObjects(sheet) {
    const vals = sheet.getDataRange().getValues();
    if (vals.length < 2) return [];
    const hdr = vals[0].map(h => toStr(h).trim());
    const out = [];
    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      const isBlank = row.every(v => toStr(v).trim() === '');
      if (isBlank) continue;
      const obj = {};
      for (let c = 0; c < hdr.length; c++) obj[hdr[c]] = row[c];
      obj.__rowNum = r + 1;
      out.push(obj);
    }
    return out;
  }

  function getField(obj, nameOrNames) {
    if (!obj) return '';
    const names = Array.isArray(nameOrNames) ? nameOrNames : [nameOrNames];
    const keys = Object.keys(obj);

    function nk(x) {
      return String(x == null ? '' : x).trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
    }

    for (let i = 0; i < names.length; i++) {
      const target = nk(names[i]);
      for (let k = 0; k < keys.length; k++) {
        if (nk(keys[k]) === target) return obj[keys[k]];
      }
    }
    return '';
  }

  function toNum(v) {
    if (v == null || v === '') return NaN;
    if (typeof v === 'number') return v;
    const s = String(v).replace(/,/g, '').trim();
    const n = Number(s);
    return isNaN(n) ? NaN : n;
  }

  function normalizeDate(v) {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    const s = toStr(v).trim();
    if (!s) return '';
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);

    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (m) {
      const mm = String(parseInt(m[1], 10)).padStart(2, '0');
      const dd = String(parseInt(m[2], 10)).padStart(2, '0');
      let yyyy = parseInt(m[3], 10);
      if (yyyy < 100) yyyy += 2000;
      return yyyy + '-' + mm + '-' + dd;
    }

    const digits = s.replace(/\D/g, '');
    if (digits.length === 8) {
      const mm = digits.substring(0, 2);
      const dd = digits.substring(2, 4);
      const yyyy = digits.substring(4, 8);
      return yyyy + '-' + mm + '-' + dd;
    }

    return s;
  }

  function normalizeTimeHHmmss(v) {
    if (v instanceof Date && !isNaN(v.getTime())) {
      return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HHmmss');
    }
    const s = toStr(v).trim();
    if (!s) return '';
    const digits = s.replace(/\D/g, '');
    if (!digits) return '';
    if (digits.length === 4) return digits + '00';
    return digits.padStart(6, '0').substring(0, 6);
  }

  function normalizeCusip(v) {
    return toStr(v).toUpperCase().replace(/[^0-9A-Z]/g, '');
  }

  function looksLikeCusip(v) {
    const s = normalizeCusip(v);
    return /^[0-9A-Z]{9}$/.test(s);
  }

  function normalizeSymbol(s) {
    return toStr(s).trim().toUpperCase().replace(/\s+/g, '');
  }

  function readCusipMap() {
    const sh = mustGetSheet('CusipMap');
    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return {};

    const headers = vals[0].map(h => toStr(h).trim());
    const idx = {};
    for (let c = 0; c < headers.length; c++) idx[normalizeHeader(headers[c])] = c;

    function col(nameOrNames) {
      const names = Array.isArray(nameOrNames) ? nameOrNames : [nameOrNames];
      for (let i = 0; i < names.length; i++) {
        const k = normalizeHeader(names[i]);
        if (idx[k] != null) return idx[k];
      }
      return -1;
    }

    const iCusip = col('CUSIP');
    const iSym = col(['Symbol', 'Ticker', 'Underlying']);
    if (iCusip < 0 || iSym < 0) throw new Error('CusipMap missing CUSIP or Symbol/Ticker/Underlying header.');

    const map = {};
    for (let r = 1; r < vals.length; r++) {
      const cusip = normalizeCusip(vals[r][iCusip]);
      const sym = normalizeSymbol(vals[r][iSym]);
      if (!cusip || !sym) continue;
      map[cusip] = sym;
    }
    return map;
  }

  function readCorpActionStockMap() {
    const sh = mustGetSheet('CorpActionStockMap');
    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return [];

    const hdr = vals[0].map(h => normalizeHeader(h));

    function idxOfHeader(name) {
      return hdr.indexOf(normalizeHeader(name));
    }

    const iAccount = idxOfHeader('Account');
    const iDate = idxOfHeader('Effective Date');
    const iPhrase = idxOfHeader('Phrase');
    const iMatchSymbol = idxOfHeader('Match Symbol');
    const iSourceSymbol = idxOfHeader('Source Symbol');
    const iEmitSymbol = idxOfHeader('Emit Symbol');
    const iEmitAction = idxOfHeader('Emit Action');
    const iQtyMultiplier = idxOfHeader('Quantity Multiplier');
    const iNotes = idxOfHeader('Notes');

    const out = [];
    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      const account = iAccount >= 0 ? toStr(row[iAccount]).trim().toUpperCase() : '';
      const dateIso = iDate >= 0 ? normalizeDate(row[iDate]) : '';
      const phrase = iPhrase >= 0 ? toStr(row[iPhrase]).trim().toUpperCase() : '';
      const matchSymbol = iMatchSymbol >= 0 ? normalizeSymbol(row[iMatchSymbol]) : '';
      const sourceSymbol = iSourceSymbol >= 0 ? normalizeSymbol(row[iSourceSymbol]) : '';
      const emitSymbol = iEmitSymbol >= 0 ? normalizeSymbol(row[iEmitSymbol]) : '';
      const emitAction = iEmitAction >= 0 ? toStr(row[iEmitAction]).trim().toUpperCase() : '';
      const qtyMultiplier = iQtyMultiplier >= 0 && !isNaN(toNum(row[iQtyMultiplier])) ? toNum(row[iQtyMultiplier]) : 1;
      const notes = iNotes >= 0 ? toStr(row[iNotes]).trim() : '';

      if (!dateIso || !phrase || !matchSymbol || !sourceSymbol || !emitSymbol || !emitAction) continue;

      out.push({
        account: account,
        dateIso: dateIso,
        phrase: phrase,
        matchSymbol: matchSymbol,
        sourceSymbol: sourceSymbol,
        emitSymbol: emitSymbol,
        emitAction: emitAction,
        qtyMultiplier: qtyMultiplier,
        notes: notes
      });
    }
    return out;
  }

   function parseCorpActionStockDescription_(descRaw, cusipMap) {
    const u = String(descRaw ?? "").trim().toUpperCase();
    if (!u) return null;

    // Must stay in lockstep with parseCorpActionStockDescription_
    // inside createMappingSheetHelpers (BuildUnifiedMappingSheets.js).
    let phrase = "";
    if (u.includes("MANDATORY - EXCHANGE")) {
      phrase = "MANDATORY - EXCHANGE";
    } else if (u.includes("NON-TAXABLE SPIN OFF/LIQUIDATION DISTRIBUTION")) {
      phrase = "NON-TAXABLE SPIN OFF/LIQUIDATION DISTRIBUTION";
    } else if (u.includes("STOCK MERGER")) {
      phrase = "STOCK MERGER";
    } else {
      // Abbreviated credit receipt, e.g. "F4 URANIUM CORP F 336.0 FFUCF"
      // Phrase is the map key. It does not appear in the TOS text.
      const abbr = u.match(
        /\b([A-Z])\s+([-+]?\d+(?:\.\d+)?)\s+([A-Z]{1,6})\s*$/,
      );
      if (!abbr) return null;
      const rawQtyAbbr = Number(abbr[2]);
      const parsedQtyAbbr = Math.abs(rawQtyAbbr);
      if (!isFinite(parsedQtyAbbr) || parsedQtyAbbr <= 0) return null;
      const rawTokenAbbr = String(abbr[3] || "").toUpperCase();
      if (!rawTokenAbbr) return null;
      return {
        phrase: "CORP ACTION RECEIVED SHARES",
        rawQty: rawQtyAbbr,
        parsedQty: parsedQtyAbbr,
        rawToken: rawTokenAbbr,
        resolvedSymbol: normalizeSymbol(rawTokenAbbr),
        tokenWasCusip: false,
        cusipMapped: false,
        rawText: u,
      };
    }

    const afterPhrase = u.substring(u.indexOf(phrase) + phrase.length).trim();
    const qtyMatch = afterPhrase.match(/([-+]?\d+(?:\.\d+)?)/);
    if (!qtyMatch) return null;

    const rawQty = Number(qtyMatch[1]);
    const parsedQty = Math.abs(rawQty);
    if (!isFinite(parsedQty) || parsedQty <= 0) return null;

    // Debit side of "Stock Merger -qty" is not a delivery.
    if (phrase === "STOCK MERGER" && rawQty <= 0) return null;

    const afterQty = afterPhrase
      .substring(qtyMatch.index + qtyMatch[0].length)
      .trim();
    const tokens = afterQty.match(/[A-Z0-9\/]{1,20}/g) || [];

    let rawToken = "";
    for (let t = 0; t < tokens.length; t++) {
      const tok = String(tokens[t] || "").trim().toUpperCase();
      if (!tok) continue;
      if (/^\d+(?:\.\d+)?$/.test(tok)) continue;
      rawToken = tok;
      break;
    }
    if (!rawToken) return null;

    let resolvedSymbol = normalizeSymbol(rawToken);
    const cusipCandidate = normalizeCusip(rawToken);
    let tokenWasCusip = false;
    let cusipMapped = false;
    if (looksLikeCusip(cusipCandidate) && /\d/.test(cusipCandidate)) {
      tokenWasCusip = true;
      const mapped = cusipMap[cusipCandidate];
      if (mapped) {
        resolvedSymbol = normalizeSymbol(mapped);
        cusipMapped = true;
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
      tokenWasCusip: tokenWasCusip,
      cusipMapped: cusipMapped,
      rawText: u,
    };
  }

  
    function findCorpActionStockMapRow_(entries, account, dateIso, phrase, matchSymbol) {
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
  
  const topRows = readSheetObjects(mustGetSheet('TosTop'));
  const cusipMap = readCusipMap();
  const corpActionStockMap = readCorpActionStockMap();

  const debugRows = [];
  debugRows.push([
    'Source Row',
    'Account',
    'TYPE',
    'Date',
    'Time',
    'Description',
    'Detected?',
    'Parsed Phrase',
    'Parsed Qty',
    'Raw Token',
    'Resolved Symbol',
    'Token Was CUSIP?',
    'CUSIP Mapped?',
    'Map Match?',
    'Map Account',
    'Map Date',
    'Map Phrase',
    'Map Match Symbol',
    'Map Source Symbol',
    'Map Emit Symbol',
    'Map Emit Action',
    'Qty Multiplier',
    'Would Emit Qty',
    'Map Notes'
  ]);

  for (let i = 0; i < topRows.length; i++) {
    const r = topRows[i];

    const account = toStr(getField(r, 'Account')).trim().toUpperCase();
    const type = toStr(getField(r, 'TYPE')).trim().toUpperCase();
    const dateIso = normalizeDate(getField(r, 'DATE'));
    const timeHHmmss = normalizeTimeHHmmss(getField(r, 'TIME')) || '000000';
    const desc = toStr(getField(r, 'DESCRIPTION')).trim();

    if (type === 'TRD') continue;

       const descU = desc.toUpperCase();
    const quickHit =
      descU.includes("MANDATORY - EXCHANGE") ||
      descU.includes("NON-TAXABLE SPIN OFF/LIQUIDATION DISTRIBUTION") ||
      descU.includes("STOCK MERGER") ||
      descU.includes("PENDING RECEIPT") ||
      descU.indexOf("FFUCF") >= 0 ||
      descU.indexOf("FUUFF") >= 0;

    const parsed = parseCorpActionStockDescription_(desc, cusipMap);

    if (!quickHit && !parsed) continue;

    const matchRow = parsed
      ? findCorpActionStockMapRow_(corpActionStockMap, account, dateIso, parsed.phrase, parsed.resolvedSymbol)
      : null;

    const wouldEmitQty = (parsed && matchRow)
      ? Math.round((parsed.parsedQty * Number(matchRow.qtyMultiplier || 1)) * 1e8) / 1e8
      : '';

    debugRows.push([
      r.__rowNum || (i + 2),
      account,
      type,
      dateIso,
      timeHHmmss,
      desc,
      parsed ? 'Y' : 'N',
      parsed ? parsed.phrase : '',
      parsed ? parsed.parsedQty : '',
      parsed ? parsed.rawToken : '',
      parsed ? parsed.resolvedSymbol : '',
      parsed ? (parsed.tokenWasCusip ? 'Y' : 'N') : '',
      parsed ? (parsed.cusipMapped ? 'Y' : 'N') : '',
      matchRow ? 'Y' : 'N',
      matchRow ? matchRow.account : '',
      matchRow ? matchRow.dateIso : '',
      matchRow ? matchRow.phrase : '',
      matchRow ? matchRow.matchSymbol : '',
      matchRow ? matchRow.sourceSymbol : '',
      matchRow ? matchRow.emitSymbol : '',
      matchRow ? matchRow.emitAction : '',
      matchRow ? matchRow.qtyMultiplier : '',
      wouldEmitQty,
      matchRow ? matchRow.notes : ''
    ]);
  }

  let outSh = getSheetByNameLoose('CorpActionStockDebug');
  if (!outSh) outSh = ss.insertSheet('CorpActionStockDebug');
  outSh.clearContents();
  outSh.getRange(1, 1, debugRows.length, debugRows[0].length).setValues(debugRows);
  outSh.autoResizeColumns(1, debugRows[0].length);
}