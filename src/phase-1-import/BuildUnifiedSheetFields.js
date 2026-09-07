/**
 * BuildUnifiedSheetFields.js
 *
 * Phase 1 – Group C: sheet / field utilities
 *
 * Sole responsibility:
 *   Read a sheet as an array of header-keyed objects, look up a field by
 *   loose header name, and compare/round prices.
 *
 * Why a factory (createSheetFieldHelpers):
 *   readSheetObjects() still uses the nested toStr() from
 *   buildUnifiedImportV3(). Passing it in keeps that behavior identical.
 *   roundTo() and pricesClose() only need the shared toNum() from Helpers.js.
 *
 * Called by:
 *   buildUnifiedImportV3() as soon as toStr exists (before any sheet is read
 *   as objects).
 *
 * Shared helpers used from Helpers.js (already global):
 *   toNum
 *
 * Still passed in from buildUnifiedImportV3 (not moved yet):
 *   toStr
 *
 * Note:
 *   getField() does its own loose header match (trim + lowercase + collapse
 *   spaces). That is NOT the same as normalizeHeader() in Helpers.js. Do not
 *   merge them in this pass.
 *
 * Related files:
 *   - BuildUnifiedImportV3.js      (orchestrator)
 *   - BuildUnifiedEnrichment.js    (Group A)
 *   - BuildUnifiedIcRetag.js       (Group B; receives roundTo from here)
 */

/**
 * createSheetFieldHelpers(opts)
 *
 * opts.toStr – nested string helper from buildUnifiedImportV3
 *
 * Returns the Group C function names that used to be nested in
 * buildUnifiedImportV3. Call-site argument lists are unchanged.
 */
function createSheetFieldHelpers(opts) {
  const toStr = opts.toStr;

  /**
   * Reads a sheet into array of objects keyed by the *exact* header text.
   * buildUnifiedImportV3 uses object access + getField for TosTop rows.
   */
  function readSheetObjects(sheet) {
    const vals = sheet.getDataRange().getValues();
    if (vals.length < 2) return [];

    const hdr = vals[0].map((h) => toStr(h).trim());
    const out = [];

    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      const isBlank = row.every((v) => toStr(v).trim() === "");
      if (isBlank) continue;

      const obj = {};
      for (let c = 0; c < hdr.length; c++) {
        obj[hdr[c]] = row[c];
      }
      out.push(obj);
    }

    return out;
  }

  function getField(obj, nameOrNames) {
    if (!obj) return "";
    const names = Array.isArray(nameOrNames) ? nameOrNames : [nameOrNames];
    const keys = Object.keys(obj);

    function normKey(x) {
      return String(x ?? "")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, " ");
    }

    for (let i = 0; i < names.length; i++) {
      const target = normKey(names[i]);
      for (let k = 0; k < keys.length; k++) {
        if (normKey(keys[k]) === target) return obj[keys[k]];
      }
    }
    return "";
  }

  function roundTo(n, decimals) {
    const x = toNum(n);
    if (isNaN(x)) return NaN;
    const p = Math.pow(10, decimals || 0);
    return Math.round(x * p) / p;
  }

  function pricesClose(a, b) {
    const x = toNum(a);
    const y = toNum(b);
    if (isNaN(x) || isNaN(y)) return false;
    if (roundTo(x, 5) === roundTo(y, 5)) return true;
    if (Math.abs(x - y) < 0.0005) return true;
    return roundTo(x, 2) === roundTo(y, 2);
  }

  return {
    readSheetObjects: readSheetObjects,
    getField: getField,
    roundTo: roundTo,
    pricesClose: pricesClose,
  };
}
