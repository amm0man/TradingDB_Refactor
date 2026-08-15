/**
 * TosSheetWriteHelpers.js
 *
 * Small helper functions for safely writing data to sheets while preserving
 * important formatting — especially preventing Google Sheets from stripping
 * leading zeros on CUSIPs, account numbers, etc.
 *
 * Key functions:
 *   - tosSetNumberFormatForHeaderColumn()   → core formatting engine
 *   - tosFormatHeaderColumnAsText()         → force one column to Plain Text
 *   - tosFormatHeaderColumnsAsText()        → force multiple columns to Plain Text
 *
 * IMPORTANT: Always call these BEFORE setValues() when writing data that
 * contains text that looks like numbers. Once Sheets coerces a value to a
 * number, leading zeros are permanently lost.
 */
/**
 * Set number format for a column identified by its header text.
 *
 * IMPORTANT: Call this BEFORE setValues() so Sheets doesn't coerce all-digit strings
 * into numbers (which permanently drops leading zeros like 090628207 -> 90628207).
 *
 * @return {boolean} true if header found and format applied, false if header not found.
 */
function tosSetNumberFormatForHeaderColumn(sheet, headerRow, headerName, rowsCount, numberFormat, startRow) {
  startRow = startRow || 1;

  if (!sheet) throw new Error('tosSetNumberFormatForHeaderColumn: sheet was null/undefined.');
  if (!Array.isArray(headerRow) || headerRow.length === 0) return false;

  const idx0 = headerRow.indexOf(headerName);
  if (idx0 < 0) return false;

  const colA1 = idx0 + 1;
  const nRows = Math.max(1, Number(rowsCount || 1));

  sheet.getRange(startRow, colA1, nRows, 1).setNumberFormat(numberFormat);
  return true;
}

/**
 * Force a header column to Plain text (@).
 * Purpose: preserve leading zeros (ex: CUSIPs like "090628207").
 */
function tosFormatHeaderColumnAsText(sheet, headerRow, headerName, rowsCount, startRow) {
  return tosSetNumberFormatForHeaderColumn(sheet, headerRow, headerName, rowsCount, '@', startRow);
}


// Convenience wrapper for multiple text columns (TIME, Symbol, etc.).
// IMPORTANT: Call this BEFORE setValues() so Sheets doesn't coerce all-digit strings.
function tosFormatHeaderColumnsAsText(sheet, headerRow, headerNames, rowsCount, startRow) {
  const names = Array.isArray(headerNames) ? headerNames : [headerNames];
  let appliedAny = false;

  for (let i = 0; i < names.length; i++) {
    const ok = tosFormatHeaderColumnAsText(sheet, headerRow, names[i], rowsCount, startRow);
    if (ok) appliedAny = true;
  }

  return appliedAny;
}
