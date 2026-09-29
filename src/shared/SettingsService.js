/**
 * SettingsService.js
 *
 * Simple centralized service for storing and retrieving script-level settings
 * using Google Apps Script's PropertiesService (ScriptProperties).
 *
 * Currently used for:
 *   - 'TOS_IMPORT_DEBUG_ALERTS' → '0' (off) or '1' (on) for import diagnostics
 *   - folder IDs, Issues write-mode, ACTIVE_IMPORT_RUN_ID (set by other files)
 *
 * accountMode / TOS_ACCOUNT_MODE were removed 2026-09-29. The ScriptProperty
 * may still appear in Show Current Settings until it is deleted by hand.
 * Do not read it. Both accounts always run.
 *
 * All values are stored as strings. Use getSetting(key, defaultValue) when
 * you want a fallback if the setting has never been set.
 *
 * Functions in this file are called from menus and from other phases.
 */
function setSetting(key, value) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(String(key), String(value));
}

function getSetting(key, defaultValue) {
  var props = PropertiesService.getScriptProperties();
  var v = props.getProperty(String(key));
  return v === null || v === undefined ? defaultValue : v;
}

function getAllSettings() {
  var props = PropertiesService.getScriptProperties().getProperties();
  return props;
}

/**
 * Toggle the TOS import DEBUG popups on or off.
 *
 * These popups are used for folder sanity checks and other diagnostics
 * during the TOS/Schwab import process (Phase 1).
 *
 * The setting is stored in ScriptProperties as:
 *   TOS_IMPORT_DEBUG_ALERTS = "0"  → OFF (default)
 *   TOS_IMPORT_DEBUG_ALERTS = "1"  → ON
 */
function promptToggleTosImportDebugAlerts() {
  var ui = SpreadsheetApp.getUi();
  var key = "TOS_IMPORT_DEBUG_ALERTS";

  // Read current value (default to "0" = OFF if never set)
  var current = String(getSetting(key, "0")).trim();
  var isOn = current === "1";

  var msg =
    "TOS Import DEBUG Alerts are currently: " +
    (isOn ? "ON" : "OFF") +
    "\n\n" +
    "Choose YES to turn ON.\n" +
    "Choose NO to turn OFF.";

  var resp = ui.alert(
    "Toggle TOS Import DEBUG Alerts",
    msg,
    ui.ButtonSet.YES_NO,
  );

  // Save the new value based on the user's choice
  var newVal = resp === ui.Button.YES ? "1" : "0";
  setSetting(key, newVal);

  ui.alert("Saved: " + key + " = " + newVal);
}

function showCurrentSettingsDialog() {
  var ui = SpreadsheetApp.getUi();
  var s = getAllSettings();
  var message = "Current script settings:\n";
  for (var k in s) {
    message += k + " = " + s[k] + "\n";
  }
  ui.alert(message);
}
