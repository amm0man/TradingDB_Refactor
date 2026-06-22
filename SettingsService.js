
function setSetting(key, value) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(String(key), String(value));
}

function getSetting(key, defaultValue) {
  var props = PropertiesService.getScriptProperties();
  var v = props.getProperty(String(key));
  return (v === null || v === undefined) ? defaultValue : v;
}

function getAllSettings() {
  var props = PropertiesService.getScriptProperties().getProperties();
  return props;
}

function promptSetAccountMode() {
  var ui = SpreadsheetApp.getUi();
  var current = getSetting('accountMode', 'DT');
  // Simple choice dialog: DT or LT
  var response = ui.prompt('Set Account Mode (DT or LT)',
    'Enter account mode: DT (Day Trading) or LT (Long Term). Current: ' + current,
    ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var val = response.getResponseText().trim().toUpperCase();
    if (val !== 'DT' && val !== 'LT') {
      ui.alert('Invalid value. Enter DT or LT (case-insensitive).');
      return;
    }
    setSetting('accountMode', val);

    // Legacy compatibility key. Safe to delete ONLY after confirming no other scripts read "TOS_ACCOUNT_MODE".
    setSetting('TOS_ACCOUNT_MODE', val);

    ui.alert('Account mode set to ' + val);
  }
}

/**
* Toggle the TOS import DEBUG popups (folder sanity checks and other TOS import diagnostics).
* Stored in ScriptProperties as: TOS_IMPORT_DEBUG_ALERTS = "0" or "1"
*/
function promptToggleTosImportDebugAlerts() {
  var ui = SpreadsheetApp.getUi();
  var key = 'TOS_IMPORT_DEBUG_ALERTS';

  var current = String(getSetting(key, '0')).trim();
  var isOn = (current === '1');

  var msg =
    'TOS Import DEBUG Alerts are currently: ' + (isOn ? 'ON' : 'OFF') + '\n\n' +
    'Choose YES to turn ON.\n' +
    'Choose NO to turn OFF.';

  var resp = ui.alert('Toggle TOS Import DEBUG Alerts', msg, ui.ButtonSet.YES_NO);
  var newVal = (resp === ui.Button.YES) ? '1' : '0';

  setSetting(key, newVal);
  ui.alert('Saved: ' + key + ' = ' + newVal);
}

function showCurrentSettingsDialog() {
  var ui = SpreadsheetApp.getUi();
  var s = getAllSettings();
  var message = 'Current script settings:\n';
  for (var k in s) {
    message += k + ' = ' + s[k] + '\n';
  }
  ui.alert(message);
}


