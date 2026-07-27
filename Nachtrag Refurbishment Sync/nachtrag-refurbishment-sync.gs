var CONFIG = {
  sourceSheetName: 'Nachtrag Refurbishment',
  targetSpreadsheetId: '13Oh7gDT8NAul2s0cwQUeaGwMcS3B2MYu0QOdFNMhXzM',
  targetSheetName: 'Refurbisment List',
  sourceStockCol: 1,
  sourceSchaedenCol: 2,
  sourceKommentarCol: 3,
  sourceStatusCol: 5,
  sourceReadyCol: 6,
  targetStockCol: 2,
  targetSchaedenCol: 23,
  targetKommentarCol: 24,
  targetStatusCol: 26,
  targetDataStartRow: 5,
  sourceDataStartRow: 2,
  readyMarker: 'x',
  doneMarker: 'OK',
  intervalMinutes: 2,
  activeFromHour: 6,
  activeToHour: 21,
  timezone: 'Europe/Berlin'
};

function setupTrigger() {
  removeTriggers_();
  ScriptApp.newTrigger('syncNachtragRefurbishment')
    .timeBased()
    .everyMinutes(1)
    .create();
  SpreadsheetApp.getUi().alert('Trigger aktiv: Sync laeuft ca. alle ' + CONFIG.intervalMinutes + ' Minuten.');
}

function removeTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncNachtragRefurbishment') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function isWithinActiveHours_() {
  var hour = Number(Utilities.formatDate(new Date(), CONFIG.timezone, 'H'));
  return hour >= CONFIG.activeFromHour && hour < CONFIG.activeToHour;
}

function syncNachtragRefurbishment() {
  if (!isWithinActiveHours_()) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    var props = PropertiesService.getScriptProperties();
    var now = Date.now();
    var last = Number(props.getProperty('nachtragSyncLastRun') || 0);
    if (now - last < (CONFIG.intervalMinutes * 60 - 5) * 1000) return;
    props.setProperty('nachtragSyncLastRun', String(now));
    runSync_();
  } finally {
    lock.releaseLock();
  }
}

function manualSync() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) {
    SpreadsheetApp.getUi().alert('Sync laeuft bereits, bitte kurz warten.');
    return;
  }
  try {
    var result = runSync_();
    SpreadsheetApp.getUi().alert(
      'Fertig.\n' +
      'Uebertragen: ' + result.transferred + '\n' +
      'Warten (Stock ID noch nicht in Refurbisment List): ' + result.waiting + '\n' +
      'Fehler: ' + result.errors
    );
  } finally {
    lock.releaseLock();
  }
}

function runSync_() {
  var result = { transferred: 0, waiting: 0, errors: 0 };
  var sourceSs = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = sourceSs.getSheetByName(CONFIG.sourceSheetName);
  if (!sourceSheet) return result;

  var sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < CONFIG.sourceDataStartRow) return result;

  var numRows = sourceLastRow - CONFIG.sourceDataStartRow + 1;
  var sourceData = sourceSheet
    .getRange(CONFIG.sourceDataStartRow, 1, numRows, CONFIG.sourceReadyCol)
    .getValues();

  var pending = [];
  for (var i = 0; i < sourceData.length; i++) {
    var ready = String(sourceData[i][CONFIG.sourceReadyCol - 1] || '').trim().toLowerCase();
    if (ready !== CONFIG.readyMarker) continue;
    var stockId = String(sourceData[i][CONFIG.sourceStockCol - 1] || '').trim();
    if (!stockId) continue;
    pending.push({
      rowIndex: i,
      sheetRow: CONFIG.sourceDataStartRow + i,
      stockId: stockId,
      schaeden: sourceData[i][CONFIG.sourceSchaedenCol - 1],
      kommentar: sourceData[i][CONFIG.sourceKommentarCol - 1],
      status: sourceData[i][CONFIG.sourceStatusCol - 1]
    });
  }

  if (!pending.length) return result;

  var targetSs = SpreadsheetApp.openById(CONFIG.targetSpreadsheetId);
  var targetSheet = targetSs.getSheetByName(CONFIG.targetSheetName);
  if (!targetSheet) {
    result.errors = pending.length;
    return result;
  }

  var targetLastRow = targetSheet.getLastRow();
  if (targetLastRow < CONFIG.targetDataStartRow) {
    result.waiting = pending.length;
    return result;
  }

  var targetNumRows = targetLastRow - CONFIG.targetDataStartRow + 1;
  var targetStockIds = targetSheet
    .getRange(CONFIG.targetDataStartRow, CONFIG.targetStockCol, targetNumRows, 1)
    .getValues();

  var stockToRow = {};
  for (var t = 0; t < targetStockIds.length; t++) {
    var id = String(targetStockIds[t][0] || '').trim().toUpperCase();
    if (!id) continue;
    if (!stockToRow.hasOwnProperty(id)) {
      stockToRow[id] = CONFIG.targetDataStartRow + t;
    }
  }

  for (var p = 0; p < pending.length; p++) {
    var item = pending[p];
    var targetRow = stockToRow[item.stockId.toUpperCase()];
    if (!targetRow) {
      result.waiting++;
      continue;
    }
    try {
      targetSheet.getRange(targetRow, CONFIG.targetSchaedenCol).setValue(item.schaeden);
      targetSheet.getRange(targetRow, CONFIG.targetKommentarCol).setValue(item.kommentar);
      targetSheet.getRange(targetRow, CONFIG.targetStatusCol).setValue(item.status);
      sourceSheet.getRange(item.sheetRow, CONFIG.sourceReadyCol).setValue(CONFIG.doneMarker);
      result.transferred++;
    } catch (err) {
      result.errors++;
    }
  }

  return result;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Nachtrag Sync')
    .addItem('Jetzt manuell syncen', 'manualSync')
    .addItem('Trigger einrichten (alle ~2 Min)', 'setupTrigger')
    .addToUi();
}
