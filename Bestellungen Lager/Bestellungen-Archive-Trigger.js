var BL_SHEET_ID = "1H-iQQNsvvZr5gkaUrgV1Xzj9pGonYFTW73XpTEPxKKQ";
var BL_SOURCE_SHEET = "Bestellungen";
var BL_ARCHIVE_SHEET = "ARCHIVE";
var BL_STATUS_COL = 7;
var BL_DATA_START_COL = 2;
var BL_DATA_END_COL = 8;
var BL_DATA_START_ROW = 5;
var BL_HEADER_ROW = 4;

function onEdit(e) {
  if (!e || !e.range) return;
  handleBestellungenGeliefertEdit_(e);
}

function handleBestellungenGeliefertEdit_(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== BL_SOURCE_SHEET) return;

  var startRow = e.range.getRow();
  var endRow = e.range.getLastRow();
  var startCol = e.range.getColumn();
  var endCol = e.range.getLastColumn();
  if (endCol < BL_STATUS_COL || startCol > BL_STATUS_COL) return;

  var rowsToMove = [];
  for (var r = startRow; r <= endRow; r++) {
    if (r < BL_DATA_START_ROW) continue;
    var statusCell = sheet.getRange(r, BL_STATUS_COL);
    if (!isGeliefertStatus_(statusCell.getValue(), statusCell.getDisplayValue())) continue;
    rowsToMove.push(r);
  }
  if (rowsToMove.length === 0) return;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    rowsToMove.sort(function (a, b) {
      return b - a;
    });
    var ss = e.source;
    var archive = ss.getSheetByName(BL_ARCHIVE_SHEET);
    if (!archive) return;
    for (var i = 0; i < rowsToMove.length; i++) {
      moveBestellungenRowToArchive_(sheet, archive, rowsToMove[i]);
    }
  } finally {
    lock.releaseLock();
  }
}

function moveBestellungenRowToArchive_(sourceSheet, archiveSheet, row) {
  var statusCell = sourceSheet.getRange(row, BL_STATUS_COL);
  if (!isGeliefertStatus_(statusCell.getValue(), statusCell.getDisplayValue())) return;

  var rowValues = sourceSheet.getRange(row, BL_DATA_START_COL, 1, BL_DATA_END_COL - BL_DATA_START_COL + 1).getValues()[0];
  if (isBestellungenRowEmpty_(rowValues)) return;

  var targetRow = getBestellungenArchiveAppendRow_(archiveSheet);
  archiveSheet.getRange(targetRow, BL_DATA_START_COL, 1, BL_DATA_END_COL - BL_DATA_START_COL + 1).setValues([rowValues]);
  sourceSheet.deleteRow(row);
}

function getBestellungenArchiveAppendRow_(archiveSheet) {
  var lastRow = archiveSheet.getLastRow();
  if (lastRow < BL_DATA_START_ROW) return BL_DATA_START_ROW;
  var scanStart = Math.max(BL_DATA_START_ROW, lastRow - 500);
  var numRows = lastRow - scanStart + 1;
  var numCols = BL_DATA_END_COL - BL_DATA_START_COL + 1;
  var values = archiveSheet.getRange(scanStart, BL_DATA_START_COL, numRows, numCols).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (!isBestellungenRowEmpty_(values[i])) return scanStart + i + 1;
  }
  return BL_DATA_START_ROW;
}

function isBestellungenRowEmpty_(rowValues) {
  for (var i = 0; i < rowValues.length; i++) {
    if (String(rowValues[i] == null ? "" : rowValues[i]).trim() !== "") return false;
  }
  return true;
}

function isGeliefertStatus_(value, displayValue) {
  return normBestellungenStatus_(value) === "geliefert" || normBestellungenStatus_(displayValue) === "geliefert";
}

function normBestellungenStatus_(val) {
  return String(val == null ? "" : val)
    .replace(/\u00a0/g, " ")
    .replace(/[\u200b-\u200d\ufeff]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function getBestellungenSpreadsheet_() {
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (err) {}
  return SpreadsheetApp.openById(BL_SHEET_ID);
}

function moveAllGeliefertToArchive() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(120000)) return;
  try {
    var ss = getBestellungenSpreadsheet_();
    var source = ss.getSheetByName(BL_SOURCE_SHEET);
    var archive = ss.getSheetByName(BL_ARCHIVE_SHEET);
    if (!source || !archive) return;

    var lastRow = source.getLastRow();
    if (lastRow < BL_DATA_START_ROW) return;

    for (var row = lastRow; row >= BL_DATA_START_ROW; row--) {
      var statusCell = source.getRange(row, BL_STATUS_COL);
      if (!isGeliefertStatus_(statusCell.getValue(), statusCell.getDisplayValue())) continue;
      moveBestellungenRowToArchive_(source, archive, row);
    }
  } finally {
    lock.releaseLock();
  }
}

function installBestellungenArchiveOnEditTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onEditBestellungenArchive") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("onEditBestellungenArchive")
    .forSpreadsheet(BL_SHEET_ID)
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert("Bestellungen-Archiv Trigger installiert.");
}

function onEditBestellungenArchive(e) {
  handleBestellungenGeliefertEdit_(e);
}
