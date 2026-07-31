var BL_SHEET_ID = "1H-iQQNsvvZr5gkaUrgV1Xzj9pGonYFTW73XpTEPxKKQ";
var BL_SOURCE_SHEET = "Bestellungen";
var BL_ARCHIVE_SHEET = "ARCHIVE";
var BL_STATUS_COL = 7;
var BL_DATA_START_COL = 2;
var BL_DATA_END_COL = 8;
var BL_DATA_START_ROW = 5;
var BL_HEADER_ROW = 4;
var BL_DEBOUNCE_MS = 5000;

function handleBestellungenGeliefertEdit_(e) {
  if (!e || !e.range) return;
  var cache = CacheService.getScriptCache();

  var sheet = e.range.getSheet();
  if (sheet.getName() !== BL_SOURCE_SHEET) return;

  var startRow = e.range.getRow();
  var endRow = e.range.getLastRow();
  var startCol = e.range.getColumn();
  var endCol = e.range.getLastColumn();
  if (endCol < BL_STATUS_COL || startCol > BL_STATUS_COL) return;
  if (endRow < BL_DATA_START_ROW) return;

  var numCols = BL_DATA_END_COL - BL_DATA_START_COL + 1;
  var candidates = [];

  if (startCol === BL_STATUS_COL && endCol === BL_STATUS_COL && typeof e.value !== "undefined") {
    if (startRow !== endRow) return;
    if (startRow < BL_DATA_START_ROW) return;
    if (!isGeliefertStatus_(e.value, e.value)) return;
    if (typeof e.oldValue !== "undefined" && isGeliefertStatus_(e.oldValue, e.oldValue)) return;
    var snap = sheet.getRange(startRow, BL_DATA_START_COL, 1, numCols).getValues()[0];
    if (isBestellungenRowEmpty_(snap)) return;
    candidates.push({ row: startRow, values: snap.slice() });
  } else {
    var statusValues = sheet.getRange(startRow, BL_STATUS_COL, endRow - startRow + 1, 1).getValues();
    var statusDisplays = sheet.getRange(startRow, BL_STATUS_COL, endRow - startRow + 1, 1).getDisplayValues();
    var rowBlock = sheet.getRange(startRow, BL_DATA_START_COL, endRow - startRow + 1, numCols).getValues();
    for (var i = 0; i < statusValues.length; i++) {
      var r = startRow + i;
      if (r < BL_DATA_START_ROW) continue;
      if (!isGeliefertStatus_(statusValues[i][0], statusDisplays[i][0])) continue;
      if (isBestellungenRowEmpty_(rowBlock[i])) continue;
      candidates.push({ row: r, values: rowBlock[i].slice() });
    }
  }

  if (candidates.length === 0) return;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    var ss = e.source || sheet.getParent();
    var archive = ss.getSheetByName(BL_ARCHIVE_SHEET);
    if (!archive) return;

    candidates.sort(function (a, b) {
      return b.row - a.row;
    });

    for (var c = 0; c < candidates.length; c++) {
      moveBestellungenCandidateToArchive_(sheet, archive, candidates[c], cache);
    }
  } finally {
    lock.releaseLock();
  }
}

function moveBestellungenCandidateToArchive_(sourceSheet, archiveSheet, candidate, cache) {
  cache = cache || CacheService.getScriptCache();
  var row = candidate.row;
  var expected = candidate.values;
  var numCols = BL_DATA_END_COL - BL_DATA_START_COL + 1;
  var statusCell = sourceSheet.getRange(row, BL_STATUS_COL);
  if (!isGeliefertStatus_(statusCell.getValue(), statusCell.getDisplayValue())) return false;

  var current = sourceSheet.getRange(row, BL_DATA_START_COL, 1, numCols).getValues()[0];
  if (!bestellungenRowValuesMatch_(expected, current)) return false;
  if (isBestellungenRowEmpty_(current)) return false;

  var dedupeKey = "blArch:" + bestellungenRowFingerprint_(current);
  if (cache.get(dedupeKey)) return false;

  var targetRow = getBestellungenArchiveAppendRow_(archiveSheet);
  archiveSheet.getRange(targetRow, BL_DATA_START_COL, 1, numCols).setValues([current]);
  cache.put(dedupeKey, "1", Math.ceil(BL_DEBOUNCE_MS / 1000));
  sourceSheet.deleteRow(row);
  return true;
}

function bestellungenRowValuesMatch_(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  for (var i = 0; i < a.length; i++) {
    if (normBestellungenCompare_(a[i]) !== normBestellungenCompare_(b[i])) return false;
  }
  return true;
}

function bestellungenRowFingerprint_(rowValues) {
  var parts = [];
  for (var i = 0; i < rowValues.length; i++) {
    parts.push(normBestellungenCompare_(rowValues[i]));
  }
  return parts.join("|");
}

function normBestellungenCompare_(val) {
  if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val.getTime())) {
    return "d:" + val.getTime();
  }
  return normBestellungenStatus_(val);
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
    var numCols = BL_DATA_END_COL - BL_DATA_START_COL + 1;

    for (var row = lastRow; row >= BL_DATA_START_ROW; row--) {
      var statusCell = source.getRange(row, BL_STATUS_COL);
      if (!isGeliefertStatus_(statusCell.getValue(), statusCell.getDisplayValue())) continue;
      var values = source.getRange(row, BL_DATA_START_COL, 1, numCols).getValues()[0];
      if (isBestellungenRowEmpty_(values)) continue;
      moveBestellungenCandidateToArchive_(source, archive, { row: row, values: values });
    }
  } finally {
    lock.releaseLock();
  }
}

function installBestellungenArchiveOnEditTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === "onEditBestellungenArchive" || fn === "onEdit") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("onEditBestellungenArchive")
    .forSpreadsheet(BL_SHEET_ID)
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert(
    "Bestellungen-Archiv Trigger installiert.\nNur noch ein Trigger aktiv (kein Doppel-Archiv mehr)."
  );
}

function onEditBestellungenArchive(e) {
  handleBestellungenGeliefertEdit_(e);
}
