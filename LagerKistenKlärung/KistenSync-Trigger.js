var KISTEN_SHEET_ID = '1UP_SaKCrrilj_K03K-8a1sVhAyEIbaH5fT8dkobuuiE';
var REFURB_SHEET_NAME = 'Refurbisment List';
var KISTEN_SHEET_NAME = 'Lager Kisten Klärung';
var COL_AB = 28;
var COL_B  = 2;
var COL_C  = 3;
var GREEN  = '#34a853';

function installTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onRefurbEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('onRefurbEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert('✅ Kisten-Sync Trigger installiert!');
}

function onRefurbEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== REFURB_SHEET_NAME) return;

  var col = e.range.getColumn();
  var row = e.range.getRow();

  if (col !== COL_AB || row <= 4) return;

  var newValue = String(e.range.getValue()).trim();
  if (newValue !== 'Tagesliste') return;

  var stockId = String(sheet.getRange(row, COL_B).getValue()).trim();
  if (!stockId) return;

  var kistenSS = SpreadsheetApp.openById(KISTEN_SHEET_ID);
  var kistenSheet = kistenSS.getSheetByName(KISTEN_SHEET_NAME);
  if (!kistenSheet) kistenSheet = kistenSS.getSheets()[0];

  var lastRow = kistenSheet.getLastRow();
  if (lastRow < 2) return;

  var stockIds = kistenSheet.getRange(2, COL_C, lastRow - 1, 1).getValues();

  for (var i = 0; i < stockIds.length; i++) {
    if (String(stockIds[i][0]).trim() === stockId) {
      var targetRow = i + 2;
      var lastCol = kistenSheet.getLastColumn();
      kistenSheet.getRange(targetRow, 2, 1, lastCol - 1).setBackground(GREEN);
      return;
    }
  }
}
