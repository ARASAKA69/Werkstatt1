var KISTEN_SHEET_ID   = '1UP_SaKCrrilj_K03K-8a1sVhAyEIbaH5fT8dkobuuiE';
var REFURB_SHEET_NAME = 'Refurbisment List';
var KISTEN_SHEET_NAME = 'Lager Kisten Klärung';

var REFURB_COL_AB = 28;
var REFURB_COL_B  = 2;
var KISTEN_COL_C  = 3;
var KISTEN_COL_F  = 6;

var GREEN = '#34a853';

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

function triggerCheck() {
  var triggers = ScriptApp.getProjectTriggers();
  var gefunden = false;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onRefurbEdit') {
      gefunden = true;
      break;
    }
  }
  if (gefunden) {
    SpreadsheetApp.getUi().alert('✅ Trigger ist aktiv, alles gut.');
  } else {
    SpreadsheetApp.getUi().alert('❌ Trigger ist NICHT aktiv!\n\nBitte "installTrigger" ausfuehren.');
  }
}

function onRefurbEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== REFURB_SHEET_NAME) return;
  syncAlleTagesliste();
}

function syncAlleTagesliste() {
  var refurbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REFURB_SHEET_NAME);
  if (!refurbSheet) return 0;

  var lastRow = refurbSheet.getLastRow();
  if (lastRow < 5) return 0;

  var abWerte = refurbSheet.getRange(5, REFURB_COL_AB, lastRow - 4, 1).getValues();
  var bWerte  = refurbSheet.getRange(5, REFURB_COL_B, lastRow - 4, 1).getValues();

  var kistenSS = SpreadsheetApp.openById(KISTEN_SHEET_ID);
  var kistenSheet = kistenSS.getSheetByName(KISTEN_SHEET_NAME);
  if (!kistenSheet) kistenSheet = kistenSS.getSheets()[0];

  var kistenLastRow = kistenSheet.getLastRow();
  if (kistenLastRow < 2) return 0;

  var kistenStockIds = kistenSheet.getRange(2, KISTEN_COL_C, kistenLastRow - 1, 1).getValues();
  var kistenStatusF  = kistenSheet.getRange(2, KISTEN_COL_F, kistenLastRow - 1, 1).getValues();
  var kistenLastCol  = kistenSheet.getLastColumn();

  var gesynct = 0;

  for (var r = 0; r < abWerte.length; r++) {
    if (String(abWerte[r][0]).trim() !== 'Tagesliste') continue;

    var stockId = String(bWerte[r][0]).trim();
    if (!stockId) continue;

    for (var k = 0; k < kistenStockIds.length; k++) {
      if (String(kistenStockIds[k][0]).trim() !== stockId) continue;
      if (String(kistenStatusF[k][0]).trim() === 'Tagesliste') continue;

      var targetRow = k + 2;
      kistenSheet.getRange(targetRow, 1, 1, kistenLastCol).setBackground(GREEN);
      kistenSheet.getRange(targetRow, KISTEN_COL_F).setValue('Tagesliste');
      kistenStatusF[k][0] = 'Tagesliste';
      gesynct++;
    }
  }

  return gesynct;
}

function manuellerSync() {
  var anzahl = syncAlleTagesliste();
  if (anzahl > 0) {
    SpreadsheetApp.getUi().alert('✅ ' + anzahl + ' Zeile(n) im Kisten-Sheet aktualisiert.');
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ Alles bereits synchron, nichts zu tun.');
  }
}
