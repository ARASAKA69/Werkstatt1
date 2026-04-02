var KISTEN_SHEET_ID = '1UP_SaKCrrilj_K03K-8a1sVhAyEIbaH5fT8dkobuuiE';
var KISTEN_SHEET_NAME = 'Lager Kisten Klärung';
var NB_SHEET_NAME = 'Nachbestellungen';

function installNbTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'syncNachbestellt') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('syncNachbestellt')
    .timeBased()
    .everyMinutes(1)
    .create();
  SpreadsheetApp.getUi().alert('✅ Nachbestellt-Sync Trigger installiert! Läuft alle ~2 Minuten.');
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) { val.setHours(0,0,0,0); return val; }
  var parts = String(val).match(/(\d{2})\.(\d{2})\.(\d{4})/);
  if (parts) {
    var d = new Date(parseInt(parts[3]), parseInt(parts[2]) - 1, parseInt(parts[1]));
    d.setHours(0,0,0,0);
    return d;
  }
  var d2 = new Date(val);
  if (isNaN(d2.getTime())) return null;
  d2.setHours(0,0,0,0);
  return d2;
}

function syncNachbestellt() {
  var nbSS = SpreadsheetApp.openById('1PuCLw8UmDjB_pBo_jCZ9rmSD3GJQESHzPoBVu_--MRo');
  var nbSheet = nbSS.getSheetByName(NB_SHEET_NAME);
  if (!nbSheet) return;

  var kistenSS = SpreadsheetApp.openById(KISTEN_SHEET_ID);
  var kistenSheet = kistenSS.getSheetByName(KISTEN_SHEET_NAME);
  if (!kistenSheet) kistenSheet = kistenSS.getSheets()[0];

  var kistenLastRow = kistenSheet.getLastRow();
  if (kistenLastRow < 2) return;
  var kistenData = kistenSheet.getRange(2, 2, kistenLastRow - 1, 5).getValues();

  var kistenMap = {};
  for (var k = 0; k < kistenData.length; k++) {
    var kDate = parseDate(kistenData[k][0]);
    var kStock = String(kistenData[k][1]).trim();
    var kComment = String(kistenData[k][4] == null ? '' : kistenData[k][4]).trim();
    if (!kStock || !kDate) continue;
    if (kComment.indexOf('Nachbestellt') !== -1) continue;
    if (!kistenMap[kStock]) kistenMap[kStock] = [];
    kistenMap[kStock].push({ row: k + 2, date: kDate, comment: kComment });
  }

  if (Object.keys(kistenMap).length === 0) return;

  var nbLastRow = nbSheet.getLastRow();
  if (nbLastRow < 2) return;
  var nbData = nbSheet.getRange(2, 1, nbLastRow - 1, 2).getValues();

  var updates = {};
  for (var n = 0; n < nbData.length; n++) {
    var nbDate = parseDate(nbData[n][0]);
    var nbStock = String(nbData[n][1]).trim();
    if (!nbDate || !nbStock) continue;
    var matches = kistenMap[nbStock];
    if (!matches) continue;
    for (var m = 0; m < matches.length; m++) {
      if (nbDate.getTime() >= matches[m].date.getTime()) {
        var r = matches[m].row;
        if (!updates[r]) {
          updates[r] = matches[m].comment ? matches[m].comment + ', Nachbestellt' : 'Nachbestellt';
        }
      }
    }
  }

  var rows = Object.keys(updates);
  for (var u = 0; u < rows.length; u++) {
    var cell = kistenSheet.getRange(parseInt(rows[u]), 6);
    cell.setValue(updates[rows[u]]);
    cell.setBackground('#ff9900');
  }
}
