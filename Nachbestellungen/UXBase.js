var UX_NB_DELBUF_KEY = "ux_nb_delbuf_v1";
var UX_NB_DELBUF_MAX = 100;

function uxBaseMenueEinrichten() {
  SpreadsheetApp.getUi()
    .createMenu("UX")
    .addItem("Gelöschte Einträge…", "uxBaseDialogGeloeschteEintraegeOeffnen")
    .addToUi();
}

function onOpen() {
  uxBaseMenueEinrichten();
}

function uxBaseDialogGeloeschteEintraegeOeffnen() {
  var html = HtmlService.createHtmlOutputFromFile("UXMenu")
    .setWidth(920)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, "UX · Gelöschte Einträge");
}

function uxNbDelBufferListe() {
  return uxNbDelBufferRead_();
}

function uxNbDelBufferEndgueltigEntfernen(id) {
  if (!id) return { ok: false, fehler: "Keine ID." };
  return uxNbDelBufferWithLock_(function (arr) {
    var next = [];
    for (var i = 0; i < arr.length; i++) {
      if (String(arr[i].id) !== String(id)) next.push(arr[i]);
    }
    return { ok: true, arr: next };
  });
}

function uxNbDelBufferWiederherstellen(id) {
  if (!id) return { ok: false, fehler: "Keine ID." };
  var lock = LockService.getScriptLock();
  var got = false;
  try {
    got = lock.tryLock(20000);
    if (!got) return { ok: false, fehler: "Konnte keine Sperre bekommen. Bitte erneut versuchen." };
    var arr = uxNbDelBufferReadRaw_();
    var entry = null;
    var rest = [];
    for (var i = 0; i < arr.length; i++) {
      if (String(arr[i].id) === String(id)) {
        entry = arr[i];
      } else {
        rest.push(arr[i]);
      }
    }
    if (!entry) return { ok: false, fehler: "Eintrag nicht gefunden (evtl. schon entfernt)." };
    var res = uxNbDelBufferRestoreEntry_(entry);
    if (res.ok) uxNbDelBufferWriteRaw_(rest);
    return res;
  } catch (e) {
    return { ok: false, fehler: String(e) };
  } finally {
    if (got) {
      try {
        lock.releaseLock();
      } catch (e2) {}
    }
  }
}

function uxNbStockIdLoeschungFuerWiederherstellungSpeichern_(ss, entryId, stockKey) {
  try {
    var snap = uxNbDelBufferBuildSnapshot_(ss, entryId, stockKey);
    if (!snap || !snap.sheets) return;
    var names = Object.keys(snap.sheets);
    if (!names.length) return;
    uxNbDelBufferWithLock_(function (arr) {
      var item = {
        id: Utilities.getUuid(),
        deletedAt: new Date().toISOString(),
        deletedBy: uxNbDelBufferUser_(),
        entryId: String(entryId || "").trim(),
        stockId: stockKey ? String(stockKey).trim().toUpperCase() : "",
        sheets: snap.sheets
      };
      arr.unshift(item);
      while (arr.length > UX_NB_DELBUF_MAX) arr.pop();
      return { ok: true, arr: arr };
    });
  } catch (e) {}
}

function uxNbDelBufferUser_() {
  try {
    var a = Session.getActiveUser().getEmail();
    if (a) return a;
  } catch (e) {}
  try {
    return Session.getEffectiveUser().getEmail() || "";
  } catch (e2) {}
  return "";
}

function uxNbDelBufferSerializeCell_(v) {
  if (v instanceof Date) return { __uxNbD: v.getTime() };
  return v;
}

function uxNbDelBufferDeserializeCell_(v) {
  if (v && typeof v === "object" && typeof v.__uxNbD === "number") return new Date(v.__uxNbD);
  return v;
}

function uxNbDelBufferSerializeRow_(row) {
  var o = [];
  for (var i = 0; i < row.length; i++) o.push(uxNbDelBufferSerializeCell_(row[i]));
  return o;
}

function uxNbDelBufferDeserializeRow_(row) {
  var o = [];
  for (var i = 0; i < row.length; i++) o.push(uxNbDelBufferDeserializeCell_(row[i]));
  return o;
}

function uxNbDelBufferBuildSnapshot_(ss, entryId, stockKey) {
  var targetId = String(entryId || "").trim();
  if (!targetId || !ss) return null;
  var sheets = {};

  try {
    var dashboard = ss.getSheetByName("Dashboard");
    if (dashboard) {
      var dLayout = getDashboardLayout(dashboard);
      var dRow = findRowByEntryId(dashboard, dLayout.cols.entryId, targetId, dLayout.dataStartRow);
      if (dRow !== -1) {
        var dLast = Math.max(dashboard.getLastColumn(), dLayout.lastCol || 0);
        var dVals = dashboard.getRange(dRow, 1, 1, dLast).getValues()[0];
        sheets.Dashboard = { row: dRow, values: uxNbDelBufferSerializeRow_(dVals) };
      }
    }
  } catch (e) {}

  try {
    var nb = ss.getSheetByName("Nachbestellung");
    if (nb && nb.getLastRow() >= 2) {
      var nbEid = getEntryIdCol(nb);
      var nbRow = findRowByEntryId(nb, nbEid, targetId, 2);
      if (nbRow !== -1) {
        var nbLast = nb.getLastColumn();
        var nbVals = nb.getRange(nbRow, 1, 1, nbLast).getValues()[0];
        sheets.Nachbestellung = { row: nbRow, values: uxNbDelBufferSerializeRow_(nbVals) };
      }
    }
  } catch (e2) {}

  for (var k = 0; k < INPUT_SHEET_NAMES.length; k++) {
    var sn = INPUT_SHEET_NAMES[k];
    try {
      var sh = ss.getSheetByName(sn);
      if (!sh || sh.getLastRow() < 2) continue;
      var inEid = getEntryIdCol(sh);
      var inRow = findRowByEntryId(sh, inEid, targetId, 2);
      if (inRow === -1) continue;
      var inLast = sh.getLastColumn();
      var inVals = sh.getRange(inRow, 1, 1, inLast).getValues()[0];
      sheets[sn] = { row: inRow, values: uxNbDelBufferSerializeRow_(inVals) };
    } catch (e3) {}
  }

  return { sheets: sheets };
}

function uxNbDelBufferReadRaw_() {
  var raw = PropertiesService.getScriptProperties().getProperty(UX_NB_DELBUF_KEY);
  if (!raw) return [];
  try {
    var arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function uxNbDelBufferWriteRaw_(arr) {
  PropertiesService.getScriptProperties().setProperty(UX_NB_DELBUF_KEY, JSON.stringify(arr || []));
}

function uxNbDelBufferRead_() {
  return uxNbDelBufferWithLock_(function (arr) {
    return { ok: true, arr: arr, readOnly: true };
  });
}

function uxNbDelBufferWithLock_(fn) {
  var lock = LockService.getScriptLock();
  var got = false;
  try {
    got = lock.tryLock(15000);
    if (!got) return { ok: false, fehler: "Sperre nicht verfügbar.", arr: [] };
    var arr = uxNbDelBufferReadRaw_();
    var out = fn(arr);
    if (out && out.readOnly) return { ok: true, arr: arr };
    if (out && out.ok && out.arr) uxNbDelBufferWriteRaw_(out.arr);
    return out || { ok: false, fehler: "Unbekannter Fehler." };
  } catch (e) {
    return { ok: false, fehler: String(e), arr: [] };
  } finally {
    if (got) {
      try {
        lock.releaseLock();
      } catch (e2) {}
    }
  }
}

function uxNbDelBufferRowExists_(sheet, entryIdCol, entryId, startRow) {
  return findRowByEntryId(sheet, entryIdCol, entryId, startRow) !== -1;
}

function uxNbDelBufferInsertRow_(sheet, desiredRow) {
  var last = sheet.getLastRow();
  var ins = Math.min(Math.max(2, desiredRow), last + 1);
  sheet.insertRowBefore(ins);
  return ins;
}

function uxNbDelBufferRestoreEntry_(entry) {
  var ss = getMainSS();
  if (!ss) return { ok: false, fehler: "Hauptmappe nicht erreichbar." };
  var entryId = String(entry.entryId || "").trim();
  if (!entryId) return { ok: false, fehler: "Keine Entry-ID im Puffer." };
  var hinweise = [];
  var restored = [];

  for (var sheetName in entry.sheets) {
    if (!Object.prototype.hasOwnProperty.call(entry.sheets, sheetName)) continue;
    var pack = entry.sheets[sheetName];
    if (!pack || !pack.values) continue;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      hinweise.push(sheetName + ": Tabellenblatt fehlt.");
      continue;
    }
    var startRow = 2;
    var eidCol = 15;
    try {
      if (sheetName === "Dashboard") {
        var lo = getDashboardLayout(sheet);
        startRow = lo.dataStartRow;
        eidCol = lo.cols.entryId;
      } else {
        eidCol = getEntryIdCol(sheet);
      }
    } catch (e) {
      eidCol = getEntryIdCol(sheet);
    }
    if (uxNbDelBufferRowExists_(sheet, eidCol, entryId, startRow)) {
      hinweise.push(sheetName + ": Zeile mit dieser Entry-ID existiert bereits, übersprungen.");
      continue;
    }
    var vals = uxNbDelBufferDeserializeRow_(pack.values);
    var packRow = pack.row;
    if (typeof packRow !== "number" || packRow < startRow) packRow = startRow;
    var rowAfter;
    try {
      rowAfter = uxNbDelBufferInsertRow_(sheet, packRow);
    } catch (e2) {
      hinweise.push(sheetName + ": Einfügen fehlgeschlagen (" + String(e2) + ").");
      continue;
    }
    var w = Math.max(sheet.getLastColumn(), vals.length);
    var rowWide = [];
    for (var c = 0; c < w; c++) rowWide.push(c < vals.length ? vals[c] : "");
    try {
      sheet.getRange(rowAfter, 1, 1, w).setValues([rowWide]);
    } catch (e3) {
      hinweise.push(sheetName + ": Schreiben fehlgeschlagen (" + String(e3) + ").");
      try {
        sheet.deleteRow(rowAfter);
      } catch (e4) {}
      continue;
    }
    if (INPUT_SHEETS[sheetName]) {
      try {
        applyStockIdProtection(sheet, rowAfter);
      } catch (e5) {}
    }
    restored.push(sheetName);
  }

  if (!restored.length) {
    return {
      ok: false,
      fehler: "Nichts wiederhergestellt. " + (hinweise.length ? hinweise.join(" ") : "")
    };
  }
  return { ok: true, wiederhergestelltIn: restored, hinweise: hinweise };
}
