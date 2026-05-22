var UX_NB_DELBUF_KEY = "ux_nb_delbuf_v1";
var UX_NB_DELBUF_MAX = 100;
var UX_NB_HUB_CACHE_KEY = "ux_nb_hub_cache_v1";
var UX_NB_NACHT_KEY = "ux_nb_nachtragen_done_v1";

function uxBaseMenueEinrichten() {
  SpreadsheetApp.getUi()
    .createMenu("Nachbestellung")
    .addItem("Control Center…", "uxBaseDialogControlCenterOeffnen")
    .addToUi();
}

function onOpen() {
  uxBaseMenueEinrichten();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function uxBaseDialogControlCenterOeffnen(initialView) {
  var t = HtmlService.createTemplateFromFile("UXMenu");
  t.initialView = initialView || "home";
  var html = t
    .evaluate()
    .setWidth(1200)
    .setHeight(850)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi().showModalDialog(html, "WMS · NB Control Center");
}

function uxNbHubCacheRead_() {
  try {
    var raw = PropertiesService.getUserProperties().getProperty(UX_NB_HUB_CACHE_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function uxNbHubCacheWrite_(data) {
  PropertiesService.getUserProperties().setProperty(UX_NB_HUB_CACHE_KEY, JSON.stringify(data || {}));
}

function uxNbNachtragenDoneRead_() {
  try {
    var raw = PropertiesService.getUserProperties().getProperty(UX_NB_NACHT_KEY);
    var arr = JSON.parse(raw || "[]");
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function uxNbNachtragenDoneWrite_(arr) {
  PropertiesService.getUserProperties().setProperty(UX_NB_NACHT_KEY, JSON.stringify(arr || []));
}

function uxNbNachtragenDoneCleanup_(missing, done) {
  var still = {};
  for (var i = 0; i < (missing || []).length; i++) {
    var eid = String(missing[i].entryId || "").trim();
    if (eid) still[eid] = true;
  }
  var next = [];
  for (var j = 0; j < (done || []).length; j++) {
    if (still[done[j]]) next.push(done[j]);
  }
  return next;
}

function uxNbHubDatenSammeln_() {
  var stats = uxNbHubDashboardStats();
  var del = uxNbDelBufferRead_();
  var sync = uxNbSyncMissingInputsListe();
  var missing = sync.ok ? sync.missing || [] : [];
  var nachtragenDone = uxNbNachtragenDoneCleanup_(missing, uxNbNachtragenDoneRead_());
  uxNbNachtragenDoneWrite_(nachtragenDone);
  return {
    ok: true,
    loadedAt: new Date().toISOString(),
    stats: stats,
    delBuffer: del.ok ? del.arr || [] : [],
    syncMissing: missing,
    syncCount: sync.ok ? sync.count || 0 : 0,
    nachtragenDone: nachtragenDone
  };
}

function uxNbHubCacheLesen() {
  var cached = uxNbHubCacheRead_();
  if (cached && cached.loadedAt) {
    return { ok: true, fromCache: true, data: cached };
  }
  return uxNbHubCacheNeuLaden();
}

function uxNbHubCacheNeuLaden() {
  try {
    var data = uxNbHubDatenSammeln_();
    uxNbHubCacheWrite_(data);
    return { ok: true, fromCache: false, data: data };
  } catch (e) {
    return { ok: false, fehler: String(e), data: null };
  }
}

function uxNbSyncNachtragen(sheetName, row, entryId) {
  var eid = String(entryId || "").trim();
  if (!eid) return { ok: false, fehler: "Keine Entry-ID." };
  var done = uxNbNachtragenDoneRead_();
  if (done.indexOf(eid) !== -1) return { ok: false, fehler: "Bereits angestoßen." };
  var ss = getMainSS();
  if (!ss) return { ok: false, fehler: "Hauptmappe nicht erreichbar." };
  var sheet = ss.getSheetByName(String(sheetName || ""));
  if (!sheet) return { ok: false, fehler: "Input-Blatt nicht gefunden." };
  var r = parseInt(row, 10);
  if (!r || r < 2) return { ok: false, fehler: "Ungültige Zeile." };
  var cell = sheet.getRange(r, 3);
  if (uxNbNormEtDiagnose_(cell.getValue()) !== "ja") {
    return { ok: false, fehler: "ET/Diagnose ist nicht Ja." };
  }
  cell.setValue("Nein");
  SpreadsheetApp.flush();
  try {
    syncInputRowToNachbestellung(sheet, r);
  } catch (e1) {}
  cell.setValue("Ja");
  SpreadsheetApp.flush();
  try {
    syncInputRowToNachbestellung(sheet, r);
  } catch (e2) {}
  try {
    enqueueInputEdit(sheet, sheet.getName(), r);
  } catch (e3) {}
  done.push(eid);
  uxNbNachtragenDoneWrite_(done);
  try {
    var cache = uxNbHubCacheRead_();
    if (cache) {
      cache.nachtragenDone = done;
      uxNbHubCacheWrite_(cache);
    }
  } catch (e4) {}
  return { ok: true, entryId: eid };
}

function uxNbCountFilledDataRows_(sheet, dataStartRow) {
  if (!sheet) return 0;
  var start = dataStartRow > 0 ? dataStartRow : 2;
  var end = sheet.getLastRow();
  if (end < start) return 0;
  var lastCol = Math.max(1, sheet.getLastColumn());
  var values = sheet.getRange(start, 1, end - start + 1, lastCol).getValues();
  var count = 0;
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    for (var c = 0; c < row.length; c++) {
      if (String(row[c] || "").trim()) {
        count++;
        break;
      }
    }
  }
  return count;
}

function uxNbHubDashboardStats() {
  var stats = {
    ok: true,
    delBufferCount: 0,
    sheetCounts: {
      "Input Mechanik": 0,
      "Input Q-Check": 0,
      "Input Lack": 0,
      "Input Exit": 0,
      Nachbestellung: 0
    },
    triggers: [],
    userEmail: uxNbDelBufferUser_()
  };
  try {
    stats.delBufferCount = uxNbDelBufferReadRaw_().length;
  } catch (e) {}
  try {
    var ss = getMainSS();
    if (ss) {
      var me = ss.getSheetByName("Input Mechanik");
      if (me) {
        stats.sheetCounts["Input Mechanik"] = uxNbCountFilledDataRows_(
          me,
          inputSheetArchiveDataStart("Input Mechanik")
        );
      }
      var qc = ss.getSheetByName("Input Q-Check");
      if (qc) {
        stats.sheetCounts["Input Q-Check"] = uxNbCountFilledDataRows_(
          qc,
          inputSheetArchiveDataStart("Input Q-Check")
        );
      }
      var la = ss.getSheetByName("Input Lack");
      if (la) {
        stats.sheetCounts["Input Lack"] = uxNbCountFilledDataRows_(
          la,
          inputSheetArchiveDataStart("Input Lack")
        );
      }
      var ex = ss.getSheetByName("Input Exit");
      if (ex) {
        stats.sheetCounts["Input Exit"] = uxNbCountFilledDataRows_(
          ex,
          inputSheetArchiveDataStart("Input Exit")
        );
      }
      var nb = ss.getSheetByName("Nachbestellung");
      if (nb) {
        var nLayout = getNachbestellungLayout(nb);
        stats.sheetCounts.Nachbestellung = uxNbCountFilledDataRows_(
          nb,
          nLayout.dataStartRow
        );
      }
    }
  } catch (e2) {}
  try {
    var watch = {
      cleanupNachbestellungStatusDate: false,
      archiveFertiggestelltRows: false,
      processQueue: false
    };
    var triggers = ScriptApp.getProjectTriggers();
    for (var t = 0; t < triggers.length; t++) {
      var fn = triggers[t].getHandlerFunction();
      if (Object.prototype.hasOwnProperty.call(watch, fn)) watch[fn] = true;
    }
    var names = [];
    if (watch.cleanupNachbestellungStatusDate) names.push("cleanupNachbestellungStatusDate");
    if (watch.archiveFertiggestelltRows) names.push("archiveFertiggestelltRows");
    if (watch.processQueue) names.push("processQueue");
    stats.triggers = names;
  } catch (e4) {}
  return stats;
}

function uxNbBuildNachbestellungEntryIdSet_(nbSheet) {
  var set = {};
  if (!nbSheet) return set;
  var layout = getNachbestellungLayout(nbSheet);
  var nbLast = nbSheet.getLastRow();
  var trim = layout.dataEndTrimBottomRows != null ? layout.dataEndTrimBottomRows : 1;
  var nbEnd = nbLast > trim ? nbLast - trim : nbLast;
  if (nbEnd < layout.dataStartRow) return set;
  var eidCol = layout.cols.entryId;
  var eids = nbSheet.getRange(layout.dataStartRow, eidCol, nbEnd - layout.dataStartRow + 1, 1).getValues();
  for (var i = 0; i < eids.length; i++) {
    var eid = String(eids[i][0] || "").trim();
    if (eid) set[eid] = true;
  }
  return set;
}

function uxNbNormEtDiagnose_(v) {
  return String(v == null ? "" : v).trim().toLowerCase();
}

function uxNbSyncMissingInputsListe() {
  try {
    var ss = getMainSS();
    if (!ss) return { ok: false, fehler: "Hauptmappe nicht erreichbar.", missing: [], count: 0 };
    var nbSheet = ss.getSheetByName("Nachbestellung");
    if (!nbSheet) return { ok: false, fehler: "Tabellenblatt Nachbestellung fehlt.", missing: [], count: 0 };
    var nbIds = uxNbBuildNachbestellungEntryIdSet_(nbSheet);
    var missing = [];
    for (var k = 0; k < INPUT_SHEET_NAMES.length; k++) {
      var sheetName = INPUT_SHEET_NAMES[k];
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      var start = inputSheetArchiveDataStart(sheetName);
      var end = inputSheetArchiveDataEnd(sheet, start);
      if (end == null || end < start) continue;
      var eidCol = getEntryIdCol(sheet);
      var dataBlock = sheet.getRange(start, 1, end, 6).getValues();
      var eidBlock = sheet.getRange(start, eidCol, end, eidCol).getValues();
      for (var r = 0; r < dataBlock.length; r++) {
        var rowData = dataBlock[r];
        var stockId = String(rowData[1] || "").trim();
        if (!stockId) continue;
        if (uxNbNormEtDiagnose_(rowData[2]) !== "ja") continue;
        if (!isInputRowComplete(rowData)) continue;
        var entryId = String(eidBlock[r][0] || "").trim();
        if (!entryId) {
          missing.push({
            sheetName: sheetName,
            row: start + r,
            stockId: stockId,
            entryId: "",
            herkunft: INPUT_SHEETS[sheetName] || ""
          });
          continue;
        }
        if (!nbIds[entryId]) {
          missing.push({
            sheetName: sheetName,
            row: start + r,
            stockId: stockId,
            entryId: entryId,
            herkunft: INPUT_SHEETS[sheetName] || ""
          });
        }
      }
    }
    missing.sort(function (a, b) {
      if (a.sheetName !== b.sheetName) return a.sheetName < b.sheetName ? -1 : 1;
      return a.row - b.row;
    });
    return { ok: true, missing: missing, count: missing.length };
  } catch (e) {
    return { ok: false, fehler: String(e), missing: [], count: 0 };
  }
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
