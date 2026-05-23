var UX_NB_DELBUF_KEY = "ux_nb_delbuf_v1";
var UX_NB_DELBUF_MAX = 100;
var UX_NB_HUB_CACHE_KEY = "ux_nb_hub_cache_v5";
var UX_NB_ALERT_STATUS_EMPTY_DAYS = 7;
var UX_NB_ALERT_DASH_RED_DAYS = 10;
var UX_NB_BULK_CHUNK = 8;
var UX_NB_NACHT_KEY = "ux_nb_nachtragen_done_v1";
var UX_NB_EVENTQUEUE_SHEET = "EventQueue";
var UX_NB_EVENTQUEUE_KEEP = 500;
var UX_NB_DASH_INPUT_SHEETS = ["Input Mechanik", "Input Q-Check", "Input Lack"];
var UX_NB_DASH_NACHT_KEY = "ux_nb_dash_nachtragen_done_v1";

function uxBaseMenueEinrichten() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Nachbestellung")
      .addItem("Control Center...", "uxBaseDialogControlCenterOeffnen")
      .addToUi();
  } catch (e) {
    try {
      Logger.log("uxBaseMenueEinrichten: " + e);
    } catch (e2) {}
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function uxAutoheroLogoDataUrl_() {
  if (typeof UX_AUTOHERO_LOGO_B64 === "undefined") return "";
  return "data:image/png;base64," + UX_AUTOHERO_LOGO_B64;
}

function uxBaseDialogControlCenterOeffnen(initialView) {
  var t = HtmlService.createTemplateFromFile("UXMenu");
  t.initialView = initialView || "home";
  t.autoheroLogoDataUrl = uxAutoheroLogoDataUrl_();
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

function uxNbDashNachtragenDoneRead_() {
  try {
    var raw = PropertiesService.getUserProperties().getProperty(UX_NB_DASH_NACHT_KEY);
    var arr = JSON.parse(raw || "[]");
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function uxNbDashNachtragenDoneWrite_(arr) {
  PropertiesService.getUserProperties().setProperty(UX_NB_DASH_NACHT_KEY, JSON.stringify(arr || []));
}

function uxNbDashNachtragenDoneCleanup_(missing, done) {
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
  try {
    if (typeof ensureDashboardStatusSyncTrigger_ === "function") ensureDashboardStatusSyncTrigger_();
  } catch (e0) {}
  var stats = uxNbHubDashboardStats();
  var del = uxNbDelBufferRead_();
  var sync = uxNbSyncMissingInputsListe();
  var missing = sync.ok ? sync.missing || [] : [];
  var nachtragenDone = uxNbNachtragenDoneCleanup_(missing, uxNbNachtragenDoneRead_());
  uxNbNachtragenDoneWrite_(nachtragenDone);
  var dash = uxNbDashMissingListe();
  var dashMissing = dash.ok ? dash.missing || [] : [];
  var dashNachtragenDone = uxNbDashNachtragenDoneCleanup_(dashMissing, uxNbDashNachtragenDoneRead_());
  uxNbDashNachtragenDoneWrite_(dashNachtragenDone);
    var dashGaps = uxNbDashStatusGapListe();
    var dashStatusGaps = dashGaps.ok ? dashGaps.gaps || [] : [];
    var integrity = uxNbIntegritySummary_();
    return {
      ok: true,
      loadedAt: new Date().toISOString(),
      stats: stats,
      delBuffer: del.ok ? del.arr || [] : [],
      syncMissing: missing,
      syncCount: sync.ok ? sync.count || 0 : 0,
      nachtragenDone: nachtragenDone,
      dashMissing: dashMissing,
      dashSyncCount: dash.ok ? dash.count || 0 : 0,
      dashNachtragenDone: dashNachtragenDone,
      dashStatusGaps: dashStatusGaps,
      dashStatusGapCount: dashGaps.ok ? dashGaps.count || 0 : 0,
      integritySummary: integrity,
      integrityCount: integrity.ok ? integrity.total || 0 : 0
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
  var values = sheet.getRange(start, 1, end, lastCol).getValues();
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

function uxNbEventQueuePlan_(queue, keepMax) {
  var plan = {
    ok: false,
    sheetName: UX_NB_EVENTQUEUE_SHEET,
    keepMax: keepMax,
    dataRows: 0,
    kept: 0,
    removable: 0,
    firstKeepRow: 2
  };
  if (!queue) return plan;
  var last = queue.getLastRow();
  if (last < 2) {
    plan.ok = true;
    return plan;
  }
  var dataRows = last - 1;
  plan.dataRows = dataRows;
  if (dataRows <= keepMax) {
    plan.kept = dataRows;
    plan.removable = 0;
    plan.firstKeepRow = 2;
    plan.ok = true;
    return plan;
  }
  plan.kept = keepMax;
  plan.removable = dataRows - keepMax;
  plan.firstKeepRow = last - keepMax + 1;
  plan.ok = true;
  return plan;
}

function uxNbEventQueueInfo_() {
  var info = {
    ok: false,
    sheetName: UX_NB_EVENTQUEUE_SHEET,
    keepMax: UX_NB_EVENTQUEUE_KEEP,
    dataRows: 0,
    kept: 0,
    removable: 0
  };
  try {
    var ss = getMainSS();
    if (!ss) {
      info.fehler = "Hauptmappe nicht erreichbar.";
      return info;
    }
    var queue = ss.getSheetByName(UX_NB_EVENTQUEUE_SHEET);
    if (!queue) {
      info.fehler = "Tabellenblatt EventQueue fehlt.";
      return info;
    }
    var plan = uxNbEventQueuePlan_(queue, UX_NB_EVENTQUEUE_KEEP);
    if (!plan.ok) {
      info.fehler = plan.fehler || "EventQueue-Plan fehlgeschlagen.";
      return info;
    }
    info.ok = true;
    info.dataRows = plan.dataRows;
    info.kept = plan.kept;
    info.removable = plan.removable;
  } catch (e) {
    info.fehler = String(e);
  }
  return info;
}

function uxNbEventQueueInfo() {
  return uxNbEventQueueInfo_();
}

function uxNbEventQueueAufraeumen() {
  var lock = LockService.getScriptLock();
  var got = false;
  try {
    got = lock.tryLock(30000);
    if (!got) return { ok: false, fehler: "Konnte keine Sperre bekommen. Bitte erneut versuchen." };
    var ss = getMainSS();
    if (!ss) return { ok: false, fehler: "Hauptmappe nicht erreichbar." };
    var queue = ss.getSheetByName(UX_NB_EVENTQUEUE_SHEET);
    if (!queue) return { ok: false, fehler: "Tabellenblatt EventQueue fehlt." };
    var plan = uxNbEventQueuePlan_(queue, UX_NB_EVENTQUEUE_KEEP);
    if (!plan.ok) return { ok: false, fehler: plan.fehler || "Planung fehlgeschlagen." };
    if (plan.removable <= 0) {
      return {
        ok: true,
        removed: 0,
        kept: plan.kept || plan.dataRows,
        dataRows: plan.dataRows,
        keepMax: UX_NB_EVENTQUEUE_KEEP
      };
    }
    queue.deleteRows(2, plan.removable);
    var after = uxNbEventQueueInfo_();
    try {
      var cache = uxNbHubCacheRead_();
      if (cache && cache.stats) {
        cache.stats.eventQueue = after;
        uxNbHubCacheWrite_(cache);
      }
    } catch (e2) {}
    return {
      ok: true,
      removed: plan.removable,
      kept: after.kept || Math.min(UX_NB_EVENTQUEUE_KEEP, plan.dataRows - plan.removable),
      dataRows: plan.dataRows,
      keepMax: UX_NB_EVENTQUEUE_KEEP
    };
  } catch (e) {
    return { ok: false, fehler: String(e) };
  } finally {
    if (got) lock.releaseLock();
  }
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
      processQueue: false,
      syncDashboardStatusSyncTrigger: false
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
    if (watch.syncDashboardStatusSyncTrigger) names.push("syncDashboardStatusSyncTrigger");
    stats.triggers = names;
  } catch (e4) {}
  try {
    stats.alerts = uxNbHubAlertsSammeln_();
  } catch (e5a) {
    stats.alerts = null;
  }
  try {
    stats.eventQueue = uxNbEventQueueInfo_();
  } catch (e5) {
    stats.eventQueue = {
      ok: false,
      sheetName: UX_NB_EVENTQUEUE_SHEET,
      keepMax: UX_NB_EVENTQUEUE_KEEP,
      dataRows: 0,
      kept: 0,
      removable: 0,
      fehler: String(e5)
    };
  }
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

function uxNbBuildDashboardEntryIdSet_(dashboard) {
  var set = {};
  if (!dashboard) return set;
  var layout = getDashboardLayout(dashboard);
  var last = dashboard.getLastRow();
  if (last < layout.dataStartRow) return set;
  var eids = dashboard.getRange(
    layout.dataStartRow,
    layout.cols.entryId,
    last,
    layout.cols.entryId
  ).getValues();
  for (var i = 0; i < eids.length; i++) {
    var eid = String(eids[i][0] || "").trim();
    if (eid) set[eid] = true;
  }
  return set;
}

function uxNbSameDashStatusDate_(a, b) {
  if (a === "" || a === null || a === undefined) {
    return b === "" || b === null || b === undefined;
  }
  if (b === "" || b === null || b === undefined) return false;
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  return String(a) === String(b);
}

function uxNbDashStatusGapListe() {
  try {
    var ss = getMainSS();
    if (!ss) return { ok: false, fehler: "Hauptmappe nicht erreichbar.", gaps: [], count: 0 };
    var dashboard = ss.getSheetByName("Dashboard");
    var nbSheet = ss.getSheetByName("Nachbestellung");
    if (!dashboard || !nbSheet) {
      return { ok: false, fehler: "Dashboard oder Nachbestellung fehlt.", gaps: [], count: 0 };
    }
    var dLayout = getDashboardLayout(dashboard);
    var last = dashboard.getLastRow();
    if (last < dLayout.dataStartRow) return { ok: true, gaps: [], count: 0 };
    var numRows = last - dLayout.dataStartRow + 1;
    var block = dashboard.getRange(dLayout.dataStartRow, 1, numRows, dLayout.lastCol).getValues();
    var exitIds = buildInputExitEntryIdSet_(ss);
    var nbByEid = buildNachbestellungStatusByEntryIdMap_(nbSheet);
    var gaps = [];
    for (var i = 0; i < block.length; i++) {
      var row = block[i];
      var entryId = String(row[dLayout.cols.entryId - 1] || "").trim();
      var stockId = String(row[dLayout.cols.stockId - 1] || "").trim();
      if (!entryId || !stockId) continue;
      if (exitIds[entryId]) continue;
      var nb = nbByEid[entryId];
      if (!nb) continue;
      var nbHas = isCellFilled(nb.i) || isCellFilled(nb.k) || isCellFilled(nb.l);
      if (!nbHas) continue;
      var curI = row[dLayout.cols.costGate - 1];
      var curK = row[dLayout.cols.status - 1];
      var curL = row[dLayout.cols.statusAenderung - 1];
      if (
        String(curI) === String(nb.i) &&
        String(curK) === String(nb.k) &&
        uxNbSameDashStatusDate_(curL, nb.l)
      ) continue;
      gaps.push({
        dashRow: dLayout.dataStartRow + i,
        stockId: stockId,
        entryId: entryId,
        herkunft: String(row[dLayout.cols.herkunft - 1] || "").trim(),
        nbCostGate: String(nb.i || ""),
        nbStatus: String(nb.k || "")
      });
    }
    return { ok: true, gaps: gaps, count: gaps.length };
  } catch (e) {
    return { ok: false, fehler: String(e), gaps: [], count: 0 };
  }
}

function uxNbDashPushRow_(sheet, row) {
  if (!sheet || !row || row < 2) return { ok: false, fehler: "Ungültige Zeile." };
  var name = sheet.getName();
  if (UX_NB_DASH_INPUT_SHEETS.indexOf(name) === -1) {
    return { ok: false, fehler: "Nur Mechanik, Q-Check und Lack." };
  }
  var ss = getMainSS();
  if (!ss) return { ok: false, fehler: "Hauptmappe nicht erreichbar." };
  var dashboard = ss.getSheetByName("Dashboard");
  var nbSheet = ss.getSheetByName("Nachbestellung");
  if (!dashboard || !nbSheet) return { ok: false, fehler: "Dashboard oder Nachbestellung fehlt." };
  var rowData = sheet.getRange(row, 1, 1, 6).getValues()[0];
  var key = String(rowData[1] || "").trim();
  if (!key) return { ok: false, fehler: "Keine Stock-ID." };
  if (uxNbNormEtDiagnose_(rowData[2]) !== "ja") return { ok: false, fehler: "ET/Diagnose ist nicht Ja." };
  if (!isInputRowComplete(rowData)) return { ok: false, fehler: "Zeile ist nicht komplett." };
  var entryId = String(sheet.getRange(row, getEntryIdCol(sheet)).getValue() || "").trim();
  if (!entryId) return { ok: false, fehler: "Keine Entry-ID." };
  if (isEntryIdFromInputExit(entryId)) return { ok: false, fehler: "Exit-Zeile gehört nicht ins Dashboard." };
  if (isInputDatumBlank(rowData[0])) {
    var now = new Date();
    sheet.getRange(row, 1).setValue(now);
    rowData[0] = now;
  }
  var code = INPUT_SHEETS[name] || "";
  var tagesMap = getTageslisteMap();
  processRow(rowData, key, dashboard, nbSheet, tagesMap, code, false, entryId);
  applyNachbestellungState(nbSheet, key, rowData, entryId);
  return { ok: true, entryId: entryId, stockId: key };
}

function uxNbDashMissingListe() {
  try {
    var ss = getMainSS();
    if (!ss) return { ok: false, fehler: "Hauptmappe nicht erreichbar.", missing: [], count: 0 };
    var dashboard = ss.getSheetByName("Dashboard");
    if (!dashboard) return { ok: false, fehler: "Tabellenblatt Dashboard fehlt.", missing: [], count: 0 };
    var dashIds = uxNbBuildDashboardEntryIdSet_(dashboard);
    var missing = [];
    for (var k = 0; k < UX_NB_DASH_INPUT_SHEETS.length; k++) {
      var sheetName = UX_NB_DASH_INPUT_SHEETS[k];
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      var start = inputSheetArchiveDataStart(sheetName);
      var end = sheet.getLastRow();
      if (end < start) continue;
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
        if (isEntryIdFromInputExit(entryId)) continue;
        if (!dashIds[entryId]) {
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

function uxNbDashNachtragen(sheetName, row, entryId) {
  var eid = String(entryId || "").trim();
  if (!eid) return { ok: false, fehler: "Keine Entry-ID." };
  var done = uxNbDashNachtragenDoneRead_();
  if (done.indexOf(eid) !== -1) return { ok: false, fehler: "Bereits angestoßen." };
  var ss = getMainSS();
  if (!ss) return { ok: false, fehler: "Hauptmappe nicht erreichbar." };
  var sheet = ss.getSheetByName(String(sheetName || ""));
  if (!sheet) return { ok: false, fehler: "Input-Blatt nicht gefunden." };
  var r = parseInt(row, 10);
  if (!r || r < 2) return { ok: false, fehler: "Ungültige Zeile." };
  var push = uxNbDashPushRow_(sheet, r);
  if (!push.ok) return push;
  try {
    enqueueInputEdit(sheet, sheet.getName(), r);
  } catch (e1) {}
  done.push(eid);
  uxNbDashNachtragenDoneWrite_(done);
  try {
    var cache = uxNbHubCacheRead_();
    if (cache) {
      cache.dashNachtragenDone = done;
      uxNbHubCacheWrite_(cache);
    }
  } catch (e2) {}
  return { ok: true, entryId: eid };
}

function uxNbDashSyncStatusNow() {
  if (typeof syncDashboardStatusFromNachbestellung !== "function") {
    return { ok: false, fehler: "Sync-Funktion fehlt im Script." };
  }
  return syncDashboardStatusFromNachbestellung(
    typeof DASH_SYNC_LOCK_WAIT_MS !== "undefined" ? DASH_SYNC_LOCK_WAIT_MS : 30000
  );
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

function uxNbIsFertigArchiveStatus_(val) {
  var s = String(val == null ? "" : val).trim().toLowerCase();
  return s === "fertiggestellt" || s === "b2a1";
}

function uxNbCountFertigRowsOnSheet_(sheet, dataStart, dataEnd, statusCol) {
  if (!sheet || dataEnd < dataStart) return 0;
  var numRows = dataEnd - dataStart + 1;
  var block = sheet.getRange(dataStart, statusCol, numRows, 1).getValues();
  var disp = sheet.getRange(dataStart, statusCol, numRows, 1).getDisplayValues();
  var n = 0;
  for (var i = 0; i < block.length; i++) {
    if (uxNbIsFertigArchiveStatus_(block[i][0]) || uxNbIsFertigArchiveStatus_(disp[i][0])) n++;
  }
  return n;
}

function uxNbArchiveReadyCount_() {
  var out = { nb: 0, exit: 0, total: 0 };
  try {
    var ss = getMainSS();
    if (!ss) return out;
    if (typeof nachbestellungFertigBounds_ === "function") {
      var nb = ss.getSheetByName("Nachbestellung");
      if (nb) {
        var bb = nachbestellungFertigBounds_(nb);
        if (bb) out.nb = uxNbCountFertigRowsOnSheet_(nb, bb.dataStart, bb.dataEnd, bb.statusCol);
      }
    }
    if (typeof inputExitFertigBounds_ === "function") {
      var ex = ss.getSheetByName("Input Exit");
      if (ex) {
        var eb = inputExitFertigBounds_(ex);
        if (eb) out.exit = uxNbCountFertigRowsOnSheet_(ex, eb.dataStart, eb.dataEnd, eb.statusCol);
      }
    }
  } catch (e) {}
  out.total = out.nb + out.exit;
  return out;
}

function uxNbDashRedDaysCount_(minDays) {
  minDays = minDays != null ? minDays : UX_NB_ALERT_DASH_RED_DAYS;
  try {
    var ss = getMainSS();
    if (!ss) return 0;
    var dashboard = ss.getSheetByName("Dashboard");
    if (!dashboard) return 0;
    var layout = getDashboardLayout(dashboard);
    var last = dashboard.getLastRow();
    if (last < layout.dataStartRow) return 0;
    var end = dashboardDataEndRow(dashboard, layout);
    var numRows = end - layout.dataStartRow + 1;
    var col = layout.cols.tageSeitLetztem;
    var vals = dashboard.getRange(layout.dataStartRow, col, numRows, 1).getValues();
    var n = 0;
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i][0];
      if (v === "" || v == null) continue;
      var d = Number(v);
      if (!isNaN(d) && d > minDays) n++;
    }
    return n;
  } catch (e) {
    return 0;
  }
}

function uxNbNbStatusEmptyOldCount_(minDays) {
  minDays = minDays != null ? minDays : UX_NB_ALERT_STATUS_EMPTY_DAYS;
  try {
    var ss = getMainSS();
    if (!ss) return 0;
    var nbSheet = ss.getSheetByName("Nachbestellung");
    if (!nbSheet) return 0;
    var layout = getNachbestellungLayout(nbSheet);
    var last = nbSheet.getLastRow();
    var trim = layout.dataEndTrimBottomRows != null ? layout.dataEndTrimBottomRows : 1;
    var dataEnd = last > trim ? last - trim : last;
    if (dataEnd < layout.dataStartRow) return 0;
    var numRows = dataEnd - layout.dataStartRow + 1;
    var width = Math.max(layout.lastCol, layout.cols.status, layout.cols.datum || 1);
    var block = nbSheet.getRange(layout.dataStartRow, 1, numRows, width).getValues();
    var stIdx = layout.cols.status - 1;
    var today = new Date();
    var n = 0;
    for (var i = 0; i < block.length; i++) {
      var row = block[i];
      var st = row[stIdx];
      if (isCellFilled(st)) continue;
      if (!String(row[layout.cols.stockId - 1] || "").trim()) continue;
      if (!String(row[layout.cols.entryId - 1] || "").trim()) continue;
      var dt = row[0];
      if (!dt) continue;
      var age = daysBetween(today, new Date(dt));
      if (age !== "" && age !== null && Number(age) >= minDays) n++;
    }
    return n;
  } catch (e) {
    return 0;
  }
}

function uxNbHubAlertsSammeln_() {
  return {
    archiveReady: uxNbArchiveReadyCount_(),
    dashRedDays: uxNbDashRedDaysCount_(UX_NB_ALERT_DASH_RED_DAYS),
    nbStatusEmptyOld: uxNbNbStatusEmptyOldCount_(UX_NB_ALERT_STATUS_EMPTY_DAYS),
    statusEmptyDays: UX_NB_ALERT_STATUS_EMPTY_DAYS,
    dashRedMin: UX_NB_ALERT_DASH_RED_DAYS
  };
}

function uxNbFindDuplicateEids_(eidList) {
  var seen = {};
  var dups = [];
  for (var i = 0; i < eidList.length; i++) {
    var eid = eidList[i].eid;
    if (!eid) continue;
    if (!seen[eid]) seen[eid] = [];
    seen[eid].push(eidList[i].row);
  }
  for (var k in seen) {
    if (seen[k].length > 1) dups.push({ entryId: k, rows: seen[k] });
  }
  return dups;
}

function uxNbIntegritySummary_() {
  var full = uxNbIntegrityCheck(true);
  return {
    ok: full.ok,
    total: full.total || 0,
    fehler: full.fehler || ""
  };
}

function uxNbIntegrityCheck(summaryOnly) {
  var maxSamples = summaryOnly ? 0 : 4;
  var issues = [];
  var total = 0;
  try {
    var ss = getMainSS();
    if (!ss) return { ok: false, fehler: "Hauptmappe nicht erreichbar.", issues: [], total: 0 };
    var nbSheet = ss.getSheetByName("Nachbestellung");
    var dashboard = ss.getSheetByName("Dashboard");
    var exitIds = buildInputExitEntryIdSet_(ss);
    if (nbSheet) {
      var nbLo = getNachbestellungLayout(nbSheet);
      var nbLast = nbSheet.getLastRow();
      var nbTrim = nbLo.dataEndTrimBottomRows != null ? nbLo.dataEndTrimBottomRows : 1;
      var nbEnd = nbLast > nbTrim ? nbLast - nbTrim : nbLast;
      if (nbEnd >= nbLo.dataStartRow) {
        var nbNum = nbEnd - nbLo.dataStartRow + 1;
        var nbEids = nbSheet.getRange(nbLo.dataStartRow, nbLo.cols.entryId, nbNum, 1).getValues();
        var nbList = [];
        for (var ni = 0; ni < nbEids.length; ni++) {
          nbList.push({ eid: String(nbEids[ni][0] || "").trim(), row: nbLo.dataStartRow + ni });
        }
        var nbDups = uxNbFindDuplicateEids_(nbList);
        if (nbDups.length) {
          total += nbDups.length;
          issues.push({
            id: "dup_nb",
            label: "Doppelte Entry-ID in Nachbestellung",
            hint: "Gleiche UUID mehrfach in NB – das verwirrt alles.",
            count: nbDups.length,
            samples: maxSamples ? nbDups.slice(0, maxSamples) : []
          });
        }
      }
    }
    if (dashboard) {
      var dLo = getDashboardLayout(dashboard);
      var dLast = dashboard.getLastRow();
      if (dLast >= dLo.dataStartRow) {
        var dNum = dLast - dLo.dataStartRow + 1;
        var dBlock = dashboard.getRange(dLo.dataStartRow, 1, dNum, dLo.lastCol).getValues();
        var dList = [];
        var exitGhost = 0;
        var exitSamples = [];
        var noNb = 0;
        var noNbSamples = [];
        var nbMap = nbSheet ? buildNachbestellungStatusByEntryIdMap_(nbSheet) : {};
        for (var di = 0; di < dBlock.length; di++) {
          var dr = dBlock[di];
          var deid = String(dr[dLo.cols.entryId - 1] || "").trim();
          var dstock = String(dr[dLo.cols.stockId - 1] || "").trim();
          if (!dstock) continue;
          if (deid) dList.push({ eid: deid, row: dLo.dataStartRow + di });
          var herk = String(dr[dLo.cols.herkunft - 1] || "").trim().toUpperCase();
          if (herk === "EX" || (deid && exitIds[deid])) {
            exitGhost++;
            if (!summaryOnly && exitSamples.length < maxSamples) {
              exitSamples.push({ dashRow: dLo.dataStartRow + di, stockId: dstock, entryId: deid });
            }
          }
          if (deid && !exitIds[deid] && !nbMap[deid]) {
            noNb++;
            if (!summaryOnly) {
              noNbSamples.push({ dashRow: dLo.dataStartRow + di, stockId: dstock, entryId: deid });
            }
          }
        }
        var dashDups = uxNbFindDuplicateEids_(dList);
        if (dashDups.length) {
          total += dashDups.length;
          issues.push({
            id: "dup_dash",
            label: "Doppelte Entry-ID im Dashboard",
            hint: "Zwei Dashboard-Zeilen mit gleicher UUID.",
            count: dashDups.length,
            samples: maxSamples ? dashDups.slice(0, maxSamples) : []
          });
        }
        if (exitGhost) {
          total += exitGhost;
          issues.push({
            id: "dash_exit",
            label: "Exit-Zeilen hängen noch im Dashboard",
            hint: "EX oder Input-Exit Entry-ID – sollten eigentlich raus.",
            count: exitGhost,
            samples: maxSamples ? exitSamples.slice(0, maxSamples) : exitSamples
          });
        }
        if (noNb) {
          total += noNb;
          issues.push({
            id: "dash_no_nb",
            label: "Dashboard ohne NB-Treffer",
            hint: "Im Dashboard aber Entry-ID fehlt in Nachbestellung komplett. Vielleicht ist es bereits im Archiv – wenn ja, aus Dashboard löschen.",
            count: noNb,
            samples: summaryOnly ? [] : noNbSamples
          });
        }
      }
    }
    var noEid = 0;
    var noEidSamples = [];
    for (var si = 0; si < INPUT_SHEET_NAMES.length; si++) {
      var sname = INPUT_SHEET_NAMES[si];
      var sh = ss.getSheetByName(sname);
      if (!sh) continue;
      var start = inputSheetArchiveDataStart(sname);
      var end = inputSheetArchiveDataEnd(sh, start);
      if (end == null || end < start) continue;
      var eidCol = getEntryIdCol(sh);
      var dataBlock = sh.getRange(start, 1, end, 6).getValues();
      var eidBlock = sh.getRange(start, eidCol, end, eidCol).getValues();
      for (var ri = 0; ri < dataBlock.length; ri++) {
        if (!String(dataBlock[ri][1] || "").trim()) continue;
        if (uxNbNormEtDiagnose_(dataBlock[ri][2]) !== "ja") continue;
        if (!isInputRowComplete(dataBlock[ri])) continue;
        if (String(eidBlock[ri][0] || "").trim()) continue;
        noEid++;
        if (!summaryOnly && noEidSamples.length < maxSamples) {
          noEidSamples.push({ sheetName: sname, row: start + ri, stockId: String(dataBlock[ri][1] || "").trim() });
        }
      }
    }
    if (noEid) {
      total += noEid;
      issues.push({
        id: "input_no_eid",
        label: "Input komplett aber ohne Entry-ID",
        hint: "ET/Diagnose Ja, Zeile voll – UUID fehlt noch.",
        count: noEid,
        samples: maxSamples ? noEidSamples.slice(0, maxSamples) : noEidSamples
      });
    }
    try {
      var queue = ss.getSheetByName(UX_NB_EVENTQUEUE_SHEET);
      if (queue && queue.getLastRow() >= 2) {
        var ql = queue.getLastRow();
        var qNum = ql - 1;
        var qData = queue.getRange(2, 1, qNum, 7).getValues();
        var stuck = 0;
        var stuckSamples = [];
        var now = Date.now();
        for (var qi = 0; qi < qData.length; qi++) {
          var act = String(qData[qi][1] || "").trim();
          if (act === "_LOG") continue;
          var st = String(qData[qi][6] || "").trim().toUpperCase();
          if (st === "DONE" || st === "SKIPPED" || st === "CANCELLED") continue;
          var ts = qData[qi][0];
          var age = ts ? now - new Date(ts).getTime() : 999999999;
          if (age < 3600000) continue;
          stuck++;
          if (!summaryOnly && stuckSamples.length < maxSamples) {
            stuckSamples.push({
              row: qi + 2,
              action: String(qData[qi][1] || ""),
              status: st,
              entryId: String(qData[qi][3] || "")
            });
          }
        }
        if (stuck) {
          total += stuck;
          issues.push({
            id: "queue_stuck",
            label: "EventQueue hängt (>1h)",
            hint: "Pending/Error seit über einer Stunde – evtl. Queue prüfen.",
            count: stuck,
            samples: maxSamples ? stuckSamples.slice(0, maxSamples) : stuckSamples
          });
        }
      }
    } catch (eq) {}
    return { ok: true, issues: issues, total: total };
  } catch (err) {
    return { ok: false, fehler: String(err), issues: [], total: 0 };
  }
}

function uxNbDashBulkEligible_(missing, doneSet) {
  var eligible = [];
  var skipNoEid = 0;
  var skipDone = 0;
  for (var i = 0; i < (missing || []).length; i++) {
    var it = missing[i];
    var eid = String(it.entryId || "").trim();
    if (!eid) {
      skipNoEid++;
      continue;
    }
    if (doneSet[eid]) {
      skipDone++;
      continue;
    }
    eligible.push(it);
  }
  return { eligible: eligible, skipNoEid: skipNoEid, skipDone: skipDone };
}

function uxNbDashNachtragenBulk(maxCount) {
  maxCount = maxCount != null ? maxCount : UX_NB_BULK_CHUNK;
  var list = uxNbDashMissingListe();
  if (!list.ok) return list;
  var done = uxNbDashNachtragenDoneRead_();
  var doneSet = {};
  for (var d = 0; d < done.length; d++) doneSet[done[d]] = true;
  var plan = uxNbDashBulkEligible_(list.missing, doneSet);
  var pushed = 0;
  var failed = 0;
  var errors = [];
  for (var i = 0; i < plan.eligible.length && pushed + failed < maxCount; i++) {
    var it = plan.eligible[i];
    var res = uxNbDashNachtragen(it.sheetName, it.row, it.entryId);
    if (res.ok) pushed++;
    else {
      failed++;
      if (errors.length < 3) errors.push(res.fehler || "?");
    }
  }
  var left = uxNbDashMissingListe();
  return {
    ok: true,
    pushed: pushed,
    failed: failed,
    skippedNoEid: plan.skipNoEid,
    skippedDone: plan.skipDone,
    remaining: left.ok ? left.count || 0 : 0,
    errors: errors
  };
}

function uxNbSyncNachtragenBulk(maxCount) {
  maxCount = maxCount != null ? maxCount : UX_NB_BULK_CHUNK;
  var list = uxNbSyncMissingInputsListe();
  if (!list.ok) return list;
  var done = uxNbNachtragenDoneRead_();
  var doneSet = {};
  for (var d = 0; d < done.length; d++) doneSet[done[d]] = true;
  var plan = uxNbDashBulkEligible_(list.missing, doneSet);
  var triggered = 0;
  var failed = 0;
  var errors = [];
  for (var i = 0; i < plan.eligible.length && triggered + failed < maxCount; i++) {
    var it = plan.eligible[i];
    var res = uxNbSyncNachtragen(it.sheetName, it.row, it.entryId);
    if (res.ok) triggered++;
    else {
      failed++;
      if (errors.length < 3) errors.push(res.fehler || "?");
    }
  }
  var left = uxNbSyncMissingInputsListe();
  return {
    ok: true,
    triggered: triggered,
    failed: failed,
    skippedNoEid: plan.skipNoEid,
    skippedDone: plan.skipDone,
    remaining: left.ok ? left.count || 0 : 0,
    errors: errors
  };
}

function uxNbDashboardDayDiffsNow() {
  if (typeof runUpdateDashboardDayDiffsForUi !== "function") {
    return { ok: false, fehler: "Day-Diff Funktion fehlt im Script." };
  }
  return runUpdateDashboardDayDiffsForUi();
}

function uxNbIntegrityCheckFull() {
  return uxNbIntegrityCheck(false);
}

function uxNbCleanupDashboardExitNow() {
  if (typeof runCleanupDashboardOfInputExitForUi !== "function") {
    return { ok: false, fehler: "Cleanup Funktion fehlt im Script." };
  }
  return runCleanupDashboardOfInputExitForUi();
}

function uxNbDashZeileInsArchiv(dashRow) {
  try {
    var ss = getMainSS();
    if (!ss) return { ok: false, fehler: "Hauptmappe nicht erreichbar." };
    var dashboard = ss.getSheetByName("Dashboard");
    if (!dashboard) return { ok: false, fehler: "Dashboard fehlt." };
    var r = parseInt(dashRow, 10);
    if (!r || r < 2) return { ok: false, fehler: "Ungültige Zeile." };
    var layout = getDashboardLayout(dashboard);
    if (r < layout.dataStartRow || r > dashboard.getLastRow()) {
      return { ok: false, fehler: "Zeile außerhalb Datenbereich." };
    }
    var stockId = String(dashboard.getRange(r, layout.cols.stockId).getValue() || "").trim();
    if (!stockId) return { ok: false, fehler: "Keine Stock-ID in Zeile." };
    var entryId = String(dashboard.getRange(r, layout.cols.entryId).getValue() || "").trim();
    if (typeof archiveDashboardRowAt_ !== "function") {
      return { ok: false, fehler: "Archiv-Funktion fehlt im Script." };
    }
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) return { ok: false, fehler: "Anderer Lauf aktiv (Lock)." };
    try {
      var width = Math.max(layout.lastCol || 0, dashboard.getLastColumn());
      var moved = archiveDashboardRowAt_(ss, dashboard, r, layout, stockId, entryId || null, width);
      if (!moved) return { ok: false, fehler: "Zeile konnte nicht archiviert werden." };
      return { ok: true, stockId: stockId, entryId: entryId, dashRow: r };
    } finally {
      try {
        lock.releaseLock();
      } catch (e) {}
    }
  } catch (err) {
    return { ok: false, fehler: String(err) };
  }
}
