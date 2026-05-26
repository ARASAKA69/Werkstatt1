var GMAIL_LOOKUP_SHEET_ID = "16QFzXPUkxvpTHwSSAtjRAeKYb5YdrQPhUrBWInygASE";
var GMAIL_LOOKUP_TAB = "Lookup";
var GMAIL_SYNC_WEEKS = 4;
var GMAIL_INCREMENTAL_OVERLAP_MS = 24 * 60 * 60 * 1000;
var GMAIL_SYNC_PROPERTY_KEY = "GMAIL_LAST_SYNC_MS";

function getGmailSyncCutoff_() {
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - GMAIL_SYNC_WEEKS * 7);
  cutoff.setHours(0, 0, 0, 0);
  return cutoff;
}

function getGmailSyncAfterQuery_(afterDate) {
  return Utilities.formatDate(afterDate, "Europe/Berlin", "yyyy/MM/dd");
}

function fetchGmailThreadsSince_(query, cutoff, seenThreads, maxPages) {
  var threads = [];
  var start = 0;
  var pageSize = 100;
  var pageLimit = maxPages || 15;

  for (var page = 0; page < pageLimit; page++) {
    var batch = GmailApp.search(query, start, pageSize);
    if (!batch || batch.length === 0) break;

    var allOlder = true;
    for (var i = 0; i < batch.length; i++) {
      var thread = batch[i];
      var tid = thread.getId();
      if (seenThreads[tid]) continue;

      var newestDate = thread.getLastMessageDate();
      if (newestDate.getTime() < cutoff.getTime()) continue;

      allOlder = false;
      seenThreads[tid] = true;
      threads.push(thread);
    }

    if (batch.length < pageSize) break;
    if (allOlder) break;
    start += pageSize;
  }

  return threads;
}

function getOrCreateLookupSheet_() {
  var ss = SpreadsheetApp.openById(GMAIL_LOOKUP_SHEET_ID);
  var sheet = ss.getSheetByName(GMAIL_LOOKUP_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(GMAIL_LOOKUP_TAB);
  }
  if (sheet.getLastRow() < 1 || String(sheet.getRange(1, 1).getValue() || "") !== "SearchKey") {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 4).setValues([["SearchKey", "StockID", "Subject", "MessageDate"]]);
    sheet.getRange(1, 6).setValue("LastSync");
  }
  return sheet;
}

function getLastSyncTime_(sheet) {
  var cellVal = sheet.getRange(1, 7).getValue();
  if (cellVal instanceof Date && !isNaN(cellVal.getTime())) return cellVal;
  var props = PropertiesService.getScriptProperties().getProperty(GMAIL_SYNC_PROPERTY_KEY);
  if (props) {
    var parsed = new Date(Number(props));
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function readExistingLookupRows_(sheet) {
  var rowMap = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return rowMap;

  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  for (var i = 0; i < data.length; i++) {
    var searchKey = String(data[i][0] || "");
    var stockId = String(data[i][1] || "");
    if (!searchKey || !stockId) continue;

    var msgDate = data[i][3];
    var msgTime = msgDate instanceof Date ? msgDate.getTime() : 0;
    var mapKey = searchKey + "|" + stockId;
    rowMap[mapKey] = {
      searchKey: searchKey,
      stockId: stockId,
      subject: data[i][2],
      msgDate: msgDate,
      msgTime: msgTime
    };
  }
  return rowMap;
}

function extractSearchKeysFromSubject_(subject) {
  var keys = {};
  var subjectText = String(subject || "");
  var subjectNorm = subjectText.replace(/\s+/g, "").toUpperCase();
  if (!subjectNorm) return [];

  var n4Match = subjectText.match(/N4PARTS\s*:\s*(N4P\d+)/i);
  if (n4Match && n4Match[1]) {
    var n4Full = String(n4Match[1]).replace(/\s+/g, "").toUpperCase();
    keys[n4Full] = true;
    var n4Num = n4Full.replace(/^N4P/i, "");
    if (n4Num.length >= 4) keys[n4Num] = true;
  }

  var parts = subjectText.split("---");
  if (parts.length > 0) {
    var orderNum = String(parts[0]).replace(/\s+/g, "").toUpperCase();
    if (orderNum.length >= 4) keys[orderNum] = true;
  }

  return Object.keys(keys);
}

function addThreadToRowMap_(thread, cutoff, rowMap) {
  var messages = thread.getMessages();
  for (var m = messages.length - 1; m >= 0; m--) {
    var msgDate = messages[m].getDate();
    if (msgDate.getTime() < cutoff.getTime()) continue;

    var subject = String(messages[m].getSubject() || "");
    var stockMatch = subject.match(/STOCK_ID\s*:\s*([A-Z]{2}\d{3,})/i);
    if (!stockMatch || !stockMatch[1]) continue;

    var stockId = String(stockMatch[1]).replace(/\s+/g, "").toUpperCase();
    var keys = extractSearchKeysFromSubject_(subject);
    for (var k = 0; k < keys.length; k++) {
      var searchKey = keys[k];
      var mapKey = searchKey + "|" + stockId;
      if (!rowMap[mapKey] || msgDate.getTime() > rowMap[mapKey].msgTime) {
        rowMap[mapKey] = {
          searchKey: searchKey,
          stockId: stockId,
          subject: subject,
          msgDate: msgDate,
          msgTime: msgDate.getTime()
        };
      }
    }
  }
}

function rowMapToSortedRows_(rowMap, cutoff) {
  var rows = [];
  var cutoffTime = cutoff.getTime();
  for (var mk in rowMap) {
    if (!Object.prototype.hasOwnProperty.call(rowMap, mk)) continue;
    var item = rowMap[mk];
    if (item.msgTime < cutoffTime) continue;
    rows.push([item.searchKey, item.stockId, item.subject, item.msgDate]);
  }
  rows.sort(function(a, b) {
    var da = a[3] instanceof Date ? a[3].getTime() : 0;
    var db = b[3] instanceof Date ? b[3].getTime() : 0;
    return db - da;
  });
  return rows;
}

function collectGmailLookupRows_(existingRowMap, incrementalSince) {
  var rowMap = existingRowMap || {};
  var cutoff = getGmailSyncCutoff_();
  var isIncremental = incrementalSince instanceof Date && !isNaN(incrementalSince.getTime());
  var searchAfter = cutoff;

  if (isIncremental) {
    searchAfter = new Date(incrementalSince.getTime() - GMAIL_INCREMENTAL_OVERLAP_MS);
    if (searchAfter.getTime() < cutoff.getTime()) searchAfter = cutoff;
  }

  var afterQuery = getGmailSyncAfterQuery_(searchAfter);
  var query = "(STOCK_ID OR label:N4P) after:" + afterQuery;
  var maxPages = isIncremental ? 5 : 15;
  var seenThreads = {};
  var threads = fetchGmailThreadsSince_(query, cutoff, seenThreads, maxPages);

  for (var t = 0; t < threads.length; t++) {
    addThreadToRowMap_(threads[t], cutoff, rowMap);
  }

  return rowMapToSortedRows_(rowMap, cutoff);
}

function isGmailQuotaError_(error) {
  return String(error && error.message ? error.message : error).indexOf("too many times") !== -1;
}

function writeLookupRows_(sheet, rows, syncedAt) {
  var dataStartRow = 2;
  var lastRow = sheet.getLastRow();
  if (lastRow >= dataStartRow) {
    sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 4).clearContent();
  }
  if (rows.length) {
    sheet.getRange(dataStartRow, 1, rows.length, 4).setValues(rows);
  }
  sheet.getRange(1, 7).setValue(syncedAt);
  PropertiesService.getScriptProperties().setProperty(GMAIL_SYNC_PROPERTY_KEY, String(syncedAt.getTime()));
  SpreadsheetApp.flush();
}

function syncGmailOrdersToSheet() {
  var sheet = getOrCreateLookupSheet_();
  var lastSync = getLastSyncTime_(sheet);
  var existingRowMap = readExistingLookupRows_(sheet);

  try {
    var rows = collectGmailLookupRows_(existingRowMap, lastSync);
  } catch (e) {
    if (isGmailQuotaError_(e)) {
      return {
        success: false,
        error: "Gmail-Tageslimit erreicht",
        message: String(e.message || e),
        rows: sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 0
      };
    }
    throw e;
  }

  var syncedAt = new Date();
  writeLookupRows_(sheet, rows, syncedAt);
  return { success: true, rows: rows.length, syncedAt: syncedAt.toISOString(), incremental: !!lastSync };
}

function syncGmailOrdersFull() {
  var sheet = getOrCreateLookupSheet_();
  try {
    var rows = collectGmailLookupRows_({}, null);
  } catch (e) {
    if (isGmailQuotaError_(e)) {
      return {
        success: false,
        error: "Gmail-Tageslimit erreicht",
        message: String(e.message || e)
      };
    }
    throw e;
  }

  var syncedAt = new Date();
  writeLookupRows_(sheet, rows, syncedAt);
  return { success: true, rows: rows.length, syncedAt: syncedAt.toISOString(), full: true };
}

function installGmailSyncTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "syncGmailOrdersToSheet") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("syncGmailOrdersToSheet").timeBased().everyMinutes(15).create();
  return { success: true, message: "Trigger alle 15 Minuten aktiv" };
}

function testSearchOrder() {
  Logger.log(JSON.stringify(searchOrderInGmail_("2611233477"), null, 2));
}

function searchOrderInGmail_(query) {
  query = String(query || "").replace(/\s+/g, "").toUpperCase();
  if (!query) return { found: false, message: "Keine Suchanfrage" };

  var cleanQuery = query.replace(/^N4P/i, "");
  if (!cleanQuery || cleanQuery.length < 4) {
    return { found: false, message: "Suchanfrage zu kurz (min. 4 Zeichen)" };
  }

  var searchTerms = ['"' + cleanQuery + '"'];
  if (query !== cleanQuery) searchTerms.push('"' + query + '"');
  searchTerms.push("N4PARTS:N4P" + cleanQuery);
  if (query.indexOf("N4P") === 0) searchTerms.push("N4PARTS:" + query);

  var threads = [];
  var seen = {};
  for (var s = 0; s < searchTerms.length; s++) {
    var batch = GmailApp.search(searchTerms[s], 0, 15);
    for (var i = 0; i < batch.length; i++) {
      var tid = batch[i].getId();
      if (seen[tid]) continue;
      seen[tid] = true;
      threads.push(batch[i]);
    }
  }

  var best = null;
  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    for (var m = messages.length - 1; m >= 0; m--) {
      var subject = String(messages[m].getSubject() || "");
      var subjectNorm = subject.replace(/\s+/g, "").toUpperCase();
      if (subjectNorm.indexOf(cleanQuery) === -1 && subjectNorm.indexOf(query) === -1) continue;

      var stockMatch = subject.match(/STOCK_ID\s*:\s*([A-Z]{2}\d{3,})/i);
      if (!stockMatch || !stockMatch[1]) continue;

      var hit = {
        found: true,
        stockId: String(stockMatch[1]).replace(/\s+/g, "").toUpperCase(),
        message: "Gefunden via Gmail",
        subject: subject,
        date: messages[m].getDate().getTime()
      };
      if (!best || hit.date > best.date) best = hit;
    }
  }

  if (best) {
    return {
      found: true,
      stockId: best.stockId,
      message: best.message + ": " + String(best.subject).substring(0, 100),
      source: "gmail"
    };
  }

  return { found: false, message: "Bestellnummer '" + query + "' nicht in Gmail gefunden." };
}

function testSyncNow() {
  Logger.log(JSON.stringify(syncGmailOrdersToSheet(), null, 2));
}

function testSyncFull() {
  Logger.log(JSON.stringify(syncGmailOrdersFull(), null, 2));
}
