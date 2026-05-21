var GMAIL_LOOKUP_SHEET_ID = "16QFzXPUkxvpTHwSSAtjRAeKYb5YdrQPhUrBWInygASE";
var GMAIL_LOOKUP_TAB = "Lookup";
var GMAIL_SYNC_MONTHS = 2;

function getGmailSyncCutoff_() {
  var cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - GMAIL_SYNC_MONTHS);
  cutoff.setHours(0, 0, 0, 0);
  return cutoff;
}

function getGmailSyncAfterQuery_() {
  return Utilities.formatDate(getGmailSyncCutoff_(), "Europe/Berlin", "yyyy/MM/dd");
}

function fetchGmailThreadsSince_(query, cutoff, seenThreads) {
  var threads = [];
  var start = 0;
  var pageSize = 100;
  var maxPages = 15;

  for (var page = 0; page < maxPages; page++) {
    var batch = GmailApp.search(query, start, pageSize);
    if (!batch || batch.length === 0) break;

    var allOlder = true;
    for (var i = 0; i < batch.length; i++) {
      var thread = batch[i];
      var tid = thread.getId();
      if (seenThreads[tid]) continue;

      var messages = thread.getMessages();
      if (!messages.length) continue;
      var newestDate = messages[messages.length - 1].getDate();
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

function collectGmailLookupRows_() {
  var rowMap = {};
  var cutoff = getGmailSyncCutoff_();
  var afterQuery = getGmailSyncAfterQuery_();
  var searches = ["STOCK_ID after:" + afterQuery, "label:N4P after:" + afterQuery];
  var seenThreads = {};

  for (var s = 0; s < searches.length; s++) {
    var threads = fetchGmailThreadsSince_(searches[s], cutoff, seenThreads);
    for (var t = 0; t < threads.length; t++) {
      var messages = threads[t].getMessages();
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
  }

  var rows = [];
  for (var mk in rowMap) {
    if (!Object.prototype.hasOwnProperty.call(rowMap, mk)) continue;
    var item = rowMap[mk];
    rows.push([item.searchKey, item.stockId, item.subject, item.msgDate]);
  }
  rows.sort(function(a, b) {
    var da = a[3] instanceof Date ? a[3].getTime() : 0;
    var db = b[3] instanceof Date ? b[3].getTime() : 0;
    return db - da;
  });
  return rows;
}

function syncGmailOrdersToSheet() {
  var sheet = getOrCreateLookupSheet_();
  var rows = collectGmailLookupRows_();
  var syncedAt = new Date();

  var dataStartRow = 2;
  var lastRow = sheet.getLastRow();
  if (lastRow > dataStartRow - 1) {
    sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 4).clearContent();
  }
  if (rows.length) {
    sheet.getRange(dataStartRow, 1, rows.length, 4).setValues(rows);
  }
  sheet.getRange(1, 7).setValue(syncedAt);
  SpreadsheetApp.flush();
  return { success: true, rows: rows.length, syncedAt: syncedAt.toISOString() };
}

function installGmailSyncTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "syncGmailOrdersToSheet") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("syncGmailOrdersToSheet").timeBased().everyMinutes(2).create();
  return { success: true, message: "Trigger alle 2 Minuten aktiv" };
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
