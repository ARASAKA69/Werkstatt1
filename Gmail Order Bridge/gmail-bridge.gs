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

var PACKZETTEL_TAB = "Packzettel";
var PACKZETTEL_DRIVE_FOLDER = "WMS Belege Archiv";
var PACKZETTEL_SYNC_WEEKS = 1;
var PACKZETTEL_SHARE_ANYONE = true;
var PACKZETTEL_MAX_TEXT = 45000;
var PACKZETTEL_MAX_DOCS_PER_RUN = 150;
var PACKZETTEL_TIME_BUDGET_MS = 1200000;
var PACKZETTEL_ENABLE_OCR = true;
var PACKZETTEL_SYNC_PROPERTY_KEY = "PACKZETTEL_LAST_SYNC_MS";
var PACKZETTEL_QUERY =
  '(from:noreply@n4.parts OR from:alfah.de OR filename:Details.pdf OR subject:Packzettel OR subject:"Auftragsbestätigung") -from:noreply@wm.de -subject:"Online Bestellung" -subject:"Rückgabeantrag"';

function pzGetSyncCutoff_() {
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - PACKZETTEL_SYNC_WEEKS * 7);
  cutoff.setHours(0, 0, 0, 0);
  return cutoff;
}

function pzGetOrCreateFolder_() {
  var it = DriveApp.getFoldersByName(PACKZETTEL_DRIVE_FOLDER);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(PACKZETTEL_DRIVE_FOLDER);
}

function pzGetOrCreateSheet_() {
  var ss = SpreadsheetApp.openById(GMAIL_LOOKUP_SHEET_ID);
  var sheet = ss.getSheetByName(PACKZETTEL_TAB);
  if (!sheet) sheet = ss.insertSheet(PACKZETTEL_TAB);
  var header = [
    "MessageDate", "Source", "Kind", "OrderNumber", "ReferenceNumber",
    "StockID", "Kennzeichen", "OrderDate", "Subject", "DriveFileId",
    "PreviewUrl", "DownloadUrl", "BodyHtml", "RawText", "DedupKey"
  ];
  var firstCell = String(sheet.getRange(1, 1).getValue() || "");
  if (sheet.getLastRow() < 1 || firstCell !== "MessageDate") {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.getRange(1, header.length + 2).setValue("LastSync");
  }
  return sheet;
}

function pzReadExistingKeys_(sheet) {
  var keys = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return keys;
  var col = sheet.getRange(2, 15, lastRow - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    var k = String(col[i][0] || "").trim();
    if (k) keys[k] = true;
  }
  return keys;
}

function pzExtractField_(text, patterns) {
  var t = String(text || "");
  for (var i = 0; i < patterns.length; i++) {
    var m = t.match(patterns[i]);
    if (m && m[1]) return String(m[1]).replace(/\s+/g, " ").trim();
  }
  return "";
}

function pzIsLabelToken_(v) {
  return /^(Referenznummer|Referenz|Kundenname|Kennzeichen|Bestellnummer|Bestelldatum|Datum|Name|NA)$/i.test(String(v || "").trim());
}

function pzExtractOrderNumber_(text) {
  var t = String(text || "");
  var v = pzExtractField_(t, [
    /Online\s+Bestellung\s+([0-9]{5,})/i,
    /Bestellung\s*Nr\.?\s*([0-9]{5,})/i,
    /Bestellnummer[:\s]+([0-9]{5,})/i,
    /Auftragsbest(?:ä|ae?)tigung\s*#\s*([0-9]{5,})/i,
    /Bestellung\s*#\s*([0-9]{5,})/i
  ]);
  if (v) return v.replace(/\s+/g, "");
  var m = t.match(/(N4P\s?\d{5,})/i);
  if (m && m[1]) return m[1].replace(/\s+/g, "");
  return "";
}

function pzIsRejectedRef_(v) {
  return /^(PDE|RDE)/i.test(String(v || "").trim());
}

function pzExtractReference_(text, orderNumber) {
  var t = String(text || "");
  var order = String(orderNumber || "").toUpperCase().replace(/\s+/g, "");

  var m = t.match(/Kunden-?Referenz[:\s]+([A-Z]{1,4}\d{3,})/i);
  if (m && m[1] && !pzIsLabelToken_(m[1]) && !pzIsRejectedRef_(m[1])) return m[1];

  m = t.match(/Referenznummer[:\s]*[\r\n: ]+([A-Z]{1,3}\d{3,})/i);
  if (m && m[1] && !pzIsLabelToken_(m[1]) && !pzIsRejectedRef_(m[1]) && !/^N4P/i.test(m[1])) return m[1];

  m = t.match(/N4P\s?\d{5,}\s+([A-Z]{1,3}\d{3,})/i);
  if (m && m[1] && !pzIsLabelToken_(m[1]) && !pzIsRejectedRef_(m[1])) return m[1];

  var labelPatterns = [
    /Kommission[:\s]+([A-Z]{2,3}\d{3,})\b/i,
    /Kennzeichen[:\s]+([A-Z]{2,3}\d{3,})\b/i,
    /Bemerkung[:\s]+([A-Z]{2,3}\d{3,})\b/i
  ];
  for (var lp = 0; lp < labelPatterns.length; lp++) {
    m = t.match(labelPatterns[lp]);
    if (m && m[1] && !pzIsLabelToken_(m[1]) && !pzIsRejectedRef_(m[1]) && !/^N4P/i.test(m[1])) {
      return m[1].toUpperCase();
    }
  }

  m = t.match(/(N4P\s?\d{5,})/i);
  if (m && m[1]) {
    var cand = m[1].replace(/\s+/g, "").toUpperCase();
    if (cand !== order) return cand;
  }

  return "";
}

function pzExtractKennzeichen_(text) {
  return pzExtractField_(text, [
    /Kennzeichen[:\s]+([A-ZÄÖÜ]{1,3}[- ][A-Z]{1,2}[- ]?\d{1,4})\b/i
  ]);
}

function pzExtractOrderDate_(text) {
  return pzExtractField_(text, [
    /Bestelldatum[:\s]+(\d{1,2}\.\d{1,2}\.\d{2,4})/i,
    /vom\s+(\d{1,2}\.\d{1,2}\.\d{2,4})/i,
    /Bestelldatum[:\s]+(\d{1,2}\.\s*[A-Za-zÄÖÜäöü]+\s+\d{4})/i
  ]);
}

function pzExtractStockId_(text) {
  var t = String(text || "");
  var m = t.match(/STOCK_?ID\s*[:\-]?\s*([A-Z]{2}\d{3,})/i);
  if (m && m[1]) return String(m[1]).replace(/\s+/g, "").toUpperCase();
  m = t.match(/Kommission[:\s]+([A-Z]{2,3}\d{3,})\b/i);
  if (m && m[1] && !pzIsRejectedRef_(m[1]) && !/^N4P/i.test(m[1])) {
    return String(m[1]).replace(/\s+/g, "").toUpperCase();
  }
  return "";
}

function pzDeriveStockFromRef_(stockId, reference) {
  if (stockId) return stockId;
  var ref = String(reference || "").trim();
  if (/^[A-Z]{2,3}\d{3,}$/i.test(ref) && !/^N4P/i.test(ref)) return ref.toUpperCase();
  return "";
}

function pzDriveConvertToDoc_(blob) {
  var stamp = "pz_ocr_temp_" + Date.now();
  if (Drive.Files && typeof Drive.Files.create === "function") {
    var resourceV3 = { name: stamp, mimeType: "application/vnd.google-apps.document" };
    return Drive.Files.create(resourceV3, blob, { ocrLanguage: "de" });
  }
  var resourceV2 = { title: stamp, mimeType: "application/vnd.google-apps.document" };
  return Drive.Files.insert(resourceV2, blob, { ocr: true, ocrLanguage: "de", convert: true });
}

function pzDriveRemove_(fileId) {
  try {
    if (Drive.Files && typeof Drive.Files.remove === "function") {
      Drive.Files.remove(fileId);
    } else if (Drive.Files && typeof Drive.Files.trash === "function") {
      Drive.Files.trash(fileId);
    }
  } catch (e) {}
}

function pzOcrPdfText_(blob) {
  if (typeof Drive === "undefined" || !Drive.Files) return "";
  var docId = null;
  try {
    var inserted = pzDriveConvertToDoc_(blob);
    docId = inserted.id;
  } catch (e) {
    return "";
  }

  var text = "";
  try {
    text = DocumentApp.openById(docId).getBody().getText();
  } catch (readErr) {
    text = "";
  }

  if (!text) {
    try {
      var url = "https://www.googleapis.com/drive/v3/files/" + docId + "/export?mimeType=text%2Fplain";
      var resp = UrlFetchApp.fetch(url, {
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) text = resp.getContentText();
    } catch (expErr) {
      text = text || "";
    }
  }

  pzDriveRemove_(docId);
  return text || "";
}

function pzShareFile_(file) {
  try {
    if (PACKZETTEL_SHARE_ANYONE) {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } else {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    }
  } catch (e) {}
}

function pzSafeFileName_(name, messageId) {
  var base = String(name || "Beleg.pdf").replace(/[\\/:*?"<>|]+/g, "_").trim();
  if (!base) base = "Beleg.pdf";
  base = base.replace(/\.pdf$/i, "");
  var idTag = String(messageId || "").slice(-10);
  return base + "__" + idTag + ".pdf";
}

function pzBuildRowFromPdf_(folder, message, attachment, sender, subject, msgDate, dedupKey) {
  var blob = attachment.copyBlob();
  var fileName = pzSafeFileName_(attachment.getName(), message.getId());
  var file;
  var existing = folder.getFilesByName(fileName);
  if (existing.hasNext()) {
    file = existing.next();
  } else {
    file = folder.createFile(blob).setName(fileName);
    pzShareFile_(file);
  }
  var fileId = file.getId();

  var ocrText = PACKZETTEL_ENABLE_OCR ? pzOcrPdfText_(blob) : "";
  var haystack = ocrText + "\n" + subject + "\n" + (message.getPlainBody ? message.getPlainBody() : "");

  var previewUrl = "https://drive.google.com/file/d/" + fileId + "/preview";
  var downloadUrl = "https://drive.google.com/uc?export=download&id=" + fileId;

  var order = pzExtractOrderNumber_(haystack);
  var ref = pzExtractReference_(haystack, order);
  var stock = pzDeriveStockFromRef_(pzExtractStockId_(haystack), ref);

  return [
    msgDate,
    sender,
    "pdf",
    order,
    ref,
    stock,
    pzExtractKennzeichen_(haystack),
    pzExtractOrderDate_(haystack),
    subject,
    fileId,
    previewUrl,
    downloadUrl,
    "",
    ocrText ? ocrText.substring(0, PACKZETTEL_MAX_TEXT) : "",
    dedupKey
  ];
}

function pzIsExcludedMessage_(sender, subject) {
  if (String(sender || "") === "wm.de") return true;
  var s = String(subject || "").toLowerCase();
  if (s.indexOf("online bestellung") !== -1) return true;
  if (s.indexOf("rückgabeantrag") !== -1 || s.indexOf("rueckgabeantrag") !== -1) return true;
  return false;
}

function pzSenderLabel_(fromRaw) {
  var f = String(fromRaw || "").toLowerCase();
  if (f.indexOf("n4.parts") !== -1) return "n4.parts";
  if (f.indexOf("wm.de") !== -1) return "wm.de";
  if (f.indexOf("alfah") !== -1) return "alfah.de";
  var m = String(fromRaw || "").match(/@([^>\s]+)/);
  return m ? m[1] : String(fromRaw || "").trim();
}

function pzCollectRows_(sheet, existingKeys, incrementalSince) {
  var cutoff = pzGetSyncCutoff_();
  var isIncremental = incrementalSince instanceof Date && !isNaN(incrementalSince.getTime());
  var searchAfter = cutoff;
  if (isIncremental) {
    searchAfter = new Date(incrementalSince.getTime() - GMAIL_INCREMENTAL_OVERLAP_MS);
    if (searchAfter.getTime() < cutoff.getTime()) searchAfter = cutoff;
  }
  var query = PACKZETTEL_QUERY + " after:" + getGmailSyncAfterQuery_(searchAfter);
  var seenThreads = {};
  var threads = fetchGmailThreadsSince_(query, cutoff, seenThreads, isIncremental ? 5 : 15);

  var folder = pzGetOrCreateFolder_();
  var buffer = [];
  var processed = 0;
  var added = 0;
  var startMs = Date.now();
  var reachedLimit = false;

  function budgetLeft() {
    if (processed >= PACKZETTEL_MAX_DOCS_PER_RUN) return false;
    if (Date.now() - startMs > PACKZETTEL_TIME_BUDGET_MS) return false;
    return true;
  }
  function flush() {
    if (buffer.length) {
      pzAppendRows_(sheet, buffer);
      added += buffer.length;
      buffer = [];
    }
  }

  for (var t = 0; t < threads.length && budgetLeft(); t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length; m++) {
      if (!budgetLeft()) { reachedLimit = true; break; }
      var message = messages[m];
      var msgDate = message.getDate();
      if (msgDate.getTime() < cutoff.getTime()) continue;

      var subject = String(message.getSubject() || "");
      var sender = pzSenderLabel_(message.getFrom());
      if (pzIsExcludedMessage_(sender, subject)) continue;
      var messageId = message.getId();
      var attachments = message.getAttachments({ includeInlineImages: false, includeAttachments: true }) || [];
      var pdfs = [];
      for (var a = 0; a < attachments.length; a++) {
        var ct = String(attachments[a].getContentType() || "").toLowerCase();
        var an = String(attachments[a].getName() || "").toLowerCase();
        if (ct.indexOf("pdf") !== -1 || an.indexOf(".pdf") !== -1) pdfs.push(attachments[a]);
      }

      if (pdfs.length) {
        for (var p = 0; p < pdfs.length; p++) {
          if (!budgetLeft()) { reachedLimit = true; break; }
          var dk = messageId + "|" + (pdfs[p].getName() || ("pdf" + p));
          if (existingKeys[dk]) continue;
          existingKeys[dk] = true;
          buffer.push(pzBuildRowFromPdf_(folder, message, pdfs[p], sender, subject, msgDate, dk));
          processed++;
          if (buffer.length >= 8) flush();
        }
      }
    }
  }

  flush();
  return { added: added, reachedLimit: reachedLimit };
}

function pzAppendRows_(sheet, rows) {
  if (rows.length) {
    var startRow = Math.max(2, sheet.getLastRow() + 1);
    sheet.getRange(startRow, 1, rows.length, 15).setValues(rows);
  }
  SpreadsheetApp.flush();
}

function pzSetLastSync_(sheet, syncedAt) {
  sheet.getRange(1, 17).setValue(syncedAt);
  PropertiesService.getScriptProperties().setProperty(PACKZETTEL_SYNC_PROPERTY_KEY, String(syncedAt.getTime()));
  SpreadsheetApp.flush();
}

function pzGetLastSync_(sheet) {
  var cellVal = sheet.getRange(1, 17).getValue();
  if (cellVal instanceof Date && !isNaN(cellVal.getTime())) return cellVal;
  var props = PropertiesService.getScriptProperties().getProperty(PACKZETTEL_SYNC_PROPERTY_KEY);
  if (props) {
    var parsed = new Date(Number(props));
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function syncPackzettelToSheet() {
  var sheet = pzGetOrCreateSheet_();
  var lastSync = pzGetLastSync_(sheet);
  var existingKeys = pzReadExistingKeys_(sheet);

  try {
    var result = pzCollectRows_(sheet, existingKeys, lastSync);
  } catch (e) {
    if (isGmailQuotaError_(e)) {
      return { success: false, error: "Gmail-Tageslimit erreicht", message: String(e.message || e) };
    }
    throw e;
  }

  var partial = !!result.reachedLimit;
  if (!partial) pzSetLastSync_(sheet, new Date());
  return { success: true, added: result.added, partial: partial, incremental: !!lastSync };
}

function syncPackzettelFull() {
  var sheet = pzGetOrCreateSheet_();
  var existingKeys = pzReadExistingKeys_(sheet);
  var result = pzCollectRows_(sheet, existingKeys, null);
  var partial = !!result.reachedLimit;
  if (!partial) pzSetLastSync_(sheet, new Date());
  return { success: true, added: result.added, partial: partial, full: true };
}

function reprocessPackzettelMissing() {
  var sheet = pzGetOrCreateSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, updated: 0, ocred: 0, remaining: 0 };

  var n = lastRow - 1;
  var values = sheet.getRange(2, 1, n, 15).getValues();
  var extractBlock = sheet.getRange(2, 4, n, 5).getValues();
  var rawBlock = sheet.getRange(2, 14, n, 1).getValues();

  var startMs = Date.now();
  var updated = 0;
  var ocred = 0;
  var remaining = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var kind = String(row[2] || "");
    if (kind !== "pdf") continue;

    var text = String(row[13] || "");
    var fileId = String(row[9] || "").trim();

    if (!text && fileId) {
      if (ocred >= PACKZETTEL_MAX_DOCS_PER_RUN || (Date.now() - startMs > PACKZETTEL_TIME_BUDGET_MS)) {
        remaining++;
        continue;
      }
      try {
        text = pzOcrPdfText_(DriveApp.getFileById(fileId).getBlob());
      } catch (e) {
        text = "";
      }
      if (text) {
        ocred++;
        rawBlock[i][0] = text.substring(0, PACKZETTEL_MAX_TEXT);
      } else {
        remaining++;
        continue;
      }
    }

    if (!text) continue;

    var haystack = text + "\n" + String(row[8] || "");
    var order = pzExtractOrderNumber_(haystack);
    var ref = pzExtractReference_(haystack, order);
    var stock = pzDeriveStockFromRef_(pzExtractStockId_(haystack), ref);
    var kennz = pzExtractKennzeichen_(haystack);
    var oDate = pzExtractOrderDate_(haystack);

    extractBlock[i][0] = order;
    extractBlock[i][1] = ref;
    extractBlock[i][2] = stock;
    extractBlock[i][3] = kennz;
    extractBlock[i][4] = oDate;
    updated++;
  }

  sheet.getRange(2, 4, n, 5).setValues(extractBlock);
  sheet.getRange(2, 14, n, 1).setValues(rawBlock);
  SpreadsheetApp.flush();
  return { success: true, updated: updated, ocred: ocred, remaining: remaining };
}

function testReprocessPackzettel() {
  Logger.log(JSON.stringify(reprocessPackzettelMissing(), null, 2));
}

function cleanupPackzettelWmRows() {
  var sheet = pzGetOrCreateSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, removed: 0, kept: 0, trashedFiles: 0 };

  var n = lastRow - 1;
  var values = sheet.getRange(2, 1, n, 15).getValues();
  var keep = [];
  var removed = 0;
  var trashedFiles = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var sender = String(row[1] || "").trim();
    var subject = String(row[8] || "");
    if (!pzIsExcludedMessage_(sender, subject)) {
      keep.push(row);
      continue;
    }
    removed++;
    var fileId = String(row[9] || "").trim();
    if (fileId) {
      try {
        DriveApp.getFileById(fileId).setTrashed(true);
        trashedFiles++;
      } catch (e) {}
    }
  }

  if (removed) {
    sheet.getRange(2, 1, n, 15).clearContent();
    if (keep.length) {
      sheet.getRange(2, 1, keep.length, 15).setValues(keep);
    }
    SpreadsheetApp.flush();
  }
  return { success: true, removed: removed, kept: keep.length, trashedFiles: trashedFiles };
}

function testCleanupPackzettelWmRows() {
  Logger.log(JSON.stringify(cleanupPackzettelWmRows(), null, 2));
}

function installPackzettelSyncTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "syncPackzettelToSheet") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("syncPackzettelToSheet").timeBased().everyMinutes(30).create();
  return { success: true, message: "Packzettel-Trigger alle 30 Minuten aktiv" };
}

function testPackzettelSyncNow() {
  Logger.log(JSON.stringify(syncPackzettelToSheet(), null, 2));
}

function testPackzettelOcr() {
  var out = {
    driveAdvancedServiceAvailable: (typeof Drive !== "undefined" && !!Drive.Files),
    foundPdf: false,
    attachment: "",
    insertOk: false,
    docAppTextLength: 0,
    exportTextLength: 0,
    finalTextLength: 0,
    orderNumber: "",
    referenceNumber: "",
    stockId: "",
    kennzeichen: "",
    error: "",
    textPreview: ""
  };

  var query = 'from:noreply@n4.parts subject:Packzettel has:attachment newer_than:14d';
  var threads = GmailApp.search(query, 0, 10);
  if (!threads.length) {
    query = PACKZETTEL_QUERY + " has:attachment newer_than:14d";
    threads = GmailApp.search(query, 0, 10);
  }

  var blob = null;
  for (var t = 0; t < threads.length && !blob; t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length && !blob; m++) {
      var atts = messages[m].getAttachments() || [];
      for (var a = 0; a < atts.length; a++) {
        var ct = String(atts[a].getContentType() || "").toLowerCase();
        var an = String(atts[a].getName() || "").toLowerCase();
        if (ct.indexOf("pdf") === -1 && an.indexOf(".pdf") === -1) continue;
        blob = atts[a].copyBlob();
        out.foundPdf = true;
        out.attachment = atts[a].getName();
        break;
      }
    }
  }

  if (!blob) {
    Logger.log(JSON.stringify(out, null, 2));
    return out;
  }

  if (!out.driveAdvancedServiceAvailable) {
    out.error = "Drive advanced service NOT enabled (Dienste/Services + -> Drive API).";
    Logger.log(JSON.stringify(out, null, 2));
    return out;
  }

  var docId = null;
  try {
    var inserted = pzDriveConvertToDoc_(blob);
    docId = inserted.id;
    out.insertOk = true;
  } catch (e) {
    out.error = "Drive convert failed: " + e.message;
    Logger.log(JSON.stringify(out, null, 2));
    return out;
  }

  var text = "";
  try {
    text = DocumentApp.openById(docId).getBody().getText();
    out.docAppTextLength = text.length;
  } catch (readErr) {
    out.error = "DocumentApp read failed: " + readErr.message;
  }

  if (!text) {
    try {
      var url = "https://www.googleapis.com/drive/v3/files/" + docId + "/export?mimeType=text%2Fplain";
      var resp = UrlFetchApp.fetch(url, {
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        text = resp.getContentText();
        out.exportTextLength = text.length;
      } else {
        out.error = (out.error ? out.error + " | " : "") + "export HTTP " + resp.getResponseCode();
      }
    } catch (expErr) {
      out.error = (out.error ? out.error + " | " : "") + "export failed: " + expErr.message;
    }
  }

  pzDriveRemove_(docId);

  out.finalTextLength = text.length;
  out.orderNumber = pzExtractOrderNumber_(text);
  out.referenceNumber = pzExtractReference_(text, out.orderNumber);
  out.stockId = pzDeriveStockFromRef_(pzExtractStockId_(text), out.referenceNumber);
  out.kennzeichen = pzExtractKennzeichen_(text);
  out.textPreview = text.substring(0, 400);

  Logger.log(JSON.stringify(out, null, 2));
  return out;
}
