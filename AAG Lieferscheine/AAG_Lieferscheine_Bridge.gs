var AAG_LIEFERSCHEINE_SHEET_ID = "1CnV8VDxHAk-xzuOzfS2qjHgVR9zVtQYa90g0L7FHUXA";
var AAG_LIEFERSCHEINE_TAB = "History 2026";
var AAG_LIEFERSCHEINE_HEADERS = [
  "Sync Zeit",
  "Lieferschein",
  "Belegdatum",
  "Kundenreferenz",
  "AAG Auftrag",
  "Preis Brutto",
  "Status",
  "Position",
  "Marke",
  "Artikelnummer",
  "Artikelbeschreibung",
  "Menge",
  "Item ID",
  "Positionsschlüssel"
];

function aagJsonOut_(value) {
  return ContentService.createTextOutput(JSON.stringify(value || {}))
    .setMimeType(ContentService.MimeType.JSON);
}

function aagSheet_() {
  var ss = SpreadsheetApp.openById(AAG_LIEFERSCHEINE_SHEET_ID);
  var sheet = ss.getSheetByName(AAG_LIEFERSCHEINE_TAB);
  if (!sheet) sheet = ss.insertSheet(AAG_LIEFERSCHEINE_TAB);
  var current = sheet.getRange(1, 1, 1, AAG_LIEFERSCHEINE_HEADERS.length).getValues()[0];
  var wrong = false;
  for (var i = 0; i < AAG_LIEFERSCHEINE_HEADERS.length; i++) {
    if (String(current[i] || "") !== AAG_LIEFERSCHEINE_HEADERS[i]) {
      wrong = true;
      break;
    }
  }
  if (wrong) {
    sheet.getRange(1, 1, 1, AAG_LIEFERSCHEINE_HEADERS.length).setValues([AAG_LIEFERSCHEINE_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, AAG_LIEFERSCHEINE_HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#00244e")
      .setFontColor("#ffffff");
  }
  return sheet;
}

function setupAagLieferscheinBridge() {
  var sheet = aagSheet_();
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 130);
  sheet.setColumnWidth(9, 130);
  sheet.setColumnWidth(10, 150);
  sheet.setColumnWidth(11, 260);
  sheet.setColumnWidth(12, 80);
  sheet.setColumnWidth(13, 130);
  sheet.hideColumns(14);
  return {
    success: true,
    sheetId: AAG_LIEFERSCHEINE_SHEET_ID,
    tab: AAG_LIEFERSCHEINE_TAB,
    syncKeyConfigured: !!PropertiesService.getScriptProperties().getProperty("AAG_SYNC_KEY")
  };
}

function createAagSyncKey() {
  var key = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "");
  PropertiesService.getScriptProperties().setProperty("AAG_SYNC_KEY", key);
  Logger.log(key);
  return key;
}

function aagText_(value, maxLen) {
  var s = String(value == null ? "" : value).trim();
  if (maxLen && s.length > maxLen) s = s.substring(0, maxLen);
  return s;
}

function aagNumber_(value) {
  var n = Number(value);
  return isFinite(n) ? n : "";
}

function aagDate_(value) {
  var n = Number(value);
  if (!isFinite(n) || n <= 0) return "";
  if (n < 100000000000) n *= 1000;
  return Utilities.formatDate(new Date(n), "Europe/Berlin", "dd.MM.yyyy");
}

function aagSyncKeyOk_(provided) {
  var stored = String(PropertiesService.getScriptProperties().getProperty("AAG_SYNC_KEY") || "");
  var incoming = String(provided || "");
  if (!stored || stored.length < 16 || stored.length !== incoming.length) return false;
  var diff = 0;
  for (var i = 0; i < stored.length; i++) diff |= stored.charCodeAt(i) ^ incoming.charCodeAt(i);
  return diff === 0;
}

function aagRowsFromReceipts_(receipts) {
  var rows = [];
  var now = new Date();
  for (var i = 0; i < (receipts || []).length; i++) {
    var receipt = receipts[i] || {};
    var detail = receipt.detailInfo || receipt.header || {};
    var ls = aagText_(detail.receiptNumber || receipt.receiptNumber, 60).toUpperCase();
    if (!/^LS[A-Z0-9-]{5,}$/i.test(ls)) continue;
    var positions = receipt.positions || receipt.data || [];
    if (!positions.length) positions = [{}];
    for (var p = 0; p < positions.length; p++) {
      var pos = positions[p] || {};
      var posNumber = aagText_(pos.positionNumber || pos.lineNumber || ("ROW" + (p + 1)), 80);
      rows.push([
        now,
        ls,
        aagDate_(detail.postingDateTimestamp || pos.postingDateTimestamp),
        aagText_(detail.customerRef || receipt.customerRef, 100).toUpperCase(),
        aagText_(detail.orderNumber || receipt.orderNumber, 100).toUpperCase(),
        aagNumber_(detail.priceGross),
        aagText_(detail.salesStatus || detail.status || pos.status, 80),
        posNumber,
        aagText_(pos.brand || pos.delivererCode, 100),
        aagText_(pos.itemNumber, 120),
        aagText_(pos.itemDescription, 300),
        aagNumber_(pos.quantityDelivered),
        aagText_(pos.itemId, 120),
        ls + "|" + posNumber
      ]);
    }
  }
  return rows;
}

function aagMillis_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") return value.getTime();
  var text = String(value || "").trim();
  var match = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (match) return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1])).getTime();
  var parsed = Date.parse(text);
  return isFinite(parsed) ? parsed : 0;
}

function aagKeepRow_(row, cutoff) {
  var when = aagMillis_(row[2]) || aagMillis_(row[0]);
  return !when || when >= cutoff;
}

function aagUpsertRows_(rows) {
  var sheet = aagSheet_();
  var lastRow = sheet.getLastRow();
  var stored = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, AAG_LIEFERSCHEINE_HEADERS.length).getValues()
    : [];
  var cutoff = new Date();
  cutoff = new Date(cutoff.getFullYear(), cutoff.getMonth(), cutoff.getDate() - 10).getTime();
  var merged = [];
  var existing = {};
  var duplicatesRemoved = 0;
  var removedOld = 0;
  for (var i = 0; i < stored.length; i++) {
    if (!aagKeepRow_(stored[i], cutoff)) {
      removedOld++;
      continue;
    }
    var storedKey = String(stored[i][13] || "");
    if (!storedKey) {
      merged.push(stored[i]);
    } else if (Object.prototype.hasOwnProperty.call(existing, storedKey)) {
      merged[existing[storedKey]] = stored[i];
      duplicatesRemoved++;
    } else {
      existing[storedKey] = merged.length;
      merged.push(stored[i]);
    }
  }
  var added = 0;
  var updated = 0;
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    var rowKey = String(row[13] || "");
    if (Object.prototype.hasOwnProperty.call(existing, rowKey)) {
      merged[existing[rowKey]] = row;
      updated++;
    } else {
      existing[rowKey] = merged.length;
      merged.push(row);
      added++;
    }
  }
  if (merged.length) {
    sheet.getRange(2, 1, merged.length, AAG_LIEFERSCHEINE_HEADERS.length).setValues(merged);
  }
  if (lastRow >= 2) {
    var oldCount = lastRow - 1;
    if (oldCount > merged.length) {
      sheet.getRange(merged.length + 2, 1, oldCount - merged.length, AAG_LIEFERSCHEINE_HEADERS.length).clearContent();
    }
  }
  SpreadsheetApp.flush();
  return { added: added, updated: updated, duplicatesRemoved: duplicatesRemoved, removedOld: removedOld };
}

function doGet(e) {
  var p = e && e.parameter || {};
  if (!aagSyncKeyOk_(p.key)) return aagJsonOut_({ success: false, error: "UNAUTHORIZED" });
  return aagJsonOut_({
    success: true,
    service: "AAG Lieferschein Sync",
    version: "1.2.1",
    sheet: AAG_LIEFERSCHEINE_TAB,
    time: new Date().toISOString()
  });
}

function doPost(e) {
  try {
    var body = JSON.parse(String(e && e.postData && e.postData.contents || "{}"));
    if (!aagSyncKeyOk_(body.key)) return aagJsonOut_({ success: false, error: "UNAUTHORIZED" });
    var action = String(body.action || "");
    if (action !== "sync" && action !== "cleanup") return aagJsonOut_({ success: false, error: "BAD_ACTION" });
    var receipts = action === "cleanup" ? [] : (body.receipts || []);
    if (Object.prototype.toString.call(receipts) !== "[object Array]") return aagJsonOut_({ success: false, error: "BAD_RECEIPTS" });
    if (receipts.length > 300) return aagJsonOut_({ success: false, error: "TOO_MANY_RECEIPTS" });
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return aagJsonOut_({ success: false, error: "SYNC_BUSY" });
    try {
      var rows = aagRowsFromReceipts_(receipts);
      var result = aagUpsertRows_(rows);
      return aagJsonOut_({
        success: true,
        receipts: receipts.length,
        rows: rows.length,
        added: result.added,
        updated: result.updated,
        duplicatesRemoved: result.duplicatesRemoved,
        removedOld: result.removedOld
      });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return aagJsonOut_({ success: false, error: String(err && err.message || err) });
  }
}
