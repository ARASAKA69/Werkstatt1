var ALFAH_SHEET_ID = "1P8s15IFPyQb3OFkjW__YZZ9Q1-Qui9b4Y5k4RJMJtpc";
var ALFAH_GID = 786084326;
var ALFAH_TAB = "Alfah";
var ALFAH_TITLE = "Alfah Retouren";
var ALFAH_BLUE = "#3c78d8";
var ALFAH_HEADERS = ["Auftrag", "Stock-ID", "R-Nummer", "Teilenummer", "Name", "Lieferant", "Hersteller", "Datum", "Uhrzeit", "Wer"];

function alfahOut_(value) {
  return ContentService.createTextOutput(JSON.stringify(value || {})).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return alfahHandle_(e && e.parameter ? e.parameter : {});
}

function doPost(e) {
  var params = {};
  try {
    if (e && e.postData && e.postData.contents) params = JSON.parse(e.postData.contents);
  } catch (err) {
    return alfahOut_({ ok: false, error: "JSON kaputt" });
  }
  return alfahHandle_(params);
}

function alfahHandle_(params) {
  try {
    var action = String((params && params.action) || "ping");
    if (action === "ping" || action === "setup") {
      alfahSheet_();
      return alfahOut_({ ok: true, tab: ALFAH_TAB });
    }
    if (action === "append") {
      var entries = params.entries || [];
      if (!Array.isArray(entries)) entries = [];
      return alfahOut_(alfahAppend_(entries));
    }
    if (action === "fixWho") return alfahOut_(alfahFixWho_(params.who));
    return alfahOut_({ ok: false, error: "unbekannt" });
  } catch (err) {
    return alfahOut_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function setupAlfahReturnsSheet() {
  alfahSheet_();
}

function alfahBook_() {
  return SpreadsheetApp.openById(ALFAH_SHEET_ID);
}

function alfahTab_(ss) {
  var sheets = ss.getSheets();
  var i;
  for (i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === ALFAH_GID) return sheets[i];
  }
  var named = ss.getSheetByName(ALFAH_TAB);
  if (named) return named;
  return ss.insertSheet(ALFAH_TAB);
}

function alfahSheet_() {
  var sheet = alfahTab_(alfahBook_());
  sheet.setTabColor(ALFAH_BLUE);
  if (!String(sheet.getRange(1, 1).getValue() || "")) {
    sheet.getRange(1, 1).setValue(ALFAH_TITLE);
    sheet.getRange(2, 1, 1, ALFAH_HEADERS.length).setValues([ALFAH_HEADERS]);
  }
  if (String(sheet.getRange(1, 1).getValue() || "") !== ALFAH_TITLE) return sheet;
  var titleRange = sheet.getRange(1, 1, 1, ALFAH_HEADERS.length);
  if (titleRange.isPartOfMerge()) titleRange.breakApart();
  titleRange.merge();
  titleRange
    .setValue(ALFAH_TITLE)
    .setBackground(ALFAH_BLUE)
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(14)
    .setFontFamily("Arial")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
  var head = sheet.getRange(2, 1, 1, ALFAH_HEADERS.length);
  head.setValues([ALFAH_HEADERS]);
  head.setBackground(ALFAH_BLUE);
  head.setFontColor("#ffffff");
  head.setFontWeight("bold");
  head.setFontSize(11);
  head.setFontFamily("Arial");
  head.setVerticalAlignment("middle");
  sheet.setFrozenRows(2);
  sheet.setRowHeight(1, 34);
  sheet.setRowHeight(2, 26);
  var widths = [110, 110, 110, 150, 280, 210, 170, 110, 80, 150];
  var c;
  for (c = 0; c < widths.length; c++) sheet.setColumnWidth(c + 1, widths[c]);
  sheet.getRange("E:E").setWrap(true);
  return sheet;
}

function alfahText_(value) {
  var text = String(value == null ? "" : value).replace(/\s+/g, " ").trim();
  if (text.length > 180) text = text.substring(0, 180);
  if (text.charAt(0) === "=") text = "'" + text;
  return text;
}

function alfahStock_(value) {
  var match = String(value || "").toUpperCase().match(/[A-Z]{2}\d{5}/);
  return match ? match[0] : "";
}

function alfahKey_(caseId, article) {
  var r = String(caseId || "").toUpperCase().match(/R-\d+/);
  var articleNo = String(article || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (!r || !articleNo) return "";
  return r[0] + "|" + articleNo;
}

function alfahLine_(line) {
  return [
    alfahText_(line.orderId),
    alfahStock_(line.stockId),
    alfahText_(line.caseId).toUpperCase(),
    alfahText_(line.article),
    alfahText_(line.name),
    alfahText_(line.supplier),
    alfahText_(line.brand),
    alfahText_(line.date),
    alfahText_(line.time),
    alfahText_(line.who)
  ];
}

function alfahFixWho_(name) {
  var who = alfahText_(name);
  if (!who || /^\d+$/.test(who)) return { ok: false, error: "Name fehlt" };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = alfahSheet_();
    var last = sheet.getLastRow();
    if (last < 3) return { ok: true, updated: 0 };
    var count = last - 2;
    var cases = sheet.getRange(3, 3, count, 1).getValues();
    var whoRange = sheet.getRange(3, 10, count, 1);
    var values = whoRange.getValues();
    var updated = 0;
    var i;
    for (i = 0; i < values.length; i++) {
      if (!/R-\d+/i.test(String(cases[i][0] || ""))) continue;
      var current = String(values[i][0] || "").trim();
      if (current && current !== "nachgetragen" && !/^\d+$/.test(current)) continue;
      if (current === who) continue;
      values[i][0] = who;
      updated++;
    }
    if (updated) {
      whoRange.setNumberFormat("@");
      whoRange.setValues(values);
    }
    return { ok: true, updated: updated };
  } finally {
    lock.releaseLock();
  }
}

function alfahAppend_(entries) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = alfahSheet_();
    var last = sheet.getLastRow();
    var seen = {};
    if (last >= 3) {
      var have = sheet.getRange(3, 3, last - 2, 2).getValues();
      var i;
      for (i = 0; i < have.length; i++) {
        var known = alfahKey_(have[i][0], have[i][1]);
        if (known) seen[known] = true;
      }
    }
    var block = [];
    var added = 0;
    var skipped = 0;
    var e;
    for (e = 0; e < entries.length; e++) {
      var lines = entries[e] && entries[e].lines ? entries[e].lines : [entries[e]];
      var fresh = [];
      var n;
      for (n = 0; n < lines.length; n++) {
        var line = lines[n] || {};
        var key = alfahKey_(line.caseId, line.article);
        if (!key) continue;
        if (seen[key]) {
          skipped++;
          continue;
        }
        seen[key] = true;
        fresh.push(alfahLine_(line));
        added++;
      }
      if (!fresh.length) continue;
      var f;
      for (f = 0; f < fresh.length; f++) block.push(fresh[f]);
      block.push(["", "", "", "", "", "", "", "", "", ""]);
    }
    if (!block.length) return { ok: true, added: 0, skipped: skipped };
    var start = last < 3 ? 3 : last + 2;
    var range = sheet.getRange(start, 1, block.length, ALFAH_HEADERS.length);
    range.setNumberFormat("@");
    range.setValues(block);
    range.setFontFamily("Arial");
    range.setFontSize(10);
    range.setVerticalAlignment("middle");
    var r;
    for (r = 0; r < block.length; r++) {
      var rowRange = sheet.getRange(start + r, 1, 1, ALFAH_HEADERS.length);
      if (!block[r].join("")) {
        rowRange.setBackground(null);
        rowRange.setBorder(false, false, false, false, false, false);
        continue;
      }
      rowRange.setBackground("#ffffff");
      rowRange.setFontColor("#1f2933");
      rowRange.setBorder(true, true, true, true, true, true, "#d6deea", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(start + r, 2).setFontWeight("bold");
      sheet.getRange(start + r, 3).setFontWeight("bold").setFontColor(ALFAH_BLUE);
    }
    return { ok: true, added: added, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}
