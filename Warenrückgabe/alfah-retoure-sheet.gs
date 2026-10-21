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

function alfahRead_(params) {
  params = params || {};
  if (params.payload) {
    try {
      var extra = JSON.parse(params.payload);
      if (extra && typeof extra === "object") {
        var key;
        for (key in extra) {
          if (Object.prototype.hasOwnProperty.call(extra, key)) params[key] = extra[key];
        }
      }
    } catch (err) {}
  }
  if (typeof params.entries === "string") {
    try { params.entries = JSON.parse(params.entries); } catch (err2) { params.entries = []; }
  }
  return params;
}

function alfahHandle_(params) {
  try {
    params = alfahRead_(params);
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
  var match = String(value || "").toUpperCase().match(/[A-Z]{1,3}\d{4,6}/);
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

var ALFAH_STATUS_COL = 11;
var ALFAH_INFO_COL = 12;
var ALFAH_LINK_COL = 13;
var ALFAH_SEARCH_COL = 14;
var ALFAH_MAIL_KEY = "alfah_mail_primed_v3";
var ALFAH_BOX_KEY = "alfah_mail_box";

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Alfah Retouren")
      .addItem("Mails jetzt lesen", "alfahScanMail")
      .addItem("Mail-Scan alle 30 Minuten", "installAlfahMailScan")
      .addToUi();
  } catch (err) {}
}

function installAlfahMailScan() {
  var triggers = ScriptApp.getProjectTriggers();
  var i;
  for (i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "alfahScanMail") ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger("alfahScanMail").timeBased().everyMinutes(30).create();
  return alfahScanMail();
}

function alfahFreshBody_(body) {
  var raw = String(body || "").replace(/\r/g, "\n");
  var cut = -1;
  var marks = [/\n\s*-----Original/i, /\n\s*Am .{0,120} schrieb/i, /\n_{5,}/];
  var i;
  for (i = 0; i < marks.length; i++) {
    var at = raw.search(marks[i]);
    if (at > 80 && (cut < 0 || at < cut)) cut = at;
  }
  if (cut > 0) raw = raw.substring(0, cut);
  return raw.replace(/\s+/g, " ").trim();
}

function alfahMailKind_(text) {
  var t = String(text || "");
  if (/stornier|\bstorno\b/i.test(t)) return "STORNO";
  if (/ablehn/i.test(t)) return "ABGELEHNT";
  if (/r[üu]ckgabefrist/i.test(t) && /[üu]berschritten/i.test(t)) return "ABGELEHNT";
  if (/nicht\s+retourf[aä]hig|retourunf[aä]hig|nicht\s+retournierbar/i.test(t)) return "ABGELEHNT";
  if (/retoure\s+wird\s+(eingeleitet|veranlasst)/i.test(t) || /r[üu]cksendung\s+(wird\s+)?eingeleitet/i.test(t)) return "RUECK";
  if (/wird\s+.{0,80}abgeholt/i.test(t) || (/durch\s+dhl/i.test(t) && /abhol/i.test(t)) || /retourenlabel/i.test(t)) return "RUECK";
  if (/retoure\s+(ist|wurde)\s+vorbereitet|r[üu]cksendung\s+vorbereitet/i.test(t)) return "VORBEREITET";
  return "";
}

function alfahMailDetail_(text, kind) {
  var t = String(text || "");
  if (/r[üu]ckgabefrist/i.test(t) && /[üu]berschritten/i.test(t)) return "Rückgabefrist überschritten";
  if (/nicht\s+retourf[aä]hig|retourunf[aä]hig|nicht\s+retournierbar/i.test(t)) return "nicht retourfähig";
  if (/35[,.]\s*00/.test(t)) return "unter 35,00 € Netto";
  if (/sonderbeschaffung/i.test(t)) return "Sonderbeschaffung";
  if (kind === "RUECK" && /dhl/i.test(t) && /abhol/i.test(t)) return "Abholung durch DHL";
  var vk = t.match(/VK-RE\d+/i);
  if (vk && kind === "RUECK") return vk[0].toUpperCase();
  var reason = t.match(/ablehnen\s*:\s*([^.]{6,90})/i);
  if (reason && kind === "ABGELEHNT") return reason[1].replace(/\s+/g, " ").trim();
  return "";
}

function alfahStatusPaint_(kind) {
  if (kind === "ABGELEHNT") return { text: "ABGELEHNT", bg: "#cc0000", fg: "#ffffff" };
  if (kind === "STORNO") return { text: "STORNO", bg: "#cc0000", fg: "#ffffff" };
  if (kind === "RUECK") return { text: "Rücksendung eingeleitet", bg: "#00b050", fg: "#ffffff" };
  if (kind === "VORBEREITET") return { text: "VORBEREITET", bg: "#f1c232", fg: "#1f2933" };
  return null;
}

function alfahMailThreads_() {
  var primed = PropertiesService.getScriptProperties().getProperty(ALFAH_MAIL_KEY) === "1";
  var query = primed ? "subject:RETOURE newer_than:3d" : "subject:RETOURE newer_than:120d";
  var threads = [];
  var start = 0;
  while (start < 400) {
    var batch = GmailApp.search(query, start, 100);
    if (!batch.length) break;
    var i;
    for (i = 0; i < batch.length; i++) threads.push(batch[i]);
    if (batch.length < 100) break;
    start += 100;
  }
  return threads;
}

function alfahCollectMail_() {
  var threads = alfahMailThreads_();
  var found = {};
  var step;
  for (step = 0; step < threads.length; step += 40) {
    var slice = threads.slice(step, step + 40);
    var groups = GmailApp.getMessagesForThreads(slice);
    var g;
    for (g = 0; g < groups.length; g++) {
      var messages = groups[g] || [];
      var m;
      for (m = 0; m < messages.length; m++) {
        var msg = messages[m];
        var subject = msg.getSubject() || "";
        var fresh = alfahFreshBody_(msg.getPlainBody() || "");
        var kind = alfahMailKind_(subject + "\n" + fresh);
        if (!kind) continue;
        var ids = subject.match(/R-\d{4,}/gi) || [];
        if (!ids.length) ids = fresh.match(/R-\d{4,}/gi) || [];
        var when = msg.getDate() ? msg.getDate().getTime() : 0;
        var seen = {};
        var n;
        for (n = 0; n < ids.length; n++) {
          var id = String(ids[n] || "").toUpperCase();
          if (!id || seen[id]) continue;
          seen[id] = true;
          var prev = found[id];
          if (prev && prev.when > when) continue;
          found[id] = {
            kind: kind,
            when: when,
            detail: alfahMailDetail_(fresh, kind),
            threadId: slice[g].getId()
          };
        }
      }
    }
  }
  return found;
}

function alfahMailbox_() {
  var stored = PropertiesService.getScriptProperties().getProperty(ALFAH_BOX_KEY) || "";
  var mail = stored;
  if (!mail) {
    try { mail = Session.getEffectiveUser().getEmail() || ""; } catch (err) {}
  }
  if (!mail) {
    try { mail = Session.getActiveUser().getEmail() || ""; } catch (err2) {}
  }
  mail = String(mail || "").trim();
  if (mail) PropertiesService.getScriptProperties().setProperty(ALFAH_BOX_KEY, mail);
  return mail;
}

function alfahThreadUrl_(threadId, mailbox) {
  var id = String(threadId || "").trim();
  if (!id) return "";
  if (mailbox) return "https://mail.google.com/mail/u/?authuser=" + encodeURIComponent(mailbox) + "#all/" + id;
  return "https://mail.google.com/mail/u/0/#all/" + id;
}

function alfahSearchUrl_(rNumber, mailbox) {
  var q = encodeURIComponent(String(rNumber || "") + " RETOURE");
  if (mailbox) return "https://mail.google.com/mail/u/?authuser=" + encodeURIComponent(mailbox) + "#search/" + q;
  return "https://mail.google.com/mail/u/0/#search/" + q;
}

function alfahOurStatus_(text) {
  var t = String(text || "");
  return t === "ABGELEHNT" || t === "STORNO" || t === "Rücksendung eingeleitet" || t === "VORBEREITET";
}

function alfahPutLink_(cell, url, label, color) {
  if (!url) {
    cell.clearContent();
    return;
  }
  var style = SpreadsheetApp.newTextStyle().setForegroundColor(color).setUnderline(true).build();
  cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(label).setLinkUrl(url).setTextStyle(style).build());
}

function alfahPaintRow_(sheet, rowNumber, paint, info, mailUrl, searchUrl) {
  var row = sheet.getRange(rowNumber, 1, 1, ALFAH_SEARCH_COL);
  sheet.getRange(rowNumber, ALFAH_STATUS_COL, 1, 2).setValues([[paint.text, info]]);
  alfahPutLink_(sheet.getRange(rowNumber, ALFAH_LINK_COL), mailUrl, "Mail öffnen", paint.fg);
  alfahPutLink_(sheet.getRange(rowNumber, ALFAH_SEARCH_COL), searchUrl, "R-Nummer suchen", paint.fg);
  row.setBackground(paint.bg);
  row.setFontColor(paint.fg);
  row.setFontFamily("Arial");
  row.setFontSize(10);
  row.setVerticalAlignment("middle");
  sheet.getRange(rowNumber, 2).setFontWeight("bold");
  sheet.getRange(rowNumber, 3).setFontWeight("bold");
  sheet.getRange(rowNumber, ALFAH_STATUS_COL).setFontWeight("bold");
}

function alfahClearAutoRow_(sheet, rowNumber) {
  sheet.getRange(rowNumber, ALFAH_STATUS_COL, 1, 4).clearContent();
  var row = sheet.getRange(rowNumber, 1, 1, ALFAH_SEARCH_COL);
  row.setBackground("#ffffff");
  row.setFontColor("#1f2933");
  row.setFontFamily("Arial");
  row.setFontSize(10);
  sheet.getRange(rowNumber, 2).setFontWeight("bold");
  sheet.getRange(rowNumber, 3).setFontWeight("bold").setFontColor(ALFAH_BLUE);
}

function alfahScanMail() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = alfahSheet_();
    if (!String(sheet.getRange(2, ALFAH_STATUS_COL).getValue() || "")) {
      sheet.getRange(2, ALFAH_STATUS_COL).setValue("Status");
    }
    if (!String(sheet.getRange(2, ALFAH_INFO_COL).getValue() || "")) {
      sheet.getRange(2, ALFAH_INFO_COL).setValue("Info");
    }
    if (!String(sheet.getRange(2, ALFAH_LINK_COL).getValue() || "")) {
      sheet.getRange(2, ALFAH_LINK_COL).setValue("Mail");
    }
    if (!String(sheet.getRange(2, ALFAH_SEARCH_COL).getValue() || "")) {
      sheet.getRange(2, ALFAH_SEARCH_COL).setValue("Suche");
    }
    sheet.getRange(2, ALFAH_STATUS_COL, 1, 4)
      .setBackground(ALFAH_BLUE)
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily("Arial");
    sheet.setColumnWidth(ALFAH_STATUS_COL, 220);
    sheet.setColumnWidth(ALFAH_INFO_COL, 220);
    sheet.setColumnWidth(ALFAH_LINK_COL, 140);
    sheet.setColumnWidth(ALFAH_SEARCH_COL, 140);
    var mailbox = alfahMailbox_();
    var found = alfahCollectMail_();
    var last = sheet.getLastRow();
    if (last < 3) {
      PropertiesService.getScriptProperties().setProperty(ALFAH_MAIL_KEY, "1");
      return { ok: true, updated: 0, mails: Object.keys(found).length };
    }
    var count = last - 2;
    var cases = sheet.getRange(3, 3, count, 1).getValues();
    var current = sheet.getRange(3, ALFAH_STATUS_COL, count, 2).getValues();
    var links = sheet.getRange(3, ALFAH_LINK_COL, count, 1).getFormulas();
    var updated = 0;
    var cleared = 0;
    var i;
    for (i = 0; i < cases.length; i++) {
      var rowNumber = 3 + i;
      var id = String(cases[i][0] || "").toUpperCase().match(/R-\d{4,}/);
      if (!id) continue;
      var hit = found[id[0]];
      var paint = hit ? alfahStatusPaint_(hit.kind) : null;
      if (!paint) {
        var auto = alfahOurStatus_(current[i][0]) && (links[i][0] || String(current[i][1] || ""));
        if (auto) {
          alfahClearAutoRow_(sheet, rowNumber);
          cleared++;
        }
        continue;
      }
      var info = alfahText_(hit.detail);
      alfahPaintRow_(sheet, rowNumber, paint, info, alfahThreadUrl_(hit.threadId, mailbox), alfahSearchUrl_(id[0], mailbox));
      updated++;
    }
    PropertiesService.getScriptProperties().setProperty(ALFAH_MAIL_KEY, "1");
    return { ok: true, updated: updated, cleared: cleared, mails: Object.keys(found).length };
  } finally {
    lock.releaseLock();
  }
}
