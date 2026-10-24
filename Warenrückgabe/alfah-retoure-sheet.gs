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
  alfahFitLayout_(sheet);
  return sheet;
}

function alfahFitLayout_(sheet) {
  var headers = ["Auftrag", "Stock-ID", "R-Nummer", "Teilenummer", "Name", "Lieferant", "Hersteller", "Datum", "Uhrzeit", "Wer", "Status", "Info", "Mail", "Suche", "Vorbereitet", "Geliefert"];
  var titleRange = sheet.getRange(1, 1, 1, ALFAH_LAST_COL);
  if (titleRange.isPartOfMerge()) titleRange.breakApart();
  titleRange.merge();
  titleRange
    .setValue(ALFAH_TITLE)
    .setBackground(ALFAH_BLUE)
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(11)
    .setFontFamily("Arial")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setWrap(true);
  var head = sheet.getRange(2, 1, 1, ALFAH_LAST_COL);
  head.setValues([headers]);
  head.setBackground(ALFAH_BLUE);
  head.setFontColor("#ffffff");
  head.setFontWeight("bold");
  head.setFontSize(ALFAH_FONT);
  head.setFontFamily("Arial");
  head.setHorizontalAlignment("center");
  head.setVerticalAlignment("middle");
  head.setWrap(true);
  sheet.setFrozenRows(2);
  sheet.setRowHeight(1, 22);
  sheet.setRowHeight(2, 32);
  var c;
  for (c = 0; c < ALFAH_WIDTHS.length; c++) sheet.setColumnWidth(c + 1, ALFAH_WIDTHS[c]);
  var last = Math.max(sheet.getLastRow(), 3);
  var rows = last - 2;
  var body = sheet.getRange(3, 1, rows, ALFAH_SEARCH_COL);
  body.setFontFamily("Arial");
  body.setFontSize(ALFAH_FONT);
  body.setWrap(true);
  body.setVerticalAlignment("middle");
  alfahFitChecks_(sheet, rows);
  alfahCompactLabels_(sheet, rows);
}

function alfahRowHasPart_(row) {
  return !!(String(row[0] || "") || String(row[1] || "") || String(row[2] || ""));
}

function alfahFitChecks_(sheet, rows) {
  if (!rows) return;
  var keys = sheet.getRange(3, 1, rows, 3).getValues();
  var boxRange = sheet.getRange(3, ALFAH_PREP_COL, rows, 2);
  var values = boxRange.getValues();
  var i;
  for (i = 0; i < keys.length; i++) {
    if (alfahRowHasPart_(keys[i])) continue;
    values[i][0] = "";
    values[i][1] = "";
  }
  boxRange.clearDataValidations();
  boxRange.setValues(values);
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  var start = -1;
  for (i = 0; i <= keys.length; i++) {
    var has = i < keys.length && alfahRowHasPart_(keys[i]);
    if (has && start < 0) start = i;
    if (!has && start >= 0) {
      var boxes = sheet.getRange(3 + start, ALFAH_PREP_COL, i - start, 2);
      boxes.setDataValidation(rule);
      boxes.setHorizontalAlignment("center");
      boxes.setVerticalAlignment("middle");
      boxes.setFontSize(ALFAH_FONT);
      start = -1;
    }
  }
}

function alfahPutChecks_(sheet, rowNumber) {
  var boxes = sheet.getRange(rowNumber, ALFAH_PREP_COL, 1, 2);
  boxes.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  boxes.setHorizontalAlignment("center");
  boxes.setVerticalAlignment("middle");
  boxes.setFontSize(ALFAH_FONT);
}

function alfahCompactLabels_(sheet, rows) {
  if (!rows) return;
  var statusRange = sheet.getRange(3, ALFAH_STATUS_COL, rows, 1);
  var status = statusRange.getValues();
  var statusChanged = false;
  var i;
  for (i = 0; i < status.length; i++) {
    if (String(status[i][0] || "") === "Rücksendung eingeleitet") {
      status[i][0] = "Rücksendung";
      statusChanged = true;
    }
  }
  if (statusChanged) statusRange.setValues(status);
  alfahRelabelColumn_(sheet, ALFAH_LINK_COL, rows, "Mail");
  alfahRelabelColumn_(sheet, ALFAH_SEARCH_COL, rows, "Suche");
}

function alfahRelabelColumn_(sheet, col, rows, label) {
  var range = sheet.getRange(3, col, rows, 1);
  var rich = range.getRichTextValues();
  var changed = false;
  var i;
  for (i = 0; i < rich.length; i++) {
    var next = alfahRelabelLink_(rich[i][0], label);
    if (!next) continue;
    rich[i][0] = next;
    changed = true;
  }
  if (changed) range.setRichTextValues(rich);
}

function alfahRelabelLink_(rich, label) {
  if (!rich) return null;
  var text = String(rich.getText() || "");
  if (!text || text === label) return null;
  var url = "";
  try { url = rich.getLinkUrl(0, text.length) || ""; } catch (err) {}
  if (!url) {
    try {
      var runs = rich.getRuns();
      var i;
      for (i = 0; i < runs.length; i++) {
        url = runs[i].getLinkUrl() || "";
        if (url) break;
      }
    } catch (err2) {}
  }
  if (!url) return null;
  var color = "#1f2933";
  try {
    var runs2 = rich.getRuns();
    if (runs2.length && runs2[0].getTextStyle() && runs2[0].getTextStyle().getForegroundColor()) {
      color = runs2[0].getTextStyle().getForegroundColor();
    }
  } catch (err3) {}
  var style = SpreadsheetApp.newTextStyle().setForegroundColor(color).setUnderline(true).setFontSize(ALFAH_FONT).setFontFamily("Arial").build();
  return SpreadsheetApp.newRichTextValue().setText(label).setLinkUrl(url).setTextStyle(style).build();
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
    range.setFontSize(ALFAH_FONT);
    range.setVerticalAlignment("middle");
    range.setWrap(true);
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
      var caseId = String(block[r][2] || "").toUpperCase().match(/R-\d{4,}/);
      if (caseId) {
        alfahPutSearch_(sheet, start + r, caseId[0], alfahBox_(), ALFAH_BLUE);
        alfahPutChecks_(sheet, start + r);
      }
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
var ALFAH_PREP_COL = 15;
var ALFAH_DELIVER_COL = 16;
var ALFAH_LAST_COL = 16;
var ALFAH_FONT = 8;
var ALFAH_WIDTHS = [60, 62, 64, 74, 100, 84, 68, 58, 38, 48, 80, 70, 36, 40, 48, 44];
var ALFAH_MAIL_KEY = "alfah_mail_primed_v3";
var ALFAH_BOX_KEY = "alfah_mail_box";
var ALFAH_INBOX = "ersatzteile.hemau@autohero.com";

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Alfah Retouren")
      .addItem("Rücksendung eintragen", "alfahMarkRuecksendung")
      .addItem("Mails jetzt lesen", "alfahScanMail")
      .addItem("Mail-Scan alle 30 Minuten", "installAlfahMailScan")
      .addToUi();
  } catch (err) {}
  try { alfahSheet_(); } catch (err2) {}
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
  if (/unter\s+wert/i.test(t)) return "ABGELEHNT";
  if (/kein(?:en|e)?\s+lagerort/i.test(t)) return "ABGELEHNT";
  if (/keine\s+r[üu]ckgabem[öo]glichkeit/i.test(t)) return "ABGELEHNT";
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
  if (/unter\s+wert/i.test(t)) return "unter Wert, keine Rückgabe";
  if (/keine\s+r[üu]ckgabem[öo]glichkeit/i.test(t)) return "keine Rückgabemöglichkeit";
  if (/kein(?:en|e)?\s+lagerort/i.test(t)) return "kein Lagerort";
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
  if (kind === "RUECK") return { text: "Rücksendung", bg: "#00b050", fg: "#ffffff" };
  if (kind === "VORBEREITET") return { text: "VORBEREITET", bg: "#f1c232", fg: "#1f2933" };
  return null;
}

function alfahHtmlText_(html) {
  var s = String(html || "");
  s = s.replace(/<style[\s\S]*?<\/style>/gi, " ");
  s = s.replace(/<script[\s\S]*?<\/script>/gi, " ");
  s = s.replace(/<br\s*\/?>/gi, "\n");
  s = s.replace(/<\/(p|div|tr|li|h\d|td)>/gi, "\n");
  s = s.replace(/<[^>]+>/g, " ");
  s = s.replace(/&nbsp;/gi, " ");
  s = s.replace(/&amp;/gi, "&");
  s = s.replace(/&quot;/gi, "\"");
  s = s.replace(/&#(\d+);/g, function(_, n) { return String.fromCharCode(parseInt(n, 10)); });
  s = s.replace(/&#x([0-9a-f]+);/gi, function(_, n) { return String.fromCharCode(parseInt(n, 16)); });
  return s;
}

function alfahMessageText_(msg) {
  var plain = "";
  var html = "";
  try { plain = msg.getPlainBody() || ""; } catch (err) {}
  try { html = alfahHtmlText_(msg.getBody() || ""); } catch (err2) {}
  if (plain && alfahMailKind_(plain)) return plain;
  if (html.length > plain.length) return html;
  return plain || html;
}

function alfahMailThreads_(query, limit) {
  var threads = [];
  var start = 0;
  var cap = limit || 400;
  while (start < cap) {
    var size = Math.min(100, cap - start);
    var batch = GmailApp.search(query, start, size);
    if (!batch.length) break;
    var i;
    for (i = 0; i < batch.length; i++) threads.push(batch[i]);
    if (batch.length < size) break;
    start += batch.length;
  }
  return threads;
}

function alfahRememberMail_(found, looked, id, when, kind, detail, threadId) {
  if (!id) return;
  looked[id] = true;
  if (!kind) return;
  var prev = found[id];
  if (prev && prev.when > when) return;
  if (prev && prev.when === when && (prev.kind === "ABGELEHNT" || prev.kind === "STORNO") && kind !== "ABGELEHNT" && kind !== "STORNO") return;
  found[id] = { kind: kind, when: when, detail: detail, threadId: threadId };
}

function alfahReadThreads_(threads, found, looked) {
  var step;
  for (step = 0; step < threads.length; step += 40) {
    var slice = threads.slice(step, step + 40);
    var groups = [];
    try { groups = GmailApp.getMessagesForThreads(slice); } catch (err) { groups = []; }
    var g;
    for (g = 0; g < groups.length; g++) {
      var messages = groups[g] || [];
      var threadId = slice[g] ? slice[g].getId() : "";
      var m;
      for (m = 0; m < messages.length; m++) {
        var msg = messages[m];
        var subject = "";
        var body = "";
        var when = 0;
        try { subject = msg.getSubject() || ""; } catch (e1) {}
        try { body = alfahMessageText_(msg); } catch (e2) {}
        try { when = msg.getDate() ? msg.getDate().getTime() : 0; } catch (e3) {}
        var fresh = alfahFreshBody_(body);
        var kind = alfahMailKind_(subject + "\n" + fresh);
        var used = fresh;
        if (!kind) {
          kind = alfahMailKind_(subject + "\n" + body);
          used = body;
        }
        var ids = subject.match(/R-\d{4,}/gi) || [];
        if (!ids.length) ids = (fresh + "\n" + body).match(/R-\d{4,}/gi) || [];
        var seen = {};
        var n;
        for (n = 0; n < ids.length; n++) {
          var id = String(ids[n] || "").toUpperCase();
          if (!id || seen[id]) continue;
          seen[id] = true;
          alfahRememberMail_(found, looked, id, when, kind, kind ? alfahMailDetail_(used, kind) : "", threadId);
        }
      }
    }
  }
}

function alfahCollectMail_(extraIds) {
  var primed = PropertiesService.getScriptProperties().getProperty(ALFAH_MAIL_KEY) === "1";
  var query = primed ? "subject:RETOURE newer_than:21d" : "subject:RETOURE newer_than:120d";
  var found = {};
  var looked = {};
  alfahReadThreads_(alfahMailThreads_(query, primed ? 400 : 400), found, looked);
  var need = [];
  var seenNeed = {};
  var e;
  for (e = 0; e < (extraIds || []).length; e++) {
    var want = String(extraIds[e] || "").toUpperCase();
    if (!/^R-\d{4,}$/.test(want) || found[want] || seenNeed[want]) continue;
    seenNeed[want] = true;
    need.push(want);
  }
  var b;
  for (b = 0; b < need.length && b < 180; b += 12) {
    var chunk = need.slice(b, b + 12);
    var q = chunk.map(function(id) { return '"' + id + '"'; }).join(" OR ") + " newer_than:60d";
    alfahReadThreads_(alfahMailThreads_(q, 40), found, looked);
  }
  found._looked = looked;
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

function alfahGmailUrl_(hash) {
  var target = "https://mail.google.com/mail/u/0/?authuser=" + encodeURIComponent(ALFAH_INBOX) + hash;
  return "https://accounts.google.com/AccountChooser?Email=" + encodeURIComponent(ALFAH_INBOX) + "&continue=" + encodeURIComponent(target);
}

function alfahThreadUrl_(threadId) {
  var id = String(threadId || "").trim();
  if (!id) return "";
  return alfahGmailUrl_("#all/" + id);
}

function alfahSearchUrl_(rNumber) {
  var q = encodeURIComponent(String(rNumber || "") + " RETOURE");
  return alfahGmailUrl_("#search/" + q);
}

function alfahInboxLink_(rich) {
  var url = "";
  try { url = rich ? String(rich.getLinkUrl() || "") : ""; } catch (err) { url = ""; }
  var text = "";
  try { text = decodeURIComponent(url); } catch (err2) { text = url; }
  return text.toLowerCase().indexOf(ALFAH_INBOX) !== -1;
}

function alfahThreadFromUrl_(url) {
  var text = String(url || "");
  try { text = decodeURIComponent(text); } catch (err) {}
  var match = text.match(/#(?:all|inbox)\/([a-f0-9]{8,})/i);
  return match ? match[1] : "";
}

function alfahIsRueck_(text) {
  var t = String(text || "");
  return t === "Rücksendung" || t === "Rücksendung eingeleitet";
}

function alfahOurStatus_(text) {
  var t = String(text || "");
  return t === "ABGELEHNT" || t === "STORNO" || alfahIsRueck_(t) || t === "VORBEREITET";
}

function alfahBox_() {
  return String(PropertiesService.getScriptProperties().getProperty(ALFAH_BOX_KEY) || "").trim();
}

function alfahSearchColor_(status) {
  if (status === "ABGELEHNT" || status === "STORNO" || alfahIsRueck_(status)) return "#ffffff";
  if (status === "VORBEREITET") return "#1f2933";
  return ALFAH_BLUE;
}

function alfahPutSearch_(sheet, rowNumber, rNumber, mailbox, color) {
  alfahPutLink_(sheet.getRange(rowNumber, ALFAH_SEARCH_COL), alfahSearchUrl_(rNumber, mailbox), "Suche", color || ALFAH_BLUE);
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
  alfahPutLink_(sheet.getRange(rowNumber, ALFAH_LINK_COL), mailUrl, "Mail", paint.fg);
  alfahPutLink_(sheet.getRange(rowNumber, ALFAH_SEARCH_COL), searchUrl, "Suche", paint.fg);
  row.setBackground(paint.bg);
  row.setFontColor(paint.fg);
  row.setFontFamily("Arial");
  row.setFontSize(ALFAH_FONT);
  row.setWrap(true);
  row.setVerticalAlignment("middle");
  sheet.getRange(rowNumber, 2).setFontWeight("bold");
  sheet.getRange(rowNumber, 3).setFontWeight("bold");
  sheet.getRange(rowNumber, ALFAH_STATUS_COL).setFontWeight("bold");
}

function alfahClearAutoRow_(sheet, rowNumber) {
  sheet.getRange(rowNumber, ALFAH_STATUS_COL, 1, 3).clearContent();
  var row = sheet.getRange(rowNumber, 1, 1, ALFAH_SEARCH_COL);
  row.setBackground("#ffffff");
  row.setFontColor("#1f2933");
  row.setFontFamily("Arial");
  row.setFontSize(ALFAH_FONT);
  row.setWrap(true);
  sheet.getRange(rowNumber, 2).setFontWeight("bold");
  sheet.getRange(rowNumber, 3).setFontWeight("bold").setFontColor(ALFAH_BLUE);
}

function alfahScanMail() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = alfahSheet_();
    var mailbox = alfahMailbox_();
    var last = sheet.getLastRow();
    var count = last >= 3 ? last - 2 : 0;
    var cases = count ? sheet.getRange(3, 3, count, 1).getValues() : [];
    var current = count ? sheet.getRange(3, ALFAH_STATUS_COL, count, 2).getValues() : [];
    var openIds = [];
    var i;
    for (i = 0; i < cases.length; i++) {
      var openId = String(cases[i][0] || "").toUpperCase().match(/R-\d{4,}/);
      if (!openId) continue;
      var openStatus = String(current[i][0] || "");
      if (openStatus === "ABGELEHNT" || openStatus === "STORNO") continue;
      openIds.push(openId[0]);
    }
    var found = alfahCollectMail_(openIds);
    var looked = found._looked || {};
    delete found._looked;
    if (last < 3) {
      PropertiesService.getScriptProperties().setProperty(ALFAH_MAIL_KEY, "1");
      return { ok: true, updated: 0, mails: Object.keys(found).length };
    }
    var links = sheet.getRange(3, ALFAH_LINK_COL, count, 1).getFormulas();
    var searchRich = sheet.getRange(3, ALFAH_SEARCH_COL, count, 1).getRichTextValues();
    var mailRich = sheet.getRange(3, ALFAH_LINK_COL, count, 1).getRichTextValues();
    var updated = 0;
    var cleared = 0;
    for (i = 0; i < cases.length; i++) {
      var rowNumber = 3 + i;
      var id = String(cases[i][0] || "").toUpperCase().match(/R-\d{4,}/);
      if (!id) continue;
      var hit = found[id[0]];
      var paint = hit ? alfahStatusPaint_(hit.kind) : null;
      if (!paint) {
        var infoNow = String(current[i][1] || "");
        var manual = infoNow === "Alfah System";
        var auto = !manual && looked[id[0]] && alfahOurStatus_(current[i][0]) && (links[i][0] || infoNow);
        if (auto) {
          alfahClearAutoRow_(sheet, rowNumber);
          alfahPutSearch_(sheet, rowNumber, id[0], mailbox, ALFAH_BLUE);
          cleared++;
        } else if (!alfahInboxLink_(searchRich[i][0])) {
          alfahPutSearch_(sheet, rowNumber, id[0], mailbox, alfahSearchColor_(String(current[i][0] || "")));
        }
        if (!auto) {
          var oldMail = "";
          try { oldMail = mailRich[i][0] ? String(mailRich[i][0].getLinkUrl() || "") : ""; } catch (errMail) { oldMail = ""; }
          var oldThread = alfahThreadFromUrl_(oldMail);
          if (oldThread && !alfahInboxLink_(mailRich[i][0])) {
            alfahPutLink_(sheet.getRange(rowNumber, ALFAH_LINK_COL), alfahThreadUrl_(oldThread), "Mail", alfahSearchColor_(String(current[i][0] || "")));
          }
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

function alfahWantedIds_(raw) {
  var text = String(raw || "").toUpperCase();
  var out = [];
  var seen = {};
  var re = /R-\d{4,}|\b\d{4,6}\b/g;
  var m;
  while ((m = re.exec(text))) {
    var id = m[0].indexOf("R-") === 0 ? m[0] : "R-" + m[0];
    if (seen[id]) continue;
    seen[id] = true;
    out.push(id);
  }
  return out;
}

function alfahPaintManualRueck_(sheet, rowNumber, rNumber) {
  var paint = alfahStatusPaint_("RUECK");
  var row = sheet.getRange(rowNumber, 1, 1, ALFAH_SEARCH_COL);
  sheet.getRange(rowNumber, ALFAH_STATUS_COL).setValue(paint.text);
  sheet.getRange(rowNumber, ALFAH_INFO_COL).setValue("Alfah System");
  row.setBackground(paint.bg);
  row.setFontColor(paint.fg);
  row.setFontFamily("Arial");
  row.setFontSize(ALFAH_FONT);
  row.setWrap(true);
  row.setVerticalAlignment("middle");
  sheet.getRange(rowNumber, 2).setFontWeight("bold");
  sheet.getRange(rowNumber, 3).setFontWeight("bold");
  sheet.getRange(rowNumber, ALFAH_STATUS_COL).setFontWeight("bold");
  alfahPutSearch_(sheet, rowNumber, rNumber, alfahMailbox_(), paint.fg);
  alfahPutChecks_(sheet, rowNumber);
}

function alfahMarkRuecksendung() {
  var html = HtmlService.createHtmlOutput(alfahRueckHtml_()).setTitle("Rücksendung");
  SpreadsheetApp.getUi().showSidebar(html);
}

function alfahApplyRuecksendung(raw) {
  var wanted = alfahWantedIds_(raw);
  if (!wanted.length) return { ok: true, done: [], skipped: [], missing: [] };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = alfahSheet_();
    var last = sheet.getLastRow();
    var foundRows = {};
    var w;
    for (w = 0; w < wanted.length; w++) foundRows[wanted[w]] = [];
    if (last >= 3) {
      var cases = sheet.getRange(3, 3, last - 2, 1).getValues();
      var i;
      for (i = 0; i < cases.length; i++) {
        var id = String(cases[i][0] || "").toUpperCase().match(/R-\d{4,}/);
        if (!id || !foundRows[id[0]]) continue;
        foundRows[id[0]].push(3 + i);
      }
    }
    var done = [];
    var skipped = [];
    var missing = [];
    for (w = 0; w < wanted.length; w++) {
      var key = wanted[w];
      var rows = foundRows[key] || [];
      if (!rows.length) {
        missing.push(key);
        continue;
      }
      var painted = 0;
      var r;
      for (r = 0; r < rows.length; r++) {
        var status = String(sheet.getRange(rows[r], ALFAH_STATUS_COL).getValue() || "");
        var info = String(sheet.getRange(rows[r], ALFAH_INFO_COL).getValue() || "");
        if (status === "ABGELEHNT" || status === "STORNO") {
          skipped.push(key + " ist " + status);
          continue;
        }
        if (alfahIsRueck_(status) && info && info !== "Alfah System") {
          skipped.push(key + " schon grün");
          continue;
        }
        alfahPaintManualRueck_(sheet, rows[r], key);
        painted++;
      }
      if (painted) done.push(key);
    }
    return { ok: true, done: done, skipped: skipped, missing: missing };
  } finally {
    lock.releaseLock();
  }
}

function alfahRueckHtml_() {
  return '<!DOCTYPE html><html><head><base target="_top"><style>'
    + 'body{font-family:Arial,sans-serif;margin:0;padding:16px;color:#1f2933}'
    + 'h2{font-size:15px;margin:0 0 8px}'
    + 'p{font-size:12px;line-height:1.45;color:#52606d;margin:0 0 12px}'
    + 'textarea{width:100%;height:140px;box-sizing:border-box;border:1px solid #cbd2d9;border-radius:8px;padding:8px;font:13px Arial,sans-serif;resize:vertical}'
    + 'button{margin-top:10px;width:100%;background:#00b050;color:#fff;border:0;border-radius:8px;padding:10px;font-weight:bold;cursor:pointer}'
    + 'button:disabled{opacity:.6}'
    + '#out{margin-top:12px;font-size:12px;white-space:pre-wrap;line-height:1.45}'
    + '.ok{color:#087f3a}.bad{color:#9b1c1c}'
    + '</style></head><body>'
    + '<h2>Rücksendung eintragen</h2>'
    + '<p>R-Nummern einfügen, die Alfah im System vorbereitet hat und zu denen keine Mail kam. Jede gefundene Zeile wird grün und bekommt Rücksendung.</p>'
    + '<textarea id="raw" placeholder="R-96434&#10;R-96432&#10;R-96431"></textarea>'
    + '<button id="go" type="button">Grün setzen</button>'
    + '<div id="out"></div>'
    + '<script>'
    + 'document.getElementById("go").onclick=function(){'
    + 'var btn=document.getElementById("go");var out=document.getElementById("out");'
    + 'btn.disabled=true;out.className="";out.textContent="Schreibe…";'
    + 'google.script.run.withSuccessHandler(function(res){'
    + 'btn.disabled=false;var lines=[];var d=(res&&res.done)||[];var s=(res&&res.skipped)||[];var m=(res&&res.missing)||[];var i;'
    + 'for(i=0;i<d.length;i++)lines.push("Grün: "+d[i]);'
    + 'for(i=0;i<s.length;i++)lines.push(s[i]);'
    + 'for(i=0;i<m.length;i++)lines.push("Nicht in der Liste: "+m[i]);'
    + 'if(!lines.length)lines.push("Keine R-Nummer erkannt.");'
    + 'out.className=(m.length||s.length)?"bad":"ok";out.textContent=lines.join("\\n");'
    + '}).withFailureHandler(function(err){'
    + 'btn.disabled=false;out.className="bad";out.textContent=String(err&&err.message?err.message:err);'
    + '}).alfahApplyRuecksendung(document.getElementById("raw").value);};'
    + '</script></body></html>';
}
