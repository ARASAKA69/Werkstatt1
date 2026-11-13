var ALFAH_SHEET_ID = "1P8s15IFPyQb3OFkjW__YZZ9Q1-Qui9b4Y5k4RJMJtpc";
var ALFAH_GID = 786084326;
var ALFAH_TAB = "Alfah";
var ALFAH_TITLE = "Alfah Retouren";
var ALFAH_REJECT_GID = 1505373902;
var ALFAH_REJECT_TAB = "Alfah Abgelehnt";
var ALFAH_REJECT_TITLE = "Alfah Abgelehnt";
var ALFAH_REJECT_RED = "#cc0000";
var ALFAH_BLUE = "#3c78d8";
var ALFAH_GREEN = "#00b050";
var ALFAH_ORANGE = "#e69138";
var ALFAH_DONE_GID = 1870949588;
var ALFAH_DONE_TAB = "Alfah Geliefert";
var ALFAH_DONE_TITLE = "Alfah Geliefert";
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
    if (action === "wmPing" || action === "wmSetup") {
      wmSheet_();
      return alfahOut_({ ok: true, tab: WM_TAB });
    }
    if (action === "wmAppend") {
      var wmEntries = params.entries || [];
      if (!Array.isArray(wmEntries)) wmEntries = [params];
      return alfahOut_(wmAppend_(wmEntries));
    }
    if (action === "aagPing" || action === "aagSetup") {
      aagSheet_();
      return alfahOut_({ ok: true, tab: AAG_TAB });
    }
    if (action === "aagAppend") {
      var aagEntries = params.entries || [];
      if (!Array.isArray(aagEntries)) aagEntries = [params];
      return alfahOut_(aagAppend_(aagEntries));
    }
    if (action === "staPing" || action === "staSetup") {
      staSheet_();
      return alfahOut_({ ok: true, tab: STA_TAB });
    }
    if (action === "staAppend") {
      var staEntries = params.entries || [];
      if (!Array.isArray(staEntries)) staEntries = [params];
      return alfahOut_(staAppend_(staEntries));
    }
    if (action === "scanMail") {
      return alfahOut_(alfahScanMail());
    }
    if (action === "scanStatus") {
      var lastScanMs = Number(PropertiesService.getScriptProperties().getProperty("alfah_mail_last_scan_ms") || 0);
      return alfahOut_({ ok: true, lastScanMs: lastScanMs, intervalMin: 30 });
    }
    if (action === "oemSync") {
      return alfahOut_(alfahSyncOemTeile());
    }
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

function alfahEnsureCols_(sheet) {
  var max = sheet.getMaxColumns();
  if (max < ALFAH_LAST_COL) sheet.insertColumnsAfter(max, ALFAH_LAST_COL - max);
}

function alfahFitLayout_(sheet) {
  var headers = ["Auftrag", "Stock-ID", "R-Nummer", "Teilenummer", "Name", "Lieferant", "Hersteller", "Datum", "Uhrzeit", "Wer", "Status", "Info", "Mail", "Suche", "Vorbereitet", "Geliefert", "Letzte Mail", "Nachfrage", "Mail-Text"];
  alfahEnsureCols_(sheet);
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
  alfahFitExtraCols_(sheet, rows);
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
var ALFAH_MAILAT_COL = 17;
var ALFAH_ASK_COL = 18;
var ALFAH_TEXT_COL = 19;
var ALFAH_LAST_COL = 19;
var ALFAH_FONT = 8;
var ALFAH_WIDTHS = [60, 62, 64, 74, 100, 84, 68, 58, 38, 48, 80, 70, 36, 40, 48, 44, 78, 78, 320];
var ALFAH_MAIL_KEY = "alfah_mail_primed_v3";
var ALFAH_BOX_KEY = "alfah_mail_box";
var ALFAH_CURSOR_KEY = "alfah_mail_cursor";
var ALFAH_INBOX = "ersatzteile.hemau@autohero.com";
var ALFAH_CHASE_DAYS = 7;
var ALFAH_CHASE_RED = "#cc0000";
var ALFAH_CHASE_ORANGE = "#e69138";
var ALFAH_SCAN_BUDGET_MS = 240000;
var ALFAH_FINAL = { RUECK: true, ABGELEHNT: true, STORNO: true, VORBEREITET: true };

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Alfah Retouren")
      .addItem("Rücksendung eintragen", "alfahMarkRuecksendung")
      .addItem("Nachfrage gesendet (markierte Zeilen)", "alfahMarkNachfrage")
      .addItem("Mails jetzt lesen", "alfahScanMail")
      .addItem("Mail-Scan alle 30 Minuten", "installAlfahMailScan")
      .addItem("OEM_Teile abgleichen", "alfahSyncOemTeileMenu")
      .addToUi();
  } catch (err) {}
  try { alfahSheet_(); } catch (err2) {}
  try { wmSheet_(); } catch (err3) {}
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getSheetId() !== ALFAH_GID) return;
    var first = e.range.getColumn();
    var lastCol = first + e.range.getNumColumns() - 1;
    if (lastCol < ALFAH_STATUS_COL || first > ALFAH_ASK_COL) return;
    alfahPaintChase_(sheet);
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

function alfahMarkNachfrage() {
  var ui = SpreadsheetApp.getUi();
  var sheet = alfahSheet_();
  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  if (!range || range.getSheet().getSheetId() !== sheet.getSheetId()) {
    ui.alert("Bitte im Tab Alfah die Zeilen markieren, für die eine Nachfrage gesendet wurde.");
    return;
  }
  var start = Math.max(range.getRow(), 3);
  var end = range.getRow() + range.getNumRows() - 1;
  if (end < start) return;
  var ids = sheet.getRange(start, 3, end - start + 1, 1).getDisplayValues();
  var stamp = alfahStampText_(Date.now());
  var marked = 0;
  var i;
  for (i = 0; i < ids.length; i++) {
    if (!/R-\d{4,}/i.test(String(ids[i][0] || ""))) continue;
    var cell = sheet.getRange(start + i, ALFAH_ASK_COL);
    cell.setNumberFormat("@");
    cell.setValue(stamp);
    marked++;
  }
  alfahPaintChase_(sheet);
  SpreadsheetApp.getActiveSpreadsheet().toast(marked + " Zeile(n) als nachgefragt markiert.", "Alfah Retouren", 4);
}

function alfahStampText_(ms) {
  return Utilities.formatDate(new Date(ms), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
}

function alfahStampMs_(value, display) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) return value.getTime();
  var s = String(display || value || "").trim();
  var m = s.match(/(\d{1,2})\.(\d{1,2})\.(\d{2,4})(?:\D+(\d{1,2}):(\d{2}))?/);
  if (!m) return 0;
  var y = Number(m[3]);
  if (y < 100) y += 2000;
  return new Date(y, Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0)).getTime();
}

function alfahCellText_(value) {
  var text = String(value == null ? "" : value).replace(/\s+/g, " ").trim();
  if (text.length > 500) text = text.substring(0, 497) + "…";
  if (/^[=+\-@]/.test(text)) text = "'" + text;
  return text;
}

function alfahFlat_(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

function alfahCutQuote_(body) {
  var raw = String(body || "").replace(/\r/g, "\n");
  var cut = -1;
  var marks = [/\n\s*-{3,}\s*(?:Original|Urspr[üu]ngliche)/i, /\n\s*Am .{0,160}schrieb/i, /\n\s*On .{0,160}wrote/i, /\n_{5,}/];
  var i;
  for (i = 0; i < marks.length; i++) {
    var at = raw.search(marks[i]);
    if (at > 80 && (cut < 0 || at < cut)) cut = at;
  }
  var heads = [];
  var re = /(^|\n)[ \t>*]*(?:Von|From)\s*:/gi;
  var m;
  while ((m = re.exec(raw))) heads.push(m.index);
  if (heads.length) {
    var before = alfahFlat_(raw.substring(0, heads[0]));
    var intro = before.length < 80 || (/folgende\s+(?:nachricht|r[üu]ckmeldung)/i.test(before) && !/gr[üu][ßs]en|regards/i.test(before));
    var headCut = intro ? (heads.length > 1 ? heads[1] : -1) : heads[0];
    if (headCut > 0 && (cut < 0 || headCut < cut)) cut = headCut;
  }
  return cut > 0 ? raw.substring(0, cut) : raw;
}

function alfahFreshBody_(body) {
  return alfahFlat_(alfahCutQuote_(body));
}

function alfahMailSnippet_(text) {
  var t = String(text || "").replace(/\r/g, "\n");
  t = t.replace(/^[ \t>*]*(?:von|from|gesendet|sent|an|to|cc|betreff|subject|datum|date|priorit[äa]t|wichtigkeit)\s*:.*$/gim, " ");
  t = t.replace(/(?:lieber\s+kunde|liebe\s+kundin|guten\s+(?:tag|morgen|abend)(?:\s+zusammen)?|hallo(?:\s+zusammen)?|sehr\s+geehrte[^,\n]{0,60}|verehrte[^,\n]{0,40}|moin)\s*[,!]?/gi, " ");
  t = t.replace(/wir\s+haben\s+(?:nach)?folgende\s+(?:nachricht|r[üu]ckmeldung)[^:\n]{0,80}:?/gi, " ");
  t = t.replace(/vielen\s+dank\s+f[üu]r\s+ihre\s+nachricht\.?/gi, " ");
  t = t.replace(/[\-\u2010-\u2015_=#'\s]*(?:please\s+)?(?:type\s+your\s+)?(?:reply|write)\s+above\s+this\s+line[\-\u2010-\u2015_=#\s]*/gi, " ");
  t = t.replace(/[\-\u2010-\u2015_=#'\s]*bitte\s+(?:nur\s+)?(?:oberhalb|über)\s+dieser\s+zeile\s+antworten[\-\u2010-\u2015_=#\s]*/gi, " ");
  t = t.replace(/\[cid:[^\]]*\]?/gi, " ");
  t = t.replace(/\[image:[^\]]*\]/gi, " ");
  var end = t.search(/mit\s+freundliche[nm]\s+gr[üu](?:[ßs]|ss)(?:en)?|freundlicher\s+gr[üu][ßs]|freundliche\s+gr[üu][ßs]e|herzliche\s+gr[üu][ßs]e|viele\s+gr[üu][ßs]e|beste\s+gr[üu][ßs]e|\bmfg\b|bei\s+r[üu]ckfragen\s+erreichen|with\s+kind(?:est)?\s+regards|best\s+regards/i);
  if (end > 0) t = t.substring(0, end);
  return alfahCellText_(t);
}

function alfahPolicyFree_(t) {
  return String(t || "")
    .replace(/(?:bis\s+zum|vom)\s+\d+\.\s*(?:bis\s+zum\s+\d+\.\s*)?tag\s*:[^:]{0,90}?m[öo]glich(?:\s+mit\s+einem\s+abzug\s+von\s+\d+\s*%)?/gi, " ")
    .replace(/nach\s+ablauf\s+von\s+\d+\s+tagen\s*:[^.]{0,90}\./gi, " ");
}

function alfahRejectText_(text) {
  var t = alfahPolicyFree_(text);
  return /ablehn|abgelehnt|abgewiesen/i.test(t)
    || /lehne[nt]?\s+(?:\S+\s+){0,5}ab\b/i.test(t)
    || /unter\s+wert/i.test(t)
    || /kein(?:en|e)?\s+lagerort/i.test(t)
    || /keine\s+lagerware/i.test(t)
    || /keine\s+r[üu]ckgabem[öo]glichkeit/i.test(t)
    || /nicht\s+retourf[aä]hig|retourunf[aä]hig|nicht\s+retournierbar|nicht\s+r[üu]ckgabef[aä]hig/i.test(t)
    || /frist\s*[üu]berschreitung|frist(?:en)?\s+(?:\S+\s+){0,3}(?:[üu]berschritten|verstrichen|abgelaufen)/i.test(t)
    || /(?:r[üu]ckgabe|r[üu]cknahme|retoure|r[üu]cksendung)[^.!?]{0,60}nicht\s+(?:mehr\s+)?m[öo]glich/i.test(t)
    || /keine\s+(?:retoure|r[üu]ckgabe|r[üu]cknahme)\s+(?:anmelden|m[öo]glich|annehmen)/i.test(t)
    || /nicht\s+(?:mehr\s+)?zur\s+retoure/i.test(t)
    || /mindest\S*wert/i.test(t)
    || /unter\s+dem\s+(?:\S+\s+){0,2}\S*wert/i.test(t)
    || /unter\s+\d+(?:[.,]\d+)?\s*(?:€|eur)/i.test(t);
}

function alfahAskText_(t) {
  return /\d+\s+m[öo]glichkeiten/i.test(t)
    || /bitte\s+nennen\s+sie\s+(?:uns\s+)?(?:ein(?:en)?\s+)?(?:abhol|termin|datum)/i.test(t)
    || /sollen\s+wir\s+(?:die\s+)?(?:abholung|retoure|r[üu]cksendung|r[üu]ckgabe)/i.test(t)
    || /(?:ben[öo]tigen|brauchen)\s+wir\s+(?:\S+\s+){0,4}(?:fotos?|bilder?)/i.test(t)
    || /bitte\s+(?:senden|schicken|[üu]bermitteln)\s+sie\s+(?:uns\s+)?(?:\S+\s+){0,3}(?:fotos|bilder)/i.test(t)
    || /(?:retoure|r[üu]ckgabe|r[üu]cknahme|r[üu]cksendung)\s+w[äa]re\s+(?:nur\s+)?(?:\S+\s+){0,2}m[öo]glich/i.test(t)
    || /(?:sind|w[äa]ren)\s+sie\s+(?:damit\s+)?einverstanden/i.test(t)
    || /(?:wie|welche)\s+(?:\S+\s+){0,2}(?:m[öo]chten|wollen|w[üu]nschen)\s+sie/i.test(t);
}

function alfahSoftAskText_(t) {
  return /bitte\s+um\s+(?:eine\s+)?(?:kurze\s+)?r[üu]ckmeldung|bitte\s+teilen\s+sie\s+uns\s+mit|um\s+(?:kurze\s+)?r[üu]ckmeldung\s+wird\s+gebeten/i.test(t);
}

function alfahPickupText_(t) {
  return /wird\s+.{0,80}abgeholt/i.test(t)
    || (/durch\s+(?:dhl|dpd|ups|gls)/i.test(t) && /abhol/i.test(t))
    || /abholer\s+(?:ist|wurde)\s+(?:erstellt|beauftragt|bestellt|angelegt)/i.test(t)
    || /abholung\s+(?:\S+\s+){0,4}(?:erstellt|beauftragt|bestellt|veranlasst|eingeleitet|angelegt|gebucht)/i.test(t)
    || /(?:holen|holt)\s+(?:\S+\s+){0,6}ab\b/i.test(t)
    || /fahrer\s+(?:\S+\s+){0,3}(?:vorbeikommen|abholen|kommen)/i.test(t)
    || /um\s+(?:das|die|den)\s+(?:\S+\s+){0,2}abzuholen/i.test(t)
    || /abholauftrag/i.test(t);
}

function alfahSelfShipText_(t) {
  return /(?:auf|per)\s+eigenem\s+(?:post)?weg/i.test(t)
    || /an\s+folgende\s+adresse\s+(?:\S+\s+){0,2}(?:zur[üu]ck|senden|schicken)/i.test(t)
    || /(?:senden|schicken)\s+sie\s+(?:uns\s+)?(?:das|die|den)\s+(?:\S+\s+){0,6}(?:zur[üu]ck|an\s+folgende)/i.test(t)
    || /uns\s+zusenden|zur[üu]cksenden\s+an/i.test(t);
}

function alfahApprovedText_(t) {
  return /genehmigt|freigegeben/i.test(t)
    || /(?:retoure|r[üu]cksendung|r[üu]ckgabe|r[üu]cknahme)\s+(?:wird|wurde|ist)\s+(?:eingeleitet|veranlasst)/i.test(t)
    || /r[üu]cksendung\s+eingeleitet|wir\s+veranlassen\s+(?:die\s+)?(?:retoure|r[üu]cksendung)/i.test(t);
}

function alfahCheckText_(t) {
  return /muss\s+(?:\S+\s+){0,2}angemeldet\s+werden/i.test(t)
    || /(?:wird|werden)\s+(?:\S+\s+){0,3}gepr[üu]ft|in\s+pr[üu]fung/i.test(t)
    || /etwas\s+geduld/i.test(t)
    || /warte[nt]?\s+(?:noch\s+)?auf\s+(?:die\s+)?(?:r[üu]ckmeldung|antwort|freigabe)/i.test(t)
    || /r[üu]ckmeldung\s+(?:\S+\s+){0,3}(?:folgt|steht\s+(?:noch\s+)?aus)/i.test(t)
    || /(?:haben|wurde)\s+(?:\S+\s+){0,3}nachgefragt/i.test(t)
    || /melden\s+uns\s+(?:\S+\s+){0,3}(?:sobald|wieder)/i.test(t)
    || /(?:retoure|r[üu]ckgabe)\s+(?:wurde|ist)\s+angemeldet/i.test(t)
    || /weitergeleitet/i.test(t);
}

function alfahMailKind_(text) {
  var t = alfahFlat_(text);
  if (!t) return "";
  if (/stornier|\bstorno\b/i.test(t)) return "STORNO";
  if (alfahRejectText_(t)) return "ABGELEHNT";
  if (alfahAskText_(t)) return "RUECKFRAGE";
  if (alfahLabelMail_(t)) return "RUECK";
  if (alfahPickupText_(t)) return "RUECK";
  if (alfahSelfShipText_(t)) return "RUECK";
  if (/retoure\s+(ist|wurde)\s+vorbereitet|r[üu]cksendung\s+vorbereitet/i.test(t)) return "VORBEREITET";
  if (alfahSoftAskText_(t)) return "RUECKFRAGE";
  if (alfahApprovedText_(t)) return "VERANLASST";
  if (alfahCheckText_(t)) return "PRUEFUNG";
  return "";
}

function alfahLabelMail_(text) {
  var t = String(text || "");
  if (/retouren?\s*-?\s*(label|unterlagen|schein)/i.test(t)) return true;
  if (/r[üu]cksende\s*-?\s*(label|unterlagen|schein)/i.test(t)) return true;
  if (/\blabels?\b.{0,100}(paket|packst[üu]ck|anbringen|befestigen|ausdrucken|aufkleben)/i.test(t)) return true;
  if (/(paket|packst[üu]ck|ausdrucken|anbringen).{0,100}\blabels?\b/i.test(t)) return true;
  if (/ihre\s+labels/i.test(t)) return true;
  return false;
}

function alfahNoiseMail_(fresh) {
  var t = String(fresh || "");
  if (/die\s+retoure?\s+muss\s+in\s+der\s+originalverpackung\s+erfolgen/i.test(t) && t.length < 2000) return true;
  if (/ihre\s+(?:anfrage|nachricht|reklamation)\s+(?:ist|wurde)\s+(?:bei\s+uns\s+)?eingegangen|eingangsbest[äa]tigung|automatische\s+antwort|abwesenheitsnotiz|out\s+of\s+office/i.test(t) && t.length < 900) return true;
  return false;
}

function alfahConditions_(t) {
  var out = [];
  var p = t.match(/(\d{1,2})\s*%\s*(?:abzug|des\s+blp|geb[üu]hr|abschlag)/i) || t.match(/(?:abzug|abschlag|geb[üu]hr)\s+(?:von\s+)?(\d{1,2})\s*%/i) || t.match(/-\s?(\d{1,2})\s*%/);
  if (p) out.push(p[1] + "% Abzug");
  var f = t.match(/(\d+(?:[.,]\d{1,2})?)\s*(?:€|eur\w*)\s*r[üu]ckfracht/i) || t.match(/r[üu]ckfracht\s+(?:von\s+)?(\d+(?:[.,]\d{1,2})?)\s*(?:€|eur)/i);
  if (f) out.push(f[1] + " € Rückfracht");
  return out;
}

function alfahMailDetail_(text, kind) {
  var t = alfahPolicyFree_(alfahFlat_(text));
  var parts = [];
  if (kind === "STORNO") parts.push("storniert");
  if (kind === "ABGELEHNT") {
    if (/frist/i.test(t) && /[üu]berschr|verstrichen|abgelaufen/i.test(t)) parts.push("Rückgabefrist überschritten");
    if (/nicht\s+retourf|retourunf|nicht\s+retournierbar|nicht\s+r[üu]ckgabef/i.test(t)) parts.push("nicht retourfähig");
    if (/keine\s+lagerware/i.test(t)) parts.push("keine Lagerware");
    if (/kein(?:en|e)?\s+lagerort/i.test(t)) parts.push("kein Lagerort");
    if (/sonderbeschaffung/i.test(t)) parts.push("Sonderbeschaffung");
    if (/mindest\S*wert|unter\s+wert|unter\s+dem\s+(?:\S+\s+){0,2}\S*wert|unter\s+\d+(?:[.,]\d+)?\s*(?:€|eur)/i.test(t)) {
      var amt = t.match(/unter\s+(\d+(?:[.,]\d+)?)\s*(?:€|eur)/i) || t.match(/warenwert\s+von\s+(\d+(?:[.,]\d+)?)\s*(?:€|eur)/i) || t.match(/(\d+[.,]\d{2})\s*(?:€|eur)\s*netto/i);
      parts.push("unter Mindestwarenwert" + (amt ? " " + amt[1] + " €" : ""));
    }
    if (/keine\s+r[üu]ckgabem[öo]glichkeit/i.test(t)) parts.push("keine Rückgabemöglichkeit");
    if (/nie\s+(?:\S+\s+){0,3}(?:an)?geliefert|nicht\s+(?:\S+\s+){0,3}(?:an)?geliefert|nicht\s+(?:bei\s+uns\s+)?(?:gekauft|bezogen)/i.test(t)) parts.push("nicht vom Lieferanten geliefert");
    var swap = t.match(/ersetzt\s+(?:\S+\s+){0,2}(?:in|durch)\s+([A-Z0-9][A-Z0-9\-]{4,})/i);
    if (swap) parts.push("ersetzt durch " + swap[1]);
    if (!parts.length && /abgelehnt|abgewiesen|lehne[nt]?\s/i.test(t)) parts.push("Retoure abgelehnt");
    var reason = t.match(/ablehnen\s*:\s*([^.]{6,90})/i);
    if (reason) parts.push(reason[1].trim());
    if (!parts.length) parts.push("Rückgabe nicht möglich");
  }
  if (kind === "RUECK") {
    if (alfahSelfShipText_(t)) {
      var addr = t.match(/adresse\s+(?:\S+\s+){0,2}zur[üu]ck\s*:?\s*([\s\S]{8,140}?)(?=\s+(?:mit\s+freundlichen|freundliche|viele\s+gr|bei\s+r[üu]ckfragen)|$)/i);
      parts.push("Selbstversand" + (addr ? " an " + addr[1].trim() : ""));
    } else if (/dhl/i.test(t) && /abhol/i.test(t)) {
      parts.push("Abholung durch DHL");
    } else if (alfahPickupText_(t)) {
      parts.push("Abholung beauftragt");
    } else {
      var vk = t.match(/VK-RE\d+/i);
      parts.push(vk ? vk[0].toUpperCase() : "Retourenlabel");
    }
  }
  if (kind === "VERANLASST") parts.push(/genehmigt|freigegeben/i.test(t) ? "Retoure genehmigt" : "Retoure veranlasst");
  if (kind === "RUECKFRAGE") {
    if (/\d+\s+m[öo]glichkeiten/i.test(t)) parts.push("Versandart wählen");
    if (/nennen\s+sie\s+(?:uns\s+)?(?:ein(?:en)?\s+)?abhol/i.test(t)) parts.push("Abholdatum nennen");
    if (/fotos?|bilder/i.test(t)) parts.push("Fotos benötigt");
    if (/sollen\s+wir/i.test(t)) parts.push("Abholung bestätigen");
    if (/w[äa]re\s+(?:nur\s+)?(?:\S+\s+){0,2}m[öo]glich|einverstanden/i.test(t)) parts.push("Bedingungen bestätigen");
    if (!parts.length) parts.push("Rückfrage beantworten");
  }
  if (kind === "PRUEFUNG") {
    if (/angemeldet\s+werden/i.test(t)) parts.push("wird beim Lieferanten angemeldet");
    else if (/nachgefragt/i.test(t)) parts.push("Alfah hat nachgefragt");
    else if (/weitergeleitet/i.test(t)) parts.push("an Lieferanten weitergeleitet");
    else parts.push("in Prüfung");
  }
  if (kind === "RUECK" || kind === "VERANLASST" || kind === "RUECKFRAGE") parts = parts.concat(alfahConditions_(t));
  return parts.join(" · ");
}

function alfahStatusPaint_(kind) {
  if (kind === "ABGELEHNT") return { text: "ABGELEHNT", bg: "#cc0000", fg: "#ffffff" };
  if (kind === "STORNO") return { text: "STORNO", bg: "#cc0000", fg: "#ffffff" };
  if (kind === "RUECK") return { text: "Rücksendung", bg: ALFAH_GREEN, fg: "#ffffff" };
  if (kind === "VERANLASST") return { text: "Rücksendung", bg: ALFAH_ORANGE, fg: "#ffffff" };
  if (kind === "VORBEREITET") return { text: "VORBEREITET", bg: "#f1c232", fg: "#1f2933" };
  if (kind === "RUECKFRAGE") return { text: "RÜCKFRAGE", bg: "#674ea7", fg: "#ffffff" };
  if (kind === "PRUEFUNG") return { text: "IN PRÜFUNG", bg: "#9fc5e8", fg: "#1f2933" };
  if (kind === "ANTWORT") return { text: "ANTWORT PRÜFEN", bg: "#c27ba0", fg: "#ffffff" };
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
  if (plain && alfahMailKind_(alfahFreshBody_(plain))) return plain;
  if (html.length > plain.length * 1.2 && alfahMailKind_(alfahFreshBody_(html))) return html;
  return plain || html;
}

function alfahMailThreads_(query, limit) {
  var threads = [];
  var start = 0;
  var cap = limit || 400;
  while (start < cap) {
    var size = Math.min(100, cap - start);
    var batch = [];
    try { batch = GmailApp.search(query, start, size); } catch (err) { batch = []; }
    if (!batch.length) break;
    var i;
    for (i = 0; i < batch.length; i++) threads.push(batch[i]);
    if (batch.length < size) break;
    start += batch.length;
  }
  return threads;
}

function alfahIdsIn_(text) {
  var list = String(text || "").toUpperCase().match(/R-\d{4,}/g) || [];
  var out = [];
  var seen = {};
  var i;
  for (i = 0; i < list.length; i++) {
    if (seen[list[i]]) continue;
    seen[list[i]] = true;
    out.push(list[i]);
  }
  return out;
}

function alfahRec_(data, id) {
  if (!data.recs[id]) data.recs[id] = { best: null, last: null, lastIn: 0, ask: 0 };
  return data.recs[id];
}

function alfahOurAddresses_() {
  var out = [ALFAH_INBOX.toLowerCase()];
  var box = String(alfahMailbox_() || "").toLowerCase();
  if (box && out.indexOf(box) < 0) out.push(box);
  return out;
}

function alfahFromUs_(from, ours) {
  var i;
  for (i = 0; i < ours.length; i++) {
    if (ours[i] && from.indexOf(ours[i]) >= 0) return true;
  }
  return false;
}

function alfahRememberMail_(rec, when, kind, detail, threadId, text) {
  if (kind) {
    var b = rec.best;
    var keep = b && (b.when > when || (b.when === when && (b.kind === "ABGELEHNT" || b.kind === "STORNO") && kind !== "ABGELEHNT" && kind !== "STORNO"));
    if (!keep) rec.best = { kind: kind, when: when, detail: detail, threadId: threadId, text: text };
  }
  if (!rec.last || when >= rec.last.when) rec.last = { kind: kind, when: when, threadId: threadId, text: text };
}

function alfahReadMessage_(msg, threadId, data, index) {
  var subject = "";
  var when = 0;
  var from = "";
  var to = "";
  try { subject = msg.getSubject() || ""; } catch (e1) {}
  try { when = msg.getDate() ? msg.getDate().getTime() : 0; } catch (e2) {}
  try { from = String(msg.getFrom() || "").toLowerCase(); } catch (e3) {}
  try { to = (String(msg.getTo() || "") + "," + String(msg.getCc() || "")).toLowerCase(); } catch (e4) {}
  var ids = alfahIdsIn_(subject);
  var n;
  if (alfahFromUs_(from, data.ours)) {
    if (!ids.length || !index || to.indexOf("alfah") < 0) return;
    for (n = 0; n < ids.length; n++) {
      var mine = alfahRec_(data, ids[n]);
      if (when > mine.ask) mine.ask = when;
    }
    return;
  }
  var body = "";
  try { body = alfahMessageText_(msg); } catch (e5) {}
  var cut = alfahCutQuote_(body);
  var fresh = alfahFlat_(cut);
  if (!ids.length) ids = alfahIdsIn_(fresh);
  if (!ids.length) return;
  var scan = fresh.length >= 40 ? fresh : alfahFlat_(body);
  var kind = alfahMailKind_(scan);
  for (n = 0; n < ids.length; n++) {
    var seenRec = alfahRec_(data, ids[n]);
    if (when > seenRec.lastIn) seenRec.lastIn = when;
  }
  if (!kind && alfahNoiseMail_(fresh)) return;
  var detail = kind ? alfahMailDetail_(scan, kind) : "";
  var text = alfahMailSnippet_(fresh.length >= 40 ? cut : body);
  for (n = 0; n < ids.length; n++) {
    alfahRememberMail_(alfahRec_(data, ids[n]), when, kind, detail, threadId, text);
  }
}

function alfahReadThreads_(threads, data) {
  var fresh = [];
  var i;
  for (i = 0; i < threads.length; i++) {
    var tid = "";
    try { tid = threads[i].getId(); } catch (err) { tid = ""; }
    if (!tid || data.done[tid]) continue;
    data.done[tid] = true;
    fresh.push(threads[i]);
  }
  var step;
  for (step = 0; step < fresh.length; step += 40) {
    var slice = fresh.slice(step, step + 40);
    var groups = [];
    try { groups = GmailApp.getMessagesForThreads(slice); } catch (err2) { groups = []; }
    var g;
    for (g = 0; g < groups.length; g++) {
      var messages = groups[g] || [];
      var threadId = slice[g] ? slice[g].getId() : "";
      var m;
      for (m = 0; m < messages.length; m++) {
        try { alfahReadMessage_(messages[m], threadId, data, m); } catch (err3) {}
      }
    }
  }
}

function alfahCollectMail_(openIds, started) {
  var props = PropertiesService.getScriptProperties();
  var primed = props.getProperty(ALFAH_MAIL_KEY) === "1";
  var data = { recs: {}, covered: {}, done: {}, ours: alfahOurAddresses_(), complete: true };
  var query = primed ? "subject:RETOURE newer_than:21d" : "subject:RETOURE newer_than:150d";
  alfahReadThreads_(alfahMailThreads_(query, 400), data);
  var need = [];
  var i;
  for (i = 0; i < openIds.length; i++) {
    var rec = data.recs[openIds[i]];
    if (rec && rec.best && ALFAH_FINAL[rec.best.kind] && rec.last && rec.last.when <= rec.best.when) continue;
    need.push(openIds[i]);
  }
  if (!need.length) return data;
  var startAt = Number(props.getProperty(ALFAH_CURSOR_KEY) || 0) % need.length;
  var order = need.slice(startAt).concat(need.slice(0, startAt));
  var processed = 0;
  var b;
  for (b = 0; b < order.length; b += 10) {
    if (Date.now() - started > ALFAH_SCAN_BUDGET_MS) {
      data.complete = false;
      break;
    }
    var chunk = order.slice(b, b + 10);
    var q = "(" + chunk.map(function(id) { return '"' + id + '"'; }).join(" OR ") + ") newer_than:180d";
    alfahReadThreads_(alfahMailThreads_(q, 80), data);
    var c;
    for (c = 0; c < chunk.length; c++) data.covered[chunk[c]] = true;
    processed += chunk.length;
  }
  props.setProperty(ALFAH_CURSOR_KEY, String(processed >= order.length ? 0 : (startAt + processed) % need.length));
  return data;
}

function alfahDecide_(rec) {
  if (!rec) return null;
  var best = rec.best;
  var last = rec.last;
  if (last && !last.kind && (!best || last.when > best.when) && (!best || !ALFAH_FINAL[best.kind])) {
    return { kind: "ANTWORT", when: last.when, detail: "neue Antwort lesen", threadId: last.threadId };
  }
  return best || null;
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
  var q = encodeURIComponent('"' + String(rNumber || "").toUpperCase() + '"');
  return alfahGmailUrl_("#search/" + q);
}

function alfahSearchCurrent_(rich, rNumber) {
  var url = "";
  try { url = rich ? String(rich.getLinkUrl() || "") : ""; } catch (err) { url = ""; }
  return !!url && url === alfahSearchUrl_(rNumber);
}

function alfahRichUrl_(rich) {
  if (!rich) return "";
  var url = "";
  try { url = String(rich.getLinkUrl() || ""); } catch (err) { url = ""; }
  if (url) return url;
  try {
    var runs = rich.getRuns();
    var i;
    for (i = 0; i < runs.length; i++) {
      var u = runs[i].getLinkUrl();
      if (u) return String(u);
    }
  } catch (err2) {}
  return "";
}

function alfahIsRueck_(text) {
  var t = String(text || "");
  return t === "Rücksendung" || t === "Rücksendung eingeleitet";
}

function alfahOurStatus_(text) {
  var t = String(text || "");
  return t === "ABGELEHNT" || t === "STORNO" || alfahIsRueck_(t) || t === "VORBEREITET" || t === "RÜCKFRAGE" || t === "IN PRÜFUNG" || t === "ANTWORT PRÜFEN";
}

function alfahBox_() {
  return String(PropertiesService.getScriptProperties().getProperty(ALFAH_BOX_KEY) || "").trim();
}

function alfahSearchColor_(status) {
  if (status === "ABGELEHNT" || status === "STORNO" || alfahIsRueck_(status) || status === "RÜCKFRAGE" || status === "ANTWORT PRÜFEN") return "#ffffff";
  if (status === "VORBEREITET" || status === "IN PRÜFUNG") return "#1f2933";
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

function alfahChaseTone_(disp, vals, rowBg, now) {
  var st = String(disp[ALFAH_STATUS_COL - 1] || "").trim();
  var bg = String(rowBg || "").toLowerCase();
  if (st === "ABGELEHNT" || st === "STORNO" || st === "VORBEREITET") return "";
  if (alfahIsRueck_(st) && bg !== ALFAH_ORANGE) return "";
  var mailMs = alfahStampMs_(vals[ALFAH_MAILAT_COL - 1], disp[ALFAH_MAILAT_COL - 1]);
  var askMs = alfahStampMs_(vals[ALFAH_ASK_COL - 1], disp[ALFAH_ASK_COL - 1]);
  var baseMs = mailMs || alfahStampMs_("", String(disp[7] || "") + " " + String(disp[8] || ""));
  var limit = ALFAH_CHASE_DAYS * 864e5;
  if (askMs && askMs >= baseMs) return now - askMs >= limit ? "red" : "orange";
  if (baseMs && now - baseMs >= limit) return "red";
  return "";
}

function alfahPaintChase_(sheet) {
  var last = sheet.getLastRow();
  if (last < 3) return 0;
  var count = last - 2;
  var range = sheet.getRange(3, 1, count, ALFAH_LAST_COL);
  var vals = range.getValues();
  var disp = range.getDisplayValues();
  var rowBg = sheet.getRange(3, 1, count, 1).getBackgrounds();
  var cell = sheet.getRange(3, ALFAH_MAILAT_COL, count, 1);
  var bgs = cell.getBackgrounds();
  var fgs = cell.getFontColors();
  var weights = cell.getFontWeights();
  var now = Date.now();
  var due = 0;
  var i;
  for (i = 0; i < count; i++) {
    if (!/R-\d{4,}/i.test(String(disp[i][2] || ""))) continue;
    var tone = alfahChaseTone_(disp[i], vals[i], rowBg[i][0], now);
    if (tone === "red") {
      bgs[i][0] = ALFAH_CHASE_RED;
      fgs[i][0] = "#ffffff";
      weights[i][0] = "bold";
      due++;
    } else if (tone === "orange") {
      bgs[i][0] = ALFAH_CHASE_ORANGE;
      fgs[i][0] = "#ffffff";
      weights[i][0] = "bold";
    } else {
      bgs[i][0] = "#ffffff";
      fgs[i][0] = "#1f2933";
      weights[i][0] = "normal";
    }
  }
  cell.setBackgrounds(bgs);
  cell.setFontColors(fgs);
  cell.setFontWeights(weights);
  return due;
}

function alfahFitExtraCols_(sheet, rows) {
  if (!rows) return;
  var extra = sheet.getRange(3, ALFAH_MAILAT_COL, rows, 3);
  extra.setFontFamily("Arial");
  extra.setFontSize(ALFAH_FONT);
  extra.setVerticalAlignment("middle");
  sheet.getRange(3, ALFAH_MAILAT_COL, rows, 2).setHorizontalAlignment("center").setWrap(true);
  sheet.getRange(3, ALFAH_TEXT_COL, rows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setHorizontalAlignment("left");
}

function alfahScanMail() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var started = Date.now();
    var sheet = alfahSheet_();
    var last = sheet.getLastRow();
    var count = last >= 3 ? last - 2 : 0;
    var cases = count ? sheet.getRange(3, 3, count, 1).getValues() : [];
    var current = count ? sheet.getRange(3, ALFAH_STATUS_COL, count, 2).getValues() : [];
    var openIds = [];
    var openSeen = {};
    var i;
    for (i = 0; i < cases.length; i++) {
      var openId = String(cases[i][0] || "").toUpperCase().match(/R-\d{4,}/);
      if (!openId || openSeen[openId[0]]) continue;
      var openStatus = String(current[i][0] || "");
      if (openStatus === "ABGELEHNT" || openStatus === "STORNO") continue;
      openSeen[openId[0]] = true;
      openIds.push(openId[0]);
    }
    var mail = alfahCollectMail_(openIds, started);
    var props = PropertiesService.getScriptProperties();
    if (!count) {
      props.setProperty(ALFAH_MAIL_KEY, "1");
      props.setProperty("alfah_mail_last_scan_ms", String(Date.now()));
      return { ok: true, updated: 0, mails: Object.keys(mail.recs).length };
    }
    var mailRich = sheet.getRange(3, ALFAH_LINK_COL, count, 1).getRichTextValues();
    var searchRich = sheet.getRange(3, ALFAH_SEARCH_COL, count, 1).getRichTextValues();
    var rowBg = sheet.getRange(3, 1, count, 1).getBackgrounds();
    var extra = sheet.getRange(3, ALFAH_MAILAT_COL, count, 3);
    var extraVals = extra.getValues();
    var extraDisp = extra.getDisplayValues();
    var extraChanged = false;
    var updated = 0;
    var review = 0;
    for (i = 0; i < cases.length; i++) {
      var rowNumber = 3 + i;
      var id = String(cases[i][0] || "").toUpperCase().match(/R-\d{4,}/);
      if (!id) continue;
      var key = id[0];
      var rec = mail.recs[key];
      if (rec && rec.lastIn) {
        var stamp = alfahStampText_(rec.lastIn);
        if (extraDisp[i][0] !== stamp) {
          extraVals[i][0] = stamp;
          extraChanged = true;
        }
      } else if (mail.covered[key] && !String(extraDisp[i][0] || "").trim()) {
        extraVals[i][0] = "keine Mail";
        extraChanged = true;
      }
      if (rec && rec.ask) {
        var oldAsk = alfahStampMs_(extraVals[i][1], extraDisp[i][1]);
        if (rec.ask > oldAsk + 60000) {
          extraVals[i][1] = alfahStampText_(rec.ask);
          extraChanged = true;
        }
      }
      var snippet = rec ? ((rec.last && rec.last.text) || (rec.best && rec.best.text) || "") : "";
      if (snippet && extraDisp[i][2] !== snippet) {
        extraVals[i][2] = snippet;
        extraChanged = true;
      }
      var res = alfahDecide_(rec);
      if (!res) {
        if (!alfahSearchCurrent_(searchRich[i][0], key)) alfahPutSearch_(sheet, rowNumber, key, "", alfahSearchColor_(String(current[i][0] || "")));
        continue;
      }
      var paint = alfahStatusPaint_(res.kind);
      if (!paint) continue;
      var curStatus = String(current[i][0] || "").trim();
      var curInfo = String(current[i][1] || "").trim();
      var curBg = String(rowBg[i][0] || "").toLowerCase();
      var manualGreen = curInfo === "Alfah System" && alfahIsRueck_(curStatus);
      if (manualGreen && res.kind !== "RUECK" && res.kind !== "ABGELEHNT" && res.kind !== "STORNO") continue;
      if (res.kind === "ANTWORT") {
        var bestPaint = rec.best ? alfahStatusPaint_(rec.best.kind) : null;
        if (curStatus && curStatus !== paint.text && curStatus !== (bestPaint ? bestPaint.text : "")) continue;
        review++;
      }
      var info = alfahText_(res.detail);
      var mailUrl = alfahThreadUrl_(res.threadId);
      if (curStatus === paint.text && curInfo === info && curBg === paint.bg && alfahRichUrl_(mailRich[i][0]) === mailUrl) continue;
      alfahPaintRow_(sheet, rowNumber, paint, info, mailUrl, alfahSearchUrl_(key));
      updated++;
    }
    if (extraChanged) {
      var r;
      var c;
      for (r = 0; r < extraVals.length; r++) {
        for (c = 0; c < 3; c++) {
          var v = extraVals[r][c];
          if (Object.prototype.toString.call(v) === "[object Date]") extraVals[r][c] = isNaN(v.getTime()) ? "" : alfahStampText_(v.getTime());
        }
      }
      extra.setNumberFormat("@");
      extra.setValues(extraVals);
    }
    props.setProperty(ALFAH_MAIL_KEY, "1");
    props.setProperty("alfah_mail_last_scan_ms", String(Date.now()));
    var moved = alfahMoveRejected_(sheet);
    var delivered = alfahMoveDelivered_(sheet);
    var due = alfahPaintChase_(sheet);
    var oem = null;
    try { oem = alfahSyncOemTeile(); } catch (errOem) { oem = { ok: false, error: String(errOem && errOem.message ? errOem.message : errOem) }; }
    return { ok: true, updated: updated, review: review, due: due, complete: mail.complete, moved: moved.moved, delivered: delivered.moved, mails: Object.keys(mail.recs).length, oem: oem };
  } finally {
    lock.releaseLock();
  }
}

function alfahSideTab_(ss, gid, name) {
  var sheets = ss.getSheets();
  var i;
  for (i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }
  var named = ss.getSheetByName(name);
  if (named) return named;
  return ss.insertSheet(name);
}

function alfahPrepareReject_(source) {
  return alfahPrepareSide_(source, ALFAH_REJECT_GID, ALFAH_REJECT_TAB, ALFAH_REJECT_TITLE, ALFAH_REJECT_RED);
}

function alfahPrepareDone_(source) {
  return alfahPrepareSide_(source, ALFAH_DONE_GID, ALFAH_DONE_TAB, ALFAH_DONE_TITLE, ALFAH_GREEN);
}

function alfahPrepareSide_(source, gid, name, title, color) {
  var sheet = alfahSideTab_(source.getParent(), gid, name);
  sheet.setTabColor(color);
  alfahEnsureCols_(sheet);
  var headers = source.getRange(2, 1, 1, ALFAH_LAST_COL).getValues();
  if (String(sheet.getRange(1, 1).getValue() || "")) {
    if (String(sheet.getRange(2, ALFAH_LAST_COL).getValue() || "")) return sheet;
    var extraHead = sheet.getRange(2, ALFAH_MAILAT_COL, 1, ALFAH_LAST_COL - ALFAH_MAILAT_COL + 1);
    extraHead.setValues([headers[0].slice(ALFAH_MAILAT_COL - 1)]);
    extraHead.setBackground(color).setFontColor("#ffffff").setFontWeight("bold").setFontSize(ALFAH_FONT).setFontFamily("Arial").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
    var w;
    for (w = ALFAH_MAILAT_COL; w <= ALFAH_LAST_COL; w++) sheet.setColumnWidth(w, ALFAH_WIDTHS[w - 1]);
    return sheet;
  }
  var titleRange = sheet.getRange(1, 1, 1, ALFAH_LAST_COL);
  if (titleRange.isPartOfMerge()) titleRange.breakApart();
  titleRange.merge();
  titleRange
    .setValue(title)
    .setBackground(color)
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(11)
    .setFontFamily("Arial")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
  var head = sheet.getRange(2, 1, 1, ALFAH_LAST_COL);
  head.setValues(headers);
  head.setBackground(color);
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
  return sheet;
}

function alfahRowBlank_(sheet, rowNumber) {
  if (rowNumber < 1 || rowNumber > sheet.getLastRow()) return false;
  var vals = sheet.getRange(rowNumber, 1, 1, ALFAH_LAST_COL).getDisplayValues()[0];
  var i;
  for (i = 0; i < vals.length; i++) {
    if (String(vals[i] || "").trim()) return false;
  }
  return true;
}

function alfahMoveRejected_(sheet) {
  var last = sheet.getLastRow();
  if (last < 3) return { moved: 0, removed: 0 };
  var count = last - 2;
  var status = sheet.getRange(3, ALFAH_STATUS_COL, count, 1).getDisplayValues();
  var bg = sheet.getRange(3, 1, count, 1).getBackgrounds();
  var rows = [];
  var i;
  for (i = 0; i < count; i++) {
    var state = String(status[i][0] || "").toUpperCase();
    var color = String(bg[i][0] || "").toLowerCase();
    if (state === "ABGELEHNT" || state === "STORNO" || color === ALFAH_REJECT_RED) rows.push(3 + i);
  }
  if (!rows.length) return { moved: 0, removed: 0 };
  return alfahMoveRows_(sheet, rows, alfahPrepareReject_(sheet), ALFAH_REJECT_RED);
}

function alfahMoveDelivered_(sheet) {
  var last = sheet.getLastRow();
  if (last < 3) return { moved: 0, removed: 0 };
  var count = last - 2;
  var status = sheet.getRange(3, ALFAH_STATUS_COL, count, 1).getDisplayValues();
  var bg = sheet.getRange(3, 1, count, 1).getBackgrounds();
  var boxes = sheet.getRange(3, ALFAH_PREP_COL, count, 2).getValues();
  var rows = [];
  var i;
  for (i = 0; i < count; i++) {
    if (!alfahIsRueck_(status[i][0])) continue;
    if (String(bg[i][0] || "").toLowerCase() !== ALFAH_GREEN) continue;
    if (!wmChecked_(boxes[i][0]) || !wmChecked_(boxes[i][1])) continue;
    rows.push(3 + i);
  }
  if (!rows.length) return { moved: 0, removed: 0 };
  return alfahMoveRows_(sheet, rows, alfahPrepareDone_(sheet), "");
}

function alfahMoveRows_(sheet, rows, dest, paint) {
  var pairs = sheet.getRange(3, 3, sheet.getLastRow() - 2, 2).getValues();
  var have = {};
  var i;
  var destLast = dest.getLastRow();
  if (destLast >= 3) {
    var existing = dest.getRange(3, 3, destLast - 2, 2).getValues();
    for (i = 0; i < existing.length; i++) {
      var known = String(existing[i][0] || "").toUpperCase() + "|" + String(existing[i][1] || "").toUpperCase();
      if (known !== "|") have[known] = true;
    }
  }
  var cursor = 3;
  if (destLast >= 3) {
    var tail = dest.getRange(destLast, 1, 1, 4).getDisplayValues()[0].join("");
    cursor = String(tail).trim() ? destLast + 2 : destLast + 1;
  }
  var moved = 0;
  for (i = 0; i < rows.length; i++) {
    var rowNumber = rows[i];
    var key = String(pairs[rowNumber - 3][0] || "").toUpperCase() + "|" + String(pairs[rowNumber - 3][1] || "").toUpperCase();
    if (key !== "|" && have[key]) continue;
    sheet.getRange(rowNumber, 1, 1, ALFAH_LAST_COL).copyTo(dest.getRange(cursor, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    if (paint) dest.getRange(cursor, 1, 1, ALFAH_LAST_COL).setBackground(paint);
    try { dest.setRowHeight(cursor, sheet.getRowHeight(rowNumber)); } catch (err) {}
    if (key !== "|") have[key] = true;
    moved++;
    cursor += 2;
  }
  for (i = rows.length - 1; i >= 0; i--) {
    var drop = rows[i];
    var above = drop <= 3 || alfahRowBlank_(sheet, drop - 1);
    var span = above && alfahRowBlank_(sheet, drop + 1) ? 2 : 1;
    sheet.deleteRows(drop, span);
  }
  return { moved: moved, removed: rows.length };
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
        var rowBg = String(sheet.getRange(rows[r], 1).getBackground() || "").toLowerCase();
        if (alfahIsRueck_(status) && info && info !== "Alfah System" && rowBg !== ALFAH_ORANGE) {
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

var WM_GID = 891953845;
var WM_TAB = "WM";
var WM_TITLE = "WM Retouren";
var WM_GREEN = "#18884a";
var WM_HEADERS = ["Auftrag", "Stock-ID", "Teilenummer", "Name", "Lieferant", "Hersteller", "Datum", "Uhrzeit", "Wer", "Info"];
var WM_WIDTHS = [72, 78, 132, 120, 72, 84, 92, 76, 72, 96];

function wmBook_() {
  return SpreadsheetApp.openById(ALFAH_SHEET_ID);
}

function wmTab_(ss) {
  var sheets = ss.getSheets();
  var i;
  for (i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === WM_GID) return sheets[i];
  }
  var named = ss.getSheetByName(WM_TAB);
  if (named) return named;
  return ss.insertSheet(WM_TAB);
}

function wmSheet_() {
  var sheet = wmTab_(wmBook_());
  sheet.setTabColor(WM_GREEN);
  wmFitLayout_(sheet);
  return sheet;
}

function wmOldLayout_(sheet) {
  var width = Math.min(Math.max(sheet.getLastColumn(), 1), 16);
  var head = sheet.getRange(2, 1, 1, width).getDisplayValues()[0];
  return String(head[2] || "") === "R-Nummer" || String(head[10] || "") === "Status" || String(head[15] || "") === "Geliefert";
}

function wmChecked_(value) {
  if (value === true) return true;
  var text = String(value || "").toUpperCase();
  return text === "TRUE" || text === "WAHR";
}

function wmMigrate_(sheet) {
  if (!wmOldLayout_(sheet)) return;
  var last = sheet.getLastRow();
  if (last < 3) return;
  var count = last - 2;
  var values = sheet.getRange(3, 1, count, 16).getValues();
  var next = [];
  var i;
  for (i = 0; i < values.length; i++) {
    var cells = values[i];
    var part = String(cells[0] || "") || String(cells[1] || "") || String(cells[3] || "") || String(cells[4] || "");
    if (!part) {
      next.push(["", "", "", "", "", "", "", "", "", "", ""]);
      continue;
    }
    next.push([
      cells[0],
      cells[1],
      cells[3],
      cells[4],
      cells[5],
      cells[6],
      cells[7],
      cells[8],
      cells[9],
      cells[11],
      wmChecked_(cells[14])
    ]);
  }
  var body = sheet.getRange(3, 1, count, 16);
  body.clearDataValidations();
  body.clearContent();
  body.setBorder(false, false, false, false, false, false);
  var fresh = sheet.getRange(3, 1, count, WM_HEADERS.length);
  fresh.setValues(next);
  sheet.getRange(3, 1, count, WM_HEADERS.length - 1).setNumberFormat("@");
}

function wmUnmerge_(range) {
  var merges = range.getMergedRanges();
  var i;
  for (i = 0; i < merges.length; i++) merges[i].breakApart();
}

function wmClearExtra_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), WM_HEADERS.length);
  var height = Math.max(sheet.getLastRow(), 2);
  wmUnmerge_(sheet.getRange(1, 1, height, lastCol));
  if (lastCol <= WM_HEADERS.length) return;
  var extra = sheet.getRange(1, WM_HEADERS.length + 1, height, lastCol - WM_HEADERS.length);
  extra.clear();
  extra.clearDataValidations();
  extra.setBorder(false, false, false, false, false, false);
}

function wmPaintRow_(sheet, rowNumber) {
  var rowRange = sheet.getRange(rowNumber, 1, 1, WM_HEADERS.length);
  rowRange.setBackground("#ffffff");
  rowRange.setFontColor("#1f2933");
  rowRange.setFontFamily("Arial");
  rowRange.setFontSize(ALFAH_FONT);
  rowRange.setVerticalAlignment("middle");
  rowRange.setWrap(true);
  rowRange.setBorder(true, true, true, true, true, true, "#d6deea", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(rowNumber, 2).setFontWeight("bold");
  sheet.getRange(rowNumber, 3).setFontWeight("bold").setFontColor(WM_GREEN);
  sheet.autoResizeRows(rowNumber, 1);
}

function wmStyleBody_(sheet) {
  var last = sheet.getLastRow();
  if (last < 3) return;
  var count = last - 2;
  var keys = sheet.getRange(3, 1, count, 4).getValues();
  var i;
  for (i = 0; i < keys.length; i++) {
    var rowNumber = 3 + i;
    var has = !!(String(keys[i][0] || "") || String(keys[i][1] || "") || String(keys[i][2] || "") || String(keys[i][3] || ""));
    var rowRange = sheet.getRange(rowNumber, 1, 1, WM_HEADERS.length);
    if (!has) {
      rowRange.setBackground(null);
      rowRange.setBorder(false, false, false, false, false, false);
      continue;
    }
    wmPaintRow_(sheet, rowNumber);
  }
}

function wmSeedIfEmpty_(sheet) {
  if (sheet.getLastRow() >= 3) return;
  if (PropertiesService.getScriptProperties().getProperty("wm_seed_hm62348") === "1") return;
  var row = ["944668", "HM62348", "F 026 407 279", "ÖLFILTER FORD", "WM", "BOH", "25.10.2026", "21:01", "Berger", "1 Stk"];
  var range = sheet.getRange(3, 1, 1, row.length);
  range.setNumberFormat("@");
  range.setValues([row]);
  wmPaintRow_(sheet, 3);
  PropertiesService.getScriptProperties().setProperty("wm_seed_hm62348", "1");
}

function wmFitLayout_(sheet) {
  wmMigrate_(sheet);
  wmClearExtra_(sheet);
  var titleRange = sheet.getRange(1, 1, 1, WM_HEADERS.length);
  wmUnmerge_(titleRange);
  titleRange.merge();
  titleRange
    .setValue(WM_TITLE)
    .setBackground(WM_GREEN)
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(11)
    .setFontFamily("Arial")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setWrap(true);
  var head = sheet.getRange(2, 1, 1, WM_HEADERS.length);
  head.setValues([WM_HEADERS]);
  head.setBackground(WM_GREEN);
  head.setFontColor("#ffffff");
  head.setFontWeight("bold");
  head.setFontSize(ALFAH_FONT);
  head.setFontFamily("Arial");
  head.setHorizontalAlignment("center");
  head.setVerticalAlignment("middle");
  head.setWrap(true);
  sheet.setFrozenRows(2);
  sheet.setRowHeight(1, 22);
  var c;
  for (c = 0; c < WM_WIDTHS.length; c++) sheet.setColumnWidth(c + 1, WM_WIDTHS[c]);
  sheet.autoResizeRows(2, 1);
  wmStyleBody_(sheet);
  wmSeedIfEmpty_(sheet);
}

function wmText_(value) {
  return alfahText_(value);
}

function wmStock_(value) {
  var match = String(value || "").toUpperCase().match(/[A-Z]{2,3}\d{4,8}/);
  return match ? match[0] : "";
}

function wmInfo_(line) {
  var bits = [];
  var qty = String(line.qty == null ? "" : line.qty).replace(/[^\d]/g, "");
  if (qty) bits.push(qty + " Stk");
  var ean = wmText_(line.ean);
  var article = wmText_(line.article || line.part || "");
  if (!article) article = ean;
  if (ean && ean !== article) bits.push("EAN " + ean);
  return bits.join(" · ");
}

function wmStamp_() {
  var text = Utilities.formatDate(new Date(), "Europe/Berlin", "dd.MM.yyyy HH:mm");
  var parts = text.split(" ");
  return { date: parts[0] || "", time: parts[1] || "" };
}

function wmHasToken_(token) {
  var id = String(token || "");
  if (!id) return false;
  var raw = String(PropertiesService.getScriptProperties().getProperty("wm_retoure_tokens") || "");
  return ("," + raw + ",").indexOf("," + id + ",") !== -1;
}

function wmRememberToken_(token) {
  var id = String(token || "");
  if (!id || wmHasToken_(id)) return;
  var props = PropertiesService.getScriptProperties();
  var raw = String(props.getProperty("wm_retoure_tokens") || "");
  var list = raw ? raw.split(",") : [];
  list.push(id);
  if (list.length > 250) list = list.slice(list.length - 250);
  props.setProperty("wm_retoure_tokens", list.join(","));
}

function wmStartRow_(sheet) {
  var last = sheet.getLastRow();
  if (last < 3) return 3;
  var filled = sheet.getRange(last, 1, 1, 5).getDisplayValues()[0].join("");
  if (!filled) return last + 1;
  return last + 2;
}

function wmAppend_(entries) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = wmSheet_();
    var added = 0;
    var skipped = 0;
    var cursor = wmStartRow_(sheet);
    var e;
    for (e = 0; e < entries.length; e++) {
      var lines = entries[e] && entries[e].lines ? entries[e].lines : [entries[e]];
      var n;
      for (n = 0; n < lines.length; n++) {
        var line = lines[n] || {};
        var token = String(line.token || "");
        if (token && wmHasToken_(token)) {
          skipped++;
          continue;
        }
        var article = wmText_(line.article || line.part || "");
        var ean = wmText_(line.ean);
        if (!article) article = ean;
        var name = wmText_(line.name);
        if (!article && !name) {
          skipped++;
          continue;
        }
        var stamp = wmStamp_();
        if (line.date && line.time) {
          stamp = { date: wmText_(line.date), time: wmText_(line.time) };
        }
        var row = [
          wmText_(line.orderId),
          wmStock_(line.stockId),
          article,
          name,
          wmText_(line.supplier) || "WM",
          wmText_(line.brand),
          stamp.date,
          stamp.time,
          wmText_(line.who),
          wmInfo_(line)
        ];
        var range = sheet.getRange(cursor, 1, 1, row.length);
        range.setNumberFormat("@");
        range.setValues([row]);
        wmPaintRow_(sheet, cursor);
        if (token) wmRememberToken_(token);
        added++;
        cursor += 2;
      }
    }
    return { ok: true, added: added, skipped: skipped, tab: WM_TAB };
  } finally {
    lock.releaseLock();
  }
}

var AAG_GID = 0;
var AAG_TAB = "KNOLL / AAG";
var AAG_TITLE = "KNOLL / AAG Retouren";
var AAG_NAVY = "#00244e";
var AAG_HEADERS = ["Auftrag", "Stock-ID", "Teilenummer", "Name", "Lieferant", "Hersteller", "Datum", "Uhrzeit", "Wer", "Info"];
var AAG_WIDTHS = [72, 78, 132, 120, 72, 84, 92, 76, 72, 96];

function aagBook_() {
  return SpreadsheetApp.openById(ALFAH_SHEET_ID);
}

function aagTab_(ss) {
  var named = ss.getSheetByName(AAG_TAB);
  if (named) return named;
  var sheets = ss.getSheets();
  var i;
  for (i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === AAG_GID) return sheets[i];
  }
  return ss.insertSheet(AAG_TAB);
}

function aagSheet_() {
  var sheet = aagTab_(aagBook_());
  sheet.setTabColor(AAG_NAVY);
  aagFitLayout_(sheet);
  return sheet;
}

function aagClearExtra_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), AAG_HEADERS.length);
  var height = Math.max(sheet.getLastRow(), 2);
  wmUnmerge_(sheet.getRange(1, 1, height, lastCol));
  if (lastCol <= AAG_HEADERS.length) return;
  var extra = sheet.getRange(1, AAG_HEADERS.length + 1, height, lastCol - AAG_HEADERS.length);
  extra.clear();
  extra.clearDataValidations();
  extra.setBorder(false, false, false, false, false, false);
}

function aagPaintRow_(sheet, rowNumber) {
  var rowRange = sheet.getRange(rowNumber, 1, 1, AAG_HEADERS.length);
  rowRange.setBackground("#ffffff");
  rowRange.setFontColor("#1f2933");
  rowRange.setFontFamily("Arial");
  rowRange.setFontSize(ALFAH_FONT);
  rowRange.setVerticalAlignment("middle");
  rowRange.setWrap(true);
  rowRange.setBorder(true, true, true, true, true, true, "#d6deea", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(rowNumber, 2).setFontWeight("bold");
  sheet.getRange(rowNumber, 3).setFontWeight("bold").setFontColor(AAG_NAVY);
  sheet.autoResizeRows(rowNumber, 1);
}

function aagFitLayout_(sheet) {
  aagClearExtra_(sheet);
  var titleRange = sheet.getRange(1, 1, 1, AAG_HEADERS.length);
  wmUnmerge_(titleRange);
  titleRange.merge();
  titleRange
    .setValue(AAG_TITLE)
    .setBackground(AAG_NAVY)
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(11)
    .setFontFamily("Arial")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setWrap(true);
  var head = sheet.getRange(2, 1, 1, AAG_HEADERS.length);
  head.setValues([AAG_HEADERS]);
  head.setBackground(AAG_NAVY);
  head.setFontColor("#ffffff");
  head.setFontWeight("bold");
  head.setFontSize(ALFAH_FONT);
  head.setFontFamily("Arial");
  head.setHorizontalAlignment("center");
  head.setVerticalAlignment("middle");
  head.setWrap(true);
  sheet.setFrozenRows(2);
  sheet.setRowHeight(1, 22);
  var c;
  for (c = 0; c < AAG_WIDTHS.length; c++) sheet.setColumnWidth(c + 1, AAG_WIDTHS[c]);
  sheet.autoResizeRows(2, 1);
}

function aagHasToken_(token) {
  var id = String(token || "");
  if (!id) return false;
  var raw = String(PropertiesService.getScriptProperties().getProperty("aag_retoure_tokens") || "");
  return ("," + raw + ",").indexOf("," + id + ",") !== -1;
}

function aagRememberToken_(token) {
  var id = String(token || "");
  if (!id || aagHasToken_(id)) return;
  var props = PropertiesService.getScriptProperties();
  var raw = String(props.getProperty("aag_retoure_tokens") || "");
  var list = raw ? raw.split(",") : [];
  list.push(id);
  if (list.length > 250) list = list.slice(list.length - 250);
  props.setProperty("aag_retoure_tokens", list.join(","));
}

function aagAppend_(entries) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = aagSheet_();
    var added = 0;
    var skipped = 0;
    var cursor = wmStartRow_(sheet);
    var e;
    for (e = 0; e < entries.length; e++) {
      var lines = entries[e] && entries[e].lines ? entries[e].lines : [entries[e]];
      var n;
      for (n = 0; n < lines.length; n++) {
        var line = lines[n] || {};
        var token = String(line.token || "");
        if (token && aagHasToken_(token)) {
          skipped++;
          continue;
        }
        var article = wmText_(line.article || line.part || "");
        var ean = wmText_(line.ean);
        if (!article) article = ean;
        var name = wmText_(line.name);
        if (!article && !name) {
          skipped++;
          continue;
        }
        var stamp = wmStamp_();
        if (line.date && line.time) stamp = { date: wmText_(line.date), time: wmText_(line.time) };
        var row = [
          wmText_(line.orderId),
          wmStock_(line.stockId),
          article,
          name,
          wmText_(line.supplier) || "AAG",
          wmText_(line.brand),
          stamp.date,
          stamp.time,
          wmText_(line.who),
          wmInfo_(line)
        ];
        var range = sheet.getRange(cursor, 1, 1, row.length);
        range.setNumberFormat("@");
        range.setValues([row]);
        aagPaintRow_(sheet, cursor);
        if (token) aagRememberToken_(token);
        added++;
        cursor += 2;
      }
    }
    return { ok: true, added: added, skipped: skipped, tab: AAG_TAB };
  } finally {
    lock.releaseLock();
  }
}

var STA_GID = 441953307;
var STA_TAB = "STA";
var STA_TITLE = "STA Retouren";
var STA_BLUE = "#1d4ed8";
var STA_HEADERS = ["Auftrag", "Stock-ID", "Teilenummer", "Name", "Lieferant", "Hersteller", "Datum", "Uhrzeit", "Wer", "Info"];
var STA_WIDTHS = [110, 90, 140, 180, 90, 110, 92, 76, 90, 160];

function staBook_() {
  return SpreadsheetApp.openById(ALFAH_SHEET_ID);
}

function staTab_(ss) {
  var sheets = ss.getSheets();
  var i;
  for (i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === STA_GID) return sheets[i];
  }
  var named = ss.getSheetByName(STA_TAB);
  if (named) return named;
  return ss.insertSheet(STA_TAB);
}

function staSheet_() {
  var sheet = staTab_(staBook_());
  sheet.setTabColor(STA_BLUE);
  staFitLayout_(sheet);
  return sheet;
}

function staPaintRow_(sheet, rowNumber) {
  var rowRange = sheet.getRange(rowNumber, 1, 1, STA_HEADERS.length);
  rowRange.setBackground("#ffffff");
  rowRange.setFontColor("#1f2933");
  rowRange.setFontFamily("Arial");
  rowRange.setFontSize(ALFAH_FONT);
  rowRange.setVerticalAlignment("middle");
  rowRange.setWrap(true);
  rowRange.setBorder(true, true, true, true, true, true, "#d6deea", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(rowNumber, 2).setFontWeight("bold");
  sheet.getRange(rowNumber, 3).setFontWeight("bold").setFontColor(STA_BLUE);
  sheet.autoResizeRows(rowNumber, 1);
}

function staFitLayout_(sheet) {
  aagClearExtra_(sheet);
  var titleRange = sheet.getRange(1, 1, 1, STA_HEADERS.length);
  wmUnmerge_(titleRange);
  titleRange.merge();
  titleRange
    .setValue(STA_TITLE)
    .setBackground(STA_BLUE)
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(11)
    .setFontFamily("Arial")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
  var head = sheet.getRange(2, 1, 1, STA_HEADERS.length);
  head.setValues([STA_HEADERS]);
  head.setBackground(STA_BLUE);
  head.setFontColor("#ffffff");
  head.setFontWeight("bold");
  head.setFontSize(ALFAH_FONT);
  head.setFontFamily("Arial");
  head.setHorizontalAlignment("center");
  head.setVerticalAlignment("middle");
  head.setWrap(true);
  sheet.setFrozenRows(2);
  sheet.setRowHeight(1, 22);
  var c;
  for (c = 0; c < STA_WIDTHS.length; c++) sheet.setColumnWidth(c + 1, STA_WIDTHS[c]);
}

function staHasToken_(token) {
  var id = String(token || "");
  if (!id) return false;
  var raw = String(PropertiesService.getScriptProperties().getProperty("sta_retoure_tokens") || "");
  return ("," + raw + ",").indexOf("," + id + ",") !== -1;
}

function staRememberToken_(token) {
  var id = String(token || "");
  if (!id || staHasToken_(id)) return;
  var props = PropertiesService.getScriptProperties();
  var raw = String(props.getProperty("sta_retoure_tokens") || "");
  var list = raw ? raw.split(",") : [];
  list.push(id);
  if (list.length > 250) list = list.slice(list.length - 250);
  props.setProperty("sta_retoure_tokens", list.join(","));
}

function staAppend_(entries) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = staSheet_();
    var added = 0;
    var skipped = 0;
    var cursor = wmStartRow_(sheet);
    var e;
    for (e = 0; e < entries.length; e++) {
      var lines = entries[e] && entries[e].lines ? entries[e].lines : [entries[e]];
      var n;
      for (n = 0; n < lines.length; n++) {
        var line = lines[n] || {};
        var token = String(line.token || "");
        if (token && staHasToken_(token)) {
          skipped++;
          continue;
        }
        var article = wmText_(line.article || line.part || "");
        var ean = wmText_(line.ean);
        if (!article) article = ean;
        var name = wmText_(line.name);
        if (!article && !name) {
          skipped++;
          continue;
        }
        var stamp = wmStamp_();
        if (line.date && line.time) stamp = { date: wmText_(line.date), time: wmText_(line.time) };
        var row = [
          wmText_(line.orderId),
          wmStock_(line.stockId),
          article,
          name,
          wmText_(line.supplier) || "STAKIS",
          wmText_(line.brand),
          stamp.date,
          stamp.time,
          wmText_(line.who),
          wmInfo_(line)
        ];
        var range = sheet.getRange(cursor, 1, 1, row.length);
        range.setNumberFormat("@");
        range.setValues([row]);
        staPaintRow_(sheet, cursor);
        if (token) staRememberToken_(token);
        added++;
        cursor += 2;
      }
    }
    return { ok: true, added: added, skipped: skipped, tab: STA_TAB };
  } finally {
    lock.releaseLock();
  }
}

var OEM_SHEET_ID = "13Oh7gDT8NAul2s0cwQUeaGwMcS3B2MYu0QOdFNMhXzM";
var OEM_GID = 1905682522;
var OEM_TAB = "OEM_Teile";
var OEM_REPORT_MAIL = "francesco.berger@auto1.com";
var OEM_AUTO_RE = /^Alfah (R-\d{4,}|ohne R-Nr)\b/;
var OEM_AUTO_STATUS = { "": true, "offen": true, "in klärung": true, "zurückgegeben": true, "nicht rückgabefähig": true, "storniert": true };

function oemFindTab_(ss, gid, name) {
  var sheets = ss.getSheets();
  var i;
  for (i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }
  return ss.getSheetByName(name);
}

function oemNorm_(value) {
  return String(value || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function oemPartKeys_(raw) {
  var text = String(raw || "").toUpperCase();
  var keys = [];
  var seen = {};
  function add(part) {
    var k = oemNorm_(part);
    if (k.length < 5 || !/\d/.test(k) || seen[k]) return;
    seen[k] = true;
    keys.push(k);
  }
  var brackets = text.match(/\[[^\]]+\]/g) || [];
  var i;
  for (i = 0; i < brackets.length; i++) add(brackets[i]);
  add(text.replace(/\[[^\]]*\]/g, " ").replace(/\([^)]*\)/g, " "));
  var tokens = text.split(/[\s\/,;:()\[\]]+/);
  for (i = 0; i < tokens.length; i++) add(tokens[i]);
  add(text);
  return keys;
}

function oemPartsMatch_(a, b) {
  var i;
  var j;
  for (i = 0; i < a.length; i++) {
    for (j = 0; j < b.length; j++) {
      var x = a[i];
      var y = b[j];
      if (x === y) return true;
      if (x.length >= 6 && y.length >= 6 && (x.indexOf(y) !== -1 || y.indexOf(x) !== -1)) return true;
    }
  }
  return false;
}

function oemResult_(kind, status, info, bg, rNumber, date) {
  var st = String(status || "").trim().toUpperCase();
  var color = String(bg || "").toLowerCase();
  var tag = "Alfah " + (rNumber || "ohne R-Nr");
  var infoTxt = String(info || "").trim();
  var extra = infoTxt ? " · " + infoTxt : "";
  if (st === "STORNO") return { status: "storniert", rank: 3, note: tag + " storniert" + extra };
  if (kind === "reject" || st === "ABGELEHNT") return { status: "nicht rückgabefähig", rank: 3, note: tag + " abgelehnt" + extra };
  if (kind === "done" || (alfahIsRueck_(status) && color === ALFAH_GREEN)) return { status: "zurückgegeben", rank: 3, note: tag + " zurückgeschickt" + extra };
  var step = st === "VORBEREITET" ? "vorbereitet" : (alfahIsRueck_(status) ? "Rücksendung veranlasst" : "angemeldet");
  return { status: "in Klärung", rank: 1, note: tag + " " + step + (date ? " " + date : "") + " · warten auf Alfah" };
}

function oemRetoureEntries_() {
  var ss = alfahBook_();
  var tabs = [
    { gid: ALFAH_GID, name: ALFAH_TAB, kind: "open" },
    { gid: ALFAH_REJECT_GID, name: ALFAH_REJECT_TAB, kind: "reject" },
    { gid: ALFAH_DONE_GID, name: ALFAH_DONE_TAB, kind: "done" }
  ];
  var byStock = {};
  var t;
  for (t = 0; t < tabs.length; t++) {
    var sheet = oemFindTab_(ss, tabs[t].gid, tabs[t].name);
    if (!sheet) continue;
    var last = sheet.getLastRow();
    if (last < 3) continue;
    var count = last - 2;
    var vals = sheet.getRange(3, 1, count, ALFAH_LAST_COL).getDisplayValues();
    var bgs = sheet.getRange(3, 1, count, 1).getBackgrounds();
    var i;
    for (i = 0; i < vals.length; i++) {
      var v = vals[i];
      var stock = alfahStock_(v[1]);
      var keys = oemPartKeys_(v[3]);
      if (!stock || !keys.length) continue;
      var r = String(v[2] || "").toUpperCase().match(/R-\d{4,}/);
      var res = oemResult_(tabs[t].kind, v[ALFAH_STATUS_COL - 1], v[ALFAH_INFO_COL - 1], bgs[i][0], r ? r[0] : "", String(v[7] || "").trim());
      if (!byStock[stock]) byStock[stock] = [];
      byStock[stock].push({ keys: keys, res: res });
    }
  }
  return byStock;
}

function oemFindCol_(headers, test) {
  var i;
  for (i = 0; i < headers.length; i++) {
    if (test(String(headers[i] || "").trim())) return i;
  }
  return -1;
}

function alfahSyncOemTeile() {
  var sheet = oemFindTab_(SpreadsheetApp.openById(OEM_SHEET_ID), OEM_GID, OEM_TAB);
  if (!sheet) return { ok: false, error: "Tab " + OEM_TAB + " nicht gefunden" };
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 3 || lastCol < 1) return { ok: true, matched: 0, updated: 0, kept: 0 };
  var top = sheet.getRange(1, 1, Math.min(6, lastRow), lastCol).getDisplayValues();
  var headerIdx = -1;
  var h;
  for (h = 0; h < top.length; h++) {
    if (oemFindCol_(top[h], function (s) { return s === "Rückgabe-Status"; }) !== -1) { headerIdx = h; break; }
  }
  if (headerIdx === -1) return { ok: false, error: "Spalte Rückgabe-Status nicht gefunden" };
  var head = top[headerIdx];
  var stockCol = oemFindCol_(head, function (s) { return /^stock/i.test(s); });
  var oeCol = oemFindCol_(head, function (s) { return /^oe\b/i.test(s) || /art\.?-?\s*nr/i.test(s); });
  var statusCol = oemFindCol_(head, function (s) { return s === "Rückgabe-Status"; });
  var noteCol = oemFindCol_(head, function (s) { return /^notiz/i.test(s); });
  if (stockCol === -1 || oeCol === -1 || noteCol === -1) return { ok: false, error: "Spalten Stock-Nr / OE / Notiz nicht gefunden" };
  var start = headerIdx + 2;
  if (lastRow < start) return { ok: true, matched: 0, updated: 0, kept: 0 };
  var data = sheet.getRange(start, 1, lastRow - start + 1, lastCol).getDisplayValues();
  var entries = oemRetoureEntries_();
  var matched = 0;
  var updated = 0;
  var kept = 0;
  var changes = [];
  var i;
  for (i = 0; i < data.length; i++) {
    var row = data[i];
    var stock = alfahStock_(row[stockCol]);
    var list = stock ? entries[stock] : null;
    if (!list) continue;
    var keys = oemPartKeys_(row[oeCol]);
    if (!keys.length) continue;
    var best = null;
    var e;
    for (e = 0; e < list.length; e++) {
      if (!oemPartsMatch_(keys, list[e].keys)) continue;
      if (!best || list[e].res.rank >= best.rank) best = list[e].res;
    }
    if (!best) continue;
    matched++;
    var curStatus = String(row[statusCol] || "").trim();
    var curNote = String(row[noteCol] || "").trim();
    var ours = OEM_AUTO_RE.test(curNote);
    var free = !curStatus || curStatus.toLowerCase() === "offen";
    if (!free && !(ours && OEM_AUTO_STATUS[curStatus.toLowerCase()])) {
      kept++;
      continue;
    }
    var cut = curNote.indexOf(" | ");
    var manual = ours ? (cut !== -1 ? curNote.substring(cut + 3) : "") : curNote;
    var nextNote = best.note + (manual ? " | " + manual : "");
    if (curStatus === best.status && curNote === nextNote) continue;
    var rowNumber = start + i;
    if (alfahStock_(sheet.getRange(rowNumber, stockCol + 1).getDisplayValue()) !== stock) continue;
    sheet.getRange(rowNumber, statusCol + 1).setValue(best.status);
    sheet.getRange(rowNumber, noteCol + 1).setValue(nextNote);
    updated++;
    changes.push({ row: rowNumber, stock: stock, oe: String(row[oeCol] || ""), oldStatus: curStatus, status: best.status, oldNote: curNote, note: nextNote });
  }
  var mailed = false;
  if (changes.length) {
    try {
      oemSendReport_(changes, matched, kept, sheet);
      mailed = true;
    } catch (errMail) {}
  }
  return { ok: true, matched: matched, updated: updated, kept: kept, mailed: mailed };
}

function oemEsc_(value) {
  return String(value == null ? "" : value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function oemSendReport_(changes, matched, kept, sheet) {
  var tz = Session.getScriptTimeZone() || "Europe/Berlin";
  var stamp = Utilities.formatDate(new Date(), tz, "dd.MM.yyyy HH:mm");
  var link = "https://docs.google.com/spreadsheets/d/" + OEM_SHEET_ID + "/edit#gid=" + sheet.getSheetId();
  var cell = "border:1px solid #d0d7de;padding:4px 8px;font:12px Arial,sans-serif;vertical-align:top;";
  var head = "background:#f6f8fa;font-weight:bold;" + cell;
  var rows = changes.map(function (c) {
    return "<tr>"
      + "<td style=\"" + cell + "\"><a href=\"" + link + "&range=A" + c.row + "\">" + c.row + "</a></td>"
      + "<td style=\"" + cell + "\">" + oemEsc_(c.stock) + "</td>"
      + "<td style=\"" + cell + "\">" + oemEsc_(c.oe) + "</td>"
      + "<td style=\"" + cell + "\">" + oemEsc_(c.oldStatus || "leer") + "</td>"
      + "<td style=\"" + cell + "font-weight:bold;\">" + oemEsc_(c.status) + "</td>"
      + "<td style=\"" + cell + "\">" + oemEsc_(c.oldNote || "leer") + "</td>"
      + "<td style=\"" + cell + "\">" + oemEsc_(c.note) + "</td>"
      + "</tr>";
  }).join("");
  var html = "<div style=\"font:13px Arial,sans-serif;\">"
    + "<p><b>OEM_Teile Abgleich " + oemEsc_(stamp) + "</b></p>"
    + "<p>Gesetzt: " + changes.length + " · Treffer gesamt: " + matched + " · Von Hand gesetzt, nicht angefasst: " + kept + "</p>"
    + "<table style=\"border-collapse:collapse;\"><tr>"
    + "<th style=\"" + head + "\">Zeile</th><th style=\"" + head + "\">Stock-ID</th><th style=\"" + head + "\">OE-Nr</th>"
    + "<th style=\"" + head + "\">Status vorher</th><th style=\"" + head + "\">Status neu</th>"
    + "<th style=\"" + head + "\">Notiz vorher</th><th style=\"" + head + "\">Notiz neu</th></tr>"
    + rows + "</table>"
    + "<p><a href=\"" + link + "\">OEM_Teile öffnen</a></p></div>";
  var text = "OEM_Teile Abgleich " + stamp + "\nGesetzt: " + changes.length + "\n\n" + changes.map(function (c) {
    return "Zeile " + c.row + " | " + c.stock + " | " + c.oe + " | " + (c.oldStatus || "leer") + " -> " + c.status + " | " + c.note;
  }).join("\n");
  MailApp.sendEmail({
    to: OEM_REPORT_MAIL,
    subject: "OEM_Teile Abgleich " + stamp + " · " + changes.length + " geändert",
    body: text,
    htmlBody: html,
    name: "Alfah Retouren"
  });
}

function alfahSyncOemTeileMenu() {
  var res;
  try { res = alfahSyncOemTeile(); } catch (err) { res = { ok: false, error: String(err && err.message ? err.message : err) }; }
  var msg = res.ok
    ? "Treffer: " + res.matched + "\nGesetzt: " + res.updated + "\nVon Hand gesetzt, nicht angefasst: " + res.kept + (res.updated ? "\nReport-Mail: " + (res.mailed ? "raus an " + OEM_REPORT_MAIL : "fehlgeschlagen") : "")
    : "Fehler: " + res.error;
  try { SpreadsheetApp.getUi().alert("OEM_Teile Abgleich", msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (err2) {}
  return res;
}
