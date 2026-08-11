var REIFEN_SHEET_ID = '1dlmZuWfJ3xiJ-LCtNYioVdTTtDN1nrR78p63mDwW18M';
var REIFEN_TAB = 'Reifen Kontrolle';
var CACHE_TAB = 'Reifen Check Cache';
var REFURB_SHEET_ID = '13Oh7gDT8NAul2s0cwQUeaGwMcS3B2MYu0QOdFNMhXzM';
var REFURB_SHEET_NAME = 'Refurbisment List';
var NACHBESTELL_SHEET_ID = '1VGCAHUbOPgsInQICA1GnrtKg1EPK1d1zWB-GkLi6iVE';
var NACHBESTELL_TAB = 'Nachbestellung';
var TAGESLISTE_SHEET_ID = '1PuCLw8UmDjB_pBo_jCZ9rmSD3GJQESHzPoBVu_--MRo';
var TAGESLISTE_TAB = 'Tagesliste';
var TAGESLISTE_GID = 1855179002;
var TL_DATUM_COL = 1;
var TL_SCHICHT_COL = 3;
var TL_REIFEN_COL = 4;
var TL_STOCK_COL = 5;
var TL_MAX_ENTRIES = 5;
var GMAIL_ACCOUNT = 'ersatzteile.hemau@autohero.com';
var CACHE_TTL_MS = 10 * 60 * 1000;
var CACHE_CHUNK = 48000;
var HOT_CACHE_PREFIX = 'reifen_v1_';
var HOT_CACHE_CHUNK = 90000;
var HOT_CACHE_TTL_SEC = 300;
var WEB_APP_URL = 'https://script.google.com/a/macros/auto1.com/s/AKfycbwsGB1o_1z0t9nCVXDx0lu3nQv8Ltj81Dgq5BVw8laLHPA4v4oLUpNvj-qx49iMjeVm/exec';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Reifen Kontrolle')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Reifen Kontrolle')
    .addItem('App öffnen', 'openReifenApp')
    .addItem('Cache neu bauen', 'menuRebuildCache')
    .addItem('Cache-Trigger einrichten', 'installCacheTrigger')
    .addToUi();
}

function openReifenApp() {
  try { ensureCacheTrigger_(); } catch (e0) {}
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><body style="margin:0;font:14px sans-serif;padding:16px;background:#111;color:#eee">' +
    '<div>Öffne Reifen Kontrolle…</div>' +
    '<script>window.onload=function(){window.open(' + JSON.stringify(WEB_APP_URL) + ',"_blank");google.script.host.close();};</script>' +
    '</body></html>'
  ).setWidth(280).setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(html, 'Reifen Kontrolle');
}

function menuRebuildCache() {
  var res = rebuildReifenCache();
  SpreadsheetApp.getUi().alert('Cache neu gebaut.\n' + (res.items ? res.items.length : 0) + ' Stock-IDs geprüft.');
}

function installCacheTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'rebuildReifenCache') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('rebuildReifenCache').timeBased().everyMinutes(5).create();
  rebuildReifenCache();
  try { SpreadsheetApp.getUi().alert('Cache-Trigger aktiv (alle 5 Minuten).'); } catch (e) {}
}

function ensureCacheTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'rebuildReifenCache') return;
  }
  ScriptApp.newTrigger('rebuildReifenCache').timeBased().everyMinutes(5).create();
}

function normalizeStockId_(value) {
  return String(value || '').replace(/\s+/g, '').toUpperCase();
}

function looksLikeStockId_(value) {
  return /^[A-Z]{2}\d{4,8}$/.test(normalizeStockId_(value));
}

function nowStamp_() {
  return Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy HH:mm');
}

function formatDateDe_(val) {
  if (!val && val !== 0) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' || val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, 'Europe/Berlin', 'dd.MM.yyyy');
  }
  return String(val).trim();
}

function gmailAuthUserParam_() {
  return 'authuser=' + encodeURIComponent(GMAIL_ACCOUNT);
}

function gmailSearchUrl_(query) {
  return 'https://mail.google.com/mail/?' + gmailAuthUserParam_() + '#search/' + encodeURIComponent(String(query || ''));
}

function gmailThreadUrl_(threadId) {
  return 'https://mail.google.com/mail/?' + gmailAuthUserParam_() + '#inbox/' + String(threadId || '');
}

function carolUrlFor_(stockId, sheetUrl) {
  var u = String(sheetUrl || '').trim();
  if (/^https?:\/\//i.test(u)) return u;
  return 'https://carol.autohero.com/en-GB/refurbishment?rsv=' + encodeURIComponent(stockId);
}

function linkUrlFromRichText_(rich) {
  if (!rich) return '';
  try {
    var direct = rich.getLinkUrl();
    if (direct) return String(direct).trim();
  } catch (e0) {}
  try {
    var runs = rich.getRuns();
    for (var i = 0; i < runs.length; i++) {
      var u = runs[i].getLinkUrl();
      if (u) return String(u).trim();
    }
  } catch (e1) {}
  return '';
}

function carolUrlFromSheetParts_(rich, formula, value) {
  var fromRich = linkUrlFromRichText_(rich);
  if (fromRich) return fromRich;
  var f = String(formula || '');
  var m = f.match(/HYPERLINK\s*\(\s*"([^"]+)"/i) || f.match(/HYPERLINK\s*\(\s*'([^']+)'/i);
  if (m && m[1]) return String(m[1]).trim();
  var v = String(value || '').trim();
  if (/^https?:\/\//i.test(v)) return v;
  return '';
}

function formatReifenLabel_(val) {
  var s = String(val || '').trim();
  if (!s) return '';
  if (/^werkstatt\s*1$/i.test(s) || /^ws\s*1$/i.test(s)) return 'Reifen da';
  return s;
}

function isReifenPart_(text) {
  var s = String(text || '');
  if (/reifen|felge|komplettrad|räder|raeder|tyre|tire/i.test(s)) return true;
  if (/\brad\b|\bräder\b/i.test(s)) return true;
  return false;
}

function normalizeSheetStatus_(status) {
  var s = String(status || '').trim().toUpperCase();
  if (!s) return '';
  if (s === 'OK') return 'OK';
  if (s === 'B2A1') return 'B2A1';
  if (s === 'COMPLETE' || s === 'COMPLETED') return 'COMPLETE';
  if (s === 'MECH') return 'MECH';
  return s;
}

function statusKey_(status) {
  var s = normalizeSheetStatus_(status);
  if (s === 'OK') return 'ok';
  if (s === 'B2A1') return 'b2a1';
  if (s === 'COMPLETE') return 'complete';
  if (s === 'MECH') return 'mech';
  if (s === 'PRÜFEN' || s === 'PRUEFEN') return 'pruefen';
  if (!s) return 'none';
  return 'other';
}

function getReifenSheet_() {
  var ss = SpreadsheetApp.openById(REIFEN_SHEET_ID);
  var sheet = ss.getSheetByName(REIFEN_TAB);
  if (!sheet) throw new Error("Tab '" + REIFEN_TAB + "' nicht gefunden");
  return sheet;
}

function getCacheSheet_() {
  var ss = SpreadsheetApp.openById(REIFEN_SHEET_ID);
  var sh = ss.getSheetByName(CACHE_TAB);
  if (!sh) {
    sh = ss.insertSheet(CACHE_TAB);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}

function dedupeReifenSheet_() {
  var removed = 0;
  try {
    var sheet = getReifenSheet_();
    var last = sheet.getLastRow();
    if (last < 3) return 0;
    var data = sheet.getRange(2, 1, last - 1, 2).getDisplayValues();
    var first = {};
    var toDelete = [];
    for (var i = 0; i < data.length; i++) {
      var sid = normalizeStockId_(data[i][0]);
      if (!sid || !looksLikeStockId_(sid)) continue;
      var status = normalizeSheetStatus_(data[i][1]);
      if (!first[sid]) {
        first[sid] = { row: i + 2, status: status };
        continue;
      }
      if (!first[sid].status && status) {
        first[sid].status = status;
        sheet.getRange(first[sid].row, 2).setValue(status);
        try { sheet.getRange(first[sid].row, 2).setFontWeight('bold').setHorizontalAlignment('center'); } catch (eFmt) {}
      }
      toDelete.push(i + 2);
    }
    for (var d = toDelete.length - 1; d >= 0; d--) {
      sheet.deleteRow(toDelete[d]);
      removed++;
    }
    if (removed) SpreadsheetApp.flush();
  } catch (e) {}
  return removed;
}

function readReifenList_() {
  var sheet = getReifenSheet_();
  var last = sheet.getLastRow();
  var list = [];
  if (last < 2) return list;
  var data = sheet.getRange(2, 1, last - 1, 2).getDisplayValues();
  var seen = {};
  for (var i = 0; i < data.length; i++) {
    var sid = normalizeStockId_(data[i][0]);
    if (!sid || !looksLikeStockId_(sid) || seen[sid]) continue;
    seen[sid] = true;
    list.push({
      stockId: sid,
      sheetRow: i + 2,
      sheetStatus: normalizeSheetStatus_(data[i][1])
    });
  }
  return list;
}

function buildRefurbMap_() {
  var map = {};
  try {
    var ss = SpreadsheetApp.openById(REFURB_SHEET_ID);
    var sheet = ss.getSheetByName(REFURB_SHEET_NAME);
    if (!sheet) return map;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return map;
    var data = sheet.getRange(2, 1, lastRow, 30).getValues();
    var formulas = sheet.getRange(2, 3, lastRow, 3).getFormulas();
    for (var i = 0; i < data.length; i++) {
      var stockId = normalizeStockId_(data[i][1]);
      if (!stockId || map[stockId]) continue;
      map[stockId] = {
        found: true,
        row: i + 2,
        carolUrl: carolUrlFromSheetParts_(null, formulas[i][0], data[i][2]),
        markeModel: String(data[i][12] || '').trim(),
        schaeden: String(data[i][22] || '').trim(),
        kommBestellung: String(data[i][23] || '').trim(),
        kommAnlieferung: String(data[i][24] || '').trim(),
        status: String(data[i][25] || '').trim(),
        regal: String(data[i][27] || '').trim(),
        reifenStatusRaw: String(data[i][29] || '').trim(),
        reifenStatus: formatReifenLabel_(data[i][29])
      };
    }
  } catch (e) {}
  return map;
}

function buildNachbestellMap_() {
  var map = {};
  try {
    var ss = SpreadsheetApp.openById(NACHBESTELL_SHEET_ID);
    var sheet = ss.getSheetByName(NACHBESTELL_TAB);
    if (!sheet) return map;
    var lastRow = Math.max(2, sheet.getLastRow());
    var lastCol = Math.min(30, Math.max(1, sheet.getLastColumn()));
    var scanRows = Math.min(8, lastRow);
    var scan = sheet.getRange(1, 1, scanRows, lastCol).getValues();
    var headerRow = 0;
    var bestScore = 0;
    for (var hr = 0; hr < scan.length; hr++) {
      var score = 0;
      for (var hc = 0; hc < scan[hr].length; hc++) {
        var ht = String(scan[hr][hc] || '').toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
        if (ht.indexOf('stock') !== -1) score += 3;
        if (ht === 'status') score += 2;
        if (ht.indexOf('lagerort') !== -1 || ht === 'regal') score += 2;
        if (ht.indexOf('datum') !== -1 || ht === 'date') score += 1;
        if (ht.indexOf('ersatzteil') !== -1 || ht.indexOf('benennung') !== -1 || ht === 'teil') score += 1;
      }
      if (score > bestScore) {
        bestScore = score;
        headerRow = hr;
      }
    }
    var header = scan[headerRow];
    var stockCol = 2;
    var teilCol = 5;
    var statusCol = 11;
    var regalCol = 13;
    var typCol = 0;
    var dateCol = 0;
    var bestellerCol = 0;
    for (var h = 0; h < header.length; h++) {
      var t = String(header[h] || '').toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
      if (t.indexOf('stock') !== -1) stockCol = h + 1;
      if (t.indexOf('ersatzteil') !== -1 || t === 'teil' || t.indexOf('benennung') !== -1 || t.indexOf('bezeichnung') !== -1) teilCol = h + 1;
      if (t === 'status') statusCol = h + 1;
      if (t.indexOf('lagerort') !== -1 || t === 'regal') regalCol = h + 1;
      if (t === 'art' || t === 'typ' || t.indexOf('artdernachbestellung') !== -1) typCol = h + 1;
      if (t === 'datum' || t === 'date' || t.indexOf('bestelldatum') !== -1 || t.indexOf('erstellt') !== -1) dateCol = h + 1;
      if (t === 'besteller') bestellerCol = h + 1;
    }
    var data = sheet.getRange(headerRow + 1, 1, lastRow, lastCol).getValues();
    for (var i = 1; i < data.length; i++) {
      var sid = normalizeStockId_(data[i][stockCol - 1]);
      if (!sid) continue;
      var rawDate = dateCol ? data[i][dateCol - 1] : '';
      var dateMs = 0;
      var dateStr = '';
      if (rawDate instanceof Date || Object.prototype.toString.call(rawDate) === '[object Date]') {
        if (!isNaN(rawDate.getTime())) {
          dateMs = rawDate.getTime();
          dateStr = formatDateDe_(rawDate);
        }
      } else if (rawDate) {
        dateStr = String(rawDate).trim();
        var parsed = new Date(rawDate);
        if (!isNaN(parsed.getTime())) {
          dateMs = parsed.getTime();
          dateStr = formatDateDe_(parsed);
        }
      }
      var typVal = typCol ? data[i][typCol - 1] : '';
      if (typVal instanceof Date || Object.prototype.toString.call(typVal) === '[object Date]') typVal = '';
      var teil = String(data[i][teilCol - 1] || '').trim();
      if (!map[sid]) map[sid] = [];
      map[sid].push({
        typ: String(typVal || '').trim(),
        teil: teil,
        besteller: bestellerCol ? String(data[i][bestellerCol - 1] || '').trim() : '',
        status: String(data[i][statusCol - 1] || '').trim(),
        regal: String(data[i][regalCol - 1] || '').trim(),
        date: dateStr,
        dateMs: dateMs,
        reifen: isReifenPart_(teil),
        sheetRow: headerRow + 1 + i
      });
    }
    var keys = Object.keys(map);
    for (var k = 0; k < keys.length; k++) {
      map[keys[k]].sort(function(a, b) {
        return (b.dateMs || 0) - (a.dateMs || 0);
      });
    }
  } catch (e1) {}
  return map;
}

function isGreenColor_(hex) {
  var h = String(hex || '').trim().toLowerCase();
  if (!h || h === '#ffffff' || h === '#fff' || h === 'white') return false;
  if (h.charAt(0) !== '#') h = '#' + h;
  if (h.length === 4) {
    h = '#' + h.charAt(1) + h.charAt(1) + h.charAt(2) + h.charAt(2) + h.charAt(3) + h.charAt(3);
  }
  var known = ['#00ff00', '#34a853', '#b7e1cd', '#d9ead3', '#a8d08d', '#93c47d', '#6aa84f', '#38761d', '#274e13', '#00b050', '#92d050'];
  if (known.indexOf(h) !== -1) return true;
  if (!/^#[0-9a-f]{6}$/.test(h)) return false;
  var r = parseInt(h.substr(1, 2), 16);
  var g = parseInt(h.substr(3, 2), 16);
  var b = parseInt(h.substr(5, 2), 16);
  return g >= 110 && g > r + 25 && g > b + 25;
}

function isReifenGestelltText_(reifen) {
  var v = String(reifen || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  if (!v) return false;
  if (v === '2' || v === '4') return true;
  if (v === '2 gestellt' || v === '4 gestellt') return true;
  if (/^(2|4)\s*gestellt$/.test(v)) return true;
  if (/\bgestellt\b/.test(v) && v.indexOf('nicht') === -1) return true;
  return false;
}

function isReifenGestellt_(reifen, bg) {
  if (isReifenGestelltText_(reifen)) return true;
  var v = String(reifen || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  if (!v || v === '2 reifen' || v === '4 reifen' || v === 'lagerbestand') return false;
  return isGreenColor_(bg);
}

function normalizeTlEntryGestellt_(entry) {
  if (!entry) return entry;
  if (isReifenGestelltText_(entry.reifen)) entry.reifenGreen = true;
  return entry;
}

function findTlGestelltEntry_(entries) {
  var list = entries || [];
  for (var i = 0; i < list.length; i++) {
    var e = normalizeTlEntryGestellt_(list[i]);
    if (e && (e.reifenGreen || isReifenGestelltText_(e.reifen))) return e;
  }
  return null;
}

function resolveTlGestellt_(tl) {
  var entries = tl || [];
  var latest = entries.length ? normalizeTlEntryGestellt_(entries[0]) : null;
  var gestelltEntry = findTlGestelltEntry_(entries);
  var fromEarlier = !!(gestelltEntry && latest && gestelltEntry.row !== latest.row);
  return {
    latest: latest,
    gestelltEntry: gestelltEntry,
    gestellt: !!gestelltEntry,
    fromEarlier: fromEarlier,
    reifen: gestelltEntry ? gestelltEntry.reifen : (latest && latest.reifen ? latest.reifen : ''),
    datum: gestelltEntry ? gestelltEntry.datum : (latest && latest.datum ? latest.datum : ''),
    url: gestelltEntry ? gestelltEntry.url : (latest && latest.url ? latest.url : '')
  };
}

function normalizeTlGestellt_(it) {
  if (!it) return it;
  if (isReifenGestelltText_(it.tlReifen)) it.tlReifenGreen = true;
  return it;
}

function applyTlGestelltToItem_(it, detail) {
  if (!it) return it;
  normalizeTlGestellt_(it);
  var entries = detail && detail.tagesliste ? (detail.tagesliste.entries || []) : [];
  var resolved = resolveTlGestellt_(entries.length ? entries : (
    (it.tlReifen || it.tlReifenGreen)
      ? [{ reifen: it.tlReifen || '', reifenGreen: !!it.tlReifenGreen, datum: it.tlDatum || '', url: it.tlUrl || '', row: 0 }]
      : []
  ));
  if (resolved.gestellt) {
    it.tlReifenGreen = true;
    it.tlReifen = resolved.reifen || it.tlReifen || '';
    if (resolved.datum) it.tlDatum = resolved.datum;
    if (resolved.url) it.tlUrl = resolved.url;
    it.tlGestelltEarlier = !!resolved.fromEarlier;
  }
  return it;
}

function tageslisteUrl_(row) {
  var u = 'https://docs.google.com/spreadsheets/d/' + TAGESLISTE_SHEET_ID + '/edit#gid=' + TAGESLISTE_GID;
  if (row) u += '&range=E' + row;
  return u;
}

function getTageslisteSheet_() {
  var ss = SpreadsheetApp.openById(TAGESLISTE_SHEET_ID);
  var sh = ss.getSheetByName(TAGESLISTE_TAB);
  if (sh) return sh;
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getSheetId() === TAGESLISTE_GID) return all[i];
  }
  return null;
}

function buildTageslisteMap_(onlyIds) {
  var map = {};
  try {
    var sheet = getTageslisteSheet_();
    if (!sheet) return map;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return map;
    var numCols = Math.max(TL_STOCK_COL, TL_REIFEN_COL);
    var data = sheet.getRange(1, 1, lastRow, numCols).getDisplayValues();
    for (var i = 0; i < data.length; i++) {
      var sid = normalizeStockId_(data[i][TL_STOCK_COL - 1]);
      if (!sid || !looksLikeStockId_(sid)) continue;
      if (onlyIds && !onlyIds[sid]) continue;
      var reifen = String(data[i][TL_REIFEN_COL - 1] || '').trim();
      var entry = {
        row: i + 1,
        datum: String(data[i][TL_DATUM_COL - 1] || '').trim(),
        schicht: String(data[i][TL_SCHICHT_COL - 1] || '').trim(),
        reifen: reifen,
        reifenGreen: isReifenGestelltText_(reifen),
        url: tageslisteUrl_(i + 1)
      };
      if (!map[sid]) map[sid] = [];
      map[sid].push(entry);
    }
    var keys = Object.keys(map);
    for (var k = 0; k < keys.length; k++) {
      var list = map[keys[k]];
      list.sort(function(a, b) { return b.row - a.row; });
      if (list.length > TL_MAX_ENTRIES) map[keys[k]] = list.slice(0, TL_MAX_ENTRIES);
    }
  } catch (e) {}
  return map;
}

function searchReturnMail_(stockId) {
  var result = { found: false, subject: '', from: '', date: '', permalink: '', message: '' };
  try {
    var queries = [
      '"Return to Auto1" "' + stockId + '"',
      '"' + stockId + '---Return to Auto1"'
    ];
    var seen = {};
    var best = null;
    for (var q = 0; q < queries.length; q++) {
      var threads = GmailApp.search(queries[q], 0, 5);
      for (var i = 0; i < threads.length; i++) {
        var id = threads[i].getId();
        if (seen[id]) continue;
        seen[id] = true;
        if (!best || threads[i].getLastMessageDate().getTime() > best.getLastMessageDate().getTime()) {
          best = threads[i];
        }
      }
      if (best) break;
    }
    if (best) {
      var msgs = best.getMessages();
      var last = msgs[msgs.length - 1];
      result.found = true;
      result.subject = String(best.getFirstMessageSubject() || '');
      result.from = String(last.getFrom() || '');
      result.date = Utilities.formatDate(last.getDate(), 'Europe/Berlin', 'dd.MM.yyyy HH:mm');
      result.permalink = gmailThreadUrl_(best.getId());
    }
  } catch (err) {
    result.message = 'Gmail: ' + String(err.message || err);
  }
  return result;
}

function hasReifenDaComment_(refurb) {
  if (!refurb || !refurb.found) return false;
  var text = String(refurb.kommBestellung || '') + ' ' + String(refurb.kommAnlieferung || '');
  return /reifen\s*(sind\s*|ist\s*)?da/i.test(text);
}

function evaluateStock_(stockId, refurb, nbs, returnMail, sheetStatus, tl) {
  refurb = refurb || { found: false };
  nbs = nbs || [];
  tl = tl || [];
  sheetStatus = normalizeSheetStatus_(sheetStatus);
  var sheetSet = !!sheetStatus && sheetStatus !== 'PRÜFEN' && sheetStatus !== 'PRUEFEN';
  var reifenWerkstatt = refurb.found && refurb.reifenStatus === 'Reifen da';
  var reifenComment = hasReifenDaComment_(refurb);
  var reifenDa = reifenWerkstatt || reifenComment;
  var statusLow = String(refurb.status || '').toLowerCase();
  var carolDone = refurb.found && /herausgegeben|handed\s*out|complete/i.test(statusLow);
  var teilweise = /teilweise/.test(statusLow);
  var komplett = /komplett/.test(statusLow);
  var nbReifen = [];
  var nbReifenOpen = [];
  for (var i = 0; i < nbs.length; i++) {
    if (!nbs[i].reifen) continue;
    nbReifen.push(nbs[i]);
    var st = String(nbs[i].status || '').toLowerCase();
    var closed = st.indexOf('komplett') !== -1 || st.indexOf('fertiggestellt') !== -1 || st === 'angeliefert' || st.indexOf('fahrzeug rr') !== -1 || st.indexOf('nicht bestellt') !== -1 || st.indexOf('nicht notwendig') !== -1;
    if (!closed) nbReifenOpen.push(nbs[i]);
  }
  var suggested = '';
  var reasons = [];
  if (returnMail && returnMail.found) {
    suggested = 'B2A1';
    reasons.push('Return-to-Auto1 Mail gefunden (' + returnMail.date + ')');
  }
  if (refurb.found) {
    if (reifenWerkstatt) reasons.push('Reifen da (Werkstatt 1 in Refurbishment)');
    if (reifenComment) reasons.push('"Reifen da" im Kommentar gefunden');
    if (teilweise) reasons.push('Status: teilweise angeliefert');
    if (komplett) reasons.push('Status: komplett angeliefert');
    if (carolDone) reasons.push('Auftrag wurde herausgegeben' + (sheetSet ? '' : ' — in Carol checken'));
    if (nbReifenOpen.length) reasons.push(nbReifenOpen.length + ' offene Reifen-Nachbestellung(en)');
  }
  if (tl.length) {
    var tlRes = resolveTlGestellt_(tl);
    var tlLatest = tlRes.latest || {};
    var tlTxt = 'War auf Tagesliste' + (tlLatest.datum ? ' am ' + tlLatest.datum : '') + (tlLatest.schicht ? ' (' + tlLatest.schicht + ')' : '');
    if (tlRes.gestellt) {
      var gWhen = (tlRes.gestelltEntry.datum ? ' am ' + tlRes.gestelltEntry.datum : '') +
        (tlRes.gestelltEntry.schicht ? ' (' + tlRes.gestelltEntry.schicht + ')' : '');
      tlTxt += ' — Reifen gestellt · bei Mechanik (nicht mehr im Lager) · Wert: ' + (tlRes.reifen || '');
      if (tlRes.fromEarlier) tlTxt += ' · gestellt in früherem Eintrag' + gWhen;
      if (tl.length > 1) tlTxt += ' · ' + tl.length + '× auf Tagesliste';
    } else if (tlLatest.reifen) {
      tlTxt += ' — Reifen nicht gestellt · noch im Lager · Wert: ' + tlLatest.reifen;
    } else {
      tlTxt += ' — nicht gestellt — kein Reifen-Eintrag (2/4)';
    }
    reasons.push(tlTxt);
  }
  if (!suggested) {
    if (!refurb.found) {
      suggested = 'PRÜFEN';
      if (!sheetSet) reasons.push('Nicht in Refurbishment List — evtl. schon Completed, in Carol checken');
    } else if (reifenDa && !nbReifenOpen.length && !carolDone) {
      suggested = 'OK';
    } else {
      suggested = 'PRÜFEN';
      if (!reifenDa && !sheetSet) reasons.push('Kein "Reifen da" Nachweis in Refurbishment');
    }
  }
  if (sheetSet) {
    if (!(suggested === 'B2A1' && sheetStatus !== 'B2A1')) suggested = '';
    reasons.unshift('Als ' + sheetStatus + ' bestätigt');
  }
  var teile = [];
  for (var t = 0; t < nbReifen.length; t++) {
    if (nbReifen[t].teil) teile.push(nbReifen[t].teil);
  }
  return {
    reifenDa: reifenDa,
    reifenWerkstatt: reifenWerkstatt,
    reifenComment: reifenComment,
    carolDone: carolDone,
    teilweise: teilweise,
    komplett: komplett,
    inRefurb: refurb.found && !carolDone,
    nbReifenCount: nbReifen.length,
    nbReifenOpenCount: nbReifenOpen.length,
    nbReifenTeile: teile,
    suggested: suggested,
    reasons: reasons
  };
}

function buildDetail_(entry, refurbMap, nbMap, tlMap, withMail) {
  var stockId = entry.stockId;
  var refurb = refurbMap[stockId] || { found: false };
  var nbs = nbMap[stockId] || [];
  var tl = (tlMap && tlMap[stockId]) || [];
  var returnMail = withMail ? searchReturnMail_(stockId) : { found: false, subject: '', from: '', date: '', permalink: '', message: '' };
  var evalRes = evaluateStock_(stockId, refurb, nbs, returnMail, entry.sheetStatus, tl);
  var tlRes = resolveTlGestellt_(tl);
  return {
    success: true,
    stockId: stockId,
    sheetRow: entry.sheetRow || 0,
    sheetStatus: entry.sheetStatus || '',
    refurb: refurb,
    nachbestellungen: nbs,
    tagesliste: {
      found: tl.length > 0,
      count: tl.length,
      entries: tl,
      latest: tlRes.latest,
      gestelltEntry: tlRes.gestelltEntry,
      anyGestellt: tlRes.gestellt,
      gestelltEarlier: tlRes.fromEarlier,
      url: tlRes.url || (tl.length ? tl[0].url : '')
    },
    returnMail: returnMail,
    evalx: evalRes,
    carolUrl: carolUrlFor_(stockId, refurb.carolUrl),
    gmailSearchUrl: gmailSearchUrl_(stockId),
    gmailReturnSearchUrl: gmailSearchUrl_('"Return to Auto1" "' + stockId + '"'),
    checkedAt: nowStamp_()
  };
}

function itemFromDetail_(d) {
  var tlRes = resolveTlGestellt_((d.tagesliste && d.tagesliste.entries) || []);
  return {
    stockId: d.stockId,
    sheetRow: d.sheetRow,
    sheetStatus: d.sheetStatus,
    statusKey: statusKey_(d.sheetStatus),
    suggested: d.evalx.suggested,
    suggestedKey: statusKey_(d.evalx.suggested),
    markeModel: (d.refurb && d.refurb.markeModel) || '',
    refurbFound: !!(d.refurb && d.refurb.found),
    refurbStatus: (d.refurb && d.refurb.status) || '',
    reifenDa: d.evalx.reifenDa,
    carolDone: d.evalx.carolDone,
    mailFound: !!(d.returnMail && d.returnMail.found),
    nbCount: (d.nachbestellungen || []).length,
    nbReifenCount: d.evalx.nbReifenCount,
    nbReifenOpenCount: d.evalx.nbReifenOpenCount,
    nbReifenTeile: d.evalx.nbReifenTeile || [],
    tlFound: !!(d.tagesliste && d.tagesliste.found),
    tlReifen: tlRes.reifen || '',
    tlReifenGreen: !!(d.tagesliste && d.tagesliste.anyGestellt) || tlRes.gestellt,
    tlGestelltEarlier: !!(d.tagesliste && d.tagesliste.gestelltEarlier) || tlRes.fromEarlier,
    tlDatum: tlRes.datum || '',
    tlUrl: (d.tagesliste && d.tagesliste.url) || tlRes.url || '',
    mismatch: !!(d.sheetStatus && d.evalx.suggested === 'B2A1' && d.sheetStatus !== 'B2A1')
  };
}

function reuseReturnMail_(detail, oldMail) {
  if (!detail || !oldMail || !oldMail.found) return detail;
  detail.returnMail = oldMail;
  detail.evalx = evaluateStock_(
    detail.stockId,
    detail.refurb,
    detail.nachbestellungen || [],
    oldMail,
    detail.sheetStatus,
    (detail.tagesliste && detail.tagesliste.entries) || []
  );
  return detail;
}

function collectOldMails_(details) {
  var out = {};
  if (!details) return out;
  var keys = Object.keys(details);
  for (var i = 0; i < keys.length; i++) {
    var mail = details[keys[i]] && details[keys[i]].returnMail;
    if (mail && mail.found) out[keys[i]] = mail;
  }
  return out;
}

function stockIdSetFromList_(list) {
  var need = {};
  for (var i = 0; i < list.length; i++) {
    if (list[i] && list[i].stockId) need[list[i].stockId] = true;
  }
  return need;
}

function rebuildReifenCache() {
  var builtAtMs = Date.now();
  var builtAt = Utilities.formatDate(new Date(builtAtMs), 'Europe/Berlin', 'dd.MM.yyyy HH:mm:ss');
  dedupeReifenSheet_();
  var list = readReifenList_();
  var oldMails = collectOldMails_((readCache_() || {}).details);
  var refurbMap = buildRefurbMap_();
  var nbMap = buildNachbestellMap_();
  var tlMap = buildTageslisteMap_(stockIdSetFromList_(list));
  var items = [];
  var details = {};
  for (var i = 0; i < list.length; i++) {
    var d = buildDetail_(list[i], refurbMap, nbMap, tlMap, false);
    reuseReturnMail_(d, oldMails[d.stockId]);
    d.fromCache = true;
    details[d.stockId] = d;
    items.push(itemFromDetail_(d));
  }
  var payload = {
    builtAt: builtAt,
    builtAtMs: builtAtMs,
    items: items,
    details: details,
    count: items.length
  };
  writeCachePayload_(payload);
  return payload;
}

var REBUILD_MAP_KEYS = {
  refurb: 'reifen_rb_refurb_',
  nb: 'reifen_rb_nb_',
  tl: 'reifen_rb_tl_',
  meta: 'reifen_rb_meta_'
};

function clearCacheJson_(prefix) {
  try {
    var cache = CacheService.getScriptCache();
    var n = Number(cache.get(prefix + 'n') || 0);
    var keys = [prefix + 'n'];
    for (var i = 0; i < n; i++) keys.push(prefix + i);
    if (keys.length) cache.removeAll(keys);
  } catch (e0) {}
}

function putCacheJson_(prefix, obj) {
  var cache = CacheService.getScriptCache();
  var raw = JSON.stringify(obj == null ? {} : obj);
  var chunk = 90000;
  var parts = Math.max(1, Math.ceil(raw.length / chunk));
  if (parts > 80) return false;
  clearCacheJson_(prefix);
  cache.put(prefix + 'n', String(parts), 600);
  for (var p = 0; p < parts; p++) {
    cache.put(prefix + p, raw.substr(p * chunk, chunk), 600);
  }
  return true;
}

function getCacheJson_(prefix) {
  try {
    var cache = CacheService.getScriptCache();
    var n = Number(cache.get(prefix + 'n') || 0);
    if (n < 1 || n > 80) return null;
    var json = '';
    for (var i = 0; i < n; i++) json += String(cache.get(prefix + i) || '');
    if (!json) return null;
    return JSON.parse(json);
  } catch (e1) {
    return null;
  }
}

function clearAllRebuildMaps_() {
  clearCacheJson_(REBUILD_MAP_KEYS.refurb);
  clearCacheJson_(REBUILD_MAP_KEYS.nb);
  clearCacheJson_(REBUILD_MAP_KEYS.tl);
  try { CacheService.getScriptCache().remove(REBUILD_MAP_KEYS.meta + 'total'); } catch (e2) {}
}

function loadRebuildMaps_() {
  var refurb = getCacheJson_(REBUILD_MAP_KEYS.refurb);
  var nb = getCacheJson_(REBUILD_MAP_KEYS.nb);
  var tl = getCacheJson_(REBUILD_MAP_KEYS.tl);
  if (!refurb || !nb || !tl) return null;
  return { refurb: refurb, nb: nb, tl: tl };
}

function rebuildStepResult_(opt) {
  return {
    success: true,
    done: !!opt.done,
    nextStep: opt.nextStep || '',
    offset: opt.offset || 0,
    total: opt.total || 0,
    percent: opt.percent || 0,
    count: opt.count != null ? opt.count : (opt.total || 0),
    currentSid: opt.currentSid || '',
    phase: opt.phase || '',
    cachedAt: opt.cachedAt || '',
    message: opt.message || ''
  };
}

function rebuildCacheStep(step, offset, batchSize) {
  try {
    step = String(step || 'start').toLowerCase();
    offset = Math.max(0, Number(offset) || 0);
    batchSize = Math.max(1, Math.min(12, Number(batchSize) || 8));

    if (step === 'start') {
      dedupeReifenSheet_();
      clearAllRebuildMaps_();
      var list0 = readReifenList_();
      try { CacheService.getScriptCache().put(REBUILD_MAP_KEYS.meta + 'total', String(list0.length), 600); } catch (e3) {}
      return rebuildStepResult_({
        done: false,
        nextStep: 'refurb',
        offset: 0,
        total: list0.length,
        percent: 4,
        currentSid: list0.length ? (list0.length + ' Stock-IDs') : 'Liste leer',
        phase: 'Liste vorbereitet',
        count: list0.length
      });
    }

    if (step === 'refurb') {
      putCacheJson_(REBUILD_MAP_KEYS.refurb, buildRefurbMap_());
      var t1 = Number(CacheService.getScriptCache().get(REBUILD_MAP_KEYS.meta + 'total') || readReifenList_().length) || 0;
      return rebuildStepResult_({
        done: false,
        nextStep: 'nb',
        offset: 0,
        total: t1,
        percent: 18,
        currentSid: 'Refurbishment ✓',
        phase: 'Refurbishment geladen',
        count: t1
      });
    }

    if (step === 'nb') {
      putCacheJson_(REBUILD_MAP_KEYS.nb, buildNachbestellMap_());
      var t2 = Number(CacheService.getScriptCache().get(REBUILD_MAP_KEYS.meta + 'total') || readReifenList_().length) || 0;
      return rebuildStepResult_({
        done: false,
        nextStep: 'tl',
        offset: 0,
        total: t2,
        percent: 32,
        currentSid: 'Nachbestellungen ✓',
        phase: 'Nachbestellungen geladen',
        count: t2
      });
    }

    if (step === 'tl') {
      var listTl = readReifenList_();
      putCacheJson_(REBUILD_MAP_KEYS.tl, buildTageslisteMap_(stockIdSetFromList_(listTl)));
      return rebuildStepResult_({
        done: false,
        nextStep: 'items',
        offset: 0,
        total: listTl.length,
        percent: 48,
        currentSid: 'Tagesliste ✓ — starte Checks',
        phase: 'Tagesliste geladen',
        count: listTl.length
      });
    }

    var maps = loadRebuildMaps_();
    if (!maps) {
      return { success: false, message: 'Vorbereitung unvollständig — bitte erneut laden', done: false, nextStep: 'start' };
    }

    var list = readReifenList_();
    var total = list.length;
    var builtAtMs = Date.now();
    var builtAt = Utilities.formatDate(new Date(builtAtMs), 'Europe/Berlin', 'dd.MM.yyyy HH:mm:ss');
    if (!total) {
      clearAllRebuildMaps_();
      writeCachePayload_({ builtAt: builtAt, builtAtMs: builtAtMs, items: [], details: {}, count: 0 });
      return rebuildStepResult_({
        done: true,
        nextStep: '',
        offset: 0,
        total: 0,
        percent: 100,
        currentSid: 'Fertig',
        phase: 'Fertig',
        count: 0,
        cachedAt: builtAt
      });
    }

    var prev = readCache_() || { items: [], details: {} };
    var oldMails = collectOldMails_(prev.details);
    var details = offset === 0 ? {} : (prev.details || {});
    if (offset === 0) details = {};

    var end = Math.min(offset + batchSize, total);
    var currentSid = '';
    for (var i = offset; i < end; i++) {
      currentSid = list[i].stockId;
      var d = buildDetail_(list[i], maps.refurb, maps.nb, maps.tl, false);
      reuseReturnMail_(d, oldMails[d.stockId]);
      d.fromCache = true;
      details[d.stockId] = d;
    }

    if (end < total && prev.details) {
      for (var k = end; k < total; k++) {
        var sidKeep = list[k].stockId;
        if (!details[sidKeep] && prev.details[sidKeep]) details[sidKeep] = prev.details[sidKeep];
      }
    }

    var items = [];
    for (var j = 0; j < total; j++) {
      var sid = list[j].stockId;
      if (details[sid]) {
        items.push(itemFromDetail_(details[sid]));
      } else {
        items.push({
          stockId: sid,
          sheetRow: list[j].sheetRow,
          sheetStatus: list[j].sheetStatus,
          statusKey: statusKey_(list[j].sheetStatus),
          suggested: '',
          suggestedKey: 'none',
          markeModel: '',
          refurbFound: false,
          refurbStatus: '',
          reifenDa: false,
          carolDone: false,
          mailFound: false,
          nbCount: 0,
          nbReifenCount: 0,
          nbReifenOpenCount: 0,
          nbReifenTeile: [],
          tlFound: false,
          tlReifen: '',
          tlReifenGreen: false,
          tlDatum: '',
          tlUrl: '',
          mismatch: false,
          uncached: true
        });
      }
    }

    var done = end >= total;
    var payload = {
      builtAt: builtAt,
      builtAtMs: builtAtMs,
      items: items,
      details: details,
      count: total
    };
    if (done) {
      clearAllRebuildMaps_();
      writeCachePayload_(payload);
    } else {
      writeHotCache_(JSON.stringify(payload));
    }

    var itemPct = total ? Math.round((end / total) * 50) : 50;
    return rebuildStepResult_({
      done: done,
      nextStep: done ? '' : 'items',
      offset: end,
      total: total,
      percent: Math.min(100, 50 + itemPct),
      currentSid: currentSid || '—',
      phase: done ? 'Fertig' : 'Stock-IDs prüfen',
      count: total,
      cachedAt: builtAt,
      message: done ? 'Cache neu gebaut' : ('Fortschritt ' + end + '/' + total)
    });
  } catch (err) {
    return { success: false, message: String(err.message || err), done: false, nextStep: step || 'start' };
  }
}

function rebuildCacheChunk(offset, batchSize) {
  if (!offset) {
    var start = rebuildCacheStep('start', 0, batchSize);
    if (!start.success) return start;
    var r = rebuildCacheStep('refurb', 0, batchSize);
    if (!r.success) return r;
    var n = rebuildCacheStep('nb', 0, batchSize);
    if (!n.success) return n;
    var t = rebuildCacheStep('tl', 0, batchSize);
    if (!t.success) return t;
    return rebuildCacheStep('items', 0, batchSize || 8);
  }
  return rebuildCacheStep('items', offset, batchSize || 8);
}

function isCacheStale_(cache) {
  if (!cache || !cache.builtAtMs) return true;
  return (Date.now() - Number(cache.builtAtMs)) > CACHE_TTL_MS;
}

function writeHotCache_(json) {
  try {
    var cache = CacheService.getScriptCache();
    var parts = Math.max(1, Math.ceil(json.length / HOT_CACHE_CHUNK));
    if (parts > 90) return;
    cache.put(HOT_CACHE_PREFIX + 'n', String(parts), HOT_CACHE_TTL_SEC);
    var batch = {};
    var batchCount = 0;
    for (var i = 0; i < parts; i++) {
      batch[HOT_CACHE_PREFIX + i] = json.substr(i * HOT_CACHE_CHUNK, HOT_CACHE_CHUNK);
      batchCount++;
      if (batchCount >= 20) {
        cache.putAll(batch, HOT_CACHE_TTL_SEC);
        batch = {};
        batchCount = 0;
      }
    }
    if (batchCount) cache.putAll(batch, HOT_CACHE_TTL_SEC);
  } catch (e) {}
}

function readHotCache_() {
  try {
    var cache = CacheService.getScriptCache();
    var nStr = cache.get(HOT_CACHE_PREFIX + 'n');
    var parts = parseInt(nStr, 10) || 0;
    if (parts < 1) return null;
    var keys = [];
    for (var i = 0; i < parts; i++) keys.push(HOT_CACHE_PREFIX + i);
    var got = cache.getAll(keys);
    var json = '';
    for (var j = 0; j < parts; j++) {
      var chunk = got[HOT_CACHE_PREFIX + j];
      if (chunk == null) return null;
      json += chunk;
    }
    if (!json) return null;
    return JSON.parse(json);
  } catch (e) {
    return null;
  }
}

function writeCachePayload_(obj) {
  var json = JSON.stringify(obj);
  var sh = getCacheSheet_();
  sh.clear();
  var parts = Math.max(1, Math.ceil(json.length / CACHE_CHUNK));
  sh.getRange(1, 1, 1, 2).setValues([[obj.builtAt || '', parts]]);
  var rows = [];
  for (var i = 0; i < parts; i++) {
    rows.push([json.substr(i * CACHE_CHUNK, CACHE_CHUNK)]);
  }
  sh.getRange(2, 1, parts, 1).setValues(rows);
  writeHotCache_(json);
}

function readCache_() {
  var hot = readHotCache_();
  if (hot) return hot;
  try {
    var sh = SpreadsheetApp.openById(REIFEN_SHEET_ID).getSheetByName(CACHE_TAB);
    if (!sh || sh.getLastRow() < 2) return null;
    var meta = sh.getRange(1, 1, 1, 2).getValues()[0];
    var parts = parseInt(meta[1], 10) || 0;
    if (parts < 1) return null;
    var chunks = sh.getRange(2, 1, parts, 1).getValues();
    var json = '';
    for (var i = 0; i < chunks.length; i++) json += String(chunks[i][0] || '');
    if (!json) return null;
    writeHotCache_(json);
    return JSON.parse(json);
  } catch (err) {
    return null;
  }
}

function filterItems_(items, filter) {
  filter = String(filter || 'alle').toLowerCase();
  var out = [];
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    if (filter === 'ohne' && it.sheetStatus) continue;
    if (filter === 'ok' && it.statusKey !== 'ok') continue;
    if (filter === 'b2a1' && it.statusKey !== 'b2a1' && !it.mailFound) continue;
    if (filter === 'complete' && it.statusKey !== 'complete' && !it.carolDone) continue;
    if (filter === 'mech' && it.statusKey !== 'mech') continue;
    if (filter === 'gestellt' && !(it.tlReifenGreen || isReifenGestelltText_(it.tlReifen))) continue;
    if ((filter === 'imlager' || filter === 'ohnegestellt') && (it.tlReifenGreen || isReifenGestelltText_(it.tlReifen))) continue;
    if (filter === 'mail' && !it.mailFound) continue;
    if (filter === 'pruefen' && it.suggested !== 'PRÜFEN' && !it.mismatch) continue;
    if (filter === 'reifenfehlt' && it.reifenDa) continue;
    out.push(it);
  }
  return out;
}

function getQueue(filter, includeDetails) {
  try {
    var cache = readCache_();
    if (!cache || !cache.items) cache = rebuildReifenCache();
    var live = readReifenList_();
    var liveMap = {};
    for (var l = 0; l < live.length; l++) liveMap[live[l].stockId] = live[l];
    var items = [];
    var seen = {};
    for (var i = 0; i < cache.items.length; i++) {
      var it = JSON.parse(JSON.stringify(cache.items[i]));
      var lv = liveMap[it.stockId];
      if (!lv) continue;
      seen[it.stockId] = true;
      it.sheetRow = lv.sheetRow;
      it.sheetStatus = lv.sheetStatus;
      it.statusKey = statusKey_(lv.sheetStatus);
      if (it.sheetStatus && !(it.suggested === 'B2A1' && it.sheetStatus !== 'B2A1')) {
        it.suggested = '';
        it.suggestedKey = 'none';
      }
      it.mismatch = !!(it.sheetStatus && it.suggested === 'B2A1' && it.sheetStatus !== 'B2A1');
      applyTlGestelltToItem_(it, cache.details && cache.details[it.stockId]);
      items.push(it);
    }
    for (var n = 0; n < live.length; n++) {
      if (seen[live[n].stockId]) continue;
      items.push({
        stockId: live[n].stockId,
        sheetRow: live[n].sheetRow,
        sheetStatus: live[n].sheetStatus,
        statusKey: statusKey_(live[n].sheetStatus),
        suggested: '',
        suggestedKey: 'none',
        markeModel: '',
        refurbFound: false,
        refurbStatus: '',
        reifenDa: false,
        carolDone: false,
        mailFound: false,
        nbCount: 0,
        nbReifenCount: 0,
        nbReifenOpenCount: 0,
        nbReifenTeile: [],
        tlFound: false,
        tlReifen: '',
        tlReifenGreen: false,
        tlDatum: '',
        tlUrl: '',
        mismatch: false,
        uncached: true
      });
    }
    items.sort(function(a, b) { return (a.sheetRow || 0) - (b.sheetRow || 0); });
    var out = {
      success: true,
      items: filterItems_(items, filter),
      total: items.length,
      cachedAt: cache.builtAt || '',
      cachedAtMs: Number(cache.builtAtMs) || 0
    };
    if (includeDetails && cache.details) {
      var details = {};
      var dkeys = Object.keys(cache.details);
      for (var di = 0; di < dkeys.length; di++) {
        var det = cache.details[dkeys[di]];
        var dlv = liveMap[dkeys[di]];
        if (!dlv) continue;
        det.sheetStatus = dlv.sheetStatus;
        det.sheetRow = dlv.sheetRow;
        det.inSheet = true;
        if (det.evalx && dlv.sheetStatus && !(det.evalx.suggested === 'B2A1' && dlv.sheetStatus !== 'B2A1')) {
          det.evalx.suggested = '';
        }
        det.cachedAt = cache.builtAt || '';
        normalizeTageslisteDetail_(det);
        details[dkeys[di]] = det;
      }
      out.details = details;
    }
    return out;
  } catch (err) {
    return { success: false, message: String(err.message || err), items: [] };
  }
}

function normalizeTageslisteDetail_(det) {
  if (!det || !det.tagesliste) return det;
  var ents = det.tagesliste.entries || [];
  for (var ei = 0; ei < ents.length; ei++) normalizeTlEntryGestellt_(ents[ei]);
  var tlFix = resolveTlGestellt_(ents);
  det.tagesliste.latest = tlFix.latest || det.tagesliste.latest;
  det.tagesliste.gestelltEntry = tlFix.gestelltEntry;
  det.tagesliste.anyGestellt = tlFix.gestellt;
  det.tagesliste.gestelltEarlier = tlFix.fromEarlier;
  if (tlFix.url) det.tagesliste.url = tlFix.url;
  return det;
}

function getStockDetail(stockId) {
  try {
    stockId = normalizeStockId_(stockId);
    if (!stockId) return { success: false, message: 'Keine Stock-ID' };
    var cache = readCache_();
    if (cache && cache.details && cache.details[stockId]) {
      var d = cache.details[stockId];
      d.cachedAt = cache.builtAt || '';
      return normalizeTageslisteDetail_(d);
    }
    return checkStockLive(stockId);
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  }
}

function checkStockLive(stockId) {
  try {
    stockId = normalizeStockId_(stockId);
    if (!stockId) return { success: false, message: 'Keine Stock-ID' };
    if (!looksLikeStockId_(stockId)) return { success: false, message: 'Ungültige Stock-ID: ' + stockId };
    var live = readReifenList_();
    var entry = null;
    for (var i = 0; i < live.length; i++) {
      if (live[i].stockId === stockId) { entry = live[i]; break; }
    }
    if (!entry) entry = { stockId: stockId, sheetRow: 0, sheetStatus: '' };
    var refurbMap = buildRefurbMap_();
    var nbMap = buildNachbestellMap_();
    var tlMap = buildTageslisteMap_();
    var d = buildDetail_(entry, refurbMap, nbMap, tlMap, true);
    d.fromCache = false;
    d.inSheet = entry.sheetRow > 0;
    try {
      var cache = readCache_();
      if (cache && cache.details) {
        cache.details[stockId] = d;
        var replaced = false;
        for (var j = 0; j < cache.items.length; j++) {
          if (cache.items[j].stockId === stockId) {
            cache.items[j] = itemFromDetail_(d);
            replaced = true;
            break;
          }
        }
        if (!replaced && entry.sheetRow > 0) cache.items.push(itemFromDetail_(d));
        writeCachePayload_(cache);
      }
    } catch (e2) {}
    return d;
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  }
}

function setSheetStatus(stockId, status) {
  try {
    stockId = normalizeStockId_(stockId);
    if (!stockId) return { success: false, message: 'Keine Stock-ID' };
    status = normalizeSheetStatus_(status);
    var sheet = getReifenSheet_();
    var last = sheet.getLastRow();
    var row = 0;
    if (last >= 2) {
      var data = sheet.getRange(2, 1, last - 1, 1).getDisplayValues();
      for (var i = 0; i < data.length; i++) {
        if (normalizeStockId_(data[i][0]) === stockId) {
          row = i + 2;
          break;
        }
      }
    }
    if (!row) {
      sheet.appendRow([stockId, status]);
      row = sheet.getLastRow();
    } else {
      sheet.getRange(row, 2).setValue(status);
    }
    try { sheet.getRange(row, 2).setFontWeight('bold').setHorizontalAlignment('center'); } catch (eFmt) {}
    SpreadsheetApp.flush();
    try {
      var cache = readCache_();
      if (cache && cache.details && cache.details[stockId]) {
        var det = cache.details[stockId];
        det.sheetStatus = status;
        det.sheetRow = row;
        if (det.evalx && status && !(det.evalx.suggested === 'B2A1' && status !== 'B2A1')) {
          det.evalx.suggested = '';
        }
        for (var j = 0; j < cache.items.length; j++) {
          if (cache.items[j].stockId !== stockId) continue;
          cache.items[j].sheetStatus = status;
          cache.items[j].statusKey = statusKey_(status);
          cache.items[j].sheetRow = row;
          if (status && !(cache.items[j].suggested === 'B2A1' && status !== 'B2A1')) {
            cache.items[j].suggested = '';
            cache.items[j].suggestedKey = 'none';
          }
          cache.items[j].mismatch = !!(status && cache.items[j].suggested === 'B2A1' && status !== 'B2A1');
          break;
        }
        writeCachePayload_(cache);
      }
    } catch (e2) {}
    return { success: true, message: (status ? status : 'Status entfernt') + ' — ' + stockId + ' gespeichert', status: status, sheetRow: row };
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  }
}

function deleteStockId(stockId) {
  var res = deleteStockIds([stockId]);
  if (!res || !res.success) return res;
  if (!res.deleted) return { success: false, message: stockId + ' nicht im Sheet gefunden' };
  return { success: true, message: stockId + ' aus Sheet gelöscht', deleted: res.deleted };
}

function deleteStockIds(ids) {
  try {
    var raw = [];
    if (Object.prototype.toString.call(ids) === '[object Array]') raw = ids;
    else raw = String(ids || '').split(/[\s,;]+/);
    var want = {};
    var order = [];
    for (var r = 0; r < raw.length; r++) {
      var sid = normalizeStockId_(raw[r]);
      if (!sid || want[sid]) continue;
      want[sid] = true;
      order.push(sid);
    }
    if (!order.length) return { success: false, message: 'Keine Stock-IDs', deleted: 0, ids: [] };
    var sheet = getReifenSheet_();
    var last = sheet.getLastRow();
    var deletedRows = 0;
    var removed = {};
    if (last >= 2) {
      var data = sheet.getRange(2, 1, last - 1, 1).getDisplayValues();
      for (var i = data.length - 1; i >= 0; i--) {
        var cur = normalizeStockId_(data[i][0]);
        if (cur && want[cur]) {
          sheet.deleteRow(i + 2);
          deletedRows++;
          removed[cur] = true;
        }
      }
    }
    if (deletedRows) SpreadsheetApp.flush();
    try {
      var cache = readCache_();
      if (cache) {
        var changed = false;
        for (var k = 0; k < order.length; k++) {
          var id = order[k];
          if (cache.details && cache.details[id]) {
            delete cache.details[id];
            changed = true;
          }
        }
        if (cache.items) {
          for (var j = cache.items.length - 1; j >= 0; j--) {
            if (want[cache.items[j].stockId]) {
              cache.items.splice(j, 1);
              changed = true;
            }
          }
        }
        if (changed) writeCachePayload_(cache);
      }
    } catch (e2) {}
    var gone = order.filter(function(id) { return !!removed[id]; });
    return {
      success: true,
      message: gone.length + ' Stock-ID' + (gone.length === 1 ? '' : 's') + ' gelöscht',
      deleted: gone.length,
      ids: gone,
      missing: order.filter(function(id) { return !removed[id]; })
    };
  } catch (err) {
    return { success: false, message: String(err.message || err), deleted: 0, ids: [] };
  }
}

function addStockIds(ids) {
  try {
    var raw = [];
    if (Object.prototype.toString.call(ids) === '[object Array]') raw = ids;
    else raw = String(ids || '').split(/[\s,;]+/);
    var sheet = getReifenSheet_();
    var last = sheet.getLastRow();
    var existing = {};
    if (last >= 2) {
      var data = sheet.getRange(2, 1, last - 1, 1).getDisplayValues();
      for (var i = 0; i < data.length; i++) {
        var ex = normalizeStockId_(data[i][0]);
        if (ex) existing[ex] = true;
      }
    }
    var toAdd = [];
    var skipped = 0;
    var invalid = [];
    var seen = {};
    for (var j = 0; j < raw.length; j++) {
      var sid = normalizeStockId_(raw[j]);
      if (!sid) continue;
      if (!looksLikeStockId_(sid)) {
        invalid.push(sid);
        continue;
      }
      if (seen[sid]) continue;
      seen[sid] = true;
      if (existing[sid]) {
        skipped++;
        continue;
      }
      toAdd.push([sid, '']);
    }
    if (toAdd.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 2).setValues(toAdd);
      SpreadsheetApp.flush();
    }
    var msg = toAdd.length + ' hinzugefügt';
    if (skipped) msg += ', ' + skipped + ' schon vorhanden';
    if (invalid.length) msg += ', ' + invalid.length + ' ungültig (' + invalid.join(', ') + ')';
    return { success: true, message: msg, added: toAdd.length, skipped: skipped, invalid: invalid };
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  }
}

function forceRebuildCache() {
  return rebuildCacheChunk(0, 5000);
}
