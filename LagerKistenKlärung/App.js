var LAGER_SHEET_ID = '11d2YPM4wqLbGMkTCL7-lZJ1GXcHydNiNjOvRmkSxPEM';
var LAGER_TAB = 'LAGER';
var REFURB_SHEET_ID = '13Oh7gDT8NAul2s0cwQUeaGwMcS3B2MYu0QOdFNMhXzM';
var REFURB_SHEET_NAME = 'Refurbisment List';
var NACHBESTELL_SHEET_ID = '1VGCAHUbOPgsInQICA1GnrtKg1EPK1d1zWB-GkLi6iVE';
var NACHBESTELL_TAB = 'Nachbestellung';
var CACHE_TAB = '_KlärungCache';
var NOTES_TAB = '_KlärungNotes';
var CHECKS_TAB = '_KlärungChecks';
var CACHE_TTL_MS = 15 * 60 * 1000;
var CACHE_CHUNK = 48000;
var SYNC_STATUS_COLOR_TO_SHEET = true;
var CHECK_STEPS = ['carol', 'parts', 'mail'];
var CHECK_TOTAL = 3;

var COLOR_B2A1 = '#ff0000';
var COLOR_ALFAH = '#ff9900';
var COLOR_TAGESLISTE = '#00ff00';
var COLOR_KONTROLLIEREN = '#ffff00';
var COLOR_NACHBESTELLT = '#4a86e8';
var COLOR_COMPLETE = '#00ffff';
var COLOR_UEBERSICHT = '#9900ff';
var COLOR_DONE = '#cfd8dc';
var COLOR_RUCKFRAGE = COLOR_ALFAH;

var WEB_APP_URL = 'https://script.google.com/a/macros/auto1.com/s/AKfycbxw1Uz9cwh7Y3w5weaZ5wURxIlLaVS2wWa0BbfJ7AFnv0j5082SREKR7ylOpp7K4Lvl9Q/exec';
var GMAIL_ACCOUNT = 'ersatzteile.hemau@autohero.com';

function gmailAuthUserParam_() {
  return 'authuser=' + encodeURIComponent(GMAIL_ACCOUNT);
}

function gmailSearchUrl_(query) {
  return 'https://mail.google.com/mail/?' + gmailAuthUserParam_() + '#search/' + encodeURIComponent(String(query || ''));
}

function gmailThreadUrl_(threadId) {
  return 'https://mail.google.com/mail/?' + gmailAuthUserParam_() + '#inbox/' + String(threadId || '');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Lager Kisten Klärung')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Kisten Klärung')
    .addItem('App öffnen', 'openKlärungApp')
    .addToUi();
}

function openKlärungApp() {
  ensureCacheTrigger_();
  try {
    var cache = readCache_();
    if (!cache || isCacheStale_(cache)) rebuildKlärungCache();
  } catch (e0) {}

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><body style="margin:0"><script>' +
    'window.onload=function(){window.open(' + JSON.stringify(WEB_APP_URL) + ',"_blank");google.script.host.close();};' +
    '</script></body></html>'
  );
  SpreadsheetApp.getUi().showModalDialog(html, 'Lager Kisten Klärung');
}

function getWebAppUrl_() {
  return WEB_APP_URL;
}

function normalizeStockId_(value) {
  return String(value || '').replace(/\s+/g, '').toUpperCase();
}

function formatDateDe_(val) {
  if (!val && val !== 0) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' || val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, 'Europe/Berlin', 'dd.MM.yyyy');
  }
  return String(val).trim();
}

function formatReifenLabel_(val) {
  var s = String(val || '').trim();
  if (!s) return '';
  if (/^werkstatt\s*1$/i.test(s) || /^ws\s*1$/i.test(s)) return 'Reifen da';
  return s;
}

function nowStamp_() {
  return Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy HH:mm');
}

function activeUser_() {
  try {
    var a = Session.getActiveUser().getEmail();
    if (a) return String(a).trim();
  } catch (e0) {}
  try {
    var e = Session.getEffectiveUser().getEmail();
    if (e) return String(e).trim();
  } catch (e1) {}
  return '';
}

function isDoneComment_(comment) {
  var c = String(comment || '').toLowerCase();
  return c.indexOf('erledigt') !== -1 || c.indexOf('done') !== -1 || c.indexOf('raus') !== -1;
}

function normalizeHex_(hex) {
  var h = String(hex || '').trim().toLowerCase();
  if (!h) return '';
  if (h.charAt(0) !== '#') h = '#' + h;
  if (h.length === 4) {
    h = '#' + h.charAt(1) + h.charAt(1) + h.charAt(2) + h.charAt(2) + h.charAt(3) + h.charAt(3);
  }
  return h;
}

function knownBgLabel_(hex) {
  var h = normalizeHex_(hex);
  var map = {
    '#ff0000': 'B2A1',
    '#8b0000': 'B2A1',
    '#ffff00': 'Kontrollieren',
    '#fff176': 'Kontrollieren',
    '#4a86e8': 'Nachbestellt',
    '#4285f4': 'Nachbestellt',
    '#00ffff': 'COMPLETE',
    '#00bcd4': 'COMPLETE',
    '#00ff00': 'Tagesliste',
    '#34a853': 'Tagesliste',
    '#ff9900': 'Rückfrage ALFAH',
    '#ff9800': 'Rückfrage ALFAH',
    '#9900ff': 'Übersicht Nachbestellung',
    '#9c27b0': 'Übersicht Nachbestellung'
  };
  return map[h] || '';
}

function statusColor_(status) {
  var s = String(status || '').toLowerCase();
  if (s === 'b2a1') return COLOR_B2A1;
  if (s === 'alfah' || s === 'rückfrage' || s === 'ruckfrage' || s.indexOf('rückfrage') !== -1 || s.indexOf('ruckfrage') !== -1) return COLOR_ALFAH;
  if (s === 'complete') return COLOR_COMPLETE;
  if (s === 'tagesliste') return COLOR_TAGESLISTE;
  if (s === 'kontrollieren' || s === 'gelb') return COLOR_KONTROLLIEREN;
  if (s === 'nachbestellt') return COLOR_NACHBESTELLT;
  if (s === 'uebersicht' || s.indexOf('übersicht') !== -1 || s.indexOf('ubersicht') !== -1) return COLOR_UEBERSICHT;
  if (s === 'erledigt') return COLOR_DONE;
  return '';
}

function normalizeColor_(color) {
  var s = String(color || '').trim().toLowerCase();
  if (s === 'b2a1' || s === '#ff0000' || s === '#8b0000') return 'b2a1';
  if (s === 'alfah' || s === 'rückfrage' || s === 'ruckfrage' || s === 'orange' || s === '#ff9900' || s === '#ff9800') return 'alfah';
  if (s === 'nachbestellt' || s === 'blau' || s === 'blue' || s === '#4a86e8' || s === '#4285f4') return 'nachbestellt';
  if (s === 'kontrollieren' || s === 'gelb' || s === 'yellow' || s === '#ffff00' || s === '#fff176') return 'kontrollieren';
  if (s === 'complete' || s === '#00ffff' || s === '#00bcd4') return 'complete';
  if (s === 'tagesliste' || s === '#00ff00' || s === '#34a853') return 'tagesliste';
  if (s.indexOf('übersicht') !== -1 || s.indexOf('ubersicht') !== -1 || s === 'uebersicht' || s === '#9900ff' || s === '#9c27b0') return 'uebersicht';
  return 'none';
}

function colorLabel_(color) {
  var c = normalizeColor_(color);
  if (c === 'b2a1') return 'B2A1';
  if (c === 'alfah') return 'ALFAH';
  if (c === 'nachbestellt') return 'Nachbestellt';
  if (c === 'kontrollieren') return 'Kontrollieren';
  if (c === 'complete') return 'COMPLETE';
  if (c === 'tagesliste') return 'Tagesliste';
  if (c === 'uebersicht') return 'Übersicht Nachbestellung';
  return '';
}

function colorPriority_(color) {
  var c = normalizeColor_(color);
  if (c === 'b2a1') return 4;
  if (c === 'alfah') return 3;
  if (c === 'nachbestellt') return 2;
  if (c === 'kontrollieren') return 1;
  return 0;
}

function normalizeStatus_(status) {
  return colorLabel_(normalizeColor_(status)) || String(status || '').trim();
}

function shortName_(email) {
  var e = String(email || '').trim();
  if (!e) return '';
  var local = e.split('@')[0] || e;
  return local.replace(/\./g, ' ');
}

function formatCommentAt_(val) {
  if (!val && val !== 0) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' || val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, 'Europe/Berlin', 'dd.MM.yyyy HH:mm');
  }
  var s = String(val).trim();
  if (/GMT|Mitteleurop|UTC/i.test(s)) {
    var d = new Date(s);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'Europe/Berlin', 'dd.MM.yyyy HH:mm');
  }
  return s;
}

function canDeleteComment_(by, me) {
  by = String(by || '').trim().toLowerCase();
  me = String(me || '').trim().toLowerCase();
  if (!by) return true;
  if (!me) return true;
  return me === by;
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

function lagerIsRegalHeader_(v) {
  return /^\s*regal\s+\d+\.\d+\s*$/i.test(String(v == null ? '' : v));
}

function lagerCategoryLabel_(v) {
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  var up = s.toUpperCase();
  var cats = ['B2A1', 'COMPLETE', 'TAGESLISTE', 'KONTROLLIEREN', 'NACHBESTELLT'];
  if (cats.indexOf(up) !== -1) return s;
  if (up.indexOf('RÜCKFRAGE') !== -1 || up.indexOf('RUCKFRAGE') !== -1) return s;
  if (up.indexOf('ÜBERSICHT') !== -1 || up.indexOf('UBERSICHT') !== -1) return s;
  return '';
}

function lagerNormalizeRegal_(v) {
  var s = String(v == null ? '' : v).trim();
  var m = s.match(/(\d+)\.(\d+)/);
  if (m) return 'Regal ' + m[1] + '.' + m[2];
  return s;
}

function lagerLooksLikeStockId_(v) {
  var s = String(v == null ? '' : v).replace(/\s+/g, '').toUpperCase();
  return /^[A-Z]{2}\d{4,8}$/.test(s);
}

function getLagerSheet_() {
  var ss = SpreadsheetApp.openById(LAGER_SHEET_ID);
  var sheet = ss.getSheetByName(LAGER_TAB);
  if (!sheet) throw new Error("Tab 'LAGER' nicht gefunden");
  return sheet;
}

function getOrCreateTab_(name) {
  var ss = SpreadsheetApp.openById(LAGER_SHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}

function getCacheSheet_() {
  return getOrCreateTab_(CACHE_TAB);
}

function getNotesSheet_() {
  var sh = getOrCreateTab_(NOTES_TAB);
  var header = String(sh.getRange(1, 1).getValue() || '');
  if (header !== 'id') {
    sh.clear();
    sh.getRange(1, 1, 1, 7).setValues([['id', 'stockId', 'cellKey', 'color', 'comment', 'createdAt', 'createdBy']]);
  }
  return sh;
}

function getChecksSheet_() {
  var sh = getOrCreateTab_(CHECKS_TAB);
  var header = String(sh.getRange(1, 1).getValue() || '');
  if (header !== 'cellKey') {
    sh.clear();
    sh.getRange(1, 1, 1, 11).setValues([[
      'cellKey', 'stockId',
      'carol', 'carolBy', 'carolAt',
      'parts', 'partsBy', 'partsAt',
      'mail', 'mailBy', 'mailAt'
    ]]);
  }
  return sh;
}

function emptyChecks_() {
  return {
    carol: false, carolBy: '', carolAt: '',
    parts: false, partsBy: '', partsAt: '',
    mail: false, mailBy: '', mailAt: '',
    done: 0,
    total: CHECK_TOTAL
  };
}

function countChecks_(ch) {
  var n = 0;
  if (ch.carol) n++;
  if (ch.parts) n++;
  if (ch.mail) n++;
  ch.done = n;
  ch.total = CHECK_TOTAL;
  return ch;
}

function readChecksMap_() {
  var map = {};
  try {
    var sh = getChecksSheet_();
    var last = sh.getLastRow();
    if (last < 2) return map;
    var data = sh.getRange(2, 1, last, 11).getValues();
    for (var i = 0; i < data.length; i++) {
      var cellKey = String(data[i][0] || '').trim();
      if (!cellKey) continue;
      var ch = emptyChecks_();
      ch.carol = String(data[i][2]) === '1' || data[i][2] === true || data[i][2] === 1;
      ch.carolBy = String(data[i][3] || '').trim();
      ch.carolAt = String(data[i][4] || '').trim();
      ch.parts = String(data[i][5]) === '1' || data[i][5] === true || data[i][5] === 1;
      ch.partsBy = String(data[i][6] || '').trim();
      ch.partsAt = String(data[i][7] || '').trim();
      ch.mail = String(data[i][8]) === '1' || data[i][8] === true || data[i][8] === 1;
      ch.mailBy = String(data[i][9] || '').trim();
      ch.mailAt = String(data[i][10] || '').trim();
      countChecks_(ch);
      map[cellKey] = ch;
    }
  } catch (e) {}
  return map;
}

function pruneChecks_(activeCellKeys) {
  try {
    var sh = getChecksSheet_();
    var last = sh.getLastRow();
    if (last < 2) return;
    var data = sh.getRange(2, 1, last, 1).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      var ck = String(data[i][0] || '').trim();
      if (!ck || !activeCellKeys[ck]) sh.deleteRow(i + 2);
    }
  } catch (e) {}
}

function newCommentId_() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 12);
}

function readAllComments_() {
  var list = [];
  try {
    var sh = getNotesSheet_();
    var last = sh.getLastRow();
    if (last < 2) return list;
    var data = sh.getRange(2, 1, last, 7).getValues();
    var me = activeUser_();
    for (var i = 0; i < data.length; i++) {
      var id = String(data[i][0] || '').trim();
      var stockId = normalizeStockId_(data[i][1]);
      if (!id || !stockId) continue;
      var by = String(data[i][6] || '').trim();
      var name = shortName_(by) || (by ? by : 'Team');
      list.push({
        id: id,
        stockId: stockId,
        cellKey: String(data[i][2] || '').trim(),
        color: normalizeColor_(data[i][3]),
        colorLabel: colorLabel_(data[i][3]),
        comment: String(data[i][4] == null ? '' : data[i][4]).trim(),
        createdAt: formatCommentAt_(data[i][5]),
        createdBy: by,
        createdByName: name,
        canDelete: canDeleteComment_(by, me),
        sheetRow: i + 2
      });
    }
  } catch (e) {}
  return list;
}

function commentsForStock_(all, stockId, cellKey) {
  stockId = normalizeStockId_(stockId);
  cellKey = String(cellKey || '');
  var out = [];
  for (var i = 0; i < all.length; i++) {
    var c = all[i];
    if (cellKey && c.cellKey === cellKey) out.push(c);
    else if (!c.cellKey && c.stockId === stockId) out.push(c);
    else if (c.stockId === stockId && (!cellKey || !c.cellKey || c.cellKey === cellKey)) out.push(c);
  }
  var seen = {};
  var uniq = [];
  for (var u = 0; u < out.length; u++) {
    if (seen[out[u].id]) continue;
    seen[out[u].id] = true;
    uniq.push(out[u]);
  }
  uniq.sort(function(a, b) {
    return String(b.createdAt).localeCompare(String(a.createdAt));
  });
  return uniq;
}

function primaryFromComments_(comments) {
  var best = null;
  var bestP = -1;
  for (var i = 0; i < comments.length; i++) {
    var p = colorPriority_(comments[i].color);
    if (p > bestP) {
      bestP = p;
      best = comments[i];
    }
  }
  return best;
}

function pruneToolNotes_(activeCellKeys, activeStockIds) {
  try {
    var sh = getNotesSheet_();
    var last = sh.getLastRow();
    if (last < 2) return;
    var data = sh.getRange(2, 1, last, 3).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      var sid = normalizeStockId_(data[i][1]);
      var ck = String(data[i][2] || '').trim();
      var keep = false;
      if (ck && activeCellKeys[ck]) keep = true;
      if (sid && activeStockIds[sid]) keep = true;
      if (!keep) sh.deleteRow(i + 2);
    }
  } catch (e) {}
}

function isCacheStale_(cache) {
  if (!cache || !cache.builtAtMs) return true;
  return (Date.now() - Number(cache.builtAtMs)) > CACHE_TTL_MS;
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
}

function readCache_() {
  try {
    var sh = SpreadsheetApp.openById(LAGER_SHEET_ID).getSheetByName(CACHE_TAB);
    if (!sh || sh.getLastRow() < 2) return null;
    var meta = sh.getRange(1, 1, 1, 2).getValues()[0];
    var parts = parseInt(meta[1], 10) || 0;
    if (parts < 1) return null;
    var chunks = sh.getRange(2, 1, parts, 1).getValues();
    var json = '';
    for (var i = 0; i < chunks.length; i++) json += String(chunks[i][0] || '');
    if (!json) return null;
    return JSON.parse(json);
  } catch (err) {
    return null;
  }
}

function ensureCacheTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'rebuildKlärungCache') return;
  }
  ScriptApp.newTrigger('rebuildKlärungCache').timeBased().everyMinutes(10).create();
}

function installCacheTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'rebuildKlärungCache') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('rebuildKlärungCache').timeBased().everyMinutes(10).create();
  var res = rebuildKlärungCache();
  SpreadsheetApp.getUi().alert('Cache-Trigger aktiv.\n' + (res.items ? res.items.length : 0) + ' Kisten aus LAGER.');
}

function parseLagerGrid_() {
  var sheet = getLagerSheet_();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();
  var numRows = values.length;
  var numCols = numRows ? values[0].length : 0;

  var colorLabelMap = {
    '#ff0000': 'B2A1',
    '#8b0000': 'B2A1',
    '#ffff00': 'Kontrollieren',
    '#fff176': 'Kontrollieren',
    '#4a86e8': 'Nachbestellt',
    '#4285f4': 'Nachbestellt',
    '#00ffff': 'COMPLETE',
    '#00bcd4': 'COMPLETE',
    '#00ff00': 'Tagesliste',
    '#34a853': 'Tagesliste',
    '#ff9900': 'Rückfrage ALFAH',
    '#ff9800': 'Rückfrage ALFAH',
    '#9900ff': 'Übersicht Nachbestellung',
    '#9c27b0': 'Übersicht Nachbestellung'
  };
  for (var lr = 0; lr < numRows; lr++) {
    for (var lc = 0; lc < numCols; lc++) {
      var label = lagerCategoryLabel_(values[lr][lc]);
      if (!label) continue;
      var hex = normalizeHex_(backgrounds[lr][lc]);
      if (hex && hex !== '#ffffff' && hex !== '#fff') {
        colorLabelMap[hex] = label;
      }
    }
  }

  var items = [];
  for (var c = 0; c < numCols; c++) {
    var currentRegal = '';
    for (var r = 0; r < numRows; r++) {
      var raw = values[r][c];
      if (lagerIsRegalHeader_(raw)) {
        currentRegal = lagerNormalizeRegal_(raw);
        continue;
      }
      if (lagerCategoryLabel_(raw)) continue;
      var txt = String(raw == null ? '' : raw).trim();
      if (!txt || !lagerLooksLikeStockId_(txt)) continue;
      var bg = normalizeHex_(backgrounds[r][c]);
      items.push({
        stockId: normalizeStockId_(txt),
        regal: currentRegal,
        sheetKategorie: colorLabelMap[bg] || knownBgLabel_(bg) || '',
        sheetBg: bg,
        sheetRow: r + 1,
        sheetCol: c + 1,
        cellKey: (r + 1) + ':' + (c + 1)
      });
    }
  }

  items.sort(function(a, b) {
    var ra = String(a.regal || '');
    var rb = String(b.regal || '');
    if (ra < rb) return -1;
    if (ra > rb) return 1;
    return String(a.stockId).localeCompare(String(b.stockId));
  });
  return { items: items, colorLabelMap: colorLabelMap };
}

function statusFromSheetKat_(kat) {
  return normalizeStatus_(kat);
}

function rebuildKlärungCache() {
  var builtAtMs = Date.now();
  var builtAt = Utilities.formatDate(new Date(builtAtMs), 'Europe/Berlin', 'dd.MM.yyyy HH:mm:ss');
  var grid = parseLagerGrid_();
  var allComments = readAllComments_();
  var checksMap = readChecksMap_();
  var refurbMap = buildRefurbMap_();
  var nbMap = buildNachbestellMap_();
  var items = [];
  var details = {};
  var activeCells = {};
  var activeStocks = {};

  for (var i = 0; i < grid.items.length; i++) {
    var g = grid.items[i];
    var stockId = g.stockId;
    var cellKey = g.cellKey;
    activeCells[cellKey] = true;
    activeStocks[stockId] = true;

    var comments = commentsForStock_(allComments, stockId, cellKey);
    var primary = primaryFromComments_(comments);
    var toolStatus = primary ? colorLabel_(primary.color) : '';
    var sheetStatus = statusFromSheetKat_(g.sheetKategorie);
    var statusHint = toolStatus || sheetStatus || '';
    var preview = comments.length ? comments[0].comment : '';
    var hasNote = comments.length > 0;
    var yellow = statusHint === 'ALFAH' || statusHint === 'Kontrollieren';
    var refurb = refurbMap[stockId] || { found: false };
    var nbs = nbMap[stockId] || [];
    var checks = countChecks_(checksMap[cellKey] ? JSON.parse(JSON.stringify(checksMap[cellKey])) : emptyChecks_());

    var kisten = {
      row: g.sheetRow,
      col: g.sheetCol,
      cellKey: cellKey,
      stockId: stockId,
      comments: comments,
      comment: preview,
      status: statusHint,
      primaryColor: primary ? primary.color : '',
      sheetKategorie: g.sheetKategorie || '',
      yellow: yellow,
      hasNote: hasNote,
      checks: checks
    };

    var item = {
      row: g.sheetRow,
      col: g.sheetCol,
      cellKey: cellKey,
      stockId: stockId,
      comment: preview,
      hasNote: hasNote,
      commentCount: comments.length,
      yellow: yellow,
      done: checks.done >= CHECK_TOTAL,
      statusHint: statusHint,
      primaryColor: primary ? primary.color : (sheetStatus ? normalizeColor_(sheetStatus) : ''),
      regal: g.regal || '',
      kategorie: statusHint || '',
      sheetKategorie: g.sheetKategorie || '',
      markeModel: refurb.markeModel || '',
      refurbStatus: refurb.status || '',
      reifenStatus: refurb.reifenStatus || '',
      nbCount: nbs.length,
      checksDone: checks.done,
      checksTotal: CHECK_TOTAL
    };
    items.push(item);
    details[cellKey] = {
      stockId: stockId,
      kisten: kisten,
      regal: g.regal || '',
      refurb: refurb,
      nachbestellungen: nbs,
      checks: checks,
      gmailSearchUrl: gmailSearchUrl_(stockId),
      carolSearchHint: 'Carol: begonnen / Fertiggestellt / B2A1 prüfen',
      fromCache: true
    };
  }

  pruneToolNotes_(activeCells, activeStocks);
  pruneChecks_(activeCells);

  var payload = {
    builtAt: builtAt,
    builtAtMs: builtAtMs,
    items: items,
    details: details,
    count: items.length,
    source: 'LAGER+Comments'
  };
  writeCachePayload_(payload);
  return payload;
}

function buildRefurbMap_() {
  var map = {};
  try {
    var ss = SpreadsheetApp.openById(REFURB_SHEET_ID);
    var sheet = ss.getSheetByName(REFURB_SHEET_NAME);
    if (!sheet) return map;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return map;
    var numRows = lastRow - 1;
    var data = sheet.getRange(2, 1, numRows, 30).getValues();
    var formulas = sheet.getRange(2, 3, numRows, 1).getFormulas();
    var richVals = sheet.getRange(2, 3, numRows, 1).getRichTextValues();
    for (var i = 0; i < data.length; i++) {
      var stockId = normalizeStockId_(data[i][1]);
      if (!stockId || map[stockId]) continue;
      map[stockId] = {
        found: true,
        row: i + 2,
        carolUrl: carolUrlFromSheetParts_(richVals[i][0], formulas[i][0], data[i][2]),
        markeModel: String(data[i][12] || '').trim(),
        schaeden: String(data[i][22] || '').trim(),
        kommBestellung: String(data[i][23] || '').trim(),
        kommAnlieferung: String(data[i][24] || '').trim(),
        status: String(data[i][25] || '').trim(),
        regal: String(data[i][27] || '').trim(),
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
    for (var h = 0; h < header.length; h++) {
      var t = String(header[h] || '').toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
      if (t.indexOf('stock') !== -1) stockCol = h + 1;
      if (t.indexOf('ersatzteil') !== -1 || t === 'teil' || t.indexOf('benennung') !== -1 || t.indexOf('bezeichnung') !== -1) teilCol = h + 1;
      if (t === 'status') statusCol = h + 1;
      if (t.indexOf('lagerort') !== -1 || t === 'regal') regalCol = h + 1;
      if (t === 'art' || t === 'typ' || t.indexOf('artdernachbestellung') !== -1) typCol = h + 1;
      if (t === 'datum' || t === 'date' || t.indexOf('bestelldatum') !== -1 || t.indexOf('erstellt') !== -1) dateCol = h + 1;
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
      if (!map[sid]) map[sid] = [];
      map[sid].push({
        source: 'Nachbestellung',
        typ: String(typVal || '').trim(),
        teil: String(data[i][teilCol - 1] || '').trim(),
        status: String(data[i][statusCol - 1] || '').trim(),
        regal: String(data[i][regalCol - 1] || '').trim(),
        date: dateStr,
        dateMs: dateMs,
        sheetRow: headerRow + 1 + i
      });
    }
    var keys = Object.keys(map);
    for (var k = 0; k < keys.length; k++) {
      var list = map[keys[k]];
      list.sort(function(a, b) {
        return (b.dateMs || 0) - (a.dateMs || 0);
      });
      var currentIdx = -1;
      for (var n = 0; n < list.length; n++) {
        var st = String(list[n].status || '').toLowerCase();
        var closed = st.indexOf('komplett') !== -1 || st.indexOf('fertiggestellt') !== -1 || st === 'angeliefert' || st.indexOf('fahrzeug rr') !== -1;
        if (!closed) {
          currentIdx = n;
          break;
        }
      }
      if (currentIdx < 0 && list.length) currentIdx = 0;
      for (var m = 0; m < list.length; m++) {
        list[m].current = m === currentIdx;
      }
    }
  } catch (e1) {}
  return map;
}

function filterItems_(items, filter) {
  filter = String(filter || 'open').toLowerCase();
  var out = [];
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var hint = String(it.statusHint || '').toLowerCase();
    var col = String(it.primaryColor || '').toLowerCase();
    if (filter === 'open' && it.done) continue;
    if (filter === 'open') { out.push(it); continue; }
    if (filter === 'yellow' && !(col === 'kontrollieren' || col === 'alfah' || it.yellow)) continue;
    if (filter === 'tagesliste' && hint.indexOf('tagesliste') === -1) continue;
    if (filter === 'nachbestellt') {
      if (col === 'uebersicht' || hint.indexOf('übersicht') !== -1 || hint.indexOf('ubersicht') !== -1) continue;
      if (col !== 'nachbestellt' && hint.indexOf('nachbestellt') === -1) continue;
    }
    if (filter === 'uebersicht' && col !== 'uebersicht' && hint.indexOf('übersicht') === -1 && hint.indexOf('ubersicht') === -1) continue;
    if (filter === 'kontrollieren' && col !== 'kontrollieren') continue;
    if (filter === 'b2a1' && col !== 'b2a1' && hint.indexOf('b2a1') === -1) continue;
    if (filter === 'alfah' && col !== 'alfah') continue;
    if (filter === 'hasnote' && !it.hasNote) continue;
    if (filter === 'nocomment' && it.hasNote) continue;
    if (filter === 'done' && !it.done) continue;
    out.push(it);
  }
  return out;
}

function getQueue(filter) {
  try {
    var cache = readCache_();
    if (!cache || !cache.items) cache = rebuildKlärungCache();
    var items = filterItems_(cache.items || [], filter);
    return {
      success: true,
      items: items,
      count: items.length,
      totalCached: (cache.items || []).length,
      cachedAt: cache.builtAt || '',
      source: cache.source || 'LAGER',
      fromCache: true
    };
  } catch (err) {
    return { success: false, message: String(err.message || err), items: [] };
  }
}

function getStockDetail(stockId, cellKey) {
  try {
    stockId = normalizeStockId_(stockId);
    cellKey = String(cellKey || '');
    if (!stockId) return { success: false, message: 'Keine Stock-ID' };
    var cache = readCache_();
    if (cache && cache.details && cellKey && cache.details[cellKey]) {
      return detailFromCache_(cache, cache.details[cellKey]);
    }
    if (cache && cache.details) {
      var keys = Object.keys(cache.details);
      for (var i = 0; i < keys.length; i++) {
        if (cache.details[keys[i]].stockId === stockId) {
          return detailFromCache_(cache, cache.details[keys[i]]);
        }
      }
    }
    return { success: false, message: 'Nicht im Cache — rebuildKlärungCache ausführen' };
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  }
}

function detailFromCache_(cache, d) {
  var checks = d.checks || (d.kisten && d.kisten.checks) || emptyChecks_();
  countChecks_(checks);
  return {
    success: true,
    stockId: d.stockId,
    kisten: d.kisten,
    regal: d.regal,
    refurb: d.refurb,
    nachbestellungen: d.nachbestellungen || [],
    checks: checks,
    gmail: { threads: [], orderHits: [], ok: true, message: 'Mails on-demand', alfahUnanswered: false },
    gmailSearchUrl: d.gmailSearchUrl,
    carolSearchHint: d.carolSearchHint,
    fromCache: true,
    cachedAt: cache.builtAt || ''
  };
}

function loadGmailForStock(stockId, kommBestellung) {
  return searchGmailForStock_(normalizeStockId_(stockId), kommBestellung);
}

function searchGmailForStock_(stockId, kommBestellung) {
  var result = { threads: [], orderHits: [], alfahUnanswered: false, message: '', ok: false };
  try {
    var terms = ['"' + stockId + '"'];
    var orderNums = String(kommBestellung || '').match(/\b(?:N4P)?\d{6,12}\b/gi) || [];
    var seenOrd = {};
    for (var o = 0; o < orderNums.length && o < 6; o++) {
      var on = String(orderNums[o]).replace(/\s+/g, '').toUpperCase();
      if (seenOrd[on]) continue;
      seenOrd[on] = true;
      terms.push('"' + on + '"');
      result.orderHits.push(on);
    }
    var seen = {};
    var threads = [];
    for (var t = 0; t < terms.length; t++) {
      var batch = GmailApp.search(terms[t], 0, 8);
      for (var i = 0; i < batch.length; i++) {
        var id = batch[i].getId();
        if (seen[id]) continue;
        seen[id] = true;
        threads.push(batch[i]);
      }
    }
    threads.sort(function(a, b) {
      return b.getLastMessageDate().getTime() - a.getLastMessageDate().getTime();
    });
    for (var j = 0; j < Math.min(12, threads.length); j++) {
      var th = threads[j];
      var msgs = th.getMessages();
      var last = msgs[msgs.length - 1];
      var from = String(last.getFrom() || '');
      var isAlfah = /alfah/i.test(from) || /alfah/i.test(th.getFirstMessageSubject());
      var unanswered = isAlfah && /alfah/i.test(from);
      if (unanswered) result.alfahUnanswered = true;
      result.threads.push({
        subject: String(th.getFirstMessageSubject() || ''),
        from: from,
        date: Utilities.formatDate(last.getDate(), 'Europe/Berlin', 'dd.MM.yyyy HH:mm'),
        snippet: String(last.getPlainBody() || '').replace(/\s+/g, ' ').substring(0, 180),
        isAlfah: isAlfah,
        unanswered: unanswered,
        permalink: gmailThreadUrl_(th.getId())
      });
    }
    result.ok = true;
    result.message = result.threads.length ? (result.threads.length + ' Mail(s)') : 'Keine Mails gefunden';
  } catch (err) {
    result.message = 'Gmail: ' + String(err.message || err);
  }
  return result;
}

function parseCellKey_(cellKey) {
  var parts = String(cellKey || '').split(':');
  if (parts.length !== 2) return null;
  var row = parseInt(parts[0], 10);
  var col = parseInt(parts[1], 10);
  if (!(row > 0) || !(col > 0)) return null;
  return { row: row, col: col };
}

function syncSheetColor_(cellKey, color) {
  if (!SYNC_STATUS_COLOR_TO_SHEET) return;
  var pos = parseCellKey_(cellKey);
  if (!pos) return;
  try {
    var cell = getLagerSheet_().getRange(pos.row, pos.col);
    var hex = statusColor_(color);
    if (hex) cell.setBackground(hex);
    else cell.setBackground(null);
    SpreadsheetApp.flush();
  } catch (e) {}
}

function readLiveSheetKat_(cellKey) {
  var pos = parseCellKey_(cellKey);
  if (!pos) return '';
  try {
    var bg = normalizeHex_(getLagerSheet_().getRange(pos.row, pos.col).getBackground());
    if (!bg || bg === '#ffffff' || bg === '#fff' || bg === '#00000000') return '';
    return knownBgLabel_(bg) || '';
  } catch (e) {
    return '';
  }
}

function refreshCellInCache_(cellKey, stockId) {
  try {
    var cache = readCache_();
    if (!cache || !cache.details || !cache.details[cellKey]) {
      rebuildKlärungCache();
      return;
    }
    var all = readAllComments_();
    var comments = commentsForStock_(all, stockId, cellKey);
    var primary = primaryFromComments_(comments);
    var liveKat = readLiveSheetKat_(cellKey);
    var toolStatus = primary ? colorLabel_(primary.color) : '';
    var statusHint = toolStatus || statusFromSheetKat_(liveKat) || '';
    var preview = comments.length ? comments[0].comment : '';
    cache.details[cellKey].kisten.comments = comments;
    cache.details[cellKey].kisten.comment = preview;
    cache.details[cellKey].kisten.hasNote = comments.length > 0;
    cache.details[cellKey].kisten.status = statusHint;
    cache.details[cellKey].kisten.primaryColor = primary ? primary.color : '';
    cache.details[cellKey].kisten.sheetKategorie = liveKat;
    for (var i = 0; i < cache.items.length; i++) {
      if (cache.items[i].cellKey !== cellKey) continue;
      cache.items[i].comment = preview;
      cache.items[i].hasNote = comments.length > 0;
      cache.items[i].commentCount = comments.length;
      cache.items[i].statusHint = statusHint;
      cache.items[i].primaryColor = primary ? primary.color : '';
      cache.items[i].kategorie = statusHint;
      cache.items[i].sheetKategorie = liveKat;
      cache.items[i].yellow = statusHint === 'ALFAH' || statusHint === 'Kontrollieren';
      break;
    }
    writeCachePayload_(cache);
  } catch (e) {
    rebuildKlärungCache();
  }
}

function forceRebuildCache() {
  try {
    var res = rebuildKlärungCache();
    return {
      success: true,
      message: 'Cache neu gebaut',
      count: (res && res.items) ? res.items.length : 0,
      cachedAt: (res && res.builtAt) || ''
    };
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  }
}

function addStockComment(cellKey, stockId, color, text) {
  try {
    cellKey = String(cellKey || '');
    stockId = normalizeStockId_(stockId);
    text = String(text || '').trim();
    color = normalizeColor_(color);
    if (!cellKey || !stockId) return { success: false, message: 'Stock/Zelle fehlt' };
    if (!text) return { success: false, message: 'Leerer Kommentar' };

    var id = newCommentId_();
    var by = activeUser_();
    var at = nowStamp_();
    var sh = getNotesSheet_();
    sh.appendRow([id, stockId, cellKey, color, text, at, by]);
    var row = sh.getLastRow();
    try {
      sh.getRange(row, 6).setNumberFormat('@').setValue(at);
      if (by) sh.getRange(row, 7).setNumberFormat('@').setValue(by);
    } catch (eFmt) {}
    syncSheetColor_(cellKey, color);
    SpreadsheetApp.flush();
    refreshCellInCache_(cellKey, stockId);

    return {
      success: true,
      message: 'Kommentar gespeichert',
      comment: {
        id: id,
        stockId: stockId,
        cellKey: cellKey,
        color: color,
        colorLabel: colorLabel_(color),
        comment: text,
        createdAt: at,
        createdBy: by,
        createdByName: shortName_(by) || 'Ich',
        canDelete: true
      }
    };
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  }
}

function deleteStockComment(commentId) {
  try {
    commentId = String(commentId || '').trim();
    if (!commentId) return { success: false, message: 'Keine Kommentar-ID' };
    var sh = getNotesSheet_();
    var last = sh.getLastRow();
    if (last < 2) return { success: false, message: 'Kein Kommentar' };
    var data = sh.getRange(2, 1, last, 7).getValues();
    var me = activeUser_();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim() !== commentId) continue;
      var by = String(data[i][6] || '').trim();
      if (!canDeleteComment_(by, me)) {
        return { success: false, message: 'Nur eigener Kommentar löschbar' };
      }
      var stockId = normalizeStockId_(data[i][1]);
      var cellKey = String(data[i][2] || '').trim();
      sh.deleteRow(i + 2);
      var left = commentsForStock_(readAllComments_(), stockId, cellKey);
      var primary = primaryFromComments_(left);
      syncSheetColor_(cellKey, primary ? primary.color : '');
      SpreadsheetApp.flush();
      refreshCellInCache_(cellKey, stockId);
      return { success: true, message: 'Kommentar gelöscht', cellKey: cellKey, stockId: stockId };
    }
    return { success: false, message: 'Kommentar nicht gefunden' };
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  }
}

function getStockComments(cellKey, stockId) {
  try {
    var comments = commentsForStock_(readAllComments_(), stockId, cellKey);
    return { success: true, comments: comments };
  } catch (err) {
    return { success: false, message: String(err.message || err), comments: [] };
  }
}

function setCheckStep(cellKey, stockId, step, checked) {
  try {
    cellKey = String(cellKey || '');
    stockId = normalizeStockId_(stockId);
    step = String(step || '').toLowerCase();
    if (!cellKey || !stockId) return { success: false, message: 'Stock/Zelle fehlt' };
    if (CHECK_STEPS.indexOf(step) === -1) return { success: false, message: 'Ungültiger Step' };

    var sh = getChecksSheet_();
    var last = sh.getLastRow();
    var hit = -1;
    var row = emptyChecks_();
    if (last >= 2) {
      var keys = sh.getRange(2, 1, last, 11).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (String(keys[i][0] || '').trim() !== cellKey) continue;
        hit = i + 2;
        row.carol = String(keys[i][2]) === '1' || keys[i][2] === true || keys[i][2] === 1;
        row.carolBy = String(keys[i][3] || '');
        row.carolAt = String(keys[i][4] || '');
        row.parts = String(keys[i][5]) === '1' || keys[i][5] === true || keys[i][5] === 1;
        row.partsBy = String(keys[i][6] || '');
        row.partsAt = String(keys[i][7] || '');
        row.mail = String(keys[i][8]) === '1' || keys[i][8] === true || keys[i][8] === 1;
        row.mailBy = String(keys[i][9] || '');
        row.mailAt = String(keys[i][10] || '');
        break;
      }
    }

    var on = !!checked;
    var who = on ? activeUser_() : '';
    var when = on ? nowStamp_() : '';
    if (step === 'carol') { row.carol = on; row.carolBy = who; row.carolAt = when; }
    if (step === 'parts') { row.parts = on; row.partsBy = who; row.partsAt = when; }
    if (step === 'mail') { row.mail = on; row.mailBy = who; row.mailAt = when; }
    countChecks_(row);

    var vals = [[
      cellKey, stockId,
      row.carol ? 1 : 0, row.carolBy, row.carolAt,
      row.parts ? 1 : 0, row.partsBy, row.partsAt,
      row.mail ? 1 : 0, row.mailBy, row.mailAt
    ]];
    if (hit > 0) sh.getRange(hit, 1, 1, 11).setValues(vals);
    else sh.appendRow(vals[0]);
    SpreadsheetApp.flush();

    try {
      var cache = readCache_();
      if (cache && cache.details && cache.details[cellKey]) {
        cache.details[cellKey].checks = row;
        if (cache.details[cellKey].kisten) cache.details[cellKey].kisten.checks = row;
        for (var j = 0; j < cache.items.length; j++) {
          if (cache.items[j].cellKey !== cellKey) continue;
          cache.items[j].checksDone = row.done;
          cache.items[j].checksTotal = CHECK_TOTAL;
          cache.items[j].done = row.done >= CHECK_TOTAL;
          break;
        }
        writeCachePayload_(cache);
      }
    } catch (e2) {}

    return {
      success: true,
      checks: row,
      message: row.done + '/' + CHECK_TOTAL + ' kontrolliert'
    };
  } catch (err) {
    return { success: false, message: String(err.message || err) };
  }
}
