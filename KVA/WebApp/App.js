var KVA_VERSION = '0.1.2';
var SOURCE_SHEET_ID = '13Oh7gDT8NAul2s0cwQUeaGwMcS3B2MYu0QOdFNMhXzM';
var SOURCE_TAB = 'Carol_Teilestatus';
var HOST_SHEET_ID = '1zeKjqR0-b9zBiHjmGWLq24DM4Qu5NV2rd_NJVn3TwO8';
var CACHE_TAB = '_KvaCache';
var CACHE_CHUNK = 45000;
var CACHE_MAX_AGE_MS = 60000;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('KVA HUD')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('KVA')
    .addItem('KVA HUD öffnen', 'openKva')
    .addToUi();
}

function openKva() {
  var url = '';
  try {
    url = ScriptApp.getService().getUrl() || '';
  } catch (e) {
    url = '';
  }
  if (!url) {
    SpreadsheetApp.getUi().alert('Die Web-App ist noch nicht bereitgestellt.');
    return;
  }
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><body style="margin:0"><script>window.onload=function(){window.open(' +
      JSON.stringify(url) +
      ',"_blank");google.script.host.close();};<\/script></body></html>'
  );
  SpreadsheetApp.getUi().showModalDialog(html, 'KVA HUD');
}

function getKvaBoard(force) {
  ensureCacheTrigger_();
  if (!force) {
    var cached = readCache_();
    if (cached && cacheFresh_(cached)) return cached;
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    var stale = readCache_();
    if (stale) return stale;
    throw e;
  }
  try {
    if (!force) {
      var again = readCache_();
      if (again && cacheFresh_(again)) return again;
    }
    var board = buildBoard_();
    writeCache_(board);
    return board;
  } finally {
    lock.releaseLock();
  }
}

function warmKvaCache() {
  getKvaBoard(true);
}

function ensureCacheTrigger_() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'warmKvaCache') return;
    }
    ScriptApp.newTrigger('warmKvaCache').timeBased().everyMinutes(1).create();
  } catch (e) {}
}

function cacheFresh_(board) {
  var at = Number(board && board.loadedAtMs);
  if (!at) return false;
  return (Date.now() - at) < CACHE_MAX_AGE_MS;
}

function cacheSpreadsheet_() {
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  return SpreadsheetApp.openById(HOST_SHEET_ID);
}

function cacheSheet_(create) {
  var ss = null;
  try { ss = cacheSpreadsheet_(); } catch (e) { ss = null; }
  if (!ss) return null;
  var sh = ss.getSheetByName(CACHE_TAB);
  if (sh || !create) return sh;
  sh = ss.insertSheet(CACHE_TAB);
  try { sh.hideSheet(); } catch (e2) {}
  return sh;
}

function readCache_() {
  var sh = null;
  try { sh = cacheSheet_(false); } catch (e) { return null; }
  if (!sh) return null;
  var meta = sh.getRange(1, 1, 1, 5).getValues()[0];
  var chunks = Number(meta[3]);
  if (!chunks) return null;
  var values = sh.getRange(2, 1, chunks, 1).getValues();
  var json = '';
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0]) return null;
    json += values[i][0];
  }
  try {
    return JSON.parse(json);
  } catch (e2) {
    return null;
  }
}

function writeCache_(board) {
  var sh = null;
  try { sh = cacheSheet_(true); } catch (e) { return; }
  if (!sh) return;
  var json = JSON.stringify(board);
  var chunks = [];
  for (var i = 0; i < json.length; i += CACHE_CHUNK) chunks.push([json.substring(i, i + CACHE_CHUNK)]);
  if (!chunks.length) return;
  var need = chunks.length + 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  sh.getRange(2, 1, chunks.length, 1).setValues(chunks);
  var extra = sh.getLastRow() - (chunks.length + 1);
  if (extra > 0) sh.getRange(chunks.length + 2, 1, extra, 1).clearContent();
  sh.getRange(1, 1, 1, 5).setValues([[board.loadedAt, board.loadedAtMs, board.rowCount, chunks.length, board.version]]);
}

function buildBoard_() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  } catch (e) {
    throw new Error('Refurbishment Hemau konnte nicht geöffnet werden. ' + e.message);
  }
  var sh = findTab_(ss, SOURCE_TAB);
  if (!sh) throw new Error('Tabellenblatt nicht gefunden: ' + SOURCE_TAB);

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return emptyBoard_();
  }

  var headerRow = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var usedCols = headerRow.length;
  while (usedCols > 1 && !String(headerRow[usedCols - 1] || '').trim()) usedCols--;
  var headers = headerRow.slice(0, usedCols);
  var numRows = lastRow - 1;
  var values = sh.getRange(2, 1, numRows, usedCols).getDisplayValues();
  var rich = null;
  try {
    rich = sh.getRange(2, 1, numRows, usedCols).getRichTextValues();
  } catch (e2) {
    rich = null;
  }

  var cols = findColumns_(headers);
  var skip = {};
  skip[cols.stock] = true;
  skip[cols.ersatzteil] = true;
  skip[cols.status] = true;
  skip[cols.kommentar] = true;
  if (cols.menge >= 0) skip[cols.menge] = true;

  var extras = [];
  for (var c = 0; c < headers.length; c++) {
    if (skip[c]) continue;
    extras.push({
      index: c,
      label: String(headers[c] || '').trim() || ('Spalte ' + colLetter_(c))
    });
  }

  var order = [];
  var byId = {};
  var dataRows = 0;

  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    var stockId = normalizeStock_(cellAt_(row, cols.stock));
    if (!stockId) continue;
    var ersatzteil = cellAt_(row, cols.ersatzteil);
    var status = cellAt_(row, cols.status);
    var kommentar = cellAt_(row, cols.kommentar);
    var extraCells = [];
    for (var e = 0; e < extras.length; e++) {
      var raw = cleanCell_(row[extras[e].index]);
      var href = '';
      if (rich && rich[r] && rich[r][extras[e].index]) href = linkFromRich_(rich[r][extras[e].index]);
      if (!href && /^https?:\/\//i.test(raw)) href = raw;
      extraCells.push({ label: extras[e].label, value: raw, href: href });
    }
    if (!ersatzteil && !status && !kommentar) {
      if (!byId[stockId]) {
        byId[stockId] = { stockId: stockId, parts: [], needsNothing: true };
        order.push(stockId);
      }
      dataRows++;
      continue;
    }

    dataRows++;
    if (!byId[stockId]) {
      byId[stockId] = { stockId: stockId, parts: [], needsNothing: false };
      order.push(stockId);
    }
    var kept = [];
    for (var k = 0; k < extraCells.length; k++) {
      if (extraCells[k].value || extraCells[k].href) kept.push(extraCells[k]);
    }
    byId[stockId].parts.push({
      ersatzteil: ersatzteil,
      status: status,
      kommentar: kommentar,
      category: categoryFromKommentar_(kommentar),
      extras: kept
    });
  }

  var stocks = [];
  for (var i = 0; i < order.length; i++) {
    var stock = byId[order[i]];
    stock.needsNothing = stock.parts.length === 0;
    stocks.push(stock);
  }

  return stampBoard_({
    sourceTab: SOURCE_TAB,
    rowCount: dataRows,
    stocks: stocks
  });
}

function emptyBoard_() {
  return stampBoard_({
    sourceTab: SOURCE_TAB,
    rowCount: 0,
    stocks: []
  });
}

function stampBoard_(board) {
  var now = new Date();
  board.version = KVA_VERSION;
  board.loadedAt = Utilities.formatDate(now, 'Europe/Berlin', 'dd.MM.yyyy HH:mm');
  board.loadedClock = Utilities.formatDate(now, 'Europe/Berlin', 'HH:mm');
  board.loadedAtMs = now.getTime();
  return board;
}

function findTab_(ss, name) {
  var exact = ss.getSheetByName(name);
  if (exact) return exact;
  var want = String(name || '').toLowerCase();
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    if (String(all[i].getName() || '').toLowerCase() === want) return all[i];
  }
  return null;
}

function findColumns_(headers) {
  var map = { stock: -1, ersatzteil: -1, status: -1, menge: -1, kommentar: -1 };
  for (var i = 0; i < headers.length; i++) {
    var s = String(headers[i] || '').toLowerCase().replace(/\./g, '').replace(/\s+/g, ' ').trim();
    if (map.stock < 0 && (s === 'stock-nr' || s === 'stock nr' || s === 'stockid' || s === 'stock id' || s === 'stock')) map.stock = i;
    else if (map.ersatzteil < 0 && s === 'ersatzteil') map.ersatzteil = i;
    else if (map.status < 0 && s === 'status') map.status = i;
    else if (map.menge < 0 && s === 'menge') map.menge = i;
    else if (map.kommentar < 0 && (s === 'kommentar' || s === 'comment')) map.kommentar = i;
  }
  if (map.stock < 0) {
    map.stock = 0;
    map.ersatzteil = 1;
    map.status = 2;
    map.menge = 3;
    map.kommentar = Math.min(4, Math.max(0, headers.length - 1));
  }
  return map;
}

function categoryFromKommentar_(text) {
  var head = String(text || '').split(':')[0].trim().toLowerCase();
  if (!head) return '';
  if (head.indexOf('vehicle services') === 0) return 'Vehicle services';
  if (head.indexOf('maintenance') === 0) return 'Maintenance';
  if (head.indexOf('damage') === 0) return 'Damage';
  return 'Sonstig';
}

function normalizeStock_(value) {
  return String(value || '').replace(/\u00a0/g, ' ').replace(/\s+/g, '').trim().toUpperCase();
}

function cellAt_(row, index) {
  if (!row || index < 0 || index >= row.length) return '';
  return cleanCell_(row[index]);
}

function cleanCell_(value) {
  return String(value || '').replace(/\u00a0/g, ' ').replace(/[ \t]+\n/g, '\n').trim();
}

function linkFromRich_(cell) {
  if (!cell) return '';
  var direct = '';
  try { direct = cell.getLinkUrl() || ''; } catch (e) { direct = ''; }
  if (direct) return direct;
  var runs = [];
  try { runs = cell.getRuns() || []; } catch (e2) { runs = []; }
  for (var i = 0; i < runs.length; i++) {
    var url = '';
    try { url = runs[i].getLinkUrl() || ''; } catch (e3) { url = ''; }
    if (url) return url;
  }
  return '';
}

function colLetter_(index) {
  var n = index + 1;
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
