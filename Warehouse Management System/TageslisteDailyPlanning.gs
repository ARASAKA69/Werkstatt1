var DPL_SOURCE_FILE_ID = '13Oh7gDT8NAul2s0cwQUeaGwMcS3B2MYu0QOdFNMhXzM';
var DPL_SOURCE_SHEET = 'Daily Planning List';
var DPL_CACHE_TTL = 18000;
var DPL_STOCK_PREFIX = 'DPL_S_';
var DPL_COLS = [2, 8, 9, 12, 13, 14, 15, 16];

function dplNormalizeStockId(v) {
  return String(v || '').replace(/\s+/g, '').toUpperCase();
}

function dplFlushCacheEntries(cache, batchEntries, ttl) {
  var keys = Object.keys(batchEntries);
  var idx = 0;
  var MAX_CHUNK = 85000;
  while (idx < keys.length) {
    var slice = {};
    var size = 0;
    while (idx < keys.length) {
      var k = keys[idx];
      var val = batchEntries[k];
      var cost = k.length + val.length + 2;
      if (size > 0 && size + cost > MAX_CHUNK) break;
      slice[k] = val;
      size += cost;
      idx++;
    }
    if (Object.keys(slice).length > 0) {
      cache.putAll(slice, ttl);
    }
  }
}

function dplLoadAllSourceEntries() {
  var sourceSs = SpreadsheetApp.openById(DPL_SOURCE_FILE_ID);
  var sourceSh = sourceSs.getSheetByName(DPL_SOURCE_SHEET);
  var sourceData = sourceSh.getDataRange().getValues();
  var strMap = {};
  var byNorm = {};
  var i;
  for (i = 1; i < sourceData.length; i++) {
    var keyNorm = dplNormalizeStockId(sourceData[i][1]);
    if (!keyNorm) continue;
    var rowValues = DPL_COLS.map(function(c) {
      return sourceData[i][c];
    });
    strMap[DPL_STOCK_PREFIX + keyNorm] = JSON.stringify(rowValues);
    byNorm[keyNorm] = rowValues;
  }
  return { strMap: strMap, byNorm: byNorm };
}

function dplProcessDownwardCore(ss, sheetName, startRow) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh || startRow < 2) return;
  var lastRow = sh.getLastRow();
  if (lastRow < startRow) return;
  var numRows = lastRow - startRow + 1;
  var colE = sh.getRange(startRow, 5, numRows, 1).getValues();
  var meta = [];
  var i;
  for (i = 0; i < colE.length; i++) {
    var inputRaw = colE[i][0] != null ? String(colE[i][0]).trim() : '';
    if (!inputRaw) break;
    meta.push({ sh: sh, row: startRow + i, norm: dplNormalizeStockId(inputRaw) });
  }
  if (!meta.length) return;
  var cache = CacheService.getScriptCache();
  var needFull = false;
  for (i = 0; i < meta.length; i++) {
    if (!cache.get(DPL_STOCK_PREFIX + meta[i].norm)) {
      needFull = true;
      break;
    }
  }
  var pack = null;
  if (needFull) {
    pack = dplLoadAllSourceEntries();
    dplFlushCacheEntries(cache, pack.strMap, DPL_CACHE_TTL);
  }
  for (i = 0; i < meta.length; i++) {
    var m = meta[i];
    var vals = null;
    var ck = cache.get(DPL_STOCK_PREFIX + m.norm);
    if (ck) {
      vals = JSON.parse(ck);
    } else if (pack && pack.byNorm[m.norm]) {
      vals = pack.byNorm[m.norm];
    }
    if (vals && vals.length) {
      var pad = vals.slice(0, 8);
      while (pad.length < 8) pad.push('');
      m.sh.getRange(m.row, 6, 1, 8).setValues([pad]);
    } else {
      m.sh.getRange(m.row, 6, 1, 8).clearContent();
    }
  }
}

function dplProcessDownward(ss, sheetName, startRow) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(280000)) return;
  try {
    dplProcessDownwardCore(ss, sheetName, startRow);
  } catch (err) {
    Logger.log(String(err));
  } finally {
    lock.releaseLock();
  }
}

function dplProcessDownFromSelection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var name = sh.getName();
  if (name !== 'Tagesliste' && name !== 'Nachbestellungen') return;
  var row = ss.getActiveRange().getRow();
  if (row < 2) return;
  dplProcessDownward(ss, name, row);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Daily Planning')
    .addItem('Ab dieser Zeile nach unten verarbeiten', 'dplProcessDownFromSelection')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  var row = e.range.getRow();
  var col = e.range.getColumn();
  if ((sheetName !== 'Tagesliste' && sheetName !== 'Nachbestellungen') || col !== 5 || row < 2) return;
  var inputValue = e.value != null ? String(e.value).trim() : '';
  if (!inputValue) {
    sheet.getRange(row, 6, 1, 8).clearContent();
    return;
  }
  dplProcessDownward(e.source, sheetName, row);
}
