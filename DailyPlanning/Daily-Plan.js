var DPL_SOURCE_FILE_ID = '13Oh7gDT8NAul2s0cwQUeaGwMcS3B2MYu0QOdFNMhXzM';

var DPL_SOURCE_SHEET = 'Daily Planning List';

var DPL_CACHE_TTL = 21600;

var DPL_REFRESH_MINUTES = 5;

var DPL_MAP_CHUNK_KEY = 'DPL_MAP_';

var DPL_MAP_COUNT_KEY = 'DPL_MAP_COUNT';

var DPL_LAST_REFRESH_KEY = 'DPL_LAST_REFRESH';

var DPL_MAX_CHUNK = 95000;

var DPL_NOINFO_COLOR = '#FF9900';

var DPL_COLS = [2, 8, 9, 12, 13, 14, 15, 16];

function dplNormalizeStockId(v) {

  return String(v == null ? '' : v).replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]+/g, '').toUpperCase();

}

function dplBuildMap() {

  var sourceSs = SpreadsheetApp.openById(DPL_SOURCE_FILE_ID);

  var sourceSh = sourceSs.getSheetByName(DPL_SOURCE_SHEET);

  var sourceData = sourceSh.getDataRange().getValues();

  var byNorm = {};

  var i;

  for (i = 0; i < sourceData.length; i++) {

    var keyNorm = dplNormalizeStockId(sourceData[i][1]);

    if (!keyNorm) continue;

    byNorm[keyNorm] = DPL_COLS.map(function(c) {

      return sourceData[i][c];

    });

  }

  return byNorm;

}

function dplStoreMap(byNorm) {

  var cache = CacheService.getScriptCache();

  var json = JSON.stringify(byNorm);

  var toPut = {};

  var count = 0;

  var pos;

  for (pos = 0; pos < json.length; pos += DPL_MAX_CHUNK) {

    toPut[DPL_MAP_CHUNK_KEY + count] = json.substring(pos, pos + DPL_MAX_CHUNK);

    count++;

  }

  if (count === 0) {

    toPut[DPL_MAP_CHUNK_KEY + '0'] = '{}';

    count = 1;

  }

  toPut[DPL_MAP_COUNT_KEY] = String(count);

  toPut[DPL_LAST_REFRESH_KEY] = String(Date.now());

  cache.putAll(toPut, DPL_CACHE_TTL);

}

function dplReadMap() {

  var cache = CacheService.getScriptCache();

  var countStr = cache.get(DPL_MAP_COUNT_KEY);

  if (!countStr) return null;

  var count = parseInt(countStr, 10);

  if (!count || count < 1) return null;

  var keys = [];

  var i;

  for (i = 0; i < count; i++) keys.push(DPL_MAP_CHUNK_KEY + i);

  var got = cache.getAll(keys);

  var json = '';

  for (i = 0; i < count; i++) {

    var part = got[DPL_MAP_CHUNK_KEY + i];

    if (part == null) return null;

    json += part;

  }

  try {

    return JSON.parse(json);

  } catch (err) {

    return null;

  }

}

function dplGetMap() {

  var map = dplReadMap();

  if (map) return map;

  try {

    map = dplBuildMap();

    dplStoreMap(map);

    return map;

  } catch (err) {

    console.error('Cache empty and cannot build from source here (no permission); relying on scheduled refresh: ' + err);

    return {};

  }

}

function dplRefreshCache() {

  var map = dplBuildMap();

  dplStoreMap(map);

  console.log('Daily Planning cache refreshed: ' + Object.keys(map).length + ' entries');

}

function dplFillRow(sh, row, vals) {

  if (vals && vals.length) {

    var pad = vals.slice(0, 8);

    while (pad.length < 8) pad.push('');

    sh.getRange(row, 6, 1, 8).setValues([pad]);

    sh.getRange(row, 5).setBackground(null);

  } else {

    sh.getRange(row, 6, 1, 8).clearContent();

    sh.getRange(row, 5).setBackground(DPL_NOINFO_COLOR);

  }

}

function dplProcessSingleRowCore(ss, sheetName, row) {

  var sh = ss.getSheetByName(sheetName);

  if (!sh || row < 2) return;

  var input = String(sh.getRange(row, 5).getValue() || '').trim();

  if (!input) {

    sh.getRange(row, 6, 1, 8).clearContent();

    sh.getRange(row, 5).setBackground(null);

    return;

  }

  if (row >= 3) {

    var ac = sh.getRange(row - 1, 1, 2, 3).getValues();

    if ((ac[1][0] === '' || ac[1][0] == null) && ac[0][0] !== '' && ac[0][0] != null) {

      sh.getRange(row, 1).setValue(ac[0][0]);

    }

    if ((ac[1][2] === '' || ac[1][2] == null) && ac[0][2] !== '' && ac[0][2] != null) {

      sh.getRange(row, 3).setValue(ac[0][2]);

    }

  }

  var norm = dplNormalizeStockId(input);

  var map = dplGetMap();

  var vals = map[norm];

  console.log('Tagesliste row ' + row + ' stockId "' + norm + '" -> ' + (vals && vals.length ? 'FOUND, filled' : 'NOT FOUND, marked orange'));

  dplFillRow(sh, row, vals);

}

function dplProcessDownwardCore(ss, sheetName, startRow) {

  var sh = ss.getSheetByName(sheetName);

  if (!sh || startRow < 2) return;

  var lastRow = sh.getLastRow();

  if (lastRow < startRow) return;

  var numRows = lastRow - startRow + 1;

  var colE = sh.getRange(startRow, 5, numRows, 1).getValues();

  var rows = [];

  var i;

  for (i = 0; i < colE.length; i++) {

    var inputRaw = colE[i][0] != null ? String(colE[i][0]).trim() : '';

    if (!inputRaw) break;

    rows.push({ row: startRow + i, norm: dplNormalizeStockId(inputRaw) });

  }

  if (!rows.length) return;

  var lastMetaRow = rows[rows.length - 1].row;

  var blockTop = startRow >= 3 ? startRow - 1 : startRow;

  var acVals = sh.getRange(blockTop, 1, lastMetaRow - blockTop + 1, 3).getValues();

  for (i = 0; i < rows.length; i++) {

    var r = rows[i].row;

    if (r < 3) continue;

    var ai = r - blockTop;

    if ((acVals[ai][0] === '' || acVals[ai][0] == null) && acVals[ai - 1][0] !== '' && acVals[ai - 1][0] != null) {

      acVals[ai][0] = acVals[ai - 1][0];

    }

    if ((acVals[ai][2] === '' || acVals[ai][2] == null) && acVals[ai - 1][2] !== '' && acVals[ai - 1][2] != null) {

      acVals[ai][2] = acVals[ai - 1][2];

    }

  }

  var aColumn = [];

  var cColumn = [];

  var rr;

  for (rr = startRow; rr <= lastMetaRow; rr++) {

    aColumn.push([acVals[rr - blockTop][0]]);

    cColumn.push([acVals[rr - blockTop][2]]);

  }

  sh.getRange(startRow, 1, aColumn.length, 1).setValues(aColumn);

  sh.getRange(startRow, 3, cColumn.length, 1).setValues(cColumn);

  var map = dplGetMap();

  var out = [];

  var eBg = [];

  for (i = 0; i < rows.length; i++) {

    var vals = map[rows[i].norm];

    if (vals && vals.length) {

      var pad = vals.slice(0, 8);

      while (pad.length < 8) pad.push('');

      out.push(pad);

      eBg.push([null]);

    } else {

      out.push(['', '', '', '', '', '', '', '']);

      eBg.push([DPL_NOINFO_COLOR]);

    }

  }

  sh.getRange(startRow, 6, out.length, 8).setValues(out);

  sh.getRange(startRow, 5, eBg.length, 1).setBackgrounds(eBg);

}

function dplProcessSingleRow(ss, sheetName, row) {

  try {

    dplProcessSingleRowCore(ss, sheetName, row);

  } catch (err) {

    console.error('dplProcessSingleRow row ' + row + ': ' + err);

  }

}

function dplProcessDownward(ss, sheetName, startRow) {

  var lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) return;

  try {

    dplProcessDownwardCore(ss, sheetName, startRow);

  } catch (err) {

    console.error('dplProcessDownward startRow ' + startRow + ': ' + err);

  } finally {

    lock.releaseLock();

  }

}

function dplProcessDownFromSelection() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sh = ss.getActiveSheet();

  var name = sh.getName();

  if (name !== 'Tagesliste') return;

  var row = ss.getActiveRange().getRow();

  if (row < 2) return;

  dplProcessDownward(ss, name, row);

}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Daily Planning')
    .addItem('Ab dieser Zeile nach unten verarbeiten', 'dplProcessDownFromSelection')
    .addItem('Cache jetzt aktualisieren', 'dplRefreshCache')
    .addItem('Trigger installieren', 'dplInstallTrigger')
    .addToUi();
}

function onEdit(e) {

  if (!e || !e.range) return;

  var sheet = e.range.getSheet();

  if (sheet.getName() !== 'Tagesliste') return;

  var row = e.range.getRow();

  var col = e.range.getColumn();

  if (col !== 5 || row < 2) return;

  console.log('onEdit fired for Tagesliste!E' + row);

  dplProcessSingleRow(e.source, 'Tagesliste', row);

}

function dplOnEdit(e) {

  onEdit(e);

}

function dplInstallTrigger() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var triggers = ScriptApp.getProjectTriggers();

  var i;

  for (i = 0; i < triggers.length; i++) {

    var fn = triggers[i].getHandlerFunction();

    if (fn === 'dplOnEdit' || fn === 'dplRefreshCache' || fn === 'onEdit') {

      ScriptApp.deleteTrigger(triggers[i]);

    }

  }

  ScriptApp.newTrigger('dplRefreshCache')
    .timeBased()
    .everyMinutes(DPL_REFRESH_MINUTES)
    .create();

  dplRefreshCache();

}
