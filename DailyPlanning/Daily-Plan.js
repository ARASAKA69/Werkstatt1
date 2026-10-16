var DPL_SOURCE_FILE_ID = '13Oh7gDT8NAul2s0cwQUeaGwMcS3B2MYu0QOdFNMhXzM';

var DPL_SOURCE_SHEET = 'Daily Planning List';

var DPL_TAGESLISTE_FILE_ID = '1PuCLw8UmDjB_pBo_jCZ9rmSD3GJQESHzPoBVu_--MRo';

var DPL_TAGESLISTE_SHEET = 'Tagesliste';

var DPL_CACHE_TTL = 21600;

var DPL_REFRESH_MINUTES = 5;

var DPL_MAP_CHUNK_KEY = 'DPL_MAP_';

var DPL_MAP_COUNT_KEY = 'DPL_MAP_COUNT';

var DPL_LAST_REFRESH_KEY = 'DPL_LAST_REFRESH';

var DPL_MAX_CHUNK = 95000;

var DPL_NOINFO_COLOR = '#FF9900';

var DPL_CAROL_GREEN_COLOR = '#00ff00';

var DPL_NACHBESTELL_BLUE_COLOR = '#3d85c6';

var DPL_QUEUE_CACHE_KEY = 'DPL_CAROL_QUEUE_V2';

var DPL_BRIDGE_STATUS_KEY = 'DPL_CAROL_BRIDGE_STATUS_V2';

var DPL_CAROL_SEARCH_BASE = 'https://carol.autohero.com/en-GB/refurbishment?rsv=';

var DPL_COLS = [2, 8, 9, 12, 13, 14, 15, 16];

function dplNormalizeStockId(v) {

  return String(v == null ? '' : v).replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF]+/g, '').toUpperCase();

}

function dplLooksLikeStockId(v) {

  return /^[A-Z]{2}\d{4,8}$/.test(dplNormalizeStockId(v));

}

function dplFormatDateCell_(val) {

  if (val === null || val === undefined || val === '') return '';

  if (val instanceof Date || Object.prototype.toString.call(val) === '[object Date]') {

    if (isNaN(val.getTime())) return '';

    return Utilities.formatDate(val, 'Europe/Berlin', 'dd.MM.yyyy');

  }

  var raw = String(val || '').trim();

  var m = raw.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);

  if (m) {

    var dd = ('0' + parseInt(m[1], 10)).slice(-2);

    var mm = ('0' + parseInt(m[2], 10)).slice(-2);

    var yyyy = m[3].length === 2 ? ('20' + m[3]) : m[3];

    return dd + '.' + mm + '.' + yyyy;

  }

  var iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);

  if (iso) {

    return iso[3] + '.' + iso[2] + '.' + iso[1];

  }

  return '';

}

function dplParseDateMs_(dateStr) {

  var s = String(dateStr || '').trim();

  var m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);

  if (!m) return 0;

  var d = new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10), 0, 0, 0, 0);

  return isNaN(d.getTime()) ? 0 : d.getTime();

}

function dplTodayMidnightMs_() {

  var nowStr = Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy');

  return dplParseDateMs_(nowStr);

}

function dplGetTageslisteSpreadsheet_() {

  try {

    var active = SpreadsheetApp.getActiveSpreadsheet();

    if (active) return active;

  } catch (e) {}

  return SpreadsheetApp.openById(DPL_TAGESLISTE_FILE_ID);

}

function dplGetTageslisteSheet_(ss) {

  var targetSs = ss || dplGetTageslisteSpreadsheet_();

  if (!targetSs) return null;

  return targetSs.getSheetByName(DPL_TAGESLISTE_SHEET);

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

  var isNbX = String(sh.getRange(row, 2).getValue() || '').trim().toLowerCase() === 'x';

  var defaultColFBg = isNbX ? DPL_NACHBESTELL_BLUE_COLOR : null;

  if (vals && vals.length) {

    var pad = vals.slice(0, 8);

    while (pad.length < 8) pad.push('');

    sh.getRange(row, 6, 1, 8).setValues([pad]);

    sh.getRange(row, 5).setBackground(null);

    sh.getRange(row, 6).setBackground(defaultColFBg);

  } else {

    sh.getRange(row, 6, 1, 8).clearContent();

    sh.getRange(row, 5).setBackground(DPL_NOINFO_COLOR);

    sh.getRange(row, 6).setBackground(defaultColFBg);

  }

}

function dplProcessSingleRowCore(ss, sheetName, row) {

  var sh = ss.getSheetByName(sheetName);

  if (!sh || row < 2) return;

  var input = String(sh.getRange(row, 5).getValue() || '').trim();

  if (!input) {

    sh.getRange(row, 6, 1, 8).clearContent();

    sh.getRange(row, 5).setBackground(null);

    sh.getRange(row, 6).setBackground(null);

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

  var fBg = [];

  for (i = 0; i < rows.length; i++) {

    var vals = map[rows[i].norm];

    var rowBVal = String(acVals[rows[i].row - blockTop][1] || '').trim().toLowerCase();

    fBg.push([rowBVal === 'x' ? DPL_NACHBESTELL_BLUE_COLOR : null]);

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

  sh.getRange(startRow, 6, fBg.length, 1).setBackgrounds(fBg);

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

  if (name !== DPL_TAGESLISTE_SHEET) return;

  var row = ss.getActiveRange().getRow();

  if (row < 2) return;

  dplProcessDownward(ss, name, row);

}

function dplOpenCarolHud() {

  var t = HtmlService.createTemplateFromFile('Tagesliste_HUD');

  t.viewMode = 'dialog';

  t.webAppUrl = dplGetWebAppUrl_();

  var html = t.evaluate()
    .setTitle('Tagesliste · Carol Herausgegeben HUD')
    .setWidth(940)
    .setHeight(710);

  SpreadsheetApp.getUi().showModelessDialog(html, 'Tagesliste · Carol Herausgegeben HUD');

}

function dplOpenCarolSidebar() {

  var t = HtmlService.createTemplateFromFile('Tagesliste_HUD');

  t.viewMode = 'sidebar';

  t.webAppUrl = dplGetWebAppUrl_();

  var html = t.evaluate()
    .setTitle('Carol Herausgegeben HUD');

  SpreadsheetApp.getUi().showSidebar(html);

}

function dplGetWebAppUrl_() {

  try {

    return String(ScriptApp.getService().getUrl() || '');

  } catch (e) {

    return '';

  }

}

function doGet(e) {

  e = e || {};

  var p = e.parameter || {};

  var action = String(p.action || p.page || '').toLowerCase();

  if (action === 'ping') {

    return dplJsonOut_({ success: true, ts: Date.now() });

  }

  if (action === 'queue') {

    return dplJsonOut_(dplGetQueueState());

  }

  if (action === 'markresult' || action === 'markgreen') {

    return dplJsonOut_(dplMarkCarolResult(
      p.stockId || p.sid || '',
      p.status || 'ok',
      p.message || '',
      p.rows ? String(p.rows).split(',') : [],
      p.dateStr || p.date || ''
    ));

  }

  var t = HtmlService.createTemplateFromFile('Tagesliste_HUD');

  t.viewMode = 'webapp';

  t.webAppUrl = dplGetWebAppUrl_();

  return t.evaluate()
    .setTitle('Tagesliste · Carol Herausgegeben HUD')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

}

function doPost(e) {

  try {

    var body = {};

    try {

      body = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    } catch (err) {

      body = {};

    }

    var action = String(body.action || (e && e.parameter && e.parameter.action) || '').toLowerCase();

    if (action === 'queue') {

      if (body.bridgeHeartbeat) {

        try {

          CacheService.getScriptCache().put(DPL_BRIDGE_STATUS_KEY, JSON.stringify({
            busy: !!body.busy,
            currentSid: String(body.currentSid || ''),
            stepText: String(body.stepText || ''),
            ts: Date.now()
          }), 60);

        } catch (eHb) {}

      }

      return dplJsonOut_(dplGetQueueState());

    }

    if (action === 'heartbeat') {

      try {

        CacheService.getScriptCache().put(DPL_BRIDGE_STATUS_KEY, JSON.stringify({
          busy: !!body.busy,
          currentSid: String(body.currentSid || ''),
          stepText: String(body.stepText || ''),
          ts: Date.now()
        }), 60);

      } catch (eHb2) {}

      return dplJsonOut_(dplGetQueueState());

    }

    if (action === 'preparebatch') {

      return dplJsonOut_(dplPrepareCarolBatch(body.stockIds || body.raw || ''));

    }

    if (action === 'clearqueue' || action === 'stopall') {

      return dplJsonOut_(dplClearCarolQueue());

    }

    if (action === 'markresult' || action === 'markgreen') {

      return dplJsonOut_(dplMarkCarolResult(
        body.stockId || body.sid || '',
        body.status || 'ok',
        body.message || '',
        body.rows || [],
        body.targetDate || body.dateStr || ''
      ));

    }

    return dplJsonOut_({ success: false, message: 'Unbekannte Aktion' });

  } catch (err2) {

    return dplJsonOut_({ success: false, message: String(err2.message || err2) });

  }

}

function dplJsonOut_(obj) {

  return ContentService.createTextOutput(JSON.stringify(obj || {}))
    .setMimeType(ContentService.MimeType.JSON);

}

function dplIsGreenHex_(hex) {

  var h = String(hex || '').trim().toLowerCase();

  return h === '#00ff00' || h === '#34a853' || h === '#56d364' || h === '#b6fcb6' || h === '#00ff66' || h === '#93c47d';

}

function dplEnsureColFUnblockedFromConditionalFormat_(sh) {

  if (!sh) return false;

  try {

    var rules = sh.getConditionalFormatRules();

    if (!rules || !rules.length) return false;

    var changed = false;

    var newRules = [];

    for (var i = 0; i < rules.length; i++) {

      var rule = rules[i];

      var ranges = rule.getRanges();

      var newRanges = [];

      var ruleChanged = false;

      for (var r = 0; r < ranges.length; r++) {

        var rng = ranges[r];

        var startCol = rng.getColumn();

        var endCol = rng.getLastColumn();

        var startRow = rng.getRow();

        var numRows = rng.getNumRows();

        if (startCol <= 6 && endCol >= 6 && !(startCol === 6 && endCol === 6)) {

          ruleChanged = true;

          changed = true;

          if (startCol <= 5) {

            newRanges.push(sh.getRange(startRow, startCol, numRows, 5 - startCol + 1));

          }

          if (endCol >= 7) {

            newRanges.push(sh.getRange(startRow, 7, numRows, endCol - 7 + 1));

          }

        } else {

          newRanges.push(rng);

        }

      }

      if (ruleChanged) {

        if (newRanges.length > 0) {

          newRules.push(rule.copy().setRanges(newRanges).build());

        }

      } else {

        newRules.push(rule);

      }

    }

    if (changed) {

      sh.setConditionalFormatRules(newRules);

    }

    return changed;

  } catch (err) {

    console.error('dplEnsureColFUnblockedFromConditionalFormat_: ' + err);

    return false;

  }

}

function dplParseStockIdList_(rawInput) {

  var list = [];

  var seen = {};

  if (Array.isArray(rawInput)) {

    for (var i = 0; i < rawInput.length; i++) {

      var s = dplNormalizeStockId(rawInput[i]);

      if (dplLooksLikeStockId(s) && !seen[s]) {

        seen[s] = true;

        list.push(s);

      }

    }

    return list;

  }

  var text = String(rawInput == null ? '' : rawInput).toUpperCase();

  var matches = text.match(/\b[A-Z]{2}\d{4,8}\b/g) || [];

  for (var m = 0; m < matches.length; m++) {

    var sid = dplNormalizeStockId(matches[m]);

    if (dplLooksLikeStockId(sid) && !seen[sid]) {

      seen[sid] = true;

      list.push(sid);

    }

  }

  return list;

}

function dplExtractUrlFromCellParts_(val, formula, richUrl) {

  if (richUrl && /^https?:\/\//i.test(String(richUrl).trim())) {

    return String(richUrl).trim();

  }

  var f = String(formula || '').trim();

  if (f) {

    var m = f.match(/HYPERLINK\s*\(\s*"([^"]+)"/i) || f.match(/HYPERLINK\s*\(\s*'([^']+)'/i);

    if (m && m[1]) return String(m[1]).trim();

  }

  var v = String(val || '').trim();

  if (/^https?:\/\//i.test(v)) return v;

  return '';

}

function dplFindLastStockRow_(sh) {

  var sheetLast = sh.getLastRow();

  if (sheetLast < 2) return 0;

  var colE = sh.getRange(2, 5, sheetLast - 1, 1).getValues();

  for (var i = colE.length - 1; i >= 0; i--) {

    var v = colE[i][0];

    if (v !== '' && v !== null && dplLooksLikeStockId(v)) {

      return i + 2;

    }

  }

  return 0;

}

function dplScanRecentTageslisteRows_(sh, windowRows) {

  dplEnsureColFUnblockedFromConditionalFormat_(sh);

  var lastStockRow = dplFindLastStockRow_(sh);

  if (lastStockRow < 2) {

    return { rows: [], activeDates: [], todayDateStr: '' };

  }

  var count = Math.min(lastStockRow - 1, windowRows || 800);

  var startRow = Math.max(2, lastStockRow - count + 1);

  var numRows = lastStockRow - startRow + 1;

  var values = sh.getRange(startRow, 1, numRows, 7).getValues();

  var formulasF = sh.getRange(startRow, 6, numRows, 1).getFormulas();

  var bgsF = sh.getRange(startRow, 6, numRows, 1).getBackgrounds();

  var carryDate = '';

  var parsedRows = [];

  var datesOrder = [];

  var dateSeen = {};

  var blueRestoreRows = [];

  for (var i = 0; i < values.length; i++) {

    var rowNum = startRow + i;

    var rawDateStr = dplFormatDateCell_(values[i][0]);

    if (rawDateStr) {

      carryDate = rawDateStr;

    }

    var isNbRow = String(values[i][1] || '').trim().toLowerCase() === 'x';

    var bgCell = bgsF[i] ? String(bgsF[i][0] || '') : '';

    var isGreen = dplIsGreenHex_(bgCell);

    if (isNbRow && !isGreen && String(bgCell).toLowerCase() !== DPL_NACHBESTELL_BLUE_COLOR) {

      blueRestoreRows.push(rowNum);

      bgCell = DPL_NACHBESTELL_BLUE_COLOR;

    }

    var sid = dplNormalizeStockId(values[i][4]);

    if (!sid || !dplLooksLikeStockId(sid)) continue;

    var rowDate = rawDateStr || carryDate || '';

    var rowDateMs = dplParseDateMs_(rowDate);

    if (rowDate && !dateSeen[rowDate]) {

      dateSeen[rowDate] = true;

      datesOrder.push(rowDate);

    }

    var urlCell = dplExtractUrlFromCellParts_(values[i][5], formulasF[i] ? formulasF[i][0] : '', '');

    var modelCell = String(values[i][6] || '').trim();

    parsedRows.push({
      row: rowNum,
      dateStr: rowDate,
      dateMs: rowDateMs,
      isNb: isNbRow,
      schicht: String(values[i][2] || '').trim(),
      stockId: sid,
      carolUrl: urlCell,
      model: modelCell,
      isGreen: isGreen
    });

  }

  if (blueRestoreRows.length > 0 && blueRestoreRows.length <= 120) {

    for (var br = 0; br < blueRestoreRows.length; br++) {

      sh.getRange(blueRestoreRows[br], 6).setBackground(DPL_NACHBESTELL_BLUE_COLOR);

    }

  }

  var todayMs = dplTodayMidnightMs_();

  var activeDateSet = {};

  var activeDates = [];

  for (var d = 0; d < datesOrder.length; d++) {

    var ds = datesOrder[d];

    var dms = dplParseDateMs_(ds);

    if (dms >= todayMs) {

      activeDateSet[ds] = true;

      activeDates.push(ds);

    }

  }

  if (activeDates.length === 0 && datesOrder.length > 0) {

    var tail = datesOrder.slice(-2);

    for (var t = 0; t < tail.length; t++) {

      activeDateSet[tail[t]] = true;

      activeDates.push(tail[t]);

    }

  }

  for (var r = 0; r < parsedRows.length; r++) {

    var pr = parsedRows[r];

    pr.isCurrentOrNextDay = !!(pr.dateStr && activeDateSet[pr.dateStr]);

  }

  return {
    rows: parsedRows,
    activeDates: activeDates,
    allRecentDates: datesOrder.slice(-6),
    todayDateStr: Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy')
  };

}

function dplReadQueueStore_() {

  try {

    var raw = CacheService.getScriptCache().get(DPL_QUEUE_CACHE_KEY);

    if (raw) return JSON.parse(raw);

  } catch (e1) {}

  try {

    var rawProp = PropertiesService.getScriptProperties().getProperty(DPL_QUEUE_CACHE_KEY);

    if (rawProp) return JSON.parse(rawProp);

  } catch (e2) {}

  return { batchId: '', activeDates: [], items: [] };

}

function dplWriteQueueStore_(qObj) {

  var str = JSON.stringify(qObj || { batchId: '', activeDates: [], items: [] });

  try {

    CacheService.getScriptCache().put(DPL_QUEUE_CACHE_KEY, str, 21600);

  } catch (e1) {}

  try {

    PropertiesService.getScriptProperties().setProperty(DPL_QUEUE_CACHE_KEY, str);

  } catch (e2) {}

}

function dplClearCarolQueue() {

  var empty = { batchId: '', activeDates: [], items: [], updatedAt: Date.now() };

  dplWriteQueueStore_(empty);

  return { success: true, items: [] };

}

function dplPrepareCarolBatch(rawInput) {

  var lock = LockService.getScriptLock();

  var gotLock = lock.tryLock(25000);

  try {

    var stockIds = dplParseStockIdList_(rawInput);

    if (!stockIds.length) {

      return { success: false, message: 'Keine gültigen Stock-IDs erkannt.', items: [] };

    }

    var ss = dplGetTageslisteSpreadsheet_();

    var sh = dplGetTageslisteSheet_(ss);

    if (!sh) {

      return { success: false, message: 'Tab "Tagesliste" nicht gefunden.', items: [] };

    }

    var map = dplGetMap();

    var scan = dplScanRecentTageslisteRows_(sh, 800);

    var parsedRows = scan.rows;

    var bySidActive = {};

    for (var i = 0; i < parsedRows.length; i++) {

      var rObj = parsedRows[i];

      if (!rObj.isCurrentOrNextDay) continue;

      var sidCell = rObj.stockId;

      if (!bySidActive[sidCell]) {

        bySidActive[sidCell] = [];

      }

      bySidActive[sidCell].push(rObj);

    }

    var existingStore = dplReadQueueStore_();

    var existingItems = (existingStore && Array.isArray(existingStore.items)) ? existingStore.items : [];

    var bySidMerged = {};

    var mergedList = [];

    for (var e = 0; e < existingItems.length; e++) {

      var ex = existingItems[e];

      if (!ex || !ex.stockId) continue;

      if (ex.status === 'pending' || ex.status === 'running') {

        bySidMerged[ex.stockId] = ex;

        mergedList.push(ex);

      }

    }

    var batchId = 'TL_' + Date.now();

    for (var k = 0; k < stockIds.length; k++) {

      var sid = stockIds[k];

      var candidates = bySidActive[sid] || [];

      var chosenRows = [];

      if (candidates.length > 0) {

        var ungreened = candidates.filter(function(c) { return !c.isGreen; });

        var pool = ungreened.length > 0 ? ungreened : candidates;

        var newestDate = pool[pool.length - 1].dateStr;

        for (var cIdx = 0; cIdx < pool.length; cIdx++) {

          if (pool[cIdx].dateStr === newestDate) {

            chosenRows.push(pool[cIdx]);

          }

        }

      }

      var mapVals = map[sid] || null;

      var mapUrl = (mapVals && mapVals[0] && /^https?:\/\//i.test(String(mapVals[0]).trim())) ? String(mapVals[0]).trim() : '';

      var mapModel = (mapVals && mapVals[1]) ? String(mapVals[1]).trim() : '';

      var rowNumbers = [];

      var sheetCarolUrl = '';

      var sheetModel = '';

      var itemDate = '';

      var allGreen = chosenRows.length > 0;

      for (var cr = 0; cr < chosenRows.length; cr++) {

        var cRow = chosenRows[cr];

        rowNumbers.push(cRow.row);

        if (!itemDate && cRow.dateStr) itemDate = cRow.dateStr;

        if (!sheetCarolUrl && cRow.carolUrl) sheetCarolUrl = cRow.carolUrl;

        if (!sheetModel && cRow.model) sheetModel = cRow.model;

        if (!cRow.isGreen) allGreen = false;

      }

      if (rowNumbers.length > 0 && !sheetCarolUrl && mapVals && mapVals.length) {

        for (var rIdx = 0; rIdx < rowNumbers.length; rIdx++) {

          dplFillRow(sh, rowNumbers[rIdx], mapVals);

        }

        sheetCarolUrl = mapUrl;

        sheetModel = sheetModel || mapModel;

        allGreen = false;

      }

      var finalUrl = sheetCarolUrl || mapUrl || (DPL_CAROL_SEARCH_BASE + encodeURIComponent(sid));

      var finalModel = sheetModel || mapModel || '';

      var newItem = {
        stockId: sid,
        dateStr: itemDate || '',
        carolUrl: finalUrl,
        model: finalModel,
        rows: rowNumbers,
        row: rowNumbers.length ? rowNumbers[rowNumbers.length - 1] : 0,
        inSheet: rowNumbers.length > 0,
        alreadyGreen: allGreen,
        status: 'pending',
        stepText: 'Wartet auf Carol…'
      };

      if (bySidMerged[sid]) {

        bySidMerged[sid].dateStr = newItem.dateStr || bySidMerged[sid].dateStr;

        bySidMerged[sid].carolUrl = newItem.carolUrl || bySidMerged[sid].carolUrl;

        bySidMerged[sid].model = newItem.model || bySidMerged[sid].model;

        bySidMerged[sid].rows = newItem.rows.length ? newItem.rows : bySidMerged[sid].rows;

        bySidMerged[sid].row = newItem.row || bySidMerged[sid].row;

        bySidMerged[sid].status = 'pending';

      } else {

        bySidMerged[sid] = newItem;

        mergedList.push(newItem);

      }

    }

    SpreadsheetApp.flush();

    var queuePayload = {
      batchId: batchId,
      activeDates: scan.activeDates,
      createdAt: Date.now(),
      updatedAt: Date.now(),
      items: mergedList
    };

    dplWriteQueueStore_(queuePayload);

    return {
      success: true,
      batchId: batchId,
      activeDates: scan.activeDates,
      items: mergedList
    };

  } catch (err) {

    return {
      success: false,
      message: String(err.message || err),
      items: []
    };

  } finally {

    if (gotLock) {

      try { lock.releaseLock(); } catch (eRel) {}

    }

  }

}

function dplMarkCarolResult(stockId, status, detailMsg, targetRows, targetDateInput) {

  try {

    var sid = dplNormalizeStockId(stockId);

    if (!sid) return { success: false, message: 'Keine Stock-ID' };

    var isOk = String(status || 'ok').toLowerCase() === 'ok' || String(status || '').toLowerCase() === 'done';

    var wantDate = dplFormatDateCell_(targetDateInput);

    var ss = dplGetTageslisteSpreadsheet_();

    var sh = dplGetTageslisteSheet_(ss);

    if (!sh) return { success: false, message: 'Tab "Tagesliste" nicht gefunden' };

    var matchedRows = {};

    var rowList = [];

    if (Array.isArray(targetRows) && targetRows.length > 0) {

      for (var t = 0; t < targetRows.length; t++) {

        var tr = parseInt(targetRows[t], 10);

        if (tr >= 2 && !matchedRows[tr]) {

          matchedRows[tr] = true;

          rowList.push(tr);

        }

      }

    }

    if (rowList.length === 0) {

      var scan = dplScanRecentTageslisteRows_(sh, 600);

      var candidates = [];

      for (var i = 0; i < scan.rows.length; i++) {

        var rEntry = scan.rows[i];

        if (rEntry.stockId !== sid) continue;

        if (!rEntry.isCurrentOrNextDay) continue;

        if (wantDate && rEntry.dateStr !== wantDate) continue;

        candidates.push(rEntry);

      }

      if (candidates.length > 0) {

        var ungreen = candidates.filter(function(c) { return !c.isGreen; });

        var usePool = ungreen.length > 0 ? ungreen : candidates;

        var newestDay = usePool[usePool.length - 1].dateStr;

        for (var u = 0; u < usePool.length; u++) {

          if (usePool[u].dateStr === newestDay && !matchedRows[usePool[u].row]) {

            matchedRows[usePool[u].row] = true;

            rowList.push(usePool[u].row);

          }

        }

      }

    }

    dplEnsureColFUnblockedFromConditionalFormat_(sh);

    for (var r = 0; r < rowList.length; r++) {

      var rowNumTarget = rowList[r];

      if (isOk) {

        sh.getRange(rowNumTarget, 6).setBackground(DPL_CAROL_GREEN_COLOR);

      } else {

        var hasX = String(sh.getRange(rowNumTarget, 2).getValue() || '').trim().toLowerCase() === 'x';

        sh.getRange(rowNumTarget, 6).setBackground(hasX ? DPL_NACHBESTELL_BLUE_COLOR : null);

      }

    }

    SpreadsheetApp.flush();

    return {
      success: true,
      stockId: sid,
      dateStr: wantDate,
      status: isOk ? 'done' : 'error',
      updatedRows: rowList,
      message: String(detailMsg || '')
    };

  } catch (err) {

    return {
      success: false,
      stockId: stockId,
      message: String(err.message || err)
    };

  }

}

function dplGetRecentUnmarkedFromSheet() {

  try {

    var ss = dplGetTageslisteSpreadsheet_();

    var sh = dplGetTageslisteSheet_(ss);

    if (!sh) return { success: false, message: 'Tab "Tagesliste" nicht gefunden', stockIds: [], items: [], activeDates: [] };

    var scan = dplScanRecentTageslisteRows_(sh, 800);

    var dayItems = [];

    var seen = {};

    var stockIds = [];

    for (var i = 0; i < scan.rows.length; i++) {

      var rObj = scan.rows[i];

      if (!rObj.isCurrentOrNextDay) continue;

      if (rObj.isGreen) continue;

      dayItems.push(rObj);

      if (!seen[rObj.stockId]) {

        seen[rObj.stockId] = true;

        stockIds.push(rObj.stockId);

      }

    }

    return {
      success: true,
      activeDates: scan.activeDates,
      stockIds: stockIds,
      items: dayItems
    };

  } catch (err) {

    return {
      success: false,
      message: String(err.message || err),
      stockIds: [],
      items: [],
      activeDates: []
    };

  }

}

function dplGetQueueState() {

  try {

    var parsed = dplReadQueueStore_();

    var items = (parsed && Array.isArray(parsed.items)) ? parsed.items : [];

    var pendingItems = items.filter(function(it) {
      return it && (it.status === 'pending' || it.status === 'running');
    });

    var bridgeStatus = null;

    try {

      var rawHb = CacheService.getScriptCache().get(DPL_BRIDGE_STATUS_KEY);

      if (rawHb) bridgeStatus = JSON.parse(rawHb);

    } catch (eHb) {}

    return {
      success: true,
      batchId: (parsed && parsed.batchId) || '',
      activeDates: (parsed && parsed.activeDates) || [],
      createdAt: (parsed && parsed.createdAt) || 0,
      updatedAt: (parsed && parsed.updatedAt) || 0,
      items: items,
      pendingCount: pendingItems.length,
      nextItem: pendingItems.length ? pendingItems[0] : null,
      bridgeStatus: bridgeStatus
    };

  } catch (err) {

    return { success: false, message: String(err.message || err), items: [], pendingCount: 0, nextItem: null };

  }

}

function onOpen() {

  SpreadsheetApp.getUi().createMenu('Daily Planning')
    .addItem('🚀 Carol Herausgegeben HUD öffnen', 'dplOpenCarolHud')
    .addItem('📌 Carol HUD als Seitenleiste', 'dplOpenCarolSidebar')
    .addSeparator()
    .addItem('Ab dieser Zeile nach unten verarbeiten', 'dplProcessDownFromSelection')
    .addToUi();

}

function onEdit(e) {

  if (!e || !e.range) return;

  var sheet = e.range.getSheet();

  if (sheet.getName() !== DPL_TAGESLISTE_SHEET) return;

  var row = e.range.getRow();

  var col = e.range.getColumn();

  var numCols = e.range.getNumColumns ? e.range.getNumColumns() : 1;

  var numRows = e.range.getNumRows ? e.range.getNumRows() : 1;

  if (row >= 2 && col <= 2 && (col + numCols - 1) >= 2) {

    dplEnsureColFUnblockedFromConditionalFormat_(sheet);

    for (var rOff = 0; rOff < numRows; rOff++) {

      var targetR = row + rOff;

      var bVal = String(sheet.getRange(targetR, 2).getValue() || '').trim().toLowerCase();

      var curFBg = String(sheet.getRange(targetR, 6).getBackground() || '').toLowerCase();

      if (bVal === 'x' && !dplIsGreenHex_(curFBg)) {

        sheet.getRange(targetR, 6).setBackground(DPL_NACHBESTELL_BLUE_COLOR);

      } else if (bVal !== 'x' && curFBg === DPL_NACHBESTELL_BLUE_COLOR) {

        sheet.getRange(targetR, 6).setBackground(null);

      }

    }

  }

  if (row < 2 || col > 5 || (col + numCols - 1) < 5) return;

  if (numRows > 1) {

    dplProcessDownward(e.source, DPL_TAGESLISTE_SHEET, row);

    return;

  }

  console.log('onEdit fired for Tagesliste!E' + row);

  dplProcessSingleRow(e.source, DPL_TAGESLISTE_SHEET, row);

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
