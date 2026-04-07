const TRACKING_SHEET_URL = "https://docs.google.com/spreadsheets/d/1PuCLw8UmDjB_pBo_jCZ9rmSD3GJQESHzPoBVu_--MRo/edit?gid=1453769469#gid=1453769469";
const REIFEN_SHEET_ID = "1NTWkl4r40VUb8hM3Zk5BYWofdxn0FgtZh4DJpOufSd8";
const NACHBESTELL_SHEET_ID = "1PuCLw8UmDjB_pBo_jCZ9rmSD3GJQESHzPoBVu_--MRo";
const NACHBESTELL_TAB = "Nachbestellungen";
const EXIT_SHEET_ID = "1OrSRkB8xdMk0uGvTGUVA_J8Q3IPX0GYF7eOXf6af1GI";
const EXIT_TAB = "Exit Repair";
const AUFTRAG_SHEET_ID = "1nE6SErc1-jmZYd_Ydviw28Pa5qdJmwNepXCiVbsdsVo";
const AUFTRAG_TAB = "BLANCO Reparaturauftrag";
const AUFTRAG_EMAIL = "francesco.berger@auto1.com";
const TAGESLISTE_SHEET_ID = "1PuCLw8UmDjB_pBo_jCZ9rmSD3GJQESHzPoBVu_--MRo";
const TAGESLISTE_TAB = "Tagesliste";
const VASOLD_WSS_TAB = "Vasold WSS";
const NACHBESTELL_REGAL_COL = 13;

function normalizeRegalKeyForCount(val) {
  if (val === "" || val == null) return "";
  if (Object.prototype.toString.call(val) === "[object Date]" || val instanceof Date) {
    var dm = Utilities.formatDate(new Date(val), "Europe/Berlin", "d.M");
    var dp = dm.split(".");
    if (dp.length >= 2) {
      var da = parseInt(dp[0], 10);
      var mo = parseInt(dp[1], 10);
      if (!isNaN(da) && !isNaN(mo) && da >= 1 && da <= 9 && mo >= 1 && mo <= 8) return "Regal " + da + "." + mo;
    }
    return "";
  }
  if (typeof val === "number" && isFinite(val)) {
    var a = Math.floor(val);
    var rest = val - a;
    var b = Math.round(rest * 10 + 1e-6);
    if (a >= 1 && a <= 9 && b >= 1 && b <= 8) return "Regal " + a + "." + b;
    return "";
  }
  var s = String(val || "").trim().replace(/,/g, ".");
  if (!s) return "";
  var low = s.toLowerCase();
  if (low === "tagesliste" || low === "lack" || low === "exit") return "";
  var m = s.match(/^regal\s+(\d+)\s*\.\s*(\d+)$/i);
  if (m) return "Regal " + parseInt(m[1], 10) + "." + parseInt(m[2], 10);
  m = s.match(/^(\d+)\s*\.\s*(\d+)$/);
  if (m) {
    var a2 = parseInt(m[1], 10);
    var b2 = parseInt(m[2], 10);
    if (a2 >= 1 && a2 <= 9 && b2 >= 1 && b2 <= 8) return "Regal " + a2 + "." + b2;
  }
  return "";
}

function nachbestellungRegalUiFromCell(val) {
  var k = normalizeRegalKeyForCount(val);
  if (k) return k.replace(/^Regal /i, "");
  return String(val || "").trim();
}

function nachbestellungLagerortVerifyMatch(expected, cellValue) {
  var e = String(expected || "").trim();
  if (e === "" && (cellValue === "" || cellValue == null)) return true;
  if (cellValue === "" || cellValue == null) return e === "";
  if (typeof cellValue === "string" && e.toLowerCase() === cellValue.trim().toLowerCase()) return true;
  var lowE = e.toLowerCase();
  if ((lowE === "tagesliste" || lowE === "lack" || lowE === "exit") && String(cellValue || "").trim().toLowerCase() === lowE) return true;
  var ne = normalizeRegalKeyForCount(e);
  var nv = normalizeRegalKeyForCount(cellValue);
  if (ne && nv && ne === nv) return true;
  return false;
}

function findNachbestellungLagerortColumn(sheet) {
  var lastRow = Math.max(1, sheet.getLastRow());
  var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
  var headerScan = sheet.getRange(1, 1, Math.min(10, lastRow), lastCol).getValues();
  var h;
  for (h = 0; h < headerScan.length; h++) {
    var colL = getColIndex(headerScan[h], ["lagerort"]);
    if (colL !== -1) return colL;
  }
  for (h = 0; h < headerScan.length; h++) {
    var row = headerScan[h];
    var hasStock = false;
    for (var c = 0; c < row.length; c++) {
      if (String(row[c] || "").toLowerCase().indexOf("stock") !== -1) {
        hasStock = true;
        break;
      }
    }
    if (hasStock) {
      var colR = getColIndex(headerScan[h], ["lagerort", "regal"]);
      if (colR !== -1) return colR;
    }
  }
  return NACHBESTELL_REGAL_COL;
}

function nachbestellungLagerortFallbackList() {
  var o = ["Tagesliste"];
  for (var i = 1; i <= 5; i++) {
    for (var j = 1; j <= 8; j++) {
      o.push(i + "." + j);
    }
  }
  o.push("LACK");
  o.push("Exit");
  return o;
}

function extractDataValidationList(dv) {
  if (!dv) return null;
  if (dv.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return null;
  var cv = dv.getCriteriaValues();
  if (!cv || cv.length === 0) return null;
  var first = cv[0];
  var out = [];
  if (Object.prototype.toString.call(first) === "[object Array]") {
    for (var i = 0; i < first.length; i++) {
      var cellVal = first[i];
      if (cellVal !== "" && cellVal != null) out.push(String(cellVal));
    }
  } else if (first !== "" && first != null) {
    out.push(String(first));
  }
  return out.length ? out : null;
}

function getNachbestellungLagerortAllowedList(sheet, col) {
  var lastRow = Math.max(2, sheet.getLastRow());
  var samples = [2, 3, 4, 5, 10, 15, 20, 50, 100, 500, 1000, 2000, 3000, 4000, 5000];
  for (var s = 0; s < samples.length; s++) {
    var rr = samples[s];
    if (rr > lastRow) continue;
    var list = extractDataValidationList(sheet.getRange(rr, col).getDataValidation());
    if (list && list.length) return list;
  }
  return nachbestellungLagerortFallbackList();
}

function nachbestellungLagerortToSheetValue(raw, allowedList) {
  raw = String(raw || "").trim();
  if (!raw) return "";
  var i;
  for (i = 0; i < allowedList.length; i++) {
    var opt = String(allowedList[i] != null ? allowedList[i] : "").trim();
    if (raw.toLowerCase() === opt.toLowerCase()) return String(allowedList[i]);
  }
  var rawNorm = normalizeRegalKeyForCount(raw);
  for (i = 0; i < allowedList.length; i++) {
    var opt2 = String(allowedList[i] != null ? allowedList[i] : "").trim();
    var optNorm = normalizeRegalKeyForCount(opt2);
    if (rawNorm && optNorm && rawNorm === optNorm) return String(allowedList[i]);
  }
  return null;
}

function mergeNachbestellungenRegalIntoCounts(counts) {
  try {
    var ss = SpreadsheetApp.openById(NACHBESTELL_SHEET_ID);
    var sheet = ss.getSheetByName(NACHBESTELL_TAB);
    if (!sheet || !counts) return;
    var col = findNachbestellungLagerortColumn(sheet);
    var lastRow = Math.max(1, sheet.getLastRow());
    var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
    if (col > lastCol) col = Math.min(NACHBESTELL_REGAL_COL, lastCol);
    var headerScan = sheet.getRange(1, 1, Math.min(10, lastRow), lastCol).getValues();
    var headerIdx = -1;
    for (var h = 0; h < headerScan.length; h++) {
      var scanRow = headerScan[h];
      var hasStock = false;
      for (var c = 0; c < scanRow.length; c++) {
        if (String(scanRow[c] || "").toLowerCase().indexOf("stock") !== -1) {
          hasStock = true;
          break;
        }
      }
      if (hasStock) {
        headerIdx = h;
        break;
      }
    }
    if (headerIdx === -1) headerIdx = 0;
    var startRow = headerIdx + 2;
    var numRows = lastRow - headerIdx - 1;
    if (numRows <= 0) return;
    var regalColData = sheet.getRange(startRow, col, numRows, 1).getValues();
    for (var r = 0; r < regalColData.length; r++) {
      var key = normalizeRegalKeyForCount(regalColData[r][0]);
      if (key && counts.hasOwnProperty(key)) counts[key]++;
    }
  } catch (err) {}
}

function normalizeStockId(value) {
    return String(value || "").replace(/\s+/g, "").toUpperCase();
  }

function cellMatchesStockId(cellVal, stockId) {
    var cv = normalizeStockId(cellVal);
    var sid = normalizeStockId(stockId);
    return sid !== "" && cv === sid;
  }

function getColIndex(headerRow, searchTerms) {
    if (!headerRow) return -1;
    for (var i = 0; i < headerRow.length; i++) {
      var cellText = String(headerRow[i] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
      for (var j = 0; j < searchTerms.length; j++) {
        var term = String(searchTerms[j] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
        if (cellText === term || cellText.indexOf(term) !== -1) return i + 1;
      }
    }
    return -1;
  }

function findHeaderRow(data, searchTerms) {
    for (var i = 0; i < Math.min(30, data.length); i++) {
      if (getColIndex(data[i], searchTerms) !== -1) return i;
    }
    return -1;
  }

function findRowFast(sheet, searchTermsHeader, stockId) {
    stockId = normalizeStockId(stockId);
    var lastRow = Math.max(1, sheet.getLastRow());
    var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
    var headerData = sheet.getRange(1, 1, Math.min(30, lastRow), lastCol).getValues();
    var headerIdx = findHeaderRow(headerData, searchTermsHeader);
    if (headerIdx === -1) return { row: -1, headerIdx: -1, stockCol: -1 };

    var stockCol = getColIndex(headerData[headerIdx], searchTermsHeader);
    if (stockCol === -1) return { row: -1, headerIdx: headerIdx, stockCol: -1 };

    var colData = sheet.getRange(1, stockCol, lastRow, 1).getValues();
    for (var i = headerIdx + 1; i < colData.length; i++) {
      if (cellMatchesStockId(colData[i][0], stockId)) {
        return { row: i + 1, headerIdx: headerIdx, stockCol: stockCol };
      }
    }
    return { row: -1, headerIdx: headerIdx, stockCol: stockCol };
  }

function getReifenSheet() {
    return SpreadsheetApp.openById(REIFEN_SHEET_ID);
  }

function getReifenSheetTab(tabName) {
    var ss = getReifenSheet();
    var name = String(tabName || "").trim();
    var sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
    var sheets = ss.getSheets();
    var normalizedName = name.replace(/\s+/g, '').toLowerCase();
    for (var i = 0; i < sheets.length; i++) {
      var sn = sheets[i].getName().replace(/\s+/g, '').toLowerCase();
      if (sn === normalizedName) return sheets[i];
    }
    return null;
  }

function parseReifenTabDateMsFromName(name) {
    var trimmed = String(name || "").trim();
    var all = trimmed.match(/\d{2}\.\d{2}\.\d{4}/g);
    if (!all || !all.length) return null;
    var dateStr = all[all.length - 1];
    var p = dateStr.split(".");
    if (p.length !== 3) return null;
    var t = new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10)).getTime();
    return isNaN(t) ? null : t;
  }

function filterReifenSheetNamesByTabDateRange(startBound, endBound) {
    var ss = getReifenSheet();
    var sheetNames = ss.getSheets().map(function(sheet) {
      return sheet.getName();
    });
    var filtered = [];
    var i;
    for (i = 0; i < sheetNames.length; i++) {
      var name = sheetNames[i];
      var tabDate = parseReifenTabDateMsFromName(name);
      if (tabDate === null) continue;
      if (tabDate < startBound || tabDate > endBound) continue;
      filtered.push(name);
    }
    filtered.sort(function(a, b) {
      var ta = parseReifenTabDateMsFromName(a);
      var tb = parseReifenTabDateMsFromName(b);
      if (ta === null || tb === null) return 0;
      return tb - ta;
    });
    return filtered;
  }

function getReifenTabNamesForStockHudWindow() {
    var now = new Date();
    var y = now.getFullYear();
    var mo = now.getMonth();
    var d = now.getDate();
    var endBound = new Date(y, mo, d, 23, 59, 59, 999).getTime();
    var startBound = new Date(y, mo, d - 7, 0, 0, 0, 0).getTime();
    return filterReifenSheetNamesByTabDateRange(startBound, endBound);
  }

function getReifenTabNamesForPaketHudWindow() {
    var now = new Date();
    var y = now.getFullYear();
    var mo = now.getMonth();
    var d = now.getDate();
    var endBound = new Date(y, mo, d, 23, 59, 59, 999).getTime();
    var startBound = new Date(y, mo, d - 21, 0, 0, 0, 0).getTime();
    return filterReifenSheetNamesByTabDateRange(startBound, endBound);
  }

function getReifenTabOptions() {
    try {
      return { success: true, tabs: getReifenTabNamesForStockHudWindow() };
    } catch (err) {
      return { success: false, message: err.message, tabs: [] };
    }
  }

function appendPaketdienstRowsFromSheet_(sh, results) {
    var tabName = sh.getName();
    var lastRow = Math.max(1, sh.getLastRow());
    var lastCol = Math.max(1, Math.min(80, sh.getLastColumn()));
    if (lastRow < 3) return;

    var headerData = sh.getRange(1, 1, Math.min(30, lastRow), lastCol).getValues();
    var headerIdx = findHeaderRow(headerData, ["stockid", "stock"]);
    if (headerIdx === -1) return;
    var header = headerData[headerIdx];

    var stockCol = getColIndex(header, ["stockid", "stock"]);
    var groesseCol = getColIndex(header, ["größe", "groesse"]);
    var lastIndexCol = getColIndex(header, ["lastindex", "last"]);
    var gwIndexCol = getColIndex(header, ["gwindex", "gw"]);
    var mengeCol = getColIndex(header, ["menge", "anzahl"]);
    var angeliefertCol = getColIndex(header, ["angeliefert"]);
    if (stockCol === -1 || groesseCol === -1) return;

    var startRow = headerIdx + 2;
    var numRows = lastRow - startRow + 1;
    if (numRows <= 0) return;
    var data = sh.getRange(startRow, 1, numRows, lastCol).getValues();

    var i;
    for (i = 0; i < data.length; i++) {
      var row = data[i];
      var hasPaket = false;
      var c;
      for (c = 0; c < row.length; c++) {
        if (String(row[c] || "").toLowerCase().indexOf("paketdienst") !== -1) { hasPaket = true; break; }
      }
      if (!hasPaket) continue;

      if (angeliefertCol !== -1) {
        var aStatus = String(row[angeliefertCol - 1] || "").trim().toLowerCase();
        if (aStatus === "ja" || aStatus === "nein") continue;
      }

      var cellGroesse = String(row[groesseCol - 1] || "").trim();
      var cellLast = lastIndexCol !== -1 ? String(row[lastIndexCol - 1] || "").trim() : "";
      var cellGw = gwIndexCol !== -1 ? String(row[gwIndexCol - 1] || "").trim() : "";
      var cellStock = stockCol !== -1 ? normalizeStockId(row[stockCol - 1]) : "";
      var cellMenge = mengeCol !== -1 ? (parseInt(row[mengeCol - 1], 10) || 1) : 1;
      if (!cellStock) continue;

      results.push({
        tabName: tabName,
        stockId: cellStock,
        groesse: cellGroesse,
        lastindex: cellLast,
        gwIndex: cellGw,
        menge: cellMenge,
        row: startRow + i
      });
    }
  }

function sortPaketdienstResults_(results) {
    results.sort(function(a, b) {
      var ma = String(a.tabName).match(/(\d{2})\.(\d{2})\.(\d{4})$/);
      var mb = String(b.tabName).match(/(\d{2})\.(\d{2})\.(\d{4})$/);
      if (ma && mb) {
        var da = new Date(parseInt(ma[3], 10), parseInt(ma[2], 10) - 1, parseInt(ma[1], 10)).getTime();
        var db = new Date(parseInt(mb[3], 10), parseInt(mb[2], 10) - 1, parseInt(mb[1], 10)).getTime();
        return db - da;
      }
      return 0;
    });
  }

function collectPaketdienstReifenRowsForTabNames(tabNames) {
    var results = [];
    var si;
    for (si = 0; si < tabNames.length; si++) {
      var sh = getReifenSheetTab(tabNames[si]);
      if (!sh) continue;
      appendPaketdienstRowsFromSheet_(sh, results);
    }
    sortPaketdienstResults_(results);
    return results;
  }

function collectPaketdienstReifenRows() {
    return collectPaketdienstReifenRowsForTabNames(getReifenTabNamesForPaketHudWindow());
  }

function searchReifenBySize(query) {
    try {
      query = String(query || "").trim();
      if (!query || query.length < 3) return { success: false, message: "Suchbegriff zu kurz (min. 3 Zeichen)", results: [] };

      var queryParts = query.replace(/\s+/g, ' ').split(' ');
      var sizeQuery = "";
      var lastQuery = "";
      var gwQuery = "";
      for (var q = 0; q < queryParts.length; q++) {
        var p = queryParts[q];
        if (p.indexOf('/') !== -1 || /^r\d+/i.test(p) || /^\d{3}\//.test(p)) {
          sizeQuery += (sizeQuery ? " " : "") + p;
        } else if (/^\d{2,3}$/.test(p) && !lastQuery) {
          lastQuery = p;
        } else if (/^[A-Za-z]{1,2}$/.test(p) && !gwQuery) {
          gwQuery = p.toUpperCase();
        } else {
          sizeQuery += (sizeQuery ? " " : "") + p;
        }
      }
      if (!sizeQuery) sizeQuery = query;
      var sizeNorm = sizeQuery.replace(/\s+/g, '').toUpperCase();

      var all = collectPaketdienstReifenRows();
      var results = [];
      for (var ri = 0; ri < all.length; ri++) {
        var item = all[ri];
        var cellGroesseNorm = String(item.groesse || "").replace(/\s+/g, '').toUpperCase();
        if (cellGroesseNorm.indexOf(sizeNorm) === -1) continue;
        var cellLast = String(item.lastindex || "").trim();
        var cellGw = String(item.gwIndex || "").trim();
        if (lastQuery && cellLast.indexOf(lastQuery) === -1) continue;
        if (gwQuery && cellGw.toUpperCase().indexOf(gwQuery) === -1) continue;
        results.push(item);
      }

      return { success: true, results: results, message: results.length + " Reifen gefunden" };
    } catch (err) {
      return { success: false, message: "Fehler: " + err.message, results: [] };
    }
  }

function getPaketdienstReifenCachePayload() {
    try {
      var rows = collectPaketdienstReifenRows();
      return { success: true, version: Date.now(), paketdienstRows: rows };
    } catch (err) {
      return { success: false, message: err.message, paketdienstRows: [] };
    }
  }

function getReifenWindowStocksCachePayload() {
    try {
      var tabsRes = getReifenTabOptions();
      if (!tabsRes.success) {
        return { success: false, message: tabsRes.message || "Tabs", tabs: [], stocksByTab: {} };
      }
      var tabs = tabsRes.tabs || [];
      var stocksByTab = {};
      var ti;
      for (ti = 0; ti < tabs.length; ti++) {
        var tn = tabs[ti];
        var idRes = getAvailableReifenStockIds(tn);
        stocksByTab[tn] = (idRes.success && idRes.ids) ? idRes.ids : [];
      }
      return { success: true, version: Date.now(), tabs: tabs, stocksByTab: stocksByTab };
    } catch (err) {
      return { success: false, message: err.message, tabs: [], stocksByTab: {} };
    }
  }

function getAvailableReifenStockIds(tabName) {
    try {
      var sheet = getReifenSheetTab(tabName);
      if (!sheet) return { success: false, message: "Tabellenblatt nicht gefunden!", ids: [] };
      var search = findRowFast(sheet, ["stockid", "stock"], "___NEVER_MATCH___");
      if (search.headerIdx === -1 || search.stockCol === -1) {
        return { success: false, message: "Kopfzeile 'Stock ID' nicht gefunden!", ids: [] };
      }

      var lastRow = Math.max(1, sheet.getLastRow());
      var startRow = search.headerIdx + 2;
      var numRows = lastRow - startRow + 1;
      if (numRows <= 0) return { success: true, ids: [] };

      var headerRow = sheet.getRange(search.headerIdx + 1, 1, 1, Math.max(1, Math.min(80, sheet.getLastColumn()))).getValues()[0];
      var angeliefertCol = getColIndex(headerRow, ["angeliefert"]);
      var stockData = sheet.getRange(startRow, search.stockCol, numRows, 1).getValues();
      var statusData = angeliefertCol !== -1 ? sheet.getRange(startRow, angeliefertCol, numRows, 1).getValues() : [];
      var ids = [];
      for (var i = 0; i < stockData.length; i++) {
        var val = normalizeStockId(stockData[i][0]);
        if (val) {
          ids.push({
            id: val,
            status: angeliefertCol !== -1 ? String(statusData[i][0] || "").trim().toLowerCase() : ""
          });
        }
      }
      return { success: true, ids: ids };
    } catch (err) {
      return { success: false, message: err.message, ids: [] };
    }
  }

function checkReifenStock(tabName, stockId) {
    try {
      stockId = normalizeStockId(stockId);
      if (!stockId) return { found: false, message: "Bitte eine Stock-ID eingeben." };

      var sheet = getReifenSheetTab(tabName);
      if (!sheet) return { found: false, message: "Bitte ein gültiges Tabellenblatt auswählen." };

      var search = findRowFast(sheet, ["stockid", "stock"], stockId);
      if (search.headerIdx === -1) return { found: false, message: "Kopfzeile 'Stock ID' in Reifenliste nicht gefunden!" };
      if (search.row === -1) return { found: false, message: "Stock-ID '" + stockId + "' in '" + sheet.getName() + "' nicht gefunden!" };
      var headerRow = sheet.getRange(search.headerIdx + 1, 1, 1, Math.max(1, Math.min(80, sheet.getLastColumn()))).getValues()[0];
      var angeliefertCol = getColIndex(headerRow, ["angeliefert"]);
      if (angeliefertCol !== -1) {
        var currentStatus = String(sheet.getRange(search.row, angeliefertCol).getValue() || "").trim().toLowerCase();
        if (currentStatus === "ja" || currentStatus === "nein") {
          return { found: false, message: "Stock-ID '" + stockId + "' wurde in '" + sheet.getName() + "' bereits verbucht!" };
        }
      }
      return { found: true, message: "Stock-ID gefunden! Bitte Status auswählen:" };
    } catch (err) {
      return { found: false, message: "Systemfehler: " + err.message };
    }
  }

function processReifenStock(tabName, stockId, isDelivered) {
    try {
      stockId = normalizeStockId(stockId);
      var sheetSeng = getReifenSheetTab(tabName);
      if (!sheetSeng) return { success: false, message: "Bitte ein gültiges Tabellenblatt auswählen." };

      var sheetHemau = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
      if (!sheetHemau) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };

      var tireInfo = "UNBEKANNT _X";
      var mengeValNum = 1;
      var search = findRowFast(sheetSeng, ["stockid", "stock"], stockId);
      if (search.headerIdx === -1) return { success: false, message: "Kopfzeile 'Stock ID' in Reifenliste nicht gefunden!" };
      if (search.row === -1) return { success: false, message: "Stock-ID '" + stockId + "' in '" + sheetSeng.getName() + "' nicht gefunden!" };

      var headerRow = sheetSeng.getRange(search.headerIdx + 1, 1, 1, Math.max(1, Math.min(80, sheetSeng.getLastColumn()))).getValues()[0];
      var angeliefertCol = getColIndex(headerRow, ["angeliefert"]);
      var mengeCol = getColIndex(headerRow, ["menge", "anzahl"]);
      var groesseCol = getColIndex(headerRow, ["größe", "groesse"]);
      var lastIndexCol = getColIndex(headerRow, ["lastindex", "last"]);
      var gwIndexCol = getColIndex(headerRow, ["gwindex", "gw"]);
      if (angeliefertCol !== -1) {
        var existingStatus = String(sheetSeng.getRange(search.row, angeliefertCol).getValue() || "").trim().toLowerCase();
        if (existingStatus === "ja" || existingStatus === "nein") {
          return { success: false, message: "Stock-ID '" + stockId + "' wurde in '" + sheetSeng.getName() + "' bereits verbucht!" };
        }
      }

      var mengeVal = mengeCol !== -1 ? sheetSeng.getRange(search.row, mengeCol).getValue() : "1";
      mengeValNum = parseInt(mengeVal, 10) || 1;
      var groesseVal = groesseCol !== -1 ? sheetSeng.getRange(search.row, groesseCol).getValue() : "GRÖSSE";
      var lastIndexVal = lastIndexCol !== -1 ? sheetSeng.getRange(search.row, lastIndexCol).getValue() : "";
      var gwIndexVal = gwIndexCol !== -1 ? sheetSeng.getRange(search.row, gwIndexCol).getValue() : "";
      tireInfo = String(groesseVal).trim() + " " + String(lastIndexVal).trim() + String(gwIndexVal).trim() + " _" + String(mengeVal).trim();

      if (angeliefertCol !== -1) {
        var statusSeng = isDelivered ? "Ja" : "Nein";
        var colorSeng = isDelivered ? "#00FF00" : "#FF0000";
        sheetSeng.getRange(search.row, angeliefertCol).setValue(statusSeng);
        var startCol = Math.min(search.stockCol, angeliefertCol);
        var numCols = Math.abs(angeliefertCol - search.stockCol) + 1;
        sheetSeng.getRange(search.row, startCol, 1, numCols).setBackground(colorSeng);
      }

      var hemauDataStock = sheetHemau.getRange(1, 2, Math.max(1, sheetHemau.getLastRow()), 1).getValues();
      var hemauRow = -1;
      for (var i = 0; i < hemauDataStock.length; i++) {
        if (cellMatchesStockId(hemauDataStock[i][0], stockId)) {
          hemauRow = i + 1;
          break;
        }
      }

      var hemauMsg = "";
      var locationText = "Lagerplatz unbekannt";
      if (hemauRow !== -1) {
        var oldLocation = String(sheetHemau.getRange(hemauRow, 28).getValue() || "").trim();
        if (oldLocation !== "") locationText = "Kiste steht in Regal " + oldLocation;

        var currentComment = String(sheetHemau.getRange(hemauRow, 25).getValue() || "");
        if (isDelivered) {
          if (currentComment.indexOf("Reifen da //") === -1) {
            var newComment = currentComment ? "Reifen da // " + currentComment : "Reifen da // ";
            sheetHemau.getRange(hemauRow, 25).setValue(newComment);
          }
          sheetHemau.getRange(hemauRow, 30).setValue("Werkstatt 1");
          hemauMsg = "Reifen als da gebucht & Refurbishment aktualisiert!";
        } else {
          sheetHemau.getRange(hemauRow, 30).setValue("Reifen nicht vorhanden");
          hemauMsg = "Reifen als fehlend gebucht & Refurbishment aktualisiert!";
        }
      } else {
        hemauMsg = "Reifen gebucht (Refurbishment übersprungen: ID nicht gefunden)";
      }

      SpreadsheetApp.flush();
      return {
        success: true,
        message: hemauMsg,
        stockId: stockId,
        tireInfo: tireInfo,
        locationText: locationText,
        menge: mengeValNum
      };
    } catch (err) {
      return { success: false, message: "Fehler: " + err.message };
    }
  }

function extractHuDateOnlyFromNachuntersuchungSegment(segment) {
  var s = String(segment || "").trim();
  if (!s) return "";
  var mNum = s.match(/(\d{1,2}\.\s*\d{1,2}\.\s*\d{2,4})/);
  if (mNum) {
    var p = mNum[1].match(/(\d{1,2})\s*\.\s*(\d{1,2})\s*\.\s*(\d{2,4})/);
    if (p) return p[1] + "." + p[2] + "." + p[3];
  }
  var reDotMo = /(\d{1,2}\.\s*(?:Jan|Feb|Mär|Mrz|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez|Januar|Februar|März|April|Juni|Juli|September|Oktober|November|Dezember)[a-zä]*\.?\s*\d{4})/i;
  var mDot = s.match(reDotMo);
  if (mDot) return mDot[1].replace(/\s+/g, " ").trim();
  var reSpMo = /(\d{1,2}\s+(?:Jan|Feb|Mär|Mrz|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez|Januar|Februar|März|April|Juni|Juli|September|Oktober|November|Dezember)[a-zä]*\.?\s*\d{4})/i;
  var mSp = s.match(reSpMo);
  if (mSp) return mSp[1].replace(/\s+/g, " ").trim();
  return "";
}

function huVasoldValueFromSchaedenText(wText) {
  var dash = "---------------";
  var m = String(wText || "").match(/Nachuntersuchung\s*bis\s*:\s*([^\n\r]+)/i);
  if (!m) return dash;
  var rest = String(m[1] || "").trim();
  if (!rest) return dash;
  var dateOnly = extractHuDateOnlyFromNachuntersuchungSegment(rest);
  if (!dateOnly) return dash;
  return dateOnly;
}

function processWssVasoldBooking(stockId, carolUrlOpt, markeOpt, gummiVorhanden) {
  try {
    stockId = normalizeStockId(stockId);
    if (!stockId) return { success: false, message: "Keine Stock-ID" };
    carolUrlOpt = String(carolUrlOpt || "").trim();
    markeOpt = String(markeOpt || "").trim();
    var gummiYes = gummiVorhanden === true || gummiVorhanden === "true";

    var sheetRef = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
    if (!sheetRef) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };

    var lastRef = Math.max(2, sheetRef.getLastRow());
    var stockCol = sheetRef.getRange(1, 2, lastRef, 1).getValues();
    var refurbRow = -1;
    for (var ir = 1; ir < stockCol.length; ir++) {
      if (cellMatchesStockId(stockCol[ir][0], stockId)) {
        refurbRow = ir + 1;
        break;
      }
    }
    if (refurbRow === -1) return { success: false, message: "Stock-ID in Refurbisment List nicht gefunden!" };

    var wText = String(sheetRef.getRange(refurbRow, 23).getValue() || "");
    if (!/wss\s*ers/i.test(wText)) {
      return { success: false, message: "Kein „WSS ers“ in Schaden (Spalte W)." };
    }

    var carol = carolUrlOpt || String(sheetRef.getRange(refurbRow, 3).getValue() || "").trim();
    var marke = markeOpt || String(sheetRef.getRange(refurbRow, 13).getValue() || "").trim();
    if (!carol) return { success: false, message: "Carol-Link fehlt oder konnte nicht gelesen werden — bitte aus Carol kopieren und im Dialog eintragen." };
    if (!marke) return { success: false, message: "Marke & Model fehlt oder konnte nicht gelesen werden — bitte aus Carol kopieren und im Dialog eintragen." };

    var huVasold = huVasoldValueFromSchaedenText(wText);

    var ssV = SpreadsheetApp.openById(TAGESLISTE_SHEET_ID);
    var sheetV = ssV.getSheetByName(VASOLD_WSS_TAB);
    if (!sheetV) return { success: false, message: "Tab „" + VASOLD_WSS_TAB + "“ in Hemau Tageslisten nicht gefunden!" };

    var searchV = findRowFast(sheetV, ["stockid", "stock"], stockId);
    if (searchV.headerIdx === -1) {
      return { success: false, message: "Vasold WSS: Kopfzeile mit Stock-ID nicht gefunden!" };
    }

    var hdr1 = searchV.headerIdx + 1;
    var targetRow;
    var isNew = false;
    if (searchV.row === -1) {
      isNew = true;
      var lrV = sheetV.getLastRow();
      targetRow = Math.max(lrV, hdr1) + 1;
    } else {
      targetRow = searchV.row;
    }

    sheetV.getRange(targetRow, 1).setValue(stockId);
    sheetV.getRange(targetRow, 2).setValue(carol);
    sheetV.getRange(targetRow, 3).setValue(marke);
    sheetV.getRange(targetRow, 4).setValue(huVasold);
    sheetV.getRange(targetRow, 5).setValue("Ja");
    if (gummiYes) sheetV.getRange(targetRow, 6).setValue("Vorhanden");

    var curCom = String(sheetRef.getRange(refurbRow, 25).getValue() || "");
    if (!/wss\s+da\b/i.test(curCom)) {
      sheetRef.getRange(refurbRow, 25).setValue(curCom ? "WSS da // " + curCom : "WSS da // ");
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: (isNew ? "Vasold WSS: neue Zeile. " : "Vasold WSS: Zeile aktualisiert. ") + "Kommentar Anlieferung: WSS da //",
      stockId: stockId,
      isNew: isNew,
      row: targetRow
    };
  } catch (err) {
    return { success: false, message: "Fehler: " + err.message };
  }
}

function getVasoldWssSyncState(stockId) {
  try {
    stockId = normalizeStockId(stockId);
    if (!stockId) return { success: false, rowFound: false, message: "Keine Stock-ID" };

    var ssV = SpreadsheetApp.openById(TAGESLISTE_SHEET_ID);
    var sheetV = ssV.getSheetByName(VASOLD_WSS_TAB);
    if (!sheetV) return { success: false, rowFound: false, message: "Tab fehlt" };

    var searchV = findRowFast(sheetV, ["stockid", "stock"], stockId);
    if (searchV.headerIdx === -1) {
      return { success: true, rowFound: false, frontscheibeJa: false, gummileisteDa: false };
    }
    if (searchV.row === -1) {
      return { success: true, rowFound: false, frontscheibeJa: false, gummileisteDa: false };
    }

    var eRaw = sheetV.getRange(searchV.row, 5).getValue();
    var fRaw = sheetV.getRange(searchV.row, 6).getValue();
    var eVal = String(eRaw != null ? eRaw : "").trim().toLowerCase();
    var fVal = String(fRaw != null ? fRaw : "").trim().toLowerCase();

    var frontscheibeJa = eVal === "ja";
    var gummileisteDa = fVal === "vorhanden" || fVal === "da" || fVal.indexOf("vorhanden") !== -1;

    return {
      success: true,
      rowFound: true,
      frontscheibeJa: frontscheibeJa,
      gummileisteDa: gummileisteDa
    };
  } catch (err) {
    return { success: false, rowFound: false, message: err.message };
  }
}

function processNachbestellVasoldWss(stockId, toggleWssJa, toggleGummi) {
  try {
    stockId = normalizeStockId(stockId);
    if (!stockId) return { success: false, message: "Keine Stock-ID" };
    var wssYes = toggleWssJa === true || toggleWssJa === "true";
    var gummiYes = toggleGummi === true || toggleGummi === "true";
    if (!wssYes && !gummiYes) {
      return { success: false, message: "Keine Option zum Übernehmen gewählt." };
    }

    var sheetRef = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
    if (!sheetRef) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };

    var lastRef = Math.max(2, sheetRef.getLastRow());
    var stockCol = sheetRef.getRange(1, 2, lastRef, 1).getValues();
    var refurbRow = -1;
    for (var ir = 1; ir < stockCol.length; ir++) {
      if (cellMatchesStockId(stockCol[ir][0], stockId)) {
        refurbRow = ir + 1;
        break;
      }
    }
    if (refurbRow === -1) return { success: false, message: "Stock-ID in Refurbisment List nicht gefunden!" };

    var wText = String(sheetRef.getRange(refurbRow, 23).getValue() || "");
    var carol = String(sheetRef.getRange(refurbRow, 3).getValue() || "").trim();
    var marke = String(sheetRef.getRange(refurbRow, 13).getValue() || "").trim();
    if (!carol) return { success: false, message: "Carol-Link fehlt in Refurbishment." };
    if (!marke) return { success: false, message: "Marke & Model fehlt in Refurbishment." };

    var huVasold = huVasoldValueFromSchaedenText(wText);

    var ssV = SpreadsheetApp.openById(TAGESLISTE_SHEET_ID);
    var sheetV = ssV.getSheetByName(VASOLD_WSS_TAB);
    if (!sheetV) return { success: false, message: "Tab „" + VASOLD_WSS_TAB + "“ in Hemau Tageslisten nicht gefunden!" };

    var searchV = findRowFast(sheetV, ["stockid", "stock"], stockId);
    if (searchV.headerIdx === -1) {
      return { success: false, message: "Vasold WSS: Kopfzeile mit Stock-ID nicht gefunden!" };
    }

    var hdr1 = searchV.headerIdx + 1;
    var targetRow;
    var isNew = false;
    if (searchV.row === -1) {
      isNew = true;
      var lrV = sheetV.getLastRow();
      targetRow = Math.max(lrV, hdr1) + 1;
    } else {
      targetRow = searchV.row;
    }

    var eRaw = sheetV.getRange(targetRow, 5).getValue();
    var fRaw = sheetV.getRange(targetRow, 6).getValue();
    var eVal = String(eRaw != null ? eRaw : "").trim().toLowerCase();
    var fVal = String(fRaw != null ? fRaw : "").trim().toLowerCase();
    var eHasJa = eVal === "ja";
    var fHasGummi = fVal === "vorhanden" || fVal === "da" || fVal.indexOf("vorhanden") !== -1;

    var needWssWrite = wssYes && (isNew || !eHasJa);
    var needGummiWrite = gummiYes && (isNew || !fHasGummi);
    if (!needWssWrite && !needGummiWrite) {
      return {
        success: true,
        skipped: true,
        message: "Vasold hatte die gewählten Werte bereits — nichts geändert."
      };
    }

    sheetV.getRange(targetRow, 1).setValue(stockId);
    sheetV.getRange(targetRow, 2).setValue(carol);
    sheetV.getRange(targetRow, 3).setValue(marke);
    sheetV.getRange(targetRow, 4).setValue(huVasold);
    if (needWssWrite) {
      sheetV.getRange(targetRow, 5).setValue("Ja");
    }
    if (needGummiWrite) {
      sheetV.getRange(targetRow, 6).setValue("Vorhanden");
    }

    var wroteEJa = needWssWrite;
    if (wroteEJa) {
      var curCom = String(sheetRef.getRange(refurbRow, 25).getValue() || "");
      if (!/wss\s+da\b/i.test(curCom)) {
        sheetRef.getRange(refurbRow, 25).setValue(curCom ? "WSS da // " + curCom : "WSS da // ");
      }
    }

    SpreadsheetApp.flush();
    var parts = [];
    if (needWssWrite) parts.push("WSS Ja");
    if (needGummiWrite) parts.push("Gummileiste");
    return {
      success: true,
      skipped: false,
      message: "Vasold WSS: " + parts.join(", ") + (isNew ? " (neue Zeile)" : " — Zeile aktualisiert"),
      stockId: stockId,
      isNew: isNew,
      row: targetRow
    };
  } catch (err) {
    return { success: false, message: "Fehler: " + err.message };
  }
}

  function onOpen() {
    SpreadsheetApp.getUi().createMenu('WMS')
      .addItem('Öffne Warehouse Management System', 'openWMS')
      .addToUi();
  }
  
  function openWMS() {
    var html = HtmlService.createHtmlOutputFromFile('WMS_HUD')
      .setWidth(2400)
      .setHeight(1600);
    SpreadsheetApp.getUi().showModelessDialog(html, 'Warehouse Management System');
  }

  function applyTrackingDateIfEmpty(stockId) {
    stockId = normalizeStockId(stockId);
    var trackingSs = SpreadsheetApp.openByUrl(TRACKING_SHEET_URL);
    var sheet = trackingSs.getSheetByName("Stock ID extern Tracking");
    if (!sheet) return { success: false, updated: false, message: "Reiter 'Stock ID extern Tracking' fehlt!" };

    var lastRow = Math.max(2, sheet.getLastRow());
    var data = sheet.getRange(1, 1, lastRow, 9).getValues();

    for (var i = 0; i < data.length; i++) {
      if (cellMatchesStockId(data[i][0], stockId)) {
        var currentDate = String(data[i][8] || "").trim();
        if (currentDate !== "") {
          return { success: true, updated: false, message: "Datum bereits gesetzt" };
        }
        var dateStr = Utilities.formatDate(new Date(), "Europe/Berlin", "dd.MM.yyyy");
        sheet.getRange(i + 1, 9).setValue(dateStr);
        SpreadsheetApp.flush();
        return { success: true, updated: true, message: "Datum gesetzt" };
      }
    }

    return { success: false, updated: false, message: "Stock-ID in Stock ID extern Tracking nicht gefunden!" };
  }
  
  function searchByOrderNumber(query) {
    try {
      query = String(query || "").replace(/\s+/g, '').toUpperCase();
      if (!query) return { found: false, message: "Keine Suchanfrage" };

      var cleanQuery = query.replace(/^N4P/i, '');
      if (!cleanQuery || cleanQuery.length < 4) return { found: false, message: "Suchanfrage zu kurz (min. 4 Zeichen)" };

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Refurbisment List");
      if (!sheet) return { found: false, message: "Reiter 'Refurbisment List' fehlt!" };

      var lastRow = Math.max(2, sheet.getLastRow());
      var kommData = sheet.getRange(1, 24, lastRow, 1).getValues();
      var stockData = sheet.getRange(1, 2, lastRow, 1).getValues();

      for (var i = 1; i < kommData.length; i++) {
        var cellText = String(kommData[i][0] || "").replace(/\s+/g, '').toUpperCase();
        if (!cellText) continue;
        if (cellText.indexOf(cleanQuery) !== -1 || cellText.indexOf(query) !== -1) {
          var stockId = String(stockData[i][0] || "").trim();
          if (stockId) {
            return { found: true, stockId: stockId, message: "Gefunden via Bestellnummer in Zeile " + (i + 1) };
          }
        }
      }

      return { found: false, message: "Bestellnummer '" + query + "' nicht in Kommentar Ersatzteile Bestellung gefunden." };
    } catch (err) {
      return { found: false, message: "Fehler: " + err.message };
    }
  }

  function getShelfCounts() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Refurbisment List");
      if (!sheet) return { success: false, shelves: [] };

      var lastRow = Math.max(2, sheet.getLastRow());
      var regalColData = sheet.getRange(1, 28, lastRow, 1).getValues();

      var counts = {};
      for (var i = 1; i <= 9; i++) {
        for (var j = 1; j <= 8; j++) {
          counts["Regal " + i + "." + j] = 0;
        }
      }
      for (var r = 1; r < regalColData.length; r++) {
        var regalVal = String(regalColData[r][0] || "").trim();
        var rk = normalizeRegalKeyForCount(regalVal);
        if (rk && counts.hasOwnProperty(rk)) counts[rk]++;
      }

      mergeNachbestellungenRegalIntoCounts(counts);

      var shelves = [];
      for (var shelf in counts) {
        shelves.push({ name: shelf, count: counts[shelf] });
      }
      return { success: true, shelves: shelves };
    } catch (err) {
      return { success: false, shelves: [] };
    }
  }

  function fetchWmsData(stockId) {
    try {
      stockId = normalizeStockId(stockId);
      if (!stockId) return { success: false, message: "Keine Stock-ID" };
  
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Refurbisment List");
      if (!sheet) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };
  
      var lastRow = Math.max(2, sheet.getLastRow());
      var stockColData = sheet.getRange(1, 2, lastRow, 1).getValues();
      var regalColData = sheet.getRange(1, 28, lastRow, 1).getValues();
      var result = { success: false };
      var hitRow = -1;
  
      var counts = {};
      for (var i = 1; i <= 9; i++) {
        for (var j = 1; j <= 8; j++) {
          counts["Regal " + i + "." + j] = 0;
        }
      }
  
      for (var r = 1; r < regalColData.length; r++) {
        var regalVal = String(regalColData[r][0] || "").trim();
        var rk2 = normalizeRegalKeyForCount(regalVal);
        if (rk2 && counts.hasOwnProperty(rk2)) counts[rk2]++;
      }

      mergeNachbestellungenRegalIntoCounts(counts);

      for (var s = 1; s < stockColData.length; s++) {
        if (cellMatchesStockId(stockColData[s][0], stockId)) {
          hitRow = s + 1;
          break;
        }
      }
  
      var availableShelves = [];
      for (var shelf in counts) {
        if (counts[shelf] < 5) {
          availableShelves.push({ name: shelf, count: counts[shelf] });
        }
      }
      result.freeShelves = availableShelves;
  
      if (hitRow > 0) {
        var rowData = sheet.getRange(hitRow, 1, 1, 30).getValues()[0];
        result.success = true;
        result.carolUrl = String(rowData[2] || "");
        result.schaeden = String(rowData[22] || "");
        result.kommBestellung = String(rowData[23] || "");
        result.kommAnlieferung = String(rowData[24] || "");
        result.status = String(rowData[25] || "");
        result.regal = String(rowData[27] || "");
        result.reifenStatus = String(rowData[29] || "");
        result.markeModel = String(rowData[12] || "");
        var curKey = normalizeRegalKeyForCount(result.regal);
        result.currentShelfCount = (curKey && counts.hasOwnProperty(curKey)) ? counts[curKey] : (counts[result.regal] || 0);
        result.currentShelfCapacity = 5;
      } else {
        result.message = "Stock-ID in Refurbisment List nicht gefunden!";
      }
  
      return result;
    } catch (err) {
      return { success: false, message: err.message };
    }
  }

function getRefurbishmentCachePayload() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Refurbisment List");
    if (!sheet) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };

    var lastRow = Math.max(2, sheet.getLastRow());
    var rawRows = [];
    if (lastRow >= 2) {
      rawRows = sheet.getRange(2, 1, lastRow, 30).getValues();
    }
    var rows = [];
    var ri;
    for (ri = 0; ri < rawRows.length; ri++) {
      var r = rawRows[ri];
      rows.push([
        r[1], r[2], r[22], r[23], r[24], r[25], r[27], r[29], r[12]
      ]);
    }

    var regalColData = sheet.getRange(1, 28, lastRow, 1).getValues();
    var counts = {};
    for (var i = 1; i <= 9; i++) {
      for (var j = 1; j <= 8; j++) {
        counts["Regal " + i + "." + j] = 0;
      }
    }
    for (var r = 1; r < regalColData.length; r++) {
      var regalVal = String(regalColData[r][0] || "").trim();
      var rk2 = normalizeRegalKeyForCount(regalVal);
      if (rk2 && counts.hasOwnProperty(rk2)) counts[rk2]++;
    }
    mergeNachbestellungenRegalIntoCounts(counts);

    return {
      success: true,
      version: Date.now(),
      lastRow: lastRow,
      rows: rows,
      counts: counts
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

  function saveKommentar(stockId, text) {
    try {
      stockId = normalizeStockId(stockId);
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
      var lastRow = Math.max(2, sheet.getLastRow());
      var data = sheet.getRange(1, 2, lastRow, 1).getValues();

      for (var i = 1; i < data.length; i++) {
        if (cellMatchesStockId(data[i][0], stockId)) {
          var row = i + 1;
          sheet.getRange(row, 25).setValue(text);
          SpreadsheetApp.flush();
          var check = sheet.getRange(row, 25).getValue();
          if (check != text) return { success: false, message: "Fehler beim Verifizieren!" };

          var dateResult = applyTrackingDateIfEmpty(stockId);
          var msg = "Kommentar gespeichert!";
          if (dateResult.updated) msg += " Datum gesetzt!";
          if (!dateResult.success) msg += " " + dateResult.message;
          return { success: true, message: msg };
        }
      }
      return { success: false, message: "Stock-ID nicht gefunden!" };
    } catch (err) {
      return { success: false, message: "Fehler: " + err.message };
    }
  }
  
  function einlagern(stockId, regal) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
    var lastRow = Math.max(2, sheet.getLastRow());
    var data = sheet.getRange(1, 2, lastRow, 1).getValues();

    for (var i = 1; i < data.length; i++) {
      if (cellMatchesStockId(data[i][0], stockId)) {
        var row = i + 1;
        sheet.getRange(row, 28).setValue(regal);
        SpreadsheetApp.flush();
        var check = sheet.getRange(row, 28).getValue();
        return (check == regal) ? { success: true, message: "In " + regal + " eingelagert!" } : { success: false, message: "Fehler beim Verifizieren!" };
      }
    }
    return { success: false, message: "Stock-ID nicht gefunden!" };
  }

  function saveKommentarUndRegal(stockId, text, regal) {
    try {
      stockId = normalizeStockId(stockId);
      regal = String(regal || "").trim();
      text = String(text || "");
      if (!stockId) return { success: false, message: "Keine Stock-ID" };
      if (!text.trim()) return { success: false, message: "Bitte erst Kommentar eintragen!" };

      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
      if (!sheet) return { success: false, message: "Reiter 'Refurbisment List' fehlt!" };

      var lastRow = Math.max(2, sheet.getLastRow());
      var data = sheet.getRange(1, 2, lastRow, 1).getValues();

      for (var i = 1; i < data.length; i++) {
        if (cellMatchesStockId(data[i][0], stockId)) {
          var row = i + 1;
          sheet.getRange(row, 25).setValue(text);
          sheet.getRange(row, 25).setBackground("#ff0000");
          sheet.getRange(row, 26).setValue("Teilweise angeliefert");
          if (regal) {
            sheet.getRange(row, 28).setValue(regal);
          }
          SpreadsheetApp.flush();

          var commentCheck = sheet.getRange(row, 25).getValue();
          var statusCheck = String(sheet.getRange(row, 26).getValue() || "").trim();
          if (commentCheck != text || statusCheck !== "Teilweise angeliefert") {
            return { success: false, message: "Fehler beim Verifizieren!" };
          }
          if (regal) {
            var regalCheck = String(sheet.getRange(row, 28).getValue() || "").trim();
            if (regalCheck !== regal) return { success: false, message: "Fehler beim Verifizieren!" };
          }

          var dateResult = applyTrackingDateIfEmpty(stockId);
          var msg = regal
            ? "Kommentar und Regal gespeichert! Status auf Teilweise angeliefert gesetzt."
            : "Kommentar gespeichert! Status auf Teilweise angeliefert gesetzt.";
          if (dateResult.updated) msg += " Datum gesetzt!";
          if (!dateResult.success) msg += " " + dateResult.message;
          return { success: true, message: msg };
        }
      }

      return { success: false, message: "Stock-ID nicht gefunden!" };
    } catch (err) {
      return { success: false, message: "Fehler: " + err.message };
    }
  }

  function getStockRegalOverview() {
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Refurbisment List");
      if (!sheet) return { success: false, message: "Reiter 'Refurbisment List' fehlt!", entries: [] };

      var lastRow = Math.max(2, sheet.getLastRow());
      var data = sheet.getRange(1, 1, lastRow, 30).getValues();
      var entries = [];

      for (var i = 1; i < data.length; i++) {
        var stockId = String(data[i][1] || "").trim();
        var regal = String(data[i][27] || "").trim();
        if (!stockId) continue;
        if (!/^Regal\s+\d+\.\d+$/i.test(regal)) continue;
        entries.push({
          stockId: stockId,
          regal: regal,
          kommBestellung: String(data[i][23] || ""),
          kommAnlieferung: String(data[i][24] || ""),
          regalReifen: String(data[i][29] || ""),
          status: String(data[i][25] || "")
        });
      }

      entries.sort(function(a, b) {
        var regalMatchA = String(a.regal).match(/^Regal\s+(\d+)\.(\d+)$/i);
        var regalMatchB = String(b.regal).match(/^Regal\s+(\d+)\.(\d+)$/i);

        if (regalMatchA && regalMatchB) {
          var aMain = parseInt(regalMatchA[1], 10);
          var bMain = parseInt(regalMatchB[1], 10);
          if (aMain !== bMain) return aMain - bMain;
          var aSub = parseInt(regalMatchA[2], 10);
          var bSub = parseInt(regalMatchB[2], 10);
          if (aSub !== bSub) return aSub - bSub;
        } else if (regalMatchA) {
          return -1;
        } else if (regalMatchB) {
          return 1;
        } else {
          var regalCompare = String(a.regal).localeCompare(String(b.regal), "de");
          if (regalCompare !== 0) return regalCompare;
        }

        return normalizeStockId(a.stockId).localeCompare(normalizeStockId(b.stockId), "de");
      });

      return { success: true, entries: entries };
    } catch (err) {
      return { success: false, message: err.message, entries: [] };
    }
  }
  
  function triggerCarolMission(stockId) {
    try {
      stockId = normalizeStockId(stockId);
      if (!stockId) return { success: false, message: "Keine Stock-ID", oldRegal: "LEER", carolUrl: "" };

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var log = [];
      var result = { success: false, oldRegal: "LEER", carolUrl: "" };

      var dateResult = applyTrackingDateIfEmpty(stockId);
      if (dateResult.updated) log.push("Datum gesetzt");
      if (dateResult.success && !dateResult.updated) log.push("Datum bereits vorhanden");
      if (!dateResult.success) log.push(dateResult.message);

      var sheetRefurb = ss.getSheetByName("Refurbisment List");
      if (!sheetRefurb) return { success: false, message: "Reiter Refurbisment List fehlt!", oldRegal: "LEER", carolUrl: "" };

      var lastRowRefurb = Math.max(2, sheetRefurb.getLastRow());
      var dataRefurb = sheetRefurb.getRange(1, 2, lastRowRefurb, 1).getValues();
      var foundRow = -1;
      for (var j = 0; j < dataRefurb.length; j++) {
        if (cellMatchesStockId(dataRefurb[j][0], stockId)) {
          foundRow = j + 1;
          break;
        }
      }
      if (foundRow === -1) {
        return { success: false, message: "Stock-ID " + stockId + " in Refurbisment List nicht gefunden!", oldRegal: "LEER", carolUrl: "" };
      }

      result.oldRegal = String(sheetRefurb.getRange(foundRow, 28).getValue() || "LEER");
      result.carolUrl = String(sheetRefurb.getRange(foundRow, 3).getValue() || "");

      sheetRefurb.getRange(foundRow, 25).setBackground("#00FF00");
      sheetRefurb.getRange(foundRow, 26).setValue("Herausgegeben");
      sheetRefurb.getRange(foundRow, 28).setValue("Tagesliste");
      SpreadsheetApp.flush();

      result.success = true;
      log.push("Zeile " + foundRow + " auf Tagesliste gesetzt");
      result.message = log.join(" | ");
      return result;
    } catch (err) {
      return { success: false, message: "Fehler: " + err.message, oldRegal: "LEER", carolUrl: "" };
    }
  }

function getNachbestellungen() {
    try {
      var ss = SpreadsheetApp.openById(NACHBESTELL_SHEET_ID);
      var sheet = ss.getSheetByName(NACHBESTELL_TAB);
      if (!sheet) return { success: false, message: "Tab '" + NACHBESTELL_TAB + "' nicht gefunden!", entries: [], lagerortOptions: [] };

      var lagerortCol = findNachbestellungLagerortColumn(sheet);
      var lastColBound = Math.max(1, Math.min(80, sheet.getLastColumn()));
      if (lagerortCol > lastColBound) lagerortCol = Math.min(NACHBESTELL_REGAL_COL, lastColBound);
      var lagerortOptions = getNachbestellungLagerortAllowedList(sheet, lagerortCol);

      var lastRow = Math.max(2, sheet.getLastRow());
      var lastCol = Math.max(1, Math.min(80, sheet.getLastColumn()));
      var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

      var headerIdx = -1;
      for (var h = 0; h < Math.min(10, data.length); h++) {
        var row = data[h].map(function(c) { return String(c || "").toLowerCase(); });
        if (row.some(function(c) { return c.indexOf("stock") !== -1; })) {
          headerIdx = h;
          break;
        }
      }
      if (headerIdx === -1) headerIdx = 0;

      var header = data[headerIdx];
      var cols = {};
      for (var c = 0; c < header.length; c++) {
        var txt = String(header[c] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
        if (!cols.date && (txt === "datum" || txt === "date")) cols.date = c;
        if (!cols.stock && (txt.indexOf("stock") !== -1)) cols.stock = c;
        if (!cols.url && (txt.indexOf("carol") !== -1 || txt.indexOf("url") !== -1 || txt.indexOf("link") !== -1)) cols.url = c;
        if (!cols.typ && (txt.indexOf("typ") !== -1 || txt.indexOf("art") !== -1 || txt.indexOf("bestellung") !== -1 || txt.indexOf("beschreibung") !== -1)) cols.typ = c;
        if (!cols.person && (txt.indexOf("team") !== -1 || txt.indexOf("name") !== -1 || txt.indexOf("person") !== -1 || txt.indexOf("mechaniker") !== -1)) cols.person = c;
        if (!cols.teil && (txt.indexOf("teil") !== -1 || txt.indexOf("article") !== -1 || txt.indexOf("ersatzteil") !== -1 || txt.indexOf("benennung") !== -1)) cols.teil = c;
        if (!cols.preis && (txt.indexOf("preis") !== -1 || txt.indexOf("kosten") !== -1 || txt.indexOf("price") !== -1)) cols.preis = c;
        if (!cols.artikel && (txt.indexOf("artikelnr") !== -1 || txt.indexOf("artikelnummer") !== -1 || txt.indexOf("article") !== -1 || txt.indexOf("teilenr") !== -1)) cols.artikel = c;
        if (!cols.status && (txt.indexOf("status") !== -1 || txt.indexOf("bestellt") !== -1 || txt.indexOf("angeliefert") !== -1)) cols.status = c;
      }

      if (cols.stock === undefined) {
        for (var bc = 0; bc < header.length; bc++) {
          var sample = String(data[headerIdx + 1] ? data[headerIdx + 1][bc] : "").trim();
          if (/^[A-Z]{2}\d{4,}/.test(sample)) { cols.stock = bc; break; }
        }
      }

      if (cols.stock === undefined) return { success: false, message: "Spalte 'Stock ID' nicht gefunden!", entries: [], lagerortOptions: lagerortOptions };

      var entries = [];
      
      for (var i = headerIdx + 1; i < data.length; i++) {
        var stockId = String(data[i][cols.stock] || "").trim();
        
        if (!stockId) continue;

        var rawStatus = cols.status !== undefined ? String(data[i][cols.status] || "").trim() : "";
        var statusVal = rawStatus.toLowerCase();
        if (statusVal === "angeliefert" || statusVal.indexOf("angeliefert/bereit") !== -1 || statusVal === "fahrzeug rr") continue;

        var dateVal = cols.date !== undefined ? data[i][cols.date] : "";
        var dateStr = "";
        if (dateVal instanceof Date) {
          dateStr = Utilities.formatDate(dateVal, "Europe/Berlin", "dd.MM.yyyy");
        } else if (Object.prototype.toString.call(dateVal) === '[object Date]') {
          dateStr = Utilities.formatDate(dateVal, "Europe/Berlin", "dd.MM.yyyy");
        } else {
          var rawDate = String(dateVal || "");
          var jsDateMatch = rawDate.match(/(\w+)\s(\w+)\s(\d{1,2})\s(\d{4})/);
          if (jsDateMatch) {
            var months = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
            var mm = months[jsDateMatch[2]] || 0;
            var dd = parseInt(jsDateMatch[3], 10);
            var yyyy = parseInt(jsDateMatch[4], 10);
            if (mm > 0) {
              dateStr = (dd < 10 ? "0" : "") + dd + "." + (mm < 10 ? "0" : "") + mm + "." + yyyy;
            } else {
              dateStr = rawDate;
            }
          } else {
            dateStr = rawDate;
          }
        }

        var typ = cols.typ !== undefined ? String(data[i][cols.typ] || "").trim() : "";

        var regalVal = "";
        if (lastCol >= lagerortCol) {
          var rawL = data[i][lagerortCol - 1];
          regalVal = nachbestellungRegalUiFromCell(rawL);
        }

        entries.push({
          row: i + 1,
          date: dateStr,
          stockId: stockId,
          url: cols.url !== undefined ? String(data[i][cols.url] || "").trim() : "",
          typ: typ,
          person: cols.person !== undefined ? String(data[i][cols.person] || "").trim() : "",
          teil: cols.teil !== undefined ? String(data[i][cols.teil] || "").trim() : "",
          preis: cols.preis !== undefined ? String(data[i][cols.preis] || "").trim() : "",
          artikel: cols.artikel !== undefined ? String(data[i][cols.artikel] || "").trim() : "",
          status: cols.status !== undefined ? String(data[i][cols.status] || "").trim() : "",
          regal: regalVal
        });
      }

      entries.sort(function(a, b) {
        var ma = String(a.date).match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
        var mb = String(b.date).match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
        if (ma && mb) {
          var da = new Date(parseInt(ma[3], 10), parseInt(ma[2], 10) - 1, parseInt(ma[1], 10)).getTime();
          var db = new Date(parseInt(mb[3], 10), parseInt(mb[2], 10) - 1, parseInt(mb[1], 10)).getTime();
          if (da !== db) return db - da;
        }
        return b.row - a.row;
      });

      return { success: true, entries: entries, lagerortOptions: lagerortOptions };
    } catch (err) {
      return { success: false, message: err.message, entries: [], lagerortOptions: [] };
    }
  }

function updateNachbestellung(sheetRow, fieldName, value) {
    try {
      var ss = SpreadsheetApp.openById(NACHBESTELL_SHEET_ID);
      var sheet = ss.getSheetByName(NACHBESTELL_TAB);
      if (!sheet) return { success: false, message: "Tab nicht gefunden!" };

      if (fieldName === "regal") {
        var lagerortCol = findNachbestellungLagerortColumn(sheet);
        var lcW = Math.max(1, sheet.getLastColumn());
        if (lagerortCol > lcW) lagerortCol = Math.min(NACHBESTELL_REGAL_COL, lcW);
        var allowedLv = getNachbestellungLagerortAllowedList(sheet, lagerortCol);
        var rawIn = String(value || "").trim();
        var regalWrite = rawIn === "" ? "" : nachbestellungLagerortToSheetValue(rawIn, allowedLv);
        if (rawIn !== "" && regalWrite === null) {
          return { success: false, message: "Lagerort nicht in der Sheet-Liste. Bitte passenden Eintrag wählen." };
        }
        var lr = sheet.getRange(sheetRow, lagerortCol);
        lr.setNumberFormat("@");
        lr.setValue(regalWrite);
        SpreadsheetApp.flush();
        var verifyCell = lr.getValue();
        if (!nachbestellungLagerortVerifyMatch(regalWrite, verifyCell)) return { success: false, message: "Lagerort konnte nicht verifiziert werden!" };
        return { success: true, message: "Lagerort gespeichert!" };
      }

      var lastCol = Math.max(1, Math.min(15, sheet.getLastColumn()));
      var headerData = sheet.getRange(1, 1, Math.min(10, sheet.getLastRow()), lastCol).getValues();
      var headerIdx = -1;
      for (var h = 0; h < headerData.length; h++) {
        var row = headerData[h].map(function(c) { return String(c || "").toLowerCase(); });
        if (row.some(function(c) { return c.indexOf("stock") !== -1; })) { headerIdx = h; break; }
      }
      if (headerIdx === -1) headerIdx = 0;

      var header = headerData[headerIdx];
      var colMap = {};
      for (var c = 0; c < header.length; c++) {
        var txt = String(header[c] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
        if (txt.indexOf("teil") !== -1 || txt.indexOf("benennung") !== -1 || txt.indexOf("ersatzteil") !== -1) colMap["teil"] = c + 1;
        if (txt.indexOf("artikelnr") !== -1 || txt.indexOf("artikelnummer") !== -1 || txt.indexOf("teilenr") !== -1) colMap["artikel"] = c + 1;
        if (txt.indexOf("status") !== -1 || txt.indexOf("bestellt") !== -1 || txt.indexOf("angeliefert") !== -1) colMap["status"] = c + 1;
      }

      var targetCol = colMap[fieldName];
      if (!targetCol) return { success: false, message: "Spalte '" + fieldName + "' nicht gefunden!" };

      sheet.getRange(sheetRow, targetCol).setValue(value);
      SpreadsheetApp.flush();

      var extraMsgs = [];
      if (fieldName === "status" && String(value || "").toLowerCase().indexOf("angeliefert") !== -1) {
        var rowData = sheet.getRange(sheetRow, 1, 1, lastCol).getValues()[0];
        var rowStockId = "";
        var rowTyp = "";
        var rowBeschreibung = "";
        for (var r = 0; r < rowData.length; r++) {
          var cellVal = String(rowData[r] || "").trim();
          if (!rowStockId && /^[A-Z]{2}\d{4,}/i.test(cellVal)) rowStockId = normalizeStockId(cellVal);
          if (!rowTyp && cellVal.toLowerCase().indexOf("exit") !== -1) rowTyp = cellVal;
        }
        if (colMap["teil"]) rowBeschreibung = String(sheet.getRange(sheetRow, colMap["teil"]).getValue() || "").trim();

        if (rowTyp && rowStockId) {
          extraMsgs.push(updateExitListStatus(rowStockId));
        }

        if (rowStockId && String(value || "").indexOf("Angeliefert/Bereit") !== -1) {
          extraMsgs.push(autoFillWerkstattauftrag(rowStockId, rowBeschreibung));
        }
      }

      var msg = "Gespeichert!";
      for (var m = 0; m < extraMsgs.length; m++) {
        if (extraMsgs[m]) msg += " | " + extraMsgs[m];
      }
      return { success: true, message: msg };
    } catch (err) {
      return { success: false, message: err.message };
    }
  }

function updateExitListStatus(stockId) {
    try {
      if (!stockId) return "Exit: Keine Stock-ID";
      var EXIT_STOCK_COL = 2;
      var EXIT_STATUS_COL = 8;
      var ss = SpreadsheetApp.openById(EXIT_SHEET_ID);
      var sh = ss.getSheetByName(EXIT_TAB);
      if (!sh) return "Exit: Tab '" + EXIT_TAB + "' nicht gefunden!";
      var lastRow = Math.max(1, sh.getLastRow());
      if (lastRow < 2) return "Exit: Tab leer";
      var stockData = sh.getRange(1, EXIT_STOCK_COL, lastRow, 1).getValues();
      for (var i = 0; i < stockData.length; i++) {
        if (cellMatchesStockId(stockData[i][0], stockId)) {
          var cell = sh.getRange(i + 1, EXIT_STATUS_COL);
          var validation = cell.getDataValidation();
          var targetValue = "komplett angeliefert";
          if (validation) {
            var criteria = validation.getCriteriaValues();
            if (criteria && criteria.length > 0 && Array.isArray(criteria[0])) {
              for (var v = 0; v < criteria[0].length; v++) {
                var opt = String(criteria[0][v] || "").toLowerCase();
                if (opt.indexOf("komplett") !== -1) {
                  targetValue = String(criteria[0][v]);
                  break;
                }
              }
            }
          }
          cell.setValue(targetValue);
          SpreadsheetApp.flush();
          var verify = String(cell.getValue() || "");
          if (verify.toLowerCase().indexOf("komplett") !== -1) {
            return "Exit Z" + (i + 1) + " → " + targetValue;
          }
          cell.clearDataValidations();
          cell.setValue(targetValue);
          if (validation) cell.setDataValidation(validation);
          SpreadsheetApp.flush();
          return "Exit Z" + (i + 1) + " → " + targetValue + " (forced)";
        }
      }
      return "Exit: '" + stockId + "' in '" + EXIT_TAB + "' nicht gefunden";
    } catch (err) {
      return "Exit Fehler: " + err.message;
    }
  }

function autoFillWerkstattauftrag(stockId, beschreibung) {
    try {
      var bearbeiter = Session.getActiveUser().getEmail();
      if (bearbeiter !== AUFTRAG_EMAIL) return "";
      if (!stockId) return "";
      var zielSs = SpreadsheetApp.openById(AUFTRAG_SHEET_ID);
      var zielSheet = zielSs.getSheetByName(AUFTRAG_TAB);
      if (!zielSheet) return "Auftrag: Tab '" + AUFTRAG_TAB + "' nicht gefunden";
      zielSheet.getRange("D10").setValue(stockId);
      zielSheet.getRange("D18").setValue(beschreibung || "");
      SpreadsheetApp.flush();
      return "Werkstattauftrag befüllt (" + stockId + ")";
    } catch (err) {
      return "Auftrag Fehler: " + err.message;
    }
  }

function getHeutigeReifenausgabe() {
  try {
    var sheet = SpreadsheetApp.openById(TAGESLISTE_SHEET_ID).getSheetByName(TAGESLISTE_TAB);
    if (!sheet) return { success: false, message: "Tab '" + TAGESLISTE_TAB + "' nicht gefunden!", entries: [] };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, entries: [] };

    var headerIdx = -1;
    for (var h = 0; h < Math.min(10, data.length); h++) {
      var rowLc = data[h].map(function(c) { return String(c || "").toLowerCase(); });
      if (rowLc.some(function(c) { return c.indexOf("stock") !== -1; })) { headerIdx = h; break; }
    }
    if (headerIdx === -1) headerIdx = 0;

    var header = data[headerIdx];
    var cols = {};
    for (var c = 0; c < header.length; c++) {
      var txt = String(header[c] || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, '');
      if (!cols.date && (txt === "datum" || txt === "date")) cols.date = c;
      if (!cols.stock && (txt.indexOf("stock") !== -1)) cols.stock = c;
      if (!cols.reifen && (txt.indexOf("reifen") !== -1)) cols.reifen = c;
    }
    if (cols.stock === undefined) return { success: false, message: "Spalte 'Stock ID' nicht gefunden!", entries: [] };
    if (cols.reifen === undefined) return { success: false, message: "Spalte 'Reifen' nicht gefunden!", entries: [] };

    var today = new Date(); today.setHours(0,0,0,0);
    var tomorrow = new Date(today); tomorrow.setDate(tomorrow.getDate() + 1);
    var dayAfter = new Date(tomorrow); dayAfter.setDate(dayAfter.getDate() + 1);

    var entries = [];
    for (var i = headerIdx + 1; i < data.length; i++) {
      var stockId = String(data[i][cols.stock] || "").trim();
      if (!stockId) continue;

      var reifenVal = String(data[i][cols.reifen] || "").trim().toLowerCase();
      var isReifen   = (reifenVal === "2 reifen"  || reifenVal === "4 reifen");
      var isGestellt = (reifenVal === "2 gestellt" || reifenVal === "4 gestellt");
      if (!isReifen && !isGestellt) continue;

      var dateVal = cols.date !== undefined ? data[i][cols.date] : null;
      var dateObj = null, dateStr = "";
      if (dateVal instanceof Date || Object.prototype.toString.call(dateVal) === '[object Date]') {
        dateObj = new Date(dateVal); dateObj.setHours(0,0,0,0);
        dateStr = Utilities.formatDate(dateVal, "Europe/Berlin", "dd.MM.yyyy");
      } else {
        var raw = String(dateVal || "");
        var m = raw.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
        if (m) { dateObj = new Date(+m[3], +m[2]-1, +m[1]); dateObj.setHours(0,0,0,0); dateStr = raw; }
      }
      if (!dateObj || dateObj < today || dateObj >= dayAfter) continue;

      entries.push({
        row: i + 1,
        date: dateStr,
        stockId: stockId,
        reifen: String(data[i][cols.reifen] || "").trim(),
        amount: reifenVal.indexOf("4") !== -1 ? 4 : 2,
        gestellt: isGestellt,
        reifenCol: cols.reifen + 1
      });
    }

    entries.sort(function(a, b) {
      if (a.gestellt !== b.gestellt) return a.gestellt ? 1 : -1;
      var ma = String(a.date).match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
      var mb = String(b.date).match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
      if (ma && mb) {
        var da = new Date(parseInt(ma[3], 10), parseInt(ma[2], 10) - 1, parseInt(ma[1], 10)).getTime();
        var db = new Date(parseInt(mb[3], 10), parseInt(mb[2], 10) - 1, parseInt(mb[1], 10)).getTime();
        if (da !== db) return da - db;
      }
      return a.row - b.row;
    });

    return { success: true, entries: entries };
  } catch (err) {
    return { success: false, message: err.message, entries: [] };
  }
}

function updateReifenGestellt(sheetRow, reifenCol, newValue) {
  try {
    var ss = SpreadsheetApp.openById(TAGESLISTE_SHEET_ID);
    var sheet = ss.getSheetByName(TAGESLISTE_TAB);
    if (!sheet) return { success: false, message: "Tab nicht gefunden!" };
    sheet.getRange(sheetRow, reifenCol).setValue(newValue);
    SpreadsheetApp.flush();
    var verify = String(sheet.getRange(sheetRow, reifenCol).getValue() || "").trim().toLowerCase();
    if (verify === newValue.toLowerCase()) {
      return { success: true, message: "Gespeichert!" };
    }
    return { success: false, message: "Verifikation fehlgeschlagen: '" + verify + "'" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function stelleAlleReifen(entries) {
  try {
    var ss = SpreadsheetApp.openById(TAGESLISTE_SHEET_ID);
    var sheet = ss.getSheetByName(TAGESLISTE_TAB);
    if (!sheet) return { success: false, message: "Tab nicht gefunden!" };
    var updated = 0;
    var failed = 0;
    for (var i = 0; i < entries.length; i++) {
      var e = entries[i];
      if (e.gestellt) continue;
      var newVal = e.amount === 4 ? "4 gestellt" : "2 gestellt";
      sheet.getRange(e.row, e.reifenCol).setValue(newVal);
      updated++;
    }
    SpreadsheetApp.flush();
    for (var j = 0; j < entries.length; j++) {
      var en = entries[j];
      if (en.gestellt) continue;
      var expected = en.amount === 4 ? "4 gestellt" : "2 gestellt";
      var actual = String(sheet.getRange(en.row, en.reifenCol).getValue() || "").trim().toLowerCase();
      if (actual !== expected.toLowerCase()) failed++;
    }
    if (failed > 0) return { success: false, message: updated + " aktualisiert, " + failed + " fehlgeschlagen!" };
    return { success: true, message: "Alle " + updated + " Reifen auf gestellt gesetzt!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}